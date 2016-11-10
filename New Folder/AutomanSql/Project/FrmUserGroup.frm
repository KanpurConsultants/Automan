VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmUserGroup 
   Caption         =   "User Groups"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   11580
   Begin VB.Frame Frame2 
      Caption         =   "Permissions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   3810
      TabIndex        =   8
      Top             =   420
      Width           =   4455
      Begin VB.CommandButton Command1 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Cancel Permission"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2955
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CheckBox opt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Account"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   225
         TabIndex        =   19
         Top             =   1215
         Width           =   1395
      End
      Begin VB.CheckBox opt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   225
         TabIndex        =   18
         Top             =   1530
         Width           =   1395
      End
      Begin VB.CheckBox opt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "WorkShop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   225
         TabIndex        =   17
         Top             =   900
         Width           =   1395
      End
      Begin VB.CheckBox opt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Spare"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   225
         TabIndex        =   16
         Top             =   585
         Width           =   1395
      End
      Begin VB.CheckBox opt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Vehicle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   15
         Top             =   270
         Width           =   1395
      End
      Begin VB.CommandButton CmdAllow 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Add All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1785
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   270
         Width           =   1035
      End
      Begin VB.CommandButton Cmdrevoke 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Revoke All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1785
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1530
         Width           =   1035
      End
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H00CFE0E0&
         Caption         =   "View All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1785
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1215
         Width           =   1035
      End
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Delete All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1785
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   900
         Width           =   1035
      End
      Begin VB.CommandButton Cmdadd 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Edit All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1785
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Save  Permission"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2955
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   735
         Width           =   1200
      End
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   5
      Top             =   540
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   2280
      Width           =   8235
      Begin MSFlexGridLib.MSFlexGrid FGridPer 
         Height          =   3705
         Left            =   90
         TabIndex        =   2
         Top             =   450
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6535
         _Version        =   393216
         Cols            =   9
         BackColorFixed  =   12243913
         BackColorBkg    =   13623520
         Redraw          =   -1  'True
         Appearance      =   0
         FormatString    =   "||Module           | Options                                                   |Add   |Edit    |Delete |View| "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Module Wise Permissions"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   195
         Width           =   2145
      End
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   661
   End
   Begin MSFlexGridLib.MSFlexGrid FGridDupli 
      Height          =   6030
      Left            =   9345
      TabIndex        =   4
      Top             =   7245
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   10636
      _Version        =   393216
      Cols            =   6
      FormatString    =   "|comp|div|form      | Module | Param"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   555
      Width           =   975
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   1695
      TabIndex        =   6
      Top             =   555
      Width           =   60
   End
End
Attribute VB_Name = "FrmUserGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''FGrid1 Constants'''''''''''''
Private Const F1SelEquip    As Byte = 0
Private Const FSiteName As Byte = 1
Private Const FSiteCode      As Byte = 2




Private ADDFLAG As Integer, DNAME As String, rs1 As Recordset, con As String, uname1 As String, setflag As Boolean
'Private Const CtrlBColOrg = &HCFE0E0                       'Orginal BackColour
'Private Const CtrlFColOrg = &H80000008                   'Orginal ForeColour
'Private Const CtrlBCol = &H0&                              'Changed BackColour
'Private Const CtrlFCol = &HFFFF&                              'Changed ForeColour
Dim FillRec As Integer
Dim RsUser As ADODB.Recordset
Dim RsComp As ADODB.Recordset
Dim RsDiv As ADODB.Recordset
Dim RsModule As ADODB.Recordset



Sub Ini_Grid()
End Sub



Private Sub Chk_Validate(Index As Integer, Cancel As Boolean)
Call Cmd_Enb(True)
End Sub

Private Sub Command1_Click()
Call Cmd_Enb(False)
End Sub

Private Sub Command3_Click()
Dim I As Integer
For I = 1 To FGridDupli.Rows - 1
    If I <= FGridDupli.Rows - 1 Then
        If FGridDupli.Rows = 2 Then
            FGridDupli.Rows = 1
            Exit For
        Else
            FGridDupli.RemoveItem (I)
            I = I - 1
        End If
    End If
Next I
For I = 1 To FGridPer.Rows - 1
           FGridDupli.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & FGridPer.TextMatrix(I, 1) & Chr(9) & FGridPer.TextMatrix(I, 2) & Chr(9) & IIf(FGridPer.TextMatrix(I, 4) = "", "*", "A") & IIf(FGridPer.TextMatrix(I, 5) = "", "*", "E") & IIf(FGridPer.TextMatrix(I, 6) = "", "*", "D") & IIf(FGridPer.TextMatrix(I, 7) = "", "*", "P")
Next
Call Cmd_Enb(False)
End Sub

Private Sub FGridComp_RowColChange()
Dim I As Integer
If FillRec = 0 Then
    For I = 1 To FGridComp.Rows - 1
        FGridComp.TextMatrix(I, 0) = ""
    Next
     FGridComp.TextMatrix(FGridComp.Row, 0) = "Ü"
     Call Fill_Line(FGridComp.Row)
     Call Set_Dupli(FGridComp.Row)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or (KeyCode = 70 And Shift = 2) Or (KeyCode = 80 And Shift = 2) Or (KeyCode = 83 And Shift = 2) Or KeyCode = vbKeyEscape Or KeyCode = vbKeyF5 Or KeyCode = vbKeyF10 Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyEnd Then TopCtrl1.TopKey_Down KeyCode, Shift
If KeyCode = 27 Then
    Unload Me
End If
Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
    '0|1 Section |2 Module Name|3 Permission
    '0|1 Code |2 Division | 3 Vehicle | 4 Spare | 5 WorkShop| 6 Account| 7 SetUp| 8 Permissiion
    '0|1 Code |2 Name of Group Company | 3 Start Date |4 Permission
    Me.left = 0: Me.top = 0
    Dim I As Byte
'    On Error GoTo err
    Set RsUser = New ADODB.Recordset
    RsUser.LockType = adLockOptimistic
    RsUser.CursorType = adOpenDynamic
    Set RsUser = G_CompCn.Execute("select * from UserGroup order by user_name")

    Ini_Grid

    
    Me.height = 7635
    Me.width = 11940

    If PubULabel = "Y" Then TopCtrl1.Tag = PubUParam
    FGridPer.ColWidth(0) = 200
    FGridPer.ColWidth(1) = 0
    FGridPer.ColWidth(8) = 0
        
    Disp_Text SETS("INI", Me, RsUser)
    Call MoveRec
    Exit Sub
err:
  CheckError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Opt_Click(Index As Integer)
Call Cmd_Enb(True)
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Call Ctrl_GetFocus(Index)
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeysA vbKeyTab, True
If KeyCode = 40 And Index <> 4 Then   'keydown = 40
    SendKeysA vbKeyTab, True
ElseIf KeyCode = 38 And ADDFLAG = 1 And Index <> 0 Then    'keyup = 38
    SendKeys "+{Tab}"
ElseIf KeyCode = 38 And ADDFLAG = 2 And Index <> 1 Then    'keyup = 38
    SendKeys "+{Tab}"
End If
If KeyCode = 40 And Index = 4 Then
    FGridComp.SetFocus
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
    keyascii = Asc(UCase(Chr(keyascii)))
    Call CheckQuote(keyascii)
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case 3, 4
        If Len(txt(Index)) = 0 Or UCase(mID(txt(Index), 1, 1)) = "N" Then
            txt(Index) = "No"
        ElseIf UCase(mID(txt(Index), 1, 1)) = "Y" Then
            txt(Index) = "Yes"
        Else
            txt(Index) = "No"
        End If
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Call Ctrl_validate(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
If txt(Index).TEXT = "" Then
    MsgBox Label1(Index) & " Is Required", vbExclamation, "Validation Check"
'    txt(Index).SetFocus
    Cancel = True
    Exit Sub
End If
Select Case Index
    Case 2
        If txt(2).TEXT <> txt(1).TEXT Then
            MsgBox "Please Retype Password For Confirmation", vbExclamation, "Validation Check"
            txt(2).TEXT = ""
            Cancel = True
            Exit Sub
        End If
    Case 4
        If txt(3) <> "Yes" Then
            txt(Index) = "No"
        End If
End Select
End Sub
Private Sub TopCtrl1_eFirst()
  BUTTONS True, Me, RsUser, 1
  Call MoveRec
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo eloop1
    Dim I As Integer
    Dim GRs As Recordset
    Dim CName As String
    Disp_Text SETS("ADD", Me, RsUser)
    ADDFLAG = 1
    txt(0) = ""
    FillRec = 1
    FGridDupli.Rows = 1
'    opt(0).Enabled = IIf(Chk(0).Value = Checked, True, False)
'    opt(1).Enabled = IIf(Chk(1).Value = Checked, True, False)
'    opt(2).Enabled = IIf(Chk(2).Value = Checked, True, False)
'    opt(3).Enabled = IIf(Chk(3).Value = Checked, True, False)
'    opt(4).Enabled = IIf(Chk(4).Value = Checked, True, False)
    
    FGridPer.Rows = 1
    Set GRs = New Recordset
    GRs.Open "select UserGroup1.Param_Str, UserGroup1.User_Name, User_Module.Form_Code,User_Module.name as FormName,User_MODULE.Module_Name as ModuleName from user_module  left join UserGroup1 on user_module.form_code +user_module.Module_Name=UserGroup1.form_code + UserGroup1.Module_Name  order by user_MODULE.Module_Name,user_MODULE.name", G_CompCn, adOpenStatic, adLockReadOnly
    Do Until GRs.EOF
        If XNull(GRs!ModuleName) <> "" Then
            FGridPer.AddItem "" & Chr(9) & GRs!Form_Code & Chr(9) & GRs!ModuleName & Chr(9) & GRs!FormName
        End If
        GRs.MoveNext
    Loop
    txt(0).SetFocus
    setflag = False
    FillRec = 0
    
    Exit Sub
eloop1:
     If err.NUMBER <> 0 Then
        MsgBox err.Description, vbInformation, "Information"
    End If
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
        ADDFLAG = 0
        If UCase(txt(0)) = "SA" Then MsgBox "SA Cannot Be Deleted.", vbInformation, "Information": Exit Sub
        If PubULabel = "Y" Then
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                G_CompCn.BeginTrans
                G_CompCn.Execute ("delete from UserGroup1 where user_name='" & RsUser!user_name & "' ")
                G_CompCn.Execute ("delete from user1 where user_name='" & RsUser!user_name & "'")
                G_CompCn.Execute ("delete from UserGroup where user_name='" & RsUser!user_name & "'")
                G_CompCn.CommitTrans
                RsUser.Requery
                Call MoveRec
                BUTTONS True, Me, RsUser, 0
            End If
        Else
            MsgBox "Only SA Can Delete Any User.", vbInformation, "Information"
            Exit Sub
        End If
eloop1:
    If err.NUMBER <> 0 Then
       GCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
    End If
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RsUser)
'    FGridComp.Row = 1
'    opt(0).Enabled = IIf(Chk(0).Value = Checked, True, False)
'    opt(1).Enabled = IIf(Chk(1).Value = Checked, True, False)
'    opt(2).Enabled = IIf(Chk(2).Value = Checked, True, False)
'    opt(3).Enabled = IIf(Chk(3).Value = Checked, True, False)
'    opt(4).Enabled = IIf(Chk(4).Value = Checked, True, False)
    txt(0).Enabled = False
    DNAME = txt(0)
    setflag = False
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    RsUser.Cancel
    Unload Me
End Sub

Private Sub TopCtrl1_eLast()
 BUTTONS True, Me, RsUser, 4
 Call MoveRec
End Sub

Private Sub TopCtrl1_eNext()
 BUTTONS True, Me, RsUser, 3
 Call MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
 BUTTONS True, Me, RsUser, 2
 Call MoveRec
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ErrorLoop
    If TopCtrl1.TopText2.CAPTION = "Add" Then Call MoveRec
    Call SETS("INI", Me, RsUser)
    Call Cmd_Enb(False)
    Call MoveRec
    ADDFLAG = 0
    Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_eSave()
    Dim I As Boolean, j As Integer, mTrans As Boolean
    Dim K As Integer
'    On Error GoTo errlbl
    If ADDFLAG = 1 And UCase(txt(0)) = "SA" Then MsgBox "You can't create User Name SA !!", vbInformation, "Validation Check": Exit Sub
    If Command3.Enabled = True Then MsgBox "First Save/Cancel Permission", vbInformation, "Validation Check ": Command3.SetFocus: Exit Sub
    If IsValid(txt(0), "User Name") = False Then Exit Sub
    
    If ADDFLAG = 2 And UCase(txt(0)) = "SA" Then
        G_CompCn.BeginTrans
        mTrans = True
        G_CompCn.Execute "update UserGroup set PASSWD='" & CODIFY(RTrim(txt(1))) & "' where user_name = 'SA'"
        G_CompCn.CommitTrans
        mTrans = False
        DNAME = txt(0)
        RsUser.Requery
        uname1 = txt(0).TEXT
        setflag = True
        ADDFLAG = 0
        RsUser.FIND "user_name = 'SA'"
        Disp_Text SETS("INI", Me, RsUser)
        Call MoveRec
        Call Cmd_Enb(False)
        Exit Sub
    End If
    
    If ADDFLAG = 1 Then
        If G_CompCn.Execute("select count(*) from UserGroup where user_name='" & txt(0) & "'").Fields(0) > 0 Then MsgBox "Duplicate User Name", vbCritical, "Validation Error": Exit Sub
    Else
        If txt(0) <> DNAME Then
            If G_CompCn.Execute("select count(*) from UserGroup where user_name='" & txt(0) & "'").Fields(0) > 0 And DNAME <> RTrim(txt(0)) Then MsgBox "Duplicate User Name", vbCritical, "Validation Error": Exit Sub
        End If
    End If
    G_CompCn.BeginTrans
    mTrans = True
    G_CompCn.Execute ("delete from user1 where user_name='" & txt(0) & "' and comp_code='" & PubCenCompCode & "'")
    G_CompCn.Execute ("delete from UserGroup1 where user_name='" & txt(0) & "' ")
    
    If ADDFLAG = 1 Then
        G_CompCn.Execute "insert into UserGroup(user_name) values('" & txt(0) & "')"
    End If
    
    For j = 1 To FGridDupli.Rows - 1
        If FGridDupli.TextMatrix(j, 5) <> "****" Then G_CompCn.Execute ("insert into UserGroup1(user_name,form_code,Module_Name,param_str) values('" & txt(0) & "','" & FGridDupli.TextMatrix(j, 3) & "','" & FGridDupli.TextMatrix(j, 4) & "','" & IIf(mID(FGridDupli.TextMatrix(j, 5), 1, 1) = "*", "*", "A") & IIf(mID(FGridDupli.TextMatrix(j, 5), 2, 1) = "*", "*", "E") & IIf(mID(FGridDupli.TextMatrix(j, 5), 3, 1) = "*", "*", "D") & IIf(mID(FGridDupli.TextMatrix(j, 5), 4, 1) = "*", "*", "P") & "')")
    Next
    
    G_CompCn.CommitTrans
    mTrans = False

    DNAME = txt(0)
    RsUser.Requery
    uname1 = txt(0).TEXT
    RsUser.Requery
    setflag = True
    ADDFLAG = 0
    RsUser.FIND "user_name = '" & uname1 & "'"
    Disp_Text SETS("INI", Me, RsUser)
    Call MoveRec
    Call Cmd_Enb(False)
    Exit Sub
errlbl:
    If mTrans Then G_CompCn.RollbackTrans
    MsgBox CStr(err.NUMBER) & " : " & err.Description, vbCritical, "User Creation Failed"
    Exit Sub
End Sub
Private Sub CmdEdit_Click()
   If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    For I = 1 To FGridPer.Rows - 1
        If FGridPer.TextMatrix(I, 3) <> "" Then
            If opt(0).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Vehicle" Then
                    FGridPer.Col = 6
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 6) = "ü"
            End If
            End If
            If opt(1).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Spare" Then
                    FGridPer.Col = 6
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 6) = "ü"
                End If
            End If
            If opt(2).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Workshop" Then
                    FGridPer.Col = 6
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 6) = "ü"
                End If
            End If
            If opt(3).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Account" Then
                    FGridPer.Col = 6
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 6) = "ü"
                End If
            End If
            If opt(4).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Setup" Then
                    FGridPer.Col = 6
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 6) = "ü"
                End If
            End If
        End If
    Next
End Sub
Private Sub Cmddel_Click()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    For I = 1 To FGridPer.Rows - 1
        If FGridPer.TextMatrix(I, 3) <> "" Then
            If opt(0).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Vehicle" Then
                    FGridPer.Col = 7
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 7) = "ü"
                End If
            End If
            If opt(1).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Spare" Then
                    FGridPer.Col = 7
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 7) = "ü"
                End If
            End If
            If opt(2).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Workshop" Then
                    FGridPer.Col = 7
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 7) = "ü"
                End If
            End If
            If opt(3).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Account" Then
                    FGridPer.Col = 7
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 7) = "ü"
                End If
            End If
            If opt(4).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Setup" Then
                    FGridPer.Col = 7
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 7) = "ü"
                End If
            End If
        End If
    Next
End Sub

Private Sub CmdAllow_Click()
     If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    For I = 1 To FGridPer.Rows - 1
        If opt(0).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Vehicle" Then
                FGridPer.Col = 4
                FGridPer.Row = I
                FGridPer.CellFontName = "wingdings"
                FGridPer.CellFontSize = 18
                FGridPer.CellForeColor = vbBlue
                FGridPer.TextMatrix(I, 4) = "ü"
            End If
        End If
        If opt(1).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Spare" Then
                FGridPer.Col = 4
                FGridPer.Row = I
                FGridPer.CellFontName = "wingdings"
                FGridPer.CellFontSize = 18
                FGridPer.CellForeColor = vbBlue
                FGridPer.TextMatrix(I, 4) = "ü"
            End If
        End If
       If opt(2).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Workshop" Then
                FGridPer.Col = 4
                FGridPer.Row = I
                FGridPer.CellFontName = "wingdings"
                FGridPer.CellFontSize = 18
                FGridPer.CellForeColor = vbBlue
                FGridPer.TextMatrix(I, 4) = "ü"
            End If
        End If
       If opt(3).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Account" Then
                FGridPer.Col = 4
                FGridPer.Row = I
                FGridPer.CellFontName = "wingdings"
                FGridPer.CellFontSize = 18
                FGridPer.CellForeColor = vbBlue
                FGridPer.TextMatrix(I, 4) = "ü"
            End If
        End If
        If opt(4).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Setup" Then
                FGridPer.Col = 4
                FGridPer.Row = I
                FGridPer.CellFontName = "wingdings"
                FGridPer.CellFontSize = 18
                FGridPer.CellForeColor = vbBlue
                FGridPer.TextMatrix(I, 4) = "ü"
            End If
        End If
    Next
End Sub
Private Sub Cmdadd_Click()
Dim I As Integer
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    For I = 1 To FGridPer.Rows - 1
        If FGridPer.TextMatrix(I, 3) <> "" Then
        If opt(0).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Vehicle" Then
                    FGridPer.Col = 5
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 5) = "ü"
            End If
        End If
        If opt(1).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Spare" Then
                    FGridPer.Col = 5
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 5) = "ü"
            End If
        End If
        If opt(2).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Workshop" Then
                    FGridPer.Col = 5
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 5) = "ü"
            End If
        End If
        If opt(3).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Account" Then
                    FGridPer.Col = 5
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 5) = "ü"
            End If
        End If
        If opt(4).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Setup" Then
                    FGridPer.Col = 5
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 5) = "ü"
            End If
        End If
        End If
    Next
End Sub

Private Sub CmdRevoke_Click()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    For I = 1 To FGridPer.Rows - 1
        If opt(0).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Vehicle" Then
               FGridPer.TextMatrix(I, 4) = ""
               FGridPer.TextMatrix(I, 5) = ""
               FGridPer.TextMatrix(I, 6) = ""
               FGridPer.TextMatrix(I, 7) = ""
            End If
        End If
        If opt(1).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Spare" Then
               FGridPer.TextMatrix(I, 4) = ""
               FGridPer.TextMatrix(I, 5) = ""
               FGridPer.TextMatrix(I, 6) = ""
               FGridPer.TextMatrix(I, 7) = ""
            End If
        End If
       If opt(2).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Workshop" Then
               FGridPer.TextMatrix(I, 4) = ""
               FGridPer.TextMatrix(I, 5) = ""
               FGridPer.TextMatrix(I, 6) = ""
               FGridPer.TextMatrix(I, 7) = ""
            End If
        End If
        If opt(3).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Account" Then
               FGridPer.TextMatrix(I, 4) = ""
               FGridPer.TextMatrix(I, 5) = ""
               FGridPer.TextMatrix(I, 6) = ""
               FGridPer.TextMatrix(I, 7) = ""
            End If
        End If
        If opt(4).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Setup" Then
               FGridPer.TextMatrix(I, 4) = ""
               FGridPer.TextMatrix(I, 5) = ""
               FGridPer.TextMatrix(I, 6) = ""
               FGridPer.TextMatrix(I, 7) = ""
            End If
       End If
    Next
End Sub

Private Sub FGridPer_Click()
    If FGridPer.Col = 0 Or FGridPer.Col = 1 Or FGridPer.Col = 2 Or FGridPer.Col = 3 Or TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGridPer.Col = 0 Or FGridPer.Col = 1 Or FGridPer.Col = 2 Or FGridPer.Col = 3 Then Exit Sub
    If FGridPer.TextMatrix(FGridPer.Row, 3) = "" Then Exit Sub
    If FGridPer.TextMatrix(FGridPer.Row, 3) = "" Then Exit Sub
'    If FGridPer.TextMatrix(FGridPer.Row, 2) = "Vehicle" And Chk(0).Value = Unchecked Then Exit Sub
'    If FGridPer.TextMatrix(FGridPer.Row, 2) = "Spare" And Chk(1).Value = Unchecked Then Exit Sub
'    If FGridPer.TextMatrix(FGridPer.Row, 2) = "Workshop" And Chk(2).Value = Unchecked Then Exit Sub
'    If FGridPer.TextMatrix(FGridPer.Row, 2) = "Account" And Chk(3).Value = Unchecked Then Exit Sub
'    If FGridPer.TextMatrix(FGridPer.Row, 2) = "Setup" And Chk(4).Value = Unchecked Then Exit Sub
    
    If FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = "" Then
        FGridPer.Col = FGridPer.Col
        FGridPer.CellFontName = "wingdings"
        FGridPer.CellFontSize = 18
        If FGridPer.Col = 4 Then
            FGridPer.CellForeColor = vbBlue
        Else
            FGridPer.CellForeColor = vbBlue
        End If
        FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = "ü"
    Else
        FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = ""
        If FGridPer.Col = 4 Then
            FGridPer.TextMatrix(FGridPer.Row, 5) = ""
            FGridPer.TextMatrix(FGridPer.Row, 6) = ""
            FGridPer.TextMatrix(FGridPer.Row, 7) = ""
        End If
    End If
Call Cmd_Enb(True)
End Sub

Private Sub FGridPer_KeyPress(keyascii As Integer)
    If FGridPer.Col = 0 Or FGridPer.Col = 1 Or TopCtrl1.TopText2.CAPTION = "Browse" Or FGridPer.TextMatrix(FGridPer.Row, 0) = "" Or keyascii <> 32 Then Exit Sub
    If FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = "" Then
        FGridPer.Col = FGridPer.Col
        FGridPer.CellFontName = "wingdings"
        FGridPer.CellFontSize = 18
        FGridPer.CellForeColor = vbBlue
        FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = "ü"
    Else
        FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = ""
    End If
End Sub


Private Sub MoveRec()
On Error GoTo ELoop
Dim GRs As ADODB.Recordset
Dim Rs As Recordset, rs1 As Recordset, Name1 As String
FillRec = 1
'FGridPer.Redraw = False
'FGridComp.Redraw = False
FGridDupli.Rows = 1
FGridPer.Rows = 1

If RsUser.RecordCount > 0 Then
    txt(0) = RsUser!user_name
    Set GRs = New Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open "select  UserGroup1.*,User_Module.name as FormName,User_MODULE.Module_Name as ModuleName from UserGroup1 left join user_module on user_module.form_code +user_module.Module_Name=UserGroup1.form_code + UserGroup1.Module_Name where  UserGroup1.user_name='" & txt(0) & "' order by user_MODULE.Module_Name,user_MODULE.name", G_CompCn, adOpenDynamic, adLockOptimistic
    Do Until GRs.EOF
        FGridDupli.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & GRs!Form_Code & Chr(9) & GRs!ModuleName & Chr(9) & GRs!param_str
        GRs.MoveNext
    Loop
End If
FillRec = 0
Call Fill_Line(1)


If FGridPer.Rows = 1 Then FGridPer.AddItem ""

TopCtrl1.tFind = False
TopCtrl1.tRef = False
TopCtrl1.tPrn = False

ELoop:
End Sub

Private Function CODIFY(txt As String) As String
    Dim xxx As String
    Dim xx As Integer, MyVal As Integer
    Randomize
    MyVal = Int((99 * Rnd) + 1)
    xxx = Chr(MyVal + 27)
    For xx = 1 To Len(txt)
        xxx = xxx + Chr(Asc(mID(txt, xx, 1)) + 27 + MyVal)
    Next
    CODIFY = xxx
End Function

Private Function DCODIFY(txt As String) As String
    Dim xxx As String
    Dim xx As Integer, MyVal As Integer
    If txt = "" Then DCODIFY = "": Exit Function
    MyVal = Asc(left(txt, 1)) - 27
    xxx = ""
    For xx = 1 To Len(txt) - 1
        xxx = xxx + Chr(Asc(mID(txt, xx + 1, 1)) - 27 - MyVal)
    Next
    DCODIFY = xxx
End Function

Private Sub Disp_Text(Enb As Boolean)
    txt(0).Enabled = Enb
    
    opt(0).Enabled = Enb
    opt(1).Enabled = Enb
    opt(2).Enabled = Enb
    opt(3).Enabled = Enb
    opt(4).Enabled = Enb
    
    Cmdadd.Enabled = Enb
    CmdEdit.Enabled = Enb
    CmdDel.Enabled = Enb
    Cmdrevoke.Enabled = Enb
    CmdAllow.Enabled = Enb
    Command3.Enabled = Enb
    Command1.Enabled = Enb
End Sub
Private Sub Ctrl_validate(Index As Integer)
txt(Index).BackColor = CtrlBColOrg
txt(Index).ForeColor = CtrlFColOrg
End Sub
Private Sub Ctrl_GetFocus(Index As Integer)
txt(Index).BackColor = CtrlBCol
txt(Index).ForeColor = CtrlFCol
End Sub
Private Sub BlankText()
Dim I As Byte
For I = 0 To 3
    txt(I).TEXT = ""
Next I
End Sub
Private Sub Fill_Line(rowval)
Dim GRs As Recordset
Dim param_val As String
If 1 = 1 Then
    
    'FGridPer.Redraw = False
    If TopCtrl1.TopText2.CAPTION <> "Browse" Then
        opt(0).Value = Unchecked
        opt(1).Value = Unchecked
        opt(2).Value = Unchecked
        opt(3).Value = Unchecked
        opt(4).Value = Unchecked
    End If
    FGridPer.Rows = 1
    Set GRs = New Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open "select Distinct UserGroup1.*,User_Module.name as FormName,User_MODULE.Module_Name as ModuleName from user_module left join UserGroup1  on user_module.form_code +user_module.Module_Name=UserGroup1.form_code + UserGroup1.Module_Name order by user_MODULE.Module_Name,user_MODULE.name", G_CompCn, adOpenStatic, adLockReadOnly
    GRs.MoveFirst
    Do Until GRs.EOF
'        If NAME1 <> GRs!Name Then
         If XNull(GRs!ModuleName) <> "" Then
            FGridPer.AddItem "" & Chr(9) & GRs!Form_Code & Chr(9) & GRs!ModuleName & Chr(9) & GRs!FormName
         End If
'        NAME1 = XNull(GRs!Name)
        param = ""
        For I = 1 To FGridDupli.Rows - 1
            If UCase(FGridDupli.TextMatrix(I, 4)) = UCase(GRs!ModuleName) And UCase(FGridDupli.TextMatrix(I, 3)) = UCase(GRs!Form_Code) Then
                paramval = FGridDupli.TextMatrix(I, 5)
                Exit For
            Else
                paramval = "****"
            End If
        Next
        If paramval <> "" Then
            FGridPer.Row = FGridPer.Rows - 1
            FGridPer.Col = 4
            FGridPer.CellFontName = "wingdings"
            FGridPer.CellFontSize = 18
            FGridPer.CellForeColor = vbBlue
            FGridPer.TextMatrix(FGridPer.Rows - 1, FGridPer.Col) = IIf(mID(paramval, 1, 1) = "*", "", "ü")
            FGridPer.Col = 5
            FGridPer.CellFontName = "wingdings"
            FGridPer.CellFontSize = 18
            FGridPer.CellForeColor = vbBlue
            FGridPer.TextMatrix(FGridPer.Rows - 1, FGridPer.Col) = IIf(mID(paramval, 2, 1) = "*", "", "ü")
            FGridPer.Col = 6
            FGridPer.CellFontName = "wingdings"
            FGridPer.CellFontSize = 18
            FGridPer.CellForeColor = vbBlue
            FGridPer.TextMatrix(FGridPer.Rows - 1, FGridPer.Col) = IIf(mID(paramval, 3, 1) = "*", "", "ü")
            FGridPer.Col = 7
            FGridPer.CellFontName = "wingdings"
            FGridPer.CellFontSize = 18
            FGridPer.CellForeColor = vbBlue
            FGridPer.TextMatrix(FGridPer.Rows - 1, FGridPer.Col) = IIf(mID(paramval, 4, 1) = "*", "", "ü")
        End If
        GRs.MoveNext
    Loop
    
    
    With FGridPer
        .Row = 0
        .Col = 3: .CellFontName = "Times New Roman"
        .TextMatrix(0, 3) = "Module/Section"
        .Col = 4: .CellFontName = "Times New Roman": .CellFontSize = 10
        .TextMatrix(0, 4) = "Add"
        .Col = 5: .CellFontName = "Times New Roman": .CellFontSize = 10
        .TextMatrix(0, 5) = "Edit"
        .Col = 6: .CellFontName = "Times New Roman": .CellFontSize = 10
        .TextMatrix(0, 6) = "Delete"
        .Col = 7: .CellFontName = "Times New Roman": .CellFontSize = 10
        .TextMatrix(0, 7) = "Print"
        .Col = 8: .CellFontName = "Times New Roman": .CellFontSize = 10
        .TextMatrix(0, 8) = "View"
    End With
    
    
    
End If
'FGridPer.Redraw = True
End Sub
Private Sub Set_Dupli(GridRow As Integer)
End Sub
Private Sub Cmd_Enb(flag As Boolean)
    Command3.Enabled = flag
    Command1.Enabled = flag
'    FGridComp.Enabled = Not Flag
End Sub



