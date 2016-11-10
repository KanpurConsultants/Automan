VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmModelCat 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Vehicle Model Category Master"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrProduce 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   5970
      TabIndex        =   11
      Top             =   2490
      Visible         =   0   'False
      Width           =   4095
      Begin MSDataGridLib.DataGrid DGProduce 
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
            DataField       =   "Produce_Code"
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
            DataField       =   "Produce_Name"
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
            DataField       =   "Produce_Code"
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
               Object.Visible         =   0   'False
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
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
         Caption         =   "List of Produce"
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
      Height          =   255
      Index           =   2
      Left            =   2355
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1125
      Width           =   4245
   End
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
      Index           =   0
      Left            =   2355
      MaxLength       =   3
      TabIndex        =   1
      Top             =   555
      Width           =   1260
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
      Index           =   1
      Left            =   2355
      MaxLength       =   15
      TabIndex        =   2
      Top             =   840
      Width           =   4245
   End
   Begin VB.Frame FrCity 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   660
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   4980
      Begin MSDataGridLib.DataGrid DGCity 
         Height          =   3225
         Left            =   30
         TabIndex        =   7
         Top             =   345
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   5689
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
            DataField       =   "ModelCat_Code"
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
            DataField       =   "ModelCat_Name"
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
            DataField       =   "ModelCat_Code"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   3435.024
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
         Caption         =   "List of Model Category"
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
         TabIndex        =   5
         Top             =   30
         Width           =   4935
      End
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produce Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   630
      TabIndex        =   10
      Top             =   1140
      Width           =   1560
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Category*"
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
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   4
      Left            =   630
      TabIndex        =   6
      Top             =   555
      Width           =   1575
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   0
      Left            =   645
      TabIndex        =   3
      Top             =   840
      Width           =   1635
   End
End
Attribute VB_Name = "frmModelCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Don't Change Tag Property of (Txt) Control as it is used in other activities
'FORM COLOR &H00C0FFFF&
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset
Dim RstProduce As ADODB.Recordset
Private Const ModelCat_Code = 0, ModelCat_NAME = 1, Produce_Name = 2



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
Me.top = 0: Me.left = 0
TopCtrl1.Tag = PubUParam    ': TopCtrl1.TopText1 = Me.Caption '"Vehicle Model Category Master"
Set RstHelp = New ADODB.Recordset
RstHelp.Open "Select ModelCat_Code,ModelCat_Name FROM MODEL_CAT where left(ModelCat_Code,1)='" & PubDivCode & "' Order by ModelCat_Name", GCn, adOpenDynamic, adLockOptimistic
Set DGCity.DataSource = RstHelp

FrCity.Visible = False
Set RstMain = New ADODB.Recordset


Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    

    RstMain.Open "Select ModelCat_Code as searchcode,MODEL_CAT.*, Produce.Produce_Name " & _
                 "From MODEL_CAT  " & _
                 "Left Join Produce On Model_Cat.Produce_Code = Produce.Produce_Code " & _
                 "where left(Model_Cat.ModelCat_Code,1)='" & PubDivCode & "' " & sitecond & " Order by ModelCat_Name", GCn, adOpenDynamic, adLockOptimistic

Set RstProduce = New ADODB.Recordset
RstProduce.Open "Select Produce_Code, Produce_Name FROM Produce where Left(Produce_Code,1)='" & PubDivCode & "' Order by Produce_NAME", GCn, adOpenDynamic, adLockOptimistic
Set DGProduce.DataSource = RstProduce


'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
CtrlClckCol
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form_Unload (-1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing
    Set RstHelp = Nothing
End Sub
Private Sub CtrlClckCol()
    Txt(ModelCat_Code).BackColor = CtrlBColOrg:      Txt(ModelCat_Code).ForeColor = CtrlFColOrg
    Txt(ModelCat_NAME).BackColor = CtrlBColOrg:      Txt(ModelCat_NAME).ForeColor = CtrlFColOrg
End Sub
Private Sub Disp_Text(Enb As Boolean)
    Txt(ModelCat_Code).Enabled = Enb
    Txt(ModelCat_NAME).Enabled = Enb
    Txt(Produce_Name).Enabled = Enb
End Sub
Private Sub MakeBlank()
    Txt(ModelCat_Code) = ""
    Txt(ModelCat_NAME) = ""
    Txt(Produce_Name) = ""
    Txt(Produce_Name).Tag = ""
End Sub
Private Sub MoveRec()
On Error GoTo ErrLoop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    MakeBlank
Else
    Txt(ModelCat_Code) = XNull(RstMain!ModelCat_Code)
    Txt(ModelCat_NAME) = XNull(RstMain!ModelCat_NAME)
    Txt(Produce_Name).Tag = XNull(RstMain!Produce_Code)
    Txt(Produce_Name) = XNull(RstMain!Produce_Name)
    
End If
TopCtrl1.tDel = False
Exit Sub
ErrLoop:        MsgBox err.Description
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrLoop
MakeBlank
ADDFLAG = 1
Disp_Text SETS("ADD", Me, RstMain)
Txt(ModelCat_Code).Tag = Txt(ModelCat_Code)
Txt_GotFocus ModelCat_Code
Txt(ModelCat_Code) = PubDivCode
Txt(ModelCat_Code).SelStart = Len(Txt(ModelCat_Code))
Txt(ModelCat_Code).SetFocus
Exit Sub
ErrLoop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ErrLoop
If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(ModelCat_Code).Enabled = False
    Txt(ModelCat_NAME).Tag = Txt(ModelCat_NAME)
    Txt_GotFocus ModelCat_NAME
    Txt(ModelCat_NAME).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
ErrLoop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ErrLoop
Dim transFalg As Byte, XBM As Variant, Res As Byte
transFalg = 0
If RstMain.RecordCount > 0 Then
    If MsgBox("Are You Sure to Delete This Record ?", vbYesNo, "Confirmation") = vbYes Then
        GCn.BeginTrans
        XBM = RstMain.Bookmark
        transFalg = 1
        GCn.Execute ("Delete From MODEL_CAT Where ModelCat_Code=" & Chk_Text(Trim(Txt(ModelCat_Code))))
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
    MsgBox "No Records To Delete", vbInformation, "Information"
End If
Exit Sub

ErrLoop:    If transFalg = 1 Then GCn.RollbackTrans
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
     Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    GSQL = "Select ModelCat_Code AS SEARCHCODE,ModelCat_Name as CategoryName,ModelCat_Code AS CategoryCode FROM MODEL_CAT where left(ModelCat_Code,1)='" & PubDivCode & "' " & sitecond & " Order by ModelCat_Name"
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
        .g_FormID = 3
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
    If Len(Trim(Txt(ModelCat_Code))) = 1 Then MsgBox "Category Code should be filled ", vbOKOnly, "Validation": Txt(ModelCat_Code).SetFocus: Exit Sub
    If IsValid(Txt(ModelCat_NAME), "Category Name") = False Then Txt_GotFocus ModelCat_NAME: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From MODEL_CAT Where ModelCat_Code=" & Chk_Text(Trim(Txt(ModelCat_Code)))).Fields(0) > 0 Then MsgBox "Category Code Already Exists", vbInformation, "Godown Code Validation": Txt_GotFocus ModelCat_Code: Txt(ModelCat_Code).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If TopCtrl1.TopText2 = "Add" Then
        GCn.Execute ("DELETE From MODEL_CAT Where ModelCat_Code=" & Chk_Text(Trim(Txt(ModelCat_Code))))
        GCn.Execute ("Insert Into MODEL_CAT (ModelCat_Code,Site_Code,ModelCat_Name, Produce_Code,U_Name,U_EntDt,U_AE) Values('" & Trim(Txt(ModelCat_Code)) & "','" & PubSiteCode & "'," & Chk_Text(Txt(ModelCat_NAME)) & "," & Chk_Text(Txt(Produce_Name).Tag) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    Else
        GCn.Execute ("update  MODEL_CAT set Site_Code='" & PubSiteCode & "',ModelCat_Name=" & Chk_Text(Txt(ModelCat_NAME)) & ",Produce_Code=" & Chk_Text(Txt(Produce_Name).Tag) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "'" & " Where ModelCat_Code=" & Chk_Text(Trim(Txt(ModelCat_Code))))
    End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    RstMain.Requery
    RstHelp.Requery
    RstMain.FIND ("ModelCat_Code=" & Chk_Text(Trim(Txt(ModelCat_Code))))
    If ADDFLAG = 1 Then
        MakeBlank
        Txt_GotFocus ModelCat_Code
        Txt(ModelCat_Code).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        FrCity.Visible = False
    End If
Exit Sub
ErrLoop:    If transFlag = 1 Then GCn.RollbackTrans
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
        FrCity.Visible = False
    End If
Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub ModelCat_CodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "ModelCat_Code >=" & Chk_Text(XNull(Trim(Txt(ModelCat_Code))))
End Sub
Private Sub ModelCat_NAMESearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "ModelCat_Name >=" & Chk_Text(XNull(Txt(ModelCat_NAME)))
End Sub
Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case ModelCat_Code, ModelCat_NAME
            FrCity.Visible = True
            FrCity.top = Txt(Index).top + Txt(Index).height + 10
            FrCity.left = Txt(Index).left
            FrCity.ZOrder 0
        Case Produce_Name
            
            'If FrModelGrp.Visible = True Then FrModelGrp.Visible = False
            FrProduce.Visible = True
            FrProduce.top = Txt(Index).top + Txt(Index).height + 10
            FrProduce.left = Txt(Index).left
            FrProduce.ZOrder 0
            
    End Select
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
    If FrCity.Visible = True Then FrCity.Visible = False
    RST_BOF_EOF RstHelp
    Txt(Index).Tag = Txt(Index)
    Txt_Click Index
    If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    DGCity.Columns(0).width = 1000.1: DGCity.Columns(1).width = 3000: DGCity.Columns(2).width = 800
    Select Case Index
        Case ModelCat_Code
            DGCity.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "ModelCat_Code ASC"
            RstHelp.Bookmark = mBookMark
            ModelCat_CodeSearch
        Case ModelCat_NAME
            DGCity.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "ModelCat_Name ASC"
            RstHelp.Bookmark = mBookMark
            ModelCat_NAMESearch
        Case Produce_Name
            DGProduce.Columns(0).width = 0: DGProduce.Columns(2).width = 0
            mBookMark = RstProduce.Bookmark
            RstProduce.Sort = "Produce_Name ASC"
            RstProduce.Bookmark = mBookMark
            ModelCat_NAMESearch
            
    End Select
    If Txt(Index) = "" Then Txt_Change Index
End Sub

Private Sub Produce_NAMESearch()
If RstProduce.RecordCount <= 0 Then Exit Sub
RstProduce.MoveFirst
RstProduce.FIND "Produce_NAME >=" & Chk_Text(XNull(Txt(Produce_Name)))
End Sub

Private Sub Txt_Click(Index As Integer)
    CtrlClckCol
    Txt(Index).ForeColor = CtrlFCol: Txt(Index).BackColor = CtrlBCol
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean
Select Case Index
    Case ModelCat_Code
        'Div Code Edit restricted
        KeyCode = RestrictCode(KeyCode, Txt(Index), Shift)
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case ModelCat_NAME
            KeyCode = RestrictCode(KeyCode, Txt(Index), Shift)
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If

        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                Txt_Validate Index, result
                If result = True Then Txt(Index).SetFocus: Exit Sub
                TopCtrl1_eSave
            Else
                Txt_Click Index
                Txt(Index).SetFocus
            End If
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case Produce_Name
        If FrProduce.Visible = True Then
            Select Case KeyCode
                Case vbKeyUp
                    If Not RstProduce.BOF Then RstProduce.MovePrevious
                Case vbKeyDown
                    If Not RstProduce.EOF Then RstProduce.MoveNext
                Case 33
                    If Not RstProduce.BOF Then RstProduce.MovePrevious
                Case 34
                    If Not RstProduce.EOF Then RstProduce.MoveNext
                Case 13
                    If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                        Txt_Validate Index, result
                        If result = True Then Txt(Index).SetFocus: Exit Sub
                        TopCtrl1_eSave
                    End If
                
                    'SendKeysA vbKeyTab, True
            End Select
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, 33, 34
                    RST_BOF_EOF RstProduce
                    If Not RstProduce.BOF And Not RstProduce.EOF Then
                        Txt(Produce_Name).Tag = XNull(RstProduce!Produce_Code): Txt(Produce_Name) = XNull(RstProduce!Produce_Name)
                        Txt(Produce_Name).SelStart = 0
                    End If
            End Select
        End If
        
End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case ModelCat_Code
        KeyAscii = RestrictCode(KeyAscii, Txt(Index), 0)
End Select

End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case ModelCat_Code
        ModelCat_CodeSearch
    Case ModelCat_NAME
        ModelCat_NAMESearch
    Case Produce_Name
        Produce_NAMESearch
End Select
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case ModelCat_Code
            Set Rst = GCn.Execute("SELECT * FROM MODEL_CAT WHERE ModelCat_Code=" & Chk_Text(Txt(ModelCat_Code)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Category Code Already Exists", vbInformation, "Validation": Txt(ModelCat_Code) = Txt(ModelCat_Code).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!ModelCat_Code <> RstMain!ModelCat_Code Then MsgBox "Category Code Already Exists", vbInformation, "Validation": Txt(ModelCat_Code) = Txt(ModelCat_Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case ModelCat_NAME
            Set Rst = GCn.Execute("SELECT * FROM MODEL_CAT WHERE ModelCat_Name=" & Chk_Text(Txt(ModelCat_NAME)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Category Name Already Exists", vbInformation, "Validation": Txt(ModelCat_NAME) = Txt(ModelCat_NAME).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!ModelCat_NAME <> RstMain!ModelCat_NAME Then MsgBox "Category Name Already Exists", vbInformation, "Validation": Txt(ModelCat_NAME) = Txt(ModelCat_NAME).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case Produce_Name
            If Not RstProduce.EOF And Not RstProduce.BOF Then
                Txt(Produce_Name).Tag = XNull(RstProduce!Produce_Code): Txt(Produce_Name) = XNull(RstProduce!Produce_Name)
            Else
                Txt(Produce_Name) = "": Txt(Produce_Name).Tag = ""
            End If
            FrProduce.Visible = False
    End Select
Set Rst = Nothing
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select ModelCat_Code as searchcode,MODEL_CAT.* From MODEL_CAT  where left(ModelCat_Code,1)='" & PubDivCode & "' And ModelCat_Code  = '" & MyValue & "' Order by ModelCat_Name")
    End If
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

