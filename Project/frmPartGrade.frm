VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmPartGrade 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Proprietary Part Grade Master"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10320
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
      Index           =   5
      Left            =   2925
      MaxLength       =   30
      TabIndex        =   15
      Top             =   1515
      Width           =   4275
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
      Index           =   4
      Left            =   2925
      MaxLength       =   30
      TabIndex        =   12
      Top             =   1275
      Width           =   4275
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
      Left            =   6225
      MaxLength       =   5
      TabIndex        =   10
      Top             =   1035
      Width           =   975
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
      Left            =   2925
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1035
      Width           =   1035
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
      Left            =   2925
      MaxLength       =   1
      TabIndex        =   1
      Top             =   555
      Width           =   675
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
      Left            =   2925
      MaxLength       =   30
      TabIndex        =   2
      Top             =   795
      Width           =   4275
   End
   Begin MSComctlLib.ListView LVWheel 
      Height          =   1605
      Left            =   2910
      TabIndex        =   4
      Top             =   1755
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   2831
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   3225
      Left            =   7695
      TabIndex        =   8
      Top             =   4545
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
         DataField       =   "PartGrade_Code"
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
         DataField       =   "PartGrade_Name"
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
         DataField       =   "PartGrade_Code"
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
   Begin MSDataGridLib.DataGrid DgSubgroup 
      Height          =   2040
      Left            =   870
      Negotiate       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4980
      Visible         =   0   'False
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   3598
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   4245.166
         EndProperty
      EndProperty
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax Form"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   855
      TabIndex        =   14
      Top             =   1530
      Width           =   1335
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Tax Form"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   855
      TabIndex        =   13
      Top             =   1290
      Width           =   1650
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Tax %"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   4515
      TabIndex        =   11
      Top             =   1050
      Width           =   1455
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vat %"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   855
      TabIndex        =   9
      Top             =   1043
      Width           =   525
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suppliers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   855
      TabIndex        =   7
      Top             =   1755
      Width           =   795
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   4
      Left            =   855
      TabIndex        =   6
      Top             =   563
      Width           =   450
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grade Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   855
      TabIndex        =   5
      Top             =   803
      Width           =   1080
   End
End
Attribute VB_Name = "frmPartGrade"
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
Private Const PartGrade_Code As Byte = 0
Private Const PartGrade_Name As Byte = 1
Private Const VatPer As Byte = 2
Private Const AddTaxPer As Byte = 3
Private Const PurachaseTaxForm As Byte = 4
Private Const SalesTaxForm As Byte = 5
Dim RsTaxForms As ADODB.Recordset
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset

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
Dim xITEM As ListItem
Me.left = 0: Me.top = 0
TopCtrl1.Tag = PubUParam: TopCtrl1.TopText1 = "Proprietary Part Grade Master"  ': TopCtrl1.TopText1.Width = 1000
'TopCtrl1.TopText2.Left = TopCtrl1.TopText2.Left - 1800: TopCtrl1.Left = 0: TopCtrl1.Top = 0: TopCtrl1.Width = Me.Width
LVWheel.ListItems.Clear

    Set RstMain = GCn.Execute("SELECT * FROM CONTRACTFINANCE WHERE FINCATG=2 ORDER BY FINNAME")
Do Until RstMain.EOF
    Set xITEM = LVWheel.ListItems.Add(, , RstMain!FinName)
    xITEM.ListSubItems.Add , , RstMain!FinCode
    RstMain.MoveNext
Loop
Set RstMain = New ADODB.Recordset
If PubMoveRecYn Then
    RstMain.Open "Select PartGrade_Code AS SEARCHCODE ,Part_GRADE.* From Part_GRADE Order by PartGrade_Name", GCn, adOpenDynamic, adLockOptimistic
Else
    Set RstMain = GCn.Execute("Select Top 1 PartGrade_Code AS SEARCHCODE ,Part_GRADE.* From Part_GRADE Order by PartGrade_Name")
End If
Set RsTaxForms = GCn.Execute("Select Form_Code  Code, Form_Desc  Name From TaxForms Order By Form_Desc")
Set DgSubgroup.DataSource = RsTaxForms

Set RstHelp = New ADODB.Recordset
RstHelp.Open "Select * From Part_GRADE Order by PartGrade_Name", GCn, adOpenDynamic, adLockOptimistic
CtrlClckCol
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0
Set DGCity.DataSource = RstHelp
DGCity.Visible = False
LVWheel.Enabled = False
End Sub
Private Sub Form_Resize()
'    TopCtrl1.Width = Me.Width
End Sub
Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing
End Sub
Private Sub CtrlClckCol()
    Txt(PartGrade_Code).BackColor = CtrlBColOrg:      Txt(PartGrade_Code).ForeColor = CtrlFColOrg
    Txt(PartGrade_Name).BackColor = CtrlBColOrg:      Txt(PartGrade_Name).ForeColor = CtrlFColOrg
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next I
    LVWheel.Enabled = Enb
End Sub
Private Sub MakeBlank()
Dim I As Integer
    Txt(PartGrade_Code) = ""
    Txt(PartGrade_Name) = ""
    Txt(VatPer) = ""
    Txt(AddTaxPer) = ""
    Txt(PurachaseTaxForm) = ""
    Txt(PurachaseTaxForm).Tag = ""
    Txt(SalesTaxForm) = ""
    Txt(SalesTaxForm).Tag = ""
    
    For I = 1 To LVWheel.ListItems.Count
        LVWheel.ListItems(I).Checked = False
    Next
End Sub
Private Sub MoveRec()
On Error GoTo ErrLoop
Dim OEMstr As String, RST1 As ADODB.Recordset, xITEM As ListItem, I As Integer
RST_BOF_EOF RstMain
TopCtrl1.tDel = False
If RstMain.RecordCount <= 0 Then
    MakeBlank
Else
    MakeBlank
    Txt(PartGrade_Code) = XNull(RstMain!PartGrade_Code)
    Txt(PartGrade_Name) = XNull(RstMain!PartGrade_Name)
    Txt(VatPer) = Format(VNull(RstMain!VatPer), "0.00")
    Txt(AddTaxPer) = Format(VNull(RstMain!AddTaxPer), "0.00")
    
    Txt(PurachaseTaxForm).Tag = XNull(RstMain!TaxFormPurchase)
    If Txt(PurachaseTaxForm).Tag <> "" Then
        Txt(PurachaseTaxForm) = GCn.Execute("Select Form_Desc from TaxForms where Form_Code = '" & Txt(PurachaseTaxForm).Tag & "'").Fields(0)
    End If
    Txt(SalesTaxForm).Tag = XNull(RstMain!TaxFormSales)
    If Txt(SalesTaxForm).Tag <> "" Then
        Txt(SalesTaxForm) = GCn.Execute("Select Form_Desc from TaxForms where Form_Code = '" & Txt(SalesTaxForm).Tag & "'").Fields(0)
    End If
    OEMstr = XNull(RstMain!OEM_CODE)
    Set RST1 = GCn.Execute("SELECT * FROM CONTRACTFINANCE WHERE FINCATG=2 ORDER BY FINCODE")
    For I = 1 To LVWheel.ListItems.Count
        LVWheel.ListItems(I).Checked = False
    Next
    If RST1.RecordCount > 0 Then
        Do While Len(OEMstr) > 0
            RST1.MoveFirst
            RST1.FIND ("FINCODE=" & Chk_Text(mID(OEMstr, 1, 6)))
            If Not RST1.EOF Then
                Set xITEM = LVWheel.FindItem(RST1!FinName)
                If xITEM Is Nothing Then
                Else
                    xITEM.Checked = True
                End If
            End If
            OEMstr = mID(OEMstr, 7, 100)
        Loop
    End If
End If
Exit Sub
ErrLoop:        MsgBox err.Description
End Sub

Private Sub LVWheel_GotFocus()
If DGCity.Visible = True Then DGCity.Visible = False
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrLoop
MakeBlank
ADDFLAG = 1
Disp_Text SETS("ADD", Me, RstMain)
Txt(PartGrade_Code).Tag = Txt(PartGrade_Code)
Txt_GotFocus PartGrade_Code
Txt(PartGrade_Code).SetFocus
Exit Sub
ErrLoop:    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo ErrLoop
If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(PartGrade_Code).Enabled = False
    Txt(PartGrade_Name).Tag = Txt(PartGrade_Name)
    Txt_GotFocus PartGrade_Name
    Txt(PartGrade_Name).SetFocus
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
            XBM = RstMain.Bookmark
                GCn.Execute ("Delete From PART_GRADE Where PartGrade_Code=" & Chk_Text(Trim(Txt(PartGrade_Code))))
                GCn.CommitTrans
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
    GSQL = "SELECT PartGrade_Code AS SEARCHCODE,PartGrade_Code,PartGrade_Name FROM PART_GRADE"
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
Dim transFlag As Byte, OEM_CODE As String, I As Integer
On Error GoTo ErrLoop
    transFlag = 0
    OEM_CODE = ""
    If IsValid(Txt(PartGrade_Code), "Code") = False Then Txt_GotFocus PartGrade_Code: Exit Sub
    If IsValid(Txt(PartGrade_Name), "Name") = False Then Txt_GotFocus PartGrade_Name: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From PART_GRADE Where PartGrade_Code=" & Chk_Text(Trim(Txt(PartGrade_Code)))).Fields(0) > 0 Then MsgBox "Godown Code Already Exists", vbInformation, "Godown Code Validation": Txt_GotFocus PartGrade_Code: Txt(PartGrade_Code).SetFocus: Exit Sub
    For I = 1 To LVWheel.ListItems.Count
        If LVWheel.ListItems(I).Checked = True Then
            
            OEM_CODE = OEM_CODE + RTrim(LVWheel.ListItems(I).ListSubItems(1).TEXT) + Space(6 - Len(RTrim(LVWheel.ListItems(I).ListSubItems(1).TEXT)))
        End If
    Next
    GCn.BeginTrans
    transFlag = 1
    If TopCtrl1.TopText2 = "Add" Then
    GCn.Execute ("DELETE From PART_GRADE Where PartGrade_Code=" & Chk_Text(Trim(Txt(PartGrade_Code))))
    GCn.Execute ("Insert Into PART_GRADE (PartGrade_Code,Site_Code,PartGrade_Name,OEM_Code, VatPer, AddTaxPer,TaxFormPurchase, TaxFormSales, U_Name,U_EntDt,U_AE) Values(" & Chk_Text(Txt(PartGrade_Code)) & ",'" & PubSiteCode & "'," & Chk_Text(Txt(PartGrade_Name)) & "," & Chk_Text(OEM_CODE) & ", " & Val(Txt(VatPer)) & ", " & Val(Txt(AddTaxPer)) & ", " & Chk_Text(Txt(PurachaseTaxForm).Tag) & ", " & Chk_Text(Txt(SalesTaxForm).Tag) & ",  '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    Else
    GCn.Execute ("update PART_GRADE set PartGrade_Code=" & Chk_Text(Txt(PartGrade_Code)) & ",Site_Code='" & PubSiteCode & "', " & _
                 "PartGrade_Name=" & Chk_Text(Txt(PartGrade_Name)) & ",OEM_Code=" & Chk_Text(OEM_CODE) & ", VatPer=" & Val(Txt(VatPer)) & ", " & _
                 "AddTaxPer = " & Val(Txt(AddTaxPer)) & ", TaxFormPurchase= " & Chk_Text(Txt(PurachaseTaxForm).Tag) & ", " & _
                 "TaxFormSales= " & Chk_Text(Txt(SalesTaxForm).Tag) & ", " & _
                 "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "'" & "Where PartGrade_Code=" & Chk_Text(Trim(Txt(PartGrade_Code))))
    End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select PartGrade_Code AS SEARCHCODE ,Part_GRADE.* From Part_GRADE Where PartGrade_Code = " & Chk_Text(Trim(Txt(PartGrade_Code))) & " Order by PartGrade_Name")
    End If
    RstHelp.Requery
    RstMain.FIND ("PartGrade_Code=" & Chk_Text(Trim(Txt(PartGrade_Code))))
    If ADDFLAG = 1 Then
        MakeBlank
        Txt_GotFocus PartGrade_Code
        Txt(PartGrade_Code).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        DGCity.Visible = False
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
        DGCity.Visible = False
    End If
Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eRef()
Dim xITEM As ListItem, RST1 As ADODB.Recordset
    RstHelp.Requery
    LVWheel.ListItems.Clear
    Set RST1 = GCn.Execute("SELECT * FROM CONTRACTFINANCE WHERE FINCATG=2 ORDER BY FINNAME")
    Do Until RST1.EOF
        Set xITEM = LVWheel.ListItems.Add(, , RST1!FinName)
        xITEM.ListSubItems.Add , , RST1!FinCode
        RST1.MoveNext
    Loop
    Call MoveRec
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub godCodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "PartGrade_Code >=" & Chk_Text(XNull(Trim(Txt(PartGrade_Code))))
End Sub
Private Sub godNameSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "PartGrade_NAME >=" & Chk_Text(XNull(Txt(PartGrade_Name)))
End Sub
Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case PartGrade_Code, PartGrade_Name
            DGCity.Visible = True
            DGCity.top = Txt(Index).top + Txt(Index).height + 10
            DGCity.left = Txt(Index).left
            DGCity.ZOrder 0
    End Select
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
    RST_BOF_EOF RstHelp
    Txt(Index).Tag = Txt(Index)
    Txt_Click Index
    If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    DGCity.Columns(0).width = 1000.1: DGCity.Columns(1).width = 3435.024: DGCity.Columns(2).width = 1000.1
    Select Case Index
        Case PartGrade_Code
            DGCity.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "PartGrade_Code ASC"
            RstHelp.Bookmark = mBookMark
            godCodeSearch
        Case PartGrade_Name
            DGCity.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "PartGrade_NAME ASC"
            RstHelp.Bookmark = mBookMark
            godNameSearch
        Case PurachaseTaxForm, SalesTaxForm
            DgSubgroup.Move Txt(Index).left, Txt(Index).top + Txt(Index).height + 30
            If RsTaxForms.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
            If RsTaxForms.EOF = True Or RsTaxForms.BOF = True Then Exit Sub
            If Txt(Index).TEXT <> RsTaxForms!Name Then
                RsTaxForms.MoveFirst
                RsTaxForms.FIND "Name ='" & Txt(Index).TEXT & "'"
            End If
    End Select
    If DGCity.Visible = True Then DGCity.Visible = False
    If DgSubgroup.Visible = True Then DgSubgroup.Visible = False
    If Txt(Index) = "" Then Txt_Change Index
End Sub
Private Sub Txt_Click(Index As Integer)
    CtrlClckCol
    Txt(Index).ForeColor = CtrlFCol: Txt(Index).BackColor = CtrlBCol
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean
Select Case Index
    Case PurachaseTaxForm, SalesTaxForm
        DGridTxtKeyDown DgSubgroup, Txt, Index, RsTaxForms, KeyCode, False, 1

    Case PartGrade_Code
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case PartGrade_Name
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case VatPer
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            If MsgBox("Do You Want to Save?", vbYesNo) = vbYes Then
                TopCtrl1_eSave
                Exit Sub
            End If
        
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
        
End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case PurachaseTaxForm, SalesTaxForm
            DGridTxtKeyPress Txt, Index, RsTaxForms, KeyAscii, "Name", False
    
        Case VatPer
            NumPress Txt(Index), KeyAscii, 2, 2
    End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case PartGrade_Code
        godCodeSearch
    Case PartGrade_Name
        godNameSearch
End Select
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case PurachaseTaxForm, SalesTaxForm
            If RsTaxForms.RecordCount > 0 And Txt(Index) <> "" And RsTaxForms.EOF = False And RsTaxForms.BOF = False Then
                Txt(Index).Tag = RsTaxForms!Code
                Txt(Index) = RsTaxForms!Name
            Else
                Txt(Index) = ""
                Txt(Index).Tag = ""
            End If
    
        Case PartGrade_Code
            Set Rst = GCn.Execute("SELECT * FROM PART_GRADE WHERE PartGrade_Code=" & Chk_Text(Txt(PartGrade_Code)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Code Already Exists", vbInformation, "Validation": Txt(PartGrade_Code) = Txt(PartGrade_Code).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!PartGrade_Code <> RstMain!PartGrade_Code Then MsgBox "Code Already Exists", vbInformation, "Validation": Txt(PartGrade_Code) = Txt(PartGrade_Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case PartGrade_Name
            Set Rst = GCn.Execute("SELECT * FROM PART_GRADE WHERE PartGrade_Name=" & Chk_Text(Txt(PartGrade_Name)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Name Already Exists", vbInformation, "Validation": Txt(PartGrade_Name) = Txt(PartGrade_Name).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!PartGrade_Name <> RstMain!PartGrade_Name Then MsgBox "Name Already Exists", vbInformation, "Validation": Txt(PartGrade_Name) = Txt(PartGrade_Name).Tag: Cancel = True: Exit Sub
                End If
            End If
    End Select
Set Rst = Nothing
End Sub
Private Sub LVWheel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = vbKeyTab Then
    If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
        TopCtrl1_eSave
    Else
        LVWheel.SetFocus
    End If
End If
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select PartGrade_Code AS SEARCHCODE ,Part_GRADE.* From Part_GRADE Where PartGrade_Code = '" & MyValue & "' Order by PartGrade_Name")
    End If
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

