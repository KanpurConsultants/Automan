VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVouPrefix 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Voucher Prefix Creation"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   11610
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   180
      TabIndex        =   21
      Top             =   3705
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   105
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   15
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
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2520
      TabIndex        =   9
      Top             =   3075
      Width           =   1980
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   8
      Top             =   2805
      Width           =   510
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   7
      Top             =   2535
      Width           =   510
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   5
      Top             =   1995
      Width           =   510
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   6
      Top             =   2265
      Width           =   1980
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   1455
      Width           =   3525
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Text            =   "0123456789"
      Top             =   1185
      Width           =   1365
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   4
      Top             =   1725
      Width           =   510
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5340
      TabIndex        =   3
      Top             =   1185
      Width           =   705
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   661
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   8325
      TabIndex        =   11
      Top             =   465
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3750
      Left            =   6540
      TabIndex        =   10
      Top             =   1485
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   6615
      _Version        =   393216
      BackColor       =   15525079
      Cols            =   5
      BackColorFixed  =   14940925
      ForeColorFixed  =   8388608
      BackColorSel    =   16308221
      ForeColorSel    =   12582912
      BackColorBkg    =   14737632
      BackColorUnpopulated=   14865856
      GridColor       =   14940925
      GridColorFixed  =   12632319
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      FormatString    =   "SrNo."
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Common Narration        Y/N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   6
      Left            =   225
      TabIndex        =   20
      Top             =   2820
      Width           =   2235
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Transaction            +/-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   4
      Left            =   225
      TabIndex        =   19
      Top             =   2010
      Width           =   2220
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Separate Narration         Y/N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   3
      Left            =   225
      TabIndex        =   18
      Top             =   2550
      Width           =   2250
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number Method              A/M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   2
      Left            =   225
      TabIndex        =   17
      Top             =   2280
      Width           =   2235
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No. from Table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   1
      Left            =   225
      TabIndex        =   16
      Top             =   3090
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   22
      Left            =   225
      TabIndex        =   15
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division Base Number  Y/N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   20
      Left            =   225
      TabIndex        =   14
      Top             =   1740
      Width           =   2235
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   19
      Left            =   225
      TabIndex        =   13
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   0
      Left            =   4110
      TabIndex        =   12
      Top             =   1200
      Width           =   1110
   End
End
Attribute VB_Name = "frmVouPrefix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Master As ADODB.Recordset
Public VouType$
Dim mAdd As Boolean, mEdit As Boolean, mDel As Boolean, mPrn As Boolean
Dim ForeColorSelEnter$
Dim BackColorSelLeave$
Dim ListArray As Variant
Dim mListItem As ListItem
Dim result As Boolean
'Private Const CellBackColLeave$ = &HECE4D7    '&HECE4D7   '&HEDF7FE
'Private Const CellForeColLeave$ = &HFF00FF
'Private Const CellBackColEnter$ = &HF0D5BF
'Private Const GridBackColorBkg$ = &HE2D5C0

Private Const Description As Byte = 0
Private Const Category As Byte = 1
Private Const DivBase As Byte = 2
Private Const VType As Byte = 3
Private Const NumberMethod As Byte = 4
Private Const StkTrn As Byte = 5
Private Const SepNarr As Byte = 6
Private Const ComNarr As Byte = 7
Private Const SerialFromTable As Byte = 8
Private Const SROff As Byte = 46
'SrNo
Private Const SrNo As Byte = 1
Private Const DateFrom As Byte = 1
Private Const Prefix As Byte = 2
Private Const StartSrlNo As Byte = 3
Private Const MaxNoExists As Byte = 4
Private Const AddEdit As Byte = 5

Dim EditName$
Dim GridKey As Integer
Dim DataAddMode As Boolean
Dim TAddMode As Boolean
Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To Txt.Count - 1
If I = 2 Or I = 4 Or I = 6 Or I = 7 Then
    Txt(I).Enabled = Enb
Else
    Txt(I).Enabled = False   ' Enb
End If
    Txt(I).ForeColor = CtrlFColOrg
Next

'If TopCtrl1.TopText2 = "Edit" Then
'    Txt(SiteCode).Enabled = False
'    Txt(Vdate).Enabled = False
'    Txt(SerialNo).Enabled = False
'End If

txtDisabled_Color Me

TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol
End Sub

'* Used for clear all text boxes used in the form
Private Sub BlankText()
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I).TEXT = ""
    Next I
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End Sub
'* Used for intialize grid columns
Private Sub Grid_Ini()
'Serial No  | Date From | Prefix |Start Srl No | Max Srl No Exists
    With FGrid
'        .left = Me.left '+ 60
'        .width = Me.width - 90
'        .top = 2550
'        .BackColor = CellBackColLeave
'        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 6

        .TextMatrix(0, 0) = " "
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 420
        
        .TextMatrix(0, DateFrom) = "Date From"
        .ColAlignmentFixed(DateFrom) = flexAlignLeftCenter
        .ColAlignment(DateFrom) = flexAlignLeftCenter
        .ColWidth(DateFrom) = 1150

        .TextMatrix(0, Prefix) = "Prefix"
        .ColAlignment(Prefix) = flexAlignLeftCenter
        .ColWidth(Prefix) = 1000
        
        .TextMatrix(0, StartSrlNo) = "Start Srl No."
        .ColAlignmentFixed(StartSrlNo) = flexAlignLeftCenter
        .ColAlignment(StartSrlNo) = flexAlignRightCenter
        .ColWidth(StartSrlNo) = 1000

        .TextMatrix(0, MaxNoExists) = "Max Srl No."
        .ColAlignmentFixed(MaxNoExists) = flexAlignLeftCenter
        .ColAlignment(MaxNoExists) = flexAlignLeftCenter
        .ColWidth(MaxNoExists) = 0 '550
        
        .TextMatrix(0, AddEdit) = ""
        .ColAlignmentFixed(AddEdit) = flexAlignLeftCenter
        .ColAlignment(AddEdit) = flexAlignLeftCenter
        .ColWidth(AddEdit) = 0
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
End Sub

Private Sub MoveRec()
Dim Master1 As ADODB.Recordset
Dim Rst As ADODB.Recordset, I As Integer
On Error GoTo ELoop
TopCtrl1.tAdd = False
TopCtrl1.tDel = False
    If Master.RecordCount > 0 Then
        If Master!SearchCode = "F_AO" Or Master!SearchCode = "SXAO" Then
            TopCtrl1.tEdit = False
        Else
            TopCtrl1.tEdit = True
        End If
'        If InStr(Me.TopCtrl1.Tag, "D") <> 0 Then Me.TopCtrl1.tDel = True
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "SELECT Description,Category,Number_Method,V_Type,DivBaseNumber,StkTrn," & _
            "Number_Method,Separate_Narr,Common_Narr,SerialNo_From_Table " & _
            "FROM Voucher_Type where V_Type='" & Master!SearchCode & "'", G_FaCn, adOpenStatic, adLockReadOnly
        
        Txt(Description).TEXT = Master1!Description
        Txt(Category).TEXT = Master1!Category
        Txt(DivBase).TEXT = IIf(Master1!DivBaseNumber = 1, "Yes", "No")
        Txt(VType).TEXT = Master1!V_Type
        Txt(NumberMethod).TEXT = Master1!Number_Method
        Txt(StkTrn).TEXT = IIf(IsNull(Master1!StkTrn), "", Master1!StkTrn)
        Txt(SepNarr).TEXT = IIf(Master1!Separate_Narr = "Y", "Yes", "No")
        Txt(ComNarr).TEXT = IIf(Master1!Common_Narr = "Y", "Yes", "No")
        Txt(SerialFromTable).TEXT = IIf(IsNull(Master1!SerialNo_From_Table), "", Master1!SerialNo_From_Table)
        
        'Select Voucher No. Table
        If Txt(SerialFromTable) = "" Then 'Serial No. from Voucher_Prefix
            GSQL = "Select * from Voucher_Prefix where Voucher_Prefix.V_Type='" & Master1!V_Type & "' AND Div_Code='" & PubDivCode & "'"
'            If Master1!DivBaseNumber = 1 Then
'                GSQL = GSQL & " and Voucher_Prefix.Div_Code='" & PubDivCode & "' "
'            End If
            GSQL = GSQL & " Order By Date_From asc"
            
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open GSQL, G_FaCn, adOpenStatic, adLockReadOnly
            
        ElseIf UCase(Txt(SerialFromTable)) = UCase("VehBill_Counter") Then 'Serial No. from VehBill_Counter, FAData.mdb
            GSQL = "Select * from VehBill_Counter where VehBill_Counter.V_Type='" & Master1!V_Type & "'"
            If Master1!DivBaseNumber = 1 Then
                GSQL = GSQL & " and VehBill_Counter.Div_Code='" & PubDivCode & "' "
            End If
            GSQL = GSQL & " Order By Date_From,Prefix"
            
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open GSQL, G_FaCn, adOpenStatic, adLockReadOnly
            
        ElseIf UCase(Txt(SerialFromTable)) = UCase("SP_OrdCoun") Then 'Serial No. from SP_OrdCoun Table, Automan.mdb
            'Only for Spare Purchase Orders
            
            GSQL = "Select SP_OrdCoun.*,'' as Date_From,Start_No as Start_Srl_No from SP_OrdCoun where SP_OrdCoun.ORD_TYPE='" & Master1!V_Type & "'"
            If Master1!DivBaseNumber = 1 Then
                GSQL = GSQL & " and SP_OrdCoun.Div_Code='" & PubDivCode & "' "
            End If
            GSQL = GSQL & " Order By Ord_Type"
            
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
        End If

        FGrid.Redraw = False
        FGrid.Rows = 1
        If Rst.RecordCount > 0 Then
 
            I = 1
            Do Until Rst.EOF
                FGrid.AddItem "" 'FGrid.Rows
                With FGrid
'                    .TextMatrix(i, SrNo) = i
                    .TextMatrix(I, DateFrom) = Rst!Date_From
                    .TextMatrix(I, Prefix) = Rst!Prefix
                    .TextMatrix(I, StartSrlNo) = Rst!start_srl_no
                    .TextMatrix(I, MaxNoExists) = ""
                    .TextMatrix(I, AddEdit) = "N"
                End With
                Rst.MoveNext
                I = I + 1
            Loop
        End If
    Else
        BlankText
        FGrid.Rows = 1
    End If
    If FGrid.Rows = 1 Then FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    FGrid.Redraw = True
    Set Rst = Nothing
    Set Master1 = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If Shift = 2 And KeyCode = 83 Then
   TxtGrid_Validate 0, True
End If
    FormKeyDown Me, KeyCode, Shift
  
Exit Sub
ELoop:
    CheckError
Exit Sub

End Sub

Private Sub Form_Load()
On Error GoTo ELoop
    TopCtrl1.tAdd = False
    TopCtrl1.tDel = False
    TopCtrl1.Tag = PubUParam
    WinSetting Me
    Grid_Ini
    DataAddMode = False
   
    Set Master = G_FaCn.Execute("SELECT V_Type as SearchCode " & _
        "FROM Voucher_Type where category not in ('FA','GenFA')Order by V_Type")
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    
  Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TopCtrl1.TopText2 <> "Browse" Then
        If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Set RsItem = Nothing
'Set RsSpot = Nothing
'Set RsGroup = Nothing
Set Master = Nothing
End Sub

'** New Code
Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
'    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
'If KeyCode = vbKeyDelete And Shift = 0 Then
'    Select Case FGrid.Col
'        Case Col_Qty, Col_Rate, Col_Amt, Col_DiscPer, Col_DiscAmt
'            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
'            Amt_Cal
'    End Select
'End If
If KeyCode = vbKeyReturn Then
       Select Case FGrid.Col
            Case DateFrom
                If FGrid.TextMatrix(FGrid.Row, 0) = "" Or FGrid.TextMatrix(FGrid.Row, 0) = "**" Then
                    Call GridDblClick(Me, FGrid, TxtGrid, 0)
                    TAddMode = False
                End If
            Case Prefix
                If FGrid.TextMatrix(FGrid.Row, DateFrom) <> "" Then
                    If FGrid.TextMatrix(FGrid.Row, 0) = "" Or FGrid.TextMatrix(FGrid.Row, 0) = "**" Then
                        Call GridDblClick(Me, FGrid, TxtGrid, 0)
                        TAddMode = False
                    End If
                Else
                    FGrid.Col = DateFrom
                End If
            Case StartSrlNo
                If FGrid.TextMatrix(FGrid.Row, Prefix) <> "" Then
                    Call GridDblClick(Me, FGrid, TxtGrid, 0)
                    TAddMode = False
                Else
                    FGrid.Col = Prefix
                End If
        End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)
On Error GoTo ELoop
SetMaxLength
'If FGrid.TextMatrix(FGrid.Row, 0) <> "**" Then
'    If mEdit Then FGrid.TextMatrix(FGrid.Row, 0) = "*"
'End If
    Select Case FGrid.Col
        Case StartSrlNo
            If FGrid.TextMatrix(FGrid.Row, DateFrom) <> "" Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, True, keyascii)
            End If
        Case Prefix
            If FGrid.TextMatrix(FGrid.Row, 0) = "" Or FGrid.TextMatrix(FGrid.Row, 0) = "**" Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, keyascii)
            End If
        Case DateFrom
            If FGrid.TextMatrix(FGrid.Row, 0) = "" Or FGrid.TextMatrix(FGrid.Row, 0) = "**" Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, keyascii)
            End If
    End Select
    If keyascii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Dim I As Integer
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        MsgBox "Check for existance & then delete"
'        If FGrid.Row  >= 1 Then
'            If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
'                If FGrid.Rows  > 2 Then
'                    FGrid.RemoveItem (FGrid.Row)
'                Else
'                    FGrid.Rows = 1
'                    FGrid.AddItem FGrid.Rows
'                    FGrid.FixedRows = 1
'                End If
'                For i = 1 To FGrid.Rows - 1
'                   FGrid.TextMatrix(i, Col_SrNo) = i
'                Next
'                'Recalculate footer values
'            End If
'        Else
'            MsgBox "No Entries To Delete!", vbCritical, "Delete Module"
'        End If
        FGrid.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
End Sub



Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    FGrid.AddItem "" 'FGrid.Rows
    Txt(2).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub TopCtrl1_eFirst()
  BUTTONS True, Me, Master, 1
  Call MoveRec
End Sub

Private Sub TopCtrl1_eLast()
 BUTTONS True, Me, Master, 4
 Call MoveRec
End Sub

Private Sub TopCtrl1_eNext()
 BUTTONS True, Me, Master, 3
 Call MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
 BUTTONS True, Me, Master, 2
 Call MoveRec
End Sub

Private Sub TopCtrl1_eCancel()
Dim I As Integer
On Error GoTo ErrorLoop
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eRef()
'    RsChassis.Requery
'    rsGod.Requery
End Sub
Private Sub TopCtrl1_eSave()
    Dim I As Integer, mPubDivCode$
    Dim mTrans As Boolean
On Error GoTo errlbl
    
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, DateFrom) <> "" Then
            If FGrid.TextMatrix(I, Prefix) = "" Then
                MsgBox "Fill Prefix in Row No " & I, vbInformation, "Validation Check": FGrid.Row = I: TxtGrid(0).left = FGrid.left: Exit Sub
            End If
        End If
    Next
    
    G_FaCn.BeginTrans
    mTrans = True
    For I = 1 To FGrid.Rows - 1
        GSQL = ""
        If FGrid.TextMatrix(I, AddEdit) = "E" Then
            GSQL = "update Voucher_Prefix set Prefix='" & FGrid.TextMatrix(I, Prefix) & "',Start_Srl_No=" & Val(FGrid.TextMatrix(I, StartSrlNo)) & _
                " where V_Type ='" & Txt(VType) & "'"
            If Txt(DivBase) = "Yes" Then
                GSQL = GSQL & " and Div_Code='" & PubDivCode & "' "
            End If
            GSQL = GSQL & " and Date_From=" & ConvertDate(FGrid.TextMatrix(I, DateFrom)) & ""
        ElseIf FGrid.TextMatrix(I, DateFrom) <> "" And FGrid.TextMatrix(I, AddEdit) = "" Then
            mPubDivCode = PubDivCode
            If Txt(DivBase) = "No" Then
                If G_FaCn.Execute("Select Div_Code from Voucher_Prefix where V_Type='" & Txt(VType) & "'").RecordCount > 0 Then
                    mPubDivCode = G_FaCn.Execute("Select Div_Code from Voucher_Prefix where V_Type='" & Txt(VType) & "'").Fields(0).Value
                End If
            End If
            GSQL = "insert into Voucher_Prefix(V_Type,Div_Code,Date_From,Prefix,Start_Srl_No)  " & _
                "values('" & Txt(VType) & "','" & mPubDivCode & "'," & ConvertDate(FGrid.TextMatrix(I, DateFrom)) & ",'" & FGrid.TextMatrix(I, Prefix) & _
                "'," & Val(FGrid.TextMatrix(I, StartSrlNo)) & ")"
        End If
        If GSQL <> "" Then
            G_FaCn.Execute (GSQL)
        End If
    Next
    GSQL = "update Voucher_Type set DivBaseNumber=" & IIf(Txt(2) = "No", 0, 1) & ",Number_Method='" & Txt(4) & "', Separate_Narr='" & IIf(Txt(6) = "No", "N", "Y") & _
        "', Common_Narr='" & IIf(Txt(7) = "No", "N", "Y") & "' where V_Type='" & Txt(3) & "'"
    G_FaCn.Execute (GSQL)
'         GSQL = "update Voucher_Prefix set Date_From=" & ConvertDate(FGrid.TextMatrix(i, DateFrom)) & " , Prefix= '" & FGrid.TextMatrix(i, Prefix) & _
'                "', Start_Srl_No=" & Val(FGrid.TextMatrix(i, StartSrlNo)) & " where V_Type='" & txt(Vtype) & "'"
'
'        If GSQL <> "" Then
     '   End If
    G_FaCn.CommitTrans
    mTrans = False
    Master.Requery
    Master.FIND "SearchCode = '" & Txt(VType) & "'"
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    Exit Sub
errlbl:
    If mTrans Then G_FaCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "SELECT V_Type as SearchCode,V_Type,Description " & _
            "FROM Voucher_Type where category not in ('FA','GenFA')Order by V_Type"
    Set SearchForm = Me
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("SearchCode='" & MyValue & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    CheckError
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop


 Select Case Index
        Case NumberMethod
           ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 600
 End Select
 If KeyCode = 13 Or KeyCode = vbKeyTab Then
    SendKeysA vbKeyTab, True
End If
Exit Sub
ELoop:
    CheckError
End Sub


Private Sub Txt_Change(Index As Integer)
If Index = NumberMethod Then
    If DataAddMode = False And Txt(VType) <> "F_AO" And Txt(VType) <> "SXAO" Then
'        Call GridDblClick(Me, FGrid, TxtGrid, 0)
    End If
End If
End Sub
Private Sub Grid_Hide()
    If ListView.Visible = True Then ListView.Visible = False
   
End Sub
Public Sub Ctrl_GetFocus(Ctrl As Object)
    Ctrl.BackColor = CtrlBCol
    Ctrl.ForeColor = CtrlFCol
    Ctrl.SelStart = Len(Ctrl)
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus Txt(Index)
TxtGrid(0).Visible = False
Grid_Hide
Select Case Index
    Case DivBase
        ListArray = Array("Yes", "No")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 2)
    Case NumberMethod
        ListArray = Array("Automatic", "Manual")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 2)
    Case SepNarr
        ListArray = Array("Yes", "No")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 2)
    Case ComNarr
        ListArray = Array("Yes", "No")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 2)

End Select

End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
On Error GoTo ELoop
Select Case Index
        Case DivBase
            If keyascii = 89 Or keyascii = 121 Then         ' Y/y
                Txt(Index).TEXT = "Yes"
                keyascii = 0
            ElseIf keyascii = 78 Or keyascii = 110 Then    ' N/n
                Txt(Index).TEXT = "No"
                keyascii = 0
            End If
            
        Case SepNarr
            If keyascii = 89 Or keyascii = 121 Then         ' Y/y
                Txt(Index).TEXT = "Yes"
                keyascii = 0
            ElseIf keyascii = 78 Or keyascii = 110 Then     ' N/n
                Txt(Index).TEXT = "No"
                keyascii = 0
            End If

        Case ComNarr
            If keyascii = 89 Or keyascii = 121 Then         ' Y/y
                Txt(Index).TEXT = "Yes"
                keyascii = 0
            ElseIf keyascii = 78 Or keyascii = 110 Then     ' N/n
                Txt(Index).TEXT = "No"
                keyascii = 0
            End If

      End Select
     
   
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
'    Select Case FGrid.Col
'        Case SpotName
'            If DgSpot.Visible = False Then DgSpot.left = TxtGrid(0).left: DgSpot.top = TxtGrid(0).top + TxtGrid(0).Height + 20
'        Case ItemName
'            If DGItem.Visible = False Then DGItem.left = TxtGrid(0).left: DGItem.top = TxtGrid(0).top + TxtGrid(0).Height + 20
'            If RsItem.RecordCount = 0 Or (RsItem.EOF = True Or RsItem.BOF = True) Or FGrid.TextMatrix(FGrid.Row, ItemName) = "" Then Exit Sub
'            If FGrid.TextMatrix(FGrid.Row, ItemName) <> RsItem!Name Then
'                RsItem.MoveFirst
'                RsItem.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, ItemName) & "'"
'            End If
'        Case ItemGroup
'            If DGGroup.Visible = False Then DGGroup.left = TxtGrid(0).left: DGGroup.top = TxtGrid(0).top + TxtGrid(0).Height + 20
'            If RsGroup.RecordCount = 0 Or (RsGroup.EOF = True Or RsGroup.BOF = True) Or FGrid.TextMatrix(FGrid.Row, ItemGroup) = "" Then Exit Sub
'            If FGrid.TextMatrix(FGrid.Row, ItemGroup) <> RsGroup!Name Then
'                RsItem.MoveFirst
'                RsItem.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, ItemGroup) & "'"
'            End If
'    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then TxtGrid(0).TEXT = TxtGrid(0).Tag: Exit Sub
    Select Case FGrid.Col
        Case DateFrom, Prefix, StartSrlNo
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, StartSrlNo, , , True
                End If
            End If
    End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If keyascii = vbKeyEscape Then Exit Sub
Call CheckQuote(keyascii)
Select Case FGrid.Col
    Case StartSrlNo
        Call NumPress(TxtGrid(Index), keyascii, 8, 0)
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
    Exit Sub
End If
Select Case FGrid.Col
    Case DateFrom, Prefix, StartSrlNo
'        MsgBox "Check for max no."
End Select
End Sub


Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop

Cancel = Not TxtGridLeave(Index, True)
'TxtGrid(0).Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim I As Integer
Select Case FGrid.Col
    Case DateFrom
        TxtGrid(0) = RetDate(TxtGrid(0))
        If TxtGrid(0) <> "" Then
            For I = 1 To FGrid.Rows - 1
                If FGrid.TextMatrix(I, DateFrom) = TxtGrid(0) And I <> FGrid.Row Then
                    MsgBox "Voucher Prefix Already Exist For This Date", vbInformation, "Validation Check"
                    TxtGridLeave = False: Exit Function
                End If
            Next
'            GSQL = "select * from Voucher_Prefix where V_Type ='" & txt(Vtype) & "' and Div_Code='" & PubDivCode & "' and Date_From=#" & TxtGrid(0) & "#"
'            If G_FACN.Execute(GSQL).RecordCount  > 0 Then
'                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
'            Else
'                MsgBox "Voucher Prefix Already Exist For This Date", vbInformation, "Validation Check"
'                TxtGridLeave = False: Exit Function
'            End If
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
            CellFontColor FGrid: FGridAddEditDel
            FGrid.TextMatrix(FGrid.Row, 0) = "**"
        End If
    Case Prefix
        If TxtGrid(0) <> "" Then
            For I = 1 To FGrid.Rows - 1
                If FGrid.TextMatrix(FGrid.Row, DateFrom) = FGrid.TextMatrix(I, DateFrom) And _
                    FGrid.TextMatrix(I, Prefix) = TxtGrid(0) And I <> FGrid.Row Then
                    MsgBox "Voucher Date + Prefix combination already Exist", vbInformation, "Validation Check"
                    TxtGridLeave = False: Exit Function
                End If
            Next
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
            CellFontColor FGrid: FGridAddEditDel
            FGrid.TextMatrix(FGrid.Row, 0) = "**"
        End If
    Case StartSrlNo
        If Val(TxtGrid(0).TEXT) < Val(FGrid.TextMatrix(FGrid.Row, FGrid.Col)) Then
            MsgBox "You Can't Fill Less Value than the existing Value"
            TxtGridLeave = False: Exit Function
        End If
        If (FGrid.TextMatrix(FGrid.Row, FGrid.Col)) <> (TxtGrid(0)) Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0")
            If FGrid.TextMatrix(FGrid.Row, 0) = "" Then
                FGrid.TextMatrix(FGrid.Row, 0) = "*"
            End If
            TxtGrid(0).Visible = False
            CellFontColor FGrid: FGridAddEditDel
        End If
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
End Function

'******* Fuctions **********
Private Sub SetMaxLength()
Dim mMaxLength As Integer, mAlignment As Byte, mHeight As Integer
mHeight = FGrid.RowHeight(0)
'Alignment :  0-mLeft  1-mRright   2-mCenter
    Select Case FGrid.Col
        Case Prefix
            If Txt(Category) = "FA" Then
                mMaxLength = 4
            Else
                mMaxLength = 5
            End If
    End Select
    TxtGrid(0).MaxLength = mMaxLength
    TxtGrid(0).Alignment = mAlignment
End Sub

Private Sub FGridAddEditDel()
If FGrid.TextMatrix(FGrid.Row, AddEdit) <> "" Then FGrid.TextMatrix(FGrid.Row, AddEdit) = "E"
End Sub

Private Sub CellFontColor(FG As MSHFlexGrid)
FG.CellForeColor = vbRed ' CellForeColLeave
End Sub

