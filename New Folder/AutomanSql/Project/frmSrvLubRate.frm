VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmSrvLubRate 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Service Rate/Lube Qty. Declaration"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   " "
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
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
      Index           =   6
      Left            =   3720
      MaxLength       =   12
      TabIndex        =   6
      Text            =   "29/MAR/2003"
      Top             =   1155
      Width           =   1320
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
      Left            =   7770
      TabIndex        =   3
      Top             =   570
      Width           =   1335
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
      Index           =   5
      Left            =   3705
      TabIndex        =   5
      Top             =   885
      Width           =   1335
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
      Index           =   4
      Left            =   9720
      MaxLength       =   12
      TabIndex        =   4
      Text            =   "29/MAR/2003"
      Top             =   570
      Width           =   1320
   End
   Begin MSDataGridLib.DataGrid DGlabS 
      Height          =   4440
      Left            =   9300
      Negotiate       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6225
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
         Caption         =   "Service Description"
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
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3825.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGLabM 
      Height          =   5250
      Left            =   930
      Negotiate       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6630
      Visible         =   0   'False
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   9260
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Code"
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
      BeginProperty Column02 
         DataField       =   "Vehicle_Type"
         Caption         =   "Veh.Type"
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
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5850.142
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1154.835
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   4845
      Left            =   120
      TabIndex        =   11
      Top             =   1545
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   8546
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   3
      Cols            =   3
      FixedRows       =   2
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      GridColorFixed  =   32896
      FocusRect       =   0
      Appearance      =   0
      FormatString    =   "ddd"
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
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   2
      Left            =   2175
      MaxLength       =   25
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   270
      Visible         =   0   'False
      Width           =   720
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
      Left            =   4275
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   570
      Width           =   765
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
      Left            =   1260
      TabIndex        =   1
      Top             =   570
      Width           =   3000
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Effective From Date(s):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   5
      Left            =   315
      TabIndex        =   16
      Top             =   885
      Width           =   2025
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sold Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   15
      Top             =   885
      Width           =   900
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Circular/Reference No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   0
      Left            =   5670
      TabIndex        =   14
      Top             =   570
      Width           =   2040
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Effective From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   3
      Left            =   1650
      TabIndex        =   13
      Top             =   1155
      Width           =   2010
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   9165
      TabIndex        =   12
      Top             =   570
      Width           =   435
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   4
      Left            =   315
      TabIndex        =   9
      Top             =   570
      Width           =   690
   End
End
Attribute VB_Name = "frmSrvLubRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset
Dim RstLabS As ADODB.Recordset, rstmodel As Recordset
Dim mFlag As Byte
Private Const Serv = 0
Private Const ServType = 1
Private Const RNo As Byte = 3
Private Const RDate As Byte = 4
Private Const SDate As Byte = 5
Private Const SEDate As Byte = 6

Dim GridKey As Integer
'Dim Gtf As Boolean
' Col Declaration
'Dim ExitCtrl As Boolean
Private Const SrNo As Byte = 0
Private Const VehType As Byte = 1
Private Const Model As Byte = 2
Private Const Mat As Byte = 3 '4
Private Const Lab As Byte = 4 ' 5
Private Const Engi As Byte = 5 ' 6
Private Const Gear As Byte = 6 ' 7
Private Const FAx As Byte = 7 ' 8
Private Const RAx As Byte = 8 '9
Private Const Stee As Byte = 9 '10
Private Const WhCat As Byte = 10 '13
Dim TAddMode As Boolean
Private Const CellBackColEnter   As String = &HF0D5BF    '&HFFC0C0
Private Const CellBackColLeave As String = &HFFFFFF    '&HBAD3C9
Private Const BackColorSelEnter As String = &HFEE0FD

Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Sub DGlabM_Click()
Fill_Data
Txt(2).Visible = False
DGLabM.Visible = False
End Sub

Private Sub DGLabs_Click()
FillModels
Txt(RNo).SetFocus
DGlabS.Visible = False
End Sub

Private Sub FGrid_Click()
If DGLabM.Visible Then DGLabM.Visible = False
If DGlabS.Visible Then DGlabS.Visible = False
Txt(2).Visible = False
'DGlabM_Click
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_EnterCell()
FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
'    FGrid.BackColorSel = BackColorSelEnter
'    FGrid.ForeColorSel = ForeColorSelEnter
    FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)
Select Case FGrid.Col
    Case Model
        Call Get_Text(Me, FGrid, Txt, 2, False, keyascii)
'    Case SDate, SEDate, RNo, RDate
'        If Len(FGrid.TextMatrix(FGrid.Row, Model)) <> 0 Then
'            Call Get_Text(Me, FGrid, Txt, 2, False, KeyAscii)
'        End If
    Case Mat, Lab, Engi, Gear, FAx, RAx, Stee, RNo
        If Len(FGrid.TextMatrix(FGrid.Row, Model)) <> 0 Then
            Call Get_Text(Me, FGrid, Txt, 2, True, keyascii)
        End If
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
            If FGrid.Rows > FGrid.FixedRows Then
                FGrid.RemoveItem (FGrid.Row)
            Else
                FGrid.Rows = 2
                FGrid.AddItem FGrid.Rows - 1
                FGrid.FixedRows = 2
            End If
            For I = 2 To FGrid.Rows - 1
                FGrid.TextMatrix(I, 0) = I - 1
            Next
            FGrid_EnterCell
         End If
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
Exit Sub
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_LostFocus()
'    FGrid.BackColorSel = BackColorSelLeave
'    FGrid.ForeColorSel = FGrid.ForeColor
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_Scroll()
    Txt(2).Visible = False
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
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
WinSetting Me: Ini_Grid: TopCtrl1.Tag = PubUParam

Set rstmodel = GCn.Execute("SELECT Model AS Code,Model_Desc AS Name,Vehicle_Type,Wheel_Catg from model Order By Model")

Set RstMain = New ADODB.Recordset
'RstMain.Open "Select max(Service_Type.Serv_Desc)as SearchCode, max(Service_Type.Serv_Desc)as ServDesc,max(Service_Rates.Serv_Type)as ServType,max(Model) as model " & _
    "From (Service_Rates left join Service_Type on Service_Rates.Serv_Type=Service_Type.Serv_Type) " & _
    "group by Service_Rates.Serv_Type,Supplier_RefNo", GCn, adOpenDynamic, adLockOptimistic
If PubMoveRecYn Then
    RstMain.Open "Select max(SrvT.Serv_Desc)as SearchCode, max(SrvT.Serv_Desc)as ServDesc,max(SrvR.Serv_Type)as ServType,max(SrvR.Supplier_RefNo) as SupplierRefNo, " & _
        "max(SrvR.Supplier_RefDate) as SupplierRefDate, max(SrvR.Sold_Date) as SoldDate, max(SrvR.Serv_EffectiveDate) as ServEffectiveDate  " & _
        "From (Service_Rates SrvR left join Service_Type SrvT on SrvR.Serv_Type=SrvT.Serv_Type) " & _
        "group by SrvR.Serv_Type,SrvR.Supplier_RefNo", GCn, adOpenDynamic, adLockOptimistic
Else
    Set RstMain = GCn.Execute("Select Top 1 max(SrvT.Serv_Desc)as SearchCode, max(SrvT.Serv_Desc)as ServDesc,max(SrvR.Serv_Type)as ServType,max(SrvR.Supplier_RefNo) as SupplierRefNo, " & _
        "max(SrvR.Supplier_RefDate) as SupplierRefDate, max(SrvR.Sold_Date) as SoldDate, max(SrvR.Serv_EffectiveDate) as ServEffectiveDate  " & _
        "From (Service_Rates SrvR left join Service_Type SrvT on SrvR.Serv_Type=SrvT.Serv_Type) " & _
        "group by SrvR.Serv_Type,SrvR.Supplier_RefNo")

End If

Set RstHelp = New ADODB.Recordset
RstHelp.Open "Select Serv_Type as Code, Serv_Desc as Name FROM Service_Type Order by Serv_Desc", GCn, adOpenDynamic, adLockOptimistic
Set DGlabS.DataSource = RstHelp
Disp_Text SETS("INI", Me, RstMain)
'CtrlClckCol
MoveRec
ADDFLAG = 0:    mFlag = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If TopCtrl1.TopText2 <> "Browse" Then
    If MsgBox("Do you want to exit", vbExclamation + vbYesNo) = vbYes Then
        Exit Sub
    Else
        Cancel = 1
    End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RstMain = Nothing: Set RstHelp = Nothing: Set RstLabS = Nothing
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
Dim result As Boolean
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) And TopCtrl1.TopText2 = "Add" Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
Select Case FGrid.Col
'    Case SDate, SEDate, RNo, RDate
'        If KeyCode = vbKeyDelete And Shift = 0 Then Call Get_Text(Me, FGrid, Txt, 2, False, 48)        'FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    Case Mat, Lab, Engi, Gear, FAx, RAx, Stee
        If KeyCode = vbKeyDelete And Shift = 0 Then Call Get_Text(Me, FGrid, Txt, 2, True, 48)       'FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
End Select
If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case Model, Mat, Lab, Engi, Gear, FAx, RAx, Stee
            Call GridDblClick(Me, FGrid, Txt, 2)
            TAddMode = False
'        Case SDate, SEDate, RNo, RDate
'            Call GridDblClick(Me, FGrid, Txt, 2)
'            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo Errloop
BlankText
Disp_Text SETS("ADD", Me, RstMain)
Txt(Serv).Tag = Txt(Serv)
Txt_GotFocus Serv
ADDFLAG = 1

Txt(Serv).SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
Dim Rs As Recordset
On Error GoTo Errloop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(Serv).Enabled = False
    Txt(RNo).Enabled = False

    ADDFLAG = 2
    FGrid.AddItem FGrid.Rows - 1
    FGrid.SetFocus
    FGrid_EnterCell
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
            GCn.Execute ("delete from Service_Rates where Serv_Type= '" & Trim(Txt(ServType)) & "' and Supplier_RefNo='" & Txt(RNo) & "'")
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
'    GSQL = "Select max(Service_Type.Serv_Desc) as searchcode, max(Service_Type.Serv_Desc)as Service_Type from (Service_Rates left join Service_Type on Service_Rates.Serv_Type=Service_Type.Serv_Type) group by Service_Rates.Serv_Type"
    GSQL = "Select max(SrvT.Serv_Type) + max(SrvR.Supplier_RefNo) as SearchCode, max(SrvT.Serv_Desc) as ServDesc,max(SrvR.Serv_Type)as ServType,max(SrvR.Supplier_RefNo) as SupplierRefNo, " & _
    "max(SrvR.Supplier_RefDate) as SupplierRefDate, max(SrvR.Sold_Date) as SoldDate, max(SrvR.Serv_EffectiveDate) as ServEffectiveDate  " & _
    "From (Service_Rates SrvR left join Service_Type SrvT on SrvR.Serv_Type=SrvT.Serv_Type) " & _
    "group by SrvR.Serv_Type,SrvR.Supplier_RefNo"
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
        Set RstMain = GCn.Execute("Select max(SrvT.Serv_Desc) as SearchCode, max(SrvT.Serv_Desc)as ServDesc,max(SrvR.Serv_Type)as ServType,max(SrvR.Supplier_RefNo) as SupplierRefNo, " & _
            "max(SrvR.Supplier_RefDate) as SupplierRefDate, max(SrvR.Sold_Date) as SoldDate, max(SrvR.Serv_EffectiveDate) as ServEffectiveDate  " & _
            "From (Service_Rates SrvR left join Service_Type SrvT on SrvR.Serv_Type=SrvT.Serv_Type) Where max(SrvT.Serv_Desc) = '" & MyValue & "' " & _
            "group by SrvR.Serv_Type,SrvR.Supplier_RefNo")
    End If
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim I As Integer, mQRY$, mRepName$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER

    mRepName = "SrvLubRate"
    mQRY = "SELECT ST.Serv_Desc,ST.FreeServCode, SR.* " & _
        " from Service_Rates SR left join Service_Type ST on SR.Serv_Type=ST.Serv_Type " & _
        " Where SR.Serv_Type='" & Txt(ServType) & _
        "' and SR.Supplier_RefNo='" & Txt(RNo) & "' Order by SR.Model,SR.Sold_Date"
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
    FGrid_LeaveCell
    Txt(2).Visible = False
    transFlag = 0
    If IsValid(Txt(Serv), "Service Details") = False Then Exit Sub
    If IsValid(Txt(RNo), Lbl(0)) = False Then Exit Sub
    If IsValid(Txt(RNo), Lbl(1)) = False Then Exit Sub
    If IsValid(Txt(SDate), Lbl(2)) = False Then Exit Sub
    If IsValid(Txt(SEDate), Lbl(3)) = False Then Exit Sub

'    For i = 2 To FGrid.Rows - 1
'        If Len(FGrid.TextMatrix(i, Model)) <> 0 Then
'            If Len(FGrid.TextMatrix(i, SDate)) = 0 Then
'                MsgBox "Sold Date is Required", vbCritical, "Validation"
'                FGrid.Row = i: FGrid.Col = SDate: FGrid.SetFocus: FGrid_EnterCell
'                Exit Sub
'            End If
'            If Len(FGrid.TextMatrix(i, SEDate)) = 0 Then
'                MsgBox "Service Effective Date is Required", vbCritical, "Validation"
'                FGrid.Row = i: FGrid.Col = SEDate: FGrid.SetFocus: FGrid_EnterCell
'                Exit Sub
'            End If
'            If CDate(FGrid.TextMatrix(i, SDate))  > CDate(FGrid.TextMatrix(i, SEDate)) Then
'                MsgBox "Sold Date " & FGrid.TextMatrix(i, SDate) & " is greater than " & vbCrLf & "Service Effective Date " & FGrid.TextMatrix(i, SEDate), vbCritical, "Validation"
'                FGrid.Row = i: FGrid.Col = SEDate: FGrid.SetFocus: FGrid_EnterCell
'                Exit Sub
'            End If
'            If Len(FGrid.TextMatrix(i, RNo)) = 0 Then
'                MsgBox "Reference Number is Required", vbCritical, "Validation"
'                FGrid.Row = i: FGrid.Col = RNo: FGrid.SetFocus: FGrid_EnterCell
'                Exit Sub
'            End If
'        End If
'    Next
    GCn.BeginTrans
    transFlag = 1
    GCn.Execute ("DELETE From Service_Rates Where Serv_Type='" & Txt(ServType) & "' and Supplier_RefNo='" & Txt(RNo) & "'")
    For I = 2 To FGrid.Rows - 1
        If Len(FGrid.TextMatrix(I, Model)) <> 0 Then
            GCn.Execute ("Insert Into Service_Rates(Serv_Type,MODEL,Sold_Date,Serv_EffectiveDate,Site_Code,Mat_Amt," & _
                "Lab_Amt,Eng_Oil,Gear_Oil,F_Axel_Oil,R_Axel_Oil,Steer_Oil,Supplier_RefNo,Supplier_RefDate,U_Name,U_EntDt,U_AE,chrg_from) " & _
                "Values('" & Txt(ServType) & "','" & FGrid.TextMatrix(I, Model) & "'," & ConvertDate(Txt(SDate)) & _
                "," & ConvertDate(Txt(SEDate)) & ",'" & PubSiteCode & "'," & Val(FGrid.TextMatrix(I, Mat)) & _
                "," & Val(FGrid.TextMatrix(I, Lab)) & "," & Val(FGrid.TextMatrix(I, Engi)) & "," & Val(FGrid.TextMatrix(I, Gear)) & _
                "," & Val(FGrid.TextMatrix(I, FAx)) & "," & Val(FGrid.TextMatrix(I, RAx)) & "," & Val(FGrid.TextMatrix(I, Stee)) & _
                ",'" & Txt(RNo) & "'," & ConvertDate(Txt(RDate)) & ",'" & pubUName & _
                "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2.CAPTION, 1) & "','A')")
        End If
    Next
    GCn.CommitTrans
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select max(SrvT.Serv_Desc) as SearchCode, max(SrvT.Serv_Desc)as ServDesc,max(SrvR.Serv_Type)as ServType,max(SrvR.Supplier_RefNo) as SupplierRefNo, " & _
        "max(SrvR.Supplier_RefDate) as SupplierRefDate, max(SrvR.Sold_Date) as SoldDate, max(SrvR.Serv_EffectiveDate) as ServEffectiveDate  " & _
        "From (Service_Rates SrvR left join Service_Type SrvT on SrvR.Serv_Type=SrvT.Serv_Type) Where max(SrvT.Serv_Desc) = '" & Txt(ServType) & Txt(RNo) & "' " & _
        "group by SrvR.Serv_Type,SrvR.Supplier_RefNo")
    End If
    RstHelp.Requery
    RstMain.FIND ("SearchCode='" & Txt(ServType) & Txt(RNo) & "'")
    Disp_Text SETS("INI", Me, RstMain)
    If ADDFLAG = 1 Then
        FGrid.CellBackColor = CellBackColLeave
    Else
        MoveRec
'        CtrlClckCol
        ADDFLAG = 0
        DGLabM.Visible = False
    End If
Exit Sub
Errloop:    If transFlag = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, RstMain)
        Me.ActiveControl.SetFocus
        MoveRec
'        CtrlClckCol
        DGlabS.Visible = False
        DGLabM.Visible = False
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    Txt(Serv).BackColor = CtrlBColOrg:      Txt(Serv).ForeColor = CtrlFColOrg
    Txt(ServType).BackColor = CtrlBColOrg:      Txt(ServType).ForeColor = CtrlFColOrg
End Sub

Private Sub MoveRec()
Dim Rs As Recordset
On Error GoTo Errloop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    Txt(Serv) = IIf(IsNull(RstMain!ServDesc), "", RstMain!ServDesc)
    Txt(ServType) = IIf(IsNull(RstMain!ServType), "", RstMain!ServType)
    Txt(RNo) = IIf(IsNull(RstMain!SupplierRefNo), "", RstMain!SupplierRefNo)
    Txt(RDate) = IIf(IsNull(RstMain!SupplierRefDate), "", RstMain!SupplierRefDate)
    Txt(SDate) = IIf(IsNull(RstMain!SoldDate), "", RstMain!SoldDate)
    Txt(SEDate) = IIf(IsNull(RstMain!ServEffectiveDate), "", RstMain!ServEffectiveDate)
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT M.Vehicle_Type,SrvR.Model,Mat_Amt,Lab_Amt,Eng_Oil,Gear_Oil,F_Axel_Oil,R_Axel_Oil,Steer_Oil " & _
                "from Service_Rates SrvR left Join Model M on SrvR.Model=M.Model " & _
                "WHERE SrvR.Serv_Type='" & Txt(ServType) & "' and SrvR.Supplier_RefNo='" & RstMain!SupplierRefNo & _
                "' Order By M.Vehicle_Type,SrvR.Model")
    If Rs.RecordCount > 0 Then
        FGrid.Rows = 2
        Do Until Rs.EOF
'            FGrid.AddItem FGrid.Rows - 1 & Chr(9) & rs!Model & Chr(9) & rs!sold_date & Chr(9) & rs!Serv_EffectiveDate & Chr(9) & IIf(rs!Mat_Amt <> 0, Format(rs!Mat_Amt, "0.00"), "") & Chr(9) & IIf(rs!Lab_Amt <> 0, Format(rs!Lab_Amt, "0.00"), "") & Chr(9) & IIf(rs!Eng_Oil <> 0, Format(rs!Eng_Oil, "0.00"), "") & Chr(9) & IIf(rs!Gear_Oil <> 0, Format(rs!Gear_Oil, "0.00"), "") & Chr(9) & IIf(rs!F_Axel_Oil <> 0, Format(rs!F_Axel_Oil, "0.00"), "") & Chr(9) & IIf(rs!R_Axel_Oil <> 0, Format(rs!R_Axel_Oil, "0.00"), "") & Chr(9) & IIf(rs!Steer_Oil <> 0, Format(rs!Steer_Oil, "0.00"), "") & Chr(9) & rs!Supplier_RefNo & Chr(9) & rs!Supplier_RefDate & Chr(9) & rs!R_Axel_Oil
            FGrid.AddItem FGrid.Rows - 1 & Chr(9) & Rs!Vehicle_Type & Chr(9) & Rs!Model & Chr(9) & IIf(Rs!Mat_Amt <> 0, Format(Rs!Mat_Amt, "0.00"), "") & Chr(9) & IIf(Rs!Lab_Amt <> 0, Format(Rs!Lab_Amt, "0.00"), "") & Chr(9) & IIf(Rs!Eng_Oil <> 0, Format(Rs!Eng_Oil, "0.00"), "") & Chr(9) & IIf(Rs!Gear_Oil <> 0, Format(Rs!Gear_Oil, "0.00"), "") & Chr(9) & IIf(Rs!F_Axel_Oil <> 0, Format(Rs!F_Axel_Oil, "0.00"), "") & Chr(9) & IIf(Rs!R_Axel_Oil <> 0, Format(Rs!R_Axel_Oil, "0.00"), "") & Chr(9) & IIf(Rs!Steer_Oil <> 0, Format(Rs!Steer_Oil, "0.00"), "") & Chr(9) & Rs!R_Axel_Oil
            Rs.MoveNext
        Loop
    Else
        FGrid.Rows = 2
        FGrid.AddItem FGrid.Rows - 1
    End If
    FGrid.FixedRows = 2
End If
Set Rs = Nothing
FGrid.Col = Model
If FGrid.Visible And FGrid.Enabled Then FGrid.SetFocus
FGrid_EnterCell
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
If DGLabM.Visible Then DGLabM.Visible = False
If DGlabS.Visible Then DGlabS.Visible = False
If Index <> 2 Then Ctrl_GetFocus Txt(Index)
If Index = 0 Then
    Txt(Index).Tag = Txt(Index).TEXT
ElseIf Index = 2 Then
    FGrid.CellBackColor = CellBackColLeave
    Txt(2).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
        Case Model
            Txt(2).MaxLength = 15
            TStr = ""
            Do Until mROW = FGrid.Rows - 1
                mROW = mROW + 1
                If mROW <> FGrid.Row Then
                    TStr = TStr + "'" + FGrid.TextMatrix(mROW, Model) + "'" + ","
                End If
            Loop
            Set rstmodel = GCn.Execute("SELECT model AS CODE,model_DESC AS NAME ,Vehicle_Type,wheel_catg from model where model NOT in (" & TStr & ")")
            rstmodel.Sort = "CODE"
            Set DGLabM.DataSource = rstmodel
            rstmodel.FIND "code  >='" & FGrid.TextMatrix(FGrid.Row, Model) & "'"
            If rstmodel.RecordCount > 0 Then
                If rstmodel.EOF = True Then rstmodel.MoveFirst
            End If
        Case SDate, SEDate, RDate
            Txt(2).MaxLength = 11
        Case RNo
            Txt(2).MaxLength = 25
        Case Mat, Lab
            Txt(2).MaxLength = 11
        Case Engi, Gear, FAx, RAx, Stee
            Txt(2).MaxLength = 8
    End Select
End If
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean
Dim I As Byte
Dim Txtdate As Boolean
Select Case Index
    Case Serv
        If KeyCode <> vbKeyEscape Then
            DGridTxtKeyDown DGlabS, Txt, Serv, RstHelp, KeyCode, False, 1
        End If
    Case 2
        If KeyCode = vbKeyEscape Then
            Txt(Index).TEXT = Txt(Index).Tag
            'Txt_KeyUp Index, KeyCode, Shift
            Txt(Index).Visible = False
            DGLabM.Visible = False
            Exit Sub
        End If
        Select Case FGrid.Col
            Case Model    '1
                If DGLabM.Visible = False Then DGridColSwap DGLabM, 0
                DGridTxtKeyDown DGLabM, Txt, 2, rstmodel, KeyCode, True, 0
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, Txt, 2, KeyCode, TAddMode, WhCat
                    End If
                End If
'            Case SDate, SEDate, RNo, RDate
'                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
'                    If TxtGridLeave = True Then
'                         GridTxtDown FGrid, Txt, 2, KeyCode, TAddMode, WhCat ', 3
'                         FGrid.CellBackColor = CellBackColLeave
'                    End If
'                End If
            Case Mat, Lab, Engi, Gear, FAx, RAx, Stee
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, Txt, 2, KeyCode, TAddMode, WhCat ', 3
                         FGrid.CellBackColor = CellBackColLeave
                    End If
                End If
        End Select
End Select

If Index <> 2 And DGlabS.Visible = False Then
        '' KEY DOWN
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        ' KEY UP
        If TopCtrl1.TopText2 = "Add" Then
            If Index <> Serv Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2 = "Edit" Then
            If Index <> RNo Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(keyascii)
Select Case Index
    Case Serv
        If DGlabS.Visible = True Then DGridTxtKeyPress Txt, Serv, RstHelp, keyascii, "name"
    Case 2
        Select Case FGrid.Col
            Case Model
                If rstmodel.RecordCount > 0 Then DGridTxtKeyPress Txt, Index, rstmodel, keyascii, "CODE"
            Case Engi, Gear, FAx, RAx, Stee
                Call NumPress(Txt(2), keyascii, 4, 3)
            Case Mat
                Call NumPress(Txt(2), keyascii, 8, 2)
            Case Lab
                Call NumPress(Txt(2), keyascii, 7, 2)
        End Select
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case Serv
        If DGlabS.Visible Then Txt(ServType).TEXT = DGlabS.TEXT
    Case 2
        Select Case FGrid.Col
            Case Model
                If KeyCode <> 13 And DGLabM.Visible = False Then Txt_KeyDown Index, GridKey, 0: DGridTxtKeyPress Txt, Index, rstmodel, KeyCode, "CODE", True
'            Case RNo
'                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Txt(Index).Text
            Case Mat, Lab, Engi, Gear, FAx, RAx, Stee
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(Val(Txt(Index)) <> 0, Format(Val(Txt(Index)), "0.00"), "")
        End Select
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
If Index = Serv Then DGLabs_Click
If Index <> 2 Then Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rs As Recordset
Select Case Index
    Case Serv
        Set Rs = GCn.Execute("Select max(Serv_Type) as serv From Service_Rates where Serv_Type='" & Txt(ServType) & "' group by Serv_Type")
        If ADDFLAG = 1 Then
           If Not Rs.EOF Then MsgBox "Service Already Exists", vbInformation, "Validation": Txt(Serv) = "": Txt(ServType) = "": Cancel = True: Set Rs = Nothing: Exit Sub
        End If
        If RstHelp.RecordCount = 0 Then Exit Sub
        If DGlabS.Visible Then
            FillModels
        End If
        
    Case SDate, SEDate, RDate
        If Txt(Index) = "" Then
           Txt(Index) = PubLoginDate
        Else
            Txt(Index) = RetDate(Txt(Index))
        End If

   Case 2
        Cancel = Not TxtGridLeave(Index, True)
End Select
Set Rs = Nothing
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid.Col
    Case Model
        If rstmodel.RecordCount = 0 Then
            TxtGridLeave = False: DGLabM.Visible = False: Exit Function
        End If
        If FGrid.TextMatrix(FGrid.Row, Model) <> Txt(2) Then Call Fill_Data
        If FGrid.TextMatrix(FGrid.Rows - 1, Model) <> "" Then FGrid.AddItem FGrid.Rows
'    Case SDate, SEDate, RDate
'        If Len(Trim(Txt(2).Text)) = 0 Then
'           Txt(2).Text = PubLoginDate
'        Else
'            Txt(2).Text = RetDate(Txt(2))
'        End If
'        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Txt(2).Text
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    Txt(2).Visible = False
End If
End Function

Private Sub DGLab_GotFocus()
    mFlag = 1
End Sub

Private Sub DGLabM_GotFocus()
    mFlag = 1
End Sub

Private Sub BlankText()
Dim I As Byte
    Txt(0).TEXT = ""
    Txt(ServType).TEXT = ""
    Txt(RNo) = ""
    Txt(RNo).Tag = ""
    Txt(RDate) = ""
    Txt(SDate) = ""
    Txt(SEDate) = ""

    Txt(2).TEXT = ""
    FGrid.Rows = 2
    FGrid.AddItem FGrid.Rows - 1
    FGrid.FixedRows = 2
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    Txt(0).Enabled = Enb
    Txt(1).Enabled = False
    Txt(3).Enabled = Enb
    Txt(4).Enabled = Enb
    Txt(5).Enabled = Enb
    Txt(6).Enabled = Enb
    
    Txt(2).Enabled = Enb
'    FGrid.Enabled = Enb
    Txt(2).BackColor = CtrlBCol
    Txt(2).ForeColor = CtrlFCol
End Sub

Private Sub Ini_Grid()
    With DGLabM
        .top = mTopScale '390
        .Columns(0).width = 1995.024
        .Columns(1).width = 6089.953
        .Columns(2).width = 1110.047
        .width = .Columns(0).width + .Columns(1).width + .Columns(2).width + 600
        .left = Me.width - (.width + 100)
        .height = .RowHeight * 25
    End With

    With DGlabS
        .top = mTopScale '390
        .Columns(0).width = 750.0473
        .Columns(1).width = 2819.906
        .width = .Columns(0).width + .Columns(1).width + 600
        .left = Me.width - (.width + 100)
        .height = .RowHeight * 10
    End With

    With FGrid
        .left = Me.left '+ 45
        .width = Me.width - 120
        .height = .RowHeight(0) * 22
        .top = 1545
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 11 '14
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
        
        .ColAlignmentFixed = flexAlignCenterCenter
        .MergeCol(SrNo) = True
        .TextMatrix(0, SrNo) = "SNo."
        .TextMatrix(1, SrNo) = .TextMatrix(0, SrNo)
        .ColAlignment(SrNo) = flexAlignRightCenter
        .ColWidth(SrNo) = 600
        
        .MergeCol(Model) = True
        .TextMatrix(0, Model) = "Model"
        .TextMatrix(1, Model) = "Model"
        .ColAlignment(Model) = flexAlignLeftCenter
        .ColWidth(Model) = 1635
        .ColAlignmentFixed(Model) = flexAlignLeftCenter
        
        .TextMatrix(0, VehType) = "Vehicle"
        .TextMatrix(1, VehType) = "Type"
        .ColWidth(VehType) = 1095
        .ColAlignment(VehType) = flexAlignCenterCenter
        
'        .TextMatrix(0, SDate) = "Applicable From Date"  'Sale Dt."
'        .TextMatrix(1, SDate) = "Sold"
'        .ColWidth(SDate) = 1095
'        .ColAlignment(SDate) = flexAlignCenterCenter
'
'        .TextMatrix(0, SEDate) = .TextMatrix(0, SDate) '"Serv Eff Dt."
'        .TextMatrix(1, SEDate) = "ServEffect"
'        .ColWidth(SEDate) = 1095
'        .ColAlignment(SEDate) = flexAlignCenterCenter

        .TextMatrix(0, Mat) = "Amount" 'Amt Mat."
        .TextMatrix(1, Mat) = "Material"
        .ColAlignment(Mat) = flexAlignRightCenter
        .ColWidth(Mat) = 720
        .ColAlignmentFixed(Mat) = flexAlignCenterCenter
        
        .TextMatrix(0, Lab) = .TextMatrix(0, Mat)   '"Amt Lab."
        .TextMatrix(1, Lab) = "Labour"
        .ColAlignment(Lab) = flexAlignRightCenter
        .ColWidth(Lab) = 720
        .ColAlignmentFixed(Lab) = flexAlignCenterCenter

        .TextMatrix(0, Engi) = "OIL" 'Eng Oil"
        .TextMatrix(1, Engi) = "Eng"
        .ColAlignment(Engi) = flexAlignRightCenter
        .ColWidth(Engi) = 675
        .ColAlignmentFixed(Engi) = flexAlignCenterCenter
        
        .TextMatrix(0, Gear) = .TextMatrix(0, Engi) '"Gear Oil"
        .TextMatrix(1, Gear) = "Gear"
        .ColAlignment(Gear) = flexAlignRightCenter
        .ColWidth(Gear) = 675
        .ColAlignmentFixed(Gear) = flexAlignCenterCenter

        .TextMatrix(0, FAx) = .TextMatrix(0, Engi) '"FAx Oil"
        .TextMatrix(1, FAx) = "FAxl"
        .ColAlignment(FAx) = flexAlignRightCenter
        .ColWidth(FAx) = 675
        .ColAlignmentFixed(FAx) = flexAlignCenterCenter
        
        .TextMatrix(0, RAx) = .TextMatrix(0, Engi) '"RAx Oil"
        .TextMatrix(1, RAx) = "RAxl"
        .ColAlignment(RAx) = flexAlignRightCenter
        .ColWidth(RAx) = 675
        .ColAlignmentFixed(RAx) = flexAlignCenterCenter

        .TextMatrix(0, Stee) = .TextMatrix(0, Engi) ' "Steer Oil"
        .TextMatrix(1, Stee) = "Steer"
        .ColAlignment(Stee) = flexAlignRightCenter
        .ColWidth(Stee) = 675
        .ColAlignmentFixed(Stee) = flexAlignCenterCenter

'        .TextMatrix(0, RNo) = "Telco"
'        .TextMatrix(1, RNo) = "Ref.No"
'        .ColAlignment(RNo) = flexAlignLeftCenter
'        .ColWidth(RNo) = 1050
'        .ColAlignmentFixed(RNo) = flexAlignCenterCenter
'
'        .TextMatrix(0, RDate) = .TextMatrix(0, RNo) '"Telco R.Dt."
'        .TextMatrix(1, RDate) = "Ref.Dt."
'        .ColWidth(RDate) = 1095
'        .ColAlignmentFixed(RDate) = flexAlignCenterCenter
        .ColWidth(WhCat) = 0
        .Rows = 3
        .FixedRows = 2
        .AddItem FGrid.Rows - 2
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    
End Sub

Private Sub Fill_Data()
If rstmodel.RecordCount = 0 Or (rstmodel.EOF = True Or rstmodel.BOF = True) Or Txt(2).TEXT = "" Then Exit Sub
FGrid.TextMatrix(FGrid.Row, Model) = rstmodel!Code
FGrid.TextMatrix(FGrid.Row, WhCat) = IIf(IsNull(rstmodel!Wheel_Catg), "", rstmodel!Wheel_Catg)
FGrid.TextMatrix(FGrid.Row, VehType) = IIf(IsNull(rstmodel!Vehicle_Type), "", rstmodel!Vehicle_Type)
End Sub

Private Sub FillModels()
Dim Rs As ADODB.Recordset

Txt(Serv).TEXT = RstHelp!Name
Txt(ServType).TEXT = RstHelp!Code
Txt(Serv).Tag = RstHelp!Name
'Fill Models
Set Rs = New Recordset
Set Rs = GCn.Execute("SELECT Vehicle_Type,Model from Model Order By Vehicle_Type,Model")
If Rs.RecordCount > 0 Then
    FGrid.Rows = 2
    Do Until Rs.EOF
        FGrid.AddItem FGrid.Rows - 1 & Chr(9) & Rs!Vehicle_Type & Chr(9) & Rs!Model
        Rs.MoveNext
    Loop
Else
    FGrid.Rows = 2
    FGrid.AddItem FGrid.Rows - 1
End If
FGrid.FixedRows = 2
Set Rs = Nothing
End Sub
