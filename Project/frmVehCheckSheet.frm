VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehCheckSheet 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Model-wise Check List Master"
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
      Index           =   10
      Left            =   8070
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1530
      Width           =   1425
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
      Index           =   11
      Left            =   9675
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1530
      Width           =   1425
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
      Index           =   8
      Left            =   8070
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1425
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
      Index           =   9
      Left            =   9675
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1425
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
      Index           =   6
      Left            =   8070
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   990
      Width           =   1425
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
      Index           =   7
      Left            =   9675
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   990
      Width           =   1425
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
      Left            =   9675
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   1425
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
      Left            =   2145
      MaxLength       =   15
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1530
      Width           =   1425
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
      Left            =   8070
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   1425
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
      Index           =   2
      Left            =   2145
      MaxLength       =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1253
      Width           =   3000
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
      Left            =   2145
      MaxLength       =   15
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   3000
   End
   Begin MSDataGridLib.DataGrid DgChassis 
      Height          =   2445
      Left            =   75
      Negotiate       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3615
      Visible         =   0   'False
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   4313
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
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Chassis Help"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Chassis No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "EngineNo"
         Caption         =   "Engine No"
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
         DataField       =   "Model"
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
      BeginProperty Column03 
         DataField       =   "Vehicle_Type"
         Caption         =   "Vehicle Type"
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
      BeginProperty Column04 
         DataField       =   "PBill_No"
         Caption         =   "Mfg.Bill No."
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
      BeginProperty Column05 
         DataField       =   "PBill_Date"
         Caption         =   "Mfg.Bill Date"
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
      BeginProperty Column06 
         DataField       =   "InDate"
         Caption         =   "Stock In Date"
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
            ColumnWidth     =   1725.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1349.858
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   10275
      MaxLength       =   25
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2685
      Visible         =   0   'False
      Width           =   690
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   4
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
      Left            =   2145
      MaxLength       =   20
      TabIndex        =   1
      Top             =   990
      Width           =   3000
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   4500
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   7938
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   5
      BackColorFixed  =   15717816
      ForeColorFixed  =   8388736
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   14737632
      GridColor       =   12632319
      GridColorFixed  =   8421631
      FocusRect       =   0
      GridLinesFixed  =   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   " "
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Srl No. && Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Index           =   5
      Left            =   5820
      TabIndex        =   23
      Top             =   720
      Width           =   2160
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Invoice No. && Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Index           =   7
      Left            =   5850
      TabIndex        =   18
      Top             =   1260
      Width           =   2130
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Document No. && Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Index           =   6
      Left            =   5250
      TabIndex        =   17
      Top             =   1537
      Width           =   2730
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Index           =   3
      Left            =   870
      TabIndex        =   14
      Top             =   1530
      Width           =   1200
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mfg. Bill No. && Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Index           =   2
      Left            =   6270
      TabIndex        =   13
      Top             =   997
      Width           =   1710
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Index           =   1
      Left            =   870
      TabIndex        =   12
      Top             =   1260
      Width           =   990
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Index           =   0
      Left            =   870
      TabIndex        =   11
      Top             =   720
      Width           =   570
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   240
      Index           =   4
      Left            =   870
      TabIndex        =   9
      Top             =   990
      Width           =   1080
   End
End
Attribute VB_Name = "frmVehCheckSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstLab As ADODB.Recordset
Dim mFlag As Byte
Private Const Chassis = 0
Private Const Model = 1
Private Const Engine = 2
Private Const VehType = 3
Private Const PVNo As Byte = 4
Private Const PVDate As Byte = 5
Private Const MfgInvNo As Byte = 6
Private Const MfgInvDate As Byte = 7
Private Const InvNo As Byte = 8
Private Const InvDate As Byte = 9
Private Const DelChNo As Byte = 10
Private Const DelChDate As Byte = 11
'    " right(Pur_DocId,8) as PVNo,Pur_VDate,right(Sal_DocId,8) as SalVNo,Sal_VDate,right(DelCh_DocId,8) as DelChNo,DelCh_Date " & _

Dim GridKey As Integer
' Col Declaration
Dim ExitCtrl As Boolean
Private Const ItemCode As Byte = 1
Private Const Description As Byte = 2
Private Const DefVal As Byte = 3
Private Const PIndex As Byte = 4
Dim TAddMode As Boolean

Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Sub DgChassis_Click()
If RstHelp.RecordCount > 0 Then
    Fill_Data RstHelp
End If
FGrid.SetFocus
DgChassis.Visible = False
End Sub

Private Sub DgChassis_GotFocus()
    mFlag = 1
End Sub

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_GotFocus()
FGrid.BackColorSel = BackColorSelEnter
FGrid.ForeColorSel = ForeColorSelEnter
Grid_Hide
FGrid.Col = DefVal
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
''Leave Cell-- > Enter Cell-- >KeyDown
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
        Case DefVal
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case DefVal
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
SetMaxLength
Select Case FGrid.Col
    Case DefVal
       Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
'If KeyCode = vbKeyD And Shift = 2 Then
'    If FGrid.Row  >= 1 Then
'         If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
'            If FGrid.Rows  > 2 Then
'                FGrid.RemoveItem (FGrid.Row)
'            Else
'                FGrid.Rows = 1
'                FGrid.AddItem FGrid.Rows
'                FGrid.FixedRows = 1
'            End If
'         End If
'         For i = 1 To FGrid.Rows - 1
'            FGrid.TextMatrix(i, 0) = i
'         Next
'    Else
'        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
'    End If
'    FGrid.SetFocus
'End If
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
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
WinSetting Me: Ini_Grid
TopCtrl1.Tag = PubUParam
Set RstMain = New ADODB.Recordset
  Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
RstMain.Open "Select distinct Model + ChassisNo as SearchCode From Veh_CheckList  " & sitecond & " ", GCn, adOpenDynamic, adLockOptimistic
'right(Pur_DocId,8) as PVNo,Pur_VDate,right(Sal_DocId,8) as SalVNo,Sal_VDate,right(DelCh_DocId,8) as DelChNo,DelCh_Date
Set RstHelp = New ADODB.Recordset
Set RstHelp = GCn.Execute("SELECT VStk.ChassisNo as code,VStk.ChassisNo as Name,VStk.ChassisNo,M.Vehicle_Type, VStk.EngineNo, VStk.MODEL, VStk.PBILL_NO, VStk.PBILL_DATE, Right(VStk.Pur_DocId,13) AS chlNo,VStk.INDATE, " & _
    " right(Pur_DocId,8) as PVNo,Pur_VDate,right(Sal_DocId,8) as SalVNo,Sal_VDate,right(DelCh_DocId,8) as DelChNo,DelCh_Date " & _
    " FROM Veh_Stock VStk left join Model M on Vstk.Model=M.Model " & _
    " where (VStk.DelCh_DocId= '' or VStk.DelCh_DocId Is Null) and VStk.ChassisNo not in (Select ChassisNo from Veh_CheckList)")
Set DgChassis.DataSource = RstHelp

Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0:    mFlag = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RstMain = Nothing: Set RstHelp = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrLoop
BlankText
Disp_Text SETS("ADD", Me, RstMain)
RstHelp.Requery
Txt(Chassis).SetFocus
ADDFLAG = 1
Exit Sub
ErrLoop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
'Dim rs As Recordset
On Error GoTo ErrLoop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(Chassis).Enabled = False
    ADDFLAG = 2
    FGrid.SetFocus
    FGrid.Col = DefVal
'    FGrid_EnterCell
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
ErrLoop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ErrLoop
Dim mTrans As Boolean
Dim XBM
Dim Res As Integer
    If RstMain.RecordCount > 0 Then
        Res = MsgBox("Do You Want to Delete Record ", 4 + vbQuestion, "Confirmation ")
        If Res = 6 Then
            GCn.BeginTrans
            XBM = RstMain.Bookmark
            mTrans = True
            GCn.Execute ("delete from Veh_CheckList where Model='" & Txt(Model) & "' and ChassisNo= '" & Txt(Chassis) & "'")
            GCn.CommitTrans
            mTrans = False
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
    Else
        MsgBox "No Records To Delete.", vbInformation, "Information"
    End If

Exit Sub
ErrLoop:    If mTrans Then GCn.RollbackTrans
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
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    GSQL = "Select max(Model) & Max(ChassisNo) as SearchCode, max(Model) as Model_Code,max(ChassisNO) as Chassis From Veh_CheckList " & sitecond & " group by Model,ChassisNo"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    RstMain.MoveFirst
    RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim I As Integer, mQry$, mRepName$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER



    mRepName = "VehCheckList"
    mQry = "SELECT '" & Txt(Model) & "' as ModelCode,'" & Txt(VehType) & "' as VehType,'" & Txt(Chassis) & "' as Chassis,'" & Txt(Engine) & _
        "' as Engine,'" & Txt(MfgInvNo) & "' as MfgInvNo," & ConvertDate(Txt(MfgInvDate)) & " as MfgInvDate," & _
        " VCL.Item_Code,MCLM.Item_Description, VCL.Default_Value,MCLM.Report_Index, VCL.U_Name,VCL.U_EntDt,VCL.U_AE " & _
        " FROM (Veh_CheckList VCL left join ModelCheckList MCL on VCL.Model & VCL.Item_Code=MCL.Model & MCL.Item_code) " & _
        " left join ModelCheckListMast MCLM on VCL.Item_Code=MCLM.Item_Code " & _
        " WHERE VCL.Model='" & Txt(Model) & "' and VCL.ChassisNo='" & Txt(Chassis) & _
        "' Order By MCLM.Report_Index"
    
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
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
Dim mTrans As Boolean
Dim I As Byte
On Error GoTo ErrLoop
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    If IsValid(Txt(Chassis), "Chassis No.") = False Then Exit Sub
    GCn.BeginTrans
    mTrans = True
    GCn.Execute ("DELETE From Veh_CheckList Where MODEL='" & Txt(Model) & "' and ChassisNo='" & Txt(Chassis) & "'")
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ItemCode) <> "" Then
            GCn.Execute ("Insert Into Veh_CheckList(MODEL,ChassisNo,Item_Code,Default_Value,Site_Code,U_Name,U_EntDt,U_AE) Values('" & Txt(Model) & "','" & Txt(Chassis) & "','" & FGrid.TextMatrix(I, ItemCode) & "','" & FGrid.TextMatrix(I, DefVal) & "','" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
        End If
    Next
    GCn.CommitTrans
    mTrans = False
    RstMain.Requery
    RstHelp.Requery
    RstMain.FIND ("SearchCode='" & Txt(Model) & Txt(Chassis) & "'")
    Disp_Text SETS("INI", Me, RstMain)
    If ADDFLAG = 1 Then
        MoveRec
        ADDFLAG = 0
    End If
Exit Sub
ErrLoop:    If mTrans Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ErrLoop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        ADDFLAG = 0
        Grid_Hide
        Disp_Text SETS("INI", Me, RstMain)
        Me.ActiveControl.SetFocus
        MoveRec
    End If
Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical
End Sub

Private Sub MoveRec()
Dim Rst As Recordset
On Error GoTo ErrLoop
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    GSQL = "SELECT M.Vehicle_Type,VStk.ChassisNo, VStk.EngineNo, VStk.MODEL, VStk.PBILL_NO, VStk.PBILL_DATE, Right(VStk.Pur_DocId,13) AS chlNo,VStk.INDATE, " & _
        " right(Pur_DocId,8) as PVNo,Pur_VDate,right(Sal_DocId,8) as SalVNo,Sal_VDate,right(DelCh_DocId,8) as DelChNo,DelCh_Date " & _
        " FROM Veh_Stock VStk left join Model M on Vstk.Model=M.Model " & _
        " where VStk.Model + VStk.ChassisNo= '" & RstMain!SearchCode & "'"
    Set Rst = New Recordset
    Set Rst = GCn.Execute(GSQL)
    Fill_Data Rst
End If
Set Rst = Nothing
Exit Sub
ErrLoop:        MsgBox err.Description
End Sub

Private Sub TopCtrl1_eRef()
    RstHelp.Requery
End Sub

Private Sub TopCtrl1_eExit()
    RstMain.Cancel
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus Txt(Index)
Grid_Hide
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Grid_Hide: Exit Sub
Select Case Index
    Case Chassis
        DGridTxtKeyDown DgChassis, Txt, Chassis, RstHelp, KeyCode, False, 0
End Select
If DgChassis.Visible = False Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case Chassis
        If DgChassis.Visible Then DGridTxtKeyPress Txt, Chassis, RstHelp, KeyAscii, "Code"
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'Select Case Index
'End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rs As Recordset, Rst As Recordset
Select Case Index
    Case Chassis
        If Txt(Chassis) = "" Then Exit Sub
        Set Rst = GCn.Execute("Select ChassisNo From Veh_CheckList where ChassisNo='" & Txt(Chassis) & "'")
        If ADDFLAG = 1 Then
            If Not Rst.EOF Then MsgBox "Chassis Already Exists", vbInformation, "Validation": Txt(Chassis) = Txt(Chassis).Tag: Cancel = True: Exit Sub
        End If
        If RstHelp.RecordCount = 0 Or (RstHelp.EOF = True Or RstHelp.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Chassis).TEXT = ""
            Txt(Chassis).Tag = ""
            Txt(Engine) = ""
            Txt(Model) = ""
            Txt(VehType) = ""
            Txt(MfgInvNo) = ""
            Txt(MfgInvDate) = ""
        Else
'            Txt(Chassis).Text = RstHelp!Name
'            Txt(Chassis).Tag = RstHelp!Code
            Fill_Data RstHelp
        End If
End Select
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    Ctrl_GetFocus TxtGrid(Index)
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then TxtGrid(0) = TxtGrid(0).Tag: Exit Sub
    Select Case FGrid.Col
        Case DefVal
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, DefVal
                End If
            End If
    End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If KeyAscii = vbKeyEscape Then Exit Sub
Call CheckQuote(KeyAscii)
'Select Case FGrid.Col
'    Case ADItem
'        If DGADItem.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsADItem, KeyAscii, "name"
'    Case Qty
'        Call NumPress(TxtGrid(Index), KeyAscii, 6, 0)
'End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'Select Case FGrid.Col
'    Case DefVal
'        FGrid.TextMatrix(FGrid.Row, DefVal) = Format(Val(TxtGrid(Index).Text), "0")
'End Select
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
    Grid_Hide
End If
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub
Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid.Col
    Case DefVal
        FGrid.TextMatrix(FGrid.Row, DefVal) = TxtGrid(Index)
'    Case Lab_Code
'        If FGrid.TextMatrix(FGrid.Row, Lab_Code) <> Txtgrid(0) Then Call Fill_Data
'    Case LCode
'        If FGrid.TextMatrix(FGrid.Row, LCode) <> Txtgrid(0) Then Call Fill_Data
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
End Function

Private Sub BlankText()
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I) = ""
        Txt(I).Tag = ""
    Next
    TxtGrid(0).TEXT = ""
    FGrid.Rows = 1
    FGrid.AddItem ""
    FGrid.FixedRows = 1
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next
    TxtGrid(0).Enabled = Enb
End Sub

Private Sub Ini_Grid()
Dim I As Byte
    With FGrid
'        .left = Me.left '+ 45
'        .width = Me.width - 120
        .RowHeightMin = PubGridRowHeight '220
        .height = .RowHeight(0) * 15
'        .top = 1575
        .Cols = 5
        .ColAlignmentFixed = flexAlignCenterCenter
        
        .TextMatrix(0, 0) = "S.No."
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 550

        .TextMatrix(0, ItemCode) = "ItemCode"
        .ColAlignment(ItemCode) = flexAlignLeftCenter
        .ColWidth(ItemCode) = 0
        
        .TextMatrix(0, Description) = "Description"
        .ColAlignment(Description) = flexAlignLeftCenter
        .ColWidth(Description) = 2500
                
        .TextMatrix(0, DefVal) = "Value"
        .ColAlignment(DefVal) = flexAlignLeftCenter
        .ColWidth(DefVal) = 1500
        .TextMatrix(0, PIndex) = "Print Index"
        .ColAlignment(PIndex) = flexAlignLeftCenter
        .ColWidth(PIndex) = 1200
    End With
    FGrid.left = (Me.width - FGrid.width) / 2
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
   
    'DgChassis.width = 7005:
    DgChassis.left = Me.width - (DgChassis.width + mRtScale): DgChassis.top = FGrid.top: DgChassis.height = Me.height - (DgChassis.top + mBotScale)
End Sub

Private Sub Grid_Hide()
If DgChassis.Visible = True Then DgChassis.Visible = False
End Sub

Private Sub SetMaxLength()
Select Case FGrid.Col   'Index
    Case DefVal
        TxtGrid(0).MaxLength = 10
    Case Else
        TxtGrid(0).MaxLength = 0
End Select
End Sub

Private Function Fill_Data(Rst As ADODB.Recordset) As Boolean
Dim Rs As ADODB.Recordset
Txt(Chassis) = Rst!ChassisNo
Txt(Chassis).Tag = Rst!ChassisNo
Txt(Engine) = Rst!EngineNo
Txt(Model) = Rst!Model
Txt(VehType) = Rst!Vehicle_Type
Txt(PVNo) = XNull(Rst!PVNo)
Txt(PVDate) = IIf(IsNull(Rst!Pur_VDate), "", Rst!Pur_VDate)
Txt(MfgInvNo) = XNull(Rst!PBILL_NO)
Txt(MfgInvDate) = IIf(IsNull(Rst!PBILL_DATE), "", Rst!PBILL_DATE)
Txt(InvNo) = IIf(IsNull(Rst!SalVNo), 0, Rst!SalVNo)
Txt(InvDate) = IIf(IsNull(Rst!Sal_VDate), "", Rst!Sal_VDate)
Txt(DelChNo) = XNull(Rst!DelChNo)
Txt(DelChDate) = IIf(IsNull(Rst!DelCh_Date), "", Rst!DelCh_Date)

If TopCtrl1.TopText2.CAPTION = "Add" Then
    GSQL = "SELECT MCL.Item_Code,MCL.Default_Value,MCLM.Item_Description,MCLM.Report_Index " & _
        "from ModelCheckList MCL left join ModelCheckListMast MCLM on MCL.Item_Code=MCLM.Item_Code " & _
        "where MCL.Model='" & Txt(Model) & "' Order By Report_Index"
Else
    GSQL = "SELECT VCL.Item_Code,VCL.Default_Value,MCLM.Item_Description,MCLM.Report_Index " & _
        "from Veh_CheckList VCL left join ModelCheckListMast MCLM on VCL.Item_Code=MCLM.Item_Code " & _
        "where VCL.ChassisNo='" & Txt(Chassis) & "' Order By Report_Index"
End If
Set Rs = New Recordset
Set Rs = GCn.Execute(GSQL)
If Rs.RecordCount > 0 Then
    FGrid.Rows = 1
    Do Until Rs.EOF
        FGrid.AddItem Rs.AbsolutePosition & Chr(9) & Rs!Item_Code & Chr(9) & Rs!Item_Description & Chr(9) & Rs!Default_Value & Chr(9) & Rs!Report_Index
        Rs.MoveNext
    Loop
    FGrid.FixedRows = 1
End If
Set Rs = Nothing
FGrid.Row = 1
FGrid.Col = DefVal
'If FGrid.Visible Then FGrid.SetFocus
End Function
