VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "Msdatgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmBankRecon 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Bank Reconsilation"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   Icon            =   "FrmBankRecon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cleared"
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
      Left            =   6570
      TabIndex        =   2
      Top             =   60
      Width           =   1170
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "All"
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
      Left            =   7740
      TabIndex        =   3
      Top             =   60
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Uncleared"
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
      Left            =   5250
      TabIndex        =   1
      Top             =   60
      Value           =   -1  'True
      Width           =   1320
   End
   Begin Petro.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   7290
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   661
   End
   Begin VB.Frame FrAc 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   2460
      Left            =   -5400
      TabIndex        =   14
      Top             =   3555
      Visible         =   0   'False
      Width           =   5325
      Begin MSDataGridLib.DataGrid DgAc 
         Height          =   2115
         Left            =   30
         TabIndex        =   15
         Top             =   315
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   3731
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   13504523
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "SubCode"
            Caption         =   ""
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
            Caption         =   ""
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
            MarqueeStyle    =   5
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   134.929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4515.024
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   5250
      End
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   9975
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   30
      Width           =   1320
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   1305
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   30
      Width           =   3720
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   4
      Left            =   10230
      TabIndex        =   11
      Text            =   "4"
      Top             =   6885
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   3
      Left            =   8850
      TabIndex        =   10
      Text            =   "3"
      Top             =   6885
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   2
      Left            =   10215
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   6495
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDE2C1&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   1
      Left            =   2430
      MaxLength       =   25
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1500
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   5940
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   435
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   10478
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   9
      FixedCols       =   0
      BackColorFixed  =   12640511
      ForeColorFixed  =   4210752
      BackColorSel    =   8421504
      BackColorBkg    =   16777215
      BackColorUnpopulated=   11794923
      GridColor       =   12632319
      WordWrap        =   -1  'True
      FocusRect       =   0
      GridLinesUnpopulated=   3
      ScrollBars      =   2
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "Date        |Particular                                          |VType |VNo|Chq.No|Chq.Dt|Bank Dt|Dr|Cr    "
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "As on Dt."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   4
      Left            =   9030
      TabIndex        =   13
      Top             =   60
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   3
      Left            =   150
      TabIndex        =   12
      Top             =   60
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Index           =   2
      Left            =   45
      Top             =   0
      Width           =   11295
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Index           =   1
      Left            =   90
      Top             =   6855
      Visible         =   0   'False
      Width           =   11625
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Index           =   0
      Left            =   90
      Top             =   6465
      Visible         =   0   'False
      Width           =   11640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Withdrwal/Deposit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   1
      Left            =   255
      TabIndex        =   8
      Top             =   6915
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance As Per Company Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   0
      Left            =   255
      TabIndex        =   7
      Top             =   6540
      Visible         =   0   'False
      Width           =   2940
   End
End
Attribute VB_Name = "FrmBankRecon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddFlag As Byte
Dim RstAc As New ADODB.Recordset
Private Const PetroHsd As Byte = 0
Private IntStaId As Integer
'******Please Don't Remove these variable
Dim GridKey As Integer
Dim ExitCtrl As Boolean
Dim TAddMode  As Boolean
Dim MyCtrl As Integer
'************** TXT
Private Const AcName  As Byte = 0
Private Const AcStDt  As Byte = 1
Private Const AcCmpBl  As Byte = 2
Private Const AcBlDr  As Byte = 3
Private Const AcBlCr  As Byte = 4
'*************Grid (1) Cols
Private Const GDate  As Byte = 0
Private Const GParti As Byte = 1
Private Const GVtype As Byte = 2
Private Const GVno As Byte = 3
Private Const GChqNo As Byte = 4
Private Const GChqDt As Byte = 5
Private Const GBnkDt As Byte = 6
Private Const GDrAmt As Byte = 7
Private Const GCrAmt As Byte = 8
Private Const GDocId As Byte = 9
Private Sub Form_Activate()
    Txt(AcName).SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i%
Dim transFlag As Byte, MySql$
On Error GoTo ErrLoop
    If KeyCode = vbKeyS And Shift = 2 Then
        i = MsgBox("Do You Want Update", vbYesNo, "Information")
        If i = vbYes Then
           transFlag = 0
            GCn.BeginTrans
                transFlag = 1
                For i = 1 To FGrid(1).Rows - 1
                    MySql = "Update Ledger Set Clg_Date=" & ConvertDate(FGrid(1).TextMatrix(i, GBnkDt)) & " Where DocId='" & FGrid(1).TextMatrix(i, GDocId) & "'"
                    GCn.Execute (MySql)
                Next
            GCn.CommitTrans
            transFlag = 0
            Txt(AcName).SetFocus
        End If
    End If
    Exit Sub
ErrLoop:
    If transFlag = 1 Then GCn.RollbackTrans
    MsgBox ERR.Description, vbCritical
    Exit Sub
End Sub
Private Sub Form_Load()
Dim i As Byte
On Error GoTo ErrLoop
    For i = 0 To Txt.Count - 1
        Txt(i).Text = ""
    Next
    Me.BackColor = FrmBackCol
    Me.Left = 210
    Me.Top = 105
    Me.Height = 7740
    Me.Width = 11355
    
    With RstAc
        .ActiveConnection = GCn
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .Open "SELECT SubCode,Name FROM SubGroup Where Nature='Bank' ORDER BY Name"
    End With
    DgAc.Columns(0).Visible = False
    Set DgAc.DataSource = RstAc
    FrAc.Left = Txt(AcName).Left
    FrAc.Top = Txt(AcName).Top + Txt(AcName).Height
    With FGrid(1)
        .BackColorBkg = GridBckCol
        .Height = 285 * 21
        '.Left = 1410
        '.Top = 2925
        .ColWidth(GDate) = 930          'Date
        .ColAlignmentFixed(GDate) = flexAlignCenterCenter
        .ColAlignment(GDate) = flexAlignLeftCenter
        
        .ColWidth(GParti) = 3225         'Paritcular
        .ColAlignmentFixed(GParti) = flexAlignLeftCenter
        .ColAlignment(GParti) = flexAlignLeftCenter
 
        .ColWidth(GVtype) = 900        'Voucher Type
        .ColAlignmentFixed(GVtype) = flexAlignRightCenter
        .ColAlignment(GVtype) = flexAlignRightCenter

        .ColWidth(GVno) = 795          'Voucher No
        .ColAlignmentFixed(GVno) = flexAlignRightCenter
        .ColAlignment(GVno) = flexAlignRightCenter
        
        .ColWidth(GChqNo) = 1005          'Chq No
        .ColAlignmentFixed(GChqNo) = flexAlignRightCenter
        .ColAlignment(GChqNo) = flexAlignRightCenter
        
        .ColWidth(GChqDt) = 960          'Chq Dt
        .ColAlignmentFixed(GChqDt) = flexAlignCenterCenter
        .ColAlignment(GChqDt) = flexAlignLeftCenter
        
        .ColWidth(GBnkDt) = 1170          'Bank Clearing Dt
        .ColAlignmentFixed(GBnkDt) = flexAlignCenterCenter
        .ColAlignment(GBnkDt) = flexAlignLeftCenter

        .ColWidth(GDrAmt) = 1185          'Debit Amount
        .ColAlignmentFixed(GDrAmt) = flexAlignRightCenter
        .ColAlignment(GDrAmt) = flexAlignRightCenter
        
        .ColWidth(GCrAmt) = 1185          'Credit Amount
        .ColAlignmentFixed(GCrAmt) = flexAlignRightCenter
        .ColAlignment(GCrAmt) = flexAlignRightCenter
        .ColWidth(GDocId) = 0
    End With
    
    
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub Form_Resize()
    'TopCtrl1.Width = Me.Width - 50
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set RstAc = Nothing
    
End Sub
Private Sub MoveRec()
Dim i As Integer, RstLocal As New ADODB.Recordset
On Error GoTo ErrLoop
    FGrid(1).Rows = 1
    FGrid(1).AddItem FGrid(1).Rows
    FGrid(1).FixedRows = 1
    i = 1
    Txt(AcBlDr).Text = ""
    Txt(AcBlCr).Text = ""
    If Option1(0).value = True Then
        Set RstLocal = GCn.Execute("SELECT Ledger.DocId,Ledger.V_Date,Ledger.V_Type,Ledger.V_No,Ledger.Narration,IIF(ISNULL(Ledger.AmtDr),0,Ledger.AmtDr) As DebitAmt,IIF(ISNULL(Ledger.AmtCr),0,Ledger.AmtCr) As CreditAmt,Subgroup.Name,Ledger.Chq_No,Ledger.Chq_Date,Ledger.Clg_Date From Ledger Left Join Subgroup on SubGroup.Subcode=Ledger.ContraSub Where Ledger.Clg_Date Is Null And Ledger.SubCode='" & Txt(AcName).Tag & "' And  V_Date<=" & ConvertDate(Txt(AcStDt).Text) & " Order By Ledger.V_date")
    ElseIf Option1(1).value = True Then
        Set RstLocal = GCn.Execute("SELECT Ledger.DocId,Ledger.V_Date,Ledger.V_Type,Ledger.V_No,Ledger.Narration,IIF(ISNULL(Ledger.AmtDr),0,Ledger.AmtDr) As DebitAmt,IIF(ISNULL(Ledger.AmtCr),0,Ledger.AmtCr) As CreditAmt,Subgroup.Name,Ledger.Chq_No,Ledger.Chq_Date,Ledger.Clg_Date From Ledger Left Join Subgroup on SubGroup.Subcode=Ledger.ContraSub Where Ledger.SubCode='" & Txt(AcName).Tag & "' And  V_Date<=" & ConvertDate(Txt(AcStDt).Text) & " Order By Ledger.V_date")
    Else
        Set RstLocal = GCn.Execute("SELECT Ledger.DocId,Ledger.V_Date,Ledger.V_Type,Ledger.V_No,Ledger.Narration,IIF(ISNULL(Ledger.AmtDr),0,Ledger.AmtDr) As DebitAmt,IIF(ISNULL(Ledger.AmtCr),0,Ledger.AmtCr) As CreditAmt,Subgroup.Name,Ledger.Chq_No,Ledger.Chq_Date,Ledger.Clg_Date From Ledger Left Join Subgroup on SubGroup.Subcode=Ledger.ContraSub Where not Ledger.Clg_Date Is Null And Ledger.SubCode='" & Txt(AcName).Tag & "' And  V_Date<=" & ConvertDate(Txt(AcStDt).Text) & " Order By Ledger.V_date")
    End If
    If RstLocal.RecordCount > 0 Then
        i = 1
        Do While Not RstLocal.EOF
            With FGrid(1)
                .AddItem ""
                .TextMatrix(i, GDate) = Format(RstLocal!V_Date, "dd/MMM/yyyy")
                .TextMatrix(i, GParti) = IIf(IsNull(RstLocal!Name), "N/a", RstLocal!Name)
                .TextMatrix(i, GVtype) = RstLocal!V_Type
                .TextMatrix(i, GVno) = RstLocal!V_No
                
                If IsNull(RstLocal!Chq_No) Then
                    .TextMatrix(i, GChqNo) = ""
                Else
                    .TextMatrix(i, GChqNo) = RstLocal!Chq_No
                End If
                
                If IsNull(RstLocal!Chq_Date) Then
                    .TextMatrix(i, GChqDt) = ""
                Else
                    .TextMatrix(i, GChqDt) = RstLocal!Chq_Date
                End If
                
                If IsNull(RstLocal!Clg_Date) Then
                    .TextMatrix(i, GBnkDt) = ""
                Else
                    .TextMatrix(i, GBnkDt) = RstLocal!Clg_Date
                End If
                
                If (Not IsNull(RstLocal!DebitAmt) And RstLocal!DebitAmt) > 0 Then
                    .TextMatrix(i, GDrAmt) = Format(RstLocal!DebitAmt, "0.00")
                    Txt(AcBlDr).Text = Format(Val(Txt(AcBlDr).Text) + RstLocal!DebitAmt, "0.00")
                Else
                    .TextMatrix(i, GDrAmt) = ""
                End If
                If (Not IsNull(RstLocal!CreditAmt) And RstLocal!CreditAmt) > 0 Then
                    .TextMatrix(i, GCrAmt) = Format(RstLocal!CreditAmt, "0.00")
                    Txt(AcBlCr).Text = Format(Val(Txt(AcBlCr).Text) + RstLocal!CreditAmt, "0.00")
                Else
                    .TextMatrix(i, GCrAmt) = ""
                End If
                .TextMatrix(i, GDocId) = RstLocal!DOCID
                i = i + 1
            End With
            RstLocal.MoveNext
        Loop
        Set GRs = GCn.Execute("SELECT SubGroup.SubCode, SubGroup.Name,Sum(Ledger.AmtDr)-Sum(Ledger.AmtCr) AS Balance FROM SubGroup LEFT JOIN Ledger ON SubGroup.SubCode = Ledger.SubCode  Where SubGroup.SubCode ='" & Txt(AcName).Tag & "' GROUP BY SubGroup.SubCode, SubGroup.Name")
        If GRs.RecordCount > 0 Then
            Txt(AcCmpBl).Text = Format(GRs!Balance, "0.00")
        End If
        FGrid(1).FixedRows = 1
        FGrid(1).Redraw = True
        TopCtrl1.TopText2.CAPTION = "ADD"
    End If
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information"
End Sub
Private Sub FGrid_Click(Index As Integer)
'    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    TxtGrid(Index).Visible = False
End Sub
Private Sub FGrid_DblClick(Index As Integer)
'    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid(Index).Col
        Case GBnkDt
            Call GridDblClick(Me, FGrid(Index), TxtGrid, Index)
    End Select
    TAddMode = False
End Sub
Private Sub FGrid_EnterCell(Index As Integer)
    'FGrid(Index).CellBackColor = CellBackColEnter
End Sub
Private Sub FGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrLoop
    'Leave Cell--> Enter Cell-->KeyDown
'    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid(Index).Tag) = (FGrid(Index).Rows - (FGrid(Index).Rows - 1)) Then
     '   FGrid(Index).CellBackColor = CellBackColLeave
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGrid(Index).Tag) = FGrid(Index).Rows - 1 Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
'            TopCtrl1_eSave
        Else
      '      FGrid(Index).CellBackColor = CellBackColEnter
            Me.ActiveControl.SetFocus
        End If
    End If
    GridKey = KeyCode
    FGrid(Index).Tag = FGrid(Index).Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGrid(Index).Col
            Case GBnkDt
                FGrid(Index).TextMatrix(FGrid(Index).Row, FGrid(Index).Col) = ""
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case FGrid(Index).Col
            Case GBnkDt
                Call GridDblClick(Me, FGrid(Index), TxtGrid, Index)
                TAddMode = False
        End Select
    End If
    KeyCode = 0
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub FGRID_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ErrLoop
'    If TopCtrl1.PrvKeyCode = 83 And MyCtrl = 2 Then FGrid(Index).CellBackColor = CellBackColLeave: KeyAscii = 0: Exit Sub
    Select Case FGrid(Index).Col
        Case GBnkDt
           Call Get_Text(Me, FGrid(Index), TxtGrid, Index, True, KeyAscii)
    End Select
    If KeyAscii <> vbKeyReturn Then TAddMode = True
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub FGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Integer
On Error GoTo ErrLoop
'    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid(Index).ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGrid(Index).Row >= 1 Then
            If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                If FGrid(Index).Rows > 2 Then
                    FGrid(Index).RemoveItem (FGrid(Index).Row)
                Else
                    FGrid(Index).Rows = 1
                    FGrid(Index).AddItem FGrid(Index).Rows
                    FGrid(Index).FixedRows = 1
                End If
            End If
            For i = 1 To FGrid(Index).Rows - 1
                FGrid(Index).TextMatrix(i, 0) = i
             Next
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid(Index).SetFocus
    End If
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub FGrid_Scroll(Index As Integer)
    TxtGrid(Index).Visible = False
 End Sub
Private Sub FGrid_LeaveCell(Index As Integer)
    'FGrid(Index).CellBackColor = CellBackColLeave
End Sub
Private Sub FGrid_Validate(Index As Integer, Cancel As Boolean)
'    FGrid(Index).CellBackColor = CellBackColLeave
End Sub


Private Sub Option1_Click(Index As Integer)
    MoveRec
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case AcName
            DGridTxtKeyUp Txt, Index, RstAc, KeyAscii, "Name"
    End Select
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ErrLoop
'    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    TxtGrid(Index).MaxLength = 9
    TxtGrid(Index).Alignment = vbLeftJustify
    FGrid(Index).CellBackColor = CellBackColLeave
    TxtGrid(Index).Tag = FGrid(Index).TextMatrix(FGrid(Index).Row, FGrid(Index).Col)
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrLoop
    If KeyCode = vbKeyEscape Then
        TxtGrid(Index).Text = TxtGrid(Index).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        TxtGrid(Index).Visible = False
        FGrid(Index).SetFocus
        Exit Sub
    End If
    Select Case FGrid(Index).Col
        Case GBnkDt
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave(Index) = True Then
                    GridTxtDown FGrid(Index), TxtGrid, Index, KeyCode, TAddMode, FGrid(Index).Cols
                Else
                    TxtGrid_LostFocus Index
                    TxtGrid(Index).SetFocus
                End If
            End If
    End Select
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ErrLoop
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
    Call CheckQuote(KeyAscii)
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
On Error GoTo ErrLoop
    Select Case FGrid(1).Cols
        Case GBnkDt
            FGrid(Index).TextMatrix(FGrid(Index).Row, FGrid(Index).Col) = TxtGrid(Index).Text
    End Select
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TxtGrid_LostFocus(Index As Integer)
On Error GoTo ErrLoop
    If ExitCtrl = False Then Exit Sub
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
Dim j As Integer
    Select Case FGrid(Index).Col
        Case GBnkDt
            FGrid(Index).TextMatrix(FGrid(Index).Row, FGrid(Index).Col) = RetDate(TxtGrid(Index))
    End Select
    TxtGrid(Index).Visible = False
End Sub
Private Function TxtGridLeave(Index As Integer) As Boolean
Dim j As Integer
On Error GoTo ErrLoop
    Select Case FGrid(Index).Col
        Case GBnkDt
            FGrid(Index).TextMatrix(FGrid(Index).Row, FGrid(Index).Col) = RetDate(TxtGrid(Index))
    End Select
    ExitCtrl = True
    TxtGridLeave = True
    TxtGrid(Index).Visible = False
    FGrid(Index).SetFocus
    Exit Function
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Function
End Function
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrLoop
    Select Case Index
        Case AcName
            DGridTxtKeyDown FrAc, Txt, Index, RstAc, KeyCode, False, 1
    End Select
    If KeyCode = vbKeyReturn Then SendKeys vbTab
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrLoop
    Select Case Index
        Case AcName
           
    End Select
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Lrs As ADODB.Recordset
On Error GoTo ErrLoop
    Select Case Index
        Case AcName
            If RstAc.EOF = True And RstAc.BOF = True Then Exit Sub
                If Txt(Index).Text <> "" Then
                    Txt(Index).Tag = RstAc!SubCode
                    Txt(Index).Text = RstAc!Name
                End If
            FrAc.Visible = False
        Case AcStDt
            Txt(Index).Text = RetDate(Txt(Index))
            MoveRec
    End Select
    Exit Sub
ErrLoop:
    MsgBox ERR.Description, vbInformation, "Information": Exit Sub
End Sub
