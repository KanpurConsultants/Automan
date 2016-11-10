VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FAFind 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      HideSelection   =   0   'False
      Left            =   345
      TabIndex        =   1
      Top             =   675
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   5445
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   9604
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   16777088
      ForeColorSel    =   128
      BackColorBkg    =   14873572
      GridColor       =   8438015
      GridColorFixed  =   65280
      GridColorUnpopulated=   16711935
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "{Press <Esc> For Cancel && Exit} {Press <Enter> For Select && Exit} {Press Ctrl+S to Sort the selected Row}"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   5475
      Width           =   10995
   End
   Begin VB.Menu popup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu FSR 
         Caption         =   "Filter Same Records"
         Index           =   0
      End
      Begin VB.Menu RF 
         Caption         =   "Remove Filter"
         Index           =   1
      End
   End
End
Attribute VB_Name = "FAFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Master As ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SortStr As String, I As Integer
If KeyCode = vbKeyEscape And TxtSearch.Visible = False Then Unload Me
If KeyCode = vbKeyS And Shift = 2 Then
    For I = GridSel.Col To 1 Step -1
        SortStr = SortStr + Master.Fields(I).Name + ","
    Next
    SortStr = Mid(SortStr, 1, Len(SortStr) - 1)
    GridSel.CellBackColor = FaCellBackColEnter1
    Master.Sort = SortStr 'Master.Fields(GridSel.Col).Name
    GridSel.Refresh
End If
End Sub
Private Sub Form_Load()
Dim aa As Integer
On Error GoTo ERR_ROUTINE
    Check = ""
    'Me.width = 8355: Me.top = 960: Me.height = 5805: Me.left = 1260: Label1.width = Me.width
    If MultiComp = True Then
        Set Master = G_CompCn.Execute(GSQL)
    Else
        Set Master = G_FaCn.Execute(GSQL)
    End If
    MultiComp = False
Set GridSel.dataSource = Master
GridSel.ColWidth(0) = 0
If Master.RecordCount > 0 Then
    For aa = 1 To Master.Fields.Count - 1
        If Master.Fields(aa).ActualSize = 0 Then
            GridSel.ColWidth(aa) = Master.Fields(aa).DefinedSize * 120
        Else
            GridSel.ColWidth(aa) = Master.Fields(aa).ActualSize * 120
        End If
    Next
End If
If GridSel.Rows = 1 Then GridSel.AddItem ""
GridSel_GotFocus
GridSel.Col = 1
Exit Sub
ERR_ROUTINE:            MsgBox err.Description
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
End Sub
Private Sub FSR_Click(Index As Integer)
    Master.Filter = Master.Fields(GridSel.Col).Name & " =  '" & GridSel.TextMatrix(GridSel.Row, GridSel.Col) & "'"
    GridSel.Refresh
End Sub
Private Sub GridSel_DblClick()
If GridSel.TextMatrix(1, 1) <> "" Then
    SearchForm.SEARCHBACK GridSel.TextMatrix(GridSel.Row, 0)
    Check = GridSel.TextMatrix(GridSel.Row, 0)
Else
    Check = ""
End If
Unload Me
Exit Sub
errorbox:    MsgBox err.Description, vbInformation
End Sub
Private Sub GridSel_GotFocus()
GridSel.CellBackColor = FaCellBackColEnter1
End Sub
Private Sub GridSel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then GridSel_DblClick
End Sub
Private Sub GridSel_KeyPress(KeyAscii As Integer)
FaSelGridKeyPress TxtSearch, GridSel, Master, KeyAscii, Master.Fields(GridSel.Col).Name, FaCellBackColEnter1, FaCellBackColLeave1
End Sub
Private Sub GridSel_LeaveCell()
GridSel.CellBackColor = FaCellBackColLeave1
End Sub
Private Sub GridSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu popup
End If
End Sub
Private Sub RF_Click(Index As Integer)
Master.Filter = adFilterNone
GridSel.Refresh
End Sub
Private Sub TxtSearch_Click()
TxtSearch.TEXT = "": GridSel.SetFocus: TxtSearch.Visible = False
End Sub
Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If FaNavigationKey(KeyCode) = True Then GridSel.SetFocus: TxtSearch.Visible = False
If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then GridSel.SetFocus: TxtSearch.Visible = False
End Sub
Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
FaSelGridKeyPress TxtSearch, GridSel, Master, KeyAscii, Master.Fields(GridSel.Col).Name, FaCellBackColEnter1, FaCellBackColLeave1: KeyAscii = 0
End Sub
Private Sub TxtSearch_LostFocus()
    TxtSearch.TEXT = "": GridSel.SetFocus: TxtSearch.Visible = False
End Sub
