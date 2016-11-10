VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FIND 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find Form"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   21914.47
   ScaleMode       =   0  'User
   ScaleWidth      =   25542.55
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFE0E0&
      Height          =   2025
      Left            =   180
      TabIndex        =   9
      Top             =   240
      Width           =   2925
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   60
         TabIndex        =   0
         Top             =   165
         Width           =   2820
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      ItemData        =   "FIND.frx":0000
      Left            =   3120
      List            =   "FIND.frx":0019
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   330
      Width           =   2970
   End
   Begin VB.CommandButton BACK 
      BackColor       =   &H00CFE0E0&
      Caption         =   "BACK"
      Height          =   330
      Left            =   5190
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2370
      Width           =   945
   End
   Begin VB.CommandButton cmdsearch 
      BackColor       =   &H00CFE0E0&
      Caption         =   "SEARCH"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4245
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2370
      Width           =   945
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1275
      MaxLength       =   15
      TabIndex        =   3
      Top             =   2385
      Width           =   2940
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FIND.frx":009F
      Height          =   1725
      Left            =   105
      TabIndex        =   1
      Top             =   2745
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   3043
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12640511
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
         DataField       =   ""
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
         DataField       =   ""
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
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   6708.863
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   16432.53
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Operator :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3150
      TabIndex        =   8
      Top             =   0
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00CFE0E0&
      Caption         =   "Fields :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   7
      Top             =   30
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Text :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   6
      Top             =   2430
      Width           =   1095
   End
End
Attribute VB_Name = "FIND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private I As Integer, j As Integer, l As Integer, r As Integer
Private WHERE, ORDER As Boolean
Private TEMPSQL As String
Private flag As Integer
Private op As String
Dim Master As ADODB.Recordset

Private Sub BACK_Click()
    If Master.AbsolutePosition <> adPosBOF And Master.AbsolutePosition <> adPosEOF And Master.AbsolutePosition <> adPosUnknown Then
        Call SearchForm.SEARCHBACK(Master!SearchCode)
        Check = Master!SearchCode
    Else
        Check = ""
    End If
    Set Master = Nothing
    Unload Me
End Sub
Private Sub BACK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Unload Me
End Sub
Private Sub cmdsearch_Click()
On Error Resume Next
Dim op1 As String
    If Text1.TEXT = "" Then Exit Sub
    If Text1.TEXT = "*" Then Master.Requery: GoTo aa
    If I = 135 And Not IsDate(RetDate(Text1)) Then MsgBox "Invalid Date", vbInformation, "Information": Exit Sub
    op1 = ""
    Select Case List2.ListIndex
        Case 0
            op = " = "
        Case 1
            op = " > "
        Case 2
            op = " < "
        Case 3
            op = " >= "
        Case 4
            op = " <= "
        Case 5
            op = " <> "
        Case 6
            op = " like "
            op1 = "%"
    End Select
    I = Master.Fields((List1.ListIndex + 1)).Type
    If I = adBSTR Or I = adChapter Or I = adDBTime Or I = adEmpty Or I = adError Or I = adFileTime Or I = adGUID Or I = adIDispatch Or I = adIUnknown Or I = adPropVariant Or I = adVariant Then MsgBox "Cannot Do Searching On This Field": Exit Sub
    If I = adChar Or I = adVarChar Or I = adDate Or I = adLongVarChar Or I = adLongVarWChar Or I = adVarWChar Or I = adWChar Or I = adDBDate Then
        Master.Requery
        Master.Filter = UCase(Master.Fields((List1.ListIndex + 1)).Name) & op & "'" & left(UCase(Text1.TEXT), Master.Fields((List1.ListIndex + 1)).DefinedSize) & op1 & "'"
    ElseIf I = 135 Then
        Master.Requery
        Master.Filter = Master.Fields((List1.ListIndex + 1)).Name & op & "" & ConvertDate(Text1.TEXT)
    Else
        Master.Requery
        If Not IsNumeric(Text1) Then MsgBox "Please Enter Numeric Value": Text1.SetFocus:  Exit Sub
        Master.Filter = Master.Fields((List1.ListIndex + 1)).Name & op & "" & Val(Text1.TEXT)
    End If
aa:
    DataGrid1.Columns(0).width = 0
End Sub
Private Sub cmdsearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeysA vbKeyTab, True
End Sub
Private Sub Form_Activate()
    Text1.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF10 Or KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo ERR_ROUTINE
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Master.Requery
    Set DataGrid1.DataSource = Master
    List2.SELECTED(0) = True
    For j = 1 To (Master.Fields.Count) - 1
        List1.AddItem ConvStr(Master.Fields(j).Name)
        DataGrid1.Columns(j).CAPTION = ConvStr(Master.Fields(j).Name)
    Next
    DataGrid1.Columns(0).width = 0
    Call SELECTED
    List1.SELECTED(0) = True
    List2.SELECTED(6) = True
    Exit Sub
ERR_ROUTINE:
    MsgBox err.Description
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set Master = Nothing
End Sub

Private Sub List1_Click()
    Call SELECTED
    I = Master.Fields((List1.ListIndex + 1)).Type
    If I = adChar Or I = adVarChar Or I = adLongVarChar Or I = adLongVarWChar Or I = adVarWChar Or I = adWChar Then
        Text1.MaxLength = Master.Fields((List1.ListIndex + 1)).DefinedSize
    ElseIf I = 135 Then
        Text1.MaxLength = 10
    Else
        l = Master.Fields((List1.ListIndex + 1)).Precision
        r = Master.Fields((List1.ListIndex + 1)).NumericScale
        Text1.MaxLength = 0
    End If
    Text1 = ""
End Sub
Private Sub List2_Click()
Dim Temp As Integer
    Temp = List2.ListIndex
    For I = 0 To List2.ListCount - 1
       If I <> Temp Then
       List2.SELECTED(I) = False
       End If
    Next
    Call SELECTED
End Sub
Private Function ConvStr(STR As String) As String
Dim MyStr As String, I As Integer, flag As Boolean
    MyStr = ""
    flag = True
    For I = 1 To Len(STR)
        If mID(STR, I, 1) = "_" Or mID(STR, I, 1) = " " Then
            MyStr = MyStr & " "
            flag = True
        Else
            MyStr = MyStr & IIf(flag, UCase(mID(STR, I, 1)), LCase(mID(STR, I, 1)))
            flag = False
        End If
    Next
    ConvStr = MyStr
End Function
Private Sub SELECTED()
    If IsSelected(List1) And IsSelected(List2) Then
        cmdsearch.Enabled = True
    Else
        cmdsearch.Enabled = False
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If I = adNumeric Then
        Call NumDown(Text1, KeyCode, l - r, r)
    End If
    If KeyCode = vbKeyReturn Then SendKeysA vbKeyTab, True
End Sub
Private Sub Text1_KeyPress(keyascii As Integer)
    If IsSelected(List1) And IsSelected(List2) Then
        If I = adNumeric Then
            Call NumPress(Text1, keyascii, l - r, r)
        ElseIf I = adInteger Or I = adSmallInt Then
            Call NumPress(Text1, keyascii, l - 1, 0)
        End If
    End If
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If IsSelected(List1) And IsSelected(List2) Then
        cmdsearch_Click
    End If
End Sub
