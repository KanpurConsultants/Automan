VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPartHelp 
   Appearance      =   0  'Flat
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   4380
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   2385
      Index           =   0
      Left            =   4200
      TabIndex        =   7
      Top             =   705
      Visible         =   0   'False
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   4207
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
      FormatLocked    =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "PART_NO"
         Caption         =   "Part No."
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
         DataField       =   "Part_Name"
         Caption         =   "Std.Name"
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
         DataField       =   "Local_Name"
         Caption         =   "Local Name"
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
         DataField       =   "PART_NO"
         Caption         =   "Part No."
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
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   3344.882
         EndProperty
         BeginProperty Column02 
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   3479.811
         EndProperty
         BeginProperty Column03 
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   1950.236
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1650
      Left            =   15
      TabIndex        =   3
      Top             =   2670
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   2910
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
      FormatLocked    =   -1  'True
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Root_Part_No"
         Caption         =   "Root No."
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
         DataField       =   "Alternate_Part_No"
         Caption         =   "Alternate Part No"
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
         DataField       =   "Cur_TP_STK"
         Caption         =   "TP.Stk"
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
         DataField       =   "Cur_TB_STK"
         Caption         =   "TB.Stk"
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
         DataField       =   "TP_SRate"
         Caption         =   "TP.Rate"
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
         DataField       =   "TB_SRate"
         Caption         =   "TB.Rate"
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
         DataField       =   "BIN_LOCA"
         Caption         =   "Bin Location"
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
      BeginProperty Column07 
         DataField       =   "Local_Name"
         Caption         =   "Local Name"
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
      BeginProperty Column08 
         DataField       =   "Part_Name"
         Caption         =   "Part Name"
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
         AllowSizing     =   0   'False
         BeginProperty Column00 
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column06 
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column07 
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   2310.236
         EndProperty
         BeginProperty Column08 
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   2160
         EndProperty
      EndProperty
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
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   1
      Top             =   315
      Width           =   3630
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
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   2
      Top             =   585
      Width           =   3630
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
      Left            =   1680
      MaxLength       =   22
      TabIndex        =   0
      Top             =   45
      Width           =   3630
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   315
      TabIndex        =   6
      Top             =   300
      Width           =   1260
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Std.Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   540
      TabIndex        =   5
      Top             =   570
      Width           =   1035
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   18
      Left            =   720
      TabIndex        =   4
      Top             =   30
      Width           =   855
   End
End
Attribute VB_Name = "frmPartHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstPartNo As ADODB.Recordset, rstAlternate As ADODB.Recordset
Private mPartNo As String, mLocalName As String, mPartName As String
Private mFlag As Byte, mFLAG1 As Byte, mResult As pHelpCls

Public Property Let partSearch(helpAttr As pHelpCls)
    mFlag = 0
    mFLAG1 = 0
    Set mResult = New pHelpCls
    Set mResult.Conn = helpAttr.Conn
    Set mResult.resultTextBox = helpAttr.resultTextBox
    Set rstPartNo = New ADODB.Recordset
'    rstPartNo.LockType = adLockOptimistic
'    rstPartNo.CursorType = adOpenKeyset
    rstPartNo.Open "select PART_NO,Local_Name,Part_Name,Part_NoHelp,Part_NameHelp from PART where div_code ='" & PubDivCode & "'", helpAttr.Conn, adOpenStatic, adLockReadOnly, adAsyncFetch
    Txt(0) = XNull(helpAttr.resultTextBox)
    Txt_GotFocus 0
    Set DataGrid(0).DataSource = rstPartNo
End Property
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Txt_Validate 0, False
    mResult.resultTextBox = Txt(0)
    Set rstPartNo = Nothing
End Sub
Private Sub Txt_GotFocus(Index As Integer)
mFlag = 0
mFLAG1 = 0
Dim mBookMark
DataGrid(0).Columns(3).width = 2085: DataGrid(0).Columns(2).width = 3345: DataGrid(0).Columns(1).width = 3345: DataGrid(0).Columns(0).width = 2085
Select Case Index
    Case 0
        DataGrid(0).Columns(3).width = 0: DataGrid(0).Columns(2).width = 0
        mBookMark = rstPartNo.Bookmark
        rstPartNo.Sort = "Part_NO ASC"
        rstPartNo.Bookmark = mBookMark
        partNoSearch
    Case 1
        DataGrid(0).Columns(1).width = 0: DataGrid(0).Columns(0).width = 0
        mBookMark = rstPartNo.Bookmark
        rstPartNo.Sort = "Local_Name ASC"
        rstPartNo.Bookmark = mBookMark
        localNameSearch
    Case 2
        DataGrid(0).Columns(2).width = 0: DataGrid(0).Columns(0).width = 0
        mBookMark = rstPartNo.Bookmark
        rstPartNo.Sort = "Part_NAME ASC"
        rstPartNo.Bookmark = mBookMark
        partNameSearch
End Select
DataGrid(0).Visible = True
DataGrid(0).top = 0     'Txt(Index).Top
DataGrid(0).left = Txt(Index).left + Txt(Index).width + 10
alternateSearch XNull(rstPartNo!Part_No)
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
mFlag = 0
mFLAG1 = 0
Select Case Index
    Case 0
        partNoSearch
    Case 1
        localNameSearch
    Case 2
        partNameSearch
End Select
alternateSearch XNull(rstPartNo!Part_No)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
If Not rstPartNo.EOF And Not rstPartNo.BOF Then
    Txt(0) = XNull(rstPartNo!Part_No): Txt(1) = XNull(rstPartNo!Local_Name): Txt(2) = XNull(rstPartNo!Part_Name)
End If
End Sub
Private Sub alternateSearch(rootPart As String)
If Not rstPartNo.EOF And Not rstPartNo.BOF Then
    Set rstAlternate = New ADODB.Recordset
    rstAlternate.LockType = adLockOptimistic
    rstAlternate.CursorType = adOpenKeyset
    rstAlternate.Open "SELECT Root_Part_No,Part_Alternate.Alternate_Part_No,TB_SRate,TP_SRate,Cur_TB_STK,Cur_TP_STK,BIN_LOCA,Local_Name,Part_Name FROM Part_Alternate LEFT JOIN Part ON Part_Alternate.Alternate_Part_No = Part.PART_NO and Part.Div_Code = Part_Alternate.Div_Code WHERE Root_Part_No=" & Chk_Text(Trim(rootPart)) & "  AND Root_Part_No<>Alternate_Part_No ORDER BY Root_Part_No,Alternate_Part_No", mResult.Conn
    Set DataGrid1.DataSource = rstAlternate
End If
End Sub
Private Sub partNoSearch()
rstPartNo.MoveFirst
rstPartNo.FIND "PART_NO >='" & XNull(Txt(0)) & "'"
If Not rstPartNo.EOF Then
    If mID(rstPartNo!Part_No, 1, Len(Trim(XNull(Txt(0))))) <> Trim(XNull(Txt(0))) Then
        partNoExSearch
    Else
        Txt(1) = XNull(rstPartNo!Local_Name): Txt(2) = XNull(rstPartNo!Part_Name)
    End If
Else
    partNoExSearch
End If
End Sub
Private Sub partNoExSearch()
Dim tempRst As ADODB.Recordset
Set tempRst = rstPartNo.Clone
tempRst.Sort = "Part_NoHelp ASC"
tempRst.FIND "Part_NoHelp >='" & FilterString(XNull(Txt(0))) & "'"
If Not tempRst.EOF Then
    rstPartNo.MoveFirst
    rstPartNo.FIND "PART_NO >='" & XNull(tempRst!Part_No) & "'"
    Txt(1) = XNull(tempRst!Local_Name): Txt(2) = XNull(tempRst!Part_Name)
Else
    Txt(1) = "": Txt(2) = ""
End If
Set tempRst = Nothing
End Sub
Private Sub partNameSearch()
rstPartNo.MoveFirst
rstPartNo.FIND "Part_Name >='" & XNull(Txt(2)) & "'"
If Not rstPartNo.EOF Then
    If mID(rstPartNo!Part_Name, 1, Len(Trim(XNull(Txt(2))))) <> Trim(XNull(Txt(2))) Then
        partNameExSearch
    Else
        Txt(0) = XNull(rstPartNo!Part_No): Txt(1) = XNull(rstPartNo!Local_Name)
    End If
Else
    partNameExSearch
End If
End Sub
Private Sub partNameExSearch()
Dim tempRst As ADODB.Recordset
Set tempRst = rstPartNo.Clone
tempRst.Sort = "Part_NameHelp ASC"
tempRst.FIND "Part_NameHelp >='" & FilterString(XNull(Txt(2))) & "'"
If Not tempRst.EOF Then
    rstPartNo.MoveFirst
    rstPartNo.FIND "Part_Name >='" & XNull(tempRst!Part_Name) & "'"
    Txt(1) = XNull(tempRst!Local_Name): Txt(0) = XNull(tempRst!Part_No)
Else
    Txt(1) = "": Txt(0) = ""
End If
Set tempRst = Nothing
End Sub
Private Sub localNameSearch()
rstPartNo.MoveFirst
rstPartNo.FIND "Local_Name >='" & XNull(Txt(1)) & "'"
If Not rstPartNo.EOF Then
    Txt(0) = XNull(rstPartNo!Part_No): Txt(2) = XNull(rstPartNo!Part_Name)
Else
    Txt(0) = "": Txt(2) = ""
End If
End Sub
Private Sub DataGrid_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
If mFlag = 1 Then
    Txt(0) = DataGrid(0).Columns(0).TEXT: Txt(1) = DataGrid(0).Columns(2).TEXT: Txt(2) = DataGrid(0).Columns(1).TEXT
    alternateSearch XNull(Txt(0))
End If
End Sub
Private Sub DataGrid_GotFocus(Index As Integer)
    mFlag = 1
End Sub
Private Sub DataGrid1_GotFocus()
    mFLAG1 = 1
End Sub
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mFLAG1 = 1 Then
    Txt(0) = DataGrid1.Columns(1).TEXT: Txt(1) = DataGrid1.Columns(7).TEXT: Txt(2) = DataGrid1.Columns(8).TEXT
    Txt_GotFocus 0
End If
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Integer
Select Case KeyCode
    Case vbKeyUp
        If Not rstPartNo.BOF Then rstPartNo.MovePrevious
    Case vbKeyDown
        If Not rstPartNo.EOF Then rstPartNo.MoveNext
    Case 33
        For i = 1 To 9
            If Not rstPartNo.BOF Then rstPartNo.MovePrevious
        Next
    Case 34
        For i = 1 To 9
            If Not rstPartNo.EOF Then rstPartNo.MoveNext
        Next
End Select
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        If Not rstPartNo.BOF And Not rstPartNo.EOF Then Txt_Validate 0, False
End Select
End Sub
