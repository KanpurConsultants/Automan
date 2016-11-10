VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmUserPass 
   BackColor       =   &H00BAD3C9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Password"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
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
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2250
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   645
      Width           =   1335
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
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2250
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   375
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00CFE0E0&
      Caption         =   "&Cancel "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2130
      Width           =   1155
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00CFE0E0&
      Caption         =   " &Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2130
      Width           =   1155
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
      Index           =   0
      Left            =   2250
      TabIndex        =   0
      Top             =   105
      Width           =   4905
   End
   Begin MSDataGridLib.DataGrid DGHelp 
      Height          =   1575
      Left            =   1170
      Negotiate       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   0   'False
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
            ColumnWidth     =   4770.142
         EndProperty
      EndProperty
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password "
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
      Left            =   540
      TabIndex        =   7
      Top             =   660
      Width           =   1590
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Password"
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
      Left            =   540
      TabIndex        =   6
      Top             =   390
      Width           =   1290
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Left            =   540
      TabIndex        =   5
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "FrmUserPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsUser As ADODB.Recordset
Private Const UName As Byte = 0
Private Const UPass As Byte = 1
Private Const CPass As Byte = 2

Private Sub Cmdsave_Click(Index As Integer)
If IsValid(Txt(UName), "UserName") = False Then Exit Sub
Select Case Index
Case 0
If Txt(UPass) <> Txt(CPass) Then
    MsgBox "Confirm Password should be same with Password", vbInformation
    Txt(CPass).SetFocus
    Exit Sub
End If
G_CompCn.Execute ("Update usermast set passwd= '" & CODIFY(Trim(Txt(UPass))) & "' where USER_NAME = '" & Txt(UName).TEXT & "'")
    If MsgBox("Record Save ! Want to change more ? ", vbYesNo) = vbYes Then
        RsUser.Requery
        Txt(UName).SetFocus
    Else
        Unload Me
    End If
Case 1
    Unload Me
End Select
End Sub


Private Sub Form_Load()
WinSetting Me, 2910, 7530
If pubUName = "SA" Then
    Set RsUser = G_CompCn.Execute("select USER_NAME as Name,PASSWD from  UserMast")
Else
    Set RsUser = G_CompCn.Execute("select USER_NAME as Name,PASSWD from  UserMast where USER_NAME = '" & pubUName & "'")
End If
Set DGHelp.DataSource = RsUser
DGHelp.left = Txt(UName).left
DGHelp.top = Txt(UName).top + Txt(UName).height + 10
End Sub

Private Sub DGHelp_Click()
    DGHelp.Visible = False
    If RsUser.RecordCount > 0 Then
        Txt(UName).TEXT = RsUser!Name
        Txt(UPass) = DCODIFY(XNull(RsUser!PASSWD))
        Txt(CPass) = DCODIFY(XNull(RsUser!PASSWD))
    End If
    Txt(UName).SetFocus
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus Txt(Index)
Grid_Hide
Select Case Index
    Case UName
        If RsUser.RecordCount = 0 Or (RsUser.EOF = True Or RsUser.BOF = True) Then Exit Sub
        If Txt(Index).TEXT <> RsUser!Name Then
            RsUser.MoveFirst
            RsUser.FIND "name ='" & Txt(Index).TEXT & "'"
        End If
   
End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case UName
        DGridTxtKeyDown DGHelp, Txt, Index, RsUser, KeyCode, False, 0
End Select
If DGHelp.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If Index <> UName Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)
Select Case Index
    Case UName
        If DGHelp.Visible = True Then DGridTxtKeyPress Txt, Index, RsUser, keyascii, "Name"
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim i As Integer
Select Case Index
     Case UName
        If IsValid(Txt(Index), "User Name") = False Then Cancel = True: Exit Sub
        If RsUser.RecordCount = 0 Or (RsUser.EOF = True Or RsUser.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(UPass) = ""
            Txt(CPass) = ""
        Else
            Txt(Index).TEXT = RsUser!Name
            Txt(UPass) = DCODIFY(XNull(RsUser!PASSWD))
            Txt(CPass) = DCODIFY(XNull(RsUser!PASSWD))
        End If
    Case UPass, CPass
        If Txt(UPass) <> Txt(CPass) And Txt(UPass) <> "" And Txt(CPass) <> "" Then
            MsgBox "Confirm Password should be same with Password", vbInformation
            Txt(Index).SetFocus
            Cancel = True
        End If
End Select
Set Rst = Nothing
End Sub
Private Function DCODIFY(Txt As String) As String
    Dim xxx As String
    Dim xx As Integer, MyVal As Integer
    If Txt = "" Then DCODIFY = "": Exit Function
    MyVal = Asc(left(Txt, 1)) - 27
    xxx = ""
    For xx = 1 To Len(Txt) - 1
        xxx = xxx + Chr(Asc(mID(Txt, xx + 1, 1)) - 27 - MyVal)
    Next
    DCODIFY = xxx
End Function
Private Function CODIFY(Txt As String) As String
    Dim xxx As String
    Dim xx As Integer, MyVal As Integer
    Randomize
    MyVal = Int((99 * Rnd) + 1)
    xxx = Chr(MyVal + 27)
    For xx = 1 To Len(Txt)
        xxx = xxx + Chr(Asc(mID(Txt, xx, 1)) + 27 + MyVal)
    Next
    CODIFY = xxx
End Function
Private Sub Grid_Hide()
    If DGHelp.Visible = True Then DGHelp.Visible = False
End Sub
