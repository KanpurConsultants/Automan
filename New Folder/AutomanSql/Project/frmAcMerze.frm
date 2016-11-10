VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAcMerze 
   Appearance      =   0  'Flat
   BackColor       =   &H00CFE0E0&
   Caption         =   "A/C MERGE "
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10200
   DrawStyle       =   5  'Transparent
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton CmdMerze 
      Caption         =   "MERGE"
      Height          =   375
      Left            =   870
      TabIndex        =   18
      Top             =   2760
      Width           =   1860
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
      Left            =   5625
      TabIndex        =   5
      Top             =   2160
      Width           =   3390
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
      Left            =   825
      TabIndex        =   4
      Top             =   2160
      Width           =   3360
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
      Left            =   5610
      TabIndex        =   3
      Top             =   1845
      Width           =   3390
   End
   Begin VB.Frame FrParty 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   2460
      Left            =   5550
      TabIndex        =   7
      Top             =   2460
      Visible         =   0   'False
      Width           =   4575
      Begin MSDataGridLib.DataGrid DgParty 
         Height          =   2100
         Left            =   15
         TabIndex        =   8
         Top             =   330
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   3704
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         ColumnHeaders   =   0   'False
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
         ColumnCount     =   1
         BeginProperty Column00 
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
               ColumnWidth     =   4004.788
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Party Help"
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
         Height          =   270
         Index           =   1
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   4515
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
      Index           =   2
      Left            =   825
      TabIndex        =   2
      Top             =   1845
      Width           =   3360
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
      Left            =   5610
      TabIndex        =   1
      Top             =   1530
      Width           =   3390
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
      Left            =   825
      TabIndex        =   0
      Top             =   1530
      Width           =   3360
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Note : It is necessary to take backup before running the utility.All Nodes should also be            Shut Down Before A/C Merging."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   345
      TabIndex        =   19
      Top             =   3555
      Width           =   9015
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Merging"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   285
      Width           =   5505
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Account"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6465
      TabIndex        =   16
      Top             =   1065
      Width           =   2025
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Source Account"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1425
      TabIndex        =   15
      Top             =   1035
      Width           =   2040
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "------------------>>"
      Height          =   240
      Left            =   4245
      TabIndex        =   14
      Top             =   2175
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "------------------>>"
      Height          =   240
      Left            =   4245
      TabIndex        =   13
      Top             =   1875
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "------------------>>"
      Height          =   240
      Left            =   4245
      TabIndex        =   12
      Top             =   1545
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   1
      Left            =   30
      TabIndex        =   11
      Top             =   1830
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   0
      Left            =   45
      TabIndex        =   10
      Top             =   2145
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   4
      Left            =   30
      TabIndex        =   6
      Top             =   1530
      Width           =   435
   End
End
Attribute VB_Name = "frmAcMerze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsParty As ADODB.Recordset

Private Const Source1 As Byte = 0
Private Const Dest1 As Byte = 1
Private Const Source2 As Byte = 2
Private Const Dest2 As Byte = 3
Private Const Source3 As Byte = 4
Private Const Dest3 As Byte = 5

Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

Private Sub CmdMerze_Click()
'MERGING First A/C

Dim RST1 As ADODB.Recordset, i As Double, TmpCon As ADODB.Connection, TmpCon1 As ADODB.Connection
Dim Path As String, Path1 As String

Set TmpCon = GCn
Set TmpCon1 = G_FaCn

Set RST1 = TmpCon.Execute("Select distinct * from AcMerge")
If RST1.RecordCount > 0 Then
    'Merging A/C 1
    If txt(Source1).Tag <> "" And txt(Dest1).Tag <> "" Then
        RST1.MoveFirst
        For i = 1 To RST1.RecordCount
            If RST1!Database = "Automan" Then
                TmpCon.Execute ("Update " & RST1!Table & " set " & RST1!FldName & " = '" & txt(Dest1).Tag & "' where " & RST1!FldName & " = '" & txt(Source1).Tag & "'")
            ElseIf RST1!Database = "FaDATA" Then
                TmpCon1.Execute ("Update " & RST1!Table & " set " & RST1!FldName & " = '" & txt(Dest1).Tag & "' where " & RST1!FldName & " = '" & txt(Source1).Tag & "'")
            End If
            RST1.MoveNext
        Next
    End If
    'Merging A/C 2
    If txt(Source2).Tag <> "" And txt(Dest2).Tag <> "" Then
        RST1.MoveFirst
        For i = 1 To RST1.RecordCount
            If RST1!Database = "Automan" Then
                TmpCon.Execute ("Update " & RST1!Table & " set " & RST1!FldName & " = '" & txt(Dest2).Tag & "' where " & RST1!FldName & " = '" & txt(Source2).Tag & "'")
            ElseIf RST1!Database = "FaDATA" Then
                TmpCon1.Execute ("Update " & RST1!Table & " set " & RST1!FldName & " = '" & txt(Dest2).Tag & "' where " & RST1!FldName & " = '" & txt(Source2).Tag & "'")
            End If
            RST1.MoveNext
        Next
    End If
        
    'Merging A/C 3
    If txt(Source3).Tag <> "" And txt(Dest3).Tag <> "" Then
        RST1.MoveFirst
        For i = 1 To RST1.RecordCount
            If RST1!Database = "Automan" Then
                TmpCon.Execute ("Update " & RST1!Table & " set " & RST1!FldName & " = '" & txt(Dest3).Tag & "' where " & RST1!FldName & " = '" & txt(Source3).Tag & "'")
            ElseIf RST1!Database = "FaDATA" Then
                TmpCon1.Execute ("Update " & RST1!Table & " set " & RST1!FldName & " = '" & txt(Dest3).Tag & "' where " & RST1!FldName & " = '" & txt(Source3).Tag & "'")
            End If
            RST1.MoveNext
        Next
    End If
    
MsgBox "Merging Completed Successfully ! Please Run Current Balance Updation for Balances."
End If
Set RST1 = Nothing
Set TmpCon = Nothing
Set TmpCon1 = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
'FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_Load()
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "select SubCOde as code,NAME from SubGroup order by name", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set RsParty = Nothing
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus (Index)
Grid_Hide
Select Case Index
    Case Source1, Dest1, Source2, Dest2, Source3, Dest3
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "code ='" & txt(Index).Tag & "'"
        End If
End Select
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Byte
Dim Txtdate As Boolean
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case Source1, Dest1, Source2, Dest2, Source3, Dest3
         DGridTxtKeyDown FrParty, txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
         FrParty.Tag = Index
End Select
        If FrParty.Visible = False Then
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> Dest3 Then
               Ctrl_DownKeyDown KeyCode, Shift
            End If
            If KeyCode = vbKeyUp And Index <> Source1 Then Ctrl_UpKeyDown KeyCode, Shift
        End If

End Sub
Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown->KeyPress->KeyUp
'Validate->LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case Source1, Dest1, Dest2, Source2, Source3, Dest3
        If FrParty.Visible = True Then DGridTxtKeyPress txt, Index, RsParty, KeyAscii, "Name"
End Select
End Sub
Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate (Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case Source1, Dest1, Dest2, Source2, Source3, Dest3
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then
        Else
            txt(Index).TEXT = RsParty!Name
            txt(Index).Tag = RsParty!Code
        End If
End Select
End Sub
Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        txt(FrParty.Tag).TEXT = RsParty!Name
        txt(FrParty.Tag).Tag = RsParty!Code
    End If
    FrParty.Visible = False
    txt(FrParty.Tag).SetFocus
End Sub
'******* Fuctions **********
Private Sub BlankText()
Dim i As Byte
For i = 0 To txt.Count - 1
    txt(i).TEXT = ""
Next i
End Sub
Private Sub Ini_Grid()
Dim i As Byte
FrParty.left = txt(Source1).left + txt(Source1).width + 10: FrParty.top = txt(Source1).top
End Sub
Private Sub Grid_Hide()
    If FrParty.Visible = True Then FrParty.Visible = False
End Sub
Private Sub Ctrl_GetFocus(Index As Integer)
    txt(Index).BackColor = CtrlBCol
    txt(Index).ForeColor = CtrlFCol
    txt(Index).BorderStyle = 1
End Sub
Private Sub Ctrl_validate(Index As Integer)
    txt(Index).BackColor = CtrlBColOrg
    txt(Index).ForeColor = CtrlFColOrg
    txt(Index).BorderStyle = 0
End Sub
