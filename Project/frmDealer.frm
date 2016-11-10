VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmDealer 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Dealer Master"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
   End
   Begin VB.Frame FrDeal 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   4185
      TabIndex        =   23
      Top             =   4635
      Visible         =   0   'False
      Width           =   5220
      Begin MSDataGridLib.DataGrid DGDeal 
         Height          =   3225
         Left            =   30
         TabIndex        =   24
         Top             =   345
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   5689
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   -2147483648
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   13504523
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "D_Code"
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
            DataField       =   "D_Name"
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
            DataField       =   "D_Code"
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
               DividerStyle    =   0
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   3089.764
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "List of Dealer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Index           =   1
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   5175
      End
   End
   Begin VB.Frame FrCity 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   0
      TabIndex        =   20
      Top             =   4635
      Visible         =   0   'False
      Width           =   4095
      Begin MSDataGridLib.DataGrid DGCity 
         Height          =   3225
         Left            =   30
         TabIndex        =   21
         Top             =   345
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   5689
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   -2147483648
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   13504523
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "CityCode"
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
            DataField       =   "CityName"
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
            DataField       =   "CityHelp"
            Caption         =   "CityHelp"
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
               DividerStyle    =   0
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   3435.024
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "List of City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Index           =   0
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   4050
      End
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
      Index           =   10
      Left            =   2805
      MaxLength       =   30
      TabIndex        =   19
      Top             =   3540
      Width           =   3390
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
      Left            =   2805
      MaxLength       =   30
      TabIndex        =   10
      Top             =   3255
      Width           =   3390
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
      Left            =   2805
      MaxLength       =   25
      TabIndex        =   9
      Top             =   2970
      Width           =   3390
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
      Left            =   2805
      MaxLength       =   6
      TabIndex        =   8
      Top             =   2685
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
      Left            =   2805
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2400
      Width           =   3390
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
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   6
      Top             =   2115
      Width           =   5205
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
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1830
      Width           =   5205
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
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1545
      Width           =   5205
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
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1260
      Width           =   5205
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
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   2
      Top             =   975
      Width           =   5205
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
      Left            =   2805
      MaxLength       =   7
      TabIndex        =   1
      Top             =   690
      Width           =   1065
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LST No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   6
      Left            =   1185
      TabIndex        =   18
      Top             =   3570
      Width           =   735
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CST No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   5
      Left            =   1185
      TabIndex        =   17
      Top             =   3285
      Width           =   765
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   4
      Left            =   1185
      TabIndex        =   16
      Top             =   3000
      Width           =   600
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pin Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   1185
      TabIndex        =   15
      Top             =   2715
      Width           =   825
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   0
      Left            =   1185
      TabIndex        =   14
      Top             =   1005
      Width           =   1395
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   1185
      TabIndex        =   13
      Top             =   1290
      Width           =   1425
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   1185
      TabIndex        =   12
      Top             =   2430
      Width           =   930
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   4
      Left            =   1185
      TabIndex        =   11
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmDealer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public MasterFormExit As Boolean
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstCity As ADODB.Recordset
Dim mFlag As Byte
Private Const D_Code = 0, D_Name = 1, D_Add1 = 2, D_Add2 = 3, D_Add3 = 4, D_Add4 = 5, CityName = 6, Pin = 7, District = 8, CST = 9, LST = 10

Private Sub DGCity_Click()
    DGCity_KeyDown vbKeyReturn, 0
End Sub

'Private Sub DGCity_DblClick()
'    Txt(CityCode).Text = RstCity!CityName
'    Txt(CityCode).Tag = RstCity!CityCode
'    DGCity_KeyDown 13, 0
'End Sub

Private Sub DGCity_KeyDown(KeyCode As Integer, Shift As Integer)
If RstCity.BOF = True Or RstCity.EOF = True Then Exit Sub
If KeyCode = vbKeyEscape Then
    txt(CityName).TEXT = ""
Else
    txt(CityName).TEXT = RstCity!CityName
    If KeyCode = vbKeyReturn Then
        If RstCity.RecordCount > 0 Then
            txt(CityName).SetFocus
        End If
    End If
End If

End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
    Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_Load()

TopCtrl1.Tag = PubUParam: TopCtrl1.TopText1 = "Dealer Master"   ': TopCtrl1.TopText1.Width = 1000
Set RstMain = New ADODB.Recordset
RstMain.CursorLocation = adUseClient
If PubMoveRecYn Then
    RstMain.Open "Select * From AMD_Dealer Order by D_Name", GCn, adOpenDynamic, adLockOptimistic
Else
    RstMain.Open "Select Top 1 * From AMD_Dealer Order by D_Name", GCn, adOpenDynamic, adLockOptimistic
End If
Set RstHelp = New ADODB.Recordset

RstHelp.Open "Select D_CODE,D_NAME FROM AMD_Dealer Order by D_Name", GCn, adOpenDynamic, adLockOptimistic
Set RstCity = New ADODB.Recordset
RstCity.Open "Select CITYCODE,CITYNAME FROM City Order by CityName", GCn, adOpenDynamic, adLockOptimistic
CtrlClckCol
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0:    mFlag = 0
Set DGDeal.DataSource = RstHelp
FrCity.Visible = False
Set DgCity.DataSource = RstCity
FrCity.Visible = False
WinSetting Me   ', 6795
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing: Set RstCity = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo Errloop
BlankText
Disp_Text SETS("ADD", Me, RstMain)
txt(D_Code).Tag = txt(D_Code)
Txt_GotFocus D_Code
ADDFLAG = 1
txt(D_Code).SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    txt(D_Code).Enabled = False
    txt(D_Name).Tag = txt(D_Name)
    Txt_GotFocus D_Name
    ADDFLAG = 2
    txt(D_Name).SetFocus
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
            GCn.Execute ("delete * from AMD_DEALER where D_CODE= '" & RstMain!D_Code) & "'"
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
    GSQL = "select D_Code as SearchCode,D_CODE,D_NAME FROM AMD_Dealer Order by D_Name"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("D_Code='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select  * From AMD_Dealer Where D_Code = '" & MyValue & "' Order by D_Name")
    End If
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim rep As CrystalReport, Form1 As frmMastList
    Set Form1 = New frmMastList
    With Form1
        .g_FormID = 13
        .LblName.CAPTION = Me.CAPTION
        .CAPTION = Me.CAPTION
        .Show
    End With
    Set Form1 = Nothing
    Set rep = Nothing
End Sub
Private Sub TopCtrl1_eSave()
Dim transFlag As Byte
On Error GoTo Errloop
    transFlag = 0
    If IsValid(txt(D_Code), "Dealer Code") = False Then Txt_GotFocus D_Code: Exit Sub
    If IsValid(txt(D_Name), "Dealer Name") = False Then Txt_GotFocus D_Name: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From AMD_Dealer Where D_Code=" & Chk_Text(Trim(txt(D_Code))) & " AND SITE_CODE='" & PubSiteCode & "'").Fields(0) > 0 Then MsgBox "Dealer Code Already Exists", vbInformation, "Duplicate Checking": Txt_GotFocus D_Code: txt(D_Code).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If ADDFLAG = 1 Then
        GCn.Execute ("DELETE From AMD_Dealer Where D_Code=" & Chk_Text(txt(D_Code)) & " AND SITE_CODE='" & PubSiteCode & "'")
        GCn.Execute ("Insert Into AMD_Dealer(D_Code,Site_Code,Div_Code,D_Name,D_ADD1,D_ADD2,D_ADD3,D_ADD4,D_CITY,D_PIN_CODE,D_DIST,D_CST_NO,D_RST_NO,U_Name,U_EntDt,U_AE) Values('" & txt(D_Code) & "','" & PubSiteCode & "','" & PubDivCode & "'," & Chk_Text(txt(D_Name)) & "," & Chk_Text(txt(D_Add1)) & "," & Chk_Text(txt(D_Add2)) & "," & Chk_Text(txt(D_Add3)) & "," & Chk_Text(txt(D_Add4)) & "," & Chk_Text(txt(CityName)) & "," & Chk_Text(txt(Pin)) & "," & Chk_Text(txt(District)) & "," & Chk_Text(txt(CST)) & "," & Chk_Text(txt(LST)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    ElseIf ADDFLAG = 2 Then
        GCn.Execute ("UPDATE AMD_Dealer SET D_Name=" & Chk_Text(txt(D_Name)) & ",D_ADD1=" & Chk_Text(txt(D_Add1)) & ",D_ADD2=" & Chk_Text(txt(D_Add2)) & ",D_ADD3=" & Chk_Text(txt(D_Add3)) & ",D_ADD4=" & Chk_Text(txt(D_Add4)) & ",D_CITY=" & Chk_Text(txt(CityName)) & ",D_PIN_CODE=" & Chk_Text(txt(Pin)) & ",D_DIST=" & Chk_Text(txt(District)) & ",D_CST_NO=" & Chk_Text(txt(CST)) & ",D_RST_NO=" & Chk_Text(txt(LST)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "' Where D_Code=" & Chk_Text(txt(D_Code)) & "")
    End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select  * From AMD_Dealer Where D_Code = " & Chk_Text(Trim(txt(D_Code))) & " Order by D_Name")
    End If
    RstHelp.Requery
    RstMain.FIND ("D_Code=" & Chk_Text(Trim(txt(D_Code))))
    If ADDFLAG = 1 Then
        BlankText
        Txt_GotFocus D_Code
        txt(D_Code).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        FrCity.Visible = False
        FrDeal.Visible = False
    End If
Exit Sub
Errloop:    If transFlag = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, RstMain)
        Me.ActiveControl.SetFocus
        MoveRec
        CtrlClckCol
        FrDeal.Visible = False
        FrCity.Visible = False
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    txt(D_Code).BackColor = CtrlBColOrg:      txt(D_Code).ForeColor = CtrlFColOrg
    txt(D_Name).BackColor = CtrlBColOrg:      txt(D_Name).ForeColor = CtrlFColOrg
    txt(D_Add1).BackColor = CtrlBColOrg:     txt(D_Add1).ForeColor = CtrlFColOrg
    txt(D_Add2).BackColor = CtrlBColOrg:     txt(D_Add2).ForeColor = CtrlFColOrg
    txt(D_Add3).BackColor = CtrlBColOrg:     txt(D_Add3).ForeColor = CtrlFColOrg
    txt(D_Add4).BackColor = CtrlBColOrg:     txt(D_Add4).ForeColor = CtrlFColOrg
    txt(CityName).BackColor = CtrlBColOrg:     txt(CityName).ForeColor = CtrlFColOrg
    txt(Pin).BackColor = CtrlBColOrg:     txt(Pin).ForeColor = CtrlFColOrg
    txt(District).BackColor = CtrlBColOrg:     txt(District).ForeColor = CtrlFColOrg
    txt(CST).BackColor = CtrlBColOrg:     txt(CST).ForeColor = CtrlFColOrg
    txt(LST).BackColor = CtrlBColOrg:     txt(LST).ForeColor = CtrlFColOrg
End Sub

Private Sub MoveRec()
On Error GoTo Errloop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    txt(D_Code) = XNull(RstMain!D_Code)
    txt(D_Name) = XNull(RstMain!D_Name)
    txt(D_Add1) = XNull(RstMain!D_Add1)
    txt(D_Add2) = XNull(RstMain!D_Add2)
    txt(D_Add3) = XNull(RstMain!D_Add3)
    txt(D_Add4) = XNull(RstMain!D_Add4)
    txt(CityName) = XNull(RstMain!D_City)
    txt(Pin) = XNull(RstMain!D_PIN_CODE)
    txt(District) = XNull(RstMain!D_DIST)
    txt(CST) = XNull(RstMain!D_CST_NO)
    txt(LST) = XNull(RstMain!D_RST_NO)
End If
TopCtrl1.tDel = False
Exit Sub
Errloop:        MsgBox err.Description
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
    RstCity.Requery
End Sub
Private Sub TopCtrl1_eExit()
    RstMain.Cancel
    Unload Me
End Sub

Private Sub DealCodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "D_CODE >=" & Chk_Text(XNull(Trim(txt(D_Code))))
End Sub
Private Sub DealNameSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "D_Name >=" & Chk_Text(XNull(txt(D_Name)))
End Sub
Private Sub cityNameSearch()
If RstCity.RecordCount <= 0 Then Exit Sub
RstCity.MoveFirst
RstCity.FIND "CITYName >=" & Chk_Text(XNull(txt(CityName)))
If Not RstCity.EOF Then
    If mID(RstCity!CityName, 1, Len(Trim(XNull(txt(CityName))))) <> Trim(XNull(txt(CityName))) Then
        CityNameExSearch
    End If
Else
    CityNameExSearch
End If
End Sub
Private Sub CityNameExSearch()
Dim tempRst As ADODB.Recordset
Set tempRst = RstCity.Clone
tempRst.Sort = "Citycode ASC"
tempRst.FIND "Cityname >='" & FilterString(XNull(txt(CityName))) & "'"
If Not tempRst.EOF Then
    RstCity.MoveFirst
    RstCity.FIND "CITYNAME >='" & XNull(tempRst!CityName) & "'"
End If
Set tempRst = Nothing
End Sub

Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case D_Code, D_Name
            If RstHelp.RecordCount = 0 Then Exit Sub
            If FrDeal.Visible = True Then FrDeal.Visible = False
            FrDeal.Visible = True
            FrDeal.top = txt(Index).top + txt(Index).height + 10
            FrDeal.left = txt(Index).left
            FrDeal.ZOrder 0
        Case CityName
            If RstCity.RecordCount = 0 Then Exit Sub
            If FrDeal.Visible = True Then FrDeal.Visible = False
            FrCity.Visible = True
            FrCity.top = txt(Index).top + txt(Index).height + 10
            FrCity.left = txt(Index).left
            FrCity.ZOrder 0
    End Select
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
DGDeal.Columns(0).width = 1000.1: DGDeal.Columns(1).width = 3535.024: DGDeal.Columns(2).width = 1000.1
Dim mBookMark
    Ctrl_GetFocus txt(Index)
mFlag = 0
    If FrDeal.Visible = True Then FrDeal.Visible = False
    If FrCity.Visible = True Then FrCity.Visible = False
    RST_BOF_EOF RstHelp
    txt(Index).Tag = txt(Index)
'    Txt_Click Index
    Select Case Index
        Case D_Code, D_Name
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
        Case CityName
            If RstCity.BOF Or RstCity.EOF Then Exit Sub
    End Select
    Select Case Index
        Case D_Code
            DGDeal.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "D_CODE ASC"
            RstHelp.Bookmark = mBookMark
            DealCodeSearch
        Case D_Name
            DGDeal.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "D_NAME ASC"
            RstHelp.Bookmark = mBookMark
            DealNameSearch
        Case CityName
            DgCity.Columns(0).width = 0: DgCity.Columns(2).width = 0
            mBookMark = RstCity.Bookmark
            RstCity.Sort = "CITYNAME ASC"
            RstCity.Bookmark = mBookMark
            cityNameSearch
    End Select
    If txt(Index) = "" Then Txt_Change Index
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer
'If KeyCode = vbKeyEscape Then Txt(Index).Text = ""
Select Case Index
    Case CityName
        If FrCity.Visible = True Then
            Select Case KeyCode
                Case vbKeyUp
                    If Not RstCity.BOF Then RstCity.MovePrevious
                Case vbKeyDown
                    If Not RstCity.EOF Then RstCity.MoveNext
                Case 33
                    For I = 1 To 9
                        If Not RstCity.BOF Then RstCity.MovePrevious
                    Next
                Case 34
                    For I = 1 To 9
                        If Not RstCity.EOF Then RstCity.MoveNext
                    Next
                Case 13
                    SendKeysA vbKeyTab, True
            End Select
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, 33, 34
                    RST_BOF_EOF RstCity
                    If Not RstCity.BOF And Not RstCity.EOF Then
                        txt(CityName) = XNull(RstCity!CityName)
                        txt(CityName).SelStart = 0
                    End If
            End Select
        End If
    Case D_Add1, D_Add2, D_Add3, D_Add4, District, CST
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case Pin
        NumDown txt(Pin), KeyCode, 6, 0
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case LST
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                Txt_Validate Index, result
                If result = True Then Txt_GotFocus Index: txt(Index).SetFocus: Exit Sub
                TopCtrl1_eSave
            Else
'                Txt_Click Index
                Txt_GotFocus Index
                txt(Index).SetFocus
            End If
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If

End Select
Select Case Index
    Case D_Code
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        'ElseIf KeyCode = vbKeyUp Then
        '    SendKeys "+{Tab}"
        '    KeyCode = 0
        End If
    Case D_Name
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case CityName
        If FrCity.Visible = False Then
            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
                SendKeysA vbKeyTab, True
                KeyCode = 0
            ElseIf KeyCode = vbKeyUp Then
                SendKeys "+{Tab}"
                KeyCode = 0
            End If
        End If
End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)
If Index = Pin Then NumPress txt(Pin), keyascii, 6, 0
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
mFlag = 0
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case D_Code
        DealCodeSearch
    Case D_Name
        DealNameSearch
    Case CityName
        cityNameSearch
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case D_Code
            Set Rst = GCn.Execute("SELECT * FROM AMD_Dealer WHERE D_CODE=" & Chk_Text(txt(D_Code)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Dealer Code Already Exists", vbInformation, "Validation": txt(D_Code) = txt(D_Code).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!D_Code <> RstMain!D_Code Then MsgBox "Dealer Code Already Exists", vbInformation, "Validation": txt(D_Code) = txt(D_Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case D_Name
            Set Rst = GCn.Execute("SELECT * FROM AMD_Dealer WHERE D_NAME=" & Chk_Text(txt(D_Name)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Dealer Name Already Exists", vbInformation, "Validation": txt(D_Name) = txt(D_Name).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!D_Name <> RstMain!D_Name Then MsgBox "Dealer Name Already Exists", vbInformation, "Validation": txt(D_Name) = txt(D_Name).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case CityName
            If Not RstCity.EOF And Not RstCity.BOF Then
                txt(CityName) = XNull(RstCity!CityName)
            Else
                txt(CityName) = "": txt(CityName) = ""
            End If
    End Select
Set Rst = Nothing
End Sub
Private Sub DGCITY_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mFlag = 1 Then
    txt(CityName) = DgCity.Columns(1).TEXT
End If
End Sub

Private Sub DGCITY_GotFocus()
    mFlag = 1
End Sub

Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
Next I
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
'CmbOrder.Enabled = IIf(AddFlag = 1, True, False)
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
Next
End Sub

'Private Sub Ini_Grid()
'    FGrid.RowHeightMin = 250
'    FGrid.ColWidth(25) = 0
'End Sub

