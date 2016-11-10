VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompany 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame FrmAddComp 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2940
      Left            =   3495
      TabIndex        =   4
      Top             =   3540
      Width           =   7230
      Begin MSDataGridLib.DataGrid DBComp 
         Height          =   2370
         Left            =   3285
         TabIndex        =   23
         Top             =   495
         Visible         =   0   'False
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   4180
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Comp_Code"
            Caption         =   "Comp_Code"
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
            DataField       =   "Comp_Name"
            Caption         =   "Comp_Name"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   2745
         TabIndex        =   19
         Top             =   2475
         Visible         =   0   'False
         Width           =   3150
      End
      Begin VB.Timer ProgTimer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   300
      End
      Begin MSComctlLib.ProgressBar ProgBar1 
         DragMode        =   1  'Automatic
         Height          =   210
         Left            =   285
         TabIndex        =   18
         Top             =   2625
         Visible         =   0   'False
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   2280
         TabIndex        =   3
         Top             =   1680
         Width           =   1140
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1410
         Width           =   4500
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1140
         Width           =   405
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1950
         Width           =   3150
      End
      Begin VB.Label LblColon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   2070
         TabIndex        =   22
         Top             =   1965
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DataBase From"
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
         Left            =   285
         TabIndex        =   21
         Top             =   1950
         Width           =   1305
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   285
         TabIndex        =   20
         Top             =   2385
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   2220
         Width           =   4500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Code"
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
         Left            =   285
         TabIndex        =   14
         Top             =   1155
         Width           =   1290
      End
      Begin VB.Label LblColon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   2070
         TabIndex        =   13
         Top             =   1155
         Width           =   45
      End
      Begin VB.Label LblColon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   2085
         TabIndex        =   12
         Top             =   2205
         Width           =   45
      End
      Begin VB.Label LblColon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   2070
         TabIndex        =   11
         Top             =   1425
         Width           =   45
      End
      Begin VB.Label LblColon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   2085
         TabIndex        =   10
         Top             =   1695
         Width           =   45
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Company Information"
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
         Index           =   6
         Left            =   2505
         TabIndex        =   15
         Top             =   150
         Width           =   2580
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name  "
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
         Left            =   285
         TabIndex        =   7
         Top             =   1425
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Central  Data  Path "
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
         Height          =   270
         Index           =   3
         Left            =   285
         TabIndex        =   6
         Top             =   2220
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year Start Date "
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
         Left            =   270
         TabIndex        =   5
         Top             =   1695
         Width           =   1275
      End
   End
   Begin VB.Frame FrmLv 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3420
      Left            =   150
      TabIndex        =   8
      Top             =   150
      Width           =   7740
      Begin MSDataGridLib.DataGrid Grid 
         Height          =   2400
         Left            =   60
         TabIndex        =   0
         Top             =   660
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   4233
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ColumnHeaders   =   -1  'True
         ForeColor       =   16711680
         HeadLines       =   0
         RowHeight       =   15
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
            DataField       =   "Comp_Code"
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
            DataField       =   "COMP_NAME"
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
         BeginProperty Column02 
            DataField       =   "CentralData_Path"
            Caption         =   "Group"
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
            DataField       =   "START_DATE"
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
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3825.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1275.024
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "     Code                      Group Company Name                    Group           Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   60
         TabIndex        =   16
         Top             =   405
         Width           =   7590
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Group Companies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Index           =   3
         Left            =   60
         TabIndex        =   9
         Top             =   150
         Width           =   7590
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "dataman"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   6825
      TabIndex        =   25
      Top             =   3675
      Width           =   1080
      WordWrap        =   -1  'True
   End
   Begin VB.Menu POPUP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu MnuAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu MnuDel 
         Caption         =   "&Delete"
      End
      Begin VB.Menu Dash1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu MnuCancel 
         Caption         =   "&Cancel"
      End
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADDFLAG As Integer, RsComp As ADODB.Recordset, CompCode As String * 2
'Private Const CtrlBColOrg = &HCFE0E0                     'Orginal BackColour
'Private Const CtrlFColOrg = &H80000008                   'Orginal ForeColour
'Private Const CtrlBCol = &H0&                            'Changed BackColour
'Private Const CtrlFCol = &HFFFF&                         'Changed ForeColour
Private Const CCode = 0                      'Company Code
Private Const CName = 1                      'Company Name
Private Const SDate = 2                      'Company Start Date
Private Const CtrlPath = 3                   'Central Data Path

Private Sub DBComp_Click()
    Txt(4).TEXT = DBComp.TEXT
    DBComp.Visible = False
    Call Txt_KeyDown(4, vbKeyReturn, 1)
End Sub

Private Sub DBComp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Txt(4).TEXT = DBComp.TEXT
        DBComp.Visible = False
        Call Txt_KeyDown(4, vbKeyReturn, 1)
    End If
    If KeyCode = vbKeyEscape Then
        DBComp.Visible = False
        Txt(4).SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then
    If FrmAddComp.Visible = False Then
        End
    End If
End If
Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
For I = 0 To 3
    Txt(I).BackColor = CtrlBColOrg
    Txt(I).ForeColor = CtrlFColOrg
Next
Me.CAPTION = PubPackage & "-Group Company"
MENUENABLE True
FrmAddComp.left = (Me.width - FrmAddComp.width) / 2
FrmAddComp.top = 105
FrmLv.left = (Me.width - (FrmLv.width + 90)) / 2
FrmLv.top = 240
Grid.left = (FrmLv.width - Grid.width) / 2
Grid.top = 660
DBComp.top = 495
DBComp.left = 3285
DBComp.height = 2370
DBComp.width = 3435
Label4.width = Grid.width
LblHelp(3).width = Grid.width
Label4.left = Grid.left
LblHelp(3).left = Grid.left
FrmAddComp.Visible = False
Set RsComp = New Recordset
If UCase(pubUName) = "SA" Or UCase(pubUName) = "ADMIN" Then
    RsComp.Open "Select * from company order by Start_Date Desc", G_CompCn, adOpenDynamic, adLockOptimistic
Else
    RsComp.Open "Select * from company where comp_code in (select distinct comp_code from user1 where user_name='" & pubUName & "') order by Start_Date Desc", G_CompCn, adOpenDynamic, adLockOptimistic
End If
Set Grid.DataSource = RsComp
Set DBComp.DataSource = RsComp
Call MoveRec
Exit Sub
ELoop:  MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub FrmAddComp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Index As Byte
If Button = 2 Then
    If Index = SDate Then Txt(SDate).TEXT = RetDate(Txt(SDate))
    PopupMenu POPUP
End If
End Sub
Private Sub MnuAdd_Click()
On Error GoTo eloop1
ADDFLAG = 1
FrmLv.Visible = False
FrmAddComp.Visible = True
FrmAddComp.ZOrder 0
Txt(CCode).Enabled = True
Label2.CAPTION = ""
BlankText
If Txt(CCode).Visible = True Then
'    txt(CCode).Visible = True
    Txt(CCode).SetFocus
End If
MENUENABLE False
Exit Sub
eloop1:     Call CheckError
End Sub

Private Sub MnuEdit_Click()
On Error GoTo eloop1
ADDFLAG = 2
FrmLv.Visible = False
Call MoveRec
FrmAddComp.Visible = True
FrmAddComp.ZOrder 0
Txt(CCode).Enabled = False
Txt(CName).SetFocus
MENUENABLE False
Exit Sub
eloop1:     Call CheckError
End Sub

Private Sub MnuDel_Click()
On Error GoTo eloop1
If RsComp.RecordCount <= 0 Then Exit Sub
If G_CompCn.Execute("select count(*) from user1 where comp_code = '" & Txt(CCode).TEXT & "' and div_code <> ''").Fields(CCode).Value > 0 Then
   MsgBox "Division Exist,Company Can't Deleted", vbExclamation, "Delete Error"
   Exit Sub
End If
If MsgBox("Are You Sure To Delete ?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
G_CompCn.Execute "DELETE FROM COMPANY WHERE COMP_CODE='" & Txt(CCode).TEXT & "'"
G_CompCn.Execute ("delete from user1 where comp_code='" & Txt(CCode).TEXT & "'")
    If PubBackEnd = "A" Then
        Kill Pub_DataPath & "\" & Txt(CtrlPath).TEXT & "\*.*"
        RmDir Pub_DataPath & "\" & Txt(CtrlPath).TEXT
    Else
        G_CompCn.Execute "Drop Database " & Txt(CtrlPath)
    End If
    RsComp.Requery
    Grid.Refresh
End If
Exit Sub
eloop1:     CheckError
End Sub

Private Sub MNUCANCEL_Click()
ADDFLAG = 0
MENUENABLE True
FrmAddComp.Visible = False
FrmLv.Visible = True
DBComp.Visible = False

End Sub

Private Sub MNUSAVE_Click()
Dim I As Byte, SourcePath$, DestPath$, SourcePathFA$, DestPathFA$, FAFolder$
Dim FS As FileSystemObject, App_Path$
On Error GoTo ELoop


For I = 0 To 4
    If I <> 3 Then
        If Txt(I).TEXT = "" Then MsgBox Label3(I).CAPTION & "Is A Required Field", vbExclamation, "Input Error":      Exit Sub
    End If
Next
If ADDFLAG = 1 Then If G_CompCn.Execute("select count(*) from company where comp_code = '" & Txt(CCode).TEXT & "'").Fields(CCode).Value > 0 Then MsgBox "Duplicate Company Code", vbExclamation, "Input Error": Exit Sub

G_CompCn.BeginTrans
CompCode = Txt(CCode).TEXT
If ADDFLAG = 1 Then
    '**********
    Label16.Visible = True
    Label16.Refresh
    ProgBar1.Value = 0
    ProgTimer1.Enabled = True
    ProgBar1.Visible = True
    '***********
'COPYING AUTOMAN.MDB TO THE DESTINATION
    Label16.CAPTION = "Please wait ! Preparing Automan.Mdb...."
    SourcePath = Pub_DataPath & "\Auto_" & Txt(4) & "\Automan.mdb"
    DestPath = Pub_DataPath & "\" & Txt(CtrlPath).TEXT & "\Automan.mdb"
        
    Set FS = New FileSystemObject
    FS.CreateFolder (Pub_DataPath & "\" & Txt(CtrlPath).TEXT)
    ProgBar1.Value = 10
    
    FS.CopyFile SourcePath, DestPath, True
    ProgBar1.Value = 35
    
    Set GCn = New ADODB.Connection
    With GCn
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & DestPath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
        .Open
    End With

'COPYING FADATA.MDB TO THE DESTINATION
    Label16.CAPTION = "Please wait ! Preparing FAData.Mdb...."
    FS.CreateFolder (Pub_DataPath & "\" & Txt(CtrlPath).TEXT) & "\FAData_" & Right(Txt(CtrlPath).TEXT, 1) & "04"
        
    SourcePathFA = Pub_DataPath & "\" & GCn.Execute("Select FADataPath from AssociatedFirms").Fields(0).Value & "\FAData.mdb"
    DestPathFA = (Pub_DataPath & "\" & Txt(CtrlPath).TEXT) & "\FAData_" & Right(Txt(CtrlPath).TEXT, 1) & "04\FAData.mdb"
    
    FS.CopyFile SourcePathFA, DestPathFA, True
    ProgBar1.Value = 55
    FAFolder = Txt(CtrlPath).TEXT & "\FAData_" & Right(Txt(CtrlPath).TEXT, 1) & "04"
    
    Label16.CAPTION = "Please wait ! Preparing Company.Mdb...."
'UPDATING Company TABLE
    G_CompCn.Execute "insert into company(comp_code,comp_name,CentralData_Path,start_date) " & " values('" & Txt(CCode).TEXT & "','" & Txt(CName).TEXT & "','" & Txt(CtrlPath).TEXT & "'," & ConvertDate(Txt(SDate).TEXT) & ")"
    G_CompCn.Execute "Update Company set OldPath= 'Auto_" & Txt(4) & "'" & ",OldPathFA='" & GCn.Execute("Select FADataPath from AssociatedFirms").Fields(0).Value & "' where Comp_Code= '" & Txt(CCode).TEXT & "'"

'UPDATING USER1 TABLE
    G_CompCn.Execute "insert into User1(User_Name,comp_code,Div_Code,Div_Name,Mod_Veh,Mod_Spr,Mod_Wsp,Mod_Acc,Mod_Set,PARAM_STR)" & _
                     " values('" & pubUName & "','" & Txt(CCode).TEXT & "','C','CVD',1,1,1,1,1,'*')"
    G_CompCn.Execute "insert into User1(User_Name,comp_code,Div_Code,Div_Name,Mod_Veh,Mod_Spr,Mod_Wsp,Mod_Acc,Mod_Set,PARAM_STR)" & _
                     " values('" & pubUName & "','" & Txt(CCode).TEXT & "','P','PCD',1,1,1,1,1,'*')"
    'For Rashmi Motors
    If UCase(left(G_CompCn.Execute("Select Comp_Name from Company").Fields(0).Value, 6)) = "RASHMI" Then
        G_CompCn.Execute "Update User1 set Mod_Veh=0"
    End If
'UPDATING ASSOCIATEDFIRMS TABLE
    GCn.Execute "Update AssociatedFirms set FADataPath = '" & FAFolder & "',AssoComp_Code='" & Right(Txt(CCode).TEXT, 1) & "' "

'UPDATING DEVISION TABLE
    GCn.Execute "Update Division set V_SecCompCode = '" & Right(Txt(CCode).TEXT, 1) & "',S_SecCompCode = '" & Right(Txt(CCode).TEXT, 1) & "',W_SecCompCode = '" & Right(Txt(CCode).TEXT, 1) & "'"
    GCn.Execute "Update Division set V_SecFADataPath = '" & FAFolder & "',S_SecFADataPath = '" & FAFolder & "',W_SecFADataPath = '" & FAFolder & "'"

    ProgBar1.Value = 70
    If pubUName <> "SA" Then
        G_CompCn.Execute ("insert into user1(user_name,comp_code) values('SA','" & Txt(CCode).TEXT & "')")
    End If
    

    
    ProgBar1.Value = 100
    ProgTimer1.Enabled = False
    Label16.Visible = False
    ProgBar1.Value = 0
    ProgTimer1.Enabled = False
    ProgBar1.Visible = False
Else
    G_CompCn.Execute "update company set comp_name='" & Txt(CName).TEXT & "' where comp_code ='" & Txt(CCode).TEXT & "'"
End If
G_CompCn.CommitTrans
ADDFLAG = 0
MENUENABLE True
RsComp.Requery
Grid.Refresh
MsgBox "New Company Successfully created.Please reload the Software"
End
'FrmAddComp.Visible = False
'FrmLv.Visible = True
Exit Sub
ELoop:
G_CompCn.RollbackTrans
CheckError
End Sub
Private Sub Grid_DblClick()
Call SelCompany
End Sub
Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 If Grid.Visible = True Then
     Call MoveRec
 End If
End Sub
Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And PubULabel = "Y" Then
    MENUENABLE True
    PopupMenu POPUP
End If
End Sub
Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call SelCompany
End Sub

Private Sub ProgTimer1_Timer()
If ProgBar1.Value <= 90 Then
    ProgBar1.Value = ProgBar1.Value + 10
Else
    ProgBar1.Value = 100
End If
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus Txt(Index)
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn Or KeyCode = vbKeyDown) And Index = 4 Then
    If MsgBox("Save Record?", vbYesNo, "Save Entry") = vbYes Then
        Call Txt_Validate(Index, True)
        Call MNUSAVE_Click
        Exit Sub
    Else
        Call MNUCANCEL_Click
        Exit Sub
    End If
End If
If KeyCode = vbKeyReturn Or (KeyCode = vbKeyDown And Index <> SDate) Then   'keydown = 40
    SendKeysA vbKeyTab, True
ElseIf KeyCode = vbKeyUp And _
    ((ADDFLAG = 1 And Index <> CCode) Or (ADDFLAG = 2 And Index <> CName)) Then     'keyup = 38
    SendKeys "+{Tab}"
ElseIf KeyCode = vbKeyEscape Then
    MNUCANCEL_Click
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
If Index = 0 Then
    keyascii = Asc(UCase(Chr(keyascii)))
End If
If Index = 4 Then
    DBComp.Visible = True
    If DBComp.Visible = True Then DBComp.SetFocus
End If
CheckQuote keyascii
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate Txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
If Index = CCode Then Txt(CtrlPath).TEXT = "Auto_" & Txt(CCode).TEXT: Label2.CAPTION = Pub_DataPath & "\" & Txt(CtrlPath).TEXT
If Txt(Index).TEXT = "" Then
    MsgBox Label3(Index) & " Is Required", vbExclamation, "Validation Check"
    Txt(Index).SetFocus
    Cancel = True
    Exit Sub
End If
If Index = SDate Then Txt(SDate).TEXT = RetDate(Txt(SDate))
End Sub

Private Sub BlankText()
Dim I As Byte
For I = 0 To 4
    Txt(I).TEXT = ""
Next I
End Sub
Private Sub SelCompany()
On Error Resume Next
Dim App_Path As String
If RsComp.RecordCount = 0 Then Call MnuAdd_Click: Exit Sub

App_Path = Pub_DataPath & "\" & Txt(CtrlPath).TEXT & "\Automan.mdb"
Set GCn = New ADODB.Connection
With GCn
    .CursorLocation = adUseClient
    If PubBackEnd = "A" Then
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & App_Path & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
    Else
        If XNull(RsComp!DbUser) <> "" Then
            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & XNull(RsComp!DbUser) & ";Password=" & XNull(RsComp!DbPass) & ";Initial Catalog=" & RsComp!CentralData_Path & ";Data Source=" & PubServerName
        Else
            If XNull(RsComp!DbUser) <> "" Then
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & XNull(RsComp!DbUser) & "; Password=" & XNull(RsComp!DbPass) & ";Initial Catalog=" & RsComp!CentralData_Path & ";Data Source=" & PubServerName
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & RsComp!CentralData_Path & ";Data Source=" & PubServerName
            End If
        End If
    End If
    
    .Open
End With
Set FaMasterRst = G_CompCn.Execute("Select * from company where comp_code='" & RsComp!Comp_Code & "'")

PubSiteCodeDisplay = "('A')"
PubSiteName = " "
PubSiteCode = " "
PubCenCompCode = RsComp!Comp_Code
PubCenDataPath = RsComp!CentralData_Path
PubDbUser = XNull(RsComp!DbUser)
PubDbPass = XNull(RsComp!DbPass)

PubStartDate = Format(RsComp!Start_Date, "dd/MMM/yyyy")
PubEndDate = DateAdd("YYYY", 1, PubStartDate) - 1
'**
    Set RsPart = New ADODB.Recordset
    RsPart.CursorLocation = adUseClient
'**

'COde is done for RSO changes.........................
RSOJPR = False
Select Case UCase(left(RsComp!Comp_Name, 5))
    Case "RDB H", "ANAND", "MUNIS", "SUMAN", "RONAK", "CHAMB", _
         "KAMAL", "BHARA", "GREEN", "JAGVIJ", "JAG VI", "URSS ", _
          "SUN C", "GANES", "BALAJ", "LOTUS", "JINDA"
        frmDivision.cmdApply(0).Visible = False: RSOJPR = True
    Case Else
End Select

'****************************************************
frmDivision.CAPTION = PubPackage & "-[" & RsComp!Comp_Code & "]" & RsComp!Comp_Name
'Unload Me

frmDivision.Show 1
PubUParam = "AEDP"
frmCompany.Hide





Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub MENUENABLE(Enb As Boolean)
    frmCompany.MnuAdd.Enabled = Enb
    frmCompany.MnuEdit.Enabled = Enb
    frmCompany.MnuDel.Enabled = Enb
    frmCompany.MnuSave.Enabled = Not Enb
    frmCompany.MnuCancel.Enabled = Not Enb
End Sub
Private Sub MoveRec()
On Error GoTo err
If RsComp.RecordCount > 0 Then
    Txt(CCode).TEXT = XNull(RsComp!Comp_Code)
    Txt(CName).TEXT = XNull(RsComp!Comp_Name)
    Txt(SDate).TEXT = IIf(IsNull(RsComp!Start_Date), "", RsComp!Start_Date)
    Txt(CtrlPath).TEXT = XNull(RsComp!CentralData_Path)
    Label2.CAPTION = Pub_DataPath & "\" & Txt(CtrlPath).TEXT
Else
     Call MnuAdd_Click
End If
Exit Sub
err:
    CheckError
End Sub

