VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmSite 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Site Master"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8595
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
   ScaleHeight     =   4215
   ScaleWidth      =   8595
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   14
      Left            =   3060
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BAD3C9&
      Height          =   2520
      Left            =   825
      TabIndex        =   22
      Top             =   1665
      Width           =   6225
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   13
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   15
         Top             =   2205
         Width           =   1320
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   12
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1965
         Width           =   1320
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   11
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   10
         Text            =   "12345678901234567890"
         Top             =   1485
         Width           =   2880
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   10
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   9
         Text            =   "1234567890"
         Top             =   1245
         Width           =   1545
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2205
         Width           =   2880
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   8
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1965
         Width           =   2880
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   7
         Left            =   1830
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1725
         Width           =   2880
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1005
         Width           =   4230
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   7
         Top             =   765
         Width           =   4230
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   6
         Top             =   525
         Width           =   4230
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   5
         Text            =   "1234567890123456789012345678901234567890"
         Top             =   285
         Width           =   4230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   31
         Top             =   1500
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   30
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.S.T. No && Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   29
         Top             =   2220
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.S.T. No && Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   1980
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   1740
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   26
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   23
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   3060
      MaxLength       =   1
      TabIndex        =   1
      Top             =   600
      Width           =   345
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   6255
      TabIndex        =   19
      Top             =   3870
      Visible         =   0   'False
      Width           =   2010
      Begin MSComctlLib.ListView ListView 
         Height          =   1815
         Left            =   0
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   3201
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   4210752
         BackColor       =   16379351
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   3060
      TabIndex        =   4
      Top             =   1320
      Width           =   1755
   End
   Begin MSDataGridLib.DataGrid DgSite 
      Height          =   2610
      Left            =   585
      Negotiate       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4020
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   4604
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   -2147483639
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Site_Code"
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
         DataField       =   "Site_Desc"
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
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2550.047
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   3060
      MaxLength       =   20
      TabIndex        =   2
      Top             =   840
      Width           =   2685
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   661
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Short Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   1350
      TabIndex        =   32
      Top             =   1095
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1365
      TabIndex        =   18
      Top             =   1335
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   1365
      TabIndex        =   21
      Top             =   615
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1365
      TabIndex        =   16
      Top             =   855
      Width           =   1350
   End
End
Attribute VB_Name = "FrmSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSite As ADODB.Recordset
Dim Master As ADODB.Recordset


Private Const SCode     As Byte = 0
Private Const SDesc     As Byte = 1
Private Const SType     As Byte = 2
Private Const Address1  As Byte = 3
Private Const Address2  As Byte = 4
Private Const Address3  As Byte = 5
Private Const City      As Byte = 6
Private Const Phone     As Byte = 7
Private Const LstNo     As Byte = 8
Private Const CstNo     As Byte = 9
Private Const PinCode   As Byte = 10
Private Const Mobile    As Byte = 11
Private Const LstDate   As Byte = 12
Private Const CstDate   As Byte = 13
Private Const ShortName   As Byte = 14




Dim EditName As String
Dim EditDesc As String
Dim ListArray As Variant
Dim mListItem As ListItem

Private Sub ListView_Click()
On Error GoTo ELoop
    txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    txt(Val(ListView.Tag)).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
    
    TopCtrl1.Tag = PubUParam
    WinSetting Me, 4500, 8715: Ini_Grid
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
      Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    If PubMoveRecYn Then
        Master.Open "select site_Code as searchcode,* from site " & sitecond & " order by site_Code", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 site_Code as searchcode,* from site " & sitecond & " order by site_Code", GCn, adOpenDynamic, adLockOptimistic
    End If
   
    Set RsSite = New ADODB.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select Site_Code as Code, Site_Code,Site_Desc  from site order by site_code", GCn, adOpenDynamic, adLockOptimistic
    Set DgSite.DataSource = RsSite
            
    Disp_Text SETS("INI", Me, Master)
    MoveRec
  Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsSite = Nothing
Set Master = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim VNo As Long
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    txt(SCode).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                GCn.BeginTrans
                GCn.Execute ("delete from site where site_Code= '" & Master!SearchCode & "'")
                GCn.CommitTrans
                Master.Requery
                Call MoveRec
                RsSite.Requery
                BUTTONS True, Me, Master, 0
            End If
eloop1:
    If err.NUMBER <> 0 Then
       GCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
    End If
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    EditName = txt(SCode).TEXT
    EditDesc = txt(SDesc).TEXT
    txt(SCode).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
        Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If

    
    GSQL = "select site_Code as searchcode,Site_Code,Site_Desc, " & cIIF("SiteType='H'", "'HO'", "'Client'") & " as Type from site " & sitecond & " order by site_code"
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
        Master.MoveFirst
        Master.FIND ("searchcode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("select Site_Code as searchcode,* from site Where Site_Code = '" & MyValue & "' order by site_Code")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eFirst()
  BUTTONS True, Me, Master, 1
  Call MoveRec
End Sub

Private Sub TopCtrl1_eLast()
 BUTTONS True, Me, Master, 4
 Call MoveRec
End Sub

Private Sub TopCtrl1_eNext()
 BUTTONS True, Me, Master, 3
 Call MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
 BUTTONS True, Me, Master, 2
 Call MoveRec
End Sub

Private Sub TopCtrl1_eCancel()
Dim I As Integer
On Error GoTo ErrorLoop
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eRef()
    RsSite.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim ItemCode As Integer
    Dim Rst As ADODB.Recordset
    On Error GoTo errlbl
   
     If IsValid(txt(SCode), "Site Code") = False Then Exit Sub
     If IsValid(txt(SDesc), "Site Desc") = False Then Exit Sub
     If IsValid(txt(SType), "Site Type") = False Then Exit Sub

    If TopCtrl1.TopText2 = "Add" Or (TopCtrl1.TopText2 = "Edit" And UCase(txt(SCode).TEXT) <> UCase(EditName)) Then
       Set Rst = New ADODB.Recordset
       Set Rst = GCn.Execute("select site_code from site where site_code = '" & txt(SCode) & "'")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Site Code", vbInformation, "Validation Check": txt(SCode).SetFocus: Exit Sub
            End If
        Set Rst = Nothing
    End If
    If TopCtrl1.TopText2 = "Add" Or (TopCtrl1.TopText2 = "Edit" And UCase(txt(SDesc).TEXT) <> UCase(EditDesc)) Then
       Set Rst = New ADODB.Recordset
       Set Rst = GCn.Execute("select site_desc from site where site_desc = '" & txt(SDesc) & "'")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Site Desc", vbInformation, "Validation Check": txt(SDesc).SetFocus: Exit Sub
            End If
        Set Rst = Nothing
    End If
    
 Grid_Hide
 GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        GCn.Execute ("delete from site where site_Code = '" & txt(SCode) & "'")
        GCn.Execute ("Insert into site(site_Code,site_desc, ShortName,SiteType, Address1, Address2, Address3, City, PinCode, Phone, Mobile, LstNo, LstDate, CstNo, CstDate,U_Name,U_EntDt,U_AE) " & _
            " values('" & txt(SCode) & "' ,'" & txt(SDesc) & "', '" & txt(ShortName) & "','" & IIf(txt(SType) = "HO", "H", "C") & "', '" & txt(Address1) & "',  '" & txt(Address2) & "',  '" & txt(Address3) & "', '" & txt(City) & "', '" & txt(PinCode) & "', '" & txt(Phone) & "', '" & txt(Mobile) & "', '" & txt(LstNo) & "', " & ConvertDate(txt(LstDate)) & ", '" & txt(CstNo) & "', " & ConvertDate(txt(CstDate)) & " ,'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')")
    Else
        GCn.Execute "Update site set Site_desc='" & txt(SDesc) & "', ShortName = '" & txt(ShortName) & "',sitetype = '" & IIf(txt(SType) = "HO", "H", "C") & "', " & _
                    "Address1='" & txt(Address1) & "', Address2='" & txt(Address2) & "', Address3='" & txt(Address3) & "'," & _
                    "City='" & txt(City) & "', PinCode = '" & txt(PinCode) & "', Phone ='" & txt(Phone) & "', Mobile='" & txt(Mobile) & "', " & _
                    "LstNo='" & txt(LstNo) & "', LstDate=" & ConvertDate(txt(LstDate)) & ", CstNo='" & txt(CstNo) & "', CstDate=" & ConvertDate(txt(CstDate)) & ", " & _
                    "U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='" & left(TopCtrl1.TopText2, 1) & "' Where site_code= '" & txt(SCode) & "'"
    End If
GCn.CommitTrans
mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select Site_Code as searchcode,* from site Where Site_Code = '" & txt(SCode) & "' order by site_Code")
    End If
    RsSite.Requery
    Master.FIND "searchcode = '" & txt(SCode) & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans = True Then
        GCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Grid_Hide
    Ctrl_GetFocus txt(Index)
    Select Case Index
    Case SType
        ListArray = Array("HO", "Branch")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Byte
Dim Txtdate As Boolean
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case SType
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case SCode
        DGridTxtKeyDown_Mast DgSite, txt, Index, RsSite, KeyCode, False, 0
End Select
If DgSite.Visible = False And FrmList.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> IIf(UCase(left(PubComp_Name, 4)) = "ENAR", CstDate, SType) Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = IIf(UCase(left(PubComp_Name, 4)) = "ENAR", CstDate, SType) Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If Index <> SCode Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case SCode
        DGridTxtKeyUp_Mast txt, Index, RsSite, KeyCode, "SITE_CODE"
    Case SType
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
    txt(I).Tag = ""
Next I
End Sub

Private Sub MoveRec()
On Error GoTo error1
If Master.RecordCount > 0 Then
    txt(SCode) = Master!SearchCode
    txt(SDesc) = Master!Site_Desc
    If Master!SiteType = "H" Then
        txt(SType).TEXT = "HO"
    Else
        txt(SType).TEXT = "Branch"
    End If
    
    txt(ShortName) = XNull(Master!ShortName)
    txt(Address1) = XNull(Master!Address1)
    txt(Address2) = XNull(Master!Address2)
    txt(Address3) = XNull(Master!Address3)
    txt(City) = XNull(Master!City)
    txt(PinCode) = XNull(Master!PinCode)
    txt(Phone) = XNull(Master!Phone)
    txt(Mobile) = XNull(Master!Mobile)
    txt(PinCode) = XNull(Master!PinCode)
    txt(LstNo) = XNull(Master!LstNo)
    txt(LstDate) = XNull(Master!LstDate)
    txt(CstNo) = XNull(Master!CstNo)
    txt(CstDate) = XNull(Master!CstDate)
End If
TopCtrl1.tPrn = False
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
    txt(I).ForeColor = CtrlFColOrg
Next
    txtDisabled_Color Me
End Sub
Private Sub Grid_Hide()
    If DgSite.Visible = True Then DgSite.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub
Private Sub Ini_Grid()
    DgSite.left = txt(SCode).left: DgSite.top = txt(SCode).top + txt(SCode).height + 20
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error Resume Next
Select Case Index
    Case SType
        If txt(Index).TEXT <> "" Then txt(Index).TEXT = ListView.SelectedItem.TEXT
    Case LstDate, CstDate
        txt(Index) = RetDate(txt(Index))
End Select
End Sub
