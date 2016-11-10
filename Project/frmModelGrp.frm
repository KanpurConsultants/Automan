VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmModelGrp 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Vehicle Model Group Master"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10320
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   7890
      TabIndex        =   22
      Top             =   345
      Visible         =   0   'False
      Width           =   2505
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   0
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   0
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3228
         View            =   3
         Arrange         =   1
         Sorted          =   -1  'True
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
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   661
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2355
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1275
      Width           =   1440
   End
   Begin VB.Frame FrModelCat 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   6030
      TabIndex        =   14
      Top             =   3570
      Visible         =   0   'False
      Width           =   4095
      Begin MSDataGridLib.DataGrid DGModelCat 
         Height          =   3225
         Left            =   30
         TabIndex        =   15
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
            DataField       =   "ModelCat_Code"
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
            DataField       =   "ModelCat_Name"
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
            DataField       =   "ModelCat_Code"
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
         Caption         =   "List of Model Category"
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
         TabIndex        =   16
         Top             =   30
         Width           =   4050
      End
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2355
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1515
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   6630
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1515
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Frame frDivision 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   6000
      TabIndex        =   18
      Top             =   2325
      Visible         =   0   'False
      Width           =   4095
      Begin MSDataGridLib.DataGrid DGDivision 
         Height          =   3225
         Left            =   30
         TabIndex        =   19
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
            DataField       =   "Div_Code"
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
            DataField       =   "Div_Name"
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
            DataField       =   "Div_Code"
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
         Caption         =   "List of Division"
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
         Index           =   2
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   4050
      End
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   6630
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1035
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2355
      MaxLength       =   5
      TabIndex        =   1
      Top             =   555
      Width           =   1260
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2355
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1035
      Width           =   4245
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2355
      MaxLength       =   40
      TabIndex        =   2
      Top             =   795
      Width           =   4245
   End
   Begin VB.Frame FrModelGrp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   660
      TabIndex        =   9
      Top             =   2835
      Visible         =   0   'False
      Width           =   4980
      Begin MSDataGridLib.DataGrid DGModelGrp 
         Height          =   3225
         Left            =   30
         TabIndex        =   17
         Top             =   345
         Width           =   4920
         _ExtentX        =   8678
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
            DataField       =   "ModelGrp_Code"
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
            DataField       =   "ModelGrp_Name"
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
            DataField       =   "ModelGrp_Code"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   3435.024
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
         Caption         =   "List of Model Group"
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
         TabIndex        =   10
         Top             =   30
         Width           =   4935
      End
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Division*"
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
      Left            =   270
      TabIndex        =   21
      Top             =   1515
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code*"
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
      Left            =   270
      TabIndex        =   13
      Top             =   555
      Width           =   555
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wheel Category*"
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
      Left            =   270
      TabIndex        =   12
      Top             =   1275
      Width           =   1485
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Category*"
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
      Left            =   270
      TabIndex        =   11
      Top             =   1035
      Width           =   1455
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Group Name*"
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
      Left            =   270
      TabIndex        =   8
      Top             =   795
      Width           =   1740
   End
End
Attribute VB_Name = "frmModelGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Don't Change Tag Property of (Txt) Control as it is used in other activities
'FORM COLOR &H00C0FFFF&
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstModelCat As ADODB.Recordset
Dim RstDivision As ADODB.Recordset, mFlag As Byte, mFLAG1 As Byte
Private Const ModelGrp_Code = 0, ModelGrp_Name = 1, ModelCat_Code = 2, ModelCat_NAME = 3
Private Const Wheel_Catg = 4 ',ModelDiv_NAME = 6, ModelDiv_Code = 5
Dim ListArray As Variant
Dim mListItem As ListItem
         
Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()

TopCtrl1.Tag = PubUParam: TopCtrl1.TopText1 = Me.CAPTION ' "Vehicle Model Group Master"
Set RstMain = New ADODB.Recordset


 Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and LEFT(MODEL_GRP.site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
If PubMoveRecYn Then
    RstMain.Open "Select MODEL_GRP.*,Model_Grp.ModelGrp_Code as SearchCode,MODEL_CAT.ModelCat_Name,DIVISION.Div_Name " & _
        " From (MODEL_GRP Left Join MODEL_CAT On MODEL_CAT.ModelCat_Code=MODEL_GRP.ModelCat_Code) " & _
        " LEFT JOIN DIVISION ON Left(MODEL_GRP.ModelGrp_Code,1)=DIVISION.Div_Code " & _
        " where left(ModelGrp_Code,1)='" & PubDivCode & "' " & sitecond & "Order by MODEL_GRP.ModelGrp_Name", GCn, adOpenDynamic, adLockOptimistic
Else
    RstMain.Open "Select Top 1 MODEL_GRP.*,Model_Grp.ModelGrp_Code as SearchCode,MODEL_CAT.ModelCat_Name,DIVISION.Div_Name " & _
        " From (MODEL_GRP Left Join MODEL_CAT On MODEL_CAT.ModelCat_Code=MODEL_GRP.ModelCat_Code) " & _
        " LEFT JOIN DIVISION ON Left(MODEL_GRP.ModelGrp_Code,1)=DIVISION.Div_Code " & _
        " where left(ModelGrp_Code,1)='" & PubDivCode & "' " & sitecond & " Order by MODEL_GRP.ModelGrp_Name", GCn, adOpenDynamic, adLockOptimistic
End If

Set RstHelp = New ADODB.Recordset
RstHelp.Open "Select ModelGrp_Code,ModelGrp_Name FROM MODEL_GRP where Left(ModelGrp_Code,1)='" & PubDivCode & "' Order by ModelGrp_Name", GCn, adOpenDynamic, adLockOptimistic
Set DGModelGrp.DataSource = RstHelp

Set RstModelCat = New ADODB.Recordset
RstModelCat.Open "Select ModelCat_Code,ModelCat_Name FROM MODEL_CAT where Left(ModelCat_Code,1)='" & PubDivCode & "' Order by ModelCat_NAME", GCn, adOpenDynamic, adLockOptimistic
Set DGModelCat.DataSource = RstModelCat

'Set RstDivision = New ADODB.Recordset
'RstDivision.Open "Select DIV_CODE,Div_Name,DIV_CODE FROM DIVISION Order by Div_Name", GCn, adOpenDynamic, adLockOptimistic
'Set DGDivision.DataSource = RstDivision
FrModelGrp.Visible = False
FrModelCat.Visible = False

frDivision.Visible = False
ListArray = Array("Two", "Three", "Four", "Six", "Ten", "Above Ten")
Set mListItem = ListView_Items(ListView, Txt, Wheel_Catg, ListArray, 6)
CtrlClckCol
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0
mFlag = 0
WinSetting Me, 5835, 9645
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing: Set RstModelCat = Nothing: Set RstDivision = Nothing
End Sub

Private Sub CtrlClckCol()
    Txt(ModelGrp_Code).BackColor = CtrlBColOrg:     Txt(ModelGrp_Code).ForeColor = CtrlFColOrg
    Txt(ModelGrp_Name).BackColor = CtrlBColOrg:     Txt(ModelGrp_Name).ForeColor = CtrlFColOrg
    Txt(ModelCat_NAME).BackColor = CtrlBColOrg:     Txt(ModelCat_NAME).ForeColor = CtrlFColOrg
    Txt(ModelCat_Code).BackColor = CtrlBColOrg:     Txt(ModelCat_Code).ForeColor = CtrlFColOrg
'    Txt(ModelDiv_Code).BackColor = CtrlBColOrg:     Txt(ModelDiv_Code).ForeColor = CtrlFColOrg
    Txt(Wheel_Catg).BackColor = CtrlBColOrg:        Txt(Wheel_Catg).ForeColor = CtrlFColOrg
'    Txt(ModelDiv_NAME).BackColor = CtrlBColOrg:     Txt(ModelDiv_NAME).ForeColor = CtrlFColOrg
End Sub

Private Sub Disp_Text(Enb As Boolean)
    Txt(ModelGrp_Code).Enabled = Enb
    Txt(ModelGrp_Name).Enabled = Enb
    Txt(ModelCat_NAME).Enabled = Enb
    Txt(ModelCat_Code).Enabled = Enb
'    Txt(ModelDiv_NAME).Enabled = Enb
'    Txt(ModelDiv_Code).Enabled = Enb
    Txt(Wheel_Catg).Enabled = Enb
End Sub

Private Sub MakeBlank()
    Txt(ModelGrp_Code) = ""
    Txt(ModelGrp_Name) = ""
    Txt(ModelCat_NAME) = ""
    Txt(ModelCat_Code) = ""
'    Txt(ModelDiv_NAME) = ""
'    Txt(ModelDiv_Code) = ""
    Txt(Wheel_Catg) = ""
End Sub

Private Sub MoveRec()
On Error GoTo ErrLoop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    MakeBlank
Else
    Txt(ModelGrp_Code) = XNull(RstMain!ModelGrp_Code)
    Txt(ModelGrp_Name) = XNull(RstMain!ModelGrp_Name)
    Txt(ModelCat_NAME) = XNull(RstMain!ModelCat_NAME)
    Txt(ModelCat_Code) = XNull(RstMain!ModelCat_Code)
'    Txt(ModelDiv_NAME) = XNull(RstMain!Div_Name)
'    Txt(ModelDiv_Code) = XNull(RstMain!ModelDiv_Code)
    Txt(Wheel_Catg) = XNull(RstMain!Wheel_Catg)
End If
TopCtrl1.tDel = False
'TopCtrl1.tPrn = False
Grid_Hide
Exit Sub
ErrLoop:        MsgBox err.Description
End Sub

Private Sub ListView_Click()
Txt(Wheel_Catg).TEXT = ListView.SelectedItem.TEXT
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrLoop
MakeBlank
If ADDFLAG <> 1 Then Disp_Text SETS("ADD", Me, RstMain)
ADDFLAG = 1
'Txt(ModelGrp_Code).Tag = Txt(ModelGrp_Code)
Txt_GotFocus ModelGrp_Code
Txt(ModelGrp_Code) = PubDivCode
Txt(ModelGrp_Code).SelStart = Len(Txt(ModelGrp_Code))
Txt(ModelGrp_Code).SetFocus
Exit Sub
ErrLoop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ErrLoop
If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(ModelGrp_Code).Enabled = False
    Txt(ModelGrp_Name).Tag = Txt(ModelGrp_Name)
    Txt_GotFocus ModelGrp_Name
    Txt(ModelGrp_Name).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
ErrLoop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub

'Private Sub TopCtrl1_eDel()
'On Error GoTo Errloop
'Dim transFalg As Byte
'transFalg = 0
'If RstMain.RecordCount  > 0 Then
'    If MsgBox("Are You Sure to Delete This Record", vbYesNo, "Confirmation") = vbYes Then
'        GCn.BeginTrans
'        transFalg = 1
'        GCn.Execute ("Delete From MODEL_GRP Where ModelGrp_Code=" & Chk_Text(Trim(Txt(ModelGrp_Code))))
'        GCn.CommitTrans
'        transFalg = 0
'        RstMain.Requery
'        RstHelp.Requery
'        Disp_Text SETS("INI", Me, RstMain)
'        MoveRec
'    End If
'Else
'    MsgBox "There Is No Record To Delete.", vbInformation, "Information"
'End If
'Exit Sub
'ErrLoop:    If transFalg = 1 Then GCn.RollbackTrans
'            MsgBox err.Description, vbExclamation, " Deletion Error "
'End Sub

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
      sitecond = "and LEFT(MODEL_GRP.site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    
    GSQL = "Select Model_Grp.ModelGrp_Code as SearchCode, Model_Grp.ModelGrp_Code, Model_Grp.ModelGrp_Name, MODEL_CAT.ModelCat_Name " & _
    "From MODEL_GRP Left Join MODEL_CAT On MODEL_CAT.ModelCat_Code=MODEL_GRP.ModelCat_Code " & _
    " where Left(MODEL_GRP.ModelGrp_Code,1)='" & PubDivCode & "' " & sitecond & " Order by MODEL_GRP.ModelGrp_Name"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim rep As CrystalReport, Form1 As frmMastList
    Set Form1 = New frmMastList
    With Form1
        .g_FormID = 4
        .LblName.CAPTION = Me.CAPTION
        .CAPTION = Me.CAPTION
        .Show
    End With
    Set Form1 = Nothing
    Set rep = Nothing

End Sub

Private Sub TopCtrl1_eSave()
Dim transFlag As Byte
On Error GoTo ErrLoop
    transFlag = 0
    If Len(Trim(Txt(ModelGrp_Code))) = 1 Then MsgBox "Group Code should be filled ", vbOKOnly, "Validation": Txt(ModelGrp_Code).SetFocus: Exit Sub
    If IsValid(Txt(ModelGrp_Name), "Group Name") = False Then Txt_GotFocus ModelGrp_Name: Exit Sub
    If IsValid(Txt(ModelCat_NAME), "Model Category") = False Then Txt_GotFocus ModelCat_NAME: Exit Sub
'    If IsValid(Txt(ModelDiv_NAME), "Model Division") = False Then Txt_GotFocus ModelDiv_NAME: Exit Sub
    If IsValid(Txt(Wheel_Catg), "Wheel Category") = False Then Txt_GotFocus Wheel_Catg: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From MODEL_GRP Where ModelGrp_Code=" & Chk_Text(Trim(Txt(ModelGrp_Code)))).Fields(0) > 0 Then MsgBox "Code Already Exists", vbInformation, "Code Validation": Txt_GotFocus ModelGrp_Code: Txt(ModelGrp_Code).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If TopCtrl1.TopText2 = "Add" Then
        GCn.Execute ("Insert Into MODEL_GRP (ModelGrp_Code,Site_Code,ModelGrp_Name,ModelCat_Code,Wheel_Catg,U_Name,U_EntDt,U_AE) Values('" & Trim(Txt(ModelGrp_Code)) & "','" & PubSiteCode & "'," & Chk_Text(Txt(ModelGrp_Name)) & "," & Chk_Text(Txt(ModelCat_Code)) & "," & Chk_Text(Txt(Wheel_Catg)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
    Else
        GCn.Execute ("update MODEL_GRP set Site_Code='" & PubSiteCode & "',ModelGrp_Name=" & Chk_Text(Txt(ModelGrp_Name)) & ",ModelCat_Code=" & Chk_Text(Txt(ModelCat_Code)) & ",Wheel_Catg=" & Chk_Text(Txt(Wheel_Catg)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' Where ModelGrp_Code=" & Chk_Text(Trim(Txt(ModelGrp_Code))))
        GCn.Execute ("Update Model set Div_Code='" & left(Txt(ModelGrp_Code), 1) & "', U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' where Grp_Code='" & Trim(Txt(ModelGrp_Code)) & "'")
    End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select MODEL_GRP.*, Model_Grp.ModelGrp_Code as SearchCode,MODEL_CAT.ModelCat_Name,DIVISION.Div_Name " & _
            " From (MODEL_GRP Left Join MODEL_CAT On MODEL_CAT.ModelCat_Code=MODEL_GRP.ModelCat_Code) " & _
            " LEFT JOIN DIVISION ON Left(MODEL_GRP.ModelGrp_Code,1)=DIVISION.Div_Code " & _
            " where left(ModelGrp_Code,1)='" & PubDivCode & "' And Model_Grp.ModelGrp_Code = " & Chk_Text(Trim(Txt(ModelGrp_Code))) & " Order by MODEL_GRP.ModelGrp_Name")
    End If
    RstHelp.Requery
    RstMain.FIND ("ModelGrp_Code=" & Chk_Text(Trim(Txt(ModelGrp_Code))))
    If ADDFLAG = 1 Then
        TopCtrl1_eAdd
'        MakeBlank
'        Txt_GotFocus ModelGrp_Code
'        Txt(ModelGrp_Code).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        frDivision.Visible = False
        frmModelCat.Visible = False
        FrModelGrp.Visible = False
        FrmList.Visible = False
    End If
Exit Sub
ErrLoop:    If transFlag = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo ErrLoop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
    If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        frDivision.Visible = False
        frmModelCat.Visible = False
        FrModelGrp.Visible = False
        FrmList.Visible = False
    End If
Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
    RstModelCat.Requery
'    RstDivision.Requery
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub ModelGrp_CodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "ModelGrp_Code >=" & Chk_Text(XNull(Trim(Txt(ModelGrp_Code))))
End Sub
Private Sub ModelGrp_NameSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "ModelGrp_Name >=" & Chk_Text(XNull(Txt(ModelGrp_Name)))
End Sub
Private Sub ModelCat_NAMESearch()
If RstModelCat.RecordCount <= 0 Then Exit Sub
RstModelCat.MoveFirst
RstModelCat.FIND "ModelCat_NAME >=" & Chk_Text(XNull(Txt(ModelCat_NAME)))
End Sub
'Private Sub ModelDiv_NAMESearch()
'If RstDivision.RecordCount <= 0 Then Exit Sub
'RstDivision.MoveFirst
'RstDivision.FIND "Div_Name >=" & Chk_Text(XNull(Txt(ModelDiv_NAME)))
'End Sub
Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case ModelGrp_Code, ModelGrp_Name
            If FrModelGrp.Visible = True Then FrModelGrp.Visible = False
            If frDivision.Visible = True Then frDivision.Visible = False
            If FrmList.Visible = True Then FrmList.Visible = False
            If FrModelCat.Visible = True Then FrModelCat.Visible = False
            FrModelGrp.Visible = True
            FrModelGrp.top = Txt(Index).top + Txt(Index).height + 10
            FrModelGrp.left = Txt(Index).left
            FrModelGrp.ZOrder 0
        Case ModelCat_NAME
            If FrmList.Visible = True Then FrmList.Visible = False
            If frDivision.Visible = True Then frDivision.Visible = False
            If FrModelGrp.Visible = True Then FrModelGrp.Visible = False
            FrModelCat.Visible = True
            FrModelCat.top = Txt(Index).top + Txt(Index).height + 10
            FrModelCat.left = Txt(Index).left
            FrModelCat.ZOrder 0
'        Case ModelDiv_NAME
'            If FrmList.Visible = True Then FrmList.Visible = False
'            If FrModelCat.Visible = True Then FrModelCat.Visible = False
'            If FrModelGrp.Visible = True Then FrModelGrp.Visible = False
'            frDivision.Visible = True
'            frDivision.top = Txt(Index).top + Txt(Index).Height + 10
'            frDivision.left = Txt(Index).left
'            frDivision.ZOrder 0
        Case Wheel_Catg
            If FrModelCat.Visible = True Then FrModelCat.Visible = False
            If FrModelGrp.Visible = True Then FrModelGrp.Visible = False
            If frDivision.Visible = True Then frDivision.Visible = False
    End Select
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
mFlag = 0
mFLAG1 = 0
Grid_Hide
    Select Case Index
        Case ModelGrp_Code, ModelGrp_Name
            RST_BOF_EOF RstHelp
        Case ModelCat_NAME
            RST_BOF_EOF RstModelCat
'        Case ModelDiv_NAME
'            RST_BOF_EOF RstDivision
    End Select
    Txt(Index).Tag = Txt(Index)
    Txt_Click Index
    Select Case Index
        Case ModelGrp_Code, ModelGrp_Name
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
        Case ModelCat_NAME
            If RstModelCat.BOF Or RstModelCat.EOF Then Exit Sub
'        Case ModelDiv_NAME
'            If RstDivision.BOF Or RstDivision.EOF Then Exit Sub
    End Select
    DGModelCat.Columns(0).width = 1000.1: DGModelCat.Columns(1).width = 3435.024: DGModelCat.Columns(2).width = 1000.1
    DGModelGrp.Columns(0).width = 1000.1: DGModelGrp.Columns(1).width = 3435.024: DGModelGrp.Columns(2).width = 1000.1
    DGDivision.Columns(0).width = 1000.1: DGDivision.Columns(1).width = 3435.024: DGDivision.Columns(2).width = 1000.1
    Select Case Index
        Case ModelGrp_Code
            DGModelGrp.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "ModelGrp_Code ASC"
            RstHelp.Bookmark = mBookMark
            ModelGrp_CodeSearch
        Case ModelGrp_Name
            DGModelGrp.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "ModelGrp_Name ASC"
            RstHelp.Bookmark = mBookMark
            ModelGrp_NameSearch
        Case ModelCat_NAME
            DGModelCat.Columns(0).width = 0: DGModelCat.Columns(2).width = 0
            mBookMark = RstModelCat.Bookmark
            RstModelCat.Sort = "ModelCat_Name ASC"
            RstModelCat.Bookmark = mBookMark
            ModelCat_NAMESearch
'        Case ModelDiv_NAME
'            DGDivision.Columns(0).width = 0: DGDivision.Columns(2).width = 0
'            mBookMark = RstDivision.Bookmark
'            RstDivision.Sort = "Div_Name ASC"
'            RstDivision.Bookmark = mBookMark
'            ModelDiv_NAMESearch
    End Select
    If Txt(Index) = "" Then Txt_Change Index
End Sub
Private Sub Txt_Click(Index As Integer)
    CtrlClckCol
    Txt(Index).ForeColor = CtrlFCol: Txt(Index).BackColor = CtrlBCol
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case ModelCat_NAME
        If FrModelCat.Visible = True Then
            Select Case KeyCode
                Case vbKeyUp
                    If Not RstModelCat.BOF Then RstModelCat.MovePrevious
                Case vbKeyDown
                    If Not RstModelCat.EOF Then RstModelCat.MoveNext
                Case 33
                    For I = 1 To 9
                        If Not RstModelCat.BOF Then RstModelCat.MovePrevious
                    Next
                Case 34
                    For I = 1 To 9
                        If Not RstModelCat.EOF Then RstModelCat.MoveNext
                    Next
                Case 13
                    SendKeysA vbKeyTab, True
            End Select
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, 33, 34
                    RST_BOF_EOF RstModelCat
                    If Not RstModelCat.BOF And Not RstModelCat.EOF Then
                        Txt(ModelCat_Code) = XNull(RstModelCat!ModelCat_Code): Txt(ModelCat_NAME) = XNull(RstModelCat!ModelCat_NAME)
                        Txt(ModelCat_NAME).SelStart = 0
                    End If
            End Select
        End If
'    Case ModelDiv_NAME
'        If frDivision.Visible = True Then
'            Select Case KeyCode
'                Case vbKeyUp
'                    If Not RstDivision.BOF Then RstDivision.MovePrevious
'                Case vbKeyDown
'                    If Not RstDivision.EOF Then RstDivision.MoveNext
'                Case 33
'                    For i = 1 To 9
'                        If Not RstDivision.BOF Then RstDivision.MovePrevious
'                    Next
'                Case 34
'                    For i = 1 To 9
'                        If Not RstDivision.EOF Then RstDivision.MoveNext
'                    Next
'                Case 13
'                    SendKeysA vbKeyTab, True
'            End Select
'            Select Case KeyCode
'                Case vbKeyUp, vbKeyDown, 33, 34
'                    RST_BOF_EOF RstDivision
'                    If Not RstDivision.BOF And Not RstDivision.EOF Then
'                        Txt(ModelDiv_Code) = XNull(RstDivision!Div_Code): Txt(ModelDiv_NAME) = XNull(RstDivision!Div_Name)
'                        Txt(ModelDiv_NAME).SelStart = 0
'                    End If
'            End Select
'        End If
    Case Wheel_Catg
            If KeyCode <> vbKeyEscape Then
                ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1600
            End If

End Select
Select Case Index
    Case ModelGrp_Code
        'Div Code Edit restricted
        KeyCode = RestrictCode(KeyCode, Txt(Index), Shift)
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case ModelGrp_Name
            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
                SendKeysA vbKeyTab, True
                KeyCode = 0
            ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
                SendKeys "+{Tab}"
                KeyCode = 0
            End If
    Case ModelCat_NAME
        If FrModelCat.Visible = False Then
            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
                SendKeysA vbKeyTab, True
                KeyCode = 0
            ElseIf KeyCode = vbKeyUp Then
                SendKeys "+{Tab}"
                KeyCode = 0
            End If
        End If
'    Case ModelDiv_NAME
'        If frDivision.Visible = False Then
'            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
'                SendKeysA vbKeyTab, True
'                KeyCode = 0
'            ElseIf KeyCode = vbKeyUp Then
'                SendKeys "+{Tab}"
'                KeyCode = 0
'            End If
'        End If
    Case Wheel_Catg
        If FrmList.Visible = False Then
WHEELLBL:
            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
                If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                    Txt_Validate Wheel_Catg, result
                    If result = True Then Txt_GotFocus Wheel_Catg: Txt(Wheel_Catg).SetFocus: Exit Sub
                    TopCtrl1_eSave
                Else
                    Txt_Click Wheel_Catg
                    Txt_GotFocus Wheel_Catg
                    Txt(Wheel_Catg).SetFocus
                End If
            ElseIf KeyCode = vbKeyUp Then
                SendKeys "+{Tab}"
                KeyCode = 0
            End If
        End If
End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case ModelGrp_Code
        KeyAscii = RestrictCode(KeyAscii, Txt(Index), 0)
End Select

End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
mFlag = 0
mFLAG1 = 0
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case Wheel_Catg
        ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
    Case ModelGrp_Code
        ModelGrp_CodeSearch
    Case ModelGrp_Name
        ModelGrp_NameSearch
    Case ModelCat_NAME
        ModelCat_NAMESearch
'    Case ModelDiv_NAME
'        ModelDiv_NAMESearch
End Select
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case ModelGrp_Code
            Set Rst = GCn.Execute("SELECT * FROM MODEL_GRP WHERE ModelGrp_Code=" & Chk_Text(Txt(ModelGrp_Code)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Code Already Exists", vbInformation, "Validation": Txt(ModelGrp_Code) = Txt(ModelGrp_Code).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!ModelGrp_Code <> RstMain!ModelGrp_Code Then MsgBox "Code Already Exists", vbInformation, "Validation": Txt(ModelGrp_Code) = Txt(ModelGrp_Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case ModelGrp_Name
            Set Rst = GCn.Execute("SELECT * FROM MODEL_GRP WHERE ModelGrp_Name=" & Chk_Text(Txt(ModelGrp_Name)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Name Already Exists", vbInformation, "Validation": Txt(ModelGrp_Name) = Txt(ModelGrp_Name).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!ModelGrp_Name <> RstMain!ModelGrp_Name Then MsgBox "Name Already Exists", vbInformation, "Validation": Txt(ModelGrp_Name) = Txt(ModelGrp_Name).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case ModelCat_NAME
            If Not RstModelCat.EOF And Not RstModelCat.BOF Then
                Txt(ModelCat_Code) = XNull(RstModelCat!ModelCat_Code): Txt(ModelCat_NAME) = XNull(RstModelCat!ModelCat_NAME)
            Else
                Txt(ModelCat_Code) = "": Txt(ModelCat_NAME) = ""
            End If
'        Case ModelDiv_NAME
'            If Not RstDivision.EOF And Not RstDivision.BOF Then
'                Txt(ModelDiv_Code) = XNull(RstDivision!Div_Code): Txt(ModelDiv_NAME) = XNull(RstDivision!Div_Name)
'            Else
'                Txt(ModelDiv_Code) = "": Txt(ModelDiv_NAME) = ""
'            End If
        Case Wheel_Catg
            Txt(Wheel_Catg).TEXT = ListView.SelectedItem.TEXT
            Grid_Hide
    End Select
Set Rst = Nothing
End Sub
Private Sub DGModelCat_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mFlag = 1 Then
    Txt(ModelCat_Code) = DGModelCat.Columns(0).TEXT: Txt(ModelCat_NAME) = DGModelCat.Columns(1).TEXT
End If
End Sub
Private Sub DGModelCat_GotFocus()
    mFlag = 1
End Sub
'Private Sub DGDivision_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If mFLAG1 = 1 Then
'    Txt(ModelDiv_Code) = DGDivision.Columns(0).Text: Txt(ModelDiv_NAME) = DGDivision.Columns(1).Text
'End If
'End Sub
'Private Sub DGDivision_GotFocus()
'    mFLAG1 = 1
'End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select MODEL_GRP.*, Model_Grp.ModelGrp_Code as SearchCode,MODEL_CAT.ModelCat_Name,DIVISION.Div_Name " & _
            " From (MODEL_GRP Left Join MODEL_CAT On MODEL_CAT.ModelCat_Code=MODEL_GRP.ModelCat_Code) " & _
            " LEFT JOIN DIVISION ON Left(MODEL_GRP.ModelGrp_Code,1)=DIVISION.Div_Code " & _
            " where left(ModelGrp_Code,1)='" & PubDivCode & "' And Model_Grp.ModelGrp_Code = '" & MyValue & "' Order by MODEL_GRP.ModelGrp_Name")
    End If
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Grid_Hide()
    If FrmList.Visible = True Then FrmList.Visible = False
        If FrModelCat.Visible = True Then FrModelCat.Visible = False
    If FrModelGrp.Visible = True Then FrModelGrp.Visible = False
    If frDivision.Visible = True Then frDivision.Visible = False

End Sub

