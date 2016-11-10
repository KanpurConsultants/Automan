VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmSubventionMast 
   BackColor       =   &H00BAD3C9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subvention Master"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
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
      Left            =   3060
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2280
      Width           =   1605
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
      Left            =   3060
      MaxLength       =   20
      TabIndex        =   6
      Top             =   2040
      Width           =   1605
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
      Left            =   3060
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1800
      Width           =   1605
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
      Left            =   3060
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1560
      Width           =   2685
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
      Left            =   3060
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1320
      Width           =   2685
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
      Left            =   3060
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1080
      Width           =   1605
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
      TabIndex        =   1
      Top             =   840
      Width           =   1605
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
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   6600
      TabIndex        =   8
      Top             =   2475
      Visible         =   0   'False
      Width           =   2010
      Begin MSComctlLib.ListView ListView 
         Height          =   1815
         Left            =   0
         TabIndex        =   9
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
      Index           =   0
      Left            =   3060
      MaxLength       =   20
      TabIndex        =   0
      Top             =   600
      Width           =   2685
   End
   Begin MSDataGridLib.DataGrid DgSubvention 
      Height          =   2100
      Left            =   3960
      Negotiate       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4575
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   3704
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "SchemeNo"
         Caption         =   "Scheme No"
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
            ColumnWidth     =   3314.835
         EndProperty
      EndProperty
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   661
   End
   Begin MSDataGridLib.DataGrid DgModelGroup 
      Height          =   2190
      Left            =   1920
      Negotiate       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4185
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   3863
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Model Group"
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
            ColumnWidth     =   3314.835
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgModel 
      Height          =   2145
      Left            =   -420
      Negotiate       =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4125
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   3784
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Model"
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
            ColumnWidth     =   3314.835
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Subvention"
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
      Left            =   1290
      TabIndex        =   21
      Top             =   2295
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tata Contribution"
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
      Left            =   1290
      TabIndex        =   20
      Top             =   2055
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Contribution"
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
      Left            =   1290
      TabIndex        =   19
      Top             =   1815
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
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
      Left            =   1290
      TabIndex        =   18
      Top             =   1575
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Group"
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
      Left            =   1290
      TabIndex        =   17
      Top             =   1335
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Till Date"
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
      Left            =   1290
      TabIndex        =   16
      Top             =   1095
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Effective From Date"
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
      Left            =   1290
      TabIndex        =   15
      Top             =   855
      Width           =   1695
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
      Left            =   1290
      TabIndex        =   14
      Top             =   2535
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scheme No"
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
      Left            =   1290
      TabIndex        =   13
      Top             =   615
      Width           =   975
   End
End
Attribute VB_Name = "FrmSubventionMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSite As adodb.Recordset
Dim Master As adodb.Recordset
Dim RsModelGroup As adodb.Recordset
Dim RsModel As adodb.Recordset



Private Const SchemeNo As Byte = 0
Private Const FromDate As Byte = 1
Private Const ToDate As Byte = 3
Private Const ModelGroup As Byte = 4
Private Const Model As Byte = 5
Private Const DealerContribution As Byte = 6
Private Const TataContribution As Byte = 7
Private Const TotalSubvention As Byte = 8
Private Const SType As Byte = 2



Dim EditName As String
Dim EditDesc As String
Dim ListArray As Variant
Dim mListItem As ListItem

Private Sub DGModel_Click()
    If RsModel.RecordCount > 0 And RsModel.EOF = False And RsModel.BOF = False Then
        txt(Model) = RsModel!Name
        txt(Model).Tag = RsModel!Code
    End If
    DgModel.Visible = False
End Sub

Private Sub DgModelGroup_Click()
    If RsModelGroup.RecordCount > 0 And RsModelGroup.EOF = False And RsModelGroup.BOF = False Then
        txt(ModelGroup) = RsModelGroup!Name
        txt(ModelGroup).Tag = RsModelGroup!Code
    End If
    DgModelGroup.Visible = False

End Sub

Private Sub ListView_Click()
On Error GoTo Eloop
    txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    txt(Val(ListView.Tag)).SetFocus
Exit Sub
Eloop:
    CheckError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Eloop
FormKeyDown Me, KeyCode, Shift
Exit Sub
Eloop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo Eloop
    
    TopCtrl1.Tag = PubUParam
    WinSetting Me, 4500, 8715: Ini_Grid
    Me.Icon = MDIForm1.Icon
    
    
    Set Master = New adodb.Recordset
    Master.CursorLocation = adUseClient
    
    If PubMoveRecYn Then
        Master.Open "select SchemeNo+ModelGroup as SearchCode,* from Subvention order by SchemeNo", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 SchemeNo+ModelGroup as SearchCode,* from Subvention order by SchemeNo", GCn, adOpenDynamic, adLockOptimistic
    End If
   
    Set RsSite = New adodb.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select SchemeNo as Code, SchemeNo from Subvention order by SchemeNo", GCn, adOpenDynamic, adLockOptimistic
    Set DgSubvention.DataSource = RsSite
            
    Set RsModelGroup = GCn.Execute("Select ModelGrp_Code As Code, ModelGrp_Name As Name From Model_Grp Order By ModelGrp_Name")
    Set DgModelGroup.DataSource = RsModelGroup
        
    Set RsModel = GCn.Execute("Select Model As Code, Model As Name, Grp_Code  From Model Order By Model")
    Set DgModel.DataSource = RsModel
    
                
    Disp_Text SETS("INI", Me, Master)
    MoveRec
  Exit Sub
Eloop:    MsgBox err.Description, vbInformation, "Information"
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
    txt(SchemeNo).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                GCn.BeginTrans
                GCn.Execute ("delete from Subvention where SchemeNo= '" & txt(SchemeNo) & "' and ModelGroup = '" & txt(ModelGroup).Tag & "' ")
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
    EditName = txt(SchemeNo).TEXT
    EditDesc = txt(FromDate).TEXT
    txt(SchemeNo).SetFocus
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
    GSQL = "select SchemeNo+ModelGroup as searchcode,SchemeNo, M.ModelGrp_Name As ModelGroup  from Subvention S Left Join Model_Grp M On S.ModelGroup=M.ModelGrp_Code order by SchemeNo"
    Set SearchForm = Me
    FIND.Show vbModal
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
        Set Master = GCn.Execute("select SchemeNo+ModelGroup as SearchCode,* from Subvention Where SchemeNo+ModelGroup = '" & MyValue & "' order by SchemeNo")
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

Private Sub TopCtrl1_ePrn()
Dim I As Integer, mQry$, mRepName$
Dim Rst As adodb.Recordset
On Error GoTo ERRORHANDLER

    mRepName = "Subvention"
    mQry = "SELECT SchemeNo, FromDate, ToDate, MG.ModelGrp_Name, Model, DealerContribution, " & _
           "TataContribution, TotalSubvention " & _
           "From Subvention S Left Join Model_Grp MG On S.ModelGroup=MG.ModelGrp_Code Order by SchemeNo"

    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
     rpt.Database.SetDataSource Rst
     rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION, , True)
    Set Rst = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub TopCtrl1_eRef()
    RsSite.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim ItemCode As Integer
    Dim Rst As adodb.Recordset
    Dim mSearchCode As String
    On Error GoTo errlbl
   
     If IsValid(txt(SchemeNo), "Scheme No") = False Then Exit Sub
     If IsValid(txt(FromDate), "From Date") = False Then Exit Sub
     If IsValid(txt(ToDate), "ToDate") = False Then Exit Sub
     If IsValid(txt(TotalSubvention), "Total Subvention") = False Then Exit Sub
     

    If TopCtrl1.TopText2 = "Add" Or (TopCtrl1.TopText2 = "Edit" And UCase(txt(SchemeNo).TEXT) <> UCase(EditName)) Then
       Set Rst = New adodb.Recordset
       Set Rst = GCn.Execute("select * from Subvention where SchemeNo = '" & txt(SchemeNo) & "' And ModelGroup='" & txt(ModelGroup).Tag & "'")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate SchemeNo And ModelGroup Combination", vbInformation, "Validation Check": txt(SchemeNo).SetFocus: Exit Sub
            End If
        Set Rst = Nothing
    End If
    mSearchCode = txt(SchemeNo) & txt(ModelGroup).Tag
 Grid_Hide
 GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        GCn.Execute ("delete from Subvention where SchemeNo = '" & txt(SchemeNo) & "' And ModelGroup='" & txt(ModelGroup).Tag & "'")
        GCn.Execute ("insert into Subvention(SchemeNo, FromDate, ToDate, ModelGroup, Model, DealerContribution, TataContribution, TotalSubvention,U_Name,U_EntDt,U_AE) " & _
            " values('" & txt(SchemeNo) & "' ," & ConvertDate(txt(FromDate)) & "," & ConvertDate(txt(ToDate)) & ", '" & txt(ModelGroup).Tag & "', '" & txt(Model).Tag & "', " & Val(txt(DealerContribution)) & ", " & Val(txt(TataContribution)) & ", " & Val(txt(TotalSubvention)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')")
    Else
        GCn.Execute "update Subvention  set FromDate=" & ConvertDate(txt(FromDate)) & ", ToDate=" & ConvertDate(txt(ToDate)) & ", ModelGroup='" & txt(ModelGroup).Tag & "', Model='" & txt(Model).Tag & "', DealerContribution=" & Val(txt(DealerContribution)) & ", TataContribution=" & Val(txt(TataContribution)) & ", TotalSubvention=" & Val(txt(TotalSubvention)) & " ,U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & left(TopCtrl1.TopText2, 1) & "' Where SchemeNo + ModelGroup = '" & Master!SearchCode & "'"
    End If
GCn.CommitTrans
mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select SchemeNo+ModelGroup as SearchCode,* from Subvention Where SchemeNo+ModelGroup = '" & mSearchCode & "' order by SchemeNo")
    End If
    RsSite.Requery
    Master.FIND "searchcode = '" & mSearchCode & "'"
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
        Case ModelGroup
            DgModelGroup.Move txt(Index).left, txt(Index).top + txt(Index).height + 20
            If txt(Index).Tag <> "" And RsModelGroup.RecordCount > 0 Then
                RsModelGroup.MoveFirst
                RsModelGroup.FIND "Code = '" & txt(Index).Tag & "'"
            End If
        Case Model
            RsModel.Filter = adFilterNone
            If txt(ModelGroup).Tag <> "" Then RsModel.Filter = "Grp_Code='" & txt(ModelGroup).Tag & "'"
            DgModel.Move txt(Index).left, txt(Index).top + txt(Index).height + 20
            If txt(Index).Tag <> "" And RsModel.RecordCount > 0 Then
                RsModel.MoveFirst
                RsModel.FIND "Code = '" & txt(Index).Tag & "'"
            End If
            
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
    Case SchemeNo
        DGridTxtKeyDown_Mast DgSubvention, txt, Index, RsSite, KeyCode, False, 0
    Case ModelGroup
        DGridTxtKeyDown DgModelGroup, txt, Index, RsModelGroup, KeyCode, False, 1, frmModelGrp, "frmModelGrp"
    Case Model
        DGridTxtKeyDown DgModel, txt, Index, RsModel, KeyCode, False, 1, frmModel, "frmModel"
        
End Select
If DgSubvention.Visible = False And DgModel.Visible = False And DgModelGroup.Visible = False And FrmList.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> TataContribution Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = TataContribution Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If Index <> SchemeNo Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
    Select Case Index
        Case ModelGroup
            DGridTxtKeyPress txt, Index, RsModelGroup, KeyAscii, "Name"
        Case Model
            DGridTxtKeyPress txt, Index, RsModel, KeyAscii, "Name"
    End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case DealerContribution, TataContribution
        Amt_Cal
    Case SchemeNo
        DGridTxtKeyUp_Mast txt, Index, RsSite, KeyCode, "SchemeNo"
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
    txt(SchemeNo).Tag = Master!SearchCode
    txt(SchemeNo) = Master!SchemeNo
    txt(FromDate) = XNull(Master!FromDate)
    txt(ToDate) = XNull(Master!ToDate)
    txt(ModelGroup).Tag = XNull(Master!ModelGroup)
    If txt(ModelGroup).Tag <> "" Then
        txt(ModelGroup) = GCn.Execute("Select ModelGrp_Name From Model_Grp Where ModelGrp_Code = '" & txt(ModelGroup).Tag & "'").Fields(0)
    Else
        txt(ModelGroup) = ""
    End If
    txt(Model).Tag = XNull(Master!Model)
    If txt(Model).Tag <> "" Then
        txt(Model) = GCn.Execute("Select Model From Model Where Model = '" & txt(Model).Tag & "'").Fields(0)
    Else
        txt(Model) = ""
    End If
    txt(DealerContribution) = Format(VNull(Master!DealerContribution), "0.00")
    txt(TataContribution) = Format(VNull(Master!TataContribution), "0.00")
    txt(TotalSubvention) = Format(VNull(Master!TotalSubvention), "0.00")
End If


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
    If DgSubvention.Visible = True Then DgSubvention.Visible = False
    If DgModelGroup.Visible = True Then DgModelGroup.Visible = False
    If DgModel.Visible = True Then DgModel.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub
Private Sub Ini_Grid()
    DgSubvention.left = txt(SchemeNo).left: DgSubvention.top = txt(SchemeNo).top + txt(SchemeNo).height + 20
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error Resume Next
Select Case Index
    Case ModelGroup
        If RsModelGroup.RecordCount > 0 And RsModel.EOF = False And RsModel.BOF = False And txt(Index) <> "" Then
            txt(Index) = RsModelGroup!Name
            txt(Index).Tag = RsModelGroup!Code
        Else
            txt(Index) = ""
            txt(Index).Tag = ""
        End If
    Case Model
        If RsModel.RecordCount > 0 And RsModel.EOF = False And RsModel.BOF = False And txt(Index) <> "" Then
            txt(Index) = RsModel!Name
            txt(Index).Tag = RsModel!Code
        Else
            txt(Index) = ""
            txt(Index).Tag = ""
        End If
        
    Case DealerContribution, TataContribution
        txt(Index) = Format(txt(Index), "0.00")
        Amt_Cal
    Case FromDate, ToDate
        txt(Index) = RetDate(txt(Index))
    Case SType
            If txt(Index).TEXT <> "" Then txt(Index).TEXT = ListView.SelectedItem.TEXT
End Select
End Sub


Sub Amt_Cal()
    txt(TotalSubvention) = Format(Val(txt(DealerContribution)) + Val(txt(TataContribution)), "0.00")
End Sub

