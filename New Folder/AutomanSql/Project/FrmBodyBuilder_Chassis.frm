VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmBodyBuilder_Chassis 
   Caption         =   "Chassis Allotment to Body Builder"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   7830
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   7
      Left            =   2415
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1785
      Width           =   4200
   End
   Begin MSDataGridLib.DataGrid DgBodyType 
      Height          =   2040
      Left            =   1860
      Negotiate       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2790
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   3598
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
         Caption         =   "Body Type"
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
            ColumnWidth     =   3479.811
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   735
      ScaleHeight     =   255
      ScaleWidth      =   270
      TabIndex        =   21
      Top             =   30
      Width           =   270
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   30
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   20
      Top             =   30
      Width           =   315
   End
   Begin MSDataGridLib.DataGrid DgChassis 
      Height          =   2040
      Left            =   705
      Negotiate       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   3598
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
         Caption         =   "Body Builder Name"
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
            ColumnWidth     =   3479.811
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgBodyBuilder 
      Height          =   2040
      Left            =   -1410
      Negotiate       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3885
      Visible         =   0   'False
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   3598
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
         Caption         =   "Chassis No"
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
            ColumnWidth     =   2610.142
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   6
      Left            =   2415
      MaxLength       =   40
      TabIndex        =   18
      Top             =   1275
      Width           =   4200
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   2415
      MaxLength       =   100
      TabIndex        =   3
      Top             =   2040
      Width           =   1470
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   5070
      MaxLength       =   100
      TabIndex        =   4
      Top             =   2040
      Width           =   1545
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   2415
      MaxLength       =   1
      TabIndex        =   9
      Top             =   765
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   6600
      TabIndex        =   7
      Top             =   2475
      Visible         =   0   'False
      Width           =   2010
      Begin MSComctlLib.ListView ListView 
         Height          =   1815
         Left            =   150
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   330
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   2415
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1020
      Width           =   2505
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   4
      Left            =   2415
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1530
      Width           =   4200
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   5
      Left            =   2415
      MaxLength       =   100
      TabIndex        =   6
      Top             =   2295
      Width           =   4200
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   384
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7836
      _ExtentX        =   13811
      _ExtentY        =   688
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body Type....."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   1230
      TabIndex        =   24
      Top             =   1815
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body Builder....."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   0
      TabIndex        =   23
      Top             =   30
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model...................."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   1230
      TabIndex        =   19
      Top             =   1305
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date......"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1215
      TabIndex        =   17
      Top             =   2070
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body Builder....."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1215
      TabIndex        =   16
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code.................."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   1215
      TabIndex        =   15
      Top             =   795
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No. ..................."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1215
      TabIndex        =   14
      Top             =   1050
      Width           =   2205
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rect. Date........."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   3975
      TabIndex        =   13
      Top             =   2055
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remark...................."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   1200
      TabIndex        =   12
      Top             =   2325
      Width           =   1875
   End
End
Attribute VB_Name = "FrmBodyBuilder_Chassis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsChassis As ADODB.Recordset
Dim RsBodyType As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim RsBodyBuilder As ADODB.Recordset

Private Const tCode         As Byte = 0
Private Const Chassis         As Byte = 1
Private Const RecDate          As Byte = 2
Private Const IssDate          As Byte = 3
Private Const BodyBuilder          As Byte = 4
Private Const Remark       As Byte = 5
Private Const Model       As Byte = 6
Private Const BodyType As Byte = 7



Dim EditName        As String
Dim EditDesc        As String
Dim ListArray       As Variant
Dim mListItem       As ListItem


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
    WinSetting Me, 4500, 8715
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "Select ChassisNo as SearchCode, ChassisNo As Chassis_No, InDate, Model, " & _
                    "BodyBuilder, BodyBuilder_BodyType, BodyBuilder_IssDate, BodyBuilder_RecDate, " & _
                    "BB.BodyBuilderDesc, VS.BodyBuilder_Remark As Remark, Bt.BodyTypeDesc " & _
                    "From ((Veh_Stock VS " & _
                    "Left Join BodyBuilder BB On VS.BodyBuilder = BB.BodyBuilderCode) " & _
                    "Left Join BodyType Bt On VS.BodyBuilder_BodyType = Bt.BodyTypeCode) " & _
                    "Order by InDate Desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "Select Top 1 ChassisNo as SearchCode, ChassisNo As Chassis_No, InDate, Model, " & _
                    "BodyBuilder, BodyBuilder_BodyType, BodyBuilder_IssDate, BodyBuilder_RecDate, " & _
                    "BB.BodyBuilderDesc, VS.BodyBuilder_Remark As Remark, Bt.BodyTypeDesc " & _
                    "From ((Veh_Stock VS " & _
                    "Left Join BodyBuilder BB On VS.BodyBuilder = BB.BodyBuilderCode) " & _
                    "Left Join BodyType Bt On VS.BodyBuilder_BodyType = Bt.BodyTypeCode) " & _
                    "Order by InDate Desc", GCn, adOpenDynamic, adLockOptimistic
    
    End If
   
    Set RsChassis = GCn.Execute("select BodyBuilderCode as Code, BodyBuilderDesc Name  from BodyBuilder Order by BodyBuilderDesc")
    Set DgChassis.DataSource = RsChassis
    
    Set RsBodyBuilder = GCn.Execute("select BodyBuilderCode as Code, BodyBuilderDesc As Name  from BodyBuilder Order by BodyBuilderDesc")
    Set DgBodyBuilder.DataSource = RsBodyBuilder
    
    Set RsBodyType = GCn.Execute("Select BodyTypeCode As Code, BodyTypeDesc As Name From BodyType Order By BodyTypeDesc ")
    Set DgBodyType.DataSource = RsBodyType
    
    Disp_Text SETS("INI", Me, Master)
    MoveRec
  
Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsChassis = Nothing
Set Master = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim VNo As Long
Dim I As Integer
    Disp_Text SETS("RecDate", Me, Master)
    Call BlankText
    txt(Chassis).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                GCn.BeginTrans
                    GCn.Execute ("delete from BodyBuilder where BodyBuilderCode= '" & Master!SearchCode & "'")
                GCn.CommitTrans
                
                Master.Requery
                Call MoveRec
                RsChassis.Requery
                
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
    EditName = txt(tCode).TEXT
    EditDesc = txt(Chassis).TEXT
    txt(BodyBuilder).SetFocus
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
    GSQL = "select VS.ChassisNo as SearchCode, VS.ChassisNo, VS.Model, BB.BodyBuilderDesc As Body_Builder, " & cCStr("VS.BodyBuilder_IssDate") & " As Issue_Date, " & cCStr("VS.BodyBuilder_RecDate") & " As Rec_Date, VS.BodyBuilder_Remark As Remark From Veh_Stock VS Left Join BodyBuilder BB On VS.BodyBuilder=BB.BodyBuilderCode Order By BB.BodyBuilderDesc"
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
        Set Master = GCn.Execute("Select ChassisNo as SearchCode, ChassisNo As Chassis_No, InDate, Model, " & _
                    "BodyBuilder, BodyBuilder_BodyType, BodyBuilder_IssDate, BodyBuilder_RecDate, " & _
                    "BB.BodyBuilderDesc, VS.BodyBuilder_Remark As Remark, Bt.BodyTypeDesc " & _
                    "From ((Veh_Stock VS " & _
                    "Left Join BodyBuilder BB On VS.BodyBuilder = BB.BodyBuilderCode) " & _
                    "Left Join BodyType Bt On VS.BodyBuilder_BodyType = Bt.BodyTypeCode) " & _
                    "Where ChassisNo = '" & MyValue & "' " & _
                    "Order by InDate Desc")
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
    RsChassis.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim ItemCode As Integer
    Dim Rst As ADODB.Recordset
    Dim mMaxId As String
    Dim mCondStr$
   On Error GoTo errlbl
   
    If txt(IssDate) <> "" And txt(RecDate) <> "" Then
        If RetDate(txt(IssDate)) > RetDate(txt(RecDate)) Then
            MsgBox "Issue Date Is Greater than Receive Date"
            txt(RecDate).SetFocus
            Exit Sub
        End If
    ElseIf txt(IssDate) = "" And txt(RecDate) <> "" Then
        MsgBox "Chassis Can't be Received from Body Builder Without Issue"
        txt(IssDate).SetFocus
        Exit Sub
    End If
   
 mMaxId = Master!SearchCode
 
 Grid_Hide
 GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Edit" Then
        GCn.Execute "update Veh_Stock set BodyBuilder='" & txt(BodyBuilder).Tag & "', BodyBuilder_BodyType = '" & txt(BodyType).Tag & "', BodyBuilder_IssDate=" & ConvertDate(txt(IssDate)) & ", BodyBuilder_RecDate=" & ConvertDate(txt(RecDate)) & ", BodyBuilder_Remark = '" & txt(Remark) & "' Where ChassisNo = '" & mMaxId & "'"
    End If
GCn.CommitTrans
mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select ChassisNo as SearchCode, ChassisNo As Chassis_No, InDate, Model, " & _
                    "BodyBuilder, BodyBuilder_BodyType, BodyBuilder_IssDate, BodyBuilder_RecDate, " & _
                    "BB.BodyBuilderDesc, VS.BodyBuilder_Remark As Remark, Bt.BodyTypeDesc " & _
                    "From ((Veh_Stock VS " & _
                    "Left Join BodyBuilder BB On VS.BodyBuilder = BB.BodyBuilderCode) " & _
                    "Left Join BodyType Bt On VS.BodyBuilder_BodyType = Bt.BodyTypeCode) " & _
                    "Where ChassisNo = '" & mMaxId & "' " & _
                    "Order by InDate Desc")
    End If
    RsChassis.Requery
    Master.FIND "SearchCode = '" & mMaxId & "'"
    
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
        Case BodyBuilder
            DgBodyBuilder.Move txt(Index).left, txt(Index).top + txt(Index).height + 30
            If RsBodyBuilder.RecordCount = 0 Or txt(Index).TEXT = "" Then Exit Sub
            If RsBodyBuilder.EOF = True Or RsBodyBuilder.BOF = True Then Exit Sub
            If txt(Index).TEXT <> RsBodyBuilder!Name Then
                RsBodyBuilder.MoveFirst
                RsBodyBuilder.FIND "Name ='" & txt(Index).TEXT & "'"
            End If
        Case BodyType
            DgBodyType.Move txt(Index).left, txt(Index).top + txt(Index).height + 30
            If RsBodyType.RecordCount = 0 Or txt(Index).TEXT = "" Then Exit Sub
            If RsBodyType.EOF = True Or RsBodyType.BOF = True Then Exit Sub
            If txt(Index).TEXT <> RsBodyType!Name Then
                RsBodyType.MoveFirst
                RsBodyType.FIND "Name ='" & txt(Index).TEXT & "'"
            End If
            
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
    Case Chassis
        DGridTxtKeyDown_Mast DgChassis, txt, Index, RsChassis, KeyCode, False, 1
    Case BodyBuilder
        DGridTxtKeyDown DgBodyBuilder, txt, Index, RsBodyBuilder, KeyCode, False, 1
    Case BodyType
        DGridTxtKeyDown DgBodyType, txt, Index, RsBodyType, KeyCode, False, 1
        
End Select
If DgChassis.Visible = False And FrmList.Visible = False And DgBodyType.Visible = False And DgBodyBuilder.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> Remark Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = Remark Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If Index <> tCode Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case Index
    Case BodyBuilder
        DGridTxtKeyPress txt, Index, RsBodyBuilder, KeyAscii, "Name", False
    Case BodyType
        DGridTxtKeyPress txt, Index, RsBodyType, KeyAscii, "Name", False
        
End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case Chassis
        DGridTxtKeyUp_Mast txt, Index, RsChassis, KeyCode, "Name"
        
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

With Master
    If .RecordCount > 0 Then
        txt(tCode) = !SearchCode
        txt(Chassis) = XNull(!Chassis_No)
        txt(Model) = XNull(!Model)
        txt(IssDate) = XNull(!BodyBuilder_IssDate)
        txt(RecDate) = XNull(!BodyBuilder_RecDate)
        txt(BodyBuilder).Tag = XNull(!BodyBuilder)
        txt(BodyBuilder) = XNull(!BodyBuilderDesc)
        txt(BodyType).Tag = XNull(!BodyBuilder_BodyType)
        txt(BodyType) = XNull(!BodyTypeDesc)
        txt(Remark) = XNull(!Remark)
    End If
End With

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
    txt(Chassis).Enabled = False
    txt(Model).Enabled = False
    
    
    TopCtrl1.tAdd = False
    TopCtrl1.tDel = False
End Sub
Private Sub Grid_Hide()
    If DgChassis.Visible = True Then DgChassis.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
    DgBodyBuilder.Visible = False
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error Resume Next
    Select Case Index
        Case BodyBuilder
            If RsBodyBuilder.RecordCount > 0 And txt(Index) <> "" And RsBodyBuilder.EOF = False And RsBodyBuilder.BOF = False Then
                txt(Index).Tag = RsBodyBuilder!Code
                txt(Index) = RsBodyBuilder!Name
            Else
                txt(Index) = ""
                txt(Index).Tag = ""
            End If
        Case BodyType
            If RsBodyType.RecordCount > 0 And txt(Index) <> "" And RsBodyType.EOF = False And RsBodyType.BOF = False Then
                txt(Index).Tag = RsBodyType!Code
                txt(Index) = RsBodyType!Name
            Else
                txt(Index) = ""
                txt(Index).Tag = ""
            End If
        Case IssDate
            txt(Index) = RetDate(txt(Index))
            If txt(Index) = "" Then txt(Index) = PubLoginDate
        Case RecDate
            txt(Index) = RetDate(txt(Index))
    End Select
End Sub










