VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehAMDMast 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Addition / Shortage / Deletion Item Master"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11820
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
   LinkTopic       =   " "
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Begin MSDataGridLib.DataGrid DGMaster 
      Height          =   5880
      Left            =   30
      Negotiate       =   -1  'True
      TabIndex        =   17
      Top             =   1290
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   10372
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Prod_Code"
         Caption         =   "Item Code"
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
         DataField       =   "Prod_Name"
         Caption         =   "Item Name"
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
         DataField       =   "Unit"
         Caption         =   "Unit"
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
         DataField       =   "Rate"
         Caption         =   "Rate"
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
         DataField       =   "StkYN"
         Caption         =   "Stock(Y/N)"
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
            DividerStyle    =   3
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4619.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1635.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1349.858
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGUnit 
      Height          =   2190
      Left            =   9750
      Negotiate       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   3863
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Unit"
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
            DividerStyle    =   3
            ColumnWidth     =   2145.26
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
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
      Height          =   270
      Index           =   2
      Left            =   2145
      MaxLength       =   6
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Txt 
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
      Height          =   270
      Index           =   4
      Left            =   8430
      MaxLength       =   3
      TabIndex        =   5
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   840
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   5295
      MaxLength       =   10
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Txt 
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
      Height          =   270
      Index           =   1
      Left            =   5295
      MaxLength       =   40
      TabIndex        =   2
      Top             =   555
      Width           =   5520
   End
   Begin VB.TextBox Txt 
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
      Height          =   270
      Index           =   0
      Left            =   2145
      MaxLength       =   10
      TabIndex        =   1
      Top             =   555
      Width           =   1590
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   661
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock(Y/N)"
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
      Index           =   0
      Left            =   6990
      TabIndex        =   15
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   840
      Visible         =   0   'False
      Width           =   855
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
      Height          =   270
      Index           =   0
      Left            =   8220
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   180
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
      Height          =   270
      Index           =   4
      Left            =   5085
      TabIndex        =   13
      Top             =   555
      Width           =   195
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
      Height          =   270
      Index           =   31
      Left            =   5085
      TabIndex        =   12
      Top             =   840
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      Index           =   34
      Left            =   3855
      TabIndex        =   11
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
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
      Index           =   10
      Left            =   3855
      TabIndex        =   10
      Top             =   555
      Width           =   915
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      Index           =   9
      Left            =   705
      TabIndex        =   9
      Top             =   555
      Width           =   855
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
      Height          =   270
      Index           =   3
      Left            =   1935
      TabIndex        =   8
      Top             =   555
      Width           =   195
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
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
      Left            =   705
      TabIndex        =   7
      Top             =   840
      Width           =   330
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
      Height          =   270
      Index           =   23
      Left            =   1935
      TabIndex        =   6
      Top             =   840
      Width           =   180
   End
End
Attribute VB_Name = "frmVehAMDMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Master As ADODB.Recordset
Dim RsUnit As ADODB.Recordset
Dim mSearchCode As String
Public MasterFormExit As Boolean
Private Const ItemCode As Byte = 0          ' Item Code
Private Const ItemName As Byte = 1          ' Item Name
Private Const Unit As Byte = 2              ' Unit
Private Const Rate As Byte = 3              ' Rate
Private Const StkYN As Byte = 4             ' Stock Y/N

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next
    DGMaster.Enabled = Not Enb
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    Master.MoveFirst
    Master.FIND ("SearchCode='" & MyValue & "'")
    MoveRec
    BUTTONS True, Me, Master, 0
Exit Sub
ELoop:
    CheckError
End Sub
'* Used for clear all text boxes used in the form
Private Sub BlankText()
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).TEXT = ""
    Next I
End Sub

Private Sub MoveRec()
On Error GoTo ELoop
If Master.RecordCount > 0 Then
    Txt(ItemCode).TEXT = Master!Prod_Code
    Txt(ItemName).TEXT = Master!Prod_Name
    Txt(Unit).TEXT = IIf(IsNull(Master!Unit), "", Master!Unit)
    Txt(Rate).TEXT = Format(Master!Rate, "0.00")
    If IsNull(Master!Stock_YN) Then
        Txt(StkYN).TEXT = ""
    Else
        Txt(StkYN).TEXT = IIf(Master!Stock_YN = 0, "No", "Yes")
    End If
End If
    Grid_Hide
'    TopCtrl1.tPrn = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Grid_Hide()
    If DGUnit.Visible = True Then DGUnit.Visible = False
End Sub

Private Sub DGMaster_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 4 Then
        Set Master = GCn.Execute("Select Prod_Code As SearchCode,Veh_AMDModel.*,Switch(Stock_YN=0,'No',Stock_YN=1,'Yes') As StkYN From Veh_AMDModel Order by Stock_YN")
    Else
        Set Master = GCn.Execute("Select Prod_Code As SearchCode,Veh_AMDModel.*,Switch(Stock_YN=0,'No',Stock_YN=1,'Yes') As StkYN From Veh_AMDModel Order by " & DGMaster.Columns(ColIndex).DataField)
    End If
    Master.Requery
    Set DGMaster.DataSource = Master
    MoveRec
End Sub

Private Sub DGMaster_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    MoveRec
End Sub

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
    CheckError
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
'to modify
    DGMaster.top = 1290
    DGMaster.left = 45 '120
    DGUnit.left = 3240: DGUnit.top = 840

    TopCtrl1.Tag = PubUParam: WinSetting Me
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
    Set RsUnit = New ADODB.Recordset
    RsUnit.CursorLocation = adUseClient
    RsUnit.Open "Select Unit_Name as Code,Unit_Name As Name From Unit Order by Unit_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGUnit.DataSource = RsUnit
    Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If



    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    Set Master = GCn.Execute("Select Prod_Code As SearchCode,Veh_AMDModel.*, " & cIIF("Stock_YN=0", "'No'", "'Yes'") & " As StkYN From Veh_AMDModel " & sitecond & " Order by Prod_Code")
    
    Set DGMaster.DataSource = Master
'    If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set Master = Nothing
    Set RsUnit = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    Txt(StkYN).TEXT = "No"
    Txt(ItemCode).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
    Txt(ItemCode).Enabled = False
    Txt(ItemName).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
Dim vBook As Variant
On Error GoTo ELoop
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        GCn.BeginTrans
            GCn.Execute ("Delete From Veh_AMDModel Where Prod_Code='" & Txt(ItemCode).TEXT & "'")
        GCn.CommitTrans
        Master.Requery
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
        End If
        BUTTONS True, Me, Master, 0
        MoveRec
    End If
Exit Sub
ELoop:
    GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub

Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub

Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
        Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(Veh_AMDModel.site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If



    GSQL = "Select " & cTrim("Prod_Code") & " As SearchCode,Prod_Code As ItemCode,Prod_Name As ItemName,Unit, " & cCStr("Rate") & " As Rate, " & cIIF("Stock_YN=0", "'No'", "'Yes'") & " As Stock FROM Veh_AMDModel " & sitecond & " Order by Prod_Code"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_ePrn()
Dim Rst As ADODB.Recordset
Dim mQry As String
    mQry = Master.Source
       
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub

        CreateFieldDefFile Rst, PubRepoPath + "\VehAMDList.TTX", True
        Set rpt = rdApp.OpenReport(PubRepoPath & "\VehAMDList.RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
        
        Call Report_View(rpt, "Vehicle Addition/Deletion/Shortage ItemList")

End Sub

Private Sub TopCtrl1_eRef()
    RsUnit.Requery
End Sub

Private Sub TopCtrl1_eSave()
Dim mTrans As Boolean
On Error GoTo ELoop
    Grid_Hide
    If IsValid(Txt(ItemCode), "Item Code") = False Then Exit Sub
    If IsValid(Txt(ItemName), "Item Name") = False Then Exit Sub

    GCn.BeginTrans
        mTrans = True
        If TopCtrl1.TopText2 = "Add" Then
            GCn.Execute "Insert Into Veh_AMDModel(" _
            & "Prod_Code,Site_Code,Prod_Name,Unit,Rate," _
            & "Stock_YN,U_Name,U_EntDt,U_AE) " _
            & "Values(" _
            & "'" & Txt(ItemCode).TEXT & "','" & PubSiteCode & "','" & Txt(ItemName).TEXT & "','" & Txt(Unit).TEXT & "'," & Val(Txt(Rate).TEXT) & "," _
            & "" & IIf(Txt(StkYN).TEXT = "No", 0, 1) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
        Else
            GCn.Execute "Update Veh_AMDModel Set " _
            & "Prod_Name='" & Txt(ItemName).TEXT & "',Unit='" & Txt(Unit).TEXT & "',Rate=" & Val(Txt(Rate).TEXT) & "," _
            & "Stock_YN=" & IIf(Txt(StkYN).TEXT = "No", 0, 1) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' " _
            & "Where Prod_Code='" & Txt(ItemCode).TEXT & "'"
        End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    mTrans = False
    mSearchCode = Txt(ItemCode)
    Master.Requery
    Master.FIND "SearchCode = '" & mSearchCode & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    If mTrans = True Then
        GCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    If MasterFormExit Then Unload Me: Exit Sub
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To Txt.Count - 1
            Txt(I).BackColor = CtrlBColOrg
            Txt(I).ForeColor = CtrlFColOrg
        Next
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub DGUnit_Click()
On Error GoTo ELoop
    DGUnit.Visible = False
    If RsUnit.RecordCount > 0 Then
        Txt(Unit).TEXT = RsUnit!Name
    End If
    Txt(Unit).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus Txt(Index)
    Grid_Hide
    Select Case Index
    Case Unit
        If RsUnit.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsUnit!Name Then
            RsUnit.MoveFirst
            RsUnit.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
    Case Unit
        DGridTxtKeyDown DGUnit, Txt, Unit, RsUnit, KeyCode, False, 1
    Case Rate
        NumDown Txt(Index), KeyCode, 7, 2
    End Select
    If DGUnit.Visible = False Then
        If Index <> StkYN Then
            If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = StkYN Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" Then
            If Index <> ItemCode And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
            If Index <> ItemName And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
    Select Case Index
    Case Unit
        If DGUnit.Visible = True Then DGridTxtKeyPress Txt, Unit, RsUnit, KeyAscii, "Name"
    Case StkYN
        If KeyAscii = 89 Or KeyAscii = 121 Or KeyAscii = 78 Or KeyAscii = 110 Then
            If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                Txt(Index).TEXT = "Yes"
                KeyAscii = 0
            ElseIf KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                Txt(Index).TEXT = "No"
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
    Case Rate
        NumPress Txt(Index), KeyAscii, 7, 2
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, I As Byte
On Error GoTo ELoop
    Select Case Index
        Case ItemCode
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select Prod_Code From Veh_AMDModel Where Prod_code='" & Txt(ItemCode).TEXT & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Item Code Not Allowed", vbInformation, "Validation"
                Cancel = True
                Txt(ItemCode).SetFocus
            End If
        Case Unit
            If RsUnit.RecordCount > 0 And Txt(Index).TEXT <> "" Then
                Txt(Index).TEXT = RsUnit!Name
            End If
        Case Rate
            If Val(Txt(Index).TEXT) = "0.00" Or Txt(Index).TEXT = "0" Then
                Txt(Index).TEXT = ""
            Else
                Txt(Index).TEXT = Format(Txt(Index), "0.00")
            End If
        Case StkYN
            If Not Trim(Txt(Index).TEXT) <> "Yes" Or Trim(Txt(Index).TEXT) <> "No" Then
                Txt(Index).TEXT = "No"
            End If
    End Select
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub
