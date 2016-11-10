VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form YearEnd 
   BackColor       =   &H80000005&
   Caption         =   "Year End Process"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameIns 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   5190
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7770
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update To New Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2610
         TabIndex        =   17
         Top             =   4680
         Width           =   2130
      End
      Begin VB.Shape Shape8 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   240
         Top             =   2865
         Width           =   180
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   255
         Top             =   2145
         Width           =   180
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   255
         Top             =   1410
         Width           =   180
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   255
         Top             =   570
         Width           =   180
      End
      Begin VB.Line Line3 
         X1              =   15
         X2              =   7770
         Y1              =   4545
         Y2              =   4545
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   7770
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do Not Switch Off the System or Stop the Process in Middle.This may damage your data.. "
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
         Height          =   480
         Left            =   750
         TabIndex        =   16
         Top             =   3720
         Width           =   6945
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape9 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   240
         Top             =   3750
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"YearEnd.frx":0000
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
         Height          =   720
         Left            =   750
         TabIndex        =   15
         Top             =   2790
         Width           =   6930
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insure that there may not be any Inconsistancy in the Power Supply while Upgrading to the next year.this may loss your data."
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
         Height          =   480
         Left            =   765
         TabIndex        =   14
         Top             =   2115
         Width           =   6930
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Take the Complete Backup of Your data before using the Utility."
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
         Height          =   480
         Left            =   765
         TabIndex        =   13
         Top             =   1395
         Width           =   6915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"YearEnd.frx":0092
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
         Height          =   720
         Left            =   795
         TabIndex        =   12
         Top             =   525
         Width           =   6915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Instructions :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   225
         TabIndex        =   11
         Top             =   60
         Width           =   2100
      End
   End
   Begin VB.TextBox Status 
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
      ForeColor       =   &H00000080&
      Height          =   1590
      Left            =   4980
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1395
      Width           =   2145
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   420
      Top             =   4545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel Process"
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
      Left            =   3015
      TabIndex        =   6
      Top             =   4560
      Width           =   2010
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   840
      TabIndex        =   8
      Top             =   4020
      Visible         =   0   'False
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   476
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008080&
      X1              =   4695
      X2              =   7290
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Automan"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   630
      Left            =   4755
      TabIndex        =   9
      Top             =   3195
      Width           =   2490
   End
   Begin VB.Image Image2 
      Height          =   225
      Index           =   4
      Left            =   780
      Picture         =   "YearEnd.frx":011D
      Top             =   3345
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   225
      Index           =   3
      Left            =   780
      Picture         =   "YearEnd.frx":042F
      Top             =   2835
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   225
      Index           =   2
      Left            =   780
      Picture         =   "YearEnd.frx":0741
      Top             =   2370
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   225
      Index           =   1
      Left            =   780
      Picture         =   "YearEnd.frx":0A53
      Top             =   1905
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   225
      Index           =   0
      Left            =   780
      Picture         =   "YearEnd.frx":0D65
      Top             =   1425
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   4230
      Left            =   195
      Shape           =   4  'Rounded Rectangle
      Top             =   195
      Width           =   7380
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00004080&
      Height          =   2760
      Left            =   4695
      Top             =   1110
      Width           =   2610
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00004080&
      FillColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   405
      Top             =   1110
      Width           =   4110
   End
   Begin VB.Image Image1 
      Height          =   210
      Index           =   4
      Left            =   585
      Picture         =   "YearEnd.frx":1077
      Stretch         =   -1  'True
      Top             =   3375
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   210
      Index           =   3
      Left            =   585
      Picture         =   "YearEnd.frx":11D5
      Stretch         =   -1  'True
      Top             =   2850
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   210
      Index           =   2
      Left            =   600
      Picture         =   "YearEnd.frx":1333
      Stretch         =   -1  'True
      Top             =   2370
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   1
      Left            =   600
      Picture         =   "YearEnd.frx":1491
      Stretch         =   -1  'True
      Top             =   1890
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   210
      Index           =   0
      Left            =   615
      Picture         =   "YearEnd.frx":15EF
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label Process 
      BackStyle       =   0  'Transparent
      Caption         =   "Completing Year End Process"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   4
      Left            =   1170
      TabIndex        =   5
      Top             =   3345
      Width           =   3180
   End
   Begin VB.Label Process 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfering FAData.mdb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   3
      Left            =   1155
      TabIndex        =   4
      Top             =   2850
      Width           =   2865
   End
   Begin VB.Label Process 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfering Automan.mdb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   2
      Left            =   1140
      TabIndex        =   3
      Top             =   2370
      Width           =   2805
   End
   Begin VB.Label Process 
      BackStyle       =   0  'Transparent
      Caption         =   "Creating new Data Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   1
      Left            =   1140
      TabIndex        =   2
      Top             =   1905
      Width           =   3210
   End
   Begin VB.Label Process 
      BackStyle       =   0  'Transparent
      Caption         =   "Searching Data Directories"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   0
      Left            =   1155
      TabIndex        =   1
      Top             =   1395
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Process Status"
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
      Height          =   240
      Left            =   3135
      TabIndex        =   0
      Top             =   480
      Width           =   1635
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   2535
      Top             =   390
      Width           =   2835
   End
End
Attribute VB_Name = "YearEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SearchDir As Byte = 0
Private Const CreateDir As Byte = 1
Private Const TransAuto As Byte = 2
Private Const TransFA As Byte = 3
Private Const CompProcess As Byte = 4
Dim G_CompCntmp As ADODB.Recordset
Dim G_Rsttmp As ADODB.Recordset
Dim j As Integer
Dim TRec1Qty As Single, TRec2Qty As Single
Dim mRec_TB_Qty As Double, mRec_TB_Val As Double, mRec_TP_Qty As Double, mRec_TP_Val As Double, StartDate$
Dim mIss_TB_Qty As Double, mIss_TB_Val As Double, mIss_TP_Qty As Double, mIss_TP_Val As Double
Dim xMOP_TBQty As Double, xMOP_TBVal As Double, xMOP_TPQty As Double, xMOP_TPVal As Double, mRate As Double
Dim mOP_TB_QTY As Double, mOP_TP_QTY As Double, mOP_TB_VAL As Double, mOP_TP_VAL As Double
Dim mQry$, Condstr$, CondDivCode$, CondMarkYN$, CondPartNos$, CondPartNos1$, CondDivCode1$, mNarr$
Dim CondStrMRP$, CondPartNosOpStk$, CondPartNosTrn$
Dim NEW_CONNECTION As ADODB.Connection

Private Sub Command1_Click()
    If MsgBox("Are You Sure to Cancel ? ", vbInformation + vbYesNo, "Year End Process") = vbYes Then
        End
    End If
End Sub

Private Sub CmdUpdate_Click()
FrameIns.Visible = False
Dim I As Integer, SourcePath$, DestPath$, SourcePathFA$, DestPathFA$, FAFolderName$, FolderName$, NoFolder As Boolean
Dim FS As FileSystemObject, App_Path$, DocID$, OpStk$, OpRate$
Set FS = New FileSystemObject

Dim mRecQty As Double, mIssQty As Double, mStkVal As Double
Dim XRecNo As Double
Dim CODE1 As Long, mCode As String
Dim cr As Double, dr As Double, SR As Long
Dim mS_NO As Long, RST1 As ADODB.Recordset
Dim mDR_BALANCE As Double, mBALANCE As Double, mPartyCode As Long, mDocId As String
Dim mVNo As Long
Dim TRec1 As ADODB.Recordset, TRec2 As ADODB.Recordset, Temp06 As ADODB.Recordset
Dim RstPart As ADODB.Recordset
Dim RstDiv As ADODB.Recordset
Dim RstStock As ADODB.Recordset, RstStock2 As ADODB.Recordset, RstStock3 As ADODB.Recordset
Dim GCN1 As ADODB.Connection
Dim RstParty As ADODB.Recordset
Dim mDelete As Boolean
Dim RsTemp As ADODB.Recordset
Dim rsTemp1 As ADODB.Recordset
Dim ObjSqlServer As Object

'On Error Resume Next
If MsgBox("Start The Year End Process", vbInformation + vbYesNo) = vbYes Then
        GCn.CommandTimeout = 1024
        'For Rashmi Motors
        StartDate = PubStartDate
        If PubStartDate = "17/Dec/2003" Then
                PubStartDate = "01/Apr/2003"
                PubEndDate = DateAdd("d", -1, DateAdd("yyyy", 1, PubStartDate))
        End If
        Set G_CompCntmp = New ADODB.Recordset
             With G_CompCntmp
                    .Open "Select * from Company", G_CompCn, adOpenDynamic, adLockOptimistic, adCmdText
             End With

'SEARCHING FOR THE FOLDERS
    If PubBackEnd = "A" Then
        Status = "Searching Folder " & FolderName & "in Data Directory..."
        For I = 1 To 99
            ImageDisp (0)
            If FS.FolderExists(Pub_DataPath & "\Auto_" & Format(I, "00")) = False Then
                NoFolder = True
                FolderName = "Auto_" & Format(I, "00")
                FAFolderName = "FAData_" & Format(I, "00")
                Status = "Folder " & FolderName & " Not Found. Creating New... "
                Exit For
            End If
        Next
    
'CREATING FOLDERS
        If NoFolder Then
            ImageDisp (1)
            FS.CreateFolder (Pub_DataPath & "\" & FolderName)
            Status = "Folder " & FolderName & " Created..."
    '       PValue (35)
        End If
    Else
        FolderName = Trim(left(PubComp_Name, 4)) & "_" & Right(Format(date, "YYYY"), 2)
    End If
'COPYING AUTOMAN.MDB TO THE DESTINATION
        ImageDisp (2)
        Status = "Transfering Vehicle,WorkShop and Spare's Data to the New Finencial Year...."
        If PubBackEnd = "A" Then
            SourcePath = Pub_DataPath & "\" & PubCenDataPath & "\" & "Automan.mdb"
            DestPath = Pub_DataPath & "\" & FolderName & "\Automan.mdb"
            FS.CopyFile SourcePath, DestPath, True
        Else
            Set ObjSqlServer = CreateObject("SQLDMO.SQLServer")
            ObjSqlServer.Connect PubServerName, "sa", ""
            PubSQLDataPath = ObjSqlServer.DataBases.Item(PubCenDataPath).PrimaryFilePath
            PubBkupDataPath = ObjSqlServer.DataBases.Item(PubCenDataPath).PrimaryFilePath
            
            DataFileName = GCn.Execute("Select File_Name(1)").Fields(0).Value
            TransactionFileName = GCn.Execute("Select File_Name(2)").Fields(0).Value
            
            StrBackupPath = PubBkupDataPath

            mBackupFile = StrBackupPath & PubCenDataPath & CStr(Format(PubLoginDate, "DDMMYY")) + ".bak"
            
            If FS.FileExists(mBackupFile) Then
                FS.DeleteFile mBackupFile
            End If
            'StrBackupPath = "\\" + PubServerName + "\" + Replace(PubBkupDataPath, ":", "")
            GCn.Execute ("BACKUP DATABASE " & PubCenDataPath & "  TO  Disk =  '" & StrBackupPath & PubCenDataPath + ".bak" & "' ")

                        
        End If
        Status = "Transfering Finencial Data to the New Finencial Year...."
'        PValue (65)
'COPYING FADATA.MDB TO THE DESTINATION
        ImageDisp (3)
        If PubBackEnd = "A" Then
            FS.CreateFolder (Pub_DataPath & "\" & FolderName & "\" & FAFolderName)
            SourcePathFA = PubFADataPath
            DestPathFA = Pub_DataPath & "\" & FolderName & "\" & FAFolderName & "\FAData.mdb"
            FS.CopyFile SourcePathFA, DestPathFA, True
            Status = "Doing Updations In WorkShop and Spare's Transfered Data.... "
            ProgressBar1.Visible = True
            Set GCn = New ADODB.Connection
                 With GCn
                        .CursorLocation = adUseClient
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .ConnectionString = "Data Source=" & DestPath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
                        .Open
                        .BeginTrans
                 End With
        Else
            mNewDatabase = Trim(left(Replace(Replace(PubComp_Name, " ", ""), ".", ""), 4)) & "_" & Right(Format(date, "YYYY"), 2)
            
            
            mCy = ""
            mDBCount = 0

'            Do While True
'                If G_CompCn.Execute("SELECT COUNT(*) FROM COMPANY WHERE CentralData_Path=" & Chk_Text(mNewDatabase)).Fields(0).Value > 0 Then
'                    mDBCount = mDBCount + 1
'                    mCy = "_" + Trim(STR(mDBCount))
'                    mNewDatabase = PubDBPrefix + MidStr(Year(PubStartDate.AddYears(1)).ToString, 2, 2) + mCy
'                Else
'                    Exit Do
'                End If
'            Loop

'            ' Restore DataBase
'            GCn.Close
'            GCn = New ADODB.Connection
'            GCn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
'            GCn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubDataBaseName & " ;Data Source=" & PubServerName
'            GCn.CommandTimeout = 1024
'            GCn.Open

            GCn.CommandTimeout = 1024
            mQry = "RESTORE DATABASE  " & mNewDatabase & " from DISK  =  '" & StrBackupPath & PubCenDataPath + ".bak" & "' " & _
                   "With MOVE  '" & DataFileName & "'   To " & Chk_Text(Trim(PubSQLDataPath) + Trim(mNewDatabase) + ".MDF") & " ," & _
                    "MOVE '" & TransactionFileName & "' To " & Chk_Text(Trim(PubSQLDataPath) + Trim(mNewDatabase) + ".LDF") & " "
            GCn.Execute (mQry)
        
            
            GCn.Close
            GCn = New ADODB.Connection
            GCn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            GCn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & mNewDatabase & " ;Data Source=" & PubServerName
            GCn.CommandTimeout = 1024
            GCn.Open
            GCn.BeginTrans
        End If
        ImageDisp (4)
'UPDATING Company TABLE
        If PubBackEnd = "A" Then
            With G_CompCntmp
                .AddNew
                .Fields("Comp_Code") = Right(FolderName, 2)
                .Fields("Comp_Name") = PubComp_Name
                .Fields("CentralData_Path") = FolderName
                .Fields("start_date") = DateAdd("yyyy", 1, PubStartDate)
            End With
            G_CompCntmp.Update
        Else
            G_CompCn.Execute "Insert Into Company(Comp_Code, Comp_Name, CentralData_Path, Start_Date) " & _
                             "Values('" & Right(Format(date, "YYYY"), 2) & "', '" & PubComp_Name & "', '" & Trim(left(PubComp_Name, 4)) & "_" & Right(Format(date, "YYYY"), 2) & "', " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & ")"
        End If
        G_CompCntmp.Close
        If PubBackEnd = "A" Then
            G_CompCn.Execute "Update Company set OldPath='" & PubCenDataPath & "\Automan.mdb" & " '" & ",OldPathFA='" & mID(PubFADataPath, Len(Pub_DataPath) + 2, Len(PubFADataPath)) & "' where Comp_Code ='" & Right(FolderName, 2) & "'"
        Else
            G_CompCn.Execute "Update Company set OldPath='" & PubCenDataPath & " '" & ",OldPathFA='" & PubCenDataPath & "' where Comp_Code ='" & Right(FolderName, 2) & "'"
        End If
'UPDATING USER1 TABLE
'        G_CompCn.Execute "insert into User1(User_Name,comp_code,Div_Code,Div_Name,Mod_Veh,Mod_Spr,Mod_Wsp,Mod_Acc,Mod_Set,PARAM_STR)" & _
'                         " values('" & pubUName & "','" & Right(FolderName, 2) & "','C','CVD',1,1,1,1,1,'*')"
'
'        G_CompCn.Execute "insert into User1(User_Name,comp_code,Div_Code,Div_Name,Mod_Veh,Mod_Spr,Mod_Wsp,Mod_Acc,Mod_Set,PARAM_STR)" & _
'                         " values('" & pubUName & "','" & Right(FolderName, 2) & "','P','PCD',1,1,1,1,1,'*')"
'
'        G_CompCn.Execute "Update User1 set Mod_Veh=0"
        Set rsTemp1 = GCn.Execute("Select Div_Code, Div_SName From Division")
        If rsTemp1.RecordCount > 0 Then
            Do Until rsTemp1.EOF
                Set RsTemp = G_CompCn.Execute("Select * From User1 Where Comp_Code='" & PubCenCompCode & "' And Div_Code='" & rsTemp1!Div_Code & "'")
                If RsTemp.RecordCount > 0 Then
                    Do Until RsTemp.EOF
                        G_CompCn.Execute "Insert Into User1(User_Name,comp_code,Div_Code,Div_Name,Mod_Veh,Mod_Spr,Mod_Wsp,Mod_Acc,Mod_Set,PARAM_STR)" & _
                                         " values('" & RsTemp!user_name & "','" & Right(FolderName, 2) & "','" & rsTemp1!Div_Code & "','" & rsTemp1!Div_SName & "'," & VNull(RsTemp!mod_veh) & "," & VNull(RsTemp!mod_spr) & "," & VNull(RsTemp!mod_wsp) & "," & VNull(RsTemp!mod_acc) & "," & VNull(RsTemp!mod_set) & ",'" & XNull(RsTemp!param_str) & "')"
                        RsTemp.MoveNext
                    Loop
                Else
                    G_CompCn.Execute "insert into User1(User_Name,comp_code,Div_Code,Div_Name,Mod_Veh,Mod_Spr,Mod_Wsp,Mod_Acc,Mod_Set,PARAM_STR)" & _
                                     " values('" & pubUName & "','" & Right(FolderName, 2) & "','" & PubDivCode & "','" & PubDivSName & "',1,1,1,1,1,'*')"
                End If
                
                rsTemp1.MoveNext
            Loop
        End If
        
        
'UPDATING USER2 TABLE
        G_CompCn.Execute ("insert into user2  select User_Name,Module_Name,Form_Code,Param_Str,'" & Right(FolderName, 2) & "' as Comp_Code,Div_Code from User2 where Comp_Code='" & PubCenCompCode & "'")

'UPDATING ASSOCIATEDFIRMS TABLE
        GCn.Execute "Update AssociatedFirms set FADataPath = '" & FolderName & "\" & FAFolderName & "',AssoComp_Code='" & Right(FolderName, 1) & "'"

'UPDATING DEVISION TABLE
        GCn.Execute "Update Division set V_SecCompCode = '" & Right(FolderName, 1) & "',S_SecCompCode = '" & Right(FolderName, 1) & "',W_SecCompCode = '" & Right(FolderName, 1) & "'"
        GCn.Execute "Update Division set V_SecFADataPath = '" & FolderName & "\" & FAFolderName & "',S_SecFADataPath = '" & FolderName & "\" & FAFolderName & "',W_SecFADataPath = '" & FolderName & "\" & FAFolderName & "'"
        ProgressBar1.Value = 5
'***********************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
' Workshop  Section Year End Updation

'UPDATING ESTIMATE TABLE
        GCn.Execute ("Delete from Estimate where V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "")
        
'UPDATING ESTIMATE1 TABLE
        GCn.Execute ("Delete from Estimate1 where DocId Not in (select DocId from Estimate)  ")

'UPDATING INDENT TABLE
        GCn.Execute ("Delete from Indent ")

'UPDATING JOB BOOKING TABLE
        GCn.Execute ("Delete from Job_Booking where Book_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "")
        ProgressBar1.Value = 12

'UPDATING JOB_CARD TABLE
        GCn.Execute ("Delete from Job_Card where len(JobCloseDate)  > 1  And Job_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & " ")
        GCn.Execute ("Delete from SP_Stock where Job_DocID not in (Select DocID from Job_Card) and V_Type in ('W_RG','W_RW') and Len(Job_DocId)  >1 ")
        GCn.Execute ("Delete from SP_Sale where Job_DocID not in (Select DocID from Job_Card) and Len(Job_DocId)  >1 ")
        ProgressBar1.Value = 15

'UPDATING JOB_CARD2 TABLE
        GCn.Execute ("Delete from Job_Card2 where DocID not in (select DocID from Job_Card)")

'UPDATING JOB_DEMAND TABLE
        GCn.Execute ("Delete from Job_Demand where Job_DocID not in (select DocID from Job_Card)")
        ProgressBar1.Value = 18

'UPDATING JOB_GATEPASS TABLE
        GCn.Execute ("Delete from Job_GatePass where Job_DocID not in (select DocID from Job_Card)")

'UPDATING JOB_GATEPASS1 TABLE
        GCn.Execute ("Delete from Job_GatePass1 where GatePassNo not in (select GatePassNo from Job_GatePass)")
        ProgressBar1.Value = 20
        
'UPDATING INSPECTION TABLE
        GCn.Execute ("Delete from Job_Inspection where Job_DocId not in (select DocId from Job_Card)")

'UPDATING INSPECTION2 TABLE
        GCn.Execute ("Delete from Job_Inspection2 where DocId not in (select DocId from Job_Inspection)")
        
'UPDATING JOB_LAB TABLE
        GCn.Execute ("Delete from Job_Lab where Job_DocId not in (select DocId from Job_Card)")

'UPDATING JOB_LAB2 TABLE
        GCn.Execute ("Delete from Job_Lab2 where Job_DocId not in (select DocId from Job_Card)")
        ProgressBar1.Value = 24

'UPDATING JOB_WARBILL TABLE
'        GCn.Execute ("Delete from Job_WarrBill")
        
'UPDATING JOB_WARR1 TABLE
        GCn.Execute ("Delete from Job_Warr1")

'UPDATING JOB_WARR2 TABLE
        GCn.Execute ("Delete from Job_Warr2")

'UPDATING RECT TABLE


'UPDATING SALETARGET TABLE
        GCn.Execute ("Delete from SaleTarget")
        
'UPDATING SP_ORDCOUN TABLE
        

'UPDATING SP_ORDER TABLE
        GCn.Execute ("Delete from SP_Order where Len(OrdClosDate)  > 1  and V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & " ")
        
'UPDATING SP_ORDER1 TABLE
        GCn.Execute ("Delete from SP_Order1 where OrderId not in (select OrderId from SP_Order)")
        ProgressBar1.Value = 26
'UPDATING SP_PURCH TABLE
        Set G_CompCntmp = New ADODB.Recordset
        G_CompCntmp.Open "Select DocID from SP_Purch where len(SP_Purch.Invoice_DocId) > 1  and V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "", GCn, adOpenDynamic, adLockOptimistic, adCmdText
        For I = 1 To G_CompCntmp.RecordCount
            GCn.Execute ("Delete from SP_Stock where DocId='" & G_CompCntmp!DocID & "'")
        Next
        G_CompCntmp.Close
        GCn.Execute ("Delete from SP_Purch where len(SP_Purch.Invoice_DocId) > 1   and V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "")
        GCn.Execute ("Delete from SP_Purch where len(SP_Purch.Invoice_DocId) = 0 and V_type <> 'SXGR'  and V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "")
        ProgressBar1.Value = 30
'UPDATINfG SP_SALE TABLE
'        Set G_CompCntmp = New ADODB.Recordset
'        G_CompCntmp.Open "Select DocID from SP_Sale where len(SP_Sale.Invoice_DocId)  > 1 ", GCn, adOpenDynamic, adLockOptimistic, adCmdText
'       For I = 1 To G_CompCntmp.RecordCount
'            GCn.Execute ("Delete from SP_Stock where DocId ='" & G_CompCntmp!DocId & "'")
'        GCn.Execute ("Delete from SP_Stock where DocId in (Select DocID from SP_Sale where len(SP_Sale.Invoice_DocId)  > 1 )")
            
'       Next
'        G_CompCntmp.Close
        GCn.Execute ("Delete from SP_Sale where len(SP_Sale.Invoice_DocId) > 1  and V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "  ")
        GCn.Execute ("Delete from SP_Sale where V_Type in ('SXSRC','SXSRR','SYSIC','SYSIR','W_SIC','W_SIR')  and V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & " ")
        ProgressBar1.Value = 35
        
'UPDATING SP_STOCK TABLE
        GCn.Execute ("Delete from SP_Stock where V_Type='SXAO'")
        If ProgressBar1.Value < 80 Then
            ProgressBar1.Value = 80
        End If
        GCn.Execute ("Delete from SP_Stock where Job_DocID not in (Select DocId from Job_Card) and len(Job_DocID)  > 1")
        GCn.Execute ("Delete from SP_Stock where V_Type in ('SYPRC','SYPRR','SXSRC','SXSRR','SXSRT','SXGRT','SYIAD','SXRAD') and V_Date< " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & " ")
        'GCn.Execute ("Delete from SP_Stock where len(Invoice_DocID)  > 1  and V_Type in ('SYSC','SYSIC','SYSIR','W_SIC','W_SIR')")
        GCn.Execute ("Delete from SP_Stock where len(Invoice_DocID)  > 1 and V_Date< " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "  ")
        
        GCn.CommitTrans
        
        GCn.BeginTrans
        ProgressBar1.Value = 84
'********************************************************************************************************************************************************************************************************

'Vehicle Section Year End Updation

'UPDATING VEH_FORECAST TABLE
        Status = "Doing Updations In Vehicle's Transfered Data.... "
        GCn.Execute ("Delete from Veh_ForeCast")

'UPDATING VEH_OFFTAKEINCENTIVE TABLE
        GCn.Execute ("Delete from Veh_OfftakeIncentive")

'UPDATING VEH_ORDER TABLE
        GCn.Execute ("Delete from Veh_Order Where len(Inv_DocID)  >  1 and Inv_Date<" & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "  and Ord_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "")
'
''UPDATING VEH_ORDER1 TABLE
        GCn.Execute ("Delete from Veh_Order1 where OrdDocID not in (Select OrdDocId from Veh_Order)")

'UPDATING VEH_INVCANCEL TABLE
        GCn.Execute ("Delete from Veh_InvCancel where OrdDocID not in (Select OrdDocId from Veh_Order)")
        
'UPDATING VEH_STOCK TABLE
        GCn.Execute ("Delete from Veh_Stock where len(Sal_DocID) > 1 and Sal_VDate < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "  and Pur_VDate < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "  ")
        Set G_CompCntmp = New ADODB.Recordset
        G_CompCntmp.Open "select  Pur_DocID,Pur_VType,Pur_DocIDHelp from Veh_Stock where (len(Sal_DocID) =0 or Sal_DocID is null)  and Pur_VDate < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "  ", GCn, adOpenDynamic, adLockOptimistic, adCmdText
        
        For I = 1 To G_CompCntmp.RecordCount
            DocID = left(G_CompCntmp!Pur_DocId, 3) & "V_OST" & " " & CStr(Year(PubStartDate)) & Right(date, 1) & Right(G_CompCntmp!Pur_DocId, 7)
            GCn.Execute "Update Veh_Purch1 set DocID='" & DocID & "' , V_Type = 'V_OST' ,DocIDHelp='" & Trim(UCase(DocID)) & "'where DocId='" & G_CompCntmp!Pur_DocId & "'"
            GCn.Execute "Update Veh_Purch2 set DocID='" & DocID & "' , V_Type = 'V_OST' where DocId='" & G_CompCntmp!Pur_DocId & "'"
            GCn.Execute "Update Veh_Stock set Pur_DocID='" & DocID & "', Pur_DocIdHelp='" & DocID & "' , Pur_VType = 'V_OST' where Pur_DocId='" & G_CompCntmp!Pur_DocId & "'"
'           G_CompCntmp!Pur_DocId = DocID
'           G_CompCntmp!Pur_DocIDHelp = Trim(UCase(DocID))
'           G_CompCntmp!Pur_VType = "V_OST"
'           G_CompCntmp.Update
            G_CompCntmp.MoveNext
        Next
        ProgressBar1.Value = 90
        G_CompCntmp.Close
'UPDATING VEH_PURCH1 TABLE
        GCn.Execute ("Delete from Veh_Purch1 where DocID not in (Select Pur_DocID from Veh_Stock)")

'UPDATING VEH_PURCH2 TABLE
        GCn.Execute ("Delete from Veh_Purch2 where DocID not in (Select DocID from Veh_Purch1)")

'UPDATING VEH_QUOT TABLE
        GCn.Execute ("Delete from Veh_Quot Where V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & " ")

'UPDATING VEH_QUOT1 TABLE
        GCn.Execute ("Delete from Veh_Quot1 where DocId Not in (Select DocId From Veh_Quot)")

'UPDATING VEH_QUOT2 TABLE
        GCn.Execute ("Delete from Veh_Quot2 where DocId Not in (Select DocId From Veh_Quot)")

'UPDATING VEH_SUBGROUPQUOT TABLE
        GCn.Execute ("Delete from Veh_SubGroupQuot Where StartDate < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & " ")

'UPDATING VEH_TARGET TABLE
        GCn.Execute ("Delete from Veh_Target")

'UPDATING VEH_TRANSFER TABLE
        GCn.Execute ("Delete from Veh_Transfer Where V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "")

'UPDATING VEH_CHECKLIST TABLE
        GCn.Execute ("Delete from Veh_CheckList where Model not in (Select Model from Veh_Stock)")

'UPDATING VISIT TABLE
'        GCn.Execute ("Delete from Visit")

'UPDATING SUBGROUP TABLE
        GCn.Execute ("Update SubGroup Set FirmCode='" & Right(FolderName, 1) & "'")

'UPDATING SUBGROUPALIAS TABLE
        GCn.Execute ("Update SubGroupAlias Set FirmCode='" & Right(FolderName, 1) & "'")
        
        ProgressBar1.Value = 93
'****************************************************************************************************************************

'FA YEAR END UPDATION MODULE

    
    Set RST1 = G_FaCn.Execute("select MAX(Nature) as NAT,MAX(GroupCode) As MCODE,MAX(LEDGER.SUBCODE) AS PARTY_CODE,SUM(AMTCR)-SUM(AMTDR) AS BALANCE FROM LEDGER LEFT JOIN SUBGROUP ON LEDGER.SUBCODE=SUBGROUP.SUBCODE WHERE GroupNature IN ('A','L') GROUP BY GroupCode,NAME,LEDGER.SUBCODE")
    
    If PubBackEnd = "A" Then
        DBEngine.RegisterDatabase "NEW_COMP", "Microsoft Access Driver (*.MDB)", True, "DBQ=" & DestPathFA & vbCr & "Maxbuffersize=2048" & vbCr & "MaxscanRows=16"
        Set NEW_CONNECTION = New ADODB.Connection
        NEW_CONNECTION.CursorLocation = adUseClient
        NEW_CONNECTION.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=NEW_COMP"
    Else
        Set NEW_CONNECTION = New ADODB.Connection
        NEW_CONNECTION.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        NEW_CONNECTION.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & mNewDatabase & " ;Data Source=" & PubServerName
        NEW_CONNECTION.CommandTimeout = 1024
        NEW_CONNECTION.Open
    
    End If
    
    
    Screen.MousePointer = vbHourglass
    NEW_CONNECTION.BeginTrans
    mBeginTrans = 1
    Status = "Making Table Updates For New Finencial Year..."
    Status.Refresh
    
    ''''' TRANSACTION FILES DELETION
    NEW_CONNECTION.Execute "delete from ledger Where V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & " "
    NEW_CONNECTION.Execute "delete from ledgerM Where V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & " "
    NEW_CONNECTION.Execute "delete from ledgerRef Where V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & " "
    NEW_CONNECTION.Execute "delete from ledgerTDS Where V_Date < " & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & " "
    NEW_CONNECTION.Execute "delete from TDSCerti "
    NEW_CONNECTION.Execute "delete from TDSChal"
    NEW_CONNECTION.Execute "delete from TDSChal1"
    NEW_CONNECTION.Execute "delete from LastVoucher"
    Status = "Updating Opening Balances For New Finencial Year..."
    Status.Refresh
    mCode = ""
    mVNo = 0
    NEW_CONNECTION.CommitTrans
    
    NEW_CONNECTION.BeginTrans
    Do Until RST1.EOF
        If mCode <> RST1!mCode Then
            mVNo = mVNo + 1
            mDocId = FaSetW(PubDivCode, 1) + PubSiteCode + PubSiteCode + FaSetW("F_AO", 5) + FaSetW(Trim(STR(Year(DateAdd("YYYY", 1, PubStartDate)))), 5) + FaSetN(STR(mVNo), 8)
            mCode = RST1!mCode
            mS_NO = 1
            NEW_CONNECTION.Execute ("INSERT INTO LEDGERM (DocId,V_Type,v_Prefix,V_No,Site_Code,V_Date,Narration,U_Name,U_EntDt,U_AE) VALUES (" & FaChk_Text(mDocId) & ",'F_AO'," & FaChk_Text(Trim(STR(Year(DateAdd("YYYY", 1, PubStartDate))))) & "," & mVNo & "," & FaChk_Text(PubSiteCode) & "," & FaConvertDate(DateAdd("d", -1, DateAdd("YYYY", 1, PubStartDate))) & ",'Opening Balance'," & FaChk_Text(pubUName) & "," & FaConvertDate(Now) & ",'A')")
        End If
        If RST1!Balance < 0 Then       'DEBIT
            NEW_CONNECTION.Execute ("INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,U_Name,U_EntDt,U_AE) VALUES (" & FaChk_Text(mDocId) & "," & mS_NO & ",'F_AO'," & mVNo & "," & FaChk_Text(Trim(STR(Year(DateAdd("YYYY", 1, PubStartDate))))) & "," & FaChk_Text(PubSiteCode) & "," & FaConvertDate(DateAdd("d", -1, DateAdd("YYYY", 1, PubStartDate))) & "," & FaChk_Text(RST1!Party_code) & ",0," & Abs(RST1!Balance) & "," & FaChk_Text(pubUName) & "," & FaConvertDate(Now) & ",'A')")
            If PubBackEnd = "A" Then
                FaCalCurrBal NEW_CONNECTION, RST1!Party_code, Abs(RST1!Balance), 0
                FaCalCurrBal GCn, RST1!Party_code, Abs(RST1!Balance), 0
            End If
        ElseIf RST1!Balance > 0 Then     'CREDIT
            NEW_CONNECTION.Execute ("INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,U_Name,U_EntDt,U_AE) VALUES (" & FaChk_Text(mDocId) & "," & mS_NO & ",'F_AO'," & mVNo & "," & FaChk_Text(Trim(STR(Year(DateAdd("YYYY", 1, PubStartDate))))) & "," & FaChk_Text(PubSiteCode) & "," & FaConvertDate(DateAdd("d", -1, DateAdd("YYYY", 1, PubStartDate))) & "," & FaChk_Text(RST1!Party_code) & "," & Abs(RST1!Balance) & ",0," & FaChk_Text(pubUName) & "," & FaConvertDate(Now) & ",'A')")
            If PubBackEnd = "A" Then
                FaCalCurrBal NEW_CONNECTION, RST1!Party_code, 0, Abs(RST1!Balance)
                FaCalCurrBal GCn, RST1!Party_code, 0, Abs(RST1!Balance)
            End If
        End If
        mS_NO = mS_NO + 1
        RST1.MoveNext
    Loop
    'UPDATING SUBGROUP TABLE
        If PubBackEnd = "A" Then NEW_CONNECTION.Execute ("Update SubGroup Set FirmCode='" & Right(FolderName, 1) & "'")
        GCn.Execute ("Update SubGroup Set FirmCode='" & Right(FolderName, 1) & "'")

    'UPDATING SUBGROUPALIAS TABLE
        If PubBackEnd = "A" Then NEW_CONNECTION.Execute ("Update SubGroupAlias Set FirmCode='" & Right(FolderName, 1) & "'")
        GCn.Execute ("Update SubGroup Set FirmCode='" & Right(FolderName, 1) & "'")
        
    'UPDATING VEH_BILLCOUNTER TABLE
        NEW_CONNECTION.Execute ("Update VehBill_Counter set Start_Srl_No=90000000,Date_From=" & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & "")
        
    'UPDATING VOUCHER_PREFIX TABLE
       
        NEW_CONNECTION.Execute ("Update Voucher_Prefix set Date_From =" & ConvertDate(DateAdd("yyyy", 1, PubStartDate)) & ",Date_To=" & ConvertDate(DateAdd("yyyy", 1, PubEndDate)) & ", Start_Srl_No= 1200000")
        NEW_CONNECTION.Execute ("Update Voucher_Type set Category='OPBAL',NCAT='OPBAL' where V_Type='F_AO'")
        
        
    If PubBackEnd = "A" Then
        Set RstParty = NEW_CONNECTION.Execute("Select SubCode from Subgroup where nature in ('Customer','Suplier')")
        If RstParty.RecordCount > 0 Then
            RstParty.MoveFirst
            For I = 1 To RstParty.RecordCount
                    mDelete = True
                   '* For FA Data Delete..
                    If NEW_CONNECTION.Execute("Select * From Ledger Where SubCode='" & RstParty!SubCode & "'").RecordCount > 0 Then
                        mDelete = False
                    End If
                    '* For Vehicle Data Delete..
                    If GCn.Execute("Select * From Veh_Purch1 Where PartyCode='" & RstParty!SubCode & "'").RecordCount > 0 Then
                        mDelete = False
                    End If
                    
                    If GCn.Execute("Select * From Veh_Order Where PartyCode='" & RstParty!SubCode & "'").RecordCount > 0 Then
                        mDelete = False
                    End If
                    
                    '* For Spare Data Delete..
                    If GCn.Execute("Select * From SP_Sale Where Party_Code='" & RstParty!SubCode & "'").RecordCount > 0 Then
                        mDelete = False
                    End If
                    
                    If GCn.Execute("Select * From SP_Purch Where Party_Code='" & RstParty!SubCode & "'").RecordCount > 0 Then
                        mDelete = False
                    End If
                    
                    If GCn.Execute("Select * From SP_Order Where Party_Code='" & RstParty!SubCode & "'").RecordCount > 0 Then
                        mDelete = False
                    End If
                    
                    If GCn.Execute("Select * From RECT Where PartyCode='" & RstParty!SubCode & "'").RecordCount > 0 Then
                        mDelete = False
                    End If
                    
                    If GCn.Execute("Select * From Rect Where AcCode='" & RstParty!SubCode & "'").RecordCount > 0 Then
                        mDelete = False
                    End If
                    
                    If mDelete = True Then
                         mTrans = True
                            GCn.Execute ("Delete From SubGroupAlias Where SubCode='" & RstParty!SubCode & "'")
                            GCn.Execute ("Delete From SubGroup Where SubCode='" & RstParty!SubCode & "'")
                
                            NEW_CONNECTION.Execute ("Delete From SubGroupAlias Where SubCode='" & RstParty!SubCode & "'")
                            NEW_CONNECTION.Execute ("Delete From SubGroup Where SubCode='" & RstParty!SubCode & "'")
                    End If
                    RstParty.MoveNext
            Next
        End If
        Set RstParty = Nothing
    End If
    ProgressBar1.Value = 96
    Status.Refresh
    Screen.MousePointer = vbDefault
    NEW_CONNECTION.CommitTrans
    mBeginTrans = 0
    Set RST1 = Nothing
    Set NEW_CONNECTION = Nothing
    Status = "Updation Complete..."
    
    GCn.CommitTrans

ProgressBar1.Value = 100
Else
End
End If


MsgBox "All Data has been updated to the new Year.Please reload the Software", vbInformation + vbOKOnly, "Year End"
End
End Sub

Private Function PValue(Val As Integer)
'    ProgressBar1.Value = 0
'    ProgressBar2.Value = 0
'    ProgressBar3.Value = 0
'    ProgressBar1.Value = val
'    ProgressBar2.Value = val - 5
'    ProgressBar3.Value = 1.5
End Function
Private Function ImageDisp(ImgNo As Integer)
    Image1(SearchDir).Visible = False
    Image1(CreateDir).Visible = False
    Image1(TransAuto).Visible = False
    Image1(TransFA).Visible = False
    
        Select Case ImgNo
            Case 0
                Image1(SearchDir).Visible = True
            Case 1
                Image1(CreateDir).Visible = True
            Case 2
                Image1(TransAuto).Visible = True
            Case 3
                Image1(TransFA).Visible = True
            Case 4
                Image1(CompProcess).Visible = True
        End Select
    If ImgNo <> 0 Then
        Process(ImgNo - 1).ForeColor = vbBlue
        Image2(ImgNo - 1).Visible = True
    End If
    Image1(ImgNo).Refresh
End Function
Private Sub X_Val1(ByRef Temp06 As ADODB.Recordset, ByRef TRec1 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec1.RecordCount <= 0 Or TRec1.EOF = True Or TRec1.BOF = True Then
        xRate = 0
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
        Exit Sub
    End If
    If xQty = TRec1!Qty Then
        TRec1.Fields("QTY") = 0
        TRec1.Update
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
        TRec1.MoveNext
    ElseIf xQty < TRec1!Qty Then
        TRec1.Fields("QTY") = TRec1!Qty - xQty
        TRec1.Update
        
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
    ElseIf xQty > TRec1!Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec1.EOF
            If TRec1!Part_No <> RstPart!Part_No Then
                GoTo MyNextRecord
            End If
            If TRec1!Qty <= TQty Then
                TQty = TQty - TRec1!Qty
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TRec1!Qty
                mOP_TB_VAL = mOP_TB_VAL - (TRec1!Qty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                        .Fields("Part_Name") = mname
                        .Fields("Narr") = left(xNARR, 25)
                        .Fields("Inv_No") = mInv_No
                        .Fields("Inv_Date") = mInv_Date
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = TRec1!Qty
                        .Fields("Tb_Val") = TRec1!Qty * xRate
                        .Fields("Tb_BQty") = mOP_TB_QTY
                        .Fields("Tb_BVal") = mOP_TB_VAL
                        
                        .Fields("Is_Tp") = 0
                        .Fields("Tp_Val") = 0
                        .Fields("Tp_BQty") = 0
                        .Fields("Tp_BVal") = 0
                        .Update
                    End With
                    TRec1.Fields("QTY") = 0
                    TRec1.Update
                End If
            Else
                TRec1.Fields("QTY") = TRec1!Qty - TQty
                TRec1.Update
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = TQty
                        .Fields("Tb_Val") = TQty * xRate
                        .Fields("Tb_BQty") = mOP_TB_QTY
                        .Fields("Tb_BVal") = mOP_TB_VAL
                        
                        .Fields("Is_Tp") = 0
                        .Fields("Tp_Val") = 0
                        .Fields("Tp_BQty") = 0
                        .Fields("Tp_BVal") = 0
                        .Update
                    End With
                    TQty = 0
                    Exit Do
                End If
            End If
MyNextRecord:
            TRec1.MoveNext
            If TRec1.EOF = True And TQty <> 0 Then
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                If mPART_ADD = False Then
                    mPART_ADD = True
                    With Temp06
                        .AddNew
                        .Fields("Part_Name") = RstPart!Part_Name
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Job_Age") = "Y"
                        .Update
                    End With
                End If
                With Temp06
                    .AddNew
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Rate") = xRate
                    
                    .Fields("Is_Tb") = TQty
                    .Fields("Tb_Val") = TQty * xRate
                    .Fields("Tb_BQty") = mOP_TB_QTY
                    .Fields("Tb_BVal") = mOP_TB_VAL
                    
                    .Fields("Is_Tp") = 0
                    .Fields("Tp_Val") = 0
                    .Fields("Tp_BQty") = 0
                    .Fields("Tp_BVal") = 0
                    .Update
                End With
            
            End If
        Loop
    End If
End Sub

Private Sub X_Val2(ByRef Temp06 As ADODB.Recordset, ByRef TRec2 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec2.RecordCount <= 0 Or TRec2.EOF = True Or TRec2.BOF = True Then
        xRate = 0
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = PrinID(RstStock!DocID)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                
                .Update
            End With
        End If
        Exit Sub
    End If
    
    If xQty = TRec2!Qty Then
        TRec2.Fields("QTY") = 0
        TRec2.Update
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                
                .Update
            End With
        End If
        TRec2.MoveNext
    ElseIf xQty < TRec2!Qty Then
        TRec2.Fields("QTY") = TRec2!Qty - xQty
        TRec2.Update
        
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                .Update
            End With
        End If
    ElseIf xQty > TRec2!Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec2.EOF
            If TRec2!Part_No <> RstPart!Part_No Then
                GoTo MyNextRecord
            End If
            If TRec2!Qty <= TQty Then
                TQty = TQty - TRec2!Qty
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TRec2!Qty
                mOP_TP_VAL = mOP_TP_VAL - (TRec2!Qty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                        .Fields("Part_Name") = mname
                        .Fields("Narr") = xNARR
                        .Fields("Inv_No") = mInv_No
                        .Fields("Inv_Date") = mInv_Date
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = 0
                        .Fields("Tb_Val") = 0
                        .Fields("Tb_BQty") = 0
                        .Fields("Tb_BVal") = 0
                        
                        .Fields("Is_Tp") = TRec2!Qty
                        .Fields("Tp_Val") = TRec2!Qty * xRate
                        .Fields("Tp_BQty") = mOP_TP_QTY
                        .Fields("Tp_BVal") = mOP_TP_VAL
                        .Update
                    End With
                    TRec2.Fields("QTY") = 0
                    TRec2.Update
                End If
            Else
                TRec2.Fields("QTY") = TRec2!Qty - TQty
                TRec2.Update
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = 0
                        .Fields("Tb_Val") = 0
                        .Fields("Tb_BQty") = 0
                        .Fields("Tb_BVal") = 0
                        
                        .Fields("Is_Tp") = TQty
                        .Fields("Tp_Val") = TQty * xRate
                        .Fields("Tp_BQty") = mOP_TP_QTY
                        .Fields("Tp_BVal") = mOP_TP_VAL
                        .Update
                    End With
                    TQty = 0
                    Exit Do
                End If
            End If
MyNextRecord:
            TRec2.MoveNext
            If TRec2.EOF = True And TQty <> 0 Then
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                If mPART_ADD = False Then
                    mPART_ADD = True
                    With Temp06
                        .AddNew
                        .Fields("Part_Name") = RstPart!Part_Name
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Job_Age") = "Y"
                        .Update
                    End With
                End If
                With Temp06
                    .AddNew
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Rate") = xRate
                    
                    .Fields("Is_Tb") = 0
                    .Fields("Tb_Val") = 0
                    .Fields("Tb_BQty") = 0
                    .Fields("Tb_BVal") = 0
                    
                    .Fields("Is_Tp") = TQty
                    .Fields("Tp_Val") = TQty * xRate
                    .Fields("Tp_BQty") = mOP_TP_QTY
                    .Fields("Tp_BVal") = mOP_TP_VAL
                    .Update
                End With
            End If
        Loop
    End If
End Sub
Private Sub X_VAL11(ByRef TRec1 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
On Error GoTo ErrLoop
    If TRec1.RecordCount <= 0 Or TRec1.EOF = True Or TRec1.BOF = True Then
        If mOP_TB_VAL <> 0 And mOP_TB_QTY <> 0 Then
            xRate = Round(mOP_TB_VAL / mOP_TB_QTY, 3)
        Else
            xRate = 0
        End If
            mOP_TB_QTY = mOP_TB_QTY - xQty
            mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
            mIss_TB_Qty = mIss_TB_Qty + xQty
            mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
          Exit Sub
    End If
    If xQty = TRec1Qty Then
        TRec1Qty = 0
'        TRec1!Qty = 0
'        TRec1.Update
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        mIss_TB_Qty = mIss_TB_Qty + xQty
        mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
        TRec1.MoveNext
        If TRec1.EOF = False Then
            TRec1Qty = TRec1!Qty
        End If
'    ElseIf xQty < TRec1!Qty Then
    ElseIf xQty < TRec1Qty Then
        TRec1Qty = TRec1Qty - xQty
'        TRec1!Qty = TRec1!Qty - xQty
'        TRec1.Update
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        mIss_TB_Qty = mIss_TB_Qty + xQty
        mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
'    ElseIf xQty  > TRec1!Qty Then
    ElseIf xQty > TRec1Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec1.EOF
'            If TRec1!Qty <= TQty Then
            If TRec1Qty <= TQty Then
'                TQty = TQty - TRec1!Qty
                TQty = TQty - TRec1Qty
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TRec1Qty 'TRec1!Qty
                mOP_TB_VAL = mOP_TB_VAL - (TRec1Qty * xRate) '(TRec1!Qty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + (TRec1Qty) '(TRec1!Qty)
                mIss_TB_Val = mIss_TB_Val + (TRec1Qty * xRate) '(TRec1!Qty * xRate)
                TRec1Qty = 0
'                TRec1!Qty = 0
'                TRec1.Update
            Else
                TRec1Qty = TRec1Qty - TQty
'                TRec1!Qty = TRec1!Qty - TQty
'                TRec1.Update
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + TQty
                mIss_TB_Val = mIss_TB_Val + (TQty * xRate)
                TQty = 0
                Exit Do
            End If
            TRec1.MoveNext
            If TRec1.EOF = True And TQty <> 0 Then
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + TQty
                mIss_TB_Val = mIss_TB_Val + (TQty * xRate)
            End If
            If TRec1.EOF = False Then
                TRec1Qty = TRec1!Qty
            End If
        Loop
    End If
ErrLoop:
     If err.NUMBER <> 0 Then CheckError
End Sub

Private Sub X_VAL22(ByRef TRec2 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec2.RecordCount <= 0 Or TRec2.EOF = True Or TRec2.BOF = True Then
        xRate = 0
        
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
        Exit Sub
    End If
'    If xQty = TRec2!Qty Then
    If xQty = TRec2Qty Then
        TRec2Qty = 0
'        TRec2!Qty = 0
'        TRec2.Update
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
        TRec2.MoveNext
        If TRec2.EOF = False Then
            TRec2Qty = TRec2!Qty
        End If
'    ElseIf xQty < TRec2!Qty Then
    ElseIf xQty < TRec2Qty Then
        TRec2Qty = TRec2Qty - xQty
'        TRec2!Qty = TRec2!Qty - xQty
'        TRec2.Update
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
'    ElseIf xQty  > TRec2!Qty Then
    ElseIf xQty > TRec2Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec2.EOF
'            If TRec2!Qty <= TQty Then
            If TRec2Qty <= TQty Then
                TQty = TQty - TRec2Qty 'TRec2!Qty
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TRec2Qty 'TRec2!Qty
                mOP_TP_VAL = mOP_TP_VAL - (TRec2Qty * xRate)   '(TRec2!Qty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + (TRec2Qty)     '(TRec2!Qty)
                mIss_TP_Val = mIss_TP_Val + (TRec2Qty * xRate) '(TRec2!Qty * xRate)
                TRec2Qty = 0
'                TRec2!Qty = 0
'                TRec2.Update
            Else
                TRec2Qty = TRec2Qty - TQty
'                TRec2!Qty = TRec2!Qty - TQty
'                TRec2.Update
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + TQty
                mIss_TP_Val = mIss_TP_Val + (TQty * xRate)
                TQty = 0
                Exit Do
            End If
            TRec2.MoveNext
            If TRec2.EOF = True And TQty <> 0 Then
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + TQty
                mIss_TP_Val = mIss_TP_Val + (TQty * xRate)
            End If
            If TRec2.EOF = False Then
                TRec2Qty = TRec2!Qty
            End If
        Loop
    End If
End Sub

