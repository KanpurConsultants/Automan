VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmVehiclePriceList 
   Caption         =   "Import Ledger Accounts & Opening Balance"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Select Excel File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   120
      TabIndex        =   5
      Top             =   2100
      Width           =   5340
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
         Height          =   435
         Left            =   150
         TabIndex        =   8
         Top             =   930
         Width           =   1005
      End
      Begin VB.CommandButton CmdFileSelect 
         Caption         =   "..."
         Height          =   390
         Left            =   4710
         TabIndex        =   7
         Top             =   435
         Width           =   465
      End
      Begin VB.TextBox Text1 
         Height          =   390
         Left            =   150
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   435
         Width           =   4530
      End
      Begin MSComctlLib.ProgressBar Prg 
         Height          =   435
         Left            =   1170
         TabIndex        =   9
         Top             =   930
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   767
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Excel File Format"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   5325
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   4095
         TabIndex        =   12
         Top             =   915
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Column Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   165
         TabIndex        =   11
         Top             =   870
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sheet Name :  Sheet1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   1440
         TabIndex        =   10
         Top             =   405
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number"
         Height          =   240
         Index           =   10
         Left            =   4410
         TabIndex        =   4
         Top             =   1455
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text (25)"
         Height          =   240
         Index           =   9
         Left            =   4290
         TabIndex        =   3
         Top             =   1170
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "2. SaleRate"
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   1455
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "1. AutomanModel"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   1170
         Width           =   1800
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5355
      Top             =   5805
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "Spare Debtor Group........."
   End
   Begin VB.Label LBL 
      Caption         =   "Label2"
      Height          =   165
      Left            =   465
      TabIndex        =   13
      Top             =   6120
      Width           =   960
   End
End
Attribute VB_Name = "FrmVehiclePriceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mFileName$
Dim mFileTitle$


Private Sub CmdFileSelect_Click()
    CD1.FileName = ""
    
    Call SelectFile
    mFileName = CD1.FileName
    Text1 = CD1.FileName
    mFileTitle = CD1.FileTitle
    If mFileName = "" Then Exit Sub
    mFileTitle = mID(mFileTitle, 1, Len(mFileTitle) - 4)
           
End Sub


Private Sub SelectFile()
    
    CD1.CancelError = False
    CD1.DialogTitle = "Select CrmDms Excel Files"
    CD1.Filter = "Excel Files (*.xls)|*.xls"
    CD1.FilterIndex = 1
    CD1.Flags = cdlOFNHideReadOnly
    CD1.ShowOpen
    
End Sub


Private Sub CmdImport_Click()
    Dim X As Long
    Dim RsTemp          As ADODB.Recordset
    Dim RsDms           As ADODB.Recordset
    Dim DmsConn As ADODB.Connection

    Dim mSubGroupCounter    As Long
    Dim mSubCode$, mDmsSubCode$, mQry$, mNarr$, mLocalCentral$, mCondStr$
    
    Dim mDocId$, mVType$, mVPrefix$, mVNo$, xDocId$, mVSNo%
    Dim mTrans As Boolean
    
On Error GoTo DispErr
                    
    
    mSubGroupCounter = G_CompCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0)
    
    If mFileName = "" Then Exit Sub
    mFileTitle = mID(mFileTitle, 1, Len(mFileTitle) - 4)
    Set DmsConn = New Connection
    DmsConn.CursorLocation = adUseClient
    DmsConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mFileName & ";Extended Properties=Excel 8.0"
    
    DmsConn.Execute "Create Table ModelNotFound (Model Text(255)) "
    
    'Set RsDms = DmsConn.Execute("Select * from [" & mFileTitle & "$]")
    Set RsDms = DmsConn.Execute("Select * from [Sheet1$]")
    
    
    For X = 1 To 9999
       ' LBL.Refresh
    Next X
    
    
    With RsDms
        If ChkFieldExist(RsDms, "AutomanModel") And ChkFieldExist(RsDms, "SaleRate") Then
            If .RecordCount > 0 Then
                Prg.Value = 0
                Prg.Visible = True
                Do Until .EOF
                    GCn.BeginTrans
                    G_FaCn.BeginTrans
                    mTrans = True
                        
                        If GCn.Execute("Select Count(*) From Model Where Model = '" & XNull(!AutomanModel) & "'").Fields(0).Value > 0 Then
                            GCn.Execute "Update Model Set Sale_Rate = " & VNull(!SaleRate) & " Where Model = '" & XNull(!AutomanModel) & "' And Div_Code = '" & PubDivCode & "' "
                        Else
                            DmsConn.Execute "Insert Into ModelNotFound Values('" & XNull(!AutomanModel) & "')"
                        End If

                        

                        If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                        .MoveNext
                    GCn.CommitTrans
                    G_FaCn.CommitTrans
                    mTrans = False
                Loop
            End If
        End If
    End With
        
    MsgBox "Price List Completed"
        
    'If ChkAllErr.Value = 0 Then mCondStr = " Where U_EntDt = " & ConvertDate(PubLoginDate) & ""
'    Set RsTemp = GCn.Execute("Select Cat As Category, [Key] as Dms_Reference, Narration From DmsErrLog " & mCondStr)
'    Set FgridErr.DataSource = RsTemp
'    Ini_Grid
        
    Set RsTemp = Nothing
    Set RsDms = Nothing
    If DmsConn.State <> 0 Then DmsConn.Close
Exit Sub
DispErr:
    MsgBox err.Description
    If mTrans Then G_FaCn.RollbackTrans
    If mTrans Then GCn.RollbackTrans
    If StrCmp(err.Description, "Table 'ModelNotFound' already exists.") Then Resume Next
    Set RsDms = Nothing
    Set RsTemp = Nothing
    If DmsConn.State <> 0 Then DmsConn.Close
    
End Sub





