VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmUpdateAccountOpening 
   Caption         =   "Import Ledger Accounts & Opening Balance"
   ClientHeight    =   6195
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
   ScaleHeight     =   6195
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
      Left            =   150
      TabIndex        =   19
      Top             =   4380
      Width           =   5295
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
         Height          =   435
         Left            =   150
         TabIndex        =   22
         Top             =   930
         Width           =   1005
      End
      Begin VB.CommandButton CmdFileSelect 
         Caption         =   "..."
         Height          =   390
         Left            =   4710
         TabIndex        =   21
         Top             =   435
         Width           =   465
      End
      Begin VB.TextBox Text1 
         Height          =   390
         Left            =   150
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   435
         Width           =   4530
      End
      Begin MSComctlLib.ProgressBar Prg 
         Height          =   435
         Left            =   1170
         TabIndex        =   23
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
      Height          =   4245
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   5325
      Begin VB.Label Label1 
         Caption         =   "8.TIN"
         Height          =   315
         Index           =   24
         Left            =   135
         TabIndex        =   28
         Top             =   3150
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text(12)"
         Height          =   240
         Index           =   22
         Left            =   2475
         TabIndex        =   27
         Top             =   3150
         Width           =   735
      End
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
         Left            =   2490
         TabIndex        =   26
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
         TabIndex        =   25
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
         Left            =   960
         TabIndex        =   24
         Top             =   480
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number"
         Height          =   240
         Index           =   17
         Left            =   2490
         TabIndex        =   18
         Top             =   3690
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number"
         Height          =   240
         Index           =   16
         Left            =   2490
         TabIndex        =   17
         Top             =   3405
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text(35)"
         Height          =   240
         Index           =   15
         Left            =   2490
         TabIndex        =   16
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text (40)"
         Height          =   240
         Index           =   14
         Left            =   2490
         TabIndex        =   15
         Top             =   2310
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text (40)"
         Height          =   240
         Index           =   13
         Left            =   2490
         TabIndex        =   14
         Top             =   2025
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text(24)"
         Height          =   240
         Index           =   12
         Left            =   2490
         TabIndex        =   13
         Top             =   2595
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text (40)"
         Height          =   240
         Index           =   11
         Left            =   2490
         TabIndex        =   12
         Top             =   1740
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer / Supplier"
         Height          =   240
         Index           =   10
         Left            =   2490
         TabIndex        =   11
         Top             =   1455
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text (40)"
         Height          =   240
         Index           =   9
         Left            =   2490
         TabIndex        =   10
         Top             =   1170
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "9. Credit"
         Height          =   315
         Index           =   8
         Left            =   150
         TabIndex        =   9
         Top             =   3690
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "8. Debit"
         Height          =   315
         Index           =   7
         Left            =   150
         TabIndex        =   8
         Top             =   3405
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "7. Phone"
         Height          =   315
         Index           =   6
         Left            =   150
         TabIndex        =   7
         Top             =   2880
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "5. City"
         Height          =   315
         Index           =   5
         Left            =   150
         TabIndex        =   6
         Top             =   2310
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "4. Address2"
         Height          =   315
         Index           =   4
         Left            =   150
         TabIndex        =   5
         Top             =   2025
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "6. Mobile"
         Height          =   315
         Index           =   3
         Left            =   150
         TabIndex        =   4
         Top             =   2595
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "3. Address1"
         Height          =   315
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   1740
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "2. Nature"
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   1455
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "1. PartyName"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   1170
         Width           =   1365
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
      TabIndex        =   29
      Top             =   6120
      Width           =   960
   End
End
Attribute VB_Name = "FrmUpdateAccountOpening"
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
    
On Error GoTo DispErr
                    
    
    mSubGroupCounter = G_CompCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0)
    
    If mFileName = "" Then Exit Sub
    mFileTitle = mID(mFileTitle, 1, Len(mFileTitle) - 4)
    Set DmsConn = New Connection
    DmsConn.CursorLocation = adUseClient
    DmsConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mFileName & ";Extended Properties=Excel 8.0"
    
    'Set RsDms = DmsConn.Execute("Select * from [" & mFileTitle & "$]")
    Set RsDms = DmsConn.Execute("Select * from [Sheet1$]")
    
    For X = 1 To 9999
       ' LBL.Refresh
    Next X
    
    
    With RsDms
        If ChkFieldExist(RsDms, "PartyName") And ChkFieldExist(RsDms, "Nature") And _
           ChkFieldExist(RsDms, "Debit") And ChkFieldExist(RsDms, "Address1") And ChkFieldExist(RsDms, "Address2") And ChkFieldExist(RsDms, "City") And ChkFieldExist(RsDms, "Mobile") And ChkFieldExist(RsDms, "Phone") And ChkFieldExist(RsDms, "Credit") And ChkFieldExist(RsDms, "TIN") Then
           
            If .RecordCount > 0 Then
                Prg.Value = 0
                Prg.Visible = True
                Do Until .EOF
                    GCn.BeginTrans
                    G_FaCn.BeginTrans

                        If InStr(1, XNull(!Nature), "Customer") > 0 Then
                            Call AutomanSubcode(XNull(!PartyName), "0020", "Customer", False, XNull(!Address1), XNull(!Address2), XNull(!City), XNull(!Mobile), XNull(!Phone), VNull(!Debit), VNull(!Credit), XNull(!TIN))
                        ElseIf InStr(1, XNull(!Nature), "Supplier") > 0 Then
                            Call AutomanSubcode(XNull(!PartyName), "0016", "Supplier", False, XNull(!Address1), XNull(!Address2), XNull(!City), XNull(!Mobile), XNull(!Phone), VNull(!Debit), VNull(!Credit), XNull(!TIN))
                        End If
                        

                        If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                        .MoveNext
                    GCn.CommitTrans
                    G_FaCn.CommitTrans
                Loop
            End If
        End If
    End With
        
    MsgBox "Import Process Completed"
        
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
    G_FaCn.RollbackTrans
    GCn.RollbackTrans
    Set RsDms = Nothing
    Set RsTemp = Nothing
    If DmsConn.State <> 0 Then DmsConn.Close

End Sub


Function AutomanSubcode(mDmsSubCode As String, mAutomanGroupCode As String, mNature As String, Optional mAskForOpeningAccounts As Boolean = True, Optional mAddress1 As String = "", Optional mAddress2 As String = "", Optional mCity As String = "", Optional mMobile As String = "", Optional mPhone As String = "", Optional mDebit As Double = 0, Optional mCredit As Double = 0, Optional mTin As String = "") As String
    Dim mConn As New ADODB.Connection
    Dim RsTemp As ADODB.Recordset
    Dim rsTemp1 As ADODB.Recordset
    Dim RsTempCity As ADODB.Recordset
    Dim mSubGroupCounter As Long
    Dim mSubCode$, mQry$, mname$, mCityCode$, mStateCode$
    Dim mLocalCentral$
    
            Dim mVType As String
            Dim mVPrefix As String
            Dim mVNo As Double
            Dim mDocId As String
            Dim mVSNo As Integer
            Dim mVDate As Date
    
    mDmsSubCode = left(mDmsSubCode, 40)
    
    
    
    mConn.CursorLocation = adUseClient
    If PubDbUser <> "" Then
        mConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
    Else
        mConn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
    End If
    mConn.Open
    
    

        Set RsTemp = mConn.Execute("Select SubCode From SubGroup With(NOLOCK) Where Name = '" & mDmsSubCode & "' ")
        
        If RsTemp.RecordCount > 0 Then
            AutomanSubcode = RsTemp!SubCode
        Else
            If mAskForOpeningAccounts Then If MsgBox("Account Name " & mDmsSubCode & "Doesn't exist in Automan. Do u want to Create it?", vbYesNo) = vbNo Then Exit Function
            mSubGroupCounter = G_CompCn.Execute("Select SubGroupAcCode From SubGroupCounter With (NOLOCK)").Fields(0)
                        
            
            mSubCode = PubSiteCode & PubFirmCode & Format(mSubGroupCounter, "000000")
            mname = mDmsSubCode
                    
            mQry = "Insert Into SubGroup (AcId, Site_Code, SubCode, FirmCode, NamePrefix, " & _
                                        "Name, NameHelp, GroupCode, Nature, Add1, " & _
                                        "Add2,  CityCode, Phone, Mobile, Email, " & _
                                        "CstNo, LstNo, ActiveYn, U_Name, " & _
                                        "U_EntDt, U_AE, GroupNature, AliasYn) " & _
                 " Values ('" & mSubCode & "', " & PubSiteCode & ", '" & mSubCode & "', " & PubFirmCode & ", '', " & _
                 "'" & mname & "', '" & mname & "', '" & mAutomanGroupCode & "', '" & mNature & "', " & _
                 "'" & mAddress1 & "', '" & mAddress2 & "', '" & mCityCode & "', '" & mPhone & "', '" & mMobile & "', '', " & _
                 "'', '" & mTin & "', 1, '" & pubUName & " ', " & _
                 "" & ConvertDate(PubLoginDate) & ", 'A', 'A', 'N')"
                         
                         
                         
            GCn.Execute mQry
            If PubBackEnd = "A" Then G_FaCn.Execute mQry
            
            
            
            G_CompCn.Execute ("Update  SubGroupCounter Set SubGroupAcCode=" & mSubGroupCounter + 1 & " ")
            
            AutomanSubcode = mSubCode
        End If

        If mDebit > 0 Or mCredit > 0 Then
            mVType = "F_AO"
            mVPrefix = "AUTO"
            mVNo = GCn.Execute("Select IsNull(Max(V_No),0)+1 From Ledger With (NoLock) Where V_Type='F_AO' And V_Prefix='Auto'").Fields(0).Value
            mDocId = PubDivCode & PubSiteCode & PubSiteCode & mVType & Space(5 - Len(mVType)) & mVPrefix & Space(5 - Len(mVPrefix)) & Space(8 - Len(mVNo)) & mVNo
            mVSNo = 1
            mVDate = DateAdd("D", -1, PubStartDate)
            
            
            
            GCn.Execute "Delete from Ledger Where Subcode = '" & mSubCode & "' And V_Type = '" & mVType & "'"
'            GCn.Execute "IF Not Exists (Select * From dbo.LedgerM Where DocId = '" & mDocId & "') INSERT INTO dbo.LedgerM (   DocId,  Site_Code,  V_Type, v_Prefix,   V_No,   V_Date, Narration,  U_Name, U_EntDt,    U_AE,      DmsRefNo    )" & _
'                        "VALUES  ('" & mDocId & "', '" & PubSiteCode & "', '" & mVType & "',   '" & mVPrefix & "',  " & Val(mVNo) & ",  " & ConvertDate(mVDate) & ",    '" & XNull(mDmsSubCode) & "', '" & pubUName & "',   '" & PubLoginDate & "',   'A',  ''   )"
            GCn.Execute "INSERT INTO dbo.Ledger  (DocId,  V_SNO,  V_Type, V_No,   v_Prefix,   Site_Code,  V_Date, SubCode,    ContraSub,  AmtDr,  AmtCr,  U_Name, U_EntDt, U_AE)" & _
                        "VALUES  ('" & mDocId & "', " & mVSNo & ", '" & mVType & "',   " & Val(mVNo) & ",  '" & mVPrefix & "',  '" & PubSiteCode & "', " & ConvertDate(mVDate) & ",  '" & mSubCode & "',   '', " & VNull(mDebit) & ", " & VNull(mCredit) & ", 'SA',  '" & PubLoginDate & "',   'A')"
        End If



    Set RsTemp = Nothing
    Set rsTemp1 = Nothing
End Function


