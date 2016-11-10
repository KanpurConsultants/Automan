VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmTallyImport 
   Caption         =   "Tally Import"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Select Ms-Excel File..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   3165
      Begin VB.CommandButton CmdImport 
         Caption         =   "Voucher"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   0
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   285
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   1
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   285
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Supplier Payment"
         Height          =   525
         Index           =   4
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4200
         Width           =   1320
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Spare Sale Return"
         Height          =   540
         Index           =   7
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4200
         Width           =   1425
      End
      Begin MSComctlLib.ProgressBar Prg 
         Height          =   270
         Left            =   165
         TabIndex        =   10
         Top             =   1365
         Visible         =   0   'False
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Error Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   11625
      Begin VB.CheckBox ChkAllErr 
         BackColor       =   &H00CFE0E0&
         Caption         =   "All Types"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6150
         TabIndex        =   5
         Top             =   15
         Width           =   1170
      End
      Begin VB.CommandButton CmdDelErr 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   10245
         TabIndex        =   4
         Top             =   2565
         Width           =   1185
      End
      Begin VB.CommandButton CmdDelErr 
         Caption         =   "Show All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   9045
         TabIndex        =   3
         Top             =   2565
         Width           =   1185
      End
      Begin VB.TextBox TxtShow 
         Appearance      =   0  'Flat
         Height          =   915
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "FrmTallyImport.frx":0000
         Top             =   1995
         Width           =   8865
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FgridErr 
         Height          =   1620
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   2858
         _Version        =   393216
         BackColorFixed  =   13623520
         BackColorBkg    =   13623520
         AllowUserResizing=   3
         Appearance      =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   3  'Align Left
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   661
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "Spare Debtor Group........."
   End
   Begin VB.Label LBL 
      Caption         =   "LBL"
      Height          =   300
      Left            =   1695
      TabIndex        =   13
      Top             =   5430
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "FrmTallyImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BtnVoucher As Byte = 0
Private Const BtnSubGroup As Byte = 1


'FGridErr Constants
Const FErr_Cat          As Byte = 1
Const FErr_DmsRef       As Byte = 2
Const FErr_Narration    As Byte = 3


Dim DmsConn As ADODB.Connection



Private Function CreateVType(mVType As String, mStartSrlNo As Double, mDivCode As String) As String
Dim VTypeAlreadyExist As Boolean
Dim ManualVType As String
Dim TmpRst As ADODB.Recordset
    Set TmpRst = G_FaCn.Execute("Select V_Type from Voucher_type where Description='" & mVType & "'")
    If TmpRst.RecordCount = 0 Then
        Do Until VTypeAlreadyExist
            If ManualVType = "" Then
                ManualVType = InputBox(mVType & " Does Not Exist in Automan.  " & vbCrLf & "Enter an Short Name of Length 5 to Create It.", "Create Voucher Type")
            Else
                ManualVType = InputBox(ManualVType & " Already Exist in Automan.  " & vbCrLf & "Enter an Short Name of Length 5 to Create It.", "Create Voucher Type")
            End If
            If ManualVType = "" Then CreateVType = "": Exit Function
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='" & left(ManualVType, 5) & "'")
            If TmpRst.RecordCount > 0 Then
                VTypeAlreadyExist = False
            Else
                VTypeAlreadyExist = True
            End If
        Loop
        
        
        G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                      "('FA','FA','" & ManualVType & "','" & mVType & "','" & Replace(mVType, " ", "") & "','" & ManualVType & "','Automatic','N','Y','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
                      
        Set TmpRst = G_FaCn.Execute("Select * From Voucher_Prefix Where V_Type='" & ManualVType & "' And Date_From=" & ConvertDate(PubStartDate) & " And Date_To=" & ConvertDate(PubEndDate) & "")
        If TmpRst.RecordCount = 0 Then
            G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('" & ManualVType & "'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'Tally'," & mStartSrlNo & ",'" & mDivCode & "')")
        End If
        CreateVType = ManualVType
    Else
        CreateVType = XNull(TmpRst(0))
    End If
End Function





Private Sub CmdImport_Click(Index As Integer)
    Dim X As Long
    Dim RsTemp          As ADODB.Recordset
    Dim RsDms           As ADODB.Recordset
    'Dim mCnt            As ADODB.Recordset
    Dim mSubGroupCounter    As Long
    Dim mSubCode$, mDmsSubCode$, mQry$, mNarr$, mLocalCentral$, mCondStr$
    Dim mFileName$, mFileTitle$, MState$, mCashCredit$, mVouCat$, mInvoiceNo$
    Dim mDocId$, mVType$, mVPrefix$, mVNo$, xDocId$, mVSNo%
    
On Error GoTo DispErr
                    
    
    'If XNull(RsDmsEnviro!CashAc) = "" Then MsgBox "Plz Define CashAc In DmsEnviro": Exit Sub
    mSubGroupCounter = G_CompCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0)
    
    CD1.FileName = ""
    
    Call SelectFile
    mFileName = CD1.FileName
    mFileTitle = CD1.FileTitle
    If mFileName = "" Then Exit Sub
    mFileTitle = mID(mFileTitle, 1, Len(mFileTitle) - 4)
    Set DmsConn = New Connection
    DmsConn.CursorLocation = adUseClient
    DmsConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mFileName & ";Extended Properties=Excel 8.0"
    
    Set RsDms = DmsConn.Execute("Select * from [" & mFileTitle & "$]")
    
    For X = 1 To 9999
        LBL.Refresh
    Next X
    
    
    With RsDms
    
                Select Case Index
                    Case BtnSubGroup
                        
                        If ChkFieldExist(RsDms, "Particulars") And ChkFieldExist(RsDms, "Nature") And _
                           ChkFieldExist(RsDms, "Debit") And ChkFieldExist(RsDms, "Credit") Then
                           
                                If .RecordCount > 0 Then
                                    Prg.Value = 0
                                    Prg.Visible = True
                                    Do Until .EOF
                                        GCn.BeginTrans
                                        G_FaCn.BeginTrans

                                            If InStr(1, XNull(!Nature), "Debtor") > 0 Then
                                                Call AutomanSubcode(XNull(!Particulars), "0020", "Customer", False)
                                            ElseIf InStr(1, XNull(!Nature), "Creditor") > 0 Then
                                                Call AutomanSubcode(XNull(!Particulars), "0016", "Supplier", False)
                                            End If
                                            

                                            If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                            .MoveNext
                                        GCn.CommitTrans
                                        G_FaCn.CommitTrans
                                    Loop
                                End If
                           
                        End If
                    Case BtnVoucher
                        'If XNull(RsDmsEnviro!SprDebtorGroupCode) = "" Then MsgBox "Plz Define SprDebtorGroupCode In DmsEnviro": Exit Sub
                        
                        
                        
                        If ChkFieldExist(RsDms, "VDate") And ChkFieldExist(RsDms, "Particulars") And _
                           ChkFieldExist(RsDms, "VType") And ChkFieldExist(RsDms, "VNo") And _
                           ChkFieldExist(RsDms, "Debit") And ChkFieldExist(RsDms, "Credit") Then
                        
                                
                                If .RecordCount > 0 Then
                                    Prg.Value = 0
                                    Prg.Visible = True
                                    Do Until .EOF
                                        GCn.BeginTrans
                                        G_FaCn.BeginTrans

                                            mInvoiceNo = XNull(!VType) & XNull(!VNo)

                                            If VNull(!Debit) > 0 Or VNull(!Credit) > 0 Then
                                                mDmsSubCode = XNull(!Particulars)
                                                mSubCode = AutomanSubcode(mDmsSubCode, "0009", "Others")
                                            End If
                                            
                                            
                                            If mSubCode = "" And (VNull(!Debit) > 0 Or VNull(!Credit) > 0) Then
                                                Call CreateErrLog(mVouCat, mInvoiceNo, "Party Name - " & XNull(!Particulars) & " Not Found In Automan")
                                            Else
                                                mVType = CreateVType(XNull(!VType), 1, PubDivCode)
                                                If mVType <> "" Then
                                                        mVPrefix = "Tally"
                                                        mVNo = CStr(Val(XNull(!VNo)))
                                                        mDocId = PubDivCode & PubSiteCode & PubSiteCode & mVType & Space(5 - Len(mVType)) & mVPrefix & Space(5 - Len(mVPrefix)) & Space(8 - Len(mVNo)) & mVNo
                                                        
                                                        If mDocId <> xDocId Then
                                                            GCn.Execute "Delete From Ledger Where DocId= '" & mDocId & "'"
                                                            GCn.Execute "Delete From LedgerM Where DocId= '" & mDocId & "'"
                                                            'GCn.Execute "Delete From DmsErrLog Where [Key]='" & mInvoiceNo & "'"
                                                            mVSNo = 1
                                                        Else
                                                            mVSNo = mVSNo + 1
                                                        End If
                                                        
                                                        
                                                        If VNull(!Debit) = 0 And VNull(!Credit) = 0 Then
                                                            GCn.Execute "IF Not Exists (Select * From dbo.LedgerM Where DocId = '" & mDocId & "') INSERT INTO dbo.LedgerM (   DocId,  Site_Code,  V_Type, v_Prefix,   V_No,   V_Date, Narration,  U_Name, U_EntDt,    U_AE,      DmsRefNo    )" & _
                                                                        "VALUES  ('" & mDocId & "', '" & PubSiteCode & "', '" & mVType & "',   '" & mVPrefix & "',  " & Val(mVNo) & ",  " & ConvertDate(XNull(!VDate)) & ",    '" & XNull(!Particulars) & "', '" & pubUName & "',   '" & PubLoginDate & "',   'A',  '" & mInvoiceNo & "'   )"
                                                        Else
                                                            GCn.Execute "INSERT INTO dbo.Ledger  (   DocId,  V_SNO,  V_Type, V_No,   v_Prefix,   Site_Code,  V_Date, SubCode,    ContraSub,  AmtDr,  AmtCr,  U_Name, U_EntDt, U_AE)" & _
                                                                        "VALUES  ('" & mDocId & "', " & mVSNo & ", '" & mVType & "',   " & Val(mVNo) & ",  '" & Val(mVPrefix) & "',  '" & PubSiteCode & "', " & ConvertDate(XNull(!VDate)) & ",  '" & mSubCode & "',   '', " & VNull(!Debit) & ", " & VNull(!Credit) & ", 'SA',  '" & PubLoginDate & "',   'A')"
                                                        End If
                                                        
                                                        xDocId = mDocId
                                                Else
                                                    Call CreateErrLog("Tally Import", mInvoiceNo, XNull(!VType) & " - Voucher Type Does Not Exit in Automan")
                                                End If
                                            End If
                                            If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                            .MoveNext
                                        GCn.CommitTrans
                                        G_FaCn.CommitTrans
                                    Loop
                                End If
                        End If
                End Select
                    
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


Private Sub SelectFile()
    
    CD1.CancelError = False
    CD1.DialogTitle = "Select CrmDms Excel Files"
    CD1.Filter = "Excel Files (*.xls)|*.xls"
    CD1.FilterIndex = 1
    CD1.Flags = cdlOFNHideReadOnly
    CD1.ShowOpen
    
End Sub



Sub CreateErrLog(mCategory As String, mKeyValue As String, mNarration As String)
    'GCn.Execute "Insert Into DmsErrLog(Cat, [Key], Narration, U_EntDt) Values('" & mCategory & "', '" & mKeyValue & "', '" & mNarration & "', " & ConvertDate(PubLoginDate) & ")"
End Sub


Private Sub Ini_Grid()
    With FgridErr
        .Cols = 4
        
        .ColWidth(0) = 400
        
        .TextMatrix(0, FErr_Cat) = "Category"
        .ColAlignment(FErr_Cat) = flexAlignLeftCenter
        .ColWidth(FErr_Cat) = 2000
        
        .TextMatrix(0, FErr_DmsRef) = "Reference"
        .ColAlignment(FErr_DmsRef) = flexAlignLeftCenter
        .ColWidth(FErr_DmsRef) = 2500
        
        .TextMatrix(0, FErr_Narration) = "Narration"
        .ColAlignment(FErr_Narration) = flexAlignLeftCenter
        .ColWidth(FErr_Narration) = 10000
    End With
End Sub

Function AutomanSubcode(mDmsSubCode As String, mAutomanGroupCode As String, mNature As String, Optional mAskForOpeningAccounts As Boolean = True) As String
    Dim mConn As New ADODB.Connection
    Dim RsTemp As ADODB.Recordset
    Dim rsTemp1 As ADODB.Recordset
    Dim RsTempCity As ADODB.Recordset
    Dim mSubGroupCounter As Long
    Dim mSubCode$, mQry$, mname$, mCityCode$, mStateCode$
    Dim mLocalCentral$
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
                 "'" & mname & "', '" & mname & "', '" & mAutomanGroupCode & "', '" & mNature & "', '', " & _
                 "'', '', '', '', '', " & _
                 "'', '', 1, '" & pubUName & " ', " & _
                 "" & ConvertDate(PubLoginDate) & ", 'A', 'A', 'N')"
                         
                         
                         
            GCn.Execute mQry
            If PubBackEnd = "A" Then G_FaCn.Execute mQry
            
            G_CompCn.Execute ("Update  SubGroupCounter Set SubGroupAcCode=" & mSubGroupCounter + 1 & " ")
            
            AutomanSubcode = mSubCode
        End If


    Set RsTemp = Nothing
    Set rsTemp1 = Nothing
End Function


