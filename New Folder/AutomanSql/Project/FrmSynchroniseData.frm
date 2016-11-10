VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form FrmSynchroniseData 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1650
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   2910
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer Timer1 
      Left            =   5670
      Top             =   45
   End
   Begin MSComctlLib.ProgressBar Prg 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1710
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Tables Current Records :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   7
      Top             =   1365
      Width           =   2250
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Table Records :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   6
      Top             =   1095
      Width           =   1395
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Table Name : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   5
      Top             =   810
      Width           =   1215
   End
   Begin VB.Label lblError 
      AutoSize        =   -1  'True
      Caption         =   "Error"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2295
      TabIndex        =   4
      Top             =   315
      Width           =   1560
   End
   Begin VB.Label LblCurrRecord 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2490
      TabIndex        =   3
      Top             =   1365
      Width           =   600
   End
   Begin VB.Label LblRecords 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2490
      TabIndex        =   2
      Top             =   1095
      Width           =   600
   End
   Begin VB.Label LblTable 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2490
      TabIndex        =   1
      Top             =   810
      Width           =   600
   End
End
Attribute VB_Name = "FrmSynchroniseData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim GcnOffLine As ADODB.Connection
Dim RsTables As Recordset

Dim mQry As String
Dim IsUploading As Boolean

Private Sub Form_Activate()
    If StrCmp(PubOffLineServer, PubServerName) Then Unload Me
End Sub

Private Sub Form_Load()
    If StrCmp(PubOffLineServer, PubServerName) Then
        MsgBox "Offline Server and Online Server is Same Can't Run Synchronization!"
    Else
        MsgBox "Offline Server : " & PubOffLineServer & vbCrLf & "Online Server : " & PubServerName & vbCrLf & "Database : " & PubCenDataPath
        Connect
        IsUploading = False
        'mQry = "Select * From Synchronisation_Fields Order By RowId"
        mQry = "SELECT F.*  FROM Synchronisation_Fields F LEFT JOIN INFORMATION_SCHEMA.TABLES T ON F.TableName =T.TABLE_NAME AND T.TABLE_SCHEMA ='dbo' WHERE T.TABLE_NAME IS Not Null  ORDER BY RowId"
    
        Set RsTables = GcnOffLine.Execute(mQry)
        Timer1.Interval = 10000
    End If
End Sub


Sub Upload()
    Dim I As Integer, j As Integer, K As Integer, l As Integer
    Dim RsRecords As ADODB.Recordset
    Dim RsCurrRecord As ADODB.Recordset
    Dim RsColumns As ADODB.Recordset
    Dim StrTableColumns() As String
    Dim StrColumns As String
    Dim mTrans As Boolean
    Dim mErrLogStr As String
    
    Dim isErrorOccured As Boolean
    
    LblTable = ""
    LblRecords = ""
    LblCurrRecord = ""
    
    
    On Error GoTo ELoop
    
    
    lblError.Visible = False
    GCn.CommandTimeout = 1024
    GcnOffLine.CommandTimeout = 1024
    GCn.BeginTrans
    GcnOffLine.BeginTrans
    mTrans = True
    
        IsUploading = True
        RsTables.MoveFirst
        For I = 0 To RsTables.RecordCount - 1
            mQry = "Select Distinct " & RsTables!UniqueKey & " as UniqueKey From " & RsTables!TableName & " Where " & RsTables!UpLoadDateField & " is Null "
            Set RsRecords = GcnOffLine.Execute(mQry)
            StrColumns = GetColumnString(RsTables!TableName, GcnOffLine)
            StrTableColumns = Split(StrColumns, ",")
            StrColumns = Replace(StrColumns, "$", "")
            LblTable = XNull(RsTables!TableName)
            For j = 0 To RsRecords.RecordCount - 1
                
                LblRecords = XNull(RsRecords.RecordCount)
                GCn.Execute "Delete From " & RsTables!TableName & " Where " & RsTables!UniqueKey & " = '" & RsRecords!UniqueKey & "'"
                mQry = "Select * From " & RsTables!TableName & " With (NoLock) Where " & RsTables!UniqueKey & " = '" & RsRecords!UniqueKey & "'"
                Set RsCurrRecord = GcnOffLine.Execute(mQry)
                LblCurrRecord = XNull(RsRecords!UniqueKey) & " ------ " & RsRecords.AbsolutePosition
                
                LblCurrRecord.Refresh

                For l = 0 To RsCurrRecord.RecordCount - 1
                    mQry = "Insert Into " & RsTables!TableName & "( " & StrColumns & " ) Values ("
                    For K = 0 To UBound(StrTableColumns)
                        If left(StrTableColumns(K), 1) = "$" Then
                            mQry = mQry & "" & ConvertDate(RsCurrRecord.Fields(Replace(StrTableColumns(K), "$", "")).Value) & "" & IIf(K < UBound(StrTableColumns), ",", ")")
                        Else
                            mQry = mQry & "'" & Replace(XNull(RsCurrRecord.Fields(Replace(StrTableColumns(K), "$", "")).Value), "'", "") & "'" & IIf(K < UBound(StrTableColumns), ",", ")")
                        End If
                    Next K
                    GCn.Execute mQry
                        
                    
                    mQry = " Update " & RsTables!TableName & "  Set " & RsTables!UpLoadDateField & " = getdate()  Where " & RsTables!UniqueKey & " = '" & RsRecords!UniqueKey & "' "
                    GCn.Execute mQry
                    GcnOffLine.Execute mQry
                    
                    
                    RsCurrRecord.MoveNext
                Next l
                'RsCurrRecord.Close
                'Set RsCurrRecord = Nothing
                RsRecords.MoveNext
            Next j
            'RsRecords.Close
            'Set RsRecords = Nothing
            RsTables.MoveNext
        Next I
        
        IsUploading = False
    GCn.CommitTrans
    GcnOffLine.CommitTrans
    mTrans = False
    'MsgBox "Data Synchronisation Completed!..."
Exit Sub
ELoop:
    lblError.Visible = True
    mErrLogStr = ""

    MsgBox "Table Name : " & XNull(RsTables!TableName) & vbCrLf & err.Description & vbCrLf & "Please Reload Synchronisation."
    CreateLog XNull(RsTables!TableName), mErrLogStr, err.Description
    If mTrans = True Then
        IsUploading = False
        GCn.RollbackTrans
        GcnOffLine.RollbackTrans
        Unload Me
    End If
End Sub

Sub CreateLog(TableName As String, SearchKey As String, Message As String)
    mQry = "Insert Into Synchronisation_Errors(TableName, SearchKey, Message) " & _
         "Values('" & TableName & "', '" & SearchKey & "', '" & Replace(Message, "'", "`") & "')   "
    GcnOffLine.Execute mQry
    ShowNotErrorRecords
End Sub

Sub ShowNotErrorRecords()
    Dim RsTemp As ADODB.Recordset
    mQry = "Select * from Synchronisation_Errors With (NoLock) Order By RowID"
    Set RsTemp = GCn.Execute(mQry)
    Set FGrid.DataSource = RsTemp
    Set RsTemp = Nothing
End Sub

Sub Connect()
    Set GcnOffLine = New ADODB.Connection
    GcnOffLine.CursorLocation = adUseClient
    GcnOffLine.IsolationLevel = adXactChaos
    GcnOffLine.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubOffLineServer
End Sub

Private Sub Timer1_Timer()
    If Not IsUploading Then
        Connect
        ConnectDb
        Upload
    End If
End Sub
