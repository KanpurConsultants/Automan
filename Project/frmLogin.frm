VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   0  'None
   Caption         =   "Automan - User Login"
   ClientHeight    =   2865
   ClientLeft      =   2790
   ClientTop       =   3150
   ClientWidth     =   6750
   FillColor       =   &H00C0E0FF&
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLogin.frx":030A
   ScaleHeight     =   1692.734
   ScaleMode       =   0  'User
   ScaleWidth      =   6337.886
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2775
      MaxLength       =   10
      TabIndex        =   1
      ToolTipText     =   "Enter User Name in this field"
      Top             =   1410
      Width           =   2475
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2775
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Type User Password in this field"
      Top             =   1740
      Width           =   2475
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "User Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   435
      Left            =   2535
      TabIndex        =   6
      Top             =   255
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      Height          =   2850
      Left            =   15
      Top             =   0
      Width           =   6735
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   42.253
      X2              =   6337.886
      Y1              =   1444.585
      Y2              =   1444.585
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   42.253
      X2              =   6337.886
      Y1              =   549.474
      Y2              =   549.474
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   450
      Left            =   330
      TabIndex        =   5
      Top             =   1140
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   1485
      TabIndex        =   4
      Top             =   1410
      UseMnemonic     =   0   'False
      Width           =   1080
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
      Height          =   225
      Index           =   20
      Left            =   2655
      TabIndex        =   3
      Top             =   1440
      Width           =   15
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1485
      TabIndex        =   2
      Top             =   1740
      UseMnemonic     =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long



'Private Const CtrlBColOrg = &HCFE0E0         'Orginal BackColour
'Private Const CtrlFColOrg = &H80000008       'Orginal ForeColour
'Private Const CtrlBCol = &H0&                'Changed BackColour
'Private Const CtrlFCol = &HFFFF&             'Changed ForeColour
Private Const UName = 0                      'pubUName Text Index No.
Private Const UPass = 1                      'UserPassword Text Index No.

Private Sub Form_Activate()
Txt(0).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = 27 Then
'    If MsgBox("Are You Sure To Quit ?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        End
'    End If
End If
Exit Sub
ELoop:      MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim buff As String, bufflen As Long, result As Long, I As Byte

Dim myFileObj As New FileSystemObject

'If myFileObj.FileExists(GetFolderPath(CSIDL_SYSTEM) & "\MSDATGRD.OCX") Then
'    Shell "regsvr32  /u/s " & GetFolderPath(CSIDL_SYSTEM) & "\MSDATGRD.OCX"
'End If
'Dim varScrptObj As New Scripting.FileSystemObject, varTxtstrm As Scripting.TextStream, varTxtstrm1 As Scripting.TextStream, Fob As New FileSystemObject
'Dim a As Long, SourcePath$, DestPath$, DataPath$, b As Long
'            Dim s As FileSystemObject
'            If Fob.FileExists("c:\Automan.ini") = False Then MsgBox "S/W ini File Missing, Contact to System Administrator", vbCritical: End
'            Set varTxtstrm = varScrptObj.OpenTextFile("c:\Automan.ini")
'            varTxtstrm.SkipLine
'            varTxtstrm.Skip 2: DataPath = varTxtstrm.ReadLine
'            varTxtstrm.SkipLine
'            varTxtstrm.SkipLine
'            varTxtstrm.SkipLine
'            varTxtstrm.SkipLine
'            varTxtstrm.SkipLine
'            varTxtstrm.SkipLine
'            If Not varTxtstrm.AtEndOfStream Then varTxtstrm.Skip 2: pub_DllPath = varTxtstrm.ReadLine Else pub_DllPath = "C:\Windows\System"
'            Set s = New FileSystemObject
'            SourcePath = Mid(DataPath, 1, Len(DataPath) - 4) & "DMFa.DLL"
'            DestPath = pub_DllPath & "\DMFa.Dll"
'            If ConvertDate(FileDateTime(DestPath)) < ConvertDate(FileDateTime(SourcePath)) Then
'                MsgBox "Your DMFa.Dll is Older Then New DMFa.Dll.Automan Will first register it.", vbOKOnly
'                'b = Shell("regsvr32 " & pub_DllPath & "\DmFa.Dll -u")
'                s.CopyFile SourcePath, DestPath, True
'                b = Shell("regsvr32 " & pub_DllPath & "\DmFa.Dll")
'                MsgBox " Please Reload the Automan"
'                End
'            End If

Me.CAPTION = PubPackage & "-User Login"
For I = 0 To 1
    Txt(I).BackColor = CtrlBColOrg
    Txt(I).ForeColor = CtrlFColOrg
    Txt(I).BackColor = CtrlBColOrg
    Txt(I).ForeColor = CtrlFColOrg
Next
Label3.left = (Me.width - Label3.width) / 2
buff = "               " ' Space(15)
bufflen = Len(buff)
RegOpenKeyEx &H80000001, "Control Panel\International", 0, &H2, result
RegQueryValueEx result, "sShortDate", 0, 2, buff, bufflen
If left(buff, 11) <> "dd/MMM/yyyy" Then
    buff = "dd/MMM/yyyy"
    bufflen = Len(buff) + 1
    RegSetValueEx result, "sShortDate", 0, 1, buff, bufflen
'    If CDate("01/12/99") <> CDate(Format(("01/12/99"), "dd/mm/yyyy")) Then MsgBox "You Must Restart Your Computer Before Proceed.", vbInformation, "Information"
End If
RegCloseKey &H80000001
RegCloseKey result
'''''
PubSec = "RAHUL"
PubBackEnd = "A"
'''''

Exit Sub
ELoop:
'    If err.Number = 62 Then
'          MsgBox "Automan.ini is Changed.Write the Changes ?"
'          Set varTxtstrm1 = varScrptObj.OpenTextFile("c:\Automan.ini", ForAppending)
'          varTxtstrm1.Write ("7=C:\Windows\System")
'
'        End
'    End If
'
'
'    If err.Number = 53 Then
'        s.CopyFile DestPath, SourcePath, True
'        End
'    End If
MsgBox err.Description, vbInformation, "Information"

End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Call Ctrl_GetFocus(Index)
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Index <> UPass Then
    SendKeysA vbKeyTab, True
ElseIf KeyCode = 13 And Index = UPass Then
    Call CompFormLoad
End If
End Sub
Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 40 And Index <> UPass Then  'keydown= 40
    SendKeysA vbKeyTab, True
ElseIf KeyCode = 38 And Index <> UName Then    'keyup =38
    SendKeys "+{Tab}"
ElseIf KeyCode = 40 And Index = UPass Then  'KeyUp = 38
    Call CompFormLoad
End If
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Call Ctrl_validate(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
If Txt(UName).TEXT = "" Then
    MsgBox "User Name Is Required", vbExclamation, "Validation Check"
    Txt(UName).SetFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub Ctrl_validate(Index As Integer)
    Txt(Index).BackColor = CtrlBColOrg
    Txt(Index).ForeColor = CtrlFColOrg
End Sub
Private Sub Ctrl_GetFocus(Index As Integer)
    Txt(Index).BackColor = CtrlBCol
    Txt(Index).ForeColor = CtrlFCol
End Sub

Public Sub CompFormLoad()
Dim I As Integer
Dim xChar As Integer
Dim mLine$
On Error Resume Next
Dim varScrptObj As New Scripting.FileSystemObject, varTxtstrm As Scripting.TextStream, varTxtstrm1 As Scripting.TextStream
Dim fob As New FileSystemObject, AdoCompany As ADODB.Recordset, Comp_Path As String
Dim mCompany$
    
    Set AdoCompany = New Recordset
    AdoCompany.CursorLocation = adUseClient
        If Len(Txt(UName).TEXT) = 0 Then
            MsgBox "Invalid User", 32, "Login"
            Txt(UName).SetFocus
            Exit Sub
        Else
            If fob.FileExists(App.Path & "\Automan.ini") Then
                Set varTxtstrm = varScrptObj.OpenTextFile(App.Path & "\Automan.ini")
            Else
                If fob.FileExists("c:\AutomanSQL.ini") = False Then
                    If fob.FileExists("c:\Automan.ini") = False Then MsgBox "S/W ini File Missing, Contact to System Administrator", vbCritical: End
                    Set varTxtstrm = varScrptObj.OpenTextFile("c:\Automan.ini")
                Else
                    Set varTxtstrm = varScrptObj.OpenTextFile("c:\AutomanSQL.ini")
                End If
            End If
            
            varTxtstrm.SkipLine
            varTxtstrm.Skip 2: Pub_DataPath = varTxtstrm.ReadLine
            varTxtstrm.Skip 2: PubRepoPath = varTxtstrm.ReadLine
            varTxtstrm.Skip 2: PubBkpPath = varTxtstrm.ReadLine
            varTxtstrm.Skip 2: PubFaReportPath = varTxtstrm.ReadLine
            
            
            If Not varTxtstrm.AtEndOfStream Then varTxtstrm.Skip 2: PubFaDosPort = varTxtstrm.ReadLine Else PubFaDosPort = "PRN"
            If Not varTxtstrm.AtEndOfStream Then varTxtstrm.Skip 2: PubRunPIF = varTxtstrm.ReadLine Else PubRunPIF = "Y"
            If Not varTxtstrm.AtEndOfStream Then varTxtstrm.Skip 2: pub_DllPath = varTxtstrm.ReadLine Else pub_DllPath = "C:\Windows\System32"
            
             If fob.FileExists(App.Path & "\Automan.ini") Then
                Set varTxtstrm = varScrptObj.OpenTextFile(App.Path & "\Automan.ini")
            Else
            If fob.FileExists("c:\AutomanSQL.ini") = False Then
                If fob.FileExists("c:\Automan.ini") = False Then MsgBox "S/W ini File Missing, Contact to System Administrator", vbCritical: End
                Set varTxtstrm = varScrptObj.OpenTextFile("c:\Automan.ini")
            Else
                Set varTxtstrm = varScrptObj.OpenTextFile("c:\AutomanSQL.ini")
            End If
            End If
            
            Do Until varTxtstrm.AtEndOfLine
                mLine = varTxtstrm.ReadLine
                                
                If UTrim(left(mLine, 10)) = "SQLSERVER=" Then
                    PubServerName = mID(mLine, 11, Len(mLine) - 10)
                ElseIf UTrim(left(mLine, 17)) = "SQLSERVERCOMPANY=" Then
                    PubServerNameCompany = mID(mLine, 18, Len(mLine) - 17)
                ElseIf UTrim(left(mLine, 8)) = "COMPANY=" Then
                    mCompany = mID(mLine, 9, Len(mLine) - 8)
                    PubCompanyDbName = mCompany
                ElseIf UTrim(left(mLine, 5)) = "USER=" Then
                    PubDbUserCompany = mID(mLine, 6, Len(mLine) - 5)
                ElseIf UTrim(left(mLine, 9)) = "PASSWORD=" Then
                    PubDbPassCompany = mID(mLine, 10, Len(mLine) - 9)
                ElseIf UTrim(left(mLine, 5)) = "DATA=" Then
                    Pub_DataPath = mID(mLine, 6, Len(mLine) - 5)
                ElseIf UTrim(left(mLine, 8)) = "REPORTS=" Then
                    PubRepoPath = mID(mLine, 9, Len(mLine) - 8)
                ElseIf UTrim(left(mLine, 7)) = "BACKUP=" Then
                    PubBkpPath = mID(mLine, 8, Len(mLine) - 6)
                ElseIf UTrim(left(mLine, 10)) = "REPORTSFA=" Then
                    PubFaReportPath = mID(mLine, 11, Len(mLine) - 10)
                ElseIf UTrim(left(mLine, 8)) = "DOSPORT=" Then
                    PubFaDosPort = mID(mLine, 9, Len(mLine) - 8)
                ElseIf UTrim(left(mLine, 7)) = "RUNPIF=" Then
                    PubRunPIF = mID(mLine, 8, Len(mLine) - 7)
                ElseIf UTrim(left(mLine, 9)) = "DLLPATH=" Then
                    pub_DllPath = mID(mLine, 10, Len(mLine) - 9)
                ElseIf UTrim(left(mLine, 14)) = "OFFLINESERVER=" Then
                    PubOffLineServer = mID(mLine, 15, Len(mLine) - 14)
                End If
                
            Loop
            
            
            
           If fob.FolderExists(PubBkpPath) = False Then fob.CreateFolder (PubBkpPath)
            
            
            
            If Trim(PubFaDosPort) = "" Then PubFaDosPort = "PRN"
            If Trim(PubRunPIF) = "" Then PubRunPIF = "Y"
            If Trim(pub_DllPath) = "" Then pub_DllPath = "C:\Windows\System32"
            If PubServerNameCompany = "" Then PubServerNameCompany = PubServerName
            If PubDbUserCompany = "" Then PubDbUserCompany = "sa"
            
                        
            PubBackEnd = IIf(Trim(PubServerName) = "", "A", "S")
            
                
                
            If PubOffLineServer <> "" Then
                If MsgBox("Do You want to Run Software Online? ", vbYesNo) = vbNo Then
                    PubServerNameCompany = PubOffLineServer
                    PubServerName = PubOffLineServer
                    PubDbUser = "SA"
                    PubDbPass = ""
                    PubDbUserCompany = "SA"
                    PubDbPassCompany = ""
                End If
            End If
                
                
                

            If PubBackEnd = "A" Then
                Comp_Path = Pub_DataPath & "\Company.mdb"
                Set G_CompCn = New Connection
                With G_CompCn
                    .CursorLocation = adUseClient
                    .Provider = "Microsoft.Jet.OLEDB.4.0"
                    .ConnectionString = "Data Source=" & Comp_Path & ";Persist Security Info=False"
                    .Open
                End With
            Else
                Set G_CompCn = New Connection
                G_CompCn.CursorLocation = adUseClient
                If mCompany = "" Then
                    G_CompCn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Company;Data Source=" & PubServerName
                Else
                    G_CompCn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUserCompany & ";Password=" & PubDbPassCompany & ";Initial Catalog=" & mCompany & ";Data Source=" & PubServerNameCompany
                End If
            End If
            
            UpdateTableStructureCompany
            
            If UCase(Txt(UName)) = "ADMIN" And UCase(Txt(UPass)) = "DMADMIN" Then
                pubUName = "ADMIN"
                PubULabel = AdoCompany!Label
                pubUAcPosting = AdoCompany!AcPosting
                Set AdoCompany = Nothing
                Set rsUserPerm = G_CompCn.Execute("select * from user2 where user_name='SA'")
                frmCompany.Show
                Exit Sub
            End If
            
            G_CompCn.Execute "Alter table UserMast add BckpY_N text (1) "
            
            If G_CompCn.Execute("SELECT COUNT(*) FROM UserMast").Fields(0).Value = 0 Then G_CompCn.Execute ("insert into UserMast(USER_NAME,PASSWD,LABEL) values('SA','','1')")
            AdoCompany.Open "select * from UserMast WHERE USER_NAME='" & Txt(UName) & "'", G_CompCn, adOpenStatic, adLockReadOnly
            If AdoCompany.RecordCount = 0 Then MsgBox "     Invalid User     ", 32, "Login": Txt(UName).TEXT = "": Txt(UName).SetFocus: Exit Sub
            If XNull(AdoCompany!BckpY_N) = "1" Then
                'MsgBox "Backup process on any machine is continued.Please Wait or Login after some time. "
                'Exit Sub
            End If
            If UCase(Txt(UPass)) = UCase(DCODIFY(XNull(AdoCompany!PASSWD))) Then
                pubUName = AdoCompany!user_name
                PubULabel = AdoCompany!Label
                pubUAcPosting = AdoCompany!AcPosting
                Set AdoCompany = Nothing
                frmCompany.Show
                Unload Me
            Else
                MsgBox "     Invalid Password     ", 32, "Login"
                Txt(UPass).TEXT = ""
                Txt(UPass).SetFocus
                Call Ctrl_GetFocus(UPass)
               Exit Sub
            End If
            Set rsUserPerm = G_CompCn.Execute("select * from user2 where user_name='" & pubUName & "'")
            
            
        End If
Exit Sub



ELoop:
    MsgBox err.Description, vbInformation, "Information"
    Exit Sub
End Sub


Private Function DCODIFY(Txt As String) As String
    If Txt = "" Then DCODIFY = "": Exit Function
    Dim xxx As String
    Dim xx As Byte, MyVal As Byte
    MyVal = Asc(left(Txt, 1)) - 27
    xxx = ""
    For xx = 1 To Len(Txt) - 1
        xxx = xxx + Chr(Asc(mID(Txt, xx + 1, 1)) - 27 - MyVal)
    Next
    DCODIFY = xxx
End Function



Sub UpdateTableStructureCompany()
    On Error Resume Next
        Dim mQry$
        
        AddNewField G_CompCn, "Company", "DbUser", "VarChar(50)", ""
        AddNewField G_CompCn, "Company", "DbPass", "VarChar(50)", ""
        'G_CompCn.Execute "Update Company Set DbUser = 'prayag09', DbPass='praydb'"
        
        
End Sub



