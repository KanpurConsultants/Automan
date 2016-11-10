VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmInsurance 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Insurance Company Master"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8595
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
   LinkTopic       =   "form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   8595
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DgCity 
      Height          =   2175
      Left            =   30
      Negotiate       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2130
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   3836
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
         Caption         =   "City"
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
      Left            =   2820
      MaxLength       =   50
      TabIndex        =   14
      Top             =   1785
      Width           =   3840
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
      Left            =   2820
      MaxLength       =   50
      TabIndex        =   12
      Top             =   1545
      Width           =   3840
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
      Left            =   2820
      MaxLength       =   50
      TabIndex        =   10
      Top             =   1305
      Width           =   3840
   End
   Begin MSDataGridLib.DataGrid DgInsurance 
      Height          =   2040
      Left            =   4305
      Negotiate       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2070
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
         Caption         =   "Insurance Company Name"
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
      Left            =   2820
      MaxLength       =   1
      TabIndex        =   1
      Top             =   585
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
         Left            =   0
         TabIndex        =   8
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
      Index           =   2
      Left            =   2820
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1065
      Width           =   3840
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
      Left            =   2820
      MaxLength       =   50
      TabIndex        =   2
      Top             =   825
      Width           =   3840
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   661
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person..."
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
      Left            =   1365
      TabIndex        =   15
      Top             =   1815
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City..................."
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
      Left            =   1365
      TabIndex        =   13
      Top             =   1575
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address2..........."
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
      Left            =   1365
      TabIndex        =   11
      Top             =   1335
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code.................."
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
      Left            =   1365
      TabIndex        =   9
      Top             =   615
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address1..........."
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
      Left            =   1365
      TabIndex        =   6
      Top             =   1095
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name...................."
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
      Left            =   1365
      TabIndex        =   4
      Top             =   855
      Width           =   1695
   End
End
Attribute VB_Name = "FrmInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsInsurance As ADODB.Recordset
Dim RsCity      As ADODB.Recordset
Dim Master As ADODB.Recordset
Public MasterFormExit As Boolean

Private Const tCode             As Byte = 0
Private Const tName             As Byte = 1
Private Const tAdd1             As Byte = 2
Private Const tAdd2             As Byte = 3
Private Const tCity             As Byte = 4
Private Const tContactPerson    As Byte = 5

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
    'SetResolutionFormLoad Me
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "Select Code as SearchCode,I.*, C.CityName from Insurance I Left Join City C On I.City=C.CityCode Order by Name", GCn, adOpenDynamic, adLockOptimistic
   
    Set RsInsurance = GCn.Execute("select Code as Code, Name  from Insurance Order by Name")
    Set DgInsurance.DataSource = RsInsurance
            
            
    Set RsCity = GCn.Execute("Select CityCode as Code, CityName As Name from City Order by CityName")
    Set DgCity.DataSource = RsCity
            
    Disp_Text SETS("INI", Me, Master)
    MoveRec
  Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsInsurance = Nothing
Set Master = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim VNo As Long
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    txt(tName).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub



Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                GCn.BeginTrans
                    GCn.Execute ("delete from Insurance where Code= '" & Master!SearchCode & "'")
                GCn.CommitTrans
                
                Master.Requery
                Call MoveRec
                RsInsurance.Requery
                
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
    EditDesc = txt(tName).TEXT
    txt(tName).SetFocus
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
    GSQL = "select I.Code as SearchCode, I.Name, I.Add1 As Address1, I.Add2 As Address2, C.CityName As City, I.ContactPerson from Insurance I Left Join City C On I.City=C.CityCode Order By I.Name"
    Set SearchForm = Me
    'FIND.Show vbModal
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal

    
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("searchcode='" & MyValue & "'")
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
    RsInsurance.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim ItemCode As Integer
    Dim Rst As ADODB.Recordset
    Dim mMaxId As Long
    Dim mCondStr$
   On Error GoTo errlbl
   
     
     If IsValid(txt(tName), "Insurance Desc") = False Then Exit Sub
     
    If TopCtrl1.TopText2 = "Edit" Then mCondStr = " And Name <> '" & Master!Name & "'"
    Set Rst = GCn.Execute("select Name from Insurance where Name = '" & txt(tName) & "' " & mCondStr & " ")
    If Rst.RecordCount > 0 Then
        MsgBox "Duplicate Insurance Company Name", vbInformation, "Validation Check": txt(tName).SetFocus: Exit Sub
    End If
    Set Rst = Nothing

    
 Grid_Hide
 GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        mMaxId = GCn.Execute("Select " & vIsNull("Max(" & cVal("Code") & ")", "0") & "+1 From Insurance").Fields(0).value
                
        GCn.Execute ("insert into Insurance(Code, Name, Add1, Add2, City, ContactPerson, U_Name, U_EntDt, U_AE) " & _
            " values('" & mMaxId & "' ,'" & txt(tName) & "','" & txt(tAdd1) & "','" & txt(tAdd2) & "','" & txt(tCity).Tag & "','" & txt(tContactPerson) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & Left(TopCtrl1.TopText2, 1) & "')")
    Else
        GCn.Execute "update Insurance  set Name='" & txt(tName) & "', Add1 = '" & txt(tAdd1) & "', Add2 = '" & txt(tAdd2) & "', City = '" & txt(tCity).Tag & "', ContactPerson='" & txt(tContactPerson) & "', U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & Left(TopCtrl1.TopText2, 1) & "' Where Code = '" & Master!SearchCode & "'"
    End If
GCn.CommitTrans
mTrans = False
    Master.Requery
    RsInsurance.Requery
    Master.FIND "SearchCode = '" & Master!SearchCode & "'"
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
        Case tCity
            'DGridColSwap DgCity, 1
            RsCity.Sort = "Name"
            If RsCity.RecordCount = 0 Or txt(Index).TEXT = "" Then Exit Sub
            If RsCity.EOF = True Or RsCity.BOF = True Then Exit Sub
            If txt(Index).TEXT <> RsCity!Name Then
                RsCity.MoveFirst
                RsCity.FIND "Name ='" & txt(Index).TEXT & "'"
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
    Case tName
        DGridTxtKeyDown_Mast DgInsurance, txt, Index, RsInsurance, KeyCode, False, 1
    Case tCity
        DGridTxtKeyDown DgCity, txt, Index, RsCity, KeyCode, False, 1, frmCity, "frmCity"
End Select
If DgInsurance.Visible = False And DgCity.Visible = False And FrmList.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> tContactPerson Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = tContactPerson Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If Index <> tCode Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case Index
    Case tCity
        DGridTxtKeyPress txt, Index, RsCity, KeyAscii, "Name"
End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case tName
        DGridTxtKeyUp_Mast txt, Index, RsInsurance, KeyCode, "Name"
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
        txt(tName) = XNull(!Name)
        txt(tAdd1) = XNull(!Add1)
        txt(tAdd2) = XNull(!Add2)
        txt(tCity).Tag = XNull(!City)
        txt(tCity) = XNull(!CityName)
        txt(tContactPerson) = !ContactPerson
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
End Sub
Private Sub Grid_Hide()
    If DgInsurance.Visible = True Then DgInsurance.Visible = False
    If DgCity.Visible = True Then DgCity.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error Resume Next
    Select Case Index
        Case tCity
            If RsCity.EOF = False And RsCity.BOF = False Then
                If txt(Index).TEXT <> "" Then
                    txt(Index).TEXT = RsCity!Name
                    txt(Index).Tag = RsCity!Code
                End If
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        
    End Select
End Sub


