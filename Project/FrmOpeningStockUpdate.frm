VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmUpdation 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Physical Stock Updation"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
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
   ScaleHeight     =   2580
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Adjustment Details"
      Height          =   1230
      Left            =   45
      TabIndex        =   12
      Top             =   15
      Width           =   3165
      Begin VB.CheckBox ChkAssumeZero 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CFE0E0&
         Caption         =   "Assume Zero if not in List"
         Height          =   330
         Left            =   330
         TabIndex        =   17
         Top             =   855
         Width           =   2685
      End
      Begin VB.TextBox Txt1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   210
         Width           =   1380
      End
      Begin VB.CheckBox ChkTB 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CFE0E0&
         Caption         =   "Taxable...."
         Height          =   435
         Left            =   315
         TabIndex        =   14
         Top             =   495
         Width           =   1215
      End
      Begin VB.CheckBox ChkTP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CFE0E0&
         Caption         =   "Taxpaid...."
         Height          =   345
         Left            =   1800
         TabIndex        =   13
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label LBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date...................."
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   16
         Top             =   270
         Width           =   1605
      End
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Exit"
      Height          =   390
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton CmdSelect 
      BackColor       =   &H00CFE0E0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5970
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1545
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   15
      TabIndex        =   6
      Top             =   1560
      Width           =   5940
   End
   Begin VB.Frame FrameColumn 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Columns Of Excel File "
      Height          =   1260
      Left            =   3270
      TabIndex        =   2
      Top             =   0
      Width           =   3030
      Begin VB.Label LblCol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5. Tp_Qty"
         Height          =   195
         Index           =   4
         Left            =   1455
         TabIndex        =   11
         Top             =   675
         Width           =   840
      End
      Begin VB.Label LblCol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Bin"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   10
         Top             =   975
         Width           =   495
      End
      Begin VB.Label LblCol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Part_No"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   5
         Top             =   330
         Width           =   900
      End
      Begin VB.Label LblCol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Tb_Qty"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   4
         Top             =   675
         Width           =   840
      End
      Begin VB.Label LblCol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Part_Name"
         Height          =   195
         Index           =   1
         Left            =   1455
         TabIndex        =   3
         Top             =   330
         Width           =   1170
      End
   End
   Begin VB.CommandButton CmdStart 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Start"
      Height          =   390
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   2055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1875
      Visible         =   0   'False
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label LblSelect 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select MS-Excel File "
      Height          =   195
      Left            =   45
      TabIndex        =   9
      Top             =   1335
      Width           =   1770
   End
End
Attribute VB_Name = "FrmUpdation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mFileName$
Dim ScreenFactor    As Single

Private Sub Check1_Click()

End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSelect_Click()
    mFileName = ""
  
    'CD1.InitDir = Pub_DataPath
    CD1.CancelError = False
    CD1.DialogTitle = LblSelect
    CD1.Filter = "Excel Files (*.xls)|*.xls"
    CD1.FilterIndex = 1
    CD1.Flags = cdlOFNHideReadOnly
  
    CD1.ShowOpen

    Txt = CD1.FileName
    mFileName = CD1.FileTitle
End Sub

Private Sub CmdStart_Click()
                       
    If Txt = "" Then MsgBox "Please Select Any File": Exit Sub
    If Txt1(0) = "" Or IsDate(Txt1(0)) = False Then MsgBox "Please Fill Adjustment Date": Exit Sub
    ProcStockAdjustment
    
    
End Sub

Private Sub Form_KeyPress(keyascii As Integer)
    If keyascii = 13 Or keyascii = 10 Then SendKeysA vbKeyTab, True
End Sub

Private Sub Form_Load()
    'SetResolutionFormLoad Me
    Me.Icon = MDIForm1.Icon
    
    DispText
End Sub


Private Sub ProcStockAdjustment()

'    On Error GoTo DispErr
    
    Dim RsPart          As ADODB.Recordset
    Dim RsTemp          As ADODB.Recordset
    Dim RsAdj As ADODB.Recordset
    Dim mVNoIssue       As Long
    Dim mVNoReceive     As Long
    Dim I               As Long
    Dim j%
    Dim mSrlIssue       As Long
    Dim mSrlReceive     As Long
    Dim mTmpTbl As Table
    Dim RsStkAdj As ADODB.Recordset
    Dim XlsConn As New ADODB.Connection
    
    Dim mCount  As Long
    Dim mQRY$, mDocIdIssue$, mDocIdReceive$, mVTypeIssue$, mVPrefixIssue$, mVDate$, mVTypeReceive$, mVPrefixReceive$
    
                        
                    
    mVTypeIssue = "SYIAD"
    mVTypeReceive = "SXRAD"
    mVDate = DateAdd("D", -1, PubStartDate)
    mVPrefixIssue = G_FaCn.Execute("Select " & xIsNull("Prefix", "") & "  " & _
                              "From Voucher_Prefix Where V_Type = 'SYIAD'").Fields(0).Value
    mVPrefixReceive = G_FaCn.Execute("Select " & xIsNull("Prefix", "") & " " & _
                              "From Voucher_Prefix Where V_Type = 'SXRAD'").Fields(0).Value
                              
    mVNoIssue = G_FaCn.Execute("Select " & Val("Start_Srl_No") & " + 1 From Voucher_Prefix " & _
                          "Where V_Type='" & mVTypeIssue & "' And Date_From = " & ConvertDate(DateAdd("D", 1, CDate(mVDate))) & "").Fields(0).Value
                                
    mVNoReceive = G_FaCn.Execute("Select " & Val("Start_Srl_No") & " + 1 From Voucher_Prefix " & _
                          "Where V_Type='" & mVTypeReceive & "' And Date_From = " & ConvertDate(DateAdd("D", 1, CDate(mVDate))) & "").Fields(0).Value
                                
    mVNoIssue = G_FaCn.Execute("Select isNull(Max(V_No),0)  + 1 From Sp_Stock " & _
                          "Where V_Type='" & mVTypeIssue & "' ").Fields(0).Value
                                
    mVNoReceive = G_FaCn.Execute("Select isNull(Max(V_No),0) + 1 From Sp_Stock " & _
                          "Where V_Type='" & mVTypeReceive & "' ").Fields(0).Value
                                
                
                
    mDocIdIssue = PubDivCode + PubSiteCode & PubSiteCode + Space(5 - Len(mVTypeIssue)) + _
                mVTypeIssue + Space(5 - Len(CStr(mVPrefixIssue))) + mVPrefixIssue + Space(8 - Len(CStr(mVNoIssue))) + CStr(mVNoIssue)
                
    mDocIdReceive = PubDivCode + PubSiteCode & PubSiteCode + Space(5 - Len(mVTypeReceive)) + _
                mVTypeReceive + Space(5 - Len(CStr(mVPrefixReceive))) + mVPrefixReceive + Space(8 - Len(CStr(mVNoReceive))) + CStr(mVNoReceive)
                
    
    
    
    G_FaCn.BeginTrans
    GCn.BeginTrans
        Dim Rst As ADODB.Recordset
        Dim RST1 As ADODB.Recordset
        Set Rst = GCn.Execute("Select * From City")
        Set RST1 = GCn.Execute("Select * From (" & Rst.Source & ") As City")
    
        XlsConn.CursorLocation = adUseClient
        XlsConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Txt & ";Extended Properties=Excel 8.0"
        XlsConn.Open
        GCn.CommandTimeout = 120
        GCn.Execute "Delete From Sp_Stock Where V_Date=" & ConvertDate(Txt1(0)) & " And U_Name='StkAdj'"
        Set RsStkAdj = XlsConn.Execute("SELECT Part_No, Max(Part_Name) As Part_Name, Sum(TB_Qty) As TB_Qty, Sum(TP_Qty) As TP_Qty, Max(Bin) As Bin FROM [Sheet1$] IN '" & Txt & "' 'EXCEL 8.0;' Group By Part_No ")
        GCn.Execute "Create Table StkAdj(Part_No VarChar(22), Part_Name VarChar(40), TB_Qty Numeric(18,3), TP_Qty Numeric(18,3), Bin VarChar(15))"
        If RsStkAdj.RecordCount > 0 Then
            Do Until RsStkAdj.EOF
                GCn.Execute "Insert Into StkAdj (Part_No, Part_Name, TB_Qty, TP_Qty, Bin) Values ('" & XNull(RsStkAdj!Part_No) & "', '" & XNull(RsStkAdj!Part_Name) & "', " & VNull(RsStkAdj!TB_Qty) & ", " & VNull(RsStkAdj!TP_Qty) & ", '" & XNull(RsStkAdj!Bin) & "')"
                
                RsStkAdj.MoveNext
            Loop
        End If
        'XlsConn.Execute ("SELECT Part_No, Max(Part_Name) As Part_Name, Sum(TB_Qty) As TB_Qty, Sum(TP_Qty) As TP_Qty, Max(Bin) As Bin Into StkAdj FROM [Sheet1$] IN '" & txt & "' 'EXCEL 8.0;' Group By Part_No ")
        
        'GCn.Execute ("Select S.Part_No, S.Part_Name InTo Diff IN '" & txt & "' 'EXCEL 8.0;'  From StkAdj S Where S.Part_No Not In (Select Part_No From Part)")
        GCn.Execute "UPDATE Part SET Part.Bin_Loca = StkAdj.Bin From StkAdj WHERE Part.Part_No = StkAdj.Part_No "
        
        
                                 
                                 
        Set RsAdj = New ADODB.Recordset
        With RsAdj
            .Fields.Append "Part_No", adVarChar, 22, adFldIsNullable
            .Fields.Append "ReqTB", adDouble, 12, adFldIsNullable
            .Fields.Append "ReqTP", adDouble, 12, adFldIsNullable
            .Fields.Append "Qty", adDouble, 12, adFldIsNullable
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .Open
        End With
                                 
                                 
                                 
       
                                 
        If ChkTB.Value = 1 And ChkTP.Value = 1 Then
            If PubBackEnd = "A" Then
                mQRY = "Select Part_No, " & _
                        "TB_Qty-(SELECT " & vIsNull("sum(qty_rec)", "0") & " - " & vIsNull("Sum(Qty_Iss)", "0") & " + " & vIsNull("Sum(Qty_Ret)", "0") & " " & _
                                "From sp_stock With (NoLOck)" & _
                                "WHERE ciif(v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & ",V_Type='SXAO',cIIF(V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(Txt1(0)) & ",V_Type<>'SXAO')) " & _
                                "And Part_No=StkAdj.Part_No And Tax_Yn=1) As ReqTB, " & _
                        "TP_Qty-(SELECT " & vIsNull("sum(qty_rec)", "0") & "- " & vIsNull("Sum(Qty_Iss)", "0") & " +" & vIsNull("Sum(Qty_Ret)", "0") & " " & _
                                "From sp_stock With (NoLock)" & _
                                "WHERE iif(v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & ",V_Type='SXAO',IIF(V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(Txt1(0)) & ",V_Type<>'SXAO')) " & _
                                "And Part_No=StkAdj.Part_No And Tax_Yn=0) As ReqTP, " & _
                        "TB_Qty+TP_Qty As Qty  " & _
                        "From StkAdj With (NoLock)"
            Else
                mQRY = "Select Part_No, " & _
                        "TB_Qty-(SELECT " & vIsNull("sum(qty_rec)", "0") & " - " & vIsNull("Sum(Qty_Iss)", "0") & " + " & vIsNull("Sum(Qty_Ret)", "0") & " " & _
                                "From sp_stock With (NoLock)" & _
                                "WHERE (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(Txt1(0)) & "", "'SXAO'") & ") " & _
                                "And Part_No=StkAdj.Part_No And Tax_Yn=1) As ReqTB, " & _
                        "TP_Qty-(SELECT " & vIsNull("sum(qty_rec)", "0") & "- " & vIsNull("Sum(Qty_Iss)", "0") & " +" & vIsNull("Sum(Qty_Ret)", "0") & " " & _
                                "From sp_stock With (NoLock)" & _
                                "WHERE (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(Txt1(0)) & "", "'SXAO'") & ") " & _
                                "And Part_No=StkAdj.Part_No And Tax_Yn=0) As ReqTP, " & _
                        "TB_Qty+TP_Qty As Qty  " & _
                        "From StkAdj With (NoLock)"
            End If
            Set RsTemp = GCn.Execute(mQRY)
            
            Fill_StkAdj RsAdj, RsTemp
            
            If ChkAssumeZero.Value = 1 Then
'                mQRY = mQRY & "Union All Select Part_No, " & _
'                     "0-IsNull((Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock S Where (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(Txt1(0)) & "", "'SXAO'") & ") And  S.Part_No=SS.Part_No And Tax_Yn=1),0) As ReqTB, " & _
'                     "0-IsNull((Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock S Where (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(Txt1(0)) & "", "'SXAO'") & ") And  S.Part_No=SS.Part_No And Tax_Yn=0),0) As ReqTP, " & _
'                     "0 As Qty From Sp_Stock SS Where SS.Part_No Not In (Select Part_No From StkAdj) Group By SS.Part_No"
                     
                mQRY = "Select Part_No, " & _
                     "0-IsNull((Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock S With (NoLock) Where (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(Txt1(0)) & "", "'SXAO'") & ") And  S.Part_No=SS.Part_No And Tax_Yn=1),0) As ReqTB, " & _
                     "0-IsNull((Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock S With (NoLock) Where (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(Txt1(0)) & "", "'SXAO'") & ") And  S.Part_No=SS.Part_No And Tax_Yn=0),0) As ReqTP, " & _
                     "0 As Qty From Sp_Stock SS With (NoLock) Where SS.Part_No Not In (Select Part_No From StkAdj With (NoLock)) Group By SS.Part_No"
                
                Set RsTemp = GCn.Execute(mQRY)
                Fill_StkAdj RsAdj, RsTemp
            End If
            Set RsTemp = RsAdj
        ElseIf ChkTB.Value = 1 And ChkTP.Value = 0 Then
            Set RsTemp = GCn.Execute("Select Part_No, " & _
                                     "TB_Qty-(SELECT " & vIsNull("sum(qty_rec)", "0") & "- " & vIsNull("Sum(Qty_Iss)", "0") & " +" & vIsNull("Sum(Qty_Ret)", "0") & " " & _
                                             "From sp_stock With (Nolock)" & _
                                             "WHERE (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(Txt1(0)) & "", "'SXAO'") & ") " & _
                                             "And Part_No=StkAdj.Part_No) As ReqTB, " & _
                                     "0 As ReqTP, " & _
                                     "TB_Qty+TB_Qty As Qty " & _
                                     "From StkAdj With (NoLOck)")
            
        ElseIf ChkTB.Value = 0 And ChkTP.Value = 1 Then
            Set RsTemp = GCn.Execute("Select Part_No, " & _
                                     "0 As ReqTB, " & _
                                     "TP_Qty-(SELECT " & vIsNull("sum(qty_rec)", "0") & " - " & vIsNull("Sum(Qty_Iss)", "0") & " +" & vIsNull("Sum(Qty_Ret)", "0") & " " & _
                                             "From sp_stock " & _
                                             "WHERE (V_Type= " & cIIF("v_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "", "'SXAO'") & " Or V_Type <> " & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(Txt1(0)) & "", "'SXAO'") & ") " & _
                                             "And Part_No=StkAdj.Part_No) As ReqTP, " & _
                                     "TB_Qty+TB_Qty As Qty " & _
                                     "From StkAdj")
        End If
                                         
        RsTemp.MoveFirst
        Do Until RsTemp.EOF
            Debug.Print RsTemp(0) & vbTab & RsTemp(1) & vbTab & RsTemp(2) & vbTab & RsTemp(3)
            RsTemp.MoveNext
        Loop
                                                                                  
                                                                                  
                                 
        If RsTemp.RecordCount > 0 Then
            mSrlIssue = 1
            mSrlReceive = 1
            mCount = 0
            ProgressBar1.Value = 0
            ProgressBar1.Visible = True
        Else
            MsgBox "No Records Found to Adjust"
            Exit Sub
        End If

        
        
        RsTemp.MoveFirst
        Do Until RsTemp.EOF
            If RsTemp!Part_No = "0002013719" Then
                MsgBox ""
            End If
            If ChkTB.Value = 1 Then
                'If IsNull(rsTemp!ReqTB) Then MsgBox rsTemp!Part_No
                If VNull(RsTemp!ReqTB) > 0 Then
                    GCn.Execute "Insert Into SP_Stock(" _
                        & "DocID,Srl_No,V_Type,V_No,V_Date,Site_Code," _
                        & "Party_Code,Remark,Part_No, Qty_Rec,Tax_YN," _
                        & "MRP_YN,Rate,MRP_Rate,Amount,Net_Amt," _
                        & "Godown,Purpose, U_Name,U_EntDt,U_AE,V_Rate) " _
                        & "Values(" _
                        & "'" & mDocIdReceive & "'," & mSrlReceive & ",'" & mVTypeReceive & "'," & mVNoReceive & "," & ConvertDate(Txt1(0)) & ",'" & PubSiteCode & PubSiteCode & "'," _
                        & "'','','" & RsTemp!Part_No & "'," & Abs(RsTemp!ReqTB) & ",1," _
                        & "1,0,0,0,0," _
                        & "'" & PubSprCounterGodown & "','','StkAdj'," & ConvertDate(PubServerDate) & ",'A',0)"
                        
                        mSrlReceive = mSrlReceive + 1
                        
                ElseIf VNull(RsTemp!ReqTB) < 0 Then
                    GCn.Execute "Insert Into SP_Stock(" _
                        & "DocID,Srl_No,V_Type,V_No,V_Date,Site_Code," _
                        & "Party_Code,Remark,Part_No, Qty_Iss,Tax_YN," _
                        & "MRP_YN,Rate,MRP_Rate,Amount,Net_Amt," _
                        & "Godown,Purpose, U_Name,U_EntDt,U_AE,V_Rate) " _
                        & "Values(" _
                        & "'" & mDocIdIssue & "'," & mSrlIssue & ",'" & mVTypeIssue & "'," & mVNoIssue & "," & ConvertDate(Txt1(0)) & ",'" & PubSiteCode & PubSiteCode & "'," _
                        & "'','','" & RsTemp!Part_No & "'," & Abs(RsTemp!ReqTB) & ",1," _
                        & "1,0,0,0,0," _
                        & "'" & PubSprCounterGodown & "','','StkAdj'," & ConvertDate(PubServerDate) & ",'A',0)"
                        
                        mSrlIssue = mSrlIssue + 1
                End If
            End If
            
            
            If ChkTP.Value = 1 Then
                If VNull(RsTemp!ReqTP) > 0 Then
                    GCn.Execute "Insert Into SP_Stock(" _
                        & "DocID,Srl_No,V_Type,V_No,V_Date,Site_Code," _
                        & "Party_Code,Remark,Part_No, Qty_Rec,Tax_YN," _
                        & "MRP_YN,Rate,MRP_Rate,Amount,Net_Amt," _
                        & "Godown,Purpose, U_Name,U_EntDt,U_AE,V_Rate) " _
                        & "Values(" _
                        & "'" & mDocIdReceive & "'," & mSrlReceive & ",'" & mVTypeReceive & "'," & mVNoReceive & "," & ConvertDate(Txt1(0)) & ",'" & PubSiteCode & PubSiteCode & "'," _
                        & "'','','" & RsTemp!Part_No & "'," & RsTemp!ReqTP & ",0," _
                        & "1,0,0,0,0," _
                        & "'" & PubSprCounterGodown & "','','StkAdj'," & ConvertDate(PubServerDate) & ",'A',0)"
                        
                        mSrlReceive = mSrlReceive + 1
                ElseIf VNull(RsTemp!ReqTP) < 0 Then
                    GCn.Execute "Insert Into SP_Stock(" _
                        & "DocID,Srl_No,V_Type,V_No,V_Date,Site_Code," _
                        & "Party_Code,Remark,Part_No, Qty_Iss,Tax_YN," _
                        & "MRP_YN,Rate,MRP_Rate,Amount,Net_Amt," _
                        & "Godown,Purpose, U_Name,U_EntDt,U_AE,V_Rate) " _
                        & "Values(" _
                        & "'" & mDocIdIssue & "'," & mSrlIssue & ",'" & mVTypeIssue & "'," & mVNoIssue & "," & ConvertDate(Txt1(0)) & ",'" & PubSiteCode & PubSiteCode & "'," _
                        & "'','','" & RsTemp!Part_No & "'," & Abs(RsTemp!ReqTP) & ",0," _
                        & "1,0,0,0,0," _
                        & "'" & PubSprCounterGodown & "','','StkAdj'," & ConvertDate(PubServerDate) & ",'A',0)"
                        
                        mSrlIssue = mSrlIssue + 1
                End If
            End If
            


            mCount = mCount + 1
            If mCount < RsTemp.RecordCount Then
                ProgressBar1.Value = mCount * 100 / RsTemp.RecordCount
            End If
            
            
            RsTemp.MoveNext
        Loop
        
        
            
                                    
        
        G_FaCn.Execute "UPDATE Voucher_Prefix SET Start_Srl_No = Start_Srl_No + 1 " & _
                       "WHERE V_Type = '" & mVTypeIssue & "'"
        G_FaCn.Execute "UPDATE Voucher_Prefix SET Start_Srl_No = Start_Srl_No + 1 " & _
                       "WHERE V_Type = '" & mVTypeReceive & "'"
                       
                       
        GCn.Execute "Drop Table StkAdj"
        
        
    G_FaCn.CommitTrans
    GCn.CommitTrans
    
    MsgBox " # Stock Adjustment Done # "
    ProgressBar1.Visible = False
    
    Set RsTemp = Nothing
Exit Sub




DispErr:
    
    G_FaCn.RollbackTrans
    GCn.RollbackTrans
    MsgBox err.Description
End Sub

Sub Fill_StkAdj(RsFillIn As ADODB.Recordset, RsFillFrom As ADODB.Recordset)
    Do Until RsFillFrom.EOF
        With RsFillIn
            .AddNew
            Debug.Print XNull(RsFillFrom!Part_No)
            !Part_No = XNull(RsFillFrom!Part_No)
            !ReqTB = VNull(RsFillFrom!ReqTB)
            !ReqTP = VNull(RsFillFrom!ReqTP)
            !Qty = VNull(RsFillFrom!Qty)
            .Update
        End With
        RsFillFrom.MoveNext
    Loop
End Sub


Private Sub DispText()

    Me.CAPTION = "Physical Stock Updation"
    LblSelect = "Select MS Excel File"
    
    LblCol(0) = "1. Part_No"
    LblCol(1) = "2. Part_Name"
    LblCol(2) = "3. Tb_Qty"
    LblCol(3) = "4. Bin"
    LblCol(4) = "5. Tp_Qty"
    
    'FrameColumn.width = 4000 * ScreenFactor
    
    Txt1(0) = ""
End Sub

Private Sub Txt1_Validate(Index As Integer, Cancel As Boolean)
    Txt1(0) = RetDate(Txt1(0))
End Sub


