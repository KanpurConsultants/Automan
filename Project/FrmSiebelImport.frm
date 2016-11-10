VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
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
      Left            =   0
      TabIndex        =   6
      Top             =   1830
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   2565
         Width           =   1185
      End
      Begin VB.TextBox TxtShow 
         Appearance      =   0  'Flat
         Height          =   915
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "FrmSiebelImport.frx":0000
         Top             =   1995
         Width           =   8865
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FgridErr 
         Height          =   1620
         Left            =   120
         TabIndex        =   11
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11595
      Begin VB.CommandButton CmdImport 
         Caption         =   "Vehicle Purchase"
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
         Index           =   9
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Index           =   8
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   285
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Supplier Payment"
         Height          =   525
         Index           =   4
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4200
         Width           =   1320
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Spare Sale Return"
         Height          =   540
         Index           =   7
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4200
         Width           =   1425
      End
      Begin MSComctlLib.ProgressBar Prg 
         Height          =   270
         Left            =   165
         TabIndex        =   3
         Top             =   1350
         Visible         =   0   'False
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Private Sub CmdImport_Click(Index As Integer)
' '' On Error GoTo Eloop
'Dim MasterCode As String, DocId As String, mV_Type As String, mPartyCode As String, mForm_Code As String
'Dim mDebitAc As String, mMfgMonth As String, mMfgYear As String, mColourCode As String, mColourName As String, mGodownCode As String
'Dim mTaxPer As Double, mDeductionCode As String, mAdditionCode As String
'Dim mLength1 As Integer, mLength2 As Integer, mTaxOnDelivery As Boolean
'Dim EditFlag As Boolean
'Dim RsX As adodb.Recordset
'Dim xDocId$
'
'
'
'    GCn.BeginTrans
'    CopyCnt = 0
'    ErrorCnt = 0
'    Set RsNew = New adodb.Recordset
'    RsNew.CursorLocation = adUseClient
'    RsNew.Open "Select * from Veh_Purch1", GCn, adOpenDynamic, adLockOptimistic
'
'    Set RsNew1 = New adodb.Recordset
'    RsNew1.CursorLocation = adUseClient
'    RsNew1.Open "Select * from Veh_Purch2", GCn, adOpenDynamic, adLockOptimistic
'
'    Set RsTemp = New adodb.Recordset
'    RsTemp.CursorLocation = adUseClient
'    RsTemp.Open "Select * from Veh_Stock", GCn, adOpenDynamic, adLockOptimistic
'
'
'    mV_Type = "V_PB"
'
'    CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from Veh_Purch1 where Left(DocID,1)='" & PubDivCode & "' and " & cMID("DocID", "2", "1") & "='" & PubSiteCode & "' and V_Type='" & mV_Type & "'").Fields(0).Value
'    Do Until Master.EOF
'        If IsNull(StringPass(Master.Fields("Invoice_No"))) Or StringPass(Master.Fields("Invoice_No")) = "" Then GoTo MyNextRecord
'        EditFlag = False
'
''Modi Arpit Because Telco Apply Vat After Some Days
''        If GCn.Execute("Select PBill_No from Veh_Purch1 where Pbill_No='" & left(StringPass(Master.Fields("Invoice_no")), 10) & "'").RecordCount > 0 Then
''            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Telco Invoice No. Already Exist in Automan")
''            GoTo MyNextRecord
''        End If
'
'        If GCn.Execute("Select PBill_No from Veh_Purch1 where Pbill_No='" & left(StringPass(Master.Fields("Invoice_no")), 10) & "'").RecordCount > 0 Then
'            EditFlag = True
'        End If
''Modi End
'
'        If StringPass(Master.Fields("Supplier_Name")) = "" Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Supplier Name is Empty")
'            GoTo MyNextRecord
'        Else
'            If ErrorGCN.Execute("select AutomanAcCode from AccountConversion where Type='Vehicle Purchase' and SiebelAc='" & StringPass(Master.Fields("Supplier_Name")) & "'").RecordCount > 0 Then
'                mPartyCode = ErrorGCN.Execute("select AutomanAcCode from AccountConversion where Type='Vehicle Purchase' and SiebelAc='" & StringPass(Master.Fields("Supplier_Name")) & "'").Fields(0).Value
'            Else
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master!Supplier_Name), "Vehicle Purchase", "Automan A/c Code is not Defined in AccountConversionTable for This Supplier")
'                GoTo MyNextRecord
'            End If
'        End If
'
'        If IsNull(Master!Invoice_Date) Or Master!Invoice_Date = "" Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Telco Invoice Date is Empty")
'            GoTo MyNextRecord
'        End If
'
'        If IsNull(StringPass(Master!Godown)) Or StringPass(Master!Godown) = "" Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Godown Name is Empty")
'            GoTo MyNextRecord
'        End If
'
'        If IsNull(StringPass(Master!VC_Number)) Or StringPass(Master!VC_Number) = "" Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "VC_Number is Empty")
'            GoTo MyNextRecord
'        Else
'            If GCn.Execute("Select Model from Model where Model='" & StringPass(Master!VC_Number) & "'").RecordCount = 0 Then
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "VC_Number is not Exist in Model Master")
'                GoTo MyNextRecord
'            End If
'        End If
'
'        If IsNull(StringPass(Master!Chassis_No)) Or StringPass(Master!Chassis_No) = "" Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Chassis Number is Empty")
'            GoTo MyNextRecord
'        End If
'
'        If GCn.Execute("Select ChassisNo from Veh_Stock where ChassisNo='" & StringPass(Master.Fields("Chassis_No")) & "'").RecordCount > 0 Then
'            If EditFlag = False Then
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master!Chassis_No, "Vehicle Purchase", "This Chassis No. Already Exist in Automan")
'                GoTo MyNextRecord
'            End If
'        End If
'
'
'        If IsNull(StringPass(Master!Narration)) Or StringPass(Master!Narration) = "" Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Narration Field is Empty (Engine Number)")
'            GoTo MyNextRecord
'        End If
'
'
'        If Len(StringPass(Master.Fields("Chassis_No"))) = 17 Then
'            If GCn.Execute("Select Name from Chas_Mth where Month_CD='" & mID(StringPass(Master!Chassis_No), 12, 1) & "'").RecordCount = 0 Then
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Chassis Mfg. Month Name is not defined in Chas_Mth Table")
'                GoTo MyNextRecord
'            Else
'                mMfgMonth = GCn.Execute("Select Name from Chas_Mth where Month_CD='" & mID(StringPass(Master!Chassis_No), 12, 1) & "'").Fields(0).Value
'            End If
'        ElseIf Len(StringPass(Master.Fields("Chassis_No"))) > 17 Then
'            mMfgMonth = Format(Master.Fields("Invoice_Date"), "MMMM")
'        Else
'            If GCn.Execute("Select Name from Chas_Mth where Month_CD='" & mID(StringPass(Master!Chassis_No), 7, 1) & "'").RecordCount = 0 Then
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Chassis Mfg. Month Name is not defined in Chas_Mth Table")
'                GoTo MyNextRecord
'            Else
'                mMfgMonth = GCn.Execute("Select Name from Chas_Mth where Month_CD='" & mID(StringPass(Master!Chassis_No), 7, 1) & "'").Fields(0).Value
'            End If
'        End If
'
'        If Len(StringPass(Master.Fields("Chassis_No"))) = 17 Then
''''            Select Case Val(Mid(StringPass(Master!Chassis_No), 10, 1))
''''                Case 9
''''                    mMfgYear = "2009"
''''                Case 0
''''                    mMfgYear = "2010"
''''                Case 1
''''                    mMfgYear = "2011"
''''                Case 2
''''                    mMfgYear = "2012"
''''                Case 3
''''                    mMfgYear = "2013"
''''                Case 4
''''                    mMfgYear = "2014"
''''                Case 5
''''                    mMfgYear = "2015"
''''                Case 6
''''                    mMfgYear = "2016"
''''                Case 7
''''                    mMfgYear = "2017"
''''                Case 8
''''                    mMfgYear = "2018"
''''            End Select
'
'        Select Case (mID(StringPass(Master!Chassis_No), 10, 1))
'                Case "9"
'                    mMfgYear = "2009"
'                Case "0"
'                    mMfgYear = "2010"
'                Case "B"
'                    mMfgYear = "2011"
'                Case "C"
'                    mMfgYear = "2012"
'                Case "D"
'                    mMfgYear = "2013"
'                Case "E"
'                    mMfgYear = "2014"
'                Case "F"
'                    mMfgYear = "2015"
'                Case "G"
'                    mMfgYear = "2016"
'                Case "H"
'                    mMfgYear = "2017"
'                Case "I"
'                    mMfgYear = "2018"
'            End Select
'        ElseIf Len(StringPass(Master.Fields("Chassis_No"))) > 17 Then
'            mMfgYear = Format(Master.Fields("Invoice_Date"), "YYYY")
'        Else
'            If GCn.Execute("Select Name from Chas_Yr where Year_Cd='" & mID(StringPass(Master!Chassis_No), 8, 2) & "'").RecordCount = 0 Then
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Chassis Mfg. Year Name is not defined in Chas_YR Table")
'                GoTo MyNextRecord
'            Else
'                mMfgYear = GCn.Execute("Select Name from Chas_Yr where Year_Cd='" & mID(StringPass(Master!Chassis_No), 8, 2) & "'").Fields(0).Value
'            End If
'        End If
'
'        If GCn.Execute("Select God_Code from Godown where Left(God_Name,20)='" & left(StringPass(Master.Fields("Godown")), 20) & "' and Appli_For=1").RecordCount = 0 Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Godown Name not found Godown Master of Automan")
'            GoTo MyNextRecord
'        Else
'            mGodownCode = GCn.Execute("Select God_Code from Godown where Left(God_Name,20)='" & left(StringPass(Master.Fields("Godown")), 20) & "' and Appli_For=1").Fields(0).Value
'        End If
'
'        mColourCode = GCn.Execute("Select Col_Code from Model where Model='" & StringPass(Master.Fields("VC_Number")) & "'").Fields(0).Value
'        If mColourCode = "" Then
'            mColourCode = ErrorGCN.Execute("Select DefaultColourCode from Enviro").Fields(0).Value
'        End If
'        mColourName = ""
''        If GCn.Execute("Select Col_Code from ColMast where Col_Code='" & mColourCode & "'").RecordCount = 0 Then
''            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Colour Name not found Colour Master of Automan")
''            GoTo MyNextRecord
''        Else
''            mColourName = GCn.Execute("Select Col_Desc from ColMast where Col_Code='" & mColourCode & "'").Fields(0).Value
''        End If
'
'        If GCn.Execute("Select Col_Code from ColMast where Col_Code='" & mColourCode & "'").RecordCount > 0 Then
'            mColourName = GCn.Execute("Select Col_Desc from ColMast where Col_Code='" & mColourCode & "'").Fields(0).Value
'        End If
'
'        If eVal(Master.Fields("TAX CST")) > 0 Then
'            mForm_Code = ErrorGCN.Execute("Select VehicleCstPurchFormCode from Enviro").Fields(0).Value
'            mTaxPer = GCn.Execute("Select Tax_Per from TaxForms Where Form_Code='" & mForm_Code & "'").Fields(0).Value
'        Else
'            mForm_Code = ErrorGCN.Execute("Select VehiclePurchFormCode from Enviro").Fields(0).Value
'            mTaxPer = ErrorGCN.Execute("Select VehiclePurchTaxPer from Enviro").Fields(0).Value
'        End If
'
'        mDeductionCode = ErrorGCN.Execute("Select ChassisDiscountItemCode from Enviro").Fields(0).Value
'        mAdditionCode = ErrorGCN.Execute("Select ChassisTransportItemCode from Enviro").Fields(0).Value
'        mTaxOnDelivery = ErrorGCN.Execute("Select TaxOnDeliveryCharges from Enviro").Fields(0).Value
'
'        mDebitAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc Where Form_Code='" & mForm_Code & "' ").Fields(0).Value
'        Dim mShortYear As String
'        If Month(Master.Fields("Invoice_Date")) > 3 Then
'            mShortYear = Right(Format(Master.Fields("Invoice_Date"), "yy"), 1) & Right(Val(Format(Master.Fields("Invoice_Date"), "yy")) + 1, 1)
'        Else
'            mShortYear = Right(Val(Format(Master.Fields("Invoice_Date"), "yy")) - 1, 1) & Right(Format(Master.Fields("Invoice_Date"), "yy"), 1)
'        End If
'
'        'DocId = PubDivCode & PubSiteCode & PubSiteCode & " " & mV_Type & "SBL" & Format(Master!Invoice_Date, "yy") & Right("00000000" & CodeCnt, 8)
'        DocId = PubDivCode & PubSiteCode & PubSiteCode & " " & mV_Type & "SBL" & mShortYear & Right("00000000" & CodeCnt, 8)
'
'
'        '' Calculation of Amount
'        Dim mTot_Amt As Double, mTax_Amt As Double, mMisc_Amt As Double
'        Dim mDeduction As Double, mAddition As Double, mAmount As Double
'
'        mTot_Amt = 0: mTax_Amt = 0: mMisc_Amt = 0
'        mDeduction = 0: mAddition = 0: mAmount = 0
'
'        If Master.Fields("Chassis_No") = "445051HRZY00517" Then
'            MsgBox ""
'        End If
'
'        If UCase(left(PubComp_Name, 3)) = "LMP" Then
'            mTot_Amt = Master!Value
'
'            If mTaxOnDelivery Then
'                mMisc_Amt = 0
'            Else
'                mMisc_Amt = VNull(Master.Fields("Delivery Charges"))
'            End If
'
'            If IsNull(Master.Fields("VatTax")) Or Master.Fields("VatTax") = "" Then
'                mTax_Amt = Round((mTot_Amt) * mTaxPer / (100 + mTaxPer), 2)     ''- mMisc_Amt
'            Else
'                mTax_Amt = Val(Master.Fields("VatTax"))
'            End If
'
'            If mTaxOnDelivery Then
'                mAddition = VNull(Master.Fields("Delivery Charges"))
'            Else
'                mAddition = 0
'            End If
'
'            mAmount = mTot_Amt + mDeduction - (mMisc_Amt + mTax_Amt + mAddition)
'        Else
'            mTot_Amt = Master!Value
'
'            If mTaxOnDelivery Then
'                mMisc_Amt = 0
'            Else
'                mMisc_Amt = VNull(Master.Fields("Delivery Charges"))
'            End If
'            If eVal(Master.Fields("Tax Cst")) > 0 Then
'                mTax_Amt = eVal(Master.Fields("Tax Cst"))
'            Else
'                If IsNull(Master.Fields("VatTax")) Or Master.Fields("VatTax") = "" Then
'                    mTax_Amt = Round((mTot_Amt) * mTaxPer / (100 + mTaxPer), 2)     ''- mMisc_Amt
'                Else
'                    mTax_Amt = Val(Master.Fields("VatTax"))
'                End If
'            End If
'
'            If mTaxOnDelivery Then
'                mAddition = VNull(Master.Fields("Delivery Charges"))
'            Else
'                mAddition = 0
'            End If
'            mDeduction = VNull(Master.Fields("Total Discount"))
'            mAmount = mTot_Amt + mDeduction - (mMisc_Amt + mTax_Amt + mAddition)
'        End If
'
'
'        If EditFlag = True Then
'            'ArpitStart
'            GCn.Execute "Update Veh_Purch1 Set Amount = " & mAmount & ", Tot_Amount = " & mTot_Amt & ", " & _
'                        "Tax_Per = " & mTaxPer & ", Tax_Amt = " & mTax_Amt & ", Addition = " & mAddition & ", " & _
'                        "Deduction = " & mDeduction & ", Misc_Amt = " & mMisc_Amt & ", U_EntDt = " & ConvertDate(date) & " " & _
'                        "Where DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(Master!Chassis_No) & "' )"
'
'
'
'            Set RsX = GCn.Execute("Select Pur_DocId From Veh_Stock Where ChassisNo = '" & Master!Chassis_No & "'")
'            If RsX.RecordCount > 0 Then
'                GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & XNull(RsX(0)) & "' And Trn_Type='D'"
'                GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & XNull(RsX(0)) & "' And Trn_Type='A'"
'            End If
'
'            If mDeduction > 0 Then
'                Set RsX = GCn.Execute("Select Pur_DocId From Veh_Stock Where ChassisNo = '" & Master!Chassis_No & "'")
'                If RsX.RecordCount > 0 Then xDocId = XNull(RsX!Pur_DocId)
'
'                If GCn.Execute("Select DocId From Veh_Purch2 Where DocId = '" & xDocId & "'").RecordCount > 0 Then
'                    'GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & xDocId & "' And Trn_Type='D'"
'                    'GCn.Execute "Update Veh_Purch2  Set Rate = " & mDeduction & " " & _
'                                "Where DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(Master!Chassis_No) & "') And Trn_Type='D'"
'                End If
'                    With RsNew1
'                        .AddNew
'                        !DocId = xDocId
'                        !Srl_No = 1
'                        !Site_Code = PubSiteCode & PubSiteCode
'                        !V_Type = mV_Type
'                        !V_NO = CodeCnt
'                        !Trn_Type = "D"
'                        !PROD_CODE = mDeductionCode
'                        !Qty = 1
'                        !Rate = mDeduction
'
'                        !U_Name = "Siebel"
'                        !U_EntDt = Format(PubLoginDate, "Short Date")
'                        !U_AE = "A"
'                        .Update
'                    End With
'                'End If
'            End If
'
'            If mAddition > 0 Then
'                Set RsX = GCn.Execute("Select Pur_DocId From Veh_Stock Where ChassisNo = '" & Master!Chassis_No & "'")
'                If RsX.RecordCount > 0 Then xDocId = XNull(RsX!Pur_DocId)
'
'                If GCn.Execute("Select DocId From Veh_Purch2 Where DocId = '" & xDocId & "'").RecordCount > 0 Then
'                    'GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & xDocId & "' And Trn_Type='A'"
'                    'GCn.Execute "Update Veh_Purch2 Set Rate = " & mAddition & " " & _
'                                "Where DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(Master!Chassis_No) & "') And Trn_Type='A'"
'                End If
'                    With RsNew1
'                        .AddNew
'                        !DocId = xDocId
'                        !Srl_No = 2
'                        !Site_Code = PubSiteCode & PubSiteCode
'                        !V_Type = mV_Type
'                        !V_NO = CodeCnt
'                        !Trn_Type = "A"
'                        !PROD_CODE = mAdditionCode
'                        !Qty = 1
'                        !Rate = mAddition
'
'                        !U_Name = "Siebel"
'                        !U_EntDt = Format(PubLoginDate, "Short Date")
'                        !U_AE = "A"
'                        .Update
'                    End With
'                'End If
'            End If
'
'            GCn.Execute "Update Veh_Stock Set Rate = " & mAmount & ", VRate = " & mTot_Amt & " " & _
'                        "Where Pur_DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(Master!Chassis_No) & "' )"
'
'            EditFlag = False
'            'ArpitEnd
'        Else
'            'Insert New Rec
'            With RsNew
'                .AddNew
'                !DocId = DocId
'                !DocIDHelp = Replace(DocId, " ", "")
'                !Site_Code = PubSiteCode & PubSiteCode
'                !V_Type = mV_Type
'                !V_NO = CodeCnt
'                !V_DATE = MakeDate(Master!Invoice_Date)
'                !PartyCode = mPartyCode
'                !PBill_No = Master!Invoice_No
'                !Pbill_Date = MakeDate(Master!Invoice_Date)
'                !BMS_Category = ErrorGCN.Execute("Select DefaultBMSCategory from Enviro").Fields(0).Value
'                !DueDate = MakeDate(Master!Invoice_Date)
'                !Gate = ""
'                !GateDate = MakeDate(Master!Invoice_Date)
'                !Form_Code = mForm_Code
'                !Amount = mAmount
'                !Addition = mAddition
'                !Deduction = mDeduction
'                !Exsice = 0
'                !Tax_Per = mTaxPer
'                !Tax_Amt = mTax_Amt
'                !Misc_Amt = mMisc_Amt
'                !Tot_AMOUNT = mTot_Amt
'                !DrAcCode = mDebitAc
'
'                !U_Name = "Siebel"
'                !U_EntDt = Format(PubLoginDate, "Short Date")
'                !U_AE = "A"
'                .Update
'            End With
'
'            If mDeduction > 0 Then
'                With RsNew1
'                    .AddNew
'                    !DocId = DocId
'                    !Srl_No = 1
'                    !Site_Code = PubSiteCode & PubSiteCode
'                    !V_Type = mV_Type
'                    !V_NO = CodeCnt
'                    !Trn_Type = "D"
'                    !PROD_CODE = mDeductionCode
'                    !Qty = 1
'                    !Rate = mDeduction
'
'                    !U_Name = "Siebel"
'                    !U_EntDt = Format(PubLoginDate, "Short Date")
'                    !U_AE = "A"
'                    .Update
'                End With
'            End If
'
'            If mAddition > 0 Then
'                With RsNew1
'                    .AddNew
'                    !DocId = DocId
'                    !Srl_No = 2
'                    !Site_Code = PubSiteCode & PubSiteCode
'                    !V_Type = mV_Type
'                    !V_NO = CodeCnt
'                    !Trn_Type = "A"
'                    !PROD_CODE = mAdditionCode
'                    !Qty = 1
'                    !Rate = mAddition
'
'                    !U_Name = "Siebel"
'                    !U_EntDt = Format(PubLoginDate, "Short Date")
'                    !U_AE = "A"
'                    .Update
'                End With
'            End If
'
'            With RsTemp
'                .AddNew
'                !ChassisNo = StringPass(Master!Chassis_No)
'                !Pur_DocId = DocId
'                !pur_SrlNo = 1
'                !Pur_DocIDHelp = Replace(DocId, " ", "")
'                !Pur_SiteCode = PubSiteCode & PubSiteCode
'                !Pur_VType = mV_Type
'                !Pur_VNo = CodeCnt
'                !Pur_VDate = MakeDate(Master!Invoice_Date)
'                !Mfg_Month = mMfgMonth
'                !Mfg_Yr = mMfgYear
'                !InDate = MakeDate(Master!Invoice_Date)
'                !Model = StringPass(Master!VC_Number)
'                !Chas_Type = left(StringPass(Master!Chassis_No), 6)
'                !Godown = mGodownCode
'
'                mLength1 = InStr(1, StringPass(Master!Narration), "Engine") + Len("Engine Number - ")
'                mLength2 = InStr(1, StringPass(Master!Narration), "Chassis")
'                mLength2 = (mLength2 - mLength1)
'                !EngineNo = Replace(Trim(mID(StringPass(Master!Narration), mLength1, mLength2)), ".", "")
'                !Rate = mAmount
'                !Fixed = 0
'                !vrate = mTot_Amt
'                !Colour_Code = mColourCode
'                !Colours = mColourName
'                !Tax_YN = 1
'                !PBill_No = StringPass(Master!Invoice_No)
'                !Pbill_Date = MakeDate(Master!Invoice_Date)
'                !PartyCode = mPartyCode
'
'                !U_Name = "Siebel"
'                !U_EntDt = Format(PubLoginDate, "Short Date")
'                !U_AE = "A"
'
'                .Update
'            End With
'        End If
'        CodeCnt = CodeCnt + 1
'        CopyCnt = CopyCnt + 1
'        lblRecCopy(Index).CAPTION = CopyCnt
'        lblRecCopy(Index).Refresh
'MyNextRecord:
'        Master.MoveNext
'    Loop
'    GCn.CommitTrans
'    ImportBtn(Index).BackColor = FinishColor
'    If Not ConvertAll Then MsgBox ImportBtn(Index).CAPTION & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
'lblExit:
'    Set RsNew = Nothing
'    Exit Sub
'Eloop:
'    ErrorCnt = ErrorCnt + 1
'    lblRecError(Index).CAPTION = ErrorCnt: lblRecError(Index).Refresh
'    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Vehicle Purchase','" & mID(Replace(err.Description, "'", "`"), 1, 250) & "')")
'    Resume Next
'
'End Sub
