VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmConvertTable 
   Caption         =   "Form1"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2085
   ScaleWidth      =   5640
   Begin VB.TextBox TxtFa 
      Height          =   315
      Left            =   150
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1005
      Width           =   4965
   End
   Begin VB.CommandButton CmdFa 
      Caption         =   "..."
      Height          =   315
      Left            =   5145
      TabIndex        =   7
      Top             =   1005
      Width           =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   585
      Top             =   2100
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   6
      Top             =   1395
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar Prg 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1785
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "..."
      Height          =   315
      Left            =   5145
      TabIndex        =   2
      Top             =   360
      Width           =   360
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   135
      Top             =   2070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Txt 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   4965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select FaData.MDB"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   9
      Top             =   720
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Table Name : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   5
      Top             =   1470
      Width           =   1095
   End
   Begin VB.Label LblTableName 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   1290
      TabIndex        =   4
      Top             =   1478
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Automan.MDB"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   105
      Width           =   1800
   End
End
Attribute VB_Name = "FrmConvertTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mFileName As String
Dim mFileNameFa As String
Dim GcnAccess As ADODB.Connection
Dim GcnFAAccess As ADODB.Connection
Dim mConverting As Boolean


Private Sub Cmd_Click()
On Error GoTo ErrHandler

    mFileName = ""
    'CommonDialog1.InitDir = Pub_DataPath
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Select MDB File For Table Conversion"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "Excel Files (*.Mdb)|*.Mdb"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen
    Txt = CommonDialog1.FileName
    mFileName = CommonDialog1.FileTitle
    
    If mFileName <> "" Then
        With GcnAccess
            Set GcnAccess = New ADODB.Connection
            GcnAccess.CursorLocation = adUseClient
            GcnAccess.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & Txt & ";Persist Security Info=False"
            GcnAccess.Open
        End With
    End If
  
Exit Sub
ErrHandler:
    MsgBox err.Description & " In Cmd_Click Procedure Of " & Me.Name
End Sub

Private Sub CmdFa_Click()
On Error GoTo ErrHandler

    mFileNameFa = ""
    'CommonDialog1.InitDir = Pub_DataPath
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Select MDB File For Table Conversion"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.Filter = "Excel Files (*.Mdb)|*.Mdb"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen
    TxtFa = CommonDialog1.FileName
    mFileNameFa = CommonDialog1.FileTitle
    
    If mFileNameFa <> "" Then
        With GcnFAAccess
            Set GcnFAAccess = New ADODB.Connection
            GcnFAAccess.CursorLocation = adUseClient
            GcnFAAccess.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & TxtFa & ";Persist Security Info=False"
            GcnFAAccess.Open
        End With
    End If
  
Exit Sub
ErrHandler:
    MsgBox err.Description & " In Cmd_Click Procedure Of " & Me.Name
End Sub



Sub ConvertTable()

Dim RsTemp As ADODB.Recordset

GCn.BeginTrans
    mConverting = True
    LblTableName = "HisCard"
    LblTableName.Refresh
    Set RsTemp = GcnAccess.Execute("Select * From HisCard")
    GCn.Execute "Delete From HisCard"
    With RsTemp
        If .RecordCount > 0 Then
            Prg.Visible = True
            Prg.Value = 0
            
            GCn.Execute "Delete From HisCard"
            Do Until .EOF
                GCn.Execute "Insert Into HisCard (CardNo, Div_Code, Site_Code, CardDate, Model, RegDate, RegNo, Chas_Type, Chassis, Engine, TransAxelNo, Delivery_Date, Dealer_Code, CouponNo, GBoxNo, RAxelNo, VehSerialNo, Supplier_BillNo, Supplier_BillDate, FAxelNo, SteerGBNo, CabinNo, BodyNo, ColourCode, DoorLockNo, SteerLockNo, FillerCapLockNo, HeadLampMake, FIP_No, Frame_No,Steer_Type,Steer_Make,Alternator,StarterMotor,Battery,Brake_Type,Radiator_Make,Tyre_FL,Tyre_FR,Tyre_ML1,Tyre_ML2,Tyre_MR1,Tyre_MR2,Tyre_RL1,Tyre_RL2,Tyre_RR1,Tyre_RR2,Spare_Wheel,Addl_Equp,FUEL,Fuel_Unit,Name,ConPerson,Add1,Add2,Add3,Area,CityCode,Pin,PhoneOff,PhoneResi,Mobile,Mail_ID,DOB,DOM,OwnDrive,OwnerRemark,Next_JobDate,Ac_Code,Govt_YN,Inv_No,Locked_Text,VehicleImage,OwnerImage,LJob_DocId,LJob_Date,LJob_AtKMsHrs,Trf_Date,U_Name,U_EntDt,U_AE,DisSprMRP,DisSprTB,DisSprTP,DisOilMRP,DisOilTB,DisOilTP,Varient,VehDet,ExtendWar) " & _
                            " Values('" & !CardNo & "', '" & !Div_Code & "', '" & !Site_Code & "',  " & ConvertDate(!CardDATE) & ",  '" & XNull(!Model) & "', " & ConvertDate(!RegDate) & ",  '" & XNull(!RegNo) & "',  '" & XNull(!Chas_Type) & "', '" & XNull(!Chassis) & "',  '" & XNull(!Engine) & "',  '" & XNull(!TransAxelNo) & "',  " & ConvertDate(!Delivery_Date) & ",  '" & XNull(!dealer_code) & "',  '" & XNull(!CouponNo) & "',  '" & XNull(!GBoxNo) & "',  '" & XNull(!RAxelNo) & "',  '" & XNull(!VehSerialNo) & "',  '" & XNull(!SUPPLIER_BILLNO) & "', " & ConvertDate(!Supplier_BillDate) & ",  '" & XNull(!FAxelNo) & "',  '" & XNull(!SteerGBNo) & "', '" & XNull(!cabinno) & "', '" & XNull(!BodyNo) & "',  '" & XNull(!Colourcode) & "',  '" & XNull(!DoorLockNo) & "',  '" & XNull(!SteerLockNo) & "', " & _
                            " '" & XNull(!FillerCapLockNo) & "', '" & XNull(!HeadLampMake) & "', '" & XNull(!FIP_No) & "', '" & XNull(!Frame_No) & "', '" & XNull(!Steer_Type) & "', '" & XNull(!Steer_Make) & "', '" & XNull(!Alternator) & "', '" & XNull(!StarterMotor) & "', '" & XNull(!Battery) & "', '" & XNull(!Brake_Type) & "', '" & XNull(!Radiator_Make) & "', '" & XNull(!Tyre_FL) & "', '" & XNull(!Tyre_FR) & "', '" & XNull(!Tyre_ML1) & "', '" & XNull(!Tyre_ML2) & "', '" & XNull(!Tyre_MR1) & "', '" & XNull(!Tyre_MR2) & "', '" & XNull(!Tyre_RL1) & "', '" & XNull(!Tyre_RL2) & "', '" & XNull(!Tyre_RR1) & "', '" & XNull(!Tyre_RR2) & "', '" & XNull(!Spare_Wheel) & "', " & _
                            " '" & XNull(!Addl_Equp) & "', '" & XNull(!FUEL) & "', '" & XNull(!Fuel_Unit) & "', '" & XNull(!Name) & "', '" & XNull(!ConPerson) & "', '" & XNull(!Add1) & "', '" & XNull(!Add2) & "', '" & XNull(!Add3) & "', '" & XNull(!Area) & "', '" & XNull(!CityCode) & "', '" & XNull(!Pin) & "', '" & XNull(!PhoneOff) & "', '" & XNull(!PhoneResi) & "', '" & XNull(!Mobile) & "', '" & XNull(!Mail_ID) & "', " & ConvertDate(!DOB) & ", " & ConvertDate(!DOM) & ", '" & XNull(!OwnDrive) & "', '" & XNull(!OwnerRemark) & "', " & ConvertDate(!Next_JobDate) & ", '" & XNull(!Ac_Code) & "', '" & XNull(!Govt_YN) & "', '" & XNull(!Inv_No) & "', '" & XNull(!Locked_Text) & "', '" & XNull(!VehicleImage) & "', '" & XNull(!OwnerImage) & "', '" & XNull(!ljob_docID) & "', " & _
                            " " & ConvertDate(!ljob_date) & ", '" & XNull(!ljob_atkmshrs) & "', Null, '" & XNull(!U_Name) & "', " & ConvertDate(!U_EntDt) & ", '" & XNull(!U_AE) & "', " & VNull(!DisSprMRP) & ", " & VNull(!DisSprTB) & ", " & VNull(!DisSprTP) & ", " & VNull(!DisOilMRP) & ", " & VNull(!DisOilTB) & ", " & VNull(!DisOilTP) & ", '" & XNull(!Varient) & "', '" & XNull(!VehDet) & "', '" & XNull(!ExtendWar) & "') "
                
                Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                .MoveNext
            Loop
        End If
    End With
    
    
    
    
    
    
    LblTableName = "Job_Card"
    LblTableName.Refresh
    Set RsTemp = GcnAccess.Execute("Select * From Job_Card")
    GCn.Execute "Delete From Job_Card"
    With RsTemp
        If .RecordCount > 0 Then
            Prg.Visible = True
            Prg.Value = 0
            

            Do Until .EOF
            
                GCn.Execute "Insert Into Job_Card (DocID, Site_Code, Job_No, Job_Date, Job_BookDivCode, Job_BookNo, Job_BookSiteCode, " & _
                            "Job_InspDivCode, Job_InspNo, Job_InspSiteCode, CardNo, Govt_YN, Serv_Type, Serv_Rate, Coupon, Coupon_Value, " & _
                            "AtKMsHrs, FUEL, Est_SpCost, Est_LabCost, ArrivalTime, Recp_Time, ExpDelDate, Body_Damage, OpenRemarks, " & _
                            "CreatedU_Name, CreatedU_EntDt, CreatedU_AE, JobCloseDate, JobComp_Dt_Time, DocId_InvSpr, DocId_InvLab, " & _
                            "GP_No, CRMemo, DrLab_AcCode, DrSpr_AcCode, BillingName, DelBy, NextSrvDate, FreeSpr_Amt, LabAmt_TB, " & _
                            "LabAmt_TP, Lab_TaxPer, Lab_TaxAmt, Lab_D_Amt, Lab_RoundOff, NetLab_Amt, Lab_Paid, Remark, DelayReason, " & _
                            "WRLabNo, WRLabDt, LabBillPrinted, ObservBy_Super, ActionBy_Super, ObservBy_Eng, ENG_OIL, GB_OIL, RA_OIL, " & _
                            "FA_OIL, TR_OIL, OIL_FLT, FUEL_FLT, TO_TSBS, FROM_TSBS, TSBS_B_NO, TSBS_B_DT, TSBS_B_SRL, ZSR_VFY_ID, " & _
                            "ZSR_VFY_DT, REJ_IND, DLAB_CODE, DLAB_SUB, ClosedU_Name, ClosedU_EntDt, ClosedU_AE, U_Name, U_EntDt, " & _
                            "U_AE, Trf_Date, Job_InspDocId, RecBy_Mechanic, RecBy_Supervisor, LabAmt_Out, KmsHrs, HrMeter, LabD_Per, " & _
                            "LastInvDocid, LastLabInvDocId, LastInvNoSuff, LastLabInvNoSuff, JobType, SiebelDocID, eCessPer, " & _
                            "eCessAmt, TempCloseDate, CreditCardNo, ChqNo, ChqDate, FreeWarrLabAmt) " & _
                            "Values ('" & XNull(!DocID) & "', '" & XNull(!Site_Code) & "', " & VNull(!Job_No) & ", " & ConvertDate(!Job_Date) & ", '" & XNull(!Job_BookDivCode) & "', '" & XNull(!Job_BookNo) & "', '" & XNull(!Job_BookSiteCode) & "', " & _
                            "'" & XNull(!Job_InspDivCode) & "', '" & XNull(!Job_InspNo) & "', '" & XNull(!Job_InspSiteCode) & "', '" & XNull(!CardNo) & "', '" & XNull(!Govt_YN) & "', '" & XNull(!Serv_Type) & "', " & VNull(!Serv_Rate) & ", '" & XNull(!Coupon) & "', " & VNull(!Coupon_Value) & ", " & _
                            "'" & XNull(!AtKMsHrs) & "', '" & XNull(!FUEL) & "', " & VNull(!Est_SpCost) & ", " & VNull(!Est_LabCost) & ", '" & XNull(!ArrivalTime) & "', '" & XNull(!Recp_Time) & "', " & ConvertDate(!ExpDelDate) & ", '" & XNull(!body_damage) & "', '" & XNull(!OpenRemarks) & "', " & _
                            "'" & XNull(!CreatedU_Name) & "', " & ConvertDate(!CreatedU_EntDt) & ", '" & XNull(!CreatedU_AE) & "', " & ConvertDate(!JobCloseDate) & ", " & ConvertDate(!JobComp_Dt_Time) & ", '" & XNull(!DocId_InvSpr) & "', '" & XNull(!DocID_InvLab) & "', " & _
                            "'" & XNull(!gp_no) & "', '" & XNull(!CrMemo) & "', '" & XNull(!DrLab_AcCode) & "', '" & XNull(!DrSpr_AcCode) & "', '" & XNull(!BillingName) & "', '" & XNull(!DelBy) & "', '" & XNull(!NextSrvDate) & "', " & VNull(!FreeSpr_Amt) & ", " & VNull(!LabAmt_TB) & ", " & _
                            "" & VNull(!LabAmt_TP) & ", " & VNull(!Lab_TaxPer) & ", " & VNull(!Lab_TaxAmt) & ", " & VNull(!Lab_D_Amt) & ", " & VNull(!Lab_RoundOff) & ", " & VNull(!NetLab_Amt) & ", '" & XNull(!Lab_Paid) & "', '" & XNull(!Remark) & "', '" & XNull(!DelayReason) & "', " & _
                            "'" & XNull(!WRLabNo) & "', " & ConvertDate(!WRLabDt) & ", '" & XNull(!LabBillPrinted) & "', '" & XNull(!ObservBy_Super) & "', '" & XNull(!ActionBy_Super) & "', '" & XNull(!ObservBy_Eng) & "', '" & XNull(!Eng_Oil) & "', '" & XNull(!GB_OIL) & "', '" & XNull(!RA_OIL) & "', " & _
                            "" & VNull(!FA_OIL) & ", " & VNull(!TR_OIL) & ", " & VNull(!OIL_FLT) & ", " & VNull(!FUEL_FLT) & ", " & VNull(!TO_TSBS) & ", " & VNull(!FROM_TSBS) & ", " & VNull(!TSBS_B_NO) & ", " & ConvertDate(!TSBS_B_DT) & ", " & XNull(!TSBS_B_SRL) & ", '" & XNull(!ZSR_VFY_ID) & "', " & _
                            "" & ConvertDate(!ZSR_VFY_DT) & ", '" & XNull(!REJ_IND) & "', '" & XNull(!DLAB_CODE) & "', '" & XNull(!DLAB_SUB) & "', '" & XNull(!ClosedU_Name) & "', " & ConvertDate(!ClosedU_EntDt) & ", '" & XNull(!ClosedU_AE) & "', '" & XNull(!U_Name) & "', " & ConvertDate(!U_EntDt) & ", " & _
                            "'" & XNull(!U_AE) & "', Null, '" & XNull(!Job_Inspdocid) & "', '" & XNull(!RecBy_Mechanic) & "', '" & XNull(!RecBy_Supervisor) & "', " & VNull(!LabAmt_Out) & ", '" & XNull(!KmsHrs) & "', '" & XNull(!HrMeter) & "', " & VNull(!LabD_Per) & ", " & _
                            "'" & XNull(!LastInvDocid) & "', '" & XNull(!LastLabInvDocId) & "', " & VNull(!LastInvNoSuff) & ", " & VNull(!LastLabInvNoSuff) & ", '" & XNull(!JobType) & "', '" & XNull(!SiebelDocID) & "', " & VNull(!eCessPer) & ", " & _
                            "" & VNull(!eCessAmt) & ", " & ConvertDate(!TempCloseDate) & ", '" & XNull(!CreditCardNo) & "', '" & XNull(!ChqNo) & "', " & ConvertDate(!ChqDate) & ", " & VNull(!FreeWarrLabAmt) & ")"
            
                
                Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                .MoveNext
            Loop
        End If
    End With
    
    
    
    
    
    
    LblTableName = "Part"
    LblTableName.Refresh
    Set RsTemp = GcnAccess.Execute("Select * From Part")
    GCn.Execute "Delete From Part"
    With RsTemp
        If .RecordCount > 0 Then
            Prg.Visible = True
            Prg.Value = 0
            

            Do Until .EOF
            
                GCn.Execute "Insert Into Part (Part_No, Div_Code, Site_Code, Part_Name, Local_Name, Part_NoHelp, Part_NameHelp, Unit, MARK_YN, Part_Grade, Part_OEM, Security_Grade, Active_YN, Value_Method, Supl_Loca, Lead_Time, Min_Lvl, Max_Lvl, ReOrd_Lvl, Disc_Factor, Bin_Loca, MRP, MRP_Effect_Dt, TB_SRate, TP_SRate, TB_Effect_Dt, New_Part, Cur_MRP_TBStk, Cur_MRP_TPStk, Cur_TB_STk, Cur_TP_Stk, Cur_MRP_TBStk_Val, Cur_MRP_TPStk_Val, Cur_TB_Stk_Val, Cur_TP_Stk_Val, Cum_Stk_Rct, Cum_Stk_Iss, High_Pur_Rate, High_MRP, High_TB_SRate, High_TP_SRate, Low_Pur_Rate, Low_MRP, Low_TB_SRate, Low_TP_SRate, Model_Grp_Code, Aggregate_Grp_Code, Veh_Type, Photo, U_Name, U_AE, U_EntDt, Trf_Date, PhyStk, PurDocId, PurDate, PurRate) " & _
                            "Values('" & XNull(!Part_No) & "', '" & XNull(!Div_Code) & "', '" & XNull(!Site_Code) & "', '" & XNull(!Part_Name) & "', '" & XNull(!Local_Name) & "', '" & XNull(!Part_NoHelp) & "', '" & XNull(!Part_NameHelp) & "', '" & XNull(!Unit) & "', '" & XNull(!MARK_YN) & "', '" & XNull(!Part_Grade) & "', '" & XNull(!Part_OEM) & "', '" & XNull(!Security_Grade) & "', '" & XNull(!Active_YN) & "', '" & XNull(!Value_Method) & "', '" & XNull(!Supl_Loca) & "', '" & XNull(!Lead_Time) & "', " & VNull(!Min_Lvl) & ", " & _
                            "" & VNull(!Max_Lvl) & ", " & VNull(!ReOrd_Lvl) & " , '" & XNull(!Disc_Factor) & "', '" & XNull(!Bin_Loca) & "', " & VNull(!MRP) & ", " & ConvertDate(!MRP_Effect_Dt) & ", " & VNull(!TB_SRate) & ", " & VNull(!TP_SRate) & ", " & ConvertDate(!TB_Effect_Dt) & ", '" & XNull(!New_Part) & "', " & VNull(!Cur_MRP_TbStk) & ", " & VNull(!Cur_MRP_TPStk) & ", " & VNull(!Cur_TB_STk) & ", " & VNull(!Cur_TP_Stk) & ", " & VNull(!Cur_MRP_TBStk_Val) & ", " & VNull(!Cur_MRP_TPStk_Val) & ", " & VNull(!Cur_TB_Stk_Val) & ", " & VNull(!Cur_TP_Stk_Val) & ", " & VNull(!Cum_Stk_Rct) & ", " & VNull(!Cum_Stk_Iss) & ", " & VNull(!high_pur_rate) & ", " & VNull(!High_MRP) & ", " & VNull(!High_TB_SRate) & ", " & VNull(!High_TP_SRate) & ", " & VNull(!low_pur_rate) & ", " & VNull(!Low_MRP) & ", " & VNull(!Low_TB_SRate) & ", " & _
                            "" & VNull(!Low_TP_SRate) & ", '" & XNull(!Model_Grp_Code) & "', '" & XNull(!Aggregate_Grp_Code) & "', '" & XNull(!Veh_Type) & "', '" & XNull(!Photo) & "', '" & XNull(!U_Name) & "', '" & XNull(!U_AE) & "', '" & XNull(!U_EntDt) & "', Null, " & VNull(!PhyStk) & ", '" & XNull(!PurDocId) & "', '" & XNull(!PurDate) & "', " & VNull(!PurRate) & ") "
                
                Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                .MoveNext
            Loop
        End If
    End With
    
    
    
    
    
    
    LblTableName = "Site"
    LblTableName.Refresh
    Set RsTemp = GcnAccess.Execute("Select * From " & LblTableName & "")
    GCn.Execute "Delete From " & LblTableName & ""
    With RsTemp
        If .RecordCount > 0 Then
            Prg.Visible = True
            Prg.Value = 0
            

            Do Until .EOF
                GCn.Execute "Insert Into Site (Site_Code, Site_Desc, U_Name, U_EntDt, U_AE, Trf_Date, SiteType, Address1, Address2, Address3, City, PinCode, Phone, Mobile, LstNo, LstDate, CstNo, CstDate) " & _
                            "Values('" & XNull(!Site_Code) & "', '" & XNull(!Site_Desc) & "', '" & XNull(!U_Name) & "', '" & XNull(!U_EntDt) & "', '" & XNull(!U_AE) & "', '" & XNull(!Trf_Date) & "', '" & XNull(!SiteType) & "', '" & XNull(!Address1) & "', '" & XNull(!Address2) & "', '" & XNull(!Address3) & "', '" & XNull(!City) & "', '" & XNull(!PinCode) & "', '" & XNull(!Phone) & "', '" & XNull(!Mobile) & "', '" & XNull(!LstNo) & "', '" & XNull(!LstDate) & "', '" & XNull(!CstNo) & "', '" & XNull(!CstDate) & "') "
                
                Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                .MoveNext
            Loop
        End If
    End With
    
    
    
    
    
    
    LblTableName = "Sp_Purch"
    LblTableName.Refresh
    Set RsTemp = GcnAccess.Execute("Select * From " & LblTableName & "")
    GCn.Execute "Delete From " & LblTableName & ""
    With RsTemp
        If .RecordCount > 0 Then
            Prg.Visible = True
            Prg.Value = 0
            

            Do Until .EOF
            
                GCn.Execute "Insert Into Sp_Purch(DocID, DocIDHelp, V_Type, V_NO, Site_Code, V_Date, Cash_Credit, Party_code, Party_Name, L_C, Form_Code, FormNo, FormIssRecDate, Party_Doc_No, Party_Doc_Date, RoadPermit_FormCode, RoadPermit_No, GR_RR_No, GR_RR_Date, Tot_No_of_Items, Tot_Doc_Qty, Tot_Phy_Qty, SprAmt_MRP_TB, SprAmt_MRP_TP, OilAmt_MRP_TB, OilAmt_MRP_TP, SprAmt_TB, SprAmt_TP, OilAmt_TB, OilAmt_TP, OilAmt, SprAmt, Tot_Amt, Tot_Disc_Amt, Tot_Ord_DiscAmt, Tot_Goods_Value, Tax_Amt, Addition, Deduction, NET_AMT, EntryTaxPer, EntryTaxAmt, Case_No, Case_Mark, Transport, Supply_Mode, Remarks, Invoice_DocId, AcPsoting_YN, DrAc_Code, Printed_YN, CancelYN, CancelRemark, U_Name, U_EntDt, U_AE, Trf_Date, Transportation, SiebelDocID) " & _
                            "Values('" & XNull(!DocID) & "', '" & XNull(!DocIDHelp) & "', '" & XNull(!V_Type) & "', '" & XNull(!V_NO) & "', '" & XNull(!Site_Code) & "', " & ConvertDate(!V_DATE) & ", '" & XNull(!Cash_Credit) & "', '" & XNull(!Party_code) & "', '" & XNull(!Party_Name) & "', '" & XNull(!L_C) & "', '" & XNull(!Form_Code) & "', '" & XNull(!FormNo) & "', " & ConvertDate(!FormIssRecDate) & ", '" & XNull(!Party_Doc_No) & "', " & ConvertDate(!Party_Doc_Date) & ", '" & XNull(!RoadPermit_FormCode) & "', " & _
                            "'" & XNull(!RoadPermit_No) & "', '" & XNull(!GR_RR_No) & "', " & ConvertDate(!GR_RR_Date) & ", '" & XNull(!Tot_No_of_Items) & "', " & VNull(!Tot_Doc_Qty) & ", " & VNull(!Tot_Phy_Qty) & ", " & VNull(!SprAmt_MRP_TB) & ", " & VNull(!SprAmt_MRP_TP) & ", " & VNull(!OilAmt_MRP_TB) & ", " & VNull(!OilAmt_MRP_TP) & ", " & VNull(!SprAmt_TB) & ", " & VNull(!SprAmt_TP) & ", " & VNull(!OilAmt_TB) & ", " & VNull(!OilAmt_TP) & ", " & VNull(!OilAmt) & ", " & VNull(!SprAmt) & ", " & VNull(!Tot_Amt) & ", " & VNull(!Tot_Disc_Amt) & ", " & VNull(!Tot_Ord_DiscAmt) & ", " & VNull(!Tot_Goods_Value) & ", " & VNull(!Tax_Amt) & ", " & VNull(!Addition) & ", " & VNull(!Deduction) & ", " & VNull(!Net_Amt) & ", " & VNull(!EntryTaxPer) & ", " & VNull(!EntryTaxAmt) & ", " & _
                            "'" & XNull(!Case_No) & "', '" & XNull(!Case_Mark) & "', '" & XNull(!Transport) & "', '" & XNull(!Supply_Mode) & "', '" & XNull(!Remarks) & "', '" & XNull(!Invoice_DocID) & "', '" & XNull(!AcPsoting_YN) & "', '" & XNull(!DrAc_Code) & "', '" & XNull(!Printed_YN) & "', '" & XNull(!CancelYN) & "', '" & XNull(!CancelRemark) & "', '" & XNull(!U_Name) & "', " & ConvertDate(!U_EntDt) & ", '" & XNull(!U_AE) & "', " & ConvertDate(!Trf_Date) & ", '" & XNull(!Transportation) & "', '" & XNull(!SiebelDocID) & "') "
                
                Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                .MoveNext
            Loop
        End If
    End With
    
    
    
    
    
    
    LblTableName = "Sp_Stock"
    LblTableName.Refresh
    Set RsTemp = GcnAccess.Execute("Select * From " & LblTableName & "")
    GCn.Execute "Delete From " & LblTableName & ""
    With RsTemp
        If .RecordCount > 0 Then
            Prg.Visible = True
            Prg.Value = 0
            

            Do Until .EOF
                GCn.Execute "Insert Into Sp_Stock(DocID, Srl_No, V_Type, V_NO, V_Date, Party_code, L_C, Job_DocID, Job_DivCode, Mech_Code, Order_DocId, Order_Srl_No, Part_No, Lub_Category, Godown, Qty_Doc, Qty_Rec, Qty_Iss, Qty_Ret, Tax_YN, MRP_YN, Rate, MRP_Rate, Disc_Per, Disc_Amt, Amount, Ord_DiscPer, Ord_DiscAmt, Net_Amt, Purpose, V_Rate, Part_SrlNo, Remark, Printed, Invoice_DocId, V_Date2, Rate2, MRP_Rate2, Disc_Per2, Disc_Amt2, Amount2, Ord_DiscPer2, Ord_DiscAmt2, Net_Amt2, Printed2, TrnComplete_YN, Site_Code, U_Name, U_EntDt, U_AE, Trf_Date, Claim_Div, Claim_Site, Claim_YearPrefix, Claim_Type, Claim_No, Claim_Date, ClaimId, TaxPer, TaxAmt, PurDocNo, PurDocDate, SiebelDocID) " & _
                            "Values('" & XNull(!DocID) & "', '" & XNull(!Srl_No) & "', '" & XNull(!V_Type) & "', '" & XNull(!V_NO) & "', " & ConvertDate(!V_DATE) & ", '" & XNull(!Party_code) & "', '" & XNull(!L_C) & "', '" & XNull(!job_docid) & "', '" & XNull(!Job_DivCode) & "', '" & XNull(!mech_code) & "', '" & XNull(!Order_DocId) & "', '" & XNull(!Order_Srl_No) & "', '" & XNull(!Part_No) & "', '" & XNull(!Lub_Category) & "', '" & XNull(!Godown) & "', " & VNull(!Qty_Doc) & ", " & VNull(!Qty_Rec) & ", " & VNull(!Qty_Iss) & ", " & VNull(!Qty_Ret) & ", " & _
                            "'" & XNull(!Tax_YN) & "', '" & XNull(!MRP_YN) & "', " & VNull(!Rate) & ", " & VNull(!MRP_Rate) & ", " & VNull(!Disc_Per) & ", " & VNull(!Disc_Amt) & ", " & VNull(!Amount) & ", " & VNull(!ord_Discper) & ", " & VNull(!ord_Discamt) & ", " & VNull(!Net_Amt) & ", '" & XNull(!Purpose) & "', " & VNull(!V_Rate) & ", '" & XNull(!Part_SrlNo) & "', '" & XNull(!Remark) & "', '" & XNull(!Printed) & "', '" & XNull(!Invoice_DocID) & "', " & ConvertDate(!V_DATE2) & ", " & VNull(!Rate2) & ", " & VNull(!MRP_Rate2) & ", " & VNull(!Disc_Per2) & ", " & VNull(!Disc_Amt2) & ", " & VNull(!Amount2) & ", " & VNull(!ord_Discper2) & ", " & VNull(!ord_Discamt2) & ", " & VNull(!Net_Amt2) & ", '" & XNull(!Printed2) & "', '" & XNull(!TrnCompLete_YN) & "', '" & XNull(!Site_Code) & "', '" & XNull(!U_Name) & "', " & _
                            "'" & XNull(!U_EntDt) & "' , '" & XNull(!U_AE) & "', " & ConvertDate(!Trf_Date) & ", '" & XNull(!Claim_Div) & "', '" & XNull(!Claim_Site) & "', '" & XNull(!claim_YearPrefix) & "', '" & XNull(!claim_type) & "', '" & XNull(!claim_no) & "', " & ConvertDate(!Claim_Date) & ", '" & XNull(!ClaimId) & "', " & VNull(!TaxPer) & ", " & VNull(!TaxAmt) & ", '" & XNull(!PurDocNo) & "', " & ConvertDate(!PurDocDate) & ", '" & XNull(!SiebelDocID) & "') "
                
                Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                .MoveNext
            Loop
        End If
    End With
    
    
    
    
    
    
    LblTableName = "SubGroupType"
    LblTableName.Refresh
    Set RsTemp = GcnAccess.Execute("Select * From " & LblTableName & "")
    GCn.Execute "Delete From " & LblTableName & ""
    With RsTemp
        If .RecordCount > 0 Then
            Prg.Visible = True
            Prg.Value = 0
            

            Do Until .EOF
                GCn.Execute "Insert Into SubGroupType(Party_Type, Description, MRP_Disc, TB_Disc, TP_Disc) " & _
                            "Values('" & XNull(!Party_Type) & "', '" & XNull(!Description) & "', '" & XNull(!mrp_Disc) & "', '" & XNull(!tb_Disc) & "', '" & XNull(!tp_Disc) & "') "
                
                Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                .MoveNext
            Loop
        End If
    End With



    
    
    
    LblTableName = "VisitObjective"
    LblTableName.Refresh
    Set RsTemp = GcnAccess.Execute("Select * From " & LblTableName & "")
    GCn.Execute "Delete From " & LblTableName & ""
    With RsTemp
        If .RecordCount > 0 Then
            Prg.Visible = True
            Prg.Value = 0
            
            Do Until .EOF
                GCn.Execute "Insert Into VisitObjective(ObjCode, ObjDesc) Values ('" & XNull(!ObjCode) & "', '" & XNull(!ObjDesc) & "')"
                            
                Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                .MoveNext
            Loop
        End If
    End With







    LblTableName = "FinBank"
    LblTableName.Refresh
    Set RsTemp = GcnAccess.Execute("Select * From " & LblTableName & "")
    GCn.Execute "Delete From " & LblTableName & ""
    With RsTemp
        If .RecordCount > 0 Then
            Prg.Visible = True
            Prg.Value = 0
            
            Do Until .EOF
                GCn.Execute "Insert Into FinBank(FinBankCode, Site_Code, FinBankName, FinGrpCode, U_Name, U_EntDt, U_AE, Trf_Date, Inv_Prefix, OldCode, xName) " & _
                            "Values('" & XNull(!FinbankCode) & "', '" & XNull(!Site_Code) & "', '" & left(XNull(!finbankname), 40) & "', '" & XNull(!Fingrpcode) & "', '" & XNull(!U_Name) & "', '" & XNull(!U_EntDt) & "', '" & XNull(!U_AE) & "', " & ConvertDate(!Trf_Date) & ", '" & XNull(!Inv_Prefix) & "', '" & XNull(!OldCode) & "', '" & XNull(!xName) & "') "
                Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                .MoveNext
            Loop
        End If
    End With






    LblTableName = "AcGroup"
    LblTableName.Refresh
    Set RsTemp = GcnFAAccess.Execute("Select * From " & LblTableName & "")
    GCn.Execute "Delete From " & LblTableName & ""
    With RsTemp
        If .RecordCount > 0 Then
            Prg.Visible = True
            Prg.Value = 0
            
            Do Until .EOF
                GCn.Execute "Insert Into AcGroup(ID, Site_Code, GroupCode, GroupName, GroupNameBiLang, GroupNature, Nature, MainGrCode,  CurrentBalance, SubLedYN, BlOrd, AliasYN, GroupHelp, SysGroup, LastYearBalance, U_Name, U_EntDt, U_AE, TradingYN) " & _
                            "Values('" & XNull(!ID) & "', '" & XNull(!Site_Code) & "', '" & XNull(!GroupCode) & "', '" & XNull(!GroupName) & "', '" & XNull(!GroupNameBiLang) & "', '" & XNull(!GroupNature) & "', '" & XNull(!Nature) & "', '" & XNull(!MainGrCode) & "',   '" & Round(VNull(!CURRENTBALANCE), 2) & "', '" & XNull(!SubLedYN) & "', '" & XNull(!BLORD) & "', '" & XNull(!AliasYN) & "', '" & XNull(!GroupHelp) & "', '" & XNull(!SysGroup) & "', '" & XNull(!LastYearBalance) & "', '" & XNull(!U_Name) & "', '" & XNull(!U_EntDt) & "', '" & XNull(!U_AE) & "', '" & XNull(!TradingYN) & "') "
                
                Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                .MoveNext
            Loop
        End If
    End With
    mConverting = False
GCn.CommitTrans
End Sub


Private Sub CmdStart_Click()
    ConvertTable
End Sub

Private Sub Timer1_Timer()
Dim I As Long
    If mConverting = True Then
        For I = 1 To 99999
            LblTableName.Refresh
        Next I
    End If
End Sub
