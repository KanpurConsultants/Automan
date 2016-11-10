Attribute VB_Name = "ModuleDec"
'pubSupDesig



Dim PubKillerActive As Boolean

Public PubImportData As Boolean
Public RsPartSiteWise As ADODB.Recordset
Public PubErrNum As Long
Public Const ErrNoVouNumManu As Long = 10
Public Const ErrNoVouNumAuto As Long = 11
Public pubLockDate$
Public pub_DllPath$
Public PubDealerID$
Public PubDiscOnLube As Byte
Public PubTOTOnLube As Byte
Public PubSrvTaxOnOutSideLab As Byte
Public PubOutSideLabDisc As Byte
Public PubEditLock As Integer
Public PubLockFinancialYear As Boolean
Public Const pubWrkDesigSuper As String = "SUPERVISOR"
Public Const pubWrkDesigRest As String = "'WASHER','MECHANIC','HELPER','ELECTRICIAN','DENTER'"
Public mPrinterName$
Public PubRsSyctrl As ADODB.Recordset
Public PubRsCompany As ADODB.Recordset
'21-04-2003
Public Const pubBrDivSysMainGrCode As String = "070"
Public Const pubSundryCrSysMainGrCode As String = "030003"
Public Const pubSundryDrSysMainGrCode As String = "060004"
'14-05-2003 lps
Public PubCompanyDbName As String

Public Const pubPurSysMainGrCode As String = "240"
Public Const pubSalSysMainGrCode As String = "230"
Public Const pubTaxSysMainGrCode As String = "030001"
'***********
Public Const PubNextSrvDays As Byte = 45
Public PubPageLength As Byte
Public PubPageLengthHalf As Byte
Public mChr14 As String   'enlarged characters 5 characters per inch
Public mChr10 As String   '10 charactergs per inch Normal Size
'Public mChr15 As String  '15 characters per inc
Public mChr17 As String   '17 characters per inc
Public mChr18 As String   '18 charactergs per inch
Public mChr20 As String   '20 characters per inch
Public mChr201 As String   'Off 20 character pitch
Public mEmph As String   'Emphasis current pitch
Public mEmph1 As String   'Off Emphasis
Public mDoub As String    'Double Strike
Public mDoub1 As String  'Off Double Strike
Public mUnd As String   'UnderLine
Public mUnd1 As String   'Off Underline
Public PubAmountPrefix$
Public PubLineFill$
Public PubAcPostingByAllUser As Boolean
Public PubChqNoReq As Boolean
'Public  mChrHt As String = Chr(27) + "w1" + Chr(27) + "W1" + Chr(27) + "E" + Chr(27) + "G" 'increase height
'Public  mChrHt1 As String = Chr(27) + "H" + Chr(27) + "F" + Chr(27) + "w0" + Chr(27) + "W0" 'Off height
Public mEject As String       '// PAGE EJECT
Public PubSpeedPrint As Boolean

' Old  Declaration  By Santosh For His Forms
Public PrintProper As Boolean
Public BiLanguage As Boolean
Public BiLanguageFont As String
Public BiLanguageName As String
Public CodeEditFlag As Boolean
Public Const mTopScale As Integer = 435
Public Const mBotScale As Integer = 450 '600
Public Const mRtScale As Integer = 150
Public Const mLtScale As Integer = 150
Public Const PubGridRowHeight As Byte = 220
'Grid Coloring Scheme
Public Const CellBackColEnter  As String = &HF0D5BF    '&HFFC0C0
Public Const CellBackColLeave As String = &HBAD3C9
'Public Const CellForeColEnter  As String = &HFFC0C0 '&HFFFF&
'Public Const CellForeColLeave As String = &H80000008
Public Const BackColorSelEnter As String = &HF8D7FD
'******
'Control like text etc coloring Scheme
Public Const CtrlBColDisabled = &HEBF0F1
Public Const CtrlBColOrg = &H80000005 '&HE3FAFD      '&H80000018   '&HCFE0E0      'Orginal BackColour
Public Const CtrlFColOrg = &H80000008   'Orginal ForeColour
'Public Const CtrlBCol = &HF0D5BF       '&H0&           'Changed BackColour
Public Const CtrlBCol = &HFDF4B5        '&HC0C0FF       'Changed BackColour
'Public Const CtrlBCol = &HFFC0FF       'Changed BackColour  (Pink)
Public Const CtrlFCol = &H80000012      '&HFFFF&        'Changed ForeColour
'Part Master
'Public Const CtrlBColOrg = &HC2D5B9        'Orginal BackColour
'Public Const CtrlFColOrg = &H80000012      'Orginal ForeColour
'Public Const CtrlBColOrg1 = &H8000000F     'Orginal BackColour
'Public Const CtrlFColOrg1 = &H80000012     'Orginal ForeColour
'Public Const CtrlBCol = &H80000008         'Changed BackColour
'Public Const CtrlFCol = &H8000000E         'Changed ForeColour
' ******* DECLARE GLOBAL VARIABLES(ACTUAL) *****
'Public Const PubIniDate As Date = "Null"
Public Const PubPackage As String = "Automan 1.0"

Public PubReSaleTaxPer As Single
Public pubTOT_On As Byte
Public pubTOTCaption$
Public PubCrLimitCheck As Byte
Public PubForm31Caption$
Public PubDivCode$
Public PubDivSName$
Public PubServiceTaxNo$
Public PubSiteCode As String * 1
Public PubSiteType As String
Public PubRSO_Code As String
Public PubOwnFinCode As String
Public PubRestrict_Godown As Byte
Public PubSprIssOnNegStk As Byte
Public PubVehGodown As String
Public PubSprCounterGodown As String
Public PubSprWorksGodown As String
Public PubRoundOffPosition As Byte
Public PubRoundOffType As String    'Standard, Upper Side, Lower Side
Public PubService_Tax As Single
Public PubSrvGatePass As Byte
Public PubSepLabourInv As Byte
Public PubServiceZone As String
Public PubLabRate_Chargable As Single
Public PubLabRate_Warranty As Single
Public PubIPO_Separate As Byte
Public PubGatePassOnSprInv As Byte
Public PubGenSurChrgOnSpr As Byte
Public pubGovtTaxFormSpr$
Public pubLocalTaxFormSpr$
Public PubTOT_YN As Byte
Public PubTOT_Rate As Byte
Public PubTaxDetOnSprInv As Byte
Public PubTBR_to_TPR As Byte
Public PubPartGrade_Lub As String
Public PubPartGrade_Consum As String
Public PubPartGrade_Tool As String
Public PubMergeGenSur_TB_Sale As Byte
Public PubCenDataPath As String
Public PubWSecFaDataPath As String
Public Pub_DataPath As String
Public PubServerName As String
Public PubServerNameCompany As String
Public PubDbUser$, PubDbPass$
Public PubVFADataPath As String
Public PubSFADataPath As String
Public PubWFADataPath As String
Public PubTaxOnFreeLabYn As Byte
Public PubComp_Contact As String
Public PubDbUserCompany As String
Public PubDbPassCompany As String
Public PubOffLineServer As String





'''Voucher Types Declaration For Siebel Integration'''''''
Public Const PubDmsVTypeSprPurCredit        As String = "D_SRP"
Public Const PubDmsVTypeSprPurCash          As String = "D_SCP"
Public Const PubDmsVTypeSprSaleCredit       As String = "D_SRS"
Public Const PubDmsVTypeSprSaleCash         As String = "D_SCS"
Public Const PubDmsVTypeWorkshopSaleCredit  As String = "D_WRS"
Public Const PubDmsVTypeWorkshopSaleCash    As String = "D_WCS"
Public Const PubDmsVTypeVehPur              As String = "D_VRP"
Public Const PubDmsVTypeVehSale             As String = "D_VRS"
Public Const PubDmsVTypeMoneyRectBank       As String = "D_BR"
Public Const PubDmsVTypeMoneyRectCash       As String = "D_CR"

Public Const PubSblImportVType As String = "('D_BR','D_CR','D_SCP','D_SCS','D_SRP','D_SRS','D_VRP','D_VRS''D_WCS','D_WRS')"





'MODISHEKHARKAPIL
Public PubFADataPath As String
Public PubConStrVFA As String
Public PubConStrSFA As String
Public PubConStrWFA As String
Public ConnectStr As String

Public GCn As ADODB.Connection
Public GCnFaV As ADODB.Connection
Public GCnFaS As ADODB.Connection
Public GCnFaW As ADODB.Connection
Public G_FaCn As ADODB.Connection
Public G_CompCn As ADODB.Connection
Public GCnTemp As ADODB.Connection
Public GRs As ADODB.Recordset
Public RsPart As ADODB.Recordset
Public RsPart1 As ADODB.Recordset
Public RsChart As ADODB.Recordset

Public PubRepLogonInfo$
Public PubRepoPath$
Public pubUAcPosting$
Public pubUName$
Public PubULabel$
Public PubUParam$
Public PubSysName$
Public PubSecName$
Public PubLoginDate As Date
Public PubStartDate As Date
Public PubEndDate As Date
Public PubBkpPath$ 'For Backup Purpose

Public Check As Variant
Public GSQL$
Public SearchForm As Form

Public G_VCTR As String
Public VrName As String
Public FldName1 As String

Public PubCenCompCode As String
Public PubFirmCode As String
Public PubVCompCode As String
Public PubSCompCode As String
Public PubWCompCode As String

Public PubComp_Name$
Public PubComp_Add$
Public PubComp_Add2$
Public PubComp_Add3$
Public PubComp_City$
Public PubComp_TINNo$
Public PubComp_CstNo$

Public PubSprCashAc As String
Public PubSrvCashAc As String
Public PubSrvLabAc As String
Public PubVATYN As Byte
Public PubSDTYN As Byte
Public PubSatYn As Byte
Public PubSiteWiseDisplayYn As Byte
Public PubVehRateIncTaxYn As Byte
Public PubSiebelActiveYn As Byte
Public PubCreditCardAc As String
Public PubChqClrAc As String
Public PubSprTaxInvPrefix As String
Public PubVehTaxInvPrefix As String




''''' Added by SKG
Public rdApp As CRAXDRT.Application
Public rpt As CRAXDRT.Report
Public rpt2 As CRAXDRT.Report
Public mREPORT
Public mRepName As String
Public RSOJPR As Boolean
Declare Function CreateFieldDefFile Lib "p2smon.dll" (X As Object, ByVal fieldDefFilePath$, ByVal bOverWriteExistingFiles%) As Integer

Public rsUserPerm As Recordset
Public WarrFrmName As Byte
Public FindFormatStr()              As Long



Public PubKillerFile As String
Public Const PubKillerFilePrefix As String = "sys"
Public PubDataSynchronisationApplicable As Boolean

Public Const Voucher_Category_Payment As String = "PYMT"
Public Const Voucher_NCat_BankPayment As String = "BPYMT"
Public Const Voucher_NCat_CashPayment As String = "CPYMT"
