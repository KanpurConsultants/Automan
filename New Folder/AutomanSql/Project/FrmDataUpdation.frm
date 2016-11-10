VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmDataUpdation 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Update Data"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
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
   ScaleHeight     =   5595
   ScaleWidth      =   11730
   Begin VB.Frame Frame4 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Excel File Format"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   45
      TabIndex        =   40
      Top             =   570
      Width           =   7935
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Part_Name (40)"
         Height          =   195
         Index           =   26
         Left            =   330
         TabIndex        =   53
         Top             =   1065
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Bin_Loca (15)"
         Height          =   195
         Index           =   25
         Left            =   330
         TabIndex        =   52
         Top             =   1740
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Disc_Factor (2)"
         Height          =   195
         Index           =   24
         Left            =   330
         TabIndex        =   51
         Top             =   1395
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Part_No (22)"
         Height          =   195
         Index           =   23
         Left            =   330
         TabIndex        =   50
         Top             =   735
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8. MRPTBStk"
         Height          =   195
         Index           =   22
         Left            =   3210
         TabIndex        =   49
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7. TPSRate"
         Height          =   195
         Index           =   21
         Left            =   3210
         TabIndex        =   48
         Top             =   1395
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6. TBSRate"
         Height          =   195
         Index           =   20
         Left            =   3210
         TabIndex        =   47
         Top             =   1065
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5. MRPRate"
         Height          =   195
         Index           =   19
         Left            =   3210
         TabIndex        =   46
         Top             =   735
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12. Part_Grade"
         Height          =   195
         Index           =   18
         Left            =   5715
         TabIndex        =   45
         Top             =   1740
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "11. TPStk"
         Height          =   195
         Index           =   17
         Left            =   5715
         TabIndex        =   44
         Top             =   1395
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10. TBStk"
         Height          =   195
         Index           =   15
         Left            =   5715
         TabIndex        =   43
         Top             =   1065
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9.  MRPTPStk"
         Height          =   195
         Index           =   14
         Left            =   5715
         TabIndex        =   42
         Top             =   735
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sheet Name : Part_Import"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2565
         TabIndex        =   41
         Top             =   255
         Width           =   2550
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Updation Summary"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4395
      Left            =   8100
      TabIndex        =   29
      Top             =   525
      Width           =   2940
      Begin VB.CommandButton CmdPrn 
         Caption         =   "Print"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2220
         Width           =   1065
      End
      Begin VB.Label LblName 
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmDataUpdation.frx":0000
         ForeColor       =   &H00C000C0&
         Height          =   1050
         Index           =   13
         Left            =   120
         TabIndex        =   39
         Top             =   3210
         Width           =   2745
      End
      Begin VB.Label LBLCNT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Not Copied"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   6
         Left            =   105
         TabIndex        =   37
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label LBLCNT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   285
         Index           =   7
         Left            =   1605
         TabIndex        =   36
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label LBLCNT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Record Copied"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   1
         Left            =   105
         TabIndex        =   35
         Top             =   1245
         Width           =   1470
      End
      Begin VB.Label LBLCNT 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   3
         Left            =   1605
         TabIndex        =   34
         Top             =   1245
         Width           =   120
      End
      Begin VB.Label LBLCNT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Records"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   33
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label LBLCNT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   1605
         TabIndex        =   32
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label LBLCNT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Table Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   4
         Left            =   105
         TabIndex        =   31
         Top             =   465
         Width           =   1575
      End
      Begin VB.Label LBLCNT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   285
         Index           =   5
         Left            =   1605
         TabIndex        =   30
         Top             =   465
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CFE2D9&
      Caption         =   "Excel File Format"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   60
      TabIndex        =   15
      Top             =   525
      Width           =   7935
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sheet Name : Part_Import"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   2565
         TabIndex        =   28
         Top             =   255
         Width           =   2550
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9. MRP_TPStk"
         Height          =   195
         Index           =   8
         Left            =   5715
         TabIndex        =   27
         Top             =   735
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10. TB_Stk"
         Height          =   195
         Index           =   9
         Left            =   5715
         TabIndex        =   26
         Top             =   1065
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "11. TP_Stk"
         Height          =   195
         Index           =   10
         Left            =   5715
         TabIndex        =   25
         Top             =   1395
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12. Part_Grade"
         Height          =   195
         Index           =   13
         Left            =   5715
         TabIndex        =   24
         Top             =   1740
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5. MRP_Rate"
         Height          =   195
         Index           =   4
         Left            =   3210
         TabIndex        =   23
         Top             =   735
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6. TB_SRate"
         Height          =   195
         Index           =   5
         Left            =   3210
         TabIndex        =   22
         Top             =   1065
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7. TP_SRate"
         Height          =   195
         Index           =   6
         Left            =   3210
         TabIndex        =   21
         Top             =   1395
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8. MRP_TBStk"
         Height          =   195
         Index           =   7
         Left            =   3210
         TabIndex        =   20
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Part_No (22)"
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   19
         Top             =   735
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Disc_Factor (2)"
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   18
         Top             =   1395
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Bin_Loca (15)"
         Height          =   195
         Index           =   3
         Left            =   330
         TabIndex        =   17
         Top             =   1740
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Part_Name (40)"
         Height          =   195
         Index           =   12
         Left            =   330
         TabIndex        =   16
         Top             =   1065
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Opening Stock Updation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   45
      TabIndex        =   6
      Top             =   2670
      Width           =   7950
      Begin VB.CommandButton CmdPart 
         Caption         =   "Start"
         Height          =   375
         Left            =   5565
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1065
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6630
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1065
         Width           =   1020
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   645
         Width           =   7080
      End
      Begin VB.CommandButton CmdSel 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   7335
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exit Form"
         Top             =   645
         Width           =   315
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1770
         Visible         =   0   'False
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Ms-Excel File"
         Height          =   210
         Left            =   255
         TabIndex        =   14
         Top             =   420
         Width           =   1845
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Updation Status"
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   2
         Left            =   255
         TabIndex        =   13
         Top             =   1530
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label LblName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%%%%%"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   6720
         TabIndex        =   12
         Top             =   1530
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   21
      Left            =   15
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   45
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton CmdBlank 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Make Blank Print File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   285
      Visible         =   0   'False
      Width           =   3090
   End
   Begin VB.CommandButton CmdPath 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Apply Path"
      Height          =   360
      Left            =   6645
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2250
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   4365
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   30
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.xls"
   End
   Begin VB.Label Label3 
      Caption         =   "Label3(16)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   16
      Left            =   2040
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label LBLCNT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Write Import Data Path"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   8
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   2040
   End
End
Attribute VB_Name = "FrmDataUpdation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldGcn As ADODB.Connection
Dim OpenCon As Boolean
Private Const FilePath = 2

Private Sub CmdBlank_Click()
GCn.Execute ("Delete * from PrnMissRec")
MsgBox "Job Completed", vbInformation
End Sub

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdPart_Click()
'Part_No,Site_Code,Part_Name,Local_Name,Part_NoHelp,Part_NameHelp,UNIT,MARK_YN,Part_Grade
'Part_OEM,Security_Grade,Active_YN,Value_Method,Supl_Loca,Lead_Time,Min_Lvl,Max_Lvl,
'ReOrd_Lvl,Disc_Factor,Bin_Loca,MRP,MRP_Effect_Dt,TB_SRate,TP_SRate,TB_Effect_Dt,New_Part
'Cur_MRP_TBStk,Cur_MRP_TPStk,Cur_TB_Stk,Cur_TP_Stk,Cur_MRP_TBStk_Val,Cur_MRP_TPStk_Val
'Cur_TB_Stk_Val,Cur_TP_Stk_Val,Cum_Stk_Rct,Cum_Stk_Iss,High_Pur_Rate,High_MRP,High_TB_SRate
'High_TP_SRate,Low_Pur_Rate,Low_MRP,Low_TB_SRate,Low_TP_SRate,Model_Grp_Code,
'Aggregate_Grp_Code,Veh_Type,Photo,U_Name,U_EntDt,U_AE,Trf_Date

'Part_No,Site_Code***,PART_NAME,LOCAL_NAME,Part_NoHelp***,Part_NameHelp***,UNIT,
'MARK,PART_GRADE,Part_OEM**,Security_Grade***,CURR_PART,Value_Method**,SUPL_LOCA
'LEAD_TIME,MIN_STK,MAX_STK,REORD_STK,DISC,BIN_LOCA,MRP_RATE,EFFECT_DT,TB_RATE,TP_RATE
'New_Part**,Cur_MRP_TBStk**,Cur_MRP_TPStk**,CUR_TP_STK,CUR_TB_STK,CUR_TP_OLD,CUR_TB_OLD
'CUM_TP_RCT_marge,CUM_TB_RCT_marge,CUM_TP_ISS_merge,CUM_TB_ISS_merge,HIGH_TP,HIGH_TB,
'USER_NAME,ENTRY_DATE,ENTRY_TIME

PartOpen
Exit Sub
'***** Updatating PartMaster  **********
Dim Rs As ADODB.Recordset
Dim SecurityGrade$, PartNo$, ProGrade$, PartName$, LocalName$, PartNameHlp$, PartNoHlp$, BinLoca$
Dim Cnt As Long
Dim Cnt1 As Long
Dim VoucherEditFlag As Boolean, mVNo$, mVType$

On Error GoTo ELoop

GSQL = "select * from part_import" 'for existing dos system
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
'rs.Open GSQL, OldGcn, adOpenDynamic, adLockOptimistic
Rs.Open GSQL, GCn, adOpenDynamic, adLockOptimistic 'for un

If Rs.RecordCount = 0 Then MsgBox "There is no records in import Database table !! Part !!", vbInformation, "No Records Found": Exit Sub
If GCn.Execute("select count(Part_No) from part").Fields(0).Value > 0 Then
    If MsgBox("There are records in Part Master !! Do you want to continue ? ", vbYesNo + vbCritical + vbDefaultButton2, "Update Database !") = vbNo Then Exit Sub
End If
LBLCNT(2).CAPTION = Rs.RecordCount
LBLCNT(2).Refresh
LBLCNT(5).CAPTION = "Part"
LBLCNT(5).Refresh
GCn.BeginTrans
'GCn.Execute ("Delete * from part")
Do Until Rs.EOF
    PartNo = PubDivCode & Replace(IIf(IsNull(Rs!Part_No), "", Rs!Part_No), "'", "`")
    PartName = Replace(IIf(IsNull(Rs!Part_Name), "", Rs!Part_Name), "'", "`")
    LocalName = Replace(IIf(IsNull(Rs!Local_Name), "", Rs!Local_Name), "'", "`")
    BinLoca = IIf(IsNull(Rs!Bin_Loca), "", Rs!Bin_Loca)
    PartNameHlp = UCase(Replace(PartName, " ", ""))
    PartNoHlp = UCase(Replace(PartNo, " ", ""))
   
'    GCn.Execute "insert into part( " & _
        "Part_No,Site_Code,Part_Name,Local_Name,Part_NoHelp, " & _
        "Part_NameHelp,UNIT,MARK_YN,Part_Grade, " & _
        "Security_Grade,Active_YN,Value_Method,Supl_Loca,Lead_Time, " & _
        "Min_Lvl,Max_Lvl,ReOrd_Lvl,Disc_Factor,Bin_Loca," & _
        "MRP,MRP_Effect_Dt,TB_SRate,TP_SRate," & _
        "New_Part,Cur_MRP_TBStk,Cur_MRP_TPStk,Cur_TB_Stk,Cur_TP_Stk," & _
        "Cum_Stk_Rct,Cum_Stk_Iss,U_Name,U_EntDt,U_AE) " & _
        "values('" & PartNo & "','" & PubSiteCode & "','" & PartName & "','" & LocalName & "','" & UCase(Replace(PartNo, " ", "")) & "'," & _
        "'" & UCase(Replace(PartName, " ", "")) & "','" & IIf(IsNull(rs!Unit), "", rs!Unit) & "','" & IIf(IsNull(rs!MARK), "", rs!MARK) & "','" & IIf(IsNull(rs!Part_Grade), "", rs!Part_Grade) & "', " & _
        "'A',0,'FIFO','F001','" & IIf(IsNull(rs!Lead_Time), "", rs!Lead_Time) & "'," & _
        "" & rs!MIN_STK & "," & rs!MAX_STK & "," & rs!REORD_STK & "," & rs!DISC & ",'" & BinLoca & "', " & _
        "" & rs!MRP_RATE & "," & ConvertDate(rs!Effect_Dt) & "," & rs!TB_RATE & "," & rs!TP_RATE & "," & _
        " 1," & (rs!Cur_TB_Stk - rs!CUR_TB_OLD) & "," & (rs!Cur_TP_Stk - rs!CUR_TP_OLD) & "," & rs!Cur_TB_Stk & "," & rs!Cur_TP_Stk & ", " & _
        "" & (rs!CUM_TP_RCT + rs!CUM_TB_RCT) & "," & (rs!CUM_TP_ISS + rs!CUM_TB_ISS) & ",'SA',#" & PubLoginDate & "#,'A')"
    

    
    'GCn.Execute "insert into part( " & _
        "Part_No,Div_Code,Site_Code,Part_Name,Local_Name,Part_NoHelp, " & _
        "Part_NameHelp,UNIT,Part_Grade, " & _
        "Security_Grade,Active_YN,Value_Method,Lead_Time,Min_Lvl,Max_Lvl,ReOrd_Lvl,Disc_Factor," & _
        "Bin_Loca,New_Part,MRP_Effect_Dt,MRP,TB_SRate,TP_SRate," & _
        "Cur_MRP_TBStk,Cur_MRP_TPStk,Cur_TB_Stk,Cur_TP_Stk," & _
        "U_Name,U_EntDt,U_AE,TB_Effect_Dt) " & _
        "values('" & PartNo & "','" & PubDivCode & "','" & PubSiteCode & "','" & PartName & "','" & LocalName & "','" & PartNoHlp & _
        "','" & PartNameHlp & "','" & IIf(IsNull(rs!Unit), "", rs!Unit) & "','" & IIf(IsNull(rs!Part_Grade), "S", rs!Part_Grade) & _
        "','A',0,'FIFO'," & VNull(rs!Lead_Time) & "," & IIf(IsNull(rs!MIN_STK), 0, rs!MIN_STK) & "," & IIf(IsNull(rs!Max_STK), 0, rs!Max_STK) & "," & IIf(IsNull(rs!REORD_STK), 0, rs!REORD_STK) & _
        ",'" & IIf(IsNull(rs!DISC), "X", rs!DISC) & "','" & BinLoca & "',1," & ConvertDate(rs!Effect_Dt) & "," & VNull(rs!MRP_RATE) & "," & VNull(rs!TB_Rate) & "," & VNull(rs!TP_RATE) & _
        "," & VNull(rs!MRP_TBSTK) & "," & VNull(rs!MRP_TPStk) & "," & VNull(rs!TB_Stk) & "," & VNull(rs!TP_Stk) & _
        ",'SA',#" & PubServerDate & "#,'A'," & ConvertDate(rs!Effect_Dt) & ")"
    'insert opening rec in sp_stock
    If VNull(Rs!MRP_TBSTK) + VNull(Rs!MRP_TPStk) + VNull(Rs!TB_Stk) + VNull(Rs!TP_Stk) <> 0 Then
        mSrlNo = 0: lblPrefix = "": mVType = "SXAO"
        mVNo = GetDocID(GCnFaS, mVType, PubStartDate, VoucherEditFlag, txt(21), Label3(16), PubSiteCode)
        If Rs!MRP_TBSTK <> 0 Then
            GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
                "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,MRP_Rate," & _
                "Site_Code,U_Name,U_EntDt,U_AE) " & _
                "VALUES ('" & mVNo & "',1,'" & mVType & "'," & txt(21) & "," & ConvertDate(PubStartDate) & ",'" & PartNo & _
                "','" & PubSprCounterGodown & "',1,1," & VNull(Rs!MRP_TBSTK) & "," & VNull(Rs!MRP_Rate) & "," & Round(VNull(Rs!MRP_TBSTK) * VNull(Rs!MRP_Rate), 2) & "," & VNull(Rs!MRP_Rate) & _
                ",'" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
        End If
        If Rs!MRP_TPStk <> 0 Then
            GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
                "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,MRP_Rate," & _
                "Site_Code,U_Name,U_EntDt,U_AE) " & _
                "VALUES ('" & mVNo & "',2,'" & mVType & "'," & txt(21) & "," & ConvertDate(PubStartDate) & ",'" & PartNo & _
                "','" & PubSprCounterGodown & "',1,0," & VNull(Rs!MRP_TPStk) & "," & VNull(Rs!MRP_Rate) & "," & Round(VNull(Rs!MRP_TPStk) * VNull(Rs!MRP_Rate), 2) & "," & VNull(Rs!MRP_Rate) & _
                ",'" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
        End If
        If Rs!TB_Stk <> 0 Then
            GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
                "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,MRP_Rate," & _
                "Site_Code,U_Name,U_EntDt,U_AE) " & _
                "VALUES ('" & mVNo & "',3,'" & mVType & "'," & txt(21) & "," & ConvertDate(PubStartDate) & ",'" & PartNo & _
                "','" & PubSprCounterGodown & "',0,1," & VNull(Rs!TB_Stk) & "," & VNull(Rs!TB_Rate) & "," & Round(VNull(Rs!TB_Stk) * VNull(Rs!TB_Rate), 2) & _
                ",0,'" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
        End If
        If VoucherEditFlag = False Then      ' if Voucher Numbering Method is Automatic
            UpdVouSrlNo GCnFaS, mVType, PubStartDate
        End If
    End If
    Rs.MoveNext
    LBLCNT(3).CAPTION = Cnt
    LBLCNT(3).Refresh
    Cnt = Cnt + 1
Loop
GCn.CommitTrans
Set Rs = Nothing
MsgBox "Updation Complete For !! Part  !!", vbInformation, "Finish Job"
Exit Sub
ELoop:
If err.NUMBER = -2147467259 Then
    Cnt1 = Cnt1 + 1: LBLCNT(7).CAPTION = Cnt1: LBLCNT(7).Refresh
    GCn.Execute ("insert into prnmissrec(code,colname,details) values('" & PartNo & "','Part No','" & mID(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
Else
    GCn.Execute ("insert into prnmissrec(code,colname,details) values('" & PartNo & "','Part No','" & mID(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End If
End Sub

Private Sub CmdPath_Click()
On Error GoTo ELoop
Dim DataPath As String
DataPath = txt(FilePath) 'Text1.Text
If DataPath = "" Then MsgBox "Give DataPath for importing database", vbInformation, "Blank Path": Text1.SetFocus: Exit Sub
If OpenCon = True Then OldGcn.Close
Set OldGcn = New ADODB.Connection
With OldGcn
    .ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dBASE Files;Initial Catalog=" & DataPath & ""
    .Open
    OpenCon = True
    MsgBox "Connected Successfully", vbInformation, "Connection"
    End With
Exit Sub
ELoop:
If err.NUMBER = -2147467259 Then MsgBox "Database Directory not found", vbExclamation, "Unrecognised database": Text1.SetFocus: Exit Sub
End Sub

Private Sub CmdPrn_Click()
Dim mQry As String, RepFileName$, RepTitle$
Dim Rst As ADODB.Recordset
Dim I As Integer
On Error GoTo ERRORHANDLER
    mQry = "SELECT * from PrnMissRec"
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    RepFileName = "PrnMissRec"
         
         CreateFieldDefFile Rst, PubRepoPath + "\" & RepFileName & ".ttx", True
          Set rpt = rdApp.OpenReport(PubRepoPath + "\" & RepFileName & ".RPT")
           rpt.Database.SetDataSource Rst
           rpt.ReadRecords
            Call Report_View(rpt, "Missing Data List Printing", , False)
            Set Rst = Nothing
            Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub CmdSel_Click(Index As Integer)
On Error GoTo ErrHandler
    mFileName = ""
  CommonDialog1.InitDir = Pub_DataPath
  ' Set CancelError is True
  CommonDialog1.CancelError = True
  CommonDialog1.DialogTitle = "Select XLS File for Stock Updation"
  'CommonDialog1.
  ' Set flags
  CommonDialog1.Flags = cdlOFNHideReadOnly
  ' Set filters
  CommonDialog1.Filter = "Excel Files (*.xls)|*.xls" '|Text Files (*.txt)|*.Txt"
  ' Specify default filter
  CommonDialog1.FilterIndex = 1
  ' Display the Open dialog box
  CommonDialog1.ShowOpen
  ' Display name of selected file
  txt(FilePath) = CommonDialog1.FileName
  mFileName = CommonDialog1.FileTitle
  
ErrHandler:
  'User pressed the Cancel button
  Exit Sub

End Sub

Private Sub Form_Load()
WinSetting Me, 6000, 11200
If PubBackEnd = "A" Then
    Frame2.Visible = True
    Frame4.Visible = False
Else
    Frame2.Visible = False
    Frame4.Visible = True
End If
OpenCon = False
'Text1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set OldGcn = Nothing
End Sub

Private Sub PartOpen()
Dim Rs As ADODB.Recordset, rsName As ADODB.Recordset
Dim SecurityGrade$, PartNo$, Part_Grade$, PartName$, LocalName$, PartNameHlp$, PartNoHlp$
Dim DiscFactor$, BinLoca$
Dim Cnt As Long, mAdded As Long, MRPEffectDt As Date, TBEffectDt As Date
Dim Cnt1 As Long, mTrans As Boolean
Dim VoucherEditFlag As Boolean, mVNo$, mVType$
Dim XlsConn As New ADODB.Connection
On Error GoTo ELoop
'***********
If Trim(txt(FilePath)) = "" Then
    MsgBox "File not Selected", vbCritical, "Select Import File"
    CmdSel(0).SetFocus
    Exit Sub
End If

If MsgBox("Start Stock Updation", vbQuestion + vbYesNo, "Confirmation") = vbNo Then Exit Sub

If XlsConn.State <> 0 Then XlsConn.Close
XlsConn.CursorLocation = adUseClient
XlsConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txt(FilePath) & ";Extended Properties=Excel 8.0"
XlsConn.Open


'XlsConn.CursorLocation = adUseClient
'XlsConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft Office\Office10\1033\FPNWIND.MDB;Persist Security Info=False"
'XlsConn.Open

If PubBackEnd = "A" Then
    
    GSQL = "Select Part_No,Disc_Factor,Bin_Loca,sum(" & vIsNull("MRP_Rate", "0") & ") as MRPRate,sum(" & vIsNull("TB_SRate", "0") & ") as TBSRate, " & _
        "sum(" & vIsNull("TP_SRate", "0") & ") as TPSRate, " & _
        "sum(" & vIsNull("MRP_TBSTK", "0") & ") as MRPTBStk,sum(" & vIsNull("MRP_TPSTK", "0") & ") as MRPTPStk, " & _
        "sum(" & vIsNull("TB_Stk", "0") & ") as TBStk,sum(" & vIsNull("TP_Stk", "0") & ") as TPStk ,Part_Grade "
    Set Rs = GCn.Execute(GSQL & " From [Part_Import$] IN '" & txt(FilePath) & "' 'EXCEL 8.0;' " & _
        "Where Part_No<>'' and Part_Name<>'' " & _
        "Group by Part_No,Part_Name,Disc_Factor,Bin_Loca,Part_Grade;")
    
Else
    Set Rs = XlsConn.Execute("Select * From [Part_Import$] In '" & txt(FilePath) & "' 'EXCEL 8.0;' ")
    
End If
If Rs.RecordCount = 0 Then MsgBox "No Data found for Stock Updation!", vbInformation, "Validation": Set Rs = Nothing: Exit Sub

With ProgressBar1
    .Min = 0
    .Max = 100
    .Value = 0
    .Visible = True
End With
LblName(2).Visible = True
LblName(0).Visible = True
LBLCNT(2).CAPTION = Rs.RecordCount
LBLCNT(2).Refresh
LBLCNT(5).CAPTION = "Part"
LBLCNT(5).Refresh
mVType = "SXAO"
'***********
'GCn.BeginTrans
'GCn.Execute "update Part_Import set Div_Code='" & PubDivCode & "'"
'GCn.CommitTrans

'GSQL = "select Part_No,Div_Code,Disc_Factor,Bin_Loca,sum(iif(isnull(MRP_Rate),0,MRP_Rate)) as MRPRate,sum(iif(isnull(TB_SRate),0,TB_SRate)) as TBSRate,sum(iif(isnull(TP_SRate),0,TP_SRate)) as TPSRate, " & _
'"sum(iif(isnull(MRP_TBSTK),0,MRP_TBStk)) as MRPTBStk,sum(iif(isnull(MRP_TPSTK),0,MRP_TPStk)) as MRPTPStk," & _
'"sum(iif(isnull(TB_Stk),0,TB_Stk)) as TBStk,sum(iif(isnull(TP_Stk),0,TP_Stk)) as TPStk " & _
'"from Part_Import Group by Part_No,Div_Code,Disc_Factor,Bin_Loca"
'
'Set rs = New ADODB.Recordset
'rs.CursorLocation = adUseClient
'rs.Open GSQL, GCn, adOpenDynamic, adLockOptimistic

'GSQL = "select Part_No&'" & PubDivCode & "' as PartNoDiv,Part_No," & PubDivCode & ",Part_Name,bin_loca from part_import Order by Part_No&'" & PubDivCode & "'"



If PubBackEnd = "A" Then
    GSQL = "select Part_No&'" & PubDivCode & "' as PartNoDiv,Part_No,'" & PubDivCode & "' as Div_Code,Part_Name,Bin_Loca "
    Set rsName = XlsConn.Execute(GSQL & " From [Part_Import$]  " & _
            "Where Part_No<>'' and Part_Name<>'' " & _
            "Order by Part_No&'" & PubDivCode & "';")
End If



GCn.BeginTrans
mTrans = True
'Delete Opening if Exists
'GCn.Execute ("Delete from SP_Stock where v_type='" & mVType & "' and left(DocID,1)='" & PubDivCode & "' and Part_No in (" & GSQL & ")")
'If rs.RecordCount  > 0 Then
'    Do While rs.EOF
'        GCn.Execute ("Delete from SP_Stock where v_type='" & mVType & "' and left(DocID,1)='" & PubDivCode & "' and Part_No='" & rs!Part_No & "'")

'        If rs.EOF Then
'            Exit Do
'        Else
'            rs.MoveNext
'        End If
'    Loop
'End If

GCn.Execute ("Delete from SP_Stock where v_type='" & mVType & "' and left(DocID,1)='" & PubDivCode & "' and " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' ")
'**EOF Delete Operation
mSrlNo = 0: lblPrefix = ""
Rs.MoveFirst
Do Until Rs.EOF
    ProgressBar1.Value = Round(Rs.AbsolutePosition * 100 / Rs.RecordCount, 0)
    ProgressBar1.Refresh
    LblName(0).CAPTION = ProgressBar1.Value & "%"
    LblName(0).Refresh
    '*********
    PartNo = Replace(IIf(IsNull(Rs!Part_No), "", Rs!Part_No), "'", "`")
    If Len(PartNo) > 40 Then
        PartNo = left(PartNo, 22)
    End If
    
    If PubBackEnd = "A" Then
        rsName.MoveFirst
        rsName.FIND ("PartNoDiv='" & PartNo & PubDivCode & "'")
        
        If rsName.EOF = False Then
            PartName = Replace(IIf(IsNull(rsName!Part_Name), "", rsName!Part_Name), "'", "`")
            If Len(PartName) > 40 Then
                PartName = left(PartName, 40)
            End If
            BinLoca = IIf(IsNull(rsName!Bin_Loca), "", rsName!Bin_Loca)
            If Len(BinLoca) > 15 Then
                BinLoca = left(BinLoca, 15)
            End If
        Else
            PartName = mPartNo & " * Name not found *"
            BinLoca = ""
        End If
    Else
        PartName = left(XNull(Rs!Part_Name), 40)
        BinLoca = left(XNull(Rs!Bin_Loca), 15)
    End If
    LocalName = ""
    PartNameHlp = UCase(Replace(PartName, " ", ""))
    PartNoHlp = UCase(Replace(PartNo, " ", ""))
    MRPEffectDt = PubStartDate
    TBEffectDt = PubStartDate
    BinLoca = IIf(IsNull(Rs!Bin_Loca), "", Rs!Bin_Loca)
    DiscFactor = IIf(IsNull(Rs!Disc_Factor), "X", Rs!Disc_Factor)
    Part_Grade = IIf(IsNull(Rs!Part_Grade), "S", Rs!Part_Grade)
    
    GSQL = "Select Part_No from Part where Part_NoHelp='" & PartNoHlp & "' and Div_Code='" & PubDivCode & "'"
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenDynamic, adLockOptimistic

'    If GCn.Execute(GSQL).RecordCount  > 0 Then
    If GRs.RecordCount > 0 Then
'        PartNo = GCn.Execute(GSQL).Fields(0).Value
        PartNo = GRs!Part_No
    End If
    GSQL = "Select Part_No from Part where part_no='" & PartNo & "' and Div_Code='" & PubDivCode & "'"
    If GCn.Execute(GSQL).RecordCount <= 0 Then 'Not Found
        GCn.Execute "insert into part( " & _
            "Part_No,Div_Code,Site_Code,Part_Name,Local_Name,Part_NoHelp, " & _
            "Part_NameHelp,UNIT,Photo,Part_Grade, " & _
            "Security_Grade,Active_YN,Value_Method,Lead_Time,Min_Lvl,Max_Lvl,ReOrd_Lvl,Disc_Factor," & _
            "Bin_Loca,New_Part,MRP_Effect_Dt,MRP,TB_Effect_Dt,TB_SRate,TP_SRate," & _
            "Cur_MRP_TBStk,Cur_MRP_TPStk,Cur_TB_Stk,Cur_TP_Stk," & _
            "U_Name,U_EntDt,U_AE) " & _
            "values('" & Replace(PartNo, " ", "") & "','" & PubDivCode & "','" & PubSiteCode & "','" & PartName & "','" & LocalName & "','" & PartNoHlp & _
            "','" & PartNameHlp & "','PCS',' ','" & Part_Grade & "'" & _
            ",'A',0,'FIFO',0,0,0,0,'" & DiscFactor & "','" & BinLoca & "',1," & ConvertDate(MRPEffectDt) & "," & VNull(Rs!MRPRate) & "," & ConvertDate(TBEffectDt) & "," & VNull(Rs!TBSRate) & "," & VNull(Rs!TPSRate) & _
            "," & VNull(Rs!MRPTBStk) & "," & VNull(Rs!MRPTPStk) & "," & VNull(Rs!TBStk) & "," & VNull(Rs!TPStk) & _
            ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
        mAdded = mAdded + 1
    Else
        GCn.Execute "Update part set Bin_Loca='" & BinLoca & "',MRP_Effect_Dt=" & ConvertDate(MRPEffectDt) & _
            ",TB_Effect_Dt=" & ConvertDate(TBEffectDt) & ",Cur_MRP_TBStk=" & VNull(Rs!MRPTBStk) & ",Cur_MRP_TPStk=" & VNull(Rs!MRPTPStk) & _
            ",Cur_TB_Stk=" & VNull(Rs!TBStk) & ",Cur_TP_Stk=" & VNull(Rs!TPStk) & _
            ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' " & _
            " where Part_No='" & PartNo & "' and Div_Code='" & PubDivCode & "'"
    End If
    'insert opening rec in sp_stock
    If VNull(Rs!MRPTBStk) + VNull(Rs!MRPTPStk) + VNull(Rs!TBStk) + VNull(Rs!TPStk) <> 0 Then
        mVNo = GetDocID(GCnFaS, mVType, PubStartDate, VoucherEditFlag, txt(21), Label3(16), PubSiteCode)
        If VNull(Rs!MRPTBStk) > 0 Then
            mSrlNo = mSrlNo + 1
            GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
                "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,MRP_Rate," & _
                "Site_Code,U_Name,U_EntDt,U_AE,V_Rate) " & _
                "VALUES ('" & mVNo & "'," & mSrlNo & ",'" & mVType & "'," & txt(21) & "," & ConvertDate(PubStartDate - 1) & ",'" & PartNo & _
                "','" & PubSprCounterGodown & "',1,1," & VNull(Rs!MRPTBStk) & "," & VNull(Rs!MRPRate) & "," & Round(VNull(Rs!MRPTBStk) * VNull(Rs!MRPRate), 2) & "," & VNull(Rs!MRPRate) & _
                ",'" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A'," & VNull(Rs!MRPRate) & ")"
        End If
        
        If VNull(Rs!MRPTPStk) <> 0 Then
            mSrlNo = mSrlNo + 1
            GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
                "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,MRP_Rate," & _
                "Site_Code,U_Name,U_EntDt,U_AE,V_Rate) " & _
                "VALUES ('" & mVNo & "'," & mSrlNo & ",'" & mVType & "'," & txt(21) & "," & ConvertDate(PubStartDate - 1) & ",'" & PartNo & _
                "','" & PubSprCounterGodown & "',1,0," & VNull(Rs!MRPTPStk) & "," & VNull(Rs!MRPRate) & "," & Round(VNull(Rs!MRPTPStk) * VNull(Rs!MRPRate), 2) & "," & VNull(Rs!MRPRate) & _
                ",'" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A'," & VNull(Rs!MRPRate) & ")"
        End If
        
        If VNull(Rs!TBStk) <> 0 Then
            mSrlNo = mSrlNo + 1
            GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
                "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,MRP_Rate," & _
                "Site_Code,U_Name,U_EntDt,U_AE,V_Rate) " & _
                "VALUES ('" & mVNo & "'," & mSrlNo & ",'" & mVType & "'," & txt(21) & "," & ConvertDate(PubStartDate - 1) & ",'" & Replace(PartNo, " ", "") & _
                "','" & PubSprCounterGodown & "',0,1," & VNull(Rs!TBStk) & "," & VNull(Rs!TBSRate) & "," & Round(VNull(Rs!TBStk) * VNull(Rs!TBSRate), 2) & ",0 " & _
                ",'" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A'," & VNull(Rs!TBSRate) & ")"
        End If
      
        If VNull(Rs!TPStk) <> 0 Then
            mSrlNo = mSrlNo + 1
            GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No," & _
                "Godown,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,MRP_Rate," & _
                "Site_Code,U_Name,U_EntDt,U_AE,V_Rate) " & _
                "VALUES ('" & mVNo & "'," & mSrlNo & ",'" & mVType & "'," & txt(21) & "," & ConvertDate(PubStartDate - 1) & ",'" & PartNo & _
                "','" & PubSprCounterGodown & "',0,0," & VNull(Rs!TPStk) & "," & VNull(Rs!TPSRate) & "," & Round(VNull(Rs!TPStk) * VNull(Rs!TPSRate), 2) & ",0 " & _
                ",'" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A'," & VNull(Rs!TPSRate) & ")"
        End If
        If VoucherEditFlag = False Then      ' if Voucher Numbering Method is Automatic
            UpdVouSrlNo GCnFaS, mVNo, PubStartDate
        End If
    End If
    Rs.MoveNext
    Cnt = Cnt + 1
    LBLCNT(3).CAPTION = "A:" & mAdded & "/M:" & Cnt - mAdded & "/T:" & Cnt
    LBLCNT(3).Refresh
Loop
GCn.CommitTrans
MsgBox "Updation completed !"
Exit Sub
ELoop:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

