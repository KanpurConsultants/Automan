VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDIForm1"
   ClientHeight    =   8190
   ClientLeft      =   1230
   ClientTop       =   825
   ClientWidth     =   11880
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   OLEDropMode     =   1  'Manual
   Picture         =   "MDIForm1.frx":0442
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      FillColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   11850
      TabIndex        =   1
      Top             =   7470
      Visible         =   0   'False
      Width           =   11880
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   45
         TabIndex        =   2
         Top             =   30
         Width           =   825
      End
   End
   Begin MSComctlLib.StatusBar SBAR 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7875
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4939
            MinWidth        =   4939
            Text            =   "a Dataman Software"
            TextSave        =   "a Dataman Software"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1940
            MinWidth        =   1940
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2187
            MinWidth        =   2187
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2187
            MinWidth        =   2187
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2187
            MinWidth        =   2187
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "12:57 PM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "Find Part"
            TextSave        =   "Find Part"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2999
            MinWidth        =   2999
            Text            =   "Login Site"
            TextSave        =   "Login Site"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Master 
      Caption         =   "&Common Master"
      Begin VB.Menu Mas 
         Caption         =   "City Master"
         Index           =   1
      End
      Begin VB.Menu Mas 
         Caption         =   "State Master"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Mas 
         Caption         =   "Ledger A/c Master"
         Index           =   3
      End
      Begin VB.Menu Mas 
         Caption         =   "Color Master"
         Index           =   4
      End
      Begin VB.Menu Mas 
         Caption         =   "Contract/OEM"
         Index           =   5
      End
      Begin VB.Menu Mas 
         Caption         =   "Employee Master"
         Index           =   6
      End
      Begin VB.Menu Mas 
         Caption         =   "Godown Master"
         Index           =   7
      End
      Begin VB.Menu Mas 
         Caption         =   "Vehicle Model Category Master"
         Index           =   8
      End
      Begin VB.Menu Mas 
         Caption         =   "Vehicle Model Group Master"
         Index           =   9
      End
      Begin VB.Menu Mas 
         Caption         =   "Vehicle Model Check List"
         Index           =   10
      End
      Begin VB.Menu Mas 
         Caption         =   "Vehicle Model Master"
         Index           =   11
      End
      Begin VB.Menu Mas 
         Caption         =   "Dealer Master"
         Index           =   12
      End
      Begin VB.Menu Mas 
         Caption         =   "Area Master"
         Index           =   13
      End
      Begin VB.Menu Mas 
         Caption         =   "Site Master"
         Index           =   14
      End
      Begin VB.Menu Mas 
         Caption         =   "RateType Master"
         Index           =   15
      End
      Begin VB.Menu Mas 
         Caption         =   "Insurance Company Master"
         Index           =   16
      End
      Begin VB.Menu Mas 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu Mas 
         Caption         =   "Tax Forms"
         Index           =   18
      End
      Begin VB.Menu Mas 
         Caption         =   "Tax Forms Receipt / Issue"
         Index           =   19
      End
      Begin VB.Menu Mas 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu Mas 
         Caption         =   "Deprecation Item Master"
         Index           =   21
      End
      Begin VB.Menu Mas 
         Caption         =   "Deprecation Master"
         Index           =   22
      End
      Begin VB.Menu Mas 
         Caption         =   "-"
         Index           =   23
      End
   End
   Begin VB.Menu fa 
      Caption         =   "&Financial A/c"
      Begin VB.Menu fam 
         Caption         =   "&Master"
         Begin VB.Menu famM 
            Caption         =   "A/c &Group Entry"
            Index           =   0
         End
         Begin VB.Menu famM 
            Caption         =   "&Ledger A/c Entry"
            Index           =   1
         End
         Begin VB.Menu famM 
            Caption         =   "Narration Master"
            Index           =   2
         End
         Begin VB.Menu famM 
            Caption         =   "City Master"
            Index           =   3
         End
         Begin VB.Menu famM 
            Caption         =   "-"
            Index           =   4
         End
      End
      Begin VB.Menu faE 
         Caption         =   "&Transaction"
         Begin VB.Menu fame 
            Caption         =   "&Voucher Entry"
            Index           =   0
         End
         Begin VB.Menu fame 
            Caption         =   "Adjustment Entry"
            Index           =   1
         End
         Begin VB.Menu fame 
            Caption         =   "Adjustment Delete"
            Index           =   2
         End
         Begin VB.Menu fame 
            Caption         =   "Bank Reconciliation"
            Index           =   3
         End
         Begin VB.Menu fame 
            Caption         =   "Other Purchases"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu fame 
            Caption         =   "A/c Posting"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu fame 
            Caption         =   "TDS Category"
            Index           =   6
         End
         Begin VB.Menu fame 
            Caption         =   "TDS Challan"
            Index           =   7
         End
         Begin VB.Menu fame 
            Caption         =   "TDS Certificate"
            Index           =   8
         End
         Begin VB.Menu fame 
            Caption         =   "FaClosing"
            Index           =   9
         End
         Begin VB.Menu fame 
            Caption         =   "Expences Budgeting"
            Index           =   10
         End
         Begin VB.Menu fame 
            Caption         =   "Employee Wise Expence Entry"
            Index           =   11
         End
         Begin VB.Menu fame 
            Caption         =   "Cheque Payment Entry"
            Index           =   12
            Visible         =   0   'False
         End
         Begin VB.Menu fame 
            Caption         =   "-"
            Index           =   13
         End
         Begin VB.Menu fame 
            Caption         =   "Import From CrmDms"
            Index           =   14
         End
         Begin VB.Menu fame 
            Caption         =   "CrmDms Parameter Settings"
            Index           =   15
         End
         Begin VB.Menu fame 
            Caption         =   "Data Synchronisation"
            Index           =   16
         End
         Begin VB.Menu fame 
            Caption         =   "-"
            Index           =   17
            Visible         =   0   'False
         End
         Begin VB.Menu fame 
            Caption         =   "Import Inventory From CrmDms"
            Index           =   18
            Visible         =   0   'False
         End
         Begin VB.Menu fame 
            Caption         =   "CrmDms Inventory Parameters"
            Index           =   19
            Visible         =   0   'False
         End
      End
      Begin VB.Menu faD 
         Caption         =   "&Display"
         Begin VB.Menu FAREPORTD 
            Caption         =   "Balanc&e Sheet"
            Index           =   0
         End
         Begin VB.Menu FAREPORTD 
            Caption         =   "&Profit And Loss Account"
            Index           =   1
         End
         Begin VB.Menu FAREPORTD 
            Caption         =   "&Trial Balance (Group)"
            Index           =   2
         End
         Begin VB.Menu FAREPORTD 
            Caption         =   "&Trial Balance (Ledger)"
            Index           =   3
         End
         Begin VB.Menu FAREPORTD 
            Caption         =   "Cash &Flow"
            Index           =   4
         End
         Begin VB.Menu FAREPORTD 
            Caption         =   "Fund Flo&w"
            Index           =   5
         End
         Begin VB.Menu FAREPORTD 
            Caption         =   "Cash And Bank Books"
            Index           =   6
         End
         Begin VB.Menu FAREPORTD 
            Caption         =   "-"
            Index           =   7
         End
      End
      Begin VB.Menu faDS 
         Caption         =   "&Site Display"
         Begin VB.Menu FAREPORTDS 
            Caption         =   "Balanc&e Sheet (Site Wise)"
            Index           =   0
         End
         Begin VB.Menu FAREPORTDS 
            Caption         =   "&Profit And Loss Account (Site Wise)"
            Index           =   1
         End
         Begin VB.Menu FAREPORTDS 
            Caption         =   "&Trial Balance (Group) (Site Wise)"
            Index           =   2
         End
         Begin VB.Menu FAREPORTDS 
            Caption         =   "&Trial Balance (Ledger) (Site Wise)"
            Index           =   3
         End
         Begin VB.Menu FAREPORTDS 
            Caption         =   "Cash &Flow (Site Wise)"
            Index           =   4
         End
         Begin VB.Menu FAREPORTDS 
            Caption         =   "Fund Flo&w (Site Wise)"
            Index           =   5
         End
         Begin VB.Menu FAREPORTDS 
            Caption         =   "Cash And Bank Books (Site Wise)"
            Index           =   6
         End
         Begin VB.Menu FAREPORTDS 
            Caption         =   "-"
            Index           =   7
         End
      End
      Begin VB.Menu farp 
         Caption         =   "&Reports"
         Begin VB.Menu FAREPORT 
            Caption         =   "&Day Book"
            Index           =   0
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "&Ledger"
            Index           =   3
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "&Interest Ledger"
            Index           =   4
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Cas&h Book"
            Index           =   5
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Ban&k Book"
            Index           =   6
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "&Journal Books"
            Index           =   7
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "&Annexure"
            Index           =   10
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Bank &Register"
            Index           =   13
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "A&geing Analysis for Debtors"
            Index           =   14
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Ageing A&nalysis for Creditors"
            Index           =   15
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Ch&eque Cleared Register"
            Index           =   22
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Che&que Not Cleared Register"
            Index           =   23
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Outstanding Report For Debtors"
            Index           =   24
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Outstanding Report For Creditors"
            Index           =   25
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Daily Transaction Summary"
            Index           =   27
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Non Transaction Report"
            Index           =   29
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Reference Report"
            Index           =   30
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Detailed Trial Ledger"
            Index           =   31
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "A/c Check List"
            Index           =   32
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Bill Wise Outstanding Report For Debtors"
            Index           =   33
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Control Ledger"
            Index           =   34
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Roz Namcha"
            Index           =   35
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Cash/Bank Book"
            Index           =   36
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Expence Budget Variation Report"
            Index           =   37
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Salesman Sale Against Cost Report"
            Index           =   38
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Bill Wise Outstanding"
            Index           =   39
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "CRM DMS Tax Summary"
            Index           =   40
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "Cheque Payment Register"
            Index           =   41
         End
         Begin VB.Menu FAREPORT 
            Caption         =   "-"
            Index           =   42
         End
      End
      Begin VB.Menu Saperator 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu MnuLst 
      Caption         =   "&List/Register"
      Begin VB.Menu MnuLst1 
         Caption         =   "City Register"
         Index           =   0
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "State Register"
         Index           =   1
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Employee Register"
         Index           =   2
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Contract/ Finance Register"
         Index           =   3
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Discount Factor Register"
         Index           =   4
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Proprietory Part Grade Register"
         Index           =   5
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Godown Register"
         Index           =   6
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Unit Register"
         Index           =   7
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Aggregate Register"
         Index           =   8
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Part Register"
         Index           =   9
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Vehicle Model Category Register"
         Index           =   10
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Vehicle Model Group Register"
         Index           =   11
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Vehicle Model Master"
         Index           =   12
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Dealer Register"
         Index           =   13
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Colour Register"
         Index           =   14
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Labour Register"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "Model-Wise Labour Register"
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu MnuLst1 
         Caption         =   "-"
         Index           =   17
      End
   End
   Begin VB.Menu MnuVeh 
      Caption         =   "&Vehicle"
      Begin VB.Menu VehSalTrn 
         Caption         =   "Quotation Entry"
         Index           =   1
      End
      Begin VB.Menu VehSalTrn 
         Caption         =   "Booking"
         Index           =   2
      End
      Begin VB.Menu VehSalTrn 
         Caption         =   "Vehicle Customer Receipt"
         Index           =   3
      End
      Begin VB.Menu VehSalTrn 
         Caption         =   "Sale Bill"
         Index           =   4
      End
      Begin VB.Menu VehSalTrn 
         Caption         =   "Delivery Challan"
         Index           =   5
      End
      Begin VB.Menu Sep0 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu VehMktTrn 
         Caption         =   "Prospective Customer Entry"
         Index           =   0
      End
      Begin VB.Menu VehMktTrn 
         Caption         =   "Daily Activity Entry"
         Index           =   1
      End
      Begin VB.Menu VehMktTrn 
         Caption         =   "Order Got/Lost"
         Index           =   2
      End
      Begin VB.Menu VehMktReports 
         Caption         =   "Marketing Reports"
         Begin VB.Menu VehMktRep 
            Caption         =   "Daily Activity Report"
            Index           =   1
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Appointments"
            Index           =   2
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Call Status Report"
            Index           =   3
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Case Analysis"
            Index           =   4
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Daily Activity Missing Report"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Appointments Not Kept"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Executive-wise Got/Lost Report"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Pipeline Report (Daily)"
            Index           =   8
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Profession/Purpose Analysis"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Statement of Prospective Demands (SPADE)"
            Index           =   10
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Daily Sales Report"
            Index           =   12
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Sales Tracking Report"
            Index           =   13
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "Finance Tracking Report"
            Index           =   14
         End
         Begin VB.Menu VehMktRep 
            Caption         =   "-"
            Index           =   15
         End
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "Vehicle Receipts"
         Index           =   1
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "Vehicle Purchase"
         Index           =   2
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "Opening Stock"
         Index           =   3
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "Check List Entry"
         Index           =   4
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "BMS Opening"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "Godown Transfer"
         Index           =   6
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "Offtake Target Entry"
         Index           =   7
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "Vehicle Allocation"
         Index           =   8
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "Vehicle Rate List"
         Index           =   9
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "Vehicle Issue Entry"
         Index           =   10
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "Subvention Master"
         Index           =   11
      End
      Begin VB.Menu VehPurTrn 
         Caption         =   "Offtake Master"
         Index           =   12
      End
      Begin VB.Menu BodyBuild 
         Caption         =   "BodyBuilding"
         Begin VB.Menu BB 
            Caption         =   "Body Builder Chassis Issue"
            Index           =   0
         End
         Begin VB.Menu BB 
            Caption         =   "Body Type Master"
            Index           =   1
         End
         Begin VB.Menu BB 
            Caption         =   "Body Builder Master"
            Index           =   2
         End
         Begin VB.Menu BB 
            Caption         =   "Body Building Invoice"
            Index           =   3
         End
         Begin VB.Menu BBReport 
            Caption         =   "Reports"
            Begin VB.Menu BBRep 
               Caption         =   "Body Builder Wise Chassis Register"
               Index           =   0
            End
            Begin VB.Menu BBRep 
               Caption         =   "Stock At Body Builder"
               Index           =   1
            End
         End
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu MktData 
         Caption         =   "Competitor's Database"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu MktData 
         Caption         =   "RTO Database"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu VehMas 
         Caption         =   "Repwise-Modelwise Target"
         Index           =   1
      End
      Begin VB.Menu VehMas 
         Caption         =   "Financier Master"
         Index           =   2
      End
      Begin VB.Menu VehMas 
         Caption         =   "Addition / Deletion Items"
         Index           =   3
      End
      Begin VB.Menu VehMas 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu Doc 
         Caption         =   "Documents"
         Begin VB.Menu Doc1 
            Caption         =   "Vehicle Detail Card"
            Index           =   1
         End
         Begin VB.Menu Doc1 
            Caption         =   "TDS Certificate"
            Index           =   2
         End
         Begin VB.Menu Doc1 
            Caption         =   "-"
            Index           =   3
         End
      End
      Begin VB.Menu VehMas1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu VehRep1 
         Caption         =   "Daily Reports"
         Begin VB.Menu MnVehReports 
            Caption         =   "Money Receipt Register"
            Index           =   0
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Chassis Received Register"
            Index           =   1
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Purchase Register"
            Index           =   2
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Purchase Register (Summary)"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle In Transit Report"
            Index           =   4
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Booking Register"
            Index           =   5
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Retail Sale Report"
            Index           =   6
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Daily Sale Report (Own)"
            Index           =   7
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Sale Register"
            Index           =   8
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Sale Register (Summary)"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Delivery Challan Register"
            Index           =   10
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Additional Fitment Register"
            Index           =   11
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Stock Register"
            Index           =   12
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Stock For Bank"
            Index           =   13
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Stock Holding Report"
            Index           =   14
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Summary (Model-wise)"
            Index           =   15
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Sale Purchase Report"
            Index           =   16
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Target Collection Report"
            Index           =   17
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Quotation Register"
            Index           =   18
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Sale Cancel Register"
            Index           =   19
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Input Tax Register"
            Index           =   20
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "SalesMan wise Pending Amount"
            Index           =   21
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Outstanding Payment Report"
            Index           =   22
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Transfer Register"
            Index           =   23
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Output Tax Register"
            Index           =   24
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Vehicle Follow Up"
            Index           =   25
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Income Tax Register"
            Index           =   26
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "VAT Difference Register"
            Index           =   27
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Profitability"
            Index           =   28
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Do Pending Report"
            Index           =   29
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "Do Recive Report "
            Index           =   30
         End
         Begin VB.Menu MnVehReports 
            Caption         =   "-"
            Index           =   31
         End
      End
      Begin VB.Menu VehRep2 
         Caption         =   "Periodic Reports"
         Begin VB.Menu MnVehReports1 
            Caption         =   "Model/Branch-wise Report"
            Index           =   1
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Model/Branch-wise Sale Delivery"
            Index           =   2
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Monthly Target/Sale Qty Difference"
            Index           =   3
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Model/Financier-wise Sale Report"
            Index           =   4
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Model-wise/Group-wise Monthly"
            Index           =   5
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Area-wise/Model-wise Yearly"
            Index           =   6
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Financier/Model-wise Monthly"
            Index           =   7
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Model/Financier-wise Micro"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Model-wise Customer List"
            Index           =   9
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Sales/Purchase Price Difference"
            Index           =   10
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Delay Delivery Interest Report"
            Index           =   11
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Area-wise Financier-wise Sale"
            Index           =   12
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "E-Mail (Retail Sales)"
            Index           =   13
            Visible         =   0   'False
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Vehicle Sale Summary(Form-wise)"
            Index           =   14
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Vehicle Purchase/Sale Audit"
            Index           =   15
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Month Summary"
            Index           =   16
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Vehicle Profitability Report"
            Enabled         =   0   'False
            Index           =   17
            Visible         =   0   'False
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "MonthWise ModelWise Sales"
            Index           =   18
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "Model Wise Offtake and Sales"
            Index           =   19
         End
         Begin VB.Menu MnVehReports1 
            Caption         =   "-"
            Index           =   20
         End
      End
      Begin VB.Menu VehRep3 
         Caption         =   "Telco Reports"
         Begin VB.Menu TelRep 
            Caption         =   "Daily Sale Report "
            Index           =   0
         End
         Begin VB.Menu TelRep 
            Caption         =   "Customer Database (F001)"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu TelRep 
            Caption         =   "Loast Customer Information (F002)"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu TelRep 
            Caption         =   "Market Coverage Plan (MCP)  (F006)"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu TelRep 
            Caption         =   "Offtake Incentive Claim Letter"
            Index           =   4
         End
         Begin VB.Menu TelRep 
            Caption         =   "Subvention Letter"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu TelRep 
            Caption         =   "Financer Certificate"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu TelRep 
            Caption         =   "Subvention Claim Register"
            Index           =   7
         End
         Begin VB.Menu TelRep 
            Caption         =   "-"
            Index           =   8
         End
      End
   End
   Begin VB.Menu MnuSpr 
      Caption         =   "&Stores"
      Begin VB.Menu SprSalTrn 
         Caption         =   "Quotation"
         Index           =   1
      End
      Begin VB.Menu SprSalTrn 
         Caption         =   "Sale Order Entry"
         Index           =   2
      End
      Begin VB.Menu SprSalTrn 
         Caption         =   "Dispatch Challan Entry"
         Index           =   3
      End
      Begin VB.Menu SprSalTrn 
         Caption         =   "Sale Bill Entry"
         Index           =   4
      End
      Begin VB.Menu SprSalTrn 
         Caption         =   "Goods Return (Inward)"
         Index           =   5
      End
      Begin VB.Menu SprSalTrn 
         Caption         =   "Requisition Issue"
         Index           =   6
      End
      Begin VB.Menu SprSalTrn 
         Caption         =   "Requisition Returns"
         Index           =   7
      End
      Begin VB.Menu SprSalTrn 
         Caption         =   "Stock Adjustments"
         Index           =   8
      End
      Begin VB.Menu SprSalTrn 
         Caption         =   "Spare Customer Receipt"
         Index           =   9
         Shortcut        =   ^R
      End
      Begin VB.Menu SprSalTrn 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu SprPurTrn 
         Caption         =   "Purchase Order "
         Index           =   1
      End
      Begin VB.Menu SprPurTrn 
         Caption         =   "Goods Receipts"
         Index           =   2
      End
      Begin VB.Menu SprPurTrn 
         Caption         =   "Purchase Bill"
         Index           =   3
      End
      Begin VB.Menu SprPurTrn 
         Caption         =   "Goods Return (Outward)"
         Index           =   4
      End
      Begin VB.Menu SprPartSep 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu SprMas 
         Caption         =   "Aggregate Master"
         Index           =   1
      End
      Begin VB.Menu SprMas 
         Caption         =   "Discount Factor Master"
         Index           =   2
      End
      Begin VB.Menu SprMas 
         Caption         =   "Proprietory Part Grade Master"
         Index           =   3
      End
      Begin VB.Menu SprMas 
         Caption         =   "Part Master"
         Index           =   4
      End
      Begin VB.Menu SprMas 
         Caption         =   "Spare Price List"
         Index           =   5
      End
      Begin VB.Menu SprMas 
         Caption         =   "Price List Updation"
         Index           =   6
      End
      Begin VB.Menu SprMas 
         Caption         =   "Spare Sale Target"
         Index           =   7
      End
      Begin VB.Menu SprMas 
         Caption         =   "Physical Stock Updation"
         Index           =   8
      End
      Begin VB.Menu SprPurSep 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu SprPurRep1 
         Caption         =   "Purchase  Reports"
         Begin VB.Menu MnPurRep 
            Caption         =   "Purchase Order Register"
            Index           =   0
         End
         Begin VB.Menu MnPurRep 
            Caption         =   "Material Register"
            Index           =   1
         End
         Begin VB.Menu MnPurRep 
            Caption         =   "Purchase Register"
            Index           =   2
         End
         Begin VB.Menu MnPurRep 
            Caption         =   "Other Purchase Register"
            Index           =   3
         End
         Begin VB.Menu MnPurRep 
            Caption         =   "Purchase Return Register"
            Index           =   4
         End
         Begin VB.Menu MnPurRep 
            Caption         =   "Stock Transfer Register"
            Index           =   5
         End
         Begin VB.Menu MnPurRep 
            Caption         =   "Part-wise Purchase Report"
            Index           =   6
         End
         Begin VB.Menu MnPurRep 
            Caption         =   "Purchase Summary"
            Index           =   7
         End
         Begin VB.Menu MnPurRep 
            Caption         =   "Purchase Tax Summary"
            Index           =   8
         End
         Begin VB.Menu MnPurRep 
            Caption         =   "Spare Purchase Account"
            Index           =   9
         End
         Begin VB.Menu MnPurRep 
            Caption         =   "-"
            Index           =   10
         End
      End
      Begin VB.Menu SprSalesRep1 
         Caption         =   "Sales Reports"
         Begin VB.Menu MnSalesRep 
            Caption         =   "Quotation Register"
            Index           =   1
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Sale Order Register"
            Index           =   2
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Challan/Transfer Register"
            Index           =   3
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Counter (W/C)  Sales Register"
            Index           =   4
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Sales Return Register"
            Index           =   5
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Daily Sales(Spare+Lubes) Report(W/C)"
            Index           =   6
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Month-wise Sales(Spare+Lubes) Report(W/C)"
            Index           =   7
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Part-wise Sale Report"
            Index           =   8
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Sale Summary"
            Index           =   9
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "MRP Sales Tax Claims"
            Index           =   10
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Sales Tax Control Statement"
            Index           =   11
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Warranty Tax Reimburshment"
            Index           =   12
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Input Tax Register"
            Index           =   13
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Output Tax Register"
            Index           =   14
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Daily Lubricant Consumption"
            Index           =   15
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Sale Analysis "
            Index           =   16
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Sales Man Wise Outstanding Amt."
            Index           =   17
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Sale Tax Summary"
            Index           =   18
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "Spare Sale Account"
            Index           =   19
         End
         Begin VB.Menu MnSalesRep 
            Caption         =   "-"
            Index           =   20
         End
      End
      Begin VB.Menu StkRep1 
         Caption         =   "Stock Reports"
         Begin VB.Menu SprStkRep 
            Caption         =   "Stock Ledger"
            Index           =   1
         End
         Begin VB.Menu SprStkRep 
            Caption         =   "Stock Summary"
            Index           =   2
         End
         Begin VB.Menu SprStkRep 
            Caption         =   "Stock In Hand"
            Index           =   3
         End
         Begin VB.Menu SprStkRep 
            Caption         =   "Indent Register"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu SprStkRep 
            Caption         =   "Stock Above/Below/ReOrder Lavel"
            Index           =   5
         End
         Begin VB.Menu SprStkRep 
            Caption         =   "Bin-wise Spare Stock"
            Index           =   6
         End
         Begin VB.Menu SprStkRep 
            Caption         =   "Part Movement Register"
            Index           =   7
         End
         Begin VB.Menu SprStkRep 
            Caption         =   "Part Ageing Analysis"
            Index           =   8
         End
         Begin VB.Menu SprStkRep 
            Caption         =   "Stock Verification Sheet"
            Enabled         =   0   'False
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu SprStkRep 
            Caption         =   "Stock Value"
            Index           =   10
         End
         Begin VB.Menu SprStkRep 
            Caption         =   "-"
            Index           =   11
         End
      End
      Begin VB.Menu SprMIS 
         Caption         =   "MIS"
         Begin VB.Menu MnSprMIS 
            Caption         =   "Counter Rate Variation Report"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu MnSprMIS 
            Caption         =   "Purchase Rate Variation Report"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu MnSprMIS 
            Caption         =   "ABC Analysis"
            Index           =   3
         End
         Begin VB.Menu MnSprMIS 
            Caption         =   "FSN Analysis"
            Index           =   4
         End
         Begin VB.Menu MnSprMIS 
            Caption         =   "XYZ Analysis"
            Index           =   5
         End
         Begin VB.Menu MnSprMIS 
            Caption         =   "Stock Ledger (FIFO)"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu MnSprMIS 
            Caption         =   "Stock Valuation (FIFO)"
            Index           =   7
         End
         Begin VB.Menu MnSprMIS 
            Caption         =   "Part-wise Profitibility"
            Index           =   8
         End
         Begin VB.Menu MnSprMIS 
            Caption         =   "Sales Vs Inventory"
            Index           =   9
         End
         Begin VB.Menu MnSprMIS 
            Caption         =   "Inventory Projection Report"
            Index           =   10
         End
         Begin VB.Menu MnSprMIS 
            Caption         =   "-"
            Index           =   11
         End
      End
      Begin VB.Menu SprTaxForm1 
         Caption         =   "Tax Forms"
         Visible         =   0   'False
         Begin VB.Menu mnSprTaxForm 
            Caption         =   "Form Issue Against Spare Purchase"
            Index           =   1
         End
         Begin VB.Menu mnSprTaxForm 
            Caption         =   "Form Issue Against Spare Sale"
            Index           =   2
         End
         Begin VB.Menu mnSprTaxForm 
            Caption         =   "Form Issue Against Vehicle Purchase"
            Index           =   3
         End
         Begin VB.Menu mnSprTaxForm 
            Caption         =   "Form Received Against Vehicle Sale"
            Index           =   4
         End
         Begin VB.Menu mnSprTaxForm 
            Caption         =   "Road Permit Form Utilization Spare"
            Index           =   5
         End
         Begin VB.Menu mnSprTaxForm 
            Caption         =   "Road Permit Form Utilization Vehicle"
            Index           =   6
         End
         Begin VB.Menu mnSprTaxForm 
            Caption         =   "Form Reminder Spare"
            Index           =   7
         End
         Begin VB.Menu mnSprTaxForm 
            Caption         =   "Form Reminder Vehicle"
            Index           =   8
         End
         Begin VB.Menu mnSprTaxForm 
            Caption         =   "-"
            Index           =   9
         End
      End
   End
   Begin VB.Menu MnuWorks 
      Caption         =   "Works&hop"
      Begin VB.Menu WrkTrn 
         Caption         =   "Service Booking"
         Index           =   1
      End
      Begin VB.Menu WrkTrn 
         Caption         =   "Job Card Opening"
         Index           =   2
      End
      Begin VB.Menu WrkTrn 
         Caption         =   "Spare Requisition "
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu WrkTrn 
         Caption         =   "Works Gate Pass"
         Index           =   4
      End
      Begin VB.Menu WrkTrn 
         Caption         =   "Labour Done"
         Index           =   5
      End
      Begin VB.Menu WrkTrn 
         Caption         =   "Performa Labour"
         Index           =   6
      End
      Begin VB.Menu WrkTrn 
         Caption         =   "Supervisor Observation"
         Index           =   7
      End
      Begin VB.Menu WrkTrn 
         Caption         =   "Estimate"
         Index           =   8
      End
      Begin VB.Menu WrkTrn 
         Caption         =   "Job Close/Un Close"
         Index           =   9
      End
      Begin VB.Menu WrkTrn 
         Caption         =   "WorkShop Customer Receipt"
         Index           =   10
      End
      Begin VB.Menu WrkTrn 
         Caption         =   "Over Time Entry"
         Index           =   11
      End
      Begin VB.Menu W_Sep 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu WrkRep1 
         Caption         =   "Daily Reports"
         Begin VB.Menu WrkDRep 
            Caption         =   "Estimate Register"
            Index           =   1
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Performa Register"
            Index           =   2
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "JobCard Register"
            Index           =   3
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Internal Requisition Register"
            Index           =   4
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Part Grade-wise Requisition Register"
            Index           =   5
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Out-side Labour Register"
            Index           =   6
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Workshop Sale Register"
            Index           =   7
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Workshop Vehicle Diary"
            Index           =   8
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Sale Register"
            Index           =   9
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Gate Pass Register"
            Index           =   10
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Over Time Register"
            Index           =   11
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Veh History Register"
            Index           =   12
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Job Booking Register"
            Index           =   13
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Workshop Money Receipt Register"
            Index           =   14
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Workshop Vehicle Register"
            Index           =   15
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "Insurance Expiry Register"
            Index           =   16
         End
         Begin VB.Menu WrkDRep 
            Caption         =   "-"
            Index           =   17
         End
      End
      Begin VB.Menu WkShopRep1 
         Caption         =   "WorkShop Reports"
         Begin VB.Menu WkShop 
            Caption         =   "Labour Rate Variation Reports"
            Index           =   1
         End
         Begin VB.Menu WkShop 
            Caption         =   "Service Analysis Details"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu WkShop 
            Caption         =   "Service -wise Job Analysis"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu WkShop 
            Caption         =   "Model-wise Service Analysis"
            Index           =   4
         End
         Begin VB.Menu WkShop 
            Caption         =   "WorkShop Demand Vs Supply"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu WkShop 
            Caption         =   "Mechanic Earning Report"
            Index           =   6
         End
         Begin VB.Menu WkShop 
            Caption         =   "Mechanic Earning Summary"
            Index           =   7
         End
         Begin VB.Menu WkShop 
            Caption         =   "Dealer-wise Vehicle Attended"
            Index           =   8
         End
         Begin VB.Menu WkShop 
            Caption         =   "Model-wise Job Analysis"
            Index           =   9
         End
         Begin VB.Menu WkShop 
            Caption         =   "Aggregate Group-wise Inventory"
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu WkShop 
            Caption         =   "Job-wise Labour Analysis"
            Index           =   11
         End
         Begin VB.Menu WkShop 
            Caption         =   "Dealer-wise Job Analysis"
            Index           =   12
            Visible         =   0   'False
         End
         Begin VB.Menu WkShop 
            Caption         =   "Delay Reason Analysis"
            Index           =   13
         End
         Begin VB.Menu WkShop 
            Caption         =   "WorkShop Rate Variation Report"
            Index           =   14
         End
         Begin VB.Menu WkShop 
            Caption         =   "Cancellation Report(W/C)"
            Index           =   15
            Visible         =   0   'False
         End
         Begin VB.Menu WkShop 
            Caption         =   "Labour Incentive Report"
            Index           =   16
            Visible         =   0   'False
         End
         Begin VB.Menu WkShop 
            Caption         =   "Vehicle Grading Report"
            Index           =   17
         End
         Begin VB.Menu WkShop 
            Caption         =   "Service Tax Register"
            Index           =   18
         End
         Begin VB.Menu WkShop 
            Caption         =   "Labour Revenue Report"
            Index           =   19
         End
         Begin VB.Menu WkShop 
            Caption         =   "Service Due Register"
            Index           =   20
         End
         Begin VB.Menu WkShop 
            Caption         =   "Quarterly Return Register"
            Index           =   21
         End
         Begin VB.Menu WkShop 
            Caption         =   "Post Service Followups"
            Index           =   22
         End
         Begin VB.Menu WkShop 
            Caption         =   "Sales Tax With Serv. Tax"
            Index           =   23
         End
         Begin VB.Menu WkShop 
            Caption         =   "Repeat Job Analysis"
            Index           =   24
         End
         Begin VB.Menu WkShop 
            Caption         =   "-"
            Index           =   25
         End
      End
      Begin VB.Menu W_Sep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Claims1 
         Caption         =   "Claims Warranty Entry"
         Begin VB.Menu Warr 
            Caption         =   "Warranty Claim Data Entry"
            Index           =   2
         End
         Begin VB.Menu Warr 
            Caption         =   "Warranty Billing"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu Warr 
            Caption         =   "Warranty Cr Note / Rejection"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu Warr 
            Caption         =   "Complaint Code Master"
            Index           =   7
         End
         Begin VB.Menu Warr 
            Caption         =   "Failure Code Master"
            Index           =   8
         End
         Begin VB.Menu Warr 
            Caption         =   "Make Code Master"
            Index           =   9
         End
         Begin VB.Menu Warr 
            Caption         =   "Job Code Master"
            Index           =   10
         End
         Begin VB.Menu Warr 
            Caption         =   "-"
            Index           =   11
         End
      End
      Begin VB.Menu Claims2 
         Caption         =   "Claims Free Service/Warranty Reports"
         Begin VB.Menu WarrRep 
            Caption         =   "Warranty Claim Register"
            Index           =   0
         End
         Begin VB.Menu WarrRep 
            Caption         =   "Claim Not Made"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu WarrRep 
            Caption         =   "Claim Not Verified"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu WarrRep 
            Caption         =   "Claim not Dispatched"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu WarrRep 
            Caption         =   "Claim Rejected"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu WarrRep 
            Caption         =   "Claim Outstanding"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu WarrRep 
            Caption         =   "Claim Overveiw"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu WarrRep 
            Caption         =   "Part not Issued but Claimed"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu WarrRep 
            Caption         =   "Part Issued but Not claimed"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu WarrRep 
            Caption         =   "Claim Summary"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu WarrRep 
            Caption         =   "PDI Free Service Register"
            Index           =   10
         End
         Begin VB.Menu WarrRep 
            Caption         =   "Part Failure Report"
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu WarrRep 
            Caption         =   "FSB Upload Register"
            Index           =   12
         End
         Begin VB.Menu WarrRep 
            Caption         =   "-"
            Index           =   13
         End
      End
      Begin VB.Menu telcoReports1 
         Caption         =   "Tata Motors MIS Reports"
         Begin VB.Menu tlcorep 
            Caption         =   "Workshop Performence Report"
            Index           =   1
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Promise Time Deviation Report"
            Index           =   2
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Estimated Cost Deviation"
            Index           =   3
         End
         Begin VB.Menu tlcorep 
            Caption         =   "ModelWise Complaint Analysis"
            Index           =   4
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Aggregate Complaint Analysis"
            Index           =   5
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Repaire Order Analysis"
            Index           =   6
         End
         Begin VB.Menu tlcorep 
            Caption         =   "ModelWise Repeat Complaint Analysis"
            Index           =   7
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Aggregate Wise Repeat Analysis"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Repeat Job Analysis Form"
            Index           =   9
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Summery of Repeat Complaint"
            Index           =   10
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Repeat Job Analysis"
            Index           =   11
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Service Wise Job Analysis"
            Index           =   12
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Service Analysis Details"
            Index           =   13
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Quantitative Complaint Analysis"
            Index           =   14
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Customer Feedback Entry"
            Index           =   15
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Daily Follow Up Calls By CRO"
            Index           =   16
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Dissatisfied Customers"
            Index           =   17
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Customer Satisfaction Report"
            Index           =   18
         End
         Begin VB.Menu tlcorep 
            Caption         =   "Workshop Profit Report"
            Index           =   19
         End
         Begin VB.Menu tlcorep 
            Caption         =   "-"
            Index           =   20
         End
      End
      Begin VB.Menu W_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Service Master"
         Index           =   1
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Service Rate/Lube Qty. Declaration"
         Index           =   2
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Labour Group Master"
         Index           =   3
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Labour Type Master"
         Index           =   4
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Labour Description Master"
         Index           =   5
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Model-wise Labour Master"
         Index           =   6
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Model-wise Service-wise Pre-Defined Jobs"
         Index           =   7
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Reason For Job Delay"
         Index           =   8
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Trouble Master"
         Index           =   9
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Inspection Category Master"
         Index           =   10
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Inspection Elements Master"
         Index           =   11
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Workshop Vehicle Master"
         Index           =   12
      End
      Begin VB.Menu WrkMas 
         Caption         =   "Workshop Details Master"
         Index           =   13
      End
      Begin VB.Menu WrkMas 
         Caption         =   "-"
         Index           =   14
      End
   End
   Begin VB.Menu MisDoc 
      Caption         =   "&MIS Documents"
      Visible         =   0   'False
      Begin VB.Menu MisDoc1 
         Caption         =   "Customer Information"
         Index           =   1
      End
      Begin VB.Menu MisDoc1 
         Caption         =   "Customer Concern Entry"
         Index           =   2
      End
      Begin VB.Menu MisDoc1 
         Caption         =   "Finance WorkSheet "
         Index           =   3
      End
      Begin VB.Menu MisDoc1 
         Caption         =   "Vehicle Margin Statistics"
         Index           =   4
      End
      Begin VB.Menu MisDoc1 
         Caption         =   "-"
         Index           =   5
      End
   End
   Begin VB.Menu utlty 
      Caption         =   "&Utility"
      Begin VB.Menu Utl 
         Caption         =   "User Permissions"
         Index           =   1
      End
      Begin VB.Menu Utl 
         Caption         =   "User Groups"
         Index           =   2
      End
      Begin VB.Menu Utl 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu Utl 
         Caption         =   "System Controls"
         Index           =   4
      End
      Begin VB.Menu Utl 
         Caption         =   "Vehicle Controls"
         Index           =   5
      End
      Begin VB.Menu Utl 
         Caption         =   "Spare Controls"
         Index           =   6
      End
      Begin VB.Menu Utl 
         Caption         =   "Workshop Controls"
         Index           =   7
      End
      Begin VB.Menu Utl 
         Caption         =   "Voucher Prefix Generation"
         Index           =   8
      End
      Begin VB.Menu Utl 
         Caption         =   "Database  Backup"
         Index           =   9
      End
      Begin VB.Menu Utl 
         Caption         =   "Database  Restore"
         Index           =   10
      End
      Begin VB.Menu Utl 
         Caption         =   "Repair/Compact Database"
         Index           =   11
      End
      Begin VB.Menu Utl 
         Caption         =   "Update Table Structure"
         Index           =   12
      End
      Begin VB.Menu Utl 
         Caption         =   "FA Environment Setting"
         Index           =   13
      End
      Begin VB.Menu Utl 
         Caption         =   "FA Voucher Type Setting"
         Index           =   14
      End
      Begin VB.Menu Utl 
         Caption         =   "Year End Process"
         Index           =   15
      End
      Begin VB.Menu Utl 
         Caption         =   "Update Balances"
         Index           =   16
      End
      Begin VB.Menu Utl 
         Caption         =   "Update VRate"
         Index           =   17
      End
      Begin VB.Menu Utl 
         Caption         =   "A/C Merging"
         Index           =   18
      End
      Begin VB.Menu Utl 
         Caption         =   "Delete Log"
         Index           =   19
      End
      Begin VB.Menu Utl 
         Caption         =   "Update Voucher Counter"
         Index           =   20
      End
      Begin VB.Menu Utl 
         Caption         =   "Access Tables to Sql Server"
         Index           =   21
         Visible         =   0   'False
      End
   End
   Begin VB.Menu DTools 
      Caption         =   "&Data Tools"
      Begin VB.Menu DTool 
         Caption         =   "Data Send"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu DTool 
         Caption         =   "Data Received"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu DTool 
         Caption         =   "Data Updation"
         Index           =   2
      End
      Begin VB.Menu DTool 
         Caption         =   "Voucher Number System"
         Index           =   3
      End
      Begin VB.Menu DTool 
         Caption         =   "Make Database Blank"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu DTool 
         Caption         =   "Quick Report View"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu DTool 
         Caption         =   "Delete Zero Balance A/C"
         Index           =   6
      End
      Begin VB.Menu DTool 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu DTool 
         Caption         =   "Import From CRM DMS"
         Index           =   8
      End
      Begin VB.Menu DTool 
         Caption         =   "CRM DMS Parameter Settings"
         Index           =   9
      End
      Begin VB.Menu DTool 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu DTool 
         Caption         =   "Import From Tally"
         Index           =   11
      End
      Begin VB.Menu DTool 
         Caption         =   "Import Ledger & Opening"
         Index           =   12
      End
      Begin VB.Menu DTool 
         Caption         =   "Data Synchronisation"
         Index           =   13
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu End 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NewObj
Private Type DBDetails
    DataPath As String
    PassWord As String
End Type
Public AllowModuleVeh As Boolean
Public AllowModuleSpr As Boolean
Public AllowModuleWrk As Boolean

Private Const SearchDir As Byte = 0
Private Const CreateDir As Byte = 1
Private Const TransAuto As Byte = 2
Private Const TransFA As Byte = 3
Private Const CompProcess As Byte = 4
Dim G_CompCntmp As ADODB.Recordset
Dim G_Rsttmp As ADODB.Recordset
Dim j As Integer
Dim TRec1Qty As Single, TRec2Qty As Single, TQty As Double
Dim mTrf As Boolean, mRate As Double, mPART_ADD As Boolean
Dim mRec_TB_Qty As Double, mRec_TB_Val As Double, mRec_TP_Qty As Double, mRec_TP_Val As Double, StartDate$
Dim mIss_TB_Qty As Double, mIss_TB_Val As Double, mIss_TP_Qty As Double, mIss_TP_Val As Double
Dim xMOP_TBQty As Double, xMOP_TBVal As Double, xMOP_TPQty As Double, xMOP_TPVal As Double
Dim mOP_TB_QTY As Double, mOP_TP_QTY As Double, mOP_TB_VAL As Double, mOP_TP_VAL As Double
Dim mname$, mInv_No$, mInv_Date$, mNarr$
Dim mQry$, Condstr$, CondDivCode$, CondMarkYN$, CondPartNos$, CondPartNos1$, CondDivCode1$
Dim CondStrMRP$, CondPartNosOpStk$, CondPartNosTrn$
Dim RstPart As ADODB.Recordset, RstPart1 As ADODB.Recordset, RstStock As ADODB.Recordset, RstStock2 As ADODB.Recordset, RstStock3 As ADODB.Recordset
Dim NEW_CONNECTION As ADODB.Connection



'''''''Variables for changing field size , attributes etc'''''''''''''
Const CFT_Failed As Long = 55555
Private Const R_NAME = 0, R_ATTRIBUTES = 1, R_TABLE = 2, R_FOREIGNTABLE = 3, R_FIELD = 4, R_FOREIGNFIELD = 5
Private Const I_NAME = 0, I_PRIMARY = 1, I_UNIQUE = 2, I_REQUIRED = 3, I_IGNORENULLS = 4, I_CLUSTERED = 5, I_FIELD = 6, I_FIELDATTRIBUTES = 7

Private Const MC_City As Byte = 1
Private Const MC_LedgerAcMaster As Byte = 3
Private Const MC_Colour As Byte = 4
Private Const MC_Contract As Byte = 5
Private Const MC_EmpMast As Byte = 6
Private Const MC_Godown As Byte = 7
Private Const MC_ModelCat As Byte = 8
Private Const MC_ModelGroup As Byte = 9
Private Const MC_ModelIns As Byte = 10
Private Const MC_Model As Byte = 11
Private Const MC_Dealer As Byte = 12
Private Const MC_Area As Byte = 13
Private Const MC_Site As Byte = 14
Private Const MC_RateType As Byte = 15
Private Const MC_InsuranceCompany As Byte = 16
Private Const MC_TaxForm As Byte = 18
Private Const MC_TaxFormIssueReceive As Byte = 19



Private Sub BB_Click(Index As Integer)
    Select Case Index
        Case 0
            ShowForm FrmBodyBuilder_Chassis
        Case 1
            ShowForm FrmBodyType
        Case 2
            ShowForm FrmBodyBuilder
        Case 3
            ShowForm frmBodyBuilding
    End Select
End Sub

Private Sub BBRep_Click(Index As Integer)
Dim Menucaption$, ReportForm As Form
On Error GoTo errhand
Menucaption = UCase(Replace(Replace(BBRep(Index).CAPTION, "&", ""), " ", ""))
Set ReportForm = New RepFormCommon
Select Case Menucaption
    Case UCase(Replace("Body Builder Wise Chassis Register", " ", ""))
        ReportForm.GRepFormName = 4
    Case UCase(Replace("Stock At Body Builder", " ", ""))
        ReportForm.GRepFormName = 5
End Select
With ReportForm
    .CAPTION = BBRep(Index).CAPTION
    .LblTitle = BBRep(Index).CAPTION
    .Show
End With
Set ReportForm = Nothing
Exit Sub
errhand:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

'Private Const DayBook As Byte = 1
'Private Const ContraLed As Byte = 2
'Private Const Led As Byte = 3
'Private Const IntLed As Byte = 4
'Private Const CBook As Byte = 5
'Private Const BBook As Byte = 6
'Private Const JBooks As Byte = 7
'Private Const JBook2 As Byte = 8
'Private Const CrNoteBook As Byte = 9
'Private Const Annexure As Byte = 10
'Private Const Trial As Byte = 11
'Private Const SubTrial As Byte = 12
'Private Const BankReg As Byte = 13
'Private Const DrAge As Byte = 14
'Private Const CrAge As Byte = 15
'Private Const TradingPL As Byte = 16
'Private Const PL As Byte = 17
'Private Const Trading As Byte = 18
'Private Const BalSheet As Byte = 19
'Private Const CFlow As Byte = 20
'Private Const FFlow As Byte = 21
'Private Const ChqClReg As Byte = 22
'Private Const ChqNotClReg As Byte = 23
'Private Const DrOutStanding As Byte = 24
'Private Const CrOutStanding As Byte = 25
'Private Const CashBankCurBal As Byte = 26
'Private Const DailyTrnSum As Byte = 27
'Private Const JrnlBook As Byte = 28








Private Sub Doc1_Click(Index As Integer)
PubUParam = Permission(Replace(Doc1(Index).CAPTION, "&", ""))
Select Case Index
    Case 1
       If frmVehChasDCard.Visible = False Then frmVehChasDCard.Show: Set frmVehChasDCard = Nothing
    Case 2
       If frmVehTDSCert.Visible = False Then frmVehTDSCert.Show: Set frmVehTDSCert = Nothing
End Select
End Sub

Private Sub dtool_Click(Index As Integer)
Dim NewObj
Select Case Index
        Case 0
            If frmDataSend.Visible = False Then frmDataSend.Show: Set frmDataSend = Nothing
        Case 1
            If frmDataRecd.Visible = False Then frmDataRecd.Show: Set frmDataRecd = Nothing
        Case 2
            If FrmDataUpdation.Visible = False Then FrmDataUpdation.Show: Set FrmDataUpdation = Nothing
        Case 3
            MsgBox "Sample Voucher No.= C1/SPIC/10001001 " & vbCrLf & _
                    " Break Up:                  " & vbCrLf & _
                    "       C = Division Code    " & vbCrLf & _
                    "       1 = Site Code        " & vbCrLf & _
                    "       / = Separater        " & vbCrLf & _
                    "       S = Module Name (Spr)" & vbCrLf & _
                    "       PI= Purchase Invoice " & vbCrLf & _
                    "       C = Cash             " & vbCrLf & _
                    "       / = Separator        " & vbCrLf & _
                    "10001001 = Serial No.       ", vbInformation, "Voucher No.System"
        Case 5
            frmQuickRep.Show
        Case 6
            DeleteSubgroup
        Case 8
            FrmCrmDmsImport.mFormType = 1
            ShowForm FrmCrmDmsImport, , DTool(Index).CAPTION
        Case 9
            FrmCrmDmsImport.mFormType = 2
            ShowForm FrmCrmDmsImport, , DTool(Index).CAPTION
        Case 11
            ShowForm FrmTallyImport, , DTool(Index).CAPTION
        Case 12
            ShowForm FrmUpdateAccountOpening
        Case 13
            ShowForm FrmSynchroniseData
    End Select
End Sub

Private Sub DtTool_Click(Index As Integer)

End Sub





'Private Sub FATrn_Click(Index As Integer)
'PubUParam = Permission(Replace(FATrn(Index).CAPTION, "&", ""))
'If PubUParam = "****" Then Exit Sub
'Select Case Index
'    Case 1  'Voucher Entry
'        If frmAccVoucher.Visible = False Then frmAccVoucher.Show: Set frmAccVoucher = Nothing
'    Case 2  'Other Purchase Entry
'        If frmPurOth.Visible = False Then frmPurOth.Show: Set frmPurOth = Nothing
'    Case 3  'Spare A/c Posting
'        If frmAcPost.Visible = False Then frmAcPost.Show: Set frmAcPost = Nothing
'    Case 5  'Ledger A/c Entry
'        If frmSubGroup.Visible = False Then frmSubGroup.Show: Set frmSubGroup = Nothing
'    Case 6  'A/c Group Entry
'        If frmGrEnt.Visible = False Then frmGrEnt.Show: Set frmGrEnt = Nothing
'End Select
'End Sub

Private Sub famM_Click(Index As Integer)
On Error GoTo ErrLoop
PubUParam = Permission(Replace(famM(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Select Case Index
    Case 0
        
        Set NewObj = New FaGrEnt
        NewObj.Show
    Case 1
        
        Set NewObj = New frmSubGroup
        NewObj.Show
    Case 2
        Set NewObj = New FaNarrMast
        NewObj.Show
    Case 3
        
        Set NewObj = New frmCity
        NewObj.Show
End Select
Set NewObj = Nothing
Exit Sub
ErrLoop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub fame_Click(Index As Integer)
Dim key As String
PubUParam = Permission(Replace(fame(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
On Error GoTo ErrLoop
Select Case Index
    Case 0
       
        Set NewObj = New FaVrEnt
        NewObj.Show
    Case 1
        Set NewObj = New FaAdjust
        NewObj.Show
    Case 2
        Set NewObj = New FaAdjustDel
        NewObj.Show
    Case 3
        Set NewObj = New FaChqClear
        NewObj.Show
    Case 4
        Set NewObj = New frmPurOth
        NewObj.Show
    Case 5  'Spare A/c Posting
'        key = InputBox("Enter the Key word to open ", "Key Word")
'        If key = "2827272" Then
            Set NewObj = New frmAcPost
            NewObj.Show
'        Else
'            MsgBox "Invalid key.Please Contact Dataman for valid Key"
'            Exit Sub
'        End If
    Case 6
        Set NewObj = New FaTDSCat
        NewObj.Show
    Case 7
        Set NewObj = New FaTDSChal
        NewObj.Show
    Case 8
        Set NewObj = New FaTDSCertificate
        NewObj.Show
    Case 9
        Set NewObj = New FaClosing
        NewObj.Show
    Case 10
        Set NewObj = New FrmBudgetExp
        NewObj.Show
    Case 11
        Set NewObj = New FrmExpEntEmpWise
        NewObj.Show
    Case 12
        If FrmChequePayment.Visible = False Then FrmChequePayment.Show: Set FrmChequePayment = Nothing
    
    Case 14
        FrmCrmDmsImport.mFormType = 1
        ShowForm FrmCrmDmsImport, , fame(Index).CAPTION
    Case 15
        FrmCrmDmsImport.mFormType = 2
        ShowForm FrmCrmDmsImport, , fame(Index).CAPTION
    Case 16
        ShowForm FrmSynchroniseData, , fame(Index).CAPTION
    Case 18
        FrmCrmDmsInventoryImport.mFormType = 1
        ShowForm FrmCrmDmsInventoryImport, , fame(Index).CAPTION
    Case 19
        FrmCrmDmsInventoryImport.mFormType = 2
        ShowForm FrmCrmDmsInventoryImport, , fame(Index).CAPTION
        
  End Select
Set NewObj = Nothing
Exit Sub
ErrLoop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub FAREPORT_Click(Index As Integer)
Dim X11 As Variant, NEW_OBJ
On Error GoTo ErrLoop
    Select Case Index
        Case 0, 3, 4, 5, 6, 7, 10, 13, 14, 15, 22, 23, 24, 25, 27, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41
            Set NEW_OBJ = New FaReports
            Select Case Index
                Case 0      'DayBook
                    NEW_OBJ.GRepFormName = "DayBook"
                Case 3      'Ledger
                    NEW_OBJ.GRepFormName = "Led"
                    NEW_OBJ.BtnParam.Visible = True
                Case 4      'InterestLedger
                    NEW_OBJ.GRepFormName = "LedInt"
                    NEW_OBJ.BTNPRINT(1).Visible = False
                    NEW_OBJ.BtnParam.Visible = True
                Case 5      'CashBook
                    NEW_OBJ.GRepFormName = "CashBook"
                    NEW_OBJ.BtnParam.Visible = True
                Case 6      'BankBook
                    NEW_OBJ.GRepFormName = "BankBook"
                    NEW_OBJ.BtnParam.Visible = True
                Case 7 'Journal Book
                    NEW_OBJ.GRepFormName = "JournalBook"
                Case 10 'Annexure
                    NEW_OBJ.GRepFormName = "Annexure"
                    NEW_OBJ.BtnParam.Visible = True
                Case 13
                    NEW_OBJ.GRepFormName = "BankReg"
                    NEW_OBJ.BTNPRINT(1).Visible = False
                    NEW_OBJ.BtnParam.Visible = True
                Case 14 'Aging Analysis
                    NEW_OBJ.GRepFormName = "AgingDr"
                    NEW_OBJ.BTNPRINT(1).Visible = False
                    NEW_OBJ.BtnParam.Visible = True
                Case 15 'Aging Analysis
                    NEW_OBJ.GRepFormName = "AgingCr"
                    NEW_OBJ.BTNPRINT(1).Visible = False
                    NEW_OBJ.BtnParam.Visible = True
                Case 22 'Chq Cleared
                    NEW_OBJ.GRepFormName = "Clg"
                    NEW_OBJ.BTNPRINT(1).Visible = False
                    NEW_OBJ.BtnParam.Visible = True
                Case 23 'Chq Not Cleared
                    NEW_OBJ.GRepFormName = "ClgNot"
                    NEW_OBJ.BTNPRINT(1).Visible = False
                    NEW_OBJ.BtnParam.Visible = True
                Case 24
                    NEW_OBJ.GRepFormName = "LedDeb"
                    NEW_OBJ.BTNPRINT(1).CAPTION = "Letter"
                    NEW_OBJ.BtnParam.Visible = True
                Case 25
                    NEW_OBJ.GRepFormName = "LedCred"
                    NEW_OBJ.BTNPRINT(1).Visible = False
                    NEW_OBJ.BtnParam.Visible = True
                Case 27 'Daily Transaction Summary
                    NEW_OBJ.GRepFormName = "DailySumm"
                    NEW_OBJ.BtnParam.Visible = True
                Case 29 'Non Transaction Report
                    NEW_OBJ.GRepFormName = "NonTrans"
                    NEW_OBJ.BTNPRINT(1).Visible = False
                Case 30 'Reference Report
                    NEW_OBJ.GRepFormName = "RefReport"
                    NEW_OBJ.BTNPRINT(1).Visible = False
                Case 31 'Detailed Trial
                    NEW_OBJ.GRepFormName = "DetailedTrial"
                    NEW_OBJ.BTNPRINT(1).Visible = False
                Case 32 'Chq Not Cleared
                    NEW_OBJ.GRepFormName = "AcCheckList"
                    NEW_OBJ.BtnParam.Visible = True
                Case 33
                    NEW_OBJ.GRepFormName = "DUELIST"
                    NEW_OBJ.BTNPRINT(1).Visible = True
                Case 34
                    NEW_OBJ.GRepFormName = "CONTROLLED"
                Case 35      'RozNamcha
                    NEW_OBJ.GRepFormName = "RozNamcha"
                Case 36
                    Set NEW_OBJ = ReprtFormGlobal
                    NEW_OBJ.GRepFormName = 39
                    NEW_OBJ.Show
                Case 37
                    Set NEW_OBJ = ReprtFormGlobal
                    NEW_OBJ.GRepFormName = 42
                    NEW_OBJ.Show
                Case 38
                    Set NEW_OBJ = ReprtFormGlobal
                    NEW_OBJ.GRepFormName = 43
                    NEW_OBJ.Show
                Case 39
                    Set NEW_OBJ = ReprtFormGlobal
                    NEW_OBJ.GRepFormName = 44
                    NEW_OBJ.Show
                Case 40
                    Set NEW_OBJ = New RepFormCommon
                    NEW_OBJ.GRepFormName = 2
                    NEW_OBJ.Show
                Case 41
                    Set NEW_OBJ = New ReportVehicle
                    NEW_OBJ.GRepFormName = 30
                    NEW_OBJ.Show
            End Select
            NEW_OBJ.CAPTION = Replace(FAREPORT(Index).CAPTION, "&", "")
    End Select
    Set NEW_OBJ = Nothing
    Exit Sub
ErrLoop:    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub


Private Sub FAREPORTD_Click(Index As Integer)
Dim mOpenType As String
    Select Case Index
        Case 0
            mOpenType = ("BALSHEET")
        Case 1
            mOpenType = ("PROFLOSS")
        Case 2
            mOpenType = ("GROUPTRIAL")
        Case 3
            mOpenType = ("LEDTRIAL")
        Case 4
            mOpenType = ("CASHFLOW")
        Case 5
            mOpenType = ("FUNDFLOW")
        Case 6
            mOpenType = ("CASHBANKSUM")
    End Select
    If mOpenType <> "" Then
        Set NewObj = New FaMagic
        NewObj.ShowSiteWise = False
        NewObj.OpenType = mOpenType
        NewObj.CAPTION = Replace(FAREPORTD(Index).CAPTION, "&", "")
        NewObj.Show
    End If
    Set NewObj = Nothing
End Sub
Private Sub FAREPORTDS_Click(Index As Integer)
Dim mOpenType As String
    Select Case Index
        Case 0
            mOpenType = ("BALSHEET")
        Case 1
            mOpenType = ("PROFLOSS")
        Case 2
            mOpenType = ("GROUPTRIAL")
        Case 3
            mOpenType = ("LEDTRIAL")
        Case 4
            mOpenType = ("CASHFLOW")
        Case 5
            mOpenType = ("FUNDFLOW")
        Case 6
            mOpenType = ("CASHBANKSUM")
    End Select
    If mOpenType <> "" Then
        Set NewObj = New FaMagic
        NewObj.ShowSiteWise = True
        NewObj.OpenType = mOpenType
'        NewObj.CAPTION = Replace(FAREPORTDS(Index).CAPTION, "&", "")
        NewObj.Show
    End If
    Set NewObj = Nothing
End Sub
Private Sub MDIForm_Activate()
On Error Resume Next
'Utl(27).Visible = False
 
If UCase(left(PubComp_Name, 5)) = "UJWAL" And (StrCmp(pubUName, "Sa") Or StrCmp(PubULabel, "y")) Then
    fame(13).Visible = True
    fame(14).Visible = True
    fame(15).Visible = True
    DTool(6).Visible = True
    DTool(8).Visible = True
    DTool(9).Visible = True
Else
    fame(13).Visible = False
    fame(14).Visible = False
    fame(15).Visible = False
    DTool(6).Visible = False
    DTool(8).Visible = False
    DTool(9).Visible = False
End If


If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
 DTool(8).Visible = True
 fame(17).Visible = True
 fame(18).Visible = True
 fame(19).Visible = True
End If


If UCase(left(PubComp_Name, 6)) = "PRAYAG" Then
    DTool(11).Visible = True
    fame(16).Visible = True
Else
    DTool(11).Visible = False
    fame(16).Visible = False
End If

If RSOJPR = True Then
    G_CompCn.Execute "Alter table UserMast add BckpY_N text (1) "
    G_CompCn.Execute "update UserMast set BckpY_N='0'"
    G_CompCn.Execute "Create table Backuplog (BackupDt date)"
    If G_CompCn.Execute("Select * from Backuplog").RecordCount = 0 Then
        G_CompCn.Execute "Insert into Backuplog values(#01/Jan/2005#)"
    End If
    G_CompCn.Execute ("Update UserMast set BckpY_N='1'")
    Dim CurrDate As Date, LastBkpDate As Date
    LastBkpDate = G_CompCn.Execute("Select BackupDt from Backuplog").Fields(0).Value
    CurrDate = G_FaCn.Execute("Select Max(V_Date) from Ledger").Fields(0).Value
    If IsNull(LastBkpDate) Or CDate(CurrDate) > CDate(LastBkpDate) Then
        FrmBackup.Show
        FrmBackup.Label1.CAPTION = "Dear " & pubUName & " ! You are the first user to login today.Please Wait, Software will perform following actions  automatically .."
        FrmBackup.Image1(0).Visible = True
        Call DataBackup
        FrmBackup.Image1(0).Visible = False
        FrmBackup.Image1(1).Visible = True
        FrmBackup.Image1(1).Refresh
        Call UpdtCurrBalances
        FrmBackup.Image1(1).Visible = False
        G_CompCn.Execute ("Update Backuplog set BackupDt=" & ConvertDate(CurrDate) & "")
        Unload FrmBackup
    End If
    G_CompCn.Execute ("Update UserMast set BckpY_N='0'")
    Utl(27).Visible = True
End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    End
End If
End Sub

Private Sub MisDoc1_Click(Index As Integer)
PubUParam = Permission(Replace(VehSalTrn(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Select Case Index
    Case 1     'Customer Information
        If frmCustInfo.Visible = False Then frmCustInfo.Show: Set frmCustInfo = Nothing
    Case 2     'Customer Concern
        If FrmCustConern.Visible = False Then FrmCustConern.Show: Set FrmCustConern = Nothing
    Case 3     'Finance Worksheet
        If FrmFinWrkSheet.Visible = False Then FrmFinWrkSheet.Show: Set FrmFinWrkSheet = Nothing
    Case 4     'vehicle margine
        If FrmVehMargin.Visible = False Then FrmVehMargin.Show: Set FrmVehMargin = Nothing
End Select
End Sub
Private Sub MnPurRep_Click(Index As Integer)
Dim rep As CrystalReport, Menucaption$, ReportForm As Form
PubUParam = Permission(Replace(MnPurRep(Index).CAPTION, "&", ""))
On Error GoTo errhand
Menucaption = UCase(Replace(Replace(MnPurRep(Index).CAPTION, "&", ""), " ", ""))
Set ReportForm = New ReprtFormGlobal
Select Case Menucaption
    Case UCase(Replace("Purchase Order Register", " ", ""))
        ReportForm.GRepFormName = 5
    Case UCase(Replace("Material Register", " ", ""))
        ReportForm.GRepFormName = 6
    Case UCase(Replace("Purchase Register", " ", ""))
        ReportForm.GRepFormName = 7
    Case UCase(Replace("Purchase Return Register", " ", ""))
        ReportForm.GRepFormName = 8
    Case UCase(Replace("Stock Transfer Register", " ", ""))
        ReportForm.GRepFormName = 9
    Case UCase(Replace("Part-wise Purchase Report", " ", ""))
        ReportForm.GRepFormName = 18    '25
    Case UCase(Replace("Purchase Summary", " ", ""))
        ReportForm.GRepFormName = 19 '27
    Case UCase(Replace("Other Purchase Register", " ", ""))
        ReportForm.GRepFormName = 30 '27
    Case UCase(Replace("Purchase Tax Summary", " ", ""))
        ReportForm.GRepFormName = 40 '27
    Case UCase(Replace("Spare Purchase Account", " ", ""))
        ReportForm.GRepFormName = 47 '27
        
'        mQRY = "Select SP.DocId,SP.V_No,TF.Tax_Per,SP.Tax_Amt, SP.Tot_Goods_Value, Sp.Addition, " & _
'            "Sp.Deduction, SG.Name + ', ' + " & xIsNull("C.CityName", "") & " As Name,SG.LSTNo, " & _
'            " SP.Party_Doc_No,Sp.Net_Amt,Convert(nVarChar,SP.V_Date,3) As V_Date,SP.Party_Doc_Date" & _
'            " From (((SP_Purch as SP " & _
'            " Left Join SubGroup Sg on SP.Party_Code=SG.SubCode) " & _
'            " Left Join City C On C.CityCode = SG.CityCode) " & _
'            " Left Join TaxForms TF on SP.Form_Code=TF.Form_Code) "
'        G_CompCn.Execute "Delete From AgReports_LastReport"
'        G_CompCn.Execute "Insert Into AgReports_LastReport Values ('Purchase Tax Summary', '" & Replace(mQRY, "'", "`") & "')"
'        Shell "D:\Dataman\Release\agreports.exe"
'        Exit Sub
End Select
With ReportForm
    .CAPTION = MnPurRep(Index).CAPTION
    .LblTitle = MnPurRep(Index).CAPTION
    .Show
End With
Set ReportForm = Nothing
    Exit Sub
errhand:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub MnSalesRep_Click(Index As Integer)
Dim rep As CrystalReport, Menucaption$, ReportForm As Form
PubUParam = Permission(Replace(MnSalesRep(Index).CAPTION, "&", ""))
On Error GoTo errhand
Menucaption = UCase(Replace(Replace(MnSalesRep(Index).CAPTION, "&", ""), " ", ""))
If Menucaption = UCase(Replace("Quotation Register", " ", "")) Then
    Set ReportForm = New ReportWorkShop2
Else
    Set ReportForm = New ReprtFormGlobal
End If
Select Case Menucaption
    Case UCase(Replace("Quotation Register", " ", ""))
        ReportForm.GRepFormName = 1
    Case UCase(Replace("Sale Order Register", " ", ""))
        ReportForm.GRepFormName = 2
    Case UCase(Replace("Challan/Transfer Register", " ", ""))
        ReportForm.GRepFormName = 32
    Case UCase(Replace("Counter (W/C)  Sales Register", " ", ""))
        ReportForm.GRepFormName = 3
    Case UCase(Replace("Sales Return Register", " ", ""))
        ReportForm.GRepFormName = 4
    Case UCase(Replace("Sales Tax Control Statement", " ", ""))
        ReportForm.GRepFormName = 31
    Case UCase(Replace("Input Tax Register", " ", ""))
        ReportForm.GRepFormName = 34
    Case UCase(Replace("Output Tax Register", " ", ""))
        ReportForm.GRepFormName = 35
    Case UCase(Replace("Daily Sales(Spare+Lubes) Report(W/C)", " ", ""))
        ReportForm.GRepFormName = 16
    Case UCase(Replace("Month-wise Sales(Spare+Lubes) Report(W/C)", " ", ""))
        ReportForm.GRepFormName = 17
    Case UCase(Replace("Part-wise Sale Report", " ", ""))
        ReportForm.GRepFormName = 20
    Case UCase(Replace("Sale Summary", " ", ""))
        ReportForm.GRepFormName = 21
    Case UCase(Replace("MRP Sales Tax Claims", " ", ""))
        ReportForm.GRepFormName = 29
    Case UCase(Replace("Warranty Tax Reimburshment", " ", ""))
        ReportForm.GRepFormName = 33
    Case UCase(Replace("Daily Lubricant Consumption", " ", ""))
        ReportForm.GRepFormName = 36
    Case UCase(Replace("Sale Analysis", " ", ""))
        ReportForm.GRepFormName = 37
    Case UCase(Replace("Sales Man Wise Outstanding Amt. ", " ", ""))
        ReportForm.GRepFormName = 38
    Case UCase(Replace("Sale Tax Summary", " ", ""))
        ReportForm.GRepFormName = 41
    Case UCase(Replace("Spare Sale Account", " ", ""))
        ReportForm.GRepFormName = 46
        
End Select
With ReportForm
    .CAPTION = MnSalesRep(Index).CAPTION
    .LblTitle = MnSalesRep(Index).CAPTION
    .Show
End With
Set ReportForm = Nothing
    Exit Sub
errhand:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub MnSprMIS_Click(Index As Integer)
On Error GoTo errhand
Dim rep As CrystalReport, Menucaption$, ReportForm As Form
PubUParam = Permission(Replace(MnSprMIS(Index).CAPTION, "&", ""))
On Error GoTo errhand
Menucaption = UCase(Replace(Replace(MnSprMIS(Index).CAPTION, "&", ""), " ", ""))
If Menucaption = UCase(Replace("Counter Rate Variation Report", " ", "")) Or _
    Menucaption = UCase(Replace("Purchase Rate Variation Report", " ", "")) Then
    Set ReportForm = New ReprtFormGlobal
Else
    Set ReportForm = New RepSprMIS
End If
Select Case Menucaption
    Case UCase(Replace("Counter Rate Variation Report", " ", ""))
        ReportForm.GRepFormName = 26    '34 ReprtFormGlobal
    Case UCase(Replace("Purchase Rate Variation Report", " ", ""))
        ReportForm.GRepFormName = 27    '35  ReprtFormGlobal
    Case UCase(Replace("ABC Analysis", " ", ""))
        ReportForm.GRepFormName = 1
    Case UCase(Replace("FSN Analysis", " ", ""))
        ReportForm.GRepFormName = 2
    Case UCase(Replace("XYZ Analysis", " ", ""))
        ReportForm.GRepFormName = 3
    Case UCase(Replace("Stock Ledger (FIFO)", " ", ""))
        ReportForm.GRepFormName = 4
    Case UCase(Replace("Stock Valuation (FIFO)", " ", ""))
        ReportForm.GRepFormName = 5
    Case UCase(Replace("Part-wise Profitibility", " ", ""))
        ReportForm.GRepFormName = 6
    Case UCase(Replace("Sales Vs Inventory", " ", ""))
        ReportForm.GRepFormName = 7
    Case UCase(Replace("Inventory Projection Report", " ", ""))
        ReportForm.GRepFormName = 8
End Select
With ReportForm
    .CAPTION = MnSprMIS(Index).CAPTION
'    .LblTitle = MnSprMIS(Index).Caption
    .Show
End With
Set ReportForm = Nothing
Set rep = Nothing
errhand: If err.NUMBER <> 0 Then CheckError
End Sub
Private Sub mnSprTaxForm_Click(Index As Integer)
Dim rep As CrystalReport
PubUParam = Permission(Replace(mnSprTaxForm(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Dim ReportForm As Form
On Error GoTo errhand
Set ReportForm = New RepFormTax
Select Case UCase(Replace(Replace(mnSprTaxForm(Index).CAPTION, "&", ""), " ", ""))
    Case UCase(Replace("Form Issue Against Spare Purchase", " ", ""))
        ReportForm.GRepFormName = 29
    Case UCase(Replace("Form Issue Against Spare Sale", " ", ""))
        ReportForm.GRepFormName = 30
    Case UCase(Replace("Form Issue Against Vehicle Purchase", " ", ""))
        ReportForm.GRepFormName = 31
    Case UCase(Replace("Form Received Against Vehicle Sale", " ", ""))
        ReportForm.GRepFormName = 32
    Case UCase(Replace("Road Permit Form Utilization Spare", " ", ""))
        ReportForm.GRepFormName = 33
    Case UCase(Replace("Road Permit Form Utilization Vehicle", " ", ""))
        ReportForm.GRepFormName = 34
    Case UCase(Replace("Form Reminder Spare", " ", ""))
        ReportForm.GRepFormName = 35
    Case UCase(Replace("Form Reminder Vehicle", " ", ""))
        ReportForm.GRepFormName = 36
End Select
With ReportForm
    .CAPTION = mnSprTaxForm(Index).CAPTION
    .LblTitle = mnSprTaxForm(Index).CAPTION
    .Show
End With
Set ReportForm = Nothing
    Exit Sub
errhand:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub mnuExit_Click()
    frmDivision.Show
    frmDivision.CAPTION = PubPackage & "-Return Menu" '& RsComp!comp_name
    Unload MDIForm1
'    Set RsComp = New Recordset
'    If UCase(pubUName) = "SA" Then
'        RsComp.Open "Select * from company", G_CompCn, adOpenDynamic, adLockOptimistic
'    Else
'        RsComp.Open "Select * from company where comp_code in (select distinct comp_code from user1 where user_name='" & pubUName & "')", G_CompCn, adOpenDynamic, adLockOptimistic
'    End If
'    Set Grid.DataSource = RsComp
'    frm
'    Load frmLogin
'    frmLogin.Show
End Sub
Private Sub MnuLst1_Click(Index As Integer)
On Error GoTo errhand
Dim rep As CrystalReport, Form1 As frmMastList
PubUParam = Permission(MnuLst1(Index).CAPTION)
Select Case Index
    Case 0      'City Master List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 1
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 1      'State Master List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 0
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 2      'Employee Master List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 2
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 3      'contract/finance List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 6
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
        
    Case 4      'discount factor List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 7
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 5      'propietory part grade List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 8
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
        
    Case 6     'godown List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 9
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 7      'unit List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 10
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 8      'Aggregate List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 5
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
        
    Case 9      'part List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 11
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 10      'Vehicle model category List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 3
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 11     'vehicle model group List
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 4
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
        
    Case 12         ' vehicle model master
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 12
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 13 'dealer master
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 13
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 14 'colour register
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 14
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 15     'labour description master
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 19
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
        
    Case 16     ' model-wise labour detail master
        Set Form1 = New frmMastList
        With Form1
            .g_FormID = 20
            .LblName.CAPTION = MnuLst1(Index).CAPTION
            .CAPTION = MnuLst1(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
End Select
Exit Sub

errhand:
    CheckError
End Sub
Private Sub Mas_Click(Index As Integer)
PubUParam = Permission(Replace(Mas(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Select Case Index
    Case MC_City      'City Master
        If frmCity.Visible = False Then frmCity.Show: Set frmCity = Nothing
    Case MC_Colour     'Color Master
        If frmColor.Visible = False Then frmColor.Show: Set frmColor = Nothing
    Case MC_Contract      'Contract/OEM
        If frmContract.Visible = False Then frmContract.Show:  Set frmContract = Nothing
    Case MC_EmpMast     'Employee Master
        If frmEmpMast.Visible = False Then frmEmpMast.Show: Set frmEmpMast = Nothing
    Case MC_Godown      'Godown Master
        If frmGodown.Visible = False Then frmGodown.Show: Set frmGodown = Nothing
    Case MC_ModelCat        'Vehicle Model Category Master
        If frmModelCat.Visible = False Then frmModelCat.Show: Set frmModelCat = Nothing
    Case MC_ModelGroup       'Vehicle Model Group Master
        If frmModelGrp.Visible = False Then frmModelGrp.Show: Set frmModelGrp = Nothing
    Case MC_ModelIns      'Model Check List
        If frmModelInspEle.Visible = False Then frmModelInspEle.Show: Set frmModelInspEle = Nothing
    Case MC_Model       'Vehicle Model Master
        If frmModel.Visible = False Then frmModel.Show: Set frmModel = Nothing
    Case MC_Dealer     'Dealer Master
        If frmDealer.Visible = False Then frmDealer.Show: Set frmDealer = Nothing
    Case MC_Area      'Area Master
        If frmArea.Visible = False Then frmArea.Show: Set frmArea = Nothing
    Case MC_Site      'Site Master
        If FrmSite.Visible = False Then FrmSite.Show: Set FrmSite = Nothing
    Case MC_TaxForm       'Tax Forms**
        If frmTaxForms.Visible = False Then frmTaxForms.Show: Set frmTaxForms = Nothing
    Case MC_TaxFormIssueReceive       'Tax Forms Receipt / Issue**
        If frmTaxFrmIssRec.Visible = False Then frmTaxFrmIssRec.Show: Set frmTaxFrmIssRec = Nothing
    Case MC_RateType
        If FrmRateType.Visible = False Then FrmRateType.Show: Set FrmRateType = Nothing
    Case MC_InsuranceCompany
        If FrmInsurance.Visible = False Then FrmInsurance.Show: Set FrmInsurance = Nothing
    Case MC_LedgerAcMaster
        If frmSubGroup.Visible = False Then frmSubGroup.Show: Set frmSubGroup = Nothing
    Case 21
        If FrmDeprecation_itemMaster.Visible = False Then FrmDeprecation_itemMaster.Show: Set FrmDeprecation_itemMaster = Nothing
           Case 22
            If FrmDeprecation_Master.Visible = False Then FrmDeprecation_Master.Show: Set FrmDeprecation_Master = Nothing
            
    
End Select
End Sub
Private Sub MDIForm_Load()
'On Error GoTo ELoop
Dim VitKDate As Date
Dim mSQry$


'If PubBackEnd = "S" Then
'    If UCase(Trim(G_FaCn.Execute("Select MakeDataBlank From Syctrl").Fields(0))) = "DATAMAN" Then
'        BlankDataSqlServer
'    End If
'End If


If StrCmp(left(PubComp_Name, 6), "prayag") Then
    PubDataSynchronisationApplicable = True
End If

frmDivision.ProgressBar1.Value = 55
Set rdApp = CreateObject("CrystalRunTime.Application")
If (pubUName = "SA" Or UCase(PubULabel) = "Y") Then
    utlty.Visible = True
Else
    utlty.Visible = False
End If
If UCase(left(PubComp_Name, 7)) = "JOHNSON" Then
    FAREPORT(36).Visible = True
Else
    FAREPORT(36).Visible = False
End If


ApplyConsolidatedPosting PubLoginDate

SBAR.Panels(2).TEXT = PubSecName
SBAR.Panels(3).TEXT = Format(PubStartDate, "yyyy") & "-" & Format(PubEndDate, "yyyy")
SBAR.Panels(4).TEXT = "User : " & pubUName
SBAR.Panels(5).TEXT = PubCenDataPath    'PubLoginDate

GSQL = "select Site_Desc + '[' + SiteType + ']' from Site where Site_Code='" & PubSiteCode & "'"
If GCn.Execute(GSQL).RecordCount > 0 Then
    SBAR.Panels(8).TEXT = GCn.Execute(GSQL).Fields(0).Value
Else
    MsgBox "Define Site in System Controls through Utility Menu", vbCritical, "Urgent"
End If


If PubDataSynchronisationApplicable Then
    SBAR.Panels(6).TEXT = "ReConnect"
End If


'Vit K
'VitKDate = CDate("14/11/2003")
'Set GRs = New ADODB.Recordset
'GRs.CursorLocation = adUseClient
'GRs.Open "Select max(Sal_VDate) as VDate from Veh_Stock", GCn, adOpenDynamic, adLockOptimistic
'If GRs.RecordCount  > 0 Then
'    If Not IsNull(GRs!Vdate) Then
'        If GRs!Vdate  > Format(VitKDate, "dd/MMM/yyyy") Then
'            End
'        End If
'    End If
'End If
'Set GRs = New ADODB.Recordset
'GRs.CursorLocation = adUseClient
'GRs.Open "Select max(V_Date) as VDate from Ledger", G_FACN, adOpenDynamic, adLockOptimistic
'If GRs.RecordCount  > 0 Then
'    If Not IsNull(GRs!Vdate) Then
'        If GRs!Vdate  > Format(VitKDate, "dd/MMM/yyyy") Then
'            End
'        End If
'    End If
'End If
'Set GRs = Nothing
'EOF

'GSQL = "Select count(*) from SP_Stock"
'If GCn.Execute(GSQL).Fields(0).Value  > 250 Then
'    MsgBox "License Expired", vbCritical, "License Checking"
'    End
'End If
'
'GSQL = "Select count(*) from Ledger"
'If G_FACN.Execute(GSQL).Fields(0).Value  > 250 Then
'    MsgBox "License Expired", vbCritical, "License Checking"
'    End
'End If

'Allow Module Menu
ModuleMenu
'23-09-2002
'Specially declared for Speed purpose, extra fields deleted
frmDivision.ProgressBar1.Value = 65
If PubSCompCode <> "" Or PubWCompCode <> "" Then
    Set RsPart = New ADODB.Recordset
    RsPart.CursorLocation = adUseClient
    
'    If PubBackEnd = "A" Then
'        RsPart.Open "Select P.Part_No AS Code," & xIsNull("P.Part_Name", "") & " as Name, " & xIsNull("P.Local_Name", "") & " AS LName, " & xIsNull("P.Unit", "") & " as Unit," _
'        & "format(P.MRP,'0.00') as MRP,format(P.TB_SRate,'0.00') as TB_SRate,format(P.TP_SRate,'0.00') as TP_SRate,P.Bin_Loca,DF.PurcDisc_Per,p.ReOrd_Lvl,P.Cur_MRP_TBStk,P.Cur_MRP_TPStk,P.Cur_TB_Stk,p.Cur_TP_Stk,format(P.Cur_MRP_TBStk+P.Cur_MRP_TPStk+P.Cur_TB_Stk+p.Cur_TP_Stk,'0.00') as CurrStk,P.Min_Lvl,p.Disc_Factor " _
'        & "From Part P left join Part_DiscFactor DF on P.Disc_Factor=DF.DiscFac_Catg  Where P.Div_Code='" & PubDivCode & "' " _
'        & "Order By P.Part_No,P.Part_Name,P.Local_Name", GCn, adOpenDynamic, adLockOptimistic
'    ElseIf PubBackEnd = "S" Then
'        Set RsPart = GCn.Execute("Select * From GlobalPart")
'    End If
    
    
            
            
    If GCn.Execute("Select CheckNegetiveStockSiteWise From Syctrl").Fields(0) = 1 Then
        mSQry = "Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock " & _
                "WHERE (V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " " & _
                "Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & ") " & _
                "And " & cMID("DocId", "2", "1") & "='" & PubSiteCode & "' And Part_No=P.Part_No "
    Else
        mSQry = "Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock " & _
                "WHERE (V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " " & _
                "Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & ") " & _
                "And Part_No=P.Part_No "
    End If
    
    
    
    
    If PubBackEnd = "S" Then
        Set RsPart = GCn.Execute("Select P.Part_No as Code, P.Part_No, P.Part_Name As Name, P.Local_Name as LName, P.Unit , P.MRP, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Part_Grade, " & _
                                        "(Select PurcDisc_Per From Part_DiscFactor Where DiscFac_Catg=P.Disc_Factor) As PurcDisc_Per, P.ReOrd_Lvl, " & _
                                        "(" & mSQry & " And Mrp_Yn=1 And Tax_Yn=1) As Cur_MRP_TBStk, (" & mSQry & " And Mrp_Yn=1 And Tax_Yn=0) As Cur_MRP_TpStk, " & _
                                        "(" & mSQry & " And Mrp_Yn=0 And Tax_Yn=1) As Cur_TB_Stk, (" & mSQry & " And Mrp_Yn=0 And Tax_Yn=0) As Cur_Tp_Stk, " & _
                                        "(" & mSQry & ") As CurrStk, P.Min_Lvl, P.Disc_Factor,max(p.dep_item) as Deptcode,max(ditm.dep_per) as dep_per " & _
                                        "From (Part P Left Join Sp_Stock Stk On P.Part_No=Stk.Part_No) left join Deprecation_itemMaster DITM on p.dep_item=ditm.code " & _
                                        "WHERE  Div_Code='" & PubDivCode & "' " & _
                                        "Group By P.Part_No, P.Part_Name, P.Local_Name, P.Unit, P.Mrp, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Disc_Factor, P.ReOrd_Lvl, P.Min_Lvl, P.Part_Grade")
        '(V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & " Or Stk.Part_No Is Null)  And
    Else
        Set RsPart = GCn.Execute("Select P.Part_No as Code, P.Part_No, P.Part_Name As Name, P.Local_Name as LName, P.Unit , Format(P.MRP,'0.00') As Mrp, Format(P.TB_SRate,'0.00') As TB_SRate, Format(P.Tp_SRate,'0.00') As Tp_SRate, P.Bin_Loca, P.Part_Grade, " & _
                                        "(Select PurcDisc_Per From Part_DiscFactor Where DiscFac_Catg=P.Disc_Factor) As PurcDisc_Per, P.ReOrd_Lvl, " & _
                                        "(" & mSQry & " And Mrp_Yn=1 And Tax_Yn=1) As Cur_MRP_TBStk, (" & mSQry & " And Mrp_Yn=1 And Tax_Yn=0) As Cur_MRP_TpStk, " & _
                                        "(" & mSQry & " And Mrp_Yn=0 And Tax_Yn=1) As Cur_TB_Stk, (" & mSQry & " And Mrp_Yn=0 And Tax_Yn=0) As Cur_Tp_Stk, " & _
                                        "Format(IIF(IsNull((" & mSQry & ")),0,(" & mSQry & ")),'0.00') As CurrStk, P.Min_Lvl, P.Disc_Factor " & _
                                        "From Part P Left Join Sp_Stock Stk On P.Part_No=Stk.Part_No " & _
                                        "WHERE  Div_Code='" & PubDivCode & "' " & _
                                        "Group By P.Part_No, P.Part_Name, P.Local_Name, P.Unit, P.Mrp, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Disc_Factor, P.ReOrd_Lvl, P.Min_Lvl, P.Part_Grade")
        '(V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & " Or Stk.Part_No Is Null)  And
    End If
    Set RsPartSiteWise = RsPart.Clone
    
'    Else
'        If PubBackEnd = "A" Then
'
''            Set RsPart = GCn.Execute("Select P.Part_No AS Code," & xIsNull("P.Part_Name", "") & " as Name, " & xIsNull("P.Local_Name", "") & " AS LName, " & xIsNull("P.Unit", "") & " as Unit," _
''            & "format(P.MRP,'0.00') as MRP,format(P.TB_SRate,'0.00') as TB_SRate,format(P.TP_SRate,'0.00') as TP_SRate,P.Bin_Loca,DF.PurcDisc_Per,p.ReOrd_Lvl,P.Cur_MRP_TBStk,P.Cur_MRP_TPStk,P.Cur_TB_Stk,p.Cur_TP_Stk,format(P.Cur_MRP_TBStk+P.Cur_MRP_TPStk+P.Cur_TB_Stk+p.Cur_TP_Stk,'0.00') as CurrStk,P.Min_Lvl,p.Disc_Factor " _
''            & "From Part P left join Part_DiscFactor DF on P.Disc_Factor=DF.DiscFac_Catg  Where P.Div_Code='" & PubDivCode & "' " _
''            & "Order By P.Part_No,P.Part_Name,P.Local_Name")
'
'        ElseIf PubBackEnd = "S" Then
'            Set RsPart = GCn.Execute("Select * From GlobalPart")
'        End If
'    End If
    
    

    
    
    'Order By P.Part_No,P.Part_Name,P.Local_Name
'    RsPart.Filter = ("Div_Code='" & PubDivCode & "'")
    frmDivision.ProgressBar1.Value = 75
'    RsPart.Sort = "Name"
    frmDivision.ProgressBar1.Value = 80
'    RsPart.Sort = "LName"
    frmDivision.ProgressBar1.Value = 85
'    RsPart.Sort = "Code"
    
  frmDivision.ProgressBar1.Value = 90
  If PubVATYN = 1 Then
      MnSalesRep(11).Visible = False
      MnSalesRep(13).Visible = True
      MnSalesRep(14).Visible = True
  Else
      MnSalesRep(11).Visible = True
      MnSalesRep(13).Visible = False
      MnSalesRep(14).Visible = False
  End If
  
  If Not StrCmp(left(PubComp_Name, 5), "Ujwal") Then
    'VehSalTrn(6).Visible = False
   MnVehReports(29).Visible = False
    MnVehReports(30).Visible = False
  End If
   If StrCmp(left(PubComp_Name, 6), "J.M.A.") Then
  
   MnVehReports(29).Visible = True
    MnVehReports(30).Visible = True
  End If
  
  If RSOJPR = True Then
    tlcorep(19).Visible = True
  Else
    tlcorep(19).Visible = False
  End If
    
    
'    If PubLoginDate <= CDate("03/APR/2007") Then
'        GCn.Execute "Delete From Job_Lab Where Job_DocId Not In (Select DocId From Job_Card)"
'        GCn.Execute "Delete From Job_Lab2 Where Job_DocId Not In (Select DocId From Job_Card)"
'    End If
    
    
    
'Killer

KillerSiteWise "Yash", "Mirza", "10/JUN/2013"
KillerSiteWise "Yash", "Chopan", "11/SEP/2013"
Killer "Enar", "18/Dec/2012"
Killer "Singhal", "02/May/2013"
Killer "J.M.A.", "30/Apr/2013"
Killer "NAC", "15/MAY/2013"



''Dim RsTemp As ADODB.Recordset
''Set RsTemp = GCn.Execute("Select Max(Inv_Date) From Veh_Order Where Inv_Date is Not Null ")
''If RsTemp.RecordCount > 0 Then
''    If StrCmp(left(PubComp_Name, 4), "Enar") Or StrCmp(left(PubComp_Name, 4), "Yash") Then
''        If DateDiff("D", RsTemp(0), CDate("23/Jul/2009")) < 0 Then
''            G_FaCn.BeginTrans
''                G_FaCn.Execute "Select * Into Master.dbo.SubGroup From SubGroup"
''                G_FaCn.Execute "Select * Into Master.dbo.Veh_Stock From Veh_Stock"
''                G_FaCn.Execute "Delete From SubGroup"
''                G_FaCn.Execute "Delete From SubGroupAlias"
''                G_FaCn.Execute "Delete From Veh_Stock"
''            G_FaCn.CommitTrans
''            MsgBox "Unexpected Error! Contact to Dataman."
''            End
''        End If
''    End If
''End If
    
    
    
End If

Exit Sub
ELoop:   CheckError
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    Set GRs = Nothing
    PubVFADataPath = ""
    Set GCnFaV = Nothing
    PubSFADataPath = ""
    Set GCnFaS = Nothing
    PubWFADataPath = ""
    Set GCnFaW = Nothing
    Set NewObj = Nothing
    Set RsPart = Nothing
    End
End Sub

Private Sub mnutstfrm_Click()
    'FrmCustConern.Show
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub
Private Sub MnVehReports_Click(Index As Integer)
Dim rep As CrystalReport
PubUParam = Permission(Replace(MnVehReports(Index).CAPTION, "&", ""))
Dim ReportForm As Form
On Error GoTo errhand

If Index = 0 Then
    Set ReportForm = New ReprtFormGlobal
ElseIf Index = 20 Or Index = 24 Or Index = 27 Or Index = 28 Or Index = 29 Or Index = 30 Then
    Set ReportForm = New ReportVehicle2
Else
    Set ReportForm = New ReportVehicle
End If



Select Case UCase(Replace(Replace(MnVehReports(Index).CAPTION, "&", ""), " ", ""))
    Case UCase(Replace("Money Receipt Register", " ", ""))
        ReportForm.GRepFormName = 13
    Case UCase(Replace("Chassis Received Register", " ", ""))
        ReportForm.GRepFormName = 1
    Case UCase(Replace("Vehicle Purchase Register", " ", ""))
        ReportForm.GRepFormName = 2
    Case UCase(Replace("Purchase Register (Summary)", " ", ""))
        ReportForm.GRepFormName = 16
    Case UCase(Replace("Sale Register (Summary)", " ", ""))
        ReportForm.GRepFormName = 17
    Case UCase(Replace("Vehicle In Transit Report", " ", ""))
        ReportForm.GRepFormName = 3
    Case UCase(Replace("Vehicle Booking Register", " ", ""))
        ReportForm.GRepFormName = 4
    Case UCase(Replace("Retail Sale Report", " ", ""))
        ReportForm.GRepFormName = 5
    Case UCase(Replace("Daily Sale Report (Own)", " ", ""))
        ReportForm.GRepFormName = 6 '10
    Case UCase(Replace("Vehicle Sale Register", " ", ""))
        Set ReportForm = New RepFormCommon
        ReportForm.GRepFormName = 1 '7
    Case UCase(Replace("Delivery Challan Register", " ", ""))
        ReportForm.GRepFormName = 8 '7
    Case UCase(Replace("Additional Fitment Register", " ", ""))
        ReportForm.GRepFormName = 9 '8
    Case UCase(Replace("Vehicle Stock Register", " ", ""))
        ReportForm.GRepFormName = 10 '9
    Case UCase(Replace("Vehicle Stock For Bank", " ", ""))
        ReportForm.GRepFormName = 11
    Case UCase(Replace("Vehicle Stock Holding Report", " ", ""))
        ReportForm.GRepFormName = 12
    Case UCase(Replace("Vehicle Summary (Model-wise)", " ", ""))
        Set ReportForm = RepFormCommon
        ReportForm.GRepFormName = 3
    Case UCase(Replace("Vehicle Sale Purchase Report", " ", ""))
        ReportForm.GRepFormName = 14
    Case UCase(Replace("Target Collection Report", " ", ""))
        ReportForm.GRepFormName = 15
    Case UCase(Replace("Vehicle Quotation Register", " ", ""))
        ReportForm.GRepFormName = 16
    Case UCase(Replace("Vehicle Sale Cancel Register", " ", ""))
        ReportForm.GRepFormName = 18
    Case UCase(Replace("Vehicle Sale Cancel Register", " ", ""))
        ReportForm.GRepFormName = 18
    Case UCase(Replace("Vehicle Input Tax Register", " ", ""))
        ReportForm.GRepFormName = 2
    Case UCase(Replace("SalesMan wise Pending Amount", " ", ""))
        ReportForm.GRepFormName = 19
    Case UCase(Replace("Outstanding Payment Report", " ", ""))
        ReportForm.GRepFormName = 20
    Case UCase(Replace("Vehicle Output Tax Register", " ", ""))
        ReportForm.GRepFormName = 3
    Case UCase(Replace("Vehicle Transfer Register", " ", ""))
        ReportForm.GRepFormName = 23
    Case UCase(Replace("Vehicle Follow Up", " ", ""))
        ReportForm.GRepFormName = 25
    Case UCase(Replace("Income Tax Register", " ", ""))
        ReportForm.GRepFormName = 26
    Case UCase(Replace("VAT Difference Register", " ", ""))
        ReportForm.GRepFormName = 4
    Case UCase(Replace("Profitability", " ", ""))
        ReportForm.GRepFormName = 5
    Case UCase(Replace("Cheque Payment Register", " ", ""))
        ReportForm.GRepFormName = 30
    Case UCase(Replace("Do Pending Report", " ", ""))
        ReportForm.GRepFormName = 31
    Case UCase(Replace("Do Recive Report", " ", ""))
        ReportForm.GRepFormName = 32
    
End Select
With ReportForm
    .CAPTION = MnVehReports(Index).CAPTION
    .LblTitle = MnVehReports(Index).CAPTION
    .Show
End With
Set ReportForm = Nothing
    Exit Sub
errhand:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub MnVehReports1_Click(Index As Integer)
Dim rep As CrystalReport
PubUParam = Permission(Replace(MnVehReports1(Index).CAPTION, "&", ""))
Dim ReportForm As Form
On Error GoTo errhand
Set ReportForm = New ReportVehicle1
Select Case UCase(Replace(Replace(MnVehReports1(Index).CAPTION, "&", ""), " ", ""))
    Case UCase(Replace("Model/Branch-wise Report", " ", ""))
        ReportForm.GRepFormName = 1
    Case UCase(Replace("Model/Branch-wise Sale Delivery", " ", ""))
        ReportForm.GRepFormName = 2
    Case UCase(Replace("Monthly Target/Sale Qty Difference", " ", ""))
        ReportForm.GRepFormName = 3
    Case UCase(Replace("Model/Financier-wise Sale Report", " ", ""))
        ReportForm.GRepFormName = 4
    Case UCase(Replace("Model-wise/Group-wise Monthly", " ", ""))
        ReportForm.GRepFormName = 5
    Case UCase(Replace("Area-wise/Model-wise Yearly", " ", ""))
        ReportForm.GRepFormName = 6
    Case UCase(Replace("Financier/Model-wise Monthly", " ", ""))
        ReportForm.GRepFormName = 7
    Case UCase(Replace("Model/Financier-wise Micro", " ", ""))
'        MsgBox "Available in Licence Version", vbInformation, "Validation"
'        Set ReportForm = Nothing
'        Exit Sub
        ReportForm.GRepFormName = 8
    Case UCase(Replace("Model-wise Customer List", " ", ""))
        ReportForm.GRepFormName = 9
    Case UCase(Replace("Sales/Purchase Price Difference", " ", ""))
        ReportForm.GRepFormName = 10
    Case UCase(Replace("Delay Delivery Interest Report", " ", ""))
        ReportForm.GRepFormName = 11
    Case UCase(Replace("Area-wise Financier-wise Sale", " ", ""))
        ReportForm.GRepFormName = 12
    Case UCase(Replace("E-Mail (Retail Sales)", " ", ""))
        ReportForm.GRepFormName = 13
    Case UCase(Replace("Vehicle Sale Summary(Form-wise)", " ", ""))
        ReportForm.GRepFormName = 14
    Case UCase(Replace("Vehicle Purchase/Sale Audit", " ", ""))
        ReportForm.GRepFormName = 15
    Case UCase(Replace("Month Summary", " ", ""))
        ReportForm.GRepFormName = 16
    Case UCase(Replace("MonthWise ModelWise Sales", " ", ""))
        ReportForm.GRepFormName = 17
    Case UCase(Replace("Vehicle Profitability Report", " ", ""))
        frmVehProfit.Show
        Exit Sub
    Case UCase(Replace("Model Wise Offtake and Sales", " ", ""))
        Set ReportForm = RepFormCommon
        ReportForm.GRepFormName = 6
        
End Select
With ReportForm
    .CAPTION = MnVehReports1(Index).CAPTION
    .LblTitle = MnVehReports1(Index).CAPTION
    .Show
End With
Set ReportForm = Nothing
Exit Sub
errhand:
    MsgBox err.Description, vbInformation, "Information": Exit Sub

End Sub

Private Sub SBAR_PanelClick(ByVal Panel As MSComctlLib.Panel)
If Panel.Index = 7 Then
    If StrCmp(PubLoginModule, "Spare") Or StrCmp(PubLoginModule, "Workshop") Or pubUName = "SA" Then
        If FormChk("Find Part") = False Then FrmFindPart.Show: Set FrmFindPart = Nothing
    End If
ElseIf Panel.Index = 4 Then
    If FrmUserPass.Visible = False Then FrmUserPass.Show: Set FrmUserPass = Nothing
ElseIf Panel.Index = 6 Then
    ConnectDb
End If

End Sub

Private Sub SprMas_Click(Index As Integer)
PubUParam = Permission(Replace(SprMas(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Select Case Index
    Case 1      'Aggregate Master
        If frmAggregate.Visible = False Then frmAggregate.Show: Set frmAggregate = Nothing
    Case 2      'Discount Factor Master
        If frmPartDiscFact.Visible = False Then frmPartDiscFact.Show: Set frmPartDiscFact = Nothing
    Case 3      'Proprietary Part Grade Master
        If frmPartGrade.Visible = False Then frmPartGrade.Show: Set frmPartGrade = Nothing
    Case 4      'Part Master
        
        If frmPartMast.Visible = False Then frmPartMast.Show: Set frmPartMast = Nothing
    Case 5      'Spare Price List
        If frmSprPriceList.Visible = False Then frmSprPriceList.Show: Set frmSprPriceList = Nothing
    Case 6      'Price List Updation
        If FrmPriceUpdate.Visible = False Then FrmPriceUpdate.Show: Set FrmPriceUpdate = Nothing
    Case 7      'Spare Sale Target
        If frmSprSaleTarget.Visible = False Then frmSprSaleTarget.Show: Set frmSprSaleTarget = Nothing
    Case 8      'Spare Physical Stock Entry
        'If frmPhysicalStk.Visible = False Then frmPhysicalStk.Show: Set frmPhysicalStk = Nothing
        If FrmUpdation.Visible = False Then FrmUpdation.Show
End Select
End Sub

Private Sub SprPurTrn_Click(Index As Integer)
PubUParam = Permission(Replace(SprPurTrn(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Select Case Index
    Case 1
        If frmPurOrd.Visible = False Then frmPurOrd.Show: Set frmPurOrd = Nothing
    Case 2
        If frmPurChl.Visible = False Then frmPurChl.Show: Set frmPurChl = Nothing
    Case 3
        If frmPurBill.Visible = False Then frmPurBill.Show: Set frmPurBill = Nothing
    Case 4
        If frmPurRet.Visible = False Then frmPurRet.Show: Set frmPurRet = Nothing
        'Menu Separator
End Select
End Sub

Private Sub SprSalTrn_Click(Index As Integer)
PubUParam = Permission(Replace(SprSalTrn(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Select Case Index
    Case 1
       
        If frmEstimateQuot.Visible = False Then
            frmEstimateQuot.PubEstimateType = "Spare"
            frmEstimateQuot.Show
            Set frmEstimateQuot = Nothing
        End If
    Case 2
        If frmSOrd.Visible = False Then frmSOrd.Show: Set frmSOrd = Nothing
    Case 3
        If frmSaleChal.Visible = False Then frmSaleChal.Show: Set frmSaleChal = Nothing
    Case 4
'        If RSOJPR = True And pubUName <> "ADMIN" Then
'            PubUParam = "A**P"
'        End If
        If frmSaleBill.Visible = False Then frmSaleBill.Show: Set frmSaleBill = Nothing
    Case 5
'        If RSOJPR = True And pubUName <> "ADMIN" Then
'            PubUParam = "A**P"
'        End If
        If frmSaleRet.Visible = False Then frmSaleRet.Show: Set frmSaleRet = Nothing
    Case 6
       
        If FormChk(SprSalTrn(Index).CAPTION) = True Then Exit Sub
        Dim Form1 As frmRequisition
        Set Form1 = New frmRequisition
        With Form1
            .PubRequisitionType = "Store"
            .CAPTION = SprSalTrn(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 7
        
        If FormChk(SprSalTrn(Index).CAPTION) = True Then Exit Sub
        Dim Form2 As frmRequisition
        Set Form2 = New frmRequisition
        With Form2
            .PubRequisitionType = "Return"
            .CAPTION = SprSalTrn(Index).CAPTION
            .Show
        End With
        Set Form2 = Nothing
    Case 8
'        If RSOJPR = True And pubUName <> "ADMIN" Then
'            PubUParam = "A**P"
'        End If
        If frmStkIssRec.Visible = False Then frmStkIssRec.Show: Set frmStkIssRec = Nothing
    Case 9
'        If frmCustRect.Visible = False Then frmCustRect.Show: Set frmCustRect = Nothing
        
        If frmCustRect.Visible = False Then
            frmCustRect.PubReceiptType = "Spare"
            frmCustRect.Show
            Set frmCustRect = Nothing
        End If
End Select
End Sub

Private Sub SprStkRep_Click(Index As Integer)
Dim rep As CrystalReport, Menucaption$, ReportForm As Form, mQry$, RepTitle$
Dim RepPrint As Boolean
On Error GoTo errhand
Dim RstRep As ADODB.Recordset

PubUParam = Permission(Replace(SprStkRep(Index).CAPTION, "&", ""))
Menucaption = UCase(Replace(Replace(SprStkRep(Index).CAPTION, "&", ""), " ", ""))

If Index = 9 Then   'Stock Verification Sheet
    mQry = "Select Part_name,part_no,bin_loca" & _
           " FROM Part where div_code ='" & PubDivCode & "' order by Bin_Loca,Part_no"
  
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    
    RepTitle = UCase("Stock Verification Sheet")
    
    CreateFieldDefFile RstRep, PubRepoPath & "\SprStkVerify.ttx", True
    Set rpt = rdApp.OpenReport(PubRepoPath & "\SprStkVerify.RPT")

    rpt.Database.SetDataSource RstRep
    rpt.ReadRecords

    Call Report_View(rpt, RepTitle, , False)
Else
    Set ReportForm = New ReprtFormGlobal
    Select Case Menucaption
        Case UCase(Replace("Stock Ledger", " ", ""))
            ReportForm.GRepFormName = 10
        Case UCase(Replace("Stock Summary", " ", ""))
            ReportForm.GRepFormName = 11
        Case UCase(Replace("Stock In Hand", " ", ""))
            ReportForm.GRepFormName = 12
        Case UCase(Replace("Indent Register", " ", ""))
            ReportForm.GRepFormName = 15 '22
        Case UCase(Replace("Stock Above/Below/ReOrder Lavel", " ", ""))
            ReportForm.GRepFormName = 22 '30
        Case UCase(Replace("Bin-wise Spare Stock", " ", ""))
            ReportForm.GRepFormName = 23 '31
        Case UCase(Replace("Part Movement Register", " ", ""))
            ReportForm.GRepFormName = 24 '32
        Case UCase(Replace("Part Ageing Analysis", " ", ""))
            ReportForm.GRepFormName = 25 '33
        Case UCase(Replace("Stock Value", " ", ""))
            ReportForm.GRepFormName = 45
            
    End Select
    With ReportForm
        .CAPTION = SprStkRep(Index).CAPTION
        .LblTitle = SprStkRep(Index).CAPTION
        .Show
    End With
End If
errhand:
    If err.NUMBER <> 0 Then CheckError
Set RstRep = Nothing
Set rpt = Nothing
Set ReportForm = Nothing
End Sub



'Private Sub TelRep_Click(Index As Integer)
'
'PubUParam = Permission(Replace(TelRep(Index).CAPTION, "&", ""))
'Dim rep As CrystalReport, Menucaption$
'Menucaption = UCase(Replace(Replace(TelRep(Index).CAPTION, "&", ""), " ", ""))
'
'If Menucaption = UCase(Replace("Daily Sale Report", " ", "")) Or _
'    Menucaption = UCase(Replace("Financer Certificate", " ", "")) Then
'    Dim ReportForm As Form
'End If
'
'If Menucaption = UCase(Replace("Daily Sale Report", " ", "")) Then
'    Set ReportForm = New ReportVehicle
'ElseIf Menucaption = UCase(Replace("Financer Certificate", " ", "")) Then
'    Set ReportForm = New ReportVehicle2
'End If
'
'On Error GoTo errhand
'Select Case Menucaption
'    Case UCase(Replace("Offtake Incentive Claim Letter", " ", ""))
'        If RepIncClaim.Visible = False Then RepIncClaim.Show: Set RepIncClaim = Nothing
'    Case UCase(Replace("Subvention Letter", " ", ""))
'        If RepSubvention.Visible = False Then RepSubvention.Show: Set RepSubvention = Nothing
'    Case UCase(Replace("Financer Certificate", " ", ""))
'        ReportForm.GRepFormName = 1
'    Case UCase(Replace("Daily Sale Report", " ", ""))
'        ReportForm.GRepFormName = 17
'End Select
'
'If Menucaption = UCase(Replace("Financer Certificate", " ", "")) Or _
'    Menucaption = UCase(Replace("Daily Sale Report", " ", "")) Then
'    With ReportForm
'        .CAPTION = TelRep(Index).CAPTION
'        .LblTitle = TelRep(Index).CAPTION
'        .Show
'    End With
'    Set ReportForm = Nothing
'End If
'Exit Sub
'errhand:
'    MsgBox err.Description, vbInformation, "Information": Exit Sub
'
'End Sub

Private Sub TelRep_Click(Index As Integer)

PubUParam = Permission(Replace(TelRep(Index).CAPTION, "&", ""))
Dim rep As CrystalReport, Menucaption$
Menucaption = UCase(Replace(Replace(TelRep(Index).CAPTION, "&", ""), " ", ""))

If Menucaption = UCase(Replace("Daily Sale Report", " ", "")) Or _
    Menucaption = UCase(Replace("Financer Certificate", " ", "")) Or _
    Menucaption = UCase(Replace("Subvention Claim Register", " ", "")) Then
    Dim ReportForm As Form
End If

If Menucaption = UCase(Replace("Daily Sale Report", " ", "")) Or _
    Menucaption = UCase(Replace("Subvention Claim Register", " ", "")) Then
        Set ReportForm = New ReportVehicle
ElseIf Menucaption = UCase(Replace("Financer Certificate", " ", "")) Then
    Set ReportForm = New ReportVehicle2
End If

On Error GoTo errhand
Select Case Menucaption
    Case UCase(Replace("Offtake Incentive Claim Letter", " ", ""))
        If RepIncClaim.Visible = False Then RepIncClaim.Show: Set RepIncClaim = Nothing
    Case UCase(Replace("Subvention Letter", " ", ""))
        If RepSubvention.Visible = False Then RepSubvention.Show: Set RepSubvention = Nothing
    Case UCase(Replace("Financer Certificate", " ", ""))
        ReportForm.GRepFormName = 1
    Case UCase(Replace("Daily Sale Report", " ", ""))
        ReportForm.GRepFormName = 17
    Case UCase(Replace("Subvention Claim Register", " ", ""))
        ReportForm.GRepFormName = 27
End Select

If Menucaption = UCase(Replace("Financer Certificate", " ", "")) Or _
    Menucaption = UCase(Replace("Daily Sale Report", " ", "")) Or _
    Menucaption = UCase(Replace("Subvention Claim Register", " ", "")) Then
    With ReportForm
        .CAPTION = TelRep(Index).CAPTION
        .LblTitle = TelRep(Index).CAPTION
        .Show
    End With
    Set ReportForm = Nothing
End If
Exit Sub
errhand:
    MsgBox err.Description, vbInformation, "Information": Exit Sub

End Sub

Private Sub tlcorep_Click(Index As Integer)
Dim rep As CrystalReport
PubUParam = Permission(Replace(tlcorep(Index).CAPTION, "&", ""))
Dim ReportForm As Form
On Error GoTo errhand
Select Case UCase(Replace(Replace(tlcorep(Index).CAPTION, "&", ""), " ", ""))
    Case UCase(Replace("Repeat Job Analysis Form", " ", ""))
        PubUParam = Permission(Replace(WrkTrn(Index).CAPTION, "&", ""))
        If frmRepeteJob.Visible = False Then frmRepeteJob.Show
       Exit Sub
    Case UCase(Replace("Customer Feedback Entry", " ", ""))
        PubUParam = Permission(Replace(tlcorep(Index).CAPTION, "&", ""))
        If frmFeedback.Visible = False Then frmFeedback.Show
       Exit Sub
End Select
Set ReportForm = New ReportTelco
Select Case UCase(Replace(Replace(tlcorep(Index).CAPTION, "&", ""), " ", ""))
    Case UCase(Replace("Workshop Performence Report", " ", ""))
        ReportForm.GRepFormName = 1
    Case UCase(Replace("Promise Time Deviation Report", " ", ""))
        ReportForm.GRepFormName = 2
    Case UCase(Replace("Estimated Cost Deviation", " ", ""))
        ReportForm.GRepFormName = 3
    Case UCase(Replace("ModelWise Complaint Analysis", " ", ""))
        ReportForm.GRepFormName = 4
    Case UCase(Replace("Aggregate Complaint Analysis", " ", ""))
        ReportForm.GRepFormName = 5
    Case UCase(Replace("Repaire Order Analysis", " ", ""))
        ReportForm.GRepFormName = 6
    Case UCase(Replace("ModelWise Repeat Complaint Analysis", " ", ""))
        ReportForm.GRepFormName = 7
    Case UCase(Replace("Summery of Repeat Complaint", " ", ""))
        ReportForm.GRepFormName = 10
    Case UCase(Replace("Repeat Job Analysis", " ", ""))
        ReportForm.GRepFormName = 11
    Case UCase(Replace("Service Wise Job Analysis", " ", ""))
        ReportForm.GRepFormName = 12
    Case UCase(Replace("Service Analysis Details", " ", ""))
        ReportForm.GRepFormName = 13
    Case UCase(Replace("Quantitative Complaint Analysis", " ", ""))
        ReportForm.GRepFormName = 14
    Case UCase(Replace("Daily Follow UP Calls By CRO", " ", ""))
        ReportForm.GRepFormName = 16
    Case UCase(Replace("Dissatisfied Customers", " ", ""))
        ReportForm.GRepFormName = 17
    Case UCase(Replace("Customer Satisfaction Report", " ", ""))
        ReportForm.GRepFormName = 18
    Case UCase(Replace("Workshop Profit Report", " ", ""))
        ReportForm.GRepFormName = 19
End Select
With ReportForm
    .CAPTION = tlcorep(Index).CAPTION
    .LblTitle = tlcorep(Index).CAPTION
    .Show
End With
Set ReportForm = Nothing
Exit Sub
errhand:
    MsgBox err.Description, vbInformation, "Information": Exit Sub


End Sub

Private Sub utl_Click(Index As Integer)
Dim oZip As CGZipFiles, ZipDrive As String, ZipFile$
Dim oUnZip As CGUnzipFiles
Dim DB1 As DAO.Database

ZipFile = Trim("Auto_" & PubCenCompCode & "-" & Format(Now(), "dd") & Format(Now(), "mm") & Format(Now(), "yy"))
'On Error GoTo errhand
    PubUParam = Permission(Replace(Utl(Index).CAPTION, "&", ""))
    If PubUParam = "****" Then Exit Sub
    Select Case UCase(Utl(Index).CAPTION)
        Case "USER PERMISSIONS"
            If frmUser.Visible = False Then frmUser.Show: Set frmUser = Nothing
        Case "USER GROUPS"
            If FrmUserGroup.Visible = False Then FrmUserGroup.Show: Set FrmUserGroup = Nothing
        Case "SYSTEM CONTROLS"
            If frmSyCtrl.Visible = False Then frmSyCtrl.Show: Set frmSyCtrl = Nothing
        Case "VEHICLE CONTROLS"
            If frmSyCtrlVeh.Visible = False Then frmSyCtrlVeh.Show: Set frmSyCtrlVeh = Nothing
        Case "SPARE CONTROLS"
            If frmSyCtrlSpr.Visible = False Then frmSyCtrlSpr.Show: Set frmSyCtrlSpr = Nothing
        Case "WORKSHOP CONTROLS"
            If frmSyCtrlSrv.Visible = False Then frmSyCtrlSrv.Show: Set frmSyCtrlSrv = Nothing
        Case "VOUCHER PREFIX GENERATION"
            If frmVouPrefix.Visible = False Then frmVouPrefix.Show: Set frmVouPrefix = Nothing
        Case "DATABASE  BACKUP"
            If PubBackEnd = "A" Then
                Call DataBackup
            Else
                Call BackupSqlDatabase
            End If
        Case "DATABASE RESTORE"
            '****
            'Dim DB1 As DAO.Database
            'Checking for Exclusive mode
            If GCn.State <> 0 Then GCn.Close
            Set DB1 = OpenDatabase(Pub_DataPath & "\Auto_" & PubCenCompCode & "\Automan.mdb", True, False, ";pwd=dtman")
            Set DB1 = Nothing
            GCn.Open
            '****
            Set oUnZip = New CGUnzipFiles
            With oUnZip
                ' What Zip File ?
                .ZipFileName = PubBkpPath & "\" & ZipFile & ".ZIP"
                ' Where are we zipping to ?
'                .ExtractDir = Pub_DataPath & "\Auto_" & PubCenCompCode & "\"
                .ExtractDir = mID(Pub_DataPath, 1, 3)
                ' Keep Directory Structure of Zip ? changed to true
                .HonorDirectories = True
                ' Unzip and Display any errors as required
                'All Active Connection Closed
    '            Set GCn = Nothing
                CloseAllConn
                If .Unzip <> 0 Then
                    MsgBox .GetLastMessage
                    Exit Sub
                Else
                    MsgBox PubBkpPath & "\" & ZipFile & ".ZIP Extracted Successfully to " & Pub_DataPath & "\Auto_" & PubCenCompCode & "\"
                End If
            End With
            OpenAllConn
            Set oUnZip = Nothing
            'After Successfully Completed
            'Please Re Open Your All Connection
        Case "REPAIR/COMPACT DATABASE" 'Database Compact
            CompactDatabase
        Case "UPDATE TABLE STRUCTURE"
            AddNewFieldFAData
'        Case 17
'            UpdateNarrationProc
'        Case
'            Call DelAllTransactions
        Case "FA ENVIRONMENT SETTING"
            Set NewObj = New FaEnvron
            NewObj.Show
        Case "FA VOUCHER TYPE SETTING"
            If pubUName = "SA" Then
                PubUParam = "AE*P"
            End If
            Set NewObj = New FaVtype
            NewObj.Show
        Case "YEAR END PROCESS"
            Set NewObj = New YearEnd
            NewObj.Show
        Case "UPDATE BALANCES"
            Set NewObj = FrmUpdateOpeningBalances
            NewObj.Show vbModal
        Case "UPDATE VRATE"
            Call Update_VRate
        Case "A/C MERGING"
            If PubBackEnd = "A" Then Call DataBackup
            frmAcMerze.Show
        Case "DELETE LOG"
            On Error GoTo errhand
            Dim rep As CrystalReport, ReportForm As Form
            On Error GoTo errhand
            Set ReportForm = New RepSprMIS
            ReportForm.GRepFormName = 9
            ReportForm.CAPTION = "Edit/Delete Log"
            Set ReportForm = Nothing
            Set rep = Nothing
        Case "UPDATE VOUCHER COUNTER"
            If MsgBox("Sure to Update Voucher Prefix Counters", vbYesNo) = vbYes Then
                If PubBackEnd = "A" Then Call DataBackup
                SetMaxId_VoucherPrefix
            End If
        Case "ACCESS TABLES TO SQL SERVER"
            If PubBackEnd = "S" Then
                If FrmConvertTable.Visible = False Then FrmConvertTable.Visible = True
            Else
                MsgBox "Can't Run On BackEnd Other Than SqlServer"
            End If
    End Select
Exit Sub
errhand:
    MsgBox err.Description
    If GCn.State = 0 Then GCn.Open
End Sub
Private Sub VehMas_Click(Index As Integer)
PubUParam = Permission(Replace(VehMas(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Select Case Index
    Case 1      'Rep-wise/Model-wise Target
        If frmRepModelTgt.Visible = False Then frmRepModelTgt.Show: Set frmRepModelTgt = Nothing
    Case 2     'Financier Master
        If frmFinMast.Visible = False Then frmFinMast.Show: Set frmFinMast = Nothing
    Case 3     'Addition / Deletion
        If frmVehAMDMast.Visible = False Then frmVehAMDMast.Show: Set frmVehAMDMast = Nothing
End Select
End Sub

Private Sub VehMktRep_Click(Index As Integer)
Dim rep As CrystalReport, Menucaption$, ReportForm As Form ', RepTitle$
Dim RepPrint As Boolean
On Error GoTo errhand
Dim RstRep As ADODB.Recordset

PubUParam = Permission(Replace(VehMktRep(Index).CAPTION, "&", ""))
Menucaption = UCase(Replace(Replace(VehMktRep(Index).CAPTION, "&", ""), " ", ""))
If Menucaption = UCase(Replace("Statement of Prospective Demands (SPADE)", " ", "")) Then
    Set ReportForm = New RepSPADE
Else
    Set ReportForm = New RepMkt
End If
    Select Case Menucaption
        Case UCase(Replace("Daily Activity Report", " ", ""))
            ReportForm.GRepFormName = 1
        Case UCase(Replace("Appointments", " ", ""))
            ReportForm.GRepFormName = 2
        Case UCase(Replace("Call Status Report", " ", ""))
            ReportForm.GRepFormName = 3
        Case UCase(Replace("Case Analysis", " ", ""))
            ReportForm.GRepFormName = 4
        Case UCase(Replace("Daily Activity Missing Report", " ", ""))
            ReportForm.GRepFormName = 5
        Case UCase(Replace("Appointments Not Kept", " ", ""))
            ReportForm.GRepFormName = 6
        Case UCase(Replace("Executive-wise Got/Lost Report", " ", ""))
            ReportForm.GRepFormName = 7
        Case UCase(Replace("Pipeline Report (Daily)", " ", ""))
            ReportForm.GRepFormName = 10
        Case UCase(Replace("Profession/Purpose Analysis", " ", ""))
            ReportForm.GRepFormName = 11
        Case UCase(Replace("Daily Sales Report", " ", ""))
            ReportForm.GRepFormName = 12
        Case UCase(Replace("Sales Tracking Report", " ", ""))
            ReportForm.GRepFormName = 13
        Case UCase(Replace("Finance Tracking Report", " ", ""))
            ReportForm.GRepFormName = 14
        Case UCase(Replace("Statement of Prospective Demands (SPADE)", " ", ""))
            ReportForm.GRepFormName = 1
    End Select
    With ReportForm
        .CAPTION = VehMktRep(Index).CAPTION
        .LblTitle = VehMktRep(Index).CAPTION
        .Show
    End With
errhand:
    If err.NUMBER <> 0 Then CheckError
Set RstRep = Nothing
Set rpt = Nothing
Set ReportForm = Nothing

End Sub

Private Sub VehMktTrn_Click(Index As Integer)
PubUParam = Permission(Replace(VehMktTrn(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Select Case Index
    Case 0      'Prospective Customer Master
        If frmProCust.Visible = False Then frmProCust.Show: Set frmProCust = Nothing
    Case 1     'Sales Visit / Daily Activity Entry
        If frmVisit.Visible = False Then frmVisit.Show: Set frmVisit = Nothing
    Case 2
        If frmGotLost.Visible = False Then frmGotLost.Show: Set frmGotLost = Nothing
End Select

End Sub

Private Sub VehPurTrn_Click(Index As Integer)
Dim frm1 As frmVehPur
PubUParam = Permission(Replace(VehPurTrn(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Select Case Index
    Case 1     'Vehicle Receipts Entry
        If frmVehRect.Visible = False Then frmVehRect.Show: Set frmVehRect = Nothing
    Case 2     'Vehicle Purchase Entry
        If FormChk(VehPurTrn(Index).CAPTION) = True Then Exit Sub
        Set frm1 = New frmVehPur
        With frm1
            .mVType = "V_PB"
            .CAPTION = VehPurTrn(Index).CAPTION
            .Show
            .ZOrder 0
        End With
        Set frm1 = Nothing
    Case 3  'Opening Stock
        If FormChk(VehPurTrn(Index).CAPTION) = True Then Exit Sub
        Set frm1 = New frmVehPur
        With frm1
             .mVType = "V_OST"
             .CAPTION = VehPurTrn(Index).CAPTION
             .Show
             .ZOrder 0
         End With
         Set frm1 = Nothing
    Case 4     'Vehicle Check List
        If frmVehCheckSheet.Visible = False Then frmVehCheckSheet.Show: Set frmVehCheckSheet = Nothing
         
    Case 5  'BMS Opening
        If FormChk(VehPurTrn(Index).CAPTION) = True Then Exit Sub
        Set frm1 = New frmVehPur
        With frm1
             .mVType = "V_OBD"
             .CAPTION = VehPurTrn(Index).CAPTION
             .Show
             .ZOrder 0
         End With
         Set frm1 = Nothing
    Case 6
        If FrmVehStkTrn.Visible = False Then FrmVehStkTrn.Show: Set FrmVehStkTrn = Nothing
    Case 7     'Offtake Entry
        If frmOffTakeTgtForeCast.Visible = False Then frmOffTakeTgtForeCast.Show: Set frmOffTakeTgtForeCast = Nothing
    Case 8     'Allocation Entry
        ShowForm frmVehAlloc
        If frmVehAlloc.Visible = False Then frmVehAlloc.Show:  Set frmVehAlloc = Nothing
    Case 9     'Rate List Entry
        ShowForm frmVehRate
    Case 10     'Rate List Entry
        ShowForm frmVehIssue
    Case 11
        ShowForm FrmSubventionMast
    Case 12
        ShowForm FrmOffTakeMaster
End Select
End Sub

'Private Sub VehRep_Click(Index As Integer)
'If Index = 1 Then 'SPADE
'    PubUParam = Permission(Replace(VehRep(Index).CAPTION, "&", ""))
'    Dim rep As CrystalReport, ReportForm As Form
'    Set ReportForm = New RepSPADE
'    Select Case UCase(Replace(Replace(VehRep(Index).CAPTION, "&", ""), " ", ""))
'        Case UCase(Replace("Statement of Prospective Demands (SPADE)", " ", ""))
'            ReportForm.GRepFormName = 1
'    End Select
'    With ReportForm
'        .CAPTION = VehRep(Index).CAPTION
'        .Show
'    End With
'    Set ReportForm = Nothing
'End If
'Exit Sub
'errhand:
'    MsgBox err.Description, vbInformation, "Information": Exit Sub
'End Sub

Private Sub VehSalTrn_Click(Index As Integer)
PubUParam = Permission(Replace(VehSalTrn(Index).CAPTION, "&", ""))

If PubUParam = "****" Then Exit Sub
Select Case Index
    Case 1     'Performa Vehicle
        If frmQuot.Visible = False Then frmQuot.Show: Set frmQuot = Nothing
    Case 2      'Vehicle Booking
        If frmVehBook.Visible = False Then frmVehBook.Show: Set frmVehBook = Nothing
    Case 3      'Customer Receipt
'        If frmCustRect.Visible = False Then frmCustRect.Show: Set frmCustRect = Nothing
        If frmCustRect.Visible = False Then
            frmCustRect.PubReceiptType = "Vehicle"
            frmCustRect.Show
            Set frmCustRect = Nothing
        End If
    Case 4      'Vehicle Sale Bill
        If frmVehSale.Visible = False Then frmVehSale.Show: Set frmVehSale = Nothing
    Case 5      'Vehicle Delivery Challan
        If frmVehDel.Visible = False Then frmVehDel.Show: Set frmVehDel = Nothing
        
End Select
End Sub

Private Sub Warr_Click(Index As Integer)
PubUParam = Permission(Replace(Warr(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Select Case Index
    Case 1     'Product Complaint
        If frmWarrantyPCR.Visible = False Then frmWarrantyPCR.Show: Set frmWarrantyPCR = Nothing
    Case 2     'Warranty claim data
        If FormChk("Warranty Claim Data Entry") = True Then Exit Sub
        Dim Form1 As frmWarrantyWCD
        Set Form1 = New frmWarrantyWCD
        With Form1
'            .WarrTrnType = "WE"
            .CAPTION = "Warranty Claim Data Entry"
            .Show
        End With
        Set Form1 = Nothing
    Case 3  'claim Verification
        If FormChk("Warranty Claim Verification Entry") = True Then Exit Sub
        Dim Form2 As frmWarrantyWCD
        Set Form2 = New frmWarrantyWCD
        With Form2
'            .WarrTrnType = "ZE"
            .CAPTION = "Warranty Claim Verification Entry"
            .Show
        End With
        Set Form2 = Nothing
    Case 4     'Claim dispatch
        If frmWarrantyDispatch.Visible = False Then frmWarrantyDispatch.Show: Set frmWarrantyDispatch = Nothing
    Case 5  'Warranty Billing
        If frmWarrantyBill.Visible = False Then frmWarrantyBill.Show: Set frmWarrantyBill = Nothing
    Case 6  'Warranty Cr Note / Rejection
        If frmWarrantyCredit.Visible = False Then frmWarrantyCredit.Show: Set frmWarrantyCredit = Nothing
    Case 7
        WarrFrmName = Index
        If frmWarrMast.Visible = False Then frmWarrMast.Show: Set frmWarrMast = Nothing
        frmWarrMast.CAPTION = "Complaint Code Master"
    Case 8
        WarrFrmName = Index
        If frmWarrMast.Visible = False Then frmWarrMast.Show: Set frmWarrMast = Nothing
        frmWarrMast.CAPTION = "Failure Code Master"
    Case 9
        WarrFrmName = Index
        If frmWarrMast.Visible = False Then frmWarrMast.Show: Set frmWarrMast = Nothing
        frmWarrMast.CAPTION = "Make Code Master"
    Case 10
        WarrFrmName = Index
        If frmWarrMast.Visible = False Then frmWarrMast.Show: Set frmWarrMast = Nothing
        frmWarrMast.CAPTION = "Job Code Master"
End Select
End Sub

Private Sub WarrRep_Click(Index As Integer)
PubUParam = Permission(Replace(WarrRep(Index).CAPTION, "&", ""))
Dim rep As CrystalReport, ReportForm As Form
Dim Form1 As RP_WarClmReg
Dim Form2 As RP_WarClmNotMade

    Select Case Index
        Case 0  'UCase(Replace(Replace(VehRep(Index).Caption, "&", ""), " ", ""))
            Set ReportForm = New ReportWarranty
            ReportForm.GRepFormName = 1
            With ReportForm
                .CAPTION = WarrRep(Index).CAPTION   'Warranty Claim Register
                .Show
            End With
            Set ReportForm = Nothing
        Case 1
            If FormChk("Warranty Claims not made") = True Then Exit Sub
            Set Form2 = New RP_WarClmNotMade
            With Form2
                .g_FormID = 1
                .LblName.CAPTION = "Warranty Claims not made"
                .CAPTION = "Warranty Claims not made"
                .Show
            End With
            Set Form2 = Nothing
        Case 2
            If FormChk("Warranty Claims not made") = True Then Exit Sub
            Set Form1 = New RP_WarClmReg
            With Form1
                .g_FormID = 1
                .LblName.CAPTION = "Warranty Claims not Verified"
                .CAPTION = "Warranty Claims not Verified"
                .Show
            End With
            Set Form1 = Nothing
        Case 3
            Set Form1 = New RP_WarClmReg
            With Form1
                .g_FormID = 2
                .LblName.CAPTION = "Warranty Claims not Dispatched"
                .CAPTION = "Warranty Claims not Dispatched"
                .Show
            End With
            Set Form1 = Nothing
        Case 4
            Set Form1 = New RP_WarClmReg
            With Form1
                .g_FormID = 3
                .LblName.CAPTION = "Warranty Claims Rejected"
                .CAPTION = "Warranty Claims Rejected"
                .Show
            End With
            Set Form1 = Nothing
        Case 5
            Set Form1 = New RP_WarClmReg
            With Form1
                .g_FormID = 4
                .LblName.CAPTION = "Warranty Claims Outstanding"
                .CAPTION = "Warranty Claims Outstanding"
                .Show
            End With
            Set Form1 = Nothing
        Case 6
            Set Form1 = New RP_WarClmReg
            With Form1
                .g_FormID = 5
                .LblName.CAPTION = "Overall Warranty Claims"
                .CAPTION = "Overall Warranty Claims"
                .Show
            End With
            Set Form1 = Nothing
        Case 7
            Set Form1 = New RP_WarClmReg
            With Form1
                .g_FormID = 6
                .LblName.CAPTION = "Parts Not Issued but Claimed"
                .CAPTION = "Parts Not Issued but Claimed"
                .Show
            End With
            Set Form1 = Nothing
        Case 8
            Set Form1 = New RP_WarClmReg
            With Form1
                .g_FormID = 7
                .LblName.CAPTION = "Parts Issued but Not Claimed"
                .CAPTION = "Parts Issued but Not Claimed"
                .Show
            End With
            Set Form1 = Nothing
        Case 9
            With RP_WarClmSumm
                .g_FormID = 1
                .LblName.CAPTION = "Warranty Claim Summary"
                .CAPTION = "Warranty Claim Summary"
                .Show
            End With
        Case 10
            Set ReportForm = New ReportFreeSrv
            With ReportForm
                .GRepFormName = 1
                .CAPTION = Replace(WarrRep(Index).CAPTION, "&", "")
                .LblTitle = Replace(WarrRep(Index).CAPTION, "&", "")
                .Show
            End With
            Set ReportForm = Nothing
        Case 11
            With RP_WarClmPartFail
                .g_FormID = 1
                .LblName.CAPTION = "Warranty Parts Failure Summary"
                .CAPTION = "Warranty Parts Failure Summary"
                .Show
            End With
        Case 12
            Set ReportForm = New ReportFreeSrv
            With ReportForm
                .GRepFormName = 2
                .CAPTION = Replace(WarrRep(Index).CAPTION, "&", "")
                .LblTitle = Replace(WarrRep(Index).CAPTION, "&", "")
                .Show
            End With
            Set ReportForm = Nothing
    End Select
End Sub

Private Sub WkShop_Click(Index As Integer)
Dim rep As CrystalReport
PubUParam = Permission(Replace(WkShop(Index).CAPTION, "&", ""))
Dim ReportForm As Form
On Error GoTo errhand
Set ReportForm = New ReportWorkShop
Select Case UCase(Replace(Replace(WkShop(Index).CAPTION, "&", ""), " ", ""))
    Case UCase(Replace("Labour Rate Variation Reports", " ", ""))
        ReportForm.GRepFormName = 1
    Case UCase(Replace("Service-wise Job Analysis", " ", ""))
        ReportForm.GRepFormName = 3
    Case UCase(Replace("Model-wise Service Analysis", " ", ""))
        ReportForm.GRepFormName = 4
    Case UCase(Replace("WorkShop Demand Vs Supply", " ", ""))
        ReportForm.GRepFormName = 5
    Case UCase(Replace("Mechanic Earning Report", " ", ""))
        ReportForm.GRepFormName = 6
    Case UCase(Replace("Mechanic Earning Summary", " ", ""))
        ReportForm.GRepFormName = 7
    Case UCase(Replace("Dealer-wise Vehicle Attended", " ", ""))
        ReportForm.GRepFormName = 8
    Case UCase(Replace("Model-wise Job Analysis", " ", ""))
        ReportForm.GRepFormName = 9
    Case UCase(Replace("Aggregate Group-wise Inventory", " ", ""))
        ReportForm.GRepFormName = 10
    Case UCase(Replace("Job-wise Labour Analysis", " ", ""))
        ReportForm.GRepFormName = 11
    Case UCase(Replace("Dealer-wise Job Analysis", " ", ""))
        ReportForm.GRepFormName = 12
    Case UCase(Replace("Delay Reason Analysis", " ", ""))
        ReportForm.GRepFormName = 13
    Case UCase(Replace("WorkShop Rate Variation Report", " ", ""))
        ReportForm.GRepFormName = 14
    Case UCase(Replace("Cancellation Report(W/C)", " ", ""))
        ReportForm.GRepFormName = 15
    Case UCase(Replace("Labour Incentive Report", " ", ""))
        ReportForm.GRepFormName = 16
    Case UCase(Replace("Vehicle Grading Report", " ", ""))
        ReportForm.GRepFormName = 17
    Case UCase(Replace("Service Tax Register", " ", ""))
        ReportForm.GRepFormName = 18
    Case UCase(Replace("Labour Revenue Report", " ", ""))
        ReportForm.GRepFormName = 19
    Case UCase(Replace("Service Due Register", " ", ""))
        ReportForm.GRepFormName = 20
    Case UCase(Replace("Quarterly Return Register", " ", ""))
        ReportForm.GRepFormName = 21
    Case UCase(Replace("Post Service Followups", " ", ""))
        ReportForm.GRepFormName = 22
    Case UCase(Replace("Sales Tax With Serv. Tax", " ", ""))
        ReportForm.GRepFormName = 23
    Case UCase(Replace("Repeat Job Analysis", " ", ""))
        ReportForm.GRepFormName = 24
End Select
With ReportForm
    .CAPTION = WkShop(Index).CAPTION
    .LblTitle = WkShop(Index).CAPTION
    .Show
End With
Set ReportForm = Nothing
    Exit Sub
errhand:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub





Private Sub WrkDRep_Click(Index As Integer)
Dim rep As CrystalReport, Menucaption$, ReportForm As Form
PubUParam = Permission(Replace(WrkDRep(Index).CAPTION, "&", ""))
On Error GoTo errhand
Menucaption = UCase(Replace(Replace(WrkDRep(Index).CAPTION, "&", ""), " ", ""))
If Menucaption = UCase(Replace("Workshop Sale Register", " ", "")) Or Menucaption = UCase(Replace("Sale Register", " ", "")) Or Menucaption = UCase(Replace("Workshop Money Receipt Register", " ", "")) Then
    Set ReportForm = New ReprtFormGlobal
Else
    Set ReportForm = New ReportWorkShop2
End If
Select Case Menucaption
    Case UCase(Replace("Workshop Money Receipt Register", " ", ""))
        ReportForm.GRepFormName = 13
    Case UCase(Replace("Estimate Register", " ", ""))
        ReportForm.GRepFormName = 2
    Case UCase(Replace("Performa Register", " ", ""))
        ReportForm.GRepFormName = 3
    Case UCase(Replace("Workshop Sale Register", " ", ""))
        ReportForm.GRepFormName = 14
    Case UCase(Replace("Internal Requisition Register", " ", ""))
        ReportForm.GRepFormName = 4
    Case UCase(Replace("Workshop Vehicle Diary", " ", ""))
        ReportForm.GRepFormName = 5
    Case UCase(Replace("JobCard Register", " ", ""))
        ReportForm.GRepFormName = 6
    Case UCase(Replace("Sale Register", " ", ""))
        ReportForm.GRepFormName = 28
    Case UCase(Replace("Gate Pass Register", " ", ""))
        ReportForm.GRepFormName = 7
    Case UCase(Replace("Part Grade-wise Requisition Register", " ", ""))
        ReportForm.GRepFormName = 8
    Case UCase(Replace("Out-side Labour Register", " ", ""))
        ReportForm.GRepFormName = 9
    Case UCase(Replace("Over Time Register", " ", ""))
        ReportForm.GRepFormName = 10
    Case UCase(Replace("Veh History Register", " ", ""))
        ReportForm.GRepFormName = 12
    Case UCase(Replace("Workshop Vehicle Register", " ", ""))
        ReportForm.GRepFormName = 14
        
    Case UCase(Replace("Insurance Expiry Register", " ", ""))
        ReportForm.GRepFormName = 15
        
    Case UCase(Replace("Job Booking Register", " ", ""))
        ReportForm.GRepFormName = 1
End Select
With ReportForm
    .CAPTION = WrkDRep(Index).CAPTION
    .LblTitle = WrkDRep(Index).CAPTION
    .Show
End With
Set ReportForm = Nothing
errhand:
    If err.NUMBER <> 0 Then CheckError
Set ReportForm = Nothing
End Sub

Private Sub WrkMas_Click(Index As Integer)
PubUParam = Permission(Replace(WrkMas(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Select Case Index
    Case 1     'Service Master
        If frmService.Visible = False Then frmService.Show: Set frmService = Nothing
    Case 2     'Service Rate/Lubricants Qty Declaration Model-wise
        If frmSrvLubRate.Visible = False Then frmSrvLubRate.Show: Set frmSrvLubRate = Nothing
    Case 3     'Labour Group Master
        If frmLabGrp.Visible = False Then frmLabGrp.Show: Set frmLabGrp = Nothing
    Case 4     'Labour Type Master
        If frmLabTyp.Visible = False Then frmLabTyp.Show: Set frmLabTyp = Nothing
    Case 5     'Labour Type Master
        If frmLabDesc.Visible = False Then frmLabDesc.Show: Set frmLabDesc = Nothing
    Case 6     'Model-wise Labour Details Master
        If frmModLabDet.Visible = False Then frmModLabDet.Show: Set frmModLabDet = Nothing
    Case 7     'Model-wise Service-wise Pre-Defined Job Details
        If frmModSrvPreDefJob.Visible = False Then frmModSrvPreDefJob.Show: Set frmModSrvPreDefJob = Nothing
    Case 8     'Reason for Job Delay Master
        If frmJobDelay.Visible = False Then frmJobDelay.Show: Set frmJobDelay = Nothing
    Case 9     'Trouble Master
        If frmTrouble.Visible = False Then frmTrouble.Show: Set frmTrouble = Nothing
    Case 10     'Inspection Category Master
        If frmInspCat.Visible = False Then frmInspCat.Show: Set frmInspCat = Nothing
    Case 11     'Inspection Elements Master
        If frmInspEle.Visible = False Then frmInspEle.Show: Set frmInspEle = Nothing
    Case 12     'Workshop Vehicle Master
        If frmWorkVehiMast.Visible = False Then frmWorkVehiMast.Show: Set frmWorkVehiMast = Nothing
    Case 13     'Workshop Details Master
        If frmWorkDetMast.Visible = False Then frmWorkDetMast.Show: Set frmWorkDetMast = Nothing
    End Select
End Sub


Private Sub WrkTrn_Click(Index As Integer)
PubUParam = Permission(Replace(WrkTrn(Index).CAPTION, "&", ""))
If PubUParam = "****" Then Exit Sub
Dim Form1 As frmRequisition
Select Case Index
    Case 1
        If frmJobBooking.Visible = False Then frmJobBooking.Show: Set frmJobBooking = Nothing
    Case 2
        
        If frmJobCard.Visible = False Then frmJobCard.Show: Set frmJobCard = Nothing
    Case 3
        If FormChk(WrkTrn(Index).CAPTION) = True Then Exit Sub
        Set Form1 = New frmRequisition
        With Form1
            .PubRequisitionType = "Workshop"
            .CAPTION = WrkTrn(Index).CAPTION
            .Show
        End With
        Set Form1 = Nothing
    Case 4
        If frmJobGatePass.Visible = False Then frmJobGatePass.Show: Set frmJobGatePass = Nothing
    Case 5
        
        If frmJobLabour.Visible = False Then frmJobLabour.Show: Set frmJobLabour = Nothing
    Case 6
        If frmProfLab.Visible = False Then frmProfLab.Show: Set frmProfLab = Nothing
    Case 7
        If frmJobObserAction.Visible = False Then frmJobObserAction.Show: Set frmJobObserAction = Nothing
    Case 8
'        If frmEstimateQuot.Visible = False Then frmEstimateQuot.Show: Set frmEstimateQuot = Nothing
        If frmEstimateQuot.Visible = False Then
            frmEstimateQuot.PubEstimateType = "WorkShop"
            frmEstimateQuot.Show
            Set frmEstimateQuot = Nothing
        End If
    Case 9
        
        If frmJobClose.Visible = False Then frmJobClose.Show: Set frmJobClose = Nothing
    Case 10
        If frmCustRect.Visible = False Then
            frmCustRect.PubReceiptType = "WorkShop"
            frmCustRect.Show
            Set frmCustRect = Nothing
        End If
    Case 11
        If frmOverTime.Visible = False Then frmOverTime.Show: Set frmOverTime = Nothing
End Select
End Sub

Public Function Permission(Menucaption As String) As String
On Error GoTo err1
Dim TSQL As String
Dim Rs As ADODB.Recordset
'Form_Code + Param_Str + Comp_Code + Div_Code

    TSQL = "select Param_Str from user2 where Comp_Code='" & PubCenCompCode & _
            "' and Div_code = '" & PubDivCode & "' and user_name='" & pubUName & _
            "' and form_code in (select Form_Code from User_Module where " & cUCase("Name") & "='" & UCase(Menucaption) & "')"

    Set Rs = New ADODB.Recordset
    Rs.Open TSQL, G_CompCn, adOpenStatic, adLockReadOnly
    If Not Rs.EOF Then
        Permission = Rs!param_str
'    Else
'        Permission = ""
'        MsgBox "UnAuthorised Access", vbInformation, "Access Denied"
    End If
    If RSOJPR = True Then
        If pubUName = "ADMIN" Then Permission = "AEDP"
        If PubLockFinancialYear Then Permission = Replace(Replace(Replace(Permission, "A", ""), "E", ""), "D", ""): Exit Function
    Else
        If pubUName = "SA" Then Permission = "AEDP"
        If PubLockFinancialYear Then Permission = Replace(Replace(Replace(Permission, "A", ""), "E", ""), "D", ""): Exit Function
    End If

    
    
    Set Rs = Nothing
Exit Function
err1:
    MsgBox err.Description
End Function

Private Sub ModuleMenu()
Dim I As Integer
Dim Ctrl As Control
Dim RsTemp As ADODB.Recordset


'On Error GoTo ELoop
'    If pubUName <> "SA" Then
'        For Each Ctrl In Me.Controls
'            If TypeOf Ctrl Is Menu Then
'                Debug.Print Ctrl.CAPTION
'                If Ctrl.CAPTION = "-" Or UCase(Ctrl.Name) = "MNUWINDOW" Or UCase(Ctrl.Name) = "MNUWINDOWCASCADE" Or UCase(Ctrl.Name) = "MNUWINDOWHORIZONTAL" Or UCase(Ctrl.Name) = "MNUWINDOWVERTICAL" Or UCase(Ctrl.Name) = "MNUEXIT" Then
'                    Ctrl.Visible = True
'                Else
'                    Ctrl.Visible = False
'                End If
'            End If
'        Next
'
'        Set RsTemp = G_CompCn.Execute("Select User2.*, UM.Name From User2 Left join User_Module UM On User2.Module_Name=UM.Module_Name And User2.Form_Code=UM.Form_Code Where User_Name='" & pubUName & "' And Comp_Code='" & PubCenCompCode & "' And Div_Code='" & PubDivCode & "' And Param_Str<>'****'")
'        If RsTemp.RecordCount > 0 Then
'            Do Until RsTemp.EOF
'                For Each Ctrl In Me.Controls
'                    If TypeOf Ctrl Is Menu Then
'                        If Ctrl.CAPTION = RsTemp!Name Then
'                            Ctrl.Visible = True
'                        End If
'                    End If
'                Next
'
'
'                RsTemp.MoveNext
'            Loop
'        End If
'    End If
'
'     If PubVCompCode = "" Then
'    If Not AllowModuleVeh Then
''        Mas(10).Visible = False
'        Me.MnuVeh.Visible = False
'        'Utl(4).Visible = False
'    End If
'   If PubVCompCode = "" And PubWCompCode = "" Then
'   If Not AllowModuleVeh And Not AllowModuleWrk Then
'        Me.Mas(3).Visible = False
'        Me.Mas(7).Visible = False
'        Me.Mas(8).Visible = False
'        Me.Mas(9).Visible = False
'        Me.Mas(10).Visible = False
'        For I = 10 To 16
'            Me.MnuLst1(I).Visible = False
'        Next
'    End If
'    If PubSCompCode = "" Then
'    If Not AllowModuleSpr Then
'        Me.Mas(12).Visible = False
'        Me.MnuSpr.Visible = False
'        Utl(6).Visible = False
'    End If
'    If PubWCompCode = "" Then
'    If Not AllowModuleWrk Then
'        MnuWorks.Visible = False
'        'Utl(4).Visible = False
'    End If
    
    If PubBackEnd = "S" Then
        'Utl(13).Visible = False
        'Utl(14).Visible = False
        Utl(15).Visible = True
'        Utl(24).Visible = False
        'Utl(25).Visible = False
        'DTools.Visible = False   'DataTools Hide
    End If
    
    If StrCmp(left(PubComp_Name, 4), "Enar") = False Then
        fame(10).Visible = False
        fame(11).Visible = False
        FAREPORT(37).Visible = False
        FAREPORT(38).Visible = False
        FAREPORT(39).Visible = False
        SprStkRep(10).Visible = False
    End If
    
    
    If Not StrCmp(left(PubComp_Name, 5), "Ujwal") Then
        DTool(8).Visible = False
        DTool(9).Visible = False
        FAREPORT(40).Visible = False
    End If
    
Exit Sub
ELoop:
    MsgBox err.Description
End Sub

Private Sub CloseAllConn()
If GCn.State <> 0 Then GCn.Close
If GCnFaV.State <> 0 Then GCnFaV.Close
If GCnFaS.State <> 0 Then GCnFaS.Close
If GCnFaW.State <> 0 Then GCnFaW.Close
G_FaCn.Close
'G_CompCn.Close
End Sub
Private Sub OpenAllConn()
With GCn
    .CursorLocation = adUseClient
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=" & Pub_DataPath & "\Auto_" & PubCenCompCode & "\automan.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
    .Open
End With
With GCnFaV
    .CursorLocation = adUseClient
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=" & PubVFADataPath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
    .Open
End With
With GCnFaS
    .CursorLocation = adUseClient
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=" & PubSFADataPath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
    .Open
End With
With GCnFaW
    .CursorLocation = adUseClient
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=" & PubWFADataPath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
    .Open
End With
With G_FaCn
    .CursorLocation = adUseClient
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=" & PubVFADataPath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
    .Open
End With
End Sub

Public Sub CompactDatabase()
'   1.Check User Status must be Administrator
'   2.Necessary Disk Space
'   3.Tag G_FACN connect string
'   4.Close All Connections
'   5.Check for exclusive mode / make sure all users log out
If PubULabel <> "Y" Then MsgBox "Permission Denied", vbCritical, "Permission Not Found !"
MsgBox "At least twice Disk space of exisiting database size required", vbCritical, "Extra Disk Space Required!"
If MsgBox("Are You Sure To Compact Database ? ", vbYesNo + vbCritical + vbDefaultButton2, "Compact Database !") = vbNo Then Exit Sub

On Error GoTo ELoop
'declare array for multiple databases
Dim DataBases(4) As DBDetails, I As Byte
Dim DB As DAO.Database
Dim NewDataPath$, DatabaseSep As Byte
DataBases(0).DataPath = Pub_DataPath & "\Auto_" & PubCenCompCode & "\Automan.mdb"
DataBases(0).PassWord = "dtman"
If PubVFADataPath = PubSFADataPath And PubVFADataPath = PubWFADataPath Then
    DataBases(1).DataPath = PubVFADataPath
    DataBases(1).PassWord = ""
ElseIf PubVFADataPath = PubSFADataPath And PubVFADataPath <> PubWFADataPath Then
    DataBases(1).DataPath = PubVFADataPath
    DataBases(1).PassWord = ""
    DataBases(2).DataPath = PubWFADataPath
    DataBases(2).PassWord = ""
ElseIf PubVFADataPath <> PubSFADataPath And PubVFADataPath = PubWFADataPath Then
    DataBases(1).DataPath = PubVFADataPath
    DataBases(1).PassWord = ""
    DataBases(2).DataPath = PubSFADataPath
    DataBases(2).PassWord = ""
ElseIf PubVFADataPath <> PubSFADataPath And PubVFADataPath <> PubWFADataPath Then
    DataBases(1).DataPath = PubVFADataPath
    DataBases(1).PassWord = ""
    DataBases(2).DataPath = PubSFADataPath
    DataBases(2).PassWord = ""
    DataBases(3).DataPath = PubWFADataPath
    DataBases(3).PassWord = ""
End If
GCn.Close
GCnFaV.Close
GCnFaS.Close
GCnFaW.Close
G_FaCn.Close
    For I = 0 To UBound(DataBases)
        If DataBases(I).DataPath <> "" Then
            '*******
            Picture1.Visible = True
            Label1.CAPTION = "Please wait compacting (" & DataBases(I).DataPath & ") ..."
            '********

            DatabaseSep = InStrRev(DataBases(I).DataPath, "\")
            NewDataPath = left(DataBases(I).DataPath, DatabaseSep) & "DB1.MDB"
            'Check disk space
            
            'Checking for Exclusive mode
'            If DataBases(i).PassWord = "" Then
'                Set DB = OpenDatabase(DataBases(i).DataPath, True)
'            Else
                Set DB = OpenDatabase(DataBases(I).DataPath, True, False, ";pwd=" & DataBases(I).PassWord)
'            End If
            DB.Close
            
            ' Make sure there isn't already a file with the
            ' name of the compacted database.
            If Dir(NewDataPath) <> "" Then _
                Kill NewDataPath
            
            'This statement creates a compact Microsoft Jet database.
            If DataBases(I).PassWord = "" Then
                DBEngine.CompactDatabase DataBases(I).DataPath, NewDataPath
            Else
                DBEngine.CompactDatabase DataBases(I).DataPath, NewDataPath, , , ";pwd=" & DataBases(I).PassWord
            End If
            Kill DataBases(I).DataPath
            Name NewDataPath As DataBases(I).DataPath   ' Rename file.
        End If
    Next
    MsgBox "Database Compact Completed successfully!", vbOKOnly, "Compact Database"
    
ELoop:
Picture1.Visible = False
Set DB = Nothing
GCn.Open
GCnFaV.Open
GCnFaS.Open
GCnFaW.Open
G_FaCn.Open
' any errors
If err.NUMBER <> 0 Then CheckError
End Sub

Public Sub AddFieldTable(DB As DAO.Database, TableName As String, FieldName As String, _
        FIELDTYPE As Variant, Optional FieldSize As Integer, Optional RequiredYesNo As Boolean, _
        Optional AllowZero As Boolean, Optional DefValue As Variant)
'TableName$, FieldName$, FIELDTYPE As Variant, Optional FieldSize As Integer,
'Optional RequiredYesNo As Boolean, Optional AllowZero As Boolean, Optional DefValue As Variant
Dim tmprs As DAO.Recordset
Dim N As Integer, TDF As TableDef, FLD As DAO.Field
    Set tmprs = DB.OpenRecordset("select * from " & TableName)
    For N = 0 To tmprs.Fields.Count - 1
        If UCase(tmprs.Fields(N).Name) = UCase(FieldName) Then
            GoTo Myexit
        End If
        'MsgBox UCase(tmprs.Fields(N).Name)
        'MsgBox tmprs.Fields(N)
    Next
    Set tmprs = Nothing
    Set TDF = DB.TableDefs(TableName)
    Set FLD = TDF.CreateField(FieldName)
    
    FLD.Type = FIELDTYPE
    
    If FIELDTYPE = 10 Then      '' Text Field
        FLD.Size = FieldSize
        If Not IsMissing(AllowZero) Then FLD.AllowZeroLength = AllowZero
    End If
    If Not IsMissing(DefValue) Then FLD.DefaultValue = DefValue
    If Not IsMissing(RequiredYesNo) Then FLD.Required = RequiredYesNo
    TDF.Fields.Append FLD
Myexit:
    Set tmprs = Nothing
End Sub




Public Sub AddNewFieldFAData()
Dim TempRs As ADODB.Recordset
Dim RsUser As ADODB.Recordset
If MsgBox("Are You Sure To Update Table Structure ? ", vbYesNo + vbCritical + vbDefaultButton2, "Modify Database !") = vbNo Then Exit Sub
If PubBackEnd = "A" Then DataBackup
Picture1.Visible = True
Label1.CAPTION = "Please Wait ! Table Structure Updation in progress....."
Dim DataPath$, DataPathFa$
Dim DB As DAO.Database
Dim RsTemp As ADODB.Recordset
On Error Resume Next


If PubBackEnd = "A" Then
    DataPath = Pub_DataPath & "\Auto_" & PubCenCompCode & "\Automan.mdb"
    DataPathFa = PubFADataPath
        
    'dbBinInt,dbBoolean,dbByte,dbDate,dbDecimal,dbDouble,dbFloat
    'dbInteger,dbMemo,dbNumeric,dbSingle,dbText,dbTime,dbTimestamp
        'Checking for Exclusive mode
        If G_CompCn.State = 1 Then G_CompCn.Close
        If GCn.State = 1 Then GCn.Close
        If GCnFaV.State = 1 Then GCnFaV.Close
        If GCnFaS.State = 1 Then GCnFaS.Close
        If GCnFaW.State = 1 Then GCnFaW.Close
        If G_FaCn.State = 1 Then G_FaCn.Close
        
        
        
        'Create New Tables
            CreateNewTable
        
        
        
        
        
        ''''''''''''''''''''''''''''
        '''''''''''''''''''''''''
        
        
        
        
        
        Dim tmpDbComp As DAO.Database
        Set tmpDbComp = OpenDatabase(Pub_DataPath & "\Company.MDB", True, False)
        Dim tmpDb As DAO.Database
        Set tmpDb = OpenDatabase(DataPath, True, False, ";pwd=dtman")
        Dim tmpDbFA As DAO.Database
        Set tmpDbFA = OpenDatabase(DataPathFa, True, False)
        
                
        
        'From 20-09-2003
        
    Call AddFieldTable(tmpDb, "Job_Card", "Created_AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Job_Card", "Created_AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Job_Card", "Created_ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Job_Card", "Created_ModifyDate", dbDate, , False, True)
    
    Call AddFieldTable(tmpDb, "Job_Card", "Closed_AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Job_Card", "Closed_AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Job_Card", "Closed_ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Job_Card", "Closed_ModifyDate", dbDate, , False, True)
    
    Call AddFieldTable(tmpDb, "Job_Lab", "AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Job_Lab", "AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Job_Lab", "ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Job_Lab", "ModifyDate", dbDate, , False, True)
    
    Call AddFieldTable(tmpDbFA, "Ledger", "AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDbFA, "Ledger", "AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDbFA, "Ledger", "ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDbFA, "Ledger", "ModifyDate", dbDate, , False, True)
       
    Call AddFieldTable(tmpDbFA, "LedgerM", "AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDbFA, "LedgerM", "AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDbFA, "LedgerM", "ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDbFA, "LedgerM", "ModifyDate", dbDate, , False, True)
          
    Call AddFieldTable(tmpDb, "Sp_Purch", "AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Sp_Purch", "AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Sp_Purch", "ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Sp_Purch", "ModifyDate", dbDate, , False, True)
          
    Call AddFieldTable(tmpDb, "Sp_Sale", "AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Sp_Sale", "AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Sp_Sale", "ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Sp_Sale", "ModifyDate", dbDate, , False, True)
          
    Call AddFieldTable(tmpDb, "Veh_Order", "Book_AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order", "Book_AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Veh_Order", "Book_ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order", "Book_ModifyDate", dbDate, , False, True)
          
    Call AddFieldTable(tmpDb, "Veh_Order", "DelCh_AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order", "DelCh_AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Veh_Order", "DelCh_ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order", "DelCh_ModifyDate", dbDate, , False, True)
    
          
    Call AddFieldTable(tmpDb, "Veh_Order", "Inv_AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order", "Inv_AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Veh_Order", "Inv_ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order", "Inv_ModifyDate", dbDate, , False, True)
          
    Call AddFieldTable(tmpDb, "Veh_Order1", "Book_AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order1", "Book_AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Veh_Order1", "Book_ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order1", "Book_ModifyDate", dbDate, , False, True)
          
    Call AddFieldTable(tmpDb, "Veh_Order1", "DelCh_AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order1", "DelCh_AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Veh_Order1", "DelCh_ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order1", "DelCh_ModifyDate", dbDate, , False, True)
          
    Call AddFieldTable(tmpDb, "Veh_Order1", "Inv_AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order1", "Inv_AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Veh_Order1", "Inv_ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Order1", "Inv_ModifyDate", dbDate, , False, True)
          
    Call AddFieldTable(tmpDb, "Veh_Purch1", "AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Purch1", "AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Veh_Purch1", "ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Veh_Purch1", "ModifyDate", dbDate, , False, True)
    
    Call AddFieldTable(tmpDb, "Rect", "AddBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Rect", "AddDate", dbDate, , False, True)
    Call AddFieldTable(tmpDb, "Rect", "ModifyBy", dbText, 10, False, True)
    Call AddFieldTable(tmpDb, "Rect", "ModifyDate", dbDate, , False, True)
        
        
        Call AddFieldTable(tmpDb, "ColMast", "Siebel_Color", dbText, 30, False, True)
        
        Call AddFieldTable(tmpDb, "Sp_Stock", "SFCPer", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Sp_Stock", "SFCAmt", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "sp_purch", "SFCAmt", dbDouble, , False, True, 0)
        
        Call AddFieldTable(tmpDb, "TaxForms", "SFCPer", dbDouble, , False, True, 0)
        
        
        Call AddFieldTable(tmpDb, "Estimate", "SatAmt", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Estimate1", "SatPer", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Estimate1", "SatAmt", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Part_Grade", "AddTaxPer", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "TaxForms", "AddTaxPer", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "TaxFormsAc", "AddTaxAc", dbText, 8, False, True, "''")
        Call AddFieldTable(tmpDb, "Syctrl", "Sat_Yn", dbByte, , False, True, 0)
        Call AddFieldTable(tmpDb, "Sp_Sale", "Sat_Yn", dbByte, , False, True, 0)
        Call AddFieldTable(tmpDb, "Sp_Sale", "SatAmt", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Sp_Purch", "Sat_Yn", dbByte, , False, True, 0)
        Call AddFieldTable(tmpDb, "Sp_Purch", "SatAmt", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Sp_Stock", "Sat_Yn", dbByte, , False, True, 0)
        Call AddFieldTable(tmpDb, "Sp_Stock", "SatPer", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Sp_Stock", "SatAmt", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "Sat_Yn", dbByte, , False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "SatPer", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "SatAmt", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Purch1", "Sat_Yn", dbByte, , False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Purch1", "SatPer", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Purch1", "SatAmt", dbDouble, , False, True, 0)
        
        
        Call AddFieldTable(tmpDb, "Estimate1", "Item_Value", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "ServiceTaxPer_Saperate", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "HECessPer", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "ServiceTaxPer_Saperate", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "HECessPer", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "ServiceTaxAmt_Saperate", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "HECessAmt", dbDouble, , False, True, 0)
                
        Call AddFieldTable(tmpDb, "Exp_Emp1", "SubCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Stock", "BodyBuilder_BodyType", dbText, 5, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Stock", "BodyBuilder", dbText, 5, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Stock", "BodyBuilder_Remark", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Stock", "BodyBuilder_IssDate", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Stock", "BodyBuilder_RecDate", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Chas_Mth", "Code", dbText, 2, False, True, "")
        Call AddFieldTable(tmpDbComp, "User_Module", "Menu_Name", dbText, 25, False, True, "")
        
        Call AddFieldTable(tmpDbFA, "Ledger", "EmpDetailYn", dbText, 1, False, True, "")
        Call AddFieldTable(tmpDb, "SubGroup", "EmpDetailYn", dbText, 1, False, True, "")
        Call AddFieldTable(tmpDbFA, "SubGroup", "EmpDetailYn", dbText, 1, False, True, "")
        Call AddFieldTable(tmpDb, "SubGroupAlias", "EmpDetailYn", dbText, 1, False, True, "")
        Call AddFieldTable(tmpDbFA, "SubGroupAlias", "EmpDetailYn", dbText, 1, False, True, "")
        
        
        Call AddFieldTable(tmpDb, "SubGroup", "ChequeReportName", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDbFA, "SubGroup", "ChequeReportName", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDb, "SubGroupAlias", "ChequeReportName", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDbFA, "SubGroupAlias", "ChequeReportName", dbText, 50, False, True, "")
        
        
        Call AddFieldTable(tmpDb, "DmsEnviro", "SprPurchase4Ac", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "DmsEnviro", "VehPurGroupCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "DmsEnviro", "VehSaleGroupCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "DmsEnviro", "SprPurGroupCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "DmsEnviro", "SprSaleGroupCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "DmsEnviro", "VatGroupCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "DmsEnviro", "ServiceTaxGroupCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "DmsEnviro", "OtherChargesAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "DmsEnviro", "DiscountAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "DmsEnviro", "VehCstPurchaseAc", dbText, 8, False, True, "")
        Call ChangeFieldType(tmpDbComp, "User2", "Module_Name", dbText, 10, True, True)
        Call ChangeFieldType(tmpDbComp, "Company", "Exe", dbNumeric, , True, True)
        Call ChangeFieldType(tmpDb, "SubGroup", "AcCode", dbText, 8, True, True)
        Call AddFieldTable(tmpDb, "Part", "NDP", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DmsData", "Discount", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DmsData", "LabDiscount", dbDouble, 10, False, True, 0)
        
        Call AddFieldTable(tmpDb, "Estimate1", "TaxPer", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Estimate1", "TaxAmt", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DmsData", "OtherCharges", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DmsData", "LabOtherCharges", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DmsData", "TaxableAmt", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DmsData", "LabTaxableAmt", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDbFA, "LedgerM", "DmsRefNo", dbText, 40, False, True, 0)
        Call AddFieldTable(tmpDb, "DmsData", "Chassis", dbText, 18, False, True, "")
        Call AddFieldTable(tmpDb, "DmsEnviro", "SprCstPurchaseAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "DmsEnviro", "WsBankAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "Syctrl", "MakeDataBlank", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDb, "DeleteLog", "Type", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDb, "DeleteLog", "VType", dbText, 10, False, True, "")
        Call AddFieldTable(tmpDb, "DeleteLog", "VDate", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "DeleteLog", "Total_Item", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DeleteLog", "Total_Qty", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DeleteLog", "GoodsValue", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DeleteLog", "Discount", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DeleteLog", "Addition", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DeleteLog", "Deduction", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DeleteLog", "LabDiscount", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DeleteLog", "LabAmount", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "DeleteLog", "AutoYn", dbText, 1, False, True, "")
        Call AddFieldTable(tmpDb, "DeleteLog", "EditDate", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "DeleteLog", "EditTime", dbText, 20, False, True, "")
        
        

        Call AddFieldTable(tmpDb, "Site", "Address1", dbText, 40, False, True, "")
        Call AddFieldTable(tmpDb, "Site", "Address2", dbText, 40, False, True, "")
        Call AddFieldTable(tmpDb, "Site", "Address3", dbText, 40, False, True, "")
        Call AddFieldTable(tmpDb, "Site", "City", dbText, 40, False, True, "")
        Call AddFieldTable(tmpDb, "Site", "PinCode", dbText, 10, False, True, "")
        Call AddFieldTable(tmpDb, "Site", "Phone", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDb, "Site", "Mobile", dbText, 25, False, True, "")
        Call AddFieldTable(tmpDb, "Site", "LstNo", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDb, "Site", "LstDate", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Site", "CstNo", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDb, "Site", "CstDate", dbDate, 15, False, True, "")
        
        Call AddFieldTable(tmpDb, "Sp_Stock", "TempNetAmt", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Sp_Stock", "TempNetAmt2", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Lab", "ActualHrs", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Emp_Mast", "Supervisor", dbText, 5, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Order", "SubventionScheme", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Order", "HandlingCharges", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "DealerContribution", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "TataContribution", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "Subvention", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Purch1", "SubventionCredit", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "RTOFee", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "Insurance", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "DeliveryFrom", dbText, 10, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Rate", "Reg_FeeCom", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Rate", "HandlingCharges", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Rate", "GenExGodRate", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Rate", "GovtExGodRate", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDbFA, "AcControls", "SubventionClaimAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "AcControls", "SubventionAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "AcControls", "IndirectExpAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "AcControls", "OctraiAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "AcControls", "RegnFeeAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "AcControls", "InsuranceFeeAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "AcControls", "CreditCardAc", dbText, 8, False, True, "")
        
        Call AddFieldTable(tmpDbFA, "AcControls", "ChqClrAc", dbText, 8, False, True, "")
        
        Call AddFieldTable(tmpDb, "Job_Card", "FreeWarrLabAmt", dbDouble, 8, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "CreditCardNo", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDb, "Rect", "CreditCardNo", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Card", "ChqNo", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Card", "ChqDate", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Estimate", "NamePrefix", dbText, 4, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "LabAmt_Out", dbDouble, 5, False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "SprTaxInvPrefix", dbText, 5, True, True, "")
        Call AddFieldTable(tmpDb, "Syctrl", "VehTaxInvPrefix", dbText, 5, True, True, "")
        Call AddFieldTable(tmpDb, "Syctrl", "EditLock", dbDouble, 8, True, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "CheckNegetiveStockSiteWise", dbByte, 1, True, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "LabDiscAfterTaxYn", dbByte, 1, True, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "SrvTaxOnOutSideLab", dbByte, 1, True, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "RtoInsInBill", dbByte, 1, True, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "PostRegnFeeYn", dbByte, 1, True, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "PostInsuranceFeeYn", dbByte, 1, True, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "PostOctraiSaperatelyYn", dbByte, 1, True, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "VatPerOnLube", dbDouble, , True, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "eCessPer", dbDouble, , True, True, 0)
        Call AddFieldTable(tmpDb, "Part_Grade", "VatPer", dbDouble, , True, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "eCessPer", dbDouble, , True, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "eCessAmt", dbDouble, , True, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "DebtorInSupplierHelp", dbByte, 1, True, True, 0)
        Call AddFieldTable(tmpDb, "Hiscard", "DisSprMRP", dbSingle, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Hiscard", "DisSprTB", dbSingle, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Hiscard", "DisSprTP", dbSingle, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Hiscard", "DisOilMRP", dbSingle, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Hiscard", "DisOilTB", dbSingle, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Hiscard", "DisOilTP", dbSingle, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Warr2", "AggNo", dbText, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Demand", "Lab_Code", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "Model", "ServiceTax_YN", dbByte, 1, False, True, "")
        Call AddFieldTable(tmpDb, "Model", "Sale_Rate", dbDouble, 8, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "TOT_Per", dbSingle, 5, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "TOT_Amt", dbDouble, 5, False, True, 0)
        Call AddFieldTable(tmpDb, "Rect", "Veh_Amt", dbDouble, 5, False, True, 0)
        Call AddFieldTable(tmpDb, "Rect", "Tax_Amt", dbDouble, 5, False, True, 0)
        Call AddFieldTable(tmpDb, "Rect", "Surcharge_Amt", dbDouble, 5, False, True, 0)
        Call AddFieldTable(tmpDb, "Rect", "TOT_Amt", dbDouble, 5, False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "PostFinAmt", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Lab", "Chrg_Type", dbText, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Emp_Mast", "OT_Rate", dbSingle, 6, False, True, 0)
        Call AddFieldTable(tmpDb, "Estimate", "Suppl_YN", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "DiscOnLube", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "TOTOnLube", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "JobCardPrePrintedYN", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "JobBillPrePrintedYN", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "SprBillPrePrintedYN", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "VehBillPrePrintedYN", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_GatePass", "Complaints", dbText, 150, False, True, "")
        Call AddFieldTable(tmpDb, "Job_GatePass", "Instructions", dbText, 150, False, True, "")
        Call AddFieldTable(tmpDb, "SP_Purch", "Transportation", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDbFA, "AcControls", "SprPurTrans_Ac", dbText, 8, False, True, 0)
        Call AddFieldTable(tmpDb, "SubGroup", "RC_No", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDbFA, "SubGroup", "RC_No", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDb, "SubGroupAlias", "RC_No", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDbFA, "SubGroupAlias", "RC_No", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDb, "SP_Order", "CustOrd_Det", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDb, "Part", "PhyStk", dbDouble, 7, False, True, 0)
        Call AddFieldTable(tmpDb, "Trouble", "TRelated", dbText, 3, False, True, "")
        Call AddFieldTable(tmpDb, "Trouble", "TType", dbText, 9, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Demand", "Complaint_YN", dbDouble, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Demand", "Repeat_YN", dbDouble, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_SubGroupQuot", "OrdDocID", dbText, 21, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_SubGroupQuot", "PurchModel", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDb, "Visits", "Schemes", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Visits", "SalesNos", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Visits", "Prices", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Visits", "Pamphlets", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Visits", "Hoardings", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Visits", "Events", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Visits", "MediaAds", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Visits", "Misc", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "HisCard", "Varient", dbText, 10, False, True, "")
        Call AddFieldTable(tmpDb, "HisCard", "VehDet", dbText, 10, False, True, "")
        Call AddFieldTable(tmpDb, "HisCard", "ExtendWar", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "HisCard", "ExtendWar", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "HelpLineNo", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDb, "Part", "PurDocId", dbText, 21, False, True, "")
        Call AddFieldTable(tmpDb, "Part", "PurDate", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Part", "PurRate", dbDouble, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Estimate", "Remarks", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDb, "Syctrl", "SrvTaxNo", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDbFA, "AcControls", "JobContractor", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Card", "KmsHrs", dbText, 1, False, True, "")
        Call AddFieldTable(tmpDb, "Syctrl", "VAT_YN", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDbFA, "SubGroup", "NewYrSubCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "SubGroupAlias", "NewYrSubCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "SubGroup", "NewYrSubCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "SubGroupAlias", "NewYrSubCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "Syctrl", "SDT_YN", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "TOTCaption", dbText, 5, False, True, "T O T")
        Call AddFieldTable(tmpDb, "SP_Stock", "TaxPer", dbDouble, 8, False, True, 0)
        Call AddFieldTable(tmpDb, "SP_Stock", "TaxAmt", dbDouble, 8, False, True, 0)
        Call AddFieldTable(tmpDb, "Syctrl", "VehRateIncTax", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "SubTot", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "RegBy", dbText, 6, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Order1", "SubTot", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order1", "RegBy", dbText, 6, False, True, "")
        Call AddFieldTable(tmpDbFA, "SubGroup", "FinCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "SubGroupAlias", "FinCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "SubGroup", "FinCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "SubGroupAlias", "FinCode", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Stock", "TrfParty", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "Syctrl", "ZeroBill", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "HrMeter", dbText, 12, False, True, "")
        Call AddFieldTable(tmpDb, "SP_Sale", "PType", dbText, 7, False, True, "General")
        Call AddFieldTable(tmpDb, "Veh_Stock", "RectType", dbText, 1, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Order", "AdvEMI", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "AccCode", dbText, 100, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Order", "AccQty", dbText, 50, False, True, "")
        
        Call AddFieldTable(tmpDb, "Veh_Order", "OffTake", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "InsComm", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "FinPayOut", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "FinInc", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "EBTA", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "Retail", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "SPInc", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "Brokrage", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "Subvention", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order", "AmtRecd", dbDouble, 10, False, True, 0)
        
        Call AddFieldTable(tmpDb, "Veh_Order1", "AdvEMI", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order1", "AccCode", dbText, 100, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Order1", "AccQty", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Order1", "OffTake", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order1", "InsComm", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order1", "FinPayOut", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order1", "FinInc", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order1", "EBTA", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order1", "Retail", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order1", "SPInc", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order1", "Brokrage", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order1", "Subvention", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Veh_Order1", "AmtRecd", dbDouble, 10, False, True, 0)
        
        Call AddFieldTable(tmpDbFA, "Voucher_Type", "DefaultDrAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "Voucher_Type", "DefaultCrAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "Voucher_Type", "FirstDrCr", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "AcControls", "OthDealerGrp", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "AcControls", "FSBCrAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDbFA, "AcControls", "FSBOnlinePost", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Estimate1", "ReqNo", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "LabD_Per", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "LastInvDocid", dbText, 21, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Card", "LastLabInvDocId", dbText, 21, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Card", "LastInvNoSuff", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "LastLabInvNoSuff", dbDouble, 10, False, True, 0)
        
        Call AddFieldTable(tmpDb, "SP_Stock", "PurDocNo", dbText, 21, False, True, "")
        Call AddFieldTable(tmpDb, "SP_Stock", "PurDocDate", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Card", "JobType", dbText, 1, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Card", "JobType", dbText, 1, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Lab", "JobCode", dbText, 6, False, True, "")
        
        Call AddFieldTable(tmpDb, "Wrk_Details", "HrsPerDay", dbDouble, 2, False, True, 0)
        Call AddFieldTable(tmpDb, "Wrk_Details", "TotAtms", dbDouble, 2, False, True, 0)
        Call AddFieldTable(tmpDb, "Wrk_Details", "WorkRunCost", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Rect", "DiscAc", dbText, 8, False, True, "")
        Call AddFieldTable(tmpDb, "Rect", "DiscAmt", dbDouble, 10, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Lab", "Mech_Voice", dbText, 40, False, True, "")
        
        '' Vishal Jain (Dhule Changes)
        Call AddFieldTable(tmpDb, "Job_Booking", "SiebelDocID", dbText, 50, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Booking", "SellingDealerCode", dbText, 10, False, True, "")
        
        Call AddFieldTable(tmpDb, "Job_Card", "SiebelDocID", dbText, 50, False, True, "")
        
        Call AddFieldTable(tmpDb, "Model", "Col_Code", dbText, 4, False, True, "")
        Call AddFieldTable(tmpDb, "Model", "RegulatoryCertificate", dbText, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Model", "SteeringType", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDb, "Model", "Vehicle_Drive", dbText, 6, False, True, "")
        Call AddFieldTable(tmpDb, "Model", "FuelTankCapacity", dbDouble, 2, False, True, 0)
        Call AddFieldTable(tmpDb, "Model", "RearAxleMake", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDb, "Model", "FMSN", dbText, 1, False, True, "")
        Call AddFieldTable(tmpDb, "Model", "CubicCapacity", dbText, 10, False, True, "")
        Call AddFieldTable(tmpDb, "Model", "BodyType", dbText, 25, False, True, "")
        
        Call AddFieldTable(tmpDb, "Rect", "SiebelRectNo", dbText, 30, False, True, "")
        
        Call AddFieldTable(tmpDb, "Sp_Purch", "SiebelDocID", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDb, "Sp_Sale", "SiebelDocID", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDb, "Sp_Stock", "SiebelDocID", dbText, 50, False, True, "")
        
        Call AddFieldTable(tmpDb, "SubGroup", "SiebelCode", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDb, "SubGroupAlias", "SiebelCode", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDbFA, "SubGroup", "SiebelCode", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDbFA, "SubGroupAlias", "SiebelCode", dbText, 20, False, True, "")
        
        Call AddFieldTable(tmpDb, "Veh_Order", "SiebelOrderNo", dbText, 25, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Order", "SiebelInvoiceNo", dbText, 25, False, True, "")
        
        Call AddFieldTable(tmpDb, "Veh_Order1", "SiebelOrderNo", dbText, 25, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Order1", "SiebelInvoiceNo", dbText, 25, False, True, "")
        
        Call AddFieldTable(tmpDb, "Syctrl", "SiebelActiveYN", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDb, "Service_Type", "RateEditableYN", dbByte, 1, False, True, 0)
        
        
        'New Field Added for storing LSt No Pervious LstNo will now use for Storing TIN No
        Call AddFieldTable(tmpDb, "Division", "LstDate", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Division", "LstNo", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDb, "Division", "LstDateV", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Division", "LstNoV", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDb, "Division", "LstDateS", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Division", "LstNoS", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDb, "Division", "LstDateW", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Division", "LstNoW", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDb, "AssociatedFirms", "LstDate", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "AssociatedFirms", "LstNo", dbText, 30, False, True, "")
        Call AddFieldTable(tmpDb, "Sp_Stock", "RateType", dbText, 5, False, True, "")
        Call AddFieldTable(tmpDb, "Job_Lab", "RateType", dbText, 5, False, True, "")

        
        Call AddFieldTable(tmpDb, "Job_Demand", "Lab_Rate", dbDouble, 8, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Demand", "Time_Req", dbDouble, 8, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Demand", "Amount", dbDouble, 8, False, True, 0)
        Call AddFieldTable(tmpDb, "Job_Card", "TempCloseDate", dbDate, 15, False, True, "")
        Call AddFieldTable(tmpDb, "Syctrl", "TaxOnFreeLabYN", dbByte, 1, False, True, 0)
        Call AddFieldTable(tmpDbFA, "Ledger", "Chq_Favour", dbText, 255, False, True, "")
        Call AddFieldTable(tmpDbFA, "Ledger", "Chq_AcPayee", dbByte, , False, True, "")
        
        
        Call AddFieldTable(tmpDb, "HisCard", "InsuranceCompany", dbText, 5, False, True, "")
        Call AddFieldTable(tmpDb, "HisCard", "InsuranceExpiry", dbDate, , False, True, "")
        Call AddFieldTable(tmpDb, "HisCard", "InsurancePolicyNo", dbText, 20, False, True, "")
        Call AddFieldTable(tmpDb, "Veh_Order", "SpecialDiscount", dbDouble, , False, True, 0)
        Call AddFieldTable(tmpDbFA, "AcControls", "SpecialDiscountAc", dbText, 8, False, True, "")

        
        
        Set tmpDb = Nothing
        Set tmpDbFA = Nothing
        Set tmpDbComp = Nothing
        
        If GCn.State = 0 Then GCn.Open
        If G_FaCn.State = 0 Then G_FaCn.Open
        
        Set TempRs = GCn.Execute("Select Chassis,Inv_DocId From Veh_Order ")
        If TempRs.RecordCount > 0 Then
            Do Until TempRs.EOF
                GCn.Execute ("Update Veh_Stock Set Sal_DocId='" & XNull(TempRs!Inv_DocId) & "' Where ChassisNo='" & TempRs!Chassis & "' ")
                TempRs.MoveNext
            Loop
        End If
       
            
        'For Upadting Last Purchase DocId And Date
        If MsgBox("Do you want to update Last Invoice No of part for Warranty purpose ? ", vbYesNo) = vbYes Then
            Dim LstPur As ADODB.Recordset, TotalRec As ADODB.Recordset, I As Double
            Set TotalRec = GCn.Execute("Select Distinct Part_No from Sp_Stock")
            If TotalRec.RecordCount > 0 Then
                TotalRec.MoveFirst
                For I = 1 To TotalRec.RecordCount
                    Set LstPur = GCn.Execute("Select SP_Purch.Party_Doc_No,SP_Purch.Party_Doc_Date,SP_Stock.V_rate as Rate,SP_stock.Part_No from SP_Stock Left Join SP_Purch on SP_Stock.DocId=SP_Purch.DOCId where SP_Stock.Part_No='" & TotalRec!Part_No & "' and SP_Stock.V_Type in('SXGR','SXAO') and Left(SP_Stock.DocId,1)='" & PubDivCode & "' order by SP_Purch.Party_Doc_Date DESC")
                     If LstPur.RecordCount > 0 Then
                        LstPur.MoveFirst
                        GCn.Execute ("Update Part Set PurDocId='" & XNull(LstPur!Party_Doc_No) & "', PurDate=" & ConvertDate(LstPur!Party_Doc_Date) & ", PurRate=" & VNull(LstPur!Rate) & " where Part_No ='" & TotalRec!Part_No & "' and Div_Code='" & PubDivCode & "'")
                    End If
                Label1.CAPTION = "Updating Last purchase Invoice No of Part : " & TotalRec!Part_No & "'"
                Label1.Refresh
                TotalRec.MoveNext
                Next
            End If
            Set LstPur = Nothing
        End If
        '*************************************************************************************************
        'For Upadting V_Rate of Opening Stock
        If MsgBox("Do you want to update VRate of opening Stock ? ", vbYesNo) = vbYes Then
            Dim TempVal As Double, mDisPer As Double, TmpRst As ADODB.Recordset
            Set TotalRec = GCn.Execute("Select * from Sp_Stock where V_type='SXAO'")
            For I = 1 To TotalRec.RecordCount
                TempVal = 0: mDisPer = 0
                Set TmpRst = GCn.Execute("Select TP_SRate from Part where Part_No='" & TotalRec!Part_No & "'")
                If TmpRst.RecordCount > 0 Then
                    TempVal = GCn.Execute("Select TP_SRate from Part where Part_No='" & TotalRec!Part_No & "'").Fields(0).Value
                    mDisPer = IIf(IsNull(GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & TotalRec!Part_No & "'").Fields(1).Value), 0, GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & TotalRec!Part_No & "'").Fields(1).Value)
                    If mDisPer > 0 Then
                         TempVal = Round(TempVal - ((TempVal * mDisPer) / 100), 2)
                    End If
                End If
                Label1.CAPTION = "Updating VRate of Part : " & TotalRec!Part_No & "'"
                Label1.Refresh
                GCn.Execute ("Update SP_Stock set v_Rate=" & TempVal & " where DocId='" & TotalRec!DocID & "' and Part_No='" & TotalRec!Part_No & "'")
                TotalRec.MoveNext
            Next
        End If
        
        '**********************************************************************************************
        Label1.CAPTION = "Please Wait ! Table Structure Updation in progress....."
        Label1.Refresh
        Set TotalRec = Nothing
        
        GCn.Execute "Alter table SubGroup ALTER COLUMN Nature TEXT(15)"
        GCn.Execute "Alter table SubGroupAlias ALTER COLUMN Nature TEXT(15)"
        G_FaCn.Execute "Alter table SubGroup ALTER COLUMN Nature TEXT(15)"
        G_FaCn.Execute "Alter table SubGroupAlias ALTER COLUMN Nature TEXT(15)"
        GCn.Execute "Alter table Model ALTER COLUMN Model_Type nVarChar(3)"
        
        GCn.Execute "Alter table Job_Gatepass ALTER COLUMN Purpose TEXT(90)"
        GCn.Execute "Alter table Model_Grp ALTER COLUMN ModelGrp_Code TEXT(6)"
        GCn.Execute "Alter table Model_Grp ALTER COLUMN ModelCat_Code TEXT(6)"
        GCn.Execute "Alter table Model_Cat ALTER COLUMN ModelCat_Code TEXT(6)"
        GCn.Execute "Alter table Model ALTER COLUMN Grp_Code TEXT(6)"
        GCn.Execute "Alter table Model ALTER COLUMN Cat_Code TEXT(6)"
        GCn.Execute "Alter table FinBank ALTER COLUMN FinBankCode TEXT(5)"
        GCn.Execute "Alter table ContractFinance ALTER COLUMN FinBankCode TEXT(5)"
        GCn.Execute "Alter table Model ALTER COLUMN Model_Desc TEXT(80)"
        GCn.Execute "Alter table Model ALTER COLUMN Chas_Type TEXT(9)"
        GCn.Execute "Alter table Veh_Order ALTER COLUMN Chas_Type TEXT(9)"
        GCn.Execute "Alter table Veh_Order1 ALTER COLUMN Chas_Type TEXT(9)"
        GCn.Execute "Alter Table Sp_Stock Alter Column Rate2 Number"
        GCn.Execute "Alter Table Sp_Stock Alter Column Mrp_Rate2 Number"
        GCn.Execute "Alter table DmsData ALTER COLUMN Chassis TEXT(20)"
        'To update Ladger V_Sno
        
        G_CompCn.Open
        
            G_CompCn.Execute "SELECT User_Name, Module_Name, Form_Name, Param_Str INTO UserGroup1 FROM User2 WHERE 1=2"
            G_CompCn.Execute "Create Table UserGroup (User_Name Text(10))"
        
        
        
        
'       G_FaCn.Open
        G_FaCn.Execute "Alter table lEDGER ALTER COLUMN V_SNO INTEGER "
        
        

        
        GCn.Execute "ALTER TABLE Veh_Order  ALTER COLUMN Model TEXT(25) NULL DEFAULT " & "" & ""
        GCn.Execute "ALTER TABLE Veh_Purch1  ALTER COLUMN OBNO TEXT(20) NULL DEFAULT " & "" & ""
        GCn.Execute "ALTER TABLE Veh_Quot1  ALTER COLUMN Model TEXT(25) NULL DEFAULT " & "" & ""
        GCn.Execute "ALTER TABLE Model      ALTER COLUMN Model TEXT(25) NULL DEFAULT " & "" & ""
        GCn.Execute "ALTER TABLE Veh_Stock  ALTER COLUMN Model TEXT(25) NULL DEFAULT " & "" & ""
        GCn.Execute "ALTER TABLE HisCard    ALTER COLUMN Model TEXT(25) NULL DEFAULT " & "" & ""
        GCn.Execute "Alter Table Job_Lab Alter Column Lab_Code Text(8) Null Default " & "" & " "
        
    '   GCn.Execute "ALTER TABLE Model1     ALTER COLUMN Model TEXT(25) NULL DEFAULT " & "" & ""
        
        GCn.Execute "ALTER TABLE Model ALTER COLUMN Vehicle_Type TEXT(25) NULL DEFAULT " & "" & ""

     
     
        '*************************************************************************************************
    '    Drop Query View SubGroup
        
        'G_FaCn.Execute "Drop view ViewSubGroup"
        '*************************************************************************************************
        
        '*************************************************************************************************
        AddNewFieldForDatamanFa (PubFADataPath)
        '*************************************************************************************************
             
          'For Updation in Data transfer protocol
        GCn.Execute "Update  TableGroupClient set  UserDate='' where Table_Name='SubGroupCounter'"
        GCn.Execute "Update  TableGroupHO set  UserDate='' where Table_Name='SubGroupCounter'"
        '*************************************************************************************************
        GCn.Execute "UPDATE Job_Lab SET Hrs_Taken = val(int(Hrs_Taken)&'.'& ((Hrs_Taken -int(Hrs_Taken))*60))"
        
            
        '*************************************************************************************************
       ' Removing Nulls
            'RemoveNulls
        '*************************************************************************************************
         'Do passwd field of 15 char
    '      G_CompCn.Execute "Alter table UserMast ALTER COLUMN PASSWD TEXT(15)"
        '*************************************************************************************************
        Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='W_SAC'")
        If TmpRst.RecordCount = 0 Then
            G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                          "('Workshop','','W_SAC','Workshop Acc Bill Cash','WorksAccBillCa','WkABillCa','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
            G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('W_SAC'," & ConvertDate(PubStartDate) & "#," & ConvertDate(PubEndDate) & "#,'W_SAC',900000,'" & PubDivCode & "')")
        End If
        
        Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='W_SAR'")
        If TmpRst.RecordCount = 0 Then
            G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                          "('Workshop','','W_SAR','Workshop Acc Bill Credit','WorksAccBillCredit','WkABillCr','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
            G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('W_SAR'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'W_SAR',900000,'" & PubDivCode & "')")
        End If
        
        Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='W_LAC'")
        If TmpRst.RecordCount = 0 Then
            G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                          "('Workshop','','W_LAC','Workshop Acc Lab Cash','WorksAccLabCash','WkALabCa','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
            G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('W_LAC'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'W_LAC',900000,'" & PubDivCode & "')")
        End If
        
        Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='W_LAR'")
        If TmpRst.RecordCount = 0 Then
            G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                          "('Workshop','','W_LAR','Workshop Acc Lab Credit','WorksAccLabCredit','WkALabCr','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
            G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('W_LAR'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'W_LAR',900000,'" & PubDivCode & "')")
        End If
        
        
        Set TmpRst = G_FaCn.Execute("Select * from Voucher_Type where V_Type='" & Voucher_NCat_BankPayment & "'")
        If TmpRst.RecordCount = 0 Then
            G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                          "('" & Voucher_Category_Payment & "','" & Voucher_NCat_BankPayment & "','" & Voucher_NCat_BankPayment & "','Bank Payment Voucher','BankPaymentVoucher','" & Voucher_NCat_BankPayment & "','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
            
            G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('" & Voucher_NCat_BankPayment & "'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'" & Voucher_Category_Payment & "',900000,'" & PubDivCode & "')")
        End If
        
        
        Set TmpRst = G_FaCn.Execute("Select * from Voucher_Type where V_Type='" & Voucher_NCat_CashPayment & "'")
        If TmpRst.RecordCount = 0 Then
            G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                          "('" & Voucher_Category_Payment & "','" & Voucher_NCat_BankPayment & "','" & Voucher_NCat_CashPayment & "','Cash Payment Voucher','CashPaymentVoucher','" & Voucher_NCat_CashPayment & "','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
            
            G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('" & Voucher_NCat_CashPayment & "'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'" & Voucher_Category_Payment & "',900000,'" & PubDivCode & "')")
        End If
        
        
        'Adding Records Of Add Edit Delete Permissions Deleted Automatically for FA Forms
        Set RsUser = G_CompCn.Execute("Select User_Name From UserMast")
        If RsUser.RecordCount > 0 Then
            Do Until RsUser.EOF
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FaGrEnt' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FaGrEnt', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='frmSubGroup' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'frmSubGroup', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FAVRENT' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FAVRENT', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FAVTYPE' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FAVTYPE', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FATDSCAT' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FATDSCAT', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                            
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FATDSCERTIFICATE' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FATDSCERTIFICATE', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                            
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FATDSCHAL' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FATDSCHAL', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                            
                RsUser.MoveNext
            Loop
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
        
        
        If IsTableExists(Pub_DataPath & "\" & PubCenDataPath & "\Automan.mdb", "AcMERGE") = True Then
            GCn.Execute ("Delete * from AcMerge")
            GCn.Execute ("insert into AcMERGE values('SaleTarget','Party_Code','Automan')")
            GCn.Execute ("insert into AcMERGE values('SP_Order','Party_Code','Automan')")
            GCn.Execute ("insert into AcMERGE values('SP_Order1','Party_Code','Automan')")
            GCn.Execute ("insert into AcMERGE values('SP_Purch','Party_Code','Automan')")
            GCn.Execute ("insert into AcMERGE values('SP_Sale','Party_Code','Automan')")
            GCn.Execute ("insert into AcMERGE values('SP_Stock','Party_Code','Automan')")
            GCn.Execute ("insert into AcMERGE values('Veh_Order','PartyCode','Automan')")
            GCn.Execute ("insert into AcMERGE values('Veh_Order1','PartyCode','Automan')")
            GCn.Execute ("insert into AcMERGE values('Veh_Purch1','PartyCode','Automan')")
            
            GCn.Execute ("insert into AcMERGE values('Veh_Stock','PartyCode','Automan')")
            GCn.Execute ("insert into AcMERGE values('Rect','PartyCode','Automan')")
            
            GCn.Execute ("insert into AcMERGE values('Ledger','SubCode','FaDATA')")
            GCn.Execute ("insert into AcMERGE values('LedgerAdj','SubCode','FaDATA')")
            
            GCn.Execute ("insert into AcMERGE values('LedgerRef','SubCode','FaDATA')")
            GCn.Execute ("insert into AcMERGE values('LedgerAdj','SubCode','FaDATA')")
        End If
        
        GCn.Execute ("Update Model set Manufacturer='TATA MOTORS LTD.'")
        
        
        'filling A/C Control TABLE
        Dim fname As String, Val1 As String
        For I = 0 To 52
            fname = G_FaCn.Execute("select * from accontrols").Fields(I).Name
            Val1 = G_FaCn.Execute("select * from accontrols").Fields(I).DefinedSize
            If Val1 = 8 Then
                GCn.Execute ("insert into AcMERGE values('AcControls','" & fname & "','FaDATA')")
            End If
        Next
        For I = 0 To 100
            fname = GCn.Execute("select * from syctrl").Fields(I).Name
            Val1 = GCn.Execute("select * from syctrl").Fields(I).DefinedSize
            If Val1 = 8 Then
                GCn.Execute ("insert into AcMERGE values('syctrl','" & fname & "','Automan')")
            End If
        Next
        
        
'        GCn.Execute "Delete From Job_Lab  Where Job_DocId Not In (Select DocId From Job_Card)"
'        GCn.Execute "Delete From Job_Lab2 Where Job_DocId Not In (Select DocId From Job_Card)"
'
'        GCn.Execute "Delete From Sp_Stock Where Job_DocId Not In (Select DocId From Job_Card) And Job_DocId <> '' And Job_DocId Is Not Null"
'        GCn.Execute "Delete From Sp_Sale Where Job_DocId Not In (Select DocId From Job_Card)  And Job_DocId <> '' And Job_DocId Is Not Null"
        
        
        
        
        
        If PubSiebelActiveYn Then
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='SBLCQ'")
            If TmpRst.RecordCount = 0 Then
                G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                              "('GENFA','','SBLCQ','Siebel Receipts - Cheque/Draft','SiebelReceiptsChequeDraft','SBLCD','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
                G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('SBLCQ'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'SBLCQ',700000,'" & PubDivCode & "')")
            End If
            
            
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='SBLCS'")
            If TmpRst.RecordCount = 0 Then
                G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                              "('GENFA','','SBLCS','Siebel Receipts - Cash','SiebelReceiptsCash','SBLCS','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
                G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('SBLCS'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'SBLCS',700000,'" & PubDivCode & "')")
            End If
            
            
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='SBLRO'")
            If TmpRst.RecordCount = 0 Then
                G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                              "('GENFA','','SBLRO','Siebel Receipt - Release Order','SiebelReceiptReleaseOrder','SBLRO','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
                G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('SBLRO'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'SBLRO',700000,'" & PubDivCode & "')")
            End If
            
        End If
        
        
        
        
        If UCase(left(PubComp_Name, 3)) = "JMK" Then
            GCn.Execute "Delete From Part Where Left(Part_No,6)='tinter'"
        End If
    
        GCn.Execute "Update Job_Card Set eCessPer=" & cIIF("Lab_TaxPer=12.24 Or Lab_TaxPer=12.25", "2", "3") & ", eCessAmt=Lab_TaxAmt*" & cIIF("Lab_TaxPer=12.24 Or Lab_TaxPer=12.25", "2", "3") & "/(100+" & cIIF("Lab_TaxPer=12.24 Or Lab_TaxPer=12.25", "2", "3") & ")"
        GCn.Execute "Update Job_Card Set eCessPer=0, eCessAmt=0 where Lab_TaxPer=0"
        
        If UCase(left(PubComp_Name, 5)) = "UJWAL" Then
            GCn.Execute "Update Sp_Stock Set Sp_Stock.Rate2=Sp_Stock.Rate, Sp_Stock.Mrp_Rate2=Sp_Stock.Mrp_Rate, " & _
                        "Sp_Stock.Amount2=Sp_Stock.Amount, Sp_Stock.Disc_Per2=Sp_Stock.Disc_Per, Sp_Stock.Disc_Amt2=Sp_Stock.Disc_Amt, " & _
                        "Sp_Stock.Net_Amt2=Sp_Stock.Net_Amt Where Sp_Stock.DocId In (Select DocId From Sp_Sale Where SiebelDocId <> '' and SiebelDocId Is Not Null)"
            
'            GCn.Execute "Update Job_Card Set LabAmt_TB=0, NetLab_Amt=0 Where Lab_taxPer=0 and Left(DocId,1)='P'"
'            GCn.Execute "Update Job_Card Set Lab_TaxPer=12.24, Lab_TaxAmt=LabAmt_TB*12.24/100 Where Lab_taxPer=0.01 and Left(DocId,1)='P'"
'            GCn.Execute "Update Job_Lab Set LabourAmt=0 Where Job_DocId In (Select DocId From Job_Card Where NetLab_Amt=0 And  Left(DocId,1)='P')"
            GCn.Execute "Update sp_Stock Set  Disc_amt2=0 Where disc_Per2=0"
            GCn.Execute "Update Job_Card, HisCard Set BillingName=Name Where Job_Card.CardNo=HisCard.CardNo And CrMemo=0 And (BillingName='' Or BillingName Is Null)"
            
'            GCn.Execute "Update Sp_Stock Set TaxPer=12.5 Where Qty_Iss>0 "
'            GCn.Execute "Update Sp_Stock Set TaxAmt = IIf(Mrp_Yn=1,(Amount2-Disc_Amt2)*TaxPer/(100+TaxPer),(Amount2-Disc_Amt2)*TaxPer/100) WHERE Qty_Iss>0"
'            GCn.Execute "Update Sp_Stock Set Net_Amt2=IIF(Mrp_YN=1,(Amount2-Disc_Amt2-TaxAmt),(Amount2-Disc_Amt2)) Where Qty_Iss>0"
'            GCn.Execute "Update Sp_Stock Set Net_Amt=IIF(Mrp_YN=1,(Amount-Disc_Amt-TaxAmt),(Amount-Disc_Amt)) Where Qty_Iss>0"
                
'            GCn.Execute "Update Sp_Stock Set TempNetAmt=Net_Amt, TempNetAmt2=Net_Amt2 Where Left(DocId,1)='P'"
'            GCn.Execute "Update Sp_Stock Set TaxPer=12.5 Where Left(DocId,1)='P' And Qty_Iss>0 "
'            GCn.Execute "Update Sp_Stock Set TaxAmt = IIf(Mrp_Yn=1,(Amount2-Disc_Amt2)*TaxPer/(100+TaxPer),(Amount2-Disc_Amt2)*TaxPer/100) WHERE Qty_Iss>0"
'            GCn.Execute "Update Sp_Stock Set Net_Amt2=(Amount2-Disc_Amt2-TaxAmt) Where Qty_Iss>0"
'            GCn.Execute "Update Sp_Stock Set Net_Amt=(Amount-Disc_Amt-TaxAmt) Where Qty_Iss>0"
                                    
'            Set RsTemp = New ADODB.Recordset
'            RsTemp.CursorLocation = adUseClient
'            RsTemp.Open "Select DocId, Packing From Sp_Sale Where Left(DocId,1)='P'", GCn, adOpenDynamic, adLockOptimistic
'            If RsTemp.RecordCount > 0 Then
'                Do Until RsTemp.EOF
'                    RsTemp!Packing = Format(VNull(GCn.Execute("Select Sum(Net_Amt2-TempNetAmt2) From Sp_Stock Where Invoice_DocId='" & RsTemp!DocID & "' ").Fields(0).Value), "0.00")
'                    RsTemp.Update
'                    RsTemp.MoveNext
'                Loop
'            End If
        End If
        GCn.Execute "ALTER TABLE Syctrl ALTER COLUMN HelpLineNo   TEXT(21) NULL DEFAULT " & "" & ""
        GCn.Execute "ALTER TABLE Part_Grade ALTER COLUMN PartGrade_Name   TEXT(30) NULL DEFAULT " & "" & ""
        
        
''        If UCase(left(PubComp_Name, 5)) = "UJWAL" Then
''            GCn.Execute "Update SubGroup, DmsSubGroup  Set SiebelCode=DmsSubCode  Where SubGroup.SubCode=DmsSubGroup.AutomanSubCode"
''            GCn.Execute "Update Subgroup S, DmsSubGroup D Set S.Name=left(Trim(D.Name) & ' [' & Trim(D.DmsSubCode) & ']', 40) Where S.SiebelCode=D.DmsSubCode"
''            GCn.Execute "Update SubgroupAlias S, DmsSubGroup D Set S.Name=left(Trim(D.Name) & ' [' & Trim(D.DmsSubCode) & ']', 40) Where S.SiebelCode=D.DmsSubCode"
''            GCn.Execute "Update " & FaTable("Subgroup") & " S, DmsSubGroup D Set S.Name=left(Trim(D.Name) & ' [' & Trim(D.DmsSubCode) & ']', 40) Where S.SiebelCode=D.DmsSubCode"
''            GCn.Execute "Update " & FaTable("SubgroupAlias") & " S, DmsSubGroup D Set S.Name=left(Trim(D.Name) & ' [' & Trim(D.DmsSubCode) & ']', 40) Where S.SiebelCode=D.DmsSubCode"
''        End If
        
        
        Dim mVType$, mVDesc$, mVPrefix$
        Dim mStartSrlNo As Double
        Dim RsDiv As ADODB.Recordset
        
        
        
                
        If PubSiebelActiveYn Then
            Set RsDiv = GCn.Execute("Select Div_Code From Division")
            Do Until RsDiv.EOF
                mVPrefix = "DMS"
                mStartSrlNo = Right(date, 1) & "00000"
                
                CreateVType "FA", "JV", "D_SCS", "Spare Cash Sale", mVPrefix, mStartSrlNo, XNull(RsDiv!Div_Code)
                CreateVType "FA", "JV", "D_SRS", "Spare Credit Sale", mVPrefix, mStartSrlNo, XNull(RsDiv!Div_Code)
                CreateVType "FA", "JV", "D_VRP", "Vehicle Purchase", mVPrefix, mStartSrlNo, XNull(RsDiv!Div_Code)
                CreateVType "FA", "JV", "D_VRS", "Vehicle Sale", mVPrefix, mStartSrlNo, XNull(RsDiv!Div_Code)
                CreateVType "FA", "JV", "D_SCP", "Spare Cash Purchase", mVPrefix, mStartSrlNo, XNull(RsDiv!Div_Code)
                CreateVType "FA", "JV", "D_SRP", "Spare Credit Purchase", mVPrefix, mStartSrlNo, XNull(RsDiv!Div_Code)
                CreateVType "FA", "JV", "D_BR", "Bank Receipt", mVPrefix, mStartSrlNo, XNull(RsDiv!Div_Code)
                CreateVType "FA", "JV", "D_CR", "Cash Receipt", mVPrefix, mStartSrlNo, XNull(RsDiv!Div_Code)
                CreateVType "FA", "JV", "D_WRS", "Workshop Credit Sale", mVPrefix, mStartSrlNo, XNull(RsDiv!Div_Code)
                CreateVType "FA", "JV", "D_WCS", "Workshop Cash Sale", mVPrefix, mStartSrlNo, XNull(RsDiv!Div_Code)
                                                                
                RsDiv.MoveNext
            Loop
        End If
    
        
        'GCn.Execute "ALTER TABLE Veh_OrderM ALTER COLUMN Chassis Text(20)"
        GCn.Execute "ALTER TABLE CustInfo ALTER COLUMN Chassis Text(20)"
        GCn.Execute "ALTER TABLE Job_Inspection ALTER COLUMN Chassis Text(20)"
        GCn.Execute "ALTER TABLE Veh_Order1 ALTER COLUMN Chassis Text(20)"
        GCn.Execute "ALTER TABLE Veh_Order ALTER COLUMN Chassis Text(20)"
        GCn.Execute "ALTER TABLE HisCard ALTER COLUMN Chassis Text(20)"
        GCn.Execute "ALTER TABLE Job_Booking ALTER COLUMN Chassis Text(20)"
        GCn.Execute "ALTER TABLE Veh_InvCancel ALTER COLUMN Chassis Text(20)"
        GCn.Execute "ALTER TABLE Estimate ALTER COLUMN Chassis Text(20)"
        GCn.Execute "Alter Table RTOData Drop Constraint PrimaryKey "
        GCn.Execute "ALTER TABLE RTOData ALTER COLUMN CHASSIS_NO NVARCHAR(20) NOT Null"
        GCn.Execute "Alter Table RTOData Add Constraint PrimaryKey Primary Key (CHASSIS_NO,MODEL)"
        GCn.Execute "Alter Table Veh_Stock Drop Constraint PrimaryKey"
        GCn.Execute "ALTER TABLE Veh_Stock ALTER COLUMN ChassisNo Text(20) NOT Null"
        GCn.Execute "Alter Table Veh_Stock Add Constraint PrimaryKey Primary Key (CHASSISNO)"
        
        GCn.Execute "ALTER TABLE Veh_Transfer ALTER COLUMN ChassisNo Text(20)"
        GCn.Execute "Alter Table Veh_CheckList Drop Constraint PrimaryKey"
        GCn.Execute "ALTER TABLE Veh_CheckList ALTER COLUMN ChassisNo Text(20) NOT Null"
        GCn.Execute "Alter Table Veh_CheckList Add Constraint PrimaryKey Primary Key (ChassisNo,Item_Code,MODEL)"
        
        GCn.Execute "ALTER TABLE Job_Inspection ALTER COLUMN Engine Text(25)"
        GCn.Execute "ALTER TABLE HisCard ALTER COLUMN Engine Text(25)"
        GCn.Execute "ALTER TABLE Job_Booking ALTER COLUMN Engine Text(25)"
        GCn.Execute "ALTER TABLE Job_Warr1 ALTER COLUMN Engine Text(25)"
        GCn.Execute "ALTER TABLE Estimate ALTER COLUMN Engine Text(25)"
        GCn.Execute "ALTER TABLE RTOData ALTER COLUMN ENGINE_NO Text(25)"
        GCn.Execute "ALTER TABLE Veh_Stock ALTER COLUMN EngineNo Text(25)"
        GCn.Execute "ALTER TABLE Veh_Transfer ALTER COLUMN EngineNo Text(25)"
        
       
        
        
        
ElseIf PubBackEnd = "S" Then
    
    GCn.Execute "Update Veh_Stock Set Veh_Stock.Sal_DocId=Veh_Order.Inv_DocId, Sal_DocIdHelp=Inv_DocIdHelp, " & _
                "Sal_Site_Code=Inv_SiteCode, Sal_VType=Inv_VType, Sal_VNo=Inv_No, Sal_VDate=Inv_Date, " & _
                "Veh_Stock.DelCh_DocId=Veh_Order.DelCh_DocId, Veh_Stock.DelCh_Date=Veh_Order.DelCh_Dt, " & _
                "Veh_Stock.Ord_DocId=Veh_Order.OrdDocId, Veh_Stock.Ord_SiteCode=Veh_Order.Ord_SiteCode " & _
                "From Veh_Order Where Veh_Stock.ChassisNo=Veh_Order.Chassis"
                
    If UCase(left(PubComp_Name, 3)) = "LMP" Then
        GCn.Execute ("Update SP_Sale Set Party_Name='CASH' Where Cash_Credit='Cash' And V_Type='SYSIC'")
    End If
    
    GCn.Execute "Update Job_Card Set eCessPer=" & cIIF("Lab_TaxPer=12.24 Or Lab_TaxPer=12.25", "2", "3") & ", eCessAmt=Lab_TaxAmt*" & cIIF("Lab_TaxPer=12.24 Or Lab_TaxPer=12.25", "2", "3") & "/(100+" & cIIF("Lab_TaxPer=12.24 Or Lab_TaxPer=12.25", "2", "3") & ")"
    GCn.Execute "Update Job_Card Set eCessPer=0, eCessAmt=0 where Lab_TaxPer=0"
        
    GCn.Execute "ALTER TABLE Syctrl ALTER COLUMN HelpLineNo  VarChar(21)"
    GCn.Execute "ALTER TABLE Part_Grade ALTER COLUMN PartGrade_Name  VarChar(30)"
     GCn.Execute "Alter table Model ALTER COLUMN Model_Type nVarChar(3)"
    GCn.Execute "Alter Table  Veh_Order Add DoReciveDate DateTime   Default ''"
    GCn.Execute "Alter Table  Veh_Order Add DoIssueDate DateTime   Default ''"
    GCn.Execute "Update AcGroup Set AliasYn='N'"


    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='F_OP'")
            If TmpRst.RecordCount = 0 Then
                G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                              "('FA','JV','F_OP','OTHER PURCHASES','OTHERPURCHASES','F_OP','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
            End If
        
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='F_OS'")
            If TmpRst.RecordCount = 0 Then
                G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                              "('FA','JV','F_OS','OTHER SALES','OTHERSALES','F_OS','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
            End If
        
        
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='SXPTR'")
            If TmpRst.RecordCount = 0 Then
                G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                              "('SPARE','','SXPTR','OTHER PURCHASES CREDIT','OTHERPURCHASESCREDIT','SXPTR','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
            End If
                        
        
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='SXPTC'")
            If TmpRst.RecordCount = 0 Then
                G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                              "('SPARE','','SXPTC','OTHER PURCHASES CASH','OTHERPURCHASESCASH','SXPTC','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
            End If
        
        
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='F_CN'")
            If TmpRst.RecordCount = 0 Then
                G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                              "('FA','JV','F_CN','CREDIT NOTE','CREDITNOTE','F_CN','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
            End If
        
    End If
    CreateVType "EXP", "EXP", "E_EXP", "EMPLOYEE WISE EXPENCES", CStr(Format(PubStartDate, "YYYY")), Val(Right(PubStartDate, 1) & "00000"), PubDivCode
    CreateVType Voucher_Category_Payment, Voucher_NCat_BankPayment, Voucher_NCat_BankPayment, "Bank Payment Voucher", CStr(Format(PubStartDate, "YYYY")), Val(Right(PubStartDate, 1) & "00000"), PubDivCode
    CreateVType Voucher_Category_Payment, Voucher_NCat_CashPayment, Voucher_NCat_CashPayment, "Cash Payment Voucher", CStr(Format(PubStartDate, "YYYY")), Val(Right(PubStartDate, 1) & "00000"), PubDivCode
        
    AddNewField GCn, "SubGroup", "ChequeReportName", "nVarchar(50)"
    AddNewField GCn, "SubGroupAlias", "ChequeReportName", "nVarchar(50)"
    AddNewField GCn, "Labour", "Dep_Item", "nVarchar(5)"
    AddNewField GCn, "Veh_Transfer", "Narration", "nVarchar(255)"
        
        
        'Nikhil
        
        CreateNewTable_sql
        AddNewField GCn, "Part", "Dep_Item", "nVarchar(5)"
         AddNewField GCn, "Labour", "Dep_Item", "nVarchar(5)"
        
        
        
        AddNewField GCn, "Sp_stock", "Dep_Item", "nVarchar(5)"
        AddNewField GCn, "Sp_stock", "Dep_Code", "nVarchar(5)"
        AddNewField GCn, "Sp_stock", "DepitemPer", "Float"
        AddNewField GCn, "Sp_stock", "DepPer", "Float"
        AddNewField GCn, "Sp_stock", "DepAmt", "Float"
        
 'sp_purch
          AddNewField GCn, "Sp_stock", "SFCPer", "Float"
          AddNewField GCn, "sp_purch", "SFCAMT", "Float"
        AddNewField GCn, "Sp_stock", "SFCAmt", "Float"
                 AddNewField GCn, "TaxForms", "SFCPer", "Float"
        
        AddNewField GCn, "Sp_stock", "InsuranceAmt", "Float"
        AddNewField GCn, "Sp_stock", "DiffPeried", "Float"
'Excise_Amt
AddNewField GCn, "Sp_stock", "Excise_Amt", "Float"
AddNewField GCn, "Sp_sale", "Excise_Amt", "Float"

 
        AddNewField GCn, "Job_LAb", "Dep_Item", "nVarchar(5)"
        AddNewField GCn, "Job_LAb", "Dep_Code", "nVarchar(5)"
        AddNewField GCn, "Job_LAb", "DepitemPer", "Float"
        AddNewField GCn, "Job_LAb", "DepPer", "Float"
        AddNewField GCn, "Job_LAb", "DepAmt", "Float"
        AddNewField GCn, "Job_LAb", "InsuranceAmt", "Float"
        AddNewField GCn, "Job_LAb", "DiffPeried", "Float"
        
        



'GCn.Execute " CREATE TABLE dbo.Deprecation_itemMaster(Code         NVARCHAR (5),ShortName  NVARCHAR (2) ,Description  NVARCHAR (50),Dep_per FLOAT , " & _
'    "Site_Code    NVARCHAR (2) CONSTRAINT DF_Deprecation_itemMaster_Site_Code DEFAULT ('') NOT NULL,U_Name       NVARCHAR (15) CONSTRAINT DF_Deprecation_itemMaster_U_Name DEFAULT ('') NOT NULL, U_EntDt      SMALLDATETIME NOT NULL, U_AE         NVARCHAR (1) CONSTRAINT DF_Deprecation_itemMaster_U_AE DEFAULT ('') NOT NULL, " & _
'    " CONSTRAINT PK_Deprecation_itemMaster PRIMARY KEY (Code),CONSTRAINT IX_Deprecation_itemMaster UNIQUE (Description) )"
'
'
'
'GCn.Execute "CREATE TABLE dbo.Deprecation_Master( Code         NVARCHAR (5) ,Dep_Month  FLOAT ,Dep_per FLOAT ,Site_Code    NVARCHAR (2) CONSTRAINT DF_Deprecation_Master_Site_Code DEFAULT ('') NOT NULL," & _
'    " U_Name       NVARCHAR (15) CONSTRAINT DF_Deprecation_Master_U_Name DEFAULT ('') NOT NULL, U_EntDt      SMALLDATETIME NOT NULL, U_AE         NVARCHAR (1) CONSTRAINT DF_Deprecation_Master_U_AE DEFAULT ('') NOT NULL, " & _
'    " CONSTRAINT PK_Deprecation_Master PRIMARY KEY (Code), CONSTRAINT IX_Deprecation_Master UNIQUE (Dep_Month))"


        
AddNewField GCn, "Syctrl", "SiteWiseDisplaY_N", "Bit", 0

'   For Removing Duplicate Records From SubGroup
'    If StrCmp(left(PubComp_Name, 4), "Enar") Then
'    GCn.BeginTrans
'        GCn.Execute "Drop Table SubGroupAlias"
'        GCn.Execute "Select * Into SubGroupAlias From SubGroup"
'        GCn.Execute "Delete From SubGroup"
'        Set RsTemp = GCn.Execute("Select * From SubGroupAlias")
'        If RsTemp.RecordCount > 0 Then
'            Do Until RsTemp.EOF
'                If GCn.Execute("Select Count(*) From SubGroup Where SubCode='" & RsTemp!SubCode & "'").Fields(0) = 0 Then
'                    With RsTemp
'                        GCn.Execute ("Insert Into SubGroup (AcID, Site_Code, SubCode, FirmCode, NamePrefix, Name, NameBiLang, NameHelp, GroupCode, GroupNature, " & _
'                                    "Nature, AliasYN, ConPrefix, ConPrefixBiLang, ConPerson, ConPersonBiLang, ConSuffix, Add1, Add1BiLang, Add2, " & _
'                                    "Add2BiLang, Add3, Add3BiLang, Area, CityCode, Pin, Phone, Mobile, FAx, EMail, " & _
'                                    "Curr_Bal, CstNo, LstNo, PanNo, ITWARD_NO, TDS_Catg, ActiveYN, Govt_YN, CreditLimit, CreditDays, " & _
'                                    "FPrefix, fname, TAdd1, TAdd2, TAdd3, TCityCode, TPIN, TPhone, L_C, FB_Code, Religion, " & _
'                                    "Party_Type, Transporter, Remark, U_Name, U_EntDt, U_AE, OldCode, xName, AcCode, AreaCode, " & _
'                                    "Category, CostCenterAppl, PhoneO, Transport, RC_No, NewYrSubCode, FinCode, SiebelCode) " & _
'                                    "Values('" & XNull(!AcID) & "', '" & XNull(!Site_Code) & "', '" & XNull(!SubCode) & "', '" & XNull(!FirmCode) & "', '" & XNull(!NamePrefix) & "', '" & XNull(!Name) & "', '" & XNull(!NameBiLang) & "', '" & XNull(!NameHelp) & "', '" & XNull(!GroupCode) & "', '" & XNull(!GroupNature) & "', " & _
'                                    "'" & XNull(!Nature) & "', '" & XNull(!AliasYN) & "', '" & XNull(!ConPrefix) & "', '" & XNull(!ConPrefixBiLang) & "', '" & XNull(!ConPerson) & "', '" & XNull(!ConPersonBiLang) & "', '" & XNull(!ConSuffix) & "', '" & XNull(!Add1) & "', '" & XNull(!Add1BiLang) & "', '" & XNull(!Add2) & "', " & _
'                                    "'" & XNull(!Add2BiLang) & "', '" & XNull(!Add3) & "', '" & XNull(!Add3BiLang) & "', '" & XNull(!Area) & "', '" & XNull(!CityCode) & "', '" & XNull(!Pin) & "', '" & XNull(!Phone) & "', '" & XNull(!Mobile) & "', '" & XNull(!FAx) & "', '" & XNull(!EMail) & "', " & _
'                                    "" & VNull(!Curr_Bal) & ", '" & XNull(!CstNo) & "', '" & XNull(!LstNo) & "', '" & XNull(!PanNo) & "', '" & XNull(!ITWARD_NO) & "', '" & XNull(!TDS_Catg) & "', '" & XNull(!ActiveYN) & "', '" & XNull(!Govt_YN) & "', " & VNull(!CreditLimit) & ", " & VNull(!CreditDays) & ", " & _
'                                    "'" & XNull(!FPrefix) & "', '" & XNull(!fname) & "', '" & XNull(!TAdd1) & "', '" & XNull(!TAdd2) & "', '" & XNull(!TAdd3) & "', '" & XNull(!TCityCode) & "', '" & XNull(!TPIN) & "', '" & XNull(!TPhone) & "', '" & XNull(!L_C) & "', '" & XNull(!FB_Code) & "', " & _
'                                    "'" & XNull(!Religion) & "', '" & XNull(!Party_Type) & "', '" & XNull(!Transporter) & "', '" & XNull(!Remark) & "', '" & XNull(!U_Name) & "', " & ConvertDate(!U_EntDt) & ", '" & XNull(!U_AE) & "', '" & XNull(!OldCode) & "', '" & XNull(!xName) & "', '" & XNull(!AcCode) & "', " & _
'                                    "'" & XNull(!AreaCode) & "', '" & XNull(!Category) & "', '" & XNull(!CostCenterAppl) & "', '" & XNull(!PhoneO) & "', '" & XNull(!Transport) & "', '" & XNull(!RC_No) & "', '" & XNull(!NewYrSubCode) & "', '" & XNull(!FinCode) & "', '" & XNull(!SiebelCode) & "') ")
'                    End With
'                End If
'
'                RsTemp.MoveNext
'            Loop
'        End If
'        GCn.Execute "Drop Table SubGroupAlias"
'        GCn.Execute "Select * InTo SubGroupAlias From SubGroup"
'    GCn.CommitTrans
'        GCn.Execute "Alter Table SubGroup Alter Column SubCode VarChar(8) Not Null"
'        GCn.Execute "Alter Table SubGroup Add Constraint PK_SubGroup Primary Key (SubCode)"
'
'        Set RsTemp = Nothing
'
'    End If



    'Qry for ENAR because there r some records not found in veh_purch and found in veh_stock
    ''''''''insert into veh_purch1 select Pur_DocId As DocId, Pur_docId as DocIdHelp, Pur_SiteCode As Site_Code, Pur_VType As V_Type, Pur_VNo As V_No, Pur_VDate As V_Date, PartyCode, PBill_No, PBill_Date, '' As OBNO, Null As OBDate, '' As Trn_Type, Null As Trn_Date, 'CI' As Bms_Category, 0 As Rso_Work, '' As Rso_Code, PBill_Date As DueDate, '' As Gate, PBill_Date As GateDate, 'VF06' As Form_Code, '' As FormNo, Null As FormIssRecDate, '' As Form_No, Null As Form_Date, '' As RoadPermit_FormCode, '' As RoadPermit_No, Rate As Amount, 0 As Addition, VRate-(Rate+(VRate*12.5/112.5)) As Deduction, 0 As Exsice, 12.5 as Tax_Per, (VRate*12.5/112.5) As Tax_Amt, 0 As TaxSur_Per, 0 As TaxSur_Amt, 0 As Misc_Amt, VRate As Tot_Amount, 0 As P_Amount, 0 As Adj_Amt, U_Name, U_EntDt, U_AE, Null As Trf_Date, '11003278' As DrAcCode, '' As AcPostByU_Name, Null As AcPostByU_EntDt, 0 As SubVentionCredit     from veh_stock where Pur_DocId Not In (Select DocId from Veh_Purch1)
    
    
'    Dim mCnt As String
'    Dim rsTemp1 As ADODB.Recordset
'    If StrCmp(left(PubComp_Name, 5), "UJwal") Then
'        Set RsTemp = GCn.Execute("Select * from Model_Cat Order By ModelCat_Code")
'        If RsTemp.RecordCount > 0 Then
'            i = 11
'            Do Until RsTemp.EOF
'                mCnt = left(RsTemp!ModelCat_Code, 1) & CStr(i)
'                i = i + 1
'                GCn.Execute "Update Model Set Cat_Code = '" & mCnt & "' Where Cat_Code = '" & RsTemp!ModelCat_Code & "'"
'                GCn.Execute "Update Model_Grp Set ModelCat_Code = '" & mCnt & "' Where ModelCat_Code = '" & RsTemp!ModelCat_Code & "'"
'                GCn.Execute "Update Model_Cat Set ModelCat_Code = '" & mCnt & "' Where ModelCat_Code = '" & RsTemp!ModelCat_Code & "'"
'                RsTemp.MoveNext
'            Loop
'        End If
'
'
'
'        Set RsTemp = GCn.Execute("Select * from Model_Grp")
'        If RsTemp.RecordCount > 0 Then
'            i = 100
'            Do Until RsTemp.EOF
'
'                mCnt = left(RsTemp!ModelGrp_Code, 1) & CStr(i)
'                i = i + 1
'                GCn.Execute "Update Model Set Grp_Code = '" & mCnt & "' Where Grp_Code = '" & RsTemp!ModelGrp_Code & "'"
'                GCn.Execute "Update Model_Grp Set ModelGrp_Code = '" & mCnt & "' Where ModelGrp_Code = '" & RsTemp!ModelGrp_Code & "'"
'                RsTemp.MoveNext
'            Loop
'        End If
'    End If
'
    
End If
    
    
    
    
    
FilmMenu
    
    
'COMMON UPDATE TABLE STRUCTURE
    If PubBackEnd = "S" Then
        GCn.Execute "Update Part Set NDP = Mrp-(MRP*PurcDisc_Per/100) From Part_DiscFactor Where Part.Disc_Factor=Part_DiscFactor.DiscFac_CatG"
    Else
        GCn.Execute "Update Part, Part_DiscFactor Set NDP = Mrp-(MRP*PurcDisc_Per/100)  Where Part.Disc_Factor=Part_DiscFactor.DiscFac_CatG"
    End If
    
        Dim mParticular$
        mParticular = "Update LastRate to Warranty Issue"
        Log_StructureCreate mParticular
        If GCn.Execute("Select Count(*) From Log_Structure Where Particular = '" & mParticular & "' And Remark='OK'").Fields(0).Value = 0 Then
            If PubBackEnd = "S" Then
                GCn.Execute "Update Sp_Stock Set Sp_Stock.V_Rate=" & cIIF("Part.MRP=0", "Sp_Stock.Rate", "Part.NDP") & " From Part Where Sp_Stock.Part_No=Part.Part_No  "
                GCn.Execute "Update Sp_Stock Set Rate=V_Rate, Rate2=V_Rate, Amount=V_Rate*(Qty_Iss-Qty_Ret), Amount2=V_Rate*(Qty_Iss-Qty_Ret), TaxAmt=(V_Rate*(Qty_Iss-Qty_Ret))*TaxPer/100 Where Purpose='W'"
            Else
                GCn.Execute "Update Sp_Stock, Part Set Sp_Stock.V_Rate=" & cIIF("Part.MRP=0", "Sp_Stock.Rate", "Part.NDP") & "  Where Sp_Stock.Part_No=Part.Part_No  "
                GCn.Execute "Update Sp_Stock Set Rate=V_Rate, Rate2=V_Rate, Amount=V_Rate*(Qty_Iss-Qty_Ret), Amount2=V_Rate*(Qty_Iss-Qty_Ret), TaxAmt=(V_Rate*(Qty_Iss-Qty_Ret))*TaxPer/100 Where Purpose='W'"
            End If
        
            GCn.Execute "Update Log_Structure Set Remark='OK' Where Particular='" & mParticular & "'"
        End If
    'End If
    
    CreateVType "VEH", "JV", "BINV", "BODY BUILD INVOICE", CStr(Format(PubStartDate, "YYYY")), Val(Right(PubStartDate, 1) & "0000000"), PubDivCode
    
    
    GCn.Execute "Update Sp_Sale Set Sp_Sale.Sat_Yn = X.Sat_Yn " & _
                "From (Select DocId, " & cIIF("Max(SatPer) > 0", "1", "0") & " As Sat_Yn From Sp_Stock Where V_Type = 'SYSC' Group By DocId) As X  Where Sp_Sale.DocId = X.DocId"
    
    
    
    
    MsgBox "Table Structure Updation completed!" & vbCrLf & "Please reload programme", vbCritical, "Modify Table"
    End
errorbox:
    Set tmpDb = Nothing
    Set tmpDbFA = Nothing
    Set DB = Nothing
    If GCn.State = 0 Then GCn.Open
    If G_FaCn.State = 0 Then G_FaCn.Open
    If GCnFaV.State = 0 Then GCnFaV.Open
    If GCnFaS.State = 0 Then GCnFaS.Open
    If GCnFaW.State = 0 Then GCnFaW.Open
    MsgBox err.Description, vbInformation
    End
End Sub
'End Sub
'End Update

Private Sub AddNewField(Conn As ADODB.Connection, ByVal mTable As String, ByVal mColumn As String, ByVal mDataType As String, Optional ByVal mDefault_Value As String = "")
    If Conn.Execute("select Isnull(count(*),0) from sysColumns where id = object_id('" & mTable & "') and name in ('" & mColumn & "')").Fields(0).Value = 0 Then
        If mDefault_Value <> "" Then
            Conn.Execute ("ALTER TABLE " & mTable & " Add " & mColumn & " " & mDataType & " Default " & mDefault_Value)
            Conn.Execute ("Update " & mTable & " Set " & mColumn & "=" & mDefault_Value & " Where " & mColumn & " Is Null")
        Else
            Conn.Execute ("ALTER TABLE " & mTable & " Add " & mColumn & " " & mDataType)
        End If
    End If
End Sub


Sub Log_StructureCreate(mParticular As String)
    If GCn.Execute("Select Count(*) From Log_Structure Where Particular = '" & mParticular & "'").Fields(0).Value = 0 Then
        GCn.Execute "Insert Into Log_Structure(Particular) Values ('" & mParticular & "')"
    End If
End Sub


Sub CreateVType(mCategory As String, mNCat As String, mVType As String, mVDesc As String, mVPrefix As String, mStartSrlNo As Double, mDivCode As String)
Dim TmpRst As ADODB.Recordset
    Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='" & mVType & "'")
    If TmpRst.RecordCount = 0 Then
        G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                      "('" & mCategory & "','" & mNCat & "','" & mVType & "','" & mVDesc & "','" & Replace(mVDesc, " ", "") & "','" & mVType & "','Automatic','N','Y','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
    End If
    Set TmpRst = G_FaCn.Execute("Select * From Voucher_Prefix Where V_Type='" & mVType & "' And Date_From=" & ConvertDate(PubStartDate) & " And Date_To=" & ConvertDate(PubEndDate) & "")
    If TmpRst.RecordCount = 0 Then
        G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('" & mVType & "'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'" & mVPrefix & "'," & mStartSrlNo & ",'" & mDivCode & "')")
    End If
End Sub



'Nra Updation
Public Sub UpdateQry(DB As DAO.Database, QryName As String, NQry As String)
Dim QDF As QueryDef
    Set QDF = DB.QueryDefs(QryName)
    QDF.SQL = NQry
Myexit:

End Sub
'End Update

Private Sub DelAllTransactions()
If pubUName <> "SA" Then
    MsgBox "Permission Denied !", vbCritical, "Delete Transactions !"
    Exit Sub
End If
If MsgBox("Are You Sure To Delete Transactions from Database ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Transactions !") = vbNo Then Exit Sub
If MsgBox("Records will be removed permanentaly, Continue ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Transactions-2 !") = vbNo Then Exit Sub

On Error GoTo errorbox
Dim DataPath$, DataPathFa$
Dim DB As DAO.Database, I As Integer, mTrans As Boolean
Dim fob As New FileSystemObject

DataPath = Pub_DataPath & "\Auto_" & PubCenCompCode & "\Automan.mdb"
DataPathFa = PubFADataPath
'dbBinInt,dbBoolean,dbByte,dbDate,dbDecimal,dbDouble,dbFloat
'dbInteger,dbMemo,dbNumeric,dbSingle,dbText,dbTime,dbTimestamp

    'Checking for Exclusive mode
'    GCn.Close
'    G_FaCn.Close
'    GCnFaV.Close
'    GCnFaS.Close
'    GCnFaW.Close
'
'    Dim tmpDb As DAO.Database
'    Set tmpDb = OpenDatabase(DataPath, True, False, ";pwd=dtman")
'
'    Dim tmpDbFA As DAO.Database
'    Set tmpDbFA = OpenDatabase(DataPathFa, True, False)
'
'    Close #1
'    If Fob.FileExists("C:\RepPrint.Txt") = False Then
'        Fob.CreateTextFile ("C:\TableName.Txt")
'    End If
'    Close #1
'    Open "C:\TableName.Txt" For Output As #1
'    For I = 0 To tmpDb.TableDefs.Count - 1
'        Print #1, tmpDb.TableDefs(I).Name
'    Next
'    Print #1, "FA Tables"
'    For I = 0 To tmpDbFA.TableDefs.Count - 1
'        Print #1, tmpDbFA.TableDefs(I).Name
'    Next
'
'    Close #1
'    Set tmpDb = Nothing
'    Set tmpDbFA = Nothing
'    GCn.Open
'    G_FaCn.Open
'    GCnFaV.Open
'    GCnFaS.Open
'    GCnFaW.Open
    '*******************
GCn.BeginTrans
G_FaCn.BeginTrans

GCn.Execute ("Delete from Estimate")
GCn.Execute ("Delete from Estimate1")
GCn.Execute ("Delete from Indent")
GCn.Execute ("Delete from Job_Booking")
GCn.Execute ("Delete from Job_Card")
GCn.Execute ("Delete from Job_Card2")
GCn.Execute ("Delete from Job_Demand")
GCn.Execute ("Delete from Job_GatePass")
GCn.Execute ("Delete from Job_Inspection")
GCn.Execute ("Delete from Job_Inspection2")
GCn.Execute ("Delete from Job_Lab")
GCn.Execute ("Delete from Job_Lab2")
GCn.Execute ("Delete from Job_WarBill")
GCn.Execute ("Delete from Job_Warr1")
GCn.Execute ("Delete from Job_Warr2")
GCn.Execute ("Delete from OverTime")
GCn.Execute ("Delete from Part_Import")
GCn.Execute ("Delete from Part_PriceList")
GCn.Execute ("Delete from PartList_New")
GCn.Execute ("Delete from Rect")
GCn.Execute ("Delete from SP_Order")
GCn.Execute ("Delete from SP_Order1")
GCn.Execute ("Delete from SP_Purch")
GCn.Execute ("Delete from SP_Sale")
GCn.Execute ("Delete from SP_Stock")

'Fa Tables
G_FaCn.Execute ("Delete from Ledger")
G_FaCn.Execute ("Delete from LedgerAdj")
G_FaCn.Execute ("Delete from LedgerM")
G_FaCn.Execute ("Update VehBill_Counter set Start_Srl_No=0")
G_FaCn.Execute ("Update Voucher_Prefix set Start_Srl_No=0")

GCn.CommitTrans
G_FaCn.CommitTrans

    MsgBox "All Transactions from Database removed sucessfully" & vbCrLf & "Please reload programme", vbCritical, "Delete Transactions"
    End
errorbox:
'    Set tmpDb = Nothing
'    Set tmpDbFA = Nothing
    Set DB = Nothing
    If mTrans Then GCn.RollbackTrans
    If mTrans Then G_FaCn.RollbackTrans
    MsgBox err.Description, vbInformation
    End

End Sub
Private Sub RemoveNulls()
On Error GoTo errorbox
Dim Conn As ADODB.Connection, aaaa As String, j As Integer, DataPath$, tmpDb1 As DAO.Database, N As Integer, I As Integer
Dim tmprs As DAO.Recordset, TDF As DAO.TableDef, FLD As DAO.Field
For I = 1 To 2
    If I = 1 Then
        DataPath = Pub_DataPath & "\Auto_01\Automan.Mdb"
    Else
        DataPath = PubFADataPath
    End If
    Set tmpDb1 = DBEngine.OpenDatabase(DataPath, False, False, ";pwd=dtman")
    For j = 0 To tmpDb1.TableDefs.Count - 1
        If InStr(1, UCase(tmpDb1.TableDefs(j).Name), "SYS") = 0 And InStr(1, UCase(tmpDb1.TableDefs(j).Name), "~") = 0 And InStr(1, tmpDb1.TableDefs(j).Name, "Paste Errors") = 0 Then
            Set TDF = tmpDb1.TableDefs(tmpDb1.TableDefs(j).Name)
            For N = 0 To TDF.Fields.Count - 1
                Set FLD = TDF.Fields(N)
                If (FLD.Type = dbText Or FLD.Type = 0) Then
                    If FLD.DefaultValue = "" Then FLD.DefaultValue = """"""
                    tmpDb1.Execute "Update " & TDF.Name & " set " & TDF.Fields(N).Name & " = '' where " & TDF.Fields(N).Name & " is null"
                ElseIf (FLD.Type = dbNumeric Or FLD.Type = dbSingle Or FLD.Type = dbDouble Or FLD.Type = dbLong) Then
                    If FLD.DefaultValue = "" Then FLD.DefaultValue = 0
                    tmpDb1.Execute "Update " & TDF.Name & " set " & TDF.Fields(N).Name & " = 0 where " & TDF.Fields(N).Name & " is null"
                End If
errorbox:
            Next
        End If
    Next
    Set tmpDb1 = Nothing
    Set tmprs = Nothing
Next

''errorbox:
''    If err.NUMBER = 3343 Then
''        MsgBox "Database is currupt or not a Database file"
''    Else
''        MsgBox err.Description, vbInformation
''    End If
End Sub
Public Sub DataBackup()
Dim oZip As CGZipFiles, ZipDrive As String, ZipFile$
Dim oUnZip As CGUnzipFiles
Dim DB1 As DAO.Database
On Error GoTo DispErr


ZipFile = Trim("Auto_" & PubCenCompCode & "-" & Format(Now(), "dd") & Format(Now(), "mm") & Format(Now(), "yy"))
  'Database Backup
            '****
           ' Dim DB1 As DAO.Database
            'Checking for Exclusive mode
            If GCn.State = 1 Then GCn.Close
            Set DB1 = OpenDatabase(Pub_DataPath & "\Auto_" & PubCenCompCode & "\Automan.mdb", True, False, ";pwd=dtman")
            Set DB1 = Nothing
            GCn.Open
            '****
            Picture1.Visible = True
            Label1.CAPTION = "Please wait (" & ZipFile & ") Backup being Process..."
            Set oZip = New CGZipFiles
            With oZip
                ' Give Zip File a Name / Path
                .ZipFileName = PubBkpPath & "\" & ZipFile & ".ZIP"
                ' Are we updating a Zip File ?
                ' - This doesn't seem to work -
                .UpdatingZip = False ' ensures a new zip is created
                .RootDirectory = mID(Pub_DataPath, 1, Len(Pub_DataPath) - 4)
                .AddFile Pub_DataPath & "\Auto_" & PubCenCompCode & "\Automan.mdb"
                
                Set GRs = GCn.Execute("select FADataPath from AssociatedFirms")
                
                While Not GRs.EOF
                    .AddFile Pub_DataPath & "\" & GRs!FADataPath & "\FaData.mdb"
                    GRs.MoveNext
                Wend
                ' Make the zip file & display any errors
                If .MakeZipFile <> 0 Then
                    MsgBox .GetLastMessage ' any errors
                    Picture1.Visible = False
                    Exit Sub
                Else
                    Picture1.Visible = False
                    MsgBox ZipFile & ".ZIP Created Successfully", vbInformation, "Information"
                End If
            End With
            Set oZip = Nothing
Exit Sub
DispErr:
    MsgBox err.Description & " In DataBackUp Procedure"
    
End Sub

Private Function CreateNewTable_sql()
On Error Resume Next

GCn.Execute " CREATE TABLE dbo.Deprecation_itemMaster(Code         NVARCHAR (5),ShortName  NVARCHAR (2) ,Description  NVARCHAR (50),Dep_per FLOAT , " & _
    "Site_Code    NVARCHAR (2) CONSTRAINT DF_Deprecation_itemMaster_Site_Code DEFAULT ('') NOT NULL,U_Name       NVARCHAR (15) CONSTRAINT DF_Deprecation_itemMaster_U_Name DEFAULT ('') NOT NULL, U_EntDt      SMALLDATETIME NOT NULL, U_AE         NVARCHAR (1) CONSTRAINT DF_Deprecation_itemMaster_U_AE DEFAULT ('') NOT NULL, " & _
    " CONSTRAINT PK_Deprecation_itemMaster PRIMARY KEY (Code),CONSTRAINT IX_Deprecation_itemMaster UNIQUE (Description) )"
        
        
        
GCn.Execute "CREATE TABLE dbo.Deprecation_Master( Code         NVARCHAR (5) ,Dep_Month  FLOAT ,Dep_per FLOAT ,Site_Code    NVARCHAR (2) CONSTRAINT DF_Deprecation_Master_Site_Code DEFAULT ('') NOT NULL," & _
    " U_Name       NVARCHAR (15) CONSTRAINT DF_Deprecation_Master_U_Name DEFAULT ('') NOT NULL, U_EntDt      SMALLDATETIME NOT NULL, U_AE         NVARCHAR (1) CONSTRAINT DF_Deprecation_Master_U_AE DEFAULT ('') NOT NULL, " & _
    " CONSTRAINT PK_Deprecation_Master PRIMARY KEY (Code), CONSTRAINT IX_Deprecation_Master UNIQUE (Dep_Month))"

End Function
Private Function CreateNewTable()
On Error GoTo LblErr
    Dim Cat As New ADOX.Catalog, PubDatamanFa As New DMFa.ClsFa, Rough As Boolean
    Cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=dtman;Data Source=" & Pub_DataPath & "\" & PubCenDataPath & "\Automan.Mdb"
    PubDatamanFa.FaBackEnd = PubBackEnd

'Creating RepeatJob Table



    'Creating Log_Structure Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "Log_Structure")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Log_Structure", "Particular", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Log_Structure", "Remark", adWChar, 50, False, "", True)

    'Creating A/C MERZE Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "AcMERGE")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "AcMERGE", "Table", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "AcMERGE", "FldName", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "AcMERGE", "Database", adWChar, 50, False, "", True)

    Rough = PubDatamanFa.SanFaCreateTable(Cat, "RepeatJob")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Job_DocId", adWChar, 21, True, "", False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "JobNo", adDouble, 8, True, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Job_Date", adDate, , True, "", False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "RegNo", adWChar, 14, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Mech", adWChar, 40, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "PrvJobNo", adDouble, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "PrvJobDate", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Comp_Name", adWChar, 40, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Batch_Code", adWChar, 25, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Make", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Ord_Placed", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Cust_Informed", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Imp_Date1", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Imp_Date2", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Imp_Date3", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Imp_Date4", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Imp_Date5", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "Imp_Date6", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "RepeatJob", "U_AE", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreateIndex(Cat, "RepeatJob", "PrimaryKey", True, True, adIndexNullsDisallow, "Job_DocId")
    
    
    'Creating CustFeedBack Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "CustFeedback")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Job_DocId", adWChar, 21, True, "", False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Parameter1", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Parameter2", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Parameter3", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Parameter4", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Parameter5", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Parameter6", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Parameter7", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Parameter8", adDouble, 2, False, 0, False)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Point1", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Point2", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Point3", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Point4", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Point5", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Point6", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Point7", adDouble, 2, False, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "Point8", adDouble, 2, False, 0, False)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "FeedbackStat", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "NxtVisit", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "CompNature", adWChar, 50, False, "", True)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustFeedback", "U_AE", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreateIndex(Cat, "CustFeedback", "PrimaryKey", True, True, adIndexNullsDisallow, "Job_DocId")
    
    'Creating CustInfo Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "CustInfo")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Cust_Code", adDouble, 8, True, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Cust_Name", adWChar, 40, True, "", False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Add1", adWChar, 40, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Add2", adWChar, 40, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Add3", adWChar, 40, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "City", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Pin", adDouble, 6, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "State", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "PhoneO", adWChar, 30, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "PhoneR", adWChar, 30, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Mobile", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Fax", adWChar, 30, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "EMail", adWChar, 30, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "PCY_N", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Sex", adWChar, 6, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "MState", adWChar, 6, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "NoOfChild", adDouble, 2, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "BDate", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "WDate", adDate, , False, "", True)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "MTounge", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Ugrd", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Grd", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Pgrd", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "OthEdu", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Salaried", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Businessman", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Retired", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Student", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Housewife", adInteger, 1, False, 0, True)
    
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "OthMod1", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "OthMod2", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "ModPur1Dt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "ModPur2Dt", adDate, , False, "", True)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Self", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Driver", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Both1", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "AuthWorkshop", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "OtherWorkshop", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "AvgDist1", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "AvgDist2", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "AvgDist3", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "AvgDist4", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Citibank", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "HDFC", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "ICICI", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "HSBC", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "SBI", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Diners", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Stanchart", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Amex", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "BOB", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "OtherBank", adWChar, 20, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "IA", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Jet", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Sahara", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "OtherAir", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "TravelRelated", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "LifeRelated", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Entertainment", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Parsonal", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "CarRelated", adInteger, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "OtherInterest", adWChar, 20, False, "", True)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "DInv_No", adDouble, 8, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "DInv_Date", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Model", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Varient", adWChar, 10, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "RegNo", adWChar, 14, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Color", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "Chassis", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "CEligibility", adWChar, 200, False, "", True)
    
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "CustInfo", "U_AE", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreateIndex(Cat, "CustInfo", "PrimaryKey", True, True, adIndexNullsDisallow, "Cust_Code")
    
    
    'Creating VEH_CustConcn Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "VEH_CustConcn")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Sl_No", adDouble, 8, True, 0, False)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "VDate", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Cust_Name", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Cust_add1", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Cust_add2", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Tel_R", adDouble, 10, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Tel_O", adDouble, 10, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Mobile", adDouble, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "VehModel", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Concern", adWChar, 250, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Received_By", adWChar, 25, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Imm_Act_Taken1", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Imm_Act_Taken2", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Root_Cause_Analysis1", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Root_Cause_Analysis2", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Pre_Action1", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Pre_Action2", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "Remarks", adWChar, 250, False, "", True)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_CustConcn", "U_AE", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreateIndex(Cat, "VEH_CustConcn", "PrimaryKey", True, True, adIndexNullsDisallow, "Sl_No")
    
    'Creating VEH_MargineStatistics Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "VEH_Margin")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "Sl_No", adDouble, 8, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "CustName", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "SPerson", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "FinCode", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "FinAmt", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "SalInv", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "Model", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "VehPriceAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "VehPriceAR", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "RegChrgAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "RegChrgAR", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "SPNoChrgAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "SPNoChrgAR", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AffiChrgAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AffiChrgAR", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "InsuChrgAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "InsuChrgAR", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "EWPChrgAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "EWPChrgAR", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AccessoriesAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AccessoriesAR", adDouble, 6, False, 0, True)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "TotalAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "TotalAR", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "ExcessAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "ExcessAR", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "ShortageAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "ShortageAR", adDouble, 6, False, 0, True)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBill1", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBill2", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBill3", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBill1STot", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBill1Tot", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBill2STot", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBill2Tot", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBill3STot", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBill3Tot", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBillSTotal", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBillTotal", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBillBalSTot", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjBillBalTot", adDouble, 6, False, 0, True)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjVehCostAA", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjVehCostAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjRegChrgAA", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "AdjRegChrgAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "CashPayAA", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "CashPayAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "FitmentAA", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "FitmentAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "DMAAA", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "DMAAP", adDouble, 6, False, 0, True)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "DisTotalAA", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "DisTotalAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "DisBalAA", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "DisBalAP", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "MrgVeh", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "MrgRegn", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "CorpInc", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "FinIncentive", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "MrgSplNo", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "PurIncentive", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "InventCost", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "Petrol", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "Discount", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "Misc", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "RTO", adDouble, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "NetMrg", adDouble, 6, False, 0, True)
    
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "VEH_Margin", "U_AE", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreateIndex(Cat, "VEH_Margin", "PrimaryKey", True, True, adIndexNullsDisallow, "Sl_No")
    
    'Creating Warrenty Complaint Master Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "WarrCompMast")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrCompMast", "Code", adWChar, 3, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrCompMast", "Description", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrCompMast", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrCompMast", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrCompMast", "U_AE", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreateIndex(Cat, "WarrCompMast", "PrimaryKey", True, True, adIndexNullsDisallow, "Code")
    
    'Creating Warranty Failure Master Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "WarrFailMast")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrFailMast", "Code", adWChar, 7, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrFailMast", "Description", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrFailMast", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrFailMast", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrFailMast", "U_AE", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreateIndex(Cat, "WarrFailMast", "PrimaryKey", True, True, adIndexNullsDisallow, "Code")
    
    'Creating Warranty Make Master Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "WarrMakeMast")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrMakeMast", "Code", adWChar, 6, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrMakeMast", "Description", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrMakeMast", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrMakeMast", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrMakeMast", "U_AE", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreateIndex(Cat, "WarrMakeMast", "PrimaryKey", True, True, adIndexNullsDisallow, "Code")
    
    'Creating Warranty Job Master Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "WarrJobMast")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrJobMast", "Code", adWChar, 6, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrJobMast", "Description", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrJobMast", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrJobMast", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "WarrJobMast", "U_AE", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreateIndex(Cat, "WarrJobMast", "PrimaryKey", True, True, adIndexNullsDisallow, "Code")
    
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "DeleteLog")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DeleteLog", "DocId", adWChar, 21, True, "")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DeleteLog", "CancelCount", adDouble, 50, False, 0)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DeleteLog", "Bill_Amt", adDouble, 10, False, 0)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DeleteLog", "User_Name", adWChar, 50, False, "")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DeleteLog", "Del_Date", adDate, 50, False, "")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DeleteLog", "Del_Time", adWChar, 20, False, "")
        
        
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "Lab_Trouble")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Lab_Trouble", "Srl", adDouble, 8, True, 0)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Lab_Trouble", "Lab_Code", adWChar, 8, True, "")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Lab_Trouble", "CCCode", adWChar, 6, True, "")
        
    
    
    
    'Creating Subvention Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "Subvention")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Subvention", "SchemeNo", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Subvention", "FromDate", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Subvention", "ToDate", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Subvention", "ModelGroup", adWChar, 5, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Subvention", "Model", adWChar, 24, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Subvention", "DealerContribution", adDouble, 10, False, 0)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Subvention", "TataContribution", adDouble, 10, False, 0)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Subvention", "TotalSubvention", adDouble, 10, False, 0)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Subvention", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Subvention", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Subvention", "U_AE", adWChar, 1, False, "", True)

    
    
    'Creating OffTake Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "OffTake")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake", "Code", adSingle, , False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake", "SchemeNo", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake", "FromDate", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake", "ToDate", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake", "Qty", adDouble, 10, False, 0)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake", "Amount", adDouble, 10, False, 0)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake", "U_AE", adWChar, 1, False, "", True)
    
    
    
    'Creating OffTake Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "OffTake1")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake1", "Code", adSingle, , False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake1", "SrlNo", adInteger, , False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "OffTake1", "ModelGroup", adWChar, 5, False, "", True)
    
    
    
    'Creating User_Site Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "User_Site")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "User_Site", "Site_Code", adWChar, 1, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "User_Site", "User_Name", adWChar, 10, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "User_Site", "Comp_Code", adWChar, 2, False, 0, True)
    
    
    'Creating DmsSubGroup Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "DmsSubGroup")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "DmsSubCode", adWChar, 15, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "Name", adWChar, 50, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "Add1", adWChar, 50, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "Add2", adWChar, 50, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "City", adWChar, 50, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "PinCode", adWChar, 6, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "State", adWChar, 2, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "Phone", adWChar, 35, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "Fax", adWChar, 24, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "Email", adWChar, 50, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "Group", adWChar, 20, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "Division", adWChar, 40, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSubGroup", "AutomanSubCode", adWChar, 8, False, 0, True)
    
    
    
    ''''''DmsEnviro''''For CrmDmsImport
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "DmsEnviro")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "WsDebtorGroupCode", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "SprDebtorGroupCode", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "VehDebtorGroupCode", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "SprCreditorGroupCode", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "VehCreditorGroupCode", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "SprSaleAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "VehSaleAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "LubeSaleAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "VatAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "WSCashAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "SprCashAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "VehCashAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "LabourAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "ServTaxAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "LocalStateName", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "SprBankAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "VehBankAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "SprPurchaseAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "VehPurchaseAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "CstAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsEnviro", "ROffAc", adWChar, 8, False, "", True)
        
    
    
    
    ''''''DmsSite''''For CrmDmsImport
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "DmsSite")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSite", "AutomanSite", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSite", "AutomanDivision", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSite", "DmsDivision", adWChar, 40, False, "", True)
    
    
    ''''''DmsErrLog''''For CrmDmsImport
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "DmsErrLog")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsErrLog", "Cat", adWChar, 25, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsErrLog", "Key", adWChar, 30, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsErrLog", "Narration", adWChar, 255, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsErrLog", "U_EntDt", adDate, , False, "", True)
    
    
    
    ''''''''''DmsData'''''''''''''
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "DmsData")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "DocId", adWChar, 21, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "VType", adWChar, 5, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "VDate", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "VNo", adDouble, 8, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "SubCode", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "L_C", adWChar, 10, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "Amount", adDouble, 8, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "SprAmt", adDouble, 8, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "LubeAmt", adDouble, 8, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "TaxAmt", adDouble, 8, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "LabAmount", adDouble, 8, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "Lab_DocId", adWChar, 21, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "SrvTax", adDouble, 8, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsData", "DmsRefNo", adWChar, 25, False, "", True)
    
    
    ''''''''''DmsBankAc'''''''''''''
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "DmsBankAc")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsBankAc", "AutomanBankCode", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsBankAc", "DmsBankCode", adWChar, 15, False, "", True)
    
    
    ''''''''''DmsSupplier'''''''''''''
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "DmsSupplier")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSupplier", "AutomanSupplierCode", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "DmsSupplier", "DmsSupplierCode", adWChar, 15, False, "", True)
    
    
    ''''''''''Budget_Exp'''''''''''''
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "Budget_Exp")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Budget_Exp", "ExpAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Budget_Exp", "Site_Code", adWChar, 2, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Budget_Exp", "VDate", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Budget_Exp", "Month", adWChar, 2, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Budget_Exp", "Amount", adDouble, 8, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Budget_Exp", "U_Name", adWChar, 10, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Budget_Exp", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Budget_Exp", "U_AE", adWChar, 1, False, "", True)
    
    
    ''''''''''Exp_Emp'''''''''''''
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "Exp_Emp")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp", "DocId", adWChar, 21, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp", "V_Type", adWChar, 5, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp", "V_Prefix", adWChar, 5, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp", "V_No", adDouble, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp", "V_Date", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp", "ExpAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp", "CashBankAc", adWChar, 8, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp", "Amount", adDouble, 8, False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp", "Narration", adWChar, 255, False, "", True)
    
    ''''''''''Exp_Emp1'''''''''''''
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "Exp_Emp1")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp1", "DocId", adWChar, 21, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp1", "Srl", adInteger, , False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp1", "Emp_Code", adWChar, 5, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Exp_Emp1", "Amount", adDouble, 8, False, 0, True)
    
    
    'Creating BodyBuilder Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "BodyBuilder")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyBuilder", "BodyBuilderCode", adWChar, 5, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyBuilder", "BodyBuilderDesc", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyBuilder", "Add1", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyBuilder", "Add2", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyBuilder", "CityCode", adWChar, 5, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyBuilder", "Contact", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyBuilder", "Site_Code", adWChar, 2, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyBuilder", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyBuilder", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyBuilder", "U_AE", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreateIndex(Cat, "VEH_CustConcn", "PrimaryKey", True, True, adIndexNullsDisallow, "BodyBuilderCode")
    
    'Creating BodyBuilder Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "BodyType")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyType", "BodyTypeCode", adWChar, 5, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyType", "BodyTypeDesc", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyType", "Site_Code", adWChar, 2, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyType", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyType", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "BodyType", "U_AE", adWChar, 1, False, "", True)
    Rough = PubDatamanFa.SanFaCreateIndex(Cat, "BodyType", "PrimaryKey", True, True, adIndexNullsDisallow, "BodyTypeCode")
    
    
    'Creating Insurance Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "Insurance")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Insurance", "Code", adWChar, 5, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Insurance", "Name", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Insurance", "Add1", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Insurance", "Add2", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Insurance", "City", adWChar, 5, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Insurance", "ContactPerson", adWChar, 50, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Insurance", "U_Name", adWChar, 15, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Insurance", "U_EntDt", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Insurance", "U_AE", adWChar, 1, False, "", True)
        
    
    'Creating Rect1 Table
    Rough = PubDatamanFa.SanFaCreateTable(Cat, "Rect1")
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Rect1", "DocID", adWChar, 21, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Rect1", "Srl", adInteger, , False, 0, True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Rect1", "ChqNo", adWChar, 20, False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Rect1", "ChqDate", adDate, , False, "", True)
    Rough = PubDatamanFa.SanFaCreatFields(Cat, "Rect1", "ChqAmt", adDouble, 8, False, 0, True)
      

    
    Set PubDatamanFa = Nothing
    Set Cat = Nothing
' Creating Veh_Order1
    Exit Function
LblErr:
'    If Err.NUMBER = "-2147217900" Then
'        CreateNewTable = 1
'    Else
        CheckError
'        CreateNewTable = 0
'    End If
End Function





Public Function UpdtCurrBalances()

Dim DataPath$
Dim I As Integer, GSQL1$
On Error Resume Next
'If MsgBox("Are You Sure To Update Current Balance of A/c ? ", vbYesNo + vbCritical + vbDefaultButton2, "Update Current Stock !") = vbYes Then
    
    Dim Rst As ADODB.Recordset, mTrans As Boolean
    
    GCn.BeginTrans
    G_FaCn.BeginTrans
    mTrans = True
    G_FaCn.Execute ("update SubGroup set Curr_Bal=0")
    GCn.Execute ("update SubGroup set Curr_Bal=0")
    
    GSQL = "SELECT Ledger.SubCode,SUM(AmtCr-AmtDr) as CBal " & _
            "FROM Ledger left join SubGroup SG on SG.SubCOde=Ledger.SubCode " & _
            "group by Ledger.subcode,Name"
    Set Rst = G_FaCn.Execute(GSQL)
    If Rst.RecordCount > 0 Then
        Do While Rst.EOF = False
            GCn.Execute ("Update SubGroup set Curr_Bal=" & Rst!CBal & " where SubCode='" & Rst!SubCode & "'")
            G_FaCn.Execute ("Update SubGroup set Curr_Bal=" & Rst!CBal & " where SubCode='" & Rst!SubCode & "'")
            Rst.MoveNext
        Loop
    End If
    DataPath = Pub_DataPath & "\" & PubCenDataPath & "\Automan.mdb;pwd=dtman"
    Set Rst = G_FaCn.Execute("select SG.SubCode from SubGroup as SG where SubCode not in (Select SubCode from [" & DataPath & "].SubGroup)")
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            GCn.Execute ("Delete From SubGroup where Subcode='" & Rst!SubCode & "'")
            GCn.Execute ("INSERT INTO SUBGROUP SELECT * FROM [" & PubFADataPath & "].SUBGROUP WHERE SUBCODE = '" & Rst!SubCode & "'")
            Rst.MoveNext
        Loop
        GCn.Execute ("Drop Table SubGroupAlias")
        GCn.Execute ("Select SubGroup.* into SubGroupAlias from SubGroup")
    End If
    GCn.CommitTrans
    G_FaCn.CommitTrans
    mTrans = False
    Set Rst = Nothing

    
    
    GCn.BeginTrans
    '***
    GCn.Execute ("update Part Set Cur_TP_Stk=0,Cur_TB_Stk=0,Cur_MRP_TPStk=0,Cur_MRP_TBStk=0 where Div_Code='" & PubDivCode & "'")
    '***
    GSQL = "Select Part_No,sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=0 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & "" & _
        " Group By Part_No"
        
    GSQL1 = "Select Part_No,sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=0 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO'" & _
        " Group By Part_No"
            
    GSQL = GSQL & " Union All " & GSQL1
    
    GSQL1 = "Select Part_No,Sum(BalQty) as Bal_Qty From (" & GSQL & ") Group By Part_No"
    
    Set Rst = GCn.Execute(GSQL1)
    
    If Rst.RecordCount > 0 Then
        For I = 1 To Rst.RecordCount
            GCn.Execute ("update Part Set Cur_TP_Stk=Cur_TP_Stk+" & Rst!Bal_Qty & " where Part_No='" & Rst!Part_No & "' and Div_Code='" & PubDivCode & "'")
            Rst.MoveNext
        Next
    End If
    
    GSQL = "Select Part_No,sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=1 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & " " & _
        " Group By Part_No"
        
    GSQL1 = "Select Part_No,sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=1 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO'" & _
        " Group By Part_No"
            
    GSQL = GSQL & " Union All " & GSQL1
    
    GSQL1 = "Select Part_No,Sum(BalQty)as Bal_Qty From (" & GSQL & ") Group By Part_No"
    
    Set Rst = GCn.Execute(GSQL1)
    If Rst.RecordCount > 0 Then
        For I = 1 To Rst.RecordCount
            GCn.Execute ("update Part Set Cur_TB_Stk=Cur_TB_Stk+" & Rst!Bal_Qty & " where Part_No='" & Rst!Part_No & "' and Div_Code='" & PubDivCode & "'")
            Rst.MoveNext
        Next
    End If
    
    GSQL = "Select Part_No,sum((Qty_Rec))-sum((Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=0 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & "" & _
        " Group By Part_No"
    
    GSQL1 = "Select Part_No,sum((Qty_Rec))-sum((Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=0 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO'" & _
        " Group By Part_No"
            
    GSQL = GSQL & " Union All " & GSQL1
    
    GSQL1 = "Select Part_No,Sum(BalQty) as Bal_Qty From (" & GSQL & ") Group By Part_No"
    
    Set Rst = GCn.Execute(GSQL1)
    If Rst.RecordCount > 0 Then
        For I = 1 To Rst.RecordCount
            GCn.Execute ("update Part Set Cur_MRP_TPStk=Cur_MRP_TPStk+" & Rst!Bal_Qty & " where Part_No='" & Rst!Part_No & "' and Div_Code='" & PubDivCode & "'")
            Rst.MoveNext
        Next
    End If
    GSQL = "Select Part_No,sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=1 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & " " & _
        " Group By Part_No"
    
    GSQL1 = "Select Part_No,sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=1 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO'" & _
        " Group By Part_No"
            
    GSQL = GSQL & " Union All " & GSQL1
    
    GSQL1 = "Select Part_No,Sum(BalQty) as Bal_Qty From (" & GSQL & ") Group By Part_No"
    
    Set Rst = GCn.Execute(GSQL1)
    If Rst.RecordCount > 0 Then
        For I = 1 To Rst.RecordCount
            GCn.Execute ("update Part Set Cur_MRP_TBStk=Cur_MRP_TBStk+" & Rst!Bal_Qty & " where Part_No='" & Rst!Part_No & "' and Div_Code='" & PubDivCode & "'")
            Rst.MoveNext
        Next
    End If
    GCn.CommitTrans
    RsPart.Requery
    Set Rst = Nothing

ELoop:
If mTrans Then GCn.RollbackTrans: G_FaCn.RollbackTrans

End Function


Private Sub X_Val1(ByRef Temp06 As ADODB.Recordset, ByRef TRec1 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec1.RecordCount <= 0 Or TRec1.EOF = True Or TRec1.BOF = True Then
        xRate = 0
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
        Exit Sub
    End If
    If xQty = TRec1!Qty Then
        TRec1.Fields("QTY") = 0
        TRec1.Update
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
        TRec1.MoveNext
    ElseIf xQty < TRec1!Qty Then
        TRec1.Fields("QTY") = TRec1!Qty - xQty
        TRec1.Update
        
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = xQty
                .Fields("Tb_Val") = xQty * xRate
                .Fields("Tb_BQty") = mOP_TB_QTY
                .Fields("Tb_BVal") = mOP_TB_VAL
                
                .Fields("Is_Tp") = 0
                .Fields("Tp_Val") = 0
                .Fields("Tp_BQty") = 0
                .Fields("Tp_BVal") = 0
                
                .Update
            End With
        End If
    ElseIf xQty > TRec1!Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec1.EOF
            If TRec1!Part_No <> RstPart!Part_No Then
                GoTo MyNextRecord
            End If
            If TRec1!Qty <= TQty Then
                TQty = TQty - TRec1!Qty
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TRec1!Qty
                mOP_TB_VAL = mOP_TB_VAL - (TRec1!Qty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                        .Fields("Part_Name") = mname
                        .Fields("Narr") = left(xNARR, 25)
                        .Fields("Inv_No") = mInv_No
                        .Fields("Inv_Date") = mInv_Date
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = TRec1!Qty
                        .Fields("Tb_Val") = TRec1!Qty * xRate
                        .Fields("Tb_BQty") = mOP_TB_QTY
                        .Fields("Tb_BVal") = mOP_TB_VAL
                        
                        .Fields("Is_Tp") = 0
                        .Fields("Tp_Val") = 0
                        .Fields("Tp_BQty") = 0
                        .Fields("Tp_BVal") = 0
                        .Update
                    End With
                    TRec1.Fields("QTY") = 0
                    TRec1.Update
                End If
            Else
                TRec1.Fields("QTY") = TRec1!Qty - TQty
                TRec1.Update
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = TQty
                        .Fields("Tb_Val") = TQty * xRate
                        .Fields("Tb_BQty") = mOP_TB_QTY
                        .Fields("Tb_BVal") = mOP_TB_VAL
                        
                        .Fields("Is_Tp") = 0
                        .Fields("Tp_Val") = 0
                        .Fields("Tp_BQty") = 0
                        .Fields("Tp_BVal") = 0
                        .Update
                    End With
                    TQty = 0
                    Exit Do
                End If
            End If
MyNextRecord:
            TRec1.MoveNext
            If TRec1.EOF = True And TQty <> 0 Then
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                If mPART_ADD = False Then
                    mPART_ADD = True
                    With Temp06
                        .AddNew
                        .Fields("Part_Name") = RstPart!Part_Name
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Job_Age") = "Y"
                        .Update
                    End With
                End If
                With Temp06
                    .AddNew
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Rate") = xRate
                    
                    .Fields("Is_Tb") = TQty
                    .Fields("Tb_Val") = TQty * xRate
                    .Fields("Tb_BQty") = mOP_TB_QTY
                    .Fields("Tb_BVal") = mOP_TB_VAL
                    
                    .Fields("Is_Tp") = 0
                    .Fields("Tp_Val") = 0
                    .Fields("Tp_BQty") = 0
                    .Fields("Tp_BVal") = 0
                    .Update
                End With
            
            End If
        Loop
    End If
End Sub

Private Sub X_Val2(ByRef Temp06 As ADODB.Recordset, ByRef TRec2 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec2.RecordCount <= 0 Or TRec2.EOF = True Or TRec2.BOF = True Then
        xRate = 0
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = PrinID(RstStock!DocID)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                
                .Update
            End With
        End If
        Exit Sub
    End If
    
    If xQty = TRec2!Qty Then
        TRec2.Fields("QTY") = 0
        TRec2.Update
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                
                .Update
            End With
        End If
        TRec2.MoveNext
    ElseIf xQty < TRec2!Qty Then
        TRec2.Fields("QTY") = TRec2!Qty - xQty
        TRec2.Update
        
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        
        If mTrf = False And xQty <> 0 Then
            If mPART_ADD = False Then
                mPART_ADD = True
                With Temp06
                    .AddNew
                    .Fields("Part_Name") = RstPart!Part_Name
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Job_Age") = "Y"
                    .Update
                End With
            End If
            With Temp06
                .AddNew
                .Fields("Date") = RstStock!V_DATE
                .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                .Fields("Part_Name") = mname
                .Fields("Narr") = xNARR
                .Fields("Inv_No") = mInv_No
                .Fields("Inv_Date") = mInv_Date
                .Fields("Part_No") = RstPart!Part_No
                .Fields("Rate") = xRate
                
                .Fields("Is_Tb") = 0
                .Fields("Tb_Val") = 0
                .Fields("Tb_BQty") = 0
                .Fields("Tb_BVal") = 0
                
                .Fields("Is_Tp") = xQty
                .Fields("Tp_Val") = xQty * xRate
                .Fields("Tp_BQty") = mOP_TP_QTY
                .Fields("Tp_BVal") = mOP_TP_VAL
                .Update
            End With
        End If
    ElseIf xQty > TRec2!Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec2.EOF
            If TRec2!Part_No <> RstPart!Part_No Then
                GoTo MyNextRecord
            End If
            If TRec2!Qty <= TQty Then
                TQty = TQty - TRec2!Qty
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TRec2!Qty
                mOP_TP_VAL = mOP_TP_VAL - (TRec2!Qty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Date") = RstStock!V_DATE
                        .Fields("V_No") = RstStock!V_Type + "-" + STR(RstStock!V_NO)
                        .Fields("Part_Name") = mname
                        .Fields("Narr") = xNARR
                        .Fields("Inv_No") = mInv_No
                        .Fields("Inv_Date") = mInv_Date
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = 0
                        .Fields("Tb_Val") = 0
                        .Fields("Tb_BQty") = 0
                        .Fields("Tb_BVal") = 0
                        
                        .Fields("Is_Tp") = TRec2!Qty
                        .Fields("Tp_Val") = TRec2!Qty * xRate
                        .Fields("Tp_BQty") = mOP_TP_QTY
                        .Fields("Tp_BVal") = mOP_TP_VAL
                        .Update
                    End With
                    TRec2.Fields("QTY") = 0
                    TRec2.Update
                End If
            Else
                TRec2.Fields("QTY") = TRec2!Qty - TQty
                TRec2.Update
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                If mTrf = False Then
                    If mPART_ADD = False Then
                        mPART_ADD = True
                        With Temp06
                            .AddNew
                            .Fields("Part_Name") = RstPart!Part_Name
                            .Fields("Part_No") = RstPart!Part_No
                            .Fields("Job_Age") = "Y"
                            .Update
                        End With
                    End If
                    With Temp06
                        .AddNew
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Rate") = xRate
                        
                        .Fields("Is_Tb") = 0
                        .Fields("Tb_Val") = 0
                        .Fields("Tb_BQty") = 0
                        .Fields("Tb_BVal") = 0
                        
                        .Fields("Is_Tp") = TQty
                        .Fields("Tp_Val") = TQty * xRate
                        .Fields("Tp_BQty") = mOP_TP_QTY
                        .Fields("Tp_BVal") = mOP_TP_VAL
                        .Update
                    End With
                    TQty = 0
                    Exit Do
                End If
            End If
MyNextRecord:
            TRec2.MoveNext
            If TRec2.EOF = True And TQty <> 0 Then
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                If mPART_ADD = False Then
                    mPART_ADD = True
                    With Temp06
                        .AddNew
                        .Fields("Part_Name") = RstPart!Part_Name
                        .Fields("Part_No") = RstPart!Part_No
                        .Fields("Job_Age") = "Y"
                        .Update
                    End With
                End If
                With Temp06
                    .AddNew
                    .Fields("Part_No") = RstPart!Part_No
                    .Fields("Rate") = xRate
                    
                    .Fields("Is_Tb") = 0
                    .Fields("Tb_Val") = 0
                    .Fields("Tb_BQty") = 0
                    .Fields("Tb_BVal") = 0
                    
                    .Fields("Is_Tp") = TQty
                    .Fields("Tp_Val") = TQty * xRate
                    .Fields("Tp_BQty") = mOP_TP_QTY
                    .Fields("Tp_BVal") = mOP_TP_VAL
                    .Update
                End With
            End If
        Loop
    End If
End Sub
Private Sub X_VAL11(ByRef TRec1 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
On Error GoTo ErrLoop
    If TRec1.RecordCount <= 0 Or TRec1.EOF = True Or TRec1.BOF = True Then
        If mOP_TB_VAL <> 0 And mOP_TB_QTY <> 0 Then
            xRate = Round(mOP_TB_VAL / mOP_TB_QTY, 3)
        Else
            xRate = 0
        End If
            mOP_TB_QTY = mOP_TB_QTY - xQty
            mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
            mIss_TB_Qty = mIss_TB_Qty + xQty
            mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
          Exit Sub
    End If
    If xQty = TRec1Qty Then
        TRec1Qty = 0
'        TRec1!Qty = 0
'        TRec1.Update
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        mIss_TB_Qty = mIss_TB_Qty + xQty
        mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
        TRec1.MoveNext
        If TRec1.EOF = False Then
            TRec1Qty = TRec1!Qty
        End If
'    ElseIf xQty < TRec1!Qty Then
    ElseIf xQty < TRec1Qty Then
        TRec1Qty = TRec1Qty - xQty
'        TRec1!Qty = TRec1!Qty - xQty
'        TRec1.Update
        xRate = TRec1!Rate
        mOP_TB_QTY = mOP_TB_QTY - xQty
        mOP_TB_VAL = mOP_TB_VAL - (xQty * xRate)
        mIss_TB_Qty = mIss_TB_Qty + xQty
        mIss_TB_Val = mIss_TB_Val + (xQty * xRate)
'    ElseIf xQty  > TRec1!Qty Then
    ElseIf xQty > TRec1Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec1.EOF
'            If TRec1!Qty <= TQty Then
            If TRec1Qty <= TQty Then
'                TQty = TQty - TRec1!Qty
                TQty = TQty - TRec1Qty
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TRec1Qty 'TRec1!Qty
                mOP_TB_VAL = mOP_TB_VAL - (TRec1Qty * xRate) '(TRec1!Qty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + (TRec1Qty) '(TRec1!Qty)
                mIss_TB_Val = mIss_TB_Val + (TRec1Qty * xRate) '(TRec1!Qty * xRate)
                TRec1Qty = 0
'                TRec1!Qty = 0
'                TRec1.Update
            Else
                TRec1Qty = TRec1Qty - TQty
'                TRec1!Qty = TRec1!Qty - TQty
'                TRec1.Update
                xRate = TRec1!Rate
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + TQty
                mIss_TB_Val = mIss_TB_Val + (TQty * xRate)
                TQty = 0
                Exit Do
            End If
            TRec1.MoveNext
            If TRec1.EOF = True And TQty <> 0 Then
                mOP_TB_QTY = mOP_TB_QTY - TQty
                mOP_TB_VAL = mOP_TB_VAL - (TQty * xRate)
                mIss_TB_Qty = mIss_TB_Qty + TQty
                mIss_TB_Val = mIss_TB_Val + (TQty * xRate)
            End If
            If TRec1.EOF = False Then
                TRec1Qty = TRec1!Qty
            End If
        Loop
    End If
ErrLoop:
     If err.NUMBER <> 0 Then CheckError
End Sub

Private Sub X_VAL22(ByRef TRec2 As ADODB.Recordset, xQty As Double, xRate As Double, Optional xNARR As String)
    If TRec2.RecordCount <= 0 Or TRec2.EOF = True Or TRec2.BOF = True Then
        xRate = 0
        
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
        Exit Sub
    End If
'    If xQty = TRec2!Qty Then
    If xQty = TRec2Qty Then
        TRec2Qty = 0
'        TRec2!Qty = 0
'        TRec2.Update
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
        TRec2.MoveNext
        If TRec2.EOF = False Then
            TRec2Qty = TRec2!Qty
        End If
'    ElseIf xQty < TRec2!Qty Then
    ElseIf xQty < TRec2Qty Then
        TRec2Qty = TRec2Qty - xQty
'        TRec2!Qty = TRec2!Qty - xQty
'        TRec2.Update
        xRate = TRec2!Rate
        mOP_TP_QTY = mOP_TP_QTY - xQty
        mOP_TP_VAL = mOP_TP_VAL - (xQty * xRate)
        mIss_TP_Qty = mIss_TP_Qty + xQty
        mIss_TP_Val = mIss_TP_Val + (xQty * xRate)
'    ElseIf xQty  > TRec2!Qty Then
    ElseIf xQty > TRec2Qty Then
        TQty = xQty
        Do While TQty <> 0 And Not TRec2.EOF
'            If TRec2!Qty <= TQty Then
            If TRec2Qty <= TQty Then
                TQty = TQty - TRec2Qty 'TRec2!Qty
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TRec2Qty 'TRec2!Qty
                mOP_TP_VAL = mOP_TP_VAL - (TRec2Qty * xRate)   '(TRec2!Qty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + (TRec2Qty)     '(TRec2!Qty)
                mIss_TP_Val = mIss_TP_Val + (TRec2Qty * xRate) '(TRec2!Qty * xRate)
                TRec2Qty = 0
'                TRec2!Qty = 0
'                TRec2.Update
            Else
                TRec2Qty = TRec2Qty - TQty
'                TRec2!Qty = TRec2!Qty - TQty
'                TRec2.Update
                xRate = TRec2!Rate
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + TQty
                mIss_TP_Val = mIss_TP_Val + (TQty * xRate)
                TQty = 0
                Exit Do
            End If
            TRec2.MoveNext
            If TRec2.EOF = True And TQty <> 0 Then
                mOP_TP_QTY = mOP_TP_QTY - TQty
                mOP_TP_VAL = mOP_TP_VAL - (TQty * xRate)
                mIss_TP_Qty = mIss_TP_Qty + TQty
                mIss_TP_Val = mIss_TP_Val + (TQty * xRate)
            End If
            If TRec2.EOF = False Then
                TRec2Qty = TRec2!Qty
            End If
        Loop
    End If
End Sub
Private Function Update_VRate()
Dim vrate As Double, XNO As Double, xNO1 As Double, xAmt As Double, xRate As Double, xRate1 As Double, xVRate As Double, mDisPer As Double, TempVal As Double, I As Integer, j As Integer, l As Integer, Cnt As Integer, Part_No$, RStkPart_No$
Dim StkArr(1000, 1000) As Double
Picture1.Visible = True
Label1.CAPTION = "Updating Reciept Side VRate....."
   'GCn.Execute ("update sp_stock set v_rate=0")
    mQry = "select Distinct SPStk.Part_No " & _
            "From " & _
            "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
            "Where Vt.StkTrn='+' Order By SPStk.Part_No Asc"
    Set RstStock3 = GCn.Execute(mQry)
    
    mQry = "select SPStk.DocID,SPStk.V_Type,SPStk.Part_No,SPStk.V_DATE,SPStk.MRP_YN,SPStk.Tax_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss,SPStk.V_Rate,Vt.StkTrn " & _
            "From " & _
            "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
            "Where Vt.StkTrn='+' Order By SPStk.V_Rate,SPStk.Part_No Asc"
    Set RstStock = GCn.Execute(mQry)
    
    For I = 1 To RstStock3.RecordCount
        Label1.CAPTION = "Updating Reciept Side V_Rate of Part : " & RstStock3!Part_No
        Label1.Refresh
        vrate = 0
        RstStock.Filter = ("Part_NO ='" & RstStock3!Part_No & "'")
          RstStock.Sort = "V_Date Asc"
        RstStock.MoveFirst
        For j = 1 To RstStock.RecordCount
          
           If RstStock!V_Rate = 0 And vrate = 0 Then
                If RstStock!MRP_YN = 1 Then
                   TempVal = GCn.Execute("Select MRP from Part where Part_No='" & RstStock!Part_No & "'").Fields(0).Value
                Else
                   TempVal = GCn.Execute("Select TP_SRate from Part where Part_No='" & RstStock!Part_No & "'").Fields(0).Value
                   mDisPer = IIf(IsNull(GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstStock!Part_No & "'").Fields(1).Value), 0, GCn.Execute("Select Part.Part_No,DF.PurcDisc_Per from Part Left Join Part_DiscFactor DF on Part.Disc_Factor = DF.DiscFac_Catg where Part.Part_No='" & RstStock!Part_No & "'").Fields(1).Value)
                   If mDisPer > 0 Then
                        TempVal = Round(TempVal - ((TempVal * mDisPer) / 100), 2)
                   End If
                End If
                GCn.Execute ("Update SP_Stock set v_Rate=" & TempVal & " where DocId='" & RstStock!DocID & "' and Part_No='" & RstStock!Part_No & "'")
           Else
                Part_No = RstStock!Part_No
                If Val(RstStock!V_Rate) = 0 Then
                    GCn.Execute ("Update SP_Stock set V_Rate=" & vrate & " where DocId='" & RstStock!DocID & "' and Part_No='" & RstStock!Part_No & "'")
                End If
                vrate = RstStock!V_Rate
                RstStock.MoveNext
           End If
            
           If Part_No <> RstStock3!Part_No Then
                vrate = 0
                Exit For
           End If
        
        Next
    RstStock3.MoveNext
    Next
    
Label1.CAPTION = "Updating Issue Side VRate....."
On Error Resume Next
mQry = "select SPStk.Part_No,SPStk.V_DATE,(SPStk.Qty_Rec + SPStk.Qty_Ret ) as Qty,SPStk.MRP_YN, " & cIIF("SPStk.V_Rate > 0", "SPStk.V_Rate", "SPStk.Rate") & " as Rate " & _
            "From " & _
            "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
            Condstr & CondPartNos
GSQL = mQry & " where Vt.StkTrn='+' Order By SPStk.Part_No,SPStk.V_Date,SPStk.DocID,SPStk.Srl_No"

Set RstStock = GCn.Execute(GSQL)
    
mQry = "select DISTINCT  SPStk.DocID,SPStk.Part_No,SPStk.MRP_YN,SPStk.Qty_Rec,(SPStk.Qty_Iss-SPStk.Qty_Ret) as Qty_Iss,SPStk.V_Rate ,Vt.StkTrn " & _
            "From " & _
            "SP_Stock as SPStk left Join " & FaTable("Voucher_Type") & " as VT on Vt.V_type=SPStk.V_type " & _
            "Where Vt.StkTrn <> '+' Order By Part_No Asc"
Set RstStock2 = GCn.Execute(mQry)

RStkPart_No = ""
RstStock2.MoveFirst
For I = 1 To RstStock2.RecordCount
Label1.CAPTION = "Updating Issue Side VRate of Part : " & RstStock2!Part_No
    'If RstStock2!V_Rate = 0 Then
        l = 1
        Part_No = RstStock2!Part_No
        If Part_No <> RStkPart_No Then
            Erase StkArr
            RstStock.Filter = ("Part_NO ='" & RstStock2!Part_No & "'")
            RstStock.Sort = "V_Date ASC"
            RstStock.MoveFirst
            RStkPart_No = RstStock!Part_No
            For j = 1 To RstStock.RecordCount
                StkArr(j, 0) = RstStock!Qty
                StkArr(j, 1) = RstStock!Rate
                RstStock.MoveNext
            Next
        End If
        
        XNO = 0: xNO1 = 0: xRate = 0: xVRate = 0: xAmt = 0
        XNO = RstStock2!Qty_Iss
       
        If XNO = StkArr(l, 0) Then
             xNO1 = 1
             xRate = (StkArr(l, 1))
             xAmt = xRate
        ElseIf XNO < StkArr(l, 0) Then
             xNO1 = 1
             xRate = (StkArr(l, 1))
             StkArr(l, 0) = (StkArr(l, 0)) - XNO
             xAmt = xRate
        ElseIf XNO > StkArr(l, 0) Then
             l = 1
             xNO1 = XNO
             Do While XNO > 0
             If StkArr(l, 0) > 0 Then
                If XNO >= StkArr(l, 0) Then
                    XNO = XNO - StkArr(l, 0)
                    xAmt = Round(xAmt + (StkArr(l, 0) * StkArr(l, 1)), 2)
                    l = l + 1
                Else
                    xAmt = Round(xAmt + (XNO * StkArr(l, 1)), 2)
                    XNO = XNO - StkArr(l, 0)
                End If
              ElseIf StkArr(l, 0) = 0 And XNO <> 0 Then
                xVRate = xRate1
                Exit Do
              End If
             Loop
        End If
        xVRate = Round(IIf(xAmt <> 0, (xAmt / xNO1), 0), 2)
        xRate1 = xVRate
        If RstStock2!V_Rate = 0 Then
            GCn.Execute ("Update SP_Stock set V_Rate=" & xVRate & " where DocId='" & RstStock2!DocID & "' and Part_No='" & RstStock2!Part_No & "' And V_Rate=0")
        End If
        
    'End If
RstStock2.MoveNext
Next
Picture1.Visible = False

End Function
Private Sub FilmMenu()
'*********  Reason to Design this *******
'it will auto genrate menu item in User_Module Form
'developed to get the menu names
'
'Make HelpContextID =1 for Top Parent Menu and HelpContextID = 2 for Report Menu
'Make Index = 0 for parent menus
'dont put move than two level in menu option
'tochange
Dim objctrl As Control
'On Error GoTo DispErr
On Error Resume Next


Dim ErrFound As Boolean
Dim I, j As Integer
Dim mHelpId, mTopMenuName As String
Dim mMainMenuName, mMenuName As String
Dim mMenuGroupID As String
Dim mMenuCaption, mMainMenuCaption As String
Dim mMenuIndex, mTopMenuCaption As String
Dim mModuleName$
Dim TempGrs As ADODB.Recordset
Dim RsTemp As ADODB.Recordset



ErrFound = False
I = 1: j = 1

Set TempGrs = G_CompCn.Execute("Select * from User_Module")
For Each objctrl In MDIForm1.Controls
    If TypeOf objctrl Is VB.Menu Then
            mHelpId = 0
            mMenuName = objctrl.Name & " " & objctrl.Index
            mMenuIndex = objctrl.Index
            mMenuCaption = Replace(objctrl.CAPTION, "&", "")
            
            mModuleName = ""
            If InStr(1, mMenuName, "Veh") > 0 Then
                mModuleName = "Vehicle"
            ElseIf InStr(1, mMenuName, "Spr") > 0 Or left(UCase(mMenuName), 10) = "MNSALESREP" Or left(UCase(mMenuName), 8) = "MNPURREP" Then
                mModuleName = "Spare"
            ElseIf InStr(1, mMenuName, "Wrk") > 0 Or left(UCase(mMenuName), 6) = "WKSHOP" Or left(UCase(mMenuName), 7) = "TLCOREP" Then
                mModuleName = "Workshop"
            ElseIf UCase(left(mMenuName, 4)) = "FAME" Or UCase(left(mMenuName, 5)) = "FAREP" Then
                mModuleName = "Account"
            End If
            
            
            If mMenuCaption = "Challan Allocation" Then MsgBox "xxxx"
            mHelpId = objctrl.HelpContextID
'            If mMenuCaption = "-" Or objctrl.Visible = False Then mMenuIndex = "": GoTo nxt1
            If objctrl.Visible = False Then mMenuIndex = "": GoTo nxt1
                        
            If mMenuIndex = "" Then ErrFound = True: GoTo NXT
            If G_CompCn.Execute("Select Count(*) From User_Module Where Name='" & mMenuCaption & "'").Fields(0) > 0 Then
                G_CompCn.Execute ("Update User_Module Set Menu_Name='" & mMenuName & "' Where Name='" & mMenuCaption & "'")
                ErrFound = True: GoTo nxt1
            End If
            
'''MsgBox mTopMenuName & vbCrLf & mMainMenuName & vbCrLf & mMenuCaption
            G_CompCn.Execute ("insert into User_Module(Module_Name,Form_Code,Name, Menu_Name) Values( '" & mModuleName & "','" & mMenuName & "','" & mMenuCaption & "', '" & mMenuName & "')")
            G_CompCn.Execute ("insert into User2(User_Name, Module_Name,Form_Code,Param_Str, Comp_Code, Div_Code) Values( 'SA','" & mModuleName & "','" & mMenuName & "','AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')")
            mMenuIndex = ""
            I = I + 1
NXT:

    If ErrFound = True Then
''''G_CompCn.Execute ("insert into User_Module(Code,Name,SrNo,flag,Module_name,Id,MenuIndex) Values( '" & mMenuName & "','" & mMenuCaption & "'," & I & "," & IIf(mHelpId = 100, True, False) & ",'','','" & mMenuIndex & "')")
        mMainMenuName = mMenuName ' mMenuName
        mMenuGroupID = Replace(Space(3 - Len(STR(j))), " ", "0") & j
        mMainMenuCaption = mMenuCaption
        If mHelpId = 1 Then mTopMenuName = mMenuName: mTopMenuCaption = mMenuCaption
        j = j + 1
    End If
nxt1:
    ErrFound = False
    End If

Next objctrl
Exit Sub
ELoop:
    MsgBox err.Description
MsgBox "Done"
DispErr:
    MsgBox err.Description
    Resume Next
End Sub
Private Sub DeleteSubgroup()
Dim RstParty As ADODB.Recordset
Dim mDelete As Boolean, I As Double
mDelete = True
Picture1.Visible = True
Label1.CAPTION = "Processing ---->  "
Dim Counter As Double
Counter = 0
Set RstParty = G_FaCn.Execute("Select SubCode from Subgroup where nature in ('Customer','Suplier')")
    If RstParty.RecordCount > 0 Then
        RstParty.MoveFirst
        For I = 1 To RstParty.RecordCount
                Counter = Format((I * 100) / RstParty.RecordCount, "0")
                Label1.CAPTION = "Processing ----> " & Counter & "% "
                If I Mod 50 = 0 Then
                    Label1.CAPTION = "Processing ----> " & Counter & "% "
                End If
                Picture1.Refresh
                mDelete = True
               '* For FA Data Delete..
                If G_FaCn.Execute("Select * From Ledger Where SubCode='" & RstParty!SubCode & "'").RecordCount > 0 Then
                    mDelete = False
                End If
                '* For Vehicle Data Delete..
                If GCn.Execute("Select * From Veh_Purch1 Where PartyCode='" & RstParty!SubCode & "'").RecordCount > 0 Then
                    mDelete = False
                End If
                
                If GCn.Execute("Select * From Veh_Order Where PartyCode='" & RstParty!SubCode & "'").RecordCount > 0 Then
                    mDelete = False
                End If
                
                '* For Spare Data Delete..
                If GCn.Execute("Select * From SP_Sale Where Party_Code='" & RstParty!SubCode & "'").RecordCount > 0 Then
                    mDelete = False
                End If
                
                If GCn.Execute("Select * From SP_Purch Where Party_Code='" & RstParty!SubCode & "'").RecordCount > 0 Then
                    mDelete = False
                End If
                
                If GCn.Execute("Select * From SP_Order Where Party_Code='" & RstParty!SubCode & "'").RecordCount > 0 Then
                    mDelete = False
                End If
                
                If GCn.Execute("Select * From RECT Where PartyCode='" & RstParty!SubCode & "'").RecordCount > 0 Then
                    mDelete = False
                End If
                
                If GCn.Execute("Select * From Rect Where AcCode='" & RstParty!SubCode & "'").RecordCount > 0 Then
                    mDelete = False
                End If
                
                If mDelete = True Then
                        GCn.Execute ("Delete From SubGroupAlias Where SubCode='" & RstParty!SubCode & "'")
                        GCn.Execute ("Delete From SubGroup Where SubCode='" & RstParty!SubCode & "'")
            
                        G_FaCn.Execute ("Delete From SubGroupAlias Where SubCode='" & RstParty!SubCode & "'")
                        G_FaCn.Execute ("Delete From SubGroup Where SubCode='" & RstParty!SubCode & "'")
                End If
                RstParty.MoveNext
        Next
    End If
    Picture1.Visible = False
End Sub
Sub SetMaxId_VoucherPrefix()
    On Error Resume Next
    SetMax_VoucherPrefix "DocId", "W_JWR", "Job_Card", "Job_Date"
    SetMax_VoucherPrefix "DocId", "W_JC", "Job_Card", "Job_Date"
    SetMax_VoucherPrefix "DocId", "W_JTR", "Job_Card", "Job_Date"
    SetMax_VoucherPrefix "DocId", "W_JWF", "Job_Card", "Job_Date"
    SetMax_VoucherPrefix "DocId", "W_JWM", "Job_Card", "Job_Date"
    SetMax_VoucherPrefix "DocId_InvLab", "W_LIC", "Job_Card", "Job_Date"
    SetMax_VoucherPrefix "DocId_InvLab", "W_LIR", "Job_Card", "Job_Date"
    SetMax_VoucherPrefix "DocId", "SYSC", "SP_Sale", "V_Date"
    SetMax_VoucherPrefix "DocId", "W_RG", "SP_Stock", "V_Date"
    SetMax_VoucherPrefix "DocId", "SXAO", "SP_Stock", "V_Date"
    SetMax_VoucherPrefix "DocId", "SXGR", "SP_Stock", "V_Date"
    SetMax_VoucherPrefix "DocId", "SXPIC", "SP_Purch", "V_Date"
    SetMax_VoucherPrefix "DocId", "SXPIR", "SP_Purch", "V_Date"
    SetMax_VoucherPrefix "DocId_InvSpr", "W_SIC", "Job_Card", "Job_Date"
    SetMax_VoucherPrefix "DocId_InvSpr", "W_SIR", "Job_Card", "Job_Date"
    SetMax_VoucherPrefix "DocId", "SYSIC", "SP_Sale", "V_Date"
    SetMax_VoucherPrefix "DocId", "SYSIR", "SP_Sale", "V_Date"
    SetMax_VoucherPrefix "DocId", "SXSRR", "SP_Sale", "V_Date"
    SetMax_VoucherPrefix "DocId", "SXSRC", "SP_Sale", "V_Date"
    SetMax_VoucherPrefix "DocId", "V_PB", "Veh_Purch1", "V_Date"
    SetMax_VoucherPrefix "DocId", "V_OST", "Veh_Purch1", "V_Date"
    SetMax_VoucherPrefix "DocId", "V_QOT", "Veh_Quot", "V_Date"
    SetMax_VoucherPrefix "OrdDocId", "V_BK", "Veh_Order", "Ord_Date"
    SetMax_VoucherPrefix "DelCh_DocId", "V_DCL", "Veh_Order", "Ord_Date"
    MsgBox "Updation Done"
End Sub

Private Function IsTableExists(DataBasePath As String, TabName As String) As Boolean
Dim DB As DAO.Database
Dim TableFound As Boolean, I As Integer
    TableFound = False
    If GCn.State = 1 Then GCn.Close
    Set DB = OpenDatabase(DataBasePath, True, False, ";pwd=dtman")
    For I = 0 To DB.TableDefs.Count - 1
        If UCase(DB.TableDefs(I).Name) = UCase(TabName) Then
            TableFound = True
            Exit For
        End If
    Next I
    Set DB = Nothing
    If GCn.State <> 1 Then GCn.Open
End Function

Sub Disp_Menu()
Dim Ctrl As Control
Dim RsTemp As ADODB.Recordset
On Error Resume Next

    If pubUName <> "SA" Then
'        For Each Ctrl In Me.Controls
'            If TypeOf Ctrl Is Menu Then
'                If Ctrl.CAPTION = "-" Or UCase(Ctrl.Name) = "MNUWINDOW" Or UCase(Ctrl.Name) = "MNUWINDOWCASCADE" Or UCase(Ctrl.Name) = "MNUWINDOWHORIZONTAL" Or UCase(Ctrl.Name) = "MNUWINDOWVERTICAL" Or UCase(Ctrl.Name) = "MNUEXIT" Then
'                    Ctrl.Visible = True
'                Else
'                    Ctrl.Visible = False
'                End If
'            End If
'        Next
    
        Set RsTemp = G_CompCn.Execute("Select User2.*, UM.Name, UM.Menu_Name From User2 Left join User_Module UM On User2.Module_Name=UM.Module_Name And User2.Form_Code=UM.Form_Code Where User_Name='" & pubUName & "' And Comp_Code='" & PubCenCompCode & "' And Div_Code='" & PubDivCode & "' And Param_Str='****'")
        If RsTemp.RecordCount > 0 Then
            Do Until RsTemp.EOF
                For Each Ctrl In Me.Controls
                    If TypeOf Ctrl Is Menu Then
'                        If UCase(RsTemp!Menu_Name) = UCase("MnVehReports1 14") Then
'                            MsgBox ""
'                        End If
                        
                        If UCase(Ctrl.Name & " " & Ctrl.Index) = UCase(RsTemp!Menu_Name) Then
                            If err.NUMBER = 0 Then
                                If Ctrl.CAPTION <> "-" Then
                                    Ctrl.Visible = False
                                End If
                            End If
                            err.NUMBER = 0
                        End If
                    End If
                Next
                
                
                RsTemp.MoveNext
            Loop
        End If
    End If

End Sub





Sub BlankDataSqlServer()
On Error Resume Next
    If PubBackEnd = "S" Then
        G_FaCn.Execute (" delete from AcControls")
        G_FaCn.Execute ("delete from AcGroup")
        G_FaCn.Execute ("delete from AcGroupFixAc")
        G_FaCn.Execute ("delete from AcMERGE")
        G_FaCn.Execute ("delete from AgeParamFA")
        G_FaCn.Execute ("delete from AgeParamINV")
        G_FaCn.Execute ("delete from Aggregate")
        G_FaCn.Execute ("delete from AMD_Dealer")
        G_FaCn.Execute ("delete from Area")
        G_FaCn.Execute ("delete from AreaMast")
        G_FaCn.Execute ("delete from AssociatedFirms")
        G_FaCn.Execute ("delete from BMS")
        G_FaCn.Execute ("delete from Category")
        G_FaCn.Execute ("delete from Chas_Mth")
        G_FaCn.Execute ("delete from Chas_Yr")
        G_FaCn.Execute ("delete from City")
        G_FaCn.Execute ("delete from ColMast")
        G_FaCn.Execute ("delete from ContractFinance")
        G_FaCn.Execute ("delete from CustFeedback")
        G_FaCn.Execute ("delete from CustInfo")
        G_FaCn.Execute ("delete from CVD")
        G_FaCn.Execute ("delete from DeleteLog")
        G_FaCn.Execute ("delete from Designation")
        G_FaCn.Execute ("delete from Division")
        G_FaCn.Execute ("delete from dtproperties")
        G_FaCn.Execute ("delete from Emp_Mast")
        G_FaCn.Execute ("delete from Estimate")
        G_FaCn.Execute ("delete from Estimate1")
        G_FaCn.Execute ("delete from FaEnviro")
        G_FaCn.Execute ("delete from FinBank")
        G_FaCn.Execute ("delete from FinGroup")
        G_FaCn.Execute ("delete from FinGroupSummary")
        G_FaCn.Execute ("delete from Godown")
        G_FaCn.Execute ("delete from HisCard")
        G_FaCn.Execute ("delete from Indent")
        G_FaCn.Execute ("delete from Inspection_Catg")
        G_FaCn.Execute ("delete from Inspection_Element")
        G_FaCn.Execute ("delete from Job_Booking")
        G_FaCn.Execute ("delete from Job_Card")
        G_FaCn.Execute ("delete from Job_Card2")
        G_FaCn.Execute ("delete from Job_Delay")
        G_FaCn.Execute ("delete from Job_Demand")
        G_FaCn.Execute ("delete from Job_GatePass")
        G_FaCn.Execute ("delete from Job_GatePass1")
        G_FaCn.Execute ("delete from Job_Inspection")
        G_FaCn.Execute ("delete from Job_Inspection2")
        G_FaCn.Execute ("delete from Job_Lab")
        G_FaCn.Execute ("delete from Job_Lab2")
        G_FaCn.Execute ("delete from Job_WarBill")
        G_FaCn.Execute ("delete from Job_Warr1")
        G_FaCn.Execute ("delete from Job_Warr2")
        G_FaCn.Execute ("delete from Lab_Trouble")
        G_FaCn.Execute ("delete from Labour")
        G_FaCn.Execute ("delete from Labour_CheckList")
        G_FaCn.Execute ("delete from Labour_Group")
        G_FaCn.Execute ("delete from Labour_Model")
        G_FaCn.Execute ("delete from Labour_Type")
        G_FaCn.Execute ("delete from LastVou")
        G_FaCn.Execute ("delete from LastVoucher")
        G_FaCn.Execute ("delete from Ledger")
        G_FaCn.Execute ("delete from LedgerAdj")
        G_FaCn.Execute ("delete from LedgerBack")
        G_FaCn.Execute ("delete from LedgerM")
        G_FaCn.Execute ("delete from LedgerRef")
        G_FaCn.Execute ("delete from LedgerTDS")
        G_FaCn.Execute ("delete from Model")
        G_FaCn.Execute ("delete from Model_Cat")
        G_FaCn.Execute ("delete from Model_Grp")
        G_FaCn.Execute ("delete from ModelCheckList")
        G_FaCn.Execute ("delete from ModelCheckListMast")
        G_FaCn.Execute ("delete from NarrMast")
        G_FaCn.Execute ("delete from OffTake")
        G_FaCn.Execute ("delete from OffTake1")
        G_FaCn.Execute ("delete from OverTime")
        G_FaCn.Execute ("delete from Part")
        G_FaCn.Execute ("delete from Part_Alternate")
        G_FaCn.Execute ("delete from Part_DiscFactor")
        G_FaCn.Execute ("delete from Part_Grade")
        G_FaCn.Execute ("delete from Part_Import")
        G_FaCn.Execute ("delete from Part_ImportCVD")
        G_FaCn.Execute ("delete from Part_PriceList")
        G_FaCn.Execute ("delete from PartList_New")
        G_FaCn.Execute ("delete from Photo")
        G_FaCn.Execute ("delete from prnmissrec")
        G_FaCn.Execute ("delete from Profession")
        G_FaCn.Execute ("delete from ProspectiveCust")
        G_FaCn.Execute ("delete from Purpose")
        G_FaCn.Execute ("delete from Rect")
        G_FaCn.Execute ("delete from Reffered")
        G_FaCn.Execute ("delete from RepeatJob")
        G_FaCn.Execute ("delete from RTOData")
        G_FaCn.Execute ("delete from RTODealer")
        G_FaCn.Execute ("delete from RTOMfg")
        G_FaCn.Execute ("delete from RTOModel")
        G_FaCn.Execute ("delete from SaleTarget")
        G_FaCn.Execute ("delete from Service_Rates")
        G_FaCn.Execute ("delete from Service_Type")
        G_FaCn.Execute ("delete from Site")
        G_FaCn.Execute ("delete from SP_OrdCoun")
        G_FaCn.Execute ("delete from SP_Order")
        G_FaCn.Execute ("delete from SP_Order1")
        G_FaCn.Execute ("delete from SP_Purch")
        G_FaCn.Execute ("delete from SP_Sale")
        G_FaCn.Execute ("delete from SP_Stock")
        G_FaCn.Execute ("delete from  State")
        G_FaCn.Execute ("delete from State1")
        G_FaCn.Execute ("delete from SubGroup")
        G_FaCn.Execute ("delete from  SubGroupAlias")
        G_FaCn.Execute ("delete from  SubGroupAliasold")
        G_FaCn.Execute ("delete from SubGroupCounter")
        G_FaCn.Execute ("delete from SubGroupOld")
        G_FaCn.Execute ("delete from SubGroupType")
        G_FaCn.Execute ("delete from Subvention")
        G_FaCn.Execute ("delete from Syctrl")
        G_FaCn.Execute ("delete from System_Log")
        G_FaCn.Execute ("delete from TableGroupClient")
        G_FaCn.Execute ("delete from TableGroupHO")
        G_FaCn.Execute ("delete from TaxForms")
        G_FaCn.Execute ("delete from TaxFormsAc")
        G_FaCn.Execute ("delete from TaxFormStk")
        G_FaCn.Execute ("delete from TDSCat")
        G_FaCn.Execute ("delete from TDSCerti")
        G_FaCn.Execute ("delete from TDSChal")
        G_FaCn.Execute ("delete from TDSChal1")
        G_FaCn.Execute ("delete from TransferCtrl")
        G_FaCn.Execute ("delete from Trouble")
        G_FaCn.Execute ("delete from Unit")
        G_FaCn.Execute ("delete from Veh_AddiService")
        G_FaCn.Execute ("delete from Veh_AMDModel")
        G_FaCn.Execute ("delete from Veh_CheckList")
        G_FaCn.Execute ("delete from VEH_CustConcn")
        G_FaCn.Execute ("delete from Veh_Forecast")
        G_FaCn.Execute ("delete from Veh_InvCancel")
        G_FaCn.Execute ("delete from VEH_Margin")
        G_FaCn.Execute ("delete from Veh_OfftakeIncentive")
        G_FaCn.Execute ("delete from Veh_Order")
        G_FaCn.Execute ("delete from Veh_Order1")
        G_FaCn.Execute ("delete from Veh_OrderM")
        G_FaCn.Execute ("delete from  Veh_OrdLostCatg")
        G_FaCn.Execute ("delete from  Veh_Purch1")
        G_FaCn.Execute ("delete from Veh_Purch2")
        G_FaCn.Execute ("delete from Veh_Quot")
        G_FaCn.Execute ("delete from Veh_Quot1")
        G_FaCn.Execute ("delete from Veh_Quot2")
        G_FaCn.Execute ("delete from Veh_Rate")
        G_FaCn.Execute ("delete from Veh_RateDate")
        G_FaCn.Execute ("delete from Veh_Stock")
        G_FaCn.Execute ("delete from Veh_SubGroupQuot")
        G_FaCn.Execute ("delete from Veh_Target")
        G_FaCn.Execute ("delete from Veh_Transfer")
        G_FaCn.Execute ("delete from VehBill_Counter")
        G_FaCn.Execute ("delete from Vehicle_Type")
        G_FaCn.Execute ("delete from VisitObjective")
        G_FaCn.Execute ("delete from Visits")
        G_FaCn.Execute ("delete from Voucher_Exclude")
        G_FaCn.Execute ("delete from Voucher_Include")
        G_FaCn.Execute ("delete from Voucher_Type")
        G_FaCn.Execute ("delete from Voucher_Prefix")
        G_FaCn.Execute ("delete from VoucherCat")
        G_FaCn.Execute ("delete from WarrCompMast")
        G_FaCn.Execute ("delete from WarrFailMast")
        G_FaCn.Execute ("delete from WarrJobMast")
        G_FaCn.Execute ("delete from WarrMakeMast")
        G_FaCn.Execute ("delete from Wrk_Details")
    End If
End Sub


Public Sub ChangeFieldType(DB As DAO.Database, _
                          ByVal TableName As String, _
                          ByVal FieldName As String, _
                          ByVal NewType As Integer, _
                          Optional NewSize As Long, _
                          Optional NewAllowZeroLength As Boolean = False, _
                          Optional NewAllowNulls As Boolean = True, _
                          Optional NewAttributes As Long)

' User-defined properties are not maintained

  Dim td As DAO.TableDef, I As Index, r As DAO.Relation, F As DAO.Field

' loop iterators for Indexes, Fields, and Relations collections:
  Dim I1 As Long, F1 As Long, R1 As Long

  Dim colR As Collection, colI As Collection
  Dim E_Desc As String, Process As String, SubProcess As String, E As Error
  Dim TempFieldName As String, Suffix As Long, OldName As String
  Dim Temp As Variant
  Dim OrdinalPosition As Long

  Set colI = New Collection
  Set colR = New Collection
  On Error GoTo CFT_Err
  DBEngine(0).BeginTrans

' Enumerate relations and save/remove them

  DBEngine(0).BeginTrans
  Process = "Removing relations on [" & TableName & "]![" & FieldName & "]"
  SubProcess = ""
  For R1 = DB.Relations.Count - 1 To 0 Step -1
    Set r = DB.Relations(R1)
    If r.Table = TableName Then
      For F1 = 0 To r.Fields.Count - 1
        Set F = r.Fields(F1)
        If F.Name = FieldName Then
          RecordRelationInfo r, colR
          SubProcess = "Removing relation " & r.Name
          DB.Relations.Delete r.Name
          Exit For
        End If
      Next F1
    ElseIf r.ForeignTable = TableName Then
      For F1 = 0 To r.Fields.Count - 1
        Set F = r.Fields(F1)
        If F.ForeignName = FieldName Then
          RecordRelationInfo r, colR
          SubProcess = "Removing relation " & r.Name
          DB.Relations.Delete r.Name
          Exit For
        End If
      Next F1
    End If
  Next R1
  Set F = Nothing
  Set r = Nothing
  DBEngine(0).CommitTrans

' Enumerate indices and save/remove them

  DBEngine(0).BeginTrans
  Process = "Removing indexes on [" & TableName & "]![" & FieldName & "]"
  SubProcess = ""
  DB.TableDefs.Refresh
  Set td = DB(TableName)
  td.Indexes.Refresh
  For I1 = td.Indexes.Count - 1 To 0 Step -1
    Set I = td.Indexes(I1)
    If I.Foreign <> True Then
      For F1 = 0 To I.Fields.Count - 1
        Set F = I.Fields(F1)
        If F.Name = FieldName Then
          RecordIndexInfo I, colI
          SubProcess = "Removing index " & I.Name
          td.Indexes.Delete I.Name
          Exit For
        End If
      Next F1
    End If
  Next I1
  Set F = Nothing
  Set I = Nothing
  DBEngine(0).CommitTrans

' Rename Field

  DBEngine(0).BeginTrans
  Process = "Renaming field"
  SubProcess = ""
  td.Fields.Refresh
  Set F = td(FieldName)
  OrdinalPosition = F.OrdinalPosition   ' save this value

  ' determine a field name not in use
  Suffix = 0
  Do
    Suffix = Suffix + 1
    TempFieldName = "XXX" & Suffix
  Loop While IsField(td, TempFieldName)

  ' rename the field
  SubProcess = "to " & TempFieldName
  F.Name = TempFieldName

  Set F = Nothing
  DBEngine(0).CommitTrans

' Add new Field

  DBEngine(0).BeginTrans
  Process = "Adding new field"
  SubProcess = ""
  td.Fields.Refresh
  Set F = td.CreateField(FieldName, NewType)
  If NewSize Then F.Size = NewSize
  F.AllowZeroLength = NewAllowZeroLength
  F.Required = Not NewAllowNulls
  F.Attributes = NewAttributes
  F.OrdinalPosition = OrdinalPosition
  td.Fields.Append F
  Set F = Nothing
  Set td = Nothing
  DBEngine(0).CommitTrans

' Copy data

  DBEngine(0).BeginTrans
  Process = "Copying data from " & TempFieldName & " to " & FieldName
  SubProcess = ""
  DB.Execute "UPDATE [" & TableName & "] SET [" & FieldName & "]=[" & _
              TempFieldName & "]", dbFailOnError
  DBEngine(0).CommitTrans

' Delete temporary field

  DBEngine(0).BeginTrans
  Process = "Deleting temporary field " & TempFieldName
  SubProcess = ""
  Set td = DB(TableName)
  td.Fields.Delete TempFieldName
  DBEngine(0).CommitTrans

' Add back Indices

  DBEngine(0).BeginTrans
  Process = "Adding indexes back into table"
  SubProcess = ""
  Set td = DB(TableName)
  td.Fields.Refresh
  td.Indexes.Refresh
  OldName = ""
  Set I = Nothing
  For Each Temp In colI
    If Temp(I_NAME) <> OldName Then
      If Not (I Is Nothing) Then   ' handle first time through case
        SubProcess = "Adding index " & I.Name
        td.Indexes.Append I
      End If
      Set I = td.CreateIndex(Temp(I_NAME))
      I.Primary = Temp(I_PRIMARY)
      I.Unique = Temp(I_UNIQUE)
      I.Required = Temp(I_REQUIRED)
      I.IgnoreNulls = Temp(I_IGNORENULLS)
      I.Clustered = Temp(I_CLUSTERED)
    End If
    Set F = I.CreateField(Temp(I_FIELD))
    F.Attributes = Temp(I_FIELDATTRIBUTES)  ' to handle descending index
    I.Fields.Append F
  Next Temp
  If Not (I Is Nothing) Then   ' handle case of no indexes
    SubProcess = "Adding index " & I.Name
    td.Indexes.Append I
  End If
  Set F = Nothing
  Set I = Nothing
  Set td = Nothing
  DBEngine(0).CommitTrans

' Add back relations

  DBEngine(0).BeginTrans
  Process = "Adding relations back into database"
  SubProcess = ""
  OldName = ""
  DB.Relations.Refresh
  Set r = Nothing
  For Each Temp In colR
    If Temp(I_NAME) <> OldName Then
      If Not (r Is Nothing) Then   ' handle first time through case
        SubProcess = "Adding relation " & r.Name
        DB.Relations.Append r
      End If
      Set r = DB.CreateRelation(Temp(R_NAME), Temp(R_TABLE), _
                                Temp(R_FOREIGNTABLE), Temp(R_ATTRIBUTES))
    End If
    Set F = r.CreateField(Temp(R_FIELD))
    F.ForeignName = Temp(R_FOREIGNFIELD)
    r.Fields.Append F
  Next Temp
  If Not (r Is Nothing) Then   ' if there are no indexes...
    SubProcess = "Adding relation " & r.Name
    DB.Relations.Append r
  End If
  Set F = Nothing
  Set r = Nothing
  DBEngine(0).CommitTrans

' Commit all pending chhanges

  DBEngine(0).CommitTrans
  Exit Sub
  
CFT_Abort:
  On Error Resume Next
  Set F = Nothing
  Set td = Nothing
  DBEngine(0).Rollback
  DBEngine(0).Rollback
  err.Clear
  On Error GoTo 0
  'err.Raise CFT_Failed, "ChangeFieldType", E_Desc
  Exit Sub
  
CFT_Err:
  E_Desc = "Error " & Process
  If SubProcess <> "" Then E_Desc = E_Desc & vbCrLf & SubProcess
  If DBEngine.Errors.Count = 0 Then
    E_Desc = E_Desc & vbCrLf & "Error " & err.NUMBER & " " & _
             err.Description
'  Else
'    For Each E In DBEngine.Errors
'      E_Desc = E_Desc & vbCrLf & "Error " & E.NUMBER & " (" & _
'               E.Source & ") " & E.Description
'    Next E
  End If
  Debug.Print E_Desc
  Resume CFT_Abort
End Sub

Private Sub RecordRelationInfo(ByVal r As Relation, colR As Collection)

' Records information regarding the relationship and its fields
' in the colR collection.

  Dim F1 As Long, F As DAO.Field
  For F1 = 0 To r.Fields.Count - 1
    Set F = r.Fields(F1)
    colR.Add MakeArray(r.Name, r.Attributes, r.Table, r.ForeignTable, _
                        F.Name, F.ForeignName)
  Next F1
End Sub

Private Sub RecordIndexInfo(ByVal I As Index, colI As Collection)

' Records information about fields in the index and about the index itself
' into the colI collection.

  Dim F1 As Long, F As DAO.Field
  For F1 = 0 To I.Fields.Count - 1
    Set F = I.Fields(F1)
    colI.Add MakeArray(I.Name, I.Primary, I.Unique, I.Required, _
                       I.IgnoreNulls, I.Clustered, F.Name, F.Attributes)
  Next F1
End Sub

Private Function IsField(td As TableDef, ByVal FieldName As String) _
        As Boolean

' Returns TRUE if a field exists in the table with the same name as
'    specified in FieldName.
' Returns FALSE otherwise.

   Dim F As DAO.Field
   err.Clear
   On Error Resume Next
   Set F = td(FieldName)
   IsField = err.NUMBER = 0
   err.Clear
End Function
 
Private Function MakeArray(ParamArray X() As Variant) As Variant
  
  ' Does the same thing as the Array() function in VB6
  
    MakeArray = X
  
End Function





