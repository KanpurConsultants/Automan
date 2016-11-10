VERSION 5.00
Begin VB.Form FaEnvron 
   BackColor       =   &H00CDCCFB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Environment Setting"
   ClientHeight    =   6510
   ClientLeft      =   180
   ClientTop       =   795
   ClientWidth     =   10455
   Icon            =   "FaEnvron.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10455
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton BtnIntegrity 
      Caption         =   "Integrity Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5670
      TabIndex        =   85
      Top             =   4995
      Width           =   2910
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00CBBE9E&
      Caption         =   "Reports"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   15
      TabIndex        =   83
      Top             =   5415
      Width           =   5460
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   37
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   20
         Top             =   195
         Width           =   450
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Separate Page for each day in Cash/Bank Book (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   21
         Left            =   780
         TabIndex        =   84
         Top             =   195
         Width           =   4005
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00CBBE9E&
      Caption         =   "Opening/Closing Stock Values"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   5460
      TabIndex        =   78
      Top             =   1950
      Width           =   2925
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   34
         Left            =   810
         TabIndex        =   35
         Top             =   435
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   35
         Left            =   810
         TabIndex        =   33
         Top             =   225
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   33
         Left            =   1710
         TabIndex        =   36
         Top             =   435
         Width           =   1140
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   32
         Left            =   1710
         TabIndex        =   34
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open.Qty"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   27
         Left            =   105
         TabIndex        =   82
         Top             =   225
         Width           =   675
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clos.Qty"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   26
         Left            =   195
         TabIndex        =   81
         Top             =   435
         Width           =   585
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   17
         Left            =   1275
         TabIndex        =   80
         Top             =   435
         Width           =   405
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   16
         Left            =   1275
         TabIndex        =   79
         Top             =   225
         Width           =   405
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00CBBE9E&
      Caption         =   "Dos Printing"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   15
      TabIndex        =   68
      Top             =   4260
      Width           =   5460
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   26
         Left            =   4845
         MaxLength       =   6
         TabIndex        =   16
         Top             =   225
         Width           =   510
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   25
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   17
         Top             =   435
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   24
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   19
         Top             =   855
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   18
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name And Address Position on the Report (L/M) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   14
         Left            =   540
         TabIndex        =   72
         Top             =   225
         Width           =   4260
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Date Should be Printed on the Report (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   13
         Left            =   1035
         TabIndex        =   71
         Top             =   435
         Width           =   3765
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Character to fill the lines in the report ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   12
         Left            =   2100
         TabIndex        =   70
         Top             =   855
         Width           =   2700
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Page No. Should be Printed on the Report (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   11
         Left            =   1230
         TabIndex        =   69
         Top             =   645
         Width           =   3570
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00CBBE9E&
      Caption         =   "Display"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   15
      TabIndex        =   62
      Top             =   2715
      Width           =   5460
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   15
         Top             =   1275
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   29
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1065
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   12
         Top             =   645
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   21
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   13
         Top             =   855
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   11
         Top             =   435
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   18
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   10
         Top             =   225
         Width           =   450
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Monthly Summary (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   18
         Left            =   2505
         TabIndex        =   76
         Top             =   1275
         Width           =   2265
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Balance in Cash/Bank Book (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   17
         Left            =   1725
         TabIndex        =   75
         Top             =   1065
         Width           =   3045
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Site Code in Ledger (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   10
         Left            =   2355
         TabIndex        =   67
         Top             =   645
         Width           =   2415
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Vr.Prefix in Ledger (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   9
         Left            =   2460
         TabIndex        =   66
         Top             =   855
         Width           =   2310
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Division Code in Ledger (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   8
         Left            =   2070
         TabIndex        =   65
         Top             =   435
         Width           =   2700
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Verticle Balance Sheet (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   7
         Left            =   2130
         TabIndex        =   63
         Top             =   225
         Width           =   2640
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00CBBE9E&
      Caption         =   "Define Voucher"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   15
      TabIndex        =   55
      Top             =   1950
      Width           =   5460
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   9
         Top             =   435
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   8
         Top             =   225
         Width           =   450
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Only System defined A/C Group in Don't Show A/C Groups (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   30
         TabIndex        =   60
         Top             =   435
         Width           =   4740
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Only System defined A/C Group in Must Show A/C Groups (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   59
         Top             =   225
         Width           =   4710
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00CBBE9E&
      Caption         =   "Voucher Entry"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   15
      TabIndex        =   54
      Top             =   -15
      Width           =   5460
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   31
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1695
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   28
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1275
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   27
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1485
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1065
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   3
         Top             =   855
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00E2FAFC&
         Height          =   195
         Index           =   6
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   0
         Top             =   225
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   1
         Top             =   435
         Width           =   450
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   2
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Online Adjustment (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   19
         Left            =   2940
         TabIndex        =   77
         Top             =   1695
         Width           =   1845
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show A/c Group && Address (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   16
         Left            =   2265
         TabIndex        =   74
         Top             =   1275
         Width           =   2520
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enable A/c Help Filter (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   15
         Left            =   2655
         TabIndex        =   73
         Top             =   1485
         Width           =   2130
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show City Name in A/C Help (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   2175
         TabIndex        =   64
         Top             =   1065
         Width           =   2610
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Ledger Current Balance (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   2085
         TabIndex        =   61
         Top             =   855
         Width           =   2700
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limit Check for Debtors (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   6
         Left            =   2565
         TabIndex        =   58
         Top             =   225
         Width           =   2220
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limit Check for Creditors (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   2505
         TabIndex        =   57
         Top             =   435
         Width           =   2280
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warn on Negative Cash Balance (Y/N) ?"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   1875
         TabIndex        =   56
         Top             =   645
         Width           =   2910
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00CBBE9E&
      Caption         =   "Ageing Parameters"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   5460
      TabIndex        =   39
      Top             =   -15
      Width           =   2925
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   10
         Left            =   1710
         TabIndex        =   27
         Top             =   495
         Width           =   1140
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   15
         Left            =   1710
         TabIndex        =   32
         Top             =   1545
         Width           =   1140
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   14
         Left            =   1710
         TabIndex        =   31
         Top             =   1335
         Width           =   1140
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   13
         Left            =   1710
         TabIndex        =   30
         Top             =   1125
         Width           =   1140
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   12
         Left            =   1710
         TabIndex        =   29
         Top             =   915
         Width           =   1140
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   11
         Left            =   1710
         TabIndex        =   28
         Top             =   705
         Width           =   1140
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   22
         Top             =   705
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   23
         Top             =   915
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   24
         Top             =   1125
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   25
         Top             =   1335
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   26
         Top             =   1545
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2FAFC&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   21
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   7
         Left            =   1500
         TabIndex        =   53
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   6
         Left            =   510
         TabIndex        =   52
         Top             =   225
         Width           =   435
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slab 1"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   1230
         TabIndex        =   51
         Top             =   495
         Width           =   450
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slab 6"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   1230
         TabIndex        =   50
         Top             =   1545
         Width           =   450
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slab 5"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   1230
         TabIndex        =   49
         Top             =   1335
         Width           =   450
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slab 4"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   1230
         TabIndex        =   48
         Top             =   1125
         Width           =   450
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slab 3"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   1230
         TabIndex        =   47
         Top             =   915
         Width           =   450
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slab 2"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   1230
         TabIndex        =   46
         Top             =   705
         Width           =   450
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period 2"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   9
         Left            =   105
         TabIndex        =   45
         Top             =   705
         Width           =   585
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period 3"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   10
         Left            =   105
         TabIndex        =   44
         Top             =   915
         Width           =   585
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period 4"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   11
         Left            =   105
         TabIndex        =   43
         Top             =   1125
         Width           =   585
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period 5"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   12
         Left            =   105
         TabIndex        =   42
         Top             =   1335
         Width           =   585
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period 6"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   13
         Left            =   105
         TabIndex        =   41
         Top             =   1545
         Width           =   585
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period 1"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   14
         Left            =   105
         TabIndex        =   40
         Top             =   495
         Width           =   585
      End
   End
   Begin VB.CommandButton btnok 
      Caption         =   "&Ok"
      DisabledPicture =   "FaEnvron.frx":030A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5670
      Picture         =   "FaEnvron.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Ok"
      Top             =   3990
      Width           =   1455
   End
   Begin VB.CommandButton BTNCANCEL 
      Caption         =   "&Cancel"
      DisabledPicture =   "FaEnvron.frx":054E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7125
      Picture         =   "FaEnvron.frx":0690
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Cancel Changes"
      Top             =   3990
      Width           =   1455
   End
End
Attribute VB_Name = "FaEnvron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CtrlBCol = &H80000008, CtrlFCol = &H8000000E
Private Const Age1 As Byte = 0, Age2 As Byte = 1, Age3 As Byte = 2, Age4 As Byte = 3
Private Const Age5 As Byte = 4, Age6 As Byte = 5, Amt1 As Byte = 10, Amt2 As Byte = 11
Private Const Amt3  As Byte = 12, Amt4 As Byte = 13, Amt5 As Byte = 14, Amt6 As Byte = 15
Private Const CreditLimit As Byte = 7, DebitLimit As Byte = 6, NegativeCashBalance As Byte = 8
Private Const ShowGroup As Byte = 9, DonotShowGroup  As Byte = 16, ShowCurrentBalance   As Byte = 17
Private Const VerticleBalanceSheet As Byte = 18, ShowCityName As Byte = 19, LedDivCode As Byte = 20
Private Const LedSiteCode As Byte = 22, LedPrefix As Byte = 21, titlerfill As Byte = 26
Private Const daterfill As Byte = 25, pagenofill As Byte = 23, linefiller As Byte = 24
Private Const FilterAC As Byte = 27, AddressHelp As Byte = 28, CashBookBalance As Byte = 29
Private Const MonthTotal As Byte = 30, OnLineAdjustment As Byte = 31
Private Const OpStockQTY As Byte = 35, OpStockValue As Byte = 32, ClStockQTY As Byte = 34, ClStockValue As Byte = 33, CashBookPage As Byte = 37
Dim mREFRESH As Boolean
Private PubDatamanFa As New DMFa.ClsFa

Public Sub Ctrl_validate(Ctrl As Object)
    Ctrl.BackColor = CtrlBColOrg
    Ctrl.ForeColor = CtrlFColOrg
End Sub
Public Sub Ctrl_GetFocus(Ctrl As Object)
    Ctrl.BackColor = &H80000008
    Ctrl.ForeColor = &H8000000E
End Sub
Private Sub SaveMsg()
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
        btnok_Click
    Else
        Txt(DebitLimit).SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn, vbKeyDown, vbKeyUp
        Select Case KeyCode
            Case vbKeyDown, vbKeyUp
        End Select
        If TypeOf Me.ActiveControl Is TextBox Then Txt_Validate Me.ActiveControl.Index, False
        If PubDatamanFa.FaManageKeysControl(Me, KeyCode, Shift) = True Then SaveMsg
        KeyCode = 0
End Select
End Sub
Private Sub BtnCancel_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set PubDatamanFa = Nothing
End Sub
Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus Txt(Index)
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case Age1, Age2, Age3, Age4, Age5, Age6
        FaNumDown Txt(Index), KeyCode, 4, 0
    Case Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, OpStockValue, ClStockValue
        FaNumDown Txt(Index), KeyCode, 10, 2
    Case OpStockQTY, ClStockQTY
        FaNumDown Txt(Index), KeyCode, 8, 3
End Select
End Sub
Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case Age1, Age2, Age3, Age4, Age5, Age6
        FaNumPress Txt(Index), KeyAscii, 4, 0
    Case Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, OpStockValue, ClStockValue
        FaNumPress Txt(Index), KeyAscii, 10, 2
    Case OpStockQTY, ClStockQTY
        FaNumPress Txt(Index), KeyAscii, 8, 3
    Case DebitLimit, CreditLimit, NegativeCashBalance, ShowGroup, DonotShowGroup, ShowCurrentBalance, VerticleBalanceSheet, ShowCityName, LedDivCode, LedSiteCode, LedPrefix, daterfill, pagenofill, FilterAC, AddressHelp, CashBookBalance, MonthTotal, OnLineAdjustment, CashBookPage
        If KeyAscii = 78 Or KeyAscii = 110 Then   'NO
            Txt(Index) = "No"
            KeyAscii = 0
        ElseIf KeyAscii = 89 Or KeyAscii = 121 Then 'Yes
            Txt(Index) = "Yes"
            KeyAscii = 0
        Else
            KeyAscii = 0
        End If
    Case titlerfill
        If KeyAscii = 77 Or KeyAscii = 109 Then
            Txt(Index) = "Middle"
            KeyAscii = 0
        ElseIf KeyAscii = 76 Or KeyAscii = 108 Then
            Txt(Index) = "Left"
            KeyAscii = 0
        Else
            KeyAscii = 0
        End If
End Select
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case OpStockValue, ClStockValue
        Txt(Index) = Format(FaValidate_Numeric(Txt(Index)), "0.00")
    Case OpStockQTY, ClStockQTY
        Txt(Index) = Format(FaValidate_Numeric(Txt(Index)), "0.000")
    Case Age1, Age2, Age3, Age4, Age5, Age6
        Txt(Index) = Format(FaValidate_Numeric(Txt(Index)), "0")
        If Val(Txt(Index)) = 0 Or IsNumeric(Txt(Index)) = False Then
            MsgBox " Nil Amount Not Accepted ", vbCritical, Me.CAPTION
            Txt(Index).TEXT = ""
            Txt(Index).SetFocus
        End If
    Case Amt1, Amt2, Amt3, Amt4, Amt5, Amt6
        Txt(Index) = Format(FaValidate_Numeric(Txt(Index)), "0.00")
        If Val(Txt(Index)) = 0 Or IsNumeric(Txt(Index)) = False Then
            MsgBox " Nil Quantity Not Accepted ", vbCritical, Me.CAPTION
            Txt(Index).TEXT = ""
            Txt(Index).SetFocus
        End If
End Select
End Sub
Private Sub Form_Load()
Dim RST1 As ADODB.Recordset, I As Integer
    Me.left = 0
    Me.top = 0
    Set RST1 = G_FaCn.Execute("select * from FaEnviro")
    If RST1.RecordCount > 0 Then
        Txt(Age1) = FaVNull(RST1!Age1)
        Txt(Age2) = FaVNull(RST1!Age2)
        Txt(Age3) = FaVNull(RST1!Age3)
        Txt(Age4) = FaVNull(RST1!Age4)
        Txt(Age5) = FaVNull(RST1!Age5)
        Txt(Age6) = FaVNull(RST1!Age6)
        Txt(Amt1) = FaSNull(RST1!Amt1)
        Txt(Amt2) = FaSNull(RST1!Amt2)
        Txt(Amt3) = FaSNull(RST1!Amt3)
        Txt(Amt4) = FaSNull(RST1!Amt4)
        Txt(Amt5) = FaSNull(RST1!Amt5)
        Txt(Amt6) = FaSNull(RST1!Amt6)
        Txt(DebitLimit) = FaXNull(RST1!DebitLimit)
        Txt(CreditLimit) = FaXNull(RST1!CreditLimit)
        Txt(NegativeCashBalance) = FaXNull(RST1!NegativeCashBalance)
        Txt(ShowGroup) = FaXNull(RST1!ShowGroup)
        Txt(DonotShowGroup) = FaXNull(RST1!DonotShowGroup)
        Txt(ShowCurrentBalance) = FaXNull(RST1!ShowCurrentBalance)
        Txt(VerticleBalanceSheet) = FaXNull(RST1!VerticleBalanceSheet)
        Txt(ShowCityName) = FaXNull(RST1!ShowCityName)
        Txt(LedDivCode) = FaXNull(RST1!LedDivCode)
        Txt(LedSiteCode) = FaXNull(RST1!LedSiteCode)
        Txt(LedPrefix) = FaXNull(RST1!LedPrefix)
        Txt(linefiller) = FaXNull(RST1!linefiller)
        Txt(daterfill) = IIf(FaXNull(RST1!daterfill) = "Y", "Yes", "No")
        Txt(titlerfill) = IIf(FaXNull(RST1!titlerfill) = "M", "Middle", "Left")
        Txt(pagenofill) = IIf(FaXNull(RST1!pagenofill) = "Y", "Yes", "No")
        Txt(FilterAC) = FaXNull(RST1!FilterAC)
        Txt(AddressHelp) = FaXNull(RST1!AddressHelp)
        Txt(CashBookBalance) = FaXNull(RST1!CashBookBalance)
        Txt(MonthTotal) = FaXNull(RST1!MonthTotal)
        Txt(OnLineAdjustment) = FaXNull(RST1!OnLineAdjustment)
        Txt(OpStockQTY) = FaXNull(RST1!OpStockQTY)
        Txt(OpStockValue) = FaXNull(RST1!OpStockValue)
        Txt(ClStockQTY) = FaXNull(RST1!ClStockQTY)
        Txt(ClStockValue) = FaXNull(RST1!ClStockValue)
        Txt(CashBookPage) = FaXNull(RST1!CashBookPage)
    Else
        For I = 0 To Txt.Count - 1
            Txt(I) = ""
        Next
    End If
Set RST1 = Nothing
End Sub
Private Sub btnok_Click()
If (Val(Txt(Age1)) >= Val(Txt(Age2)) Or Val(Txt(Age2)) >= Val(Txt(Age3)) Or Val(Txt(Age3)) >= Val(Txt(Age4)) Or Val(Txt(Age4)) >= Val(Txt(Age5)) Or Val(Txt(Age5)) >= Val(Txt(Age6))) Then
    MsgBox " Time Periods Must be in Increasing Order ", vbCritical, Me.CAPTION
    Exit Sub
ElseIf (Val(Txt(Amt1)) >= Val(Txt(Amt2)) Or Val(Txt(Amt2)) >= Val(Txt(Amt3)) Or Val(Txt(Amt3)) >= Val(Txt(Amt4)) Or Val(Txt(Amt4)) >= Val(Txt(Amt5)) Or Val(Txt(Amt5)) >= Val(Txt(Amt6))) Then
    MsgBox " Amount Slab Must be in Increasing Order ", vbCritical, Me.CAPTION
    Exit Sub
Else
    If G_FaCn.Execute("SELECT COUNT(*) FROM FaEnviro").Fields(0) <= 0 Then G_FaCn.Execute "INSERT INTO FAENVIRO (AGE1) VALUES (0)"
    G_FaCn.Execute "update FaEnviro SET AGE1=" & FaVNull(Txt(Age1)) & ",AGE2=" & FaVNull(Txt(Age2)) & ",AGE3=" & FaVNull(Txt(Age3)) & ",AGE4=" & FaVNull(Txt(Age4)) & ",AGE5=" & FaVNull(Txt(Age5)) & ",AGE6=" & FaVNull(Txt(Age6)) & ",AMT1=" & FaVNull(Txt(Amt1)) & ",AMT2=" & FaVNull(Txt(Amt2)) & ",AMT3=" & FaVNull(Txt(Amt3)) & ",AMT4=" & FaVNull(Txt(Amt4)) & ",AMT5=" & FaVNull(Txt(Amt5)) & ",AMT6=" & FaVNull(Txt(Amt6)) & ",DebitLimit=" & FaChk_Text(Txt(DebitLimit)) & ",CreditLimit=" & FaChk_Text(Txt(CreditLimit)) & ",NegativeCashBalance=" & FaChk_Text(Txt(NegativeCashBalance)) & ",ShowGroup=" & FaChk_Text(Txt(ShowGroup)) & ",DonotShowGroup=" & FaChk_Text(Txt(DonotShowGroup)) & ",ShowCurrentBalance=" & FaChk_Text(Txt(ShowCurrentBalance)) & ",VerticleBalanceSheet=" & FaChk_Text(Txt(VerticleBalanceSheet)) & ",ShowCityName=" & FaChk_Text(Txt(ShowCityName)) & ",LedDivCode=" & FaChk_Text(Txt(LedDivCode)) & "," & _
    "LedSiteCode=" & FaChk_Text(Txt(LedSiteCode)) & ",LedPrefix=" & FaChk_Text(Txt(LedPrefix)) & ",linefiller=" & FaChk_Text(Txt(linefiller)) & ",daterfill=" & FaChk_Text(left(Txt(daterfill), 1)) & ",titlerfill=" & FaChk_Text(left(Txt(titlerfill), 1)) & ",pagenofill=" & FaChk_Text(left(Txt(pagenofill), 1)) & ",FilterAC=" & FaChk_Text(Txt(FilterAC)) & ",AddressHelp=" & FaChk_Text(Txt(AddressHelp)) & ",CashBookBalance=" & FaChk_Text(Txt(CashBookBalance)) & ",MonthTotal=" & FaChk_Text(Txt(MonthTotal)) & ",OnLineAdjustment=" & FaChk_Text(Txt(OnLineAdjustment)) & ",OpStockQTY=" & FaVNull(Txt(OpStockQTY)) & ",OpStockValue=" & FaVNull(Txt(OpStockValue)) & ",ClStockQTY=" & FaVNull(Txt(ClStockQTY)) & ",ClStockValue=" & FaVNull(Txt(ClStockValue)) & ",CashBookPage=" & FaChk_Text(Txt(CashBookPage))
End If
Unload Me
End Sub
Private Sub BtnIntegrity_Click()
Dim fob As New Scripting.FileSystemObject, varTxtstrm As Scripting.TextStream
Dim Rst As ADODB.Recordset
'On Error GoTo ERRORHANDLER
    MousePointer = vbHourglass
    fob.CreateTextFile "C:\FaCheckList.Log", True
    Set varTxtstrm = fob.OpenTextFile("C:\FaCheckList.Log", ForAppending)
    varTxtstrm.Write "***** DATAMAN COMPUTER SYSTEMS (P) LTD.  " + vbCrLf
    varTxtstrm.Write "-----------------------------------------" + vbCrLf
    varTxtstrm.Write "***** DESCRIPTION : F.A. Integrity Tool  " + vbCrLf
    varTxtstrm.Write "***** WRITTEN BY  : SANJEEV KUMAR GUPTA  " + vbCrLf
    varTxtstrm.Write "***** DATE UPDATED: 12/MAR/2005" + vbCrLf
    varTxtstrm.Write "-----------------------------------------" + vbCrLf
    Set Rst = G_FaCn.Execute("SELECT * FROM LEDGER WHERE SUBCODE IS NULL OR SUBCODE=''")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "***** These Ledger DocId Have Blank SubCode *****" + vbCrLf
        varTxtstrm.Write "---------------------" + vbCrLf
        varTxtstrm.Write "DocId" + vbCrLf
        varTxtstrm.Write "---------------------" + vbCrLf
        Do Until Rst.EOF
            varTxtstrm.Write FaXNull(Rst!DocID) + vbCrLf
            Rst.MoveNext
        Loop
        varTxtstrm.Write "---------------------" + vbCrLf
    End If
    varTxtstrm.Write " " + vbCrLf
    varTxtstrm.Write "***** Updating Ledger for Null Values in LedgerM/Ledger Table *****" + vbCrLf
    G_FaCn.Execute "UPDATE LEDGER SET AMTDR=0 WHERE AMTDR IS NULL"
    G_FaCn.Execute "UPDATE LEDGER SET AMTCR=0 WHERE AMTCR IS NULL"
    G_FaCn.Execute "UPDATE LEDGER SET NARRATION='' WHERE NARRATION IS NULL"
    G_FaCn.Execute "UPDATE LEDGERM SET NARRATION='' WHERE NARRATION IS NULL"
        
    Set Rst = G_FaCn.Execute("SELECT LEDGER.*,SUBGROUP.NAME FROM LEDGER LEFT JOIN SUBGROUP ON LEDGER.SUBCODE=SUBGROUP.SUBCODE WHERE SUBGROUP.NAME IS NULL")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "***** These Ledger Subcode(S) are not Available in Subgroup *****" + vbCrLf
        varTxtstrm.Write "--------" + vbCrLf
        varTxtstrm.Write "SUBCODE " + vbCrLf
        varTxtstrm.Write "--------" + vbCrLf
        Do Until Rst.EOF
            varTxtstrm.Write FaXNull(Rst!SubCode) + vbCrLf
            Rst.MoveNext
        Loop
        varTxtstrm.Write "--------" + vbCrLf
    End If
    
    Set Rst = G_FaCn.Execute("SELECT * FROM SUBGROUP WHERE GROUPCODE NOT IN (SELECT GROUPCODE FROM ACGROUP)")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "***** These SUBGROUP GroupCode(S) are not Available in ACgroup *****" + vbCrLf
        varTxtstrm.Write "---------" + vbCrLf
        varTxtstrm.Write "GROUPCODE" + vbCrLf
        varTxtstrm.Write "---------" + vbCrLf
        Do Until Rst.EOF
            varTxtstrm.Write FaXNull(Rst!GroupCode) + vbCrLf
            Rst.MoveNext
        Loop
        varTxtstrm.Write "---------" + vbCrLf
    End If
    
    Set Rst = G_FaCn.Execute("SELECT SUBGROUP.*,ACGROUP.GROUPNATURE AS GRNAT FROM SUBGROUP LEFT JOIN ACGROUP ON SUBGROUP.GROUPCODE=ACGROUP.GROUPCODE WHERE SUBGROUP.GROUPNATURE<>ACGROUP.GROUPNATURE")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "***** These GROUP NATURE MISMATCH *****" + vbCrLf
        varTxtstrm.Write "----------------------------------------------------------------------------------" + vbCrLf
        varTxtstrm.Write "SUBCODE  A/CNAME                                            SUBGROUP        GROUP " + vbCrLf
        varTxtstrm.Write "                                                            Nature          Nature" + vbCrLf
        varTxtstrm.Write "----------------------------------------------------------------------------------" + vbCrLf
        Do Until Rst.EOF
            G_FaCn.Execute ("UPDATE SUBGROUP SET GROUPNATURE='" & Rst!GRNAT & "' WHERE SUBCODE='" & Rst!SubCode & "'")
            varTxtstrm.Write FaSetW(FaXNull(Rst!SubCode), 8) + " " + FaSetW(FaXNull(Rst!Name), 50) + " " + FaSetW(FaXNull(Rst!GroupNature), 15) + " " + FaSetW(FaXNull(Rst!GRNAT), 15) + vbCrLf
            Rst.MoveNext
        Loop
        varTxtstrm.Write "----------------------------------------------------------------------------------" + vbCrLf
    End If
    
    Set Rst = G_FaCn.Execute("SELECT * FROM SUBGROUP WHERE GROUPNATURE IS NULL OR GROUPNATURE=''")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "***** These SUBGROUP GroupNature is Blank/Null *****" + vbCrLf
        Do Until Rst.EOF
            varTxtstrm.Write FaXNull(Rst!SubCode) + " " + FaXNull(Rst!Name) + vbCrLf
            Rst.MoveNext
        Loop
    End If
    
    Set Rst = G_FaCn.Execute("SELECT * FROM SUBGROUP WHERE ALIASYN IS NULL OR ALIASYN=''")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "***** These SUBGROUP AliasYN is Blank/Null *****" + vbCrLf
        Do Until Rst.EOF
            varTxtstrm.Write FaXNull(Rst!SubCode) + " " + FaXNull(Rst!Name) + vbCrLf
            Rst.MoveNext
        Loop
    End If
    
    G_FaCn.Execute "UPDATE SUBGROUP SET AliasYN='N' WHERE AliasYN='0' OR ALIASYN IS NULL OR ALIASYN=''"
    Set Rst = G_FaCn.Execute("SELECT * FROM ACGROUP WHERE GROUPNATURE IS NULL OR GROUPNATURE=''")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write " ***** These ACGROUP GroupNature is Blank/Null *****" + vbCrLf
        varTxtstrm.Write "-----------------------------------------------------" + vbCrLf
        Do Until Rst.EOF
            varTxtstrm.Write FaXNull(Rst!GroupCode) + " " + FaXNull(Rst!GroupName) + vbCrLf
            Rst.MoveNext
        Loop
    End If

   G_FaCn.Execute "UPDATE SUBGROUP SET ACCODE=SUBCODE WHERE ACCODE IS NULL OR ACCODE=''"
    
    G_FaCn.Execute "UPDATE ACGROUP SET TRADINGYN='N' WHERE GROUPNATURE IN ('A','L')"
    
    G_FaCn.Execute "UPDATE ACGROUP SET TRADINGYN='Y' WHERE GROUPNAME='Income (Direct)'"
    G_FaCn.Execute "UPDATE ACGROUP SET TRADINGYN='Y' WHERE GROUPNAME='Expense (Direct)'"

    G_FaCn.Execute "UPDATE ACGROUP SET TRADINGYN='N' WHERE GROUPNAME='Income (Indirect)'"
    G_FaCn.Execute "UPDATE ACGROUP SET TRADINGYN='N' WHERE GROUPNAME='Expense (Indirect)'"

    Set Rst = G_FaCn.Execute("SELECT * FROM ACGROUP WHERE TRADINGYN IS NULL OR TRADINGYN=''")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "-----------------------------------------------------" + vbCrLf
        varTxtstrm.Write "***** These ACGROUP TradingYN is Blank/Null *****" + vbCrLf
        varTxtstrm.Write "-----------------------------------------------------" + vbCrLf
        Do Until Rst.EOF
            varTxtstrm.Write FaXNull(Rst!GroupCode) + " " + FaXNull(Rst!GroupName) + vbCrLf
            Rst.MoveNext
        Loop
    End If
    
    G_FaCn.Execute ("UPDATE ACGROUP SET ALIASYN='N' WHERE ALIASYN IS NULL OR ALIASYN=''")
    
    Set Rst = G_FaCn.Execute("SELECT * FROM ACGROUP WHERE ALIASYN IS NULL OR ALIASYN=''")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "-----------------------------------------------------" + vbCrLf
        varTxtstrm.Write "***** These ACGROUP AliasYN is Blank/Null *****" + vbCrLf
        varTxtstrm.Write "-----------------------------------------------------" + vbCrLf
        Do Until Rst.EOF
            varTxtstrm.Write FaXNull(Rst!GroupCode) + " " + FaXNull(Rst!GroupName) + vbCrLf
            Rst.MoveNext
        Loop
    End If
    
    Set Rst = G_FaCn.Execute("SELECT *  FROM LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE WHERE V_DATE<" & FaConvertDate(PubStartDate) & " AND GROUPNATURE IN ('E','R')")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "***** Opening Balance Exist in Expenditure & Revenue *****" + vbCrLf
        varTxtstrm.Write "-----------------------------------------------------------------------------" + vbCrLf
        varTxtstrm.Write "DocID                              DEBIT   CREDIT  A/C NAME" + vbCrLf
        varTxtstrm.Write "-----------------------------------------------------------------------------" + vbCrLf
        Do Until Rst.EOF
            varTxtstrm.Write FaXNull(Rst!DocID) + " " + FaSetN(FaBNull(Rst!AmtDr), 13) + " " + FaSetN(FaBNull(Rst!AmtCr), 13) + " " + FaXNull(Rst!Name) + vbCrLf
            Rst.MoveNext
        Loop
        varTxtstrm.Write "-----------------------------------------------------------------------------" + vbCrLf
    End If
    
    Set Rst = G_FaCn.Execute("SELECT SUM(AMTCR)-SUM(AMTDR) AS BAL FROM LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE WHERE V_DATE<" & FaConvertDate(PubStartDate) & " AND GROUPNATURE NOT IN ('E','R')")
    If Rst.RecordCount > 0 Then
        If Round(FaVNull(Rst!BAL), 2) <> 0 Then
            varTxtstrm.Write " " + vbCrLf
            varTxtstrm.Write "***** Opening Balance Difference *****" + vbCrLf
            varTxtstrm.Write "-----------------" + vbCrLf
            varTxtstrm.Write "DIFFERANCE AMOUNT" + vbCrLf
            varTxtstrm.Write "-----------------" + vbCrLf
            Do Until Rst.EOF
                varTxtstrm.Write FaSetN(FaBNull(Rst!BAL), 13) + vbCrLf
                Rst.MoveNext
            Loop
            varTxtstrm.Write "-----------------" + vbCrLf
        End If
    End If
    
    Set Rst = G_FaCn.Execute("SELECT * FROM LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE WHERE V_DATE>" & FaConvertDate(PubLoginDate))
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "***** Transaction Exist in BIG Date *****" + vbCrLf
        varTxtstrm.Write "-------------------------------------------------------------" + vbCrLf
        varTxtstrm.Write "DOCID                 DATE                DEBIT        CREDIT" + vbCrLf
        varTxtstrm.Write "-------------------------------------------------------------" + vbCrLf
        Do Until Rst.EOF
            varTxtstrm.Write FaXNull(Rst!DocID) + " " + CStr(Rst!V_Date) + " " + FaSetN(FaBNull(Rst!AmtDr), 13) + " " + FaSetN(FaBNull(Rst!AmtCr), 13) + vbCrLf
            Rst.MoveNext
        Loop
        varTxtstrm.Write "-------------------------------------------------------------" + vbCrLf
    End If
    
    Set Rst = G_FaCn.Execute("SELECT SUM(AMTCR)-SUM(AMTDR) AS BAL FROM LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE WHERE V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(PubLoginDate))
    If Rst.RecordCount > 0 Then
        If Round(FaVNull(Rst!BAL), 2) <> 0 Then
            varTxtstrm.Write " " + vbCrLf
            varTxtstrm.Write "***** Transaction Summary Difference *****" + vbCrLf
            varTxtstrm.Write "------------------" + vbCrLf
            varTxtstrm.Write "DIFF. AMOUNT" + vbCrLf
            varTxtstrm.Write "------------------" + vbCrLf
            Do Until Rst.EOF
                varTxtstrm.Write FaSetN(FaBNull(Rst!BAL), 13) + vbCrLf
                Rst.MoveNext
            Loop
            varTxtstrm.Write "------------------" + vbCrLf
        End If
    End If
    
    Set Rst = G_FaCn.Execute("SELECT V_DATE,SUM(AMTCR)-SUM(AMTDR) AS BAL FROM LEDGER WHERE V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(PubLoginDate) & " GROUP BY V_DATE Having Sum (AmtCr)-Sum(AmtDr) <> 0")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "***** Date Wise Difference in Transaction Sum *****" + vbCrLf
        varTxtstrm.Write "-------------------------" + vbCrLf
        varTxtstrm.Write "DATE              BALANCE" + vbCrLf
        varTxtstrm.Write "-------------------------" + vbCrLf
        Do Until Rst.EOF
            If Round(FaVNull(Rst!BAL), 2) <> 0 Then
                varTxtstrm.Write FaXNull(Rst!V_Date) + " " + FaSetN(FaBNull(Rst!BAL), 13) + vbCrLf
            End If
            Rst.MoveNext
        Loop
        varTxtstrm.Write "-------------------------" + vbCrLf
    End If
    
    Set Rst = G_FaCn.Execute("SELECT DOCID,V_DATE,SUM(AMTCR)-SUM(AMTDR) AS BAL FROM LEDGER WHERE V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(PubLoginDate) & " GROUP BY DOCID,V_DATE Having Sum (AmtCr)-Sum(AmtDr) <> 0")
    If Rst.RecordCount > 0 Then
        varTxtstrm.Write " " + vbCrLf
        varTxtstrm.Write "***** DocId Wise/ Date Wise Difference in Transaction Sum *****" + vbCrLf
        varTxtstrm.Write "-----------------------------------------------" + vbCrLf
        varTxtstrm.Write "DOCID                 DATE              BALANCE" + vbCrLf
        varTxtstrm.Write "-----------------------------------------------" + vbCrLf
        Do Until Rst.EOF
            If Round(FaVNull(Rst!BAL), 2) <> 0 Then
                varTxtstrm.Write FaXNull(Rst!DocID) + " " + FaXNull(Rst!V_Date) + " " + FaSetN(FaBNull(Rst!BAL), 13) + vbCrLf
            End If
            Rst.MoveNext
        Loop
        varTxtstrm.Write "-----------------------------------------------" + vbCrLf
    End If
    varTxtstrm.Close
    MousePointer = vbDefault
    Exit Sub
ERRORHANDLER:               MsgBox err.Description
End Sub
