VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmPriceUpdate 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Price List Updation"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   11505
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Excel File Format"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   90
      TabIndex        =   26
      Top             =   45
      Width           =   2970
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0->No/1-Yes"
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
         Index           =   23
         Left            =   1725
         TabIndex        =   41
         Top             =   2595
         Width           =   1095
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Index           =   22
         Left            =   1740
         TabIndex        =   40
         Top             =   2295
         Width           =   405
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number"
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
         Index           =   21
         Left            =   1740
         TabIndex        =   39
         Top             =   1980
         Width           =   675
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2 Char"
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
         Index           =   20
         Left            =   1740
         TabIndex        =   38
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "40 Char"
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
         Index           =   19
         Left            =   1740
         TabIndex        =   37
         Top             =   1365
         Width           =   690
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "21 Char"
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
         Left            =   1740
         TabIndex        =   36
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Part_No"
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
         Left            =   150
         TabIndex        =   35
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Part_Name"
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
         Left            =   150
         TabIndex        =   34
         Top             =   1365
         Width           =   1170
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Disc_Code"
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
         Index           =   6
         Left            =   150
         TabIndex        =   33
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Rate"
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
         Index           =   7
         Left            =   150
         TabIndex        =   32
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5. Effect_Dt"
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
         Index           =   8
         Left            =   150
         TabIndex        =   31
         Top             =   2295
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6. MRP_YN"
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
         Index           =   9
         Left            =   150
         TabIndex        =   30
         Top             =   2595
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Col. Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   10
         Left            =   195
         TabIndex        =   29
         Top             =   765
         Width           =   825
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Col. Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   11
         Left            =   1740
         TabIndex        =   28
         Top             =   765
         Width           =   765
      End
      Begin VB.Label LblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sheet Name : PartList_New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   12
         Left            =   255
         TabIndex        =   27
         Top             =   360
         Width           =   2220
      End
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Index           =   4
      Left            =   9705
      MaxLength       =   5
      TabIndex        =   10
      Top             =   2693
      Width           =   465
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Index           =   3
      Left            =   9705
      MaxLength       =   5
      TabIndex        =   7
      Top             =   2190
      Width           =   465
   End
   Begin VB.OptionButton OptBaseRate 
      BackColor       =   &H00CFE0E0&
      Caption         =   "List Price"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   1
      Left            =   6060
      TabIndex        =   3
      Top             =   1343
      Width           =   1485
   End
   Begin VB.OptionButton OptBaseRate 
      BackColor       =   &H00CFE0E0&
      Caption         =   "MRP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   0
      Left            =   5190
      TabIndex        =   2
      Top             =   1343
      Width           =   660
   End
   Begin VB.CheckBox ChkTB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00CFE0E0&
      Caption         =   "Taxable Price"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3285
      TabIndex        =   9
      Top             =   2700
      Width           =   2160
   End
   Begin VB.CheckBox ChkTP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00CFE0E0&
      Caption         =   "Taxpaid Price"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3285
      TabIndex        =   8
      Top             =   2430
      Width           =   2160
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   1
      Top             =   405
      Width           =   6870
   End
   Begin VB.CommandButton CmdSel 
      BackColor       =   &H00D3BEC9&
      Caption         =   "Select &Text File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4605
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Exit Form"
      Top             =   5895
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton CmdSel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Select E&xcel File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   8205
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Form"
      Top             =   975
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   5070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.xls"
   End
   Begin VB.CheckBox ChkMRP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00CFE0E0&
      Caption         =   "UMRP Price"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3285
      TabIndex        =   6
      Top             =   2190
      Width           =   2160
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   2850
      TabIndex        =   17
      Top             =   4860
      Visible         =   0   'False
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton BtnUpdate 
      BackColor       =   &H00D3BEC9&
      Caption         =   "&Update Price List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2850
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Exit Form"
      Top             =   5115
      Width           =   2175
   End
   Begin VB.CommandButton BtnExit 
      BackColor       =   &H00D3BEC9&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exit Form"
      Top             =   5115
      Width           =   2175
   End
   Begin VB.CommandButton BtnPrint 
      BackColor       =   &H00D3BEC9&
      Caption         =   "&Print Variation List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5025
      MaskColor       =   &H00800080&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Print Reports"
      Top             =   5115
      Width           =   2175
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Index           =   1
      Left            =   5235
      MaxLength       =   12
      TabIndex        =   4
      Top             =   1650
      Width           =   1605
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Index           =   0
      Left            =   5235
      TabIndex        =   5
      Top             =   1920
      Width           =   1605
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   24
      Left            =   3120
      TabIndex        =   42
      Top             =   150
      Width           =   1305
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Price=List Price)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   18
      Left            =   7110
      TabIndex        =   25
      Top             =   3315
      Width           =   2130
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(if Base Price List=List Price and % is zero then "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   17
      Left            =   7110
      TabIndex        =   24
      Top             =   3000
      Width           =   4125
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conversion % from MRP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   16
      Left            =   7125
      TabIndex        =   23
      Top             =   2700
      Width           =   2100
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conversion % from List Price"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   15
      Left            =   7125
      TabIndex        =   22
      Top             =   2190
      Width           =   2505
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Base Price List ---->"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   14
      Left            =   3315
      TabIndex        =   21
      Top             =   1380
      Width           =   1740
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note : First Row of Sheet just after Heading row must contain Character Values in Part_No,Part_Name && Disc_Code columns."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   13
      Left            =   75
      TabIndex        =   20
      Top             =   4080
      Width           =   10740
   End
   Begin VB.Label LblName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%%%%%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   8460
      TabIndex        =   19
      Top             =   4575
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Updation Status"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   2850
      TabIndex        =   18
      Top             =   4545
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref. No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   3315
      TabIndex        =   16
      Top             =   1650
      Width           =   690
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price Effective Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   4
      Left            =   3315
      TabIndex        =   15
      Top             =   1920
      Width           =   1680
   End
End
Attribute VB_Name = "FrmPriceUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADDFLAG As Byte, mFileName$
Dim RstNew As ADODB.Recordset, RstPart As ADODB.Recordset
Dim mFlag As Byte
Private Const Eff_Date = 0
Private Const RefNo = 1
Private Const FilePath = 2
Private Const ConvPerc = 3
Private Const ConvPerc2 = 4

Dim mGen_Per As Double, mLST_Per As Double, mLST_Sur As Double
Private Const OptMRP As Byte = 0
Private Const OptListPrice As Byte = 1

Private Sub btnexit_Click()
    Unload Me
End Sub

Private Sub BTNPRINT_Click()
On Error GoTo lblErrorBox
Dim mDate As Double
Dim RepFileName$, TTXFileName$
Dim mOldRef$, mNewRef$
Dim mReportCount As Integer
Dim Rst As ADODB.Recordset

    If IsValid(txt(Eff_Date), "Effective Date") = False Then Exit Sub
    
    Set GRs = GCn.Execute("select Distinct Effect_Dt,ref_no From Part_PriceList where Div_Code='" & PubDivCode & "' and Effect_Dt<=" & ConvertDate(txt(Eff_Date)) & " and MRP_Yn=" & ChkMRP.Value & " order By Effect_Dt desc")
    
    If GRs.RecordCount = 0 Then
        MsgBox "No Price List Found for mentioned Effective Date", vbInformation, "Validation"
        Exit Sub
    Else
        If GRs!Effect_Dt <> CDate(txt(Eff_Date).TEXT) Then
            txt(Eff_Date) = GRs!Effect_Dt
        End If
        mNewRef = GRs!Ref_No
    End If
    
    Set GRs = GCn.Execute("select Distinct Effect_Dt,Ref_no From Part_PriceList where Div_Code='" & PubDivCode & "' and Effect_Dt<" & ConvertDate(txt(Eff_Date)) & " and MRP_Yn=" & ChkMRP.Value & " order By Effect_Dt desc")
    If GRs.RecordCount = 0 Then
        MsgBox "No Old Price List Found before mentioned Effective Date", vbInformation, "Validation"
        Exit Sub
    Else
        mDate = GRs!Effect_Dt
        mOldRef = GRs!Ref_No
    End If
    
    If MsgBox("Sure to Print Variation List", vbQuestion + vbYesNo, "Confirmation") = vbNo Then Exit Sub
    
    If ChkMRP.Value = 1 Then
'        GSQL = "select Part.Part_Name,PN.Part_No,PN.New_Part,PN.MRP_YN,PN.MRP as Rate,iif(isnull(PO.MRP),0,PO.MRP) as ORate " & _
'            "From (Part_PriceList PN Left Join Part_PriceList PO on PN.Part_No=PO.Part_No) " & _
'            "Left Join Part on PN.Part_No=Part.Part_No and Part.Div_Code = PN.div_code " & _
'            "where PN.New_Part=0 and PN.Effect_Dt=" & ConvertDate(Txt(Eff_Date)) & " and PO.Effect_Dt=" & ConvertDate(mDate) & " and PN.MRP_Yn=" & ChkMRP.Value & " and PO.MRP_YN=" & ChkMRP.Value & _
'            " Union " & _
'            "select Part.Part_Name,PN.Part_No,PN.New_Part,PN.MRP_YN,PN.MRP as Rate,0 as ORate " & _
'            "From Part_PriceList PN Left Join Part on PN.Part_No=Part.Part_No and Part.Div_Code = PN.div_code " & _
'            "where PN.New_Part=1 and PN.Effect_Dt=" & ConvertDate(Txt(Eff_Date)) & " and PN.MRP_Yn=" & ChkMRP.Value
        GSQL = "select Part.Part_Name,PN.Part_No,PN.New_Part,PN.MRP_YN,PN.MRP as Rate, " & vIsNull("PO.MRP", "0") & " as ORate, " & _
            "PN.Disc_Factor as Disc_Factor, " & xIsNull("PO.Disc_Factor", "") & " as oDisc_Factor,PDFN.PurcDisc_Per,PDFN.SalDisc_Per,PDFO.PurcDisc_Per as oPurcDisc_Per, PDFO.SalDisc_Per as oSalDisc_Per " & _
            "From ((((Part_PriceList PN Left Join Part_PriceList PO on PN.Part_No=PO.Part_No) " & _
            "Left Join Part on PN.Part_No=Part.Part_No and Part.Div_Code = PN.div_code) " & _
            "Left Join Part_DiscFactor PDFN On PDFN.DiscFac_Catg=PN.Disc_Factor) " & _
            "Left Join Part_DiscFactor PDFO On PDFO.DiscFac_Catg=PO.Disc_Factor) " & _
            "where PN.New_Part=0 and PN.Effect_Dt=" & ConvertDate(txt(Eff_Date)) & " and PO.Effect_Dt=" & ConvertDate(mDate) & " and PN.MRP_Yn=" & ChkMRP.Value & " and PO.MRP_YN=" & ChkMRP.Value & _
            " Union " & _
            "select Part.Part_Name,PN.Part_No,PN.New_Part,PN.MRP_YN,PN.MRP as Rate,0 as ORate,PN.Disc_Factor,'' as oDisc_Factor,PDF.PurcDisc_Per,PDF.SalDisc_Per,0 as oPurcDisc_Per,0 as oSalDisc_Per " & _
            "From ((Part_PriceList PN Left Join Part on PN.Part_No=Part.Part_No and Part.Div_Code = PN.div_code) " & _
            "Left Join Part_DiscFactor PDF On PDF.DiscFac_Catg=PN.Disc_Factor) " & _
            "where PN.New_Part=1 and PN.Effect_Dt=" & ConvertDate(txt(Eff_Date)) & " and PN.MRP_Yn=" & ChkMRP.Value
    
    Else
        GSQL = "select Part.Part_Name,PN.Part_No,PN.New_Part,PN.MRP_YN,PN.TB_SRate as Rate, " & vIsNull("PO.TB_SRate", "0") & " as ORate From (Part_PriceList PN Left Join Part_PriceList PO on PN.Part_No=PO.Part_No) Left Join Part on PN.Part_No=Part.Part_No and Part.Div_Code = PN.div_code where PN.New_Part=0 and  PN.Effect_Dt=" & ConvertDate(txt(Eff_Date)) & " and PO.Effect_Dt=" & ConvertDate(mDate) & " and PN.MRP_Yn=" & ChkMRP.Value & " and PO.MRP_YN=" & ChkMRP.Value & _
            " Union " & _
            "select Part.Part_Name,PN.Part_No,PN.New_Part,PN.MRP_YN,PN.TB_SRate as Rate,0 as ORate From Part_PriceList PN Left Join Part on PN.Part_No=Part.Part_No and Part.Div_Code = PN.div_code where PN.New_Part=1 and PN.Effect_Dt=" & ConvertDate(txt(Eff_Date)) & " and PN.MRP_Yn=" & ChkMRP.Value
    End If

    TTXFileName = "PriceListDiff"
    RepFileName = "PriceListDiff"
    Set Rst = GCn.Execute(GSQL)
    
    If Rst.RecordCount > 0 Then
        CreateFieldDefFile Rst, PubRepoPath + "\" & TTXFileName & ".TTX", True
        Set rpt = rdApp.OpenReport(PubRepoPath + "\" & RepFileName & ".RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
        For mReportCount = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(mReportCount).FormulaFieldName)
                Case UCase("NewDate")
                    rpt.FormulaFields(mReportCount).TEXT = "'New Date : " & txt(Eff_Date).TEXT & "'"
                Case UCase("OldDate")
                    rpt.FormulaFields(mReportCount).TEXT = "'Old Date : " & Format(mDate, "dd/MMM/yyyy") & "'"
                Case UCase("UMRP")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & IIf(ChkMRP.Value = 0, "List Price Based Price List", "UMRP Price Based Price List") & "'"
                Case UCase("OldRef")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & mOldRef & "'"
                Case UCase("NewRef")
                    rpt.FormulaFields(mReportCount).TEXT = "'" & mNewRef & "'"
            End Select
        Next
        Call Report_View(rpt, Me.CAPTION)
    Else
        MsgBox "No Records to Print", vbInformation, "Information"
        Exit Sub
    End If
    Set Rst = Nothing
    Exit Sub
lblErrorBox:
    Set Rst = Nothing
    ProcErrorMsg
End Sub


Private Sub BtnUpdate_Click()
'On Error GoTo ErrorTrap
Dim XlsConn As New ADODB.Connection
Dim mTran As Boolean, PartNo$, PartName$
Dim mRate As Single, mTB_SRate As Double, mTP_SRate As Double, mMRP As Double
    
    ''  Changes in Structure Made :
    ''  Table : Part            : Field NewPart - > Default Value changed to 1 From 0
    ''          Part_PriceList  : New Field  :   MRP_YN
    ''                          : New Field  :   New_Part
    ''                          : Change Made in Primary Key
    ''  New Table : PartList_New
    
    If Trim(txt(FilePath)) = "" Then
        MsgBox "File not Selected", vbCritical, "Select Import File"
        CmdSel(0).SetFocus
        Exit Sub
    End If
    If IsValid(txt(Eff_Date), "Effective Date") = False Then Exit Sub
    If IsValid(txt(RefNo), "Ref. No.") = False Then Exit Sub
    If GCn.Execute("select Distinct Effect_Dt From Part_PriceList where Div_Code='" & PubDivCode & "' and Effect_Dt=" & ConvertDate(txt(Eff_Date)) & " and MRP_Yn=" & ChkMRP.Value).RecordCount > 0 Then
        MsgBox "PriceList is already updated", vbInformation, "Validation"
        Exit Sub
    End If
    
    If MsgBox("Start Price List Updation", vbQuestion + vbYesNo, "Confirmation") = vbNo Then Exit Sub
    
    XlsConn.CursorLocation = adUseClient
    XlsConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txt(FilePath) & ";Extended Properties=Excel 8.0"
    XlsConn.Open
''    XlsConn.Provider = "Microsoft.Jet.OLEDB.4.0"
''    XlsConn.ConnectionString = "Data Source=" & Txt(FilePath) & "x" & ";Persist Security Info=False; Extended Properties=Excel 8.0;"
''    XlsConn.Open
'
''
''    XlsConn.ConnectionString = "Data Source=" & Txt(FilePath) & ";Microsoft.Jet.OLEDB.4.0; Extended Properties=Excel 8.0;"
''    XlsConn.Open
'
'
''    Set RstNew = GCn.Execute("select * From  PartList_New")
'Set RstNew = GCn.Execute("SELECT * FROM " & _
'                            "OpenQ('Microsoft.Jet.OLEDB.4.0', " & _
'                            "'Data Source=" & Txt(FilePath) & ";" & _
'                            "Extended Properties=Excel 8.0')...[PartList_New$]")
'
'    Set RstNew = GCn.Execute("SELECT * FROM OPENROWSET('Microsoft.Jet.OLEDB.4.0','Excel 8.0;Database=" & Txt(FilePath) & "', 'SELECT * FROM [PartList_New$]')")

    
    
    'Set RstNew = XlsConn.Execute("select distinct * From [PartList_New$] IN '" & Txt(FilePath) & "' 'EXCEL 8.0;' where Part_No <> '' and Part_Name <> '' ORDER BY 1;")
    Set RstNew = XlsConn.Execute("select distinct * From [PartList_New$]  where Part_No <> '' and Part_Name <> '' ORDER BY 1")
    
    
    Set RstPart = GCn.Execute("select Part_No,Photo From Part where Div_Code='" & PubDivCode & "' Order By Part_No")
    
    If RstNew.RecordCount = 0 Then MsgBox "No Data found for Updation in New Price List Table", vbInformation, "Validation": Exit Sub
    
    GCn.BeginTrans
    mTran = True
    
    GCn.Execute ("Update part set New_Part=0 Where Div_Code='" & PubDivCode & "'")
    
    With ProgressBar1
        .Min = 0
        .Max = 100
        .Value = 0
        .Visible = True
    End With
    LblName(2).Visible = True
    LblName(0).Visible = True
    Do While Not RstNew.EOF
        ProgressBar1.Value = Round(RstNew.AbsolutePosition * 100 / RstNew.RecordCount, 0)
        ProgressBar1.Refresh
        LblName(0).CAPTION = ProgressBar1.Value & "%"
        LblName(0).Refresh
        '*********
        PartNo = Replace(IIf(IsNull(RstNew!Part_No), "", RstNew!Part_No), "'", "`")
        PartName = Replace(IIf(IsNull(RstNew!Part_Name), "", RstNew!Part_Name), "'", "`")
        If Len(PartNo) > 40 Then
            PartNo = left(PartNo, 22)
        End If
        If Len(PartName) > 40 Then
            PartName = left(PartName, 40)
        End If
        If RstPart.RecordCount > 0 Then
            RstPart.MoveFirst
            RstPart.FIND ("Part_No='" & PartNo & "'")
        End If
        

        If IsNull(RstNew!Rate) Then
            mRate = 0
        Else
            mRate = Val(Replace(RstNew!Rate, ",", ""))
        End If
        mMRP = 0
        mTB_SRate = 0
        mTP_SRate = 0
        If OptBaseRate(OptMRP) Then
            If ChkMRP.Value = 1 Then mMRP = mRate
            If ChkTP.Value = 1 Then mTP_SRate = mRate
        Else
            If ChkTP.Value = 1 Then mTP_SRate = ConvertTPRate(mRate)
            If ChkMRP.Value = 1 Then
                mMRP = Round(mRate + (mRate * Val(txt(ConvPerc)) / 100), 2)
                If ChkTP.Value = 1 Then mTP_SRate = mMRP
            End If
        End If
        If ChkTB.Value = 1 Then
           If OptBaseRate(OptListPrice) And Val(txt(ConvPerc2)) = 0 Then
                mTB_SRate = mRate
            Else
                mTB_SRate = Round(mMRP - ((mMRP * Val(txt(ConvPerc2))) / (100 + Val(txt(ConvPerc2)))), 2)
            End If
        End If
       
        If RstPart.EOF = True Then      '' not found
                '' Insert Record in Part
                GSQL = "insert into Part(" & _
                    "Div_Code,Site_Code,Part_No,Part_Name,Local_Name,Part_NoHelp,Part_NameHelp," & _
                    "Part_Grade,Photo,Security_Grade,Value_Method,Disc_Factor,MRP,MRP_Effect_Dt," & _
                    "TB_SRate,TP_SRate,TB_Effect_Dt," & _
                    "New_Part,U_Name, U_EntDt, U_AE) " & _
                    " values(" & _
                    "'" & PubDivCode & "','" & PubSiteCode & "','" & PartNo & "','" & PartName & "','" & PartName & "','" & Replace(PartNo, " ", "") & "','" & Replace(PartName, " ", "") & _
                    "','S',' ','A','FIFO','" & IIf(IsNull(RstNew!Disc_code), "X", RstNew!Disc_code) & "'," & mMRP & "," & IIf(ChkMRP.Value = 1, ConvertDate(txt(Eff_Date)), "Null") & _
                    "," & mTB_SRate & "," & mTP_SRate & "," & IIf(ChkTP.Value = 1 Or ChkTB.Value = 1, ConvertDate(txt(Eff_Date)), "Null") & _
                    ",1,'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
            Else
                '' Update Record in Part
                GSQL = "Update Part set Disc_Factor='" & IIf(IsNull(RstNew!Disc_code), "X", RstNew!Disc_code) & "',Photo ='" & IIf(IsNull(RstPart!Photo), " ", RstPart!Photo) & "',"
               ' If ChkMRP.Value = 1 Then 'MRP
               '     GSQL = GSQL & "MRP=" & mMRP & ",MRP_Effect_Dt=" & ConvertDate(txt(Eff_Date)) & ","
               '     If ChkTP.Value = 1 Then 'Taxpaid
               '         GSQL = GSQL & "TP_SRate=" & mTP_SRate & ",TB_Effect_Dt=" & ConvertDate(txt(Eff_Date)) & ","
               '     End If
               ' Else
                    If ChkMRP.Value = 1 Then 'MRP
                        GSQL = GSQL & "MRP=" & mMRP & ",MRP_Effect_Dt=" & ConvertDate(txt(Eff_Date)) & ","
                    End If
                    If ChkTB.Value = 1 Then 'Taxable Rate
                        GSQL = GSQL & "TB_SRate=" & mTB_SRate & ",TB_Effect_Dt=" & ConvertDate(txt(Eff_Date)) & ","
                    End If
                    If ChkTP.Value = 1 Then 'Taxpaid converted from TB Rate
                        GSQL = GSQL & "TP_SRate=" & mTP_SRate & ","
                    End If
              'End If
                GSQL = GSQL & "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'" & _
                    "where part_no='" & PartNo & "' and Div_Code='" & PubDivCode & "'"
            End If
        GCn.Execute GSQL
        '' Insert Record in Part_PriceList
        GSQL = "insert into Part_PriceList(" _
                & "Div_Code,Site_Code,Part_No," _
                & "MRP_YN,Ref_No,New_Part,Effect_Dt," _
                & "MRP,TB_SRate,TP_SRate,Disc_Factor," _
                & " U_Name, U_EntDt, U_AE) " _
                & " values(" _
                & "'" & PubDivCode & "','" & PubSiteCode & "','" & PartNo & "'," _
                & ChkMRP.Value & ",'" & txt(RefNo) & "'," & IIf(RstPart.EOF = True, 1, 0) & "," & ConvertDate(txt(Eff_Date)) & _
                "," & mMRP & "," & mTB_SRate & "," & mTP_SRate & ",'" & IIf(IsNull(RstNew!Disc_code), "X", RstNew!Disc_code) & "'," _
                & "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
        GCn.Execute GSQL
        RstNew.MoveNext
    Loop
    
        If PubBackEnd = "S" Then
            GCn.Execute "Update Part Set NDP = Mrp-(MRP*PurcDisc_Per/100) From Part_DiscFactor Where Part.Disc_Factor=Part_DiscFactor.DiscFac_CatG"
        Else
            GCn.Execute "Update Part, Part_DiscFactor Set NDP = Mrp-(MRP*PurcDisc_Per/100)  Where Part.Disc_Factor=Part_DiscFactor.DiscFac_CatG"
        End If
    
    GCn.CommitTrans
    mTran = False
    MsgBox "Price List Updated Successfully!", vbOKOnly, "Validation"
ErrorTrap:
    If mTran = True Then GCn.RollbackTrans
    If err.NUMBER <> 0 Then MsgBox err.Description, vbCritical, "Error Message"
End Sub

Private Sub ChkMRP_Click()
If OptBaseRate(OptMRP).Value Then
    txt(ConvPerc) = ""
    txt(ConvPerc).Enabled = False
Else
    If ChkMRP.Value = 1 Then
        txt(ConvPerc).Enabled = True
    Else
        txt(ConvPerc) = ""
        txt(ConvPerc).Enabled = False
    End If
End If
End Sub

Private Sub CmdSel_Click(Index As Integer)
On Error GoTo ErrHandler
    mFileName = ""
  CommonDialog1.InitDir = Pub_DataPath
  ' Set CancelError is True
  CommonDialog1.CancelError = True
  CommonDialog1.DialogTitle = "Select XLS File for Price List Updation"
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
'    FormKeyDown Me, KeyCode, Shift
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
Call WinSetting(Me)
    'Me.top = 0: Me.left = 0
    Set GRs = GCn.Execute("select GenSurChrgOnSpr,TF.Tax_Per,TF.Tax_Sur_Per From (Syctrl left join TaxForms as TF on Syctrl.LocalTaxFormSpr=TF.Form_Code)")
    If GRs.RecordCount > 0 Then
        mGen_Per = IIf(IsNull(GRs!GenSurChrgOnSpr), 0, GRs!GenSurChrgOnSpr)
        mLST_Per = IIf(IsNull(GRs!Tax_Per), 0, GRs!Tax_Per)
        mLST_Sur = IIf(IsNull(GRs!Tax_Sur_Per), 0, GRs!Tax_Sur_Per)
    Else
        mGen_Per = 0
        mLST_Per = 0
        mLST_Sur = 0
    End If
    Disp_Text True
    CtrlClckCol
    ADDFLAG = 0:    mFlag = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form_Unload (True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RstNew = Nothing: Set RstPart = Nothing
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
    txt(Eff_Date).BackColor = CtrlBColOrg:
    txt(RefNo).BackColor = CtrlBColOrg:
End Sub

Private Sub OptBaseRate_Click(Index As Integer)
If Index = OptMRP Then 'MRP
    ChkMRP.Value = 1
    ChkTP.Value = 0
    ChkTB.Value = 0
'    ChkTB.Enabled = False
    txt(ConvPerc).Enabled = False
ElseIf Index = OptListPrice Then
    ChkMRP.Value = 0
    ChkTP.Value = 0
    ChkTB.Value = 1
'    ChkTB.Enabled = True
End If
txt(ConvPerc) = ""
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus txt(Index)
    txt(Index).Tag = txt(Index)
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> Eff_Date Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf Index <> RefNo Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
    End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case Eff_Date
            txt(Eff_Date) = RetDate(txt(Index))
    End Select
Set Rst = Nothing
End Sub

Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
Next I
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
Next
txt(ConvPerc).Enabled = False
End Sub

Private Function ConvertTPRate(ByVal mRate As Double) As Double
Dim xRate As Double
    mRate = mRate + Round(mRate * mGen_Per / 100, 2)
    xRate = Round(mRate * mLST_Per / 100, 2)
    mRate = mRate + xRate + Round(xRate * mLST_Sur / 100, 2)
    ConvertTPRate = mRate
End Function
