VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FaReports 
   BackColor       =   &H00DFE7C0&
   Caption         =   "ReprtForm"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   11820
   Begin VB.Frame FrDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7C5F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   2520
      TabIndex        =   17
      Top             =   1530
      Visible         =   0   'False
      Width           =   8160
      Begin VB.TextBox TEXT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   13
         Left            =   6795
         MaxLength       =   5
         TabIndex        =   38
         Text            =   "Yes"
         ToolTipText     =   "Enter the Interest Rate."
         Top             =   105
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox TEXT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   12
         Left            =   4770
         MaxLength       =   20
         TabIndex        =   39
         ToolTipText     =   "Enter the Interest Rate."
         Top             =   405
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.TextBox TEXT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   11
         Left            =   4770
         MaxLength       =   5
         TabIndex        =   37
         Text            =   "Yes"
         ToolTipText     =   "Enter the Interest Rate."
         Top             =   105
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox TEXT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   10
         Left            =   4230
         MaxLength       =   5
         TabIndex        =   40
         Text            =   ">="
         ToolTipText     =   "Enter the Interest Rate."
         Top             =   1305
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox TEXT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   1245
         MaxLength       =   5
         TabIndex        =   31
         Text            =   ">="
         ToolTipText     =   "Enter the Interest Rate."
         Top             =   405
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox TEXT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   1245
         MaxLength       =   5
         TabIndex        =   29
         Text            =   ">="
         ToolTipText     =   "Enter the Interest Rate."
         Top             =   105
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton BtnClose 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Close"
         DownPicture     =   "FaReports.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   43
         ToolTipText     =   "Print Report"
         Top             =   840
         Width           =   1620
      End
      Begin VB.TextBox TEXT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   1725
         MaxLength       =   5
         TabIndex        =   36
         Text            =   "No"
         ToolTipText     =   "Enter the Interest Rate."
         Top             =   1605
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox TEXT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   4695
         MaxLength       =   5
         TabIndex        =   42
         Text            =   "10"
         ToolTipText     =   "Enter the Interest Rate."
         Top             =   1605
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox TEXT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   1725
         TabIndex        =   30
         Text            =   "0"
         Top             =   105
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox TEXT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   1725
         TabIndex        =   32
         Text            =   "0"
         Top             =   405
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox TEXT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   4695
         TabIndex        =   41
         Text            =   "0"
         Top             =   1305
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox TEXT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   1725
         TabIndex        =   33
         Top             =   705
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.TextBox TEXT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   1725
         TabIndex        =   34
         Top             =   1005
         Visible         =   0   'False
         Width           =   4440
      End
      Begin VB.TextBox TEXT 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   1725
         MaxLength       =   5
         TabIndex        =   35
         Text            =   "No"
         ToolTipText     =   "Enter the Interest Rate."
         Top             =   1305
         Visible         =   0   'False
         Width           =   1395
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2310
         Left            =   45
         TabIndex        =   18
         Top             =   1995
         Visible         =   0   'False
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   4075
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Header Remarks"
         TabPicture(0)   =   "FaReports.frx":3132
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "TxtHeader(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "TxtHeader(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "TxtHeader(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "TxtHeader(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "TxtHeader(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Footer Remarks"
         TabPicture(1)   =   "FaReports.frx":314E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TxtFooter(0)"
         Tab(1).Control(1)=   "TxtFooter(1)"
         Tab(1).Control(2)=   "TxtFooter(2)"
         Tab(1).Control(3)=   "TxtFooter(3)"
         Tab(1).Control(4)=   "TxtFooter(4)"
         Tab(1).ControlCount=   5
         Begin VB.TextBox TxtHeader 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   15
            MaxLength       =   75
            TabIndex        =   28
            Top             =   1785
            Width           =   7980
         End
         Begin VB.TextBox TxtHeader 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   15
            MaxLength       =   75
            TabIndex        =   27
            Top             =   1455
            Width           =   7980
         End
         Begin VB.TextBox TxtHeader 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   15
            MaxLength       =   75
            TabIndex        =   26
            Top             =   1125
            Width           =   7980
         End
         Begin VB.TextBox TxtHeader 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   15
            MaxLength       =   75
            TabIndex        =   25
            Top             =   795
            Width           =   7980
         End
         Begin VB.TextBox TxtHeader 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   15
            MaxLength       =   75
            TabIndex        =   24
            Top             =   465
            Width           =   7980
         End
         Begin VB.TextBox TxtFooter 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   -74985
            MaxLength       =   75
            TabIndex        =   23
            Top             =   1785
            Width           =   7980
         End
         Begin VB.TextBox TxtFooter 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   -74985
            MaxLength       =   75
            TabIndex        =   22
            Top             =   1455
            Width           =   7980
         End
         Begin VB.TextBox TxtFooter 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   -74985
            MaxLength       =   75
            TabIndex        =   21
            Top             =   1125
            Width           =   7980
         End
         Begin VB.TextBox TxtFooter 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   -74985
            MaxLength       =   75
            TabIndex        =   20
            Top             =   795
            Width           =   7980
         End
         Begin VB.TextBox TxtFooter 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   -74985
            MaxLength       =   75
            TabIndex        =   19
            Top             =   465
            Width           =   7980
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print Vr.No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   7
         Left            =   5790
         TabIndex        =   63
         Top             =   150
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For Site"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   6
         Left            =   3360
         TabIndex        =   62
         Top             =   450
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print Narration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   5
         Left            =   3360
         TabIndex        =   60
         Top             =   150
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Narration in Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   53
         Top             =   1650
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   3180
         TabIndex        =   50
         Top             =   1650
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Narr. Not Having"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1050
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Narr. Having"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   750
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dr.Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   150
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cr.Balance "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   450
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TxN.Amt.  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   45
         Top             =   1350
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "As Per Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   44
         Top             =   1350
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   3330
      Left            =   6120
      Negotiate       =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12648447
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Code"
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
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "Site Name"
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGGroup 
      Height          =   3330
      Left            =   3600
      Negotiate       =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12648447
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Code"
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
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "Account Group"
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   210
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   4005
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   3
      Left            =   6585
      TabIndex        =   57
      Top             =   45
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton BtnParam 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Parameters"
      DownPicture     =   "FaReports.frx":316A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   56
      ToolTipText     =   "Print Report"
      Top             =   0
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Dos Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3015
      TabIndex        =   55
      ToolTipText     =   "Print Report"
      Top             =   6090
      Width           =   1620
   End
   Begin VB.TextBox TXTACC_CODE1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10290
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2295
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DGAccount 
      Height          =   3300
      Left            =   10200
      Negotiate       =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   5821
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12648447
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Code"
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
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "Account Name"
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
      BeginProperty Column02 
         DataField       =   "Address"
         Caption         =   "Address"
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
      BeginProperty Column03 
         DataField       =   "CityName"
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
      BeginProperty Column04 
         DataField       =   "GroupName"
         Caption         =   "Group Name"
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
            DividerStyle    =   1
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   1
            ColumnWidth     =   3105.071
         EndProperty
         BeginProperty Column02 
            DividerStyle    =   1
            ColumnWidth     =   4500.284
         EndProperty
         BeginProperty Column03 
            DividerStyle    =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            DividerStyle    =   1
            ColumnWidth     =   1995.024
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TXTACC_CODE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10095
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   7020
      TabIndex        =   14
      Top             =   210
      Visible         =   0   'False
      Width           =   4560
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   3228
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3942
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.TextBox TXTS_DATE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10185
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox TXTE_DATE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10185
      TabIndex        =   12
      Top             =   1410
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton BTNEXIT 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&xit"
      DownPicture     =   "FaReports.frx":629C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6255
      TabIndex        =   6
      ToolTipText     =   "Exit Form"
      Top             =   6090
      Width           =   1620
   End
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Windows &Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4635
      TabIndex        =   5
      ToolTipText     =   "Print Report"
      Top             =   6090
      Width           =   1620
   End
   Begin VB.PictureBox Pic 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11820
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6045
      Width           =   11820
      Begin VB.Label LblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   7230
         TabIndex        =   11
         Top             =   0
         Width           =   4470
      End
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      HideSelection   =   0   'False
      Left            =   840
      TabIndex        =   8
      Top             =   1470
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   2
      Left            =   5010
      TabIndex        =   3
      Top             =   1830
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1845
      Visible         =   0   'False
      Width           =   915
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   90
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   285
      Visible         =   0   'False
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   661
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   1785
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2910
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   12632256
      ForeColorSel    =   128
      BackColorBkg    =   14673856
      GridColor       =   8438015
      GridColorFixed  =   192
      GridColorUnpopulated=   16711935
      AllowUserResizing=   1
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   2
      Left            =   4920
      TabIndex        =   4
      Top             =   1785
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   2910
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   12632256
      ForeColorSel    =   128
      BackColorBkg    =   14673856
      GridColor       =   8438015
      GridColorFixed  =   192
      GridColorUnpopulated=   16711935
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1560
      Left            =   1575
      TabIndex        =   0
      Top             =   60
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2752
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16512
      Rows            =   5
      Cols            =   3
      FixedRows       =   0
      BackColorFixed  =   14673856
      ForeColorFixed  =   16384
      BackColorSel    =   16711680
      ForeColorSel    =   12648447
      BackColorBkg    =   14673856
      GridColor       =   13166810
      GridColorFixed  =   13166810
      GridColorUnpopulated=   12648447
      GridLinesFixed  =   1
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   3
      Left            =   6540
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   2910
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   12632256
      ForeColorSel    =   128
      BackColorBkg    =   14873572
      GridColor       =   8438015
      GridColorFixed  =   192
      GridColorUnpopulated=   16711935
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   1170
      TabIndex        =   9
      Top             =   1230
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "FaReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GridRowHeight As Integer = 270
Public GRepFormName As String
Private Const CellBackColLeave As String = &HFFFFFF, CellBackColEnter As String = &HFFFFC0
Private Const CellBackColLeave1 As String = &HEDF7FE, CellBackColEnter1 As String = &HFFFFC0
Private Const Date1 As Byte = 0, Date2 As Byte = 1, List1 As Byte = 2, List2 As Byte = 3, List3 As Byte = 4
Private Const Cat1 As Byte = 5, Cat2 As Byte = 6, Cat3 As Byte = 7, Cat4 As Byte = 8, Cat5 As Byte = 9
Private Const DrBalance As Byte = 0, CrBalance As Byte = 1, TxNAmount As Byte = 2, NarrationHaving As Byte = 3, NarrationNotHaving As Byte = 4
Private Const Interest As Byte = 5, AsPerDetail As Byte = 6, AsPerDetailNarration As Byte = 7
Private Const SignDr As Byte = 8, SignCr As Byte = 9, SignAmt As Byte = 10, PrintNarration As Byte = 11, SiteCode As Byte = 12, PrintVrNo As Byte = 13
Dim RsGrid1 As ADODB.Recordset, RsGrid2 As ADODB.Recordset, RsGrid3 As ADODB.Recordset, RstEnviro As ADODB.Recordset, RstGroup As ADODB.Recordset, RstAccount As ADODB.Recordset, RstSite As ADODB.Recordset
Dim GridString1 As String, GridString2 As String, GridString3 As String, GridRow1() As Integer, GridRow2() As Integer, GridRow3() As Integer
Dim FormulaString1 As String, FormulaString2 As String, FormulaString3 As String
Dim RepTitle As String, RepName As String, SpeedPrn As Boolean, RepPrint As Boolean, TOT_AMTDR As Double
Private PubDatamanFa As New DMFa.ClsFa
Dim mLastRow As Integer, mFirstRow As Integer, mHelpGridNo, GridKey As Integer, TAddMode As Boolean
Dim ListArray As Variant, mGridStartRow As Integer, mGridEndRow As Integer, mListItem As ListItem

Private Sub BTNPRINT_Click(Index As Integer)
'On Error GoTo ERRORHANDLER
RepPrint = True
Select Case GRepFormName
    Case "DayBook"
        DayBook Index
    Case "Led"
        Led Index, "Led"
    Case "LedInt"
        Led Index, "LedInt"
    Case "CashBook"
        CashBook Index
    Case "BankBook"
        CashBook Index
    Case "JournalBook"
        JournalBook Index
    Case "Annexure"
        Annexure Index
    Case "AgingDr"
        Aging Index, "Dr"
    Case "AgingCr"
        Aging Index, "Cr"
    Case "BankReg"
        BankReg Index
    Case "Clg"
        Clg Index
    Case "ClgNot"
        ClgNot Index
    Case "LedDeb"
        Led Index, "LedDeb"
    Case "LedCred"
        Led Index, "LedCred"
    Case "DailySumm"
        DailySumm Index
    Case "NonTrans"
        NonTrans Index
    Case "DetailedTrial"
        DetailedTrial Index
    Case "RefReport"
        RefReport Index
    Case "AcCheckList"
        Led Index, "AcCheckList"
    Case "DUELIST"
        DUELIST
    Case "CONTROLLED"
        ControlLed Index
    Case "RozNamcha"
        RozNamcha Index
End Select
If RepPrint = False Then Exit Sub
If Index = 1 Then
    
Else
    Select Case GRepFormName
        Case "AgingDr", "AgingCr"
       
        Case Else
            Formulas
    End Select
    rpt.ReadRecords
    FaReport_View rpt, Index, RepTitle, True
End If
SpeedPrn = False
Set rpt = Nothing
Exit Sub
ERRORHANDLER:       MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub DGGroup_Click()
    DGGroup.Visible = False
    If RstGroup.RecordCount > 0 Then
        TxtGrid(Val(DGGroup.Tag)).Tag = RstGroup!Code
        TxtGrid(Val(DGGroup.Tag)).TEXT = RstGroup!Name
    End If
    TxtGrid(Val(DGGroup.Tag)).SetFocus
End Sub
Private Sub DGSite_Click()
    If GRepFormName = "DayBook" Or GRepFormName = "JournalBook" Then
        DGSite.Visible = False
        If RstSite.RecordCount > 0 Then
            TxtGrid(Val(DGSite.Tag)).Tag = RstSite!Code
            TxtGrid(Val(DGSite.Tag)).TEXT = RstSite!Name
        End If
        TxtGrid(Val(DGSite.Tag)).SetFocus
    Else
        DGSite.Visible = False
        If RstSite.RecordCount > 0 Then
            TEXT(SiteCode).Tag = RstSite!Code
            TEXT(SiteCode).TEXT = RstSite!Name
        End If
        TEXT(SiteCode).SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FaFormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:     If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub GridSel_RowColChange(Index As Integer)
    If Index = 1 Then
        If GRepFormName = "Led" Or GRepFormName = "LedInt" Or GRepFormName = "LedDeb" Or GRepFormName = "LedCred" Or GRepFormName = "AcCheckList" Then TxtDetails = GridSel(1).TextMatrix(GridSel(1).Row, 3) + vbCrLf + GridSel(1).TextMatrix(GridSel(1).Row, 4)
    End If
End Sub
Private Sub btnexit_Click()
    Unload Me
End Sub
Private Sub Check1_Click(Index As Integer)
    If Check1(Index).Value = Unchecked Then
        GridSel(Index).Enabled = True
        If GridSel(Index).Rows > 1 Then
            GridSel(Index).Row = 1
            GridSel(Index).Col = 1
        End If
    Else
        GridSel(Index).Enabled = False
        If GridSel(Index).Rows > 1 Then
            GridSel(Index).Row = 0
            GridSel(Index).Col = 0
            GridSel(Index).RowSel = GridSel(Index).Rows - 1
        End If
    End If
End Sub
Private Sub Check1_GotFocus(Index As Integer)
    Check1(Index).BackColor = &HFF&
End Sub
Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub
Private Sub Check1_Validate(Index As Integer, Cancel As Boolean)
Check1(Index).BackColor = &H800000
End Sub
Private Sub Form_Load()
On Error GoTo ELoop
    Me.left = 0
    Me.top = 0
    Me.width = 11900
    Me.height = 7085
    Global_Grid
    TopCtrl1.TopText2 = "Add"
    ''''''''''''''''''''''
    PubDatamanFa.FaBackEnd = PubBackEnd
    PubDatamanFa.FaPubLoginDate = PubLoginDate
    PubDatamanFa.FaPubDivCode = PubDivCode
    PubDatamanFa.FaPubSiteCode = PubSiteCode
    PubDatamanFa.FaPubSiteCodeDisplay = PubSiteCodeDisplay
    PubDatamanFa.FaPubSiteName = PubSiteName
    PubDatamanFa.FapubUName = pubUName
    PubDatamanFa.FaDosPort = PubFaDosPort
    PubDatamanFa.FaRunPIF = PubRunPIF
    PubDatamanFa.FaPubSiteType = PubFaSiteType
    Set PubDatamanFa.SetG_FaCn = G_FaCn
    Set PubDatamanFa.SetG_CompCn = G_CompCn
    Set PubDatamanFa.SetrsUserPerm = rsUserPerm.Clone
    Set PubDatamanFa.SetMasterRst = FaMasterRst.Clone
    ''''''''''''''''''''''
    Set RstEnviro = G_FaCn.Execute("SELECT * FROM FAENVIRO")
    If RstEnviro.RecordCount > 0 Then
        TxtHeader(0) = FaXNull(RstEnviro!TagadaHeader1)
        TxtHeader(1) = FaXNull(RstEnviro!TagadaHeader2)
        TxtHeader(2) = FaXNull(RstEnviro!TagadaHeader3)
        TxtHeader(3) = FaXNull(RstEnviro!TagadaHeader4)
        TxtHeader(4) = FaXNull(RstEnviro!TagadaHeader5)
        TxtFooter(0) = FaXNull(RstEnviro!TagadaFooter1)
        TxtFooter(1) = FaXNull(RstEnviro!TagadaFooter2)
        TxtFooter(2) = FaXNull(RstEnviro!TagadaFooter3)
        TxtFooter(3) = FaXNull(RstEnviro!TagadaFooter4)
        TxtFooter(4) = FaXNull(RstEnviro!TagadaFooter5)
    End If
    Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set RsGrid1 = Nothing
Set RsGrid2 = Nothing
Set mListItem = Nothing
Set RstEnviro = Nothing
Set RstGroup = Nothing
Set RstAccount = Nothing
Set RstSite = Nothing
Set rpt = Nothing
Set PubDatamanFa = Nothing
End Sub
Private Sub GridSel_Click(Index As Integer)
    If Index = 1 Then
        If GRepFormName = "Led" Or GRepFormName = "LedInt" Or GRepFormName = "LedDeb" Or GRepFormName = "LedCred" Or GRepFormName = "AcCheckList" Then TxtDetails = GridSel(1).TextMatrix(GridSel(1).Row, 3) + vbCrLf + GridSel(1).TextMatrix(GridSel(1).Row, 4)
    End If
End Sub
Private Sub GridSel_EnterCell(Index As Integer)
    GridSel(Index).CellBackColor = CellBackColEnter1
End Sub
Private Sub GridSel_GotFocus(Index As Integer)
    GridSel(Index).CellBackColor = CellBackColEnter1
End Sub
Private Sub GridSel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Integer
If KeyCode = 13 Then SendKeys vbTab
If GridSel(Index).Rows < 1 Then Exit Sub
If KeyCode = vbKeySpace And GridSel(Index).Col = 0 Then
    GridSel(Index).CellFontName = "WINGDINGS"
    GridSel(Index).CellFontSize = 14
    GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = IIf(GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "ü", " ", "ü")
    Select Case Index
        Case 1
            I = UBound(GridRow1) + 1
            ReDim Preserve GridRow1(I)
            GridRow1(I) = GridSel(Index).Row
        Case 2
            I = UBound(GridRow2) + 1
            ReDim Preserve GridRow2(I)
            GridRow2(I) = GridSel(Index).Row
        Case 3
            I = UBound(GridRow3) + 1
            ReDim Preserve GridRow3(I)
            GridRow3(I) = GridSel(Index).Row
    End Select
End If
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyPageDown, vbKeyPageUp
        If GRepFormName = "Led" Or GRepFormName = "LedInt" Or GRepFormName = "LedDeb" Or GRepFormName = "LedCred" Or GRepFormName = "AcCheckList" Then TxtDetails = GridSel(1).TextMatrix(GridSel(1).Row, 3) + vbCrLf + GridSel(1).TextMatrix(GridSel(1).Row, 4)
End Select
End Sub
Private Sub GridSel_KeyPress(Index As Integer, KeyAscii As Integer)
If GridSel(Index).Col = 0 Or GridSel(Index).Row = 0 Then Exit Sub
Select Case Index
    Case 1
       FaSelGridKeyPress TxtSearch, GridSel(Index), RsGrid1, KeyAscii, RsGrid1.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1: KeyAscii = 0
       If GRepFormName = "Led" Or GRepFormName = "LedInt" Or GRepFormName = "LedDeb" Or GRepFormName = "LedCred" Or GRepFormName = "AcCheckList" Then TxtDetails = GridSel(1).TextMatrix(GridSel(1).Row, 3) + vbCrLf + GridSel(1).TextMatrix(GridSel(1).Row, 4)
    Case 2
       FaSelGridKeyPress TxtSearch, GridSel(Index), RsGrid2, KeyAscii, RsGrid2.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1: KeyAscii = 0
    Case 3
       FaSelGridKeyPress TxtSearch, GridSel(Index), RsGrid3, KeyAscii, RsGrid3.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1: KeyAscii = 0
End Select
TxtSearch.Tag = Index
End Sub
Private Sub TxtSearch_Click()
TxtSearch.TEXT = ""
GridSel(Val(TxtSearch.Tag)).SetFocus
TxtSearch.Visible = False
End Sub
Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If FaNavigationKey(KeyCode) = True Then GridSel(Val(TxtSearch.Tag)).SetFocus: TxtSearch.Visible = False
If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then GridSel(Val(TxtSearch.Tag)).SetFocus: TxtSearch.Visible = False
If GRepFormName = "Led" Or GRepFormName = "LedInt" Or GRepFormName = "LedDeb" Or GRepFormName = "LedCred" Or GRepFormName = "AcCheckList" Then TxtDetails = GridSel(1).TextMatrix(GridSel(1).Row, 3) + vbCrLf + GridSel(1).TextMatrix(GridSel(1).Row, 4)
End Sub
Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
GridSel_KeyPress Val(TxtSearch.Tag), KeyAscii
End Sub
Private Sub TxtSearch_LostFocus()
    TxtSearch.TEXT = ""
    TxtSearch.Visible = False
End Sub
Private Sub GridSel_LeaveCell(Index As Integer)
GridSel(Index).CellBackColor = CellBackColLeave1
End Sub
Private Sub GridSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If GridSel(Index).Col <> 0 Then Exit Sub
mGridStartRow = GridSel(Index).Row
End Sub
Private Sub GridSel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer, J As Integer
If GridSel(Index).Col <> 0 Or mGridStartRow = 0 Then Exit Sub
mGridEndRow = GridSel(Index).RowSel
For J = mGridStartRow To mGridEndRow
    GridSel(Index).Row = J
    GridSel(Index).Col = 0
    GridSel(Index).CellFontName = "WINGDINGS"
    GridSel(Index).CellFontSize = 14
    GridSel(Index).TextMatrix(J, 0) = IIf(GridSel(Index).TextMatrix(J, 0) = "ü", " ", "ü")
    Select Case Index
        Case 1
            I = UBound(GridRow1) + 1
            ReDim Preserve GridRow1(I)
            GridRow1(I) = GridSel(Index).Row
        Case 2
            I = UBound(GridRow2) + 1
            ReDim Preserve GridRow2(I)
            GridRow2(I) = GridSel(Index).Row
        Case 3
            I = UBound(GridRow3) + 1
            ReDim Preserve GridRow3(I)
            GridRow3(I) = GridSel(Index).Row
    End Select
Next
mGridStartRow = 0
End Sub
Private Sub GridSel_Validate(Index As Integer, Cancel As Boolean)
    GridSel(Index).CellBackColor = CellBackColLeave1
End Sub
Private Sub ListView_Click()
If FrDetail.Visible = True Then
    TEXT(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    TEXT(Val(ListView.Tag)).SetFocus
Else
    TxtGrid(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    TxtGrid(Val(ListView.Tag)).SetFocus
End If
End Sub
Private Sub TxtGrid_GotFocus(Index As Integer)
Dim Rst1 As ADODB.Recordset
FGrid.CellBackColor = CellBackColLeave
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
Select Case FGrid.Row
    Case List1
        Select Case GRepFormName
            Case "DUELIST"
                ListArray = Array("All", "Pending")
                Set mListItem = FaListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                ListArray = Array("Selected", "Merge", "Group(Selected)", "Group(Range)", "City Wise")
                Set mListItem = FaListView_Items(ListView, TxtGrid, Index, ListArray, 5)
            Case "CashBook", "BankBook", "DayBook", "RozNamcha", "DetailedTrial"
                ListArray = Array("Yes", "No")
                Set mListItem = FaListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case "JournalBook", "Annexure"
                If RstGroup.RecordCount = 0 Or (RstGroup.EOF = True Or RstGroup.BOF = True) Or TxtGrid(Index).TEXT = "" Then Exit Sub
                If TxtGrid(Index).TEXT <> RstGroup!Name Then
                    RstGroup.MoveFirst
                    RstGroup.FIND "Name =" & FaChk_Text(TxtGrid(Index).TEXT)
                End If
                DGGroup.left = TxtGrid(Index).left
                DGGroup.top = TxtGrid(Index).top + TxtGrid(Index).height
                DGGroup.Tag = Index
            Case "DailySumm", "AgingDr", "AgingCr", "Clg", "ClgNot"
                If RstAccount.RecordCount = 0 Or (RstAccount.EOF = True Or RstAccount.BOF = True) Or TxtGrid(Index).TEXT = "" Then Exit Sub
                If TxtGrid(Index).TEXT <> RstAccount!Name Then
                    RstAccount.MoveFirst
                    RstAccount.FIND "Name =" & FaChk_Text(TxtGrid(Index).TEXT)
                End If
                DGAccount.left = 0
                DGAccount.top = TxtGrid(Index).top + TxtGrid(Index).height
                DGAccount.Tag = Index
        End Select
    Case List2
        Select Case GRepFormName
            Case "Annexure", "RefReport", "JournalBook", "DetailedTrial"
                ListArray = Array("Yes", "No")
                Set mListItem = FaListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            Case "DayBook"
                If RstSite.RecordCount = 0 Or (RstSite.EOF = True Or RstSite.BOF = True) Or TxtGrid(Index).TEXT = "" Then Exit Sub
                If TxtGrid(Index).TEXT <> RstSite!Name Then
                    RstSite.MoveFirst
                    RstSite.FIND "Name =" & FaChk_Text(TxtGrid(Index).TEXT)
                End If
                DGAccount.left = 0
                DGAccount.top = TxtGrid(Index).top + TxtGrid(Index).height
                DGAccount.Tag = Index
            Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                If UCase(FGrid.TextMatrix(List1, 1)) = "GROUP(RANGE)" Then
                    If RstGroup.RecordCount = 0 Or (RstGroup.EOF = True Or RstGroup.BOF = True) Or TxtGrid(Index).TEXT = "" Then Exit Sub
                    If TxtGrid(Index).TEXT <> RstGroup!Name Then
                        RstGroup.MoveFirst
                        RstGroup.FIND "Name =" & FaChk_Text(TxtGrid(Index).TEXT)
                    End If
                    DGGroup.left = TxtGrid(Index).left
                    DGGroup.top = TxtGrid(Index).top + TxtGrid(Index).height
                    DGGroup.Tag = Index
                Else
                    If RstAccount.RecordCount = 0 Or (RstAccount.EOF = True Or RstAccount.BOF = True) Or TxtGrid(Index).TEXT = "" Then Exit Sub
                    If TxtGrid(Index).TEXT <> RstAccount!Name Then
                        RstAccount.MoveFirst
                        RstAccount.FIND "Name =" & FaChk_Text(TxtGrid(Index).TEXT)
                    End If
                    DGAccount.left = 0
                    DGAccount.top = TxtGrid(Index).top + TxtGrid(Index).height
                    DGAccount.Tag = Index
                End If
        End Select
    Case List3
        Select Case GRepFormName
            Case "CashBook", "BankBook", "Led", "LedInt", "AcCheckList", "LedDeb", "LedCred", "RozNamcha"
                If RstAccount.RecordCount = 0 Or (RstAccount.EOF = True Or RstAccount.BOF = True) Or TxtGrid(Index).TEXT = "" Then Exit Sub
                If TxtGrid(Index).TEXT <> RstAccount!Name Then
                    RstAccount.MoveFirst
                    RstAccount.FIND "Name =" & FaChk_Text(TxtGrid(Index).TEXT)
                End If
                DGAccount.left = 0
                DGAccount.top = TxtGrid(Index).top + TxtGrid(Index).height
                DGAccount.Tag = Index
            Case "JournalBook"
                If RstSite.RecordCount = 0 Or (RstSite.EOF = True Or RstSite.BOF = True) Or TxtGrid(Index).TEXT = "" Then Exit Sub
                If TxtGrid(Index).TEXT <> RstSite!Name Then
                    RstSite.MoveFirst
                    RstSite.FIND "Name =" & FaChk_Text(TxtGrid(Index).TEXT)
                End If
                DGAccount.left = 0
                DGAccount.top = TxtGrid(Index).top + TxtGrid(Index).height
                DGAccount.Tag = Index
        End Select
    Case Cat1
        Select Case GRepFormName
            Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                If RstAccount.RecordCount = 0 Or (RstAccount.EOF = True Or RstAccount.BOF = True) Or TxtGrid(Index).TEXT = "" Then Exit Sub
                If TxtGrid(Index).TEXT <> RstAccount!Name Then
                    RstAccount.MoveFirst
                    RstAccount.FIND "Name =" & FaChk_Text(TxtGrid(Index).TEXT)
                End If
                DGAccount.left = 0
                DGAccount.top = TxtGrid(Index).top + TxtGrid(Index).height
                DGAccount.Tag = Index
            Case "CashBook", "BankBook"
                ListArray = Array("Yes", "No")
                Set mListItem = FaListView_Items(ListView, TxtGrid, Index, ListArray, 2)
        End Select
End Select
Set Rst1 = Nothing
End Sub
Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Integer
If KeyCode = vbKeyEscape Then
    TxtGrid(0).TEXT = TxtGrid(0).Tag
    TxtGrid_KeyUp Index, KeyCode, Shift
    FGrid.SetFocus
    TxtGrid(0).Visible = False
    Grid_Hide
    Exit Sub
End If
Select Case FGrid.Row
    Case List1
        Select Case GRepFormName
            Case "DUELIST"
                ListView.ColumnHeaders(1).width = 1500
                ListView.ColumnHeaders(2).width = 0
                FaListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), 1750, 800
            Case "CashBook", "BankBook", "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList", "DayBook", "RozNamcha", "DetailedTrial"
                FaListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), TxtGrid(0).width
            Case "JournalBook", "Annexure"
                FaDGridTxtKeyDown DGGroup, TxtGrid, Index, RstGroup, KeyCode, True, 1
            Case "DailySumm", "AgingDr", "AgingCr", "Clg", "ClgNot"
                FaDGridTxtKeyDown DGAccount, TxtGrid, Index, RstAccount, KeyCode, True, 1
        End Select
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
    Case List2
        Select Case GRepFormName
            Case "Annexure", "RefReport", "JournalBook", "DetailedTrial"
                FaListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), TxtGrid(0).width
            Case "DayBook"
                If PubFaSiteType <> 0 Then
                    FaDGridTxtKeyDown DGSite, TxtGrid, Index, RstSite, KeyCode, True, 1
                End If
            Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                If UCase(FGrid.TextMatrix(List1, 1)) = "GROUP(RANGE)" Then
                    FaDGridTxtKeyDown DGGroup, TxtGrid, Index, RstGroup, KeyCode, True, 1
                Else
                    FaDGridTxtKeyDown DGAccount, TxtGrid, Index, RstAccount, KeyCode, True, 1
                End If
        End Select
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
    Case List3
        Select Case GRepFormName
            Case "CashBook", "BankBook", "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList", "RozNamcha"
                FaDGridTxtKeyDown DGAccount, TxtGrid, Index, RstAccount, KeyCode, True, 1
            Case "JournalBook"
                If PubFaSiteType <> 0 Then
                    FaDGridTxtKeyDown DGSite, TxtGrid, Index, RstSite, KeyCode, True, 1
                End If
        End Select
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
    Case Cat1
        Select Case GRepFormName
            Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                FaDGridTxtKeyDown DGAccount, TxtGrid, Index, RstAccount, KeyCode, True, 1
            Case "CashBook", "BankBook"
                FaListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), TxtGrid(0).width
        End Select
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
    Case Date1, Date2, Cat2, Cat3, Cat4, Cat5
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
End Select
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
FaCheckQuote KeyAscii
Select Case FGrid.Row
    Case List1
        Select Case GRepFormName
            Case "JournalBook", "Annexure"
                If DGGroup.Visible = True Then FaDGridTxtKeyPress TxtGrid, Index, RstGroup, KeyAscii, "Name"
            Case "DailySumm", "AgingDr", "AgingCr", "Clg", "ClgNot"
                If DGAccount.Visible = True Then FaDGridTxtKeyPress TxtGrid, Index, RstAccount, KeyAscii, "Name"
        End Select
    Case List2
        Select Case GRepFormName
            Case "DueListPeriod"
                FaNumPress TxtGrid(Index), KeyAscii, 3, 0
            Case "CashBook", "BankBook", "RozNamcha"
                FaNumPress TxtGrid(Index), KeyAscii, 4, 0
            Case "DayBook"
                If PubFaSiteType <> 0 Then
                    If DGSite.Visible = True Then FaDGridTxtKeyPress TxtGrid, Index, RstSite, KeyAscii, "Name"
                End If
            Case "Led", "LedInt"
                If UCase(FGrid.TextMatrix(List1, 1)) = "GROUP(RANGE)" Then
                    If DGGroup.Visible = True Then FaDGridTxtKeyPress TxtGrid, Index, RstGroup, KeyAscii, "Name"
                Else
                    If DGAccount.Visible = True Then FaDGridTxtKeyPress TxtGrid, Index, RstAccount, KeyAscii, "Name"
                End If
        End Select
    Case List3
        Select Case GRepFormName
            Case "CashBook", "BankBook", "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList", "RozNamcha"
                If DGAccount.Visible = True Then FaDGridTxtKeyPress TxtGrid, Index, RstAccount, KeyAscii, "Name"
            Case "JournalBook"
                If PubFaSiteType <> 0 Then
                    If DGSite.Visible = True Then FaDGridTxtKeyPress TxtGrid, Index, RstSite, KeyAscii, "Name"
                End If
        End Select
    Case Cat1
        Select Case GRepFormName
            Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                If DGAccount.Visible = True Then FaDGridTxtKeyPress TxtGrid, Index, RstAccount, KeyAscii, "Name"
        End Select
End Select
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case FGrid.Row
    Case List1
        Select Case GRepFormName
            Case "JournalBook", "Annexure"
                If KeyCode <> 13 And DGGroup.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: FaDGridTxtKeyPress TxtGrid, Index, RstGroup, KeyCode, "Name", True
            Case "DailySumm", "AgingDr", "AgingCr", "Clg", "ClgNot"
                If KeyCode <> 13 And DGAccount.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: FaDGridTxtKeyPress TxtGrid, Index, RstAccount, KeyCode, "Name", True
            Case Else
                If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
                FaListView_KeyUp ListView, TxtGrid, 0, KeyCode, mListItem
            End Select
    Case List2
        Select Case GRepFormName
            Case "Annexure", "RefReport", "JournalBook"
                If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
                FaListView_KeyUp ListView, TxtGrid, 0, KeyCode, mListItem
            Case "DayBook"
                If PubFaSiteType <> 0 Then
                    If KeyCode <> 13 And DGSite.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: FaDGridTxtKeyPress TxtGrid, Index, RstSite, KeyCode, "Name", True
                End If
            Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                If UCase(FGrid.TextMatrix(List1, 1)) = "GROUP(RANGE)" Then
                    If KeyCode <> 13 And DGAccount.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: FaDGridTxtKeyPress TxtGrid, Index, RstGroup, KeyCode, "Name", True
                Else
                    If KeyCode <> 13 And DGAccount.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: FaDGridTxtKeyPress TxtGrid, Index, RstAccount, KeyCode, "Name", True
                End If
        End Select
    Case List3
        Select Case GRepFormName
            Case "CashBook", "BankBook", "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList", "RozNamcha"
                If KeyCode <> 13 And DGAccount.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: FaDGridTxtKeyPress TxtGrid, Index, RstAccount, KeyCode, "Name", True
            Case "JournalBook"
                If PubFaSiteType <> 0 Then
                    If KeyCode <> 13 And DGSite.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: FaDGridTxtKeyPress TxtGrid, Index, RstSite, KeyCode, "Name", True
                End If
            Case Else
                If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
                FaListView_KeyUp ListView, TxtGrid, 0, KeyCode, mListItem
        End Select
    Case Cat1
        Select Case GRepFormName
            Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                If KeyCode <> 13 And DGAccount.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: FaDGridTxtKeyPress TxtGrid, Index, RstAccount, KeyCode, "Name", True
            Case Else
                If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
                FaListView_KeyUp ListView, TxtGrid, 0, KeyCode, mListItem
        End Select
End Select
End Sub
Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
    Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:  MsgBox err.Description, vbInformation
End Sub
Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim mSiteHlpSubgroup As String, mSiteHlpAndSubgroup As String
Dim mSiteHlpViewSubgroup As String, mSiteHlpAndViewSubgroup As String
mSiteHlpSubgroup = ""
mSiteHlpAndSubgroup = ""
mSiteHlpViewSubgroup = ""
mSiteHlpAndViewSubgroup = ""
If PubSiteCodeWiseHelp = True Then
    mSiteHlpSubgroup = " Where Subgroup.Site_Code='" & PubSiteCode & "'"
    mSiteHlpAndSubgroup = " And Subgroup.Site_Code='" & PubSiteCode & "'"
    mSiteHlpViewSubgroup = " Where ViewSubgroup.Site_Code='" & PubSiteCode & "'"
    mSiteHlpAndViewSubgroup = " And ViewSubgroup.Site_Code='" & PubSiteCode & "'"
End If
Select Case FGrid.Row
    Case Cat1, Cat2, Cat3, Cat4, Cat5
        Select Case GRepFormName
            Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                If FGrid.Row = Cat1 Then
                    If RstAccount.RecordCount = 0 Or (RstAccount.EOF = True Or RstAccount.BOF = True) Or TxtGrid(Index).TEXT = "" Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                        FGrid.TextMatrix(FGrid.Row, 2) = ""
                    Else
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RstAccount!Name
                        FGrid.TextMatrix(FGrid.Row, 2) = RstAccount!Code
                    End If
                End If
            Case "CashBook", "BankBook"
                If TxtGrid(0).TEXT <> "" Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
                    FGrid.TextMatrix(FGrid.Row, 2) = ListView.SelectedItem.SubItems(1)
                End If
            Case Else
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
        End Select
    Case List1
        Select Case GRepFormName
            Case "DUELIST"
                If TxtGrid(0).TEXT <> "" Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
                    FGrid.TextMatrix(FGrid.Row, 2) = ListView.SelectedItem.SubItems(1)
                End If
            Case "CashBook", "BankBook", "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList", "DayBook", "RozNamcha", "DetailedTrial"
                If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
                Select Case GRepFormName
                    Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                        If UCase(FGrid.TextMatrix(FGrid.Row, FGrid.Col)) = "SELECTED" Then
                            TxtDetails.Visible = True
                            GridSel(1).Visible = True
                            Check1(1).Visible = True
                            GridSel(2).Visible = False
                            Check1(2).Visible = False
                            FGrid.RowHeight(List2) = 0
                            FGrid.RowHeight(List3) = 0
                            FGrid.RowHeight(Cat1) = 0
                            mLastRow = List1
                            FGrid.height = 310 * 3
                        End If
                        If UCase(FGrid.TextMatrix(FGrid.Row, FGrid.Col)) = "MERGE" Then
                            TxtDetails.Visible = False
                            FGrid.TextMatrix(List2, 0) = "From Account        ": FGrid.RowHeight(List2) = GridRowHeight
                            FGrid.TextMatrix(List3, 0) = "To Account          ": FGrid.RowHeight(List3) = GridRowHeight
                            FGrid.RowHeight(List2) = GridRowHeight
                            FGrid.RowHeight(List3) = GridRowHeight
                            FGrid.RowHeight(Cat1) = 0
                            mLastRow = List3
                            FGrid.height = 310 * 5
                            GridSel(1).Visible = False
                            Check1(1).Visible = False
                            GridSel(2).Visible = False
                            Check1(2).Visible = False
                            FGrid.TextMatrix(List2, 1) = RstAccount!Name
                            FGrid.TextMatrix(List2, 2) = RstAccount!Code
                            FGrid.TextMatrix(List3, 1) = RstAccount!Name
                            RstAccount.Filter = adFilterNone
                            FGrid.TextMatrix(List3, 2) = RstAccount!Code
                        End If
                        If UCase(FGrid.TextMatrix(FGrid.Row, FGrid.Col)) = "GROUP(SELECTED)" Then
                            TxtDetails.Visible = False
                            GridInitialise 2, "SELECT '' as O,GroupName,GroupCode as Code from AcGroup Order by GroupName"
                            GridSel(1).Visible = False
                            Check1(1).Visible = False
                            GridSel(2).Visible = True
                            Check1(2).Visible = True
                            GridSel(2).height = GridSel(1).height
                            GridSel(2).left = GridSel(1).left
                            GridSel(2).width = GridSel(1).width
                            GridSel(2).top = GridSel(1).top
                            Check1(2).height = Check1(1).height
                            Check1(2).left = Check1(1).left
                            Check1(2).width = Check1(1).width
                            Check1(2).top = Check1(1).top
                            FGrid.RowHeight(List2) = 0
                            FGrid.RowHeight(List3) = 0
                            FGrid.RowHeight(Cat1) = 0
                            mLastRow = List1
                            FGrid.height = 310 * 3
                        End If
                        If UCase(FGrid.TextMatrix(List1, 1)) = "GROUP(RANGE)" Then
                            TxtDetails.Visible = False
                            FGrid.TextMatrix(List2, 0) = "For Group           ": FGrid.RowHeight(List2) = GridRowHeight
                            FGrid.TextMatrix(List3, 0) = "From Account        ": FGrid.RowHeight(List3) = GridRowHeight
                            FGrid.TextMatrix(Cat1, 0) = "To Account          ": FGrid.RowHeight(Cat1) = GridRowHeight
                            FGrid.TextMatrix(List2, 1) = RstGroup!Name
                            FGrid.TextMatrix(List2, 2) = RstGroup!Code
                            RstAccount.Filter = adFilterNone
                            RstAccount.Filter = " GroupCode='" & RstGroup!Code & "'"
                            FGrid.TextMatrix(List3, 1) = ""
                            FGrid.TextMatrix(List3, 2) = ""
                            FGrid.TextMatrix(Cat1, 1) = ""
                            FGrid.TextMatrix(Cat1, 2) = ""
                            mLastRow = Cat1
                            FGrid.height = 310 * 6
                            GridSel(1).Visible = False
                            Check1(1).Visible = False
                            GridSel(2).Visible = False
                            Check1(2).Visible = False
                        End If
                        If UCase(FGrid.TextMatrix(FGrid.Row, FGrid.Col)) = "CITY WISE" Then
                            TxtDetails.Visible = False
                            GridInitialise 2, "SELECT '' as O,CityName,CityCode as Code from City Order by CityName"
                            GridSel(2).Visible = True
                            Check1(2).Visible = True
                            GridSel(1).Visible = False
                            Check1(1).Visible = False
                            GridSel(2).height = GridSel(1).height
                            GridSel(2).left = GridSel(1).left
                            GridSel(2).width = GridSel(1).width
                            GridSel(2).top = GridSel(1).top
                            Check1(2).height = Check1(1).height
                            Check1(2).left = Check1(1).left
                            Check1(2).width = Check1(1).width
                            Check1(2).top = Check1(1).top
                            FGrid.RowHeight(List2) = 0
                            FGrid.RowHeight(List3) = 0
                            FGrid.RowHeight(Cat1) = 0
                            mLastRow = List1
                            FGrid.height = 310 * 3
                        End If
                End Select
            Case "JournalBook", "Annexure"
                If RstGroup.RecordCount = 0 Or (RstGroup.EOF = True Or RstGroup.BOF = True) Or TxtGrid(Index).TEXT = "" Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    FGrid.TextMatrix(FGrid.Row, 2) = ""
                Else
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RstGroup!Name
                    FGrid.TextMatrix(FGrid.Row, 2) = RstGroup!Code
                End If
                Select Case GRepFormName
                    Case "Annexure"
                        GridInitialise 1, "SELECT '' as O,NAME as Account,SUBCODE as AccId from SUBGROUP where GROUPCODE='" & FGrid.TextMatrix(List1, 2) & "' " & mSiteHlpAndSubgroup & " Order by Name"
                        GridSel(1).height = Me.height - FGrid.height - Pic.height - 1200
                End Select
            Case "DailySumm", "AgingDr", "AgingCr", "Clg", "ClgNot"
                If RstAccount.RecordCount = 0 Or (RstAccount.EOF = True Or RstAccount.BOF = True) Or TxtGrid(Index).TEXT = "" Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    FGrid.TextMatrix(FGrid.Row, 2) = ""
                Else
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RstAccount!Name
                    FGrid.TextMatrix(FGrid.Row, 2) = RstAccount!Code
                End If
        End Select
    Case List2
        Select Case GRepFormName
            Case "DUELIST"
                If TxtGrid(0).TEXT <> "" Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
                End If
            Case "Annexure", "RefReport", "JournalBook", "DetailedTrial"
                If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
        End Select
        Select Case GRepFormName
            Case "DayBook"
                If PubFaSiteType <> 0 Then
                    If RstSite.RecordCount = 0 Or (RstSite.EOF = True Or RstSite.BOF = True) Or TxtGrid(Index).TEXT = "" Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                        FGrid.TextMatrix(FGrid.Row, 2) = ""
                    Else
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RstSite!Name
                        FGrid.TextMatrix(FGrid.Row, 2) = RstSite!Code
                    End If
                End If
            Case "CashBook", "BankBook", "RozNamcha"
                If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = FaValidate_Numeric(TxtGrid(0).TEXT)
            Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                If UCase(FGrid.TextMatrix(List1, 1)) = "GROUP(RANGE)" Then
                    If RstGroup.RecordCount = 0 Or (RstGroup.EOF = True Or RstGroup.BOF = True) Or TxtGrid(Index).TEXT = "" Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                        FGrid.TextMatrix(FGrid.Row, 2) = ""
                        FGrid.TextMatrix(List3, 1) = ""
                        FGrid.TextMatrix(List3, 2) = ""
                        FGrid.TextMatrix(Cat1, 1) = ""
                        FGrid.TextMatrix(Cat1, 2) = ""
                    Else
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RstGroup!Name
                        FGrid.TextMatrix(FGrid.Row, 2) = RstGroup!Code
                        FGrid.TextMatrix(List3, 1) = ""
                        FGrid.TextMatrix(List3, 2) = ""
                        FGrid.TextMatrix(Cat1, 1) = ""
                        FGrid.TextMatrix(Cat1, 2) = ""
                        RstAccount.Filter = adFilterNone
                        RstAccount.Filter = " GroupCode='" & RstGroup!Code & "'"
                    End If
                Else
                    If RstAccount.RecordCount = 0 Or (RstAccount.EOF = True Or RstAccount.BOF = True) Or TxtGrid(Index).TEXT = "" Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                        FGrid.TextMatrix(FGrid.Row, 2) = ""
                    Else
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RstAccount!Name
                        FGrid.TextMatrix(FGrid.Row, 2) = RstAccount!Code
                    End If
                End If
        End Select
    Case List3
        Select Case GRepFormName
            Case "CashBook", "BankBook", "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList", "RozNamcha"
                If RstAccount.RecordCount = 0 Or (RstAccount.EOF = True Or RstAccount.BOF = True) Or TxtGrid(Index).TEXT = "" Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    FGrid.TextMatrix(FGrid.Row, 2) = ""
                Else
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RstAccount!Name
                    FGrid.TextMatrix(FGrid.Row, 2) = RstAccount!Code
                End If
            Case "JournalBook"
                If PubFaSiteType <> 0 Then
                    If RstSite.RecordCount = 0 Or (RstSite.EOF = True Or RstSite.BOF = True) Or TxtGrid(Index).TEXT = "" Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                        FGrid.TextMatrix(FGrid.Row, 2) = ""
                    Else
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RstSite!Name
                        FGrid.TextMatrix(FGrid.Row, 2) = RstSite!Code
                    End If
                End If
        End Select
    Case Date1, Date2
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = PubDatamanFa.FaRetDateFunc(TxtGrid(0))
End Select
TxtGridLeave = True
If ValidateCall = False Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
End Function
Private Sub BTNCLOSE_Click()
    FrDetail.Visible = False
    If GRepFormName = "LedDeb" Then
        G_FaCn.Execute "UPDATE FAENVIRO SET " & "TagadaHeader1=" & FaChk_Text(TxtHeader(0)) & ",TagadaHeader2= " & FaChk_Text(TxtHeader(1)) & ",TagadaHeader3= " & FaChk_Text(TxtHeader(2)) & ",TagadaHeader4= " & FaChk_Text(TxtHeader(3)) & ",TagadaHeader5= " & FaChk_Text(TxtHeader(4)) & ",TagadaFooter1= " & FaChk_Text(TxtFooter(0)) & ",TagadaFooter2= " & FaChk_Text(TxtFooter(1)) & ",TagadaFooter3= " & FaChk_Text(TxtFooter(2)) & ",TagadaFooter4= " & FaChk_Text(TxtFooter(3)) & ",TagadaFooter5= " & FaChk_Text(TxtFooter(4))
    End If
End Sub
Private Sub BtnParam_Click()
    FrDetail.Visible = True
    FrDetail.left = 1395
    FrDetail.top = 1900
    Select Case GRepFormName
        Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
            Label1(7).Visible = True
            TEXT(PrintVrNo).Visible = True
    End Select
    If PubFaSiteType <> 0 Then
        Label1(6).Visible = True
        TEXT(SiteCode).Visible = True
        DGSite.left = FrDetail.left + TEXT(SiteCode).left
        DGSite.top = FrDetail.top + TEXT(SiteCode).top + TEXT(SiteCode).height
        DGSite.Tag = SiteCode
    Else
        Label1(6).Visible = False
        TEXT(SiteCode).Visible = False
    End If
End Sub
Private Sub Global_Grid()
Dim I As Integer, cnt As Integer
Pic.top = Me.top - Pic.width - 10
BtnPrint(1).left = (Pic.width - (BtnPrint(1).width + BtnPrint(0).width + BtnPrint(0).width)) / 2: BtnPrint(1).top = Pic.top + 10
BtnPrint(0).left = BtnPrint(1).left + BtnPrint(1).width: BtnPrint(0).top = Pic.top + 10
BTNEXIT.left = BtnPrint(0).left + BtnPrint(0).width: BTNEXIT.top = Pic.top + 10
TxtGrid(0) = ""
FGrid.left = (Me.width - FGrid.width) / 2
FGrid.top = 75
FGrid.Rows = 10  '5
FGrid.Cols = 3
FGrid.FixedCols = 1
FGrid.ColWidth(0) = 2200
FGrid.ColWidth(1) = 2000
FGrid.ColWidth(2) = 0
FGrid.ColAlignment(1) = flexAlignLeftCenter
For I = 0 To FGrid.Rows - 1
    FGrid.RowHeight(I) = 0
Next
Ini_Grid
For I = 1 To 2
    If GridSel(I).Visible = True Then cnt = cnt + 1
Next
FGrid.height = ((mLastRow - mFirstRow + 1) * PubGridRowHeight) + 500
Select Case mHelpGridNo
    Case 0
        FGrid.top = 1000
    Case 1
        Select Case GRepFormName
            Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
                Label1(0).Visible = True
                Label1(1).Visible = True
                Label1(3).Visible = True
                Label1(4).Visible = True
                Label1(5).Visible = True
                TEXT(DrBalance).Visible = True
                TEXT(CrBalance).Visible = True
                TEXT(SignDr).Visible = True
                TEXT(SignCr).Visible = True
                TEXT(AsPerDetail).Visible = True
                TEXT(AsPerDetailNarration).Visible = True
                TEXT(PrintNarration).Visible = True
                Label1(0).left = 120: Label1(0).top = 150
                Label1(1).left = 120: Label1(1).top = 450
                TEXT(DrBalance).left = 1725: TEXT(DrBalance).top = 105
                TEXT(SignDr).left = 1725 - TEXT(SignDr).width - 50: TEXT(SignDr).top = 105
                TEXT(CrBalance).left = 1725: TEXT(CrBalance).top = 405
                TEXT(SignCr).left = 1725 - TEXT(SignCr).width - 50: TEXT(SignCr).top = 405
                If GRepFormName = "LedDeb" Then SSTab1.Visible = True
                Select Case GRepFormName
                    Case "Led", "LedDeb", "LedCred"
                        Label1(3).left = 120: Label1(3).top = 750
                        Label1(4).left = 120: Label1(4).top = 1050
                        TEXT(AsPerDetail).left = 1725: TEXT(AsPerDetail).top = 705
                        TEXT(AsPerDetailNarration).left = 1725: TEXT(AsPerDetailNarration).top = 1005
                    Case "LedInt"
                        Label4.Visible = True: TEXT(Interest).Visible = True
                        Label4.left = 120: Label4.top = 750
                        Label1(3).left = 120: Label1(3).top = 1050
                        Label1(4).left = 120: Label1(4).top = 1350
                        TEXT(Interest).left = 1725: TEXT(Interest).top = 705
                        TEXT(AsPerDetail).left = 1725: TEXT(AsPerDetail).top = 1005
                        TEXT(AsPerDetailNarration).left = 1725: TEXT(AsPerDetailNarration).top = 1305
                    Case "AcCheckList"
                        Label1(2).Visible = True: TEXT(TxNAmount).Visible = True: TEXT(SignAmt).Visible = True
                        Label5.Visible = True: TEXT(NarrationHaving).Visible = True
                        Label6.Visible = True: TEXT(NarrationNotHaving).Visible = True
                        Label1(2).left = 120: Label1(2).top = 750
                        TEXT(TxNAmount).left = 1725: TEXT(TxNAmount).top = 705
                        TEXT(SignAmt).left = 1725 - TEXT(SignAmt).width - 50: TEXT(SignAmt).top = 705
                        Label1(3).left = 120: Label1(3).top = 1050
                        Label1(4).left = 120: Label1(4).top = 1350
                        TEXT(AsPerDetail).left = 1725: TEXT(AsPerDetail).top = 1005
                        TEXT(AsPerDetailNarration).left = 1725: TEXT(AsPerDetailNarration).top = 1305
                        Label5.left = 120: Label5.top = 1650
                        TEXT(NarrationHaving).left = 1725: TEXT(NarrationHaving).top = 1605
                        Label6.left = 120: Label6.top = 1950
                        TEXT(NarrationNotHaving).left = 1725: TEXT(NarrationNotHaving).top = 1905
                End Select
        End Select
        If GRepFormName = "AcCheckList" Then
            GridSel(1).left = (Me.width / 2 - GridSel(1).width) / 2
            GridSel(1).top = FGrid.top + FGrid.height + 700
            GridSel(1).height = Me.height - FGrid.height - Pic.height - 2200
            Check1(1).top = GridSel(1).top + 20: Check1(1).left = GridSel(1).left + 40
            
            GridSel(3).left = Me.width / 2 + (Me.width / 2 - GridSel(1).width) / 2
            GridSel(3).top = FGrid.top + FGrid.height + 700
            GridSel(3).height = Me.height - FGrid.height - Pic.height - 2200
            Check1(3).top = GridSel(3).top + 20: Check1(3).left = GridSel(3).left + 40
            If GRepFormName = "Led" Or GRepFormName = "LedInt" Or GRepFormName = "LedDeb" Or GRepFormName = "LedCred" Or GRepFormName = "AcCheckList" Then
                TxtDetails.Visible = True
                TxtDetails.left = GridSel(1).left
                TxtDetails.width = GridSel(1).width
                TxtDetails.top = GridSel(1).top + GridSel(1).height
            End If
        Else
            GridSel(1).left = (Me.width - GridSel(1).width) / 2
            GridSel(1).top = FGrid.top + FGrid.height + 700
            GridSel(1).height = Me.height - FGrid.height - Pic.height - 2200
            Check1(1).top = GridSel(1).top + 20: Check1(1).left = GridSel(1).left + 40
            If GRepFormName = "Led" Or GRepFormName = "LedInt" Or GRepFormName = "LedDeb" Or GRepFormName = "LedCred" Or GRepFormName = "AcCheckList" Then
                TxtDetails.Visible = True
                TxtDetails.left = GridSel(1).left
                TxtDetails.width = GridSel(1).width
                TxtDetails.top = GridSel(1).top + GridSel(1).height
            End If
        End If
    Case 2
        GridSel(1).left = (Me.width / 2 - GridSel(1).width) / 2
        GridSel(1).top = FGrid.top + FGrid.height + 700
        GridSel(1).height = Me.height - FGrid.height - Pic.height - 2200
        Check1(1).top = GridSel(1).top + 20
        Check1(1).left = GridSel(1).left + 40
        GridSel(2).left = Me.width / 2 + (Me.width / 2 - GridSel(1).width) / 2
        GridSel(2).top = FGrid.top + FGrid.height + 700
        GridSel(2).height = Me.height - FGrid.height - Pic.height - 2200
        Check1(2).top = GridSel(2).top + 20
        Check1(2).left = GridSel(2).left + 40
        If GRepFormName = "Led" Or GRepFormName = "LedInt" Or GRepFormName = "LedDeb" Or GRepFormName = "LedCred" Or GRepFormName = "AcCheckList" Then
            TxtDetails.Visible = True
            TxtDetails.left = GridSel(1).left
            TxtDetails.width = GridSel(1).width
            TxtDetails.top = GridSel(1).top + GridSel(1).height
        End If
End Select
End Sub
Private Sub Grid_Hide()
If FrmList.Visible = True Then FrmList.Visible = False
If DGAccount.Visible = True Then DGAccount.Visible = False
If DGGroup.Visible = True Then DGGroup.Visible = False
End Sub
Private Sub FGrid_DblClick()
    Select Case FGrid.Row
        Case Date1, Date2, List1, List2, List3, Cat1, Cat2, Cat3, Cat4, Cat5
            Call FaGridDblClick(Me, FGrid, TxtGrid, 0)
    End Select
TAddMode = False
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp And Val(FGrid.Tag) = mFirstRow Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = mLastRow Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys vbTab
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    FGrid.TextMatrix(FGrid.Row, 2) = ""
End If
Select Case FGrid.Col
    Case List3
        Select Case GRepFormName
            Case "CashBook", "BankBook"
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        End Select
End Select
If KeyCode = vbKeyReturn Then
    Select Case FGrid.Row
        Case Date1, Date2, List1, List2, List3, Cat1, Cat2, Cat3, Cat4, Cat5
            Call FaGridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
Dim I As Integer
    Select Case FGrid.Row
        Case Cat2, Cat3, Cat4, Cat5
            FaGet_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
        Case Date1, Date2
            FaGet_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        Case List1
            FaGet_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        Case List2
            FaGet_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        Case List3, Cat1
            FaGet_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            Select Case GRepFormName
                Case "CashBook", "BankBook"
                    'TxtGrid_KeyPress 0, KeyAscii
                    FaGet_Text Me, FGrid, TxtGrid, 0, False, Asc(UCase(Chr(KeyAscii)))
                    If Len(TxtGrid(0).TEXT) <= 1 Then TxtGrid(0).TEXT = TxtGrid(0).TEXT
                    TxtGrid(0).SelStart = Len(TxtGrid(0).TEXT)
            End Select
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub
Private Sub FGrid_EnterCell()
    FGrid.CellBackColor = CellBackColEnter
End Sub
Private Sub FGrid_GotFocus()
   FGrid.CellBackColor = CellBackColEnter
   TxtGrid(0).Visible = False
   Grid_Hide
End Sub
Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub
Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub
Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
End Sub
Private Function FillString(GridArray As Variant, Gridindex As Integer, DataType As Byte) As String
On Error GoTo ELoop
Dim ac_str As String, I As Integer, GridRow As Integer, FormulaString As String
    ac_str = ""
    For I = 0 To UBound(GridArray)
        If GridArray(I) = 0 Then GoTo NXT:
        GridRow = GridArray(I)
        If GridSel(Gridindex).TextMatrix(GridRow, 0) = "ü" Then
                If DataType = 0 Then
                   ac_str = ac_str + IIf(ac_str = "", GridSel(Gridindex).TextMatrix(GridRow, 2), "," + GridSel(Gridindex).TextMatrix(GridRow, 2))
                ElseIf DataType = 1 Then
                   ac_str = ac_str + IIf(ac_str = "", "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'", "," + "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'")
                End If
                If Len(FormulaString + GridSel(Gridindex).TextMatrix(GridRow, 1)) < 255 Then
                    FormulaString = FormulaString + IIf(FormulaString = "", GridSel(Gridindex).TextMatrix(GridRow, 1), "," + GridSel(Gridindex).TextMatrix(GridRow, 1))
                End If
            GridSel(Gridindex).TextMatrix(GridRow, 0) = ""
        Else
            GridArray(I) = 0
        End If
NXT:
    Next
    For I = 0 To UBound(GridArray)
        GridRow = GridArray(I)
        If GridArray(I) <> 0 Then
            GridSel(Gridindex).TextMatrix(GridRow, 0) = "ü"
        End If
    Next
    If ac_str = "" Then
        MsgBox "Select " & GridSel(Gridindex).TextMatrix(0, 1), vbInformation
        GridSel(Gridindex).SetFocus
        RepPrint = False
        Exit Function
    End If
    FillString = ac_str
    Select Case Gridindex
        Case 1
            FormulaString1 = FormulaString
        Case 2
            FormulaString2 = FormulaString
    End Select
    Exit Function
ELoop:      RepPrint = False
            MsgBox err.Description
End Function
Private Sub TxtKeyDown()
Dim I As Integer
    If FGrid.Row = mLastRow Then SendKeys vbTab: Exit Sub
    For I = FGrid.Row To FGrid.Rows - 1
         If FGrid.RowHeight(I + 1) <> 0 Then FGrid.Row = I + 1: Exit For
    Next
End Sub
Private Sub GridInitialise(Gridindex As Integer, GridSql As String)
Dim Index As Integer
Index = Gridindex
Select Case Index
    Case 1
        Set RsGrid1 = New ADODB.Recordset: RsGrid1.CursorLocation = adUseClient
        RsGrid1.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid1
        ReDim Preserve GridRow1(0)
        GridRow1(0) = 0
    Case 2
        Set RsGrid2 = New ADODB.Recordset: RsGrid2.CursorLocation = adUseClient
        RsGrid2.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid2
        ReDim Preserve GridRow2(0)
        GridRow2(0) = 0
    Case 3
        Set RsGrid3 = New ADODB.Recordset: RsGrid3.CursorLocation = adUseClient
        RsGrid3.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid3
        ReDim Preserve GridRow3(0)
        GridRow3(0) = 0
End Select
GridSel(Index).height = 1700
GridSel(Index).Visible = True
GridSel(Index).Enabled = True
Check1(Index).Visible = True
GridSel(Index).width = 5200
GridSel(Index).ColWidth(0) = 600
GridSel(Index).ColWidth(2) = 0
GridSel(Index).ColWidth(1) = 4000
Check1(Index).width = 580
Check1(Index).height = GridSel(Index).RowHeight(0) + 20
Check1(Index).Value = Unchecked
End Sub
Private Function IsNotBlank(FieldRow As Integer, FieldCaption As String) As Boolean
If FGrid.TextMatrix(FieldRow, 1) = "" Then
    MsgBox Trim(FieldCaption) & " Should not be Blank.", vbInformation, "Validation Check"
    FGrid.SetFocus
    FGrid.Row = FieldRow
    FGrid.Col = 1
    IsNotBlank = False
Else
    IsNotBlank = True
End If
End Function
Private Sub Formulas()
On Error GoTo ELoop
Dim I As Integer
For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("Title")
            Select Case GRepFormName
                Case GRepFormName = "DUELIST"
                    rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & " More Than " & FGrid.TextMatrix(List2, 1) & " Days""'"
                Case "DayBook"    '" And PubFaSiteType <> 0
                    rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION + " (" + FGrid.TextMatrix(List2, 1) + ")" + "'"
                Case "JournalBook"
                    If Trim(FGrid.TextMatrix(List3, 2)) <> "" Then
                        rpt.FormulaFields(I).TEXT = "'" & Trim(FGrid.TextMatrix(List1, 1)) + " (" + Trim(FGrid.TextMatrix(List3, 1)) + ")" & "'"
                    Else
                        rpt.FormulaFields(I).TEXT = "'" & Trim(FGrid.TextMatrix(List1, 1)) & "'"
                    End If
                Case "Annexure", "BankReg", "CashBook", "BankBook", "Clg", "ClgNot", "DailySumm", "RozNamcha", "CONTROLLED", "DUELIST", "RefReport", "NonTrans"
                    If Trim(TEXT(SiteCode)) <> "" Then
                        rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION + " (" + Trim(TEXT(SiteCode)) + ")" + "'"
                    Else
                        rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
                    End If
                Case Else
                    rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
            End Select
        Case UCase("FORDATE"), UCase("DATE"), UCase("FROMDATE")
            Select Case GRepFormName
                Case "Annexure"
                    rpt.FormulaFields(I).TEXT = "'Upto Date : " & TXTE_DATE & "'"
                Case Else
                    rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            End Select
        Case UCase("PageNo")
            rpt.FormulaFields(I).TEXT = "'" & RstEnviro!pagenofill & "'"
        Case UCase("DT")
            rpt.FormulaFields(I).TEXT = "'" & RstEnviro!daterfill & "'"
        Case UCase("SepPage")
            rpt.FormulaFields(I).TEXT = "'" & UCase(RstEnviro!CashBookPage) & "'"
        Case UCase("PageNoIni")
            rpt.FormulaFields(I).TEXT = Val(FGrid.TextMatrix(List2, 1))
        Case UCase("ACNAME")
            Select Case GRepFormName
                Case "CashBook", "BankBook"
                    rpt.FormulaFields(I).TEXT = "'" & FGrid.TextMatrix(List3, 1) & "'"
                Case Else
                    rpt.FormulaFields(I).TEXT = "'From A/C : " & TXTACC_CODE & "'"
            End Select
        Case UCase("PARTYNAME"), UCase("FORPARTY")
            rpt.FormulaFields(I).TEXT = "'For A/c : " & TXTACC_CODE & "'"
        Case UCase("MyOpBal")
            rpt.FormulaFields(I).TEXT = TOT_AMTDR
        Case UCase("VRTOT")
            rpt.FormulaFields(I).TEXT = "'" & FGrid.TextMatrix(List2, 1) & "'"
        Case UCase("VRWISETOT")
            rpt.FormulaFields(I).TEXT = "'" & FGrid.TextMatrix(List1, 1) & "'"
        Case UCase("BetweenDate")
            rpt.FormulaFields(I).TEXT = "'Upto Date :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "'"
        Case UCase("ForParty")
            If Check1(1).Value = Unchecked Then
                rpt.FormulaFields(I).TEXT = Trim(left("'For Party :" & FormulaString1, 255)) & "'"
            Else
                rpt.FormulaFields(I).TEXT = "'For Party : All'"
            End If
    End Select
Next
Exit Sub
ELoop:      MsgBox err.Description
End Sub
Private Sub Ini_Grid()
Dim mQRYx As String, mQRYy As String, mSiteHlpSubgroup As String, mSiteHlpAndSubgroup As String
Dim mSiteHlpViewSubgroup As String, mSiteHlpAndViewSubgroup As String
mSiteHlpSubgroup = ""
mSiteHlpAndSubgroup = ""
mSiteHlpViewSubgroup = ""
mSiteHlpAndViewSubgroup = ""
If PubSiteCodeWiseHelp = True Then
    mSiteHlpSubgroup = " Where Subgroup.Site_Code='" & PubSiteCode & "'"
    mSiteHlpAndSubgroup = " And Subgroup.Site_Code='" & PubSiteCode & "'"
    mSiteHlpViewSubgroup = " Where ViewSubgroup.Site_Code='" & PubSiteCode & "'"
    mSiteHlpAndViewSubgroup = " And ViewSubgroup.Site_Code='" & PubSiteCode & "'"
End If
TxtDetails.Visible = False
Select Case GRepFormName
    Case "Led", "LedInt", "LedDeb", "LedCred", "AcCheckList"
        TxtDetails.Visible = True
        mQRYx = ""
        Select Case GRepFormName
            Case "LedDeb"
                mQRYx = " WHERE SubGroup.Nature='Customer' AND A.Nature='Customer'" + mSiteHlpAndSubgroup
                mQRYy = " WHERE ViewSubgroup.Nature='Customer' " + mSiteHlpAndViewSubgroup
            Case "LedCred"
                mQRYx = " WHERE SubGroup.Nature='Supplier' AND A.Nature='Supplier'" + mSiteHlpAndSubgroup
                mQRYy = " WHERE ViewSubgroup.Nature='Supplier' " + mSiteHlpAndViewSubgroup
            Case Else
                mQRYx = mSiteHlpSubgroup
                mQRYy = mSiteHlpViewSubgroup
        End Select
        With FGrid
            .TextMatrix(Date1, 0) = "From Date           ": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "To Date             ": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Account Selection   ": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Form Account        ": .RowHeight(List2) = 0
            .TextMatrix(List3, 0) = "To Account          ": .RowHeight(List3) = 0
            .TextMatrix(Cat1, 0) = "To Account          ": .RowHeight(Cat1) = 0
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Selected"
        End With
        ListView.Font.Size = 10
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List1: mHelpGridNo = 1
        FGrid.ColWidth(0) = 2000: FGrid.ColWidth(1) = 5000
        FGrid.width = FGrid.ColWidth(0) + FGrid.ColWidth(1) + 100
        FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75
        Set RstAccount = New ADODB.Recordset
        RstAccount.CursorLocation = adUseClient
        If PubBackEnd = "A" Then
            RstAccount.Open "Select SubCode AS Code,Name,RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3)) AS Address,CityName,SubGroup.GroupCode,A.GROUPNAME From (SubGroup Left Join City C on C.CityCode=SubGroup.CityCode) LEFT JOIN ACGROUP A ON A.GROUPCODE=SubGroup.GROUPCODE " & mQRYx & " order by SubGroup.Name", G_FaCn, adOpenForwardOnly, adLockReadOnly
        ElseIf PubBackEnd = "S" Then
            RstAccount.Open "Select SubCode AS Code,Name,RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,'')) AS Address,CityName,SubGroup.GroupCode,A.GROUPNAME From (SubGroup Left Join City C on C.CityCode=SubGroup.CityCode) LEFT JOIN ACGROUP A ON A.GROUPCODE=SubGroup.GROUPCODE " & mQRYx & " order by SubGroup.Name", G_FaCn, adOpenForwardOnly, adLockReadOnly
        End If
        Set DGAccount.DataSource = RstAccount
        If RstAccount.RecordCount > 0 Then
            FGrid.TextMatrix(List2, 1) = RstAccount!Name
            FGrid.TextMatrix(List2, 2) = RstAccount!Code
            FGrid.TextMatrix(List3, 1) = RstAccount!Name
            FGrid.TextMatrix(List3, 2) = RstAccount!Code
        End If
        Set RstGroup = New ADODB.Recordset
        RstGroup.CursorLocation = adUseClient
        RstGroup.Open "Select GROUPCODE AS Code,GROUPNAME AS NAME From AcGroup WHERE AliasYN<> 'Y' order by GroupName", G_FaCn, adOpenForwardOnly, adLockReadOnly
        Set DGGroup.DataSource = RstGroup
        If PubBackEnd = "A" Then
            GridInitialise 1, "SELECT '' as O,NAME as Account,SUBCODE as AccId,GNAME AS GroupName,RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3))+RTRIM(IIF(ISNULL(CityName),'',','+CityName)) AS Address From ViewSubgroup " & mQRYy & " order by name"
        ElseIf PubBackEnd = "S" Then
            GridInitialise 1, "SELECT '' as O,NAME as Account,SUBCODE as AccId,GNAME AS GroupName,RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,''))+RTRIM(ISNULL(CityName,'')) AS Address From ViewSubgroup " & mQRYy & " order by name"
        End If
        GridInitialise 2, "SELECT '' as O,GroupName,GroupCode as Code from AcGroup Order by GroupName"
        FGrid.height = 5 * 300
        GridSel(2).Visible = False
        Check1(2).Visible = False
        If GRepFormName = "AcCheckList" Then
            GridInitialise 3, "SELECT '' as O,Description AS VrType,V_TYPE as Code from Voucher_type Order by Description"
        End If
    Case "Annexure"
        With FGrid
            .TextMatrix(Date1, 0) = "From Date        ": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date        ": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "For Account      ": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Amount Slab Wise ": .RowHeight(List2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List2, 1) = "No"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List2: mHelpGridNo = 1
        FGrid.ColWidth(0) = 2000: FGrid.ColWidth(1) = 3500
        FGrid.width = FGrid.ColWidth(0) + FGrid.ColWidth(1) + 100
        FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75
        Set RstGroup = New ADODB.Recordset
        RstGroup.CursorLocation = adUseClient
        RstGroup.Open "Select GROUPCODE AS Code,GROUPNAME AS NAME From AcGroup WHERE AliasYN<> 'Y' order by GroupName", G_FaCn, adOpenForwardOnly, adLockReadOnly
        Set DGGroup.DataSource = RstGroup
        If RstGroup.RecordCount > 0 Then
            FGrid.TextMatrix(List1, 1) = RstGroup!Name
            FGrid.TextMatrix(List1, 2) = RstGroup!Code
            GridInitialise 1, "SELECT '' as O,NAME as Account,SUBCODE as AccId from SUBGROUP where GROUPCODE='" & FGrid.TextMatrix(List1, 2) & "' " & mSiteHlpAndSubgroup & " Order by Name"
        End If
    Case "DetailedTrial"
        With FGrid
            .TextMatrix(Date1, 0) = "From Date      ": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date      ": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 0) = "With Op.Bal.   ": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List1, 1) = "Yes"
            .TextMatrix(List2, 0) = "Datailed       ": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List2, 1) = "Yes"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List2: mHelpGridNo = 0
        FGrid.ColWidth(0) = 1500: FGrid.ColWidth(1) = 1500
        FGrid.width = FGrid.ColWidth(0) + FGrid.ColWidth(1) + 100
        FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75
    Case "DayBook"
        With FGrid
            .TextMatrix(Date1, 0) = "From Date         ": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date         ": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Voucher Wise Total": .RowHeight(List1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "No"
        End With
        FGrid.ColWidth(0) = 2500: FGrid.ColWidth(1) = 1500
        FGrid.width = FGrid.ColWidth(0) + FGrid.ColWidth(1) + 100
        FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75
        If PubFaSiteType <> 0 Then
            FGrid.TextMatrix(List2, 0) = "For Site": FGrid.RowHeight(List2) = GridRowHeight
            FGrid.TextMatrix(List2, 1) = ""
            Set RstSite = New ADODB.Recordset
            RstSite.CursorLocation = adUseClient
            RstSite.Open "Select Site_Code AS Code,Site_Desc AS NAME From Site Order by Site_Desc", G_FaCn, adOpenForwardOnly, adLockReadOnly
            Set DGSite.DataSource = RstSite
            mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List2: mHelpGridNo = 0
        Else
            mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List1: mHelpGridNo = 0
        End If
    Case "CashBook", "BankBook", "RozNamcha"
        mQRYx = ""
        Select Case GRepFormName
            Case "CashBook", "RozNamcha"
                mQRYx = " WHERE SubGroup.Nature='Cash' AND A.Nature='Cash'" + mSiteHlpAndSubgroup
            Case "BankBook"
                mQRYx = " WHERE SubGroup.Nature='Bank' AND A.Nature='Bank'" + mSiteHlpAndSubgroup
        End Select
        Select Case GRepFormName
            Case "CashBook", "RozNamcha"
                With FGrid
                    .TextMatrix(Date1, 0) = "From Date           ": .RowHeight(Date1) = GridRowHeight
                    .TextMatrix(Date2, 0) = "UpTo Date           ": .RowHeight(Date2) = GridRowHeight
                    .TextMatrix(List1, 0) = "As Per Detail       ": .RowHeight(List1) = GridRowHeight
                    .TextMatrix(List2, 0) = "Starting Page No.   ": .RowHeight(List2) = GridRowHeight
                    .TextMatrix(List3, 0) = "For Account         ": .RowHeight(List3) = GridRowHeight
                    .TextMatrix(Cat1, 0) = "Detail With Narration": .RowHeight(Cat1) = GridRowHeight
                    .TextMatrix(Date1, 1) = PubStartDate
                    .TextMatrix(Date2, 1) = PubLoginDate
                    .TextMatrix(List1, 1) = "No"
                    .TextMatrix(List2, 1) = "1"
                    .TextMatrix(List3, 1) = ""
                    .TextMatrix(Cat1, 1) = "No"
                End With
                mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Cat1: mHelpGridNo = 0
            Case Else
                With FGrid
                    .TextMatrix(Date1, 0) = "From Date           ": .RowHeight(Date1) = GridRowHeight
                    .TextMatrix(Date2, 0) = "UpTo Date           ": .RowHeight(Date2) = GridRowHeight
                    .TextMatrix(List1, 0) = "As Per Detail       ": .RowHeight(List1) = GridRowHeight
                    .TextMatrix(List2, 0) = "Starting Page No.   ": .RowHeight(List2) = GridRowHeight
                    .TextMatrix(List3, 0) = "For Account         ": .RowHeight(List3) = GridRowHeight
                    .TextMatrix(Date1, 1) = PubStartDate
                    .TextMatrix(Date2, 1) = PubLoginDate
                    .TextMatrix(List1, 1) = "No"
                    .TextMatrix(List2, 1) = "1"
                    .TextMatrix(List3, 1) = ""
                End With
                mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List3: mHelpGridNo = 0
        End Select
        FGrid.ColWidth(0) = 2000: FGrid.ColWidth(1) = 3500
        FGrid.width = FGrid.ColWidth(0) + FGrid.ColWidth(1) + 100
        FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75
        Set RstAccount = New ADODB.Recordset
        RstAccount.CursorLocation = adUseClient
        If PubBackEnd = "A" Then
            RstAccount.Open "Select SubCode AS Code,Name,RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3)) AS Address,CityName,SubGroup.GroupCode,A.GROUPNAME From (SubGroup Left Join City C on C.CityCode=SubGroup.CityCode) LEFT JOIN ACGROUP A ON A.GROUPCODE=SubGroup.GROUPCODE " & mQRYx & " order by SubGroup.Name", G_FaCn, adOpenForwardOnly, adLockReadOnly
        ElseIf PubBackEnd = "S" Then
            RstAccount.Open "Select SubCode AS Code,Name,RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,'')) AS Address,CityName,SubGroup.GroupCode,A.GROUPNAME From (SubGroup Left Join City C on C.CityCode=SubGroup.CityCode) LEFT JOIN ACGROUP A ON A.GROUPCODE=SubGroup.GROUPCODE " & mQRYx & " order by SubGroup.Name", G_FaCn, adOpenForwardOnly, adLockReadOnly
        End If
        Set DGAccount.DataSource = RstAccount
    Case "JournalBook", "DailySumm", "Clg", "ClgNot"
        With FGrid
            .TextMatrix(Date1, 0) = "From Date   ": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date   ": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = ""
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List1: mHelpGridNo = 0
        FGrid.ColWidth(0) = 1700: FGrid.ColWidth(1) = 2500
        FGrid.width = FGrid.ColWidth(0) + FGrid.ColWidth(1) + 100
        FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75
        Select Case GRepFormName
            Case "JournalBook"
                Set RstGroup = New ADODB.Recordset
                RstGroup.CursorLocation = adUseClient
                FGrid.TextMatrix(List1, 0) = "Voucher Type": FGrid.RowHeight(List1) = GridRowHeight
                FGrid.TextMatrix(List2, 0) = "Day Total": FGrid.RowHeight(List2) = GridRowHeight
                FGrid.TextMatrix(List2, 1) = "No"
                mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List2: mHelpGridNo = 0
                RstGroup.Open "Select V_TYPE AS Code,Description+' Book' AS Name from Voucher_type WHERE Category='FA' AND V_TYPE<>'F_AO' order by Description+' Book'", G_FaCn, adOpenForwardOnly, adLockReadOnly
                Set DGGroup.DataSource = RstGroup
                If PubFaSiteType <> 0 Then
                    FGrid.TextMatrix(List3, 0) = "For Site": FGrid.RowHeight(List3) = GridRowHeight
                    FGrid.TextMatrix(List3, 1) = ""
                    Set RstSite = New ADODB.Recordset
                    RstSite.CursorLocation = adUseClient
                    RstSite.Open "Select Site_Code AS Code,Site_Desc AS NAME From Site Order by Site_Desc", G_FaCn, adOpenForwardOnly, adLockReadOnly
                    Set DGSite.DataSource = RstSite
                    mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List3: mHelpGridNo = 0
                End If
            Case "DailySumm"
                Set RstAccount = New ADODB.Recordset
                RstAccount.CursorLocation = adUseClient
                FGrid.TextMatrix(List1, 0) = "For Account": FGrid.RowHeight(List1) = GridRowHeight
                If PubBackEnd = "A" Then
                    RstAccount.Open "Select SubCode AS Code,Name,RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3)) AS Address,CityName,SubGroup.GroupCode,A.GROUPNAME From (SubGroup Left Join City C on C.CityCode=SubGroup.CityCode) LEFT JOIN ACGROUP A ON A.GROUPCODE=SubGroup.GROUPCODE " & mSiteHlpSubgroup & " order by SubGroup.Name", G_FaCn, adOpenForwardOnly, adLockReadOnly
                ElseIf PubBackEnd = "S" Then
                    RstAccount.Open "Select SubCode AS Code,Name,RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,'')) AS Address,CityName,SubGroup.GroupCode,A.GROUPNAME From (SubGroup Left Join City C on C.CityCode=SubGroup.CityCode) LEFT JOIN ACGROUP A ON A.GROUPCODE=SubGroup.GROUPCODE " & mSiteHlpSubgroup & " order by SubGroup.Name", G_FaCn, adOpenForwardOnly, adLockReadOnly
                End If
                Set DGAccount.DataSource = RstAccount
            Case "Clg", "ClgNot"
                Set RstAccount = New ADODB.Recordset
                RstAccount.CursorLocation = adUseClient
                FGrid.TextMatrix(List1, 0) = "For Account": FGrid.RowHeight(List1) = GridRowHeight
                If PubBackEnd = "A" Then
                    RstAccount.Open "Select SubCode AS Code,Name,RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3)) AS Address,CityName,SubGroup.GroupCode,A.GROUPNAME From (SubGroup Left Join City C on C.CityCode=SubGroup.CityCode) LEFT JOIN ACGROUP A ON A.GROUPCODE=SubGroup.GROUPCODE Where SubGroup.nature='Bank' AND A.nature='Bank' " & mSiteHlpAndSubgroup & " order by SubGroup.Name", G_FaCn, adOpenForwardOnly, adLockReadOnly
                ElseIf PubBackEnd = "S" Then
                    RstAccount.Open "Select SubCode AS Code,Name,RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,'')) AS Address,CityName,SubGroup.GroupCode,A.GROUPNAME From (SubGroup Left Join City C on C.CityCode=SubGroup.CityCode) LEFT JOIN ACGROUP A ON A.GROUPCODE=SubGroup.GROUPCODE Where SubGroup.nature='Bank' AND A.nature='Bank'  " & mSiteHlpAndSubgroup & " order by SubGroup.Name", G_FaCn, adOpenForwardOnly, adLockReadOnly
                End If
                Set DGAccount.DataSource = RstAccount
        End Select
    Case "NonTrans", "BankReg", "CONTROLLED"
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "To Date  ": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date2: mHelpGridNo = 1
        FGrid.ColWidth(0) = 1200: FGrid.ColWidth(1) = 1500
        FGrid.width = FGrid.ColWidth(0) + FGrid.ColWidth(1) + 100
        FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75
        Select Case GRepFormName
            Case "NonTrans", "CONTROLLED"
                GridInitialise 1, "SELECT '' as O,GroupName as Account,GroupCode as AccId from AcGroup WHERE AliasYN<> 'Y' order by GroupName"
            Case "BankReg"
                GridInitialise 1, "SELECT '' as O,NAME as Account,SUBCODE as AccId from SUBGROUP where nature='Bank' " & mSiteHlpAndSubgroup & " order by name "
        End Select
    Case "RefReport"
        With FGrid
            .TextMatrix(List2, 0) = "Pending Only": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List2, 1) = "Yes"
        End With
        mFirstRow = List2: FGrid.Row = mFirstRow: mLastRow = List2: mHelpGridNo = 1
        GridInitialise 1, "SELECT '' as O,NAME as Account,SUBCODE as AccId from SUBGROUP  " & mSiteHlpSubgroup & " Order by Name"
    Case "AgingDr", "AgingCr"
        With FGrid
            .TextMatrix(Date1, 0) = "UpTo Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = Date1: mHelpGridNo = 1
        FGrid.ColWidth(0) = 1200: FGrid.ColWidth(1) = 1500
        FGrid.width = FGrid.ColWidth(0) + FGrid.ColWidth(1) + 100
        FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75
        If GRepFormName = "AgingDr" Then
            GridInitialise 1, "SELECT '' as O,GroupName as Account,GroupCode as AccId from AcGroup Where AliasYN<>'Y' AND Nature='Customer' order by GroupName"
        ElseIf GRepFormName = "AgingCr" Then
            GridInitialise 1, "SELECT '' as O,GroupName as Account,GroupCode as AccId from AcGroup Where AliasYN<>'Y' AND  Nature='Supplier' order by GroupName"
        Else
            GridInitialise 1, "SELECT '' as O,GroupName as Account,GroupCode as AccId from AcGroup Where AliasYN<>'Y' Order By GroupName"
        End If
    Case "DUELIST"
        With FGrid
            .TextMatrix(Date1, 0) = "UpTo Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date1, 1) = PubLoginDate
            .TextMatrix(List1, 0) = "All/Pending": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List1, 1) = "Pending"
            .TextMatrix(List2, 0) = "More Than Days": .RowHeight(List2) = GridRowHeight
            .TextMatrix(List2, 1) = "0"
        End With
        mFirstRow = Date1: FGrid.Row = mFirstRow: mLastRow = List1: mHelpGridNo = 1
        mLastRow = List2
        If PubBackEnd = "A" Then
            GridInitialise 1, "Select '' as O,SubGroup.Name as PartyName,SubGroup.Subcode as Code,RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(CityName) OR LEN(CityName)=0,'',','+CityName)) AS NameWithADDR From SubGroup Left Join City C on SubGroup.CityCode=C.CityCode Where SubGroup.Nature='Customer' " & mSiteHlpAndSubgroup & " Order by SubGroup.Name"
        ElseIf PubBackEnd = "S" Then
            GridInitialise 1, "Select '' as O,SubGroup.Name as PartyName,SubGroup.Subcode as Code,RTrim(IsNull(Add1,''))+','+RTrim(IsNull(Add2,''))+','+RTrim(IsNull(CityName,'')) AS NameWithADDR From SubGroup Left Join City C on SubGroup.CityCode=C.CityCode Where SubGroup.Nature='Customer' " & mSiteHlpAndSubgroup & " Order by SubGroup.Name"
        End If
        GridSel(1).ColWidth(3) = 0
End Select
End Sub
Private Sub RefReport(Index)
On Error GoTo ELoop
Dim RstTmp As ADODB.Recordset, Condstr As String, Rst1 As ADODB.Recordset, Rst2 As ADODB.Recordset, RST3 As ADODB.Recordset, Rst4 As ADODB.Recordset, Condstr2 As String
Dim X11, I As Integer, mSiteCode As String
GridString1 = ""
Condstr = ""
If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
If GridString1 <> "" Then Condstr = " Where SUBCODE IN (" & GridString1 & ")"
If GridString1 <> "" Then Condstr2 = " Where Ledger.SUBCODE IN (" & GridString1 & ")"
If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(LEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And LEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
End If
Set RstTmp = New ADODB.Recordset
Set RstTmp = PubDatamanFa.FaAdjRst(RstTmp)
Set Rst1 = G_FaCn.Execute("SELECT SUBCODE,NAME FROM SUBGROUP  " & Condstr & " ORDER BY SUBCODE")
Set Rst2 = G_FaCn.Execute("SELECT Ledger.*,LEDGERREF.AGREFTYPE FROM Ledger LEFT JOIN ledgerRef ON (Ledger.V_SNo=ledgerRef.V_SNo) AND (Ledger.DocId = ledgerRef.DocId) " & Condstr2 & " " & mSiteCode & " ORDER BY Ledger.SUBCODE,Ledger.V_DATE,Ledger.DOCId,Ledger.V_SNO")
Set RST3 = G_FaCn.Execute("SELECT LedgerRef.*  FROM LedgerRef " & Condstr & " ORDER BY SUBCode,V_DATE,DocId,V_SNO")

If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
If Rst2.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
Do Until Rst1.EOF
    If RST3.RecordCount > 0 Then
        RST3.MoveFirst
        RST3.FIND "SubCode='" & Rst1!SubCode & "'"
        If RST3.EOF = False Then
            Do While True
                If FGrid.TextMatrix(List2, 1) = "Yes" And RST3!AgRefType <> "On Account" Then
                    Set Rst4 = G_FaCn.Execute("SELECT AgRefNo,SUM(Cr)-SUM(Dr) AS BalanceAmt From ledgerRef " & Condstr + IIf(Condstr = "", " Where ", " AND ") & " AGREFTYPE IN ('Advance','New Ref','Ag.Ref.') AND AGREFNO='" & RST3!AgRefNo & "' GROUP BY SubCode,AgRefNo")
                    If Rst4.RecordCount > 0 Then
                        If Rst4!BalanceAmt = 0 Then GoTo NextGo
                    End If
                End If
                RstTmp.AddNew
                RstTmp!DocId = RST3!DocId
                RstTmp!V_SNo = RST3!V_SNo
                RstTmp!V_tYPE = mID(RST3!DocId, 4, 5)
                RstTmp!V_NO = Right(RST3!DocId, 8)
                RstTmp!V_DATE = RST3!V_DATE
                RstTmp!SubCode = RST3!SubCode
                RstTmp!Name = Rst1!Name
                RstTmp!cr = RST3!cr
                RstTmp!dr = RST3!dr
                Select Case RST3!AgRefType
                    Case "Advance", "New Ref"
                        RstTmp!AgRefType = RST3!AgRefType
                        RstTmp!AgRefNo = Trim(RST3!AgRefNo)
                        RstTmp!RefSort = "1"
                        RstTmp!GRPSort = "1"
                    Case "Ag.Ref."
                        RstTmp!AgRefType = RST3!AgRefType
                        RstTmp!AgRefNo = Trim(RST3!AgRefNo)
                        RstTmp!GRPSort = "1"
                        RstTmp!RefSort = "2"
                    Case "On Account"
                        RstTmp!AgRefNo = "On Account"
                        RstTmp!AgRefType = ""
                        RstTmp!GRPSort = "5"
                        RstTmp!RefSort = "1"
                End Select
                RstTmp.Update
NextGo:
                RST3.MoveNext
                If RST3.EOF = True Then Exit Do
                If Rst1!SubCode <> RST3!SubCode Then Exit Do
            Loop
        End If
    End If
    Rst2.MoveFirst
    Rst2.FIND "SubCode='" & Rst1!SubCode & "'"
    If Rst2.EOF = False Then
        Do While True
            If IsNull(Rst2!AgRefType) Or Rst2!AgRefType = "" Then
                RstTmp.AddNew
                RstTmp!DocId = Rst2!DocId
                RstTmp!V_SNo = Rst2!V_SNo
                RstTmp!V_tYPE = Rst2!V_tYPE
                RstTmp!V_NO = Rst2!V_NO
                RstTmp!V_DATE = Rst2!V_DATE
                RstTmp!SubCode = Rst2!SubCode
                RstTmp!Name = Rst1!Name
                RstTmp!cr = Rst2!AmtCr
                RstTmp!dr = Rst2!AmtDr
                RstTmp!AgRefNo = "On Account"
                RstTmp!AgRefType = ""
                RstTmp!GRPSort = "5"
                RstTmp!RefSort = "1"
                RstTmp.Update
            End If
            Rst2.MoveNext
            If Rst2.EOF = True Then Exit Do
            If Rst1!SubCode <> Rst2!SubCode Then Exit Do
        Loop
    End If
    Rst1.MoveNext
Loop
If RstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
X11 = CreateFieldDefFile(RstTmp, PubFaReportPath + "\FaAgRefNo.ttx", True)
Set rpt = rdApp.OpenReport(PubFaReportPath + "\FaAgRefNo.RPT")

'Set rpt = PubDatamanFa.FaAgRefNoRpt
rpt.Database.SetDataSource RstTmp
EXIT_SUB:
    Set RstTmp = Nothing
    Set Rst1 = Nothing
    Set Rst2 = Nothing
    Set RST3 = Nothing
    Set Rst4 = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub Annexure(Index As Integer)
On Error GoTo ELoop
Dim AGE_rs As ADODB.Recordset, RstTmp As ADODB.Recordset, Rst1 As ADODB.Recordset
Dim TinTin As Integer, mQRY As String, Condstr As String, CondStr1 As String, ac_str As String, dr As Double, cr As Double, X As Double, I As Integer, mAcCode As String
Dim ARR(7) As Double, Dr1 As Double, DR2 As Double, DR3 As Double, DR4 As Double, DR5 As Double, DR6 As Double, DR7 As Double, DAYS As Integer, mQRY1 As String, mNARR1 As String
Dim mSiteCode As String, mSiteCode1 As String

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
GridString1 = ""
Condstr = ""
CondStr1 = ""
If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub

If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(LEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
        mSiteCode1 = " And RIGHT(ViewLedger.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And LEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
        mSiteCode1 = " And LEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
    mSiteCode1 = ""
End If

TXTACC_CODE = Trim(FGrid.TextMatrix(List1, 1))
mAcCode = Trim(FGrid.TextMatrix(List1, 2))
If GridString1 <> "" Then Condstr = " AND VIEWLEDGER.PARTY IN (" & GridString1 & ")"
If GridString1 <> "" Then CondStr1 = " AND VIEWSUBGROUP.SUBCODE IN (" & GridString1 & ")"
Set AGE_rs = G_FaCn.Execute("SELECT * FROM FaEnviro")
mQRY1 = ""
If FGrid.TextMatrix(List2, 1) = "No" Then
    Set Rst1 = G_FaCn.Execute("SELECT MAINGRCODE,GroupNature FROM ACGROUP WHERE GROUPCODE='" & mAcCode & "' AND ALIASYN='N'")
    If Rst1.RecordCount > 0 Then
        mQRY1 = Rst1!MainGrCode
        mNARR1 = Rst1!GroupNature
        Set Rst1 = New ADODB.Recordset
        Rst1.Sort = "PARTYNAME ASC,GrCode ASC,PARTYCODE ASC"
        Select Case mNARR1
            Case "E", "R"
                If PubBackEnd = "A" Then
                    Rst1.Open ("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS OP_CR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS OP_dR,0 AS BALANCECR,0 AS BALANCEDR,0 As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & CondStr1 & " AND ViewSubgroup.GROUPCODE='" & mAcCode & "' AND ACGROUP.AliasYN='N' " & mSiteCode & " GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE " & _
                    "Union SELECT 2 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS BALANCECR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS BALANCEDR,0 As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " " & CondStr1 & " AND ViewSubgroup.GROUPCODE='" & mAcCode & "' AND ACGROUP.AliasYN='N'  " & mSiteCode & " GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE " & _
                    "Union  SELECT 3 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS PARTYCODE,ACGROUP.GROUPNAME AS PARTYNAME,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS OP_CR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) As OP_DR,0 AS BALANCECR,0 AS BALANCEDR,0 AS Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & mQRY1 & "'))='" & mQRY1 & "' AND LEN(MAINGRCODE)=LEN('" & mQRY1 & "')+3  " & mSiteCode & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME " & _
                    "Union SELECT 4 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS PARTYCODE,ACGROUP.GROUPNAME AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS BALANCECR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS BALANCEDR,0 As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & mQRY1 & "'))='" & mQRY1 & "' AND LEN(MAINGRCODE)=LEN('" & mQRY1 & "')+3  " & mSiteCode & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME"), G_FaCn, adOpenDynamic, adLockOptimistic
                ElseIf PubBackEnd = "S" Then
                    Rst1.Open ("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,ISNULL(sum(AMTCR),0) AS OP_CR,ISNULL(SUM(AMTDR),0) AS OP_dR,0 AS BALANCECR,0 AS BALANCEDR,0 As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & CondStr1 & " AND ViewSubgroup.GROUPCODE='" & mAcCode & "' AND ACGROUP.AliasYN='N'  " & mSiteCode & " GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE " & _
                    "Union SELECT 2 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,ISNULL(sum(AMTCR),0) AS BALANCECR,ISNULL(SUM(AMTDR),0) AS BALANCEDR,0 As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " " & CondStr1 & " AND ViewSubgroup.GROUPCODE='" & mAcCode & "' AND ACGROUP.AliasYN='N'  " & mSiteCode & " GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE " & _
                    "Union  SELECT 3 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS PARTYCODE,ACGROUP.GROUPNAME AS PARTYNAME,ISNULL(sum(AMTCR),0) AS OP_CR,ISNULL(SUM(AMTDR),0) As OP_DR,0 AS BALANCECR,0 AS BALANCEDR,0 AS Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & mQRY1 & "'))='" & mQRY1 & "' AND LEN(MAINGRCODE)=LEN('" & mQRY1 & "')+3  " & mSiteCode & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME " & _
                    "Union  SELECT 4 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS PARTYCODE,ACGROUP.GROUPNAME AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,ISNULL(sum(AMTCR),0) AS BALANCECR,ISNULL(SUM(AMTDR),0) AS BALANCEDR,0 As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & mQRY1 & "'))='" & mQRY1 & "' AND LEN(MAINGRCODE)=LEN('" & mQRY1 & "')+3  " & mSiteCode & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME"), G_FaCn, adOpenDynamic, adLockOptimistic
                End If
            Case "A", "L"
                If PubBackEnd = "A" Then
                    Rst1.Open ("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS OP_CR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS OP_dR,0 AS BALANCECR,0 AS BALANCEDR,0 As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE<" & FaConvertDate(TXTS_DATE) & " " & CondStr1 & " AND ViewSubgroup.GROUPCODE='" & mAcCode & "' AND ACGROUP.AliasYN='N'  " & mSiteCode & " GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE " & _
                    "Union SELECT 2 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS BALANCECR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS BALANCEDR,0 As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " " & CondStr1 & " AND ViewSubgroup.GROUPCODE='" & mAcCode & "' AND ACGROUP.AliasYN='N'  " & mSiteCode & " GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE " & _
                    "Union  SELECT 3 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS PARTYCODE,ACGROUP.GROUPNAME AS PARTYNAME,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS OP_CR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) As OP_DR,0 AS BALANCECR,0 AS BALANCEDR,0 AS Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE<" & FaConvertDate(TXTS_DATE) & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & mQRY1 & "'))='" & mQRY1 & "' AND LEN(MAINGRCODE)=LEN('" & mQRY1 & "')+3  " & mSiteCode & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME " & _
                    "Union  SELECT 4 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS PARTYCODE,ACGROUP.GROUPNAME AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS BALANCECR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS BALANCEDR,0 As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & mQRY1 & "'))='" & mQRY1 & "' AND LEN(MAINGRCODE)=LEN('" & mQRY1 & "')+3  " & mSiteCode & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME"), G_FaCn, adOpenDynamic, adLockOptimistic
                ElseIf PubBackEnd = "S" Then
                    Rst1.Open ("SELECT 1 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,ISNULL(sum(AMTCR),0) AS OP_CR,ISNULL(SUM(AMTDR),0) AS OP_dR,0 AS BALANCECR,0 AS BALANCEDR,0 As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE<" & FaConvertDate(TXTS_DATE) & " " & CondStr1 & " AND ViewSubgroup.GROUPCODE='" & mAcCode & "' AND ACGROUP.AliasYN='N'  " & mSiteCode & " GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE " & _
                    "Union SELECT 2 AS TT,MAX(ACGROUP.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,ISNULL(sum(AMTCR),0) AS BALANCECR,ISNULL(SUM(AMTDR),0) AS BALANCEDR,0 As Bal FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=VIEWSUBGROUP.GROUPCODE WHERE V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " " & CondStr1 & " AND ViewSubgroup.GROUPCODE='" & mAcCode & "' AND ACGROUP.AliasYN='N'  " & mSiteCode & " GROUP BY ACGROUP.GROUPNAME,LEDGER.SUBCODE " & _
                    "Union  SELECT 3 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS PARTYCODE,ACGROUP.GROUPNAME AS PARTYNAME,ISNULL(SUM(AMTCR),0) AS OP_CR,ISNULL(sum(AMTDR),0) As OP_DR,0 AS BALANCECR,0 AS BALANCEDR,0 AS Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE<" & FaConvertDate(TXTS_DATE) & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & mQRY1 & "'))='" & mQRY1 & "' AND LEN(MAINGRCODE)=LEN('" & mQRY1 & "')+3  " & mSiteCode & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME " & _
                    "Union  SELECT 4 AS TT,ACGROUP.GROUPNAME,MAX(ACGROUP.GROUPCODE) AS GrCode,'' AS PARTYCODE,ACGROUP.GROUPNAME AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,ISNULL(sum(AMTCR),0) AS BALANCECR,ISNULL(SUM(AMTDR),0) AS BALANCEDR,0 As Bal FROM (ACGROUP INNER JOIN ViewSubgroup ON ACGROUP.MAINGRCODE=" & IIf(PubBackEnd = "S", "SUBSTRING", "MID") & " (ViewSubgroup.MAINGRCODES,1,LEN(ACGROUP.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE WHERE V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND ACGROUP.ALIASYN='N' AND ViewSubgroup.AliasYN='N' AND LEFT(MAINGRCODE,LEN('" & mQRY1 & "'))='" & mQRY1 & "' AND LEN(MAINGRCODE)=LEN('" & mQRY1 & "')+3  " & mSiteCode & " GROUP BY ACGROUP.MAINGRCODE,ACGROUP.GROUPNAME"), G_FaCn, adOpenDynamic, adLockOptimistic
                End If
        End Select
    End If
    If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    Select Case Index
        Case 1
            MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
            TinTin = PubDatamanFa.FaAnnexureDosPrinting(Me, Rst1)
            GoTo EXIT_SUB
        Case Else
'            X1 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaANNEXURE.ttx", True)
            Set rpt = PubDatamanFa.FaAnnexureRpt
            rpt.Database.SetDataSource Rst1
    End Select
Else
    Set RstTmp = New ADODB.Recordset
    Set RstTmp = PubDatamanFa.FaAgeTmp(RstTmp)
    Set Rst1 = New ADODB.Recordset
    Rst1.Sort = "PARTYNAME ASC"
    Rst1.Open "SELECT MAX(ACGROUP.GROUPNAME)As GroupName,MAX(PARTY) AS PARTYCODE,MAX(PARTY_NAME) AS PARTYNAME,MAX(ADD1) AS ADDR1,MAX(ADD2) AS ADDR2,MAX(CITY_NAME) AS CITYNAME,sum(credit)-SUM(DEBIT) As Bal FROM (VIEWLEDGER LEFT JOIN PARTY_LIST ON PARTY_LIST.SUBCODE=VIEWLEDGER.PARTY) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=PARTY_LIST.GROUPCODE WHERE V_DATE<" & FaConvertDate(TXTE_DATE) & " " & Condstr & " AND PARTY_LIST.GroupNature in ('A','L') AND ALIASYN='N' AND CODE='" & mAcCode & "'  " & mSiteCode1 & " GROUP BY ACGROUP.GROUPNAME,PARTY " & _
    " Union SELECT MAX(ACGROUP.GROUPNAME)As GroupName,MAX(PARTY) AS PARTYCODE,MAX(PARTY_NAME) AS PARTYNAME,MAX(ADD1) AS ADDR1,MAX(ADD2) AS ADDR2,MAX(CITY_NAME) AS CITYNAME,sum(credit)-SUM(DEBIT) As Bal FROM (VIEWLEDGER LEFT JOIN PARTY_LIST ON PARTY_LIST.SUBCODE=VIEWLEDGER.PARTY) LEFT JOIN ACGROUP ON ACGROUP.GROUPCODE=PARTY_LIST.GROUPCODE WHERE V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(TXTE_DATE) & " " & Condstr & " AND PARTY_LIST.GroupNature NOT in ('A','L') AND ALIASYN='N' AND CODE='" & mAcCode & "'  " & mSiteCode1 & " GROUP BY ACGROUP.GROUPNAME,PARTY", G_FaCn, adOpenDynamic, adLockOptimistic
    Do Until Rst1.EOF
        If Rst1!Bal <> 0 Then
            RstTmp.AddNew
            RstTmp!ACC_NAME = left(Rst1!PartyName, 35)
            RstTmp!AName = Rst1!GroupName
            If Abs(Rst1!Bal) <= AGE_rs!Amt1 Then
                RstTmp!DEBIT1 = Rst1!Bal
            ElseIf Abs(Rst1!Bal) > AGE_rs!Amt1 And Abs(Rst1!Bal) <= AGE_rs!Amt2 Then
                RstTmp!DEBIT2 = Rst1!Bal
            ElseIf Abs(Rst1!Bal) > AGE_rs!Amt2 And Abs(Rst1!Bal) <= AGE_rs!Amt3 Then
                RstTmp!DEBIT3 = Rst1!Bal
            ElseIf Abs(Rst1!Bal) > AGE_rs!Amt3 And Abs(Rst1!Bal) <= AGE_rs!Amt4 Then
                RstTmp!DEBIT4 = Rst1!Bal
            ElseIf Abs(Rst1!Bal) > AGE_rs!Amt4 And Abs(Rst1!Bal) <= AGE_rs!Amt5 Then
                RstTmp!DEBIT5 = Rst1!Bal
            ElseIf Abs(Rst1!Bal) > AGE_rs!Amt5 And Abs(Rst1!Bal) <= AGE_rs!Amt6 Then
                RstTmp!DEBIT6 = Rst1!Bal
            Else
                RstTmp!DEBIT = Rst1!Bal
            End If
            RstTmp.Update
        End If
        Rst1.MoveNext
    Loop
    If RstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    Select Case Index
        Case 1
            MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
            TinTin = PubDatamanFa.FaAnnexure2DosPrinting(Me, RstTmp)
            GoTo EXIT_SUB
        Case Else
'            X1 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaANNEXURE2.ttx", True)
            Set rpt = PubDatamanFa.FaAnnexure2Rpt
            For I = 1 To rpt.FormulaFields.Count
                Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                    Case "P1"
                        rpt.FormulaFields(I).TEXT = "'" & "0 - " + STR(AGE_rs!Amt1) & "'"
                    Case "P2"
                        rpt.FormulaFields(I).TEXT = "'" & STR(AGE_rs!Amt1 + 1) & " - " & STR(AGE_rs!Amt2) & "'"
                    Case "P3"
                        rpt.FormulaFields(I).TEXT = "'" & STR(AGE_rs!Amt2 + 1) + " - " + STR(AGE_rs!Amt3) & "'"
                    Case "P4"
                        rpt.FormulaFields(I).TEXT = "'" & STR(AGE_rs!Amt3 + 1) + " - " + STR(AGE_rs!Amt4) & "'"
                    Case "P5"
                        rpt.FormulaFields(I).TEXT = "'" & STR(AGE_rs!Amt4 + 1) + " - " + STR(AGE_rs!Amt5) & "'"
                    Case "P6"
                        rpt.FormulaFields(I).TEXT = "'" & STR(AGE_rs!Amt5 + 1) + " - " + STR(AGE_rs!Amt6) & "'"
                    Case "P7"
                        rpt.FormulaFields(I).TEXT = "'Above " & STR(AGE_rs!Amt6) & "'"
                End Select
            Next
            rpt.Database.SetDataSource RstTmp
    End Select
End If
EXIT_SUB:
    Set AGE_rs = Nothing
    Set RstTmp = Nothing
    Set Rst1 = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub Aging(Index As Integer, DrCrType As String)
On Error GoTo ELoop
Dim AGE_rs As ADODB.Recordset, SUBGROUP_rs As ADODB.Recordset, RstTmp As ADODB.Recordset, mGROUP_rs As ADODB.Recordset
Dim TinTin As Integer, mQRY As String, Condstr As String, ac_str As String, dr As Double, cr As Double, X As Double, I As Integer
Dim ARR(7) As Double, Dr1 As Double, DR2 As Double, DR3 As Double, DR4 As Double, DR5 As Double, DR6 As Double, DR7 As Double, DAYS As Integer, mSiteCode As String

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And VIEWLEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
End If
GridString1 = ""
Condstr = ""
If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
If GridString1 <> "" Then Condstr = " AND ACGROUP.GROUPCODE IN (" & GridString1 & ")"
TXTE_DATE = FGrid.TextMatrix(Date1, 1)
If FaValidDateChk(TXTE_DATE, "Date Upto") = False Then RepPrint = False: Exit Sub
Set AGE_rs = G_FaCn.Execute("SELECT * FROM FaEnviro")
If AGE_rs.RecordCount = 0 Then MsgBox " ** Please Set Ageing Parameters ** ", vbInformation, Me.CAPTION: Exit Sub
If DrCrType = "Dr" Then Condstr = Condstr + " AND AcGroup.Nature='Customer'"
If DrCrType = "Cr" Then Condstr = Condstr + " AND AcGroup.Nature='Supplier'"
If PubSiteCodeWiseHelp = True Then
    Set SUBGROUP_rs = G_FaCn.Execute("SELECT SUBGROUP.SUBCODE,SUBGROUP.GROUPCODE AS CODE,SUBGROUP.NAME,ACGROUP.GroupName FROM SUBGROUP LEFT JOIN ACGROUP ON SUBGROUP.GROUPCODE=ACGROUP.GROUPCODE WHERE ACGROUP.ALIASYN='N' AND SUBGROUP.SITE_CODE='" & PubSiteCode & "' " & Condstr & " ORDER BY SUBGROUP.NAME")
Else
    Set SUBGROUP_rs = G_FaCn.Execute("SELECT SUBGROUP.SUBCODE,SUBGROUP.GROUPCODE AS CODE,SUBGROUP.NAME,ACGROUP.GroupName FROM SUBGROUP LEFT JOIN ACGROUP ON SUBGROUP.GROUPCODE=ACGROUP.GROUPCODE WHERE ACGROUP.ALIASYN='N' " & Condstr & " ORDER BY SUBGROUP.NAME")
End If
ac_str = ""
Set RstTmp = New ADODB.Recordset
Set RstTmp = PubDatamanFa.FaAgeTmp(RstTmp)
RstTmp.Sort = "ACC_NAME ASC"
Do Until SUBGROUP_rs.EOF
    Dr1 = 0: DR2 = 0: DR3 = 0: DR4 = 0: DR5 = 0: DR6 = 0: DR7 = 0: dr = 0: cr = 0
    Set mGROUP_rs = G_FaCn.Execute("SELECT V_DATE,DEBIT,CREDIT,PARTY AS SUBCODE FROM VIEWLEDGER WHERE '" & SUBGROUP_rs!SubCode & "'=VIEWLEDGER.PARTY AND V_DATE<=" & FaConvertDate(TXTE_DATE) & " " & mSiteCode & "")
    Do Until mGROUP_rs.EOF
        If DrCrType = "Dr" Then
            If mGROUP_rs!CREDIT > 0 Then
                cr = mGROUP_rs!CREDIT + cr
            Else
                DAYS = DateDiff("D", mGROUP_rs!V_DATE, TXTE_DATE)
                If DAYS <= AGE_rs!Age1 Then
                    Dr1 = Dr1 + mGROUP_rs!DEBIT
                ElseIf DAYS > AGE_rs!Age1 And DAYS <= AGE_rs!Age2 Then
                    DR2 = DR2 + mGROUP_rs!DEBIT
                ElseIf DAYS > AGE_rs!Age2 And DAYS <= AGE_rs!Age3 Then
                    DR3 = DR3 + mGROUP_rs!DEBIT
                ElseIf DAYS > AGE_rs!Age3 And DAYS <= AGE_rs!Age4 Then
                    DR4 = DR4 + mGROUP_rs!DEBIT
                ElseIf DAYS > AGE_rs!Age4 And DAYS <= AGE_rs!Age5 Then
                    DR5 = DR5 + mGROUP_rs!DEBIT
                ElseIf DAYS > AGE_rs!Age5 And DAYS <= AGE_rs!Age6 Then
                    DR6 = DR6 + mGROUP_rs!DEBIT
                Else
                    DR7 = DR7 + mGROUP_rs!DEBIT
                End If
            End If
        ElseIf DrCrType = "Cr" Then
            If mGROUP_rs!DEBIT > 0 Then
                cr = mGROUP_rs!DEBIT + cr
            Else
                DAYS = DateDiff("D", mGROUP_rs!V_DATE, TXTE_DATE)
                If DAYS <= AGE_rs!Age1 Then
                    Dr1 = Dr1 + mGROUP_rs!CREDIT
                ElseIf DAYS > AGE_rs!Age1 And DAYS <= AGE_rs!Age2 Then
                    DR2 = DR2 + mGROUP_rs!CREDIT
                ElseIf DAYS > AGE_rs!Age2 And DAYS <= AGE_rs!Age3 Then
                    DR3 = DR3 + mGROUP_rs!CREDIT
                ElseIf DAYS > AGE_rs!Age3 And DAYS <= AGE_rs!Age4 Then
                    DR4 = DR4 + mGROUP_rs!CREDIT
                ElseIf DAYS > AGE_rs!Age4 And DAYS <= AGE_rs!Age5 Then
                    DR5 = DR5 + mGROUP_rs!CREDIT
                ElseIf DAYS > AGE_rs!Age5 And DAYS <= AGE_rs!Age6 Then
                    DR6 = DR6 + mGROUP_rs!CREDIT
                Else
                    DR7 = DR7 + mGROUP_rs!CREDIT
                End If
            End If
        End If
        mGROUP_rs.MoveNext
    Loop
    ARR(0) = Dr1: ARR(1) = DR2: ARR(2) = DR3: ARR(3) = DR4: ARR(4) = DR5: ARR(5) = DR6: ARR(6) = DR7: X = 7
    Do While X <> 0
        If ARR(X - 1) > 0 Then
            If cr >= ARR(X - 1) Then
                cr = cr - ARR(X - 1)
                ARR(X - 1) = 0
            ElseIf ARR(X - 1) > cr Then
                ARR(X - 1) = ARR(X - 1) - cr
                cr = 0
            End If
        End If
        X = X - 1
    Loop
    X = ARR(0) + ARR(1) + ARR(2) + ARR(3) + ARR(4) + ARR(5) + ARR(6)
    If Not ((Round(X, 2) = 0) And (Round(cr, 2) = 0)) Then
        With RstTmp
            .AddNew
            !DEBIT1 = ARR(0)
            !DEBIT2 = ARR(1)
            !DEBIT3 = ARR(2)
            !DEBIT4 = ARR(3)
            !DEBIT5 = ARR(4)
            !DEBIT6 = ARR(5)
            !DEBIT = ARR(6)
            !TOTALDR = ARR(0) + ARR(1) + ARR(2) + ARR(3) + ARR(4) + ARR(5) + ARR(6)
            !CREDIT = cr
            !ACC_NAME = left(SUBGROUP_rs!Name, 35)
            !AName = left(SUBGROUP_rs!GroupName, 35)
            .Update
        End With
    End If
    SUBGROUP_rs.MoveNext
Loop
If RstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
'X1 = CreateFieldDefFile(RstTmp, PubFaReportPath + "\FaAGEING.ttx", True)
Set rpt = PubDatamanFa.FaAgeingRpt
For I = 1 To rpt.FormulaFields.Count
    Select Case rpt.FormulaFields(I).FormulaFieldName
        Case "title"
            If Trim(TEXT(SiteCode)) = "" Then
                rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
            Else
                rpt.FormulaFields(I).TEXT = "'" + Me.CAPTION + " (" + Trim(TEXT(SiteCode)) + ")" + "'"
            End If
        Case "ACNAME"
            rpt.FormulaFields(I).TEXT = "'For A/C  : " & TXTACC_CODE.TEXT & "'"
        Case "DATE"
            rpt.FormulaFields(I).TEXT = "'Upto Date : " & TXTE_DATE & "'"
        Case "P1"
            rpt.FormulaFields(I).TEXT = "'" & "0 - " + STR(AGE_rs!Age1) & "'"
        Case "P2"
            rpt.FormulaFields(I).TEXT = "'" & STR(AGE_rs!Age1 + 1) & " - " & STR(AGE_rs!Age2) & "'"
        Case "P3"
            rpt.FormulaFields(I).TEXT = "'" & STR(AGE_rs!Age2 + 1) + " - " + STR(AGE_rs!Age3) & "'"
        Case "P4"
            rpt.FormulaFields(I).TEXT = "'" & STR(AGE_rs!Age3 + 1) + " - " + STR(AGE_rs!Age4) & "'"
        Case "P5"
            rpt.FormulaFields(I).TEXT = "'" & STR(AGE_rs!Age4 + 1) + " - " + STR(AGE_rs!Age5) & "'"
        Case "P6"
            rpt.FormulaFields(I).TEXT = "'" & STR(AGE_rs!Age5 + 1) + " - " + STR(AGE_rs!Age6) & "'"
        Case "P7"
            rpt.FormulaFields(I).TEXT = "'Above " & STR(AGE_rs!Age6) & "'"
        Case "P8"
            If DrCrType = "Dr" Then
                rpt.FormulaFields(I).TEXT = "'Total Debit'"
            ElseIf DrCrType = "Cr" Then
                rpt.FormulaFields(I).TEXT = "'Total Credit'"
            End If
        Case "P9"
            If DrCrType = "Dr" Then
                rpt.FormulaFields(I).TEXT = "'Total Credit'"
            ElseIf DrCrType = "Cr" Then
                rpt.FormulaFields(I).TEXT = "'Total Debit'"
            End If
        Case "HEADI"
            If DrCrType = "Dr" Then
                rpt.FormulaFields(I).TEXT = "'<----------------------------- AMOUNT DEBITED FROM DAYS ------------------------------>'"
            ElseIf DrCrType = "Cr" Then
                rpt.FormulaFields(I).TEXT = "'<----------------------------- AMOUNT CREDITED FROM DAYS ------------------------------>'"
            End If
        Case "PageNo"
            rpt.FormulaFields(I).TEXT = "'" & RstEnviro!pagenofill & "'"
        Case "DT"
             rpt.FormulaFields(I).TEXT = "'" & RstEnviro!daterfill & "'"
    End Select
Next
rpt.Database.SetDataSource RstTmp
EXIT_SUB:
    Set AGE_rs = Nothing
    Set SUBGROUP_rs = Nothing
    Set RstTmp = Nothing
    Set mGROUP_rs = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub BankReg(Index As Integer)
On Error GoTo ELoop
Dim Rst1 As ADODB.Recordset, TinTin As Integer, Condstr As String, mSiteCode As String

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
GridString1 = ""
Condstr = ""
If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
If GridString1 <> "" Then Condstr = " AND VIEWLEDGER.PARTY IN (" & GridString1 & ")"
TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub

If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And VIEWLEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
End If

Set Rst1 = G_FaCn.Execute("SELECT MAX(VIEWLEDGER.PARTY) AS PARTY_cODE,MAX(SUBGROUP.NAME) AS PARTY_NAME, " & FaConvertDate(TXTS_DATE) & " AS V_DATE,Sum(DEBIT) AS DR,Sum(CREDIT) AS CR,'' AS v_type, 0 AS v_no,'' AS v_add,'' AS CHQ_NO,'' AS CHQ_DATE,'' AS CLG_DATE,'' AS NARRATION,'Opening Balance' AS Name1 FROM VIEWLEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE = VIEWLEDGER.PARTY WHERE V_DATE<" & FaConvertDate(TXTS_DATE) & " " & Condstr & " AND SUBGROUP.NATURE='Bank' " & mSiteCode & " GROUP BY PARTY " & _
"Union SELECT VIEWLEDGER.PARTY AS PARTY_CODE,SUBGROUP.NAME AS PARTY_NAME,V_DATE,DEBIT AS DR,CREDIT AS CR,v_type,v_no,v_add,CHQ_NO,CHQ_DATE,CLG_DATE,NARR AS NARRATION,SUBGROUP1.NAME AS NAME1 FROM (VIEWLEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE = VIEWLEDGER.PARTY) LEFT JOIN SUBGROUP SUBGROUP1 ON VIEWLEDGER.party1=SUBGROUP1.SUBCODE WHERE V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " " & Condstr & " AND SUBGROUP.NATURE='Bank' " & mSiteCode & "")
If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
'X1 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaBANKREG.ttx", True)
Set rpt = PubDatamanFa.FaBankregRpt
rpt.Database.SetDataSource Rst1
EXIT_SUB:
    Set Rst1 = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub NonTrans(Index As Integer)
On Error GoTo ELoop
Dim Rst1 As ADODB.Recordset, RstTmp As ADODB.Recordset, MyRst As ADODB.Recordset
Dim TinTin As Integer, MyDrStr As String, MyCrStr As String, Condstr As String, MyOpBal As Double, MyCloBal As Double, mSiteCode As String

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
GridString1 = ""
Condstr = ""
If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
If GridString1 <> "" Then Condstr = " AND SUBGROUP.GROUPCODE IN (" & GridString1 & ") AND ACGROUP.GROUPCODE IN (" & GridString1 & ")"
TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub
If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(LEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And LEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
End If
Set Rst1 = G_FaCn.Execute("Select SUBGROUP.GroupCode,Name,SubCode,MAX(GroupName) AS GNAME From SubGroup LEFT JOIN AcGroup on AcGroup.GROUPCODE=SUBGROUP.GROUPCODE Where SubCode Not In (Select DISTINCT SubCode From Ledger Where V_Date Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & ") " & Condstr & " GROUP BY SubGroup.GroupCode,SubGroup.Name,SubGroup.SubCode")
If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
Set RstTmp = New ADODB.Recordset
Set RstTmp = PubDatamanFa.FaNonTrans(RstTmp)
While Not Rst1.EOF
    MyDrStr = ""
    MyCrStr = ""
    If PubBackEnd = "S" Then
        Set MyRst = G_FaCn.Execute("Select ISNULL(sum(AmtDr),0)-ISNULL(sum(AmtCr),0) As OpBal From Ledger Where SubCode='" & Rst1!SubCode & "' And V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "")
        If MyRst.RecordCount > 0 Then MyOpBal = IIf(IsNull(MyRst!OpBal), 0, MyRst!OpBal)
        Set MyRst = G_FaCn.Execute("SELECT ISNULL(sum(AmtDr),0)-ISNULL(sum(AmtCr),0) AS ClBal FROM Ledger Where SubCode='" & Rst1!SubCode & "' And V_Date<=" & FaConvertDate(TXTE_DATE) & " " & mSiteCode & "")
        If MyRst.RecordCount > 0 Then MyCloBal = IIf(IsNull(MyRst!clbal), 0, MyRst!clbal)
    ElseIf PubBackEnd = "A" Then
        Set MyRst = G_FaCn.Execute("Select sum(IIF(ISNULL(AmtDr),0,AmtDr))-sum(IIF(ISNULL(AmtCr),0,AmtCr)) As OpBal From Ledger Where SubCode='" & Rst1!SubCode & "' And V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "")
        If MyRst.RecordCount > 0 Then MyOpBal = IIf(IsNull(MyRst!OpBal), 0, MyRst!OpBal)
        Set MyRst = G_FaCn.Execute("SELECT sum(IIF(ISNULL(AmtDr),0,AmtDr))-sum(IIF(ISNULL(AmtCr),0,AmtCr)) AS ClBal FROM Ledger Where SubCode='" & Rst1!SubCode & "' And V_Date<=" & FaConvertDate(TXTE_DATE) & " " & mSiteCode & "")
        If MyRst.RecordCount > 0 Then MyCloBal = IIf(IsNull(MyRst!clbal), 0, MyRst!clbal)
    End If
    Set MyRst = G_FaCn.Execute("SELECT subgroup.Name,V_DATE,AMTCR,AMTDR,v_type,v_no FROM LEDGER LEFT JOIN subgroup ON LEDGER.CONTRASUB=subgroup.SubCode WHERE LEDGER.SUBCODE='" & Rst1!SubCode & "' AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND AMTDR>0 " & mSiteCode & " ORDER BY LEDGER.V_DATE DESC")
    If MyRst.RecordCount > 0 Then
        MyDrStr = CStr(MyRst!V_DATE) + Space(1) + CStr(MyRst!V_tYPE) + Space(1) + CStr(MyRst!V_NO) + Space(1) + CStr(MyRst!AmtDr)
    End If
    Set MyRst = G_FaCn.Execute("SELECT subgroup.Name,V_DATE,AMTcr,AMTDR,v_type,v_no FROM LEDGER LEFT JOIN subgroup ON LEDGER.CONTRASUB=subgroup.SubCode WHERE LEDGER.SUBCODE='" & Rst1!SubCode & "' AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND AMTCR>0 " & mSiteCode & " ORDER BY LEDGER.V_DATE DESC")
    If MyRst.RecordCount > 0 Then
        MyCrStr = CStr(MyRst!V_DATE) + Space(1) + CStr(MyRst!V_tYPE) + Space(1) + CStr(MyRst!V_NO) + Space(1) + CStr(MyRst!AmtCr)
    End If
    With RstTmp
        .AddNew
        !ACCNAME = Rst1!Name
        !OpBal = Format(MyOpBal, "0.00")
        !clbal = Format(MyCloBal, "0.00")
        !LastDrVNo = MyDrStr
        !LastCrVNo = MyCrStr
        !GroupName = Rst1!GName
        .Update
    End With
    Rst1.MoveNext
Wend
'X11 = CreateFieldDefFile(RstTmp, PubFaReportPath + "\FaNonTran.ttx", True)
If RstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
Set rpt = PubDatamanFa.FaNonTranRpt
rpt.Database.SetDataSource RstTmp
EXIT_SUB:
    Set Rst1 = Nothing
    Set RstTmp = Nothing
    Set MyRst = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub Clg(Index As Integer)
On Error GoTo ELoop
Dim Rst1 As ADODB.Recordset, mAcCode As String, TinTin As Integer, mQRY1 As String, mSiteCode As String

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub
If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And VIEWLEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
End If

TXTACC_CODE = Trim(FGrid.TextMatrix(List1, 1))
mAcCode = Trim(FGrid.TextMatrix(List1, 2))
mQRY1 = " WHERE PARTY='" & mAcCode & "' AND CLG_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND CLG_DATE IS NOT NULL"
Set Rst1 = G_FaCn.Execute("SELECT 1 AS TT," & FaConvertDate(TXTS_DATE) & " AS V_DATE,SUM(credit) AS CR,SUM(debit) AS DR,'OP' AS v_type,0 AS v_no,'' AS v_add,'' AS CHQ_NO,NULL AS CHQ_DATE," & FaConvertDate(TXTS_DATE) & " AS CLG_DATE,'' AS NARRATION,MAX(PARTY_NAME) AS PARTYNAME,MAX(PARTY_LIST.ADD1) AS ADDRESS1,MAX(PARTY_LIST.ADD2) AS ADDRESS2,MAX(CITY_NAME) AS CITYNAME,'Opening Balance' AS CONTRA_NAME FROM (VIEWLEDGER LEFT JOIN SUBGROUP ON VIEWLEDGER.party1=SUBGROUP.SUBCODE) LEFT JOIN PARTY_LIST ON VIEWLEDGER.party=PARTY_LIST.SUBCODE WHERE PARTY='" & mAcCode & "' AND CLG_DATE<" & FaConvertDate(TXTS_DATE) & " AND CLG_DATE IS NOT NULL " & mSiteCode & " GROUP BY VIEWLEDGER.PARTY " & _
        "UNION SELECT 2 AS TT,V_DATE,credit AS CR,debit AS DR,v_type,v_no,v_add,CHQ_NO,CHQ_DATE,CLG_DATE,MNARR+' '+NARR AS NARRATION,PARTY_NAME AS PARTYNAME,PARTY_LIST.ADD1 AS ADDRESS1,PARTY_LIST.ADD2 AS ADDRESS2,CITY_NAME AS CITYNAME,SUBGROUP.NAME AS CONTRA_NAME FROM (VIEWLEDGER LEFT JOIN SUBGROUP ON VIEWLEDGER.party1=SUBGROUP.SUBCODE) LEFT JOIN PARTY_LIST ON VIEWLEDGER.party=PARTY_LIST.SUBCODE " & mQRY1 & "  " & mSiteCode & "")
If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
'X1 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaCLGCLEAR.ttx", True)
Set rpt = PubDatamanFa.FaClgClearRpt
rpt.Database.SetDataSource Rst1
EXIT_SUB:
    Set Rst1 = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub ClgNot(Index As Integer)
On Error GoTo ELoop
Dim Rst1 As ADODB.Recordset, mAcCode As String, TinTin As Integer, mQRY1 As String, mSiteCode As String

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub
If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And VIEWLEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
End If
TXTACC_CODE = Trim(FGrid.TextMatrix(List1, 1))
mAcCode = Trim(FGrid.TextMatrix(List1, 2))
mQRY1 = " WHERE PARTY='" & mAcCode & "' AND V_date BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND CHQ_NO<> '' AND CHQ_NO IS NOT NULL AND CLG_DATE IS NULL"
Set Rst1 = G_FaCn.Execute("SELECT V_DATE,credit,debit,v_type,v_no,v_add,CHQ_NO,CHQ_DATE,CLG_DATE,NARR AS NARRATION,PARTY_NAME,PARTY_LIST.ADD1,PARTY_LIST.ADD2,CITY_NAME,SUBGROUP.NAME AS CONTRA_NAME FROM (VIEWLEDGER LEFT JOIN SUBGROUP ON VIEWLEDGER.party1=SUBGROUP.SUBCODE) LEFT JOIN PARTY_LIST ON VIEWLEDGER.party=PARTY_LIST.SUBCODE " & mQRY1 & " " & mSiteCode)
If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
'X1 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaCLG.ttx", True)
Set rpt = PubDatamanFa.FaClgRpt
rpt.Database.SetDataSource Rst1
EXIT_SUB:
    Set Rst1 = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub DailySumm(Index As Integer)
On Error GoTo ELoop
Dim Rst1 As ADODB.Recordset, G_Rs As ADODB.Recordset
Dim mAcCode As String, mDocNo As String, mDocNo1 As String, TinTin As Integer, mSiteCode As String

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub
If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And VIEWLEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
End If
TXTACC_CODE = Trim(FGrid.TextMatrix(List1, 1))
mAcCode = Trim(FGrid.TextMatrix(List1, 2))
Set Rst1 = G_FaCn.Execute("SELECT GROUPNATURE FROM PARTY_LIST WHERE SUBCODE='" & mAcCode & "'")
If Rst1.RecordCount > 0 Then
    If Rst1!GroupNature = "A" Or Rst1!GroupNature = "L" Then
        Set G_Rs = G_FaCn.Execute("SELECT SUM(CREDIT)-SUM(DEBIT) AS OPBAL from VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "")
    Else
        Set G_Rs = G_FaCn.Execute("SELECT SUM(CREDIT)-SUM(DEBIT) AS OPBAL from VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_dATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "")
    End If
End If
If G_Rs.RecordCount > 0 Then
    TOT_AMTDR = FaVNull(G_Rs!OpBal)
Else
    TOT_AMTDR = 0
End If
Select Case Index
    Case 1
        If PubBackEnd = "A" Then
            Set Rst1 = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) AS DEB,IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) AS CRED,V_DATE from VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " " & mSiteCode & " GROUP BY V_DATE ORDER BY V_DATE")
        ElseIf PubBackEnd = "S" Then
            Set Rst1 = G_FaCn.Execute("SELECT ISNULL(SUM(DEBIT),0) AS DEB,ISNULL(SUM(CREDIT),0) AS CRED,V_DATE from VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " " & mSiteCode & " GROUP BY V_DATE ORDER BY V_DATE")
        End If
        If Rst1.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: RepPrint = False: GoTo EXIT_SUB
        MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
        TinTin = PubDatamanFa.FaDailyTransSummaryDosPrinting(Me, Rst1, TOT_AMTDR)
        GoTo EXIT_SUB
    Case Else
        Set Rst1 = G_FaCn.Execute("SELECT V_TYPE,V_NO,V_ADD,V_SNO,PARTY,0 AS OPBAL,DEBIT AS DEB,CREDIT AS CRED,V_DATE,GROUPNATURE from VIEWLEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=VIEWLEDGER.PARTY WHERE PARTY='" & mAcCode & "' AND V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " " & mSiteCode & "")
        If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: GoTo EXIT_SUB
'        X11 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaDAILYSUM.ttx", True)
        Set rpt = PubDatamanFa.FaDailysumRpt
        rpt.Database.SetDataSource Rst1
End Select
EXIT_SUB:
    Set Rst1 = Nothing
    Set G_Rs = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub JournalBook(Index As Integer)
'On Error GoTo ELoop
Dim Rst1 As ADODB.Recordset, TmpRst As ADODB.Recordset
Dim mAcCode As String, mDocNo As String, mDocNo1 As String, TinTin As Integer, mSiteCode As String
If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub
TXTACC_CODE = Trim(FGrid.TextMatrix(List1, 1))
mAcCode = Trim(FGrid.TextMatrix(List1, 2))
If Trim(FGrid.TextMatrix(List3, 2)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & FGrid.TextMatrix(List3, 2) & "'"
    ElseIf PubFaSiteType = 2 Then
        If PubSiteCodeWidth = 1 Then
            mSiteCode = " And VIEWLEDGER.Site_Code='" & FGrid.TextMatrix(List3, 2) & "'"
        Else
            mSiteCode = " And VIEWLEDGER.Site_Code='" & Trim(FGrid.TextMatrix(List3, 2)) + Trim(FGrid.TextMatrix(List3, 2)) & "'"
        End If
    End If
Else
    mSiteCode = ""
End If
Set Rst1 = G_FaCn.Execute("SELECT V_DATE,credit,debit,v_type,v_no,v_add,CHQ_NO,CHQ_DATE,NARR,NAME,MNARR,V_SNO,DOCID,VIEWLEDGER.SITE_CODE FROM VIEWLEDGER LEFT JOIN SUBGROUP ON VIEWLEDGER.party=SUBGROUP.SUBCODE where V_date BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND V_TYPE='" & mAcCode & "' " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,V_ADD,V_SNO")
If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
Set TmpRst = New ADODB.Recordset
Set TmpRst = PubDatamanFa.FaJournal(TmpRst)
Do Until Rst1.EOF
    With TmpRst
        If PubFaSiteType = 1 Then
            mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(Rst1!DocId, 1), "")
            mDocNo = mDocNo + IIf(RstEnviro!LedSiteCode = "Yes", Trim(Right(Rst1!Site_Code, 1)), "")
            mDocNo = mDocNo + IIf(RstEnviro!LedPrefix = "Yes", IIf(mDocNo = "", "", "/") + Trim(Rst1!V_ADD), "")
            mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + left(Trim(Rst1!V_tYPE), 1) + Trim(mID(Trim(Rst1!V_tYPE), 3, 3))
            mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(Rst1!V_NO))
        Else
            mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(Rst1!DocId, 1), "") + IIf(RstEnviro!LedSiteCode = "Yes", Trim(left(Rst1!Site_Code, 1)), "")
            mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + IIf(RstEnviro!LedPrefix = "Yes", Trim(Rst1!V_ADD), "")
            mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(Rst1!V_tYPE)
            mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(Rst1!V_NO))
        End If
        .AddNew
        !V_ADD = FaXNull(Rst1!V_ADD)
        !V_NO = Rst1!V_NO
        !DOCNO = mDocNo
        !Name = Trim(FaXNull(Rst1!Name))
        !V_DATE = Format(Rst1!V_DATE, "dd/MMM/yyyy")
        !CREDIT = Rst1!CREDIT
        !DEBIT = Rst1!DEBIT
        !V_tYPE = Rst1!V_tYPE
        !V_SNo = FaVNull(Rst1!V_SNo)
        !mNarr = Rst1!mNarr
        !Narr = Rst1!Narr
        If Trim(FaXNull(Rst1!Chq_No)) <> "" Then
            !Chq_No = Rst1!Chq_No
        End If
        If Not IsNull(Rst1!Chq_Date) Then
            !Chq_Date = Rst1!Chq_Date
        End If
    End With
    Rst1.MoveNext
Loop
If TmpRst.RecordCount > 0 Then TmpRst.MoveFirst
Select Case Index
    Case 1
        MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
        TinTin = PubDatamanFa.FaJBookDosPrinting(Me, TmpRst, FGrid.TextMatrix(List2, 1))
'        JBookDos Me, TmpRst, FGrid.TextMatrix(List2, 1)
        GoTo EXIT_SUB
    Case Else
    Dim X11
        X11 = CreateFieldDefFile(TmpRst, PubFaReportPath + "\FaJRNL.ttx", True)
        Set rpt = PubDatamanFa.FaJRNLRpt
        rpt.Database.SetDataSource TmpRst
End Select
EXIT_SUB:
    Set Rst1 = Nothing
    Set TmpRst = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub CashBook(Index As Integer)
'On Error GoTo ELoop
Dim Rst1 As ADODB.Recordset, RstTmp As ADODB.Recordset, mGROUP_rs As ADODB.Recordset, SUBGROUP_rs As ADODB.Recordset, TmpGrs As ADODB.Recordset, TmpGrs1 As ADODB.Recordset
Dim DrAc As String, CrAc As String, oBAL As Double, mAcCode As String, mDocNo As String, mDocNo1 As String
Dim mNARR1 As String, mNARR2 As String, TmpDate As Date, mDate1 As Date, mDate2 As Date, TinTin As Integer
Dim mFLAG1 As Boolean, mFLAG2 As Boolean, mFLAG11 As Boolean, mFLAG22 As Boolean, mFLAG111 As Boolean, mFLAG222 As Boolean, mFLAG1111 As Boolean, mFLAG2222 As Boolean
Dim DNarrStr As String, DmNarrStr1 As String, DmNarrStr2 As String
Dim D1NarrStr As String, D1mNarrStr1 As String, D1mNarrStr2 As String, mSiteCode As String

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List3, FGrid.TextMatrix(List3, 0)) = False Then RepPrint = False: Exit Sub
TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub

If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And VIEWLEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
End If
TXTACC_CODE = Trim(FGrid.TextMatrix(List3, 1))
mAcCode = Trim(FGrid.TextMatrix(List3, 2))
DrAc = ""
CrAc = ""
oBAL = 0
Set Rst1 = G_FaCn.Execute("SELECT GROUPNATURE FROM PARTY_LIST WHERE SUBCODE='" & mAcCode & "'")
If Rst1.RecordCount <= 0 Then Exit Sub
If Rst1!GroupNature = "A" Or Rst1!GroupNature = "L" Then
    If PubBackEnd = "S" Then
        oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode).Fields(0)
    ElseIf PubBackEnd = "A" Then
        oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE<" & FaConvertDate(TXTS_DATE) & "  " & mSiteCode).Fields(0)
    End If
Else
    If PubBackEnd = "S" Then
        oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode).Fields(0)
    ElseIf PubBackEnd = "A" Then
        oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode).Fields(0)
    End If
End If
Set RstTmp = New ADODB.Recordset
Set RstTmp = PubDatamanFa.FaCASHTMP1(RstTmp)
If oBAL <> 0 Then
    RstTmp.AddNew
    RstTmp!V_DATE = TXTS_DATE
    If oBAL < 0 Then
        RstTmp!Name = "OPENING BALANCE"
        RstTmp!cr = Abs(oBAL)
    Else
        RstTmp!Name1 = "OPENING BALANCE"
        RstTmp!ADJAMT = Abs(oBAL)
    End If
    RstTmp.Update
End If
If PubBackEnd = "S" Then
    Set mGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.V_ADD,DocID,VIEWLEDGER.Site_Code,VIEWLEDGER.V_NO,subgroup.NAME, VIEWLEDGER.V_DATE, VIEWLEDGER.CREDIT AS AMOUNT, VIEWLEDGER.V_TYPE,VIEWLEDGER.MNARR,VIEWLEDGER.NARR, VIEWLEDGER.V_SNO,VIEWLEDGER.CHQ_NO,CONVERT(VARCHAR,VIEWLEDGER.CHQ_DATE,103)AS CHQDATE FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY1=subgroup.SUBCODE WHERE VIEWLEDGER.PARTY='" & mAcCode & "' AND VIEWLEDGER.V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND CREDIT>0  " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
    Set SUBGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.V_ADD,DocID,VIEWLEDGER.Site_Code,VIEWLEDGER.V_NO,subgroup.NAME, VIEWLEDGER.V_DATE, VIEWLEDGER.DEBIT AS AMOUNT, VIEWLEDGER.V_TYPE,VIEWLEDGER.MNARR,VIEWLEDGER.NARR, VIEWLEDGER.V_SNO,VIEWLEDGER.CHQ_NO,CONVERT(VARCHAR,VIEWLEDGER.CHQ_DATE,103) AS CHQDATE FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY1=subgroup.SUBCODE WHERE VIEWLEDGER.PARTY='" & mAcCode & "' AND VIEWLEDGER.V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND DEBIT>0  " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
ElseIf PubBackEnd = "A" Then
    Set mGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.V_ADD,DocID,VIEWLEDGER.Site_Code,VIEWLEDGER.V_NO,subgroup.NAME, VIEWLEDGER.V_DATE, VIEWLEDGER.CREDIT AS AMOUNT, VIEWLEDGER.V_TYPE,VIEWLEDGER.MNARR,VIEWLEDGER.NARR, VIEWLEDGER.V_SNO,VIEWLEDGER.CHQ_NO,FORMAT(VIEWLEDGER.CHQ_DATE,'DD/MM/YY') AS CHQDATE FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY1=subgroup.SUBCODE WHERE VIEWLEDGER.PARTY='" & mAcCode & "' AND VIEWLEDGER.V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND CREDIT>0   " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
    Set SUBGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.V_ADD,DocID,VIEWLEDGER.Site_Code,VIEWLEDGER.V_NO,subgroup.NAME, VIEWLEDGER.V_DATE, VIEWLEDGER.DEBIT AS AMOUNT, VIEWLEDGER.V_TYPE,VIEWLEDGER.MNARR,VIEWLEDGER.NARR, VIEWLEDGER.V_SNO,VIEWLEDGER.CHQ_NO,FORMAT(VIEWLEDGER.CHQ_DATE,'DD/MM/YY') AS CHQDATE FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY1=subgroup.SUBCODE WHERE VIEWLEDGER.PARTY='" & mAcCode & "' AND VIEWLEDGER.V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND DEBIT>0   " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
End If
If Not (mGROUP_rs.EOF) Then mDate2 = mGROUP_rs!V_DATE
If Not (SUBGROUP_rs.EOF) Then mDate1 = SUBGROUP_rs!V_DATE
mFLAG1 = False
mFLAG2 = False
mFLAG11 = False
mFLAG22 = False
mFLAG111 = False
mFLAG222 = False
mFLAG1111 = False
mFLAG2222 = False
mNARR1 = ""
mNARR2 = ""
TmpDate = TXTS_DATE
Do Until mGROUP_rs.EOF And SUBGROUP_rs.EOF
    If mDate1 = TmpDate Or mDate2 = TmpDate Then
        RstTmp.AddNew
        If mDate1 = TmpDate Then
            RstTmp!V_tYPE = SUBGROUP_rs!V_tYPE
            RstTmp!V_DATE = mDate1
            RstTmp!V_NO = SUBGROUP_rs!V_NO
            RstTmp!V_SNo = SUBGROUP_rs!V_SNo
            If PubFaSiteType = 1 Then
                mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(SUBGROUP_rs!DocId, 1), "")
                mDocNo = mDocNo + IIf(RstEnviro!LedSiteCode = "Yes", Trim(Right(SUBGROUP_rs!Site_Code, 1)), "")
                mDocNo = mDocNo + IIf(RstEnviro!LedPrefix = "Yes", IIf(mDocNo = "", "", "/") + Trim(SUBGROUP_rs!V_ADD), "")
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + left(Trim(SUBGROUP_rs!V_tYPE), 1) + Trim(mID(Trim(SUBGROUP_rs!V_tYPE), 3, 3))
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(SUBGROUP_rs!V_NO))
            Else
                mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(SUBGROUP_rs!DocId, 1), "") + IIf(RstEnviro!LedSiteCode = "Yes", Trim(left(SUBGROUP_rs!Site_Code, 1)), "")
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + IIf(RstEnviro!LedPrefix = "Yes", Trim(SUBGROUP_rs!V_ADD), "")
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(SUBGROUP_rs!V_tYPE)
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(SUBGROUP_rs!V_NO))
            End If
            RstTmp!DOCNO = mDocNo
            If mFLAG1 = False Or mFLAG11 = False Or mFLAG111 = False Then
                mFLAG1 = True
                mNARR1 = ""
                If FaXNull(Trim(SUBGROUP_rs!Chq_No)) <> "" Then mNARR1 = mNARR1 + "Ch.No:" + Trim(FaXNull(SUBGROUP_rs!Chq_No)) + " Ch.Dt: " + CStr(FaXNull(SUBGROUP_rs!ChqDate))
                mNARR1 = mNARR1 + Trim(FaXNull(SUBGROUP_rs!mNarr)) + Trim(FaXNull(SUBGROUP_rs!Narr))
                If Len(FaXNull(SUBGROUP_rs!Name)) <> 0 Then
                    RstTmp!Name = FaXNull(SUBGROUP_rs!Name)
                    RstTmp!cr = Format(SUBGROUP_rs!AMOUNT, "0.00")
                    oBAL = oBAL - Format(SUBGROUP_rs!AMOUNT, "0.00")
                    mFLAG1111 = True
                    mFLAG111 = True
                    mFLAG11 = True
                Else
                    If mFLAG11 = False And Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then
                        If mFLAG111 = False Then
                            RstTmp!cr = Format(SUBGROUP_rs!AMOUNT, "0.00")
                            oBAL = oBAL - Format(SUBGROUP_rs!AMOUNT, "0.00")
                            RstTmp!Name = "As Per Detail"
                            mFLAG111 = True
                            mFLAG1111 = False
                            If Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then Set TmpGrs = G_FaCn.Execute("SELECT subgroup.NAME,VIEWLEDGER.* FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE VIEWLEDGER.DOCID='" & SUBGROUP_rs!DocId & "' AND V_SNO<>" & SUBGROUP_rs!V_SNo)
                        Else
                            If mFLAG1111 = False Then
                                If Not TmpGrs.EOF Then
                                    If TmpGrs!DEBIT > 0 Then
                                        RstTmp!Name = Space(2) + FaSetW(TmpGrs!Name, 20) + " " + FaSetN(FaSNull(TmpGrs!DEBIT), 12) + " Dr"
                                    Else
                                        RstTmp!Name = Space(2) + FaSetW(TmpGrs!Name, 20) + " " + FaSetN(FaSNull(TmpGrs!CREDIT), 12) + " Cr"
                                    End If
                                    DNarrStr = ""
                                    DmNarrStr1 = ""
                                    DmNarrStr2 = ""
                                    If FGrid.TextMatrix(Cat1, 1) = "Yes" Then
                                        If FaXNull(TmpGrs!Chq_No) <> "" Then DNarrStr = DNarrStr + "Chq.No:" + Trim(FaXNull(TmpGrs!Chq_No))
                                        If Not IsNull(TmpGrs!Chq_Date) Then DNarrStr = DNarrStr + " Dt: " + CStr(Format(TmpGrs!Chq_Date, "dd/MM/yy"))
                                        DmNarrStr2 = Trim(FaXNull(TmpGrs!Narr))
                                        If DNarrStr <> "" Or DmNarrStr1 <> "" Or DmNarrStr2 <> "" Then
                                            mFLAG1111 = True
                                        Else
                                            mFLAG1111 = False
                                        End If
                                    Else
                                        mFLAG1111 = False
                                    End If
                                    TmpGrs.MoveNext
                                End If
                            Else
                                If FGrid.TextMatrix(Cat1, 1) = "Yes" Then
                                    If Trim(DNarrStr) <> "" Then
                                        RstTmp!Name = Space(2) + FaSetW(DNarrStr, 35)
                                        DNarrStr = mID(DNarrStr, 36, 300)
                                    End If
                                    If Trim(DmNarrStr2) <> "" Then
                                        RstTmp!Name = Space(2) + FaSetW(DmNarrStr2, 35)
                                        DmNarrStr2 = mID(DmNarrStr2, 36, 300)
                                    End If
                                End If
                            End If
                            If DNarrStr = "" And DmNarrStr1 = "" And DmNarrStr2 = "" Then
                                mFLAG1111 = False
                            End If
                            If TmpGrs.EOF = True Then
                                If DNarrStr = "" And DmNarrStr1 = "" And DmNarrStr2 = "" Then
                                    mFLAG11 = True
                                    mFLAG111 = True
                                End If
                            End If
                        End If
                    Else
                        RstTmp!Name = Space(2) + Trim(mID(mNARR1, 1, 36))
                        RstTmp!cr = Format(SUBGROUP_rs!AMOUNT, "0.00")
                        oBAL = oBAL - Format(SUBGROUP_rs!AMOUNT, "0.00")
                        mNARR1 = Trim(mID(mNARR1, 37, 510))
                        mFLAG1111 = True
                        mFLAG111 = True
                        mFLAG11 = True
                    End If
                End If
                If Len(mNARR1) <= 0 And mFLAG11 = True And mFLAG111 = True Then
                    mFLAG1 = False
                    mFLAG11 = False
                    mFLAG111 = False
                    mFLAG1111 = False
                    SUBGROUP_rs.MoveNext
                    If Not SUBGROUP_rs.EOF Then
                        mDate1 = SUBGROUP_rs!V_DATE
                    Else
                        mDate1 = DateAdd("D", 1, TXTE_DATE)
                    End If
                End If
            Else
                mNARR1 = Trim(mNARR1)
                RstTmp!Name = Space(2) + Trim(mID(mNARR1, 1, 36))
                mNARR1 = Trim(mID(mNARR1, 37, 510))
                If Len(mNARR1) <= 0 Then
                    mFLAG1 = False
                    mFLAG11 = False
                    mFLAG111 = False
                    mFLAG1111 = False
                    SUBGROUP_rs.MoveNext
                    If Not SUBGROUP_rs.EOF Then
                        mDate1 = SUBGROUP_rs!V_DATE
                    Else
                        mDate1 = DateAdd("D", 1, TXTE_DATE)
                    End If
                End If
            End If
        End If
        If mDate2 = TmpDate Then
            RstTmp!VType = mGROUP_rs!V_tYPE
            RstTmp!VNo = mGROUP_rs!V_NO
            RstTmp!V_DATE = mDate2
            RstTmp!VSNo = mGROUP_rs!V_SNo
            If PubFaSiteType = 1 Then
                mDocNo1 = IIf(RstEnviro!LedDivCode = "Yes", left(mGROUP_rs!DocId, 1), "")
                mDocNo1 = mDocNo1 + IIf(RstEnviro!LedSiteCode = "Yes", Trim(Right(mGROUP_rs!Site_Code, 1)), "")
                mDocNo1 = mDocNo1 + IIf(RstEnviro!LedPrefix = "Yes", IIf(mDocNo1 = "", "", "/") + Trim(mGROUP_rs!V_ADD), "")
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + left(Trim(mGROUP_rs!V_tYPE), 1) + Trim(mID(Trim(mGROUP_rs!V_tYPE), 3, 3))
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + Trim(STR(mGROUP_rs!V_NO))
            Else
                mDocNo1 = IIf(RstEnviro!LedDivCode = "Yes", left(mGROUP_rs!DocId, 1), "") + IIf(RstEnviro!LedSiteCode = "Yes", Trim(left(mGROUP_rs!Site_Code, 1)), "")
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + IIf(RstEnviro!LedPrefix = "Yes", Trim(mGROUP_rs!V_ADD), "")
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + Trim(mGROUP_rs!V_tYPE)
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + Trim(STR(mGROUP_rs!V_NO))
            End If
            RstTmp!DocNo1 = mDocNo1
            If mFLAG2 = False Or mFLAG22 = False Or mFLAG222 = False Then
                mFLAG2 = True
                mNARR2 = ""
                If FaXNull(Trim(mGROUP_rs!Chq_No)) <> "" Then mNARR2 = mNARR2 + "Ch.No:" + Trim(FaXNull(mGROUP_rs!Chq_No)) + " Ch.Dt: " + CStr(FaXNull(mGROUP_rs!ChqDate))
                mNARR2 = mNARR2 + Trim(FaXNull(mGROUP_rs!mNarr)) + Trim(FaXNull(mGROUP_rs!Narr))
                If Len(FaXNull(mGROUP_rs!Name)) <> 0 Then
                    RstTmp!Name1 = mGROUP_rs!Name
                    RstTmp!ADJAMT = Format(mGROUP_rs!AMOUNT, "0.00")
                    oBAL = oBAL + Format(mGROUP_rs!AMOUNT, "0.00")
                    mFLAG2222 = True
                    mFLAG222 = True
                    mFLAG22 = True
                Else
                    If mFLAG22 = False And Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then
                        If mFLAG222 = False Then
                            RstTmp!ADJAMT = Format(mGROUP_rs!AMOUNT, "0.00")
                            oBAL = oBAL + Format(mGROUP_rs!AMOUNT, "0.00")
                            RstTmp!Name1 = "As Per Detail"
                            mFLAG222 = True
                            mFLAG2222 = False
                            If Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then Set TmpGrs1 = G_FaCn.Execute("SELECT subgroup.NAME,VIEWLEDGER.* FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE VIEWLEDGER.DOCID='" & mGROUP_rs!DocId & "' AND V_SNO<>" & mGROUP_rs!V_SNo)
                        Else
                            If mFLAG2222 = False Then
                                If Not TmpGrs1.EOF Then
                                    If TmpGrs1!DEBIT > 0 Then
                                        RstTmp!Name1 = Space(2) + FaSetW(TmpGrs1!Name, 20) + " " + FaSetN(FaSNull(TmpGrs1!DEBIT), 12) + " Dr"
                                    Else
                                        RstTmp!Name1 = Space(2) + FaSetW(TmpGrs1!Name, 20) + " " + FaSetN(FaSNull(TmpGrs1!CREDIT), 12) + " Cr"
                                    End If
                                    D1NarrStr = ""
                                    D1mNarrStr1 = ""
                                    D1mNarrStr2 = ""
                                    If FGrid.TextMatrix(Cat1, 1) = "Yes" Then
                                        If FaXNull(TmpGrs1!Chq_No) <> "" Then DNarrStr = DNarrStr + "Chq.No:" + Trim(FaXNull(TmpGrs1!Chq_No))
                                        If Not IsNull(TmpGrs1!Chq_Date) Then DNarrStr = DNarrStr + " Dt: " + CStr(Format(TmpGrs1!Chq_Date, "dd/MM/yy"))
                                        D1mNarrStr2 = Trim(FaXNull(TmpGrs1!Narr))
                                        If D1NarrStr <> "" Or D1mNarrStr1 <> "" Or D1mNarrStr2 <> "" Then
                                            mFLAG2222 = True
                                        Else
                                            mFLAG2222 = False
                                        End If
                                    Else
                                        mFLAG2222 = False
                                    End If
                                    TmpGrs1.MoveNext
                                End If
                            Else
                                If FGrid.TextMatrix(Cat1, 1) = "Yes" Then
                                    If Trim(D1NarrStr) <> "" Then
                                        RstTmp!Name1 = Space(2) + FaSetW(D1NarrStr, 35)
                                        D1NarrStr = mID(D1NarrStr, 36, 300)
                                    End If
                                    If Trim(D1mNarrStr2) <> "" Then
                                        RstTmp!Name1 = Space(2) + FaSetW(D1mNarrStr2, 35)
                                        D1mNarrStr2 = mID(D1mNarrStr2, 36, 300)
                                    End If
                                End If
                            End If
                            If D1NarrStr = "" And D1mNarrStr1 = "" And D1mNarrStr2 = "" Then
                                mFLAG2222 = False
                            End If
                            If TmpGrs1.EOF = True Then
                                If D1NarrStr = "" And D1mNarrStr1 = "" And D1mNarrStr2 = "" Then
                                    mFLAG22 = True
                                    mFLAG222 = True
                                End If
                            End If
                        End If
                    Else
                        RstTmp!Name1 = Space(2) + Trim(mID(mNARR2, 1, 36))
                        mNARR2 = Trim(mID(mNARR2, 37, 510))
                        RstTmp!ADJAMT = Format(mGROUP_rs!AMOUNT, "0.00")
                        oBAL = oBAL + Format(mGROUP_rs!AMOUNT, "0.00")
                        mFLAG2222 = True
                        mFLAG222 = True
                        mFLAG22 = True
                    End If
                End If
                If Len(mNARR2) <= 0 And mFLAG22 = True And mFLAG222 = True Then
                    mFLAG2 = False
                    mFLAG22 = False
                    mFLAG222 = False
                    mFLAG2222 = False
                    mGROUP_rs.MoveNext
                    If Not mGROUP_rs.EOF Then
                        mDate2 = mGROUP_rs!V_DATE
                    Else
                        mDate2 = DateAdd("D", 1, TXTE_DATE)
                    End If
                End If
            Else
                mNARR2 = Trim(mNARR2)
                RstTmp!Name1 = Space(2) + Trim(mID(mNARR2, 1, 36))
                mNARR2 = Trim(mID(mNARR2, 37, 510))
                If Len(mNARR2) <= 0 Then
                    mFLAG2 = False
                    mFLAG22 = False
                    mFLAG222 = False
                    mFLAG2222 = False
                    mGROUP_rs.MoveNext
                    If Not mGROUP_rs.EOF Then
                        mDate2 = mGROUP_rs!V_DATE
                    Else
                        mDate2 = DateAdd("D", 1, TXTE_DATE)
                    End If
                End If
            End If
        End If
        RstTmp.Update
    Else
        If mDate1 <= mDate2 Then
            If mDate1 = CDate("12:00:00 AM") Then
                TmpDate = mDate2
            Else
                TmpDate = mDate1
            End If
        Else
            If mDate2 = CDate("12:00:00 AM") Then
                TmpDate = mDate1
            Else
                TmpDate = mDate2
            End If
        End If
        If oBAL <> 0 Then
            RstTmp.AddNew
            RstTmp!V_DATE = TmpDate
            If oBAL < 0 Then
                RstTmp!Name = "OPENING BALANCE"
                RstTmp!cr = Abs(oBAL)
            Else
                RstTmp!Name1 = "OPENING BALANCE"
                RstTmp!ADJAMT = Abs(oBAL)
            End If
            RstTmp.Update
        End If
    End If
Loop
If RstTmp.RecordCount > 0 Then RstTmp.MoveFirst
If RstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
Select Case Index
    Case 1
        MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
        TinTin = PubDatamanFa.FaCBookDosPrinting(Me, RstTmp, Val(FGrid.TextMatrix(List2, 1)))
        GoTo EXIT_SUB
    Case Else
'        X11 = CreateFieldDefFile(RstTmp, PubFaReportPath + "\Facashbook.ttx", True)
        If RstEnviro!LedDivCode = "No" And RstEnviro!LedSiteCode = "No" And RstEnviro!LedPrefix = "No" Then
            Set rpt = PubDatamanFa.FaCashbookPortraitRpt
        Else
            Set rpt = PubDatamanFa.FaCashbookRpt
        End If
        rpt.Database.SetDataSource RstTmp
End Select
EXIT_SUB:
    Set Rst1 = Nothing
    Set mGROUP_rs = Nothing
    Set SUBGROUP_rs = Nothing
    Set TmpGrs = Nothing
    Set TmpGrs1 = Nothing
    Set RstTmp = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub DayBook(Index As Integer)
On Error GoTo ELoop
Dim Rst1 As ADODB.Recordset, RstTmp As ADODB.Recordset
Dim TinTin As Integer, I As Integer, mSiteCode As String
Dim mV_NO As Long, mNarr As String, mNARR1 As String, mNARR2 As String, mDocNo As String
Dim TmpDate As Date, mDocId As String, mS_NO As Long
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    TXTS_DATE = FGrid.TextMatrix(Date1, 1)
    TXTE_DATE = FGrid.TextMatrix(Date2, 1)
    If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub
    
    If Trim(FGrid.TextMatrix(List2, 2)) <> "" Then
        If PubFaSiteType = 1 Then
            mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & FGrid.TextMatrix(List2, 2) & "'"
        ElseIf PubFaSiteType = 2 Then
            If PubSiteCodeWidth = 1 Then
                mSiteCode = " And VIEWLEDGER.Site_Code='" & FGrid.TextMatrix(List2, 2) & "'"
            Else
                mSiteCode = " And VIEWLEDGER.Site_Code='" & Trim(FGrid.TextMatrix(List2, 2)) + Trim(FGrid.TextMatrix(List2, 2)) & "'"
            End If
        End If
    Else
        mSiteCode = ""
    End If
    Set Rst1 = G_FaCn.Execute("SELECT VIEWLEDGER.*,SUBGROUP.NAME FROM VIEWLEDGER LEFT JOIN SUBGROUP ON VIEWLEDGER.PARTY=SUBGROUP.SUBCODE WHERE V_dATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,V_SNo")
    Set RstTmp = New ADODB.Recordset
    Set RstTmp = PubDatamanFa.FaDayBook(RstTmp)
    mV_NO = 0
    Do Until Rst1.EOF
        If mV_NO = 0 Then mV_NO = Rst1!V_NO
        mNarr = ""
        mNARR1 = ""
        mNARR2 = ""
        mDocNo = ""
        If FaXNull(Rst1!Chq_No) <> "" Then mNarr = mNarr + "Chq.No:" + Trim(FaXNull(Rst1!Chq_No))
        If Not IsNull(Rst1!Chq_Date) Then mNarr = mNarr + " Dt: " + CStr(Format(Rst1!Chq_Date, "dd/MM/yy"))
        mNARR1 = FaXNull(Rst1!Narr)
        mNARR2 = FaXNull(Rst1!mNarr)
        TmpDate = Format(Rst1!V_DATE, "dd/MMM/yyyy")
        mDocId = Rst1!DocId
        mS_NO = Rst1!V_SNo
        With RstTmp
            .AddNew
            !PDATE = Rst1!V_DATE
            !V_DATE = Rst1!V_DATE
            !V_tYPE = Rst1!V_tYPE
            !V_NO = Rst1!V_NO
            !V_SNo = Rst1!V_SNo
            !SubCode = FaXNull(Rst1!Party)
            !Name = FaXNull(Rst1!Name)
            !cr = Format(Rst1!CREDIT, "0.00")
            !dr = Format(Rst1!DEBIT, "0.00")
            If PubFaSiteType = 1 Then
                mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(Rst1!DocId, 1), "")
                mDocNo = mDocNo + IIf(RstEnviro!LedSiteCode = "Yes", Trim(Right(Rst1!Site_Code, 1)), "")
                mDocNo = mDocNo + IIf(RstEnviro!LedPrefix = "Yes", IIf(mDocNo = "", "", "/") + Trim(Rst1!V_ADD), "")
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + left(Trim(Rst1!V_tYPE), 1) + Trim(mID(Trim(Rst1!V_tYPE), 3, 3))
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(Rst1!V_NO))
            Else
                mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(Rst1!DocId, 1), "") + IIf(RstEnviro!LedSiteCode = "Yes", Trim(left(Rst1!Site_Code, 1)), "")
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + IIf(RstEnviro!LedPrefix = "Yes", Trim(Rst1!V_ADD), "")
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(Rst1!V_tYPE)
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(Rst1!V_NO))
            End If
            !DOCNO = mDocNo
            .Update
        End With
        If Trim(mNARR1) <> "" Then
            Do While Len(mNARR1) > 0
                With RstTmp
                    .MoveLast
                    If Trim(RstTmp.Fields("Name")) <> "" Then .AddNew
                    .Fields("Name") = left(mNARR1, 50)
                    .Fields("DocId") = Rst1!DocId
                    .Fields("PDate") = Rst1!V_DATE
                    !V_SNo = mS_NO
                    !DOCNO = mDocNo
                    If Rst1!V_tYPE = "OP" Then
                        .Fields("Val") = "1"
                    Else
                        .Fields("Val") = IIf(Rst1!CREDIT > 0, "2", "3")
                    End If
                    .Update
                End With
                mNARR1 = mID(mNARR1, 51, 300)
            Loop
        End If
        Rst1.MoveNext
        If Rst1.EOF = True Then
NARRLOOP:
            If Trim(mNARR1) <> "" Then
                Do While Len(mNARR1) > 0
                    With RstTmp
                        .MoveLast
                        If Trim(RstTmp.Fields("Name")) <> "" Then .AddNew
                        .Fields("Name") = left(mNARR1, 50)
                        .Fields("DocId") = mDocId
                        .Fields("PDate") = TmpDate
                        !V_SNo = mS_NO
                        !DOCNO = mDocNo
                        .Fields("Val") = "4"
                        .Update
                    End With
                    mNARR1 = mID(mNARR1, 51, 300)
                Loop
            End If
            If Trim(mNARR2) <> "" Then
                Do While Len(mNARR2) > 0
                    With RstTmp
                        .MoveLast
                        If Trim(RstTmp.Fields("Name")) <> "" Then .AddNew
                        .Fields("Name") = left(mNARR2, 50)
                        .Fields("DocId") = mDocId
                        .Fields("PDate") = TmpDate
                        !V_SNo = mS_NO
                        !DOCNO = mDocNo
                        .Fields("Val") = "5"
                        .Update
                    End With
                    mNARR2 = mID(mNARR2, 51, 300)
                Loop
            End If
        Else
            If mV_NO <> Rst1!V_NO Then
                mV_NO = 0
                GoTo NARRLOOP
            End If
        End If
    Loop
    If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "DayBook"
    Select Case Index
        Case 1
            MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
            TinTin = PubDatamanFa.FaDayBookDosPrinting(Me, RstTmp, FGrid.TextMatrix(List1, 1))
            GoTo EXIT_SUB
        Case 0
            TinTin = CreateFieldDefFile(RstTmp, PubFaReportPath + "\FaDayBook.ttx", True)
            Set rpt = PubDatamanFa.FaDayBookRpt
            rpt.Database.SetDataSource RstTmp
    End Select
EXIT_SUB:
    Set Rst1 = Nothing
    Set RstTmp = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub DetailedTrial(Index As Integer)
'On Error GoTo ELoop
Dim mQRY As String, mQRY1 As String, mQRY_OP As String, mQRY1_OP As String, Rst1 As ADODB.Recordset, Rst As ADODB.Recordset, Rst2 As ADODB.Recordset
Dim TinTin As Integer, I As Integer
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    TXTS_DATE = FGrid.TextMatrix(Date1, 1)
    TXTE_DATE = FGrid.TextMatrix(Date2, 1)
    If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub
    If FGrid.TextMatrix(List2, 1) = "Yes" Then
        mQRY1 = " Where ((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND VIEWSUBGROUP.GroupNature IN ('E','R')) OR (V_DATE<" & FaConvertDate(TXTS_DATE) & " AND VIEWSUBGROUP.GroupNature NOT IN ('E','R')))"
        mQRY = " Where ((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND VIEWSUBGROUP.GroupNature IN ('E','R')) OR (V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND VIEWSUBGROUP.GroupNature NOT IN ('E','R')))"
        Set Rst1 = New ADODB.Recordset
        If FGrid.TextMatrix(List1, 1) = "Yes" Then
            If PubBackEnd = "A" Then
                Rst1.Open ("SELECT 1 AS TT,MAX(ViewSubgroup.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS OP_CR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS OP_dR,0 AS BALANCECR,0 AS BALANCEDR,0 As Bal,MAX(ViewSubgroup.MAINGRCODES) AS mgrcode,Max(AA.GROUPNAME) AS GRName,Max(AA.GROUPNATURE) AS GRNATURE FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE ) LEFT JOIN ACGROUP  AA ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,3) " & mQRY1 & " AND GAliasYN='N' AND aa.AliasYN='N' GROUP BY VIEWSUBGROUP.GROUPNAME,LEDGER.SUBCODE Union " & _
                           "SELECT 2 AS TT,MAX(ViewSubgroup.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS BALANCECR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS BALANCEDR,0 As Bal,MAX(ViewSubgroup.MAINGRCODES) AS mgrcode,Max(AA.GROUPNAME) AS GRName,Max(AA.GROUPNATURE) AS GRNATURE FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE ) LEFT JOIN ACGROUP  AA ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,3) " & mQRY & " AND GAliasYN='N' AND aa.AliasYN='N' GROUP BY VIEWSUBGROUP.GROUPNAME,LEDGER.SUBCODE "), G_FaCn, adOpenDynamic, adLockOptimistic
            ElseIf PubBackEnd = "S" Then
                Rst1.Open ("SELECT 1 AS TT,MAX(ViewSubgroup.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,ISNULL(sum(AMTCR),0) AS OP_CR,ISNULL(SUM(AMTDR),0) AS OP_dR,0 AS BALANCECR,0 AS BALANCEDR,0 As Bal,MAX(ViewSubgroup.MAINGRCODES) AS mgrcode,Max(AA.GROUPNAME) AS GRName ,Max(AA.GROUPNATURE) AS GRNATURE FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE ) LEFT JOIN ACGROUP  AA ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,3) " & mQRY1 & " AND GAliasYN='N' AND aa.AliasYN='N' GROUP BY VIEWSUBGROUP.GROUPNAME,LEDGER.SUBCODE Union " & _
                           "SELECT 2 AS TT,MAX(ViewSubgroup.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,ISNULL(sum(AMTCR),0) AS BALANCECR,ISNULL(SUM(AMTDR),0) AS BALANCEDR,0 As Bal,MAX(ViewSubgroup.MAINGRCODES) AS mgrcode,Max(AA.GROUPNAME) AS GRName ,Max(AA.GROUPNATURE) AS GRNATURE FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE ) LEFT JOIN ACGROUP  AA ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,3) " & mQRY & " AND GAliasYN='N' AND aa.AliasYN='N' GROUP BY VIEWSUBGROUP.GROUPNAME,LEDGER.SUBCODE "), G_FaCn, adOpenDynamic, adLockOptimistic
            End If
        Else
            If PubBackEnd = "A" Then
                Rst1.Open ("SELECT 2 AS TT,MAX(ViewSubgroup.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS BALANCECR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS BALANCEDR,0 As Bal,MAX(ViewSubgroup.MAINGRCODES) AS mgrcode,Max(AA.GROUPNAME) AS GRName ,Max(AA.GROUPNATURE) AS GRNATURE FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE ) LEFT JOIN ACGROUP  AA ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,3) " & mQRY & " AND GAliasYN='N' AND aa.AliasYN='N' GROUP BY VIEWSUBGROUP.GROUPNAME,LEDGER.SUBCODE "), G_FaCn, adOpenDynamic, adLockOptimistic
            ElseIf PubBackEnd = "S" Then
                Rst1.Open ("SELECT 2 AS TT,MAX(ViewSubgroup.GROUPNAME)As GroupName,'' AS GrCode,MAX(LEDGER.SUBCODE) AS PARTYCODE,MAX(ViewSubgroup.NAME) AS PARTYNAME,0 AS OP_CR,0 AS OP_DR,ISNULL(sum(AMTCR),0) AS BALANCECR,ISNULL(SUM(AMTDR),0) AS BALANCEDR,0 As Bal,MAX(ViewSubgroup.MAINGRCODES) AS mgrcode,Max(AA.GROUPNAME) AS GRName ,Max(AA.GROUPNATURE) AS GRNATURE FROM LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE " & mQRY & " AND GAliasYN='N' AND aa.AliasYN='N' GROUP BY VIEWSUBGROUP.GROUPNAME,LEDGER.SUBCODE "), G_FaCn, adOpenDynamic, adLockOptimistic
            End If
        End If
    Else
        mQRY_OP = " Where (((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND AA.GroupNature IN ('E','R') AND AA.AliasYN='N' AND VIEWSUBGROUP.GroupNature IN ('E','R') AND VIEWSUBGROUP.AliasYN='N') OR (V_DATE<" & FaConvertDate(TXTS_DATE) & " AND AA.GroupNature NOT IN ('E','R') AND AA.AliasYN='N' AND VIEWSUBGROUP.GroupNature NOT IN ('E','R') AND VIEWSUBGROUP.AliasYN='N')))"
        mQRY1_OP = " Where (((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND AA.GroupNature IN ('E','R')  AND AA.AliasYN='N' AND VIEWSUBGROUP.GroupNature IN ('E','R') AND VIEWSUBGROUP.AliasYN='N') OR (V_DATE<" & FaConvertDate(TXTS_DATE) & " AND AA.GroupNature NOT IN ('E','R') AND AA.AliasYN='N' AND VIEWSUBGROUP.GroupNature NOT IN ('E','R') AND VIEWSUBGROUP.AliasYN='N')))"
        
        mQRY = "  Where (((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND AA.GroupNature IN ('E','R') AND AA.ALIASYN='N' AND VIEWSUBGROUP.GroupNature IN ('E','R') AND VIEWSUBGROUP.AliasYN='N') OR (V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND AA.GroupNature NOT IN ('E','R') AND AA.AliasYN='N' AND VIEWSUBGROUP.GroupNature NOT IN ('E','R') AND VIEWSUBGROUP.AliasYN='N')))"
        mQRY1 = " Where (((V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND AA.GroupNature IN ('E','R') AND AA.ALIASYN='N' AND VIEWSUBGROUP.GroupNature IN ('E','R') AND VIEWSUBGROUP.AliasYN='N') OR (V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND AA.GroupNature NOT IN ('E','R') AND AA.AliasYN='N' AND VIEWSUBGROUP.GroupNature NOT IN ('E','R') AND VIEWSUBGROUP.AliasYN='N')))"
        
        If PubBackEnd = "A" Then
            If FGrid.TextMatrix(List1, 1) = "Yes" Then
                Set Rst = G_FaCn.Execute("SELECT 1 AS TT,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS OP_CR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS OP_dR,0                                    AS BALANCECR,0                                    AS BALANCEDR,0 As Bal,MAX(BB.MAINGRCODE) AS mgrcode,MAX(BB.GROUPNAME) AS GRName,MAX(BB.GroupCode) As GRCode,MAX(AA.MAINGRCODE)   AS mgrcodeSUB,MAX(AA.GROUPNAME) AS GRNameSUB ,MAX(AA.GROUPCODE)  AS GRCodeSUB FROM ((ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE) LEFT JOIN ACGROUP BB ON BB.MAINGRCODE=LEFT(AA.MAINGRCODE ,3) " & mQRY_OP & "  AND BB.AliasYN='N' GROUP BY AA.MAINGRCODE,AA.GROUPCODE HAVING LEN(AA.MAINGRCODE)=6  UNION " & _
                                          "SELECT 2 AS TT,0                                    AS OP_CR,0                                    AS OP_DR,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS BALANCECR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS BALANCEDR,0 As Bal,MAX(BB.MAINGRCODE) AS mgrcode,MAX(BB.GROUPNAME) AS GRName,MAX(BB.GroupCode) As GRCode,MAX(AA.MAINGRCODE)   AS mgrcodeSUB,MAX(AA.GROUPNAME) AS GRNameSUB ,MAX(AA.GROUPCODE)  AS GRCodeSUB FROM ((ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE) LEFT JOIN ACGROUP BB ON BB.MAINGRCODE=LEFT(AA.MAINGRCODE ,3) " & mQRY & "  AND BB.AliasYN='N' GROUP BY AA.MAINGRCODE,AA.GROUPCODE HAVING LEN(AA.MAINGRCODE)=6 ")
            Else
                Set Rst = G_FaCn.Execute("SELECT 2 AS TT,0                                    AS OP_CR,0                                    AS OP_DR,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS BALANCECR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS BALANCEDR,0 As Bal,MAX(BB.MAINGRCODE) AS mgrcode,MAX(BB.GROUPNAME) AS GRName,MAX(BB.GroupCode) As GRCode,MAX(AA.MAINGRCODE)   AS mgrcodeSUB,MAX(AA.GROUPNAME) AS GRNameSUB ,MAX(AA.GROUPCODE)  AS GRCodeSUB FROM ((ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE) LEFT JOIN ACGROUP BB ON BB.MAINGRCODE=LEFT(AA.MAINGRCODE ,3) " & mQRY & "  AND BB.AliasYN='N' GROUP BY AA.MAINGRCODE,AA.GROUPCODE HAVING LEN(AA.MAINGRCODE)=6 ")
            End If
        ElseIf PubBackEnd = "S" Then
            If FGrid.TextMatrix(List1, 1) = "Yes" Then
                Set Rst = G_FaCn.Execute("SELECT 1 AS TT,ISNULL(sum(AMTCR),0) AS OP_CR,ISNULL(SUM(AMTDR),0) AS OP_dR,0                    AS BALANCECR,0                    AS BALANCEDR,0 As Bal,MAX(BB.MAINGRCODE) AS mgrcode,MAX(BB.GROUPNAME) AS GRName,MAX(BB.GroupCode) As GRCode,MAX(AA.MAINGRCODE)   AS mgrcodeSUB,MAX(AA.GROUPNAME) AS GRNameSUB ,MAX(AA.GROUPCODE)  AS GRCodeSUB FROM ((ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE) LEFT JOIN ACGROUP BB ON BB.MAINGRCODE=LEFT(AA.MAINGRCODE ,3) " & mQRY_OP & "   AND BB.AliasYN='N'   GROUP BY AA.MAINGRCODE,AA.GROUPCODE HAVING LEN(AA.MAINGRCODE)=6 UNION " & _
                                          "SELECT 2 AS TT,0                    AS OP_CR,0                    AS OP_DR,ISNULL(sum(AMTCR),0) AS BALANCECR,ISNULL(SUM(AMTDR),0) AS BALANCEDR,0 As Bal,MAX(BB.MAINGRCODE) AS mgrcode,MAX(BB.GROUPNAME) AS GRName,MAX(BB.GroupCode) As GRCode,MAX(AA.MAINGRCODE)   AS mgrcodeSUB,MAX(AA.GROUPNAME) AS GRNameSUB ,MAX(AA.GROUPCODE)  AS GRCodeSUB FROM ((ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE) LEFT JOIN ACGROUP BB ON BB.MAINGRCODE=LEFT(AA.MAINGRCODE ,3) " & mQRY & "  AND BB.AliasYN='N' GROUP BY AA.MAINGRCODE,AA.GROUPCODE HAVING LEN(AA.MAINGRCODE)=6 ")
            Else
                Set Rst = G_FaCn.Execute("SELECT 2 AS TT,0                    AS OP_CR,0                    AS OP_DR,ISNULL(sum(AMTCR),0) AS BALANCECR,ISNULL(SUM(AMTDR),0) AS BALANCEDR,0 As Bal,MAX(BB.MAINGRCODE) AS mgrcode,MAX(BB.GROUPNAME) AS GRName,MAX(BB.GroupCode) As GRCode,MAX(AA.MAINGRCODE)   AS mgrcodeSUB,MAX(AA.GROUPNAME) AS GRNameSUB ,MAX(AA.GROUPCODE)  AS GRCodeSUB FROM ((ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE) LEFT JOIN ACGROUP BB ON BB.MAINGRCODE=LEFT(AA.MAINGRCODE ,3) " & mQRY & "  AND BB.AliasYN='N' GROUP BY AA.MAINGRCODE,AA.GROUPCODE HAVING LEN(AA.MAINGRCODE)=6 ")
            End If
        End If
        Set Rst1 = New ADODB.Recordset
        Set Rst1 = mDetailTrial(Rst1)
        Do Until Rst.EOF
            With Rst1
                .AddNew
                .Fields("TT") = Rst!TT
                .Fields("GroupName") = "" 'Rst!GRName
                .Fields("GRCODE") = "" 'Rst!mgrcode
                .Fields("PARTYCODE") = Rst!mgrcodeSUB
                .Fields("PARTYNAME") = Rst!GRNameSUB
                .Fields("OP_CR") = Rst!OP_CR
                .Fields("OP_dR") = Rst!OP_DR
                .Fields("BALANCECR") = Rst!BALANCECR
                .Fields("BALANCEDR") = Rst!BALANCEDR
                .Fields("Bal") = Rst!Bal
                .Fields("mgrcode") = Rst!mgrcode
                .Fields("GRName") = Rst!GRName
                .Update
            End With
            Rst.MoveNext
        Loop
        
         If PubBackEnd = "A" Then
            If FGrid.TextMatrix(List1, 1) = "Yes" Then
                Set Rst = G_FaCn.Execute("SELECT 1 AS TT,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS OP_CR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS OP_dR,0                                    AS BALANCECR,0                                    AS BALANCEDR,0 As Bal,AA.MAINGRCODE AS mgrcode,MAX(AA.GROUPNAME) AS GRName,MAX(aa.GroupCode) As GRCode,MAX(VIEWSUBGROUP.MAINGRCODES)   AS mgrcodeSUB,MAX(VIEWSUBGROUP.NAME) AS GRNameSUB ,MAX(VIEWSUBGROUP.SUBCODE)  AS GRCodeSUB FROM (ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mQRY1_OP & "    GROUP BY AA.GROUPCODE,AA.MAINGRCODE,VIEWSUBGROUP.GROUPCODE,VIEWSUBGROUP.SUBCODE HAVING LEN(MAX(VIEWSUBGROUP.MAINGRCODES))=3 UNION " & _
                                          "SELECT 2 AS TT,0                                    AS OP_CR,0                                    AS OP_DR,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS BALANCECR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS BALANCEDR,0 As Bal,AA.MAINGRCODE AS mgrcode,MAX(AA.GROUPNAME) AS GRName,MAX(aa.GroupCode) As GRCode,MAX(VIEWSUBGROUP.MAINGRCODES)   AS mgrcodeSUB,MAX(VIEWSUBGROUP.NAME) AS GRNameSUB ,MAX(VIEWSUBGROUP.SUBCODE)  AS GRCodeSUB FROM (ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mQRY1 & " GROUP BY AA.GROUPCODE,AA.MAINGRCODE,VIEWSUBGROUP.GROUPCODE,VIEWSUBGROUP.SUBCODE HAVING LEN(MAX(VIEWSUBGROUP.MAINGRCODES))=3 ")
            Else
                Set Rst = G_FaCn.Execute("SELECT 2 AS TT,0                                    AS OP_CR,0                                    AS OP_DR,IIF(ISNULL(sum(AMTCR)),0,sum(AMTCR)) AS BALANCECR,IIF(ISNULL(SUM(AMTDR)),0,SUM(AMTDR)) AS BALANCEDR,0 As Bal,AA.MAINGRCODE AS mgrcode,MAX(AA.GROUPNAME) AS GRName,MAX(aa.GroupCode) As GRCode,MAX(VIEWSUBGROUP.MAINGRCODES)   AS mgrcodeSUB,MAX(VIEWSUBGROUP.NAME) AS GRNameSUB ,MAX(VIEWSUBGROUP.SUBCODE)  AS GRCodeSUB FROM (ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mQRY1 & " GROUP BY AA.GROUPCODE,AA.MAINGRCODE,VIEWSUBGROUP.GROUPCODE,VIEWSUBGROUP.SUBCODE HAVING LEN(MAX(VIEWSUBGROUP.MAINGRCODES))=3")
            End If
        ElseIf PubBackEnd = "S" Then
            If FGrid.TextMatrix(List1, 1) = "Yes" Then
                Set Rst = G_FaCn.Execute("SELECT 1 AS TT,ISNULL(sum(AMTCR),0) AS OP_CR,ISNULL(SUM(AMTDR),0) AS OP_dR,0                    AS BALANCECR,0                    AS BALANCEDR,0 As Bal,AA.MAINGRCODE AS mgrcode,MAX(AA.GROUPNAME) AS GRName,MAX(aa.GroupCode) As GRCode,MAX(VIEWSUBGROUP.MAINGRCODES)   AS mgrcodeSUB,MAX(VIEWSUBGROUP.NAME) AS GRNameSUB ,MAX(VIEWSUBGROUP.SUBCODE)  AS GRCodeSUB FROM (ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mQRY1_OP & "    GROUP BY AA.GROUPCODE,AA.MAINGRCODE,VIEWSUBGROUP.GROUPCODE,VIEWSUBGROUP.SUBCODE HAVING LEN(MAX(VIEWSUBGROUP.MAINGRCODES))=3 UNION " & _
                                          "SELECT 2 AS TT,0                    AS OP_CR,0                    AS OP_DR,ISNULL(sum(AMTCR),0) AS BALANCECR,ISNULL(SUM(AMTDR),0) AS BALANCEDR,0 As Bal,AA.MAINGRCODE AS mgrcode,MAX(AA.GROUPNAME) AS GRName,MAX(aa.GroupCode) As GRCode,MAX(VIEWSUBGROUP.MAINGRCODES)   AS mgrcodeSUB,MAX(VIEWSUBGROUP.NAME) AS GRNameSUB ,MAX(VIEWSUBGROUP.SUBCODE)  AS GRCodeSUB FROM (ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mQRY1 & " GROUP BY AA.GROUPCODE,AA.MAINGRCODE,VIEWSUBGROUP.GROUPCODE,VIEWSUBGROUP.SUBCODE HAVING LEN(MAX(VIEWSUBGROUP.MAINGRCODES))=3")
            Else
                Set Rst = G_FaCn.Execute("SELECT 2 AS TT,0                    AS OP_CR,0                    AS OP_DR,ISNULL(sum(AMTCR),0) AS BALANCECR,ISNULL(SUM(AMTDR),0) AS BALANCEDR,0 As Bal,AA.MAINGRCODE AS mgrcode,MAX(AA.GROUPNAME) AS GRName,MAX(aa.GroupCode) As GRCode,MAX(VIEWSUBGROUP.MAINGRCODES)   AS mgrcodeSUB,MAX(VIEWSUBGROUP.NAME) AS GRNameSUB ,MAX(VIEWSUBGROUP.SUBCODE)  AS GRCodeSUB FROM (ACGROUP AA INNER JOIN VIEWSUBGROUP ON AA.MAINGRCODE=LEFT(VIEWSUBGROUP.MAINGRCODES,LEN(AA.MAINGRCODE))) LEFT JOIN LEDGER ON LEDGER.SUBCODE=VIEWSUBGROUP.SUBCODE " & mQRY1 & " GROUP BY AA.GROUPCODE,AA.MAINGRCODE,VIEWSUBGROUP.GROUPCODE,VIEWSUBGROUP.SUBCODE HAVING LEN(MAX(VIEWSUBGROUP.MAINGRCODES))=3")
            End If
        End If
        Do Until Rst.EOF
            With Rst1
                .AddNew
                .Fields("TT") = Rst!TT
                .Fields("GroupName") = "" 'Rst!GRName
                .Fields("GRCODE") = "" 'Rst!mgrcode
                .Fields("PARTYCODE") = Rst!mgrcodeSUB
                .Fields("PARTYNAME") = Rst!GRNameSUB
                .Fields("OP_CR") = Rst!OP_CR
                .Fields("OP_dR") = Rst!OP_DR
                .Fields("BALANCECR") = Rst!BALANCECR
                .Fields("BALANCEDR") = Rst!BALANCEDR
                .Fields("Bal") = Rst!Bal
                .Fields("mgrcode") = Rst!mgrcode
                .Fields("GRName") = Rst!GRName
                .Update
            End With
            Rst.MoveNext
        Loop
        
    End If
    If Rst1.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    Select Case Index
        Case 1
                MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
                TinTin = PubDatamanFa.FaDetailedTrialDosPrinting(Me, Rst1)
                GoTo EXIT_SUB
            Case Else
            Dim X1
                X1 = CreateFieldDefFile(Rst1, PubFaReportPath + "\FaDetTrial.ttx", True)
                Set rpt = PubDatamanFa.FaDetTrialRpt
                rpt.Database.SetDataSource Rst1
    End Select
EXIT_SUB:
    Set Rst = Nothing
    Set Rst1 = Nothing
    Set Rst2 = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Sub DGAccount_Click()
    DGAccount.Visible = False
    If RstAccount.RecordCount > 0 Then
        TxtGrid(Val(DGAccount.Tag)).Tag = RstAccount!Code
        TxtGrid(Val(DGAccount.Tag)).TEXT = RstAccount!Name
    End If
    TxtGrid(Val(DGAccount.Tag)).SetFocus
End Sub
Private Sub Led(Index As Integer, Optional mType As String)
Dim Rst1 As ADODB.Recordset, TmpRst As ADODB.Recordset, Rst2 As ADODB.Recordset, RstCheck21 As ADODB.Recordset
Dim mReportStype As String, Condstr As String, ac_str As String, Ac_Name As String, Ac_Code As String
Dim NarrStr As String, mNarrStr1 As String, mNarrStr2 As String, mDocNo As String, X1, TinTin As Integer, mSepratePage As String
Dim DNarrStr As String, DmNarrStr1 As String, DmNarrStr2 As String, Qry As String, mSiteCode As String
Dim oBAL As Double, NextDate As Date, I As Integer
Dim RstDrX As ADODB.Recordset, RstCrX As ADODB.Recordset
On Error GoTo Errloop
If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then GoTo ExitLoop
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then GoTo ExitLoop
If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then GoTo ExitLoop
TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
'If FaValidDate(Me) = 0 Then GoTo ExitLoop
If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And VIEWLEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
End If
If mType = "LedInt" Then If FaIsValid(TEXT(Interest), "Enter Interest Rate") = False Then GoTo ExitLoop
mReportStype = FGrid.TextMatrix(List1, 1)
If mType = "AcCheckList" Then
    Qry = ""
    GridString3 = ""
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then GoTo ExitLoop
    If GridString3 <> "" Then Qry = " AND V_TYPE IN (" & GridString3 & ") "
    If Val(TEXT(TxNAmount)) > 0 Then Qry = Qry + " AND VIEWLEDGER.DEBIT+VIEWLEDGER.CREDIT" & Trim(TEXT(10)) & " " & Val(TEXT(TxNAmount))
    If PubBackEnd = "S" Then
        If Trim(TEXT(NarrationHaving)) <> "" Then Qry = Qry + " AND VIEWLEDGER.NARR LIKE '%" & Trim(TEXT(NarrationHaving)) & "%'"
        If Trim(TEXT(NarrationNotHaving)) <> "" Then Qry = Qry + " AND VIEWLEDGER.NARR NOT LIKE '%" & Trim(TEXT(NarrationNotHaving)) & "%'"
    ElseIf PubBackEnd = "A" Then
        If Trim(TEXT(NarrationHaving)) <> "" Then Qry = Qry + " AND (INSTR(1,UCASE(VIEWLEDGER.NARR),UCASE(TRIM('" & TEXT(NarrationHaving) & "'))) OR INSTR(1,UCASE(VIEWLEDGER.MNARR),UCASE(TRIM('" & TEXT(NarrationHaving) & "'))))"
        If Trim(TEXT(NarrationNotHaving)) <> "" Then Qry = Qry + " AND NOT (INSTR(1,UCASE(VIEWLEDGER.NARR),UCASE(TRIM('" & TEXT(NarrationNotHaving) & "'))) OR INSTR(1,UCASE(VIEWLEDGER.MNARR),UCASE(TRIM('" & TEXT(NarrationNotHaving) & "'))))"
    End If
End If
Select Case mReportStype
    Case "Selected"
        GridString1 = ""
        Condstr = ""
        If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then GoTo ExitLoop
        If GridString1 <> "" Then Condstr = " WHERE SUBCODE IN (" & GridString1 & ") "
        Select Case mType
            Case "LedDeb"
                Condstr = Condstr + IIf(Condstr = "", " Where", " AND") + " Nature='Customer'"
            Case "LedCred"
                Condstr = Condstr + IIf(Condstr = "", " Where", " AND") + " Nature='Supplier'"
        End Select
        Set Rst1 = G_FaCn.Execute("SELECT GROUPNATURE,groupCode as CODE,SUBCODE,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST " & Condstr & " ORDER BY NAME")
    Case "Merge"
        If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then GoTo ExitLoop
        If IsNotBlank(List3, FGrid.TextMatrix(List3, 0)) = False Then GoTo ExitLoop
        TXTACC_CODE = FGrid.TextMatrix(List2, 1)
        TXTACC_CODE.Tag = FGrid.TextMatrix(List2, 2)
        TXTACC_CODE1 = FGrid.TextMatrix(List3, 1)
        TXTACC_CODE1.Tag = FGrid.TextMatrix(List3, 2)
        If TXTACC_CODE = TXTACC_CODE1 Then MsgBox " ** Both A/C are Same** ", vbCritical, Me.CAPTION:  GoTo ExitLoop
        Set Rst1 = G_FaCn.Execute("SELECT GROUPNATURE,CODE,SUBCODE,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST WHERE SUBCODE IN ('" & TXTACC_CODE.Tag & "','" & TXTACC_CODE1.Tag & "') ORDER BY NAME")
    Case "Group(Selected)"
        GridString2 = ""
        Condstr = ""
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then GoTo ExitLoop
        If GridString2 <> "" Then Condstr = " WHERE CODE IN (" & GridString2 & ") "
        Set Rst1 = G_FaCn.Execute("SELECT GROUPNATURE,groupCode as CODE,SUBCODE,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST " & Condstr & " ORDER BY NAME")
    Case "Group(Range)"
        If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then GoTo ExitLoop
        If IsNotBlank(List3, FGrid.TextMatrix(List3, 0)) = False Then GoTo ExitLoop
        If IsNotBlank(Cat1, FGrid.TextMatrix(Cat1, 0)) = False Then GoTo ExitLoop
        TXTACC_CODE = FGrid.TextMatrix(List3, 1)
        TXTACC_CODE.Tag = FGrid.TextMatrix(List3, 2)
        TXTACC_CODE1 = FGrid.TextMatrix(Cat1, 1)
        TXTACC_CODE1.Tag = FGrid.TextMatrix(Cat1, 2)
        Set Rst1 = G_FaCn.Execute("SELECT GROUPNATURE,CODE,SUBCODE,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST WHERE CODE='" & FGrid.TextMatrix(List2, 2) & "' AND NAME BETWEEN '" & TXTACC_CODE & "' AND '" & TXTACC_CODE1 & "' ORDER BY NAME")
    Case "City Wise"
        GridString2 = ""
        Condstr = ""
        If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then GoTo ExitLoop
        If GridString2 <> "" Then Condstr = " WHERE CITYCODE IN (" & GridString2 & ") "
        Select Case mType
            Case "LedDeb"
                Condstr = Condstr + IIf(Condstr = "", " Where", " AND") + " Nature='Customer'"
            Case "LedCred"
                Condstr = Condstr + IIf(Condstr = "", " Where", " AND") + " Nature='Supplier'"
        End Select
        Set Rst1 = G_FaCn.Execute("SELECT GROUPNATURE,groupCode as CODE,SUBCODE,NAME,ADD1,ADD2,CITY_NAME FROM PARTY_LIST " & Condstr & " ORDER BY NAME")
End Select
If Rst1.RecordCount <= 0 Then MsgBox ("No A/c to Print"): GoTo ExitLoop
Set TmpRst = New ADODB.Recordset
Set TmpRst = PubDatamanFa.FaADTMP1(TmpRst)
ac_str = ""
Ac_Name = ""
Do Until Rst1.EOF
    ac_str = ac_str + IIf(Trim(ac_str) = "", "", ",") + "'" + Trim(Rst1!SubCode) + "'"
    Ac_Name = Ac_Name + IIf(Trim(Ac_Name) = "", "", ",") + Trim(Rst1!Name)
    Rst1.MoveNext
Loop
Select Case GRepFormName
    Case "Led", "LedInt", "AcCheckList"
        If RstEnviro!ShowCityName = "Yes" Then
            If PubBackEnd = "A" Then
                Set Rst2 = G_FaCn.Execute("SELECT VIEWLEDGER.*,NameWithCity AS NAME FROM VIEWLEDGER LEFT JOIN ViewSubgroup ON VIEWLEDGER.PARTY1=ViewSubgroup.SUBCODE WHERE VIEWLEDGER.PARTY IN (" & ac_str & ") AND VIEWLEDGER.V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " " & Qry & "  " & mSiteCode & " ORDER BY PARTY,V_DATE,IIF(CREDIT>0,3,2),DOCID,V_SNo")
            Else
                Set Rst2 = G_FaCn.Execute("SELECT VIEWLEDGER.*,NameWithCity AS NAME FROM VIEWLEDGER LEFT JOIN ViewSubgroup ON VIEWLEDGER.PARTY1=ViewSubgroup.SUBCODE WHERE VIEWLEDGER.PARTY IN (" & ac_str & ") AND VIEWLEDGER.V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " " & Qry & "  " & mSiteCode & " ORDER BY PARTY,V_DATE,CASE WHEN CREDIT>0 THEN 3 WHEN DEBIT>0 THEN 2 ELSE 1 END,DOCID,V_SNo")
            End If
        Else
            If PubBackEnd = "A" Then
                Set Rst2 = G_FaCn.Execute("SELECT VIEWLEDGER.*,SUBGROUP.NAME FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY1=subgroup.SUBCODE WHERE VIEWLEDGER.PARTY IN (" & ac_str & ") AND VIEWLEDGER.V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " " & Qry & "  " & mSiteCode & " ORDER BY PARTY,V_DATE,IIF(CREDIT>0,3,2),DOCID,V_SNo")
            Else
                Set Rst2 = G_FaCn.Execute("SELECT VIEWLEDGER.*,SUBGROUP.NAME FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY1=subgroup.SUBCODE WHERE VIEWLEDGER.PARTY IN (" & ac_str & ") AND VIEWLEDGER.V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " " & Qry & "  " & mSiteCode & " ORDER BY PARTY,V_DATE,CASE WHEN CREDIT>0 THEN 3 WHEN DEBIT>0 THEN 2 ELSE 1 END,DOCID,V_SNo")
            End If
        End If
End Select
If Rst1.RecordCount > 0 Then Rst1.MoveFirst
Do Until Rst1.EOF
    If Val(TEXT(DrBalance)) > 0 Then
        If PubBackEnd = "S" Then
            Set RstDrX = G_FaCn.Execute("SELECT ISNULL(SUM(DEBIT),0)-ISNULL(SUM(CREDIT),0) AS BALANCE FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE<=" & FaConvertDate(TXTE_DATE))
        ElseIf PubBackEnd = "A" Then
            Set RstDrX = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT))-IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) AS BALANCE FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE<=" & FaConvertDate(TXTE_DATE))
        End If
        If RstDrX.RecordCount > 0 Then
            Select Case Trim(TEXT(8))
                Case "="
                    If RstDrX!Balance <> Val(TEXT(DrBalance)) Then GoTo EXIT_LOOP
                Case "<"
                    If RstDrX!Balance >= Val(TEXT(DrBalance)) Then GoTo EXIT_LOOP
                Case "<="
                    If RstDrX!Balance > Val(TEXT(DrBalance)) Then GoTo EXIT_LOOP
                Case ">"
                    If RstDrX!Balance <= Val(TEXT(DrBalance)) Then GoTo EXIT_LOOP
                Case ">="
                    If RstDrX!Balance < Val(TEXT(DrBalance)) Then GoTo EXIT_LOOP
                Case "<>"
                    If RstDrX!Balance = Val(TEXT(DrBalance)) Then GoTo EXIT_LOOP
            End Select
        End If
    End If
    If Val(TEXT(CrBalance)) > 0 Then
        If PubBackEnd = "S" Then
            Set RstDrX = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) AS BALANCE FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE<=" & FaConvertDate(TXTE_DATE))
        ElseIf PubBackEnd = "A" Then
            Set RstDrX = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) AS BALANCE FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE<=" & FaConvertDate(TXTE_DATE))
        End If
        If RstDrX.RecordCount > 0 Then
            Select Case Trim(TEXT(9))
                Case "="
                    If RstDrX!Balance <> Val(TEXT(CrBalance)) Then GoTo EXIT_LOOP
                Case "<"
                    If RstDrX!Balance >= Val(TEXT(CrBalance)) Then GoTo EXIT_LOOP
                Case "<="
                    If RstDrX!Balance > Val(TEXT(CrBalance)) Then GoTo EXIT_LOOP
                Case ">"
                    If RstDrX!Balance <= Val(TEXT(CrBalance)) Then GoTo EXIT_LOOP
                Case ">="
                    If RstDrX!Balance < Val(TEXT(CrBalance)) Then GoTo EXIT_LOOP
                Case "<>"
                    If RstDrX!Balance = Val(TEXT(CrBalance)) Then GoTo EXIT_LOOP
            End Select
        End If
    End If
    oBAL = 0
    Ac_Code = IIf(mReportStype = "Merge", 0, Rst1!SubCode)
    If mReportStype <> "Merge" Then Ac_Name = Trim(Rst1!Name)
    Ac_Name = left(Ac_Name, 50)
    Select Case mType
        Case "LedDeb"
            If Rst1!GroupNature <> "E" And Rst1!GroupNature <> "R" Then
                If PubBackEnd = "S" Then
                    oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE <=" & FaConvertDate(TXTE_DATE) & " AND CREDIT >0 " & mSiteCode & "").Fields(0)
                    oBAL = oBAL - G_FaCn.Execute("SELECT ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND DEBIT>0 " & mSiteCode & "").Fields(0)
                ElseIf PubBackEnd = "A" Then
                    oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE <=" & FaConvertDate(TXTE_DATE) & " AND CREDIT >0 " & mSiteCode & "").Fields(0)
                    oBAL = oBAL - G_FaCn.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND DEBIT>0 " & mSiteCode & "").Fields(0)
                End If
            Else
                If PubBackEnd = "S" Then
                    oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTE_DATE) & " AND CREDIT >0 " & mSiteCode & "").Fields(0)
                    oBAL = oBAL - G_FaCn.Execute("SELECT ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND DEBIT>0 " & mSiteCode & "").Fields(0)
                ElseIf PubBackEnd = "A" Then
                    oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTE_DATE) & " AND CREDIT >0 " & mSiteCode & "").Fields(0)
                    oBAL = oBAL - G_FaCn.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND DEBIT>0 " & mSiteCode & "").Fields(0)
                End If
            End If
            If oBAL < 0 Then
                With TmpRst
                    .AddNew
                    !V_tYPE = ""
                    !V_NO = 0
                    !V_ADD = ""
                    !V_SNo = 0
                    !Name = "Opening Balance"
                    !V_DATE = Format(TXTS_DATE, "dd/MMM/yyyy")
                    !PDATE = Format(TXTS_DATE, "dd/MMM/yyyy")
                    !cr = Abs(oBAL)
                    !ADJQTY = Abs(oBAL)
                    !Val = "0"
                    !Name1 = Ac_Name
                    !SubCode = Ac_Code
                    !CITY_NAME = FaXNull(Rst1!CITY_NAME)
                    !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                    .Update
                End With
                oBAL = 0
            End If
            Set Rst2 = G_FaCn.Execute("SELECT VIEWLEDGER.*,SUBGROUP.NAME FROM VIEWLEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=VIEWLEDGER.PARTY1 WHERE VIEWLEDGER.PARTY=" & FaChk_Text(Rst1!SubCode) & " AND VIEWLEDGER.V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND VIEWLEDGER.DEBIT>0  " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO")
            Do Until Rst2.EOF
                If oBAL >= Rst2!DEBIT Then
                    oBAL = oBAL - Rst2!DEBIT
                Else
                    NarrStr = ""
                    mNarrStr1 = ""
                    mNarrStr2 = ""
                    mDocNo = ""
                    If Trim(FaXNull(Rst2!Chq_No)) <> "" Then NarrStr = NarrStr + "Chq.No:" + Trim(FaXNull(Rst2!Chq_No))
                    If Not IsNull(Rst2!Chq_Date) Then NarrStr = NarrStr + " Dt: " + CStr(Format(Rst2!Chq_Date, "dd/MMM/yyyy"))
                    mNarrStr1 = IIf(Trim(FaXNull(Rst2!mNarr)) <> "", Trim(FaXNull(Rst2!mNarr)), Trim(FaXNull(Rst2!Narr)))
                    mNarrStr2 = IIf(Trim(FaXNull(Rst2!mNarr)) <> "", Trim(Rst2!Narr), "")
                    If PubFaSiteType = 1 Then
                        mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(Rst2!DocId, 1), "")
                        mDocNo = mDocNo + IIf(RstEnviro!LedSiteCode = "Yes", Trim(Right(Rst2!Site_Code, 1)), "")
                        mDocNo = mDocNo + IIf(RstEnviro!LedPrefix = "Yes", IIf(mDocNo = "", "", "/") + Trim(Rst2!V_ADD), "")
                        mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + left(Trim(Rst2!V_tYPE), 1) + Trim(mID(Trim(Rst2!V_tYPE), 3, 3))
                        mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(Rst2!V_NO))
                    Else
                        mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(Rst2!DocId, 1), "") + IIf(RstEnviro!LedSiteCode = "Yes", Trim(left(Rst2!Site_Code, 1)), "")
                        mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + IIf(RstEnviro!LedPrefix = "Yes", Trim(Rst2!V_ADD), "")
                        mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(Rst2!V_tYPE)
                        mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(Rst2!V_NO))
                    End If
                    With TmpRst
                        .AddNew
                        !V_ADD = FaXNull(Rst2!V_ADD)
                        !V_NO = Rst2!V_NO
                        .Fields("DocNo") = mDocNo
                        !Name = Trim(FaXNull(Rst2!Name))
                        !V_DATE = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                        !cr = Rst2!DEBIT
                        !ADJQTY = Rst2!DEBIT - oBAL
                        !V_tYPE = Rst2!V_tYPE
                        !V_SNo = Rst2!V_SNo
                        !Name1 = Ac_Name
                        !SubCode = Ac_Code
                        !CITY_NAME = FaXNull(Rst1!CITY_NAME)
                        !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                        .Fields("DocId") = Rst2!DocId
                        .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                        If Rst2!V_tYPE = "OP" Then
                            !Val = "1"
                        Else
                            !Val = IIf(Rst2!CREDIT > 0, "3", "2")
                        End If
                        .Update
                    End With
                    If TEXT(AsPerDetail) = "Yes" And Len(Rst2!Party1) = 0 Then
                        With TmpRst
                            .MoveLast
                            .Fields("Name") = "As Per Detail"
                            .Fields("DocId") = Rst2!DocId
                            .Fields("Name1") = Ac_Name
                            .Fields("SubCode") = Ac_Code
                            .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                            .Update
                        End With
                        Set RstCheck21 = G_FaCn.Execute("SELECT SUBGROUP.SUBCODE,SUBGROUP.NAME,VIEWLEDGER.* FROM VIEWLEDGER LEFT JOIN SUBGROUP ON VIEWLEDGER.PARTY=SUBGROUP.SUBCODE WHERE DOCID=" & FaChk_Text(Rst2!DocId) & " AND V_SNo<>" & Rst2!V_SNo)
                        Do Until RstCheck21.EOF
                            DNarrStr = ""
                            DmNarrStr1 = ""
                            DmNarrStr2 = ""
                            If FaXNull(RstCheck21!Chq_No) <> "" Then DNarrStr = DNarrStr + "Chq.No:" + Trim(FaXNull(RstCheck21!Chq_No))
                            If Not IsNull(RstCheck21!Chq_Date) Then DNarrStr = DNarrStr + " Dt: " + CStr(Format(RstCheck21!Chq_Date, "dd/MM/yy"))
                            DmNarrStr2 = Trim(FaXNull(RstCheck21!Narr))
                            With TmpRst
                                .AddNew
                                If Rst2!V_tYPE = "OP" Then
                                    !Val = "1"
                                Else
                                    !Val = IIf(Rst2!CREDIT > 0, "3", "2")
                                End If
                                .Fields("PDate") = Format(RstCheck21!V_DATE, "dd/MMM/yyyy")
                                .Fields("Name1") = Ac_Name
                                .Fields("DocId") = RstCheck21!DocId
                                If RstCheck21!DEBIT > 0 Then
                                    .Fields("Sub") = "*"
                                    .Fields("Name") = Space(2) + FaSetW(RstCheck21!Name, 20) + " " + FaSetN(FaSNull(RstCheck21!DEBIT), 12) + " Dr"
                                Else
                                    .Fields("Sub") = "*"
                                    .Fields("Name") = Space(2) + FaSetW(RstCheck21!Name, 20) + " " + FaSetN(FaSNull(RstCheck21!CREDIT), 12) + " Cr"
                                End If
                                .Update
                            End With
                            If TEXT(AsPerDetailNarration) = "Yes" Then
                                If Trim(DNarrStr) <> "" Then
                                    Do While Len(DNarrStr) > 0
                                        With TmpRst
                                            .AddNew
                                            .Fields("Name") = Space(2) + left(DNarrStr, 35)
                                            .Fields("DocId") = RstCheck21!DocId
                                            .Fields("Name1") = Ac_Name
                                            .Fields("Sub") = "*"
                                            .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                            If Rst2!V_tYPE = "OP" Then
                                                .Fields("Val") = "1"
                                            Else
                                                .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                            End If
                                            .Update
                                        End With
                                        DNarrStr = mID(DNarrStr, 36, 300)
                                    Loop
                                End If
                                If Trim(DmNarrStr2) <> "" Then
                                    Do While Len(DmNarrStr2) > 0
                                        With TmpRst
                                            .AddNew
                                            .Fields("Name") = Space(2) + left(DmNarrStr2, 35)
                                            .Fields("DocId") = RstCheck21!DocId
                                            .Fields("Name1") = Ac_Name
                                            .Fields("Sub") = "*"
                                            .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                            If Rst2!V_tYPE = "OP" Then
                                                .Fields("Val") = "1"
                                            Else
                                                .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                            End If
                                            .Update
                                        End With
                                        DmNarrStr2 = mID(DmNarrStr2, 36, 300)
                                    Loop
                                End If
                            End If
                            RstCheck21.MoveNext
                        Loop
                    End If
                    If Trim(NarrStr) <> "" Then
                        Do While Len(NarrStr) > 0
                            With TmpRst
                                .MoveLast
                                If Trim(TmpRst.Fields("Name")) <> "" Then .AddNew
                                .Fields("Name") = Space(2) + left(NarrStr, 48)
                                .Fields("DocId") = Rst2!DocId
                                .Fields("Name1") = Ac_Name
                                .Fields("SubCode") = Ac_Code
                                .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                If Rst2!V_tYPE = "OP" Then
                                    .Fields("Val") = "1"
                                Else
                                    .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                End If
                                .Update
                            End With
                            NarrStr = mID(NarrStr, 49, 300)
                        Loop
                    End If
                    If Trim(mNarrStr1) <> "" Then
                        Do While Len(mNarrStr1) > 0
                            With TmpRst
                                .MoveLast
                                If Trim(TmpRst.Fields("Name")) <> "" Then .AddNew
                                .Fields("Name") = Space(2) + left(mNarrStr1, 48)
                                .Fields("SubCode") = Ac_Code
                                .Fields("Name1") = Ac_Name
                                .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                If Rst2!V_tYPE = "OP" Then
                                    .Fields("Val") = "1"
                                Else
                                    .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                End If
                                .Fields("DocId") = Rst2!DocId
                                .Update
                            End With
                            mNarrStr1 = mID(mNarrStr1, 49, 300)
                        Loop
                    End If
                    If Trim(mNarrStr2) <> "" Then
                        Do While Len(mNarrStr2) > 0
                            With TmpRst
                                .MoveLast
                                If Trim(TmpRst.Fields("Name")) <> "" Then .AddNew
                                .Fields("Name") = Space(2) + left(mNarrStr2, 48)
                                .Fields("DocId") = Rst2!DocId
                                .Fields("Name1") = Ac_Name
                                .Fields("SubCode") = Ac_Code
                                .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                If Rst2!V_tYPE = "OP" Then
                                    .Fields("Val") = "1"
                                Else
                                    .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                End If
                                .Update
                            End With
                            mNarrStr2 = mID(mNarrStr2, 49, 300)
                        Loop
                    End If
                    oBAL = 0
                End If
                Rst2.MoveNext
            Loop
            If oBAL > 0 Then
                With TmpRst
                    .AddNew
                    !V_tYPE = ""
                    !V_NO = 0
                    !V_ADD = ""
                    !V_SNo = 0
                    !Name = "Excess Credit"
                    !V_DATE = TXTE_DATE
                    !ADJAMT = Abs(oBAL)
                    !Name1 = Ac_Name
                    !SubCode = Ac_Code
                    !Address1 = Trim(IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), "", ", " + Rst1!Add2)))
                    !CITY_NAME = FaXNull(Rst1!CITY_NAME)
                    .Update
                End With
            End If
        Case "LedCred"
            If Rst1!GroupNature <> "E" And Rst1!GroupNature <> "R" Then
                If PubBackEnd = "S" Then
                    oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE <=" & FaConvertDate(TXTE_DATE) & " AND DEBIT >0 " & mSiteCode & "").Fields(0)
                    oBAL = oBAL - G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND CREDIT>0 " & mSiteCode & "").Fields(0)
                ElseIf PubBackEnd = "A" Then
                    oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE <=" & FaConvertDate(TXTE_DATE) & " AND DEBIT >0 " & mSiteCode & "").Fields(0)
                    oBAL = oBAL - G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND CREDIT>0 " & mSiteCode & "").Fields(0)
                End If
            Else
                If PubBackEnd = "S" Then
                    oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(TXTE_DATE) & " AND DEBIT >0 " & mSiteCode & "").Fields(0)
                    oBAL = oBAL - G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND CREDIT>0 " & mSiteCode & "").Fields(0)
                ElseIf PubBackEnd = "A" Then
                    oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(TXTE_DATE) & " AND DEBIT >0 " & mSiteCode & "").Fields(0)
                    oBAL = oBAL - G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT)) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " AND CREDIT>0 " & mSiteCode & "").Fields(0)
                End If
            End If
            If oBAL < 0 Then
                With TmpRst
                    .AddNew
                    !V_tYPE = ""
                    !V_NO = 0
                    !V_ADD = ""
                    !V_SNo = 0
                    !Name = "Opening Balance"
                    !V_DATE = Format(TXTS_DATE, "dd/MMM/yyyy")
                    !PDATE = TXTS_DATE
                    !cr = Abs(oBAL)
                    !ADJQTY = Abs(oBAL)
                    !Val = "0"
                    !Name1 = Ac_Name
                    !SubCode = Ac_Code
                    !Address1 = Trim(IIf(Trim(FaXNull(Rst1!Add1)) = "", "", Rst1!Add1) + Trim(IIf(Trim(FaXNull(Rst1!Add2)) = "", "", ", " + Rst1!Add2)))
                    !CITY_NAME = FaXNull(Rst1!CITY_NAME)
                    .Update
                End With
                oBAL = 0
            End If
            Set Rst2 = G_FaCn.Execute("SELECT VIEWLEDGER.*,SUBGROUP.NAME FROM VIEWLEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=VIEWLEDGER.PARTY1 WHERE VIEWLEDGER.PARTY=" & FaChk_Text(Rst1!SubCode) & " AND VIEWLEDGER.V_DATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " AND VIEWLEDGER.CREDIT>0  " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO")
            Do Until Rst2.EOF
                If oBAL >= Rst2!CREDIT Then
                    oBAL = oBAL - Rst2!CREDIT
                Else
                    NarrStr = ""
                    mNarrStr1 = ""
                    mNarrStr2 = ""
                    mDocNo = ""
                    If FaXNull(Rst2!Chq_No) <> "" Then NarrStr = NarrStr + "Chq.No:" + Trim(FaXNull(Rst2!Chq_No))
                    If Not IsNull(Rst2!Chq_Date) Then NarrStr = NarrStr + " Dt: " + CStr(Format(Rst2!Chq_Date, "dd/MM/yy"))
                    mNarrStr1 = IIf(FaXNull(Rst2!mNarr) <> "", Trim(FaXNull(Rst2!mNarr)), Trim(FaXNull(Rst2!Narr)))
                    mNarrStr2 = IIf(FaXNull(Rst2!mNarr) <> "", Trim(Rst2!Narr), "")
                    If PubFaSiteType = 1 Then
                        mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(Rst2!DocId, 1), "")
                        mDocNo = mDocNo + IIf(RstEnviro!LedSiteCode = "Yes", Trim(Right(Rst2!Site_Code, 1)), "")
                        mDocNo = mDocNo + IIf(RstEnviro!LedPrefix = "Yes", IIf(mDocNo = "", "", "/") + Trim(Rst2!V_ADD), "")
                        mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + left(Trim(Rst2!V_tYPE), 1) + Trim(mID(Trim(Rst2!V_tYPE), 3, 3))
                        mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(Rst2!V_NO))
                    Else
                        mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(Rst2!DocId, 1), "") + IIf(RstEnviro!LedSiteCode = "Yes", Trim(left(Rst2!Site_Code, 1)), "")
                        mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + IIf(RstEnviro!LedPrefix = "Yes", Trim(Rst2!V_ADD), "")
                        mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(Rst2!V_tYPE)
                        mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(Rst2!V_NO))
                    End If
                    With TmpRst
                        .AddNew
                        !V_ADD = FaXNull(Rst2!V_ADD)
                        !V_NO = Rst2!V_NO
                        .Fields("DocNo") = mDocNo
                        !Name = Trim(FaXNull(Rst2!Name))
                        !V_DATE = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                        !cr = Rst2!CREDIT
                        !ADJQTY = Rst2!CREDIT - oBAL
                        !V_tYPE = Rst2!V_tYPE
                        !V_SNo = Rst2!V_SNo
                        !Name1 = Ac_Name
                        !SubCode = Ac_Code
                        !CITY_NAME = FaXNull(Rst1!CITY_NAME)
                        !Address1 = Trim(IIf(Trim(FaXNull(Rst1!Add1)) = "", "", Rst1!Add1) + Trim(IIf(Trim(FaXNull(Rst1!Add2)) = "", "", ", " + Rst1!Add2)))
                        .Fields("DocId") = Rst2!DocId
                        .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                        If Rst2!V_tYPE = "OP" Then
                            !Val = "0"
                        Else
                            !Val = IIf(Rst2!CREDIT > 0, "3", "2")
                        End If
                        .Update
                    End With
                    If TEXT(AsPerDetail) = "Yes" And Len(Rst2!Party1) = 0 Then
                        With TmpRst
                            .MoveLast
                            .Fields("Name") = "As Per Detail"
                            .Fields("DocId") = Rst2!DocId
                            .Fields("Name1") = Ac_Name
                            .Fields("SubCode") = Ac_Code
                            .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                            .Update
                        End With
                        Set RstCheck21 = G_FaCn.Execute("SELECT SUBGROUP.SUBCODE,SUBGROUP.NAME,VIEWLEDGER.* FROM VIEWLEDGER LEFT JOIN SUBGROUP ON VIEWLEDGER.PARTY=SUBGROUP.SUBCODE WHERE DOCID=" & FaChk_Text(Rst2!DocId) & " AND V_SNo<>" & Rst2!V_SNo)
                        Do Until RstCheck21.EOF
                            DNarrStr = ""
                            DmNarrStr1 = ""
                            DmNarrStr2 = ""
                            If Trim(FaXNull(RstCheck21!Chq_No)) <> "" Then DNarrStr = DNarrStr + "Chq.No:" + Trim(FaXNull(RstCheck21!Chq_No))
                            If Not IsNull(RstCheck21!Chq_Date) Then DNarrStr = DNarrStr + " Dt: " + CStr(Format(RstCheck21!Chq_Date, "dd/MM/yy"))
                            DmNarrStr2 = Trim(FaXNull(RstCheck21!Narr))
                            With TmpRst
                                .AddNew
                                If Rst2!V_tYPE = "OP" Then
                                    !Val = "1"
                                Else
                                    !Val = IIf(Rst2!CREDIT > 0, "3", "2")
                                End If
                                .Fields("PDate") = Format(RstCheck21!V_DATE, "dd/MMM/yyyy")
                                .Fields("Name1") = Ac_Name
                                .Fields("DocId") = RstCheck21!DocId
                                If RstCheck21!DEBIT > 0 Then
                                    .Fields("Sub") = "*"
                                    .Fields("Name") = Space(2) + FaSetW(RstCheck21!Name, 20) + " " + FaSetN(FaSNull(RstCheck21!DEBIT), 12) + " Dr"
                                Else
                                    .Fields("Sub") = "*"
                                    .Fields("Name") = Space(2) + FaSetW(RstCheck21!Name, 20) + " " + FaSetN(FaSNull(RstCheck21!CREDIT), 12) + " Cr"
                                End If
                                .Update
                            End With
                            If TEXT(AsPerDetailNarration) = "Yes" Then
                                If Trim(DNarrStr) <> "" Then
                                    Do While Len(DNarrStr) > 0
                                        With TmpRst
                                            .AddNew
                                            .Fields("Name") = Space(2) + left(DNarrStr, 35)
                                            .Fields("DocId") = RstCheck21!DocId
                                            .Fields("Name1") = Ac_Name
                                            .Fields("Sub") = "*"
                                            .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                            If Rst2!V_tYPE = "OP" Then
                                                .Fields("Val") = "1"
                                            Else
                                                .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                            End If
                                            .Update
                                        End With
                                        DNarrStr = mID(DNarrStr, 36, 300)
                                    Loop
                                End If
                                If Trim(DmNarrStr2) <> "" Then
                                    Do While Len(DmNarrStr2) > 0
                                        With TmpRst
                                            .AddNew
                                            .Fields("Name") = Space(2) + left(DmNarrStr2, 35)
                                            .Fields("DocId") = RstCheck21!DocId
                                            .Fields("Name1") = Ac_Name
                                            .Fields("Sub") = "*"
                                            .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                            If Rst2!V_tYPE = "OP" Then
                                                .Fields("Val") = "1"
                                            Else
                                                .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                            End If
                                            .Update
                                        End With
                                        DmNarrStr2 = mID(DmNarrStr2, 36, 300)
                                    Loop
                                End If
                            End If
                            RstCheck21.MoveNext
                        Loop
                    End If
                    If Trim(NarrStr) <> "" Then
                        Do While Len(NarrStr) > 0
                            With TmpRst
                                .MoveLast
                                If Trim(TmpRst.Fields("Name")) <> "" Then .AddNew
                                .Fields("Name") = Space(2) + left(NarrStr, 48)
                                .Fields("DocId") = Rst2!DocId
                                .Fields("Name1") = Ac_Name
                                .Fields("SubCode") = Ac_Code
                                .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                If Rst2!V_tYPE = "OP" Then
                                    .Fields("Val") = "1"
                                Else
                                    .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                End If
                                .Update
                            End With
                            NarrStr = mID(NarrStr, 49, 300)
                        Loop
                    End If
                    If Trim(mNarrStr1) <> "" Then
                        Do While Len(mNarrStr1) > 0
                            With TmpRst
                                .MoveLast
                                If Trim(TmpRst.Fields("Name")) <> "" Then .AddNew
                                .Fields("Name") = Space(2) + left(mNarrStr1, 48)
                                .Fields("DocId") = Rst2!DocId
                                .Fields("Name1") = Ac_Name
                                .Fields("SubCode") = Ac_Code
                                .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                If Rst2!V_tYPE = "OP" Then
                                    .Fields("Val") = "1"
                                Else
                                    .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                End If
                                .Update
                            End With
                            mNarrStr1 = mID(mNarrStr1, 49, 300)
                        Loop
                    End If
                    If Trim(mNarrStr2) <> "" Then
                        Do While Len(mNarrStr2) > 0
                            With TmpRst
                                .MoveLast
                                If Trim(TmpRst.Fields("Name")) <> "" Then .AddNew
                                .Fields("Name") = Space(2) + left(mNarrStr2, 48)
                                .Fields("DocId") = Rst2!DocId
                                .Fields("Name1") = Ac_Name
                                .Fields("SubCode") = Ac_Code
                                .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                If Rst2!V_tYPE = "OP" Then
                                    .Fields("Val") = "1"
                                Else
                                    .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                End If
                                .Update
                            End With
                            mNarrStr2 = mID(mNarrStr2, 49, 300)
                        Loop
                    End If
                    oBAL = 0
                End If
                Rst2.MoveNext
            Loop
            If oBAL > 0 Then
                With TmpRst
                    .AddNew
                    !V_tYPE = ""
                    !V_NO = 0
                    !V_ADD = ""
                    !V_SNo = 0
                    !Name = "Excess Debit"
                    !V_DATE = Format(TXTE_DATE, "dd/MMM/yyyy")
                    !ADJAMT = Abs(oBAL)
                    !Name1 = Ac_Name
                    !SubCode = Ac_Code
                    !Address1 = Trim(IIf(Trim(FaXNull(Rst1!Add1)) = "", "", Rst1!Add1) + Trim(IIf(Trim(FaXNull(Rst1!Add2)) = "", "", ", " + Rst1!Add2)))
                    !CITY_NAME = FaXNull(Rst1!CITY_NAME)
                    .Update
                End With
            End If
        Case Else
'            Set Rst2 = G_FaCn.Execute("SELECT VIEWLEDGER.*,SUBGROUP.NAME FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY1=subgroup.SUBCODE WHERE VIEWLEDGER.PARTY=" & FaChk_Text(RST1!SubCode) & " AND VIEWLEDGER.V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " " & Qry & " ORDER BY V_DATE,SUBGROUP.NAME,CREDIT,DEBIT")
            If mType <> "AcCheckList" Then
                If Rst1!GroupNature <> "E" And Rst1!GroupNature <> "R" Then
                    If PubBackEnd = "S" Then
                        oBAL = oBAL + G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & "  " & mSiteCode & "").Fields(0)
                    ElseIf PubBackEnd = "A" Then
                        oBAL = oBAL + G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & "  " & mSiteCode & "").Fields(0)
                    End If
                Else
                    If PubBackEnd = "S" Then
                        oBAL = oBAL + G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & "  " & mSiteCode & "").Fields(0)
                    ElseIf PubBackEnd = "A" Then
                        oBAL = oBAL + G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY=" & FaChk_Text(Rst1!SubCode) & " AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & "  " & mSiteCode & "").Fields(0)
                    End If
                End If
                If mType = "LedInt" Then
                    If Rst2.RecordCount > 0 Then
                        Rst2.MoveFirst
                        Rst2.FIND "Party='" & Rst1!SubCode & "'"
                        If Rst2.EOF = False Then
                            If oBAL = 0 Then Rst2.MoveNext
                            If Rst2.EOF = True Then
                                NextDate = Format(TXTE_DATE, "dd/MMM/yyyy")
                            Else
                                If Rst2!Party = Rst1!SubCode Then
                                    NextDate = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                Else
                                    NextDate = Format(TXTE_DATE, "dd/MMM/yyyy")
                                End If
                            End If
                            If oBAL = 0 Then Rst2.MovePrevious
                        Else
                            NextDate = Format(TXTE_DATE, "dd/MMM/yyyy")
                        End If
                    Else
                        NextDate = Format(TXTE_DATE, "dd/MMM/yyyy")
                    End If
                End If
                If oBAL <> 0 Then
                    With TmpRst
                        .AddNew
                        !V_tYPE = ""
                        !V_NO = 0
                        !V_ADD = ""
                        !V_SNo = 0
                        !Name = "OPENING BALANCE"
                        !V_DATE = Format(TXTS_DATE, "dd/MMM/yyyy")
                        !cr = IIf(oBAL > 0, Abs(oBAL), 0)
                        !ADJAMT = IIf(oBAL < 0, Abs(oBAL), 0)
                        !Val = "1"
                        !Name1 = Ac_Name
                        !Address1 = Trim(IIf(IsNull(Rst1!Add1), " ", Rst1!Add1) + Trim(IIf(IsNull(Rst1!Add2), " ", ", " + Rst1!Add2)))
                        !CITY_NAME = FaXNull(Rst1!CITY_NAME)
                        !SubCode = Ac_Code
                        .Fields("PDate") = Format(TXTS_DATE, "dd/MMM/yyyy")
                        If mType = "LedInt" Then .Fields("NextDate") = Format(NextDate, "dd/MMM/yyyy")
                        .Update
                    End With
                End If
            End If
            If Rst2.RecordCount > 0 Then
                Rst2.MoveFirst
                Rst2.FIND "Party='" & Rst1!SubCode & "'"
                If Rst2.EOF = False Then
                    Do While True
                        If mType = "LedInt" And Rst1!SubCode = Rst2!Party Then
                            If Rst2.RecordCount > 0 Then
                                Rst2.MoveNext
                                If Rst2.EOF = True Then
                                    NextDate = Format(TXTE_DATE, "dd/MMM/yyyy")
                                Else
                                    If Rst2!Party = Rst1!SubCode Then
                                        NextDate = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                    Else
                                        NextDate = Format(TXTE_DATE, "dd/MMM/yyyy")
                                    End If
                                End If
                                Rst2.MovePrevious
                            End If
                        End If
                        With TmpRst
                            .AddNew
                            !V_ADD = FaXNull(Rst2!V_ADD)
                            !V_NO = Rst2!V_NO
                            NarrStr = ""
                            mNarrStr1 = ""
                            mNarrStr2 = ""
                            mDocNo = ""
                            If Trim(FaXNull(Rst2!Chq_No)) <> "" Then NarrStr = NarrStr + "Chq.No:" + Trim(FaXNull(Rst2!Chq_No))
                            If Not IsNull(Rst2!Chq_Date) Then NarrStr = NarrStr + " Dt: " + CStr(Format(Rst2!Chq_Date, "dd/MM/yy"))
                            mNarrStr1 = IIf(FaXNull(Rst2!mNarr) <> "", Trim(FaXNull(Rst2!mNarr)), Trim(FaXNull(Rst2!Narr)))
                            mNarrStr2 = IIf(FaXNull(Rst2!mNarr) <> "", Trim(FaXNull(Rst2!Narr)), "")
                            If PubFaSiteType = 1 Then
                                mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(Rst2!DocId, 1), "")
                                mDocNo = mDocNo + IIf(RstEnviro!LedSiteCode = "Yes", Trim(Right(Rst2!Site_Code, 1)), "")
                                mDocNo = mDocNo + IIf(RstEnviro!LedPrefix = "Yes", IIf(mDocNo = "", "", "/") + Trim(Rst2!V_ADD), "")
                                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + left(Trim(Rst2!V_tYPE), 1) + Trim(mID(Trim(Rst2!V_tYPE), 3, 3))
                                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(Rst2!V_NO))
                            Else
                                mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(Rst2!DocId, 1), "") + IIf(RstEnviro!LedSiteCode = "Yes", Trim(left(Rst2!Site_Code, 1)), "")
                                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + IIf(RstEnviro!LedPrefix = "Yes", Trim(Rst2!V_ADD), "")
                                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(Rst2!V_tYPE)
                                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(Rst2!V_NO))
                            End If
                            .Fields("DocNo") = mDocNo
                            !Name = Trim(FaXNull(Rst2!Name))
                            !V_DATE = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                            !cr = Rst2!CREDIT
                            !ADJAMT = Rst2!DEBIT
                            !V_tYPE = Rst2!V_tYPE
                            !V_SNo = Rst2!V_SNo
                            If Rst2!V_tYPE = "OP" Then
                                !Val = "1"
                            Else
                                !Val = IIf(Rst2!CREDIT > 0, "3", "2")
                            End If
                            !Name1 = Ac_Name
                            !SubCode = Ac_Code
                            !Address1 = Trim(IIf(Trim(FaXNull(Rst1!Add1)) = "", "", Rst1!Add1) + Trim(IIf(Trim(FaXNull(Rst1!Add2)) = "", "", ", " + Rst1!Add2)))
                            !CITY_NAME = FaXNull(Rst1!CITY_NAME)
                            .Fields("DocId") = Rst2!DocId
                            .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                            If mType = "LedInt" Then .Fields("NextDate") = Format(NextDate, "dd/MMM/yyyy")
                            .Update
                        End With
                        If TEXT(AsPerDetail) = "Yes" And Len(Rst2!Party1) = 0 Then
                            With TmpRst
                                .MoveLast
                                .Fields("Name") = "As Per Detail"
                                .Fields("DocId") = Rst2!DocId
                                .Fields("Name1") = Ac_Name
                                .Fields("SubCode") = Ac_Code
                                .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                .Update
                            End With
                            Set RstCheck21 = G_FaCn.Execute("SELECT SUBGROUP.SUBCODE,SUBGROUP.NAME,VIEWLEDGER.* FROM VIEWLEDGER LEFT JOIN SUBGROUP ON VIEWLEDGER.PARTY=SUBGROUP.SUBCODE WHERE DOCID=" & FaChk_Text(Rst2!DocId) & " AND V_SNo<>" & Rst2!V_SNo)
                            Do Until RstCheck21.EOF
                                DNarrStr = ""
                                DmNarrStr1 = ""
                                DmNarrStr2 = ""
                                If FaXNull(RstCheck21!Chq_No) <> "" Then DNarrStr = DNarrStr + "Chq.No:" + Trim(FaXNull(RstCheck21!Chq_No))
                                If Not IsNull(RstCheck21!Chq_Date) Then DNarrStr = DNarrStr + " Dt: " + CStr(Format(RstCheck21!Chq_Date, "dd/MM/yy"))
                                DmNarrStr2 = Trim(FaXNull(RstCheck21!Narr))
                                With TmpRst
                                    .AddNew
                                    If Rst2!V_tYPE = "OP" Then
                                        !Val = "1"
                                    Else
                                        !Val = IIf(Rst2!CREDIT > 0, "3", "2")
                                    End If
                                    .Fields("PDate") = Format(RstCheck21!V_DATE, "dd/MMM/yyyy")
                                    .Fields("Name1") = Ac_Name
                                    .Fields("DocId") = RstCheck21!DocId
                                    If RstCheck21!DEBIT > 0 Then
                                        .Fields("Sub") = "*"
                                        .Fields("Name") = Space(2) + FaSetW(RstCheck21!Name, 20) + " " + FaSetN(FaSNull(RstCheck21!DEBIT), 12) + " Dr"
                                    Else
                                        .Fields("Sub") = "*"
                                        .Fields("Name") = Space(2) + FaSetW(RstCheck21!Name, 20) + " " + FaSetN(FaSNull(RstCheck21!CREDIT), 12) + " Cr"
                                    End If
                                    .Update
                                End With
                                If TEXT(AsPerDetailNarration) = "Yes" Then
                                    If Trim(DNarrStr) <> "" Then
                                        Do While Len(DNarrStr) > 0
                                            With TmpRst
                                                .AddNew
                                                .Fields("Name") = Space(2) + left(DNarrStr, 35)
                                                .Fields("DocId") = RstCheck21!DocId
                                                .Fields("Name1") = Ac_Name
                                                .Fields("Sub") = "*"
                                                .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                                If Rst2!V_tYPE = "OP" Then
                                                    .Fields("Val") = "1"
                                                Else
                                                    .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                                End If
                                                .Update
                                            End With
                                            DNarrStr = mID(DNarrStr, 36, 300)
                                        Loop
                                    End If
                                    If Trim(DmNarrStr2) <> "" Then
                                        Do While Len(DmNarrStr2) > 0
                                            With TmpRst
                                                .AddNew
                                                .Fields("Name") = Space(2) + left(DmNarrStr2, 35)
                                                .Fields("DocId") = RstCheck21!DocId
                                                .Fields("Name1") = Ac_Name
                                                .Fields("Sub") = "*"
                                                .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                                If Rst2!V_tYPE = "OP" Then
                                                    .Fields("Val") = "1"
                                                Else
                                                    .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                                End If
                                                .Update
                                            End With
                                            DmNarrStr2 = mID(DmNarrStr2, 36, 300)
                                        Loop
                                    End If
                                End If
                                RstCheck21.MoveNext
                            Loop
                        End If
                        If TEXT(PrintNarration) = "Yes" Then
                            If Trim(NarrStr) <> "" Then
                                Do While Len(NarrStr) > 0
                                    With TmpRst
                                        .MoveLast
                                        If Trim(TmpRst.Fields("Name")) <> "" Then .AddNew
                                        .Fields("Name") = Space(2) + left(NarrStr, 48)
                                        .Fields("DocId") = Rst2!DocId
                                        .Fields("Name1") = Ac_Name
                                        .Fields("SubCode") = Ac_Code
                                        .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                        If Rst2!V_tYPE = "OP" Then
                                            .Fields("Val") = "1"
                                        Else
                                            .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                        End If
                                        .Update
                                    End With
                                    NarrStr = mID(NarrStr, 49, 300)
                                Loop
                            End If
                            If Trim(mNarrStr1) <> "" Then
                                Do While Len(mNarrStr1) > 0
                                    With TmpRst
                                        .MoveLast
                                        If Trim(TmpRst.Fields("Name")) <> "" Then .AddNew
                                        .Fields("Name") = Space(2) + left(mNarrStr1, 48)
                                        .Fields("DocId") = Rst2!DocId
                                        .Fields("Name1") = Ac_Name
                                        .Fields("SubCode") = Ac_Code
                                        .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                        If Rst2!V_tYPE = "OP" Then
                                            .Fields("Val") = "1"
                                        Else
                                            .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                        End If
                                        .Update
                                    End With
                                    mNarrStr1 = mID(mNarrStr1, 49, 300)
                                Loop
                            End If
                            If Trim(mNarrStr2) <> "" Then
                                Do While Len(mNarrStr2) > 0
                                    With TmpRst
                                        .MoveLast
                                        If Trim(TmpRst.Fields("Name")) <> "" Then .AddNew
                                        .Fields("Name") = Space(2) + left(mNarrStr2, 48)
                                        .Fields("DocId") = Rst2!DocId
                                        .Fields("Name1") = Ac_Name
                                        .Fields("SubCode") = Ac_Code
                                        .Fields("PDate") = Format(Rst2!V_DATE, "dd/MMM/yyyy")
                                        If Rst2!V_tYPE = "OP" Then
                                            .Fields("Val") = "1"
                                        Else
                                            .Fields("Val") = IIf(Rst2!CREDIT > 0, "3", "2")
                                        End If
                                        .Update
                                    End With
                                    mNarrStr2 = mID(mNarrStr2, 49, 300)
                                Loop
                            End If
                        End If
                        Rst2.MoveNext
                        If Rst2.EOF Then Exit Do
                        If Rst2!Party <> Rst1!SubCode Then Exit Do
                    Loop
                End If
            End If
    End Select
EXIT_LOOP:
    Rst1.MoveNext
Loop
FrDetail.Visible = False
If TmpRst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: GoTo ExitLoop
X1 = CreateFieldDefFile(TmpRst, PubFaReportPath + "\FaLEDGER.ttx", True)
If mType = "LedDeb" And Index = 1 Then
    Set rpt = rdApp.OpenReport(PubFaReportPath + "\FaTAGADALET.RPT")
    G_FaCn.Execute "UPDATE FaEnviro SET TagadaHeader1=" & FaChk_Text(TxtHeader(0)) & ",TagadaHeader2=" & FaChk_Text(TxtHeader(1)) & ",TagadaHeader3=" & FaChk_Text(TxtHeader(2)) & ",TagadaHeader4=" & FaChk_Text(TxtHeader(3)) & ",TagadaHeader5=" & FaChk_Text(TxtHeader(4)) & ",TagadaFooter1=" & FaChk_Text(TxtFooter(0)) & ",TagadaFooter2=" & FaChk_Text(TxtFooter(1)) & ",TagadaFooter3=" & FaChk_Text(TxtFooter(2)) & ",TagadaFooter4=" & FaChk_Text(TxtFooter(3)) & ",TagadaFooter5=" & FaChk_Text(TxtFooter(4))
Else
    If MsgBox("Do You Want Each A/C on Separate Page", vbYesNo + vbQuestion + vbDefaultButton2, Me.CAPTION) = vbYes Then
        mSepratePage = "Y"
    Else
        mSepratePage = "N"
    End If
    Select Case Index
        Case 1
            Select Case mType
                Case "Led", "AcCheckList"
                    MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
                    TinTin = PubDatamanFa.FaLedDosPrinting(Me, TmpRst, mSepratePage, TEXT(PrintVrNo))
                    GoTo ExitLoop
                    Exit Sub
                Case "LedInt"
                    Set rpt = PubDatamanFa.FaINTERESTRpt
                Case "LedDeb", "LedCred"
                    Set rpt = rdApp.OpenReport(PubFaReportPath + "\FaTAGADADR.RPT")
            End Select
        Case Else
            Select Case mType
                Case "Led", "AcCheckList"
                    X1 = CreateFieldDefFile(TmpRst, PubFaReportPath + "\FaLedger.ttx", True)
                    Set rpt = PubDatamanFa.FaLedgerRpt
                Case "LedInt"
                    X1 = CreateFieldDefFile(TmpRst, PubFaReportPath + "\FaINTEREST.ttx", True)
                    Set rpt = PubDatamanFa.FaINTERESTRpt
                Case "LedDeb", "LedCred"
                    X1 = CreateFieldDefFile(TmpRst, PubFaReportPath + "\FaTAGADADR.ttx", True)
                    Set rpt = PubDatamanFa.FaTAGADADRRpt
            End Select
    End Select
End If
For I = 1 To rpt.FormulaFields.Count
    Select Case rpt.FormulaFields(I).FormulaFieldName
        Case "TITLE"
            If Trim(TEXT(SiteCode)) <> "" Then
                rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION + " (" + TEXT(SiteCode) + ")" + "'"
            Else
                rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
            End If
        Case "DT"
            rpt.FormulaFields(I).TEXT = "'From : " & TXTS_DATE & " To : " & TXTE_DATE & "'"
        Case "INT_RATE"
            If mType = "LedInt" Then rpt.FormulaFields(I).TEXT = "" & Val(TEXT(Interest)) & ""
        Case "END_DATE"
            If mType = "LedInt" Then rpt.FormulaFields(I).TEXT = "DATE(#" & Format(CDate(TXTE_DATE), "YYYY,MM,DD") & "#)"
        Case "SEPRATEPAGE"
            rpt.FormulaFields(I).TEXT = "'" & mSepratePage & "'"
        Case "PrintVrNo"
            rpt.FormulaFields(I).TEXT = "'" & TEXT(PrintVrNo) & "'"
        Case "DRCRTYPE"
            Select Case mType
                Case "LedDeb"
                    rpt.FormulaFields(I).TEXT = "'DR'"
                Case "LedCred"
                    rpt.FormulaFields(I).TEXT = "'CR'"
            End Select
        Case "TagadaHeader1"
            rpt.FormulaFields(I).TEXT = FaChk_Text(TxtHeader(0))
        Case "TagadaHeader2"
            rpt.FormulaFields(I).TEXT = FaChk_Text(TxtHeader(1))
        Case "TagadaHeader3"
            rpt.FormulaFields(I).TEXT = FaChk_Text(TxtHeader(2))
        Case "TagadaHeader4"
            rpt.FormulaFields(I).TEXT = FaChk_Text(TxtHeader(3))
        Case "TagadaHeader5"
            rpt.FormulaFields(I).TEXT = FaChk_Text(TxtHeader(4))
        Case "TagadaFooter1"
            rpt.FormulaFields(I).TEXT = FaChk_Text(TxtFooter(0))
        Case "TagadaFooter2"
            rpt.FormulaFields(I).TEXT = FaChk_Text(TxtFooter(1))
        Case "TagadaFooter3"
            rpt.FormulaFields(I).TEXT = FaChk_Text(TxtFooter(2))
        Case "TagadaFooter4"
            rpt.FormulaFields(I).TEXT = FaChk_Text(TxtFooter(3))
        Case "TagadaFooter5"
            rpt.FormulaFields(I).TEXT = FaChk_Text(TxtFooter(4))
        Case "PageNo"
            rpt.FormulaFields(I).TEXT = "'" & RstEnviro!pagenofill & "'"
        Case "Date1"
            rpt.FormulaFields(I).TEXT = "'" & RstEnviro!daterfill & "'"
    End Select
Next
rpt.Database.SetDataSource TmpRst
rpt.ReadRecords
Select Case mType
    Case "Led", "LedDeb"
        FaReport_View rpt, Index, Me.CAPTION, True
    Case Else
        FaReport_View rpt, 0, Me.CAPTION, True
End Select
ExitLoop:
Set Rst1 = Nothing
Set Rst2 = Nothing
Set TmpRst = Nothing
Set RstCheck21 = Nothing
Set RstDrX = Nothing
Set RstCrX = Nothing
TXTACC_CODE = ""
TXTACC_CODE1 = ""
Set rpt = Nothing
RepPrint = False
Exit Sub
Errloop:    MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub Text_GotFocus(Index As Integer)
Dim Rst1 As ADODB.Recordset
Select Case Index
    Case SignAmt, SignCr, SignDr
        ListArray = Array("=", "<", "<=", ">", ">=", "<>")
        Set mListItem = FaListView_Items(ListView, TEXT, Index, ListArray, 6)
    Case SiteCode
        Set RstSite = New ADODB.Recordset
        RstSite.CursorLocation = adUseClient
        RstSite.Open "Select Site_Code AS Code,Site_Desc AS NAME From Site Order by Site_Desc", G_FaCn, adOpenForwardOnly, adLockReadOnly
        Set DGSite.DataSource = RstSite
        If RstSite.RecordCount = 0 Or (RstSite.EOF = True Or RstSite.BOF = True) Or TEXT(Index).TEXT = "" Then Exit Sub
        If TEXT(Index).TEXT <> RstSite!Name Then
            RstSite.MoveFirst
            RstSite.FIND "Name =" & FaChk_Text(TEXT(Index).TEXT)
        End If
End Select
Set Rst1 = Nothing
End Sub
Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
    Select Case Index
        Case 0, 1, 2
            FaNumDown TEXT(Index), KeyCode, 10, 2
        Case 5
            FaNumDown TEXT(Index), KeyCode, 3, 2
        Case SignAmt, SignCr, SignDr
            FaListViewReport_KeyDown FrmList, ListView, TEXT, Index, KeyCode, Shift, FrDetail.left + TEXT(Index).left, (FrDetail.top + TEXT(Index).top + TEXT(Index).height + 25), TEXT(Index).width
        Case SiteCode
            If PubFaSiteType <> 0 Then
                FaDGridTxtKeyDown DGSite, TEXT, Index, RstSite, KeyCode, True, 1
            End If
    End Select
End Sub
Private Sub TEXT_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0, 1, 2
        FaNumPress TEXT(Index), KeyAscii, 10, 2
    Case 5
        FaNumPress TEXT(Index), KeyAscii, 3, 2
    Case 6, 7, PrintNarration, PrintVrNo
        If KeyAscii = 78 Or KeyAscii = 110 Then   'NO
            TEXT(Index) = "No"
            KeyAscii = 0
        ElseIf KeyAscii = 89 Or KeyAscii = 121 Then 'Yes
            TEXT(Index) = "Yes"
            KeyAscii = 0
        Else
            KeyAscii = 0
        End If
    Case SiteCode
        If PubFaSiteType <> 0 Then
            If DGSite.Visible = True Then FaDGridTxtKeyPress TEXT, Index, RstSite, KeyAscii, "Name"
        End If
End Select
End Sub
Private Sub TExt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case SignDr, SignCr, SignAmt
        If KeyCode <> 13 And FrmList.Visible = False Then Text_KeyDown Index, GridKey, 0
        FaListView_KeyUp ListView, TEXT, Index, KeyCode, mListItem
    Case SiteCode
        If PubFaSiteType <> 0 Then
            If KeyCode <> 13 And DGSite.Visible = False Then Text_KeyDown Index, GridKey, 0: FaDGridTxtKeyPress TEXT, Index, RstSite, KeyCode, "Name", True
        End If
End Select
End Sub
Private Sub TEXT_LostFocus(Index As Integer)
    TEXT_Validate Index, False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub
Private Sub TEXT_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case 0, 1, 2, 5
        TEXT(Index) = Format(FaValidate_Numeric(TEXT(Index)), "0.00")
    Case SignAmt, SignCr, SignDr
        TEXT(Index).TEXT = ListView.SelectedItem.TEXT
    Case SiteCode
        If PubFaSiteType <> 0 Then
            If RstSite.RecordCount = 0 Or (RstSite.EOF = True Or RstSite.BOF = True) Or TEXT(Index).TEXT = "" Then
                TEXT(SiteCode).Tag = ""
                TEXT(SiteCode) = ""
            Else
                TEXT(SiteCode) = RstSite!Name
                TEXT(SiteCode).Tag = RstSite!Code
            End If
        End If
End Select
End Sub
Private Sub DUELIST()
On Error GoTo ELoop
Dim I As Byte, RstRep As ADODB.Recordset, X1
Dim Rst1 As ADODB.Recordset, Rst2 As ADODB.Recordset, RST3 As ADODB.Recordset, Rst4 As ADODB.Recordset, RST5 As ADODB.Recordset, oBAL As Double
Dim Ac_Name As String, Ac_Code As String, mDueDr As Double, mDueCr As Double
Dim mDR As Double, mCR As Double, mDrSum As Double, mCrSum As Double
Dim TotCr As Double, TotDr As Double, tSubCode As String, tCount As Double, mSiteCode As String

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If Trim(TEXT(SiteCode)) <> "" Then
        If PubFaSiteType = 1 Then
            mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
        ElseIf PubFaSiteType = 2 Then
            mSiteCode = " And VIEWLEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
        End If
    Else
        mSiteCode = ""
    End If
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If FGrid.TextMatrix(List2, 1) = "" Then MsgBox "Specify More Than Days": RepPrint = False: Exit Sub
    Set RstRep = New ADODB.Recordset
    Set RstRep = RstDueList(RstRep)
    If GridString1 = "" Then
        Set Rst1 = G_FaCn.Execute("Select SG.SubCode,SG.Name,SG.Add1,SG.Add2,C.CityName,SG.ConPerson,SG.PIN,SG.EMail,SG.PhoneO,VIEWLEDGER.* From (SubGroup SG Left Join City C on SG.CityCode=C.CityCode) LEFT JOIN VIEWLEDGER ON VIEWLEDGER.PARTY=SG.SUBCODE Where SG.Nature='" & IIf(GRepFormName = "DueListCreditorOverLay", "Supplier", "Customer") & "' AND VIEWLEDGER.V_DATE<=" & FaConvertDate(FGrid.TextMatrix(Date1, 1)) & " " & mSiteCode & " Order By Name,SG.SUBCODE,VIEWLEDGER.V_DATE,VIEWLEDGER.DOCID")
        Set Rst2 = G_FaCn.Execute("Select SubCode+'-'+DocId1+'-'+ LTRIM(RTRIM(V_SNo1)) AS SCODE,SUBCODE,DocId1,V_SNo1,SUM(CR) AS ADJCr  FROM LEDGERADJ GROUP BY SUBCODE,DocId1,V_SNo1")
        Set RST3 = G_FaCn.Execute("Select SubCode+'-'+DocId2+'-'+ LTRIM(RTRIM(V_SNo2)) AS SCODE,SUBCODE,DocId2,V_SNo2,SUM(CR) AS ADJDr  FROM LEDGERADJ GROUP BY SUBCODE,DocId2,V_SNo2")
    Else
        Set Rst1 = G_FaCn.Execute("Select SG.SubCode,SG.Name,SG.Add1,SG.Add2,C.CityName,SG.ConPerson,SG.PIN,SG.EMail,SG.PhoneO,VIEWLEDGER.* From (SubGroup SG Left Join City C on SG.CityCode=C.CityCode) LEFT JOIN VIEWLEDGER ON VIEWLEDGER.PARTY=SG.SUBCODE Where SG.Nature='" & IIf(GRepFormName = "DueListCreditorOverLay", "Supplier", "Customer") & "' AND SG.SUBCODE IN (" & GridString1 & ") AND VIEWLEDGER.V_DATE<=" & FaConvertDate(FGrid.TextMatrix(Date1, 1)) & " " & mSiteCode & " Order By Name,SG.SUBCODE,VIEWLEDGER.V_DATE,VIEWLEDGER.DOCID")
        Set Rst2 = G_FaCn.Execute("Select SubCode+'-'+DocId1+'-'+ LTRIM(RTRIM(V_SNo1)) AS SCODE,SUBCODE,DocId1,V_SNo1,SUM(CR) AS ADJCr  FROM LEDGERADJ WHERE SUBCODE IN (" & GridString1 & ") GROUP BY SUBCODE,DocId1,V_SNo1")
        Set RST3 = G_FaCn.Execute("Select SubCode+'-'+DocId2+'-'+ LTRIM(RTRIM(V_SNo2)) AS SCODE,SUBCODE,DocId2,V_SNo2,SUM(CR) AS ADJDr  FROM LEDGERADJ WHERE SUBCODE IN (" & GridString1 & ") GROUP BY SUBCODE,DocId2,V_SNo2")
    End If
'    Set Rst4 = G_FaCn.Execute("Select (S.V_Type+'-'+RTrim(LTrim(Convert(Char(8),S.V_No)))) As tCode,C.PartyOrdNo,C.PartyOrdDate,S.PaymentTerms,D.Name As DocThroughName From ((SBill S Left Join SBill1 S1 on S.DocID=S1.DocID) Left Join DocThrough D on S.DocThrough=D.Code) Left Join CN C on S1.CNDocID=C.DocID Order By (S.V_Type+'-'+RTrim(LTrim(Convert(Char(8),S.V_No))))")
    Do Until Rst1.EOF
        Ac_Name = Trim(IIf(IsNull(Rst1!Name), "", Rst1!Name))
        Ac_Code = Rst1!Party
        mDR = 0: mCR = 0: mDrSum = 0: mCrSum = 0
        Do While Rst1!Party = Ac_Code
            mDueDr = 0
            mDueCr = 0
            If FaVNull(Rst1!DEBIT) > 0 Then
                If RST3.RecordCount > 0 Then RST3.MoveFirst
                RST3.FIND "SCODE='" & Rst1!SubCode + "-" + Rst1!DocId + "-" + Trim(Rst1!V_SNo) & "'"
                If RST3.EOF = False Then
                    mDueDr = FaVNull(RST3!ADJDR)
                End If
                If FGrid.TextMatrix(List1, 1) = "Pending" And Format(Rst1!DEBIT, "0.00") - Format(mDueDr, "0.00") = 0 Then GoTo NEXTLOOP
                With RstRep
                    .AddNew
                    .Fields("VDATE1") = Rst1!V_DATE
                    .Fields("VNO1") = Rst1!V_NO
                    .Fields("VTYPE1") = Rst1!V_tYPE
                    .Fields("VSNO1") = Rst1!V_SNo
                    .Fields("VADD1") = Rst1!V_ADD
                    .Fields("Dr") = Rst1!DEBIT
                    .Fields("PendDr") = Rst1!DEBIT - mDueDr
                    .Fields("AgeDr") = DateDiff("D", Rst1!V_DATE, CDate(FGrid.TextMatrix(Date1, 1)))
                    .Fields("SUBCODE") = Rst1!Party
                    .Fields("NAME") = Rst1!Name
                    .Fields("ADDRESS") = IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + "," + IIf(IsNull(Rst1!Add2), "", Rst1!Add2)
                    .Fields("CITY_NAME") = IIf(IsNull(Rst1!CityName), "", Rst1!CityName)
                    .Fields("ConPerson") = IIf(IsNull(Rst1!ConPerson), "", Rst1!ConPerson)
                    .Fields("PIN") = IIf(IsNull(Rst1!Pin), "", Rst1!Pin)
                    .Fields("EMail") = IIf(IsNull(Rst1!EMail), "", Rst1!EMail)
                    .Fields("NARRATION1") = FaXNull(Rst1!mNarr)
                    .Fields("NARRATION2") = FaXNull(Rst1!Narr)
                    .Update
                End With
            ElseIf FaVNull(Rst1!CREDIT) > 0 Then
                If Rst2.RecordCount > 0 Then Rst2.MoveFirst
                Rst2.FIND "SCODE='" & Rst1!SubCode + "-" + Rst1!DocId + "-" + Trim(Rst1!V_SNo) & "'"
                If Rst2.EOF = False Then
                    mDueCr = FaVNull(Rst2!ADJCR)
                End If
                If FGrid.TextMatrix(List1, 1) = "Pending" And Format(Rst1!CREDIT, "0.00") - Format(mDueCr, "0.00") = 0 Then GoTo NEXTLOOP
                With RstRep
                    .AddNew
                    .Fields("VDATE2") = Rst1!V_DATE
                    .Fields("VNO2") = Rst1!V_NO
                    .Fields("VTYPE2") = Rst1!V_tYPE
                    .Fields("VSNO2") = Rst1!V_SNo
                    .Fields("VADD2") = Rst1!V_ADD
                    .Fields("Cr") = Rst1!CREDIT
                    .Fields("PendCr") = Rst1!CREDIT - mDueCr
                    .Fields("AgeCr") = DateDiff("D", Rst1!V_DATE, CDate(FGrid.TextMatrix(Date1, 1)))
                    .Fields("SUBCODE") = Rst1!Party
                    .Fields("NAME") = Rst1!Name
                    .Fields("ADDRESS") = IIf(IsNull(Rst1!Add1), "", Rst1!Add1) + "," + IIf(IsNull(Rst1!Add2), "", Rst1!Add2)
                    .Fields("CITY_NAME") = IIf(IsNull(Rst1!CityName), "", Rst1!CityName)
                    .Fields("NARRATION1") = FaXNull(Rst1!mNarr)
                    .Fields("NARRATION2") = FaXNull(Rst1!Narr)
                    .Update
                End With
            End If
NEXTLOOP:
            Rst1.MoveNext
            If Rst1.EOF = True Then Exit Do
        Loop
    Loop
    '''/// For Period
    If Rst1.RecordCount > 0 Then Rst1.MoveFirst
    While Not Rst1.EOF
        tSubCode = Rst1!Party
        tCount = 0
        RstRep.Filter = adFilterNone
        If RstRep.RecordCount > 0 Then RstRep.MoveFirst
        While Not RstRep.EOF
            If (RstRep!SubCode = tSubCode) And (RstRep!AgeDr >= Val(FGrid.TextMatrix(List2, 1))) Then
                tCount = tCount + 1
                If tCount > 0 Then GoTo NextLoopPeriod Else GoTo NextLoopDelete
            End If
        RstRep.MoveNext
        Wend
        If tCount = 0 Then GoTo NextLoopDelete Else GoTo NextLoopPeriod
NextLoopDelete:
        If tSubCode = Rst1!Party Then
            If RstRep.RecordCount > 0 Then RstRep.MoveFirst
            RstRep.Filter = ("SubCode='" & tSubCode & "'")
            If RstRep.EOF = False Then
                While Not RstRep.EOF
                    RstRep.Delete
                    RstRep.MoveNext
                Wend
            End If
        End If
NextLoopPeriod:
    Rst1.MoveNext
    Wend
    RstRep.Filter = adFilterNone
    '''/// End For Period
    Set Rst1 = Nothing
    Set Rst2 = Nothing
    Set RST3 = Nothing
    Set Rst4 = Nothing
    Set RST5 = Nothing
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False
    RepTitle = UCase(Me.CAPTION)
    X1 = CreateFieldDefFile(RstRep, PubFaReportPath + "\FaDueList.ttx", True)
    Set rpt = rdApp.OpenReport(PubFaReportPath + "\FaDueList.RPT")
    rpt.Database.SetDataSource RstRep
    Set RstRep = Nothing
Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub ControlLed(Index As Integer)
On Error GoTo ELoop
Dim Rst1 As ADODB.Recordset, RstTmp As ADODB.Recordset, Rst2 As ADODB.Recordset, RST3 As ADODB.Recordset
Dim mAcCode As String, mDocNo As String, mDocNo1 As String, TinTin As Integer, Condstr As String, X11, mSiteCode As String

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub
If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And VIEWLEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
End If

If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then GoTo EXIT_SUB
If GridString1 <> "" Then Condstr = " AND GROUPCODE IN (" & GridString1 & ") "
Set RstTmp = New ADODB.Recordset
Set RstTmp = ControlLedRst(RstTmp)

Set RST3 = G_FaCn.Execute("SELECT SUM(CREDIT)-SUM(DEBIT) AS BALANCE FROM VIEWLEDGER LEFT JOIN SUBGROUP ON VIEWLEDGER.party=SUBGROUP.SUBCODE where V_date<" & FaConvertDate(TXTS_DATE) & " " & Condstr & "  " & mSiteCode & "")
Set Rst2 = G_FaCn.Execute("SELECT VIEWLEDGER.*,SUBGROUP.NAME FROM VIEWLEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=VIEWLEDGER.PARTY WHERE V_dATE BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & "  " & mSiteCode & " ORDER BY DOCID")
Set Rst1 = G_FaCn.Execute("SELECT DISTINCT DOCID FROM VIEWLEDGER LEFT JOIN SUBGROUP ON VIEWLEDGER.party=SUBGROUP.SUBCODE where V_date BETWEEN " & FaConvertDate(TXTS_DATE) & " AND " & FaConvertDate(TXTE_DATE) & " " & Condstr & "  " & mSiteCode & " ORDER BY DOCID")
If RST3.RecordCount > 0 Then
    If RST3!Balance <> 0 Then
        With RstTmp
            .AddNew
            !V_DATE = TXTS_DATE
            If RST3!Balance > 0 Then
                !CREDIT = Abs(RST3!Balance)
                !DEBIT = 0
            Else
                !CREDIT = 0
                !DEBIT = Abs(RST3!Balance)
            End If
            !V_tYPE = "OP"
            !V_NO = 0
            !V_ADD = ""
            !Chq_No = ""
            .Update
        End With
    End If
End If
Do Until Rst1.EOF
    Rst2.MoveFirst
    Rst2.FIND "DocId='" & Rst1!DocId & "'"
    mDocNo1 = Rst1!DocId
    If Rst2.EOF = False Then
        Do While mDocNo1 = Rst2!DocId
            With RstTmp
                .AddNew
                !V_DATE = Rst2!V_DATE
                !CREDIT = Rst2!CREDIT
                !DEBIT = Rst2!DEBIT
                !V_tYPE = Rst2!V_tYPE
                !V_NO = Rst2!V_NO
                !V_ADD = Rst2!V_ADD
                !Chq_No = Rst2!Chq_No
                !Chq_Date = Rst2!Chq_Date
                !Narr = Rst2!Narr
                !Name = Rst2!Name
                !mNarr = Rst2!mNarr
                .Update
            End With
            Rst2.MoveNext
            If Rst2.EOF = True Then Exit Do
        Loop
    End If
    Rst1.MoveNext
Loop
If RstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
Select Case Index
    Case 1
        MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
        TinTin = PubDatamanFa.FaControlLedgerDosPrinting(Me, RstTmp)
        GoTo EXIT_SUB
    Case Else
        X11 = CreateFieldDefFile(RstTmp, PubFaReportPath + "\FaControlLed.ttx", True)
        Set rpt = PubDatamanFa.FaControlLedRpt
        rpt.Database.SetDataSource RstTmp
End Select
EXIT_SUB:
    Set Rst1 = Nothing
    Set RstTmp = Nothing
    Set Rst2 = Nothing
    Set RST3 = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Private Function RstDueList(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "VDATE1", adDate, , adFldIsNullable
    .Fields.Append "VNO1", adInteger, , adFldIsNullable
    .Fields.Append "VTYPE1", adVarChar, 5, adFldIsNullable
    .Fields.Append "VSNO1", adInteger, , adFldIsNullable
    .Fields.Append "VADD1", adVarChar, 5, adFldIsNullable
    .Fields.Append "VDATE2", adDate, , adFldIsNullable
    .Fields.Append "VNO2", adInteger, , adFldIsNullable
    .Fields.Append "VTYPE2", adVarChar, 5, adFldIsNullable
    .Fields.Append "VSNO2", adInteger, , adFldIsNullable
    .Fields.Append "VADD2", adVarChar, 5, adFldIsNullable
    .Fields.Append "Cr", adDouble, , adFldIsNullable
    .Fields.Append "Dr", adDouble, , adFldIsNullable
    .Fields.Append "PendCr", adDouble, , adFldIsNullable
    .Fields.Append "PendDr", adDouble, , adFldIsNullable
    .Fields.Append "AgeCr", adInteger, , adFldIsNullable
    .Fields.Append "AgeDr", adInteger, , adFldIsNullable
    .Fields.Append "SUBCODE", adVarChar, 10
    .Fields.Append "NAME", adVarChar, 50
    .Fields.Append "ADDRESS", adVarChar, 150
    .Fields.Append "CITY_NAME", adVarChar, 50
    .Fields.Append "ConPerson", adVarChar, 50, adFldIsNullable
    .Fields.Append "PIN", adVarChar, 6, adFldIsNullable
    .Fields.Append "EMail", adVarChar, 50, adFldIsNullable
    .Fields.Append "NARRATION1", adVarChar, 255
    .Fields.Append "NARRATION2", adVarChar, 255
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set RstDueList = Rst
End Function
Private Function ControlLedRst(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "V_DATE", adDate, , adFldIsNullable
    .Fields.Append "credit", adDouble, , adFldIsNullable
    .Fields.Append "debit", adDouble, , adFldIsNullable
    .Fields.Append "v_type", adVarChar, 5, adFldIsNullable
    .Fields.Append "v_no", adInteger, , adFldIsNullable
    .Fields.Append "v_add", adVarChar, 5, adFldIsNullable
    .Fields.Append "Chq_No", adVarChar, 20, adFldIsNullable
    .Fields.Append "CHQ_DATE", adDate, , adFldIsNullable
    .Fields.Append "NARR", adVarChar, 255, adFldIsNullable
    .Fields.Append "NAME", adVarChar, 50, adFldIsNullable
    .Fields.Append "MNARR", adVarChar, 255, adFldIsNullable
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set ControlLedRst = Rst
End Function
Private Sub RozNamcha(Index As Integer)
On Error GoTo ELoop
Dim Rst1 As ADODB.Recordset, RstTmp As ADODB.Recordset, mGROUP_rs As ADODB.Recordset, SUBGROUP_rs As ADODB.Recordset, TmpGrs As ADODB.Recordset, TmpGrs1 As ADODB.Recordset
Dim DrAc As String, CrAc As String, oBAL As Double, mAcCode As String, mDocNo As String, mDocNo1 As String
Dim mNARR1 As String, mNARR2 As String, TmpDate As Date, mDate1 As Date, mDate2 As Date, TinTin As Integer
Dim mFLAG1 As Boolean, mFLAG2 As Boolean, mFLAG11 As Boolean, mFLAG22 As Boolean, mFLAG111 As Boolean, mFLAG222 As Boolean, mSiteCode As String

If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List2, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
If IsNotBlank(List3, FGrid.TextMatrix(List3, 0)) = False Then RepPrint = False: Exit Sub
TXTS_DATE = FGrid.TextMatrix(Date1, 1)
TXTE_DATE = FGrid.TextMatrix(Date2, 1)
If FaValidDate(Me) = 0 Then RepPrint = False: Exit Sub
If Trim(TEXT(SiteCode)) <> "" Then
    If PubFaSiteType = 1 Then
        mSiteCode = " And RIGHT(VIEWLEDGER.Site_Code,1)='" & TEXT(SiteCode).Tag & "'"
    ElseIf PubFaSiteType = 2 Then
        mSiteCode = " And VIEWLEDGER.Site_Code='" & TEXT(SiteCode).Tag & "'"
    End If
Else
    mSiteCode = ""
End If
TXTACC_CODE = Trim(FGrid.TextMatrix(List3, 1))
mAcCode = Trim(FGrid.TextMatrix(List3, 2))
DrAc = ""
CrAc = ""
oBAL = 0

Set Rst1 = G_FaCn.Execute("SELECT GROUPNATURE FROM PARTY_LIST WHERE SUBCODE='" & mAcCode & "'")

If Rst1.RecordCount <= 0 Then Exit Sub
If Rst1!GroupNature = "A" Or Rst1!GroupNature = "L" Then
    If PubBackEnd = "S" Then
        oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "").Fields(0)
    ElseIf PubBackEnd = "A" Then
        oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "").Fields(0)
    End If
Else
    If PubBackEnd = "S" Then
        oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "").Fields(0)
    ElseIf PubBackEnd = "A" Then
        oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<" & FaConvertDate(TXTS_DATE) & " " & mSiteCode & "").Fields(0)
    End If
End If
Set RstTmp = New ADODB.Recordset
Set RstTmp = PubDatamanFa.FaCASHTMP1(RstTmp)
If oBAL <> 0 Then
    RstTmp.AddNew
    RstTmp!V_DATE = TXTS_DATE
    If oBAL < 0 Then
        RstTmp!Name = "OPENING BALANCE"
        RstTmp!cr = Abs(oBAL)
    Else
        RstTmp!Name1 = "OPENING BALANCE"
        RstTmp!ADJAMT = Abs(oBAL)
    End If
    RstTmp.Update
End If

If PubBackEnd = "S" Then
    Set SUBGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.*,subgroup.NAME,CONVERT(VARCHAR,VIEWLEDGER.CHQ_DATE,103)AS CHQDATE  FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND CREDIT>0 AND PARTY<>'" & mAcCode & "' " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
    Set mGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.*,subgroup.NAME,CONVERT(VARCHAR,VIEWLEDGER.CHQ_DATE,103)AS CHQDATE  FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND DEBIT>0 AND PARTY<>'" & mAcCode & "' " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
ElseIf PubBackEnd = "A" Then
    Set SUBGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.*,subgroup.NAME,FORMAT(VIEWLEDGER.CHQ_DATE,'DD/MM/YY') AS CHQDATE  FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND CREDIT>0 AND PARTY<>'" & mAcCode & "' " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
    Set mGROUP_rs = G_FaCn.Execute("SELECT VIEWLEDGER.*,subgroup.NAME,FORMAT(VIEWLEDGER.CHQ_DATE,'DD/MM/YY') AS CHQDATE  FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE V_DATE Between " & FaConvertDate(TXTS_DATE) & " And " & FaConvertDate(TXTE_DATE) & " AND DEBIT>0 AND PARTY<>'" & mAcCode & "' " & mSiteCode & " ORDER BY V_DATE,V_TYPE,V_NO,v_add,V_SNO")
End If

If Not (mGROUP_rs.EOF) Then mDate2 = mGROUP_rs!V_DATE
If Not (SUBGROUP_rs.EOF) Then mDate1 = SUBGROUP_rs!V_DATE
mFLAG1 = False
mFLAG2 = False
mFLAG11 = False
mFLAG22 = False
mFLAG111 = False
mFLAG222 = False
mNARR1 = ""
mNARR2 = ""
TmpDate = TXTS_DATE
Do Until mGROUP_rs.EOF And SUBGROUP_rs.EOF
    If mDate1 = TmpDate Or mDate2 = TmpDate Then
        RstTmp.AddNew
        If mDate1 = TmpDate Then
            RstTmp!V_tYPE = SUBGROUP_rs!V_tYPE
            RstTmp!V_DATE = mDate1
            RstTmp!V_NO = SUBGROUP_rs!V_NO
            RstTmp!V_SNo = SUBGROUP_rs!V_SNo
            If PubFaSiteType = 1 Then
                mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(SUBGROUP_rs!DocId, 1), "")
                mDocNo = mDocNo + IIf(RstEnviro!LedSiteCode = "Yes", Trim(Right(SUBGROUP_rs!Site_Code, 1)), "")
                mDocNo = mDocNo + IIf(RstEnviro!LedPrefix = "Yes", IIf(mDocNo = "", "", "/") + Trim(SUBGROUP_rs!V_ADD), "")
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + left(Trim(SUBGROUP_rs!V_tYPE), 1) + Trim(mID(Trim(SUBGROUP_rs!V_tYPE), 3, 3))
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(SUBGROUP_rs!V_NO))
            Else
                mDocNo = IIf(RstEnviro!LedDivCode = "Yes", left(SUBGROUP_rs!DocId, 1), "") + IIf(RstEnviro!LedSiteCode = "Yes", Trim(left(SUBGROUP_rs!Site_Code, 1)), "")
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + IIf(RstEnviro!LedPrefix = "Yes", Trim(SUBGROUP_rs!V_ADD), "")
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(SUBGROUP_rs!V_tYPE)
                mDocNo = mDocNo + IIf(mDocNo = "", "", "/") + Trim(STR(SUBGROUP_rs!V_NO))
            End If
            RstTmp!DOCNO = mDocNo
            If mFLAG1 = False Or mFLAG11 = False Or mFLAG111 = False Then
                mFLAG1 = True
                mNARR1 = ""
                If FaXNull(Trim(SUBGROUP_rs!Chq_No)) <> "" Then mNARR1 = mNARR1 + "Ch.No:" + Trim(FaXNull(SUBGROUP_rs!Chq_No)) + " Ch.Dt: " + CStr(FaXNull(SUBGROUP_rs!ChqDate))
                mNARR1 = mNARR1 + Trim(FaXNull(SUBGROUP_rs!mNarr)) + Trim(FaXNull(SUBGROUP_rs!Narr))
                If Len(FaXNull(SUBGROUP_rs!Name)) <> 0 Then
                    RstTmp!Name = FaXNull(SUBGROUP_rs!Name)
                    RstTmp!cr = Format(SUBGROUP_rs!CREDIT, "0.00")
'                    If mAcCode = SUBGROUP_rs!Party Then oBAL = oBAL - Format(SUBGROUP_rs!CREDIT, "0.00")
                    mFLAG111 = True
                    mFLAG11 = True
                Else
                    If mFLAG11 = False And Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then
                        If mFLAG111 = False Then
                            RstTmp!cr = Format(SUBGROUP_rs!CREDIT, "0.00")
'                            If mAcCode = SUBGROUP_rs!Party Then oBAL = oBAL - Format(SUBGROUP_rs!CREDIT, "0.00")
                            RstTmp!Name = "As Per Detail"
                            mFLAG111 = True
                            If Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then Set TmpGrs = G_FaCn.Execute("SELECT subgroup.NAME,VIEWLEDGER.DEBIT,VIEWLEDGER.Credit FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE VIEWLEDGER.DOCID='" & SUBGROUP_rs!DocId & "' AND V_SNO<>" & SUBGROUP_rs!V_SNo)
                        Else
                            If Not TmpGrs.EOF Then
                                If TmpGrs!DEBIT > 0 Then
                                    RstTmp!Name = Space(2) + FaSetW(TmpGrs!Name, 22) + " " + FaSetN(FaSNull(TmpGrs!DEBIT), 12) + " Dr"
                                Else
                                    RstTmp!Name = Space(2) + FaSetW(TmpGrs!Name, 22) + " " + FaSetN(FaSNull(TmpGrs!CREDIT), 12) + " Cr"
                                End If
                                TmpGrs.MoveNext
                            End If
                            If TmpGrs.EOF = True Then mFLAG11 = True
                        End If
                    Else
                        RstTmp!Name = Space(2) + Trim(mID(mNARR1, 1, 36))
                        RstTmp!cr = Format(SUBGROUP_rs!CREDIT, "0.00")
'                        If mAcCode = SUBGROUP_rs!Party Then oBAL = oBAL - Format(SUBGROUP_rs!CREDIT, "0.00")
                        mNARR1 = Trim(mID(mNARR1, 37, 510))
                        mFLAG111 = True
                        mFLAG11 = True
                    End If
                End If
                If Len(mNARR1) <= 0 And mFLAG11 = True And mFLAG111 = True Then
                    mFLAG1 = False
                    SUBGROUP_rs.MoveNext
                    If Not SUBGROUP_rs.EOF Then
                        mDate1 = SUBGROUP_rs!V_DATE
                    Else
                        mDate1 = DateAdd("D", 1, TXTE_DATE)
                    End If
                End If
            Else
                mNARR1 = Trim(mNARR1)
                RstTmp!Name = Space(2) + Trim(mID(mNARR1, 1, 36))
                mNARR1 = Trim(mID(mNARR1, 37, 510))
                If Len(mNARR1) <= 0 Then
                    mFLAG1 = False
                    mFLAG11 = False
                    mFLAG111 = False
                    SUBGROUP_rs.MoveNext
                    If Not SUBGROUP_rs.EOF Then
                        mDate1 = SUBGROUP_rs!V_DATE
                    Else
                        mDate1 = DateAdd("D", 1, TXTE_DATE)
                    End If
                End If
            End If
        End If
        If mDate2 = TmpDate Then
            RstTmp!VType = mGROUP_rs!V_tYPE
            RstTmp!VNo = mGROUP_rs!V_NO
            RstTmp!V_DATE = mDate2
            RstTmp!VSNo = mGROUP_rs!V_SNo
            If PubFaSiteType = 1 Then
                mDocNo1 = IIf(RstEnviro!LedDivCode = "Yes", left(mGROUP_rs!DocId, 1), "")
                mDocNo1 = mDocNo1 + IIf(RstEnviro!LedSiteCode = "Yes", Trim(Right(mGROUP_rs!Site_Code, 1)), "")
                mDocNo1 = mDocNo1 + IIf(RstEnviro!LedPrefix = "Yes", IIf(mDocNo1 = "", "", "/") + Trim(mGROUP_rs!V_ADD), "")
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + left(Trim(mGROUP_rs!V_tYPE), 1) + Trim(mID(Trim(mGROUP_rs!V_tYPE), 3, 3))
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + Trim(STR(mGROUP_rs!V_NO))
            Else
                mDocNo1 = IIf(RstEnviro!LedDivCode = "Yes", left(mGROUP_rs!DocId, 1), "") + IIf(RstEnviro!LedSiteCode = "Yes", Trim(left(mGROUP_rs!Site_Code, 1)), "")
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + IIf(RstEnviro!LedPrefix = "Yes", Trim(mGROUP_rs!V_ADD), "")
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + Trim(mGROUP_rs!V_tYPE)
                mDocNo1 = mDocNo1 + IIf(mDocNo1 = "", "", "/") + Trim(STR(mGROUP_rs!V_NO))
            End If
            RstTmp!DocNo1 = mDocNo1
            If mFLAG2 = False Or mFLAG22 = False Or mFLAG222 = False Then
                mFLAG2 = True
                mNARR2 = ""
                If FaXNull(Trim(mGROUP_rs!Chq_No)) <> "" Then mNARR2 = mNARR2 + "Ch.No:" + Trim(FaXNull(mGROUP_rs!Chq_No)) + " Ch.Dt: " + CStr(FaXNull(mGROUP_rs!ChqDate))
                mNARR2 = mNARR2 + Trim(FaXNull(mGROUP_rs!mNarr)) + Trim(FaXNull(mGROUP_rs!Narr))
                If Len(FaXNull(mGROUP_rs!Name)) <> 0 Then
                    RstTmp!Name1 = mGROUP_rs!Name
                    RstTmp!ADJAMT = Format(mGROUP_rs!DEBIT, "0.00")
                    mFLAG222 = True
                    mFLAG22 = True
                Else
                    If mFLAG22 = False And Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then
                        If mFLAG222 = False Then
                            RstTmp!ADJAMT = Format(mGROUP_rs!DEBIT, "0.00")
                            RstTmp!Name1 = "As Per Detail"
                            mFLAG222 = True
                            If Trim(FGrid.TextMatrix(List1, 1)) = "Yes" Then Set TmpGrs1 = G_FaCn.Execute("SELECT subgroup.NAME,VIEWLEDGER.DEBIT,VIEWLEDGER.Credit FROM VIEWLEDGER LEFT JOIN subgroup ON VIEWLEDGER.PARTY=subgroup.SUBCODE WHERE VIEWLEDGER.DOCID='" & mGROUP_rs!DocId & "' AND V_SNO<>" & mGROUP_rs!V_SNo)
                        Else
                            If Not TmpGrs1.EOF Then
                                If TmpGrs1!DEBIT > 0 Then
                                    RstTmp!Name1 = Space(2) + FaSetW(TmpGrs1!Name, 22) + " " + FaSetN(FaSNull(TmpGrs1!DEBIT), 12) + " Dr"
                                Else
                                    RstTmp!Name1 = Space(2) + FaSetW(TmpGrs1!Name, 22) + " " + FaSetN(FaSNull(TmpGrs1!CREDIT), 12) + " Cr"
                                End If
                                TmpGrs1.MoveNext
                            End If
                            If TmpGrs1.EOF = True Then mFLAG22 = True
                        End If
                    Else
                        RstTmp!Name1 = Space(2) + Trim(mID(mNARR2, 1, 36))
                        mNARR2 = Trim(mID(mNARR2, 37, 510))
                        RstTmp!ADJAMT = Format(mGROUP_rs!DEBIT, "0.00")
'                        If mAcCode = mGROUP_rs!Party Then oBAL = oBAL + Format(mGROUP_rs!DEBIT, "0.00")
                        mFLAG222 = True
                        mFLAG22 = True
                    End If
                End If
                If Len(mNARR2) <= 0 And mFLAG22 = True And mFLAG222 = True Then
                    mFLAG2 = False
                    mGROUP_rs.MoveNext
                    If Not mGROUP_rs.EOF Then
                        mDate2 = mGROUP_rs!V_DATE
                    Else
                        mDate2 = DateAdd("D", 1, TXTE_DATE)
                    End If
                End If
            Else
                mNARR2 = Trim(mNARR2)
                RstTmp!Name1 = Space(2) + Trim(mID(mNARR2, 1, 36))
                mNARR2 = Trim(mID(mNARR2, 37, 510))
                If Len(mNARR2) <= 0 Then
                    mFLAG2 = False
                    mFLAG22 = False
                    mFLAG222 = False
                    mGROUP_rs.MoveNext
                    If Not mGROUP_rs.EOF Then
                        mDate2 = mGROUP_rs!V_DATE
                    Else
                        mDate2 = DateAdd("D", 1, TXTE_DATE)
                    End If
                End If
            End If
        End If
        RstTmp.Update
    Else
        If Rst1!GroupNature = "A" Or Rst1!GroupNature = "L" Then
            If PubBackEnd = "S" Then
                oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
            ElseIf PubBackEnd = "A" Then
                oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
            End If
        Else
            If PubBackEnd = "S" Then
                oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
            ElseIf PubBackEnd = "A" Then
                oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
            End If
        End If
        If oBAL <> 0 Then
            RstTmp.AddNew
            RstTmp!V_DATE = TmpDate
            If oBAL > 0 Then
                RstTmp!Name = "CLOSING BALANCE"
                RstTmp!cr = Abs(oBAL)
            Else
                RstTmp!Name1 = "CLOSING BALANCE"
                RstTmp!ADJAMT = Abs(oBAL)
            End If
            RstTmp.Update
        End If
        If mDate1 <= mDate2 Then
            If mDate1 = CDate("12:00:00 AM") Then
                TmpDate = mDate2
            Else
                TmpDate = mDate1
            End If
        Else
            If mDate2 = CDate("12:00:00 AM") Then
                TmpDate = mDate1
            Else
                TmpDate = mDate2
            End If
        End If
        
        If oBAL <> 0 Then
            RstTmp.AddNew
            RstTmp!V_DATE = TmpDate
            If oBAL < 0 Then
                RstTmp!Name = "OPENING BALANCE"
                RstTmp!cr = Abs(oBAL)
            Else
                RstTmp!Name1 = "OPENING BALANCE"
                RstTmp!ADJAMT = Abs(oBAL)
            End If
            RstTmp.Update
        End If
    End If
Loop
If Rst1!GroupNature = "A" Or Rst1!GroupNature = "L" Then
    If PubBackEnd = "S" Then
        oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
    ElseIf PubBackEnd = "A" Then
        oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
    End If
Else
    If PubBackEnd = "S" Then
        oBAL = G_FaCn.Execute("SELECT ISNULL(SUM(CREDIT),0)-ISNULL(SUM(DEBIT),0) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
    ElseIf PubBackEnd = "A" Then
        oBAL = G_FaCn.Execute("SELECT IIF(ISNULL(SUM(CREDIT)),0,SUM(CREDIT))-IIF(ISNULL(SUM(DEBIT)),0,SUM(DEBIT)) FROM VIEWLEDGER WHERE PARTY='" & mAcCode & "' AND V_DATE>=" & FaConvertDate(PubStartDate) & " AND V_DATE<=" & FaConvertDate(TmpDate) & " " & mSiteCode & "").Fields(0)
    End If
End If
If oBAL <> 0 Then
    RstTmp.AddNew
    RstTmp!V_DATE = TmpDate
    If oBAL > 0 Then
        RstTmp!Name = "CLOSING BALANCE"
        RstTmp!cr = Abs(oBAL)
    Else
        RstTmp!Name1 = "CLOSING BALANCE"
        RstTmp!ADJAMT = Abs(oBAL)
    End If
    RstTmp.Update
End If
If RstTmp.RecordCount > 0 Then RstTmp.MoveFirst
If RstTmp.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
Select Case Index
    Case 1
        MsgBox "Please Set 80 Col. Paper in Your D.M.Printer", vbDefaultButton1 + vbInformation, "Paper Setting"
        TinTin = PubDatamanFa.FaRozNamchaDosPrinting(Me, RstTmp, Val(FGrid.TextMatrix(List2, 1)))
        GoTo EXIT_SUB
    Case Else
        Dim X11
        X11 = CreateFieldDefFile(RstTmp, PubFaReportPath + "\FaRozNamch.ttx", True)
        If RstEnviro!LedDivCode = "No" And RstEnviro!LedSiteCode = "No" And RstEnviro!LedPrefix = "No" Then
            Set rpt = PubDatamanFa.FaRozNamchaPortraitRpt
        Else
            Set rpt = PubDatamanFa.FaRozNamchaRpt
        End If
        rpt.Database.SetDataSource RstTmp
End Select
EXIT_SUB:
    Set Rst1 = Nothing
    Set mGROUP_rs = Nothing
    Set SUBGROUP_rs = Nothing
    Set TmpGrs = Nothing
    Set TmpGrs1 = Nothing
    Set RstTmp = Nothing
    Exit Sub
ELoop:  RepPrint = False
        MsgBox err.Description
End Sub
Public Function mDetailTrial(Rst As ADODB.Recordset) As ADODB.Recordset
With Rst
    .Fields.Append "TT", adInteger, , adFldIsNullable
    .Fields.Append "GroupName", adVarChar, 255, adFldIsNullable
    .Fields.Append "GRCODE", adVarChar, 255, adFldIsNullable
    .Fields.Append "PARTYCODE", adVarChar, 255, adFldIsNullable
    .Fields.Append "PARTYNAME", adVarChar, 255, adFldIsNullable
    .Fields.Append "OP_CR", adDouble, , adFldIsNullable
    .Fields.Append "OP_dR", adDouble, , adFldIsNullable
    .Fields.Append "BALANCECR", adDouble, , adFldIsNullable
    .Fields.Append "BALANCEDR", adDouble, , adFldIsNullable
    .Fields.Append "Bal", adInteger, , adFldIsNullable
    .Fields.Append "mgrcode", adVarChar, 255, adFldIsNullable
    .Fields.Append "GRName", adVarChar, 255, adFldIsNullable
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With
Set mDetailTrial = Rst
End Function



'''''Public Sub JBookDos(FormName As Object, RstRowSource As ADODB.Recordset, DayTot As String)
'''''On Error GoTo ELoop
'''''Dim mROW As Integer, mPAGE As Integer, PrintLine As String, RstEnvi As ADODB.Recordset
'''''Dim mDR As Double, mCR As Double, mDrSum As Double, mCrSum As Double, mDate As Date
'''''Dim mNarr As String, Narr As String, mChq As String, mVrNo As String, mSepNarr As Boolean, mComNarr As Boolean
'''''mROW = 0
'''''mPAGE = 1
'''''mDR = 0
'''''mCR = 0
'''''mDrSum = 0
'''''mCrSum = 0
'''''Set RstEnvi = G_FaCn.Execute("SELECT * FROM FAENVIRO")
'''''Dater = IIf(FaXNull(RstEnvi!daterfill) = "", "Y", FaXNull(RstEnvi!daterfill))
'''''Pager = IIf(FaXNull(RstEnvi!pagenofill) = "", "Y", FaXNull(RstEnvi!pagenofill))
'''''Titler = IIf(FaXNull(RstEnvi!titlerfill) = "", "M", FaXNull(RstEnvi!titlerfill))
'''''LineFil = IIf(FaXNull(RstEnvi!linefiller) = "", "-", FaXNull(RstEnvi!linefiller))
'''''Open "C:\reptmp\TEXTFILE.TXT" For Output As #1
'''''PrintLine = String(123, LineFil)
'''''Do Until RstRowSource.EOF
'''''    If mROW = 0 Then
'''''        mROW = FaPRNTIT(mPAGE, mROW, 80)
'''''        If PubFaSiteType <> 0 And FormName.FGrid.TextMatrix(4, 1) <> "" Then
'''''            Print #1, FaPRN_TIT(Trim(FormName.FGrid.TextMatrix(2, 1)) + " (" + FormName.FGrid.TextMatrix(4, 1) + ")", "B", 80)
'''''        Else
'''''            Print #1, FaPRN_TIT(Trim(FormName.FGrid.TextMatrix(2, 1)), "B", 80)
'''''        End If
'''''        Print #1, "From Date : " & FormName.TXTS_DATE & " To " & FormName.TXTE_DATE
'''''        Print #1, Chr(15) + PrintLine + Chr(18)
'''''        Print #1, Chr(15) + "Vr.Date    Vr No.             Particulars                                                          Debit             Credit" + Chr(18)
'''''        Print #1, Chr(15) + PrintLine + Chr(18)
'''''        mROW = mROW + 5
'''''    End If
'''''    If mDate <> RstRowSource!V_Date Then
'''''        mDR = 0
'''''        mCR = 0
'''''        mDate = RstRowSource!V_Date
'''''    End If
'''''    If mVrNo <> Trim(RstRowSource!v_TYPE) + "/" + Trim(CStr(RstRowSource!V_NO)) + "/" + Trim(RstRowSource!V_ADD) Then
'''''        mVrNo = Trim(RstRowSource!v_TYPE) + "/" + Trim(CStr(RstRowSource!V_NO)) + "/" + Trim(RstRowSource!V_ADD)
''''''        mVrNo = Trim(RstRowSource!DOCNO)
'''''        Print #1, Chr(15) + FaSetW(CStr(Format(RstRowSource!V_Date, "dd/MM/yy")), 10) + " " + FaSetW(RstRowSource!DOCNO, 18) + " " + FaSetW(FaXNull(RstRowSource!Name), 60) + "  " + FaSetN(FaBNull(RstRowSource!DEBIT), 12) + "       " + FaSetN(FaBNull(RstRowSource!CREDIT), 12) + Chr(18)
'''''        mROW = mROW + 1
'''''        If mROW >= 60 Then mPAGE = mPAGE + 1: mROW = PageChgJournal(FormName, PrintLine, mPAGE, mROW, mDrSum, mCrSum)
'''''        If FaXNull(Trim(RstRowSource!Chq_No)) <> "" Then mChq = mChq + "Ch.No:" + Trim(FaXNull(RstRowSource!Chq_No)) + " Ch.Dt: " + CStr(Format(FaXNull(RstRowSource!Chq_Date), "dd/MM/yy"))
'''''        mNarr = FaXNull(RstRowSource!mNarr)
'''''        Narr = FaXNull(RstRowSource!Narr)
'''''        mDR = mDR + FaSNull(RstRowSource!DEBIT)
'''''        mCR = mCR + FaSNull(RstRowSource!CREDIT)
'''''        mDrSum = mDrSum + FaSNull(RstRowSource!DEBIT)
'''''        mCrSum = mCrSum + FaSNull(RstRowSource!CREDIT)
'''''        If Len(Trim(Narr)) > 0 Then mSepNarr = True
'''''        If Len(Trim(mNarr)) > 0 Then mComNarr = True
'''''    Else
'''''        If mSepNarr = False Then
'''''            Print #1, Chr(15) + Space(30) + FaSetW(FaXNull(RstRowSource!Name), 60) + "  " + FaSetN(FaBNull(RstRowSource!DEBIT), 12) + "       " + FaSetN(FaBNull(RstRowSource!CREDIT), 12) + Chr(18)
'''''            mROW = mROW + 1
'''''            If mROW >= 58 Then mPAGE = mPAGE + 1: mROW = PageChgJournal(FormName, PrintLine, mPAGE, mROW, mDrSum, mCrSum)
'''''            Narr = FaXNull(RstRowSource!Narr)
'''''            mDR = mDR + FaSNull(RstRowSource!DEBIT)
'''''            mCR = mCR + FaSNull(RstRowSource!CREDIT)
'''''            mDrSum = mDrSum + FaSNull(RstRowSource!DEBIT)
'''''            mCrSum = mCrSum + FaSNull(RstRowSource!CREDIT)
'''''            If Len(Trim(Narr)) > 0 Then mSepNarr = True
'''''        End If
'''''    End If
'''''    If Len(Trim(Narr)) > 0 Then
'''''        Print #1, Chr(15) + Space(30) + FaSetW(FaXNull(left(Narr, 60)), 60) + Chr(18)
'''''        mROW = mROW + 1
'''''        Narr = Mid(Narr, 61, 250)
'''''        If mROW >= 58 Then mPAGE = mPAGE + 1: mROW = PageChgJournal(FormName, PrintLine, mPAGE, mROW, mDrSum, mCrSum)
'''''    Else
'''''        mSepNarr = False
'''''        RstRowSource.MoveNext
'''''    End If
'''''    If RstRowSource.EOF = True Then
'''''        GoTo LastMPrint
'''''    Else
'''''        If mVrNo <> Trim(RstRowSource!v_TYPE) + "/" + Trim(CStr(RstRowSource!V_NO)) + "/" + Trim(RstRowSource!V_ADD) Then
'''''LastMPrint:
'''''            If Len(mChq) > 0 Then
'''''                Print #1, Chr(15) + Space(30) + FaSetW(FaXNull(mChq), 60) + Chr(18)
'''''                mChq = ""
'''''                mROW = mROW + 1
'''''            End If
'''''            Do Until mComNarr = False
'''''                If Len(Trim(mNarr)) > 0 Then
'''''                    If mROW >= 58 Then mPAGE = mPAGE + 1: mROW = PageChgJournal(FormName, PrintLine, mPAGE, mROW, mDrSum, mCrSum)
'''''                    Print #1, Chr(15) + Space(30) + FaSetW(FaXNull(left(mNarr, 60)), 60) + Chr(18)
'''''                    mNarr = Mid(mNarr, 61, 250)
'''''                    mROW = mROW + 1
'''''                Else
'''''                    mComNarr = False
'''''                End If
'''''            Loop
'''''            If RstRowSource.EOF = True Then GoTo LastPrint
'''''        End If
'''''        If mDate <> RstRowSource!V_Date Then
'''''LastPrint:
'''''            If DayTot = "Yes" Then
'''''                If mROW >= 58 Then mPAGE = mPAGE + 1: mROW = PageChgJournal(FormName, PrintLine, mPAGE, mROW, mDrSum, mCrSum)
'''''                Print #1, Chr(15) + Space(92) + String(31, "-") + Chr(18)
'''''                Print #1, Chr(15) + FaSetW("Total", 92) + FaSetN(FaBNull(mDR), 12) + "       " + FaSetN(FaBNull(mCR), 12) + Chr(18)
'''''                Print #1, Chr(15) + Space(92) + String(31, "-") + Chr(18)
'''''                mROW = mROW + 3
'''''            End If
'''''        End If
'''''    End If
'''''    If mROW >= 60 Then
'''''        Print #1, Chr(15) + PrintLine + Chr(18)
'''''        Print #1, Chr(15) + FaSetW("Total", 92) + FaSetN(FaBNull(mDrSum), 12) + "       " + FaSetN(FaBNull(mCrSum), 12) + Chr(18)
'''''        Print #1, Chr(15) + PrintLine + Chr(18)
'''''        Print #1, Chr(12)
'''''        mPAGE = mPAGE + 1
'''''        mROW = 0
'''''    End If
'''''Loop
'''''Print #1, Chr(15) + PrintLine + Chr(18)
'''''Print #1, Chr(15) + FaSetW("Total", 92) + FaSetN(FaBNull(mDrSum), 12) + "       " + FaSetN(FaBNull(mCrSum), 12) + Chr(18)
'''''Print #1, Chr(15) + PrintLine + Chr(18)
'''''Print #1, Chr(12)
'''''Close #1
'''''PrintOnDosPort
'''''Set RstEnvi = Nothing
'''''Exit Sub
'''''ELoop:    CheckError
'''''End Sub
'''''Public Function FaPRN_TIT(st1 As String, mFont As String, LNT As Integer) As String
'''''Dim LEN1 As Integer, WDT
'''''FaPRN_TIT = ""
'''''st1 = Trim(st1)
'''''LEN1 = Len(st1)
'''''Select Case mFont
'''''    Case "A"
'''''        WDT = Int(LNT / 2)
'''''        FaPRN_TIT = Chr(18) + Chr(14) + Chr(27) + "G" + Space((WDT - LEN1) / 2) + st1 + Chr(27) + "H"
'''''    Case "B"
'''''        WDT = Int(LNT * 5 / 6)
'''''        FaPRN_TIT = Chr(14) + Chr(15) + Space((WDT - LEN1) / 2) + st1 + Chr(18)
'''''    Case "C"
'''''        WDT = LNT
'''''        FaPRN_TIT = Chr(18) + Space((WDT - LEN1) / 2) + st1
'''''End Select
'''''End Function
'''''Public Function FaPRNTIT(PageNo As Integer, RowNo As Integer, Paper As Integer) As Integer
'''''Dim PrnStr As String, PrnStr1 As String, PrnStr2 As String, PrnStr3 As String, PrnStr4 As String
'''''PrnStr = ""
'''''PrnStr1 = ""
'''''PrnStr2 = ""
'''''PrnStr3 = ""
'''''PrnStr4 = ""
'''''If Dater = "Y" Then PrnStr = PrnStr + "DATE " + CStr(Format(PubLoginDate, "dd/MM/yy"))
'''''If Pager = "Y" Then PrnStr = PrnStr + Space(Paper - IIf(Dater = "Y", Len("DATE " + CStr(Format(PubLoginDate, "dd/MM/yy"))), 0) - Len(" PAGE NO " + Trim(STR(PageNo)))) + " PAGE NO " + Trim(STR(PageNo))
'''''Select Case Titler
'''''    Case "L"
'''''        If Len(PubComp_Name) <= 40 Then
'''''            PrnStr1 = PrnStr1 + Chr(18) + Chr(14) + Chr(27) + Chr(69) + PubComp_Name + Chr(27) + Chr(70) + Chr(14)
'''''        Else
'''''            PrnStr1 = PrnStr1 + Chr(18) + Chr(14) + Chr(15) + PubComp_Name + Chr(18)
'''''        End If
'''''        If Len(PubComp_Add) > 0 Then PrnStr2 = Chr(14) + Chr(15) + PubComp_Add + Chr(18)
'''''        If Len(PubComp_Add2) > 0 Then PrnStr3 = Chr(14) + Chr(15) + PubComp_Add2 + Chr(18)
'''''        If Len(PubComp_City) > 0 Then PrnStr4 = Chr(14) + Chr(15) + PubComp_City + Chr(18)
'''''    Case "M"
'''''        If Len(PubComp_Name) <= 40 Then
'''''            PrnStr1 = PrnStr1 + FaPRN_TIT(PubComp_Name, "A", Val(Paper))
'''''        Else
'''''            PrnStr1 = PrnStr1 + FaPRN_TIT(PubComp_Name, "B", Val(Paper))
'''''        End If
'''''        If Len(PubComp_Add) > 0 Then PrnStr2 = PrnStr2 + FaPRN_TIT(PubComp_Add, "B", Val(Paper))
'''''        If Len(PubComp_Add2) > 0 Then PrnStr3 = PrnStr3 + FaPRN_TIT(PubComp_Add2, "B", Val(Paper))
'''''        If Len(PubComp_City) > 0 Then PrnStr4 = PrnStr4 + FaPRN_TIT(PubComp_City, "B", Val(Paper))
'''''End Select
'''''If Len(PrnStr) > 0 Then Print #1, PrnStr: RowNo = RowNo + 1
'''''If Len(PrnStr1) > 0 Then Print #1, PrnStr1: RowNo = RowNo + 1
'''''If Len(PrnStr2) > 0 Then Print #1, PrnStr2: RowNo = RowNo + 1
'''''If Len(PrnStr3) > 0 Then Print #1, PrnStr3: RowNo = RowNo + 1
'''''If Len(PrnStr4) > 0 Then Print #1, PrnStr4: RowNo = RowNo + 1
'''''FaPRNTIT = RowNo
'''''End Function
'''''Private Function PageChgJournal(FormName As Object, PrintLine As String, mPAGE As Integer, mROW As Integer, mDrSum As Double, mCrSum As Double) As Integer
'''''    Print #1, Chr(15) + PrintLine + Chr(18)
'''''    Print #1, Chr(15) + FaSetW("Total", 92) + FaSetN(FaBNull(mDrSum), 12) + "       " + FaSetN(FaBNull(mCrSum), 12) + Chr(18)
'''''    Print #1, Chr(15) + PrintLine + Chr(18)
'''''    Print #1, Chr(12)
'''''    mROW = 0
'''''    mROW = FaPRNTIT(mPAGE, mROW, 80)
'''''    If PubFaSiteType <> 0 And FormName.FGrid.TextMatrix(4, 1) <> "" Then
'''''        Print #1, FaPRN_TIT(Trim(FormName.FGrid.TextMatrix(2, 1)) + " (" + FormName.FGrid.TextMatrix(4, 1) + ")", "B", 80)
'''''    Else
'''''        Print #1, FaPRN_TIT(Trim(FormName.FGrid.TextMatrix(2, 1)), "B", 80)
'''''    End If
'''''    Print #1, "From Date : " & FormName.TXTS_DATE & " To " & FormName.TXTE_DATE
'''''    Print #1, Chr(15) + PrintLine + Chr(18)
'''''    Print #1, Chr(15) + "Vr.Date    Vr No.             Particulars                                                          Debit             Credit" + Chr(18)
'''''    Print #1, Chr(15) + PrintLine + Chr(18)
'''''    mROW = mROW + 5
'''''    PageChgJournal = mROW
'''''End Function
'''''Private Sub PrintOnDosPort()
'''''Dim connectionId
'''''Open "C:\reptmp\SBill.BAT" For Output As #1
'''''Print #1, "TYPE C:\reptmp\TEXTFILE.TXT>" + PubFaDosPort
'''''Close #1
'''''If PubRunPIF = "Y" Then
'''''    connectionId = Shell("C:\reptmp\SBill.PIF", vbHide)
'''''Else
'''''    connectionId = Shell("C:\reptmp\SBill.BAT", vbHide)
'''''End If
'''''End Sub
'''''Public Sub CheckError()
'''''If err.NUMBER <> 0 Then
'''''    MsgBox err.Description, vbInformation, "Validation"
'''''End If
'''''End Sub
