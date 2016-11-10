VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmSyCtrlVeh 
   Appearance      =   0  'Flat
   BackColor       =   &H00CFE0E0&
   Caption         =   "Vehicle A/c Control Declaration"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   630
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DGGrp 
      Height          =   3330
      Left            =   -30
      Negotiate       =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   8010
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   5874
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "Name"
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
      BeginProperty Column01 
         DataField       =   "Code"
         Caption         =   "Group Code"
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
         DataField       =   "GroupNature"
         Caption         =   "GroupNature"
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
         DataField       =   "MainGrCode"
         Caption         =   "MainGrCode"
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
         DataField       =   "GroupLevel"
         Caption         =   "GroupLevel"
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
      BeginProperty Column05 
         DataField       =   "CurrentCount"
         Caption         =   "CurrentCount"
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
      BeginProperty Column06 
         DataField       =   "CurrentBalance"
         Caption         =   "CurrentBalance"
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
      BeginProperty Column07 
         DataField       =   "SubLedYN"
         Caption         =   "SubLedYN"
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
      BeginProperty Column08 
         DataField       =   "AliasYN"
         Caption         =   "AliasYN"
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
      BeginProperty Column09 
         DataField       =   "GroupHelp"
         Caption         =   "GroupHelp"
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
      BeginProperty Column10 
         DataField       =   "Nature"
         Caption         =   "Nature"
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
            ColumnWidth     =   5040
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGGodown 
      Height          =   3330
      Left            =   45
      Negotiate       =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   7545
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5874
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Godown Name"
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
         DataField       =   "Code"
         Caption         =   "GodownCode"
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
            ColumnWidth     =   2789.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5490
      Left            =   60
      TabIndex        =   1
      Top             =   405
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9684
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   15718112
      TabCaption(0)   =   "1. General Settings"
      TabPicture(0)   =   "frmSyCtrlVeh.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "2. A/c Settings"
      TabPicture(1)   =   "frmSyCtrlVeh.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Lbl(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Lbl(9)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Lbl(10)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Lbl(23)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Lbl(17)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Lbl(14)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Lbl(12)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Lbl(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Lbl(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Lbl(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Lbl(4)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Lbl(5)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Lbl(6)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Lbl(7)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Lbl(8)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Lbl(11)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Lbl(13)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Lbl(15)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Lbl(16)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Txt(3)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Txt(2)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Txt(0)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Txt(4)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Txt(5)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Txt(6)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Txt(1)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Txt(7)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Txt(8)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Txt(9)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Txt(10)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Txt(24)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Txt(27)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Txt(28)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Txt(32)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Txt(33)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Txt(34)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Txt(35)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Txt(37)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).ControlCount=   38
      Begin VB.TextBox Txt 
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
         Height          =   225
         Index           =   37
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   106
         Top             =   3915
         Width           =   4785
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   210
         Index           =   35
         Left            =   6750
         MaxLength       =   50
         TabIndex        =   32
         Top             =   4845
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   210
         Index           =   34
         Left            =   6750
         MaxLength       =   50
         TabIndex        =   31
         Top             =   4605
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   33
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   30
         Top             =   3675
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   32
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   29
         Top             =   3435
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   28
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   28
         Top             =   3195
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   27
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   27
         Top             =   2955
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   24
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   25
         Top             =   2475
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   10
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   26
         Top             =   2715
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   9
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   24
         Top             =   2235
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   8
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   23
         Top             =   1995
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   7
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1755
         Width           =   4785
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00BAD3C9&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   5010
         Left            =   -74955
         TabIndex        =   42
         Top             =   390
         Width           =   11670
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Index           =   36
            Left            =   5940
            MaxLength       =   5
            TabIndex        =   104
            Text            =   "999"
            Top             =   1275
            Width           =   1050
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            Height          =   210
            Index           =   31
            Left            =   2655
            MaxLength       =   3
            TabIndex        =   98
            Text            =   "Yes/No"
            Top             =   1755
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox Txt 
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
            Height          =   210
            Index           =   30
            Left            =   2670
            MaxLength       =   3
            TabIndex        =   96
            Text            =   "Yes/No"
            Top             =   1515
            Width           =   600
         End
         Begin VB.TextBox Txt 
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
            Height          =   210
            Index           =   29
            Left            =   2655
            MaxLength       =   3
            TabIndex        =   94
            Text            =   "Yes/No"
            Top             =   1275
            Width           =   600
         End
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Index           =   17
            Left            =   6390
            MaxLength       =   3
            TabIndex        =   87
            Text            =   "999"
            Top             =   1020
            Width           =   600
         End
         Begin VB.TextBox Txt 
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
            Height          =   210
            Index           =   26
            Left            =   10890
            MaxLength       =   3
            TabIndex        =   85
            Text            =   "Yes/No"
            Top             =   1035
            Width           =   600
         End
         Begin VB.TextBox Txt 
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
            Height          =   210
            Index           =   25
            Left            =   6390
            MaxLength       =   3
            TabIndex        =   84
            Text            =   "Yes/No"
            Top             =   780
            Width           =   600
         End
         Begin VB.TextBox Txt 
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
            Height          =   210
            Index           =   23
            Left            =   2655
            MaxLength       =   3
            TabIndex        =   10
            Text            =   "Yes/No"
            Top             =   1035
            Width           =   600
         End
         Begin VB.TextBox Txt 
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
            Height          =   210
            Index           =   22
            Left            =   4755
            TabIndex        =   3
            Text            =   "Default Godown"
            Top             =   300
            Width           =   2235
         End
         Begin VB.TextBox Txt 
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
            Height          =   225
            Index           =   18
            Left            =   10905
            MaxLength       =   3
            TabIndex        =   9
            Text            =   "Yes/N"
            Top             =   780
            Width           =   585
         End
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
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
            Height          =   210
            Index           =   13
            Left            =   2655
            MaxLength       =   3
            TabIndex        =   5
            Text            =   "999"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox Txt 
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
            Height          =   210
            Index           =   15
            Left            =   8310
            MaxLength       =   40
            TabIndex        =   7
            Text            =   "X(40)"
            Top             =   540
            Width           =   3180
         End
         Begin VB.TextBox Txt 
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
            Height          =   210
            Index           =   11
            Left            =   1845
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "X(10)"
            Top             =   300
            Width           =   1410
         End
         Begin VB.TextBox Txt 
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
            Height          =   210
            Index           =   14
            Left            =   6390
            MaxLength       =   3
            TabIndex        =   6
            Text            =   "Yes/No"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox Txt 
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
            Height          =   210
            Index           =   12
            Left            =   8310
            MaxLength       =   20
            TabIndex        =   4
            Text            =   "012345678901234567890"
            Top             =   300
            Width           =   3180
         End
         Begin VB.TextBox Txt 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Index           =   16
            Left            =   2655
            MaxLength       =   5
            TabIndex        =   8
            Text            =   "99.99"
            Top             =   780
            Width           =   600
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   2025
            Left            =   30
            TabIndex        =   11
            Top             =   2535
            Width           =   11610
            _ExtentX        =   20479
            _ExtentY        =   3572
            _Version        =   393216
            TabHeight       =   520
            ShowFocusRect   =   0   'False
            BackColor       =   12632064
            ForeColor       =   8388736
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "1. Quotation"
            TabPicture(0)   =   "frmSyCtrlVeh.frx":0038
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Txt(19)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "2. Booking"
            TabPicture(1)   =   "frmSyCtrlVeh.frx":0054
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Txt(20)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "3. Invoice"
            TabPicture(2)   =   "frmSyCtrlVeh.frx":0070
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Txt(21)"
            Tab(2).ControlCount=   1
            Begin VB.TextBox Txt 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1440
               Index           =   21
               Left            =   -74910
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               Top             =   465
               Width           =   11400
            End
            Begin VB.TextBox Txt 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1440
               Index           =   20
               Left            =   -74910
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   13
               Top             =   465
               Width           =   11400
            End
            Begin VB.TextBox Txt 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1440
               Index           =   19
               Left            =   90
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               Top             =   465
               Width           =   11400
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   44
               Left            =   -74145
               TabIndex        =   72
               Top             =   1560
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Header"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   54
               Left            =   -74775
               TabIndex        =   71
               Top             =   1530
               Width           =   615
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   43
               Left            =   -74145
               TabIndex        =   70
               Top             =   1095
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Middle"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   53
               Left            =   -74775
               TabIndex        =   69
               Top             =   1065
               Width           =   540
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   42
               Left            =   -74145
               TabIndex        =   68
               Top             =   615
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Header"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   52
               Left            =   -74805
               TabIndex        =   67
               Top             =   615
               Width           =   615
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Header"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   51
               Left            =   -74790
               TabIndex        =   66
               Top             =   510
               Width           =   615
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   41
               Left            =   -74130
               TabIndex        =   65
               Top             =   510
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Middle"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   50
               Left            =   -74760
               TabIndex        =   64
               Top             =   960
               Width           =   540
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   40
               Left            =   -74130
               TabIndex        =   63
               Top             =   990
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Header"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   49
               Left            =   -74760
               TabIndex        =   62
               Top             =   1425
               Width           =   615
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   39
               Left            =   -74130
               TabIndex        =   61
               Top             =   1455
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Header"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   48
               Left            =   -74640
               TabIndex        =   60
               Top             =   555
               Width           =   615
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   38
               Left            =   -73980
               TabIndex        =   59
               Top             =   555
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Middle"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   47
               Left            =   -74610
               TabIndex        =   58
               Top             =   1005
               Width           =   540
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   37
               Left            =   -73980
               TabIndex        =   57
               Top             =   1035
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Header"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   46
               Left            =   -74610
               TabIndex        =   56
               Top             =   1470
               Width           =   615
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   36
               Left            =   -73980
               TabIndex        =   55
               Top             =   1500
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Header"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   42
               Left            =   -74775
               TabIndex        =   54
               Top             =   480
               Width           =   615
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   32
               Left            =   -74115
               TabIndex        =   53
               Top             =   480
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Middle"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   41
               Left            =   -74745
               TabIndex        =   52
               Top             =   930
               Width           =   540
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   31
               Left            =   -74115
               TabIndex        =   51
               Top             =   960
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Header"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   40
               Left            =   -74745
               TabIndex        =   50
               Top             =   1395
               Width           =   615
            End
            Begin VB.Label LblColon 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   30
               Left            =   -74115
               TabIndex        =   49
               Top             =   1425
               Width           =   90
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Invoice Prefix :"
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
            Left            =   3375
            TabIndex        =   105
            Top             =   1290
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Octrai Saperately"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   225
            TabIndex        =   99
            Top             =   1770
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Registration Fee"
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
            Left            =   240
            TabIndex        =   97
            Top             =   1530
            Width           =   1800
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Post Insurance Fee"
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
            Left            =   225
            TabIndex        =   95
            Top             =   1290
            Width           =   1635
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mfg. Invoice No. on Sale Invoice :"
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
            Index           =   33
            Left            =   3375
            TabIndex        =   91
            Top             =   540
            Width           =   2925
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rebate Days :"
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
            Left            =   3375
            TabIndex        =   90
            Top             =   1035
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Godown :"
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
            Index           =   113
            Left            =   3375
            TabIndex        =   89
            Top             =   300
            Width           =   1500
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Rate Inclusive Tax:"
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
            Left            =   3390
            TabIndex        =   88
            Top             =   780
            Width           =   2340
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show Debtors In Supplier Help"
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
            Left            =   7035
            TabIndex        =   86
            Top             =   1050
            Width           =   2640
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A/c Posting By All Users :"
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
            Left            =   255
            TabIndex        =   82
            Top             =   1050
            Width           =   2190
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Detail on Sale Invoice :"
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
            Left            =   7035
            TabIndex        =   74
            Top             =   780
            Width           =   2400
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Document Footers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   55
            Left            =   75
            TabIndex        =   48
            Top             =   2250
            Width           =   1530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quotation Validity (Days) :"
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
            Index           =   37
            Left            =   255
            TabIndex        =   47
            Top             =   540
            Width           =   2295
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Own Financer :"
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
            Index           =   36
            Left            =   7035
            TabIndex        =   46
            Top             =   540
            Width           =   1290
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RSO Code :"
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
            Index           =   35
            Left            =   255
            TabIndex        =   45
            Top             =   300
            Width           =   1020
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Visit Objective :"
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
            Index           =   32
            Left            =   7035
            TabIndex        =   44
            Top             =   300
            Width           =   1365
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delay Intt. Rate % :"
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
            Index           =   112
            Left            =   255
            TabIndex        =   43
            Top             =   780
            Width           =   1740
         End
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   1
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   16
         ToolTipText     =   "Press L-> Local or C-> Central"
         Top             =   795
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   6
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1515
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   5
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   20
         ToolTipText     =   "Press L-> Local or C-> Central"
         Top             =   1275
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   4
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1035
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   0
         Left            =   6765
         MaxLength       =   50
         TabIndex        =   15
         ToolTipText     =   "Press L-> Local or C-> Central"
         Top             =   5085
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   2
         Left            =   6765
         MaxLength       =   50
         TabIndex        =   17
         Top             =   5325
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.TextBox Txt 
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
         Height          =   210
         Index           =   3
         Left            =   6765
         MaxLength       =   50
         TabIndex        =   18
         Top             =   5565
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Special Discount A/c"
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
         Index           =   16
         Left            =   450
         TabIndex        =   107
         Top             =   3900
         Width           =   1755
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subvention Claim A/c"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   4350
         TabIndex        =   103
         Top             =   4830
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subvention A/c"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   4350
         TabIndex        =   102
         Top             =   4590
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Indirect Exepences A/c"
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
         Index           =   11
         Left            =   450
         TabIndex        =   101
         Top             =   3660
         Width           =   1980
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Octrai A/c"
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
         Left            =   435
         TabIndex        =   100
         Top             =   3420
         Width           =   855
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance A/c"
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
         Left            =   435
         TabIndex        =   93
         Top             =   3180
         Width           =   1200
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration Fee A/c"
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
         Left            =   435
         TabIndex        =   92
         Top             =   2940
         Width           =   1725
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turn Over Tax A/c Name"
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
         Left            =   435
         TabIndex        =   83
         Top             =   2460
         Width           =   2145
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Round Off A/c Name"
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
         Left            =   435
         TabIndex        =   78
         Top             =   2700
         Width           =   1755
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TDS A/c Name"
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
         Left            =   435
         TabIndex        =   77
         Top             =   2220
         Width           =   1260
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interest A/c Name"
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
         Left            =   435
         TabIndex        =   76
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Charge A/c Name"
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
         Left            =   435
         TabIndex        =   75
         Top             =   1755
         Width           =   2235
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fitment A/c Name"
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
         Index           =   12
         Left            =   435
         TabIndex        =   41
         Top             =   1005
         Width           =   1530
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel A/c Name"
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
         Index           =   14
         Left            =   435
         TabIndex        =   40
         Top             =   1260
         Width           =   1245
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank A/c Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   17
         Left            =   4350
         TabIndex        =   39
         Top             =   5565
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stamp Duty A/c Name"
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
         Left            =   435
         TabIndex        =   38
         Top             =   1500
         Width           =   1920
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash A/c Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   10
         Left            =   4350
         TabIndex        =   37
         Top             =   5325
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sundry Creditors A/c Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   9
         Left            =   4350
         TabIndex        =   36
         Top             =   5085
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sundry Debtors A/c Group"
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
         Left            =   435
         TabIndex        =   35
         ToolTipText     =   "Press L-> Local or C-> Central"
         Top             =   795
         Width           =   2280
      End
   End
   Begin MSDataGridLib.DataGrid DGAc 
      Height          =   3330
      Left            =   975
      Negotiate       =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   8340
      Visible         =   0   'False
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   5874
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "A/c Name"
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
         DataField       =   "Code"
         Caption         =   "Ac.Code"
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
            ColumnWidth     =   5220.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   661
   End
   Begin MSDataGridLib.DataGrid DGObj 
      Height          =   4515
      Left            =   4200
      Negotiate       =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   7980
      Visible         =   0   'False
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   7964
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
         Caption         =   "Objective"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGFin 
      Height          =   4515
      Left            =   6300
      Negotiate       =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   8130
      Visible         =   0   'False
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   7964
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   16777215
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
         Caption         =   "Financer Name"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RSO Code :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   2
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmSyCtrlVeh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'FA Connection for Works FAData
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim rsGrp As ADODB.Recordset
Dim rsAc As ADODB.Recordset
Dim rsObj As ADODB.Recordset
Dim rsFin As ADODB.Recordset
Dim RsGodown As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim Syctrl As ADODB.Recordset
Dim ExitCtrl As Boolean
'grid color scheme
Private Const CellBackColLeave As String = &HC8E8DA
Private Const CellForeColLeave As String = &H0&
Private Const CellBackColEnter As String = &HC0E0FF
Private Const GridBackColorBkg As String = &HBAD3C9

Dim MyIndex As Byte
Private Const VehCreGrp As Byte = 0
Private Const VehDebGrp As Byte = 1
Private Const VehCashAc As Byte = 2
Private Const VehBankAc As Byte = 3
Private Const FitmentAc As Byte = 4
Private Const FuelAc As Byte = 5
Private Const StampDutyAc As Byte = 6
Private Const ServiceChrgAc As Byte = 7
Private Const InterestAc As Byte = 8
Private Const TDSAc As Byte = 9
Private Const TOTAc As Byte = 24
Private Const VehRoffAc As Byte = 10

Private Const RSO_Code As Byte = 11
Private Const VisitObjCode As Byte = 12
Private Const Valid_Day As Byte = 13
Private Const SupInvOnVehSaleInv As Byte = 14
Private Const OwnFinCode As Byte = 15
Private Const DelayInttRate As Byte = 16
Private Const RebDays As Byte = 17
Private Const TaxDetOnVehInv As Byte = 18

Private Const VehQuotFooter As Byte = 19
Private Const VehBookFooter As Byte = 20
Private Const VehSaleInvFooter      As Byte = 21
Private Const VehGodown             As Byte = 22
Private Const AcPostingByAllUser    As Byte = 23
Private Const VehRateInclTax        As Byte = 25
Private Const DebtorInSupplierHelp  As Byte = 26
Private Const RegnFeeAc             As Byte = 27
Private Const InsuranceFeeAc        As Byte = 28
Private Const PostRegnFeeYn         As Byte = 29
Private Const PostInsuranceFeeYn    As Byte = 30
Private Const PostOctraiSaperatelyYn As Byte = 31
Private Const OctraiAc              As Byte = 32
Private Const IndirectExpAc         As Byte = 33
Private Const SubventionAc         As Byte = 34
Private Const SubventionClaimAc         As Byte = 35
Private Const VehTaxInvPrefix         As Byte = 36
Private Const SpecialDiscountAc         As Byte = 37

'Private Const VehDelivChFooter As Byte = 29
'Private Const VehSaleCertiFooter As Byte = 30
'--EOF of Vehicle Section


Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next
'    Txt(SprMoneyRectFooter).Enabled = True
    Txt(VehQuotFooter).Enabled = True
    Txt(VehBookFooter).Enabled = True
    Txt(VehSaleInvFooter).Enabled = True
'    Txt(SprMoneyRectFooter).Locked = True
    Txt(VehQuotFooter).Locked = True
    Txt(VehBookFooter).Locked = True
    Txt(VehSaleInvFooter).Locked = True
End Sub

'* Used for intialize grid columns
Private Sub Grid_Ini()
    DGObj.left = Me.left: DGObj.top = mTopScale
    DGFin.left = Me.left: DGFin.top = mTopScale
    DGGrp.left = Me.width - (DGGrp.width + mRtScale): DGGrp.top = mTopScale
    DGAc.left = Me.width - (DGAc.width + mRtScale): DGAc.top = mTopScale
End Sub

Private Sub Grid_Hide()
    If DGObj.Visible Then DGObj.Visible = False
    If DGFin.Visible Then DGFin.Visible = False
    If DGAc.Visible = True Then DGAc.Visible = False
    If DGGrp.Visible = True Then DGGrp.Visible = False
End Sub

Private Sub MoveRec()
Dim Master1 As ADODB.Recordset
On Error GoTo ELoop
SSTab1.Tab = 0
SSTab2.Tab = 0
'General Settings
Txt(RSO_Code) = IIf(IsNull(Syctrl!RSO_Code), "", Syctrl!RSO_Code)
If Syctrl!VehGodown <> "" Then
    RsGodown.MoveFirst
    RsGodown.FIND ("Code ='" & Syctrl!VehGodown & "'")
    Txt(VehGodown).Tag = Syctrl!VehGodown
    Txt(VehGodown) = IIf(RsGodown.EOF, "", RsGodown!Name)
Else
    Txt(VehGodown) = ""
    Txt(VehGodown).Tag = ""
End If
If RsGodown.RecordCount > 0 And RsGodown.EOF Then RsGodown.MoveFirst

If Syctrl!VisitObjCode <> "" Then
    rsObj.MoveFirst
    rsObj.FIND ("Code ='" & Syctrl!VisitObjCode & "'")
    Txt(VisitObjCode).Tag = Syctrl!VisitObjCode
    Txt(VisitObjCode) = IIf(rsObj.EOF, "", rsObj!Name)
Else
    Txt(VisitObjCode) = ""
    Txt(VisitObjCode).Tag = ""
End If
If rsObj.RecordCount > 0 And rsObj.EOF Then rsObj.MoveFirst

Txt(Valid_Day) = IIf(IsNull(Syctrl!Valid_Day) Or Syctrl!Valid_Day = 0, "", Syctrl!Valid_Day)
If Syctrl!OwnFinCode <> "" Then
    rsObj.MoveFirst
    rsObj.FIND ("Code ='" & Syctrl!OwnFinCode & "'")
    Txt(OwnFinCode).Tag = Syctrl!OwnFinCode
    Txt(OwnFinCode) = IIf(rsObj.EOF, "", rsObj!Name)
Else
    Txt(OwnFinCode) = ""
    Txt(OwnFinCode).Tag = ""
End If
If rsObj.RecordCount > 0 And rsObj.EOF Then rsObj.MoveFirst

Txt(DelayInttRate) = IIf(IsNull(Syctrl!DelayInttRate) Or Syctrl!DelayInttRate = 0, "", Syctrl!DelayInttRate)
Txt(RebDays) = IIf(IsNull(Syctrl!RebDays) Or Syctrl!RebDays = 0, "", Syctrl!RebDays)
'Txt(VehSaleInv_Prefix) = IIf(IsNull(Syctrl!VehSaleInv_Prefix), "", Syctrl!VehSaleInv_Prefix)
Txt(SupInvOnVehSaleInv) = IIf(Syctrl!SupInvOnVehSaleInv = 1, "Yes", "No")
Txt(DebtorInSupplierHelp) = IIf(Syctrl!DebtorInSupplierHelp = 1, "Yes", "No")
Txt(PostRegnFeeYn) = IIf(Syctrl!PostRegnFeeYn = 1, "Yes", "No")
Txt(PostInsuranceFeeYn) = IIf(Syctrl!PostInsuranceFeeYn = 1, "Yes", "No")
Txt(PostOctraiSaperatelyYn) = IIf(Syctrl!PostOctraiSaperatelyYn = 1, "Yes", "No")


Txt(TaxDetOnVehInv) = IIf(Syctrl!TaxDetOnVehInv = 1, "Yes", "No")
Txt(AcPostingByAllUser) = IIf(Syctrl!AcPostingByAllUser = 1, "Yes", "No")
Txt(VehRateInclTax) = IIf(Syctrl!VehRateIncTax = 1, "Yes", "No")
Txt(VehTaxInvPrefix) = XNull(Syctrl!VehTaxInvPrefix)

Txt(VehQuotFooter) = IIf(IsNull(Syctrl!VehQuotFooter), "", Syctrl!VehQuotFooter)
Txt(VehBookFooter) = IIf(IsNull(Syctrl!VehBookFooter), "", Syctrl!VehBookFooter)
Txt(VehSaleInvFooter) = IIf(IsNull(Syctrl!VehSaleInvFooter), "", Syctrl!VehSaleInvFooter)
'Txt(SprMoneyRectFooter) = IIf(IsNull(Syctrl!SprMoneyRectFooter), "", Syctrl!SprMoneyRectFooter)
'***
    Txt(VehCreGrp) = ""
    Txt(VehCreGrp).Tag = ""
    Txt(VehDebGrp) = ""
    Txt(VehDebGrp).Tag = ""
    Txt(VehCashAc) = ""
    Txt(VehCashAc).Tag = ""
    Txt(VehBankAc) = ""
    Txt(VehBankAc).Tag = ""
    Txt(FitmentAc) = ""
    Txt(FitmentAc).Tag = ""
    Txt(FuelAc) = ""
    Txt(FuelAc).Tag = ""
    Txt(StampDutyAc) = ""
    Txt(StampDutyAc).Tag = ""
    Txt(ServiceChrgAc) = ""
    Txt(ServiceChrgAc).Tag = ""
    Txt(InterestAc) = ""
    Txt(InterestAc).Tag = ""
    Txt(TDSAc) = ""
    Txt(TDSAc).Tag = ""
    Txt(TOTAc) = ""
    Txt(TOTAc).Tag = ""
    Txt(RegnFeeAc) = ""
    Txt(SpecialDiscountAc) = ""
    Txt(SpecialDiscountAc).Tag = ""
    Txt(RegnFeeAc).Tag = ""
    Txt(OctraiAc) = ""
    Txt(OctraiAc).Tag = ""
    Txt(IndirectExpAc) = ""
    Txt(IndirectExpAc).Tag = ""
    Txt(SubventionAc) = ""
    Txt(SubventionAc).Tag = ""
    Txt(SubventionClaimAc) = ""
    Txt(SubventionClaimAc).Tag = ""
    
    Txt(InsuranceFeeAc) = ""
    Txt(InsuranceFeeAc).Tag = ""
    Txt(VehRoffAc) = ""
    Txt(VehRoffAc).Tag = ""
'** A/c Section
If Master!VehCre_Grp <> Null Or Master!VehCre_Grp <> "" Then
    Txt(VehCreGrp) = GCnFaV.Execute("Select GroupName from AcGroup where GroupCode='" & Master!VehCre_Grp & "'").Fields(0).Value
    Txt(VehCreGrp).Tag = Master!VehCre_Grp
End If
If Master!VehDeb_Grp <> Null Or Master!VehDeb_Grp <> "" Then
    Txt(VehDebGrp) = GCnFaV.Execute("Select GroupName from AcGroup where GroupCode='" & Master!VehDeb_Grp & "'").Fields(0).Value
    Txt(VehDebGrp).Tag = Master!VehDeb_Grp
End If

Set Master1 = New Recordset
Master1.CursorLocation = adUseClient
Master1.Open "Select SubCode,Name from SubGroup Order by SubCode", GCnFaV, adOpenStatic, adLockReadOnly
If Master!VehCash_Ac <> Null Or Master!VehCash_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!VehCash_Ac & "'")
    Txt(VehCashAc) = Master1!Name
    Txt(VehCashAc).Tag = Master!VehCash_Ac
End If
If Master!VehBank_Ac <> Null Or Master!VehBank_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!VehBank_Ac & "'")
    Txt(VehBankAc) = Master1!Name
    Txt(VehBankAc).Tag = Master!VehBank_Ac
End If
If Master!Fitment_Ac <> Null Or Master!Fitment_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!Fitment_Ac & "'")
    Txt(FitmentAc) = Master1!Name
    Txt(FitmentAc).Tag = Master!Fitment_Ac
End If
If Master!Fuel_Ac <> Null Or Master!Fuel_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!Fuel_Ac & "'")
    Txt(FuelAc) = Master1!Name
    Txt(FuelAc).Tag = Master!Fuel_Ac
End If
If Master!StampDuty_Ac <> Null Or Master!StampDuty_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!StampDuty_Ac & "'")
    Txt(StampDutyAc) = Master1!Name
    Txt(StampDutyAc).Tag = Master!StampDuty_Ac
End If
If Master!ServiceChrg_Ac <> Null Or Master!ServiceChrg_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!ServiceChrg_Ac & "'")
    Txt(ServiceChrgAc) = Master1!Name
    Txt(ServiceChrgAc).Tag = Master!ServiceChrg_Ac
End If
If Master!Interest_Ac <> Null Or Master!Interest_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!Interest_Ac & "'")
    Txt(InterestAc) = Master1!Name
    Txt(InterestAc).Tag = Master!Interest_Ac
End If
If Master!TDS_Ac <> Null Or Master!TDS_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!TDS_Ac & "'")
    Txt(TDSAc) = Master1!Name
    Txt(TDSAc).Tag = Master!TDS_Ac
End If
If Master!TOTax_Ac <> Null Or Master!TOTax_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!TOTax_Ac & "'")
    Txt(TOTAc) = Master1!Name
    Txt(TOTAc).Tag = Master!TOTax_Ac
End If
If Master!RegnFeeAc <> Null Or Master!RegnFeeAc <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!RegnFeeAc & "'")
    Txt(RegnFeeAc) = Master1!Name
    Txt(RegnFeeAc).Tag = Master!RegnFeeAc
End If

If Master!SpecialDiscountAc <> Null Or Master!SpecialDiscountAc <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SpecialDiscountAc & "'")
    Txt(SpecialDiscountAc) = Master1!Name
    Txt(SpecialDiscountAc).Tag = Master!SpecialDiscountAc
End If


If Master!OctraiAc <> Null Or Master!OctraiAc <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!OctraiAc & "'")
    Txt(OctraiAc) = Master1!Name
    Txt(OctraiAc).Tag = Master!OctraiAc
End If

If Master!IndirectExpAc <> Null Or Master!IndirectExpAc <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!IndirectExpAc & "'")
    Txt(IndirectExpAc) = Master1!Name
    Txt(IndirectExpAc).Tag = Master!IndirectExpAc
End If


If Master!SubventionAc <> Null Or Master!SubventionAc <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SubventionAc & "'")
    Txt(SubventionAc) = Master1!Name
    Txt(SubventionAc).Tag = Master!SubventionAc
End If
If Master!SubventionClaimAc <> Null Or Master!SubventionClaimAc <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SubventionClaimAc & "'")
    Txt(SubventionClaimAc) = Master1!Name
    Txt(SubventionClaimAc).Tag = Master!SubventionClaimAc
End If


If Master!InsuranceFeeAc <> Null Or Master!InsuranceFeeAc <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!InsuranceFeeAc & "'")
    Txt(InsuranceFeeAc) = Master1!Name
    Txt(InsuranceFeeAc).Tag = Master!InsuranceFeeAc
End If

If Master!VehROff_Ac <> Null Or Master!VehROff_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!VehROff_Ac & "'")
    Txt(VehRoffAc) = Master1!Name
    Txt(VehRoffAc).Tag = Master!VehROff_Ac
End If
Set Master1 = Nothing
TopCtrl1.tAdd = False
TopCtrl1.tDel = False
TopCtrl1.tFirst = False
TopCtrl1.tPrev = False
TopCtrl1.tNext = False
TopCtrl1.tLast = False
TopCtrl1.tFind = False
TopCtrl1.tPrn = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGFin_Click()
On Error GoTo ELoop
    If rsFin.RecordCount > 0 Then
        Txt(MyIndex).TEXT = rsFin!Name
        Txt(MyIndex).Tag = rsFin!Code
    End If
    Txt(MyIndex).SetFocus
    DGFin.Visible = False
Exit Sub
ELoop:
    CheckError

End Sub

Private Sub DGGodown_Click()
On Error GoTo ELoop
    If RsGodown.RecordCount > 0 Then
        Txt(MyIndex).TEXT = RsGodown!Name
        Txt(MyIndex).Tag = RsGodown!Code
    End If
    Txt(MyIndex).SetFocus
    DGGodown.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub dgobj_Click()
On Error GoTo ELoop
    If rsObj.RecordCount > 0 Then
        Txt(MyIndex).TEXT = rsObj!Name
        Txt(MyIndex).Tag = rsObj!Code
    End If
    Txt(MyIndex).SetFocus
    DGObj.Visible = False
Exit Sub
ELoop:
    CheckError

End Sub

Private Sub DGAc_Click()
On Error GoTo ELoop
    If rsAc.RecordCount > 0 Then
        Txt(MyIndex).TEXT = rsAc!Name
        Txt(MyIndex).Tag = rsAc!Code
    End If
    Txt(MyIndex).SetFocus
    DGAc.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGGrp_Click()
On Error GoTo ELoop
    If rsGrp.RecordCount > 0 Then
        Txt(MyIndex).TEXT = rsGrp!Name
        Txt(MyIndex).Tag = rsGrp!Code
    End If
    Txt(MyIndex).SetFocus
    DGGrp.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
'FA Connection for Vehicle FAData
    TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini
    If PubSDTYN = 1 Then
        Lbl(5) = "S D T A/c Name"
    End If
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg '&HDFF4F2
        Txt(I).ForeColor = CtrlFColOrg
    Next
    Set RsGodown = New ADODB.Recordset
    RsGodown.CursorLocation = adUseClient
    RsGodown.Open "Select God_Code as Code, God_Name As Name From Godown where Appli_For=1 Order by God_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGodown.DataSource = RsGodown
    
    Set rsObj = New ADODB.Recordset
    rsObj.CursorLocation = adUseClient
    rsObj.Open "Select ObjCode as Code, ObjDesc As Name From VisitObjective Order by ObjDesc", GCn, adOpenDynamic, adLockOptimistic
    Set DGObj.DataSource = rsObj
    
    Set rsFin = New ADODB.Recordset
    rsFin.CursorLocation = adUseClient
    rsFin.Open "Select FinCode as Code, FinName As Name,Add1 From ContractFinance where FinCatg=0 Order by FinName,Add1", GCn, adOpenDynamic, adLockOptimistic
    Set DGFin.DataSource = rsFin
    
    Set rsGrp = New ADODB.Recordset
    rsGrp.CursorLocation = adUseClient
    rsGrp.Open "Select GroupCode As Code,GroupName As Name,GroupNature,MainGrCode From AcGroup Where MainGrCode <>'999' Order by GroupName", GCnFaV, adOpenDynamic, adLockOptimistic
    Set DGGrp.DataSource = rsGrp
    
    Set rsAc = New ADODB.Recordset
    rsAc.CursorLocation = adUseClient
    rsAc.Open "Select SubCode as Code,Name From SubGroup Order by Name", GCnFaV, adOpenDynamic, adLockOptimistic
    Set DGAc.DataSource = rsAc

    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    Master.Open "Select * from AcControls Where Div_Code = '" & PubDivCode & "'", GCnFaV, adOpenDynamic, adLockOptimistic
    'Set Master = GCnFaV.Execute("Select * from AcControls")
    If Master.RecordCount <= 0 Then
        Master.AddNew
        Master!Div_Code = PubDivCode
        Master.Update
    End If

    Set Syctrl = New ADODB.Recordset
    Syctrl.LockType = adLockOptimistic
    Syctrl.CursorLocation = adUseClient
    Syctrl.CursorType = adOpenDynamic
    Set Syctrl = GCn.Execute("Select * from Syctrl")
    
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If TopCtrl1.TopText2 <> "Browse" Then
    If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
        Exit Sub
    Else
        Cancel = 1
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsObj = Nothing
    Set rsFin = Nothing
    Set rsGrp = Nothing
    Set rsAc = Nothing
    Set Master = Nothing
    Set Syctrl = Nothing
    Set RsGodown = Nothing
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
'    Txt(SprMoneyRectFooter).Locked = False
    Txt(VehQuotFooter).Locked = False
    Txt(VehBookFooter).Locked = False
    Txt(VehSaleInvFooter).Locked = False

    Txt(RSO_Code).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eRef()
On Error GoTo ELoop
    rsGrp.Requery
    rsAc.Requery
    rsObj.Requery
    Master.Requery
    Syctrl.Requery
    RsGodown.Requery
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eSave()
Dim mTrans As Boolean, MasterSql$
On Error GoTo ELoop

'Apply necessary validations
'If IsValid(Txt(Party), "Party Name") = False Then Exit Sub
'**********
    GCnFaV.BeginTrans
    GCn.BeginTrans
        mTrans = True
        If TopCtrl1.TopText2 = "Edit" Then   'Edit
            GSQL = "update Syctrl set AcPostingByAllUser=" & IIf(Txt(AcPostingByAllUser) = "Yes", 1, 0) & _
                ", VehGodown = '" & Txt(VehGodown).Tag & "', RSO_Code = '" & Txt(RSO_Code) & _
                "',VisitObjCode='" & Txt(VisitObjCode).Tag & "' , Valid_Day= " & Val(Txt(Valid_Day)) & _
                ", OwnFinCode = '" & Txt(OwnFinCode).Tag & "', DelayInttRate = " & Val(Txt(DelayInttRate)) & _
                ", RebDays = " & Val(Txt(RebDays)) & ", DebtorInSupplierHelp=" & IIf(Txt(DebtorInSupplierHelp) = "Yes", 1, 0) & ",SupInvOnVehSaleInv = " & IIf(Txt(SupInvOnVehSaleInv) = "Yes", 1, 0) & _
                ", PostRegnFeeYn=" & IIf(Txt(PostRegnFeeYn) = "Yes", 1, 0) & ", PostInsuranceFeeYn=" & IIf(Txt(PostInsuranceFeeYn) = "Yes", 1, 0) & ", PostOctraiSaperatelyYn=" & IIf(Txt(PostOctraiSaperatelyYn) = "Yes", 1, 0) & ", TaxDetOnVehInv = " & IIf(Txt(TaxDetOnVehInv) = "Yes", 1, 0) & _
                ", VehQuotFooter = '" & Txt(VehQuotFooter) & "',VehBookFooter = '" & Txt(VehBookFooter) & _
                "',VehSaleInvFooter = '" & Txt(VehSaleInvFooter) & "', VehTaxInvPrefix = '" & Txt(VehTaxInvPrefix) & "',VehRateIncTax=" & IIf(Txt(VehRateInclTax) = "Yes", 1, 0)
                
            GCn.Execute GSQL

            GCnFaV.Execute "Update AcControls Set " _
                & "VehCre_Grp='" & IIf(IsNull(Txt(VehCreGrp).Tag), "", Txt(VehCreGrp).Tag) & _
                "',VehDeb_Grp='" & IIf(IsNull(Txt(VehDebGrp).Tag), "", Txt(VehDebGrp).Tag) & _
                "',VehCash_Ac='" & IIf(IsNull(Txt(VehCashAc).Tag), "", Txt(VehCashAc).Tag) & _
                "',VehBank_Ac='" & IIf(IsNull(Txt(VehBankAc).Tag), "", Txt(VehBankAc).Tag) & _
                "',Fitment_Ac='" & IIf(IsNull(Txt(FitmentAc).Tag), "", Txt(FitmentAc).Tag) & _
                "',Fuel_Ac='" & IIf(IsNull(Txt(FuelAc).Tag), "", Txt(FuelAc).Tag) & _
                "',StampDuty_Ac='" & IIf(IsNull(Txt(StampDutyAc).Tag), "", Txt(StampDutyAc).Tag) & _
                "',ServiceChrg_Ac='" & IIf(IsNull(Txt(ServiceChrgAc).Tag), "", Txt(ServiceChrgAc).Tag) & _
                "',Interest_Ac='" & IIf(IsNull(Txt(InterestAc).Tag), "", Txt(InterestAc).Tag) & _
                "',TDS_Ac='" & IIf(IsNull(Txt(TDSAc).Tag), "", Txt(TDSAc).Tag) & _
                "',VehROff_Ac='" & IIf(IsNull(Txt(VehRoffAc).Tag), "", Txt(VehRoffAc).Tag) & _
                "',IndirectExpAc='" & IIf(IsNull(Txt(IndirectExpAc).Tag), "", Txt(IndirectExpAc).Tag) & "',SubventionAc='" & IIf(IsNull(Txt(SubventionAc).Tag), "", Txt(SubventionAc).Tag) & "',SubventionClaimAc='" & IIf(IsNull(Txt(SubventionClaimAc).Tag), "", Txt(SubventionClaimAc).Tag) & "',OctraiAc='" & IIf(IsNull(Txt(OctraiAc).Tag), "", Txt(OctraiAc).Tag) & "',RegnFeeAc='" & IIf(IsNull(Txt(RegnFeeAc).Tag), "", Txt(RegnFeeAc).Tag) & "',SpecialDiscountAc ='" & IIf(IsNull(Txt(SpecialDiscountAc).Tag), "", Txt(SpecialDiscountAc).Tag) & "', InsuranceFeeAc='" & IIf(IsNull(Txt(InsuranceFeeAc).Tag), "", Txt(InsuranceFeeAc).Tag) & "', TOTax_Ac='" & IIf(IsNull(Txt(TOTAc).Tag), "", Txt(TOTAc).Tag) & _
                "' where Div_Code = '" & PubDivCode & "'"
        End If
    GCn.CommitTrans
    GCnFaV.CommitTrans
    mTrans = False
    Master.Requery
    Syctrl.Requery
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    If mTrans = True Then GCn.RollbackTrans: GCnFaV.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To Txt.Count - 1
            Txt(I).BackColor = CtrlBColOrg
            Txt(I).ForeColor = CtrlFColOrg
        Next
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
On Error GoTo ELoop
Ctrl_GetFocus Txt(Index)
Grid_Hide
MyIndex = Index
Select Case Index
    Case VehGodown
        If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsGodown!Name Then
            RsGodown.MoveFirst
            RsGodown.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case VisitObjCode
        If rsObj.RecordCount = 0 Or (rsObj.EOF = True Or rsObj.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsObj!Name Then
            rsObj.MoveFirst
            rsObj.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case OwnFinCode
        If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsFin!Name Then
            rsFin.MoveFirst
            rsFin.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case VehCreGrp, VehDebGrp
        If rsGrp.RecordCount = 0 Or (rsGrp.EOF = True Or rsGrp.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsGrp!Name Then
            rsGrp.MoveFirst
            rsGrp.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case VehCashAc, VehBankAc, FitmentAc, FuelAc, StampDutyAc, ServiceChrgAc, InterestAc, TDSAc, VehRoffAc
        If rsAc.RecordCount = 0 Or (rsAc.EOF = True Or rsAc.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsAc!Name Then
            rsAc.MoveFirst
            rsAc.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
        Case VehGodown
            DGridTxtKeyDown DGGodown, Txt, Index, RsGodown, KeyCode, False, 1, frmGodown, "frmGodown"
        Case VisitObjCode
            DGridTxtKeyDown DGObj, Txt, Index, rsObj, KeyCode, False, 1
        Case OwnFinCode
            DGridTxtKeyDown DGFin, Txt, Index, rsFin, KeyCode, False, 1, frmFinMast, "frmFinMast"
        Case VehCreGrp, VehDebGrp
            DGridTxtKeyDown DGGrp, Txt, Index, rsGrp, KeyCode, False, 1, frmGrEnt, "frmGrEnt"
       Case VehCashAc, VehBankAc, FitmentAc, FuelAc, StampDutyAc, ServiceChrgAc, InterestAc, TDSAc, VehRoffAc, TOTAc, RegnFeeAc, SpecialDiscountAc, InsuranceFeeAc, OctraiAc, IndirectExpAc, SubventionAc, SubventionClaimAc
            DGridTxtKeyDown DGAc, Txt, Index, rsAc, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
    End Select
    If Txt(Index).MultiLine = False Then
        If DGObj.Visible = False And DGFin.Visible = False And DGGrp.Visible = False And DGAc.Visible = False Then   'Arrow Key
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VehRoffAc Then
               If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
            End If
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> VehRoffAc Then Ctrl_DownKeyDown KeyCode, Shift
            If TopCtrl1.TopText2.CAPTION = "Edit" Then
                If Index <> VehCreGrp And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
Select Case Index
    Case VehGodown
        If DGGodown.Visible = True Then DGridTxtKeyPress Txt, VehGodown, RsGodown, KeyAscii, "Name"
    Case Valid_Day, RebDays
        NumPress Txt(Index), KeyAscii, 3, 0
    Case DelayInttRate
        NumPress Txt(Index), KeyAscii, 2, 2
    Case AcPostingByAllUser, SupInvOnVehSaleInv, TaxDetOnVehInv, VehRateInclTax, DebtorInSupplierHelp, PostRegnFeeYn, PostInsuranceFeeYn, PostOctraiSaperatelyYn
        If UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index).TEXT = "Yes"
            KeyAscii = 0
        Else    'If KeyAscii = 87 Or KeyAscii = 119 Then   ' W/w
            If KeyAscii <> vbKeyReturn Then
                Txt(Index).TEXT = "No"
                KeyAscii = 0
            End If
        End If
    Case VehCreGrp, VehDebGrp
        If DGGrp.Visible = True Then DGridTxtKeyPress Txt, Index, rsGrp, KeyAscii, "Name"
    Case VehCashAc, VehBankAc, FitmentAc, FuelAc, StampDutyAc, ServiceChrgAc, InterestAc, TDSAc, VehRoffAc, TOTAc, RegnFeeAc, SpecialDiscountAc, InsuranceFeeAc, OctraiAc, IndirectExpAc, SubventionAc, SubventionClaimAc
        If DGAc.Visible = True Then DGridTxtKeyPress Txt, Index, rsAc, KeyAscii, "Name"
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
    Select Case Index
        Case VehGodown
            If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or Txt(Index).TEXT = "" Then
                Txt(Index) = ""
                Txt(Index).Tag = ""
            Else
                Txt(Index).Tag = RsGodown!Code
                Txt(Index) = RsGodown!Name
            End If
        Case VisitObjCode
            If rsObj.RecordCount = 0 Or (rsObj.EOF = True Or rsObj.BOF = True) Or Txt(Index).TEXT = "" Then
                Txt(Index) = ""
                Txt(Index).Tag = ""
            Else
                Txt(Index).Tag = rsObj!Code
                Txt(Index) = rsObj!Name
            End If
        Case OwnFinCode
            If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or Txt(Index).TEXT = "" Then
                Txt(Index) = ""
                Txt(Index).Tag = ""
            Else
                Txt(Index).Tag = rsFin!Code
                Txt(Index) = rsFin!Name
            End If
        Case VehCreGrp, VehDebGrp
            If rsGrp.RecordCount > 0 Or (rsGrp.EOF = False Or rsGrp.BOF = False) Then
                If Txt(Index).TEXT <> "" Then
                    Txt(Index).TEXT = rsGrp!Name
                    Txt(Index).Tag = rsGrp!Code
                End If
            Else
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            End If
        Case VehCashAc, VehBankAc, FitmentAc, FuelAc, StampDutyAc, ServiceChrgAc, InterestAc, TDSAc, VehRoffAc
            If rsAc.RecordCount = 0 Or (rsAc.EOF = True Or rsAc.BOF = True) Or Txt(Index).TEXT = "" Then
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            Else
                Txt(Index).TEXT = rsAc!Name
                Txt(Index).Tag = rsAc!Code
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

