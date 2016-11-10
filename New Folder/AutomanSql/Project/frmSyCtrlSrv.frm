VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmSyCtrlSrv 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Workshop A/c Control Declaration"
   ClientHeight    =   8760
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
   ScaleHeight     =   8760
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DGAc 
      Height          =   3330
      Left            =   90
      Negotiate       =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6600
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
            DividerStyle    =   3
            ColumnWidth     =   5220.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGGrp 
      Height          =   3330
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   6060
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
            DividerStyle    =   3
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5640
      Left            =   30
      TabIndex        =   1
      Top             =   255
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   9948
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   12640511
      TabCaption(0)   =   "1. General Settings"
      TabPicture(0)   =   "frmSyCtrlSrv.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameService"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "2. A/c Settings"
      TabPicture(1)   =   "frmSyCtrlSrv.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "LblColon(23)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Lbl(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Lbl(9)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Lbl(10)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Lbl(23)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Lbl(17)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Lbl(14)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Lbl(12)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Lbl(0)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Lbl(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Lbl(2)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Lbl(4)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Lbl(5)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Lbl(6)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Lbl(7)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Lbl(8)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Txt(3)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Txt(2)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Txt(0)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Txt(4)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Txt(5)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Txt(6)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Txt(1)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Txt(18)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Txt(19)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Txt(20)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Txt(27)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Txt(29)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Txt(31)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Txt(32)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Txt(33)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).ControlCount=   31
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
         TabIndex        =   97
         ToolTipText     =   "Press L-> Local or C-> Central"
         Top             =   4380
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
         TabIndex        =   95
         ToolTipText     =   "Press L-> Local or C-> Central"
         Top             =   4140
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
         Index           =   31
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   30
         Top             =   3420
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
         Index           =   29
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   29
         Top             =   3180
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
         TabIndex        =   28
         Top             =   2940
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
         Index           =   20
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   31
         Top             =   3660
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
         Index           =   19
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   32
         Top             =   3900
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
         Index           =   18
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   24
         Top             =   1980
         Width           =   4785
      End
      Begin VB.Frame FrameService 
         Appearance      =   0  'Flat
         BackColor       =   &H00BAD3C9&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   4785
         Left            =   -74970
         TabIndex        =   43
         Top             =   330
         Width           =   11595
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   36
            Left            =   10890
            MaxLength       =   40
            TabIndex        =   103
            Text            =   "99.99"
            Top             =   1335
            Width           =   540
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   35
            Left            =   10890
            MaxLength       =   40
            TabIndex        =   101
            Text            =   "99.99"
            Top             =   855
            Width           =   540
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   34
            Left            =   10890
            MaxLength       =   40
            TabIndex        =   100
            Text            =   "99.99"
            Top             =   1095
            Width           =   540
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   30
            Left            =   10995
            MaxLength       =   25
            TabIndex        =   4
            Text            =   "Yes/No"
            Top             =   300
            Width           =   420
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   28
            Left            =   9225
            MaxLength       =   30
            TabIndex        =   90
            Top             =   2295
            Width           =   2190
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   26
            Left            =   9225
            MaxLength       =   21
            TabIndex        =   16
            Top             =   2055
            Width           =   2190
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            Height          =   225
            Index           =   25
            Left            =   3300
            MaxLength       =   40
            TabIndex        =   12
            Text            =   "Spr/Lab inv"
            Top             =   1305
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   240
            Index           =   24
            Left            =   6855
            MaxLength       =   40
            TabIndex        =   15
            Text            =   "Default Godown"
            Top             =   2310
            Width           =   420
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   23
            Left            =   6840
            MaxLength       =   40
            TabIndex        =   14
            Text            =   "Spr/Lab inv"
            Top             =   2055
            Width           =   420
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   240
            Index           =   22
            Left            =   3300
            MaxLength       =   40
            TabIndex        =   13
            Text            =   "Default Godown"
            Top             =   1560
            Width           =   420
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   21
            Left            =   3300
            MaxLength       =   40
            TabIndex        =   11
            Text            =   "Spr/Lab inv"
            Top             =   1050
            Width           =   420
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   9
            Left            =   10980
            MaxLength       =   40
            TabIndex        =   5
            Top             =   -45
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   11
            Left            =   6285
            MaxLength       =   40
            TabIndex        =   7
            Text            =   "Yes/No"
            Top             =   540
            Width           =   420
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   10
            Left            =   2985
            MaxLength       =   40
            TabIndex        =   6
            Text            =   "9999.99"
            Top             =   540
            Width           =   765
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   7
            Left            =   2595
            MaxLength       =   5
            TabIndex        =   2
            Text            =   "X(5)"
            Top             =   300
            Width           =   660
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   13
            Left            =   2160
            MaxLength       =   40
            TabIndex        =   9
            Text            =   "Spr/Lab inv"
            Top             =   795
            Width           =   1590
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   12
            Left            =   8235
            MaxLength       =   25
            TabIndex        =   8
            Text            =   "Yes/No"
            Top             =   540
            Width           =   420
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   225
            Index           =   8
            Left            =   10875
            MaxLength       =   40
            TabIndex        =   3
            Text            =   "99.99"
            Top             =   540
            Width           =   540
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
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
            Height          =   240
            Index           =   14
            Left            =   6285
            MaxLength       =   40
            TabIndex        =   10
            Text            =   "Default Godown"
            Top             =   780
            Width           =   2370
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   1860
            Left            =   0
            TabIndex        =   53
            Top             =   2850
            Width           =   11625
            _ExtentX        =   20505
            _ExtentY        =   3281
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
            TabCaption(0)   =   "1. Job Card"
            TabPicture(0)   =   "frmSyCtrlSrv.frx":0038
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Txt(15)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "2. Works Spare Invoice "
            TabPicture(1)   =   "frmSyCtrlSrv.frx":0054
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Txt(16)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "3. Labour Invoice"
            TabPicture(2)   =   "frmSyCtrlSrv.frx":0070
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Txt(17)"
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
               Height          =   1215
               Index           =   17
               Left            =   -74865
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   19
               Top             =   465
               Width           =   11385
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
               Height          =   1275
               Index           =   16
               Left            =   -74850
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               Top             =   450
               Width           =   11355
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
               Height          =   1320
               Index           =   15
               Left            =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   420
               Width           =   11205
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
               TabIndex        =   77
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
               TabIndex        =   76
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
               TabIndex        =   75
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
               TabIndex        =   74
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
               TabIndex        =   73
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
               TabIndex        =   72
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
               TabIndex        =   71
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
               TabIndex        =   70
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
               TabIndex        =   69
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
               TabIndex        =   64
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
               TabIndex        =   63
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
               TabIndex        =   62
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
               TabIndex        =   61
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
               TabIndex        =   60
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
               TabIndex        =   59
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
               TabIndex        =   58
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
               TabIndex        =   57
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
               TabIndex        =   56
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
               TabIndex        =   55
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
               TabIndex        =   54
               Top             =   1425
               Width           =   90
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "H. E. Cess% "
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
            Index           =   11
            Left            =   9555
            TabIndex        =   104
            Top             =   1335
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Tax %"
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
            Index           =   9
            Left            =   9555
            TabIndex        =   102
            Top             =   855
            Width           =   1140
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cess% :"
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
            Index           =   10
            Left            =   10095
            TabIndex        =   99
            Top             =   1095
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FSB Online Posting :"
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
            Index           =   8
            Left            =   9060
            TabIndex        =   93
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Tax No."
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
            Index           =   7
            Left            =   7410
            TabIndex        =   91
            Top             =   2295
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "24-Hrs Help Line No."
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
            Index           =   6
            Left            =   7410
            TabIndex        =   88
            Top             =   2055
            Width           =   1740
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Tax on OutSide Jobs   :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   5
            Left            =   600
            TabIndex        =   87
            Top             =   1305
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complementary Issue :"
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
            Index           =   4
            Left            =   4800
            TabIndex        =   86
            Top             =   2325
            Width           =   1920
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Issue             :"
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
            Index           =   3
            Left            =   4800
            TabIndex        =   85
            Top             =   2055
            Width           =   1935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Print on Spare Bill :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   2
            Left            =   4800
            TabIndex        =   84
            Top             =   1800
            Width           =   1560
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Separate Bill for Labour            :"
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
            Index           =   1
            Left            =   600
            TabIndex        =   83
            Top             =   1575
            Width           =   2520
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Labour Disc. on OutSide Jobs :"
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
            Index           =   0
            Left            =   600
            TabIndex        =   82
            Top             =   1050
            Width           =   2550
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
            Left            =   15
            TabIndex        =   52
            Top             =   1815
            Width           =   1530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Separate Series for Spare Inv. :"
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
            Index           =   39
            Left            =   8355
            TabIndex        =   51
            Top             =   -45
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hour Details in Labour Inv. :"
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
            Index           =   37
            Left            =   3930
            TabIndex        =   50
            Top             =   540
            Width           =   2265
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Major Labour Amt. Limit :"
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
            Index           =   36
            Left            =   885
            TabIndex        =   49
            Top             =   540
            Width           =   2025
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Zone Code :"
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
            Index           =   35
            Left            =   885
            TabIndex        =   48
            Top             =   300
            Width           =   1650
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gate Pass :"
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
            Index           =   34
            Left            =   7125
            TabIndex        =   47
            Top             =   540
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gate Pass On :"
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
            Index           =   33
            Left            =   870
            TabIndex        =   46
            Top             =   795
            Width           =   1245
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Tax with Cess% :"
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
            Index           =   32
            Left            =   8730
            TabIndex        =   45
            Top             =   555
            Width           =   2055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Godown :"
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
            Index           =   112
            Left            =   4800
            TabIndex        =   44
            Top             =   780
            Width           =   1410
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
         TabIndex        =   21
         ToolTipText     =   "Press L-> Local or C-> Central"
         Top             =   1260
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
         Index           =   6
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   27
         Top             =   2700
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
         TabIndex        =   26
         ToolTipText     =   "Press L-> Local or C-> Central"
         Top             =   2460
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
         TabIndex        =   25
         Top             =   2220
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
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   20
         ToolTipText     =   "Press L-> Local or C-> Central"
         Top             =   1020
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
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1500
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
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   23
         Top             =   1740
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque Clearing A/c"
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
         Left            =   90
         TabIndex        =   98
         Top             =   4365
         Width           =   1785
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card A/c"
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
         Left            =   90
         TabIndex        =   96
         Top             =   4125
         Width           =   1350
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FSB Credit A/C"
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
         Left            =   90
         TabIndex        =   94
         Top             =   3435
         Width           =   1305
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Dealer Group"
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
         Left            =   90
         TabIndex        =   92
         Top             =   3195
         Width           =   1695
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Conractor A/c"
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
         Left            =   90
         TabIndex        =   89
         Top             =   2925
         Width           =   1545
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telco AMC A/c"
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
         Index           =   2
         Left            =   90
         TabIndex        =   81
         Top             =   3675
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer AMC A/c"
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
         Index           =   1
         Left            =   90
         TabIndex        =   80
         Top             =   3915
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Labour Charge A/c  Taxable"
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
         Left            =   90
         TabIndex        =   79
         Top             =   1965
         Width           =   2415
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "                            Taxpaid"
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
         Left            =   90
         TabIndex        =   42
         Top             =   2205
         Width           =   2355
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Tax A/c Name"
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
         Left            =   90
         TabIndex        =   41
         Top             =   2445
         Width           =   1920
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank A/c Name"
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
         Index           =   17
         Left            =   90
         TabIndex        =   40
         Top             =   1740
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Labour Round Off A/c Name"
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
         Left            =   90
         TabIndex        =   39
         Top             =   2685
         Width           =   2400
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash A/c Name"
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
         Index           =   10
         Left            =   90
         TabIndex        =   38
         Top             =   1500
         Width           =   1335
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sundry Creditors A/c Group"
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
         Index           =   9
         Left            =   90
         TabIndex        =   37
         Top             =   1020
         Visible         =   0   'False
         Width           =   2400
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
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   36
         ToolTipText     =   "Press L-> Local or C-> Central"
         Top             =   1260
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Label LblColon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Height          =   195
         Index           =   23
         Left            =   90
         TabIndex        =   35
         Top             =   1020
         Visible         =   0   'False
         Width           =   75
      End
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
   Begin MSDataGridLib.DataGrid DGGodown 
      Height          =   3330
      Left            =   15
      Negotiate       =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   16
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
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
            DividerStyle    =   3
            ColumnWidth     =   2789.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSyCtrlSrv"
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
Private Const SrvCreGrp As Byte = 0
Private Const SrvDebGrp As Byte = 1
Private Const SrvCashAc As Byte = 2
Private Const SrvBankAc As Byte = 3
Private Const SrvLabourAc As Byte = 4
Private Const SrvTaxAc As Byte = 5
Private Const SrvRoffAc As Byte = 6
Private Const SrvLabourAcTB As Byte = 18

'Private Const JobCard_Type As Byte = 3 replaced by Division
Private Const ServiceZone As Byte = 7
Private Const Service_Tax As Byte = 8
'Private Const Srv_SeparateSprInvSrlNo As Byte = 9
Private Const MajorLabLimit As Byte = 10
Private Const HrDetOnLabInv As Byte = 11
Private Const SrvGatePass As Byte = 12
Private Const SrvGatePass_On As Byte = 13
Private Const SprWorksGodown As Byte = 14
Private Const JobCardFooter As Byte = 15
Private Const WorkShopInvFooter As Byte = 16
Private Const LabInvFooter As Byte = 17
Private Const DlrAMC As Byte = 19
Private Const TelcoAMC As Byte = 20
Private Const OutSideLabDisc As Byte = 21
Private Const SepLabourInv As Byte = 22   'SepLabourInv
Private Const PrintCompanyIssue As Byte = 23
Private Const PrintComplIssue As Byte = 24
Private Const TaxOnOutSideLab As Byte = 25
Private Const HelpLineNo As Byte = 26
Private Const JobContractor As Byte = 27
Private Const SrvTaxNo As Byte = 28
Private Const OtherDealerGroup As Byte = 29
Private Const FSBOnlinePost As Byte = 30
Private Const FSBCrAcCode As Byte = 31
Private Const CreditCardAc As Byte = 32
Private Const ChqClrAc As Byte = 33
Private Const eCessPer As Byte = 34
Private Const ServiceTaxPer_Saperate As Byte = 35
Private Const HECessPer As Byte = 36

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To txt.Count - 1
        txt(I).Enabled = Enb
    Next
txt(JobCardFooter).Enabled = True
txt(WorkShopInvFooter).Enabled = True
txt(LabInvFooter).Enabled = True
txt(JobCardFooter).Locked = True
txt(WorkShopInvFooter).Locked = True
txt(LabInvFooter).Locked = True
End Sub
'* Used for intialize grid columns
Private Sub Grid_Ini()
    DGGodown.left = Me.width - (DGGodown.width + mRtScale): DGGodown.top = mTopScale
    DGGrp.left = Me.width - (DGGrp.width + mRtScale): DGGrp.top = mTopScale
    DGAc.left = Me.width - (DGAc.width + mRtScale): DGAc.top = mTopScale
End Sub
Private Sub Grid_Hide()
    If DGGodown.Visible Then DGGodown.Visible = False
    If DGAc.Visible = True Then DGAc.Visible = False
    If DGGrp.Visible = True Then DGGrp.Visible = False
End Sub
Private Sub MoveRec()
Dim Master1 As ADODB.Recordset
On Error Resume Next 'by lps at Cuttack GoTo ELoop
SSTab1.Tab = 0
SSTab2.Tab = 0
'General Settings
txt(ServiceZone) = IIf(IsNull(Syctrl!ServiceZone), "", Syctrl!ServiceZone)
txt(Service_Tax) = IIf(IsNull(Syctrl!Service_Tax), "", Syctrl!Service_Tax)
txt(eCessPer) = IIf(IsNull(Syctrl!eCessPer), "", Syctrl!eCessPer)
txt(MajorLabLimit) = IIf(IsNull(Syctrl!MajorLabLimit), "", Syctrl!MajorLabLimit)
txt(HrDetOnLabInv) = IIf(Syctrl!HrDetOnLabInv = 1, "Yes", "No")
txt(SrvGatePass) = IIf(Syctrl!SrvGatePass = 1, "Yes", "No")
txt(OutSideLabDisc) = IIf(Syctrl!OutSideLabDisc = 1, "Yes", "No")
txt(SepLabourInv) = IIf(Syctrl!SepLabourInv = 1, "Yes", "No")
txt(PrintCompanyIssue) = IIf(Syctrl!PrintCompanyIssue = 1, "Yes", "No")
txt(PrintComplIssue) = IIf(Syctrl!PrintComplIssue = 1, "Yes", "No")
txt(TaxOnOutSideLab) = IIf(Syctrl!SrvTaxOnOutSideLab = 1, "Yes", "No")
txt(FSBOnlinePost) = IIf(Master!FSBOnlinePost = 1, "Yes", "No")
txt(HelpLineNo) = Syctrl!HelpLineNo
txt(SrvTaxNo) = XNull(Syctrl!SrvTaxNo)
txt(ServiceTaxPer_Saperate) = VNull(Syctrl!ServiceTaxPer_Saperate)
txt(HECessPer) = VNull(Syctrl!HECessPer)
If Syctrl!SrvGatePass = 1 Then
    If Syctrl!SrvGatePass_On = "S" Then
        txt(SrvGatePass_On) = "Spare Invoice"
        txt(SrvGatePass_On).Tag = "S"
    Else
        txt(SrvGatePass_On) = "Labour Invoice"
        txt(SrvGatePass_On).Tag = "L"
    End If
Else
    txt(SrvGatePass_On).Tag = ""
    txt(SrvGatePass_On) = ""
End If
If Syctrl!SprWorksGodown <> "" Then
    RsGodown.MoveFirst
    RsGodown.FIND ("Code ='" & Syctrl!SprWorksGodown & "'")
    txt(SprWorksGodown).Tag = Syctrl!SprWorksGodown
    txt(SprWorksGodown) = IIf(RsGodown.EOF, "", RsGodown!Name)
Else
    txt(SprWorksGodown) = ""
    txt(SprWorksGodown).Tag = ""
End If
If RsGodown.RecordCount > 0 And RsGodown.EOF Then RsGodown.MoveFirst

txt(JobCardFooter) = IIf(IsNull(Syctrl!JobCardFooter), "", Syctrl!JobCardFooter)
txt(WorkShopInvFooter) = IIf(IsNull(Syctrl!WorkShopInvFooter), "", Syctrl!WorkShopInvFooter)
txt(LabInvFooter) = IIf(IsNull(Syctrl!LabInvFooter), "", Syctrl!LabInvFooter)
'**
    txt(SrvCreGrp) = ""
    txt(SrvCreGrp).Tag = ""
    txt(SrvDebGrp) = ""
    txt(SrvDebGrp).Tag = ""
    txt(SrvCashAc) = ""
    txt(SrvCashAc).Tag = ""
    txt(SrvBankAc) = ""
    txt(SrvBankAc).Tag = ""
    txt(SrvLabourAcTB) = ""
    txt(SrvLabourAcTB).Tag = ""
    txt(SrvLabourAc) = ""
    txt(SrvLabourAc).Tag = ""
    txt(SrvTaxAc) = ""
    txt(SrvTaxAc).Tag = ""
    txt(SrvRoffAc) = ""
    txt(SrvRoffAc).Tag = ""
    txt(JobContractor) = ""
    txt(JobContractor).Tag = ""
    txt(DlrAMC) = ""
    txt(DlrAMC).Tag = ""
    txt(TelcoAMC) = ""
    txt(TelcoAMC).Tag = ""
'** A/c Section
If Master!SrvCre_Grp <> Null Or Master!SrvCre_Grp <> "" Then
    txt(SrvCreGrp) = GCnFaW.Execute("Select GroupName from AcGroup where GroupCode='" & Master!SrvCre_Grp & "'").Fields(0).Value
    txt(SrvCreGrp).Tag = Master!SrvCre_Grp
End If
If Master!SrvDeb_Grp <> Null Or Master!SrvDeb_Grp <> "" Then
    txt(SrvDebGrp) = GCnFaW.Execute("Select GroupName from AcGroup where GroupCode='" & Master!SrvDeb_Grp & "'").Fields(0).Value
    txt(SrvDebGrp).Tag = Master!SrvDeb_Grp
End If
Set Master1 = New Recordset
Master1.CursorLocation = adUseClient
Master1.Open "Select SubCode,Name from SubGroup Order by SubCode", GCnFaW, adOpenStatic, adLockReadOnly
If Master!SrvCash_Ac <> Null Or Master!SrvCash_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SrvCash_Ac & "'")
    txt(SrvCashAc) = Master1!Name
    txt(SrvCashAc).Tag = Master1!SubCode
End If
If Master!SrvBank_Ac <> Null Or Master!SrvBank_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SrvBank_Ac & "'")
    txt(SrvBankAc) = Master1!Name
    txt(SrvBankAc).Tag = Master1!SubCode
End If
If Master!SrvLabourTB_Ac <> Null Or Master!SrvLabourTB_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SrvLabourTB_Ac & "'")
    txt(SrvLabourAcTB) = Master1!Name
    txt(SrvLabourAcTB).Tag = Master1!SubCode
End If
If Master!SrvLabour_Ac <> Null Or Master!SrvLabour_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SrvLabour_Ac & "'")
    txt(SrvLabourAc) = Master1!Name
    txt(SrvLabourAc).Tag = Master1!SubCode
End If
If Master!SrvTax_Ac <> Null Or Master!SrvTax_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SrvTax_Ac & "'")
    txt(SrvTaxAc) = Master1!Name
    txt(SrvTaxAc).Tag = Master1!SubCode
End If
If Master!SrvROff_Ac <> Null Or Master!SrvROff_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SrvROff_Ac & "'")
    txt(SrvRoffAc) = Master1!Name
    txt(SrvRoffAc).Tag = Master1!SubCode
End If
If Master!DlrAMC_Ac <> Null Or Master!DlrAMC_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!DlrAMC_Ac & "'")
    txt(DlrAMC) = Master1!Name
    txt(DlrAMC).Tag = Master1!SubCode
End If
If Master!TelcoAMC_Ac <> Null Or Master!TelcoAMC_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!TelcoAMC_Ac & "'")
    txt(TelcoAMC) = Master1!Name
    txt(TelcoAMC).Tag = Master1!SubCode
End If
If Master!JobContractor <> Null Or Master!JobContractor <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!JobContractor & "'")
    txt(JobContractor) = Master1!Name
    txt(JobContractor).Tag = Master1!SubCode
End If
If Master!FSBCrAc <> Null Or Master!FSBCrAc <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!FSBCrAc & "'")
    txt(FSBCrAcCode) = Master1!Name
    txt(FSBCrAcCode).Tag = Master1!SubCode
End If
If Master!CreditCardAc <> Null Or Master!CreditCardAc <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!CreditCardAc & "'")
    txt(CreditCardAc) = Master1!Name
    txt(CreditCardAc).Tag = Master1!SubCode
End If
If Master!ChqClrAc <> Null Or Master!ChqClrAc <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!ChqClrAc & "'")
    txt(ChqClrAc) = Master1!Name
    txt(ChqClrAc).Tag = Master1!SubCode
End If

If Master!OthDealerGrp <> Null Or Master!OthDealerGrp <> "" Then
    txt(OtherDealerGroup) = GCnFaW.Execute("Select GroupName from AcGroup where GroupCode='" & Master!OthDealerGrp & "'").Fields(0).Value
    txt(OtherDealerGroup).Tag = Master!OthDealerGrp
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
Private Sub DGGodown_Click()
On Error GoTo ELoop
    If RsGodown.RecordCount > 0 Then
        txt(MyIndex).TEXT = RsGodown!Name
        txt(MyIndex).Tag = RsGodown!Code
    End If
    txt(MyIndex).SetFocus
    DGGodown.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGAc_Click()
On Error GoTo ELoop
    If rsAc.RecordCount > 0 Then
        txt(MyIndex).TEXT = rsAc!Name
        txt(MyIndex).Tag = rsAc!Code
    End If
    txt(MyIndex).SetFocus
    DGAc.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGGrp_Click()
On Error GoTo ELoop
    If rsGrp.RecordCount > 0 Then
        txt(MyIndex).TEXT = rsGrp!Name
        txt(MyIndex).Tag = rsGrp!Code
    End If
    txt(MyIndex).SetFocus
    DGGrp.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Activate()
Dim UnLoadFrm As Boolean, MsgStr$
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
    Call TopCtrl1_eRef
End If
If RsGodown.RecordCount <= 0 Then
    MsgStr = "No Records in Godown Master"
    UnLoadFrm = True
End If
If UnLoadFrm Then
    MsgBox MsgStr, vbInformation, "Validation"
    Unload Me
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
'FA Connection for Works FAData
    TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini

    For I = 0 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg '&HDFF4F2
        txt(I).ForeColor = CtrlFColOrg
    Next
    Set RsGodown = New ADODB.Recordset
    RsGodown.CursorLocation = adUseClient
    RsGodown.Open "Select God_Code as Code, God_Name As Name From Godown  where Appli_For=0  Order by God_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGodown.DataSource = RsGodown
    
    Set rsGrp = New ADODB.Recordset
    rsGrp.CursorLocation = adUseClient
    rsGrp.Open "Select GroupCode As Code,GroupName As Name,GroupNature,MainGrCode From AcGroup Where MainGrCode<>'999' Order by GroupName", GCnFaW, adOpenDynamic, adLockOptimistic
    Set DGGrp.DataSource = rsGrp
    
    Set rsAc = New ADODB.Recordset
    rsAc.CursorLocation = adUseClient
    rsAc.Open "Select SubCode as Code,Name From SubGroup Order by Name", GCnFaW, adOpenDynamic, adLockOptimistic
    Set DGAc.DataSource = rsAc

    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    Master.Open "Select * from AcControls where Div_Code='" & PubDivCode & "'", GCnFaW, adOpenDynamic, adLockOptimistic
    If Master.RecordCount <= 0 Then
        Master.AddNew
        Master!Div_Code = PubDivCode
        Master.Update
    End If
'    Set Master = GCnFaW.Execute("Select * from AcControls")
    
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
    Set rsGrp = Nothing
    Set rsAc = Nothing
    Set Master = Nothing
    Set RsGodown = Nothing
    Set Syctrl = Nothing
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
    txt(JobCardFooter).Locked = False
    txt(WorkShopInvFooter).Locked = False
    txt(LabInvFooter).Locked = False
    txt(ServiceZone).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eRef()
On Error GoTo ELoop
    rsGrp.Requery
    rsAc.Requery
    Master.Requery
    RsGodown.Requery
    Syctrl.Requery
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eSave()
Dim mTrans As Boolean, MasterSql$
On Error GoTo ELoop
'Apply necessary validations
'If IsValid(Txt(Party), "Party Name") = False Then Exit Sub
If txt(SrvGatePass) = "Yes" Then
    If IsValid(txt(SrvGatePass_On), "Gate Pass On") = False Then Exit Sub
End If

    GCnFaW.BeginTrans
    GCn.BeginTrans
        mTrans = True
        If TopCtrl1.TopText2 = "Edit" Then   'Edit Bill
            GSQL = "update Syctrl set ServiceZone = '" & txt(ServiceZone) & "', eCessPer=" & Val(txt(eCessPer)) & ", Service_Tax = " & Val(txt(Service_Tax)) & _
                " , MajorLabLimit=" & Val(txt(Service_Tax)) & " , HrDetOnLabInv= " & IIf(txt(HrDetOnLabInv) = "Yes", 1, 0) & _
                ", SrvGatePass = " & IIf(txt(SrvGatePass) = "Yes", 1, 0) & ", SrvGatePass_On = '" & left(txt(SrvGatePass_On), 1) & _
                "',SprworksGodown = '" & txt(SprWorksGodown).Tag & "', JobCardFooter = '" & txt(JobCardFooter) & _
                "',WorkShopInvFooter = '" & txt(WorkShopInvFooter) & "',LabInvFooter = '" & txt(LabInvFooter) & _
                "',OutSideLabDisc = " & IIf(txt(OutSideLabDisc) = "Yes", 1, 0) & ",SepLabourInv = " & IIf(txt(SepLabourInv) = "Yes", 1, 0) & _
                ",PrintCompanyIssue = " & IIf(txt(PrintCompanyIssue) = "Yes", 1, 0) & ",PrintComplIssue = " & IIf(txt(PrintComplIssue) = "Yes", 1, 0) & _
                ",SrvTaxOnOutSideLab =" & IIf(txt(TaxOnOutSideLab) = "Yes", 1, 0) & ",HelpLineNo ='" & txt(HelpLineNo) & "',SrvTaxNo ='" & txt(SrvTaxNo) & "', ServiceTaxPer_Saperate=" & Val(txt(ServiceTaxPer_Saperate)) & ", HECessPer = " & Val(txt(HECessPer)) & ""
            GCn.Execute GSQL

            GCnFaW.Execute "Update AcControls Set " _
                & "SrvCre_Grp='" & txt(SrvCreGrp).Tag & "',SrvDeb_Grp='" & txt(SrvDebGrp).Tag & _
                "',SrvCash_Ac='" & txt(SrvCashAc).Tag & "',SrvBank_Ac='" & txt(SrvBankAc).Tag & _
                "',SrvLabour_Ac='" & txt(SrvLabourAc).Tag & "',SrvTax_Ac='" & txt(SrvTaxAc).Tag & _
                "',SrvROff_Ac='" & txt(SrvRoffAc).Tag & "',SrvLabourTB_Ac='" & txt(SrvLabourAcTB).Tag & "',JobContractor='" & txt(JobContractor).Tag & _
                "',DlrAMC_Ac = '" & txt(DlrAMC).Tag & "',TelcoAMC_Ac = '" & txt(TelcoAMC).Tag & _
                "',OthDealerGrp='" & txt(OtherDealerGroup).Tag & "',FSBCrAc='" & txt(FSBCrAcCode).Tag & "', CreditCardAc='" & txt(CreditCardAc).Tag & "', ChqClrAc='" & txt(ChqClrAc).Tag & "',FSBOnlinePost =" & IIf(txt(FSBOnlinePost) = "Yes", 1, 0) & " where Div_Code='" & PubDivCode & "'"
        End If
    GCn.CommitTrans
    GCnFaW.CommitTrans
    mTrans = False
    Master.Requery
    Syctrl.Requery
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    If mTrans = True Then GCn.RollbackTrans: GCnFaW.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To txt.Count - 1
            txt(I).BackColor = CtrlBColOrg
            txt(I).ForeColor = CtrlFColOrg
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
Ctrl_GetFocus txt(Index)
Grid_Hide
MyIndex = Index
Select Case Index
    Case SprWorksGodown
        If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsGodown!Name Then
            RsGodown.MoveFirst
            RsGodown.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case SrvCreGrp, SrvDebGrp
        If rsGrp.RecordCount = 0 Or (rsGrp.EOF = True Or rsGrp.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsGrp!Name Then
            rsGrp.MoveFirst
            rsGrp.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case SrvCashAc, SrvBankAc, SrvLabourAc, SrvTaxAc, SrvRoffAc, SrvLabourAcTB, JobContractor, TelcoAMC, DlrAMC, FSBCrAcCode, CreditCardAc, ChqClrAc
        If rsAc.RecordCount = 0 Or (rsAc.EOF = True Or rsAc.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsAc!Name Then
            rsAc.MoveFirst
            rsAc.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
        Case SprWorksGodown
            DGridTxtKeyDown DGGodown, txt, Index, RsGodown, KeyCode, False, 1, frmGodown, "frmGodown"
        Case SrvCreGrp, SrvDebGrp, OtherDealerGroup
            DGridTxtKeyDown DGGrp, txt, Index, rsGrp, KeyCode, False, 1, frmGrEnt, "frmGrEnt"
        Case SrvCashAc, SrvBankAc, SrvLabourAc, SrvTaxAc, SrvRoffAc, SrvLabourAcTB, JobContractor, TelcoAMC, DlrAMC, FSBCrAcCode, CreditCardAc, ChqClrAc
            DGridTxtKeyDown DGAc, txt, Index, rsAc, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
    End Select
    If txt(Index).MultiLine = False Then
        If DGGodown.Visible = False And DGGrp.Visible = False And DGAc.Visible = False Then  'Arrow Key
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = OtherDealerGroup Then
               If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
            End If
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> SrvRoffAc Then Ctrl_DownKeyDown KeyCode, Shift
            If TopCtrl1.TopText2.CAPTION = "Edit" Then
                If Index <> SrvCreGrp And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
On Error GoTo ELoop
If keyascii = 39 Then keyascii = 0: Exit Sub
Select Case Index
    Case Service_Tax, eCessPer
        NumPress txt(Index), keyascii, 2, 2
    Case MajorLabLimit
        NumPress txt(Index), keyascii, 4, 2
    Case SrvGatePass_On
        If UCase(Chr(keyascii)) = "S" Then
            txt(Index).TEXT = "Spare Invoice"
            keyascii = 0
        Else    'If KeyAscii = 87 Or KeyAscii = 119 Then   ' W/w
            If keyascii <> vbKeyReturn Then
                txt(Index).TEXT = "Labour Invoice"
                keyascii = 0
            End If
        End If
    Case TaxOnOutSideLab, HrDetOnLabInv, SrvGatePass, OutSideLabDisc, SepLabourInv, PrintCompanyIssue, PrintComplIssue, FSBOnlinePost
        If UCase(Chr(keyascii)) = "Y" Then
            txt(Index).TEXT = "Yes"
            keyascii = 0
        Else    'If KeyAscii = 87 Or KeyAscii = 119 Then   ' W/w
            If keyascii <> vbKeyReturn Then
                txt(Index).TEXT = "No"
                keyascii = 0
            End If
        End If
    Case SrvCreGrp, SrvDebGrp, OtherDealerGroup
        If DGGrp.Visible = True Then DGridTxtKeyPress txt, Index, rsGrp, keyascii, "Name"
    Case SrvCashAc, SrvBankAc, SrvLabourAc, SrvTaxAc, SrvRoffAc, SrvLabourAcTB, JobContractor, TelcoAMC, DlrAMC, FSBCrAcCode, CreditCardAc, ChqClrAc
        If DGAc.Visible = True Then DGridTxtKeyPress txt, Index, rsAc, keyascii, "Name"
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
    Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
    Select Case Index
        Case SprWorksGodown
            If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or txt(Index).TEXT = "" Then
                txt(Index) = ""
                txt(Index).Tag = ""
            Else
                txt(Index).Tag = RsGodown!Code
                txt(Index) = RsGodown!Name
            End If
        Case SrvCreGrp, SrvDebGrp
            If rsGrp.RecordCount > 0 Or (rsGrp.EOF = False Or rsGrp.BOF = False) Then
                If txt(Index).TEXT <> "" Then
                    txt(Index).TEXT = rsGrp!Name
                    txt(Index).Tag = rsGrp!Code
                End If
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        Case SrvCashAc, SrvBankAc, SrvLabourAc, SrvTaxAc, SrvRoffAc, SrvLabourAcTB, JobContractor, TelcoAMC, DlrAMC, FSBCrAcCode, CreditCardAc, ChqClrAc
            If rsAc.RecordCount = 0 Or (rsAc.EOF = True Or rsAc.BOF = True) Or txt(Index).TEXT = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            Else
                txt(Index).TEXT = rsAc!Name
                txt(Index).Tag = rsAc!Code
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

