VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmSyCtrlSpr 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Stores Control Declaration"
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
   Begin MSDataGridLib.DataGrid DGPartGrade 
      Height          =   3375
      Left            =   9165
      Negotiate       =   -1  'True
      TabIndex        =   146
      TabStop         =   0   'False
      Top             =   8415
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   5953
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
         Caption         =   "Description"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGAc 
      Height          =   6810
      Left            =   7290
      Negotiate       =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   8490
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   12012
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
   Begin VB.CommandButton CmdAcCurBalUpd 
      Caption         =   "Update Current Balance A/c"
      Enabled         =   0   'False
      Height          =   390
      Left            =   5190
      TabIndex        =   150
      Top             =   465
      Width           =   3000
   End
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   3330
      Left            =   8010
      Negotiate       =   -1  'True
      TabIndex        =   149
      TabStop         =   0   'False
      Top             =   8385
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Form Name"
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
         Caption         =   "Form Code"
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
         DataField       =   "Tax_Per"
         Caption         =   "Tax%"
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
         DataField       =   "Tax_Sur_Per"
         Caption         =   "S.Charge%"
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
            ColumnWidth     =   4004.788
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
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGGrp 
      Height          =   3330
      Left            =   7650
      Negotiate       =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   8430
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
      RowDividerStyle =   1
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
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   2535
      TabIndex        =   142
      Top             =   5685
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   60
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   225
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3228
         View            =   3
         Arrange         =   1
         Sorted          =   -1  'True
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton CmdStkUpd 
      Caption         =   "Update Current Stock(Spare)"
      Enabled         =   0   'False
      Height          =   390
      Left            =   8265
      TabIndex        =   56
      Top             =   465
      Width           =   3000
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   661
   End
   Begin MSDataGridLib.DataGrid DGPartyType 
      Height          =   2295
      Left            =   6630
      Negotiate       =   -1  'True
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   7365
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4048
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
         Size            =   8.25
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
         Caption         =   "Party Type"
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
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2295
      Left            =   4545
      Negotiate       =   -1  'True
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   6255
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4048
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
      RowHeight       =   18
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
         Caption         =   "Description"
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
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGGodown 
      Height          =   3330
      Left            =   1035
      Negotiate       =   -1  'True
      TabIndex        =   147
      TabStop         =   0   'False
      Top             =   6840
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
   Begin MSDataGridLib.DataGrid DGPerson 
      Height          =   3330
      Left            =   1980
      Negotiate       =   -1  'True
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   6435
      _ExtentX        =   11351
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
         Size            =   8.25
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
         Caption         =   "Sales Person"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   5265.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   494.929
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6840
      Left            =   45
      TabIndex        =   46
      Top             =   885
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   12065
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   12243913
      TabCaption(0)   =   "General Settings"
      TabPicture(0)   =   "frmSyCtrlSpr.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameSpare"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "A/c Settings"
      TabPicture(1)   =   "frmSyCtrlSpr.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Txt(58)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Txt(57)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Txt(53)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Txt(56)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Txt(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Txt(55)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Txt(54)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Txt(52)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Txt(51)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Txt(10)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Txt(8)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Txt(3)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Txt(9)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Txt(2)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Txt(4)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Txt(5)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Txt(6)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Txt(7)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Txt(14)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Txt(15)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Txt(1)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Txt(11)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Txt(13)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Txt(12)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Lbl(18)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Lbl(4)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Lbl(20)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Lbl(0)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Lbl(7)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Lbl(24)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Lbl(16)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Lbl(15)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Lbl(14)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Lbl(13)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Lbl(12)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Lbl(11)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Lbl(8)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Lbl(27)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Lbl(5)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Lbl(6)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Lbl(34)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Lbl(40)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Lbl(37)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Lbl(3)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Lbl(9)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Lbl(39)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Lbl(10)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Lbl(17)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "Lbl(21)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "Lbl(22)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "Lbl(1)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Lbl(2)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).ControlCount=   52
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
         Height          =   270
         Index           =   58
         Left            =   -67380
         MaxLength       =   50
         TabIndex        =   61
         Top             =   780
         Width           =   4020
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
         Height          =   270
         Index           =   57
         Left            =   -67590
         MaxLength       =   50
         TabIndex        =   59
         Top             =   495
         Width           =   4230
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
         Height          =   270
         Index           =   53
         Left            =   -67575
         MaxLength       =   50
         TabIndex        =   67
         Top             =   2145
         Width           =   4230
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
         Height          =   270
         Index           =   56
         Left            =   -67560
         MaxLength       =   50
         TabIndex        =   73
         Top             =   3930
         Visible         =   0   'False
         Width           =   4230
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
         Height          =   270
         Index           =   0
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   4230
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
         Height          =   270
         Index           =   55
         Left            =   -67560
         MaxLength       =   50
         TabIndex        =   71
         Top             =   3645
         Visible         =   0   'False
         Width           =   4230
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
         Height          =   270
         Index           =   54
         Left            =   -67575
         MaxLength       =   50
         TabIndex        =   69
         Top             =   3030
         Visible         =   0   'False
         Width           =   4230
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
         Height          =   270
         Index           =   52
         Left            =   -67575
         MaxLength       =   50
         TabIndex        =   65
         Top             =   1860
         Width           =   4230
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
         Height          =   270
         Index           =   51
         Left            =   -67590
         MaxLength       =   50
         TabIndex        =   63
         Top             =   1320
         Width           =   4230
      End
      Begin VB.Frame FrameSpare 
         Appearance      =   0  'Flat
         BackColor       =   &H00BAD3C9&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   6360
         Left            =   45
         TabIndex        =   82
         Top             =   315
         Width           =   11655
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataSource      =   "9"
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
            Height          =   210
            Index           =   62
            Left            =   3540
            MaxLength       =   5
            TabIndex        =   170
            Text            =   "Yes/No"
            Top             =   2910
            Width           =   960
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
            Index           =   61
            Left            =   3855
            MaxLength       =   40
            TabIndex        =   2
            Text            =   "Yes/No"
            Top             =   750
            Width           =   645
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataSource      =   "9"
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
            Height          =   210
            Index           =   17
            Left            =   3855
            TabIndex        =   7
            Text            =   "Yes/No"
            Top             =   1950
            Width           =   645
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   16
            Left            =   3855
            TabIndex        =   0
            Text            =   "Yes/No"
            Top             =   270
            Width           =   645
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
            Left            =   2415
            TabIndex        =   5
            Text            =   "Yes/No"
            Top             =   1470
            Width           =   2085
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   43
            Left            =   3855
            TabIndex        =   3
            Text            =   "Yes/No"
            Top             =   990
            Width           =   645
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   42
            Left            =   4845
            MaxLength       =   40
            TabIndex        =   168
            Text            =   "Yes/No"
            Top             =   4125
            Visible         =   0   'False
            Width           =   645
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   37
            Left            =   3855
            TabIndex        =   167
            Text            =   "Yes/No"
            Top             =   2430
            Width           =   645
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   39
            Left            =   2415
            MaxLength       =   40
            TabIndex        =   6
            Text            =   "Yes/No"
            Top             =   1710
            Width           =   2085
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   38
            Left            =   3855
            MaxLength       =   40
            TabIndex        =   4
            Text            =   "Yes/No"
            Top             =   1230
            Width           =   645
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
            Left            =   3855
            MaxLength       =   40
            TabIndex        =   1
            Text            =   "Yes/No"
            Top             =   510
            Width           =   645
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataSource      =   "9"
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
            Height          =   210
            Index           =   59
            Left            =   3855
            TabIndex        =   9
            Text            =   "Yes/No"
            Top             =   2670
            Width           =   645
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataSource      =   "10"
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
            Height          =   210
            Index           =   18
            Left            =   3855
            TabIndex        =   8
            Text            =   "Y/N"
            Top             =   2190
            Width           =   645
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataSource      =   "10"
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
            Height          =   225
            Index           =   60
            Left            =   4800
            TabIndex        =   28
            Text            =   "99.99"
            Top             =   3585
            Visible         =   0   'False
            Width           =   540
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   35
            Left            =   9930
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "99.99"
            Top             =   2910
            Width           =   540
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   36
            Left            =   10740
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "99.99"
            Top             =   2910
            Width           =   540
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   41
            Left            =   10755
            TabIndex        =   20
            Text            =   "99.99"
            Top             =   2190
            Width           =   540
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
            Index           =   34
            Left            =   7605
            TabIndex        =   24
            Text            =   "Govt Tax Form"
            Top             =   2910
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
            Height          =   210
            Index           =   31
            Left            =   7605
            TabIndex        =   21
            Text            =   "Tax Form"
            Top             =   2670
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
            Height          =   210
            Index           =   25
            Left            =   9045
            MaxLength       =   3
            TabIndex        =   14
            Text            =   "999"
            Top             =   1230
            Width           =   480
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
            Left            =   9870
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "99.99"
            Top             =   1230
            Width           =   510
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
            Left            =   10755
            MaxLength       =   3
            TabIndex        =   16
            Text            =   "99.99"
            Top             =   1230
            Width           =   540
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
            Left            =   10755
            TabIndex        =   18
            Text            =   "99.99"
            Top             =   1710
            Width           =   540
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
            Left            =   9045
            TabIndex        =   11
            Text            =   "Yes/No"
            Top             =   510
            Width           =   2250
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
            Left            =   9045
            TabIndex        =   12
            Text            =   "Yes/No"
            Top             =   750
            Width           =   2250
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
            Left            =   9045
            TabIndex        =   13
            Text            =   "Yes/No"
            Top             =   990
            Width           =   2250
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
            Left            =   10755
            TabIndex        =   17
            Text            =   "99.99"
            Top             =   1470
            Width           =   540
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
            Index           =   21
            Left            =   9045
            TabIndex        =   10
            Text            =   "Default Godown"
            Top             =   270
            Width           =   2250
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
            Height          =   240
            Index           =   20
            Left            =   10935
            TabIndex        =   48
            Text            =   "Yes/No"
            Top             =   4125
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
               Strikethrough   =   -1  'True
            EndProperty
            Height          =   240
            Index           =   44
            Left            =   6810
            MaxLength       =   40
            TabIndex        =   29
            Text            =   "Store Incharge Name"
            Top             =   3855
            Visible         =   0   'False
            Width           =   4545
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
            Height          =   240
            Index           =   45
            Left            =   6810
            MaxLength       =   20
            TabIndex        =   30
            Text            =   "Designation"
            Top             =   4110
            Visible         =   0   'False
            Width           =   2235
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   40
            Left            =   10755
            TabIndex        =   19
            Text            =   "99.99"
            Top             =   1950
            Width           =   540
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   33
            Left            =   10740
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "99.99"
            Top             =   2670
            Width           =   540
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   32
            Left            =   9930
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "99.99"
            Top             =   2670
            Width           =   540
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   1590
            Left            =   15
            TabIndex        =   50
            Top             =   3330
            Width           =   11580
            _ExtentX        =   20426
            _ExtentY        =   2805
            _Version        =   393216
            Tabs            =   5
            Tab             =   2
            TabsPerRow      =   5
            TabHeight       =   520
            ShowFocusRect   =   0   'False
            BackColor       =   12243913
            ForeColor       =   8388736
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "1. Purchase Order"
            TabPicture(0)   =   "frmSyCtrlSpr.frx":0038
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Txt(46)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "2. Purchase Return"
            TabPicture(1)   =   "frmSyCtrlSpr.frx":0054
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Txt(47)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "3. Quot./Estimate"
            TabPicture(2)   =   "frmSyCtrlSpr.frx":0070
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Txt(48)"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "3. Invoice"
            TabPicture(3)   =   "frmSyCtrlSpr.frx":008C
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Txt(49)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "4. Sale Return"
            TabPicture(4)   =   "frmSyCtrlSpr.frx":00A8
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "Txt(50)"
            Tab(4).ControlCount=   1
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
               Height          =   1035
               Index           =   50
               Left            =   -74820
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   35
               Top             =   435
               Width           =   11205
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
               Height          =   1035
               Index           =   49
               Left            =   -74850
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   34
               Top             =   465
               Width           =   11205
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
               Height          =   1035
               Index           =   48
               Left            =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   33
               Top             =   435
               Width           =   11205
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
               Height          =   1035
               Index           =   47
               Left            =   -74880
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   32
               Top             =   420
               Width           =   11205
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
               Height          =   1035
               Index           =   46
               Left            =   -74895
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Top             =   435
               Width           =   11385
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
               Index           =   53
               Left            =   -74115
               TabIndex        =   112
               Top             =   1425
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
               Index           =   82
               Left            =   -74745
               TabIndex        =   111
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
               Index           =   52
               Left            =   -74115
               TabIndex        =   110
               Top             =   960
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
               Index           =   81
               Left            =   -74745
               TabIndex        =   109
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
               Index           =   51
               Left            =   -74115
               TabIndex        =   108
               Top             =   480
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
               Index           =   80
               Left            =   -74775
               TabIndex        =   107
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
               Index           =   50
               Left            =   -73980
               TabIndex        =   106
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
               Index           =   79
               Left            =   -74610
               TabIndex        =   105
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
               Index           =   49
               Left            =   -73980
               TabIndex        =   104
               Top             =   1035
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
               Index           =   78
               Left            =   -74610
               TabIndex        =   103
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
               Index           =   48
               Left            =   -73980
               TabIndex        =   102
               Top             =   555
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
               Index           =   77
               Left            =   -74640
               TabIndex        =   101
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
               Index           =   47
               Left            =   -74130
               TabIndex        =   100
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
               Index           =   76
               Left            =   -74760
               TabIndex        =   99
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
               Index           =   46
               Left            =   -74130
               TabIndex        =   98
               Top             =   990
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
               Index           =   75
               Left            =   -74760
               TabIndex        =   97
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
               Index           =   45
               Left            =   -74130
               TabIndex        =   96
               Top             =   510
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
               Index           =   74
               Left            =   -74790
               TabIndex        =   95
               Top             =   510
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
               Index           =   73
               Left            =   -74805
               TabIndex        =   94
               Top             =   615
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
               Index           =   35
               Left            =   -74145
               TabIndex        =   93
               Top             =   615
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
               Index           =   72
               Left            =   -74775
               TabIndex        =   92
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
               Index           =   34
               Left            =   -74145
               TabIndex        =   91
               Top             =   1095
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
               Index           =   71
               Left            =   -74775
               TabIndex        =   90
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
               Index           =   33
               Left            =   -74145
               TabIndex        =   89
               Top             =   1560
               Width           =   90
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "3."
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
               Index           =   67
               Left            =   -74880
               TabIndex        =   88
               Top             =   1020
               Width           =   150
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "2."
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
               Index           =   66
               Left            =   -74880
               TabIndex        =   87
               Top             =   765
               Width           =   150
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1."
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
               Index           =   65
               Left            =   -74880
               TabIndex        =   86
               Top             =   510
               Width           =   150
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "3."
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
               Index           =   64
               Left            =   -74895
               TabIndex        =   85
               Top             =   1035
               Width           =   150
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "2."
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
               Index           =   63
               Left            =   -74895
               TabIndex        =   84
               Top             =   780
               Width           =   150
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1."
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
               Index           =   62
               Left            =   -74895
               TabIndex        =   83
               Top             =   525
               Width           =   150
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Invoice Prefix............................."
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
            Left            =   405
            TabIndex        =   171
            Top             =   2895
            Width           =   3300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Negative Stock Site Wise................."
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
            Left            =   405
            TabIndex        =   169
            Top             =   750
            Width           =   3765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Separate Warranty Requisition..............."
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
            Index           =   87
            Left            =   405
            TabIndex        =   166
            Top             =   1950
            Width           =   3525
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Details in Sale Document.................."
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
            Index           =   103
            Left            =   405
            TabIndex        =   165
            Top             =   2190
            Width           =   3615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VAT % On Lubricant................................................"
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
            Left            =   255
            TabIndex        =   164
            Top             =   3585
            Visible         =   0   'False
            Width           =   4605
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zero Amt. Bill Creation Stop....................."
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
            Left            =   405
            TabIndex        =   163
            Top             =   2655
            Width           =   3690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local Sale Tax Form Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   12
            Left            =   7545
            TabIndex        =   141
            Top             =   2415
            Width           =   2295
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Form............"
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
            Left            =   6210
            TabIndex        =   140
            Top             =   2670
            Width           =   1530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Issue on Negative Stock......................."
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
            Index           =   18
            Left            =   405
            TabIndex        =   139
            Top             =   510
            Width           =   3465
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Re Sale Tax % ......................................................"
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
            Index           =   17
            Left            =   6210
            TabIndex        =   138
            Top             =   2190
            Width           =   4575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOT Y/N................................................"
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
            Index           =   15
            Left            =   405
            TabIndex        =   137
            Top             =   1245
            Width           =   3585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOT On...................................................."
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
            Left            =   405
            TabIndex        =   136
            Top             =   1725
            Width           =   3765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Govt Tax Form...."
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
            Index           =   13
            Left            =   6210
            TabIndex        =   135
            Top             =   2910
            Width           =   1515
         End
         Begin VB.Label LblColon 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[A]"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   62
            Left            =   8760
            TabIndex        =   134
            Top             =   1230
            Width           =   270
         End
         Begin VB.Label LblColon 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[B]"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   59
            Left            =   9555
            TabIndex        =   133
            Top             =   1230
            Width           =   270
         End
         Begin VB.Label LblColon 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[C]"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   54
            Left            =   10425
            TabIndex        =   132
            Top             =   1230
            Width           =   285
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Profit % in Warranty Amount ..................................."
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
            Index           =   38
            Left            =   6210
            TabIndex        =   131
            Top             =   1710
            Width           =   4635
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Profit % in General Surcharge................................."
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
            Index           =   106
            Left            =   6210
            TabIndex        =   130
            Top             =   1470
            Width           =   4560
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lubricant Grade............................"
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
            Index           =   107
            Left            =   6210
            TabIndex        =   129
            Top             =   510
            Width           =   3045
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consumable Grade...................."
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
            Index           =   108
            Left            =   6210
            TabIndex        =   128
            Top             =   750
            Width           =   2850
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tool Grade ................................"
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
            Index           =   109
            Left            =   6210
            TabIndex        =   127
            Top             =   1005
            Width           =   2925
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part Security Grade Day......."
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
            Index           =   92
            Left            =   6210
            TabIndex        =   126
            Top             =   1230
            Width           =   2520
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Taxable to Taxpaid Rate Conversion........"
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
            Index           =   105
            Left            =   1395
            TabIndex        =   125
            Top             =   4125
            Visible         =   0   'False
            Width           =   3600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Merge Gen. Surcharge to TB Sales A/c...."
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
            Index           =   110
            Left            =   405
            TabIndex        =   124
            Top             =   990
            Width           =   3555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Principal Party Type ............................."
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
            Left            =   390
            TabIndex        =   123
            Top             =   1470
            Width           =   3510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Godown............................"
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
            Left            =   6210
            TabIndex        =   122
            Top             =   270
            Width           =   3045
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MRP Applicable :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   111
            Left            =   9510
            TabIndex        =   121
            Top             =   4125
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Designation :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   85
            Left            =   5640
            TabIndex        =   120
            Top             =   4110
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Store Incharge Name :"
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
            Index           =   86
            Left            =   4860
            TabIndex        =   119
            Top             =   3810
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOT % .................................................................."
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
            Index           =   104
            Left            =   6210
            TabIndex        =   118
            Top             =   1950
            Width           =   4605
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Surch %"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   91
            Left            =   10545
            TabIndex        =   117
            Top             =   2415
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax %"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   84
            Left            =   9915
            TabIndex        =   116
            Top             =   2415
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "General Surcharge % on Spare Sale......."
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
            Index           =   90
            Left            =   405
            TabIndex        =   115
            Top             =   2445
            Width           =   3555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gate Pass on Spare Invoice ...................."
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
            Index           =   89
            Left            =   405
            TabIndex        =   114
            Top             =   270
            Width           =   3645
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Document Footers"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   83
            Left            =   60
            TabIndex        =   113
            Top             =   4200
            Width           =   1560
         End
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
         Height          =   270
         Index           =   10
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   47
         ToolTipText     =   " "
         Top             =   3330
         Width           =   4230
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
         Height          =   270
         Index           =   8
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   44
         ToolTipText     =   " "
         Top             =   2760
         Width           =   4230
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
         Height          =   270
         Index           =   3
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   39
         ToolTipText     =   " "
         Top             =   1335
         Visible         =   0   'False
         Width           =   4230
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
         Height          =   270
         Index           =   9
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   45
         ToolTipText     =   " "
         Top             =   3045
         Width           =   4230
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
         Height          =   270
         Index           =   2
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   38
         ToolTipText     =   " "
         Top             =   1057
         Width           =   4230
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
         Height          =   270
         Index           =   4
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   40
         ToolTipText     =   " "
         Top             =   1620
         Width           =   4230
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
         Height          =   270
         Index           =   5
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   41
         ToolTipText     =   " "
         Top             =   1905
         Width           =   4230
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
         Height          =   270
         Index           =   6
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   42
         ToolTipText     =   " "
         Top             =   2190
         Width           =   4230
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
         Height          =   270
         Index           =   7
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   43
         ToolTipText     =   " "
         Top             =   2475
         Width           =   4230
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
         Height          =   270
         Index           =   14
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   55
         Top             =   4470
         Width           =   4230
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   15
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   57
         Top             =   4755
         Width           =   4230
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
         Height          =   270
         Index           =   1
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   37
         ToolTipText     =   " "
         Top             =   765
         Visible         =   0   'False
         Width           =   4230
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
         Height          =   270
         Index           =   11
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   49
         ToolTipText     =   " "
         Top             =   3615
         Width           =   4230
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
         Height          =   270
         Index           =   13
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   53
         ToolTipText     =   " "
         Top             =   4185
         Width           =   4230
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
         Height          =   270
         Index           =   12
         Left            =   -73050
         MaxLength       =   50
         TabIndex        =   51
         ToolTipText     =   " "
         Top             =   3900
         Width           =   4230
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pur.Trans. A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   18
         Left            =   -68595
         TabIndex        =   162
         Top             =   810
         Width           =   1140
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Tax A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   4
         Left            =   -68595
         TabIndex        =   161
         Top             =   495
         Width           =   1020
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taxpaid A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   20
         Left            =   -68595
         TabIndex        =   160
         Top             =   2175
         Width           =   930
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taxpaid A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   0
         Left            =   -68595
         TabIndex        =   159
         Top             =   3915
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spare Sale A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   7
         Left            =   -68595
         TabIndex        =   158
         Top             =   1065
         Width           =   1200
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oil Purchase A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   24
         Left            =   -68595
         TabIndex        =   157
         Top             =   3375
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taxable A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   16
         Left            =   -68595
         TabIndex        =   156
         Top             =   3645
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spare Purchase A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   15
         Left            =   -68595
         TabIndex        =   155
         Top             =   2745
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taxpaid A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   14
         Left            =   -68595
         TabIndex        =   154
         Top             =   3030
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taxable A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   13
         Left            =   -68595
         TabIndex        =   153
         Top             =   1860
         Width           =   930
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Oil Sale A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   12
         Left            =   -68595
         TabIndex        =   152
         Top             =   1620
         Width           =   930
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taxpaid A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   11
         Left            =   -68595
         TabIndex        =   151
         Top             =   1305
         Width           =   930
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Misc. Charges A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   8
         Left            =   -74805
         TabIndex        =   81
         Top             =   1620
         Width           =   1470
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Surch. on Tax A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   27
         Left            =   -74805
         TabIndex        =   80
         Top             =   2760
         Width           =   1395
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Surch. on CST A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   5
         Left            =   -74805
         TabIndex        =   79
         Top             =   3330
         Width           =   1470
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Central Tax A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   6
         Left            =   -74805
         TabIndex        =   78
         Top             =   3045
         Width           =   1215
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Tax A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Index           =   34
         Left            =   -74805
         TabIndex        =   77
         Top             =   2475
         Width           =   1065
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General Surch. A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   40
         Left            =   -74805
         TabIndex        =   76
         Top             =   2190
         Width           =   1515
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transportation A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   37
         Left            =   -74805
         TabIndex        =   75
         Top             =   1905
         Width           =   1485
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sundry Debtors Grp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   3
         Left            =   -74805
         TabIndex        =   74
         ToolTipText     =   "Press L-> Local or C-> Central"
         Top             =   765
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sundry Creditors Grp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   9
         Left            =   -74805
         TabIndex        =   72
         Top             =   480
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disc. A/c Taxable "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   39
         Left            =   -74805
         TabIndex        =   70
         Top             =   4470
         Width           =   1440
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
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   10
         Left            =   -74805
         TabIndex        =   68
         Top             =   1050
         Width           =   1290
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
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   17
         Left            =   -74805
         TabIndex        =   66
         Top             =   1335
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turn Over Tax A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   21
         Left            =   -74805
         TabIndex        =   64
         Top             =   3615
         Width           =   1410
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spr Round Off A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   22
         Left            =   -74805
         TabIndex        =   62
         Top             =   4185
         Width           =   1440
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Sale Tax A/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   1
         Left            =   -74805
         TabIndex        =   60
         Top             =   3900
         Width           =   1290
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disc. A/cTaxpaid "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   2
         Left            =   -74805
         TabIndex        =   58
         Top             =   4755
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmSyCtrlSpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TAddMode As Boolean
Dim GridKey As Integer
Dim rsGrp As ADODB.Recordset
Dim rsAc As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim Syctrl As ADODB.Recordset
Dim RsSite As ADODB.Recordset
Dim rsPartyType As ADODB.Recordset
Dim RsGodown As ADODB.Recordset
Dim rsPartGrade As ADODB.Recordset
Dim rsForm As ADODB.Recordset
Dim RsPerson As ADODB.Recordset
Dim ExitCtrl As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

Private Const mTOTCaption As String = "TOT on Sub Total (B)"
Private Const mTOTCaption1 As String = "TOT on SubTot(BefTax)"

'grid color scheme
Private Const CellBackColLeave As String = &HC8E8DA
Private Const CellForeColLeave As String = &H0&
Private Const CellBackColEnter As String = &HC0E0FF
Private Const GridBackColorBkg As String = &HBAD3C9

Dim MyIndex As Byte
Private Const SprCreGrp As Byte = 0
Private Const SprDebGrp As Byte = 1
Private Const SprCashAc As Byte = 2
Private Const SprBankAc As Byte = 3
Private Const MiscChrgAc As Byte = 4
Private Const TransportationAc As Byte = 5
Private Const SprGenSurAc As Byte = 6
Private Const LocalTaxAc As Byte = 7
Private Const LocalTaxSurAc As Byte = 8
Private Const CentralTaxAc As Byte = 9
Private Const CentralTaxSurAc As Byte = 10
Private Const TOTaxAc As Byte = 11
Private Const ReSaleTaxAc As Byte = 12
Private Const SprROffAc As Byte = 13
Private Const SprDiscTBAc As Byte = 14
Private Const SprDiscTPAc As Byte = 15
Private Const SprSalTPAc As Byte = 51
Private Const OilSalTBAc As Byte = 52
Private Const OilSalTPAc As Byte = 53
Private Const SprPurTPAc As Byte = 54
Private Const OilPurTBAc As Byte = 55
Private Const OilPurTPAc As Byte = 56
Private Const EntryTaxAc As Byte = 57
Private Const SprPurTransAc As Byte = 58


'Spare Specific
Private Const SGatePassOnSprInv As Byte = 16    'Yes/No
Private Const IPO_Separate As Byte = 17         'Yes/No
Private Const TaxDetOnSprInv As Byte = 18       'Yes/No
Private Const PartyType As Byte = 19            'List
Private Const MRP_YN As Byte = 20               'Yes/No
Private Const SprCounterGodown As Byte = 21     'DGGod
Private Const PartGrade_Lub As Byte = 22        'DGGrade
Private Const PartGrade_Consum As Byte = 23     'DGGrade
Private Const PartGrade_Tool As Byte = 24       'DGGrade
Private Const PartSecurityGradeDay1 As Byte = 25    'Numeric Value
Private Const PartSecurityGradeDay2 As Byte = 26    'Numeric Value
Private Const PartSecurityGradeDay3 As Byte = 27    'Numeric Value
Private Const GenSurProfitPer As Byte = 28    'Numeric Value
Private Const WarrProfitPer As Byte = 29    'Numeric Value
Private Const SprIssOnNegStk As Byte = 30   'Yes/No
Private Const LocalTaxFormSpr As Byte = 31  'Form Name from DGForm
Private Const LST_Rate As Byte = 32         'DGForm.Numeric Value
Private Const LSTSur_Rate As Byte = 33      'DGForm.Numeric Value
Private Const GovtTaxFormSpr As Byte = 34   'Form Name from DGForm
Private Const LST_RateGovt As Byte = 35     'DGForm.Numeric Value
Private Const LSTSur_RateGovt As Byte = 36  'DGForm.Numeric Value
Private Const GenSurChrgOnSpr As Byte = 37  'Numeric Value
Private Const TOT_YN As Byte = 38           'Yes/No
Private Const TOT_On As Byte = 39           'List View
Private Const TOT_Rate As Byte = 40         'Numeric Value
Private Const ReSaleTax_Per As Byte = 41    'Numeric Value
Private Const TBR_to_TPR As Byte = 42       'Yes/No
Private Const MergeGenSur_TB_Sale As Byte = 43  'Yes/No
Private Const Spr_IC_Name As Byte = 44          'Text Value
Private Const Spr_IC_Designation As Byte = 45   'Text Value
Private Const SprPurOrdFooter As Byte = 46 '49
Private Const SprPurRetFooter               As Byte = 47 '48
Private Const EstiInvFooter                 As Byte = 48 '50
Private Const SprInvFooter                  As Byte = 49        'Text Value MEmo 3 lines Max
Private Const SprRetInvFooter               As Byte = 50
Private Const ZeroBillCreate                As Byte = 59
Private Const VatPerOnLube                  As Byte = 60
Private Const CheckNegetiveStockSiteWise    As Byte = 61
Private Const SprTaxInvoicePrefix              As Byte = 62

Private Sub Disp_Text(Enb As Boolean)
Dim i As Integer
    For i = 0 To Txt.Count - 1
        Txt(i).Enabled = Enb
    Next
    Txt(SprPurOrdFooter).Enabled = True
    Txt(SprPurRetFooter).Enabled = True
    Txt(EstiInvFooter).Enabled = True
    Txt(SprInvFooter).Enabled = True
    Txt(SprRetInvFooter).Enabled = True
    
    Txt(SprPurOrdFooter).Locked = True
    Txt(SprPurRetFooter).Locked = True
    Txt(EstiInvFooter).Locked = True
    Txt(SprInvFooter).Locked = True
    Txt(SprRetInvFooter).Locked = True

End Sub

'* Used for intialize grid columns
Private Sub Grid_Ini()
    DGGodown.left = Me.width - (DGGodown.width + mRtScale): DGGodown.top = mTopScale
    DGPartGrade.left = Me.width - (DGPartGrade.width + mRtScale): DGPartGrade.top = mTopScale
    DGForm.left = Me.left: DGForm.top = mTopScale
    DGPartyType.left = Me.width - (DGPartyType.width + mRtScale): DGPartyType.top = mTopScale
    DGGrp.left = Me.width - (DGGrp.width + mRtScale): DGGrp.top = mTopScale
    DGAc.left = Me.width - (DGAc.width + mRtScale): DGAc.top = mTopScale
End Sub

Private Sub Grid_Hide()
    If FrmList.Visible Then FrmList.Visible = False
    If DGGodown.Visible Then DGGodown.Visible = False
    If DGPartGrade.Visible Then DGPartGrade.Visible = False
    If DGForm.Visible Then DGForm.Visible = False
    If DGPartyType.Visible Then DGPartyType.Visible = False
    If DGAc.Visible = True Then DGAc.Visible = False
    If DGGrp.Visible = True Then DGGrp.Visible = False
End Sub

Private Sub MoveRec()
Dim Master1 As ADODB.Recordset
On Error Resume Next 'by lps at Cuttack GoTo ELoop

Txt(SGatePassOnSprInv) = IIf(Syctrl!GatePassOnSprInv = 1, "Yes", "No")
Txt(IPO_Separate) = IIf(Syctrl!IPO_Separate = 1, "Yes", "No")
Txt(TaxDetOnSprInv) = IIf(Syctrl!TaxDetOnSprInv = 1, "Yes", "No")
GSQL = "Select Description from SubGroupType where Party_Type=" & Syctrl!PartyType & ""
Txt(PartyType) = IIf(IsNull(Syctrl!PartyType), "", GCn.Execute(GSQL).Fields(0).Value)
'Txt(MRP_YN) = IIf(Syctrl!PartyType = 1, "Yes", "No")
If Syctrl!SprCounterGodown <> "" Then
    RsGodown.MoveFirst
    RsGodown.FIND ("Code ='" & Syctrl!SprCounterGodown & "'")
    Txt(SprCounterGodown).Tag = Syctrl!SprCounterGodown
    Txt(SprCounterGodown) = IIf(RsGodown.EOF, "", RsGodown!Name)
Else
    Txt(SprCounterGodown) = ""
    Txt(SprCounterGodown).Tag = ""
End If
If RsGodown.RecordCount > 0 And RsGodown.EOF Then RsGodown.MoveFirst

If Syctrl!PartGrade_Lub <> "" Then
    rsPartGrade.MoveFirst
    rsPartGrade.FIND ("Code ='" & Syctrl!PartGrade_Lub & "'")
    Txt(PartGrade_Lub).Tag = Syctrl!PartGrade_Lub
    Txt(PartGrade_Lub) = IIf(rsPartGrade.EOF, "", rsPartGrade!Name)
Else
    Txt(PartGrade_Lub) = ""
    Txt(PartGrade_Lub).Tag = ""
End If
If Syctrl!PartGrade_Consum <> "" Then
    rsPartGrade.MoveFirst
    rsPartGrade.FIND ("Code ='" & Syctrl!PartGrade_Consum & "'")
    Txt(PartGrade_Consum).Tag = Syctrl!PartGrade_Consum
    Txt(PartGrade_Consum) = IIf(rsPartGrade.EOF, "", rsPartGrade!Name)
Else
    Txt(PartGrade_Consum).Tag = ""
    Txt(PartGrade_Consum) = ""
End If
If Syctrl!PartGrade_Tool <> "" Then
    rsPartGrade.MoveFirst
    rsPartGrade.FIND ("Code ='" & Syctrl!PartGrade_Tool & "'")
    Txt(PartGrade_Tool).Tag = Syctrl!PartGrade_Tool
    Txt(PartGrade_Tool) = IIf(rsPartGrade.EOF, "", rsPartGrade!Name)
Else
    Txt(PartGrade_Tool).Tag = ""
    Txt(PartGrade_Tool) = ""
End If
If rsPartGrade.RecordCount > 0 And rsPartGrade.EOF Then rsPartGrade.MoveFirst

Txt(PartSecurityGradeDay1) = IIf(IsNull(Syctrl!PartSecurityGradeDay1), "", Syctrl!PartSecurityGradeDay1)
Txt(PartSecurityGradeDay2) = IIf(IsNull(Syctrl!PartSecurityGradeDay2), "", Syctrl!PartSecurityGradeDay2)
Txt(PartSecurityGradeDay3) = IIf(IsNull(Syctrl!PartSecurityGradeDay3), "", Syctrl!PartSecurityGradeDay3)
Txt(GenSurProfitPer) = IIf(IsNull(Syctrl!GenSurProfitPer), "", Syctrl!GenSurProfitPer)
Txt(WarrProfitPer) = IIf(IsNull(Syctrl!WarrProfitPer), "", Syctrl!WarrProfitPer)
Txt(SprIssOnNegStk) = IIf(IsNull(Syctrl!SprIssOnNegStk), "", IIf(Syctrl!SprIssOnNegStk = 1, "Yes", "No"))
Txt(CheckNegetiveStockSiteWise) = IIf(IsNull(Syctrl!CheckNegetiveStockSiteWise), "", IIf(Syctrl!CheckNegetiveStockSiteWise = 1, "Yes", "No"))
If Not IsNull(Syctrl!LocalTaxFormSpr) Then
    GSQL = "Select Form_Desc from TaxForms where Form_Code='" & Syctrl!LocalTaxFormSpr & "'"
    rsForm.MoveFirst
    rsForm.FIND ("Code ='" & Syctrl!LocalTaxFormSpr & "'")
    If rsForm.EOF = False Then
        Txt(LocalTaxFormSpr).Tag = Syctrl!LocalTaxFormSpr
        Txt(LocalTaxFormSpr) = rsForm!Name
        Txt(LST_Rate) = rsForm!Tax_Per
        Txt(LSTSur_Rate) = rsForm!Tax_Sur_Per
    Else
        Txt(LocalTaxFormSpr) = ""
        Txt(LocalTaxFormSpr).Tag = ""
        Txt(LST_Rate) = ""
        Txt(LSTSur_Rate) = ""
    End If
Else
    Txt(LocalTaxFormSpr) = ""
    Txt(LocalTaxFormSpr).Tag = ""
    Txt(LST_Rate) = ""
    Txt(LSTSur_Rate) = ""
End If
If Not IsNull(Syctrl!GovtTaxFormSpr) Then
    GSQL = "Select Form_Desc from TaxForms where Form_Code='" & Syctrl!GovtTaxFormSpr & "'"
    rsForm.MoveFirst
    rsForm.FIND ("Code ='" & Syctrl!GovtTaxFormSpr & "'")
    If rsForm.EOF = False Then
        Txt(GovtTaxFormSpr).Tag = Syctrl!GovtTaxFormSpr
        Txt(GovtTaxFormSpr) = rsForm!Name
        Txt(LST_RateGovt) = rsForm!Tax_Per
        Txt(LSTSur_RateGovt) = rsForm!Tax_Sur_Per
    Else
        Txt(GovtTaxFormSpr) = ""
        Txt(GovtTaxFormSpr).Tag = ""
        Txt(LST_RateGovt) = ""
        Txt(LSTSur_RateGovt) = ""
    End If
Else
    Txt(LocalTaxFormSpr) = ""
    Txt(LocalTaxFormSpr).Tag = ""
    Txt(LST_RateGovt) = ""
    Txt(LSTSur_RateGovt) = ""
End If
If rsForm.RecordCount > 0 And rsForm.EOF Then rsForm.MoveFirst

Txt(GenSurChrgOnSpr) = IIf(IsNull(Syctrl!GenSurChrgOnSpr), "", IIf(Syctrl!GenSurChrgOnSpr = 1, "Yes", "No"))
Txt(TOT_YN) = IIf(IsNull(Syctrl!TOT_YN), "", IIf(Syctrl!TOT_YN = 1, "Yes", "No"))
Txt(TOT_On) = IIf(IsNull(Syctrl!TOT_On) Or Syctrl!TOT_On = 0, mTOTCaption, mTOTCaption1)
Txt(TOT_Rate) = IIf(IsNull(Syctrl!TOT_Rate) Or Syctrl!TOT_Rate = 0, "", Format(Syctrl!TOT_Rate, "0.00"))
Txt(ReSaleTax_Per) = IIf(IsNull(Syctrl!ReSaleTax_Per) Or Syctrl!ReSaleTax_Per = 0, "", Format(Syctrl!ReSaleTax_Per, "0.00"))
Txt(TBR_to_TPR) = IIf(IsNull(Syctrl!TBR_to_TPR), "", IIf(Syctrl!TBR_to_TPR = 1, "Yes", "No"))
Txt(MergeGenSur_TB_Sale) = IIf(IsNull(Syctrl!MergeGenSur_TB_Sale), "", IIf(Syctrl!MergeGenSur_TB_Sale = 1, "Yes", "No"))
Txt(Spr_IC_Name) = IIf(IsNull(Syctrl!Spr_IC_Name), "", Syctrl!Spr_IC_Name)
Txt(Spr_IC_Designation) = IIf(IsNull(Syctrl!Spr_IC_Designation), "", Syctrl!Spr_IC_Designation)
Txt(SprPurOrdFooter) = IIf(IsNull(Syctrl!SprPurOrdFooter), "", Syctrl!SprPurOrdFooter)
Txt(SprPurRetFooter) = IIf(IsNull(Syctrl!SprPurRetFooter), "", Syctrl!SprPurRetFooter)
Txt(EstiInvFooter) = IIf(IsNull(Syctrl!EstiInvFooter), "", Syctrl!EstiInvFooter)
Txt(SprInvFooter) = IIf(IsNull(Syctrl!SprInvFooter), "", Syctrl!SprInvFooter)
Txt(SprRetInvFooter) = IIf(IsNull(Syctrl!SprRetInvFooter), "", Syctrl!SprRetInvFooter)
Txt(ZeroBillCreate) = IIf(IsNull(Syctrl!ZeroBill), "", IIf(Syctrl!ZeroBill = 1, "Yes", "No"))
Txt(VatPerOnLube) = VNull(Syctrl!VatPerOnLube)
Txt(SprTaxInvoicePrefix) = XNull(Syctrl!SprTaxInvPrefix)
'*******
    Txt(SprCreGrp) = ""
    Txt(SprCreGrp).Tag = ""
    Txt(SprDebGrp) = ""
    Txt(SprDebGrp).Tag = ""
    Txt(SprCashAc) = ""
    Txt(SprCashAc).Tag = ""
    Txt(SprBankAc) = ""
    Txt(SprBankAc).Tag = ""
    Txt(MiscChrgAc) = ""
    Txt(MiscChrgAc).Tag = ""
    Txt(TransportationAc) = ""
    Txt(TransportationAc).Tag = ""
    Txt(SprGenSurAc) = ""
    Txt(SprGenSurAc).Tag = ""
    Txt(LocalTaxAc) = ""
    Txt(LocalTaxAc).Tag = ""
    Txt(LocalTaxSurAc) = ""
    Txt(LocalTaxSurAc).Tag = ""
    Txt(CentralTaxAc) = ""
    Txt(CentralTaxAc).Tag = ""
    Txt(CentralTaxSurAc) = ""
    Txt(CentralTaxSurAc).Tag = ""
    Txt(TOTaxAc) = ""
    Txt(TOTaxAc).Tag = ""
    Txt(ReSaleTaxAc) = ""
    Txt(ReSaleTaxAc).Tag = ""
    Txt(SprROffAc) = ""
    Txt(SprROffAc).Tag = ""
    Txt(SprDiscTBAc) = ""
    Txt(SprDiscTBAc).Tag = ""
    Txt(SprDiscTPAc) = ""
    Txt(SprDiscTPAc).Tag = ""
    Txt(SprSalTPAc) = ""
    Txt(SprSalTPAc).Tag = ""
    Txt(OilSalTPAc) = ""
    Txt(OilSalTPAc).Tag = ""
    Txt(OilSalTBAc) = ""
    Txt(OilSalTBAc).Tag = ""
    Txt(EntryTaxAc) = ""
    Txt(EntryTaxAc).Tag = ""
    Txt(SprPurTransAc) = ""
    Txt(SprPurTransAc).Tag = ""

'***
'If Master.RecordCount <= 0 Then GoTo ExitLoop
If Master!SprCre_Grp <> Null Or Master!SprCre_Grp <> "" Then
    Txt(SprCreGrp) = GCnFaS.Execute("Select GroupName from AcGroup where GroupCode='" & Master!SprCre_Grp & "'").Fields(0).Value
    Txt(SprCreGrp).Tag = Master!SprCre_Grp
End If
If Master!SprDeb_Grp <> Null Or Master!SprDeb_Grp <> "" Then
    Txt(SprDebGrp) = GCnFaS.Execute("Select GroupName from AcGroup where GroupCode='" & Master!SprDeb_Grp & "'").Fields(0).Value
    Txt(SprDebGrp).Tag = Master!SprDeb_Grp
End If
Set Master1 = New Recordset
Master1.CursorLocation = adUseClient
Master1.Open "Select SubCode,Name from SubGroup Order by SubCode", GCnFaS, adOpenStatic, adLockReadOnly
If Master!SprCash_Ac <> Null Or Master!SprCash_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SprCash_Ac & "'")
    Txt(SprCashAc) = Master1!Name
    Txt(SprCashAc).Tag = Master1!SubCode
End If
If Master!SprBank_Ac <> Null Or Master!SprBank_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SprBank_Ac & "'")
    Txt(SprBankAc) = Master1!Name
    Txt(SprBankAc).Tag = Master1!SubCode
End If
If Master!MiscChrg_Ac <> Null Or Master!MiscChrg_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!MiscChrg_Ac & "'")
    Txt(MiscChrgAc) = Master1!Name
    Txt(MiscChrgAc).Tag = Master1!SubCode
End If
If Master!Transportation_Ac <> Null Or Master!Transportation_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!Transportation_Ac & "'")
    Txt(TransportationAc) = Master1!Name
    Txt(TransportationAc).Tag = Master1!SubCode
End If
If Master!SprGenSur_Ac <> Null Or Master!SprGenSur_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SprGenSur_Ac & "'")
    Txt(SprGenSurAc) = Master1!Name
    Txt(SprGenSurAc).Tag = Master1!SubCode
End If
If Master!LocalTax_Ac <> Null Or Master!LocalTax_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!LocalTax_Ac & "'")
    Txt(LocalTaxAc) = Master1!Name
    Txt(LocalTaxAc).Tag = Master1!SubCode
End If
If Master!LocalTaxSur_Ac <> Null Or Master!LocalTaxSur_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!LocalTaxSur_Ac & "'")
    Txt(LocalTaxSurAc) = Master1!Name
    Txt(LocalTaxSurAc).Tag = Master1!SubCode
End If
If Master!CentralTax_Ac <> Null Or Master!CentralTax_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!CentralTax_Ac & "'")
    Txt(CentralTaxAc) = Master1!Name
    Txt(CentralTaxAc).Tag = Master1!SubCode
End If
If Master!CentralTaxSur_Ac <> Null Or Master!CentralTaxSur_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!CentralTaxSur_Ac & "'")
    Txt(CentralTaxSurAc) = Master1!Name
    Txt(CentralTaxSurAc).Tag = Master1!SubCode
End If
If Master!TOTax_Ac <> Null Or Master!TOTax_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!TOTax_Ac & "'")
    Txt(TOTaxAc) = Master1!Name
    Txt(TOTaxAc).Tag = Master1!SubCode
End If
If Master!ReSaleTax_Ac <> Null Or Master!ReSaleTax_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!ReSaleTax_Ac & "'")
    Txt(ReSaleTaxAc) = Master1!Name
    Txt(ReSaleTaxAc).Tag = Master1!SubCode
End If
If Master!SprROff_Ac <> Null Or Master!SprROff_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SprROff_Ac & "'")
    Txt(SprROffAc) = Master1!Name
    Txt(SprROffAc).Tag = Master1!SubCode
End If
If Master!SprDiscTB_Ac <> Null Or Master!SprDiscTB_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SprDiscTB_Ac & "'")
    Txt(SprDiscTBAc) = Master1!Name
    Txt(SprDiscTBAc).Tag = Master1!SubCode
End If
If Master!SprDiscTP_Ac <> Null Or Master!SprDiscTP_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SprDiscTP_Ac & "'")
    Txt(SprDiscTPAc) = Master1!Name
    Txt(SprDiscTPAc).Tag = Master1!SubCode
End If
If Master!SprSalTP_Ac <> Null Or Master!SprSalTP_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SprSalTP_Ac & "'")
    Txt(SprSalTPAc) = Master1!Name
    Txt(SprSalTPAc).Tag = Master1!SubCode
End If
If Master!OilSalTP_Ac <> Null Or Master!OilSalTP_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!OilSalTP_Ac & "'")
    Txt(OilSalTPAc) = Master1!Name
    Txt(OilSalTPAc).Tag = Master1!SubCode
End If
If Master!OilSalTB_Ac <> Null Or Master!OilSalTB_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!OilSalTB_Ac & "'")
    Txt(OilSalTBAc) = Master1!Name
    Txt(OilSalTBAc).Tag = Master1!SubCode
End If
'EntryTaxAc
If Master!EntryTax_Ac <> Null Or Master!EntryTax_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!EntryTax_Ac & "'")
    Txt(EntryTaxAc) = Master1!Name
    Txt(EntryTaxAc).Tag = Master1!SubCode
End If
'Spare Purchase Transportation Ac
If Master!SprPurTrans_Ac <> Null Or Master!SprPurTrans_Ac <> "" Then
    Master1.MoveFirst
    Master1.FIND ("SubCode ='" & Master!SprPurTrans_Ac & "'")
    Txt(SprPurTransAc) = Master1!Name
    Txt(SprPurTransAc).Tag = Master1!SubCode
End If
'16-05-03 lps
'If Master!CPSprAc <> Null Or Master!CPSprAc <> "" Then
'    Master1.MoveFirst
'    Master1.FIND ("SubCode ='" & Master!CPSprAc & "'")
'    Txt(SprPurTempAc) = Master1!Name
'    Txt(SprPurTempAc).Tag = Master1!SubCode
'End If
'If Master!CSSprAc <> Null Or Master!CSSprAc <> "" Then
'    Master1.MoveFirst
'    Master1.FIND ("SubCode ='" & Master!CSSprAc & "'")
'    Txt(SprSalTempAc) = Master1!Name
'    Txt(SprSalTempAc).Tag = Master1!SubCode
'End If
'********
'If Master!SprPurTP_Ac <> Null Or Master!SprPurTP_Ac <> "" Then
'    Master1.MoveFirst
'    Master1.FIND ("SubCode ='" & Master!SprPurTP_Ac & "'")
'    Txt(SprPurTPAc) = Master1!Name
'    Txt(SprPurTPAc).Tag = Master1!SubCode
'End If
'If Master!OilPurTB_Ac <> Null Or Master!OilPurTB_Ac <> "" Then
'    Master1.MoveFirst
'    Master1.FIND ("SubCode ='" & Master!OilPurTB_Ac & "'")
'    Txt(OilPurTBAc) = Master1!Name
'    Txt(OilPurTBAc).Tag = Master1!SubCode
'End If
'If Master!OilPurTP_Ac <> Null Or Master!OilPurTP_Ac <> "" Then
'    Master1.MoveFirst
'    Master1.FIND ("SubCode ='" & Master!OilPurTP_Ac & "'")
'    Txt(OilPurTPAc) = Master1!Name
'    Txt(OilPurTPAc).Tag = Master1!SubCode
'End If

'***********Nra Modi For SDT********
If PubSDTYN = 1 Then
    Label3(15).CAPTION = pubTOTCaption
    Txt(TOT_YN) = IIf(VNull(GCn.Execute("Select SDT_YN from Syctrl").Fields(0).Value) = 1, "Yes", "No")
    Label3(14).CAPTION = "SDT ON"
    Txt(TOT_On) = "Taxable Total"
    Label3(104).CAPTION = "SDT %"
    Txt(TOT_Rate) = VNull(GCn.Execute("Select TOT_Rate from Syctrl").Fields(0).Value)
End If
'***********************************

ExitLoop:
'****
Set Master1 = Nothing
    TopCtrl1.tAdd = False
    TopCtrl1.tDel = False
    TopCtrl1.tFirst = False
    TopCtrl1.tPrev = False
    TopCtrl1.tNext = False
    TopCtrl1.tLast = False
    TopCtrl1.tFind = False
    TopCtrl1.tPrn = False
SSTab1.Tab = 0
SSTab3.Tab = 0
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub CmdAcCurBalUpd_Click()
Dim DataPath$

If MsgBox("Are You Sure To Update Current Balance of A/c ? ", vbYesNo + vbCritical + vbDefaultButton2, "Update Current Stock !") = vbYes Then
    CmdAcCurBalUpd.CAPTION = "Updation in progress.."
    CmdAcCurBalUpd.Enabled = False
    
    Dim Rst As ADODB.Recordset, mTrans As Boolean
    
    'GCn.BeginTrans
    'G_FaCn.BeginTrans
    mTrans = True
    If PubBackEnd = "A" Then G_FaCn.Execute ("update SubGroup set Curr_Bal=0")
    GCn.Execute ("update SubGroup set Curr_Bal=0 ")
    
    GSQL = "SELECT Ledger.SubCode,SUM(AmtCr-AmtDr) as CBal " & _
            "FROM Ledger left join SubGroup SG on SG.SubCOde=Ledger.SubCode " & _
            "group by Ledger.subcode,Name"
    Set Rst = G_FaCn.Execute(GSQL)
    If Rst.RecordCount > 0 Then
        Do While Rst.EOF = False
            GCn.Execute ("Update SubGroup set Curr_Bal=" & Rst!CBal & " where SubCode='" & Rst!SubCode & "'")
            If PubBackEnd = "A" Then G_FaCn.Execute ("Update SubGroup set Curr_Bal=" & Rst!CBal & " where SubCode='" & Rst!SubCode & "'")
            Rst.MoveNext
        Loop
    End If
'    DataPath = Pub_DataPath & "\" & PubCenDataPath & "\Automan.mdb;pwd=dtman"
'    Set Rst = G_FaCn.Execute("select SG.SubCode from SubGroup as SG where SubCode not in (Select SubCode from [" & DataPath & "].SubGroup where FirmCode='" & PubFirmCode & "')")
'    If Rst.RecordCount > 0 Then
'        Do Until Rst.EOF
'            GCn.Execute ("Delete From SubGroup where Subcode='" & Rst!SubCode & "'")
'            GCn.Execute ("INSERT INTO SUBGROUP SELECT * FROM [" & PubFADataPath & "].SUBGROUP WHERE SUBCODE = '" & Rst!SubCode & "'")
'            Rst.MoveNext
'        Loop
'    End If
    GCn.Execute ("Drop Table SubGroupAlias")
    GCn.Execute ("Select SubGroup.* into SubGroupAlias from SubGroup")
    
    'GCn.CommitTrans
    'G_FaCn.CommitTrans
    mTrans = False
    CmdAcCurBalUpd.CAPTION = "Updated Successfully"
    MsgBox "Updated SuccessFully"
    CmdAcCurBalUpd.Enabled = False
    Set Rst = Nothing
End If
Exit Sub

ELoop:
'If mTrans Then GCn.RollbackTrans: G_FaCn.RollbackTrans
CmdStkUpd.CAPTION = "Updation Failed"
MsgBox "A/c Current Balance Updation failed, contact Adminstrator!", vbCritical, "Stock Updation"

End Sub
Private Sub CmdStkUpd_Click()
On Error Resume Next
Dim i As Integer, GSQL1$
If MsgBox("Are You Sure To Update Current Stock ? ", vbYesNo + vbCritical + vbDefaultButton2, "Update Current Stock !") = vbYes Then
    CmdStkUpd.CAPTION = "Updation in progress.."
    CmdStkUpd.Enabled = False
    'GCn.BeginTrans
    '***
    GCn.Execute ("update Part Set Cur_TP_Stk=0,Cur_TB_Stk=0,Cur_MRP_TPStk=0,Cur_MRP_TBStk=0 where Div_Code='" & PubDivCode & "'")
    '***
    
    Dim Rst As ADODB.Recordset
    Dim mSQry$, mQRY$
    
    
    mSQry = "Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock " & _
            "WHERE (V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " " & _
            "Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & ") " & _
            "And Part_No=P.Part_No "

    
    If PubBackEnd = "S" Then
        mQRY = "Select P.Part_No as Code, P.Part_No, P.Part_Name As Name, P.Local_Name as LName, P.Unit , P.MRP, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, " & _
                        "(Select PurcDisc_Per From Part_DiscFactor Where DiscFac_Catg=P.Disc_Factor) As PurcDisc_Per, P.ReOrd_Lvl, " & _
                        "(" & mSQry & " And Mrp_Yn=1 And Tax_Yn=1) As Cur_MRP_TBStk, (" & mSQry & " And Mrp_Yn=1 And Tax_Yn=0) As Cur_MRP_TpStk, " & _
                        "(" & mSQry & " And Mrp_Yn=0 And Tax_Yn=1) As Cur_TBStk, (" & mSQry & " And Mrp_Yn=0 And Tax_Yn=0) As Cur_TpStk, " & _
                        "(" & mSQry & ") As CurrStk, P.Min_Lvl, P.Disc_Factor " & _
                        "From Part P Left Join Sp_Stock Stk On P.Part_No=Stk.Part_No " & _
                        "WHERE (V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & " Or Stk.Part_No Is Null)  And Div_Code='" & PubDivCode & "' " & _
                        "Group By P.Part_No, P.Part_Name, P.Local_Name, P.Unit, P.Mrp, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Disc_Factor, P.ReOrd_Lvl, P.Min_Lvl"
    Else
        mQRY = "Select P.Part_No as Code, P.Part_No, P.Part_Name As Name, P.Local_Name as LName, P.Unit , Format(P.MRP,'0.00') As Mrp, Format(P.TB_SRate,'0.00') As TB_SRate, Format(P.Tp_SRate,'0.00') As Tp_SRate, P.Bin_Loca, " & _
                        "(Select PurcDisc_Per From Part_DiscFactor Where DiscFac_Catg=P.Disc_Factor) As PurcDisc_Per, P.ReOrd_Lvl, " & _
                        "(" & mSQry & " And Mrp_Yn=1 And Tax_Yn=1) As Cur_MRP_TBStk, (" & mSQry & " And Mrp_Yn=1 And Tax_Yn=0) As Cur_MRP_TpStk, " & _
                        "(" & mSQry & " And Mrp_Yn=0 And Tax_Yn=1) As Cur_TBStk, (" & mSQry & " And Mrp_Yn=0 And Tax_Yn=0) As Cur_TpStk, " & _
                        "(" & mSQry & ") As CurrStk, P.Min_Lvl, P.Disc_Factor " & _
                        "From Part P Left Join Sp_Stock Stk On P.Part_No=Stk.Part_No " & _
                        "WHERE (V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & " Or Stk.Part_No Is Null)  And Div_Code='" & PubDivCode & "' " & _
                        "Group By P.Part_No, P.Part_Name, P.Local_Name, P.Unit, P.Mrp, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Disc_Factor, P.ReOrd_Lvl, P.Min_Lvl"
    End If
    
    GCn.Execute "Drop Table TmpTbl"
    GCn.Execute ("Select * Into TmpTbl From (" & mQRY & ")")
    
    GCn.Execute ("Update Part, TmpTbl Set Part.Cur_TP_Stk=TmpTbl.Cur_TPStk, " & _
                 "Part.Cur_TB_Stk=TmpTbl.Cur_TBStk, Part.Cur_Mrp_TpStk=TmpTbl.Cur_MRP_TPStk, " & _
                 "Part.Cur_Mrp_TBStk=TmpTbl.Cur_MRP_TbStk where Part.Part_No=TmpTbl.Part_No and Part.Div_Code='" & PubDivCode & "'")
    CmdStkUpd.CAPTION = "Updated Sucessfully"
    CmdStkUpd.Enabled = False
    
    
'    RsPart.Requery
End If
Set Rst = Nothing
MsgBox "Spare Current Stock Updated! Please Reload the Software", vbOKOnly, "Message"
End
Exit Sub

ELoop:
MsgBox err.Description
GCn.RollbackTrans
CmdStkUpd.CAPTION = "Updation Failed"
MsgBox "Stock Updation failed, contact Adminstrator!", vbCritical, "Stock Updation"

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

Private Sub DGForm_Click()
On Error GoTo ELoop
    If rsForm.RecordCount > 0 Then
        Txt(MyIndex).TEXT = rsForm!Name
        Txt(MyIndex).Tag = rsForm!Code
    End If
    Txt(MyIndex).SetFocus
    DGForm.Visible = False
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

Private Sub DGPartGrade_Click()
On Error GoTo ELoop
    If rsPartGrade.RecordCount > 0 Then
        Txt(MyIndex).TEXT = rsPartGrade!Name
        Txt(MyIndex).Tag = rsPartGrade!Code
    End If
    Txt(MyIndex).SetFocus
    DGPartGrade.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGPartyType_Click()
On Error GoTo ELoop
    If rsPartyType.RecordCount > 0 Then
        Txt(MyIndex).TEXT = rsPartyType!Name
        Txt(MyIndex).Tag = rsPartyType!Code
    End If
    Txt(MyIndex).SetFocus
    DGPartyType.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGSite_Click()
On Error GoTo ELoop
    If RsSite.RecordCount > 0 Then
        Txt(MyIndex).TEXT = RsSite!Name
        Txt(MyIndex).Tag = RsSite!Code
    End If
    Txt(MyIndex).SetFocus
    DgSite.Visible = False
Exit Sub
ELoop:
    CheckError

End Sub

Private Sub Form_Activate()
Dim UnLoadFrm As Boolean, MsgStr$
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
Dim i As Byte
    TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini
    If PubSDTYN = 1 Then
        LBL(21) = "S D T A/c Name"
    End If
    For i = 0 To Txt.Count - 1
        Txt(i).BackColor = CtrlBColOrg '&HDFF4F2
        Txt(i).ForeColor = CtrlFColOrg
'        Txt(I).BorderStyle = 1
    Next
    If pubUName = "SA" Then CmdStkUpd.Enabled = True
    If pubUName = "SA" Then CmdAcCurBalUpd.Enabled = True
    
    Set RsGodown = New ADODB.Recordset
    RsGodown.CursorLocation = adUseClient
    RsGodown.Open "Select God_Code as Code, God_Name As Name From Godown where Appli_For=0 Order by God_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGodown.DataSource = RsGodown
    
    Set rsPartGrade = New ADODB.Recordset
    rsPartGrade.CursorLocation = adUseClient
    rsPartGrade.Open "Select PartGrade_Code as Code, PartGrade_Name As Name From Part_Grade Order by PartGrade_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGPartGrade.DataSource = rsPartGrade
    
    Set rsForm = New ADODB.Recordset
    rsForm.CursorLocation = adUseClient
    rsForm.Open "Select Form_Code as Code,Form_Desc As Name,Tax_Per,Tax_Sur_Per From TaxForms Where Spare_YN=1 and Trn_Type='Sale' Order by Form_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGForm.DataSource = rsForm

    Set RsPerson = New ADODB.Recordset
    RsPerson.CursorLocation = adUseClient
    RsPerson.Open "Select Emp_Code as Code, Emp_Name as Name From Emp_Mast Where Emp_Type=0 and (LeftOn Is Null or LeftOn< " & ConvertDate(PubLoginDate) & ") Order By Emp_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGPerson.DataSource = RsPerson

    Set rsPartyType = New ADODB.Recordset
    rsPartyType.CursorLocation = adUseClient
    rsPartyType.Open "Select Party_Type As Code,Description As Name From SubGroupType Order by Description", GCn, adOpenDynamic, adLockOptimistic
    Set DGPartyType.DataSource = rsPartyType

    Set rsGrp = New ADODB.Recordset
    rsGrp.CursorLocation = adUseClient
    rsGrp.Open "Select GroupCode As Code,GroupName As Name,GroupNature,MainGrCode From AcGroup Where MainGrCode<>'999' Order by GroupName", GCnFaS, adOpenDynamic, adLockOptimistic
    Set DGGrp.DataSource = rsGrp
    
    Set rsAc = New ADODB.Recordset
    rsAc.CursorLocation = adUseClient
    rsAc.Open "Select SubCode as Code,Name From SubGroup Order by Name", GCnFaS, adOpenDynamic, adLockOptimistic
    Set DGAc.DataSource = rsAc

    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    Master.Open "Select * from AcControls where Div_Code='" & PubDivCode & "'", GCnFaS, adOpenDynamic, adLockOptimistic
'    Set Master = GCnFaS.Execute("Select * from AcControls where Div_Code='" & PubDivCode & "'")
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
    Set rsGrp = Nothing
    Set rsAc = Nothing
    Set Master = Nothing
    Set RsGodown = Nothing
    Set rsPartGrade = Nothing
    Set rsForm = Nothing
    Set RsPerson = Nothing
    Set rsPartyType = Nothing
    Set Syctrl = Nothing
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
    Txt(SprPurOrdFooter).Locked = False
    Txt(SprPurRetFooter).Locked = False
    Txt(EstiInvFooter).Locked = False
    Txt(SprInvFooter).Locked = False
    Txt(SprRetInvFooter).Locked = False
    Txt(SGatePassOnSprInv).SetFocus
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
    rsPartGrade.Requery
    rsForm.Requery
    RsPerson.Requery
    rsPartyType.Requery
    Syctrl.Requery
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eSave()
Dim mTrans As Boolean, MasterSql$
On Error GoTo ELoop
Grid_Hide
'Apply necessary validations
'If IsValid(Txt(Party), "Party Name") = False Then Exit Sub
    
    GCnFaS.BeginTrans
    GCn.BeginTrans
        mTrans = True
        If TopCtrl1.TopText2 = "Edit" Then   'Edit Bill
            GCn.Execute "update Syctrl set GatePassOnSprInv = " & IIf(Txt(SGatePassOnSprInv) = "Yes", 1, 0) & _
                ", IPO_Separate = " & IIf(Txt(IPO_Separate) = "Yes", 1, 0) & ", TaxDetOnSprInv = " & IIf(Txt(TaxDetOnSprInv) = "Yes", 1, 0) & _
                ", PartyType = " & Val(Txt(PartyType).Tag) & ", SprCounterGodown = '" & Txt(SprCounterGodown).Tag & _
                "',PartGrade_Lub = '" & Txt(PartGrade_Lub).Tag & "',PartGrade_Consum = '" & Txt(PartGrade_Consum).Tag & "',PartGrade_Tool = '" & Txt(PartGrade_Tool).Tag & _
                "',PartSecurityGradeDay1 = " & Val(Txt(PartSecurityGradeDay1)) & ", PartSecurityGradeDay2 = " & Val(Txt(PartSecurityGradeDay2)) & ", PartSecurityGradeDay3 = " & Val(Txt(PartSecurityGradeDay3)) & _
                ", GenSurProfitPer = " & Val(Txt(GenSurProfitPer)) & ", WarrProfitPer = " & Val(Txt(WarrProfitPer)) & ",CheckNegetiveStockSiteWise = " & IIf(Txt(CheckNegetiveStockSiteWise) = "Yes", 1, 0) & ",SprIssOnNegStk = " & IIf(Txt(SprIssOnNegStk) = "Yes", 1, 0) & _
                ", LocalTaxFormSpr = '" & Txt(LocalTaxFormSpr).Tag & "', GovtTaxFormSpr = '" & Txt(GovtTaxFormSpr).Tag & "', GenSurChrgOnSpr = " & IIf(Txt(GenSurChrgOnSpr) = "Yes", 1, 0) & _
                ", TOT_YN = " & IIf(Txt(TOT_YN) = "Yes", 1, 0) & ", TOT_On = " & IIf(Txt(TOT_On) = mTOTCaption, 0, 1) & ", TOT_Rate = " & Val(Txt(TOT_Rate)) & ", ReSaleTax_Per = " & Val(Txt(ReSaleTax_Per)) & _
                ", TBR_to_TPR = " & IIf(Txt(TBR_to_TPR) = "Yes", 1, 0) & ", MergeGenSur_TB_Sale = " & IIf(Txt(MergeGenSur_TB_Sale) = "Yes", 1, 0) & ", Spr_IC_Name = '" & Txt(Spr_IC_Name) & _
                "',Spr_IC_Designation = '" & Txt(Spr_IC_Designation) & "', SprPurOrdFooter = '" & Txt(SprPurOrdFooter) & "', SprPurRetFooter = '" & Txt(SprPurRetFooter) & _
                "',EstiInvFooter = '" & Txt(EstiInvFooter) & "', VatPerOnLube=" & Val(Txt(VatPerOnLube)) & ", SprInvFooter = '" & Txt(SprInvFooter) & "', SprRetInvFooter = '" & Txt(SprRetInvFooter) & "', SprTaxInvPrefix = '" & Txt(SprTaxInvoicePrefix) & "',ZeroBill = " & IIf(Txt(ZeroBillCreate) = "Yes", 1, 0) & ""
            GCn.CommitTrans
            GCnFaS.Execute "Update AcControls Set " _
                & "SprCre_Grp='" & Txt(SprCreGrp).Tag & "',SprDeb_Grp='" & Txt(SprDebGrp).Tag & _
                "',SprCash_Ac='" & Txt(SprCashAc).Tag & "',SprBank_Ac='" & Txt(SprBankAc).Tag & _
                "',MiscChrg_Ac='" & Txt(MiscChrgAc).Tag & "',Transportation_Ac='" & Txt(TransportationAc).Tag & _
                "',SprGenSur_Ac='" & Txt(SprGenSurAc).Tag & "',LocalTax_Ac='" & Txt(LocalTaxAc).Tag & _
                "',LocalTaxSur_Ac='" & Txt(LocalTaxSurAc).Tag & "',CentralTax_Ac='" & Txt(CentralTaxAc).Tag & _
                "',CentralTaxSur_Ac='" & Txt(CentralTaxSurAc).Tag & _
                "',SprDiscTB_Ac='" & Txt(SprDiscTBAc).Tag & "',SprDiscTP_Ac='" & Txt(SprDiscTPAc).Tag & "',TOTax_Ac='" & Txt(TOTaxAc).Tag & _
                "',ReSaleTax_Ac='" & Txt(ReSaleTaxAc).Tag & "',SprROff_Ac='" & Txt(SprROffAc).Tag & _
                "',SprSalTP_Ac='" & Txt(SprSalTPAc).Tag & _
                "',OilSalTB_Ac='" & Txt(OilSalTBAc).Tag & "',OilSalTP_Ac='" & Txt(OilSalTPAc).Tag & _
                "',SprPurTP_Ac='" & Txt(SprPurTPAc).Tag & "',EntryTax_Ac='" & Txt(EntryTaxAc).Tag & _
                "',SprPurTrans_Ac='" & Txt(SprPurTransAc).Tag & "' where Div_Code='" & PubDivCode & "'"
'                "',OilPurTB_Ac='" & Txt(OilPurTBAc).Tag & "',OilpurTP_Ac='" & Txt(OilPurTPAc).Tag & "'"
        End If

    GCnFaS.CommitTrans
    mTrans = False
'    mSearchCode = Txt(Docid)
    Master.Requery
    Syctrl.Requery
'    Master.FIND "SearchCode = '" & mSearchCode & "'"
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    If mTrans = True Then CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim i As Byte
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For i = 0 To Txt.Count - 1
            Txt(i).BackColor = CtrlBColOrg
            Txt(i).ForeColor = CtrlFColOrg
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
    Case PartyType
        If rsPartyType.RecordCount = 0 Or (rsPartyType.EOF = True Or rsPartyType.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsPartyType!Name Then
            rsPartyType.MoveFirst
            rsPartyType.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case SprCounterGodown
        If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsGodown!Name Then
            RsGodown.MoveFirst
            RsGodown.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case PartGrade_Lub, PartGrade_Consum, PartGrade_Tool
        If rsPartGrade.RecordCount = 0 Or (rsPartGrade.EOF = True Or rsPartGrade.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsPartGrade!Name Then
            rsPartGrade.MoveFirst
            rsPartGrade.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case LocalTaxFormSpr, GovtTaxFormSpr
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsForm!Name Then
            rsForm.MoveFirst
            rsForm.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case TOT_On
        If Txt(TOT_YN) = "Yes" Then
            ListArray = Array(mTOTCaption, mTOTCaption1)
            Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 2)
        End If
    Case SprCreGrp, SprDebGrp
        If rsGrp.RecordCount = 0 Or (rsGrp.EOF = True Or rsGrp.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsGrp!Name Then
            rsGrp.MoveFirst
            rsGrp.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case SprCashAc, SprBankAc, MiscChrgAc, TransportationAc, SprGenSurAc, LocalTaxAc, _
        LocalTaxSurAc, CentralTaxAc, CentralTaxSurAc, TOTaxAc, SprROffAc, SprDiscTBAc, _
        SprDiscTPAc, ReSaleTaxAc    ', SprPurTempAc, SprSalTempAc
        'Help grid Positioning
        DGAc.left = Me.width - (DGAc.width + mRtScale): DGAc.top = mTopScale
        'eof positioning
        If rsAc.RecordCount = 0 Or (rsAc.EOF = True Or rsAc.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsAc!Name Then
            rsAc.MoveFirst
            rsAc.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
        
    Case SprPurTransAc, EntryTaxAc, SprSalTPAc, OilSalTBAc, OilSalTPAc ', SprPurTPAc, OilPurTBAc, OilPurTPAc
        'Help grid Positioning
        DGAc.left = mLtScale: DGAc.top = mTopScale
        'eof positioning
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
        Case PartyType
            DGridTxtKeyDown DGPartyType, Txt, Index, rsPartyType, KeyCode, False, 1
        Case SprCounterGodown
            DGridTxtKeyDown DGGodown, Txt, Index, RsGodown, KeyCode, False, 1, frmGodown
        Case PartGrade_Lub, PartGrade_Consum, PartGrade_Tool
            DGridTxtKeyDown DGPartGrade, Txt, Index, rsPartGrade, KeyCode, False, 1, frmPartGrade
        Case LocalTaxFormSpr, GovtTaxFormSpr
            DGridTxtKeyDown DGForm, Txt, Index, rsForm, KeyCode, False, 1, frmTaxForms
        Case TOT_On
            If Txt(TOT_YN) = "Yes" Then
                ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 700
            End If
        Case SprCreGrp, SprDebGrp
            DGridTxtKeyDown DGGrp, Txt, Index, rsGrp, KeyCode, False, 1, frmGrEnt
        Case SprPurTransAc, EntryTaxAc, SprCashAc, SprBankAc, MiscChrgAc, TransportationAc, SprGenSurAc, LocalTaxAc, _
            LocalTaxSurAc, CentralTaxAc, CentralTaxSurAc, TOTaxAc, SprROffAc, SprDiscTBAc, _
            SprDiscTPAc, ReSaleTaxAc, SprSalTPAc, _
            OilSalTBAc, OilSalTPAc ', SprPurTPAc, OilPurTBAc, OilPurTPAc ', SprPurTempAc, SprSalTempAc
            DGridTxtKeyDown DGAc, Txt, Index, rsAc, KeyCode, False, 1, frmSubGroup
    End Select
    If Txt(Index).MultiLine = False Then
        If FrmList.Visible = False And DGGodown.Visible = False _
            And DGPartGrade.Visible = False And DGForm.Visible = False _
            And DGPartyType.Visible = False And DGGrp.Visible = False And DGAc.Visible = False Then    'Arrow Key
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = OilSalTPAc Then
               If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
            End If
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> OilSalTPAc Then Ctrl_DownKeyDown KeyCode, Shift
            If TopCtrl1.TopText2.CAPTION = "Edit" Then
                If Index <> SprCreGrp And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
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
    Case SGatePassOnSprInv, IPO_Separate, TaxDetOnSprInv, MRP_YN, CheckNegetiveStockSiteWise, SprIssOnNegStk, _
        TOT_YN, TBR_to_TPR, MergeGenSur_TB_Sale, ZeroBillCreate
        If UCase(Chr(keyascii)) = "Y" Then
            Txt(Index).TEXT = "Yes"
            keyascii = 0
        Else    'If KeyAscii = 87 Or KeyAscii = 119 Then   ' W/w
            If keyascii <> vbKeyReturn Then
                Txt(Index).TEXT = "No"
                keyascii = 0
            End If
        End If
    Case PartSecurityGradeDay1, PartSecurityGradeDay2, PartSecurityGradeDay3
        NumPress Txt(Index), keyascii, 3, 0
    Case GenSurProfitPer, WarrProfitPer
        NumPress Txt(Index), keyascii, 2, 3
    Case GenSurChrgOnSpr, TOT_Rate, ReSaleTax_Per, VatPerOnLube
        NumPress Txt(Index), keyascii, 2, 2
    Case PartyType
        If DGPartyType.Visible = True Then DGridTxtKeyPress Txt, PartyType, rsPartyType, keyascii, "Name"
    Case SprCounterGodown
        If DGGodown.Visible = True Then DGridTxtKeyPress Txt, SprCounterGodown, RsGodown, keyascii, "Name"
    Case PartGrade_Lub, PartGrade_Consum, PartGrade_Tool
        If DGPartGrade.Visible = True Then DGridTxtKeyPress Txt, Index, rsPartGrade, keyascii, "Name"
    Case LocalTaxFormSpr, GovtTaxFormSpr
        If DGForm.Visible = True Then DGridTxtKeyPress Txt, Index, rsForm, keyascii, "Name"
    Case SprCreGrp, SprDebGrp
        If DGGrp.Visible = True Then DGridTxtKeyPress Txt, Index, rsGrp, keyascii, "Name"
    Case SprPurTransAc, EntryTaxAc, SprCashAc, SprBankAc, MiscChrgAc, TransportationAc, SprGenSurAc, LocalTaxAc, _
        LocalTaxSurAc, CentralTaxAc, CentralTaxSurAc, TOTaxAc, SprROffAc, SprDiscTBAc, _
        SprDiscTPAc, ReSaleTaxAc, SprSalTPAc, _
        OilSalTBAc, OilSalTPAc ', SprPurTPAc, OilPurTBAc, OilPurTPAc ', SprPurTempAc, SprSalTempAc
        If DGAc.Visible = True Then DGridTxtKeyPress Txt, Index, rsAc, keyascii, "Name"
  
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case TOT_On
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
End Select
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
        Case SprCounterGodown
            If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or Txt(Index).TEXT = "" Then
                Txt(Index) = ""
                Txt(Index).Tag = ""
            Else
                Txt(Index).Tag = RsGodown!Code
                Txt(Index) = RsGodown!Name
            End If
        Case PartGrade_Lub, PartGrade_Consum, PartGrade_Tool
            If Txt(Index) = "" Then Txt(Index).Tag = ""
        Case LocalTaxFormSpr
            If Txt(Index) <> "" And (rsForm.RecordCount > 0 Or (rsForm.EOF = False Or rsForm.BOF = False)) Then
                Txt(LST_Rate) = IIf(Txt(Index).TEXT <> "", rsForm!Tax_Per, "")
                Txt(LSTSur_Rate) = IIf(Txt(Index).TEXT <> "", rsForm!Tax_Sur_Per, "")
            Else
                Txt(Index).Tag = ""
                Txt(LST_Rate) = ""
                Txt(LSTSur_Rate) = ""
            End If
        Case GovtTaxFormSpr
            If Txt(Index) <> "" And (rsForm.RecordCount > 0 Or (rsForm.EOF = False Or rsForm.BOF = False)) Then
                Txt(LST_RateGovt) = IIf(Txt(Index).TEXT <> "", rsForm!Tax_Per, "")
                Txt(LSTSur_RateGovt) = IIf(Txt(Index).TEXT <> "", rsForm!Tax_Sur_Per, "")
            Else
                Txt(Index).Tag = ""
                Txt(LST_RateGovt) = ""
                Txt(LSTSur_RateGovt) = ""
            End If
        Case PartyType
            If rsPartyType.RecordCount > 0 Or (rsPartyType.EOF = False Or rsPartyType.BOF = False) Then
                If Txt(Index).TEXT <> "" Then
                    Txt(Index).TEXT = rsPartyType!Name
                    Txt(Index).Tag = rsPartyType!Code
                Else
                    Txt(Index).TEXT = ""
                    Txt(Index).Tag = ""
                End If
            End If
        Case SprCreGrp, SprDebGrp
            If rsGrp.RecordCount > 0 Or (rsGrp.EOF = False Or rsGrp.BOF = False) Then
                If Txt(Index).TEXT <> "" Then
                    Txt(Index).TEXT = rsGrp!Name
                    Txt(Index).Tag = rsGrp!Code
                End If
            Else
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            End If
        Case SprPurTransAc, EntryTaxAc, SprCashAc, SprBankAc, MiscChrgAc, TransportationAc, SprGenSurAc, LocalTaxAc, _
            LocalTaxSurAc, CentralTaxAc, CentralTaxSurAc, TOTaxAc, SprROffAc, SprDiscTBAc, _
            SprDiscTPAc, ReSaleTaxAc, SprSalTPAc, _
            OilSalTBAc, OilSalTPAc, SprPurTPAc, OilPurTBAc, OilPurTPAc ', SprPurTempAc, SprSalTempAc
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

