VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form FaGrEnt 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Group Accounts Entry"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   Icon            =   "FaGrEnt.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   9675
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
      Index           =   6
      Left            =   1890
      MaxLength       =   15
      TabIndex        =   6
      ToolTipText     =   "Group behavior"
      Top             =   2265
      Width           =   1740
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   661
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   2370
      Left            =   7215
      TabIndex        =   16
      Top             =   4020
      Visible         =   0   'False
      Width           =   2325
      Begin MSComctlLib.ListView ListView 
         Height          =   2310
         Left            =   15
         TabIndex        =   17
         Top             =   30
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   4075
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
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid DGUnderAc 
      Height          =   3330
      Left            =   2115
      Negotiate       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Parent Group Account Name Help"
      Top             =   4200
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1.5
      RowHeight       =   18
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
   Begin MSDataGridLib.DataGrid DGAcAlias 
      Height          =   3330
      Left            =   960
      Negotiate       =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Alias Group Account Name Help"
      Top             =   3945
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1.5
      RowHeight       =   18
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
   Begin MSDataGridLib.DataGrid DGAcName 
      Height          =   3330
      Left            =   390
      Negotiate       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Group Account Name Help"
      Top             =   3420
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1.5
      RowHeight       =   18
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
      Index           =   5
      Left            =   1890
      MaxLength       =   15
      TabIndex        =   5
      ToolTipText     =   "Group behavior"
      Top             =   1995
      Width           =   1740
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
      Index           =   4
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   4
      ToolTipText     =   "Parent Group Account Name"
      Top             =   1725
      Width           =   4980
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
      Index           =   3
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1455
      Width           =   4980
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
      Index           =   2
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   2
      ToolTipText     =   "Alias Group Account Name"
      Top             =   1185
      Width           =   4980
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
      Index           =   1
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   1
      ToolTipText     =   "Group Account Name (BiLangual)"
      Top             =   915
      Width           =   4980
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
      Index           =   0
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   0
      ToolTipText     =   "Group Account Name"
      Top             =   645
      Width           =   4980
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trading A/C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   6
      Left            =   750
      TabIndex        =   20
      Top             =   2250
      Width           =   1080
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   3
      Left            =   6990
      TabIndex        =   18
      Top             =   630
      Width           =   45
   End
   Begin VB.Label LblAliasBiLang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Hindi)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   750
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label LblNameBiLang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Hindi)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   750
      TabIndex        =   11
      Top             =   900
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alias Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   2
      Left            =   750
      TabIndex        =   10
      Top             =   1170
      Width           =   1050
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   5
      Left            =   750
      TabIndex        =   9
      Top             =   1980
      Width           =   600
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Under"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   4
      Left            =   750
      TabIndex        =   8
      Top             =   1710
      Width           =   555
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   0
      Left            =   750
      TabIndex        =   7
      Top             =   630
      Width           =   555
   End
End
Attribute VB_Name = "FaGrEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mSearchCode As Integer, Alias As String, BasicGroup As Byte, SysGroup As String
Dim OldGroupName As String, OldGroupUnderAc As String, OldGroupCode As String, OldGroupUnderAcCode As String
Dim xName  As ListItem, mListItem As ListItem, NewGroupUnderAc As String
Dim Master As ADODB.Recordset, RsAcName As ADODB.Recordset
Dim RsAcNameHelp As ADODB.Recordset, RsAcAlias As ADODB.Recordset, RsUnderAc As ADODB.Recordset
Private Const AcName As Byte = 0, AcNameBiLang As Byte = 1, AcAlias As Byte = 2, AcAliasBiLang As Byte = 3
Private Const UnderAc As Byte = 4, Nature As Byte = 5, TradingYN As Byte = 6
Private PubDatamanFa As New DMFa.ClsFa

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Dim Rst As ADODB.Recordset, SameName As Byte, SameName1 As Byte, GroupCode As String
If KeyCode = vbKeyEscape Then Grid_Hide: Exit Sub
Select Case Index
    Case AcName
        FaDGridTxtKeyDown_Mast DGAcName, Txt, Index, RsAcName, KeyCode, False, 1
    Case AcAlias
        FaDGridTxtKeyDown_Mast DGAcAlias, Txt, Index, RsAcAlias, KeyCode, False, 1
        If SysGroup = "Y" Then
            If KeyCode = vbKeyReturn Then
                If TopCtrl1.TopText2 = "Edit" Then     ' For Edit Mode
                    If UCase(Trim(Txt(AcAlias).Text)) = UCase(Trim(Txt(AcName).Text)) Then SameName = 1
                    Set Rst = G_FaCn.Execute("Select GroupHelp From AcGroup Where GroupHelp='" & FaFilterString(Txt(AcAlias).Text) & "' and GroupHelp<>'" & FaFilterString(Alias) & "'")
                    If Rst.RecordCount > 0 Then SameName1 = 1
                    If SameName = 1 Or SameName1 = 1 Then
                        MsgBox "Duplicate Alias not Allowed", vbInformation, "Validation"
                        Txt(AcAlias).SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
    Case UnderAc
        FaDGridTxtKeyDown DGUnderAc, Txt, Index, RsUnderAc, KeyCode, False, 1
        If BasicGroup = 0 Then
            If KeyCode = vbKeyReturn Then
                If RsUnderAc.RecordCount > 0 Or (RsUnderAc.EOF = False Or RsUnderAc.BOF = False) Or Txt(Index).Text <> "" Then
                    GroupCode = RsUnderAc!Code
                    Set Rst = G_FaCn.Execute("Select GroupCode,GroupName,MainGrCode,Nature,AliasYN From AcGroup Where GroupCode='" & GroupCode & "'")
                    If Rst.RecordCount > 0 Then
                        Txt(Nature) = IIf(IsNull(Rst!Nature), "", Rst!Nature)
                        While Not Rst.EOF
                            If Rst!AliasYN = "N" Then
                                Txt(UnderAc) = Trim(Rst!GroupName)
                                Txt(UnderAc).Tag = Rst!GroupCode
                                NewGroupUnderAc = Txt(UnderAc)
                            End If
                            Rst.MoveNext
                        Wend
                    End If
                End If
                Exit Sub
            End If
        End If
    Case Nature
        FaListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left + Txt(Index).width, 700, Txt(Index).width, 4000
End Select
Set Rst = Nothing
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
Select Case Index
    Case UnderAc
        If DGUnderAc.Visible = True Then FaDGridTxtKeyPress Txt, Index, RsUnderAc, KeyAscii, "Name"
End Select
Exit Sub
ELoop:          If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case AcName
        If DGAcName.Visible = True Then FaDGridTxtKeyUp_Mast Txt, Index, RsAcName, KeyCode, "Name"
    Case AcAlias
        If DGAcAlias.Visible = True Then FaDGridTxtKeyUp_Mast Txt, Index, RsAcAlias, KeyCode, "Name"
    Case Nature
        If FrmList.Visible = True Then FaListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
End Select
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    FaCtrl_validate Txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, SameName As Byte, SameName1 As Byte, GroupCode As String
On Error GoTo ELoop
SameName = 0
SameName1 = 0
Select Case Index
    Case AcName
        If Txt(Index).Text = "" Then Exit Sub
        If TopCtrl1.TopText2 = "Add" Then         ' For Add Mode
            Set Rst = G_FaCn.Execute("Select GroupHelp From AcGroup Where GroupHelp='" & FaFilterString(Txt(AcName).Text) & "'")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Account Group not Allowed", vbInformation, "Validation"
                Txt(AcName).SetFocus
                Cancel = True
                Exit Sub
            End If
            Set RsUnderAc = G_FaCn.Execute("Select GroupCode As Code,GroupName As Name,GroupNature,MainGrCode,CurrentBalance,SubLedYN,AliasYN,GroupHelp,Nature From AcGroup Where MainGrCode<>'999' Order by GroupName")
            Set DGUnderAc.DataSource = RsUnderAc
        ElseIf TopCtrl1.TopText2 = "Edit" Then      ' For Edit Mode
            Set Rst = G_FaCn.Execute("Select GroupHelp From AcGroup Where GroupHelp='" & FaFilterString(Txt(AcName).Text) & "' and GroupHelp<>'" & FaFilterString(OldGroupName) & "'")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Account Group not Allowed", vbInformation, "Validation"
                Txt(AcName).SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
    Case AcAlias
        If Txt(Index).Text = "" Then Exit Sub
        If TopCtrl1.TopText2 = "Add" Then         ' For Add Mode
            If UCase(Trim(Txt(AcAlias).Text)) = UCase(Trim(Txt(AcName).Text)) Then SameName = 1
            Set Rst = G_FaCn.Execute("Select GroupHelp From AcGroup Where GroupHelp='" & FaFilterString(Txt(AcAlias).Text) & "'")
            If Rst.RecordCount > 0 Then SameName1 = 1
            If SameName = 1 Or SameName1 = 1 Then
                MsgBox "Duplicate Alias not Allowed", vbInformation, "Validation"
                Txt(AcAlias).SetFocus
                Cancel = True
                Exit Sub
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then     ' For Edit Mode
            If UCase(Trim(Txt(AcAlias).Text)) = UCase(Trim(Txt(AcName).Text)) Then SameName = 1
            Set Rst = G_FaCn.Execute("Select GroupHelp From AcGroup Where GroupHelp='" & FaFilterString(Txt(AcAlias).Text) & "' and GroupHelp<>'" & FaFilterString(Alias) & "'")
            If Rst.RecordCount > 0 Then SameName1 = 1
            If SameName = 1 Or SameName1 = 1 Then
                MsgBox "Duplicate Alias not Allowed", vbInformation, "Validation"
                Txt(AcAlias).SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
    Case UnderAc
        If Txt(Index).Text = "" Then Exit Sub
        If RsUnderAc.RecordCount > 0 Or (RsUnderAc.EOF = False Or RsUnderAc.BOF = False) Or Txt(Index).Text <> "" Then
            GroupCode = RsUnderAc!Code
            Set Rst = G_FaCn.Execute("Select GroupCode,GroupName,MainGrCode,Nature,AliasYN,TRADINGYN,GroupNature From AcGroup Where GroupCode='" & GroupCode & "'")
            If Rst.RecordCount > 0 Then
                Txt(Nature) = IIf(IsNull(Rst!Nature), "", Rst!Nature)
                While Not Rst.EOF
                    If Rst!AliasYN = "N" Then
                        Txt(UnderAc) = Trim(Rst!GroupName)
                        Txt(UnderAc).Tag = Rst!GroupCode
                        NewGroupUnderAc = Txt(UnderAc)
                        If Rst!GroupNature = "E" Or Rst!GroupNature = "R" Then
                            Txt(TradingYN) = IIf(Rst!TradingYN = "Y", "Yes", "No")
                        Else
                            Txt(TradingYN) = ""
                        End If
                    End If
                    Rst.MoveNext
                Wend
            End If
        End If
    Case Nature
        Txt(Index).Text = ListView.SelectedItem.Text
    End Select
Set Rst = Nothing
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub SaveMsg(Index As Integer)
    Grid_Hide
    If FaIsValid(Txt(AcName), "Group Name") = False Then Exit Sub
    If FaIsValid(Txt(UnderAc), "Under Group") = False Then Exit Sub
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
        TopCtrl1_eSave
    Else
        Txt(Index).SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case KeyCode
    Case vbKeyReturn, vbKeyDown, vbKeyUp
        Select Case KeyCode
            Case vbKeyDown, vbKeyUp
                If DGAcName.Visible = True Or DGAcAlias.Visible = True Or DGUnderAc.Visible = True Or FrmList.Visible = True Or ListView.Visible = True Then Exit Sub
        End Select
        If TypeOf Me.ActiveControl Is TextBox Then Txt_Validate Me.ActiveControl.Index, False
        If PubDatamanFa.FaManageKeysControl(Me, KeyCode, Shift) = True Then SaveMsg Nature
        KeyCode = 0
    Case Else
        FaFormKeyDown Me, KeyCode, Shift
End Select
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyReturn
        KeyAscii = 0
End Select
End Sub
Private Sub Form_Load()
Dim I As Byte
    TopCtrl1.Tag = "AEDP": TopCtrl1.TopText1 = Me.CAPTION
    If PubSec = "SANJEEV" Then
        If rsUserPerm.RecordCount > 0 Then
            rsUserPerm.MoveFirst
            rsUserPerm.Find ("FORM_NAME='" & Me.CAPTION & "'")
            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
        End If
    ElseIf PubSec = "RAHUL" Then
        If rsUserPerm.RecordCount > 0 Then
            rsUserPerm.MoveFirst
            rsUserPerm.Find ("FORM_CODE='" & Me.Name & "'")
            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
        End If
    End If
    Me.left = 0
    Me.top = 0
    '''''''''''''
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
    '''''''''''''
    ConvBiLanguage BiLanguage
    FaFormIni Me, CtrlBColOrg, CtrlFColOrg
    Set RsAcName = New ADODB.Recordset
    RsAcName.CursorLocation = adUseClient
    RsAcName.Open "Select GroupCode As Code,GroupName As Name,GroupNature,MainGrCode,CurrentBalance,SubLedYN,AliasYN,GroupHelp,Nature From AcGroup Where MainGrCode<>'999' Order by GroupName", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGAcName.DataSource = RsAcName
    Set RsAcAlias = RsAcName
    Set DGAcAlias.DataSource = RsAcAlias
    Set RsUnderAc = RsAcName
    Set DGUnderAc.DataSource = RsUnderAc
    Set RsAcNameHelp = New ADODB.Recordset
    RsAcNameHelp.CursorLocation = adUseClient
    RsAcNameHelp.Open "Select ID,GroupCode,GroupName,GroupHelp,Nature From AcGroup Where MainGrCode<>'999' Order by GroupHelp", G_FaCn, adOpenDynamic, adLockOptimistic
    '* For Group Nature Filling
    With ListView.ListItems
        Set xName = .Add(, , "Bank")
        Set xName = .Add(, , "Broker")
        Set xName = .Add(, , "Cash")
        Set xName = .Add(, , "Customer")
        Set xName = .Add(, , "Electrician")
        Set xName = .Add(, , "Employee")
        Set xName = .Add(, , "Expenses")
        Set xName = .Add(, , "Mukadim")
        Set xName = .Add(, , "Others")
        Set xName = .Add(, , "PDC")
        Set xName = .Add(, , "Purchase")
        Set xName = .Add(, , "Revenue")
        Set xName = .Add(, , "Sale")
        Set xName = .Add(, , "SalesMan")
        Set xName = .Add(, , "SalesRep")
        Set xName = .Add(, , "Supplier")
        Set xName = .Add(, , "T.D.S.")
        Set xName = .Add(, , "Transporter")
        Set xName = .Add(, , "Unsecured Loan")
    End With
    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    If PubSiteCodeWiseMasterRst = True Then
        Set Master = G_FaCn.Execute("Select ID as SearchCode,ID,Site_Code,GroupCode,GroupName,GroupNameBiLang,GroupNature,MainGrCode,CurrentBalance,SubLedYN,BlOrd,AliasYN,GroupHelp,Nature,SysGroup,TRADINGYN From AcGroup Where SITE_CODE='" & PubSiteCode & "' AND AliasYN<>'Y' Order by GroupName")
    Else
        Set Master = G_FaCn.Execute("Select ID as SearchCode,ID,Site_Code,GroupCode,GroupName,GroupNameBiLang,GroupNature,MainGrCode,CurrentBalance,SubLedYN,BlOrd,AliasYN,GroupHelp,Nature,SysGroup,TRADINGYN From AcGroup Where AliasYN<>'Y' Order by GroupName")
    End If
    MoveRec
    Disp_Text SETS("INI", Me, Master)
End Sub
Private Sub Form_Resize()
    DGAcName.left = Txt(AcName).left
    DGAcName.top = Txt(AcName).top + Txt(AcName).height + 15
    DGAcAlias.left = Txt(AcAlias).left
    DGAcAlias.top = Txt(AcAlias).top + Txt(AcAlias).height + 15
    DGUnderAc.left = Txt(UnderAc).left
    DGUnderAc.top = Txt(UnderAc).top + Txt(UnderAc).height + 15
    FrmList.left = Txt(Nature).left
    FrmList.top = Txt(Nature).top + Txt(Nature).height + 15
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RsAcName = Nothing
    Set RsAcNameHelp = Nothing
    Set RsAcAlias = Nothing
    Set RsUnderAc = Nothing
    Set PubDatamanFa = Nothing
End Sub
Private Sub Grid_Hide()
    If DGAcName.Visible = True Then DGAcName.Visible = False
    If DGAcAlias.Visible = True Then DGAcAlias.Visible = False
    If DGUnderAc.Visible = True Then DGUnderAc.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub
Private Sub Disp_Text(Enb As Boolean)
    Txt(AcName).Enabled = Enb
    Txt(AcNameBiLang).Enabled = Enb
    Txt(AcAlias).Enabled = Enb
    Txt(AcAliasBiLang).Enabled = Enb
    Txt(UnderAc).Enabled = Enb
    Txt(Nature).Enabled = Enb
    Txt(TradingYN).Enabled = False
    
End Sub
Private Sub BlankText()
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).Text = ""
    Next I
    FaFormIni Me, CtrlBColOrg, CtrlFColOrg
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    Master.MoveFirst
    Master.Find ("SearchCode='" & MyValue & "'")
    MoveRec
    BUTTONS True, Me, Master, 0
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub MoveRec()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset
    If Master.RecordCount > 0 Then
        mSearchCode = Master!ID
        Lbl(3) = IIf(Master!SysGroup = "Y", "System Defined", "User Defined")
        Txt(AcName).Tag = Master!GroupCode
        Txt(AcName).Text = Master!GroupName
        OldGroupName = Txt(AcName).Text
        Txt(AcNameBiLang).Text = IIf(IsNull(Master!GroupNameBiLang), "", Master!GroupNameBiLang)
        Txt(Nature).Text = IIf(IsNull(Master!Nature), "", Master!Nature)
        If Master!GroupNature = "E" Or Master!GroupNature = "R" Then
            Txt(TradingYN) = IIf(Master!TradingYN = "Y", "Yes", "No")
        Else
            Txt(TradingYN) = ""
        End If
        SysGroup = Master!SysGroup
        '**** To Gather information of group
        Set Rst = G_FaCn.Execute("Select ID,GroupCode,GroupName,GroupNameBiLang,MainGrCode,Nature,SubLedYN,AliasYN,GroupHelp From AcGroup Where GroupCode='" & Txt(AcName).Tag & "'")
        If Rst.RecordCount > 0 Then
            If Len(Rst!MainGrCode) = 3 Then
                BasicGroup = 1
                Txt(UnderAc).Text = "Basic Group"
                Txt(UnderAc).Enabled = False
            Else
                BasicGroup = 0
                Txt(UnderAc).Text = G_FaCn.Execute("Select GroupName From AcGroup Where MainGrCode='" & left(Rst!MainGrCode, Len(Rst!MainGrCode) - 3) & "'").Fields(0).Value
                Txt(UnderAc).Tag = G_FaCn.Execute("Select GroupCode From AcGroup Where MainGrCode='" & left(Rst!MainGrCode, Len(Rst!MainGrCode) - 3) & "'").Fields(0).Value
            End If
            OldGroupUnderAc = Txt(UnderAc).Text
            OldGroupUnderAcCode = Txt(UnderAc).Tag
            While Not Rst.EOF
                If Rst.RecordCount > 1 And Rst!AliasYN = "Y" Then
                    Txt(AcAlias) = Rst!GroupName
                    Alias = Rst!GroupName
                    Txt(AcAliasBiLang) = IIf(IsNull(Rst!GroupNameBiLang), "", Rst!GroupNameBiLang)
                Else
                    Txt(AcAlias) = ""
                    Alias = ""
                    Txt(AcAliasBiLang) = ""
                End If
                Rst.MoveNext
            Wend
        End If
    Else
        BlankText
    End If
Set Rst = Nothing
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub ConvBiLanguage(Enb As Boolean)
    If Enb = True Then
        LblNameBiLang.left = Lbl(0).left
        LblNameBiLang.top = 900
        Txt(AcNameBiLang).left = Txt(AcName).left
        Txt(AcNameBiLang).top = 915
        LblAliasBiLang.left = Lbl(0).left
        LblAliasBiLang.top = 1440
        Txt(AcAliasBiLang).left = Txt(AcName).left
        Txt(AcAliasBiLang).top = 1455
        LblNameBiLang.CAPTION = "(" & BiLanguageName & ")"
        Txt(AcNameBiLang).Font = BiLanguageFont
        LblAliasBiLang.CAPTION = "(" & BiLanguageName & ")"
        Txt(AcAliasBiLang).Font = BiLanguageFont
        LblNameBiLang.Visible = True
        LblAliasBiLang.Visible = True
    Else
        LblNameBiLang.Visible = False
        Txt(AcNameBiLang).Visible = False
        LblAliasBiLang.Visible = False
        Txt(AcAliasBiLang).Visible = False
'        * Alias
        Lbl(2).left = Lbl(0).left
        Lbl(2).top = 900
        Txt(AcAlias).left = Txt(AcName).left
        Txt(AcAlias).top = 915
'        * Under
        Lbl(4).left = Lbl(0).left
        Lbl(4).top = 1170
        Txt(UnderAc).left = Txt(AcName).left
        Txt(UnderAc).top = 1185
'        * Nature
        Lbl(5).left = Lbl(0).left
        Lbl(5).top = 1440
        Txt(Nature).left = Txt(AcName).left
        Txt(Nature).top = 1455
'        * Sub Ledger
        Lbl(6).left = Lbl(0).left
        Lbl(6).top = 1710
        Txt(TradingYN).left = Txt(AcName).left
        Txt(TradingYN).top = 1725
        
    End If
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    Txt(AcName).Enabled = True
    Txt(UnderAc).Enabled = True
    SysGroup = "N"
    Lbl(3) = "User Defined"
    Txt(AcName).SetFocus
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    If Master.RecordCount > 0 Then
        OldGroupCode = Master!GroupCode
        OldGroupName = Master!GroupName
        Disp_Text SETS("EDIT", Me, Master)
        If SysGroup = "Y" Then
            Txt(AcName).Enabled = False
            Txt(UnderAc).Enabled = False
            If BiLanguage = True Then
                Txt(AcNameBiLang).SetFocus
            Else
                Txt(AcAlias).SetFocus
            End If
        Else
            Txt(AcName).Enabled = True
            Txt(UnderAc).Enabled = True
            Txt(AcName).SetFocus
        End If
    Else
        MsgBox "There Is No Record To Edit.", vbInformation, "Information"
    End If
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub ListView_Click()
On Error GoTo ELoop
    Txt(Val(ListView.Tag)).Text = ListView.SelectedItem.Text
    FrmList.Visible = False
    Txt(Val(ListView.Tag)).SetFocus
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eDel()
    UpdateDataBaseDelete
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "Select ID As SearchCode,GroupName,GroupNature,Nature FROM AcGroup Where AliasYN<>'Y' Order by GroupName"
    Set SearchForm = Me
    FAFind.Show vbModal
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_ePrn()
Dim RST1 As ADODB.Recordset, X11, I As Integer
On Error GoTo ERRORHANDLER
If PubBackEnd = "A" Then
    Set RST1 = G_FaCn.Execute("select SYSGROUP,GroupName,Nature,IIF(GroupNature='A','A S S E T S',IIF(GroupNature='E','E X P E N D I T U R E',IIF(GroupNature='L','L I A B I L I T Y',IIF(GroupNature ='R','R E V E N U E','Others')))) AS GNature,MainGrCode from acgroup order by groupname")
ElseIf PubBackEnd = "S" Then
    Set RST1 = G_FaCn.Execute("select SYSGROUP,GroupName,Nature,GNature= CASE GroupNature WHEN 'A' THEN 'A S S E T S' WHEN 'E' THEN 'E X P E N D I T U R E' WHEN 'L' THEN 'L I A B I L I T Y' WHEN 'R' THEN 'R E V E N U E' ELSE 'Others' END,MainGrCode from acgroup order by groupname")
End If
If RST1.RecordCount = 0 Then MsgBox "No record Found to Print": Exit Sub
X11 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaGrplist.ttx", True)
If MsgBox("Do You Want Tree Like List", vbQuestion + vbDefaultButton1 + vbYesNo, "A/C Group List") = vbYes Then
    Set rpt = PubDatamanFa.FaGrpListTreeRpt
Else
    Set rpt = PubDatamanFa.FaGrpListRpt
End If
For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("Title")
            rpt.FormulaFields(I).Text = "'Group List'"
    End Select
Next
rpt.Database.SetDataSource RST1
rpt.ReadRecords
FaReport_View rpt, 0, Me.CAPTION, True
Set RST1 = Nothing
Exit Sub
ERRORHANDLER:  MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub TopCtrl1_eRef()
    RsAcName.Requery
    RsAcNameHelp.Requery
    RsAcAlias.Requery
    RsUnderAc.Requery
'    Master.Requery
End Sub
Private Sub DGUnderAc_Click()
On Error GoTo ELoop
    DGUnderAc.Visible = False
    If RsUnderAc.RecordCount > 0 Then
        Txt(UnderAc).Text = RsUnderAc!Name
        Txt(UnderAc).Tag = RsUnderAc!Code
    End If
    Txt(UnderAc).SetFocus
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eSave()
    Grid_Hide
    If FaIsValid(Txt(AcName), "Group Name") = False Then Exit Sub
    If FaIsValid(Txt(UnderAc), "Under Group") = False Then Exit Sub
    If Txt(Nature).Enabled = True Then If FaIsValid(Txt(Nature), "Nature") = False Then Exit Sub
    If Txt(AcName) = Txt(UnderAc) Then MsgBox "A/c Group And Under group Can not be same": Txt(UnderAc).SetFocus: Exit Sub
    If TopCtrl1.TopText2 = "Add" Then
        If PubDatamanFa.FaGrAdd(Me) = True Then
            mSearchCode = Me.Tag 'ID
            Master.Requery
            RsAcName.Requery
            RsAcAlias.Requery
            RsAcNameHelp.Requery
            RsUnderAc.Requery
            Master.Find "SearchCode = " & mSearchCode
            TopCtrl1_eAdd
        End If
    Else
        UpdateDataBaseEdit
    End If
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To Txt.Count - 1
            Txt(I).BackColor = CtrlBColOrg
            Txt(I).ForeColor = CtrlFColOrg
        Next
    End If
Exit Sub
ELoop:   If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub Txt_GotFocus(Index As Integer)
FaCtrl_GetFocus Txt(Index)
Grid_Hide
Select Case Index
    Case AcName
        If RsAcName.RecordCount = 0 Or (RsAcName.EOF = True Or RsAcName.BOF = True) Or Txt(Index).Text = "" Then Exit Sub
        If Txt(Index).Text <> RsAcName!Name Then
            RsAcName.MoveFirst
            RsAcName.Find "Name='" & Txt(Index).Text & "'"
        End If
    Case AcAlias
        If RsAcAlias.RecordCount = 0 Or (RsAcAlias.EOF = True Or RsAcAlias.BOF = True) Or Txt(Index).Text = "" Then Exit Sub
        If Txt(Index).Text <> RsAcAlias!Name Then
            RsAcAlias.MoveFirst
            RsAcAlias.Find "Name ='" & Txt(Index).Text & "'"
        End If
    Case UnderAc
        If RsUnderAc.RecordCount = 0 Or (RsUnderAc.EOF = True Or RsUnderAc.BOF = True) Or Txt(Index).Text = "" Then Exit Sub
        If Txt(Index).Text <> RsUnderAc!Name Then
            RsUnderAc.MoveFirst
            RsUnderAc.Find "Name ='" & Txt(Index).Text & "'"
        End If
    Case Nature
        Set mListItem = ListView.FindItem(Txt(Index), 0, , 1)
        If mListItem Is Nothing Then
            Exit Sub
        Else
            mListItem.EnsureVisible
            mListItem.Selected = True
        End If
    End Select
End Sub
Private Sub UpdateDataBaseDelete()
On Error GoTo ELoop
Dim vBook As Variant, mTrans As Boolean
If SysGroup = "Y" Then MsgBox "System Group, Can not be Deleted", vbInformation, "Validation Check": Exit Sub
If G_FaCn.Execute("SELECT COUNT(*) from ACGROUP WHERE LEFT(MAINGRCODE,LEN(MAINGRCODE)-3) ='" & Master!MainGrCode & "'").Fields(0) > 0 Then MsgBox "Childs Exist Can't Delete it": Exit Sub
If G_FaCn.Execute("SELECT COUNT(*) from SUBGROUP WHERE GROUPCODE='" & Master!GroupCode & "'").Fields(0) > 0 Then MsgBox "Ledger A/c Exist Can't Delete it": Exit Sub
If Master.RecordCount > 0 Then
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        G_FaCn.BeginTrans
        mTrans = True
        G_FaCn.Execute ("Delete From AcGroup Where GroupCode='" & Master!GroupCode & "'")
        G_FaCn.CommitTrans
        mTrans = False
        Master.Requery
        RsAcName.Requery
        RsAcAlias.Requery
        RsAcNameHelp.Requery
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
        End If
        BUTTONS True, Me, Master, 0
        MoveRec
    End If
Else
    MsgBox "No Records To Delete!", vbInformation, "Information"
End If
Exit Sub
ELoop:
    If mTrans = True Then
        G_FaCn.RollbackTrans: If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
    Else
        If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
    End If
End Sub
Private Sub UpdateDataBaseEdit()
On Error GoTo ELoop
Dim RstOld As ADODB.Recordset, RstNew As ADODB.Recordset, mTrans As Boolean, ID As Integer, MyID As Integer
Dim NewGroupNature As String, NewUnderCode As String, NewMainGrCode As String, CurCount As Integer
Dim NewBlOrd As Integer, NewSysGroup As String, NewTradingYN As String, MainGrCode As String
Dim GroupCode As String, OldCurBal As Double, OldMainGrCode As String
Set RstOld = G_FaCn.Execute("Select * FROM ACGROUP WHERE GROUPCODE='" & OldGroupCode & "' AND GROUPNAME='" & OldGroupName & "'")
If RstOld.RecordCount > 0 Then
    OldCurBal = FaVNull(RstOld!CURRENTBALANCE)
    OldMainGrCode = RstOld!MainGrCode
    NewSysGroup = FaXNull(RstOld!SysGroup)
    ID = RstOld!ID
    MyID = RstOld!ID
    If Txt(UnderAc).Tag <> OldGroupUnderAcCode Then
        Set RstNew = G_FaCn.Execute("Select * FROM ACGROUP WHERE GROUPCODE='" & Txt(UnderAc).Tag & "' AND GROUPNAME='" & Txt(UnderAc) & "'")
        If RstNew.RecordCount > 0 Then
            NewGroupUnderAc = Txt(UnderAc)
            NewGroupNature = RstNew!GroupNature
            NewBlOrd = FaVNull(RstNew!BLORD)
            NewTradingYN = RstNew!TradingYN
            MainGrCode = RstNew!MainGrCode
            CurCount = 1
            Do While True
                NewMainGrCode = MainGrCode & Format(CurCount, "000")
                If G_FaCn.Execute("Select COUNT(*) from AcGroup Where MAINGRCODE='" & NewMainGrCode & "'").Fields(0).Value > 0 Then
                    CurCount = CurCount + 1
                Else
                    Exit Do
                End If
            Loop
        End If
    Else
        NewMainGrCode = RstOld!MainGrCode
        NewGroupUnderAc = Txt(UnderAc)
        NewGroupNature = RstOld!GroupNature
        NewBlOrd = FaVNull(RstOld!BLORD)
        NewTradingYN = FaXNull(RstOld!TradingYN)
    End If
End If
G_FaCn.BeginTrans
mTrans = True
Set RstOld = G_FaCn.Execute("SELECT * FROM ACGROUP WHERE LEFT(MAINGRCODE,LEN('" & OldMainGrCode & "'))='" & OldMainGrCode & "'")
If RstOld.RecordCount > 0 Then
    RstOld.MoveFirst
    Do Until RstOld.EOF
        G_FaCn.Execute ("UPDATE ACGROUP SET MAINGRCODE='" & NewMainGrCode & "'+" & IIf(PubBackEnd = "A", "MID", "SUBSTRING") & "(MAINGRCODE,LEN('" & OldMainGrCode & "')+1,255) WHERE MAINGRCODE='" & RstOld!MainGrCode & "'")
        RstOld.MoveNext
    Loop
End If
G_FaCn.Execute ("Update AcGroup Set GroupName='" & Txt(AcName) & "',GroupNameBiLang='" & Txt(AcNameBiLang) & "',GroupNature='" & NewGroupNature & "',MainGrCode='" & NewMainGrCode & "',Nature='" & Txt(Nature) & "',AliasYN='N',GroupHelp='" & FaFilterString(Txt(AcName)) & "',U_Name='" & pubUName & "',U_EntDt=" & FaConvertDate(PubLoginDate) & ",U_AE='E',BlOrd=" & NewBlOrd & ",SysGroup=" & FaChk_Text(NewSysGroup) & ",TradingYN=" & FaChk_Text(NewTradingYN) & "  Where ID=" & ID)

G_FaCn.Execute ("Update AcGroup Set GroupNature='" & NewGroupNature & "',Nature='" & Txt(Nature) & "',TradingYN=" & FaChk_Text(NewTradingYN) & " WHERE LEFT(MAINGRCODE,LEN('" & NewMainGrCode & "'))='" & NewMainGrCode & "'")
'G_FaCn.Execute ("Update SubGroup LEFT JOIN ACGROUP ON SUBGROUP.GROUPCODE=ACGROUP.GROUPCODE Set SubGroup.GroupNature='" & NewGroupNature & "',SubGroup.Nature='" & Txt(Nature) & "' WHERE LEFT(MAINGRCODE,LEN('" & NewMainGrCode & "'))='" & NewMainGrCode & "'")
'G_FaCn.Execute ("Update SubGroupAlias LEFT JOIN ACGROUP ON SubGroupAlias.GROUPCODE=ACGROUP.GROUPCODE Set SubGroupAlias.GroupNature='" & NewGroupNature & "',SubGroupAlias.Nature='" & Txt(Nature) & "' WHERE LEFT(MAINGRCODE,LEN('" & NewMainGrCode & "'))='" & NewMainGrCode & "'")
G_FaCn.Execute ("Update SubGroup Set SubGroup.GroupNature='" & NewGroupNature & "',SubGroup.Nature='" & Txt(Nature) & "' WHERE GROUPCODE='" & OldGroupCode & "'")
G_FaCn.Execute ("Update SubGroupAlias Set SubGroupAlias.GroupNature='" & NewGroupNature & "',SubGroupAlias.Nature='" & Txt(Nature) & "' WHERE GROUPCODE='" & OldGroupCode & "'")

Set RstOld = G_FaCn.Execute("Select ID From AcGroup Where GroupName='" & Alias & "'")
If RstOld.RecordCount > 0 Then
    ID = G_FaCn.Execute("Select ID From AcGroup Where GroupName='" & Alias & "'").Fields(0).Value
    G_FaCn.Execute ("Delete From AcGroup Where ID=" & ID)
End If
If Trim(Txt(AcAlias)) <> "" Then
    If PubBackEnd = "A" Then
        ID = G_FaCn.Execute("Select iif(isnull(max(ID)),0,max(ID)) From AcGroup").Fields(0).Value + 1
    ElseIf PubBackEnd = "S" Then
        ID = G_FaCn.Execute("Select isnull(max(ID),0) From AcGroup").Fields(0).Value + 1
    End If
    G_FaCn.Execute ("Insert Into AcGroup(ID,Site_Code,GroupCode,GroupName,GroupNameBiLang,GroupNature,MainGrCode,Nature,AliasYN,GroupHelp,U_Name,U_EntDt,U_AE,BlOrd,SysGroup,TradingYN) Values (" & ID & ",'" & PubSiteCode & "','" & Txt(AcName).Tag & "','" & Txt(AcAlias) & "','" & Txt(AcAliasBiLang) & "','" & _
    NewGroupNature & "','" & NewMainGrCode & "','" & Txt(Nature) & "','Y','" & FaFilterString(Txt(AcAlias)) & "','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'E'," & NewBlOrd & "," & FaChk_Text(NewSysGroup) & "," & FaChk_Text(NewTradingYN) & ")")
End If
G_FaCn.CommitTrans
mTrans = False
mSearchCode = MyID
Master.Requery
RsAcName.Requery
RsAcAlias.Requery
RsAcNameHelp.Requery
RsUnderAc.Requery
Master.Find "SearchCode = " & mSearchCode
Disp_Text SETS("INI", Me, Master)
MoveRec
Set RstOld = Nothing
Set RstNew = Nothing
Exit Sub
ELoop:
    If mTrans = True Then
        G_FaCn.RollbackTrans: If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
    Else
        If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
    End If
End Sub
