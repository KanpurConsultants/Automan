VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmGrEnt 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Group Entry"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   9675
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   2370
      Left            =   8205
      TabIndex        =   25
      Top             =   4170
      Visible         =   0   'False
      Width           =   2325
      Begin MSComctlLib.ListView ListView 
         Height          =   2340
         Left            =   45
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   4128
         View            =   3
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
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid DGUnderAc 
      Height          =   3330
      Left            =   2115
      Negotiate       =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4200
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
   Begin MSDataGridLib.DataGrid DGAcAlias 
      Height          =   3330
      Left            =   960
      Negotiate       =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3945
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
   Begin MSDataGridLib.DataGrid DGAcName 
      Height          =   3330
      Left            =   360
      Negotiate       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3375
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
      Height          =   225
      Index           =   6
      Left            =   5475
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2265
      Visible         =   0   'False
      Width           =   630
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
      Height          =   225
      Index           =   5
      Left            =   1890
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1920
      Width           =   1740
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
      Height          =   225
      Index           =   4
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1665
      Width           =   4980
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
      Height          =   225
      Index           =   3
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1410
      Width           =   4980
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
      Height          =   225
      Index           =   2
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1155
      Width           =   4980
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
      Height          =   225
      Index           =   1
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   2
      Top             =   900
      Width           =   4980
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
      Height          =   225
      Index           =   0
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   1
      Top             =   645
      Width           =   4980
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   661
   End
   Begin VB.Label LblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Edit Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   6945
      TabIndex        =   27
      Top             =   405
      Width           =   2700
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   5280
      TabIndex        =   21
      Top             =   2265
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label LblAliasBiLang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Hindi)"
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
      Left            =   450
      TabIndex        =   20
      Top             =   1410
      Visible         =   0   'False
      Width           =   570
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1710
      TabIndex        =   19
      Top             =   1410
      Width           =   75
   End
   Begin VB.Label LblNameBiLang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Hindi)"
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
      Left            =   450
      TabIndex        =   18
      Top             =   900
      Visible         =   0   'False
      Width           =   570
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   1710
      TabIndex        =   17
      Top             =   900
      Width           =   75
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   1710
      TabIndex        =   16
      Top             =   1920
      Width           =   75
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   1710
      TabIndex        =   15
      Top             =   1665
      Width           =   75
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1710
      TabIndex        =   14
      Top             =   1155
      Width           =   75
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1710
      TabIndex        =   13
      Top             =   645
      Width           =   75
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alias Name"
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
      Left            =   450
      TabIndex        =   12
      Top             =   1155
      Width           =   960
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Behaves Like a Sub Ledger (Yes->Y/No->N) ?"
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
      Left            =   450
      TabIndex        =   11
      Top             =   2265
      Visible         =   0   'False
      Width           =   4530
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature"
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
      Left            =   450
      TabIndex        =   10
      Top             =   1920
      Width           =   570
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Under"
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
      Left            =   450
      TabIndex        =   9
      Top             =   1665
      Width           =   510
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   450
      TabIndex        =   8
      Top             =   645
      Width           =   495
   End
End
Attribute VB_Name = "frmGrEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Don't Change Tag Property of (TxtUnder and Txt(AcName)) Control as it is used in other activities
Option Explicit
Public MasterFormExit As Boolean
Dim mSearchCode As Integer
Dim Alias As String                         ' For Alias
Dim BasicGroup As Byte                      ' For Basic Group Tracking(1- > Basic Group,0- > Non Basic Group)
Dim SysGroup As String                      ' For System Group Tracking(Y- > System Group,N- > Non System Group)
Dim OldGroupName As String
Dim OldGroupLevel As Integer
Dim OldGroupUnderAc As String, NewGroupUnderAc As String
Dim xName  As ListItem
Dim mListItem As ListItem
Dim Master As ADODB.Recordset
Dim RsAcName As ADODB.Recordset
Dim RsAcNameHelp As ADODB.Recordset
Dim RsAcAlias As ADODB.Recordset
Dim RsUnderAc As ADODB.Recordset

Private Const AcName As Byte = 0                ' Ac Name
Private Const AcNameBiLang As Byte = 1          ' Ac Name Bi Language
Private Const AcAlias As Byte = 2               ' Ac Alias
Private Const AcAliasBiLang As Byte = 3         ' Alias Bi Language
Private Const UnderAc As Byte = 4               ' Under Ac
Private Const Nature As Byte = 5                ' Nature
Private Const SubLedYN As Byte = 6              ' Sub Leder Ac Yes/No

Private Sub Disp_Text(Enb As Boolean)
    Txt(AcName).Enabled = Enb
    Txt(AcNameBiLang).Enabled = Enb
    Txt(AcAlias).Enabled = Enb
    Txt(AcAliasBiLang).Enabled = Enb
    Txt(UnderAc).Enabled = Enb
    Txt(Nature).Enabled = False
    Txt(SubLedYN).Enabled = Enb
End Sub
'To Make Controls Blank
Private Sub BlankText()
Dim i As Byte
    For i = 0 To Txt.Count - 1
        Txt(i).TEXT = ""
    Next i
End Sub

Private Sub Grid_Hide()
    If DGAcName.Visible = True Then DGAcName.Visible = False
    If DGAcAlias.Visible = True Then DGAcAlias.Visible = False
    If DGUnderAc.Visible = True Then DGUnderAc.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    Master.MoveFirst
    Master.FIND ("SearchCode='" & MyValue & "'")
    MoveRec
    BUTTONS True, Me, Master, 0
Exit Sub
ELoop:
    CheckError
End Sub
'Movement of Records
Private Sub MoveRec()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset, SiteType$
Dim mTAdd As Boolean, mTDel As Boolean
    If Master.RecordCount > 0 Then
        mSearchCode = Master!ID
        Txt(AcName).Tag = Master!GroupCode
        Txt(AcName).TEXT = Master!GroupName
        OldGroupName = Txt(AcName).TEXT
        Txt(AcNameBiLang).TEXT = IIf(IsNull(Master!GroupNameBiLang), "", Master!GroupNameBiLang)
        Txt(Nature).TEXT = IIf(IsNull(Master!Nature), "", Master!Nature)
        Txt(SubLedYN) = IIf(Master!SubLedYN = "Y", "Yes", "No")
        SysGroup = Master!SysGroup
        '**** To Gather information of group
        Set Rst = G_FaCn.Execute("Select ID,GroupCode,GroupName,GroupNameBiLang,MainGrCode,GroupLevel,Nature,SubLedYN,AliasYN,GroupHelp From AcGroup Where GroupCode='" & Txt(AcName).Tag & "'")
        If Rst.RecordCount > 0 Then
            If Rst!GroupLevel = 0 Then
                BasicGroup = 1
                Txt(UnderAc).TEXT = "Basic Group"
                Txt(UnderAc).Enabled = False
            Else
                BasicGroup = 0
                Txt(UnderAc).TEXT = G_FaCn.Execute("Select GroupName From AcGroup Where MainGrCode='" & left(Rst!MainGrCode, Rst!GroupLevel * 3) & "'").Fields(0).Value
                Txt(UnderAc).Tag = G_FaCn.Execute("Select GroupCode From AcGroup Where MainGrCode='" & left(Rst!MainGrCode, Rst!GroupLevel * 3) & "'").Fields(0).Value
            End If
            OldGroupLevel = Rst!GroupLevel
            OldGroupUnderAc = Txt(UnderAc).TEXT

            While Not Rst.EOF
                If Rst.RecordCount > 1 And Rst!AliasYN = "Y" Then
                    Txt(AcAlias) = Rst!GroupName: Alias = Rst!GroupName
                    Txt(AcAliasBiLang) = IIf(IsNull(Rst!GroupNameBiLang), "", Rst!GroupNameBiLang)
                Else
                    Txt(AcAlias) = "": Alias = ""
                    Txt(AcAliasBiLang) = ""
                End If
                Rst.MoveNext
            Wend
        End If
    Else
        BlankText
    End If
    Set Rst = Nothing
    SiteType = GCn.Execute("Select SiteType from Site where Site_Code='" & PubSiteCode & "'").Fields(0).Value
    If InStr(Me.TopCtrl1.Tag, "A") <> 0 Then mTAdd = True
    If InStr(Me.TopCtrl1.Tag, "D") <> 0 Then mTDel = True
    If SiteType = "H" Then  'HO
        TopCtrl1.tAdd = mTAdd
        TopCtrl1.tDel = mTAdd
        LblStatus = "Add/Delete Enabled"
    Else
        TopCtrl1.tAdd = False
        TopCtrl1.tDel = False
        LblStatus = "Add/Delete disabled"
    End If
    TopCtrl1.tFind = False

Exit Sub
ELoop:
    CheckError
End Sub
'replaced by lib function
'This Function is used to Maintain Current Balance of Group
'Private Sub CalBalance(MainGrCode As String, Amt As Double, PlusMinus As String)
'Dim ControlStr As String, i As Integer, Length As Integer
'    ControlStr = ""
'    Length = Len(MainGrCode) - 3
'    For i = Length To 3 Step -3
'        If ControlStr = "" Then ControlStr = "'" & left(MainGrCode, i) & "'" Else ControlStr = ControlStr & ",'" & left(MainGrCode, i) & "'"
'    Next
'    If ControlStr <> "" Then
'        G_FACN.Execute ("Update AcGroup Set CurrentBalance=CurrentBalance " & PlusMinus & " " & Amt & " Where MainGrCode In(" & ControlStr & ")")
'    End If
'End Sub

'This Function is used to change the position of control when Bi language is True or false
Private Sub ConvBiLanguage(Enb As Boolean)
    If Enb = True Then
        LblNameBiLang.left = Lbl(0).left: LblNameBiLang.top = 945
        Txt(AcNameBiLang).left = Txt(AcName).left: Txt(AcNameBiLang).top = 945
        LblAliasBiLang.left = Lbl(0).left: LblAliasBiLang.top = 1545
        Txt(AcAliasBiLang).left = Txt(AcName).left: Txt(AcAliasBiLang).top = 1545

        LblNameBiLang.CAPTION = "(" & BiLanguageName & ")"
        Txt(AcNameBiLang).Font = BiLanguageFont
        LblAliasBiLang.CAPTION = "(" & BiLanguageName & ")"
        Txt(AcAliasBiLang).Font = BiLanguageFont

        LblNameBiLang.Visible = True
        LblAliasBiLang.Visible = True
    Else
        LblNameBiLang.Visible = False: LblColon(1).Visible = False: Txt(AcNameBiLang).Visible = False
        LblAliasBiLang.Visible = False: LblColon(3).Visible = False: Txt(AcAliasBiLang).Visible = False
        '* Name
        Lbl(0).left = 450: Lbl(0).top = 660
        LblColon(0).left = 1710: LblColon(0).top = 660
        Txt(AcName).left = 1935: Txt(AcName).top = 660
        '* Alias
        Lbl(2).left = 450: Lbl(2).top = Lbl(0).top + Lbl(0).height + 30
        LblColon(2).left = 1710: LblColon(2).top = Lbl(2).top
        Txt(AcAlias).left = 1935: Txt(AcAlias).top = Lbl(2).top
        '* Under
        Lbl(4).left = 450: Lbl(4).top = Lbl(2).top + Lbl(2).height + 30
        LblColon(4).left = 1710: LblColon(4).top = Lbl(4).top
        Txt(UnderAc).left = 1935: Txt(UnderAc).top = Lbl(4).top
        '* Nature
        Lbl(5).left = 450: Lbl(5).top = 1800
        LblColon(5).left = 1710: LblColon(5).top = 1800
        Txt(Nature).left = 1935: Txt(Nature).top = 1800
        '* Sub Ledger
        Lbl(6).left = 450: Lbl(6).top = 2085
        LblColon(6).left = 5250: LblColon(6).top = Lbl(6).top
        Txt(SubLedYN).left = 5385: Txt(SubLedYN).top = 2085
    End If
End Sub
'Database updation procedure For Addition
Private Sub UpdateDataBaseAdd()
On Error GoTo ELoop
'* Variable Declaration
Dim ID As Integer, GroupCode As String * 4, mTrans As Boolean
Dim GroupNature As String * 1, MainGrCode As String
Dim Level As Byte, CurCount As Integer
'* Database Updation
    ID = G_FaCn.Execute("Select " & vIsNull("max(ID)", "0") & " From AcGroup").Fields(0).Value + 1
    GroupCode = Format(ID, "0000")
    GroupNature = G_FaCn.Execute("Select GroupNature From AcGroup Where GroupCode='" & Txt(UnderAc).Tag & "'").Fields(0).Value
    MainGrCode = G_FaCn.Execute("Select MainGrCode From AcGroup Where GroupCode='" & Txt(UnderAc).Tag & "'").Fields(0).Value & Format(G_FaCn.Execute("Select CurrentCount From AcGroup Where GroupCode='" & Txt(UnderAc).Tag & "'").Fields(0).Value, "000")
    Level = G_FaCn.Execute("Select GroupLevel From AcGroup Where GroupCode='" & Txt(UnderAc).Tag & "'").Fields(0).Value + 1
    CurCount = 1

    G_FaCn.BeginTrans
        mTrans = True
        G_FaCn.Execute ("Insert Into AcGroup(ID,Site_Code,GroupCode,GroupName,GroupNameBiLang," _
            & "GroupNature,MainGrCode,GroupLevel,CurrentCount,Nature," _
            & "SubLedYN,AliasYN,GroupHelp,U_Name,U_EntDt,U_AE) " _
            & "Values(" & ID & ",'" & PubSiteCode & "','" & GroupCode & "','" & Txt(AcName) & "','" & Txt(AcNameBiLang) & "'," _
            & "'" & GroupNature & "','" & MainGrCode & "'," & Level & "," & CurCount & ",'" & Txt(Nature) & "'," _
            & "'" & left(Txt(SubLedYN), 1) & "','N','" & FilterString(Txt(AcName)) & "','" & pubUName & "',#" & PubServerDate & "#,'A')")
        'Used For Alias of Group
        If Trim(Txt(AcAlias).TEXT) <> "" Then
            G_FaCn.Execute ("Insert Into AcGroup(ID,Site_Code,GroupCode,GroupName,GroupNameBiLang," _
                & "GroupNature,MainGrCode,GroupLevel,CurrentCount,Nature," _
                & "SubLedYN,AliasYN,GroupHelp,U_Name,U_EntDt,U_AE) " _
                & "Values(" & ID + 1 & ",'" & PubSiteCode & "','" & GroupCode & "','" & Txt(AcAlias) & "','" & Txt(AcAliasBiLang) & "'," _
                & "'" & GroupNature & "','" & MainGrCode & "'," & Level & "," & CurCount & ",'" & Txt(Nature) & "'," _
                & "'" & left(Txt(SubLedYN), 1) & "','Y','" & FilterString(Txt(AcAlias)) & "','" & pubUName & "',#" & PubServerDate & "#,'A')")
        End If
        If Txt(SubLedYN) = "Yes" Then
            '9999 To update subgroup table
'            G_FACN.Execute ("Insert Into SubGroup(ID,GroupCode,GroupName,GroupNature,MainGrCode,GroupLevel,CurrentCount,SubLedYN) " & _
                                 " Values(" & ID & ",'" & GroupCode & "'," & ChkText(Txt(AcName)) & ",'" & GroupNature & "','" & MainGrCode & "'," & Level & "," & CurCount & ",'" & Left(Txt(SubLedYN), 1) & "')")
        End If
        G_FaCn.Execute ("Update AcGroup Set CurrentCount=CurrentCount+1 Where GroupCode='" & Txt(UnderAc).Tag & "'")
    G_FaCn.CommitTrans
    mTrans = False
    mSearchCode = ID
    Master.Requery
    RsAcName.Requery
    RsAcAlias.Requery
    RsAcNameHelp.Requery

    Master.FIND "SearchCode = " & mSearchCode
    TopCtrl1_eAdd
Exit Sub
ELoop:
    If mTrans = True Then
        G_FaCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
End Sub
'Database updation procedure For Edit
Private Sub UpdateDataBaseEdit()
On Error GoTo ELoop
'* Variable Declaration
Dim Rst As ADODB.Recordset, mTrans As Boolean
Dim ID As Integer               ' For Assigning The Row ID of Record
Dim OldMainGrCode As String     ' For Assigning Old MainGrCode
Dim OldGroupLevel As Integer    ' For Assigning Old GroupLevel
Dim OldGroupNature As String * 1
Dim OldCurBal As Double
Dim OldNature As String * 10
Dim OldUnderCode As String
Dim NewUnderCode As String
Dim NewMainGrCode As String     ' For Assigning New MainGrCode
Dim NewGroupLevel As Integer    ' For Assigning New GroupLevel
Dim GroupLevelDiff As Integer   ' For Assigning Difference in GroupLevel(Upward(-)/Downward(+))
Dim NewGroupNature As String * 1
Dim NewNature As String * 10
Dim GroupCode As String * 4, CurCount As Integer
Dim MyID As Integer
    '* For Getting Old Values
    Set Rst = G_FaCn.Execute("Select ID,GroupCode,GroupNature,MainGrCode,GroupLevel,CurrentCount,CurrentBalance,Nature From AcGroup Where GroupHelp='" & FilterString(OldGroupName) & "'")
    ID = Rst!ID
    MyID = ID
    OldGroupNature = Rst!GroupNature
    OldMainGrCode = Rst!MainGrCode
    GroupCode = Rst!GroupCode
    OldGroupLevel = Rst!GroupLevel
    CurCount = Rst!CurrentCount
    OldCurBal = Rst!CURRENTBALANCE
    OldNature = IIf(IsNull(Rst!Nature), "", Rst!Nature)
    If BasicGroup = 0 Then
        OldUnderCode = G_FaCn.Execute("Select MainGrCode From AcGroup Where GroupName='" & OldGroupUnderAc & "'").Fields(0).Value
    End If
    '* For Getting New Values
    If BasicGroup = 0 Then
        Set Rst = G_FaCn.Execute("Select GroupNature,MainGrCode,GroupLevel,CurrentCount,CurrentBalance,Nature From AcGroup Where GroupCode='" & Txt(UnderAc).Tag & "'")
        NewGroupUnderAc = Txt(UnderAc)
        NewGroupNature = Rst!GroupNature
        OldNature = IIf(IsNull(Rst!Nature), "", Rst!Nature)
        NewMainGrCode = Rst!MainGrCode & Format(Rst!CurrentCount, "000")
        NewGroupLevel = Rst!GroupLevel + 1
        GroupLevelDiff = NewGroupLevel - OldGroupLevel
        NewUnderCode = G_FaCn.Execute("Select MainGrCode From AcGroup Where GroupName='" & NewGroupUnderAc & "'").Fields(0).Value
    End If

    '* Database Updation
    G_FaCn.BeginTrans
        mTrans = True
        If BasicGroup = 0 And Trim(OldGroupUnderAc) <> Trim(NewGroupUnderAc) Then      ' Under A/c is defferent than previously stored
            G_FaCn.Execute ("Update AcGroup Set " _
                & "GroupName='" & Txt(AcName) & "',GroupNameBiLang='" & Txt(AcNameBiLang) & "'," _
                & "GroupNature='" & NewGroupNature & "',MainGrCode='" & NewMainGrCode & "'," _
                & "GroupLevel=" & NewGroupLevel & ",Nature='" & Txt(Nature) & "'," _
                & "SubLedYN='" & left(Txt(SubLedYN), 1) & "',AliasYN='N'," _
                & "GroupHelp='" & FilterString(Txt(AcName)) & "',U_Name='" & pubUName & "'," _
                & "U_EntDt=#" & PubServerDate & "#,U_AE='E' " _
                & "Where ID=" & ID)
            '* For Reducing CurrentCount by 1 of the Old Under account
            G_FaCn.Execute ("Update AcGroup Set CurrentCount=CurrentCount-1 Where MainGrCode='" & OldUnderCode & "'")
            '* Update All GroupLevel=GroupLevel+1 of old group childs
            G_FaCn.Execute ("Update AcGroup Set GroupLevel=GroupLevel+" & GroupLevelDiff & " Where Left(MainGrCode,(" & OldGroupLevel & "+1)*3)='" & OldMainGrCode & "'")
            '* Replace All Maching Child's MainGrCode With New MainGrCode using the formula below
            '* Replace code with Newstring+substring(code,4,(len(code)-len(oldstring))
            G_FaCn.Execute ("Update AcGroup Set MainGrCode='" & NewMainGrCode & "'+Mid(MainGrCode,len((" & OldMainGrCode & "))+1,255-len((" & OldMainGrCode & "))) Where Left(MainGrCode,(" & OldGroupLevel & "+1)*3)='" & OldMainGrCode & "'")
            'Current balance Maintenance
            CalBalAcGroup "AcGroup", G_FaCn, OldMainGrCode, OldCurBal, "-"
            CalBalAcGroup "AcGroup", G_FaCn, NewMainGrCode, OldCurBal, "+"

            '* Used For Alias of Group
            '* If Previously alias exists and now it is blank
            If Alias <> "" And Trim(Txt(AcAlias).TEXT) = "" Then
                ID = G_FaCn.Execute("Select ID From AcGroup Where GroupName='" & Alias & "'").Fields(0).Value
                G_FaCn.Execute ("Delete * From AcGroup Where ID=" & ID)
            '* If Previously alias is Blank and now it Exists
            ElseIf Alias = "" And Trim(Txt(AcAlias).TEXT) <> "" Then
                ID = G_FaCn.Execute("Select " & vIsNull("max(ID)", "0") & " From AcGroup").Fields(0).Value + 1
                G_FaCn.Execute ("Insert Into AcGroup(ID,Site_Code,GroupCode,GroupName,GroupNameBiLang," _
                    & "GroupNature,MainGrCode,GroupLevel,CurrentCount,Nature," _
                    & "SubLedYN,AliasYN,GroupHelp,U_Name,U_EntDt,U_AE) " _
                    & "Values(" & ID & ",'" & PubSiteCode & "','" & GroupCode & "','" & Txt(AcAlias) & "','" & Txt(AcAliasBiLang) & "'," _
                    & "'" & NewGroupNature & "','" & NewMainGrCode & "'," & NewGroupLevel & "," & CurCount & ",'" & Txt(Nature) & "'," _
                    & "'" & left(Txt(SubLedYN), 1) & "','Y','" & FilterString(Txt(AcAlias)) & "','" & pubUName & "',#" & PubServerDate & "#,'E')")
            '* If Alias Changed
            ElseIf Alias <> Txt(AcAlias).TEXT Or Trim(Txt(AcAliasBiLang)) <> "" Then
                ID = G_FaCn.Execute("Select ID From AcGroup Where GroupName='" & Alias & "'").Fields(0).Value
                G_FaCn.Execute ("Update AcGroup Set " _
                    & "GroupName='" & Txt(AcAlias) & "',GroupNameBiLang='" & Txt(AcAliasBiLang) & "'," _
                    & "GroupNature='" & NewGroupNature & "',MainGrCode='" & NewMainGrCode & "'," _
                    & "GroupLevel=" & NewGroupLevel & ",CurrentCount=" & CurCount & "," _
                    & "Nature='" & Txt(Nature) & "',SubLedYN='" & left(Txt(SubLedYN), 1) & "'," _
                    & "GroupHelp='" & FilterString(Txt(AcAlias)) & "',U_Name='" & pubUName & "'," _
                    & "U_EntDt=#" & PubServerDate & "#,U_AE='E' " _
                    & "Where ID=" & ID)
            End If
            '* For Incrementing CurrentCount by 1 of the New Under account
            G_FaCn.Execute ("Update AcGroup Set CurrentCount=CurrentCount+1 Where MainGrCode='" & NewUnderCode & "'")
        Else        ' Under A/c is Same as previously stored
            G_FaCn.Execute ("Update AcGroup Set " _
                & "GroupName='" & Txt(AcName) & "',GroupNameBiLang='" & Txt(AcNameBiLang) & "'," _
                & "Nature='" & Txt(Nature) & "',SubLedYN='" & left(Txt(SubLedYN), 1) & "'," _
                & "AliasYN='N',GroupHelp='" & FilterString(Txt(AcName)) & "'," _
                & "U_Name='" & pubUName & "',U_EntDt=#" & PubServerDate & "#," _
                & "U_AE='E' Where ID=" & ID)
            '* Used For Alias of Group
            '* If Previously alias exists and now it is blank
            If Alias <> "" And Trim(Txt(AcAlias).TEXT) = "" Then
                ID = G_FaCn.Execute("Select ID From AcGroup Where GroupName='" & Alias & "'").Fields(0).Value
                G_FaCn.Execute ("Delete * From AcGroup Where ID=" & ID)
            '* If Previously alias is Blank and now it Exists
            ElseIf Alias = "" And Trim(Txt(AcAlias).TEXT) <> "" Then
                ID = G_FaCn.Execute("Select " & vIsNull("max(ID)", "0") & " From AcGroup").Fields(0).Value + 1
                G_FaCn.Execute ("Insert Into AcGroup(ID,Site_Code,GroupCode,GroupName,GroupNameBiLang," _
                    & "GroupNature,MainGrCode,GroupLevel,CurrentCount,Nature," _
                    & "SubLedYN,AliasYN,GroupHelp,U_Name,U_EntDt,U_AE) " _
                    & "Values(" & ID & ",'" & PubSiteCode & "','" & GroupCode & "','" & Txt(AcAlias) & "','" & Txt(AcAliasBiLang) & "'," _
                    & "'" & OldGroupNature & "','" & OldMainGrCode & "'," & OldGroupLevel & "," & CurCount & ",'" & Txt(Nature) & "'," _
                    & "'" & left(Txt(SubLedYN), 1) & "','Y','" & FilterString(Txt(AcAlias)) & "','" & pubUName & "',#" & PubServerDate & "#,'E')")
            '* If Alias Changed
            ElseIf Alias <> Txt(AcAlias).TEXT Or Trim(Txt(AcAliasBiLang)) <> "" Then
                ID = G_FaCn.Execute("Select ID From AcGroup Where GroupName='" & Alias & "'").Fields(0).Value
                G_FaCn.Execute ("Update AcGroup Set " _
                    & "GroupName='" & Txt(AcAlias) & "',GroupNameBiLang='" & Txt(AcAliasBiLang) & "'," _
                    & "GroupNature='" & OldGroupNature & "',MainGrCode='" & OldMainGrCode & "'," _
                    & "GroupLevel=" & OldGroupLevel & ",CurrentCount=" & CurCount & "," _
                    & "Nature='" & Txt(Nature) & "',SubLedYN='" & left(Txt(SubLedYN), 1) & "'," _
                    & "AliasYN='Y',GroupHelp='" & FilterString(Txt(AcAlias)) & "'," _
                    & "U_Name='" & pubUName & "',U_EntDt=#" & PubServerDate & "#,U_AE='E' " _
                    & "Where ID=" & ID)
            End If
        End If
    G_FaCn.CommitTrans
    mTrans = False
    mSearchCode = MyID
    Master.Requery
    RsAcName.Requery
    RsAcAlias.Requery
    RsAcNameHelp.Requery
    Master.FIND "SearchCode = " & mSearchCode
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Set Rst = Nothing
Exit Sub
ELoop:
    If mTrans = True Then
        G_FaCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
End Sub
'Database updation procedure For Delete
Private Sub UpdateDataBaseDelete()
On Error GoTo ELoop
'* Variable Declaration
Dim vBook As Variant, Rst As ADODB.Recordset, mTrans As Boolean
Dim MainGrCode As String
Dim CurBal As Double
Dim GroupUnderAc As String          ' For Assigning Under A/c Name
Dim GroupCodeUnderAc As String * 4  ' For Assigning Under A/c Code
Dim GroupCode As String * 4, CurCount As Integer
Dim MyID As Integer
    If SysGroup = "Y" Then
        MsgBox "System Group, Can not be Deleted", vbInformation, "Validation Check"
        Exit Sub
    End If
    '* For Getting Values
    Set Rst = G_FaCn.Execute("Select ID,GroupCode,MainGrCode,CurrentCount,CurrentBalance From AcGroup Where GroupHelp='" & FilterString(Txt(AcName)) & "'")
    GroupCode = Rst!GroupCode
    MainGrCode = Rst!MainGrCode
    CurBal = Rst!CURRENTBALANCE
    CurCount = Rst!CurrentCount
    GroupUnderAc = Txt(UnderAc)
    If BasicGroup = 1 Then
        MsgBox "Basic Group, Can not be Deleted", vbInformation, "Validation Check"
        Exit Sub
    End If
    If CurCount > 1 Then
        MsgBox "Child(s) Exist Under This Group," & vbCrLf & "Can not Delete this Record", vbInformation, "Validation Check"
        Exit Sub
    End If
    If G_FaCn.Execute("Select GroupCode From SubGroup Where GroupCode='" & GroupCode & "'").RecordCount > 0 Then
        MsgBox "Ledger A/c's exists, Delete Denied!", vbCritical, "Delete Denied"
        Exit Sub
    End If
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            vBook = Master.AbsolutePosition
            G_FaCn.BeginTrans
                mTrans = True
                G_FaCn.Execute ("Delete * From AcGroup Where GroupCode='" & GroupCode & "'")
                'Current balance Maintenance
                CalBalAcGroup "AcGroup", G_FaCn, MainGrCode, CurBal, "-"
                'For Reducing CurrentCount by 1 of the Under account
                GroupCodeUnderAc = G_FaCn.Execute("Select GroupCode From AcGroup Where GroupHelp='" & FilterString(GroupUnderAc) & "'").Fields(0).Value
                G_FaCn.Execute ("Update AcGroup Set CurrentCount=CurrentCount-1 Where GroupCode='" & GroupCodeUnderAc & "'")
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
Set Rst = Nothing
Exit Sub
ELoop:
    If mTrans = True Then
        G_FaCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
End Sub

Private Sub DGUnderAc_Click()
On Error GoTo ELoop
    DGUnderAc.Visible = False
    If RsUnderAc.RecordCount > 0 Then
        Txt(UnderAc).TEXT = RsUnderAc!Name
        Txt(UnderAc).Tag = RsUnderAc!Code
    End If
    Txt(UnderAc).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub ListView_Click()
On Error GoTo ELoop
    Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    Txt(Val(ListView.Tag)).SetFocus
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
    FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Load()
Dim i As Byte
On Error GoTo ELoop
    Me.height = 6930
    Me.left = 0
    Me.top = 0

    TopCtrl1.Tag = PubUParam
    ConvBiLanguage BiLanguage
    For i = 0 To Txt.Count - 1
        Txt(i).BackColor = CtrlBColOrg
        Txt(i).ForeColor = CtrlFColOrg
    Next

    Set RsAcName = New ADODB.Recordset
    RsAcName.CursorLocation = adUseClient
    RsAcName.Open "Select GroupCode As Code,GroupName As Name,GroupNature,MainGrCode,GroupLevel,CurrentCount,CurrentBalance,SubLedYN,AliasYN,GroupHelp,Nature From AcGroup Where MainGrCode<>'999' Order by GroupName", G_FaCn, adOpenDynamic, adLockOptimistic
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
        Set xName = .Add(, , "Cash")
        Set xName = .Add(, , "Bank")
        Set xName = .Add(, , "Customer")
        Set xName = .Add(, , "Supplier")
        Set xName = .Add(, , "Employee")
        Set xName = .Add(, , "Expenses")
        Set xName = .Add(, , "Revenue")
        Set xName = .Add(, , "Others")
    End With

    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    Set Master = G_FaCn.Execute("Select ID as SearchCode,ID,Site_Code,GroupCode,GroupName,GroupNameBiLang,GroupNature,MainGrCode,GroupLevel,CurrentCount,CurrentBalance,SubLedYN,BlOrd,AliasYN,GroupHelp,Nature,SysGroup From AcGroup Where MainGrCode<>'999' and AliasYN<>'Y' Order by GroupName")
'    If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Resize()
    DGAcName.left = Txt(AcName).left: DGAcName.top = Txt(AcName).top + Txt(AcName).height + 15
    DGAcAlias.left = Txt(AcAlias).left: DGAcAlias.top = Txt(AcAlias).top + Txt(AcAlias).height + 15
    DGUnderAc.left = Txt(UnderAc).left: DGUnderAc.top = Txt(UnderAc).top + Txt(UnderAc).height + 15
    FrmList.left = Txt(Nature).left: FrmList.top = Txt(Nature).top + Txt(Nature).height + 15
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set Master = Nothing: Set RsAcName = Nothing
    Set RsAcNameHelp = Nothing: Set RsAcAlias = Nothing
    Set RsUnderAc = Nothing
End Sub
'******* Top Bar
Public Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    Txt(SubLedYN).TEXT = "No"
    Txt(AcName).Enabled = True
    Txt(UnderAc).Enabled = True
    SysGroup = "N"
    Txt(AcName).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    If Master.RecordCount > 0 Then
        Disp_Text SETS("EDIT", Me, Master)
' Don't activate the commented section
'        If BasicGroup = 1 Then
'            Txt(UnderAc).Enabled = False
'            Txt(Nature).Enabled = True
'        Else
'            Txt(UnderAc).Enabled = True
'            Txt(Nature).Enabled = False
'        End If
        If SysGroup = "Y" Then
            Txt(AcName).Enabled = False
            Txt(UnderAc).Enabled = False
            Txt(Nature).Enabled = False
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
ELoop:
    CheckError
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
    GSQL = "Select ID As SearchCode,GroupName,GroupNature,Nature FROM AcGroup Where MainGrCode<>'999' And AliasYN<>'Y' Order by GroupName"
    Set SearchForm = Me
    FIND.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_ePrn()
Dim RST1 As ADODB.Recordset, X11, i As Integer
On Error GoTo ERRORHANDLER
Set RST1 = G_FaCn.Execute("select SYSGROUP,GroupName,Nature,MainGrCode," & cIIF("GroupNature='A'", "'Assets'", cIIF("GroupNature='E'", "'Expenditure'", cIIF("GroupNature='L'", "'Liability'", cIIF("GroupNature ='R'", "'Revenue'", "'** Not Defined**'")))) & " AS GNature from AcGroup order by MainGrCode,GroupName")
If RST1.RecordCount = 0 Then MsgBox "No record Found to Print": Exit Sub
X11 = CreateFieldDefFile(RST1, PubRepoPath + "\FaGrplist.ttx", True)
Set rpt = rdApp.OpenReport(PubRepoPath + "\FaGrpList.RPT")
For i = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
        Case UCase("Title")
            rpt.FormulaFields(i).TEXT = "'A/c Group List'"
    End Select
Next
rpt.Database.SetDataSource RST1
rpt.ReadRecords
Set RST1 = Nothing
Report_View rpt, "A/c Group List", 0, False
Exit Sub
ERRORHANDLER:  MsgBox err.Description, vbCritical, Me.CAPTION

End Sub

Private Sub TopCtrl1_eRef()
    RsAcName.Requery
    RsAcNameHelp.Requery
    RsAcAlias.Requery
    RsUnderAc.Requery
    Master.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Grid_Hide
    If IsValid(Txt(AcName), "Group Name") = False Then Exit Sub
    If IsValid(Txt(UnderAc), "Under Group") = False Then Exit Sub

    If TopCtrl1.TopText2 = "Add" Then
        UpdateDataBaseAdd
    Else
        UpdateDataBaseEdit
    End If
    If MasterFormExit Then Unload Me: Exit Sub
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim i As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
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
    Ctrl_GetFocus Txt(Index)
    Grid_Hide
    Select Case Index
    Case AcName
        If RsAcName.RecordCount = 0 Or (RsAcName.EOF = True Or RsAcName.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsAcName!Name Then
            RsAcName.MoveFirst
            RsAcName.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case AcAlias
        If RsAcAlias.RecordCount = 0 Or (RsAcAlias.EOF = True Or RsAcAlias.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsAcAlias!Name Then
            RsAcAlias.MoveFirst
            RsAcAlias.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case UnderAc
        If RsUnderAc.RecordCount = 0 Or (RsUnderAc.EOF = True Or RsUnderAc.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsUnderAc!Name Then
            RsUnderAc.MoveFirst
            RsUnderAc.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case Nature
        Set mListItem = ListView.FindItem(Txt(Index), 0, , 1)
        If mListItem Is Nothing Then
            Exit Sub
        Else
            mListItem.EnsureVisible
            mListItem.SELECTED = True
        End If
    End Select
End Sub

Private Sub SaveMsg(Index As Integer)
    Grid_Hide
    If IsValid(Txt(AcName), "Group Name") = False Then Exit Sub
    If IsValid(Txt(UnderAc), "Under Group") = False Then Exit Sub
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
        TopCtrl1_eSave
    Else
        Txt(Index).SetFocus
    End If
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Dim Rst As ADODB.Recordset
Dim SameName As Byte, SameName1 As Byte, GroupCode As String
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
    Case AcName
        DGridTxtKeyDown_Mast DGAcName, Txt, Index, RsAcName, KeyCode, False, 1
    Case AcAlias
        DGridTxtKeyDown_Mast DGAcAlias, Txt, Index, RsAcAlias, KeyCode, False, 1
        If SysGroup = "Y" Then
            If KeyCode = vbKeyReturn Then
                If TopCtrl1.TopText2 = "Edit" Then     ' For Edit Mode
                    If UCase(Trim(Txt(AcAlias).TEXT)) = UCase(Trim(Txt(AcName).TEXT)) Then SameName = 1
                    Set Rst = G_FaCn.Execute("Select GroupHelp From AcGroup Where GroupHelp='" & FilterString(Txt(AcAlias).TEXT) & "' and GroupHelp<>'" & FilterString(Alias) & "'")
                    If Rst.RecordCount > 0 Then SameName1 = 1
                    If SameName = 1 Or SameName1 = 1 Then
                        MsgBox "Duplicate Alias not Allowed", vbInformation, "Validation"
                        Txt(AcAlias).SetFocus
                        Exit Sub
                    End If
                    If BiLanguage = False Then SaveMsg Index
                End If
            End If
        End If
    Case AcAliasBiLang
        If SysGroup = "Y" Then
            If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
                SaveMsg Index
            End If
        End If
    Case UnderAc
        DGridTxtKeyDown DGUnderAc, Txt, Index, RsUnderAc, KeyCode, False, 1
        If BasicGroup = 0 Then
            If KeyCode = vbKeyReturn Then
                If RsUnderAc.RecordCount > 0 Or (RsUnderAc.EOF = False Or RsUnderAc.BOF = False) Or Txt(Index).TEXT <> "" Then
                    GroupCode = RsUnderAc!Code
                    Set Rst = G_FaCn.Execute("Select GroupCode,GroupName,MainGrCode,GroupLevel,Nature,AliasYN From AcGroup Where GroupCode='" & GroupCode & "'")
                    If Rst.RecordCount > 0 Then
                        Txt(Nature) = IIf(IsNull(Rst!Nature), "", Rst!Nature)
                        If TopCtrl1.TopText2 = "Add" Then
                            If Rst!GroupLevel > 50 Then
                                MsgBox "Maximum Level Exceed, Can not Create Account Under This Account", vbInformation, "Validation"
                                Txt(UnderAc).SetFocus
                                Exit Sub
                            End If
                        Else
                            If Rst!GroupLevel + OldGroupLevel > 50 Then
                                MsgBox "Maximum Level Exceed, Can not Create Account Under This Account", vbInformation, "Validation"
                                Txt(UnderAc).SetFocus
                                Exit Sub
                            End If
                        End If
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
                SaveMsg Index
                Exit Sub
            End If
        End If
    Case Nature
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 3000
        If BasicGroup = 1 Then
            If KeyCode = vbKeyReturn Then
                SaveMsg Index
            End If
        End If
    End Select
    If FrmList.Visible = False And DGAcName.Visible = False And DGAcAlias.Visible = False And DGUnderAc.Visible = False Then
        If SysGroup = "Y" Then
            If BiLanguage = True Then
                If Index <> AcAliasBiLang And (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
            ElseIf BiLanguage = False Then
                If Index <> AcAlias And (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
            End If
        Else
            If Index <> UnderAc And (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        End If
        If SysGroup = "Y" Then
            If BiLanguage = True Then
                If Index <> AcNameBiLang And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            Else
                If Index <> AcAlias And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        Else
            If Index <> AcName And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
    End If
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
    Select Case Index
    Case UnderAc
        If DGUnderAc.Visible = True Then DGridTxtKeyPress Txt, Index, RsUnderAc, KeyAscii, "Name"
    Case SubLedYN
        If KeyAscii = 89 Or KeyAscii = 121 Or KeyAscii = 78 Or KeyAscii = 110 Then
            If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                Txt(Index).TEXT = "Yes"
                KeyAscii = 0
            ElseIf KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                Txt(Index).TEXT = "No"
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
    Case AcName
        If DGAcName.Visible = True Then DGridTxtKeyUp_Mast Txt, Index, RsAcName, KeyCode, "Name"
    Case AcAlias
        If DGAcAlias.Visible = True Then DGridTxtKeyUp_Mast Txt, Index, RsAcAlias, KeyCode, "Name"
    Case Nature
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
Dim Rst As ADODB.Recordset
Dim SameName As Byte, SameName1 As Byte, GroupCode As String
On Error GoTo ELoop
    SameName = 0: SameName1 = 0
    Select Case Index
    Case AcName
        If Txt(Index).TEXT = "" Then Exit Sub
        If TopCtrl1.TopText2 = "Add" Then         ' For Add Mode
            Set Rst = G_FaCn.Execute("Select GroupHelp From AcGroup Where GroupHelp='" & FilterString(Txt(AcName).TEXT) & "'")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Account Group not Allowed", vbInformation, "Validation"
                Txt(AcName).SetFocus
                Cancel = True
                Exit Sub
            End If
            Set RsUnderAc = G_FaCn.Execute("Select GroupCode As Code,GroupName As Name,GroupNature,MainGrCode,GroupLevel,CurrentCount,CurrentBalance,SubLedYN,AliasYN,GroupHelp,Nature From AcGroup Where MainGrCode<>'999' Order by GroupName")
            Set DGUnderAc.DataSource = RsUnderAc
        ElseIf TopCtrl1.TopText2 = "Edit" Then      ' For Edit Mode
            Set Rst = G_FaCn.Execute("Select GroupHelp From AcGroup Where GroupHelp='" & FilterString(Txt(AcName).TEXT) & "' and GroupHelp<>'" & FilterString(OldGroupName) & "'")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Account Group not Allowed", vbInformation, "Validation"
                Txt(AcName).SetFocus
                Cancel = True
                Exit Sub
            End If
            '**** For Gather information of group
            Set Rst = G_FaCn.Execute("Select ID,GroupCode,GroupName,GroupNameBiLang,MainGrCode,GroupLevel,Nature,SubLedYN,AliasYN,GroupHelp From AcGroup Where GroupCode='" & Txt(AcName).Tag & "'")
            If Rst.RecordCount > 0 Then
                If Rst!GroupLevel <> 0 Then          ' For Non Primary Groups
                    '**** To Fill Corrusponding Groups of A/C Group(All Groups above the selected group except it and its childs)
                    Set RsUnderAc = G_FaCn.Execute("Select GroupCode As Code,GroupName As Name,GroupNature,MainGrCode,GroupLevel,CurrentCount,CurrentBalance,SubLedYN,AliasYN,GroupHelp,Nature From AcGroup WHERE Left(MainGrCode,(" & Rst!GroupLevel & "+1)*3)<>'" & Rst!MainGrCode & "' AND AcGroup.MainGrCode<>'" & Rst!MainGrCode & "' order by GroupName")
                    Set DGUnderAc.DataSource = RsUnderAc
                End If
            End If
        End If
    Case AcAlias
        If Txt(Index).TEXT = "" Then Exit Sub
        If TopCtrl1.TopText2 = "Add" Then         ' For Add Mode
            If UCase(Trim(Txt(AcAlias).TEXT)) = UCase(Trim(Txt(AcName).TEXT)) Then SameName = 1
            Set Rst = G_FaCn.Execute("Select GroupHelp From AcGroup Where GroupHelp='" & FilterString(Txt(AcAlias).TEXT) & "'")
            If Rst.RecordCount > 0 Then SameName1 = 1
            If SameName = 1 Or SameName1 = 1 Then
                MsgBox "Duplicate Alias not Allowed", vbInformation, "Validation"
                Txt(AcAlias).SetFocus
                Cancel = True
                Exit Sub
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then     ' For Edit Mode
            If UCase(Trim(Txt(AcAlias).TEXT)) = UCase(Trim(Txt(AcName).TEXT)) Then SameName = 1
            Set Rst = G_FaCn.Execute("Select GroupHelp From AcGroup Where GroupHelp='" & FilterString(Txt(AcAlias).TEXT) & "' and GroupHelp<>'" & FilterString(Alias) & "'")
            If Rst.RecordCount > 0 Then SameName1 = 1
            If SameName = 1 Or SameName1 = 1 Then
                MsgBox "Duplicate Alias not Allowed", vbInformation, "Validation"
                Txt(AcAlias).SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
    Case UnderAc
        If Txt(Index).TEXT = "" Then Exit Sub
        '**** For Alias Eqivalent Name Searching
        If RsUnderAc.RecordCount > 0 Or (RsUnderAc.EOF = False Or RsUnderAc.BOF = False) Or Txt(Index).TEXT <> "" Then
            GroupCode = RsUnderAc!Code
            Set Rst = G_FaCn.Execute("Select GroupCode,GroupName,MainGrCode,GroupLevel,Nature,AliasYN From AcGroup Where GroupCode='" & GroupCode & "'")
            If Rst.RecordCount > 0 Then
                Txt(Nature) = IIf(IsNull(Rst!Nature), "", Rst!Nature)
                If TopCtrl1.TopText2 = "Add" Then
                    If Rst!GroupLevel > 50 Then
                        MsgBox "Maximum Level Exceed, Can not Create Account Under This Account", vbInformation, "Validation"
                        Txt(UnderAc).SetFocus
                        Cancel = True
                        Exit Sub
                    End If
                Else
                    If Rst!GroupLevel + OldGroupLevel > 50 Then
                        MsgBox "Maximum Level Exceed, Can not Create Account Under This Account", vbInformation, "Validation"
                        Txt(UnderAc).SetFocus
                        Cancel = True
                        Exit Sub
                    End If
                End If
    
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
    Case Nature
        Txt(Index).TEXT = ListView.SelectedItem.TEXT
    End Select
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub
