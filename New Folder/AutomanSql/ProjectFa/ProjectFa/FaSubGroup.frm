VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "topctl.ocx"
Begin VB.Form FaSubGroup 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Ledger Accounts Entry"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   Icon            =   "FaSubGroup.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   10260
   Begin MSDataGridLib.DataGrid DGTDSCat 
      Height          =   2985
      Left            =   6015
      Negotiate       =   -1  'True
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   5265
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "T.D.S. Category"
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
            ColumnWidth     =   4124.977
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1785
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1170
      Width           =   4140
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   0
      TabIndex        =   108
      Top             =   0
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   661
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DFE7C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1770
      Index           =   5
      Left            =   3090
      TabIndex        =   99
      Top             =   4320
      Width           =   6180
      Begin VB.Frame Frame2 
         BackColor       =   &H00DFE7C0&
         Caption         =   "A/c Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   870
         Index           =   0
         Left            =   450
         TabIndex        =   100
         Top             =   255
         Width           =   5115
         Begin VB.OptionButton Opt3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DFE7C0&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   255
            Index           =   2
            Left            =   420
            TabIndex        =   102
            Top             =   165
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton Opt3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00DFE7C0&
            Caption         =   "Particular Group"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   285
            Index           =   3
            Left            =   2910
            TabIndex        =   101
            Top             =   150
            Width           =   1800
         End
         Begin MSDataListLib.DataCombo DCGroup 
            Height          =   315
            Left            =   270
            TabIndex        =   103
            Top             =   450
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   15527904
            ForeColor       =   12582912
            Text            =   ""
         End
      End
      Begin VB.OptionButton Opt6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CFE0E0&
         Caption         =   "Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   360
         Index           =   0
         Left            =   1680
         TabIndex        =   107
         Top             =   330
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.OptionButton Opt6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CFE0E0&
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Index           =   1
         Left            =   3330
         TabIndex        =   106
         Top             =   405
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton BtnPrint 
         BackColor       =   &H00D3BEC9&
         Caption         =   "&Print"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         MaskColor       =   &H00800080&
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Print Reports"
         Top             =   1185
         Width           =   2955
      End
      Begin VB.CommandButton BtnExit 
         BackColor       =   &H00D3BEC9&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3045
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Exit Form"
         Top             =   1185
         Width           =   3060
      End
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00B7DBC8&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   285
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2085
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   6270
      MaxLength       =   50
      TabIndex        =   98
      Top             =   1890
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox txtCurrBal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   9315
      MaxLength       =   15
      TabIndex        =   96
      Top             =   2130
      Width           =   1410
   End
   Begin MSDataGridLib.DataGrid DGArea 
      Height          =   2985
      Left            =   10650
      Negotiate       =   -1  'True
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   7605
      Visible         =   0   'False
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   5265
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "AreaName"
         Caption         =   "Area Name"
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
         DataField       =   "AreaCode"
         Caption         =   "Area Code"
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
            ColumnWidth     =   4124.977
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGUnderAc 
      Height          =   3330
      Left            =   9900
      Negotiate       =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   7665
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
      Left            =   8580
      Negotiate       =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   7485
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
      ColumnCount     =   5
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "NameHelp"
         Caption         =   "NameHelp"
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
         DataField       =   "GroupCode"
         Caption         =   "Under Group"
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
      EndProperty
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   44
      Left            =   9315
      TabIndex        =   93
      Top             =   2130
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   43
      Left            =   9315
      TabIndex        =   92
      Top             =   1853
      Width           =   1410
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   10440
      TabIndex        =   87
      Top             =   7050
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1800
         Left            =   45
         TabIndex        =   94
         Top             =   30
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   3175
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
   Begin MSDataGridLib.DataGrid DGPartyType 
      Height          =   3330
      Left            =   9180
      Negotiate       =   -1  'True
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   5445
      Visible         =   0   'False
      Width           =   4425
      _ExtentX        =   7805
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
            ColumnWidth     =   3720.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1785
      MaxLength       =   50
      TabIndex        =   4
      Top             =   900
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2355
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2820
      Width           =   3570
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   42
      Left            =   7680
      MaxLength       =   50
      TabIndex        =   43
      Top             =   5025
      Width           =   3945
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   41
      Left            =   7680
      MaxLength       =   50
      TabIndex        =   42
      Top             =   4755
      Width           =   3945
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   40
      Left            =   7680
      MaxLength       =   50
      TabIndex        =   41
      Top             =   4485
      Width           =   3945
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   39
      Left            =   8220
      MaxLength       =   50
      TabIndex        =   40
      Top             =   4215
      Width           =   3405
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   38
      Left            =   7680
      MaxLength       =   4
      TabIndex        =   39
      Text            =   "Mr."
      Top             =   4215
      Width           =   525
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   18
      Left            =   4620
      TabIndex        =   20
      Top             =   6270
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   19
      Left            =   4665
      TabIndex        =   21
      Top             =   5505
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   17
      Left            =   4665
      TabIndex        =   19
      Top             =   5250
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   1995
      MaxLength       =   35
      TabIndex        =   38
      Top             =   7935
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   36
      Left            =   4755
      MaxLength       =   6
      TabIndex        =   37
      Top             =   7665
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   34
      Left            =   1995
      MaxLength       =   50
      TabIndex        =   35
      Top             =   7380
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   1995
      MaxLength       =   50
      TabIndex        =   34
      Top             =   7110
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   1995
      MaxLength       =   50
      TabIndex        =   33
      Top             =   6840
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   35
      Left            =   1995
      TabIndex        =   36
      Top             =   7650
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   1995
      MaxLength       =   4
      TabIndex        =   31
      Text            =   "Mr."
      Top             =   6570
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   2565
      MaxLength       =   50
      TabIndex        =   32
      Top             =   6570
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1785
      MaxLength       =   4
      TabIndex        =   8
      Text            =   "Mr."
      Top             =   2820
      Width           =   525
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   7800
      MaxLength       =   20
      TabIndex        =   29
      Top             =   5865
      Visible         =   0   'False
      Width           =   2790
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   2985
      Left            =   7110
      Negotiate       =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   5265
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "City Name"
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
         Caption         =   "City Code"
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
            ColumnWidth     =   4124.977
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   23
      Left            =   4665
      MaxLength       =   3
      TabIndex        =   25
      Text            =   "Yes"
      Top             =   6015
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1785
      MaxLength       =   8
      TabIndex        =   0
      Top             =   360
      Width           =   1200
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   7605
      MaxLength       =   200
      TabIndex        =   30
      Top             =   3630
      Width           =   4065
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   45
      Left            =   6135
      MaxLength       =   50
      TabIndex        =   97
      Top             =   1455
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   1785
      TabIndex        =   13
      Top             =   3900
      Width           =   2565
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1785
      MaxLength       =   50
      TabIndex        =   1
      Top             =   630
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   1785
      MaxLength       =   50
      TabIndex        =   10
      Top             =   3090
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1785
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3360
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   1785
      MaxLength       =   50
      TabIndex        =   12
      Top             =   3630
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   1785
      MaxLength       =   25
      TabIndex        =   14
      Top             =   4170
      Width           =   2565
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   1785
      MaxLength       =   35
      TabIndex        =   15
      Top             =   4440
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   1785
      MaxLength       =   24
      TabIndex        =   16
      Top             =   4710
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   1785
      MaxLength       =   24
      TabIndex        =   17
      Top             =   4980
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   1785
      MaxLength       =   50
      TabIndex        =   18
      Top             =   5250
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   1785
      TabIndex        =   23
      Top             =   5790
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   1785
      TabIndex        =   22
      Top             =   5520
      Width           =   1290
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   22
      Left            =   4665
      MaxLength       =   3
      TabIndex        =   24
      Text            =   "Yes"
      Top             =   5760
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   7605
      MaxLength       =   20
      TabIndex        =   28
      Top             =   3360
      Width           =   2790
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   7605
      MaxLength       =   30
      TabIndex        =   27
      Top             =   3090
      Width           =   2790
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   7605
      MaxLength       =   30
      TabIndex        =   26
      Top             =   2820
      Width           =   2790
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7605
      MaxLength       =   50
      TabIndex        =   3
      Top             =   975
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      HelpContextID   =   2
      Index           =   2
      Left            =   7605
      MaxLength       =   50
      TabIndex        =   2
      Top             =   705
      Visible         =   0   'False
      Width           =   4140
   End
   Begin MSDataGridLib.DataGrid DGAcName 
      Height          =   3330
      Left            =   5970
      Negotiate       =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   8145
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
      ColumnCount     =   5
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "NameHelp"
         Caption         =   "NameHelp"
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
         DataField       =   "GroupCode"
         Caption         =   "Under Group"
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
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1155
      Left            =   150
      TabIndex        =   7
      Top             =   1545
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2037
      _Version        =   393216
      BackColor       =   14085097
      Cols            =   8
      BackColorFixed  =   128
      ForeColorFixed  =   65535
      BackColorSel    =   13297348
      BackColorBkg    =   13623520
      GridColor       =   128
      AllowUserResizing=   1
      BorderStyle     =   0
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.D.S.Category"
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
      Index           =   33
      Left            =   225
      TabIndex        =   109
      Top             =   1170
      Width           =   1380
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00800000&
      Height          =   1140
      Left            =   6075
      Shape           =   4  'Rounded Rectangle
      Top             =   4185
      Width           =   5640
   End
   Begin VB.Label LblHindi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hindi Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   6105
      TabIndex        =   90
      Top             =   3975
      Width           =   1230
   End
   Begin VB.Label LblAddBiLang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   89
      Top             =   4485
      Width           =   765
   End
   Begin VB.Label LblConPerBiLang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   88
      Top             =   4215
      Width           =   1365
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
      Height          =   240
      Index           =   35
      Left            =   7515
      TabIndex        =   86
      Top             =   1620
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      Height          =   885
      Left            =   7350
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   4170
   End
   Begin VB.Label LblOpBalType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dr/Cr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   10785
      TabIndex        =   85
      Top             =   1860
      Width           =   555
   End
   Begin VB.Label LblCurBalType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dr/Cr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   10785
      TabIndex        =   84
      Top             =   2130
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance"
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
      Index           =   32
      Left            =   7515
      TabIndex        =   83
      Top             =   1860
      Width           =   1560
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Balance"
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
      Index           =   31
      Left            =   7515
      TabIndex        =   82
      Top             =   2130
      Width           =   1425
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bussiness Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   29
      Left            =   435
      TabIndex        =   81
      Top             =   6195
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   1710
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   6525
      Visible         =   0   'False
      Width           =   5835
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Type"
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
      Height          =   255
      Index           =   30
      Left            =   3450
      TabIndex        =   80
      Top             =   6225
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local/Central"
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
      Height          =   255
      Index           =   28
      Left            =   3450
      TabIndex        =   79
      Top             =   5520
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
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
      Height          =   255
      Index           =   27
      Left            =   3450
      TabIndex        =   78
      Top             =   5265
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No(s)"
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
      Height          =   255
      Index           =   26
      Left            =   465
      TabIndex        =   77
      Top             =   7935
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Height          =   255
      Index           =   25
      Left            =   465
      TabIndex        =   76
      Top             =   6840
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Height          =   255
      Index           =   24
      Left            =   480
      TabIndex        =   75
      Top             =   7650
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
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
      Height          =   255
      Index           =   23
      Left            =   4200
      TabIndex        =   74
      Top             =   7665
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name"
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
      Index           =   22
      Left            =   465
      TabIndex        =   73
      Top             =   6570
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Cat."
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
      Index           =   21
      Left            =   6270
      TabIndex        =   72
      Top             =   6150
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IT Ward No."
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
      Index           =   20
      Left            =   6270
      TabIndex        =   71
      Top             =   5880
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt. Party"
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
      Height          =   210
      Index           =   19
      Left            =   3450
      TabIndex        =   66
      Top             =   6015
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Height          =   255
      Index           =   18
      Left            =   225
      TabIndex        =   65
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remark"
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
      Height          =   255
      Index           =   17
      Left            =   6075
      TabIndex        =   64
      Top             =   3645
      Width           =   720
   End
   Begin VB.Label LblNature 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   9315
      TabIndex        =   63
      Top             =   1620
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   180
      X2              =   11955
      Y1              =   2760
      Y2              =   2745
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Days"
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
      Height          =   255
      Index           =   16
      Left            =   225
      TabIndex        =   62
      Top             =   5820
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit"
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
      Height          =   255
      Index           =   15
      Left            =   225
      TabIndex        =   61
      Top             =   5535
      Width           =   975
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Active"
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
      Height          =   255
      Index           =   14
      Left            =   3450
      TabIndex        =   60
      Top             =   5775
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAN No."
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
      Height          =   255
      Index           =   13
      Left            =   6075
      TabIndex        =   59
      Top             =   3375
      Width           =   780
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LST No."
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
      Height          =   255
      Index           =   12
      Left            =   6075
      TabIndex        =   58
      Top             =   3090
      Width           =   735
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CST No."
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
      Height          =   255
      Index           =   11
      Left            =   6075
      TabIndex        =   57
      Top             =   2820
      Width           =   765
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
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
      Height          =   255
      Index           =   10
      Left            =   225
      TabIndex        =   56
      Top             =   5265
      Width           =   570
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Index           =   9
      Left            =   225
      TabIndex        =   55
      Top             =   4995
      Width           =   330
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
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
      Height          =   255
      Index           =   8
      Left            =   225
      TabIndex        =   54
      Top             =   4725
      Width           =   615
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
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
      Left            =   210
      TabIndex        =   53
      Top             =   4170
      Width           =   435
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Index           =   5
      Left            =   225
      TabIndex        =   52
      Top             =   3900
      Width           =   330
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Height          =   255
      Index           =   4
      Left            =   225
      TabIndex        =   51
      Top             =   3090
      Width           =   765
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
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
      Height          =   255
      Index           =   3
      Left            =   225
      TabIndex        =   50
      Top             =   2820
      Width           =   1365
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
      Height          =   255
      Left            =   6630
      TabIndex        =   49
      Top             =   975
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
      Height          =   255
      Left            =   6615
      TabIndex        =   48
      Top             =   705
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alias"
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
      Height          =   255
      Index           =   1
      Left            =   6150
      TabIndex        =   47
      Top             =   1215
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No(s)"
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
      Height          =   255
      Index           =   7
      Left            =   225
      TabIndex        =   46
      Top             =   4455
      Width           =   1125
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Under Group"
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
      Height          =   255
      Index           =   2
      Left            =   225
      TabIndex        =   45
      Top             =   900
      Width           =   1155
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
      Height          =   255
      Index           =   0
      Left            =   225
      TabIndex        =   44
      Top             =   630
      Width           =   555
   End
End
Attribute VB_Name = "FaSubGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CodeEditFlag As Boolean
Private Const mVType As String = "F_AO"
Private Const SubCode As Byte = 0, AcName As Byte = 1, AcNameBiLang As Byte = 2, TDSCat As Byte = 3
Private Const AcAliasBiLang As Byte = 4, UnderGroup As Byte = 5, ConPersonPrefix As Byte = 6
Private Const ConPerson As Byte = 7, Add1 As Byte = 8, Add2 As Byte = 9, Add3 As Byte = 10
Private Const City As Byte = 11, Area As Byte = 12, Phone As Byte = 13, Mobile As Byte = 14
Private Const Fax As Byte = 15, EMail As Byte = 16, Religion As Byte = 17, PartyType As Byte = 18
Private Const LC As Byte = 19, CrLimit As Byte = 20, CrDays As Byte = 21, ActiveYN As Byte = 22
Private Const GovtPartyYN As Byte = 23, CST As Byte = 24, LST As Byte = 25, PAN As Byte = 26
Private Const ITWardNo As Byte = 27, TDS_CAT As Byte = 28, Remark As Byte = 29, ConPersonPrefixB As Byte = 30
Private Const ConPersonB As Byte = 31, Add1B As Byte = 32, Add2B As Byte = 33, Add3B As Byte = 34
Private Const CityB As Byte = 35, PinB As Byte = 36, PhoneB As Byte = 37, ConPersonPrefixBiLang As Byte = 38
Private Const ConPersonBiLang As Byte = 39, Add1BiLang As Byte = 40, Add2BiLang As Byte = 41
Private Const Add3BiLang As Byte = 42, OpBal As Byte = 43, CurBal As Byte = 44, AcAlias = 45
'************* Constant defined for useing in place of Index in Line File
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_RefDate = 1                   ' Ref Date
Private Const Col_RefNo = 2                     ' Ref No
Private Const Col_Amount = 3                    ' Amount
Private Const Col_CrDr = 4                      ' Cr/Dr
Private Const Col_VNo = 5                       ' Voucher No
Dim TAddMode As Boolean, GridKey As Integer, ExitCtrl As Boolean, SysGroup As String
Dim VNo As Long, VPrefixUpdateFlag As Byte, mSearchCode As String, mDocId As String
Dim OldName As String, OldAlias As String, OldMainGrCode As String, OldCurBal As Double
Dim OldCurBalType As String, DetailFlag As Byte, AliasName As String, Master As ADODB.Recordset
Dim RsAcName As ADODB.Recordset, RsAcNameHelp As ADODB.Recordset, RsAcAlias As ADODB.Recordset
Dim RsUnderAc As ADODB.Recordset, Rscity As ADODB.Recordset, RsArea As ADODB.Recordset
Dim RsPartyType As ADODB.Recordset, RsTDSCat As ADODB.Recordset
Dim TmpSQL As String, ListArray As Variant, mListItem As ListItem
Private PubDatamanFa As New DMFa.ClsFa

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case KeyCode
    Case vbKeyReturn, vbKeyDown, vbKeyUp
        Select Case KeyCode
            Case vbKeyDown, vbKeyUp
                If DGAcName.Visible = True Or DGAcAlias.Visible = True Or DGUnderAc.Visible = True Or DGCity.Visible = True Or FrmList.Visible = True Or DGPartyType.Visible = True Or DGArea.Visible = True Then Exit Sub
        End Select
        If TypeOf Me.ActiveControl Is TextBox Then Txt_Validate Me.ActiveControl.Index, False
        If PubDatamanFa.FaManageKeysControl(Me, KeyCode, Shift) = True Then
            If UCase(Me.ActiveControl.Name) = "TXTGRID" Or UCase(Me.ActiveControl.Name) = "FGRID" Then
                If KeyCode = vbKeyDown Then SaveMsg Me.ActiveControl
            Else
                SaveMsg Me.ActiveControl
            End If
        End If
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
'On Error GoTo errorbox
Dim I As Byte
    Me.height = 7065
    Me.width = 11900
    Me.top = 0
    Me.left = 0
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
    DetailFlag = 1
    VPrefixUpdateFlag = 1
    ConvBiLanguage BiLanguage
    FaFormIni Me, CtrlBColOrg, CtrlFColOrg
    txtCurrBal.BackColor = CtrlBColOrg
    txtCurrBal.ForeColor = CtrlFColOrg
    CodeEditFlag = False
    If CodeEditFlag = True Then
        Lbl(18).Visible = True
        Txt(SubCode).Visible = True
    Else
        Lbl(18).Visible = False
        Txt(SubCode).Visible = False
    End If
    Set RsAcName = New ADODB.Recordset
    RsAcName.CursorLocation = adUseClient
    RsAcName.Open "Select SubCode As Code,Name,AliasYN,NameHelp,GroupCode From SubGroupAlias Order by Name", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGAcName.DataSource = RsAcName
    Set RsAcAlias = RsAcName
    Set DGAcAlias.DataSource = RsAcAlias
    Set RsUnderAc = New ADODB.Recordset
    RsUnderAc.CursorLocation = adUseClient
    RsUnderAc.Open "Select GroupCode As Code,GroupName As Name,GroupNature,MainGrCode,CurrentBalance,SubLedYN,AliasYN,GroupHelp,Nature From AcGroup Order by GroupName", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGUnderAc.DataSource = RsUnderAc
    Set Rscity = New ADODB.Recordset
    Rscity.CursorLocation = adUseClient
    Rscity.Open "Select CityCode As Code,CityName As Name,CityHelp From City Order by CityName", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGCity.DataSource = Rscity

    Set RsArea = New ADODB.Recordset
    RsArea.CursorLocation = adUseClient
    RsArea.Open "Select AreaCode,AreaName From areamast Order by areaName", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGArea.DataSource = RsArea
    
    Set RsTDSCat = New ADODB.Recordset
    RsTDSCat.CursorLocation = adUseClient
    RsTDSCat.Open "Select Code,Name From TDSCAT Order by Name", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGTDSCat.DataSource = RsTDSCat
    
    Set RsPartyType = New ADODB.Recordset
    RsPartyType.CursorLocation = adUseClient
    RsPartyType.Open "Select Party_Type As Code,Description As Name From SubGroupType Order by Description", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGPartyType.DataSource = RsPartyType
    
    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    If PubSiteCodeWiseMasterRst = True Then
        Set Master = G_FaCn.Execute("Select SubCode As SearchCode,SubCode,Site_Code,Name From SubGroup Where SITE_CODE='" & PubSiteCode & "' AND AliasYN<>'Y' order by Name")
    Else
        Set Master = G_FaCn.Execute("Select SubCode As SearchCode,SubCode,Site_Code,Name From SubGroup Where AliasYN<>'Y' order by Name")
    End If
    MoveRec
    Disp_Text SETS("INI", Me, Master)
    Frame2(5).Visible = False
    Me.TopCtrl1.TopText1.left = 5800
    Exit Sub
errorbox:       MsgBox err.Description, vbInformation
End Sub
Private Sub Form_Resize()
    Grid_Ini
    DGAcName.left = Txt(AcName).left
    DGAcName.top = Txt(AcName).top + Txt(AcName).height + 15
    DGAcAlias.left = Txt(AcAlias).left
    DGAcAlias.top = Txt(AcAlias).top + Txt(AcAlias).height + 15
    DGUnderAc.left = Txt(UnderGroup).left
    DGUnderAc.top = Txt(UnderGroup).top + Txt(UnderGroup).height + 15
    DGCity.left = Txt(City).left
    DGCity.top = Txt(City).top + Txt(City).height + 15
    DGArea.left = Txt(Area).left
    DGArea.top = Txt(Area).top + Txt(Area).height + 15
    DGPartyType.left = Txt(PartyType).left
    DGPartyType.top = Txt(PartyType).top + Txt(PartyType).height + 15
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RsAcName = Nothing
    Set RsAcAlias = Nothing
    Set RsAcNameHelp = Nothing
    Set RsUnderAc = Nothing
    Set Rscity = Nothing
    Set RsPartyType = Nothing
    Set RsTDSCat = Nothing
    Set PubDatamanFa = Nothing
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    For I = 0 To Txt.Count - 1
        If I = OpBal Or I = CurBal Then
        Else
            Txt(I).Enabled = Enb
        End If
    Next
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    Master.MoveFirst
    Master.Find ("SearchCode='" & MyValue & "'")
    MoveRec
    BUTTONS True, Me, Master, 0
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub BlankText()
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I) = ""
        Txt(I).Tag = ""
    Next
    LblNature.CAPTION = ""
    LblOpBalType.CAPTION = ""
    LblCurBalType.CAPTION = ""
    txtCurrBal = ""
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    FaFormIni Me, CtrlBColOrg, CtrlFColOrg
    txtCurrBal.BackColor = CtrlBColOrg
    txtCurrBal.ForeColor = CtrlFColOrg
End Sub
Private Sub Grid_Hide()
    If DGAcName.Visible = True Then DGAcName.Visible = False
    If DGAcAlias.Visible = True Then DGAcAlias.Visible = False
    If DGUnderAc.Visible = True Then DGUnderAc.Visible = False
    If DGCity.Visible = True Then DGCity.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGPartyType.Visible = True Then DGPartyType.Visible = False
    If DGArea.Visible = True Then DGArea.Visible = False
    If DGTDSCat.Visible = True Then DGTDSCat.Visible = False
End Sub
Private Sub ConvBiLanguage(Enb As Boolean)
    If Enb = True Then
        LblNameBiLang.CAPTION = "(" & BiLanguageName & ")"
        LblNameBiLang.Visible = True
        Txt(AcNameBiLang).Font = BiLanguageFont
        Txt(AcNameBiLang).Visible = True
        LblAliasBiLang.CAPTION = "(" & BiLanguageName & ")"
        LblAliasBiLang.Visible = True
        Txt(AcAliasBiLang).Font = BiLanguageFont
        Txt(AcAliasBiLang).Visible = True
        LblConPerBiLang.Visible = True
        Txt(ConPersonPrefixBiLang).Font = BiLanguageFont
        Txt(ConPersonPrefixBiLang).Visible = True
        Txt(ConPersonBiLang).Font = BiLanguageFont
        Txt(ConPersonBiLang).Visible = True
        LblAddBiLang.Visible = True
        Txt(Add1BiLang).Font = BiLanguageFont
        Txt(Add1BiLang).Visible = True
        Txt(Add2BiLang).Font = BiLanguageFont
        Txt(Add2BiLang).Visible = True
        Txt(Add3BiLang).Font = BiLanguageFont
        Txt(Add3BiLang).Visible = True
    Else
        Shape3.Visible = False
        LblHindi.Visible = False
        LblNameBiLang.Visible = False
        Txt(AcNameBiLang).Visible = False
        LblAliasBiLang.Visible = False
        Txt(AcAliasBiLang).Visible = False
        LblConPerBiLang.Visible = False
        Txt(ConPersonPrefixBiLang).Visible = False
        Txt(ConPersonBiLang).Visible = False
        LblAddBiLang.Visible = False
        Txt(Add1BiLang).Visible = False
        Txt(Add2BiLang).Visible = False
        Txt(Add3BiLang).Visible = False
    End If
    LblAliasBiLang.Visible = False
    Txt(AcAliasBiLang).Visible = False
End Sub
Private Sub Grid_Ini()
    With FGrid
        .left = 195
        .top = 1500
        .RowHeightMin = 220
        .Cols = 6
        .TextMatrix(0, Col_SrNo) = "S.No"
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 450
        .TextMatrix(0, Col_VNo) = "Voucher No"
        .ColAlignment(Col_VNo) = flexAlignLeftCenter
        .ColWidth(Col_VNo) = 0
        .TextMatrix(0, Col_RefDate) = "Ref. Date"
        .ColAlignment(Col_RefDate) = flexAlignLeftCenter
        .ColWidth(Col_RefDate) = 1200
        .TextMatrix(0, Col_RefNo) = "Narration"
        .ColAlignment(Col_RefNo) = flexAlignLeftCenter
        .ColWidth(Col_RefNo) = 1600
        .TextMatrix(0, Col_Amount) = "Amount"
        .ColAlignmentFixed(Col_Amount) = flexAlignRightCenter
        .ColWidth(Col_Amount) = 1800
        .TextMatrix(0, Col_CrDr) = "Cr/Dr"
        .ColAlignment(Col_CrDr) = flexAlignLeftCenter
        .ColWidth(Col_CrDr) = 600
    End With
End Sub
Private Sub CalOpBal()
Dim I As Integer, cr As Double, dr As Double, FinalCrDr As Double
    cr = 0
    dr = 0
    FinalCrDr = 0
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_CrDr) = "Cr" Then
            cr = cr + Val(FGrid.TextMatrix(I, Col_Amount))
        ElseIf FGrid.TextMatrix(I, Col_CrDr) = "Dr" Then
            dr = dr + Val(FGrid.TextMatrix(I, Col_Amount))
        End If
    Next
    FinalCrDr = Format(dr - cr, "0.00")
    Txt(OpBal) = Format(Abs(FinalCrDr), "0.00")
    If FinalCrDr < 0 Then
        LblOpBalType.CAPTION = "Cr"
    Else
        LblOpBalType.CAPTION = "Dr"
    End If
    Txt(CurBal) = Format(Abs(Val(Txt(CurBal).Tag) + FinalCrDr), "0.00")
    txtCurrBal = FaCurrBal(Txt(SubCode).TEXT)
    If Val(Txt(CurBal).Tag) + FinalCrDr < 0 Then
        LblCurBalType.CAPTION = "Cr"
    Else
        LblCurBalType.CAPTION = "Dr"
    End If
End Sub
Private Function GetReligion() As Byte
    If Txt(Religion) = "N/A" Then
        GetReligion = 0
    ElseIf Txt(Religion) = "Hindu" Then
        GetReligion = 1
    ElseIf Txt(Religion) = "Muslim" Then
        GetReligion = 2
    ElseIf Txt(Religion) = "Sikh" Then
        GetReligion = 3
    ElseIf Txt(Religion) = "Christian" Then
        GetReligion = 4
    End If
End Function
Private Function VoucherNo() As String
Dim Rst As ADODB.Recordset, VType As String, mVPrefix As String
    VType = "F_AO"
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & VType & "' Order By VP.Date_From DESC", G_FaCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount > 0 Then
        VNo = Rst!start_srl_no + 1
        mVPrefix = Rst!Prefix
    End If
    mDocId = PubDivCode + PubSiteCode + PubSiteCode + Space(5 - Len(CStr(VType))) + VType + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(VNo))) + CStr(VNo)
    VoucherNo = mDocId
Set Rst = Nothing
End Function
Private Sub Opt3_Click(Index As Integer)
Select Case Index
    Case 3
        Call FaIniCombo(" Select GroupCode,GroupName,maingrcode From AcGroup Where MainGrCode<>'999' Order by GroupName", DCGroup, "GroupName", "GroupCode")
End Select
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To Txt.Count - 1
            If I <> 3 Then
                Txt(I).BackColor = CtrlBColOrg
                Txt(I).ForeColor = CtrlFColOrg
            End If
        Next
    End If
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
End Sub
Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub
Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub
Private Sub DetailEnb(DetFlag As Byte)
Dim I As Byte
    If DetFlag = 1 Then
        For I = 6 To 42
            Txt(I).Enabled = True
        Next
    Else
        For I = 6 To 42
            Txt(I).Enabled = False
            Txt(I).TEXT = ""
        Next
    End If
End Sub
Private Sub DGArea_Click()
On Error GoTo ELoop
    DGArea.Visible = False
    If RsArea.RecordCount > 0 Then
        Txt(DGArea.Tag).TEXT = RsArea!Name
        Txt(DGArea.Tag).Tag = RsArea!Code
    End If
    Txt(DGArea.Tag).SetFocus
Exit Sub
ELoop:
    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub DGCity_Click()
On Error GoTo ELoop
    DGCity.Visible = False
    If Rscity.RecordCount > 0 Then
        Txt(DGCity.Tag).TEXT = Rscity!Name
        Txt(DGCity.Tag).Tag = Rscity!Code
    End If
    Txt(DGCity.Tag).SetFocus
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub DGPartyType_Click()
On Error GoTo ELoop
    DGPartyType.Visible = False
    If RsPartyType.RecordCount > 0 Then
        Txt(PartyType).TEXT = RsPartyType!Name
        Txt(PartyType).Tag = RsPartyType!Code
    End If
    Txt(PartyType).SetFocus
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub DGTDSCat_Click()
On Error GoTo ELoop
    DGTDSCat.Visible = False
    If RsTDSCat.RecordCount > 0 Then
        Txt(TDSCat).TEXT = RsTDSCat!Name
        Txt(TDSCat).Tag = RsTDSCat!Code
    End If
    Txt(TDSCat).SetFocus
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub DGUnderAc_Click()
On Error GoTo ELoop
    DGUnderAc.Visible = False
    If RsUnderAc.RecordCount > 0 Then
        Txt(UnderGroup).TEXT = RsUnderAc!Name
        Txt(UnderGroup).Tag = RsUnderAc!Code
    End If
    Txt(UnderGroup).SetFocus
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub ListView_Click()
On Error GoTo ELoop
    Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    Txt(Val(ListView.Tag)).SetFocus
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Function TxtGridLeave() As Boolean
Select Case FGrid.Col
    Case Col_RefDate
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = PubDatamanFa.FaRetDateFunc(TxtGrid(0))
    Case Col_RefNo, Col_CrDr
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
        If FGrid.Col = Col_CrDr Then
            CalOpBal
            If FGrid.TextMatrix(FGrid.Rows - 1, Col_CrDr) <> "" And Val(FGrid.TextMatrix(FGrid.Rows - 1, Col_Amount)) <> 0 Then FGrid.AddItem FGrid.Rows
        End If
    Case Col_Amount
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
        CalOpBal
        If FGrid.TextMatrix(FGrid.Rows - 1, Col_CrDr) <> "" And Val(FGrid.TextMatrix(FGrid.Rows - 1, Col_Amount)) <> 0 Then FGrid.AddItem FGrid.Rows
    End Select
    ExitCtrl = True
    TxtGridLeave = True
    TxtGrid(0).Visible = False
    FGrid.SetFocus
End Function
Private Sub MoveRec()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset, CrAmt As Double, DrAmt As Double, CrDr As String, CrDrAmt As Double
Dim I As Integer
If Master.RecordCount > 0 Then
    Txt(SubCode).TEXT = Master!SubCode
    Set Rst = G_FaCn.Execute("Select S.*,G.GroupNature AS GNature,G.GroupName,G.MainGrCode,a.areaname,C.CityName,C1.CityName As CityNameB,T.CATEGORY_Desc,ST.Description As PartyType,TDSCAT.NAME AS TDSNAME From ((((((SubGroup S Left Join AcGroup G on S.GroupCode=G.GroupCode) Left Join City C on S.CityCode=C.CityCode) Left Join City C1 on S.TCityCode=C1.CityCode) Left Join CATEGORY T on S.CatEGORY=T.CatEGORY) left join areamast a on s.areacode=a.areacode) Left Join SubGroupType ST on S.Party_Type=ST.Party_Type) " & _
    "LEFT JOIN TDSCAT ON TDSCAT.CODE=S.TDS_Catg Where S.SubCode='" & Master!SubCode & "'")
    If Rst.RecordCount > 0 Then
        If Master!Name = "Cash" Then
            SysGroup = "Y"
        Else
            SysGroup = "N"
        End If
        Txt(AcName) = Rst!Name
        OldName = Txt(AcName)
        Txt(TDSCat) = FaXNull(Rst!TDSNAME)
        Txt(AcNameBiLang) = IIf(IsNull(Rst!NameBiLang), "", Rst!NameBiLang)
        Txt(UnderGroup) = IIf(IsNull(Rst!GroupName), "", Rst!GroupName)
        Txt(UnderGroup).Tag = Rst!GroupCode
        OldMainGrCode = Rst!MainGrCode
        LblNature.CAPTION = Rst!Nature
        If Rst!GNature = "A" Or Rst!GNature = "L" Then
            DetailFlag = 1
        Else
            DetailFlag = 0
        End If
        If PubBackEnd = "A" Then
            CrAmt = G_FaCn.Execute("Select iif(isnull(Sum(AmtCr)),0,Sum(AmtCr)) From Ledger Where V_Type='" & mVType & "' and SubCode='" & Txt(SubCode) & "'").Fields(0).Value
            DrAmt = G_FaCn.Execute("Select iif(isnull(Sum(AmtDr)),0,Sum(AmtDr)) From Ledger Where V_Type='" & mVType & "' and SubCode='" & Txt(SubCode) & "'").Fields(0).Value
        ElseIf PubBackEnd = "S" Then
            CrAmt = G_FaCn.Execute("Select isnull(Sum(AmtCr),0) From Ledger Where V_Type='" & mVType & "' and SubCode='" & Txt(SubCode) & "'").Fields(0).Value
            DrAmt = G_FaCn.Execute("Select isnull(Sum(AmtDr),0) From Ledger Where V_Type='" & mVType & "' and SubCode='" & Txt(SubCode) & "'").Fields(0).Value
        End If
        Txt(OpBal).TEXT = Format(Abs(DrAmt - CrAmt), "0.00")
        LblOpBalType.CAPTION = IIf(CrAmt > DrAmt, "Cr", IIf(CrAmt < DrAmt, "Dr", ""))
        Txt(CurBal).TEXT = Format(IIf(IsNull(Rst!Curr_Bal), 0, Abs(Rst!Curr_Bal)), "0.00")
        txtCurrBal = FaCurrBal(Txt(SubCode).TEXT)
        LblCurBalType.CAPTION = IIf(Rst!Curr_Bal > 0, "Dr", IIf(Rst!Curr_Bal < 0, "Cr", ""))
        OldCurBal = Rst!Curr_Bal
        OldCurBalType = IIf(Rst!Curr_Bal > 0, "Dr", "Cr")
        Txt(CurBal).Tag = Rst!Curr_Bal + (CrAmt - DrAmt)
        Txt(ConPersonPrefix) = IIf(IsNull(Rst!ConPrefix), "", Rst!ConPrefix)
        Txt(ConPerson) = IIf(IsNull(Rst!ConPerson), "", Rst!ConPerson)
        Txt(Add1) = IIf(IsNull(Rst!Add1), "", Rst!Add1)
        Txt(Add2) = IIf(IsNull(Rst!Add2), "", Rst!Add2)
        Txt(Add3) = IIf(IsNull(Rst!Add3), "", Rst!Add3)
        Txt(City) = IIf(IsNull(Rst!CityName), "", Rst!CityName)
        Txt(City).Tag = IIf(IsNull(Rst!CityCode), "", Rst!CityCode)
        Txt(Area) = IIf(IsNull(Rst!areaname), 0, Rst!areaname)
        Txt(Area).Tag = IIf(IsNull(Rst!AreaCode), 0, Rst!AreaCode)
        Txt(Phone) = IIf(IsNull(Rst!Phone), "", Rst!Phone)
        Txt(Mobile) = IIf(IsNull(Rst!Mobile), "", Rst!Mobile)
        Txt(Fax) = IIf(IsNull(Rst!Fax), "", Rst!Fax)
        Txt(EMail) = IIf(IsNull(Rst!EMail), "", Rst!EMail)

        If Rst!Religion = 0 Or IsNull(Rst!Religion) Then
            Txt(Religion) = "N/A"
        ElseIf Rst!Religion = 1 Then
            Txt(Religion) = "Hindu"
        ElseIf Rst!Religion = 2 Then
            Txt(Religion) = "Muslim"
        ElseIf Rst!Religion = 3 Then
            Txt(Religion) = "Sikh"
        ElseIf Rst!Religion = 4 Then
            Txt(Religion) = "Christian"
        End If

        Txt(PartyType) = IIf(IsNull(Rst!PartyType), "", Rst!PartyType)
        Txt(PartyType).Tag = IIf(IsNull(Rst!Party_Type), 0, Rst!Party_Type)
        Txt(LC).TEXT = IIf(Rst!L_C = "L", "Local", "Central")
        Txt(CrLimit) = IIf(IsNull(Rst!CreditLimit), "", Rst!CreditLimit)
        Txt(CrDays) = IIf(IsNull(Rst!CreditDays), "", Rst!CreditDays)
        Txt(ActiveYN) = IIf(Rst!ActiveYN = 0, "No", "Yes")
        Txt(GovtPartyYN) = IIf(Rst!Govt_YN = 0, "No", "Yes")
        Txt(CST) = IIf(IsNull(Rst!CstNo), "", Rst!CstNo)
        Txt(LST) = IIf(IsNull(Rst!LstNo), "", Rst!LstNo)
        Txt(PAN) = IIf(IsNull(Rst!PANNo), "", Rst!PANNo)
        Txt(ITWardNo) = IIf(IsNull(Rst!ITWARD_NO), "", Rst!ITWARD_NO)
        Txt(Remark) = IIf(IsNull(Rst!Remark), "", Rst!Remark)
        Txt(ConPersonPrefixB) = IIf(IsNull(Rst!FPrefix), "", Rst!FPrefix)
        Txt(ConPersonB) = IIf(IsNull(Rst!FName), "", Rst!FName)
        Txt(Add1B) = IIf(IsNull(Rst!TAdd1), "", Rst!TAdd1)
        Txt(Add2B) = IIf(IsNull(Rst!TAdd2), "", Rst!TAdd2)
        Txt(Add3B) = IIf(IsNull(Rst!TAdd3), "", Rst!TAdd3)
        Txt(CityB) = IIf(IsNull(Rst!CityNameB), "", Rst!CityNameB)
        Txt(CityB).Tag = IIf(IsNull(Rst!TCityCode), "", Rst!TCityCode)
        Txt(PinB) = IIf(IsNull(Rst!TPin), "", Rst!TPin)
        Txt(PhoneB) = IIf(IsNull(Rst!TPhone), "", Rst!TPhone)
        Txt(ConPersonPrefixBiLang) = IIf(IsNull(Rst!ConPrefixBiLang), "", Rst!ConPrefixBiLang)
        Txt(ConPersonBiLang) = IIf(IsNull(Rst!ConPersonBiLang), "", Rst!ConPersonBiLang)
        Txt(Add1BiLang) = IIf(IsNull(Rst!Add1BiLang), "", Rst!Add1BiLang)
        Txt(Add2BiLang) = IIf(IsNull(Rst!Add2BiLang), "", Rst!Add2BiLang)
        Txt(Add3BiLang) = IIf(IsNull(Rst!Add3BiLang), "", Rst!Add3BiLang)
        ' For Alias
        Set Rst = G_FaCn.Execute("Select SubCode,Name,NameBiLang,AliasYN From SubGroupAlias Where SubCode='" & Txt(SubCode) & "' and AliasYN='Y'")
        If Rst.RecordCount > 0 Then
            Txt(AcAlias) = Rst!Name
            OldAlias = Rst!Name
            Txt(AcAliasBiLang) = IIf(IsNull(Rst!NameBiLang), "", Rst!NameBiLang)
        Else
            Txt(AcAlias) = ""
            OldAlias = ""
            Txt(AcAliasBiLang) = ""
        End If
        FGrid.Rows = 1
        Set Rst = G_FaCn.Execute("Select NARRATION,DocID,V_SNo,V_Type,V_No,V_Date,Site_Code,SubCode,AmtCr,AmtDr,Chq_No,Chq_Date From Ledger Where V_Type='" & mVType & "' and SubCode='" & Txt(SubCode) & "' Order by V_SNo")
        If Rst.RecordCount > 0 Then
            I = 1
            mDocId = Rst!DocID
            VNo = Rst!V_NO
            VPrefixUpdateFlag = 1
            Do Until Rst.EOF
                If Rst!AmtCr = 0 Then
                    CrDr = "Dr"
                    CrDrAmt = FaVNull(Rst!AmtDr)
                ElseIf Rst!AmtDr = 0 Then
                    CrDr = "Cr"
                    CrDrAmt = FaVNull(Rst!AmtCr)
                End If
                FGrid.AddItem I & Chr(9) & Format(Rst!V_Date, "dd/MMM/yyyy") & Chr(9) & FaXNull(Rst!Narration) & Chr(9) & Format(FaVNull(CrDrAmt), "0.00") & Chr(9) & CrDr & Chr(9) & FaVNull(Rst!V_NO)
                I = I + 1
                Rst.MoveNext
            Loop
            FGrid.FixedRows = 1
        Else
            mDocId = ""
            VNo = 0
            VPrefixUpdateFlag = 0
            FGrid.AddItem FGrid.Rows
            FGrid.FixedRows = 1
        End If
        Set Rst = Nothing
    End If
Else
    BlankText
End If
Set Rst = Nothing
Exit Sub
ELoop:          If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub SubGroupUpdate(ByRef xType As String, ByRef xTableName As String, ByRef xAcID As String, ByRef xSubCode As String, ByRef xSubName As String, ByRef xSubNameBiLang As String, ByRef xAliasYN As String, xA_E As String)
Dim RST1 As ADODB.Recordset, Nature As String, GroupNature As String, MyCurrBal As Double
Set RST1 = G_FaCn.Execute("SELECT * From AcGroup Where GroupCode='" & Txt(UnderGroup).Tag & "'")
If RST1.RecordCount > 0 Then
    Nature = FaXNull(RST1!Nature)
    GroupNature = FaXNull(RST1!GroupNature)
Else
    Nature = "Other"
    GroupNature = ""
End If
If xType = "Add" Then
    TmpSQL = "Insert Into " & xTableName & " (AcID,AcCode,Site_Code,SubCode,Name,NameBiLang,NameHelp,GroupCode,GroupNature,Nature,AliasYN,ConPrefix,ConPerson,Add1,Add2,Add3,CityCode,areacode,Phone,Mobile,Fax,EMail,Religion,Party_Type,L_C,CreditLimit,CreditDays,ActiveYN,Govt_YN,CSTNo,LSTNo,PANNo,ITWard_No,Remark,FPrefix,FName,TAdd1,TAdd2,TAdd3,TCityCode,TPIN,TPhone,ConPrefixBiLang,ConPersonBiLang,Add1BiLang,Add2BiLang,Add3BiLang,U_Name,U_EntDt,U_AE,TDS_Catg) Values ('" & _
    xAcID & "','" & xAcID & "','" & PubSiteCode & "','" & xSubCode & "','" & xSubName & "','" & xSubNameBiLang & "','" & FaFilterString(xSubName) & "','" & Txt(UnderGroup).Tag & "','" & GroupNature & "','" & Nature & "','" & xAliasYN & "','" & Txt(ConPersonPrefix) & "','" & Txt(ConPerson) & "','" & Txt(Add1) & "','" & Txt(Add2) & "','" & Txt(Add3) & "','" & Txt(City).Tag & "'," & Val(Txt(Area).Tag) & ",'" & Txt(Phone) & "','" & Txt(Mobile) & "','" & Txt(Fax) & "','" & _
    Txt(EMail) & "'," & "" & GetReligion & "," & Val(Txt(PartyType).Tag) & ",'" & IIf(Txt(LC) = "Local", "L", "C") & "'," & Val(Txt(CrLimit)) & "," & Val(Txt(CrDays)) & "," & IIf(Txt(ActiveYN) = "Yes", 1, 0) & "," & IIf(Txt(GovtPartyYN) = "Yes", 1, 0) & ",'" & Txt(CST) & "','" & Txt(LST) & "','" & Txt(PAN) & "','" & Txt(ITWardNo) & "','" & Txt(Remark) & "','" & Txt(ConPersonPrefixB) & "','" & Txt(ConPersonB) & "','" & Txt(Add1B) & "','" & Txt(Add2B) & "','" & _
    Txt(Add3B) & "','" & Txt(CityB).Tag & "','" & Txt(PinB) & "','" & Txt(PhoneB) & "','" & Txt(ConPersonPrefixBiLang) & "','" & Txt(ConPersonBiLang) & "','" & Txt(Add1BiLang) & "','" & Txt(Add2BiLang) & "','" & Txt(Add3BiLang) & "','" & pubUName & "'," & FaConvertDate(Now) & ",'" & xA_E & "','" & Txt(TDSCat).Tag & "')"
ElseIf xType = "Edit" Then
    TmpSQL = "Update " & xTableName & " Set Name='" & xSubName & "',NameBiLang='" & xSubNameBiLang & "',NameHelp='" & FaFilterString(xSubName) & "',GroupCode='" & Txt(UnderGroup).Tag & "',GroupNature='" & GroupNature & "',Nature='" & Nature & "',AliasYN='" & xAliasYN & "',ConPrefix='" & Txt(ConPersonPrefix) & "',ConPerson='" & Txt(ConPerson) & "',Add1='" & Txt(Add1) & "',Add2='" & Txt(Add2) & "',Add3='" & Txt(Add3) & "',CityCode='" & Txt(City).Tag & "',areacode=" & _
    Val(Txt(Area).Tag) & ",Phone='" & Txt(Phone) & "',Mobile='" & Txt(Mobile) & "',Fax='" & Txt(Fax) & "',EMail='" & Txt(EMail) & "',Religion=" & GetReligion & ",Party_Type=" & Val(Txt(PartyType).Tag) & ",L_C='" & IIf(Txt(LC) = "Local", "L", "C") & "',CreditLimit=" & Val(Txt(CrLimit)) & ",CreditDays=" & Val(Txt(CrDays)) & ",ActiveYN=" & IIf(Txt(ActiveYN) = "Yes", 1, 0) & ",Govt_YN=" & IIf(Txt(GovtPartyYN) = "Yes", 1, 0) & ",CSTNo='" & Txt(CST) & "',LSTNo='" & _
    Txt(LST) & "'," & "PANNo='" & Txt(PAN) & "',ITWard_No='" & Txt(ITWardNo) & "',Remark='" & Txt(Remark) & "',FPrefix='" & Txt(ConPersonPrefixB) & "',FName='" & Txt(ConPersonB) & "',TAdd1='" & Txt(Add1B) & "',TAdd2='" & Txt(Add2B) & "',TAdd3='" & Txt(Add3B) & "',TCityCode='" & Txt(CityB).Tag & "',TPIN='" & Txt(PinB) & "',TPhone='" & Txt(PhoneB) & "',ConPrefixBiLang='" & Txt(ConPersonPrefixBiLang) & "',ConPersonBiLang='" & Txt(ConPersonBiLang) & "',Add1BiLang='" & _
    Txt(Add1BiLang) & "',Add2BiLang='" & Txt(Add2BiLang) & "',Add3BiLang='" & Txt(Add3BiLang) & "',U_Name='" & pubUName & "',U_EntDt=" & FaConvertDate(Now) & ",U_AE='" & xA_E & "',TDS_Catg='" & Txt(TDSCat).Tag & "' Where AcID='" & xAcID & "'"
End If
End Sub
Private Sub UpdateDataBaseAdd()
On Error GoTo ELoop
Dim AcCodeAlias As String * 8, ID As Integer, mTrans As Boolean, Rst As ADODB.Recordset, I As Integer, AmtCrDr As String
Dim RST1 As ADODB.Recordset, VType As String, mVPrefix As String
If CodeEditFlag = True Then
    Set Rst = G_FaCn.Execute("Select SubCode From SubGroupAlias Where SubCode='" & Txt(SubCode) & "'")
    If Rst.RecordCount > 0 Then
        MsgBox "Code Already Exists", vbInformation, "Validation"
        Txt(SubCode) = PubSiteCode & Format(G_FaCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0).Value, "0000000")
        Txt(SubCode).Tag = Txt(SubCode)
        Txt(SubCode).SetFocus
        Exit Sub
    End If
End If
G_FaCn.BeginTrans
mTrans = True
ID = G_FaCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0).Value
SubGroupUpdate "Add", "SubGroup", Txt(SubCode), Txt(SubCode), Txt(AcName), Txt(AcNameBiLang), "N", "A"
G_FaCn.Execute (TmpSQL)
SubGroupUpdate "Add", "SubGroupAlias", Txt(SubCode), Txt(SubCode), Txt(AcName), Txt(AcNameBiLang), "N", "A"
G_FaCn.Execute (TmpSQL)
G_FaCn.Execute ("Update SubGroupCounter Set SubGroupAcCode=SubGroupAcCode+1")
If Trim(Txt(AcAlias)) <> "" Then
    ID = G_FaCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0).Value
    AcCodeAlias = PubSiteCode & Format(CStr(ID), "0000000")
    SubGroupUpdate "Add", "SubGroupAlias", AcCodeAlias, Txt(SubCode), Txt(AcAlias), Txt(AcAliasBiLang), "Y", "A"
    G_FaCn.Execute (TmpSQL)
    G_FaCn.Execute ("Update SubGroupCounter Set SubGroupAcCode=SubGroupAcCode+1")
End If
VType = "F_AO"
Set RST1 = New ADODB.Recordset
RST1.CursorLocation = adUseClient
RST1.Open "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & VType & "' Order By VP.Date_From DESC", G_FaCn, adOpenDynamic, adLockOptimistic
If RST1.RecordCount > 0 Then
    VNo = RST1!start_srl_no + 1
    mVPrefix = RST1!Prefix
End If
mDocId = PubDivCode + PubSiteCode + PubSiteCode + Space(5 - Len(CStr(VType))) + VType + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(VNo))) + CStr(VNo)
 mDocId = VoucherNo
For I = 1 To FGrid.Rows - 1
    If Val(FGrid.TextMatrix(I, Col_Amount)) <> 0 Then
        VPrefixUpdateFlag = 0
        If FGrid.TextMatrix(I, Col_CrDr) = "Cr" Then
            AmtCrDr = "AmtCr"
            FaCalCurrBal G_FaCn, Txt(SubCode), 0, Val(FGrid.TextMatrix(I, Col_Amount))
        Else
            AmtCrDr = "AmtDr"
            FaCalCurrBal G_FaCn, Txt(SubCode), Val(FGrid.TextMatrix(I, Col_Amount)), 0
        End If
        G_FaCn.Execute "Insert Into Ledger(DocId,V_SNo,V_Type,V_No,Site_Code,V_Date,SubCode," & AmtCrDr & ",NARRATION,U_Name,U_EntDt,U_AE,v_Prefix) Values ('" & mDocId & "'," & I & ",'" & mVType & "'," & VNo & ",'" & PubSiteCode & PubSiteCode & "'," & FaConvertDate(IIf(FGrid.TextMatrix(I, Col_RefDate) = "", PubStartDate - 1, FGrid.TextMatrix(I, Col_RefDate))) & ",'" & Txt(SubCode) & "'," & Val(FGrid.TextMatrix(I, Col_Amount)) & ",'" & FGrid.TextMatrix(I, Col_RefNo) & "','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'," & FaChk_Text(mVPrefix) & ")"
    End If
Next
' To Update Voucher_Prefix Serial No
If VPrefixUpdateFlag = 0 Then
    If FGrid.Rows >= 2 Then
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & mVType & "'", G_FaCn, adOpenDynamic, adLockOptimistic
        If Rst.RecordCount > 0 Then
            G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=Start_Srl_No+1 Where V_Type='" & Rst!V_tYPE & "'"
        End If
    End If
End If
G_FaCn.CommitTrans
RsAcName.Requery
RsAcAlias.Requery
mTrans = False
mSearchCode = Txt(SubCode)
Master.Requery
Master.Find "SearchCode = '" & mSearchCode & "'"
TopCtrl1_eAdd
Set Rst = Nothing
Set RST1 = Nothing
Exit Sub
ELoop:      If mTrans = True Then
                G_FaCn.RollbackTrans
                If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
            Else
                If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
            End If
End Sub
Private Sub UpdateDataBaseEdit()
'On Error GoTo ELoop
Dim Rst As ADODB.Recordset, mTrans As Boolean, I As Byte, ID As Integer, RST1 As ADODB.Recordset
Dim AcID As String, AcIDAlias As String, NewMainGrCode As String, AmtCrDr As String
NewMainGrCode = G_FaCn.Execute("Select MainGrCode From AcGroup Where GroupCode='" & Txt(UnderGroup).Tag & "'").Fields(0).Value
G_FaCn.BeginTrans
mTrans = True

Set RST1 = G_FaCn.Execute("SELECT * From Ledger Where DocId='" & mDocId & "' and V_Type='" & mVType & "' and SubCode= '" & Txt(SubCode) & "'")
Do Until RST1.EOF
    FaCalCurrBal G_FaCn, RST1!SubCode, RST1!AmtCr, RST1!AmtDr
    RST1.MoveNext
Loop
G_FaCn.Execute ("Delete From Ledger Where DocId='" & mDocId & "' and V_Type='" & mVType & "' and SubCode= '" & Txt(SubCode) & "'")
SubGroupUpdate "Edit", "SubGroup", Txt(SubCode), Txt(SubCode), Txt(AcName), Txt(AcNameBiLang), "N", "E"
G_FaCn.Execute (TmpSQL)
SubGroupUpdate "Edit", "SubGroupAlias", Txt(SubCode), Txt(SubCode), Txt(AcName), Txt(AcNameBiLang), "N", "E"
G_FaCn.Execute (TmpSQL)
If OldAlias <> "" And Trim(Txt(AcAlias)) = "" Then
    AcID = G_FaCn.Execute("Select AcID From SubGroupAlias Where Name='" & OldAlias & "'").Fields(0).Value
    G_FaCn.Execute ("Delete * From SubGroupAlias Where AcID='" & AcID & "'")
ElseIf OldAlias = "" And Trim(Txt(AcAlias)) <> "" Then
    ID = G_FaCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0).Value
    AcIDAlias = PubSiteCode & Format(CStr(ID), "0000000")
    SubGroupUpdate "Add", "SubGroupAlias", AcIDAlias, Txt(SubCode), Txt(AcAlias), Txt(AcAliasBiLang), "Y", "E"
    G_FaCn.Execute (TmpSQL)
    G_FaCn.Execute ("Update SubGroupCounter Set SubGroupAcCode=SubGroupAcCode+1")
ElseIf OldAlias <> "" Then
    AcID = G_FaCn.Execute("Select AcID From SubGroupAlias Where Name='" & OldAlias & "'").Fields(0).Value
    SubGroupUpdate "Edit", "SubGroupAlias", AcID, Txt(SubCode), Txt(AcAlias), Txt(AcAliasBiLang), "Y", "E"
    G_FaCn.Execute (TmpSQL)
End If
If mDocId = "" Then mDocId = VoucherNo
For I = 1 To FGrid.Rows - 1
    If Val(FGrid.TextMatrix(I, Col_Amount)) <> 0 Then
        If FGrid.TextMatrix(I, Col_CrDr) = "Cr" Then
            AmtCrDr = "AmtCr"
            FaCalCurrBal G_FaCn, Txt(SubCode), 0, Val(FGrid.TextMatrix(I, Col_Amount))
        Else
            AmtCrDr = "AmtDr"
            FaCalCurrBal G_FaCn, Txt(SubCode), Val(FGrid.TextMatrix(I, Col_Amount)), 0
        End If
        G_FaCn.Execute "Insert Into Ledger(DocId,V_SNo,V_Type,V_No,Site_Code,V_Date,SubCode," & AmtCrDr & ",Narration,U_Name,U_EntDt,U_AE) Values ('" & mDocId & "'," & I & ",'" & mVType & "'," & VNo & ",'" & PubSiteCode & PubSiteCode & "'," & FaConvertDate(IIf(FGrid.TextMatrix(I, Col_RefDate) = "", PubStartDate - 1, FGrid.TextMatrix(I, Col_RefDate))) & ",'" & Txt(SubCode) & "'," & Val(FGrid.TextMatrix(I, Col_Amount)) & ",'" & FGrid.TextMatrix(I, Col_RefNo) & "','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'E')"
    End If
Next
If VPrefixUpdateFlag = 0 Then
    If FGrid.Rows >= 2 Then
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & mVType & "'", G_FaCn, adOpenDynamic, adLockOptimistic
        If Rst.RecordCount > 0 Then
            G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=Start_Srl_No+1 Where V_Type='" & Rst!V_tYPE & "'"
        End If
    End If
End If
G_FaCn.CommitTrans
RsAcName.Requery
RsAcAlias.Requery
mTrans = False
mSearchCode = Txt(SubCode)
Master.Requery
Master.Find "SearchCode = '" & mSearchCode & "'"
Disp_Text SETS("INI", Me, Master)
MoveRec
Set Rst = Nothing
Exit Sub
ELoop:
    If mTrans = True Then
        G_FaCn.RollbackTrans: If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
    Else
        If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
    End If
End Sub
Private Sub UpdateDataBaseDelete()
On Error GoTo ELoop
Dim vBook As Variant, mTrans As Boolean, RST1 As ADODB.Recordset
    If SysGroup = "Y" Then
        MsgBox "You Can not Delete this A/c", vbInformation, "Validation Check"
        Exit Sub
    End If
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        G_FaCn.BeginTrans
        mTrans = True
        Set RST1 = G_FaCn.Execute("SELECT * From Ledger Where DocId='" & mDocId & "' and V_Type='" & mVType & "' and SubCode= '" & Txt(SubCode) & "'")
        Do Until RST1.EOF
            FaCalCurrBal G_FaCn, RST1!SubCode, RST1!AmtCr, RST1!AmtDr
            RST1.MoveNext
        Loop
        G_FaCn.Execute ("Delete From Ledger Where DocId='" & mDocId & "' and V_Type='" & mVType & "' and SubCode= '" & Txt(SubCode) & "'")
        G_FaCn.Execute ("Delete From SubGroupAlias Where AcCode='" & Txt(SubCode) & "'")
        G_FaCn.Execute ("Delete From SubGroup Where AcCode='" & Txt(SubCode) & "'")
        G_FaCn.CommitTrans
        mTrans = False
        TopCtrl1_eRef
        Master.Requery
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
        End If
        BUTTONS True, Me, Master, 0
        MoveRec
    End If
Exit Sub
ELoop:
    If mTrans = True Then
        G_FaCn.RollbackTrans: If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
    Else
        If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
    End If
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    DetailEnb 0
    Txt(ActiveYN).TEXT = "Yes"
    Txt(GovtPartyYN).TEXT = "No"
    Txt(LC).TEXT = "Local"
    Txt(Religion).TEXT = "N/A"
    Txt(SubCode) = PubSiteCode & Format(G_FaCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0).Value, "0000000")
    Txt(SubCode).Tag = Txt(SubCode).TEXT
    If CodeEditFlag = True Then
        Txt(SubCode).SetFocus
    Else
        Txt(AcName).SetFocus
    End If
    OldCurBal = 0
    OldCurBalType = ""
    SysGroup = "N"
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    If Master.RecordCount > 0 Then
        Disp_Text SETS("EDIT", Me, Master)
        Txt(SubCode).Enabled = False
        DetailEnb DetailFlag
        FGrid.AddItem FGrid.Rows
        If SysGroup = "Y" Then
            Txt(AcName).Enabled = False
            Txt(UnderGroup).Enabled = False
            Txt(AcAlias).SetFocus
        Else
            Txt(AcName).SetFocus
        End If
    Else
        MsgBox "There Is No Record To Edit.", vbInformation, "Information"
    End If
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eDel()
    If Master.RecordCount > 0 Then
        If G_FaCn.Execute("SELECT COUNT(*) From Ledger Where SubCode= '" & Txt(SubCode) & "' AND V_TYPE<>'F_AO'").Fields(0) > 0 Then
            MsgBox "Transactions Exist Can't Delete it", vbInformation, "Information"
            Exit Sub
        End If
        UpdateDataBaseDelete
        
        MoveRec
    Else
        MsgBox "There Is No Record To Delete", vbInformation, "Information"
    End If
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
GSQL = "Select SubCode As SearchCode,Name,G.GroupName As UnderGroup,ConPrefix,ConPerson,Add1 As Address1,Add2 As Address2,Add3 As Address3,C.CityName,a.areaname,Phone,Mobile,FAX,EMail,CSTNo As CST,LSTNo As LST,PANNo As PAN,Switch(ActiveYN=0,'No',ActiveYN=1,'Yes') As Active,Switch(Govt_YN=0,'No',Govt_YN=1,'Yes') As GovtParty,CreditLimit,CreditDays,Remark From (((SubGroup S Left Join AcGroup G on (S.GroupCode=G.GroupCode)) Left Join City C on (S.CityCode=C.CityCode)) " & _
"left join areamast a on (s.areacode=a.areacode)) Where S.AliasYN<>'Y' AND G.AliasYN<>'Y' Order by Name"
Set SearchForm = Me
FAFind.Show vbModal
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eRef()
    RsAcName.Requery
    'RsAcNameHelp.Requery
    RsAcAlias.Requery
    RsUnderAc.Requery
    Rscity.Requery
    RsPartyType.Requery
    RsTDSCat.Requery
'    Master.Requery
End Sub
Private Sub TopCtrl1_eSave()
Dim I As Integer, Rst As ADODB.Recordset
If TxtGrid(0).Visible = True Then
    If TxtGridLeave = False Then
        TxtGrid_LostFocus 0
        TxtGrid(0).SetFocus
        Exit Sub
    End If
End If
Grid_Hide
If FaIsValid(Txt(SubCode), "A/C Code") = False Then Exit Sub
If FaIsValid(Txt(AcName), "A/C Name") = False Then Exit Sub
If FaIsValid(Txt(UnderGroup), "Under Group") = False Then Exit Sub
If TopCtrl1.TopText2 = "Add" Then         ' For Add Mode
    Set Rst = G_FaCn.Execute("Select NameHelp From SubGroupAlias Where NameHelp='" & FaFilterString(Txt(AcName).TEXT) & "'")
    If Rst.RecordCount > 0 Then
        MsgBox "Duplicate A/c Name not Allowed", vbInformation, "Validation"
        Txt(AcName).SetFocus
        Exit Sub
    End If
ElseIf TopCtrl1.TopText2 = "Edit" Then      ' For Edit Mode
    Set Rst = G_FaCn.Execute("Select NameHelp From SubGroupAlias Where NameHelp='" & FaFilterString(Txt(AcName).TEXT) & "' and Name<>'" & OldName & "'")
    If Rst.RecordCount > 0 Then
        MsgBox "Duplicate A/c Name not Allowed", vbInformation, "Validation"
        Txt(AcName) = OldName
        Txt(AcName).SetFocus
        Exit Sub
    End If
End If
For I = 1 To FGrid.Rows - 1
    If Val(FGrid.TextMatrix(I, Col_Amount)) <> 0 Then
        If FGrid.TextMatrix(I, Col_CrDr) = "" Then MsgBox "Please Specify Cr/Dr in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_CrDr: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
    End If
Next
If TopCtrl1.TopText2 = "Add" Then
    UpdateDataBaseAdd
Else
    UpdateDataBaseEdit
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
FaCtrl_GetFocus Txt(Index)
Grid_Hide
Select Case Index
    Case AcName
        If RsAcName.RecordCount = 0 Or (RsAcName.EOF = True Or RsAcName.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsAcName!Name Then
            RsAcName.MoveFirst
            RsAcName.Find "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case AcAlias
        If RsAcAlias.RecordCount = 0 Or (RsAcAlias.EOF = True Or RsAcAlias.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsAcAlias!Name Then
            RsAcAlias.MoveFirst
            RsAcAlias.Find "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case UnderGroup
        If RsUnderAc.RecordCount = 0 Or (RsUnderAc.EOF = True Or RsUnderAc.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsUnderAc!Name Then
            RsUnderAc.MoveFirst
            RsUnderAc.Find "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case TDSCat
        If RsTDSCat.RecordCount = 0 Or (RsTDSCat.EOF = True Or RsTDSCat.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsTDSCat!Name Then
            RsTDSCat.MoveFirst
            RsTDSCat.Find "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case ConPersonPrefix
        ListArray = Array("Mr.", "Mrs.", "Miss", "M/S")
        Set mListItem = FaListView_Items(ListView, Txt, Index, ListArray, 4)
    Case ConPersonPrefixB
        ListArray = Array("S/O", "W/O", "D/O", "C/O", "And ", "U/C")
        Set mListItem = FaListView_Items(ListView, Txt, Index, ListArray, 6)
    Case City, CityB
        DGCity.Tag = Index
        DGCity.left = Txt(Index).left
        DGCity.top = Txt(Index).top + Txt(Index).height + 15
        If Rscity.RecordCount = 0 Or (Rscity.EOF = True Or Rscity.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> Rscity!Name Then
            Rscity.MoveFirst
            Rscity.Find "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case Area
        DGArea.Tag = Index
        DGArea.left = Txt(Index).left
        DGArea.top = Txt(Index).top + Txt(Index).height + 15
        If RsArea.RecordCount = 0 Or (RsArea.EOF = True Or RsArea.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsArea!areaname Then
            RsArea.MoveFirst
            RsArea.Find "AreaName ='" & Txt(Index).TEXT & "'"
        End If
    Case Religion
        ListArray = Array("N/A", "Hindu", "Muslim", "Sikh", "Christian")
        Set mListItem = FaListView_Items(ListView, Txt, Religion, ListArray, 5)
    Case PartyType
        If RsPartyType.RecordCount = 0 Or (RsPartyType.EOF = True Or RsPartyType.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsPartyType!Name Then
            RsPartyType.MoveFirst
            RsPartyType.Find "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case LC
        ListArray = Array("Local", "Central")
        Set mListItem = FaListView_Items(ListView, Txt, Index, ListArray, 2)
End Select
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case AcName
        FaDGridTxtKeyDown_Mast DGAcName, Txt, Index, RsAcName, KeyCode, False, 1
    Case AcAlias
        FaDGridTxtKeyDown_Mast DGAcAlias, Txt, Index, RsAcAlias, KeyCode, False, 1
    Case UnderGroup
        FaDGridTxtKeyDown DGUnderAc, Txt, Index, RsUnderAc, KeyCode, False, 1
    Case TDSCat
        FaDGridTxtKeyDown DGTDSCat, Txt, Index, RsTDSCat, KeyCode, False, 1
    Case ConPersonPrefix
        FaListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1200
    Case ConPersonPrefixB
        FaListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1800
    Case City, CityB
        FaDGridTxtKeyDown DGCity, Txt, Index, Rscity, KeyCode, False, 1
    Case Area
        FaDGridTxtKeyDown DGArea, Txt, Index, RsArea, KeyCode, False, 1
    Case Religion
        FaListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1500
    Case PartyType
        FaDGridTxtKeyDown DGPartyType, Txt, Index, RsPartyType, KeyCode, False, 1
    Case LC
        FaListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 600
    Case PhoneB
        If BiLanguage = False Then
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
                If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
                    TopCtrl1_eSave
                Else
                    Txt(Index).SetFocus
                    Exit Sub
                End If
            End If
        End If
    Case Add3BiLang
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
                TopCtrl1_eSave
            Else
                Txt(Index).SetFocus
            End If
        End If
End Select
'If FrmList.Visible = False And DGAcName.Visible = False And DGAcAlias.Visible = False And DGUnderAc.Visible = False And DGCity.Visible = False And DGCategory.Visible = False And DGPartyType.Visible = False And DGArea.Visible = False Then
'    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
'    If TopCtrl1.TopText2.CAPTION = "Add" Then
'        If CodeEditFlag = True Then
'            If Index <> SubCode And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
'        Else
'            If Index <> AcName And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
'        End If
'    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
'        If SysGroup = "Y" Then
'            If BiLanguage = True Then
'                If Index <> AcNameBiLang And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
'            Else
'                If Index <> AcAlias And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
'            End If
'        Else
'            If Index <> AcName And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
'        End If
'    End If
'End If
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
    Select Case Index
    Case UnderGroup
        If DGUnderAc.Visible = True Then FaDGridTxtKeyPress Txt, Index, RsUnderAc, KeyAscii, "Name"
    Case City, CityB
        If DGCity.Visible = True Then FaDGridTxtKeyPress Txt, Index, Rscity, KeyAscii, "Name"
    Case Area
        If DGArea.Visible = True Then FaDGridTxtKeyPress Txt, Index, RsArea, KeyAscii, "AreaName"
    Case TDSCat
        If DGTDSCat.Visible = True Then FaDGridTxtKeyPress Txt, Index, RsTDSCat, KeyAscii, "Name"
    Case PartyType
        If DGPartyType.Visible = True Then FaDGridTxtKeyPress Txt, Index, RsPartyType, KeyAscii, "Name"
    Case CrLimit
        FaNumPress Txt(CrLimit), KeyAscii, 9, 2
    Case CrDays
        FaNumPress Txt(CrDays), KeyAscii, 3, 0
    Case ActiveYN, GovtPartyYN
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
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case AcName
        If DGAcName.Visible = True Then FaDGridTxtKeyUp_Mast Txt, Index, RsAcName, KeyCode, "Name"
    Case AcAlias
        If DGAcAlias.Visible = True Then FaDGridTxtKeyUp_Mast Txt, Index, RsAcAlias, KeyCode, "Name"
    Case ConPersonPrefix, ConPersonPrefixB, Religion, LC
        If FrmList.Visible = True Then FaListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
End Select
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    FaCtrl_validate Txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, SameName As Byte, SameName1 As Byte
Select Case Index
    Case SubCode
        If Txt(SubCode).Visible = True Then
            If Txt(SubCode).TEXT = "" Then Txt(SubCode) = Txt(SubCode).Tag: Txt(SubCode).SetFocus: Exit Sub
            If G_FaCn.Execute("Select AcCode From SubGroup Where AcCode='" & Txt(SubCode) & "'").RecordCount > 0 Then
                MsgBox "A/c Code Already Exists", vbInformation, "Validation"
                Txt(SubCode) = Txt(SubCode).Tag
                Txt(SubCode).SetFocus
                Exit Sub
            End If
        End If
    Case AcName
        If Txt(Index).TEXT = "" Then Exit Sub
        If TopCtrl1.TopText2 = "Add" Then         ' For Add Mode
            Set Rst = G_FaCn.Execute("Select NameHelp From SubGroupAlias Where NameHelp='" & FaFilterString(Txt(AcName).TEXT) & "'")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate A/c Name not Allowed", vbInformation, "Validation"
                Txt(AcName).SetFocus
                Cancel = True
                Exit Sub
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then      ' For Edit Mode
            Set Rst = G_FaCn.Execute("Select NameHelp From SubGroupAlias Where NameHelp='" & FaFilterString(Txt(AcName).TEXT) & "' and Name<>'" & OldName & "'")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate A/c Name not Allowed", vbInformation, "Validation"
                Txt(AcName) = OldName
                Txt(AcName).SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
    Case AcAlias
        If Txt(Index).TEXT = "" Then Exit Sub
        If TopCtrl1.TopText2 = "Add" Then         ' For Add Mode
            If UCase(Trim(Txt(AcAlias).TEXT)) = UCase(Trim(Txt(AcName).TEXT)) Then SameName = 1
            Set Rst = G_FaCn.Execute("Select NameHelp From SubGroupAlias Where NameHelp='" & FaFilterString(Txt(AcAlias)) & "'")
            If Rst.RecordCount > 0 Then SameName1 = 1
            If SameName = 1 Or SameName1 = 1 Then
                MsgBox "Duplicate Alias not Allowed", vbInformation, "Validation"
                Txt(AcAlias).SetFocus
                Cancel = True
                Exit Sub
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then     ' For Edit Mode
            If UCase(Trim(Txt(AcAlias).TEXT)) = UCase(Trim(Txt(AcName).TEXT)) Then SameName = 1
            Set Rst = G_FaCn.Execute("Select NameHelp From SubGroupAlias Where NameHelp='" & FaFilterString(Txt(AcAlias)) & "' and NameHelp<>'" & FaFilterString(OldAlias) & "'")
            If Rst.RecordCount > 0 Then SameName1 = 1
            If SameName = 1 Or SameName1 = 1 Then
                MsgBox "Duplicate Alias not Allowed", vbInformation, "Validation"
                Txt(AcAlias).SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
    Case UnderGroup
        If Txt(Index).TEXT = "" Then Exit Sub
        If RsUnderAc.RecordCount > 0 Or (RsUnderAc.EOF = False Or RsUnderAc.BOF = False) Or Txt(Index).TEXT <> "" Then
            Set Rst = G_FaCn.Execute("Select ID,GroupCode,GroupName,Nature,AliasYN,GroupNature From AcGroup Where GroupCode='" & RsUnderAc!Code & "'")
            If Rst.RecordCount > 0 Then
                DetailFlag = 1
                DetailEnb DetailFlag
                LblNature = IIf(IsNull(Rst!Nature), "", Rst!Nature)
                If Rst!GroupNature = "A" Or Rst!GroupNature = "L" Then
                    DetailFlag = 1
                Else
                    DetailFlag = 0
                End If
                DetailEnb DetailFlag
                While Not Rst.EOF
                    If Rst!AliasYN = "N" Then
                        Txt(UnderGroup) = Trim(Rst!GroupName)
                        Txt(UnderGroup).Tag = Rst!GroupCode     'Rst!ID
                    End If
                    Rst.MoveNext
                Wend
            End If
        End If
    Case ConPersonPrefix, ConPersonPrefixB, Religion, LC
        If Txt(Index).TEXT <> "" Then Txt(Index).TEXT = ListView.SelectedItem.TEXT
    Case City, CityB
        If Rscity.RecordCount > 0 Or (Rscity.EOF = False Or Rscity.BOF = False) Then
            If Txt(Index).TEXT <> "" Then
                Txt(Index).TEXT = Rscity!Name
                Txt(Index).Tag = Rscity!Code
            Else
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            End If
        End If
     Case Area
            FaRstBofEof RsArea
            If Txt(Index).TEXT <> "" Then
                Txt(Index).TEXT = RsArea!areaname
                Txt(Index).Tag = RsArea!AreaCode
            Else
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            End If
     Case TDSCat
            FaRstBofEof RsTDSCat
            If Txt(Index).TEXT <> "" Then
                Txt(Index).TEXT = RsTDSCat!Name
                Txt(Index).Tag = RsTDSCat!Code
            Else
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            End If
    Case PartyType
        If RsPartyType.RecordCount > 0 Or (RsPartyType.EOF = False Or RsPartyType.BOF = False) Then
            If Txt(Index).TEXT <> "" Then
                Txt(Index).TEXT = RsPartyType!Name
                Txt(Index).Tag = RsPartyType!Code
            Else
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            End If
        End If
End Select
Set Rst = Nothing
End Sub
Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
    FaCtrl_GetFocus TxtGrid(Index)
    Grid_Hide
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
    Case Col_RefNo
        TxtGrid(0).MaxLength = 15
    End Select
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        TxtGrid(0).TEXT = TxtGrid(0).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        TxtGrid(0).Visible = False
        FGrid.SetFocus
        Exit Sub
    End If
    Select Case FGrid.Col
    Case Col_RefDate, Col_RefNo, Col_CrDr, Col_Amount
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then
                FaGridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_CrDr
            Else
                TxtGrid_LostFocus 0
                TxtGrid(0).SetFocus
            End If
        End If
    End Select
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    FaCheckQuote KeyAscii
    Select Case FGrid.Col
    Case Col_Amount
        FaNumPress TxtGrid(Index), KeyAscii, 10, 2
    End Select
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case FGrid.Col
    Case Col_CrDr
        If Len(TxtGrid(Index)) = 0 Or UCase(Mid(TxtGrid(Index), 1, 1)) = "C" Then
            TxtGrid(Index) = "Cr"
        ElseIf UCase(Mid(TxtGrid(Index), 1, 1)) = "D" Then
            TxtGrid(Index) = "Dr"
        Else
            TxtGrid(Index) = "Cr"
        End If
    Case Col_Amount
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
End Select
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TxtGrid_LostFocus(Index As Integer)
On Error GoTo ELoop
    If ExitCtrl = False Then Exit Sub
    FaCtrl_validate TxtGrid(Index)
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Select Case FGrid.Col
    Case Col_RefDate
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = PubDatamanFa.FaRetDateFunc(TxtGrid(Index))
    Case Col_RefNo, Col_CrDr
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(Index).TEXT
        If FGrid.Col = Col_CrDr Then
            CalOpBal
            If FGrid.TextMatrix(FGrid.Rows - 1, Col_CrDr) <> "" And Val(FGrid.TextMatrix(FGrid.Rows - 1, Col_Amount)) <> 0 Then FGrid.AddItem FGrid.Rows
        End If
    Case Col_Amount
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        CalOpBal
        If FGrid.TextMatrix(FGrid.Rows - 1, Col_CrDr) <> "" And Val(FGrid.TextMatrix(FGrid.Rows - 1, Col_Amount)) <> 0 Then FGrid.AddItem FGrid.Rows
End Select
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
End Sub
Private Sub FGrid_DblClick()
On Error GoTo ELoop
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid.Col
    Case Col_RefDate, Col_RefNo, Col_CrDr, Col_Amount
        FaGridDblClick Me, FGrid, TxtGrid, 0
End Select
TAddMode = False
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FGrid_EnterCell()
    FGrid.CellBackColor = CellBackColEnter
End Sub
Private Sub FGrid_GotFocus()
    FGrid.CellBackColor = CellBackColEnter
    Grid_Hide
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
        SendKeys "{Tab}"
        If DetailFlag = 0 Then If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    Select Case FGrid.Col
        Case Col_RefDate, Col_RefNo, Col_CrDr, Col_Amount
            If KeyCode = vbKeyDelete And Shift = 0 Then
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            End If
    End Select
    If KeyCode = vbKeyReturn Then
        Select Case FGrid.Col
            Case Col_RefDate, Col_RefNo, Col_CrDr, Col_Amount
                FaGridDblClick Me, FGrid, TxtGrid, 0
        End Select
        TAddMode = False
    End If
    KeyCode = 0
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
Select Case FGrid.Col
    Case Col_RefDate, Col_RefNo, Col_CrDr
        FaGet_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
    Case Col_Amount
        FaGet_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGrid.Row >= 1 Then
            If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                If FGrid.Rows > 2 Then
                    FGrid.RemoveItem (FGrid.Row)
                    CalOpBal
                Else
                    FGrid.Rows = 1
                    FGrid.AddItem FGrid.Rows
                    FGrid.FixedRows = 1
                End If
            End If
            For I = 1 To FGrid.Rows - 1
                FGrid.TextMatrix(I, Col_SrNo) = I
            Next
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid.SetFocus
    End If
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub SaveMsg(xObject As Object)
Grid_Hide
If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
    TopCtrl1_eSave
Else
    xObject.SetFocus
End If
End Sub
Private Sub TopCtrl1_ePrn()
    Frame2(5).left = 1000
    Frame2(5).top = 1000
    Frame2(5).Visible = True
    Frame2(5).ZOrder 0
End Sub
Private Sub BTNPRINT_Click()
On Error GoTo errorbox
Dim RST1 As ADODB.Recordset, X11, I As Integer, SqlQry As String
If Master.RecordCount <= 0 Then Exit Sub
SqlQry = ""
If Opt3(3).Value = True Then
    SqlQry = SqlQry & " Where g.groupname='" & DCGroup.TEXT & "'"
End If
If MsgBox("Do You Want Tree Like List", vbQuestion + vbDefaultButton1 + vbYesNo, "A/C Group List") = vbYes Then
    If PubBackEnd = "A" Then
        Set RST1 = G_FaCn.Execute("Select G.MAINGRCODE,G.GROUPCODE,SubCode,Name,G.GroupName As UnderGroup,ConPrefix,ConPerson,Add1 As Address1,Add2 As Address2,Add3 As Address3,C.CityName,Phone,Mobile,FAX,EMail,CSTNo As CST,LSTNo As LST,PANNo As PAN,Switch(ActiveYN=0,'No',ActiveYN=1,'Yes') As Active,Switch(Govt_YN=0,'No',Govt_YN=1,'Yes') As GovtParty,CreditLimit,CreditDays,Remark From (AcGroup G  Left Join SubGroup S on S.GroupCode=G.GroupCode) Left Join City C on S.CityCode=C.CityCode  " & SqlQry & " ORDER BY G.MAINGRCODE,G.GROUPNAME,S.NAME")
    ElseIf PubBackEnd = "S" Then
        Set RST1 = G_FaCn.Execute("Select G.MAINGRCODE,G.GROUPCODE,SubCode,Name,G.GroupName As UnderGroup,ConPrefix,ConPerson,Add1 As Address1,Add2 As Address2,Add3 As Address3,C.CityName,Phone,Mobile,FAX,EMail,CSTNo As CST,LSTNo As LST,PANNo As PAN,Switch(ActiveYN=0,'No',ActiveYN=1,'Yes') As Active,Switch(Govt_YN=0,'No',Govt_YN=1,'Yes') As GovtParty,CreditLimit,CreditDays,Remark From (AcGroup G  Left Join SubGroup S on S.GroupCode=G.GroupCode) Left Join City C on S.CityCode=C.CityCode  " & SqlQry & " ORDER BY G.MAINGRCODE,G.GROUPNAME,S.NAME")
    End If
    If RST1.RecordCount = 0 Then MsgBox "No record Found to Print": Exit Sub
    X11 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaPartyMast.ttx", True)
    Set rpt = rdApp.OpenReport(PubFaReportPath + "\FaPartyTree.RPT")
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("Title")
                rpt.FormulaFields(I).TEXT = "'Ledger A/C List'"
        End Select
    Next
    rpt.Database.SetDataSource RST1
    rpt.ReadRecords
Else
    If PubBackEnd = "A" Then
        Set RST1 = G_FaCn.Execute("Select SubCode,Name,G.GroupName As UnderGroup,ConPrefix,ConPerson,Add1 As Address1,Add2 As Address2,Add3 As Address3,C.CityName,Phone,Mobile,FAX,EMail,CSTNo As CST,LSTNo As LST,PANNo As PAN,Switch(ActiveYN=0,'No',ActiveYN=1,'Yes') As Active,Switch(Govt_YN=0,'No',Govt_YN=1,'Yes') As GovtParty,CreditLimit,CreditDays,Remark From (SubGroup S Left Join AcGroup G on S.GroupCode=G.GroupCode) Left Join City C on S.CityCode=C.CityCode " & SqlQry)
    ElseIf PubBackEnd = "S" Then
        Set RST1 = G_FaCn.Execute("Select SubCode,Name,G.GroupName As UnderGroup,ConPrefix,ConPerson,Add1 As Address1,Add2 As Address2,Add3 As Address3,C.CityName,Phone,Mobile,FAX,EMail,CSTNo As CST,LSTNo As LST,PANNo As PAN,Active=CASE ActiveYN WHEN 0 THEN 'No' WHEN 1 THEN 'Yes' END,GovtParty=CASE Govt_YN WHEN 0 THEN 'No' WHEN 1 THEN 'Yes' END,CreditLimit,CreditDays,Remark From (SubGroup S Left Join AcGroup G on S.GroupCode=G.GroupCode) Left Join City C on S.CityCode=C.CityCode " & SqlQry)
    End If
    If RST1.RecordCount = 0 Then MsgBox "No record Found to Print": Exit Sub
    X11 = CreateFieldDefFile(RST1, PubFaReportPath + "\FaPartyMast.ttx", True)
    Set rpt = rdApp.OpenReport(PubFaReportPath + "\FaPartyMast.RPT")
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("Title")
                rpt.FormulaFields(I).TEXT = "'Ledger A/C List'"
        End Select
    Next
    rpt.Database.SetDataSource RST1
    rpt.ReadRecords
End If
FaReport_View rpt, 0, Me.CAPTION, True
Set RST1 = Nothing
Exit Sub
errorbox:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub btnexit_Click()
    Frame2(5).Visible = False
End Sub
