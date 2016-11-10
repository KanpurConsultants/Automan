VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmBodyBuilding 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Vehicle Purchase "
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   11880
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txt 
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
      Left            =   1155
      MaxLength       =   4
      TabIndex        =   87
      Top             =   5295
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt 
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
      Left            =   630
      MaxLength       =   4
      TabIndex        =   86
      Top             =   5340
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt 
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
      Left            =   780
      MaxLength       =   4
      TabIndex        =   85
      Top             =   5235
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt 
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
      Left            =   525
      MaxLength       =   4
      TabIndex        =   84
      Top             =   5040
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt 
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
      Left            =   1215
      MaxLength       =   4
      TabIndex        =   83
      Top             =   5145
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt 
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
      Left            =   810
      MaxLength       =   4
      TabIndex        =   82
      Top             =   5115
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt 
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
      Left            =   675
      MaxLength       =   4
      TabIndex        =   81
      Top             =   4920
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt 
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
      Left            =   1110
      MaxLength       =   4
      TabIndex        =   80
      Top             =   5010
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt 
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
      Left            =   975
      MaxLength       =   4
      TabIndex        =   79
      Top             =   4830
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt 
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
      Index           =   32
      Left            =   10095
      MaxLength       =   18
      TabIndex        =   75
      Top             =   5790
      Width           =   1470
   End
   Begin VB.TextBox txt 
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
      Index           =   31
      Left            =   9555
      TabIndex        =   74
      Top             =   5775
      Width           =   495
   End
   Begin VB.TextBox txt 
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
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   30
      Left            =   6660
      MaxLength       =   12
      TabIndex        =   23
      Top             =   7440
      Visible         =   0   'False
      Width           =   1470
   End
   Begin MSDataGridLib.DataGrid DGCol 
      Height          =   2445
      Left            =   4650
      Negotiate       =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   8580
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   4313
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
      RowDividerStyle =   0
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
      Caption         =   "Colour Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   "Colors"
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
            ColumnWidth     =   4380.095
         EndProperty
      EndProperty
   End
   Begin VB.TextBox lblGroup 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   0
      TabIndex        =   71
      Top             =   3150
      Visible         =   0   'False
      Width           =   1785
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2955
      Left            =   525
      Negotiate       =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   9060
      Visible         =   0   'False
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   5212
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777152
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
         Weight          =   400
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
      Caption         =   "Party Help"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Party Name"
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
         DataField       =   "Add1"
         Caption         =   "Add1"
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
         DataField       =   "Add2"
         Caption         =   "Add2"
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
         DataField       =   "Add3"
         Caption         =   "Add3"
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
         DataField       =   "City"
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
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
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
      Left            =   10575
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Text            =   "VFa"
      Top             =   1950
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DGMod 
      Height          =   2865
      Left            =   4650
      Negotiate       =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   8715
      Visible         =   0   'False
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   5054
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
      Caption         =   "Model Help"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Model Code"
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
         DataField       =   "ModelGroup"
         Caption         =   "Model Group"
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
         DataField       =   "Colour"
         Caption         =   "Colour"
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
         DataField       =   "Chas_Type"
         Caption         =   "Chassis Type"
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
         DataField       =   "Name"
         Caption         =   "Model Name"
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
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   4995.213
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2175
      Left            =   2310
      Negotiate       =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   7860
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
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
         Weight          =   400
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
      Caption         =   "Site Help"
      ColumnCount     =   2
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "code"
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
            ColumnWidth     =   2310.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgChassis 
      Height          =   2445
      Left            =   330
      Negotiate       =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   7695
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   4313
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
      Caption         =   "Chassis Help"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "CODE"
         Caption         =   "Chassis No."
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
         DataField       =   "EngineNo"
         Caption         =   "Engine No."
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
            ColumnWidth     =   2250.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2160
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
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
      Index           =   16
      Left            =   8790
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Text            =   "0123456789"
      Top             =   1965
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txt 
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
      Index           =   17
      Left            =   10095
      MaxLength       =   4
      TabIndex        =   11
      Top             =   3990
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt 
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
      Index           =   18
      Left            =   9555
      TabIndex        =   17
      Top             =   5535
      Width           =   495
   End
   Begin VB.TextBox txt 
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
      Index           =   19
      Left            =   3015
      TabIndex        =   19
      Top             =   7020
      Width           =   495
   End
   Begin VB.TextBox txt 
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
      Index           =   23
      Left            =   10095
      MaxLength       =   12
      TabIndex        =   15
      Top             =   4950
      Width           =   1470
   End
   Begin VB.TextBox txt 
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
      Index           =   26
      Left            =   3555
      MaxLength       =   18
      TabIndex        =   20
      Top             =   7260
      Width           =   1470
   End
   Begin VB.TextBox txt 
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
      Index           =   27
      Left            =   10095
      MaxLength       =   12
      TabIndex        =   21
      Top             =   6030
      Width           =   1470
   End
   Begin MSDataGridLib.DataGrid DGPCat 
      Height          =   4935
      Left            =   -2070
      Negotiate       =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   7560
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
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
      Caption         =   "Purchase Category "
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Purchase Category"
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
   Begin VB.TextBox txt 
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
      Left            =   9330
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1065
      Width           =   2325
   End
   Begin VB.TextBox txt 
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
      Left            =   6000
      MaxLength       =   12
      TabIndex        =   7
      Top             =   825
      Width           =   1275
   End
   Begin VB.TextBox txt 
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
      Left            =   2010
      MaxLength       =   25
      TabIndex        =   6
      Top             =   825
      Width           =   2775
   End
   Begin VB.TextBox txt 
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
      Index           =   25
      Left            =   10095
      MaxLength       =   18
      TabIndex        =   18
      Top             =   5550
      Width           =   1470
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   9090
      TabIndex        =   49
      Top             =   7185
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   510
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   0
         Width           =   2325
         _ExtentX        =   4101
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
            Name            =   "Arial"
            Size            =   8.25
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
   Begin MSDataGridLib.DataGrid DGADItem 
      Height          =   4935
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   7665
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
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
      RowDividerStyle =   0
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
      Caption         =   "Add/Del Item Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Item Name"
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
   Begin VB.TextBox txt 
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
      Left            =   10485
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1545
      Width           =   1170
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBF0F1&
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
      Index           =   4
      Left            =   2010
      MaxLength       =   40
      TabIndex        =   5
      Top             =   585
      Width           =   5265
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   3555
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   9330
      MaxLength       =   21
      TabIndex        =   1
      Top             =   570
      Width           =   2340
   End
   Begin VB.TextBox txt 
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
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   28
      Left            =   10095
      MaxLength       =   12
      TabIndex        =   22
      Top             =   6270
      Width           =   1470
   End
   Begin VB.TextBox txt 
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
      Index           =   22
      Left            =   10095
      MaxLength       =   12
      TabIndex        =   14
      Top             =   4710
      Width           =   1470
   End
   Begin VB.TextBox txt 
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
      Index           =   21
      Left            =   10095
      MaxLength       =   12
      TabIndex        =   13
      Top             =   4470
      Width           =   1470
   End
   Begin VB.TextBox txt 
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
      Index           =   24
      Left            =   10095
      MaxLength       =   12
      TabIndex        =   16
      Top             =   5310
      Width           =   1470
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   210
      Index           =   20
      Left            =   10095
      MaxLength       =   15
      TabIndex        =   12
      Top             =   4230
      Width           =   1470
   End
   Begin VB.TextBox txt 
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
      Height          =   240
      Index           =   15
      Left            =   2010
      MaxLength       =   40
      TabIndex        =   8
      Top             =   1065
      Width           =   5265
   End
   Begin VB.TextBox txt 
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
      Left            =   9330
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1305
      Width           =   2325
   End
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   4935
      Left            =   8745
      Negotiate       =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   7365
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
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
         Weight          =   400
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
      Caption         =   "Tax Form Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Form Description"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1635
      Left            =   60
      TabIndex        =   10
      Top             =   2310
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   2884
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   18
      BackColorFixed  =   13623520
      ForeColorFixed  =   0
      BackColorSel    =   15196124
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   16744576
      FocusRect       =   0
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   18
   End
   Begin MSDataGridLib.DataGrid DGGod 
      Height          =   2445
      Left            =   8400
      Negotiate       =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   7440
      Visible         =   0   'False
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   4313
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
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
         Weight          =   400
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
      Caption         =   "Godown Help"
      ColumnCount     =   1
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
   Begin VB.Label LblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   45
      TabIndex        =   78
      Top             =   6645
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Surcharge"
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
      Left            =   7905
      TabIndex        =   77
      Top             =   5790
      Width           =   1260
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   9480
      TabIndex        =   76
      Top             =   5760
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subvention Credit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   14
      Left            =   4470
      TabIndex        =   73
      Top             =   7440
      Visible         =   0   'False
      Width           =   1545
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   6255
      TabIndex        =   72
      Top             =   7875
      Width           =   75
   End
   Begin VB.Label LblAcPostDt 
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10125
      TabIndex        =   68
      Top             =   1965
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LblAcPostBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Posting By :"
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
      Left            =   7395
      TabIndex        =   67
      Top             =   1950
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date* :"
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
      Left            =   4935
      TabIndex        =   64
      Top             =   855
      Width           =   645
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   7935
      X2              =   11550
      Y1              =   5250
      Y2              =   5250
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
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
      Left            =   7905
      TabIndex        =   61
      Top             =   4950
      Width           =   810
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   16
      Left            =   9465
      TabIndex        =   60
      Top             =   4920
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   15
      Left            =   2925
      TabIndex        =   59
      Top             =   7005
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Surcharge"
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
      Left            =   1365
      TabIndex        =   58
      Top             =   7260
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Misc Charges"
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
      Left            =   7905
      TabIndex        =   57
      Top             =   6030
      Width           =   1140
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   14
      Left            =   9465
      TabIndex        =   56
      Top             =   6000
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code*"
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
      Left            =   7875
      TabIndex        =   55
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer Bill No.*"
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
      Left            =   75
      TabIndex        =   53
      Top             =   825
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax"
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
      Left            =   7905
      TabIndex        =   52
      Top             =   5550
      Width           =   315
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   9465
      TabIndex        =   51
      Top             =   5520
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Type*"
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
      Left            =   75
      TabIndex        =   47
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      Height          =   1425
      Left            =   7770
      Top             =   510
      Width           =   3975
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Prefix"
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
      Left            =   9780
      TabIndex        =   44
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Srl No.*"
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
      Left            =   7875
      TabIndex        =   43
      Top             =   1545
      Width           =   1530
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division "
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
      Left            =   7875
      TabIndex        =   42
      Top             =   825
      Width           =   735
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code "
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
      Left            =   9930
      TabIndex        =   41
      Top             =   825
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOC ID"
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
      Index           =   42
      Left            =   7860
      TabIndex        =   40
      Top             =   585
      Width           =   675
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   22
      Left            =   9465
      TabIndex        =   38
      Top             =   6240
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
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
      Index           =   41
      Left            =   7905
      TabIndex        =   37
      Top             =   6270
      Width           =   1005
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   21
      Left            =   9465
      TabIndex        =   36
      Top             =   4680
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction"
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
      Index           =   40
      Left            =   7905
      TabIndex        =   35
      Top             =   4710
      Width           =   855
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   20
      Left            =   9465
      TabIndex        =   34
      Top             =   4440
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Addition"
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
      Index           =   39
      Left            =   7905
      TabIndex        =   33
      Top             =   4470
      Width           =   690
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   19
      Left            =   9465
      TabIndex        =   32
      Top             =   5280
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excise"
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
      Left            =   7905
      TabIndex        =   31
      Top             =   5310
      Width           =   540
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   18
      Left            =   9465
      TabIndex        =   30
      Top             =   4200
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Goods Value"
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
      Index           =   37
      Left            =   7905
      TabIndex        =   29
      Top             =   4230
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity"
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
      Left            =   7905
      TabIndex        =   28
      Top             =   3990
      Width           =   1200
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   4
      Left            =   9465
      TabIndex        =   27
      Top             =   3960
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier*"
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
      Left            =   75
      TabIndex        =   26
      Top             =   585
      Width           =   810
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
      Index           =   91
      Left            =   9195
      TabIndex        =   25
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Date*"
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
      Left            =   7875
      TabIndex        =   24
      Top             =   1320
      Width           =   1350
   End
End
Attribute VB_Name = "frmBodyBuilding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim ForeColorSelEnter$
Dim BackColorSelLeave$



Dim mReposting As Boolean
Dim mRePostCounter As Integer


Dim RsParty As ADODB.Recordset
Dim RsChassis As ADODB.Recordset
Dim RsMod As ADODB.Recordset
Dim RsSite As ADODB.Recordset
Dim RsADItem As ADODB.Recordset
Dim rsGod As ADODB.Recordset
Dim rsForm As ADODB.Recordset
Dim RsCol As ADODB.Recordset
Dim RsPCat As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim FirmAddFlag As Byte
Dim GridKey As Integer
Dim DocID As String * 21
Public mVType As String
Dim VoucherEditFlag As Boolean
Dim vPrefix As String
Private Const PurVType As String = "BINV"
Private Const TxtDocID As Byte = 0
Private Const SiteCode As Byte = 1
Private Const VDate As Byte = 2
Private Const SerialNo As Byte = 3
Private Const Party As Byte = 4
Private Const TelcoInvNo As Byte = 5
Private Const TelcoInvDate As Byte = 6
Private Const SuppInvNo As Byte = 7
Private Const SuppInvDate As Byte = 8
Private Const PCat As Byte = 9
Private Const RsoYn As Byte = 10
Private Const RsoCode As Byte = 11
Private Const DueDate As Byte = 12
Private Const ExGP As Byte = 13
Private Const ExDate As Byte = 14
Private Const FormType As Byte = 15
Private Const TotQty As Byte = 17
Private Const TaxPer As Byte = 18
Private Const TaxSurPer As Byte = 19
Private Const TotGoods As Byte = 20
Private Const Addition As Byte = 21
Private Const Deduction As Byte = 22
Private Const SubAmt  As Byte = 23
Private Const ExAmt  As Byte = 24
Private Const TaxAmt As Byte = 25
Private Const TaxSurch As Byte = 26
Private Const MisCharge As Byte = 27
Private Const Gtot As Byte = 28
Private Const AcPostByName      As Byte = 16
Private Const AcPostDate        As Byte = 29
Private Const SubventionCredit  As Byte = 30
Private Const SatPer  As Byte = 31
Private Const SatAmt  As Byte = 32


Private Const SrNo As Byte = 0
Private Const Model1 As Byte = 1
Private Const ChassisNo As Byte = 2
Private Const EngineNo As Byte = 3
Private Const Rate As Byte = 4
Private Const Godown  As Byte = 5
Private Const LdRate  As Byte = 6
Private Const God  As Byte = 7
Private Const ChassisDocid  As Byte = 8



Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

Private Sub cmdPost_Click()
Dim I As Integer, mStartdate As String, mEndDate As String
    mStartdate = InputBox("Posting Required from which Date ?", "Start Date for Posting", PubLoginDate)
    mEndDate = InputBox("Posting Required upto which Date ?", "Last Date for Posting", PubLoginDate)

    mReposting = True
    mRePostCounter = 1
    
    If mStartdate = "" Or mEndDate = "" Then Exit Sub
    mStartdate = MakeDate(mStartdate)
    mEndDate = MakeDate(mEndDate)
    
    If Master.RecordCount > 0 Then Master.MoveFirst
    Do Until Master.EOF
        If IsNull(Master!V_DATE) Then GoTo MyNextRecord
        If Master!V_DATE < CDate(mStartdate) Then GoTo MyNextRecord
        If Master!V_DATE > CDate(mEndDate) Then GoTo MyNextRecord
        
        MoveRec
        
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        
        'Call TopCtrl1_eEdit
        Disp_Text SETS("EDIT", Me, Master)
        'txt(Party).SetFocus
        FGrid.AddItem FGrid.Rows
        
        Call TopCtrl1_eSave
        If Master.EOF = True Then Exit Do
MyNextRecord:
        Master.MoveNext
        Me.Refresh
    Loop
    
    
    
    mRePostCounter = 0
    
    If mStartdate = "" Or mEndDate = "" Then Exit Sub
    If Master.RecordCount > 0 Then Master.MoveFirst
    Do Until Master.EOF
        If IsNull(Master!V_DATE) Then GoTo MyNextRecord1
        If Master!V_DATE < CDate(mStartdate) Then GoTo MyNextRecord1
        If Master!V_DATE > CDate(mEndDate) Then GoTo MyNextRecord1
        
        
        MoveRec
        
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        
        'Call TopCtrl1_eEdit
        Disp_Text SETS("EDIT", Me, Master)
        'txt(Party).SetFocus
        FGrid.AddItem FGrid.Rows
        
        Call TopCtrl1_eSave
        If Master.EOF = True Then Exit Do
MyNextRecord1:
        Master.MoveNext
        Me.Refresh
    Loop
    
    mReposting = False
    
    If Master.RecordCount > 0 Then Call TopCtrl1_eFirst
    MsgBox "Updation Complete", vbInformation, "Re-Updation"
End Sub

Private Sub Command1_Click()
    TopCtrl1.TopText2 = "Edit"
    ProcAcPost False
    TopCtrl1.TopText2 = "Browse"
End Sub

Private Sub DGSite_Click()
    DGSite.Visible = False
    If RsSite.RecordCount > 0 Then
        txt(SiteCode).TEXT = RsSite!Name
        txt(SiteCode).Tag = RsSite!Code
    End If
    txt(SiteCode).SetFocus
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
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
'On Error GoTo ELoop
    Dim RsTemp As ADODB.Recordset
    
    frmBodyBuilding.CAPTION = " Body Building Invoice "
    
    TopCtrl1.Tag = PubUParam: WinSetting Me
    mVType = "BINV"
    If PubVATYN = 1 Then
        Label3(12) = "V A T @"
    End If
    Ini_Grid
    If mVType = PurVType Then
        '*** A/c Posting Status
        txt(AcPostByName).Visible = True
        txt(AcPostDate).Visible = True
        LblAcPostBy.Visible = True
        LblAcPostDt.Visible = True
    End If

    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "Select DocID as Searchcode,Body_Purch.* from Body_Purch where v_type = '" & mVType & "' and left(Docid,1)='" & PubDivCode & "' Order by V_NO desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "Select Top 1 DocID as Searchcode,Body_Purch.* from Body_Purch where v_type = '" & mVType & "' and left(Docid,1)='" & PubDivCode & "' Order by V_NO desc", GCn, adOpenDynamic, adLockOptimistic
    End If
    
    Set RsCol = New ADODB.Recordset
    RsCol.CursorLocation = adUseClient
    RsCol.Open "select Col_code as code,col_Desc  as name from colmast order by col_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGCol.DataSource = RsCol
    
    Set rsGod = New ADODB.Recordset
    rsGod.CursorLocation = adUseClient
    rsGod.Open "select god_code as code,god_name as name from godown where appli_for = 1 order by god_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGod.DataSource = rsGod
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
'    RsParty.Open "select SubGroup.Subcode as code,SubGroup.NAME from SubGroup Where firmCode = '" & PubFirmCode & "' and Nature='Supplier'  order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    If GCn.Execute("Select " & vIsNull("DebtorInSupplierHelp", "0") & " From Syctrl").Fields(0) = 1 Then
        GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,SubGroup.Add1,SubGroup.Add2,SubGroup.Add3,City.CityName as City from ((SubGroup " & _
            "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode)Left Join City on SubGroup.CityCode=City.CityCode )" & _
            "Where  " & _
            "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
            "order by SubGroup.name"
    Else
        GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,SubGroup.Add1,SubGroup.Add2,SubGroup.Add3,City.CityName as City from ((SubGroup " & _
            "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode)Left Join City on SubGroup.CityCode=City.CityCode )" & _
            "Where  " & _
            "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "') " & _
            "order by SubGroup.name"
    End If
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set RsSite = New ADODB.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select site_code as code,site_desc as name from site order by site_desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGSite.DataSource = RsSite
    
    If PubSiebelActiveYn = 1 Then
        Set RsMod = New ADODB.Recordset
        RsMod.CursorLocation = adUseClient
        RsMod.Open "select Model as code,ModelGrp_Name as ModelGroup,Col_desc  as Colour,Model_Desc as NAME, Chas_Type from (Model Left join Model_Grp on model.Grp_Code=Model_Grp.ModelGrp_Code) Left Join ColMast on Model.Col_Code=ColMast.Col_Code where Div_Code='" & PubDivCode & "' order by Model", GCn, adOpenDynamic, adLockOptimistic
        Set DGMod.DataSource = RsMod
    Else
        Set RsMod = New ADODB.Recordset
        RsMod.CursorLocation = adUseClient
        RsMod.Open "select Model as code,Model_Desc as NAME, Chas_Type from Model where (Div_Code='" & PubDivCode & "' or Div_Code='') order by Model", GCn, adOpenDynamic, adLockOptimistic
        Set DGMod.DataSource = RsMod
        DGMod.Columns(1).width = 0
        DGMod.Columns(2).width = 0
    End If

    Set RsChassis = New ADODB.Recordset
    RsChassis.CursorLocation = adUseClient
    RsChassis.Open ("SELECT Veh_Stock.ChassisNo as code, Veh_Stock.EngineNo,Veh_Stock.Chassis_RctDocNo, Veh_Stock.INDATE, Veh_Stock.Srv_BookNo, Veh_Stock.Mfg_Month, Veh_Stock.Mfg_Yr, Veh_Stock.SDM_STM_NO, Veh_Stock.TAX_YN, Godown.God_Name, ColMast.Col_Desc, Veh_Stock.Colour_Code, Veh_Stock.Godown, Veh_Stock.Model " & _
        "FROM (Veh_Stock LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code) " & _
        "LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code LEFT JOIN Body_PurchDetail ON Veh_Stock.ChassisNo=Body_PurchDetail.ChassisNo  " & _
        "where Veh_Stock.ChassisNo NOT IN (select ChassisNo from Body_PurchDetail) AND 1=1"), GCn, adOpenDynamic, adLockOptimistic
    Set DgChassis.DataSource = RsChassis
    
    Set rsForm = New ADODB.Recordset
    With rsForm
        .CursorLocation = adUseClient
'        .Open "SELECT Form_Code as Code, Form_Desc as Name,Tax_Sur_Per,Tax_Per,PurSal_Ac_Code FROM TaxForms where Vehicle_YN = 1 and Trn_Type = 'Purchase' order by Form_Desc ", GCn, adOpenDynamic, adLockOptimistic
        .Open "SELECT T.Form_Code as Code, T.Form_Desc as Name,T.Tax_Sur_Per,T.Tax_Per, T.AddTaxPer,T1.PurSal_Ac_Code,t1.Tax_Ac_Code,t.L_C, T1.AddTaxAc " & _
            "FROM TaxForms as T left join TaxFormsAc T1 on T.Form_Code+'" & PubDivCode & "'=T1.Form_Code+T1.Div_Code " & _
            "where T.Vehicle_YN = 1 and T.Trn_Type = 'Purchase' Order by Form_Desc ", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGForm.DataSource = rsForm
    
    Set RsADItem = New ADODB.Recordset
    With RsADItem
        .CursorLocation = adUseClient
        .Open "SELECT  Prod_Code as code,Prod_name as name,Rate FROM veh_amdModel order by  veh_amdModel.Prod_name ", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGADItem.DataSource = RsADItem
    
    Set RsPCat = New ADODB.Recordset
    RsPCat.CursorLocation = adUseClient
    RsPCat.Open "SELECT BMS_Code as code,BMS_name as name,CREDIT_BMS,days  FROM BMS order by  BMS_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGPCat.DataSource = RsPCat
    
    Call MoveRec
    Disp_Text SETS("INI", Me, Master)
Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TopCtrl1.TopText2 <> "Browse" Then
        If MsgBox("Do you want to exit", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
Set RsParty = Nothing
Set RsMod = Nothing
Set RsADItem = Nothing
Set rsGod = Nothing
Set RsSite = Nothing
Set rsForm = Nothing
Set RsChassis = Nothing
Set RsPCat = Nothing
Set RsCol = Nothing
Set Master = Nothing
Set mListItem = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
'On Error GoTo ErrorLoop
    Call TopCtrl1_eRef
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    LblVPrefix.CAPTION = ""

    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    
    txt(RsoYn) = "Yes"
    txt(RsoCode) = GCn.Execute("select rso_code from syctrl").Fields(0).Value
    'If UCase(left(PubComp_Name, 4)) = "ENAR" Then
        txt(SiteCode).Tag = PubSiteCode
        txt(SiteCode) = PubSiteName
        txt(VDate).SetFocus
    'Else
    '    Txt(SiteCode).SetFocus
    'End If
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim LedgAry(1) As LedgRec, mResult As Byte, mExists As Boolean ', rst As ADODB.Recordset

Dim I As Integer ', mNarr$

'    GSQL = "Select OfftakeIncentiveSrlNo from veh_stock where ChassisNo  = '" & Txt(ChassisNo).Text & "'"
'    Set rst = New ADODB.Recordset
'    rst.CursorLocation = adUseClient
'    rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
'    If rst.RecordCount  > 0 Then
'        If rst!OfftakeIncentiveSrlNo <> "" Then
'            MsgBox "Offtake Incentive Claim made." & vbCrLf & "Deletion denied!", vbCritical, "Deletion Denied"
'            Set rst = Nothing
'            Exit Sub
'        End If
'    End If
'    Set rst = Nothing
    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
    
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ChassisNo) <> "" Then
            GSQL = "Select Sal_Docid from veh_stock where ChassisNo  = '" & FGrid.TextMatrix(I, ChassisNo) & "'"
            If GCn.Execute(GSQL).Fields(0).Value <> "" Then
                mExists = True
            End If
        End If
    Next
    If mExists Then
        MsgBox "Vehicle Sold !", vbCritical, "Delete Check !"
        Exit Sub
    End If
    
    If mVType = PurVType Then
        If AcPostAuthorisation(txt(AcPostByName)) = False Then Exit Sub
    End If
    
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
        If mVType = PurVType Then
            GCnFaV.BeginTrans
            'Unpost Ledger a/c
            CreateLog Me, Master!SearchCode, mReposting
            
            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaV, txt(TxtDocID))
            If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
            'Unposting of Ledger completed
        End If
        For I = 1 To FGrid.Rows - 1
            If GCn.Execute("select count(*) From Veh_order where Chassis = '" & FGrid.TextMatrix(I, ChassisNo) & "'").Fields(0).Value = 0 Then
                'GCn.Execute ("delete from Veh_stock where Chassis_RctSiteCode  = '" & PubSiteCode & "' and Chassis_RctDivCode  = '" & PubDivCode & "' and Chassisno = '" & FGrid.TextMatrix(I, ChassisNo) & "'")
                GCn.Execute ("delete from Body_PurchDetail where  Chassisno = '" & FGrid.TextMatrix(I, ChassisNo) & "'")
'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
'                GCn.Execute "delete from hiscard Where chassis='" & FGrid.TextMatrix(i, ChassisNo) & "'"
            Else
                MsgBox "Chassis No " & FGrid.TextMatrix(I, ChassisNo) & " is Sold" & vbCrLf & "Deletion Denied", vbInformation, "Deletion Denied"
                Exit Sub
            End If
        Next
        GCn.Execute ("delete from Body_Purch where docId = '" & Master!DocID & "'")
        
        If mVType = PurVType Then GCnFaV.CommitTrans
        GCn.CommitTrans
        Master.Requery
        BUTTONS True, Me, Master, 0
        Call MoveRec
    End If
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then GCn.RollbackTrans: GCnFaV.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
Dim I As Integer
     If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
     
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, ChassisNo) <> "" Then
                GSQL = "Select Inv_Docid from Veh_Order where Chassis  = '" & FGrid.TextMatrix(I, ChassisNo) & "'"
                Debug.Print GCn.Execute(GSQL).RecordCount
               If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
                  MsgBox "Vehicle Sold ! Edit ? ", vbCritical
               Else
                If GCn.Execute(GSQL).RecordCount > 0 Then
                    If XNull(GCn.Execute(GSQL).Fields(0)) <> "" Then
                        If Not mReposting Then MsgBox "Vehicle Sold !", vbCritical, "Edit Denied!"
                        If UCase(left(PubComp_Name, 3)) = "LMP" Or UCase(left(PubComp_Name, 4)) = "ENAR" Then
                        Else
                            Exit Sub
                        End If
                    End If
                End If
                End If
            End If
        Next
     
    If AcPostAuthorisation(txt(AcPostByName)) = False Then Exit Sub
    
    Disp_Text SETS("EDIT", Me, Master)
    txt(Party).SetFocus
    FGrid.AddItem FGrid.Rows
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub TopCtrl1_eFirst()
  BUTTONS True, Me, Master, 1
  Call MoveRec
End Sub

Private Sub TopCtrl1_eLast()
 BUTTONS True, Me, Master, 4
 Call MoveRec
End Sub

Private Sub TopCtrl1_eNext()
 BUTTONS True, Me, Master, 3
 Call MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
 BUTTONS True, Me, Master, 2
 Call MoveRec
End Sub

Private Sub TopCtrl1_eCancel()
Dim I As Integer
On Error GoTo ErrorLoop
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_ePrn()
'On Error GoTo ELoop
'Dim RstRep As ADODB.Recordset, RstRep1 As ADODB.Recordset
'Dim mQry As String, I As Integer, X11
'
'
'    mQry = "Select Vp.DocID,Vp.DocIDHelp,Vp.Site_Code,Vp.V_Type,Vp.V_NO,Vp.V_Date, " & _
'            "Vp.PARTYCODE,Vp.PBILL_NO,Vp.PBILL_DATE,Vp.OBNO, " & _
'            "Vp.OBDate,Vp.BMS_CATEGORY,Vp.RSO_WORK,Vp.RSO_Code,Vp.DueDate, " & _
'            "Vp.GATE,Vp.GATEDATE,Vp.Form_Code,Vp.AMOUNT,Vp.Addition,Vp.Deduction,Vp.Exsice, " & _
'            "Vp.Tax_Per,Vp.TaxSur_Per,Vp.Tax_Amt,Vp.TaxSur_Amt,Vp.Misc_Amt, " & _
'            "Vp.Tot_Amount, Vp.U_Name, Vp.U_EntDt, Vp.U_AE,Vp.AcPostByU_Name,Vp.AcPostByU_EntDt,Vp.DrAcCode," & _
'            "Vs.Chassis_RctDocNo ,Vs.Pur_VDate, Vs.Mfg_Month, Vs.Mfg_Yr, Vs.RSO_WORK,Vs.InDate, " & _
'            "Vs.MODEL,Vs.Godown,Vs.ChassisNo,Vs.EngineNo,Vs.VehSerialNo, " & _
'            "Vs.Srv_BookNo,Vs.RATE,Vs.vrate,Vs.Colour_Code,Vs.TAX_YN,Vs.SDM_STM_NO, " & _
'            "Vs.PBILL_NO,Vs.PBILL_DATE,Vs.PartyCode, " & _
'            "Vs.OfftakeIncentiveSrlNo,Vs.OfftakeIncentive,Vs.TgtLinkIncentive,Vs.SubventionSrlNo,Vs.MfgShare, " & _
'            "Sg.Name as PartyName,Sg.Add1,Sg.Add2,Sg.Add3,Sg.LstNo,Sg.CstNo,City.CityName,TF.Form_Desc,TF.Printing_Desc,ColMast.Col_Desc,BMS.BMS_name, M.Model_Desc, Mg.ModelGrp_Name " & _
'            "From (((((((Body_Purch As Vp Left Join Veh_Stock as Vs On Vp.DocId = Pur_DocId) " & _
'            "                         Left Join SubGroup As Sg  On Sg.SubCode   = Vp.PartyCode)    " & _
'            "                         Left Join City            On Sg.CityCode  = City.CityCode)     " & _
'            "                         Left Join TaxForms As TF  On Tf.Form_Code = Vp.Form_Code)      " & _
'            "                         Left Join ColMast         On VS.Colour_Code = ColMast.Col_Code) Left Join BMS On Vp.BMS_CATEGORY = BMS.BMS_Code) " & _
'            "                         Left Join Model M On M.Model=Vs.Model) " & _
'            "                         Left join Model_Grp Mg On Mg.ModelGrp_Code=M.Grp_Code " & _
'            "Where Vp.DocId='" & Master!DocID & "' "
'
'
'    Set RstRep = GCn.Execute(mQry)
'
'
'    mQry = "Select Vp2.DocId,Vp2.Srl_No,Vp2.Site_Code,Vp2.V_TYPE,Vp2.V_NO,Vp2.PROD_CODE,Vp2.trn_type,Vp2.QTY,Vp2.RATE," & cIIF("Vp2.trn_type='A'", "Vp2.QTY*Vp2.RATE", "-1*Vp2.QTY*Vp2.RATE") & " As Amount,veh_amdModel.Prod_name " & _
'            " From Veh_Purch2 As Vp2 Left Join Veh_amdModel On veh_amdModel.Prod_Code=Vp2.Prod_Code Where Vp2.DocId = '" & Master!DocID & "'"
'
'    Set RstRep1 = GCn.Execute(mQry)
'
'    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: Exit Sub
'    X11 = CreateFieldDefFile(RstRep, PubRepoPath + "\Veh_PurchaseBill.ttx", True)
'    X11 = CreateFieldDefFile(RstRep1, PubRepoPath & "\Veh_PurchaseBill1.ttx", True)
'    Set rpt = rdApp.OpenReport(PubRepoPath + "\Veh_PurchaseBill.RPT")
'    rpt.Database.SetDataSource RstRep
'    rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
'    For I = 1 To rpt.FormulaFields.Count
'        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
'            Case UCase("comp_name")
'                rpt.FormulaFields(I).TEXT = "'" & PubComp_Name & "'"
'            Case UCase("comp_add1")
'                rpt.FormulaFields(I).TEXT = "'" & PubComp_Add & "'"
'            Case UCase("comp_add2")
'                rpt.FormulaFields(I).TEXT = "'" & PubComp_Add2 & "'"
'            Case UCase("comp_city")
'                rpt.FormulaFields(I).TEXT = "'" & PubComp_City & "'"
'            Case UCase("title")
'                rpt.FormulaFields(I).TEXT = "'" & "Vehicle Purchase Bill" & "'"
'        End Select
'    Next
'    rpt.ReadRecords
'
'    Call Report_View(rpt, Me.CAPTION, 0, True)
'
'    Set RstRep = Nothing
'    Exit Sub
'ELoop:
'    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eRef()
    RsParty.Requery
    RsMod.Requery
    RsCol.Requery
    RsSite.Requery
    rsForm.Requery
    RsPCat.Requery
    RsADItem.Requery
    RsChassis.Requery
    rsGod.Requery
End Sub
Private Sub TopCtrl1_eSave()
    Dim I As Integer, DocIdHlp$, CardNo$
    Dim Rst As ADODB.Recordset
    Dim mTrans As Boolean
    
    
    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
'    On Error GoTo errlbl
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    If IsValid(txt(SiteCode), "Site Code") = False Then Exit Sub
    If IsValid(txt(VDate), "Purchase Date") = False Then Exit Sub
    If IsValid(txt(SerialNo), "Purchase Srl Number") = False Then Exit Sub
    If IsValid(txt(Party), "Supplier Name") = False Then Exit Sub
    If IsValid(txt(TelcoInvNo), "Manufacturer Bill No.") = False Then Exit Sub
    If IsValid(txt(TelcoInvDate), "Manufacturer Bill Date") = False Then Exit Sub
    'If IsValid(txt(PCat), "Purchase Category") = False Then Exit Sub
    'If txt(SuppInvNo) <> "" Then If IsValid(txt(SuppInvDate), "Supplier Invoice Date") = False Then Exit Sub
    'If IsValid(txt(RsoYn), "RSO Purchase YN") = False Then Exit Sub
    If IsValid(txt(FormType), "Form Type") = False Then Exit Sub
    
    If txt(TelcoInvDate) <> "" Then
        If CDate(txt(TelcoInvDate)) > CDate(txt(VDate)) Then
            MsgBox "Mfg. Bill Date is greater than Purchase Date", vbCritical, "Validation"
            txt(TelcoInvDate).SetFocus
            Exit Sub
        End If
    End If
    If txt(SuppInvDate) <> "" Then
        If CDate(txt(SuppInvDate)) > CDate(txt(VDate)) Then
            MsgBox "Other Dlr Bill Date is greater than Purchase Date", vbCritical, "Validation"
            txt(SuppInvDate).SetFocus
            Exit Sub
        End If
    End If
    If txt(ExDate) <> "" Then
        If CDate(txt(ExDate)) > CDate(txt(VDate)) Then
            MsgBox "Excise Gate Pass Date is greater than Purchase Date", vbCritical, "Validation"
            txt(ExDate).SetFocus
            Exit Sub
        End If
    End If
    If FGrid.Rows = 2 And FGrid.TextMatrix(1, Model1) = "" Then MsgBox "Fill Transaction Data", vbInformation, "Required data": FGrid.Row = 1: FGrid.Col = Model1: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub

    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Model1) <> "" Then
            If Not mReposting Then
                If FGrid.TextMatrix(I, ChassisNo) = "" Then MsgBox "Fill Chassis No. in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = ChassisNo: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
                If FGrid.TextMatrix(I, EngineNo) = "" Then MsgBox "Fill Engine No. in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = EngineNo: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
                If Val(FGrid.TextMatrix(I, Rate)) = 0 Then MsgBox "Fill Rate in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Rate: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
                If FGrid.TextMatrix(I, Godown) = "" Then MsgBox "Fill Godown in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Godown: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
            End If
        End If
    Next
    Amt_Cal True
    '********* cHECKING pOSTING cOTROLS
    If mVType = PurVType Then
        If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
            If ProcAcPost(True) = False Then Me.ActiveControl.SetFocus: Exit Sub
            txt(AcPostByName) = pubUName
            txt(AcPostDate) = PubServerDate
        End If
    End If
    '**********
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        DocID = txt(TxtDocID)
        If GCn.Execute("select count(*) from Body_Purch where Left(DocID,1)='" & PubDivCode & "' And V_Type = '" & mVType & "' And V_No=" & Val(txt(SerialNo)) & "").Fields(0) > 0 Then
            If VoucherEditFlag Then 'And Txt(SerialNo).Visible Then
                MsgBox "Purchase Document No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                Exit Sub
            Else
                SetMax_VoucherPrefix "DocId", "V_PB", "Body_Purch", "V_Date"
                SetMax_VoucherPrefix "DocId", "V_OST", "Body_Purch", "v_date"
                txt(TxtDocID) = GetDocID(GCnFaV, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix, txt(SiteCode).Tag)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
                    MsgBox "Purchase Document No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    Exit Sub
                End If
            End If
        End If
    End If
    DocIdHlp = Replace(txt(TxtDocID), " ", "")
    
    GCn.BeginTrans
    If mVType = PurVType Then GCnFaV.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2 = "Add" Then
'        GCn.Execute ("delete from Body_Purch where DocID='" & txt(TxtDocID) & "'")
        GCn.Execute ("insert into Body_Purch( " & _
            "DocID,DocIDHelp,Site_Code,V_Type,V_NO,V_Date, " & _
            "PARTYCODE,PBILL_NO,PBILL_DATE, " & _
            "Form_Code,AMOUNT,Addition,Deduction,Exsice, " & _
            "Tax_Per,TaxSur_Per,Tax_Amt,TaxSur_Amt, SatPer, SatAmt,Misc_Amt, " & _
            "Tot_Amount, U_Name, U_EntDt, U_AE,AcPostByU_Name,AcPostByU_EntDt, AddBy, AddDate,DrAcCode) " & _
            "values( '" & txt(TxtDocID) & "','" & DocIdHlp & "','" & txt(SiteCode).Tag & txt(SiteCode).Tag & "','" & mVType & "'," & Val(txt(SerialNo)) & "," & ConvertDate(txt(VDate)) & _
            " ,'" & txt(Party).Tag & "','" & txt(TelcoInvNo) & "'," & ConvertDate(txt(TelcoInvDate)) & _
            " ,'" & txt(FormType).Tag & "','" & txt(TotGoods) & "'," & Val(txt(Addition)) & "," & Val(txt(Deduction)) & _
            " , " & Val(txt(ExAmt)) & "," & Val(txt(TaxPer)) & "," & Val(txt(TaxSurPer)) & "," & Val(txt(TaxAmt)) & "," & Val(txt(TaxSurch)) & ", " & Val(txt(SatPer)) & ", " & Val(txt(SatAmt)) & "," & Val(txt(MisCharge)) & _
            " , " & Val(txt(Gtot)) & ", '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & txt(AcPostByName) & "'," & ConvertDate(txt(AcPostDate)) & ", '" & pubUName & "', " & ConvertDateTime(PubServerDate) & ",'" & rsForm!PurSal_Ac_Code & "')")
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaV, txt(TxtDocID), txt(VDate)
    Else    'edit
        CreateLog Me, Master!SearchCode, mReposting
            GCn.Execute ("update Body_Purch set V_Date=" & ConvertDate(txt(VDate)) & ", PARTYCODE='" & txt(Party).Tag & "',PBILL_NO='" & txt(TelcoInvNo) & "',PBILL_DATE=" & ConvertDate(txt(TelcoInvDate)) & ",Form_Code='" & txt(FormType).Tag & "',AMOUNT=" & Val(txt(TotGoods)) & ",Addition=" & Val(txt(Addition)) & ",Deduction=" & Val(txt(Deduction)) & _
                " ,Exsice = " & Val(txt(ExAmt)) & ",TAX_Amt=" & Val(txt(TaxAmt)) & ",TaxSur_Amt=" & Val(txt(TaxSurch)) & ", SatPer = " & Val(txt(SatPer)) & ", SatAmt = " & Val(txt(SatAmt)) & ",TAX_PER=" & Val(txt(TaxPer)) & ",TaxSur_Per=" & Val(txt(TaxSurPer)) & ", MISC_AMT=" & Val(txt(MisCharge)) & _
                " ,Tot_Amount=" & Val(txt(Gtot)) & ", U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE= 'E',AcPostByU_Name='" & txt(AcPostByName) & "',AcPostByU_EntDt=" & ConvertDate(txt(AcPostDate)) & ", ModifyBy = '" & pubUName & "', ModifyDate = " & ConvertDateTime(PubServerDate) & ",DrAcCode= '" & rsForm!PurSal_Ac_Code & _
                "' where DocID='" & txt(TxtDocID) & "'")
    End If
            GCn.Execute ("Delete from Body_PurchDetail where Pur_DocId='" & txt(TxtDocID) & "'")
            For I = 1 To FGrid.Rows - 1
                If FGrid.TextMatrix(I, Model1) <> "" Then
                   If Val(FGrid.TextMatrix(I, ChassisDocid)) = 0 Then
                        GCn.Execute ("insert into Body_PurchDetail " & _
                            "(Pur_DocId,Pur_SrlNo,Pur_DocIDHelp,Pur_SiteCode,Pur_VType,Pur_VNO, " & _
                            "Pur_VDate, " & _
                            "MODEL,Godown,ChassisNo,EngineNo,VehSerialNo, " & _
                            "RATE,vrate, " & _
                            "PBILL_NO,PBILL_DATE,PartyCode, U_Name, U_EntDt,U_AE) " & _
                            "values('" & txt(TxtDocID).TEXT & "'," & I & ",'" & DocIdHlp & "','" & PubSiteCode & txt(SiteCode).Tag & "','" & mVType & "'," & Val(txt(SerialNo).TEXT) & ", " & _
                            "" & ConvertDate(txt(VDate).TEXT) & ", " & _
                            "'" & FGrid.TextMatrix(I, Model1) & "','" & FGrid.TextMatrix(I, God) & "','" & FGrid.TextMatrix(I, ChassisNo) & "','" & FGrid.TextMatrix(I, EngineNo) & "','" & I & "' , " & _
                            "" & Val(FGrid.TextMatrix(I, Rate)) & "," & Val(FGrid.TextMatrix(I, LdRate)) & ", " & _
                            "'" & txt(TelcoInvNo).TEXT & "'," & ConvertDate(txt(TelcoInvDate).TEXT) & ",'" & txt(Party).Tag & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'E')")
                            
                   Else
                        GCn.Execute "update veh_stock set " & _
                            "  Pur_DocId='" & txt(TxtDocID).TEXT & "',Pur_SrlNo=" & I & ",Pur_DocIDHelp='" & DocIdHlp & "',Pur_SiteCode='" & PubSiteCode & txt(SiteCode).Tag & "',Pur_VType='" & mVType & "',Pur_VNO=" & Val(txt(SerialNo).TEXT) & _
                            ", Pur_VDate=" & ConvertDate(txt(VDate).TEXT) & ", MODEL='" & FGrid.TextMatrix(I, Model1) & "',Godown='" & FGrid.TextMatrix(I, God) & "',EngineNo='" & FGrid.TextMatrix(I, EngineNo) & "',VehSerialNo='" & I & _
                            "',RATE=" & Val(FGrid.TextMatrix(I, Rate)) & ",vrate=" & Val(FGrid.TextMatrix(I, LdRate)) & ",PBILL_NO='" & txt(TelcoInvNo).TEXT & "',PBILL_DATE=" & ConvertDate(txt(TelcoInvDate).TEXT) & ",PartyCode='" & txt(Party).Tag & "', U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E' where ChassisNo='" & FGrid.TextMatrix(I, ChassisNo) & "' and  Chassis_RctDocNo = " & Val(FGrid.TextMatrix(I, ChassisDocid)) & ""
                   End If
                End If
            Next
    
    
    
    'A/c Posting
    If mVType = PurVType Then
        If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
            If mRePostCounter = 0 Then ProcAcPost
        End If
    End If
    
    'EOF of A/c Posting Section
    If mVType = PurVType Then GCnFaV.CommitTrans
    GCn.CommitTrans
    mTrans = False
    Set Rst = Nothing
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select DocID as Searchcode,Body_Purch.* from Body_Purch where v_type = '" & mVType & "' and left(Docid,1)='" & PubDivCode & "' And DocId ='" & txt(TxtDocID) & "' Order by V_NO desc")
    End If
    
    RsPCat.Requery
    Master.FIND "DocId = '" & txt(TxtDocID) & "'"
        
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(SerialNo)) > Val(DeCodeDocID(DocID, Document_No)) Then
            MsgBox "Purchase Document No." & Trim(DeCodeDocID(DocID, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
   
errlbl:
    If mTrans Then GCn.RollbackTrans: GCnFaV.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "select VP1.DocID as searchcode, " & cCStr("VP1.V_NO") & " As V_No," & cDt("VP1.V_Date") & " As V_Date,VP1.PBILL_NO as Mfg_BNo," & cDt("VP1.PBILL_DATE") & " as Mfg_BDate,VStk.MODEL,VStk.ChassisNo,VStk.EngineNo,SG.Name " & _
    " from (Body_Purch VP1 left join Body_PurchDetail VStk on VP1.DocID=VSTK.Pur_DocID) " & _
    " left join SubGroup SG on VP1.PartyCode=SG.SubCode " & _
    " where v_type = '" & mVType & "' and left(Docid,1)='" & PubDivCode & "' order by V_Date desc"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("searchcode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select DocID as Searchcode,Body_Purch.* from Body_Purch where v_type = '" & mVType & "' and left(Docid,1)='" & PubDivCode & "' And DocId ='" & MyValue & "' Order by V_NO desc")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
TxtGrid(0).Visible = False
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case Party
        Set DGParty.DataSource = RsParty
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case SiteCode
        Set DGSite.DataSource = RsSite
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Then Exit Sub
        If txt(Index).TEXT = "" Then
            RsSite.MoveFirst
            RsSite.FIND "code ='" & PubSiteCode & "'"
            txt(Index).Tag = RsSite!Code
            txt(Index).TEXT = RsSite!Name
        Else
            If txt(Index).TEXT <> RsSite!Name Then
                RsSite.MoveFirst
                RsSite.FIND "name ='" & txt(Index).TEXT & "'"
            End If
        End If
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsForm!Name Then
            rsForm.MoveFirst
            rsForm.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case PCat
        If RsPCat.RecordCount = 0 Or (RsPCat.EOF = True Or RsPCat.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsPCat!Name Then
            RsPCat.MoveFirst
            RsPCat.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case SerialNo, ExAmt, TaxAmt, TaxSurch, TaxPer, TaxSurPer, SatAmt, SatPer
'        SendKeys "{HOME}+{END}"
End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case Party
        DGridTxtKeyDown DGParty, txt, Party, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
    Case SiteCode
        DGridTxtKeyDown DGSite, txt, Index, RsSite, KeyCode, False, 1
    Case FormType
        DGridTxtKeyDown DGForm, txt, FormType, rsForm, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
    Case PCat
        DGridTxtKeyDown DGPCat, txt, Index, RsPCat, KeyCode, False, 1
End Select
If FrmList.Visible = False And DGSite.Visible = False And DGPCat.Visible = False And DGParty.Visible = False And DGForm.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VDate Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> IIf(UCase(left(PubComp_Name, 3)) = "LMP", SubventionCredit, MisCharge) Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = IIf(UCase(left(PubComp_Name, 3)) = "LMP", SubventionCredit, MisCharge) Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> SiteCode Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> Party Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case SiteCode
        If DGSite.Visible = True Then DGridTxtKeyPress txt, Index, RsSite, KeyAscii, "Name"
    Case Party
        If DGParty.Visible = True Then DGridTxtKeyPress txt, Party, RsParty, KeyAscii, "Name"
        lblGroup.Visible = True: lblGroup.Locked = True: lblGroup.BackColor = vbBlack: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
    Case FormType
        If DGForm.Visible = True Then DGridTxtKeyPress txt, FormType, rsForm, KeyAscii, "Name"
    Case PCat
        If DGPCat.Visible = True Then DGridTxtKeyPress txt, Index, RsPCat, KeyAscii, "Name"
    Case RsoYn
        If UCase(Chr(KeyAscii)) = "Y" Then
            txt(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            txt(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txt(Index) = ""
        End If
        KeyAscii = 0
Case SerialNo
    Call NumPress(txt(Index), KeyAscii, 7, 0)
Case ExAmt, TaxAmt, TaxSurch, SubventionCredit, SatAmt
    Call NumPress(txt(Index), KeyAscii, 8, 2)
Case TaxPer, TaxSurPer
    Call NumPress(txt(Index), KeyAscii, 3, 2)
End Select

'KeyAscii = RetDGKeyAscii()
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs

Select Case Index
'    Case Party
'        If DGParty.Visible = True Then DGridTxtKeyUp txt, Party, RsParty, KeyCode, "Name"
    Case FormType
        If DGForm.Visible = True Then
            If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            txt(TaxPer).TEXT = IIf(IsNull(rsForm!Tax_Per), 0, rsForm!Tax_Per)
            txt(TaxAmt).TEXT = (Val(txt(SubAmt).TEXT) - Val(txt(ExAmt).TEXT)) * Val(txt(TaxPer).TEXT) / 100
            txt(SatPer).TEXT = IIf(IsNull(rsForm!AddTaxPer), 0, rsForm!AddTaxPer)
            txt(SatAmt).TEXT = (Val(txt(SubAmt).TEXT) - Val(txt(ExAmt).TEXT)) * Val(txt(SatPer).TEXT) / 100
            txt(TaxSurPer).TEXT = IIf(IsNull(rsForm!Tax_Sur_Per), 0, rsForm!Tax_Sur_Per)
            txt(TaxSurch).TEXT = Val(txt(TaxSurPer).TEXT) * Val(txt(TaxAmt).TEXT) / 100
            Amt_Cal False
        End If
    Case TaxPer
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = 16 Then Exit Sub
        txt(TaxAmt).TEXT = Format((Val(txt(SubAmt).TEXT) + Val(txt(ExAmt).TEXT)) * Val(txt(TaxPer).TEXT) / 100, "0.00")
        txt(SatAmt).TEXT = Format((Val(txt(SubAmt).TEXT) + Val(txt(ExAmt).TEXT)) * Val(txt(SatPer).TEXT) / 100, "0.00")
        txt(TaxSurch).TEXT = Format(Val(txt(TaxSurPer).TEXT) * Val(txt(TaxAmt).TEXT) / 100, "0.00")
        Amt_Cal False
    Case TaxSurPer
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = 16 Then Exit Sub
        txt(TaxSurch).TEXT = Format(Val(txt(TaxSurPer).TEXT) * Val(txt(TaxAmt).TEXT) / 100, "0.00")
        Amt_Cal False
    Case ExAmt
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = 16 Then Exit Sub
        txt(TaxAmt).TEXT = Format((Val(txt(SubAmt).TEXT) - Val(txt(ExAmt).TEXT)) * Val(txt(TaxPer).TEXT) / 100, "0.00")
        txt(SatAmt).TEXT = Format((Val(txt(SubAmt).TEXT) - Val(txt(ExAmt).TEXT)) * Val(txt(SatPer).TEXT) / 100, "0.00")
        txt(TaxSurch).TEXT = Format(Val(txt(TaxSurPer).TEXT) * Val(txt(TaxAmt).TEXT) / 100, "0.00")
        Amt_Cal False
    Case MisCharge, TaxAmt, TaxSurch, SatAmt, Addition, Deduction
        Amt_Cal False
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim I As Integer
Dim mDays As Integer
Select Case Index
    Case TelcoInvNo
        If txt(TelcoInvNo) <> "" Then
            If GCn.Execute("Select PBILL_NO from Body_Purch where PBILL_NO='" & txt(TelcoInvNo) & "'").RecordCount > 0 Then
                If MsgBox("Mfg. Bill No. already exists, continue ?", vbYesNo + vbCritical + vbDefaultButton2, "Validation") = vbNo Then
                    Cancel = True
                End If
                Me.ActiveControl.SetFocus
            End If
        End If
    Case Party
        If IsValid(txt(Index), "Party") = False Then Cancel = True: Exit Sub
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsParty!Name
            txt(Index).Tag = RsParty!Code
        End If
    Case SiteCode
        If IsValid(txt(Index), "Site Code") = False Then Cancel = True: Exit Sub
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsSite!Name
            txt(Index).Tag = RsSite!Code
        End If
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = rsForm!Name
            txt(Index).Tag = rsForm!Code
        End If
    Case PCat
        If RsPCat.RecordCount = 0 Or (RsPCat.EOF = True Or RsPCat.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
            txt(DueDate).TEXT = ""
        Else
            txt(Index).TEXT = RsPCat!Name
            txt(Index).Tag = RsPCat!Code
'            If IIf(IsNull(RsPCat!CREDIT_BMS), 0, RsPCat!CREDIT_BMS) = 1 Then
                mDays = IIf(IsNull(RsPCat!DAYS), 0, RsPCat!DAYS)
                If txt(TelcoInvDate) <> "" Then
                    txt(DueDate).TEXT = DateAdd("D", mDays, txt(TelcoInvDate))
                Else
                    txt(DueDate).TEXT = ""
                End If
'            End If
        End If
    Case SuppInvDate, ExDate, TelcoInvDate, DueDate
        txt(Index).TEXT = RetDate(txt(Index))
    Case VDate
        If Len(Trim(txt(VDate).TEXT)) = 0 Then
             txt(VDate).TEXT = PubLoginDate
        Else
            txt(Index).TEXT = RetDate(txt(Index))
        End If
        If TopCtrl1.TopText2 = "Add" Then
            If mVType = "V_OST" Or (mVType <> "V_OST" And CheckFinYear(txt(Index))) Then
                txt(TxtDocID) = GetDocID(GCnFaV, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix, txt(SiteCode).Tag)
                DocID = txt(TxtDocID)
            Else
                Cancel = True
            End If
        End If
    Case SerialNo
        If IsValid(txt(SerialNo), "Serial No.") = False Then Cancel = True:   Exit Sub
            If VoucherEditFlag Then      ' Manual
                txt(TxtDocID) = GetDocID(GCnFaV, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix, txt(SiteCode).Tag)
                DocID = txt(TxtDocID)
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select * From Body_Purch Where docid='" & DocID & "'", GCn, adOpenDynamic, adLockOptimistic
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                    Cancel = True
                    txt(SerialNo).SetFocus
                End If
            End If
    Case MisCharge, TaxAmt, TaxPer, TaxSurch, TaxSurPer, MisCharge, ExAmt, SatPer, SatAmt
        Amt_Cal False
    Case SubventionCredit
        txt(Index) = Format(txt(Index), "0.00")
End Select
Set Rst = Nothing
End Sub
Private Sub DGPCat_Click()
    DGPCat.Visible = False
    If RsPCat.RecordCount > 0 Then
        txt(PCat).TEXT = RsPCat!Name
        txt(PCat).Tag = RsPCat!Code
    End If
    txt(PCat).SetFocus
End Sub

Private Sub DgChassis_Click()
    DgChassis.Visible = False
    If RsChassis.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsChassis!Code
    End If
    TxtGrid(0).SetFocus
End Sub
Private Sub DGMod_Click()
If RsMod.RecordCount > 0 Then
    TxtGrid(0).TEXT = RsMod!Code
    FGrid.TextMatrix(FGrid.Row, Model1) = RsMod!Code
End If
TxtGrid(0).SetFocus
DGMod.Visible = False
End Sub

Private Sub DGForm_Click()
    DGForm.Visible = False
    If rsForm.RecordCount > 0 Then
        txt(FormType).TEXT = rsForm!Name
        txt(FormType).Tag = rsForm!Code
    End If
    txt(FormType).SetFocus
End Sub
Private Sub DGGod_Click()
    DGGod.Visible = False
    If rsGod.RecordCount > 0 Then
        TxtGrid(0).TEXT = rsGod!Name
         FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
         FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
    End If
   TxtGrid(0).SetFocus
End Sub

Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        txt(Party).TEXT = RsParty!Name
        txt(Party).Tag = RsParty!Code
    End If
    DGParty.Visible = False
    lblGroup.Visible = False
    txt(Party).SetFocus
End Sub

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
'    FGrid.CellBackColor = CellBackColEnter
    Grid_Hide
   TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
     Select Case FGrid.Col
        Case Model1, ChassisNo, EngineNo
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        Case Godown
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, God) = ""
        Case Rate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "0.00"
    End Select
    Amt_Cal False
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case Model1, ChassisNo, EngineNo, Rate, Godown
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If txt(TelcoInvDate) = "" Then
    MsgBox "Please Enter Manufacturer Bill No. & Date!", vbOKOnly, "Validation"
    txt(TelcoInvDate).SetFocus
    Exit Sub
End If
Dim gg As String
SetMaxLength
    Select Case FGrid.Col
        Case ChassisNo
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
                TxtGrid(0).SelStart = 6
        Case Model1, EngineNo
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
        Case Rate
            Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
            
        Case Godown
           Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid.Rows > 2 Then
                If GCn.Execute("select count(*) From Veh_order where Chassis = '" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "'").Fields(0).Value > 0 Then
                  MsgBox "Chassis Sold" & vbCrLf & "Deletion Denied", vbInformation, "Deletion Denied": FGrid.SetFocus: Exit Sub
                End If
                FGrid.RemoveItem (FGrid.Row)
            Else
                If GCn.Execute("select count(*) From Veh_order where Chassis = '" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "'").Fields(0).Value > 0 Then
                    MsgBox "Chassis Sold" & vbCrLf & "Deletion Denied", vbInformation, "Deletion Denied": FGrid.SetFocus: Exit Sub
                End If
                FGrid.Rows = 1
                FGrid.AddItem FGrid.Rows
                FGrid.FixedRows = 1
            End If
         End If
         For I = 1 To FGrid.Rows - 1
            FGrid.TextMatrix(I, 0) = I
         Next
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
Exit Sub
End Sub
Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
Next I
End Sub
Private Sub MoveRec()
Dim Rs As Recordset, I As Integer
'On Error GoTo error1

'TopCtrl1.tPrn = False
If Master.RecordCount > 0 And Master.EOF = False And Master.BOF = False Then

    DocID = Master!DocID
    txt(TxtDocID).TEXT = Master!DocID
    LblDiv.CAPTION = "Division : " & left(Master!DocID, 1)
    LblSite.CAPTION = "Site Code : " & mID(Master!Site_Code, 1, 1)
    txt(SiteCode).Tag = mID(Master!Site_Code, 2, 1)
    txt(SiteCode).TEXT = GCn.Execute("select site_desc from site where site_code = '" & txt(SiteCode).Tag & "'").Fields(0).Value
    LblVPrefix.CAPTION = DeCodeDocID(Master!DocID, Document_Prefix)
    txt(SerialNo).TEXT = Master!V_NO
    txt(VDate).TEXT = Master!V_DATE
    txt(Party).Tag = Master!PartyCode
    If txt(Party).Tag <> "" Then
        Set Rs = GCn.Execute("select NAME from SubGroup where Subcode = '" & txt(Party).Tag & "'")
        If Rs.RecordCount > 0 Then
            txt(Party).TEXT = Rs(0)
        Else
            txt(Party).TEXT = ""
        End If
    Else
        txt(Party).TEXT = ""
    End If
    LblUser = IIf(Not IsNull(Master!AddDate), "Add By : " & XNull(Master!AddBy) & "  Dated : " & XNull(Master!AddDate), "") & IIf(Not IsNull(Master!ModifyDate), "     Modify By : " & XNull(Master!ModifyBy) & "  Dated : " & XNull(Master!ModifyDate), "")

    
    txt(FormType).Tag = IIf(IsNull(Master!Form_Code), "", Master!Form_Code)
    If txt(FormType).Tag <> "" Then
        txt(FormType).TEXT = GCn.Execute("select form_desc from taxforms where form_code = '" & txt(FormType).Tag & "'").Fields(0).Value
    Else
        txt(FormType).TEXT = ""
    End If

    txt(TelcoInvNo).TEXT = IIf(IsNull(Master!PBILL_NO), "", Master!PBILL_NO)
    txt(TelcoInvDate).TEXT = IIf(IsNull(Master!PBILL_DATE), "", Master!PBILL_DATE)

    txt(TotGoods).TEXT = Format(IIf(IsNull(Master!Amount) Or Master!Amount = 0, "", Master!Amount), "0.00")
    txt(Addition) = Format(IIf(IsNull(Master!Addition) Or Master!Addition = 0, "", Master!Addition), "0.00")
    txt(Deduction) = Format(IIf(IsNull(Master!Deduction) Or Master!Deduction = 0, "", Master!Deduction), "0.00")
    txt(ExAmt).TEXT = Format(IIf(IsNull(Master!exsice) Or Master!exsice = 0, "", Master!exsice), "0.00")
    txt(TaxPer).TEXT = Format(IIf(IsNull(Master!Tax_Per) Or Master!Tax_Per = 0, "", Master!Tax_Per), "0.00")
    txt(TaxSurPer).TEXT = Format(IIf(IsNull(Master!TaxSur_Per) Or Master!TaxSur_Per = 0, "", Master!TaxSur_Per), "0.00")
    txt(TaxAmt).TEXT = Format(IIf(IsNull(Master!Tax_Amt) Or Master!Tax_Amt = 0, "", Master!Tax_Amt), "0.00")
    txt(SatPer).TEXT = Format(IIf(IsNull(Master!SatPer) Or Master!SatPer = 0, "", Master!SatPer), "0.00")
    txt(SatAmt).TEXT = Format(IIf(IsNull(Master!SatAmt) Or Master!SatAmt = 0, "", Master!SatAmt), "0.00")
    txt(TaxSurch).TEXT = Format(IIf(IsNull(Master!TaxSur_Amt) Or Master!TaxSur_Amt = 0, "", Master!TaxSur_Amt), "0.00")
    txt(MisCharge).TEXT = Format(IIf(IsNull(Master!Misc_Amt) Or Master!Misc_Amt = 0, "", Master!Misc_Amt), "0.00")
    txt(Gtot).TEXT = Format(IIf(IsNull(Master!Tot_Amount) Or Master!Tot_Amount = 0, "", Master!Tot_Amount), "0.00")
    '*** A/c Posting Status
    txt(AcPostByName) = IIf(IsNull(Master!AcPostByU_Name), "", Master!AcPostByU_Name)
    txt(AcPostDate) = IIf(IsNull(Master!AcPostByU_EntDt), "", Master!AcPostByU_EntDt)
    '***
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT Godown.God_Name, Body_PurchDetail.pur_SrlNo,  Body_PurchDetail.MODEL, Body_PurchDetail.ChassisNo, Body_PurchDetail.EngineNo, Body_PurchDetail.VehSerialNo,  Body_PurchDetail.RATE, Body_PurchDetail.VRATE, Body_PurchDetail.godown " & _
            "  " & _
            " FROM (Body_PurchDetail LEFT JOIN Godown ON Body_PurchDetail.Godown = Godown.God_Code)  " & _
            " where Body_PurchDetail.Pur_DocId = '" & Master!DocID & "'")
    FGrid.Rows = 1: FGrid.Redraw = False
    I = 1
    If Rs.RecordCount > 0 Then
        Do Until Rs.EOF
            With FGrid
                .AddItem ""
                .TextMatrix(I, 0) = Rs!Pur_SrlNo
                .TextMatrix(I, Model1) = Rs!Model
                .TextMatrix(I, ChassisNo) = Rs!ChassisNo
                .TextMatrix(I, EngineNo) = IIf(IsNull(Rs!EngineNo), "", Rs!EngineNo)
                .TextMatrix(I, Rate) = Format(IIf(IsNull(Rs!Rate), "", Rs!Rate), "0.00")
                .TextMatrix(I, Godown) = IIf(IsNull(Rs!God_Name), "", Rs!God_Name)
                .TextMatrix(I, LdRate) = Format(IIf(IsNull(Rs!vrate), "", Rs!vrate), "0.00")
                .TextMatrix(I, God) = IIf(IsNull(Rs!Godown), "", Rs!Godown)
            End With
            Rs.MoveNext
           I = I + 1
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    FGrid.Redraw = True
    
    
    
Else
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    
End If
Grid_Hide
Amt_Cal False
Exit Sub
error1:
        CheckError
End Sub

Private Sub Ini_Grid()
 'SrNo.0|Model 1|Chassis No 2|Engine No 3|Serial No 4|Colour 5|Tax 6|Rate 7|Mfg Month 8
'|Year 9|Service Book No 10|SDM/STM No 11|InDate 12|Godown 13|Ld. Rate 14| ColCode 15|God 16
'SrNo.0|Add/Del Item 1|Type 2|Qty 3|Rate 4|Amount 5|Itemcode 6
'Dim i As Byte

    With FGrid
        .left = Me.left '+45
        .width = Me.width - 90
        .top = 2550
        .Cols = 9
        .RowHeightMin = PubGridRowHeight

        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 500

        .TextMatrix(0, Model1) = " Model Code"
        .ColAlignment(Model1) = flexAlignLeftCenter
        .ColWidth(Model1) = 1700
        
        .TextMatrix(0, ChassisNo) = "Chassis No"
        .ColAlignment(ChassisNo) = flexAlignLeftCenter
        .ColWidth(ChassisNo) = 1900
        
        .TextMatrix(0, EngineNo) = "Engine No"
        .ColAlignment(EngineNo) = flexAlignLeftCenter
        .ColWidth(EngineNo) = 2000
        
        
        .TextMatrix(0, Rate) = "Rate"
        .ColAlignmentFixed(Rate) = flexAlignRightCenter
        .ColWidth(Rate) = 1100
      
        
        .TextMatrix(0, Godown) = "Godown"
        .ColAlignment(Godown) = flexAlignLeftCenter
        .ColWidth(Godown) = 1200
        
        .TextMatrix(0, LdRate) = "VRate"
        .ColAlignmentFixed(LdRate) = flexAlignRightCenter
        .ColWidth(LdRate) = 1100
        
        .ColWidth(God) = 0
        .ColWidth(ChassisDocid) = 0
    End With
    
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    
    
DGGod.left = 6630: DGGod.top = mTopScale: DGGod.height = FGrid.top - mTopScale
DgChassis.left = 3165: DgChassis.top = mTopScale: DgChassis.height = FGrid.top - mTopScale
DGMod.left = 0: DGMod.width = Me.width - 90: DGMod.top = FGrid.top + FGrid.height: DGMod.height = Me.height - (DGMod.top + mBotScale)
DGSite.left = 4275: DGSite.top = mTopScale
DGCol.left = 6630: DGCol.top = mTopScale: DGCol.height = FGrid.top - mTopScale
DGParty.width = 11535:   DGParty.left = 0: DGParty.top = FGrid.top  '390
DGParty.height = 5160
DGForm.left = 6630: DGForm.top = mTopScale
DGADItem.left = 6630: DGADItem.top = mTopScale
DGPCat.left = 6630: DGPCat.top = mTopScale
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer


For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
    txt(I).ForeColor = CtrlFColOrg
Next

txt(SiteCode).Enabled = False

If TopCtrl1.TopText2 = "Edit" Then
    txt(SiteCode).Enabled = False
    'txt(VDate).Enabled = False
    txt(SerialNo).Enabled = False
    txt(ExGP).Enabled = True
    txt(ExDate).Enabled = True
    txt(SubventionCredit).Enabled = True
'    If mSoldVehicle Then
'        FGrid.Enabled = False
'    End If
End If

txt(TxtDocID).Enabled = False
txt(TotQty).Enabled = False
txt(TotGoods).Enabled = False
txt(Gtot).Enabled = False
txt(SubAmt).Enabled = False

txtDisabled_Color Me

TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol


End Sub

Private Sub Grid_Hide()
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGMod.Visible = True Then DGMod.Visible = False
    If DgChassis.Visible = True Then DgChassis.Visible = False
    If DGCol.Visible = True Then DGCol.Visible = False
    If DGForm.Visible = True Then DGForm.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If lblGroup.Visible = True Then lblGroup.Visible = False
    If DGADItem.Visible = True Then DGADItem.Visible = False
    If DGGod.Visible = True Then DGGod.Visible = False
    If DGPCat.Visible = True Then DGPCat.Visible = False
    If DGSite.Visible = True Then DGSite.Visible = False
End Sub
Private Sub DGParty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGParty.Row >= 0 Then
    lblGroup.TEXT = G_FaCn.Execute("Select AcGroup.GroupName from (AcGroup Left Join SubGroup on SubGroup.GroupCode=AcGroup.GroupCode) where SubGroup.SubCode='" & RsParty!Code & "'").Fields(0).Value
    lblGroup.Refresh
End If
End Sub
Private Sub TxtGrid_GotFocus(Index As Integer)
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
         Case Model1
            If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Model1) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Model1) <> RsMod!Code Then
                RsMod.MoveFirst
                RsMod.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, Model1) & "'"
            End If
        Case ChassisNo
'            If RsChassis.State <> 0 Then RsChassis.Close
'            If FGrid.TextMatrix(FGrid.Row, Model1) = "" Then MsgBox "Select Model First", vbInformation, "Validation": FGrid.Col = Model1: TxtGrid(0).Visible = False: FGrid.SetFocus: Exit Sub
'                Set RsChassis = New ADODB.Recordset
'                RsChassis.CursorLocation = adUseClient
'            RsChassis.Open ("SELECT Veh_Stock.ChassisNo as code, Veh_Stock.EngineNo,Veh_Stock.Chassis_RctDocNo, Veh_Stock.INDATE, Veh_Stock.Srv_BookNo, Veh_Stock.Mfg_Month, Veh_Stock.Mfg_Yr, Veh_Stock.SDM_STM_NO, Veh_Stock.TAX_YN, Godown.God_Name, ColMast.Col_Desc, Veh_Stock.Colour_Code, Veh_Stock.Godown " & _
'                "FROM (Veh_Stock LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code " & _
'                "where Veh_Stock.MODEL  = '" & FGrid.TextMatrix(FGrid.Row, Model1) & "' and (Veh_Stock.Pur_DocId='' or Veh_Stock.Pur_DocId Is Null)"), GCn, adOpenDynamic, adLockOptimistic
'            Set DgChassis.DataSource = RsChassis
            If RsChassis.RecordCount = 0 Or (RsChassis.EOF = True Or RsChassis.BOF = True) Or FGrid.TextMatrix(FGrid.Row, ChassisNo) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, ChassisNo) <> RsChassis!Code Then
                RsChassis.MoveFirst
                RsChassis.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "'"
            End If
        Case Godown
            If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Godown) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Godown) <> rsGod!Name Then
                rsGod.MoveFirst
                rsGod.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, Godown) & "'"
            End If
         Case Rate
'            SendKeys "{HOME}+{END}"
    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    TxtGrid(0).TEXT = TxtGrid(0).Tag
    TxtGrid_KeyUp Index, KeyCode, Shift
    TxtGrid(0).Visible = False
    Grid_Hide
    FGrid.SetFocus
    Exit Sub
End If
Select Case FGrid.Col

    Case ChassisNo
    
        DGridTxtKeyDown DgChassis, TxtGrid, Index, RsChassis, KeyCode, True, 0
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, FGrid.Cols - 1
            End If
        End If
    Case Model1    '1
        DGridTxtKeyDown DGMod, TxtGrid, Index, RsMod, KeyCode, True, 0, frmModel, "frmModel"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, FGrid.Cols - 1
            End If
        End If
    Case Godown
        DGridTxtKeyDown DGGod, TxtGrid, 0, rsGod, KeyCode, True, 1, frmGodown, "frmGodown"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, FGrid.Cols - 1, 10
            End If
        End If
    Case EngineNo, Rate
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, FGrid.Cols - 1
            End If
        End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case FGrid.Col
    Case ChassisNo
        If StrCmp(left(PubComp_Name, 4), "yash") Then
        Debug.Print TxtGrid(Index).SelStart
            If TxtGrid(Index).SelStart <= 6 Then
                If TxtGrid(Index).SelStart < 6 Then
                    KeyAscii = 0
                ElseIf TxtGrid(Index).SelStart = 6 And (KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
                    KeyAscii = 0
                End If
            End If
        End If
        If DgChassis.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsChassis, KeyAscii, "Code"
    Case Model1
        If DGMod.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsMod, KeyAscii, "Code"
    Case Godown
        If DGGod.Visible = True Then DGridTxtKeyPress TxtGrid, 0, rsGod, KeyAscii, "Name"
    Case Rate
       Call NumPress(TxtGrid(0), KeyAscii, 8, 2)
End Select
'KeyAscii = RetDGKeyAscii(GridHelp, KeyAscii)
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
        Select Case FGrid.Col
            Case Model1
                If KeyCode <> 13 And DGMod.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsMod, KeyCode, "code", True
            Case ChassisNo
                If KeyCode <> 13 And DgChassis.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
            Case Godown
                If KeyCode <> 13 And DGGod.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, 0, rsGod, KeyCode, "Name", True
            Case Rate
                FGrid.TextMatrix(FGrid.Row, Rate) = Format(Val(TxtGrid(Index).TEXT), "0.00")
                Amt_Cal False
        End Select
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim j As Integer
Dim Rst As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim GridCol As Byte
GridCol = FGrid.Col
Select Case GridCol
        Case Model1
            If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, Model1) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, Model1) = RsMod!Code
                If TopCtrl1.TopText2 = "Add" And left(FGrid.TextMatrix(FGrid.Row, ChassisNo), 6) <> RsMod!Chas_Type Then
                    FGrid.TextMatrix(FGrid.Row, ChassisNo) = RsMod!Chas_Type
                
                    If txt(TelcoInvDate) <> "" Then
                        Set Rst = New ADODB.Recordset
                        Rst.CursorLocation = adUseClient
                        Rst.Open "Select top 1 TAXABLE_YN,P_RATE From Veh_Rate Where model='" & FGrid.TextMatrix(FGrid.Row, Model1) & "' And Effective_Date<=" & ConvertDate(txt(TelcoInvDate)) & " Order by Effective_Date Desc", GCn, adOpenDynamic, adLockOptimistic
                        If Rst.RecordCount > 0 Then
                            FGrid.TextMatrix(FGrid.Row, Rate) = Format(IIf(IsNull(Rst!p_rate), 0, Rst!p_rate), "0.00")
                        Else
                            FGrid.TextMatrix(FGrid.Row, Rate) = ""
                        End If
                    End If
                End If
                
                Amt_Cal False
            End If
            If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
        Case ChassisNo
            FGrid = TxtGrid(0)
            If ChkDul_Chassis = True Then TxtGridLeave = False: Exit Function
'Modi Shekhar

            If FGrid.TextMatrix(FGrid.Row, ChassisNo) <> "" Then
            If GCn.Execute("select count(*) From Veh_order where Chassis = '" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "'").Fields(0).Value > 0 Then
                If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
                  MsgBox "Chassis Sold and Issue To Body Builder" '& vbCrLf & "Editing Denied", vbInformation, "Editing Denied" ': FGrid.SetFocus: TxtGridLeave = False: Exit Function
                Else
                  MsgBox "Chassis Sold" & vbCrLf & "Editing Denied", vbInformation, "Editing Denied": FGrid.SetFocus: TxtGridLeave = False: Exit Function
                End If
            End If
            End If
            'And TopCtrl1.TopText2 <> "Edit"
            If FGrid.TextMatrix(FGrid.Row, ChassisNo) <> "" Then
                If RsChassis.RecordCount = 0 Or (RsChassis.EOF = True Or RsChassis.BOF = True) Or TxtGrid(0).TEXT = "" Then
                    Fill_Data False
                Else
                    If UCase(Trim(TxtGrid(0).TEXT)) <> UCase(RsChassis!Code) Then
                        Fill_Data False
                    Else
                        Fill_Data True
                    End If
                End If
            End If
            FGrid.TextMatrix(FGrid.Row, ChassisNo) = UCase(TxtGrid(0).TEXT)
'end modi
            If FGrid.TextMatrix(FGrid.Row, ChassisNo) <> "" Then
                If PubVehGodown <> "" Then
                    If rsGod.RecordCount > 0 And Trim(FGrid.TextMatrix(FGrid.Row, Godown)) = "" Then
                        rsGod.MoveFirst
                        rsGod.FIND "Code ='" & PubVehGodown & "'"
                        FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
                        FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
                    End If
                End If
            End If
         Case EngineNo
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = UCase(TxtGrid(0).TEXT)

        Case Godown
            If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, Godown) = ""
                FGrid.TextMatrix(FGrid.Row, God) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
                FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
            End If
         Case Rate
                FGrid.TextMatrix(FGrid.Row, Rate) = Format(Val(TxtGrid(0).TEXT), "0.00")
                Amt_Cal False
End Select
Set Rst = Nothing
    TxtGridLeave = True
    If ValidateCall = False Then
        FGrid.SetFocus
        TxtGrid(0).Visible = False
    End If
End Function

Private Function ChkDuplicate() As Boolean
Dim I As Integer
Dim X As String, Y As String
Dim Col1 As Byte, Col2 As Byte
    Select Case FGrid.Col
    Case Model1
        Col2 = Model1
        Col1 = ChassisNo
    Case ChassisNo
        Col1 = Model1
        Col2 = ChassisNo
    End Select
    X = UCase(CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col1))) + CStr(Trim(TxtGrid(0).TEXT)))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col1))) + CStr(Trim(FGrid.TextMatrix(I, Col2))))
        If X = Y And Y <> "" Then
            MsgBox "Duplicate Item Not Allowed", vbInformation, "Validation"
            TxtGrid(0).SetFocus
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function





Private Sub FGrid1_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub




 Private Sub Amt_Cal(Lrate As Boolean)
 Dim I As Byte
 Dim ICnt As Integer
 Dim TOTAmt As Double
 Dim TotAdd As Double
 Dim TotDel As Double
 Dim LAdd As Double
 Dim LAdd1 As Double
 Dim LRateItem As Double
 Dim LRateVal As Double
 
    For I = 1 To FGrid.Rows - 1
       If FGrid.TextMatrix(I, Model1) <> "" Then
            TOTAmt = TOTAmt + Val(FGrid.TextMatrix(I, Rate))
            ICnt = ICnt + 1
       End If
    Next I
    
    
    txt(TotQty) = Format(ICnt, "0")
    txt(TotGoods).TEXT = Format(TOTAmt, "0.00")
    'txt(Addition).TEXT = Format(Val(txt(Addition)), "0.00")
    'txt(Deduction).TEXT = Format(Val(txt(Deduction)), "0.00")
    txt(SubAmt).TEXT = Format((TOTAmt + Val(txt(Addition)) - Val(txt(Deduction))), "0.00")
    txt(TaxAmt).TEXT = Format((Val(txt(SubAmt).TEXT) + Val(txt(ExAmt).TEXT)) * Val(txt(TaxPer).TEXT) / 100, "0.00")
    txt(SatAmt).TEXT = Format((Val(txt(SubAmt).TEXT) + Val(txt(ExAmt).TEXT)) * Val(txt(SatPer).TEXT) / 100, "0.00")
    txt(TaxSurch).TEXT = Format(Val(txt(TaxSurPer).TEXT) * Val(txt(TaxAmt).TEXT) / 100, "0.00")

    txt(Gtot).TEXT = Format((Val(txt(SubAmt).TEXT) + Val(txt(ExAmt).TEXT) + Val(txt(TaxAmt).TEXT) + Val(txt(SatAmt).TEXT) + Val(txt(TaxSurch).TEXT) + Val(txt(MisCharge).TEXT)), "0.00")
     
    If Lrate = True Then
        LAdd = (Val(txt(ExAmt).TEXT) + Val(txt(TaxAmt).TEXT) + Val(txt(SatAmt).TEXT) + Val(txt(TaxSurch).TEXT) + Val(txt(MisCharge).TEXT))
        For I = 1 To FGrid.Rows - 1
           If FGrid.TextMatrix(I, Model1) <> "" Then
                LRateItem = Val(FGrid.TextMatrix(I, Rate)) + TotAdd - TotDel
                FGrid.TextMatrix(I, LdRate) = LRateItem + ((LRateItem * LAdd) / Val(txt(SubAmt).TEXT))
           End If
           LRateVal = LRateVal + Val(FGrid.TextMatrix(I, LdRate))
        Next I
        If Val(txt(Gtot).TEXT) - LRateVal <> 0 Then
            For I = 1 To FGrid.Rows - 1
               If FGrid.TextMatrix(I, Model1) <> "" Then
                    FGrid.TextMatrix(1, LdRate) = Val(FGrid.TextMatrix(I, LdRate)) + (Val(txt(Gtot).TEXT) - LRateVal)
                    Exit For
               End If
            Next I
        End If
    End If
End Sub


Private Sub Fill_Data(Enb As Boolean)
If Enb = True Then
    FGrid.TextMatrix(FGrid.Row, Model1) = XNull(RsChassis!Model)
    FGrid.TextMatrix(FGrid.Row, EngineNo) = IIf(IsNull(RsChassis!EngineNo), "", RsChassis!EngineNo)
    FGrid.TextMatrix(FGrid.Row, God) = IIf(IsNull(RsChassis!Godown), "", RsChassis!Godown)
    FGrid.TextMatrix(FGrid.Row, Godown) = IIf(IsNull(RsChassis!God_Name), "", RsChassis!God_Name)
    FGrid.TextMatrix(FGrid.Row, ChassisDocid) = IIf(IsNull(RsChassis!Chassis_RctDocNo), 0, RsChassis!Chassis_RctDocNo)
Else
    FGrid.TextMatrix(FGrid.Row, EngineNo) = ""
    FGrid.TextMatrix(FGrid.Row, God) = ""
    FGrid.TextMatrix(FGrid.Row, Godown) = ""
    FGrid.TextMatrix(FGrid.Row, ChassisDocid) = ""
End If
End Sub

Private Function ChkDul_Chassis() As Boolean
Dim I As Integer

ChkDul_Chassis = False
End Function

Private Function ProcAcPost(Optional CheckCtrls As Boolean) As Boolean
        rsForm.MoveFirst        'Vehicle Purchase A/c Code
        rsForm.FIND "Name ='" & txt(FormType) & "'"
        If IsNull(rsForm!PurSal_Ac_Code) Or rsForm!PurSal_Ac_Code = "" Then
            MsgBox "Please Define Purchase A/c in Tax Forms, Purchase Bill Add/Edit Denied !" & vbCrLf & "Please Define A/c in Tax Forms", vbInformation, "Validation"
            ProcAcPost = False: Exit Function
        End If
        If CheckCtrls Then
            ProcAcPost = True: Exit Function
        End If
        
        Dim I As Integer
        'A/c Posting related declarations
        Dim LedgAry(4) As LedgRec, mResult As Byte, mNarr$
        
        mNarr = "Through Vehicle Purchase Mfg Bill No." & txt(TelcoInvNo) & " Date " & txt(TelcoInvDate) & " Chassis No."
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Model1) <> "" Then
                mNarr = mNarr & FGrid.TextMatrix(I, ChassisNo) & "."
            End If
        Next
        I = 0
        
        If Val(txt(Gtot)) <> 0 Then
            If PubVATYN = 1 And rsForm!L_C = "Local" Then
            'Purchase A/c
                LedgAry(I).SubCode = rsForm!PurSal_Ac_Code
                LedgAry(I).AmtDr = Val(txt(Gtot)) - Val(txt(TaxAmt)) - Val(txt(SatAmt))
                LedgAry(I).Narration = mNarr
                I = I + 1
                
                'Tax Amt A/c
                If Val(txt(TaxAmt)) > 0 Then
                    LedgAry(I).SubCode = rsForm!Tax_Ac_Code
                    LedgAry(I).AmtDr = Val(txt(TaxAmt))
                    LedgAry(I).Narration = mNarr
                    I = I + 1
                End If
                
                'SAT Amt A/c
                If Val(txt(SatAmt)) > 0 Then
                    LedgAry(I).SubCode = rsForm!AddTaxAc
                    LedgAry(I).AmtDr = Val(txt(SatAmt))
                    LedgAry(I).Narration = mNarr
                    I = I + 1
                End If
                
            Else
                'Purchase A/c
                LedgAry(I).SubCode = rsForm!PurSal_Ac_Code
                LedgAry(I).AmtDr = Val(txt(Gtot))
                LedgAry(I).Narration = mNarr
                I = I + 1
            End If
            
            'Party A/c
            LedgAry(I).SubCode = txt(Party).Tag
            LedgAry(I).AmtCr = Val(txt(Gtot))
            LedgAry(I).Narration = mNarr
        End If
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, txt(TxtDocID), CDate(txt(VDate)), mNarr & "[Common]")
        If mResult <> 1 Then
            MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
            ProcAcPost = False
        Else
            ProcAcPost = True
        End If
End Function

Private Sub SetMaxLength()
    Select Case FGrid.Col
        Case ChassisNo
            TxtGrid(0).MaxLength = 20
        Case EngineNo
            TxtGrid(0).MaxLength = 25
        Case Model1, EngineNo, Godown
             TxtGrid(0).MaxLength = 0
    End Select
End Sub

