VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehPur 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Vehicle Purchase "
   ClientHeight    =   7770
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
   ScaleHeight     =   7770
   ScaleWidth      =   11880
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "AcPost"
      Height          =   330
      Left            =   6825
      TabIndex        =   99
      Top             =   15
      Width           =   1290
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
      TabIndex        =   95
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
      TabIndex        =   94
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
      Left            =   10095
      MaxLength       =   12
      TabIndex        =   33
      Top             =   6510
      Width           =   1470
   End
   Begin MSDataGridLib.DataGrid DGCol 
      Height          =   2445
      Left            =   2640
      Negotiate       =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   7470
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
   Begin VB.CommandButton cmdPost 
      Caption         =   "Re-Update Ledger"
      Height          =   330
      Left            =   8715
      TabIndex        =   90
      Top             =   -15
      Width           =   2040
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
      TabIndex        =   89
      Top             =   3150
      Visible         =   0   'False
      Width           =   1785
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2955
      Left            =   -4545
      Negotiate       =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   7635
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
      TabIndex        =   87
      TabStop         =   0   'False
      Text            =   "VFa"
      Top             =   1950
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DGMod 
      Height          =   2865
      Left            =   -690
      Negotiate       =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   7500
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
      TabIndex        =   83
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
      TabIndex        =   79
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
      TabIndex        =   78
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
      TabIndex        =   21
      Top             =   3990
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
      TabIndex        =   27
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
      TabIndex        =   29
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
      TabIndex        =   25
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
      TabIndex        =   30
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
      TabIndex        =   31
      Top             =   6030
      Width           =   1470
   End
   Begin MSDataGridLib.DataGrid DGPCat 
      Height          =   4935
      Left            =   -2070
      Negotiate       =   -1  'True
      TabIndex        =   70
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
      TabIndex        =   28
      Top             =   5550
      Width           =   1470
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   9090
      TabIndex        =   65
      Top             =   7185
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   510
         TabIndex        =   66
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
   Begin VB.TextBox TxtGrid1 
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
      Left            =   1245
      TabIndex        =   19
      Top             =   5175
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSDataGridLib.DataGrid DGADItem 
      Height          =   4935
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   64
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
      Index           =   12
      Left            =   6000
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1065
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
      Index           =   13
      Left            =   2010
      MaxLength       =   10
      TabIndex        =   14
      Top             =   1785
      Width           =   2760
   End
   Begin VB.TextBox txt 
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
      Index           =   14
      Left            =   6000
      MaxLength       =   12
      TabIndex        =   15
      Top             =   1785
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
      Index           =   10
      Left            =   2010
      MaxLength       =   4
      TabIndex        =   12
      Top             =   1545
      Width           =   495
   End
   Begin VB.TextBox txt 
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
      Index           =   11
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1545
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
      Index           =   3
      Left            =   10485
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1545
      Width           =   1170
   End
   Begin VB.TextBox txt 
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
      Index           =   9
      Left            =   2010
      MaxLength       =   40
      TabIndex        =   8
      Top             =   1065
      Width           =   2775
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
      TabIndex        =   17
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
      TabIndex        =   32
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   26
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
      TabIndex        =   22
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
      Height          =   210
      Index           =   15
      Left            =   2010
      MaxLength       =   40
      TabIndex        =   16
      Top             =   2025
      Width           =   5265
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Index           =   8
      Left            =   6000
      MaxLength       =   12
      TabIndex        =   11
      Top             =   1305
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
      Index           =   7
      Left            =   2010
      MaxLength       =   20
      TabIndex        =   10
      Top             =   1305
      Width           =   2775
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
      TabIndex        =   49
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
      TabIndex        =   18
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
      FormatString    =   $"frmVehPur.frx":0000
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
      TabIndex        =   56
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   1950
      Left            =   15
      TabIndex        =   20
      Top             =   4395
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   3440
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   7
      BackColorFixed  =   12179694
      ForeColorFixed  =   128
      BackColorSel    =   15196124
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   8421631
      FocusRect       =   0
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "SrNo.|Add/Del Item |Type      |Qty|Rate  |Amount|Itemcode"
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
      _Band(0).Cols   =   7
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
      TabIndex        =   98
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
      TabIndex        =   97
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
      TabIndex        =   96
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
      Left            =   7905
      TabIndex        =   93
      Top             =   6510
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
      Left            =   9465
      TabIndex        =   92
      Top             =   6480
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Dlr Bill No. "
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
      Left            =   75
      TabIndex        =   91
      Top             =   1320
      Width           =   1515
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
      TabIndex        =   86
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
      TabIndex        =   85
      Top             =   1950
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
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
      Left            =   4935
      TabIndex        =   82
      Top             =   1320
      Width           =   540
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
      TabIndex        =   81
      Top             =   855
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date :"
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
      Left            =   4935
      TabIndex        =   80
      Top             =   1080
      Width           =   945
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
      TabIndex        =   77
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
      TabIndex        =   76
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
      TabIndex        =   75
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
      TabIndex        =   74
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
      TabIndex        =   73
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
      TabIndex        =   72
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
      TabIndex        =   71
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
      TabIndex        =   69
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
      TabIndex        =   68
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
      TabIndex        =   67
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
      TabIndex        =   63
      Top             =   2040
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excise GP. No."
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
      Left            =   75
      TabIndex        =   62
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Addition/Deduction/Shortage Item Detail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Index           =   4
      Left            =   15
      TabIndex        =   60
      Top             =   4125
      Width           =   6840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
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
      Left            =   4935
      TabIndex        =   59
      Top             =   1785
      Width           =   540
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
      Index           =   9
      Left            =   4935
      TabIndex        =   58
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RSO Purch Y/N"
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
      Left            =   75
      TabIndex        =   57
      Top             =   1560
      Width           =   1275
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
      TabIndex        =   55
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
      TabIndex        =   54
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
      TabIndex        =   53
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
      TabIndex        =   52
      Top             =   825
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purch Catg"
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
      Index           =   43
      Left            =   75
      TabIndex        =   51
      Top             =   1065
      Width           =   945
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
      TabIndex        =   50
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
      TabIndex        =   48
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
      TabIndex        =   47
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
      TabIndex        =   46
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
      TabIndex        =   45
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
      TabIndex        =   44
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
      TabIndex        =   43
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
      TabIndex        =   42
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
      TabIndex        =   41
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
      TabIndex        =   40
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
      TabIndex        =   39
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
      TabIndex        =   38
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
      TabIndex        =   37
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
      Top             =   1320
      Width           =   1350
   End
End
Attribute VB_Name = "frmVehPur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim ForeColorSelEnter$
Dim BackColorSelLeave$



Dim mReposting As Boolean
Dim mRePostCounter As Integer
Dim mSoldVehicle As Boolean

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
Private Const PurVType As String = "V_PB"
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
'Private Const Srlno As Byte = 4
Private Const Colours As Byte = 4 '
Private Const Taxable As Byte = 5 '6
Private Const Rate As Byte = 6 '7
Private Const MfgMth As Byte = 7 '8
Private Const MfgYr As Byte = 8 '9
Private Const SBookNo As Byte = 9 '10
Private Const SDM_STM_NO  As Byte = 10 '11
Private Const InDate  As Byte = 11 '12
Private Const Godown  As Byte = 12 '13
Private Const Srlno As Byte = 13 '14
Private Const LdRate  As Byte = 14
Private Const ColCode  As Byte = 15
Private Const God  As Byte = 16
Private Const ChassisDocid  As Byte = 17
Private Const OfftakeIncentiveSrlNo  As Byte = 18
Private Const OfftakeIncentive  As Byte = 19
Private Const TgtLinkIncentive  As Byte = 20
Private Const SubventionSrlNo  As Byte = 21
Private Const MfgShare As Byte = 22

' Model1,ChassisNo,EngineNo,SrlNo,Colours,Taxable,Rate,MfgMonth,MfgYr,SBookNo,SDM_STM_NO,InDate,Godown
'SrNo.0|Model 1|Chassis No 2|Engine No 3|Serial No 4|Colour 5|Tax 6|Rate 7|Mfg Month 8
'|Year 9|Service Book No 10|SDM/STM No 11|InDate 12|Godown 13|Ld. Rate 14| ColCode 15|God 16
'SrNo.0|Add/Del Item 1|Type 2|Qty 3|Rate 4|Amount 5|Itemcode 6

Private Const ADItem  As Byte = 1
Private Const ADType  As Byte = 2
Private Const Qty1  As Byte = 3
Private Const Rate1 As Byte = 4
Private Const Amt1 As Byte = 5
Private Const ADItemCode  As Byte = 6

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
        FGrid1.AddItem FGrid1.Rows
        
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
        FGrid1.AddItem FGrid1.Rows
        
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
On Error GoTo ELoop
    Dim RsTemp As ADODB.Recordset
    
    TopCtrl1.Tag = PubUParam: WinSetting Me
    
    If PubVATYN = 1 Then
        Label3(12) = "V A T @"
    End If
    Ini_Grid
    If mVType <> "V_OST" Then
        Me.BackColor = &HBBDBB3
    End If
    If mVType = PurVType Then
        '*** A/c Posting Status
        txt(AcPostByName).Visible = True
        txt(AcPostDate).Visible = True
        LblAcPostBy.Visible = True
        LblAcPostDt.Visible = True
    End If
    
    'Reminder for missing pbills
    GSQL = "SELECT Veh_Stock.ChassisNo FROM Veh_Stock where Veh_Stock.Pur_DocId = '' and Chassis_RctDocNo<>0 "
    If GCn.Execute(GSQL).RecordCount > 0 Then
        MsgBox "Purchase Bill for " & GCn.Execute(GSQL).RecordCount & " Vehicle(s) is pending", vbCritical, "Validation"
    End If
    'eof reminder
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Dim sitecond As String
    
    If mVType = PurVType Then
        sitecond = " And V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    End If
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("Veh_Purch1.DocId", "3", "1") & "='" & PubSiteCode & "'"
    End If

    
    If PubMoveRecYn Then
        Master.Open "Select DocID as Searchcode,Veh_Purch1.* from Veh_Purch1 where v_type = '" & mVType & "' and left(Docid,1)='" & PubDivCode & "' " & sitecond & " Order by V_NO desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "Select Top 1 DocID as Searchcode,Veh_Purch1.* from Veh_Purch1 where v_type = '" & mVType & "' and left(Docid,1)='" & PubDivCode & "' " & sitecond & "  Order by V_NO desc", GCn, adOpenDynamic, adLockOptimistic
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
    RsChassis.Open ("SELECT Veh_Stock.ChassisNo as code, Veh_Stock.EngineNo,Veh_Stock.Chassis_RctDocNo, Veh_Stock.INDATE, Veh_Stock.Srv_BookNo, Veh_Stock.Mfg_Month, Veh_Stock.Mfg_Yr, Veh_Stock.SDM_STM_NO, Veh_Stock.TAX_YN, Godown.God_Name, ColMast.Col_Desc, Veh_Stock.Colour_Code, Veh_Stock.Godown " & _
        "FROM (Veh_Stock LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code) " & _
        "LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code " & _
        "where (Veh_Stock.Pur_DocId='' or Veh_Stock.Pur_DocId Is Null)"), GCn, adOpenDynamic, adLockOptimistic
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
Private Sub ListView_Click()
    txtgrid1(0).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    txtgrid1(0).SetFocus
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
    mSoldVehicle = False
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    LblVPrefix.CAPTION = ""

    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
    
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
                GCn.Execute ("delete from Veh_stock where  Chassisno = '" & FGrid.TextMatrix(I, ChassisNo) & "'")
'Hiscard creation stopped by LP Singh at Udaipur 15-04-2003
'                GCn.Execute "delete from hiscard Where chassis='" & FGrid.TextMatrix(i, ChassisNo) & "'"
            Else
                MsgBox "Chassis No " & FGrid.TextMatrix(I, ChassisNo) & " is Sold" & vbCrLf & "Deletion Denied", vbInformation, "Deletion Denied"
                Exit Sub
            End If
        Next
        GCn.Execute ("delete from veh_purch1 where docId = '" & Master!DocID & "'")
        GCn.Execute ("delete from veh_purch2 where docId = '" & Master!DocID & "'")
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Model1) <> "" Then
                If Val(FGrid.TextMatrix(I, ChassisDocid)) = 0 Then
                    GCn.Execute ("delete from veh_stock where Pur_DocId='" & txt(TxtDocID) & "' and Chassis_RctDocNo =0")
                Else
                    GCn.Execute "update veh_stock set " & _
                        "Pur_DocId='',Pur_SrlNo=null,Pur_DocIDHelp='',Pur_SiteCode='',Pur_VType='',Pur_VNO=null, " & _
                        " Pur_VDate=null, PBILL_NO='',PBILL_DATE=null,PartyCode='' where ChassisNo='" & FGrid.TextMatrix(I, ChassisNo) & "' and  Chassis_RctDocNo = " & Val(FGrid.TextMatrix(I, ChassisDocid)) & ""
                End If
            End If
        Next
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
     
        mSoldVehicle = False
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, ChassisNo) <> "" Then
                GSQL = "Select Inv_Docid from Veh_Order where Chassis  = '" & FGrid.TextMatrix(I, ChassisNo) & "'"
                Debug.Print GCn.Execute(GSQL).RecordCount
                If GCn.Execute(GSQL).RecordCount > 0 Then
                    If XNull(GCn.Execute(GSQL).Fields(0)) <> "" Then
                        If Not mReposting Then MsgBox "Vehicle Sold !", vbCritical, "Edit Denied!"
                        If UCase(left(PubComp_Name, 3)) = "LMP" Or UCase(left(PubComp_Name, 4)) = "ENAR" Then
                            mSoldVehicle = True
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next
     
    If AcPostAuthorisation(txt(AcPostByName)) = False Then Exit Sub
    
    Disp_Text SETS("EDIT", Me, Master)
    If Not mSoldVehicle Then
        txt(Party).SetFocus
    Else
        txt(ExGP).SetFocus
    End If
    FGrid.AddItem FGrid.Rows
    FGrid1.AddItem FGrid1.Rows
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
On Error GoTo ELoop
Dim RstRep As ADODB.Recordset, RstRep1 As ADODB.Recordset
Dim mQry As String, I As Integer, X11


    mQry = "Select Vp.DocID,Vp.DocIDHelp,Vp.Site_Code,Vp.V_Type,Vp.V_NO,Vp.V_Date, " & _
            "Vp.PARTYCODE,Vp.PBILL_NO,Vp.PBILL_DATE,Vp.OBNO, " & _
            "Vp.OBDate,Vp.BMS_CATEGORY,Vp.RSO_WORK,Vp.RSO_Code,Vp.DueDate, " & _
            "Vp.GATE,Vp.GATEDATE,Vp.Form_Code,Vp.AMOUNT,Vp.Addition,Vp.Deduction,Vp.Exsice, " & _
            "Vp.Tax_Per,Vp.TaxSur_Per,Vp.Tax_Amt,Vp.TaxSur_Amt,Vp.Misc_Amt, " & _
            "Vp.Tot_Amount, Vp.U_Name, Vp.U_EntDt, Vp.U_AE,Vp.AcPostByU_Name,Vp.AcPostByU_EntDt,Vp.DrAcCode," & _
            "Vs.Chassis_RctDocNo ,Vs.Pur_VDate, Vs.Mfg_Month, Vs.Mfg_Yr, Vs.RSO_WORK,Vs.InDate, " & _
            "Vs.MODEL,Vs.Godown,Vs.ChassisNo,Vs.EngineNo,Vs.VehSerialNo, " & _
            "Vs.Srv_BookNo,Vs.RATE,Vs.vrate,Vs.Colour_Code,Vs.TAX_YN,Vs.SDM_STM_NO, " & _
            "Vs.PBILL_NO,Vs.PBILL_DATE,Vs.PartyCode, " & _
            "Vs.OfftakeIncentiveSrlNo,Vs.OfftakeIncentive,Vs.TgtLinkIncentive,Vs.SubventionSrlNo,Vs.MfgShare, " & _
            "Sg.Name as PartyName,Sg.Add1,Sg.Add2,Sg.Add3,Sg.LstNo,Sg.CstNo,City.CityName,TF.Form_Desc,TF.Printing_Desc,ColMast.Col_Desc,BMS.BMS_name, M.Model_Desc, Mg.ModelGrp_Name " & _
            "From (((((((Veh_Purch1 As Vp Left Join Veh_Stock as Vs On Vp.DocId = Pur_DocId) " & _
            "                         Left Join SubGroup As Sg  On Sg.SubCode   = Vp.PartyCode)    " & _
            "                         Left Join City            On Sg.CityCode  = City.CityCode)     " & _
            "                         Left Join TaxForms As TF  On Tf.Form_Code = Vp.Form_Code)      " & _
            "                         Left Join ColMast         On VS.Colour_Code = ColMast.Col_Code) Left Join BMS On Vp.BMS_CATEGORY = BMS.BMS_Code) " & _
            "                         Left Join Model M On M.Model=Vs.Model) " & _
            "                         Left join Model_Grp Mg On Mg.ModelGrp_Code=M.Grp_Code " & _
            "Where Vp.DocId='" & Master!DocID & "' "
    
        
    Set RstRep = GCn.Execute(mQry)
    
    
    mQry = "Select Vp2.DocId,Vp2.Srl_No,Vp2.Site_Code,Vp2.V_TYPE,Vp2.V_NO,Vp2.PROD_CODE,Vp2.trn_type,Vp2.QTY,Vp2.RATE," & cIIF("Vp2.trn_type='A'", "Vp2.QTY*Vp2.RATE", "-1*Vp2.QTY*Vp2.RATE") & " As Amount,veh_amdModel.Prod_name " & _
            " From Veh_Purch2 As Vp2 Left Join Veh_amdModel On veh_amdModel.Prod_Code=Vp2.Prod_Code Where Vp2.DocId = '" & Master!DocID & "'"
    
    Set RstRep1 = GCn.Execute(mQry)
    
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    X11 = CreateFieldDefFile(RstRep, PubRepoPath + "\Veh_PurchaseBill.ttx", True)
    X11 = CreateFieldDefFile(RstRep1, PubRepoPath & "\Veh_PurchaseBill1.ttx", True)
    Set rpt = rdApp.OpenReport(PubRepoPath + "\Veh_PurchaseBill.RPT")
    rpt.Database.SetDataSource RstRep
    rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("comp_name")
                rpt.FormulaFields(I).TEXT = "'" & PubComp_Name & "'"
            Case UCase("comp_add1")
                rpt.FormulaFields(I).TEXT = "'" & PubComp_Add & "'"
            Case UCase("comp_add2")
                rpt.FormulaFields(I).TEXT = "'" & PubComp_Add2 & "'"
            Case UCase("comp_city")
                rpt.FormulaFields(I).TEXT = "'" & PubComp_City & "'"
            Case UCase("title")
                rpt.FormulaFields(I).TEXT = "'" & "Vehicle Purchase Bill" & "'"
        End Select
    Next
    rpt.ReadRecords
    
    Call Report_View(rpt, Me.CAPTION, 0, True)
    
    Set RstRep = Nothing
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
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
    If txtgrid1(0).Visible = True Then
        If TxtGridLeave1 = False Then
            txtgrid1(0).SetFocus
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
    If txt(SuppInvNo) <> "" Then If IsValid(txt(SuppInvDate), "Supplier Invoice Date") = False Then Exit Sub
    If IsValid(txt(RsoYn), "RSO Purchase YN") = False Then Exit Sub
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
                If StrCmp(left(PubComp_Name, 4), "Enar") = False Then
                    If FGrid.TextMatrix(I, Colours) = "" Then MsgBox "Fill Colour in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Colours: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
                End If
                If FGrid.TextMatrix(I, Taxable) = "" Then MsgBox "Fill Taxable Yes/No in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Taxable: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
                If Val(FGrid.TextMatrix(I, Rate)) = 0 Then MsgBox "Fill Rate in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Rate: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
                If FGrid.TextMatrix(I, Godown) = "" Then MsgBox "Fill Godown in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Godown: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
            End If
        End If
    Next
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, ADItem) <> "" Then
            If FGrid1.TextMatrix(I, ADType) = "" Then MsgBox "Fill Type in Row No " & I, vbInformation, "Required data": FGrid1.Row = I: FGrid1.Col = ADType: FGrid1.SetFocus: FGrid1.CellBackColor = CellBackColEnter: Exit Sub
            If Val(FGrid1.TextMatrix(I, Qty1)) = 0 Then MsgBox "Fill Quantity in Row No " & I, vbInformation, "Required data": FGrid1.Row = I: FGrid1.Col = Qty1: FGrid1.SetFocus: FGrid1.CellBackColor = CellBackColEnter: Exit Sub
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
        If GCn.Execute("select count(*) from veh_purch1 where Left(DocID,1)='" & PubDivCode & "' And V_Type = '" & mVType & "' And V_No=" & Val(txt(SerialNo)) & "").Fields(0) > 0 Then
            If VoucherEditFlag Then 'And Txt(SerialNo).Visible Then
                MsgBox "Purchase Document No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                Exit Sub
            Else
                SetMax_VoucherPrefix "DocId", "V_PB", "Veh_Purch1", "V_Date"
                SetMax_VoucherPrefix "DocId", "V_OST", "Veh_Purch1", "V_Date"
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
'        GCn.Execute ("delete from veh_purch1 where DocID='" & txt(TxtDocID) & "'")
        GCn.Execute ("insert into Veh_Purch1( " & _
            "DocID,DocIDHelp,Site_Code,V_Type,V_NO,V_Date, " & _
            "PARTYCODE,PBILL_NO,PBILL_DATE,OBNO, " & _
            "OBDate,BMS_CATEGORY,RSO_WORK,RSO_Code,DueDate, " & _
            "GATE,GATEDATE,Form_Code,AMOUNT,Addition,Deduction,Exsice, " & _
            "Tax_Per,TaxSur_Per,Tax_Amt,TaxSur_Amt, SatPer, SatAmt,Misc_Amt, " & _
            "Tot_Amount, SubventionCredit, U_Name, U_EntDt, U_AE,AcPostByU_Name,AcPostByU_EntDt, AddBy, AddDate,DrAcCode) " & _
            "values( '" & txt(TxtDocID) & "','" & DocIdHlp & "','" & txt(SiteCode).Tag & txt(SiteCode).Tag & "','" & mVType & "'," & Val(txt(SerialNo)) & "," & ConvertDate(txt(VDate)) & _
            " ,'" & txt(Party).Tag & "','" & txt(TelcoInvNo) & "'," & ConvertDate(txt(TelcoInvDate)) & ",'" & txt(SuppInvNo) & "'," & ConvertDate(txt(SuppInvDate)) & _
            " ,'" & txt(PCat).Tag & "'," & IIf(txt(RsoYn) = "Yes", 1, 0) & ",'" & txt(RsoCode) & "'," & ConvertDate(txt(DueDate)) & _
            " ,'" & txt(ExGP) & "'," & ConvertDate(txt(ExDate)) & ",'" & txt(FormType).Tag & "','" & txt(TotGoods) & "'," & Val(txt(Addition)) & "," & Val(txt(Deduction)) & _
            " , " & Val(txt(ExAmt)) & "," & Val(txt(TaxPer)) & "," & Val(txt(TaxSurPer)) & "," & Val(txt(TaxAmt)) & "," & Val(txt(TaxSurch)) & ", " & Val(txt(SatPer)) & ", " & Val(txt(SatAmt)) & "," & Val(txt(MisCharge)) & _
            " , " & Val(txt(Gtot)) & ", " & Val(txt(SubventionCredit)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & txt(AcPostByName) & "'," & ConvertDate(txt(AcPostDate)) & ", '" & pubUName & "', " & ConvertDateTime(PubServerDate) & ",'" & rsForm!PurSal_Ac_Code & "')")
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaV, txt(TxtDocID), txt(VDate)
    Else    'edit
        CreateLog Me, Master!SearchCode, mReposting
        If Not mSoldVehicle Then
            GCn.Execute ("update veh_purch1 set V_Date=" & ConvertDate(txt(VDate)) & ", PARTYCODE='" & txt(Party).Tag & "',PBILL_NO='" & txt(TelcoInvNo) & "',PBILL_DATE=" & ConvertDate(txt(TelcoInvDate)) & ",OBNO='" & txt(SuppInvNo) & _
                "',OBDate=" & ConvertDate(txt(SuppInvDate)) & ",BMS_CATEGORY='" & txt(PCat).Tag & "',RSO_WORK=" & IIf(txt(RsoYn) = "Yes", 1, 0) & ",RSO_Code='" & txt(RsoCode) & "',DUEDATE=" & ConvertDate(txt(DueDate)) & _
                " ,GATE='" & txt(ExGP) & "',GATEDATE=" & ConvertDate(txt(ExDate)) & ",Form_Code='" & txt(FormType).Tag & "',AMOUNT=" & Val(txt(TotGoods)) & ",Addition=" & Val(txt(Addition)) & ",Deduction=" & Val(txt(Deduction)) & _
                " ,Exsice = " & Val(txt(ExAmt)) & ",TAX_Amt=" & Val(txt(TaxAmt)) & ",TaxSur_Amt=" & Val(txt(TaxSurch)) & ", SatPer = " & Val(txt(SatPer)) & ", SatAmt = " & Val(txt(SatAmt)) & ",TAX_PER=" & Val(txt(TaxPer)) & ",TaxSur_Per=" & Val(txt(TaxSurPer)) & ", MISC_AMT=" & Val(txt(MisCharge)) & _
                " ,Tot_Amount=" & Val(txt(Gtot)) & ", SubventionCredit = " & Val(txt(SubventionCredit)) & ", U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE= 'E',AcPostByU_Name='" & txt(AcPostByName) & "',AcPostByU_EntDt=" & ConvertDate(txt(AcPostDate)) & ", ModifyBy = '" & pubUName & "', ModifyDate = " & ConvertDateTime(PubServerDate) & ",DrAcCode= '" & rsForm!PurSal_Ac_Code & _
                "' where DocID='" & txt(TxtDocID) & "'")
        Else
            GCn.Execute "Update Veh_Purch1 Set Gate='" & txt(ExGP) & "', GateDate=" & ConvertDate(txt(ExDate)) & ", SubventionCredit=" & Val(txt(SubventionCredit)) & " Where  DocID='" & txt(TxtDocID) & "'"
        End If
    End If
        If Not mSoldVehicle Then
            GCn.Execute ("Delete from Veh_Stock where Pur_DocId='" & txt(TxtDocID) & "' and (Chassis_RctDocNo =0 or Chassis_RctDocNo iS Null)")
            For I = 1 To FGrid.Rows - 1
                If FGrid.TextMatrix(I, Model1) <> "" Then
                   If Val(FGrid.TextMatrix(I, ChassisDocid)) = 0 Then
                        GCn.Execute ("insert into veh_stock " & _
                            "(Pur_DocId,Pur_SrlNo,Pur_DocIDHelp,Pur_SiteCode,Pur_VType,Pur_VNO, " & _
                            "Chassis_RctDocNo ,Pur_VDate, Mfg_Month, Mfg_Yr, RSO_WORK,InDate, " & _
                            "MODEL,Godown,ChassisNo,EngineNo,VehSerialNo, " & _
                            "Srv_BookNo,RATE,vrate,Colour_Code,TAX_YN,SDM_STM_NO, " & _
                            "PBILL_NO,PBILL_DATE,PartyCode, U_Name, U_EntDt,U_AE, " & _
                            "OfftakeIncentiveSrlNo,OfftakeIncentive,TgtLinkIncentive,SubventionSrlNo,MfgShare) " & _
                            "values('" & txt(TxtDocID).TEXT & "'," & I & ",'" & DocIdHlp & "','" & PubSiteCode & txt(SiteCode).Tag & "','" & mVType & "'," & Val(txt(SerialNo).TEXT) & ", " & _
                            "" & Val(FGrid.TextMatrix(I, ChassisDocid)) & "," & ConvertDate(txt(VDate).TEXT) & ",'" & FGrid.TextMatrix(I, MfgMth) & "','" & FGrid.TextMatrix(I, MfgYr) & "'," & IIf(txt(RsoYn).TEXT = "Yes", 1, 0) & "," & ConvertDate(FGrid.TextMatrix(I, InDate)) & ", " & _
                            "'" & FGrid.TextMatrix(I, Model1) & "','" & FGrid.TextMatrix(I, God) & "','" & FGrid.TextMatrix(I, ChassisNo) & "','" & FGrid.TextMatrix(I, EngineNo) & "','" & FGrid.TextMatrix(I, Srlno) & "' , " & _
                            "'" & FGrid.TextMatrix(I, SBookNo) & "'," & Val(FGrid.TextMatrix(I, Rate)) & "," & Val(FGrid.TextMatrix(I, LdRate)) & ",'" & FGrid.TextMatrix(I, ColCode) & "'," & IIf(FGrid.TextMatrix(I, Taxable) = "Yes", 1, 0) & ",'" & FGrid.TextMatrix(I, SDM_STM_NO) & "', " & _
                            "'" & txt(TelcoInvNo).TEXT & "'," & ConvertDate(txt(TelcoInvDate).TEXT) & ",'" & txt(Party).Tag & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'E', " & _
                            "'" & FGrid.TextMatrix(I, OfftakeIncentiveSrlNo) & "'," & Val(FGrid.TextMatrix(I, OfftakeIncentive)) & _
                            ", " & Val(FGrid.TextMatrix(I, TgtLinkIncentive)) & ",'" & FGrid.TextMatrix(I, SubventionSrlNo) & _
                            "', " & Val(FGrid.TextMatrix(I, MfgShare)) & ")")
                            
                            If GCn.Execute("Select * from hiscard Where chassis='" & FGrid.TextMatrix(I, ChassisNo) & "'").RecordCount = 0 Then
                                CardNo = PubSiteCode + STR(GCn.Execute("select " & vIsNull("max(" & cVal(cMID("cardno", "3", "len(cardno)-2")) & ")", "0") & " + 1 from hiscard").Fields(0))
            
                                Dim urs As Recordset
                                Set urs = GCn.Execute("select Max(" & cVal(cMID("CardNo", "3", "6")) & "), Max(CardDate) from HisCard")
                                CardNo = PubSiteCode + PubDivCode + Right("000000" & IIf(IsNull((urs.Fields(0).Value)), 1, (urs.Fields(0).Value) + 1), 6)
                                Set urs = Nothing
            
                                
                                GCn.Execute "insert into hiscard(cardno,Site_Code,Div_Code,carddate,Name,model,chassis,engine,U_Name, U_EntDt, U_AE) " & _
                                "values('" & CardNo & "','" & PubSiteCode & "','" & PubDivCode & "'," & ConvertDate(txt(VDate)) & ",'" & PubComp_Name & "','" & FGrid.TextMatrix(I, Model1) & "','" & FGrid.TextMatrix(I, ChassisNo) & "','" & FGrid.TextMatrix(I, EngineNo) & "', " & _
                                "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
                            End If
                   Else
                        GCn.Execute "update veh_stock set " & _
                            "  Pur_DocId='" & txt(TxtDocID).TEXT & "',Pur_SrlNo=" & I & ",Pur_DocIDHelp='" & DocIdHlp & "',Pur_SiteCode='" & PubSiteCode & txt(SiteCode).Tag & "',Pur_VType='" & mVType & "',Pur_VNO=" & Val(txt(SerialNo).TEXT) & _
                            ", Pur_VDate=" & ConvertDate(txt(VDate).TEXT) & ", Mfg_Month='" & FGrid.TextMatrix(I, MfgMth) & "', Mfg_Yr='" & FGrid.TextMatrix(I, MfgYr) & "', RSO_WORK=" & IIf(txt(RsoYn).TEXT = "Yes", 1, 0) & ", InDate=" & ConvertDate(FGrid.TextMatrix(I, InDate)) & _
                            ", MODEL='" & FGrid.TextMatrix(I, Model1) & "',Godown='" & FGrid.TextMatrix(I, God) & "',EngineNo='" & FGrid.TextMatrix(I, EngineNo) & "',VehSerialNo='" & FGrid.TextMatrix(I, Srlno) & _
                            "',Srv_BookNo='" & FGrid.TextMatrix(I, SBookNo) & "',RATE=" & Val(FGrid.TextMatrix(I, Rate)) & ",vrate=" & Val(FGrid.TextMatrix(I, LdRate)) & ",Colour_Code='" & FGrid.TextMatrix(I, ColCode) & "',TAX_YN=" & IIf(FGrid.TextMatrix(I, Taxable) = "Yes", 1, 0) & ",SDM_STM_NO='" & FGrid.TextMatrix(I, SDM_STM_NO) & _
                            "',PBILL_NO='" & txt(TelcoInvNo).TEXT & "',PBILL_DATE=" & ConvertDate(txt(TelcoInvDate).TEXT) & ",PartyCode='" & txt(Party).Tag & "', U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E', OfftakeIncentiveSrlNo= '" & FGrid.TextMatrix(I, OfftakeIncentiveSrlNo) & "',OfftakeIncentive=" & Val(FGrid.TextMatrix(I, OfftakeIncentive)) & _
                            ",TgtLinkIncentive=" & Val(FGrid.TextMatrix(I, TgtLinkIncentive)) & ",SubventionSrlNo='" & FGrid.TextMatrix(I, SubventionSrlNo) & _
                            "',MfgShare= " & Val(FGrid.TextMatrix(I, MfgShare)) & _
                            " where ChassisNo='" & FGrid.TextMatrix(I, ChassisNo) & "' and  Chassis_RctDocNo = " & Val(FGrid.TextMatrix(I, ChassisDocid)) & ""
                   End If
                End If
            Next
    
    
            GCn.Execute ("delete from veh_purch2 where DocId='" & txt(TxtDocID) & "'")
            For I = 1 To FGrid1.Rows - 1
                If FGrid1.TextMatrix(I, ADItem) <> "" And Val(FGrid1.TextMatrix(I, Qty1)) <> 0 Then
                    GCn.Execute ("insert into veh_purch2(DocId,Srl_No,Site_Code,V_TYPE,V_NO,PROD_CODE,trn_type,QTY,RATE, U_Name, U_EntDt, U_AE) " & _
                        "values('" & txt(TxtDocID).TEXT & "'," & I & ",'" & PubSiteCode & txt(SiteCode).Tag & "','" & mVType & "','" & txt(SerialNo).TEXT & "', " & _
                        "'" & FGrid1.TextMatrix(I, ADItemCode) & "','" & left(FGrid1.TextMatrix(I, ADType), 1) & "'," & Val(FGrid1.TextMatrix(I, Qty1)) & "," & Val(FGrid1.TextMatrix(I, Rate1)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'E')")
                End If
            Next
    Else
        For I = 1 To FGrid.Rows - 1
            GCn.Execute "update veh_stock set Godown='" & FGrid.TextMatrix(I, God) & "' where ChassisNo='" & FGrid.TextMatrix(I, ChassisNo) & "' and  Chassis_RctDocNo = " & Val(FGrid.TextMatrix(I, ChassisDocid)) & ""
        Next I
    End If
    
    'A/c Posting
    If mVType = PurVType Then
        If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
            If mRePostCounter = 0 Then ProcAcPost
        End If
    End If
    
    mSoldVehicle = False
    'EOF of A/c Posting Section
    If mVType = PurVType Then GCnFaV.CommitTrans
    GCn.CommitTrans
    mTrans = False
    Set Rst = Nothing
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select DocID as Searchcode,Veh_Purch1.* from Veh_Purch1 where v_type = '" & mVType & "' and left(Docid,1)='" & PubDivCode & "' And DocId ='" & txt(TxtDocID) & "' Order by V_NO desc")
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
    
    Dim sitecond As String
    
    If mVType = PurVType Then
        sitecond = " And V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    Else
        sitecond = " And V_Date <= " & ConvertDate(DateAdd("D", -1, PubStartDate)) & " "
    End If
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
        sitecond = sitecond & " and " & cMID("vp1.DocId", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    GSQL = "select VP1.DocID as searchcode, " & cCStr("VP1.V_NO") & " As V_No," & cDt("VP1.V_Date") & " As V_Date,VP1.PBILL_NO as Mfg_BNo," & cDt("VP1.PBILL_DATE") & " as Mfg_BDate,VStk.MODEL,VStk.ChassisNo,VStk.EngineNo,SG.Name " & _
    " from (Veh_Purch1 VP1 left join Veh_Stock VStk on VP1.DocID=VSTK.Pur_DocID) " & _
    " left join SubGroup SG on VP1.PartyCode=SG.SubCode " & _
    " where v_type = '" & mVType & "' and left(Docid,1)='" & PubDivCode & "' " & sitecond & " order by V_Date desc"
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
        Set Master = GCn.Execute("Select DocID as Searchcode,Veh_Purch1.* from Veh_Purch1 where v_type = '" & mVType & "' and left(Docid,1)='" & PubDivCode & "' And DocId ='" & MyValue & "' Order by V_NO desc")
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
    Case MisCharge, TaxAmt, TaxSurch, SatAmt
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
            If GCn.Execute("Select PBILL_NO from Veh_Purch1 where PBILL_NO='" & txt(TelcoInvNo) & "'").RecordCount > 0 Then
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
                Rst.Open "Select * From veh_purch1 Where docid='" & DocID & "'", GCn, adOpenDynamic, adLockOptimistic
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
Private Sub DGCol_Click()
    DGCol.Visible = False
    If RsCol.RecordCount > 0 Then
        TxtGrid(0).TEXT = rsGod!Name
         FGrid.TextMatrix(FGrid.Row, Colours) = RsCol!Name
         FGrid.TextMatrix(FGrid.Row, ColCode) = RsCol!Code
    End If
   TxtGrid(0).SetFocus
End Sub
Private Sub DGPCat_Click()
    DGPCat.Visible = False
    If RsPCat.RecordCount > 0 Then
        txt(PCat).TEXT = RsPCat!Name
        txt(PCat).Tag = RsPCat!Code
    End If
    txt(PCat).SetFocus
End Sub

Private Sub DGADItem_Click()
    DGADItem.Visible = False
    If RsADItem.RecordCount > 0 Then
        txtgrid1(0).TEXT = RsADItem!Name
         FGrid1.TextMatrix(FGrid1.Row, ADItem) = RsADItem!Name
         FGrid1.TextMatrix(FGrid1.Row, ADItemCode) = RsADItem!Code
    End If
   txtgrid1(0).SetFocus
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
        Case Model1, ChassisNo, EngineNo, Srlno, Taxable, MfgMth, MfgYr, SBookNo, SDM_STM_NO, InDate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        Case Godown
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, God) = ""
        Case Colours
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            FGrid.TextMatrix(FGrid.Row, ColCode) = ""
        Case Rate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "0.00"
    End Select
    Amt_Cal False
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case Model1, ChassisNo, EngineNo, Srlno, Colours, Taxable, Rate, MfgMth, MfgYr, SBookNo, SDM_STM_NO, InDate, Godown
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
            If Not mSoldVehicle Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
                TxtGrid(0).SelStart = 6
            End If
        Case Model1, EngineNo, Srlno, Colours, MfgMth, SBookNo, SDM_STM_NO, InDate
            If Not mSoldVehicle Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
            End If
        Case Taxable
            If Not mSoldVehicle Then
                If UCase(Chr(KeyAscii)) = "N" Then
                    FGrid.TextMatrix(FGrid.Row, Taxable) = "No"
                ElseIf UCase(Chr(KeyAscii)) = "Y" Then
                    FGrid.TextMatrix(FGrid.Row, Taxable) = "Yes"
                Else
                    FGrid.TextMatrix(FGrid.Row, Taxable) = ""
                End If
                KeyAscii = 0
            End If
        Case Rate, MfgYr
            If Not mSoldVehicle Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
            End If
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
On Error GoTo error1

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

    txt(PCat).Tag = Master!BMS_CATEGORY
    If txt(PCat).Tag <> "" Then
        txt(PCat).TEXT = GCn.Execute("SELECT  BMS_name  FROM BMS  where BMS_Code = '" & txt(PCat).Tag & "'").Fields(0).Value
    Else
        txt(PCat).TEXT = ""
    End If
    
    txt(FormType).Tag = IIf(IsNull(Master!Form_Code), "", Master!Form_Code)
    If txt(FormType).Tag <> "" Then
        txt(FormType).TEXT = GCn.Execute("select form_desc from taxforms where form_code = '" & txt(FormType).Tag & "'").Fields(0).Value
    Else
        txt(FormType).TEXT = ""
    End If

    txt(SuppInvNo).TEXT = IIf(IsNull(Master!OBNO), "", Master!OBNO)
    txt(SuppInvDate).TEXT = IIf(IsNull(Master!OBDate), "", Master!OBDate)
    txt(TelcoInvNo).TEXT = IIf(IsNull(Master!PBILL_NO), "", Master!PBILL_NO)
    txt(TelcoInvDate).TEXT = IIf(IsNull(Master!PBILL_DATE), "", Master!PBILL_DATE)
    txt(RsoYn).TEXT = IIf(Master!RSO_WORK = 1, "Yes", "No")
    txt(RsoCode).TEXT = IIf(IsNull(Master!RSO_Code), "", Master!RSO_Code)
    txt(DueDate).TEXT = IIf(IsNull(Master!DueDate), "", Master!DueDate)
    txt(ExGP).TEXT = IIf(IsNull(Master!GATE), "", Master!GATE)
    txt(ExDate).TEXT = IIf(IsNull(Master!GATEDATE), "", Master!GATEDATE)

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
    txt(SubventionCredit).TEXT = Format(IIf(IsNull(Master!SubventionCredit) Or Master!SubventionCredit = 0, "", Master!SubventionCredit), "0.00")
    '*** A/c Posting Status
    txt(AcPostByName) = IIf(IsNull(Master!AcPostByU_Name), "", Master!AcPostByU_Name)
    txt(AcPostDate) = IIf(IsNull(Master!AcPostByU_EntDt), "", Master!AcPostByU_EntDt)
    '***
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT Godown.God_Name, ColMast.Col_Desc,  Veh_Stock.pur_SrlNo,Veh_Stock.Chassis_RctDocNo,  Veh_Stock.Mfg_Month, Veh_Stock.Mfg_Yr, Veh_Stock.INDATE, Veh_Stock.MODEL, Veh_Stock.ChassisNo, Veh_Stock.EngineNo, Veh_Stock.VehSerialNo, Veh_Stock.Srv_BookNo, Veh_Stock.RATE, Veh_Stock.VRATE, Veh_Stock.TAX_YN,Veh_Stock.sdm_stm_no ,Veh_Stock.godown,Veh_Stock.Colour_Code, " & _
            " OfftakeIncentiveSrlNo,OfftakeIncentive,TgtLinkIncentive,SubventionSrlNo,MfgShare " & _
            " FROM (Veh_Stock LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code " & _
            " where Veh_Stock.Pur_DocId = '" & Master!DocID & "'")
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
                .TextMatrix(I, Colours) = IIf(IsNull(Rs!Col_Desc), "", Rs!Col_Desc)
                .TextMatrix(I, Taxable) = IIf(IsNull(Rs!Tax_YN) Or Rs!Tax_YN = 0, "No", "Yes")
                .TextMatrix(I, Rate) = Format(IIf(IsNull(Rs!Rate), "", Rs!Rate), "0.00")
                .TextMatrix(I, MfgMth) = IIf(IsNull(Rs!Mfg_Month), "", Rs!Mfg_Month)
                .TextMatrix(I, MfgYr) = IIf(IsNull(Rs!Mfg_Yr), "", Rs!Mfg_Yr)
                .TextMatrix(I, Srlno) = IIf(IsNull(Rs!VehSerialNo), "", Rs!VehSerialNo)
                .TextMatrix(I, SBookNo) = IIf(IsNull(Rs!Srv_BookNo), "", Rs!Srv_BookNo)
                .TextMatrix(I, SDM_STM_NO) = IIf(IsNull(Rs!SDM_STM_NO), "", Rs!SDM_STM_NO)
                .TextMatrix(I, InDate) = IIf(IsNull(Rs!InDate), "", Rs!InDate)
                .TextMatrix(I, Godown) = IIf(IsNull(Rs!God_Name), "", Rs!God_Name)
                .TextMatrix(I, LdRate) = Format(IIf(IsNull(Rs!vrate), "", Rs!vrate), "0.00")
                .TextMatrix(I, ColCode) = IIf(IsNull(Rs!Colour_Code), "", Rs!Colour_Code)
                .TextMatrix(I, God) = IIf(IsNull(Rs!Godown), "", Rs!Godown)
                .TextMatrix(I, ChassisDocid) = IIf(IsNull(Rs!Chassis_RctDocNo), "", Rs!Chassis_RctDocNo)
                .TextMatrix(I, OfftakeIncentiveSrlNo) = IIf(IsNull(Rs!OfftakeIncentiveSrlNo), "", Rs!OfftakeIncentiveSrlNo)
                .TextMatrix(I, OfftakeIncentive) = IIf(IsNull(Rs!OfftakeIncentive), "", Rs!OfftakeIncentive)
                .TextMatrix(I, TgtLinkIncentive) = IIf(IsNull(Rs!TgtLinkIncentive), "", Rs!TgtLinkIncentive)
                .TextMatrix(I, SubventionSrlNo) = IIf(IsNull(Rs!SubventionSrlNo), "", Rs!SubventionSrlNo)
                .TextMatrix(I, MfgShare) = IIf(IsNull(Rs!MfgShare), "", Rs!MfgShare)
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
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT Veh_AMDModel.Prod_Name, Veh_Purch2.Srl_No, Veh_Purch2.PROD_CODE, Veh_Purch2.QTY, Veh_Purch2.RATE, Veh_Purch2.Trn_Type " & _
        "FROM Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
        "where Veh_purch2.DocId = '" & Master!DocID & "' ")
    FGrid1.Rows = 1: FGrid1.Redraw = False
    I = 1
    If Rs.RecordCount > 0 Then
        Do Until Rs.EOF
            With FGrid1
                .AddItem ""
                .TextMatrix(I, 0) = Rs!Srl_No
                .TextMatrix(I, ADItem) = XNull(Rs!Prod_Name)
                .TextMatrix(I, ADType) = IIf(Rs!Trn_Type = "A", "Addition", IIf(Rs!Trn_Type = "D", "Deletion", "Shortage"))
                .TextMatrix(I, Qty1) = Format(IIf(IsNull(Rs!Qty), "", Rs!Qty), "0")
                .TextMatrix(I, Rate1) = Format(IIf(IsNull(Rs!Rate), "", Rs!Rate), "0.00")
                .TextMatrix(I, Amt1) = Format(.TextMatrix(I, Qty1) * .TextMatrix(I, Rate1), "0.00")
                .TextMatrix(I, ADItemCode) = Rs!Prod_Code
            End With
            Rs.MoveNext
           I = I + 1
        Loop
        FGrid1.FixedRows = 1
    Else
        FGrid1.AddItem FGrid1.Rows
        FGrid1.FixedRows = 1
    End If
    FGrid1.Redraw = True
    Set Rs = Nothing
    
    
    
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, ChassisNo) <> "" Then
                GSQL = "Select Inv_Docid from Veh_Order where Chassis  = '" & FGrid.TextMatrix(I, ChassisNo) & "'"
                Debug.Print GCn.Execute(GSQL).RecordCount
                If GCn.Execute(GSQL).RecordCount > 0 Then
                    If XNull(GCn.Execute(GSQL).Fields(0)) <> "" Then
                        mSoldVehicle = True
                    End If
                End If
            End If
        Next
    
Else
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
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
        .Cols = 23
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
        
        .TextMatrix(0, Srlno) = "Srl. NO"
        .ColAlignment(Srlno) = flexAlignLeftCenter
        .ColWidth(Srlno) = 1800
        
        .TextMatrix(0, Colours) = "Colours"
        .ColAlignment(Colours) = flexAlignLeftCenter
        .ColWidth(Colours) = 1500
        
        .TextMatrix(0, Taxable) = "TaxOnSale"
        .ColAlignment(Taxable) = flexAlignLeftCenter
        .ColWidth(Taxable) = 885
        
        .TextMatrix(0, Rate) = "Rate"
        .ColAlignmentFixed(Rate) = flexAlignRightCenter
        .ColWidth(Rate) = 1100
      
        .TextMatrix(0, MfgMth) = "Mfg Month"
        .ColAlignment(MfgMth) = flexAlignLeftCenter
        .ColWidth(MfgMth) = 900
        
        .TextMatrix(0, MfgYr) = "Year"
        .ColAlignment(MfgYr) = flexAlignLeftCenter
        .ColWidth(MfgYr) = 570
        
        .TextMatrix(0, SBookNo) = "SrvBookNo"
        .ColAlignment(SBookNo) = flexAlignLeftCenter
        .ColWidth(SBookNo) = 975
        
        .TextMatrix(0, InDate) = "In Date"
        .ColAlignment(InDate) = flexAlignLeftCenter
        .ColWidth(InDate) = 1200
        
        .TextMatrix(0, SDM_STM_NO) = "SDM/STM No"
        .ColAlignment(SDM_STM_NO) = flexAlignLeftCenter
        .ColWidth(SDM_STM_NO) = 1200
        
        .TextMatrix(0, Godown) = "Godown"
        .ColAlignment(Godown) = flexAlignLeftCenter
        .ColWidth(Godown) = 1200
        
        .TextMatrix(0, LdRate) = "VRate"
        .ColAlignmentFixed(LdRate) = flexAlignRightCenter
        .ColWidth(LdRate) = 1100
        
        .ColWidth(ColCode) = 0
        .ColWidth(God) = 0
        .ColWidth(ChassisDocid) = 0
        .ColWidth(OfftakeIncentiveSrlNo) = 0
        .ColWidth(OfftakeIncentive) = 0
        .ColWidth(TgtLinkIncentive) = 0
        .ColWidth(SubventionSrlNo) = 0
        .ColWidth(MfgShare) = 0
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    
    With FGrid1
        .left = Me.left
        .width = 7410
        .top = 4620
        .RowHeightMin = PubGridRowHeight

        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, ADItem) = "Add Del Item"
        .ColAlignment(ADItem) = flexAlignLeftCenter
        .ColWidth(ADItem) = 2400
        
        .TextMatrix(0, ADType) = "Type"
        .ColAlignment(ADType) = flexAlignLeftCenter
        .ColWidth(ADType) = 1200
       
        .TextMatrix(0, Qty1) = "Qty"
        .ColAlignmentFixed(Qty1) = flexAlignRightCenter
        .ColWidth(Qty1) = 645

        .TextMatrix(0, Rate1) = "Rate"
        .ColAlignmentFixed(Rate1) = flexAlignRightCenter
        .ColWidth(Rate1) = 855
        
        .TextMatrix(0, Amt1) = "Amount"
        .ColAlignmentFixed(Amt1) = flexAlignRightCenter
        .ColWidth(Amt1) = 1065
        
        .ColWidth(ADItemCode) = 0
    End With
    Label3(4).left = FGrid1.left: Label3(4).top = FGrid1.top - (Label3(4).height + 15): Label3(4).width = FGrid1.width
    
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

If mSoldVehicle Then Enb = False

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
txt(Addition).Enabled = False
txt(Deduction).Enabled = False
txt(SubAmt).Enabled = False

txtDisabled_Color Me

TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol

txtgrid1(0).BackColor = CtrlBCol
txtgrid1(0).ForeColor = CtrlFCol

If PubSiebelActiveYn = 1 And pubUName = "SA" Then
    cmdPost.Visible = True
Else
    cmdPost.Visible = False
End If
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
            If RsChassis.State <> 0 Then RsChassis.Close
            If FGrid.TextMatrix(FGrid.Row, Model1) = "" Then MsgBox "Select Model First", vbInformation, "Validation": FGrid.Col = Model1: TxtGrid(0).Visible = False: FGrid.SetFocus: Exit Sub
'                Set RsChassis = New ADODB.Recordset
'                RsChassis.CursorLocation = adUseClient
            RsChassis.Open ("SELECT Veh_Stock.ChassisNo as code, Veh_Stock.EngineNo,Veh_Stock.Chassis_RctDocNo, Veh_Stock.INDATE, Veh_Stock.Srv_BookNo, Veh_Stock.Mfg_Month, Veh_Stock.Mfg_Yr, Veh_Stock.SDM_STM_NO, Veh_Stock.TAX_YN, Godown.God_Name, ColMast.Col_Desc, Veh_Stock.Colour_Code, Veh_Stock.Godown " & _
                "FROM (Veh_Stock LEFT JOIN Godown ON Veh_Stock.Godown = Godown.God_Code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code " & _
                "where Veh_Stock.MODEL  = '" & FGrid.TextMatrix(FGrid.Row, Model1) & "' and (Veh_Stock.Pur_DocId='' or Veh_Stock.Pur_DocId Is Null)"), GCn, adOpenDynamic, adLockOptimistic
            Set DgChassis.DataSource = RsChassis
'            If RsChassis.RecordCount = 0 Or (RsChassis.EOF = True Or RsChassis.BOF = True) Or FGrid.TextMatrix(FGrid.Row, ChassisNo) = "" Then Exit Sub
'            If FGrid.TextMatrix(FGrid.Row, ChassisNo) <> RsChassis!code Then
'                RsChassis.MoveFirst
'                RsChassis.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "'"
'            End If
        Case Godown
            If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Godown) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Godown) <> rsGod!Name Then
                rsGod.MoveFirst
                rsGod.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, Godown) & "'"
            End If
         Case Colours
            TxtGrid(0).MaxLength = 15
            If RsCol.RecordCount = 0 Or (RsCol.EOF = True Or RsCol.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Colours) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Colours) <> RsCol!Name Then
                RsCol.MoveFirst
                RsCol.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, Colours) & "'"
            End If
         Case Rate, MfgYr
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
    
        DGridTxtKeyDown_Mast DgChassis, TxtGrid, Index, RsChassis, KeyCode, True, 0
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Srlno
            End If
        End If
    Case Model1    '1
        DGridTxtKeyDown DGMod, TxtGrid, Index, RsMod, KeyCode, True, 0, frmModel, "frmModel"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Srlno
            End If
        End If
    Case Godown
        DGridTxtKeyDown DGGod, TxtGrid, 0, rsGod, KeyCode, True, 1, frmGodown, "frmGodown"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Srlno
            End If
        End If
    Case Colours
        DGridTxtKeyDown DGCol, TxtGrid, 0, RsCol, KeyCode, True, 1, frmColor, "frmColor"
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
               GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Srlno
            End If
        End If
    Case EngineNo, Taxable, Srlno, Rate, MfgMth, MfgYr, SBookNo, SDM_STM_NO, InDate
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Srlno
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
    Case Model1
        If DGMod.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsMod, KeyAscii, "Code"
    Case Godown
        If DGGod.Visible = True Then DGridTxtKeyPress TxtGrid, 0, rsGod, KeyAscii, "Name"
    Case Colours
        If DGCol.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsCol, KeyAscii, "name"
    Case Taxable
        If UCase(Chr(KeyAscii)) = "Y" Then
            TxtGrid(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            TxtGrid(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            TxtGrid(Index) = ""
        End If
        KeyAscii = 0
    Case Rate
       Call NumPress(TxtGrid(0), KeyAscii, 8, 2)
    Case MfgYr
       Call NumPress(TxtGrid(0), KeyAscii, 4, 0)
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
                DGridTxtKeyUp_Mast TxtGrid, Index, RsChassis, KeyCode, "code"
            Case Godown
                If KeyCode <> 13 And DGGod.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, 0, rsGod, KeyCode, "Name", True
            Case Colours
                If KeyCode <> 13 And DGCol.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsCol, KeyCode, "name", True
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
                    FGrid.TextMatrix(FGrid.Row, Taxable) = IIf(txt(RsoYn) = "Yes", "No", "Yes")
                
                    If txt(TelcoInvDate) <> "" Then
                        Set Rst = New ADODB.Recordset
                        Rst.CursorLocation = adUseClient
                        Rst.Open "Select top 1 TAXABLE_YN,P_RATE From Veh_Rate Where model='" & FGrid.TextMatrix(FGrid.Row, Model1) & "' And Effective_Date<=" & ConvertDate(txt(TelcoInvDate)) & " Order by Effective_Date Desc", GCn, adOpenDynamic, adLockOptimistic
                        If Rst.RecordCount > 0 Then
                            FGrid.TextMatrix(FGrid.Row, Rate) = Format(IIf(IsNull(Rst!p_rate), 0, Rst!p_rate), "0.00")
                            FGrid.TextMatrix(FGrid.Row, Taxable) = IIf(Rst!TAXABLE_YN = 1, "Yes", "No")
                        Else
                            FGrid.TextMatrix(FGrid.Row, Rate) = ""
                            FGrid.TextMatrix(FGrid.Row, Taxable) = "No"
                        End If
                    End If
                End If
                
                Amt_Cal False
            End If
            If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
        Case ChassisNo
            If ChkDul_Chassis = True Then TxtGridLeave = False: Exit Function
'Modi Shekhar
            If FGrid.TextMatrix(FGrid.Row, ChassisNo) <> "" Then
                If GCn.Execute("select count(*) From Veh_order where Chassis = '" & FGrid.TextMatrix(FGrid.Row, ChassisNo) & "'").Fields(0).Value > 0 Then
                  MsgBox "Chassis Sold" & vbCrLf & "Editing Denied", vbInformation, "Editing Denied": FGrid.SetFocus: TxtGridLeave = False: Exit Function
                End If
            End If
            If FGrid.TextMatrix(FGrid.Row, ChassisNo) <> "" And TopCtrl1.TopText2 <> "Edit" Then
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
            If FGrid.TextMatrix(FGrid.Row, InDate) = "" Then
                FGrid.TextMatrix(FGrid.Row, InDate) = txt(VDate)
            End If
            If FGrid.TextMatrix(FGrid.Row, ChassisNo) <> "" Then
                If FGrid.TextMatrix(FGrid.Row, MfgMth) = "" Or FGrid.TextMatrix(FGrid.Row, MfgYr) = "" Then
                    FGrid.TextMatrix(FGrid.Row, MfgMth) = DeCodeChassis(FGrid.TextMatrix(FGrid.Row, ChassisNo), MfgMonth)
                    FGrid.TextMatrix(FGrid.Row, MfgYr) = DeCodeChassis(FGrid.TextMatrix(FGrid.Row, ChassisNo), MfgYear)
                End If
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
            If FGrid <> "" Then
                Set RsTemp = GCn.Execute("Select Count(*) From Veh_Stock Where EngineNo = '" & FGrid & "' And Pur_DocId <> '" & DocID & "' ")
                If RsTemp(0) > 0 Then
                    MsgBox "Engine No Already Exist!"
                    TxtGridLeave = False
                    Exit Function
                End If
            End If
        Case Godown
            If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, Godown) = ""
                FGrid.TextMatrix(FGrid.Row, God) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
                FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
            End If
        Case Colours
            If RsCol.RecordCount = 0 Or (RsCol.EOF = True Or RsCol.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, Colours) = ""
                FGrid.TextMatrix(FGrid.Row, ColCode) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, Colours) = RsCol!Name
                FGrid.TextMatrix(FGrid.Row, ColCode) = RsCol!Code
            End If
         Case Rate
                FGrid.TextMatrix(FGrid.Row, Rate) = Format(Val(TxtGrid(0).TEXT), "0.00")
                Amt_Cal False
         Case InDate
                TxtGrid(0).TEXT = RetDate(TxtGrid(0))
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
         Case MfgMth
                TxtGrid(0).TEXT = RetMonth(TxtGrid(0))
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
         Case Srlno, Taxable, MfgYr, SBookNo, SDM_STM_NO
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
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

Private Sub TxtGrid1_GotFocus(Index As Integer)
    txtgrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    Select Case FGrid1.Col
        Case ADType
            ListArray = Array("Addition", "Deletion", "Shortage")
            Set mListItem = ListView_Items(ListView, txtgrid1, 0, ListArray, 3)
         Case ADItem
            If RsADItem.RecordCount = 0 Or (RsADItem.EOF = True Or RsADItem.BOF = True) Or FGrid1.TextMatrix(FGrid1.Row, ADItemCode) = "" Then Exit Sub
            If FGrid1.TextMatrix(FGrid1.Row, ADItem) <> RsADItem!Code Then
                RsADItem.MoveFirst
                RsADItem.FIND "code ='" & FGrid1.TextMatrix(FGrid1.Row, ADItemCode) & "'"
            End If
         Case Rate1, Qty1
'                         SendKeys "{HOME}+{END}"
     End Select
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyEscape Then
                txtgrid1(0).TEXT = txtgrid1(0).Tag
                TxtGrid1_KeyUp Index, KeyCode, Shift
                txtgrid1(0).Visible = False
                Grid_Hide
                FGrid1.SetFocus
                Exit Sub
            End If
            Select Case FGrid1.Col
                Case ADType
                    ListView_KeyDown FrmList, ListView, txtgrid1, 0, KeyCode, Shift, txtgrid1(0).left, (txtgrid1(0).top + txtgrid1(0).height + 25), txtgrid1(0).width, 900
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave1 = True Then
                             GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, Rate1
                        End If
                    End If
                Case ADItem    '1
                    DGridTxtKeyDown DGADItem, txtgrid1, Index, RsADItem, KeyCode, True, 1, frmVehAMDMast, "frmVehAMDMast"
                    If KeyCode = vbKeyReturn Then
                            If TxtGridLeave1 = True Then
                                 GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, Rate1
                            End If
                    End If

                Case Qty1, Rate1
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave1 = True Then
                             GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, Rate1
                        End If
                End If
                End Select
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
Select Case FGrid1.Col
    Case ADItem
        If DGADItem.Visible = True Then DGridTxtKeyPress txtgrid1, Index, RsADItem, KeyAscii, "name"
    Case Rate1
        Call NumPress(txtgrid1(Index), KeyAscii, 8, 2)
    Case Qty1
        Call NumPress(txtgrid1(Index), KeyAscii, 6, 0)
End Select
End Sub


Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
        Select Case FGrid1.Col
            Case ADType
                If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0
                ListView_KeyUp ListView, txtgrid1, 0, KeyCode, mListItem
            Case ADItem
                If KeyCode <> 13 And DGADItem.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0: DGridTxtKeyPress txtgrid1, Index, RsADItem, KeyCode, "name", True
            Case Qty1
                FGrid1.TextMatrix(FGrid1.Row, Qty1) = Format(Val(txtgrid1(Index).TEXT), "0.000")
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = Format((Val(FGrid1.TextMatrix(FGrid1.Row, Rate1)) * Val(FGrid1.TextMatrix(FGrid1.Row, Qty1))), "0.00")
                Amt_Cal False
            Case Rate1
                FGrid1.TextMatrix(FGrid1.Row, Rate1) = Format(Val(txtgrid1(Index).TEXT), "0.00")
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = Format((Val(FGrid1.TextMatrix(FGrid1.Row, Rate1)) * Val(FGrid1.TextMatrix(FGrid1.Row, Qty1))), "0.00")
                Amt_Cal False
        End Select
End Sub


Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
Dim j As Integer
Select Case FGrid1.Col
        Case ADType
            If txtgrid1(0).TEXT <> "" Then txtgrid1(0).TEXT = ListView.SelectedItem.TEXT
            FGrid1.TextMatrix(FGrid1.Row, ADType) = txtgrid1(0).TEXT
            If FGrid1.TextMatrix(FGrid1.Row, ADType) = "Shortage" Then
                FGrid1.TextMatrix(FGrid1.Row, Rate1) = "0.00"
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = "0.00"
            End If
            Amt_Cal False
        Case ADItem
            If RsADItem.RecordCount = 0 Or (RsADItem.EOF = True Or RsADItem.BOF = True) Or txtgrid1(0).TEXT = "" Then
                FGrid1.TextMatrix(FGrid1.Row, ADItem) = ""
                FGrid1.TextMatrix(FGrid1.Row, ADItemCode) = ""
            Else
                FGrid1.TextMatrix(FGrid1.Row, ADItemCode) = RsADItem!Code
                FGrid1.TextMatrix(FGrid1.Row, ADItem) = RsADItem!Name
                FGrid1.TextMatrix(FGrid1.Row, Rate1) = Format(IIf(IsNull(RsADItem!Rate), 0, RsADItem!Rate), "0.00")
            End If
            
            If FGrid1.TextMatrix(FGrid1.Rows - 1, 1) <> "" Then FGrid1.AddItem FGrid1.Rows
        Case Qty1
                FGrid1.TextMatrix(FGrid1.Row, Qty1) = Format(Val(txtgrid1(Index).TEXT), "0.000")
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = Format((Val(FGrid1.TextMatrix(FGrid1.Row, Rate1)) * Val(FGrid1.TextMatrix(FGrid1.Row, Qty1))), "0.00")
                Amt_Cal False
        Case Rate1
                FGrid1.TextMatrix(FGrid1.Row, Rate1) = Format(Val(txtgrid1(Index).TEXT), "0.00")
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = Format((Val(FGrid1.TextMatrix(FGrid1.Row, Rate1)) * Val(FGrid1.TextMatrix(FGrid1.Row, Qty1))), "0.00")
                Amt_Cal False
End Select
End Sub

Private Function TxtGridLeave1() As Boolean
Dim j As Integer
Dim GridCol As Byte
GridCol = FGrid1.Col
Select Case GridCol
        Case ADType
            If txtgrid1(0).TEXT <> "" Then txtgrid1(0).TEXT = ListView.SelectedItem.TEXT
            FGrid1.TextMatrix(FGrid1.Row, ADType) = txtgrid1(0).TEXT
            If FGrid1.TextMatrix(FGrid1.Row, ADType) = "Shortage" Then
                FGrid1.TextMatrix(FGrid1.Row, Rate1) = ""
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = ""
            End If
        Case ADItem
            If RsADItem.RecordCount = 0 Or (RsADItem.EOF = True Or RsADItem.BOF = True) Or txtgrid1(0).TEXT = "" Then
                FGrid1.TextMatrix(FGrid1.Row, ADItem) = ""
                FGrid1.TextMatrix(FGrid1.Row, ADItemCode) = ""
            Else
                FGrid1.TextMatrix(FGrid1.Row, ADItemCode) = RsADItem!Code
                FGrid1.TextMatrix(FGrid1.Row, ADItem) = RsADItem!Name
                FGrid1.TextMatrix(FGrid1.Row, Rate1) = Format(IIf(IsNull(RsADItem!Rate), 0, RsADItem!Rate), "0.00")
            End If
            If FGrid1.TextMatrix(FGrid1.Rows - 1, 1) <> "" Then FGrid1.AddItem FGrid1.Rows
        Case Qty1
                FGrid1.TextMatrix(FGrid1.Row, Qty1) = Format(Val(txtgrid1(0).TEXT), "0.000")
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = Format((Val(FGrid1.TextMatrix(FGrid1.Row, Rate1)) * Val(FGrid1.TextMatrix(FGrid1.Row, Qty1))), "0.00")
                Amt_Cal False
        Case Rate1
                FGrid1.TextMatrix(FGrid1.Row, Rate1) = Format(Val(txtgrid1(0).TEXT), "0.00")
                FGrid1.TextMatrix(FGrid1.Row, Amt1) = Format((Val(FGrid1.TextMatrix(FGrid1.Row, Rate1)) * Val(FGrid1.TextMatrix(FGrid1.Row, Qty1))), "0.00")
                Amt_Cal False
End Select
    TxtGridLeave1 = True
    txtgrid1(0).Visible = False
    FGrid1.SetFocus
End Function

Private Sub FGrid1_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid1_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid1.Col
        Case ADItem, ADType, Qty1, Rate1
            Call GridDblClick(Me, FGrid1, txtgrid1, 0)
    End Select
TAddMode = False
End Sub

Private Sub FGrid1_GotFocus()
    FGrid1.BackColorSel = BackColorSelEnter
    FGrid1.ForeColorSel = ForeColorSelEnter
    txtgrid1(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
'    FGrid1.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
'    FGrid1.CellBackColor = CellBackColLeave
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid1.Tag = FGrid1.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid1.Col
        Case Model1
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
        Case ADType
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
        Case Qty1
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
            FGrid1.TextMatrix(FGrid1.Row, Amt1) = ""
        Case Rate1
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
            FGrid1.TextMatrix(FGrid1.Row, Amt1) = ""
    End Select
Amt_Cal False
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid1.Col
        Case ADItem, ADType, Qty1, Rate1
            Call GridDblClick(Me, FGrid1, txtgrid1, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid1.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid1.Rows > 2 Then
                FGrid1.RemoveItem (FGrid1.Row)
            Else
                FGrid1.Rows = 1
                FGrid1.AddItem FGrid1.Rows
                FGrid1.FixedRows = 1
            End If
         End If
         For I = 1 To FGrid1.Rows - 1
            FGrid1.TextMatrix(I, 0) = I
         Next
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid1.SetFocus
End If
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid1.Col
        Case ADItem, ADType
           Call Get_Text(Me, FGrid1, txtgrid1, 0, False, KeyAscii)
        Case Amt1
            FGrid1.Col = FGrid1.Col + 1
            FGrid1.SetFocus
        Case Qty1, Rate1
           Call Get_Text(Me, FGrid1, txtgrid1, 0, True, KeyAscii)
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub
    
Private Sub FGrid1_LostFocus()
    FGrid1.BackColorSel = BackColorSelLeave
    FGrid1.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid1_Scroll()
txtgrid1(0).Visible = False
Grid_Hide
End Sub

 Private Sub Amt_Cal(Lrate As Boolean)
 Dim I As Byte
 Dim ICnt As Integer
 Dim TOTAmt As Double
 Dim TotAdd As Double
 Dim TotDel As Double
 Dim TotAdd1 As Double
 Dim TotDel1 As Double
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
    
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, ADItem) <> "" Then
            If FGrid1.TextMatrix(I, ADType) = "Shortage" Then
                FGrid1.TextMatrix(I, Rate1) = "0.00"
                FGrid1.TextMatrix(I, Amt1) = "0.00"
            End If
            If FGrid1.TextMatrix(I, ADType) = "Addition" Then
                TotAdd = TotAdd + Val(FGrid1.TextMatrix(I, Amt1))
            ElseIf FGrid1.TextMatrix(I, ADType) = "Deletion" Then
                TotDel = TotDel + Val(FGrid1.TextMatrix(I, Amt1))
            End If
        End If
    Next
    TotAdd1 = TotAdd * ICnt
    TotDel1 = TotDel * ICnt
    
    txt(TotQty) = Format(ICnt, "0")
    txt(TotGoods).TEXT = Format(TOTAmt, "0.00")
    txt(Addition).TEXT = IIf(TotAdd1 = 0, "", Format(TotAdd1, "0.00"))
    txt(Deduction).TEXT = IIf(TotDel = 0, "", Format(TotDel1, "0.00"))
    txt(SubAmt).TEXT = Format((TOTAmt + TotAdd1 - TotDel1), "0.00")
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
    FGrid.TextMatrix(FGrid.Row, EngineNo) = IIf(IsNull(RsChassis!EngineNo), "", RsChassis!EngineNo)
    FGrid.TextMatrix(FGrid.Row, SBookNo) = IIf(IsNull(RsChassis!Srv_BookNo), "", RsChassis!Srv_BookNo)
    FGrid.TextMatrix(FGrid.Row, InDate) = IIf(IsNull(RsChassis!InDate), "", RsChassis!InDate)
    FGrid.TextMatrix(FGrid.Row, SDM_STM_NO) = IIf(IsNull(RsChassis!SDM_STM_NO), "", RsChassis!SDM_STM_NO)
    FGrid.TextMatrix(FGrid.Row, Taxable) = IIf(RsChassis!Tax_YN = 1, "Yes", "No")
    FGrid.TextMatrix(FGrid.Row, Colours) = IIf(IsNull(RsChassis!Col_Desc), "", RsChassis!Col_Desc)
    FGrid.TextMatrix(FGrid.Row, ColCode) = IIf(IsNull(RsChassis!Colour_Code), "", RsChassis!Colour_Code)
    FGrid.TextMatrix(FGrid.Row, God) = IIf(IsNull(RsChassis!Godown), "", RsChassis!Godown)
    FGrid.TextMatrix(FGrid.Row, Godown) = IIf(IsNull(RsChassis!God_Name), "", RsChassis!God_Name)
    FGrid.TextMatrix(FGrid.Row, ChassisDocid) = IIf(IsNull(RsChassis!Chassis_RctDocNo), 0, RsChassis!Chassis_RctDocNo)
    FGrid.TextMatrix(FGrid.Row, MfgMth) = IIf(IsNull(RsChassis!Mfg_Month), "", RsChassis!Mfg_Month)
    FGrid.TextMatrix(FGrid.Row, MfgYr) = IIf(IsNull(RsChassis!Mfg_Yr), "", RsChassis!Mfg_Yr)
Else
    FGrid.TextMatrix(FGrid.Row, EngineNo) = ""
    FGrid.TextMatrix(FGrid.Row, SBookNo) = ""
    FGrid.TextMatrix(FGrid.Row, SDM_STM_NO) = ""
    FGrid.TextMatrix(FGrid.Row, Taxable) = ""
    FGrid.TextMatrix(FGrid.Row, Colours) = ""
    FGrid.TextMatrix(FGrid.Row, InDate) = ""
    FGrid.TextMatrix(FGrid.Row, ColCode) = ""
    FGrid.TextMatrix(FGrid.Row, God) = ""
    FGrid.TextMatrix(FGrid.Row, Godown) = ""
    FGrid.TextMatrix(FGrid.Row, ChassisDocid) = ""
    FGrid.TextMatrix(FGrid.Row, MfgMth) = ""
    FGrid.TextMatrix(FGrid.Row, MfgYr) = ""
End If
End Sub

Private Function ChkDul_Chassis() As Boolean
Dim I As Integer
If TxtGrid(0).TEXT = FGrid.TextMatrix(FGrid.Row, ChassisNo) Then
    ChkDul_Chassis = False
    Exit Function
End If
If GCn.Execute("select COUNT(*) from veh_stock where ChassisNo = '" & TxtGrid(0).TEXT & "' AND Pur_DocId<>''").Fields(0).Value > 0 Then
    MsgBox "Same Chassis No already exist in stock", vbInformation, "Duplicate Chassis"
    ChkDul_Chassis = True
    Exit Function
End If
For I = 1 To FGrid.Rows - 1
    If I <> FGrid.Row Then
        If FGrid.TextMatrix(I, ChassisNo) = TxtGrid(0).TEXT Then
            MsgBox "Same Chassis No already taken ", vbInformation, "Duplicate Chassis"
            ChkDul_Chassis = True
            Exit Function
        End If
    End If
Next
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
        Case ChassisNo, SDM_STM_NO
            TxtGrid(0).MaxLength = 20
        Case MfgMth
            TxtGrid(0).MaxLength = 9
        Case MfgYr
            TxtGrid(0).MaxLength = 4
        Case Srlno, EngineNo
            TxtGrid(0).MaxLength = 25
        Case SBookNo
            TxtGrid(0).MaxLength = 10
        Case Model1, EngineNo, Colours, Rate, InDate, Godown
             TxtGrid(0).MaxLength = 0
    End Select
End Sub

