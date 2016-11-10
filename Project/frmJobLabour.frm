VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmJobLabour 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Labour Entry"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11805
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
   LinkTopic       =   " "
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11805
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
      Left            =   5160
      MaxLength       =   12
      TabIndex        =   108
      Top             =   6540
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   35
      Left            =   2895
      MaxLength       =   12
      TabIndex        =   107
      Top             =   6570
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DgRateType 
      Height          =   2100
      Left            =   7335
      Negotiate       =   -1  'True
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   5940
      Visible         =   0   'False
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   3704
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   14413565
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Rate Type"
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
            DividerStyle    =   1
            ColumnWidth     =   2594.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   29
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   2730
      Width           =   975
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
      Index           =   33
      Left            =   6630
      MaxLength       =   12
      TabIndex        =   103
      Top             =   2730
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSDataGridLib.DataGrid DGJobCode 
      Height          =   2730
      Left            =   3840
      Negotiate       =   -1  'True
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   7170
      Visible         =   0   'False
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   4815
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
         DataField       =   "Description"
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
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4320
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
      Index           =   34
      Left            =   6630
      MaxLength       =   12
      TabIndex        =   99
      Top             =   2985
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   -1920
      TabIndex        =   97
      Top             =   3645
      Visible         =   0   'False
      Width           =   2115
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   75
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   15
         Width           =   1935
         _ExtentX        =   3413
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
      Index           =   32
      Left            =   6630
      MaxLength       =   12
      TabIndex        =   27
      Top             =   3255
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSDataGridLib.DataGrid DGGatePass 
      Height          =   2910
      Left            =   990
      Negotiate       =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   7515
      Visible         =   0   'False
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   5133
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "GP No."
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
         DataField       =   "GatePassDate"
         Caption         =   "GP Date"
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
         DataField       =   "ContractRecdDate"
         Caption         =   "Recd.Date"
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
         DataField       =   "ContractAmt"
         Caption         =   "Contract Amt."
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
         DataField       =   "FinName"
         Caption         =   "Contractor Name"
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
         DataField       =   "Remarks"
         Caption         =   "Purpose"
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
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   4830.236
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      Index           =   30
      Left            =   2325
      MaxLength       =   12
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2985
      Width           =   660
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   31
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2985
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      Index           =   28
      Left            =   2325
      MaxLength       =   12
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2730
      Width           =   660
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   240
      HideSelection   =   0   'False
      Left            =   8430
      TabIndex        =   90
      Top             =   2070
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
      Index           =   12
      Left            =   885
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1965
      Width           =   1980
   End
   Begin MSDataGridLib.DataGrid DGJob 
      Height          =   2520
      Left            =   120
      Negotiate       =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   7140
      Visible         =   0   'False
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   4445
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "FindJobNo"
         Caption         =   "Job No."
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
         DataField       =   "Chassis"
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
      BeginProperty Column02 
         DataField       =   "RegNo"
         Caption         =   "Reg. No"
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
         DataField       =   "Model"
         Caption         =   "Model"
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
         DataField       =   "VehSerialNo"
         Caption         =   "Veh.Srl No."
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
         DataField       =   "Name"
         Caption         =   "Owner Name"
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
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   3
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3195.213
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1275
      Index           =   1
      Left            =   7995
      TabIndex        =   4
      Top             =   2025
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   2249
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   12632319
      ForeColorFixed  =   8388608
      BackColorSel    =   -2147483624
      ForeColorSel    =   12582912
      BackColorBkg    =   15132390
      GridColor       =   8438015
      GridColorFixed  =   8421504
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      GridLineWidthFixed=   1
      FormatString    =   "SSS"
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
      Index           =   15
      Left            =   4395
      MaxLength       =   4
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1965
      Width           =   2460
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   1785
      MaxLength       =   4
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "1234"
      Top             =   690
      Width           =   570
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   26
      Left            =   8460
      MaxLength       =   40
      TabIndex        =   33
      Top             =   1710
      Width           =   3225
   End
   Begin MSDataGridLib.DataGrid DGLabour 
      Height          =   2730
      Left            =   6975
      Negotiate       =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   4815
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
      ColumnCount     =   3
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
         Caption         =   "Labour Description"
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
         DataField       =   "LabGrp_Desc"
         Caption         =   "Group"
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
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4320
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   5640
      MaxLength       =   12
      TabIndex        =   5
      Text            =   "29/DEC/2003"
      Top             =   435
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1785
      MaxLength       =   25
      TabIndex        =   1
      Top             =   435
      Width           =   450
   End
   Begin VB.TextBox TxtGrid1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
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
      Index           =   0
      Left            =   9000
      MaxLength       =   40
      TabIndex        =   34
      Top             =   3960
      Visible         =   0   'False
      Width           =   705
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
      Index           =   27
      Left            =   10605
      TabIndex        =   32
      Top             =   1455
      Width           =   1080
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   11
      Left            =   885
      MaxLength       =   40
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "0123456789012345678901234567890123456789"
      Top             =   1710
      Width           =   4320
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1785
      MaxLength       =   14
      TabIndex        =   8
      Top             =   945
      Width           =   1740
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3585
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "Help"
      Top             =   435
      Width           =   1050
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   4
      Left            =   5640
      MaxLength       =   12
      TabIndex        =   7
      Top             =   690
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   19
      Left            =   3960
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "999999.99"
      Top             =   2475
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   17
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2220
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      Index           =   20
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3255
      Width           =   975
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      Index           =   21
      Left            =   2010
      MaxLength       =   8
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "99999.99"
      Top             =   3240
      Width           =   975
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
      Index           =   13
      Left            =   6405
      MaxLength       =   4
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1710
      Width           =   450
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   18
      Left            =   2325
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "999.99"
      Top             =   2475
      Width           =   660
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   661
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   9
      Top             =   945
      Width           =   2055
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
      Index           =   9
      Left            =   1785
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1455
      Width           =   1740
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   8460
      MaxLength       =   20
      TabIndex        =   30
      Top             =   1200
      Width           =   3225
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   8
      Left            =   4800
      MaxLength       =   25
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      Index           =   16
      Left            =   2325
      MaxLength       =   12
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2220
      Width           =   660
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   8460
      MaxLength       =   8
      TabIndex        =   28
      Top             =   945
      Width           =   990
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   23
      Left            =   10605
      MaxLength       =   25
      TabIndex        =   29
      Top             =   945
      Width           =   1080
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   10
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1455
      Width           =   2055
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Index           =   25
      Left            =   8460
      MaxLength       =   8
      TabIndex        =   31
      Text            =   "999999"
      Top             =   1455
      Width           =   705
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
      Index           =   7
      Left            =   1785
      MaxLength       =   15
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1740
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   2085
      Left            =   195
      TabIndex        =   3
      Top             =   3645
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   3678
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   15
      BackColorFixed  =   12640511
      ForeColorFixed  =   8388608
      BackColorSel    =   -2147483624
      ForeColorSel    =   12582912
      BackColorBkg    =   14145495
      GridColor       =   12640511
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      GridLineWidthFixed=   1
      FormatString    =   "SSS"
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
      _Band(0).Cols   =   15
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSDataGridLib.DataGrid DGMech 
      Height          =   2865
      Left            =   1245
      Negotiate       =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   7245
      Visible         =   0   'False
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   5054
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
      BeginProperty Column01 
         DataField       =   "name"
         Caption         =   "Staff Name"
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
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4199.811
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGridLab2Copy 
      Height          =   1425
      Left            =   1905
      TabIndex        =   91
      Top             =   7560
      Visible         =   0   'False
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   2514
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   12632319
      ForeColorFixed  =   8388608
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   15132390
      GridColor       =   8438015
      GridColorFixed  =   8421504
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      GridLineWidthFixed=   1
      FormatString    =   "|S_No|Lab_Code|Mech_Code|Mech_Name             "
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGridLab2 
      Height          =   1425
      Left            =   585
      TabIndex        =   89
      Top             =   7530
      Visible         =   0   'False
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   2514
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   12632319
      ForeColorFixed  =   8388608
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   15132390
      GridColor       =   8438015
      GridColorFixed  =   8421504
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      GridLineWidthFixed=   1
      FormatString    =   "|S_No|Lab_Code|Mech_Code|Mech_Name             "
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coupon No. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   24
      Left            =   4935
      TabIndex        =   105
      Top             =   2738
      Visible         =   0   'False
      Width           =   1080
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
      Left            =   180
      TabIndex        =   102
      Top             =   5700
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coupon Value :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   22
      Left            =   4980
      TabIndex        =   100
      Top             =   2993
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Completion Date :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   21
      Left            =   4995
      TabIndex        =   96
      Top             =   3263
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Dlr+Self ->   Hrs :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   20
      Left            =   375
      TabIndex        =   95
      Top             =   2985
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   19
      Left            =   3180
      TabIndex        =   94
      Top             =   2985
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer ->   Hrs :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   10
      Left            =   465
      TabIndex        =   93
      Top             =   2730
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   6
      Left            =   3180
      TabIndex        =   92
      Top             =   2730
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   4
      Left            =   165
      TabIndex        =   88
      Top             =   1965
      Width           =   300
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   780
      TabIndex        =   87
      Top             =   1965
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   225
      Left            =   180
      TabIndex        =   86
      Top             =   2235
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Charged :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   14
      Left            =   3090
      TabIndex        =   85
      Top             =   3240
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Advisor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   9
      Left            =   3060
      TabIndex        =   84
      Top             =   1965
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   2
      Left            =   165
      TabIndex        =   83
      Top             =   705
      Width           =   1080
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
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   5
      Left            =   8355
      TabIndex        =   60
      Top             =   1725
      Width           =   45
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Mechanic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   16
      Left            =   7140
      TabIndex        =   59
      Top             =   1710
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Dt. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   81
      Top             =   435
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Owner"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   39
      Left            =   165
      TabIndex        =   80
      Top             =   1710
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chargeable -> Hrs :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   17
      Left            =   645
      TabIndex        =   79
      Top             =   2475
      Width           =   1620
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "External -> Paid :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   23
      Left            =   615
      TabIndex        =   78
      Top             =   3240
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   40
      Left            =   3150
      TabIndex        =   77
      Top             =   2475
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle's on Floor :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   7
      Left            =   9675
      TabIndex        =   76
      Top             =   720
      Width           =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Service"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   27
      Left            =   7140
      TabIndex        =   75
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Job No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   36
      Left            =   7155
      TabIndex        =   74
      Top             =   945
      Width           =   1035
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division            :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   7140
      TabIndex        =   73
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label lblDocId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job DocID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   225
      Left            =   7140
      TabIndex        =   72
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Jobs Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   13
      Left            =   165
      TabIndex        =   71
      Top             =   450
      Width           =   1305
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   29
      Left            =   1680
      TabIndex        =   70
      Top             =   435
      Width           =   45
   End
   Begin VB.Line Line1 
      X1              =   165
      X2              =   11760
      Y1              =   3525
      Y2              =   3525
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   33
      Left            =   6300
      TabIndex        =   67
      Top             =   1710
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JC Close Dt. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   225
      Index           =   11
      Left            =   4500
      TabIndex        =   66
      Top             =   705
      Width           =   1110
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
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   30
      Left            =   10530
      TabIndex        =   65
      Top             =   1455
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "History Srl No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   15
      Left            =   9330
      TabIndex        =   64
      Top             =   1455
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   29
      Left            =   3180
      TabIndex        =   63
      Top             =   2220
      Width           =   720
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   10
      Left            =   1680
      TabIndex        =   62
      Top             =   705
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt. Vehicle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   18
      Left            =   5205
      TabIndex        =   61
      Top             =   1710
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard No. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   12
      Left            =   2400
      TabIndex        =   58
      Top             =   435
      Width           =   1125
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   57
      Top             =   945
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   56
      Top             =   945
      Width           =   1035
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   55
      Top             =   1455
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Serial No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   165
      TabIndex        =   54
      Top             =   1455
      Width           =   1455
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   90
      Left            =   1680
      TabIndex        =   53
      Top             =   945
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   165
      TabIndex        =   52
      Top             =   945
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   1530
      Left            =   7050
      Top             =   465
      Width           =   4725
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code      :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   8910
      TabIndex        =   51
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   37
      Left            =   3600
      TabIndex        =   50
      Top             =   1455
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warranty ->   Hrs :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   35
      Left            =   810
      TabIndex        =   49
      Top             =   2220
      Width           =   1455
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
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   16
      Left            =   8355
      TabIndex        =   48
      Top             =   945
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   12
      Left            =   8355
      TabIndex        =   47
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last KMs "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   32
      Left            =   7140
      TabIndex        =   46
      Top             =   1455
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   18
      Left            =   1680
      TabIndex        =   45
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   38
      Left            =   165
      TabIndex        =   44
      Top             =   1200
      Width           =   495
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
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   14
      Left            =   10530
      TabIndex        =   43
      Top             =   945
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Job Dt."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   34
      Left            =   9525
      TabIndex        =   42
      Top             =   945
      Width           =   975
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   26
      Left            =   780
      TabIndex        =   41
      Top             =   1710
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   8
      Left            =   4665
      TabIndex        =   40
      Top             =   1455
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   39
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   33
      Left            =   3720
      TabIndex        =   38
      Top             =   1200
      Width           =   915
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
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   9
      Left            =   8355
      TabIndex        =   37
      Top             =   1455
      Width           =   45
   End
   Begin VB.Label LblTotVeh 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   11370
      TabIndex        =   36
      Top             =   720
      Width           =   315
   End
End
Attribute VB_Name = "frmJobLabour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TAddMode As Boolean
Dim GridKey As Integer
Dim ADDFLAG$
Dim ApplyServTax As Byte
Dim ForSiteCode$
Dim MyIndex As Byte
Dim Rst As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim RsGatePass As ADODB.Recordset
Dim RsJob As ADODB.Recordset
Dim RsLab As ADODB.Recordset
Dim RsMech As ADODB.Recordset
Dim RsJobCode As ADODB.Recordset
Dim blnRateEditableYn As Boolean


Dim ListArray As Variant
Dim mListItem As ListItem

'Dim GridRow1() As Integer
Dim mGridStartRow As Integer
Dim mGridEndRow As Integer


Private Const JCType As Byte = 1
Private Const JobNo As Byte = 2
Private Const JobDt As Byte = 3
Private Const JobCDt As Byte = 4
Private Const VehRegNo As Byte = 5
Private Const Chassis As Byte = 6
Private Const Model As Byte = 7
Private Const Engine As Byte = 8
Private Const VehSrlNo As Byte = 9
Private Const SrvType As Byte = 10
Private Const OwnerName As Byte = 11
Private Const City As Byte = 12
Private Const GovtYn As Byte = 13
Private Const VehicleType As Byte = 14
Private Const SrvAdvisor As Byte = 15
Private Const WarrHrs As Byte = 16
Private Const WarrAmt As Byte = 17
Private Const ChgHrs As Byte = 18
Private Const ChgAmt As Byte = 19
Private Const ExtLab As Byte = 21
Private Const ExtLabChg As Byte = 20
Private Const LastJobNo As Byte = 22
Private Const LastJobDt As Byte = 23
Private Const LastSrv As Byte = 24
Private Const LastKMS As Byte = 25
Private Const LastMech As Byte = 26
Private Const HistNo As Byte = 27
Private Const MfgHrs As Byte = 28
Private Const MfgAmt As Byte = 29
Private Const OthHrs As Byte = 30
Private Const OthAmt As Byte = 31
Private Const RateSystemYN As Byte = 32
Private Const Coupon As Byte = 33
Private Const CouponValue As Byte = 34
Private Const SoldDate As Byte = 35

'Text Box (Grid)
Private Const mTxtGrid1 As Byte = 1

'Fgrid1 Columns
Private Const C_LabCode As Byte = 1
Private Const C_LabName As Byte = 2
Private Const C_MechVoice As Byte = 3
Private Const C_Fixed As Byte = 4
Private Const C_TaxYN As Byte = 5   'Introduced on 28-05-03 but not activated
Private Const C_PaidBy As Byte = 6
Private Const C_ChrgType As Byte = 7
Private Const C_ActHrs As Byte = 8
Private Const C_Hrs As Byte = 9
Private Const C_Rate As Byte = 10
Private Const C_Amt As Byte = 11
Private Const C_External As Byte = 12
Private Const C_GPNo As Byte = 13
Private Const C_Remarks As Byte = 14
Private Const C_ContName As Byte = 15
Private Const C_WIssueDt As Byte = 16
Private Const C_WRecdDt As Byte = 17
Private Const C_ContAmt As Byte = 18
Private Const C_ContCode As Byte = 19
Private Const C_Major As Byte = 20
Private Const C_JobCode As Byte = 21


Private Const Col_DepItem As Byte = 22
Private Const Col_DepitemPer As Byte = 23
Private Const Col_DepCode As Byte = 24
Private Const Col_DepPer As Byte = 25
Private Const Col_DepAmt As Byte = 26
Private Const Col_InsuranceAmt As Byte = 27
Private Const Col_DiffPeried As Byte = 28


'FGrid Columns
Private Const Srlno As Byte = 1
Private Const LabCode As Byte = 2
Private Const MechCode As Byte = 3
Private Const MechName As Byte = 4

Private Const BackColorSelEnter As String = &HEBB7EC   '&HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Dim mAddBy$, mAddDate As String

Private Sub DGJob_Click()
If Master.RecordCount > 0 Then
    Call History_Field
End If
txt(MyIndex).SetFocus
DGJob.Visible = False
End Sub

Private Sub DGGatePass_Click()
If RsGatePass.RecordCount > 0 Then
    txtgrid1(0) = RsGatePass!Code
End If
txtgrid1(0).SetFocus
DGGatePass.Visible = False
End Sub
Private Sub DGLabour_Click()
If RsLab.RecordCount > 0 Then
    txtgrid1(0).Tag = RsLab!Code
    txtgrid1(0).TEXT = RsLab!Name
End If
txtgrid1(0).SetFocus
DGLabour.Visible = False
End Sub
Private Sub DGJobCode_Click()
If RsJobCode.RecordCount > 0 Then
    txtgrid1(0).TEXT = RsJobCode!Code
End If
txtgrid1(0).SetFocus
DGJobCode.Visible = False
End Sub
Private Sub FGrid1_LostFocus()
    FGrid1.BackColorSel = BackColorSelLeave
    FGrid1.ForeColorSel = FGrid1.ForeColor
End Sub
Private Sub FGrid1_RowColChange()
Dim I As Integer
    'Filling Mechanic to Display from FGridLab2
    GridSel(1).Redraw = False
    GridSel(1).Rows = 1
    I = 1
    For I = 1 To FGridLab2.Rows - 1
        If (FGridLab2.TextMatrix(I, LabCode) = FGrid1.TextMatrix(FGrid1.Row, C_LabCode)) And FGridLab2.TextMatrix(I, LabCode) <> "" And Val(FGrid1.TextMatrix(FGrid1.Row, 0)) = FGridLab2.RowData(I) Then
            GridSel(1).AddItem ""
            GridSel(1).Row = GridSel(1).Rows - 1
            GridSel(1).Col = 0
            GridSel(1).CellFontName = "WINGDINGS"
            GridSel(1).CellFontSize = 14
            With GridSel(1)
                .TextMatrix(GridSel(1).Rows - 1, 0) = "ü"
                .TextMatrix(GridSel(1).Rows - 1, LabCode) = FGridLab2.TextMatrix(I, LabCode)
                .TextMatrix(GridSel(1).Rows - 1, MechCode) = FGridLab2.TextMatrix(I, MechCode)
                .TextMatrix(GridSel(1).Rows - 1, MechName) = FGridLab2.TextMatrix(I, MechName)
            End With
        End If
    Next
    '******** Rahul UN AutoMobiles 10-04-2003 Because is No Mechanic Sepecify
    '******** GridSel(1) Show Blank and Default One Row Selected
    
'    If GridSel(1).Rows = 1 Then
'        GridSel(1).Visible = False 'on 24-07-2003
'    Else
'        GridSel(1).Visible = True 'on 24-07-2003
        GridSel(1).AddItem ""
        GridSel(1).FixedRows = 1
'        GridSel(1).Redraw = True
'    End If
    GridSel(1).Visible = True  'on 24-07-2003
    GridSel(1).Redraw = True
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
Dim I As Byte
Dim SrNo As Integer
    '' pending points :
    '' Editing not allowed for Chargeable Labour if JobCard is Closed -- REQ
    '' There should be a grid on form, in which complaints of driver will appear -- REQ through button
    '' After Job Close User Can only modify
    '' Contractor related information and Row Deletion not allowed -- REQ
    WinSetting Me:    Ini_Grid
    TopCtrl1.Tag = PubUParam
    ForSiteCode = PubSiteCode
    Call BlankText
    '**Speed
    Me.Show
    DoEvents
    '**
    
     Dim sitecond As String
     
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and  " & cMID("JL.JOB_DOCID", "3", "1") & "='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "select Distinct JL.JOB_DOCID AS SEARCHCODE from Job_Lab as JL where left(JL.Job_DocId,1)='" & PubDivCode & "' " & sitecond & " order by JL.Job_DocID desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select  Distinct Top 1 JL.JOB_DOCID AS SEARCHCODE from Job_Lab as JL where left(JL.Job_DocId,1)='" & PubDivCode & "' " & sitecond & " order by JL.Job_DocID desc", GCn, adOpenDynamic, adLockOptimistic
    End If
    
    
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and  " & cMID("J.DOCID", "3", "1") & "='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    Set RsJob = New ADODB.Recordset
    With RsJob
        .CursorLocation = adUseClient
        .Open "select J.DocId AS CODE," & cCStr("J.Job_No") & " As FindJobNo,J.JobCloseDate,J.CardNo,HC.Model,HC.RegNo," _
            & "HC.Chassis,HC.Engine,HC.VehSerialNo,HC.Name " _
            & "from (job_card as J left Join Hiscard as HC on J.CardNo=HC.CardNo) " _
            & "where left(J.DocId,1)='" & PubDivCode & "' " & sitecond & " and (right(j.DocId_InvSpr,8) <> 'Cancelld' Or DocId_InvSpr Is Null) order by J.Job_No", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGJob.DataSource = RsJob
    
    Set RsGatePass = New ADODB.Recordset
    RsGatePass.CursorLocation = adUseClient
    RsGatePass.Open "Select GatePassNo as Code,GatePassNo as Name,GatePassDate,Job_DocID,ContractRecdDate,ContractAmt,Remarks,ContractCode,CF.FinName " _
        & "FROM Job_GatePass left join ContractFinance as CF on Job_GatePass.ContractCode=CF.FinCode " _
        & "where left(GatePassNo,1)='" & PubDivCode & "' Order by GatePassNo", GCn, adOpenDynamic, adLockOptimistic
    Set DGGatePass.DataSource = RsGatePass
'    RsGatePass.Sort = "Code"
    
    Set RsLab = New ADODB.Recordset
    'Setting Mechaninc recordset for searching
    GSQL = "Select Emp_Code as code,Emp_Name as Name FROM Emp_Mast where Div_Code='" & PubDivCode & "' And  Emp_type=1 and Designation ='MECHANIC' and (LeftOn Is Null or LeftOn >=" & ConvertDate(PubLoginDate) & ") Order by Emp_name"
    Set RsMech = New ADODB.Recordset
    RsMech.CursorLocation = adUseClient
    RsMech.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    Set DGMech.DataSource = RsMech
    
    Set RsJobCode = New ADODB.Recordset
    With RsJobCode
        .CursorLocation = adUseClient
        .Open "Select Code,Description from WarrJobMast", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGJobCode.DataSource = RsJobCode
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If TopCtrl1.TopText2 <> "Browse" Then
    If ADDFLAG <> "B" Then
        If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RsJob = Nothing
    Set RsLab = Nothing
    Set RsJobCode = Nothing
    Set RsGatePass = Nothing
End Sub

Private Sub GridSel_GotFocus(Index As Integer)
    GridSel(1).BackColorSel = BackColorSelEnter
    GridSel(1).ForeColorSel = ForeColorSelEnter
End Sub

Private Sub GridSel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeysA vbKeyTab, True
If GridSel(Index).Rows < 1 Then Exit Sub
If (KeyCode = vbKeySpace Or KeyCode = vbKeyReturn) And GridSel(Index).Col = 0 Then
    GridSel(Index).CellFontName = "WINGDINGS"
    GridSel(Index).CellFontSize = 14
    GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = IIf(GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "ü", " ", "ü")
End If
End Sub

Private Sub GridSel_KeyPress(Index As Integer, KeyAscii As Integer)
If GridSel(Index).Col = 0 Or GridSel(Index).Row = 0 Then Exit Sub
Select Case Index
    Case 1
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsMech, KeyAscii, "Name", vbCyan
End Select
TxtSearch.Tag = Index
End Sub

Private Sub GridSel_LeaveCell(Index As Integer)
GridSel(1).CellBackColor = &H80000018
End Sub

Private Sub GridSel_LostFocus(Index As Integer)
'Dim i As Integer
'GridSel(1).BackColorSel = BackColorSelLeave
'GridSel(1).ForeColorSel = GridSel(1).ForeColor
''** Job_Lab2 for Multiple Mechanic
'FGridLab2Copy.Rows = 1
'GridSel(1).Redraw = False
'GridSel(1).Rows = 1
'For i = 1 To FGridLab2.Rows - 1
''    If GridSel(1).TextMatrix(GridSel(1).Row, 0) = "ü" Then
'    If FGridLab2.TextMatrix(i, LabCode) <> FGrid1.TextMatrix(FGrid1.Row, C_LabCode) Then
'        FGridLab2Copy.AddItem ""
''        GridSel(1).Row = GridSel(1).Rows - 1
''        GridSel(1).Col = 0
''        GridSel(1).CellFontName = "WINGDINGS"
''        GridSel(1).CellFontSize = 14
'        With FGridLab2Copy
'            .TextMatrix(FGridLab2Copy.Rows - 1, LabCode) = FGridLab2.TextMatrix(i, LabCode)
'            .TextMatrix(FGridLab2Copy.Rows - 1, MechCode) = FGridLab2.TextMatrix(i, MechCode)
'            .TextMatrix(FGridLab2Copy.Rows - 1, MechName) = FGridLab2.TextMatrix(i, MechName)
'        End With
'    End If
'Next

End Sub

Private Sub GridSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ADDFLAG = "B" Then Exit Sub
If GridSel(Index).Col <> 0 Then Exit Sub
mGridStartRow = GridSel(Index).Row
End Sub

Private Sub GridSel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim j As Integer
If ADDFLAG = "B" Then Exit Sub
If GridSel(Index).Col <> 0 Or mGridStartRow = 0 Then Exit Sub
mGridEndRow = GridSel(Index).RowSel
For j = mGridStartRow To mGridEndRow
    GridSel(Index).Row = j
    GridSel(Index).Col = 0
    GridSel(Index).CellFontName = "WINGDINGS"
    GridSel(Index).CellFontSize = 14
    GridSel(Index).TextMatrix(j, 0) = IIf(GridSel(Index).TextMatrix(j, 0) = "ü", " ", "ü")
Next
mGridStartRow = 0
End Sub

Private Sub GridSel_Validate(Index As Integer, Cancel As Boolean)
Dim I As Integer

GridSel(1).BackColorSel = BackColorSelLeave
GridSel(1).ForeColorSel = GridSel(1).ForeColor
'** Job_Lab2 for Multiple Mechanic
FGridLab2Copy.Rows = 1
'Making Copy of Total Mechanics excluding currently selected Labour
For I = 1 To FGridLab2.Rows - 1
    If FGridLab2.TextMatrix(I, LabCode) <> FGrid1.TextMatrix(FGrid1.Row, C_LabCode) Then
        FGridLab2Copy.AddItem ""
        With FGridLab2Copy
            .TextMatrix(FGridLab2Copy.Rows - 1, LabCode) = FGridLab2.TextMatrix(I, LabCode)
            .TextMatrix(FGridLab2Copy.Rows - 1, MechCode) = FGridLab2.TextMatrix(I, MechCode)
            .TextMatrix(FGridLab2Copy.Rows - 1, MechName) = FGridLab2.TextMatrix(I, MechName)
        End With
    End If
Next
'**
'Restoring copy to Total Grid
FGridLab2.Rows = 1
For I = 1 To FGridLab2Copy.Rows - 1
    FGridLab2.AddItem ""
    With FGridLab2
        .TextMatrix(FGridLab2.Rows - 1, LabCode) = FGridLab2Copy.TextMatrix(I, LabCode)
        .TextMatrix(FGridLab2.Rows - 1, MechCode) = FGridLab2Copy.TextMatrix(I, MechCode)
        .TextMatrix(FGridLab2.Rows - 1, MechName) = FGridLab2Copy.TextMatrix(I, MechName)
    End With
Next
'**
'Add Mechanic for selected Labour
For I = 1 To GridSel(1).Rows - 1
    If Trim(GridSel(1).TextMatrix(I, 0)) <> "" Then
        FGridLab2.AddItem ""
        With FGridLab2
            .TextMatrix(FGridLab2.Rows - 1, LabCode) = GridSel(1).TextMatrix(I, LabCode)
            .TextMatrix(FGridLab2.Rows - 1, MechCode) = GridSel(1).TextMatrix(I, MechCode)
            .TextMatrix(FGridLab2.Rows - 1, MechName) = GridSel(1).TextMatrix(I, MechName)
        End With
    End If
Next
FGridLab2Copy.Rows = 1
End Sub

Private Sub ListView_Click()
On Error GoTo ELoop
        txtgrid1(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
        FrmList.Visible = False
        txtgrid1(Val(ListView.Tag)).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    txt(JCType).TEXT = "Yes"
    txt(JCType).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant, mTrans As Boolean
If txt(JobCDt) <> "" Then
    MsgBox "Job Closed. Edit/Delete denied!", vbOKOnly, "Validation": Exit Sub
End If
    If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        GCn.BeginTrans
        mTrans = True
        GCn.Execute "Delete from Job_Lab  where job_Docid='" & lblDocId.CAPTION & "'"
        GCn.Execute "Delete from Job_Lab2  where job_Docid='" & lblDocId.CAPTION & "'"
        GCn.CommitTrans
        mTrans = False
        
        Master.Requery
        Call UpdRequery
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
            Call MoveRec
        Else
            Call BlankText
        End If
        BUTTONS True, Me, Master, 0
    End If
    Exit Sub
eloop1:
    If mTrans Then GCn.RollbackTrans
    MsgBox err.Description, vbCritical, "Deletion Message"
End Sub

Private Sub TopCtrl1_eEdit()
Dim I As Integer
On Error GoTo eloop1
If txt(JobCDt) <> "" Then
    MsgBox "Job Closed. Edit/Delete denied!", vbOKOnly, "Validation": Exit Sub
End If
    
    Disp_Text SETS("EDIT", Me, Master)
    For I = 0 To txt.Count - 1
        txt(I).Enabled = False
    Next I
    FGrid1.AddItem FGrid1.Rows
    FGrid1.Col = C_LabName: FGrid1.SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
Unload Me
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    'If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    
     Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and  " & cMID("JL.JOB_DOCID", "3", "1") & "='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    GSQL = "select Distinct JL.Job_DocId as searchcode, J.Job_No," & cTrim(cMID("JL.Job_DocId", "9", "5")) & " as Prefix, J.Job_Date, J.JobCloseDate, " & _
            "J.Govt_YN, HC.Model,HC.RegNo, HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name, HC.Add1, HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc, City.CityName" & _
            " from  (((Job_Lab as JL Left Join job_card as J on JL.Job_DocId=J.DocId)" & _
            "left Join Hiscard as HC on J.CardNo=HC.CardNo) " & _
            "left Join Service_Type as ST on J.Serv_Type=ST.Serv_Type) " & _
            "Left Join City on HC.CityCode=City.CityCode " & _
        "where left(J.docid,1)='" & PubDivCode & "' " & sitecond & " order by J.Job_No"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("searchcode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select Distinct JL.JOB_DOCID AS SEARCHCODE from Job_Lab as JL where left(JL.Job_DocId,1)='" & PubDivCode & "' And JL.JOB_DOCID ='" & MyValue & "' order by JL.Job_DocID desc")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
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
    If MsgBox("Cancel Entry ?", vbExclamation + vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        Call MoveRec
    Else
        Me.ActiveControl.SetFocus
    End If
    Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eRef()
    Call UpdRequery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer, j As Integer
    Dim mTrans As Boolean, MechFound As Boolean
    Dim SrNo As Integer
    Dim mExternal$, mTaxYN As Byte
    Dim mChgRate As Single, mWarRate As Single, mLabAmt As Single
    Dim mChgHrs As Single, mWarHrs As Single, mChgAmt As Single, mWarAmt As Single
'    On Error GoTo errlbl
    
    If txtgrid1(0).Visible = True Then
        If TxtGrid1Leave = False Then
            txtgrid1(0).SetFocus
            Exit Sub
        Else
            txtgrid1(0).Visible = False
        End If
    End If
    Grid_Hide
    'Checking Job Closed by other user during add/edit of labour
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select JobCloseDate,ClosedU_Name,ClosedU_EntDt from Job_Card where Job_Card.DocId='" & lblDocId.CAPTION & "'", GCn, adOpenDynamic, adLockOptimistic
    If Not IsNull(Rst!JobCloseDate) Then 'Job Closed
        MsgBox "Job Already Closed by User " & Rst!ClosedU_Name & " Dt." & Rst!ClosedU_EntDt
        GoTo errlbl
    End If
    Set Rst = Nothing
    'eof of checking
    If IsValid(txt(JobNo), "Job Card No.") = False Then Exit Sub
    '' checking for data in fgrid1
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, C_LabCode) <> "" And FGrid1.TextMatrix(I, C_Amt) <> "" Then GoTo Mynxt
    Next I
    MsgBox "No Labour Details Feeded" & vbCrLf & " or " & vbCrLf & "Amount is Zero ", vbCritical, "Validation"
    FGrid1.SetFocus: Exit Sub
Mynxt:
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, C_ChrgType) = "Warranty" Then
            If left(FGrid1.TextMatrix(I, C_PaidBy), 1) <> "M" Then
                MsgBox "Paid By must be Manufacturer"
                FGrid1.Row = I
                FGrid1.Col = C_PaidBy
                FGrid1.SetFocus: Exit Sub
                Exit For
            End If
        End If
    Next I
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, C_ChrgType) = "Free Service" Or FGrid1.TextMatrix(I, C_ChrgType) = "PDI" Then
            If left(FGrid1.TextMatrix(I, C_PaidBy), 1) = "C" Then
                MsgBox "Paid By must be Manufacturer / Other Dealer"
                FGrid1.Row = I
                FGrid1.Col = C_PaidBy
                FGrid1.SetFocus: Exit Sub
                Exit For
            End If
        End If
    Next I
    '' eof : checking of data in fgrid1
    For I = 1 To FGrid1.Rows - 1
        '** Job_Lab2 for Multiple Mechanic
        If FGrid1.TextMatrix(I, C_LabCode) <> "" Then
            MechFound = False
            For j = 1 To FGridLab2.Rows - 1
                If FGridLab2.TextMatrix(j, LabCode) = FGrid1.TextMatrix(I, C_LabCode) Then
                    MechFound = True
                    Exit For
                End If
            Next
            If MechFound = False Then
                MsgBox "Mechanic Name is Required in Row No." & I, vbInformation, "Validation"
                FGrid1.Row = I
                FillMechGrid I
                GridSel(1).SetFocus
                Exit Sub
            End If
        End If
        If RSOJPR = True Then
            If FGrid1.TextMatrix(I, C_ChrgType) = "Warranty" Then
                If FGrid1.TextMatrix(I, C_JobCode) = "" Then
                    MsgBox "Job code For Warranty Labour is Required in Row No." & I, vbInformation, "Validation"
                    FGrid1.Row = I
                    FGrid1.Col = C_JobCode
                    FGrid1.SetFocus
                    SendKeys vbKeyReturn
                    Exit Sub
                End If
            End If
        End If
    Next I
    GCn.BeginTrans
    mTrans = True
'    If AddFlag = "E" Then  bvn
    GCn.Execute "Delete from Job_Lab  where JOB_Docid='" & lblDocId.CAPTION & "'"
    GCn.Execute "Delete from Job_Lab2  where JOB_Docid='" & lblDocId.CAPTION & "'"
'    End If
    
    SrNo = 1
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, C_LabCode) <> "" And (Not IsNull(FGrid1.TextMatrix(I, C_LabCode))) Then
            If FGrid1.TextMatrix(I, C_ChrgType) = "Chargable" Or FGrid1.TextMatrix(I, C_ChrgType) = "Chargeable" Then
                mChgHrs = Val(FGrid1.TextMatrix(I, C_Hrs))
                mChgRate = Val(FGrid1.TextMatrix(I, C_Rate))
                mWarHrs = 0
                mWarRate = 0
            Else
                mChgHrs = 0
                mChgRate = 0
                mChgAmt = 0
                mWarHrs = Val(FGrid1.TextMatrix(I, C_Hrs))
                mWarRate = Val(FGrid1.TextMatrix(I, C_Rate))
            End If
            mLabAmt = Val(FGrid1.TextMatrix(I, C_Amt))
            mExternal = IIf(FGrid1.TextMatrix(I, C_External) = "Yes", "1", "0")
            mTaxYN = IIf(FGrid1.TextMatrix(I, C_TaxYN) = "Yes", 1, 0)
            GSQL = "insert into Job_lab(" _
                & "Job_DocId,Site_Code,s_No,lab_code," _
                & "ActualHrs, hrs_taken,lab_rate,hrs_war,war_lab_rate,labourAmt," _
                & "major_yn,external_yn,Chrg_From,ExtJobGatePassNo," _
                & "Contract_Remarks, RateType, U_Name, U_EntDt, U_AE,Tax_YN,Chrg_Type,JobCode,Mech_Voice,Dep_Item , Dep_Code, DepitemPer, DepPer, DepAmt, InsuranceAmt,DiffPeried ) " _
                & " values(" _
                & "'" & lblDocId.CAPTION & "','" & PubSiteCode & "'," & SrNo & ",'" & FGrid1.TextMatrix(I, C_LabCode) & "'," _
                & "" & Val(FGrid1.TextMatrix(I, C_ActHrs)) & "," & mChgHrs & "," & mChgRate & "," & mWarHrs & ", " & mWarRate & "," & mLabAmt & "," _
                & "'" & FGrid1.TextMatrix(I, C_Major) & "','" & mExternal & "','" & left(FGrid1.TextMatrix(I, C_PaidBy), 1) & "','" & FGrid1.TextMatrix(I, C_GPNo) & "'," _
                & "'" & FGrid1.TextMatrix(I, C_Remarks) & "', '','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & ADDFLAG & "'," & mTaxYN & ",'" & left(FGrid1.TextMatrix(I, C_ChrgType), 1) & "','" & FGrid1.TextMatrix(I, C_JobCode) & "','" & FGrid1.TextMatrix(I, C_MechVoice) & "','" & FGrid1.TextMatrix(I, Col_DepItem) & "','" & FGrid1.TextMatrix(I, Col_DepCode) & "'," & Val(FGrid1.TextMatrix(I, Col_DepitemPer)) & "," & Val(FGrid1.TextMatrix(I, Col_DepPer)) & "," & Val(FGrid1.TextMatrix(I, Col_DepAmt)) & "," & Val(FGrid1.TextMatrix(I, Col_InsuranceAmt)) & "," & Val(FGrid1.TextMatrix(I, Col_DiffPeried)) & ")"
            GCn.Execute GSQL
            
            If TopCtrl1.TopText2 = "Add" Then
                GCn.Execute "Update Job_Lab Set AddBy = '" & pubUName & "', AddDate = " & ConvertDateTime(PubServerDate) & " Where Job_DocId = '" & lblDocId & "' "
            Else
                GCn.Execute "Update Job_Lab Set  AddBy = '" & mAddBy & "', AddDate = " & ConvertDateTime(mAddDate) & ", ModifyBy = '" & pubUName & "', ModifyDate = " & ConvertDateTime(PubServerDate) & " Where Job_DocId = '" & lblDocId & "' "
            End If
            
            '** Job_Lab2 for Multiple Mechanic
            For j = 1 To FGridLab2.Rows - 1
'                If TopCtrl1.TopText2 = "Add" Then
                    If FGridLab2.TextMatrix(j, LabCode) = FGrid1.TextMatrix(I, C_LabCode) Then
                        GSQL = "insert into Job_lab2(" _
                            & "Job_DocId,Site_Code,s_No,lab_code,mech_code," _
                            & "U_Name, U_EntDt, U_AE) " _
                            & " values('" & lblDocId.CAPTION & "','" & PubSiteCode & "'," & SrNo & ",'" & FGridLab2.TextMatrix(j, LabCode) & "','" & FGridLab2.TextMatrix(j, MechCode) & _
                            "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & ADDFLAG & "')"
                        GCn.Execute GSQL
                    End If
'                Else
'                    If FGridLab2.TextMatrix(J, LabCode) = FGrid1.TextMatrix(I, C_LabCode) And FGridLab2.RowData(J) = Val(FGrid1.TextMatrix(I, 0)) Then
'                        GSQL = "insert into Job_lab2(" _
'                            & "Job_DocId,Site_Code,s_No,lab_code,mech_code," _
'                            & "U_Name, U_EntDt, U_AE) " _
'                            & " values('" & lblDocId.CAPTION & "','" & PubSiteCode & "'," & SrNo & ",'" & FGridLab2.TextMatrix(J, LabCode) & "','" & FGridLab2.TextMatrix(J, MechCode) & _
'                            "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & ADDFLAG & "')"
'                        GCn.Execute GSQL
'                    End If
'                End If
            Next
            '********
            SrNo = SrNo + 1
        End If
    Next I
'    GCn.Execute ("Update Job_Card='" & IIf(Txt(RateSystemYN) = "Yes", 1, 0) & "' where DocID='" & lblDocId & "'")
    GCn.CommitTrans
    mTrans = False
    Set Rst = Nothing
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select Distinct JL.JOB_DOCID AS SEARCHCODE from Job_Lab as JL where left(JL.Job_DocId,1)='" & PubDivCode & "' And JL.JOB_DOCID ='" & lblDocId.CAPTION & "' order by JL.Job_DocID desc")
    End If
    Call UpdRequery
    
    Master.FIND "searchcode = '" & lblDocId.CAPTION & "'"
    If ADDFLAG = "A" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub

errlbl:
    If mTrans Then GCn.RollbackTrans
    Set Rst = Nothing
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus txt(Index)
    txtgrid1(0).Visible = False
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case JobNo
            DGridColSwap DGJob, 0
            RsJob.Sort = "FindJOBNO"
            If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).Tag <> RsJob!Code Then
                RsJob.MoveFirst
                RsJob.FIND ("Code='" & txt(Index).Tag & "'")
            End If
        Case Chassis
            DGridColSwap DGJob, 1
            RsJob.Sort = "CHASSIS"
            If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).Tag <> RsJob!Code Then
                RsJob.MoveFirst
                RsJob.FIND ("CHASSIS='" & txt(Index).TEXT & "'")
            End If
        Case VehRegNo
            DGridColSwap DGJob, 2
            RsJob.Sort = "REGNO"
            If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).Tag <> RsJob!Code Then
                RsJob.FIND ("REGNO='" & txt(Index).TEXT & "'")
            End If
            
'by lps
'        Case Model
'            DGridColSwap DGJob, 3
'            RsJob.Sort = "Model"
'            If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or Txt(Index).Text = "" Then Exit Sub
'            If Txt(Index).Tag <> RsJob!Code Then
'                RsJob.MoveFirst
'                RsJob.FIND ("MODEL='" & Txt(Index).Text & "'")
'            End If
'        Case VehSrlNo
'            DGridColSwap DGJob, 4
'            RsJob.Sort = "VehSerialNo"
'            If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or Txt(Index).Text = "" Then Exit Sub
'            If Txt(Index).Tag <> RsJob!Code Then
'                RsJob.MoveFirst
'                RsJob.FIND ("VEHSERIALNO='" & Txt(Index).Text & "'")
'            End If
'        Case OwnerName
'            DGridColSwap DGJob, 5
'            RsJob.Sort = "name"
'            If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or Txt(Index).Text = "" Then Exit Sub
'            If Txt(Index).Tag <> RsJob!Code Then
'                RsJob.MoveFirst
'                RsJob.FIND ("NAME='" & Txt(Index).Text & "'")
'            End If
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case JobNo
            DGridTxtKeyDown DGJob, txt, Index, RsJob, KeyCode, False, 1
        Case VehRegNo
            DGridTxtKeyDown DGJob, txt, Index, RsJob, KeyCode, False, 3
        Case Chassis
            DGridTxtKeyDown DGJob, txt, Index, RsJob, KeyCode, False, 4
'        Case Model 'by lps
'            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 2, frmModel
'        Case VehSrlNo  'by lps
'            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 6
'        Case OwnerName
'            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 7
    End Select
    If DGJob.Visible = False And DGMech.Visible = False And DGLabour.Visible = False And DGGatePass.Visible = False And DgRateType.Visible = False Then
        '' KEY DOWN
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then 'And Index <> Remarks Then
            Ctrl_DownKeyDown KeyCode, Shift
        End If
        ' KEY UP
'        If TopCtrl1.TopText2 = "Add" Then
        If ADDFLAG = "A" Then
            If Index <> JCType Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
    Select Case Index
        Case JobNo
            If DGJob.Visible = True Then DGridTxtKeyPress txt, JobNo, RsJob, KeyAscii, "FindJobNo"
        Case VehRegNo
            DGridTxtKeyPress txt, Index, RsJob, KeyAscii, "regno"
        Case Chassis
            DGridTxtKeyPress txt, Index, RsJob, KeyAscii, "chassis"
            
'        Case Model 'by lps
'            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "model"
'        Case VehSrlNo 'by lps
'            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "vehserialno"
'        Case OwnerName
'            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "name"
        Case RateSystemYN
            If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                txt(Index).TEXT = "Yes"
                KeyAscii = 0
            Else    'If KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                txt(Index).TEXT = "No"
                KeyAscii = 0
            End If
        Case JCType
            If KeyAscii = 89 Or KeyAscii = 121 Or KeyAscii = 78 Or KeyAscii = 110 Then
                If ((KeyAscii = 89 Or KeyAscii = 121) And txt(Index).TEXT = "No") Or ((KeyAscii = 78 Or KeyAscii = 110) And txt(Index).TEXT = "Yes") Then
                    Call BlankText
                End If
                If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                    txt(Index).TEXT = "Yes"
                    KeyAscii = 0
                ElseIf KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                    txt(Index).TEXT = "No"
                    KeyAscii = 0
                End If
            Else
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
    txt(Index).Tag = RsJob!Code
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rs As ADODB.Recordset
Dim I As Integer
    Select Case Index
        Case JobNo ', VehRegNo, Chassis
            If txt(Index).Tag <> "" Then
                RsJob.Sort = "CODE"
                RsJob.FIND ("CODE='" & txt(Index).Tag & "'")
            End If
            If RsJob.BOF = True Or RsJob.EOF = True Then Exit Sub
            lblDocId = RsJob!Code
            blnRateEditableYn = VNull(GCn.Execute("Select " & vIsNull("RateEditableYn", "0") & " From Service_Type Where Serv_Type = (Select Serv_Type From Job_Card Where DocId ='" & lblDocId & "')").Fields(0))
            txt(VehRegNo).Tag = XNull(RsJob!Code)
            Call History_Field
            Call Fill_Grid(RsJob!Code, RsJob!Model)
            
            If FGrid1.Rows >= 2 And FGrid1.TextMatrix(1, C_LabCode) = "" Then
                Set Rs = GCn.Execute("Select J.Lab_Code, L.Lab_Desc, L.Time_Req, L.Lab_Rate " & _
                                     "From (Job_Demand J " & _
                                     "Left Join Labour L On J.Lab_Code=L.Lab_Code) " & _
                                     "Where J.Job_DocId='" & txt(JobNo).Tag & "' And J.Lab_Code Is Not Null And J.Lab_Code <> ''")
                FGrid1.Rows = 1
                I = 1
                If Rs.RecordCount > 0 Then
                    Do Until Rs.EOF
                        FGrid1.AddItem ""
                        If XNull(Rs!Lab_Code) <> "" Then
                            
                            FGrid1.TextMatrix(I, C_LabCode) = Rs!Lab_Code
                            FGrid1.TextMatrix(I, C_LabName) = XNull(Rs!Lab_Desc)
                            FGrid1.TextMatrix(I, C_Fixed) = "No"
                            FGrid1.TextMatrix(I, C_TaxYN) = "Yes"
                            FGrid1.TextMatrix(I, C_PaidBy) = "Customer"
                            FGrid1.TextMatrix(I, C_ChrgType) = "Chargable"
                            FGrid1.TextMatrix(I, C_Rate) = VNull(Rs!Lab_Rate)
                            FGrid1.TextMatrix(I, C_Hrs) = VNull(Rs!TIME_REQ)
                            FGrid1.TextMatrix(I, C_Amt) = Format(VNull(Rs!TIME_REQ) * VNull(Rs!Lab_Rate), "0.00")
                            
                            I = I + 1
                        End If
                        
                        
                        Rs.MoveNext
                    Loop
                    FGrid1.AddItem ""
                    FGrid1.FixedRows = 1
                Else
                    FGrid1.AddItem ""
                    FGrid1.FixedRows = 1
                End If
            End If
            If txt(JobNo) <> "" Then
                txt(VehRegNo).Enabled = False
            Else
                txt(VehRegNo).Enabled = True
            End If
        Case JCType
            If txt(JCType).TEXT = "Yes" Then
                RsJob.Filter = ("JobCloseDate=null")
            Else
                RsJob.Filter = ("")
            End If
            
    End Select
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    For I = 0 To txt.Count - 1
        txt(I).TEXT = ""
        txt(I).Tag = ""
    Next I
    
    
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
    
    
    GridSel(1).Rows = 1
    GridSel(1).AddItem GridSel(1).Rows
    GridSel(1).FixedRows = 1
        
    lblDocId.CAPTION = ""
    lblDocId.Refresh
End Sub

Private Sub MoveRec()
Dim Master1 As ADODB.Recordset
On Error GoTo error1
    If Master.RecordCount > 0 Then
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "select J.Coupon,J.Coupon_Value,JL.Job_DocId, JL.JOB_DOCID AS SEARCHCODE, JL.SITE_CODE,J.Job_No, J.Job_Date, J.JobCloseDate, J.CARDNO,J.Govt_YN, HC.Model,HC.RegNo, HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name, HC.Add1, HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc, City.CityName,JL.JobCode,JL.Mech_Voice, Jl.AddBy, Jl.AddDate, Jl.ModifyBy, Jl.ModifyDate, JL.RateType, RateType.Description As RateTypeDescription, RateType.VariationPer ,hc.Delivery_Date as SOldDate  " & _
            " from ((((Job_Lab as JL Left Join job_card as J on JL.Job_DocId=J.DocId) " & _
            " left Join Hiscard as HC on J.CardNo=HC.CardNo) " & _
            " left Join Service_Type as ST on J.Serv_Type=ST.Serv_Type) " & _
            " Left Join City on HC.CityCode=City.CityCode) " & _
            " Left Join RateType on JL.RateType=RateType.Code " & _
            " where J.DocId='" & Master!SearchCode & "' Order by JL.S_No", GCn, adOpenStatic, adLockReadOnly
        
        LblDiv.CAPTION = "Division : " & left(Master1!job_docid, 1)
        LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        lblDocId.CAPTION = Master1!job_docid
        txt(JobNo).TEXT = Master1!Job_No
        txt(JobDt).TEXT = Master1!Job_Date
        txt(JobCDt).TEXT = IIf(Master1!JobCloseDate = #1/1/1900# Or IsNull(Master1!JobCloseDate), "", Master1!JobCloseDate)
                
        blnRateEditableYn = VNull(GCn.Execute("Select " & vIsNull("RateEditableYn", "0") & " From Service_Type Where Serv_Type = (Select Serv_Type From Job_Card Where DocId ='" & lblDocId & "')").Fields(0))
        txt(SoldDate).TEXT = IIf(IsNull(Master1!SoldDate), "", Master1!SoldDate)
        
        txt(SrvType).TEXT = XNull(Master1!Serv_Desc)
        txt(HistNo).TEXT = Master1!CardNo
        txt(HistNo).Tag = Master1!CardNo
        txt(Coupon).TEXT = XNull(Master1!Coupon)
        txt(CouponValue).Tag = IIf(IsNull(Master1!Coupon_Value), "", Master1!Coupon_Value)
        
        txt(VehRegNo).TEXT = XNull(Master1!RegNo)
        txt(Chassis).TEXT = XNull(Master1!Chassis)
        txt(Model).TEXT = XNull(Master1!Model)
        txt(Engine).TEXT = XNull(Master1!Engine)
        txt(VehSrlNo).TEXT = XNull(Master1!VehSerialNo)
        txt(GovtYn).TEXT = IIf(Master1!Govt_YN = 1, "Yes", "No ")
        txt(OwnerName).TEXT = XNull(Master1!Name)
        txt(City).TEXT = XNull(Master1!CityName)
'        Txt(RateSystemYN).Text = IIf(Master1!ApplyLabRateYN = 1, "Yes", "No")
        
        
        
        LblUser = IIf(Not IsNull(Master1!AddDate), "Add By : " & XNull(Master1!AddBy) & "  Dated : " & XNull(Master1!AddDate), "") & IIf(Not IsNull(Master1!ModifyDate), "     Modify By : " & XNull(Master1!ModifyBy) & "  Dated : " & XNull(Master1!ModifyDate), "")
        mAddBy = XNull(Master1!AddBy)
        mAddDate = XNull(Master1!AddDate)
        
        UpdLastJC
        Call Fill_Grid(Master!SearchCode, XNull(Master1!Model))
        Call veh_count
    End If
    Grid_Hide
    FGrid1_GotFocus
    Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
    With FGrid1
        .width = Me.width - 120
        .left = 0
        .top = 3635
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 29
        .height = .RowHeight(0) * 5
        '.AllowUserResizing = flexResizeNone
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 400
        
        .TextMatrix(0, C_LabCode) = "Lab.Code"
        .ColAlignment(C_LabCode) = flexAlignLeftCenter
        .ColWidth(C_LabCode) = 780
        
        .TextMatrix(0, C_LabName) = "Labour Description"
        .ColAlignment(C_LabName) = flexAlignLeftCenter
        .ColWidth(C_LabName) = 2500
        
        .TextMatrix(0, C_MechVoice) = "Mech.Voice"
        .ColAlignment(C_MechVoice) = flexAlignLeftCenter
        .ColWidth(C_MechVoice) = 2200
        
        .TextMatrix(0, C_Fixed) = "Fixed"
        .ColAlignment(C_Fixed) = flexAlignLeftCenter
        .ColWidth(C_Fixed) = 465
        
        .TextMatrix(0, C_TaxYN) = "Tax YN"
        .ColAlignment(C_TaxYN) = flexAlignLeftCenter
        .ColWidth(C_TaxYN) = 650
        
        .TextMatrix(0, C_PaidBy) = "PaidBy"
        .ColAlignment(C_PaidBy) = flexAlignLeftCenter
        .ColWidth(C_PaidBy) = 585

        .TextMatrix(0, C_ChrgType) = "Type"
        .ColAlignmentFixed(C_ChrgType) = flexAlignCenterCenter
        .ColAlignment(C_ChrgType) = flexAlignLeftCenter
        .ColWidth(C_ChrgType) = 885


        .TextMatrix(0, C_ActHrs) = "Act.Hrs"
        .ColAlignmentFixed(C_ActHrs) = flexAlignCenterCenter
        .ColAlignment(C_ActHrs) = flexAlignRightCenter
        .ColWidth(C_ActHrs) = 700

        .TextMatrix(0, C_Hrs) = "Hrs Minut" 'Ch.Hrs." '
        .ColAlignmentFixed(C_Hrs) = flexAlignCenterCenter
        .ColAlignment(C_Hrs) = flexAlignRightCenter
        .ColWidth(C_Hrs) = 700

        .TextMatrix(0, C_Rate) = "Rate/Hr" 'Ch.Amt." '
        .ColAlignmentFixed(C_Rate) = flexAlignCenterCenter
        .ColAlignment(C_Rate) = flexAlignRightCenter
        .ColWidth(C_Rate) = 700

        .TextMatrix(0, C_Amt) = "Amount" 'Ch.Amt." '
        .ColAlignmentFixed(C_Amt) = flexAlignCenterCenter
        .ColAlignment(C_Amt) = flexAlignRightCenter
        .ColWidth(C_Amt) = 795

        .TextMatrix(0, C_External) = "Extl"
        .ColAlignment(C_External) = flexAlignLeftCenter
        .ColWidth(C_External) = 400

        .TextMatrix(0, C_GPNo) = "GP No."
        .ColAlignment(C_GPNo) = flexAlignLeftCenter
        .ColWidth(C_GPNo) = 800
        .TextMatrix(0, C_Remarks) = "Remarks"
        .ColAlignment(C_Remarks) = flexAlignLeftCenter
        .ColWidth(C_Remarks) = 1815

        .TextMatrix(0, C_ContName) = "Contractor Name"
        .ColAlignment(C_ContName) = flexAlignLeftCenter
        .ColWidth(C_ContName) = 2280

        .TextMatrix(0, C_WIssueDt) = "Issue Dt."
        .ColAlignment(C_WIssueDt) = flexAlignRightCenter
        .ColWidth(C_WIssueDt) = 1100

        .TextMatrix(0, C_WRecdDt) = "Recd. Dt."
        .ColAlignment(C_WRecdDt) = flexAlignRightCenter
        .ColWidth(C_WRecdDt) = 1100

        .TextMatrix(0, C_ContAmt) = "Cont. Amt."
        .ColAlignment(C_ContAmt) = flexAlignRightCenter
        .ColWidth(C_ContAmt) = 950

        .TextMatrix(0, C_ContCode) = "Contractor Code"
        .ColAlignment(C_ContCode) = flexAlignLeftCenter
        .ColWidth(C_ContCode) = 0

        .TextMatrix(0, C_Major) = "Major Y/N"
        .ColAlignment(C_Major) = flexAlignLeftCenter
        .ColWidth(C_Major) = 0
        
        .TextMatrix(0, C_JobCode) = "JobCode"
        .ColAlignment(C_JobCode) = flexAlignLeftCenter
        .ColWidth(C_JobCode) = 1500
        
        
        .TextMatrix(0, Col_DepItem) = "Deprecation Item"
        .ColWidth(Col_DepItem) = 0

        .TextMatrix(0, Col_DepitemPer) = "Deprecation Item Per"
        .ColAlignment(Col_DepitemPer) = flexAlignLeftCenter
        .ColWidth(Col_DepitemPer) = 1000

        .TextMatrix(0, Col_DepCode) = "Deprecation Code"
        .ColAlignment(Col_DepCode) = flexAlignLeftCenter
        .ColWidth(Col_DepCode) = 0


        .TextMatrix(0, Col_DepPer) = "Deprecation Per"
        .ColAlignment(Col_DepPer) = flexAlignLeftCenter
        .ColWidth(Col_DepPer) = 1000


        .TextMatrix(0, Col_DepAmt) = "Deprecation Amt"
        .ColAlignment(Col_DepAmt) = flexAlignLeftCenter
        .ColWidth(Col_DepAmt) = 1000

       .TextMatrix(0, Col_InsuranceAmt) = "Insurance Amt"
        .ColAlignment(Col_InsuranceAmt) = flexAlignLeftCenter
        .ColWidth(Col_InsuranceAmt) = 1000

        .TextMatrix(0, Col_DiffPeried) = "Diffrence Peried"
        .ColAlignment(Col_DiffPeried) = flexAlignLeftCenter
        .ColWidth(Col_DiffPeried) = 0



        
    End With
    BackColorSelLeave = FGrid1.BackColorSel
    ForeColorSelEnter = FGrid1.ForeColorSel

    With GridSel(1)
        '.width = Me.width - 120
        '.left = 0
        '.top = 3465
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 5
        .Rows = 6
        .height = (.RowHeight(0) * 6) + 15
        .AllowUserResizing = flexResizeNone
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 400
        
        .TextMatrix(0, Srlno) = "SrlNo"
        .ColAlignment(Srlno) = flexAlignLeftCenter
        .ColWidth(Srlno) = 0
        
        .TextMatrix(0, LabCode) = "LabCode"
        .ColAlignment(LabCode) = flexAlignLeftCenter
        .ColWidth(LabCode) = 0
        
        .TextMatrix(0, MechCode) = "Mech Code"
        .ColAlignment(MechCode) = flexAlignLeftCenter
        .ColWidth(MechCode) = 0
        
        .TextMatrix(0, MechName) = "Mechanic Name"
        .ColAlignment(MechName) = flexAlignLeftCenter
        .ColWidth(MechName) = 3000
    End With
'    ReDim Preserve GridRow1(1)
'    GridRow1(1) = 0

    DGJob.width = Me.width - 60: DGJob.left = FGrid1.left: DGJob.top = FGrid1.top: DGJob.height = Me.height - (DGJob.top + mBotScale)
    DGLabour.width = 7000:
    DGLabour.left = Me.width - (DGLabour.width + mRtScale): DGLabour.top = mTopScale: DGLabour.height = Me.height - (DGLabour.top + mBotScale)
    DGMech.width = DGLabour.width: DGMech.left = Me.width - (DGMech.width + mRtScale): DGMech.top = mTopScale: DGMech.height = FGrid1.top - mTopScale
    DGGatePass.width = Me.width - 60: DGGatePass.left = FGrid1.left: DGGatePass.top = mTopScale: DGGatePass.height = FGrid1.top - mTopScale
    DGJobCode.left = Me.width - (DGJobCode.width + mRtScale): DGJobCode.top = mTopScale: DGJobCode.height = Me.height / 2 - 500
End Sub

Public Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    ADDFLAG = left(TopCtrl1.TopText2, 1)
    For I = 0 To txt.Count - 1
        txt(I).Enabled = Enb
    Next
    
    For I = 0 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg
        txt(I).ForeColor = CtrlFColOrg
    Next
    
    txtgrid1(0).BackColor = CtrlBCol
    txtgrid1(0).ForeColor = CtrlFCol
    txtgrid1(0).Enabled = Enb
    txt(JobCDt).Enabled = False
    txt(JobDt).Enabled = False
    txt(VehRegNo).Enabled = False
    txt(Chassis).Enabled = False
    txt(Engine).Enabled = False
    txt(SrvType).Enabled = False
    txt(City).Enabled = False
    txt(GovtYn).Enabled = False
    
    txt(LastJobDt).Enabled = False
    txt(LastJobNo).Enabled = False
    txt(LastSrv).Enabled = False
    txt(LastKMS).Enabled = False
    txt(LastMech).Enabled = False
    txt(HistNo).Enabled = False

    txt(ChgHrs).Enabled = False
    txt(ChgAmt).Enabled = False
    txt(WarrAmt).Enabled = False
    txt(WarrHrs).Enabled = False
'    Txt(TotAmt).Enabled = False
'    Txt(TotHrs).Enabled = False
    txt(ExtLab).Enabled = False
    txt(ExtLabChg).Enabled = False
    
    
End Sub

Private Sub Grid_Hide()
    If DGJob.Visible = True Then DGJob.Visible = False
    If DGGatePass.Visible = True Then DGGatePass.Visible = False
    If DGMech.Visible = True Then DGMech.Visible = False
    If DGLabour.Visible = True Then DGLabour.Visible = False
End Sub

Private Sub veh_count()
    If txt(JobDt).TEXT <> "" Then
        LblTotVeh.CAPTION = GCn.Execute("select count(*) from job_Card where JobCloseDate = " & ConvertDate("01/Jan/1900") & " or JobCloseDate Is Null and left(Docid,1)='" & PubDivCode & "'").Fields(0)
    End If
End Sub

Private Sub UpdRequery()
    RsJob.Requery
    RsLab.Requery
    DGLabour.Refresh
    RsJobCode.Requery
    RsGatePass.Requery
End Sub

Private Sub History_Field()
On Error GoTo ErrLoop
Dim rsJob2 As Recordset
    Set rsJob2 = GCn.Execute("select J.Job_Date,J.Govt_YN,City.CityName,ST.Serv_Desc,hc.Delivery_Date as Solddate " _
        & "from (((job_card as J left Join Hiscard as HC on J.CardNo=HC.CardNo) " _
        & "left Join Service_Type as ST on J.Serv_Type=ST.Serv_Type) " _
        & "Left Join City on HC.CityCode=City.CityCode) " _
        & "where J.DocId='" & RsJob!Code & "' ")
    
    txt(HistNo).Tag = RsJob!CardNo
    txt(HistNo).TEXT = RsJob!CardNo
    
    txt(VehRegNo).Tag = RsJob!Code
    txt(Chassis).Tag = RsJob!Code
    txt(Model).Tag = RsJob!Code
    txt(VehSrlNo).Tag = RsJob!Code
    txt(OwnerName).Tag = RsJob!Code
    txt(JobNo).Tag = RsJob!Code
    
    txt(JobNo).TEXT = RsJob!FindJobNo
    txt(JobDt).TEXT = rsJob2!Job_Date
    txt(JobCDt).TEXT = IIf(RsJob!JobCloseDate = #1/1/1900# Or IsNull(RsJob!JobCloseDate), "", RsJob!JobCloseDate)
    txt(VehRegNo).TEXT = IIf(IsNull(RsJob!RegNo), "", RsJob!RegNo)
    txt(Chassis).TEXT = IIf(IsNull(RsJob!Chassis), "", RsJob!Chassis)
    txt(Model).TEXT = IIf(IsNull(RsJob!Model), "", RsJob!Model)
    txt(Engine).TEXT = IIf(IsNull(RsJob!Engine), "", RsJob!Engine)
    txt(VehSrlNo).TEXT = IIf(IsNull(RsJob!VehSerialNo), "", RsJob!VehSerialNo)
    txt(GovtYn).TEXT = IIf(rsJob2!Govt_YN = 0, "No", "Yes")
    txt(OwnerName).TEXT = XNull(RsJob!Name)
    txt(City).TEXT = XNull(rsJob2!CityName)
        txt(SoldDate).TEXT = IIf(IsNull(rsJob2!SoldDate), "", rsJob2!SoldDate)
    Set rsJob2 = Nothing
    Call UpdLastJC
    Exit Sub
ErrLoop:
    Set rsJob2 = Nothing
    CheckError
End Sub

Private Sub UpdLastJC()
    If txt(JobDt) = "" Then Exit Sub
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    RsTemp.CursorLocation = adUseClient
    RsTemp.Open "SELECT Top 1 JOB_NO,JOB_DATE,AtKMsHrs,Srv.Serv_SrlNo,Srv.Serv_Type,Srv.SERV_DESC AS SrvDesc,EMP_MAST.EMP_NAME AS MECH_NAME " & _
            " FROM (JOB_CARD LEFT JOIN Service_Type Srv ON JOB_CARD.SERV_TYPE=Srv.SERV_TYPE) " & _
            " LEFT JOIN EMP_MAST ON JOB_CARD.RECBY_MECHANIC=EMP_MAST.EMP_CODE " & _
            " WHERE CARDNO='" & txt(HistNo).TEXT & _
            "' and Job_Date< " & ConvertDate(txt(JobDt)) & _
            " ORDER BY JOB_DATE Desc ", GCn, adOpenStatic, adLockReadOnly
    If RsTemp.RecordCount > 0 Then
        txt(LastJobNo).TEXT = RsTemp!Job_No
        txt(LastJobDt).TEXT = RsTemp!Job_Date
        txt(LastKMS).TEXT = RsTemp!AtKMsHrs
        txt(LastSrv).TEXT = RsTemp!SrvDesc
        txt(LastSrv).Tag = RsTemp!Serv_SrlNo
        txt(LastMech).TEXT = IIf(IsNull(RsTemp!MECH_NAME), "*No Mechanic*", RsTemp!MECH_NAME)
    Else
        txt(LastJobNo).TEXT = "":           txt(LastJobDt).TEXT = ""
        txt(LastKMS).TEXT = "":             txt(LastSrv).TEXT = ""
        txt(LastMech).TEXT = "":            txt(LastSrv).Tag = ""
    End If
    Set RsTemp = Nothing
End Sub

Private Sub FGrid1_Click()
    If ADDFLAG = "B" Then Exit Sub
End Sub
Private Sub FGrid1_DblClick()
On Error GoTo ELoop
FGrid1_KeyPress vbKeyReturn
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub FGrid1_GotFocus()
    FGrid1.BackColorSel = BackColorSelEnter
    FGrid1.ForeColorSel = ForeColorSelEnter
    txtgrid1(0).Visible = False
    Grid_Hide
    If ADDFLAG = "A" Then
        txt(VehRegNo).Enabled = False
        txt(Chassis).Enabled = False
        txt(Model).Enabled = False
        txt(VehSrlNo).Enabled = False
        txt(OwnerName).Enabled = False
        txt(JobNo).Enabled = False
    End If
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If ADDFLAG = "B" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
    If ADDFLAG <> "E" Then SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
    If MsgBox("Save Entry ?", vbInformation + vbYesNo, "Save Data") = vbYes Then
        TopCtrl1_eSave
        Exit Sub
    Else
        FGrid1.Tag = FGrid1.Row
    End If
    FGrid1.SetFocus
    KeyCode = 0
End If
GridKey = KeyCode
FGrid1.Tag = FGrid1.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    If txt(JobCDt) = "" Then        ' Exit Sub
        Select Case FGrid1.Col
            Case C_Hrs, C_Rate
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
                TotHrAmt
            Case C_Amt
                If FGrid1.TextMatrix(FGrid1.Row, C_Fixed) = "Yes" Then Exit Sub
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
                TotHrAmt
            Case C_MechVoice
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
        End Select
    End If
    Select Case FGrid1.Col
        Case C_GPNo
            FGrid1.TextMatrix(FGrid1.Row, C_GPNo) = ""
            FGrid1.TextMatrix(FGrid1.Row, C_ContName) = ""
            FGrid1.TextMatrix(FGrid1.Row, C_ContCode) = ""
            FGrid1.TextMatrix(FGrid1.Row, C_WIssueDt) = ""
            FGrid1.TextMatrix(FGrid1.Row, C_WRecdDt) = ""
            FGrid1.TextMatrix(FGrid1.Row, C_ContAmt) = ""
            FGrid1.TextMatrix(FGrid1.Row, C_Remarks) = ""
        Case C_Remarks
            FGrid1.TextMatrix(FGrid1.Row, C_Remarks) = ""
    End Select
End If
KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub FGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
SetMaxLength
    
    If FGrid1.TextMatrix(FGrid1.Row, C_LabCode) <> "" Then
        Select Case FGrid1.Col
            Case C_Fixed ',C_PaidBy
                FGrid1.Col = C_ChrgType
            Case C_PaidBy
                Get_Text Me, FGrid1, txtgrid1, 0, False, KeyAscii
            Case C_ChrgType
                If UCase(Chr(KeyAscii)) = "W" Then
                    FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Warranty"
                ElseIf UCase(Chr(KeyAscii)) = "F" Then
                    FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Free Service"
                ElseIf UCase(Chr(KeyAscii)) = "A" Then
                    FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "AMC"
                ElseIf UCase(Chr(KeyAscii)) = "P" Then
                    FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "PDI"
                Else 'If UCase(Chr(KeyAscii)) = "Y" Then
                    If FGrid1.TextMatrix(FGrid1.Row, C_PaidBy) = "Customer" Then
                        FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Chargeable"
                    Else
                        MsgBox "Chargeable Can Be Set with only Customer."
                        FGrid1.SetFocus
                    End If
                End If
                KeyAscii = 0
                FGrid1.Col = FGrid1.Col + 1
                TotHrAmt
            Case C_ActHrs
                Get_Text Me, FGrid1, txtgrid1, 0, True, KeyAscii
            Case C_Hrs
                If blnRateEditableYn Then
                    Get_Text Me, FGrid1, txtgrid1, 0, True, KeyAscii
                Else
                    FGrid1.Col = C_External
                    GridSel(1).Col = 0: GridSel(1).Row = 1
                    GridSel(1).SetFocus
                End If
                
            Case C_Rate
               ' If Val(FGrid1.TextMatrix(FGrid1.Row, C_Amt)) <= 0 Then
                    If blnRateEditableYn Then
                        Get_Text Me, FGrid1, txtgrid1, 0, True, KeyAscii
                    Else
                        FGrid1.Col = C_External
                        GridSel(1).Col = 0: GridSel(1).Row = 1
                        GridSel(1).SetFocus
                    End If
               ' Else
               '     KeyAscii = 0
               '     FGrid1.Col = FGrid1.Col + 1
               'End If
            Case C_MechVoice
                    Get_Text Me, FGrid1, txtgrid1, 0, True, KeyAscii
            Case C_Amt
                If FGrid1.TextMatrix(FGrid1.Row, C_Fixed) = "Yes" Then Exit Sub
                If Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)) <= 0 Then
                    Get_Text Me, FGrid1, txtgrid1, 0, True, KeyAscii
                Else
                    KeyAscii = 0
                    FGrid1.Col = C_Rate
                End If
            Case C_TaxYN
                If UCase(Chr(KeyAscii)) = "Y" And ApplyServTax = 1 Then
                    FGrid1.TextMatrix(FGrid1.Row, C_TaxYN) = "Yes"
                Else
                    FGrid1.TextMatrix(FGrid1.Row, C_TaxYN) = "No"
                End If
                KeyAscii = 0
                FGrid1.Col = FGrid1.Col + 1
                
            Case C_External
                If FGrid1.TextMatrix(FGrid1.Row, C_External) = "Yes" And FGrid1.TextMatrix(FGrid1.Row, C_GPNo) <> "" Then
                    KeyAscii = 0
                    MsgBox "Please Delete GP No. first!", vbCritical, "Validation"
                    FGrid1.SetFocus
                Else
                    If UCase(Chr(KeyAscii)) = "Y" Then
                        FGrid1.TextMatrix(FGrid1.Row, C_External) = "Yes"
                        FGrid1.Col = FGrid1.Col + 1
                    Else 'If UCase(Chr(KeyAscii)) = "Y" Then
                        FGrid1.TextMatrix(FGrid1.Row, C_External) = "No"
                        FGrid1.Row = FGrid1.Row + 1
                        FGrid1.Col = C_LabCode
                    End If
                End If
                KeyAscii = 0
            Case C_GPNo
                If FGrid1.TextMatrix(FGrid1.Row, C_External) = "Yes" Then
                    RsGatePass.Filter = ("Job_DocID='" & lblDocId & "'")
                    If RsGatePass.RecordCount = 1 Then
                        FGrid1.TextMatrix(FGrid1.Row, C_GPNo) = RsGatePass!Code
                        FGrid1.TextMatrix(FGrid1.Row, C_ContCode) = RsGatePass!ContractCode
                        FGrid1.TextMatrix(FGrid1.Row, C_ContName) = RsGatePass!FinName
                        FGrid1.TextMatrix(FGrid1.Row, C_WIssueDt) = RsGatePass!GatePassDate
                        FGrid1.TextMatrix(FGrid1.Row, C_WRecdDt) = IIf(IsNull(RsGatePass!ContractRecdDate), "", RsGatePass!ContractRecdDate)
                    Else
                        Get_Text Me, FGrid1, txtgrid1, 0, False, KeyAscii
                    End If
                End If
            Case C_Remarks
                If FGrid1.TextMatrix(FGrid1.Row, C_External) = "Yes" Then
                    Get_Text Me, FGrid1, txtgrid1, 0, False, KeyAscii
                End If
            Case C_ContName, C_WIssueDt, C_WRecdDt, C_ContAmt
                FGrid1.Col = C_Remarks
            Case C_JobCode
                Get_Text Me, FGrid1, txtgrid1, 0, False, KeyAscii
        End Select
    End If
    If FGrid1.TextMatrix(FGrid1.Row, C_Amt) = "" Or UCase(left(PubComp_Name, 3)) = "LMP" Then
        Select Case FGrid1.Col
            Case C_LabCode
                Get_Text Me, FGrid1, txtgrid1, 0, False, KeyAscii
            Case C_LabName
                If FGrid1.TextMatrix(FGrid1.Row, C_LabCode) = "" Then
                    Get_Text Me, FGrid1, txtgrid1, 0, False, KeyAscii
                Else
                    FGrid1.Col = C_ChrgType 'FGrid1.Col + 4
                End If
'            Case Else
'                FGrid1.Col = C_LabCode
        End Select
    End If
    If KeyAscii <> vbKeyReturn Then TAddMode = True
    Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer, j As Integer
On Error GoTo ELoop
    If ADDFLAG = "B" Then Exit Sub
    If FGrid1.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If txt(JobCDt).TEXT = "" Then
            If FGrid1.Row >= 1 Then
                If MsgBox("Are You Sure To Delete Entry ?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                    '** Job_Lab2 for Multiple Mechanic
                    FGridLab2Copy.Rows = 1
                    For I = 1 To FGridLab2.Rows - 1
                        If FGridLab2.TextMatrix(I, LabCode) <> FGrid1.TextMatrix(FGrid1.Row, C_LabCode) Then
                            FGridLab2Copy.AddItem ""
                            With FGridLab2Copy
                                .TextMatrix(FGridLab2Copy.Rows - 1, LabCode) = FGridLab2.TextMatrix(I, LabCode)
                                .TextMatrix(FGridLab2Copy.Rows - 1, MechCode) = FGridLab2.TextMatrix(I, MechCode)
                                .TextMatrix(FGridLab2Copy.Rows - 1, MechName) = FGridLab2.TextMatrix(I, MechName)
                            End With
                        End If
                    Next
                    '**
                    FGridLab2.Rows = 1
                    For I = 1 To FGridLab2Copy.Rows - 1
                        FGridLab2.AddItem ""
                        With FGridLab2
                            .TextMatrix(FGridLab2.Rows - 1, LabCode) = FGridLab2Copy.TextMatrix(I, LabCode)
                            .TextMatrix(FGridLab2.Rows - 1, MechCode) = FGridLab2Copy.TextMatrix(I, MechCode)
                            .TextMatrix(FGridLab2.Rows - 1, MechName) = FGridLab2Copy.TextMatrix(I, MechName)
                        End With
                    Next
                    '**
                    FGridLab2Copy.Rows = 1
                    'Delete Rows
                    If FGrid1.Rows > 2 Then
                        FGrid1.RemoveItem (FGrid1.Row)
                        FillMechGrid FGrid1.Row
                    Else
                        FGrid1.Rows = 1
                        FGrid1.AddItem FGrid1.Rows
                        FGrid1.FixedRows = 1
                        FGridLab2.Rows = 1: FGridLab2.AddItem "": FGridLab2.FixedRows = 1
                        GridSel(1).Rows = 1: GridSel(1).AddItem "": GridSel(1).FixedRows = 1
                    End If
                End If
                For I = 1 To FGrid1.Rows - 1
                    FGrid1.TextMatrix(I, 0) = I
                Next
                FGrid1_RowColChange
                TotHrAmt
            Else
                MsgBox "No Entries To Delete", vbCritical, "Delete Module"
            End If
        Else
            MsgBox "JobCard is Closed, You can't Delete Labour", vbInformation, "Validation"
        End If
        FGrid1.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_Scroll()
    txtgrid1(0).Visible = False
    DGLabour.Visible = False
    DGMech.Visible = False
    DGGatePass.Visible = False
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
On Error GoTo ELoop
If FGrid1.TextMatrix(FGrid1.Row, C_External) = "No" Then
    If FGrid1.Col = C_ContAmt Or FGrid1.Col = C_ContName Or FGrid1.Col = C_Remarks Or FGrid1.Col = C_WIssueDt Or FGrid1.Col = C_WRecdDt Then Exit Sub
End If
Grid_Hide
txtgrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
txtgrid1(0).MaxLength = 0
Select Case FGrid1.Col
    Case C_LabCode
        If RsLab.EOF = True Or RsLab.BOF = True Or txtgrid1(Index).TEXT = "" Then Exit Sub
        RsLab.MoveFirst
        RsLab.Sort = "Code"
        RsLab.FIND "Code='" & FGrid1.TextMatrix(FGrid1.Row, C_LabCode) & "'"
    Case C_LabName
        If RsLab.EOF = True Or RsLab.BOF = True Or txtgrid1(Index).TEXT = "" Then Exit Sub
        RsLab.MoveFirst
        RsLab.Sort = "name"
        RsLab.FIND "name='" & FGrid1.TextMatrix(FGrid1.Row, C_LabName) & "'"
    Case C_PaidBy
        ListArray = Array("Customer", "Manufacturer", "Self", "Other Dealer")
        Set mListItem = ListView_Items(ListView, txtgrid1, Index, ListArray, 4)
    Case C_GPNo
        If RsGatePass.EOF = True Or RsGatePass.BOF = True Or txtgrid1(Index).TEXT = "" Then Exit Sub
        RsGatePass.MoveFirst
        RsGatePass.FIND "Code='" & FGrid1.TextMatrix(FGrid1.Row, C_GPNo) & "'"
    Case C_JobCode
        If RsJobCode.EOF = True Or RsJobCode.BOF = True Or txtgrid1(Index).TEXT = "" Then Exit Sub
        RsJobCode.MoveFirst
        RsJobCode.Sort = "Code"
        RsJobCode.FIND "Code='" & FGrid1.TextMatrix(FGrid1.Row, C_JobCode) & "'"
    Case C_MechVoice
        txtgrid1(0).MaxLength = 40
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then txtgrid1(0).TEXT = txtgrid1(0).Tag: Exit Sub
    Select Case FGrid1.Col
        Case C_LabCode
            If DGLabour.Visible = False Then DGridColSwap DGLabour, 0
            DGridTxtKeyDown DGLabour, txtgrid1, 0, RsLab, KeyCode, True, 0, frmLabDesc, "frmLabDesc"
            If KeyCode = vbKeyReturn Then
                If TxtGrid1Leave Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks, 4
                End If
            End If
        Case C_LabName
            If DGLabour.Visible = False Then DGridColSwap DGLabour, 1
            DGridTxtKeyDown DGLabour, txtgrid1, 0, RsLab, KeyCode, True, 1, frmLabDesc, "frmLabDesc"
            If KeyCode = vbKeyReturn Then
                If TxtGrid1Leave Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks, 3
                End If
            End If
        Case C_PaidBy
            ListView_KeyDown FrmList, ListView, txtgrid1, 0, KeyCode, Shift, 5400, 2500, 1400, 1200
            If KeyCode = vbKeyReturn Then
                If TxtGrid1Leave = True Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks
                End If
            End If
        Case C_Hrs  'lps   , C_WarAmt, C_WarHrs
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If LabTimeChk(C_Hrs) Or UCase(left(PubComp_Name, 3)) = "LMP" Then
                    If TxtGrid1Leave Then
                        GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks, 1  'Skip to Amount Col
                    End If
                Else
                    MsgBox "Invalid Time Entry.Please Enter Hrs.Minuts", vbInformation
                    Exit Sub
                End If
                If FGrid1.Col = C_LabCode Then FGrid1.Col = C_LabName
            End If
        Case C_ActHrs
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If LabTimeChk(C_ActHrs) Or UCase(left(PubComp_Name, 3)) = "LMP" Then
                    If TxtGrid1Leave Then
                        GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks, 1  'Skip to Amount Col
                    End If
                Else
                    MsgBox "Invalid Time Entry.Please Enter Hrs.Minuts", vbInformation
                    Exit Sub
                End If
            End If
            
        Case C_Rate 'for Kota 23-07-2003
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks, 1
'                    If AddFlag <> "A" Then
                        GridSel(1).Col = 0: GridSel(1).Row = 1
                        GridSel(1).SetFocus
'                    End If
                End If
                If FGrid1.Col = C_LabCode Then FGrid1.Col = C_LabName
            End If
        
        Case C_Amt
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode) Then
                If TxtGrid1Leave Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks
'                    If AddFlag <> "A" Then
                        GridSel(1).Col = 0: GridSel(1).Row = 1
                        GridSel(1).SetFocus
'                    End If
                End If
                If FGrid1.Col = C_LabCode Then FGrid1.Col = C_LabName
            End If
        Case C_MechVoice
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode) Then
                If TxtGrid1Leave Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks
                End If
            End If
        Case C_GPNo
            DGridTxtKeyDown DGGatePass, txtgrid1, 0, RsGatePass, KeyCode, True, 1
            If KeyCode = vbKeyReturn Then
                If TxtGrid1Leave = True Then
                    If FGrid1.TextMatrix(FGrid1.Row, C_External) = "Yes" Then
                        GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks
                    End If
                End If
            End If
        Case C_Remarks
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave = True Then
                    If FGrid1.TextMatrix(FGrid1.Row, C_External) = "Yes" Then
                        GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks
                    End If
                End If
                If FGrid1.Col = C_LabCode Then FGrid1.Col = C_LabName
            End If
        Case C_JobCode
            If DGJobCode.Visible = False Then DGridColSwap DGJobCode, 0
            DGridTxtKeyDown DGJobCode, txtgrid1, 0, RsJobCode, KeyCode, True, 0
            If KeyCode = vbKeyReturn Then
                If TxtGrid1Leave Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks, 4
                End If
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Exit Sub
On Error GoTo ELoop
    CheckQuote KeyAscii
    Select Case FGrid1.Col
        Case C_LabCode
            If DGLabour.Visible = True Then DGridTxtKeyPress txtgrid1, Index, RsLab, KeyAscii, "Code"
        Case C_LabName
            If DGLabour.Visible = True Then DGridTxtKeyPress txtgrid1, Index, RsLab, KeyAscii, "Name"
        Case C_Hrs, C_ActHrs
            NumPress txtgrid1(Index), KeyAscii, 3, 2
        Case C_Rate
'            If Txt(RateSystemYN) = "No" Then Exit Sub
                NumPress txtgrid1(Index), KeyAscii, 4, 2
        Case C_Amt
            If FGrid1.TextMatrix(FGrid1.Row, C_Fixed) = "Yes" Then Exit Sub
            NumPress txtgrid1(Index), KeyAscii, 5, 2
        Case C_GPNo
            If FGrid1.TextMatrix(FGrid1.Row, C_External) <> "Yes" Then Exit Sub
            If DGGatePass.Visible = True Then DGridTxtKeyPress txtgrid1, Index, RsGatePass, KeyAscii, "Code"
        Case C_JobCode
            If DGJobCode.Visible = True Then DGridTxtKeyPress txtgrid1, Index, RsJobCode, KeyAscii, "Code"
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case 0
        Select Case FGrid1.Col
            Case C_LabCode
                If KeyCode <> 13 And DGLabour.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0: DGridTxtKeyPress txtgrid1, Index, RsLab, KeyCode, "Code", True
            Case C_LabName
                If KeyCode <> 13 And DGLabour.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0: DGridTxtKeyPress txtgrid1, Index, RsLab, KeyCode, "Name", True
            Case C_PaidBy
                If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0
                If FrmList.Visible = True Then ListView_KeyUp ListView, txtgrid1, Index, KeyCode, mListItem
            Case C_GPNo
                If FGrid1.TextMatrix(FGrid1.Row, C_External) <> "Yes" Then Exit Sub
                If KeyCode <> 13 And DGGatePass.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0: DGridTxtKeyPress txtgrid1, Index, RsGatePass, KeyCode, "Code", True
            Case C_JobCode
                If KeyCode <> 13 And DGJobCode.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0: DGridTxtKeyPress txtgrid1, Index, RsJobCode, KeyCode, "Code", True
        End Select
    End Select
    If KeyCode = vbKeyEscape Then
        FGrid1.Col = C_LabName
        FGrid1.SetFocus
        txtgrid1(0).Visible = False
        Grid_Hide
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
Dim I As Integer
On Error GoTo ELoop
Cancel = Not TxtGrid1Leave(Index)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid1Leave(Optional Index As Integer) As Boolean
Dim I As Integer, mLabAmt As Single

Dim mDatediff As Integer
Dim MrstDEp As ADODB.Recordset


    Select Case FGrid1.Col
        Case C_LabCode, C_LabName
            If RsLab.EOF = True Or RsLab.BOF = True Or txtgrid1(0).TEXT = "" Then
                FGrid1.TextMatrix(FGrid1.Row, C_LabCode) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_LabName) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Hrs) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Rate) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Amt) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Fixed) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_External) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Major) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_PaidBy) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_TaxYN) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_MechVoice) = ""
            Else
                If UCase(left(PubComp_Name, 3)) <> "LMP" Then
                    For I = 1 To FGrid1.Rows - 1
                        If FGrid1.TextMatrix(I, C_LabCode) = RsLab!Code And I <> FGrid1.Row Then
                            MsgBox "Duplicate Labour Not Allowed", vbInformation, "Validation"
                            GoTo NXT
                        End If
                    Next I
                End If
                FGrid1.TextMatrix(FGrid1.Row, C_LabCode) = RsLab!Code
                FGrid1.TextMatrix(FGrid1.Row, C_LabName) = RsLab!Name
                If StrCmp(left(PubComp_Name, 7), "Vandana") Then
                    FGrid1.TextMatrix(FGrid1.Row, C_MechVoice) = ""
                Else
                    FGrid1.TextMatrix(FGrid1.Row, C_MechVoice) = RsLab!Name
                End If
                FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Chargeable"
                FGrid1.TextMatrix(FGrid1.Row, C_TaxYN) = IIf(ApplyServTax = 1, "Yes", "No")
                'shekharkapil
                FGrid1.TextMatrix(FGrid1.Row, C_Hrs) = IIf(IsNull(RsLab!TIMEREQ) Or RsLab!TIMEREQ = 0, "", Round(RsLab!TIMEREQ, 2))
                FGrid1.TextMatrix(FGrid1.Row, C_ActHrs) = IIf(IsNull(RsLab!TIMEREQ) Or RsLab!TIMEREQ = 0, "", Round(RsLab!TIMEREQ, 2))
                FGrid1.TextMatrix(FGrid1.Row, C_Rate) = IIf(IsNull(RsLab!LabRate) Or RsLab!LabRate = 0, "", RsLab!LabRate)
                
                
                
                  'Start****************************Nikhil
    'Deptcode
    FGrid1.TextMatrix(FGrid1.Row, Col_DepItem) = XNull(RsLab!Deptcode)
    FGrid1.TextMatrix(FGrid1.Row, Col_DepitemPer) = XNull(RsLab!Dep_per)
    
    If txt(SoldDate) <> "" Then
        mDatediff = DateDiff("M", CDate(txt(SoldDate)), CDate(txt(JobDt)))
        FGrid1.TextMatrix(FGrid1.Row, Col_DiffPeried) = VNull(mDatediff)
    End If
        
    Set MrstDEp = GCn.Execute("SELECT  TOP 1 * FROM Deprecation_Master WHERE Dep_Month>" & mDatediff & " ORDER BY Dep_Month ")
    If MrstDEp.RecordCount > 0 Then
    
    FGrid1.TextMatrix(FGrid1.Row, Col_DepCode) = XNull(MrstDEp!Code)
    FGrid1.TextMatrix(FGrid1.Row, Col_DepPer) = XNull(MrstDEp!Dep_per)
    Else
    FGrid1.TextMatrix(FGrid1.Row, Col_DepCode) = ""
    FGrid1.TextMatrix(FGrid1.Row, Col_DepPer) = ""
    
    End If
'****************End



                

                If UCase(left(PubComp_Name, 3)) = "LMP" Then
                    FGrid1.TextMatrix(FGrid1.Row, C_Amt) = Round(Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)) * Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)), 2)
                Else
                    FGrid1.TextMatrix(FGrid1.Row, C_Amt) = IIf(IsNull(RsLab!LabRate) Or RsLab!LabRate = 0, "", Format(RsLab!LabRate, "0.00"))
                End If
                FGrid1.TextMatrix(FGrid1.Row, C_Fixed) = IIf(IsNull(RsLab!Fixed), "No", "Yes")
                FGrid1.TextMatrix(FGrid1.Row, C_External) = IIf(RsLab!External_yn = "1", "Yes", "No")
                FGrid1.TextMatrix(FGrid1.Row, C_Major) = RsLab!Major_YN
'                FGrid1.TextMatrix(FGrid1.Row, C_PaidBy) = RsLab!CHRG_FROM
                If RsLab!Chrg_From = "C" Then
                    FGrid1.TextMatrix(FGrid1.Row, C_PaidBy) = "Customer"  'RsLab!CHRG_FROM
                    FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Chargeable"
                ElseIf RsLab!Chrg_From = "M" Then
                    FGrid1.TextMatrix(FGrid1.Row, C_PaidBy) = "Manufacturer"  'RsLab!CHRG_FROM
                    FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Free Service"
                ElseIf RsLab!Chrg_From = "S" Then
                    FGrid1.TextMatrix(FGrid1.Row, C_PaidBy) = "Self"  'RsLab!CHRG_FROM
                    FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Free Service"
                ElseIf RsLab!Chrg_From = "O" Then
                    FGrid1.TextMatrix(FGrid1.Row, C_PaidBy) = "Other Dealer"  'RsLab!CHRG_FROM
                    FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Free Service"
                End If
            End If
            If FGrid1.TextMatrix(FGrid1.Row, C_External) <> "Yes" Then
                FGrid1.TextMatrix(FGrid1.Row, C_GPNo) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_ContCode) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_ContName) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_WIssueDt) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_WRecdDt) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_ContAmt) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Remarks) = ""
            End If
            If FGrid1.TextMatrix(FGrid1.Rows - 1, C_LabCode) <> "" Then FGrid1.AddItem FGrid1.Rows
            FillMechGrid FGrid1.Row
        Case C_PaidBy
            txtgrid1(0).TEXT = ListView.SelectedItem.TEXT
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = txtgrid1(0).TEXT
            If FGrid1.TextMatrix(FGrid1.Row, C_PaidBy) = "Customer" Then
                FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Chargeable"
            ElseIf FGrid1.TextMatrix(FGrid1.Row, C_PaidBy) = "Manufacturer" Then
                FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Free Service"
            ElseIf FGrid1.TextMatrix(FGrid1.Row, C_PaidBy) = "Self" Then
                FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Free Service"
            ElseIf FGrid1.TextMatrix(FGrid1.Row, C_PaidBy) = "Other Dealer" Then
                FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Free Service"
            End If
        Case C_ActHrs
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = IIf(Val(txtgrid1(0)) <> 0, Format(txtgrid1(0), "0.00"), "")
            
            If Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)) > 0 Then
                If UCase(left(PubComp_Name, 3)) <> "LMP" Then
                    mLabAmt = Round(Int(Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs))) * Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)), 2)
                    If (Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)) - Int(Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)))) > 0 Then
                        mLabAmt = mLabAmt + Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)) / (0.6 / (Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)) - (Int(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)))))
                    Else
                        mLabAmt = Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)) * (Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)))
                    End If
                Else
                    mLabAmt = Round(Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)) * Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)), 2)
                End If
                If mLabAmt <= 0 Then
                    FGrid1.TextMatrix(FGrid1.Row, C_Amt) = ""
                Else
                    FGrid1.TextMatrix(FGrid1.Row, C_Amt) = Format(mLabAmt, "0.00")
                End If
            End If
            TotHrAmt
            
        Case C_Hrs
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = IIf(Val(txtgrid1(0)) <> 0, Format(txtgrid1(0), "0.00"), "")
'            If Txt(RateSystemYN) = "Yes" Then
            
            If Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)) > 0 Then
                If UCase(left(PubComp_Name, 3)) <> "LMP" Then
                    mLabAmt = Round(Int(Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs))) * Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)), 2)
                    If (Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)) - Int(Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)))) > 0 Then
                        mLabAmt = mLabAmt + Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)) / (0.6 / (Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)) - (Int(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)))))
                    Else
                        mLabAmt = Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)) * (Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)))
                    End If
                Else
                    mLabAmt = Round(Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)) * Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)), 2)
                End If
                If mLabAmt <= 0 Then
                    FGrid1.TextMatrix(FGrid1.Row, C_Amt) = ""
                Else
                    FGrid1.TextMatrix(FGrid1.Row, C_Amt) = Format(mLabAmt, "0.00")
                End If
            End If
            TotHrAmt
        Case C_Rate 'for Kota 23-07-2003
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = IIf(Val(txtgrid1(0)) <> 0, Format(txtgrid1(0), "0.00"), "")
            If Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)) = 0 Then FGrid1.TextMatrix(FGrid1.Row, C_Hrs) = 1
            If UCase(left(PubComp_Name, 3)) <> "LMP" Then
                mLabAmt = Round(Val(Int(FGrid1.TextMatrix(FGrid1.Row, C_Hrs))) * Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)), 2)
                If (Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)) - (Int(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)))) > 0 Then
                    mLabAmt = mLabAmt + Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)) / (0.6 / (Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)) - (Int(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)))))
                Else
                    mLabAmt = Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)) * (Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)))
                End If
            Else
                mLabAmt = Round(Val(FGrid1.TextMatrix(FGrid1.Row, C_Hrs)) * Val(FGrid1.TextMatrix(FGrid1.Row, C_Rate)), 2)
            End If
            If mLabAmt <= 0 Then
                FGrid1.TextMatrix(FGrid1.Row, C_Amt) = ""
            Else
                FGrid1.TextMatrix(FGrid1.Row, C_Amt) = Format(mLabAmt, "0.00")
            End If
            TotHrAmt
            FillMechGrid FGrid1.Row
        Case C_Amt ', C_WarAmt
            If FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) = "Warranty" Or _
                (FGrid1.TextMatrix(FGrid1.Row, C_ChrgType) <> "Warranty" And FGrid1.TextMatrix(FGrid1.Row, C_Fixed) <> "Yes") Then
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = IIf(Val(txtgrid1(0)) <> 0, Format(txtgrid1(0), "0.00"), "")
            End If
            TotHrAmt
            FillMechGrid FGrid1.Row
        Case C_GPNo
            If FGrid1.TextMatrix(FGrid1.Row, C_External) <> "Yes" Then Exit Function
            If RsGatePass.BOF = True Or RsGatePass.EOF = True Or txtgrid1(0).TEXT = "" Then
                FGrid1.TextMatrix(FGrid1.Row, C_GPNo) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_ContCode) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_ContName) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_WIssueDt) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_WRecdDt) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_ContAmt) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Remarks) = ""
            Else
                If IsNull(RsGatePass!ContractRecdDate) Then
                    MsgBox "External Job not Recd.", vbCritical, "External Job"
                    Exit Function
                End If
                FGrid1.TextMatrix(FGrid1.Row, C_GPNo) = RsGatePass!Code
                FGrid1.TextMatrix(FGrid1.Row, C_ContCode) = RsGatePass!ContractCode
                FGrid1.TextMatrix(FGrid1.Row, C_ContName) = RsGatePass!FinName
                FGrid1.TextMatrix(FGrid1.Row, C_WIssueDt) = RsGatePass!GatePassDate
                FGrid1.TextMatrix(FGrid1.Row, C_WRecdDt) = IIf(IsNull(RsGatePass!ContractRecdDate), "", RsGatePass!ContractRecdDate)
                FGrid1.TextMatrix(FGrid1.Row, C_ContAmt) = IIf(IsNull(RsGatePass!ContractAmt) Or RsGatePass!ContractAmt = 0, "", Format(RsGatePass!ContractAmt, "0.00"))
                If Val(FGrid1.TextMatrix(FGrid1.Row, C_ContAmt)) > Val(FGrid1.TextMatrix(FGrid1.Row, C_Amt)) Then
                    MsgBox "Amount is less than External Job Amount", vbCritical, "External Job"
                    FGrid1.SetFocus 'Col = C_Amt
                End If
            End If
        Case C_Remarks
            If FGrid1.TextMatrix(FGrid1.Row, C_External) <> "Yes" Then Exit Function
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = txtgrid1(0).TEXT
        Case C_JobCode, C_MechVoice
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = txtgrid1(0).TEXT
    End Select
    Call Amt_Cal
NXT:
    TxtGrid1Leave = True
End Function
Private Sub Amt_Cal()
 'Nikhil
    Dim LblNetValue As Double
    
    LblNetValue = Format(Val(FGrid1.TextMatrix(FGrid1.Row, C_Amt)), "0.00")
      If Val(FGrid1.TextMatrix(FGrid1.Row, Col_DepitemPer)) > 0 And Val(FGrid1.TextMatrix(FGrid1.Row, Col_DepPer)) > 0 Then
     
        
        FGrid1.TextMatrix(FGrid1.Row, Col_DepAmt) = Format((Val(LblNetValue) * Val(FGrid1.TextMatrix(FGrid1.Row, Col_DepitemPer)) / 100) * Val(FGrid1.TextMatrix(FGrid1.Row, Col_DepPer)) / 100, "0.00")
           FGrid1.TextMatrix(FGrid1.Row, Col_InsuranceAmt) = Format(Val(LblNetValue) - Val(FGrid1.TextMatrix(FGrid1.Row, Col_DepAmt)), "0.00")
        Else
        FGrid1.TextMatrix(FGrid1.Row, Col_DepAmt) = ""
        FGrid1.TextMatrix(FGrid1.Row, Col_InsuranceAmt) = ""
      End If
      
      
End Sub
Private Sub Fill_Grid(ByVal DocID As String, ByVal Model As String)
Dim mVehicleType$
Dim MyRst As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim mDatediff As Integer
Dim MrstDEp As ADODB.Recordset


Dim I As Integer
    'by lps 08-12-2K2 Model Specific Labour Considered
    Set RsTemp = GCn.Execute("Select Vehicle_Type from Model where Model='" & Model & "'")
    If RsTemp.RecordCount > 0 Then
        mVehicleType = XNull(RsTemp!Vehicle_Type)
    End If
    Set RsTemp = GCn.Execute("Select ServiceTax_YN from Model where Model='" & Model & "'")
    If RsTemp.RecordCount > 0 Then
        ApplyServTax = VNull(RsTemp!ServiceTax_YN)
    End If
    
    

'        GSQL = "select L.Lab_Code as Code, L.Lab_Desc as Name,LG.LabGrp_Desc, L.External_YN, L.Major_YN, L.Chrg_From," & _
'            " switch(LM.Lab_Rate<>0, LM.Lab_Rate,isnull(LM.Lab_Rate) or LM.Lab_Rate=0,L.Lab_Rate) as LabRate," & _
'            " switch(LM.Time_Req<>0, LM.Time_Req,isnull(LM.Time_Req) or LM.Time_Req=0,L.Time_Req) as TimeReq," & _
'            " switch(LM.WTime_Req<>0, LM.WTime_Req,isnull(LM.WTime_Req) or LM.WTime_Req=0,L.WTime_Req) as WTimeReq," & _
'            " LM.Fixed" & _
'            " from ((Labour as L left Join Labour_Model as LM on L.Lab_Code=LM.Lab_Code) " & _
'            " left join Labour_Group as LG on L.Lab_Group=LG.Lab_Group) " & _
'            " WHERE l.ModelBased = 0 Or " & _
'            " (l.ModelBased = 1 And l.Lab_Code = LM.Lab_Code and LM.Vehicle_Type='" & mVehicleType & "')"

        GSQL = "select L.Lab_Code as Code, L.Lab_Desc as Name,LG.LabGrp_Desc, L.External_YN, L.Major_YN, L.Chrg_From," & _
            " " & cIIF(vIsNull("LM.Lab_Rate", "0") & "<> 0", "LM.Lab_Rate", "L.Lab_Rate") & " as LabRate," & _
            " " & cIIF(vIsNull("LM.Time_Req", "0") & "<>0", "LM.Time_Req", "L.Time_Req") & " as TimeReq," & _
            " " & cIIF(vIsNull("LM.WTime_Req", "0") & "<>0", "LM.WTime_Req", "L.WTime_Req") & " as WTimeReq," & _
            " LM.Fixed ,l.dep_item as Deptcode,ditm.dep_per as dep_per" & _
            " from ((Labour as L left Join Labour_Model as LM on L.Lab_Code=LM.Lab_Code) " & _
            " left join Labour_Group as LG on L.Lab_Group=LG.Lab_Group)left join Deprecation_itemMaster DITM on l.dep_item=ditm.code  " & _
            " WHERE l.ModelBased = 0 Or " & _
            " (l.ModelBased = 1 And l.Lab_Code = LM.Lab_Code and LM.Vehicle_Type='" & mVehicleType & "')"

 
    
    Set RsLab = New ADODB.Recordset
    RsLab.CursorLocation = adUseClient
    RsLab.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGLabour.DataSource = RsLab
    RsLab.Sort = "code"
    RsLab.Sort = "name"
    
    FGrid1.Redraw = False
    FGrid1.Rows = 1
    If GCn.Execute("Select Lab_code from Job_Lab").RecordCount <= 0 Then GoTo ExitLoop
    Set MyRst = New ADODB.Recordset
    MyRst.CursorLocation = adUseClient
    GSQL = "Select JL.*, L.Lab_Desc AS LabName,CF.FinName AS ContName,LM.Fixed,GP.GatePassDate,GP.ContractRecdDate,GP.ContractAmt,GP.ContractCode " & _
        " From ((((Job_Lab as JL left join labour as L on JL.Lab_Code=L.Lab_Code) " & _
        " Left Join Labour_Model LM on JL.Lab_Code=LM.Lab_Code) " & _
        " left join Job_GatePass as GP on JL.ExtJobGatePassNo=GP.GatePassNo) " & _
        " Left Join ContractFinance as CF ON GP.ContractCode=CF.FinCode) " & _
        " Where JL.Job_DocId='" & DocID & _
        "' and (LM.Vehicle_Type='" & mVehicleType & "' or l.ModelBased = 0 Or L.ModelBased Is Null) Order by JL.S_No"
    MyRst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    I = 1
    If MyRst.RecordCount > 0 Then
        Do Until MyRst.EOF
            FGrid1.AddItem ""
            With FGrid1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, C_LabCode) = MyRst!Lab_Code
                .TextMatrix(I, C_LabName) = XNull(MyRst!LabName)
                .TextMatrix(I, C_MechVoice) = XNull(MyRst!Mech_Voice)
                .TextMatrix(I, C_ActHrs) = Format(VNull(MyRst!ActualHrs), "0.00")
'                If MyRst!Hrs_Taken + MyRst!Lab_Rate  > 0 Then
                If MyRst!Chrg_From = "M" Or MyRst!Chrg_From = "O" Then
                    If MyRst!Chrg_Type = "W" Then 'Warranty
                        .TextMatrix(I, C_ChrgType) = "Warranty"
                        .TextMatrix(I, C_Hrs) = IIf(MyRst!Hrs_War = 0, "", Format(MyRst!Hrs_War, "0.00"))
                        .TextMatrix(I, C_Rate) = IIf(MyRst!War_Lab_Rate = 0, "", Format(MyRst!War_Lab_Rate, "0.00"))
                    ElseIf MyRst!Chrg_Type = "P" Then 'PDI
                        .TextMatrix(I, C_ChrgType) = "PDI"
                        .TextMatrix(I, C_Hrs) = IIf(MyRst!Hrs_Taken = 0, "", Format(MyRst!Hrs_Taken, "0.00"))
                        .TextMatrix(I, C_Rate) = IIf(MyRst!Lab_Rate = 0, "", Format(MyRst!Lab_Rate, "0.00"))
                    ElseIf MyRst!Chrg_Type = "C" Then
                        .TextMatrix(I, C_ChrgType) = "Chargable"
                        .TextMatrix(I, C_Hrs) = IIf(MyRst!Hrs_Taken = 0, "", Format(MyRst!Hrs_Taken, "0.00"))
                        .TextMatrix(I, C_Rate) = IIf(MyRst!Lab_Rate = 0, "", Format(MyRst!Lab_Rate, "0.00"))
                    ElseIf MyRst!Chrg_Type = "A" Then
                        .TextMatrix(I, C_ChrgType) = "AMC"
                        .TextMatrix(I, C_Hrs) = IIf(MyRst!Hrs_Taken = 0, "", Format(MyRst!Hrs_Taken, "0.00"))
                        .TextMatrix(I, C_Rate) = IIf(MyRst!Lab_Rate = 0, "", Format(MyRst!Lab_Rate, "0.00"))
                    Else    'Free Service
                        .TextMatrix(I, C_ChrgType) = "Free Service"
                        .TextMatrix(I, C_Hrs) = IIf(MyRst!Hrs_Taken = 0, "", Format(MyRst!Hrs_Taken, "0.00"))
                        .TextMatrix(I, C_Rate) = IIf(MyRst!Lab_Rate = 0, "", Format(MyRst!Lab_Rate, "0.00"))
                    End If
                Else
                    .TextMatrix(I, C_ChrgType) = IIf(MyRst!Chrg_Type = "A", "AMC", "Chargable")
                    .TextMatrix(I, C_Hrs) = IIf(MyRst!Hrs_Taken = 0, "", Format(MyRst!Hrs_Taken, "0.00"))
                    .TextMatrix(I, C_Rate) = IIf(MyRst!Lab_Rate = 0, "", Format(MyRst!Lab_Rate, "0.00"))
                End If
                .TextMatrix(I, C_Amt) = IIf(MyRst!LabourAmt = 0, "", Format(MyRst!LabourAmt, "0.00"))
                .TextMatrix(I, C_GPNo) = XNull(MyRst!ExtJobGatePassNo)
                .TextMatrix(I, C_ContName) = XNull(MyRst!ContName)
                .TextMatrix(I, C_WIssueDt) = IIf(IsNull(MyRst!GatePassDate), "", MyRst!GatePassDate)
                .TextMatrix(I, C_WRecdDt) = IIf(IsNull(MyRst!ContractRecdDate), "", MyRst!ContractRecdDate)
                .TextMatrix(I, C_ContAmt) = IIf(MyRst!ContractAmt = 0, "", Format(MyRst!ContractAmt, "0.00"))
                .TextMatrix(I, C_Remarks) = XNull(MyRst!Contract_Remarks)
                .TextMatrix(I, C_ContCode) = XNull(MyRst!ContractCode)
                .TextMatrix(I, C_External) = IIf(MyRst!External_yn = "1", "Yes", "No")
                .TextMatrix(I, C_Major) = MyRst!Major_YN
                .TextMatrix(I, C_Fixed) = IIf(IsNull(MyRst!Fixed) Or MyRst!Fixed = 0, "No", "Yes")
                .TextMatrix(I, C_TaxYN) = IIf(MyRst!Tax_YN = "1", "Yes", "No")
                
                
                'Start****************************Nikhil
    'Deptcode
''''''''  .TextMatrix(I, Col_DepItem) = XNull(MyRst!Deptcode)
''''''''    .TextMatrix(I, Col_DepitemPer) = XNull(MyRst!Dep_per)
''''''''    mDatediff = DateDiff("M", CDate(txt(SoldDate)), CDate(txt(JobDt)))
''''''''    .TextMatrix(I, Col_DiffPeried) = VNull(mDatediff)
''''''''
''''''''    Set MrstDEp = GCn.Execute("SELECT  TOP 1 * FROM Deprecation_Master WHERE Dep_Month>" & mDatediff & " ORDER BY Dep_Month ")
''''''''    If MrstDEp.RecordCount > 0 Then
''''''''
''''''''    .TextMatrix(I, Col_DepCode) = XNull(MrstDEp!Code)
''''''''    .TextMatrix(I, Col_DepPer) = XNull(MrstDEp!Dep_per)
''''''''    Else
''''''''    .TextMatrix(I, Col_DepCode) = ""
''''''''    .TextMatrix(I, Col_DepPer) = ""
''''''''
''''''''    End If


                     .TextMatrix(I, Col_DepItem) = IIf(IsNull(MyRst!Dep_Item), "", MyRst!Dep_Item)
                     .TextMatrix(I, Col_DepCode) = IIf(IsNull(MyRst!Dep_Code), "", MyRst!Dep_Code)
                     .TextMatrix(I, Col_DepitemPer) = Format(MyRst!DepitemPer, "0.00")
                     .TextMatrix(I, Col_DepPer) = Format(MyRst!DepPer, "0.00")
                     .TextMatrix(I, Col_DepAmt) = Format(MyRst!DepAmt, "0.00")
                     
                     .TextMatrix(I, Col_InsuranceAmt) = Format(MyRst!InsuranceAmt, "0.00")
                     .TextMatrix(I, Col_DiffPeried) = Format(MyRst!DiffPeried, "0.00")
                     
                     
'****************End


                '"Customer", "Manufacturer", "Self", "Other Dealer"
                If MyRst!Chrg_From = "C" Then
                    .TextMatrix(I, C_PaidBy) = "Customer"  'MyRst!CHRG_FROM
                ElseIf MyRst!Chrg_From = "M" Then
                    .TextMatrix(I, C_PaidBy) = "Manufacturer"  'MyRst!CHRG_FROM
                ElseIf MyRst!Chrg_From = "S" Then
                    .TextMatrix(I, C_PaidBy) = "Self"  'MyRst!CHRG_FROM
                ElseIf MyRst!Chrg_From = "O" Then
                    .TextMatrix(I, C_PaidBy) = "Other Dealer"  'MyRst!CHRG_FROM
                End If
                .TextMatrix(I, C_JobCode) = XNull(MyRst!JobCode)
            End With
            I = I + 1
            MyRst.MoveNext
        Loop
        FGrid1.AddItem ""
        FGrid1.FixedRows = 1
    Else
ExitLoop:
        FGrid1.Rows = FGrid1.Rows
        FGrid1.AddItem 1
        FGrid1.FixedRows = 1
    End If
    FGrid1.Redraw = True
    FGrid1.Row = 1
    Set MyRst = Nothing
    TotHrAmt
    '*************
    'by lps 29-03-2K3 Multiple Mechanic in Single Labour Considered
    FGridLab2.Rows = 1
    If Master.RecordCount > 0 Then
        Set MyRst = New ADODB.Recordset
        MyRst.CursorLocation = adUseClient
        GSQL = "Select JL2.*, Emp.Emp_Name AS MechName " & _
            " From (Job_LAB2 as JL2 left join Emp_Mast Emp ON JL2.Mech_Code=Emp.Emp_Code) where Div_Code='" & PubDivCode & "' " & _
            " and JL2.Job_DocId='" & DocID & _
            "' Order by JL2.S_No"
        MyRst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
        I = 1
        If MyRst.RecordCount > 0 Then
            Do Until MyRst.EOF
                FGridLab2.AddItem ""
                With FGridLab2
                    .RowData(I) = MyRst!S_No
                    .TextMatrix(I, 0) = I
                    .TextMatrix(I, LabCode) = MyRst!Lab_Code
                    .TextMatrix(I, MechCode) = XNull(MyRst!mech_code)
                    .TextMatrix(I, MechName) = XNull(MyRst!MechName)
                End With
                I = I + 1
                MyRst.MoveNext
            Loop
        End If
        Set MyRst = Nothing
    End If
    If FGridLab2.Rows = 1 Then FGridLab2.AddItem ""
    FGridLab2.FixedRows = 1
    FGridLab2.Redraw = True
    FGrid1_RowColChange
    '*************
End Sub

Private Sub SetMaxLength()
    Select Case FGrid1.Col
        Case C_LabCode
            txtgrid1(0).MaxLength = 6
            txtgrid1(0).Alignment = 0
        Case C_LabName ', C_ContName ', C_MechName
            txtgrid1(0).MaxLength = 40
            txtgrid1(0).Alignment = 0
            
        Case C_Hrs, C_ActHrs, C_Amt, C_ContAmt ', C_WarHrs, C_WarAmt
            txtgrid1(0).MaxLength = 0
            txtgrid1(0).Alignment = 1
        Case C_MechVoice
            txtgrid1(0).MaxLength = 40
        Case C_WIssueDt, C_WRecdDt
            txtgrid1(0).MaxLength = 0
            txtgrid1(0).Alignment = 0
            
        Case C_Remarks
            txtgrid1(0).MaxLength = 20
            txtgrid1(0).Alignment = 0
        Case Else
            txtgrid1(0).MaxLength = 0
            txtgrid1(0).Alignment = 0
    End Select
End Sub


Public Sub SelGridKeyPressLocal(txt As Object, FGrid As Object, Index As Integer, Rst As ADODB.Recordset, ByRef KeyAscii As Integer, FindFldName As String, Optional CellBackColEnter As ColorConstants, Optional CellBackColLeave As ColorConstants)
Dim FindStr$    ' As String
Dim LPlace As Byte
'    If FilterKeyCode(KeyAscii) = True Then Exit Sub
    If FGrid(Index).Rows < 1 Then Exit Sub
    If Rst.RecordCount <= 0 Then txt.TEXT = "": Exit Sub
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyDelete Then Exit Sub
        If KeyAscii = vbKeyBack And Len(txt.SelText) <> 1 Then
            txt.SelLength = Len(txt.SelText) - 1
            FindStr = txt.SelText
        Else
            FindStr = txt.SelText + Chr(KeyAscii)
        End If
        Rst.MoveFirst
        If Rst.Fields(FindFldName).Type = adInteger Then    'Numeric Search
            Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
        Else    'character serach
            Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
        End If
        KeyAscii = 0
       If Rst.AbsolutePosition <> adPosEOF And Rst.AbsolutePosition <> adPosBOF Then
            FGrid(Index).CellBackColor = CellBackColLeave
            FGrid(Index).Row = Rst.AbsolutePosition
            FGrid(Index).CellBackColor = CellBackColEnter
            txt.TEXT = Rst.Fields(FindFldName).Value
            txt.SelLength = Len(FindStr)
            txt.left = FGrid(Index).CellLeft + FGrid(Index).left
            txt.top = FGrid(Index).CellTop + FGrid(Index).top
            If txt.Visible = False Then
                txt.Visible = True: txt.ZOrder 0: txt.SetFocus: txt.BackColor = FGrid(Index).CellBackColor
                 txt.ForeColor = FGrid(Index).CellForeColor: txt.width = FGrid(Index).CellWidth: txt.height = FGrid(Index).CellHeight
            End If
       End If
End Sub

Private Sub TxtSearch_Click()
TxtSearch.Visible = False: TxtSearch.TEXT = "": GridSel(Val(TxtSearch.Tag)).SetFocus
End Sub

Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If NavigationKey(KeyCode) = True Then TxtSearch.Visible = False: GridSel(Val(TxtSearch.Tag)).SetFocus
If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then TxtSearch.Visible = False: GridSel(Val(TxtSearch.Tag)).SetFocus
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
Select Case TxtSearch.Tag
    Case 1
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsMech, KeyAscii, "Name", vbCyan
End Select
End Sub

Private Sub TxtSearch_LostFocus()
TxtSearch.Visible = False: TxtSearch.TEXT = ""
End Sub
Private Sub FillMechGrid(FGrid1Row As Integer)
Dim I As Integer, j As Integer
    '**********fILL mECHANIC gRID
    GSQL = "Select Emp_Code as code,Emp_Name as Name FROM Emp_Mast where Div_Code='" & PubDivCode & "' And  Emp_type=1 and Designation ='MECHANIC' and (LeftOn Is Null or LeftOn >=" & ConvertDate(PubLoginDate) & ") Order by Emp_name"
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    GridSel(1).Redraw = False
    GridSel(1).Rows = 1
    If GRs.RecordCount > 0 Then
        I = 1
        Do Until GRs.EOF
            GridSel(1).AddItem ""
            With GridSel(1)
                .TextMatrix(I, 0) = ""
                .TextMatrix(I, LabCode) = FGrid1.TextMatrix(FGrid1Row, C_LabCode)
                .TextMatrix(I, MechCode) = GRs!Code
                .TextMatrix(I, MechName) = GRs!Name
            End With
            I = I + 1
            GRs.MoveNext
        Loop
        I = 1
        For I = 1 To FGridLab2.Rows - 1
            If FGridLab2.TextMatrix(I, LabCode) = FGrid1.TextMatrix(FGrid1Row, C_LabCode) Then
                For j = 1 To GridSel(1).Rows - 1
                    If FGridLab2.TextMatrix(I, MechCode) = GridSel(1).TextMatrix(j, MechCode) Then
                        GridSel(1).Row = j
                        GridSel(1).Col = 0
                        GridSel(1).CellFontName = "WINGDINGS"
                        GridSel(1).CellFontSize = 14
                        GridSel(1).TextMatrix(j, 0) = "ü"
                    End If
                Next
            End If
        Next
    Else
        GridSel(1).AddItem 1
    End If
    Set GRs = Nothing
    GridSel(1).Visible = True
    GridSel(1).FixedRows = 1
    GridSel(1).Redraw = True
    GridSel(1).Col = 0
    GridSel(1).Row = 1
'    GridSel(1).SetFocus
End Sub
Private Sub TotHrAmt()
Dim I As Integer, TotContAmtChg As Double, TotContAmtPaid As Double
Dim TotWarHrs As Single, TotWarAmt As Double
Dim TotChgHrs As Single, TotChgAmt As Double
Dim TotMfgHrs As Single, TotMfgAmt As Double
Dim TotOthHrs As Single, TotOthAmt As Double

For I = 1 To FGrid1.Rows - 1
    Select Case left(FGrid1.TextMatrix(I, C_PaidBy), 1)
        Case "M"    '"Manufacturer"
            If FGrid1.TextMatrix(I, C_ChrgType) = "Warranty" Then
                TotWarHrs = TotWarHrs + Val(FGrid1.TextMatrix(I, C_Hrs))
                TotWarAmt = TotWarAmt + Val(FGrid1.TextMatrix(I, C_Amt))
            Else    'Free Service
                TotMfgHrs = TotMfgHrs + Val(FGrid1.TextMatrix(I, C_Hrs))
                TotMfgAmt = TotMfgAmt + Val(FGrid1.TextMatrix(I, C_Amt))
            End If
        Case "S", "O"   '"Self", "Other Dealer
            TotOthHrs = TotOthHrs + Val(FGrid1.TextMatrix(I, C_Hrs))
            TotOthAmt = TotOthAmt + Val(FGrid1.TextMatrix(I, C_Amt))
        Case Else
            'Case "C"    'Customer"
            TotChgHrs = TotChgHrs + Val(FGrid1.TextMatrix(I, C_Hrs))
            TotChgAmt = TotChgAmt + Val(FGrid1.TextMatrix(I, C_Amt))
    End Select
    If FGrid1.TextMatrix(I, C_External) = "Yes" Then
        TotContAmtChg = TotContAmtChg + Val(FGrid1.TextMatrix(I, C_Amt))
        TotContAmtPaid = TotContAmtPaid + Val(FGrid1.TextMatrix(I, C_ContAmt))
    End If
Next
txt(WarrHrs) = IIf(TotWarHrs <> 0, Format(TotWarHrs, "0.00"), "")
txt(WarrAmt) = IIf(TotWarAmt <> 0, Format(TotWarAmt, "0.00"), "")
txt(ChgHrs) = IIf(TotChgHrs <> 0, Format(TotChgHrs, "0.00"), "")
txt(ChgAmt) = IIf(TotChgAmt <> 0, Format(TotChgAmt, "0.00"), "")
txt(MfgHrs) = IIf(TotMfgHrs <> 0, Format(TotMfgHrs, "0.00"), "")
txt(MfgAmt) = IIf(TotMfgAmt <> 0, Format(TotMfgAmt, "0.00"), "")
txt(OthHrs) = IIf(TotOthHrs <> 0, Format(TotOthHrs, "0.00"), "")
txt(OthAmt) = IIf(TotOthAmt <> 0, Format(TotOthAmt, "0.00"), "")
txt(ExtLab) = IIf(TotContAmtPaid <> 0, Format(TotContAmtPaid, "0.00"), "")
txt(ExtLabChg) = IIf(TotContAmtChg <> 0, Format(TotContAmtChg, "0.00"), "")
End Sub

Private Function LabTimeChk(Index As Integer)
    Dim Hr As Double, Min As Double, time As Double
    time = Val(txtgrid1(0).TEXT)
    Hr = Int(time)
    time = time - Hr
    If time > 0.6 Then
        LabTimeChk = False
    Else
        LabTimeChk = True
    End If
End Function
