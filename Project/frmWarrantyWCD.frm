VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmWarrantyWCD 
   Appearance      =   0  'Flat
   BackColor       =   &H00E7DFDC&
   Caption         =   "Warranty PCR"
   ClientHeight    =   7230
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
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DGFailure 
      Height          =   3150
      Left            =   3810
      Negotiate       =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4350
      Visible         =   0   'False
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   5556
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5999.812
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGJobCode 
      Height          =   3180
      Left            =   2955
      Negotiate       =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   4725
      Visible         =   0   'False
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   5609
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5999.812
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGMake 
      Height          =   3195
      Left            =   2895
      Negotiate       =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4860
      Visible         =   0   'False
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   5636
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5999.812
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGComplaint 
      Height          =   3225
      Left            =   3135
      Negotiate       =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   5145
      Visible         =   0   'False
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   5689
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4995.213
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   3405
      TabIndex        =   59
      Top             =   3480
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   60
         TabIndex        =   60
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
   Begin VB.TextBox TxtNarr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   1095
      Left            =   7095
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   2505
      Width           =   4530
   End
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   2670
      Left            =   495
      Negotiate       =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   7020
      Visible         =   0   'False
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   4710
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Part No."
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
         Caption         =   "Part Name"
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
         DataField       =   "Unit"
         Caption         =   "Unit"
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
         DataField       =   "MRP"
         Caption         =   "      MRP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "TB_SRate"
         Caption         =   "   TB Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "TP_SRate"
         Caption         =   "   TP Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "LName"
         Caption         =   "Part Local Name"
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
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2564.788
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   24
      Left            =   5430
      MaxLength       =   1
      TabIndex        =   25
      Top             =   3030
      Width           =   270
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   23
      Left            =   3450
      MaxLength       =   1
      TabIndex        =   24
      Top             =   3030
      Width           =   270
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   22
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   23
      Top             =   3030
      Width           =   270
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   21
      Left            =   5430
      MaxLength       =   1
      TabIndex        =   22
      Top             =   2760
      Width           =   270
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   20
      Left            =   3045
      MaxLength       =   5
      TabIndex        =   21
      Text            =   "12345"
      Top             =   2760
      Width           =   675
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   19
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   20
      Top             =   2760
      Width           =   270
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   18
      Left            =   9570
      MaxLength       =   20
      TabIndex        =   19
      Top             =   1335
      Width           =   1665
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   17
      Left            =   9570
      MaxLength       =   20
      TabIndex        =   18
      Top             =   1065
      Width           =   1665
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   16
      Left            =   9570
      MaxLength       =   20
      TabIndex        =   17
      Top             =   795
      Width           =   1665
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   15
      Left            =   9000
      MaxLength       =   20
      TabIndex        =   16
      Top             =   1950
      Width           =   1140
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
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   14
      Left            =   5430
      MaxLength       =   6
      TabIndex        =   15
      Top             =   2490
      Width           =   1335
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   13
      Left            =   5430
      MaxLength       =   10
      TabIndex        =   14
      Top             =   2220
      Width           =   1665
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   12
      Left            =   5430
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1950
      Width           =   1665
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   11
      Left            =   5430
      MaxLength       =   50
      TabIndex        =   12
      Top             =   1680
      Width           =   5805
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   10
      Left            =   9000
      MaxLength       =   10
      TabIndex        =   11
      Top             =   2220
      Width           =   1140
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   9
      Left            =   1875
      MaxLength       =   6
      TabIndex        =   10
      Top             =   2490
      Width           =   1335
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   8
      Left            =   1875
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2220
      Width           =   450
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   7
      Left            =   1875
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1950
      Width           =   270
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   6
      Left            =   1875
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   5
      Left            =   1875
      MaxLength       =   17
      TabIndex        =   6
      Top             =   1410
      Width           =   2625
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   4
      Left            =   3630
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1065
      Width           =   870
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   3
      Left            =   1290
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1065
      Width           =   870
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   2
      Left            =   6690
      MaxLength       =   2
      TabIndex        =   3
      Top             =   795
      Width           =   450
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   1
      Left            =   4410
      MaxLength       =   20
      TabIndex        =   2
      Text            =   "02/06/2003"
      Top             =   795
      Width           =   1140
   End
   Begin MSDataGridLib.DataGrid DGJob 
      Height          =   3075
      Left            =   450
      Negotiate       =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6855
      Visible         =   0   'False
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   5424
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "JobDocId"
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
         Caption         =   "Job No"
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
         DataField       =   "Job_Date"
         Caption         =   "Date"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
         DataField       =   "OwnerName"
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column03 
            DividerStyle    =   3
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3179.906
         EndProperty
      EndProperty
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
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   4530
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   0
      Left            =   1290
      MaxLength       =   15
      TabIndex        =   1
      Top             =   795
      Width           =   1665
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3270
      Left            =   15
      TabIndex        =   26
      Top             =   3660
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   5768
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   12632319
      ForeColorFixed  =   16384
      BackColorSel    =   16777215
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   12632319
      GridColorFixed  =   32896
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
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
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label LblCSVMade 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* CSV File Made *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   225
      Left            =   5625
      TabIndex        =   58
      Top             =   420
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblWarr 
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
      ForeColor       =   &H00C000C0&
      Height          =   225
      Left            =   195
      TabIndex        =   57
      Top             =   390
      Width           =   4260
   End
   Begin VB.Label lblDocId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DocID:"
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
      Left            =   7815
      TabIndex        =   54
      Top             =   420
      Width           =   3555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Claim Catg.(1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   24
      Left            =   3855
      TabIndex        =   53
      Top             =   3045
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Catg.(1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   23
      Left            =   1920
      TabIndex        =   52
      Top             =   3045
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Road Type(1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   22
      Left            =   105
      TabIndex        =   51
      Top             =   3045
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operation Code(1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   21
      Left            =   3855
      TabIndex        =   50
      Top             =   2775
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Load(5)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   20
      Left            =   1920
      TabIndex        =   49
      Top             =   2775
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status Code(1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   19
      Left            =   105
      TabIndex        =   48
      Top             =   2775
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PCR Date(10)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   18
      Left            =   7770
      TabIndex        =   47
      Top             =   1350
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Repair Date(10)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   17
      Left            =   7770
      TabIndex        =   46
      Top             =   1080
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Complaint Date(10)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   16
      Left            =   7770
      TabIndex        =   45
      Top             =   810
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Veh. Sale Date(10)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   15
      Left            =   7290
      TabIndex        =   44
      Top             =   1965
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kms(6)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   14
      Left            =   3675
      TabIndex        =   43
      Top             =   2505
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Dealer(10)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   13
      Left            =   3675
      TabIndex        =   42
      Top             =   2235
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Code(10)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   12
      Left            =   3675
      TabIndex        =   41
      Top             =   1965
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name(50)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   11
      Left            =   3675
      TabIndex        =   40
      Top             =   1695
      Width           =   1710
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RegistrationNo.(10)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   10
      Left            =   7290
      TabIndex        =   39
      Top             =   2235
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis Serial(6)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   9
      Left            =   105
      TabIndex        =   38
      Top             =   2505
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis Year(2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   8
      Left            =   105
      TabIndex        =   37
      Top             =   2235
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis Month(1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   7
      Left            =   105
      TabIndex        =   36
      Top             =   1965
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis Type(6)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   6
      Left            =   105
      TabIndex        =   35
      Top             =   1695
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine/Aggr. No.(17)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   5
      Left            =   105
      TabIndex        =   34
      Top             =   1425
      Width           =   1710
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prowac Year(4)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   4
      Left            =   2295
      TabIndex        =   33
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prowac No(4)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   3
      Left            =   135
      TabIndex        =   32
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Year(2)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   2
      Left            =   5670
      TabIndex        =   31
      Top             =   810
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Date(10)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   1
      Left            =   3195
      TabIndex        =   30
      Top             =   810
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job No.(15)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Index           =   0
      Left            =   135
      TabIndex        =   28
      Top             =   810
      Width           =   960
   End
End
Attribute VB_Name = "frmWarrantyWCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsJob As ADODB.Recordset
Dim RsComplaint  As ADODB.Recordset
Dim RsFailure As ADODB.Recordset
Dim RsMake As ADODB.Recordset
Dim RsJobCode As ADODB.Recordset
Dim mDocId As String
Dim Master As ADODB.Recordset
Dim GridKey As Integer

Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const JobNo As Byte = 0 'To Be save
Private Const JobDt As Byte = 1
Private Const JobYr As Byte = 2
Private Const ProwNo As Byte = 3    'To Be save
Private Const ProwYr As Byte = 4 'To Be save
Private Const Engine As Byte = 5
Private Const ChasType As Byte = 6
Private Const ChasMth As Byte = 7
Private Const ChasYr As Byte = 8
Private Const ChasSrl As Byte = 9
Private Const RegNo As Byte = 10
Private Const CustName As Byte = 11
Private Const DlrCode As Byte = 12   'Cvd
Private Const SellDlrCode As Byte = 13
Private Const Kms As Byte = 14
Private Const VehSaleDt As Byte = 15
Private Const Cmpl_Date As Byte = 16 'To Be save
Private Const Repair_Date As Byte = 17 'To Be save
Private Const PCR_Date As Byte = 18 'To Be save
Private Const Status_Code As Byte = 19 'To Be save
Private Const PayLoad As Byte = 20 ' CVD     To Be save
Private Const OperationCode As Byte = 21 'CVD             To Be save
Private Const RoadCode As Byte = 22 'CVD         To Be save
Private Const Cust_Catg As Byte = 23 'CVD        To Be save
Private Const Claim_Catg As Byte = 24 'CVD   To Be save

' Col Declaration

Private Const AggNo As Byte = 1
Private Const FailureCode As Byte = 2
Private Const CustComplCode As Byte = 3
Private Const ComplFailCode As Byte = 4
Private Const MakeCode As Byte = 5
Private Const MakeCodeRepl As Byte = 6
Private Const Part_No  As Byte = 7
Private Const MRP_YN  As Byte = 8
Private Const Tax_YN  As Byte = 9
Private Const NoOfCompl As Byte = 10
Private Const JobCode As Byte = 11
Private Const Labour_Amt  As Byte = 12
Private Const Spl_Chrg  As Byte = 13
Private Const Misc_Chrg  As Byte = 14
Private Const TotQty  As Byte = 15
Private Const StkQty  As Byte = 16
Private Const FloatQty As Byte = 17
Private Const Price As Byte = 18
Private Const ComplDesc As Byte = 19
Private Const ComplDesc2 As Byte = 20
Private Const ActionTaken As Byte = 21
Private Const SourceCode As Byte = 22
Private Const SPD_InvNo As Byte = 23
Private Const SPD_InvDt As Byte = 24
Private Const IPOSrNo As Byte = 25
Private Const IPODocID As Byte = 26
Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem
Private Sub DGJob_Click()
On Error GoTo ELoop
    If RsJob.RecordCount > 0 Then
        txt(JobNo).TEXT = RsJob!Name
        txt(JobNo).Tag = RsJob!Code
    End If
    txt(JobNo).SetFocus
    DGJob.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub DGPart_Click()
If RsPart.RecordCount > 0 Then
    TxtGrid(0).TEXT = RsPart!Name
End If
If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
DGPart.Visible = False
End Sub
Private Sub FGrid_RowColChange()
If FGrid.Col = ComplDesc Or FGrid.Col = ComplDesc2 Or FGrid.Col = ActionTaken Then
    TxtNarr = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
End If
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
    TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid
    
    
    Dim SiteCond As String
    
    SiteCond = " And PCR_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
        SiteCond = SiteCond & " and   " & cMID("Job_Warr1.DocId", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    Set Master = GCn.Execute("SELECT DocId AS searchcode FROM Job_Warr1 where left(DocID,1) = '" & PubDivCode & "' " & SiteCond & "  order by PCR_Date desc, DociD desc")
    
    Set RsJob = GCn.Execute("SELECT Job_Card.DocId as code, " & cCStr("Job_Card.Job_No") & " as Name,hiscard.model, Job_Card.Job_Date, Job_Card.AtKMsHrs, Job_Card.JobCloseDate, HisCard.Chassis, HisCard.Engine, HisCard.Name as OwnerName, HisCard.RegNo, HisCard.Delivery_Date, HisCard.Chas_Type, HisCard.Dealer_Code " & _
        "FROM Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo  " & _
        "WHERE left(Job_Card.docid,1)='" & PubDivCode & "' Order By Job_Card.Job_No") 'and Not isnull(Job_Card.JobCloseDate)
    Set DGJob.DataSource = RsJob
    Set DGPart.DataSource = RsPart
    
    Set RsComplaint = New ADODB.Recordset
    RsComplaint.CursorLocation = adUseClient
    RsComplaint.Open "select Code,Description from WarrCompMast", GCn, adOpenDynamic, adLockOptimistic
    Set DGComplaint.DataSource = RsComplaint
    
    Set RsFailure = New ADODB.Recordset
    RsFailure.CursorLocation = adUseClient
    RsFailure.Open "select Code,Description from WarrFailMast", GCn, adOpenDynamic, adLockOptimistic
    Set DGFailure.DataSource = RsFailure
    
    Set RsMake = New ADODB.Recordset
    RsMake.CursorLocation = adUseClient
    RsMake.Open "select Code,Description from WarrMakeMast", GCn, adOpenDynamic, adLockOptimistic
    Set DGMake.DataSource = RsMake
    
    Set RsJobCode = New ADODB.Recordset
    RsJobCode.CursorLocation = adUseClient
    RsJobCode.Open "select Code,Description from WarrJobMast", GCn, adOpenDynamic, adLockOptimistic
    Set DGJobCode.DataSource = RsJobCode
    
    Call MoveRec
    Disp_Text SETS("INI", Me, Master)
    txt(PCR_Date).Tag = PubLoginDate
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
Set RsJob = Nothing
Set Master = Nothing
End Sub

Private Sub ListView_Click()
On Error GoTo ELoop
    If TxtGrid(0).Visible = True Then
        TxtGrid(0).TEXT = ListView.SelectedItem.TEXT
        TxtGrid(0).SetFocus
        FrmList.Visible = False
    End If
    txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    txt(Val(ListView.Tag)).SetFocus
    FrmList.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim VouNo As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    txt(PCR_Date) = txt(PCR_Date).Tag
    FGrid.Rows = 2
    FGrid.AddItem FGrid.Rows - 1
    FGrid.FixedRows = 2
    txt(JobNo).SetFocus
    If GCn.Execute("select count(*) from job_warr1 where left(docid,1) = '" & PubDivCode & "' and " & cMID("docid", "2", "2") & " = '" & PubSiteCode & PubSiteCode & "'").Fields(0).Value > 0 Then
        VouNo = GCn.Execute("select MAX(right(Docid,8)) from job_warr1 where left(docid,1) = '" & PubDivCode & "' and " & cMID("docid", "2", "2") & " = '" & PubSiteCode & PubSiteCode & "'").Fields(0).Value + 1
    Else
        VouNo = 1
    End If
    lblDocId = PubDivCode & PubSiteCode & PubSiteCode & "WrClm" & "WrClm" & Space(8 - Len(CStr(VouNo))) & CStr(VouNo)
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim I As Integer
If GCn.Execute("select WBill_DocId from Job_Warr1 where docid = '" & mDocId & "'").Fields(0).Value <> "" Then
    MsgBox "Warranty Bill raised ,Deletion Denied", vbInformation: Exit Sub
End If
If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
    GCn.BeginTrans
    GCn.Execute ("delete from Job_Warr1 where Docid='" & mDocId & "'")
    GCn.Execute ("delete from Job_Warr2 where Docid='" & mDocId & "'")
    GCn.CommitTrans
    Master.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
eloop1:
    If err.NUMBER <> 0 Then
       GCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
    End If
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
 Dim I As Integer
 If GCn.Execute("select WBill_DocId from Job_Warr1 where docid = '" & mDocId & "'").Fields(0).Value <> "" Then
        MsgBox "Warranty Bill raised ,Editing Denied", vbInformation: Exit Sub
    End If
    Disp_Text SETS("EDIT", Me, Master)
    FGrid.AddItem FGrid.Rows - 1
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
On Error GoTo ErrorLoop
Grid_Hide
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim Rst As ADODB.Recordset
Dim I As Integer
Dim PrintStr$, CSVFName$
Dim fob As New FileSystemObject
CSVFName = "W" & Trim(txt(JobNo)) & "-" & Format(txt(PCR_Date), "DD") & ".txt"
If fob.FileExists(Pub_DataPath & "\Warranty\" & CSVFName) = False Then
    If fob.FolderExists(Pub_DataPath & "\Warranty") = False Then
        If PubBackEnd = "A" Then
            fob.CreateFolder (Pub_DataPath & "\Warranty")
            fob.CreateTextFile (Pub_DataPath & "\Warranty\" & CSVFName)
        Else
            fob.CreateFolder (PubBkpPath & "\Warranty")
            fob.CreateTextFile (PubBkpPath & "\Warranty\" & CSVFName)
        End If
    End If
    
End If
Close #1
If PubBackEnd = "A" Then
    Open Pub_DataPath & "\Warranty\" & CSVFName For Output As #1
Else
    Open PubBkpPath & "\Warranty\" & CSVFName For Output As #1
End If

If lblWarr = "CVD Warranty PCR" Then
    GSQL = "SELECT J1.DealerCode,J1.ProwNo, J1.ProwYr, J1.Cust_Catg, " & _
        "HC.Chassis,HC.Dealer_Code, JC.AtKMsHrs,J1.Status_Code, J1.Claim_Catg, " & _
        "'" & txt(VehSaleDt) & "' as VehSaleDate,J1.Cmpl_Date,J1.Repair_Date,J1.PCR_Date,J2.FailureCode, HC.RegNo," & _
        " HC.Name,J1.PayLoad, J1.OperationCode, J1.RoadCode,JC.Job_No,JC.Job_Date," & _
        " J2.ComplDesc, J2.ComplDesc2, J2.ActionTaken, J2.CustComplCode," & _
        "J2.ComplFailCode, J2.MakeCode, J2.MakeCodeRepl,J2.JobCode,J2.NoOfCompl," & _
        "J2.AggNo,J2.PART_NO,J2.SourceCode, J2.Price, J2.Labour_Amt, J2.Spl_Chrg," & _
        "J2.TotQty,J2.SPD_InvNo, J2.SPD_InvDt " & _
        "FROM (Job_Warr1 J1 INNER JOIN Job_Warr2 J2 ON J1.DocID = J2.DocID) " & _
        "LEFT JOIN (Job_Card JC LEFT JOIN HisCard HC ON JC.CardNo = HC.CardNo) " & _
        "ON J1.Job_DocId = JC.DocId where J1.DocID = '" & mDocId & "'"
ElseIf lblWarr = "MUV Warranty PCR" Then
    GSQL = "SELECT JC.Job_No,'" & txt(JobYr) & "' as JobYr,J1.Engine, HC.Chas_Type, " & _
        "'" & txt(ChasMth) & "' as ChasMth,'" & txt(ChasYr) & "' as ChasYr,'" & txt(ChasSrl) & "' as ChasSrl," & _
        "HC.Dealer_Code, JC.AtKMsHrs, J1.Status_Code,'" & txt(VehSaleDt) & "' as VehSaleDate, J1.Cmpl_Date,J1.Repair_Date, " & _
        "J2.FailureCode, J2.CustComplCode,J2.ComplFailCode, J2.MakeCode,J2.PART_NO," & _
        "J2.JobCode, J2.NoOfCompl, J2.Labour_Amt, J2.Spl_Chrg," & _
        "J2.TotQty, J2.StkQty, J2.FloatQty,J2.ComplDesc " & _
        "FROM (Job_Warr1 J1 INNER JOIN Job_Warr2 J2 ON J1.DocID = J2.DocID) " & _
        "LEFT JOIN (Job_Card JC LEFT JOIN HisCard HC ON JC.CardNo = HC.CardNo) " & _
        "ON J1.Job_DocId = JC.DocId where J1.DocID = '" & mDocId & "'"
ElseIf lblWarr = "CAR Warranty PCR" Then
    GSQL = "SELECT JC.Job_No,'" & txt(JobYr) & "' as JobYr,J1.Engine, HC.Chas_Type, " & _
        "'" & txt(ChasMth) & "' as ChasMth,'" & txt(ChasYr) & "' as ChasYr,'" & txt(ChasSrl) & "' as ChasSrl," & _
        "HC.Dealer_Code, JC.AtKMsHrs , J1.Status_Code,'" & txt(VehSaleDt) & "' as VehSaleDate, J1.Cmpl_Date,J1.Repair_Date, J1.PCR_Date, " & _
        "J2.FailureCode, J2.CustComplCode,J2.ComplFailCode, J2.MakeCode,J2.PART_NO," & _
        "J2.JobCode, J2.NoOfCompl, J2.Labour_Amt, J2.Spl_Chrg," & _
        "J2.TotQty, J2.StkQty, J2.FloatQty,J2.ComplDesc " & _
        "FROM (Job_Warr1 J1 INNER JOIN Job_Warr2 J2 ON J1.DocID = J2.DocID) " & _
        "LEFT JOIN (Job_Card JC LEFT JOIN HisCard HC ON JC.CardNo = HC.CardNo) " & _
        "ON J1.Job_DocId = JC.DocId where J1.DocID = '" & mDocId & "'"
End If

Set Rst = GCn.Execute(GSQL)
        Print #1, "DealerCode[10],ProwNo[4],ProwYear[4],CustCatagory[1],ChassisNo[18],Sell.DealerCode[10],KMS[6],StatusCode[1],ClaimCatagory[1],VehSaleDate[10],CmplReportDate[10],VehRepairdate[10],PCRDate[10],FailureCode[1],RegistrationNo[10],CustomerName[50],PayLoad[5],OperationType[1],RoadType[1],JobNo[15],JobDate[10],ComplDesc[240],Investigation[480],ActionTaken[240],CustComplaint[3],Complaintcode[7],MakeCodeFail[6],MakeCodeRepl[6],JobCode[6],NoofCompl[1],AggregateNo[15],PartNo[14],Source[1],Price[8.2],LabChrgs[8.2],SPLLabCharges[8.2],TotalQty[3],SPDInvoiceNo[10],SPDInvoiceDate[10]"
Do Until Rst.EOF
    PrintStr = ""
    If lblWarr = "CVD Warranty PCR" Then
        'PrintStr = """" & Rst!DealerCode & """" & "," & """" & Rst!ProwNo & """" & "," & """" & Rst!ProwYr & """" & "," & """" & Rst!Cust_Catg
        'PrintStr = PrintStr & """" & "," & """" & Rst!Chassis & """" & "," & """" & Rst!dealer_code & """" & "," & """" & Rst!AtKMsHrs & """" & "," & """" & Rst!Status_Code & """" & "," & """" & Rst!Claim_Catg
        'PrintStr = PrintStr & """" & "," & """" & Txt(VehSaleDt) & """" & "," & """" & Rst!Cmpl_Date & """" & "," & """" & Rst!Repair_Date & """" & "," & """" & Rst!PCR_Date & """" & "," & """" & Rst!FailureCode & """" & "," & """" & Rst!RegNo
        'PrintStr = PrintStr & """" & "," & """" & Rst!Name & """" & "," & """" & Rst!PayLoad & """" & "," & """" & Rst!OperationCode & """" & "," & """" & Rst!RoadCode & """" & "," & """" & Rst!Job_No & """" & "," & """" & Format(Rst!Job_Date, "dd/mm/yyyy") & """" & "," & """" & Rst!ComplDesc2
        'PrintStr = PrintStr & """" & "," & """" & Rst!ComplDesc & """" & "," & """" & Rst!ActionTaken & """" & "," & """" & Rst!CustComplCode
        'PrintStr = PrintStr & """" & "," & """" & Rst!ComplFailCode & """" & "," & """" & Rst!MakeCode & """" & "," & """" & Rst!MakeCodeRepl & """" & "," & """" & Rst!JobCode & """" & "," & """" & Rst!NoOfCompl
        'PrintStr = PrintStr & """" & "," & """" & Rst!AggNo & """" & "," & """" & Rst!Part_No & """" & "," & """" & Rst!SourceCode & """" & "," & """" & Rst!Price & """" & "," & """" & Rst!Labour_Amt & """" & "," & """" & Rst!Spl_Chrg
        'PrintStr = PrintStr & """" & "," & """" & Rst!TotQty & """" & "," & """" & Rst!SPD_InvNo & """" & "," & """" & Rst!SPD_InvDt
        
        PrintStr = "" & Rst!DealerCode & "," & Rst!ProwNo & "," & Rst!ProwYr & "," & """" & Rst!Cust_Catg
        PrintStr = PrintStr & """" & "," & """" & Rst!Chassis & """" & "," & Rst!dealer_code & "," & Rst!AtKMsHrs & "," & """" & Rst!Status_Code & """" & "," & """" & Rst!Claim_Catg
        PrintStr = PrintStr & """" & "," & """" & txt(VehSaleDt) & """" & "," & """" & Rst!Cmpl_Date & """" & "," & """" & Rst!Repair_Date & """" & "," & """" & Rst!PCR_Date & """" & "," & """" & Rst!FailureCode & """" & "," & """" & Rst!RegNo
        PrintStr = PrintStr & """" & "," & """" & Rst!Name & """" & "," & Rst!PayLoad & "," & """" & Rst!OperationCode & """" & "," & """" & Rst!RoadCode & """" & "," & """" & Rst!Job_No & """" & "," & """" & RetDate(Rst!Job_Date) & """" & "," & """" & Rst!ComplDesc2
        PrintStr = PrintStr & """" & "," & """" & Rst!ComplDesc & """" & "," & """" & Rst!ActionTaken & """" & "," & """" & Rst!CustComplCode
        PrintStr = PrintStr & """" & "," & """" & Rst!ComplFailCode & """" & "," & """" & Rst!MakeCode & """" & "," & """" & Rst!MakeCodeRepl & """" & "," & """" & Rst!JobCode & """" & "," & Rst!NoOfCompl
        PrintStr = PrintStr & "," & """" & Rst!AggNo & """" & "," & """" & Rst!Part_No & """" & "," & """" & Rst!SourceCode & """" & "," & Rst!Price & "," & Rst!Labour_Amt & "," & Rst!Spl_Chrg
        PrintStr = PrintStr & "," & Rst!TotQty & "," & Rst!SPD_InvNo & "," & """" & Rst!SPD_InvDt
        
        Print #1, PrintStr
       
    ElseIf lblWarr = "MUV Warranty PCR" Then
        PrintStr = PrintStr & """" & "," & """" & Rst!Job_No & """" & "," & """" & txt(JobYr) & """" & "," & """" & Rst!Engine & """" & "," & """" & Rst!Chas_Type
        PrintStr = PrintStr & """" & "," & """" & txt(ChasMth) & """" & "," & """" & txt(ChasYr) & """" & "," & """" & txt(ChasSrl)
        PrintStr = PrintStr & """" & "," & """" & Rst!dealer_code & """" & "," & """" & Rst!AtKMsHrs & """" & "," & """" & Rst!Status_Code & """" & "," & """" & txt(VehSaleDt) & """" & "," & """" & Rst!Cmpl_Date & """" & "," & """" & Rst!Repair_Date
        PrintStr = PrintStr & """" & "," & """" & Rst!FailureCode & """" & "," & """" & Rst!CustComplCode & """" & "," & """" & Rst!ComplFailCode & """" & "," & """" & Rst!MakeCode & """" & "," & """" & Rst!Part_No
        PrintStr = PrintStr & """" & "," & """" & Rst!JobCode & """" & "," & """" & Rst!NoOfCompl & """" & "," & """" & Rst!Labour_Amt & """" & "," & """" & Rst!Spl_Chrg
        PrintStr = PrintStr & """" & "," & """" & Rst!TotQty & """" & "," & """" & Rst!StkQty & """" & "," & """" & Rst!FloatQty & """" & "," & """" & Rst!ComplDesc
        Print #1, PrintStr
    ElseIf lblWarr = "CAR Warranty PCR" Then
        PrintStr = PrintStr & """" & "," & """" & Rst!Job_No & """" & "," & """" & txt(JobYr) & """" & "," & """" & Rst!Engine & """" & "," & """" & Rst!Chas_Type
        PrintStr = PrintStr & """" & "," & """" & txt(ChasMth) & """" & "," & """" & txt(ChasYr) & """" & "," & """" & txt(ChasSrl)
        PrintStr = PrintStr & """" & "," & """" & Rst!dealer_code & """" & "," & """" & Rst!AtKMsHrs & """" & "," & """" & Rst!Status_Code & """" & "," & """" & txt(VehSaleDt) & """" & "," & """" & Rst!Cmpl_Date & """" & "," & """" & Rst!Repair_Date & """" & "," & """" & Rst!PCR_Date
        PrintStr = PrintStr & """" & "," & """" & Rst!FailureCode & """" & "," & """" & Rst!CustComplCode & """" & "," & """" & Rst!ComplFailCode & """" & "," & """" & Rst!MakeCode & """" & "," & """" & Rst!Part_No
        PrintStr = PrintStr & """" & "," & """" & Rst!JobCode & """" & "," & """" & Rst!NoOfCompl & """" & "," & """" & Rst!Labour_Amt & """" & "," & """" & Rst!Spl_Chrg
        PrintStr = PrintStr & """" & "," & """" & Rst!TotQty & """" & "," & """" & Rst!StkQty & """" & "," & """" & Rst!FloatQty & """" & "," & """" & Rst!ComplDesc
        Print #1, PrintStr
    End If
    Rst.MoveNext
Loop
Close #1
MsgBox "Warranty CSV File " & vbCrLf & CSVFName & vbCrLf & "Made!"
End Sub

Private Sub TopCtrl1_eRef()
    RsJob.Requery
End Sub
Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim CardNo As String
    Dim VouNo As Long
    Dim mEditMode As String
'    On Error GoTo errlbl

    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

    If IsValid(txt(JobNo), "Job No") = False Then Exit Sub
    If IsValid(txt(JobDt), "Job Date") = False Then Exit Sub
    If IsValid(txt(Cmpl_Date), "Complaint Date") = False Then Exit Sub
    If IsValid(txt(Repair_Date), "Repair Date") = False Then Exit Sub
    If IsValid(txt(PCR_Date), "PCR Date") = False Then Exit Sub
    
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Part_No) <> "" And (FGrid.TextMatrix(I, MRP_YN) = "" Or FGrid.TextMatrix(I, Tax_YN) = "") Then
            MsgBox "MRP YN/TAX YN in Row No " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = MRP_YN: FGrid.SetFocus: Exit Sub
        End If
    Next
    mDocId = lblDocId
    If TopCtrl1.TopText2 = "Add" And GCn.Execute("select count(*) from job_warr1 where docid = '" & mDocId & "'").Fields(0).Value > 0 Then
        VouNo = GCn.Execute("select MAX(right(Docid,8)) from job_warr1 where left(docid,1) = '" & PubDivCode & "' and " & cMID("DocId", "2", "2") & " = '" & PubSiteCode & PubSiteCode & "'").Fields(0).Value + 1
        lblDocId = PubDivCode & PubSiteCode & PubSiteCode & "WrClm" & "WrClm" & Space(8 - Len(CStr(VouNo))) & CStr(VouNo)
        mDocId = lblDocId
    End If
    
    GCn.BeginTrans
    mTrans = True
    
    If TopCtrl1.TopText2 = "Add" Then
        mEditMode = "A"
    Else
        mEditMode = "E"
    End If
    If mEditMode = "A" Then
        GCn.Execute ("insert into job_warr1(DocID,Div_Code,site_code,ProwNo,ProwYr,Cust_Catg,Claim_Catg,Job_DocId,Engine, " & _
            "Status_Code,PayLoad,OperationCode,RoadCode,DealerCode,Cmpl_Date,Repair_Date,PCR_Date,WBill_DocId,U_Name,U_EntDt,U_AE) values( " & _
            "'" & mDocId & "','" & PubDivCode & "','" & PubSiteCode & "', '" & txt(ProwNo) & "','" & txt(ProwYr) & "','" & txt(Cust_Catg) & "','" & txt(Claim_Catg) & "','" & txt(JobNo).Tag & "','" & txt(Engine) & "', " & _
            "'" & txt(Status_Code) & "','" & txt(PayLoad) & "','" & txt(OperationCode) & "','" & txt(RoadCode) & "','" & txt(DlrCode) & "','" & txt(Cmpl_Date) & "','" & txt(Repair_Date) & "','" & txt(PCR_Date) & "','','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & mEditMode & "')")
    Else
        GCn.Execute ("update job_warr1 set ProwNo='" & txt(ProwNo) & "',ProwYr='" & txt(ProwYr) & "',Cust_Catg='" & txt(Cust_Catg) & "'," & _
            "Claim_Catg='" & txt(Claim_Catg) & "',Job_DocId='" & txt(JobNo).Tag & "',Engine='" & txt(Engine) & "', " & _
            "Status_Code='" & txt(Status_Code) & "',PayLoad='" & txt(PayLoad) & "',OperationCode='" & txt(OperationCode) & "'," & _
            "RoadCode='" & txt(RoadCode) & "',DealerCode='" & txt(DlrCode) & "',Cmpl_Date='" & txt(Cmpl_Date) & "',Repair_Date='" & txt(Repair_Date) & "'," & _
            "PCR_Date='" & txt(PCR_Date) & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='" & mEditMode & "' where DocID = '" & mDocId & "'")
    End If
    GCn.Execute ("delete from Job_Warr2 where Docid='" & mDocId & "'")
    For I = 2 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Part_No) <> "" Or FGrid.TextMatrix(I, JobCode) <> "" Then
            GSQL = "insert into Job_Warr2(DocID,SrlNo,Div_Code,site_code,ProwNo,ProwYr,IPODocID,iPOSrNo,FailureCode,CustComplCode,ComplFailCode, " & _
                "MakeCode,MakeCodeRepl,PART_NO,mrp_yn,tax_yn,JobCode,NoOfCompl,Labour_Amt,Misc_Chrg,Spl_Chrg, " & _
                "TotQty,StkQty,FloatQty,Price,ComplDesc,ComplDesc2,ActionTaken, " & _
                "SourceCode,SPD_InvNo,SPD_InvDt,U_Name,U_EntDt,U_AE,AggNo) Values( " & _
                "'" & mDocId & "'," & I & ",'" & PubDivCode & "','" & PubSiteCode & "', '" & txt(ProwNo) & "','" & txt(ProwYr) & "','" & FGrid.TextMatrix(I, IPODocID) & "'," & Val(FGrid.TextMatrix(I, IPOSrNo)) & ",'" & FGrid.TextMatrix(I, FailureCode) & "','" & FGrid.TextMatrix(I, CustComplCode) & "','" & FGrid.TextMatrix(I, ComplFailCode) & "', " & _
                "'" & FGrid.TextMatrix(I, MakeCode) & "','" & FGrid.TextMatrix(I, MakeCodeRepl) & "','" & FGrid.TextMatrix(I, Part_No) & "'," & IIf(FGrid.TextMatrix(I, MRP_YN) = "No", 0, 1) & "," & IIf(FGrid.TextMatrix(I, Tax_YN) = "No", 0, 1) & ", '" & FGrid.TextMatrix(I, JobCode) & "'," & Val(FGrid.TextMatrix(I, NoOfCompl)) & "," & Val(FGrid.TextMatrix(I, Labour_Amt)) & "," & Val(FGrid.TextMatrix(I, Misc_Chrg)) & "," & Val(FGrid.TextMatrix(I, Spl_Chrg)) & ", " & _
                "" & Val(FGrid.TextMatrix(I, TotQty)) & "," & Val(FGrid.TextMatrix(I, StkQty)) & "," & Val(FGrid.TextMatrix(I, FloatQty)) & "," & Val(FGrid.TextMatrix(I, Price)) & ",'" & FGrid.TextMatrix(I, ComplDesc) & "','" & FGrid.TextMatrix(I, ComplDesc2) & "','" & FGrid.TextMatrix(I, ActionTaken) & "', " & _
                "'" & FGrid.TextMatrix(I, SourceCode) & "','" & FGrid.TextMatrix(I, SPD_InvNo) & "','" & FGrid.TextMatrix(I, SPD_InvDt) & "' ,'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & mEditMode & "','" & FGrid.TextMatrix(I, AggNo) & "')"
            
            If FGrid.TextMatrix(I, IPODocID) <> "" And FGrid.TextMatrix(I, IPOSrNo) <> "" Then
                GCn.Execute ("update sp_stock set ClaimId = '" & mDocId & "' where DocID= '" & FGrid.TextMatrix(I, IPODocID) & "' and Srl_No = " & Val(FGrid.TextMatrix(I, IPOSrNo)) & "")
            End If
            GCn.Execute GSQL
        End If
    Next

GCn.CommitTrans
mTrans = False
    Master.Requery
    Master.FIND "SearchCode = '" & mDocId & "'"
    If mEditMode = "A" Then
        txt(PCR_Date).Tag = txt(PCR_Date)
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    Dim SiteCond As String
    SiteCond = " And PCR_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " where   " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "'"
    End If
    GSQL = "SELECT DociD AS SearchCode, Job_Warr1.DocID, Job_Warr1.ProwNo, Job_Warr1.ProwYr, Job_Warr1.Job_DocId, Job_Warr1.Cmpl_Date, Job_Warr1.Repair_Date, Job_Warr1.PCR_Date FROM Job_Warr1 " & SiteCond & " order by docid"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND "searchcode='" & MyValue & "'"
     BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub Txt_GotFocus(Index As Integer)
TxtGrid(0).Visible = False
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case OperationCode
        ListArray = Array("1 - > Drive Away", "2 - > Long Route", "3 - > City Route", "4 - > Construction", "5 - > Mining", "6 - > Forest", "7 - > Marine", "8 - > Others")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 8)
    Case RoadCode
        ListArray = Array("1 - > Plain Metalled", "2 - > Plain Kutcha", "3 - > Off Road", "4 - > Hilly Metalled", "5 - > Hilly Kutcha", "6 - > Desert", "7 - > Others")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 7)
    Case Claim_Catg
        ListArray = Array("1 - > Normal Warranty", "2 - > Repeat Failure", "3 - > Goodwill Commercial", "4 - > Retrofitment", "5 - > PIW", "6 - > Industrial Engine", "7 - > Drive Away", "8 - > Pre Sale Total Loss", "9 - > Spare Part Claim", "0 -  > Cummins Engine", "a -  > Insurance Claim", "b - > Goodwill Technical", "c - > Goodwill Spare", "d - > Atlas Copco", "e - > 207 DI Engine")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 15)
    Case Cust_Catg
        ListArray = Array("1 - > Retail Customer", "2 - > STUs", "3 - > Govt. Organisation", "4 - > Army")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 4)
    Case Status_Code
            ListArray = Array("1 - > Drive Away", "2 - > Sold")
            Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
    Case JobNo
    If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
    If txt(Index).Tag <> RsJob!Code Then
        RsJob.MoveFirst
        RsJob.FIND "Code ='" & txt(Index).Tag & "'"
    End If
End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Byte
Dim Txtdate As Boolean
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case JobNo
        DGridTxtKeyDown DGJob, txt, JobNo, RsJob, KeyCode, False, 1
    Case OperationCode
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), 2000, 2400
    Case Claim_Catg
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), 2500, 4200
    Case Cust_Catg
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), 2000, 1200
    Case RoadCode
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), 2000, 2100
    Case Status_Code
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), 2000, 800
End Select
If DGJob.Visible = False And FrmList.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> JobNo Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> ProwNo Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case JobNo
        If DGJob.Visible = True Then DGridTxtKeyPress txt, JobNo, RsJob, KeyAscii, "Name"
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case OperationCode, RoadCode, Claim_Catg, Cust_Catg, Status_Code
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case JobNo
        If RsJob.RecordCount <> 0 And Trim(txt(Index).TEXT = "") Then
            MsgBox "Please Select Job No.", vbInformation, "Information"
            txt(Index).SetFocus
            Cancel = True
            Exit Sub
        End If
        If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then
            txt(JobNo).TEXT = ""
            txt(JobNo).Tag = ""
        Else
            txt(JobNo).TEXT = RsJob!Name
            txt(JobNo).Tag = RsJob!Code
            FillData 'Then TopCtrl1_eCancel
        End If
    Case Cmpl_Date, Repair_Date, PCR_Date
        If Len(Trim(txt(Index).TEXT)) = 0 Then
             txt(Index).TEXT = PubLoginDate
        Else
            txt(Index).TEXT = RetDate(RetDate(txt(Index)))
        End If
    Case OperationCode, RoadCode, Claim_Catg, Cust_Catg, Status_Code
        txt(Index).TEXT = left(ListView.SelectedItem.TEXT, 1)
End Select
End Sub

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress (vbKeyReturn)
TAddMode = False
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case FailureCode, CustComplCode, ComplFailCode, MakeCode, MakeCodeRepl, Part_No, JobCode, NoOfCompl, Labour_Amt, Spl_Chrg, TotQty, StkQty, FloatQty, Price, ComplDesc, ComplDesc2, ActionTaken, SourceCode, SPD_InvNo, SPD_InvDt
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
End If
If KeyCode = 13 Then TAddMode = False
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
SetMaxLength
Select Case FGrid.Col
    Case AggNo, FailureCode, CustComplCode, ComplFailCode, MakeCode, MakeCodeRepl, Part_No, JobCode, NoOfCompl, ComplDesc, ComplDesc2, ActionTaken, SourceCode, SPD_InvNo, SPD_InvDt
       Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    Case Labour_Amt, Spl_Chrg, TotQty, StkQty, FloatQty, Price, Misc_Chrg
       Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
    Case MRP_YN, Tax_YN
        If UCase(Chr(KeyAscii)) = "N" Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "No"
        ElseIf UCase(Chr(KeyAscii)) = "Y" Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Yes"
        End If
        KeyAscii = 0
        FGrid.Col = FGrid.Col + 1
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub
Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid.Rows > 2 Then
                FGrid.RemoveItem (FGrid.Row)
            Else
                FGrid.Rows = 2
                FGrid.AddItem FGrid.Rows - 1
                FGrid.FixedRows = 2
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
Dim Master1 As ADODB.Recordset, RsJob1 As ADODB.Recordset, Rs As ADODB.Recordset, I As Integer
Dim WarrType As Byte
On Error GoTo error1
If Master.RecordCount > 0 Then
    Set Master1 = GCn.Execute("select * from Job_warr1 where docid = '" & Master!SearchCode & "'")
    Set RsJob1 = GCn.Execute("SELECT Job_Card.DocId, Job_Card.Job_No, Job_Card.Job_Date, Job_Card.AtKMsHrs, Job_Card.JobCloseDate, HisCard.Chassis, HisCard.Engine,HisCard.Model, HisCard.Name, HisCard.RegNo, HisCard.Delivery_Date, HisCard.Chas_Type, HisCard.Dealer_Code " & _
        "FROM Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo  " & _
        "WHERE Job_Card.DocID = '" & Master1!job_docid & "' and left(Job_Card.docid,1)='" & PubDivCode & "'  Order By Job_Card.Job_No")    'and Not isnull(Job_Card.JobCloseDate)
    lblDocId = Master1!DocID
    mDocId = lblDocId
    Set Rs = GCn.Execute("SELECT Vehicle_Type.Warr_Type FROM Model LEFT JOIN Vehicle_Type ON Model.Vehicle_Type = Vehicle_Type.Vehicle_Type where Model.Model ='" & RsJob1!Model & "'")
    If Rs.RecordCount > 0 Then WarrType = VNull(Rs(0))
    lblWarr = IIf(WarrType = 0, "CVD Warranty PCR", IIf(WarrType = 1, "MUV Warranty PCR", "CAR Warranty PCR"))
    txt(JobNo) = Trim(DeCodeDocID(Master1!job_docid, Document_No))
    txt(JobNo).Tag = RsJob1!DocID
    txt(JobDt) = RetDate(RsJob1!Job_Date)
    txt(JobYr) = Format(RsJob1!Job_Date, "YY")
    txt(ProwNo) = Master1!ProwNo
    txt(ProwYr) = Master1!ProwYr
    txt(Engine) = Master1!Engine
    txt(ChasType) = DeCodeChassis(RsJob1!Chassis, ChasType)
    txt(ChasMth) = DeCodeChassis(RsJob1!Chassis, MfgMonth)
    txt(ChasYr) = DeCodeChassis(RsJob1!Chassis, MfgYear)
    txt(ChasSrl) = DeCodeChassis(RsJob1!Chassis, ChasSerialNo)
    txt(RegNo) = RsJob1!RegNo
    txt(CustName) = RsJob1!Name
    txt(DlrCode) = GCn.Execute("select Dealer_ID from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
    txt(SellDlrCode) = RsJob1!dealer_code
    txt(Kms) = RsJob1!AtKMsHrs
    txt(VehSaleDt) = RetDate(RsJob1!Delivery_Date)
    txt(Cmpl_Date) = RetDate(Master1!Cmpl_Date)
    txt(Repair_Date) = RetDate(Master1!Repair_Date)
    txt(PCR_Date) = RetDate(Master1!PCR_Date)
    txt(Status_Code) = Master1!Status_Code
    txt(PayLoad) = Master1!PayLoad
    txt(OperationCode) = Master1!OperationCode
    txt(RoadCode) = Master1!RoadCode
    txt(Cust_Catg) = Master1!Cust_Catg
    txt(Claim_Catg) = Master1!Claim_Catg
    
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT * from job_warr2 where job_warr2.docid = '" & Master!SearchCode & "'")
    FGrid.Rows = 2
    If Rs.RecordCount > 0 Then
        I = 2
        Do Until Rs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, 0) = Rs!Srlno
                .TextMatrix(I, IPODocID) = Rs!IPODocID
                .TextMatrix(I, IPOSrNo) = Rs!IPOSrNo
                .TextMatrix(I, AggNo) = XNull(Rs!AggNo)
                .TextMatrix(I, FailureCode) = XNull(Rs!FailureCode)
                .TextMatrix(I, CustComplCode) = Rs!CustComplCode
                .TextMatrix(I, ComplFailCode) = Rs!ComplFailCode
                .TextMatrix(I, MakeCode) = XNull(Rs!MakeCode)
                .TextMatrix(I, MakeCodeRepl) = Rs!MakeCodeRepl
                .TextMatrix(I, Part_No) = Rs!Part_No
                .TextMatrix(I, MRP_YN) = IIf(Rs!MRP_YN = 0, "No", "Yes")
                .TextMatrix(I, Tax_YN) = IIf(Rs!Tax_YN = 0, "No", "Yes")
                .TextMatrix(I, JobCode) = XNull(Rs!JobCode)
                .TextMatrix(I, NoOfCompl) = XNull(Rs!NoOfCompl)
                .TextMatrix(I, Labour_Amt) = Rs!Labour_Amt
                .TextMatrix(I, Misc_Chrg) = Rs!Misc_Chrg
                .TextMatrix(I, Spl_Chrg) = Rs!Spl_Chrg
                .TextMatrix(I, TotQty) = Rs!TotQty
                .TextMatrix(I, StkQty) = Rs!StkQty
                .TextMatrix(I, FloatQty) = Rs!FloatQty
                .TextMatrix(I, Price) = Rs!Price
                .TextMatrix(I, ComplDesc) = Rs!ComplDesc
                .TextMatrix(I, ComplDesc2) = Rs!ComplDesc2
                .TextMatrix(I, ActionTaken) = Rs!ActionTaken
                .TextMatrix(I, SourceCode) = Rs!SourceCode
                .TextMatrix(I, SPD_InvNo) = Rs!SPD_InvNo
                .TextMatrix(I, SPD_InvDt) = Rs!SPD_InvDt
            End With
            Rs.MoveNext
            I = I + 1
        Loop
    End If
    Set Rs = Nothing
Else
    Call BlankText
End If
If FGrid.Rows = 2 Then FGrid.AddItem FGrid.Rows - 1
FGrid.FixedRows = 2
Set Rs = Nothing
Set RsJob1 = Nothing
Set Master1 = Nothing
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Ini_Grid()
    With FGrid
'        .FixedRows = 2
        .Cols = 27
        .left = Me.left '+45
        .width = Me.width - 90
        .top = 3660
        .RowHeightMin = PubGridRowHeight
        
        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450
        .ColWidth(IPODocID) = 1000
        .ColWidth(IPOSrNo) = 450
        
        .TextMatrix(0, AggNo) = "Eng./Agg.No."
        .TextMatrix(1, AggNo) = "    (15)"
        .ColAlignmentFixed(AggNo) = flexAlignCenterCenter
        .ColAlignment(AggNo) = flexAlignLeftCenter
        .ColWidth(AggNo) = 1620
        
        .TextMatrix(0, FailureCode) = "Failure"
        .TextMatrix(1, FailureCode) = "Code(1)"
        .ColAlignmentFixed(FailureCode) = flexAlignCenterCenter
        .ColAlignment(FailureCode) = flexAlignLeftCenter
        .ColWidth(FailureCode) = 645
        
        .TextMatrix(0, CustComplCode) = "CustCompl"
        .TextMatrix(1, CustComplCode) = "Code(3)"
        .ColAlignmentFixed(CustComplCode) = flexAlignCenterCenter
        .ColAlignment(CustComplCode) = flexAlignLeftCenter
        .ColWidth(CustComplCode) = 840
        
        .TextMatrix(0, ComplFailCode) = "ComplFail"
        .TextMatrix(1, ComplFailCode) = "Code(7)"
        .ColAlignmentFixed(ComplFailCode) = flexAlignCenterCenter
        .ColAlignment(ComplFailCode) = flexAlignLeftCenter
        .ColWidth(ComplFailCode) = 825
        
        .TextMatrix(0, MakeCode) = "Make"
        .TextMatrix(1, MakeCode) = "Code(6)"
        .ColAlignmentFixed(MakeCode) = flexAlignCenterCenter
        .ColAlignment(MakeCode) = flexAlignLeftCenter
        .ColWidth(MakeCode) = 690

        .TextMatrix(0, MakeCodeRepl) = "MakeCode"
        .TextMatrix(1, MakeCodeRepl) = "Repl(6)"
        .ColAlignmentFixed(MakeCodeRepl) = flexAlignCenterCenter
        .ColAlignment(MakeCodeRepl) = flexAlignLeftCenter
        .ColWidth(MakeCodeRepl) = 855

        .TextMatrix(0, Part_No) = "PART NO"
        .TextMatrix(1, Part_No) = "(14)"
        .ColAlignmentFixed(Part_No) = flexAlignCenterCenter
        .ColAlignment(Part_No) = flexAlignLeftCenter
        .ColWidth(Part_No) = 1620
        
        .TextMatrix(0, MRP_YN) = "MRP"
        .TextMatrix(1, MRP_YN) = "Y/N"
        .ColAlignmentFixed(MRP_YN) = flexAlignCenterCenter
        .ColAlignment(MRP_YN) = flexAlignLeftCenter
        .ColWidth(MRP_YN) = 435
        
        .TextMatrix(0, Tax_YN) = "TAX"
        .TextMatrix(1, Tax_YN) = "Y/N"
        .ColAlignmentFixed(Tax_YN) = flexAlignCenterCenter
        .ColAlignment(Tax_YN) = flexAlignLeftCenter
        .ColWidth(Tax_YN) = 435
        
        .TextMatrix(0, NoOfCompl) = "No.of"
        .TextMatrix(1, NoOfCompl) = "Compl(1)"
        .ColAlignmentFixed(NoOfCompl) = flexAlignCenterCenter
        .ColAlignment(NoOfCompl) = flexAlignLeftCenter
        .ColWidth(NoOfCompl) = 735
        

        .TextMatrix(0, JobCode) = "Job"
        .TextMatrix(1, JobCode) = "Code(6)"
        .ColAlignmentFixed(JobCode) = flexAlignCenterCenter
        .ColAlignment(JobCode) = flexAlignLeftCenter
        .ColWidth(JobCode) = 645
        
        .TextMatrix(0, Price) = "Price"
        .TextMatrix(1, Price) = "(10)"
        .ColAlignmentFixed(Price) = flexAlignCenterCenter
        .ColAlignment(Price) = flexAlignLeftCenter
        .ColWidth(Price) = 975
        
        .TextMatrix(0, Labour_Amt) = "Labour_Amt"
        .TextMatrix(1, Labour_Amt) = "(8,2)"
        .ColAlignmentFixed(Labour_Amt) = flexAlignCenterCenter
        .ColAlignment(Labour_Amt) = flexAlignLeftCenter
        .ColWidth(Labour_Amt) = 1215
        
        .TextMatrix(0, Spl_Chrg) = "Spl Lab"
        .TextMatrix(1, Spl_Chrg) = "(8,2)"
        .ColAlignmentFixed(Spl_Chrg) = flexAlignCenterCenter
        .ColAlignment(Spl_Chrg) = flexAlignLeftCenter
        .ColWidth(Spl_Chrg) = 1215
        
        .TextMatrix(0, Misc_Chrg) = "Misc_Chrg"
        .TextMatrix(1, Misc_Chrg) = "(10)"
        .ColAlignmentFixed(Misc_Chrg) = flexAlignCenterCenter
        .ColAlignment(Misc_Chrg) = flexAlignLeftCenter
        .ColWidth(Misc_Chrg) = 1215
        
        .TextMatrix(0, TotQty) = "TotQty"
        .TextMatrix(1, TotQty) = "(3)"
        .ColAlignmentFixed(TotQty) = flexAlignCenterCenter
        .ColAlignment(TotQty) = flexAlignLeftCenter
        .ColWidth(TotQty) = 555

        .TextMatrix(0, StkQty) = "StkQty"
        .TextMatrix(1, StkQty) = "(1)"
        .ColAlignmentFixed(StkQty) = flexAlignCenterCenter
        .ColAlignment(StkQty) = flexAlignLeftCenter
        .ColWidth(StkQty) = 645

        .TextMatrix(0, FloatQty) = "FloatQty"
        .TextMatrix(1, FloatQty) = "(1)"
        .ColAlignmentFixed(FloatQty) = flexAlignCenterCenter
        .ColAlignment(FloatQty) = flexAlignLeftCenter
        .ColWidth(FloatQty) = 705
        
        .TextMatrix(0, ComplDesc) = "ComplDesc (Investi for CVD)"
        .TextMatrix(1, ComplDesc) = "(480)"
        .ColAlignmentFixed(ComplDesc) = flexAlignCenterCenter
        .ColAlignment(ComplDesc) = flexAlignLeftCenter
        .ColWidth(ComplDesc) = 3135
        
        .TextMatrix(0, ComplDesc2) = "ComplDesc2(CVD)"
        .TextMatrix(1, ComplDesc2) = "(240)"
        .ColAlignmentFixed(ComplDesc2) = flexAlignCenterCenter
        .ColAlignment(ComplDesc2) = flexAlignLeftCenter
        .ColWidth(ComplDesc2) = 3135
        
        .TextMatrix(0, ActionTaken) = "ActionTaken(CVD)"
        .TextMatrix(1, ActionTaken) = "(240)"
        .ColAlignmentFixed(ActionTaken) = flexAlignCenterCenter
        .ColAlignment(ActionTaken) = flexAlignLeftCenter
        .ColWidth(ActionTaken) = 3135
        
        .TextMatrix(0, SourceCode) = "Source"
        .TextMatrix(1, SourceCode) = "Code(1)"
        .ColAlignmentFixed(SourceCode) = flexAlignCenterCenter
        .ColAlignment(SourceCode) = flexAlignLeftCenter
        .ColWidth(SourceCode) = 735
        
        .TextMatrix(0, SPD_InvNo) = "SPD_InvNo"
        .TextMatrix(1, SPD_InvNo) = "(10)"
        .ColAlignmentFixed(SPD_InvNo) = flexAlignCenterCenter
        .ColAlignment(SPD_InvNo) = flexAlignLeftCenter
        .ColWidth(SPD_InvNo) = 1380
        
        .TextMatrix(0, SPD_InvDt) = "SPD_InvDt"
        .TextMatrix(1, SPD_InvDt) = "(10)"
        .ColAlignmentFixed(SPD_InvDt) = flexAlignCenterCenter
        .ColAlignment(SPD_InvDt) = flexAlignLeftCenter
        .ColWidth(SPD_InvDt) = 1380
End With
BackColorSelLeave = FGrid.BackColorSel
ForeColorSelEnter = FGrid.ForeColorSel
DGPart.left = FGrid.left: DGPart.top = mTopScale
DGJob.left = FGrid.left: DGJob.top = FGrid.top: DGJob.height = FGrid.height
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
    txt(I).ForeColor = CtrlFColOrg
Next
    txt(JobDt).Enabled = False
    txt(JobYr).Enabled = False
    txt(ChasType).Enabled = False
    txt(ChasMth).Enabled = False
    txt(ChasYr).Enabled = False
    txt(ChasSrl).Enabled = False
    txt(RegNo).Enabled = False
    txt(CustName).Enabled = False
    txt(DlrCode).Enabled = False
    txt(SellDlrCode).Enabled = False
    txt(VehSaleDt).Enabled = False
    txt(Kms).Enabled = False

txtDisabled_Color Me
TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol
End Sub
Private Sub Grid_Hide()
    If ListView.Visible = True Then ListView.Visible = False
    If DGJob.Visible = True Then DGJob.Visible = False
    If DGPart.Visible = True Then DGPart.Visible = False
'    If TxtNarr.Visible = True Then TxtNarr.Visible = False
End Sub
Private Sub TxtGrid_GotFocus(Index As Integer)
    Grid_Hide
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
        Case FailureCode
            ListArray = Array("1 - > OE", "2 - > Repeat", "3 - > Spare Parts")
            Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
        Case ComplDesc2, ComplDesc, ActionTaken
            TxtGrid(0).height = FGrid.RowHeight(0) * 3
'            If TxtNarr.Visible = False Then TxtNarr.Text = TxtGrid(0).Text: TxtNarr.Visible = True: TxtNarr.top = FGrid.top - TxtNarr.Height - 20: TxtNarr.left = FGrid.CellLeft
        Case Part_No
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "CODE"
            If FGrid.TextMatrix(FGrid.Row, Part_No) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, Part_No) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case CustComplCode
            'DGComplaint.left = txtGrid(0).left: DGJob.top = txtGrid(0).top + txtGrid(0).height
            If RsComplaint.RecordCount = 0 Or (RsComplaint.EOF = True Or RsComplaint.BOF = True) Then Exit Sub
            RsComplaint.Sort = "CODE"
            If FGrid.TextMatrix(FGrid.Row, CustComplCode) <> "" Then
                RsComplaint.MoveFirst
                RsComplaint.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, CustComplCode) & "'"
                If RsComplaint.EOF = True Then RsComplaint.MoveFirst
            End If
        Case ComplFailCode
            'DGFailure.left = txtGrid(0).left: DGJob.top = txtGrid(0).top + txtGrid(0).height
            If RsFailure.RecordCount = 0 Or (RsFailure.EOF = True Or RsFailure.BOF = True) Then Exit Sub
            RsFailure.Sort = "CODE"
            If FGrid.TextMatrix(FGrid.Row, ComplFailCode) <> "" Then
                RsFailure.MoveFirst
                RsFailure.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, ComplFailCode) & "'"
                If RsFailure.EOF = True Then RsFailure.MoveFirst
            End If
        Case MakeCode, MakeCodeRepl
            'DGMake.left = txtGrid(0).left: DGJob.top = txtGrid(0).top + txtGrid(0).height
            If RsMake.RecordCount = 0 Or (RsMake.EOF = True Or RsMake.BOF = True) Then Exit Sub
            RsMake.Sort = "CODE"
            If FGrid.TextMatrix(FGrid.Row, MakeCode) <> "" Then
                RsMake.MoveFirst
                RsMake.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, MakeCode) & "'"
                If RsMake.EOF = True Then RsMake.MoveFirst
            End If
        Case JobCode
            'DGMake.left = txtGrid(0).left: DGJob.top = txtGrid(0).top + txtGrid(0).height
            If RSOJPR = True Then
                Set RsJobCode = New ADODB.Recordset
                RsJobCode.CursorLocation = adUseClient
                RsJobCode.Open "select WarrJobMast.Code,WarrJobMast.Description from WarrJobMast Left Join Job_Lab on WarrJobMast.Code=Job_Lab.JobCode where Job_Lab.Job_DocId='" & txt(JobNo).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
                Set DGJobCode.DataSource = RsJobCode
            End If
            If RsJobCode.RecordCount = 0 Or (RsJobCode.EOF = True Or RsJobCode.BOF = True) Then Exit Sub
            RsJobCode.Sort = "CODE"
            If FGrid.TextMatrix(FGrid.Row, JobCode) <> "" Then
                RsJobCode.MoveFirst
                RsJobCode.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, JobCode) & "'"
                If RsJobCode.EOF = True Then RsJobCode.MoveFirst
            End If
    End Select
End Sub
Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then TxtGrid(0) = TxtGrid(0).Tag: Exit Sub
            Select Case FGrid.Col
                Case Part_No
                     DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 0, frmPartMast, "frmPartMast"
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SPD_InvDt
                        End If
                    End If
                Case CustComplCode
                     DGridTxtKeyDown DGComplaint, TxtGrid, Index, RsComplaint, KeyCode, True, 0
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SPD_InvDt
                        End If
                    End If
                Case ComplFailCode
                     DGridTxtKeyDown DGFailure, TxtGrid, Index, RsFailure, KeyCode, True, 0
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SPD_InvDt
                        End If
                    End If
                Case MakeCode, MakeCodeRepl
                     DGridTxtKeyDown DGMake, TxtGrid, Index, RsMake, KeyCode, True, 0
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SPD_InvDt
                        End If
                    End If
                Case JobCode
                     DGridTxtKeyDown DGJobCode, TxtGrid, Index, RsJobCode, KeyCode, True, 0
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SPD_InvDt
                        End If
                    End If
                Case ComplDesc2, ComplDesc, ActionTaken
                    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SPD_InvDt
                        End If
                    End If
                Case CustComplCode, ComplFailCode, MakeCode, JobCode, NoOfCompl, SourceCode, SPD_InvNo, SPD_InvDt, Labour_Amt, Spl_Chrg, TotQty, StkQty, FloatQty, Price, Misc_Chrg
                    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SPD_InvDt
                        End If
                    End If
                Case FailureCode
                    ListView_KeyDown FrmList, ListView, TxtGrid, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), 1500, 900
                    If FrmList.Visible = False Then
                        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                            If TxtGridLeave = True Then
                                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, SPD_InvDt
                            End If
                        End If
                    End If
            End Select
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If KeyAscii = vbKeyEscape Then Exit Sub
Call CheckQuote(KeyAscii)
Select Case FGrid.Col
    Case Part_No
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsPart, KeyAscii, "CODE"
    Case ComplFailCode
        If DGFailure.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsFailure, KeyAscii, "CODE"
    Case CustComplCode
        If DGComplaint.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsComplaint, KeyAscii, "CODE"
    Case MakeCode, MakeCodeRepl
        If DGMake.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsMake, KeyAscii, "CODE"
    Case JobCode
        If RSOJPR = True Then
            Set RsJobCode = New ADODB.Recordset
            RsJobCode.CursorLocation = adUseClient
            RsJobCode.Open "select WarrJobMast.Code,WarrJobMast.Description from WarrJobMast Left Join Job_Lab on WarrJobMast.Code=Job_Lab.JobCode where Job_Lab.Job_DocId='" & txt(JobNo).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
            Set DGJobCode.DataSource = RsJobCode
        End If
        If DGJobCode.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsJobCode, KeyAscii, "CODE"
End Select
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case ComplDesc2, ComplDesc, ActionTaken
        TxtNarr = TxtGrid(0)
    Case Part_No
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0:   DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "CODE", True
    Case CustComplCode
        If KeyCode <> 13 And DGComplaint.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0:   DGridTxtKeyPress TxtGrid, Index, RsComplaint, KeyCode, "CODE", True
    Case ComplFailCode
        If KeyCode <> 13 And DGFailure.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0:   DGridTxtKeyPress TxtGrid, Index, RsFailure, KeyCode, "CODE", True
    Case MakeCode, MakeCodeRepl
        If KeyCode <> 13 And DGMake.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0:   DGridTxtKeyPress TxtGrid, Index, RsMake, KeyCode, "CODE", True
    Case JobCode
        If KeyCode <> 13 And DGJobCode.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0:   DGridTxtKeyPress TxtGrid, Index, RsJobCode, KeyCode, "CODE", True
    Case CustComplCode
        If KeyCode <> 13 And DGComplaint.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0:   DGridTxtKeyPress TxtGrid, Index, RsComplaint, KeyCode, "CODE", True
    Case FailureCode
        If FrmList.Visible = True Then ListView_KeyUp ListView, TxtGrid, Index, KeyCode, mListItem
    Case JobCode
        If KeyCode <> 13 And DGJobCode.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0:   DGridTxtKeyPress TxtGrid, Index, RsJobCode, KeyCode, "CODE", True
End Select
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
    Grid_Hide
End If
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
Select Case FGrid.Col
        Case Part_No
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, Part_No) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, Part_No) = RsPart!Code
            End If
        Case CustComplCode
            If RsComplaint.RecordCount = 0 Or (RsComplaint.EOF = True Or RsComplaint.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, CustComplCode) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, CustComplCode) = RsComplaint!Code
            End If
        Case ComplFailCode
            If RsFailure.RecordCount = 0 Or (RsFailure.EOF = True Or RsFailure.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, ComplFailCode) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, ComplFailCode) = RsFailure!Code
            End If
        Case JobCode
            If RsJobCode.RecordCount = 0 Or (RsJobCode.EOF = True Or RsJobCode.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, JobCode) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, JobCode) = RsJobCode!Code
            End If
        Case ComplDesc2, ComplDesc, ActionTaken
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
        Case AggNo, FailureCode, CustComplCode, ComplFailCode, MakeCode, MakeCodeRepl, JobCode, NoOfCompl, SourceCode, SPD_InvNo, Labour_Amt, Spl_Chrg, TotQty, StkQty, FloatQty, Price, Misc_Chrg
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
        Case SPD_InvDt
            TxtGrid(0).TEXT = RetDate(RetDate(TxtGrid(0)))
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
        
    End Select
    TxtGridLeave = True
    If ValidateCall = False Then
        FGrid.SetFocus
        TxtGrid(0).Visible = False
    End If
End Function
Private Function FillData() As Boolean
Dim Rs As ADODB.Recordset, RsJob1 As ADODB.Recordset, I As Integer
Dim WarrType As Byte
If RsJob.EOF = False And RsJob.BOF = False And RsJob.RecordCount > 0 Then
    Set RsJob1 = GCn.Execute("SELECT Job_Card.DocId, Job_Card.Job_No, Job_Card.Job_Date, Job_Card.AtKMsHrs, Job_Card.JobCloseDate,HisCard.model, HisCard.Chassis, HisCard.Engine, HisCard.Name, HisCard.RegNo, HisCard.Delivery_Date, HisCard.Chas_Type, HisCard.Dealer_Code " & _
        "FROM Job_Card LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo  " & _
        "WHERE Job_Card.DocID  = '" & RsJob!Code & "'")
    Set Rs = GCn.Execute("SELECT Vehicle_Type.Warr_Type FROM Model LEFT JOIN Vehicle_Type ON Model.Vehicle_Type = Vehicle_Type.Vehicle_Type where Model.Model ='" & RsJob1!Model & "'")
    
    If Rs.RecordCount > 0 Then WarrType = VNull(Rs(0))
    lblWarr = IIf(WarrType = 0, "CVD Warranty PCR", IIf(WarrType = 1, "MUV Warranty PCR", "CAR Warranty PCR"))
    txt(JobNo) = Trim(DeCodeDocID(RsJob1!DocID, Document_No))
    txt(JobNo).Tag = RsJob1!DocID
    txt(JobDt) = RetDate(RsJob1!Job_Date)
    txt(JobYr) = Format(RsJob1!Job_Date, "YY")
    txt(Engine) = RsJob1!Engine
    txt(ChasType) = DeCodeChassis(RsJob1!Chassis, ChasType)
    txt(ChasMth) = DeCodeChassis(RsJob1!Chassis, MfgMonth)
    txt(ChasYr) = DeCodeChassis(RsJob1!Chassis, MfgYear)
    txt(ChasSrl) = DeCodeChassis(RsJob1!Chassis, ChasSerialNo)
    txt(RegNo) = RsJob1!RegNo
    txt(CustName) = RsJob1!Name
    txt(DlrCode) = GCn.Execute("select Dealer_ID from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
    txt(SellDlrCode) = RsJob1!dealer_code
    txt(Kms) = RsJob1!AtKMsHrs
    txt(VehSaleDt) = RetDate(RsJob1!Delivery_Date)
    txt(Cmpl_Date) = RetDate(RsJob1!Job_Date)
    txt(Repair_Date) = RetDate(RsJob1!JobCloseDate)
   'txt(PCR_Date) = txt(PCR_Date).Tag
    txt(PCR_Date) = RetDate(RsJob1!Job_Date)

    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT * from sp_stock where sp_stock.Job_DocID = '" & RsJob1!DocID & "' and sp_stock.purpose = 'W'")
    FGrid.Rows = 2
    If Rs.RecordCount > 0 Then
        I = 2
        Do Until Rs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, 0) = I - 1
                .TextMatrix(I, IPODocID) = Rs!DocID
                .TextMatrix(I, IPOSrNo) = Rs!Srl_No
                .TextMatrix(I, Part_No) = Rs!Part_No
                .TextMatrix(I, MRP_YN) = IIf(Rs!MRP_YN = 0, "No", "Yes")
                .TextMatrix(I, Tax_YN) = IIf(Rs!Tax_YN = 0, "No", "Yes")
'                .TextMatrix(I, Labour_Amt) = "0"
'                .TextMatrix(I, Misc_Chrg) = "0"
'                .TextMatrix(I, Spl_Chrg) = "0"
                .TextMatrix(I, TotQty) = Rs!Qty_Iss
'                .TextMatrix(I, StkQty) = "0"
'                .TextMatrix(I, FloatQty) = "0"
                .TextMatrix(I, Price) = Format(Rs!Rate, "0.00")
            End With
            Rs.MoveNext
            I = I + 1
        Loop
    End If
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT  S_No,Lab_Code,LabourAmt from Job_Lab where Job_DocID  = '" & RsJob1!DocID & "' and War_Lab_Rate <>0")
    If Rs.RecordCount > 0 Then
        Do Until Rs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, 0) = I
                .TextMatrix(I, JobCode) = Rs!Lab_Code
                .TextMatrix(I, Labour_Amt) = Rs!LabourAmt
'                .TextMatrix(I, Misc_Chrg) = "0"
'                .TextMatrix(I, Spl_Chrg) = "0"
'                .TextMatrix(I, TotQty) = "0"
'                .TextMatrix(I, StkQty) = "0"
'                .TextMatrix(I, FloatQty) = "0"
'                .TextMatrix(I, Price) = "0"
            End With
            Rs.MoveNext
            I = I + 1
        Loop
    End If
End If
If FGrid.Rows = 2 Then FGrid.AddItem FGrid.Rows - 1
FGrid.FixedRows = 2
Set Rs = Nothing
Set RsJob1 = Nothing
If txt(SellDlrCode) = "" Then
    MsgBox "Selling Dealer Code missing, Fill through Workshop Vehicle Master" & vbCrLf & "Terminating !", vbCritical, "Validation"
    FillData = True
End If
If txt(VehSaleDt) = "" Then
    MsgBox "Sale / Delivery Date missing, Fill through Workshop Vehicle Master" & vbCrLf & "Terminating !", vbCritical, "Validation"
    FillData = True
End If
End Function

Private Function ConvertDateLocal(Temp)
    If IsNull(Temp) Or Temp = "" Then
        ConvertDateLocal = "Null"
    Else
        ConvertDateLocal = "" & ConvertDate(RetDate(CDate(Temp))) & ""
    End If
End Function

Private Sub SetMaxLength()
Dim mMaxLength As Integer, mAlignment As Byte, mHeight As Integer
mHeight = FGrid.RowHeight(0)
'Alignment :  0-mLeft  1-mRright   2-mCenter
    Select Case FGrid.Col
        Case AggNo 'N,Common
            mMaxLength = 15
        Case FailureCode 'N,Common
            mMaxLength = 1
        Case CustComplCode 'C, Common
            mMaxLength = 3
        Case ComplFailCode 'C, Common
            mMaxLength = 7
        Case MakeCode 'C, Common
            mMaxLength = 6
        Case MakeCodeRepl   'C CVD only
            mMaxLength = 6
        Case JobCode    'N for Car/MUV, C for CVD
            mMaxLength = 6
        Case NoOfCompl  'For CAR/CVd No. of Complaints, for MUV No. of Job Codes
            mMaxLength = 1
        Case ComplDesc 'C Investigation for CVD
            mMaxLength = 480
            mHeight = FGrid.RowHeight(0) * 3
        
'If lblWarr = "CVD Warranty PCR" Then
'ElseIf lblWarr = "MUV Warranty PCR" Then
'ElseIf lblWarr = "CAR Warranty PCR" Then
        
        Case ComplDesc2, ActionTaken ' For CVD
            mMaxLength = 240
            mHeight = FGrid.RowHeight(0) * 3
        Case SourceCode ' For CVD
            mMaxLength = 1
        Case Price ' For CVD
            mMaxLength = 10 '8,2
            mAlignment = 1
        Case Labour_Amt 'Lab Chgs
            mMaxLength = 10
            mAlignment = 1
        Case Spl_Chrg    'Spl Lab Chgs
            mMaxLength = 10
            mAlignment = 1
        Case TotQty 'N, Common
            mMaxLength = IIf(lblWarr = "CAR Warranty PCR", 1, 3)
            mAlignment = 1
        Case StkQty 'N, CAR/MUV
            mMaxLength = IIf(lblWarr = "MUV Warranty PCR", 3, 1)
            mAlignment = 1
        Case FloatQty 'N, CAR/MUV
            mMaxLength = IIf(lblWarr = "MUV Warranty PCR", 3, 1)
            mAlignment = 1
        Case Misc_Chrg
            mMaxLength = 10
            mAlignment = 1
        Case SPD_InvNo, SPD_InvDt
            mMaxLength = 10
            mAlignment = 0
    End Select
    TxtGrid(0).MaxLength = mMaxLength
    TxtGrid(0).Alignment = mAlignment
End Sub

