VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmLabDesc 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Labour Description Master"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   14985
   Begin MSDataGridLib.DataGrid DGDep_Item 
      Height          =   4455
      Left            =   4755
      Negotiate       =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   7858
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
      RowHeight       =   20
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Deprecation Item"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "code"
         Caption         =   "code"
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
         Caption         =   "Deprecation Item Master"
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
      Height          =   210
      Index           =   13
      Left            =   6795
      MaxLength       =   25
      TabIndex        =   35
      Top             =   2370
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DGTrouble 
      Height          =   2730
      Left            =   6000
      Negotiate       =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4815
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
         Caption         =   "Trouble Name"
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
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4575.118
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtgrid1 
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
      Index           =   1
      Left            =   4845
      MaxLength       =   40
      TabIndex        =   27
      Top             =   4350
      Visible         =   0   'False
      Width           =   705
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
      Height          =   210
      Index           =   12
      Left            =   2805
      MaxLength       =   3
      TabIndex        =   11
      Top             =   2370
      Width           =   720
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
      Height          =   210
      Index           =   11
      Left            =   7290
      MaxLength       =   6
      TabIndex        =   10
      Top             =   2130
      Width           =   720
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
      Height          =   210
      Index           =   10
      Left            =   2805
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2130
      Width           =   720
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
      Height          =   210
      Index           =   9
      Left            =   7290
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1890
      Width           =   720
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
      Height          =   210
      Index           =   8
      Left            =   2805
      MaxLength       =   25
      TabIndex        =   7
      Top             =   1890
      Width           =   1815
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   8490
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   2505
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   75
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   30
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
            Name            =   "Verdana"
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
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   661
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
      Height          =   210
      Index           =   7
      Left            =   7065
      MaxLength       =   25
      TabIndex        =   20
      Top             =   1410
      Width           =   945
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
      Height          =   210
      Index           =   6
      Left            =   7065
      MaxLength       =   25
      TabIndex        =   19
      Top             =   1170
      Width           =   945
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
      Height          =   210
      Index           =   5
      Left            =   7290
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1650
      Width           =   720
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
      Height          =   210
      Index           =   4
      Left            =   2805
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1650
      Width           =   720
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
      Height          =   210
      Index           =   3
      Left            =   2805
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1410
      Width           =   4230
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
      Height          =   210
      Index           =   2
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1170
      Width           =   4230
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
      Height          =   210
      Index           =   1
      Left            =   2805
      MaxLength       =   40
      TabIndex        =   2
      Top             =   930
      Width           =   5205
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
      Height          =   210
      Index           =   0
      Left            =   2805
      MaxLength       =   6
      TabIndex        =   1
      Top             =   690
      Width           =   1065
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   1695
      Left            =   930
      TabIndex        =   28
      Top             =   3255
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   0
      FocusRect       =   0
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "ddd"
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
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSDataGridLib.DataGrid DGLabD 
      Height          =   3225
      Left            =   5460
      TabIndex        =   31
      Top             =   5880
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   5689
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "LAB_Code"
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
         DataField       =   "LAB_DESC"
         Caption         =   "Name"
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
         DataField       =   "LAB_Code"
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   0
            Locked          =   -1  'True
            ColumnWidth     =   3089.764
         EndProperty
         BeginProperty Column02 
            DividerStyle    =   0
            Locked          =   -1  'True
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGLabT 
      Height          =   3225
      Left            =   495
      TabIndex        =   32
      Top             =   6090
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   5689
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "LAB_TYPE"
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
         DataField       =   "LAB_DESC"
         Caption         =   "Name"
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
         DataField       =   ""
         Caption         =   ""
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   0
            Locked          =   -1  'True
            ColumnWidth     =   3435.024
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGLabG 
      Height          =   3225
      Left            =   6345
      TabIndex        =   33
      Top             =   6150
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   5689
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "LAB_GROUP"
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
         DataField       =   "LABGRP_DESC"
         Caption         =   "Name"
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
         DataField       =   ""
         Caption         =   ""
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            DividerStyle    =   0
            Locked          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   0
            Locked          =   -1  'True
            ColumnWidth     =   3435.024
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deprecation Iteml...................."
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
      Left            =   4890
      TabIndex        =   34
      Top             =   2400
      Width           =   2745
   End
   Begin VB.Label LblGrid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Applicable Against Complaints"
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
      Left            =   945
      TabIndex        =   29
      Top             =   3030
      Width           =   2580
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Based (Y/N)*"
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
      Left            =   945
      TabIndex        =   26
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warranty Hrs."
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
      Left            =   5535
      TabIndex        =   25
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chargeable Rate"
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
      Left            =   945
      TabIndex        =   24
      Top             =   2160
      Width           =   1440
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chargeable Hrs."
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
      Left            =   5550
      TabIndex        =   23
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chargeable From*"
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
      Left            =   945
      TabIndex        =   18
      Top             =   1920
      Width           =   1590
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Major (Y/N)"
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
      Left            =   5535
      TabIndex        =   17
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "External Job (Y/N)"
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
      Left            =   945
      TabIndex        =   16
      Top             =   1680
      Width           =   1560
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour Description*"
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
      Left            =   945
      TabIndex        =   15
      Top             =   960
      Width           =   1710
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour Type*"
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
      Left            =   945
      TabIndex        =   14
      Top             =   1200
      Width           =   1170
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour Group*"
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
      Left            =   945
      TabIndex        =   13
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour Code*"
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
      Left            =   945
      TabIndex        =   12
      Top             =   720
      Width           =   1200
   End
End
Attribute VB_Name = "frmLabDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstLabT As ADODB.Recordset, RstLabG As ADODB.Recordset
Dim mFlag As Byte
Dim GridKey As Integer
Dim RsTrb As ADODB.Recordset

Private Const CCCode As Byte = 1
Private Const CCDesc As Byte = 2
Dim RstDEp_item As ADODB.Recordset
Private Const Lab_Code = 0, Lab_Desc = 1, Lab_Type = 2, Lab_Group = 3, External = 4, Major = 5
Private Const Lab_Tcode = 6, Lab_Gcode = 7, Ch_From = 8, Ch_Hrs = 9, Ch_Rate = 10, Wr_Hrs = 11, Dep_Item As Byte = 13
Private Const ModelBased = 12
Dim ListArray As Variant
Dim mListItem As ListItem



Private Sub DGLabT_Click()
    DGLabT_KeyDown vbKeyReturn, 0
End Sub

Private Sub DGLabG_Click()
    DGLabG_KeyDown vbKeyReturn, 0
End Sub
Private Sub DGDep_Item_click()
    txt(Dep_Item) = RstDEp_item!Name
    txt(Dep_Item).Tag = RstDEp_item!Code
    txt(Dep_Item).SetFocus
    DGDep_Item.Visible = False
End Sub
Private Sub DGLabG_DblClick()
    txt(Lab_Group).TEXT = RstLabG!LabGrp_Desc
    txt(Lab_Gcode).TEXT = RstLabG!Lab_Code
    txt(Lab_Group).Tag = RstLabG!Lab_Code
    DGLabG_KeyDown 13, 0
End Sub

Private Sub DGLabT_DblClick()
    txt(Lab_Type).TEXT = RstLabG!Lab_Desc
    txt(Lab_Tcode).TEXT = RstLabG!Lab_Type
    txt(Lab_Type).Tag = RstLabG!Lab_Type
    DGLabT_KeyDown 13, 0
End Sub

Private Sub DGLabT_KeyDown(KeyCode As Integer, Shift As Integer)
If RstLabT.BOF = True Or RstLabT.EOF = True Then Exit Sub
If KeyCode = vbKeyEscape Then
    txt(Lab_Type).TEXT = ""
Else
    txt(Lab_Type).TEXT = RstLabT!Lab_Desc
    If KeyCode = vbKeyReturn Then
        If RstLabT.RecordCount > 0 Then
            txt(Lab_Type).SetFocus
        End If
    End If
End If

End Sub

Private Sub DGLabG_KeyDown(KeyCode As Integer, Shift As Integer)
If RstLabG.BOF = True Or RstLabG.EOF = True Then Exit Sub
If KeyCode = vbKeyEscape Then
    txt(Lab_Group).TEXT = ""
Else
    txt(Lab_Group).TEXT = RstLabG!LabGrp_Desc
    If KeyCode = vbKeyReturn Then
        If RstLabG.RecordCount > 0 Then
            txt(Lab_Group).SetFocus
        End If
    End If
End If

End Sub

Private Sub FGrid1_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
txtgrid1(1).Visible = False
End Sub

Private Sub FGrid1_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid1.Col
    Case CCCode, CCDesc
        Call GridDblClick(Me, FGrid1, txtgrid1, 1)
End Select

End Sub

Private Sub FGrid1_EnterCell()
'FGrid1.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid1_GotFocus()
    FGrid1.BackColorSel = FaBackColorSelEnter

    FGrid1.Col = CCCode
    txtgrid1(1).Visible = False
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
    If MsgBox("Do You Want to Save?", vbYesNo) = vbYes Then TopCtrl1_eSave
'    SendKeysA vbKeyTab, True
'    KeyCode = 0
End If
GridKey = KeyCode
FGrid1.Tag = FGrid1.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid1.Col
        Case CCCode, CCDesc
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
    End Select
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid1.Col
        Case CCCode, CCDesc
            Call GridDblClick(Me, FGrid1, txtgrid1, 1)
            
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
Select Case FGrid1.Col
    Case CCCode
        RsTrb.Sort = "Code"
       Call Get_Text(Me, FGrid1, txtgrid1, 1, False, KeyAscii)
    Case CCDesc
        RsTrb.Sort = "Name"
        Call Get_Text(Me, FGrid1, txtgrid1, 1, False, KeyAscii)
End Select

End Sub

Private Sub FGrid1_LostFocus()
FGrid1.BackColorSel = FaCellBackColLeave1

FGrid1_Validate (True)
End Sub

Private Sub FGrid1_Scroll()
txtgrid1(1).Visible = False
DGTrouble.Visible = False
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid1.ColSel = False Then Exit Sub
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
Exit Sub
End Sub

Private Sub FGrid1_Validate(Cancel As Boolean)
'    FGrid1.CellBackColor = CellBackColLeave
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
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
Me.top = 0: Me.left = 0
Me.width = 9345: Me.height = 6870
TopCtrl1.Tag = PubUParam

Set RstMain = New ADODB.Recordset
If PubMoveRecYn Then
    RstMain.Open "Select Labour.Lab_Code as SearchCode,Labour.*,Ditm.Description as DepItemname  From Labour left join Deprecation_itemMaster ditm on Labour.Dep_Item=ditm.code  Order by Lab_Desc", GCn, adOpenDynamic, adLockOptimistic
Else
    Set RstMain = GCn.Execute("Select Top 1 Labour.Lab_Code as SearchCode,Labour.*,Ditm.Description as DepItemname  From Labour left join Deprecation_itemMaster ditm on Labour.Dep_Item=ditm.code  Order by Lab_Desc")
End If

Set RstHelp = New ADODB.Recordset
RstHelp.Open "Select Lab_Code,Lab_Desc FROM Labour Order by Lab_Desc", GCn, adOpenDynamic, adLockOptimistic

Set RstLabT = New ADODB.Recordset
RstLabT.Open "Select Lab_Type,Lab_Desc FROM Labour_type Order by Lab_Desc", GCn, adOpenDynamic, adLockOptimistic

Set RstLabG = New ADODB.Recordset
RstLabG.Open "Select Lab_Group,LabGrp_Desc FROM Labour_Group Order by LabGrp_Desc", GCn, adOpenDynamic, adLockOptimistic

Set RsTrb = GCn.Execute("Select Trouble_code as code,Trouble_Name as name FROM trouble Order by trouble_Code")
Set DGTrouble.DataSource = RsTrb



'Nikhil
Set RstDEp_item = New ADODB.Recordset
RstDEp_item.CursorLocation = adUseClient
RstDEp_item.Open "Select CODE as code ,Description as name From Deprecation_itemMaster ", GCn, adOpenDynamic, adLockOptimistic
Set DGDep_Item.DataSource = RstDEp_item



ListArray = Array("Customer", "Manufacturer", "Self", "Other Dealer")
Set mListItem = ListView_Items(ListView, txt, Ch_From, ListArray, 4)
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0:    mFlag = 0
Set DGLabD.DataSource = RstHelp
Set DGLabT.DataSource = RstLabT
Set DGLabG.DataSource = RstLabG

DGDep_Item.Visible = False

Ini_Grid
Grid_Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing: Set RstLabT = Nothing: Set RstLabG = Nothing: Set RstDEp_item = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrLoop
BlankText
Disp_Text SETS("ADD", Me, RstMain)
'Txt(Lab_Code).Tag = Txt(Lab_Code)
Txt_GotFocus Lab_Code
ADDFLAG = 1
txt(Lab_Tcode).Enabled = False
txt(Lab_Gcode).Enabled = False
txt(External) = "No"
txt(Major) = "No"
txt(ModelBased) = "Yes"
txt(Ch_From) = "Customer"
'Chrg_From.ListIndex = 0
txt(Lab_Code).SetFocus


FGrid1.Rows = 1
FGrid1.AddItem ""
FGrid1.FixedRows = 1

Exit Sub
ErrLoop:    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo ErrLoop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    txt(Lab_Code).Enabled = False
    txt(Lab_Tcode).Enabled = False
    txt(Lab_Gcode).Enabled = False
    txt(Lab_Desc).Tag = txt(Lab_Desc)
    Txt_GotFocus Lab_Desc
    ADDFLAG = 2
    txt(Lab_Desc).SetFocus
    
    
    FGrid1.AddItem ""
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
ErrLoop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
On Error GoTo ErrLoop
Dim transFalg As Byte
transFalg = 0
Dim XBM
Dim Res As Integer
    If RstMain.RecordCount > 0 Then
        If GCn.Execute("Select Lab_Code from Labour_Model where Lab_Code='" & txt(Lab_Code).Tag & "'").RecordCount > 0 Then
            MsgBox "Transaction in Model-wise Labour exists" & vbCrLf & "Delete Denied !", vbCritical, "Delete Denied!"
            Exit Sub
        End If
        Res = MsgBox("Do You Want to Delete Record ", 4 + vbQuestion, "Confirmation ")
        If Res = 6 Then
            GCn.BeginTrans
            XBM = RstMain.Bookmark
            transFalg = 1
            GCn.Execute ("delete  from Labour where Lab_CODE= '" & txt(Lab_Code).Tag & "'")
            GCn.CommitTrans
            transFalg = 0
            RstMain.Requery
            RstHelp.Requery
            If RstMain.RecordCount >= XBM Then
                RstMain.Bookmark = XBM
            Else
                If RstMain.EOF = False Then RstMain.MoveLast
            End If
            Call MoveRec
        End If
    Else
        MsgBox "No Records To Delete.", vbInformation, "Information"
    End If

Exit Sub
ErrLoop:    If transFalg = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbExclamation, " Deletion Error "
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, RstMain, 1
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, RstMain, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, RstMain, 3
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, RstMain, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "Select Lab_Code as SearchCode, Lab_Code,Left(Lab_Desc,40) as Description from Labour Order By Lab_Desc"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SearchCode='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select Labour.Lab_Code as SearchCode,Labour.* From Labour Where Labour.Lab_Code= '" & MyValue & "' Order by Lab_Desc")
    End If
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim I As Integer, mQry$, mRepName$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER

    mRepName = "LabourDesc"
    If PubBackEnd = "A" Then
        mQry = "Select L.*, switch(l.Chrg_From='C','Cutomer',l.Chrg_From='O','Other Dealer',l.Chrg_From='M','Manufacturer',l.Chrg_From='S','Self') as ChrgFrm," & _
            " " & cIIF("L.External_YN=1", "'Yes'", "'No'") & " as ExternalYN," & _
            " " & cIIF("L.Major_YN=1", "'Yes'", "'No'") & " as  MajorYN, " & _
            " LG.LabGrp_Desc, LT.Lab_Desc,switch(L.ModelBased=0,'No',L.ModelBased=1,'Yes') as  ModelBaseYN " & _
            " from (Labour L left join Labour_Group LG on L.Lab_Group=LG.Lab_Group) " & _
            " Left join Labour_Type LT on L.Lab_Type=LT.Lab_Type " & _
            " Order by L.Lab_Group,L.Lab_Type,L.Lab_Code"
    ElseIf PubBackEnd = "S" Then
        mQry = "Select L.*, Case l.Chrg_From When 'C' Then 'Cutomer' When 'O' Then 'Other Dealer' When 'M' Then 'Manufacturer' When 'S' Then 'Self' End as ChrgFrm," & _
            " " & cIIF("L.External_YN=1", "'Yes'", "'No'") & " as ExternalYN," & _
            " " & cIIF("L.Major_YN=1", "'Yes'", "'No'") & " as  MajorYN, " & _
            " LG.LabGrp_Desc, LT.Lab_Desc, " & cIIF("L.ModelBased=1", "'Yes'", "'No'") & " as  ModelBaseYN " & _
            " from (Labour L left join Labour_Group LG on L.Lab_Group=LG.Lab_Group) " & _
            " Left join Labour_Type LT on L.Lab_Type=LT.Lab_Type " & _
            " Order by L.Lab_Group,L.Lab_Type,L.Lab_Code"
    End If
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION, , False)
    Set Rst = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub TopCtrl1_eSave()
Dim mTrans As Byte
Dim I As Integer
Dim mSearchCode$
On Error GoTo ErrLoop
    mTrans = 0
    If IsValid(txt(Lab_Code), "Labour Code") = False Then Txt_GotFocus Lab_Code: Exit Sub
    If IsValid(txt(Lab_Desc), "Labour Description") = False Then Txt_GotFocus Lab_Desc: Exit Sub
    If IsValid(txt(Lab_Type), "Labour Type") = False Then Txt_GotFocus Lab_Type: Exit Sub
    If IsValid(txt(Lab_Group), "Labour Group") = False Then Txt_GotFocus Lab_Group: Exit Sub
    If IsValid(txt(Ch_From), "Chargable From") = False Then Txt_GotFocus Ch_From: Exit Sub
    If IsValid(txt(ModelBased), "Model Based") = False Then Txt_GotFocus ModelBased: Exit Sub
    
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From Labour Where Lab_Code=" & Chk_Text(Trim(txt(Lab_Code)))).Fields(0) > 0 Then MsgBox "Labour Code Already Exists", vbInformation, "Duplicate Checking": Txt_GotFocus Lab_Code: txt(Lab_Code).SetFocus: Exit Sub
    
    
    
    GCn.BeginTrans
    mTrans = 1
        
        GCn.Execute "Delete From Lab_Trouble Where Lab_Code='" & txt(Lab_Code) & "'"
        
        If ADDFLAG = 1 Then
            GCn.Execute ("Insert Into Labour(lab_Code,Site_Code,Div_Code,Lab_Desc,External_YN,Major_YN,Lab_Type,Lab_Group,Chrg_From,time_req,lab_rate,wtime_req,U_Name,U_EntDt,U_AE,ModelBased,Dep_Item) Values('" & txt(Lab_Code) & "','" & PubSiteCode & "','" & PubDivCode & "','" & txt(Lab_Desc) & "'," & IIf(txt(External) = "Yes", 1, 0) & "," & IIf(txt(Major) = "Yes", 1, 0) & ",'" & txt(Lab_Tcode) & "','" & txt(Lab_Gcode) & "','" & left(txt(Ch_From), 1) & "'," & Val(txt(Ch_Hrs)) & "," & Val(txt(Ch_Rate)) & "," & Val(txt(Wr_Hrs)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "', " & IIf(txt(ModelBased) = "Yes", 1, 0) & ",'" & txt(Dep_Item).Tag & "' )")
            mSearchCode = txt(Lab_Code)
        ElseIf ADDFLAG = 2 Then
            GCn.Execute ("UPDATE Labour SET Lab_Desc='" & txt(Lab_Desc) & "',External_YN=" & IIf(txt(External) = "Yes", 1, 0) & ",Major_YN=" & IIf(txt(Major) = "Yes", 1, 0) & ",Lab_Type='" & txt(Lab_Tcode) & "',Lab_Group='" & txt(Lab_Gcode) & "',Chrg_From='" & left(txt(Ch_From), 1) & "',time_req=" & Val(txt(Ch_Hrs)) & ",lab_rate=" & Val(txt(Ch_Rate)) & ",wtime_req=" & Val(txt(Wr_Hrs)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & left(TopCtrl1.TopText2, 1) & "', ModelBased=" & IIf(txt(ModelBased) = "Yes", 1, 0) & " ,Dep_Item='" & txt(Dep_Item).Tag & "' Where Lab_Code='" & RstMain!Lab_Code & "'")
            mSearchCode = RstMain!Lab_Code
        End If
    
    
        For I = 1 To FGrid1.Rows - 1
            If FGrid1.TextMatrix(I, CCCode) <> "" Then
                GCn.Execute "Insert Into Lab_Trouble (Srl, CCCode, Lab_Code) Values (" & I & ", '" & FGrid1.TextMatrix(I, CCCode) & "', '" & txt(Lab_Code) & "')"
            End If
        Next I
    
    GCn.CommitTrans
    mTrans = 0
    
    
    
    If MasterFormExit Then Unload Me: Exit Sub
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select Labour.Lab_Code as SearchCode,Labour.* From Labour Where Labour.Lab_Code= '" & mSearchCode & "' Order by Lab_Desc")
    End If
    RstHelp.Requery
    RstMain.FIND ("Lab_Code='" & txt(Lab_Code).Tag & "'")
    If ADDFLAG = 1 Then
        BlankText
        Txt_GotFocus Lab_Code
        txt(Lab_Code).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        ADDFLAG = 0
        
        Grid_Hide
    End If
    ADDFLAG = 0
Exit Sub
ErrLoop:    If mTrans = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ErrLoop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, RstMain)
        Me.ActiveControl.SetFocus
        MoveRec
        DGLabT.Visible = False
        DGLabG.Visible = False
        DGLabD.Visible = False
    End If
Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub MoveRec()
Dim Rs As Recordset
Dim I As Integer
On Error GoTo ErrLoop
RST_BOF_EOF RstMain
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT LT.LAB_DESC as LTDesc,LG.LABGRP_DESC as LGDesc " & _
        " FROM LABOUR_TYPE LT,LABOUR_GROUP LG " & _
        " WHERE LT.LAB_TYPE='" & XNull(RstMain!Lab_Type) & "' AND LG.LAB_GROUP='" & XNull(RstMain!Lab_Group) & "'")
    'Txt(Lab_Code) = mID(XNull(RstMain!Lab_Code), 2, 5)
    txt(Lab_Code) = XNull(RstMain!Lab_Code)
    txt(Lab_Code).Tag = XNull(RstMain!Lab_Code)
    txt(Lab_Desc) = XNull(RstMain!Lab_Desc)
    If Rs.RecordCount > 0 Then
        txt(Lab_Type) = XNull(Rs!LTDesc)
        txt(Lab_Group) = XNull(Rs!LGDesc)
    Else
        txt(Lab_Type) = ""
        txt(Lab_Group) = ""
    End If
    
    'Nikhil
    txt(Dep_Item).Tag = IIf(IsNull(RstMain!Dep_Item), "", RstMain!Dep_Item)
    txt(Dep_Item) = IIf(IsNull(RstMain!DepItemname), "", RstMain!DepItemname)


    txt(Lab_Tcode) = XNull(RstMain!Lab_Type)
    txt(Lab_Gcode) = XNull(RstMain!Lab_Group)
    txt(External) = IIf(XNull(RstMain!External_yn) = "1", "Yes", "No")
    txt(Major) = IIf(XNull(RstMain!Major_YN) = "1", "Yes", "No")
'    Chrg_From.ListIndex = RstMain!Chrg_From
    Select Case XNull(RstMain!Chrg_From)
        Case "C"
            txt(Ch_From) = "Customer"
        Case "M"
            txt(Ch_From) = "Manufacturer"
        Case "S"
            txt(Ch_From) = "Self"
        Case "O"
            txt(Ch_From) = "Other Dealer"
        Case Else
            txt(Ch_From) = ""
    End Select
    txt(Ch_Hrs) = Format(IIf(IsNull(RstMain!TIME_REQ), 0, RstMain!TIME_REQ), "0.00")
    txt(Ch_Rate) = Format(IIf(IsNull(RstMain!Lab_Rate), 0, RstMain!Lab_Rate), "0.00")
    txt(Wr_Hrs) = Format(IIf(IsNull(RstMain!WTime_Req), 0, RstMain!WTime_Req), "0.00")
    txt(ModelBased) = IIf(XNull(RstMain!ModelBased) = "1", "Yes", "No")
    
    
    Set Rs = GCn.Execute("Select L.Srl, L.CCCode, T.Trouble_Name   From Lab_Trouble L Left Join Trouble T On L.CCCode=T.Trouble_Code Where Lab_Code='" & RstMain!Lab_Code & "' Order By L.Srl")
    FGrid1.Rows = 1
    If Rs.RecordCount > 0 Then
        
        I = 1
        Do Until Rs.EOF
            FGrid1.AddItem ""
            
            
            FGrid1.TextMatrix(I, 0) = Rs!Srl
            FGrid1.TextMatrix(I, CCCode) = Rs!CCCode
            FGrid1.TextMatrix(I, CCDesc) = Rs!trouble_name
            
            I = I + 1
            Rs.MoveNext
        Loop
        FGrid1.FixedRows = 1
    Else
        FGrid1.AddItem ""
        FGrid1.FixedRows = 1
    End If
Grid_Hide
    
End If
Exit Sub
ErrLoop:        MsgBox err.Description
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
    RstLabT.Requery
    RstLabG.Requery
     RstDEp_item.Requery
End Sub
Private Sub TopCtrl1_eExit()
    RstMain.Cancel
    Unload Me
End Sub

Private Sub LabCodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "Lab_CODE  >=" & Chk_Text(XNull(Trim(txt(Lab_Code))))
End Sub

Private Sub LabNameSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "Lab_Desc  >=" & Chk_Text(XNull(txt(Lab_Desc)))
End Sub

Private Sub LabTSearch()
If RstLabT.RecordCount <= 0 Then Exit Sub
RstLabT.MoveFirst
RstLabT.FIND "Lab_Desc  >=" & Chk_Text(XNull(txt(Lab_Type)))
If Not RstLabT.EOF Then
    If RstLabT!Lab_Desc <> Trim(XNull(txt(Lab_Type))) Then
        LabTExSearch
    End If
Else
    LabTExSearch
End If
End Sub


Private Sub DEPSearch()
If RstDEp_item.RecordCount <= 0 Then Exit Sub
RstDEp_item.MoveFirst
RstDEp_item.FIND "NAme  >=" & Chk_Text(XNull(txt(Dep_Item)))
If Not RstDEp_item.EOF Then
    If RstDEp_item!Name <> Trim(XNull(txt(Dep_Item))) Then
        DepTExSearch
    End If
Else
    DepTExSearch
End If
End Sub

Private Sub DepTExSearch()
Dim tempRst As ADODB.Recordset
Set tempRst = RstDEp_item.Clone
tempRst.Sort = "Name ASC"
tempRst.FIND "NAme  >='" & FilterString(XNull(txt(Dep_Item))) & "'"
If Not tempRst.EOF Then
    RstDEp_item.MoveFirst
    RstDEp_item.FIND "name  >=" & Chk_Text(XNull(tempRst!Name))
End If
Set tempRst = Nothing
End Sub

Private Sub LabTExSearch()
Dim tempRst As ADODB.Recordset
Set tempRst = RstLabT.Clone
tempRst.Sort = "Lab_Desc ASC"
tempRst.FIND "Lab_Desc  >='" & FilterString(XNull(txt(Lab_Type))) & "'"
If Not tempRst.EOF Then
    RstLabT.MoveFirst
    RstLabT.FIND "Lab_Desc  >=" & Chk_Text(XNull(tempRst!Lab_Desc))
End If
Set tempRst = Nothing
End Sub
Private Sub LabGSearch()
If RstLabG.RecordCount <= 0 Then Exit Sub
RstLabG.MoveFirst
RstLabG.FIND "LabGrp_Desc   >=" & Chk_Text(XNull(txt(Lab_Group)))
If Not RstLabG.EOF Then
    If RstLabG!LabGrp_Desc <> Trim(XNull(txt(Lab_Group))) Then
        LabGExSearch
    End If
Else
    LabGExSearch
End If
End Sub
Private Sub LabGExSearch()
Dim tempRst As ADODB.Recordset
Set tempRst = RstLabG.Clone
tempRst.Sort = "LabGrp_Desc ASC"
tempRst.FIND "LabGrp_Desc   >='" & FilterString(XNull(txt(Lab_Group))) & "'"
If Not tempRst.EOF Then
    RstLabG.MoveFirst
    RstLabG.FIND "LabGrp_Desc  >=" & Chk_Text(XNull(tempRst!LabGrp_Desc))
End If
Set tempRst = Nothing
End Sub

Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case Lab_Code, Lab_Desc
            If RstHelp.RecordCount = 0 Then Exit Sub
            DGLabD.top = txt(Index).top + txt(Index).height + 10
            DGLabD.left = txt(Index).left
            DGLabD.ZOrder 0
        Case Lab_Type
            DGLabT.top = txt(Index).top + txt(Index).height + 10
            DGLabT.left = txt(Index).left
            DGLabT.ZOrder 0
        Case Dep_Item
         DGDep_Item.top = txt(Index).top + txt(Index).height + 10
            DGDep_Item.left = txt(Index).left
            DGDep_Item.ZOrder 0
        Case Lab_Group
            DGLabG.top = txt(Index).top + txt(Index).height + 10
            DGLabG.left = txt(Index).left
            DGLabG.ZOrder 0
    End Select
End If
End Sub

Private Sub Txt_GotFocus(Index As Integer)
DGLabD.Columns(0).width = 1000.1: DGLabD.Columns(1).width = 3535.024: DGLabD.Columns(2).width = 1000.1
Dim mBookMark
    Ctrl_GetFocus txt(Index)
    mFlag = 0
    If DGLabD.Visible = True Then DGLabD.Visible = False
    If DGLabT.Visible = True Then DGLabT.Visible = False
    If DGLabG.Visible = True Then DGLabG.Visible = False
    RST_BOF_EOF RstHelp
    txt(Index).Tag = txt(Index)
    Select Case Index
        Case Lab_Code, Lab_Desc
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
        Case Lab_Type
            If RstLabT.BOF Or RstLabT.EOF Then Exit Sub
          Case Dep_Item
            If RstDEp_item.BOF Or RstDEp_item.EOF Then Exit Sub
        Case Lab_Group
            If RstLabG.BOF Or RstLabG.EOF Then Exit Sub
    End Select
    Select Case Index
        Case Lab_Code
            DGLabD.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "lab_CODE ASC"
            RstHelp.Bookmark = mBookMark
            LabCodeSearch
        Case Lab_Desc
            DGLabD.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "lab_desc ASC"
            RstHelp.Bookmark = mBookMark
            LabNameSearch
        Case Lab_Type
            DGLabT.Columns(0).width = 0: DGLabT.Columns(2).width = 0
            mBookMark = RstLabT.Bookmark
            RstLabT.Sort = "lab_desc ASC"
            RstLabT.Bookmark = mBookMark
            LabTSearch
            
        Case Dep_Item
            DGDep_Item.Columns(0).width = 0: DGDep_Item.Columns(0).width = 0
            mBookMark = RstDEp_item.Bookmark
            RstDEp_item.Sort = "name ASC"
            RstDEp_item.Bookmark = mBookMark
           DEPSearch
            
        
        Case Lab_Group
            DGLabG.Columns(0).width = 0: DGLabG.Columns(2).width = 0
            mBookMark = RstLabG.Bookmark
            RstLabG.Sort = "labgrp_desc ASC"
            RstLabG.Bookmark = mBookMark
            LabGSearch
    End Select
    If txt(Index) = "" Then Txt_Change Index
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer
Select Case Index
    Case Ch_From
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 1200
        If KeyCode = 13 Or KeyCode = vbKeyTab Then
           FrmList.Visible = False
        End If
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown And FrmList.Visible = False Then
            'If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
                SendKeysA vbKeyTab, True
                KeyCode = 0
            'Else
            '    txt(Ch_From).SetFocus
            'End If
        ElseIf KeyCode = vbKeyUp Then
            If FrmList.Visible = False Then SendKeys "+{Tab}": KeyCode = 0
        End If
        'Nikhil
    Case Dep_Item
        DGridTxtKeyDown DGDep_Item, txt, Index, RstDEp_item, KeyCode, False, 1, FrmDeprecation_itemMaster, "FrmDeprecation_itemMaster"
        
        
    Case Lab_Code, Lab_Desc
        DGLabD.Visible = True
    Case Lab_Type
        DGLabT.Visible = True
        If DGLabT.Visible = True Then
            Select Case KeyCode
                Case vbKeyUp
                    If Not RstLabT.BOF Then RstLabT.MovePrevious
                Case vbKeyDown
                    If Not RstLabT.EOF Then RstLabT.MoveNext
                Case 33
                    For I = 1 To 9
                        If Not RstLabT.BOF Then RstLabT.MovePrevious
                    Next
                Case 34
                    For I = 1 To 9
                        If Not RstLabT.EOF Then RstLabT.MoveNext
                    Next
                Case 13
                    SendKeysA vbKeyTab, True
            End Select
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, 33, 34
                    RST_BOF_EOF RstLabT
                    If Not RstLabT.BOF And Not RstLabT.EOF Then
                        txt(Lab_Type) = XNull(RstLabT!Lab_Desc)
                        txt(Lab_Tcode) = XNull(RstLabT!Lab_Type)
                        txt(Lab_Type).SelStart = 0
                    End If
            End Select
        End If
    Case Lab_Group
        DGLabG.Visible = True
        If DGLabG.Visible = True Then
            Select Case KeyCode
                Case vbKeyUp
                    If Not RstLabG.BOF Then RstLabG.MovePrevious
                Case vbKeyDown
                    If Not RstLabG.EOF Then RstLabG.MoveNext
                Case 33
                    For I = 1 To 9
                        If Not RstLabG.BOF Then RstLabG.MovePrevious
                    Next
                Case 34
                    For I = 1 To 9
                        If Not RstLabG.EOF Then RstLabG.MoveNext
                    Next
                Case 13
                    SendKeysA vbKeyTab, True
            End Select
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, 33, 34
                    RST_BOF_EOF RstLabG
                    If Not RstLabT.BOF And Not RstLabG.EOF Then
                        txt(Lab_Group) = XNull(RstLabG!LabGrp_Desc)
                        txt(Lab_Gcode) = XNull(RstLabG!Lab_Group)
                        txt(Lab_Group).SelStart = 0
                    End If
            End Select
        End If
End Select
Select Case Index
    Case Lab_Type
        If DGLabT.Visible = False Then
            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
                SendKeysA vbKeyTab, True
                KeyCode = 0
            ElseIf KeyCode = vbKeyUp Then
                SendKeys "+{Tab}"
                KeyCode = 0
            End If
        End If
    Case Lab_Group
        If DGLabG.Visible = False Then
            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
                SendKeysA vbKeyTab, True
                KeyCode = 0
            ElseIf KeyCode = vbKeyUp Then
                SendKeys "+{Tab}"
                KeyCode = 0
            End If
        End If
End Select

Select Case Index
    Case External, Major, Ch_Hrs, Ch_Rate, Wr_Hrs
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case Lab_Code
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case Lab_Desc
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case ModelBased
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
        If UCase(left(PubComp_Name, 3)) <> "LMP" Then
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then: TopCtrl1_eSave
        End If
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If

End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case Index
    Case Wr_Hrs, Ch_Hrs
        NumPress txt(Index), KeyAscii, 3, 2
    Case Ch_Rate
        NumPress txt(Index), KeyAscii, 5, 2
End Select
'If Index = Pin Then NumPress Txt(Pin), KeyAscii, 6, 0
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
mFlag = 0
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case External, Major, ModelBased
        If Len(txt(Index)) = 0 Or UCase(Chr(KeyCode)) = "N" Then
            txt(Index) = "No"
        ElseIf UCase(Chr(KeyCode)) = "Y" Then
            txt(Index) = "Yes"
        Else
            txt(Index) = "No"
        End If
End Select
Select Case Index

 'Nikhil
      Case Dep_Item
         DEPSearch
        
        
    Case Ch_From
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case Lab_Code
        LabCodeSearch
    Case Lab_Desc
        LabNameSearch
    Case Lab_Type
        LabTSearch
    Case Lab_Group
        LabGSearch
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case Lab_Code
            Set Rst = GCn.Execute("SELECT Lab_Code FROM Labour WHERE lab_CODE=" & Chk_Text(txt(Lab_Code)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Labour Code Already Exists", vbInformation, "Validation": Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Lab_Code <> RstMain!Lab_Code Then MsgBox "Labour Code Already Exists", vbInformation, "Validation": Cancel = True: Exit Sub
                End If
            End If
            txt(Lab_Code).Tag = PubSiteCode & txt(Lab_Code)
            
        Case Lab_Desc
            Set Rst = GCn.Execute("SELECT Lab_Desc FROM Labour WHERE Lab_Desc=" & Chk_Text(txt(Lab_Desc)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Labour Description Already Exists", vbInformation, "Validation": txt(Lab_Desc) = txt(Lab_Desc).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Lab_Desc <> RstMain!Lab_Desc Then MsgBox "Labour Description Already Exists", vbInformation, "Validation": txt(Lab_Desc) = txt(Lab_Desc).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case Lab_Type
            If Not RstLabT.EOF And Not RstLabT.BOF Then
                txt(Lab_Type) = XNull(RstLabT!Lab_Desc)
                txt(Lab_Tcode) = XNull(RstLabT!Lab_Type)
            Else
                txt(Lab_Type) = "": txt(Lab_Type) = "": txt(Lab_Tcode) = ""
            End If
        Case Lab_Group
            If Not RstLabG.EOF And Not RstLabG.BOF Then
                txt(Lab_Group) = XNull(RstLabG!LabGrp_Desc)
                txt(Lab_Gcode) = XNull(RstLabG!Lab_Group)
            Else
                txt(Lab_Group) = "": txt(Lab_Group) = "": txt(Lab_Gcode) = ""
            End If
        Case Ch_Hrs, Ch_Rate, Wr_Hrs
            txt(Index) = Format(txt(Index), "0.00")
         'Nikhil
    Case Dep_Item
        txt(Dep_Item) = RstDEp_item!Name
        txt(Dep_Item).Tag = RstDEp_item!Code
        
    End Select
Set Rst = Nothing
End Sub

Private Sub DGLabT_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mFlag = 1 Then
    txt(Lab_Type) = DGLabT.Columns(1).TEXT
    txt(Lab_Tcode) = DGLabT.Columns(0).TEXT
End If
End Sub

Private Sub DGLabT_GotFocus()
    mFlag = 1
End Sub
Private Sub DGLabG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mFlag = 1 Then
    txt(Lab_Group) = DGLabG.Columns(1).TEXT
    txt(Lab_Gcode) = DGLabG.Columns(0).TEXT
End If
End Sub

Private Sub DGLabG_GotFocus()
    mFlag = 1
End Sub

Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
    txt(Lab_Code).Tag = ""
    txt(Lab_Desc).Tag = ""
Next I
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
Next

If UCase(left(PubComp_Name, 3)) = "LMP" Then
    FGrid1.Visible = True
    LblGrid.Visible = True
Else
    FGrid1.Visible = False
    LblGrid.Visible = False
End If
'Chrg_From.Enabled = Enb
End Sub

'Private Sub Ini_Grid()
'    FGrid.RowHeightMin = 250
'    FGrid.ColWidth(25) = 0
'End Sub

Sub Grid_Hide()
    DGLabD.Visible = False
    DGLabG.Visible = False
    DGLabT.Visible = False
    DGTrouble.Visible = False
        If DGDep_Item.Visible = True Then DGDep_Item.Visible = False
        
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
Ctrl_GetFocus txtgrid1(Index)
    Grid_Hide
    txtgrid1(1).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    
    Select Case FGrid1.Col
        Case CCCode
            DGTrouble.Move FGrid1.left, txtgrid1(1).top + txtgrid1(1).height + 20
            If RsTrb.RecordCount = 0 Or FGrid1.TextMatrix(FGrid1.Row, CCCode) = "" Then Exit Sub
            RsTrb.Sort = "Code"
            RsTrb.MoveFirst
            RsTrb.FIND "Code ='" & FGrid1.TextMatrix(FGrid1.Row, CCCode) & "'"
            If RsTrb.EOF = True Then RsTrb.MoveFirst
        Case CCDesc
            DGTrouble.Move FGrid1.left, txtgrid1(1).top + txtgrid1(1).height + 20
            If RsTrb.RecordCount = 0 Or FGrid1.TextMatrix(FGrid1.Row, CCDesc) = "" Then Exit Sub
            RsTrb.Sort = "name"
            RsTrb.MoveFirst
            RsTrb.FIND "name ='" & FGrid1.TextMatrix(FGrid1.Row, CCDesc) & "'"
            If RsTrb.EOF = True Then RsTrb.MoveFirst
    End Select
End Sub



Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtgrid1(1).TEXT = txtgrid1(1).Tag
        TxtGrid1_KeyUp Index, KeyCode, Shift
        FGrid1.SetFocus
        txtgrid1(1).Visible = False
        Exit Sub
    End If
    Select Case FGrid1.Col
        Case CCCode
            If DGTrouble.Visible = False Then DGridColSwap DGTrouble, 1
            DGridTxtKeyDown DGTrouble, txtgrid1, Index, RsTrb, KeyCode, False, 1, frmTrouble, "frmTrouble"
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DGTrouble.Visible = False) Then
                If TxtGrid1Leave = True Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, True, CCCode, 1
                End If
            End If
            
        Case CCDesc
            If DGTrouble.Visible = False Then DGridColSwap DGTrouble, 1
            DGridTxtKeyDown DGTrouble, txtgrid1, Index, RsTrb, KeyCode, False, 1, frmTrouble, "frmTrouble"
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DGTrouble.Visible = False) Then
                If TxtGrid1Leave = True Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, True, CCCode
                End If
            End If
            
            
    End Select
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
    Select Case FGrid1.Col
        Case CCCode
            DGridTxtKeyPress txtgrid1, Index, RsTrb, KeyAscii, "Code"
        Case CCDesc
            DGridTxtKeyPress txtgrid1, Index, RsTrb, KeyAscii, "Name"
    End Select
End Sub

Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        Select Case FGrid1.Col
            Case CCCode
                If KeyCode <> 13 And DGTrouble.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0
                DGridTxtKeyUp_Mast txtgrid1, Index, RsTrb, KeyCode, "Code"
            Case CCDesc
                If KeyCode <> 13 And DGTrouble.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0
                DGridTxtKeyUp_Mast txtgrid1, Index, RsTrb, KeyCode, "Name"
                
        End Select
End Sub

Private Sub TxtGrid1_LostFocus(Index As Integer)
    'If ExitCtrl = False Then Exit Sub
    txtgrid1(Index).Visible = False
End Sub

Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGrid1Leave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid1Leave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim Repeat$
Select Case FGrid1.Col
    Case CCCode, CCDesc
        If RsTrb.RecordCount = 0 Or txtgrid1(1).TEXT = "" Or RsTrb.EOF = True Or RsTrb.BOF = True Then
            FGrid1.TextMatrix(FGrid1.Row, CCCode) = ""
            FGrid1.TextMatrix(FGrid1.Row, CCDesc) = ""
        Else
            FGrid1.TextMatrix(FGrid1.Row, 0) = FGrid1.Row
            FGrid1.TextMatrix(FGrid1.Row, CCCode) = RsTrb!Code
            FGrid1.TextMatrix(FGrid1.Row, CCDesc) = RsTrb!Name
        End If
            
        
'    Case CCDesc
'        If RsTrb.RecordCount = 0 Or txtgrid1(1).TEXT = "" Then
'            FGrid1.TextMatrix(FGrid1.Row, CCCode) = ""
'            FGrid1.TextMatrix(FGrid1.Row, CCDesc) = ""
'        Else
'            FGrid1.TextMatrix(FGrid1.Row, CCCode) = RsTrb!Code
'            If UCase(Trim(txtgrid1(1).TEXT)) <> UCase(left(RsTrb!Name, Len(Trim(txtgrid1(1).TEXT)))) Then
'                FGrid1.TextMatrix(FGrid1.Row, CCDesc) = txtgrid1(1).TEXT
'            Else
'                txtgrid1(1).TEXT = RsTrb!Name
'                FGrid1.TextMatrix(FGrid1.Row, CCDesc) = RsTrb!Name
'            End If
'        End If
End Select
TxtGrid1Leave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid1.SetFocus
    txtgrid1(1).Visible = False
End If
End Function


Sub Ini_Grid()
    With FGrid1
    
        .TextMatrix(0, 0) = "Srl."
                
        .TextMatrix(0, CCCode) = "Code"
        .ColAlignment(CCCode) = flexAlignLeftCenter
        .ColWidth(CCCode) = 1200
        
        
        .TextMatrix(0, CCDesc) = "Trouble Name"
        .ColAlignment(CCDesc) = flexAlignLeftCenter
        .ColWidth(CCDesc) = 5000
    End With
End Sub
