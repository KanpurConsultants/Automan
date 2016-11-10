VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmProCust 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Prospective Customer"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11655
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   11655
   Visible         =   0   'False
   Begin MSDataGridLib.DataGrid DGRep 
      Height          =   4935
      Left            =   5175
      Negotiate       =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   5190
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
      Caption         =   "Representative Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Representative Name"
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
   Begin MSDataGridLib.DataGrid DGArea 
      Height          =   4845
      Left            =   8580
      Negotiate       =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   4710
      Visible         =   0   'False
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   8546
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
      Caption         =   "Area Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   2729.764
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGProf 
      Height          =   4920
      Left            =   7140
      Negotiate       =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   8678
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
      Caption         =   "Profession Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Profession"
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
            ColumnWidth     =   2640.189
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   4845
      Left            =   8205
      Negotiate       =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4470
      Visible         =   0   'False
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   8546
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
      Caption         =   "City Help"
      ColumnCount     =   1
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   2910.047
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGRef 
      Height          =   4890
      Left            =   195
      Negotiate       =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   5205
      Visible         =   0   'False
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   8625
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
      Caption         =   "Referred By Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Referred By"
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
            ColumnWidth     =   3165.166
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
      Index           =   0
      Left            =   7590
      TabIndex        =   30
      Top             =   3765
      Width           =   1965
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
      Index           =   34
      Left            =   7590
      MaxLength       =   5
      TabIndex        =   24
      Top             =   2145
      Width           =   600
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
      Left            =   1485
      MaxLength       =   4
      TabIndex        =   1
      Top             =   525
      Width           =   630
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
      Left            =   2130
      MaxLength       =   40
      TabIndex        =   2
      Top             =   525
      Width           =   4275
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
      Index           =   31
      Left            =   7590
      MaxLength       =   5
      TabIndex        =   33
      Top             =   4305
      Width           =   600
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
      Index           =   32
      Left            =   7590
      MaxLength       =   1
      TabIndex        =   32
      Top             =   4035
      Width           =   600
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
      Index           =   33
      Left            =   4470
      MaxLength       =   1
      TabIndex        =   35
      Top             =   4995
      Visible         =   0   'False
      Width           =   255
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
      Left            =   7590
      MaxLength       =   5
      TabIndex        =   25
      Top             =   2415
      Width           =   600
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
      Index           =   26
      Left            =   7590
      MaxLength       =   15
      TabIndex        =   27
      Top             =   2955
      Width           =   1965
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
      Index           =   25
      Left            =   7590
      MaxLength       =   5
      TabIndex        =   26
      Top             =   2685
      Width           =   600
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
      Index           =   27
      Left            =   7590
      MaxLength       =   12
      TabIndex        =   28
      Top             =   3225
      Width           =   1560
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
      Index           =   28
      Left            =   4275
      MaxLength       =   12
      TabIndex        =   31
      Top             =   5280
      Visible         =   0   'False
      Width           =   1905
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
      Index           =   29
      Left            =   7590
      MaxLength       =   12
      TabIndex        =   29
      Top             =   3495
      Width           =   1560
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
      Index           =   30
      Left            =   4275
      MaxLength       =   5
      TabIndex        =   34
      Top             =   5595
      Visible         =   0   'False
      Width           =   945
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   661
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   9450
      TabIndex        =   55
      Top             =   2625
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   510
         TabIndex        =   56
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
      Left            =   2130
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1065
      Width           =   4275
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
      Left            =   1485
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1065
      Width           =   630
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
      Left            =   1485
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1605
      Width           =   4275
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   12
      Left            =   1485
      TabIndex        =   12
      Text            =   "0123456789012345678901234"
      Top             =   1875
      Width           =   2895
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
      Left            =   1485
      MaxLength       =   24
      TabIndex        =   17
      Top             =   2955
      Width           =   4275
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
      Left            =   1485
      MaxLength       =   24
      TabIndex        =   19
      Top             =   3495
      Width           =   4275
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
      Left            =   1485
      MaxLength       =   35
      TabIndex        =   15
      Top             =   2415
      Width           =   4275
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
      Left            =   4935
      MaxLength       =   6
      TabIndex        =   13
      Text            =   "012345"
      Top             =   1875
      Width           =   825
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
      Left            =   1485
      MaxLength       =   40
      TabIndex        =   21
      Top             =   4035
      Width           =   4275
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   16
      Left            =   1485
      MaxLength       =   35
      TabIndex        =   16
      Top             =   2685
      Width           =   4275
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   14
      Left            =   1485
      MaxLength       =   6
      TabIndex        =   14
      Top             =   2145
      Width           =   1470
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
      Left            =   6420
      MaxLength       =   1
      TabIndex        =   3
      Top             =   525
      Width           =   225
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
      Left            =   1485
      MaxLength       =   15
      TabIndex        =   4
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
      Index           =   7
      Left            =   1485
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1335
      Width           =   4275
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
      Left            =   1485
      MaxLength       =   40
      TabIndex        =   23
      Top             =   4575
      Visible         =   0   'False
      Width           =   4275
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
      Left            =   1485
      MaxLength       =   40
      TabIndex        =   22
      Top             =   4305
      Width           =   4275
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
      Left            =   1485
      MaxLength       =   40
      TabIndex        =   20
      Top             =   3765
      Width           =   4275
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
      Left            =   1485
      MaxLength       =   50
      TabIndex        =   18
      Top             =   3225
      Width           =   4275
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
      Left            =   7320
      TabIndex        =   10
      Top             =   1875
      Width           =   4155
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
      Left            =   7320
      MaxLength       =   40
      TabIndex        =   9
      Top             =   1605
      Width           =   4155
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
      Left            =   7320
      MaxLength       =   40
      TabIndex        =   8
      Top             =   1335
      Width           =   4155
   End
   Begin MSDataGridLib.DataGrid DGMod 
      Height          =   2865
      Left            =   225
      Negotiate       =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   4890
      Visible         =   0   'False
      Width           =   11385
      _ExtentX        =   20082
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
      Caption         =   "Model Help"
      ColumnCount     =   3
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
      BeginProperty Column02 
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   7304.882
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model (possible)"
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
      Left            =   6135
      TabIndex        =   77
      Top             =   3780
      Width           =   1380
   End
   Begin VB.Label LblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Code"
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
      Left            =   9510
      TabIndex        =   72
      Top             =   615
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt. Y/N"
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
      Index           =   2
      Left            =   6135
      TabIndex        =   71
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
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
      Left            =   105
      TabIndex        =   70
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      Height          =   780
      Left            =   7500
      Top             =   510
      Width           =   3945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Index           =   3
      Left            =   105
      TabIndex        =   69
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Last"
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
      Index           =   16
      Left            =   6135
      TabIndex        =   68
      Top             =   4050
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Active Y/N"
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
      Index           =   15
      Left            =   6135
      TabIndex        =   67
      Top             =   4320
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
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
      Index           =   14
      Left            =   2820
      TabIndex        =   66
      Top             =   5010
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Sale Y/N"
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
      Index           =   13
      Left            =   6135
      TabIndex        =   65
      Top             =   2700
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Model"
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
      Index           =   12
      Left            =   6135
      TabIndex        =   64
      Top             =   2970
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Vehicle Y/N"
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
      Left            =   6135
      TabIndex        =   63
      Top             =   2430
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Call Status"
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
      Left            =   2625
      TabIndex        =   62
      Top             =   5610
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Visit Date"
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
      Left            =   2625
      TabIndex        =   61
      Top             =   5355
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Sale Date"
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
      Left            =   6135
      TabIndex        =   60
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Visit Date"
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
      Left            =   6135
      TabIndex        =   59
      Top             =   3510
      Width           =   1185
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division                     :"
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
      Left            =   7605
      TabIndex        =   58
      Top             =   930
      Width           =   1650
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
      Left            =   9750
      TabIndex        =   57
      Top             =   930
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Office"
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
      Left            =   105
      TabIndex        =   54
      Top             =   2430
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FAX"
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
      Left            =   105
      TabIndex        =   53
      Top             =   3510
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
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
      Left            =   105
      TabIndex        =   52
      Top             =   2970
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
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
      Index           =   38
      Left            =   105
      TabIndex        =   51
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
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
      Left            =   4530
      TabIndex        =   50
      Top             =   1890
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
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
      Index           =   45
      Left            =   105
      TabIndex        =   49
      Top             =   810
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Second Rep"
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
      Index           =   43
      Left            =   105
      TabIndex        =   48
      Top             =   4590
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profession"
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
      Index           =   46
      Left            =   105
      TabIndex        =   47
      Top             =   1350
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FName"
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
      Height          =   255
      Index           =   39
      Left            =   1485
      TabIndex        =   46
      Top             =   1065
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resi"
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
      Index           =   28
      Left            =   765
      TabIndex        =   45
      Top             =   2700
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STD Code"
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
      Index           =   31
      Left            =   105
      TabIndex        =   44
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City Name"
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
      Index           =   33
      Left            =   105
      TabIndex        =   43
      Top             =   1890
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reffered By"
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
      Index           =   44
      Left            =   105
      TabIndex        =   42
      Top             =   4050
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMail"
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
      Left            =   105
      TabIndex        =   41
      Top             =   3255
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
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
      Index           =   41
      Left            =   105
      TabIndex        =   40
      Top             =   3780
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Executive"
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
      Index           =   40
      Left            =   105
      TabIndex        =   39
      Top             =   4320
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Code      :"
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
      Index           =   4
      Left            =   7605
      TabIndex        =   38
      Top             =   615
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   6120
      TabIndex        =   37
      Top             =   1350
      Width           =   690
   End
End
Attribute VB_Name = "frmProCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim RsCity As ADODB.Recordset
Dim RsMod  As ADODB.Recordset
Dim RsRef As ADODB.Recordset
Dim RsRep As ADODB.Recordset
Dim RsProf As ADODB.Recordset
Dim RsArea As ADODB.Recordset
Dim Master As ADODB.Recordset
Public DGridInsertCall As Boolean
Dim DocID As String * 8
'Call_Status LastDate NextDateTime Active_YN Close_Lost LABEL
Private Const Model As Byte = 0
Private Const NPrefix As Byte = 1
Private Const PName As Byte = 2
Private Const NSuffix As Byte = 3
Private Const FPrefix As Byte = 4
Private Const fname As Byte = 5
Private Const Religion As Byte = 6
Private Const Profession As Byte = 7
Private Const Add1 As Byte = 8
Private Const Add2 As Byte = 9
Private Const Add3 As Byte = 10
Private Const ConPerson As Byte = 11
Private Const City As Byte = 12
Private Const Pin As Byte = 13
Private Const STD As Byte = 14
Private Const PhoneOff As Byte = 15
Private Const PhoneResi As Byte = 16
Private Const Mobile  As Byte = 17
Private Const EMail As Byte = 18
Private Const FAx As Byte = 19
Private Const Area As Byte = 20
Private Const RefPer As Byte = 21
Private Const Rep1 As Byte = 22
Private Const Rep2 As Byte = 23
Private Const Govt_YN As Byte = 34
Private Const FirstVeh As Byte = 24
Private Const FirstSale As Byte = 25
Private Const LModel As Byte = 26
Private Const LSaleDate As Byte = 27
Private Const LastDate As Byte = 28
Private Const NextDate As Byte = 29
Private Const CallStat As Byte = 30
Private Const ActiveYN As Byte = 31
Private Const CloseYN As Byte = 32
Private Const Lbl As Byte = 33
Dim ListArray As Variant
Dim mListItem As ListItem

Private Sub DGMod_Click()
    If RsMod.RecordCount > 0 Then
        Txt(Model).TEXT = RsMod!Code
        Txt(Model).Tag = RsMod!Code
    End If
    Txt(Model).SetFocus
    DGMod.Visible = False
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

'Code Site_Code NPrefix Name NSuffix FPrefix FName Govt_YN ConPerson Add1 Add2 Add3 CityCode PIN STD PhoneOff PhoneResi
'Mobile  FAX EMail AREA  REF_CODE REP_CODE REP_CODE2 Profession Religion FirstVeh_YN FirstSal_YN L_Model L_Sale_Date
'Call_Status LastDate NextDateTime Active_YN Close_Lost LABEL

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo ELoop

'Dim i As Byte
WinSetting Me, 6405, 11655
TopCtrl1.Tag = PubUParam: Ini_Grid
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
     Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    If PubMoveRecYn Then
        Master.Open "select ProspectiveCust.cust_CODE as searchcode,ProspectiveCust.* from ProspectiveCust " & sitecond & " ", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 ProspectiveCust.cust_CODE as searchcode,ProspectiveCust.* from ProspectiveCust " & sitecond & " ", GCn, adOpenDynamic, adLockOptimistic
    End If
   
    Set RsCity = New ADODB.Recordset
    RsCity.CursorLocation = adUseClient
    RsCity.Open "select citycode as code,cityname as name from city order by cityname,citycode", GCn, adOpenDynamic, adLockOptimistic
    Set DGCity.DataSource = RsCity
    
    Set RsRef = New ADODB.Recordset
    RsRef.CursorLocation = adUseClient
    RsRef.Open "select RefCode as code,RefName as name from reffered order by Refname", GCn, adOpenDynamic, adLockOptimistic
    Set DGRef.DataSource = RsRef
  
    Set RsArea = New ADODB.Recordset
    RsArea.CursorLocation = adUseClient
    RsArea.Open "select AreaCode as code,AreaName as name from Area order by AreaName", GCn, adOpenDynamic, adLockOptimistic
    Set DGArea.DataSource = RsArea
    
    Set RsProf = New ADODB.Recordset
    RsProf.CursorLocation = adUseClient
    RsProf.Open "select ProfessionCode as code,Professionname as name from Profession order by Professionname", GCn, adOpenDynamic, adLockOptimistic
    Set DGProf.DataSource = RsProf
  
    Set RsRep = New ADODB.Recordset
    RsRep.CursorLocation = adUseClient
    RsRep.Open "select Emp_code as code,emp_name as name from emp_mast where emp_type = 0  order by Emp_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGRep.DataSource = RsRep
    
    Set RsMod = New ADODB.Recordset
    RsMod.CursorLocation = adUseClient
    RsMod.Open "select Model as code,Model_Desc as NAME, Chas_Type from model where (div_code='" & PubDivCode & "' or Div_Code='') order by model", GCn, adOpenDynamic, adLockOptimistic
    Set DGMod.DataSource = RsMod
'    If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
    MoveRec
    Disp_Text SETS("INI", Me, Master)
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
DGridInsertCall = False
Set RsCity = Nothing
Set RsProf = Nothing
Set RsRep = Nothing
Set RsRef = Nothing
Set RsArea = Nothing
Set Master = Nothing
Set RsMod = Nothing
End Sub

Private Sub ListView_Click()
Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
FrmList.Visible = False
Txt(Val(ListView.Tag)).SetFocus
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim VNo As Long
'Dim i As Integer

    Disp_Text SETS("ADD", Me, Master)
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    Call BlankText
    If GCn.Execute("select count(*) from ProspectiveCust where site_code = '" & PubSiteCode & "'").Fields(0).Value > 0 Then
        VNo = GCn.Execute("select MAX(right(cust_CODE,7)) from ProspectiveCust  where site_code = '" & PubSiteCode & "'").Fields(0).Value + 1
    Else
        VNo = 1
    End If
    DocID = PubSiteCode + Space(7 - Len(CStr(VNo))) + CStr(VNo)
    LblCode.CAPTION = DocID
    Txt(Govt_YN) = "No"
    Txt(ActiveYN) = "Yes"
    If DGridInsertCall Then
        Label3(7).Visible = False
        Txt(NextDate).Visible = False
        Label3(6).Visible = False
        Txt(Model).Visible = False
    Else
        Label3(7).Visible = True
        Txt(NextDate).Visible = True
        Label3(6).Visible = True
        Txt(Model).Visible = True
    End If
    Txt(NPrefix).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant, mTrans As Boolean
Dim CheckStr1$, CheckStr2$, CheckStr3$
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        CheckStr1 = "Select Party_Code from Veh_Quot where Party_Code='" & Master!Cust_code & "'"
        CheckStr2 = "Select PartyCode from Veh_SubGroupQuot where ProspectiveCust_SubGroup=0 and PartyCode='" & Master!Cust_code & "'"
        CheckStr3 = "Select ProspectiveCust_SubGroup from Visits where ProspectiveCust_SubGroup=0 and Party_Code='" & Master!Cust_code & "'"
        If GCn.Execute(CheckStr1).RecordCount > 0 Or _
            GCn.Execute(CheckStr2).RecordCount > 0 Or _
            GCn.Execute(CheckStr3).RecordCount > 0 Then
             MsgBox "Transaction Exists !" & vbCrLf & "Can't Delete this Reocord", vbInformation, "Validation"
             Exit Sub
        End If
        vBook = Master.AbsolutePosition
        GCn.BeginTrans
        mTrans = True
        GCn.Execute ("Delete from ProspectiveCust where cust_CODE = '" & Master!Cust_code & "'")
        GCn.CommitTrans
        mTrans = False
        Master.Requery
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
        End If
        Call MoveRec
        BUTTONS True, Me, Master, 0
    End If
Exit Sub
eloop1:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    Txt(NextDate).Enabled = False
    Txt(Model).Enabled = False
    
    Txt(NPrefix).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then CheckError
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
      Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    GSQL = "select ProspectiveCust.cust_CODE as searchcode,Name+NSuffix as NameSfx,FName,ConPerson,Add1,Add2 from ProspectiveCust " & sitecond & " order by name,nsuffix"
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
        Set Master = GCn.Execute("select ProspectiveCust.cust_CODE as searchcode,ProspectiveCust.* from ProspectiveCust Where ProspectiveCust.cust_CODE = '" & MyValue & "' ")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    CheckError
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
If MasterFormExit Then Unload Me: Exit Sub
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
Dim Rst As ADODB.Recordset
Dim mQry As String
  Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(ProspectiveCust.site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    mQry = "SELECT City.CityName, Reffered.RefName AS RefferPerson, Area.AreaName, Profession.ProfessionName, Emp_Mast.Emp_Name AS RepName, ProspectiveCust.NPrefix, ProspectiveCust.Name, ProspectiveCust.NSuffix, ProspectiveCust.Add1, ProspectiveCust.Add2, ProspectiveCust.Add3, ProspectiveCust.PIN, ProspectiveCust.STD, ProspectiveCust.PhoneOff, ProspectiveCust.PhoneResi, ProspectiveCust.Mobile, ProspectiveCust.NextDateTime as NextDate, ProspectiveCust.Call_Status, ProspectiveCust.EMail, ProspectiveCust.FAX " & _
    "FROM ((((ProspectiveCust LEFT JOIN City ON ProspectiveCust.CityCode = City.CityCode) LEFT JOIN Reffered ON ProspectiveCust.REF_CODE = Reffered.RefCode) LEFT JOIN Profession ON ProspectiveCust.Profession = Profession.ProfessionCode) LEFT JOIN Area ON ProspectiveCust.AREA = Area.AreaCode) LEFT JOIN Emp_Mast ON ProspectiveCust.REP_CODE = Emp_Mast.Emp_Code " & sitecond & " "
       
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub

        CreateFieldDefFile Rst, PubRepoPath + "\VehProCust.TTX", True
        Set rpt = rdApp.OpenReport(PubRepoPath & "\VehProCust.RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
        
        Call Report_View(rpt, "Prospective Customer List")
End Sub

Private Sub TopCtrl1_eRef()
    RsMod.Requery
    RsCity.Requery
    RsRep.Requery
    RsRef.Requery
    RsProf.Requery
    RsArea.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer, mVisitDate As Date
    Dim mTrans As Boolean
    Dim Relg As Integer
    Dim StCall As Integer
    Dim Srlno As Integer, VisitObj As String

    On Error GoTo errlbl
    Grid_Hide
    RemoveTxtNull
    If IsValid(Txt(NPrefix), "Name Prefix") = False Then Exit Sub
    If IsValid(Txt(PName), "Name") = False Then Exit Sub
    If IsValid(Txt(Profession), "Profession") = False Then Exit Sub
    If IsValid(Txt(ConPerson), "Contact Person") = False Then Exit Sub
    If IsValid(Txt(City), "City") = False Then Exit Sub
    If IsValid(Txt(Area), "Area") = False Then Exit Sub
    If IsValid(Txt(RefPer), "Reffered By") = False Then Exit Sub
    If IsValid(Txt(Rep1), "Sales Executive") = False Then Exit Sub
'    If IsValid(Txt(Rep2), "Second Rep") = False Then Exit Sub
    
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Txt(NextDate) <> "" Then
            If CDate(Format(Txt(NextDate), "dd/mmm/yyyy")) < CDate(Format(PubLoginDate, "dd/mmm/yyyy")) Then
                MsgBox "Next Date " & Txt(NextDate) & " is Greater than User Login Date!", vbOKOnly, "Visit Date Validation"
                Txt(NextDate).SetFocus
                Exit Sub
'                If MsgBox("Next Date " & Txt(NextDate) & " is Greater than User Login Date, Proceed ?", vbYesNo + vbCritical + vbDefaultButton2, "Check for Visit Entry") = vbYes Then
'                    mVisitDate = CDate(Txt(NextDate)) - 1
'                Else
'                    Txt(NextDate).SetFocus: Exit Sub
'                End If
            Else
                mVisitDate = PubLoginDate
            End If
        Else
            mVisitDate = PubLoginDate
        End If
        GSQL = "Select Name from ProspectiveCust where Name+NSuffix= '" & Txt(PName).TEXT & Txt(NSuffix).TEXT & "'"
        If GCn.Execute(GSQL).RecordCount > 0 Then
            MsgBox "Prospective Customer Name already exists", vbOKOnly, "Duplicate Check"
            Txt(PName).SetFocus: Exit Sub
        End If
    End If
    Relg = FxReligion(Txt(Religion).TEXT)
    
    Select Case Txt(CallStat).TEXT
        Case "Cold"
            StCall = 0
        Case "Warm"
            StCall = 1
        Case "Hot"
            StCall = 2
    End Select
    GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        GCn.Execute ("delete from ProspectiveCust where cust_Code='" & DocID & "'")
        GCn.Execute ("insert into ProspectiveCust(cust_Code,Site_Code,NPrefix, Name,NSuffix,FPrefix,FName,Govt_YN,ConPerson,Add1,Add2,Add3,CityCode,PIN,STD,PhoneOff,PhoneResi, " & _
            " Mobile,FAX,EMail,AREA,REF_CODE,REP_CODE,REP_CODE2,Profession,Religion,FirstVeh_YN,FirstSal_YN,L_Model,L_Sale_Date,Call_Status,LastDate,NextDateTime,Active_YN,Close_Lost,LABEL, U_Name, U_EntDt, U_AE ) " & _
            " values('" & DocID & "','" & PubSiteCode & "','" & Txt(NPrefix).TEXT & "','" & Txt(PName).TEXT & "' ,'" & Txt(NSuffix).TEXT & "','" & Txt(FPrefix).TEXT & "','" & Txt(fname).TEXT & "'," & IIf(Txt(Govt_YN).TEXT = "Yes", 1, 0) & ",'" & Txt(ConPerson).TEXT & "','" & Txt(Add1).TEXT & "','" & Txt(Add2).TEXT & "','" & Txt(Add3).TEXT & "','" & Txt(City).Tag & "','" & Txt(Pin).TEXT & "','" & Txt(STD).TEXT & "','" & Txt(PhoneOff).TEXT & "','" & Txt(PhoneResi).TEXT & "', " & _
            " '" & Txt(Mobile).TEXT & "','" & Txt(FAx).TEXT & "','" & Txt(EMail).TEXT & "','" & Txt(Area).Tag & "','" & Txt(RefPer).Tag & "','" & Txt(Rep1).Tag & "','" & Txt(Rep2).Tag & "','" & Txt(Profession).Tag & "'," & Relg & "," & IIf(Txt(FirstVeh).TEXT = "Yes", 1, 0) & "," & IIf(Txt(FirstSale).TEXT = "Yes", 1, 0) & ",'" & Txt(LModel).TEXT & "'," & ConvertDate(Txt(LSaleDate).TEXT) & "," & StCall & "," & ConvertDate(Txt(LastDate).TEXT) & "," & ConvertDate(Txt(NextDate).TEXT) & "," & IIf(Txt(ActiveYN).TEXT = "Yes", 1, 0) & ",'" & Txt(CloseYN).TEXT & "','" & Txt(Lbl).TEXT & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
        If Txt(NextDate) <> "" Then
            VisitObj = GCn.Execute("select VisitObjcode from syctrl").Fields(0).Value
            Srlno = GCn.Execute("select " & vIsNull("max(srlno)", "1") & " from Visits where visitdate = " & ConvertDate(PubLoginDate) & " AND Rep_Code = '" & Txt(Rep1).Tag & "' AND  Site_Code = '" & PubSiteCode & "'").Fields(0).Value + 1
            GCn.Execute ("Insert into visits( VisitDate, Rep_Code, SrlNo, Div_Code, Site_Code, ProspectiveCust_SubGroup, Party_Code,NewEnquiry, " & _
                " Visit_Call, OBJECTIVE,Call_Status, NEXT_DATE,U_Name, U_EntDt, U_AE ) " & _
                " values(" & ConvertDate(mVisitDate) & ", '" & Txt(Rep1).Tag & "', " & Srlno & ", '" & PubDivCode & "', '" & PubSiteCode & "',0,'" & DocID & "',1, " & _
                "1,'" & VisitObj & "',0, " & ConvertDate(Txt(NextDate)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
            If Txt(Model) <> "" Then
                GCn.Execute ("delete from Veh_SubGroupQuot where PartyCode ='" & DocID & "' and  Model = '" & Txt(Model) & "'")
                GCn.Execute ("insert into Veh_SubGroupQuot(PartyCode, StartDate, Model, " & _
                    "ProspectiveCust_SubGroup, QuotDocId, QuotSrl_No, " & _
                    "Site_Code,REP_CODE,Call_Status, " & _
                    " U_Name, U_EntDt, U_AE ) " & _
                    " values('" & DocID & "'," & ConvertDate(mVisitDate) & ",'" & Txt(Model) & _
                    "',0,'',0,'" & PubSiteCode & PubSiteCode & "','" & Txt(Rep1).Tag & "',0, " & _
                    " '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
            End If
        End If
    Else
        GCn.Execute ("update ProspectiveCust set NPrefix='" & Txt(NPrefix).TEXT & "', Name='" & Txt(PName).TEXT & "',NSuffix='" & Txt(NSuffix).TEXT & "',FPrefix='" & Txt(FPrefix).TEXT & "',FName='" & Txt(fname).TEXT & "',Govt_YN=" & IIf(Txt(Govt_YN).TEXT = "Yes", 1, 0) & ",ConPerson='" & Txt(ConPerson).TEXT & "',Add1='" & Txt(Add1).TEXT & "',Add2='" & Txt(Add2).TEXT & "',Add3='" & Txt(Add3).TEXT & "',CityCode='" & Txt(City).Tag & "',PIN='" & Txt(Pin).TEXT & "',STD='" & Txt(STD).TEXT & "',PhoneOff='" & Txt(PhoneOff).TEXT & "',PhoneResi='" & Txt(PhoneResi).TEXT & "', " & _
            " Mobile='" & Txt(Mobile).TEXT & "',FAX='" & Txt(FAx).TEXT & "',EMail='" & Txt(EMail).TEXT & "',AREA='" & Txt(Area).Tag & "',REF_CODE='" & Txt(RefPer).Tag & "',REP_CODE='" & Txt(Rep1).Tag & "',REP_CODE2='" & Txt(Rep2).Tag & "',Profession='" & Txt(Profession).Tag & "',Religion=" & Relg & ",FirstVeh_YN=" & IIf(Txt(FirstVeh).TEXT = "Yes", 1, 0) & ",FirstSal_YN=" & IIf(Txt(FirstSale).TEXT = "Yes", 1, 0) & ",L_Model='" & Txt(LModel).TEXT & "',L_Sale_Date=" & ConvertDate(Txt(LSaleDate).TEXT) & ",Call_Status=" & StCall & ",LastDate=" & ConvertDate(Txt(LastDate).TEXT) & ",Active_YN=" & IIf(Txt(ActiveYN).TEXT = "Yes", 1, 0) & ",Close_Lost='" & Txt(CloseYN).TEXT & "',LABEL='" & Txt(Lbl).TEXT & "', U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE= 'A' where cust_Code = '" & DocID & "'")
    End If
GCn.CommitTrans
If MasterFormExit Then Unload Me: Exit Sub
mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select ProspectiveCust.cust_CODE as searchcode,ProspectiveCust.* from ProspectiveCust Where ProspectiveCust.cust_CODE = '" & DocID & "' ")
    End If
    Master.FIND "searchcode = '" & DocID & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Grid_Hide
Ctrl_GetFocus Txt(Index)
Select Case Index
    Case Religion
        ListArray = Array("N/A", "Hindu", "Muslim", "Sikh", "Christian")
        Set mListItem = ListView_Items(ListView, Txt, Religion, ListArray, 5)
    Case NPrefix
        ListArray = Array("Mr.", "Mrs.", "Miss", "M/S")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 4)
    Case FPrefix
        ListArray = Array("S/O", "W/O", "D/O", "C/O", "And ", "U/C")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 6)
    Case CallStat
        ListArray = Array("Cold", "Warm", "Hot")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 3)
    Case City
        If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or Txt(City).TEXT = "" Then Exit Sub
        If Txt(City).TEXT <> RsCity!Name Then
            RsCity.MoveFirst
            RsCity.FIND "name ='" & Txt(City).TEXT & "'"
        End If
    Case Area
        If RsArea.RecordCount = 0 Or (RsArea.EOF = True Or RsArea.BOF = True) Or Txt(Area).TEXT = "" Then Exit Sub
        If Txt(Area).TEXT <> RsArea!Name Then
            RsArea.MoveFirst
            RsArea.FIND "name ='" & Txt(Area).TEXT & "'"
        End If

    Case RefPer
        If RsRef.RecordCount = 0 Or (RsRef.EOF = True Or RsRef.BOF = True) Or Txt(RefPer).TEXT = "" Then Exit Sub
        If Txt(RefPer).TEXT <> RsRef!Name Then
            RsRef.MoveFirst
            RsRef.FIND "name ='" & Txt(RefPer).TEXT & "'"
        End If
    Case Rep1
        DGRep.Tag = 1
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or Txt(Rep1).TEXT = "" Then Exit Sub
        If Txt(Rep1).TEXT <> RsRep!Name Then
            RsRep.MoveFirst
            RsRep.FIND "name ='" & Txt(Rep1).TEXT & "'"
        End If
    Case Rep2
        DGRep.Tag = 2
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or Txt(Rep2).TEXT = "" Then Exit Sub
        If Txt(Rep2).TEXT <> RsRep!Name Then
            RsRep.MoveFirst
            RsRep.FIND "name ='" & Txt(Rep2).TEXT & "'"
        End If
    Case Profession
        If RsProf.RecordCount = 0 Or (RsProf.EOF = True Or RsProf.BOF = True) Or Txt(Profession).TEXT = "" Then Exit Sub
        If Txt(Profession).TEXT <> RsProf!Name Then
            RsProf.MoveFirst
            RsProf.FIND "name ='" & Txt(Profession).TEXT & "'"
        End If
    Case Model
        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or Txt(Model).TEXT = "" Then Exit Sub
        If Txt(Model).TEXT <> RsMod!Code Then
            RsMod.MoveFirst
            RsMod.FIND "code ='" & Txt(Model).TEXT & "'"
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
    Case NPrefix
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1200
    Case FPrefix
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1800
    Case Religion
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1500
    Case CallStat
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 900
    Case City
        DGridTxtKeyDown DGCity, Txt, City, RsCity, KeyCode, False, 1, frmCity, "frmCity"
    Case Area
        DGridTxtKeyDown DGArea, Txt, Index, RsArea, KeyCode, False, 1, frmArea, "frmArea"
    Case RefPer
        DGridTxtKeyDown DGRef, Txt, Index, RsRef, KeyCode, False, 1
    Case Rep1
        DGridTxtKeyDown DGRep, Txt, Index, RsRep, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
    Case Rep2
        DGridTxtKeyDown DGRep, Txt, Index, RsRep, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
    Case Profession
        DGridTxtKeyDown DGProf, Txt, Index, RsProf, KeyCode, False, 1
    Case Model
        DGridTxtKeyDown DGMod, Txt, Index, RsMod, KeyCode, False, 0, frmModel, "frmModel"
End Select
If DGMod.Visible = False And FrmList.Visible = False And DGCity.Visible = False And DGRep.Visible = False And DGRef.Visible = False And DGProf.Visible = False And DGArea.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> ActiveYN Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = ActiveYN Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If Index <> NPrefix Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case Index
    Case City
        If DGCity.Visible = True Then DGridTxtKeyPress Txt, Index, RsCity, KeyAscii, "Name"
    Case Area
        If DGArea.Visible = True Then DGridTxtKeyPress Txt, Index, RsArea, KeyAscii, "Name"
    Case Rep1
        If DGRep.Visible = True Then DGridTxtKeyPress Txt, Index, RsRep, KeyAscii, "Name"
    Case Rep2
        If DGRep.Visible = True Then DGridTxtKeyPress Txt, Index, RsRep, KeyAscii, "Name"
    Case RefPer
        If DGRef.Visible = True Then DGridTxtKeyPress Txt, Index, RsRef, KeyAscii, "Name"
    Case Profession
        If DGProf.Visible = True Then DGridTxtKeyPress Txt, Index, RsProf, KeyAscii, "Name"
    Case FirstVeh, FirstSale, ActiveYN, Govt_YN, CloseYN
        If UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            Txt(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            Txt(Index) = ""
        End If
        KeyAscii = 0
    Case Model
        If DGMod.Visible = True Then DGridTxtKeyPress Txt, Index, RsMod, KeyAscii, "code"
End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case Religion, CallStat, NPrefix, FPrefix
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
        
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case City
        If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsCity!Name
            Txt(Index).Tag = RsCity!Code
        End If
    Case Area
        If RsArea.RecordCount = 0 Or (RsArea.EOF = True Or RsArea.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsArea!Name
            Txt(Index).Tag = RsArea!Code
        End If
    Case Rep1
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsRep!Name
            Txt(Index).Tag = RsRep!Code
        End If
    Case Rep2
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsRep!Name
            Txt(Index).Tag = RsRep!Code
        End If
    Case Profession
        If RsProf.RecordCount = 0 Or (RsProf.EOF = True Or RsProf.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsProf!Name
            Txt(Index).Tag = RsProf!Code
        End If
    Case RefPer
        If RsRef.RecordCount = 0 Or (RsRef.EOF = True Or RsRef.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsRef!Name
            Txt(Index).Tag = RsRef!Code
        End If
    Case NPrefix, FPrefix, Religion, CallStat
        If Txt(Index).TEXT <> "" Then Txt(Index).TEXT = ListView.SelectedItem.TEXT
    Case LSaleDate, LastDate, NextDate
        Txt(Index).TEXT = RetDate(Txt(Index))
    Case Model
        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsMod!Code
            Txt(Index).Tag = RsMod!Code
        End If
End Select
End Sub

Private Sub DGCity_Click()
    DGCity.Visible = False
    If RsCity.RecordCount > 0 Then
        Txt(City).TEXT = RsCity!Name
        Txt(City).Tag = RsCity!Code
    End If
    Txt(City).SetFocus
End Sub
Private Sub DGArea_Click()
    DGArea.Visible = False
    If RsArea.RecordCount > 0 Then
        Txt(Area).TEXT = RsArea!Name
        Txt(Area).Tag = RsArea!Code
    End If
    Txt(Area).SetFocus
End Sub

Private Sub DGProf_Click()
    DGProf.Visible = False
    If RsProf.RecordCount > 0 Then
        Txt(Profession).TEXT = RsProf!Name
        Txt(Profession).Tag = RsProf!Code
    End If
    Txt(Profession).SetFocus

End Sub

Private Sub DGRef_Click()
    DGRef.Visible = False
    If RsRef.RecordCount > 0 Then
        Txt(RefPer).TEXT = RsRef!Name
        Txt(RefPer).Tag = RsRef!Code
    End If
    Txt(RefPer).SetFocus

End Sub

Private Sub DGRep_Click()
    DGRep.Visible = False
    If DGRep.Tag = 1 Then
        If RsRep.RecordCount > 0 Then
            Txt(Rep1).TEXT = RsRep!Name
            Txt(Rep1).Tag = RsRep!Code
        End If
        Txt(Rep1).SetFocus
    ElseIf DGRep.Tag = 2 Then
        If RsRep.RecordCount > 0 Then
            Txt(Rep2).TEXT = RsRep!Name
            Txt(Rep2).Tag = RsRep!Code
        End If
        Txt(Rep2).SetFocus
    End If
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
    Txt(I).Tag = ""
    
Next I
End Sub

Private Sub MoveRec()
On Error GoTo error1
If Master.RecordCount > 0 Then
    DocID = Master!Cust_code
    LblCode = Master!Cust_code
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & Master!Site_Code
    Txt(NPrefix) = IIf(IsNull(Master!NPrefix), "", Master!NPrefix)
    Txt(PName) = IIf(IsNull(Master!Name), "", Master!Name)
    Txt(NSuffix) = IIf(IsNull(Master!NSuffix), "", Master!NSuffix)
    Txt(FPrefix) = IIf(IsNull(Master!FPrefix), "", Master!FPrefix)
    Txt(fname) = IIf(IsNull(Master!fname), "", Master!fname)
    Txt(Religion).TEXT = FxReligion(IIf(IsNull(Master!Religion), 0, Master!Religion))
    Txt(Add1) = IIf(IsNull(Master!Add1), "", Master!Add1)
    Txt(Add2) = IIf(IsNull(Master!Add2), "", Master!Add2)
    Txt(Add3) = IIf(IsNull(Master!Add3), "", Master!Add3)
    Txt(ConPerson) = IIf(IsNull(Master!ConPerson), "", Master!ConPerson)
    
    Txt(Profession).Tag = IIf(IsNull(Master!Profession), "", Master!Profession)
    If Txt(Profession).Tag <> "" And GCn.Execute("select Professionname from Profession where Professioncode = '" & Txt(Profession).Tag & "'").RecordCount > 0 Then
        Txt(Profession).TEXT = GCn.Execute("select Professionname from Profession where Professioncode = '" & Txt(Profession).Tag & "'").Fields(0).Value
    Else
        Txt(Profession).TEXT = ""
    End If
    Txt(City).Tag = IIf(IsNull(Master!CityCode), "", Master!CityCode)
    If Txt(City).Tag <> "" And GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").RecordCount > 0 Then
        Txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").Fields(0).Value
    Else
        Txt(City).TEXT = ""
    End If
    Txt(Area).Tag = IIf(IsNull(Master!Area), "", Master!Area)
    If Txt(Area).Tag <> "" And GCn.Execute("select AREAname from AREA where AREAcode = '" & Txt(Area).Tag & "'").RecordCount > 0 Then
        Txt(Area).TEXT = GCn.Execute("select AREAname from AREA where AREAcode = '" & Txt(Area).Tag & "'").Fields(0).Value
    Else
        Txt(Area).TEXT = ""
    End If
    Txt(Rep1).Tag = IIf(IsNull(Master!REP_CODE), "", Master!REP_CODE)
    If Txt(Rep1).Tag <> "" And GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(Rep1).Tag & "'").RecordCount > 0 Then
        Txt(Rep1).TEXT = GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(Rep1).Tag & "'").Fields(0).Value
    Else
        Txt(Rep1).TEXT = ""
    End If

    Txt(Rep2).Tag = IIf(IsNull(Master!Rep_Code2), "", Master!Rep_Code2)
    If Txt(Rep2).Tag <> "" And GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(Rep2).Tag & "'").RecordCount > 0 Then
        Txt(Rep2).TEXT = GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(Rep2).Tag & "'").Fields(0).Value
    Else
        Txt(Rep2).TEXT = ""
    End If
    
    Txt(RefPer).Tag = IIf(IsNull(Master!REF_CODE), "", Master!REF_CODE)
    If Txt(RefPer).Tag <> "" And GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(RefPer).Tag & "'").RecordCount > 0 Then
        Txt(RefPer).TEXT = GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(RefPer).Tag & "'").Fields(0).Value
    Else
        Txt(RefPer).TEXT = ""
    End If
    
    Txt(Pin) = IIf(IsNull(Master!Pin), "", Master!Pin)
    Txt(STD) = IIf(IsNull(Master!STD), "", Master!STD)
    Txt(PhoneOff) = IIf(IsNull(Master!PhoneOff), "", Master!PhoneOff)
    Txt(PhoneResi) = IIf(IsNull(Master!PhoneResi), "", Master!PhoneResi)
    Txt(Mobile) = IIf(IsNull(Master!Mobile), "", Master!Mobile)
    Txt(EMail) = IIf(IsNull(Master!EMail), "", Master!EMail)
    Txt(FAx) = IIf(IsNull(Master!FAx), "", Master!FAx)
    Txt(Govt_YN) = IIf(Master!Govt_YN = 1, "Yes", "No")
    Txt(FirstVeh) = IIf(Master!FirstVeh_YN = 1, "Yes", "No")
    Txt(FirstSale) = IIf(Master!FirstSal_YN = 1, "Yes", "No")
    Txt(LModel) = IIf(IsNull(Master!L_Model), "", Master!L_Model)
    Txt(LSaleDate) = IIf(IsNull(Master!L_Sale_Date), "", Master!L_Sale_Date)
    Txt(NextDate) = IIf(IsNull(Master!NextDateTime), "", Master!NextDateTime)
    
    Txt(LastDate) = IIf(IsNull(Master!LastDate), "", Master!LastDate)
    If Not IsNull(Master!Call_Status) Then
        Select Case Master!Call_Status
            Case 0
               Txt(CallStat) = "Cold"
            Case 1
               Txt(CallStat) = "Warm"
            Case 2
                Txt(CallStat) = "Hot"
        End Select
    End If
    Txt(CloseYN) = IIf(Master!Close_lost = 1, "Yes", "No")
    Txt(Lbl) = IIf(IsNull(Master!Label), "", Master!Label)
    Txt(ActiveYN) = IIf(Master!Active_YN = 1, "Yes", "No")
    Txt(Model) = "" 'Applicable only in add.
Else
    Call BlankText
End If
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Ini_Grid()
DGCity.left = Me.width - (DGCity.width + mRtScale): DGCity.top = mTopScale: DGCity.height = 4935
DGRep.left = Me.width - (DGRep.width + mRtScale): DGRep.top = mTopScale
DGProf.left = Me.width - (DGProf.width + mRtScale): DGProf.top = mTopScale
DGRef.left = Me.width - (DGRef.width + mRtScale): DGRef.top = mTopScale
DGArea.left = Me.width - (DGArea.width + mRtScale): DGArea.top = mTopScale
DGMod.left = Me.left: DGMod.width = Me.width - mRtScale
DGMod.top = Txt(Model).top + Txt(Model).height: DGMod.height = Me.height - (DGMod.top + mTopScale)

End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
    Txt(I).BackColor = CtrlBColOrg
    Txt(I).ForeColor = CtrlFColOrg
Next
Txt(LastDate).Enabled = False
Txt(CallStat).Enabled = False
Txt(CloseYN).Enabled = False

txtDisabled_Color Me
End Sub

Private Sub Grid_Hide()
    If DGCity.Visible = True Then DGCity.Visible = False
    If DGRep.Visible = True Then DGRep.Visible = False
    If DGRef.Visible = True Then DGRef.Visible = False
    If DGProf.Visible = True Then DGProf.Visible = False
    If DGArea.Visible = True Then DGArea.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Sub RemoveTxtNull()
Dim I As Integer
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = IIf(IsNull(Txt(I).TEXT), "", Txt(I).TEXT)
Next I
End Sub

