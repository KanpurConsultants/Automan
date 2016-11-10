VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FaVtype 
   BackColor       =   &H00CAF1FD&
   Caption         =   "Define Voucher Type"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11400
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   11400
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   15
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   13
      Text            =   "No"
      Top             =   3015
      Width           =   480
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   13
      Left            =   7890
      MaxLength       =   30
      TabIndex        =   15
      Top             =   915
      Width           =   3360
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   12
      Left            =   7890
      MaxLength       =   30
      TabIndex        =   14
      Top             =   705
      Width           =   3360
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   14
      Left            =   7890
      MaxLength       =   3
      TabIndex        =   16
      Top             =   1125
      Width           =   480
   End
   Begin MSDataGridLib.DataGrid DGGroup 
      Height          =   3330
      Left            =   6825
      Negotiate       =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   6810
      Visible         =   0   'False
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12176853
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3150
      Left            =   135
      TabIndex        =   18
      Top             =   3510
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   5556
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Prefix Details"
      TabPicture(0)   =   "FaVtype.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Must Have A/C Groups"
      TabPicture(1)   =   "FaVtype.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Don't Show A/C Groups"
      TabPicture(2)   =   "FaVtype.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   2565
         Left            =   195
         TabIndex        =   39
         Top             =   450
         Width           =   9570
         Begin VB.TextBox TxtGrid 
            Appearance      =   0  'Flat
            BackColor       =   &H00FDF4B5&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000012&
            Height          =   240
            Index           =   0
            Left            =   1725
            TabIndex        =   41
            Top             =   465
            Visible         =   0   'False
            Width           =   1485
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
            Height          =   2400
            Left            =   15
            TabIndex        =   17
            Top             =   120
            Width           =   9510
            _ExtentX        =   16775
            _ExtentY        =   4233
            _Version        =   393216
            BackColor       =   13166810
            ForeColor       =   0
            BackColorFixed  =   12632319
            ForeColorFixed  =   128
            BackColorSel    =   15718112
            ForeColorSel    =   12582912
            BackColorBkg    =   12243913
            GridColor       =   12632319
            GridColorFixed  =   33023
            FocusRect       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
            RowSizingMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2550
         Left            =   -74775
         TabIndex        =   38
         Top             =   420
         Width           =   9570
         Begin VB.TextBox TxtGrid 
            Appearance      =   0  'Flat
            BackColor       =   &H00FDF4B5&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000012&
            Height          =   240
            Index           =   1
            Left            =   1695
            TabIndex        =   43
            Top             =   750
            Visible         =   0   'False
            Width           =   1485
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
            Height          =   2385
            Left            =   15
            TabIndex        =   40
            Top             =   150
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   4207
            _Version        =   393216
            BackColor       =   13166810
            ForeColor       =   0
            BackColorFixed  =   12632319
            ForeColorFixed  =   128
            BackColorSel    =   15718112
            ForeColorSel    =   12582912
            BackColorBkg    =   12243913
            GridColor       =   12632319
            GridColorFixed  =   33023
            FocusRect       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
            RowSizingMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2670
         Left            =   -74790
         TabIndex        =   36
         Top             =   375
         Width           =   9555
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid2 
            Height          =   2460
            Left            =   30
            TabIndex        =   37
            Top             =   150
            Width           =   9510
            _ExtentX        =   16775
            _ExtentY        =   4339
            _Version        =   393216
            BackColor       =   13166810
            ForeColor       =   0
            BackColorFixed  =   12632319
            ForeColorFixed  =   128
            BackColorSel    =   15718112
            ForeColorSel    =   12582912
            BackColorBkg    =   12243913
            GridColor       =   12632319
            GridColorFixed  =   33023
            FocusRect       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
            RowSizingMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   270
      TabIndex        =   34
      Top             =   6570
      Visible         =   0   'False
      Width           =   2010
      Begin MSComctlLib.ListView ListView 
         Height          =   1815
         Left            =   15
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   45
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   3201
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
         BackColor       =   12176853
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
      Height          =   195
      Index           =   11
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   12
      Text            =   "No"
      Top             =   2805
      Width           =   480
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   2
      Left            =   1710
      MaxLength       =   5
      TabIndex        =   1
      Top             =   495
      Width           =   990
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   8
      Left            =   1710
      MaxLength       =   150
      TabIndex        =   9
      Top             =   2175
      Width           =   5640
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   10
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "No"
      Top             =   2595
      Width           =   480
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   9
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "No"
      Top             =   2385
      Width           =   480
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   1
      Left            =   1710
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1335
      Width           =   1005
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "No"
      Top             =   1965
      Width           =   480
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   3
      Left            =   1710
      MaxLength       =   30
      TabIndex        =   2
      Top             =   705
      Width           =   3360
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   0
      Left            =   1710
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1125
      Width           =   2025
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   4
      Left            =   1710
      MaxLength       =   10
      TabIndex        =   3
      Top             =   915
      Width           =   2025
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   5
      Left            =   1710
      MaxLength       =   9
      TabIndex        =   6
      Text            =   "Automatic"
      Top             =   1545
      Width           =   945
   End
   Begin VB.TextBox Txt 
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
      Height          =   195
      Index           =   6
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "Yes"
      Top             =   1755
      Width           =   480
   End
   Begin MSDataGridLib.DataGrid DGNCat 
      Height          =   3330
      Left            =   2490
      Negotiate       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7380
      Visible         =   0   'False
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12176853
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Caption         =   "NCat"
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
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGVType 
      Height          =   3330
      Left            =   2505
      Negotiate       =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7050
      Visible         =   0   'False
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12176853
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Caption         =   "VType"
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
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCategory 
      Height          =   3330
      Left            =   2490
      Negotiate       =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6780
      Visible         =   0   'False
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12176853
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Caption         =   "Category"
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
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      tAdd            =   0   'False
   End
   Begin MSDataGridLib.DataGrid DGDR 
      Height          =   3330
      Left            =   6615
      Negotiate       =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2895
      Visible         =   0   'False
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12176853
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Caption         =   "Debit A/c"
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
            ColumnWidth     =   4919.811
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCR 
      Height          =   3330
      Left            =   5760
      Negotiate       =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   3090
      Visible         =   0   'False
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12176853
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Caption         =   "Credit A/c"
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
            ColumnWidth     =   4919.811
         EndProperty
      EndProperty
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Wise Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   22
      Left            =   105
      TabIndex        =   57
      Top             =   3015
      Width           =   1530
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Y)es / (N)o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   21
      Left            =   2265
      TabIndex        =   56
      Top             =   3015
      Width           =   825
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Dr/Cr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   20
      Left            =   6780
      TabIndex        =   53
      Top             =   1125
      Width           =   900
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(D)r / (C)r"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   19
      Left            =   8475
      TabIndex        =   52
      Top             =   1125
      Width           =   660
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Default Debit A/C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   18
      Left            =   6150
      TabIndex        =   51
      Top             =   915
      Width           =   1530
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Default Credit A/C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   17
      Left            =   6105
      TabIndex        =   50
      Top             =   705
      Width           =   1575
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Y)es / (N)o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   16
      Left            =   2280
      TabIndex        =   49
      Top             =   2805
      Width           =   825
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Y)es / (N)o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   15
      Left            =   2280
      TabIndex        =   48
      Top             =   2595
      Width           =   825
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Y)es / (N)o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   14
      Left            =   2280
      TabIndex        =   47
      Top             =   2385
      Width           =   825
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Y)es / (N)o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   13
      Left            =   2280
      TabIndex        =   46
      Top             =   1965
      Width           =   825
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Y)es / (N)o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   12
      Left            =   2280
      TabIndex        =   45
      Top             =   1755
      Width           =   825
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(M)anual / (A)utomatic / (S)emiauto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   8
      Left            =   2730
      TabIndex        =   44
      Top             =   1545
      Width           =   2490
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clg Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   11
      Left            =   900
      TabIndex        =   33
      Top             =   2805
      Width           =   750
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   1215
      TabIndex        =   32
      Top             =   495
      Width           =   435
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1065
      TabIndex        =   31
      Top             =   1335
      Width           =   585
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   885
      TabIndex        =   30
      Top             =   1125
      Width           =   765
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   675
      TabIndex        =   29
      Top             =   705
      Width           =   975
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numbering Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   5
      Left            =   45
      TabIndex        =   28
      Top             =   1545
      Width           =   1605
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chq Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   10
      Left            =   840
      TabIndex        =   27
      Top             =   2595
      Width           =   810
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Short Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   645
      TabIndex        =   26
      Top             =   915
      Width           =   1005
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seperate Narration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   6
      Left            =   30
      TabIndex        =   25
      Top             =   1755
      Width           =   1620
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Common Narration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   7
      Left            =   90
      TabIndex        =   24
      Top             =   1965
      Width           =   1560
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chq No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   9
      Left            =   945
      TabIndex        =   23
      Top             =   2385
      Width           =   705
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "FaVtype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BackColorSelLeave As String
Dim RsCateg As ADODB.Recordset, RsVType As ADODB.Recordset, RsNCat As ADODB.Recordset, RsAcGroup As ADODB.Recordset, RsEnv As ADODB.Recordset
Dim Master As ADODB.Recordset, RsDr As ADODB.Recordset, RsCr As ADODB.Recordset
Dim TmpSQL1 As String, moldVType As String, TempMsg As Byte
Dim GridKey As Integer, TAddMode As Boolean, ListArray As Variant, mListItem As ListItem
Private Const Categ As Byte = 0, NCat As Byte = 1, VType As Byte = 2, Desc As Byte = 3
Private Const ShortName As Byte = 4, VouNumMethod As Byte = 5, SepNarr As Byte = 6
Private Const CommNarr As Byte = 7, Narration As Byte = 8, ChqNo As Byte = 9
Private Const ChqDate As Byte = 10, ClgDate As Byte = 11, DefaultCrAC As Byte = 12, DefaultDrAC As Byte = 13, FirstDrCr As Byte = 14, SiteBaseNumber As Byte = 15

Private Const FSNo As Byte = 0, FFromDate As Byte = 1, FToDate As Byte = 2, FPrefix As Byte = 3, FStartSrNo As Byte = 4, FVType As Byte = 5
Private Const FSno1 As Byte = 0, FGroupName1 As Byte = 1, FDr1 As Byte = 2, FCr1 As Byte = 3, FGroupCode1 As Byte = 4
Private Const FSno2 As Byte = 0, FGroupName2 As Byte = 1, FDr2 As Byte = 2, FCr2 As Byte = 3, FGroupCode2 As Byte = 4
Private PubDatamanFa As New DMFa.ClsFa

Private Sub DGCategory_Click()
    DgCategory.Visible = False
    If RsCateg.RecordCount > 0 Then
        txt(Categ).Tag = RsCateg!Code
        txt(Categ).TEXT = RsCateg!Name
    End If
    txt(Categ).SetFocus
End Sub
Private Sub DGNCat_Click()
    DGNCat.Visible = False
    If RsNCat.RecordCount > 0 Then
        txt(NCat).Tag = RsNCat!Code
        txt(NCat).TEXT = RsNCat!Name
    End If
    txt(NCat).SetFocus
End Sub
Private Sub DGVType_Click()
    DGVType.Visible = False
    If RsVType.RecordCount > 0 Then
        txt(VType).Tag = RsVType!Code
        txt(VType).TEXT = RsVType!Name
    End If
    txt(VType).SetFocus
End Sub
Private Sub DGGroup_Click()
    DGGroup.Visible = False
    If RsAcGroup.RecordCount > 0 Then
        txtgrid(1).TEXT = RsAcGroup!Name
        FGrid1.TextMatrix(FGrid.Row, FGroupCode1) = RsAcGroup!Code
        FGrid1.TextMatrix(FGrid.Row, FGroupName1) = RsAcGroup!Name
    End If
    TxtGridLeave
    If txtgrid(1).Visible = True Then txtgrid(1).SetFocus
End Sub
Private Sub FGrid1_Click()
If TopCtrl1.TopText2 = "Browse" Then Exit Sub
If FGrid1.Col = FDr1 Or FGrid1.Col = FCr1 Then
    If FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = "" Then
        FGrid1.Col = FGrid1.Col
        FGrid1.Row = FGrid1.Row
        FGrid1.CellFontName = "wingdings"
        FGrid1.CellFontSize = 18
        FGrid1.CellForeColor = vbRed
        FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = "ü"
    Else
        FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
    End If
Else
   txtgrid(1).Visible = False
End If
End Sub
Private Sub FGrid2_Click()
If TopCtrl1.TopText2 = "Browse" Then Exit Sub
If FGrid2.Col = FDr2 Or FGrid2.Col = FCr2 Then
    If FGrid2.TextMatrix(FGrid2.Row, FGrid2.Col) = "" Then
        FGrid2.Col = FGrid2.Col
        FGrid2.Row = FGrid2.Row
        FGrid2.CellFontName = "wingdings"
        FGrid2.CellFontSize = 18
        FGrid2.CellForeColor = vbRed
        FGrid2.TextMatrix(FGrid2.Row, FGrid2.Col) = "ü"
    Else
        FGrid2.TextMatrix(FGrid2.Row, FGrid2.Col) = ""
    End If
End If
End Sub
Private Sub FGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   If FGrid2.Col = FDr2 Or FGrid2.Col = FCr2 Then
        If FGrid2.TextMatrix(FGrid2.Row, FGrid2.Col) = "" Then
            FGrid2.Col = FGrid2.Col
            FGrid2.Row = FGrid2.Row
            FGrid2.CellFontName = "wingdings"
            FGrid2.CellFontSize = 18
            FGrid2.CellForeColor = vbRed
            FGrid2.TextMatrix(FGrid2.Row, FGrid2.Col) = "ü"
        Else
            FGrid2.TextMatrix(FGrid2.Row, FGrid2.Col) = ""
        End If
    End If
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FaFormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:     If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Form_Load()
On Error GoTo ELoop
    TopCtrl1.Tag = "AEDP": TopCtrl1.TopText1 = Me.CAPTION
    TopCtrl1.Tag = PubUParam
'    If PubSec = "SANJEEV" Then
'        If rsUserPerm.RecordCount > 0 Then
'            rsUserPerm.MoveFirst
'            rsUserPerm.FIND ("FORM_NAME='" & Me.CAPTION & "'")
'            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
'        End If
'    ElseIf PubSec = "RAHUL" Then
'        If rsUserPerm.RecordCount > 0 Then
'            rsUserPerm.MoveFirst
'            rsUserPerm.FIND ("FORM_CODE='" & Me.Name & "'")
'            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
'        End If
'    End If
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
    Set RsCateg = New ADODB.Recordset
    RsCateg.CursorLocation = adUseClient
    RsCateg.Open "Select DISTINCT Category As Code,Category as Name From VoucherCat Order by Category", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DgCategory.DataSource = RsCateg
    
    Set RsVType = New ADODB.Recordset
    RsVType.CursorLocation = adUseClient
    RsVType.Open "Select Category As Code,V_Type as Name ,Category From Voucher_Type order by V_type", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGVType.DataSource = RsVType
    
    Set RsNCat = New ADODB.Recordset
    RsNCat.CursorLocation = adUseClient
    RsNCat.Open "Select Category as Code,NCat as Name ,Category From VoucherCat Order by Category", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGNCat.DataSource = RsNCat
    
    Set RsEnv = New ADODB.Recordset
    RsEnv.CursorLocation = adUseClient
    RsEnv.Open "Select ShowGroup,DonotShowGroup From FaEnviro ", G_FaCn, adOpenDynamic, adLockOptimistic
    
    Set RsCr = New ADODB.Recordset
    RsCr.CursorLocation = adUseClient
    RsCr.Open "Select SubCode As Code,Name From SubGroup Order by Name", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGCR.DataSource = RsCr

    Set RsDr = New ADODB.Recordset
    RsDr.CursorLocation = adUseClient
    RsDr.Open "Select SubCode As Code,Name From SubGroup Order by Name", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGDR.DataSource = RsDr

    Set RsAcGroup = New ADODB.Recordset
    RsAcGroup.CursorLocation = adUseClient
    If RsEnv!ShowGroup = "Yes" Then
        RsAcGroup.Open "Select GroupCode as Code,GroupName as Name From AcGroup Where SysGroup = 'Y' And AliasYN <> 'Y' Order by GroupName", G_FaCn, adOpenDynamic, adLockOptimistic
    Else
        RsAcGroup.Open "Select GroupCode as Code,GroupName as Name From AcGroup Where AliasYN <> 'Y' Order by GroupName", G_FaCn, adOpenDynamic, adLockOptimistic
    End If
    Set DGGroup.DataSource = RsAcGroup

    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "Select V.V_Type as SearchCode,V.* From Voucher_Type V Order by V.V_type", G_FaCn, adOpenDynamic, adLockOptimistic
    
    Disp_Text SETS("INI", Me, Master)
    Ini_Grid
    FaWinSetting Me
    Me.height = 7590
    Me.width = 11940
    MoveRec
    Exit Sub
ELoop:      MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form_Unload (-1)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set RsCateg = Nothing
    Set Master = Nothing
    Set RsNCat = Nothing
    Set RsVType = Nothing
    Set RsCr = Nothing
    Set RsDr = Nothing
    Set PubDatamanFa = Nothing
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    Disp_Text SETS("ADD", Me, Master)
    BlankText
    txt(VType).SetFocus
    txt(VouNumMethod) = "Manual"
    txt(SepNarr) = "Yes"
    txt(CommNarr) = "Yes"
    txt(ChqNo) = "Yes"
    txt(ChqDate) = "Yes"
    txt(SiteBaseNumber) = "Yes"
    txt(ClgDate) = "Yes"
    FGrid.TextMatrix(1, FStartSrNo) = 1
    Call BlankFGrid1
    Call FillFGrid2
Exit Sub
ELoop: If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eDel()
Dim XBM, j As Byte, TmpSQL As String, Rst As ADODB.Recordset
On Error GoTo ELoop
    If VoucherTypeCheck(Master!V_Type) = True Then MsgBox "Transactions Exist For this Voucher Type,Can't Delete it", vbCritical, "Voucher Type Validation": Exit Sub
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        G_FaCn.BeginTrans
        XBM = Master.Bookmark
        G_FaCn.Execute "Delete From Voucher_Prefix Where V_Type='" & Master!SearchCode & "'"
        G_FaCn.Execute "Delete From Voucher_Include Where V_Type='" & Master!SearchCode & "'"
        G_FaCn.Execute "Delete From Voucher_Exclude Where V_Type='" & Master!SearchCode & "'"
        G_FaCn.Execute "Delete From Voucher_Type Where V_Type='" & Master!SearchCode & "'"
        G_FaCn.CommitTrans
        Master.Requery
        If Master.RecordCount >= XBM Then
            Master.Bookmark = XBM
        Else
            If Master.EOF = False Then Master.MoveLast
        End If
        MoveRec
        BUTTONS True, Me, Master, 0
    End If
    Set Rst = Nothing
Exit Sub
ELoop:      G_FaCn.RollbackTrans
            MsgBox err.Description, vbCritical, " Deletion Message"
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
    moldVType = txt(VType).TEXT
    FGrid.AddItem FGrid.Rows
    txt(Categ).Enabled = False
    txt(VType).Enabled = False
    txt(NCat).Enabled = False
    txt(Desc).SetFocus
    Call FillFGrid2
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
GSQL = "Select V.V_Type as SearchCode,V.Category,V.NCat,V.Description,V.V_Type,V.Short_Name,V.Number_Method,V.Separate_Narr,V.Common_Narr,V.Narration,V.ChqNo,V.ChqDt,V.ClgDt From Voucher_Type V Order by V.V_Type"
Set SearchForm = Me
FAFind.Show vbModal
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    Master.MoveFirst
    Master.FIND ("SearchCode='" & MyValue & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
    End If
Exit Sub
ELoop:      If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_eRef()
    RsCateg.Requery
    RsNCat.Requery
    RsVType.Requery
End Sub
Private Sub TopCtrl1_eSave()
Dim Rst As ADODB.Recordset, mTrans As Boolean, SearchCode As String, I As Integer, j As Integer, Count As Integer, X As String, Y As String
Dim SepNarr1 As String, CommNarr1 As String, ChqNo1 As String, ChqDt1 As String, ClgDt1 As String
On Error GoTo ELoop
    If FaIsValid(txt(Categ), "Category") = False Then Exit Sub
    If FaIsValid(txt(VType), "Voucher Type") = False Then Exit Sub
    If FaIsValid(txt(Desc), "Description") = False Then Exit Sub
    If FGrid.TextMatrix(1, FVType) = "" Then
        MsgBox "Please Fill Voucher Type Detail", vbInformation, Me.CAPTION: FGrid.Row = 1: FGrid.SetFocus: Exit Sub
    End If
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, FVType) <> "" Then
            If FGrid.TextMatrix(1, FFromDate) = "" Then MsgBox "From Date is a Required Field": FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(1, FToDate) = "" Then MsgBox "To Date is a Required Field": FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(1, FPrefix) = "" Then MsgBox "Prefix is a Required Field": FGrid.SetFocus: Exit Sub
        ElseIf FGrid.TextMatrix(I, FVType) = "" Then
            FGrid.RemoveItem (I)
        End If
    Next
    For I = 1 To FGrid.Rows - 1
        X = Trim(CStr(FGrid.TextMatrix(I, FPrefix)))
        Count = 0
        For j = I + 1 To FGrid.Rows - 1
            Y = Trim(CStr(FGrid.TextMatrix(j, FPrefix)))
            If X = Y Then Count = Count + 1
            If Count > 1 Then
                MsgBox "Duplicate Voucher Prefix ", vbInformation, "Grid Validation"
                FGrid.SetFocus
                Exit Sub
            End If
        Next
    Next
    SepNarr1 = IIf(txt(SepNarr) = "Yes", "Y", "N")
    CommNarr1 = IIf(txt(CommNarr) = "Yes", "Y", "N")
    ChqNo1 = IIf(txt(ChqNo) = "Yes", "Y", "N")
    ChqDt1 = IIf(txt(ChqDate) = "Yes", "Y", "N")
    ClgDt1 = IIf(txt(ClgDate) = "Yes", "Y", "N")
    G_FaCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If PubFaSiteType = 1 And PubSeparateVrNoForSite = 1 Then
            G_FaCn.Execute ("Delete From Voucher_Prefix Where V_Type='" & txt(VType) & "' AND SITE_CODE='" & PubSeparateLogSite & "'")
        ElseIf PubFaSiteType = 2 Then
            G_FaCn.Execute ("Delete From Voucher_Prefix Where V_Type='" & txt(VType) & "' AND SITE_CODE='" & PubSiteCode & "'")
        Else
            G_FaCn.Execute ("Delete From Voucher_Prefix Where V_Type='" & txt(VType) & "' And (Div_Code='" & PubDivCode & "' Or IsNull(Div_Code,'')='') And (Site_Code='" & PubSiteCode & "' Or IsNull(Site_Code,'')='') ")
        End If
        G_FaCn.Execute "Delete From Voucher_Include Where V_Type='" & txt(VType) & "'"
        G_FaCn.Execute "Delete From Voucher_Exclude Where V_Type='" & txt(VType) & "'"
        G_FaCn.Execute "Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt, Print_Vno,U_Name,U_EntDt,U_AE,DefaultCrAC,DefaultDrAC,FirstDrCr, SiteBaseNumber) Values ('" & txt(Categ) & "','" & txt(NCat) & "','" & txt(VType) & "','" & txt(Desc) & "', '" & txt(Desc) & "','" & txt(ShortName) & "','" & txt(VouNumMethod) & "','" & SepNarr1 & "','" & CommNarr1 & "','" & txt(Narration) & "','" & ChqNo1 & "','" & ChqDt1 & "','" & ClgDt1 & "', 1,'" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'," & FaChk_Text(txt(DefaultCrAC).Tag) & "," & FaChk_Text(txt(DefaultDrAC).Tag) & "," & FaChk_Text(txt(FirstDrCr)) & ", " & FaChk_Text(IIf(txt(SiteBaseNumber) = "Yes", "Y", "N")) & ")"
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, 1) <> "" And FGrid.TextMatrix(I, 2) <> "" And FGrid.TextMatrix(I, 3) <> "" Then
                If PubFaSiteType = 1 And PubSeparateVrNoForSite = 1 Then
                    G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,SITE_CODE, Div_Code) Values ('" & FGrid.TextMatrix(I, FVType) & "'," & FaConvertDate(FGrid.TextMatrix(I, FFromDate)) & "," & FaConvertDate(FGrid.TextMatrix(I, FToDate)) & ",'" & FGrid.TextMatrix(I, FPrefix) & "'," & Val(FGrid.TextMatrix(I, FStartSrNo)) & ",'" & PubSeparateLogSite & "', '" & PubDivCode & "')")
                ElseIf PubFaSiteType = 2 Then
                    G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,SITE_CODE, Div_Code) Values ('" & FGrid.TextMatrix(I, FVType) & "'," & FaConvertDate(FGrid.TextMatrix(I, FFromDate)) & "," & FaConvertDate(FGrid.TextMatrix(I, FToDate)) & ",'" & FGrid.TextMatrix(I, FPrefix) & "'," & Val(FGrid.TextMatrix(I, FStartSrNo)) & ",'" & PubSiteCode & "', '" & PubDivCode & "')")
                Else
                    G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No, Div_Code, Site_Code) Values ('" & FGrid.TextMatrix(I, FVType) & "'," & FaConvertDate(FGrid.TextMatrix(I, FFromDate)) & "," & FaConvertDate(FGrid.TextMatrix(I, FToDate)) & ",'" & FGrid.TextMatrix(I, FPrefix) & "'," & Val(FGrid.TextMatrix(I, FStartSrNo)) & ", '" & PubDivCode & "', '" & PubSiteCode & "')")
                End If
            End If
        Next
        '''////entry in voucher include
        For I = 1 To FGrid1.Rows - 1
            If FGrid1.TextMatrix(I, 1) <> "" And (FGrid1.TextMatrix(I, 2) <> "" Or FGrid1.TextMatrix(I, 3) <> "") Then
                G_FaCn.Execute ("Insert Into Voucher_include(V_Type,GroupCode,DR,CR) Values ('" & txt(VType) & "','" & FGrid1.TextMatrix(I, FGroupCode1) & "','" & IIf(FGrid1.TextMatrix(I, FDr1) = "ü", "Y", "N") & "','" & IIf(FGrid1.TextMatrix(I, FCr1) = "ü", "Y", "N") & "')")
            End If
        Next
        '''////entry in voucher exclude
        For I = 1 To FGrid2.Rows - 1
            If FGrid2.TextMatrix(I, 1) <> "" And (FGrid2.TextMatrix(I, 2) <> "" Or FGrid2.TextMatrix(I, 3) <> "") Then
                G_FaCn.Execute ("Insert Into Voucher_Exclude(V_Type,GroupCode,DR,CR) Values ('" & txt(VType) & "','" & FGrid2.TextMatrix(I, FGroupCode2) & "','" & IIf(FGrid2.TextMatrix(I, FDr2) = "ü", "Y", "N") & "','" & IIf(FGrid2.TextMatrix(I, FCr2) = "ü", "Y", "N") & "')")
            End If
        Next
    Else
        If (PubFaSiteType = 1 And PubSeparateVrNoForSite = 1) Then
            G_FaCn.Execute ("Delete From Voucher_Prefix Where V_Type='" & txt(VType) & "' AND SITE_CODE='" & PubSeparateLogSite & "'")
        ElseIf PubFaSiteType = 2 Then
            G_FaCn.Execute ("Delete From Voucher_Prefix Where V_Type='" & txt(VType) & "' AND SITE_CODE='" & PubSiteCode & "'")
        Else
            G_FaCn.Execute ("Delete From Voucher_Prefix Where V_Type='" & txt(VType) & "'  And (Div_Code='" & PubDivCode & "' Or IsNull(Div_Code,'')='')  And (Site_Code='" & PubSiteCode & "' Or IsNull(Site_Code,'')='') ")
        End If
        G_FaCn.Execute "Delete From Voucher_Include Where V_Type='" & txt(VType) & "'"
        G_FaCn.Execute "Delete From Voucher_Exclude Where V_Type='" & txt(VType) & "'"
        G_FaCn.Execute "Update Voucher_Type Set DefaultCrAC=" & FaChk_Text(txt(DefaultCrAC).Tag) & ",DefaultDrAC=" & FaChk_Text(txt(DefaultDrAC).Tag) & ",FirstDrCr=" & FaChk_Text(txt(FirstDrCr)) & ",Description='" & txt(Desc) & "',Short_Name='" & txt(ShortName) & "',Number_Method='" & txt(VouNumMethod) & "',Separate_Narr='" & SepNarr1 & "',Common_Narr='" & CommNarr1 & "',Narration='" & txt(Narration) & "',ChqNo='" & ChqNo1 & "',ChqDt='" & ChqDt1 & "', ClgDt='" & ClgDt1 & "', SiteBaseNumber=" & FaChk_Text(IIf(txt(SiteBaseNumber) = "Yes", "Y", "N")) & ",U_Name='" & pubUName & "',U_EntDt=" & FaConvertDate(PubLoginDate) & ", U_AE='E' Where V_Type='" & txt(VType) & "'"
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, FVType) <> "" And FGrid.TextMatrix(I, FFromDate) <> "" And FGrid.TextMatrix(I, FPrefix) <> "" Then
                If (PubFaSiteType = 1 And PubSeparateVrNoForSite = 1) Then
                    G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,SITE_CODE) Values ('" & FGrid.TextMatrix(I, FVType) & "'," & FaConvertDate(FGrid.TextMatrix(I, FFromDate)) & "," & FaConvertDate(FGrid.TextMatrix(I, FToDate)) & ",'" & FGrid.TextMatrix(I, FPrefix) & "'," & Val(FGrid.TextMatrix(I, FStartSrNo)) & ",'" & PubSeparateLogSite & "')")
                ElseIf PubFaSiteType = 2 Then
                    G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,SITE_CODE) Values ('" & FGrid.TextMatrix(I, FVType) & "'," & FaConvertDate(FGrid.TextMatrix(I, FFromDate)) & "," & FaConvertDate(FGrid.TextMatrix(I, FToDate)) & ",'" & FGrid.TextMatrix(I, FPrefix) & "'," & Val(FGrid.TextMatrix(I, FStartSrNo)) & ",'" & PubSiteCode & "')")
                Else
                    G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No, Div_Code,Site_Code) Values ('" & FGrid.TextMatrix(I, FVType) & "'," & FaConvertDate(FGrid.TextMatrix(I, FFromDate)) & "," & FaConvertDate(FGrid.TextMatrix(I, FToDate)) & ",'" & FGrid.TextMatrix(I, FPrefix) & "'," & Val(FGrid.TextMatrix(I, FStartSrNo)) & ", '" & PubDivCode & "', '" & PubSiteCode & "')")
                End If
            End If
        Next
         '''////entry in voucher include
        For I = 1 To FGrid1.Rows - 1
            If FGrid1.TextMatrix(I, 1) <> "" And (FGrid1.TextMatrix(I, 2) <> "" Or FGrid1.TextMatrix(I, 3) <> "") Then
                G_FaCn.Execute ("Insert Into Voucher_include(V_Type,GroupCode,DR,CR) Values ('" & txt(VType) & "','" & FGrid1.TextMatrix(I, FGroupCode1) & "','" & IIf(FGrid1.TextMatrix(I, FDr1) = "ü", "Y", "N") & "','" & IIf(FGrid1.TextMatrix(I, FCr1) = "ü", "Y", "N") & "')")
            End If
        Next
        '''////entry in voucher exclude
        For I = 1 To FGrid2.Rows - 1
            If FGrid2.TextMatrix(I, 1) <> "" And (FGrid2.TextMatrix(I, 2) <> "" Or FGrid2.TextMatrix(I, 3) <> "") Then
                G_FaCn.Execute ("Insert Into Voucher_Exclude(V_Type,GroupCode,DR,CR) Values ('" & txt(VType) & "','" & FGrid2.TextMatrix(I, FGroupCode2) & "','" & IIf(FGrid2.TextMatrix(I, FDr2) = "ü", "Y", "N") & "','" & IIf(FGrid2.TextMatrix(I, FCr2) = "ü", "Y", "N") & "')")
            End If
        Next
    End If
    G_FaCn.CommitTrans
    mTrans = False
    SearchCode = txt(VType)
    Master.Requery
    Master.FIND "SearchCode ='" & SearchCode & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Set Rst = Nothing
    Exit Sub
ELoop:      If mTrans = True Then G_FaCn.RollbackTrans
            If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
Exit Sub
End Sub
Private Sub Txt_GotFocus(Index As Integer)
    FaCtrl_GetFocus txt(Index)
    Grid_Hide
    Select Case Index
        Case Categ
            If RsCateg.RecordCount = 0 Or (RsCateg.EOF = True Or RsCateg.BOF = True) Or txt(Index) = "" Then Exit Sub
            If txt(Index).TEXT <> RsCateg!Name Then
                RsCateg.MoveFirst
                RsCateg.FIND "Name ='" & txt(Index).TEXT & "'"
            End If
        Case VType
            If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or txt(Index) = "" Then Exit Sub
            If txt(Index).TEXT <> RsVType!Name Then
                RsVType.MoveFirst
                RsVType.FIND "Name ='" & txt(Index).TEXT & "'"
            End If
        Case NCat
            RsNCat.Filter = adFilterNone
            RsNCat.Filter = "Category='" & txt(Categ) & "'"
            If RsNCat.RecordCount = 0 Or (RsNCat.EOF = True Or RsNCat.BOF = True) Or txt(Index) = "" Then Exit Sub
            If txt(Index).TEXT <> RsNCat!Name Then
                RsNCat.MoveFirst
                RsNCat.FIND "Name ='" & txt(Index).TEXT & "'"
            End If
        Case DefaultCrAC
            If RsCr.RecordCount = 0 Or (RsCr.EOF = True Or RsCr.BOF = True) Or txt(Index) = "" Then Exit Sub
            If txt(Index).TEXT <> RsCr!Name Then
                RsCr.MoveFirst
                RsCr.FIND "Name ='" & txt(Index).TEXT & "'"
            End If
        Case DefaultDrAC
            If RsDr.RecordCount = 0 Or (RsDr.EOF = True Or RsDr.BOF = True) Or txt(Index) = "" Then Exit Sub
            If txt(Index).TEXT <> RsDr!Name Then
                RsDr.MoveFirst
                RsDr.FIND "Name ='" & txt(Index).TEXT & "'"
            End If
    End Select
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case DefaultCrAC
            FaDGridTxtKeyDown DGCR, txt, Index, RsCr, KeyCode, False, 1
        Case DefaultDrAC
            FaDGridTxtKeyDown DGDR, txt, Index, RsDr, KeyCode, False, 1
        Case VType
            FaDGridTxtKeyDown_Mast DGVType, txt, Index, RsVType, KeyCode, False, 1
        Case Categ
            FaDGridTxtKeyDown DgCategory, txt, Index, RsCateg, KeyCode, False, 1
        Case NCat
            FaDGridTxtKeyDown DGNCat, txt, Index, RsNCat, KeyCode, False, 1
    End Select
    If DGDR.Visible = False And DGCR.Visible = False And DgCategory.Visible = False And DGNCat.Visible = False And DGVType.Visible = False And FrmList.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
            If txt(CommNarr) = "Yes" Then
                txt(Narration).Enabled = True
            Else
                txt(Narration).Enabled = False
                txt(Narration) = ""
            End If
            FaCtrl_DownKeyDown KeyCode, Shift
        End If
        If TopCtrl1.TopText2 = "Add" And Index <> Categ Then
            If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then FaCtrl_UpKeyDown KeyCode, Shift
        End If
        If TopCtrl1.TopText2 = "Edit" And Index <> Desc Then
            If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then FaCtrl_UpKeyDown KeyCode, Shift
        End If
    End If
End Sub
Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
    FaCheckQuote keyascii
    Select Case Index
        Case DefaultCrAC
            If DGCR.Visible = True Then FaDGridTxtKeyPress txt, Index, RsCr, keyascii, "Name"
        Case DefaultDrAC
            If DGDR.Visible = True Then FaDGridTxtKeyPress txt, Index, RsDr, keyascii, "Name"
        Case Categ
            If DgCategory.Visible = True Then FaDGridTxtKeyPress txt, Index, RsCateg, keyascii, "Name"
        Case NCat
            If DGNCat.Visible = True Then FaDGridTxtKeyPress txt, Index, RsNCat, keyascii, "Name"
        Case VouNumMethod
            If keyascii = 77 Or keyascii = 109 Then
                txt(Index) = "Manual"
                keyascii = 0
            ElseIf keyascii = 65 Or keyascii = 97 Then
                txt(Index) = "Automatic"
                keyascii = 0
            ElseIf keyascii = 83 Or keyascii = 115 Then
                txt(Index) = "SemiAuto"
                keyascii = 0
            Else
                keyascii = 0
            End If
        Case SepNarr, CommNarr, ChqNo, ChqDate, ClgDate, SiteBaseNumber
            If keyascii = 78 Or keyascii = 110 Then   'NO
                txt(Index) = "No"
                keyascii = 0
            ElseIf keyascii = 89 Or keyascii = 121 Then 'Yes
                txt(Index) = "Yes"
                keyascii = 0
            Else
                keyascii = 0
            End If
            If Index = CommNarr Then
                If txt(CommNarr) = "Yes" Then
                    txt(Narration).Enabled = True
                Else
                    txt(Narration).Enabled = False
                    txt(Narration) = ""
                End If
            End If
        Case FirstDrCr
            If Asc(UCase(Chr(keyascii))) = vbKeyD Then
                txt(FirstDrCr) = "Dr"
                keyascii = 0
            ElseIf Asc(UCase(Chr(keyascii))) = vbKeyC Then
                txt(FirstDrCr) = "Cr"
                keyascii = 0
            Else
                keyascii = 0
            End If
    End Select
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case VType
            If DGVType.Visible = True Then FaDGridTxtKeyUp_Mast txt, Index, RsVType, KeyCode, "Name"
        Case VouNumMethod
            If FrmList.Visible = True Then FaListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    End Select
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    FaCtrl_validate txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case DefaultCrAC
            If RsCr.RecordCount = 0 Or (RsCr.EOF = True Or RsCr.BOF = True) Or txt(Index) = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsCr!Name
            txt(Index).Tag = RsCr!Code
        End If
        Case DefaultDrAC
            If RsDr.RecordCount = 0 Or (RsDr.EOF = True Or RsDr.BOF = True) Or txt(Index) = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            Else
                txt(Index).TEXT = RsDr!Name
                txt(Index).Tag = RsDr!Code
            End If
        Case Categ
            If FaIsValid(txt(Categ), "Category") = False Then txt(Index).SetFocus: Cancel = True: Exit Sub
            If RsCateg.RecordCount = 0 Or (RsCateg.EOF = True Or RsCateg.BOF = True) Or txt(Index).TEXT = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            Else
                txt(Index).TEXT = RsCateg!Name
                txt(Index).Tag = RsCateg!Code
            End If
        Case NCat
            If FaIsValid(txt(NCat), "NCat") = False Then txt(Index).SetFocus: Cancel = True: Exit Sub
            If RsNCat.RecordCount = 0 Or (RsNCat.EOF = True Or RsNCat.BOF = True) Or txt(Index).TEXT = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            Else
                txt(Index).TEXT = RsNCat!Name
                txt(Index).Tag = RsNCat!Code
            End If
        Case VType
            If FaIsValid(txt(VType), "Voucher Type") = False Then txt(Index).SetFocus: Cancel = True: Exit Sub
            If Validate = True Then Cancel = True: Exit Sub
            FGrid.TextMatrix(1, FVType) = txt(VType)
    End Select
End Sub
Private Function Validate() As Boolean
    If RsVType.RecordCount = 0 Then Exit Function
    If TopCtrl1.TopText2 = "Add" Then
        If G_FaCn.Execute("Select Count(*) From Voucher_Type Where V_Type='" & txt(VType) & "'").Fields(0).Value > 0 Then MsgBox "Duplicate Voucher Type", vbInformation, "Information": txt(VType).SetFocus: Validate = True: Exit Function
    Else
        If G_FaCn.Execute("Select Count(*) From Voucher_Type Where V_Type='" & txt(VType) & "' And V_Type <>'" & moldVType & "'").Fields(0).Value > 0 Then MsgBox "Duplicate Voucher Type", vbInformation, "Information": txt(VType).SetFocus: Validate = True: Exit Function
    End If
End Function
Private Sub FGrid_Click()
    txtgrid(0).Visible = False
End Sub
Private Sub FGrid_DblClick()
    FGrid_KeyPress (vbKeyReturn)
    TAddMode = False
End Sub
Private Sub FGrid_GotFocus()
    If FGrid.BackColorSel = BackColorSelLeave Then FGrid.Col = 1
    FGrid.BackColorSel = FaBackColorSelEnter
    txtgrid(0).Visible = False
    Grid_Hide
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
        SendKeys "+{Tab}"
        KeyCode = 0
       ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 And FGrid.TextMatrix(1, FVType) <> "" Then
         If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave: Exit Sub
    ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
        SendKeysA vbKeyTab, True
    End If
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGrid.Col
            Case FPrefix, FFromDate, FToDate, FStartSrNo
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        End Select
    End If
    If KeyCode = 13 Then TAddMode = False
    KeyCode = 0
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FGrid_KeyPress(keyascii As Integer)
On Error GoTo ELoop
    Select Case FGrid.Col
        Case FPrefix, FFromDate, FToDate
            Call FaGet_Text(Me, FGrid, txtgrid, 0, False, keyascii)
        Case FStartSrNo
            Call FaGet_Text(Me, FGrid, txtgrid, 0, True, keyascii)
    End Select
    If keyascii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Dim I As Integer
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGrid.Row >= 1 Then
            If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                If FGrid.Rows > 2 Then
                    FGrid.RemoveItem (FGrid.Row)
                Else
                    FGrid.Rows = 1
                    FGrid.AddItem FGrid.Rows
                    FGrid.FixedRows = 1
                End If
                For I = 1 To FGrid.Rows - 1
                   FGrid.TextMatrix(I, 0) = I
                Next
            End If
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid.SetFocus
    End If
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FGrid_LostFocus()
    If txtgrid(0).Visible = False Then FGrid.BackColorSel = BackColorSelLeave
End Sub
Private Sub FGrid_Scroll()
    txtgrid(0).Visible = False
    Grid_Hide
End Sub
Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
Grid_Hide
Select Case Index
    Case 0
        txtgrid(Index).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Case 1
        txtgrid(Index).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
End Select
txtgrid(Index).MaxLength = 0
Select Case Index
Case 0
Case 1
    Select Case FGrid1.Col
        Case FGroupName1
            If RsAcGroup.RecordCount = 0 Or (RsAcGroup.EOF = True Or RsAcGroup.BOF = True) Or FGrid1.TextMatrix(FGrid1.Row, FGroupName1) = "" Then Exit Sub
            If FGrid1.TextMatrix(FGrid1.Row, FGroupName1) <> "" Then
                RsAcGroup.MoveFirst
                RsAcGroup.FIND "Code='" & FGrid1.TextMatrix(FGrid1.Row, FGroupCode1) & "'"
                If RsAcGroup.EOF = True Then RsAcGroup.MoveFirst
            End If
    End Select
End Select
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        txtgrid(Index).TEXT = txtgrid(Index).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        Select Case Index
            Case 0
                FGrid.SetFocus
            Case 1
                FGrid1.SetFocus
        End Select
        txtgrid(Index).Visible = False
        Exit Sub
    End If
    Select Case Index
    Case 0
        Select Case FGrid.Col
            Case FPrefix
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                         FaGridTxtDown FGrid, txtgrid, Index, KeyCode, TAddMode, FStartSrNo
                         FGrid.TextMatrix(1, FVType) = txt(VType)
                    End If
                End If
            Case FFromDate
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                         FaGridTxtDown FGrid, txtgrid, Index, KeyCode, TAddMode, FStartSrNo
                         FGrid.TextMatrix(1, FVType) = txt(VType)
                    End If
                End If
            Case FToDate
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                         FaGridTxtDown FGrid, txtgrid, Index, KeyCode, TAddMode, FStartSrNo
                         FGrid.TextMatrix(1, FVType) = txt(VType)
                    End If
                End If
            Case FStartSrNo
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave(0) = True Then
                         FaGridTxtDown FGrid, txtgrid, Index, KeyCode, TAddMode, FStartSrNo
                         FGrid.TextMatrix(1, FVType) = txt(VType)
                    End If
                End If
        End Select
    Case 1
        If FGrid1.Col = FGroupName1 Then
            FaDGridTxtKeyDown DGGroup, txtgrid, Index, RsAcGroup, KeyCode, True, 1
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave1 = True Then
                     FaGridTxtDown FGrid1, txtgrid, Index, KeyCode, TAddMode, FGroupName1
                End If
            End If
        End If
    End Select
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, keyascii As Integer)
On Error GoTo ELoop
FaCheckQuote keyascii
Select Case Index
Case 0
    Select Case FGrid.Col
        Case FStartSrNo
            FaNumPress txtgrid(0), keyascii, 8, 0
    End Select
Case 1
    Select Case FGrid1.Col
        Case FGroupName1
            If DGGroup.Visible = True Then FaDGridTxtKeyPress txtgrid, Index, RsAcGroup, keyascii, "Name"
    End Select
End Select
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case 0
        Select Case FGrid.Col
            Case FPrefix
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = txtgrid(0).TEXT
            Case FStartSrNo
                If txtgrid(0).TEXT = "" Then txtgrid(0).TEXT = 0
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = txtgrid(0).TEXT
        End Select
    Case 1
        Select Case FGrid1.Col
        Case FGroupName1
            If KeyCode <> 13 And DGGroup.Visible = False Then
                TxtGrid_KeyDown Index, GridKey, 0
                FaDGridTxtKeyPress txtgrid, Index, RsAcGroup, KeyCode, "Name", True
            End If
        End Select
End Select
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Select Case Index
Case 0
    Cancel = Not TxtGridLeave(Index, True)
Case 1
    Cancel = Not TxtGridLeave1(Index, True)
End Select
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim I As Integer
Select Case Index
Case 0
    Select Case FGrid.Col
        Case FPrefix
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = txtgrid(0).TEXT
            FGrid.TextMatrix(FGrid.Row, FVType) = txt(VType)
        Case FStartSrNo
            If txtgrid(0).TEXT = "" Then txtgrid(0).TEXT = 0
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = txtgrid(0).TEXT
        Case FFromDate
            If Len(Trim(txtgrid(0).TEXT)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, FFromDate) = PubLoginDate
                FGrid.TextMatrix(FGrid.Row, FPrefix) = Year(PubLoginDate)
            Else
                FGrid.TextMatrix(FGrid.Row, FFromDate) = RetDate(txtgrid(0))
                FGrid.TextMatrix(FGrid.Row, FPrefix) = Year(RetDate(txtgrid(0)))
                FGrid.TextMatrix(FGrid.Row, FVType) = txt(VType)
            End If
        Case FToDate
            If Len(Trim(txtgrid(0).TEXT)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, FToDate) = PubEndDate
            Else
                FGrid.TextMatrix(FGrid.Row, FToDate) = RetDate(txtgrid(0))
            End If
            FGrid.TextMatrix(FGrid.Row, FVType) = txt(VType)
    End Select
Case 1
End Select
    TxtGridLeave = True
    If ValidateCall = False Then
        FGrid.SetFocus
        txtgrid(0).Visible = False
    End If
End Function
Private Function TxtGridLeave1(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim I As Integer
    Select Case FGrid1.Col
        Case FGroupName1
            If RsAcGroup.RecordCount = 0 Or (RsAcGroup.EOF = True Or RsAcGroup.BOF = True) Or txtgrid(1).TEXT = "" Then
                FGrid1.TextMatrix(FGrid1.Row, FGroupName1) = ""
                FGrid1.TextMatrix(FGrid1.Row, FGroupCode1) = ""
            Else
                If ChkDuplicate1 = False Then TxtGridLeave1 = False: Exit Function
                FGrid1.TextMatrix(FGrid1.Row, FGroupName1) = RsAcGroup!Name
                FGrid1.TextMatrix(FGrid1.Row, FGroupCode1) = RsAcGroup!Code
            End If
    End Select
    TxtGridLeave1 = True
    If ValidateCall = False Then
        FGrid1.SetFocus
        txtgrid(1).Visible = False
    End If
End Function
Private Sub FGrid1_DblClick()
    FGrid1_KeyPress (vbKeyReturn)
    TAddMode = False
End Sub
Private Sub FGrid1_GotFocus()
    If FGrid1.BackColorSel = BackColorSelLeave Then FGrid1.Col = 1
    FGrid1.BackColorSel = FaBackColorSelEnter
    txtgrid(1).Visible = False
    Grid_Hide
End Sub
Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
        SendKeys "+{Tab}"
        KeyCode = 0
       ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 And FGrid1.TextMatrix(1, FGroupName1) <> "" Then
         If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave: Exit Sub
    ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
        SendKeysA vbKeyTab, True
    End If
    GridKey = KeyCode
    FGrid1.Tag = FGrid1.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGrid1.Col
            Case FGroupName1
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
                FGrid1.TextMatrix(FGrid1.Row, FGroupCode1) = ""
        End Select
    End If
    If KeyCode = 13 Then
        TAddMode = False
        If FGrid1.Col = FDr1 Or FGrid1.Col = FCr1 Then
            If FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = "" Then
                FGrid1.Col = FGrid1.Col
                FGrid1.Row = FGrid1.Row
                FGrid1.CellFontName = "wingdings"
                FGrid1.CellFontSize = 18
                FGrid1.CellForeColor = vbRed
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = "ü"
            Else
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
            End If
        End If
    End If
    KeyCode = 0
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FGrid1_KeyPress(keyascii As Integer)
On Error GoTo ELoop
    Select Case FGrid1.Col
        Case FGroupName1
            Call FaGet_Text(Me, FGrid1, txtgrid, 1, False, keyascii)
    End Select
    If keyascii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
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
                For I = 1 To FGrid1.Rows - 1
                   FGrid1.TextMatrix(I, 0) = I
                Next
            End If
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid1.SetFocus
    End If
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub FGrid1_LostFocus()
    If txtgrid(1).Visible = False Then FGrid1.BackColorSel = BackColorSelLeave
End Sub
Private Sub FGrid1_Scroll()
    txtgrid(1).Visible = False
    Grid_Hide
End Sub
Private Function ChkDuplicate() As Boolean
Dim I As Integer, X As String, Y As String, Col1 As Byte, Col2 As Byte
    Select Case FGrid.Col
        Case FPrefix
            Col1 = FPrefix
    End Select
    X = UCase(CStr(Trim(txtgrid(0))))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col1))))
        If X = Y And Y <> "" Then
            MsgBox "Duplicate Prefix Not Allowed", vbInformation, "Validation"
            txtgrid(0).SetFocus
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function
Private Function ChkDuplicate1() As Boolean
Dim I As Integer, X As String, Y As String
Dim Col1 As Byte
    Select Case FGrid1.Col
        Case FGroupName1
            Col1 = FGroupName1
    End Select
    X = UCase(CStr(Trim(txtgrid(1))))
    For I = 1 To FGrid1.Rows - 1
        If I = FGrid1.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid1.TextMatrix(I, Col1))))
        If X = Y And Y <> "" Then
            MsgBox "Duplicate Group Name Not Allowed", vbInformation, "Validation"
            txtgrid(1).SetFocus
            ChkDuplicate1 = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate1 = True
End Function
Private Sub BlankText()
Dim I As Byte
    moldVType = ""
    For I = 0 To txt.Count - 1
        txt(I).TEXT = ""
        txt(I).Tag = ""
    Next I
    txt(DefaultCrAC).Tag = ""
    txt(DefaultDrAC).Tag = ""
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End Sub
Private Sub MoveRec()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset, LedgRS As ADODB.Recordset, I As Integer, j As Integer
If Master.RecordCount > 0 Then
    txt(Categ) = Master!Category
    txt(NCat) = FaXNull(Master!NCat)
    txt(VType) = Master!V_Type
    txt(Desc) = Master!Description
    txt(ShortName) = FaXNull(Master!Short_Name)
    txt(VouNumMethod) = FaXNull(Master!Number_Method)
    If Not IsNull(Master!Separate_Narr) Then
        txt(SepNarr) = IIf(Master!Separate_Narr = "Y", "Yes", "No")
    Else
        txt(SepNarr) = ""
    End If
    If Not IsNull(Master!Common_Narr) Then
        txt(CommNarr) = IIf(Master!Common_Narr = "Y", "Yes", "No")
    Else
        txt(CommNarr) = ""
    End If
    txt(Narration) = FaXNull(Master!Narration)
    If Not IsNull(Master!ChqNo) Then
        txt(ChqNo) = IIf(Master!ChqNo = "Y", "Yes", "No")
    Else
        txt(CommNarr) = ""
    End If
    txt(SiteBaseNumber) = IIf(XNull(Master!SiteBaseNumber) = "Y", "Yes", "No")
    txt(ChqDate) = IIf(Master!ChqDT = "Y", "Yes", "No")
    txt(ClgDate) = IIf(Master!CLGDT = "Y", "Yes", "No")
    txt(FirstDrCr) = FaXNull(Master!FirstDrCr)
    FGrid.Redraw = False
    FGrid.Rows = 1
    I = 1
    txt(DefaultCrAC).Tag = FaXNull(Master!DefaultCrAC)
    txt(DefaultDrAC).Tag = FaXNull(Master!DefaultDrAC)
    Set Rst = G_FaCn.Execute("SELECT NAME FROM SUBGROUP WHERE SUBCODE='" & Master!DefaultCrAC & "'")
    If Rst.RecordCount > 0 Then
        txt(DefaultCrAC) = FaXNull(Rst!Name)
    Else
        txt(DefaultCrAC) = ""
    End If
    Set Rst = G_FaCn.Execute("SELECT NAME FROM SUBGROUP WHERE SUBCODE='" & Master!DefaultDrAC & "'")
    If Rst.RecordCount > 0 Then
        txt(DefaultDrAC) = FaXNull(Rst!Name)
    Else
        txt(DefaultDrAC) = ""
    End If
    If (PubFaSiteType = 1 And PubSeparateVrNoForSite = 1) Then
        Set Rst = G_FaCn.Execute("Select V_Type,Date_From,Date_to,Prefix,Start_Srl_No From Voucher_Prefix Where V_Type='" & Master!V_Type & "' AND SITE_CODE='" & PubSeparateLogSite & "' And Div_Code = '" & PubDivCode & "' Order By Date_From")
    ElseIf PubFaSiteType = 2 Then
        Set Rst = G_FaCn.Execute("Select V_Type,Date_From,Date_to,Prefix,Start_Srl_No From Voucher_Prefix Where V_Type='" & Master!V_Type & "' AND SITE_CODE='" & PubSiteCode & "' And Div_Code = '" & PubDivCode & "'  Order By Date_From")
    Else
        Set Rst = G_FaCn.Execute("Select V_Type,Date_From,Date_to,Prefix,Start_Srl_No From Voucher_Prefix Where V_Type='" & Master!V_Type & "' And (Div_Code = '" & PubDivCode & "' Or IsNull(Div_Code,'')='') And (Site_Code = '" & PubSiteCode & "' Or IsNull(Site_Code,'')='')   Order By Date_From")
    End If
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, FSNo) = I
                .TextMatrix(I, FPrefix) = FaXNull(Rst!Prefix)
                .TextMatrix(I, FFromDate) = FaXNull(Rst!Date_From)
                .TextMatrix(I, FToDate) = FaXNull(Rst!Date_to)
                .TextMatrix(I, FStartSrNo) = FaVNull(Rst!start_srl_no)
                .TextMatrix(I, FVType) = Rst!V_Type
            End With
            I = I + 1
            Rst.MoveNext
        Loop
        FGrid.FixedRows = 1
    End If
    FGrid.Redraw = True
    If I = 1 Then
        FGrid.Rows = 1
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    FGrid1.Redraw = False
    FGrid1.Rows = 1
    I = 1
    Set Rst = G_FaCn.Execute("Select V_Type,Voucher_Include.GroupCode,GroupName ,Dr,Cr From Voucher_Include Left Join AcGroup  on Voucher_Include.GroupCode = AcGroup.GroupCode Where AcGroup.AliasYN <>'Y' And V_Type='" & Master!V_Type & "' Order By Voucher_Include.GroupCode")
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            FGrid1.AddItem ""
            With FGrid1
                .TextMatrix(I, FSno1) = I
                .TextMatrix(I, FGroupName1) = Rst!GroupName
                If Not IsNull(Rst!dr) Then
                    .TextMatrix(I, FDr1) = IIf(Rst!dr = "Y", "ü", "")
                Else
                    .TextMatrix(I, FDr1) = ""
                End If
                If Not IsNull(Rst!cr) Then
                    .TextMatrix(I, FCr1) = IIf(Rst!cr = "Y", "ü", "")
                Else
                    .TextMatrix(I, FCr1) = ""
                End If
                .TextMatrix(I, FGroupCode1) = Rst!GroupCode
            End With
            I = I + 1
            Rst.MoveNext
        Loop
        FGrid1.FixedRows = 1
    End If
    FGrid1.Redraw = True
    If I = 1 Then
        FGrid1.Rows = 1
        FGrid1.AddItem FGrid1.Rows
        FGrid1.FixedRows = 1
    End If
    For j = 1 To FGrid1.Rows - 1
        FGrid1.Row = j
        FGrid1.Col = FDr2
        FGrid1.CellFontName = "wingdings"
        FGrid1.CellFontSize = 18
        FGrid1.CellForeColor = vbRed
        FGrid1.Col = FCr2
        FGrid1.CellFontName = "wingdings"
        FGrid1.CellFontSize = 18
        FGrid1.CellForeColor = vbRed
    Next
    FGrid2.Redraw = False
    FGrid2.Rows = 1
    I = 1
    Set Rst = G_FaCn.Execute("Select VE.V_Type,VE.GroupCode,A.GroupName ,VE.Dr,VE.Cr From Voucher_Exclude VE Left Join AcGroup A on VE.GroupCode = A.GroupCode Where A.AliasYN <>'Y' And VE.V_Type='" & Master!V_Type & "' Order By VE.GroupCode")
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            FGrid2.AddItem ""
            With FGrid2
                .TextMatrix(I, FSno2) = I
                .TextMatrix(I, FGroupName2) = Rst!GroupName
                If Not IsNull(Rst!dr) Then
                    .TextMatrix(I, FDr2) = IIf(Rst!dr = "Y", "ü", "")
                Else
                    .TextMatrix(I, FDr2) = ""
                End If
                If Not IsNull(Rst!cr) Then
                    .TextMatrix(I, FCr2) = IIf(Rst!cr = "Y", "ü", "")
                Else
                    .TextMatrix(I, FCr2) = ""
                End If
                .TextMatrix(I, FGroupCode2) = Rst!GroupCode
            End With
            I = I + 1
            Rst.MoveNext
        Loop
        FGrid2.FixedRows = 1
    End If
    FGrid2.Redraw = True
    If I = 1 Then
        FGrid2.Rows = 1
        FGrid2.AddItem FGrid2.Rows
        FGrid2.FixedRows = 1
    End If
    For j = 1 To FGrid2.Rows - 1
        FGrid2.Row = j
        FGrid2.Col = FDr2
        FGrid2.CellFontName = "wingdings"
        FGrid2.CellFontSize = 18
        FGrid2.CellForeColor = vbRed
        FGrid2.Col = FCr2
        FGrid2.CellFontName = "wingdings"
        FGrid2.CellFontSize = 18
        FGrid2.CellForeColor = vbRed
    Next
Else
    BlankText
End If
Grid_Hide
Set Rst = Nothing
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    For I = 0 To txt.Count - 1
        txt(I).Enabled = Enb
    Next
End Sub
Private Sub Grid_Hide()
    If DgCategory.Visible = True Then DgCategory.Visible = False
    If DGNCat.Visible = True Then DGNCat.Visible = False
    If DGVType.Visible = True Then DGVType.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGGroup.Visible = True Then DGGroup.Visible = False
    If DGCR.Visible = True Then DGCR.Visible = False
End Sub
Private Sub SaveMsg(Index As Integer)
    Grid_Hide
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
        TopCtrl1_eSave
    Else
        txt(Index).SetFocus
    End If
End Sub
Private Sub Ini_Grid()
    SSTab1.Tab = 0
    SSTab1.left = 60: SSTab1.top = 3500
    FGrid.left = 0: FGrid.top = 0: FGrid.width = Frame3.width: FGrid.height = Frame3.height                '2085
    FGrid1.left = 0: FGrid1.top = 0: FGrid1.width = Frame2.width: FGrid1.height = Frame2.height                 '2085
    FGrid2.left = 0: FGrid2.top = 0: FGrid2.width = Frame1.width: FGrid2.height = Frame1.height                 '2085
    DgCategory.left = 7665: DgCategory.top = 465
    DGNCat.left = 7665: DGNCat.top = 465
    DGVType.left = 7665: DGVType.top = 465
    DGGroup.left = 7665: DGGroup.top = 465
    DGCR.left = txt(DefaultCrAC).left: DGCR.top = txt(DefaultCrAC).top + txt(DefaultCrAC).height
    DGDR.left = txt(DefaultDrAC).left: DGDR.top = txt(DefaultDrAC).top + txt(DefaultDrAC).height
    With FGrid
        .Cols = 6
         BackColorSelLeave = .BackColor
        .ColWidth(FSNo) = 500                           ' marker
        .TextMatrix(0, FVType) = "Voucher Type"         ' vtype
        .ColWidth(FVType) = 0
        .TextMatrix(0, FFromDate) = "From Date"                  'from date
        .ColAlignmentFixed(FFromDate) = flexAlignLeftCenter
        .ColAlignment(FFromDate) = flexAlignLeftCenter
        .ColWidth(FFromDate) = 2500
        .TextMatrix(0, FToDate) = "To Date"                  'To date
        .ColAlignmentFixed(FToDate) = flexAlignLeftCenter
        .ColAlignment(FToDate) = flexAlignLeftCenter
        .ColWidth(FToDate) = 2500
        .TextMatrix(0, FPrefix) = "Prefix"
        .ColAlignmentFixed(FPrefix) = flexAlignLeftCenter
        .ColAlignment(FPrefix) = flexAlignLeftCenter
        .ColWidth(FPrefix) = 2000
        .TextMatrix(0, FStartSrNo) = "Start Sr. No."
        .ColAlignmentFixed(FStartSrNo) = flexAlignRightCenter
        .ColAlignment(FStartSrNo) = flexAlignRightCenter
        .ColWidth(FStartSrNo) = 2000
    End With
    With FGrid1
        .Cols = 5
         BackColorSelLeave = .BackColor
        .ColWidth(FSno1) = 500                           ' marker
        .TextMatrix(0, FGroupName1) = "Group Name"                   'grp name
        .ColAlignmentFixed(FGroupName1) = flexAlignLeftCenter
        .ColAlignment(FGroupName1) = flexAlignLeftCenter
        .ColWidth(FGroupName1) = 3000
        .TextMatrix(0, FDr1) = "Dr"
        .ColAlignmentFixed(FDr1) = flexAlignLeftCenter
        .ColAlignment(FDr1) = flexAlignLeftCenter
        .ColWidth(FDr1) = 1000
        .TextMatrix(0, FCr1) = "Cr"
        .ColAlignmentFixed(FCr1) = flexAlignLeftCenter
        .ColAlignment(FCr1) = flexAlignLeftCenter
        .ColWidth(FCr1) = 1000
        .ColWidth(FGroupCode1) = 0
    End With
    With FGrid2
        .Cols = 5
         BackColorSelLeave = .BackColor
        .ColWidth(FSno2) = 500                           ' marker
        .TextMatrix(0, FGroupName2) = "Group Name"                   'grp name
        .ColAlignmentFixed(FGroupName2) = flexAlignLeftCenter
        .ColAlignment(FGroupName2) = flexAlignLeftCenter
        .ColWidth(FGroupName2) = 3000
        .TextMatrix(0, FDr2) = "Dr"
        .ColAlignmentFixed(FDr2) = flexAlignLeftCenter
        .ColAlignment(FDr2) = flexAlignLeftCenter
        .ColWidth(FDr2) = 1000
        .TextMatrix(0, FCr2) = "Cr"
        .ColAlignmentFixed(FCr2) = flexAlignLeftCenter
        .ColAlignment(FCr2) = flexAlignLeftCenter
        .ColWidth(FCr2) = 1000
        .ColWidth(FGroupCode2) = 0
    End With
End Sub
Private Sub BlankFGrid1()
Dim I As Integer
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
End Sub
Private Sub FillFGrid2()
Dim I As Integer, rsgroup As New ADODB.Recordset, Rst As ADODB.Recordset, j As Integer
FGrid2.Rows = 1
FGrid2.AddItem FGrid2.Rows
FGrid2.FixedRows = 1
If RsEnv!DonotShowGroup = "Yes" Then
    rsgroup.Open "Select GroupCode as Code,GroupName as Name From AcGroup Where SysGroup='Y' and AliasYN='N' Order by GroupName", G_FaCn, adOpenDynamic, adLockOptimistic
Else
    rsgroup.Open "Select GroupCode as Code,GroupName as Name From AcGroup Where AliasYN='N' Order by GroupName", G_FaCn, adOpenDynamic, adLockOptimistic
End If
    If rsgroup.RecordCount > 0 Then
        I = 1
        Do Until rsgroup.EOF
            FGrid2.AddItem ""
            With FGrid2
                .TextMatrix(I, FSno2) = I
                .TextMatrix(I, FGroupName2) = rsgroup!Name
                .TextMatrix(I, FDr2) = ""
                .TextMatrix(I, FCr2) = ""
                .TextMatrix(I, FGroupCode2) = rsgroup!Code
            End With
            I = I + 1
            rsgroup.MoveNext
        Loop
        FGrid2.FixedRows = 1
    End If
    If TopCtrl1.TopText2.CAPTION = "Edit" Then
        Set Rst = G_FaCn.Execute(" Select GroupCode,CR,DR From Voucher_Exclude Where V_Type='" & txt(VType) & "'")
        While Not Rst.EOF
            For j = 1 To FGrid2.Rows - 1
                If Rst!GroupCode = FGrid2.TextMatrix(j, FGroupCode2) Then
                    FGrid2.Row = j
                    FGrid2.Col = FDr2
                    FGrid2.CellFontName = "wingdings"
                    FGrid2.CellFontSize = 18
                    FGrid2.CellForeColor = vbRed
                    FGrid2.TextMatrix(j, FDr1) = IIf(Rst!dr = "Y", "ü", "")
                    FGrid2.Row = j
                    FGrid2.Col = FCr2
                    FGrid2.CellFontName = "wingdings"
                    FGrid2.CellFontSize = 18
                    FGrid2.CellForeColor = vbRed
                    FGrid2.TextMatrix(j, FCr2) = IIf(Rst!cr = "Y", "ü", "")
                    Exit For
                End If
            Next
        Rst.MoveNext
        Wend
    End If
    Set rsgroup = Nothing: Set Rst = Nothing
End Sub
Private Sub DGDR_Click()
    DGDR.Visible = False
    If RsDr.RecordCount > 0 Then
        txt(DefaultDrAC).Tag = RsDr!Code
        txt(DefaultDrAC).TEXT = RsDr!Name
    End If
    txt(DefaultDrAC).SetFocus
End Sub
Private Sub DGCR_Click()
    DGCR.Visible = False
    If RsCr.RecordCount > 0 Then
        txt(DefaultCrAC).Tag = RsCr!Code
        txt(DefaultCrAC).TEXT = RsCr!Name
    End If
    txt(DefaultCrAC).SetFocus
End Sub
