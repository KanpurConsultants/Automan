VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmPartMast 
   AutoRedraw      =   -1  'True
   Caption         =   " Part Master"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14400
   FillColor       =   &H00C0E0FF&
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   14400
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
      Height          =   210
      Index           =   23
      Left            =   8730
      MaxLength       =   15
      TabIndex        =   20
      Top             =   2400
      Width           =   1185
   End
   Begin MSDataGridLib.DataGrid DGDep_Item 
      Height          =   4455
      Left            =   10170
      Negotiate       =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   2850
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   22
      Left            =   5490
      TabIndex        =   22
      Top             =   2400
      Width           =   1185
   End
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   3180
      Left            =   360
      Negotiate       =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   5609
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
      Caption         =   "PART HELP"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Part No"
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
         Caption         =   "MRP"
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
         DataField       =   "CurStk"
         Caption         =   "Curr.Stock"
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
         DataField       =   "Local_Name"
         Caption         =   "Local Name"
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   2759.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4050.142
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   4050.142
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmdbin 
      Caption         =   "Change Bin Location"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   0
      Width           =   1875
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3180
      Left            =   75
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3165
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   5609
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388736
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Opening Stock"
      TabPicture(0)   =   "frmPartMast.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Alternate Parts"
      TabPicture(1)   =   "frmPartMast.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00CFE0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2445
         Left            =   -74940
         TabIndex        =   54
         Top             =   390
         Width           =   11730
         Begin VB.TextBox TxtGrid 
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   0
            Left            =   1500
            TabIndex        =   55
            Top             =   1680
            Visible         =   0   'False
            Width           =   1080
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
            Height          =   3840
            Left            =   105
            TabIndex        =   24
            Top             =   75
            Width           =   11790
            _ExtentX        =   20796
            _ExtentY        =   6773
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   10
            BackColorFixed  =   13623520
            ForeColorFixed  =   16384
            BackColorSel    =   16761024
            BackColorBkg    =   13623520
            GridColor       =   0
            GridColorFixed  =   8421504
            FocusRect       =   0
            AllowUserResizing=   3
            BorderStyle     =   0
            FormatString    =   $"frmPartMast.frx":0038
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
            _Band(0).Cols   =   10
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00CFE0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2115
         Left            =   60
         TabIndex        =   52
         Top             =   540
         Width           =   11730
         Begin VB.TextBox TxtGrid 
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   1
            Left            =   3045
            TabIndex        =   53
            Top             =   1365
            Visible         =   0   'False
            Width           =   1080
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
            Height          =   1950
            Left            =   0
            TabIndex        =   23
            Top             =   15
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   3440
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   11
            BackColorFixed  =   13623520
            ForeColorFixed  =   0
            BackColorSel    =   16761024
            ForeColorSel    =   12648447
            BackColorBkg    =   13623520
            GridColor       =   0
            GridColorFixed  =   32768
            FocusRect       =   0
            AllowUserResizing=   3
            BorderStyle     =   0
            Appearance      =   0
            FormatString    =   $"frmPartMast.frx":00E6
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
            _Band(0).Cols   =   11
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Index           =   17
      Left            =   1935
      TabIndex        =   17
      Top             =   2160
      Width           =   1185
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   18
      Left            =   8715
      TabIndex        =   19
      Top             =   2160
      Width           =   1185
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   19
      Left            =   5490
      TabIndex        =   18
      Top             =   2160
      Width           =   1185
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   20
      Left            =   1935
      TabIndex        =   21
      Top             =   2400
      Width           =   1185
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   11
      Left            =   5490
      MaxLength       =   15
      TabIndex        =   5
      Top             =   1200
      Width           =   1185
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   12
      Left            =   1935
      TabIndex        =   7
      Top             =   1440
      Width           =   1185
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   13
      Left            =   8715
      TabIndex        =   10
      Top             =   1440
      Width           =   1185
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
      Height          =   210
      Index           =   14
      Left            =   5490
      TabIndex        =   12
      Top             =   1680
      Width           =   1185
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
      Height          =   210
      Index           =   15
      Left            =   1935
      MaxLength       =   1
      TabIndex        =   11
      Top             =   1680
      Width           =   1185
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
      Height          =   210
      Index           =   16
      Left            =   8715
      MaxLength       =   6
      TabIndex        =   13
      Top             =   1680
      Width           =   1185
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
      Height          =   240
      Index           =   7
      Left            =   9810
      MaxLength       =   8
      TabIndex        =   43
      Text            =   "Yes"
      Top             =   3990
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   8
      Left            =   8715
      TabIndex        =   16
      Top             =   1920
      Width           =   1185
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   9
      Left            =   1935
      TabIndex        =   14
      Top             =   1920
      Width           =   1185
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   10
      Left            =   5490
      TabIndex        =   15
      Top             =   1920
      Width           =   1185
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   1935
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1200
      Width           =   1185
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
      Height          =   210
      Index           =   4
      Left            =   8715
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1200
      Width           =   1185
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
      Height          =   210
      Index           =   6
      Left            =   5490
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1440
      Width           =   480
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   5
      Left            =   6315
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSDataGridLib.DataGrid DGPartGrade 
      Height          =   3375
      Left            =   5475
      Negotiate       =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5895
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   5953
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
      Caption         =   "PROPRIETARY PART GRADE HELP"
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
         Caption         =   "Dealer "
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
   Begin MSDataGridLib.DataGrid DGDisFact 
      Height          =   3615
      Left            =   5625
      Negotiate       =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5580
      Visible         =   0   'False
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   6376
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
      Caption         =   "DISCOUNT FACTOR HELP"
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
         Caption         =   "Purch. %"
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
         DataField       =   "Name1"
         Caption         =   "Sale %"
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
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1305.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGGod 
      Height          =   4455
      Left            =   10275
      Negotiate       =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1800
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
      Caption         =   "GODOWN HELP"
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
         Caption         =   "Godown"
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
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   21
      Left            =   10845
      MaxLength       =   22
      TabIndex        =   35
      Top             =   3720
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   8940
      TabIndex        =   32
      Top             =   5520
      Visible         =   0   'False
      Width           =   2010
      Begin MSComctlLib.ListView ListView 
         Height          =   1815
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   -30
         Width           =   1800
         _ExtentX        =   3175
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
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   661
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
      Left            =   1935
      MaxLength       =   40
      TabIndex        =   3
      Top             =   945
      Width           =   4740
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
      Height          =   210
      Index           =   1
      Left            =   1935
      MaxLength       =   40
      TabIndex        =   2
      Top             =   705
      Width           =   4740
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
      Height          =   210
      Index           =   0
      Left            =   1935
      MaxLength       =   22
      TabIndex        =   1
      Top             =   465
      Width           =   2190
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
      Left            =   6795
      TabIndex        =   63
      Top             =   2415
      Width           =   2745
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N.D.P. .........................."
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
      Left            =   3225
      TabIndex        =   61
      Top             =   2415
      Width           =   2160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lead Time.........................."
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
      Left            =   165
      TabIndex        =   59
      Top             =   2168
      Width           =   2445
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min.Stock Level...................."
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
      Left            =   6795
      TabIndex        =   58
      Top             =   2168
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max.Stock Level...................."
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
      Left            =   3225
      TabIndex        =   57
      Top             =   2168
      Width           =   2595
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Order Level.........................."
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
      Left            =   165
      TabIndex        =   56
      Top             =   2408
      Width           =   2865
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Y)es/(N)o"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   1
      Left            =   10635
      TabIndex        =   51
      Top             =   4005
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bin Location............................"
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
      Index           =   55
      Left            =   3225
      TabIndex        =   50
      Top             =   1215
      Width           =   2715
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Stock Qty.................."
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
      Index           =   54
      Left            =   165
      TabIndex        =   49
      Top             =   1448
      Width           =   2505
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Stock Value................"
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
      Index           =   53
      Left            =   6795
      TabIndex        =   48
      Top             =   1448
      Width           =   2700
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Valuation Method.............."
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
      Index           =   52
      Left            =   3225
      TabIndex        =   47
      Top             =   1688
      Width           =   2850
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Security Grade*...................."
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
      Index           =   51
      Left            =   165
      TabIndex        =   46
      Top             =   1695
      Width           =   2595
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special Part (Y/N).................."
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
      Index           =   50
      Left            =   6795
      TabIndex        =   45
      Top             =   1695
      Width           =   2610
   End
   Begin VB.Label LblName 
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
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   3
      Left            =   9150
      TabIndex        =   44
      Top             =   4035
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MRP.................................."
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
      Left            =   6795
      TabIndex        =   42
      Top             =   1935
      Width           =   2400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable Selling Rate............."
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
      Left            =   165
      TabIndex        =   41
      Top             =   1935
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxpaid Selling Rate............."
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
      Left            =   3225
      TabIndex        =   40
      Top             =   1928
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit..................................."
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
      Left            =   165
      TabIndex        =   39
      Top             =   1185
      Width           =   2430
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proprietory Grade*...................."
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
      Left            =   6795
      TabIndex        =   38
      Top             =   1208
      Width           =   2460
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Factor*................."
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
      Left            =   3225
      TabIndex        =   37
      Top             =   1455
      Width           =   2445
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prefix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   16
      Left            =   9270
      TabIndex        =   34
      Top             =   3660
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No.*............................"
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
      Left            =   165
      TabIndex        =   27
      Top             =   465
      Width           =   2475
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part Desc. (Std.)*........."
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
      Left            =   165
      TabIndex        =   26
      Top             =   720
      Width           =   2085
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part Desc. (Local).........."
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
      Index           =   20
      Left            =   165
      TabIndex        =   25
      Top             =   960
      Width           =   2130
   End
End
Attribute VB_Name = "frmPartMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim GridKey As Integer
Dim TAddMode As Boolean
Private Const mVType As String = "SXAO"
Dim VoucherEditFlag As Boolean
Dim ExitCtrl As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem
Dim urec As Integer
Dim rsPartRefresh As Boolean
Private Const Part_No As Byte = 0, Part_Name As Byte = 1
Private Const Local_Name As Byte = 2, Unit As Byte = 3
Private Const Part_Grade_Name As Byte = 4, Part_Grade As Byte = 5
Private Const Disc_Factor As Byte = 6, Active_YN As Byte = 7, MRP As Byte = 8
Private Const TB_SRate As Byte = 9, TP_SRate As Byte = 10, Bin_Loca As Byte = 11
Private Const Curr_Stock As Byte = 12, Curr_Stock_Value As Byte = 13
Private Const Value_Method As Byte = 14, Security_Grade As Byte = 15, MARK_YN As Byte = 16
Private Const Lead_Time As Byte = 17, Min_Lvl As Byte = 18, Max_Lvl As Byte = 19
Private Const ReOrd_Lvl As Byte = 20, SerialNo As Byte = 21, NDP As Byte = 22, Dep_Item As Byte = 23
'for listview
Private Const lvVehType As Byte = 0, lvModelGroup As Byte = 1, lvAggregate As Byte = 2
Private Const lvOEM As Byte = 3, lvCity As Byte = 4, lvStockValuation As Byte = 5
Private Const MaxCol1 As Byte = 10, MaxCol2 As Byte = 6
'FOR ALETRNATE
Private Const AltNo As Byte = 1, StaPart As Byte = 2, LocPart As Byte = 3, MRPp As Byte = 4, TxRate As Byte = 5, TPRate As Byte = 6, Bin As Byte = 7
'FOR OPENING
Private Const ODate As Byte = 1, OGodown As Byte = 2, ORefNo As Byte = 3, OMRPYN As Byte = 4, OTAXYN As Byte = 5, OQuan As Byte = 6, OVal As Byte = 7, ORate As Byte = 8
Private Const OGodCode As Byte = 9, OLastInvNo As Byte = 10, OLastInvDt As Byte = 11
'for recordset
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstPartGrade As ADODB.Recordset
Dim RstDiscFact As ADODB.Recordset, RstUnit As ADODB.Recordset
Dim RstGod As ADODB.Recordset
Dim RstDEp_item As ADODB.Recordset
Dim ADDFLAG As Byte, mFlag As Byte, mFLAG1 As Byte
Dim OldRateMRP As Double, OldRateTB As Double, OldRateTP As Double, DiffRate As Double
Private Sub Cmdbin_Click()
    frmBinChange.Show vbModal
End Sub
Private Sub DGDisFact_Click()
txt(Disc_Factor).TEXT = DGDisFact.TEXT
DGDisFact.Visible = False
txt(Disc_Factor).SetFocus
End Sub

Private Sub DGDisFact_GotFocus()
    mFlag = 1
End Sub

Private Sub DGDisFact_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mFlag = 1 Then
    txt(Disc_Factor) = DGDisFact.Columns(0).TEXT
End If
End Sub

Private Sub DGGod_Click()
If RstGod.RecordCount > 0 Then
    FGrid1.TextMatrix(FGrid1.Row, OGodown) = RstGod!Name
    FGrid1.TextMatrix(FGrid1.Row, OGodCode) = RstGod!Code
    TxtGrid(1) = RstGod!Name
End If
TxtGrid(1).SetFocus
DGGod.Visible = False
End Sub
Private Sub DGPart_Click()
    DGPart.Visible = False
End Sub
Private Sub DGPartGrade_Click()
    txt(Part_Grade_Name) = RstPartGrade!Name
    txt(Part_Grade) = RstPartGrade!Code
    txt(Part_Grade_Name).SetFocus
    DGPartGrade.Visible = False
End Sub
Private Sub DGDep_Item_click()
    txt(Dep_Item) = RstDEp_item!Name
    txt(Dep_Item).Tag = RstDEp_item!Code
    txt(Dep_Item).SetFocus
    DGDep_Item.Visible = False
End Sub

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid.Col
    Case AltNo, StaPart, LocPart, Bin, MRPp, TxRate, TPRate
        Call GridDblClick(Me, FGrid, TxtGrid, 0)
End Select
TAddMode = False
End Sub

Private Sub FGrid_GotFocus()
'    FGrid.CellBackColor = CellBackColEnter
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
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case MRPp, TxRate, TPRate
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case AltNo, StaPart, LocPart, Bin
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
Select Case FGrid.Col
    Case AltNo, StaPart, LocPart, Bin
       Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    Case MRPp, TxRate, TPRate
       Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
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
End Sub

Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid1_Click()
'TxtGrid(1).Visible = False
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
SetMaxLength
End Sub

Private Sub FGrid1_DblClick()
FGrid1_KeyPress vbKeyReturn
End Sub

Private Sub FGrid1_GotFocus()
    Grid_Hide
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid1.Tag = FGrid1.Row

If KeyCode = vbKeyDelete And Shift = 0 Then
Select Case FGrid1.Col
    Case OGodown
        If KeyCode = vbKeyDelete And Shift = 0 Then Exit Sub
    Case ORefNo
        If KeyCode = vbKeyDelete And Shift = 0 Then FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
    Case OQuan, ORate, OLastInvNo, OLastInvDt
        If RSOJPR = True And TopCtrl1.TopText2 = "Edit" And FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) <> "" Then
            Exit Sub
        Else
            If KeyCode = vbKeyDelete And Shift = 0 Then FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = "": FGrid1.TextMatrix(FGrid1.Row, OVal) = ""
        End If
End Select
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid1.Col
        Case OGodown, ORefNo, OLastInvNo, OLastInvDt
            Call GridDblClick(Me, FGrid1, TxtGrid, 1)
            TAddMode = False
        Case OMRPYN, OTAXYN, OQuan, ORate
            If RSOJPR = True And TopCtrl1.TopText2 = "Edit" And FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) <> "" Then
                Exit Sub
            Else
                Call GridDblClick(Me, FGrid1, TxtGrid, 1)
                TAddMode = False
            End If

    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
Dim mUChr As String
mUChr = UCase(Chr(KeyAscii))
SetMaxLength
Select Case FGrid1.Col
    Case ODate, OGodown, ORefNo
       Call Get_Text(Me, FGrid1, TxtGrid, 1, False, KeyAscii)
    Case OLastInvNo, OLastInvDt
       Call Get_Text(Me, FGrid1, TxtGrid, 1, True, KeyAscii)
    Case OQuan, ORate, OVal
       If RSOJPR = True And TopCtrl1.TopText2 = "Edit" And FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) <> "" Then
            Exit Sub
       Else
            Call Get_Text(Me, FGrid1, TxtGrid, 1, True, KeyAscii)
       End If
            
    Case OMRPYN, OTAXYN
        If mUChr = "N" Then
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = "No"
        ElseIf mUChr = "Y" Then
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = "Yes"
        End If
        KeyAscii = 0
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid1.ColSel = False Then Exit Sub
If RSOJPR = True And TopCtrl1.TopText2 = "Edit" And FGrid1.TextMatrix(FGrid1.Row, ODate) <> "" Then
    Exit Sub
End If
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

Private Sub FGrid1_Scroll()
    TxtGrid(1).Visible = False
    Grid_Hide
End Sub

Private Sub Form_Activate()
Dim MsgStr$
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
    Call TopCtrl1_eRef
End If
If GCn.Execute("SELECT FINCATG FROM ContractFinance WHERE FINCATG=2 ORDER BY FINNAME").RecordCount <= 0 Then
    MsgStr = "OEM Master"
End If
If GCn.Execute("SELECT CITYCODE FROM CITY ORDER BY CITYNAME").RecordCount <= 0 Then
    MsgStr = "City Master"
End If
If GCn.Execute("SELECT Vehicle_Type FROM Vehicle_Type ORDER BY Vehicle_Type").RecordCount <= 0 Then
    MsgStr = "Vehicle_Type"
End If
If GCn.Execute("SELECT ModelGrp_Code FROM MODEL_GRP ORDER BY ModelGrp_Name").RecordCount <= 0 Then
    MsgStr = "Model Group"
End If
If GCn.Execute("SELECT Aggre_Code FROM Aggregate ORDER BY Aggre_Name").RecordCount <= 0 Then
    MsgStr = "Aggregate"
End If
'If MsgStr <> "" Then
'    MsgBox "Add atleast 1 record in " & MsgStr & vbCrLf & " Part Master Loading Aborted!", vbCritical, "Validation !"
''    Unload Me
'End If
End Sub
Private Sub Form_Deactivate()
    If rsPartRefresh Then RsPart.Requery: rsPartRefresh = False
'    If MasterFormExit  Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
ELoop:
Exit Sub
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
'On Error GoTo ELoop
Dim xITEM As ListItem, mMousePointer As Long
TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini
'**Speed
Me.Show
DoEvents
mMousePointer = Me.MousePointer
Me.MousePointer = vbHourglass
'**
rsPartRefresh = False
ADDFLAG = 0: mFlag = 0: mFLAG1 = 0
SSTab1.Tab = 0

Set RstMain = New ADODB.Recordset
RstMain.CursorLocation = adUseClient
If PubMoveRecYn Then
    Set RstMain = RsPart.Clone
Else
    Set RstMain = GCn.Execute("Select Top 1 P.Part_No as Code, P.Part_No, P.Part_Name As Name, P.Local_Name as LName, " & _
                              "P.Unit , P.MRP, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, (Select PurcDisc_Per From Part_DiscFactor Where DiscFac_Catg=P.Disc_Factor) As PurcDisc_Per, P.ReOrd_Lvl, " & _
                              "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=1 And Tax_Yn=1) As Cur_MRP_TBStk, " & _
                              "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=1 And Tax_Yn=0) As Cur_MRP_TpStk, " & _
                              "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "  Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=0 And Tax_Yn=1) As Cur_TB_Stk, " & _
                              "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=0 And Tax_Yn=0) As Cur_Tp_Stk, " & _
                              "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No ) As CurrStk, " & _
                              "P.Min_Lvl, P.Disc_Factor " & _
                              "From Part P " & _
                              "Left Join Sp_Stock Stk On P.Part_No=Stk.Part_No " & _
                              "WHERE  Div_Code='C' " & _
                              "Group By P.Part_No, P.Part_Name, P.Local_Name, P.Unit, P.Mrp, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Disc_Factor, P.ReOrd_Lvl, P.Min_Lvl")
End If
'RstMain.Open "SELECT PART_NO AS SEARCHCODE from part where Div_Code='" & PubDivCode & "' order by Part_No", GCn, adOpenDynamic, adLockOptimistic ', adAsyncFetch

Set RstHelp = New ADODB.Recordset
RstHelp.CursorLocation = adUseClient
Set RstHelp = RsPart.Clone
'RstHelp.Open "SELECT PART_NO,PART_NAME,UNIT,TP_SRate,MRP,LOCAL_NAME,Bin_Loca,TB_SRate,(Cur_MRP_TBStk+Cur_MRP_TPStk+Cur_TB_Stk+Cur_TP_Stk) AS CURSTK FROM Part where Div_Code='" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
Set DGPart.DataSource = RstHelp
'RstHelp.Sort = "Part_No"
'RstHelp.Sort = "Part_Name"
'RstHelp.Sort = "Local_Name"

Set RstPartGrade = New ADODB.Recordset
RstPartGrade.CursorLocation = adUseClient
RstPartGrade.Open "Select partgrade_code as code ,partgrade_name as name From Part_GRADE", GCn, adOpenDynamic, adLockOptimistic
Set DGPartGrade.DataSource = RstPartGrade

Set RstDiscFact = New ADODB.Recordset
RstDiscFact.CursorLocation = adUseClient
RstDiscFact.Open "Select DiscFac_Catg as code ,PurcDisc_Per as name,SalDisc_Per as name1 From Part_DiscFactor", GCn, adOpenDynamic, adLockOptimistic
Set DGDisFact.DataSource = RstDiscFact

Set RstGod = New ADODB.Recordset
RstGod.CursorLocation = adUseClient
RstGod.Open "Select GOD_CODE as code ,GOD_NAME as name From GODOWN WHERE APPLI_FOR=0", GCn, adOpenDynamic, adLockOptimistic
Set DGGod.DataSource = RstGod


'Nikhil
Set RstDEp_item = New ADODB.Recordset
RstDEp_item.CursorLocation = adUseClient
RstDEp_item.Open "Select CODE as code ,Description as name From Deprecation_itemMaster ", GCn, adOpenDynamic, adLockOptimistic
Set DGDep_Item.DataSource = RstDEp_item


DGPartGrade.Visible = False

DGDep_Item.Visible = False

Disp_Text SETS("INI", Me, RstMain)
MoveRec
Me.MousePointer = mMousePointer

Exit Sub
ELoop:
MsgBox err.Description, vbCritical: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing: Set RstPartGrade = Nothing: Set RstDiscFact = Nothing: Set RstDEp_item = Nothing
    If rsPartRefresh Then RsPart.Requery
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To 23
    txt(I).Enabled = Enb
Next
txt(Curr_Stock).Enabled = False
txt(Curr_Stock_Value).Enabled = False
End Sub

Private Sub MoveRec()
Dim xITEM As ListItem
Dim fob As New FileSystemObject
Dim Rstmain1 As Recordset
Dim Tss As String
Dim URST As ADODB.Recordset
On Error GoTo ErrLoop
RST_BOF_EOF RstMain
TopCtrl1.tDel = False
OldRateMRP = 0
OldRateTB = 0
OldRateTP = 0
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    Set Rstmain1 = New Recordset
    Rstmain1.CursorLocation = adUseClient
    Rstmain1.Open "SELECT PART.PART_NO AS SEARCHCODE,part.*,Part_Grade.PartGrade_Name,Part_DiscFactor.DiscFac_Catg,Part_DiscFactor.PurcDisc_Per,Part_DiscFactor.SalDisc_Per,Ditm.Description as DepItemname  " & _
    " FROM ((Part LEFT JOIN Part_Grade ON Part.Part_Grade = Part_Grade.PartGrade_Code) LEFT JOIN Part_DiscFactor ON Part.Disc_Factor = Part_DiscFactor.DiscFac_Catg) " & _
    " left join Deprecation_itemMaster ditm on part.Dep_Item=ditm.code " & _
    " where part.part_no+Part.Div_Code = '" & RstMain!Code & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic  ', adAsyncFetch

    txt(Part_No) = IIf(IsNull(Rstmain1!Part_No), "", Rstmain1!Part_No)
    txt(Part_Name) = IIf(IsNull(Rstmain1!Part_Name), "", Rstmain1!Part_Name)
    txt(Local_Name) = IIf(IsNull(Rstmain1!Local_Name), "", Rstmain1!Local_Name)
    txt(Unit) = IIf(IsNull(Rstmain1!Unit), "", Rstmain1!Unit)
    txt(Part_Grade) = IIf(IsNull(Rstmain1!Part_Grade), "", Rstmain1!Part_Grade)
    txt(Part_Grade_Name) = IIf(IsNull(Rstmain1!PartGrade_Name), "", Rstmain1!PartGrade_Name)
    txt(Disc_Factor) = IIf(IsNull(Rstmain1!Disc_Factor), "", Rstmain1!Disc_Factor)
    txt(MRP) = IIf(IsNull(Rstmain1!MRP) Or Rstmain1!MRP = 0, "", Format(Rstmain1!MRP, "0.000"))
    txt(TB_SRate) = IIf(IsNull(Rstmain1!TB_SRate) Or Rstmain1!TB_SRate = 0, "", Format(Rstmain1!TB_SRate, "0.000"))
    txt(TP_SRate) = IIf(IsNull(Rstmain1!TP_SRate) Or Rstmain1!TP_SRate = 0, "", Format(Rstmain1!TP_SRate, "0.000"))
    txt(NDP) = Format(VNull(Rstmain1!NDP), "0.00")
    txt(Bin_Loca) = IIf(IsNull(Rstmain1!Bin_Loca), "", Rstmain1!Bin_Loca)
    txt(Curr_Stock) = IIf(VNull(Rstmain1!Cur_MRP_TbStk) + VNull(Rstmain1!Cur_MRP_TPStk) + VNull(Rstmain1!Cur_TB_STk) + VNull(Rstmain1!Cur_TP_Stk) = 0, "", VNull(Rstmain1!Cur_MRP_TbStk) + VNull(Rstmain1!Cur_MRP_TPStk) + VNull(Rstmain1!Cur_TB_STk) + VNull(Rstmain1!Cur_TP_Stk))
    txt(Curr_Stock_Value) = IIf(VNull(Rstmain1!Cur_MRP_TBStk_Val) + VNull(Rstmain1!Cur_MRP_TPStk_Val) + VNull(Rstmain1!Cur_TB_Stk_Val) + VNull(Rstmain1!Cur_TP_Stk_Val) = 0, "", VNull(Rstmain1!Cur_MRP_TBStk_Val) + VNull(Rstmain1!Cur_MRP_TPStk_Val) + VNull(Rstmain1!Cur_TB_Stk_Val) + VNull(Rstmain1!Cur_TP_Stk_Val))
    txt(Value_Method) = XNull(Rstmain1!Value_Method)
    txt(Security_Grade) = XNull(Rstmain1!Security_Grade)
    txt(MARK_YN) = IIf(XNull(Rstmain1!MARK_YN) = "Y", "Yes", "No")
    txt(Lead_Time) = IIf(IsNull(Rstmain1!Lead_Time) Or Rstmain1!Lead_Time = 0, "", Format(Rstmain1!Lead_Time, "0.00"))
    txt(Min_Lvl) = IIf(IsNull(Rstmain1!Min_Lvl) Or Rstmain1!Min_Lvl = 0, "", Format(Rstmain1!Min_Lvl, "0.00"))
    txt(Max_Lvl) = IIf(IsNull(Rstmain1!Max_Lvl) Or Rstmain1!Max_Lvl = 0, "", Format(Rstmain1!Max_Lvl, "0.00"))
    txt(ReOrd_Lvl) = IIf(IsNull(Rstmain1!ReOrd_Lvl) Or Rstmain1!ReOrd_Lvl = 0, "", Format(Rstmain1!ReOrd_Lvl, "0.00"))
    'Txt(Active_YN) = IIf(VNull(Rstmain1!Active_YN) = 1, "Yes", "No")
'Nikhil
    txt(Dep_Item).Tag = IIf(IsNull(Rstmain1!Dep_Item), "", Rstmain1!Dep_Item)
    txt(Dep_Item) = IIf(IsNull(Rstmain1!DepItemname), "", Rstmain1!DepItemname)
    
On Error Resume Next
  '***********************ALTERNATE PART
    Set URST = GCn.Execute("SELECT Part_Name,Alternate_Part_No,Local_Name,TB_SRate,TP_SRate,Bin_Loca " & _
            "from PART_ALTERNATE LEFT JOIN PART ON (PART_ALTERNATE.Alternate_Part_No=PART.PART_NO and Part.Div_Code = PART_ALTERNATE.Div_Code ) " & _
            "WHERE Root_Part_No='" & txt(Part_No) & "' and PART_ALTERNATE.Div_Code = '" & PubDivCode & "'")
    If URST.RecordCount > 0 Then
        FGrid.Rows = 1
        Do Until URST.EOF
            FGrid.AddItem FGrid.Rows & Chr(9) & URST!Alternate_Part_No & Chr(9) & URST!Part_Name & Chr(9) & URST!Local_Name & Chr(9) & Format(URST!TB_SRate, "0.00") & Chr(9) & Format(URST!TP_SRate, "0.00") & Chr(9) & URST!Bin_Loca
            URST.MoveNext
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.Rows = 1
        FGrid.AddItem "1"
        FGrid.FixedRows = 1
    End If
'***********************OPENING
    Set URST = GCn.Execute("SELECT SRL_NO,V_DATE,god_name,GOD_CODE,remark,MRP_YN,TAX_YN,QTY_REC,RATE,AMOUNT,PurDocNo,PurDocDate " & _
        "from SP_STOCK left join godown on (sp_stock.godown=godown.god_code) " & _
        "WHERE V_TYPE='" & mVType & "' AND Part_No='" & txt(Part_No) & "' AND Left(SP_Stock.DocID,1)='" & PubDivCode & "' and " & cMID("Sp_Stock.DocId", "3", "1") & "='" & PubSiteCode & "' ORDER BY SRL_NO")
    If URST.RecordCount > 0 Then
        FGrid1.Rows = 1
        Do Until URST.EOF
            FGrid1.AddItem FGrid1.Rows & Chr(9) & URST!V_DATE & Chr(9) & URST!God_Name & Chr(9) & URST!Remark & Chr(9) & IIf(URST!MRP_YN = 1, "Yes", "No") & Chr(9) & IIf(URST!Tax_YN = 1, "Yes", "No") & Chr(9) & Format(URST!Qty_Rec, "0.00") & Chr(9) & Format(URST!Amount, "0.00") & Chr(9) & Format(URST!Rate, "0.00") & Chr(9) & URST!God_Code & Chr(9) & URST!PurDocNo & Chr(9) & URST!PurDocDate
            URST.MoveNext
        Loop
        FGrid1.FixedRows = 1
    Else
        FGrid1.Rows = 1
        FGrid1.AddItem "1"
        FGrid1.FixedRows = 1
    End If
End If
ErrLoop:
    Set Rstmain1 = Nothing
    If err.NUMBER <> 0 Then MsgBox err.Description
End Sub

Private Sub ListView_Click()
txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
FrmList.Visible = False
txt(Val(ListView.Tag)).SetFocus
End Sub

Private Sub Picture1_dblClick(Index As Integer)
On Error GoTo ELoop
Dim PartFileN$, dQot$
If TopCtrl1.TopText2 = "Browse" Then Exit Sub
dQot = Chr(34)
'If Trim(Picture1(index).Tag) = Trim(Pub_DataPath & "\Pictures\Parts" & "\nopicture.bmp") Then
'        CDlg.Action = 1
'        FileCopy CDlg.FileName, Pub_DataPath & "\Pictures\Parts" & "\" & Txt(Part_No).Text
'        Picture1(index).Tag = Pub_DataPath & "\Pictures\Parts" & "\" & Txt(Part_No).Text
'        Picture1(index).Picture = LoadPicture(Picture1(index).Tag)
'Else
'    If MsgBox("Do You Wnat To Change Photo,Previous Will Be Removed ?", vbYesNo, "Cancel") = vbYes Then

        'Characters not permitted in File Name  <>/\""|:*?
        PartFileN = txt(Part_No)
        If InStr(1, PartFileN, "<", vbTextCompare) > 0 Then
            PartFileN = Replace(PartFileN, "<", " ")
        End If
        If InStr(1, PartFileN, " >", vbTextCompare) > 0 Then
            PartFileN = Replace(PartFileN, " >", " ")
        End If
        If InStr(1, PartFileN, "/", vbTextCompare) > 0 Then
            PartFileN = Replace(PartFileN, "/", " ")
        End If
        If InStr(1, PartFileN, "\", vbTextCompare) > 0 Then
            PartFileN = Replace(PartFileN, "\", " ")
        End If
        If InStr(1, PartFileN, ":", vbTextCompare) > 0 Then
            PartFileN = Replace(PartFileN, ":", " ")
        End If
        If InStr(1, PartFileN, "|", vbTextCompare) > 0 Then
            PartFileN = Replace(PartFileN, "|", " ")
        End If
        If InStr(1, PartFileN, dQot, vbTextCompare) > 0 Then
            PartFileN = Replace(PartFileN, dQot, " ")
        End If
'        If Fob.FileExists(Pub_DataPath & "\Pictures\Parts" & "\" & PartFileN) Then
'            MsgBox "Part Picture File already exists !" & vbCrLf & "Please wait creating Unique Picture File Name", vbOKOnly, "Picture File Name Checking"
'        End If
'        Do While True
'            If mSNO  >= 100 Then
'                For mSNO = 1 To 100
'                    mFILE_NAME = "C:\REPTMP\REP" + Trim(str(mSNO))
'                    If Fob.FileExists(Trim(mFILE_NAME) + ".TXT") Then Fob.DeleteFile (Trim(mFILE_NAME) + ".TXT")
'                Next
'                mSNO = 1
'            End If
'            mFILE_NAME = "C:\REPTMP\REP" + Trim(str(mSNO))
'            If Fob.FileExists(Trim(mFILE_NAME) + ".TXT") Then
'                mSNO = mSNO + 1
'            Else
'                Exit Do
'            End If
'        Loop

'    End If
'End If
Exit Sub
ELoop:
    If err.NUMBER <> 0 Then MsgBox err.Description
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Grid_Hide
End Sub

Public Sub TopCtrl1_eAdd()
'    Set RstHelp = New ADODB.Recordset
'If RstHelp.State = 0 Then
'    RstHelp.CursorLocation = adUseClient
'    RstHelp.Open "SELECT PART_NO,PART_NAME,UNIT,TP_SRate,MRP,LOCAL_NAME,Bin_Loca,TB_SRate,(Cur_MRP_TBStk+Cur_MRP_TPStk+Cur_TB_Stk+Cur_TP_Stk) AS CURSTK FROM Part where Div_Code='" & PubDivCode & "' order by Part_Name", GCn, adOpenDynamic, adLockOptimistic
'    Set DGPart.DataSource = RstHelp
''    RstHelp.Sort = "Part_No"
''    RstHelp.Sort = "Part_Name"
''    RstHelp.Sort = "Local_Name"
'End If
On Error GoTo ErrLoop
BlankText
ADDFLAG = 1
Disp_Text SETS("ADD", Me, RstMain)
FillListV
FGrid1.TextMatrix(FGrid1.Rows - 1, OMRPYN) = "Yes"
FGrid1.TextMatrix(FGrid1.Rows - 1, OTAXYN) = "Yes"
'Txt(Part_No) = PubDivCode
'Txt(Part_No).SelStart = Len(Txt(Part_No))
txt(Part_No).SetFocus
Exit Sub
ErrLoop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ErrLoop
'**********************
' Set RstHelp = New ADODB.Recordset
' If RstHelp.State = 0 Then
'    RstHelp.CursorLocation = adUseClient
'    RstHelp.Open "SELECT PART_NO,PART_NAME,UNIT,TP_SRate,MRP,LOCAL_NAME,Bin_Loca,TB_SRate,(Cur_MRP_TBStk+Cur_MRP_TPStk+Cur_TB_Stk+Cur_TP_Stk) AS CURSTK FROM Part where Div_Code='" & PubDivCode & "' order by Part_Name", GCn, adOpenDynamic, adLockOptimistic
'    Set DGPart.DataSource = RstHelp
''    RstHelp.Sort = "Part_No"
''    RstHelp.Sort = "Part_Name"
''    RstHelp.Sort = "Local_Name"
'End If

If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RstMain)
    txt(Part_No).Enabled = False
    txt(Part_Name).Tag = txt(Part_Name)
    Txt_GotFocus Part_Name
    txt(Part_Name).SetFocus
    FillListV
    txt(Part_Name).Tag = txt(Part_Name)
    txt(Local_Name).Tag = txt(Local_Name)
    
    If RSOJPR = True Then
        If FGrid1.TextMatrix(1, ODate) <> "" Then
            FGrid1.Enabled = False
        Else
            FGrid1.Enabled = True
        End If
    End If
'   vikash add for  rdb highway
    
   FGrid1.Enabled = True
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub

ErrLoop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub

Private Sub TopCtrl1_eDel()
Dim XBM
On Error GoTo ErrLoop
    If MsgBox("Are You Sure to Delete This Record", vbYesNo, "Confirmation") = vbYes Then
        GCn.BeginTrans
        XBM = RstMain.Bookmark
'        GCn.Execute ("DELETE From SP_STOCK Where V_Type='" & mVType & "' AND Part_No='" & txt(Part_No) & "'")
        GCn.Execute ("DELETE From SP_STOCK Where Part_No='" & txt(Part_No) & "'")
        GCn.Execute ("DELETE From Part_Alternate Where Root_Part_No='" & txt(Part_No) & "'")
        GCn.Execute ("DELETE From Part Where PART_NO='" & txt(Part_No) & "' AND div_code ='" & PubDivCode & "'")
        GCn.CommitTrans
        RstMain.Requery
        RstHelp.Requery
        If RstMain.RecordCount >= XBM Then
            RstMain.Bookmark = XBM
        Else
            If RstMain.EOF = False Then RstMain.MoveLast
        End If
        Call MoveRec
        BUTTONS True, Me, RstMain, 0
    End If
Exit Sub

ErrLoop:
    If err.NUMBER <> 0 Then
       GCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
    End If
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
    GSQL = "SELECT PART_NO as searchcode,Part_No,Part_Name,Local_Name FROM PART where div_code ='" & PubDivCode & "' order by PART.PART_NO"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
    Dim X1
    Dim mQry$, mRepName$
    Dim Rst As ADODB.Recordset
    Dim I   As Integer
    On Error GoTo ELoop
                
    
    mQry = "Select * From Part Order By Part_No"
               
               
    mRepName = "PartMaster"
    Set Rst = GCn.Execute(mQry)
    X1 = CreateFieldDefFile(Rst, PubRepoPath + "\" & mRepName & ".TTX", True)
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Report_View rpt, "Part Master", 0, False
    Set Rst = Nothing
Exit Sub
ELoop:
    MsgBox err.Description
End Sub

Private Sub TopCtrl1_eSave()
Dim transFlag As Byte, j As Integer, I As Integer, mVNo$, RST1 As ADODB.Recordset, mDocId$
Dim mModel_Grp_Code$, mVeh_Type$, mPart_OEM$, mSupl_Loca$
Dim Aggregate_Grp_Code$, mOpTrn As Boolean
On Error GoTo ErrLoop
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    If TxtGrid(1).Visible = True Then
        If TxtGridLeave1 = False Then
            TxtGrid(1).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    transFlag = 0
    If IsValid(txt(Part_No), "Part No.") = False Then Txt_GotFocus Part_No: Exit Sub
    If IsValid(txt(Part_Name), "Part Name") = False Then Txt_GotFocus Part_Name: Exit Sub
    If IsValid(txt(Security_Grade), "Security Grade") = False Then Txt_GotFocus Security_Grade: Exit Sub
    If IsValid(txt(Part_Grade_Name), "Proprietary Grade") = False Then Txt_GotFocus Part_Grade_Name: Exit Sub
    
    If IsValid(txt(Disc_Factor), "Discount Factor") = False Then Txt_GotFocus Disc_Factor: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From PART Where PART_NO=" & Chk_Text(Trim(txt(Part_No))) & " and div_code = '" & PubDivCode & "'").Fields(0) > 0 Then MsgBox "Part No. Already Exists for this division", vbInformation, "Part Validation": Txt_GotFocus Part_No: txt(Part_No).SetFocus: Exit Sub
    mModel_Grp_Code = "":   Aggregate_Grp_Code = ""
    mVeh_Type = "":         mPart_OEM = "":     mSupl_Loca = ""
    
    If Len(mSupl_Loca) > 28 Then MsgBox "Max 7 Supply Location's allowed, Unselect ", vbOKOnly, "Supply Location Selection": Exit Sub
    
    GCn.BeginTrans
    transFlag = 1
    If ADDFLAG = 1 Then
        GCn.Execute ("Insert Into PART (Div_Code,Site_Code,U_Name,U_EntDt,U_AE,PART_NO,Part_Name,Local_Name,Part_NoHelp,Part_NameHelp,UNIT,MARK_YN,Part_Grade,Security_Grade,Active_YN,Value_Method,Lead_Time,Min_Lvl,Max_Lvl,ReOrd_Lvl,Disc_Factor,Bin_Loca,MRP,TB_SRate,TP_SRate,Cur_TP_STk,Cur_TB_STk,Cur_MRP_TPSTk,Cur_MRP_TBSTk, NDP,Dep_Item ) " & _
            "  Values('" & PubDivCode & "','" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & _
            "','" & txt(Part_No) & "','" & txt(Part_Name) & "','" & txt(Local_Name) & "','" & Replace(txt(Part_No), " ", "") & "','" & Replace(txt(Part_Name), " ", "") & _
            "','" & txt(Unit) & "','" & left(txt(MARK_YN), 1) & "','" & txt(Part_Grade) & "','" & txt(Security_Grade) & "',1 " & _
            " ,'" & txt(Value_Method) & "'," & Val(txt(Lead_Time)) & "," & Val(txt(Min_Lvl)) & "," & Val(txt(Max_Lvl)) & "," & Val(txt(ReOrd_Lvl)) & _
            ",'" & txt(Disc_Factor) & "','" & txt(Bin_Loca) & "'," & Val(txt(MRP)) & "," & Val(txt(TB_SRate)) & "," & Val(txt(TP_SRate)) & _
            ",0,0,0,0, " & Val(txt(NDP)) & ",'" & txt(Dep_Item).Tag & "' )")
    Else
        Set GRs = GCn.Execute("select docid From SP_STOCK Where V_Type='" & mVType & "' AND Part_No='" & txt(Part_No) & "' and left(docid,1)='" & PubDivCode & "'")
        If GRs.RecordCount > 0 Then
            mVNo = GRs!DocID
            UpdStkTableToTable mVNo, "-", "R"
        End If
        Set GRs = Nothing
        GCn.Execute ("DELETE From SP_STOCK Where V_Type='" & mVType & "' AND Part_No='" & txt(Part_No) & "' and left(docid,1) = '" & PubDivCode & "' And " & cMID("DocId", "3", "1") & "='" & PubSiteCode & "' ")
        GCn.Execute ("UPDATE Part SET Site_Code='" & PubSiteCode & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'," & _
            "Part_Name='" & txt(Part_Name) & "',Local_Name='" & txt(Local_Name) & "',Part_NameHelp='" & Replace(txt(Part_Name), " ", "") & _
            "',UNIT='" & txt(Unit) & "',MARK_YN='" & left(txt(MARK_YN), 1) & "',Part_Grade='" & txt(Part_Grade) & "',Security_Grade='" & txt(Security_Grade) & _
            "',Active_YN=1,Value_Method='" & txt(Value_Method) & "',Lead_Time=" & Val(txt(Lead_Time)) & _
            ",Min_Lvl=" & Val(txt(Min_Lvl)) & ",Max_Lvl=" & Val(txt(Max_Lvl)) & ",ReOrd_Lvl=" & Val(txt(ReOrd_Lvl)) & ",Disc_Factor='" & txt(Disc_Factor) & _
            "',Bin_Loca='" & txt(Bin_Loca) & "',MRP=" & Val(txt(MRP)) & ",TB_SRate=" & Val(txt(TB_SRate)) & ", NDP=" & Val(txt(NDP)) & ",Dep_Item='" & txt(Dep_Item).Tag & "',TP_SRate=" & Val(txt(TP_SRate)) & _
            " Where PART_NO = '" & txt(Part_No) & "' and div_code = '" & PubDivCode & "'")
    End If
    GCn.Execute ("DELETE From Part_Alternate Where Root_Part_No='" & txt(Part_No) & "'  and div_code = '" & PubDivCode & "'")
    For j = 1 To FGrid.Rows - 1 'ALTERNATE PART
        If FGrid.TextMatrix(j, 1) <> "" Then GCn.Execute ("INSERT INTO Part_Alternate (Root_Part_No,Alternate_Part_No,Div_Code,Site_Code,U_Name,U_EntDt,U_AE) VALUES ('" & txt(Part_No) & "'," & Chk_Text(FGrid.TextMatrix(j, 1)) & ",'" & PubDivCode & "','" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    Next
    
    For j = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(j, OQuan) <> "" Then
            mOpTrn = True
            Exit For
        End If
    Next
    
    For j = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(j, ODate) <> "" And FGrid1.TextMatrix(j, OLastInvDt) <> "" Then
            If CDate(FGrid1.TextMatrix(j, ODate)) < CDate(FGrid1.TextMatrix(j, OLastInvDt)) Then
                MsgBox "Last Inv Date is greater than Opening Date "
                FGrid1.Row = j
                FGrid1.Col = OLastInvDt
                FGrid1.SetFocus
                Exit Sub
            End If
        End If
    Next
    
    If mOpTrn Then
        mVNo = GetDocID(GCnFaS, mVType, PubStartDate, VoucherEditFlag, txt(SerialNo), Label3(16), PubSiteCode)
        For j = 1 To FGrid1.Rows - 1
            If FGrid1.TextMatrix(j, OQuan) <> "" Then            'MRP TB Qty (kapil 28/07/2003 following Pub)
                GCn.Execute "INSERT INTO SP_STOCK (DocID,Srl_No,V_Type,V_No,V_Date,Part_No,Godown,remark,MRP_YN,TAX_YN,Qty_Rec,Rate,Amount,MRP_Rate,Site_Code,U_Name,U_EntDt,U_AE,V_Rate,PurDocNo,PurDocDate) " & _
                    "VALUES ('" & mVNo & "'," & j & ",'" & mVType & "'," & txt(SerialNo) & "," & ConvertDate(PubStartDate - 1) & ",'" & txt(Part_No) & "','" & FGrid1.TextMatrix(j, OGodCode) & _
                    "','" & FGrid1.TextMatrix(j, ORefNo) & "'," & IIf(FGrid1.TextMatrix(j, OMRPYN) = "Yes", 1, 0) & "," & IIf(FGrid1.TextMatrix(j, OTAXYN) = "Yes", 1, 0) & "," & Val(FGrid1.TextMatrix(j, OQuan)) & _
                    "," & Val(FGrid1.TextMatrix(j, ORate)) & "," & Val(FGrid1.TextMatrix(j, OVal)) & "," & IIf(FGrid1.TextMatrix(j, OMRPYN) = "Yes", Val(FGrid1.TextMatrix(j, ORate)), 0) & _
                    ",'" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "'," & Val(FGrid1.TextMatrix(j, ORate)) & ",'" & FGrid1.TextMatrix(j, OLastInvNo) & "'," & ConvertDate(FGrid1.TextMatrix(j, OLastInvDt)) & ")"
        'modi lps at Cuttack 02.09.03
                Call UpdStkGridToTable(txt(Part_No), "+", FGrid1.TextMatrix(j, OMRPYN), FGrid1.TextMatrix(j, OTAXYN), FGrid1.TextMatrix(j, OQuan))
            End If
        Next
        UpdVouSrlNo GCnFaS, mVNo, PubStartDate
    End If
    GCn.CommitTrans
    transFlag = 0
    rsPartRefresh = True
    If PubMoveRecYn Then
        RsPartSiteWise.Requery
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select P.Part_No as Code, P.Part_No, P.Part_Name As Name, P.Local_Name as LName, " & _
                                  "P.Unit , P.MRP, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, (Select PurcDisc_Per From Part_DiscFactor Where DiscFac_Catg=P.Disc_Factor) As PurcDisc_Per, P.ReOrd_Lvl, " & _
                                  "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=1 And Tax_Yn=1) As Cur_MRP_TBStk, " & _
                                  "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=1 And Tax_Yn=0) As Cur_MRP_TpStk, " & _
                                  "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "  Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=0 And Tax_Yn=1) As Cur_TB_Stk, " & _
                                  "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=0 And Tax_Yn=0) As Cur_Tp_Stk, " & _
                                  "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No ) As CurrStk, " & _
                                  "P.Min_Lvl, P.Disc_Factor " & _
                                  "From Part P " & _
                                  "Left Join Sp_Stock Stk On P.Part_No=Stk.Part_No " & _
                                  "WHERE  Div_Code='C' And P.Part_No = " & Chk_Text(txt(Part_No)) & " " & _
                                  "Group By P.Part_No, P.Part_Name, P.Local_Name, P.Unit, P.Mrp, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Disc_Factor, P.ReOrd_Lvl, P.Min_Lvl")
    
    End If
    
    If MasterFormExit Then Unload Me: Exit Sub
    RstMain.FIND ("Code=" & Chk_Text(txt(Part_No)))
    If ADDFLAG = 1 Then
        BlankText
        Txt_GotFocus Part_No
        txt(Part_No).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        ADDFLAG = 0
        'LV(4).Visible = False
    End If
    Set RST1 = Nothing
Exit Sub
ErrLoop:    If transFlag = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ErrLoop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        Grid_Hide
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        If TxtGrid(0).Visible = True Then TxtGrid(0).Visible = False
        If TxtGrid(1).Visible = True Then TxtGrid(1).Visible = False
    End If
Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eRef()
    RstHelp.Requery
    RstPartGrade.Requery
    RstDiscFact.Requery
    RstDEp_item.Requery
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
On Error GoTo ELoop
Dim I As Integer
Dim XXA() As String
TxtGrid(0).Visible = False
TxtGrid(1).Visible = False
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case Part_No
        RstHelp.Sort = "Code"
    Case Part_Name
        RstHelp.Sort = "Name"
    Case Local_Name
        RstHelp.Sort = "LName"
    Case Value_Method
        ListArray = Array("FIFO", "LIFO", "RAR", "FAR", "LPR")
        Set mListItem = ListView_Items(ListView, txt, Value_Method, ListArray, 5)
    Case Disc_Factor
        If RstDiscFact.RecordCount = 0 Or (RstDiscFact.EOF = True Or RstDiscFact.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RstDiscFact!Code Then
            RstDiscFact.MoveFirst
            RstDiscFact.FIND "Code ='" & Trim(txt(Index).TEXT) & "'"
        End If
    Case Part_Grade_Name
        If RstPartGrade.RecordCount = 0 Or (RstPartGrade.EOF = True Or RstPartGrade.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RstPartGrade!Name Then
            RstPartGrade.MoveFirst
            RstPartGrade.FIND "name ='" & txt(Index).TEXT & "'"
        End If
        'Nikhil
    Case Dep_Item
        If RstDEp_item.RecordCount = 0 Or (RstDEp_item.EOF = True Or RstDEp_item.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RstDEp_item!Name Then
            RstDEp_item.MoveFirst
            RstDEp_item.FIND "name ='" & txt(Index).TEXT & "'"
        End If
        
        
    Case Security_Grade
        ListArray = Array("A", "B", "C")
        Set mListItem = ListView_Items(ListView, txt, Security_Grade, ListArray, 3)
    Case Unit
        SSTab1.Tab = 0
        txt(Unit).Tag = txt(Unit)
        'Case Unit
        Set RstUnit = New ADODB.Recordset
        With RstUnit
             .CursorLocation = adUseClient
             .Open "SELECT Unit_Name  from unit order by unit_name", GCn, adOpenDynamic, adLockOptimistic
        End With
        Do While Not RstUnit.EOF
            I = I
            ReDim Preserve XXA(I)
            XXA(I) = RstUnit!Unit_Name
            I = I + 1
            RstUnit.MoveNext
        Loop
        urec = RstUnit.RecordCount
        Set mListItem = ListView_Items(ListView, txt, Unit, XXA, RstUnit.RecordCount)
        'Set rsdes = Nothing
    Case MRP
        If OldRateMRP = 0 Then
            OldRateMRP = Val(txt(Index))
        End If
    Case TB_SRate
        If OldRateTB = 0 Then
            OldRateTB = Val(txt(Index))
        End If
    Case TP_SRate
        If OldRateTP = 0 Then
            OldRateTP = Val(txt(Index))
        End If

'    Case Lead_Time
'        SSTab1.Tab = 1
End Select
Exit Sub
ELoop:
MsgBox err.Description, vbCritical: Exit Sub
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case Part_No
        If DGPart.Visible = False Then DGridColSwap DGPart, 0
        DGridTxtKeyDown_Mast DGPart, txt, Index, RstHelp, KeyCode, True, 0
    Case Part_Name
        If DGPart.Visible = False Then DGridColSwap DGPart, 1
        DGridTxtKeyDown_Mast DGPart, txt, Index, RstHelp, KeyCode, True, 1
    Case Local_Name
        If DGPart.Visible = False Then DGridColSwap DGPart, 5
        DGridTxtKeyDown_Mast DGPart, txt, Index, RstHelp, KeyCode, True, 5
    Case Value_Method
        ListView_KeyDown FrmList, ListView, txt, Value_Method, KeyCode, Shift, txt(Value_Method).left, (txt(Value_Method).top + txt(Value_Method).height), txt(Value_Method).width, 260 * 5
    Case Security_Grade
        ListView_KeyDown FrmList, ListView, txt, Security_Grade, KeyCode, Shift, txt(Security_Grade).left, txt(Security_Grade).top + txt(Security_Grade).height, txt(Security_Grade).width, 260 * 3
    Case Unit
        ListView_KeyDown FrmList, ListView, txt, Unit, KeyCode, Shift, txt(Unit).left, txt(Unit).top + txt(Unit).height, txt(Unit).width, 260 * urec
    Case Part_Grade_Name
        DGridTxtKeyDown DGPartGrade, txt, Index, RstPartGrade, KeyCode, False, 1, frmPartGrade, "frmPartGrade"
        'Nikhil
    Case Dep_Item
        DGridTxtKeyDown DGDep_Item, txt, Index, RstDEp_item, KeyCode, False, 1, FrmDeprecation_itemMaster, "FrmDeprecation_itemMaster"
        
    Case Disc_Factor
        DGridTxtKeyDown DGDisFact, txt, Index, RstDiscFact, KeyCode, False, 0, frmPartDiscFact, "frmPartDiscFact"
End Select
If FrmList.Visible = False And DGPart.Visible = False And DGPartGrade.Visible = False And DGDep_Item.Visible = False And DGDisFact.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> Part_No Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> Part_Name Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Dim mUChr As String
mUChr = UCase(Chr(KeyAscii))
Select Case Index
    Case Lead_Time
        NumPress txt(Index), KeyAscii, 3, 0
    Case Min_Lvl, Max_Lvl, ReOrd_Lvl
        NumPress txt(Index), KeyAscii, 9, 2
    Case MRP, TB_SRate, TP_SRate, NDP
        NumPress txt(Index), KeyAscii, 9, 3
    Case Curr_Stock, Curr_Stock_Value
        NumPress txt(Index), KeyAscii, 7, 3
    Case Part_Grade_Name
        If DGPartGrade.Visible = True Then DGridTxtKeyPress txt, Index, RstPartGrade, KeyAscii, "name"
    'Nikhil
      Case Dep_Item
        If DGDep_Item.Visible = True Then DGridTxtKeyPress txt, Index, RstDEp_item, KeyAscii, "name"
        
    Case Disc_Factor
        If DGDisFact.Visible = True Then DGridTxtKeyPress txt, Index, RstDiscFact, KeyAscii, "code"
    Case MARK_YN
        If mUChr = "N" Then
            txt(Index) = "No"
        ElseIf mUChr = "Y" Then
            txt(Index) = "Yes"
        End If
        KeyAscii = 0
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case Part_No
        DGridTxtKeyUp_Mast txt, Index, RstHelp, KeyCode, "Code"
    Case Part_Name
       If DGPart.Visible = True Then DGridTxtKeyUp_Mast txt, Index, RstHelp, KeyCode, "Name"
    Case Local_Name
        If DGPart.Visible = True Then DGridTxtKeyUp_Mast txt, Index, RstHelp, KeyCode, "LName"
    Case Value_Method
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Value_Method, KeyCode, mListItem
    Case Security_Grade
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Security_Grade, KeyCode, mListItem
    Case Unit
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Unit, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim NewRate As Double
'Dim Rst As ADODB.Recordset
Select Case Index
    Case MRP
        txt(Index).TEXT = Format(txt(Index), "0.000")
        NewRate = Val(txt(Index).TEXT)
        If RSOJPR = True Then
            If OldRateMRP > 0 Then
                DiffRate = Format((OldRateMRP * 25) / 100, "0.000")
                If NewRate > (OldRateMRP + DiffRate) Or NewRate < (OldRateMRP - DiffRate) Then
                    MsgBox "Rate change limit exceeded.Change can not be made.", vbInformation
                    txt(Index) = OldRateMRP
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    Case TB_SRate
        txt(Index).TEXT = Format(txt(Index), "0.000")
        NewRate = Val(txt(Index).TEXT)
        If RSOJPR = True Then
            If OldRateTB > 0 Then
                DiffRate = Format((OldRateTB * 25) / 100, "0.000")
                If NewRate > (OldRateTB + DiffRate) Or NewRate < (OldRateTB - DiffRate) Then
                    MsgBox "Rate change limit exceeded.Change can not be made.", vbInformation
                    txt(Index) = OldRateTB
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    Case TP_SRate
        txt(Index).TEXT = Format(txt(Index), "0.000")
        NewRate = Val(txt(Index).TEXT)
        If RSOJPR = True Then
            If OldRateTP > 0 Then
                DiffRate = Format((OldRateTP * 25) / 100, "0.000")
                If NewRate > (OldRateTP + DiffRate) Or NewRate < (OldRateTP - DiffRate) Then
                    MsgBox "Rate change limit exceeded.Change can not be made.", vbInformation
                    txt(Index) = OldRateTP
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    Case Lead_Time, Curr_Stock, Curr_Stock_Value, Min_Lvl, Max_Lvl, ReOrd_Lvl, NDP
        txt(Index).TEXT = Format(txt(Index), "0.00")
    Case Part_No
        If GCn.Execute("select count(*) from part where part_no='" & txt(Part_No).TEXT & "' and div_code = '" & PubDivCode & "'").Fields(0) > 0 Then
            MsgBox "Duplicate Part  No.", vbCritical, "Validation Error"
            txt(Index).TEXT = ""
            Cancel = True
            Exit Sub
        End If
        'Nikhil
    Case Dep_Item
        txt(Dep_Item) = RstDEp_item!Name
        txt(Dep_Item).Tag = RstDEp_item!Code
        
    Case Part_Grade_Name
        txt(Part_Grade_Name) = RstPartGrade!Name
        txt(Part_Grade) = RstPartGrade!Code
End Select
'Set Rst = Nothing
End Sub

Private Sub LV_Click(Index As Integer)
Select Case Index
    Case lvStockValuation
        Txt_Validate Value_Method, True
End Select
End Sub

Private Sub LV_GotFocus(Index As Integer)
'   CtrlClckCol
Select Case Index
    Case lvOEM
        SSTab1.Tab = 0
    Case lvModelGroup
        SSTab1.Tab = 3
End Select
End Sub


Private Sub DGPartGrade_GotFocus()
    mFLAG1 = 1
End Sub


Private Sub DGDep_Item_GotFocus()
    mFLAG1 = 1
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
Ctrl_GetFocus TxtGrid(Index)
Grid_Hide
Select Case Index
    Case 0
'        FGrid.CellBackColor = CellBackColLeave
    Case 1
'        FGrid1.CellBackColor = CellBackColLeave
End Select
Select Case Index
    Case 0  'Alternate No. Grid
        TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
        Select Case FGrid.Col
            Case AltNo
                If RstHelp.RecordCount = 0 Or (RstHelp.EOF = True And RstHelp.BOF = True) Or FGrid.TextMatrix(FGrid.Row, AltNo) = "" Then Exit Sub
                RstHelp.Sort = "Code"
                RstHelp.MoveFirst
                RstHelp.FIND "Code ='" & FGrid.TextMatrix(FGrid.Row, AltNo) & "'"
            Case StaPart
                If RstHelp.RecordCount = 0 Or (RstHelp.EOF = True And RstHelp.BOF = True) Or FGrid.TextMatrix(FGrid.Row, StaPart) = "" Then Exit Sub
                RstHelp.Sort = "Name"
                RstHelp.MoveFirst
                RstHelp.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, StaPart) & "'"
                If RstHelp.EOF = True Then RstHelp.MoveFirst
            Case LocPart
                If RstHelp.RecordCount = 0 Or (RstHelp.EOF = True And RstHelp.BOF = True) Or FGrid.TextMatrix(FGrid.Row, LocPart) = "" Then Exit Sub
                RstHelp.Sort = "lname"
                RstHelp.MoveFirst
                RstHelp.FIND "lname ='" & FGrid.TextMatrix(FGrid.Row, LocPart) & "'"
                If RstHelp.EOF = True Then RstHelp.MoveFirst
        End Select
    Case 1  'Opening Stock Grid
        Select Case FGrid1.Col
            Case OGodown
                If RstGod.RecordCount = 0 Or (RstGod.EOF = True Or RstGod.BOF = True) Then Exit Sub
                If FGrid1.TextMatrix(FGrid1.Row, OGodown) = "" Then
                    RstGod.Sort = "Name"
                Else
                    RstGod.Sort = "Name"
                    RstGod.MoveFirst
                    RstGod.FIND "name ='" & FGrid1.TextMatrix(FGrid1.Row, OGodown) & "'"
                    If RstGod.EOF = True Then RstGod.MoveFirst
                End If
        End Select
End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrLoop
If KeyCode = vbKeyEscape Then TxtGrid(Index).TEXT = TxtGrid(Index).Tag: Exit Sub
Select Case Index
    Case 0  'Alternate No. Grid
        Select Case FGrid.Col
            Case AltNo    '1
                If DGPart.Visible = False Then DGridColSwap DGPart, 0
                DGridTxtKeyDown DGPart, TxtGrid, Index, RstHelp, KeyCode, True, 0, frmPartMast, "frmPartMast"
                If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, True, 7, 7
                        Else
                            TxtGrid_LostFocus 0
                            TxtGrid(0).SetFocus
                        End If
                End If
            Case StaPart
                If DGPart.Visible = False Then DGridColSwap DGPart, 1
                DGridTxtKeyDown DGPart, TxtGrid, Index, RstHelp, KeyCode, True, 1, frmPartMast, "frmPartMast"
                If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 2
                        Else
                            TxtGrid_LostFocus 0
                            TxtGrid(0).SetFocus
                        End If
                End If
            Case LocPart   '3
                If DGPart.Visible = False Then DGridColSwap DGPart, 5
                DGridTxtKeyDown DGPart, TxtGrid, Index, RstHelp, KeyCode, True, 5, frmPartMast, "frmPartMast"
                If KeyCode = vbKeyReturn Then
                        If TxtGridLeave = True Then
                             GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 7, , LocPart
                        Else
                            TxtGrid_LostFocus 0
                            TxtGrid(0).SetFocus
                        End If
                End If
        End Select
    Case 1  'Opening Stock Grid
        Select Case FGrid1.Col
            Case OGodown
                If DGGod.Visible = False Then DGridColSwap DGGod, 1
                DGridTxtKeyDown DGGod, TxtGrid, Index, RstGod, KeyCode, True, 1, frmGodown, "frmGodown"
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave1 = True Then
                        GridTxtDown FGrid1, TxtGrid, 1, KeyCode, TAddMode, 8
                    Else
                        TxtGrid_LostFocus 1
                        TxtGrid(1).SetFocus
                    End If
                End If
            Case ODate, ORefNo, OMRPYN, OTAXYN, OQuan, OVal
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave1 = True Then
                         GridTxtDown FGrid1, TxtGrid, 1, KeyCode, TAddMode, 8, 0
    ' Nra Updation
                         'FGrid1.TextMatrix(FGrid1.Row, ORate) = Tax_Calc(IIf(FGrid1.TextMatrix(FGrid1.Row, OMRPYN) = "", 0, Val(IIf(FGrid1.TextMatrix(FGrid1.Row, OMRPYN) = "Yes", 1, 0))), IIf(FGrid1.TextMatrix(FGrid1.Row, OTAXYN) = "", 0, Val(IIf(FGrid1.TextMatrix(FGrid1.Row, OTAXYN) = "Yes", 1, 0))))
    ' End update
                    Else
                        TxtGrid_LostFocus 0
                        TxtGrid(1).SetFocus
                    End If
                End If
                If Index <> 1 And DGGod.Visible = False Then
                    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
                End If
            Case ORate, OLastInvNo, OLastInvDt
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave1 = True Then
                         GridTxtDown FGrid1, TxtGrid, 1, KeyCode, TAddMode, 10, 0
                    Else
                        TxtGrid_LostFocus 0
                        TxtGrid(1).SetFocus
                    End If
                End If
            
        End Select
End Select
ErrLoop:
CheckError
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If KeyAscii = vbKeyEscape Then Exit Sub
Dim mUChr As String
mUChr = UCase(Chr(KeyAscii))
Call CheckQuote(KeyAscii)
Select Case Index
    Case 0
        Select Case FGrid.Col
            Case AltNo
                If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RstHelp, KeyAscii, "part_no"
            Case StaPart
                If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RstHelp, KeyAscii, "part_name"
            Case LocPart
                If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RstHelp, KeyAscii, "Local_name"
        End Select
    Case 1
        Select Case FGrid1.Col
            Case OGodown
                If DGGod.Visible = True Then DGridTxtKeyPress TxtGrid, 1, RstGod, KeyAscii, "name"
            Case OMRPYN, OTAXYN
                If mUChr = "N" Then
                    TxtGrid(Index) = "No"
                ElseIf mUChr = "Y" Then
                    TxtGrid(Index) = "Yes"
                End If
                KeyAscii = 0
        End Select
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case 0
        Select Case FGrid.Col
            Case AltNo
                If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
                DGridTxtKeyPress TxtGrid, Index, RstHelp, KeyCode, "Code"
            Case StaPart
                If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
                DGridTxtKeyPress TxtGrid, Index, RstHelp, KeyCode, "name"
            Case LocPart
                If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
                DGridTxtKeyPress TxtGrid, Index, RstHelp, KeyCode, "Local_name"
        End Select
    Case 1  'Opening Stock Grid
        Select Case FGrid1.Col
            Case ODate
                FGrid1.TextMatrix(FGrid1.Row, ODate) = TxtGrid(1).TEXT
            Case OGodown
                If KeyCode <> 13 And DGGod.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, 1, RstGod, KeyCode, "name", True
            Case ORefNo
                FGrid1.TextMatrix(FGrid1.Row, ORefNo) = TxtGrid(1).TEXT
            Case OQuan
                FGrid1.TextMatrix(FGrid1.Row, OQuan) = Format(Val(TxtGrid(1).TEXT), "0.00")
            Case ORate
                FGrid1.TextMatrix(FGrid1.Row, ORate) = Format(Val(TxtGrid(1).TEXT), "0.00")
            Case OVal
                FGrid1.TextMatrix(FGrid1.Row, OVal) = Format(Val(TxtGrid(Index).TEXT), "0.00")
            Case OLastInvNo
                FGrid1.TextMatrix(FGrid1.Row, OLastInvNo) = TxtGrid(Index).TEXT
            Case OLastInvDt
                FGrid1.TextMatrix(FGrid1.Row, OLastInvDt) = RetDate(TxtGrid(Index))
                
        End Select
End Select
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    TxtGrid(Index).Visible = False
    Grid_Hide
End If
End Sub

Private Sub Grid_Hide()
    If DGPart.Visible = True Then DGPart.Visible = False
    If DGDisFact.Visible = True Then DGDisFact.Visible = False
    If DGPartGrade.Visible = True Then DGPartGrade.Visible = False
    If DGDep_Item.Visible = True Then DGDep_Item.Visible = False
    If DGGod.Visible = True Then DGGod.Visible = False
If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Function ChkDuplicate() As Boolean
Dim I As Integer
Dim X As String, Y As String
Dim Col1 As Byte, Col2 As Byte, Col3 As Byte
    Select Case FGrid.Col
    Case AltNo
        Col1 = AltNo
    Case StaPart
        Col1 = StaPart
    Case LocPart
        Col1 = LocPart
    End Select
    X = UCase(CStr(Trim(TxtGrid(0).TEXT)))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col1))))
        If X = Y And Y <> "" Then
            MsgBox "Duplicate Item Not Allowed", vbInformation, "Validation"
            TxtGrid(0).SetFocus
            Ctrl_GetFocus TxtGrid(0)
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function

Private Function TxtGridLeave() As Boolean
Dim j As Integer
Select Case FGrid.Col
        Case AltNo, StaPart, LocPart
            If ChkDuplicate = False Then TxtGridLeave = False: ExitCtrl = False: Exit Function
            If RstHelp.RecordCount = 0 Or (RstHelp.EOF = True Or RstHelp.BOF = True) Or TxtGrid(0).TEXT = "" Then
                FGrid.TextMatrix(FGrid.Row, AltNo) = ""
                FGrid.TextMatrix(FGrid.Row, StaPart) = ""
                FGrid.TextMatrix(FGrid.Row, LocPart) = ""
                FGrid.TextMatrix(FGrid.Row, MRPp) = "0.00"
                FGrid.TextMatrix(FGrid.Row, TxRate) = "0.00"
                FGrid.TextMatrix(FGrid.Row, TPRate) = "0.00"
                FGrid.TextMatrix(FGrid.Row, TPRate) = "0.00"
                FGrid.TextMatrix(FGrid.Row, Bin) = ""
            Else
                FGrid.TextMatrix(FGrid.Row, AltNo) = RstHelp!Code
                FGrid.TextMatrix(FGrid.Row, StaPart) = RstHelp!Name
                FGrid.TextMatrix(FGrid.Row, LocPart) = RstHelp!LName
                FGrid.TextMatrix(FGrid.Row, MRPp) = Format(RstHelp!TB_SRate, "0.00")
                FGrid.TextMatrix(FGrid.Row, TxRate) = Format(RstHelp!TP_SRate, "0.00")
                FGrid.TextMatrix(FGrid.Row, TPRate) = Format(RstHelp!MRP, "0.00")
                FGrid.TextMatrix(FGrid.Row, Bin) = IIf(IsNull(RstHelp!Bin_Loca), "", RstHelp!Bin_Loca)
            End If
End Select
    ExitCtrl = True
    TxtGridLeave = True
    TxtGrid(0).Visible = False
    FGrid.SetFocus
End Function

Private Function TxtGridLeave1() As Boolean
Dim j As Integer
Select Case FGrid1.Col
    Case OGodown, ORefNo
        'If ChkDuplicate = False Then TxtGridLeave1 = False: ExitCtrl = False: Exit Function
        FGrid1.TextMatrix(FGrid1.Row, OGodown) = RstGod!Name
        FGrid1.TextMatrix(FGrid1.Row, OGodCode) = RstGod!Code
    Case ODate
        TxtGrid(1).TEXT = RetDate(TxtGrid(1))
        FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = TxtGrid(1).TEXT
    Case ORefNo, OMRPYN, OTAXYN
        FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = TxtGrid(1).TEXT
    Case OQuan, ORate, OVal
        FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = Format(Val(TxtGrid(1).TEXT), "0.00")
        FGrid1.TextMatrix(FGrid1.Row, ORate) = Format(Val(FGrid1.TextMatrix(FGrid1.Row, OVal)) / Val(FGrid1.TextMatrix(FGrid1.Row, OQuan)), "0.00")
    Case OLastInvNo
        FGrid1.TextMatrix(FGrid1.Row, OLastInvNo) = TxtGrid(1).TEXT
    Case OLastInvDt
        FGrid1.TextMatrix(FGrid1.Row, OLastInvDt) = RetDate(TxtGrid(1))
End Select
ExitCtrl = True
TxtGridLeave1 = True
TxtGrid(1).Visible = False
FGrid1.SetFocus
End Function

Private Sub TxtGrid_LostFocus(Index As Integer)
If ExitCtrl = False Then Exit Sub
  Ctrl_validate TxtGrid(Index)
Select Case FGrid.Col
    Case OLastInvDt
        If CDate(FGrid1.TextMatrix(FGrid1.Row, ODate)) < CDate(FGrid1.TextMatrix(FGrid1.Row, OLastInvDt)) Then
            MsgBox "Invoice Date is > then Opening Date !"
            FGrid1.TextMatrix(FGrid1.Row, OLastInvDt) = ""
            Exit Sub
        End If
End Select
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case 1
        Select Case FGrid1.Col
        Case ODate
            TxtGrid(Index).TEXT = RetDate(TxtGrid(Index))
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = TxtGrid(Index).TEXT
        Case OQuan, ORate, OVal
            FGrid1.TextMatrix(FGrid1.Row, OVal) = Format(Val(FGrid1.TextMatrix(FGrid1.Row, OQuan)) * Val(FGrid1.TextMatrix(FGrid1.Row, ORate)), "0.00")
            If Val(FGrid1.TextMatrix(FGrid1.Row, OVal)) = 0 Then
                FGrid1.TextMatrix(FGrid1.Row, OVal) = ""
            End If
        End Select
        
    Case 0
        Select Case FGrid.Col
            Case AltNo, StaPart, LocPart
           If ChkDuplicate = False Then Cancel = True: Exit Sub
        End Select
End Select
End Sub
Private Sub BlankText()
Dim I As Byte
For I = 0 To 23
    txt(I).TEXT = ""
Next I
        FGrid.Rows = 1
        FGrid.AddItem ""
        FGrid.FixedRows = 1
        FGrid1.Rows = 1
        FGrid1.AddItem ""
        FGrid1.FixedRows = 1
        
End Sub

Private Sub FillListV()
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("CODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select P.Part_No as Code, P.Part_No, P.Part_Name As Name, P.Local_Name as LName, " & _
                                  "P.Unit , P.MRP, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, (Select PurcDisc_Per From Part_DiscFactor Where DiscFac_Catg=P.Disc_Factor) As PurcDisc_Per, P.ReOrd_Lvl, " & _
                                  "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=1 And Tax_Yn=1) As Cur_MRP_TBStk, " & _
                                  "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=1 And Tax_Yn=0) As Cur_MRP_TpStk, " & _
                                  "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & "  Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=0 And Tax_Yn=1) As Cur_TB_Stk, " & _
                                  "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No  And Mrp_Yn=0 And Tax_Yn=0) As Cur_Tp_Stk, " & _
                                  "(Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock WHERE (V_Type=(Case When V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)) & " Then 'SXAO'  End) Or V_Type<>(Case When V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate) & " Then 'SXAO'  End)) And Part_No=P.Part_No ) As CurrStk, " & _
                                  "P.Min_Lvl, P.Disc_Factor " & _
                                  "From Part P " & _
                                  "Left Join Sp_Stock Stk On P.Part_No=Stk.Part_No " & _
                                  "WHERE  Div_Code='C' And P.Part_No = '" & MyValue & "' " & _
                                  "Group By P.Part_No, P.Part_Name, P.Local_Name, P.Unit, P.Mrp, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Disc_Factor, P.ReOrd_Lvl, P.Min_Lvl")
    End If
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Grid_Ini()
'* Used for intialize grid columns
With FGrid
    .ColAlignmentFixed = flexAlignCenterCenter
    .ColAlignmentFixed(0) = flexAlignRightCenter
    .ColAlignment(AltNo) = flexAlignLeftCenter
    .ColAlignment(StaPart) = flexAlignLeftCenter
    .ColAlignment(LocPart) = flexAlignLeftCenter
    .ColAlignment(Bin) = flexAlignLeftCenter
    .ColWidth(9) = 0
    .ColWidth(8) = 0
End With
    With FGrid1
        .left = 0
        .RowHeightMin = PubGridRowHeight
        .Cols = 12
        .ColAlignmentFixed = flexAlignCenterCenter

        .ColAlignmentFixed(0) = flexAlignRightCenter
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 350
        
        .TextMatrix(0, ODate) = "Date"
        .ColAlignment(ODate) = flexAlignCenterCenter
        .ColWidth(ODate) = 1260
        .TextMatrix(0, OGodown) = "Godown" ' Location"
        .ColAlignment(OGodown) = flexAlignLeftCenter
        .ColWidth(OGodown) = 2505

        .TextMatrix(0, ORefNo) = "Particulars"
        .ColAlignment(ORefNo) = flexAlignLeftCenter
        .ColWidth(ORefNo) = 1515

        .TextMatrix(0, OMRPYN) = "MRP Y/N"
        .ColAlignment(OMRPYN) = flexAlignCenterCenter
        .ColWidth(OMRPYN) = 855

        .TextMatrix(0, OTAXYN) = "Tax Y/N"
        .ColAlignment(OTAXYN) = flexAlignCenterCenter
        .ColWidth(OTAXYN) = 855
        
        .TextMatrix(0, OQuan) = "Quantity"
        .ColAlignment(OQuan) = flexAlignRightCenter
        .ColWidth(OQuan) = 1005
        .TextMatrix(0, OVal) = "Value"
        .ColAlignment(ORate) = flexAlignRightCenter
        .ColWidth(ORate) = 1005
        
        .TextMatrix(0, ORate) = "Rate"
        .ColAlignment(OVal) = flexAlignRightCenter
        .ColWidth(OVal) = 1140
        .ColWidth(OGodCode) = 0
        
        .TextMatrix(0, OLastInvNo) = "Last Inv.No."
        .ColAlignment(OLastInvNo) = flexAlignRightCenter
        .ColWidth(OLastInvNo) = 1500
        
        .TextMatrix(0, OLastInvDt) = "Last Inv.DT."
        .ColAlignment(OLastInvDt) = flexAlignRightCenter
        .ColWidth(OLastInvDt) = 1500

    End With
    
DGPart.left = Me.left ' + 45
DGPart.top = 3795
DGDisFact.top = 400
DGDisFact.left = 7245
DGPartGrade.top = 400
DGPartGrade.left = 6720
DGGod.top = 400
DGGod.left = 6720

End Sub

Private Sub SetMaxLength()
Select Case FGrid1.Col
        Case ORefNo
            TxtGrid(1).MaxLength = 30
            TxtGrid(1).Alignment = 0
        Case Else
            TxtGrid(1).MaxLength = 0
    End Select
End Sub
' Nra Updation
Private Function Tax_Calc(MRPRate As Integer, TAX As Integer)
Set RstDiscFact = New ADODB.Recordset
RstDiscFact.CursorLocation = adUseClient
RstDiscFact.Open "Select SalDisc_Per as Discount From Part_DiscFactor where DiscFac_Catg='" & txt(Disc_Factor) & "'", GCn, adOpenDynamic, adLockOptimistic
    If MRPRate = 1 And TAX = 1 Then
         Tax_Calc = Val(txt(MRP)) - Val(txt(MRP)) * RstDiscFact!Discount / 100
    ElseIf MRPRate = 0 And TAX = 1 Then
         Tax_Calc = Val(txt(TB_SRate)) - Val(txt(TB_SRate)) * RstDiscFact!Discount / 100
    ElseIf MRPRate = 1 And TAX = 0 Then
         Tax_Calc = Val(txt(MRP)) - Val(txt(MRP)) * RstDiscFact!Discount / 100
    ElseIf MRPRate = 0 And TAX = 0 Then
         Tax_Calc = Val(txt(TP_SRate)) - Val(txt(TP_SRate)) * RstDiscFact!Discount / 100
    End If
End Function
' End update
