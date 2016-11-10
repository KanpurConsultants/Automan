VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVisit 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Daily Activity Entry"
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.ListBox LstMarket 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   2190
      ItemData        =   "frmVisit.frx":0000
      Left            =   8865
      List            =   "frmVisit.frx":001C
      Style           =   1  'Checkbox
      TabIndex        =   20
      Top             =   2700
      Width           =   2820
   End
   Begin MSDataGridLib.DataGrid DGMod 
      Height          =   2250
      Left            =   4770
      Negotiate       =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   4965
      Visible         =   0   'False
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   3969
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
      Left            =   5790
      TabIndex        =   6
      Top             =   1800
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
      Index           =   21
      Left            =   8865
      MaxLength       =   15
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2070
      Width           =   2820
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
      Left            =   8865
      MaxLength       =   15
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2820
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
      Left            =   8865
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "0123456789012345678901234"
      Top             =   1530
      Width           =   2820
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   315
      TabIndex        =   32
      Top             =   5820
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   225
         TabIndex        =   33
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
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDF4B5&
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
      Left            =   510
      TabIndex        =   48
      Top             =   4200
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSDataGridLib.DataGrid DGProsCust 
      Height          =   4935
      Left            =   4185
      Negotiate       =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5970
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
         DataField       =   "name"
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
         DataField       =   "NSuffix"
         Caption         =   "Sfix"
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
            ColumnWidth     =   3945.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   450.142
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2175
      Left            =   195
      TabIndex        =   21
      Top             =   4800
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   3836
      _Version        =   393216
      BackColorFixed  =   12632319
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   12640511
      GridColorFixed  =   12640511
      GridColorUnpopulated=   13623520
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Left            =   7140
      MaxLength       =   6
      TabIndex        =   17
      Top             =   3960
      Width           =   615
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
      IMEMode         =   3  'DISABLE
      Index           =   18
      Left            =   2970
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1050
      Width           =   1245
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
      Left            =   3900
      MaxLength       =   6
      TabIndex        =   12
      Top             =   2880
      Width           =   615
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
      Left            =   2970
      MaxLength       =   6
      TabIndex        =   11
      Top             =   2880
      Width           =   615
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
      Left            =   9945
      MaxLength       =   12
      TabIndex        =   43
      Text            =   "VFalse"
      Top             =   390
      Visible         =   0   'False
      Width           =   1530
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9000
      MaxLength       =   12
      TabIndex        =   1
      Top             =   780
      Width           =   1380
   End
   Begin VB.TextBox txt 
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
      Index           =   3
      Left            =   2970
      MaxLength       =   40
      TabIndex        =   2
      Top             =   780
      Width           =   4785
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
      Left            =   2970
      MaxLength       =   12
      TabIndex        =   16
      Top             =   3960
      Width           =   1380
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
      Left            =   2970
      MaxLength       =   50
      TabIndex        =   19
      Top             =   4500
      Width           =   4785
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
      Left            =   2970
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "YES"
      Top             =   2070
      Width           =   570
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
      Left            =   2970
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1530
      Width           =   4785
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
      Left            =   2970
      MaxLength       =   50
      TabIndex        =   13
      Top             =   3150
      Width           =   4785
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
      Left            =   5310
      MaxLength       =   20
      TabIndex        =   10
      Text            =   "01234567890123456789"
      Top             =   2610
      Width           =   2445
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   14
      Left            =   2970
      MaxLength       =   8
      TabIndex        =   18
      Text            =   "99999.99"
      Top             =   4230
      Width           =   930
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
      Left            =   2970
      MaxLength       =   50
      TabIndex        =   14
      Top             =   3420
      Width           =   4785
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
      Index           =   13
      Left            =   2970
      MaxLength       =   5
      TabIndex        =   15
      Top             =   3690
      Width           =   1380
   End
   Begin VB.TextBox txt 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   9000
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1050
      Width           =   705
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
      Left            =   2970
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1800
      Width           =   570
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
      Left            =   2970
      MaxLength       =   40
      TabIndex        =   8
      Top             =   2340
      Width           =   4785
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
      Left            =   2970
      MaxLength       =   6
      TabIndex        =   9
      Top             =   2610
      Width           =   990
   End
   Begin MSDataGridLib.DataGrid DGRep 
      Height          =   4515
      Left            =   5580
      Negotiate       =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6225
      Visible         =   0   'False
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   7964
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
   Begin MSDataGridLib.DataGrid dGObj 
      Height          =   4515
      Left            =   4800
      Negotiate       =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   7964
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Objective"
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
   Begin MSDataGridLib.DataGrid DgParty 
      Height          =   4515
      Left            =   3060
      Negotiate       =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   6435
      Visible         =   0   'False
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   7964
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
      ColumnCount     =   1
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Also Collected Market Information : "
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
      Left            =   7920
      TabIndex        =   57
      Top             =   2415
      Width           =   2880
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
      Index           =   9
      Left            =   4335
      TabIndex        =   55
      Top             =   1815
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profession"
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
      Left            =   7890
      TabIndex        =   54
      Top             =   2085
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
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
      Left            =   8415
      TabIndex        =   52
      Top             =   1815
      Width           =   375
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   15
      Left            =   8490
      TabIndex        =   51
      Top             =   1530
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Meeting Time"
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
      Index           =   6
      Left            =   5460
      TabIndex        =   47
      Top             =   3960
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Index           =   5
      Left            =   1275
      TabIndex        =   46
      Top             =   1065
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Index           =   4
      Left            =   3660
      TabIndex        =   45
      Top             =   2880
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Meeting Time  From"
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
      Index           =   2
      Left            =   1275
      TabIndex        =   39
      Top             =   2880
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activity Date"
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
      Index           =   26
      Left            =   7905
      TabIndex        =   38
      Top             =   795
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transfered Y/N"
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
      Left            =   1275
      TabIndex        =   37
      Top             =   2085
      Width           =   1230
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
      Index           =   3
      Left            =   1275
      TabIndex        =   36
      Top             =   795
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Remarks"
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
      Left            =   1275
      TabIndex        =   35
      Top             =   4515
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Visit/Call Date"
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
      Left            =   1275
      TabIndex        =   34
      Top             =   3960
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Rs."
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
      Left            =   1275
      TabIndex        =   31
      Top             =   4245
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
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
      Index           =   32
      Left            =   1275
      TabIndex        =   30
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Todays Remark"
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
      Index           =   38
      Left            =   1260
      TabIndex        =   29
      Top             =   3165
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Enquiry Y/N"
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
      Left            =   1275
      TabIndex        =   28
      Top             =   1815
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transfered From"
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
      Index           =   46
      Left            =   1275
      TabIndex        =   27
      Top             =   2355
      Width           =   1380
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
      Index           =   31
      Left            =   1275
      TabIndex        =   26
      Top             =   3705
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Objective"
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
      Index           =   33
      Left            =   4455
      TabIndex        =   24
      Top             =   2625
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No"
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
      Index           =   27
      Left            =   8190
      TabIndex        =   23
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visit / Call "
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
      Left            =   1275
      TabIndex        =   22
      Top             =   2625
      Width           =   855
   End
End
Attribute VB_Name = "frmVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim ExitCtrl As Boolean

Dim RsRep As ADODB.Recordset
Dim RsParty As ADODB.Recordset
Dim rsProsCust As ADODB.Recordset
Dim rsObj As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim RsMod As ADODB.Recordset
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Dim ListArray As Variant
Dim mListItem As ListItem

Private Const VisitDate As Byte = 1
Private Const Srlno As Byte = 2
Private Const REP_CODE As Byte = 3
Private Const NewEnquiry As Byte = 4
Private Const Party_code As Byte = 5
Private Const Trf_YN As Byte = 6
Private Const TrfFrom_RepCode As Byte = 7
Private Const Visit_Call As Byte = 8
Private Const Next_Date As Byte = 9
Private Const Remark1 As Byte = 10
Private Const Remark2 As Byte = 11
Private Const Objective As Byte = 12
Private Const Call_Status As Byte = 13
Private Const Expence As Byte = 14
Private Const ExpRemark As Byte = 15
Private Const Meet_TimeFrom As Byte = 16
Private Const Meet_TimeTo As Byte = 17
Private Const RepPWD As Byte = 18
Private Const Next_Time As Byte = 19
Private Const Area As Byte = 20
Private Const Profession As Byte = 21
Private Const City As Byte = 22
Private Const Model As Byte = 23

'Grid Columns
Private Const PartyCode As Byte = 1
Private Const Site_Code As Byte = 2
Private Const Rep_Code2 As Byte = 3
Private Const StartDate As Byte = 4
Private Const Model2 As Byte = 5
Private Const Call_Status2 As Byte = 6
Private Const Got_Lost As Byte = 7
Private Const GotLost_Date As Byte = 8
Private Const Lost_Cat As Byte = 9
Private Const QuotDocId As Byte = 10
Private Const QuotSrl_No As Byte = 11
Private Const U_Name As Byte = 12
Private Const U_EntDt As Byte = 13
Private Const U_AE As Byte = 14
Private Const Trf_Date As Byte = 15
Private Const Call_Status2Old As Byte = 16

Private Sub Ini_Grid()
Dim I As Byte
    
    With FGrid
        .left = 195
        .Cols = 17
        .top = 5000
'        .BackColor = CellBackColLeave
'        .BackColorBkg = GridBackColorBkg
'        .RowHeightMin = PubGridRowHeight
        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 525
        .ColWidth(PartyCode) = 0
        .ColWidth(Site_Code) = 0
        .ColWidth(Rep_Code2) = 0
        
        .TextMatrix(0, StartDate) = "Start Date"
        .ColAlignmentFixed(StartDate) = flexAlignCenterCenter
        .ColAlignment(StartDate) = flexAlignLeftCenter
        .ColWidth(StartDate) = 1245
       
        .TextMatrix(0, Model2) = "Model"
        .ColAlignmentFixed(Model2) = flexAlignCenterCenter
        .ColAlignment(Model2) = flexAlignLeftCenter
        .ColWidth(Model2) = 1680
        
        .TextMatrix(0, Call_Status2) = "Call Status"
        .ColAlignmentFixed(Call_Status2) = flexAlignCenterCenter
        .ColAlignmentFixed(Call_Status2) = flexAlignLeftCenter
        .ColWidth(Call_Status2) = 1050
        
        .TextMatrix(0, Got_Lost) = "Got/Lost"
        .ColAlignmentFixed(Got_Lost) = flexAlignCenterCenter
        .ColAlignmentFixed(Got_Lost) = flexAlignLeftCenter
        .ColWidth(Got_Lost) = 1000
        
        .TextMatrix(0, GotLost_Date) = "Got/Lost Date"
        .ColAlignmentFixed(GotLost_Date) = flexAlignCenterCenter
        .ColAlignment(GotLost_Date) = flexAlignLeftCenter
        .ColWidth(GotLost_Date) = 1380
        
        .TextMatrix(0, Lost_Cat) = "Lost Cat"
        .ColAlignmentFixed(Lost_Cat) = flexAlignCenterCenter
        .ColAlignment(Lost_Cat) = flexAlignLeftCenter
        .ColWidth(Lost_Cat) = 1860
  
        .TextMatrix(0, QuotDocId) = "QuotDocId"
        .ColAlignmentFixed(QuotDocId) = flexAlignCenterCenter
        .ColAlignment(QuotDocId) = flexAlignRightCenter
        .ColWidth(QuotDocId) = 1650
        
        .TextMatrix(0, QuotSrl_No) = "SNo"
        .ColAlignmentFixed(QuotSrl_No) = flexAlignRightCenter
        .ColWidth(QuotSrl_No) = 555
        
        .ColWidth(U_Name) = 0
        .ColWidth(U_EntDt) = 0
        .ColWidth(U_AE) = 0
        .ColWidth(Trf_Date) = 0
        .ColWidth(Call_Status2Old) = 0
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    DGRep.left = 6660: DGRep.top = mTopScale
    DGParty.left = 6660: DGParty.top = mTopScale
    DGProsCust.left = 6660: DGProsCust.top = mTopScale
    dGObj.left = 6660: dGObj.top = mTopScale
'    DGMod.left = Me.width - (DGMod.width + mRtScale): DGMod.top = mTopScale
    DGMod.left = Me.left: DGMod.width = Me.width - mRtScale
    DGMod.top = txt(Model).top + txt(Model).height: DGMod.height = Me.height - (DGMod.top + mTopScale)
End Sub

Private Sub DGMod_Click()
'If RsMod.RecordCount  > 0 Then
'    TxtGrid(0).Text = RsMod!Code
'    FGrid.TextMatrix(FGrid.Row, Model2) = RsMod!Code
'End If
'TxtGrid(0).SetFocus
'DGMod.Visible = False
    If RsMod.RecordCount > 0 Then
        txt(Model).TEXT = RsMod!Code
        txt(Model).Tag = RsMod!Code
    End If
    txt(Model).SetFocus
    DGMod.Visible = False
End Sub

Private Sub FGrid_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
FGrid_KeyPress (vbKeyReturn)
TAddMode = False
End Sub

Private Sub FGrid_GotFocus()
    If FGrid.BackColorSel = BackColorSelLeave Then FGrid.Col = StartDate
    FGrid.BackColorSel = BackColorSelEnter
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
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
        TopCtrl1_eSave
    End If
    Exit Sub
'        SendKeysA vbKeyTab, True
'        KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case Model2
'            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
End If
KeyCode = 0
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
'Dim gg As String
SetMaxLength
    Select Case FGrid.Col
'        Case StartDate
'           Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
'        Case Model2
'            If FGrid.TextMatrix(FGrid.Row, StartDate) <> "" Then
'               Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
'            End If
        Case Call_Status2
            If FGrid.TextMatrix(FGrid.Row, Model2) <> "" Then
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            End If
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_LostFocus()
    If TxtGrid(0).Visible = False Then FGrid.BackColorSel = BackColorSelLeave
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
'at udaipur
'If KeyCode = vbKeyD And Shift = 2 Then
'    If FGrid.Row  >= 1 Then
'        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
'            If FGrid.Rows  > 2 Then
'                FGrid.RemoveItem (FGrid.Row)
'            Else
'                FGrid.Rows = 1
'                FGrid.AddItem FGrid.Rows
'                FGrid.FixedRows = 1
'            End If
'         End If
'         For i = 1 To FGrid.Rows - 1
'            FGrid.TextMatrix(i, 0) = i
'         Next
'    Else
'        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
'    End If
'    FGrid.SetFocus
'End If
Exit Sub
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
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
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
        Dim SiteCond As String
        SiteCond = " Where  VisitDate Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " And LEFT(site_code,1)='" & PubSiteCode & "'"
    End If
    
    
    If PubMoveRecYn Then
        If PubBackEnd = "A" Then
            Master.Open "select (Visits.site_Code + " & cCStr("Visits.VisitDate") & "  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") as searchcode,Visits.* from Visits  " & SiteCond & " order by Visits.VisitDate", GCn, adOpenDynamic, adLockOptimistic
        ElseIf PubBackEnd = "S" Then
            Master.Open "select (Visits.site_Code + Convert(nVarChar,Visits.VisitDate,3)  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") as searchcode,Visits.* from Visits " & SiteCond & " order by Visits.VisitDate", GCn, adOpenDynamic, adLockOptimistic
        End If
    Else
        If PubBackEnd = "A" Then
            Master.Open "select Top 1 (Visits.site_Code + " & cCStr("Visits.VisitDate") & "  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") as searchcode,Visits.* from Visits  " & SiteCond & " order by Visits.VisitDate", GCn, adOpenDynamic, adLockOptimistic
        ElseIf PubBackEnd = "S" Then
            Master.Open "select Top 1 (Visits.site_Code + Convert(nVarChar,Visits.VisitDate,3)  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") as searchcode,Visits.* from Visits  " & SiteCond & " order by Visits.VisitDate", GCn, adOpenDynamic, adLockOptimistic
        End If
    End If
   
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "select SubCode as code, Name, ' ' as NSuffix  from subgroup where Nature='Customer' order by Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set rsProsCust = New ADODB.Recordset
    rsProsCust.CursorLocation = adUseClient
    rsProsCust.Open "select Cust_code as Code, Name,NSuffix,Profession,CityCode,Area from ProspectiveCust order by Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGProsCust.DataSource = rsProsCust
    
    Set rsObj = New ADODB.Recordset
    rsObj.CursorLocation = adUseClient
    rsObj.Open "select objcode as code,objdesc as name from VisitObjective order by objdesc", GCn, adOpenDynamic, adLockOptimistic
    Set dGObj.DataSource = rsObj
    
    Set RsRep = New ADODB.Recordset
    RsRep.CursorLocation = adUseClient
    RsRep.Open "select Emp_code as code,emp_name as name from emp_mast where emp_type = 0  order by Emp_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGRep.DataSource = RsRep
    
    Set RsMod = New ADODB.Recordset
    RsMod.CursorLocation = adUseClient
    RsMod.Open "select Model as code,Model_Desc as NAME, Chas_Type from model where (div_code='" & PubDivCode & "' or Div_Code='') order by model", GCn, adOpenDynamic, adLockOptimistic
    Set DGMod.DataSource = RsMod
    
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    txt(VisitDate).Tag = PubLoginDate
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
Set RsMod = Nothing
Set RsParty = Nothing
Set RsRep = Nothing
Set rsProsCust = Nothing
Set rsObj = Nothing
Set Master = Nothing
End Sub

Private Sub ListView_Click()
If TxtGrid(0).Visible Then
    TxtGrid(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    TxtGrid(Val(ListView.Tag)).SetFocus
Else
    txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    txt(Val(ListView.Tag)).SetFocus
End If
FrmList.Visible = False
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim VNo As Long
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    txt(Srlno).Enabled = False
    txt(VisitDate) = Format(txt(VisitDate).Tag, "dd/mm/yyyy")
    txt(NewEnquiry) = "No"
    txt(Trf_YN) = "No"
    txt(VisitDate).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim mTrans As Boolean, vBook As Variant
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        mTrans = True
        GCn.BeginTrans

        GCn.Execute ("Delete from visits where Rep_code='" & txt(REP_CODE).Tag & "' and VisitDate = " & ConvertDate(txt(VisitDate)) & " and SrlNo = " & Val(txt(Srlno).TEXT) & "")
    'GCn.Execute ("Delete from Veh_SubGroupQuot where Rep_code='" & txt(REP_CODE).Tag & "' and PartyCode='" & txt(Party_code).Tag & "' ")
        GCn.CommitTrans
        mTrans = False
        Master.Requery
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
        End If
        BUTTONS True, Me, Master, 0
        Call MoveRec
    End If
eloop1:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    FGrid.AddItem FGrid.Rows
    txt(Party_code).SetFocus
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
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
       Dim SiteCond As String
       SiteCond = " Where  VisitDate Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and LEFT(v.site_code,1)='" & PubSiteCode & "'"
    End If
    If PubBackEnd = "A" Then
        GSQL = "select (V.site_Code + " & cCStr("V.VisitDate") & " + V.Rep_Code + " & cCStr("V.SrlNo") & ") as searchcode," & _
            " V.VisitDate as Visit_Date,V.SrlNo,Emp.Emp_Name,V.Objective,V.Next_Date " & _
            " from Visits V left Join Emp_Mast Emp on V.Rep_Code=Emp.Emp_Code " & SiteCond & "" & _
            " Order by V.VisitDate"
    ElseIf PubBackEnd = "S" Then
        GSQL = "select (V.site_Code + Convert(nVarChar,V.VisitDate,3) + V.Rep_Code + " & cCStr("V.SrlNo") & ") as searchcode," & _
            " V.VisitDate as Visit_Date,V.SrlNo,Emp.Emp_Name,V.Objective,V.Next_Date " & _
            " from Visits V left Join Emp_Mast Emp on V.Rep_Code=Emp.Emp_Code " & SiteCond & " " & _
            " Order by V.VisitDate"
    End If
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
        If PubBackEnd = "A" Then
            Set Master = GCn.Execute("select (Visits.site_Code + " & cCStr("Visits.VisitDate") & "  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") as searchcode,Visits.* from Visits Where (Visits.site_Code + " & cCStr("Visits.VisitDate") & "  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") = '" & MyValue & "'  order by Visits.VisitDate")
        ElseIf PubBackEnd = "S" Then
            Set Master = GCn.Execute("select (Visits.site_Code + Convert(nVarChar,Visits.VisitDate,3)  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") as searchcode,Visits.* from Visits Where (Visits.site_Code + Convert(nVarChar,Visits.VisitDate,3)  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") = '" & MyValue & "'  order by Visits.VisitDate")
        End If
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

Private Sub TopCtrl1_eRef()
    RsRep.Requery
    rsObj.Requery
    RsParty.Requery
    rsProsCust.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer, mRepPWD$, mStr$
    Dim mTrans As Boolean
    Dim Viscall As Integer
    Dim StCall As Integer
    Dim mNextDate As Variant
'    On Error GoTo errlbl
    
    Grid_Hide
    If IsValid(txt(VisitDate), "Visit Date") = False Then Exit Sub
    If IsValid(txt(REP_CODE), "Representative name") = False Then Exit Sub
    If IsValid(txt(Srlno), "Serial Number") = False Then Exit Sub
    If IsValid(txt(NewEnquiry), "New Enquiry") = False Then Exit Sub
    If txt(NewEnquiry) = "Yes" Then
        If IsValid(txt(Model), "Model") = False Then
            Exit Sub
        Else
            GSQL = "Select PartyCode from Veh_SubGroupQuot " & _
                "where PartyCode='" & txt(Party_code).Tag & "' and StartDate=" & ConvertDate(txt(VisitDate)) & " and Model='" & txt(Model) & "'"
            If GCn.Execute(GSQL).RecordCount > 0 Then
                MsgBox "Model Already found for this party, Select another Model", vbOKOnly, "Duplicate Model"
                txt(Model).SetFocus
                Exit Sub
            End If
        End If
    End If
    If IsValid(txt(Visit_Call), "Visit / Call") = False Then Exit Sub
    If txt(Meet_TimeFrom) = "00:00" Then MsgBox "Meeting Time From is required", vbOKOnly, "Validation": txt(Meet_TimeFrom).SetFocus: Exit Sub
    If txt(Meet_TimeTo) = "00:00" Then MsgBox "Meeting Time To is required", vbOKOnly, "Validation": txt(Meet_TimeTo).SetFocus: Exit Sub
    If IsValid(txt(Call_Status), "Call Status") = False Then Exit Sub
    
    mRepPWD = GCn.Execute("Select Access_PWD from Emp_Mast where Emp_Code='" & txt(REP_CODE).Tag & "'").Fields(0).Value
    If txt(RepPWD) <> mRepPWD Then
        MsgBox "Please enter valid password!", vbOKOnly, "Authorisation Checking"
        txt(RepPWD).SetFocus
        Exit Sub
    End If
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, StartDate) <> "" Then
            If FGrid.TextMatrix(I, Model2) = "" Then MsgBox "Fill Model in Row No " & I, vbInformation, "Required Data": FGrid.Row = I: FGrid.Col = Model2: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Call_Status2) = "" Then MsgBox "Fill Call Status in Row No " & I, vbInformation, "Required Data": FGrid.Row = I: FGrid.Col = Call_Status2: FGrid.SetFocus: Exit Sub
        End If
    Next
 RemoveTxtNull
 GCn.BeginTrans
    mTrans = True
    Select Case txt(Visit_Call).TEXT
        Case "Visit"
            Viscall = 0
        Case "Call"
            Viscall = 1
    End Select
    StCall = FxCallStatus(txt(Call_Status), False)
    If IsNull(txt(Next_Date)) Or txt(Next_Date) = "" Then
        mNextDate = "Null"
    Else
'        ConvertDate = "#" & Format(CDate(temp), "dd/MMM/yyyy") & "#"
        mNextDate = "" & ConvertDate(Format(txt(Next_Date) & " " & txt(Next_Time), "dd/MMM/yyyy hh:mm")) & ""
    End If

    If TopCtrl1.TopText2.CAPTION = "Add" Then
        GCn.Execute ("delete from visits where  VisitDate = " & ConvertDate(txt(VisitDate)) & " and Rep_Code = '" & txt(REP_CODE).Tag & "' and  SrlNo = " & Val(txt(Srlno)) & " and  Site_Code = '" & PubSiteCode & "'")
        GCn.Execute ("Insert into visits( VisitDate , Rep_Code, SrlNo, Div_Code, Site_Code, ProspectiveCust_SubGroup, Party_Code, Trf_YN, TrfFrom_RepCode, " & _
            " Visit_Call, Meet_TimeFrom, Meet_Timeto, REMARK1, OBJECTIVE, REMARK2, Call_Status, NEXT_DATE, EXPENCE, EXPREMARK,U_Name, U_EntDt, U_AE,Schemes,SalesNos,Prices,Pamphlets,Hoardings,Events,MediaAds,Misc ) " & _
            " values(" & ConvertDate(txt(VisitDate)) & ", '" & txt(REP_CODE).Tag & "', " & Val(txt(Srlno)) & ", '" & PubDivCode & "', '" & PubSiteCode & "', 0,'" & txt(Party_code).Tag & "', " & IIf(txt(Trf_YN) = "Yes", 1, 0) & " , '" & txt(TrfFrom_RepCode).Tag & _
            "'," & Viscall & ",'" & txt(Meet_TimeFrom) & "','" & txt(Meet_TimeTo) & "','" & txt(Remark1) & "','" & txt(Objective).Tag & "','" & txt(Remark2) & "', " & StCall & ", " & mNextDate & ", " & Val(txt(Expence)) & ",'" & txt(ExpRemark) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A' ," & IIf(LstMarket.SELECTED(0) = True, 1, 0) & "," & IIf(LstMarket.SELECTED(1) = True, 1, 0) & "," & IIf(LstMarket.SELECTED(2) = True, 1, 0) & "," & IIf(LstMarket.SELECTED(3) = True, 1, 0) & "," & IIf(LstMarket.SELECTED(4) = True, 1, 0) & "," & IIf(LstMarket.SELECTED(5) = True, 1, 0) & " ," & IIf(LstMarket.SELECTED(6) = True, 1, 0) & " ," & IIf(LstMarket.SELECTED(7) = True, 1, 0) & ")")
    Else
        GCn.Execute ("update visits  set ProspectiveCust_SubGroup = 0, Party_Code='" & txt(Party_code).Tag & "', Trf_YN=" & IIf(txt(Trf_YN) = "Yes", 1, 0) & " , TrfFrom_RepCode = '" & txt(TrfFrom_RepCode).Tag & _
            "',Visit_Call=" & Viscall & ", Meet_TimeFrom='" & txt(Meet_TimeFrom) & "', Meet_Timeto='" & txt(Meet_TimeTo) & "', REMARK1 = '" & txt(Remark1) & "', OBJECTIVE = '" & txt(Objective).Tag & "', REMARK2 = '" & txt(Remark2) & "', Call_Status=" & StCall & _
            " , NEXT_DATE =" & mNextDate & ",EXPENCE = " & Val(txt(Expence)) & ", EXPREMARK =  '" & txt(ExpRemark) & "',U_Name = '" & pubUName & "', U_EntDt = " & ConvertDate(PubServerDate) & ", U_AE ='E',Schemes = " & IIf(LstMarket.SELECTED(0) = True, 1, 0) & ", SalesNos = " & IIf(LstMarket.SELECTED(1) = True, 1, 0) & ",Prices = " & IIf(LstMarket.SELECTED(2) = True, 1, 0) & ",Pamphlets = " & IIf(LstMarket.SELECTED(3) = True, 1, 0) & ", Hoardings = " & IIf(LstMarket.SELECTED(4) = True, 1, 0) & ",Events = " & IIf(LstMarket.SELECTED(5) = True, 1, 0) & " ,MediaAds = " & IIf(LstMarket.SELECTED(6) = True, 1, 0) & " ,Misc = " & IIf(LstMarket.SELECTED(7) = True, 1, 0) & "" & _
            " where VisitDate = " & ConvertDate(txt(VisitDate)) & " and Rep_Code = '" & txt(REP_CODE).Tag & "' and  SrlNo = " & Val(txt(Srlno)) & "")
    End If
'    GCn.Execute ("Delete from Veh_SubGroupQuot where PartyCode='" & Txt(Party_code).Tag & "' and rep_code='" & Txt(REP_CODE).Tag & "'")
'    For i = 1 To FGrid.Rows - 1
'        If FGrid.TextMatrix(i, StartDate) <> "" Then
'            GCn.Execute ("insert into Veh_SubGroupQuot " & _
'                "(StartDate, Model,PartyCode, Site_Code, Rep_Code, Call_Status, " & _
'                "Got_Lost, GotLost_Date, Lost_Cat, QuotDocId, QuotSrl_No, " & _
'                "U_Name , U_EntDt, U_AE, Trf_Date) " & _
'                "values(" & ConvertDate(FGrid.TextMatrix(i, StartDate)) & ",'" & FGrid.TextMatrix(i, Model2) & "','" & Txt(Party_code).Tag & _
'                "','" & PubSiteCode & PubSiteCode & "','" & Txt(REP_CODE).Tag & "'," & FxCallStatus(FGrid.TextMatrix(i, Call_Status2), False) & _
'                " ,'" & FGrid.TextMatrix(i, Got_Lost) & "'," & ConvertDate(FGrid.TextMatrix(i, GotLost_Date)) & ",'" & FGrid.TextMatrix(i, Lost_Cat) & _
'                "','" & FGrid.TextMatrix(i, QuotDocId) & "'," & Val(FGrid.TextMatrix(i, QuotSrl_No)) & _
'                ",'" & pubUName & "',#" & PubLoginDate & "#,'" & left(TopCtrl1.TopText2.Caption, 1) & "'," & ConvertDate(FGrid.TextMatrix(i, GotLost_Date)) & ")")
'        End If
'    Next
    
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Call_Status2) <> FGrid.TextMatrix(I, Call_Status2Old) Then
            GCn.Execute ("Update Veh_SubGroupQuot set " & _
                "Call_Status=" & FxCallStatus(FGrid.TextMatrix(I, Call_Status2), False) & _
                " ,U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'" & _
                " where PartyCode='" & txt(Party_code).Tag & "' and StartDate=" & ConvertDate(FGrid.TextMatrix(I, StartDate)) & " and Model='" & FGrid.TextMatrix(I, Model2) & "' and Rep_Code='" & txt(REP_CODE).Tag & "'")
        End If
    Next
    If txt(NewEnquiry) = "Yes" Then
        'Insert model in Veh_subGroupQuot
        GCn.Execute ("insert into Veh_SubGroupQuot " & _
            "(StartDate, Model,PartyCode, Site_Code, Rep_Code, Call_Status, " & _
            "U_Name , U_EntDt, U_AE) " & _
            "values(" & ConvertDate(txt(VisitDate)) & ",'" & txt(Model) & "','" & txt(Party_code).Tag & _
            "','" & PubSiteCode & PubSiteCode & "','" & txt(REP_CODE).Tag & "'," & FxCallStatus(txt(Call_Status), False) & _
            ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
    End If
GCn.CommitTrans
mTrans = False

    If PubBackEnd = "A" Then
        mStr = PubSiteCode & txt(VisitDate) & txt(REP_CODE).Tag & txt(Srlno)
    ElseIf PubBackEnd = "S" Then
        mStr = PubSiteCode & Format(txt(VisitDate), "DD/MM/YY") & txt(REP_CODE).Tag & txt(Srlno)
    End If


    If PubMoveRecYn Then
        Master.Requery
    Else
        If PubBackEnd = "A" Then
            Set Master = GCn.Execute("select (Visits.site_Code + " & cCStr("Visits.VisitDate") & "  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") as searchcode,Visits.* from Visits Where (Visits.site_Code + " & cCStr("Visits.VisitDate") & "  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") = '" & mStr & "'  order by Visits.VisitDate")
        ElseIf PubBackEnd = "S" Then
            Set Master = GCn.Execute("select (Visits.site_Code + Convert(nVarChar,Visits.VisitDate,3)  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") as searchcode,Visits.* from Visits Where (Visits.site_Code + Convert(nVarChar,Visits.VisitDate,3)  + Visits.Rep_Code + " & cCStr("Visits.SrlNo") & ") = '" & mStr & "'  order by Visits.VisitDate")
        End If
    End If
        Master.FIND "SearchCode = '" & mStr & "'"
            
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        txt(VisitDate).Tag = Format(txt(VisitDate), "dd/mm/yyyy")
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
If TxtGrid(0).Visible = True Then TxtGrid(0).Visible = False
Grid_Hide
Ctrl_GetFocus txt(Index)
Select Case Index
    Case Call_Status
        ListArray = Array("Cold", "Warm", "Hot")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 3)
    Case Visit_Call
        ListArray = Array("Visit", "Call")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
    Case REP_CODE
         DGRep.Tag = 1
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or txt(REP_CODE).TEXT = "" Then Exit Sub
        If txt(REP_CODE).TEXT <> RsRep!Name Then
            RsRep.MoveFirst
            RsRep.FIND "name ='" & txt(REP_CODE).TEXT & "'"
        End If
    Case TrfFrom_RepCode
        DGRep.Tag = 2
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or txt(TrfFrom_RepCode).TEXT = "" Then Exit Sub
        If txt(TrfFrom_RepCode).TEXT <> RsRep!Name Then
            RsRep.MoveFirst
            RsRep.FIND "name ='" & txt(TrfFrom_RepCode).TEXT & "'"
        End If
    Case Party_code
'        If txt(ProsCust_SubGroup).Text = "Yes" Then
'            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Party_code).Text = "" Then Exit Sub
'            If txt(Party_code).Text <> RsParty!Name Then
'                RsParty.MoveFirst
'                RsParty.FIND "name ='" & txt(Party_code).Text & "'"
'            End If
'        Else
            If rsProsCust.RecordCount = 0 Or rsProsCust.EOF = True Or rsProsCust.BOF = True Or txt(Party_code).TEXT = "" Then Exit Sub
            If txt(Party_code).TEXT <> rsProsCust!Name Then
                rsProsCust.MoveFirst
                rsProsCust.FIND "name ='" & txt(Party_code).TEXT & "'"
            End If
'        End If
    Case Objective
        If rsObj.RecordCount = 0 Or (rsObj.EOF = True Or rsObj.BOF = True) Or txt(Objective).TEXT = "" Then Exit Sub
        If txt(Objective).TEXT <> rsObj!Name Then
            rsObj.MoveFirst
            rsObj.FIND "name ='" & txt(Objective).TEXT & "'"
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
    txt(Index).TEXT = ""
    Grid_Hide
    Exit Sub
End If

Select Case Index
    Case Model
        DGridTxtKeyDown DGMod, txt, Index, RsMod, KeyCode, False, 0, frmModel, "frmModel"
    Case Visit_Call
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case Call_Status
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 900
    Case REP_CODE
        DGridTxtKeyDown DGRep, txt, Index, RsRep, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
    Case TrfFrom_RepCode
        DGridTxtKeyDown DGRep, txt, Index, RsRep, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
    Case Party_code
'        If txt(ProsCust_SubGroup).Text = "Yes" Then
'            DGridTxtKeyDown DgParty, txt, Index, RsParty, KeyCode, False, 1, frmSubGroup
'        Else
            DGridTxtKeyDown DGProsCust, txt, Index, rsProsCust, KeyCode, False, 1, frmProCust, "frmProCust"
'        End If
    Case Objective
        DGridTxtKeyDown dGObj, txt, Index, rsObj, KeyCode, False, 1
End Select
If DGMod.Visible = False And FrmList.Visible = False And DGRep.Visible = False And DGProsCust.Visible = False And dGObj.Visible = False And DGParty.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
'        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = ExpRemark Then
'            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
'        End If
    If TopCtrl1.TopText2 = "Add" And Index <> VisitDate Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    If TopCtrl1.TopText2 = "Edit" And Index <> Party_code Then If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub
Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case Index
    Case Model
        If DGMod.Visible = True Then DGridTxtKeyPress txt, Index, RsMod, KeyAscii, "code"
    Case Meet_TimeFrom, Meet_TimeTo, Next_Time
        Call NumPress(txt(Index), KeyAscii, 2, 2)
    Case Party_code
'        If txt(ProsCust_SubGroup).Text = "Yes" Then
'            If DgParty.Visible = True Then DGridTxtKeyPress txt, Index, RsParty, KeyAscii, "Name"
'        Else
            If DGProsCust.Visible = True Then DGridTxtKeyPress txt, Index, rsProsCust, KeyAscii, "Name"
'        End If
    Case REP_CODE
        If DGRep.Visible = True Then DGridTxtKeyPress txt, Index, RsRep, KeyAscii, "Name"
    Case TrfFrom_RepCode
        If DGRep.Visible = True Then DGridTxtKeyPress txt, Index, RsRep, KeyAscii, "Name"
    Case Objective
        If dGObj.Visible = True Then DGridTxtKeyPress txt, Index, rsObj, KeyAscii, "Name"
    Case NewEnquiry
        If UCase(Chr(KeyAscii)) = "Y" Then
            txt(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            txt(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txt(Index) = ""
        End If
        KeyAscii = 0
        
    Case Trf_YN
        If UCase(Chr(KeyAscii)) = "Y" Then
            txt(Index) = "Yes"
            txt(TrfFrom_RepCode).Enabled = True
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            txt(Index) = "No"
            txt(TrfFrom_RepCode).Enabled = False
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            txt(Index) = ""
            txt(TrfFrom_RepCode).Enabled = False
        End If
        KeyAscii = 0
    Case Expence
        Call NumPress(txt(Index), KeyAscii, 5, 2)
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case Visit_Call
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case Call_Status
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case Model
        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsMod!Code
            txt(Index).Tag = RsMod!Code
        End If
    Case Meet_TimeFrom, Meet_TimeTo, Next_Time
        txt(Index) = Format(txt(Index), "hh:mm")
    Case Expence
        txt(Expence) = IIf(Val(txt(Expence)) = 0, "", Format(txt(Expence), "0.00"))
    Case Objective
        If rsObj.RecordCount = 0 Or (rsObj.EOF = True Or rsObj.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = rsObj!Name
            txt(Index).Tag = rsObj!Code
        End If
    Case Party_code
'        If txt(ProsCust_SubGroup).Text = "Yes" Then
'            If RsParty.RecordCount = 0 Or RsParty.EOF = True Or RsParty.BOF = True Or txt(Index).Text = "" Then
'                txt(Index).Text = ""
'                txt(Index).Tag = ""
'            Else
'                txt(Index).Text = RsParty!Name
'                txt(Index).Tag = RsParty!Code
'                FillModelGrid
'            End If
'        Else
            If rsProsCust.RecordCount = 0 Or rsProsCust.EOF = True Or rsProsCust.BOF = True Or txt(Index).TEXT = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
                txt(Area) = ""
                txt(Profession) = ""
                txt(City) = ""
            Else
                txt(Index).TEXT = rsProsCust!Name
                txt(Index).Tag = rsProsCust!Code
                FillPartyDetail
                FillModelGrid
            End If
'        End If
    Case Model
        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or txt(Model).TEXT = "" Then Exit Sub
        If txt(Model).TEXT <> RsMod!Code Then
            RsMod.MoveFirst
            RsMod.FIND "code ='" & txt(Model).TEXT & "'"
        End If
        
    Case REP_CODE
        If IsValid(txt(REP_CODE), "Representative name") = False Then Exit Sub
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsRep!Name
            txt(Index).Tag = RsRep!Code
            FillModelGrid
        End If
        If txt(VisitDate).TEXT = "" Then
            txt(VisitDate).SetFocus
        Else
            txt(Srlno).TEXT = IIf(GCn.Execute("select count(*) from visits where VisitDate = " & ConvertDate(txt(VisitDate).TEXT) & " and Rep_Code  = '" & txt(REP_CODE).Tag & "'").Fields(0).Value > 0, GCn.Execute("select MAX(srlno) from visits  where VisitDate = " & ConvertDate(txt(VisitDate).TEXT) & " and Rep_Code  = '" & txt(REP_CODE).Tag & "'").Fields(0).Value + 1, 1)
        End If
    Case TrfFrom_RepCode
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsRep!Name
            txt(Index).Tag = RsRep!Code
        End If
    
    Case Call_Status, Visit_Call
        If txt(Index).TEXT <> "" Then txt(Index).TEXT = ListView.SelectedItem.TEXT
    Case VisitDate
        txt(Index).TEXT = RetDate(txt(Index))
        If txt(REP_CODE).TEXT = "" Then
            txt(REP_CODE).SetFocus
        Else
            txt(Srlno).TEXT = IIf(GCn.Execute("select count(*) from visits where VisitDate = " & ConvertDate(txt(VisitDate).TEXT) & " and Rep_Code  = '" & txt(REP_CODE).Tag & "'").Fields(0).Value > 0, GCn.Execute("select MAX(srlno) from visits  where VisitDate = " & ConvertDate(txt(VisitDate).TEXT) & " and Rep_Code  = '" & txt(REP_CODE).Tag & "'").Fields(0).Value + 1, 1)
        End If
    Case Next_Date
        txt(Index).TEXT = RetDate(txt(Index))
End Select
End Sub

Private Sub DGRep_Click()
    If DGRep.Tag = 1 Then
        If RsRep.RecordCount > 0 Then
            txt(REP_CODE).TEXT = RsRep!Name
            txt(REP_CODE).Tag = RsRep!Code
        End If
        txt(REP_CODE).SetFocus
    ElseIf DGRep.Tag = 2 Then
        If RsRep.RecordCount > 0 Then
            txt(TrfFrom_RepCode).TEXT = RsRep!Name
            txt(TrfFrom_RepCode).Tag = RsRep!Code
        End If
        txt(TrfFrom_RepCode).SetFocus
    End If
    DGRep.Visible = False
End Sub

Private Sub DGParty_Click()
'    If txt(ProsCust_SubGroup).Text = "Yes" Then
'        If RsParty.RecordCount  > 0 Then
'            txt(Party_code).Text = RsParty!Name
'            txt(Party_code).Tag = RsParty!Code
'        End If
'    Else
        If rsProsCust.RecordCount > 0 Then
            txt(Party_code).TEXT = rsProsCust!Name
            txt(Party_code).Tag = rsProsCust!Code
        End If
'    End If
    txt(Party_code).SetFocus
    DGParty.Visible = False
End Sub
Private Sub dgobj_Click()
    If rsObj.RecordCount > 0 Then
        txt(Objective).TEXT = rsObj!Name
        txt(Objective).Tag = rsObj!Code
    End If
    txt(Objective).SetFocus
    dGObj.Visible = False
End Sub
Private Sub DGProsCust_Click()
    If rsProsCust.RecordCount > 0 Then
        txt(Party_code).TEXT = rsProsCust!Name
        txt(Party_code).Tag = rsProsCust!Code
    End If
    txt(Party_code).SetFocus
    DGProsCust.Visible = False
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    If I <> VisitDate Then txt(I).TEXT = "": txt(I).Tag = ""
Next I
For I = 0 To 7
    LstMarket.SELECTED(I) = False
Next
FGrid.Rows = 1
FGrid.AddItem FGrid.Rows:  FGrid.FixedRows = 1

End Sub

Private Sub MoveRec()
Dim Rs As Recordset, I As Integer
On Error GoTo error1
TopCtrl1.tPrn = False
If Master.RecordCount > 0 Then
    txt(VisitDate) = Master!VisitDate
    txt(Srlno) = Master!Srlno
    txt(Meet_TimeFrom) = Format(XNull(Master!Meet_TimeFrom), "hh:mm")
    txt(Meet_TimeTo) = Format(XNull(Master!Meet_TimeTo), "hh:mm")
    txt(Remark1) = IIf(IsNull(Master!Remark1), "", Master!Remark1)
    txt(Remark2) = IIf(IsNull(Master!Remark2), "", Master!Remark2)
    txt(Next_Date) = IIf(IsNull(Master!Next_Date), "", Master!Next_Date)
    txt(Next_Time) = IIf(IsNull(Master!Next_Date), "", Format(Master!Next_Date, "hh:mm"))
    txt(Expence) = Format(IIf(IsNull(Master!Expence), "", Master!Expence), "0.00")
    txt(ExpRemark) = IIf(IsNull(Master!ExpRemark), "", Master!ExpRemark)
    txt(NewEnquiry) = IIf(Master!NewEnquiry = 1, "Yes", "No")
    txt(Party_code).Tag = IIf(IsNull(Master!Party_code), "", Master!Party_code)
    If Master!Schemes = 1 Then LstMarket.SELECTED(0) = True Else LstMarket.SELECTED(0) = False
    If Master!SalesNos = 1 Then LstMarket.SELECTED(1) = True Else LstMarket.SELECTED(1) = False
    If Master!Prices = 1 Then LstMarket.SELECTED(2) = True Else LstMarket.SELECTED(2) = False
    If Master!Pamphlets = 1 Then LstMarket.SELECTED(3) = True Else LstMarket.SELECTED(3) = False
    If Master!Hoardings = 1 Then LstMarket.SELECTED(4) = True Else LstMarket.SELECTED(4) = False
    If Master!Events = 1 Then LstMarket.SELECTED(5) = True Else LstMarket.SELECTED(5) = False
    If Master!MediaAds = 1 Then LstMarket.SELECTED(6) = True Else LstMarket.SELECTED(6) = False
    If Master!Misc = 1 Then LstMarket.SELECTED(7) = True Else LstMarket.SELECTED(7) = False
    
    
'    If txt(ProsCust_SubGroup).Text = "Yes" Then
'        If txt(Party_code).Tag <> "" And GCn.Execute("select name from subgroup where SubCode = '" & txt(Party_code).Tag & "'").RecordCount  > 0 Then
'            txt(Party_code).Text = GCn.Execute("select name from subgroup where SubCode = '" & txt(Party_code).Tag & "'").Fields(0).Value
'        Else
'            txt(Party_code).Text = ""
'        End If
'    Else
        If txt(Party_code).Tag <> "" And GCn.Execute("select name from ProspectiveCust where CUST_CODE = '" & txt(Party_code).Tag & "'").RecordCount > 0 Then
            txt(Party_code).TEXT = GCn.Execute("select name from ProspectiveCust where CUST_CODE = '" & txt(Party_code).Tag & "'").Fields(0).Value
        Else
            txt(Party_code).TEXT = ""
        End If
'    End If
    txt(Trf_YN) = IIf(Master!Trf_YN = 1, "Yes", "No")
    If Not IsNull(Master!Visit_Call) Then
        Select Case Master!Visit_Call
            Case 0
                txt(Visit_Call).TEXT = "Visit"
            Case 1
                txt(Visit_Call).TEXT = "Call"
        End Select
    End If
    txt(Call_Status) = FxCallStatus(Master!Call_Status, True)
    txt(Objective).Tag = IIf(IsNull(Master!Objective), "", Master!Objective)
    If txt(Objective).Tag <> "" And GCn.Execute("select objdesc from VisitObjective where objcode= '" & txt(Objective).Tag & "'").RecordCount > 0 Then
        txt(Objective).TEXT = GCn.Execute("select objdesc from VisitObjective where objcode= '" & txt(Objective).Tag & "'").Fields(0).Value
    Else
        txt(Objective).TEXT = ""
    End If
    txt(REP_CODE).Tag = Master!REP_CODE
    If txt(REP_CODE).Tag <> "" And GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & txt(REP_CODE).Tag & "'").RecordCount > 0 Then
        txt(REP_CODE).TEXT = GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & txt(REP_CODE).Tag & "'").Fields(0).Value
    Else
        txt(REP_CODE).TEXT = ""
    End If
    txt(TrfFrom_RepCode).Tag = IIf(IsNull(Master!TrfFrom_RepCode), "", Master!TrfFrom_RepCode)
    If txt(TrfFrom_RepCode).Tag <> "" And GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & txt(TrfFrom_RepCode).Tag & "'").RecordCount > 0 Then
        txt(TrfFrom_RepCode).TEXT = GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & txt(TrfFrom_RepCode).Tag & "'").Fields(0).Value
    Else
        txt(TrfFrom_RepCode).TEXT = ""
    End If
    'Fill Party Details
    rsProsCust.MoveFirst
    rsProsCust.FIND "Code = '" & Master!Party_code & "'"
    FillPartyDetail
    'Fill VehSubGroupQuot Records
    FillModelGrid
Else
    Call BlankText
End If
If FGrid.Rows = 1 Then FGrid.AddItem FGrid.Rows:  FGrid.FixedRows = 1
FGrid.Redraw = True
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
'    txt(i).ForeColor = CtrlFColOrg
Next
LstMarket.Enabled = Enb
If TopCtrl1.TopText2 = "Edit" Then
    txt(VisitDate).Enabled = False
    txt(REP_CODE).Enabled = False
    txt(Srlno).Enabled = False
End If
'txtDisabled_Color Me
End Sub
Private Sub Grid_Hide()
    If DGMod.Visible = True Then DGMod.Visible = False
    If DGRep.Visible = True Then DGRep.Visible = False
    If DGProsCust.Visible = True Then DGProsCust.Visible = False
    If dGObj.Visible = True Then dGObj.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Sub RemoveTxtNull()
Dim I As Integer
For I = 0 To txt.Count - 1
    txt(I).TEXT = IIf(IsNull(txt(I).TEXT), "", txt(I).TEXT)
Next I
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
'GridHelp = False
'If DGMod.Visible = True Then Exit Sub
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    
    Select Case FGrid.Col
         Case Model2
            If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Model2) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Model2) <> RsMod!Code Then
                RsMod.MoveFirst
                RsMod.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, Model2) & "'"
            End If
         Case Call_Status2
            ListArray = Array("Cold", "Warm", "Hot", "Nill")
            Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 4)
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
            Case StartDate
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, QuotSrl_No
                    End If
                End If
            Case Model2    '1
                DGridTxtKeyDown DGMod, TxtGrid, Index, RsMod, KeyCode, True, 0, frmModel, "frmModel"
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                         GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, QuotSrl_No
                    End If
                End If
            Case Call_Status2
                ListView_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, FGrid.top - 1200, 1250, 1200
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, QuotSrl_No
                    End If
                End If
        End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
Select Case FGrid.Col
'Filling on MoveRec only     Case Got_Lost, GotLost_Date, Lost_Cat
    Case Model2
        If DGMod.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsMod, KeyAscii, "code"
End Select
End Sub


Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case Model2
        If KeyCode <> 13 And DGMod.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsMod, KeyCode, "code", True
    Case Call_Status2
        If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
        ListView_KeyUp ListView, TxtGrid, Index, KeyCode, mListItem
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
Dim GridCol As Byte
GridCol = FGrid.Col
Select Case GridCol
    Case StartDate
        TxtGrid(0).TEXT = RetDate(TxtGrid(0))
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
        If FGrid.TextMatrix(FGrid.Rows - 1, StartDate) <> "" Then FGrid.AddItem FGrid.Rows
    Case Model2
        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or TxtGrid(0).TEXT = "" Then
            FGrid.TextMatrix(FGrid.Row, Model2) = ""
        Else
            FGrid.TextMatrix(FGrid.Row, Model2) = RsMod!Code
        End If
    Case Call_Status2
        TxtGrid(0).TEXT = ListView.SelectedItem.TEXT
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    TxtGrid(0).Visible = False
    FGrid.SetFocus
End If
End Function

Private Function FxCallStatus(CallStatus As Variant, VarTypeNumber As Boolean) As Variant
If VarTypeNumber Then
    If Not IsNull(CallStatus) Then
        Select Case CallStatus
            Case 0
                FxCallStatus = "Cold"
            Case 1
                FxCallStatus = "Warm"
            Case 2
                FxCallStatus = "Hot"
            Case 3
                FxCallStatus = "Nill"
        End Select
    End If
Else
    If Not IsNull(CallStatus) Then
        Select Case CallStatus
            Case "Cold"
                FxCallStatus = 0
            Case "Warm"
                FxCallStatus = 1
            Case "Hot"
                FxCallStatus = 2
            Case "Nill"
                FxCallStatus = 3
        End Select
    End If
End If
End Function

Private Sub SetMaxLength()
    Select Case FGrid.Col
        Case StartDate
            TxtGrid(0).MaxLength = 12
            TxtGrid(0).Alignment = 0   '0-Left Align
        Case Else
            TxtGrid(0).MaxLength = 0
    End Select
End Sub
Private Sub FillModelGrid()
'Fill Fgrid
Dim Rs As ADODB.Recordset, I As Integer
If txt(Party_code).Tag = "" Or txt(REP_CODE).Tag = "" Then Exit Sub
    Set Rs = New Recordset
    GSQL = "Select * from Veh_SubGroupQuot where PartyCode='" & txt(Party_code).Tag & "' and Rep_Code='" & txt(REP_CODE).Tag & "'"
    Set Rs = GCn.Execute(GSQL)
    FGrid.Rows = 1: FGrid.Redraw = False
    I = 1
    If Rs.RecordCount > 0 Then
        Do Until Rs.EOF
            With FGrid
                .AddItem I
                .TextMatrix(I, PartyCode) = Rs!PartyCode
                .TextMatrix(I, Site_Code) = Rs!Site_Code
                .TextMatrix(I, Rep_Code2) = Rs!REP_CODE
                .TextMatrix(I, StartDate) = Rs!StartDate
                .TextMatrix(I, Model2) = Rs!Model
                .TextMatrix(I, Call_Status2) = FxCallStatus(Rs!Call_Status, True)
                .TextMatrix(I, Call_Status2Old) = FxCallStatus(Rs!Call_Status, True)
                .TextMatrix(I, Got_Lost) = Rs!Got_Lost
                .TextMatrix(I, GotLost_Date) = Format(Rs!GotLost_Date, "dd/mmm/yyyy")
                If IsNull(Rs!Lost_Cat) Or Rs!Lost_Cat = "" Then
                Else
                    .TextMatrix(I, Lost_Cat) = GCn.Execute("Select NAME from Veh_OrdLostCatg where code='" & Rs!Lost_Cat & "'").Fields(0).Value
                End If
                .TextMatrix(I, QuotDocId) = IIf(IsNull(Rs!QuotDocId), "", Rs!QuotDocId)
                .TextMatrix(I, QuotSrl_No) = IIf(IsNull(Rs!QuotSrl_No), "", Rs!QuotSrl_No)
                .TextMatrix(I, U_Name) = IIf(IsNull(Rs!U_Name), "", Rs!U_Name)
                .TextMatrix(I, U_EntDt) = IIf(IsNull(Rs!U_EntDt), "", Rs!U_EntDt)
                .TextMatrix(I, U_AE) = IIf(IsNull(Rs!U_AE), "", Rs!U_AE)
                .TextMatrix(I, Trf_Date) = IIf(IsNull(Rs!Trf_Date), "", Rs!Trf_Date)
            End With
            Rs.MoveNext
           I = I + 1
        Loop
        FGrid.FixedRows = 1
    End If
Set Rs = Nothing
If FGrid.Rows = 1 Then FGrid.AddItem FGrid.Rows:  FGrid.FixedRows = 1
FGrid.Redraw = True

End Sub

Private Sub FillPartyDetail()

    txt(Profession).Tag = IIf(IsNull(rsProsCust!Profession), "", rsProsCust!Profession)
    If txt(Profession).Tag <> "" And GCn.Execute("select Professionname from Profession where Professioncode = '" & txt(Profession).Tag & "'").RecordCount > 0 Then
        txt(Profession).TEXT = GCn.Execute("select Professionname from Profession where Professioncode = '" & txt(Profession).Tag & "'").Fields(0).Value
    Else
        txt(Profession) = ""
        txt(Profession).Tag = ""
    End If
    txt(City).Tag = IIf(IsNull(rsProsCust!CityCode), "", rsProsCust!CityCode)
    If txt(City).Tag <> "" And GCn.Execute("select cityname from city where citycode = '" & txt(City).Tag & "'").RecordCount > 0 Then
        txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & txt(City).Tag & "'").Fields(0).Value
    Else
        txt(City) = ""
        txt(City).Tag = ""
    End If
    txt(Area).Tag = IIf(IsNull(rsProsCust!Area), "", rsProsCust!Area)
    If txt(Area).Tag <> "" And GCn.Execute("select AREAname from AREA where AREAcode = '" & txt(Area).Tag & "'").RecordCount > 0 Then
        txt(Area).TEXT = GCn.Execute("select AREAname from AREA where AREAcode = '" & txt(Area).Tag & "'").Fields(0).Value
    Else
        txt(Area) = ""
        txt(Area).Tag = ""
    End If

End Sub
