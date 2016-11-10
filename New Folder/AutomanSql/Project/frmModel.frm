VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmModel 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Vehicle Model Master"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11820
   Begin VB.CommandButton CmdUpdatePriceList 
      Caption         =   "Update Price List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8145
      TabIndex        =   113
      Top             =   15
      Width           =   1665
   End
   Begin MSDataGridLib.DataGrid DGCol 
      Height          =   2445
      Left            =   6765
      Negotiate       =   -1  'True
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   8490
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   50
      Left            =   8220
      MaxLength       =   25
      TabIndex        =   50
      ToolTipText     =   "Registered Laden Weight"
      Top             =   3930
      Width           =   3465
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   49
      Left            =   8220
      MaxLength       =   10
      TabIndex        =   49
      ToolTipText     =   "Registered Laden Weight"
      Top             =   3690
      Width           =   3465
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   48
      Left            =   8220
      MaxLength       =   1
      TabIndex        =   48
      ToolTipText     =   "Registered Laden Weight"
      Top             =   3450
      Width           =   3465
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   47
      Left            =   8220
      MaxLength       =   30
      TabIndex        =   47
      ToolTipText     =   "Registered Laden Weight"
      Top             =   3210
      Width           =   3465
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      Index           =   46
      Left            =   8220
      MaxLength       =   10
      TabIndex        =   46
      ToolTipText     =   "Registered Laden Weight"
      Top             =   2970
      Width           =   3465
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   45
      Left            =   8220
      MaxLength       =   6
      TabIndex        =   45
      ToolTipText     =   "Registered Laden Weight"
      Top             =   2730
      Width           =   3465
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   44
      Left            =   8220
      TabIndex        =   44
      ToolTipText     =   "Registered Laden Weight"
      Top             =   2490
      Width           =   3465
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   43
      Left            =   10695
      MaxLength       =   15
      TabIndex        =   43
      ToolTipText     =   "Registered Laden Weight"
      Top             =   2250
      Width           =   990
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   42
      Left            =   8220
      MaxLength       =   15
      TabIndex        =   42
      ToolTipText     =   "Registered Laden Weight"
      Top             =   2250
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   41
      Left            =   6060
      MaxLength       =   8
      TabIndex        =   17
      Text            =   "Yes"
      Top             =   3180
      Width           =   540
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDF4B5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   8445
      TabIndex        =   99
      Top             =   7770
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DGItem 
      Height          =   4395
      Left            =   5685
      Negotiate       =   -1  'True
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   8535
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7752
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777152
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Item Description"
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
            ColumnWidth     =   3330.142
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   5
      Left            =   6300
      MaxLength       =   50
      TabIndex        =   95
      Top             =   2460
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame FrModel 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   135
      TabIndex        =   56
      Top             =   8520
      Visible         =   0   'False
      Width           =   2820
      Begin MSDataGridLib.DataGrid DGModel 
         Height          =   3225
         Left            =   30
         TabIndex        =   57
         Top             =   345
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   5689
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   -2147483648
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   13504523
         HeadLines       =   1
         RowHeight       =   15
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "MODEL"
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
            DataField       =   "MODEL"
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
               DividerStyle    =   0
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "List of Models"
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
         Height          =   270
         Index           =   0
         Left            =   30
         TabIndex        =   58
         Top             =   30
         Width           =   2760
      End
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   20
      Left            =   8220
      MaxLength       =   15
      TabIndex        =   30
      ToolTipText     =   "Registered Laden Weight"
      Top             =   810
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   19
      Left            =   1785
      MaxLength       =   30
      TabIndex        =   27
      Top             =   4380
      Width           =   4815
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   4755
      TabIndex        =   91
      Top             =   7695
      Visible         =   0   'False
      Width           =   1335
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   0
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
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
      Index           =   18
      Left            =   5280
      TabIndex        =   5
      Top             =   780
      Width           =   1320
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   17
      Left            =   10695
      TabIndex        =   40
      Text            =   "999999"
      Top             =   1770
      Width           =   990
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   27
      Left            =   10695
      TabIndex        =   38
      Top             =   1290
      Width           =   990
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   40
      Left            =   10695
      MaxLength       =   10
      TabIndex        =   39
      Top             =   1530
      Width           =   990
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   39
      Left            =   10695
      TabIndex        =   36
      Top             =   810
      Width           =   990
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   38
      Left            =   10695
      MaxLength       =   10
      TabIndex        =   37
      Top             =   1050
      Width           =   990
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   37
      Left            =   8220
      MaxLength       =   15
      TabIndex        =   35
      Top             =   2010
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   36
      Left            =   8220
      MaxLength       =   15
      TabIndex        =   34
      Top             =   1770
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   35
      Left            =   8220
      MaxLength       =   15
      TabIndex        =   33
      Top             =   1530
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   34
      Left            =   8220
      MaxLength       =   15
      TabIndex        =   32
      Top             =   1290
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   16
      Left            =   8220
      MaxLength       =   15
      TabIndex        =   31
      ToolTipText     =   "Registered Laden Weight"
      Top             =   1050
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   7
      Left            =   1785
      MaxLength       =   9
      TabIndex        =   2
      Top             =   780
      Width           =   1320
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   9
      Left            =   5280
      TabIndex        =   4
      Top             =   1020
      Width           =   1320
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   33
      Left            =   8220
      TabIndex        =   28
      Top             =   570
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   32
      Left            =   10695
      TabIndex        =   29
      Top             =   570
      Width           =   990
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   31
      Left            =   4845
      MaxLength       =   15
      TabIndex        =   21
      Text            =   "012345678901234"
      Top             =   3660
      Width           =   1755
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   30
      Left            =   4845
      MaxLength       =   15
      TabIndex        =   23
      Top             =   3900
      Width           =   1755
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   29
      Left            =   3210
      MaxLength       =   15
      TabIndex        =   25
      Top             =   4140
      Width           =   1380
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   28
      Left            =   1785
      MaxLength       =   35
      TabIndex        =   18
      Top             =   2700
      Width           =   4815
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   26
      Left            =   5730
      TabIndex        =   26
      Top             =   4140
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   25
      Left            =   1785
      TabIndex        =   20
      Top             =   3660
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   24
      Left            =   1785
      TabIndex        =   22
      Top             =   3900
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   23
      Left            =   1785
      TabIndex        =   24
      Top             =   4140
      Width           =   870
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   22
      Left            =   1785
      MaxLength       =   50
      TabIndex        =   19
      Top             =   3420
      Width           =   4815
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   21
      Left            =   10695
      TabIndex        =   41
      Top             =   2010
      Width           =   990
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   15
      Left            =   1785
      MaxLength       =   10
      TabIndex        =   16
      Top             =   3180
      Width           =   1845
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   14
      Left            =   6060
      MaxLength       =   8
      TabIndex        =   15
      Text            =   "Yes"
      Top             =   2940
      Width           =   540
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   13
      Left            =   1785
      MaxLength       =   20
      TabIndex        =   11
      Top             =   2220
      Width           =   4485
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   12
      Left            =   6300
      MaxLength       =   50
      TabIndex        =   12
      Top             =   2220
      Visible         =   0   'False
      Width           =   300
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   6
      Left            =   1785
      MaxLength       =   30
      TabIndex        =   13
      Top             =   2460
      Width           =   4485
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   4
      Left            =   1785
      MaxLength       =   20
      TabIndex        =   14
      Top             =   2940
      Width           =   1845
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   11
      Left            =   1785
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1740
      Width           =   4815
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   10
      Left            =   1785
      MaxLength       =   80
      TabIndex        =   7
      Top             =   1500
      Width           =   4815
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   8
      Left            =   1785
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1020
      Width           =   1320
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   6300
      MaxLength       =   50
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   300
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   1785
      MaxLength       =   20
      TabIndex        =   1
      Top             =   540
      Width           =   4815
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   1785
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1980
      Width           =   4485
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   1785
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1260
      Width           =   4815
   End
   Begin VB.Frame FrModelGrp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   840
      TabIndex        =   52
      Top             =   8475
      Visible         =   0   'False
      Width           =   4155
      Begin MSDataGridLib.DataGrid DGModelGrp 
         Height          =   3225
         Left            =   30
         TabIndex        =   59
         Top             =   345
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5689
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   -2147483648
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   13504523
         HeadLines       =   1
         RowHeight       =   15
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "ModelGrp_Code"
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
            DataField       =   "ModelGrp_Name"
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
            DataField       =   "ModelCat_Code"
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
         BeginProperty Column03 
            DataField       =   "ModelCat_Name"
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
         BeginProperty Column04 
            DataField       =   "ModelDiv_Code"
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
         BeginProperty Column05 
            DataField       =   "Div_Name"
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
         BeginProperty Column06 
            DataField       =   "Wheel_Catg"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   3435.024
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   0
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "List of Model Group"
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
         Height          =   270
         Index           =   1
         Left            =   30
         TabIndex        =   53
         Top             =   30
         Width           =   4095
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1665
      Left            =   165
      TabIndex        =   96
      Top             =   4950
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   2937
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   0
      Cols            =   5
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   192
      GridColorUnpopulated=   16761024
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
      FormatString    =   "SrNo."
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
      _Band(0).Cols   =   5
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4395
      Left            =   8745
      Negotiate       =   -1  'True
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   8220
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7752
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777152
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Item Description"
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
            ColumnWidth     =   3330.142
         EndProperty
      EndProperty
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warranty(Months)"
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
      Left            =   9150
      TabIndex        =   112
      Top             =   2025
      Width           =   1545
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body Type............."
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
      Index           =   30
      Left            =   6675
      TabIndex        =   109
      Top             =   3915
      Width           =   1695
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cubic Capacity....."
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
      Index           =   29
      Left            =   6675
      TabIndex        =   108
      Top             =   3675
      Width           =   1590
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FMSN....................."
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
      Index           =   28
      Left            =   6675
      TabIndex        =   107
      Top             =   3435
      Width           =   1725
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rear Axle Make....."
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
      Index           =   27
      Left            =   6675
      TabIndex        =   106
      Top             =   3195
      Width           =   1650
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tank Capacity........."
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
      Index           =   26
      Left            =   6675
      TabIndex        =   105
      Top             =   2955
      Width           =   1770
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Drive.........."
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
      Index           =   25
      Left            =   6675
      TabIndex        =   104
      Top             =   2730
      Width           =   1740
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colour Name........."
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
      Index           =   24
      Left            =   6675
      TabIndex        =   103
      Top             =   2475
      Width           =   1665
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Steering Type..........."
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
      Index           =   23
      Left            =   9150
      TabIndex        =   102
      Top             =   2250
      Width           =   1860
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Regulatory Cert...."
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
      Left            =   6660
      TabIndex        =   101
      Top             =   2250
      Width           =   1605
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Tax Y/N.............."
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
      Left            =   3915
      TabIndex        =   100
      Top             =   3180
      Width           =   2220
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check List Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Index           =   18
      Left            =   150
      TabIndex        =   98
      Top             =   4665
      Width           =   1440
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gear Box No. ........."
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
      Index           =   21
      Left            =   6660
      TabIndex        =   94
      Top             =   795
      Width           =   1755
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tyre Details ............."
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
      Left            =   180
      TabIndex        =   93
      Top             =   4380
      Width           =   1875
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Type*.........."
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
      Left            =   3480
      TabIndex        =   90
      Top             =   780
      Width           =   1800
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cylinder................."
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
      Left            =   9150
      TabIndex        =   89
      Top             =   1290
      Width           =   1740
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel....................."
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
      Left            =   9150
      TabIndex        =   88
      Top             =   1530
      Width           =   1605
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wheel Base(mm)"
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
      Index           =   14
      Left            =   9150
      TabIndex        =   87
      Top             =   810
      Width           =   1485
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horse Power..............."
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
      Left            =   9150
      TabIndex        =   86
      Top             =   1050
      Width           =   1980
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rear Axle Wt.*........"
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
      Left            =   6660
      TabIndex        =   85
      Top             =   1995
      Width           =   1770
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Front Axle Wt.*......"
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
      Left            =   6660
      TabIndex        =   84
      Top             =   1755
      Width           =   1680
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Wt.* ............"
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
      Left            =   6660
      TabIndex        =   83
      Top             =   1515
      Width           =   1725
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unladen Wt.* .........."
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
      Left            =   6660
      TabIndex        =   82
      Top             =   1275
      Width           =   1800
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RL Wt.* ................"
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
      Left            =   6660
      TabIndex        =   81
      Top             =   1035
      Width           =   1680
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total No. of RIMS"
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
      Left            =   6660
      TabIndex        =   80
      Top             =   555
      Width           =   1500
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seating Capacity.."
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
      Left            =   9150
      TabIndex        =   79
      Top             =   585
      Width           =   1575
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size..............."
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
      Left            =   3675
      TabIndex        =   78
      Top             =   3660
      Width           =   1260
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size.................."
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
      Left            =   3675
      TabIndex        =   77
      Top             =   3900
      Width           =   1440
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
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
      Left            =   2745
      TabIndex        =   76
      Top             =   4148
      Width           =   360
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trade No.............."
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
      Left            =   180
      TabIndex        =   75
      Top             =   2700
      Width           =   1620
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Tyres"
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
      Left            =   4665
      TabIndex        =   74
      Top             =   4148
      Width           =   960
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Front Tyres...."
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
      Left            =   180
      TabIndex        =   73
      Top             =   3660
      Width           =   1695
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Middle Tyres"
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
      Left            =   180
      TabIndex        =   72
      Top             =   3900
      Width           =   1560
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rear Tyres....."
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
      Left            =   180
      TabIndex        =   71
      Top             =   4140
      Width           =   1725
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer..........."
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
      Index           =   14
      Left            =   180
      TabIndex        =   70
      Top             =   3420
      Width           =   1785
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warranty(Kms)......"
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
      Left            =   9150
      TabIndex        =   69
      Top             =   1770
      Width           =   1680
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Rate"
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
      Left            =   180
      TabIndex        =   68
      Top             =   3165
      Width           =   825
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Active Y/N.............."
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
      Left            =   3915
      TabIndex        =   67
      Top             =   2940
      Width           =   2280
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Category...."
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
      Left            =   180
      TabIndex        =   66
      Top             =   2220
      Width           =   1590
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wheel Category....."
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
      Left            =   180
      TabIndex        =   65
      Top             =   2940
      Width           =   1680
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Division........."
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
      Left            =   180
      TabIndex        =   64
      Top             =   2460
      Width           =   1770
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Description*"
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
      Left            =   180
      TabIndex        =   63
      Top             =   1500
      Width           =   1620
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Ind.*..............."
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
      Left            =   3480
      TabIndex        =   62
      Top             =   1020
      Width           =   1905
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Type*.........."
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
      Left            =   180
      TabIndex        =   61
      Top             =   1020
      Width           =   1680
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis Type* ......."
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
      Left            =   180
      TabIndex        =   60
      Top             =   780
      Width           =   1725
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Code* .........."
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
      Left            =   180
      TabIndex        =   55
      Top             =   540
      Width           =   1770
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Group*..........."
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
      Left            =   180
      TabIndex        =   54
      Top             =   1980
      Width           =   1845
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Designation*"
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
      Left            =   180
      TabIndex        =   51
      Top             =   1260
      Width           =   1545
   End
End
Attribute VB_Name = "frmModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte, mFlag As Byte
Dim RstItem As ADODB.Recordset
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset, RstModelGrp As ADODB.Recordset
Dim RsCol As ADODB.Recordset

Private Const Model = 0, Sales_Desc = 1, Grp_Code = 2, Grp_Name = 3, Wheel_Catg = 4, Div_Code = 5, Div_Name = 6
Private Const Chas_Type = 7, Model_Type = 8, Model_Ind = 9, Model_Desc = 10, Model_Desc1 = 11
Private Const Cat_Code = 12, Cat_Name = 13, Active_YN = 14, SaleRate = 15, RLW = 16, Warr_KMS = 17
Private Const Warr_Mth = 21, Manufacturer = 22, Tyre_R = 23, Tyre_M = 24, Tyre_F = 25, Tyres = 26
Private Const Trade_NO = 28, Tyre_RS = 29, Tyre_MS = 30, Tyre_FS = 31, Seat = 32, Rims = 33, Unladen_Wt = 34
Private Const Gross_Wt = 35, Front_A_Wt = 36, Rear_A_Wt = 37, HorsePower = 38, WHEELBASE = 39, FUEL = 40, Cylinder = 27
Private Const VehicleType As Byte = 18, TyreDetails As Byte = 19, GearBox As Byte = 20
Private Const ServiceTaxYN As Byte = 41

Private Const Regulatory As Byte = 42
Private Const SteeringType As Byte = 43
Private Const ColourName As Byte = 44
Private Const VehicleDrive As Byte = 45
Private Const FuelTankCapacity As Byte = 46
Private Const RearAxleMake As Byte = 47
Private Const FMSN As Byte = 48
Private Const CubicCapacity As Byte = 49
Private Const BodyType As Byte = 50




Private Const ItemCode As Byte = 1
Private Const Description As Byte = 2
Private Const DefVal As Byte = 3
Private Const PIndex As Byte = 4
Dim ForeColorSelEnter$
Dim BackColorSelLeave$


Dim GridKey As Integer
Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem
Dim OldTrnType As String

Private Sub CmdUpdatePriceList_Click()
    ShowForm FrmVehiclePriceList, True, "Vehicle Price List"
End Sub

Private Sub DGItem_Click()
On Error GoTo ELoop
    If RstItem.RecordCount > 0 Then
        TxtGrid(0).TEXT = RstItem!Name
    End If
    TxtGridValid_Description
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
    DGItem.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
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
        Case DefVal
           FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
End If

If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case Description
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
        Case DefVal
            If FGrid.TextMatrix(FGrid.Row, Description) <> "" Then
                Call GridDblClick(Me, FGrid, TxtGrid, 0)
                TAddMode = False
            End If
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
SetMaxLength
Select Case FGrid.Col
    Case Description
        Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
    Case DefVal
        If FGrid.TextMatrix(FGrid.Row, Description) <> "" Then
           Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        End If
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

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
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
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
Me.top = 0: Me.left = 0: Me.width = 11940: Me.height = 7640: Ini_Grid
TopCtrl1.Tag = PubUParam
Set RstMain = New ADODB.Recordset

    Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
        sitecond = "and LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
        sitecond = ""
    End If
    
If PubMoveRecYn Then
    RstMain.Open "Select MODEL AS SEARCHCODE,MODEL.* FROM MODEL where (Div_Code='" & PubDivCode & "' or Div_Code='') " & sitecond & "Order by Model", GCn, adOpenDynamic, adLockOptimistic
Else
    RstMain.Open "Select Top 1 MODEL AS SEARCHCODE,MODEL.* FROM MODEL where (Div_Code='" & PubDivCode & "' or Div_Code='') " & sitecond & " Order by Model", GCn, adOpenDynamic, adLockOptimistic
End If

Set RstHelp = New ADODB.Recordset
RstHelp.Open "Select * FROM MODEL where (Div_Code='" & PubDivCode & "' or Div_Code='') Order by Model", GCn, adOpenDynamic, adLockOptimistic
Set DgModel.DataSource = RstHelp

Set RstModelGrp = New ADODB.Recordset
RstModelGrp.Open "Select MODEL_GRP.*,MODEL_CAT.ModelCat_Name,DIVISION.Div_Name " & _
                "From (MODEL_GRP Left Join MODEL_CAT On MODEL_GRP.ModelCat_Code=MODEL_CAT.ModelCat_Code) " & _
                "LEFT JOIN DIVISION ON left(MODEL_GRP.ModelGrp_Code,1)=DIVISION.Div_Code " & _
                " where left(MODEL_GRP.ModelGrp_Code,1)='" & PubDivCode & "' Order by MODEL_GRP.ModelGrp_Name", GCn, adOpenDynamic, adLockOptimistic
Set DgModelGrp.DataSource = RstModelGrp

Set RsCol = New ADODB.Recordset
RsCol.CursorLocation = adUseClient
RsCol.Open "select Col_code as code,col_Desc  as name from colmast order by col_Desc", GCn, adOpenDynamic, adLockOptimistic
Set DGCol.DataSource = RsCol

Set RstItem = New ADODB.Recordset
RstItem.Open "Select Item_Code as Code,Item_Description as Name,Default_Value,Report_Index FROM ModelCheckListMast Order by Item_Description", GCn, adOpenDynamic, adLockOptimistic
Set DGItem.DataSource = RstItem
RstItem.Sort = "Name"

'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0
mFlag = 0
LvVehicleType
FrModel.Visible = False
FrModelGrp.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing: Set RstModelGrp = Nothing
    Set RsCol = Nothing
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To txt.Count - 1
        txt(I).Enabled = Enb
    Next
    
    txt(Grp_Code).Enabled = False
    txt(Cat_Code).Enabled = False
    txt(Cat_Name).Enabled = False
    txt(Div_Code).Enabled = False
    txt(Div_Name).Enabled = False
    txt(Wheel_Catg).Enabled = False
End Sub
Private Sub MoveRec()
On Error GoTo ErrLoop
Grid_Hide
RST_BOF_EOF RstMain
TopCtrl1.tDel = False
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    CopyDetails RstMain!Model
End If
Exit Sub
ErrLoop:        MsgBox err.Description
End Sub

Private Sub ListView_Click()
txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
FrmList.Visible = False
txt(Val(ListView.Tag)).SetFocus
End Sub
Public Sub TopCtrl1_eAdd()
On Error GoTo ErrLoop
Dim ModelCopyFrom$
ModelCopyFrom = txt(Model)
BlankText
ADDFLAG = 1
Disp_Text SETS("ADD", Me, RstMain)
txt(Model).Tag = txt(Model)
If RstMain.RecordCount > 0 Then
    CopyDetails ModelCopyFrom
End If
txt(Active_YN) = "Yes"
Txt_GotFocus Model
txt(Model).SetFocus
Exit Sub
ErrLoop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ErrLoop
If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RstMain)
    FGrid.AddItem FGrid.Rows
    txt(Model).Enabled = False
    txt(Chas_Type).Tag = txt(Chas_Type)
    Txt_GotFocus Chas_Type
    txt(Chas_Type).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
ErrLoop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
Dim XBM, mTrans As Boolean
On Error GoTo eloop1
        If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            GCn.BeginTrans
            mTrans = True
            XBM = RstMain.Bookmark
            GCn.Execute ("Delete From MODEL Where Model=" & Chk_Text(Trim(txt(Model))))
            GCn.CommitTrans
            mTrans = False
            
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
eloop1:
    If mTrans Then GCn.RollbackTrans
    CheckError
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
    
      Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    GSQL = "SELECT MODEL as searchcode,MODEL,Model_Desc,Model_Desc1,Model_Desc2,Chas_Type,Model_Type,Sales_Desc,Wheel_Catg,CYLINDER,FUEL,Manufacturer FROM MODEL where (Div_Code='" & PubDivCode & "' or Div_Code='') " & sitecond & " Order by Model"

    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_ePrn()
Dim rep As CrystalReport, Form1 As frmMastList
    Set Form1 = New frmMastList
    With Form1
        .g_FormID = 12
        .LblName.CAPTION = Me.CAPTION
        .CAPTION = Me.CAPTION
        .Show
    End With
    Set Form1 = Nothing
    Set rep = Nothing
End Sub
Private Sub TopCtrl1_eSave()
Dim transFlag As Byte, I As Integer
On Error GoTo ErrLoop
    transFlag = 0
    If IsValid(txt(Model), "Model") = False Then Txt_GotFocus Model: Exit Sub
    If IsValid(txt(Chas_Type), "Chassis Type") = False Then Txt_GotFocus Chas_Type: Exit Sub
    If IsValid(txt(Model_Type), "Model Type") = False Then Txt_GotFocus Model_Type: Exit Sub
    If IsValid(txt(Model_Ind), "Model Ind.") = False Then Txt_GotFocus Model_Ind: Exit Sub
    If IsValid(txt(Sales_Desc), "Sale Designation") = False Then Txt_GotFocus Sales_Desc: Exit Sub
    If IsValid(txt(Model_Desc), "Model Description") = False Then Txt_GotFocus Model_Desc: Exit Sub
    If IsValid(txt(Grp_Name), "Model Group") = False Then Txt_GotFocus Grp_Name: Exit Sub
    
    If IsValid(txt(RLW), "RLW") = False Then Exit Sub
    If IsValid(txt(Unladen_Wt), "Unladen Wt.") = False Then Exit Sub
    If IsValid(txt(Gross_Wt), "Gross Wt.") = False Then Exit Sub
    If IsValid(txt(Front_A_Wt), "Front Axle Wt") = False Then Exit Sub
    If IsValid(txt(Rear_A_Wt), "Rear Axle Wt") = False Then Exit Sub
    
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From MODEL Where Model=" & Chk_Text(Trim(txt(Model)))).Fields(0) > 0 Then MsgBox "Code Already Exists", vbInformation, "Code Validation": Txt_GotFocus Model: txt(Model).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    If TopCtrl1.TopText2 = "Add" Then
        GCn.Execute ("DELETE From MODEL Where Model=" & Chk_Text(Trim(txt(Model))))
        GCn.Execute ("Insert Into MODEL (MODEL,Chas_Type,Model_Type,Model_Ind,Sales_Desc,Model_Desc,Model_Desc1,Grp_Code,Cat_Code,Div_Code,Active_YN,TYRES,TYRE_R,TYRE_M,TYRE_F,TYRE_RS,TYRE_MS,TYRE_FS,RIMS,RLW,SEAT,HORSEPOWER,FRONT_A_WT,REAR_A_WT,UNLADEN_WT,GROSS_WT,WHEELBASE,CYLINDER,FUEL,TRADE_NO,Manufacturer,Warr_KMS,Warr_Mth,Wheel_Catg,Site_Code,U_Name,U_EntDt,U_AE,Vehicle_Type,TyreDetails,GearBoxNo,ServiceTax_YN,Sale_Rate,RegulatoryCertificate,SteeringType,Col_Code,Vehicle_Drive,FuelTankCapacity,RearAxleMake,FMSN,CubicCapacity,BodyType) Values(" & _
            Chk_Text(txt(Model)) & "," & Chk_Text(txt(Chas_Type)) & "," & Chk_Text(txt(Model_Type)) & "," & _
            VNull(txt(Model_Ind)) & "," & Chk_Text(txt(Sales_Desc)) & "," & Chk_Text(txt(Model_Desc)) & "," & _
            Chk_Text(txt(Model_Desc1)) & "," & Chk_Text(txt(Grp_Code)) & " , " & Chk_Text(txt(Cat_Code)) & "," & _
            Chk_Text(txt(Div_Code)) & "," & IIf(txt(Active_YN) = "Yes", 1, 0) & "," & VNull(txt(Tyres)) & "," & _
            VNull(txt(Tyre_R)) & "," & VNull(txt(Tyre_M)) & "," & VNull(txt(Tyre_F)) & "," & _
            Chk_Text(txt(Tyre_RS)) & "," & Chk_Text(txt(Tyre_MS)) & "," & Chk_Text(txt(Tyre_FS)) & "," & _
            VNull(txt(Rims)) & "," & Chk_Text(txt(RLW)) & "," & VNull(txt(Seat)) & "," & Chk_Text(txt(HorsePower)) & "," & _
            Chk_Text(txt(Front_A_Wt)) & "," & Chk_Text(txt(Rear_A_Wt)) & "," & Chk_Text(txt(Unladen_Wt)) & "," & Chk_Text(txt(Gross_Wt)) & "," & _
            VNull(txt(WHEELBASE)) & "," & VNull(txt(Cylinder)) & "," & Chk_Text(txt(FUEL)) & "," & Chk_Text(txt(Trade_NO)) & "," & _
            Chk_Text(txt(Manufacturer)) & "," & VNull(txt(Warr_KMS)) & "," & VNull(txt(Warr_Mth)) & "," & Chk_Text(txt(Wheel_Catg)) & ", '" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "','" & txt(VehicleType) & "','" & txt(TyreDetails) & "','" & txt(GearBox) & "'," & IIf(txt(ServiceTaxYN) = "Yes", 1, 0) & "," & Val(txt(SaleRate)) & ",'" & txt(Regulatory).TEXT & "','" & txt(SteeringType).TEXT & "','" & txt(ColourName).Tag & "','" & txt(VehicleDrive).TEXT & "'," & Val(txt(FuelTankCapacity).TEXT) & ",'" & txt(RearAxleMake).TEXT & "','" & txt(FMSN).TEXT & "','" & txt(CubicCapacity).TEXT & "','" & txt(BodyType).TEXT & "')")
    Else
        GCn.Execute ("update MODEL set MODEL=" & Chk_Text(txt(Model)) & ",Chas_Type=" & Chk_Text(txt(Chas_Type)) & ",Model_Type=" & Chk_Text(txt(Model_Type)) & ",Model_Ind= " & VNull(txt(Model_Ind)) & "," & _
            " Sales_Desc=" & Chk_Text(txt(Sales_Desc)) & ",Model_Desc=" & Chk_Text(txt(Model_Desc)) & ",Model_Desc1= " & Chk_Text(txt(Model_Desc1)) & ",Grp_Code=" & Chk_Text(txt(Grp_Code)) & " ,Cat_Code=" & Chk_Text(txt(Cat_Code)) & ", " & _
            " Div_Code=" & Chk_Text(txt(Div_Code)) & ",Active_YN=" & IIf(txt(Active_YN) = "Yes", 1, 0) & ",TYRES=" & VNull(txt(Tyres)) & ",TYRE_R=" & VNull(txt(Tyre_R)) & ",TYRE_M=" & VNull(txt(Tyre_M)) & ", " & _
            " TYRE_F=" & VNull(txt(Tyre_F)) & ",TYRE_RS=" & Chk_Text(txt(Tyre_RS)) & ",TYRE_MS=" & Chk_Text(txt(Tyre_MS)) & ",TYRE_FS=" & Chk_Text(txt(Tyre_FS)) & ",RIMS=" & VNull(txt(Rims)) & "," & _
            " RLW=" & Chk_Text(txt(RLW)) & ",SEAT=" & VNull(txt(Seat)) & ",HORSEPOWER=" & Chk_Text(txt(HorsePower)) & ",FRONT_A_WT=" & Chk_Text(txt(Front_A_Wt)) & ",REAR_A_WT=" & Chk_Text(txt(Rear_A_Wt)) & ", " & _
            " UNLADEN_WT=" & Chk_Text(txt(Unladen_Wt)) & ",GROSS_WT=" & Chk_Text(txt(Gross_Wt)) & ",WHEELBASE= " & VNull(txt(WHEELBASE)) & ",CYLINDER=" & VNull(txt(Cylinder)) & ",FUEL=" & Chk_Text(txt(FUEL)) & ", " & _
            " TRADE_NO=" & Chk_Text(txt(Trade_NO)) & ",Manufacturer=" & Chk_Text(txt(Manufacturer)) & ",Warr_KMS=" & VNull(txt(Warr_KMS)) & ",Warr_Mth=" & VNull(txt(Warr_Mth)) & ",Wheel_Catg=" & Chk_Text(txt(Wheel_Catg)) & ", " & _
            " RegulatoryCertificate=" & Chk_Text(txt(Regulatory)) & ",SteeringType=" & Chk_Text(txt(SteeringType)) & ",Col_Code=" & Chk_Text(txt(ColourName).Tag) & ",Vehicle_Drive=" & Chk_Text(txt(VehicleDrive)) & ",FuelTankCapacity=" & VNull(txt(FuelTankCapacity)) & ", " & _
            " RearAxleMake=" & Chk_Text(txt(RearAxleMake)) & ",FMSN=" & Chk_Text(txt(FMSN)) & ",CubicCapacity=" & Chk_Text(txt(CubicCapacity)) & ",BodyType=" & Chk_Text(txt(BodyType)) & "," & _
            " Site_Code='" & PubSiteCode & "',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "', Vehicle_Type='" & txt(VehicleType) & "', TyreDetails='" & txt(TyreDetails) & "', GearBoxNo='" & txt(GearBox) & "',ServiceTax_YN=" & IIf(txt(ServiceTaxYN) = "Yes", 1, 0) & " ,sale_Rate=" & Val(txt(SaleRate)) & " Where Model=" & Chk_Text(Trim(txt(Model))) & " and Div_Code='" & PubDivCode & "'")
    End If
    'Add/Edit Check List Items
    GCn.Execute ("delete from ModelCheckList where Model='" & txt(Model) & "'")
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, ItemCode) <> "" Then
            GCn.Execute ("insert into ModelCheckList(MODEL,Item_Code,Default_Value," & _
                " Site_Code, U_Name, U_EntDt, U_AE) " & _
                " values('" & txt(Model) & "','" & FGrid.TextMatrix(I, ItemCode) & "','" & FGrid.TextMatrix(I, DefVal) & _
                "','" & PubSiteCode & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')")
        End If
    Next
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select MODEL AS SEARCHCODE,MODEL.* FROM MODEL where (Div_Code='" & PubDivCode & "' or Div_Code='') And MODEL = " & Chk_Text(Trim(txt(Model))) & " Order by Model")
    End If
    RstHelp.Requery
    RstMain.FIND ("Model=" & Chk_Text(Trim(txt(Model))))
    If ADDFLAG = 1 Then
        BlankText
        Txt_GotFocus Model
        txt(Model).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        'CtrlClckCol
        ADDFLAG = 0
        FrModel.Visible = False
        FrModelGrp.Visible = False
    End If
Exit Sub
ErrLoop:    If transFlag = 1 Then GCn.RollbackTrans
            CheckError
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo ErrLoop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        'CtrlClckCol
        FrModel.Visible = False
        FrModelGrp.Visible = False
    End If
Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eRef()
    RstItem.Requery
    RstHelp.Requery
    RstModelGrp.Requery
    LvVehicleType
    RsCol.Requery
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub ModelSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "Model >=" & Chk_Text(left(Trim(txt(Model)), RstHelp.Fields("Model").DefinedSize))
'RstHelp.FIND "Model LIKE " & Chk_Text(XNull(Trim(Txt(Model))))
End Sub

Private Sub ModelGrp_NameSearch()
If RstModelGrp.RecordCount <= 0 Then Exit Sub
RstModelGrp.MoveFirst
RstModelGrp.FIND "ModelGrp_Name >=" & Chk_Text(XNull(txt(Grp_Name)))
End Sub
Private Sub DGCol_Click()
    DGCol.Visible = False
    If RsCol.RecordCount > 0 Then
         txt(ColourName) = RsCol!Name
         txt(ColourName).Tag = RsCol!Code
    End If
   txt(ColourName).SetFocus
End Sub



Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case Model
            If FrModelGrp.Visible = True Then FrModelGrp.Visible = False
            FrModel.Visible = True
            FrModel.top = txt(Index).top + txt(Index).height + 10
            FrModel.left = txt(Index).left
            FrModel.ZOrder 0
        Case Grp_Name
            If FrModel.Visible = True Then FrModel.Visible = False
            FrModelGrp.Visible = True
            FrModelGrp.top = txt(Index).top + txt(Index).height + 10
            FrModelGrp.left = txt(Index).left
            FrModelGrp.ZOrder 0
        Case ColourName
            If DGCol.Visible = True Then DGCol.Visible = False
            DGCol.Visible = True
            DGCol.top = txt(Index).top + txt(Index).height + 10
            DGCol.left = txt(Index).left
            DGCol.ZOrder 0
    End Select
End If
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
    Ctrl_GetFocus txt(Index)
    mFlag = 0
    If FrModelGrp.Visible = True Then FrModelGrp.Visible = False
    If FrModel.Visible = True Then FrModel.Visible = False
    Select Case Index
        Case Grp_Name
            RST_BOF_EOF RstModelGrp
        Case Model
            RST_BOF_EOF RstHelp
    End Select
    txt(Index).Tag = txt(Index)
    Txt_Click Index
    Select Case Index
        Case Model
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
        Case Grp_Name
            If RstModelGrp.BOF Or RstModelGrp.EOF Then Exit Sub
    End Select
    DgModelGrp.Columns(0).width = 1000.1: DgModelGrp.Columns(1).width = 3435.024: DgModelGrp.Columns(2).width = 1000.1
    Select Case Index
        Case Model
            DgModelGrp.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Model ASC"
            RstHelp.Bookmark = mBookMark
            ModelSearch
        Case Grp_Name
            DgModelGrp.Columns(0).width = 0
            mBookMark = RstModelGrp.Bookmark
            RstModelGrp.Sort = "ModelGrp_Name ASC"
            RstModelGrp.Bookmark = mBookMark
            ModelGrp_NameSearch
        Case VehicleType
            OldTrnType = txt(VehicleType).TEXT
        Case ColourName
            TxtGrid(0).MaxLength = 15
            If RsCol.RecordCount = 0 Or (RsCol.EOF = True Or RsCol.BOF = True) Or txt(ColourName) = "" Then Exit Sub
            If txt(ColourName) <> RsCol!Name Then
                RsCol.MoveFirst
                RsCol.FIND "name ='" & txt(ColourName) & "'"
            End If
    End Select
    If txt(Index) = "" Then Txt_Change Index
End Sub

Private Sub Txt_Click(Index As Integer)
    'CtrlClckCol
    txt(Index).ForeColor = CtrlFCol: txt(Index).BackColor = CtrlBCol
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer, LVHeight As Integer
Select Case Index
    Case Model_Ind, Rims, Tyres, Tyre_R, Tyre_M, Tyre_F, Seat, Cylinder, Warr_Mth
        NumDown txt(Index), KeyCode, 2, 0
    Case RLW, Front_A_Wt, Rear_A_Wt, Unladen_Wt, Gross_Wt, WHEELBASE
        NumDown txt(Index), KeyCode, 4, 2
    Case Warr_KMS
        NumDown txt(Index), KeyCode, 6, 0
    Case ColourName
        DGridTxtKeyDown DGCol, txt, Index, RsCol, KeyCode, False, 1
        If KeyCode = vbKeyReturn Then SendKeysA vbKeyTab, True
    Case VehicleType
        If ListView.ListItems.Count > 20 Then
            LVHeight = 20 * 300
        Else
            LVHeight = ListView.ListItems.Count * 300
        End If
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, LVHeight
            
End Select
Select Case Index
    Case Grp_Name
        If FrModelGrp.Visible = True Then
            Select Case KeyCode
                Case vbKeyUp
                    If Not RstModelGrp.BOF Then RstModelGrp.MovePrevious
                Case vbKeyDown
                    If Not RstModelGrp.EOF Then RstModelGrp.MoveNext
                Case 33
                    For I = 1 To 9
                        If Not RstModelGrp.BOF Then RstModelGrp.MovePrevious
                    Next
                Case 34
                    For I = 1 To 9
                        If Not RstModelGrp.EOF Then RstModelGrp.MoveNext
                    Next
                Case 13
                    SendKeysA vbKeyTab, True
            End Select
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, 33, 34
                    RST_BOF_EOF RstModelGrp
                    If Not RstModelGrp.BOF And Not RstModelGrp.EOF Then
                        txt(Grp_Code) = XNull(RstModelGrp!ModelGrp_Code)
                        txt(Grp_Name) = XNull(RstModelGrp!ModelGrp_Name)
                        txt(Cat_Code) = XNull(RstModelGrp!ModelCat_Code)
                        txt(Cat_Name) = XNull(RstModelGrp!ModelCat_NAME)
                        txt(Div_Code) = left(RstModelGrp!ModelGrp_Code, 1)
                        txt(Div_Name) = XNull(RstModelGrp!Div_Name)
                        txt(Wheel_Catg) = XNull(RstModelGrp!Wheel_Catg)
                        txt(Grp_Name).SelStart = 0
                    End If
            End Select
        End If
End Select
Select Case Index
    Case Model
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case Chas_Type
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp And TopCtrl1.TopText2 = "Add" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case VehicleType
        If ListView.Visible = False Then
            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
                SendKeysA vbKeyTab, True
                KeyCode = 0
            ElseIf KeyCode = vbKeyUp Then
                SendKeys "+{Tab}"
                KeyCode = 0
            End If
        End If
    Case TyreDetails, GearBox, Sales_Desc, SaleRate, Model_Type, Model_Ind, Model_Desc, Model_Desc1, Active_YN, SaleRate, RLW, Warr_KMS, Manufacturer, Tyre_R, Tyre_M, Tyre_F, Tyres, Trade_NO, Tyre_RS, Tyre_MS, Tyre_FS, Seat, Rims, Unladen_Wt, Gross_Wt, Front_A_Wt, Rear_A_Wt, HorsePower, WHEELBASE, FUEL, Cylinder, Regulatory, SteeringType, VehicleDrive, FuelTankCapacity, RearAxleMake, FMSN, CubicCapacity, ServiceTaxYN, Warr_Mth
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case Grp_Name
        If FrModelGrp.Visible = False Then
            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
                SendKeysA vbKeyTab, True
                KeyCode = 0
            ElseIf KeyCode = vbKeyUp Then
                SendKeys "+{Tab}"
                KeyCode = 0
            End If
        End If
    Case BodyType
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
                Txt_Validate Warr_Mth, result
                If result = True Then Txt_GotFocus Warr_Mth: txt(Warr_Mth).SetFocus: Exit Sub
                TopCtrl1_eSave
            Else
                Txt_Click Warr_Mth
                Txt_GotFocus Warr_Mth
                txt(Warr_Mth).SetFocus
            End If
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    Case ColourName
'        If DGCol.Visible = False Then
'            If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
'                If MsgBox("Save Record Yes/No", vbYesNo, "Save Record") = vbYes Then
'                    Txt_Validate Warr_Mth, result
'                    If result = True Then Txt_GotFocus Warr_Mth: Txt(Warr_Mth).SetFocus: Exit Sub
'                    TopCtrl1_eSave
'                Else
'                    Txt_Click Warr_Mth
'                    Txt_GotFocus Warr_Mth
'                    Txt(Warr_Mth).SetFocus
'                End If
'            ElseIf KeyCode = vbKeyUp Then
'                SendKeys "+{Tab}"
'                KeyCode = 0
'            End If
'        End If
End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case Model_Ind, Rims, Tyres, Tyre_R, Tyre_M, Tyre_F, Seat, Cylinder, Warr_Mth
        NumPress txt(Index), KeyAscii, 2, 0
    Case WHEELBASE
        NumPress txt(Index), KeyAscii, 4, 2
    Case Warr_KMS
        NumPress txt(Index), KeyAscii, 6, 0
    Case FuelTankCapacity
        NumPress txt(Index), KeyAscii, 10, 0
    Case ColourName
        If DGCol.Visible = True Then DGridTxtKeyPress txt, Index, RsCol, KeyAscii, "Name"
    Case SaleRate
        'Call NumPress(Txt(Index), KeyAscii, 7, 2)
        'MsgBox "Invalid Action ! Update Vehicle Rate List from Vehicle Menu", vbInformation + vbOKOnly
        'Txt(ServiceTaxYN).SetFocus
    
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
mFlag = 0
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, 33, 34
        Exit Sub
End Select
Select Case Index
    Case Model
        ModelSearch
    Case Grp_Name
        ModelGrp_NameSearch
    Case VehicleType
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case Active_YN, ServiceTaxYN
        If Len(txt(Index)) = 0 Or UCase(mID(txt(Index), 1, 1)) = "Y" Then
            txt(Index) = "Yes"
        ElseIf UCase(mID(txt(Index), 1, 1)) = "N" Then
            txt(Index) = "No"
        Else
            txt(Index) = "Yes"
        End If
    Case Tyre_R, Tyre_M, Tyre_F
        txt(Tyres) = Val(txt(Tyre_R)) + Val(txt(Tyre_M)) + Val(txt(Tyre_F))
        txt(Rims) = Val(txt(Tyre_R)) + Val(txt(Tyre_M)) + Val(txt(Tyre_F))
    Case ColourName
        'If KeyCode <> 13 And DGCol.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsCol, KeyCode, "name", True
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
    Select Case Index
        Case VehicleType
            If txt(Index).TEXT <> "" Then txt(Index).TEXT = ListView.SelectedItem.TEXT
        Case Model
            Set Rst = GCn.Execute("SELECT * FROM MODEL WHERE Model=" & Chk_Text(txt(Model)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Code Already Exists", vbInformation, "Validation": txt(Model) = txt(Model).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Model <> RstMain!Model Then MsgBox "Code Already Exists", vbInformation, "Validation": txt(Model) = txt(Model).Tag: Cancel = True: Exit Sub
                End If
            End If
               'TYRES,TYRE_F,TYRE_M,TYRE_R,RIMS,TYRE_RS,TYRE_MS,TYRE_FS
        Case Grp_Name
            If Not RstModelGrp.EOF And Not RstModelGrp.BOF Then
                txt(Grp_Code) = XNull(RstModelGrp!ModelGrp_Code): txt(Grp_Name) = XNull(RstModelGrp!ModelGrp_Name): txt(Cat_Code) = XNull(RstModelGrp!ModelCat_Code): txt(Cat_Name) = XNull(RstModelGrp!ModelCat_NAME): txt(Div_Code) = left(RstModelGrp!ModelGrp_Code, 1): txt(Div_Name) = XNull(RstModelGrp!Div_Name): txt(Wheel_Catg) = XNull(RstModelGrp!Wheel_Catg)
            Else
                txt(Grp_Code) = "": txt(Grp_Name) = "": txt(Cat_Code) = "": txt(Cat_Name) = "": txt(Div_Code) = "": txt(Div_Name) = "": txt(Wheel_Catg) = ""
            End If
            If UCase(txt(Wheel_Catg)) = "TWO" Then
                txt(Tyre_M).Visible = False
                txt(Tyre_MS).Visible = False
                Lbl(5).Visible = False
                Lbl(9).Visible = False
            Else
                txt(Tyre_M).Visible = True
                txt(Tyre_MS).Visible = True
                Lbl(5).Visible = True
                Lbl(9).Visible = True
            End If
    End Select
Set Rst = Nothing
End Sub

Private Sub DGModelGrp_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If mFlag = 1 Then
    txt(Grp_Code) = DgModelGrp.Columns(0).TEXT: txt(Grp_Name) = DgModelGrp.Columns(1).TEXT: txt(Cat_Code) = DgModelGrp.Columns(2).TEXT: txt(Cat_Name) = DgModelGrp.Columns(3).TEXT: txt(Div_Code) = DgModelGrp.Columns(4).TEXT: txt(Div_Name) = DgModelGrp.Columns(5).TEXT: txt(Wheel_Catg) = DgModelGrp.Columns(6).TEXT
End If
End Sub

Private Sub DGModelGrp_GotFocus()
    mFlag = 1
End Sub

Private Sub BlankText()
Dim I As Integer
    For I = 0 To 40
        txt(I).TEXT = ""
    Next I
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select MODEL AS SEARCHCODE,MODEL.* FROM MODEL where (Div_Code='" & PubDivCode & "' or Div_Code='') And MODEL = '" & MyValue & "' Order by Model")
    End If
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub LvVehicleType()
Set GRs = New Recordset
Set GRs = GCn.Execute("Select Vehicle_Type as Name From Vehicle_Type")
Set mListItem = ListView_Items_RecordSet(ListView, txt, VehicleType, GRs)
Set GRs = Nothing
ListView.width = txt(VehicleType).width
End Sub

Private Sub CopyDetails(ModelCode As String)
Dim RST1 As ADODB.Recordset, I As Integer
Set GRs = New Recordset
Set GRs = GCn.Execute("Select * from Model where Model='" & ModelCode & "'")

    txt(Model) = XNull(GRs!Model)
    txt(Chas_Type) = XNull(GRs!Chas_Type)
    txt(Model_Type) = XNull(GRs!Model_Type)
    txt(Model_Ind) = VNull(GRs!Model_Ind)
    txt(Sales_Desc) = XNull(GRs!Sales_Desc)
    txt(Model_Desc) = XNull(GRs!Model_Desc)
    txt(Model_Desc1) = XNull(GRs!Model_Desc1)
    txt(Grp_Code) = XNull(GRs!Grp_Code)
    txt(Cat_Code) = XNull(GRs!Cat_Code)
    txt(Div_Code) = XNull(GRs!Div_Code)
    txt(Active_YN) = IIf(GRs!Active_YN = 1, "Yes", "No")
    txt(ServiceTaxYN) = IIf(GRs!ServiceTax_YN = 1, "Yes", "No")
    
    txt(Tyres) = VNull(GRs!Tyres)
    txt(Tyre_R) = VNull(GRs!Tyre_R)
    txt(Tyre_M) = VNull(GRs!Tyre_M)
    txt(Tyre_F) = VNull(GRs!Tyre_F)
    txt(Tyre_RS) = XNull(GRs!Tyre_RS)
    txt(Tyre_MS) = XNull(GRs!Tyre_MS)
    txt(Tyre_FS) = XNull(GRs!Tyre_FS)
    txt(TyreDetails) = XNull(GRs!TyreDetails)
    txt(Rims) = VNull(GRs!Rims)
    txt(RLW) = VNull(GRs!RLW)
    txt(Seat) = VNull(GRs!Seat)
    txt(HorsePower) = XNull(GRs!HorsePower)
    txt(Front_A_Wt) = VNull(GRs!Front_A_Wt)
    txt(Rear_A_Wt) = VNull(GRs!Rear_A_Wt)
    txt(Unladen_Wt) = VNull(GRs!Unladen_Wt)
    txt(Gross_Wt) = VNull(GRs!Gross_Wt)
    txt(WHEELBASE) = VNull(GRs!WHEELBASE)
    txt(Cylinder) = VNull(GRs!Cylinder)
    txt(FUEL) = XNull(GRs!FUEL)
    txt(Trade_NO) = XNull(GRs!Trade_NO)
    txt(Manufacturer) = XNull(GRs!Manufacturer)
    txt(Warr_KMS) = VNull(GRs!Warr_KMS)
    txt(Warr_Mth) = VNull(GRs!Warr_Mth)
    txt(VehicleType) = IIf(IsNull(GRs!Vehicle_Type), "", GRs!Vehicle_Type)
    txt(GearBox) = XNull(GRs!GearBoxNo)
    
    
    txt(Regulatory).TEXT = XNull(GRs!RegulatoryCertificate)
    txt(SteeringType).TEXT = XNull(GRs!SteeringType)
    txt(VehicleDrive).TEXT = XNull(GRs!Vehicle_Drive)
    txt(FuelTankCapacity).TEXT = XNull(GRs!FuelTankCapacity)
    txt(RearAxleMake).TEXT = XNull(GRs!RearAxleMake)
    txt(FMSN).TEXT = XNull(GRs!FMSN)
    txt(CubicCapacity).TEXT = XNull(GRs!CubicCapacity)
    txt(BodyType).TEXT = XNull(GRs!BodyType)
    txt(ColourName).Tag = XNull(GRs!Col_Code)
    If txt(ColourName).Tag <> "" Then
        If GCn.Execute("Select Col_desc From ColMast where Col_Code='" & txt(ColourName).Tag & "'").RecordCount > 0 Then
            txt(ColourName).TEXT = GCn.Execute("Select Col_desc From ColMast where Col_Code='" & txt(ColourName).Tag & "'").Fields(0).Value
        Else
            txt(ColourName).Tag = ""
            txt(ColourName).TEXT = ""
        End If
    Else
        txt(ColourName).TEXT = ""
    End If
    
    If GCn.Execute("SELECT s_RATE FROM VEH_RATE WHERE MODEL='" & txt(Model) & "'" & " AND EFFECTIVE_DATE<=" & ConvertDate(date) & "").RecordCount = 1 Then
        txt(SaleRate) = GCn.Execute("SELECT s_RATE FROM VEH_RATE WHERE MODEL='" & txt(Model) & "'" & " AND EFFECTIVE_DATE<=" & ConvertDate(date) & "").Fields(0).Value
    ElseIf GCn.Execute("SELECT s_RATE FROM VEH_RATE WHERE MODEL='" & txt(Model) & "'" & " AND EFFECTIVE_DATE<=" & ConvertDate(date) & "").RecordCount > 1 Then
        txt(SaleRate) = GCn.Execute("SELECT s_RATE FROM VEH_RATE WHERE MODEL='" & txt(Model) & "'" & " AND EFFECTIVE_DATE<=" & ConvertDate(date) & "" & " ORDER BY EFFECTIVE_DATE desc").Fields(0).Value
    Else
        txt(SaleRate) = 0
       
    End If
    txt(SaleRate) = Format(VNull(GRs!Sale_Rate), "0.00")
    Set RST1 = GCn.Execute("Select MODEL_GRP.*,MODEL_CAT.ModelCat_Name,DIVISION.Div_Name " & _
                    "From (MODEL_GRP Left Join MODEL_CAT On MODEL_GRP.ModelCat_Code=MODEL_CAT.ModelCat_Code) " & _
                    "LEFT JOIN DIVISION ON left(MODEL_GRP.ModelGrp_Code,1)=DIVISION.Div_Code " & _
                    "WHERE ModelGrp_Code =" & Chk_Text(GRs!Grp_Code))
    If RST1.RecordCount > 0 Then
        txt(Grp_Name) = XNull(RST1!ModelGrp_Name)
        txt(Div_Name) = XNull(RST1!Div_Name)
        txt(Cat_Name) = XNull(RST1!ModelCat_NAME)
        txt(Wheel_Catg) = XNull(RST1!Wheel_Catg)
    Else
        txt(Grp_Name) = ""
        txt(Div_Name) = ""
        txt(Cat_Name) = ""
        txt(Wheel_Catg) = ""
    End If
    If UCase(txt(Wheel_Catg)) = "TWO" Then
        txt(Tyre_M).Visible = False
        txt(Tyre_MS).Visible = False
        Lbl(5).Visible = False
        Lbl(9).Visible = False
    Else
        txt(Tyre_M).Visible = True
        txt(Tyre_MS).Visible = True
        Lbl(5).Visible = True
        Lbl(9).Visible = True
    End If
    Set RST1 = Nothing
    Set GRs = New Recordset
    Set GRs = GCn.Execute("Select MCL.Item_Code,MCLM.Item_Description,MCL.Default_Value,MCLM.Report_Index " & _
                    " from ModelCheckList MCL " & _
                    " left join ModelCheckListMast MCLM on MCL.Item_Code=MCLM.Item_Code" & _
                    " where MCL.Model='" & ModelCode & "' Order by MCLM.Report_Index")
    FGrid.Rows = 1
    If GRs.RecordCount > 0 Then
        I = 1
        Do Until GRs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, 0) = I
                .TextMatrix(I, ItemCode) = GRs!Item_Code
                .TextMatrix(I, Description) = GRs!Item_Description
                .TextMatrix(I, DefVal) = GRs!Default_Value
                .TextMatrix(I, PIndex) = GRs!Report_Index
            End With
            GRs.MoveNext
            I = I + 1
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    Set GRs = Nothing
End Sub

Private Sub Ini_Grid()
    With FGrid
        .left = 100
        .top = 4950
        .width = 8000
        .Cols = 5
        .RowHeightMin = PubGridRowHeight
        .height = .RowHeight(0) * 8
        
        .TextMatrix(0, 0) = "S.No."
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 550

        .TextMatrix(0, ItemCode) = "ItemCode"
        .ColAlignment(ItemCode) = flexAlignLeftCenter
        .ColWidth(ItemCode) = 0
        
        .TextMatrix(0, Description) = "Description"
        .ColAlignment(Description) = flexAlignLeftCenter
        .ColWidth(Description) = 4000
                
        .TextMatrix(0, DefVal) = "Defalt Value"
        .ColAlignment(DefVal) = flexAlignLeftCenter
        .ColWidth(DefVal) = 1500
        
        .TextMatrix(0, PIndex) = "PrintInd"
        .ColAlignment(PIndex) = flexAlignRightCenter
        .ColWidth(PIndex) = 1000
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    DGItem.top = mTopScale: DGItem.left = mLtScale
    DGCol.top = mTopScale: DGCol.left = mLtScale
End Sub

Private Sub Grid_Hide()
    If DGItem.Visible = True Then DGItem.Visible = False
    If DGCol.Visible = True Then DGCol.Visible = False
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
Grid_Hide
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
Select Case FGrid.Col
    Case Description
        If RstItem.RecordCount = 0 Or (RstItem.EOF = True Or RstItem.BOF = True) Then Exit Sub
        If FGrid.TextMatrix(FGrid.Row, Description) <> "" Then
            RstItem.MoveFirst
            RstItem.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, Description) & "'"
            If RstItem.EOF = True Then RstItem.MoveFirst
        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    TxtGrid(0).TEXT = TxtGrid(0).Tag
    TxtGrid_KeyUp Index, KeyCode, Shift
    FGrid.SetFocus
    TxtGrid(0).Visible = False
    Grid_Hide
    Exit Sub
End If
Select Case FGrid.Col
    Case DefVal
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, DefVal
            End If
        End If
    Case Description
        DGridTxtKeyDown DGItem, TxtGrid, Index, RstItem, KeyCode, True, 1
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then
                 GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, DefVal
            End If
        End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case FGrid.Col
    Case Description
        If DGItem.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RstItem, KeyAscii, "name"
End Select
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case Description
        If KeyCode <> 13 And DGItem.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RstItem, KeyCode, "name", True
    Case PIndex
        FGrid.TextMatrix(FGrid.Row, PIndex) = TxtGrid(Index)
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
Select Case FGrid.Col
    Case Description
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        TxtGridValid_Description
    Case DefVal
        FGrid.TextMatrix(FGrid.Row, DefVal) = TxtGrid(Index)
    Case ColourName
        If RsCol.RecordCount = 0 Or (RsCol.EOF = True Or RsCol.BOF = True) Or TxtGrid(0).TEXT = "" Then
            txt(ColourName).TEXT = ""
            txt(ColourName).Tag = ""
        Else
            txt(ColourName).TEXT = RsCol!Name
            txt(ColourName).Tag = RsCol!Code
        End If
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
End Function

Private Function ChkDuplicate() As Boolean
Dim I As Integer
Dim X As String, Y As String
Dim Col1 As Byte, Col2 As Byte, Col3 As Byte
Select Case FGrid.Col
    Case Description
        Col1 = ItemCode
        Col2 = Description
    End Select
    X = UCase(Trim(TxtGrid(0).TEXT))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, FGrid.Col))))
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

Private Sub SetMaxLength()
Select Case FGrid.Col   'Index
    Case Description
        TxtGrid(0).MaxLength = 25
    Case DefVal
        TxtGrid(0).MaxLength = 10
    Case PIndex
        TxtGrid(0).MaxLength = 2
    Case Else
        TxtGrid(0).MaxLength = 0
End Select
End Sub

Private Sub TxtGridValid_Description()
If RstItem.RecordCount = 0 Or (RstItem.EOF = True Or RstItem.BOF = True) Or TxtGrid(0).TEXT = "" Then
    FGrid.TextMatrix(FGrid.Row, ItemCode) = ""
    FGrid.TextMatrix(FGrid.Row, Description) = ""
    FGrid.TextMatrix(FGrid.Row, DefVal) = ""
    FGrid.TextMatrix(FGrid.Row, PIndex) = ""
Else
    FGrid.TextMatrix(FGrid.Row, ItemCode) = IIf(IsNull(RstItem!Code), "", RstItem!Code)
    FGrid.TextMatrix(FGrid.Row, Description) = IIf(IsNull(RstItem!Name), "", RstItem!Name)
    FGrid.TextMatrix(FGrid.Row, DefVal) = IIf(IsNull(RstItem!Default_Value), "", RstItem!Default_Value)
    FGrid.TextMatrix(FGrid.Row, PIndex) = IIf(IsNull(RstItem!Report_Index), "", RstItem!Report_Index)
End If
If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
End Sub
