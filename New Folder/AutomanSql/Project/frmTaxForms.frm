VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmTaxForms 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Tax Forms"
   ClientHeight    =   6600
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
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   11880
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   7080
      TabIndex        =   51
      Text            =   "999.99"
      Top             =   1800
      Width           =   675
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
      Height          =   210
      Index           =   14
      Left            =   4500
      MaxLength       =   50
      TabIndex        =   48
      Top             =   3735
      Width           =   3660
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   4500
      MaxLength       =   40
      TabIndex        =   45
      Top             =   2295
      Width           =   675
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   9045
      TabIndex        =   40
      Top             =   5595
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   135
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   120
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
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   2865
      Left            =   9885
      Negotiate       =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Form_Code"
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
         DataField       =   "Form_Desc"
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
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3509.858
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGTaxAc 
      Height          =   2865
      Left            =   8205
      Negotiate       =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6015
      Visible         =   0   'False
      Width           =   5340
      _ExtentX        =   9419
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
         Caption         =   "A/c Name"
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
            ColumnWidth     =   30.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4710.047
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   3165
      Left            =   8595
      Negotiate       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6045
      Visible         =   0   'False
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   5583
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
         DataField       =   "Name"
         Caption         =   "A/c Name"
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
            ColumnWidth     =   30.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4710.047
         EndProperty
      EndProperty
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
      Height          =   210
      Index           =   2
      Left            =   4500
      MaxLength       =   25
      TabIndex        =   2
      Text            =   "0123456789012345678901234"
      Top             =   855
      Width           =   3165
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
      Height          =   210
      Index           =   3
      Left            =   4500
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1095
      Width           =   3165
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
      Height          =   210
      Index           =   12
      Left            =   4500
      MaxLength       =   50
      TabIndex        =   13
      Top             =   3975
      Width           =   3660
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
      Height          =   210
      Index           =   11
      Left            =   4500
      MaxLength       =   50
      TabIndex        =   12
      Top             =   3495
      Width           =   3660
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
      Height          =   210
      Index           =   10
      Left            =   4500
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3255
      Width           =   3660
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
      Height          =   210
      Index           =   8
      Left            =   4500
      MaxLength       =   7
      TabIndex        =   9
      Text            =   "VisFals"
      Top             =   2775
      Visible         =   0   'False
      Width           =   1350
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
      Height          =   210
      Index           =   7
      Left            =   4500
      MaxLength       =   8
      TabIndex        =   8
      Top             =   2535
      Width           =   1350
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   4500
      MaxLength       =   40
      TabIndex        =   7
      Top             =   2055
      Width           =   675
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   4500
      TabIndex        =   6
      Text            =   "999.99"
      Top             =   1815
      Width           =   675
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
      Height          =   210
      Index           =   4
      Left            =   4500
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "Local"
      Top             =   1335
      Width           =   1350
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
      Height          =   210
      Index           =   9
      Left            =   4500
      MaxLength       =   8
      TabIndex        =   10
      Text            =   "Yes"
      Top             =   3015
      Width           =   1350
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
      Height          =   210
      Index           =   1
      Left            =   4500
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "0123"
      Top             =   615
      Width           =   675
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
      Height          =   210
      Index           =   0
      Left            =   4500
      MaxLength       =   12
      TabIndex        =   5
      Text            =   "Pur/Sal/Perm"
      Top             =   1575
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SFC Percentage"
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
      Left            =   5640
      TabIndex        =   52
      Top             =   1800
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Tax A/c Name*"
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
      Left            =   1950
      TabIndex        =   50
      Top             =   3750
      Width           =   2220
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   4260
      TabIndex        =   49
      Top             =   3735
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   4260
      TabIndex        =   47
      Top             =   2310
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Tax Percentage"
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
      Left            =   1950
      TabIndex        =   46
      Top             =   2310
      Width           =   2235
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   4260
      TabIndex        =   44
      Top             =   3975
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   4260
      TabIndex        =   43
      Top             =   3495
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   4260
      TabIndex        =   42
      Top             =   3240
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Code*"
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
      Left            =   1950
      TabIndex        =   21
      Top             =   630
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Description*"
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
      Left            =   1950
      TabIndex        =   15
      Top             =   870
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Purchase / Sale A/c*"
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
      Left            =   1950
      TabIndex        =   37
      Top             =   3975
      Width           =   2475
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surcharge A/c Name*"
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
      Left            =   1950
      TabIndex        =   36
      Top             =   3510
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax A/c Name*"
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
      Left            =   1950
      TabIndex        =   35
      Top             =   3270
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt / Issue / Both"
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
      Left            =   1950
      TabIndex        =   34
      Top             =   2790
      Visible         =   0   'False
      Width           =   1875
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   4260
      TabIndex        =   33
      Top             =   2790
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Transaction Type*"
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
      Left            =   1950
      TabIndex        =   32
      Top             =   2550
      Width           =   2070
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   4260
      TabIndex        =   31
      Top             =   2550
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surcharge Percentage"
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
      Left            =   1950
      TabIndex        =   30
      Top             =   2070
      Width           =   1905
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   4260
      TabIndex        =   29
      Top             =   2070
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Percentage"
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
      Left            =   1950
      TabIndex        =   28
      Top             =   1830
      Width           =   1335
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   4260
      TabIndex        =   27
      Top             =   1830
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local / Central*"
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
      Left            =   1950
      TabIndex        =   26
      Top             =   1350
      Width           =   1365
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   4260
      TabIndex        =   25
      Top             =   1350
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   4260
      TabIndex        =   24
      Top             =   1110
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Description"
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
      Left            =   1950
      TabIndex        =   23
      Top             =   1110
      Width           =   1665
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   4260
      TabIndex        =   22
      Top             =   615
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   4260
      TabIndex        =   20
      Top             =   1590
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase / Sale / Permit*"
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
      Left            =   1950
      TabIndex        =   19
      Top             =   1590
      Width           =   2205
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   4260
      TabIndex        =   18
      Top             =   3030
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Applicable For*"
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
      Left            =   1950
      TabIndex        =   17
      Top             =   3030
      Width           =   1305
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   4260
      TabIndex        =   14
      Top             =   870
      Width           =   75
   End
End
Attribute VB_Name = "frmTaxForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim TAddMode As Boolean
Dim RsParty As ADODB.Recordset
Dim RsTaxAc As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim mSearchCode As String
Dim ListArray As Variant
Dim mListItem As ListItem
Dim OldTrnType As String
Dim MyIndex As Byte
Private Const TrnType As Byte = 0               ' Form Type Purchase / Sale
Private Const FormCode As Byte = 1              ' Form Code
Private Const FormDesc As Byte = 2              ' FormDesc
Private Const PrintDesc As Byte = 3             ' PrintDesc
Private Const L_C As Byte = 4                   ' Local / Central
Private Const TaxPer As Byte = 5                ' Tax Percentage
Private Const TaxSurPer As Byte = 6             ' Tax Surcharge Percentage
Private Const FormTrnType As Byte = 7           ' Form Transaction Type
'Private Const FormYN As Byte = 7               ' Form YN
'Private Const RecIss As Byte = 8               ' Tax Percentage
Private Const AppFor As Byte = 9                ' Applicable for Vehicle / Spare
Private Const TaxAcCode As Byte = 10            ' Tax A/c Code
Private Const SurAcCode As Byte = 11            ' Tax Surcharge A/c Code
Private Const PurSalAcCode As Byte = 12         ' Taxable Pur Sale A/c Code
Private Const AddTaxPer As Byte = 13
Private Const AddTaxAc As Byte = 14
Private Const SFCPer As Byte = 15

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    For I = 0 To txt.Count - 1
        If I <> 8 Then txt(I).Enabled = Enb
    Next I
'    txt(12).Enabled = False
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("SearchCode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select T.Form_Code as Code,T.Form_Desc as Name,T.Form_Code+T.Trn_Type As SearchCode,T.*, " & _
            " S.Name as TaxAcName, S1.Name as SurAcName, S2.Name as PurSalAcName " & _
            " from ((((TaxForms T LEFT JOIN TaxFormsAc as T1 on T.Form_Code+'" & PubDivCode & "'=T1.Form_Code+T1.Div_Code ) " & _
            " left join SubGroup S on S.SubCode=T1.Tax_Ac_Code) " & _
            " left join SubGroup S1 on S1.SubCode=T1.Sur_Ac_Code) " & _
            " left join SubGroup S2 on S2.SubCode=T1.PurSal_Ac_Code) " & _
            " Where T.Form_Code+T.Trn_Type = '" & MyValue & "' Order by T.Form_Code,T.Trn_Type,T.Vehicle_YN,T.Spare_YN")
    End If
    MoveRec
    BUTTONS True, Me, Master, 0
Exit Sub
ELoop:
    CheckError
End Sub

'* Used for clear all text boxes used in the form
Private Sub BlankText()
Dim I As Byte
    For I = 0 To txt.Count - 1
        txt(I).TEXT = ""
        txt(I).Tag = ""
    Next I
End Sub

Private Sub Grid_Hide()
    If DGForm.Visible = True Then DGForm.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If DGTaxAc.Visible = True Then DGTaxAc.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Sub MoveRec()
Dim Master1 As ADODB.Recordset, I As Byte
On Error GoTo ELoop
    If Master.RecordCount > 0 Then
        mSearchCode = Master!Form_Code & Master!Trn_Type
        txt(TrnType).TEXT = Master!Trn_Type
        
        txt(FormCode).TEXT = Master!Form_Code
        txt(FormDesc).TEXT = Master!form_desc
        txt(PrintDesc).TEXT = Master!Printing_Desc
        txt(L_C).TEXT = Master!L_C
        txt(TaxPer).TEXT = IIf(IsNull(Master!Tax_Per) Or Master!Tax_Per = 0, "", Format(Master!Tax_Per, "0.00"))
        txt(SFCPer).TEXT = IIf(IsNull(Master!SFCPer) Or Master!SFCPer = 0, "", Format(Master!SFCPer, "0.00"))
        txt(TaxSurPer).TEXT = IIf(IsNull(Master!Tax_Sur_Per) Or Master!Tax_Sur_Per = 0, "", Format(Master!Tax_Sur_Per, "0.00"))
        txt(AddTaxPer).TEXT = IIf(IsNull(Master!AddTaxPer) Or Master!AddTaxPer = 0, "", Format(Master!AddTaxPer, "0.00"))
        
        txt(FormTrnType).TEXT = FxFormTrnType(Master!FormTrnType)
'            Txt(FormYN).Text = IIf(Master!Form_YN = 0, "No", "Yes")
'            Txt(RecIss).Text = Master!rec_iss
        If Master!Vehicle_YN = 1 Then
            txt(AppFor) = "Vehicle"
        Else
            txt(AppFor) = "Spare"
        End If
        If txt(TrnType) = "Sale" Then
            Label3(15).CAPTION = txt(AppFor) & " Sale A/c"
            RsParty.Requery
            RsParty.Filter = "MainGrCode like '" & pubSalSysMainGrCode & "*'"
        Else
            Label3(15).CAPTION = txt(AppFor) & " Purchase A/c"
            RsParty.Requery
            RsParty.Filter = "MainGrCode like '" & pubPurSysMainGrCode & "*'"
        End If

        'from TaxFormsAc
        Set Master1 = New ADODB.Recordset
        Master1.CursorLocation = adUseClient
        Set Master1 = GCn.Execute("Select T1.*, " & _
            " S.Name as TaxAcName, S1.Name as SurAcName, S2.Name as PurSalAcName, S3.Name As AddTaxAcName " & _
            " from (((TaxFormsAc T1 LEFT JOIN SubGroup S on T1.Tax_Ac_Code=S.SubCode) " & _
            " left join SubGroup S1 on T1.Sur_Ac_Code=S1.SubCode) " & _
            " left join SubGroup S2 on T1.PurSal_Ac_Code=S2.SubCode) " & _
            " left join SubGroup S3 on T1.AddTaxAc=S3.SubCode " & _
            " where T1.Form_Code+T1.Div_Code='" & Master!Form_Code & PubDivCode & "'")
        If Master1.RecordCount > 0 Then
            txt(TaxAcCode).Tag = IIf(IsNull(Master1!Tax_Ac_Code), "", Master1!Tax_Ac_Code)
            txt(TaxAcCode).TEXT = IIf(IsNull(Master1!TaxAcName), "", Master1!TaxAcName)
            txt(SurAcCode).Tag = IIf(IsNull(Master1!Sur_Ac_Code), "", Master1!Sur_Ac_Code)
            txt(SurAcCode).TEXT = IIf(IsNull(Master1!SurAcName), "", Master1!SurAcName)
            txt(AddTaxAc).Tag = IIf(IsNull(Master1!AddTaxAc), "", Master1!AddTaxAc)
            txt(AddTaxAc).TEXT = IIf(IsNull(Master1!AddTaxAcName), "", Master1!AddTaxAcName)
            txt(PurSalAcCode).Tag = IIf(IsNull(Master1!PurSal_Ac_Code), "", Master1!PurSal_Ac_Code)
            txt(PurSalAcCode).TEXT = IIf(IsNull(Master1!PurSalAcName), "", Master1!PurSalAcName)
        Else
            txt(TaxAcCode).Tag = ""
            txt(TaxAcCode).TEXT = ""
            txt(SurAcCode).Tag = ""
            txt(SurAcCode).TEXT = ""
            txt(PurSalAcCode).Tag = ""
            txt(PurSalAcCode).TEXT = ""
            txt(AddTaxAc).Tag = ""
            txt(AddTaxAc) = ""
        End If
    Else
        BlankText
    End If
    Set Master1 = Nothing
    RDisp Master, Me
    TopCtrl1.tDel = False
    TopCtrl1.tPrn = False
    Grid_Hide
Exit Sub
ELoop:
    Set Master1 = Nothing
    CheckError
End Sub

Private Sub DGTaxAc_Click()
    DGTaxAc.Visible = False
    If RsTaxAc.RecordCount > 0 Then
        txt(Val(DGForm.Tag)).TEXT = RsTaxAc!Name
        txt(Val(DGForm.Tag)).Tag = RsTaxAc!Code
    End If
    txt(Val(DGForm.Tag)).SetFocus
    DGForm.Tag = ""
End Sub

Private Sub DGParty_Click()
    DGParty.Visible = False
    If RsParty.RecordCount > 0 Then
        txt(MyIndex).TEXT = RsParty!Name
        txt(MyIndex).Tag = RsParty!Code
    End If
    txt(MyIndex).SetFocus
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
    CheckError
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
    TopCtrl1.Tag = PubUParam    '"AEDP"
    WinSetting Me
    For I = 0 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg '&HDFF4F2
        txt(I).ForeColor = CtrlFColOrg
    Next
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Party_Type,AcGroup.MainGrCode from SubGroup " & _
        "left join  " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,3) in ('" & pubPurSysMainGrCode & "','" & pubSalSysMainGrCode & "') " & _
        "order by SubGroup.name"
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Party_Type from SubGroup " & _
        "left join  " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) in ('" & pubTaxSysMainGrCode & "') " & _
        "order by SubGroup.name"
    Set RsTaxAc = New ADODB.Recordset
    RsTaxAc.CursorLocation = adUseClient
    RsTaxAc.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGTaxAc.DataSource = RsTaxAc
    
    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    If PubMoveRecYn Then
        Set Master = GCn.Execute("Select T.Form_Code as Code,T.Form_Desc as Name,T.Form_Code+T.Trn_Type As SearchCode,T.*, " & _
            " S.Name as TaxAcName, S1.Name as SurAcName, S2.Name as PurSalAcName " & _
            " from ((((TaxForms T LEFT JOIN TaxFormsAc as T1 on T.Form_Code+'" & PubDivCode & "'=T1.Form_Code+T1.Div_Code ) " & _
            " left join SubGroup S on S.SubCode=T1.Tax_Ac_Code) " & _
            " left join SubGroup S1 on S1.SubCode=T1.Sur_Ac_Code) " & _
            " left join SubGroup S2 on S2.SubCode=T1.PurSal_Ac_Code) " & _
            " Order by T.Form_Code,T.Trn_Type,T.Vehicle_YN,T.Spare_YN")
    Else
        Set Master = GCn.Execute("Select Top 1 T.Form_Code as Code,T.Form_Desc as Name,T.Form_Code+T.Trn_Type As SearchCode,T.*, " & _
            " S.Name as TaxAcName, S1.Name as SurAcName, S2.Name as PurSalAcName " & _
            " from ((((TaxForms T LEFT JOIN TaxFormsAc as T1 on T.Form_Code+'" & PubDivCode & "'=T1.Form_Code+T1.Div_Code ) " & _
            " left join SubGroup S on S.SubCode=T1.Tax_Ac_Code) " & _
            " left join SubGroup S1 on S1.SubCode=T1.Sur_Ac_Code) " & _
            " left join SubGroup S2 on S2.SubCode=T1.PurSal_Ac_Code) " & _
            " Order by T.Form_Code,T.Trn_Type,T.Vehicle_YN,T.Spare_YN")
    End If
    Set DGForm.DataSource = Master
    
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    
Exit Sub
ELoop:
    CheckError
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

Private Sub Form_Resize()
    DGForm.left = 6500: DGForm.top = mTopScale
    DGParty.left = 6600: DGParty.top = mTopScale
    DGTaxAc.left = 6600: DGTaxAc.top = mTopScale: DGTaxAc.height = 2350
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RsParty = Nothing
    Set RsTaxAc = Nothing
    Set Master = Nothing
End Sub

Private Sub ListView_Click()
txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
FrmList.Visible = False
txt(Val(ListView.Tag)).SetFocus
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    txt(FormCode).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
Dim I As Integer
    Disp_Text SETS("EDIT", Me, Master)
    'Txt(12).Enabled = True
    txt(TrnType).Enabled = False
    txt(FormCode).Enabled = False
    txt(FormDesc).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim vBook As Variant
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            vBook = Master.AbsolutePosition
            GCn.BeginTrans
            GCn.Execute ("Delete From TaxForms Where FormCode='" & txt(FormCode).TEXT & "'")
            GCn.CommitTrans
            Master.Requery
            If Master.RecordCount > 0 Then
                If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
            End If
            BUTTONS True, Me, Master, 0
            MoveRec
        End If
    Else
        MsgBox "No Records To Delete!", vbInformation, "Information"
    End If
Exit Sub
ELoop:
    GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub

Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub

Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    'If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    
    GSQL = "Select T.Form_Code+T.Trn_Type as SearchCode,Form_Code,Form_Desc,Printing_Desc," & _
        "L_C as Local_Central,Trn_Type," & cCStr("Vehicle_YN") & " as Vehicle, " & cCStr("Spare_YN") & " as Spare " & _
        " from TaxForms T Order by Form_Code,Trn_Type,Vehicle_YN,Spare_YN"
        
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
     FAFind.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eRef()
On Error GoTo ELoop
    RsParty.Requery
    RsTaxAc.Requery
    Master.Requery
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean, ReOrderQty As Double
Dim Rst As ADODB.Recordset, DocIdHlp As String, TmpStr As String, VehYn As Byte, SprYN As Byte
On Error GoTo ELoop
    Grid_Hide
    
    If IsValid(txt(TrnType), Label3(4).CAPTION) = False Then Exit Sub
    If IsValid(txt(FormCode), Label3(5).CAPTION) = False Then Exit Sub
    If IsValid(txt(FormDesc), Label3(3).CAPTION) = False Then Exit Sub
    If IsValid(txt(L_C), Label3(8).CAPTION) = False Then Exit Sub
'    If IsValid(Txt(FormYN), Label3(11).CAPTION) = False Then Exit Sub
    If IsValid(txt(FormTrnType), Label3(11).CAPTION) = False Then Exit Sub
    If IsValid(txt(AppFor), Label3(1).CAPTION) = False Then Exit Sub
    If txt(TrnType) <> "Permit" Then
        If IsValid(txt(TaxAcCode), Label3(13).CAPTION) = False Then Exit Sub
        If IsValid(txt(SurAcCode), Label3(14).CAPTION) = False Then Exit Sub
        If IsValid(txt(PurSalAcCode), Label3(15).CAPTION) = False Then Exit Sub
    End If
    
    If txt(AppFor) = "Vehicle" Then
        VehYn = 1
    Else
        SprYN = 1
    End If
    
    If TopCtrl1.TopText2 = "Add" Then
        If GCn.Execute("Select count(*) From TaxForms Where Form_Code='" & txt(FormCode) & "'").Fields(0) > 0 Then
            MsgBox "Form Code " & txt(FormCode) & " Already Exists", vbCritical, "Validation Error"
            txt(FormCode).SetFocus
            Exit Sub
        End If
    End If
    GCn.BeginTrans
        mTrans = True
        If TopCtrl1.TopText2 = "Add" Then
            GSQL = "Insert into TaxForms (Trn_Type, Form_Code, Form_Desc, Printing_Desc, " & _
                " L_C, Tax_Per, Tax_Sur_Per, AddTaxPer, FormTrnType, " & _
                " Vehicle_YN, Spare_YN, U_Name, U_EntDt, U_AE,SFCPer )" & _
                " values ('" & txt(TrnType) & "', '" & txt(FormCode) & "', '" & txt(FormDesc) & _
                "', '" & txt(PrintDesc) & "', '" & txt(L_C) & "', " & Val(txt(TaxPer)) & _
                ", " & Val(txt(TaxSurPer)) & ", " & Val(txt(AddTaxPer)) & ", " & FxFormTrnType(txt(FormTrnType).TEXT) & _
                ", " & VehYn & ", " & SprYN & ", '" & pubUName & "'," & ConvertDate(PubServerDate) & ", 'A', " & Val(txt(SFCPer)) & ")"
        Else    'Edit
            GSQL = "Update TaxForms set Form_Desc= '" & txt(FormDesc) & "', Printing_Desc='" & txt(PrintDesc) & _
                "', L_C='" & txt(L_C) & "', Tax_Per=" & Val(txt(TaxPer)) & ", SFCPer=" & Val(txt(SFCPer)) & ", AddTaxPer = " & Val(txt(AddTaxPer)) & ", Tax_Sur_Per=" & Val(txt(TaxSurPer)) & _
                ", FormTrnType=" & FxFormTrnType(txt(FormTrnType).TEXT) & _
                ", U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E' where form_code='" & txt(FormCode) & "'"
        End If
        GCn.Execute GSQL
        GCn.Execute ("Delete from TaxFormsAc where Form_Code='" & txt(FormCode) & "' and Div_Code='" & PubDivCode & "'")
        GSQL = "Insert into TaxFormsAc (Form_Code,Div_Code," & _
                " Tax_Ac_Code, Sur_Ac_Code, AddTaxAc," & _
                " PurSal_Ac_Code, U_Name, U_EntDt, U_AE )" & _
                " values ('" & txt(FormCode) & "', '" & PubDivCode & _
                "', '" & txt(TaxAcCode).Tag & "', '" & txt(SurAcCode).Tag & "', '" & txt(AddTaxAc).Tag & _
                "','" & txt(PurSalAcCode).Tag & "', '" & pubUName & "'," & ConvertDate(PubServerDate) & ", 'A')"
        GCn.Execute GSQL
    GCn.CommitTrans
    mTrans = False
    mSearchCode = txt(FormCode) & txt(TrnType)
    Master.Requery
    If MasterFormExit Then Unload Me: Exit Sub
    If PubMoveRecYn Then
        Master.FIND "SearchCode = '" & mSearchCode & "'"
    Else
        Set Master = GCn.Execute("Select T.Form_Code as Code,T.Form_Desc as Name,T.Form_Code+T.Trn_Type As SearchCode,T.*, " & _
            " S.Name as TaxAcName, S1.Name as SurAcName, S2.Name as PurSalAcName " & _
            " from ((((TaxForms T LEFT JOIN TaxFormsAc as T1 on T.Form_Code+'" & PubDivCode & "'=T1.Form_Code+T1.Div_Code ) " & _
            " left join SubGroup S on S.SubCode=T1.Tax_Ac_Code) " & _
            " left join SubGroup S1 on S1.SubCode=T1.Sur_Ac_Code) " & _
            " left join SubGroup S2 on S2.SubCode=T1.PurSal_Ac_Code) " & _
            " Where T.Form_Code+T.Trn_Type = '" & mSearchCode & "' " & _
            " Order by T.Form_Code,T.Trn_Type,T.Vehicle_YN,T.Spare_YN")
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    If MasterFormExit Then Unload Me: Exit Sub
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To txt.Count - 1
            txt(I).BackColor = CtrlBColOrg
            txt(I).ForeColor = CtrlFColOrg
        Next
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
On Error GoTo ELoop
Ctrl_GetFocus txt(Index)
Grid_Hide
MyIndex = Index
Select Case Index
    Case TrnType
        ListArray = Array("Purchase", "Sale", "Permit")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 3)
        OldTrnType = txt(TrnType).TEXT
        
    Case FormTrnType
        ListArray = Array("NA", "Issue", "Receipt") ', "Both")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 3)
        OldTrnType = txt(FormTrnType).TEXT
        
    Case AppFor
        ListArray = Array("Vehicle", "Spare")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
        OldTrnType = txt(AppFor)
        
    Case TaxAcCode, SurAcCode, AddTaxAc
        If RsTaxAc.RecordCount = 0 Or (RsTaxAc.EOF = True Or RsTaxAc.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsTaxAc!Name Then
            RsTaxAc.MoveFirst
            RsTaxAc.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case PurSalAcCode
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then Grid_Hide:  Exit Sub

Select Case Index
    Case FormCode
        DGridTxtKeyDown_Mast DGForm, txt, Index, Master, KeyCode, False, 1
    Case TrnType, FormTrnType, AppFor
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 850
    Case L_C
        KeyCode = IIf(KeyCode = vbKeyDelete Or KeyCode = vbKeyBack, 0, KeyCode)
    Case PurSalAcCode
        DGridTxtKeyDown DGParty, txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
    Case TaxAcCode, SurAcCode, AddTaxAc
        DGridTxtKeyDown DGTaxAc, txt, Index, RsTaxAc, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
End Select

If ListView.Visible = False Then
    If DGTaxAc.Visible = False Then
        If DGParty.Visible = False Then
            If (txt(TrnType) = "Permit" And Index = AppFor) Or Index = PurSalAcCode Then
                If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
                    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave: Exit Sub Else txt(Index).SetFocus: Exit Sub
                ElseIf Index = PurSalAcCode Then
                    DGridTxtKeyDown DGParty, txt, PurSalAcCode, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
                End If
            End If
            If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
                Ctrl_DownKeyDown KeyCode, Shift
            ElseIf KeyCode = vbKeyUp Then
                If TopCtrl1.TopText2.CAPTION = "Add" Then
                    If Index <> FormCode Then Ctrl_UpKeyDown KeyCode, Shift
                ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
                    If Index <> FormDesc Then Ctrl_UpKeyDown KeyCode, Shift
                End If
            End If
        End If
    End If
End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
Dim PressChr As String
PressChr = Asc(UCase(Chr(KeyAscii)))
Select Case Index
    Case FormCode
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case L_C
        If PressChr = vbKeyL Then
            txt(Index) = "Local"
        ElseIf PressChr = vbKeyC Then
            txt(Index) = "Central"
        End If
        KeyAscii = 0
    Case TaxPer, TaxSurPer, AddTaxPer
        Call NumPress(txt(Index), KeyAscii, 2, 2)
    Case TaxAcCode, SurAcCode, AddTaxAc
        If DGTaxAc.Visible = True Then DGridTxtKeyPress txt, Index, RsTaxAc, KeyAscii, "Name"
    Case PurSalAcCode
        If DGParty.Visible = True Then DGridTxtKeyPress txt, PurSalAcCode, RsParty, KeyAscii, "Name"
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case TrnType, FormTrnType, AppFor
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
        If Index = TrnType Then Txt_Validate TrnType, False
    Case FormCode
        DGridTxtKeyUp_Mast txt, Index, Master, KeyCode, "Code"
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate txt(Index)
    Select Case Index
        Case TaxPer, TaxSurPer, AddTaxPer
            txt(Index).TEXT = IIf(Val(txt(Index)) <> 0, Format(Val(txt(Index)), "0.00"), "")
        Case TaxAcCode, SurAcCode, AddTaxAc
            DGForm.Tag = Index
    End Select
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, TmpStr As String, ApplyTax As Boolean
Dim I As Integer
On Error GoTo ELoop
Select Case Index
    Case TrnType, FormTrnType
        If txt(Index).TEXT <> "" Then txt(Index).TEXT = ListView.SelectedItem.TEXT
        If Master!Vehicle_YN = 1 Then
            txt(AppFor) = "Vehicle"
        Else
            txt(AppFor) = "Spare"
        End If
        If txt(TrnType) = "Sale" Then
            ApplyTax = True
            Label3(15).CAPTION = txt(AppFor) & " Sale A/c"
            RsParty.Requery
            RsParty.Filter = "MainGrCode like '" & pubSalSysMainGrCode & "*'"
        ElseIf txt(TrnType) = "Purchase" Then
            ApplyTax = True
            Label3(15).CAPTION = txt(AppFor) & " Purchase A/c"
            RsParty.Requery
            RsParty.Filter = "MainGrCode like '" & pubPurSysMainGrCode & "*'"
        End If
        txt(TaxPer).Enabled = ApplyTax
        txt(AddTaxPer).Enabled = ApplyTax
        txt(TaxSurPer).Enabled = ApplyTax
        txt(TaxAcCode).Enabled = ApplyTax
        txt(SurAcCode).Enabled = ApplyTax
        txt(PurSalAcCode).Enabled = ApplyTax

    Case AppFor
        If txt(Index).TEXT <> "" Then txt(Index).TEXT = ListView.SelectedItem.TEXT
'        txt(12).Enabled = False
        If txt(TrnType) = "Sale" Then
            ApplyTax = True
            Label3(15).CAPTION = txt(AppFor) & " Sale A/c"
        ElseIf txt(TrnType) = "Purchase" Then
            ApplyTax = True
            Label3(15).CAPTION = txt(AppFor) & " Purchase A/c"
        End If
    Case TaxAcCode, SurAcCode, AddTaxAc
        If RsTaxAc.RecordCount > 0 Then
            If txt(Index).TEXT <> "" Then
                txt(Index).TEXT = RsTaxAc!Name
                txt(Index).Tag = RsTaxAc!Code
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        Else
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        End If
    Case PurSalAcCode
        If RsParty.RecordCount > 0 Then
            If txt(Index).TEXT <> "" Then
                txt(Index).TEXT = RsParty!Name
                txt(Index).Tag = RsParty!Code
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        Else
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        End If
    End Select
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

