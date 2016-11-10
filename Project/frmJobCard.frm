VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmJobCard 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Job Card Entry"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12195
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
   ScaleHeight     =   8835
   ScaleWidth      =   12195
   Begin MSDataGridLib.DataGrid DgInsuranceCompany 
      Height          =   2730
      Left            =   150
      Negotiate       =   -1  'True
      TabIndex        =   164
      TabStop         =   0   'False
      Top             =   7095
      Visible         =   0   'False
      Width           =   4440
      _ExtentX        =   7832
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
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
      Height          =   210
      Index           =   54
      Left            =   1530
      MaxLength       =   20
      TabIndex        =   20
      Top             =   3360
      Width           =   1485
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
      Height          =   210
      Index           =   53
      Left            =   4755
      MaxLength       =   25
      TabIndex        =   21
      Top             =   3360
      Width           =   1620
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
      Height          =   210
      Index           =   52
      Left            =   1530
      MaxLength       =   25
      TabIndex        =   19
      Text            =   "Help"
      Top             =   3135
      Width           =   4845
   End
   Begin MSDataGridLib.DataGrid DGLab 
      Height          =   2730
      Left            =   -1860
      Negotiate       =   -1  'True
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   8490
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
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
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4575.118
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrmCancel 
      Height          =   465
      Left            =   6750
      TabIndex        =   157
      Top             =   -120
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<<- Cancelled ->>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   45
         TabIndex        =   158
         Top             =   150
         Width           =   2130
      End
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   3165
      TabIndex        =   153
      Top             =   8175
      Visible         =   0   'False
      Width           =   2505
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   150
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   75
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
      Index           =   51
      Left            =   1530
      MaxLength       =   15
      TabIndex        =   0
      Top             =   435
      Visible         =   0   'False
      Width           =   1965
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
      Height          =   210
      Index           =   0
      Left            =   3375
      TabIndex        =   31
      Text            =   "9999999"
      Top             =   3900
      Width           =   1050
   End
   Begin MSDataGridLib.DataGrid DGModel 
      Height          =   1155
      Left            =   -1485
      Negotiate       =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   8280
      Visible         =   0   'False
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   2037
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
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "Model Description"
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
            ColumnWidth     =   5325.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   8040.189
         EndProperty
      EndProperty
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
      Index           =   50
      Left            =   900
      MaxLength       =   4
      TabIndex        =   55
      Text            =   "Kms."
      Top             =   3900
      Visible         =   0   'False
      Width           =   510
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
      Index           =   49
      Left            =   8025
      MaxLength       =   25
      TabIndex        =   33
      Text            =   "23-APR-2002"
      Top             =   3900
      Width           =   1305
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
      Height          =   210
      Index           =   48
      Left            =   6585
      MaxLength       =   20
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Extra Field"
      Top             =   435
      Width           =   615
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
      Index           =   46
      Left            =   7665
      MaxLength       =   10
      TabIndex        =   25
      Top             =   2910
      Width           =   1305
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
      Index           =   47
      Left            =   10470
      MaxLength       =   11
      TabIndex        =   26
      Top             =   2910
      Visible         =   0   'False
      Width           =   1305
   End
   Begin MSDataGridLib.DataGrid DGTrouble 
      Height          =   2730
      Left            =   6990
      Negotiate       =   -1  'True
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   8520
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
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
            DividerStyle    =   3
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4575.118
         EndProperty
      EndProperty
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
      Index           =   45
      Left            =   9450
      TabIndex        =   56
      Text            =   "Extra Field"
      Top             =   6465
      Width           =   615
   End
   Begin VB.Frame FrmPrn 
      BackColor       =   &H00CAECF0&
      Caption         =   "Printing Option"
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
      Height          =   1605
      Left            =   1620
      TabIndex        =   130
      Top             =   7440
      Visible         =   0   'False
      Width           =   5025
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00CAECF0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4695
         MousePointer    =   99  'Custom
         Picture         =   "frmJobCard.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "Delete Current Record"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   15
         Picture         =   "frmJobCard.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmJobCard.frx":0678
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3420
         MaskColor       =   &H00FFC0FF&
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmJobCard.frx":0982
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   3420
         MaskColor       =   &H00EFD5B8&
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmJobCard.frx":0C8C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   3420
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Printer "
         Top             =   285
         Width           =   1590
      End
      Begin VB.TextBox txtPrint 
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
         Left            =   7425
         TabIndex        =   135
         Top             =   555
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtPrint 
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
         Left            =   7080
         TabIndex        =   134
         Top             =   285
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtPrint 
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
         Left            =   7470
         TabIndex        =   133
         Top             =   300
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton Optpre 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "PrePrinted "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1725
         TabIndex        =   132
         Top             =   720
         Width           =   1200
      End
      Begin VB.OptionButton OptPlain 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "Plain"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   300
         TabIndex        =   131
         Top             =   720
         Width           =   750
      End
      Begin VB.Line Line8 
         X1              =   1470
         X2              =   1470
         Y1              =   510
         Y2              =   600
      End
      Begin VB.Line Line7 
         X1              =   2820
         X2              =   2820
         Y1              =   630
         Y2              =   735
      End
      Begin VB.Line Line5 
         X1              =   360
         X2              =   360
         Y1              =   615
         Y2              =   720
      End
      Begin VB.Line Line6 
         X1              =   2820
         X2              =   345
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Stationary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   41
         Left            =   -105
         TabIndex        =   143
         Top             =   285
         Width           =   3315
      End
      Begin VB.Label LblPrinter 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Current Active Printer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   330
         TabIndex        =   142
         Top             =   1275
         Width           =   4650
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Printer Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   0
         TabIndex        =   141
         Top             =   0
         Width           =   4695
      End
   End
   Begin MSDataGridLib.DataGrid DGService 
      Height          =   2730
      Left            =   2190
      Negotiate       =   -1  'True
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   7635
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
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
         Caption         =   "Service Desc."
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2640.189
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGInsp 
      Height          =   2775
      Left            =   -315
      Negotiate       =   -1  'True
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   8175
      Visible         =   0   'False
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "code"
         Caption         =   "Sheet No."
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
         DataField       =   "Regno"
         Caption         =   "Reg No."
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGMech 
      Height          =   2865
      Left            =   6495
      Negotiate       =   -1  'True
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   8700
      Visible         =   0   'False
      Width           =   5340
      _ExtentX        =   9419
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
         Caption         =   "Mechanic Name"
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3495.118
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGHist 
      Height          =   2520
      Left            =   2970
      Negotiate       =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   8580
      Visible         =   0   'False
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   4445
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "RegNo"
         Caption         =   "Reg. No."
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "PhoneOff"
         Caption         =   "Phone (O)"
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
         DataField       =   "Govt"
         Caption         =   "Govt"
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
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4004.788
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGBook 
      Height          =   3330
      Left            =   1920
      Negotiate       =   -1  'True
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   8385
      Visible         =   0   'False
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   5874
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Book_No"
         Caption         =   "Book No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Regno"
         Caption         =   "Reg No."
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
         DataField       =   "Chassis"
         Caption         =   "Chassis"
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
         DataField       =   "OName"
         Caption         =   "Owner"
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
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4995.213
         EndProperty
      EndProperty
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
      ForeColor       =   &H00FF00FF&
      Height          =   210
      Index           =   44
      Left            =   4785
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   127
      Top             =   885
      Width           =   1305
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
      Height          =   210
      Index           =   3
      Left            =   1530
      MaxLength       =   25
      TabIndex        =   4
      Top             =   885
      Width           =   1965
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   2730
      Left            =   -510
      Negotiate       =   -1  'True
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   7830
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2505.26
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtgrid2 
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
      Left            =   7455
      MaxLength       =   8
      TabIndex        =   42
      Top             =   5610
      Visible         =   0   'False
      Width           =   705
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
      Left            =   4005
      MaxLength       =   40
      TabIndex        =   40
      Top             =   5670
      Visible         =   0   'False
      Width           =   705
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
      Height          =   210
      Index           =   16
      Left            =   1530
      MaxLength       =   25
      TabIndex        =   15
      Text            =   "Help"
      Top             =   2685
      Width           =   4845
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
      Index           =   25
      Left            =   10635
      TabIndex        =   53
      Top             =   1590
      Width           =   1080
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
      Left            =   1530
      MaxLength       =   40
      TabIndex        =   11
      Text            =   "Help"
      Top             =   1785
      Width           =   4845
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
      Index           =   13
      Left            =   1530
      MaxLength       =   40
      TabIndex        =   12
      Top             =   2010
      Width           =   4845
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
      Left            =   1530
      MaxLength       =   40
      TabIndex        =   13
      Top             =   2235
      Width           =   4845
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
      Index           =   15
      Left            =   1530
      MaxLength       =   40
      TabIndex        =   14
      Top             =   2460
      Width           =   4845
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
      Index           =   6
      Left            =   1530
      MaxLength       =   14
      TabIndex        =   5
      Text            =   "Help"
      Top             =   1110
      Width           =   1965
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
      Index           =   1
      Left            =   2205
      MaxLength       =   8
      TabIndex        =   1
      Top             =   660
      Width           =   1290
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
      Height          =   210
      Index           =   5
      Left            =   10470
      MaxLength       =   8
      TabIndex        =   47
      Text            =   "Help"
      Top             =   2685
      Visible         =   0   'False
      Width           =   1305
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
      Height          =   210
      Index           =   2
      Left            =   4785
      MaxLength       =   25
      TabIndex        =   2
      Top             =   435
      Width           =   1305
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
      Height          =   210
      Index           =   4
      Left            =   4785
      MaxLength       =   25
      TabIndex        =   155
      Top             =   660
      Width           =   1305
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   35
      Left            =   1440
      TabIndex        =   35
      Text            =   "999999.99"
      Top             =   4125
      Width           =   1050
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
      Index           =   34
      Left            =   5610
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Text            =   "Extra Field"
      Top             =   3675
      Width           =   1110
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   36
      Left            =   5610
      TabIndex        =   36
      Top             =   4125
      Width           =   1110
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
      Index           =   37
      Left            =   8025
      MaxLength       =   25
      TabIndex        =   37
      Text            =   "23-APR-2002"
      Top             =   4125
      Width           =   1305
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
      Index           =   41
      Left            =   2310
      MaxLength       =   40
      TabIndex        =   44
      Text            =   "Help"
      Top             =   6450
      Width           =   4275
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
      Index           =   40
      Left            =   2310
      MaxLength       =   40
      TabIndex        =   45
      Text            =   "Help"
      Top             =   6675
      Width           =   4275
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
      Index           =   42
      Left            =   7530
      MaxLength       =   50
      TabIndex        =   46
      Top             =   6225
      Width           =   4275
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
      Index           =   39
      Left            =   2310
      MaxLength       =   50
      TabIndex        =   43
      Top             =   6225
      Width           =   4275
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
      Height          =   210
      Index           =   33
      Left            =   5610
      MaxLength       =   8
      TabIndex        =   32
      Text            =   "01234567"
      Top             =   3900
      Width           =   1110
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
      Left            =   4785
      MaxLength       =   4
      TabIndex        =   10
      Top             =   1560
      Width           =   630
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
      Height          =   210
      Index           =   32
      Left            =   1440
      TabIndex        =   30
      Text            =   "9999999"
      Top             =   3900
      Width           =   1050
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
      Height          =   210
      Index           =   24
      Left            =   8640
      MaxLength       =   40
      TabIndex        =   52
      Top             =   1815
      Width           =   3075
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
      Index           =   27
      Left            =   7665
      MaxLength       =   50
      TabIndex        =   23
      Text            =   "Help"
      Top             =   2460
      Width           =   4110
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
      Height          =   210
      Index           =   7
      Left            =   4785
      MaxLength       =   20
      TabIndex        =   6
      Text            =   "Help"
      Top             =   1110
      Width           =   2400
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
      Height          =   210
      Index           =   19
      Left            =   4725
      MaxLength       =   10
      TabIndex        =   18
      Top             =   2910
      Width           =   1650
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
      Height          =   210
      Index           =   18
      Left            =   3075
      MaxLength       =   25
      TabIndex        =   17
      Top             =   2910
      Width           =   1395
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
      Left            =   1530
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1560
      Width           =   1965
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
      Height          =   210
      Index           =   22
      Left            =   8640
      MaxLength       =   20
      TabIndex        =   50
      Top             =   1365
      Width           =   3075
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
      Height          =   210
      Index           =   9
      Left            =   4785
      MaxLength       =   25
      TabIndex        =   8
      Top             =   1335
      Width           =   2400
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
      Index           =   30
      Left            =   8025
      MaxLength       =   12
      TabIndex        =   28
      Text            =   "012345678901"
      Top             =   3675
      Width           =   1305
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
      Height          =   210
      Index           =   20
      Left            =   8640
      MaxLength       =   8
      TabIndex        =   48
      Text            =   "99999999"
      Top             =   1140
      Width           =   855
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
      Height          =   210
      Index           =   21
      Left            =   10635
      MaxLength       =   25
      TabIndex        =   49
      Top             =   1140
      Width           =   1080
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
      Height          =   210
      Index           =   29
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   27
      Top             =   3675
      Width           =   2985
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
      Height          =   210
      Index           =   17
      Left            =   1530
      MaxLength       =   25
      TabIndex        =   16
      Top             =   2910
      Width           =   1260
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
      Height          =   210
      Index           =   23
      Left            =   8640
      MaxLength       =   8
      TabIndex        =   51
      Top             =   1590
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
      Index           =   8
      Left            =   1530
      TabIndex        =   7
      Text            =   "Help"
      Top             =   1335
      Width           =   1965
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
      Height          =   210
      Index           =   31
      Left            =   10980
      TabIndex        =   29
      Text            =   "9999.99"
      Top             =   3675
      Width           =   780
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
      Index           =   26
      Left            =   7665
      MaxLength       =   10
      TabIndex        =   22
      Text            =   "Help"
      Top             =   2235
      Width           =   1305
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
      Height          =   210
      Index           =   28
      Left            =   7665
      MaxLength       =   25
      TabIndex        =   24
      Top             =   2685
      Width           =   1305
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   1695
      Left            =   75
      TabIndex        =   39
      Top             =   4470
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
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
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid2 
      Height          =   1695
      Left            =   5910
      TabIndex        =   41
      Top             =   4455
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   7
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      GridColorFixed  =   0
      FocusRect       =   0
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "dd"
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
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
      Index           =   38
      Left            =   10980
      TabIndex        =   34
      Text            =   "Extra Field"
      Top             =   3900
      Width           =   780
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
      Height          =   210
      Index           =   43
      Left            =   10980
      MaxLength       =   20
      TabIndex        =   38
      Text            =   "Extra Field"
      Top             =   4125
      Width           =   780
   End
   Begin MSDataGridLib.DataGrid DGDealer 
      Height          =   2730
      Left            =   7110
      Negotiate       =   -1  'True
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   4440
      _ExtentX        =   7832
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
         Caption         =   "Dealer name"
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   384
      Left            =   0
      TabIndex        =   156
      Top             =   0
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   688
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insu. Policy No"
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
      Left            =   60
      TabIndex        =   163
      Top             =   3375
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Expiry Dt."
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
      Index           =   51
      Left            =   3090
      TabIndex        =   162
      Top             =   3360
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Co."
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
      Left            =   45
      TabIndex        =   161
      Top             =   3135
      Width           =   1215
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
      Left            =   6630
      TabIndex        =   160
      Top             =   6720
      Width           =   45
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   30
      X2              =   11730
      Y1              =   4395
      Y2              =   4395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard Type*"
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
      Index           =   49
      Left            =   60
      TabIndex        =   152
      Top             =   420
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label LblJobType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
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
      Left            =   10560
      TabIndex        =   151
      Top             =   465
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobType :"
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
      Left            =   9720
      TabIndex        =   150
      Top             =   465
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hr.Meter"
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
      Index           =   48
      Left            =   2550
      TabIndex        =   149
      Top             =   3900
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Date*"
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
      Index           =   47
      Left            =   6855
      TabIndex        =   148
      Top             =   3915
      Width           =   1140
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
      Height          =   255
      Index           =   35
      Left            =   8415
      TabIndex        =   147
      Top             =   3915
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
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
      Index           =   46
      Left            =   6090
      TabIndex        =   146
      Top             =   450
      Width           =   495
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   45
      Left            =   9060
      TabIndex        =   145
      Top             =   2925
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mfg. Inv. No. "
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
      Index           =   63
      Left            =   6405
      TabIndex        =   144
      Top             =   2910
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Rate"
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
      Index           =   44
      Left            =   4455
      TabIndex        =   129
      Top             =   3690
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Date*"
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
      Index           =   21
      Left            =   3615
      TabIndex        =   128
      Top             =   885
      Width           =   1050
   End
   Begin VB.Label lblDocId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DocID:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   7305
      TabIndex        =   125
      Top             =   690
      Width           =   615
   End
   Begin VB.Label lblPrefix 
      BackStyle       =   0  'Transparent
      Caption         =   "VPrefix"
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
      Height          =   225
      Left            =   1530
      TabIndex        =   124
      Top             =   660
      Width           =   630
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard Entry Completion Time"
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
      Left            =   6615
      TabIndex        =   118
      Top             =   6480
      Width           =   2730
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
      Height          =   270
      Index           =   36
      Left            =   9345
      TabIndex        =   117
      Top             =   6450
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp Del.Time*"
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
      Left            =   9405
      TabIndex        =   116
      Top             =   4140
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Left            =   45
      TabIndex        =   115
      Top             =   2685
      Width           =   345
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   45
      X2              =   11745
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Date*"
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
      Left            =   3615
      TabIndex        =   114
      Top             =   435
      Width           =   855
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
      ForeColor       =   &H00C000C0&
      Height          =   240
      Index           =   30
      Left            =   10530
      TabIndex        =   113
      Top             =   1710
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "History Srl No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   15
      Left            =   9330
      TabIndex        =   112
      Top             =   1590
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insp. Sheet No."
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
      Left            =   9045
      TabIndex        =   111
      Top             =   2715
      Visible         =   0   'False
      Width           =   1335
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
      Index           =   28
      Left            =   1275
      TabIndex        =   110
      Top             =   660
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Time*"
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
      Index           =   42
      Left            =   9405
      TabIndex        =   109
      Top             =   3900
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Est Spares Amt."
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
      Left            =   60
      TabIndex        =   108
      Top             =   4140
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp Del. Dt.*"
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
      Left            =   6855
      TabIndex        =   107
      Top             =   4140
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Est Lab Amt."
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
      Left            =   4455
      TabIndex        =   106
      Top             =   4140
      Width           =   1080
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Advisor Name"
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
      Left            =   75
      TabIndex        =   105
      Top             =   6450
      Width           =   1905
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mechanic Name"
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
      Left            =   75
      TabIndex        =   104
      Top             =   6690
      Width           =   1335
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
      Height          =   270
      Index           =   15
      Left            =   7455
      TabIndex        =   103
      Top             =   6210
      Width           =   45
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   6660
      TabIndex        =   102
      Top             =   6225
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body Damage/Shortage"
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
      Left            =   75
      TabIndex        =   101
      Top             =   6225
      Width           =   2070
   End
   Begin VB.Label LblFuel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6420
      TabIndex        =   100
      Top             =   3900
      Width           =   390
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel in Tank*"
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
      Left            =   4455
      TabIndex        =   99
      Top             =   3915
      Width           =   1140
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
      Height          =   105
      Index           =   10
      Left            =   4980
      TabIndex        =   98
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt. Vehicle"
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
      Left            =   3615
      TabIndex        =   97
      Top             =   1560
      Width           =   1140
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
      Height          =   255
      Index           =   7
      Left            =   1590
      TabIndex        =   96
      Top             =   3915
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kms*"
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
      Left            =   60
      TabIndex        =   95
      Top             =   3915
      Width           =   480
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
      ForeColor       =   &H00C000C0&
      Height          =   240
      Index           =   5
      Left            =   8535
      TabIndex        =   94
      Top             =   1860
      Width           =   45
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Mechanic"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   16
      Left            =   7305
      TabIndex        =   93
      Top             =   1815
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Dt"
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
      Left            =   3615
      TabIndex        =   92
      Top             =   660
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No."
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
      Left            =   45
      TabIndex        =   91
      Top             =   885
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard No.*"
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
      Left            =   60
      TabIndex        =   90
      Top             =   660
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Name "
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
      Left            =   6405
      TabIndex        =   87
      Top             =   2460
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No.*"
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
      Left            =   3615
      TabIndex        =   86
      Top             =   1110
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(M)"
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
      Left            =   4470
      TabIndex        =   85
      Top             =   2910
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(R)"
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
      Index           =   5
      Left            =   2805
      TabIndex        =   84
      Top             =   2910
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(O)"
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
      Left            =   1245
      TabIndex        =   83
      Top             =   2910
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Srl No."
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
      Left            =   45
      TabIndex        =   82
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration No.*"
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
      Left            =   45
      TabIndex        =   81
      Top             =   1110
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name*"
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
      Left            =   45
      TabIndex        =   80
      Top             =   1785
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   1695
      Left            =   7230
      Top             =   420
      Width           =   4560
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division :"
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
      Left            =   7305
      TabIndex        =   79
      Top             =   480
      Width           =   810
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code :"
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
      Left            =   8520
      TabIndex        =   78
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Type*"
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
      Index           =   37
      Left            =   60
      TabIndex        =   77
      Top             =   3690
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Job No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   36
      Left            =   7305
      TabIndex        =   76
      Top             =   1140
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coupon No."
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
      Index           =   35
      Left            =   6855
      TabIndex        =   75
      Top             =   3690
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
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   16
      Left            =   8535
      TabIndex        =   74
      Top             =   1140
      Width           =   75
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
      ForeColor       =   &H00C000C0&
      Height          =   240
      Index           =   12
      Left            =   8535
      TabIndex        =   73
      Top             =   1365
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last KMs/Hrs "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   32
      Left            =   7305
      TabIndex        =   72
      Top             =   1590
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model*"
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
      Left            =   45
      TabIndex        =   71
      Top             =   1335
      Width           =   600
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
      ForeColor       =   &H00C000C0&
      Height          =   240
      Index           =   14
      Left            =   10530
      TabIndex        =   70
      Top             =   1140
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Job Dt."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   34
      Left            =   9525
      TabIndex        =   69
      Top             =   1140
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
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
      Index           =   31
      Left            =   60
      TabIndex        =   68
      Top             =   2925
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No."
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
      Index           =   33
      Left            =   3615
      TabIndex        =   67
      Top             =   1335
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
      ForeColor       =   &H00C000C0&
      Height          =   240
      Index           =   9
      Left            =   8535
      TabIndex        =   66
      Top             =   1590
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Service"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   27
      Left            =   7305
      TabIndex        =   65
      Top             =   1365
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   45
      TabIndex        =   64
      Top             =   2010
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coupon/Srv Value"
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
      Index           =   41
      Left            =   9405
      TabIndex        =   63
      Top             =   3675
      Width           =   1575
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
      Index           =   2
      Left            =   10215
      TabIndex        =   62
      Top             =   900
      Width           =   45
   End
   Begin VB.Label LblTotVeh 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   11250
      TabIndex        =   61
      Top             =   900
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Vehicle for Service Date"
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
      Index           =   7
      Left            =   7305
      TabIndex        =   59
      Top             =   900
      Width           =   2565
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sold Date "
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
      Left            =   6405
      TabIndex        =   58
      Top             =   2685
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Code "
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
      Left            =   6405
      TabIndex        =   57
      Top             =   2235
      Width           =   1140
   End
End
Attribute VB_Name = "frmJobCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TAddMode As Boolean
Dim ExitCtrl As Boolean
Dim GridKey As Integer

Dim VoucherEditFlag As Boolean
Dim JobDocID$
Dim ForSiteCode$
Dim VType As String
Dim MyIndex As Byte
Dim Rst As ADODB.Recordset
Dim Master As ADODB.Recordset

Dim RSBook As ADODB.Recordset
Dim RsInsp As ADODB.Recordset
Dim RsHist As ADODB.Recordset
Dim RsModel As ADODB.Recordset
Dim RsServ As ADODB.Recordset
Dim RsDealer As ADODB.Recordset
Dim RsMech As ADODB.Recordset
Dim RsCity As ADODB.Recordset
Dim RsInsuranceCompany As ADODB.Recordset
Dim RsSuper As ADODB.Recordset
Dim RsTrb As ADODB.Recordset
Dim RsElem As ADODB.Recordset
Dim RsLab As ADODB.Recordset
Private Const BackColorSelEnter$ = &HEBB7EC ' &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$
Dim OldBookingID$
Dim OldInsSheetID$
'Text Box (Form)
Dim mAddFlag$

Private Const JobClDt As Byte = 44
Private Const HrMeter As Byte = 0
Private Const JobNo As Byte = 1
Private Const JobDt As Byte = 2
Private Const BookSrl As Byte = 3
Private Const BookDate As Byte = 4
Private Const Ins_Sheet As Byte = 5
Private Const VehRegNo As Byte = 6
Private Const Chassis As Byte = 7
Private Const Model As Byte = 8
Private Const Engine As Byte = 9
Private Const VehSrlNo As Byte = 10
Private Const GovtYn As Byte = 11
Private Const OwnerName As Byte = 12
Private Const Address1 As Byte = 13
Private Const Address2 As Byte = 14
Private Const Address3 As Byte = 15
Private Const City As Byte = 16
Private Const PhoneOff As Byte = 17
Private Const PhoneResi As Byte = 18
Private Const Mobile As Byte = 19
Private Const LastJobNo As Byte = 20
Private Const LastJobDt As Byte = 21
Private Const LastSrv As Byte = 22
Private Const LastKMS As Byte = 23
Private Const LastMech As Byte = 24
Private Const HistNo As Byte = 25
Private Const DCode As Byte = 26
Private Const DNAME As Byte = 27
Private Const SaleDate As Byte = 28
Private Const SrvType As Byte = 29
Private Const Coupon As Byte = 30
Private Const CouponVal As Byte = 31
Private Const CurrentKMS As Byte = 32
Private Const FUEL As Byte = 33
Private Const SrvRate As Byte = 34
Private Const EstSpr As Byte = 35
Private Const EstLab As Byte = 36
Private Const DelDt             As Byte = 37
Private Const Damage            As Byte = 39
Private Const Mechanic          As Byte = 40
Private Const Supervisor        As Byte = 41
Private Const Remarks           As Byte = 42

Private Const ArrTime           As Byte = 38
Private Const DelTime           As Byte = 43
Private Const JCTime            As Byte = 45
Private Const InvNo             As Byte = 46
Private Const InvDate           As Byte = 47
Private Const RecpTime          As Byte = 48
Private Const ArrDate           As Byte = 49
Private Const KmsHrs            As Byte = 50
Private Const JobType           As Byte = 51
Private Const InsuranceCompany  As Byte = 52
Private Const InsuranceExpiry   As Byte = 53
Private Const InsurancePolicyNo As Byte = 54

'Text Box (Grid)
Private Const mTxtGrid1 As Byte = 1
Private Const mTxtGrid2 As Byte = 1


'Fgrid1 Columns
Private Const Col_Code      As Byte = 1
Private Const Col_Trouble   As Byte = 2
Private Const Col_Repeat    As Byte = 3
Private Const Col_Lab_Code  As Byte = 4
Private Const Col_Lab_Desc  As Byte = 5
Private Const Col_Lab_Rate  As Byte = 6
Private Const Col_Time_Req  As Byte = 7
Private Const Col_Amount    As Byte = 8


'Fgrid2 Columns
Private Const ElCode        As Byte = 1
Private Const ElNature      As Byte = 2
Private Const ElDefault     As Byte = 3
Private Const ElName        As Byte = 4
Private Const ElValue       As Byte = 5

Private Const PWindows      As Byte = 0
Private Const PScreen       As Byte = 1
Private Const PDos          As Byte = 2
Private Const PClose        As Byte = 3
Private Const PSetUp        As Byte = 4
Dim mRepName$
Dim ListArray As Variant
Dim mListItem As ListItem
Private Sub CmdJobType_Click()
End Sub
Private Sub DGBook_Click()
If RSBook.RecordCount > 0 Then
    Txt(BookSrl).TEXT = RSBook!Book_no
    Txt(BookSrl).Tag = RSBook!JobBookingID
    FillBookingData
End If
Txt(BookSrl).SetFocus
DGBook.Visible = False
End Sub
Private Sub DGHist_Click()
If RsHist.RecordCount > 0 Then
    Call History_Field(False, True)
End If
Txt(MyIndex).SetFocus
DGHist.Visible = False
End Sub
Private Sub DGInsp_Click()
If RsInsp.RecordCount > 0 Then
    Txt(Ins_Sheet).TEXT = RsInsp!Code
    Txt(Ins_Sheet).Tag = RsInsp!DocID
    If RsInsp!RegNo <> "" Then
        RsHist.FIND ("Regno='" & RsInsp!RegNo & "'")
    ElseIf RSBook!Chassis <> "" Then
        RsHist.FIND ("chassis='" & RsInsp!Chassis & "'")
    End If
    If RsHist.EOF = True Or RsHist.BOF = True Then
        Txt(HistNo).Tag = ""
        Txt(HistNo).TEXT = ""
        Txt(VehRegNo).TEXT = RsInsp!RegNo
        Txt(Chassis).TEXT = RsInsp!Name
        Txt(Model).Tag = RsInsp!Model
        Txt(Model).TEXT = RsInsp!Model
        Txt(Engine).TEXT = RsInsp!Engine
        Call History_Enb(True)
    Else
        Call History_Field(False)
    End If
End If
Txt(Ins_Sheet).SetFocus
DGInsp.Visible = False
End Sub

Private Sub DgInsuranceCompany_Click()
If RsInsuranceCompany.RecordCount > 0 Then
    Txt(InsuranceCompany).Tag = RsInsuranceCompany!Code
    Txt(InsuranceCompany).TEXT = RsInsuranceCompany!Name
End If
Txt(InsuranceCompany).SetFocus
DgInsuranceCompany.Visible = False
End Sub

Private Sub DGMech_Click()
If DGMech.Columns(0).CAPTION = "Mechanic Name" Then
    If RsMech.RecordCount > 0 Then
        Txt(MyIndex).TEXT = RsMech!Name
        Txt(MyIndex).Tag = RsMech!Code
    End If
ElseIf DGMech.Columns(0).CAPTION = "Workshop Staff" Then
    If RsSuper.RecordCount > 0 Then
        Txt(MyIndex).TEXT = RsSuper!Name
        Txt(MyIndex).Tag = RsSuper!Code
    End If
End If
Txt(MyIndex).SetFocus
DGMech.Visible = False
End Sub

Private Sub DGModel_Click()
If RsModel.RecordCount > 0 Then
    Txt(Model).TEXT = RsModel!Code
    Txt(Model).Tag = RsModel!Code
End If
Txt(Model).SetFocus
DgModel.Visible = False
End Sub
Private Sub DGService_Click()
If RsServ.RecordCount > 0 Then
    Txt(SrvType).Tag = RsServ!Code
    Txt(SrvType).TEXT = RsServ!Name
    GSQL = "select SR.Lab_Amt from Service_Rates SR " & _
        " where Sr.Serv_Type='" & Txt(SrvType).Tag & "' and SR.Model='" & Txt(Model) & _
        "' order by sold_date desc"
    If GCn.Execute(GSQL).RecordCount > 0 Then
        Txt(SrvRate) = Format(GCn.Execute(GSQL).Fields(0).Value, "0.00")
        If RsServ!serv_catg = "F" Then
            Txt(CouponVal) = Txt(SrvRate)    'Coupon
        Else
            Txt(CouponVal) = ""
            Txt(Coupon) = ""
        End If
    Else
        Txt(SrvRate) = ""
        Txt(CouponVal) = ""
        Txt(Coupon) = ""
    End If
End If
Txt(SrvType).SetFocus
DGService.Visible = False
End Sub

Private Sub DGCity_Click()
If RsCity.RecordCount > 0 Then
    Txt(City).Tag = RsCity!Code
    Txt(City).TEXT = RsCity!Name
End If
Txt(City).SetFocus
DGCity.Visible = False
End Sub
Private Sub DGDealer_Click()
If RsDealer.RecordCount > 0 Then
    Txt(DCode).TEXT = RsDealer!Code
    Txt(DNAME).Tag = RsDealer!Code
    Txt(DNAME).TEXT = RsDealer!Name
End If
Txt(MyIndex).SetFocus
DGDealer.Visible = False
End Sub

Private Sub DGTrouble_Click()
If RsTrb.RecordCount > 0 Then
    txtgrid1(1).Tag = RsTrb!Code
    txtgrid1(1).TEXT = RsTrb!Name
End If
txtgrid1(1).SetFocus
DGTrouble.Visible = False
End Sub

Private Sub FGrid1_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
txtgrid1(1).Visible = False
End Sub

Private Sub FGrid1_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid1.Col
    Case Col_Trouble
        Call GridDblClick(Me, FGrid1, txtgrid1, 1)
End Select
TAddMode = False
End Sub

Private Sub FGrid1_EnterCell()
'FGrid1.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid1_GotFocus()
    FGrid1.BackColorSel = BackColorSelEnter
    FGrid1.ForeColorSel = ForeColorSelEnter
'    FGrid1.CellBackColor = CellBackColEnter
    'FGrid1.Col = Col_Trouble
    txtgrid1(1).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case Col_Trouble
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
    End Select
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid1.Col
        Case Col_Trouble
            Call GridDblClick(Me, FGrid1, txtgrid1, 1)
            TAddMode = False
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
Select Case FGrid1.Col
    Case Col_Trouble
        Call Get_Text(Me, FGrid1, txtgrid1, 1, False, KeyAscii)
    Case Col_Lab_Desc
'        Set RsLab = GCn.Execute("Select Lab_Code As Code, Lab_Desc As Name From Labour Where Lab_Code In (Select Lab_Code From Lab_Trouble Where CCCode='" & FGrid1.TextMatrix(FGrid1.Row, Col_Code) & "')")
'        Set DGLab.DataSource = RsLab
        Call Get_Text(Me, FGrid1, txtgrid1, 1, False, KeyAscii)
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid1_LostFocus()
FGrid1.BackColorSel = BackColorSelLeave
FGrid1.ForeColorSel = FGrid1.ForeColor
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

Private Sub FGrid2_Click()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
txtgrid2(1).Visible = False
End Sub

Private Sub FGrid2_DblClick()
On Error GoTo ELoop
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid2.Col = ElName Then Exit Sub
Select Case FGrid2.Col
    Case ElValue
        Call GridDblClick(Me, FGrid2, txtgrid2, 1)
End Select
TAddMode = False
ELoop:
    CheckError
End Sub

Private Sub FGrid2_EnterCell()
'FGrid2.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid2_GotFocus()
    FGrid2.BackColorSel = BackColorSelEnter
    FGrid2.ForeColorSel = ForeColorSelEnter
'    FGrid2.CellBackColor = CellBackColEnter
    FGrid2.Col = ElValue
    txtgrid2(1).Visible = False
End Sub

Private Sub FGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid2.Tag) = (FGrid2.Rows - (FGrid2.Rows - 1)) Then
'    FGrid2.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid2.Tag) = FGrid2.Rows - 1 Then
'    FGrid2.CellBackColor = CellBackColLeave
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid2.Tag = FGrid2.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid2.Col
        Case ElValue
            FGrid2.TextMatrix(FGrid2.Row, ElValue) = FGrid2.TextMatrix(FGrid2.Row, ElDefault)
    End Select
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid2.Col
        Case ElValue
            Call GridDblClick(Me, FGrid2, txtgrid2, 1)
            TAddMode = False
    End Select
End If
KeyCode = 0
ELoop:
    CheckError
End Sub

Private Sub FGrid2_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
    Select Case FGrid2.Col
        Case ElValue
           Call Get_Text(Me, FGrid2, txtgrid2, 1, False, KeyAscii)
        Case ElName
            FGrid2_LeaveCell
            FGrid2.Col = FGrid2.Col + 1
            FGrid2_EnterCell
            FGrid2.SetFocus
    End Select
    If KeyAscii <> vbKeyReturn Then TAddMode = True
ELoop:
    CheckError
End Sub

Private Sub FGrid2_LeaveCell()
'    FGrid2.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid2_LostFocus()
FGrid2.BackColorSel = BackColorSelLeave
FGrid2.ForeColorSel = FGrid2.ForeColor
FGrid2_Validate (True)
End Sub

Private Sub FGrid2_Scroll()
    txtgrid2(1).Visible = False
End Sub
Private Sub FGrid2_Validate(Cancel As Boolean)
'    FGrid2.CellBackColor = CellBackColLeave
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
Dim SrNo As Integer
    VType = "W_JC"
    ListArray = Array("Regular", "On Site Repair", "Quick Repair")
    Set mListItem = ListView_Items(ListView, Txt, JobType, ListArray, 3)
    
    If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
        Label3(49).Visible = True
        Txt(JobType).Visible = True
    End If
    WinSetting Me:    Ini_Grid
    TopCtrl1.Tag = PubUParam
    ForSiteCode = PubSiteCode
    Txt(JobDt).Tag = PubLoginDate
    Call BlankText
    '**Speed
    Me.Show
    DoEvents
    '**
    
     Dim sitecond As String
     sitecond = " And  Job_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("j.Docid", "3", "1") & "='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "select J.Job_No as SearchCode,J.DocID from Job_Card J where left(J.DocId,1)='" & PubDivCode & "' " & sitecond & " order by J.Job_Date desc,j.job_no desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 J.Job_No as SearchCode,J.DocID from Job_Card J where left(J.DocId,1)='" & PubDivCode & "' " & sitecond & "  order by J.Job_Date desc,j.job_no desc", GCn, adOpenDynamic, adLockOptimistic
    End If
    
    
    Set RsLab = GCn.Execute("Select Lab_Code As Code, Lab_Desc As Name, Lab_Rate, Time_Req, Lab_Rate*Time_Req As Amount  From Labour Where Lab_Code In (Select Lab_Code From Lab_Trouble Where CCCode='" & FGrid1.TextMatrix(FGrid1.Row, Col_Code) & "') Order By Lab_Desc")
    Set DGLab.DataSource = RsLab
    
    Set RSBook = New ADODB.Recordset
    With RSBook
        .CursorLocation = adUseClient
        .Open "SELECT " & cCStr("Book_no") & " AS CODE," & cCStr("Book_no") & " AS Name,Book_No, Book_Date,regno,chassis, name as OName ,Div_Code + " & cCStr("Book_No") & " + Site_Code as JobBookingID " & _
            " from Job_Booking " & _
            " where Div_Code='" & PubDivCode & "' and Job_DocID is null or Job_DocID=''" & _
            " order by Book_No", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGBook.DataSource = RSBook

    Set RsModel = New ADODB.Recordset
    RsModel.CursorLocation = adUseClient
    'RsModel.Open "Select MODEL as code,model as name ,model_desc as Listname,Chas_Type FROM Model where div_code='" & PubDivCode & "' Order by MODEL", GCn, adOpenDynamic, adLockOptimistic
    RsModel.Open "Select MODEL as code,model_desc as name, Model,Chas_Type FROM Model where (div_code='" & PubDivCode & "' or Div_Code='') Order by Model", GCn, adOpenDynamic, adLockOptimistic
    Set DgModel.DataSource = RsModel
    'RsModel.Sort = "code"
    
    Set RsServ = New ADODB.Recordset
    RsServ.CursorLocation = adUseClient
    RsServ.Open "Select Serv_type as code,serv_desc as name, Serv_Catg FROM Service_Type Order by Serv_DESC", GCn, adOpenDynamic, adLockOptimistic

    Set DGService.DataSource = RsServ
    RsServ.Sort = "name"
    
    
    Set RsInsuranceCompany = GCn.Execute("Select Code, Name From Insurance Order By Name")
    Set DgInsuranceCompany.DataSource = RsInsuranceCompany
    
    Set RsInsp = New ADODB.Recordset
    RsInsp.CursorLocation = adUseClient
    RsInsp.Open "Select " & cCStr("Insp_No") & " as code,Chassis as name,Insp_No,Regno,Model,engine,DocId " & _
            " FROM Job_Inspection " & _
            " where left(DocId,1)='" & PubDivCode & "' and Job_DocId is null or Job_DocId=''" & _
            " Order by Insp_no", GCn, adOpenDynamic, adLockOptimistic
    Set DGInsp.DataSource = RsInsp
    RsInsp.Sort = "Code"
    If RsInsp.RecordCount > 0 Then
        Label3(19).Visible = True
        LblColon(32).Visible = True
        Txt(Ins_Sheet).Visible = True
    End If
    Set RsHist = New ADODB.Recordset
    RsHist.CursorLocation = adUseClient
    RsHist.Open "Select CardNo as Code,Chassis,RegNo,Model,Name,CardNo,Engine,PhoneOff " & _
            " FROM Hiscard " & _
            " Where HISCARD.Div_Code='" & PubDivCode & "' Order by Regno", GCn, adOpenDynamic, adLockOptimistic
    Set DGHist.DataSource = RsHist
    RsHist.Sort = "Code"
    
    Set RsDealer = New ADODB.Recordset
    RsDealer.CursorLocation = adUseClient
    RsDealer.Open "Select d_Code as code,D_Name as name FROM Amd_Dealer Order by D_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGDealer.DataSource = RsDealer
    RsDealer.Sort = "code"
    
    Set RsMech = New ADODB.Recordset
    RsMech.CursorLocation = adUseClient
    RsMech.Open "Select Emp_Code as code,Emp_Name as name FROM Emp_Mast where Div_Code='" & PubDivCode & "' And Designation  in (" & pubWrkDesigRest & ") Order by Emp_name", GCn, adOpenDynamic, adLockOptimistic
    RsMech.Sort = "Name"
    
    Set RsSuper = New ADODB.Recordset
    RsSuper.CursorLocation = adUseClient
    RsSuper.Open "Select Emp_Code as code,Emp_Name as name FROM Emp_Mast where Div_Code='" & PubDivCode & "' And Designation in ('" & pubWrkDesigSuper & "') Order by Emp_name", GCn, adOpenDynamic, adLockOptimistic
    RsSuper.Sort = "Name"
    
    Set RsCity = New ADODB.Recordset
    RsCity.CursorLocation = adUseClient
    RsCity.Open "Select CityCode as code,CityName as name FROM City Order by CityName", GCn, adOpenDynamic, adLockOptimistic
    Set DGCity.DataSource = RsCity
    RsCity.Sort = "Name"
    
    Set RsTrb = New ADODB.Recordset
    RsTrb.CursorLocation = adUseClient
    RsTrb.Open "Select Trouble_code as code,Trouble_Name as name FROM trouble Order by trouble_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGTrouble.DataSource = RsTrb
    RsTrb.Sort = "Name"
    
    Set RsElem = New ADODB.Recordset
    RsElem.CursorLocation = adUseClient
    RsElem.Open "select IC.insp_code as IC_Code, IC.Insp_Description as IC_Name, IC.Report_Index as IC_Index, IE.InspElem_Code as IE_Code,IE.InspElem_Description as IE_Name, IE.Insp_ValueType as IE_Type, IE.Default_Value as IE_Value,IE.Report_Index as IE_Index from Inspection_Element as IE left join Inspection_Catg as IC on IE.Inspection_Catg=IC.Insp_Code where IC.Print_on='J' order by IC.Report_Index,IE.Report_Index", GCn, adOpenStatic, adLockReadOnly
                                                                                                                                                                                                                        
    SrNo = 1
    FGrid2.Rows = 1
    If RsElem.RecordCount > 0 Then
        Do Until RsElem.EOF
                         ' s no                   IE Code                   Nature                   Defa Value             desc      User Value     Remarks
            FGrid2.AddItem SrNo & Chr(9) & RsElem!IE_Code & Chr(9) & RsElem!IE_Type & Chr(9) & RsElem!IE_Value & Chr(9) & RsElem!IE_Name
            RsElem.MoveNext
            SrNo = SrNo + 1
        Loop
    Else
        FGrid2.AddItem FGrid2.Rows
    End If
    
    FGrid2.FixedRows = 1
    Set RsElem = Nothing
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
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
    Set Master = Nothing
    Set RSBook = Nothing
    Set RsInsp = Nothing
    Set RsHist = Nothing
    Set RsModel = Nothing
    Set RsServ = Nothing
    Set RsDealer = Nothing
    Set RsMech = Nothing
    Set RsSuper = Nothing
    Set RsCity = Nothing
    Set RsTrb = Nothing
'    RsElem.Requery
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    mAddFlag = "A"
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        If MsgBox("OPen P.C.D JobCard ?", vbYesNo, App.Title) = vbYes Then
            VType = "W_JC"
        Else
            VType = "W_JCO"
        End If
    End If
   
    Txt(JobDt).TEXT = Txt(JobDt).Tag 'Format(Date, "dd/MMM/yyyy")
    JobDocID = GetDocID(GCnFaW, VType, Txt(JobDt).TEXT, VoucherEditFlag, Txt(JobNo), lblPrefix, ForSiteCode)
    lblDocId.CAPTION = "DocId : " & JobDocID
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    
    Txt(ArrTime) = Format(time, "hh:mm")
    Txt(DelTime) = Format("17:00", "hh:mm")
    Txt(RecpTime) = Format(time, "hh:mm")
    If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
        Txt(JobType).SetFocus
    ElseIf Txt(JobNo).Enabled = True Then
        Txt(JobNo).SetFocus
    Else
        Txt(JobDt).SetFocus
    End If
    FGrid1.Col = Col_Trouble
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant, mTrans As Boolean
    If RSOJPR = True Then
        If MsgBox("Are You Sure To Cancel This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            GCn.BeginTrans
            GCn.Execute ("Update Job_Card  Set Docid_InvSpr='-------------Cancelld',Docid_InvLab='------------Cancelld' where docid='" & JobDocID & "'")
            'GCn.Execute ("Insert into Deletelog Values('" & JobDocID & "',1,0,'" & pubUName & "'," & ConvertDate(date$) & ",'" & Time$ & "')")
            GCn.CommitTrans
            Master.Requery
            MoveRec
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
        vBook = Master.AbsolutePosition
        mTrans = True
        'If Txt(BookSrl).Text <> "" Then
            GCn.Execute ("update job_booking set job_docid='' where Job_DocId='" & JobDocID & "'")
        'End If
        'If Txt(Ins_Sheet).Text <> "" Then
            GCn.Execute ("update job_inspection set job_docid='' where Job_DocId='" & JobDocID & "'")
        'End If
        GCn.Execute ("delete from job_demand where job_docid='" & JobDocID & "'")
        GCn.Execute ("delete from Job_Inspection2 where docid='" & JobDocID & "'")
        GCn.Execute "Delete from Job_Card  where Docid='" & JobDocID & "'"
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
    MsgBox err.Description, vbCritical, " Deletion Message"
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
'    If Not IsNull(Master!JOBCLOSEDATE) Then
'        MsgBox "JobCard is Closed, Can't Edit it", vbInformation, "Validation"
'        Exit Sub
'    End If
    If IsEditable(RetDate(Txt(JobDt))) = False Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    mAddFlag = "E"
    Call History_Enb(False)
    Txt(RecpTime).Enabled = False
    Txt(JobDt).SetFocus
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
    ''" & JobDocID & "'
    Dim sitecond As String
    sitecond = " And  Job_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("jc.Docid", "3", "1") & "='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    GSQL = "SELECT JC.Job_NO as searchcode,Jc.Job_no," & cTrim(cMID("JC.DocID", "9", "5")) & " as Prefix,JC.Job_date,JC.JobCloseDate, Service_Type.SERV_Desc,Coupon,hiscard.regno,hiscard.chassis,hiscard.name  FROM (JOB_card as jc left join  Service_Type on jc.serv_type=Service_Type.serv_type) left join HISCARD ON jc.cardno=hiscard.cardno WHERE LEFT(JC.DOCID,1)='" & PubDivCode & "' " & sitecond & " order by jc.JOB_NO"
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Public Sub SEARCHBACK(ByVal MyValue$)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("searchcode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("select J.Job_No as SearchCode,J.DocID from Job_Card J where left(J.DocId,1)='" & PubDivCode & "' And  J.DocId  = '" & MyValue & "' order by J.Job_Date desc,j.job_no desc")
    End If
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

Private Sub TopCtrl1_ePrn()
FrmPrn.top = 2220
FrmPrn.left = (Me.width - FrmPrn.width) / 2
FrmPrn.Visible = True
FrmPrn.ZOrder 0
OptPlain.Value = True
LblPrinter.CAPTION = Printer.DeviceName
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus
End Sub

Private Sub TopCtrl1_eRef()
    Call UpdRequery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim SrNo As Integer
    Dim DocIdHlp$, DupCheck$, mBookDivCode$, mBookSiteCode$
    Dim DupCount As Integer, mDelDtTm$, mArrivalDtTm$, LastHrs As Double
    Dim TmpRst As ADODB.Recordset, rsHist1 As ADODB.Recordset
    Dim hrs As Double
'   On Error GoTo errlbl
    If IsEditable(RetDate(Txt(JobDt))) = False Then Exit Sub
    Grid_Hide
    
   ' If TopCtrl1.TopText2 = "Browse" Then Exit Sub

    If IsValid(Txt(JobDt), "JobCard Date") = False Then Exit Sub
    If TopCtrl1.TopText2 = "Add" And VoucherEditFlag = True Then
        If IsValid(Txt(JobNo), "JobCard No") = False Then Exit Sub
    End If
    If Txt(VehRegNo) = "" And Txt(Chassis) = "" Then
        MsgBox "Registration No. or Chassis No. should have some data"
        Txt(VehRegNo).SetFocus
        Exit Sub
    End If
    If IsValid(Txt(Chassis), "Chassis") = False Then Exit Sub
    If IsValid(Txt(Model), "Model") = False Then Exit Sub
    If IsValid(Txt(OwnerName), "Owner Name") = False Then Exit Sub
    If IsValid(Txt(SrvType), "Service Type") = False Then Exit Sub
    
    If StrCmp(left(PubComp_Name, 3), "JMK") Then
        If StrCmp(Txt(InsuranceCompany), "") Then
            If MsgBox("Insurance Company Not Defined. Do You Want to Continue?", vbYesNo) = vbNo Then
                If IsValid(Txt(InsuranceCompany), "Insurance Company") = False Then Exit Sub
                If IsValid(Txt(InsuranceExpiry), "Insurance Expiry Date") = False Then Exit Sub
            End If
        Else
            If IsValid(Txt(InsuranceExpiry), "Insurance Expiry Date") = False Then Exit Sub
        End If
    End If
    
    
    If Txt(KmsHrs) = "Hrs." Then
        Set TmpRst = GCn.Execute("Select AtKmsHrs from Job_Card where KmsHrs='H' and CardNo='" & Txt(HistNo).TEXT & "' order by Job_Date DESC")
        If TmpRst.RecordCount > 0 Then
            TmpRst.MoveFirst
            LastHrs = VNull(TmpRst!AtKMsHrs)
            If LastHrs > 0 And Val(Txt(CurrentKMS)) < LastHrs Then
                MsgBox "Current Kms/Hrs less than Last KMs/Hrs", vbOKOnly, "Validation": Txt(CurrentKMS).SetFocus: Exit Sub
            End If
        End If
    Else
        If Val(Txt(LastKMS)) > 0 And Val(Txt(CurrentKMS)) < Val(Txt(LastKMS)) Then
            MsgBox "Current Kms/Hrs less than Last KMs/Hrs", vbOKOnly, "Validation": Txt(CurrentKMS).SetFocus: Exit Sub
        End If
    End If
    
    
    If Val(Txt(HrMeter)) > 0 Then
        Set TmpRst = GCn.Execute("Select HrMeter from Job_Card where CardNo='" & Txt(HistNo).TEXT & "' order by Job_Date DESC")
        If TmpRst.RecordCount > 0 Then
            TmpRst.MoveFirst
            LastHrs = VNull(TmpRst!HrMeter)
            If LastHrs > 0 And Val(Txt(HrMeter)) < LastHrs Then
                MsgBox "Current Hrs less than Last Hrs.", vbOKOnly, "Validation": Txt(HrMeter).SetFocus: Exit Sub
            End If
        End If
    End If
    
    If IsValid(Txt(FUEL), "Fuel in Tank") = False Then Exit Sub
    If IsValid(Txt(ArrDate), "Arrival Date") = False Then Exit Sub
    If RSOJPR = True Then
        If IsValid(Txt(EstSpr), "Estd Spares Amt.") = False Then Exit Sub
        If IsValid(Txt(EstLab), "Estd Labour Amt.") = False Then Exit Sub
    End If
'    If Val(txt(ArrTime)) <= 0 Then
'        MsgBox "Invalid Arrival Time", vbOKOnly, "Validation": txt(ArrTime).SetFocus: Exit Sub
'    End If
    If Txt(ArrDate) <> "" Then
        If CDate(Txt(ArrDate)) > CDate(Txt(JobDt)) Then
            MsgBox "Arrival Date is Grater than Job Date", vbOKOnly, "Validation"
            Txt(ArrDate).SetFocus: Exit Sub
        End If
        If Val(Txt(ArrTime)) <= 0 Then
            MsgBox "Invalid Arrival Time", vbOKOnly, "Validation": Txt(ArrTime).SetFocus: Exit Sub
        End If
        Txt(ArrTime) = Format(Txt(ArrTime), "hh:mm")
        If PubBackEnd = "A" Then
            mArrivalDtTm = "#" & Format(Txt(ArrDate) & " " & Txt(ArrTime), "dd/MMM/yyyy hh:mm") & "#"
        ElseIf PubBackEnd = "S" Then
            mArrivalDtTm = "'" & Format(Txt(ArrDate) & " " & Txt(ArrTime), "dd/MMM/yyyy hh:mm") & "'"
        End If
    Else
        Txt(ArrTime) = ""
        mArrivalDtTm = "Null"
    End If
    
    If IsValid(Txt(DelDt), "Expected Delivery Date") = False Then Exit Sub
    If Txt(DelDt) <> "" Then
        If CDate(Txt(DelDt)) < CDate(Txt(JobDt)) Then
            MsgBox "Expected Delivery Date is less than Job Date", vbOKOnly, "Validation"
            Txt(DelDt).SetFocus: Exit Sub
        End If
        If Val(Txt(DelTime)) <= 0 Then
            MsgBox "Invalid Delivery Time", vbOKOnly, "Validation": Txt(DelTime).SetFocus: Exit Sub
        End If
        Txt(DelTime) = Format(Txt(DelTime), "hh:mm")
        If PubBackEnd = "A" Then
            mDelDtTm = "#" & Format(Txt(DelDt) & " " & Txt(DelTime), "dd/MMM/yyyy hh:mm") & "#"
        ElseIf PubBackEnd = "S" Then
            mDelDtTm = "'" & Format(Txt(DelDt) & " " & Txt(DelTime), "dd/MMM/yyyy hh:mm") & "'"
        End If
    Else
        Txt(DelTime) = ""
        mDelDtTm = "Null"
    End If
    
    'hrs = CalcHrs(CDate(Txt(ArrDate)), CDate(Txt(DelDt)), Txt(ArrTime), Txt(DelTime), 0)
    
    If CDate(Txt(ArrDate) & " " & Txt(ArrTime)) > CDate(Txt(DelDt) & " " & Txt(DelTime)) Then
        MsgBox "Delivery Time Must Be > Arrival Time ", vbOKOnly, "Validation": Txt(DelTime).SetFocus: Exit Sub
    End If
    RsServ.MoveFirst
    RsServ.FIND ("Name = '" & Txt(SrvType) & "'")
    If RsServ!serv_catg = "F" Then
        IsValid Txt(Coupon), "Coupon No."
        If IsValid(Txt(CouponVal), "Coupon Value") = False Then Exit Sub
    End If
    'If IsValid(txt(Supervisor), "Supervisor Name") = False Then Exit Sub
    'If IsValid(txt(Mechanic), "Mechanic Name") = False Then Exit Sub
    '-------- Already done Service checking for PDI/Free Service
    If Txt(HistNo).Tag <> "" Then
        Dim SrvCat$, mFreeServCode$, RsTemp As ADODB.Recordset
        SrvCat = GCn.Execute("Select Serv_Catg from Service_Type where Serv_Type='" & Txt(SrvType).Tag & "'").Fields(0).Value
        If (SrvCat = "P" Or SrvCat = "F") Then
            Set RsTemp = New ADODB.Recordset
            RsTemp.CursorLocation = adUseClient
            RsTemp.Open "Select Serv_Catg,FreeServCode,Serv_Type from Service_Type where Serv_Catg in ('P','F') and Serv_Type in (select distinct Serv_Type from Job_Card where CardNo='" & Txt(HistNo).Tag & "' ) Order By Serv_Catg,FreeServCode", GCn, adOpenStatic, adLockReadOnly
            mFreeServCode = GCn.Execute("Select FreeServCode from Service_Type where Serv_Type='" & Txt(SrvType).Tag & "'").Fields(0).Value
            
            If RsTemp.RecordCount > 0 Then
                Do While RsTemp.EOF = False
                    If RsTemp!serv_catg = SrvCat And TopCtrl1.TopText2 = "Add" Then
                        If (Val(RsTemp!FREESERVCODE) >= Val(mFreeServCode)) Then
                            MsgBox "Selected Service Already Done!", vbOKOnly, "Service Checking"
                            Set RsTemp = Nothing
                            Txt(SrvType).SetFocus: Exit Sub
                        End If
                    End If
                    RsTemp.MoveNext
                Loop
                Set RsTemp = Nothing
            End If
        End If
    End If
    '--------
    If Txt(HistNo).Tag = "" Then
        If Txt(VehRegNo) <> "" Then
            If GCn.Execute("select " & xIsNull("Chassis", "") & " from hiscard where regno='" & Txt(VehRegNo) & "'").RecordCount > 0 Then
                DupCheck = GCn.Execute("select " & xIsNull("Chassis", "") & " from hiscard where regno='" & Txt(VehRegNo) & "'").Fields(0)
                If Txt(Chassis) <> DupCheck And DupCheck <> "" Then
                    MsgBox " Registration No. is already allocated with Chassis No. " & DupCheck, vbInformation, "Validation"
                    Txt(VehRegNo).SetFocus
                    Exit Sub
                End If
            End If
        End If
        If Txt(Chassis) <> "" Then
            If GCn.Execute("select " & xIsNull("regno", "") & " from hiscard where Chassis='" & Txt(Chassis) & "'").RecordCount > 0 Then
                DupCheck = GCn.Execute("select " & xIsNull("regno", "") & " from hiscard where Chassis='" & Txt(Chassis) & "'").Fields(0)
                If Txt(VehRegNo) <> DupCheck And DupCheck <> "" Then
                    MsgBox " Chassis No. is already allocated with Registration No. " & DupCheck, vbInformation, "Validation"
                    Txt(Chassis).SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    If Txt(BookSrl) <> "" Then
        mBookDivCode = PubDivCode
        mBookSiteCode = Right(Txt(BookSrl), 1)
    Else
        mBookDivCode = ""
        mBookSiteCode = ""
    End If
   'If mTrans = True Then Exit Sub
    GCn.BeginTrans
    mTrans = True
    
    Select Case mAddFlag
        Case "A"
            
            Set rsHist1 = Nothing
            If GCn.Execute("select count(*) from Job_Card where DocID='" & JobDocID & "' And Job_No= " & Val(Txt(JobNo)) & " ").Fields(0) > 0 Then
                If VoucherEditFlag Then
                    MsgBox "JobCard No. " & Txt(JobNo) & " Already Exists", vbCritical, "Validation Error"
                    Txt(JobNo).Tag = ""
                    Txt(JobNo).SetFocus
                    GoTo errlbl
                Else
                    Set rsHist1 = GCn.Execute("Select DocID,Job_Date,JobCloseDate FROM Job_Card " & _
                                              " Where Job_Card.CardNo='" & RsHist!Code & "'")
                    If rsHist1.RecordCount > 0 Then
                         For I = 1 To rsHist1.RecordCount
                             If IsNull(rsHist1!JobCloseDate) = True Or rsHist1!JobCloseDate = "" Then
                                    MsgBox "Job No.: " & PrinID(rsHist1!DocID) & vbCrLf & " Dt.: " & rsHist1!Job_Date & vbCrLf & "Already exists for selected Vehicle. Entry Aborted... ", vbCritical + vbOKOnly, "Job Already Exists"
                                    GCn.RollbackTrans
                                    Exit Sub
                             End If
                             rsHist1.MoveNext
                         Next
                    End If
                    Set rsHist1 = Nothing
                
                    Txt(JobNo).Tag = Txt(JobNo)
                    JobDocID = GetDocID(GCnFaW, VType, Txt(JobDt), VoucherEditFlag, Txt(JobNo), lblPrefix, ForSiteCode)
                    If Val(Txt(JobNo)) <= Val(Txt(JobNo).Tag) Then
                        MsgBox "Job No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                        GoTo errlbl
                    End If
                End If
            End If
            
            If Txt(HistNo) = "" Then   'Insert New Vehicle
                Txt(HistNo) = PubSiteCode + Right("000000" & GCn.Execute("select max(" & cVal(cMID("cardno", "3", "Len(CardNo) - 2")) & ")+1 from hiscard").Fields(0), 6)
                Txt(HistNo).Tag = Txt(HistNo)
                GCn.Execute "insert into hiscard(cardno,Site_Code,Div_Code,carddate,model,regno,chassis,engine,vehserialno,govt_yn,dealer_code,delivery_date,couponno,name,add1,add2,add3,citycode,phoneoff,phoneresi,mobile,SUPPLIER_BILLNO,U_Name, U_EntDt, U_AE, InsuranceCompany, InsuranceExpiry, InsurancePolicyNo) " & _
                    " values('" & Txt(HistNo) & "','" & PubSiteCode & "','" & PubDivCode & "'," & ConvertDate(Txt(JobDt)) & ",'" & Txt(Model).Tag & "','" & Txt(VehRegNo) & "','" & Txt(Chassis) & "','" & Txt(Engine) & _
                    "','" & Txt(VehSrlNo) & "'," & IIf(Txt(GovtYn) = "Yes", 1, 0) & ",'" & Txt(DCode) & "'," & ConvertDate(Txt(SaleDate)) & ",'" & Txt(Coupon) & "','" & Txt(OwnerName) & _
                    "','" & Txt(Address1) & "','" & Txt(Address2) & "','" & Txt(Address3) & "'," & Val(Txt(City).Tag) & ",'" & Txt(PhoneOff) & "','" & Txt(PhoneResi) & "','" & Txt(Mobile) & "','" & Txt(InvNo) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A', '" & Txt(InsuranceCompany).Tag & "', " & ConvertDate(Txt(InsuranceExpiry)) & ", '" & Txt(InsurancePolicyNo) & "')"
            End If
            
            If RSOJPR = True Then
                GSQL = "insert into Job_Card(DocId, Site_Code, Job_No, Job_Date,Job_BookDivCode,Job_BookSiteCode,Job_BookNo,Job_InspDocId, cardno,Govt_yn,Serv_type,Serv_Rate,Coupon,Coupon_Value,atkmsHrs,Fuel,Est_spCost,Est_LabCost,ArrivalTime,Recp_Time,ExpDelDate,Body_Damage,OpenRemarks,RecBy_Mechanic,RecBy_Supervisor, CreatedU_Name, CreatedU_EntDt, CreatedU_AE, U_Name, U_EntDt, U_AE, Created_AddBy, Created_AddDate,KmsHrs,HrMeter,JobType) " & _
                    " values('" & JobDocID & "','" & PubSiteCode & "'," & Txt(JobNo) & "," & ConvertDate(Txt(JobDt)) & ",'" & PubDivCode & "','" & mBookSiteCode & "'," & Val(Txt(BookSrl)) & ",'" & Txt(Ins_Sheet).Tag & "','" & Txt(HistNo).Tag & "'," & IIf(Txt(GovtYn) = "Yes", 1, 0) & ",'" & Txt(SrvType).Tag & "'," & Val(Txt(SrvRate)) & ",'" & Txt(Coupon) & "'," & Val(Txt(CouponVal)) & "," & _
                    " " & Val(Txt(CurrentKMS)) & ",'" & Txt(FUEL) & "'," & Val(Txt(EstSpr)) & "," & Val(Txt(EstLab)) & "," & mArrivalDtTm & "," & cTime(Format(Txt(RecpTime), "hh:mm")) & "," & mDelDtTm & ",'" & (Txt(Damage)) & _
                    "','" & Txt(Remarks) & "','" & Txt(Mechanic).Tag & "','" & Txt(Supervisor).Tag & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & pubUName & "'," & ConvertDateTime(PubServerDate) & "," & cIIF("'" & Txt(KmsHrs) & "'='Hrs.'", "'H'", "'K'") & ",'" & Txt(HrMeter) & "','" & left(LblJobType.CAPTION, 1) & "')"
            Else
                GSQL = "insert into Job_Card(DocId, Site_Code, Job_No, Job_Date,Job_BookDivCode,Job_BookSiteCode,Job_BookNo,Job_InspDocId, cardno,Govt_yn,Serv_type,Serv_Rate,Coupon,Coupon_Value,atkmsHrs,Fuel,Est_spCost,Est_LabCost,ArrivalTime,Recp_Time,ExpDelDate,Body_Damage,OpenRemarks,RecBy_Mechanic,RecBy_Supervisor, CreatedU_Name, CreatedU_EntDt, CreatedU_AE, U_Name, U_EntDt, U_AE, Created_AddBy, Created_AddDate,KmsHrs,HrMeter) " & _
                    " values('" & JobDocID & "','" & PubSiteCode & "'," & Txt(JobNo) & "," & ConvertDate(Txt(JobDt)) & ",'" & PubDivCode & "','" & mBookSiteCode & "'," & Val(Txt(BookSrl)) & ",'" & Txt(Ins_Sheet).Tag & "','" & Txt(HistNo).Tag & "'," & IIf(Txt(GovtYn) = "Yes", 1, 0) & ",'" & Txt(SrvType).Tag & "'," & Val(Txt(SrvRate)) & ",'" & Txt(Coupon) & "'," & Val(Txt(CouponVal)) & "," & _
                    " " & Val(Txt(CurrentKMS)) & ",'" & Txt(FUEL) & "'," & Val(Txt(EstSpr)) & "," & Val(Txt(EstLab)) & "," & mArrivalDtTm & "," & cTime(Format(Txt(RecpTime), "hh:mm")) & "," & mDelDtTm & ",'" & (Txt(Damage)) & _
                    "','" & Txt(Remarks) & "','" & Txt(Mechanic).Tag & "','" & Txt(Supervisor).Tag & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & pubUName & "'," & ConvertDateTime(PubServerDate) & "," & cIIF("'" & Txt(KmsHrs) & "'='Hrs.'", "'H'", "'K'") & ",'" & Txt(HrMeter) & "')"
            End If
            GCn.Execute GSQL
           
            UpdVouSrlNo GCnFaS, JobDocID, Txt(JobDt)

        Case "E"
            GCn.Execute ("delete from job_demand where job_docid='" & JobDocID & "'")
            GCn.Execute ("delete from Job_Inspection2 where docid='" & JobDocID & "'")
            
            If OldBookingID <> "" Then
                GCn.Execute ("update job_booking set job_docid='' where div_Code+ " & cCStr("book_no") & " +Site_Code='" & OldBookingID & "'")
            End If
            If OldInsSheetID <> "" Then
                GCn.Execute ("update job_inspection set job_docid='' where DocId='" & OldInsSheetID & "'")
            End If
            
            
            '' note: Recp_time is not updated in edit mode because it is valid only for job opening
            If RSOJPR = True Then
                GCn.Execute "Update Job_Card Set Job_Date= " & ConvertDate(Txt(JobDt)) & _
                    ", Job_BookDivCode='" & mBookDivCode & "',Job_BookSiteCode='" & mBookSiteCode & "',Job_BookNo=" & Val(Txt(BookSrl)) & _
                    ", Job_InspDocId='" & Txt(Ins_Sheet).Tag & "', cardno='" & Txt(HistNo).Tag & "',Govt_yn=" & IIf(Txt(GovtYn) = "Yes", 1, 0) & _
                    ", Serv_type='" & Txt(SrvType).Tag & "',Serv_Rate=" & Val(Txt(SrvRate)) & ",Coupon='" & Txt(Coupon) & _
                    "',Coupon_Value=" & Val(Txt(CouponVal)) & ",atkmsHrs=" & Val(Txt(CurrentKMS)) & _
                    ",Fuel='" & Txt(FUEL) & "',Est_spCost=" & Val(Txt(EstSpr)) & _
                    ",Est_LabCost=" & Val(Txt(EstLab)) & ",ArrivalTime=" & mArrivalDtTm & ",ExpDelDate=" & mDelDtTm & ",Body_Damage='" & (Txt(Damage)) & "',OpenRemarks='" & Txt(Remarks) & _
                    "',RecBy_Mechanic='" & Txt(Mechanic).Tag & "',RecBy_Supervisor='" & Txt(Supervisor).Tag & _
                    "',CreatedU_Name='" & pubUName & "', CreatedU_EntDt=" & ConvertDate(PubServerDate) & _
                    ",CreatedU_AE='E',U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E', Created_ModifyBy = '" & pubUName & "', Created_ModifyDate=" & ConvertDateTime(PubServerDate) & ",KmsHrs=iif('" & Txt(KmsHrs) & "'='Hrs.','H','K'),HrMeter='" & Txt(HrMeter) & "',JobType='" & left(LblJobType.CAPTION, 1) & "' where Docid='" & JobDocID & "'"
            Else
                GCn.Execute "Update Job_Card Set Job_Date= " & ConvertDate(Txt(JobDt)) & _
                    ", Job_BookDivCode='" & mBookDivCode & "',Job_BookSiteCode='" & mBookSiteCode & "',Job_BookNo=" & Val(Txt(BookSrl)) & _
                    ", Job_InspDocId='" & Txt(Ins_Sheet).Tag & "', cardno='" & Txt(HistNo).Tag & "',Govt_yn=" & IIf(Txt(GovtYn) = "Yes", 1, 0) & _
                    ", Serv_type='" & Txt(SrvType).Tag & "',Serv_Rate=" & Val(Txt(SrvRate)) & ",Coupon='" & Txt(Coupon) & _
                    "',Coupon_Value=" & Val(Txt(CouponVal)) & ",atkmsHrs=" & Val(Txt(CurrentKMS)) & _
                    ",Fuel='" & Txt(FUEL) & "',Est_spCost=" & Val(Txt(EstSpr)) & _
                    ",Est_LabCost=" & Val(Txt(EstLab)) & ",ArrivalTime=" & mArrivalDtTm & ",ExpDelDate=" & mDelDtTm & ",Body_Damage='" & (Txt(Damage)) & "',OpenRemarks='" & Txt(Remarks) & _
                    "',RecBy_Mechanic='" & Txt(Mechanic).Tag & "',RecBy_Supervisor='" & Txt(Supervisor).Tag & _
                    "',CreatedU_Name='" & pubUName & "', CreatedU_EntDt=" & ConvertDate(PubServerDate) & _
                    ",CreatedU_AE='E',U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E',Created_ModifyBy = '" & pubUName & "', Created_ModifyDate=" & ConvertDateTime(PubServerDate) & ",KmsHrs=" & cIIF("'" & Txt(KmsHrs) & "'='Hrs.'", "'H'", "'K'") & ",HrMeter='" & Txt(HrMeter) & "' where Docid='" & JobDocID & "'"
            End If
    End Select
    'Lock Job Booking
    If Txt(BookSrl) <> "" Then
        GCn.Execute ("update job_booking set job_docid='" & JobDocID & "'where div_Code+ " & cCStr("book_no") & "+Site_Code='" & Txt(BookSrl).Tag & "'")
    End If
    'Lock Inspection Sheet
    If Txt(Ins_Sheet) <> "" Then
        GCn.Execute ("update job_inspection set job_docid='" & JobDocID & "'where DocId='" & Txt(Ins_Sheet).Tag & "'")
    End If
    '
    SrNo = 1
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, Col_Trouble) <> "" And (Not IsNull(FGrid1.TextMatrix(I, Col_Trouble))) Then
            GSQL = "Insert into Job_Demand(Job_DocId , Site_Code, S_No, Code, Details, Repeat_YN, Lab_Code, Lab_Rate, Time_Req, Amount, U_Name, U_EntDt, U_AE) " & _
                   "Values('" & JobDocID & "','" & PubSiteCode & "'," & SrNo & ",'" & IIf(FGrid1.TextMatrix(I, Col_Code) = "", " ", FGrid1.TextMatrix(I, Col_Code)) & "','" & FGrid1.TextMatrix(I, Col_Trouble) & "','" & Val(FGrid1.TextMatrix(I, Col_Repeat)) & "', '" & FGrid1.TextMatrix(I, Col_Lab_Code) & "', " & Val(FGrid1.TextMatrix(I, Col_Lab_Rate)) & ", " & Val(FGrid1.TextMatrix(I, Col_Time_Req)) & ", " & Val(FGrid1.TextMatrix(I, Col_Amount)) & ", '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & mAddFlag & "')"
            GCn.Execute GSQL
            SrNo = SrNo + 1
        End If
    Next I
    
    For I = 1 To FGrid2.Rows - 1
        If FGrid2.TextMatrix(I, ElCode) <> "" Then
            GSQL = "insert into Job_Inspection2(DocId , Site_Code, s_No, element_code, remarks, U_Name, U_EntDt, U_AE) " & _
                    " values('" & JobDocID & "','" & PubSiteCode & "'," & Val(FGrid2.TextMatrix(I, 0)) & ",'" & FGrid2.TextMatrix(I, ElCode) & "','" & FGrid2.TextMatrix(I, ElValue) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & mAddFlag & "')"
            GCn.Execute GSQL
        End If
    Next I
    
    If Txt(HistNo) <> "" Then
        GSQL = "Update HisCard Set InsuranceCompany = '" & Txt(InsuranceCompany).Tag & "', InsuranceExpiry=" & ConvertDate(Txt(InsuranceExpiry)) & ", InsurancePolicyNo = '" & Txt(InsurancePolicyNo) & "' Where CardNo = '" & Txt(HistNo) & "'"
        GCn.Execute GSQL
    End If
    
    
    GCn.CommitTrans
'   mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select J.Job_No as SearchCode,J.DocID from Job_Card J where left(J.DocId,1)='" & PubDivCode & "' And  J.DocId  = '" & JobDocID & "' order by J.Job_Date desc,j.job_no desc")
    End If
    Call UpdRequery
    If TopCtrl1.TopText2 = "Add" Then
        Txt(JobDt).Tag = Txt(JobDt)
    End If
    
    Master.FIND "Docid = '" & JobDocID & "'"
    TopCtrl1_ePrn
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    mTrans = False
    Exit Sub

errlbl:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus Txt(Index)
    txtgrid1(1).Visible = False
    txtgrid2(1).Visible = False
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case BookSrl
            If RSBook.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
            If RSBook.EOF = True Or RSBook.BOF = True Then Exit Sub
            If Txt(Index).TEXT <> RSBook!Code Then
                RSBook.MoveFirst
                RSBook.FIND "code ='" & Txt(Index).TEXT & "'"
            End If
        Case Ins_Sheet
            If RsInsp.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
            If RsInsp.EOF = True Or RsInsp.BOF = True Then Exit Sub
            If Txt(Index).TEXT <> RsInsp!Code Then
                RsInsp.MoveFirst
                RsInsp.FIND "code ='" & Txt(Index).TEXT & "'"
            End If
        Case VehRegNo
            DGridColSwap DGHist, 0
            RsHist.Sort = "REGNO"
            If RsHist.RecordCount = 0 Or Txt(Index) = "" Then Exit Sub
            If UCase(Txt(Index)) <> UCase(XNull(RsHist!RegNo)) Then
                RsHist.MoveFirst
                RsHist.FIND "RegNo ='" & Txt(Index) & "'"
            End If
        Case Chassis
            DGridColSwap DGHist, 1
            RsHist.Sort = "CHASSIS"
            If RsHist.RecordCount = 0 Or Txt(Index) = "" Then Exit Sub
            If UCase(Txt(Index)) <> UCase(RsHist!Chassis) Then
                RsHist.MoveFirst
                RsHist.FIND "Chassis ='" & Txt(Index) & "'"
            End If
        Case OwnerName
            DGridColSwap DGHist, 3
            RsHist.Sort = "name"
        Case City
            DGridColSwap DGCity, 1
            If RsCity.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
            If RsInsp.EOF = True Or RsInsp.BOF = True Then Exit Sub
            If Txt(Index).TEXT <> RsCity!Name Then
                RsCity.MoveFirst
                RsCity.FIND "name ='" & Txt(Index).TEXT & "'"
            End If
        Case SrvType
            DGridColSwap DGService, 1
            If RsServ.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
            If RsServ.EOF = True Or RsServ.BOF = True Then Exit Sub
            
            If Txt(Index).TEXT <> RsServ!Code Then
                RsServ.MoveFirst
                RsServ.FIND "name ='" & Txt(Index).TEXT & "'"
            End If
        Case Model
            If RsModel.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
            If RsModel.EOF = True Or RsModel.BOF = True Then Exit Sub
            If Txt(Index).TEXT <> RsModel!Code Then
                RsModel.MoveFirst
                RsModel.FIND "cODE ='" & Txt(Index).TEXT & "'"
            End If
        Case DNAME
            RsDealer.Sort = "name"
            DGridColSwap DGDealer, 1
            If RsDealer.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
            If RsDealer.EOF = True Or RsDealer.BOF = True Then Exit Sub
            If Txt(Index).TEXT <> RsDealer!Name Then
                RsDealer.MoveFirst
                RsDealer.FIND "name ='" & Txt(Index).TEXT & "'"
            End If
            
        Case InsuranceCompany
            DgInsuranceCompany.Move Txt(Index).left, Txt(Index).top + Txt(Index).height + 30
            DGridColSwap DgInsuranceCompany, 1
            If RsInsuranceCompany.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
            If RsInsuranceCompany.EOF = True Or RsInsuranceCompany.BOF = True Then Exit Sub
            If Txt(Index).TEXT <> RsInsuranceCompany!Name Then
                RsInsuranceCompany.MoveFirst
                RsInsuranceCompany.FIND "name ='" & Txt(Index).TEXT & "'"
            End If
            
        Case DCode
            RsDealer.Sort = "code"
            DGridColSwap DGDealer, 0
            If RsDealer.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
            If RsDealer.EOF = True Or RsDealer.BOF = True Then Exit Sub
            If Txt(Index).TEXT <> RsDealer!Code Then
                RsDealer.MoveFirst
                RsDealer.FIND "code ='" & Txt(Index).TEXT & "'"
            End If
            
        Case Mechanic
            DGMech.Columns(0).CAPTION = "Mechanic Name"
            Set DGMech.DataSource = RsMech
            DGridColSwap DGMech, 1
            RsMech.Sort = "name"
            If RsMech.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
            If RsMech.EOF = True Or RsMech.BOF = True Then Exit Sub
            If Txt(Index).TEXT <> RsMech!Name Then
                RsMech.MoveFirst
                RsMech.FIND "name ='" & Txt(Index).TEXT & "'"
            End If
            
        Case Supervisor
            DGMech.Columns(0).CAPTION = "Workshop Staff"
            Set DGMech.DataSource = RsSuper
            DGridColSwap DGMech, 1
            RsSuper.Sort = "name"
            If RsSuper.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
            If RsSuper.EOF = True Or RsSuper.BOF = True Then Exit Sub
            If Txt(Index).TEXT <> RsSuper!Name Then
                RsSuper.MoveFirst
                RsSuper.FIND "name ='" & Txt(Index).TEXT & "'"
            End If
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case JobType
            ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1200
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
        Case VehRegNo
            If TopCtrl1.TopText2 = "Add" Then
                DGridTxtKeyDown_Mast DGHist, Txt, Index, RsHist, KeyCode, False, 2
            Else
                DGridTxtKeyDown DGHist, Txt, Index, RsHist, KeyCode, False, 2
            End If
        Case Chassis
            If TopCtrl1.TopText2 = "Add" Then
                DGridTxtKeyDown_Mast DGHist, Txt, Index, RsHist, KeyCode, False, 1
            Else
                DGridTxtKeyDown DGHist, Txt, Index, RsHist, KeyCode, False, 1
            End If
        Case OwnerName
            DGridTxtKeyDown_Mast DGHist, Txt, Index, RsHist, KeyCode, False, 4
        Case BookSrl
            DGridTxtKeyDown DGBook, Txt, Index, RSBook, KeyCode, False, 0, frmJobBooking, "frmJobBooking"
        Case City
            DGridTxtKeyDown DGCity, Txt, Index, RsCity, KeyCode, False, 1, frmCity, "frmCity"
        Case Ins_Sheet
            DGridTxtKeyDown DGInsp, Txt, Index, RsInsp, KeyCode, False, 0
        Case Model
            DGridTxtKeyDown DgModel, Txt, Index, RsModel, KeyCode, False, 0, frmModel, "frmModel"
        Case SrvType
            DGridTxtKeyDown DGService, Txt, Index, RsServ, KeyCode, False, 1, frmService, "frmService"
        Case DCode
            DGridTxtKeyDown DGDealer, Txt, Index, RsDealer, KeyCode, False, 0, frmDealer, "frmDealer"
        Case DNAME
            DGridTxtKeyDown DGDealer, Txt, Index, RsDealer, KeyCode, False, 1, frmDealer, "frmDealer"
        Case InsuranceCompany
            DGridTxtKeyDown DgInsuranceCompany, Txt, Index, RsInsuranceCompany, KeyCode, False, 1, FrmInsurance, "frmInsurance"
        Case Supervisor
            DGridTxtKeyDown DGMech, Txt, Index, RsSuper, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
        Case Mechanic
            DGridTxtKeyDown DGMech, Txt, Index, RsMech, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
        Case CurrentKMS
            NumDown Txt(Index), KeyCode, 6, 0
        Case CouponVal, EstSpr, EstLab
            NumDown Txt(Index), KeyCode, 6, 2
    End Select
    If DGBook.Visible = False And DgInsuranceCompany.Visible = False And DGHist.Visible = False And DGMech.Visible = False And DgModel.Visible = False And DGService.Visible = False And DGDealer.Visible = False And DGInsp.Visible = False And DGCity.Visible = False And ListView.Visible = False Then
        '' KEY DOWN
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> Remarks Then
            Ctrl_DownKeyDown KeyCode, Shift
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = Remarks Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        ' KEY UP
        If TopCtrl1.TopText2 = "Add" Then
            If RSOJPR = True Then
                If (Txt(JobNo).Enabled = False And Index <> JobDt) Or (Txt(JobNo).Enabled = True And Index <> JobType) Then
                    If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
                End If
            Else
                If (Txt(JobNo).Enabled = False And Index <> JobDt) Or (Txt(JobNo).Enabled = True And Index <> JobNo Or Txt(JobType).Enabled = True And Index <> JobType) Then
                    If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
                End If
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then
            If Index <> JobDt Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
    Select Case Index
        Case VehRegNo, Chassis, Model
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select

    Select Case Index
        Case VehRegNo
            If TopCtrl1.TopText2 <> "Add" Then
                If DGHist.Visible = True Then DGridTxtKeyPress Txt, Index, RsHist, KeyAscii, "RegNo"
            End If
        Case Chassis
            If TopCtrl1.TopText2 <> "Add" Then
                If DGHist.Visible = True Then DGridTxtKeyPress Txt, Index, RsHist, KeyAscii, "Chassis"
            End If
        Case ArrTime, DelTime, JCTime
            Call NumPress(Txt(Index), KeyAscii, 2, 2)
        Case JobNo
            Call NumPress(Txt(Index), KeyAscii, 8, 0)
        Case BookSrl
            DGridTxtKeyPress Txt, Index, RSBook, KeyAscii, "code"
        Case CurrentKMS
            Call NumPress(Txt(Index), KeyAscii, 6, 0)
        Case CouponVal, EstSpr, EstLab
            Call NumPress(Txt(Index), KeyAscii, 6, 2)
        Case SrvType
            DGridTxtKeyPress Txt, Index, RsServ, KeyAscii, "name"
        Case Model
            DGridTxtKeyPress Txt, Index, RsModel, KeyAscii, "Code"
        Case Ins_Sheet
            DGridTxtKeyPress Txt, Index, RsInsp, KeyAscii, "Code"
        Case DCode
            DGridTxtKeyPress Txt, Index, RsDealer, KeyAscii, "code"
        Case DNAME
            DGridTxtKeyPress Txt, Index, RsDealer, KeyAscii, "name"
        Case InsuranceCompany
            DGridTxtKeyPress Txt, Index, RsInsuranceCompany, KeyAscii, "name"
            
        Case City
            DGridTxtKeyPress Txt, Index, RsCity, KeyAscii, "name"
        Case Supervisor
            DGridTxtKeyPress Txt, Index, RsSuper, KeyAscii, "name"
        Case Mechanic
            DGridTxtKeyPress Txt, Index, RsMech, KeyAscii, "name"
        Case KmsHrs
            If Asc("H") = KeyAscii Or Asc("h") = KeyAscii Then
                Txt(KmsHrs) = "Hrs."
            ElseIf Asc("K") = KeyAscii Or Asc("k") = KeyAscii Then
                Txt(KmsHrs) = "Kms."
            End If
    End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
'       Case ArrTime, DelTime, JCTime
'            Txt(Index) = Format(Txt(Index), "hh:mm") '"00:00"),Format(
        Case JobType
            ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
        Case VehRegNo
            If TopCtrl1.TopText2 = "Add" Then
                DGridTxtKeyUp_Mast Txt, Index, RsHist, KeyCode, "Regno"
            End If
        Case Chassis
            If TopCtrl1.TopText2 = "Add" Then
                DGridTxtKeyUp_Mast Txt, Index, RsHist, KeyCode, "Chassis"
            End If
        Case OwnerName
            If TopCtrl1.TopText2 = "Add" Then
                DGridTxtKeyUp_Mast Txt, Index, RsHist, KeyCode, "Name"
            End If
        Case GovtYn
            If Len(Txt(Index)) = 0 Or UCase(mID(Txt(Index), 1, 1)) = "N" Then
                Txt(Index) = "No"
            ElseIf UCase(mID(Txt(Index), 1, 1)) = "Y" Then
                Txt(Index) = "Yes"
            Else
                Txt(Index) = "No"
            End If
    End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
  Select Case Index
        Case ArrTime, DelTime, JCTime
            Txt(Index) = Format(Txt(Index), "hh:mm") '"00:00"),Format(
End Select
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim RsCard As ADODB.Recordset
Dim xChange As Boolean
Select Case Index
    Case JobType
        If RSOJPR = True Then
            Label4.Visible = True: LblJobType.Visible = True
            LblJobType = Txt(JobType).TEXT
        End If
    Case ArrTime, DelTime, JCTime
        Txt(Index) = Format(Txt(Index), "hh:mm") '"00:00"),Format(
    Case JobNo
        lblPrefix.CAPTION = XNull(lblPrefix.CAPTION)
        JobDocID = GetDocID(GCnFaW, VType, Txt(JobDt).TEXT, VoucherEditFlag, Txt(JobNo), lblPrefix, ForSiteCode)
        lblDocId = "DocId : " & JobDocID
        If VoucherEditFlag = True Then    ' Manual
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select Docid From Job_card Where DocID='" & JobDocID & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                If Txt(JobNo).Enabled = True Then
                    Txt(JobNo).SetFocus
                End If
            End If
        End If
    Case BookSrl
        If Txt(BookSrl).TEXT <> "" Then
            If RSBook.BOF = True Or RSBook.EOF = True Then Exit Sub
            Txt(BookSrl).TEXT = RSBook!Book_no
            Txt(BookSrl).Tag = RSBook!JobBookingID
            FillBookingData
        End If
    Case Ins_Sheet
        If Txt(Ins_Sheet).TEXT <> "" Then
            If RsInsp.BOF = True Or RsInsp.EOF = True Then Exit Sub
            Txt(Ins_Sheet).Tag = RsInsp!DocID
            If RsInsp!RegNo <> "" Then
                RsHist.FIND ("Regno='" & RsInsp!RegNo & "'")
            ElseIf RsInsp!Chassis <> "" Then
                RsHist.FIND ("chassis='" & RsInsp!Chassis & "'")
            End If
            If RsHist.EOF = True Or RsHist.BOF = True Then
                Txt(HistNo).Tag = ""
                Txt(HistNo).TEXT = ""
                Txt(VehRegNo).TEXT = RsInsp!RegNo
                Txt(Chassis).TEXT = RsInsp!Name
                Txt(Model).Tag = RsInsp!Model
                Txt(Model).TEXT = RsInsp!Model
                Txt(Engine).TEXT = RsInsp!Engine
                Call History_Enb(True)
            Else
                Call History_Field(False)
            End If
        End If
    Case OwnerName
        If RsHist.EOF = True Or RsHist.BOF = True Then Exit Sub
        If RsHist.RecordCount > 0 Then
            If Txt(OwnerName).TEXT <> "" And UCase(Trim(Txt(OwnerName).TEXT)) = UCase(Trim(RsHist!Name)) Then
                Set RsCard = New ADODB.Recordset
                RsCard.CursorLocation = adUseClient
                
                RsCard.Open "Select H.Name,H.Add1,H.Add2,H.Add3,H.PhoneResi,H.PhoneOff,H.Mobile,H.CityCode,C.CityName " & _
                        " FROM Hiscard as H left join City as C on H.CityCode=C.CityCode " & _
                        " Where H.Div_Code='" & PubDivCode & "'  and Name='" & Txt(OwnerName) & "' Order by Name", GCn, adOpenDynamic, adLockOptimistic
                
                Txt(Address1).TEXT = XNull(RsCard!Add1)
                Txt(Address2).TEXT = XNull(RsCard!Add2)
                Txt(Address3).TEXT = XNull(RsCard!Add3)
                Txt(PhoneResi).TEXT = XNull(RsCard!PhoneResi)
                Txt(PhoneOff).TEXT = XNull(RsCard!PhoneOff)
                Txt(Mobile).TEXT = XNull(RsCard!Mobile)
                Txt(City).Tag = XNull(RsCard!CityCode)
                Txt(City).TEXT = XNull(RsCard!CityName)
                Set RsCard = Nothing
            End If
        End If
    Case SrvType
        If RsServ.EOF = False And RsServ.BOF = False Then
            If Txt(SrvType).TEXT <> "" Then
                xChange = IIf(TopCtrl1.TopText2 = "Edit" And Txt(Index).Tag <> RsServ!Code, False, True)
                
                Txt(SrvType).TEXT = RsServ!Name
                Txt(SrvType).Tag = RsServ!Code
                
                
                If xChange = False Then
                    If RsServ!serv_catg = "F" Or RsServ!serv_catg = "P" Then
                        If Txt(HistNo) <> "" Then
                            Txt(Coupon) = GCn.Execute("Select H.CouponNo FROM Hiscard as H Where H.Div_Code='" & PubDivCode & "' and CardNo='" & Txt(HistNo) & "' ").Fields(0).Value
                        End If
                    Else
                        Txt(Coupon) = ""
                    End If
                
                    GSQL = "select SR.Lab_Amt from Service_Rates SR " & _
                        " where Sr.Serv_Type='" & Txt(SrvType).Tag & "' and SR.Model='" & Txt(Model) & _
                        "' order by sold_date desc"
                    If GCn.Execute(GSQL).RecordCount > 0 Then
                        Txt(SrvRate) = Format(GCn.Execute(GSQL).Fields(0).Value, "0.00")
                        If RsServ!serv_catg = "F" Then
                            Txt(CouponVal) = Txt(SrvRate)    'Coupon
                        Else
                            Txt(CouponVal) = ""
                        End If
                    Else
                        Txt(SrvRate) = ""
                        Txt(CouponVal) = ""
                    End If
                End If
            End If
        Else
            Txt(SrvType).TEXT = ""
            Txt(SrvType).Tag = ""
        End If
    Case Coupon
        If RsServ!serv_catg = "F" Then
            If Txt(Coupon) = "" Then
                MsgBox "Coupon No. not feeded", vbCritical, "Coupon No."
            End If
        Else
            If RsServ!serv_catg <> "P" Then
                Txt(CouponVal) = ""
                Txt(Coupon) = ""
            End If
        End If
    Case JobDt
        Txt(JobDt) = RetDate(Txt(JobDt))
        Cancel = Not CheckFinYear(Txt(Index))
        Txt(ArrDate) = Txt(JobDt)
        Txt(DelDt) = Txt(JobDt)
        
    Case SaleDate
        Txt(SaleDate).TEXT = RetDate(Txt(SaleDate))
        If Txt(SaleDate) <> "" Then
            If Format(Txt(SaleDate), "yyyymmdd") > Format(Txt(JobDt), "yyyymmdd") Then
                MsgBox "Sold Date " & Txt(SaleDate) & " is greater than Job Date " & Txt(JobDt) & " !", vbCritical, "Sale Date Validation"
                Cancel = True
            End If
        End If
    Case InvDate
        Txt(InvDate).TEXT = RetDate(Txt(InvDate))
        If Txt(InvDate) <> "" Then
            If Format(Txt(InvDate), "yyyymmdd") > Format(Txt(SaleDate), "yyyymmdd") Then
                MsgBox "Telco Invoice Date " & Txt(SaleDate) & " is greater than Sold Date " & Txt(JobDt) & " !", vbCritical, "Sale Date Validation"
                Cancel = True
            End If
        End If
    Case ArrDate
        Txt(ArrDate).TEXT = RetDate(Txt(ArrDate))
        If Txt(ArrDate) <> "" Then
            If Format(Txt(ArrDate), "yyyymmdd") > Format(Txt(JobDt), "yyyymmdd") Then
                MsgBox "Arrival Date " & Txt(ArrDate) & " is greater than Job Date " & Txt(JobDt) & " !", vbCritical, "Date Validation"
                Cancel = True
            End If
        End If
    Case DelDt
        Txt(DelDt).TEXT = RetDate(Txt(DelDt))
        If Txt(DelDt) <> "" Then
            If Format(Txt(DelDt), "yyyymmdd") < Format(Txt(JobDt), "yyyymmdd") Then
                MsgBox "Expected Delivery Date " & Txt(DelDt) & " is Less than Job Date " & Txt(JobDt) & " !", vbCritical, "Date Validation"
                Cancel = True
            End If
        End If
    Case Model
        If RsModel.EOF = False And RsModel.BOF = False Then
            If Txt(Model).TEXT <> "" Then
                Txt(Model).TEXT = RsModel!Code
                Txt(Model).Tag = RsModel!Code
            End If
        Else
            Txt(Model).TEXT = ""
            Txt(Model).Tag = ""
        End If
    Case Mechanic
        If RsMech.EOF = False And RsMech.BOF = False Then
            If Txt(Index).TEXT <> "" Then
                Txt(Index).TEXT = RsMech!Name
                Txt(Index).Tag = RsMech!Code
            End If
        Else
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        End If
'Nra upation
     Case Supervisor
        If RsSuper.EOF = False And RsSuper.BOF = False Then
            If Txt(Index).TEXT <> "" Then
                Txt(Index).TEXT = RsSuper!Name
                Txt(Index).Tag = RsSuper!Code
            End If
        Else
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        End If
' end update
    Case DCode
        If RsDealer.EOF = False And RsDealer.BOF = False Then
            If Txt(Index).TEXT <> "" Then
                Txt(Index).TEXT = RsDealer!Code
                Txt(Index).Tag = RsDealer!Code
                Txt(DNAME).TEXT = RsDealer!Name
                Txt(DNAME).Tag = RsDealer!Code
            End If
        Else
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
            Txt(DNAME).TEXT = ""
            Txt(DNAME).Tag = ""
        End If
    Case DNAME
        If RsDealer.EOF = False And RsDealer.BOF = False Then
            If Txt(Index).TEXT <> "" Then
                Txt(Index).TEXT = RsDealer!Name
                Txt(Index).Tag = RsDealer!Code
                Txt(DCode).TEXT = RsDealer!Code
            End If
        Else
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
            Txt(DCode).TEXT = ""
        End If
    Case InsuranceCompany
        If RsInsuranceCompany.EOF = False And RsInsuranceCompany.BOF = False Then
            If Txt(Index).TEXT <> "" Then
                Txt(Index).TEXT = RsInsuranceCompany!Name
                Txt(Index).Tag = RsInsuranceCompany!Code
            End If
        Else
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        End If
        
    Case City
        If RsCity.EOF = False And RsCity.BOF = False Then
            If Txt(Index).TEXT <> "" Then
                Txt(Index).TEXT = RsCity!Name
                Txt(Index).Tag = RsCity!Code
            End If
        Else
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        End If
    Case InsuranceExpiry
        Txt(Index) = RetDate(Txt(Index))
End Select
If Index = VehRegNo Or Index = Chassis Then
    If Txt(VehRegNo).TEXT = "" And Txt(Chassis).TEXT = "" Then
        Txt(HistNo).Tag = ""
        Txt(HistNo).TEXT = ""
        Call History_Enb(True)
    Else
        If Txt(BookSrl) = "" Then
            If RsHist.RecordCount > 0 Then
                If Index = VehRegNo And Trim(Txt(VehRegNo).TEXT) <> "" Then
                    RsHist.FIND ("Regno='" & Txt(VehRegNo).TEXT & "'")
                End If
                If Index = Chassis And Trim(Txt(Chassis).TEXT) <> "" Then
                    RsHist.FIND ("chassis='" & Txt(Chassis).TEXT & "'")
                End If
                If RsHist.EOF = True Or RsHist.BOF = True Then
                    Txt(HistNo).Tag = ""
                    Txt(HistNo).TEXT = ""
                    Call History_Enb(True)
                Else
                    Call History_Field(False)
                    Txt(HistNo).Tag = RsHist!CardNo
                    Txt(HistNo).TEXT = RsHist!CardNo
                End If
            End If
        End If
    End If
End If
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).TEXT = ""
        If I <> JobDt Then
            Txt(I).Tag = ""
        End If
    Next I
    
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
    
    FGrid2.FixedRows = 1
End Sub
Private Sub MoveRec()
Dim Master1 As ADODB.Recordset
Dim mVor$
Dim I As Integer
'On Error GoTo error1
    mAddFlag = ""
    If InStr(Me.TopCtrl1.Tag, "E") <> 0 Then Me.TopCtrl1.tEdit = True
    If InStr(Me.TopCtrl1.Tag, "D") <> 0 Then Me.TopCtrl1.tDel = True
    If Master.RecordCount > 0 Then
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "select JOB_Card.*, H.regno,H.chassis,H.model, H.engine, H.name, H.add1, H.add2,H.add3,H.citycode, H.phoneoff, H.phoneresi, " & _
                    " H.mobile,H.fuel_unit,H.vehserialno,H.dealer_code,H.delivery_date,Srv.Serv_type as SrvType,Srv.Serv_desc as SrvDesc, JB.Book_no, Jb.Book_date " & _
                    " From ((Job_card left join Service_Type Srv on Job_card.Serv_type=Srv.serv_type) left join Hiscard H on Job_Card.Cardno=H.cardno) left join Job_booking JB on Job_card.DocId=jb.job_DocID " & _
                    " where Job_Card.DocId='" & Master!DocID & "'", GCn, adOpenStatic, adLockReadOnly
        
        LblDiv.CAPTION = "Division : " & left(Master1!DocID, 1)
        LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        If RSOJPR = True Or Trim(UCase(left(PubComp_Name, 5))) = "KANOD" Then
            Label4.Visible = True: LblJobType.Visible = True
            If XNull(Master1!JobType) = "R" Or XNull(Master1!JobType) = "" Then
                LblJobType = "Regular"
                Txt(JobType) = "Regular"
            ElseIf XNull(Master1!JobType) = "O" Then
                LblJobType = "On Site Repair"
                Txt(JobType) = "On Site Repair"
            ElseIf XNull(Master1!JobType) = "Q" Then
                LblJobType = "Quick Repair"
                Txt(JobType) = "Quick Repair"
            End If
        Else
            Label4.Visible = False: LblJobType.Visible = False
        End If
        JobDocID = Master1!DocID
        lblDocId.CAPTION = "DocId : " & Master1!DocID
        lblPrefix.CAPTION = mID(Master1!DocID, 9, 5)
        If RSOJPR = True Then
            If XNull(Right(Master1!DocId_InvSpr, 8)) = "Cancelld" Then
                FrmCancel.Visible = True
            Else
                FrmCancel.Visible = False
            End If
        End If
        Txt(JobNo) = Master1!Job_No
        Txt(JobDt) = Master1!Job_Date
        Txt(JobClDt) = XNull(Master1!JobCloseDate)
        LblUser = IIf(Not IsNull(Master1!Created_AddDate), "Add By : " & XNull(Master1!Created_AddBy) & "  Dated : " & XNull(Master1!Created_AddDate), "") & IIf(Not IsNull(Master1!Created_ModifyDate), "     Modify By : " & XNull(Master1!Created_ModifyBy) & "  Dated : " & XNull(Master1!Created_ModifyDate), "")
'        GSQL = "SELECT Book_no, Book_Date from job_booking where Div_Code&Book_no&Site_Code ='" & Master1!Job_BookDivCode & Master1!Job_BookNo & Master1!Job_BookSiteCode & "'"
        GSQL = "SELECT Book_no, Book_Date from job_booking where Job_DocID ='" & Master!DocID & "'"
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
        
        If Rst.RecordCount > 0 Then
            Txt(BookSrl) = Master1!Job_BookNo
            Txt(BookSrl).Tag = Master1!Job_BookDivCode & Master1!Job_BookNo & Master1!Job_BookSiteCode
            Txt(BookDate) = XNull(Rst!Book_Date)
        Else
            Txt(BookSrl) = ""
            Txt(BookSrl).Tag = ""
            Txt(BookDate) = ""
        End If
        Set Rst = Nothing
        OldBookingID = Txt(BookSrl).Tag
        If Master1!Job_Inspdocid <> "" Then
            Txt(Ins_Sheet) = DeCodeDocID(Master1!Job_Inspdocid, Document_No)
            Txt(Ins_Sheet).Tag = Master1!Job_Inspdocid
            OldInsSheetID = Master1!Job_Inspdocid
        Else
            Txt(Ins_Sheet) = ""
            Txt(Ins_Sheet).Tag = ""
            OldInsSheetID = ""
        End If
        RsHist.Sort = "code"
        RsHist.FIND ("CODE='" & Master1!CardNo & "'")
        'RsModel.Sort = "code"
        RsModel.MoveFirst
        RsModel.FIND ("CODE='" & Master1!Model & "'")
        
        Call History_Field(False)
        Txt(VehRegNo) = XNull(Master1!RegNo)
        Txt(Chassis) = XNull(Master1!Chassis)
        Txt(Model).Tag = XNull(Master1!Model)
        Txt(Model) = XNull(Master1!Model)
        Txt(Engine) = XNull(Master1!Engine)
        Txt(VehSrlNo) = XNull(Master1!VehSerialNo)
        Txt(GovtYn) = IIf(Master1!Govt_YN = 1, "Yes", "No ")
        Txt(OwnerName) = XNull(Master1!Name)
        Txt(Address1) = XNull(Master1!Add1)
        Txt(Address2) = XNull(Master1!Add2)
        Txt(Address3) = XNull(Master1!Add3)
        Txt(PhoneOff) = XNull(Master1!PhoneOff)
        Txt(PhoneResi) = XNull(Master1!PhoneResi)
        Txt(Mobile) = XNull(Master1!Mobile)
        Txt(City).Tag = ""
        Txt(City) = ""
        If (Not IsNull(Master1!CityCode)) Or Master1!CityCode <> "" Then
            RsCity.Sort = "code"
            RsCity.FIND ("CODE='" & Replace(Master1!CityCode, "'", "") & "'")
            If RsCity.BOF = False And RsCity.EOF = False Then
                Txt(City) = RsCity!Name
                Txt(City).Tag = Master1!CityCode
            End If
        Else
            
        End If
        Txt(DNAME) = ""
        Txt(DNAME).Tag = ""
        Txt(DCode) = ""
        Txt(DCode).Tag = ""
        If (Not IsNull(Master1!dealer_code)) Or Master1!dealer_code <> "" Then
            RsDealer.Sort = "code"
            RsDealer.FIND ("CODE='" & Master1!dealer_code & "'")
            If RsDealer.BOF = False And RsDealer.EOF = False Then
                Txt(DNAME) = RsDealer!Name
                Txt(DNAME).Tag = RsDealer!Code
                Txt(DCode) = RsDealer!Code
                Txt(DCode).Tag = RsDealer!Code
            End If
        End If
        Txt(SaleDate) = IIf(IsNull(Master1!Delivery_Date), "", Master1!Delivery_Date)
        Txt(SrvType).Tag = XNull(Master1!SrvType)
        Txt(SrvType) = XNull(Master1!SrvDesc)
        Txt(Coupon) = XNull(Master1!Coupon)
        Txt(CouponVal) = XNull(Master1!Coupon_Value)
        Txt(CurrentKMS) = XNull(Master1!AtKMsHrs)
        Txt(KmsHrs) = IIf(XNull(Master1!KmsHrs) = "H", "Hrs.", "Kms.")
        Txt(FUEL) = XNull(Master1!FUEL)
        Txt(HrMeter) = XNull(Master1!HrMeter)
        LblFuel.CAPTION = XNull(Master1!Fuel_Unit)
        Txt(RecpTime) = Format(Master1!Recp_Time, "hh:mm")
        Txt(ArrDate) = Format(Master1!ArrivalTime, "dd/MMM/yyyy")
        Txt(ArrTime) = Format(Master1!ArrivalTime, "hh:mm")
        Txt(EstSpr) = Format(Master1!Est_SpCost, "0.00")
        Txt(EstLab) = Format(Master1!Est_LabCost, "0.00")
        Txt(DelDt) = Format(Master1!ExpDelDate, "dd/MMM/yyyy")
        Txt(DelTime) = Format(Master1!ExpDelDate, "hh:mm")
        Txt(Damage) = XNull(Master1!body_damage)
        Txt(Remarks) = XNull(Master1!OpenRemarks)
        
        Txt(Mechanic).Tag = ""
        Txt(Mechanic) = ""
        If (Not IsNull(Master1!RecBy_Mechanic)) Or Master1!RecBy_Mechanic <> "" Then
            RsMech.Sort = "code"
            RsMech.FIND ("CODE='" & Master1!RecBy_Mechanic & "'")
            If RsMech.BOF = False And RsMech.EOF = False Then
                Txt(Mechanic) = RsMech!Name
                Txt(Mechanic).Tag = RsMech!Code
            End If
        End If
        
        Txt(Supervisor).Tag = ""
        Txt(Supervisor) = ""
        If (Not IsNull(Master1!RecBy_Supervisor)) Or Master1!RecBy_Supervisor <> "" Then
            RsSuper.Sort = "code"
            RsSuper.FIND ("CODE='" & Master1!RecBy_Supervisor & "'")
            If RsSuper.BOF = False And RsSuper.EOF = False Then
                Txt(Supervisor) = RsSuper!Name
                Txt(Supervisor).Tag = RsSuper!Code
            End If
        End If
        Txt(HistNo) = Master1!CardNo
        Txt(HistNo).Tag = Master1!CardNo
        
        If IsNull(Master1!Recp_Time) Then
            Txt(JCTime) = "00:00"
        Else
            Txt(JCTime) = Format(Master1!Recp_Time, "hh:mm")
        End If
        
        FGrid1.Rows = 1
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select J.S_NO,J.CODE,J.Details, J.Lab_Code, L.Lab_Desc, J.Lab_Rate, J.Time_Req, J.Amount  From Job_Demand J Left Join Labour L On J.Lab_Code=L.Lab_Code Where j.Job_DocId='" & Master1!DocID & "'", GCn, adOpenStatic, adLockReadOnly
                
        I = 1
        If Rst.RecordCount > 0 Then
            Do Until Rst.EOF
                FGrid1.AddItem ""
                         
                FGrid1.TextMatrix(I, 0) = Rst!S_No
                FGrid1.TextMatrix(I, Col_Code) = Rst!Code
                FGrid1.TextMatrix(I, Col_Trouble) = Rst!Details
                If UCase(left(PubComp_Name, 3)) = "LMP" Then
                    FGrid1.TextMatrix(I, Col_Lab_Code) = XNull(Rst!Lab_Code)
                    FGrid1.TextMatrix(I, Col_Lab_Desc) = XNull(Rst!Lab_Desc)
                    FGrid1.TextMatrix(I, Col_Lab_Rate) = VNull(Rst!Lab_Rate)
                    FGrid1.TextMatrix(I, Col_Time_Req) = VNull(Rst!TIME_REQ)
                    FGrid1.TextMatrix(I, Col_Amount) = VNull(Rst!Amount)
                    'FGrid1.AddItem Rst!S_No & Chr(9) & Rst!Code & Chr(9) & Rst!Details & Chr(9) & IIf(UCase(left(PubComp_Name, 3)) = "LMP", Chr(9) & Rst!Lab_Code & Chr(9) & Rst!Lab_Desc, "")
                End If
                
                I = I + 1
                Rst.MoveNext
            Loop
        Else
            FGrid1.Rows = FGrid1.Rows
            FGrid1.AddItem ""
        End If
        FGrid1.FixedRows = 1
        
        '' Note : Following Lines are for updation of grid2 with default value
        For I = 1 To FGrid2.Rows - 1
            FGrid2.TextMatrix(I, ElValue) = FGrid2.TextMatrix(I, ElDefault)
        Next I
        
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select JI.S_NO,JI.Element_code as code,JI.In_Status as Value1,JI.Remarks From Job_Inspection2 as JI Where JI.DocId='" & Master1!DocID & "'", GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount > 0 Then
            Do Until Rst.EOF
                For I = 1 To FGrid2.Rows - 1
                    If FGrid2.TextMatrix(I, ElCode) = Rst!Code Then
                        If FGrid2.TextMatrix(I, ElNature) = "Numeric" Then
                            FGrid2.TextMatrix(I, ElValue) = Val(Rst!Remarks)
                        ElseIf FGrid2.TextMatrix(I, ElNature) = "Character" Then
                            FGrid2.TextMatrix(I, ElValue) = Rst!Remarks
                        ElseIf FGrid2.TextMatrix(I, ElNature) = "Boolean" Then
                            FGrid2.TextMatrix(I, ElValue) = Rst!Remarks
                        End If
                    End If
                Next I
                Rst.MoveNext
            Loop
            FGrid2.FixedRows = 1
        End If
        Set Rst = Nothing
        UpdLastJC
        Call veh_count
'        If Txt(JobClDt) = "" Then
'            TopCtrl1.tEdit = CheckPerm(Ed, Me)
'            TopCtrl1.tDel = CheckPerm(De, Me)
'        Else
'            TopCtrl1.tEdit = False
'            TopCtrl1.tDel = False
'        End If
        If GCn.Execute("Select Job_Docid from Job_Lab where Job_Docid='" & Master!DocID & "'").RecordCount > 0 Or _
            GCn.Execute("Select Job_Docid from SP_Stock where Job_Docid='" & Master!DocID & "'").RecordCount > 0 Then
            TopCtrl1.tDel = False
        End If
        If Txt(JobClDt) <> "" Or FrmCancel.Visible = True Then
            TopCtrl1.tEdit = False
            TopCtrl1.tDel = False
        End If
    End If
    Set Master1 = Nothing
    Grid_Hide
    FGrid1_GotFocus
    Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
    
    With FGrid1
        .left = 0
        .top = 4470
        .width = IIf(UCase(left(PubComp_Name, 3)) = "LMP", 7395, 5535)
        .RowHeightMin = PubGridRowHeight '220
        .height = FGrid1.RowHeight(0) * 7
        .Cols = 6
        
        .TextMatrix(0, 0) = "Srl."
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 500
        
        .TextMatrix(1, 0) = "1"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 500
        
        .TextMatrix(0, Col_Code) = ""
        .ColAlignment(Col_Code) = flexAlignLeftCenter
        .ColWidth(Col_Code) = 0
        
    
        .TextMatrix(0, Col_Trouble) = "Owner/Driver's Compalaints"
        .ColAlignment(Col_Trouble) = flexAlignLeftCenter
        .ColWidth(Col_Trouble) = IIf(UCase(left(PubComp_Name, 3)) = "LMP", 3250, 4750)
        
        .TextMatrix(0, Col_Repeat) = ""
        .ColAlignment(Col_Repeat) = flexAlignLeftCenter
        .ColWidth(Col_Repeat) = 0
        
        .ColWidth(Col_Lab_Code) = 0
        
        
        .TextMatrix(0, Col_Lab_Desc) = "Applicable Labour"
        .ColAlignment(Col_Lab_Desc) = flexAlignLeftCenter
        .ColWidth(Col_Lab_Desc) = IIf(UCase(left(PubComp_Name, 3)) = "LMP", 3250, 0)
        
        
        .ColWidth(Col_Lab_Rate) = 0
        .ColWidth(Col_Time_Req) = 0
        .ColWidth(Col_Amount) = 0
        
        
    End With
    BackColorSelLeave = FGrid1.BackColorSel
    ForeColorSelEnter = FGrid1.ForeColorSel
    
    With FGrid2
        .left = IIf(UCase(left(PubComp_Name, 3)) = "LMP", 7500, 5640)
        .width = IIf(UCase(left(PubComp_Name, 3)) = "LMP", 4300, 6150)
        .top = 4470
        .RowHeightMin = PubGridRowHeight '220
        .height = FGrid1.RowHeight(0) * 7
        .Cols = 6
        
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 500
        
        .TextMatrix(0, ElCode) = "Code"
        .ColAlignment(ElCode) = flexAlignLeftCenter
        .ColWidth(ElCode) = 0

        .TextMatrix(0, ElNature) = "Field Nature/Type"
        .ColAlignment(ElNature) = flexAlignLeftCenter
        .ColWidth(ElNature) = 0
    
        .TextMatrix(0, ElDefault) = "Default value of Field"
        .ColAlignment(ElDefault) = flexAlignLeftCenter
        .ColWidth(ElDefault) = 0
        
        .TextMatrix(0, ElName) = "Particulars"
        .ColAlignment(ElName) = flexAlignLeftCenter
        .ColWidth(ElName) = IIf(UCase(left(PubComp_Name, 3)) = "LMP", 2400, 3645)
        
        .TextMatrix(0, ElValue) = "Value"
        .ColAlignment(ElValue) = flexAlignLeftCenter
        .ColWidth(ElValue) = IIf(UCase(left(PubComp_Name, 3)) = "LMP", 1000, 1695)
    End With
    
    DGBook.width = Me.width - 60: DGBook.left = FGrid1.left: DGBook.top = 3705: DGBook.height = Me.height - (DGBook.top + mBotScale) 'FGrid1.height
    DgModel.width = DGBook.width: DgModel.left = DGBook.left: DgModel.top = DGBook.top: DgModel.height = DGBook.height
    DGInsp.width = DGBook.width: DGInsp.left = DGBook.left: DGInsp.top = DGBook.top: DGInsp.height = DGBook.height
    DGHist.width = DGBook.width: DGHist.left = DGBook.left: DGHist.top = DGBook.top: DGHist.height = DGBook.height
    DGDealer.width = DGBook.width: DGDealer.left = DGBook.left: DGDealer.top = DGBook.top: DGDealer.height = DGBook.height
    '6525
    DGService.width = 4100: DGService.left = Me.width - (DGService.width + mRtScale): DGService.top = mTopScale: DGService.height = 2865
    DGCity.width = 4100: DGCity.left = Me.width - (DGCity.width + mRtScale): DGCity.top = mTopScale: DGCity.height = 2865
    DGMech.width = 5160: DGMech.left = Me.width - (DGMech.width + mRtScale): DGMech.top = DGBook.top: DGMech.height = Me.height - (DGMech.top + mBotScale) '2865
    DGTrouble.left = Me.width - (DGTrouble.width + mRtScale): DGTrouble.top = mTopScale: DGTrouble.height = Me.height - (DGTrouble.top + mBotScale) 'FGrid2.height
    DGLab.left = Me.width - (DGLab.width + mRtScale): DGLab.top = mTopScale: DGLab.height = Me.height - (DGLab.top + mBotScale) 'FGrid2.height
End Sub

Public Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
    txtgrid1(1).BackColor = CtrlBColOrg
    txtgrid1(1).ForeColor = CtrlFColOrg
    txtgrid1(1).Enabled = Enb
    
    txtgrid2(1).BackColor = CtrlBColOrg
    txtgrid2(1).ForeColor = CtrlFColOrg
    txtgrid2(1).Enabled = Enb
    
    Txt(LastJobDt).Enabled = False
    Txt(LastJobNo).Enabled = False
    Txt(LastSrv).Enabled = False
    Txt(LastKMS).Enabled = False
    Txt(JobNo).Enabled = False
    Txt(LastMech).Enabled = False
    Txt(HistNo).Enabled = False
    Txt(BookDate).Enabled = False
'    txt(SrvRate).Enabled = false
End Sub

Private Sub Grid_Hide()
    If DGBook.Visible = True Then DGBook.Visible = False
    If DgModel.Visible = True Then DgModel.Visible = False
    If DGService.Visible = True Then DGService.Visible = False
    If DGDealer.Visible = True Then DGDealer.Visible = False
    If DGMech.Visible = True Then DGMech.Visible = False
    If DGHist.Visible = True Then DGHist.Visible = False
    If DGCity.Visible = True Then DGCity.Visible = False
    If DGInsp.Visible = True Then DGInsp.Visible = False
    If DGTrouble.Visible = True Then DGTrouble.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGLab.Visible = True Then DGLab.Visible = False
End Sub

Private Sub veh_count()
    If Txt(JobDt).TEXT <> "" Then
        LblTotVeh.CAPTION = GCn.Execute("select count(*) from job_Card where Job_Date=" & ConvertDate(Txt(JobDt).TEXT) & "  and left(Docid,1)='" & PubDivCode & "'").Fields(0)
    End If
End Sub

Private Sub UpdRequery()
    RSBook.Requery
    RsInsp.Requery
    RsHist.Requery
    RsModel.Requery
    RsServ.Requery
    RsDealer.Requery
    RsMech.Requery
    Set DGMech.DataSource = RsMech
    RsSuper.Requery
    RsCity.Requery
    RsTrb.Requery
'    RsElem.Requery
End Sub

Private Sub History_Field(flag As Boolean, Optional GridClick As Boolean)
If RsHist.EOF Then MsgBox "History Card Not found", vbCritical, "Error Message": Exit Sub
Dim rsHist1  As ADODB.Recordset
    Set rsHist1 = GCn.Execute("Select GOVT_YN,HisCard.Add1,HisCard.Add2,HisCard.Add3,HISCARD.CityCode,City.CityName," & _
        " PhoneOff,PhoneResi,Mobile,Dealer_Code,VehSerialNo,Fuel_Unit,Locked_Text," & _
        " amd_dealer.D_Name as DealerName,Delivery_Date,SUPPLIER_BILLNO, InsuranceCompany, InsuranceExpiry, InsurancePolicyNo, Insurance.Name As InsuranceCompanyName " & _
        " FROM ((Hiscard " & _
        " left join Insurance on Hiscard.InsuranceCompany=Insurance.Code) " & _
        " left join city on Hiscard.CityCode=City.CityCode) " & _
        " left join amd_dealer on hiscard.dealer_code=amd_dealer.d_code " & _
        " Where HISCARD.CardNo='" & RsHist!Code & "'")
    
    Txt(HistNo).Tag = RsHist!CardNo
    Txt(HistNo) = RsHist!CardNo
    If MyIndex = Chassis Then
        Txt(VehRegNo) = XNull(RsHist!RegNo)
        If GridClick = True Then
            Txt(Chassis) = XNull(RsHist!Chassis)
        End If
    ElseIf MyIndex = VehRegNo Then
        If GridClick = True Then
            Txt(VehRegNo) = XNull(RsHist!RegNo)
        End If
        Txt(Chassis) = XNull(RsHist!Chassis)
    Else
        Txt(VehRegNo) = XNull(RsHist!RegNo)
        Txt(Chassis) = XNull(RsHist!Chassis)
    End If
    Txt(Model).Tag = XNull(RsHist!Model)
    Txt(Model) = XNull(RsHist!Model)
    Txt(Engine) = XNull(RsHist!Engine)
    
    Txt(VehSrlNo) = XNull(rsHist1!VehSerialNo)
    Txt(GovtYn) = IIf(rsHist1!Govt_YN = 0, "No", "Yes")
    Txt(OwnerName) = XNull(RsHist!Name)
    Txt(Address1) = XNull(rsHist1!Add1)
    Txt(Address2) = XNull(rsHist1!Add2)
    Txt(Address3) = XNull(rsHist1!Add3)
    Txt(City).Tag = XNull(rsHist1!CityCode)
    Txt(City) = XNull(rsHist1!CityName)
    Txt(PhoneOff) = XNull(rsHist1!PhoneOff)
    Txt(PhoneResi) = XNull(rsHist1!PhoneResi)
    Txt(Mobile) = XNull(rsHist1!Mobile)
    Txt(DCode) = XNull(rsHist1!dealer_code)
    Txt(DNAME) = XNull(rsHist1!DealerName)
    Txt(SaleDate) = XNull(rsHist1!Delivery_Date)
    Txt(InvNo) = XNull(rsHist1!SUPPLIER_BILLNO)
    Txt(InsuranceCompany).Tag = XNull(rsHist1!InsuranceCompany)
    Txt(InsuranceCompany) = XNull(rsHist1!InsuranceCompanyName)
    Txt(InsuranceExpiry) = XNull(rsHist1!InsuranceExpiry)
    Txt(InsurancePolicyNo) = XNull(rsHist1!InsurancePolicyNo)
    LblFuel.CAPTION = XNull(rsHist1!Fuel_Unit)
    If rsHist1.RecordCount > 0 Then
        If rsHist1!Locked_Text <> "" Then
            MsgBox rsHist1!Locked_Text, vbCritical, "Previous Job Close Remarks"
        End If
    End If
    Set rsHist1 = Nothing
    Dim I As Integer
    If mAddFlag = "A" Then
       Set rsHist1 = GCn.Execute("Select DocID,Job_Date,JobCloseDate FROM Job_Card " & _
            " Where Job_Card.CardNo='" & RsHist!Code & "'")
       If rsHist1.RecordCount > 0 Then
            For I = 1 To rsHist1.RecordCount
                If IsNull(rsHist1!JobCloseDate) = True Or rsHist1!JobCloseDate = "" Then
                    MsgBox "Job No.: " & PrinID(rsHist1!DocID) & vbCrLf & " Dt.: " & rsHist1!Job_Date & vbCrLf & "Already exists for selected Vehicle", vbCritical, "Job Already Exists"
                End If
            rsHist1.MoveNext
            Next
       End If
        Set rsHist1 = Nothing
        
    End If
    Call UpdLastJC
    Call History_Enb(flag)
End Sub

Private Sub History_Enb(flag As Boolean)
    Txt(Model).Enabled = flag
    Txt(Engine).Enabled = flag
    Txt(VehSrlNo).Enabled = flag
    Txt(GovtYn).Enabled = flag
    Txt(OwnerName).Enabled = flag
    Txt(Address1).Enabled = flag
    Txt(Address2).Enabled = flag
    Txt(Address3).Enabled = flag
    Txt(City).Enabled = flag
    Txt(PhoneOff).Enabled = flag
    Txt(PhoneResi).Enabled = flag
    Txt(Mobile).Enabled = flag
    Txt(DCode).Enabled = flag
    Txt(DNAME).Enabled = flag
    Txt(SaleDate).Enabled = flag
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
Ctrl_GetFocus txtgrid1(Index)
    Grid_Hide
    txtgrid1(1).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    Select Case FGrid1.Col
        Case Col_Trouble
            If RsTrb.RecordCount = 0 Or FGrid1.TextMatrix(FGrid1.Row, Col_Trouble) = "" Then Exit Sub
            RsTrb.Sort = "name"
            RsTrb.MoveFirst
            RsTrb.FIND "name ='" & FGrid1.TextMatrix(FGrid1.Row, Col_Trouble) & "'"
            If RsTrb.EOF = True Then RsTrb.MoveFirst
            
        Case Col_Lab_Desc
            Set RsLab = GCn.Execute("Select Lab_Code As Code, Lab_Desc As Name, Lab_Rate, Time_Req, Lab_Rate*Time_Req As Amount  From Labour Where Lab_Code In (Select Lab_Code From Lab_Trouble Where CCCode='" & FGrid1.TextMatrix(FGrid1.Row, Col_Code) & "') Order By NAME")
            Set DGLab.DataSource = RsLab
        
            If RsLab.RecordCount = 0 Or FGrid1.TextMatrix(FGrid1.Row, Col_Lab_Code) = "" Then Exit Sub
            'RsLab.Sort = "Code"
            RsLab.MoveFirst
            RsLab.FIND "Code ='" & FGrid1.TextMatrix(FGrid1.Row, Col_Lab_Code) & "'"
            If RsLab.EOF = True Then RsLab.MoveFirst

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
        Case Col_Trouble
            If DGTrouble.Visible = False Then DGridColSwap DGTrouble, 1
'            DGridTxtKeyDown_Mast DGTrouble, TxtGrid1, Index, RsTrb, KeyCode, False, 1
            DGridTxtKeyDown DGTrouble, txtgrid1, Index, RsTrb, KeyCode, False, 1, frmTrouble, "frmTrouble"
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True And DGTrouble.Visible = False) Then
                If TxtGrid1Leave = True Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, IIf(UCase(left(PubComp_Name, 3)) = "LMP", Col_Lab_Desc, 1), 1
                End If
            End If
            
        Case Col_Lab_Desc
            If DGLab.Visible = False Then DGridColSwap DGLab, 1
            DGridTxtKeyDown DGLab, txtgrid1, Index, RsLab, KeyCode, False, 1, frmLabDesc, "frmLabDesc"
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True And DGLab.Visible = False) Then
                If TxtGrid1Leave = True Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, 1
                End If
            End If
            
    End Select
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
    Select Case FGrid1.Col
        Case Col_Trouble
            DGridTxtKeyPress txtgrid1, Index, RsTrb, KeyAscii, "name"
        Case Col_Lab_Desc
            DGridTxtKeyPress txtgrid1, Index, RsLab, KeyAscii, "Name"
    End Select
End Sub

Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        Select Case FGrid1.Col
            Case Col_Trouble
                If KeyCode <> 13 And DGTrouble.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0
                DGridTxtKeyUp_Mast txtgrid1, Index, RsTrb, KeyCode, "name"
            Case Col_Lab_Desc
                If KeyCode <> 13 And DGLab.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0
                DGridTxtKeyUp_Mast txtgrid1, Index, RsLab, KeyCode, "Name"
        End Select
End Sub

Private Sub TxtGrid1_LostFocus(Index As Integer)
    If ExitCtrl = False Then Exit Sub
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
Dim I As Integer
Select Case FGrid1.Col

    Case Col_Trouble
        If RsTrb.RecordCount = 0 Or txtgrid1(1).TEXT = "" Then
            FGrid1.TextMatrix(FGrid1.Row, Col_Code) = ""
            FGrid1.TextMatrix(FGrid1.Row, Col_Trouble) = ""
        Else
            FGrid1.TextMatrix(FGrid1.Row, Col_Code) = RsTrb!Code
            If UCase(Trim(txtgrid1(1).TEXT)) <> UCase(left(RsTrb!Name, Len(Trim(txtgrid1(1).TEXT)))) Then
                FGrid1.TextMatrix(FGrid1.Row, Col_Trouble) = txtgrid1(1).TEXT
            Else
                txtgrid1(1).TEXT = RsTrb!Name
                FGrid1.TextMatrix(FGrid1.Row, Col_Trouble) = RsTrb!Name
            End If
        End If
            Repeat = CheckRepeat(Txt(VehRegNo), FGrid1.Row)
            If Repeat <> "" Then
                If MsgBox("This Complaint is repeated from " & Repeat & "Mark Repeated !", vbYesNo) = vbNo Then
                    FGrid1.TextMatrix(FGrid1.Row, Col_Repeat) = 0
                Else
                    FGrid1.TextMatrix(FGrid1.Row, Col_Repeat) = 1
                End If
            End If
        'If FGrid1.TextMatrix(FGrid1.Rows - 1, 1) <> "" Then FGrid1.AddItem FGrid1.Rows
        
    Case Col_Lab_Desc
        If RsLab.RecordCount = 0 Or RsLab.EOF = True Or RsLab.EditMode = True Or txtgrid1(1) = "" Then
            FGrid1.TextMatrix(FGrid1.Row, Col_Lab_Code) = ""
            FGrid1.TextMatrix(FGrid1.Row, Col_Lab_Desc) = ""
            FGrid1.TextMatrix(FGrid1.Row, Col_Lab_Rate) = ""
            FGrid1.TextMatrix(FGrid1.Row, Col_Time_Req) = ""
            FGrid1.TextMatrix(FGrid1.Row, Col_Amount) = ""
            
        Else
            FGrid1.TextMatrix(FGrid1.Row, Col_Lab_Code) = RsLab!Code
            FGrid1.TextMatrix(FGrid1.Row, Col_Lab_Desc) = RsLab!Name
            FGrid1.TextMatrix(FGrid1.Row, Col_Lab_Rate) = VNull(RsLab!Lab_Rate)
            FGrid1.TextMatrix(FGrid1.Row, Col_Time_Req) = VNull(RsLab!TIME_REQ)
            FGrid1.TextMatrix(FGrid1.Row, Col_Amount) = VNull(RsLab!Amount)
            Txt(EstLab) = "0.00"
            For I = 1 To FGrid1.Rows - 1
                Txt(EstLab) = Val(Txt(EstLab)) + Val(FGrid1.TextMatrix(I, Col_Amount))
            Next I
            Txt(EstLab) = Format(Txt(EstLab), "0.00")
        End If
        
End Select
TxtGrid1Leave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid1.SetFocus
    txtgrid1(1).Visible = False
End If
End Function

Private Sub UpdLastJC()
Dim RsTemp As ADODB.Recordset
    If Txt(JobDt) = "" Then Exit Sub
    Set RsTemp = New ADODB.Recordset
    RsTemp.CursorLocation = adUseClient
    RsTemp.Open "SELECT Top 1 JOB_NO,JOB_DATE,AtKMsHrs,Srv.Serv_SrlNo,Srv.Serv_Type,Srv.SERV_DESC AS SrvDesc,EMP_MAST.EMP_NAME AS MECH_NAME " & _
            " FROM (JOB_CARD LEFT JOIN Service_Type Srv ON JOB_CARD.SERV_TYPE=Srv.SERV_TYPE) " & _
            " LEFT JOIN EMP_MAST ON JOB_CARD.RECBY_MECHANIC=EMP_MAST.EMP_CODE " & _
            " WHERE CARDNO='" & Txt(HistNo).TEXT & _
            "' and Job_Date< " & ConvertDate(Txt(JobDt)) & _
            " ORDER BY JOB_DATE Desc ", GCn, adOpenStatic, adLockReadOnly
    If RsTemp.RecordCount > 0 Then
        Txt(LastJobNo).TEXT = RsTemp!Job_No
        Txt(LastJobDt).TEXT = RsTemp!Job_Date
        Txt(LastKMS).TEXT = RsTemp!AtKMsHrs
        Txt(LastSrv).TEXT = RsTemp!SrvDesc
        Txt(LastSrv).Tag = RsTemp!Serv_Type
        Txt(LastMech).TEXT = IIf(IsNull(RsTemp!MECH_NAME), "*No Mechanic*", RsTemp!MECH_NAME)
        
    Else
        Txt(LastJobNo).TEXT = "":           Txt(LastJobDt).TEXT = ""
        Txt(LastKMS).TEXT = "":             Txt(LastSrv).TEXT = ""
        Txt(LastMech).TEXT = "":            Txt(LastSrv).Tag = ""
    End If
    Set RsTemp = Nothing
End Sub

Private Sub TxtGrid2_GotFocus(Index As Integer)
On Error GoTo ELoop
If ExitCtrl = False Then Exit Sub
    Ctrl_GetFocus txtgrid2(Index)
'    FGrid2.CellBackColor = CellBackColLeave
    txtgrid2(1).Tag = FGrid2.TextMatrix(FGrid2.Row, FGrid2.Col)
    Select Case FGrid2.Col
        Case ElValue
            If FGrid2.TextMatrix(FGrid2.Row, ElNature) = "Character" Then
                txtgrid2(1).MaxLength = 25
            ElseIf FGrid2.TextMatrix(FGrid2.Row, ElNature) = "Numeric" Then
                txtgrid2(1).MaxLength = 10
            ElseIf FGrid2.TextMatrix(FGrid2.Row, ElNature) = "Boolean" Then
                txtgrid2(1).MaxLength = 4
            End If
    End Select
    Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        txtgrid2(1).TEXT = txtgrid2(1).Tag
        TxtGrid2_KeyUp Index, KeyCode, Shift
        FGrid2.SetFocus
        txtgrid2(1).Visible = False
        Exit Sub
    End If
    Select Case FGrid2.Col
        Case ElValue
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True And DGTrouble.Visible = False) Then
                TxtGrid2Leave
                If KeyCode = vbKeyReturn Then
                    If FGrid2.Row < FGrid2.Rows - 1 Then
                        FGrid2.Row = FGrid2.Row + 1
                    End If
                    FGrid2.Col = ElValue
                End If
            End If
    End Select
    Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txtGrid2_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
    Select Case FGrid2.Col
        Case ElValue
            If FGrid2.TextMatrix(FGrid2.Row, ElNature) = "Character" Then
            
            ElseIf FGrid2.TextMatrix(FGrid2.Row, ElNature) = "Numeric" Then
                NumPress txtgrid2(Index), KeyAscii, 8, 2
            ElseIf FGrid2.TextMatrix(FGrid2.Row, ElNature) = "Boolean" Then
        
            End If
    End Select
End Sub

Private Sub TxtGrid2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case FGrid2.Col
        Case ElValue
            If FGrid2.TextMatrix(FGrid2.Row, ElNature) = "Character" Then
            
            ElseIf FGrid2.TextMatrix(FGrid2.Row, ElNature) = "Numeric" Then
                
            ElseIf FGrid2.TextMatrix(FGrid2.Row, ElNature) = "Boolean" Then
                If Len(txtgrid2(Index)) = 0 Or UCase(mID(txtgrid2(Index), 1, 1)) = "N" Then
                    txtgrid2(Index) = "No"
                ElseIf UCase(mID(txtgrid2(Index), 1, 1)) = "Y" Then
                    txtgrid2(Index) = "Yes"
                Else
                    txtgrid2(Index) = "No"
                End If
            End If
    End Select
End Sub

Private Sub TxtGrid2_LostFocus(Index As Integer)
    If ExitCtrl = False Then Exit Sub
    txtgrid2(Index).Visible = False
End Sub

Private Sub TxtGrid2_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGrid2Leave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid2Leave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean

Select Case FGrid2.Col
    Case ElValue
        FGrid2.TextMatrix(FGrid2.Row, FGrid2.Col) = txtgrid2(1).TEXT
End Select
txtgrid2(1).MaxLength = 25
txtgrid1(1).Visible = False
TxtGrid2Leave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid2.SetFocus
    txtgrid2(1).Visible = False
End If
End Function

Private Sub FillBookingData()
            
        If RSBook!RegNo <> "" Then
            RsHist.FIND ("Regno='" & RSBook!RegNo & "'")
        ElseIf RSBook!Chassis <> "" Then
            RsHist.FIND ("chassis='" & RSBook!Chassis & "'")
        End If
        
        GSQL = "SELECT J.Name As OwnerName, Add1, Add2, Add3, J.CityCode, CityName, PhoneOff, PhoneResi, Mobile, Model, Engine,  " & _
               "Advance,ForServiceDate,Remarks,Service_Type, Service_Type.serv_desc " & _
               "from (Job_Booking J " & _
               "left join Service_Type on J.Service_Type=Service_Type.serv_type) " & _
               "Left Join City C On J.CityCode=C.CityCode " & _
               "where Book_No=" & Txt(BookSrl).TEXT & " "
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
        
        If RsHist.EOF = True Or RsHist.BOF = True Then
            Txt(HistNo).Tag = ""
            Txt(HistNo).TEXT = ""
            
            Txt(VehRegNo).TEXT = RSBook!RegNo
            Txt(Chassis).TEXT = RSBook!Chassis
            Txt(Model).Tag = Rst!Model
            Txt(Model).TEXT = Rst!Model
            Txt(Engine).TEXT = Rst!Engine
            Txt(OwnerName).TEXT = Rst!OwnerName
            Txt(Address1).TEXT = XNull(Rst!Add1)
            Txt(Address2).TEXT = XNull(Rst!Add2)
            Txt(Address3).TEXT = XNull(Rst!Add3)
            Txt(City).Tag = XNull(Rst!CityCode)
            Txt(City) = XNull(Rst!CityName)
            Txt(PhoneOff).TEXT = XNull(Rst!PhoneOff)
            Txt(PhoneResi).TEXT = XNull(Rst!PhoneResi)
            Txt(Mobile).TEXT = XNull(Rst!Mobile)
            Call History_Enb(True)
        Else
            Call History_Field(False)
        End If
        Txt(SrvType).Tag = XNull(Rst!Service_Type)
        Txt(SrvType).TEXT = XNull(Rst!Serv_Desc)
        Txt(BookDate).TEXT = XNull(RSBook!Book_Date)
        Set Rst = Nothing

End Sub
Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        FrmPrn.Visible = False
        If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
            If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
            Disp_Text SETS("INI", Me, Master)
            Call MoveRec
        End If
    End If
End Sub

Private Sub CmdPrint_Click(Index As Integer)
On Error Resume Next
Dim PrePrn As Boolean
Dim I As Integer
PrePrn = IIf(Optpre.Value = True, True, False)
GSQL = "SELECT syctrl.jobcardfooter,Emp_Mast.Emp_Name As MechName,City.CityName,H.Model,H.RegNo,H.RegDate,H.Chas_Type,H.Chassis,H.Engine,H.VehSerialNo,H.Supplier_BillNo," & _
    "H.Supplier_BillDate ,H.Delivery_Date,H.Dealer_Code,H.CouponNo,H.GBoxNo,H.RAxelNo,H.FAxelNo,H.TransAxelNo,H.SteerGBNo,H.CabinNo,H.BodyNo,Model.Col_Code As ColourCode ,H.DoorLockNo," & _
    "H.SteerLockNo,H.FillerCapLockNo,H.HeadLampMake,H.FIP_No,H.Frame_No,H.Steer_Type ,H.Steer_Make,H.Alternator,H.StarterMotor,H.Battery,H.Brake_Type,H.Radiator_Make," & _
    "H.Tyre_FL,H.Tyre_FR,H.Tyre_ML1,H.Tyre_ML2,H.Tyre_MR1,H.Tyre_MR2,H.Tyre_RL1,H.Tyre_RL2,H.Tyre_RR1," & _
    "H.Tyre_RR2,H.Spare_Wheel,H.Addl_Equp,H.Fuel as FuelType,H.Fuel_Unit,H.Name ,H.ConPerson,H.Add1,H.Add2,H.Add3,H.AREA,H.CityCode,H.Pin,H.PhoneOff,H.PhoneResi,H.Govt_YN," & _
    "J.DocId ,J.Job_No,J.Job_Date,J.Job_BookDivCode,J.Job_BookNo,J.Job_BookSiteCode,J.Job_InspDocID,J.CardNo As HisCardNo,J.Govt_YN As GovtYn,J.Serv_Type,j.Serv_Rate ,j.Coupon,j.Coupon_Value,j.AtKMsHrs," & _
    "j.FUEL,j.Est_SpCost,j.Est_LabCost,j.ArrivalTime,j.Recp_Time ,j.ExpDelDate,j.Body_Damage,j.OpenRemarks, j.RecBy_Mechanic,j.RecBy_Supervisor,j.CreatedU_Name,j.CreatedU_EntDt,j.CreatedU_AE,j.Remark,j.U_Name,j.U_EntDt," & _
    "Model.Model_Desc,Model.RLW,Model.Manufacturer,Service_Type.Serv_Desc,Job_Demand.S_No, Job_Demand.Details, Job_Demand.Code As Cust_Comp_Code, Job_Demand.Repeat_YN,J.HrMeter,H.Mobile, Col.Col_Desc, Ad.D_Name As DealerName, Job_demand.Lab_Code, L.Lab_Desc, L.Time_Req As Lab_Time, L.Time_Req*L.Lab_Rate As Est_Lab_Amount, Emp_Mast_1.Emp_Name As Supervisor_Name, H.Varient,Model.UNLADEN_WT AS ULW " & _
    "FROM ((((((((((Job_Card as J LEFT JOIN HisCard as H ON J.CardNo = H.CardNo) " & _
    "LEFT JOIN Emp_Mast ON J.RecBy_Mechanic = Emp_Mast.Emp_Code) " & _
    "LEFT JOIN Emp_Mast AS Emp_Mast_1 ON J.RecBy_Supervisor = Emp_Mast_1.Emp_Code) " & _
    "LEFT JOIN City ON H.CityCode = City.CityCode) LEFT JOIN Model ON H.Model = Model.MODEL) " & _
    "LEFT JOIN Service_Type ON J.Serv_Type = Service_Type.Serv_Type) LEFT JOIN Job_Demand ON J.DocId = Job_Demand.Job_DocID) " & _
    "Left Join ColMast Col On Col.col_Code=Model.Col_Code) Left Join Amd_Dealer AD On AD.D_Code = H.Dealer_Code ) Left Join Labour L On L.Lab_Code =  Job_Demand.Lab_Code)" & _
    "LEFT JOIN Syctrl ON Syctrl.LinkTable  >=J.U_AE " & _
    "where J.DocId = '" & JobDocID & "'"

Select Case Index
    Case PScreen, PWindows
        Set GRs = GCn.Execute("SELECT Vehicle_Type.Warr_Type FROM Model LEFT JOIN Vehicle_Type ON Model.Vehicle_Type = Vehicle_Type.Vehicle_Type where Model.Model ='" & Txt(Model) & "'")
        If UCase(left(PubComp_Name, 3)) = "LMP" Then
            mRepName = "JobCardSiebel"
            Set Rst = GCn.Execute(GSQL)
            CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
            If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
            Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
            rpt.Database.SetDataSource Rst
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
                        rpt.FormulaFields(I).TEXT = "'" & "Spare Purchase Bill" & "'"
                End Select
            Next
            rpt.ReadRecords
            'rpt.PrintOut
            
            Call Report_View(rpt, Me.CAPTION, 0, True)
            
            Set Rst = Nothing
            Exit Sub
        Else
            mRepName = IIf(VNull(GRs!warr_type) = 0, "JobCardCVD", "JobCardPCD")
        End If
        If GRs.RecordCount > 0 And VNull(GRs!warr_type) = 0 Then
            Call WindowsPrintCVD(Index, GSQL, PrePrn)
        Else
            Call WindowsPrintPCD(Index, GSQL, PrePrn)
        End If
        FrmPrn.Visible = False
    Case PDos
        Set GRs = GCn.Execute("SELECT Vehicle_Type.Warr_Type FROM Model LEFT JOIN Vehicle_Type ON Model.Vehicle_Type = Vehicle_Type.Vehicle_Type where Model.Model ='" & Txt(Model) & "'")
        If GRs.RecordCount > 0 Then
            If GRs!warr_type = 0 Then
                If PrePrn = True Then 'And GCn.Execute("select Dealer_ID from Syctrl").Fields(0).Value = "000001" Then  'for Katak
                    If UCase(left(PubComp_Name, 6)) = "RASHMI" Then
                        Call SpeedPrintCVD000001Rashmi(PrePrn)
                    ElseIf UCase(left(PubComp_Name, 7)) = "SOCIETY" Then
                        Call SpeedPrintCVD000001Society(PrePrn)
                    Else
                        Call SpeedPrintCVD000001(PrePrn)
                    End If
                Else
                    Call SpeedPrintCVD(GSQL, PrePrn)
                End If
            Else
                If PrePrn = True Then 'And GCn.Execute("select Dealer_ID from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value = "000001" Then  'for Katak
                    Call SpeedPrintPCD000001(PrePrn)
                Else
                    Call SpeedPrintCVD(GSQL, PrePrn)
                    'Call SpeedPrintPCD(PrePrn)
                End If
            End If
        Else
                If PrePrn = True Then 'And GCn.Execute("select Dealer_ID from Syctrl").Fields(0).Value = "000001" Then  'for Katak
                    If UCase(left(PubComp_Name, 6)) = "RASHMI" Then
                        Call SpeedPrintCVD000001Rashmi(PrePrn)
                    ElseIf UCase(left(PubComp_Name, 7)) = "SOCIETY" Then
                        Call SpeedPrintCVD000001Society(PrePrn)
                    Else
                        Call SpeedPrintCVD000001(PrePrn)
                    End If
                Else
                    Call SpeedPrintCVD(GSQL, PrePrn)
                End If
        End If
        FrmPrn.Visible = False
    Case PSetUp
        Set GRs = GCn.Execute("SELECT Vehicle_Type.Warr_Type FROM Model LEFT JOIN Vehicle_Type ON Model.Vehicle_Type = Vehicle_Type.Vehicle_Type where Model.Model ='" & Txt(Model) & "'")
        mRepName = IIf(GRs!warr_type = 0, "JobCardCVD", "JobCardPCD")
        Call PrinerSetUp
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
'TopCtrl1_eAdd
If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
End If
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub WindowsPrintPCD(Index As Integer, mQry$, PrePrn As Boolean)
Dim RepTitle$
Dim Condstr$
Dim Cnt As Integer, j As Integer
Dim RST1 As ADODB.Recordset
Dim Speciality$
Dim Rst As ADODB.Recordset
Dim I As Integer
Dim ColorName$

On Error GoTo ERRORHANDLER
Dim GRs As ADODB.Recordset
Set GRs = GCn.Execute("SELECT ColMast.Col_Desc FROM HisCard LEFT JOIN ColMast ON HisCard.ColourCode = ColMast.Col_Code where Hiscard.chassis = '" & Txt(Chassis) & "'")
If GRs.RecordCount > 0 And GRs.EOF = False And GRs.EOF = False Then
    ColorName = IIf(IsNull(GRs!Col_Desc), "", GRs!Col_Desc)
Else
    ColorName = ""
End If
Set GRs = Nothing

Set Rst = New ADODB.Recordset
With Rst
    .Fields.Append "Col1", adVarChar, 3, adFldIsNullable
    .Fields.Append "Col2", adVarChar, 40, adFldIsNullable
    .Fields.Append "Col3", adVarChar, 40, adFldIsNullable
    .Fields.Append "Col4", adVarChar, 1, adFldIsNullable
    .Fields.Append "Col5", adVarChar, 40, adFldIsNullable
    .Fields.Append "Col6", adVarChar, 40, adFldIsNullable
    .Fields.Append "Col7", adVarChar, 40, adFldIsNullable
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
End With
Cnt = 1
For I = 1 To FGrid1.Rows - 1
With Rst
        .AddNew
        .Fields("Col1") = I
        .Fields("Col2") = FGrid1.TextMatrix(I, Col_Trouble)
        Select Case I
            Case 1
                If PrePrn = False Then
                    .Fields("Col3") = "Labour"
                    .Fields("Col4") = "B"
                    .Fields("Col5") = "Parts"
                    .Fields("Col6") = "Consumables"
                    .Fields("Col7") = "Total"
                End If
            Case 2
                .Fields("Col3") = Format(Txt(EstLab), "0.00")
                .Fields("Col5") = Format(Txt(EstSpr), "0.00")
                .Fields("Col7") = Format(Val(Txt(EstLab)) + Val(Txt(EstSpr)), "0.00")
            Case 3
                If PrePrn = False Then
                    .Fields("Col3") = "Delivery"
                    .Fields("Col4") = "B"
                End If
            Case 4
                If PrePrn = False Then
                    .Fields("Col3") = "Date"
                    .Fields("Col4") = "B"
                    .Fields("Col5") = "Estimated Time"
                    .Fields("Col7") = "Time Out"
                End If
            Case 5
                .Fields("Col3") = Txt(DelDt)
                .Fields("Col5") = Txt(DelTime)
            Case 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
                .Fields("Col4") = "A"
            Case 21
                If PrePrn = False Then
                    .Fields("Col3") = "Inventory Of Accessory"
                    .Fields("Col4") = "B"
                End If
            Case 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32
                .Fields("Col3") = FGrid2.TextMatrix(Cnt, ElName)
                .Fields("Col5") = FGrid2.TextMatrix(Cnt, ElValue)
                Cnt = Cnt + 1
                .Fields("Col6") = FGrid2.TextMatrix(Cnt, ElName)
                .Fields("Col7") = FGrid2.TextMatrix(Cnt, ElValue)
                Cnt = Cnt + 1
        End Select
        .Update
End With
Next
If FGrid1.Rows < 33 Then
Do Until I > 32
    With Rst
        .AddNew
        Select Case I
        Case 1
                If PrePrn = False Then
                    .Fields("Col3") = "Labour"
                    .Fields("Col4") = "B"
                    .Fields("Col5") = "Parts"
                    .Fields("Col6") = "Consumables"
                    .Fields("Col7") = "Total"
                End If
            Case 2
                .Fields("Col3") = Format(Txt(EstLab), "0.00")
                .Fields("Col5") = Format(Txt(EstSpr), "0.00")
                .Fields("Col7") = Format(Val(Txt(EstLab)) + Val(Txt(EstSpr)), "0.00")
            Case 3
                If PrePrn = False Then
                    .Fields("Col3") = "Delivery"
                    .Fields("Col4") = "B"
                End If
            Case 4
                If PrePrn = False Then
                    .Fields("Col3") = "Date"
                    .Fields("Col4") = "B"
                    .Fields("Col5") = "Estimated Time"
                    .Fields("Col7") = "Time Out"
                End If
            Case 5
                .Fields("Col3") = Txt(DelDt)
                .Fields("Col5") = Txt(DelTime)
            Case 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
                .Fields("Col4") = "A"
            Case 21
                If PrePrn = False Then
                    .Fields("Col3") = "Inventory Of Accessory"
                    .Fields("Col4") = "B"
                End If
            Case 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32
'                .Fields("Col3") = FGrid2.TextMatrix(cnt, ElName)
'                .Fields("Col5") = FGrid2.TextMatrix(cnt, ElValue)
'                cnt = cnt + 1
'                .Fields("Col6") = FGrid2.TextMatrix(cnt, ElName)
'                .Fields("Col7") = FGrid2.TextMatrix(cnt, ElValue)
'                cnt = cnt + 1
        End Select
        .Update
    End With
    I = I + 1
Loop
End If

RepTitle = GCn.Execute("Select Div_SName from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
Speciality = GCn.Execute("Select W_SecSpeciality from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")

Set RST1 = New Recordset
RST1.CursorLocation = adUseClient
RST1.Open "select W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'", GCn, adOpenDynamic, adLockOptimistic

For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("SubTitle")
            rpt.FormulaFields(I).TEXT = "'" & Speciality & "'"
        Case UCase("Phone")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecPhone & "'"
        Case UCase("Fax")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecFax & "'"
        Case UCase("JobNo")
            rpt.FormulaFields(I).TEXT = "'" & PrinID(lblDocId.CAPTION) & "'"
        Case UCase("Name")
            rpt.FormulaFields(I).TEXT = "'" & Txt(OwnerName) & "'"
        Case UCase("Add1")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Address1) & " " & Txt(Address2) & "'"
        Case UCase("Add2")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Address3) & " " & Txt(City) & "'"
        Case UCase("PhoneCust")
            rpt.FormulaFields(I).TEXT = "'" & Txt(PhoneOff) & " " & Txt(PhoneResi) & Txt(Mobile) & "'"
'       Case UCase("JobNo")
'           rpt.FormulaFields(i).Text = "'" & left(JobDocID, 1) & Mid(JobDocID, 2, 1) & Txt(JobNo) & "'"
        Case UCase("JobDate")
            rpt.FormulaFields(I).TEXT = "'" & Txt(JobDt) & "'"
        Case UCase("InTime")
            rpt.FormulaFields(I).TEXT = "'" & Txt(ArrTime) & "'"
        Case UCase("Kms")
            rpt.FormulaFields(I).TEXT = "'" & Txt(CurrentKMS) & "'"
        Case UCase("Fuel")
            rpt.FormulaFields(I).TEXT = "'" & Txt(FUEL) & "'"
        Case UCase("Chassis")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Chassis) & "'"
        Case UCase("Engine")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Engine) & "'"
        Case UCase("RegNo")
            rpt.FormulaFields(I).TEXT = "'" & Txt(VehRegNo) & "'"
        Case UCase("Model")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Model) & "'"
        Case UCase("ColorName")
            rpt.FormulaFields(I).TEXT = "'" & ColorName & "'"
        Case UCase("DOS")
            rpt.FormulaFields(I).TEXT = "'" & Txt(SaleDate) & "'"
        Case UCase("Dealer")
            rpt.FormulaFields(I).TEXT = "'" & Txt(DNAME) & "'"
        Case UCase("SerAdv")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Supervisor) & "'"
        Case UCase("Mechanic")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Mechanic) & "'"
        Case UCase("Ele23")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(23, ElName) & "'"
        Case UCase("EleVal23")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(23, ElValue) & "'"
        Case UCase("Ele24")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(24, ElName) & "'"
        Case UCase("EleVal24")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(24, ElValue) & "'"
        Case UCase("Ele25")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(25, ElName) & "'"
        Case UCase("EleVal25")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(25, ElValue) & "'"
        Case UCase("Ele26")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(26, ElName) & "'"
        Case UCase("EleVal26")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(26, ElValue) & "'"
        Case UCase("Ele27")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(27, ElName) & "'"
        Case UCase("EleVal27")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(27, ElValue) & "'"
        Case UCase("Ele28")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(28, ElName) & "'"
        Case UCase("EleVal28")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(28, ElValue) & "'"
        Case UCase("Ele29")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(29, ElName) & "'"
        Case UCase("EleVal29")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(29, ElValue) & "'"
        Case UCase("Ele30")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(30, ElName) & "'"
        Case UCase("EleVal30")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(30, ElValue) & "'"
        Case UCase("Ele31")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(31, ElName) & "'"
        Case UCase("EleVal31")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(31, ElValue) & "'"
        Case UCase("Ele32")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(32, ElName) & "'"
        Case UCase("EleVal32")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(32, ElValue) & "'"
        Case UCase("Ele33")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(33, ElName) & "'"
        Case UCase("EleVal33")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(33, ElValue) & "'"
        Case UCase("Ele34")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(34, ElName) & "'"
        Case UCase("EleVal34")
            rpt.FormulaFields(I).TEXT = "'" & FGrid2.TextMatrix(34, ElValue) & "'"
    End Select
Next
           
rpt.Database.SetDataSource Rst
rpt.ReadRecords
Select Case Index
    Case PWindows  'Printer
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
            Case UCase("Title")
                rpt.FormulaFields(I).TEXT = "'" & "Job Card" & "'"
        End Select
        Next
        rpt.PrintOut False
    Case PDos
        Call Report_View(rpt, "Job Card" & IIf(IsNull(RepTitle), "", " (" & RepTitle & ")"), 1, True)
    Case PScreen  'screen
            Call Report_View(rpt, "Job Card" & IIf(IsNull(RepTitle), "", " (" & RepTitle & ")"), , True)
End Select
CmdPrint(PSetUp).Tag = ""
Set Rst = Nothing
Set RST1 = Nothing
Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub WindowsPrintPCDNew(Index As Integer, mQry$, PrePrn As Boolean)
Dim RepTitle$
Dim Condstr$
Dim Cnt As Integer, j As Integer
Dim RST1 As ADODB.Recordset
Dim Speciality$
Dim Rst As ADODB.Recordset
Dim I As Integer, l As Integer, TotalTroub As Integer
Dim ColorName$, PhoneOff$, PhoneResi$, Mobile$, Mail_ID$, Varient$, VehDet$, EvalStr$
Dim ExtendWar As Integer
On Error GoTo ERRORHANDLER
Dim GRs As ADODB.Recordset
Set GRs = GCn.Execute("SELECT ColMast.Col_Desc,HisCard.PhoneOff,HisCard.PhoneResi,HisCard.Mobile,HisCard.Mail_ID,HisCard.Varient,HisCard.VehDet,HisCard.ExtendWar FROM HisCard LEFT JOIN ColMast ON HisCard.ColourCode = ColMast.Col_Code where Hiscard.chassis = '" & Txt(Chassis) & "'")
If GRs.RecordCount > 0 And GRs.EOF = False And GRs.EOF = False Then
    ColorName = IIf(IsNull(GRs!Col_Desc), "", GRs!Col_Desc)
    PhoneOff = XNull(GRs!PhoneOff)
    PhoneResi = XNull(GRs!PhoneResi)
    Mobile = XNull(GRs!Mobile)
    Mail_ID = XNull(GRs!Mail_ID)
    Varient = XNull(GRs!Varient)
    VehDet = XNull(GRs!VehDet)
    ExtendWar = VNull(GRs!ExtendWar)
End If
Set GRs = Nothing
Set Rst = New ADODB.Recordset
With Rst
    .Fields.Append "Col1", adVarChar, 3, adFldIsNullable
    .Fields.Append "Col2", adVarChar, 40, adFldIsNullable
    .Fields.Append "Col3", adVarChar, 40, adFldIsNullable
    .Fields.Append "Col4", adVarChar, 1, adFldIsNullable
    .Fields.Append "Col5", adVarChar, 40, adFldIsNullable
    .Fields.Append "Col6", adVarChar, 40, adFldIsNullable
    .Fields.Append "Col7", adVarChar, 40, adFldIsNullable
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
End With
Cnt = 1
TotalTroub = FGrid1.Rows - 1
For I = 1 To TotalTroub
With Rst
        .AddNew
        .Fields("Col1") = I
        .Fields("Col2") = FGrid1.TextMatrix(I, Col_Trouble)
        .Update
End With
Next
If TotalTroub < 18 Then
    For I = TotalTroub To 18
        With Rst
            .AddNew
            .Fields("Col1") = I
            .Fields("Col2") = ""
            .Update
        End With
    Next
End If
RepTitle = GCn.Execute("Select Div_SName from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
Speciality = GCn.Execute("Select W_SecSpeciality from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
Set RST1 = New Recordset
RST1.CursorLocation = adUseClient
RST1.Open "select W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("SubTitle")
            rpt.FormulaFields(I).TEXT = "'" & Speciality & "'"
        Case UCase("Phone")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecPhone & "'"
        Case UCase("Fax")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecFax & "'"
'        Case UCase("JobNo")
'            rpt.FormulaFields(I).TEXT = "'" & PrinID(lblDocId.CAPTION) & "'"
        Case UCase("Name")
            rpt.FormulaFields(I).TEXT = "'" & Txt(OwnerName) & "'"
        Case UCase("Add1")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Address1) & " " & Txt(Address2) & "'"
        Case UCase("Add2")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Address3) & " " & Txt(City) & "'"
'        Case UCase("PhoneCust")
'            rpt.FormulaFields(I).TEXT = "'" & Txt(PhoneOff) & " " & Txt(PhoneResi) & Txt(Mobile) & "'"
       Case UCase("JobNo")
           rpt.FormulaFields(I).TEXT = "'" & PrinID(JobDocID) & "'"
        Case UCase("JobDate")
            rpt.FormulaFields(I).TEXT = "'" & Txt(JobDt) & "'"
        Case UCase("InTime")
            rpt.FormulaFields(I).TEXT = "'" & Txt(ArrTime) & "'"
        Case UCase("Kms")
            rpt.FormulaFields(I).TEXT = "'" & Txt(CurrentKMS) & "'"
        Case UCase("Fuel")
            rpt.FormulaFields(I).TEXT = "'" & Txt(FUEL) & "'"
        Case UCase("Chassis")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Chassis) & "'"
        Case UCase("Engine")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Engine) & "'"
        Case UCase("RegNo")
            rpt.FormulaFields(I).TEXT = "'" & Txt(VehRegNo) & "'"
        Case UCase("Model")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Model) & "'"
        Case UCase("ColorName")
            rpt.FormulaFields(I).TEXT = "'" & ColorName & "'"
        Case UCase("DOS")
            rpt.FormulaFields(I).TEXT = "'" & Txt(SaleDate) & "'"
        Case UCase("Dealer")
            rpt.FormulaFields(I).TEXT = "'" & Txt(DNAME) & "'"
        Case UCase("SerAdv")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Supervisor) & "'"
        Case UCase("Mechanic")
            rpt.FormulaFields(I).TEXT = "'" & Txt(Mechanic) & "'"
            
'
        Case UCase("EMail")
            rpt.FormulaFields(I).TEXT = "'" & Mail_ID & "'"
        Case UCase("Mobile")
            rpt.FormulaFields(I).TEXT = "'" & Mobile & "'"
        Case UCase("PhoneOff")
            rpt.FormulaFields(I).TEXT = "'" & PhoneOff & "'"
        Case UCase("PhoneRes")
            rpt.FormulaFields(I).TEXT = "'" & PhoneResi & "'"
        Case UCase("Varient")
            rpt.FormulaFields(I).TEXT = "'" & Varient & "'"
        Case UCase("VehDet")
            rpt.FormulaFields(I).TEXT = "'" & VehDet & "'"
        Case UCase("EstLab+Parts")
            rpt.FormulaFields(I).TEXT = "" & Val(Txt(EstSpr)) + Val(Txt(EstLab)) & ""
        Case UCase("ExtendWar")
            rpt.FormulaFields(I).TEXT = "'" & IIf(ExtendWar = 1, "True", "False") & "'"
        Case UCase("LastAttOn")
            rpt.FormulaFields(I).TEXT = "'" & Txt(LastJobDt) & "'"
        Case UCase("LastAttFor")
            rpt.FormulaFields(I).TEXT = "'" & Txt(LastSrv) & "'"
        Case UCase("EstDelTime")
            rpt.FormulaFields(I).TEXT = "'" & Txt(DelTime) & "'"
        Case UCase("EstdelDate")
            rpt.FormulaFields(I).TEXT = "'" & Txt(DelDt) & "'"
        
        Case UCase("Ele1")
            For l = 1 To 34
                If FGrid2.TextMatrix(l, ElValue) = "Yes" Then
                    EvalStr = EvalStr & FGrid2.TextMatrix(l, ElName) & " , "
                End If
            Next
        rpt.FormulaFields(I).TEXT = "'" & EvalStr & "'"
    End Select
Next
rpt.Database.SetDataSource Rst
rpt.ReadRecords
Select Case Index
    Case PWindows  'Printer
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
            Case UCase("Title")
                rpt.FormulaFields(I).TEXT = "'" & "Job Card" & "'"
        End Select
        Next
        rpt.PrintOut False
    Case PDos
        Call Report_View(rpt, "Job Card" & IIf(IsNull(RepTitle), "", " (" & RepTitle & ")"), 1, True)
    Case PScreen  'screen
            Call Report_View(rpt, "Job Card" & IIf(IsNull(RepTitle), "", " (" & RepTitle & ")"), , True)
End Select
CmdPrint(PSetUp).Tag = ""
Set Rst = Nothing
Set RST1 = Nothing
Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub PrinerSetUp()
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.PrinterSetup (0)
CmdPrint(PSetUp).Tag = "1"
LblPrinter.CAPTION = rpt.PrinterName
End Sub

Private Sub SpeedPrintPCD(PrePrn As Boolean)
On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, j As Integer, ColorName$, Cnt As Integer
    Dim PrintStr$
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, RepTitle$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double, FormulaStr1$, FormulaStr2$
    Dim fob As New FileSystemObject, SecondStr As Boolean
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, mQry$
'    Dim GRs As ADODB.Recordset
    Set GRs = GCn.Execute("SELECT ColMast.Col_Desc FROM HisCard LEFT JOIN ColMast ON HisCard.ColourCode = ColMast.Col_Code where Hiscard.chassis = '" & Txt(Chassis) & "'")
    If GRs.RecordCount > 0 And GRs.EOF = False And GRs.EOF = False Then
        ColorName = IIf(IsNull(GRs!Col_Desc), "", GRs!Col_Desc)
    Else
        ColorName = ""
    End If
    Set GRs = Nothing
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select JobCardFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
 
    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
        
    mHeader = 14
    mFooter = 20
    mFooter = mFooter + FooterCnt
      
    'Header
    RepTitle = GCn.Execute("Select Div_SName from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
    mDocStr = "Job Card" & IIf(RepTitle = "", "", " (" & RepTitle & ")")
    Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
        
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!W_SecSpeciality) <> "" Then
       Print #1, PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
       mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    If PubComp_Add2 <> "" Then
       Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
       mHeader = mHeader + 1
    End If
    If PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT("PHONE No :" & XNull(RstCompDet!W_SecPhone) & IIf(XNull(RstCompDet!W_SecFax) = "", "", "FAX : " & XNull(RstCompDet!W_SecFax)), "C", PageWidth)
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, PRN_TIT("** " & mDocStr & " **", "A", PageWidth) & mChr18 & mEmph
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, PSTR("Owner Name & Address : ", 23) & PSTR("Job Card No. : ", 18) & mID(JobDocID, 1, 2) & PSTR(Txt(JobNo), 8, , AlignRight) & mEmph1 & Space(1) & mChr17 & PSTR(" Chassis No.", 12) & " : " & Txt(Chassis)
    mHeader = mHeader + 1
    Print #1, PSTR(Txt(OwnerName), 40) & mChr18 & mEmph & PSTR("Job Card Date : ", 18) & PSTR(Txt(JobDt), 11) & mEmph1 & Space(1) & mChr17 & PSTR("Engine No.", 12) & " : " & Txt(Engine)
    mHeader = mHeader + 1
    Print #1, PSTR(Txt(Address1), 40) & mChr18 & PSTR("Incoming Time : ", 18) & PSTR(Txt(ArrTime), 11) & mEmph1 & Space(1) & mChr17 & PSTR("Reg. No.", 12) & " : " & Txt(VehRegNo)
    mHeader = mHeader + 1
    Print #1, PSTR(Txt(Address2), 40) & mChr18 & Space(29) & mEmph1 & Space(1) & mChr17 & PSTR("Model", 12) & " : " & Txt(Model)
    mHeader = mHeader + 1
    Print #1, PSTR(Txt(Address3) & Txt(City), 40) & mChr18 & Space(29) & Space(1) & mChr17 & PSTR("Color", 12) & " : " & ColorName
    mHeader = mHeader + 1
    Print #1, PSTR("Ph." & Txt(PhoneOff) & Txt(PhoneResi) & Txt(Mobile), 40) & mChr18 & mEmph & PSTR("Mileage in ", 14) & PSTR(" Fuel Reading", 15) & mEmph1 & Space(1) & mChr17 & PSTR("  Sale Date", 12) & " : " & Txt(SaleDate)
    mHeader = mHeader + 1
    Print #1, Space(40) & mChr18 & PSTR(Txt(CurrentKMS), 14) & PSTR(Txt(FUEL), 15) & Space(1) & mChr17 & PSTR("Dealer Name", 12) & " : " & Txt(DNAME) & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-") & mEmph
    mHeader = mHeader + 1
    Print #1, PSTR("Srl", 3) & " | " & PSTR("Customer Complaint's/Operation", 40) & " | " & "Estimate Cost"
    mHeader = mHeader + 1
    Print #1, PSTR("No.", 3) & " | " & PSTR("Labour Description", 40) & " | " & mEmph1
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    mFix = PageLength - (mHeader + mFooter)
    Page = 1
    mLine = 1
    mSlNo = 1
    Cnt = 1
    If FGrid1.Rows > 0 Then
        I = 1
        Do Until I = FGrid1.Rows
            If I = 33 Then Exit Do
            PrintStr = mChr17 & PSTR(I, 5) & mChr18 & " | " & mChr17 & PSTR(FGrid1.TextMatrix(I, Col_Trouble), 68) & mChr18 & " | " & mChr17
            Select Case I
                Case 1
                    If PrePrn = False Then
                        PrintStr = PrintStr & PSTR("Labour", 12) & PSTR("Parts", 12) & PSTR("Consumables", 12) & PSTR("Total", 12)
                    End If
                Case 2
                    PrintStr = PrintStr & PSTR(Format(Txt(EstLab), "0.00"), 12, AlignRight) & PSTR(Format(Txt(EstSpr), "0.00"), 12) & Space(12) & PSTR(Format(Val(Txt(EstLab)) + Val(Txt(EstSpr)), "0.00"), 12)
                Case 3
                    If PrePrn = False Then
                        PrintStr = PrintStr & mChr18 & mEmph & "Estimated Delivery " & mEmph1 & mChr17
                    End If
                Case 4
                    If PrePrn = False Then
                        PrintStr = PrintStr & PSTR("Date", 12) & PSTR("Time", 12) & Space(12) & PSTR("Time Out", 12)
                    End If
                Case 5
                    PrintStr = PrintStr & PSTR(Txt(DelDt), 12) & PSTR(Txt(DelTime), 12)
                Case 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
                    PrintStr = PrintStr
                Case 21
                    If PrePrn = False Then
                        PrintStr = PrintStr & mChr18 & mEmph & "Inventory Of Accessory" & mEmph1 & mChr17
                    End If
                Case 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32
                    PrintStr = PrintStr & PSTR(FGrid2.TextMatrix(Cnt, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt, ElValue), 7, , AlignRight) & Space(1) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElValue), 7, , AlignRight)
                    Cnt = Cnt + 2
            End Select
            Print #1, PrintStr
            I = I + 1
            mSlNo = mSlNo + 1
            mLine = mLine + 1
        Loop
        If FGrid1.Rows < 33 Then
            Do Until I > 32
                PrintStr = mChr17 & Space(5) & mChr18 & " | " & mChr17 & Space(68) & mChr18 & " | " & mChr17
                Select Case I
                    Case 1
                        If PrePrn = False Then
                            PrintStr = PrintStr & PSTR("Labour", 12) & PSTR("Parts", 12) & PSTR("Consumables", 12) & PSTR("Total", 12)
                        End If
                    Case 2
                        PrintStr = PrintStr & PSTR(Format(Txt(EstLab), "0.00"), 12, AlignRight) & PSTR(Format(Txt(EstSpr), "0.00"), 12) & Space(12) & PSTR(Format(Val(Txt(EstLab)) + Val(Txt(EstSpr)), "0.00"), 12)
                    Case 3
                        If PrePrn = False Then
                            PrintStr = PrintStr & mChr18 & mEmph & "Estimated Delivery " & mEmph1 & mChr17
                        End If
                    Case 4
                        If PrePrn = False Then
                            PrintStr = PrintStr & PSTR("Date", 12) & PSTR("Time", 12) & Space(12) & PSTR("Time Out", 12)
                        End If
                    Case 5
                        PrintStr = PrintStr & PSTR(Txt(DelDt), 12) & PSTR(Txt(DelTime), 12)
                    Case 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
                        PrintStr = PrintStr
                    Case 21
                        If PrePrn = False Then
                            PrintStr = PrintStr & mChr18 & mEmph & "Inventory Of Accessory" & mEmph1 & mChr17
                        End If
                    Case 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32
                        If FGrid2.Rows > 1 And FGrid2.TextMatrix(1, ElName) <> "" Then
                            PrintStr = PrintStr & PSTR(FGrid2.TextMatrix(Cnt, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt, ElValue), 7, , AlignRight) & Space(1) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElValue), 7, , AlignRight)
                            Cnt = Cnt + 2
                        End If
                End Select
                Print #1, PrintStr
                I = I + 1
                mSlNo = mSlNo + 1
                mLine = mLine + 1
            Loop
        End If
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop
    
    'FOOTER
    Cnt = 23
    If FGrid2.Rows > 1 And FGrid2.TextMatrix(1, ElName) <> "" Then
        Print #1, mChr17 & PSTR("I Hereby authorise the above work to be done alongwith necessary spares.", 78) & mChr18 & " | " & mChr17 & PSTR(FGrid2.TextMatrix(Cnt, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt, ElValue), 7, , AlignRight) & Space(1) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElValue), 7, , AlignRight)
        Cnt = Cnt + 2
        Print #1, PSTR("materials at my cost.Any additional work if required shall be done at my cost.", 78) & mChr18 & " | " & mChr17 & PSTR(FGrid2.TextMatrix(Cnt, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt, ElValue), 7, , AlignRight) & Space(1) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElValue), 7, , AlignRight)
        Cnt = Cnt + 2
        Print #1, PSTR("also authorised the vehicle to be store,repaired & driven at my risk.", 78) & mChr18 & " | " & mChr17 & PSTR(FGrid2.TextMatrix(Cnt, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt, ElValue), 7, , AlignRight) & Space(1) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElValue), 7, , AlignRight)
        Cnt = Cnt + 2
        Print #1, PSTR("Invoice for above job has to be settled" & " (* Delete as necessary)", 78) & mChr18 & " | " & mChr17 & PSTR(FGrid2.TextMatrix(Cnt, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt, ElValue), 7, , AlignRight) & Space(1) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElValue), 7, , AlignRight)
        Cnt = Cnt + 2
        Print #1, PSTR("Delivery of the vehicle.", 78) & mChr18 & " | " & mChr17 & PSTR(FGrid2.TextMatrix(Cnt, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt, ElValue), 7, , AlignRight) & Space(1) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElValue), 7, , AlignRight)
        Cnt = Cnt + 2
        Print #1, PSTR("I agree to the above work being undertaken" & " I wish you to proceed  without", 78) & mChr18 & " | " & mChr17 & PSTR(FGrid2.TextMatrix(Cnt, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt, ElValue), 7, , AlignRight) & Space(1) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElName), 18) & PSTR(FGrid2.TextMatrix(Cnt + 1, ElValue), 7, , AlignRight)
    End If
    Print #1, Space(35) & PSTR("Further autority", 56) & mChr18 & " | " & mChr17
    Print #1, Space(78) & mChr18 & " | " & mChr17
    Print #1, PSTR("Signature of Customer(or Agent)" & " * Await my phone/written authority", 78) & mChr18 & " | " & mChr17
    Print #1, PSTR("Service Advisor Name    : " & Txt(Supervisor), 78) & mChr18 & " | " & mChr17
    Print #1, PSTR("Signature               : ", 78) & mChr18 & " | " & mChr17
    Print #1, PSTR("Technical Inspection by : " & Txt(Mechanic), 78) & mChr18 & " | " & mChr17
    Print #1, PSTR("Signature               : ", 78) & mChr18 & " | "
    Print #1, Replace(Space(PageWidth), " ", "-")
 
    Print #1, mChr17 & "" & Space((PageWidth * 1.7) - Len("" & pubUName & "   " & PubServerDate)) & pubUName & "   " & PubServerDate & mChr18
    Print #1, mEject
                If FGrid1.Rows > 33 Then
                    I = 33
                    'Header On Second Page
                    mHeader = 0
'                    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
'                    mHeader = mHeader + 1
'                    If XNull(RstCompDet!W_SecSpeciality) <> "" Then
'                       Print #1, PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
'                       mHeader = mHeader + 1
'                    End If
'                    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
'                    If PubComp_Add2 <> "" Then
'                       Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
'                       mHeader = mHeader + 1
'                    End If
'                    If PubComp_City <> "" Then
'                        Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
'                        mHeader = mHeader + 1
'                    End If
'                    Print #1, PRN_TIT("PHONE No :" & XNull(RstCompDet!W_SecPhone) & IIf(XNull(RstCompDet!W_SecFax) = "", "", "FAX : " & XNull(RstCompDet!W_SecFax)), "C", PageWidth)
'                    mHeader = mHeader + 1
'                    Print #1, ""
'                    mHeader = mHeader + 1
'                    Print #1, PRN_TIT("** " & mDocStr & " **", "A", PageWidth) & mChr18 & mEmph
'                    mHeader = mHeader + 1
'                    Print #1, ""
'                    mHeader = mHeader + 1
                    Print #1, Space(23) & PSTR("Job Card No. : ", 18) & mID(JobDocID, 1, 2) & PSTR(Txt(JobNo), 8, , AlignRight) & mEmph1 & Space(1) & mChr17 & PSTR(" Chassis No.", 12) & " : " & Txt(Chassis) & mChr18
                    mHeader = mHeader + 1
'                    Print #1, PSTR(txt(OwnerName), 40) & mChr18 & mEmph & PSTR("Job Card Date : ", 18) & PSTR(txt(JobDt), 11) & mEmph1 & Space(1) & mChr17 & PSTR("Engine No.", 12) & " : " & txt(Engine)
'                    mHeader = mHeader + 1
'                    Print #1, PSTR(txt(Address1), 40) & mChr18 & PSTR("Incoming Time : ", 18) & PSTR(txt(ArrTime), 11) & mEmph1 & Space(1) & mChr17 & PSTR("Reg. No.", 12) & " : " & txt(VehRegNo)
'                    mHeader = mHeader + 1
'                    Print #1, PSTR(txt(Address2), 40) & mChr18 & Space(29) & mEmph1 & Space(1) & mChr17 & PSTR("Model", 12) & " : " & txt(Model)
'                    mHeader = mHeader + 1
'                    Print #1, PSTR(txt(Address3) & txt(City), 40) & mChr18 & Space(29) & Space(1) & mChr17 & PSTR("Color", 12) & " : " & ColorName
'                    mHeader = mHeader + 1
'                    Print #1, PSTR("Ph." & txt(PhoneOff) & txt(PhoneResi) & txt(Mobile), 40) & mChr18 & mEmph & PSTR("Mileage in ", 14) & PSTR(" Fuel Reading", 15) & mEmph1 & Space(1) & mChr17 & PSTR("  Sale Date", 12) & " : " & txt(DelDt)
'                    mHeader = mHeader + 1
'                    Print #1, Space(40) & mChr18 & PSTR(txt(CurrentKMS), 14) & PSTR(txt(FUEL), 15) & Space(1) & mChr17 & PSTR("Dealer Name", 12) & " : " & txt(DNAME) & mChr18
'                    mHeader = mHeader + 1
'                    mHeader = mHeader + 1
                    Print #1, Replace(Space(PageWidth), " ", "-") & mEmph
                    mHeader = mHeader + 1
                    Print #1, PSTR("Srl", 3) & " | " & PSTR("Customer Complaint's/Operation", 40) & " | " & "Estimate Cost"
                    mHeader = mHeader + 1
                    Print #1, PSTR("No.", 3) & " | " & PSTR("Labour Description", 40) & " | " & mEmph1
                    mHeader = mHeader + 1
                    Print #1, Replace(Space(PageWidth), " ", "-")
                    mHeader = mHeader + 1
                    Print #1, Space(PageWidth - Len("Contd. from last page.." + STR(Page))) & "Contd. from last page.." & STR(Page)
                    mHeader = mHeader + 1
                    Page = Page + 1
                   Do Until I = FGrid1.Rows
                        PrintStr = mChr17 & PSTR(I, 5) & mChr18 & " | " & mChr17 & PSTR(FGrid1.TextMatrix(I, Col_Trouble), 68) & mChr18 & " | " & mChr17
                         Print #1, PrintStr
                        I = I + 1
                   Loop
                   Print #1, mChr18 & Replace(Space(PageWidth), " ", "-")
                End If

    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''    Else
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
''    End If
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub
Private Sub SpeedPrintPCDNew(PrePrn As Boolean)
On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, j As Integer, ColorName$, Cnt As Integer
    Dim PrintStr$
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, RepTitle$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Double, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double, FormulaStr1$, FormulaStr2$
    Dim fob As New FileSystemObject, SecondStr As Boolean
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, mQry$
    Dim PhoneOff$, PhoneResi$, Mobile$, Mail_ID$, Varient$, VehDet$, EvalStr$, HlpLineNo$
    Dim ExtendWar As Integer, TotalRows As Integer, TotalRows1 As Integer, prnPos As Integer
    Dim multiPage As Boolean
'    On Error GoTo ERRORHANDLER
    
    Set GRs = GCn.Execute("SELECT ColMast.Col_Desc,HisCard.PhoneOff,HisCard.PhoneResi,HisCard.Mobile,HisCard.Mail_ID,HisCard.Varient,HisCard.VehDet,HisCard.ExtendWar FROM HisCard LEFT JOIN ColMast ON HisCard.ColourCode = ColMast.Col_Code where Hiscard.chassis = '" & Txt(Chassis) & "'")
    If GRs.RecordCount > 0 And GRs.EOF = False And GRs.EOF = False Then
        ColorName = IIf(IsNull(GRs!Col_Desc), "", GRs!Col_Desc)
        PhoneOff = XNull(GRs!PhoneOff)
        PhoneResi = XNull(GRs!PhoneResi)
        Mobile = XNull(GRs!Mobile)
        Mail_ID = XNull(GRs!Mail_ID)
        Varient = XNull(GRs!Varient)
        VehDet = XNull(GRs!VehDet)
        ExtendWar = VNull(GRs!ExtendWar)
        VehDet = XNull(GRs!VehDet)
    End If
    Set GRs = Nothing
    HlpLineNo = GCn.Execute("Select HelpLineNo from Syctrl").Fields(0).Value
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select JobCardFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
        
    mHeader = 14
    mFooter = 20
    mFooter = mFooter + FooterCnt
      
    'Header
    RepTitle = GCn.Execute("Select Div_SName from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
    mDocStr = "Job Card" & IIf(RepTitle = "", "", " (" & RepTitle & ")")
    Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
    Print #1, Chr(27) & Chr(69) & PSTR(PubComp_Name, PageWidth / 3) & Chr(27) & Chr(70) & Chr(27) & Chr(69) & PSTR("** " & mDocStr & " **", PageWidth / 3) & Chr(27) & Chr(70) & PSTR("Job Card No.  :", 15) & PrinID(JobDocID)
    mHeader = mHeader + 1
    Print #1, PSTR(PubComp_Add, PageWidth / 3) & Space(PageWidth / 3) & PSTR("Date          :", 15) & Txt(JobDt)
    mHeader = mHeader + 1
    Print #1, PSTR(PubComp_Add2, PageWidth / 3) & Space(PageWidth / 3) & PSTR("Veh. Reg. No. :", 15) & Txt(VehRegNo)
    mHeader = mHeader + 1
    Print #1, PSTR(PubComp_City, PageWidth / 3) & Space(PageWidth / 3) & PSTR("Current Kms   :", 15) & Txt(CurrentKMS)
    mHeader = mHeader + 1
    Print #1, "PHONE No :" & XNull(RstCompDet!W_SecPhone)
    mHeader = mHeader + 1
    Print #1, Chr(27) & Chr(45) & Chr(1) & Chr(27) & Chr(69) & PSTR("24 - Hrs Helpline No.  :" & HlpLineNo, 90) & Chr(27) & Chr(70) & Chr(27) & Chr(45) & Chr(0)
    mHeader = mHeader + 1
'    Print #1, Replace(Space(PageWidth), " ", "-") & mEmph
'    mHeader = mHeader + 1
    Print #1, Chr(27) & Chr(45) & Chr(1) & Chr(27) & Chr(69) & PSTR("CUSTOMER DETAILS" & Chr(27) & Chr(70) & Space(20) & Chr(27) & Chr(69) & "Vehicle Detail  " & Chr(27) & Chr(70) & IIf(VehDet = "Personal", Chr(251), "") & "Personal " & IIf(VehDet = "Taxi", Chr(251), "") & "Taxi " & IIf(VehDet = "Corporate", Chr(251), "") & "Corporate ", PageWidth) & Space(10) & Chr(27) & Chr(45) & Chr(0)
    mHeader = mHeader + 1
   
    Print #1, PSTR("Name : " & Txt(OwnerName), PageWidth / 3) & " | " & PSTR("Model  :" & Txt(Model), 25) & PSTR("Varient:" & Varient, 20) & PSTR("Colour: " & ColorName, 20)
    mHeader = mHeader + 1
    Print #1, PSTR("Address : " & Txt(Address1), PageWidth / 3) & " | " & PSTR("Sold By:" & Txt(DCode), 15) & PSTR("On: " & Txt(SaleDate), 15) & PSTR("Extended Warr.: " & IIf(ExtendWar = 1, "Yes", "No"), 25)
    mHeader = mHeader + 1
    Print #1, PSTR(Txt(Address2), PageWidth / 3) & " | " & PSTR("Chassis:" & Txt(Chassis), 30)
    mHeader = mHeader + 1
    Print #1, PSTR(Txt(Address3), PageWidth / 3) & " | " & PSTR("Engine :" & Txt(Engine), 30)
    mHeader = mHeader + 1
    Print #1, PSTR("Phone Res : ", PageWidth / 3) & " | " & PSTR("Last Attended At. : ", PageWidth / 3) & PSTR("On :" & Txt(LastJobDt), PageWidth / 3)
    mHeader = mHeader + 1
    Print #1, PSTR("Mobile : ", PageWidth / 3) & " | " & PSTR("For         : " & Txt(LastSrv), 30)
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("E-MailID : ", PageWidth + 10) & Chr(27) + Chr(45) + Chr(0)
    For I = 1 To 34
        EvalStr = ""
        For j = I To I + 3
            If FGrid2.TextMatrix(I, ElValue) = "Yes" Then
                EvalStr = EvalStr & FGrid2.TextMatrix(I, ElName) & ","
            End If
            I = I + 1
        Next
            If I = 5 Then
                Print #1, PSTR(EvalStr, 60) & " |Tyre make & |Other              "
            ElseIf I = 10 Then
                Print #1, PSTR(EvalStr, 60) & " |Condition   |Observations:      "
            Else
                Print #1, PSTR(EvalStr, 60) & " |            |"
            End If
            mHeader = mHeader + 1
    Next
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & Space(PageWidth + 10) & Chr(27) + Chr(45) + Chr(0)
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & Space(PageWidth / 3) & "|" & PSTR("Final Estimation", PageWidth / 3) & "|" & PSTR("Actual", (PageWidth / 3) - 3) & Space(10) & Chr(27) + Chr(45) + Chr(0)
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Charges(Lab.+Parts)", PageWidth / 3) & "|" & PSTR(Val(Txt(EstLab)) + Val(Txt(EstSpr)), PageWidth / 3, , AlignLeft) & "|" & Space((PageWidth / 3) - 3) & Space(10) & Chr(27) + Chr(45) + Chr(0)
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Time and Date", PageWidth / 3) & "|" & Space(PageWidth / 3) & "|" & Space((PageWidth / 3) - 3) & Space(10) & Chr(27) + Chr(45) + Chr(0)
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Customer Request", 40) & "|" & PSTR("Repair advice by service advisor", 40) & Space(9) & Chr(27) + Chr(45) + Chr(0)
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Comp Code", 10) & "|" & PSTR("Complaint Detail", 23) & "|" & PSTR("JobCode", 10) & "|" & PSTR("Repair Det.", 10) & "|" & PSTR("Std.Hrs", 7) & "|" & PSTR("Parts Est.", 7) & "|" & PSTR("Lab. Est.", 7) & Space(10) & Chr(27) + Chr(45) + Chr(0); ""
    mHeader = mHeader + 1
    TotalRows1 = FGrid1.Rows - 1
    TotalRows = TotalRows1
    If TotalRows1 > 18 Then TotalRows1 = 18: multiPage = True
    If TotalRows1 <= 18 Then
        For I = 1 To TotalRows1
            Print #1, PSTR(FGrid1.TextMatrix(I, Col_Code), 10) & "|" & PSTR(FGrid1.TextMatrix(I, Col_Trouble), 23) & "|" & Space(10) & "|" & Space(10) & "|" & Space(7) & "|"; Space(7) & "|"; Space(7) & ""
            mHeader = mHeader + 1
        Next
        For I = TotalRows1 To 18
            Print #1, Space(10) & "|" & Space(23) & "|" & Space(10) & "|" & Space(10) & "|" & Space(7) & "|"; Space(7) & "|"; Space(7) & ""
            mHeader = mHeader + 1
        Next
    End If
    Print #1, Chr(27) + Chr(45) + Chr(1) & Space(PageWidth) & Space(10) & Chr(27) + Chr(45) + Chr(0)
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Estimate Charged :", PageWidth / 3) & "|" & PSTR("Est.Del.Time :" & Txt(DelTime), PageWidth / 3) & "|" & PSTR("Est.Del.Date :" & Txt(DelDt), (PageWidth / 3) - 3) & Space(10) & Chr(27) + Chr(45) + Chr(0)
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Revision in the Job :", PageWidth) & Space(10) & Chr(27) + Chr(45) + Chr(0)
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Dscription", 16) & "|" & PSTR("Est.Parts", 10) & "|" & PSTR("Est.Lab", 10) & "|" & PSTR("Est.Time", 10) & "|" & PSTR("Cust. Permission for additional job", 40) & Chr(27) + Chr(45) + Chr(0); ""
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & Space(16) & "|" & Space(10) & "|" & Space(10) & "|" & Space(10); "|" & PSTR("Time of calling up", 30) & Space(10) & Chr(27) + Chr(45) + Chr(0); ""
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & Space(16) & "|" & Space(10) & "|" & Space(10) & "|" & Space(10); "|" & PSTR("Cust. Aggrees for add. work (Y/N)", 40) & Chr(27) + Chr(45) + Chr(0); ""
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Total", 16) & "|" & Space(10) & "|" & Space(10) & "|" & Space(10); "|" & Space(40) & Chr(27) + Chr(45) + Chr(0); ""
    mHeader = mHeader + 1
    Print #1, PSTR("I authorise to execute the jobs", 35) & "|" & PSTR("I Certify that the work has", 35) & "|" & PSTR("Payment Details", 25) & ""
    mHeader = mHeader + 1
    Print #1, PSTR("described herein using necessary", 35) & "|" & PSTR("been done to my satisfact-", 35) & "|" & Space(25) & ""
    mHeader = mHeader + 1
    Print #1, PSTR("material at my cost.I understand", 35) & "|" & PSTR("ion and i have taken the", 35) & "|" & "Cash   Card  Warr" & ""
    mHeader = mHeader + 1
    Print #1, PSTR("that the vehicle is being stored", 35) & "|" & PSTR("delivery of vehicle.", 35) & "|" & Space(25) & ""
    mHeader = mHeader + 1
    Print #1, PSTR("repaired and tested at my risk", 30) & "|" & Space(30) & "|" & PSTR("Invoice No.", 25) & ""
    mHeader = mHeader + 1
    Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Time            Customer Sign.", 30) & "|" & PSTR("Time            Customer Sign.", 30) & "|" & PSTR("Date", 18) & Space(10) & Chr(27) + Chr(45) + Chr(0); ""
    
    mHeader = mHeader + 1
    Print #1, "All deliveries against payment by cash or card only.Please fill the correct timings"
    mHeader = mHeader + 1
    Print #1, "above in order to help us serve you faster.THANK YOU"
    mHeader = mHeader + 1
    Print #1, mChr17 & "" & Space((PageWidth * 1.7) - Len("" & pubUName & "   " & PubServerDate)) & pubUName & "   " & PubServerDate & mChr18
    Print #1, mEject
    If multiPage Then
        TotalRows = TotalRows - 18
        prnPos = 18
        mHeader = 0
        While TotalRows > 0
            Print #1, Chr(27) & Chr(69) & PSTR(PubComp_Name, PageWidth / 3) & Chr(27) & Chr(70) & Chr(27) & Chr(69) & PSTR("** " & mDocStr & " **", PageWidth / 3) & Chr(27) & Chr(70) & PSTR("Job Card No.  :", 15) & PrinID(JobDocID)
            mHeader = mHeader + 1
            Print #1, PSTR(PubComp_Add, PageWidth / 3) & Space(PageWidth / 3) & PSTR("Date          :", 15) & Txt(JobDt)
            mHeader = mHeader + 1
            Print #1, PSTR(PubComp_Add2, PageWidth / 3) & Space(PageWidth / 3) & PSTR("Veh. Reg. No. :", 15) & Txt(VehRegNo)
            mHeader = mHeader + 1
            Print #1, PSTR(PubComp_City, PageWidth / 3) & Space(PageWidth / 3) & PSTR("Current Kms   :", 15) & Txt(CurrentKMS)
            mHeader = mHeader + 1
            Print #1, "PHONE No :" & XNull(RstCompDet!W_SecPhone)
            mHeader = mHeader + 1
            Print #1, Chr(27) & Chr(45) & Chr(1) & Chr(27) & Chr(69) & PSTR("24 - Hrs Helpline No.  :" & HlpLineNo, 90) & Chr(27) & Chr(70) & Chr(27) & Chr(45) & Chr(0)
            mHeader = mHeader + 1
        '    Print #1, Replace(Space(PageWidth), " ", "-") & mEmph
        '    mHeader = mHeader + 1
            Print #1, Chr(27) & Chr(45) & Chr(1) & PSTR("CUSTOMER DETAILS" & Space(25) & "Vehicle Detail      Personal      Taxi      Corporate ", PageWidth) & Space(10) & Chr(27) & Chr(45) & Chr(0)
            mHeader = mHeader + 1
           
            Print #1, PSTR("Name : " & Txt(OwnerName), PageWidth / 3) & " | " & PSTR("Model  :" & Txt(Model), 25) & PSTR("Varient:" & Varient, 20) & PSTR("Colour: " & ColorName, 20)
            mHeader = mHeader + 1
            Print #1, PSTR("Address : " & Txt(Address1), PageWidth / 3) & " | " & PSTR("Sold By:" & Txt(DCode), 15) & PSTR("On: " & Txt(SaleDate), 15) & PSTR("Extended Warr.: " & IIf(ExtendWar = 1, "Yes", "No"), 25)
            mHeader = mHeader + 1
            Print #1, PSTR(Txt(Address2), PageWidth / 3) & " | " & PSTR("Chassis:" & Txt(Chassis), 30)
            mHeader = mHeader + 1
            Print #1, PSTR(Txt(Address3), PageWidth / 3) & " | " & PSTR("Engine :" & Txt(Engine), 30)
            mHeader = mHeader + 1
            Print #1, PSTR("Phone Res : ", PageWidth / 3) & " | " & PSTR("Last Attended At. : ", PageWidth / 3) & PSTR("On :" & Txt(LastJobDt), PageWidth / 3)
            mHeader = mHeader + 1
            Print #1, PSTR("Mobile : ", PageWidth / 3) & " | " & PSTR("For         : " & Txt(LastSrv), 30)
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("E-MailID : ", PageWidth + 10) & Chr(27) + Chr(45) + Chr(0)
            For I = 1 To 34
                EvalStr = ""
                For j = I To I + 3
                    If FGrid2.TextMatrix(I, ElValue) = "Yes" Then
                        EvalStr = EvalStr & FGrid2.TextMatrix(I, ElName) & ","
                    End If
                    I = I + 1
                Next
                    If I = 5 Then
                        Print #1, PSTR(EvalStr, 60) & " |Tyre make & |Other              "
                    ElseIf I = 10 Then
                        Print #1, PSTR(EvalStr, 60) & " |Condition   |Observations:      "
                    Else
                        Print #1, PSTR(EvalStr, 60) & " |            |"
                    End If
                    mHeader = mHeader + 1
            Next
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & Space(PageWidth + 10) & Chr(27) + Chr(45) + Chr(0)
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & Space(PageWidth / 3) & "|" & PSTR("Final Estimation", PageWidth / 3) & "|" & PSTR("Actual", (PageWidth / 3) - 3) & Space(10) & Chr(27) + Chr(45) + Chr(0)
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Charges(Lab.+Parts)", PageWidth / 3) & "|" & PSTR(Val(Txt(EstLab)) + Val(Txt(EstSpr)), PageWidth / 3, , AlignLeft) & "|" & Space((PageWidth / 3) - 3) & Space(10) & Chr(27) + Chr(45) + Chr(0)
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Time and Date", PageWidth / 3) & "|" & Space(PageWidth / 3) & "|" & Space((PageWidth / 3) - 3) & Space(10) & Chr(27) + Chr(45) + Chr(0)
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Customer Request", 40) & "|" & PSTR("Repair advice by service advisor", 40) & Space(9) & Chr(27) + Chr(45) + Chr(0)
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Comp Code", 10) & "|" & PSTR("Complaint Detail", 23) & "|" & PSTR("JobCode", 10) & "|" & PSTR("Repair Det.", 10) & "|" & PSTR("Std.Hrs", 7) & "|" & PSTR("Parts Est.", 7) & "|" & PSTR("Lab. Est.", 7) & Space(10) & Chr(27) + Chr(45) + Chr(0); ""
            mHeader = mHeader + 1
            If TotalRows > 18 Then TotalRows1 = prnPos + 18 Else TotalRows1 = prnPos + TotalRows
            
                For I = prnPos To TotalRows1
                    Print #1, PSTR(FGrid1.TextMatrix(I, Col_Code), 10) & "|" & PSTR(FGrid1.TextMatrix(I, Col_Trouble), 23) & "|" & Space(10) & "|" & Space(10) & "|" & Space(7) & "|"; Space(7) & "|"; Space(7) & ""
                    mHeader = mHeader + 1
                Next
                For I = TotalRows1 To prnPos + 18
                    Print #1, Space(10) & "|" & Space(23) & "|" & Space(10) & "|" & Space(10) & "|" & Space(7) & "|"; Space(7) & "|"; Space(7) & ""
                    mHeader = mHeader + 1
                Next
            
            Print #1, Chr(27) + Chr(45) + Chr(1) & Space(PageWidth) & Space(10) & Chr(27) + Chr(45) + Chr(0)
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Estimate Charged :", PageWidth / 3) & "|" & PSTR("Est.Del.Time :" & Txt(DelTime), PageWidth / 3) & "|" & PSTR("Est.Del.Date :" & Txt(DelDt), (PageWidth / 3) - 3) & Space(10) & Chr(27) + Chr(45) + Chr(0)
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Revision in the Job :", PageWidth) & Space(10) & Chr(27) + Chr(45) + Chr(0)
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Dscription", 16) & "|" & PSTR("Est.Parts", 10) & "|" & PSTR("Est.Lab", 10) & "|" & PSTR("Est.Time", 10) & "|" & PSTR("Cust. Permission for additional job", 40) & Chr(27) + Chr(45) + Chr(0); ""
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & Space(16) & "|" & Space(10) & "|" & Space(10) & "|" & Space(10); "|" & PSTR("Time of calling up", 30) & Space(10) & Chr(27) + Chr(45) + Chr(0); ""
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & Space(16) & "|" & Space(10) & "|" & Space(10) & "|" & Space(10); "|" & PSTR("Cust. Aggrees for add. work (Y/N)", 40) & Chr(27) + Chr(45) + Chr(0); ""
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Total", 16) & "|" & Space(10) & "|" & Space(10) & "|" & Space(10); "|" & Space(40) & Chr(27) + Chr(45) + Chr(0); ""
            mHeader = mHeader + 1
            Print #1, PSTR("I authorise to execute the jobs", 35) & "|" & PSTR("I Certify that the work has", 35) & "|" & PSTR("Payment Details", 25) & ""
            mHeader = mHeader + 1
            Print #1, PSTR("described herein using necessary", 35) & "|" & PSTR("been done to my satisfact-", 35) & "|" & Space(25) & ""
            mHeader = mHeader + 1
            Print #1, PSTR("material at my cost.I understand", 35) & "|" & PSTR("ion and i have taken the", 35) & "|" & "Cash   Card  Warr" & ""
            mHeader = mHeader + 1
            Print #1, PSTR("that the vehicle is being stored", 35) & "|" & PSTR("delivery of vehicle.", 35) & "|" & Space(25) & ""
            mHeader = mHeader + 1
            Print #1, PSTR("repaired and tested at my risk", 35) & "|" & Space(35) & "|" & PSTR("Invoice No.", 18) & ""
            mHeader = mHeader + 1
            Print #1, Chr(27) + Chr(45) + Chr(1) & PSTR("Time            Customer Sign.", 35) & "|" & PSTR("Time            Customer Sign.", 35) & "|" & PSTR("Date", 18) & Chr(27) + Chr(45) + Chr(0); ""
            
            mHeader = mHeader + 1
            Print #1, "All deliveries against payment by cash or card only.Please fill the correct timings"
            mHeader = mHeader + 1
            Print #1, "above in order to help us serve you faster.THANK YOU"
            mHeader = mHeader + 1
            Print #1, mChr17 & "" & Space((PageWidth * 1.7) - Len("" & pubUName & "   " & PubServerDate)) & pubUName & "   " & PubServerDate & mChr18
            prnPos = prnPos + 18
            TotalRows = TotalRows - 18
            Print #1, mEject
        Wend
    End If
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''    Else
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
''    End If
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub SpeedPrintPCD000001(PrePrn As Boolean)
On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, j As Integer, ColorName$, Cnt As Integer
    Dim PrintStr$, mPhone$
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, RepTitle$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double, FormulaStr1$, FormulaStr2$
    Dim fob As New FileSystemObject, SecondStr As Boolean
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, mQry$
    Dim GRs As ADODB.Recordset
    
    Set GRs = GCn.Execute("SELECT ColMast.Col_Desc FROM HisCard LEFT JOIN ColMast ON HisCard.ColourCode = ColMast.Col_Code where Hiscard.chassis = '" & Txt(Chassis) & "'")
    If GRs.RecordCount > 0 And GRs.EOF = False And GRs.EOF = False Then
        ColorName = IIf(IsNull(GRs!Col_Desc), "", GRs!Col_Desc)
    Else
        ColorName = ""
    End If
    Set GRs = Nothing
    
'    If RstJob.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.Caption: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 1
'    For Page = 1 To 5
'        Print #1, "Page No." & str(Page)
'        Print #1, mEject
'    Next
'    GoTo PrnLoop

'    Footer = XNull(GCn.Execute("select JobCardFooter from Syctrl").Fields(0).Value)
'    For i = 1 To Len(Footer)
'        If Mid(Footer, i, 1) = vbLf Then
'            FooterCnt = FooterCnt + 1
'        End If
'    Next
 
    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
        
    mHeader = 14
    mFooter = 17
    mFooter = mFooter '+ FooterCnt
      
    'Header
      RepTitle = GCn.Execute("Select Div_SName from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
      mDocStr = "Job Card" & IIf(RepTitle = "", "", " (" & RepTitle & ")")
      
         
        Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
        
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        mHeader = 9 '7
        Print #1, mChr17 & PSTR(Txt(OwnerName), 40) & mChr18 & Space(16) & mEmph & PSTR(PrinID(JobDocID), 17, , AlignLeft) & Space(10) & Txt(Chassis) & mEmph1 & mChr17
        mHeader = mHeader + 1
        Print #1, PSTR(Txt(Address1), 40) & mChr18 & mEmph & Space(16) & PSTR(Txt(JobDt), 11) & mEmph1 & mChr17 & Space(17) & Txt(Engine)
        mHeader = mHeader + 1
        Print #1, PSTR(Txt(Address2), 40) & mChr18 & Space(16) & PSTR(Txt(ArrTime), 11) & Space(10) & mEmph & Txt(VehRegNo) & mEmph1 & mChr17
        mHeader = mHeader + 1
        Print #1, PSTR(Txt(Address3) & Txt(City), 40) & mChr18 & Space(37) & Txt(Model) & mChr17
        mHeader = mHeader + 1
        Print #1, Space(40) & mChr18 & Space(37) & ColorName & mChr17
        mHeader = mHeader + 1
        Print #1, Space(40) & Space(20) & PSTR(Txt(CurrentKMS), 14) & PSTR(Txt(FUEL), 15) & mChr18 & Space(9) & Txt(SaleDate) & mChr17
        mHeader = mHeader + 1
        mPhone = "       " & IIf(Txt(PhoneOff) = "", Txt(PhoneResi), Txt(PhoneOff))
        mPhone = left(mPhone + Space(40), 40)
        Print #1, mPhone & Space(45) & Space(19) & mChr17 & Txt(DNAME) & mChr18
        mHeader = mHeader + 1
        Print #1, "": Print #1, ""
        mHeader = mHeader + 3
        mFix = PageLength - (mHeader + mFooter)
        Page = 1
        mLine = 1
        mSlNo = 1
        Cnt = 1
        If FGrid1.Rows > 0 Then
            I = 1
            Do Until I = FGrid1.Rows
                If I = 33 Then Exit Do
                PrintStr = mChr17 & PSTR(I, 3) & Space(4) & PSTR(FGrid1.TextMatrix(I, Col_Trouble), 70) & mChr18 & Space(3) & mChr17
                Select Case I
                    Case 1
                        PrintStr = PrintStr & Space(5) & PSTR(Format(Txt(EstLab), "0.00"), 12, AlignRight) & PSTR(Format(Txt(EstSpr), "0.00"), 12) & Space(20) & PSTR(Format(Val(Txt(EstLab)) + Val(Txt(EstSpr)), "0.00"), 12)
                    Case 5
                        PrintStr = PrintStr & Space(5) & PSTR(Txt(DelDt), 12) & PSTR(Txt(DelTime), 12)
                    Case 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32
                        PrintStr = PrintStr
                End Select
                Print #1, PrintStr
                I = I + 1
                mSlNo = mSlNo + 1
                mLine = mLine + 1
            Loop
            If FGrid1.Rows < 33 Then
                Do Until I > 32
                    PrintStr = mChr17 & Space(3) & Space(4) & Space(70) & mChr18 & Space(3) & mChr17
                    Select Case I
                        Case 1
                            PrintStr = PrintStr & Space(5) & PSTR(Format(Txt(EstLab), "0.00"), 12, AlignRight) & PSTR(Format(Txt(EstSpr), "0.00"), 12) & Space(12) & PSTR(Format(Val(Txt(EstLab)) + Val(Txt(EstSpr)), "0.00"), 12)
                        Case 5
                            PrintStr = PrintStr & Space(5) & PSTR(Txt(DelDt), 12) & PSTR(Txt(DelTime), 12)
                        Case 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32
                            PrintStr = PrintStr
                    End Select
                    Print #1, PrintStr
                    I = I + 1
                    mSlNo = mSlNo + 1
                    mLine = mLine + 1
                Loop
            End If
        End If
        Do Until mLine >= mFix
            Print #1, ""
            mLine = mLine + 1
        Loop
        'FOOTER
        Cnt = 23
        Print #1, "": Print #1, "": Print #1, "": Print #1, "" ': Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, Space(33) & Txt(Supervisor)
        Print #1, "": Print #1, "":
        Print #1, Space(33) & Txt(Mechanic)
        Print #1, ""
        Print #1, mChr17 & "" & Space((PageWidth * 1.7) - Len("" & pubUName & "   " & PubServerDate)) & pubUName & "   " & PubServerDate & mChr18


        Print #1, mEject
        If FGrid1.Rows > 33 Then
            I = 33
            mHeader = 0
            Print #1, "": Print #1, "": Print #1, "": Print #1, ""
            Print #1, "": Print #1, "": Print #1, ""
            mHeader = 7
            Print #1, mChr17 & PSTR(Txt(OwnerName), 40) & mChr18 & Space(16) & mEmph & PSTR(PrinID(JobDocID), 17, , AlignLeft) & Space(10) & Txt(Chassis) & mEmph1 & mChr17
            mHeader = mHeader + 1
            Print #1, PSTR(Txt(Address1), 40) & mChr18 & mEmph & Space(16) & PSTR(Txt(JobDt), 11) & mEmph1 & mChr17 & Space(17) & Txt(Engine)
            mHeader = mHeader + 1
            Print #1, PSTR(Txt(Address2), 40) & mChr18 & Space(16) & PSTR(Txt(ArrTime), 11) & Space(10) & mEmph & Txt(VehRegNo) & mEmph1 & mChr17
            mHeader = mHeader + 1
            Print #1, PSTR(Txt(Address3) & Txt(City), 40) & mChr18 & Space(37) & mChr17 & Txt(Model)
            mHeader = mHeader + 1
            Print #1, "Contd. from previous page.." & STR(Page)
            mHeader = mHeader + 1
            Print #1, ""
            mHeader = mHeader + 1
            Print #1, ""
            mHeader = mHeader + 1
            Print #1, "": Print #1, "": Print #1, ""
            mHeader = mHeader + 3
            mFix = PageLength - (mHeader + mFooter)
            
            Do Until I = FGrid1.Rows
                PrintStr = mChr17 & PSTR(I, 3) & Space(4) & PSTR(FGrid1.TextMatrix(I, Col_Trouble), 70) & mChr18
                Print #1, PrintStr
                I = I + 1
            Loop
        End If
PrnLoop:
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''    Else
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
''    End If
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub WindowsPrintCVD(Index As Integer, mQry$, PrePrn As Boolean)
Dim RepTitle$, Speciality$, FormulaStr1$, FormulaStr2$, Condstr$
Dim Rst As ADODB.Recordset, RST1 As ADODB.Recordset
Dim I As Integer
Dim SecondStr  As Boolean
On Error GoTo ERRORHANDLER

Set Rst = New Recordset
Rst.CursorLocation = adUseClient
Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic

If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
RepTitle = GCn.Execute("Select Div_SName from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
Speciality = GCn.Execute("Select W_SecSpeciality from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value

CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")

FormulaStr1 = ""
FormulaStr2 = ""
SecondStr = False
Dim Cnt As Integer
Cnt = 80
For I = 1 To FGrid2.Rows - 1
If FGrid2.TextMatrix(I, ElName) <> "" Then
    If Len(FormulaStr1 & IIf(FormulaStr1 = "", "", " | ") & FGrid2.TextMatrix(I, ElName) & " : " & FGrid2.TextMatrix(I, ElValue)) > 255 Then
        SecondStr = True
        FormulaStr2 = FormulaStr2 & " | " & FGrid2.TextMatrix(I, ElNature) & " : " & FGrid2.TextMatrix(I, ElDefault)
    Else
        FormulaStr1 = FormulaStr1 & IIf(FormulaStr1 = "", "", " | ") & FGrid2.TextMatrix(I, ElName) & " : " & FGrid2.TextMatrix(I, ElValue)
        'FormulaStr2 = FormulaStr1
    End If
End If
Next
FormulaStr1 = left(FormulaStr1, 250)
FormulaStr2 = ""
'FormulaStr2 = Trim(FormulaStr2)
Set RST1 = New Recordset
RST1.CursorLocation = adUseClient
RST1.Open "select W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'", GCn, adOpenDynamic, adLockOptimistic

For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("SubTitle")
            rpt.FormulaFields(I).TEXT = "'" & Speciality & "'"
        Case UCase("FormulaStr1")
            rpt.FormulaFields(I).TEXT = "'" & FormulaStr1 & "' "
        Case UCase("FormulaStr2")
            If SecondStr = True Then
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr2 & "'"
            Else
                rpt.FormulaFields(I).TEXT = ""
            End If
        Case UCase("Phone")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecPhone & "'"
        Case UCase("Fax")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecFax & "'"
    End Select
Next

FormulaStr1 = left(FormulaStr1, 200)
rpt.Database.SetDataSource Rst
rpt.ReadRecords
Select Case Index
    Case PWindows  'Printer
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
            Case UCase("Title")
              If StrCmp(left(PubComp_Name, 4), "yash") Then
                 If MsgBox(" Customer Copy ", vbYesNo) = vbYes Then
                     rpt.FormulaFields(I).TEXT = "'" & "Job Card - Customer Copy" & "'"
                 ElseIf MsgBox(" Office Copy ", vbYesNo) = vbYes Then
                     rpt.FormulaFields(I).TEXT = "'" & "Job Card - Office Copy" & "'"
                 ElseIf MsgBox(" Shop Floor Copy ", vbYesNo) = vbYes Then
                     rpt.FormulaFields(I).TEXT = "'" & "Job Card - Shop Floor Copy" & "'"
                 End If
              
              Else
                rpt.FormulaFields(I).TEXT = "'" & "Job Card" & IIf(IsNull(RepTitle), "", " (" & RepTitle & ")") & "'"
             End If
        End Select
        Next
        rpt.PrintOut False
    Case PScreen  'screen
    'kunal
     If StrCmp(left(PubComp_Name, 4), "yash") Then
            If MsgBox(" Customer Copy ", vbYesNo) = vbYes Then
                Call Report_View(rpt, "Job Card - Customer Copy") '& IIf(IsNull(RepTitle), "", " (" & RepTitle & ")"), , True)
            ElseIf MsgBox(" Office Copy ", vbYesNo) = vbYes Then
                Call Report_View(rpt, "Job Card - Office Copy") '& IIf(IsNull(RepTitle), "", " (" & RepTitle & ")"), , True)
            ElseIf MsgBox(" Shop Floor Copy ", vbYesNo) = vbYes Then
                Call Report_View(rpt, "Job Card - Shop Floor Copy") '& IIf(IsNull(RepTitle), "", " (" & RepTitle & ")"), , True)
            End If
     Else
                Call Report_View(rpt, "Job Card" & IIf(IsNull(RepTitle), "", " (" & RepTitle & ")"), , True)
     End If
End Select
CmdPrint(PSetUp).Tag = ""
Set Rst = Nothing
Set RST1 = Nothing
Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub SpeedPrintCVD(mQry$, PrePrn As Boolean)

    Dim I As Integer, j As Integer
    Dim PrintStr$
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, RepTitle$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double, FormulaStr1$, FormulaStr2$
    Dim fob As New FileSystemObject, SecondStr As Boolean
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double ', mQry$
    Dim mDocId$, X(5) As String
    
    Set RstJob = GCn.Execute(mQry)
    If RstJob.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1

    FooterCnt = 1
    Footer = XNull(GCn.Execute("select JobCardFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next

    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7

    mDocId = left(RstJob!DocID, 3) & "-" & Trim(DeCodeDocID(RstJob!DocID, Document_No))
    
    mHeader = 0   'Ideal 17
    mFooter = 14
    mFooter = mFooter + FooterCnt

    'Header
    RepTitle = GCn.Execute("Select Div_SName from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
    mDocStr = "Job Card" & IIf(RepTitle = "", "", " (" & RepTitle & ")")
    
    For I = 1 To FGrid2.Rows - 1
        If FGrid2.TextMatrix(I, ElName) <> "" Then
            If Len(X(j) & IIf(X(j) = "", "", " | ") & FGrid2.TextMatrix(I, ElName) & ":" & FGrid2.TextMatrix(I, ElValue)) > 130 Then
                j = j + 1
                X(j) = FGrid2.TextMatrix(I, ElName) & ":" & FGrid2.TextMatrix(I, ElValue)
            Else
                X(j) = X(j) & IIf(X(j) = "", "", " | ") & FGrid2.TextMatrix(I, ElName) & " : " & FGrid2.TextMatrix(I, ElValue)
            End If
        End If
    Next
        Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")

        Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
        mHeader = mHeader + 1
        If XNull(RstCompDet!W_SecSpeciality) <> "" Then
           Print #1, PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
           mHeader = mHeader + 1
        End If
        Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
        If PubComp_Add2 <> "" Then
           Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
           mHeader = mHeader + 1
        End If
        If PubComp_City <> "" Then
            Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        'Print #1, PSTR("PHONE No :" & XNull(RstCompDet!W_SecPhone), 40) & PSTR(IIf(XNull(RstCompDet!W_SecFax) = "", "", "FAX : " & XNull(RstCompDet!W_SecFax)), 40, , AlignRight, " ")
        'mHeader = mHeader + 1

        Print #1, PRN_TIT("** " & mDocStr & " **", "A", PageWidth) & mChr18 & mEmph
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PSTR("Job Card No.", 15) & ": " & PSTR(mDocId, 21) & mEmph1 & Space(4) & PSTR("Model", 14) & ": " & XNull(RstJob!Model) '& mEmph
        mHeader = mHeader + 1
        Print #1, mEmph & PSTR("Job Card Date", 15) & ": " & PSTR(STR(RstJob!Job_Date), 21) & Space(4) & PSTR("Reg. No. & Dt.", 14) & ": " & XNull(RstJob!RegNo) & " " & XNull(RstJob!RegDate)
        mHeader = mHeader + 1
        Print #1, "Owner:" & PSTR(RstJob!Name, 35) & Space(1) & PSTR("Chassis No.", 14) & ": " & XNull(RstJob!Chassis) & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstJob!Add1), 40) & Space(2) & PSTR("Engine No.", 14) & ": " & XNull(RstJob!Engine)
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(2) & PSTR("Date Of Sale", 14) & ": " & XNull(RstJob!Delivery_Date)
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstJob!CityName), 25) & Space(17) & PSTR("Telco Inv. No.", 14) & ": " & XNull(RstJob!SUPPLIER_BILLNO) & " Dt." & IIf(IsNull(RstJob!Supplier_BillDate), "", RstJob!Supplier_BillDate)
        mHeader = mHeader + 1
        Print #1, PSTR("Phone No.", 9) & ": " & PSTR(RstJob!PhoneOff & IIf(RstJob!PhoneOff = "", "", "(O)") & RstJob!PhoneResi & IIf(RstJob!PhoneResi = "", "", "(R)") & RstJob!Mobile & IIf(RstJob!Mobile = "", "", "(M)"), 30) & Space(1) & PSTR("Dealer Code", 12) & ": " & XNull(RstJob!dealer_code) & "   RLW: " & XNull(RstJob!RLW)
        mHeader = mHeader + 1
        Print #1, mEmph & PSTR("Service Type ", 15) & ": " & PSTR(RstJob!Serv_Desc, 22) & mEmph1 & Space(3) & PSTR("FI Pump", 8) & ": " & XNull(RstJob!FIP_No) & mChr17 & "   Coupan No.: " & XNull(RstJob!CouponNo) & mChr18
        mHeader = mHeader + 1
        Print #1, PSTR("Axle No.(Front)", 15) & ": " & PSTR(RstJob!FAxelNo, 22) & Space(3) & PSTR("GearBox No.", 12) & ": " & XNull(RstJob!GBoxNo)
        mHeader = mHeader + 1
        Print #1, PSTR("Axle No.(Rear)", 15) & ": " & PSTR(RstJob!RAxelNo, 22) & Space(3) & PSTR("Battery Make", 12) & ": " & XNull(RstJob!Battery)
        mHeader = mHeader + 1
        Print #1, PSTR("Steer GBox No.", 15) & ": " & PSTR(RstJob!SteerGBNo, 22) & Space(3) & PSTR("Kms:", 12) & ": " & Trim(STR(RstJob!AtKMsHrs))
        mHeader = mHeader + 1
        Print #1, Space(42) & PSTR("Hr.Meter", 12) & ": " & XNull(RstJob!HrMeter)
        mHeader = mHeader + 1
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
        Print #1, Space(30) & "Time(HH:MM) " & PSTR("Signature", 10) & " | " & mEmph & "Exp. Delivery Date & Time" & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR("Arrival Of Customer", 30) & PSTR(CStr(Format(RstJob!ArrivalTime, "HH:MM")), 12) & Replace(Space(10), " ", "_") & " | " & XNull(RstJob!ExpDelDate)
        mHeader = mHeader + 1
        Print #1, PSTR("Reception Time Of Customer", 30) & PSTR(CStr(Format(RstJob!Recp_Time, "HH:MM")), 12) & Replace(Space(10), " ", "_") & " | " & mEmph & "Estimated Cost Of Repair" & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR("Customer Attended By Srv.Adv.", 42) & PSTR("_", 10) & " | " & PSTR("Labour(Rs.)", 13) & " : " & STR(RstJob!Est_LabCost)
        mHeader = mHeader + 1
        Print #1, PSTR("Completion Of JobCard", 42) & PSTR("_", 10) & " | " & PSTR("Spares(Rs.)", 13) & " : " & STR(RstJob!Est_SpCost)
        mHeader = mHeader + 1
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
        For I = 0 To UBound(X())
            If X(I) <> "" Then
                Print #1, mChr17 & X(I) & mChr18
                mHeader = mHeader + 1
            End If
        Next
        Print #1, "Damage to the Vehicle/Shortage : " & XNull(RstJob!body_damage)
        mHeader = mHeader + 1
        Print #1, "Remarks                        : " & XNull(RstJob!OpenRemarks)
        mHeader = mHeader + 1

        Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
        mHeader = mHeader + 1
        Print #1, PSTR("SRL.", 4) & "|" & PSTR("COMPLAINTS", 28) & "|" & PSTR("IF REPEAT,LAST", 14) & "|" & PSTR("ACTION", 10) & "|" & PSTR("ATTD. BY", 9) & "|" & PSTR("LABOUR", 8)
        mHeader = mHeader + 1
        Print #1, PSTR("NO.", 4) & "|" & Space(28) & "|" & PSTR("ATTENDED ON/BY", 14) & "|" & Space(10) & "|" & Space(9) & "|" & PSTR("CHARGE", 8) & mDoub1
        mHeader = mHeader + 1
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
        mFix = PageLength - (mHeader + mFooter)
        Page = 1
        mLine = 1
        mSlNo = 1
        If RstJob.RecordCount > 0 Then
        I = 1
            Do Until RstJob.EOF
                If mLine > mFix Then
                  
'                    Print #1, mChr18 & Replace(Space(PageWidth), " ", "-")
                    Print #1, Space(PageWidth - Len("Contd. on next page.." + STR(Page))) & "Contd. on next page.." & STR(Page)
'                    Do Until mLine  >= mFix + mFooter - 2
'                        Print #1, ""
'                        mLine = mLine + 1
'                    Loop
                    If Page = 1 Then
                        RstJob.MoveFirst
                        Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
                        Footer = Footer + vbLf
                        j = 1
                        For I = 1 To Len(Footer)
                            If mID(Footer, I, 1) = vbLf Then
                                Print #1, RTrim(mID(Footer, j, I - j))
                                j = I + 1
                            End If
                         Next
                    
                        Print #1, mChr18 & ""
                        Print #1, PSTR("Service Advisor", 55) & "Customer's Signature"
                        Print #1, ""
                        Print #1, "Signature : "
                        Print #1, "Name : " & RstJob!Name
                        Print #1, PSTR("Mechanic Name : " & XNull(RstJob!Emp_Name), 55) & "Name :"
                        Print #1, Replace(Space(PageWidth), " ", "-")
                        Print #1, "Vehicle delivered vide gate pass no.____________________dated____________"
                        Print #1, "Vehicle received after repair of all jobs to my satisfaction Yes/No._____"
                        Print #1, "Jobs at Srl.No.______not attended."
                        Print #1, "Received on _______________Time____________"
                        Print #1, "                                            Customer's Signature"
                        Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("")) / 2) & "" & mChr18
                    End If
                    Page = Page + 1
                    Print #1, mEject

                    'Header On Second Page
                    mHeader = 0
'                    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
'                    mHeader = mHeader + 1
'                    If XNull(RstCompDet!W_SecSpeciality) <> "" Then
'                        Print #1, PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth)
'                        mHeader = mHeader + 1
'                    End If

'                    Print #1, PRN_TIT("** " & mDocStr & " **", "A", PageWidth) & mChr18 & mEmph
'                    mHeader = mHeader + 1

                    Print #1, PSTR("Job Card No.", 15) & ": " & PSTR(mDocId, 14) & mEmph1 & Space(18) & PSTR("Model", 14) & " : " & XNull(RstJob!Model) & mEmph
                    mHeader = mHeader + 1
'                    Print #1, PSTR("Job Card Date", 15) & ": " & PSTR(str(RstJob!JOB_DATE), 14) & mEmph1 & Space(18) & PSTR("Reg. No. & Dt.", 14) & " : " & XNull(RstJob!RegNo) & " " & XNull(RstJob!RegDate) & mEmph
'                    mHeader = mHeader + 1
'                    Print #1, "Owner : " & PSTR(RstJob!Name, 40) & mEmph1 & Space(1) & PSTR("Chassis No.", 14) & " : " & XNull(RstJob!Chassis)
'                    mHeader = mHeader + 1
'                    Print #1, PSTR(XNull(RstJob!Add1), 40) & Space(5) & PSTR("Engine No.", 14) & " : " & XNull(RstJob!Engine)
'                    mHeader = mHeader + 1
'                    Print #1, PSTR(XNull(RstJob!Add2), 40) & Space(1) & PSTR("Date Of Sale", 14) & " : " & XNull(RstJob!Delivery_Date) & " Kms : " & str(RstJob!ATKMSHRS)
'                    mHeader = mHeader + 1
'                    Print #1, PSTR(XNull(RstJob!CityName), 25) & Space(16) & PSTR("Telco Inv. No.", 14) & " : " & RstJob!SUPPLIER_BILLNO & " Dt. " & XNull(RstJob!Supplier_BillDate)
'                    mHeader = mHeader + 1
'                    Print #1, PSTR("Phone No.", 15) & ": " & PSTR(RstJob!PhoneOff & "(O)" & RstJob!PhoneResi & "(R)", 30) & Space(1) & PSTR("Dealer Code", 12) & " : " & XNull(RstJob!dealer_code) & " Coupan No. : " & XNull(RstJob!CouponNo)
'                    mHeader = mHeader + 1

                    Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
                    mHeader = mHeader + 1
                    Print #1, PSTR("SRL.", 4) & "|" & PSTR("COMPLAINTS", 28) & "|" & PSTR("IF REPEAT,LAST", 14) & "|" & PSTR("ACTION", 10) & "|" & PSTR("ATTD. BY", 9) & "|" & PSTR("LABOUR", 8)
                    mHeader = mHeader + 1
                    Print #1, PSTR("NO.", 4) & "|" & Space(28) & "|" & PSTR("ATTENDED ON/BY", 14) & "|" & Space(10) & "|" & Space(9) & "|" & PSTR("CHARGE", 8) & mDoub1
                    mHeader = mHeader + 1
                    Print #1, Replace(Space(PageWidth), " ", "-")
                    mHeader = mHeader + 1
                    mFix = PageLength - mHeader
                    mLine = 1
                End If
            PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 4) & "|" & mChr17 & PSTR(RstJob!Details, 48) & mChr18 & "|" & Space(14) & "|" & Space(10) & "|" & Space(9) & "|"

            Print #1, PrintStr
            RstJob.MoveNext
            mSlNo = mSlNo + 1
            mLine = mLine + 1
        Loop
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop

    ' FOOTER
    If Page = 1 Then
        RstJob.MoveFirst
        Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
        Footer = Footer + vbLf
        j = 1
        For I = 1 To Len(Footer)
            If mID(Footer, I, 1) = vbLf Then
                Print #1, RTrim(mID(Footer, j, I - j))
                j = I + 1
            End If
         Next
    
        Print #1, mChr18 & ""
        Print #1, "Supervisor Signature : "
        Print #1, PSTR("Name: " & Txt(Supervisor), 40) & "Customer's Signature:"
        Print #1, PSTR("Mechanic Name: " & XNull(RstJob!MechName), 40); "Name:" & RstJob!Name
        Print #1, Replace(Space(PageWidth), " ", "-")
        Print #1, "Vehicle delivered vide gate pass no.____________________dated____________"
        Print #1, "Vehicle received after repair of all jobs to my satisfaction Yes/No._____"
        Print #1, "Jobs at Srl.No.______not attended."
        Print #1, "Received on _______________Time____________"
        Print #1, "                                                     Customer's Signature"
        Print #1, mChr17 & "" & Space((PageWidth * 1.7) - Len("" & pubUName & "   " & PubServerDate)) & pubUName & "   " & PubServerDate & mChr18
    End If
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
'        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
'    Else
'        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub SpeedPrintCVD000001(PrePrn As Boolean)
On Error GoTo ELoop
    Dim I As Integer, j As Integer, Str2$
    Dim Rst As ADODB.Recordset
    Dim GBNo$, FAxle$, RAxle$, FIP$, STBox$
    Dim PrintStr$
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, RepTitle$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double, FormulaStr1$, FormulaStr2$
    Dim fob As New FileSystemObject, SecondStr As Boolean
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, mQry$

    GBNo = ""
    FAxle = ""
    RAxle = ""
    FIP = ""
    STBox = ""
    Set Rst = GCn.Execute("SELECT GBoxNo,RAxelNo,SteerGBNo,FIP_No,FaxelNo From HisCard where CardNo = '" & Txt(HistNo) & "'")

    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1

    FooterCnt = 1
    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7

    mHeader = 0   'Ideal 17
    mFooter = 35
    
    'Header
    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
    Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
    mHeader = 9 '7
        
    Print #1, Space(7) & mEmph & PSTR(PrinID(JobDocID), 21, , AlignRight) & Space(18) & Txt(JobDt) & Space(1) & Txt(ArrTime) & mEmph1
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, Space(12) & Txt(OwnerName)
    mHeader = mHeader + 1
    Print #1, Space(12) & mChr17 & PSTR(Txt(Address1) & " " & Txt(Address2), 80) & mChr18
    mHeader = mHeader + 1
    Print #1, Space(12) & mChr17 & PSTR(Txt(Address3) & " " & Txt(City), 65) & mChr18
    mHeader = mHeader + 1
    Print #1, Space(12) & Txt(PhoneOff) & " " & Txt(PhoneResi) & " " & Txt(Mobile)
    mHeader = mHeader + 1
    Print #1, Space(12) & PSTR(Txt(Model), 35) & mEmph & Txt(VehRegNo) & mEmph1
    mHeader = mHeader + 1
    Print #1, Space(12) & mEmph & PSTR(Txt(Chassis), 35) & mEmph1 & Txt(Engine)
    mHeader = mHeader + 1
    Print #1, Space(12) & PSTR(Txt(CurrentKMS), 15) & Space(20) & Txt(InvNo)
    mHeader = mHeader + 1
    Print #1, Space(12) & PSTR(Txt(SrvType), 35) & PSTR(Txt(SaleDate), 28) & mChr17 & Txt(DCode) & mChr18
    mHeader = mHeader + 1
    Print #1, Space(12) & PSTR(Txt(Supervisor), 25)
    mHeader = mHeader + 1
    Print #1, "": Print #1, "": Print #1, ""
    mHeader = mHeader + 3
        
    mFix = PageLength - (mHeader + mFooter)
    Page = 1
    mLine = 1
    mSlNo = 1
    If FGrid1.Rows > 1 Then
        I = 1
        Do Until I > FGrid1.Rows - 1
            If mLine > mFix Then
                Print #1, Space(PageWidth - Len("Contd. on next page.." + STR(Page + 1))) & "Contd. on next page.." & STR(Page + 1)
                If Page = 1 Then
                    Print #1, ""
'                    Set Rst = GCn.Execute("SELECT GBoxNo,RAxelNo,SteerGBNo,FIP_No,FaxelNo From HisCard where CardNo = '" & Txt(Model).Tag & "'")
                    If Rst.RecordCount > 0 And Rst.EOF = False And Rst.BOF = False Then
                        GBNo = IIf(IsNull(Rst!GBoxNo), "", Rst!GBoxNo): RAxle = IIf(IsNull(Rst!RAxelNo), "", Rst!RAxelNo)
                        FAxle = IIf(IsNull(Rst!FAxelNo), "", Rst!FAxelNo): FIP = IIf(IsNull(Rst!FIP_No), "", Rst!FIP_No)
                        STBox = IIf(IsNull(Rst!SteerGBNo), "", Rst!SteerGBNo)
                    End If
                    Set Rst = Nothing
                    Print #1, Space(7) & PSTR(GBNo, 15) & Space(10) & PSTR(RAxle, 15) & Space(10) & PSTR(FAxle, 15)
                    Print #1, Space(7) & PSTR(FIP, 15) & Space(10) & PSTR(STBox, 15)
                    
                    Print #1, ""
                    Print #1, "" 'Space(17) & PSTR(Txt(DelDt), 11) & Space(6) & PSTR(Txt(DelTime), 7) & Space(17) & PSTR(Txt(EstSpr), 10) & Space(9) & PSTR(Txt(EstLab), 10)
                    Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "": Print #1, "": Print #1, ""
                        
                    Print #1, ""
                    Print #1, ""
                    Print #1, Space(52) & Txt(ArrTime)
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "" 'Space(44) & PSTR(txt(JobClDt), 20) & Space(12) & txt(DelDt)
                    Print #1, "" 'Space(44) & PSTR(txt(JCTime), 20) & Space(12) & txt(DelTime)
                    Print #1, ""
                    Print #1, "" 'Space(35) & mChr17 & PSTR(txt(Remarks), 60) & mChr18
                    Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("")) / 2) & "" & mChr18
                End If
                Page = Page + 1
                Print #1, mEject

                'Header On Second Page
                Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                mHeader = 7
                Print #1, mChr18 & Space(7) & mEmph & PSTR(PrinID(JobDocID), 17, , AlignRight) & Space(29) & Txt(JobDt) & Space(1) & Txt(ArrTime) & mEmph1
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, Space(12) & Txt(OwnerName)
                mHeader = mHeader + 1
                Print #1, "Contd. from last page.." + STR(Page - 1)
                mHeader = mHeader + 1
                '***
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, Space(12) & PSTR(Txt(Model), 35) & mEmph & Txt(VehRegNo) & mEmph1
                mHeader = mHeader + 1
                Print #1, Space(12) & mEmph & PSTR(Txt(Chassis), 35) & mEmph1 & Txt(Engine)
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, "": Print #1, "": Print #1, ""
                mHeader = mHeader + 3
                '***
                mLine = 1
            End If
            If I + 1 <= FGrid1.Rows - 1 Then
                Str2 = FGrid1.TextMatrix(I + 1, Col_Trouble)
            Else
                Str2 = ""
            End If
            PrintStr = mChr17 & PSTR(FGrid1.TextMatrix(I, Col_Trouble), 72) & Str2 & mChr18
            I = I + 2
            Print #1, PrintStr
            mSlNo = mSlNo + 1
            mLine = mLine + 1
        Loop
    End If
    
    ' FOOTER
    If Page = 1 Then
        Do Until mLine >= mFix
            Print #1, ""
            mLine = mLine + 1
        Loop
        Print #1, ""
        'Set Rst = GCn.Execute("SELECT GBoxNo,RAxelNo,SteerGBNo,FIP_No,FaxelNo From HisCard where model = '" & Txt(Model) & "'")
        If Rst.RecordCount > 0 And Rst.EOF = False And Rst.BOF = False Then
            GBNo = IIf(IsNull(Rst!GBoxNo), "", Rst!GBoxNo): RAxle = IIf(IsNull(Rst!RAxelNo), "", Rst!RAxelNo)
            FAxle = IIf(IsNull(Rst!FAxelNo), "", Rst!FAxelNo): FIP = IIf(IsNull(Rst!FIP_No), "", Rst!FIP_No)
            STBox = IIf(IsNull(Rst!SteerGBNo), "", Rst!SteerGBNo)
        End If
        Set Rst = Nothing
        Print #1, Space(7) & PSTR(GBNo, 15) & Space(10) & PSTR(RAxle, 15) & Space(10) & PSTR(FAxle, 15)
        Print #1, Space(7) & PSTR(FIP, 15) & Space(10) & PSTR(STBox, 15)
        
        Print #1, ""
        Print #1, Space(17) & PSTR(Txt(DelDt), 11) & Space(6) & PSTR(Txt(DelTime), 7) & Space(17) & PSTR(Txt(EstSpr), 10) & Space(7) & PSTR(Txt(EstLab), 9)

        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, ""
            
        Print #1, ""
        Print #1, ""
        Print #1, ""; Space(52) & Txt(ArrTime)
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, mChr17 & "" & Space((PageWidth * 1.7) - Len("" & pubUName & "   " & PubServerDate)) & pubUName & "   " & PubServerDate & mChr18
    End If

    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub
Private Function CheckRepeat(RegNo As String, Row As Integer)
    Dim mQry$
    Dim RstJob As ADODB.Recordset, rstJobDemand As ADODB.Recordset, I As Integer, j As Integer
    mQry = "Select JC.DocID,JC.Job_No,JC.Job_Date from Job_Card JC Left Join HisCard H on JC.CardNo=H.CardNo where H.RegNo='" & RegNo & "'"
    Set RstJob = GCn.Execute(mQry)
    RstJob.Sort = "Job_Date Desc"
    If RstJob.RecordCount = 1 Then Exit Function
    For I = 1 To IIf(RstJob.RecordCount <= 3, RstJob.RecordCount, 3)
    
        RstJob.MoveFirst:    RstJob.MoveNext
        If RstJob.EOF = True Then: Exit Function
        Set rstJobDemand = GCn.Execute("Select Code From Job_Demand where Job_DocId='" & RstJob!DocID & "'")
            If rstJobDemand.RecordCount > 0 Then
                For j = 1 To rstJobDemand.RecordCount
                  If FGrid1.TextMatrix(Row, Col_Code) = rstJobDemand!Code Then
                        CheckRepeat = "Job No : " & RstJob!Job_No & " and Date & : " & RstJob!Job_Date
                        Exit Function
                  End If
                rstJobDemand.MoveNext
                Next
            End If
        RstJob.MoveNext
    Next
End Function
Private Sub SpeedPrintCVD000001Rashmi(PrePrn As Boolean)
On Error GoTo ELoop
    Dim I As Integer, j As Integer, Str2$
    Dim Rst As ADODB.Recordset
    Dim GBNo$, FAxle$, RAxle$, FIP$, STBox$
    Dim PrintStr$
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, RepTitle$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double, FormulaStr1$, FormulaStr2$
    Dim fob As New FileSystemObject, SecondStr As Boolean
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, mQry$
    Set Rst = GCn.Execute("SELECT GBoxNo,RAxelNo,SteerGBNo,FIP_No,FaxelNo From HisCard where CardNo = '" & Txt(HistNo) & "'")
                    If Rst.RecordCount > 0 And Rst.EOF = False And Rst.BOF = False Then
                        GBNo = IIf(IsNull(Rst!GBoxNo), "", Rst!GBoxNo): RAxle = IIf(IsNull(Rst!RAxelNo), "", Rst!RAxelNo)
                        FAxle = IIf(IsNull(Rst!FAxelNo), "", Rst!FAxelNo): FIP = IIf(IsNull(Rst!FIP_No), "", Rst!FIP_No)
                        STBox = IIf(IsNull(Rst!SteerGBNo), "", Rst!SteerGBNo)
                    End If
   
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    FooterCnt = 1
    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 -> X * 0.56
    'chr 10 to chr 17 -> X * 1.7
    mHeader = 0   'Ideal 17
    mFooter = 35
    
    'Header
    Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
    mHeader = 7
    Print #1, Space(11) & mDoub & PSTR(Txt(JobNo), 15, , AlignLeft) & mDoub1 & Space(12) & PSTR(Txt(JobDt), 15, , AlignRight) & Space(12) & PSTR(Txt(JCTime), 15, , AlignRight)
    mHeader = mHeader + 1
    Print #1, ""
    mHeader = mHeader + 1
    Print #1, Space(11) & Chr(15) & PSTR(Txt(OwnerName), 30, , AlignLeft) & Chr(18) & Space(13) & mDoub & PSTR(Txt(VehRegNo), 15, , AlignLeft) & mDoub1 & Space(12) & Chr(15) & mDoub & PSTR(Txt(Chassis), 20, , AlignLeft) & mDoub1 & Chr(18)
    mHeader = mHeader + 1
    Print #1, Space(11) & PSTR(Txt(Address1), 19, , AlignLeft) & Space(12) & PSTR(Txt(Model), 15, , AlignLeft) & Space(12) & Chr(15) & PSTR(Txt(Engine), 17, , AlignLeft) & Chr(18)
    mHeader = mHeader + 1
    Print #1, Space(11) & PSTR(Txt(Address2), 15, , AlignLeft) & Space(16) & PSTR(Txt(SaleDate), 15, , AlignLeft) & Space(12) & PSTR(GBNo, 13, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, Space(11) & PSTR(Txt(PhoneOff), 15, , AlignLeft) & Space(16) & PSTR(Txt(InvNo), 15, , AlignLeft) & Space(12) & PSTR(RAxle, 13, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, Space(11) & PSTR(Txt(Mobile), 15, , AlignLeft) & Space(16) & PSTR(Txt(DCode), 15, , AlignLeft) & Space(12) & PSTR(FAxle, 13, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, Space(11) & PSTR(Txt(SrvType), 15, , AlignLeft) & Space(16) & PSTR(Txt(CurrentKMS), 15, , AlignLeft) & Space(12) & PSTR(STBox, 13, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, Space(11) & PSTR(Txt(Supervisor), 15, , AlignLeft) & Space(16) & PSTR(Txt(HrMeter), 15, , AlignLeft) & Space(12) & PSTR(FIP, 15, , AlignLeft)
    Print #1, "": Print #1, "": Print #1, ""
    Print #1, "": Print #1, "" ': Print #1, ""
    mHeader = mHeader + 5
     
    mFix = PageLength - (mHeader + mFooter)
    Page = 1
    mLine = 1
     If FGrid1.Rows > 1 Then
        I = 1
        Do Until I > FGrid1.Rows - 1
            If mLine > mFix Then
                Print #1, Space(PageWidth - Len("Contd. on next page.." + STR(Page + 1))) & "Contd. on next page.." & STR(Page + 1)
                If Page = 1 Then
                    Print #1, ""
                    Print #1, ""
                    Print #1, ""
                    Print #1, ""
                    Print #1, ""
                    Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "": Print #1, "": Print #1, ""
                    Print #1, ""
                    Print #1, ""
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                End If
                Page = Page + 1
                Print #1, mEject

                'Header On Second Page
                Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                mHeader = 7
                mHeader = mHeader + 1
                Print #1, "Contd. from last page.." + STR(Page - 1)
                mHeader = mHeader + 1
                '***
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, "": Print #1, "": Print #1, ""
                mHeader = mHeader + 3
                '***
                mLine = 1
            End If
            If I + 1 <= FGrid1.Rows - 1 Then
                Str2 = mChr17 & PSTR(FGrid1.TextMatrix(I + 1, Col_Trouble), 30, , AlignLeft) & mChr18
            Else
                Str2 = ""
            End If
            PrintStr = mChr17 & PSTR(FGrid1.TextMatrix(I, Col_Trouble), 30, , AlignLeft) & mChr18
            I = I + 2
            Print #1, PrintStr
            Print #1, Str2
            mLine = mLine + 1
        Loop
    End If
    
    ' FOOTER
    If Page = 1 Then
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, Space(9) & PSTR(Txt(DelDt), 11) & Space(10) & PSTR(Txt(DelTime), 7, , AlignLeft) & Space(12) & PSTR(Txt(EstSpr), 10, , AlignLeft) & Space(9) & PSTR(Txt(EstLab), 9)
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, ""
        Print #1, Space(10) & ""
    End If
     Close #1
    Open "C:\RepPrint.Bat" For Output As #1
    If fob.FolderExists("c:\Windows") Then
        If Len(Printer.DeviceName) > 0 Then
            mPrinterName = "Prn"
            If left(Printer.DeviceName, 2) = "\\" Then
                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
            End If
        Else
            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
        End If
    Else
        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
    End If
    Print #1, "Type C:\RepPrint.Txt>" & mPrinterName
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Sub SpeedPrintCVD000001Society(PrePrn As Boolean)
Dim TmpRst As ADODB.Recordset
On Error GoTo ELoop
    Dim ChasisNo, EngineNo, VehType As String
    Dim Exwar$
    Dim I As Integer, j As Integer, Str2$
    Dim Rst As ADODB.Recordset
    Dim GBNo$, FAxle$, RAxle$, FIP$, STBox$
    Dim PrintStr$
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstJob As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, RepTitle$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double, FormulaStr1$, FormulaStr2$
    Dim fob As New FileSystemObject, SecondStr As Boolean
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, mQry$
    Set Rst = GCn.Execute("SELECT GBoxNo,RAxelNo,SteerGBNo,FIP_No,FaxelNo From HisCard where CardNo = '" & Txt(HistNo) & "'")
                    If Rst.RecordCount > 0 And Rst.EOF = False And Rst.BOF = False Then
                        GBNo = IIf(IsNull(Rst!GBoxNo), "", Rst!GBoxNo): RAxle = IIf(IsNull(Rst!RAxelNo), "", Rst!RAxelNo)
                        FAxle = IIf(IsNull(Rst!FAxelNo), "", Rst!FAxelNo): FIP = IIf(IsNull(Rst!FIP_No), "", Rst!FIP_No)
                        STBox = IIf(IsNull(Rst!SteerGBNo), "", Rst!SteerGBNo)
                    End If
   
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    FooterCnt = 1
    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 -> X * 0.56
    'chr 10 to chr 17 -> X * 1.7
    mHeader = 0   'Ideal 17
    mFooter = 35
    
    'Header
    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
    mHeader = 4
    Print #1, Space(67) & mDoub & PSTR(Txt(JobNo), 15, , AlignLeft) & mDoub1
    mHeader = mHeader + 1
    Print #1, Space(57) & PSTR(Txt(JobDt), 15, , AlignRight) & Space(5) & PSTR(Txt(JCTime), 15, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, Space(64) & PSTR(Txt(VehRegNo), 15, , AlignRight)
    mHeader = mHeader + 1
    Print #1, Space(66) & PSTR(Txt(CurrentKMS), 15, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, ""
    Set TmpRst = GCn.Execute("Select * from HisCard where RegNo='" & Txt(VehRegNo).TEXT & "'")
    
    If Not TmpRst.EOF Then VehType = TmpRst!VehDet
    If VehType <> "" Then
        If VehType = "Personal" Then
            PrintStr = Space(10) & mDoub & "ü" & mDoub1
        ElseIf VehType = "Taxi" Then
            PrintStr = Space(20) & mDoub & "ü" & mDoub1
        ElseIf VehType = "Corporate" Then
            PrintStr = Space(30) & mDoub & "ü" & mDoub1
        End If
        Print #1, PrintStr
        mHeader = mHeader + 1
        Exwar = IIf(TmpRst!ExtendWar = 1, "Yes", "No")
    End If
    Print #1, Space(8) & Chr(15) & PSTR(Txt(OwnerName), 30, , AlignLeft) & Chr(18) & Space(17) & mDoub & PSTR(Txt(Model), 15, , AlignLeft) & mDoub1
    mHeader = mHeader + 1
    Print #1, Space(8) & PSTR(Txt(Address1), 19, , AlignLeft) & Space(17) & Chr(15) & PSTR(Txt(DNAME), 20, , AlignLeft) & Space(7) & PSTR(Txt(SaleDate), 13, , AlignLeft) & Chr(18) & Space(15) & PSTR(Exwar, 3, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, Space(8) & PSTR(Txt(Address2), 15, , AlignLeft)
    mHeader = mHeader + 1
    I = Len(Txt(Chassis))
    j = Len(Txt(Engine))
    For I = 1 To Len(Txt(Chassis))
        ChasisNo = ChasisNo & mID(Txt(Chassis), I, 1) & " "
    Next
    For j = 1 To Len(Txt(Engine))
        EngineNo = EngineNo & mID(Txt(Engine), j, 1) & " "
    Next
    Print #1, Space(46) & mDoub & PSTR(ChasisNo, 50, , AlignLeft) & mDoub1
    Print #1, Space(46) & mDoub & PSTR(EngineNo, 50, , AlignLeft) & mDoub1
    mHeader = mHeader + 2
    Print #1, Space(11) & PSTR(Txt(PhoneOff), 15, , AlignLeft) '& Space(16) & PSTR(Txt(InvNo), 15, , AlignLeft) & Space(12) & PSTR(RAxle, 13, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, Space(11) & PSTR(Txt(Mobile), 15, , AlignLeft) & Space(21) & PSTR(Txt(LastMech), 15, , AlignLeft) & Space(10) & PSTR(Txt(LastJobDt), 13, , AlignLeft)
    mHeader = mHeader + 1
    Print #1, Space(45) & PSTR(Txt(LastSrv), 15, , AlignLeft) & Space(5) & PSTR(Txt(LastKMS) & "Kms.", 15, AlignLeft)
    mHeader = mHeader + 1
    
    Print #1, ""
    mHeader = mHeader + 1
    PrintStr = TmpRst!Tyre_FL & Space(5) & TmpRst!Tyre_FR
    Print #1, Space(48) & Chr(15) & PSTR(PrintStr, 25, , AlignLeft) & Chr(18)
    mHeader = mHeader + 1
    
    PrintStr = TmpRst!Tyre_RL1 & Space(5) & TmpRst!Tyre_RR1
    Print #1, Space(48) & Chr(15) & PSTR(PrintStr, 25, , AlignLeft) & Chr(18)
    mHeader = mHeader + 1
    
    Print #1, "": Print #1, ""
    mHeader = mHeader + 2
    Print #1, Space(50) & Format(PSTR(Val(Txt(EstLab)) + Val(Txt(EstSpr)), 15, , AlignLeft), "0.00")
    mHeader = mHeader + 1
    Print #1, Space(50) & Chr(15) & PSTR(Txt(DelDt), 11) & Space(1) & PSTR(Txt(DelTime), 11) & Chr(18)
    mHeader = mHeader + 1
    'Print #1, Space(11) & PSTR(Txt(Supervisor), 15, , AlignLeft) & Space(16) & PSTR(IIf(Txt(KmsHrs) = "Hrs.", Txt(CurrentKMS), ""), 15, , AlignLeft) & Space(12) & PSTR(FIP, 15, , AlignLeft)
    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
    mHeader = mHeader + 3
    mFix = PageLength - (mHeader + mFooter)
    Page = 1
    mLine = 1
     If FGrid1.Rows > 1 Then
        I = 1
        Do Until I > FGrid1.Rows - 1
            If mLine > mFix Then
                Print #1, Space(PageWidth - Len("Contd. on next page.." + STR(Page + 1))) & "Contd. on next page.." & STR(Page + 1)
                If Page = 1 Then
                    Print #1, ""
                    Print #1, ""
                    Print #1, ""
                    Print #1, ""
                    Print #1, ""
                    Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "": Print #1, "": Print #1, ""
                    Print #1, ""
                    Print #1, ""
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                End If
                Page = Page + 1
                Print #1, mEject

                'Header On Second Page
                Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                Print #1, "": Print #1, "": Print #1, ""
                mHeader = 7
                mHeader = mHeader + 1
                Print #1, "Contd. from last page.." + STR(Page - 1)
                mHeader = mHeader + 1
                '***
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, ""
                mHeader = mHeader + 1
                Print #1, "": Print #1, "": Print #1, ""
                mHeader = mHeader + 3
                '***
                mLine = 1
            End If
            If I + 1 <= FGrid1.Rows - 1 Then
                Str2 = PSTR(FGrid1.TextMatrix(I + 1, Col_Trouble), 30, , AlignLeft)
            Else
                Str2 = ""
            End If
            PrintStr = Space(10) & PSTR(FGrid1.TextMatrix(I, Col_Trouble), 30, , AlignLeft)
            I = I + 2
            Print #1, PrintStr
            Print #1, Space(10) & Str2
            mLine = mLine + 1
        Loop
    End If
    
    ' FOOTER
    If Page = 1 Then
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, Space(15) & PSTR(Txt(DelDt), 11) & Space(15) & PSTR(Txt(DelTime), 7, , AlignLeft) '& Space(12) & PSTR(Txt(EstSpr), 10, , AlignLeft) & Space(9) & PSTR(Txt(EstLab), 9)
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, ""
        Print #1, Space(5) & PSTR(Txt(JCTime), 11) & Space(20) & PSTR(Txt(JCTime), 11)
        Print #1, "": Print #1, ""
        Print #1, Space(10) & ""
    End If
     Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt>" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub
