VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmRequisition 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Requisition From Supervisor Entry"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14010
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
   ScaleHeight     =   10290
   ScaleWidth      =   14010
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt 
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
      Index           =   28
      Left            =   10365
      MaxLength       =   10
      TabIndex        =   131
      Top             =   2250
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   29
      Left            =   5070
      MaxLength       =   11
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   510
      Visible         =   0   'False
      Width           =   1305
   End
   Begin MSDataGridLib.DataGrid DgRateType 
      Height          =   2100
      Left            =   555
      Negotiate       =   -1  'True
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   5445
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
   Begin VB.CommandButton cmdPost 
      Caption         =   "Re-Post"
      Height          =   345
      Left            =   7185
      TabIndex        =   125
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
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
      Left            =   3420
      TabIndex        =   111
      Top             =   4965
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
         Picture         =   "frmRequisition.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   121
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
         Picture         =   "frmRequisition.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   120
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmRequisition.frx":0678
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
         TabIndex        =   119
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmRequisition.frx":0982
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
         TabIndex        =   118
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmRequisition.frx":0C8C
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
         TabIndex        =   117
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
         TabIndex        =   116
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
         TabIndex        =   115
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
         TabIndex        =   114
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
         TabIndex        =   113
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
         TabIndex        =   112
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
         TabIndex        =   124
         Top             =   300
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
         TabIndex        =   123
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
         TabIndex        =   122
         Top             =   0
         Width           =   4695
      End
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   15
      TabIndex        =   46
      Top             =   4215
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   75
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   15
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
   Begin VB.Frame FrmDetail 
      BackColor       =   &H00CAF1FD&
      BorderStyle     =   0  'None
      DragIcon        =   "frmRequisition.frx":0F96
      ForeColor       =   &H00C00000&
      Height          =   2205
      Left            =   6645
      TabIndex        =   79
      Top             =   4515
      Visible         =   0   'False
      Width           =   6285
      Begin VB.Line Line3 
         X1              =   3750
         X2              =   3750
         Y1              =   1035
         Y2              =   2070
      End
      Begin VB.Line Line2 
         X1              =   2760
         X2              =   2475
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line1 
         X1              =   1755
         X2              =   75
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bin Location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   32
         Left            =   3765
         TabIndex        =   110
         Top             =   255
         Width           =   1020
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Bin Loca>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   4920
         TabIndex        =   109
         Top             =   255
         Width           =   930
      End
      Begin VB.Label LblFrm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<Part No>"
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
         Height          =   225
         Index           =   0
         Left            =   1140
         TabIndex        =   108
         Top             =   255
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
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
         Height          =   225
         Index           =   20
         Left            =   75
         TabIndex        =   107
         Top             =   270
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   21
         Left            =   1800
         TabIndex        =   106
         Top             =   930
         Width           =   660
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000.00"
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
         Left            =   2745
         TabIndex        =   105
         Top             =   1185
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last"
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
         Height          =   225
         Index           =   23
         Left            =   4920
         TabIndex        =   104
         Top             =   1185
         Width           =   360
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Index           =   11
         Left            =   3285
         TabIndex        =   103
         Top             =   1875
         Width           =   360
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   210
         Index           =   14
         Left            =   5460
         TabIndex        =   102
         Top             =   1185
         Width           =   765
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Height          =   255
         Index           =   6
         Left            =   2115
         TabIndex        =   101
         Top             =   1635
         Width           =   360
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   2115
         TabIndex        =   100
         Top             =   1875
         Width           =   360
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Height          =   255
         Index           =   10
         Left            =   3285
         TabIndex        =   99
         Top             =   1635
         Width           =   360
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   13
         Left            =   5460
         TabIndex        =   98
         Top             =   1657
         Width           =   765
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000.000"
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
         Left            =   5130
         TabIndex        =   97
         Top             =   930
         Width           =   1095
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   2100
         TabIndex        =   96
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Part Local Name>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   1140
         TabIndex        =   95
         Top             =   675
         Width           =   1590
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
         BeginProperty Font 
            Name            =   "Arial"
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
         Left            =   5460
         TabIndex        =   94
         Top             =   1410
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   19
         Left            =   4920
         TabIndex        =   93
         Top             =   1410
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00BBDBB3&
         BackStyle       =   0  'Transparent
         Caption         =   "Pur. Rate"
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
         Height          =   225
         Index           =   18
         Left            =   3930
         TabIndex        =   92
         Top             =   1185
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   17
         Left            =   75
         TabIndex        =   91
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
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
         Height          =   225
         Index           =   16
         Left            =   4920
         TabIndex        =   90
         Top             =   1650
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   15
         Left            =   2805
         TabIndex        =   89
         Top             =   930
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Paid"
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
         Index           =   14
         Left            =   75
         TabIndex        =   88
         Top             =   1875
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MRP Taxable"
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
         Left            =   75
         TabIndex        =   87
         Top             =   1185
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Stock"
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
         Index           =   12
         Left            =   3930
         TabIndex        =   86
         Top             =   915
         Width           =   1110
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item Detail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Index           =   11
         Left            =   0
         TabIndex        =   85
         Top             =   0
         Width           =   6285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taxable"
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
         Height          =   255
         Index           =   10
         Left            =   75
         TabIndex        =   84
         Top             =   1635
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Name"
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
         Left            =   75
         TabIndex        =   83
         Top             =   465
         Width           =   885
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Part Name>"
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
         Left            =   1140
         TabIndex        =   82
         Top             =   465
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MRP Taxpaid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   8
         Left            =   75
         TabIndex        =   81
         Top             =   1410
         Width           =   1080
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2115
         TabIndex        =   80
         Top             =   1410
         Width           =   360
      End
      Begin VB.Line Line4 
         X1              =   3660
         X2              =   3885
         Y1              =   1035
         Y2              =   1035
      End
   End
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   2670
      Left            =   1470
      Negotiate       =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   9270
      Visible         =   0   'False
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   4710
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
      RowDividerStyle =   5
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
      ColumnCount     =   8
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
         DataField       =   "Bin_Loca"
         Caption         =   "Bin Loc."
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
         DataField       =   "CurrStk"
         Caption         =   "Curr.Stk."
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
      BeginProperty Column06 
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
      BeginProperty Column07 
         DataField       =   "MRP"
         Caption         =   "    MRP"
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
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1230.236
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
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2564.788
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   27
      Left            =   5070
      MaxLength       =   11
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   750
      Width           =   1305
   End
   Begin MSDataGridLib.DataGrid DGGodown 
      Height          =   2100
      Left            =   3840
      Negotiate       =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   7185
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
         Caption         =   "Godown Name"
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
   Begin MSDataGridLib.DataGrid DGMech 
      Height          =   2865
      Left            =   540
      Negotiate       =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   7500
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
   Begin MSDataGridLib.DataGrid DGJob 
      Height          =   3075
      Left            =   -210
      Negotiate       =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   7365
      Visible         =   0   'False
      Width           =   11985
      _ExtentX        =   21140
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "DocId"
         Caption         =   "Job_DocID"
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
         DataField       =   "DispJob_No"
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
      BeginProperty Column07 
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
         BeginProperty Column07 
            ColumnWidth     =   1275.024
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
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
      Index           =   4
      Left            =   8820
      MaxLength       =   10
      TabIndex        =   26
      Top             =   2640
      Width           =   1395
   End
   Begin VB.TextBox txt 
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
      Index           =   25
      Left            =   11340
      MaxLength       =   2
      TabIndex        =   25
      Top             =   1875
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txt 
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
      Left            =   2715
      MaxLength       =   3
      TabIndex        =   30
      Top             =   6870
      Width           =   645
   End
   Begin VB.TextBox txt 
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
      Index           =   23
      Left            =   8820
      MaxLength       =   40
      TabIndex        =   28
      Top             =   3120
      Width           =   2955
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   19
      Left            =   1635
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2910
      Width           =   2970
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   11
      Left            =   1635
      TabIndex        =   11
      Top             =   1230
      Width           =   1890
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   20
      Left            =   1635
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3150
      Width           =   1260
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   8
      Left            =   5070
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1470
      Width           =   2055
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   13
      Left            =   5070
      MaxLength       =   25
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1230
      Width           =   2055
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   14
      Left            =   1635
      TabIndex        =   13
      Top             =   1470
      Width           =   1890
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   21
      Left            =   3210
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3150
      Width           =   1395
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   22
      Left            =   4890
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3150
      Width           =   1650
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   12
      Left            =   5070
      MaxLength       =   20
      TabIndex        =   10
      Top             =   990
      Width           =   2055
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   9
      Left            =   1635
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1710
      Width           =   1890
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   7
      Left            =   1635
      MaxLength       =   11
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   750
      Width           =   1890
   End
   Begin VB.TextBox txt 
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
      Index           =   6
      Left            =   1635
      TabIndex        =   6
      Top             =   510
      Width           =   1890
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   10
      Left            =   1635
      TabIndex        =   9
      Top             =   990
      Width           =   1890
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   18
      Left            =   1635
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2670
      Width           =   5625
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   17
      Left            =   1635
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2430
      Width           =   5625
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   16
      Left            =   1635
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2190
      Width           =   5625
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   15
      Left            =   1635
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1950
      Width           =   5625
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14010
      _ExtentX        =   24712
      _ExtentY        =   661
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   0
      Left            =   9405
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   525
      Width           =   2280
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
      Left            =   2520
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4395
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txt 
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
      Index           =   1
      Left            =   9765
      TabIndex        =   3
      ToolTipText     =   "Press C-> Cash or R-> Credit"
      Top             =   1050
      Width           =   1200
   End
   Begin VB.TextBox txt 
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
      Index           =   26
      Left            =   10395
      MaxLength       =   2
      TabIndex        =   24
      Top             =   1875
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txt 
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
      Index           =   5
      Left            =   8820
      MaxLength       =   11
      TabIndex        =   27
      Top             =   2880
      Width           =   1395
   End
   Begin VB.TextBox txt 
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
      Left            =   9765
      MaxLength       =   11
      TabIndex        =   4
      Top             =   1290
      Width           =   1560
   End
   Begin VB.TextBox txt 
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
      Index           =   3
      Left            =   10425
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1530
      Width           =   900
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3285
      Left            =   15
      TabIndex        =   29
      Top             =   3525
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   5794
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   36
      BackColorFixed  =   13623520
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      GridColorFixed  =   12632319
      FocusRect       =   0
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "dddd"
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
      _Band(0).Cols   =   36
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
      Height          =   255
      Index           =   10
      Left            =   5820
      TabIndex        =   128
      Top             =   6870
      Width           =   180
   End
   Begin VB.Label LblNetValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   6045
      TabIndex        =   127
      Top             =   6870
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Value +  Tax"
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
      Index           =   34
      Left            =   4245
      TabIndex        =   126
      Top             =   6870
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Date........."
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
      Left            =   3720
      TabIndex        =   77
      Top             =   765
      Width           =   1485
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   10860
      TabIndex        =   76
      Top             =   1890
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Left            =   9945
      TabIndex        =   75
      Top             =   1950
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No."
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
      Left            =   8460
      TabIndex        =   74
      Top             =   2670
      Width           =   285
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
      Index           =   24
      Left            =   2580
      TabIndex        =   71
      Top             =   6870
      Width           =   180
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Is This Slip is Final (yes/No)"
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
      Index           =   24
      Left            =   210
      TabIndex        =   0
      Top             =   6870
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mechanic"
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
      Left            =   7470
      TabIndex        =   70
      Top             =   3120
      Width           =   780
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
      Index           =   28
      Left            =   75
      TabIndex        =   69
      Top             =   2910
      Width           =   345
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
      Index           =   27
      Left            =   75
      TabIndex        =   68
      Top             =   2190
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No.........."
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
      Left            =   3720
      TabIndex        =   67
      Top             =   1230
      Width           =   1455
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
      Left            =   75
      TabIndex        =   66
      Top             =   3150
      Width           =   870
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
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
      Left            =   75
      TabIndex        =   65
      Top             =   1230
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Type......."
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
      Height          =   270
      Index           =   37
      Left            =   3720
      TabIndex        =   64
      Top             =   1470
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name"
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
      Left            =   75
      TabIndex        =   63
      Top             =   1950
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration No."
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
      Left            =   75
      TabIndex        =   62
      Top             =   990
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Serial No."
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
      Left            =   75
      TabIndex        =   61
      Top             =   1470
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(O)"
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
      Left            =   1215
      TabIndex        =   60
      Top             =   3150
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(R)"
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
      Left            =   2925
      TabIndex        =   59
      Top             =   3150
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(M)"
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
      Left            =   4620
      TabIndex        =   58
      Top             =   3150
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No......."
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
      Left            =   3720
      TabIndex        =   57
      Top             =   990
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard No."
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
      Left            =   75
      TabIndex        =   56
      Top             =   510
      Width           =   1050
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kms Reading"
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
      Left            =   75
      TabIndex        =   55
      Top             =   1725
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JC Open Date"
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
      Left            =   75
      TabIndex        =   54
      Top             =   750
      Width           =   1185
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   1425
      Left            =   8520
      Top             =   450
      Width           =   3240
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Prefix"
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
      Left            =   9765
      TabIndex        =   53
      Top             =   1530
      Width           =   675
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
      Index           =   25
      Left            =   9285
      TabIndex        =   51
      Top             =   525
      Width           =   75
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOC ID"
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
      Left            =   8610
      TabIndex        =   50
      Top             =   525
      Width           =   675
   End
   Begin VB.Label LblSite 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code"
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
      Left            =   10320
      TabIndex        =   49
      Top             =   825
      Width           =   840
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
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
      Left            =   8610
      TabIndex        =   48
      Top             =   825
      Width           =   675
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doc. Type"
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
      Left            =   8610
      TabIndex        =   45
      ToolTipText     =   "Press C-> Cash or R-> Credit"
      Top             =   1050
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total No. of Items"
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
      Height          =   255
      Index           =   7
      Left            =   7095
      TabIndex        =   44
      Top             =   6870
      Width           =   1470
   End
   Begin VB.Label LblQty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.000"
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
      Height          =   255
      Left            =   10935
      TabIndex        =   43
      Top             =   6870
      Width           =   465
   End
   Begin VB.Label LblIVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   8895
      TabIndex        =   42
      Top             =   6870
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity"
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
      Height          =   255
      Index           =   25
      Left            =   9450
      TabIndex        =   41
      Top             =   6870
      Width           =   1170
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
      Height          =   255
      Index           =   28
      Left            =   8670
      TabIndex        =   40
      Top             =   6870
      Width           =   180
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
      Height          =   255
      Index           =   27
      Left            =   10725
      TabIndex        =   39
      Top             =   6870
      Width           =   180
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warr Claim"
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
      Left            =   7470
      TabIndex        =   37
      Top             =   2670
      Width           =   975
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   7470
      TabIndex        =   36
      Top             =   2880
      Width           =   405
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
      Index           =   90
      Left            =   9630
      TabIndex        =   35
      Top             =   1080
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
      Index           =   92
      Left            =   9630
      TabIndex        =   34
      Top             =   1560
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
      Index           =   93
      Left            =   9630
      TabIndex        =   33
      Top             =   1320
      Width           =   75
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   8595
      TabIndex        =   32
      Top             =   1290
      Width           =   405
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No."
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
      Left            =   8580
      TabIndex        =   31
      Top             =   1530
      Width           =   840
   End
End
Attribute VB_Name = "frmRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PubRequisitionType$             ' Used For Various Reuisition Type Like - "Workshop","Stores","Return"
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim RsJob As ADODB.Recordset
Dim RsMech As ADODB.Recordset
Dim RsGodown As ADODB.Recordset
Dim Master As ADODB.Recordset

Dim mVType$, mVPrefix$
Dim mSearchCode$
Dim LockYN As Boolean
Dim ExitCtrl As Boolean
' Under observation
Dim VoucherEditFlag As Boolean                  ' Used for whether we can edit voucher no or not
' End Under observation
Dim ListArray As Variant
Dim mListItem As ListItem
Dim mServCatg$

Dim mCheckNegetiveStockSiteWise As Boolean

Private mDisSprMRP As Single, mDisSprTB As Single, mDisSprTP As Single
Private mDisOilMRP As Single, mDisOilTB As Single, mDisOilTP As Single

Private Const BackColorSelEnter$ = &HEBB7EC '&HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const DocID As Byte = 0                 ' Doc.ID
Private Const DocType As Byte = 1               ' Document Type
Private Const VDate As Byte = 2                 ' Date
Private Const SerialNo As Byte = 3              ' Serial No.
Private Const WarrNo As Byte = 4                ' Warranty Claim No
Private Const WarrDate As Byte = 5              ' Warranty Date
Private Const JobNo As Byte = 6                 ' JC No
Private Const JobDt As Byte = 7                 ' JC Date
Private Const SrvType As Byte = 8               ' Service Type
Private Const CurrentKMS As Byte = 9            ' Kilometer Reading
Private Const VehRegNo As Byte = 10             ' Reg. No
Private Const Model As Byte = 11                ' Model
Private Const Chassis As Byte = 12              ' Chassis NO
Private Const Engine As Byte = 13               ' Engine No
Private Const VehSrlNo As Byte = 14             ' Veh. Srl. No
Private Const OwnerName As Byte = 15            ' Owner
Private Const Address1 As Byte = 16             ' Address1
Private Const Address2 As Byte = 17             ' Address2
Private Const Address3 As Byte = 18             ' Address3
Private Const City As Byte = 19                 ' City
Private Const PhoneOff As Byte = 20             ' Phone Office
Private Const PhoneResi As Byte = 21            ' Phone Resi
Private Const Mobile As Byte = 22               ' Mobile
Private Const Mechanic As Byte = 23             ' Mechanic
Private Const FinSlipYN As Byte = 24            ' Final Slip(Y/N)
Private Const WarrType As Byte = 25             ' Warranty Claim No
Private Const WarrYear As Byte = 26             ' Warranty Claim No
Private Const JobClDt As Byte = 27
Private Const RateType As Byte = 28
Private Const SoldDate As Byte = 29

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_PNo As Byte = 1               ' Part No
Private Const Col_Unit As Byte = 2              ' Unit
Private Const Col_Purpose As Byte = 3          ' Purpose
Private Const Col_MRP As Byte = 4               ' MRP Yes/No
Private Const Col_Taxable As Byte = 5           ' Taxable Yes/No
Private Const Col_QtyReq As Byte = 6            ' Qty Required
Private Const Col_QtyIss As Byte = 7            ' Qty Issued
Private Const Col_QtyRet As Byte = 8            ' Qty Return
Private Const Col_Rate As Byte = 9              ' Rate
Private Const Col_MRPRate As Byte = 10          ' MRP Rate
Private Const Col_Amt As Byte = 11              ' Amt
Private Const Col_DiscPer As Byte = 12          ' Disc. %
Private Const Col_DiscAmt As Byte = 13          ' Disc. Amt.
Private Const Col_TaxPer As Byte = 14           ' Tax Per.
Private Const Col_TaxAmt As Byte = 15           ' Tax Amt.
Private Const Col_SatPer As Byte = 16           ' Tax Per.
Private Const Col_SatAmt As Byte = 17           ' Tax Amt.
Private Const Col_ItemVal As Byte = 18          ' Item Value
Private Const Col_GodownCode As Byte = 19       ' Godown Code
Private Const Col_Godown As Byte = 20           ' Godown
Private Const Col_LubCat As Byte = 21           ' Lubricant Category
Private Const Col_RemWs As Byte = 22            ' Remark Workshop
Private Const Col_RemStores As Byte = 23        ' Remark Stores
Private Const Col_PName As Byte = 24            ' Part Name
Private Const Col_LName As Byte = 25            ' Local Name
Private Const Col_MRPStkTB As Byte = 26         ' MRP Stock TB
Private Const Col_MRPStkTP As Byte = 27         ' MRP Qty TP
Private Const Col_TBStk As Byte = 28            ' Taxbale Qty
Private Const Col_TPStk As Byte = 29            ' Tax Paid Qty
Private Const Col_TBRate As Byte = 30           ' Taxbale Rate
Private Const Col_TPRate As Byte = 31           ' Tax Paid Rate
Private Const Col_Bin As Byte = 32              ' Bin
Private Const Col_LastRate As Byte = 33         ' Last Purchase Rate
Private Const Col_HPRate As Byte = 34           ' High Purchase Rate
Private Const Col_LPRate As Byte = 35           ' Low Purchase Rate
Private Const Col_PartGrade As Byte = 36        ' Part Grade (Used for Oil Item)
Private Const Col_DiscFact As Byte = 37         ' Discount Factor (Used for Disc%)
Private Const Col_EffectDate As Byte = 38       ' MRP Effective Date/TB Effective Date
Private Const Col_SrlNo As Byte = 39            ' SP_Stock SrlNo (DocID+SrlNo)
Private Const Col_PurDocId As Byte = 40
Private Const Col_PurDate As Byte = 41
Private Const Col_Supplier As Byte = 42

Private Const Col_DepItem As Byte = 43
Private Const Col_DepitemPer As Byte = 44
Private Const Col_DepCode As Byte = 45
Private Const Col_DepPer As Byte = 46
Private Const Col_DepAmt As Byte = 47
Private Const Col_InsuranceAmt As Byte = 48
Private Const Col_DiffPeried As Byte = 49

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName$

Dim mSatYn As Boolean

Dim rsTaxPer As ADODB.Recordset

Private Sub Disp_Text(Enb As Boolean)
    txt(DocType).Enabled = Enb
    txt(VDate).Enabled = Enb
    txt(SerialNo).Enabled = Enb
    txt(WarrNo).Enabled = Enb
    txt(WarrType).Enabled = Enb
    txt(WarrYear).Enabled = Enb
    txt(WarrDate).Enabled = Enb
    txt(JobNo).Enabled = Enb
    txt(Mechanic).Enabled = Enb
    txt(FinSlipYN).Enabled = Enb
End Sub

'* Used for clear all text boxes used in the form
Private Sub BlankText()
Dim I As Integer
    For I = 0 To txt.Count - 1
        txt(I).TEXT = ""
    Next I
    txt(DocID).Tag = ""
    LblDiv.CAPTION = "Division : "
    LblSite.CAPTION = "Site Code : "
    LblVPrefix.CAPTION = ""
    LblIVal.CAPTION = ""
    LblQty.CAPTION = ""

    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End Sub

'Used for Grid Coloumn Width Initialization on various Option Like (Workshop/Store/Return)
Private Sub Grid_IniColWidth( _
        WCol_MRP As Integer, WCol_Taxable As Integer, _
        WCol_QtyRet As Integer, WCol_Rate As Integer, WCol_Amt As Integer, _
        WCol_DiscPer As Integer, WCol_DiscAmt As Integer, WCol_ItemVal As Integer, _
        WCol_Godown As Integer, WCol_LubCat As Integer, WCol_Purpose As Integer, _
        WCol_RemWs As Integer, WCol_RemStores As Integer, WCol_PurDocId As Integer, WCol_PurDate As Integer)
    With FGrid
        .ColWidth(Col_MRP) = WCol_MRP
        .ColWidth(Col_Taxable) = WCol_Taxable
        .ColWidth(Col_QtyRet) = WCol_QtyRet
        .ColWidth(Col_Rate) = WCol_Rate
        .ColWidth(Col_Amt) = WCol_Amt
        .ColWidth(Col_DiscPer) = WCol_DiscPer
        .ColWidth(Col_DiscAmt) = WCol_DiscAmt
        .ColWidth(Col_ItemVal) = WCol_ItemVal
        .ColWidth(Col_Godown) = WCol_Godown
        .ColWidth(Col_LubCat) = WCol_LubCat
        .ColWidth(Col_Purpose) = WCol_Purpose
        .ColWidth(Col_RemWs) = WCol_RemWs
        .ColWidth(Col_RemStores) = WCol_RemStores
        .ColWidth(Col_PurDocId) = WCol_PurDocId
        .ColWidth(Col_PurDate) = WCol_PurDate
    End With
End Sub

'* Used for intialize grid columns
Private Sub Grid_Ini()
Dim FGridLeft As Long, FGridwidth As Long
    
    With FGrid
        .left = Me.left ' + 45
        .width = Me.width - 90
        .top = 3525
        .RowHeightMin = PubGridRowHeight
        .Cols = 50
        
        .TextMatrix(0, Col_SrNo) = "S.No"
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 450

        .TextMatrix(0, Col_PNo) = "Part No"
        .ColAlignment(Col_PNo) = flexAlignLeftCenter
        .ColWidth(Col_PNo) = 2430

        .TextMatrix(0, Col_PName) = "Part Name"
        .ColAlignment(Col_PName) = flexAlignLeftCenter
        .ColWidth(Col_PName) = 2500

        .TextMatrix(0, Col_Unit) = "Unit"
        .ColAlignment(Col_Unit) = flexAlignLeftCenter
        .ColWidth(Col_Unit) = 435

        .TextMatrix(0, Col_MRP) = "MRP"
        .ColAlignment(Col_MRP) = flexAlignLeftCenter
'        .ColWidth(Col_MRP) = 450

        .TextMatrix(0, Col_Taxable) = "Tax"
        .ColAlignment(Col_Taxable) = flexAlignLeftCenter
'        .ColWidth(Col_Taxable) = 420

        .TextMatrix(0, Col_QtyReq) = "Qty Req"
        .ColAlignmentFixed(Col_QtyReq) = flexAlignRightCenter
        .ColWidth(Col_QtyReq) = 0 '960

        .TextMatrix(0, Col_QtyIss) = "Qty Issue"
        .ColAlignmentFixed(Col_QtyIss) = flexAlignRightCenter
        .ColWidth(Col_QtyIss) = 810

        .TextMatrix(0, Col_QtyRet) = "Qty Return"
        .ColAlignmentFixed(Col_QtyRet) = flexAlignRightCenter
'        .ColWidth(Col_QtyRet) = 960

        .TextMatrix(0, Col_Rate) = "Rate"
        .ColAlignmentFixed(Col_Rate) = flexAlignRightCenter
'        .ColWidth(Col_Rate) = 870

        .TextMatrix(0, Col_MRPRate) = "MRP Rate"
        .ColAlignmentFixed(Col_MRPRate) = flexAlignRightCenter
        .ColWidth(Col_MRPRate) = 0

        .TextMatrix(0, Col_Amt) = "Amount"
        .ColAlignmentFixed(Col_Amt) = flexAlignRightCenter
'        .ColWidth(Col_Amt) = 1065

        .TextMatrix(0, Col_DiscPer) = "Disc%"
        .ColAlignmentFixed(Col_DiscPer) = flexAlignRightCenter
'        .ColWidth(Col_DiscPer) = 555

        .TextMatrix(0, Col_DiscAmt) = "Disc.Amt"
        .ColAlignmentFixed(Col_DiscAmt) = flexAlignRightCenter
'        .ColWidth(Col_DiscAmt) = 840

        If PubVATYN = 1 Then
            .TextMatrix(0, Col_TaxPer) = "TaxPer"
            .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
            .ColWidth(Col_TaxPer) = 840
            
            .TextMatrix(0, Col_TaxAmt) = "TaxAmt"
            .ColAlignmentFixed(Col_TaxAmt) = flexAlignRightCenter
            .ColWidth(Col_TaxAmt) = 840
                        
            If PubSatYn = 1 Then
                .TextMatrix(0, Col_SatPer) = "SAT %"
                .ColAlignmentFixed(Col_SatPer) = flexAlignRightCenter
                .ColWidth(Col_SatPer) = 840
                
                .TextMatrix(0, Col_SatAmt) = "SAT Amt"
                .ColAlignmentFixed(Col_SatAmt) = flexAlignRightCenter
                .ColWidth(Col_SatAmt) = 840
            Else
                .ColWidth(Col_SatPer) = 0
                .ColWidth(Col_SatAmt) = 0
            End If
        Else
            .TextMatrix(0, Col_TaxPer) = ""
            .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
            .ColWidth(Col_TaxPer) = 0
            
            .TextMatrix(0, Col_TaxAmt) = ""
            .ColAlignmentFixed(Col_TaxAmt) = flexAlignRightCenter
            .ColWidth(Col_TaxAmt) = 0
            
            .ColWidth(Col_SatPer) = 0
            .ColWidth(Col_SatAmt) = 0
        End If

        .TextMatrix(0, Col_ItemVal) = "Item Value"
        .ColAlignmentFixed(Col_ItemVal) = flexAlignRightCenter
'        .ColWidth(Col_ItemVal) = 1095

        .TextMatrix(0, Col_LName) = "Local Name"
        .ColAlignment(Col_LName) = flexAlignLeftCenter
        .ColWidth(Col_LName) = 2000




        .TextMatrix(0, Col_GodownCode) = "Godown Code"
        .ColAlignment(Col_GodownCode) = flexAlignLeftCenter
        .ColWidth(Col_GodownCode) = 0

        .TextMatrix(0, Col_Godown) = "Godown"
        .ColAlignment(Col_Godown) = flexAlignLeftCenter
'        .ColWidth(Col_Godown) = 1200

        .TextMatrix(0, Col_LubCat) = "Lubricant Cat."
        .ColAlignment(Col_LubCat) = flexAlignLeftCenter
'        .ColWidth(Col_LubCat) = 1200

        .TextMatrix(0, Col_Purpose) = "Purpose"
        .ColAlignment(Col_Purpose) = flexAlignLeftCenter
'        .ColWidth(Col_Purpose) = 1200

        .TextMatrix(0, Col_RemWs) = "Issue Remarks"
        .ColAlignment(Col_RemWs) = flexAlignLeftCenter
'        .ColWidth(Col_RemWs) = 1300

        .TextMatrix(0, Col_RemStores) = "Return Remarks"
        .ColAlignment(Col_RemStores) = flexAlignLeftCenter
'        .ColWidth(Col_RemStores) = 1300

        .TextMatrix(0, Col_MRPStkTB) = "Current Stock TB"
        .ColWidth(Col_MRPStkTB) = 0

        .TextMatrix(0, Col_MRPStkTP) = "Current Stock TP"
        .ColWidth(Col_MRPStkTP) = 0

        .TextMatrix(0, Col_TBStk) = "Taxable Qty"
        .ColWidth(Col_TBStk) = 0

        .TextMatrix(0, Col_TPStk) = "Tax Paid Qty"
        .ColWidth(Col_TPStk) = 0

        .TextMatrix(0, Col_TBRate) = "Taxbale Rate"
        .ColWidth(Col_TBRate) = 0

        .TextMatrix(0, Col_TPRate) = "Tax Paid Rate"
        .ColWidth(Col_TPRate) = 0

        .TextMatrix(0, Col_Bin) = "Bin"
        .ColWidth(Col_Bin) = 600

        .TextMatrix(0, Col_LastRate) = "Last Purchase Rate"
        .ColWidth(Col_LastRate) = 0

        .TextMatrix(0, Col_HPRate) = "High Purchase Rate"
        .ColWidth(Col_HPRate) = 0

        .TextMatrix(0, Col_LPRate) = "Low Purchase Rate"
        .ColWidth(Col_LPRate) = 0

        .TextMatrix(0, Col_PartGrade) = "Part Grade"
        .ColWidth(Col_PartGrade) = 0

        .TextMatrix(0, Col_DiscFact) = "Discount Factor"
        .ColWidth(Col_DiscFact) = 0

        .TextMatrix(0, Col_EffectDate) = "Rate Effective Date"
        .ColWidth(Col_EffectDate) = 0

        .TextMatrix(0, Col_SrlNo) = "Stock Srl No"
        .ColWidth(Col_SrlNo) = 0
        
        .TextMatrix(0, Col_PurDocId) = "Purch Doc No"
        .ColWidth(Col_PurDocId) = 0
        
        .TextMatrix(0, Col_PurDate) = "Purch Doc Date"
        .ColWidth(Col_PurDate) = 0
        If RSOJPR = True Then
            .TextMatrix(0, Col_Supplier) = "Supplier"
            .ColWidth(Col_Supplier) = 2500
        End If
        
        

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
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    FGridLeft = FGrid.left
    FGridwidth = FGrid.width
    DGPart.width = FGridwidth: DGPart.left = FGridLeft: DGPart.top = mTopScale: DGPart.height = 2900
    DGGodown.left = FGridLeft: DGGodown.top = DGPart.top: DGGodown.height = 2350
    FrmDetail.width = 6285: FrmDetail.left = 5595: FrmDetail.top = 405: FrmDetail.height = 2130
    DGMech.left = 60: DGMech.top = 435
    DGJob.left = FGridLeft: DGJob.top = FGrid.top: DGJob.height = FGrid.height
    DGGodown.left = FGridLeft: DGGodown.top = DGPart.top: DGGodown.height = 2350

    If PubRequisitionType = "Workshop" Or PubRequisitionType = "Store" Then
'        Grid_IniColWidth 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1200, 1200, 1300, 0
        Grid_IniColWidth 450, 420, 0, 870, 945, 555, 840, 1095, 1170, 1200, 765, 1300, 0, 2000, 1500
'        Grid_IniColWidth 450, 420, 960, 0, 870, 1065, 555, 840, 1095, 1200, 1200, 1200, 1300, 1300
    ElseIf PubRequisitionType = "Return" Then
        Grid_IniColWidth 420, 420, 810, 870, 945, 555, 840, 1095, 1170, 1200, 765, 0, 1300, 2000, 1500
    End If
'    With DGPart
'        .Columns(6).width = 2564.788
'        .Columns(5).width = 1005.165
'        .Columns(4).width = 1005.165
'        .Columns(3).width = 1005.165
'        .Columns(2).width = 494.9292
'        .Columns(1).width = 3225.26
'        .Columns(0).width = 2759.811 '1950.236
'    End With
End Sub

Private Sub Grid_Hide()
    If DGJob.Visible = True Then DGJob.Visible = False
    If DGGodown.Visible = True Then DGGodown.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGPart.Visible = True Then DGPart.Visible = False
    If DgRateType.Visible = False Then DgRateType.Visible = False
End Sub

Private Sub MoveRec()
Dim Rst As ADODB.Recordset, I As Integer, TmpStr$
Dim RstTmp As ADODB.Recordset
Dim WrkEdit As Boolean
On Error GoTo ELoop
    FrmDetail.Visible = False
    If Master.RecordCount > 0 Then
        WrkEdit = True
        
        Set Rst = GCn.Execute("Select P.Part_Name ,P.Local_Name ,P.Unit ,P.MRP ,P.TB_SRate ,P.TP_SRate ," _
            & "P.MRP_Effect_Dt ,P.TB_Effect_Dt ,P.Part_Grade ,Disc_Factor ," _
            & "P.Cur_MRP_TBStk, P.Cur_MRP_TPStk, P.Cur_TB_Stk, P.Cur_TP_Stk, " _
            & "P.Bin_Loca ,P.High_Pur_Rate ,P.Low_Pur_Rate ," _
            & "Emp_Mast.Emp_Name,Godown.God_Name, RateType.Description as RateTypeDescription, RateType.VariationPer,SP_Stock.*,h.Delivery_Date as SOldDate   " _
            & "From (((((SP_Stock Left Join Part P On SP_Stock.Part_No=P.Part_No and P.Div_Code = left(SP_Stock.DocID,1)) " _
            & "Left Join Emp_Mast on SP_Stock.Mech_Code=Emp_Mast.Emp_Code) " _
            & "Left Join Godown on SP_Stock.Godown=Godown.God_Code) " _
            & "Left Join RateType on SP_Stock.RateType=RateType.Code)Left join Job_card j on SP_Stock.job_docid=j.docid )Left Join HisCard H on J.CardNo=H.CardNo " _
            & "Where SP_Stock.DocID='" & Master!SearchCode & "' Order By SP_Stock.Srl_No")
            
        FGrid.Redraw = False
        If Rst.RecordCount > 0 Then
            If GCn.Execute("Select S.Serv_Catg From Job_card J Left Join Service_type S on J.Serv_type=S.Serv_Type Where J.DocId='" & Rst!job_docid & "'").RecordCount > 0 Then
                mServCatg = GCn.Execute("Select S.Serv_Catg From Job_card J Left Join Service_type S on J.Serv_type=S.Serv_Type Where J.DocId='" & Rst!job_docid & "'").Fields(0).Value
            Else
                mServCatg = "****"
            End If
            
            txt(DocID).TEXT = Master!SearchCode
            mSearchCode = txt(DocID)
            LblDiv.CAPTION = "Division : " & left(Rst!DocID, 1)
            LblSite.CAPTION = "Site Code : " & Rst!Site_Code
            If PubBackEnd = "A" Then
                mSatYn = IIf(VNull(Master!SAT_YN) = 1, True, False)
            Else
                mSatYn = IIf(VNull(Master!SAT_YN) = True, True, False)
            End If
            IniGrid_Vat
            mVType = Rst!V_Type
            If mVType = "W_RW" Then
                txt(DocType).TEXT = "Warranty"
            ElseIf mVType = "W_RG" Or mVType = "W_RGO" Then
                txt(DocType).TEXT = "General"
            End If
            
            txt(VDate).TEXT = Rst!V_DATE
            LblVPrefix.CAPTION = DeCodeDocID(Rst!DocID, Document_Prefix)
            txt(SerialNo).TEXT = Rst!V_NO
            txt(WarrNo).TEXT = IIf(IsNull(Rst!claim_no), "", Rst!claim_no)
            txt(WarrYear).TEXT = IIf(IsNull(Rst!claim_YearPrefix), "", Rst!claim_YearPrefix)
            txt(WarrType).TEXT = IIf(IsNull(Rst!claim_type), "", Rst!claim_type)
            txt(WarrDate).TEXT = IIf(IsNull(Rst!Claim_Date), "", Rst!Claim_Date)
            txt(Mechanic).Tag = IIf(IsNull(Rst!mech_code), "", Rst!mech_code)
            txt(Mechanic).TEXT = IIf(IsNull(Rst!Emp_Name), "", Rst!Emp_Name)
            txt(FinSlipYN).TEXT = IIf(Rst!TrnCompLete_YN = 0, "No", "Yes")
            txt(JobNo).Tag = Rst!job_docid
            'Txt(JobNo).Text = Trim(Mid(Master!Job_DocID, 8, 5)) + CStr(Trim(Right(Master!Job_DocID, 8)))
            txt(JobNo).TEXT = CStr(Trim(Right(Rst!job_docid, 8)))
            txt(JobClDt).TEXT = IIf(IsNull(Rst!V_DATE2), "", Rst!V_DATE2)
            
txt(SoldDate).TEXT = IIf(IsNull(Rst!SoldDate), "", Rst!SoldDate)

            JobCardDetail txt(JobNo).Tag
            
            FGrid.Rows = 1
            I = 1
            Do Until Rst.EOF
                FGrid.AddItem ""
                With FGrid
                    .TextMatrix(I, Col_SrNo) = I
                    .TextMatrix(I, Col_PNo) = Rst!Part_No
                    .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                    .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 0, "No", "Yes")
                    .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 0, "No", "Yes")
                    .TextMatrix(I, Col_QtyReq) = Format(Rst!Qty_Doc, "0.000")
                    .TextMatrix(I, Col_QtyIss) = IIf(Rst!Qty_Iss = 0, "", Format(Rst!Qty_Iss, "0.000"))
                    .TextMatrix(I, Col_QtyRet) = IIf(Rst!Qty_Ret = 0, "", Format(Rst!Qty_Ret, "0.000"))
                    .TextMatrix(I, Col_Rate) = IIf(Rst!Rate = 0, "", Format(Rst!Rate, IIf(UCase(left(PubComp_Name, 3)) = "LMP", "0.000", "0.000")))
                    .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP, "0.00")
                    .TextMatrix(I, Col_Amt) = Format(Rst!Amount, "0.00") 'Format(((Rst!Qty_Iss - Rst!Qty_ret) * Rst!Rate), "0.00")
                    .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per, "0.00")
                    .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt, "0.00")
                    If PubVATYN = 1 Then
                        .TextMatrix(I, Col_TaxPer) = Format(Rst!TaxPer, "0.00")
                        .TextMatrix(I, Col_TaxAmt) = Format(Rst!TaxAmt, "0.00")
                        
                        If mSatYn Then
                            .TextMatrix(I, Col_SatPer) = Format(Rst!SatPer, "0.00")
                            .TextMatrix(I, Col_SatAmt) = Format(Rst!SatAmt, "0.00")
                        End If
                    End If
                    .TextMatrix(I, Col_ItemVal) = Format(Rst!Net_Amt, "0.00")
                    .TextMatrix(I, Col_GodownCode) = Rst!Godown
                    .TextMatrix(I, Col_Godown) = IIf(IsNull(Rst!God_Name), "", Rst!God_Name)
                    
        
                     .TextMatrix(I, Col_DepItem) = IIf(IsNull(Rst!Dep_Item), "", Rst!Dep_Item)
                     .TextMatrix(I, Col_DepCode) = IIf(IsNull(Rst!Dep_Code), "", Rst!Dep_Code)
                     .TextMatrix(I, Col_DepitemPer) = Format(Rst!DepitemPer, "0.00")
                     .TextMatrix(I, Col_DepPer) = Format(Rst!DepPer, "0.00")
                     .TextMatrix(I, Col_DepAmt) = Format(Rst!DepAmt, "0.00")
                     
                     .TextMatrix(I, Col_InsuranceAmt) = Format(Rst!InsuranceAmt, "0.00")
                     .TextMatrix(I, Col_DiffPeried) = Format(Rst!DiffPeried, "0.00")
                     
                     
                    'O- >Oil Filter, F- >Fuel Filter,E- >Engine Oil,G- >Gear Oil,R- >Rear Axle Oil,A- >Front Axle Oil,S- >Steering Oil,N- > N.A.
                    If Rst!Lub_Category = "O" Then
                        TmpStr = "Oil Filter"
                    ElseIf Rst!Lub_Category = "F" Then
                        TmpStr = "Fuel Filter"
                    ElseIf Rst!Lub_Category = "E" Then
                        TmpStr = "Engine Oil"
                    ElseIf Rst!Lub_Category = "G" Then
                        TmpStr = "Gear Oil"
                    ElseIf Rst!Lub_Category = "R" Then
                        TmpStr = "Rear Axle Oil"
                    ElseIf Rst!Lub_Category = "A" Then
                        TmpStr = "Front Axle Oil"
                    ElseIf Rst!Lub_Category = "S" Then
                        TmpStr = "Steering Oil"
                    ElseIf Rst!Lub_Category = "N" Then
                        TmpStr = "N.A."
                    Else
                        TmpStr = ""
                    End If
                    .TextMatrix(I, Col_LubCat) = TmpStr
                    'P- >PDI,F- >Free Service, C- >Chargable,W- >Warranty,O- >Company Vehicle,L- >Complementary
                    If Rst!Purpose = "P" Then
                        TmpStr = "PDI"
                    ElseIf Rst!Purpose = "F" Then
                        TmpStr = "Free Service"
                    ElseIf Rst!Purpose = "C" Then
                        TmpStr = "Charge"
                    ElseIf Rst!Purpose = "W" Then
                        TmpStr = "Warranty"
                    ElseIf Rst!Purpose = "O" Then
                        TmpStr = "Company Vehicle"
                    ElseIf Rst!Purpose = "L" Then
                        TmpStr = "Complementary"
                    ElseIf Rst!Purpose = "A" Then
                        TmpStr = "AMC"
                    Else
                        TmpStr = ""
                    End If
                    .TextMatrix(I, Col_Purpose) = TmpStr

                    TmpStr = IIf(IsNull(Rst!Remark), Space(30), Rst!Remark)
                    .TextMatrix(I, Col_RemWs) = Trim(left(TmpStr, 20))
                    .TextMatrix(I, Col_RemStores) = Trim(Right(TmpStr, 10))
                    .TextMatrix(I, Col_PName) = IIf(IsNull(Rst!Part_Name), "", Rst!Part_Name)
                    .TextMatrix(I, Col_LName) = IIf(IsNull(Rst!Local_Name), "", Rst!Local_Name)
                    .TextMatrix(I, Col_MRPStkTB) = IIf(IsNull(Rst!Cur_MRP_TbStk), "", Rst!Cur_MRP_TbStk)
                    .TextMatrix(I, Col_MRPStkTP) = IIf(IsNull(Rst!Cur_MRP_TPStk), "", Rst!Cur_MRP_TPStk)
                    .TextMatrix(I, Col_TBStk) = IIf(IsNull(Rst!Cur_TB_STk), "", Rst!Cur_TB_STk)
                    .TextMatrix(I, Col_TPStk) = IIf(IsNull(Rst!Cur_TP_Stk), "", Rst!Cur_TP_Stk)
                    .TextMatrix(I, Col_TBRate) = IIf(IsNull(Rst!TB_SRate), "", Rst!TB_SRate)
                    .TextMatrix(I, Col_TPRate) = IIf(IsNull(Rst!TP_SRate), "", Rst!TP_SRate)
                    .TextMatrix(I, Col_Bin) = IIf(IsNull(Rst!Bin_Loca), "", Rst!Bin_Loca)
                    .TextMatrix(I, Col_LastRate) = ""
                    .TextMatrix(I, Col_HPRate) = IIf(IsNull(Rst!high_pur_rate), "", Rst!high_pur_rate)
                    .TextMatrix(I, Col_LPRate) = IIf(IsNull(Rst!low_pur_rate), "", Rst!low_pur_rate)
                    .TextMatrix(I, Col_PartGrade) = IIf(IsNull(Rst!Part_Grade), "", Rst!Part_Grade)
                    .TextMatrix(I, Col_DiscFact) = IIf(IsNull(Rst!Disc_Factor), "", Rst!Disc_Factor)
                    .TextMatrix(I, Col_EffectDate) = Format(IIf(Rst!MRP_YN = 1, IIf(IsNull(Rst!MRP_Effect_Dt), "", Rst!MRP_Effect_Dt), IIf(IsNull(Rst!TB_Effect_Dt), "", Rst!TB_Effect_Dt)), "dd/MMM/yyyy")
                    .TextMatrix(I, Col_SrlNo) = Rst!Srl_No
                    .TextMatrix(I, Col_PurDocId) = XNull(Rst!PurDocNo)
                    .TextMatrix(I, Col_PurDate) = XNull(Rst!PurDocDate)
                    If XNull(Rst!PurDocNo) <> "" Then
                        Set RstTmp = GCn.Execute("Select Party_Name from SP_Purch where Party_Doc_No='" & XNull(Rst!PurDocNo) & "'")
                        If RstTmp.RecordCount > 0 Then
                            .TextMatrix(I, Col_Supplier) = XNull(RstTmp!Party_Name)
                        End If
                    End If
                    Set RstTmp = Nothing
                End With
                If WrkEdit Then
                    WrkEdit = IIf(Rst!Qty_Iss > 0, False, True)
                End If
                Rst.MoveNext
                I = I + 1
            Loop
            FGrid.FixedRows = 1
            If PubRequisitionType = "Workshop" Or PubRequisitionType = "Store" Then
                If txt(JobClDt) <> "" Then
                    TopCtrl1.tEdit = False
                    TopCtrl1.tDel = False
                Else
'                    If WrkEdit Then
                        If InStr(Me.TopCtrl1.Tag, "E") <> 0 Then Me.TopCtrl1.tEdit = True
                        If InStr(Me.TopCtrl1.Tag, "D") <> 0 Then Me.TopCtrl1.tDel = True
'                    Else
'                        TopCtrl1.tEdit = False
'                        TopCtrl1.tDel = False
'                    End If
                End If
            ElseIf PubRequisitionType = "Return" Then
                If txt(JobClDt) <> "" Then
                    TopCtrl1.tEdit = False
                Else
                    If InStr(Me.TopCtrl1.Tag, "E") <> 0 Then Me.TopCtrl1.tEdit = True
                End If
            End If
            CountItem
        End If
    Else
        BlankText
    End If
Set Rst = Nothing
    If FGrid.Rows = 1 Then FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    FGrid.Redraw = True
    Grid_Hide
    If PubRequisitionType = "Return" Then
        TopCtrl1.tAdd = False
        TopCtrl1.tDel = False
    End If
    LblNetValue = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_ItemVal)) + Val(FGrid.TextMatrix(FGrid.Row, Col_TaxAmt)), "0.00")
Exit Sub
ELoop:
    CheckError
End Sub
' Used For Checking Duplicate Items in the Grid
Private Function ChkDuplicate() As Boolean
Dim I As Integer, X$, Y$
Dim TmpRst As ADODB.Recordset
Dim sStr As String
Dim Col1 As Byte, Col2 As Byte, Col3 As Byte
    Select Case FGrid.Col
    Case Col_PNo, Col_PName, Col_LName
        Col1 = Col_MRP
        Col2 = Col_Taxable
        Col3 = FGrid.Col
    Case Col_MRP
        Col1 = Col_PNo
        Col2 = Col_Taxable
        Col3 = Col_MRP
    Case Col_Taxable
        Col1 = Col_PNo
        Col2 = Col_MRP
        Col3 = Col_Taxable
    End Select
    If RSOJPR = False Then
        X = UCase(CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col1))) + CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col2))) + CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col3))))
        For I = 1 To FGrid.Rows - 1
            If I = FGrid.Row Then GoTo nxt1
            Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col1))) + CStr(Trim(FGrid.TextMatrix(I, Col2))) + CStr(Trim(FGrid.TextMatrix(I, Col3))))
            If X = Y And Y <> "" Then
                MsgBox "Duplicate Item Not Allowed", vbInformation, "Validation"
                FGrid.Row = FGrid.Row: FGrid.Col = Col3: FGrid.SetFocus
                
                'txtGrid(0).SetFocus
                ChkDuplicate = False
                Exit Function
            End If
nxt1:
        Next
        ChkDuplicate = True
    End If
    If RSOJPR = True Then
        Col1 = Col_PNo
        X = UCase(CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col1))))
        For I = 1 To FGrid.Rows - 1
            If I = FGrid.Row Then GoTo NXT
            Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col1))))
            If X = Y And Y <> "" Then
                MsgBox "Duplicate Item. ", vbInformation, "Validation"
                FGrid.Row = FGrid.Row: FGrid.Col = Col3: FGrid.SetFocus
                ChkDuplicate = False
                Exit Function
            End If
NXT:
        Next
        ChkDuplicate = True
        '*******For Duplicate Parts Checking in Requi***************
        Set TmpRst = GCn.Execute("Select docid from SP_Stock where Job_Docid='" & txt(JobNo).Tag & "' and Part_no='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "'")
        If TmpRst.RecordCount > 0 Then
            For I = 1 To TmpRst.RecordCount
                sStr = sStr & "  " & Right(TmpRst!DocID, 8)
                TmpRst.MoveNext
            Next
            MsgBox "This Part is already issued on this Job in " & sStr & " Requisition No.(s)", vbInformation + vbOKOnly
            Exit Function
        End If
        '**********************************************************
    End If
End Function

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid.Col
    
    Case Col_PNo, Col_PName, Col_LName
        TxtGridValid_PNo
        'Call ChkDuplicate '= False Then TxtGridLeave = False: Exit Function
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
    Case Col_Taxable, Col_MRP
        TxtGridValid_TaxMRP
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        
    Case Col_QtyReq, Col_QtyRet
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.000")
        Amt_Cal
    Case Col_QtyIss
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.000")
        FGrid.TextMatrix(FGrid.Row, Col_QtyReq) = Format(Val(TxtGrid(0).TEXT), "0.000")
        If Val(FGrid.TextMatrix(FGrid.Row, Col_QtyRet)) > 0 Then
            If Val(FGrid.TextMatrix(FGrid.Row, Col_QtyRet)) > Val(FGrid.TextMatrix(FGrid.Row, Col_QtyIss)) Then
                MsgBox "Reurn Qty is Greater than Issue Qty", vbOKOnly, "Check Qty"
                TxtGrid(0).SetFocus: TxtGridLeave = False:  Exit Function
            End If
        End If
        If CheckSprStock(FGrid, FGrid.Row, Col_MRP, Col_Taxable, Col_QtyIss, Col_MRPStkTB, Col_MRPStkTP, Col_TBStk, Col_TPStk) = False Then
            TxtGrid(0) = "": TxtGrid(0).SetFocus: TxtGridLeave = False: Exit Function
        End If
        Amt_Cal
        If RsGodown.RecordCount > 0 Or Trim(FGrid.TextMatrix(FGrid.Row, Col_Godown)) = "" Then
            RsGodown.MoveFirst
            RsGodown.FIND "Code ='" & PubSprWorksGodown & "'"
            FGrid.TextMatrix(FGrid.Row, Col_GodownCode) = RsGodown!Code
            FGrid.TextMatrix(FGrid.Row, Col_Godown) = RsGodown!Name
        End If
    Case Col_Rate
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), IIf(UCase(left(PubComp_Name, 3)) = "LMP", "0.000", "0.000"))
        Amt_Cal
    Case Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
        Amt_Cal
    Case Col_LubCat, Col_Purpose
        TxtGrid(0).TEXT = ListView.SelectedItem.TEXT
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    Case Col_Godown
        TxtGridValid_Godown
    Case Col_RemWs, Col_RemStores
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    Case Col_PurDocId
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    Case Col_PurDate
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
    End Select
    TxtGridLeave = True
    'Important at the time of validating  a control if you are making the visibility of
    'control false forcefully it will generate error
    If ValidateCall = False Then
        FGrid.SetFocus
        TxtGrid(0).Visible = False
    End If
End Function
'* Used for Calculate the Amount
Private Sub Amt_Cal()
Dim mAmount As Double, TaxAmt As Double, DisAmt As Double, OrdDisAmt1 As Double
Dim TTaxAmt As Double, mTaxableAmt As Double
Dim mQty As Double
mQty = Val(FGrid.TextMatrix(FGrid.Row, Col_QtyIss)) - Val(FGrid.TextMatrix(FGrid.Row, Col_QtyRet))
    
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        If UCase(FGrid.TextMatrix(FGrid.Row, Col_MRP)) = "YES" Then
            FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) * mQty), "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) * mQty), "0.00")
        End If
    Else
        If UCase(FGrid.TextMatrix(FGrid.Row, Col_MRP)) = "YES" Then
            FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)) * mQty), "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) * mQty), "0.00")
        End If
    End If
    
    If StrCmp(left(PubComp_Name, 4), "Enar") Then
        If FGrid.TextMatrix(FGrid.Row, Col_Purpose) = "Charge" Or FGrid.TextMatrix(FGrid.Row, Col_Purpose) = "AMC" Then
            FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = Format(((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) * Val(FGrid.TextMatrix(FGrid.Row, Col_DiscPer))) / 100), "0.00")
            FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) - Val(FGrid.TextMatrix(FGrid.Row, Col_DiscAmt))), "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = ""
            FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = ""
        End If
    Else
        If FGrid.TextMatrix(FGrid.Row, Col_Purpose) = "Charge" Then
            FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = Format(((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) * Val(FGrid.TextMatrix(FGrid.Row, Col_DiscPer))) / 100), "0.00")
            FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) - Val(FGrid.TextMatrix(FGrid.Row, Col_DiscAmt))), "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = ""
            FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = ""
        End If
    End If
    'FGrid.TextMatrix(FGrid.Row, Col_Amt)
'      ******************* For Tax in Line File *************************
    If PubVATYN = 1 Then
            If FGrid.TextMatrix(FGrid.Row, Col_TaxPer) <> "" Then
                mAmount = Val(FGrid.TextMatrix(FGrid.Row, Col_Amt))
                DisAmt = Val(FGrid.TextMatrix(FGrid.Row, Col_DiscAmt))
                If StrCmp(left(PubComp_Name, 4), "Enar") Then
                    If FGrid.TextMatrix(FGrid.Row, Col_MRP) = "Yes" And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" And Not (StrCmp(FGrid.TextMatrix(FGrid.Row, Col_Purpose), "Warranty")) Then ' Or StrCmp(FGrid.TextMatrix(FGrid.Row, Col_Purpose), "Amc")) Then
                        If mSatYn Then
                            mTaxableAmt = Format((mAmount - DisAmt) * 100 / (100 + Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) + Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer))), "0.00")
                            FGrid.TextMatrix(FGrid.Row, Col_TaxAmt) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / 100, "0.00")
                            FGrid.TextMatrix(FGrid.Row, Col_SatAmt) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer)) / 100, "0.00")
                        Else
                            FGrid.TextMatrix(FGrid.Row, Col_TaxAmt) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / (100 + Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer))), "0.00")
                        End If
                        If Val(FGrid.TextMatrix(FGrid.Row, Col_ItemVal)) > 0 Then
                            FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_ItemVal)) - Val(FGrid.TextMatrix(FGrid.Row, Col_TaxAmt)) - Val(FGrid.TextMatrix(FGrid.Row, Col_SatAmt)), "0.00")
                        End If
                    'ElseIf (FGrid.TextMatrix(FGrid.Row, Col_MRP) = "No" Or StrCmp(FGrid.TextMatrix(FGrid.Row, Col_Purpose), "Warranty") Or StrCmp(FGrid.TextMatrix(FGrid.Row, Col_Purpose), "Amc")) And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" Then
                    ElseIf (FGrid.TextMatrix(FGrid.Row, Col_MRP) = "No" Or StrCmp(FGrid.TextMatrix(FGrid.Row, Col_Purpose), "Warranty")) And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" Then
                        FGrid.TextMatrix(FGrid.Row, Col_TaxAmt) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / 100, "0.00")
                        If mSatYn Then
                            FGrid.TextMatrix(FGrid.Row, Col_SatAmt) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer)) / 100, "0.00")
                        End If
                    Else
                        FGrid.TextMatrix(FGrid.Row, Col_TaxAmt) = ""
                        FGrid.TextMatrix(FGrid.Row, Col_SatAmt) = ""
                    End If
                Else
                    If FGrid.TextMatrix(FGrid.Row, Col_MRP) = "Yes" And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" And Not (StrCmp(FGrid.TextMatrix(FGrid.Row, Col_Purpose), "Warranty")) Or StrCmp(FGrid.TextMatrix(FGrid.Row, Col_Purpose), "Amc") Then
                        
                        
                        If mSatYn Then
                            mTaxableAmt = Format((mAmount - DisAmt) * 100 / (100 + Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) + Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer))), "0.00")
                            FGrid.TextMatrix(FGrid.Row, Col_TaxAmt) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / 100, "0.00")
                            FGrid.TextMatrix(FGrid.Row, Col_SatAmt) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer)) / 100, "0.00")
                        Else
                            FGrid.TextMatrix(FGrid.Row, Col_TaxAmt) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / (100 + Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer))), "0.00")
                        End If
                        
                        If Val(FGrid.TextMatrix(FGrid.Row, Col_ItemVal)) > 0 Then
                            FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_ItemVal)) - Val(FGrid.TextMatrix(FGrid.Row, Col_TaxAmt)) - Val(FGrid.TextMatrix(FGrid.Row, Col_SatAmt)), "0.00")
                        End If
                    ElseIf (FGrid.TextMatrix(FGrid.Row, Col_MRP) = "No" Or StrCmp(FGrid.TextMatrix(FGrid.Row, Col_Purpose), "Warranty") Or StrCmp(FGrid.TextMatrix(FGrid.Row, Col_Purpose), "Amc")) And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" Then
                        FGrid.TextMatrix(FGrid.Row, Col_TaxAmt) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / 100, "0.00")
                        If mSatYn Then
                            FGrid.TextMatrix(FGrid.Row, Col_SatAmt) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer)) / 100, "0.00")
                        End If
                    Else
                        FGrid.TextMatrix(FGrid.Row, Col_TaxAmt) = ""
                        FGrid.TextMatrix(FGrid.Row, Col_SatAmt) = ""
                    End If
                End If
       End If
       
    End If
    '*******************************************************************
    'Nikhil
    
    LblNetValue = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_ItemVal)) + Val(FGrid.TextMatrix(FGrid.Row, Col_TaxAmt)) + Val(FGrid.TextMatrix(FGrid.Row, Col_SatAmt)), "0.00")
      If Val(FGrid.TextMatrix(FGrid.Row, Col_DepitemPer)) > 0 And Val(FGrid.TextMatrix(FGrid.Row, Col_DepPer)) > 0 Then
     
     ' FGrid.TextMatrix(FGrid.Row, Col_DepAmt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_ItemVal)) * Val(FGrid.TextMatrix(FGrid.Row, Col_DepitemPer)) / 100) * Val(FGrid.TextMatrix(FGrid.Row, Col_DepPer)) / 100, "0.00")
     
        FGrid.TextMatrix(FGrid.Row, Col_DepAmt) = Format((Val(LblNetValue) * Val(FGrid.TextMatrix(FGrid.Row, Col_DepitemPer)) / 100) * Val(FGrid.TextMatrix(FGrid.Row, Col_DepPer)) / 100, "0.00")
           FGrid.TextMatrix(FGrid.Row, Col_InsuranceAmt) = Format(Val(LblNetValue) - Val(FGrid.TextMatrix(FGrid.Row, Col_DepAmt)), "0.00")
        Else
        FGrid.TextMatrix(FGrid.Row, Col_DepAmt) = ""
        FGrid.TextMatrix(FGrid.Row, Col_InsuranceAmt) = ""
      End If
     
    
    
End Sub

Private Sub CountItem()
Dim I As Integer, TotItems As Integer, TotQty As Double
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            TotQty = TotQty + Val(FGrid.TextMatrix(I, Col_QtyReq))
            TotItems = TotItems + 1
        End If
    Next I
    LblIVal.CAPTION = Format(TotItems, "0")
    LblQty.CAPTION = Format(TotQty, "0.000")
End Sub

Private Sub JobCardDetail(JobCardNo$)
Dim Rst As ADODB.Recordset
    Set Rst = GCn.Execute("Select J.DocId,J.Job_No,J.Job_Date,J.JobCloseDate,J.AtKMsHrs,H.RegNo,H.Chassis,H.Model,H.Engine,H.Name,H.Add1,H.Add2,H.Add3,H.Citycode,H.PhoneOff,H.PhoneResi,H.Mobile,H.VehSerialNo,S.Serv_Type as SrvType,S.Serv_Desc as SrvDesc,S.Serv_Catg,City.CityName,H.DisSprMRP,H.DisSprTB,H.DisSprTP,H.DisOilMRP,H.DisOilTB,H.DisOilTP,Emp_Mast.Emp_Name,h.Delivery_Date as Solddate " _
        & "From ((((Job_card J Left Join Service_type S on J.Serv_type=S.Serv_Type) " _
        & "Left Join HisCard H on J.CardNo=H.CardNo) " _
        & "Left Join City on H.CityCode=City.CityCode) " _
        & "Left Join Emp_Mast on J.RecBy_Mechanic=Emp_Mast.Emp_Code) " _
        & "Where J.DocId='" & JobCardNo & "' " _
        & "Order By J.Job_No")
    If Rst.RecordCount > 0 Then
        mServCatg = Rst!serv_catg
        txt(JobNo).TEXT = IIf(IsNull(Rst!Job_No), "", Rst!Job_No)
        txt(JobNo).Tag = IIf(IsNull(Rst!DocID), "", Rst!DocID)
        txt(JobDt).TEXT = IIf(IsNull(Rst!Job_Date), "", Rst!Job_Date)
        txt(JobClDt).TEXT = IIf(IsNull(Rst!JobCloseDate), "", Rst!JobCloseDate)
        txt(SrvType).TEXT = IIf(IsNull(Rst!SrvDesc), "", Rst!SrvDesc)
        txt(CurrentKMS).TEXT = IIf(IsNull(Rst!AtKMsHrs), "", Rst!AtKMsHrs)
        txt(VehRegNo).TEXT = IIf(IsNull(Rst!RegNo), "", Rst!RegNo)
        txt(Model).TEXT = IIf(IsNull(Rst!Model), "", Rst!Model)
        txt(Chassis).TEXT = IIf(IsNull(Rst!Chassis), "", Rst!Chassis)
        txt(Engine).TEXT = IIf(IsNull(Rst!Engine), "", Rst!Engine)
        txt(VehSrlNo).TEXT = IIf(IsNull(Rst!VehSerialNo), "", Rst!VehSerialNo)
        txt(OwnerName).TEXT = IIf(IsNull(Rst!Name), "", Rst!Name)
        txt(Address1).TEXT = IIf(IsNull(Rst!Add1), "", Rst!Add1)
        txt(Address2).TEXT = IIf(IsNull(Rst!Add2), "", Rst!Add2)
        txt(Address3).TEXT = IIf(IsNull(Rst!Add3), "", Rst!Add3)
        txt(City).TEXT = IIf(IsNull(Rst!CityName), "", Rst!CityName)
        txt(PhoneOff).TEXT = IIf(IsNull(Rst!PhoneOff), "", Rst!PhoneOff)
        txt(PhoneResi).TEXT = IIf(IsNull(Rst!PhoneResi), "", Rst!PhoneResi)
        txt(Mobile).TEXT = IIf(IsNull(Rst!Mobile), "", Rst!Mobile)
        txt(Mechanic).TEXT = IIf(IsNull(Rst!Emp_Name), "", Rst!Emp_Name)
        mDisSprMRP = IIf(IsNull(Rst!DisSprMRP), 0, Rst!DisSprMRP)
        mDisSprTB = IIf(IsNull(Rst!DisSprTB), 0, Rst!DisSprTB)
        mDisSprTP = IIf(IsNull(Rst!DisSprTP), 0, Rst!DisSprTP)
        mDisOilMRP = IIf(IsNull(Rst!DisOilMRP), 0, Rst!DisOilMRP)
        mDisOilTB = IIf(IsNull(Rst!DisOilTB), 0, Rst!DisOilTB)
        mDisOilTP = IIf(IsNull(Rst!DisOilTP), 0, Rst!DisOilTP)
         txt(SoldDate).TEXT = IIf(IsNull(Rst!SoldDate), "", Rst!SoldDate)
    Else
        txt(JobNo).TEXT = "": txt(JobNo).Tag = "": txt(JobDt).TEXT = ""
        txt(SrvType).TEXT = "": txt(CurrentKMS).TEXT = "": txt(VehRegNo).TEXT = ""
        txt(Model).TEXT = "": txt(Chassis).TEXT = "": txt(Engine).TEXT = ""
        txt(VehSrlNo).TEXT = "": txt(OwnerName).TEXT = "": txt(Address1).TEXT = ""
        txt(Address2).TEXT = "": txt(Address3).TEXT = "": txt(City).TEXT = ""
        txt(PhoneOff).TEXT = "": txt(PhoneResi).TEXT = "": txt(Mobile).TEXT = ""
        mDisSprMRP = 0
        mDisSprTB = 0
        mDisSprTP = 0
        mDisOilMRP = 0
        mDisOilTB = 0
        mDisOilTP = 0
    End If
Set Rst = Nothing
End Sub

Private Function GetLubCatOrPurpose(FRow As Integer, Fcol As Integer, FType$) As String
    If FType = "LubCat" Then
        If FGrid.TextMatrix(FRow, Fcol) = "Oil Filter" Then
            GetLubCatOrPurpose = "O"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "Fuel Filter" Then
            GetLubCatOrPurpose = "F"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "Engine Oil" Then
            GetLubCatOrPurpose = "E"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "Gear Oil" Then
            GetLubCatOrPurpose = "G"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "Rear Axle Oil" Then
            GetLubCatOrPurpose = "R"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "Front Axle Oil" Then
            GetLubCatOrPurpose = "A"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "Steering Oil" Then
            GetLubCatOrPurpose = "S"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "N.A." Then
            GetLubCatOrPurpose = "N"
        End If
    ElseIf FType = "Purpose" Then
        If FGrid.TextMatrix(FRow, Fcol) = "PDI" Then
            GetLubCatOrPurpose = "P"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "Free Service" Then
            GetLubCatOrPurpose = "F"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "Charge" Then
            GetLubCatOrPurpose = "C"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "Warranty" Then
            GetLubCatOrPurpose = "W"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "Company Vehicle" Then
            GetLubCatOrPurpose = "O"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "Complementary" Then
            GetLubCatOrPurpose = "L"
        ElseIf FGrid.TextMatrix(FRow, Fcol) = "AMC" Then
            GetLubCatOrPurpose = "A"
        End If
    End If
End Function

Private Sub cmdPost_Click()
Dim I As Double
Dim j As Double
    If PubRequisitionType <> "Store" Then MsgBox "Only From Requisition (Store) Module", vbInformation: Exit Sub
    If Master.RecordCount > 0 Then Master.MoveFirst
    Do Until Master.EOF
        Call MoveRec
        For I = 0 To txt.Count - 1
            txt(I).Refresh
            
        Next
        Call TopCtrl1_eEdit
        For I = 1 To FGrid.Rows - 1
            FGrid.Row = I
            FGrid.Col = Col_DiscPer
            FGrid_DblClick
            Call TxtGrid_Validate(0, False)
            FGrid.Col = Col_TaxPer
            FGrid_DblClick
            Call TxtGrid_Validate(0, False)
            TxtGrid(0).Refresh
'            If Val(FGrid.TextMatrix(i, Col_TaxAmt)) = 0 And Val(FGrid.TextMatrix(i, Col_Amt)) > 0 Then
'                MsgBox ""
'            End If
            If FGrid.TextMatrix(I, Col_PNo) <> "" Then
                GCn.BeginTrans
                GCn.Execute "Update SP_Stock Set " _
                    & "TaxPer=" & Val(FGrid.TextMatrix(I, Col_TaxPer)) & ",TaxAmt=" & Val(FGrid.TextMatrix(I, Col_TaxAmt)) & ",SatPer=" & Val(FGrid.TextMatrix(I, Col_SatPer)) & ",SatAmt=" & Val(FGrid.TextMatrix(I, Col_SatAmt)) & ",Net_Amt=" & Val(FGrid.TextMatrix(I, Col_ItemVal)) & "," _
                    & "Godown='" & FGrid.TextMatrix(I, Col_GodownCode) & "',Remark='" & FGrid.TextMatrix(I, Col_RemWs) + Space(20 - Len(FGrid.TextMatrix(I, Col_RemWs))) + FGrid.TextMatrix(I, Col_RemStores) + Space(10 - Len(FGrid.TextMatrix(I, Col_RemStores))) & "' " _
                    & "Where DocID='" & txt(DocID).TEXT & "' And Srl_No=" & Val(FGrid.TextMatrix(I, Col_SrlNo))
                GCn.CommitTrans
            End If
        Next
MyNextRecord:
        Disp_Text SETS("INI", Me, Master)
        Master.MoveNext
    Loop
End Sub

Private Sub DGJob_Click()
    DGJob.Visible = False
    If RsJob.RecordCount > 0 Then
        JobCardDetail RsJob!Code
    End If
End Sub

Private Sub DGMech_Click()
    DGMech.Visible = False
    If RsMech.RecordCount > 0 Then
        txt(Mechanic).TEXT = RsMech!Name
        txt(Mechanic).Tag = RsMech!Code
    End If
    txt(Mechanic).SetFocus
End Sub

Private Sub DGPart_Click()
On Error GoTo ELoop
    DGPart.Visible = False
    If RsPart.RecordCount > 0 Then
        Select Case FGrid.Col
        Case Col_PNo
            TxtGrid(0).TEXT = RsPart!Code
        Case Col_PName
            TxtGrid(0).TEXT = RsPart!Name
        Case Col_LName
            TxtGrid(0).TEXT = RsPart!LName
        End Select
        TxtGridValid_PNo
    End If
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGGodown_Click()
On Error GoTo ELoop
    DGGodown.Visible = False
    If RsGodown.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsGodown!Name
        FGrid.TextMatrix(FGrid.Row, Col_GodownCode) = RsGodown!Code
        FGrid.TextMatrix(FGrid.Row, Col_Godown) = RsGodown!Name
    End If
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
    If FrmDetail.Visible = True Then FrmDetail.Visible = False
'    If TopCtrl1.TopText2.CAPTION <> "Browse" Then
'        If TxtGrid(0).Visible = False Then SprMRP1
'    End If
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
    CheckError
End Sub
Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
    WinSetting Me: Grid_Ini
    Call Ini_Pub
    TopCtrl1.Tag = PubUParam
    For I = 0 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg
        txt(I).ForeColor = CtrlFColOrg
    Next
    txt(VDate).Tag = PubLoginDate
    
    Me.Show
    DoEvents
    '**
    Set DGPart.DataSource = RsPart

    Set RsJob = New ADODB.Recordset
    RsJob.CursorLocation = adUseClient
        
         Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and  " & cMID("J.DocId", "3", "1") & "='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    RsJob.Open "Select J.DocId AS Code, " & cCStr("J.Job_No") & " As Name,J.Job_No as DispJob_No,J.Job_Date,J.JobCloseDate,H.RegNo,H.Chassis,H.Model,H.Engine,H.Name As OwnerName,H.VehSerialNo,h.Supplier_BillDate as SOldDate " _
        & "From (Job_card J Left Join HisCard H on J.CardNo=H.CardNo) " _
        & "WHERE left(J.docid,1)='" & PubDivCode & "' " & sitecond & " and (JobCloseDate is null or len(JobCloseDate)=0) and (right(j.DocId_InvSpr,8) <> 'Cancelld' Or J.DocId_InvSpr Is Null)" _
        & "Order By J.Job_No", GCn, adOpenDynamic, adLockOptimistic
    Set DGJob.DataSource = RsJob
    
    Set RsMech = New ADODB.Recordset
    RsMech.CursorLocation = adUseClient
    RsMech.Open "Select Emp_Code as code,Emp_Name as name FROM Emp_Mast where Div_Code='" & PubDivCode & "' And Designation  in (" & pubWrkDesigRest & ") Order by Emp_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGMech.DataSource = RsMech

    Set RsGodown = New ADODB.Recordset
    RsGodown.CursorLocation = adUseClient
    RsGodown.Open "Select God_Code as Code,God_Name As Name From Godown Where Appli_For=0 Order by God_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGodown.DataSource = RsGodown


    sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("s.DocId", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    
    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    If PubMoveRecYn Then
        Set Master = GCn.Execute("Select Distinct S.DocID As SearchCode, Sat_Yn " _
            & "From SP_Stock S " _
            & "Where left(s.DocId,1)='" & PubDivCode & "' " & sitecond & " and S.V_Type In ('W_RG','W_RW','W_RGO') " _
            & " Order by S.DocID desc")
    Else
        Set Master = GCn.Execute("Select  Distinct Top 1 S.DocID As SearchCode, Sat_Yn " _
            & "From SP_Stock S " _
            & "Where left(s.DocId,1)='" & PubDivCode & "' and S.V_Type In ('W_RG','W_RW','W_RGO') " _
            & " Order by S.DocID desc")
    End If
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    
    If PubRequisitionType = "Workshop" Or PubRequisitionType = "Store" Then
        Lbl(FinSlipYN).Visible = False
        LblColon(FinSlipYN).Visible = False
        txt(FinSlipYN).Visible = False
    ElseIf PubRequisitionType = "Return" Then
'        If PubRequisitionType = "Store" Then
'            Me.BackColor = &HC6E4E6   '&HCAF1FD
'        Else
            Me.BackColor = &HD7C6C8
'        End If
        FGrid.BackColorBkg = Me.BackColor
        TopCtrl1.tAdd = False
        TopCtrl1.tDel = False
    End If
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        LblFrm(6).Visible = False
        LblFrm(7).Visible = False
        Label3(10).Visible = False
        Label3(14).Visible = False
    End If
    
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsJob = Nothing
    Set RsMech = Nothing
    Set RsGodown = Nothing
    Set Master = Nothing
End Sub

Private Sub FrmDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FrmDetail.MousePointer = 15
End Sub

Private Sub FrmDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
FrmDetail.MousePointer = 0
FrmDetail.Move X, Y
End Sub

Private Sub ListView_Click()
On Error GoTo ELoop
    If TxtGrid(0).Visible = True Then
        TxtGrid(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
        FrmList.Visible = False
        TxtGrid(Val(ListView.Tag)).SetFocus
    Else
        txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
        FrmList.Visible = False
        txt(Val(ListView.Tag)).SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    LockYN = False
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    txt(VDate).TEXT = txt(VDate).Tag
    txt(DocType).TEXT = "General"
    
    mSatYn = IIf(PubSatYn = 1, True, False)
    IniGrid_Vat
    mVType = "W_RG"
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        If MsgBox("Open P.C.D Req.Slip ?", vbYesNo, App.Title) = vbYes Then
            mVType = "W_RG"
        Else
            mVType = "W_RGO"
        End If
    End If

    If PubIPO_Separate = 0 Then         ' Separate IPO is No
        txt(DocType).Enabled = False
        txt(VDate).SetFocus
    Else
        txt(DocType).Enabled = True
        txt(DocType).SetFocus
    End If
    
    txt(FinSlipYN).TEXT = "Yes"
    txt(DocID).TEXT = GetDocID(GCnFaW, mVType, txt(VDate).TEXT, VoucherEditFlag, txt(SerialNo), LblVPrefix)
    txt(DocID).Tag = txt(DocID)
    RsJob.Requery
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset
'Check for existance of transactions
'    If PubRequisitionType = "Workshop" Then
'        If Val(FGrid.TextMatrix(1, Col_QtyIss)) <> 0 Or Val(FGrid.TextMatrix(1, Col_QtyRet)) <> 0 Then
'            MsgBox "Stores Issue Made Against this Requisition, " & vbCrLf & "Can't Edit the Reocord", vbInformation, "Validation"
'            Exit Sub
'        End If
'        Set Rst = New ADODB.Recordset
'        Rst.CursorLocation = adUseClient
'        Rst.Open "Select Qty_Iss,Qty_ret from SP_Stock Where DocId='" & txt(DocId) & "'", GCn, adOpenDynamic, adLockOptimistic
'        If Rst.RecordCount  > 0 Then
'            If Rst!Qty_Iss  > 0 Or Rst!Qty_Ret  > 0 Then
'                MsgBox "Stores Issue Made Against this Requisition, " & vbCrLf & "Can't Edit the Reocord", vbInformation, "Validation"
'                Exit Sub
'            End If
'        End If
'    End If
    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    txt(VDate).Enabled = False
    txt(DocType).Enabled = False
    txt(SerialNo).Enabled = False
    txt(JobNo).Enabled = False
    txt(WarrNo).Enabled = False
    txt(WarrType).Enabled = False
    txt(WarrYear).Enabled = False
    txt(WarrDate).Enabled = False
    If PubRequisitionType = "Workshop" Or PubRequisitionType = "Store" Then
'        FGrid.AddItem FGrid.Rows
'        Txt(Mechanic).SetFocus
'    ElseIf PubRequisitionType = "Store" Then
        txt(WarrNo).Enabled = False
        txt(WarrType).Enabled = False
        txt(WarrYear).Enabled = False
        txt(WarrDate).Enabled = False
        txt(Mechanic).Enabled = False
        LockYN = False
        If CDate(txt(VDate)) < PubLoginDate And RSOJPR = True Then
            MsgBox "Back Date Edit Denied ! You Can Not Edit Qty.Please Make New Requisition.", vbInformation
            LockYN = True
        End If
        FGrid.AddItem FGrid.Rows
        FGrid.Col = Col_PNo
        FGrid.SetFocus
    ElseIf PubRequisitionType = "Return" Then
        txt(VDate).Enabled = False
        txt(WarrNo).Enabled = False
        txt(WarrType).Enabled = False
        txt(WarrYear).Enabled = False
        txt(WarrDate).Enabled = False
        txt(Mechanic).Enabled = False
        FGrid.Col = Col_QtyRet
        FGrid.SetFocus
    End If
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim mTrans As Boolean, vBook As Variant, Rst As ADODB.Recordset
'Check for existance of transactions
    If PubRequisitionType = "Workshop" Then
        If Val(FGrid.TextMatrix(1, Col_QtyIss)) <> 0 Or Val(FGrid.TextMatrix(1, Col_QtyRet)) <> 0 Then
            MsgBox "Stores Issue Made Against this Requisition, " & vbCrLf & "Can't Delete the Reocord", vbInformation, "Validation"
            Exit Sub
        End If
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select Qty_Iss,Qty_ret from SP_Stock Where DocId='" & txt(DocID) & "'", GCn, adOpenDynamic, adLockOptimistic
        If Rst.RecordCount > 0 Then
            If Rst!Qty_Iss > 0 Or Rst!Qty_Ret > 0 Then
                MsgBox "Stores Issue Made Against this Requisition, " & vbCrLf & "Can't Delete the Reocord", vbInformation, "Validation"
                Exit Sub
            End If
        End If
    End If
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            vBook = Master.AbsolutePosition
            GCn.BeginTrans
            mTrans = True
            GCn.Execute ("Delete From SP_Stock Where DocID='" & txt(DocID).TEXT & "'")
            UpdStkTableToTable txt(DocID), "+", "R"
            GCn.CommitTrans
            mTrans = False
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
Set Rst = Nothing
Exit Sub
ELoop:
    If mTrans Then GCn.RollbackTrans
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
    Dim sitecond As String
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("s.DocId", "3", "1") & "='" & PubSiteCode & "'"
    End If


    If PubBackEnd = "A" Then
        GSQL = "Select Distinct S.DocId As SearchCode,Switch(S.V_Type='W_RG','General',S.V_Type='W_RW','Warranty') As DocType, " & cTrim(cMID("S.DocID", "9", "5")) & " As VPrefix, " & cCStr("S.V_No") & " As V_No,HisCard.RegNo, " & cDt("S.V_Date") & " AS VDate, " & cCStr(cTrim("Right(S.Job_DocID, 8)")) & " as Job_No,E.Emp_Name As Mechanic from(((SP_Stock S Left Join Emp_Mast E on S.Mech_Code=E.Emp_Code) Left Join Job_Card on S.Job_DocID = Job_Card.DocId) left join HisCard on Job_Card.CardNo=HisCard.CardNo)Where S.V_Type In ('W_RG','W_RW','W_RGO') " & sitecond & " and left(s.docid,1)='" & PubDivCode & " ' "
    ElseIf PubBackEnd = "S" Then
        GSQL = "Select Distinct S.DocId As SearchCode,Case When S.V_Type='W_RG' Then 'General' When S.V_Type='W_RW' Then 'Warranty' End As DocType, " & _
                "" & cTrim(cMID("S.DocID", "9", "5")) & " As VPrefix, " & cCStr("S.V_No") & " As V_No,HisCard.RegNo, " & cDt("S.V_Date") & " AS VDate, " & _
                "" & cCStr(cTrim("Convert(Numeric,Right(S.Job_DocID, 8))")) & " as Job_No,E.Emp_Name As Mechanic " & _
                "From(((SP_Stock S Left Join Emp_Mast E on S.Mech_Code=E.Emp_Code) " & _
                "Left Join Job_Card on S.Job_DocID = Job_Card.DocId) " & _
                "left join HisCard on Job_Card.CardNo=HisCard.CardNo) " & _
                "Where S.V_Type In ('W_RG','W_RW','W_RGO') " & sitecond & " and left(s.docid,1)='" & PubDivCode & " ' " & _
                "Order by VDate"
    End If
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
       FAFind.Show vbModal
    Else
    ' FAFind.Show vbModal
       FIND2.Show vbModal
    End If
Exit Sub
ELoop:
    CheckError
End Sub
Public Sub SEARCHBACK(ByVal MyValue$)
On Error GoTo ELoop
    If PubMoveRecYn Then
        Master.MoveFirst
    Else
        Set Master = GCn.Execute("Select Distinct S.DocID As SearchCode, Sat_Yn " _
            & "From SP_Stock S " _
            & "Where left(s.DocId,1)='" & PubDivCode & "' and S.V_Type In ('W_RG','W_RW','W_RGO') And S.DocID  = '" & MyValue & "' " _
            & " Order by S.DocID desc")
    End If
    Master.FIND ("SearchCode='" & MyValue & "'")
    MoveRec
    BUTTONS True, Me, Master, 0
Exit Sub
ELoop:
    CheckError
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
On Error GoTo ELoop
    RsJob.Requery
    RsPart.Requery
    RsMech.Requery
    RsGodown.Requery
    'Master.Requery
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean, mGridFilled As Boolean, ReOrderQty As Double
Dim Rst As ADODB.Recordset, DocIdHlp$, LubCat$, Purpose$, TmpStr$
Dim TmpNo As Double
On Error GoTo ELoop
    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    'Checking Job Closed by other user during add/edit of requisition
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select JobCloseDate,ClosedU_Name,ClosedU_EntDt from Job_Card where Job_Card.DocId='" & txt(JobNo).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
    If Not IsNull(Rst!JobCloseDate) Then 'Job Closed
        MsgBox "Job Already Closed by User " & Rst!ClosedU_Name & " Dt." & Rst!ClosedU_EntDt
'        GoTo ELoop
    End If
    Set Rst = Nothing
    'eof of checking
    
    If IsValid(txt(DocType), "Document Type") = False Then Exit Sub
    If IsValid(txt(VDate), "Date") = False Then Exit Sub
    If IsValid(txt(SerialNo), "Serial Number") = False Then Exit Sub
    If IsValid(txt(JobNo), "Job Card No") = False Then Exit Sub
    
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            If PubRequisitionType = "Workshop" Or PubRequisitionType = "Store" Then
                If Val(FGrid.TextMatrix(I, Col_QtyReq)) = 0 Then MsgBox "Please Specify Quantity Requied in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_QtyReq: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
                If FGrid.TextMatrix(I, Col_PartGrade) = PubPartGrade_Lub Then
                    If FGrid.TextMatrix(I, Col_LubCat) = "" Then MsgBox "Please Specify Lubricant Category in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_LubCat: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
                End If
                If FGrid.TextMatrix(I, Col_Purpose) = "" Then MsgBox "Please Specify Part Purpose in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Purpose: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
                If StrCmp(left(PubComp_Name, 4), "Enar") Then
                    'If FGrid.TextMatrix(i, Col_Purpose) = "Warranty" Or FGrid.TextMatrix(i, Col_Purpose) = "Complementary" Or FGrid.TextMatrix(i, Col_Purpose) = "AMC" Then
                    If FGrid.TextMatrix(I, Col_Purpose) = "Warranty" Or FGrid.TextMatrix(I, Col_Purpose) = "Complementary" Then
                        If FGrid.TextMatrix(I, Col_RemWs) = "" Then MsgBox "Please Specify Issue Remark in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_RemWs: FGrid.SetFocus: Exit Sub
                    End If
                Else
                    If FGrid.TextMatrix(I, Col_Purpose) = "Warranty" Or FGrid.TextMatrix(I, Col_Purpose) = "Complementary" Or FGrid.TextMatrix(I, Col_Purpose) = "AMC" Then
                        If FGrid.TextMatrix(I, Col_RemWs) = "" Then MsgBox "Please Specify Issue Remark in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_RemWs: FGrid.SetFocus: Exit Sub
                    End If
                End If
                If FGrid.TextMatrix(I, Col_MRP) = "" Then MsgBox "Please Specify MRP Yes/No in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_MRP: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
                If FGrid.TextMatrix(I, Col_Taxable) = "" Then MsgBox "Please Specify Taxable Yes/No in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Taxable: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
                If Val(FGrid.TextMatrix(I, Col_QtyIss)) > Val(FGrid.TextMatrix(I, Col_QtyReq)) Then
                    MsgBox "Quantity Issued is Greater Than Quantity Reqiured in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_QtyIss: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
                End If
                If Val(FGrid.TextMatrix(I, Col_QtyIss)) > 0 And FGrid.TextMatrix(I, Col_Godown) = "" Then MsgBox "Please Specify Godown in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Godown: FGrid.SetFocus: Exit Sub  ': FGrid.CellBackColor = CellBackColEnter
                If Val(FGrid.TextMatrix(I, Col_Rate)) = 0 Then
'                   If PubULabel <> "Y" Then
                        MsgBox "Please Specify Rate in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Rate: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
'                   End If
                End If
            ElseIf PubRequisitionType = "Return" Then
                If Val(FGrid.TextMatrix(I, Col_QtyRet)) > Val(FGrid.TextMatrix(I, Col_QtyIss)) Then
                    MsgBox "Quantity Returned is Greater Than Quantity Issued in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_QtyRet: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
                End If
            End If
            mGridFilled = True
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Item Detail", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = Col_PNo: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
    If TopCtrl1.TopText2 = "Add" Then
        'lp 12-03-03
        txt(DocID).Tag = txt(DocID)
        If GCn.Execute("Select Count(*) From SP_Stock Where DocID='" & txt(DocID) & "'").Fields(0) > 0 Then
            If VoucherEditFlag Then
                MsgBox "Serial No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                Exit Sub
            Else
                txt(DocID) = GetDocID(GCnFaW, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(txt(DocID).Tag, Document_No)) Then
                    MsgBox "Serial No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    Exit Sub
                End If
            End If
        End If
    End If
    DocIdHlp = UCase(Replace(txt(DocID), " ", ""))
    GCn.BeginTrans
    mTrans = True
    
    
    


    If PubRequisitionType = "Workshop" Or PubRequisitionType = "Store" Then       ' For Workshop
        If TopCtrl1.TopText2 = "Edit" Then
            UpdStkTableToTable txt(DocID), "+", "I"
        End If
        GCn.Execute ("Delete From SP_Stock Where DocID='" & txt(DocID).TEXT & "'")
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" And Val(FGrid.TextMatrix(I, Col_QtyReq)) <> 0 Then
                LubCat = GetLubCatOrPurpose(I, Col_LubCat, "LubCat")
                Purpose = GetLubCatOrPurpose(I, Col_Purpose, "Purpose")
                GCn.Execute "Insert Into SP_Stock(" _
                    & "DocID,Srl_No,V_Type,V_No,V_Date,Site_Code," _
                    & "Job_DocID,Job_DivCode,Claim_No,Claim_Date,Mech_Code," _
                    & "TrnComplete_YN,Part_No,Qty_Doc,Qty_Iss,Qty_Ret," _
                    & "Tax_YN,MRP_YN,Rate,MRP_Rate,Amount," _
                    & "Disc_Per,Disc_Amt,Net_Amt,Godown,Lub_Category," _
                    & "claim_YearPrefix,Claim_Type,Claim_div,Claim_Site," _
                    & "Purpose,Remark, RateType,U_Name,U_EntDt,U_AE,V_Rate,TaxPer,TaxAmt, SatPer, SatAmt, Sat_Yn,PurDocNo,PurDocDate, " & _
                    " Dep_Item , Dep_Code, DepitemPer, DepPer, DepAmt, InsuranceAmt,DiffPeried) " _
                    & "Values(" _
                    & "'" & txt(DocID).TEXT & "'," & I & ",'" & mVType & "'," & txt(SerialNo).TEXT & "," & ConvertDate(txt(VDate).TEXT) & ",'" & PubSiteCode & PubSiteCode & "'," _
                    & "'" & txt(JobNo).Tag & "','" & PubDivCode & "','" & txt(WarrNo).TEXT & "'," & ConvertDate(txt(WarrDate).TEXT) & ",'" & txt(Mechanic).Tag & "'," _
                    & "" & IIf(txt(FinSlipYN).TEXT = "Yes", 1, 0) & ",'" & FGrid.TextMatrix(I, Col_PNo) & "'," & Val(FGrid.TextMatrix(I, Col_QtyReq)) & "," & Val(FGrid.TextMatrix(I, Col_QtyIss)) & "," & Val(FGrid.TextMatrix(I, Col_QtyRet)) & "," _
                    & "" & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & "," & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & "," & Val(FGrid.TextMatrix(I, Col_Rate)) & "," & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," & Val(FGrid.TextMatrix(I, Col_Amt)) & "," _
                    & "" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & "," & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," & Val(FGrid.TextMatrix(I, Col_ItemVal)) & ",'" & FGrid.TextMatrix(I, Col_GodownCode) & "','" & LubCat & "'," _
                    & "'" & txt(WarrYear) & "','" & txt(WarrType) & "','" & IIf(txt(WarrNo) = "", "", PubDivCode) & "','" & IIf(txt(WarrNo) = "", "", PubSiteCode) & "'," _
                    & "'" & Purpose & "','" & FGrid.TextMatrix(I, Col_RemWs) + Space(20 - Len(FGrid.TextMatrix(I, Col_RemWs))) & "', '','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A'," & Val(FGrid.TextMatrix(I, Col_LastRate)) & "," & Val(FGrid.TextMatrix(I, Col_TaxPer)) & "," & Val(FGrid.TextMatrix(I, Col_TaxAmt)) & "," & Val(FGrid.TextMatrix(I, Col_SatPer)) & "," & Val(FGrid.TextMatrix(I, Col_SatAmt)) & ", " & IIf(mSatYn, 1, 0) & ",'" & FGrid.TextMatrix(I, Col_PurDocId) & "'," & ConvertDate(FGrid.TextMatrix(I, Col_PurDate)) & " , " & _
                    " '" & FGrid.TextMatrix(I, Col_DepItem) & "','" & FGrid.TextMatrix(I, Col_DepCode) & "'," & Val(FGrid.TextMatrix(I, Col_DepitemPer)) & ", " & _
                    " " & Val(FGrid.TextMatrix(I, Col_DepPer)) & "," & Val(FGrid.TextMatrix(I, Col_DepAmt)) & "," & Val(FGrid.TextMatrix(I, Col_InsuranceAmt)) & "," & Val(FGrid.TextMatrix(I, Col_DiffPeried)) & " )"
                    
                Call UpdStkGridToTable(FGrid.TextMatrix(I, Col_PNo), "-", FGrid.TextMatrix(I, Col_MRP), FGrid.TextMatrix(I, Col_Taxable), Val(FGrid.TextMatrix(I, Col_QtyIss)))
            End If
        Next
        If TopCtrl1.TopText2 = "Add" Then
            'Voucher Serial No. Updation LPS 21-05-03
            'update Table only when DocSrlNo >Table.SerialNo
            UpdVouSrlNo GCnFaS, txt(DocID), txt(VDate)
        End If
        'If txt(WarrNo) <> "" Then
        '    Call Update_warranty
        'End If
    ElseIf PubRequisitionType = "Return" Then      ' For Return
        'Stock Updation during Return Edit
        UpdStkTableToTable txt(DocID), "+", "I"
        'eof edit stokc upd
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" Then
                GCn.Execute "Update SP_Stock Set " _
                    & "TrnComplete_YN=" & IIf(txt(FinSlipYN).TEXT = "Yes", 1, 0) & ",Qty_Ret=" & Val(FGrid.TextMatrix(I, Col_QtyRet)) & "," _
                    & "Rate=" & Val(FGrid.TextMatrix(I, Col_Rate)) & ",MRP_Rate=" & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," _
                    & "Amount=" & Val(FGrid.TextMatrix(I, Col_Amt)) & ",Disc_Per=" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & "," _
                    & "Disc_Amt=" & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & ",TaxPer=" & Val(FGrid.TextMatrix(I, Col_TaxPer)) & ",TaxAmt=" & Val(FGrid.TextMatrix(I, Col_TaxAmt)) & ",SatPer=" & Val(FGrid.TextMatrix(I, Col_SatPer)) & ",SatAmt=" & Val(FGrid.TextMatrix(I, Col_SatAmt)) & ",Net_Amt=" & Val(FGrid.TextMatrix(I, Col_ItemVal)) & "," _
                    & "Godown='" & FGrid.TextMatrix(I, Col_GodownCode) & "',Remark='" & FGrid.TextMatrix(I, Col_RemWs) + Space(20 - Len(FGrid.TextMatrix(I, Col_RemWs))) + FGrid.TextMatrix(I, Col_RemStores) + Space(10 - Len(FGrid.TextMatrix(I, Col_RemStores))) & "', RateType =  ''," _
                    & "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E',PurDocNo='" & FGrid.TextMatrix(I, Col_PurDocId) & "',PurDocDate=" & ConvertDate(FGrid.TextMatrix(I, Col_PurDate)) & " " _
                    & "Where DocID='" & txt(DocID).TEXT & "' And Srl_No=" & Val(FGrid.TextMatrix(I, Col_SrlNo))
'by lps after QC of Vikash
'                Call UpdStkGridToTable(FGrid.TextMatrix(i, Col_PNo), "-", FGrid.TextMatrix(i, Col_MRP), FGrid.TextMatrix(i, Col_Taxable), Val(FGrid.TextMatrix(i, Col_QtyRet)))
            End If
        Next
        'Stock Updation during Return Edit
        UpdStkTableToTable txt(DocID), "-", "I"
        'eof edit stokc upd
        'If txt(WarrNo) <> "" Then
        '    Call Update_warranty
        'End If
    End If
    GCn.CommitTrans
    'Updating Curr Stock
        For I = 1 To FGrid.Rows - 1
            StkUpd (FGrid.TextMatrix(I, Col_PNo))
        Next
    'End Update
    mTrans = False
    Set Rst = Nothing
    mSearchCode = txt(DocID)
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select Distinct S.DocID As SearchCode, Sat_Yn " _
            & "From SP_Stock S " _
            & "Where left(s.DocId,1)='" & PubDivCode & "' and S.V_Type In ('W_RG','W_RW','W_RGO') And S.DocID  = '" & mSearchCode & "' " _
            & " Order by S.DocID desc")
    End If
    Master.FIND "SearchCode = '" & mSearchCode & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(SerialNo)) > Val(DeCodeDocID(txt(DocID).Tag, Document_No)) Then
            MsgBox "Serial No." & Trim(DeCodeDocID(txt(DocID).Tag, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
        TopCtrl1_ePrn
        Disp_Text SETS("INI", Me, Master)
        MoveRec
    Else
        Disp_Text SETS("INI", Me, Master)
        MoveRec
    End If
    
Exit Sub
ELoop:
'    If mTrans Then GCn.RollbackTrans
'    Set Rst = Nothing
'    CheckError
    MsgBox err.Description & vbCr & "In TopCtrl1_eSave Procedure Of " & Me.Name
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
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
TxtGrid(0).Visible = False
Grid_Hide
Select Case Index
    Case DocType
        ListArray = Array("General", "Warranty")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
        
    Case JobNo
        If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsJob!Name Then
            RsJob.MoveFirst
            RsJob.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
        
    Case Mechanic
        If RsMech.RecordCount = 0 Or (RsMech.EOF = True Or RsMech.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsMech!Name Then
            RsMech.MoveFirst
            RsMech.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
        
        
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
    Case DocType
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case SerialNo
        NumDown txt(Index), KeyCode, 8, 0
    Case JobNo
        DGridTxtKeyDown DGJob, txt, JobNo, RsJob, KeyCode, False, 1
    Case Mechanic
        DGridTxtKeyDown DGMech, txt, Mechanic, RsMech, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
    Case FinSlipYN
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
    End Select
    If FrmList.Visible = False And DGJob.Visible = False And DGMech.Visible = False And DgRateType.Visible = False Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
        If TopCtrl1.TopText2.CAPTION = "Add" Then
            If Index <> DocType And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
            If Index <> VDate And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
Select Case Index
    Case SerialNo
        NumPress txt(Index), KeyAscii, 8, 0
    Case JobNo
        If DGJob.Visible = True Then DGridTxtKeyPress txt, JobNo, RsJob, KeyAscii, "Name"
    Case Mechanic
        If DGMech.Visible = True Then DGridTxtKeyPress txt, Mechanic, RsMech, KeyAscii, "Name"
    Case FinSlipYN
        If UCase(Chr(KeyAscii)) = "Y" Or UCase(Chr(KeyAscii)) = "N" Then
            txt(Index).TEXT = IIf(UCase(Chr(KeyAscii)) = "Y", "Yes", "No")
        End If
        KeyAscii = 0
End Select
Exit Sub

ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
    Case DocType
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, I As Integer
'On Error GoTo ELoop
Select Case Index
    Case DocType
        txt(Index).TEXT = ListView.SelectedItem.TEXT
        If Not Trim(txt(Index).TEXT) <> "General" Or Trim(txt(Index).TEXT) <> "Warranty" Then
            txt(Index).TEXT = "General"
        End If
        If Trim(txt(Index).TEXT) = "General" Then
            mVType = "W_RG"
            txt(WarrNo).Enabled = False
            txt(WarrType).Enabled = False
            txt(WarrYear).Enabled = False
            txt(WarrDate).Enabled = False
            
            For I = 1 To FGrid.Rows - 1
                If FGrid.TextMatrix(I, Col_PNo) <> "" And FGrid.TextMatrix(I, Col_Purpose) = "Warranty" Then
                    FGrid.TextMatrix(I, Col_Purpose) = ""
                End If
            Next
        ElseIf Trim(txt(Index).TEXT) = "Warranty" Then
            mVType = "W_RW"
            txt(WarrNo).Enabled = True
            txt(WarrType).Enabled = True
            txt(WarrYear).Enabled = True
            txt(WarrDate).Enabled = True
            For I = 1 To FGrid.Rows - 1
                If FGrid.TextMatrix(I, Col_PNo) <> "" Then
                    FGrid.TextMatrix(I, Col_Purpose) = "Warranty"
                End If
            Next
        End If
        txt(DocID) = GetDocID(GCnFaW, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
        txt(DocID).Tag = txt(DocID)
    Case VDate
        txt(Index) = RetDate(txt(Index))
        Cancel = Not CheckFinYear(txt(Index))
        If Cancel Then Exit Sub
        If txt(JobDt).TEXT <> "" Then
            If CDate(txt(Index).TEXT) < CDate(txt(JobDt).TEXT) Then
                MsgBox "Requisition Date is Less than Job Card Date", vbInformation, "Validation"
                Cancel = True: Exit Sub
            End If
        End If
        txt(DocID) = GetDocID(GCnFaW, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
        txt(DocID).Tag = txt(DocID)
        RsJob.Filter = ("Job_Date<=" & ConvertDate(Format(txt(VDate), "dd/MMM/yyyy")) & "")
    Case SerialNo
        If VoucherEditFlag = True Then      ' Manual
            txt(DocID) = GetDocID(GCnFaW, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            txt(DocID).Tag = txt(DocID)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select V_No From SP_Stock Where DocID='" & txt(DocID).TEXT & "'", GCn, adOpenStatic, adLockReadOnly
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                txt(SerialNo).SetFocus
                Cancel = True
            End If
        End If
    Case JobNo
'        If RsJob.RecordCount <> 0 And Trim(Txt(Index).Text = "") Then
'            MsgBox "Please Select Job No.", vbInformation, "Information"
'            Txt(Index).SetFocus
'            Cancel = True
'            Exit Sub
'        End If
        If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then
        Else
            JobCardDetail RsJob!Code
        End If
    Case WarrDate
        If txt(WarrDate).TEXT <> "" Then
            txt(Index).TEXT = RetDate(txt(Index))
            If CDate(txt(WarrDate).TEXT) < CDate(txt(JobDt).TEXT) Then
                MsgBox "Warranty Date is Less than Job Card Date", vbInformation, "Validation"
                txt(Index).SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
    Case Mechanic
        If RsMech.RecordCount = 0 Or (RsMech.EOF = True Or RsMech.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsMech!Name
            txt(Index).Tag = RsMech!Code
        End If
        
    Case FinSlipYN
        If Not Trim(txt(Index).TEXT) <> "Yes" Or Trim(txt(Index).TEXT) <> "No" Then
            txt(Index).TEXT = "Yes"
        End If
    Case WarrNo, WarrType, WarrYear
        If txt(WarrYear) = "" And txt(WarrType) = "" And txt(WarrNo) = "" Then Exit Sub
        If Index = WarrYear Then
           If txt(WarrYear) <> "" And (txt(WarrType) = "" Or txt(WarrNo) = "") Then MsgBox "Please Type Proper Claim No.", vbInformation, "Validation": txt(Index).SetFocus: Exit Sub
        End If
        If Index = WarrType Then
            If txt(WarrType) <> "" And (txt(WarrNo) = "" Or txt(WarrYear) = "") Then MsgBox "Please Type Proper Claim No.", vbInformation, "Validation": txt(Index).SetFocus: Exit Sub
        End If
        If Index = WarrNo Then
            'If Txt(WarrNo) <> "" And (Txt(WarrType) = "" Or Txt(WarrYear) = "") Then MsgBox "Please Type Proper Claim No.", vbInformation, "Validation": Txt(Index).SetFocus: Exit Sub
        End If
        
        'Set Rst = GCn.Execute("select job_docid,Claim_Date From job_warr3 where claim_type='" & Txt(WarrType) & "' and Year_Prefix='" & Txt(WarrYear) & "' and claim_no='" & Txt(WarrNo) & "' and div_code='" & PubDivCode & "' and site_code='" & PubSiteCode & "'")
        'If Rst.RecordCount > 0 Then
        '    Txt(WarrDate) = Format(Rst!Claim_Date, "dd/MMM/yyyy")
        '    Txt(JobNo).Tag = Rst!Job_DocID
        '    Txt(JobNo) = DeCodeDocID(Rst!Job_DocID, Document_No)
        '    Txt(WarrDate).Enabled = False
        '    Txt(JobNo).Enabled = False
        '    JobCardDetail Txt(JobNo).Tag
        'Else
        '    Txt(WarrDate).Enabled = True
        '    Txt(JobNo).Enabled = True
        'End If
End Select
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
Dim Rst As ADODB.Recordset
Grid_Hide
If FrmDetail.Visible = False Then FrmDetail.Visible = True
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
Select Case FGrid.Col
    Case Col_PNo
        If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
        RsPart.Sort = "Code"
        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
            RsPart.MoveFirst
            RsPart.FIND "Code='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "'"
            If RsPart.EOF = True Then RsPart.MoveFirst
        End If
    Case Col_LubCat
        ListArray = Array("N.A.", "Oil Filter", "Fuel Filter", "Engine Oil", "Gear Oil", "Rear Axle Oil", "Front Axle Oil", "Steering Oil")
        Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 8)
    Case Col_Purpose
        If PubIPO_Separate = 0 Then         ' Separate IPO=No
           ' If mServCatg = "P" Or mServCatg = "F" Then
                ListArray = Array("PDI", "Free Service", "Charge", "Company Vehicle", "Complementary", "AMC", "Warranty")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 7)
           ' Else
           '     ListArray = Array("Charge", "Company Vehicle", "Complementary", "Warranty")
           '     Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 4)
           ' End If
        Else
'            If mServCatg = "P" Or mServCatg = "F" Then
                ListArray = Array("PDI", "Free Service", "Charge", "Company Vehicle", "Complementary", "AMC")
                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 6)
'            Else
'                ListArray = Array("Charge", "Company Vehicle", "Complementary")
'                Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
'            End If
        End If
    Case Col_Godown
        If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Col_Godown) = "" Then Exit Sub
        If FGrid.TextMatrix(FGrid.Row, Col_Godown) <> RsGodown!Name Then
            RsGodown.MoveFirst
            RsGodown.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, Col_Godown) & "'"
        End If
    Case Col_PName
        If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
        RsPart.Sort = "Name"
        If FGrid.TextMatrix(FGrid.Row, Col_PName) <> "" Then
            RsPart.MoveFirst
            RsPart.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, Col_PName) & "'"
            If RsPart.EOF = True Then RsPart.MoveFirst
        End If
    Case Col_LName
        If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
        RsPart.Sort = "LName"
        If FGrid.TextMatrix(FGrid.Row, Col_LName) <> "" Then
            RsPart.MoveFirst
            RsPart.FIND "LName ='" & FGrid.TextMatrix(FGrid.Row, Col_LName) & "'"
            If RsPart.EOF = True Then RsPart.MoveFirst
        End If
End Select
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then TxtGrid(0) = TxtGrid(0).Tag: Exit Sub
'On Error GoTo ELoop
    If PubRequisitionType = "Workshop" Or PubRequisitionType = "Store" Then
        Select Case FGrid.Col
            Case Col_PNo
                If DGPart.Visible = False Then DGridColSwap DGPart, 0
                DGridTxtKeyDown DGPart, TxtGrid, 0, RsPart, KeyCode, True, 0, frmPartMast, "frmPartMast"
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName, 1 ' , Col_QtyReq
                    End If
                End If
            Case Col_PName
                If DGPart.Visible = False Then DGridColSwap DGPart, 1
                DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 1, frmPartMast, "frmPartMast"
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName
                    End If
                End If
            Case Col_LName
                If DGPart.Visible = False Then DGridColSwap DGPart, 6
                DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 2, frmPartMast, "frmPartMast"
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName ', 1
                    End If
                End If
'            Case Col_QtyReq
'                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
'                    If TxtGridLeave = True Then
'                        If FGrid.TextMatrix(FGrid.Row, Col_PartGrade) = PubPartGrade_Lub Then
'                            GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName, , Col_LubCat
'                        ElseIf FGrid.TextMatrix(FGrid.Row, Col_PartGrade) = PubPartGrade_Lub And txt(DocType).Text = "General" Then
'                            GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName, , Col_Purpose
'                        ElseIf FGrid.TextMatrix(FGrid.Row, Col_PartGrade) <> PubPartGrade_Lub And txt(DocType).Text = "General" Then
'                            GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName, , Col_Purpose
'                        ElseIf FGrid.TextMatrix(FGrid.Row, Col_PartGrade) <> PubPartGrade_Lub And txt(DocType).Text = "Warranty" Then
'                            GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName, , Col_RemWs
'                        End If
'                    End If
'                End If
            Case Col_LubCat
                ListView_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, 60, 435, 2500, 2400
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, IIf(left(FGrid.TextMatrix(FGrid.Row, Col_Purpose), 1) = "W", Col_LName, Col_PNo)
                    End If
                End If
            Case Col_Purpose
                ListView_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, 3115, 1865, 1800, 1620
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName
                    End If
                End If
            Case Col_RemWs  '', Col_PurDocId, Col_PurDate
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PNo
                    End If
                End If
            Case Col_PurDocId, Col_PurDate
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PurDate
                    End If
                End If
            Case Col_MRP
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_RemStores, , Col_Taxable
                    End If
                End If
            Case Col_Taxable
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_RemStores, , Col_QtyIss
                    End If
                End If
            Case Col_QtyIss
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_RemStores, , Col_Rate
                    End If
                End If
            Case Col_Rate
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_RemStores, , Col_DiscPer
                    End If
                End If
            Case Col_DiscPer, Col_TaxPer, Col_SatPer
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_RemStores, , Col_DiscAmt
                    End If
                End If
                Amt_Cal
            Case Col_DiscAmt
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                        If PubRestrict_Godown = 1 Then      ' Restrict Godown is "YES"
                            'Purpose not Clear, Redesign
'                            FGrid_LeaveCell
                            GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Col_RemStores, , Col_Godown
                        Else
                            GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_RemStores, , Col_Godown
                        End If
                    End If
                End If
            Case Col_Godown
                DGridTxtKeyDown DGGodown, TxtGrid, Index, RsGodown, KeyCode, True, 1, frmGodown, "frmGodown"
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, IIf(FGrid.TextMatrix(FGrid.Row, Col_PartGrade) = PubPartGrade_Lub, Col_LubCat, IIf(left(FGrid.TextMatrix(FGrid.Row, Col_Purpose), 1) = "W", Col_RemStores, Col_PNo)), 1, IIf(FGrid.TextMatrix(FGrid.Row, Col_PartGrade) = PubPartGrade_Lub, Col_LubCat, IIf(left(FGrid.TextMatrix(FGrid.Row, Col_Purpose), 1) = "W", Col_RemWs, 0))
                    End If
                End If
            Case Col_RemStores
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_RemStores
'                        If FGrid.Row <> FGrid.Rows - 1 Then
'                            FGrid.Row = FGrid.Row + 1
'                            FGrid.Col = Col_MRP
'                            FGrid.SetFocus
'                        Else
'                            SendKeysA vbKeyTab, True
'                        End If
    '                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName
                    End If
                End If
        End Select
    ElseIf PubRequisitionType = "Return" Then
        Select Case FGrid.Col
            Case Col_QtyRet
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName, , Col_RemStores
                    End If
                End If
            Case Col_RemStores
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave = True Then
'                        If FGrid.Row <> FGrid.Rows - 1 Then
'                            FGrid.Row = FGrid.Row + 1
'                            FGrid.Col = Col_QtyRet
'                            FGrid.SetFocus
'                        Else
                            SendKeysA vbKeyTab, True
'                        End If
    '                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName
                    End If
                End If
        End Select
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Exit Sub
On Error GoTo ELoop
CheckQuote KeyAscii
Select Case FGrid.Col
    Case Col_PNo
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "Code"
    Case Col_PName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "Name"
    Case Col_LName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "LName"
    Case Col_Godown
        If DGGodown.Visible = True Then
            If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                DGridTxtKeyPress TxtGrid, Index, RsGodown, KeyAscii, "Name"
            End If
        End If
    Case Col_QtyReq, Col_QtyIss, Col_QtyRet
        NumPress TxtGrid(Index), KeyAscii, 8, 3
    Case Col_DiscPer, Col_TaxPer, Col_SatPer
        NumPress TxtGrid(Index), KeyAscii, 3, 2
    Case Col_Rate, Col_DiscAmt
        NumPress TxtGrid(Index), KeyAscii, 8, 3
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case FGrid.Col
    Case Col_PNo
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Code", True
    Case Col_PName
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Name", True
    Case Col_LName
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "LName", True
    Case Col_Godown
        If KeyCode <> 13 And DGGodown.Visible = False Then
            TxtGrid_KeyDown Index, GridKey, 0
            If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                DGridTxtKeyPress TxtGrid, Index, RsGodown, KeyCode, "Name", True
            End If
        End If
    Case Col_LubCat
        If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
        ListView_KeyUp ListView, TxtGrid, Index, KeyCode, mListItem
    Case Col_QtyReq
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.000")
        CountItem
    Case Col_Purpose
        If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
        If FrmList.Visible = True Then ListView_KeyUp ListView, TxtGrid, Index, KeyCode, mListItem
        Amt_Cal
    Case Col_QtyIss, Col_QtyRet
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.000")
        CountItem
        Amt_Cal
    Case Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
    Case Col_Rate
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        FGrid.TextMatrix(FGrid.Row, Col_MRPRate) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        
        Amt_Cal
    Case Col_RemWs, Col_RemStores 'Col_LubCat, Col_Purpose, Col_RemWs, Col_RemStores
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(Index).TEXT
End Select
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)

Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
    FGrid.Tag = FGrid.Row
End Sub

Private Sub FGrid_DblClick()
On Error GoTo ELoop
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
FGrid_KeyPress (vbKeyReturn)
TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_GotFocus()
    If TopCtrl1.TopText2 = "Add" Then
        If Trim(txt(JobNo).TEXT) = "" Then
            MsgBox "Please Select Job No.", vbInformation, "Information"
            txt(JobNo).SetFocus
            Exit Sub
        End If
    End If
    
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    TxtGrid(0).Visible = False
    If TopCtrl1.TopText2 <> "Browse" Then
'        MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, Col_PNo), _
            FGrid.TextMatrix(FGrid.Row, Col_PName), FGrid.TextMatrix(FGrid.Row, Col_LName), _
            Col_MRPStkTB, Col_MRPStkTP, _
            Col_TBStk, Col_TPStk, _
            Col_MRPRate, Col_TBRate, _
            Col_TPRate, Col_Bin, _
            Col_LastRate , Col_HPRate, Col_LPRate
        FrmDetail.Visible = True
    End If
    Grid_Hide
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
        If PubRequisitionType = "Workshop" Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    'ElseIf KeyCode = vbKeyDown And FGrid.Row = FGrid.Rows - 1 Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
            TopCtrl1_eSave
        Else
            FGrid.Tag = FGrid.Row
        End If
        Exit Sub
'        SendKeysA vbKeyTab, True
'        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        If PubRequisitionType = "Workshop" Or PubRequisitionType = "Store" Then
            If FGrid.TextMatrix(FGrid.Row, Col_QtyRet) <> "" Then MsgBox "Return Qty Exists !", vbOKOnly, "Qty Checking": FGrid.SetFocus: Exit Sub
            Select Case FGrid.Col
                Case Col_QtyReq, Col_RemWs, Col_PurDocId, Col_PurDate
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    Amt_Cal
                Case Col_LubCat
                    If FGrid.TextMatrix(FGrid.Row, Col_PartGrade) = PubPartGrade_Lub Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    End If
                Case Col_Purpose
                    If txt(DocType).TEXT = "General" Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    End If
                Case Col_MRP, Col_Taxable
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                Case Col_QtyIss, Col_Rate, Col_DiscPer, Col_DiscAmt, Col_RemStores, Col_TaxPer, Col_SatPer
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    Amt_Cal
                Case Col_Godown
                    If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    End If
            End Select
        ElseIf PubRequisitionType = "Return" Then
            Select Case FGrid.Col
                Case Col_QtyRet, Col_RemStores
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    Amt_Cal
            End Select
        End If
    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
On Error GoTo ELoop
Dim mRate As Double
Dim xColVal$
SetMaxLength
    If PubRequisitionType = "Workshop" Or PubRequisitionType = "Store" Then
        Select Case FGrid.Col
            Case Col_PNo, Col_RemWs
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            Case Col_PName, Col_LName
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) = "" Then
                    Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
                End If
'            Case Col_QtyReq
'                Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
            Case Col_LubCat
                If FGrid.TextMatrix(FGrid.Row, Col_PartGrade) = PubPartGrade_Lub Then
                    Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
                End If
            Case Col_Purpose
                If txt(DocType).TEXT = "General" Then
                    Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
                End If
            Case Col_MRP, Col_Taxable
                xColVal = left(FGrid, 1)
                'If FGrid.TextMatrix(FGrid.Row, FGrid.Col) <> "" Then
                    If UCase(Chr(KeyAscii)) = "Y" Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Yes"
                    Else
                        If PubVATYN = 1 And RSOJPR = False And KeyAscii = 13 Then
                            If FGrid.Col = Col_MRP Then
                                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "No"
                            ElseIf FGrid.Col = Col_Taxable Then
                                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Yes"
                            End If
                        Else
                            If FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "" Then
                                If KeyAscii = Asc("Y") Or KeyAscii = Asc("y") Then
                                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Yes"
                                Else
                                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "No"
                                End If
                            Else
                                If KeyAscii <> 13 And KeyAscii <> 10 Then
                                    If KeyAscii = Asc("Y") Or KeyAscii = Asc("y") Then
                                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Yes"
                                    Else
                                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "No"
                                    End If
                                End If
                            End If
                        End If
                    End If
                If PubSprIssOnNegStk = 0 Then
                    FGrid.TextMatrix(FGrid.Row, Col_QtyIss) = ""
                    FGrid.TextMatrix(FGrid.Row, Col_QtyReq) = ""
                End If
                If xColVal <> left(FGrid, 1) Then
                    FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(0, FGrid, CDate(txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate, FGrid.TextMatrix(FGrid.Row, Col_Purpose), Val(FGrid.TextMatrix(FGrid.Row, Col_LastRate)))
                End If
                KeyAscii = 0
                If ChkDuplicate = False Then Exit Sub
                If FGrid.Col = Col_MRP Then
                    FGrid.Col = FGrid.Col + 1
                Else
                    FGrid.Col = Col_QtyIss
                End If
            Case Col_RemStores
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            Case Col_QtyIss
                If LockYN = False Then
                    Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
                End If
            Case Col_QtyIss, Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer
                Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
            Case Col_Rate
                If UCase(left(PubComp_Name, 6)) <> "RASHMI" Then
                    Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
                End If
            Case Col_Godown
                If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                    Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
                End If
                'If UCase(left(PubComp_Name, 3)) <> "LMP" Then
'                    mRate = GetRate(0, FGrid, CDate(Txt(Vdate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
'                    FGrid.TextMatrix(FGrid.Row, Col_Rate) = IIf(mRate <= 0, "", Format(mRate, "0.000"))
'                    Amt_Cal
                'End If
            Case Col_PurDocId, Col_PurDate
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        End Select
    ElseIf PubRequisitionType = "Return" Then
        Select Case FGrid.Col
            Case Col_QtyRet
                Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
            Case Col_RemStores
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        End Select
    End If
    If KeyAscii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If PubRequisitionType = "Return" Then Exit Sub
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
           FGrid.TextMatrix(I, Col_SrNo) = I
        Next
        CountItem
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_RowColChange()
    If TopCtrl1.TopText2.CAPTION <> "Browse" Then
        MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, Col_PNo), _
            FGrid.TextMatrix(FGrid.Row, Col_PName), FGrid.TextMatrix(FGrid.Row, Col_LName), _
            Col_MRPStkTB, Col_MRPStkTP, _
            Col_TBStk, Col_TPStk, _
            Col_MRPRate, Col_TBRate, _
            Col_TPRate, Col_Bin, _
            Col_LastRate, Col_HPRate, Col_LPRate, mCheckNegetiveStockSiteWise
    End If
    
    LblNetValue = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_ItemVal)) + Val(FGrid.TextMatrix(FGrid.Row, Col_TaxAmt)) + Val(FGrid.TextMatrix(FGrid.Row, Col_SatAmt)), "0.00")
End Sub

Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub Update_warranty()
Dim I As Integer
Dim MySrlNo As Integer
Dim Rst As ADODB.Recordset
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_Purpose) = "Warranty" Then
            Set Rst = GCn.Execute("select ipo_Docid From job_warr3 where ipo_Docid='" & txt(DocID).TEXT & "' and ipo_srl=" & I)
            If Rst.RecordCount = 0 Then
                Set Rst = GCn.Execute("select Job_DocId From job_warr1 " & _
                    " where Div_Code&Site_Code&Year_Prefix&Claim_Type&Claim_No='" & PubDivCode & PubSiteCode & txt(WarrYear) & txt(WarrType) & txt(WarrNo) & "'")
                If Rst.RecordCount <= 0 Then
                    GSQL = ("Insert into Job_Warr1(" & _
                            "Div_code,site_code,year_prefix,claim_type,Claim_no," & _
                            "Claim_date,Job_DocId," & _
                            "U_Name,U_EntDt,U_AE" & _
                            ") values(" & _
                            " '" & PubDivCode & "','" & PubSiteCode & "','" & txt(WarrYear) & "','" & txt(WarrType) & "','" & txt(WarrNo) & "'," & _
                            " " & ConvertDate(txt(WarrDate).TEXT) & ",'" & txt(JobNo).Tag & "'," & _
                            "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
                    GCn.Execute GSQL
                End If
                MySrlNo = GCn.Execute("select " & vIsNull("max(srl_no)", "0") & "+1 from job_warr3 where claim_type='" & txt(WarrType) & "' and Year_Prefix='" & txt(WarrYear) & "' and claim_no='" & txt(WarrNo) & "' and div_code='" & PubDivCode & "' and site_code='" & PubSiteCode & "'").Fields(0).Value
                GSQL = "Insert into Job_Warr3(" & _
                        "Div_code,site_code,year_prefix,claim_type,Claim_no," & _
                        "Claim_date,Srl_no,IPO_Docid,IPO_No,IPO_Srl," & _
                        "IPO_Date,Job_DocId,Part_No,Iss_qty,MRP_YN," & _
                        "Tax_YN,Rate,Mrp_Rate,Disc_Per,NDP," & _
                        "U_Name,U_EntDt,U_AE" & _
                        ") values(" & _
                        " '" & PubDivCode & "','" & PubSiteCode & "','" & txt(WarrYear) & "','" & txt(WarrType) & "','" & txt(WarrNo) & "'," & _
                        " " & ConvertDate(txt(WarrDate).TEXT) & "," & MySrlNo & ",'" & txt(DocID).TEXT & "','" & txt(SerialNo) & "'," & I & "," & _
                        " " & ConvertDate(txt(VDate).TEXT) & ",'" & txt(JobNo).Tag & "','" & FGrid.TextMatrix(I, Col_PNo) & "'," & Val(FGrid.TextMatrix(I, Col_QtyIss)) - Val(FGrid.TextMatrix(I, Col_QtyRet)) & "," & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & "," & _
                        " " & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & "," & Val(FGrid.TextMatrix(I, Col_Rate)) & "," & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," & Val(FGrid.TextMatrix(I, Col_DiscPer)) & "," & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," & _
                        " '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A'" & _
                        ")"
                GCn.Execute GSQL
            Else
                GSQL = "Update Job_Warr3 set IPO_Date=" & ConvertDate(txt(VDate).TEXT) & ",Job_DocId= '" & txt(JobNo).Tag & "',Part_No='" & FGrid.TextMatrix(I, Col_PNo) & "'," & _
                        "Iss_qty=" & Val(FGrid.TextMatrix(I, Col_QtyIss)) - Val(FGrid.TextMatrix(I, Col_QtyRet)) & ",MRP_YN=" & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & "," & _
                        "Tax_YN=" & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & ",Rate=" & Val(FGrid.TextMatrix(I, Col_Rate)) & ",Mrp_Rate=" & Val(FGrid.TextMatrix(I, Col_MRPRate)) & ",Disc_Per=" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & ",NDP=" & Val(FGrid.TextMatrix(I, Col_MRPRate)) & _
                        " where ipo_Docid='" & txt(DocID).TEXT & "' and ipo_srl=" & I
                GCn.Execute GSQL
            End If
        End If
    Next
End Sub

Private Sub TxtGridValid_PNo()
'Called from TxtGrid_Validate & TxtGridLeave procedures
Dim OldPNo$, LstPur As ADODB.Recordset, RetVal As String
Dim mDatediff As Integer
Dim MrstDEp As ADODB.Recordset

If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Or TxtGrid(0).TEXT = "" Then
    FGrid.TextMatrix(FGrid.Row, Col_PNo) = ""
    FGrid.TextMatrix(FGrid.Row, Col_PName) = ""
    FGrid.TextMatrix(FGrid.Row, Col_LName) = ""
    FGrid.TextMatrix(FGrid.Row, Col_Purpose) = ""   'Warranty
    
    MainLib.Fill_Data 0, LblFrm, FGrid, _
        "", "", "", _
        Col_Unit, Col_MRP, Col_Taxable, Col_MRPStkTB, Col_MRPStkTP, _
        Col_TBStk, Col_TPStk, _
        Col_MRPRate, Col_TBRate, _
        Col_TPRate, Col_Bin, _
        Col_HPRate, Col_LPRate, _
        Col_LastRate, Col_PartGrade, _
        Col_EffectDate, Col_DiscPer, mCheckNegetiveStockSiteWise
Else
    OldPNo = FGrid.TextMatrix(FGrid.Row, Col_PNo)
    FGrid.TextMatrix(FGrid.Row, Col_PNo) = RsPart!Code
    FGrid.TextMatrix(FGrid.Row, Col_PName) = RsPart!Name
    FGrid.TextMatrix(FGrid.Row, Col_LName) = RsPart!LName
    
   'Start****************************Nikhil
    'Deptcode
    FGrid.TextMatrix(FGrid.Row, Col_DepItem) = XNull(RsPart!Deptcode)
    FGrid.TextMatrix(FGrid.Row, Col_DepitemPer) = XNull(RsPart!Dep_per)
    If txt(SoldDate) <> "" Then
        mDatediff = DateDiff("M", CDate(txt(SoldDate)), CDate(txt(JobDt)))
        FGrid.TextMatrix(FGrid.Row, Col_DiffPeried) = VNull(mDatediff)
    End If
    Set MrstDEp = GCn.Execute("SELECT  TOP 1 * FROM Deprecation_Master WHERE Dep_Month>" & mDatediff & " ORDER BY Dep_Month ")
    If MrstDEp.RecordCount > 0 Then
    
    FGrid.TextMatrix(FGrid.Row, Col_DepCode) = XNull(MrstDEp!Code)
    FGrid.TextMatrix(FGrid.Row, Col_DepPer) = XNull(MrstDEp!Dep_per)
    Else
    FGrid.TextMatrix(FGrid.Row, Col_DepCode) = ""
    FGrid.TextMatrix(FGrid.Row, Col_DepPer) = ""
    
    End If
'****************End

    If PubIPO_Separate <> 0 And txt(DocType) = "Warranty" Then          ' Separate IPO='Yes'
        FGrid.TextMatrix(FGrid.Row, Col_Purpose) = "Warranty"
    End If
    MainLib.Fill_Data 0, LblFrm, FGrid, _
        RsPart!Code, RsPart!Name, RsPart!LName, _
        Col_Unit, Col_MRP, Col_Taxable, Col_MRPStkTB, Col_MRPStkTP, _
        Col_TBStk, Col_TPStk, _
        Col_MRPRate, Col_TBRate, _
        Col_TPRate, Col_Bin, _
        Col_HPRate, Col_LPRate, _
        Col_LastRate, Col_PartGrade, _
        Col_EffectDate, Col_DiscPer, mCheckNegetiveStockSiteWise
    If RSOJPR = True Then
        FIFOStkIss (RsPart!Code)
    End If
    If PubVATYN = 1 Then
        GSQL = "Select TAX_Per, AddTaxPer from TaxForms where Form_Code=(Select LocalTaxFormSpr from syctrl)"
        Set rsTaxPer = GCn.Execute(GSQL)
        If rsTaxPer.RecordCount > 0 Then
            FGrid.TextMatrix(FGrid.Row, Col_TaxPer) = rsTaxPer!Tax_Per
            FGrid.TextMatrix(FGrid.Row, Col_SatPer) = VNull(rsTaxPer!AddTaxPer)
              
            Set rsTaxPer = GCn.Execute("Select VatPer, AddTaxPer From Part_Grade Where PartGrade_Code='" & FGrid.TextMatrix(FGrid.Row, Col_PartGrade) & "'")
            If rsTaxPer.RecordCount > 0 Then
                If VNull(rsTaxPer!VatPer) > 0 Then FGrid.TextMatrix(FGrid.Row, Col_TaxPer) = Format(rsTaxPer!VatPer, "0.00")
                If VNull(rsTaxPer!AddTaxPer) > 0 Then FGrid.TextMatrix(FGrid.Row, Col_SatPer) = Format(rsTaxPer!AddTaxPer, "0.00")
            End If
                
         End If
    End If
        
    
    FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(0, FGrid, CDate(txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate, FGrid.TextMatrix(FGrid.Row, Col_Purpose), Val(FGrid.TextMatrix(FGrid.Row, Col_LastRate)))
        
    If RSOJPR = False Then
        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> OldPNo Then
                FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(0, FGrid, CDate(txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate, FGrid.TextMatrix(FGrid.Row, Col_Purpose), Val(FGrid.TextMatrix(FGrid.Row, Col_LastRate)))
                
                'Display Last Purchase
                Set LstPur = GCn.Execute("Select PurDocId,PurDate from Part where Part_No='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "' and Div_Code='" & PubDivCode & "'")
                If LstPur.RecordCount > 0 Then
                    FGrid.TextMatrix(FGrid.Row, Col_PurDocId) = XNull(LstPur!PurDocId)
                    FGrid.TextMatrix(FGrid.Row, Col_PurDate) = XNull(LstPur!PurDate)
                End If
    '           FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsPart!SalDisc_Per, "0.00")
            End If
        End If
    End If
    

    
    
End If
Amt_Cal
If FGrid.TextMatrix(FGrid.Rows - 1, Col_PNo) <> "" Then FGrid.AddItem FGrid.Rows
End Sub

Private Sub TxtGridValid_TaxMRP()
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
'        If TopCtrl1.TopText2 = "Add" Or _
            TopCtrl1.TopText2 = "Edit" And Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) = 0 Then
            FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(0, FGrid, CDate(txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate, FGrid.TextMatrix(FGrid.Row, Col_Purpose), Val(FGrid.TextMatrix(FGrid.Row, Col_LastRate)))
            
'        End If
    End If
    If FGrid.TextMatrix(FGrid.Row, Col_PartGrade) = PubPartGrade_Lub Then
        If (FGrid.TextMatrix(FGrid.Row, Col_MRP) = "Yes" And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes") Then
            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = mDisOilMRP
        ElseIf (FGrid.TextMatrix(FGrid.Row, Col_MRP) <> "Yes" And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes") Then
            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = mDisOilTB
        Else
            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = mDisOilTP
        End If
    Else    'If FGrid.TextMatrix(FGrid.Row, Col_PartGrade) <> PubPartGrade_Lub Then
        If (FGrid.TextMatrix(FGrid.Row, Col_MRP) = "Yes" And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes") Then
            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = mDisSprMRP
        ElseIf (FGrid.TextMatrix(FGrid.Row, Col_MRP) <> "Yes" And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes") Then
            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = mDisSprTB
        Else
            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = mDisSprTP
        End If
    End If
    
    Amt_Cal
End Sub

Private Sub TxtGridValid_Godown()

    If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or TxtGrid(0).TEXT = "" Then
        FGrid.TextMatrix(FGrid.Row, Col_GodownCode) = ""
        FGrid.TextMatrix(FGrid.Row, Col_Godown) = ""
    Else
        FGrid.TextMatrix(FGrid.Row, Col_GodownCode) = RsGodown!Code
        FGrid.TextMatrix(FGrid.Row, Col_Godown) = RsGodown!Name
    End If

End Sub

Private Sub SetMaxLength()
    Select Case FGrid.Col
        Case Col_PNo
            TxtGrid(0).MaxLength = 22
            TxtGrid(0).Alignment = 0   '0-Left Align
        Case Col_LubCat, Col_Purpose, Col_RemWs, Col_Godown
            TxtGrid(0).MaxLength = 20
            TxtGrid(0).Alignment = 0   '0-Left Align
        Case Col_RemStores
            TxtGrid(0).MaxLength = 10
            TxtGrid(0).Alignment = 0   '0-Left Align
        Case Col_PName, Col_LName
            TxtGrid(0).MaxLength = 40
            TxtGrid(0).Alignment = 0   '0-Left Align
        Case Else
            TxtGrid(0).MaxLength = 0
    End Select
End Sub



Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        FrmPrn.Visible = False
        If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
            If TopCtrl1.TopText2.CAPTION = "Add" Then
                txt(VDate).Tag = txt(VDate).TEXT
                TopCtrl1_eAdd
                Exit Sub
            End If
            Disp_Text SETS("INI", Me, Master)
            If PubRequisitionType = "Store" Or PubRequisitionType = "Return" Then
                TopCtrl1.tAdd = False
                TopCtrl1.tDel = False
            End If
            MoveRec
        End If
    End If
End Sub

Private Sub CmdPrint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
Dim sQryPurBillNo$, sQryPurBillDt$, SQryPurRate$

    sQryPurBillNo = "SELECT TOP 1 Pur.Party_Doc_No " & _
                    "FROM SP_Purch Pur " & _
                    "LEFT JOIN SP_Stock S ON Pur.DocID =S.Invoice_DocId " & _
                    "WHERE S.Part_No = Sp_Stock.Part_No ORDER BY Pur.V_Date Desc "
    sQryPurBillDt = "SELECT TOP 1 Pur.Party_Doc_Date " & _
                    "FROM SP_Purch Pur " & _
                    "LEFT JOIN SP_Stock S ON Pur.DocID =S.Invoice_DocId " & _
                    "WHERE S.Part_No = Sp_Stock.Part_No ORDER BY Pur.V_Date Desc "
    SQryPurRate = "SELECT TOP 1 S.Rate " & _
                    "FROM SP_Purch Pur " & _
                    "LEFT JOIN SP_Stock S ON Pur.DocID =S.Invoice_DocId " & _
                    "WHERE S.Part_No = Sp_Stock.Part_No ORDER BY Pur.V_Date Desc "


    GSQL = "Select Sp_Stock.Srl_No,p.Part_No,P.Part_Name,SP_Stock.Rate,P.Bin_Loca,SP_Stock.Purpose,SP_Stock.Qty_Doc AS QReq,SP_Stock.Qty_Iss AS QIss,SP_Stock.Qty_Ret AS QRec,SP_Stock.TAX_YN,Sp_Stock.Printed,SP_Stock.Amount,'" & txt(SrvType) & "' as SrvType ,SP_Stock.Purpose,(" & sQryPurBillNo & ") as PurDocId, (" & sQryPurBillDt & ") as PurDate, (" & SQryPurRate & ") As PurRate " & _
        " From ((SP_Stock Left Join Part P On SP_Stock.Part_No=P.Part_No and P.Div_Code = left(SP_Stock.DocID,1)) " & _
        " Left Join Emp_Mast on SP_Stock.Mech_Code=Emp_Mast.Emp_Code) " & _
        " Where SP_Stock.DocID='" & Master!SearchCode & "' Order By SP_Stock.Srl_No"

Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "WorkRequi", "WorkRequi")
        Call WindowsPrint(Index, GSQL)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint(GSQL)
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "WorkRequi", "WorkRequi")
        Call PrinerSetUp
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        txt(VDate).Tag = txt(VDate).TEXT
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    If PubRequisitionType = "Store" Or PubRequisitionType = "Return" Then
        TopCtrl1.tAdd = False
        TopCtrl1.tDel = False
    End If
    MoveRec
End If
'TopCtrl1_eAdd
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub WindowsPrint(Index As Integer, mQry As String)
Dim I As Integer, mDocStr$, mQuanIRR$, mComm_Mes$, mFootMes$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER
'mqry = "Select Sp_Stock.Srl_No,p.Part_No,P.Part_Name,SP_Stock.Rate,P.Bin_Loca,SP_Stock.Purpose,SP_Stock.Qty_Doc AS QReq,SP_Stock.Qty_Iss AS QIss,SP_Stock.Qty_Rec AS QRec,SP_Stock.TAX_YN,Sp_Stock.Printed " & _
'    " From ((SP_Stock Left Join Part P On SP_Stock.Part_No=P.Part_No and P.Div_Code = left(SP_Stock.DocID,1)) " & _
'    " Left Join Emp_Mast on SP_Stock.Mech_Code=Emp_Mast.Emp_Code) " & _
'    " Where SP_Stock.DocID='" & Master!searchcode & "' Order By SP_Stock.Srl_No"
Set Rst = GCn.Execute(mQry)

If Me.CAPTION = "Requisition &Returns" Then
    mDocStr = "GOODS RETURN SLIP (WORKS TO STORES)"
    mQuanIRR = "   REQU.   ISSUED   RETURNED"
    mComm_Mes = "Please acknowledge the following Goods as Return against Issue to Workshop"
    mFootMes = "Goods Received By:" + pubUName + "  Date:" + Format(date, "dd/mm/yy") + "  Time:" + Format(time, "hh:mm") + "     Mechanic:" + txt(Mechanic)
Else
    mDocStr = "REQUISITION SLIP"
    mQuanIRR = "   REQU.   ISSUED"
    mComm_Mes = "Please issue the following goods against the original copy of this Requisition Slip"
    mFootMes = "Mechanic : " + txt(Mechanic) + Space(10) + "Supervisor/Manager" + Space(10) + "Stores"
End If
mDocStr = "** " + mDocStr + IIf(Rst!Printed = 1, " (DUPLICATE)", "") + " **"

CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("JOB_NO")
            rpt.FormulaFields(I).TEXT = "'" & txt(JobNo) & "'"
        Case UCase("JOB_DT")
            rpt.FormulaFields(I).TEXT = "'" & Format(txt(JobDt), "DD/MM/YY") & "'"
        Case UCase("REQ_NO")
            rpt.FormulaFields(I).TEXT = "'" & Replace(mID(txt(DocID), 9, 15), " ", "") & "'"
        Case UCase("REQ_DT")
            rpt.FormulaFields(I).TEXT = "'" & Format(txt(VDate), "DD/MM/YY") & "'"
        Case UCase("MODEL")
            rpt.FormulaFields(I).TEXT = "'" & txt(Model) & "'"
        Case UCase("VEH_NO")
            rpt.FormulaFields(I).TEXT = "'" & txt(VehRegNo) & "'"
        Case UCase("CHA_NO")
            rpt.FormulaFields(I).TEXT = "'" & txt(Chassis) & "'"
        Case UCase("ENG_NO")
            rpt.FormulaFields(I).TEXT = "'" & txt(Engine) & "'"
        Case UCase("CLAIM_NO")
            rpt.FormulaFields(I).TEXT = "'" & txt(WarrNo) & "'"
        Case UCase("SERV_TYPE")
            rpt.FormulaFields(I).TEXT = "'" & txt(SrvType) & "'"
        Case UCase("OwnerName")
            rpt.FormulaFields(I).TEXT = "'" & txt(OwnerName) & "'"
        Case UCase("TITLE1")
            rpt.FormulaFields(I).TEXT = "'" & mDocStr & "'"
        Case UCase("QUANIRR")
            rpt.FormulaFields(I).TEXT = "'" & mQuanIRR & "'"
        Case UCase("COMM_MES")
            rpt.FormulaFields(I).TEXT = "'" & mComm_Mes & "'"
        Case UCase("FOOTMES")
            rpt.FormulaFields(I).TEXT = "'" & mFootMes & "'"
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
            End Select
        Next
        rpt.PrintOut False
    Case PScreen  'screen
        Call Report_View(rpt, "", , True)
End Select
Set Rst = Nothing
Set rpt = Nothing
CmdPrint(PSetUp).Tag = ""
'TopCtrl1_eAdd
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

Private Sub SpeedPrint(mQry As String)
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
    Dim I As Integer, j As Integer
    Dim PrintStr$
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstReqi As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, RepTitle$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double, FormulaStr1$, FormulaStr2$
    Dim fob As New FileSystemObject, SecondStr As Boolean
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, mAdd$
    Dim TOTAmt As Double
    Dim RstTmp As ADODB.Recordset, Supplier As String
    
'    mQRY = "Select Sp_Stock.Srl_No,p.Part_No,P.Part_Name,SP_Stock.Rate,P.Bin_Loca,SP_Stock.Purpose,SP_Stock.Qty_Doc AS QReq,SP_Stock.Qty_Iss AS QIss,SP_Stock.Qty_Rec AS QRec,SP_Stock.TAX_YN,SP_Stock.Printed,SP_Stock.Amount " & _
'        " From ((SP_Stock Left Join Part P On SP_Stock.Part_No=P.Part_No and P.Div_Code = left(SP_Stock.DocID,1)) " & _
'        " Left Join Emp_Mast on SP_Stock.Mech_Code=Emp_Mast.Emp_Code) " & _
'        " Where SP_Stock.DocID='" & Master!searchcode & "' Order By SP_Stock.Srl_No"
    Set RstReqi = GCn.Execute(mQry)
    
    If RstReqi.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    PageLength = PubPageLengthHalf
    'PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
        
    mHeader = 0   'Ideal 17
    mFooter = 4
    
   'Header
    If Me.CAPTION = "Requisition &Returns" Then
        mDocStr = "GOODS RETURN SLIP (WORKS TO STORES)"
    Else
       mDocStr = "REQUISITION SLIP"
    End If
    mAdd = Trim(PubComp_Add) & IIf(Trim(PubComp_Add2) = "", "", "," & Trim(PubComp_Add2))
    mDocStr = "** " & mDocStr & IIf(RstReqi!Printed = 1, " (DUPLICATE)", "") & " **"
    Print #1, Chr(27) + Chr(67) + Chr(36) & PRN_TIT(PubComp_Name, "A", PageWidth) 'small paper size
    mHeader = mHeader + 1
    If mAdd <> "" Then
        Print #1, PRN_TIT(mAdd, "C", PageWidth)
        mHeader = mHeader + 1
    End If
'    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
'    mHeader = mHeader + 1
'    If PubComp_Add2 <> "" Then
'        Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
'        mHeader = mHeader + 1
'    End If
    If PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(mDocStr, "B", PageWidth) & mChr18
    mHeader = mHeader + 1
    Print #1, PSTR("MODEL", 5) & ": " & mChr17 & PSTR(txt(Model), 15) & mChr18 & "  " & PSTR("VEH.NO.", 8) & ": " & PSTR(txt(VehRegNo), 14) & Space(1) & mEmph & PSTR("REQ. NO.& Dt.", 14) & ": " & Replace(mID(txt(DocID), 9, 15), " ", "") & " " & Format(txt(VDate), "DD/MM/YY") & mEmph1
    mHeader = mHeader + 1
    Print #1, mEmph & PSTR("CHASSIS NO. ", 12) & ": " & PSTR(txt(Chassis), 20) & mEmph1 & Space(6) & PSTR("JOB NO. & Dt.", 14) & ": " & PrinID(txt(JobNo).Tag) & "  " & Format(txt(JobDt), "DD/MM/YY")
    mHeader = mHeader + 1
    Print #1, PSTR("ENGINE NO. ", 12) & ": " & PSTR(txt(Engine), 20) & Space(6) & PSTR("SERVICE TYPE", 14) & ": " & txt(SrvType)
    mHeader = mHeader + 1
    Print #1, PSTR("CLAIM NO. & Dt. ", 16) & ": " & PSTR(txt(WarrNo) & " " & txt(WarrDate), 20) & Space(6) & "OWNER'S NAME  : " & mChr17 & txt(OwnerName) & mChr18
    mHeader = mHeader + 1
    Print #1, mChr17 & IIf(Me.CAPTION = "Requisition &Returns", "Please acknowledge the following Goods as Return against Issue to Workshop", "Please issue the following goods against the original copy of this Requisition Slip") & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-") ' & mDoub
    mHeader = mHeader + 1
    Print #1, mChr17 & PSTR("SRL.", 4) & PSTR("PART NO.", 18) & PSTR("DESCRIPTION", 25) & "BIN       " & PSTR("PUR-", 4) & " <-----QUANTITY----- > " & PSTR("VALUE", 12, , AlignRight) & PSTR("Pur. No.", 10, , AlignRight) & PSTR("Pur.Dt.", 12, , AlignRight) & Space(6) & PSTR("Supplier", 25, , AlignLeft) & mChr18    ' & " T"
    mHeader = mHeader + 1
    Print #1, mChr17 & PSTR("NO.", 4) & Space(18) & Space(25) & "LOCATION " & PSTR("POSE", 4) & PSTR("REQ", 8, , AlignRight) & PSTR("ISS", 8, , AlignRight) & PSTR("RET", 8, , AlignRight) & PSTR("APPROX", 12, , AlignRight) & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    mFix = PageLength - (mHeader + mFooter) - 2
    Page = 1
    mLine = 1
    mSlNo = 1
    If RstReqi.RecordCount > 0 Then
        I = 1
        Do Until RstReqi.EOF
            If mLine > mFix Then
                Page = Page + 1
                Print #1, mChr18 & Replace(Space(PageWidth), " ", "-")
                Print #1, Space(PageWidth - Len("Contd. on next page.." + STR(Page))) & "Contd. on next page.." & STR(Page)
                Do Until mLine >= mFix + mFooter - 2
                    Print #1, ""
                    mLine = mLine + 1
                Loop
                Print #1, mEject
                
                'Header On Second Page
                mHeader = 0
                Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, PRN_TIT(" & mDocStr & ", "B", PageWidth) + mChr18
                mHeader = mHeader + 1
                Print #1, PSTR("MODEL", 5) & ": " & mChr17 & PSTR(txt(Model), 15) & mChr18 & "  " & PSTR("VEH.NO.", 8) & ": " & PSTR(txt(VehRegNo), 14) & Space(6) & mEmph & PSTR("REQ. NO.& Dt.", 14) & ": " & Replace(mID(txt(DocID), 9, 15), " ", "") & " " & Format(txt(VDate), "DD/MM/YY") & mEmph1
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-") ' & mDoub
                mHeader = mHeader + 1
                Print #1, mChr17 & PSTR("SRL.", 4) & PSTR("PART NO.", 18) & PSTR("DESCRIPTION", 25) & "BIN       " & PSTR("PUR-", 4) & " <-----QUANTITY----- > " & PSTR("VALUE", 12, , AlignRight) & PSTR("Pur. No.", 10, , AlignRight) & PSTR("Pur.Dt.", 12, , AlignRight) & Space(2) & PSTR("Supplier", 25, , AlignLeft) & mChr18   ' & " T"
                mHeader = mHeader + 1
                Print #1, mChr17 & PSTR("NO.", 4) & Space(18) & Space(25) & "LOCATION " & PSTR("POSE", 4) & PSTR("REQ", 8, , AlignRight) & PSTR("ISS", 8, , AlignRight) & PSTR("RET", 8, , AlignRight) & PSTR("APPROX", 12, , AlignRight) & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                mFix = PageLength - (mHeader + mFooter)
                mLine = 1
            End If
            If XNull(RstReqi!PurDocId) <> "" Then
                Set RstTmp = GCn.Execute("Select Party_Name from SP_Purch where Party_Doc_No='" & XNull(RstReqi!PurDocId) & "'")
                If RstTmp.RecordCount > 0 Then
                    Supplier = XNull(RstTmp!Party_Name)
                End If
            End If
            Set RstTmp = Nothing
            TOTAmt = TOTAmt + RstReqi!Amount
            PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 4) & mChr17 & PSTR(RstReqi!Part_No, 18, , AlignLeft) & PSTR(RstReqi!Part_Name, 25, , AlignLeft) & PSTR(RstReqi!Bin_Loca, 8, , AlignLeft) & " " & PSTR(RstReqi!Purpose, 4, , AlignLeft) & PSTR(RstReqi!qreq, 8, 2, AlignLeft) & PSTR(RstReqi!qiss, 8, 2, AlignLeft) & PSTR(RstReqi!qrec, 8, 2, AlignLeft) & PSTR(RstReqi!Amount, 12, 2, AlignLeft) & _
                       "" & Space(2) & PSTR(XNull(RstReqi!PurDocId), 12, 2, AlignLeft) & PSTR(XNull(RstReqi!PurDate), 12, 2, AlignLeft) & PSTR(XNull(Supplier), 25, 2, AlignLeft) & mChr18
            Print #1, mChr17 & PrintStr & mChr18
            RstReqi.MoveNext
            mSlNo = mSlNo + 1
            mLine = mLine + 1
        Loop
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop
    
    ' FOOTER
    RstReqi.MoveFirst
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, PSTR("Total Amount : ", 17) & Format(TOTAmt, "0.00")
    Print #1, IIf(Me.CAPTION = "Requisition &Returns", "Goods Received By:" & pubUName & "  Date:" & Format(date, "dd/mm/yy") & "  Time:" & Format(time, "hh:mm") & "  Mechanic:" & txt(Mechanic), "Mechanic : " & txt(Mechanic) & Space(10) & "Supervisor/Manager" & Space(10) & "Stores")
    'Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
    Print #1, mChr17 & "*a dataman software*" & Space((PageWidth * 1.7) - Len("* a dataman software *" & pubUName & "   " & PubServerDate)) & pubUName & "   " & PubServerDate & mChr18
    If UCase(left(PubComp_Name, 7)) = "SHANKAR" Or UCase(left(PubComp_Name, 6)) = "MAURYA" Then
    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
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
Private Sub StkUpd(PNo As String)
Dim I As Integer
Dim mSQry$, mQry$
    Dim Rst As ADODB.Recordset
    GCn.BeginTrans
            
        mSQry = "Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock " & _
                "WHERE (V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " " & _
                "Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & ") " & _
                "And Part_No='" & PNo & "' "
    
        
        If PubBackEnd = "S" Then
            mQry = "Select P.Part_No as Code, P.Part_No, P.Part_Name As Name, P.Local_Name as LName, P.Unit , P.MRP, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, " & _
                            "(Select PurcDisc_Per From Part_DiscFactor Where DiscFac_Catg=P.Disc_Factor) As PurcDisc_Per, P.ReOrd_Lvl, " & _
                            "(" & mSQry & " And Mrp_Yn=1 And Tax_Yn=1) As Cur_MRP_TBStk, (" & mSQry & " And Mrp_Yn=1 And Tax_Yn=0) As Cur_MRP_TpStk, " & _
                            "(" & mSQry & " And Mrp_Yn=0 And Tax_Yn=1) As Cur_TBStk, (" & mSQry & " And Mrp_Yn=0 And Tax_Yn=0) As Cur_TpStk, " & _
                            "(" & mSQry & ") As CurrStk, P.Min_Lvl, P.Disc_Factor " & _
                            "From Part P Left Join Sp_Stock Stk On P.Part_No=Stk.Part_No " & _
                            "WHERE (V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & " Or Stk.Part_No Is Null)  And Div_Code='" & PubDivCode & "' and P.Part_No='" & PNo & "' " & _
                            "Group By P.Part_No, P.Part_Name, P.Local_Name, P.Unit, P.Mrp, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Disc_Factor, P.ReOrd_Lvl, P.Min_Lvl"
        Else
            mQry = "Select P.Part_No as Code, P.Part_No, P.Part_Name As Name, P.Local_Name as LName, P.Unit , Format(P.MRP,'0.00') As Mrp, Format(P.TB_SRate,'0.00') As TB_SRate, Format(P.Tp_SRate,'0.00') As Tp_SRate, P.Bin_Loca, " & _
                            "(Select PurcDisc_Per From Part_DiscFactor Where DiscFac_Catg=P.Disc_Factor) As PurcDisc_Per, P.ReOrd_Lvl, " & _
                            "(" & mSQry & " And Mrp_Yn=1 And Tax_Yn=1) As Cur_MRP_TBStk, (" & mSQry & " And Mrp_Yn=1 And Tax_Yn=0) As Cur_MRP_TpStk, " & _
                            "(" & mSQry & " And Mrp_Yn=0 And Tax_Yn=1) As Cur_TBStk, (" & mSQry & " And Mrp_Yn=0 And Tax_Yn=0) As Cur_TpStk, " & _
                            "(" & mSQry & ") As CurrStk, P.Min_Lvl, P.Disc_Factor " & _
                            "From Part P Left Join Sp_Stock Stk On P.Part_No=Stk.Part_No " & _
                            "WHERE (V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & " Or Stk.Part_No Is Null)  And Div_Code='" & PubDivCode & "' And P.Part_No = '" & PNo & "' " & _
                            "Group By P.Part_No, P.Part_Name, P.Local_Name, P.Unit, P.Mrp, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Disc_Factor, P.ReOrd_Lvl, P.Min_Lvl"
        End If
        
        
        Set Rst = GCn.Execute(mQry)
        If Rst.RecordCount > 0 Then
            With Rst
                GCn.Execute ("Update Part Set Part.Cur_TP_Stk=" & VNull(!Cur_TPStk) & ", " & _
                             "Part.Cur_TB_Stk=" & VNull(!Cur_TBStk) & ", Part.Cur_Mrp_TpStk=" & VNull(!Cur_MRP_TPStk) & ", " & _
                             "Part.Cur_Mrp_TBStk=" & VNull(!Cur_MRP_TbStk) & " where Part.Part_No='" & !Part_No & "'and Part.Div_Code='" & PubDivCode & "'")
            End With
        End If
    GCn.CommitTrans

Set Rst = Nothing
Exit Sub
End Sub
Private Function FIFOStkIss(Part_No As String)
Dim RstTmp As ADODB.Recordset
Dim TBRst As ADODB.Recordset, TPRst As ADODB.Recordset, TmpRst As ADODB.Recordset
Dim MRPTBRst As ADODB.Recordset, MRPTPRst As ADODB.Recordset
Dim TBCurrStk, TPCurrStk As Double, I As Double
Dim MRPTBCurrStk, MRPTPCurrStk As Double
Dim TBStkDate$, TBStkPurDocId$, TBStkPurDate$, TPStkDate$, TPStkPurDocId$, TPStkPurDate$
Dim MRPTBStkDate$, MRPTBStkPurDocId$, MRPTBStkPurDate$, MRPTPStkDate$, MRPTPStkPurDocId$, MRPTPStkPurDate$

'Set TBRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec,iif(SP_Stock.V_Type='SXAO',SP_Stock.PurDocNo,SP_Purch.Party_Doc_No) as Party_Doc_No,iif(SP_Stock.V_Type='SXAO',SP_Stock.PurDocDate,SP_Purch.Party_Doc_Date) as Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId where Qty_Rec > 0 and Tax_YN=1 and MRP_YN=0 and Part_No='" & Part_No & "'")
Set TBRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec," & cIIF("SP_Stock.PurDocNo='' or SP_Stock.PurDocNo Is Null", "SP_Purch.Party_Doc_No", "SP_Stock.PurDocNo") & " as Party_Doc_No, " & cIIF("SP_Stock.PurDocNo='' or SP_Stock.PurDocNo Is Null", "SP_Stock.PurDocDate", "SP_Purch.Party_Doc_Date") & " as Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId where Qty_Rec > 0 and Tax_YN=1 and MRP_YN=0 and Part_No='" & Part_No & "'")
'Set TPRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec,iif(SP_Stock.V_Type='SXAO',SP_Stock.PurDocNo,SP_Purch.Party_Doc_No) as Party_Doc_No,iif(SP_Stock.V_Type='SXAO',SP_Stock.PurDocDate,SP_Purch.Party_Doc_Date) as Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId  where Qty_Rec > 0 and Tax_YN=0 and MRP_YN=0 and Part_No='" & Part_No & "'")
Set TPRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec," & cIIF("SP_Stock.PurDocNo='' or SP_Stock.PurDocNo Is Null", "SP_Purch.Party_Doc_No", "SP_Stock.PurDocNo") & " as Party_Doc_No, " & cIIF("SP_Stock.PurDocNo='' or SP_Stock.PurDocNo Is Null", "SP_Purch.Party_Doc_Date", "SP_Stock.PurDocDate") & " as Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId  where Qty_Rec > 0 and Tax_YN=0 and MRP_YN=0 and Part_No='" & Part_No & "'")
'Set MRPTBRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec,iif(SP_Stock.V_Type='SXAO',SP_Stock.PurDocNo,SP_Purch.Party_Doc_No) as Party_Doc_No,iif(SP_Stock.V_Type='SXAO',SP_Stock.PurDocDate,SP_Purch.Party_Doc_Date) as Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId  where Qty_Rec > 0 and Tax_YN=1 and MRP_YN=1 and Part_No='" & Part_No & "'")
Set MRPTBRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec," & cIIF("SP_Stock.PurDocNo='' or SP_Stock.PurDocNo Is Null", "SP_Purch.Party_Doc_No", "SP_Stock.PurDocNo") & " as Party_Doc_No, " & cIIF("SP_Stock.PurDocNo='' or SP_Stock.PurDocNo Is Null) ", "SP_Purch.Party_Doc_Date", "SP_Stock.PurDocDate") & " as Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId  where Qty_Rec > 0 and Tax_YN=1 and MRP_YN=1 and Part_No='" & Part_No & "'")
'Set MRPTPRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec,iif(SP_Stock.V_Type='SXAO',SP_Stock.PurDocNo,SP_Purch.Party_Doc_No) as Party_Doc_No,iif(SP_Stock.V_Type='SXAO',SP_Stock.PurDocDate,SP_Purch.Party_Doc_Date) as Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId  where Qty_Rec > 0 and Tax_YN=0 and MRP_YN=1 and Part_No='" & Part_No & "'")
Set MRPTPRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec," & cIIF("SP_Stock.PurDocNo='' or SP_Stock.PurDocNo Is Null", "SP_Purch.Party_Doc_No", "SP_Stock.PurDocNo") & " as Party_Doc_No, " & cIIF("SP_Stock.PurDocNo='' or SP_Stock.PurDocNo Is Null", "SP_Purch.Party_Doc_Date", "SP_Stock.PurDocDate") & " as Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId  where Qty_Rec > 0 and Tax_YN=0 and MRP_YN=1 and Part_No='" & Part_No & "'")

TBCurrStk = VNull(GCn.Execute("Select sum(Qty_Rec-(Qty_Iss-Qty_Ret)) as CurrStkTB from Sp_Stock where Tax_YN=1 and MRP_YN=0 and Part_No='" & Part_No & "'").Fields(0).Value)
TPCurrStk = VNull(GCn.Execute("Select sum(Qty_Rec-(Qty_Iss-Qty_Ret)) as CurrStkTP from Sp_Stock where Tax_YN=0 and MRP_YN=0 and Part_No='" & Part_No & "'").Fields(0).Value)
MRPTBCurrStk = VNull(GCn.Execute("Select sum(Qty_Rec-(Qty_Iss-Qty_Ret)) as CurrStkTB from Sp_Stock where Tax_YN=1 and MRP_YN=1 and Part_No='" & Part_No & "'").Fields(0).Value)
MRPTPCurrStk = VNull(GCn.Execute("Select sum(Qty_Rec-(Qty_Iss-Qty_Ret)) as CurrStkTP from Sp_Stock where Tax_YN=0 and MRP_YN=1 and Part_No='" & Part_No & "'").Fields(0).Value)

If TBRst.RecordCount > 0 Then
    TBRst.Sort = "V_Date Desc"
    For I = 1 To TBRst.RecordCount
        If TBCurrStk >= TBRst!Qty_Rec Then
            TBStkDate = XNull(TBRst!V_DATE)
            TBStkPurDocId = XNull(TBRst!Party_Doc_No)
            TBStkPurDate = XNull(TBRst!Party_Doc_Date)
            Exit For
        Else
            TBCurrStk = TBCurrStk - TBRst!Qty_Rec
            TBRst.MoveNext
        End If
    Next
End If
If TPRst.RecordCount > 0 Then
    TPRst.Sort = "V_Date Desc"
    For I = 1 To TPRst.RecordCount
        If TPCurrStk >= TPRst!Qty_Rec Then
            TPStkDate = XNull(TPRst!V_DATE)
            TPStkPurDocId = XNull(TPRst!Party_Doc_No)
            TPStkPurDate = XNull(TPRst!Party_Doc_Date)
            Exit For
        Else
            TPCurrStk = TPCurrStk - TPRst!Qty_Rec
            TPRst.MoveNext
        End If
    Next
End If
If MRPTBRst.RecordCount > 0 Then
    MRPTBRst.Sort = "V_Date Desc"
    For I = 1 To MRPTBRst.RecordCount
        If MRPTBCurrStk >= MRPTBRst!Qty_Rec Then
            MRPTBStkDate = XNull(MRPTBRst!V_DATE)
            MRPTBStkPurDocId = XNull(MRPTBRst!Party_Doc_No)
            MRPTBStkPurDate = XNull(MRPTBRst!Party_Doc_Date)
            Exit For
        Else
            MRPTBCurrStk = MRPTBCurrStk - MRPTBRst!Qty_Rec
            MRPTBRst.MoveNext
        End If
    Next
End If
If MRPTPRst.RecordCount > 0 Then
    MRPTPRst.Sort = "V_Date Desc"
    For I = 1 To MRPTPRst.RecordCount
        If MRPTPCurrStk >= MRPTPRst!Qty_Rec Then
            MRPTPStkDate = XNull(MRPTPRst!V_DATE)
            MRPTPStkPurDocId = XNull(MRPTPRst!Party_Doc_No)
            MRPTPStkPurDate = XNull(MRPTPRst!Party_Doc_Date)
            Exit For
        Else
            MRPTPCurrStk = MRPTPCurrStk - MRPTPRst!Qty_Rec
            MRPTPRst.MoveNext
        End If
    Next
End If
Set TmpRst = New ADODB.Recordset
With TmpRst
    .Fields.Append "xDate", adVarChar, 11
    .Fields.Append "PurDocId", adVarChar, 21
    .Fields.Append "PurDate", adVarChar, 11
    .Fields.Append "Cond", adVarChar, 2
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open
End With
DoEvents
If TPStkDate <> "" Then
    TmpRst.AddNew
    TmpRst.Fields(0) = TPStkDate: TmpRst!cond = "00"
    TmpRst!PurDocId = TPStkPurDocId
    TmpRst!PurDate = TPStkPurDate
End If
If MRPTPStkDate <> "" Then
    TmpRst.AddNew
    TmpRst.Fields(0) = MRPTPStkDate: TmpRst!cond = "10"
    TmpRst!PurDocId = MRPTPStkPurDocId
    TmpRst!PurDate = MRPTPStkPurDate
End If
If TBStkDate <> "" Then
    TmpRst.AddNew
    TmpRst.Fields(0) = TBStkDate: TmpRst!cond = "01"
    TmpRst!PurDocId = TBStkPurDocId
    TmpRst!PurDate = TBStkPurDate
End If
If MRPTBStkDate <> "" Then
    TmpRst.AddNew
    TmpRst.Fields(0) = MRPTBStkDate: TmpRst!cond = "11"
    TmpRst!PurDocId = MRPTBStkPurDocId
    TmpRst!PurDate = MRPTBStkPurDate
End If
TmpRst.Sort = "xDate"
If TmpRst.RecordCount > 0 Then
    TmpRst.MoveFirst
    Select Case TmpRst!cond
        Case "00"
            FGrid.TextMatrix(FGrid.Row, Col_MRP) = "No"
            FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "No"
            FGrid.TextMatrix(FGrid.Row, Col_PurDocId) = TPStkPurDocId
            FGrid.TextMatrix(FGrid.Row, Col_PurDate) = TPStkPurDate
        Case "01"
            FGrid.TextMatrix(FGrid.Row, Col_MRP) = "No"
            FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes"
            FGrid.TextMatrix(FGrid.Row, Col_PurDocId) = TBStkPurDocId
            FGrid.TextMatrix(FGrid.Row, Col_PurDate) = TBStkPurDate
        Case "10"
            FGrid.TextMatrix(FGrid.Row, Col_MRP) = "Yes"
            FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "No"
            FGrid.TextMatrix(FGrid.Row, Col_PurDocId) = MRPTPStkPurDocId
            FGrid.TextMatrix(FGrid.Row, Col_PurDate) = MRPTPStkPurDate
        Case "11"
            FGrid.TextMatrix(FGrid.Row, Col_MRP) = "Yes"
            FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes"
            FGrid.TextMatrix(FGrid.Row, Col_PurDocId) = MRPTBStkPurDocId
            FGrid.TextMatrix(FGrid.Row, Col_PurDate) = MRPTBStkPurDate
    End Select
    If FGrid.TextMatrix(FGrid.Row, Col_PurDocId) <> "" Then
        Set RstTmp = GCn.Execute("Select Party_Name from SP_Purch where Party_Doc_No='" & FGrid.TextMatrix(FGrid.Row, Col_PurDocId) & "'")
        If RstTmp.RecordCount > 0 Then
            FGrid.TextMatrix(FGrid.Row, Col_Supplier) = XNull(RstTmp!Party_Name)
        End If
    End If
    Set RstTmp = Nothing
End If
End Function

Sub Ini_Pub()
    Dim RsTemp As ADODB.Recordset
    
    Set RsTemp = GCn.Execute("Select CheckNegetiveStockSiteWise From Syctrl")
    If RsTemp.RecordCount > 0 Then
        mCheckNegetiveStockSiteWise = VNull(RsTemp!CheckNegetiveStockSiteWise)
    End If
End Sub



Sub IniGrid_Vat()
    With FGrid
        If PubVATYN = 1 Then
            .TextMatrix(0, Col_TaxPer) = "TaxPer"
            .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
            .ColWidth(Col_TaxPer) = 840
            
            .TextMatrix(0, Col_TaxAmt) = "TaxAmt"
            .ColAlignmentFixed(Col_TaxAmt) = flexAlignRightCenter
            .ColWidth(Col_TaxAmt) = 840
                        
'            If mSatYn Then
'                .TextMatrix(0, Col_SatPer) = "SAT %"
'                .ColAlignmentFixed(Col_SatPer) = flexAlignRightCenter
'                .ColWidth(Col_SatPer) = 840
'
'                .TextMatrix(0, Col_SatAmt) = "SAT Amt"
'                .ColAlignmentFixed(Col_SatAmt) = flexAlignRightCenter
'                .ColWidth(Col_SatAmt) = 840
'            Else
'                .ColWidth(Col_SatPer) = 0
'                .ColWidth(Col_SatAmt) = 0
'            End If
        Else
            .TextMatrix(0, Col_TaxPer) = ""
            .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
            .ColWidth(Col_TaxPer) = 0
            
            .TextMatrix(0, Col_TaxAmt) = ""
            .ColAlignmentFixed(Col_TaxAmt) = flexAlignRightCenter
            .ColWidth(Col_TaxAmt) = 0
            
            .ColWidth(Col_SatPer) = 0
            .ColWidth(Col_SatAmt) = 0
        End If
    End With
End Sub
