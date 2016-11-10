VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmSaleRet 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Goods Return (Inward)"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12945
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
   ScaleHeight     =   9255
   ScaleWidth      =   12945
   Visible         =   0   'False
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   51
      Left            =   10260
      TabIndex        =   193
      Top             =   5325
      Width           =   1215
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Re-Post"
      Height          =   345
      Left            =   8205
      TabIndex        =   192
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox lblGroup 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   330
      Left            =   105
      TabIndex        =   191
      Top             =   3180
      Visible         =   0   'False
      Width           =   1785
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
      Left            =   5040
      TabIndex        =   176
      Top             =   2610
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
         Picture         =   "frmSaleRet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   186
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
         Picture         =   "frmSaleRet.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   185
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmSaleRet.frx":0678
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
         TabIndex        =   184
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmSaleRet.frx":0982
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
         TabIndex        =   183
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmSaleRet.frx":0C8C
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
         TabIndex        =   182
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
         TabIndex        =   181
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
         TabIndex        =   180
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
         TabIndex        =   179
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
         TabIndex        =   178
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
         TabIndex        =   177
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
         TabIndex        =   189
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
         TabIndex        =   188
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
         TabIndex        =   187
         Top             =   0
         Width           =   4695
      End
   End
   Begin VB.TextBox Txt 
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   18
      Left            =   7695
      TabIndex        =   175
      Text            =   "Addition"
      Top             =   1485
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox Txt 
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   33
      Left            =   5445
      TabIndex        =   174
      Text            =   "Addition"
      Top             =   1635
      Visible         =   0   'False
      Width           =   525
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   50
      Left            =   10260
      TabIndex        =   47
      Top             =   6405
      Width           =   1215
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   49
      Left            =   9585
      TabIndex        =   46
      Top             =   6405
      Width           =   600
   End
   Begin MSDataGridLib.DataGrid DGPerson 
      Height          =   3330
      Left            =   -6285
      Negotiate       =   -1  'True
      TabIndex        =   172
      TabStop         =   0   'False
      Top             =   4215
      Visible         =   0   'False
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   5874
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Sales Person"
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
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   5265.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   494.929
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   2670
      Left            =   -4080
      Negotiate       =   -1  'True
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   9045
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
   Begin VB.Frame FrmDetail 
      BackColor       =   &H00CAF1FD&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   2205
      Left            =   -5970
      TabIndex        =   140
      Top             =   3915
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
         Index           =   12
         Left            =   3765
         TabIndex        =   171
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
         TabIndex        =   170
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
         TabIndex        =   169
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
         TabIndex        =   168
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
         TabIndex        =   167
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
         TabIndex        =   166
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
         TabIndex        =   165
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
         TabIndex        =   164
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
         TabIndex        =   163
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
         TabIndex        =   162
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
         TabIndex        =   161
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
         TabIndex        =   160
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
         TabIndex        =   159
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
         TabIndex        =   158
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
         TabIndex        =   157
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
         TabIndex        =   156
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
         TabIndex        =   155
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
         TabIndex        =   154
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
         TabIndex        =   153
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
         TabIndex        =   152
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
         TabIndex        =   151
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
         Index           =   14
         Left            =   2805
         TabIndex        =   150
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
         Index           =   13
         Left            =   75
         TabIndex        =   149
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
         Index           =   11
         Left            =   75
         TabIndex        =   148
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
         Index           =   10
         Left            =   3930
         TabIndex        =   147
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
         Index           =   9
         Left            =   0
         TabIndex        =   146
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
         Index           =   8
         Left            =   75
         TabIndex        =   145
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
         Index           =   0
         Left            =   75
         TabIndex        =   144
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
         TabIndex        =   143
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
         Index           =   1
         Left            =   75
         TabIndex        =   142
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
         TabIndex        =   141
         Top             =   1395
         Width           =   360
      End
      Begin VB.Line Line4 
         X1              =   3660
         X2              =   3885
         Y1              =   1035
         Y2              =   1035
      End
   End
   Begin MSDataGridLib.DataGrid DGBaseDoc 
      Height          =   2775
      Left            =   4785
      Negotiate       =   -1  'True
      TabIndex        =   138
      TabStop         =   0   'False
      Top             =   7980
      Visible         =   0   'False
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   4895
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Document No."
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
         Caption         =   "DocID"
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
         DataField       =   "V_Date"
         Caption         =   "Doc. Date"
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
            ColumnWidth     =   1844.787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1454.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   16
      Left            =   7020
      MaxLength       =   3
      TabIndex        =   16
      Text            =   "Yes"
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   1950
      Visible         =   0   'False
      Width           =   525
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12945
      _ExtentX        =   22834
      _ExtentY        =   661
   End
   Begin MSDataGridLib.DataGrid DGGodown 
      Height          =   3330
      Left            =   9270
      Negotiate       =   -1  'True
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   8445
      Visible         =   0   'False
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   5874
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
            DividerStyle    =   3
            ColumnWidth     =   5220.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGSONo 
      Height          =   2775
      Left            =   4260
      Negotiate       =   -1  'True
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   8115
      Visible         =   0   'False
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   4895
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
         DataField       =   "Name"
         Caption         =   "Sale Order No"
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
         Caption         =   "OrderID"
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
         DataField       =   "V_Date"
         Caption         =   "Order Date"
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
         DataField       =   "Qty"
         Caption         =   "Order Qty"
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
         DataField       =   "Rate"
         Caption         =   "Rate"
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
         DataField       =   "MRPYN"
         Caption         =   "MRP"
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
         DataField       =   "TAXYN"
         Caption         =   "Taxable"
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
            ColumnWidth     =   3195.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   945.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGForm31 
      Height          =   3330
      Left            =   5550
      Negotiate       =   -1  'True
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   8895
      Visible         =   0   'False
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1.5
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
         DataField       =   "Name"
         Caption         =   "Form 31 Name"
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
         Caption         =   "Form Code"
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
            ColumnWidth     =   5220.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00F1C7CF&
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
      Index           =   10
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   10
      Top             =   1950
      Width           =   1080
   End
   Begin VB.TextBox Txt 
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
      Index           =   9
      Left            =   1455
      MaxLength       =   40
      TabIndex        =   9
      Top             =   1950
      Width           =   1515
   End
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   3330
      Left            =   5550
      Negotiate       =   -1  'True
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   8865
      Visible         =   0   'False
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   5874
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Form Name"
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
         Caption         =   "Form Code"
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
         DataField       =   "Tax_Per"
         Caption         =   "Tax%"
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
         DataField       =   "Tax_Sur_Per"
         Caption         =   "S.Charge%"
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
            ColumnWidth     =   5295.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCrAc 
      Height          =   3330
      Left            =   1395
      Negotiate       =   -1  'True
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   9105
      Visible         =   0   'False
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   5874
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
      ColumnCount     =   4
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
      BeginProperty Column01 
         DataField       =   "Code"
         Caption         =   "Ac.Code"
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
         DataField       =   "Add1"
         Caption         =   "Address"
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
         DataField       =   "CityName"
         Caption         =   "City"
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
            ColumnWidth     =   4305.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3344.882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2399.811
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   3330
      Left            =   -2295
      Negotiate       =   -1  'True
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   9015
      Visible         =   0   'False
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   5874
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
      ColumnCount     =   4
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
      BeginProperty Column01 
         DataField       =   "Code"
         Caption         =   "Ac.Code"
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
         DataField       =   "Add1"
         Caption         =   "Address"
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
         DataField       =   "CityName"
         Caption         =   "City"
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
            ColumnWidth     =   5220.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2250.142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2190.047
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
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
      Index           =   14
      Left            =   6405
      MaxLength       =   7
      TabIndex        =   14
      ToolTipText     =   "Press L-> Local or C-> Central"
      Top             =   1095
      Width           =   1050
   End
   Begin VB.TextBox Txt 
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
      Index           =   4
      Left            =   1035
      MaxLength       =   40
      TabIndex        =   4
      Top             =   525
      Width           =   3930
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   9375
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   525
      Width           =   2415
   End
   Begin VB.TextBox Txt 
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   15
      Left            =   6405
      MaxLength       =   3
      TabIndex        =   15
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   1380
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   -345
      TabIndex        =   112
      Top             =   8085
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   60
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   225
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
            Size            =   8.25
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
      Index           =   13
      Left            =   6405
      MaxLength       =   10
      TabIndex        =   13
      Top             =   810
      Width           =   2040
   End
   Begin VB.TextBox Txt 
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
      Index           =   8
      Left            =   1455
      MaxLength       =   40
      TabIndex        =   8
      Top             =   1665
      Width           =   2715
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   48
      Left            =   5190
      TabIndex        =   50
      Top             =   6765
      Width           =   1440
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   38
      Left            =   10260
      TabIndex        =   37
      Top             =   4785
      Width           =   1215
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
      Enabled         =   0   'False
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
      Height          =   255
      Index           =   47
      Left            =   10260
      TabIndex        =   49
      Top             =   6945
      Width           =   1215
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
      Enabled         =   0   'False
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
      Index           =   22
      Left            =   5115
      TabIndex        =   23
      Top             =   5055
      Width           =   1215
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
      Enabled         =   0   'False
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
      Index           =   21
      Left            =   2985
      TabIndex        =   22
      Top             =   5055
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      Index           =   20
      Left            =   5115
      TabIndex        =   21
      Top             =   4785
      Width           =   1215
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
      Left            =   2445
      TabIndex        =   18
      Top             =   2955
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Txt 
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
      Index           =   5
      Left            =   1035
      MaxLength       =   40
      TabIndex        =   5
      Top             =   810
      Width           =   3930
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   9570
      TabIndex        =   1
      Top             =   1080
      Width           =   2220
   End
   Begin VB.TextBox Txt 
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
      Index           =   6
      Left            =   1035
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1095
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.TextBox Txt 
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   17
      Left            =   7020
      MaxLength       =   50
      TabIndex        =   17
      Top             =   2235
      Width           =   4770
   End
   Begin VB.TextBox Txt 
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1380
      Width           =   2085
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
      Index           =   46
      Left            =   10260
      TabIndex        =   48
      Top             =   6675
      Width           =   1215
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   45
      Left            =   10260
      TabIndex        =   45
      Top             =   6135
      Width           =   1215
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   44
      Left            =   9585
      TabIndex        =   44
      ToolTipText     =   "Turn Over Tax %"
      Top             =   6135
      Width           =   600
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
      Enabled         =   0   'False
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
      Index           =   43
      Left            =   10260
      TabIndex        =   43
      Top             =   5865
      Width           =   1215
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   42
      Left            =   10950
      TabIndex        =   41
      Top             =   7590
      Width           =   1215
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   41
      Left            =   10275
      TabIndex        =   40
      ToolTipText     =   "Surcharge % on Local Sales Tax"
      Top             =   7590
      Width           =   600
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
      Index           =   40
      Left            =   10260
      TabIndex        =   39
      Top             =   5055
      Width           =   1215
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   39
      Left            =   9585
      TabIndex        =   38
      ToolTipText     =   "Local Sales Tax %"
      Top             =   5055
      Width           =   600
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   2985
      TabIndex        =   36
      Top             =   6675
      Width           =   1215
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   36
      Left            =   2985
      TabIndex        =   35
      Top             =   6405
      Width           =   1215
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   35
      Left            =   2280
      TabIndex        =   34
      ToolTipText     =   "General Surcharge %"
      Top             =   6405
      Width           =   600
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   19
      Left            =   2985
      TabIndex        =   20
      Top             =   4785
      Width           =   1215
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
      Left            =   10260
      TabIndex        =   42
      Top             =   5595
      Width           =   1215
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
      Enabled         =   0   'False
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
      Index           =   32
      Left            =   5115
      TabIndex        =   33
      Top             =   6135
      Width           =   1215
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
      Enabled         =   0   'False
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
      Index           =   31
      Left            =   2985
      TabIndex        =   32
      Text            =   "99999999.99"
      Top             =   6135
      Width           =   1215
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
      Left            =   5115
      TabIndex        =   31
      Top             =   5865
      Width           =   1215
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
      Left            =   4425
      TabIndex        =   30
      ToolTipText     =   "Discount % Taxpaid"
      Top             =   5865
      Width           =   600
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
      Left            =   2985
      TabIndex        =   29
      Text            =   "99999999.99"
      Top             =   5865
      Width           =   1215
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
      Left            =   2295
      TabIndex        =   28
      Text            =   "99.99"
      ToolTipText     =   "Discount % Taxable"
      Top             =   5865
      Width           =   600
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
      Enabled         =   0   'False
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
      Index           =   26
      Left            =   5115
      TabIndex        =   27
      Top             =   5595
      Width           =   1215
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
      Enabled         =   0   'False
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
      Index           =   25
      Left            =   2985
      TabIndex        =   26
      Top             =   5595
      Width           =   1215
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
      Enabled         =   0   'False
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
      Index           =   24
      Left            =   5115
      TabIndex        =   25
      Top             =   5325
      Width           =   1215
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
      Enabled         =   0   'False
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
      Index           =   23
      Left            =   2985
      TabIndex        =   24
      Top             =   5325
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Index           =   12
      Left            =   6405
      TabIndex        =   12
      Top             =   525
      Width           =   1050
   End
   Begin VB.TextBox Txt 
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
      Index           =   2
      Left            =   9570
      MaxLength       =   11
      TabIndex        =   2
      Top             =   1365
      Width           =   1560
   End
   Begin VB.TextBox Txt 
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
      Index           =   11
      Left            =   1455
      MaxLength       =   40
      TabIndex        =   11
      Top             =   2235
      Width           =   4185
   End
   Begin VB.TextBox Txt 
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
      Height          =   270
      Index           =   3
      Left            =   10230
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1650
      Width           =   900
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1935
      Left            =   60
      TabIndex        =   19
      Top             =   2550
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   14940925
      Cols            =   34
      BackColorFixed  =   13300221
      ForeColorFixed  =   128
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   8438015
      GridColorFixed  =   33023
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "KKKK"
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
      _Band(0).Cols   =   34
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Tax "
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
      Index           =   5
      Left            =   7575
      TabIndex        =   195
      Top             =   5325
      Width           =   1185
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   0
      Left            =   9105
      TabIndex        =   194
      Top             =   5325
      Width           =   180
   End
   Begin VB.Label LblCancel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Cancelled*"
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
      Left            =   6255
      TabIndex        =   190
      Top             =   1710
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReSale Tax             :"
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
      Index           =   42
      Left            =   7575
      TabIndex        =   173
      Top             =   6420
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   29
      Left            =   6900
      TabIndex        =   137
      Top             =   1950
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Posting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   4
      Left            =   5880
      TabIndex        =   136
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   1950
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008080FF&
      Height          =   1485
      Left            =   8490
      Shape           =   4  'Rounded Rectangle
      Top             =   465
      Width           =   3360
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Prefix"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   9570
      TabIndex        =   135
      Top             =   1650
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   36
      Left            =   4470
      TabIndex        =   131
      Top             =   1950
      Width           =   45
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form 31 No"
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
      Height          =   270
      Index           =   36
      Left            =   3090
      TabIndex        =   130
      Top             =   1950
      Width           =   975
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   35
      Left            =   1350
      TabIndex        =   129
      Top             =   1950
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form 31"
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
      Height          =   270
      Index           =   35
      Left            =   180
      TabIndex        =   128
      Top             =   1950
      Width           =   690
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   23
      Left            =   6285
      TabIndex        =   124
      Top             =   1095
      Width           =   180
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dispatch Type"
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
      Height          =   270
      Index           =   3
      Left            =   5010
      TabIndex        =   123
      ToolTipText     =   "Press L-> Local or C-> Central"
      Top             =   1095
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   3
      Left            =   915
      TabIndex        =   122
      Top             =   525
      Width           =   195
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
      Index           =   25
      Left            =   9255
      TabIndex        =   121
      Top             =   525
      Width           =   45
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOC ID"
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
      Height          =   270
      Index           =   31
      Left            =   8580
      TabIndex        =   120
      Top             =   525
      Width           =   585
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party"
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
      Height          =   270
      Index           =   9
      Left            =   180
      TabIndex        =   119
      Top             =   525
      Width           =   405
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print Tax Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   29
      Left            =   4860
      TabIndex        =   117
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   1380
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   21
      Left            =   6285
      TabIndex        =   116
      Top             =   1380
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label LblSite 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code"
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
      Height          =   270
      Left            =   10320
      TabIndex        =   115
      Top             =   825
      Width           =   810
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
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
      Height          =   270
      Left            =   8580
      TabIndex        =   114
      Top             =   825
      Width           =   660
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Case Mark"
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
      Height          =   270
      Index           =   40
      Left            =   5295
      TabIndex        =   111
      Top             =   810
      Width           =   885
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   34
      Left            =   6285
      TabIndex        =   110
      Top             =   810
      Width           =   195
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   32
      Left            =   1350
      TabIndex        =   109
      Top             =   1665
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form"
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
      Height          =   270
      Index           =   39
      Left            =   180
      TabIndex        =   108
      Top             =   1665
      Width           =   435
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   22
      Left            =   915
      TabIndex        =   107
      Top             =   1095
      Width           =   195
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cr A/c"
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
      Height          =   270
      Index           =   38
      Left            =   180
      TabIndex        =   106
      Top             =   1095
      Width           =   480
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doc. Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   1
      Left            =   8580
      TabIndex        =   105
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spares Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   11
      Left            =   180
      TabIndex        =   104
      Top             =   5325
      Width           =   1320
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item-wise Disc Total"
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
      Height          =   255
      Index           =   19
      Left            =   180
      TabIndex        =   103
      Top             =   4785
      Width           =   1680
   End
   Begin VB.Label Lbl 
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   10
      Left            =   180
      TabIndex        =   102
      Top             =   810
      Width           =   690
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MRP Item's Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   33
      Left            =   180
      TabIndex        =   101
      Top             =   5055
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   600
      Left            =   4950
      Shape           =   4  'Rounded Rectangle
      Top             =   6495
      Width           =   2055
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET REFUNDABLE AMT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   30
      Left            =   5040
      TabIndex        =   89
      Top             =   6510
      Width           =   1920
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAXABLE TOTAL"
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
      Index           =   37
      Left            =   7575
      TabIndex        =   100
      Top             =   4785
      Width           =   1365
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   33
      Left            =   9105
      TabIndex        =   99
      Top             =   4785
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET SPARE/LUB AMT :"
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
      Index           =   32
      Left            =   7575
      TabIndex        =   98
      Top             =   6945
      Width           =   1860
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   30
      Left            =   1965
      TabIndex        =   97
      Top             =   5055
      Width           =   180
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   7
      Left            =   7275
      TabIndex        =   96
      Top             =   4530
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11085
      TabIndex        =   95
      Top             =   4545
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   9105
      TabIndex        =   94
      Top             =   4530
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   25
      Left            =   9630
      TabIndex        =   93
      Top             =   4530
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   28
      Left            =   8850
      TabIndex        =   92
      Top             =   4530
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   27
      Left            =   10905
      TabIndex        =   91
      Top             =   4530
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   4
      Left            =   915
      TabIndex        =   90
      Top             =   810
      Width           =   195
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   20
      Left            =   6900
      TabIndex        =   88
      Top             =   2235
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Index           =   28
      Left            =   6045
      TabIndex        =   87
      Top             =   2258
      Width           =   765
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   19
      Left            =   1560
      TabIndex        =   86
      Top             =   1380
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference Doc."
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
      Height          =   270
      Index           =   27
      Left            =   180
      TabIndex        =   85
      Top             =   1380
      Width           =   1275
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   18
      Left            =   9105
      TabIndex        =   84
      Top             =   6675
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Round Off"
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
      Height          =   255
      Index           =   26
      Left            =   7575
      TabIndex        =   83
      Top             =   6675
      Width           =   1365
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOT (TB+TP)"
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
      Height          =   255
      Index           =   25
      Left            =   7575
      TabIndex        =   82
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   6135
      Width           =   1065
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   16
      Left            =   9105
      TabIndex        =   81
      Top             =   5865
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total (B)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   24
      Left            =   7575
      TabIndex        =   80
      Top             =   5865
      Width           =   1080
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   15
      Left            =   9795
      TabIndex        =   79
      Top             =   7590
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surcharge on Tax"
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
      Height          =   255
      Index           =   23
      Left            =   8265
      TabIndex        =   78
      Top             =   7590
      Width           =   1440
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   14
      Left            =   9105
      TabIndex        =   77
      Top             =   5055
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax"
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
      Height          =   255
      Index           =   22
      Left            =   7575
      TabIndex        =   76
      Top             =   5055
      Width           =   810
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   13
      Left            =   1965
      TabIndex        =   75
      Top             =   6675
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transportation"
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
      Height          =   255
      Index           =   21
      Left            =   180
      TabIndex        =   74
      Top             =   6675
      Width           =   1200
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   12
      Left            =   1965
      TabIndex        =   73
      Top             =   6405
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General Surcharge"
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
      Height          =   255
      Index           =   20
      Left            =   180
      TabIndex        =   72
      Top             =   6405
      Width           =   1560
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   11
      Left            =   1965
      TabIndex        =   71
      Top             =   4785
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   10
      Left            =   9105
      TabIndex        =   70
      Top             =   5595
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Misc. Charges"
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
      Height          =   255
      Index           =   18
      Left            =   7575
      TabIndex        =   69
      Top             =   5595
      Width           =   1185
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   8
      Left            =   1965
      TabIndex        =   68
      Top             =   6135
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total (A)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   16
      Left            =   180
      TabIndex        =   67
      Top             =   6135
      Width           =   1080
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   7
      Left            =   1965
      TabIndex        =   66
      Top             =   5865
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Height          =   255
      Index           =   15
      Left            =   180
      TabIndex        =   65
      Top             =   5865
      Width           =   735
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   6
      Left            =   1965
      TabIndex        =   64
      Top             =   5595
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oil Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   14
      Left            =   180
      TabIndex        =   63
      Top             =   5595
      Width           =   930
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Paid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   13
      Left            =   5565
      TabIndex        =   62
      Top             =   4530
      Width           =   765
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   12
      Left            =   3525
      TabIndex        =   61
      Top             =   4530
      Width           =   675
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   5
      Left            =   1965
      TabIndex        =   60
      Top             =   5325
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   2
      Left            =   6285
      TabIndex        =   59
      Top             =   525
      Width           =   195
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Case No"
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
      Height          =   270
      Index           =   8
      Left            =   5445
      TabIndex        =   58
      Top             =   525
      Width           =   735
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person"
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
      Height          =   270
      Index           =   7
      Left            =   180
      TabIndex        =   57
      Top             =   2235
      Width           =   1125
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   1
      Left            =   1350
      TabIndex        =   56
      Top             =   2235
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
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   90
      Left            =   9450
      TabIndex        =   55
      Top             =   1080
      Width           =   45
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
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   92
      Left            =   9450
      TabIndex        =   54
      Top             =   1650
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   93
      Left            =   9450
      TabIndex        =   53
      Top             =   1365
      Width           =   180
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Height          =   270
      Index           =   0
      Left            =   8580
      TabIndex        =   52
      Top             =   1365
      Width           =   390
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   2
      Left            =   8580
      TabIndex        =   51
      Top             =   1650
      Width           =   810
   End
End
Attribute VB_Name = "frmSaleRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMRevDisTBPer As Double, mMRevDisTPPer As Double
Dim mTBDisAmtMRP As Double, mTPDisAmtMRP As Double
Dim mMRPTax As Double, mMRPTaxSur As Double, mMRPTOT As Double, mMRPReSales As Double
Dim mMRPLubeTB As Double, mMRPLubeTP  As Double


Dim mCheckNegetiveStockSiteWise As Boolean
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim RsParty As ADODB.Recordset
Dim RsCrAc As ADODB.Recordset
Dim rsForm As ADODB.Recordset
Dim rsForm31 As ADODB.Recordset
Dim RsBaseDoc As ADODB.Recordset
Dim RsPerson As ADODB.Recordset
Dim RsGodown As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim rsCtrlAc As ADODB.Recordset

Dim mVType As String, mVPrefix As String
Dim mSearchCode As String
Dim mPartyType As Byte         'Used to Detect Part Rate in GetRate Function
Dim BaseDocType$

Private Const SalChalType As String = "SYSC"
Private Const TrfChalType As String = "SYSCT"

Private Const RetCashSalVType As String = "SXSRC"
Private Const RetCrSalVType As String = "SXSRR"
Private Const RetTrfIssVType As String = "SXSRT"

Private Const CashSalVType As String = "SYSIC"
Private Const CrSalVType As String = "SYSIR"
Private Const TrfIssVType As String = "SYSCT"

'grid color scheme
Private Const CellBackColLeave As String = &HE3FAFD
Private Const GridBackColorBkg As String = &HB9D8EE ' me.backColor=&HB9D8EE
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

' Under observation
Dim VoucherEditFlag As Boolean                  ' Used for whether we can edit voucher no or not
' End Under observation
Dim ListArray As Variant
Dim mListItem As ListItem
Dim OldBaseDoc As String                        ' Used For Base Doc.
Dim FillDetailFlag As Byte

Private Const DocID As Byte = 0                 ' Doc.ID
Private Const DocType As Byte = 1               ' Document Type
Private Const VDate As Byte = 2                 ' Date
Private Const SerialNo As Byte = 3              ' Serial No.
Private Const Party As Byte = 4                 ' Party Name
Private Const Address1 As Byte = 5              ' Address1
Private Const CrAc As Byte = 6                  ' Cr A/c
Private Const BaseDoc As Byte = 7               ' Base Document No  (Reference Doc)
Private Const FormName As Byte = 8              ' Form Name
Private Const Form31Name As Byte = 9            ' Form 31 Name
Private Const Form31No As Byte = 10             ' Form 31 No
Private Const SPerson As Byte = 11              ' Sales Person
Private Const CaseNo As Byte = 12               ' Case No.
Private Const CaseMark As Byte = 13             ' Case Mark
Private Const LC As Byte = 14                   ' Dispatch Type(Local/Central)
Private Const TaxDet As Byte = 15               ' Print Tax Detail(Y/N)
Private Const AcPost As Byte = 16               ' A/c Posting (Y/N)
Private Const Remark As Byte = 17               ' Remark

Private Const IWDiscTotTB As Byte = 19          ' Item-wise Disc Total Taxable
Private Const IWDiscTotTP As Byte = 20          ' Item-wise Disc Total Taxpaid
Private Const MRPAmtTB As Byte = 21         ' MRP Item's Amount Taxable
Private Const MRPAmtTP As Byte = 22         ' MRP Item's Amount Taxpaid
Private Const SprAmtTB As Byte = 23             ' Spares Amount Taxable
Private Const SprAmtTP As Byte = 24             ' Spares Amount Taxpaid
Private Const OilAmtTB As Byte = 25             ' Oil Amount Taxable
Private Const OilAmtTP As Byte = 26             ' Oil Amount Taxpaid
Private Const DiscPerTB As Byte = 27            '
Private Const DiscAmtTB As Byte = 28            '
Private Const DiscPerTP As Byte = 29            '
Private Const DiscAmtTP As Byte = 30            '
Private Const STotATB As Byte = 31              '
Private Const STotATP As Byte = 32              '
Private Const Addition As Byte = 33            ' Withdrawn
Private Const PackCrg As Byte = 34              '
Private Const GenSurPer As Byte = 35           '
Private Const GenSurAmt As Byte = 36           '
Private Const TransAmt As Byte = 37             '
Private Const TaxableTot As Byte = 38           '
Private Const STaxPer As Byte = 39              '
Private Const STaxAmt As Byte = 40              '
Private Const TaxSurPer As Byte = 41            '
Private Const TaxSurAmt As Byte = 42            '
Private Const STotB As Byte = 43                '
Private Const TurnOverPer As Byte = 44          '
Private Const TurnOverAmt As Byte = 45          '
Private Const SROff As Byte = 46                '
Private Const NetSprAmt As Byte = 47            '
Private Const NetAmt As Byte = 48               '
Private Const ReSalTaxPer As Byte = 49          '
Private Const ReSalTaxAmt As Byte = 50          '
Private Const SatAmt As Byte = 51          '

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_PNo As Byte = 1               ' Part No
Private Const Col_ChalNoCode As Byte = 2       ' Sale Challan No Code
Private Const Col_ChalNo As Byte = 3           ' Sale Challan No Name
Private Const Col_ChalSrNo As Byte = 4         ' Sale Challan Serial No
Private Const Col_SONoCode As Byte = 5          ' Sale Order No Code
Private Const Col_SONo As Byte = 6              ' Sale Order No Name
Private Const Col_SOSrNo As Byte = 7            ' Sale Order Serial No
Private Const Col_Unit As Byte = 8              ' Unit
Private Const Col_MRP As Byte = 9               ' MRP Yes/No
Private Const Col_Taxable As Byte = 10          ' Taxable Yes/No
Private Const Col_Qty As Byte = 11              ' Qty
Private Const Col_Rate As Byte = 12             ' Rate
Private Const Col_MRPRate As Byte = 13          ' MRP Rate
Private Const Col_Amt As Byte = 14              ' Amt
Private Const Col_DiscPer As Byte = 15          ' Disc. %
Private Const Col_DiscAmt As Byte = 16          ' Disc. Amt.
Private Const Col_TaxPer As Byte = 17           ' Tax Per.
Private Const Col_TaxAmt1 As Byte = 18
Private Const Col_SatPer As Byte = 19           ' Tax Per.
Private Const Col_SatAmt1 As Byte = 20

Private Const Col_ItemVal As Byte = 21          ' Item Value
Private Const Col_GodownCode As Byte = 22       ' Godown Code
Private Const Col_Godown As Byte = 23           ' Godown
Private Const Col_PartSrlNo As Byte = 24        ' Part Serial No
Private Const Col_PName As Byte = 25            ' Part Name
Private Const Col_LName As Byte = 26            ' Local Name
Private Const Col_MRPStkTP As Byte = 27          ' MRP TP Qty 'Current Stock Qty
Private Const Col_MRPStkTB As Byte = 28          ' MRP TB Qty
Private Const Col_TBStk As Byte = 29             ' Taxbale Qty
Private Const Col_TPStk As Byte = 30             ' Tax Paid Qty
Private Const Col_TBRate As Byte = 31           ' Taxbale Rate
Private Const Col_TPRate As Byte = 32           ' Tax Paid Rate
Private Const Col_Bin As Byte = 33              ' Bin
Private Const Col_LastRate As Byte = 34         ' Last Purchase Rate
Private Const Col_HPRate As Byte = 35           ' High Purchase Rate
Private Const Col_LPRate As Byte = 36           ' Low Purchase Rate
Private Const Col_PartGrade As Byte = 37        ' Part Grade (Used for Oil Item)
Private Const Col_EffectDate As Byte = 38       ' MRP Effective Date/TB Effective Date
Private Const Col_IssQty As Byte = 39

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String
Dim mSatYn As Boolean
Dim rsTaxPer As ADODB.Recordset

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To txt.Count - 1
        If I = DocID Or I = IWDiscTotTB Or I = IWDiscTotTP Or I = MRPAmtTB Or I = 18 _
            Or I = MRPAmtTP Or I = SprAmtTB Or I = OilAmtTB Or I = OilAmtTP Or I = SprAmtTP Or I = STotATB _
            Or I = STotATP Or I = TaxableTot Or I = STotB Or I = SROff Or I = NetSprAmt Or I = NetAmt Then
        Else
            txt(I).Enabled = Enb
        End If
    Next
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("SearchCode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select S.DocID As SearchCode From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' and S.V_Type In ('" & RetCashSalVType & "','" & RetCrSalVType & "','" & RetTrfIssVType & "') And S.DocID = '" & MyValue & "' " _
            & "Order by S.V_Date desc,S.DocID Desc")
    End If
    MoveRec
    BUTTONS True, Me, Master, 0
Exit Sub
ELoop:
    CheckError
End Sub

'* Used for clear all text boxes used in the form
Private Sub BlankText()
Dim I As Integer
    For I = 0 To txt.Count - 1
        If I = 18 Then
        Else
            txt(I).TEXT = ""
            If I <> VDate Then
                txt(I).Tag = ""
            End If
        End If
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

'* Used for intialize grid columns
Private Sub Grid_Ini()
Dim MeWidth As Long
    With FGrid
        .left = Me.left '+ 60
        .width = Me.width - 90
        .top = 2550
        .BackColor = CellBackColLeave
        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 40 '34

        .TextMatrix(0, Col_SrNo) = "S.No"
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 450

        .TextMatrix(0, Col_PNo) = "Part No."
        .ColAlignment(Col_PNo) = flexAlignLeftCenter
        .ColWidth(Col_PNo) = 1500

        .TextMatrix(0, Col_ChalNoCode) = "Challan No.Code"
        .ColAlignment(Col_ChalNoCode) = flexAlignLeftCenter
        .ColWidth(Col_ChalNoCode) = 0

        .TextMatrix(0, Col_ChalNo) = "Challan No."
        .ColAlignment(Col_ChalNo) = flexAlignLeftCenter
        .ColWidth(Col_ChalNo) = 0
        
        .TextMatrix(0, Col_ChalSrNo) = "Challan Serial No."
        .ColAlignment(Col_ChalSrNo) = flexAlignLeftCenter
        .ColWidth(Col_ChalSrNo) = 0

        .TextMatrix(0, Col_SONoCode) = "SO No. Code"
        .ColAlignment(Col_SONoCode) = flexAlignLeftCenter
        .ColWidth(Col_SONoCode) = 0

        .TextMatrix(0, Col_SONo) = "SO No."
        .ColAlignment(Col_SONo) = flexAlignLeftCenter
        .ColWidth(Col_SONo) = 0

        .TextMatrix(0, Col_SOSrNo) = "SO Serial No."
        .ColAlignment(Col_SOSrNo) = flexAlignLeftCenter
        .ColWidth(Col_SOSrNo) = 0

        .TextMatrix(0, Col_Unit) = "Unit"
        .ColAlignment(Col_Unit) = flexAlignLeftCenter
        .ColWidth(Col_Unit) = 550

        .TextMatrix(0, Col_MRP) = "MRP"
        .ColAlignment(Col_MRP) = flexAlignLeftCenter
        .ColWidth(Col_MRP) = 450

        .TextMatrix(0, Col_Taxable) = "Tax"
        .ColAlignment(Col_Taxable) = flexAlignLeftCenter
        .ColWidth(Col_Taxable) = 420

        .TextMatrix(0, Col_Qty) = "Qty"
        .ColAlignmentFixed(Col_Qty) = flexAlignRightCenter
        .ColWidth(Col_Qty) = 960

        .TextMatrix(0, Col_Rate) = "Rate"
        .ColAlignmentFixed(Col_Rate) = flexAlignRightCenter
        .ColWidth(Col_Rate) = 870

        .TextMatrix(0, Col_MRPRate) = "MRP Rate"
        .ColAlignmentFixed(Col_MRPRate) = flexAlignRightCenter
        .ColWidth(Col_MRPRate) = 870

        .TextMatrix(0, Col_Amt) = "Amount"
        .ColAlignmentFixed(Col_Amt) = flexAlignRightCenter
        .ColWidth(Col_Amt) = 1065

        .TextMatrix(0, Col_DiscPer) = "Disc%"
        .ColAlignmentFixed(Col_DiscPer) = flexAlignRightCenter
        .ColWidth(Col_DiscPer) = 555

        .TextMatrix(0, Col_DiscAmt) = "Disc.Amt"
        .ColAlignmentFixed(Col_DiscAmt) = flexAlignRightCenter
        .ColWidth(Col_DiscAmt) = 840

        If PubVATYN = 1 Then
            .TextMatrix(0, Col_TaxPer) = "TaxPer"
            .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
            .ColWidth(Col_TaxPer) = 840
            
            .TextMatrix(0, Col_TaxAmt1) = "TaxAmt"
            .ColAlignmentFixed(Col_TaxAmt1) = flexAlignRightCenter
            .ColWidth(Col_TaxAmt1) = 840
        
            .TextMatrix(0, Col_SatPer) = "SatPer"
            .ColAlignmentFixed(Col_SatPer) = flexAlignRightCenter
            .ColWidth(Col_SatPer) = 840
            
            .TextMatrix(0, Col_SatAmt1) = "SatAmt"
            .ColAlignmentFixed(Col_SatAmt1) = flexAlignRightCenter
            .ColWidth(Col_SatAmt1) = 840
        Else
            .TextMatrix(0, Col_TaxPer) = ""
            .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
            .ColWidth(Col_TaxPer) = 0
            
            .TextMatrix(0, Col_TaxAmt1) = ""
            .ColAlignmentFixed(Col_TaxAmt1) = flexAlignRightCenter
            .ColWidth(Col_TaxAmt1) = 0
            
            .ColWidth(Col_SatPer) = 0
            .ColWidth(Col_SatAmt1) = 0
        End If
        
        .TextMatrix(0, Col_ItemVal) = "Item Value"
        .ColAlignmentFixed(Col_ItemVal) = flexAlignRightCenter
        .ColWidth(Col_ItemVal) = 1095

        .TextMatrix(0, Col_GodownCode) = "Godown Code"
        .ColAlignment(Col_GodownCode) = flexAlignLeftCenter
        .ColWidth(Col_GodownCode) = 0

        .TextMatrix(0, Col_Godown) = "Godown"
        .ColAlignment(Col_Godown) = flexAlignLeftCenter
        .ColWidth(Col_Godown) = 1200
        
        .TextMatrix(0, Col_PartSrlNo) = "Part SrlNo"
        .ColAlignmentFixed(Col_PartSrlNo) = flexAlignLeftCenter
        .ColAlignment(Col_PartSrlNo) = flexAlignLeftCenter
        .ColWidth(Col_PartSrlNo) = 1200

        .TextMatrix(0, Col_PName) = "Part Name"
        .ColAlignment(Col_PName) = flexAlignLeftCenter
        .ColWidth(Col_PName) = 2500
 
        .TextMatrix(0, Col_LName) = "Local Name"
        .ColAlignment(Col_LName) = flexAlignLeftCenter
        .ColWidth(Col_LName) = 2000
        
        .TextMatrix(0, Col_MRPStkTP) = "MRP Stock TP"
        .ColAlignmentFixed(Col_MRPStkTP) = flexAlignRightCenter
        .ColWidth(Col_MRPStkTP) = 0

        .TextMatrix(0, Col_MRPStkTB) = "MRP Stk TB"
        .ColAlignmentFixed(Col_MRPStkTB) = flexAlignRightCenter
        .ColWidth(Col_MRPStkTB) = 0

        .TextMatrix(0, Col_TBStk) = "Taxable Qty"
        .ColAlignmentFixed(Col_TBStk) = flexAlignRightCenter
        .ColWidth(Col_TBStk) = 0

        .TextMatrix(0, Col_TPStk) = "Tax Paid Qty"
        .ColAlignmentFixed(Col_TPStk) = flexAlignRightCenter
        .ColWidth(Col_TPStk) = 0

        .TextMatrix(0, Col_TBRate) = "Taxbale Rate"
        .ColAlignmentFixed(Col_TBRate) = flexAlignRightCenter
        .ColWidth(Col_TBRate) = 0

        .TextMatrix(0, Col_TPRate) = "Tax Paid Rate"
        .ColAlignmentFixed(Col_TPRate) = flexAlignRightCenter
        .ColWidth(Col_TPRate) = 0

        .TextMatrix(0, Col_Bin) = "Bin"
        .ColAlignmentFixed(Col_Bin) = flexAlignLeftCenter
        .ColWidth(Col_Bin) = 600

        .TextMatrix(0, Col_LastRate) = "Last Purchase Rate"
        .ColAlignmentFixed(Col_LastRate) = flexAlignRightCenter
        .ColWidth(Col_LastRate) = 0

        .TextMatrix(0, Col_HPRate) = "High Purchase Rate"
        .ColAlignmentFixed(Col_HPRate) = flexAlignRightCenter
        .ColWidth(Col_HPRate) = 0

        .TextMatrix(0, Col_LPRate) = "Low Purchase Rate"
        .ColAlignmentFixed(Col_LPRate) = flexAlignRightCenter
        .ColWidth(Col_LPRate) = 0

        .TextMatrix(0, Col_PartGrade) = "Part Grade"
        .ColWidth(Col_PartGrade) = 0

        .TextMatrix(0, Col_EffectDate) = "Rate Effective Date"
        .ColWidth(Col_EffectDate) = 0
        
        .TextMatrix(0, Col_IssQty) = ""
        .ColWidth(Col_IssQty) = 0
        
    End With
    MeWidth = Me.width
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    
    DGPart.width = FGrid.width: DGPart.left = FGrid.left: DGPart.top = FGrid.top + FGrid.height: DGPart.height = Me.height - (DGPart.top + mBotScale)
    DGSONo.left = (MeWidth - DGSONo.width) / 2: DGSONo.top = DGPart.top: DGSONo.height = Me.height - (FGrid.top + FGrid.height + mBotScale)
    DGGodown.left = MeWidth - (DGGodown.width + mRtScale): DGGodown.top = DGPart.top: DGGodown.height = Me.height - (FGrid.top + FGrid.height + mBotScale)
    FrmDetail.width = 6285: FrmDetail.left = 5595: FrmDetail.top = 405: FrmDetail.height = 2130
    DGParty.left = MeWidth - (DGParty.width + mRtScale): DGParty.top = mTopScale
    DGCrAc.left = MeWidth - (DGCrAc.width + mRtScale): DGCrAc.top = mTopScale
    DGForm.left = MeWidth - (DGForm.width + mRtScale): DGForm.top = mTopScale
    DGForm31.left = MeWidth - (DGForm31.width + mRtScale): DGForm31.top = mTopScale
    FrmPrn.left = (MeWidth - FrmPrn.width) / 2: FrmPrn.top = (Me.height - FrmPrn.height) / 2
    DGBaseDoc.left = (MeWidth - DGBaseDoc.width): DGBaseDoc.top = mTopScale
    DGPerson.left = MeWidth - (DGPerson.width + mRtScale): DGPerson.top = mTopScale
    With DGPart
        .Columns(6).width = 2564.788
        .Columns(5).width = 1005.165
        .Columns(4).width = 1005.165
        .Columns(3).width = 1005.165
        .Columns(2).width = 494.9292
        .Columns(1).width = 3225.26
        .Columns(0).width = 1950.236
    End With

End Sub

Private Sub Grid_Hide()
    If ListView.Visible = True Then ListView.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If lblGroup.Visible = True Then lblGroup.Visible = False
    If DGCrAc.Visible = True Then DGCrAc.Visible = False
    If DGForm.Visible = True Then DGForm.Visible = False
    If DGForm31.Visible = True Then DGForm31.Visible = False
    If DGBaseDoc.Visible = True Then DGBaseDoc.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGPart.Visible = True Then DGPart.Visible = False
    If DGSONo.Visible = True Then DGSONo.Visible = False
    If DGGodown.Visible = True Then DGGodown.Visible = False
    If DGPerson.Visible = True Then DGPerson.Visible = False
End Sub

Private Sub cmdPost_Click()
Dim I As Integer
    If Master.RecordCount > 0 Then Master.MoveFirst
    Do Until Master.EOF
        Call MoveRec
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        Call TopCtrl1_eEdit
        FGrid.Col = Col_TaxPer
        For I = 1 To FGrid.Rows - 1
            If Val(FGrid.TextMatrix(I, Col_TaxAmt1)) = 0 Or Val(txt(STaxAmt).TEXT) = 0 Then
                FGrid.Row = I
                FGrid_Click
                FGrid_DblClick
                Call TxtGrid_GotFocus(0)
                Call TxtGrid_Validate(0, False)
             End If
        Next
        Call Txt_Validate(STaxAmt, False)
        Call TopCtrl1_eSave
MyNextRecord:
        Master.MoveNext
    Loop

End Sub

Private Sub DGParty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGParty.Row >= 0 Then
    lblGroup.TEXT = G_FaCn.Execute("Select AcGroup.GroupName from (AcGroup Left Join SubGroup on SubGroup.GroupCode=AcGroup.GroupCode) where SubGroup.SubCode='" & RsParty!Code & "'").Fields(0).Value
    lblGroup.Refresh
End If
End Sub
Private Sub MoveRec()
Dim Master1 As ADODB.Recordset
Dim Rst As ADODB.Recordset, I As Integer
Dim mItemDiscTotTB As Double, mItemDiscTotTP As Double
On Error GoTo ELoop
    mMRevDisTBPer = 0
    mMRevDisTPPer = 0
    mTBDisAmtMRP = 0
    mTPDisAmtMRP = 0
    mMRPTax = 0
    mMRPTaxSur = 0
    mMRPTOT = 0
    mMRPReSales = 0
    mMRPLubeTB = 0
    mMRPLubeTP = 0
    FrmDetail.Visible = False
    If Master.RecordCount > 0 Then
        If InStr(Me.TopCtrl1.Tag, "E") <> 0 Then Me.TopCtrl1.tEdit = True
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "Select S.*,SubGroup.Name As PartyName,SubGroup.Party_Type,SubGroupCr.Name As CreditAcName,TaxForms.Form_Desc As FormName,TaxForms31.Form_Desc As Form31Name,Emp.Emp_Name " _
            & "From (((((SP_Sale S Left Join SubGroup on S.Party_Code=SubGroup.SubCode) " _
            & "Left Join SubGroup SubGroupCr on S.CrAc=SubGroupCr.SubCode) " _
            & "Left Join TaxForms on S.Form_Code=TaxForms.Form_Code) " _
            & "Left Join TaxForms TaxForms31 on S.RoadPermit_FormCode=TaxForms31.Form_Code) " _
            & "Left Join Emp_Mast Emp on S.Rep_Code=Emp.Emp_Code) " _
            & "Where S.DocID='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
        If Master1!CancelYN = 1 Then
            TopCtrl1.tEdit = False
            LblCancel.Visible = True
        Else
            LblCancel.Visible = False
        End If
        txt(DocID).TEXT = Master1!DocID
        mSearchCode = txt(DocID)
        LblDiv.CAPTION = "Division : " & left(Master1!DocID, 1)
        LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        mVType = Master1!V_Type
        If mVType = RetCashSalVType Then
            txt(DocType).TEXT = "Sale Return Cash"
            mPartyType = 0
        ElseIf mVType = RetCrSalVType Then
            txt(DocType).TEXT = "Sale Return Credit"
            mPartyType = Master1!Party_Type
        ElseIf mVType = RetTrfIssVType Then
            txt(DocType).TEXT = "Transfer Issue Return"
            mPartyType = Master1!Party_Type
        End If
        
        If PubBackEnd = "A" Then
            mSatYn = IIf(VNull(Master1!SAT_YN) = 1, True, False)
        Else
            mSatYn = IIf(VNull(Master1!SAT_YN) = True, True, False)
        End If
        DispTextVat
        
        txt(VDate).TEXT = Master1!V_Date
        LblVPrefix.CAPTION = mID(Master1!DocID, 9, 5)
        txt(SerialNo).TEXT = Master1!V_NO
        txt(Party).Tag = Master1!Party_code
        
        If mVType = RetCrSalVType Or mVType = RetTrfIssVType Then
            txt(Party).TEXT = IIf(IsNull(Master1!PartyName), "", Master1!PartyName)
        ElseIf mVType = RetCashSalVType Then
            txt(Party).TEXT = Master1!Party_Name
        End If
        
        txt(Address1).TEXT = XNull(Master1!Address)
        txt(CrAc).Tag = XNull(Master1!CrAc)
        txt(CrAc).TEXT = IIf(IsNull(Master1!CreditAcName), "", Master1!CreditAcName)
        If Master1!Invoice_DocID = "" Then
            txt(BaseDoc).Tag = ""
            txt(BaseDoc).TEXT = ""
            OldBaseDoc = ""
        Else
            txt(BaseDoc).Tag = XNull(Master1!Invoice_DocID)
            txt(BaseDoc).TEXT = Trim(mID(XNull(Master1!Invoice_DocID), 9, 5)) + CStr(Trim(Right(XNull(Master1!Invoice_DocID), 8)))
            OldBaseDoc = XNull(Master1!Invoice_DocID)
        End If
        txt(FormName).Tag = Master1!Form_Code
        txt(FormName).TEXT = IIf(IsNull(Master1!FormName), "", Master1!FormName)
        txt(Form31Name).Tag = XNull(Master1!RoadPermit_FormCode)
        txt(Form31Name).TEXT = IIf(IsNull(Master1!Form31Name), "", Master1!Form31Name)
        txt(Form31No).TEXT = XNull(Master1!RoadPermit_No)
        txt(SPerson).Tag = XNull(Master1!REP_CODE)
        txt(SPerson).TEXT = IIf(IsNull(Master1!Emp_Name), "", Master1!Emp_Name)
        txt(Remark).TEXT = XNull(Master1!Remarks)
        If Master1!L_C = "L" Then
            txt(LC).TEXT = "Local"
        ElseIf Master1!L_C = "C" Then
            txt(LC).TEXT = "Central"
        End If
        txt(CaseNo).TEXT = Master1!Case_No
        txt(CaseMark).TEXT = XNull(Master1!Case_Mark)
        txt(TaxDet).TEXT = IIf(Master1!Det_Tax = 0, "No", "Yes")
        txt(MRPAmtTB).TEXT = Format(Master1!SprAmt_MRP_TB + Master1!OilAmt_MRP_TB, "0.00")
        txt(MRPAmtTP).TEXT = Format(Master1!SprAmt_MRP_TP + Master1!OilAmt_MRP_TP, "0.00")
        mMRPLubeTB = Master1!OilAmt_MRP_TB
        mMRPLubeTP = Master1!OilAmt_MRP_TP
        txt(SprAmtTB).TEXT = Format(Master1!SprAmt_TB, "0.00")
        txt(SprAmtTP).TEXT = Format(Master1!SprAmt_TP, "0.00")
        txt(OilAmtTB).TEXT = Format(Master1!OilAmt_TB, "0.00")
        txt(OilAmtTP).TEXT = Format(Master1!OilAmt_TP, "0.00")
        txt(DiscPerTB).TEXT = Format(Master1!D_Per_TB, "0.00")
        txt(DiscAmtTB).TEXT = Format(Master1!D_Amt_TB, "0.00")
        txt(DiscPerTP).TEXT = Format(Master1!D_Per_TP, "0.00")
        txt(DiscAmtTP).TEXT = Format(Master1!D_Amt_TP, "0.00")
        txt(STotATB).TEXT = Format((Master1!SprAmt_MRP_TB + Master1!OilAmt_MRP_TB + Master1!SprAmt_TB + Master1!OilAmt_TB) - Master1!D_Amt_TB, "0.00")
        txt(STotATP).TEXT = Format((Master1!SprAmt_MRP_TP + Master1!OilAmt_MRP_TP + Master1!SprAmt_TP + Master1!OilAmt_TP) - Master1!D_Amt_TP, "0.00")
'        Txt(Addition).Text = Format(Master1!Addition, "0.00")
        txt(PackCrg).TEXT = Format(Master1!Packing, "0.00")
        txt(GenSurPer).TEXT = Format(Master1!Gen_Sur_Per, "0.00")
        txt(GenSurAmt).TEXT = Format(Master1!Gen_Sur_Amt, "0.00")
        txt(TransAmt).TEXT = Format(Master1!Trans_Amt, "0.00")
'        Txt(TaxableTot) = Format(Val(Txt(STotATB)) + Val(Txt(Addition)) + Val(Txt(PackCrg)) + Val(Txt(GenSurAmt)) + Val(Txt(TransAmt)), "0.00")
        txt(TaxableTot) = Format(Val(txt(STotATB)) + Val(txt(PackCrg)) + Val(txt(GenSurAmt)) + Val(txt(TransAmt)), "0.00")
        txt(STaxPer).TEXT = Format(Master1!Tax_Per, "0.00")
        txt(STaxAmt).TEXT = Format(Master1!Tax_Amt, "0.00")
        txt(SatAmt).TEXT = Format(Master1!SatAmt, "0.00")
        txt(TaxSurPer).TEXT = Format(Master1!Tax_Sur_Per, "0.00")
        txt(TaxSurAmt).TEXT = Format(Master1!Tax_Sur_Amt, "0.00")
        txt(STotB) = Format(Val(txt(TaxableTot)) + Val(txt(STaxAmt)) + Val(txt(TaxSurAmt)), "0.00")
        txt(TurnOverPer).TEXT = Format(Master1!TOT_Per, "0.00")
        txt(TurnOverAmt).TEXT = Format(Master1!Tot_Amt, "0.00")
        txt(ReSalTaxPer).TEXT = Format(Master1!ReSalTax_Per, "0.00")
        txt(ReSalTaxAmt).TEXT = Format(Master1!ReSalTax_Amt, "0.00")
        txt(SROff).TEXT = Format(Master1!Rounded, "0.00")
        txt(NetSprAmt) = Format(Val(txt(STotB)) + Val(txt(STotATP)) + Val(txt(TurnOverAmt)) + Val(txt(SROff)), "0.00")
        txt(NetAmt).TEXT = Format(Master1!Total_Amt, "0.00")

        FGrid.Rows = 1
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select P.Part_Name ,P.Local_Name ,P.Unit ,P.MRP ,P.MRP_Effect_Dt ,P.TB_SRate ,P.TP_SRate ,P.TB_Effect_Dt ,P.Part_Grade ,P.Cur_MRP_TBStk, P.Cur_MRP_TPStk, P.Cur_TB_Stk ,P.Cur_TP_Stk ,P.Bin_Loca ,P.High_Pur_Rate ,P.Low_Pur_Rate,Godown.God_Name," & cTrim(cMID("SP_Stock.DocID", "9", "5")) & " + " & cCStr(cTrim("Right(SP_Stock.DocID,8)")) & " As ChallIDDisp," & cTrim(cMID("SP_Stock.Order_DocID", "9", "5")) & " +  " & cCStr(cTrim("Right(SP_Stock.Order_DocID,8)")) & " As OrderIDDisp,SP_Stock.* " & _
            "From (SP_Stock Left Join Part P On SP_Stock.Part_No=P.Part_No and P.Div_Code = left(SP_Stock.DocID,1)) " & _
            "Left Join Godown on SP_Stock.Godown=Godown.God_Code " & _
            "Where SP_Stock.DocId='" & Master1!DocID & "'", GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount > 0 Then
            I = 1
            Do Until Rst.EOF
                            '|0 Col_SrNo |1 Col_PNo |2 Col_ChalNoCode |3 Col_ChalNo |4 Col_ChalSrNo |5 Col_SONoCode |6 Col_SONo |7 Col_SOSrNo |8 Col_Unit |9 Col_MRP |10 Col_Taxable |11 Col_Qty |12 Col_Rate |13 Col_MRPRate |14 Col_Amt |15 Col_DiscPer |16 Col_DiscAmt |17 Col_ItemVal |18 Col_GodownCode |19 Col_Godown |20 Col_PName |21 Col_LName |22 Col_MRPStkTP |23 Col_MRPStkTB |24 Col_TBStk |25 Col_TPStk |26 Col_TBRate |27 Col_TPRate |28 Col_Bin |29 Col_LastRate |30 Col_HPRate |31 Col_LPRate |32 Col_PartGrade |33 Col_EffectDate
'                FGrid.AddItem i & Chr(9) & Rst!Part_No & Chr(9) & Rst!DocId & Chr(9) & Rst!ChallIDDisp & Chr(9) & Rst!Srl_No & Chr(9) & Rst!order_docid & Chr(9) & Rst!OrderIDDisp & Chr(9) & Rst!Order_Srl_No & Chr(9) & Rst!Unit & Chr(9) & IIf(Rst!MRP_YN = 1, "Yes", "No") & Chr(9) & IIf(Rst!Tax_YN = 1, "Yes", "No") & Chr(9) & Format(Rst!Qty_iss, "0.000") & Chr(9) & Format(Rst!Rate, "0.00") & Chr(9) & Format(Rst!MRP, "0.00") & Chr(9) & Format((Rst!Qty_iss * Rst!Rate), "0.00") & Chr(9) & Format(Rst!Disc_Per, "0.00") & Chr(9) & Format(Rst!Disc_Amt, "0.00") & Chr(9) & Format(Rst!Net_Amt, "0.00") & Chr(9) & Rst!Godown & Chr(9) & Rst!God_Name & Chr(9) & Rst!Part_Name & Chr(9) & Rst!Local_Name & Chr(9) & Rst!Curstk & Chr(9) & Rst!MRPQty & Chr(9) & Rst!Cur_TB_Stk & Chr(9) & Rst!Cur_TP_Stk & Chr(9) & Rst!TB_SRate & Chr(9) & Rst!TP_SRate & Chr(9) & Rst!Bin_Loca & Chr(9) & " " & Chr(9) & Rst!high_pur_rate & Chr(9) & Rst!low_pur_rate & Chr(9) & Rst!Part_Grade & Chr(9) & _
                        Format(IIf(Rst!MRP_YN = 1, Rst!MRP_Effect_Dt, Rst!TB_Effect_Dt), "dd/MMM/yyyy")
                             '0                  1                     2                        3                      4                         5                         6                            7                     8                                 9                                           10                                      11                                12                                     13                                        14                                              15                                        16                                     17                           18                      19                    20                        21                      22                   23                      24                        25                          26                      27                      28               29                   30                          31                         32                                                              33
                FGrid.AddItem ""
                With FGrid
                    .TextMatrix(I, Col_SrNo) = I
                    .TextMatrix(I, Col_PNo) = Rst!Part_No
                    .TextMatrix(I, Col_ChalNoCode) = Rst!DocID
                    .TextMatrix(I, Col_ChalNo) = IIf(IsNull(Rst!ChallIDDisp), "", Rst!ChallIDDisp)
                    .TextMatrix(I, Col_ChalSrNo) = Rst!Srl_No
                    .TextMatrix(I, Col_SONoCode) = Rst!Order_DocId
                    .TextMatrix(I, Col_SONo) = IIf(IsNull(Rst!OrderIDDisp), "", Rst!OrderIDDisp)
                    .TextMatrix(I, Col_SOSrNo) = Rst!Order_Srl_No
                    .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                    .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Qty) = Format(Rst!Qty_Rec, "0.000")
                    .TextMatrix(I, Col_IssQty) = Format(Rst!Qty_Rec, "0.000")
                    .TextMatrix(I, Col_Rate) = Format(Rst!Rate, "0.00")
                    .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP_Rate, "0.00")           'Format(Rst!MRP, "0.00")
                    If Rst!MRP_YN = 1 Then
                        .TextMatrix(I, Col_Amt) = Format((Rst!Qty_Rec * Rst!MRP_Rate), "0.00")
                    Else
                        .TextMatrix(I, Col_Amt) = Format((Rst!Qty_Rec * Rst!Rate), "0.00")
                    End If
                    .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per, "0.00")
                    .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt, "0.00")
                    If PubVATYN = 1 Then
                        .TextMatrix(I, Col_TaxPer) = Format(Rst!TaxPer, "0.00")
                        .TextMatrix(I, Col_TaxAmt1) = Format(Rst!TaxAmt, "0.00")
                        .TextMatrix(I, Col_SatPer) = Format(Rst!SatPer, "0.00")
                        .TextMatrix(I, Col_SatAmt1) = Format(Rst!SatAmt, "0.00")
                    End If
                    .TextMatrix(I, Col_ItemVal) = Format(Rst!Net_Amt, "0.00")
                    .TextMatrix(I, Col_GodownCode) = Rst!Godown
                    .TextMatrix(I, Col_Godown) = IIf(IsNull(Rst!God_Name), "", Rst!God_Name)
                    .TextMatrix(I, Col_PartSrlNo) = IIf(IsNull(Rst!Part_SrlNo), "", Rst!Part_SrlNo)
                    .TextMatrix(I, Col_PName) = IIf(IsNull(Rst!Part_Name), "", Rst!Part_Name)
                    .TextMatrix(I, Col_LName) = IIf(IsNull(Rst!Local_Name), "", Rst!Local_Name)
                    .TextMatrix(I, Col_MRPStkTP) = IIf(IsNull(Rst!Cur_MRP_TPStk), "", Rst!Cur_MRP_TPStk)
                    .TextMatrix(I, Col_MRPStkTB) = IIf(IsNull(Rst!Cur_MRP_TbStk), "", Rst!Cur_MRP_TbStk)
                    .TextMatrix(I, Col_TBStk) = IIf(IsNull(Rst!Cur_TB_STk), "", Rst!Cur_TB_STk)
                    .TextMatrix(I, Col_TPStk) = IIf(IsNull(Rst!Cur_TP_Stk), "", Rst!Cur_TP_Stk)
                    .TextMatrix(I, Col_TBRate) = IIf(IsNull(Rst!TB_SRate), "", Rst!TB_SRate)
                    .TextMatrix(I, Col_TPRate) = IIf(IsNull(Rst!TP_SRate), "", Rst!TP_SRate)
                    .TextMatrix(I, Col_Bin) = IIf(IsNull(Rst!Bin_Loca), "", Rst!Bin_Loca)
                    .TextMatrix(I, Col_LastRate) = ""
                    .TextMatrix(I, Col_HPRate) = IIf(IsNull(Rst!high_pur_rate), "", Rst!high_pur_rate)
                    .TextMatrix(I, Col_LPRate) = IIf(IsNull(Rst!low_pur_rate), "", Rst!low_pur_rate)
                    .TextMatrix(I, Col_PartGrade) = IIf(IsNull(Rst!Part_Grade), "", Rst!Part_Grade)
                    .TextMatrix(I, Col_EffectDate) = Format(IIf(Rst!MRP_YN = 1, IIf(IsNull(Rst!MRP_Effect_Dt), "", Rst!MRP_Effect_Dt), IIf(IsNull(Rst!TB_Effect_Dt), "", Rst!TB_Effect_Dt)), "dd/MMM/yyyy")
                End With
                If Rst!Tax_YN = 1 Then
                    mItemDiscTotTB = mItemDiscTotTB + Rst!Disc_Amt
                Else
                    mItemDiscTotTP = mItemDiscTotTP + Rst!Disc_Amt
                End If
                Rst.MoveNext
                I = I + 1
            Loop
            txt(IWDiscTotTB).TEXT = Format(mItemDiscTotTB, "0.00")
            txt(IWDiscTotTP).TEXT = Format(mItemDiscTotTP, "0.00")
            FGrid.FixedRows = 1
            CountItem
        Else
            FGrid.AddItem FGrid.Rows
            FGrid.FixedRows = 1
        End If
    Else
        BlankText
    End If
    Grid_Hide
Set Rst = Nothing
Set Master1 = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

' Used For Checking Duplicate Items in the Grid
Private Function ChkDuplicate() As Boolean
Dim I As Integer, X As String, Y As String
Dim Col1 As Byte, Col2 As Byte, Col3 As Byte, Col4 As Byte
    Select Case FGrid.Col
    Case Col_PNo, Col_PName, Col_LName
        Col4 = FGrid.Col
        Col1 = Col_SONo
        Col2 = Col_Taxable
        Col3 = Col_MRP
    Case Col_MRP
        Col1 = Col_PNo
        Col2 = Col_SONo
        Col3 = Col_Taxable
        Col4 = Col_MRP
    Case Col_Taxable
        Col1 = Col_PNo
        Col2 = Col_SONo
        Col4 = Col_Taxable
        Col3 = Col_MRP
    Case Col_SONo
        Col1 = Col_PNo
        Col4 = Col_SONo
        Col2 = Col_Taxable
        Col3 = Col_MRP
    End Select
    X = UCase(CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col1))) + CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col2))) + CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col3))) + CStr(Trim(TxtGrid(0).TEXT)))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col1))) + CStr(Trim(FGrid.TextMatrix(I, Col2))) + CStr(Trim(FGrid.TextMatrix(I, Col3))) + CStr(Trim(FGrid.TextMatrix(I, Col4))))
        If X = Y And Y <> "" Then
            MsgBox "Duplicate Item Not Allowed", vbInformation, "Validation"
            If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
    Select Case FGrid.Col
'        Case Col_SONo
'            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
'            'Old code equivalent to Validate removed by LPS
'            TxtGridValid_SONo

        Case Col_PNo, Col_PName, Col_LName
'            If RsPart.RecordCount = 0 Then TxtGridLeave = False: Exit Function
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            TxtGridValid_PNo
            
        Case Col_Taxable, Col_MRP
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            'Old code equivalent to Validate removed by LPS
            TxtGridValid_TaxMRP
            
        Case Col_Rate, Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
            Amt_Cal
        Case Col_DiscAmt
            If Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) < Val(TxtGrid(0)) Then
                MsgBox "Item-wsie Disc. Amount is greater than Item Value", vbOKOnly, "Item-wise Disc. Checking"
                TxtGridLeave = False: Exit Function
            End If
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
            Amt_Cal
        Case Col_Qty
            If Val(FGrid.TextMatrix(FGrid.Row, Col_IssQty)) > 0 Then
                If Val(TxtGrid(0).TEXT) > Val(FGrid.TextMatrix(FGrid.Row, Col_IssQty)) Then
                    MsgBox "Qty can not be greater than issue qty.", vbInformation
                    TxtGrid(0) = ""
                    Exit Function
                Else
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.000")
                End If
            End If
            Amt_Cal
            If RsGodown.RecordCount > 0 And Trim(FGrid.TextMatrix(FGrid.Row, Col_Godown)) = "" Then
                RsGodown.MoveFirst
                RsGodown.FIND "Code ='" & PubSprCounterGodown & "'"
                FGrid.TextMatrix(FGrid.Row, Col_GodownCode) = RsGodown!Code
                FGrid.TextMatrix(FGrid.Row, Col_Godown) = RsGodown!Name
            End If
            
        Case Col_PartSrlNo
            FGrid.TextMatrix(FGrid.Row, Col_PartSrlNo) = TxtGrid(0)
        
        Case Col_Godown
            TxtGridValid_Godown
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
'Dim i As Integer
'Dim TotItemDisTB As Double, TotItemDisTP As Double
'Dim TotMRPAmtTB As Double, TotMRPAmtTP As Double
'Dim TotSprAmtTB As Double, TotSprAmtTP As Double
'Dim TotOilAmtTB As Double, TotOilAmtTP As Double
'Dim TotMRPItemDisTB As Double, TotMRPItemDisTP As Double
Dim mAmount As Double, TaxAmt As Double, DisAmt As Double, OrdDisAmt1 As Double
 Dim TTaxAmt As Double, mTaxableAmt As Double
    If FillDetailFlag <> 1 Then
        If UCase(left(PubComp_Name, 3)) = "JMK" Then
            If UCase(FGrid.TextMatrix(FGrid.Row, Col_MRP)) = "YES" Then
                FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty))), "0.00")
            Else
                FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty))), "0.00")
            End If
        Else
            If UCase(FGrid.TextMatrix(FGrid.Row, Col_MRP)) = "YES" Then
                FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty))), "0.00")
            Else
                FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty))), "0.00")
            End If
        End If
    
    
    
'        If FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" Then
'            FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty))), "0.00")
'        Else
'            FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty))), "0.00")
'        End If
        FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = Format(((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) * Val(FGrid.TextMatrix(FGrid.Row, Col_DiscPer))) / 100), "0.00")
        FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) - Val(FGrid.TextMatrix(FGrid.Row, Col_DiscAmt))), "0.00")
    End If
    '******************** For Tax in Line File *************************
    If PubVATYN = 1 Then
       If txt(FormName).Tag <> "" Then
            If FGrid.TextMatrix(FGrid.Row, Col_TaxPer) <> "" Then
                mAmount = Val(FGrid.TextMatrix(FGrid.Row, Col_Amt))
                DisAmt = Val(FGrid.TextMatrix(FGrid.Row, Col_DiscAmt))
                If FGrid.TextMatrix(FGrid.Row, Col_MRP) = "Yes" And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" Then
                    If mSatYn Then
                        mTaxableAmt = Format((mAmount - DisAmt) * 100 / (100 + Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) + Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer))), "0.00")
                        FGrid.TextMatrix(FGrid.Row, Col_TaxAmt1) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / 100, "0.00")
                        FGrid.TextMatrix(FGrid.Row, Col_SatAmt1) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer)) / 100, "0.00")
                    Else
                        FGrid.TextMatrix(FGrid.Row, Col_TaxAmt1) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / (100 + Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer))), "0.00")
                        FGrid.TextMatrix(FGrid.Row, Col_SatAmt1) = 0
                    End If
                    FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_ItemVal)) - Val(FGrid.TextMatrix(FGrid.Row, Col_TaxAmt1)) - Val(FGrid.TextMatrix(FGrid.Row, Col_SatAmt1)), "0.00")
                ElseIf FGrid.TextMatrix(FGrid.Row, Col_MRP) = "No" And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" Then
                    FGrid.TextMatrix(FGrid.Row, Col_TaxAmt1) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / 100, "0.00")
                    FGrid.TextMatrix(FGrid.Row, Col_SatAmt1) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer)) / 100, "0.00")
                Else
                    FGrid.TextMatrix(FGrid.Row, Col_TaxAmt1) = ""
                    FGrid.TextMatrix(FGrid.Row, Col_SatAmt1) = ""
                End If
            End If
       End If
       
    End If
    '*******************************************************************
    
    If PubVATYN = 1 Then
       
        MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
            txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
            txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
            txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
            txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
            txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt1, txt(SatAmt)
    Else
        MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
            txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
            txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
            txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
            txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
            txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
    End If
'MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
    Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
    Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
    Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
    Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
    Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
    Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
    Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
End Sub

Private Sub CountItem()
Dim I As Integer, TotItems As Integer, TotQty As Double
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            TotQty = TotQty + Val(FGrid.TextMatrix(I, Col_Qty))
            TotItems = TotItems + 1
        End If
    Next I
    LblIVal.CAPTION = Format(TotItems, "0")
    LblQty.CAPTION = Format(TotQty, "0.000")
End Sub

' Used For Enable/Disable Controls in case of Base Document Option
Private Sub CtrlEnbChallan(Enb As Boolean)
'    Txt(FormName).Enabled = Enb
    txt(Form31Name).Enabled = Enb
    txt(Form31No).Enabled = Enb
    txt(SPerson).Enabled = Enb
    txt(CaseNo).Enabled = Enb
    txt(CaseMark).Enabled = Enb
    txt(LC).Enabled = Enb
End Sub

Private Sub FillItemDet()
Dim Rst As ADODB.Recordset, I As Integer, DocIDName$, TSQL$
    If Trim(txt(BaseDoc).TEXT) <> "" Then
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient

        If txt(DocType) = "Sale Return Credit" Or txt(DocType) = "Sale Return Cash" Then
            DocIDName = "Invoice_DocID"
            'modi lps at Cuttack 17.02.04
            TSQL = " and " & cTrim(cMID("SP_Stock.DocID", "4", "5")) & "='" & SalChalType & "'"
        Else
            DocIDName = "DocID"
        End If
        Rst.Open "Select SS.L_C As LC,SS.Form_Code,SS.RoadPermit_FormCode,SS.RoadPermit_No,SS.Rep_Code,Emp.Emp_Name,SS.Case_No,SS.Case_Mark,P.Part_Name ,P.Local_Name ,P.Unit ,P.MRP ,P.MRP_Effect_Dt,P.TB_SRate ,P.TP_SRate ,P.TB_Effect_Dt ,P.Part_Grade, P.Cur_MRP_TBStk, P.Cur_MRP_TPStk, P.Cur_TB_Stk, P.Cur_TP_Stk, P.Bin_Loca, P.High_Pur_Rate, P.Low_Pur_Rate, Godown.God_Name, " & cTrim(cMID("SP_Stock.DocID", "9", "5")) & " + " & cCStr(cTrim("Right(SP_Stock.DocID,8)")) & " As ChallIDDisp, " & cTrim(cMID("SP_Stock.Order_DocID", "9", "5")) & " + " & cCStr(cTrim("Right(SP_Stock.Order_DocID,8)")) & " As OrderIDDisp,SP_Stock.* " _
            & "From (((SP_Stock Left Join SP_Sale SS on SP_Stock.Invoice_DocId=SS.DocID) " _
            & "Left Join Part P On SP_Stock.Part_No=P.Part_No and P.Div_Code = left(SP_Stock.DocID,1)) " _
            & "Left Join Godown on SP_Stock.Godown=Godown.God_Code) " _
            & "Left Join Emp_Mast Emp on SS.Rep_Code=Emp.Emp_Code " _
            & "Where SP_Stock." & DocIDName & "='" & txt(BaseDoc).Tag & "'" & TSQL, GCn, adOpenStatic, adLockReadOnly
        
        FGrid.Rows = 1
        If Rst.RecordCount > 0 Then
            FillDetailFlag = 1
            If Rst!Form_Code <> "" Then
                txt(FormName).TEXT = GCn.Execute("Select Form_Desc From TaxForms Where Form_Code='" & Rst!Form_Code & "'").Fields(0).Value
            End If
            If Rst!RoadPermit_FormCode <> "" Then
                txt(Form31Name).TEXT = GCn.Execute("Select Form_Desc From TaxForms Where Form_Code='" & Rst!RoadPermit_FormCode & "'").Fields(0).Value
            End If
            txt(Form31No).TEXT = XNull(Rst!RoadPermit_No)
            txt(SPerson).Tag = XNull(Rst!REP_CODE)
            txt(SPerson).TEXT = IIf(IsNull(Rst!Emp_Name), "", Rst!Emp_Name)
            txt(CaseNo).TEXT = XNull(Rst!Case_No)
            txt(CaseMark).TEXT = XNull(Rst!Case_Mark)
            txt(LC).TEXT = IIf(Rst!LC = "L", "Local", "Central")
            I = 1
            Do Until Rst.EOF
                    '|0 Col_SrNo |1 Col_PNo |2 Col_ChalNoCode |3 Col_ChalNo |4 Col_ChalSrNo |5 Col_SONoCode |6 Col_SONo |7 Col_SOSrNo |8 Col_Unit |9 Col_MRP |10 Col_Taxable |11 Col_Qty |12 Col_Rate |13 Col_MRPRate |14 Col_Amt |15 Col_DiscPer |16 Col_DiscAmt |17 Col_ItemVal |18 Col_GodownCode |19 Col_Godown |20 Col_PName |21 Col_LName |22 Col_MRPStkTP |23 Col_MRPStkTB |24 Col_TBStk |25 Col_TPStk |26 Col_TBRate |27 Col_TPRate |28 Col_Bin |29 Col_LastRate |30 Col_HPRate |31 Col_LPRate |32 Col_PartGrade |33 Col_EffectDate
'                    FGrid.AddItem i & Chr(9) & Rst!Part_No & Chr(9) & Rst!DocId & Chr(9) & Rst!ChallIDDisp & Chr(9) & Rst!Srl_No & Chr(9) & Rst!order_docid & Chr(9) & Rst!OrderIDDisp & Chr(9) & Rst!Order_Srl_No & Chr(9) & Rst!Unit & Chr(9) & Rst!MRPYN & Chr(9) & Rst!TaxYN & Chr(9) & Format(Rst!Qty_Iss, "0.000") & Chr(9) & Format(Rst!Rate2, "0.00") & Chr(9) & Format(Rst!MRP, "0.00") & Chr(9) & Format((Rst!Qty_Iss * Rst!Rate2), "0.00") & Chr(9) & Format(Rst!Disc_Per2, "0.00") & Chr(9) & Format(Rst!Disc_Amt2, "0.00") & Chr(9) & Format(Rst!Net_Amt2, "0.00") & Chr(9) & Rst!Godown & Chr(9) & Rst!God_Name & Chr(9) & Rst!Part_Name & Chr(9) & Rst!Local_Name & Chr(9) & Rst!Curstk & Chr(9) & Rst!MRPQty & Chr(9) & Rst!Cur_TB_Stk & Chr(9) & Rst!Cur_TP_Stk & Chr(9) & Rst!TB_SRate & Chr(9) & Rst!TP_SRate & Chr(9) & Rst!Bin_Loca & Chr(9) & " " & Chr(9) & Rst!high_pur_rate & Chr(9) & Rst!low_pur_rate & Chr(9) & Rst!Part_Grade & Chr(9) & Format(IIf(Rst!MRPYN = "Yes", Rst!MRP_Effect_Dt, Rst!TB_Effect_Dt), "dd/MMM/yyyy")
                FGrid.AddItem ""
                With FGrid
                    .TextMatrix(I, Col_SrNo) = I
                    .TextMatrix(I, Col_PNo) = Rst!Part_No
                    .TextMatrix(I, Col_ChalNoCode) = Rst!DocID
                    .TextMatrix(I, Col_ChalNo) = IIf(IsNull(Rst!ChallIDDisp), "", Rst!ChallIDDisp)
                    .TextMatrix(I, Col_ChalSrNo) = Rst!Srl_No
                    .TextMatrix(I, Col_SONoCode) = Rst!Order_DocId
                    .TextMatrix(I, Col_SONo) = IIf(IsNull(Rst!OrderIDDisp), "", Rst!OrderIDDisp)
                    .TextMatrix(I, Col_SOSrNo) = Rst!Order_Srl_No
                    .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                    .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Qty) = Format(Rst!Qty_Iss, "0.000")
                    .TextMatrix(I, Col_IssQty) = Format(Rst!Qty_Iss, "0.000")
                    
                    If txt(DocType) = "Sale Return Credit" Or txt(DocType) = "Sale Return Cash" Then
                        .TextMatrix(I, Col_Rate) = Format(Rst!Rate2, "0.00")
                        .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP_Rate, "0.00")           'Format(Rst!MRP, "0.00")
                        If Rst!MRP_YN = 1 Then
                            .TextMatrix(I, Col_Amt) = Format((Rst!Qty_Iss * Rst!MRP_Rate), "0.00")
                        Else
                            .TextMatrix(I, Col_Amt) = Format((Rst!Qty_Iss * Rst!Rate2), "0.00")
                        End If
                        .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per2, "0.00")
                        .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt2, "0.00")
                        .TextMatrix(I, Col_ItemVal) = Format(Rst!Net_Amt2, "0.00")
                    Else
    '                    FGrid.AddItem i & Chr(9) & Rst!Part_No & Chr(9) & Rst!DocId & Chr(9) & Rst!ChallIDDisp & Chr(9) & Rst!Srl_No & Chr(9) & Rst!order_docid & Chr(9) & Rst!OrderIDDisp & Chr(9) & Rst!Order_Srl_No & Chr(9) & Rst!Unit & Chr(9) & Rst!MRPYN & Chr(9) & Rst!TaxYN & Chr(9) & Format(Rst!Qty_Iss, "0.000") & Chr(9) & Format(Rst!Rate, "0.00") & Chr(9) & Format(Rst!MRP, "0.00") & Chr(9) & Format((Rst!Qty_Iss * Rst!Rate), "0.00") & Chr(9) & Format(Rst!Disc_Per, "0.00") & Chr(9) & Format(Rst!Disc_Amt, "0.00") & Chr(9) & Format(Rst!Net_Amt, "0.00") & Chr(9) & Rst!Godown & Chr(9) & Rst!God_Name & Chr(9) & Rst!Part_Name & Chr(9) & Rst!Local_Name & Chr(9) & Rst!Curstk & Chr(9) & Rst!MRPQty & Chr(9) & Rst!Cur_TB_Stk & Chr(9) & Rst!Cur_TP_Stk & Chr(9) & Rst!TB_SRate & Chr(9) & Rst!TP_SRate & Chr(9) & Rst!Bin_Loca & Chr(9) & " " & Chr(9) & Rst!high_pur_rate & Chr(9) & Rst!low_pur_rate & Chr(9) & Rst!Part_Grade & Chr(9) & Format(IIf(Rst!MRPYN = "Yes", Rst!MRP_Effect_Dt, Rst!TB_Effect_Dt), "dd/MMM/yyyy")
                        .TextMatrix(I, Col_Rate) = Format(Rst!Rate, "0.00")
                        .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP_Rate, "0.00")           'Format(Rst!MRP, "0.00")
                        If Rst!MRP_YN = 1 Then
                            .TextMatrix(I, Col_Amt) = Format((Rst!Qty_Iss * Rst!MRP_Rate), "0.00")
                        Else
                            .TextMatrix(I, Col_Amt) = Format((Rst!Qty_Iss * Rst!Rate), "0.00")
                        End If
                        .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per, "0.00")
                        .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt, "0.00")
                        .TextMatrix(I, Col_ItemVal) = Format(Rst!Net_Amt, "0.00")
                    End If
                    If PubVATYN = 1 Then
                        .TextMatrix(I, Col_TaxPer) = Format(VNull(Rst!TaxPer), "0")
                        .TextMatrix(I, Col_TaxAmt1) = Format(VNull(Rst!TaxAmt), "0.00")
                        .TextMatrix(I, Col_SatPer) = Format(VNull(Rst!SatPer), "0")
                        .TextMatrix(I, Col_SatAmt1) = Format(VNull(Rst!SatAmt), "0.00")
                    End If
                    .TextMatrix(I, Col_GodownCode) = Rst!Godown
                    .TextMatrix(I, Col_Godown) = IIf(IsNull(Rst!God_Name), "", Rst!God_Name)
                    .TextMatrix(I, Col_PartSrlNo) = IIf(IsNull(Rst!Part_SrlNo), "", Rst!Part_SrlNo)
                    .TextMatrix(I, Col_PName) = IIf(IsNull(Rst!Part_Name), "", Rst!Part_Name)
                    .TextMatrix(I, Col_LName) = IIf(IsNull(Rst!Local_Name), "", Rst!Local_Name)
                    .TextMatrix(I, Col_MRPStkTP) = IIf(IsNull(Rst!Cur_MRP_TPStk), "", Rst!Cur_MRP_TPStk)
                    .TextMatrix(I, Col_MRPStkTB) = IIf(IsNull(Rst!Cur_MRP_TbStk), "", Rst!Cur_MRP_TbStk)
                    .TextMatrix(I, Col_TBStk) = IIf(IsNull(Rst!Cur_TB_STk), "", Rst!Cur_TB_STk)
                    .TextMatrix(I, Col_TPStk) = IIf(IsNull(Rst!Cur_TP_Stk), "", Rst!Cur_TP_Stk)
                    .TextMatrix(I, Col_TBRate) = IIf(IsNull(Rst!TB_SRate), "", Rst!TB_SRate)
                    .TextMatrix(I, Col_TPRate) = IIf(IsNull(Rst!TP_SRate), "", Rst!TP_SRate)
                    .TextMatrix(I, Col_Bin) = IIf(IsNull(Rst!Bin_Loca), "", Rst!Bin_Loca)
                    .TextMatrix(I, Col_LastRate) = ""
                    .TextMatrix(I, Col_HPRate) = IIf(IsNull(Rst!high_pur_rate), "", Rst!high_pur_rate)
                    .TextMatrix(I, Col_LPRate) = IIf(IsNull(Rst!low_pur_rate), "", Rst!low_pur_rate)
                    .TextMatrix(I, Col_PartGrade) = IIf(IsNull(Rst!Part_Grade), "", Rst!Part_Grade)
                    .TextMatrix(I, Col_EffectDate) = Format(IIf(Rst!MRP_YN = 1, IIf(IsNull(Rst!MRP_Effect_Dt), "", Rst!MRP_Effect_Dt), IIf(IsNull(Rst!TB_Effect_Dt), "", Rst!TB_Effect_Dt)), "dd/MMM/yyyy")
                End With
                Rst.MoveNext
                I = I + 1
            Loop
            Amt_Cal
            FillDetailFlag = 0
            FGrid.FixedRows = 1
        End If
    Else
        FGrid.Rows = 1
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    CountItem
    Set Rst = Nothing
End Sub
' Used For Updation of Sale Order in case of Edit and Delete
'999
Private Sub UpdateSO()
'Dim Rst As ADODB.Recordset, I As Byte
'    Set Rst = New ADODB.Recordset
'    Rst.CursorLocation = adUseClient
'    Rst.Open "Select * From SP_Stock Where DocId='" & Txt(DocID).Text & "'", GCn, adOpenDynamic, adLockOptimistic
'    If Rst.RecordCount  > 0 Then
'    While Not Rst.EOF
'        If Rst!order_docid <> "" Then
'            GCn.Execute "Update SP_Order1 Set Sup_Qty=Sup_Qty+" & Rst!Qty_iss & " Where OrderId='" & Rst!order_docid & "' and Part_No='" & Rst!Part_No & "'"
'        End If
'        Rst.MoveNext
'    Wend
'    End If
'Set Rst = Nothing
End Sub

Private Sub DGParty_Click()
On Error GoTo ELoop
    If RsParty.RecordCount > 0 Then
        txt(Party).TEXT = RsParty!Name
        txt(Party).Tag = RsParty!Code
    End If
    txt(Party).SetFocus
    DGParty.Visible = False
    lblGroup.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGCrAc_Click()
On Error GoTo ELoop
    If RsCrAc.RecordCount > 0 Then
        txt(CrAc).TEXT = RsCrAc!Name
        txt(CrAc).Tag = RsCrAc!Code
    End If
    txt(CrAc).SetFocus
    DGCrAc.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGBaseDoc_Click()
On Error GoTo ELoop
    If RsBaseDoc.RecordCount > 0 Then
        txt(BaseDoc).TEXT = RsBaseDoc!Name
        txt(BaseDoc).Tag = RsBaseDoc!Code
        FillItemDet
    End If
    txt(BaseDoc).SetFocus
    DGBaseDoc.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGForm_Click()
On Error GoTo ELoop
    If rsForm.RecordCount > 0 Then
        txt(FormName).TEXT = rsForm!Name
        txt(FormName).Tag = rsForm!Code
        If TopCtrl1.TopText2.CAPTION = "Add" Then   ' To Assign Tax% in case of Add
            txt(STaxPer).TEXT = rsForm!Tax_Per
            txt(TaxSurPer).TEXT = rsForm!Tax_Sur_Per
            Amt_Cal
        End If
    End If
    txt(FormName).SetFocus
    DGForm.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGForm31_Click()
On Error GoTo ELoop
    If rsForm31.RecordCount > 0 Then
        txt(Form31Name).TEXT = rsForm31!Name
        txt(Form31Name).Tag = rsForm31!Code
    End If
    txt(Form31Name).SetFocus
    DGForm31.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGPart_Click()
On Error GoTo ELoop
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
    DGPart.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

'Private Sub DGSONo_Click()
'On Error GoTo ELoop
'    DGSONo.Visible = False
'    If RsSONo.RecordCount  > 0 Then
'        TxtGrid(0).Text = RsSONo!Name
'        FGrid.TextMatrix(FGrid.Row, Col_SONoCode) = RsSONo!Code
'        FGrid.TextMatrix(FGrid.Row, Col_SONo) = RsSONo!Name
'    End If
'    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
'Exit Sub
'ELoop:
'    CheckError
'End Sub

Private Sub DGGodown_Click()
On Error GoTo ELoop
    If RsGodown.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsGodown!Name
        FGrid.TextMatrix(FGrid.Row, Col_GodownCode) = RsGodown!Code
        FGrid.TextMatrix(FGrid.Row, Col_Godown) = RsGodown!Name
    End If
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
    DGGodown.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGPerson_Click()
On Error GoTo ELoop
    If RsPerson.RecordCount > 0 Then
        txt(SPerson).TEXT = RsPerson!Name
        txt(SPerson).Tag = RsPerson!Code
    End If
    txt(SPerson).SetFocus
    DGPerson.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Activate()
Dim UnLoadFrm As Boolean, MsgStr$
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
    Call TopCtrl1_eRef
End If
If rsCtrlAc.RecordCount <= 0 Then
    MsgStr = "No Records in Spare A/c Controls, Please Fill Spare "
    UnLoadFrm = True
ElseIf rsCtrlAc!SprSalTP_Ac = "" Or _
    rsCtrlAc!OilSalTB_Ac = "" Or rsCtrlAc!OilSalTP_Ac = "" Or _
    rsCtrlAc!SprCash_Ac = "" Or rsCtrlAc!SprDiscTB_Ac = "" Or rsCtrlAc!SprGenSur_Ac = "" Or _
    rsCtrlAc!Transportation_Ac = "" Or rsCtrlAc!ReSaleTax_Ac = "" Or _
    rsCtrlAc!MiscChrg_Ac = "" Or rsCtrlAc!TOTax_Ac = "" Or rsCtrlAc!SprROff_Ac = "" Then
    MsgStr = "Please Fill Spare"
    UnLoadFrm = True
End If
If PubVATYN = 1 Then
    Lbl(22).CAPTION = "V A T   "
    txt(39).Visible = False
End If
'EOF Spare A/c control checking
If UnLoadFrm Then
    MsgBox "Spare Sale Return Loading Aborted !" & vbCrLf & MsgStr & " A/c Controls through Utility Menu", vbInformation, "Validation"
    Unload Me
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
    TopCtrl1.Tag = PubUParam:    WinSetting Me:     Grid_Ini
    Call Ini_Pub
    For I = 0 To txt.Count - 1
        If I = 18 Then
        Else
            txt(I).BackColor = CtrlBColOrg '&HDFF4F2
            txt(I).ForeColor = CtrlFColOrg
        End If
    Next
    TxtGrid(0).BackColor = CtrlBCol
    TxtGrid(0).ForeColor = CtrlFCol
    Lbl(35) = PubForm31Caption
    Lbl(36) = PubForm31Caption & " No."
    Lbl(25) = pubTOTCaption
    If PubReSaleTaxPer = 0 Then
        Lbl(42).Visible = False
        txt(ReSalTaxPer).Visible = False
        txt(ReSalTaxAmt).Visible = False
    End If
    mVType = RetCashSalVType
    txt(VDate).Tag = PubLoginDate

    'A/c Pstong Control Checking
    Set rsCtrlAc = New ADODB.Recordset
    rsCtrlAc.CursorLocation = adUseClient
    'CSSprAc=Temp Sale A/c
    'SprSalTB_Ac shifted to Tax Forms
    rsCtrlAc.Open "Select SprSalTP_Ac,OilSalTB_Ac,OilSalTP_Ac,CSSprAc,SprGenSur_Ac,ReSaleTax_Ac,SprCash_Ac,SprDiscTB_Ac,Transportation_Ac,MiscChrg_Ac,TOTax_Ac,SprROff_Ac From AcControls", GCnFaS, adOpenDynamic, adLockOptimistic
    'eof checking
    Set DGPart.DataSource = RsPart
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Add1,Transporter,Party_Type,City.CityName from ((SubGroup " & _
        "left Join City On City.CityCode=SubGroup.CityCode) " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode) " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty

    Set RsCrAc = New ADODB.Recordset
    RsCrAc.CursorLocation = adUseClient
    RsCrAc.Open "Select SubCode as Code,Name From SubGroup Where  Nature='Revenue' Order by Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGCrAc.DataSource = RsCrAc

    Set RsBaseDoc = New ADODB.Recordset
    RsBaseDoc.CursorLocation = adUseClient
    RsBaseDoc.Open "Select DocID as Code,left(DocID,1)+ " & cMID("DocID", "3", "1") & "+'/'+left(" & cTrim(cMID("DocID", "4", "5")) & ",1) + right(" & cMID("Docid", "4", "5") & ",len(" & cMID("Docid", "4", "5") & ")-2)+'/' + " & cTrim("Right(DocID,8)") & " As Name,V_Date,Party_Code From SP_Sale Where V_Type='" & BaseDocType & "' Order By DocID", GCn, adOpenDynamic, adLockOptimistic
    Set DGBaseDoc.DataSource = RsBaseDoc

    Set rsForm = New ADODB.Recordset
    rsForm.CursorLocation = adUseClient
    rsForm.Open "Select T.Form_Code as Code,T.Form_Desc As Name,T.Tax_Per,T.Tax_Sur_Per,T1.Tax_Ac_Code,T1.Sur_Ac_Code,T1.PurSal_Ac_Code " & _
        "From TaxForms as T left Join TaxFormsAc as T1 on T.Form_Code+'" & PubDivCode & "'=T1.Form_Code+T1.Div_Code " & _
        "Where Trn_Type='Sale' and Spare_YN=1 Order by Form_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGForm.DataSource = rsForm

    Set rsForm31 = New ADODB.Recordset
    rsForm31.CursorLocation = adUseClient
    rsForm31.Open "Select Form_Code as Code,Form_Desc As Name From TaxForms Where Spare_YN=1 and Trn_Type='Form 31' Order by Form_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGForm31.DataSource = rsForm31

'    Set RsSONo = New ADODB.Recordset
'    RsSONo.CursorLocation = adUseClient
'    RsSONo.Open "Select OrderID as Code,Trim(Mid(OrderID,9,5))+CStr(Trim(Right(OrderID,8))) As Name,V_Date,Qty,Rate,Switch(TAX_YN=1,'Yes',TAX_YN=0,'No') As TAXYN,Switch(MRP_YN=1,'Yes',MRP_YN=0,'No') As MRPYN From SP_Order1 Where Order_Type='S_SO' Order By OrderID", GCn, adOpenDynamic, adLockOptimistic
'    Set DGSONo.DataSource = RsSONo

    Set RsGodown = New ADODB.Recordset
    RsGodown.CursorLocation = adUseClient
    RsGodown.Open "Select God_Code as Code,God_Name As Name From Godown Where Appli_For=0 Order by God_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGodown.DataSource = RsGodown
    
    Set RsPerson = New ADODB.Recordset
    RsPerson.CursorLocation = adUseClient
    RsPerson.Open "Select Emp_Code as Code, Emp_Name as Name From Emp_Mast Where Emp_Type=0 and (LeftOn Is Null or LeftOn< " & ConvertDate(PubLoginDate) & ") Order By Emp_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGPerson.DataSource = RsPerson

    Dim SiteCond As String
    SiteCond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and  " & cMID("s.docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    If PubMoveRecYn Then
        Set Master = GCn.Execute("Select S.DocID As SearchCode From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' " & SiteCond & " and S.V_Type In ('" & RetCashSalVType & "','" & RetCrSalVType & "','" & RetTrfIssVType & "') " _
            & "Order by S.V_Date desc,S.DocID Desc")
    Else
        Set Master = GCn.Execute("Select Top 1 S.DocID As SearchCode From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' and S.V_Type In ('" & RetCashSalVType & "','" & RetCrSalVType & "','" & RetTrfIssVType & "') " _
            & "Order by S.V_Date desc,S.DocID Desc")
    Disp_Text SETS("INI", Me, Master)
    
    End If
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsParty = Nothing
    Set RsCrAc = Nothing
    Set rsForm = Nothing
    Set rsForm31 = Nothing
    Set RsBaseDoc = Nothing
'    Set RsSONo = Nothing
    Set RsGodown = Nothing
    Set Master = Nothing
End Sub

Private Sub ListView_Click()
On Error GoTo ELoop
    txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    txt(Val(ListView.Tag)).SetFocus
    FrmList.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    If PubSatYn = 1 Then mSatYn = True
    DispTextVat
    'CtrlEnbChallan True
    txt(VDate).TEXT = txt(VDate).Tag
    txt(DocType).TEXT = "Sale Return Credit"
    txt(LC).TEXT = "Local"
    txt(TaxDet).TEXT = "Yes"
    txt(AcPost).TEXT = "Yes"
    txt(ReSalTaxPer) = IIf(PubReSaleTaxPer = 0, "", Format(PubReSaleTaxPer, "0.00"))
    mPartyType = 0
    txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
    txt(DocID).Tag = txt(DocID)
    txt(DocType).SetFocus
    txt(TurnOverPer) = MainLib.TOTCal()
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
    txt(DocType).Enabled = False
    txt(VDate).Enabled = False
    txt(SerialNo).Enabled = False
    txt(BaseDoc).Enabled = False
    If txt(BaseDoc).TEXT <> "" Then CtrlEnbChallan False Else CtrlEnbChallan True
    FGrid.AddItem FGrid.Rows
    'Enable / Disable Text Box if values zero
    DisableEnableFooter txt(MRPAmtTB), txt(MRPAmtTP), txt(SprAmtTB), txt(SprAmtTP), _
            txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), txt(DiscPerTP), _
            txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), _
            txt(GenSurPer), txt(GenSurAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt)
    'EOF enable / disable section
    txt(Party).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim vBook As Variant, mTrans As Boolean, Rst As ADODB.Recordset
Dim LedgAry(1) As LedgRec, mResult As Byte, MsgStr$, mTitle$
    If Master.RecordCount > 0 Then
        If GCn.Execute("Select CancelYN from SP_Sale where DocID='" & Master!SearchCode & "'").Fields(0).Value = 1 Then
            MsgStr = "Are You Sure To Delete This ? "
            mTitle = "Delete Entry!"
        Else
            MsgStr = "Are You Sure To Cancel This ? "
            mTitle = "Cancel Entry!"
        End If
        If MsgBox(MsgStr, vbYesNo + vbCritical + vbDefaultButton2, mTitle) = vbYes Then
            vBook = Master.AbsolutePosition
            GCn.BeginTrans
'            UpdateSO
            GCnFaS.BeginTrans
            mTrans = True
            'Stock unposting
            UpdStkTableToTable txt(DocID), "-", "I"
            'eof stock unposting
            GCn.Execute ("Delete From SP_Stock Where DocID='" & txt(DocID) & "'")
            If mTitle = "Delete Entry!" Then
                GCn.Execute ("Delete From SP_Sale Where DocID='" & txt(DocID) & "'")
            Else
                GCn.Execute "Update SP_Sale Set " _
                    & "CancelYN=1,RoadPermit_FormCode='',RoadPermit_No='',CrAc='',SprAmt_MRP_TB=0 " & _
                    ",SprAmt_MRP_TP=0,SprAmt_TB=0,OilAmt_MRP_TB=0,OilAmt_MRP_TP=0,SprAmt_TP=0," & _
                    "OilAmt_TB=0,OilAmt_TP=0,D_Per_TB=0,D_Amt_TB=0,D_Per_TP=0,D_Amt_TP=0,Packing=0 " & _
                    ",Gen_Sur_Per=0,Gen_Sur_Amt=0,Trans_Amt=0,Tax_Per=0,Tax_Amt=0,Tax_Sur_Per=0 " & _
                    ",Tax_Sur_Amt=0,TOT_Per=0,TOT_Amt=0,ReSalTax_Per=0,ReSalTax_Amt=0,Rounded=0 " & _
                    ",Total_Amt=0,U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' " & _
                    ",D_Per_MRP_TB=0,D_Amt_MRP_TB=0,D_Per_MRP_TP=0,D_Amt_MRP_TP=0,Tax_AmtMRP=0" & _
                    ",TaxSur_AmtMRP=0, TOT_AmtMRP= 0 " & _
                    " Where DocID='" & txt(DocID) & "'"
            End If
            'Unpost Ledger a/c
            If txt(DocType) = "Sale Return Cash" Then
                'A/c Posting
                ProcAcPost rsCtrlAc
                'EOF Posting
            Else
                mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, txt(DocID))
                If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
                'Unposting of Ledger completed
            End If
            '**eoc A/c Posting
            GCnFaS.CommitTrans
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
Exit Sub
ELoop:
    If mTrans Then GCnFaS.RollbackTrans: GCn.RollbackTrans
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
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    Dim SiteCond As String
    SiteCond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and  " & cMID("s.docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    If PubBackEnd = "A" Then
        GSQL = "Select S.DocId As SearchCode,S.Site_Code, " & cTrim("Right(Invoice_DocId,8)") & " As Ref_V_No, Switch(S.V_Type='" & RetCashSalVType & "','Sale Return Cash',S.V_Type='" & RetCrSalVType & "','Sale Return Credit',S.V_Type='" & RetTrfIssVType & "','Transfer Issue Return') As VType,Trim(" & cMID("S.DocID", "9", "5") & ") As VPrefix, S.V_No, " & cDt("S.V_Date") & " AS VDate, S.Party_Name as PartyName From SP_Sale S Where left(S.DocID,1)='" & PubDivCode & "' " & SiteCond & " and S.V_Type In ('" & RetCashSalVType & "','" & RetCrSalVType & "','" & RetTrfIssVType & "') Order by S.V_Date Desc,S.V_Type"
    ElseIf PubBackEnd = "S" Then
        GSQL = "Select S.DocId As SearchCode,S.Site_Code, " & cTrim("Right(Invoice_DocId,8)") & " As Ref_V_No, Case When S.V_Type='" & RetCashSalVType & "' Then 'Sale Return Cash' When S.V_Type='" & RetCrSalVType & "' Then 'Sale Return Credit' When S.V_Type='" & RetTrfIssVType & "' Then 'Transfer Issue Return' End As VType, " & cTrim(cMID("S.DocID", "9", "5")) & " As VPrefix, S.V_No, " & cDt("S.V_Date") & " AS VDate, S.Party_Name as PartyName From SP_Sale S Where left(S.DocID,1)='" & PubDivCode & "' " & SiteCond & " and S.V_Type In ('" & RetCashSalVType & "','" & RetCrSalVType & "','" & RetTrfIssVType & "') Order by S.V_Date Desc,S.V_Type"
    End If
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
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
    RsParty.Requery
    RsCrAc.Requery
    rsForm.Requery
    rsForm31.Requery
    RsBaseDoc.Requery
    RsPart.Requery
'    RsSONo.Requery
    RsGodown.Requery
    'Master.Requery
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean, mGridFilled As Boolean
Dim Rst As ADODB.Recordset, DocIdHlp As String
On Error GoTo ELoop
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid_LostFocus 0
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    
    If IsValid(txt(DocType), "Document Type") = False Then Exit Sub
    If IsValid(txt(VDate), "Date") = False Then Exit Sub
    If IsValid(txt(SerialNo), "Serial Number") = False Then Exit Sub
    If IsValid(txt(Party), "Party Name") = False Then Exit Sub
    If IsValid(txt(FormName), "Form") = False Then Exit Sub
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            If FGrid.TextMatrix(I, Col_MRP) = "" Then MsgBox "Please Specify MRP Yes/No in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_MRP: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Col_Taxable) = "" Then MsgBox "Please Specify Taxable Yes/No in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Taxable: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, Col_Qty)) = 0 Then MsgBox "Please Specify Quantity in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Qty: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Col_Godown) = "" Then MsgBox "Please Specify Godown in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Godown: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, Col_Rate)) = 0 Then
'                If PubULabel <> "Y" Then
                    MsgBox "Please Specify Rate in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Qty: FGrid.SetFocus: Exit Sub
'                End If
            End If
            mGridFilled = True
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Item Detail", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = Col_PNo: FGrid.SetFocus: Exit Sub
    'Amount Calculation
    MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
            Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
            Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
            Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
    
    'MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
        Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
        Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
        Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
        Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
        Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
        Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
        Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
        Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
    If PubVATYN = 1 Then
       MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
            txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
            txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
            txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
            txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
            txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt1, txt(SatAmt)
    Else
        MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
            txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
            txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
            txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
            txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
            txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
    End If
    'EOF Amount Calculation
    GCn.BeginTrans
    GCnFaS.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2 = "Add" Then
        'lp 12-03-03
        txt(DocID).Tag = txt(DocID)
        If GCn.Execute("Select Count(*) From SP_Sale Where Left(DocID,1)='" & PubDivCode & "'  And V_Type = '" & mVType & "'  And V_No = " & Val(txt(SerialNo)) & " ").Fields(0) > 0 Then
            If VoucherEditFlag Then
                MsgBox "Sale Return No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                GoTo ELoop
            Else
                txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(txt(DocID).Tag, Document_No)) Then
                    MsgBox "Sale Return No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo ELoop
                End If
            End If
        End If
        DocIdHlp = UCase(Replace(txt(DocID), " ", ""))
        '**************
        GCn.Execute "Insert Into SP_Sale(" _
            & "DocID ,DocIDHelp ,V_Type,V_No,Site_Code,Cash_Credit," _
            & "V_Date ,Party_Code ,Party_Name ,Address ,L_C ," _
            & "Form_Code ,RoadPermit_FormCode ,RoadPermit_No ,CrAc ,Case_No ," _
            & "Case_Mark ,Rep_Code ,Remarks ,Det_Tax," _
            & "Invoice_DocID,SprAmt_MRP_TB ,SprAmt_MRP_TP ," _
            & "OilAmt_MRP_TB,OilAmt_MRP_TP,SprAmt_TB ,SprAmt_TP ,OilAmt_TB ,OilAmt_TP ," _
            & "D_Per_TB ,D_Amt_TB ,D_Per_TP ,D_Amt_TP ,Addition ," _
            & "Packing ,Gen_Sur_Per ,Gen_Sur_Amt ,Trans_Amt ,Tax_Per ," _
            & "Tax_Amt ,Tax_Sur_Per ,Tax_Sur_Amt ,TOT_Per ,TOT_Amt ," _
            & "ReSalTax_Per, ReSalTax_Amt,Rounded ,Total_Amt,U_Name ,U_EntDt ,U_AE, " _
            & "D_Per_MRP_TB,D_Amt_MRP_TB, D_Per_MRP_TP , D_Amt_MRP_TP, Tax_AmtMRP, TaxSur_AmtMRP, TOT_AmtMRP, SatAmt, Sat_Yn) " _
            & "Values('" & txt(DocID) & "','" & DocIdHlp & "','" & mVType & "'," & txt(SerialNo) & ",'" & PubSiteCode & PubSiteCode & "','" & IIf(mVType = RetCashSalVType, "Cash", "Credit") & _
            "'," & ConvertDate(Format(txt(VDate), "dd/MMM/yyyy")) & ",'" & txt(Party).Tag & "','" & txt(Party) & "','" & txt(Address1) & "','" & left(txt(LC), 1) & _
            "','" & txt(FormName).Tag & "','" & txt(Form31Name).Tag & "','" & txt(Form31No) & "','" & txt(CrAc).Tag & "'," & Val(txt(CaseNo)) & _
            ",' " & txt(CaseMark) & "','" & txt(SPerson).Tag & "','" & txt(Remark) & "'," & IIf(txt(TaxDet) = "Yes", 1, 0) & ",'" & txt(BaseDoc).Tag & "'," & Val(txt(MRPAmtTB)) - mMRPLubeTB & _
            " , " & Val(txt(MRPAmtTP)) - mMRPLubeTP & "," & mMRPLubeTB & "," & mMRPLubeTP & "," & Val(txt(SprAmtTB)) & "," & Val(txt(SprAmtTP)) & "," & Val(txt(OilAmtTB)) & "," & Val(txt(OilAmtTP)) & _
            " , " & Val(txt(DiscPerTB)) & "," & Val(txt(DiscAmtTB)) & "," & Val(txt(DiscPerTP)) & "," & Val(txt(DiscAmtTP)) & "," & Val(txt(Addition)) & _
            " , " & Val(txt(PackCrg)) & "," & Val(txt(GenSurPer)) & "," & Val(txt(GenSurAmt)) & "," & Val(txt(TransAmt)) & "," & Val(txt(STaxPer)) & _
            " , " & Val(txt(STaxAmt)) & "," & Val(txt(TaxSurPer)) & "," & Val(txt(TaxSurAmt)) & "," & Val(txt(TurnOverPer)) & "," & Val(txt(TurnOverAmt)) & _
            " , " & Val(txt(ReSalTaxPer)) & "," & Val(txt(ReSalTaxAmt)) & "," & Val(txt(SROff)) & "," & Val(txt(NetAmt)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & _
            ",'A'," & mMRevDisTBPer & "," & mTBDisAmtMRP & "," & mMRevDisTPPer & "," & mTPDisAmtMRP & "," & mMRPTax & "," & mMRPTaxSur & ", " & mMRPTOT & ", " & Val(txt(SatAmt)) & ", " & IIf(mSatYn, 1, 0) & ")"
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaS, txt(DocID), txt(VDate)
    Else
'            UpdateSO
        'Stock unposting
        UpdStkTableToTable txt(DocID), "-", "R"
        'eof stock unposting
        GCn.Execute ("Delete From SP_Stock Where DocID='" & txt(DocID) & "'")
        GCn.Execute "Update SP_Sale Set " _
            & "Party_Code='" & txt(Party).Tag & "',Party_Name='" & txt(Party) & "',Address='" & txt(Address1) & _
            "',L_C='" & left(txt(LC), 1) & "',Form_Code='" & txt(FormName).Tag & "',RoadPermit_FormCode='" & txt(Form31Name).Tag & _
            "',RoadPermit_No='" & txt(Form31No) & "',CrAc='" & txt(CrAc).Tag & "',Case_No=" & Val(txt(CaseNo)) & _
            ",Case_Mark='" & txt(CaseMark) & "',Rep_Code='" & txt(SPerson).Tag & "',Remarks='" & txt(Remark) & _
            "',Det_Tax=" & IIf(txt(TaxDet) = "Yes", 1, 0) & ",Invoice_DocID='" & txt(BaseDoc).Tag & "',SprAmt_MRP_TB=" & Val(txt(MRPAmtTB)) - mMRPLubeTB & _
            ",SprAmt_MRP_TP=" & Val(txt(MRPAmtTP)) - mMRPLubeTP & ",SprAmt_TB=" & Val(txt(SprAmtTB)) & _
            ",OilAmt_MRP_TB=" & mMRPLubeTB & ",OilAmt_MRP_TP=" & mMRPLubeTP & _
            ",SprAmt_TP=" & Val(txt(SprAmtTP)) & ",OilAmt_TB=" & Val(txt(OilAmtTB)) & _
            ",OilAmt_TP=" & Val(txt(OilAmtTP)) & ",D_Per_TB=" & Val(txt(DiscPerTB)) & _
            ",D_Amt_TB=" & Val(txt(DiscAmtTB)) & ",D_Per_TP=" & Val(txt(DiscPerTP)) & _
            ",D_Amt_TP=" & Val(txt(DiscAmtTP)) & ",Addition=" & Val(txt(Addition)) & _
            ",Packing=" & Val(txt(PackCrg)) & ",Gen_Sur_Per=" & Val(txt(GenSurPer)) & _
            ",Gen_Sur_Amt=" & Val(txt(GenSurAmt)) & ",Trans_Amt=" & Val(txt(TransAmt)) & _
            ",Tax_Per=" & Val(txt(STaxPer)) & ",Tax_Amt=" & Val(txt(STaxAmt)) & _
            ",Tax_Sur_Per=" & Val(txt(TaxSurPer)) & ",Tax_Sur_Amt=" & Val(txt(TaxSurAmt)) & _
            ",TOT_Per=" & Val(txt(TurnOverPer)) & ",TOT_Amt=" & Val(txt(TurnOverAmt)) & _
            ",ReSalTax_Per=" & Val(txt(ReSalTaxPer)) & ",ReSalTax_Amt=" & Val(txt(ReSalTaxAmt)) & _
            ",Rounded=" & Val(txt(SROff)) & ",Total_Amt=" & Val(txt(NetAmt)) & _
            ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' " & _
            ",D_Per_MRP_TB=" & mMRevDisTBPer & ",D_Amt_MRP_TB=" & mTBDisAmtMRP & _
            ",D_Per_MRP_TP=" & mMRevDisTPPer & " , D_Amt_MRP_TP=" & mTPDisAmtMRP & _
            ",Tax_AmtMRP=" & mMRPTax & ", TaxSur_AmtMRP=" & mMRPTaxSur & ", SatAmt = " & Val(txt(SatAmt)) & ", TOT_AmtMRP= " & mMRPTOT & _
            " Where DocID='" & txt(DocID) & "'"
    End If

    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
'            If FGrid.TextMatrix(i, Col_SONo) <> "" Then
'                GCn.Execute "Update SP_Order1 Set Sup_Qty=Sup_Qty-" & Val(FGrid.TextMatrix(i, Col_Qty)) & " Where OrderId='" & FGrid.TextMatrix(i, Col_SONoCode) & "' and Part_No='" & FGrid.TextMatrix(i, Col_PNo) & "'"
'            End If
            GCn.Execute "Insert Into SP_Stock(" _
                & "DocID,Srl_No,V_Type,V_No,V_Date,Site_Code," _
                & "Party_Code,L_C,Order_DocId,Order_Srl_No,Part_No," _
                & "Godown,Qty_Rec,Tax_YN,MRP_YN,Rate," _
                & "MRP_Rate,Disc_Per,Disc_Amt,Amount,Net_Amt," _
                & "Invoice_DocId,Part_SrlNo,U_Name,U_EntDt,U_AE,TaxPer,TaxAmt, SatPer, SatAmt) " _
                & "Values('" & txt(DocID) & "'," & I & ",'" & mVType & "'," & txt(SerialNo) & "," & ConvertDate(Format(txt(VDate), "dd/MMM/yyyy")) & ",'" & PubSiteCode & PubSiteCode & _
                "','" & txt(Party).Tag & "','" & left(txt(LC), 1) & "','" & FGrid.TextMatrix(I, Col_SONoCode) & "'," & Val(FGrid.TextMatrix(I, Col_SOSrNo)) & ",'" & FGrid.TextMatrix(I, Col_PNo) & _
                "','" & FGrid.TextMatrix(I, Col_GodownCode) & "'," & Val(FGrid.TextMatrix(I, Col_Qty)) & "," & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & "," & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & "," & Val(FGrid.TextMatrix(I, Col_Rate)) & _
                " , " & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," & Val(FGrid.TextMatrix(I, Col_DiscPer)) & "," & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," & Val(FGrid.TextMatrix(I, Col_Amt)) & "," & Val(FGrid.TextMatrix(I, Col_ItemVal)) & _
                " ,'" & txt(BaseDoc).Tag & "','" & FGrid.TextMatrix(I, Col_PartSrlNo) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(TopCtrl1.TopText2 = "Add", "A", "E") & "'," & Val(FGrid.TextMatrix(I, Col_TaxPer)) & ", " & Val(FGrid.TextMatrix(I, Col_TaxAmt1)) & "," & Val(FGrid.TextMatrix(I, Col_SatPer)) & ", " & Val(FGrid.TextMatrix(I, Col_SatAmt1)) & ")"
            Call UpdStkGridToTable(FGrid.TextMatrix(I, Col_PNo), "+", FGrid.TextMatrix(I, Col_MRP), FGrid.TextMatrix(I, Col_Taxable), FGrid.TextMatrix(I, Col_Qty))
        End If
    Next
    'A/c Posting
    'If UCase(Txt(AcPost)) = "YES" Then
        ProcAcPost rsCtrlAc
    'End If
    'EOf of A/c Posting
    
    GCnFaS.CommitTrans
    GCn.CommitTrans
    mTrans = False
    mSearchCode = txt(DocID)
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select S.DocID As SearchCode From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' and S.V_Type In ('" & RetCashSalVType & "','" & RetCrSalVType & "','" & RetTrfIssVType & "') And S.DocID = '" & mSearchCode & "' " _
            & "Order by S.V_Date desc,S.DocID Desc")
    End If
    Master.FIND "SearchCode = '" & mSearchCode & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(SerialNo)) > Val(DeCodeDocID(txt(DocID).Tag, Document_No)) Then
            MsgBox "Document No." & Trim(DeCodeDocID(txt(DocID).Tag, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
        txt(VDate).Tag = txt(VDate).TEXT
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    If mTrans Then GCnFaS.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Master.FIND "SearchCode='" & mSearchCode & "'"
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To txt.Count - 1
            If I = 18 Then
            Else
                txt(I).BackColor = CtrlBColOrg
                txt(I).ForeColor = CtrlFColOrg
            End If
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
        ListArray = Array("Sale Return Cash", "Sale Return Credit", "Transfer Issue Return")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 3)
    Case Party
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case SPerson
        If RsPerson.RecordCount = 0 Or (RsPerson.EOF = True Or RsPerson.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsPerson!Name Then
            RsPerson.MoveFirst
            RsPerson.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case FormName
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsForm!Name Then
            rsForm.MoveFirst
            rsForm.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case Form31Name
        If rsForm31.RecordCount = 0 Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsForm31!Name Then
            rsForm31.MoveFirst
            rsForm31.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case BaseDoc
        OldBaseDoc = txt(BaseDoc).TEXT
    Case LC
        ListArray = Array("Local", "Central")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
    Case DiscPerTB, DiscPerTP, GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ReSalTaxPer
        txt(Index).Tag = txt(Index)
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
        Case DocType
            ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 900
        Case SerialNo
            NumDown txt(Index), KeyCode, 8, 0
        Case LC
            ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
        Case Party
            If txt(DocType).TEXT = "Sale Return Credit" Or txt(DocType).TEXT = "Transfer Issue Return" Then
                DGridTxtKeyDown DGParty, txt, Party, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
            End If
        Case SPerson
            DGridTxtKeyDown DGPerson, txt, SPerson, RsPerson, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
        Case BaseDoc
            DGridTxtKeyDown DGBaseDoc, txt, BaseDoc, RsBaseDoc, KeyCode, False, 1
        Case CrAc
            If txt(DocType).TEXT = "Sale Return Credit" Then
                DGridTxtKeyDown DGCrAc, txt, CrAc, RsCrAc, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
            End If
        Case FormName
            DGridTxtKeyDown DGForm, txt, FormName, rsForm, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
        Case Form31Name
            DGridTxtKeyDown DGForm31, txt, Form31Name, rsForm31, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
        Case CaseNo
            NumDown txt(Index), KeyCode, 8, 0
        Case DiscAmtTB, DiscAmtTP, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, PackCrg, TurnOverAmt, ReSalTaxAmt
            NumDown txt(Index), KeyCode, 8, 2
        Case DiscPerTB, DiscPerTP, GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ReSalTaxPer
            NumDown txt(Index), KeyCode, 2, 2
    End Select
    If FrmList.Visible = False And DGParty.Visible = False And DGCrAc.Visible = False And _
        DGBaseDoc.Visible = False And DGForm.Visible = False And DGForm31.Visible = False And DGPerson.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And _
            ((PubReSaleTaxPer = 0 And Index = TurnOverAmt) Or _
            (PubReSaleTaxPer <> 0 And Index = ReSalTaxAmt)) Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        Else
            If Index <> SROff Then
                If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
            End If
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" Then
            If Index <> DocType And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
            If Index <> Party And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
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
    Case Party
        If DGParty.Visible = True Then DGridTxtKeyPress txt, Party, RsParty, KeyAscii, "Name"
        lblGroup.Visible = True: lblGroup.Locked = True: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
    Case CrAc
        If txt(DocType).TEXT = "Sale Return Credit" Or txt(DocType).TEXT = "Transfer Issue Return" Then
            If DGCrAc.Visible = True Then DGridTxtKeyPress txt, CrAc, RsCrAc, KeyAscii, "Name"
        End If
    Case SPerson
        If DGPerson.Visible = True Then DGridTxtKeyPress txt, SPerson, RsPerson, KeyAscii, "Name"
    Case BaseDoc
        If txt(DocType).TEXT = "Sale Return Credit" Or txt(DocType).TEXT = "Transfer Issue Return" Then
            If DGBaseDoc.Visible = True Then DGridTxtKeyPress txt, BaseDoc, RsBaseDoc, KeyAscii, "Name"
        End If
    Case FormName
        If DGForm.Visible = True Then DGridTxtKeyPress txt, FormName, rsForm, KeyAscii, "Name"
    Case Form31Name
        If DGForm31.Visible = True Then DGridTxtKeyPress txt, Form31Name, rsForm31, KeyAscii, "Name"
    Case CaseNo
        NumPress txt(Index), KeyAscii, 8, 0
    Case TaxDet, AcPost
        If KeyAscii = 89 Or KeyAscii = 121 Or KeyAscii = 78 Or KeyAscii = 110 Then
            If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                txt(Index).TEXT = "Yes"
                KeyAscii = 0
            ElseIf KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                txt(Index).TEXT = "No"
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
    Case DiscAmtTB, DiscAmtTP, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, PackCrg, TurnOverAmt, ReSalTaxAmt
        NumPress txt(Index), KeyAscii, 8, 2
    Case DiscPerTB, DiscPerTP, GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ReSalTaxPer
        NumPress txt(Index), KeyAscii, 2, 2
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
    Case LC
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case DiscPerTB, DiscAmtTB, DiscPerTP, DiscAmtTP, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, TurnOverPer, PackCrg, TurnOverAmt, ReSalTaxPer, ReSalTaxAmt
        If Val(txt(MRPAmtTB)) + Val(txt(MRPAmtTP)) <> 0 Then
            MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
                Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
                Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
                Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
        End If
        If PubVATYN = 1 Then
               MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt1, txt(SatAmt)
         Else
                MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
        End If
       ' MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
            Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
            Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
            Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
            Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
            Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
            Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, I As Byte
On Error GoTo ELoop
    Select Case Index
        Case DocType
            txt(Index).TEXT = ListView.SelectedItem.TEXT
            If Not (Trim(txt(Index).TEXT) <> "Sale Return Cash" Or Trim(txt(Index).TEXT) <> "Sale Return Credit" Or Trim(txt(Index).TEXT) <> "Transfer Issue Return") Then
                txt(Index).TEXT = "Sale Return Credit"
            End If
            If Trim(txt(Index).TEXT) = "Sale Return Cash" Then
                txt(Party).Tag = PubSprCashAc
                txt(CrAc).Enabled = False
                txt(CrAc).TEXT = ""
                mVType = RetCashSalVType    ' "S_SRC"
            ElseIf Trim(txt(Index).TEXT) = "Sale Return Credit" Then
                txt(CrAc).Enabled = True
                mVType = RetCrSalVType  ' "S_SRR"
            ElseIf Trim(txt(Index).TEXT) = "Transfer Issue Return" Then
                txt(CrAc).Enabled = True
                mVType = RetTrfIssVType '"S_SRT"
            End If
            txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            txt(DocID).Tag = txt(DocID)
        Case VDate
            txt(Index).TEXT = RetDate(txt(Index))
            Cancel = Not CheckFinYear(txt(Index))
            If Cancel = False Then
                txt(DocID).TEXT = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                txt(DocID).Tag = txt(DocID)
            End If
        Case SerialNo
            If VoucherEditFlag = True Then      ' Manual
                txt(DocID).TEXT = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                txt(DocID).Tag = txt(DocID)
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select V_No From SP_Sale Where DocID='" & txt(DocID).TEXT & "'", GCn, adOpenStatic, adLockReadOnly
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                    Cancel = True
                    txt(SerialNo).SetFocus
                End If
            End If
        Case Party
            If Trim(txt(Index).TEXT = "") Then
                MsgBox "Please Select Party", vbInformation, "Information"
                txt(Index).SetFocus
                Cancel = True
                Exit Sub
            End If
            ' To Populate Base Doc. Data Grid for the Customer
            If txt(DocType).TEXT = "Sale Return Cash" Then
                txt(Party).Tag = PubSprCashAc
                BaseDocType = CashSalVType  ' "S_SIC"
                mPartyType = 0
            ElseIf txt(DocType).TEXT = "Sale Return Credit" Then
                If RsParty.RecordCount > 0 And txt(Index).TEXT <> "" Then
                    txt(Index).TEXT = RsParty!Name
                    txt(Index).Tag = RsParty!Code
                    txt(Address1).TEXT = RsParty!Add1
                    mPartyType = 0
                End If
                BaseDocType = CrSalVType ' "S_SIR"
            ElseIf txt(DocType).TEXT = "Transfer Issue Return" Then
                If RsParty.RecordCount > 0 And txt(Index).TEXT <> "" Then
                    txt(Index).TEXT = RsParty!Name
                    txt(Index).Tag = RsParty!Code
                    txt(Address1).TEXT = RsParty!Add1
                    mPartyType = 0
               End If
                BaseDocType = TrfIssVType '"S_DCT"
            End If
            Set RsBaseDoc = New ADODB.Recordset
            RsBaseDoc.CursorLocation = adUseClient
            RsBaseDoc.Open "Select DocID as Code,left(DocID,1)+ " & cMID("DocID", "3", "1") & "+'/' + left(" & cTrim(cMID("DocID", "4", "5")) & ",1)+right(" & cMID("Docid", "4", "5") & ",len(" & cMID("Docid", "4", "5") & ")-2)+'/'+ " & cTrim("Right(DocID,8)") & " As Name,V_Date,Party_Code From SP_Sale Where left(DocID,1)='" & PubDivCode & "' and V_Type='" & BaseDocType & "' and Party_Code='" & txt(Party).Tag & "' Order By DocID", GCn, adOpenDynamic, adLockOptimistic
'            RsBaseDoc.Open "Select DocID as Code,Trim(Mid(DocID,9,5))+CStr(Trim(Right(DocID,8))) as Name,V_Date From SP_Sale Where left(DocID,1)='" & PubDivCode & "' and V_Type='" & BaseDocType & "' and Party_Code='" & Txt(Party).Tag & "' Order By DocID", GCn, adOpenDynamic, adLockOptimistic
            Set DGBaseDoc.DataSource = RsBaseDoc
        Case BaseDoc
            If txt(Index).TEXT = OldBaseDoc Then Exit Sub
            If RsBaseDoc.RecordCount > 0 Then
                If txt(Index).TEXT <> "" Then
                    txt(Index).TEXT = RsBaseDoc!Name
                    txt(Index).Tag = RsBaseDoc!Code
                    Set GRs = New ADODB.Recordset
                    GRs.CursorLocation = adUseClient
                    GRs.Open "Select SP_Sale.Form_Code,TF.Form_Desc from SP_Sale left join TaxForms as TF on SP_Sale.Form_Code=TF.Form_Code " & _
                        " where DocID='" & txt(Index).Tag & "'", GCn, adOpenStatic, adLockReadOnly
                    txt(FormName) = XNull(GRs!form_desc)
                    txt(FormName).Tag = XNull(GRs!Form_Code)
                    Set GRs = Nothing
                    CtrlEnbChallan False
                Else
                    CtrlEnbChallan True
                    txt(Index).Tag = ""
                    txt(FormName).TEXT = ""
                    txt(FormName).Tag = ""
                    txt(Form31Name).TEXT = ""
                    txt(Form31No).TEXT = ""
                    txt(SPerson).TEXT = ""
                    txt(CaseNo).TEXT = ""
                    txt(CaseMark).TEXT = ""
                    txt(LC).TEXT = ""
                End If
                FillItemDet
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        Case LC
            txt(Index).TEXT = ListView.SelectedItem.TEXT
        Case FormName
            If rsForm.RecordCount > 0 Then
                If Trim(txt(Index).TEXT = "") Then
                    MsgBox "Please Select Form Type", vbInformation, "Information"
                    txt(Index).SetFocus
                    Cancel = True
                    Exit Sub
                Else        'If Txt(Index).Text <> "" Then
                    txt(Index).TEXT = rsForm!Name
                    txt(Index).Tag = rsForm!Code
                    If TopCtrl1.TopText2.CAPTION = "Add" Then   ' To Assign Tax% in case of Add
                        txt(STaxPer).TEXT = rsForm!Tax_Per
                        txt(TaxSurPer).TEXT = rsForm!Tax_Sur_Per
                        MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
                                Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
                                Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
                                Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
            If PubVATYN = 1 Then
               MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt1, txt(SatAmt)
            Else
                MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
            End If
                       ' MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                            Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
                            Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
                            Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
                            Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
                            Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
                            Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
                            Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
                    End If
                End If
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        Case Form31Name
            If rsForm31.RecordCount > 0 Then
                If txt(Index).TEXT <> "" Then
                    txt(Index).TEXT = rsForm31!Name
                    txt(Index).Tag = rsForm31!Code
                End If
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        Case TaxDet
            If Not Trim(txt(Index).TEXT) <> "Yes" Or Trim(txt(Index).TEXT) <> "No" Then
                txt(Index).TEXT = "Yes"
            End If
        Case CaseNo, DiscPerTB, DiscAmtTB, DiscPerTP, DiscAmtTP, PackCrg, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, TurnOverPer, TurnOverAmt, ReSalTaxPer, ReSalTaxAmt, SROff, SatAmt
            If Index <> CaseNo Then
                If Val(txt(Index).TEXT) = 0 Then
                    txt(Index).TEXT = ""
                Else
                    txt(Index).TEXT = Format(txt(Index), "0.00")
                End If
            End If
    End Select
    'Removing Tag Value for SprCalc purpose
    Select Case Index
        Case DiscPerTB, DiscPerTP, GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ReSalTaxPer
            txt(Index).Tag = txt(Index)
    End Select
    'EOF Tag Value Removal
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
    Grid_Hide
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
        Case Col_PartSrlNo
            TxtGrid(Index).MaxLength = 20
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
'        Case Col_SONo
'            If RsSONo.RecordCount = 0 Or FGrid.TextMatrix(FGrid.Row, Col_SONo) = "" Then Exit Sub
'            If Txt(BaseDoc).Text = "" Then
'                Set RsSONo = New ADODB.Recordset
'                RsSONo.CursorLocation = adUseClient
'                RsSONo.Open "Select OrderID as Code,Trim(Mid(OrderID,9,5))+CStr(Trim(Right(OrderID,8))) as Name,V_Date,Qty,(Qty-Sup_Qty) As PendQty,Rate,Switch(TAX_YN=1,'Yes',TAX_YN=0,'No') As TAXYN,Switch(MRP_YN=1,'Yes',MRP_YN=0,'No') As MRPYN From SP_Order1 Where Order_Type='S_SO' and Part_No='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "' and Party_Code='" & Txt(Party).Tag & "' and (Qty-Sup_Qty) >0 Order By OrderID", GCn, adOpenDynamic, adLockOptimistic
'                Set DGSONo.DataSource = RsSONo
'            End If
'            If FGrid.TextMatrix(FGrid.Row, Col_SONoCode) <> RsSONo!Code Then
'                RsSONo.MoveFirst
'                RsSONo.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, Col_SONo) & "'"
'            End If
        Case Col_Godown
            If RsGodown.RecordCount = 0 Or FGrid.TextMatrix(FGrid.Row, Col_Godown) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Col_Godown) <> RsGodown!Name Then
                RsGodown.MoveFirst
                RsGodown.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, Col_Godown) & "'"
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        TxtGrid(0).TEXT = TxtGrid(0).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        FGrid.SetFocus
        TxtGrid(0).Visible = False
        Exit Sub
    End If
    Select Case FGrid.Col
        Case Col_PNo
            If DGPart.Visible = False Then DGridColSwap DGPart, 0
            DGridTxtKeyDown DGPart, TxtGrid, 0, RsPart, KeyCode, True, 0, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, 7
                End If
            End If
'        Case Col_SONo
'            DGridTxtKeyDown DGSONo, TxtGrid, 0, RsSONo, KeyCode, True, 1
'            If KeyCode = vbKeyReturn Then
'                If TxtGridLeave = True Then
'                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 21, 2
'                Else
'                    TxtGrid_LostFocus 0
'                    TxtGrid(0).SetFocus
'                End If
'            End If
        Case Col_Godown
            DGridTxtKeyDown DGGodown, TxtGrid, Index, RsGodown, KeyCode, True, 1, frmGodown, "frmGodown"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo
                End If
            End If
        Case Col_Taxable, Col_MRP, Col_Qty, Col_DiscPer, Col_TaxPer, Col_SatPer
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo
                End If
            End If
'        Case Col_Qty
'            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
'                If TxtGridLeave = True Then
'                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 21
'                Else
'                    TxtGrid_LostFocus 0
'                    TxtGrid(0).SetFocus
'                End If
'            End If
        Case Col_Rate
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, 2
                End If
            End If
'        Case Col_DiscPer
'            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
'                If TxtGridLeave = True Then
'                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 21
'                Else
'                    TxtGrid_LostFocus 0
'                    TxtGrid(0).SetFocus
'                End If
'            End If
        Case Col_DiscAmt
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                    If PubRestrict_Godown = 1 Then      ' Restrict Godown is "YES"
                        GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Col_PartSrlNo, 1
                    Else
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo
                    End If
                End If
            End If
        Case Col_PName
            If DGPart.Visible = False Then DGridColSwap DGPart, 1
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 1, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo
                End If
            End If
        Case Col_LName
            If DGPart.Visible = False Then DGridColSwap DGPart, 2
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 2, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo
                End If
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    CheckQuote KeyAscii
    Select Case FGrid.Col
        Case Col_PNo
            If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "Code"
        Case Col_PName
            If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "Name"
        Case Col_LName
            If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "LName"
'        Case Col_SONo
'            If DGSONo.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsSONo, KeyAscii, "Name"
        Case Col_Godown
            If DGGodown.Visible = True Then
                If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                    DGridTxtKeyPress TxtGrid, Index, RsGodown, KeyAscii, "Name"
                End If
            End If
        Case Col_Qty
            NumPress TxtGrid(Index), KeyAscii, 8, 3
        Case Col_Rate, Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer
            NumPress TxtGrid(Index), KeyAscii, 8, 2
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
'        Case Col_SONo
'            If KeyCode <> 13 And DGSONo.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsSONo, KeyCode, "Name", True
        Case Col_Godown
            If KeyCode <> 13 And DGGodown.Visible = False Then
                TxtGrid_KeyDown Index, GridKey, 0
                If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                    DGridTxtKeyPress TxtGrid, Index, RsGodown, KeyCode, "Name", True
                End If
            End If
        Case Col_Taxable, Col_MRP
            If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
                TxtGrid(Index) = ""
            ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Then
                TxtGrid(Index) = "Yes"
            Else
                TxtGrid(Index) = "No"
            End If
        Case Col_DiscPer, Col_DiscAmt, Col_Rate, Col_TaxPer, Col_SatPer
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        Case Col_Qty
            FGrid.TextMatrix(FGrid.Row, Col_Qty) = Format(Val(TxtGrid(Index).TEXT), "0.000")
            CountItem
        Case Col_PartSrlNo
            FGrid.TextMatrix(FGrid.Row, Col_PartSrlNo) = TxtGrid(Index)
    End Select
    Amt_Cal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_LostFocus(Index As Integer)
TxtGrid(Index).Visible = False
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
End Sub

Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    TxtGrid(0).Visible = False
    If TopCtrl1.TopText2 <> "Browse" Then
        MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, Col_PNo), _
            FGrid.TextMatrix(FGrid.Row, Col_PName), FGrid.TextMatrix(FGrid.Row, Col_LName), _
            Col_MRPStkTB, Col_MRPStkTP, _
            Col_TBStk, Col_TPStk, _
            Col_MRPRate, Col_TBRate, _
            Col_TPRate, Col_Bin, _
            Col_LastRate, Col_HPRate, Col_LPRate, mCheckNegetiveStockSiteWise
        FrmDetail.Visible = True
    End If
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
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
        If Trim(txt(BaseDoc).TEXT) = "" Then   ' If Base Doc. is not specified
            Select Case FGrid.Col
                Case Col_SONo
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    Amt_Cal
            End Select
        End If
        Select Case FGrid.Col
            Case Col_MRP, Col_Taxable
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            Case Col_Qty, Col_Rate, Col_Amt, Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                Amt_Cal
            Case Col_Godown
                If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                End If
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        If Trim(txt(BaseDoc).TEXT) = "" Then   ' If Base Doc. is not specified
            Select Case FGrid.Col
                Case Col_PNo, Col_PName, Col_LName
                    GridDblClick Me, FGrid, TxtGrid, 0
                Case Col_MRP, Col_Taxable ', Col_SONo
                    If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                        GridDblClick Me, FGrid, TxtGrid, 0
                    End If
            End Select
        End If
        Select Case FGrid.Col
            Case Col_Qty, Col_Rate, Col_DiscPer, Col_DiscAmt, Col_PartSrlNo, Col_TaxPer, Col_SatPer
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    GridDblClick Me, FGrid, TxtGrid, 0
                End If
            Case Col_Amt
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Col_PartSrlNo, , Col_DiscPer
                End If
            Case Col_ItemVal
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    If PubRestrict_Godown = 1 Then      ' Restrict Godown is "YES"
                        GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Col_PartSrlNo
                    End If
                End If
            Case Col_Godown
                If FGrid.TextMatrix(FGrid.Row, Col_Qty) <> "" Then
                    If PubRestrict_Godown = 0 Then      ' Restrict Godown is "Yes"
                        GridDblClick Me, FGrid, TxtGrid, 0
                    Else
                        GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Col_PartSrlNo
                    End If
                End If
        End Select
        TAddMode = False
    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
    If Trim(txt(BaseDoc).TEXT) = "" Then   ' If Base Doc. is not specified
        Select Case FGrid.Col
            Case Col_PNo, Col_PName, Col_LName   ', Col_SONo
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            Case Col_Unit
                FGrid.Col = FGrid.Col + 1
                FGrid.SetFocus
            Case Col_MRP, Col_Taxable
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
                End If
        End Select
    End If
    
    Select Case FGrid.Col
        Case Col_Qty, Col_Rate, Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
            End If
        Case Col_PartSrlNo
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            End If
        Case Col_Godown
            If FGrid.TextMatrix(FGrid.Row, Col_Qty) <> "" Then
                If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                    Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
                End If
            End If
    End Select
    If KeyAscii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
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
                   FGrid.TextMatrix(I, Col_SrNo) = I
                Next
                CountItem
                'Recalculate footer values
            If PubVATYN = 1 Then
               MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt1, txt(SatAmt)
            Else
                MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
            End If
                'MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
                    Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
                    Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
                    Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
                    Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
                    Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
                    Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
            End If
            MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, Col_PNo), _
                FGrid.TextMatrix(FGrid.Row, Col_PName), FGrid.TextMatrix(FGrid.Row, Col_LName), _
                Col_MRPStkTB, Col_MRPStkTP, _
                Col_TBStk, Col_TPStk, _
                Col_MRPRate, Col_TBRate, _
                Col_TPRate, Col_Bin, _
                Col_LastRate, Col_HPRate, Col_LPRate, mCheckNegetiveStockSiteWise
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_LostFocus()
FGrid.BackColorSel = BackColorSelLeave
FGrid.ForeColorSel = FGrid.ForeColor
If TopCtrl1.TopText2.CAPTION <> "Browse" Then
    If TxtGrid(0).Visible = False Then
        MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
            Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
            Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
            Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
        If PubVATYN = 1 Then
               MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt1, txt(SatAmt)
        Else
                MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, txt(IWDiscTotTB), txt(IWDiscTotTP), txt(MRPAmtTB), txt(MRPAmtTP), _
                    txt(SprAmtTB), txt(SprAmtTP), txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), _
                    txt(DiscPerTP), txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), txt(STotATP), _
                    txt(GenSurPer), txt(GenSurAmt), txt(TransAmt), txt(TaxableTot), _
                    txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt), txt(PackCrg), _
                    txt(STotB), txt(TurnOverPer), txt(TurnOverAmt), txt(ReSalTaxPer), txt(ReSalTaxAmt), _
                    txt(SROff), txt(NetSprAmt), txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
        End If
        'MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
            Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
            Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
            Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
            Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
            Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
            Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
    End If
End If
If FrmDetail.Visible = True Then FrmDetail.Visible = False
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
End Sub

Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
'    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub TxtGridValid_PNo()
'Called from TxtGrid_Validate & TxtGridLeave procedures
Dim OldPNo$
Dim RsTemp As ADODB.Recordset

If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Or TxtGrid(0).TEXT = "" Then
    FGrid.TextMatrix(FGrid.Row, Col_PNo) = ""
    FGrid.TextMatrix(FGrid.Row, Col_PName) = ""
    FGrid.TextMatrix(FGrid.Row, Col_LName) = ""
    MainLib.Fill_Data mPartyType, LblFrm, FGrid, _
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
    
    MainLib.Fill_Data mPartyType, LblFrm, FGrid, _
        RsPart!Code, RsPart!Name, RsPart!LName, _
        Col_Unit, Col_MRP, Col_Taxable, Col_MRPStkTB, Col_MRPStkTP, _
        Col_TBStk, Col_TPStk, _
        Col_MRPRate, Col_TBRate, _
        Col_TPRate, Col_Bin, _
        Col_HPRate, Col_LPRate, _
        Col_LastRate, Col_PartGrade, _
        Col_EffectDate, Col_DiscPer, mCheckNegetiveStockSiteWise
'by LPS 27-04-2K2
    If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> OldPNo Then
            FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(mPartyType, FGrid, CDate(txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
'            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsPart!SalDisc_Per, "0.00")
        End If
    End If
    If PubVATYN = 1 Then
         Set rsTaxPer = GCn.Execute("Select Tax_Per, AddTaxPer, L_C from TaxForms where Form_Code='" & txt(FormName).Tag & "'")
         If rsTaxPer.RecordCount > 0 Then
            FGrid.TextMatrix(FGrid.Row, Col_TaxPer) = VNull(rsTaxPer!Tax_Per)
            FGrid.TextMatrix(FGrid.Row, Col_SatPer) = VNull(rsTaxPer!AddTaxPer)
         End If
         
        If UTrim(XNull(rsTaxPer!L_C)) = "LOCAL" Then
           Set RsTemp = GCn.Execute("Select VatPer, AddTaxPer From Part_Grade Where PartGrade_Code = '" & FGrid.TextMatrix(FGrid.Row, Col_PartGrade) & "'")
           If RsTemp.RecordCount > 0 Then
               If VNull(RsTemp!VatPer) > 0 Then FGrid.TextMatrix(FGrid.Row, Col_TaxPer) = Format(VNull(RsTemp!VatPer), "0.00")
               If VNull(RsTemp!AddTaxPer) > 0 Then FGrid.TextMatrix(FGrid.Row, Col_SatPer) = Format(VNull(RsTemp!AddTaxPer), "0.00")
           End If
        End If
         
    End If
End If
    If Trim(txt(BaseDoc).TEXT) = "" Then    ' In Case No Base Doc. Specified
'        Set RsSONo = New ADODB.Recordset
'        RsSONo.CursorLocation = adUseClient
'        RsSONo.Open "Select OrderID as Code,Trim(Mid(OrderID,9,5))+CStr(Trim(Right(OrderID,8))) as Name,V_Date,Qty,Rate,Switch(TAX_YN=1,'Yes',TAX_YN=0,'No') As TAXYN,Switch(MRP_YN=1,'Yes',MRP_YN=0,'No') As MRPYN From SP_Order1 Where Order_Type='S_SO' and Part_No='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "' and Party_Code='" & Txt(Party).Tag & "' Order By OrderID", GCn, adOpenDynamic, adLockOptimistic
'        Set DGSONo.DataSource = RsSONo
    End If
    If FGrid.TextMatrix(FGrid.Rows - 1, Col_PNo) <> "" Then FGrid.AddItem FGrid.Rows
End Sub

'Private Sub TxtGridValid_SONo()
'    If RsSONo.RecordCount = 0 Or (RsSONo.EOF = True Or RsSONo.BOF = True) Or TxtGrid(Index).Text = "" Then
'        FGrid.TextMatrix(FGrid.Row, Col_SONo) = ""
'        FGrid.TextMatrix(FGrid.Row, Col_SONoCode) = ""
'    Else
'        FGrid.TextMatrix(FGrid.Row, Col_SONoCode) = RsSONo!Code
'        FGrid.TextMatrix(FGrid.Row, Col_SONo) = RsSONo!Name
'    End If
'End Sub

Private Sub TxtGridValid_TaxMRP()
FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
'        If TopCtrl1.TopText2 = "Add" Or _
            TopCtrl1.TopText2 = "Edit" And Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) = 0 Then
            FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(mPartyType, FGrid, CDate(txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
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

Private Sub ProcAcPost(rsCtrlAc As ADODB.Recordset)
On Error GoTo lblExit
Dim xMRPSprTp As Double, xMRPOilTp As Double
Dim xSprTp As Double, xOilTp As Double
Dim mShare As Single, mShareAmt As Double, mShare2Amt As Double
Dim xNetAmt As Double, xRoundAmt As Double, xSprAmtMRPTB As Double, xSprAmtMRPTP As Double
Dim xOilAmtMRPTB As Double, xOilAmtMRPTP As Double
Dim xSprAmtTB  As Double, xSprAmtTP As Double, xOilAmtTB As Double, xOilAmtTP As Double
Dim xDisAmtTB As Double, xDisAmtTP As Double, xDisAmtMRPTB As Double, xDisAmtMRPTP As Double
Dim xGenSurAmt As Double, xTrans As Double, xTaxAmt As Double, xTaxAmtMRP As Double, xPack As Double
Dim xTurnOver As Double, xReSaleTaxAmt As Double, mFADocID$, mQry$
Dim RsTemp As ADODB.Recordset, rsTemp1 As ADODB.Recordset
'A/c Posting related declarations
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, mNarr$, TaxSQL$, I As Integer, j As Integer
Dim mSprAmtMRPTB As Double, mSprAmtTB As Double
Dim mOilAmtMRPTB As Double, mOilAmtTB As Double
Dim mTotMRPOilTB As Double, mTotOilTB As Double, mTotShareAmt As Double
Dim mShareSpr As Single, mShareAmtSpr As Double, mShare2AmtSpr As Double
Dim mTot1ShareAmt As Double, mTot2ShareAmt As Double, mTot3ShareAmt As Double
Dim PartyCode$

    TaxSQL = "select TF.Tax_Ac_Code,TF.Sur_Ac_Code, Tf.AddTaxAc,sum(Tax_Amt+Tax_AmtMRP) as TaxAmt,sum(Tax_Sur_Amt+TaxSur_AmtMRP) as TaxSurAmt, Sum(SatAmt) As SatAmt " & _
        " from SP_Sale left join TaxFormsAc as TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code"
    
    If txt(DocType) = "Sale Return Cash" Then
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
            "where V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(DocID), 8) & _
            "' group by TF.PurSal_Ac_Code"
            
        mQry = "select " & _
            "sum(Total_Amt) as NetAmt,sum(rounded) as RoundAmt," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(SprAmt_MRP_TP) as SprAmtMRPTP, " & _
            "sum(OilAmt_MRP_TB) as OilAmtMRPTB, sum(OilAmt_MRP_TP) as OilAmtMRPTP, " & _
            "sum(SprAmt_TB) as SprAmtTB, sum(SprAmt_TP) as SprAmtTP, " & _
            "sum(OilAmt_TB) as OilAmtTB, sum(OilAmt_TP) as OilAmtTP, " & _
            "sum(D_Amt_TB) as DisAmtTB, sum(D_Amt_TP) as DisAmtTP, " & _
            "sum(D_Amt_MRP_TB) as DisAmtMRPTB, sum(D_Amt_MRP_TP) as DisAmtMRPTP," & _
            "sum(Gen_Sur_Amt) as GenSurAmt,sum(Trans_Amt) as Trans," & _
            "sum(Tax_Amt+Tax_Sur_Amt+Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmt," & _
            "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, sum(TOT_Amt) as TurnOver, " & _
            "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
            "where V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(DocID), 8) & "'"
        'for tax
        TaxSQL = TaxSQL & " where  V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(DocID), 8) & _
            "' Group by TF.Tax_Ac_Code,TF.Sur_Ac_Code, TF.AddTaxAc"
        mNarr = "Through Counter Cash Sale Return (Daily Posting)"
        mCommNarr = mNarr & " [Common]"
        mFADocID = left(txt(DocID), 8) & "AAAAA" & "  " & Format(txt(VDate), "yymmdd")
        PartyCode = PubSprCashAc
    Else
        PartyCode = txt(Party).Tag
        mFADocID = txt(DocID)
        If txt(DocType) = "Sale Return Credit" Then
            mNarr = "Through Counter Cr Sale Return "
        Else
            mNarr = "Through Transfer Return "
        End If
        mCommNarr = mNarr & " [Common]"
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
            "where docid='" & txt(DocID) & _
            "' group by TF.PurSal_Ac_Code"

        mQry = "select " & _
            "sum(Total_Amt) as NetAmt,sum(rounded) as RoundAmt," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(SprAmt_MRP_TP) as SprAmtMRPTP, " & _
            "sum(OilAmt_MRP_TB) as OilAmtMRPTB, sum(OilAmt_MRP_TP) as OilAmtMRPTP, " & _
            "sum(SprAmt_TB) as SprAmtTB, sum(SprAmt_TP) as SprAmtTP, " & _
            "sum(OilAmt_TB) as OilAmtTB, sum(OilAmt_TP) as OilAmtTP, " & _
            "sum(D_Amt_TB) as DisAmtTB, sum(D_Amt_TP) as DisAmtTP, " & _
            "sum(D_Amt_MRP_TB) as DisAmtMRPTB, sum(D_Amt_MRP_TP) as DisAmtMRPTP," & _
            "sum(Gen_Sur_Amt) as GenSurAmt,sum(Trans_Amt) as Trans," & _
            "sum(Tax_Amt+Tax_Sur_Amt+Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmt," & _
            "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, sum(TOT_Amt) as TurnOver, " & _
            "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
            "where docid='" & txt(DocID) & "'"
        'for tax
        TaxSQL = TaxSQL & " where docid='" & txt(DocID) & _
            "' Group by TF.Tax_Ac_Code,TF.Sur_Ac_Code, TF.AddTaxAc"
    End If
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open mQry, GCn, adOpenStatic, adLockReadOnly
    
    'for tax purpose
    Set RsTemp = New ADODB.Recordset
    RsTemp.CursorLocation = adUseClient
    RsTemp.Open TaxSQL, GCn, adOpenStatic, adLockReadOnly

'        1.MRP Spr TB = SprAmtMRPTB-OilAmtMRPTB - part of DisAmtMRPTB - part of TaxAmtMRP
'        3.MRP Oil TB = OilAmtMRPTB - part of DisAmtMRPTB - part of TaxAmtMRP

'        2.MRP Spr TP = SprAmtMRPTP-OilAmtMRPTP - part of DisAmtMRPTP
'        4.MRP Oil TP = OilAmtMRPTP - part of DisAmtMRPTP

'        1.Spr TB = SprAmtTB - part of DisAmtTB
'        3.Oil TB = OilAmtTB - part of DisAmtTB
'        2.Spr TP = SprAmtTP - part of DisAmtTP
'        4.Oil TP = OilAmtTP - part of DisAmtTP
    xNetAmt = IIf(IsNull(rsTemp1!NetAmt), 0, rsTemp1!NetAmt): xRoundAmt = IIf(IsNull(rsTemp1!RoundAmt), 0, rsTemp1!RoundAmt)
    xSprAmtMRPTB = IIf(IsNull(rsTemp1!SprAmtMrpTB), 0, rsTemp1!SprAmtMrpTB)
    xOilAmtMRPTB = IIf(IsNull(rsTemp1!OilAmtMrpTB), 0, rsTemp1!OilAmtMrpTB)
    xSprAmtMRPTP = IIf(IsNull(rsTemp1!SprAmtMrpTP), 0, rsTemp1!SprAmtMrpTP)
    xOilAmtMRPTP = IIf(IsNull(rsTemp1!OilAmtMrpTP), 0, rsTemp1!OilAmtMrpTP)
    xSprAmtTB = IIf(IsNull(rsTemp1!SprAmtTB), 0, rsTemp1!SprAmtTB)
    xOilAmtTB = IIf(IsNull(rsTemp1!OilAmtTB), 0, rsTemp1!OilAmtTB)
    xSprAmtTP = IIf(IsNull(rsTemp1!SprAmtTP), 0, rsTemp1!SprAmtTP)
    xOilAmtTP = IIf(IsNull(rsTemp1!OilAmtTP), 0, rsTemp1!OilAmtTP)
    xDisAmtTB = IIf(IsNull(rsTemp1!DisAmtTB), 0, rsTemp1!DisAmtTB)
    xDisAmtTP = IIf(IsNull(rsTemp1!DisAmtTP), 0, rsTemp1!DisAmtTP)
    xDisAmtMRPTB = IIf(IsNull(rsTemp1!DisAmtMRPTB), 0, rsTemp1!DisAmtMRPTB)
    xDisAmtMRPTP = IIf(IsNull(rsTemp1!DisAmtMRPTP), 0, rsTemp1!DisAmtMRPTP)
    xGenSurAmt = IIf(IsNull(rsTemp1!GenSurAmt), 0, rsTemp1!GenSurAmt)
    xTrans = IIf(IsNull(rsTemp1!Trans), 0, rsTemp1!Trans)
    xTaxAmt = IIf(IsNull(rsTemp1!TaxAmt), 0, rsTemp1!TaxAmt)
    xTaxAmtMRP = IIf(IsNull(rsTemp1!TaxAmtMRP), 0, rsTemp1!TaxAmtMRP)
    xPack = IIf(IsNull(rsTemp1!Pack), 0, rsTemp1!Pack)
    xTurnOver = IIf(IsNull(rsTemp1!TurnOver), 0, rsTemp1!TurnOver)
    xReSaleTaxAmt = IIf(IsNull(rsTemp1!ReSaleTaxAmt), 0, rsTemp1!ReSaleTaxAmt)
    '*** Sale Amount Row
    I = 1
    ReDim Preserve LedgAry(1)
    '**Taxable Spr / Oil Calculation
     Do While GRs.EOF = False
        mOilAmtMRPTB = IIf(IsNull(GRs!OilAmtMrpTB), 0, GRs!OilAmtMrpTB)
        mSprAmtMRPTB = IIf(IsNull(GRs!SprAmtMrpTB), 0, GRs!SprAmtMrpTB) ' - mOilAmtMRPTB
        mSprAmtTB = IIf(IsNull(GRs!SprAmtTB), 0, GRs!SprAmtTB)
        mOilAmtTB = IIf(IsNull(GRs!OilAmtTB), 0, GRs!OilAmtTB)
        'Allocate values in their proportions
        If (mSprAmtMRPTB + mOilAmtMRPTB) <> 0 Then
            mShare = Round((mSprAmtMRPTB + mOilAmtMRPTB) * 100 / (xSprAmtMRPTB + xOilAmtMRPTB), 2)
            mShareAmt = Round(xDisAmtMRPTB * mShare / 100, 2)
            mShare2Amt = Round(xTaxAmtMRP * mShare / 100, 2)
            mShareSpr = Round(mSprAmtMRPTB * 100 / (mSprAmtMRPTB + mOilAmtMRPTB), 2)
            mShareAmtSpr = Round(mShareAmt * mShareSpr / 100, 2)
            mShare2AmtSpr = Round(mShare2Amt * mShareSpr / 100, 2)
            mTot1ShareAmt = mTot1ShareAmt + mShareAmt
            mTot2ShareAmt = mTot2ShareAmt + mShare2Amt
            If GRs.AbsolutePosition = GRs.RecordCount Then
                mShareAmt = mShareAmt + ((xDisAmtMRPTB) - mTot1ShareAmt)
                mShare2Amt = mShare2Amt + ((xTaxAmtMRP) - mTot2ShareAmt)
            End If
            mSprAmtMRPTB = mSprAmtMRPTB - (mShareAmtSpr + mShare2AmtSpr)
            mOilAmtMRPTB = mOilAmtMRPTB - ((mShareAmt + mShare2Amt) - (mShareAmtSpr + mShare2AmtSpr))
        End If
        '*****
        If (mSprAmtTB + mOilAmtTB) <> 0 Then
            mShare = Round((mSprAmtTB + mOilAmtTB) * 100 / (xSprAmtTB + xOilAmtTB), 2)
            mShareAmt = Round((xDisAmtTB - xDisAmtMRPTB) * mShare / 100, 2)
            mShareSpr = Round(mSprAmtTB * 100 / (mSprAmtTB + mOilAmtTB), 2)
            mShareAmtSpr = Round(mShareAmt * mShareSpr / 100, 2)
            mTot3ShareAmt = mTot3ShareAmt + mShareAmt
            If GRs.AbsolutePosition = GRs.RecordCount Then
                mShareAmt = mShareAmt + ((xDisAmtTB - xDisAmtMRPTB) - mTot3ShareAmt)
            End If
            mSprAmtTB = mSprAmtTB - (mShareAmtSpr)
            mOilAmtTB = mOilAmtTB - (mShareAmt - mShareAmtSpr)
        End If
        'Spare Sale A/c Taxable
        mTotMRPOilTB = mTotMRPOilTB + mOilAmtMRPTB
        mTotOilTB = mTotOilTB + mOilAmtTB
        '*****
        If mSprAmtMRPTB + mSprAmtTB <> 0 Then
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = GRs!PurSal_Ac_Code
            LedgAry(I).AmtDr = Round(mSprAmtMRPTB + mSprAmtTB, 2)
            LedgAry(I).AmtCr = 0
            LedgAry(I).Narration = mNarr ' & " Spare"
        End If
        GRs.MoveNext
    Loop
    If (xSprAmtMRPTP + xOilAmtMRPTP) <> 0 Then
        xMRPSprTp = xSprAmtMRPTP '- xOilAmtMRPTP
        xMRPOilTp = xOilAmtMRPTP
        mShare = Round(xMRPSprTp * 100 / (xSprAmtMRPTP + xOilAmtMRPTP), 2)
        mShareAmt = Round(xDisAmtMRPTP * mShare / 100, 2)
        xMRPSprTp = xMRPSprTp - (mShareAmt)
        xMRPOilTp = xMRPOilTp - (xDisAmtMRPTP - (mShareAmt))
    End If
    If (xSprAmtTP + xOilAmtTP) <> 0 Then
        mShare = Round(xSprAmtTP * 100 / (xSprAmtTP + xOilAmtTP), 2)
        mShareAmt = Round((xDisAmtTP - xDisAmtMRPTP) * mShare / 100, 2)
        xSprTp = xSprAmtTP - (mShareAmt)
        xOilTp = xOilAmtTP - ((xDisAmtTP - xDisAmtMRPTP) - (mShareAmt))
    End If

'   0.Taxable Spr = MRP Spr TB + SPR TB = Dr
'   1.Taxpaid Spr = MRP Spr TP + SPR TP = Dr
'   2.Taxable Oil = MRP Oil TB + Oil TB = Dr
'   3.Taxable Oil = MRP Oil TP + Oil TP = Dr
'   4.xGenSurAmt = Dr
'   5.xPack = Dr
'   6.xTurnOver = Dr
'   7.xReSaleTaxAmt = Dr
'   8.Party A/c or Cash A/c = Cr
    'Spare Sale A/c Taxpaid
    If xMRPSprTp + xSprTp <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!SprSalTP_Ac
        LedgAry(I).AmtDr = Round(xMRPSprTp + xSprTp, 2)
        LedgAry(I).AmtCr = 0
        LedgAry(I).Narration = mNarr ' & " Spare"
    End If
    'Oil Sale A/c Taxable
    If mTotMRPOilTB + mTotOilTB <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!OilSalTB_Ac
        LedgAry(I).AmtDr = Round(mTotMRPOilTB + mTotOilTB, 2)
        LedgAry(I).AmtCr = 0
        LedgAry(I).Narration = mNarr ' & " Spare"
    End If
     'Oil Sale A/c Taxpaid
     If xMRPOilTp + xOilTp <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!OilSalTP_Ac
        LedgAry(I).AmtDr = Round(xMRPOilTp + xOilTp, 2)
        LedgAry(I).AmtCr = 0
        LedgAry(I).Narration = mNarr ' & " Spare"
     End If
      'GenSurAmt
     If xGenSurAmt <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!SprGenSur_Ac
        LedgAry(I).AmtDr = Round(xGenSurAmt, 2)
        LedgAry(I).AmtCr = 0
        LedgAry(I).Narration = mNarr ' & " Sale Tax"
     End If
    'Transportation
     If xTrans <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!Transportation_Ac
        If xTrans > 0 Then
            LedgAry(I).AmtDr = Round(xTrans, 2)
            LedgAry(I).AmtCr = 0
        Else
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(Abs(xTrans), 2)
        End If
        LedgAry(I).Narration = mNarr '& " Transportation"
     End If
     If RsTemp.RecordCount > 0 Then
         Do While RsTemp.EOF = False
            If RsTemp!TaxAmt <> 0 Then
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
                LedgAry(I).SubCode = RsTemp!Tax_Ac_Code
                If RsTemp!TaxAmt > 0 Then
                    LedgAry(I).AmtDr = Round(RsTemp!TaxAmt, 2)
                    LedgAry(I).AmtCr = 0
                Else
                    LedgAry(I).AmtDr = 0
                    LedgAry(I).AmtCr = Round(Abs(RsTemp!TaxAmt), 2)
                End If
                 LedgAry(I).Narration = mNarr '& " Sales Tax & Surcharge"
            End If
            If RsTemp!TaxSurAmt <> 0 Then
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
                LedgAry(I).SubCode = RsTemp!Sur_Ac_Code
                If RsTemp!TaxSurAmt > 0 Then
                    LedgAry(I).AmtDr = Round(RsTemp!TaxSurAmt, 2)
                    LedgAry(I).AmtCr = 0
                Else
                    LedgAry(I).AmtDr = 0
                    LedgAry(I).AmtCr = Round(Abs(RsTemp!TaxSurAmt), 2)
                End If
                 LedgAry(I).Narration = mNarr '& " Sales Tax & Surcharge"
             End If
            If RsTemp!SatAmt <> 0 Then
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
                LedgAry(I).SubCode = RsTemp!AddTaxAc
                If RsTemp!TaxAmt > 0 Then
                    LedgAry(I).AmtDr = Round(RsTemp!SatAmt, 2)
                    LedgAry(I).AmtCr = 0
                Else
                    LedgAry(I).AmtDr = 0
                    LedgAry(I).AmtCr = Round(Abs(RsTemp!SatAmt), 2)
                End If
                 LedgAry(I).Narration = mNarr '& " Sales Tax & Surcharge"
            End If
             
             
             RsTemp.MoveNext
         Loop
     End If
    'Misc / Packing Chrg
    If xPack <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!MiscChrg_Ac
        If Val(xPack) > 0 Then
            LedgAry(I).AmtDr = Round(xPack, 2)
            LedgAry(I).AmtCr = 0
        Else
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(Abs(xPack), 2)
        End If
        LedgAry(I).Narration = mNarr '& " Misc Charges"
    End If
    If xTurnOver <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!TOTax_Ac
        If xTurnOver > 0 Then
            LedgAry(I).AmtDr = Round(xTurnOver, 2)
            LedgAry(I).AmtCr = 0
        Else
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(Abs(xTurnOver), 2)
        End If
        LedgAry(I).Narration = mNarr '& " TurnOver Amt"
    End If
    If xReSaleTaxAmt <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!ReSaleTax_Ac
        If xReSaleTaxAmt > 0 Then
            LedgAry(I).AmtDr = Round(xReSaleTaxAmt, 2)
            LedgAry(I).AmtCr = 0
        Else
            LedgAry(I).AmtCr = Round(Abs(xReSaleTaxAmt), 2)
            LedgAry(I).AmtDr = 0
        End If
        LedgAry(I).Narration = mNarr '& " ReSale Tax Amount"
    End If
    If xRoundAmt <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!SprROff_Ac
        If xRoundAmt > 0 Then
            LedgAry(I).AmtDr = Round(xRoundAmt, 2)
            LedgAry(I).AmtCr = 0
        Else
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(Abs(xRoundAmt), 2)
        End If
        LedgAry(I).Narration = mNarr '& " Round Off"
    End If
    'Cash / Party A/c
    I = UBound(LedgAry) + 1
    ReDim Preserve LedgAry(I)
    LedgAry(I).SubCode = PartyCode
    LedgAry(I).AmtDr = 0
    LedgAry(I).AmtCr = Round(xNetAmt, 2)
    LedgAry(I).Narration = mNarr
    
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, mFADocID, CDate(txt(VDate)), mCommNarr)
    If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
lblExit:
    Set GRs = Nothing
    Set RsTemp = Nothing
    If err.NUMBER <> 0 Then MsgBox err.Description
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
            MoveRec
        End If
    End If
End Sub
Private Sub CmdPrint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
GSQL = "SELECT Syctrl.SprInvFooter,s.DocID,s.V_Type,s.V_No,s.V_Date,s.Cash_Credit,s.Job_DocID,s.Party_Code,s.Party_Name,s.Address," & _
    "SG.NamePrefix,SG.Name,SG.Add1, SG.Add2, SG.Add3,City.CityName,SG.PIN,SG.Phone,SG.CSTNo,s.L_C,s.REP_CODE, s.Form_Code,s.RoadPermit_FormCode," & _
    "s.GR_RR_No, s.GR_RR_Date,s.CrAc, s.Case_No, s.Case_Mark, s.Mode_Dispatch, s.Transport, s.Remarks,s.SprAmt_MRP_TB,s.SprAmt_MRP_TP," & _
    "s.OilAmt_MRP_TB,s.OilAmt_MRP_TP,s.SprAmt_TB,s.SprAmt_TP, s.OilAmt_TB, s.OilAmt_TP, s.D_Per_TB, s.D_Amt_TB, s.D_Per_TP,s.D_Amt_TP," & _
    "s.D_Per_MRP_TB,s.D_Amt_MRP_TB,s.D_Per_MRP_TP,s.D_Amt_MRP_TP,s.Addition, s.Gen_Sur_Per, s.Gen_Sur_Amt, s.Trans_Amt, s.LineFileTaxSum," & _
    "s.Tax_Per, s.Tax_Amt, s.Tax_AmtMRP, s.Tax_Sur_Per,s.Tax_Sur_Amt,s.TaxSur_AmtMRP,s.Packing, s.TOT_Per, s.Tot_Amt, s.TOT_AmtMRP," & _
    "s.ReSalTax_Per, s.ReSalTax_Amt,s.Total_Amt, s.Rounded, s.Det_Tax,s.GP_No,s.GP_Date,s.Printed_YN,s.Invoice_DocId, s.U_Name, s.U_EntDt,S.CancelYN," & _
    "" & vIsNull("SPStk.Srl_No", "0") & " as Srl_No, " & xIsNull("SPStk.V_Date", "") & " as SPStk_V_Date, " & xIsNull("SPStk.Party_Code", "") & " as SPStkParty_Code,SPStk.L_C," & _
    "" & xIsNull("SPStk.Job_DocID", "") & " as Job_DocID,SPStk.Mech_Code, SPStk.Order_DocId,SPStk.Order_Srl_No,SPStk.Part_No,Part.Part_Name, SPStk.Lub_Category, SPStk.Godown," & _
    "" & vIsNull("SPStk.Qty_Doc", "0") & " as Qty_Doc, " & vIsNull("SPStk.Qty_Rec", "0") & " as Qty_Rec, " & vIsNull("SPStk.Qty_Iss", "0") & " as Qty_Iss," & _
    "" & vIsNull("SPStk.Qty_Ret", "0") & " as Qty_Ret, " & vIsNull("SPStk.Tax_YN", "0") & " as Tax_YN, " & vIsNull("SPStk.MRP_YN", "0") & " as MRP_YN," & _
    "" & vIsNull("SPStk.Rate", "0") & " as Rate, " & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate, " & vIsNull("SPStk.Disc_Per", "0") & " as Disc_Per," & _
    "" & vIsNull("SPStk.Disc_Amt", "0") & " as Disc_Amt, " & vIsNull("SPStk.AMOUNT", "0") & " as AMOUNT, " & vIsNull("SPStk.Ord_DiscPer", "0") & " as Ord_DiscPer," & _
    "" & vIsNull("SPStk.Ord_DiscAmt", "0") & " as Ord_DiscAmt, " & vIsNull("SPStk.Net_Amt", "0") & " as Net_Amt, " & xIsNull("SPStk.Purpose", "") & " as Purpose," & _
    "SPStk.Part_SrlNo,SPStk.Remark,SPStk.Invoice_DocId as SPStk_Invoice_DocId, SPStk.V_Date2, " & vIsNull("SPStk.Rate2", "0") & " as Rate2, " & vIsNull("SPStk.MRP_Rate2", "0") & " as MRP_Rate2," & _
    "" & vIsNull("SPStk.Disc_Per2", "0") & " as Disc_Per2, " & vIsNull("SPStk.Disc_Amt2", "0") & " as Disc_Amt2, " & vIsNull("SPStk.Amount2", "0") & " as Amount2," & _
    "" & vIsNull("SPStk.Ord_DiscPer2", "0") & " as Ord_DiscPer2, " & vIsNull("SPStk.Ord_DiscAmt2", "0") & " as Ord_DiscAmt2, " & vIsNull("SPStk.Net_Amt2", "0") & " as Net_Amt2,SPStk.Printed2,TF.Printing_Desc, S.SatAmt " & _
    " FROM ((((SP_Sale as S left JOIN SP_Stock as SPStk ON S.DocID = SPStk.DocId) " & _
    "left JOIN Part ON SPStk.Part_No = Part.PART_NO and Part.Div_Code = left(SPStk.Docid,1)) " & _
    "LEFT JOIN (SubGroup as SG LEFT JOIN City ON SG.CityCode = City.CityCode) ON S.Party_Code = SG.SubCode) " & _
    "Left Join TaxForms TF on S.Form_Code=TF.Form_Code) " & _
    "LEFT JOIN Syctrl ON Syctrl.LinkTable<>S.U_AE " & _
    "where S.DocId='" & Master!SearchCode & "'"
Select Case Index
    Case PScreen, PWindows
'        mRepName = IIf(OptPlain.Value = True, "SprSaleChlRet", "SprSaleChlRet")
        mRepName = IIf(OptPlain.Value = True, "SprSaleRet", "SprSaleRet")
        Call WindowsPrint(Index, GSQL)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint(GSQL)
        FrmPrn.Visible = False
    Case PSetUp
'        mRepName = IIf(OptPlain.Value = True, "SprSaleChlRet", "SprSaleChlRet")
        mRepName = IIf(OptPlain.Value = True, "SprSaleRet", "SprSaleRet")
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
    MoveRec
End If
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub WindowsPrint(Index As Integer, mQry As String)
Dim RST1 As ADODB.Recordset
Dim RstRep As ADODB.Recordset
'Dim mQRY$
Dim I As Integer, j As Integer
On Error GoTo ERRORHANDLER
'        mQRY = "SELECT City.CityName, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, SubGroup.PIN, SubGroup.Phone, Syctrl.SprInvFooter, Syctrl.SprRetInvFooter, Part.Part_Name, Estimate1.*, Estimate.* " & _
        "FROM (((SP_Sale as Estimate left JOIN SP_Stock as Estimate1 ON Estimate.DocID = Estimate1.DocId) left JOIN Part ON Estimate1.Part_No = Part.PART_NO and Part.Div_Code = left(Estimate1.DocID,1)) LEFT JOIN (SubGroup LEFT JOIN City ON SubGroup.CityCode = City.CityCode) ON Estimate.Party_Code = SubGroup.SubCode) LEFT JOIN Syctrl ON Syctrl.LinkTable >=Estimate.U_AE " & _
        "where Estimate.DocId='" & Master!SearchCode & "'"
        
        Set RstRep = GCn.Execute(mQry)
        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
        
        Set RST1 = GCn.Execute("select S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
               
        CreateFieldDefFile RstRep, PubRepoPath + "\" & mRepName & ".ttx", True
        If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("LST")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecLST & "'"
                Case UCase("LSTDate")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecLST_Date & "'"
                Case UCase("CST")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecCST & "'"
                Case UCase("CSTDate")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecCST_Date & "'"
                Case UCase("Phone")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecPhone & "'"
                Case UCase("Fax")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!S_SecFax & "'"
                Case UCase("VouType")
                    rpt.FormulaFields(I).TEXT = "'SYSIR'"
                Case UCase("TOTCaption")
                    rpt.FormulaFields(I).TEXT = "'" & pubTOTCaption & "'"
            End Select
        Next
        rpt.Database.SetDataSource RstRep
        rpt.ReadRecords
        Select Case Index
            Case 0  'Printer
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
                            rpt.FormulaFields(I).TEXT = "'" & IIf(mVType = RetTrfIssVType, "TRANSFER ISSUE RETURN", "SALE RETURN") & "'"
                    End Select
                Next
                rpt.PrintOut False
                If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
                    GCn.Execute "update Sp_Sale set Printed_YN = 1  where Sp_Sale.docid='" & Master!SearchCode & "' "
                End If
            Case 1  'screen
                Call Report_View(rpt, IIf(mVType = RetTrfIssVType, "TRANSFER ISSUE RETURN", "SALE RETURN"), , True)
        End Select
        Set RST1 = Nothing
        Set RstRep = Nothing
        CmdPrint(PSetUp).Tag = ""
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
    Dim PrintStr As String
    Dim RstCompDet As ADODB.Recordset, RstRep As ADODB.Recordset, RstStock As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mGoods_Amt As Double
    Dim Footer As String, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select SprRetInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
'   mGatePass = 0
    mFooter = 19    'Line For Gate Pass =9 ,Line For NonTax Detail = 5
    mFooter = mFooter + FooterCnt
    mDocStr = IIf(mVType = RetTrfIssVType, "TRANSFER ISSUE RETURN", "SALE RETURN")
    mDocStr = mDocStr & IIf(mVType = RetCashSalVType, " CASH", "")
    mDocStr = mDocStr & IIf(mVType = RetCrSalVType, " CREDIT", "")
    mDupStr = IIf(RstRep!Printed_YN = 1, "(DUPLICATE)", "")
    Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")

    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!S_SecSpeciality) <> "" Then
        Print #1, PRN_TIT(RstCompDet!S_SecSpeciality, "C", PageWidth)
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
    Print #1, PSTR(XNull(RstCompDet!S_SecLST) & IIf(XNull(RstCompDet!S_SecLST_Date) = "", "", " Dt. " & RstCompDet!S_SecLST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!S_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!S_SecPhone)), 27, , AlignRight, " ")
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstCompDet!S_SecCST) & IIf(XNull(RstCompDet!S_SecCST_Date) = "", "", " Dt. " & RstCompDet!S_SecCST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!S_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!S_SecFax)), 27, , AlignRight, " ")
    mHeader = mHeader + 1
    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, mChr18 & "To," & mEmph
    mHeader = mHeader + 1
    Print #1, PSTR(RstRep!NamePrefix & " " & RstRep!Party_Name, 40) & Space(1) & PSTR(mDocStr & " NO.", Len(mDocStr) + 5) & " : " & PrinID(RstRep!DocID) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(IIf(mVType <> RetCashSalVType, XNull(RstRep!Add1), XNull(RstRep!Address)), 40) & Space(1) & mEmph & PSTR(mDocStr & " DATE", Len(mDocStr) + 5) & " : " & CDate(RstRep!V_Date) & mEmph1
    mHeader = mHeader + 1
    Print #1, IIf(mVType <> RetCashSalVType, XNull(RstRep!Add2), "") & Space(1) & mEmph & IIf(RstRep!CancelYN = 1, "** CANCELLED **", "") & mEmph1
    mHeader = mHeader + 1
    Print #1, IIf(mVType <> RetCashSalVType, XNull(RstRep!Add3) & IIf(XNull(RstRep!Add3) <> "" And XNull(RstRep!CityName) <> "", ",", "") & XNull(RstRep!CityName), "")
    mHeader = mHeader + 1
    Print #1, mDoub & "CST NO." & XNull(RstRep!CstNo) & mDoub1
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-") & mChr17 & mDoub
    mHeader = mHeader + 1
    Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 35) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
    mHeader = mHeader + 1
    Print #1, Space(89) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mDoub1 & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
    mHeader = mHeader + 1
    mFix = PageLength - (mHeader + mFooter)
    Page = 1
    mLine = 1
    mSlNo = 1
    LAdd = VNull(RstRep!Gen_Sur_Amt) + VNull(RstRep!Trans_Amt) + VNull(RstRep!Tax_Amt) + VNull(RstRep!Tax_Sur_Amt) + VNull(RstRep!Packing) + VNull(RstRep!ReSalTax_Amt) + VNull(RstRep!Tot_Amt)
    SubTot = Val(txt(STotATP)) + Val(txt(STotATB))
    If RstRep.RecordCount > 0 Then
        Do Until RstRep.EOF
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
                Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, mChr18 & "To," & mEmph
                mHeader = mHeader + 1
                Print #1, PSTR(RstRep!NamePrefix & " " & RstRep!Party_Name, 40) & Space(1) & PSTR(mDocStr & " NO.", Len(mDocStr) + 5) & " : " & PrinID(RstRep!DocID) & mEmph1
                mHeader = mHeader + 1
                Print #1, PSTR(IIf(mVType <> RetCashSalVType, XNull(RstRep!Add1), XNull(RstRep!Address)), 40) & Space(1) & mEmph & PSTR(mDocStr & " DATE", Len(mDocStr) + 5) & " : " & CDate(RstRep!V_Date) & mEmph1
                mHeader = mHeader + 1
                Print #1, IIf(mVType <> RetCashSalVType, XNull(RstRep!Add2), "")
                mHeader = mHeader + 1
                Print #1, IIf(mVType <> RetCashSalVType, XNull(RstRep!Add3) & IIf(XNull(RstRep!Add3) <> "" And XNull(RstRep!CityName) <> "", ",", "") & XNull(RstRep!CityName), "")
                mHeader = mHeader + 1
                Print #1, mDoub & "CST NO." & XNull(RstRep!CstNo) & mDoub1
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-") & mChr17 & mDoub
                mHeader = mHeader + 1
                Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 35) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
                mHeader = mHeader + 1
                Print #1, Space(89) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mDoub1 & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
                mHeader = mHeader + 1
                mFix = PageLength - (mHeader + mFooter)
                mLine = 1
            End If
            mRate = IIf(RstRep!MRP_YN = 1, RstRep!MRP_Rate, RstRep!Rate)
            PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstRep!Part_No, 22, , AlignLeft) & PSTR(RstRep!Part_Name, 35) & PSTR(RstRep!Qty_Rec, 12, 3)
            PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstRep!MRP_YN = 1, "M", "L") & _
            PSTR(RstRep!Disc_Per, 8, 2) & " %" & PSTR(RstRep!Disc_Amt, 10, 2) & _
            IIf(RstRep!Tax_YN = 0, PSTR(RstRep!Net_Amt, 12, 2) & PSTR(0, 12, 2), PSTR(0, 12, 2) & PSTR(RstRep!Net_Amt, 12, 2))
            
            Print #1, PrintStr
            RstRep.MoveNext
            mSlNo = mSlNo + 1
            mLine = mLine + 1
        Loop
    End If
    Do Until mLine >= mFix
        Print #1, ""
        mLine = mLine + 1
    Loop
    RstRep.MovePrevious
    Print #1, mChr18 & "Customer's Signature"
' SALE FOOTER
    '22 space maintain between heading and :
    Print #1, Replace(Space(21), " ", "-") & "TaxPaid" & Replace(Space(12), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")

    Print #1, PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 12, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
    ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstRep!Tax_Per, 5, 2) & "%" & PSTR(RstRep!Tax_Amt, 12, 2) & mDoub
    
    If PubSatYn Then
        Print #1, PSTR("MRP Items Amt", 16) & PSTR(RstRep!SprAmt_MRP_TP + RstRep!OilAmt_MRP_TP, 12, 2) & Space(8) & PSTR(RstRep!SprAmt_MRP_TB + RstRep!OilAmt_MRP_TB, 12, 2) & mDoub1 _
        ; " | " & PSTR("S A T ", 16, 0) & PSTR(RstRep!SatAmt, 12, 2) & mDoub
    Else
        Print #1, PSTR("MRP Items Amt", 16) & PSTR(RstRep!SprAmt_MRP_TP + RstRep!OilAmt_MRP_TP, 12, 2) & Space(8) & PSTR(RstRep!SprAmt_MRP_TB + RstRep!OilAmt_MRP_TB, 12, 2) & mDoub1 _
        ; " | " & PSTR("Tax Surc. ", 10, 0) & PSTR(RstRep!Tax_Sur_Per, 5, 2) & "%" & PSTR(RstRep!Tax_Sur_Amt, 12, 2) & mDoub
    End If
  
    Print #1, PSTR("Spares Amount", 16) & PSTR(RstRep!SprAmt_TP, 12, 2) & Space(8) & PSTR(RstRep!SprAmt_TB, 12, 2) & mDoub1 _
    ; " | " & PSTR("Misc. Charges", 16) & PSTR(RstRep!Packing, 12, 2) & mDoub

'"Itemwise Dis.Amt 01234567.12 00.00% 01234567.12 | Itemwise Dis.Amt 01234567.12"
'col1(16) col2(28) col3(35) col4(47) ,col5(50) ,col6(66) ,col7(78)
'col1(16) col2(12) col3(7) col4(12) ,col5(3) ,col6(16) ,col7(12)

    Print #1, PSTR("Oil Amount ", 16) & PSTR(RstRep!OilAmt_TP, 12, 2) & Space(8) & PSTR(RstRep!OilAmt_TB, 12, 2) & mDoub1 _
    ; " | " & mEmph & PSTR("Sub Total[TP&TB]", 16) & PSTR(Val(txt(STotB)), 12, 2) & mEmph1
    
    Print #1, PSTR("Discount ", 10, 0) & PSTR(RstRep!D_Per_TP, 5, 2) & "%" & PSTR(RstRep!D_Amt_TP, 12, 2) & PSTR(RstRep!D_Per_TB, 7, 2) & "%" & PSTR(RstRep!D_Amt_TB, 12, 2) _
    ; " | " & PSTR(pubTOTCaption, 10, 0) & PSTR(RstRep!TOT_Per, 5, 2) & "%" & PSTR(RstRep!Tot_Amt, 12, 2) & mEmph
    
    Print #1, PSTR("Sub Total [A]", 16) & PSTR(Val(txt(STotATP)), 12, 2) & Space(8) & PSTR(Val(txt(STotATB)), 12, 2) & mEmph1 _
    ; " | " & PSTR("ReSale Tax", 10, 0) & PSTR(RstRep!ReSalTax_Per, 5, 2) & "%" & PSTR(RstRep!ReSalTax_Amt, 12, 2)
    
    Print #1, PSTR("Gen Surch ", 10, 0) & PSTR(RstRep!Gen_Sur_Per, 5, 2) & "%" & PSTR(0, 12, 2) & PSTR(RstRep!Gen_Sur_Amt, 20, 2) _
    ; " | " & PSTR("Round Off", 16) & PSTR(RstRep!Rounded, 12, 2)
   
    Print #1, PSTR("Transportation", 16) & PSTR(0, 12, 2) & PSTR(RstRep!Trans_Amt, 20, 2) _
    ; " | " & mEmph & PSTR("Net Payble Rs.", 16) & PSTR(Val(txt(NetAmt)), 12, 2) & mEmph1
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mDoub & ntow(Val(txt(NetAmt)), "Rupees", "Paise") & mDoub1
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, "Remark : " & XNull(RstRep!Remarks)
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, PSTR(RstRep!Printing_Desc, 25) & PSTR("E & OE", PageWidth - 25, , AlignRight)
    Print #1, ""
    Print #1, Space(PageWidth - Len("For " & PubComp_Name)) & "For " & mEmph & PubComp_Name & mEmph1 & mDoub
    Print #1, "Terms & Conditions:" & mDoub1 & Replace(Space(PageWidth - 18), " ", "-") & mChr17
    Footer = Footer + vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    Print #1, Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
    Print #1, mEject
    Close #1
    Set RstRep = Nothing
    Set RstCompDet = Nothing
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
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        GCn.Execute "update Sp_Sale set Printed_YN = 1  where Sp_Sale.docid='" & Master!SearchCode & "' "
    End If
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update Sp_Sale set Printed_YN = 1  where Sp_Sale.docid='" & Master!SearchCode & "' "
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Sub Ini_Pub()
    Dim RsTemp As ADODB.Recordset
    
    Set RsTemp = GCn.Execute("Select CheckNegetiveStockSiteWise From Syctrl")
    If RsTemp.RecordCount > 0 Then
        mCheckNegetiveStockSiteWise = VNull(RsTemp!CheckNegetiveStockSiteWise)
    End If
End Sub

Sub DispTextVat()
        With FGrid
            If PubVATYN = 1 Then
                .TextMatrix(0, Col_TaxPer) = "TaxPer"
                .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
                .ColWidth(Col_TaxPer) = 840
                
                .TextMatrix(0, Col_TaxAmt1) = "TaxAmt"
                .ColAlignmentFixed(Col_TaxAmt1) = flexAlignRightCenter
                .ColWidth(Col_TaxAmt1) = 840
                
                If mSatYn Then
                    .TextMatrix(0, Col_SatPer) = "SAT %"
                    .ColAlignmentFixed(Col_SatPer) = flexAlignRightCenter
                    .ColWidth(Col_SatPer) = 840
                    
                    .TextMatrix(0, Col_SatAmt1) = "SAT Amt"
                    .ColAlignmentFixed(Col_SatAmt1) = flexAlignRightCenter
                    .ColWidth(Col_SatAmt1) = 840
                Else
                    .ColWidth(Col_SatPer) = 0
                    .ColWidth(Col_SatAmt1) = 0
                End If
            Else
                
                .TextMatrix(0, Col_TaxPer) = ""
                .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
                .ColWidth(Col_TaxPer) = 0
                
                .TextMatrix(0, Col_TaxAmt1) = ""
                .ColAlignmentFixed(Col_TaxAmt1) = flexAlignRightCenter
                .ColWidth(Col_TaxAmt1) = 0
                
                .ColWidth(Col_SatPer) = 0
                .ColWidth(Col_SatAmt1) = 0
            End If
        End With
End Sub

