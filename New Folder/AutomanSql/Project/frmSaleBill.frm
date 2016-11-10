VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmSaleBill 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Sale Bill Entry"
   ClientHeight    =   8595
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
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
      Height          =   225
      Index           =   54
      Left            =   9420
      TabIndex        =   205
      Top             =   5025
      Width           =   1890
   End
   Begin VB.CommandButton CmdTransPost 
      Caption         =   "Post Trans."
      Height          =   330
      Left            =   9720
      TabIndex        =   202
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Re-Post"
      Height          =   345
      Left            =   8490
      TabIndex        =   201
      Top             =   -15
      Width           =   1200
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
      Height          =   225
      Index           =   53
      Left            =   4410
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1170
      Width           =   885
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   3330
      Left            =   900
      Negotiate       =   -1  'True
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   8445
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17171
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
            ColumnWidth     =   3885.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3479.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2310.236
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   2670
      Left            =   -495
      Negotiate       =   -1  'True
      TabIndex        =   152
      TabStop         =   0   'False
      Top             =   8460
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
      ColumnCount     =   9
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
         DataField       =   "Disc_Factor"
         Caption         =   "Disc_Fact."
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
         Caption         =   "Cur.Stk"
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
         DataField       =   "Mrp"
         Caption         =   "   MRP"
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
      BeginProperty Column08 
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
            ColumnWidth     =   1049.953
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
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2564.788
         EndProperty
      EndProperty
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
      Left            =   2400
      TabIndex        =   199
      Top             =   3300
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
      Left            =   60
      TabIndex        =   186
      Top             =   7830
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
         Picture         =   "frmSaleBill.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   193
         ToolTipText     =   "Exit"
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
         Picture         =   "frmSaleBill.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   192
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmSaleBill.frx":0678
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
         TabIndex        =   191
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmSaleBill.frx":0982
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
         TabIndex        =   190
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmSaleBill.frx":0C8C
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
         TabIndex        =   189
         ToolTipText     =   "Printer "
         Top             =   285
         Width           =   1590
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
         TabIndex        =   188
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
         TabIndex        =   187
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
         TabIndex        =   196
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
         TabIndex        =   195
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
         TabIndex        =   194
         Top             =   0
         Width           =   4695
      End
   End
   Begin MSDataGridLib.DataGrid DGPerson 
      Height          =   3330
      Left            =   3375
      Negotiate       =   -1  'True
      TabIndex        =   185
      TabStop         =   0   'False
      Top             =   8385
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
   Begin VB.Frame FrmDetail 
      BackColor       =   &H00CAF1FD&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   2205
      Left            =   -6195
      TabIndex        =   153
      Top             =   2130
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
         TabIndex        =   184
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
         TabIndex        =   183
         Top             =   255
         Width           =   930
      End
      Begin VB.Label LblFrm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<Part No.>"
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
         TabIndex        =   182
         Top             =   255
         Width           =   870
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
         TabIndex        =   181
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
         TabIndex        =   180
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
         TabIndex        =   179
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
         TabIndex        =   178
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
         TabIndex        =   177
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
         TabIndex        =   176
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
         TabIndex        =   175
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
         TabIndex        =   174
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
         TabIndex        =   173
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
         TabIndex        =   172
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
         TabIndex        =   171
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
         TabIndex        =   170
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
         TabIndex        =   169
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
         TabIndex        =   168
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
         TabIndex        =   167
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
         TabIndex        =   166
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
         TabIndex        =   165
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
         TabIndex        =   164
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
         TabIndex        =   163
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
         TabIndex        =   162
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
         TabIndex        =   161
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
         TabIndex        =   160
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
         TabIndex        =   159
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
         TabIndex        =   158
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
         TabIndex        =   157
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
         TabIndex        =   156
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
         TabIndex        =   155
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
         TabIndex        =   154
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
   Begin VB.TextBox Txt 
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
      Height          =   225
      Index           =   52
      Left            =   10140
      TabIndex        =   151
      Text            =   "Sale A/c"
      ToolTipText     =   "Sale A/c Code"
      Top             =   7665
      Width           =   1215
   End
   Begin VB.TextBox Txt 
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
      Height          =   225
      Index           =   51
      Left            =   9465
      TabIndex        =   150
      Text            =   "Sur A/c"
      ToolTipText     =   "Tax Sur A/c"
      Top             =   7665
      Width           =   600
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
      Left            =   3435
      TabIndex        =   149
      Text            =   "ChalSrlNo"
      ToolTipText     =   "Challan Serial No"
      Top             =   -15
      Visible         =   0   'False
      Width           =   705
   End
   Begin MSDataGridLib.DataGrid DGOrdPart 
      Height          =   2625
      Left            =   -10215
      Negotiate       =   -1  'True
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   -1935
      Visible         =   0   'False
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   4630
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   16
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Srl_No"
         Caption         =   "Srl.No."
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
         DataField       =   "Part_No"
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
      BeginProperty Column02 
         DataField       =   "Part_Name"
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
      BeginProperty Column03 
         DataField       =   "Qty"
         Caption         =   "Ord.Qty"
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
         DataField       =   "Sup_Qty"
         Caption         =   "Sup.Qty"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PendQty"
         Caption         =   "Pending Qty"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
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
            Alignment       =   1
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3149.858
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1289.764
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGSONo 
      Height          =   2775
      Left            =   -15
      Negotiate       =   -1  'True
      TabIndex        =   147
      TabStop         =   0   'False
      Top             =   8415
      Visible         =   0   'False
      Width           =   7110
      _ExtentX        =   12541
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
      RowHeight       =   16
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Our Doc.No."
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
         Caption         =   "Our Doc. Date"
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
         DataField       =   "Party_RefDoc_No"
         Caption         =   "Party Ref No."
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
         DataField       =   "Party_RefDoc_Date"
         Caption         =   "Ref. Date"
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
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1530.142
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   2040
      Negotiate       =   -1  'True
      TabIndex        =   146
      TabStop         =   0   'False
      Top             =   8355
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Voucher No"
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
            ColumnWidth     =   2865.26
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGTrans 
      Height          =   3330
      Left            =   1650
      Negotiate       =   -1  'True
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   8490
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
      RowHeight       =   16
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
         Caption         =   "Transporter"
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
      EndProperty
   End
   Begin VB.Frame FrmSel 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2130
      Left            =   -30
      TabIndex        =   139
      Top             =   7800
      Visible         =   0   'False
      Width           =   5115
      Begin VB.CommandButton CmdSel 
         BackColor       =   &H00CFE0E0&
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   1800
         Width           =   1155
      End
      Begin VB.CommandButton CmdSel 
         BackColor       =   &H00CFE0E0&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3930
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   1800
         Width           =   1155
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGridSel 
         Height          =   1695
         Left            =   45
         TabIndex        =   142
         Top             =   30
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   12243913
         Cols            =   5
         BackColorFixed  =   128
         ForeColorFixed  =   65535
         BackColorSel    =   16711680
         BackColorBkg    =   13623520
         GridColor       =   16512
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
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
   Begin MSDataGridLib.DataGrid DGGodown 
      Height          =   4845
      Left            =   6870
      Negotiate       =   -1  'True
      TabIndex        =   137
      TabStop         =   0   'False
      Top             =   8550
      Visible         =   0   'False
      Width           =   7710
      _ExtentX        =   13600
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
   Begin MSDataGridLib.DataGrid DGForm31 
      Height          =   3330
      Left            =   8565
      Negotiate       =   -1  'True
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   8490
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
      Height          =   225
      Index           =   50
      Left            =   4215
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1425
      Width           =   1080
   End
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   3330
      Left            =   825
      Negotiate       =   -1  'True
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   8430
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
      RowHeight       =   16
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
            ColumnWidth     =   5265.071
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
      Left            =   5220
      Negotiate       =   -1  'True
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   8505
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
      Height          =   225
      Index           =   10
      Left            =   1395
      MaxLength       =   7
      TabIndex        =   8
      ToolTipText     =   "Press L-> Local or C-> Central"
      Top             =   915
      Width           =   1095
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
      Height          =   225
      Index           =   0
      Left            =   9480
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   465
      Width           =   2280
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
      Height          =   225
      Index           =   15
      Left            =   6765
      MaxLength       =   3
      TabIndex        =   20
      Text            =   "Yes"
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   1680
      Width           =   435
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   -360
      TabIndex        =   118
      Top             =   8520
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   60
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   225
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
      Height          =   225
      Index           =   18
      Left            =   3795
      MaxLength       =   10
      TabIndex        =   16
      Top             =   1935
      Width           =   1500
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
      Height          =   225
      Index           =   8
      Left            =   1395
      MaxLength       =   40
      TabIndex        =   10
      Top             =   1170
      Width           =   2115
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
      Height          =   225
      Index           =   48
      Left            =   4395
      TabIndex        =   52
      Top             =   6345
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
      Height          =   225
      Index           =   38
      Left            =   9420
      TabIndex        =   41
      Top             =   4515
      Width           =   1890
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
      Height          =   225
      Index           =   47
      Left            =   9420
      TabIndex        =   51
      Top             =   6300
      Width           =   1890
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
      Height          =   225
      Index           =   22
      Left            =   5040
      TabIndex        =   27
      Top             =   4800
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
      Height          =   225
      Index           =   21
      Left            =   2925
      TabIndex        =   26
      Top             =   4800
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
      Height          =   225
      Index           =   20
      Left            =   5040
      TabIndex        =   25
      Top             =   4545
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
      TabIndex        =   54
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
      Height          =   225
      Index           =   6
      Left            =   1395
      MaxLength       =   40
      TabIndex        =   5
      Top             =   660
      Width           =   3900
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
      Height          =   225
      Index           =   1
      Left            =   9840
      TabIndex        =   1
      ToolTipText     =   "Press C-> Cash or R-> Credit"
      Top             =   915
      Width           =   915
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
      Height          =   225
      Index           =   13
      Left            =   6765
      MaxLength       =   15
      TabIndex        =   18
      Top             =   1170
      Width           =   1500
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
      Height          =   225
      Index           =   14
      Left            =   6765
      MaxLength       =   11
      TabIndex        =   19
      Top             =   1425
      Width           =   1500
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
      Height          =   225
      Index           =   11
      Left            =   3150
      MaxLength       =   25
      TabIndex        =   9
      Top             =   915
      Width           =   2145
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
      Height          =   225
      Index           =   12
      Left            =   6210
      MaxLength       =   40
      TabIndex        =   17
      Top             =   915
      Width           =   2355
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
      Height          =   225
      Index           =   46
      Left            =   9420
      TabIndex        =   50
      Top             =   6045
      Width           =   1890
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
      Height          =   225
      Index           =   45
      Left            =   10095
      TabIndex        =   49
      Top             =   5790
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
      Height          =   225
      Index           =   44
      Left            =   9420
      TabIndex        =   48
      ToolTipText     =   "Turn Over Tax %"
      Top             =   5790
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
      Height          =   225
      Index           =   43
      Left            =   9420
      TabIndex        =   47
      Top             =   5535
      Width           =   1890
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
      Height          =   225
      Index           =   42
      Left            =   10170
      TabIndex        =   45
      Top             =   7410
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
      Height          =   225
      Index           =   41
      Left            =   9495
      TabIndex        =   44
      ToolTipText     =   "Surcharge % on Local Sales Tax"
      Top             =   7410
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
      Height          =   225
      Index           =   40
      Left            =   10095
      TabIndex        =   43
      Top             =   4770
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
      Height          =   225
      Index           =   39
      Left            =   9420
      TabIndex        =   42
      ToolTipText     =   "Local Sales Tax %"
      Top             =   4770
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
      Height          =   225
      Index           =   37
      Left            =   2925
      TabIndex        =   40
      Top             =   6330
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
      Height          =   225
      Index           =   36
      Left            =   2925
      TabIndex        =   39
      Top             =   6075
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
      Height          =   225
      Index           =   35
      Left            =   2235
      TabIndex        =   38
      ToolTipText     =   "General Surcharge %"
      Top             =   6075
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
      Height          =   225
      Index           =   19
      Left            =   2925
      TabIndex        =   24
      Top             =   4545
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
      Height          =   225
      Index           =   34
      Left            =   9420
      TabIndex        =   46
      Top             =   5280
      Width           =   1890
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   33
      Left            =   5610
      TabIndex        =   53
      Text            =   "WithDrawn"
      Top             =   7560
      Visible         =   0   'False
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
      Height          =   225
      Index           =   32
      Left            =   5040
      TabIndex        =   37
      Top             =   5820
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
      Height          =   225
      Index           =   31
      Left            =   2925
      TabIndex        =   36
      Text            =   "99999999.99"
      Top             =   5820
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
      Height          =   225
      Index           =   30
      Left            =   5040
      TabIndex        =   35
      Top             =   5565
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
      Height          =   225
      Index           =   29
      Left            =   4350
      TabIndex        =   34
      ToolTipText     =   "Discount % Taxpaid"
      Top             =   5565
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
      Height          =   225
      Index           =   28
      Left            =   2925
      TabIndex        =   33
      Text            =   "99999999.99"
      Top             =   5565
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
      Height          =   225
      Index           =   27
      Left            =   2235
      TabIndex        =   32
      Text            =   "99.99"
      ToolTipText     =   "Discount % Taxable"
      Top             =   5565
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
      Height          =   225
      Index           =   26
      Left            =   5040
      TabIndex        =   31
      Top             =   5310
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
      Height          =   225
      Index           =   25
      Left            =   2925
      TabIndex        =   30
      Top             =   5310
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
      Height          =   225
      Index           =   24
      Left            =   5040
      TabIndex        =   29
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
      Height          =   225
      Index           =   23
      Left            =   2925
      TabIndex        =   28
      Top             =   5055
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
      Height          =   225
      Index           =   2
      Left            =   9840
      MaxLength       =   11
      TabIndex        =   2
      Top             =   1170
      Width           =   1560
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
      Height          =   225
      Index           =   3
      Left            =   10500
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1410
      Width           =   900
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
      Height          =   225
      Index           =   16
      Left            =   6765
      MaxLength       =   50
      TabIndex        =   21
      Top             =   1935
      Width           =   4680
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
      Index           =   7
      Left            =   7635
      MaxLength       =   40
      TabIndex        =   22
      Top             =   8010
      Visible         =   0   'False
      Width           =   3855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1965
      Left            =   90
      TabIndex        =   23
      Top             =   2175
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   3466
      _Version        =   393216
      BackColor       =   13166810
      ForeColor       =   0
      Cols            =   34
      BackColorFixed  =   12632319
      ForeColorFixed  =   128
      BackColorSel    =   13166810
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   12632319
      GridColorFixed  =   33023
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "MW"
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
   Begin VB.OptionButton OptChal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00BAD3C9&
      Caption         =   "Select Challan"
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
      Index           =   0
      Left            =   5325
      TabIndex        =   6
      Top             =   645
      Width           =   1500
   End
   Begin VB.OptionButton OptChal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00BAD3C9&
      Caption         =   "Create Challan"
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
      Index           =   1
      Left            =   6885
      TabIndex        =   7
      Top             =   645
      Width           =   1545
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
      Height          =   225
      Index           =   5
      Left            =   1395
      MaxLength       =   40
      TabIndex        =   4
      Top             =   405
      Width           =   3900
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
      Height          =   225
      Index           =   9
      Left            =   1395
      MaxLength       =   40
      TabIndex        =   14
      Top             =   1680
      Width           =   3900
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
      Height          =   225
      Index           =   17
      Left            =   1395
      MaxLength       =   4
      TabIndex        =   15
      Top             =   1935
      Width           =   1260
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
      Height          =   225
      Index           =   49
      Left            =   1395
      MaxLength       =   40
      TabIndex        =   12
      Top             =   1425
      Width           =   1515
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
      Left            =   210
      TabIndex        =   208
      Top             =   6660
      Width           =   45
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additonal Tax"
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
      Left            =   7500
      TabIndex        =   207
      Top             =   5025
      Width           =   1095
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
      TabIndex        =   206
      Top             =   5025
      Width           =   180
   End
   Begin VB.Label LblCurrBal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curr Bal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5355
      TabIndex        =   204
      Top             =   435
      Width           =   720
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
      Left            =   6660
      TabIndex        =   203
      Top             =   1935
      Width           =   45
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Type"
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
      Index           =   4
      Left            =   3540
      TabIndex        =   200
      Top             =   1170
      Width           =   840
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
      Left            =   7380
      TabIndex        =   198
      Top             =   1680
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label LblVPrefix2 
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
      Left            =   2745
      TabIndex        =   197
      Top             =   30
      Visible         =   0   'False
      Width           =   600
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
      Left            =   3465
      TabIndex        =   65
      Top             =   4290
      Width           =   675
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   600
      Left            =   4230
      Shape           =   4  'Rounded Rectangle
      Top             =   6075
      Width           =   1770
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      Height          =   600
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   6075
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gate Pass No."
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
      Height          =   240
      Index           =   43
      Left            =   6165
      TabIndex        =   145
      Top             =   6075
      Width           =   1200
   End
   Begin VB.Label lblGatePass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GP No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   240
      Left            =   6345
      TabIndex        =   144
      Top             =   6345
      Width           =   660
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      Height          =   1335
      Left            =   8610
      Top             =   405
      Width           =   3180
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
      Left            =   9840
      TabIndex        =   138
      Top             =   1425
      Width           =   600
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
      Height          =   225
      Index           =   36
      Left            =   3180
      TabIndex        =   135
      Top             =   1440
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
      Height          =   225
      Index           =   35
      Left            =   1320
      TabIndex        =   134
      Top             =   1410
      Width           =   45
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
      Left            =   75
      TabIndex        =   133
      Top             =   1425
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
      Left            =   1320
      TabIndex        =   129
      Top             =   915
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
      Left            =   75
      TabIndex        =   128
      ToolTipText     =   "Press L-> Local or C-> Central"
      Top             =   870
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
      Height          =   195
      Index           =   3
      Left            =   1320
      TabIndex        =   127
      Top             =   435
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   25
      Left            =   9360
      TabIndex        =   126
      Top             =   465
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
      Left            =   8685
      TabIndex        =   125
      Top             =   465
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   9
      Left            =   75
      TabIndex        =   124
      Top             =   405
      Width           =   450
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
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   29
      Left            =   5340
      TabIndex        =   122
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   1680
      Width           =   1335
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
      Left            =   10425
      TabIndex        =   121
      Top             =   660
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
      Left            =   8685
      TabIndex        =   120
      Top             =   690
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
      Left            =   2850
      TabIndex        =   117
      Top             =   1935
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
      Height          =   225
      Index           =   32
      Left            =   1320
      TabIndex        =   116
      Top             =   1170
      Width           =   45
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Form"
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
      Index           =   39
      Left            =   75
      TabIndex        =   115
      Top             =   1170
      Width           =   795
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Credit"
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
      Left            =   8685
      TabIndex        =   112
      ToolTipText     =   "Press C-> Cash or R-> Credit"
      Top             =   915
      Width           =   990
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
      TabIndex        =   111
      Top             =   5055
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
      TabIndex        =   110
      Top             =   4545
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
      Left            =   75
      TabIndex        =   109
      Top             =   660
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
      TabIndex        =   108
      Top             =   4800
      Width           =   1665
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET PAYABLE AMT"
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
      Height          =   210
      Index           =   30
      Left            =   4320
      TabIndex        =   91
      Top             =   6090
      Width           =   1590
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
      Height          =   225
      Index           =   37
      Left            =   7500
      TabIndex        =   107
      Top             =   4530
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
      TabIndex        =   106
      Top             =   4515
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode"
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
      Index           =   34
      Left            =   2565
      TabIndex        =   105
      Top             =   915
      Width           =   450
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
      Index           =   31
      Left            =   3930
      TabIndex        =   104
      Top             =   915
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NET SPARE AMT"
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
      Left            =   7500
      TabIndex        =   103
      Top             =   6315
      Width           =   1380
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
      Index           =   26
      Left            =   9105
      TabIndex        =   102
      Top             =   6300
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
      Index           =   30
      Left            =   1965
      TabIndex        =   101
      Top             =   4800
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
      TabIndex        =   100
      Top             =   4260
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
      Left            =   11115
      TabIndex        =   99
      Top             =   4290
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
      Left            =   9075
      TabIndex        =   98
      Top             =   4260
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
      TabIndex        =   97
      Top             =   4260
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
      TabIndex        =   96
      Top             =   4260
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
      TabIndex        =   95
      Top             =   4290
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
      Left            =   1320
      TabIndex        =   94
      Top             =   660
      Width           =   195
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LR/RR No."
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
      Index           =   6
      Left            =   5340
      TabIndex        =   93
      Top             =   1170
      Width           =   885
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LR/RR Date"
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
      Index           =   5
      Left            =   5340
      TabIndex        =   92
      Top             =   1425
      Width           =   990
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
      Height          =   270
      Index           =   28
      Left            =   5355
      TabIndex        =   90
      Top             =   1935
      Width           =   765
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transport"
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
      Left            =   5340
      TabIndex        =   89
      Top             =   900
      Width           =   795
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
      TabIndex        =   88
      Top             =   6045
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
      Left            =   7500
      TabIndex        =   87
      Top             =   6045
      Width           =   1365
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOT on Sub Total (B)"
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
      Index           =   25
      Left            =   7500
      TabIndex        =   86
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   5805
      Width           =   1710
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total (B) TB+TP"
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
      Index           =   24
      Left            =   7500
      TabIndex        =   85
      Top             =   5550
      Width           =   1680
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
      Left            =   9180
      TabIndex        =   84
      Top             =   7410
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
      Left            =   7575
      TabIndex        =   83
      Top             =   7410
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
      TabIndex        =   82
      Top             =   4770
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
      Left            =   7500
      TabIndex        =   81
      Top             =   4770
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
      TabIndex        =   80
      Top             =   6330
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
      TabIndex        =   79
      Top             =   6330
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
      TabIndex        =   78
      Top             =   6075
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
      TabIndex        =   77
      Top             =   6075
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
      TabIndex        =   76
      Top             =   4545
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
      TabIndex        =   75
      Top             =   5280
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
      Height          =   225
      Index           =   18
      Left            =   7500
      TabIndex        =   74
      Top             =   5295
      Width           =   1185
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReSaleTax                :"
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
      Index           =   17
      Left            =   7545
      TabIndex        =   73
      Top             =   7680
      Width           =   1665
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
      TabIndex        =   72
      Top             =   5820
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
      TabIndex        =   71
      Top             =   5820
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
      TabIndex        =   70
      Top             =   5565
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
      TabIndex        =   69
      Top             =   5565
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
      TabIndex        =   68
      Top             =   5310
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
      TabIndex        =   67
      Top             =   5310
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
      Left            =   5490
      TabIndex        =   66
      Top             =   4290
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   5
      Left            =   1965
      TabIndex        =   64
      Top             =   5055
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
      Height          =   225
      Index           =   2
      Left            =   1320
      TabIndex        =   63
      Top             =   1935
      Width           =   45
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
      Left            =   75
      TabIndex        =   62
      Top             =   1935
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
      Left            =   75
      TabIndex        =   61
      Top             =   1680
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
      Height          =   225
      Index           =   1
      Left            =   1320
      TabIndex        =   60
      Top             =   1680
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
      Index           =   90
      Left            =   9705
      TabIndex        =   59
      Top             =   915
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
      Left            =   9705
      TabIndex        =   58
      Top             =   1425
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
      Left            =   9705
      TabIndex        =   57
      Top             =   1170
      Width           =   180
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doc. Date"
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
      Index           =   0
      Left            =   8685
      TabIndex        =   56
      Top             =   1170
      Width           =   810
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doc. No."
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
      Left            =   8685
      TabIndex        =   55
      Top             =   1425
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   22
      Left            =   7515
      TabIndex        =   114
      Top             =   8010
      Visible         =   0   'False
      Width           =   45
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   38
      Left            =   6780
      TabIndex        =   113
      Top             =   8010
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmSaleBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMRevDisTBPer As Double, mMRevDisTPPer As Double
Dim mTBDisAmtMRP As Double, mTPDisAmtMRP As Double
Dim mMRPTax As Double, mMRPTaxSur As Double, mMRPTOT As Double, mMRPReSales As Double
Dim mMRPLubeTB As Double, mMRPLubeTP  As Double
Private Const PageWidth As Byte = 80
Dim FirstPrint As Boolean
Dim CurrStk%
Dim mReposting As Boolean
Dim mRePostCounter As Integer
Dim mVatYn As Byte



Dim mCheckNegetiveStockSiteWise As Boolean
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim RsParty As ADODB.Recordset
Dim RsCrAc As ADODB.Recordset
Dim rsForm As ADODB.Recordset
Dim rsForm31 As ADODB.Recordset
Dim rsTrans As ADODB.Recordset
Dim RsPerson As ADODB.Recordset
Dim RsSONo As ADODB.Recordset
Dim RsGodown As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim rsCtrlAc As ADODB.Recordset

Dim mAddFlag$
Dim mVType As String, mVPrefix As String
Dim mSearchCode As String
Dim FillChallanFlag As Byte
Dim mPartyType As Byte         'Used to Detect Part Rate in GetRate Function
Dim ForSiteCode As String
Private Const SalChalType As String = "SYSC"
Private Const TrfChalType As String = "SYSCT"
Private Const SalCrVType As String = "SYSIR"
Private Const SalCashVType As String = "SYSIC"

'grid color scheme
Private Const CellBackColLeave As String = &HC8E8DA
'Private Const CellForeColLeave As String = &H0&
'Private Const CellBackColEnter As String = &HC0E0FF
Private Const GridBackColorBkg As String = &HBAD3C9
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

' Under observation
Dim VoucherEditFlag As Boolean                  ' Used for whether we can edit voucher no or not
' End Under observation
Dim ListArray As Variant
Dim mListItem As ListItem

Private Const DocID As Byte = 0                 ' Doc.ID
Private Const DocType As Byte = 1               ' Document Type
Private Const VDate As Byte = 2                 ' Date
Private Const SerialNo As Byte = 3              ' Serial No.
Private Const SerialNo2 As Byte = 4              ' Serial No. Challan
'Private Const CashCr As Byte = 4                ' Cash/Credit
Private Const Party As Byte = 5                 ' Party Name
Private Const Address1 As Byte = 6              ' Address1
Private Const CrAc As Byte = 7                  ' Debit A/c
Private Const FormName As Byte = 8              ' Form Name
Private Const SPerson As Byte = 9               ' Sales Person
Private Const LC As Byte = 10                   ' Dispatch Type(Local/Central)
Private Const DispMode As Byte = 11             ' Dispatch Mode
Private Const Transport As Byte = 12            ' Transport
Private Const LRNo As Byte = 13                 ' LR/RR No.
Private Const LRDate As Byte = 14               ' LR/RR Date
Private Const TaxDet As Byte = 15               ' Print Tax Detail(Y/N)
Private Const Remark As Byte = 16               ' Remark
Private Const CaseNo As Byte = 17               ' Case No.
Private Const CaseMark As Byte = 18             ' Case Mark

Private Const IWDiscTotTB As Byte = 19          ' Item-wise Disc Total Taxable
Private Const IWDiscTotTP As Byte = 20          ' Item-wise Disc Total Taxpaid
Private Const MRPAmtTB As Byte = 21             ' MRP Item's Amount Taxable
Private Const MRPAmtTP As Byte = 22             ' MRP Item's Amount Taxpaid
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
Private Const Addition As Byte = 33             '
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
Private Const Form31Name As Byte = 49           ' Form 31 Name
Private Const Form31No As Byte = 50             ' Form 31 No
Private Const ReSalTaxPer As Byte = 51          '
Private Const ReSalTaxAmt As Byte = 52            '
Private Const PType As Byte = 53            '
Private Const SatAmt As Byte = 54            '

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_SONoCode As Byte = 2          ' Sale Order No Code
Private Const Col_SONo As Byte = 1              ' Sale Order No Name
Private Const Col_SOSrNo As Byte = 3            ' Sale Order Serial No
Private Const Col_PNo As Byte = 4               ' Part No
Private Const Col_ChalNoCode As Byte = 5       ' Sale Challan No Code
Private Const Col_ChalNo As Byte = 6           ' Sale Challan No Name
Private Const Col_ChalSrNo As Byte = 7         ' Sale Challan Serial No
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
Private Const Col_TaxAmt1 As Byte = 18          ' Tax. Amt.
Private Const Col_SatPer As Byte = 19           ' Tax Per.
Private Const Col_SatAmt As Byte = 20          ' Tax. Amt.
Private Const Col_ItemVal As Byte = 21         ' Item Value
Private Const Col_GodownCode As Byte = 22       ' Godown Code
Private Const Col_Godown As Byte = 23           ' Godown
Private Const Col_PartSrlNo As Byte = 24        ' Part Serial No
Private Const Col_PName As Byte = 25            ' Part Name
Private Const Col_LName As Byte = 26            ' Local Name
Private Const Col_MRPStkTB As Byte = 27         ' MRP Qty TB 'Current Stock Qty
Private Const Col_MRPStkTP As Byte = 28         ' MRP Qty TP
Private Const Col_TBStk As Byte = 29            ' Taxbale Qty
Private Const Col_TPStk As Byte = 30            ' Tax Paid Qty
Private Const Col_TBRate As Byte = 31           ' Taxbale Rate
Private Const Col_TPRate As Byte = 32           ' Tax Paid Rate
Private Const Col_Bin As Byte = 33              ' Bin
Private Const Col_LastRate As Byte = 34         ' Last Purchase Rate
Private Const Col_HPRate As Byte = 35           ' High Purchase Rate
Private Const Col_LPRate As Byte = 36           ' Low Purchase Rate
Private Const Col_PartGrade As Byte = 37        ' Part Grade (Used for Oil Item)
Private Const Col_EffectDate As Byte = 38       ' MRP Effective Date/TB Effective Date
Private Const Col_PurDocId As Byte = 39
Private Const Col_PurDate As Byte = 40


'* Challan Selection Grid
Private Const SCol_SrNo As Byte = 0             ' Serial No
Private Const SCol_ChalNoCode As Byte = 1       ' Sale Challan No Code
Private Const SCol_ChalNo As Byte = 2           ' Sale Challan No Name
Private Const SCol_ChalSrNo As Byte = 3         ' Sale Challan Serial No
Private Const SCol_ChalDate As Byte = 4         ' Challan Date

Private Const FromVno As Byte = 0
Private Const ToVno As Byte = 1
Private Const VType1 As Byte = 2
Private Const ChalSelect As Byte = 0
Private Const ChalCreate As Byte = 1

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String
Dim CustOrdDet As String
Dim rsTaxPer As ADODB.Recordset


Dim mSatYn As Boolean

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To Txt.Count - 1
        If I = DocID Or I = IWDiscTotTB Or I = IWDiscTotTP Or I = MRPAmtTB _
            Or I = MRPAmtTP Or I = SprAmtTB Or I = SprAmtTP Or I = OilAmtTB _
            Or I = OilAmtTP Or I = STotATB Or I = STotATP Or I = TaxableTot _
            Or I = STotB Or I = SROff Or I = NetSprAmt Or I = NetAmt Then
        Else
            Txt(I).Enabled = Enb
        End If
    Next
    OptChal(ChalSelect).Enabled = Enb
    OptChal(ChalCreate).Enabled = Enb
    If PubSiebelActiveYn = 1 And pubUName = "SA" Then
        cmdPost.Visible = True
    Else
        cmdPost.Visible = False
    End If
    
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("SearchCode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select S.DocID As SearchCode,U_EntDt, V_Date From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' and S.V_Type In ('" & SalCashVType & "','" & SalCrVType & "') And S.DocID  = '" & MyValue & "' " _
            & "Order by S.V_Date desc,S.DocID desc")
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
    For I = 0 To Txt.Count - 1
        Txt(I).TEXT = ""
'        txt(i).Tag = ""
    Next I
    Txt(DocID).Tag = ""
    LblDiv.CAPTION = "Division : "
    LblSite.CAPTION = "Site Code : "
    LblVPrefix.CAPTION = ""
    LblIVal.CAPTION = ""
    LblQty.CAPTION = ""
    lblGatePass.CAPTION = ""
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

    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End Sub
'* Used for intialize grid columns
Private Sub Grid_Ini()
    With FGrid
        .left = Me.left '+ 60
        .width = Me.width - 90
        .top = 2200 '2550 '2610
        .BackColor = CellBackColLeave
        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 41

        .TextMatrix(0, Col_SrNo) = "S.No"
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 420

        .TextMatrix(0, Col_SONoCode) = "SO No. Code"
        .ColAlignment(Col_SONoCode) = flexAlignLeftCenter
        .ColWidth(Col_SONoCode) = 0

        .TextMatrix(0, Col_SONo) = "SO No."
        .ColAlignment(Col_SONo) = flexAlignLeftCenter
        .ColWidth(Col_SONo) = 1150

        .TextMatrix(0, Col_SOSrNo) = "SO Srl No."
        .ColAlignment(Col_SOSrNo) = flexAlignLeftCenter
        .ColWidth(Col_SOSrNo) = 0
        
        .TextMatrix(0, Col_PNo) = "Part No"
        .ColAlignment(Col_PNo) = flexAlignLeftCenter
        .ColWidth(Col_PNo) = 1500

        .TextMatrix(0, Col_ChalNoCode) = "Challan No. Code"
        .ColAlignment(Col_ChalNoCode) = flexAlignLeftCenter
        .ColWidth(Col_ChalNoCode) = 0

        .TextMatrix(0, Col_ChalNo) = "Challan No."
        .ColAlignment(Col_ChalNo) = flexAlignLeftCenter
        .ColWidth(Col_ChalNo) = 1150

        .TextMatrix(0, Col_ChalSrNo) = "Challan Srl No."
        .ColAlignment(Col_ChalSrNo) = flexAlignLeftCenter
        .ColWidth(Col_ChalSrNo) = 0

        .TextMatrix(0, Col_Unit) = "Unit"
        .ColAlignment(Col_Unit) = flexAlignLeftCenter
        .ColWidth(Col_Unit) = 500

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
        .ColWidth(Col_Rate) = 1000

        .TextMatrix(0, Col_MRPRate) = "MRP Rate"
        .ColAlignmentFixed(Col_MRPRate) = flexAlignRightCenter
        .ColWidth(Col_MRPRate) = 1000

        .TextMatrix(0, Col_Amt) = "Amount"
        .ColAlignmentFixed(Col_Amt) = flexAlignRightCenter
        .ColWidth(Col_Amt) = 1065

        .TextMatrix(0, Col_DiscPer) = "Disc%"
        .ColAlignmentFixed(Col_DiscPer) = flexAlignRightCenter
        .ColWidth(Col_DiscPer) = 700

        .TextMatrix(0, Col_DiscAmt) = "Disc.Amt"
        .ColAlignmentFixed(Col_DiscAmt) = flexAlignRightCenter
        .ColWidth(Col_DiscAmt) = 840
        
        If mVatYn = 1 Then
            .TextMatrix(0, Col_TaxPer) = "TaxPer"
            .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
            .ColWidth(Col_TaxPer) = 840
            
            .TextMatrix(0, Col_TaxAmt1) = "TaxAmt"
            .ColAlignmentFixed(Col_TaxAmt1) = flexAlignRightCenter
            .ColWidth(Col_TaxAmt1) = 840
        
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
            
            .TextMatrix(0, Col_TaxAmt1) = ""
            .ColAlignmentFixed(Col_TaxAmt1) = flexAlignRightCenter
            .ColWidth(Col_TaxAmt1) = 0
            
            .ColWidth(Col_SatPer) = 0
            .ColWidth(Col_SatAmt) = 0
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

        .TextMatrix(0, Col_MRPStkTP) = "MRP Qty TP"
        .ColAlignmentFixed(Col_MRPStkTP) = flexAlignRightCenter
        .ColWidth(Col_MRPStkTP) = 0

        .TextMatrix(0, Col_MRPStkTB) = "MRP Qty TB"
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
        
        .TextMatrix(0, Col_PurDocId) = "Purch Doc No"
        .ColWidth(Col_PurDocId) = 2500
        
        .TextMatrix(0, Col_PurDate) = "Purch Doc Date"
        .ColWidth(Col_PurDate) = 2000
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    With FGridSel
        .RowHeightMin = PubGridRowHeight
        .Cols = 5

        .TextMatrix(0, SCol_SrNo) = ""
        .ColAlignment(SCol_SrNo) = flexAlignLeftCenter
        .ColWidth(SCol_SrNo) = 450

        .TextMatrix(0, SCol_ChalNoCode) = "Sale Challan No Code"
        .ColAlignment(SCol_ChalNoCode) = flexAlignLeftCenter
        .ColWidth(SCol_ChalNoCode) = 0

        .TextMatrix(0, SCol_ChalNo) = "Challan No"
        .ColAlignment(SCol_ChalNo) = flexAlignLeftCenter
        .ColWidth(SCol_ChalNo) = 2000

        .TextMatrix(0, SCol_ChalSrNo) = "Sale Challan Serial No"
        .ColAlignment(SCol_ChalSrNo) = flexAlignLeftCenter
        .ColWidth(SCol_ChalSrNo) = 0

        .TextMatrix(0, SCol_ChalDate) = "Date"
        .ColAlignment(SCol_ChalDate) = flexAlignLeftCenter
        .ColWidth(SCol_ChalDate) = 1500
    End With
    DGPart.width = FGrid.width: DGPart.left = FGrid.left: DGPart.top = FGrid.top + FGrid.height: DGPart.height = Me.height - (DGPart.top + mBotScale)
    DGSONo.left = FGrid.left: DGSONo.top = DGPart.top: DGSONo.height = DGPart.height
    DGGodown.left = Me.width - (DGGodown.width + mRtScale): DGGodown.top = DGPart.top: DGGodown.height = DGPart.height
    FrmDetail.width = 6285: FrmDetail.left = 5595: FrmDetail.top = 200: FrmDetail.height = 2130
    'DGParty.left = Me.width - (DGParty.width + mRtScale): DGParty.top = mTopScale
    DGParty.left = mRtScale: DGParty.top = mTopScale + FGrid.top
    DGCrAc.left = Me.width - (DGCrAc.width + mRtScale): DGCrAc.top = mTopScale
    DGForm.left = Me.width - (DGForm.width + mRtScale): DGForm.top = mTopScale
    DGForm31.left = Me.width - (DGForm31.width + mRtScale): DGForm31.top = mTopScale
    DGTrans.left = mLtScale: DGTrans.top = mTopScale
    FrmSel.left = Me.width - (FrmSel.width + mRtScale): FrmSel.top = mTopScale
    DGOrdPart.left = (Me.width - DGOrdPart.width) / 2: DGOrdPart.top = FGrid.top + FGrid.height: DGOrdPart.height = Me.height - (DGOrdPart.top + mBotScale)
    DGPerson.left = Me.width - (DGPerson.width + mRtScale): DGPerson.top = mTopScale
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
    If DGTrans.Visible = True Then DGTrans.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGPart.Visible = True Then DGPart.Visible = False
    If DGSONo.Visible = True Then DGSONo.Visible = False
End Sub

Private Sub cmdPost_Click()
Dim I As Integer, mStartdate As String, mEndDate As String
Dim DupMaster As ADODB.Recordset

If Master.RecordCount > 0 Then
    Master.MoveFirst
    mStartdate = InputBox("Posting Required from which Date ?", "Start Date for Posting", PubLoginDate)
    mEndDate = InputBox("Posting Required upto which Date ?", "Last Date for Posting", PubLoginDate)
    
    If mStartdate = "" Or mEndDate = "" Then Exit Sub
    mStartdate = MakeDate(mStartdate)
    mEndDate = MakeDate(mEndDate)
    
    mReposting = True
    mRePostCounter = 1
    
    Do Until Master.EOF
        If IsNull(Master!V_DATE) Then GoTo MyNextRecord
        If Master!V_DATE < CDate(mStartdate) Then GoTo MyNextRecord
        If Master!V_DATE > CDate(mEndDate) Then GoTo MyNextRecord
        
        Call MoveRec
        
        For I = 0 To Txt.Count - 1
            Txt(I).Refresh
        Next
        
        Call TopCtrl1_eEdit
        If Txt(GenSurAmt).Enabled = True Then Txt(GenSurAmt).SetFocus
'        For I = 1 To FGrid.Rows - 1
'            FGrid.Col = Col_TaxPer
'            FGrid.Row = I
'            FGrid_Click
'            FGrid_DblClick
'            Call TxtGrid_GotFocus(0)
'            Call TxtGrid_Validate(0, False)
'
'            FGrid.Col = Col_DiscPer
'            FGrid.Row = I
'            FGrid_Click
'            FGrid_DblClick
'            Call TxtGrid_GotFocus(0)
'            Call TxtGrid_Validate(0, False)
'        Next
'        Call Txt_Validate(STaxAmt, False)
        
        Call TopCtrl1_eSave
MyNextRecord:
        Master.MoveNext
    Loop
    
    
    
    mRePostCounter = 0
    
    
    Master.MoveFirst
    Do Until Master.EOF
        If IsNull(Master!V_DATE) Then GoTo MyNextRecord
        If Master!V_DATE < CDate(mStartdate) Then GoTo MyNextRecord1
        If Master!V_DATE > CDate(mEndDate) Then GoTo MyNextRecord1
        
        Call MoveRec
        
        For I = 0 To Txt.Count - 1
            Txt(I).Refresh
        Next
        Call TopCtrl1_eEdit
        If Txt(GenSurAmt).Enabled = True Then Txt(GenSurAmt).SetFocus
'        For i = 1 To FGrid.Rows - 1
'            FGrid.Col = Col_TaxPer
'            FGrid.Row = i
'            FGrid_Click
'            FGrid_DblClick
'            Call TxtGrid_GotFocus(0)
'            Call TxtGrid_Validate(0, False)
'
'            FGrid.Col = Col_DiscPer
'            FGrid.Row = i
'            FGrid_Click
'            FGrid_DblClick
'            Call TxtGrid_GotFocus(0)
'            Call TxtGrid_Validate(0, False)
'        Next
        'Call Txt_Validate(STaxAmt, False)
        
        Call TopCtrl1_eSave
MyNextRecord1:
        Master.MoveNext
    Loop
    
    MsgBox "Re-Posting Completed", vbInformation
    mReposting = False
    Unload Me
End If
End Sub

Private Sub CmdTransPost_Click()


Master.MoveFirst
    Do Until Master.EOF
        Call MoveRec
        Disp_Text SETS("EDIT", Me, Master)
        'txt(Vdate).SetFocus
        FGrid.AddItem FGrid.Rows
        Amt_Cal
        mAddFlag = "E"
        TopCtrl1_eSave
        FrmPrn.Visible = False
        Me.Refresh
MyNextRecord:
        Master.MoveNext
    Loop
    
End Sub

Private Sub DGParty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
Dim lblstring As String, Bal As Double
If DGParty.Row >= 0 Then
    lblstring = G_FaCn.Execute("Select AcGroup.GroupName from (AcGroup Left Join SubGroup on SubGroup.GroupCode=AcGroup.GroupCode) where SubGroup.SubCode='" & RsParty!Code & "'").Fields(0).Value
    lblGroup = lblstring
    Bal = Abs(VNull(G_FaCn.Execute("Select Sum(AmtDr)-Sum(AmtCr) from Ledger where SubCode='" & RsParty!Code & "'").Fields(0).Value))
    If Bal > 0 Then
        lblGroup.TEXT = lblstring & "  |  " & Format(Bal, "0.00") & " Dr.  "
    ElseIf Bal < 0 Then
        lblGroup.TEXT = lblstring & "  |  " & Format(Bal, "0.00") & " Cr.  "
    End If
    lblGroup.Refresh
End If
End Sub
Private Sub FillChallan()
Dim Rst As ADODB.Recordset
Dim I As Integer, j As Integer, Cnt As Integer
    FGrid.Redraw = False
    FGrid.Rows = 1
    For I = 1 To FGridSel.Rows - 1
        If FGridSel.TextMatrix(I, SCol_SrNo) <> "" Then
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select P.Part_Name ,P.Local_Name ,P.Unit ,P.MRP ,P.MRP_Effect_Dt ,P.TB_SRate ,P.TP_SRate ,P.TB_Effect_Dt ,P.Part_Grade, P.Cur_MRP_TBStk, P.Cur_MRP_TPStk, P.Cur_TB_Stk, P.Cur_TP_Stk, P.Bin_Loca, P.High_Pur_Rate, P.Low_Pur_Rate,Godown.God_Name, " & cTrim(cMID("SP_Stock.DocID", "9", "5")) & "+" & cCStr(cTrim("Right(SP_Stock.DocID,8)")) & " As ChallIDDisp, " & cTrim(cMID("SP_Stock.Order_DocID", "9", "5")) & "+ " & cCStr(cTrim("Right(SP_Stock.Order_DocID,8)")) & " As OrderIDDisp,SP_Stock.* From (SP_Stock Left Join Part P On SP_Stock.Part_No=P.Part_No and P.Div_Code = left(SP_Stock.Docid,1)) Left Join Godown on SP_Stock.Godown=Godown.God_Code Where Sp_Stock.docId='" & FGridSel.TextMatrix(I, SCol_ChalNoCode) & "'", GCn, adOpenStatic, adLockReadOnly
            If Rst.RecordCount > 0 Then
                FillChallanFlag = 1
                Do Until Rst.EOF
'                    Cnt = Cnt + 1
                    j = j + 1
                                 '|0 Col_SrNo |1 Col_PNo |2 Col_ChalNoCode |3 Col_ChalNo |4 Col_ChalSrNo |5 Col_SONoCode |6 Col_SONo |7 Col_SOSrNo |8 Col_Unit |9 Col_MRP |10 Col_Taxable |11 Col_Qty |12 Col_Rate |13 Col_MRPRate |14 Col_Amt |15 Col_DiscPer |16 Col_DiscAmt |17 Col_ItemVal |18 Col_GodownCode |19 Col_Godown |20 Col_PName |21 Col_LName |22 Col_MRPStkTB |23 Col_MRPStkTP |24 Col_TBStk |25 Col_TPStk |26 Col_TBRate |27 Col_TPRate |28 Col_Bin |29 Col_LastRate |30 Col_HPRate |31 Col_LPRate |32 Col_PartGrade |33 Col_EffectDate
'                    FGrid.AddItem Cnt & Chr(9) & Rst!order_docid & Chr(9) & Rst!OrderIDDisp & Chr(9) & Rst!Order_Srl_No & Chr(9) & Rst!Part_No & Chr(9) & Rst!DocId & Chr(9) & Rst!ChallIDDisp & Chr(9) & Rst!Srl_No & Chr(9) & Rst!Unit & Chr(9) & IIf(Rst!MRP_YN = 0, "No", "Yes") & Chr(9) & IIf(Rst!Tax_YN = 0, "No", "Yes") & Chr(9) & Format(Rst!Qty_Iss, "0.000") & Chr(9) & Format(Rst!Rate, "0.00") & Chr(9) & Format(Rst!MRP_Rate, "0.00") & Chr(9) & Format((Rst!Qty_Iss * Rst!Rate), "0.00") & Chr(9) & Format(Rst!Disc_Per, "0.00") & Chr(9) & Format(Rst!Disc_Amt, "0.00") & Chr(9) & Format(Rst!Net_Amt, "0.00") & Chr(9) & Rst!Godown & Chr(9) & Rst!God_Name & Chr(9) & Rst!Part_Name & Chr(9) & Rst!Local_Name & Chr(9) & Rst!Curstk & Chr(9) & Rst!MRPQty & Chr(9) & Rst!Cur_TB_Stk & Chr(9) & Rst!Cur_TP_Stk & Chr(9) & Rst!TB_SRate & Chr(9) & Rst!TP_SRate & Chr(9) & Rst!Bin_Loca & Chr(9) & " " & Chr(9) & Rst!high_pur_rate & Chr(9) & Rst!low_pur_rate & Chr(9) & Rst!Part_Grade _
                                          & Chr(9) & Format(IIf(Rst!MRP_YN = 1, Rst!MRP_Effect_Dt, Rst!TB_Effect_Dt), "dd/MMM/yyyy")
                                     '0                  1                     2                        3                       4                          5                        6                            7                      8                                9                                          10                                          11                                  12                                      13                                        14                                              15                                      16                                    17                            18                      19                    20                          21                       22                   23                      24                        25                         26                     27                     28                  29                     30                          31                          32                  33
                    FGrid.AddItem ""
                    With FGrid
                        .TextMatrix(j, Col_SrNo) = FGrid.Rows
                        .TextMatrix(j, Col_SONoCode) = Rst!Order_DocId
                        .TextMatrix(j, Col_SONo) = IIf(IsNull(Rst!OrderIDDisp), "", Rst!OrderIDDisp)
                        .TextMatrix(j, Col_PNo) = Rst!Part_No
                        .TextMatrix(j, Col_SOSrNo) = Rst!Order_Srl_No
                        .TextMatrix(j, Col_ChalNoCode) = Rst!DocID
                        .TextMatrix(j, Col_ChalNo) = IIf(IsNull(Rst!ChallIDDisp), "", Rst!ChallIDDisp)
                        .TextMatrix(j, Col_ChalSrNo) = Rst!Srl_No
                        .TextMatrix(j, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                        .TextMatrix(j, Col_MRP) = IIf(Rst!MRP_YN = 1, "Yes", "No")
                        .TextMatrix(j, Col_Taxable) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                        .TextMatrix(j, Col_Qty) = Format(Rst!Qty_Iss, "0.000")
                        .TextMatrix(j, Col_Rate) = Format(Rst!Rate, "0.0000")
                        .TextMatrix(j, Col_MRPRate) = Format(Rst!MRP_Rate, "0.0000")
                        If Rst!MRP_YN = 1 Then
                            .TextMatrix(j, Col_Amt) = Format((Rst!Qty_Iss * Rst!MRP_Rate), "0.00")
                        Else
                            .TextMatrix(j, Col_Amt) = Format((Rst!Qty_Iss * Rst!Rate), "0.00")
                        End If
                        .TextMatrix(j, Col_DiscPer) = Format(Rst!Disc_Per, "0.0000")
                        .TextMatrix(j, Col_DiscAmt) = Format(Rst!Disc_Amt, "0.00")
                        
                        If mVatYn = 1 Then
                            .TextMatrix(j, Col_TaxPer) = Format(VNull(Rst!TaxPer), "0.000")
                            .TextMatrix(j, Col_TaxAmt1) = Format(VNull(Rst!TaxAmt), "0.00")
                            .TextMatrix(j, Col_SatPer) = Format(VNull(Rst!SatPer), "0.000")
                            .TextMatrix(j, Col_SatAmt) = Format(VNull(Rst!SatAmt), "0.00")
                        End If
                        
                        .TextMatrix(j, Col_ItemVal) = Format(Rst!Net_Amt, "0.00")
                        .TextMatrix(j, Col_GodownCode) = Rst!Godown
                        .TextMatrix(j, Col_Godown) = IIf(IsNull(Rst!God_Name), "", Rst!God_Name)
                        .TextMatrix(j, Col_PName) = IIf(IsNull(Rst!Part_Name), "", Rst!Part_Name)
                        .TextMatrix(j, Col_LName) = IIf(IsNull(Rst!Local_Name), "", Rst!Local_Name)
                        .TextMatrix(j, Col_MRPStkTP) = IIf(IsNull(Rst!Cur_MRP_TPStk), "", Rst!Cur_MRP_TPStk)
                        .TextMatrix(j, Col_MRPStkTB) = IIf(IsNull(Rst!Cur_MRP_TbStk), "", Rst!Cur_MRP_TbStk)
                        .TextMatrix(j, Col_TBStk) = IIf(IsNull(Rst!Cur_TB_STk), "", Rst!Cur_TB_STk)
                        .TextMatrix(j, Col_TPStk) = IIf(IsNull(Rst!Cur_TP_Stk), "", Rst!Cur_TP_Stk)
                        .TextMatrix(j, Col_TBRate) = IIf(IsNull(Rst!TB_SRate), "", Rst!TB_SRate)
                        .TextMatrix(j, Col_TPRate) = IIf(IsNull(Rst!TP_SRate), "", Rst!TP_SRate)
                        .TextMatrix(j, Col_Bin) = IIf(IsNull(Rst!Bin_Loca), "", Rst!Bin_Loca)
                        .TextMatrix(j, Col_LastRate) = ""
                        .TextMatrix(j, Col_HPRate) = IIf(IsNull(Rst!high_pur_rate), "", Rst!high_pur_rate)
                        .TextMatrix(j, Col_LPRate) = IIf(IsNull(Rst!low_pur_rate), "", Rst!low_pur_rate)
                        .TextMatrix(j, Col_PartGrade) = IIf(IsNull(Rst!Part_Grade), "", Rst!Part_Grade)
                        .TextMatrix(j, Col_EffectDate) = Format(IIf(Rst!MRP_YN = 1, IIf(IsNull(Rst!MRP_Effect_Dt), "", Rst!MRP_Effect_Dt), IIf(IsNull(Rst!TB_Effect_Dt), "", Rst!TB_Effect_Dt)), "dd/MMM/yyyy")
                    End With
    '                If Rst!Tax_YN = 1 Then
    '                    mItemDiscTotTB = mItemDiscTotTB + Rst!Disc_Amt
    '                Else
    '                    mItemDiscTotTP = mItemDiscTotTP + Rst!Disc_Amt
    '                End If
                    Rst.MoveNext
                Loop
                Amt_Cal
                FillChallanFlag = 0
            Else
                FGridSel.AddItem FGridSel.Rows
                FGridSel.FixedRows = 1
            End If
        End If
    Next
    Set Rst = Nothing
    FGrid.FixedRows = 1
    FGrid.Redraw = True
End Sub

Private Sub MoveRec()
Dim Master1 As ADODB.Recordset
Dim Rst As ADODB.Recordset, I As Integer, mAmt As Double
Dim mItemDiscTotTB As Double, mItemDiscTotTP As Double
On Error GoTo ELoop
    
    mAddFlag = ""
    FrmDetail.Visible = False
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
    If Master.RecordCount > 0 Then
        If InStr(Me.TopCtrl1.Tag, "E") <> 0 Then Me.TopCtrl1.tEdit = True
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "Select S.*,SubGroup.Name As PartyName,SubGroup.Curr_Bal,SubGroup.Party_Type,SubGroupCr.Name As CreditAcName, " _
            & "TaxForms.Form_Desc As FormName,TaxForms31.Form_Desc As Form31Name,Emp.Emp_Name " _
            & "From (((((SP_Sale S Left Join SubGroup on S.Party_Code=SubGroup.SubCode) " _
            & "Left Join SubGroup SubGroupCr on S.CrAc=SubGroupCr.SubCode) " _
            & "Left Join TaxForms on S.Form_Code=TaxForms.Form_Code) " _
            & "Left Join TaxForms TaxForms31 on S.RoadPermit_FormCode=TaxForms31.Form_Code) " _
            & "Left Join Emp_Mast Emp on S.Rep_Code=Emp.Emp_Code) " _
            & "Where S.DocID = '" & Master!SearchCode & "' " _
            & "Order by S.V_Date,S.V_Type", GCn, adOpenStatic, adLockReadOnly
            
        If UCase(left(PubComp_Name, 7)) = "SOCIETY" Then
            If AllowEditDel(pubUName, Master1!V_DATE, PubLoginDate) = False Then
                TopCtrl1.tDel = False
                TopCtrl1.tEdit = False
            Else
                TopCtrl1.tDel = True
                TopCtrl1.tEdit = True
            End If
        End If
    
            
        If Master1!CancelYN = 1 Then
            TopCtrl1.tEdit = False
            LblCancel.Visible = True
        Else
            LblCancel.Visible = False
        End If
        Txt(DocID).TEXT = Master1!DocID
        Txt(DocID).Tag = Master1!DocID
        mSearchCode = Txt(DocID)
        LblDiv.CAPTION = "Division : " & left(Master1!DocID, 1)
        LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        LblUser = IIf(Not IsNull(Master1!AddDate), "Add By : " & XNull(Master1!AddBy) & "  Dated : " & XNull(Master1!AddDate), "") & IIf(Not IsNull(Master1!ModifyDate), "     Modify By : " & XNull(Master1!ModifyBy) & "  Dated : " & XNull(Master1!ModifyDate), "")
        mVType = Master1!V_Type
        If mVType = SalCashVType Then
            Txt(DocType).TEXT = "Cash"
            Txt(PType).Visible = True
            Lbl(4).Visible = True
        ElseIf mVType = SalCrVType Then
            Txt(DocType).TEXT = "Credit"
            Txt(PType).Visible = False
            Lbl(4).Visible = False
        End If
        Txt(VDate).TEXT = Master1!V_DATE
        
        mVatYn = PubVATYN
        If CDate(Master1!V_DATE) < CDate("1/Jan/2008") And StrCmp(left(PubComp_Name, 3), "Jmk") Then
            mVatYn = 0
        End If
        
        If PubBackEnd = "A" Then
            mSatYn = IIf(VNull(Master1!SAT_YN) = 1, True, False)
        Else
            mSatYn = IIf(VNull(Master1!SAT_YN) = True, True, False)
        End If
        
        DispTextVat
        
        
        LblVPrefix.CAPTION = mID(Master1!DocID, 9, 5)
        lblGatePass.CAPTION = IIf(IsNull(Master1!gp_no), "", Master1!gp_no)
        Txt(SerialNo).TEXT = Master1!V_NO
        Txt(Party).Tag = Master1!Party_code
        RsParty.MoveFirst
        RsParty.FIND ("Code ='" & Txt(Party).Tag & "'")
        If Master1!Cash_Credit = "Cash" Then
            Txt(Party).TEXT = Master1!Party_Name
        Else
            Txt(Party).TEXT = IIf(IsNull(Master1!PartyName), "", Master1!PartyName)
        End If
        LblCurrBal = "Bal. " & Format(Abs(Master1!Curr_Bal), "0.00")
        LblCurrBal = LblCurrBal & IIf(Master1!Curr_Bal > 0, " Cr", IIf(Master1!Curr_Bal < 0, " Dr", ""))

        Txt(Address1).TEXT = XNull(Master1!Address)
        mPartyType = IIf(IsNull(Master1!Party_Type), 0, Master1!Party_Type)
        Txt(CrAc).Tag = XNull(Master1!CrAc)
        Txt(CrAc).TEXT = IIf(IsNull(Master1!CreditAcName), "", Master1!CreditAcName)
        Txt(FormName).Tag = Master1!Form_Code
        Txt(PType) = XNull(Master1!PType)
        Txt(FormName).TEXT = IIf(IsNull(Master1!FormName), "", Master1!FormName)
        Txt(Form31Name).Tag = XNull(Master1!RoadPermit_FormCode)
        Txt(Form31Name).TEXT = IIf(IsNull(Master1!Form31Name), "", Master1!Form31Name)
        Txt(Form31No).TEXT = XNull(Master1!RoadPermit_No)
        
        Txt(SPerson).TEXT = XNull(Master1!REP_CODE)
        Txt(SPerson).TEXT = IIf(IsNull(Master1!Emp_Name), "", Master1!Emp_Name)
        Txt(Remark).TEXT = XNull(Master1!Remarks)
        If Master1!L_C = "L" Then
            Txt(LC).TEXT = "Local"
        ElseIf Master1!L_C = "C" Then
            Txt(LC).TEXT = "Central"
        End If
        Txt(DispMode).TEXT = XNull(Master1!Mode_Dispatch)
        Txt(Transport).TEXT = XNull(Master1!Transport)
        Txt(LRNo).TEXT = XNull(Master1!GR_RR_No)
        Txt(LRDate).TEXT = IIf(IsNull(Master1!GR_RR_Date), "", Master1!GR_RR_Date)
        If Master1!Det_Tax = 0 Then
            Txt(TaxDet).TEXT = "No"
        ElseIf Master1!Det_Tax = 1 Then
            Txt(TaxDet).TEXT = "Yes"
        End If
        Txt(CaseNo).TEXT = XNull(Master1!Case_No)
        Txt(CaseMark).TEXT = XNull(Master1!Case_Mark)

        Txt(MRPAmtTB).TEXT = Format(Master1!SprAmt_MRP_TB + Master1!OilAmt_MRP_TB, "0.00")
        Txt(MRPAmtTP).TEXT = Format(Master1!SprAmt_MRP_TP + Master1!OilAmt_MRP_TP, "0.00")
        mMRPLubeTB = Master1!OilAmt_MRP_TB
        mMRPLubeTP = Master1!OilAmt_MRP_TP
        Txt(SprAmtTB).TEXT = Format(Master1!SprAmt_TB, "0.00")
        Txt(SprAmtTP).TEXT = Format(Master1!SprAmt_TP, "0.00")
        Txt(OilAmtTB).TEXT = Format(Master1!OilAmt_TB, "0.00")
        Txt(OilAmtTP).TEXT = Format(Master1!OilAmt_TP, "0.00")
        Txt(DiscPerTB).TEXT = Format(Master1!D_Per_TB, "0.0000")
        Txt(DiscAmtTB).TEXT = Format(Master1!D_Amt_TB, "0.00")
        Txt(DiscPerTP).TEXT = Format(Master1!D_Per_TP, "0.0000")
        Txt(DiscAmtTP).TEXT = Format(Master1!D_Amt_TP, "0.00")
        Txt(STotATB).TEXT = Format((Master1!SprAmt_MRP_TB + Master1!OilAmt_MRP_TB + Master1!SprAmt_TB + Master1!OilAmt_TB) - Master1!D_Amt_TB, "0.00")
        Txt(STotATP).TEXT = Format((Master1!SprAmt_MRP_TP + Master1!OilAmt_MRP_TP + Master1!SprAmt_TP + Master1!OilAmt_TP) - Master1!D_Amt_TP, "0.00")
'        Txt(Addition).Text = Format(Master1!Addition, "0.00")
        Txt(GenSurPer).TEXT = Format(Master1!Gen_Sur_Per, "0.00")
        Txt(GenSurAmt).TEXT = Format(Master1!Gen_Sur_Amt, "0.00")
        Txt(TransAmt).TEXT = Format(Master1!Trans_Amt, "0.00")
'        Txt(TaxableTot) = Format(Val(Txt(STotATB)) + Val(Txt(Addition)) + Val(Txt(PackCrg)) + Val(Txt(GenSurAmt)) + Val(Txt(TransAmt)), "0.00")
        Txt(TaxableTot) = Format(Val(Txt(STotATB)) + Val(Txt(GenSurAmt)) + Val(Txt(TransAmt)), "0.00")
        Txt(STaxPer).TEXT = Format(Master1!Tax_Per, "0.00")
        Txt(STaxAmt).TEXT = Format(Master1!Tax_Amt, "0.00")
        Txt(TaxSurPer).TEXT = Format(Master1!Tax_Sur_Per, "0.00")
        Txt(TaxSurAmt).TEXT = Format(Master1!Tax_Sur_Amt, "0.00")
        Txt(PackCrg).TEXT = Format(Master1!Packing, "0.00")
'        Txt(STotB) = Format(Val(Txt(TaxableTot)) + Val(Txt(STaxAmt)) + Val(Txt(TaxSurAmt)), "0.00")
        Txt(STotB) = Format(Val(Txt(STotATP)) + Val(Txt(TaxableTot)) + Val(Txt(PackCrg)) + Val(Txt(STaxAmt)) + Val(Txt(TaxSurAmt)), "0.00")
        Txt(TurnOverPer).TEXT = Format(Master1!TOT_Per, "0.00")
        Txt(TurnOverAmt).TEXT = Format(Master1!Tot_Amt, "0.00")
        Txt(ReSalTaxPer).TEXT = Format(Master1!ReSalTax_Per, "0.00")
        Txt(ReSalTaxAmt).TEXT = Format(Master1!ReSalTax_Amt, "0.00")
        Txt(SROff).TEXT = Format(Master1!Rounded, "0.00")
'        Txt(NetSprAmt) = Format(Val(Txt(STotB)) + Val(Txt(STotATP)) + Val(Txt(TurnOverAmt)) + Val(Txt(SROff)), "0.00")
        Txt(NetSprAmt) = Format(Val(Txt(STotB)) + Val(Txt(TurnOverAmt)) + Val(Txt(ReSalTaxAmt)) + Val(Txt(SROff)), "0.00")
        Txt(NetAmt).TEXT = Format(Master1!Total_Amt, "0.00")

        mTBDisAmtMRP = Master1!D_Amt_MRP_TB
        mTPDisAmtMRP = Master1!D_Amt_MRP_TP
        mMRPTax = Master1!Tax_AmtMRP
        mMRPTaxSur = Master1!TaxSur_AmtMRP
        mMRPTOT = Master1!Tot_AmtMrp

        FGrid.Rows = 1
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select P.Part_Name,P.Local_Name,P.Unit,P.MRP,P.MRP_Effect_Dt,P.TB_SRate,P.TP_SRate,P.TB_Effect_Dt,P.Part_Grade ,P.Cur_MRP_TBStk,P.Cur_MRP_TPStk,P.Cur_TB_Stk ,P.Cur_TP_Stk ,P.Bin_Loca ,P.High_Pur_Rate, P.Low_Pur_Rate, Godown.God_Name," & cTrim(cMID("SP_Stock.DocID", "9", "5")) & " + " & cCStr(cTrim("Right(SP_Stock.DocID,8)")) & " As ChallIDDisp, " & cTrim(cMID("SP_Stock.Order_DocID", "9", "5")) & " + " & cCStr(cTrim("Right(SP_Stock.Order_DocID,8)")) & " As OrderIDDisp,SP_Stock.* " & _
            " From (SP_Stock Left Join Part P On SP_Stock.Part_No=P.Part_No and P.Div_Code = left(SP_Stock.Docid,1)) " & _
            " Left Join Godown on SP_Stock.Godown=Godown.God_Code " & _
            " Where SP_Stock.Invoice_DocId='" & Master1!DocID & "' and SP_Stock.V_Type='" & SalChalType & "'", GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount > 0 Then
            I = 1
            FGrid.Redraw = False
            Do Until Rst.EOF
            '|0 Col_SrNo |1 Col_SONoCode |2 Col_SONo |3 Col_SOSrNo | 4 Col_PNo |5 Col_ChalNoCode |6 Col_ChalNo |7 Col_ChalSrNo |8 Col_Unit |9 Col_MRP |10 Col_Taxable |11 Col_Qty |12 Col_Rate |13 Col_MRPRate |14 Col_Amt |15 Col_DiscPer |16 Col_DiscAmt |17 Col_ItemVal |18 Col_GodownCode |19 Col_Godown |20 Col_PName |21 Col_LName |22 Col_MRPStkTB |23 Col_MRPStkTP |24 Col_TBStk |25 Col_TPStk |26 Col_TBRate |27 Col_TPRate |28 Col_Bin |29 Col_LastRate |30 Col_HPRate |31 Col_LPRate |32 Col_PartGrade |33 Col_EffectDate
                If Rst!MRP_YN = 1 Then
                    mAmt = (Rst!Qty_Iss * Rst!MRP_Rate2)
                Else
                    mAmt = (Rst!Qty_Iss * Rst!Rate2)
                End If
'                FGrid.AddItem i & Chr(9) & Rst!order_docid & Chr(9) & Rst!OrderIDDisp & Chr(9) & Rst!Order_Srl_No & Chr(9) & Rst!Part_No & Chr(9) & Rst!DocId & Chr(9) & Rst!ChallIDDisp & Chr(9) & Rst!Srl_No & Chr(9) & Rst!Unit & Chr(9) & Rst!MRPYN & Chr(9) & Rst!TaxYN & Chr(9) & Format(Rst!Qty_iss, "0.000") & Chr(9) & Format(Rst!Rate2, "0.00") & Chr(9) & Format(Rst!MRP_Rate2, "0.00") & Chr(9) & Format(mAmt, "0.00") & Chr(9) & Format(Rst!Disc_Per2, "0.00") & Chr(9) & Format(Rst!Disc_Amt2, "0.00") & Chr(9) & Format(Rst!Net_Amt2, "0.00") & Chr(9) & Rst!Godown & Chr(9) & Rst!God_Name & Chr(9) & Rst!Part_Name & Chr(9) & Rst!Local_Name & Chr(9) & Rst!Curstk & Chr(9) & Rst!MRPQty & Chr(9) & Rst!Cur_TB_Stk & Chr(9) & Rst!Cur_TP_Stk & Chr(9) & Rst!TB_SRate & Chr(9) & Rst!TP_SRate & Chr(9) & Rst!Bin_Loca & Chr(9) & " " & Chr(9) & Rst!high_pur_rate & Chr(9) & Rst!low_pur_rate & Chr(9) & Rst!Part_Grade & Chr(9) & Format(IIf(Rst!MRPYN = "Yes", Rst!MRP_Effect_Dt, Rst!TB_Effect_Dt), "dd/MMM/yyyy")
                             '0                  1                     2                        3                      4                         5                       6                            7                        8                   9                    10                              11                                12                                      13                                           14                                               15                                        16                                     17                              18                       19                    20                        21                      22                    23                      24                        25                        26                      27                      28                29                     30                           31                         32                                                              33
                FGrid.AddItem ""
                CustOrdDet = ""
                If Rst!Order_DocId <> "" Then
                        CustOrdDet = GCn.Execute("Select SPO.CustOrd_Det from SP_Order SPO where OrderId='" & Rst!Order_DocId & "'").Fields(0).Value
                End If
                With FGrid
                    .TextMatrix(I, Col_SrNo) = I
                    .TextMatrix(I, Col_SONoCode) = XNull(Rst!Order_DocId)
                    .TextMatrix(I, Col_SONo) = IIf(IsNull(Rst!OrderIDDisp), "", Rst!OrderIDDisp)
                    .TextMatrix(I, Col_SOSrNo) = XNull(Rst!Order_Srl_No)
                    .TextMatrix(I, Col_PNo) = Rst!Part_No
                    .TextMatrix(I, Col_ChalNoCode) = Rst!DocID
                    .TextMatrix(I, Col_ChalNo) = Rst!ChallIDDisp
                    .TextMatrix(I, Col_ChalSrNo) = Rst!Srl_No
                    .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                    .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Qty) = Format(Rst!Qty_Iss, "0.000")
                    .TextMatrix(I, Col_Rate) = Format(Rst!Rate2, "0.0000")
                    .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP_Rate2, "0.0000")
                    .TextMatrix(I, Col_Amt) = Format(mAmt, "0.00")
                    .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per2, "0.0000")
                    .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt2, "0.00")
                    
                    If mVatYn = 1 Then
                        .TextMatrix(I, Col_TaxPer) = Format(Rst!TaxPer, "0.000")
                        .TextMatrix(I, Col_TaxAmt1) = Format(Rst!TaxAmt, "0.00")
                        
                        If mSatYn Then
                            .TextMatrix(I, Col_SatPer) = Format(Rst!SatPer, "0.000")
                            .TextMatrix(I, Col_SatAmt) = Format(Rst!SatAmt, "0.00")
                        End If
                    End If
                    
                    
                    
                    .TextMatrix(I, Col_ItemVal) = Format(Rst!Net_Amt2, "0.00")
                    .TextMatrix(I, Col_GodownCode) = Rst!Godown
                    .TextMatrix(I, Col_Godown) = IIf(IsNull(Rst!God_Name), "", Rst!God_Name)
                    .TextMatrix(I, Col_PartSrlNo) = IIf(IsNull(Rst!Part_SrlNo), "", Rst!Part_SrlNo)
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
                    .TextMatrix(I, Col_EffectDate) = Format(IIf(Rst!MRP_YN = 1, IIf(IsNull(Rst!MRP_Effect_Dt), "", Rst!MRP_Effect_Dt), IIf(IsNull(Rst!TB_Effect_Dt), "", Rst!TB_Effect_Dt)), "dd/MMM/yyyy")
                    .TextMatrix(I, Col_PurDocId) = XNull(Rst!PurDocNo)
                    .TextMatrix(I, Col_PurDate) = XNull(Rst!PurDocDate)
                End With
                If Rst!Tax_YN = 1 Then
                    mItemDiscTotTB = mItemDiscTotTB + Rst!Disc_Amt2
                Else
                    mItemDiscTotTP = mItemDiscTotTP + Rst!Disc_Amt2
                End If
                Rst.MoveNext
                I = I + 1
            Loop
            Txt(IWDiscTotTB).TEXT = Format(mItemDiscTotTB, "0.00")
            Txt(IWDiscTotTP).TEXT = Format(mItemDiscTotTP, "0.00")
            FGrid.FixedRows = 1
            FGrid.Redraw = True
            CountItem
        Else
            FGrid.AddItem FGrid.Rows
            FGrid.FixedRows = 1
        End If
    Else
        BlankText
    End If
    Grid_Hide
    Amt_Cal
Set Rst = Nothing
Set Master1 = Nothing
Txt(STaxPer).Enabled = False
Txt(STaxAmt).Enabled = False
Txt(TaxSurPer).Enabled = False
Txt(TaxSurAmt).Enabled = False
Txt(GenSurPer).Enabled = False
Txt(GenSurAmt).Enabled = False
Txt(DiscAmtTB).Enabled = False
Txt(DiscAmtTP).Enabled = False
Txt(DiscPerTB).Enabled = False
Txt(DiscPerTP).Enabled = False
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
            TxtGrid(0).SetFocus
            ChkDuplicate = False
            Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
    Select Case FGrid.Col
        Case Col_SONo  ' OptChal(ChalCreate).Value = True  Create Challan
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            TxtGridValid_SONo
            FillOrderDetail
        Case Col_PNo, Col_PName, Col_LName ' OptChal(ChalCreate).Value = True  Create Challan
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            TxtGridValid_PNo
        Case Col_Taxable, Col_MRP
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            TxtGridValid_TaxMRP
            If mVatYn = 1 Then
                If Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) = 0 And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" Then
                    MsgBox "Wrong Attempt ! Tax Percentage is Zero.Select the Appropriate Tax Form.", vbOKOnly
                    Exit Function
                End If
            Else
                If Val(Txt(STaxPer)) = 0 And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" Then
                    MsgBox "Wrong Attempt ! Tax Percentage is Zero.Select the Appropriate Tax Form.", vbOKOnly
                    Exit Function
                End If
            End If
        Case Col_Rate, Col_DiscPer, Col_TaxPer, Col_SatPer
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.0000")
            Amt_Cal
        Case Col_DiscAmt
            If Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) < Val(TxtGrid(0)) Then
                MsgBox "Item-wsie Disc. Amount is greater than Item Value", vbOKOnly, "Item-wise Disc. Checking"
                TxtGridLeave = False: Exit Function
            End If
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
            Amt_Cal
        Case Col_Qty
            FGrid.TextMatrix(FGrid.Row, Col_Qty) = Format(Val(TxtGrid(0).TEXT), "0.000")
            If CheckSprStock(FGrid, FGrid.Row, Col_MRP, Col_Taxable, Col_Qty, Col_MRPStkTB, Col_MRPStkTP, Col_TBStk, Col_TPStk) = False Then
                TxtGrid(0) = "": TxtGrid(0).SetFocus: TxtGridLeave = False: Exit Function
            End If
            If Val(RsPart!ReOrd_Lvl) > 0 And Val(RsPart!Min_Lvl) > 0 Then
                CurrStk = GCn.Execute("Select " & vIsNull("Cur_MRP_TBStk", "0") & " + " & vIsNull("Cur_MRP_TPStk", "0") & " + " & vIsNull("Cur_TB_STk", "0") & " + " & vIsNull("Cur_Tp_Stk", "0") & " from Part where Part_No='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "' And Div_Code='" & PubDivCode & "' ").Fields(0).Value
                If CurrStk - Val(FGrid.TextMatrix(FGrid.Row, Col_Qty)) < Val(RsPart!ReOrd_Lvl) Then
                    MsgBox "Stock for this Part is Below Reorder Level", vbInformation, App.Title & "[validation Check]"
                End If
            
                If CurrStk - Val(FGrid.TextMatrix(FGrid.Row, Col_Qty)) < Val(RsPart!Min_Lvl) Then
                    MsgBox "Stock for this Part is Below Minimum Level", vbInformation, App.Title & "[validation Check]"
                End If
            End If
            Amt_Cal
            If RsGodown.RecordCount > 0 And Trim(FGrid.TextMatrix(FGrid.Row, Col_Godown)) = "" Then
                RsGodown.MoveFirst
                RsGodown.FIND "Code ='" & PubSprCounterGodown & "'"
                FGrid.TextMatrix(FGrid.Row, Col_GodownCode) = RsGodown!Code
                FGrid.TextMatrix(FGrid.Row, Col_Godown) = RsGodown!Name
            End If
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
'Dim TotMRPItemAmtTB As Double, TotMRPItemAmtTP As Double
'Dim TotSprAmtTB As Double, TotSprAmtTP As Double
'Dim TotOilAmtTB As Double, TotOilAmtTP As Double
'Dim TotMRPItemDisTB As Double, TotMRPItemDisTP As Double
' To Change
'PubPartGrade_Lub = "O"
'---
 Dim mAmount As Double, TaxAmt As Double, DisAmt As Double, OrdDisAmt1 As Double
 Dim TTaxAmt As Double
 Dim mTaxableAmt As Double
    If FillChallanFlag <> 1 Then
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

        FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = Format(((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) * Val(FGrid.TextMatrix(FGrid.Row, Col_DiscPer))) / 100), "0.00")
        FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) - Val(FGrid.TextMatrix(FGrid.Row, Col_DiscAmt))), "0.00")
    End If
    '******************** For Tax in Line File *************************
    If mVatYn = 1 Then
       If Txt(FormName).Tag <> "" Then
            If FGrid.TextMatrix(FGrid.Row, Col_TaxPer) <> "" Then
                mAmount = Val(FGrid.TextMatrix(FGrid.Row, Col_Amt))
                DisAmt = Val(FGrid.TextMatrix(FGrid.Row, Col_DiscAmt))
                If FGrid.TextMatrix(FGrid.Row, Col_MRP) = "Yes" And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" Then
                    If mSatYn Then
                        mTaxableAmt = Format((mAmount - DisAmt) * 100 / (100 + Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) + Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer))), "0.00")
                        FGrid.TextMatrix(FGrid.Row, Col_TaxAmt1) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / 100, "0.00")
                        FGrid.TextMatrix(FGrid.Row, Col_SatAmt) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer)) / 100, "0.00")
                    Else
                        FGrid.TextMatrix(FGrid.Row, Col_TaxAmt1) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / (100 + Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer))), "0.00")
                    End If
                    FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_ItemVal)) - Val(FGrid.TextMatrix(FGrid.Row, Col_TaxAmt1)) - Val(FGrid.TextMatrix(FGrid.Row, Col_SatAmt)), "0.00")
                ElseIf FGrid.TextMatrix(FGrid.Row, Col_MRP) = "No" And FGrid.TextMatrix(FGrid.Row, Col_Taxable) = "Yes" Then
                    FGrid.TextMatrix(FGrid.Row, Col_TaxAmt1) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_TaxPer)) / 100, "0.00")
                    If mSatYn Then
                        FGrid.TextMatrix(FGrid.Row, Col_SatAmt) = Format((mAmount - DisAmt) * Val(FGrid.TextMatrix(FGrid.Row, Col_SatPer)) / 100, "0.00")
                    Else
                        FGrid.TextMatrix(FGrid.Row, Col_SatAmt) = 0
                    End If
                Else
                    FGrid.TextMatrix(FGrid.Row, Col_TaxAmt1) = ""
                    FGrid.TextMatrix(FGrid.Row, Col_SatAmt) = 0
                End If
            End If
       End If
       
    End If
    '*******************************************************************
    MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
            Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
            Val(Txt(DiscPerTB)), Val(Txt(DiscPerTP)), _
            Val(Txt(STaxPer)), Val(Txt(TaxSurPer)), Val(Txt(TurnOverPer))
            
    If mVatYn = 1 Then
       MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
            Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
            Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
            Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
            Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
            Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
            Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, Txt(SatAmt)
    Else
        MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
            Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
            Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
            Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
            Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
            Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
            Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
    End If
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
' Used For Enable/Disable Controls in case of Create/Select Challan Option
Private Sub CtrlEnbChallan(Enb As Boolean)
    Txt(SPerson).Enabled = Enb
    Txt(CaseNo).Enabled = Enb
    Txt(CaseMark).Enabled = Enb
    Txt(DispMode).Enabled = Enb
    Txt(Transport).Enabled = Enb
    Txt(LRNo).Enabled = Enb
    Txt(LRDate).Enabled = Enb
End Sub
' Used For Updation of Sale Order in case of Edit and Delete of Challan
Private Sub UpdateSO(ChalDocID As String)
Dim Rst As ADODB.Recordset, I As Byte
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select order_docid,Order_Srl_No,Qty_iss From SP_Stock Where Invoice_DocId='" & ChalDocID & "'", GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        While Not Rst.EOF
            If Rst!Order_DocId <> "" Then
                GCn.Execute "Update SP_Order1 Set Sup_Qty=Sup_Qty-" & Rst!Qty_Iss & " Where OrderId='" & Rst!Order_DocId & "' and Srl_No=" & Rst!Order_Srl_No & ""
            End If
            Rst.MoveNext
        Wend
    End If
Set Rst = Nothing
End Sub

Private Sub DGOrdPart_dblClick()
FGrid.TextMatrix(FGrid.Row, Col_SOSrNo) = GRs!Srl_No
Set GRs = Nothing
FGrid.SetFocus
DGOrdPart.Visible = False
End Sub

Private Sub DGParty_Click()
On Error GoTo ELoop
    If RsParty.RecordCount > 0 Then
        Txt(Party).TEXT = RsParty!Name
        Txt(Party).Tag = RsParty!Code
        Txt(Address1).TEXT = RsParty!Add1
        If Txt(Transport) = "" Then
            Txt(Transport).TEXT = IIf(IsNull(RsParty!Transporter), "", RsParty!Transporter)
        End If
    End If
    Txt(Party).SetFocus
    DGParty.Visible = False
    lblGroup.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGCrAc_Click()
On Error GoTo ELoop
    If RsCrAc.RecordCount > 0 Then
        Txt(CrAc).TEXT = RsCrAc!Name
        Txt(CrAc).Tag = RsCrAc!Code
    End If
    Txt(CrAc).SetFocus
    DGCrAc.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGForm_Click()
On Error GoTo ELoop
    If rsForm.RecordCount > 0 Then
        Txt(FormName).TEXT = rsForm!Name
        Txt(FormName).Tag = rsForm!Code
        If TopCtrl1.TopText2.CAPTION = "Add" Then   ' To Assign Tax% in case of Add
            Txt(STaxPer).TEXT = IIf(Val(rsForm!Tax_Per) = 0, "", Format(rsForm!Tax_Per, "0.00"))
            Txt(TaxSurPer).TEXT = IIf(Val(rsForm!Tax_Sur_Per) = 0, "", Format(rsForm!Tax_Sur_Per, "0.00"))
            Amt_Cal
        End If
    End If
    Txt(FormName).SetFocus
    DGForm.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGForm31_Click()
On Error GoTo ELoop
    If rsForm31.RecordCount > 0 Then
        Txt(Form31Name).TEXT = rsForm31!Name
        Txt(Form31Name).Tag = rsForm31!Code
    End If
    Txt(Form31Name).SetFocus
    DGForm31.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGPerson_Click()
On Error GoTo ELoop
    If RsPerson.RecordCount > 0 Then
        Txt(SPerson).TEXT = RsPerson!Name
        Txt(SPerson).Tag = RsPerson!Code
    End If
    Txt(SPerson).SetFocus
    DGPerson.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGTrans_Click()
    If rsTrans.RecordCount > 0 Then
        Txt(Transport).TEXT = rsTrans!Name
    End If
    Txt(Transport).SetFocus
    DGTrans.Visible = False
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
    End If
    TxtGridValid_PNo
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
    DGPart.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGSONo_Click()
On Error GoTo ELoop
    If RsSONo.RecordCount > 0 Then
        TxtGrid(0).TEXT = RsSONo!Name
        FGrid.TextMatrix(FGrid.Row, Col_SONoCode) = RsSONo!Code
        FGrid.TextMatrix(FGrid.Row, Col_SONo) = RsSONo!Name
    End If
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
    DGSONo.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

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

Private Sub Form_Activate()
Dim UnLoadFrm As Boolean, MsgStr$
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
    Call TopCtrl1_eRef
End If
If rsCtrlAc.RecordCount <= 0 Then
    MsgStr = "No Records in Spare A/c Controls"
    UnLoadFrm = True
End If
If rsCtrlAc!SprSalTP_Ac = "" Or _
    rsCtrlAc!OilSalTB_Ac = "" Or rsCtrlAc!OilSalTP_Ac = "" Or _
    rsCtrlAc!SprCash_Ac = "" Or rsCtrlAc!SprDiscTB_Ac = "" Or rsCtrlAc!SprGenSur_Ac = "" Or _
    rsCtrlAc!Transportation_Ac = "" Or rsCtrlAc!ReSaleTax_Ac = "" Or _
    rsCtrlAc!MiscChrg_Ac = "" Or rsCtrlAc!TOTax_Ac = "" Or rsCtrlAc!SprROff_Ac = "" Then
    MsgStr = "Please Fill Spare"
    UnLoadFrm = True
End If
    
'EOF Spare A/c control checking
If UnLoadFrm Then
    MsgBox "Spare Sale Bill Loading Aborted !" & vbCrLf & MsgStr & " A/c Controls through Utility Menu", vbInformation, "Validation"
    Unload Me
End If
If mVatYn = 1 Then
    Lbl(22).CAPTION = "V A T   "
    Txt(39).Visible = False
End If
'Do Until Master.EOF
'        Disp_Text SETS("INI", Me, Master)
'        MoveRec
'        TopCtrl1_eEdit
'        TopCtrl1_eSave
'        Master.MoveNext
' Loop

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
TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini
    Call Ini_Pub
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg '&HDFF4F2
        Txt(I).ForeColor = CtrlFColOrg
    Next
    Lbl(35) = PubForm31Caption
    Lbl(36) = PubForm31Caption & " No."
    Lbl(25) = pubTOTCaption
    If PubReSaleTaxPer = 0 Then
        Lbl(17).Visible = False
        Txt(ReSalTaxPer).Visible = False
        Txt(ReSalTaxAmt).Visible = False
    End If
    mVType = SalCashVType
    ForSiteCode = PubSiteCode
    Txt(VDate).Tag = PubLoginDate
    
    'A/c Pstong Control Checking
    Set rsCtrlAc = New ADODB.Recordset
    rsCtrlAc.CursorLocation = adUseClient
    'CSSprAc=Temp Sale A/c
    'SprSalTB_Ac shifted to Tax Forms
    rsCtrlAc.Open "Select SprSalTP_Ac,OilSalTB_Ac,OilSalTP_Ac,CSSprAc,SprGenSur_Ac,ReSaleTax_Ac,SprCash_Ac,SprDiscTB_Ac,Transportation_Ac,MiscChrg_Ac,TOTax_Ac,SprROff_Ac From AcControls where div_Code='" & PubDivCode & "'", GCnFaS, adOpenDynamic, adLockOptimistic
    'eof checking
    
    Set DGPart.DataSource = RsPart

    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Curr_Bal,Add1,Transporter,Party_Type,City.CityName, LstNo from (SubGroup " & _
        "left Join City on City.CityCode=SubGroup.CityCode) " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
        "Where  " & _
        "(Left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') Or SubGroup.Nature='Cash')" & _
        " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty

    Set RsCrAc = New ADODB.Recordset
    RsCrAc.CursorLocation = adUseClient
    RsCrAc.Open "Select SubCode as Code,Name From SubGroup Where  Nature='Revenue' and SubGroup.AliasYN<>'Y' Order by Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGCrAc.DataSource = RsCrAc
    
    Set rsForm = New ADODB.Recordset
    rsForm.CursorLocation = adUseClient
    rsForm.Open "Select T.Form_Code as Code,T.Form_Desc As Name,T.Tax_Per,T.Tax_Sur_Per, T.AddTaxPer,T1.Tax_Ac_Code,T1.Sur_Ac_Code, T1.AddTaxAc,T1.PurSal_Ac_Code, T.L_C " & _
            "From TaxForms as T left join TaxFormsAc as T1 on T.Form_Code+'" & PubDivCode & "'=T1.Form_Code+T1.Div_Code " & _
            "Where Trn_Type='Sale' and Spare_YN=1  Order by Form_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGForm.DataSource = rsForm
    
    Set rsForm31 = New ADODB.Recordset
    rsForm31.CursorLocation = adUseClient
    rsForm31.Open "Select Form_Code as Code,Form_Desc As Name From TaxForms Where Spare_YN=1 and Trn_Type='Permit' Order by Form_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGForm31.DataSource = rsForm31

    Set rsTrans = New ADODB.Recordset
    rsTrans.CursorLocation = adUseClient
    rsTrans.Open "Select Distinct Transport as Name From  SP_Sale Where Transport<>'' Order By Transport", GCn, adOpenDynamic, adLockOptimistic
    Set DGTrans.DataSource = rsTrans
    
    Set RsPerson = New ADODB.Recordset
    RsPerson.CursorLocation = adUseClient
    RsPerson.Open "Select Emp_Code as Code, Emp_Name as Name From Emp_Mast Where Emp_Type=0 and (LeftOn Is Null or LeftOn< " & ConvertDate(PubLoginDate) & ") Order By Emp_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGPerson.DataSource = RsPerson

    Set RsSONo = New ADODB.Recordset
    RsSONo.CursorLocation = adUseClient
    RsSONo.Open "Select OrderID as Code," & cTrim(cMID("OrderID", "9", "5")) & " + " & cCStr(cTrim("Right(OrderID,8)")) & " As Name,V_Date,Qty,(Qty-Sup_Qty) As PendQty,Rate, " & cIIF("TAX_YN=1", "'Yes'", "'No'") & " As TAXYN, " & cIIF("MRP_YN=1", "'Yes'", "'No'") & " As MRPYN From SP_Order1 Where left(OrderID,1)='" & PubDivCode & "' and Order_Type='S_SO' Order By OrderID", GCn, adOpenDynamic, adLockOptimistic
    Set DGSONo.DataSource = RsSONo

    Set RsGodown = New ADODB.Recordset
    RsGodown.CursorLocation = adUseClient
    RsGodown.Open "Select God_Code as Code,God_Name As Name From Godown Where Appli_For=0 Order by God_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGodown.DataSource = RsGodown
    
     Dim sitecond As String
     sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("S.docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    If PubMoveRecYn Then
        Set Master = GCn.Execute("Select S.DocID As SearchCode,U_EntDt, V_Date From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' " & sitecond & " and S.V_Type In ('" & SalCashVType & "','" & SalCrVType & "') " _
            & "Order by S.V_Date desc,S.DocID desc")
    Else
        Set Master = GCn.Execute("Select Top 1 S.DocID As SearchCode,U_EntDt, V_Date From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' " & sitecond & " and S.V_Type In ('" & SalCashVType & "','" & SalCrVType & "') " _
            & "Order by S.V_Date desc,S.DocID desc")
    End If
    
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TopCtrl1.TopText2 <> "Browse" Then
        If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsParty = Nothing
    Set RsCrAc = Nothing
    Set rsForm = Nothing
    Set rsForm31 = Nothing
    Set rsTrans = Nothing
    Set RsSONo = Nothing
    Set RsGodown = Nothing
    Set Master = Nothing
    Set rsCtrlAc = Nothing
End Sub

Private Sub ListView_Click()
On Error GoTo ELoop
    Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    Txt(Val(ListView.Tag)).SetFocus
    FrmList.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Map1_Click()
MsgBox "Helo Mr. Vikash Verma"
End Sub

Private Sub OptChal_Click(Index As Integer)
Dim Rst As ADODB.Recordset
If OptChal(ChalSelect).Enabled = False And OptChal(ChalCreate).Enabled = False Then Exit Sub
Select Case Index
    Case 0          ' Select Challan
        If Txt(DocType).TEXT = "Credit" And Txt(Party).TEXT = "" Then
            MsgBox "Please Select Party", vbInformation, "Validation"
            Txt(Party).SetFocus
            Exit Sub
        End If
        CtrlEnbChallan False
        FrmSel.Visible = True: FrmSel.ZOrder 0: FGridSel.SetFocus

        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
'        If Txt(DocType).Text = "Credit" Then
        Rst.Open "Select Distinct S.DocID, " & cTrim(cMID("S.DocID", "9", "5")) & "+ " & cCStr(cTrim("Right(S.DocID,8)")) & " As ChallNo,S.V_Date " _
            & "From SP_Stock S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' and S.V_Type='SYSC' And S.Party_Code='" & Txt(Party).Tag & "' And (S.Invoice_DocId Is Null or S.Invoice_DocId='')", GCn, adOpenDynamic, adLockOptimistic
'        Else
'            Rst.Open "Select Distinct S.DocID, " & cTrim(cMid("S.DocID","8","5")) & "+ " & cCStr(cTrim("Right(S.DocID,8)")) &  " As ChallNo,S.V_Date " _
'            & "From SP_Stock S " _
'            & "Where S.V_Type='SYSC' And (Isnull(S.Invoice_DocId) or S.Invoice_DocId='')", GCn, adOpenDynamic, adLockOptimistic
'        End If
        FGridSel.Rows = 1
        If Rst.RecordCount > 0 Then
            Do Until Rst.EOF
                FGridSel.AddItem "" & Chr(9) & Rst!DocID & Chr(9) & Rst!ChallNo & Chr(9) & 0 & Chr(9) & Rst!V_DATE
                Rst.MoveNext
            Loop
            FGridSel.FixedRows = 1
        End If
    Case 1          ' Create Challan
        CtrlEnbChallan True
        FrmSel.Visible = False
    End Select
End Sub

Private Sub OptChal_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeysA vbKeyTab, True
    
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
Dim RsTemp As ADODB.Recordset
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    mAddFlag = "A"
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    CtrlEnbChallan True
    Txt(VDate).TEXT = Txt(VDate).Tag
    Txt(DocType).TEXT = "Credit"
    Txt(LC).TEXT = "Local"
    If PubTaxDetOnSprInv = 1 Then
        Txt(TaxDet).TEXT = "Yes"
    Else
        Txt(TaxDet).TEXT = "No"
    End If
    mSatYn = IIf(PubSatYn = 1, True, False)
    
    DispTextVat
    Set RsTemp = G_FaCn.Execute("Select Prefix From Voucher_Prefix Where V_Type = '" & mVType & "' And Prefix<>'" & PubSprTaxInvPrefix & "' And Div_Code='" & PubDivCode & "' And Date_From<= " & ConvertDate(Txt(VDate)) & " And Date_To>= " & ConvertDate(Txt(VDate)) & " ")
    If RsTemp.RecordCount > 0 Then LblVPrefix = RsTemp(0)
    Txt(DocID) = GetDocIDmPrefix(GCnFaS, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
    Txt(DocID).Tag = Txt(DocID)
    mPartyType = 0
    Txt(DocType).SetFocus
    Txt(TurnOverPer) = MainLib.TOTCal()
    FGrid.Col = Col_SONo
    
Set RsTemp = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    If IsEditable(RetDate(Txt(VDate))) = False Then Exit Sub
    
    
    Disp_Text SETS("EDIT", Me, Master)
    mAddFlag = "E"
    Txt(DocType).Enabled = False
    Txt(VDate).Enabled = False
    Txt(SerialNo).Enabled = False
    Txt(Party).Enabled = True
    OptChal(ChalSelect).Enabled = False
    OptChal(ChalCreate).Enabled = False
    OptChal(ChalSelect).Value = True
    FGrid.AddItem FGrid.Rows
    
    If PubTaxDetOnSprInv = 1 Then
        Txt(TaxDet).TEXT = "Yes"
    Else
        Txt(TaxDet).TEXT = "No"
    End If

    'Enable / Disable Text Box if values zero
    DisableEnableFooter Txt(MRPAmtTB), Txt(MRPAmtTP), Txt(SprAmtTB), Txt(SprAmtTP), _
            Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), Txt(DiscPerTP), _
            Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), _
            Txt(GenSurPer), Txt(GenSurAmt), Txt(TaxableTot), _
            Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt)
    'EOF enable / disable section
    Txt(Address1).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim vBook As Variant, Rst As ADODB.Recordset, I As Integer, mTrans As Boolean
Dim mChalDocID$, MsgStr$, mTitle$
Dim LedgAry(1) As LedgRec, mResult As Byte
    'Check for existance of transactions
'    Set Rst = New ADODB.Recordset
'    Rst.CursorLocation = adUseClient
'    Rst.Open "Select Order_DocId from SP_Stock Where Order_DocId='" & txt(DocID) & "'", GCn, adOpenDynamic, adLockOptimistic
'    If Rst.RecordCount  > 0 Then
'        MsgBox "Dispatch Challan Exists of this Sale Order, " & vbCrLf & "Can't Delete the Reocord", vbInformation, "Validation"
'        Exit Sub
'    End If
    If IsEditable(RetDate(Txt(VDate))) = False Then Exit Sub
    
    
    ApplyConsolidatedPosting CDate(Txt(VDate))
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
            GCnFaS.BeginTrans
            mTrans = True
            
            CreateLog Me, Master!SearchCode, mReposting
            
            mChalDocID = FGrid.TextMatrix(1, Col_ChalNoCode)
            For I = 1 To FGrid.Rows - 1
                If FGrid.TextMatrix(I, Col_PNo) <> "" And Val(FGrid.TextMatrix(I, Col_Qty)) <> 0 Then
                    GCn.Execute ("Update SP_Stock Set " _
                        & "V_Date2 = " & ConvertDate("") & ",Invoice_DocId = ''," _
                        & "Rate2= 0,MRP_Rate2=0," _
                        & "Disc_Per2=0,Disc_Amt2=0," _
                        & "Amount2=0,Net_Amt2=0 " _
                        & "Where DocID='" & FGrid.TextMatrix(I, Col_ChalNoCode) & "' And Srl_No = " & Val(FGrid.TextMatrix(I, Col_ChalSrNo)))
                End If
            Next
            If Txt(DocType).TEXT = "Cash" Then
                'Delete Challan
                'one challan for cashmemo is necessary
                UpdateSO mChalDocID
                UpdStkTableToTable mChalDocID, "+", "I"
                GCn.Execute ("Delete From SP_Stock Where DocID='" & mChalDocID & "'")
                GCn.Execute ("Delete From SP_Sale Where DocID='" & mChalDocID & "'")
            Else
                GCn.Execute ("Update SP_Sale set Invoice_DocID='' Where Invoice_DocID='" & Txt(DocID) & "'")
            End If
            If mTitle = "Delete Entry!" Then
                GSQL = "Delete From SP_Sale Where DocID='" & Txt(DocID) & "'"
            Else
                GSQL = "Update SP_Sale Set " _
                    & "CancelYN=1,RoadPermit_FormCode='',RoadPermit_No='',CrAc='',SprAmt_MRP_TB=0,SprAmt_MRP_TP=0,OilAmt_MRP_TB=0,OilAmt_MRP_TP=0" & _
                    ",SprAmt_TB=0,SprAmt_TP=0,OilAmt_TB=0,OilAmt_TP=0,D_Per_TB=0,D_Amt_TB=0,D_Per_TP=0,D_Amt_TP=0,Addition=0" & _
                    ",Packing=0,Gen_Sur_Per=0,Gen_Sur_Amt=0,Trans_Amt=0,Tax_Per=0,Tax_Amt=0,Tax_Sur_Per=0,Tax_Sur_Amt=0" & _
                    ",TOT_Per=0,TOT_Amt=0,ReSalTax_Per=0,ReSalTax_Amt=0,Rounded=0,Total_Amt=0,U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'" & _
                    ",D_Per_MRP_TB=0,D_Amt_MRP_TB=0,D_Per_MRP_TP =0,D_Amt_MRP_TP=0,Tax_AmtMRP=0,TaxSur_AmtMRP=0,TOT_AmtMRP=0, SatAmt=0 " & _
                    "  Where DocID='" & Txt(DocID) & "'"
                
                'GCn.Execute ("Insert into Deletelog Values('" & txt(DocID) & "',1," & Val(txt(NetAmt)) & ",'" & pubUName & "'," & ConvertDate(date$) & ",'" & Time$ & "')")
            End If
            GCn.Execute GSQL
            'Unpost Ledger a/c
            If Txt(DocType).TEXT = "Cash" And IsConsolidatedPosting Then
                'A/c Posting
                ProcAcPost rsCtrlAc
                'EOF Posting
            Else
                mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, Txt(DocID))
                If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
                'Unposting of Ledger completed
            End If
            '***eof A/c Posting
            
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
Set Rst = Nothing
Exit Sub
ELoop:
    If mTrans Then GCn.RollbackTrans: GCnFaS.RollbackTrans
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
 sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("S.docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "Select S.DocId As SearchCode,S.Site_Code, " & cIIF("S.V_Type='" & SalCashVType & "'", "'Cash'", "'Credit'") & " As VType, " & cTrim(cMID("S.DocID", "9", "5")) & " As VPrefix, " & cCStr("S.V_No") & " As V_No, " & cDt("S.V_Date") & " AS VDate, S.Party_Name as PartyName From SP_Sale S Where left(S.DocID,1)='" & PubDivCode & "' " & sitecond & " and S.V_Type In ('" & SalCashVType & "','" & SalCrVType & "') Order by S.V_Date Desc,S.V_Type"
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
    rsTrans.Requery
    RsPart.Requery
    
    
    RsSONo.Requery
    RsGodown.Requery
    'Master.Requery
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean, mGridFilled As Boolean
Dim Rst As ADODB.Recordset, DocIdHlp$, MyGPNo$, ChalDocID$, MasterSql$, MyGPNoInv$
Dim mCrLimit As Double, mCurrBal As Double, mEditValue As Double, mGatePassOnSprInv$



'On Error GoTo ELoop
    If IsEditable(RetDate(Txt(VDate))) = False Then Exit Sub

    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    If IsValid(Txt(DocType), "Document Type") = False Then Exit Sub
    If IsValid(Txt(VDate), "Date") = False Then Exit Sub
    If IsValid(Txt(SerialNo), "Serial Number") = False Then Exit Sub
    
    
    If IsValid(Txt(Party), "Party Name") = False Then Exit Sub
    If IsValid(Txt(FormName), "Form Type") = False Then Exit Sub
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            If FGrid.TextMatrix(I, Col_MRP) = "" Then MsgBox "Please Specify MRP Yes/No in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_MRP: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Col_Taxable) = "" Then MsgBox "Please Specify Taxable Yes/No in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Taxable: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, Col_Qty)) = 0 Then MsgBox "Please Specify Quantity in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Qty: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Col_Godown) = "" Then MsgBox "Please Godown in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Godown: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, Col_Rate)) = 0 And UTrim(FGrid.TextMatrix(I, Col_MRP)) <> "YES" Then
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
            Val(Txt(DiscPerTB)), Val(Txt(DiscPerTP)), _
            Val(Txt(STaxPer)), Val(Txt(TaxSurPer)), Val(Txt(TurnOverPer))
    
    'MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
        Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
        Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
        Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
        Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
        Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
        Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
        Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
        Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
        
    If mVatYn = 1 Then
       MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
            Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
            Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
            Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
            Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
            Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
            Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, Txt(SatAmt)
    Else
        MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
            Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
            Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
            Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
            Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
            Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
            Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
    End If
    'EOF Amount Calculation
    'Check Cr Limit for Challans
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        Txt(TurnOverAmt) = Format((Val(Txt(STotATB)) + Val(Txt(STaxAmt))) * Val(Txt(TurnOverPer)) / 100, "0.00")
        Txt(NetSprAmt).TEXT = Format(Val(Txt(STotB)) + Val(Txt(TurnOverAmt)), "0.00")
        Txt(NetAmt) = Format(Val(Txt(NetSprAmt)), "0.00")
        Txt(NetAmt) = Format(Round(Val(Txt(NetSprAmt)), 0), "0.00")
    End If
    
    
    If PubCrLimitCheck = 1 And Txt(DocType) <> "Cash" Then
        mCurrBal = 0
        mEditValue = 0
        mCurrBal = VNull(G_FaCn.Execute("Select Sum(AmtDr)-Sum(AmtCr) from Ledger where SubCode='" & Txt(Party).Tag & "'").Fields(0).Value)
        mCrLimit = VNull(GCn.Execute("Select CreditLimit from SubGroup where SubCode='" & Txt(Party).Tag & "'").Fields(0).Value)
        If TopCtrl1.TopText2 <> "Add" Then
            mEditValue = VNull(GCn.Execute("Select Total_Amt from SP_Sale S Where S.DocID = '" & Master!SearchCode & "'").Fields(0).Value)
        End If
        mCurrBal = (mCurrBal - mEditValue) + Val(Txt(NetAmt))
        If mCurrBal > 0 Then    'Dr Balance
            If mCurrBal > mCrLimit And mCrLimit > 0 Then
                MsgBox "Cr Limit Rs." & mCrLimit & " Exceeds by Rs." & mCurrBal - mCrLimit & vbCrLf & "Add/Edit Denied !", vbInformation, "Cr Limit Checking"
                Me.ActiveControl.SetFocus: Exit Sub
            End If
        End If
    End If
    'EOF Cr Limit Checking
    
'    'Calculating Net Rate for Each Part, to be modified
'If MRP=N
'   IF Tax=Y
'      BasVRate = TBItemValue - TBDisAmt + GenSurg
'   Else: TaxPaid
'      BasVRate = TBItemValue - TPDisAmt
'   End If
'Else
'   IF Tax=Y
'      BasVRate = TBMRPItemValue - TBMRPDisAmt
'   Else: TaxPaid
'      BasVRate = TPMRPItemValue - TPMRPDisAmt
'   End If
'End If
'BasVRate = BasVRate + Transportation + MisChrg
'    mTotDiffAmt = Val(txt(TaxAmt)) + Val(txt(Addition)) + Val(txt(Deduction))
'    mDiffPosted = 0
'    If mTotDiffAmt <> 0 Then
'        For i = 1 To FGrid.Rows - 1
'            If FGrid.TextMatrix(i, PNo) <> "" Then
'                mItemVal = Val(FGrid.TextMatrix(i, ItemVal))
'                mItemQty = Val(FGrid.TextMatrix(i, PQty))
'                mDiffPerc = Round((mItemVal * 100) / Val(txt(TotGoods)), 2)
'                mDiffAmt = Round(mTotDiffAmt * mDiffPerc / 100, 2)
'                mDiffPosted = mDiffPosted + mDiffAmt
'                FGrid.TextMatrix(i, NDP) = Round((mItemVal + mDiffAmt) / mItemQty, 2)
'                LastI = i
'            End If
'        Next
'    End If
'    If mTotDiffAmt - mDiffPosted <> 0 Then
'        mItemVal = Val(FGrid.TextMatrix(LastI, ItemVal))
'        mItemQty = Val(FGrid.TextMatrix(LastI, PQty))
'        FGrid.TextMatrix(LastI, NDP) = Round((mItemVal + mTotDiffAmt - mDiffPosted) / mItemQty, 2)
'    End If
'    'EOF Landed Rate Calculation
    GCn.BeginTrans
    G_FaCn.BeginTrans
    mTrans = True
    
    If mAddFlag = "A" Then  'Case of ADD
        'lp 12-03-03
        Txt(DocID).Tag = Txt(DocID)
'
'        If GCn.Execute("Select Count(*) From SP_Sale Where Left(DocID,1)='" & PubDivCode & "' And V_Type = '" & mVType & "' And " & cTrim(cMID("DocID", "9", "5")) & "='" & LblVPrefix & "'  And V_No= " & Val(txt(SerialNo)) & " ").Fields(0) > 0 Then
'            If VoucherEditFlag Then
'                MsgBox "Document No. already exists, Retry", vbCritical, "Validation Error"
'                txt(SerialNo).SetFocus
'                GoTo Eloop
'            Else
'                txt(DocId) = GetDocIDmPrefix(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
'                If Val(txt(SerialNo)) <= Val(DeCodeDocID(txt(DocId).Tag, Document_No)) Then
'                    'SetMax_VoucherPrefix "DocID", mVType, "SP_Sale"
'                    MsgBox "Document No. Already Exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
'                    GoTo Eloop
'                End If
'            End If
'        End If

        If GCn.Execute("Select Count(*) From SP_Sale Where DocID='" & Txt(DocID) & "' ").Fields(0) > 0 Then
            If VoucherEditFlag Then
                MsgBox "Document No. already exists, Retry", vbCritical, "Validation Error"
                Txt(SerialNo).SetFocus
                GoTo ELoop
            Else
                Txt(DocID) = GetDocIDmPrefix(GCnFaS, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                If Val(Txt(SerialNo)) <= Val(DeCodeDocID(Txt(DocID).Tag, Document_No)) Then
                    'SetMax_VoucherPrefix "DocID", mVType, "SP_Sale"
                    MsgBox "Document No. Already Exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo ELoop
                End If
            End If
        End If

        DocIdHlp = UCase(Replace(Txt(DocID), " ", ""))
        '********* modi lps 02.09.03
        mGatePassOnSprInv = GCn.Execute("Select GatePassOnSprInv from Syctrl").Fields(0).Value
        If (OptChal(ChalCreate).Value = True And mGatePassOnSprInv <> 1) Or _
            (mVType = SalCashVType Or mGatePassOnSprInv = 1) Then
            ' For Gate Pass
            MyGPNo = "00000" & GCn.Execute("select " & vIsNull("max(" & cVal("right(gp_no,5)") & ")", "0") & "+1 from SP_Sale where left(gp_no,1)='" & PubDivCode & "' AND " & cMID("gp_no", "2", "1") & "='" & PubSiteCode & "' And (Job_DocID Is Null or " & cTrim("Job_DocID") & "='')").Fields(0).Value
            MyGPNo = PubDivCode & PubSiteCode & ForSiteCode & Right(MyGPNo, 5)
            lblGatePass = MyGPNo
        End If
        
        If (mVType = SalCashVType Or mGatePassOnSprInv = 1) Then
            MyGPNoInv = MyGPNo
            MyGPNo = ""
        End If
        If OptChal(ChalCreate).Value = True Then 'Create Challan
            ChallanCreate ChalDocID, MyGPNo
        End If
        MasterSql = "Insert Into SP_Sale(" _
            & "DocID ,DocIDHelp ,V_Type ,V_No ,Site_Code ," _
            & "V_Date ,Cash_Credit ,Party_Code ,Party_Name ,Address ," _
            & "L_C ,Form_Code ,RoadPermit_FormCode ,RoadPermit_No ,GR_RR_No ," _
            & "GR_RR_Date ,CrAc ,Case_No ,Case_Mark ,Mode_Dispatch ," _
            & "Transport ,Rep_Code,Remarks ,Det_Tax ,SprAmt_MRP_TB ," _
            & "SprAmt_MRP_TP,OilAmt_MRP_TB,OilAmt_MRP_TP,SprAmt_TB ,SprAmt_TP ,OilAmt_TB ,OilAmt_TP ," _
            & "D_Per_TB ,D_Amt_TB ,D_Per_TP ,D_Amt_TP ,Addition ," _
            & "Packing ,Gen_Sur_Per ,Gen_Sur_Amt ,Trans_Amt ,Tax_Per ," _
            & "Tax_Amt ,Tax_Sur_Per ,Tax_Sur_Amt ,TOT_Per ,TOT_Amt ," _
            & "ReSalTax_Per, ReSalTax_Amt,Rounded ,Total_Amt,U_Name ,U_EntDt,U_AE, AddBy, AddDate,D_Per_MRP_TB, " _
            & "D_Amt_MRP_TB, D_Per_MRP_TP, D_Amt_MRP_TP, Tax_AmtMRP, TaxSur_AmtMRP, TOT_AmtMRP,GP_No,GP_Date,PType, SatAmt, Sat_Yn ) " _
            & "Values('" & Txt(DocID) & "','" & DocIdHlp & "','" & mVType & "'," & Txt(SerialNo) & ",'" & PubSiteCode & ForSiteCode & _
            "'," & ConvertDate(Format(Txt(VDate), "dd/MMM/yyyy")) & ",'" & Txt(DocType) & "','" & Txt(Party).Tag & "','" & Txt(Party) & "','" & Txt(Address1) & _
            "','" & left(Txt(LC), 1) & "','" & Txt(FormName).Tag & "','" & Txt(Form31Name).Tag & "','" & Txt(Form31No) & "','" & Txt(LRNo) & _
            "', " & ConvertDate(Txt(LRDate)) & ",'" & Txt(CrAc).Tag & "'," & Val(Txt(CaseNo)) & ",'" & Txt(CaseMark) & "','" & Txt(DispMode) & _
            "','" & Txt(Transport) & "','" & Txt(SPerson).Tag & "','" & Txt(Remark) & "'," & IIf(Txt(TaxDet) = "Yes", 1, 0) & "," & Val(Txt(MRPAmtTB)) - mMRPLubeTB & _
            " , " & Val(Txt(MRPAmtTP)) - mMRPLubeTP & "," & mMRPLubeTB & "," & mMRPLubeTP & "," & Val(Txt(SprAmtTB)) & "," & Val(Txt(SprAmtTP)) & "," & Val(Txt(OilAmtTB)) & "," & Val(Txt(OilAmtTP)) & _
            " , " & Val(Txt(DiscPerTB)) & "," & Val(Txt(DiscAmtTB)) & "," & Val(Txt(DiscPerTP)) & "," & Val(Txt(DiscAmtTP)) & "," & Val(Txt(Addition)) & _
            " , " & Val(Txt(PackCrg)) & "," & Val(Txt(GenSurPer)) & "," & Val(Txt(GenSurAmt)) & "," & Val(Txt(TransAmt)) & "," & Val(Txt(STaxPer)) & _
            " , " & Val(Txt(STaxAmt)) & "," & Val(Txt(TaxSurPer)) & "," & Val(Txt(TaxSurAmt)) & "," & Val(Txt(TurnOverPer)) & "," & Val(Txt(TurnOverAmt)) & _
            " , " & Val(Txt(ReSalTaxPer)) & "," & Val(Txt(ReSalTaxAmt)) & "," & Val(Txt(SROff)) & "," & Val(Txt(NetAmt)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A', '" & pubUName & "', " & ConvertDateTime(PubServerDate) & "," & mMRevDisTBPer & _
            " , " & mTBDisAmtMRP & "," & mMRevDisTPPer & "," & mTPDisAmtMRP & "," & mMRPTax & "," & mMRPTaxSur & ", " & mMRPTOT & ", '" & MyGPNoInv & "', " & ConvertDate(IIf(MyGPNoInv = "", "", PubServerDate)) & ",'" & Txt(PType) & "', " & Val(Txt(SatAmt)) & ", " & IIf(mSatYn, 1, 0) & ")"
            
        GCn.Execute MasterSql
        ' For Updation of Selected Challans in SP_Sale Table
        For I = 1 To FGridSel.Rows - 1
            If FGridSel.TextMatrix(I, SCol_SrNo) <> "" And Val(FGridSel.TextMatrix(I, SCol_ChalNoCode)) <> 0 Then
                GCn.Execute ("Update SP_Sale Set Invoice_DocId = '" & Txt(DocID).TEXT & "' Where DocId = '" & IIf(OptChal(ChalCreate).Value = True, ChalDocID, FGridSel.TextMatrix(I, SCol_ChalNoCode)) & "'")
            End If
        Next
        ' Line File Updation (Invoice No,Date, Rate Etc..)
        'Separate for Add because Invoice No. Updated
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" And Val(FGrid.TextMatrix(I, Col_Qty)) <> 0 Then
                GCn.Execute ("Update SP_Stock Set " _
                    & "Invoice_DocId ='" & Txt(DocID).TEXT & "',V_Date2=" & ConvertDate(Txt(VDate).TEXT) & _
                    ",Rate2=" & Val(FGrid.TextMatrix(I, Col_Rate)) & ",MRP_Rate2=" & Val(FGrid.TextMatrix(I, Col_MRPRate)) & _
                    ",Disc_Per2=" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & ",Disc_Amt2=" & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & _
                    ",Amount2=" & Val(FGrid.TextMatrix(I, Col_Amt)) & ",Net_Amt2=" & Val(FGrid.TextMatrix(I, Col_ItemVal)) & _
                    ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E',TaxPer=" & Val(FGrid.TextMatrix(I, Col_TaxPer)) & ",TaxAmt=" & Val(FGrid.TextMatrix(I, Col_TaxAmt1)) & ",  " & _
                    " SatPer = " & Val(FGrid.TextMatrix(I, Col_SatPer)) & ", SatAmt = " & Val(FGrid.TextMatrix(I, Col_SatAmt)) & " " _
                    & " Where DocID='" & IIf(OptChal(ChalCreate).Value = True, ChalDocID, FGrid.TextMatrix(I, Col_ChalNoCode)) & _
                    "' And Srl_No=" & Val(FGrid.TextMatrix(I, Col_ChalSrNo)))
            End If
        Next
        'Sale Numbering
        'Voucher Serial No. Updation LPS 21-05-03
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaS, Txt(DocID), Txt(VDate)
            
    ElseIf mAddFlag = "E" Then   'Edit Bill
        CreateLog Me, Master!SearchCode, mReposting
                
    
    
        'If Txt(DocType) = "Cash" Then 'Edit Challan
        ChallanEdit Txt(DocID)
        'End If
        GCn.Execute "Update SP_Sale Set " _
            & "Party_Code='" & Txt(Party).Tag & "',Party_Name='" & Txt(Party) & "',Address='" & Txt(Address1) & _
            "',L_C='" & left(Txt(LC), 1) & "',Form_Code='" & Txt(FormName).Tag & _
            "',RoadPermit_FormCode='" & Txt(Form31Name).Tag & "',RoadPermit_No='" & Txt(Form31No) & _
            "',GR_RR_No='" & Txt(LRNo) & "',GR_RR_Date=" & ConvertDate(Txt(LRDate)) & _
            " ,CrAc='" & Txt(CrAc).Tag & "',Case_No=" & Val(Txt(CaseNo)) & ",Case_Mark='" & Txt(CaseMark) & _
            "',Mode_Dispatch='" & Txt(DispMode) & "',Transport='" & Txt(Transport) & "',Rep_Code='" & Txt(SPerson).Tag & _
            "',Remarks='" & Txt(Remark) & "',Det_Tax=" & IIf(Txt(TaxDet) = "Yes", 1, 0) & ",SprAmt_MRP_TB=" & Val(Txt(MRPAmtTB)) - mMRPLubeTB & _
            " ,SprAmt_MRP_TP=" & Val(Txt(MRPAmtTP)) - mMRPLubeTP & ",OilAmt_MRP_TB=" & mMRPLubeTB & ",OilAmt_MRP_TP=" & mMRPLubeTP & ",SprAmt_TB=" & Val(Txt(SprAmtTB)) & ",SprAmt_TP=" & Val(Txt(SprAmtTP)) & _
            " ,OilAmt_TB=" & Val(Txt(OilAmtTB)) & ",OilAmt_TP=" & Val(Txt(OilAmtTP)) & ",D_Per_TB=" & Val(Txt(DiscPerTB)) & _
            " ,D_Amt_TB=" & Val(Txt(DiscAmtTB)) & ",D_Per_TP=" & Val(Txt(DiscPerTP)) & ",D_Amt_TP=" & Val(Txt(DiscAmtTP)) & _
            " ,Addition=" & Val(Txt(Addition)) & ",Packing=" & Val(Txt(PackCrg)) & ",Gen_Sur_Per=" & Val(Txt(GenSurPer)) & _
            " ,Gen_Sur_Amt=" & Val(Txt(GenSurAmt)) & ",Trans_Amt=" & Val(Txt(TransAmt)) & ",Tax_Per=" & Val(Txt(STaxPer)) & _
            " ,Tax_Amt=" & Val(Txt(STaxAmt)) & ",Tax_Sur_Per=" & Val(Txt(TaxSurPer)) & ",Tax_Sur_Amt=" & Val(Txt(TaxSurAmt)) & _
            " ,TOT_Per=" & Val(Txt(TurnOverPer)) & ",TOT_Amt=" & Val(Txt(TurnOverAmt)) & _
            " ,ReSalTax_Per=" & Val(Txt(ReSalTaxPer)) & ", ReSalTax_Amt=" & Val(Txt(ReSalTaxAmt)) & ",Rounded=" & Val(Txt(SROff)) & _
            " ,Total_Amt=" & Val(Txt(NetAmt)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'" & _
            " , ModifyBy = '" & pubUName & "', ModifyDate = " & ConvertDateTime(PubServerDate) & ",D_Per_MRP_TB=" & mMRevDisTBPer & ",D_Amt_MRP_TB=" & mTBDisAmtMRP & ", D_Per_MRP_TP =" & mMRevDisTPPer & ", D_Amt_MRP_TP=" & mTPDisAmtMRP & _
            " ,Tax_AmtMRP=" & mMRPTax & ",TaxSur_AmtMRP= " & mMRPTaxSur & ", TOT_AmtMRP= " & mMRPTOT & _
            " ,PType='" & Txt(PType) & "', SatAmt = " & Val(Txt(SatAmt)) & " Where DocID='" & Txt(DocID) & "'"
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" And Val(FGrid.TextMatrix(I, Col_Qty)) <> 0 Then
                If mReposting Or PubSiebelActiveYn = 1 Then
                    GCn.Execute ("Update SP_Stock " _
                        & "Set Rate2=" & Val(FGrid.TextMatrix(I, Col_Rate)) & ", MRP_Rate2=" & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," _
                        & "Disc_Per2=" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & ",Disc_Amt2=" & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," _
                        & "Amount2=" & Val(FGrid.TextMatrix(I, Col_Amt)) & ",Net_Amt2=" & Val(FGrid.TextMatrix(I, Col_ItemVal)) & "," _
                        & "U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E',TaxPer=" & Val(FGrid.TextMatrix(I, Col_TaxPer)) & ",TaxAmt=" & Val(FGrid.TextMatrix(I, Col_TaxAmt1)) & ",SatPer=" & Val(FGrid.TextMatrix(I, Col_SatPer)) & ", SATAmt=" & Val(FGrid.TextMatrix(I, Col_SatAmt)) & "  " _
                        & "Where DocID='" & FGrid.TextMatrix(I, Col_ChalNoCode) & "' and  Part_No='" & FGrid.TextMatrix(I, Col_PNo) & "' And Mrp_Yn=" & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & " And Tax_Yn=" & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & " And Qty_Iss=" & Val(FGrid.TextMatrix(I, Col_Qty)) & "")
                Else
                    GCn.Execute ("Update SP_Stock " _
                        & "Set Rate2=" & Val(FGrid.TextMatrix(I, Col_Rate)) & ", MRP_Rate2=" & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," _
                        & "Disc_Per2=" & Val(FGrid.TextMatrix(I, Col_DiscPer)) & ",Disc_Amt2=" & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," _
                        & "Amount2=" & Val(FGrid.TextMatrix(I, Col_Amt)) & ",Net_Amt2=" & Val(FGrid.TextMatrix(I, Col_ItemVal)) & "," _
                        & "U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E',TaxPer=" & Val(FGrid.TextMatrix(I, Col_TaxPer)) & ",TaxAmt=" & Val(FGrid.TextMatrix(I, Col_TaxAmt1)) & ",SatPer=" & Val(FGrid.TextMatrix(I, Col_SatPer)) & ", SATAmt=" & Val(FGrid.TextMatrix(I, Col_SatAmt)) & "  " _
                        & "Where DocID='" & FGrid.TextMatrix(I, Col_ChalNoCode) & "' and  Part_No='" & FGrid.TextMatrix(I, Col_PNo) & "' And Mrp_Yn=" & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & " And Tax_Yn=" & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & " And Srl_No=" & Val(FGrid.TextMatrix(I, Col_ChalSrNo)) & "")
                End If
            End If
        Next
    End If
    'A/c Posting
    If mRePostCounter = 0 Then ProcAcPost rsCtrlAc
    'EOF Posting
    'Nra Update for ENAR
    CustOrdDet = ""
    If RsSONo.RecordCount > 0 Then
        If RsSONo!Code <> "" Then
            CustOrdDet = GCn.Execute("Select SPO.CustOrd_Det from SP_Order SPO where OrderId='" & RsSONo!Code & "'").Fields(0).Value
        End If
    End If
    'End upadte
    G_FaCn.CommitTrans
    GCn.CommitTrans
    mTrans = False
    mSearchCode = Txt(DocID)
    If PubMoveRecYn Then
        If TopCtrl1.TopText2 = "Add" Then Master.Requery
    Else
        Set Master = GCn.Execute("Select S.DocID As SearchCode,U_EntDt, V_Date From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' and S.V_Type In ('" & SalCashVType & "','" & SalCrVType & "') And S.DocID  = '" & mSearchCode & "' " _
            & "Order by S.V_Date desc,S.DocID desc")
    End If
    Master.FIND "SearchCode = '" & mSearchCode & "'"
    'lp 12-03-03
    'VIK TEMP
    If mAddFlag = "A" Then
        If Val(Txt(SerialNo)) > Val(DeCodeDocID(Txt(DocID).Tag, Document_No)) Then
            MsgBox "Document No." & Trim(DeCodeDocID(Txt(DocID).Tag, Document_No)) & " already exists ! " & vbCrLf & "New No. " & Txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
        TopCtrl1_ePrn
    Else
        Disp_Text SETS("INI", Me, Master)
        MoveRec
    End If
    
Exit Sub
ELoop:
    If mTrans Then GCn.RollbackTrans: G_FaCn.CommitTrans
    CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
'        Master.FIND "SearchCode='" & mSearchCode & "'"
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To Txt.Count - 1
            Txt(I).BackColor = CtrlBColOrg
            Txt(I).ForeColor = CtrlFColOrg
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
Ctrl_GetFocus Txt(Index)
TxtGrid(0).Visible = False
Grid_Hide

Select Case Index
    Case DocType
        ListArray = Array("Cash", "Credit")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 2)
    Case Party
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case FormName
        rsForm.Filter = adFilterNone
        rsForm.Filter = "L_C = '" & Txt(LC) & "'"
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsForm!Name Then
            rsForm.MoveFirst
            rsForm.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case Form31Name
        If rsForm31.RecordCount = 0 Or (rsForm31.EOF = True Or rsForm31.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> rsForm31!Name Then
            rsForm31.MoveFirst
            rsForm31.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case SPerson
        If RsPerson.RecordCount = 0 Or (RsPerson.EOF = True Or RsPerson.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsPerson!Name Then
            RsPerson.MoveFirst
            RsPerson.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case LC
        ListArray = Array("Local", "Central")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 2)
    Case DiscPerTB, DiscPerTP, GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ReSalTaxPer
        Txt(Index).Tag = Txt(Index)
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
        Case DocType
            ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 600
        Case SerialNo
            NumDown Txt(Index), KeyCode, 8, 0
        Case LC
            ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 600
        Case Party
            If Txt(DocType).TEXT = "Credit" Then
                DGridTxtKeyDown DGParty, Txt, Party, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
                If KeyCode = 13 Then
                    If RsParty.BOF = False And RsParty.EOF = False Then
                        LblCurrBal = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
                        LblCurrBal = LblCurrBal & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))
                    End If
                End If
            End If
        Case CrAc
            If Txt(DocType).TEXT = "Credit" Then
                DGridTxtKeyDown DGCrAc, Txt, CrAc, RsCrAc, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
            End If
        Case SPerson
            DGridTxtKeyDown DGPerson, Txt, SPerson, RsPerson, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
        Case FormName
            DGridTxtKeyDown DGForm, Txt, FormName, rsForm, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
        Case Form31Name
            DGridTxtKeyDown DGForm31, Txt, Form31Name, rsForm31, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
        Case Transport
            DGridTxtKeyDown_Mast DGTrans, Txt, Transport, rsTrans, KeyCode, False, 0
        Case CaseNo
            NumDown Txt(Index), KeyCode, 8, 0
        Case DiscAmtTB, DiscAmtTP, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, PackCrg, TurnOverAmt, ReSalTaxAmt
            NumDown Txt(Index), KeyCode, 8, 2
        Case GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ReSalTaxPer
            NumDown Txt(Index), KeyCode, 2, 2
        Case DiscPerTB, DiscPerTP
            NumDown Txt(Index), KeyCode, 2, 4
    End Select
    If FrmList.Visible = False And DGParty.Visible = False And DGCrAc.Visible = False And DGForm.Visible = False And DGForm31.Visible = False And DGTrans.Visible = False And DGPerson.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And _
            ((PubReSaleTaxPer = 0 And Index = TurnOverAmt) Or _
            (PubReSaleTaxPer <> 0 And Index = ReSalTaxAmt)) Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave    ' Else Me.ActiveControl.SetFocus
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
        NumPress Txt(Index), KeyAscii, 8, 0
    Case Party
        If DGParty.Visible = True Then DGridTxtKeyPress Txt, Party, RsParty, KeyAscii, "Name"
        lblGroup.Visible = True: lblGroup.Locked = True: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
    Case CrAc
        If Txt(DocType).TEXT = "Credit" Then
            If DGCrAc.Visible = True Then DGridTxtKeyPress Txt, CrAc, RsCrAc, KeyAscii, "Name"
        End If
    Case SPerson
        If DGPerson.Visible = True Then DGridTxtKeyPress Txt, SPerson, RsPerson, KeyAscii, "Name"
    Case FormName
        If DGForm.Visible = True Then DGridTxtKeyPress Txt, FormName, rsForm, KeyAscii, "Name"
    Case Form31Name
        If DGForm31.Visible = True Then DGridTxtKeyPress Txt, Form31Name, rsForm31, KeyAscii, "Name"
    Case TaxDet
        If KeyAscii = 89 Or KeyAscii = 121 Or KeyAscii = 78 Or KeyAscii = 110 Then
            If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                Txt(Index).TEXT = "Yes"
                KeyAscii = 0
            ElseIf KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                Txt(Index).TEXT = "No"
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
    Case PType
        If Asc("g") = KeyAscii Or Asc("G") = KeyAscii Then
            Txt(PType) = "General"
        ElseIf Asc("d") = KeyAscii Or Asc("D") = KeyAscii Then
            Txt(PType) = "Dealer"
        End If
        KeyAscii = 0
    Case CaseNo
        NumPress Txt(Index), KeyAscii, 8, 0
    Case DiscAmtTB, DiscAmtTP, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, PackCrg, TurnOverAmt, ReSalTaxAmt
        NumPress Txt(Index), KeyAscii, 8, 2
    Case GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ReSalTaxPer
        NumPress Txt(Index), KeyAscii, 2, 2
    Case DiscPerTB, DiscPerTP
        NumPress Txt(Index), KeyAscii, 2, 4
'    Case SROff
'        NumPress Txt(Index), KeyAscii, 0, 2
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
        Case DocType
            If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
        Case LC
            If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
        Case Transport
            If DGTrans.Visible = True Then DGridTxtKeyUp_Mast Txt, Transport, rsTrans, KeyCode, "Name"
        Case DiscPerTB, DiscAmtTB, DiscPerTP, DiscAmtTP, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, TurnOverPer, PackCrg, TurnOverAmt, ReSalTaxPer, ReSalTaxAmt
            Amt_Cal
            If Val(Txt(MRPAmtTB)) + Val(Txt(MRPAmtTP)) <> 0 Then
                MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
                        Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
                        Val(Txt(DiscPerTB)), Val(Txt(DiscPerTP)), _
                        Val(Txt(STaxPer)), Val(Txt(TaxSurPer)), Val(Txt(TurnOverPer))
            End If
            If mVatYn = 1 Then
               MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
                    Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
                    Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
                    Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
                    Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
                    Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
                    Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, Txt(SatAmt)
            Else
                MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                    Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                    Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
                    Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
                    Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
                    Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
                    Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
                    Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
                    Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
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
'            If UCase(left(PubComp_Name, 3)) = "JMK" Then
'                    Txt(TurnOverAmt) = Format((Val(Txt(STotATB)) + Val(Txt(STaxAmt))) * Val(Txt(TurnOverPer)) / 100, "0.00")
'                    Txt(NetSprAmt).TEXT = Format(Val(Txt(STotB)) + Val(Txt(TurnOverAmt)), "0.00")
'                    Txt(NetAmt) = Format(Val(Txt(NetSprAmt)), "0.00")
'                    Txt(NetAmt) = Format(Round(Val(Txt(NetSprAmt)), 0), "0.00")
'            Else
'                Txt(NetAmt) = Format(Round(Val(Txt(NetSprAmt)), 0), "0.00")
'            End If
             If UCase(left(PubComp_Name, 3)) = "JMK" Then
                Txt(TurnOverAmt) = Format((Val(Txt(STotATB)) + Val(Txt(STaxAmt))) * Val(Txt(TurnOverPer)) / 100, "0.00")
                Txt(NetSprAmt).TEXT = Format(Val(Txt(STotB)) + Val(Txt(TurnOverAmt)), "0.00")
                Txt(NetAmt) = Format(Val(Txt(NetSprAmt)), "0.00")
                Txt(NetAmt) = Format(Round(Val(Txt(NetSprAmt)), 0), "0.00")
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, I As Byte
Dim RsTemp As ADODB.Recordset
On Error GoTo ELoop
    Select Case Index
        Case DocType
            Txt(Index).TEXT = ListView.SelectedItem.TEXT
            If Not Trim(Txt(Index).TEXT) <> "Cash" Or Trim(Txt(Index).TEXT) <> "Credit" Then
                Txt(Index).TEXT = "Cash"
            End If
            If Trim(Txt(Index).TEXT) = "Cash" Then
                Txt(Party).Tag = PubSprCashAc
                Txt(CrAc).Enabled = False
                Txt(CrAc).TEXT = ""
                mVType = SalCashVType
                OptChal(0).Value = False
                OptChal(1).Value = True
                Txt(PType).Visible = True
                Lbl(4).Visible = True
                Txt(PType) = "General"
            ElseIf Trim(Txt(Index).TEXT) = "Credit" Then
                Txt(CrAc).Enabled = True
                mVType = SalCrVType
                OptChal(0).Value = False
                OptChal(1).Value = True
                Txt(PType) = ""
                Txt(PType).Visible = False
                Lbl(4).Visible = False
            End If
            
            Set RsTemp = G_FaCn.Execute("Select Prefix From Voucher_Prefix Where V_Type = '" & mVType & "' And Prefix<>'" & PubSprTaxInvPrefix & "' And Div_Code='" & PubDivCode & "' And Date_From<= " & ConvertDate(Txt(VDate)) & " And Date_To>= " & ConvertDate(Txt(VDate)) & " ")
            If RsTemp.RecordCount > 0 Then LblVPrefix = RsTemp(0)
            Txt(DocID) = GetDocIDmPrefix(GCnFaS, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
            Txt(DocID).Tag = Txt(DocID)
        Case VDate
            Txt(Index).TEXT = RetDate(Txt(Index))
            Cancel = Not CheckFinYear(Txt(Index))
            If Cancel = False Then
                Txt(DocID) = GetDocIDmPrefix(GCnFaS, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                Txt(DocID).Tag = Txt(DocID)
            End If
        Case SerialNo
            If VoucherEditFlag = True Then      ' Manual
                Txt(DocID) = GetDocIDmPrefix(GCnFaS, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                Txt(DocID).Tag = Txt(DocID)
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select V_No From SP_Sale Where DocID='" & Txt(DocID).TEXT & "'", GCn, adOpenDynamic, adLockOptimistic
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                    Cancel = True
                    Txt(SerialNo).SetFocus
                End If
            End If
        Case Party
            If Trim(Txt(Index).TEXT = "") Then
                MsgBox "Please Select Party", vbInformation, "Information"
                Txt(Index).SetFocus
                Cancel = True
                Exit Sub
            End If
            ' To Populate Sale Orders Data Grid by the Customer
'             If Txt(Index).Tag = GEditText Then Exit Sub
            If Txt(DocType).TEXT = "Credit" Then
                If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then
                    Txt(Index).TEXT = ""
                    Txt(Index).Tag = ""
                    Txt(Address1).TEXT = ""
                    mPartyType = 0
                    GSQL = ""
                Else
                    Txt(Index).TEXT = RsParty!Name
                    Txt(Index).Tag = RsParty!Code
                    LblCurrBal = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
                    LblCurrBal = LblCurrBal & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))

                    Txt(Address1).TEXT = RsParty!Add1
                    If Txt(Transport) = "" Then
                        Txt(Transport).TEXT = IIf(IsNull(RsParty!Transporter), "", RsParty!Transporter)
                    End If
                    mPartyType = VNull(RsParty!Party_Type)
                    GSQL = "Select OrderID as Code," & cTrim(cMID("OrderID", "9", "5")) & " + " & cCStr(cTrim("Right(OrderID,8)")) & " as Name,V_Date From SP_Order Where left(OrderID,1)='" & PubDivCode & "' and Order_Type='S_SO' and Party_Code='" & Txt(Party).Tag & "'  and V_Date<=" & ConvertDate(Format(Txt(VDate), "dd-mm-yyyy")) & " and OrdClosDate is null Order By OrderID"
                    
                    If TopCtrl1.TopText2 = "Add" Then
                        If XNull(RsParty!LstNo) <> "" Then
                            If PubSprTaxInvPrefix <> "" Then
                                LblVPrefix = PubSprTaxInvPrefix
                            End If
                        Else
                            Set RsTemp = G_FaCn.Execute("Select Prefix From Voucher_Prefix Where V_Type = '" & mVType & "' And Prefix<>'" & PubSprTaxInvPrefix & "' And Div_Code='" & PubDivCode & "' And Date_From<= " & ConvertDate(Txt(VDate)) & " And Date_To>= " & ConvertDate(Txt(VDate)) & " ")
                            If RsTemp.RecordCount > 0 Then LblVPrefix = RsTemp(0)
                        End If
                        
                        Txt(DocID) = GetDocIDmPrefix(GCnFaS, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                        Txt(DocID).Tag = Txt(DocID)
                    End If
                End If
            Else
                Txt(Party).Tag = PubSprCashAc
                mPartyType = 0
                GSQL = "Select OrderID as Code," & cTrim(cMID("OrderID", "9", "5")) & " + " & cCStr(cTrim("Right(OrderID,8)")) & " as Name,V_Date From SP_Order Where left(OrderID,1)='" & PubDivCode & "' and Order_Type='S_SO' and V_Date<=" & ConvertDate(Format(Txt(VDate), "dd-mm-yyyy")) & " and OrdClosDate is null Order By OrderID"
            End If
            If GSQL <> "" Then
                Set RsSONo = New ADODB.Recordset
                RsSONo.CursorLocation = adUseClient
                RsSONo.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
                Set DGSONo.DataSource = RsSONo
            End If
        Case LC
            Txt(Index).TEXT = ListView.SelectedItem.TEXT
        Case FormName
            If rsForm.RecordCount > 0 Or (rsForm.EOF = False Or rsForm.BOF = False) Then
                If Trim(Txt(Index).TEXT = "") Then
                    MsgBox "Please Select Form Type", vbInformation, "Information"
                    Txt(Index).SetFocus
                    Cancel = True
                    Exit Sub
                Else        'If Txt(Index).Text <> "" Then
                    If (IsNull(rsForm!Tax_Ac_Code) Or rsForm!Tax_Ac_Code = "") Or _
                        (IsNull(rsForm!Sur_Ac_Code) Or rsForm!Sur_Ac_Code = "") Or _
                        (IsNull(rsForm!PurSal_Ac_Code) Or rsForm!PurSal_Ac_Code = "") Then
                        MsgBox "Please define Tax / Sales A/c in selected Form !", vbCritical, "Define A/c Code"
                        Txt(Index).SetFocus
                    End If
                    Txt(Index).TEXT = rsForm!Name
                    Txt(Index).Tag = rsForm!Code
                    If TopCtrl1.TopText2.CAPTION = "Add" Then   ' To Assign Tax% in case of Add
                        Txt(STaxPer).TEXT = rsForm!Tax_Per
                        Txt(TaxSurPer).TEXT = rsForm!Tax_Sur_Per
                        MainLib.SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
                                Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
                                Val(Txt(DiscPerTB)), Val(Txt(DiscPerTP)), _
                                Val(Txt(STaxPer)), Val(Txt(TaxSurPer)), Val(Txt(TurnOverPer))
                        If mVatYn = 1 Then
                           MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                                Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                                Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
                                Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
                                Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
                                Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
                                Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
                                Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
                                Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, Txt(SatAmt)
                        Else
                            MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                                Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                                Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
                                Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
                                Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
                                Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
                                Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
                                Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
                                Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
                        End If
                        If UCase(left(PubComp_Name, 3)) = "JMK" Then
                            Txt(TurnOverAmt) = Format((Val(Txt(STotATB)) + Val(Txt(STaxAmt))) * Val(Txt(TurnOverPer)) / 100, "0.00")
                            Txt(NetSprAmt).TEXT = Format(Val(Txt(STotB)) + Val(Txt(TurnOverAmt)), "0.00")
                            Txt(NetAmt) = Format(Val(Txt(NetSprAmt)), "0.00")
                            Txt(NetAmt) = Format(Round(Val(Txt(NetSprAmt)), 0), "0.00")
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
            Else
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            End If
        Case Form31Name
            If rsForm31.RecordCount > 0 Or (rsForm31.EOF = False Or rsForm31.BOF = False) Then
                If Txt(Index).TEXT <> "" Then
                    Txt(Index).TEXT = rsForm31!Name
                    Txt(Index).Tag = rsForm31!Code
                End If
            Else
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            End If
        Case TaxDet
            If Not Trim(Txt(Index).TEXT) <> "Yes" Or Trim(Txt(Index).TEXT) <> "No" Then
                Txt(Index).TEXT = "Yes"
            End If
        Case LRDate
            Txt(Index).TEXT = RetDate(Txt(Index))
        Case CaseNo, DiscAmtTB, DiscAmtTP, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, TurnOverPer, PackCrg, TurnOverAmt, ReSalTaxPer, ReSalTaxAmt, SROff
            If Index <> CaseNo Then
                If Val(Txt(Index).TEXT) = 0 Then
                    Txt(Index).TEXT = ""
                Else
                    Txt(Index).TEXT = Format(Txt(Index), "0.00")
                End If
            End If
        Case DiscPerTB, DiscPerTP
            If Index <> CaseNo Then
                If Val(Txt(Index).TEXT) = 0 Then
                    Txt(Index).TEXT = ""
                Else
                    Txt(Index).TEXT = Format(Txt(Index), "0.0000")
                End If
            End If
    End Select
    'Removing Tag Value for SprCalc purpose
    Select Case Index
        Case DiscPerTB, DiscPerTP, GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ReSalTaxPer
            Txt(Index).Tag = Txt(Index)
    End Select
    'EOF Tag Value Removal
Set Rst = Nothing
Set RsTemp = Nothing
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
Grid_Hide
If FrmDetail.Visible = False Then FrmDetail.Visible = True
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
TxtGrid(Index).MaxLength = 0
Select Case FGrid.Col
    Case Col_SONo
        If RsSONo.RecordCount = 0 Or (RsSONo.EOF = True Or RsSONo.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Col_SONo) = "" Then Exit Sub
        If FGrid.TextMatrix(FGrid.Row, Col_SONoCode) <> RsSONo!Code Then
            RsSONo.MoveFirst
            RsSONo.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, Col_SONo) & "'"
        End If
    Case Col_PNo
        If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
        RsPart.Sort = "Code"
        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
            RsPart.MoveFirst
            RsPart.FIND "Code='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "'"
            If RsPart.EOF = True Then RsPart.MoveFirst
        End If
    Case Col_Godown
        If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Col_Godown) = "" Then Exit Sub
        If FGrid.TextMatrix(FGrid.Row, Col_Godown) <> RsGodown!Name Then
            RsGodown.MoveFirst
            RsGodown.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, Col_Godown) & "'"
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
End Select
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then TxtGrid(0).TEXT = TxtGrid(0).Tag: Exit Sub
    Select Case FGrid.Col
        Case Col_SONo
            DGridTxtKeyDown DGSONo, TxtGrid, 0, RsSONo, KeyCode, True, 1
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, 2 '21
                End If
            End If
        Case Col_PNo
            If DGPart.Visible = False Then DGridColSwap DGPart, 0
            DGridTxtKeyDown DGPart, TxtGrid, 0, RsPart, KeyCode, True, 0, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, 4 '21, 3
                End If
            End If
        Case Col_MRP, Col_Taxable, Col_Qty, Col_PartSrlNo, Col_TaxAmt1, Col_SatPer
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo '21
                End If
            End If
            
        Case Col_TaxPer
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, , Col_Godown '21
                End If
            End If
            
        Case Col_DiscPer
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, , Col_Godown '21
                End If
            End If
            
'        Case Col_Qty
'            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
'                If TxtGridLeave = True Then
'                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Godown '21
'                Else
'                    TxtGrid_LostFocus 0
'                    TxtGrid(0).SetFocus
'                End If
'            End If
        Case Col_Rate
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, 2 '21, 2
                End If
            End If
        Case Col_DiscAmt
            
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                    If PubRestrict_Godown = 1 Then      ' Restrict Godown is "YES"
                        'Purpose not Clear, Redesign
                        GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Col_LName, , Col_PName
                    Else
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, 1 ' 21, 2
                    End If
                End If
            End If
        Case Col_Godown
            DGridTxtKeyDown DGGodown, TxtGrid, Index, RsGodown, KeyCode, True, 1, frmGodown, "frmGodown"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_SONo  ' 18
                End If
            End If
        Case Col_PName
            If DGPart.Visible = False Then DGridColSwap DGPart, 1
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 1, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Godown '21
                End If
            End If
        Case Col_LName
            If DGPart.Visible = False Then DGridColSwap DGPart, 2
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 2, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Godown, 1  '21, 1
                End If
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
If KeyAscii = vbKeyEscape Then Exit Sub
CheckQuote KeyAscii
Select Case FGrid.Col
    Case Col_SONo
        If DGSONo.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsSONo, KeyAscii, "Name"
    Case Col_PNo
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "Code"
    Case Col_PName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "Name"
    Case Col_LName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "LName"
    Case Col_Godown
        If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
            If DGGodown.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsGodown, KeyAscii, "Name"
        End If
    Case Col_Qty
        NumPress TxtGrid(Index), KeyAscii, 8, 3
   Case Col_DiscPer, Col_TaxPer, Col_SatPer
        NumPress TxtGrid(Index), KeyAscii, 2, 4
    Case Col_DiscAmt
        NumPress TxtGrid(Index), KeyAscii, 8, 2
    Case Col_Rate
        NumPress TxtGrid(Index), KeyAscii, 8, 4
    End Select
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case 0
    Select Case FGrid.Col
        Case Col_SONo
            If KeyCode <> 13 And DGSONo.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsSONo, KeyCode, "Name", True
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
        Case Col_MRP, Col_Taxable
            If TxtGrid(Index) <> "" Then
                If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
                    TxtGrid(Index) = ""
                ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Then
                    TxtGrid(Index) = "Yes"
                Else
                    TxtGrid(Index) = "No"
                End If
            End If
        Case Col_DiscPer, Col_TaxPer
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.0000")
        Case Col_DiscAmt
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        Case Col_Rate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.0000")
            FGrid.TextMatrix(FGrid.Row, Col_MRPRate) = Format(Val(TxtGrid(Index).TEXT), "0.0000")
 
        Case Col_Qty
            FGrid.TextMatrix(FGrid.Row, Col_Qty) = Format(Val(TxtGrid(Index).TEXT), "0.000")
            CountItem
        Case Col_PartSrlNo
            FGrid.TextMatrix(FGrid.Row, Col_PartSrlNo) = TxtGrid(Index)
    End Select
    Amt_Cal
    
End Select
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
    Grid_Hide
End If
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub TxtGrid_LostFocus(Index As Integer)
TxtGrid(0).Visible = False
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
        If TopCtrl1.TopText2 = "Add" Then
            If OptChal(0).Value = False And OptChal(1).Value = False Then
                MsgBox "Select Create / Select Challan Option", vbCritical, "Select/Create Challan"
                OptChal(0).SetFocus: Exit Sub
            End If
        End If
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
        'If OptChal(ChalCreate).Value = True Then 'Or Txt(DocType) = "Cash" Then
            Select Case FGrid.Col
                Case Col_MRP, Col_Taxable, Col_PartSrlNo
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                Case Col_SONo, Col_Qty, Col_Rate, Col_Amt
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    Amt_Cal
                Case Col_Godown
                    If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    End If
            End Select
        'End If
        Select Case FGrid.Col
            Case Col_DiscPer, Col_DiscAmt, Col_TaxPer
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                Amt_Cal
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        'If OptChal(ChalCreate).Value = True Then 'Or txt(DocType) = "Cash" Then
            Select Case FGrid.Col
                Case Col_SONo, Col_PNo, Col_PName, Col_LName
                    GridDblClick Me, FGrid, TxtGrid, 0
                Case Col_Taxable, Col_MRP, Col_Qty, Col_PartSrlNo
                    If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                        GridDblClick Me, FGrid, TxtGrid, 0
                    End If
                Case Col_Godown
                    If FGrid.TextMatrix(FGrid.Row, Col_Qty) <> "" Then
                        If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                            GridDblClick Me, FGrid, TxtGrid, 0
                        Else
                            GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Col_LName, , Col_PName
                        End If
                    End If
            End Select
        'End If
        Select Case FGrid.Col
            Case Col_Rate, Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer, Col_TaxAmt1
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    GridDblClick Me, FGrid, TxtGrid, 0
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
    If OptChal(ChalCreate).Value = True Then 'Or Txt(DocType) = "Cash" Then
        Select Case FGrid.Col
            Case Col_SONo, Col_PNo, Col_PName, Col_LName
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            Case Col_Unit
                FGrid.Col = FGrid.Col + 1
                FGrid.SetFocus
            Case Col_MRP, Col_Taxable
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
                End If
                If mVatYn = 1 And KeyAscii = 13 Then
                    If FGrid.Col = Col_MRP Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "No"
                        TxtGrid(0) = "No"
                    ElseIf FGrid.Col = Col_Taxable Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Yes"
                        TxtGrid(0) = "Yes"
                    End If
                Else
                    If Asc("Y") = KeyAscii Or Asc("y") = KeyAscii Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Yes"
                        TxtGrid(0) = "Yes"
                    ElseIf Asc("N") = KeyAscii Or Asc("n") = KeyAscii Then
                        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "No"
                        TxtGrid(0) = "No"
                    End If
                End If
            Case Col_PartSrlNo
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
                End If
            Case Col_Qty, Col_TaxPer, Col_SatPer
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
                End If
            Case Col_Godown
                If FGrid.TextMatrix(FGrid.Row, Col_Qty) <> "" Then
                    If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                        Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
                    End If
                End If
            
        End Select
    End If
    Select Case FGrid.Col
        Case Col_PNo
            Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        Case Col_Rate, Col_DiscPer, Col_DiscAmt, Col_TaxPer, Col_SatPer, Col_Qty, Col_Taxable, Col_MRP
            'If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
            'End If
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
    'If OptChal(ChalSelect).Value = True Then Exit Sub 'And Txt(DocType) <> "Cash" Then Exit Sub
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
                'Recalculate Footer Values
                If mVatYn = 1 Then
                   MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                        Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                        Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
                        Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
                        Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
                        Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
                        Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
                        Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
                        Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, Txt(SatAmt)
                Else
                    MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
                        Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
                        Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
                        Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
                        Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
                        Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
                        Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
                        Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
                        Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
                End If
                If UCase(left(PubComp_Name, 3)) = "JMK" Then
                    Txt(TurnOverAmt) = Format((Val(Txt(STotATB)) + Val(Txt(STaxAmt))) * Val(Txt(TurnOverPer)) / 100, "0.00")
                    Txt(NetSprAmt).TEXT = Format(Val(Txt(STotB)) + Val(Txt(TurnOverAmt)), "0.00")
                    Txt(NetAmt) = Format(Val(Txt(NetSprAmt)), "0.00")
                    Txt(NetAmt) = Format(Round(Val(Txt(NetSprAmt)), 0), "0.00")
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
                    Val(Txt(DiscPerTB)), Val(Txt(DiscPerTP)), _
                    Val(Txt(STaxPer)), Val(Txt(TaxSurPer)), Val(Txt(TurnOverPer))
    If mVatYn = 1 Then
       MainLib.SprCalcVAT NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Col_TaxPer, Col_TaxAmt1, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
            Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
            Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
            Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
            Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
            Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
            Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP, Col_SatPer, Col_SatAmt, Txt(SatAmt)
    Else
        MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
            Col_PNo, Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_ItemVal, Col_PartGrade, _
            Col_DiscAmt, Txt(IWDiscTotTB), Txt(IWDiscTotTP), Txt(MRPAmtTB), Txt(MRPAmtTP), _
            Txt(SprAmtTB), Txt(SprAmtTP), Txt(OilAmtTB), Txt(OilAmtTP), Txt(DiscPerTB), _
            Txt(DiscPerTP), Txt(DiscAmtTB), Txt(DiscAmtTP), Txt(STotATB), Txt(STotATP), _
            Txt(GenSurPer), Txt(GenSurAmt), Txt(TransAmt), Txt(TaxableTot), _
            Txt(STaxPer), Txt(STaxAmt), Txt(TaxSurPer), Txt(TaxSurAmt), Txt(PackCrg), _
            Txt(STotB), Txt(TurnOverPer), Txt(TurnOverAmt), Txt(ReSalTaxPer), Txt(ReSalTaxAmt), _
            Txt(SROff), Txt(NetSprAmt), Txt(NetAmt), mMRPTax, mMRPTaxSur, mMRPTOT, mMRPReSales, mMRPLubeTB, mMRPLubeTP
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

Private Sub FGridSel_Click()
    If FGridSel.TextMatrix(FGridSel.Row, SCol_SrNo) = "" Then
        FGridSel.Col = SCol_SrNo
        FGridSel.CellFontName = "wingdings"
        FGridSel.CellFontSize = 16
        FGridSel.TextMatrix(FGridSel.Row, SCol_SrNo) = "ü"
    Else
        FGridSel.TextMatrix(FGridSel.Row, SCol_SrNo) = ""
    End If
End Sub

Private Sub CmdSel_Click(Index As Integer)
Dim I As Integer, Cnt As Integer
Select Case Index
    Case 1
        For I = 1 To FGridSel.Rows - 1
            If FGridSel.TextMatrix(I, SCol_SrNo) = "ü" Then Cnt = Cnt + 1
        Next
        If Cnt = 0 Then
            MsgBox "Please Select at Least One Challan", vbInformation, "Validation"
            FGridSel.SetFocus
            Exit Sub
        End If
        FillChallan
        FrmSel.Visible = False
    Case 2
        FrmSel.Visible = False
End Select
End Sub

Private Sub SBILL_HD()
Dim mDocStr$, mDupStr$
''       mDocStr = IIf(left(mBILL, 2) = "S1", "INVOICE", IIf(left(mBILL, 2) = "S2", "CASHMEMO", IIf(left(mBILL, 2) = "T1", "TRANSFER MEMO", IIf(left(mBILL, 2) = "R1", "GOODS RETURN INWARD (CREDIT)", IIf(left(mBILL, 2) = "R2", "GOODS RETURN INWARD (CASH)", "GOODS RETURN INWARD (TRANSFER)")))))
'       mDocStr = IIf(mVType = SalCrVType, "Invoice", "Cash Memo")
''       mDupSTR=IIF(SPSALE- >PRINTED,' DUPLICATE','')
'
'        PRN_TIT TITL, "A", PageWidth, "18"
'        If Page <= 1 Then
'            If AUTHO_FOR <> "" Then
'                PRN_TIT AUTHO_FOR, "C", PageWidth, "18"
'            End If
'            PRN_TIT Add, "C", PageWidth, "18"
'            If Add1 <> "" Then
'                PRN_TIT Add1, "C", PageWidth, "18"
'            End If
'            If Add2 <> "" Then
'                PRN_TIT Add2, "C", PageWidth, "18"
'            End If
'            If Add3 <> "" Then
'                PRN_TIT Add3, "C", PageWidth, "18"
'            End If
'        End If
'
'        PRN_TIT "** " & mDocStr & mDupStr & " **", "A", PageWidth, "18"
'        If Page <= 1 Then
'            ?UPTT+SPACE(80-LEN(UPTT+'PHONE : '+PHONE))+'PHONE : '+PHONE
'            ?CST +SPACE(80-LEN(CST +'GRAM  : '+GRAM))  +'GRAM  : '+GRAM
'        End If
'        Print mChr18 + "To,"
'        Print "M/s " + mEmph + mName + Space(PageWidth - Len("M/s " + mName + mDocStr + " No. : " + SUBS(mBILL, 3, 8))) + mDOC_STR + " NO. : " + SUBS(mBILL, 3, 8) + mEmph1
'
'        Print IIf(mVType = SalCrVType,mADD1,SPACE(30))              +SPACE(PageWidth-LEN(mADD1+mDOC_STR+" DATE: "+DTOC(SPSALE- >DATE)))  +SPACE(LEN(mDOC_STR))+mEMPH+" DATE: "+DTOC(SPSALE- >DATE)+mEMPH1
'        Print IIf(mVType = SalCrVType,mADD2,SPACE(30))              +SPACE(PageWidth-LEN(mADD2+"ORDER NO.  : "+PSTR(SPSALE- >ORD_NO,8)))+IIF(LEFT(mBILL,2)="S1","ORDER NO.  : "+PSTR(SPSALE- >ORD_NO,8)," ")
'        Print IIf(mVType = SalCrVType,mCITY,SPACE(15))              +SPACE(PageWidth-LEN(mCITY+"      DATE : "+DTOC(SPSALE- >ORD_DT)))  +IIF(LEFT(mBILL,2)="S1","      DATE : "+DTOC(SPSALE- >ORD_DT)," ")
'        Print mDoub + "SALES TAX REF.NO." + mCST_NO + mDoub1
'        '---------------'
'
'        Print LN
'        If DET_OF_TAX Then
'            Print "SL.PART NO.     DESCRIPTION        QTY     RATE    RATE  DISC. DISC. <--A M O U N T- >"
'            Print "NO.                                                       (%)  AMT.  TAXPAID TAXABLE"
'        Else
'            Print "SL. PART NO.     DESCRIPTION      QTY. MRP RATE    RATE DISC %  DISC.      AMOUNT"
'        End If
'        Print LN
'        First = False
'        mRow = 18
'        mLine = 1
End Sub
Private Sub TxtGridValid_PNo()
'Called from TxtGrid_Validate & TxtGridLeave procedures

Dim OldPNo$
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
                FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(mPartyType, FGrid, CDate(Txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
'                FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsPart!SalDisc_Per, "0.00")
            End If
        End If
    End If
    If RSOJPR = True Then
        FIFOStkIss (RsPart!Code)
    End If
    If mVatYn = 1 Then
         Set rsTaxPer = GCn.Execute("Select Tax_Per, AddTaxPer, L_C from TaxForms where Form_Code='" & Txt(FormName).Tag & "'")
         If rsTaxPer.RecordCount > 0 Then
            FGrid.TextMatrix(FGrid.Row, Col_TaxPer) = rsTaxPer!Tax_Per
            FGrid.TextMatrix(FGrid.Row, Col_SatPer) = VNull(rsTaxPer!AddTaxPer)
            
                If UTrim(XNull(rsTaxPer!L_C)) = "LOCAL" Then
                   Set rsTaxPer = GCn.Execute("Select VatPer, AddTaxPer From Part_Grade Where PartGrade_Code = '" & FGrid.TextMatrix(FGrid.Row, Col_PartGrade) & "'")
                   If rsTaxPer.RecordCount > 0 Then
                       If VNull(rsTaxPer!VatPer) > 0 Then FGrid.TextMatrix(FGrid.Row, Col_TaxPer) = Format(rsTaxPer!VatPer, "0.00")
                       If VNull(rsTaxPer!AddTaxPer) > 0 Then FGrid.TextMatrix(FGrid.Row, Col_SatPer) = Format(rsTaxPer!AddTaxPer, "0.00")
                   End If
                End If
         End If
    End If
    
If FGrid.TextMatrix(FGrid.Row, Col_SONoCode) <> "" Then
    GSQL = "Select s1.Srl_No From SP_Order1 S1 Where OrderID='" & FGrid.TextMatrix(FGrid.Row, Col_SONoCode) & "' and Part_No='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "'"
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    If GRs.RecordCount > 1 Or GRs.RecordCount <= 0 Then
        Set GRs = New ADODB.Recordset
        GRs.CursorLocation = adUseClient
        GRs.Open "Select s1.Srl_No,s1.PART_NO,P.Part_Name,s1.QTY,s1.Sup_Qty,(s1.Qty-s1.Sup_Qty) As PendQty From SP_Order1 S1 Left Join Part P on S1.Part_no=P.Part_No and P.Div_Code = left(s1.orderid,1) Where OrderID='" & FGrid.TextMatrix(FGrid.Row, Col_SONoCode) & "'", GCn, adOpenStatic, adLockReadOnly
        Set DGOrdPart.DataSource = GRs
        GRs.FIND ("Part_No='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "'")
        If GRs.EOF Then
            GRs.MoveFirst
        End If
        DGOrdPart.Visible = True
        DGOrdPart.ZOrder 0
        DGOrdPart.SetFocus
    Else
        FGrid.TextMatrix(FGrid.Row, Col_SOSrNo) = GRs!Srl_No
        Set GRs = Nothing
        FGrid.SetFocus
        DGOrdPart.Visible = False
    End If
End If
    If FGrid.TextMatrix(FGrid.Rows - 1, Col_PNo) <> "" Then FGrid.AddItem FGrid.Rows
End Sub

Private Sub TxtGridValid_SONo()
If RsSONo.RecordCount = 0 Or (RsSONo.EOF = True Or RsSONo.BOF = True) Or TxtGrid(0).TEXT = "" Then
    FGrid.TextMatrix(FGrid.Row, Col_SONo) = ""
    FGrid.TextMatrix(FGrid.Row, Col_SONoCode) = ""
Else
    FGrid.TextMatrix(FGrid.Row, Col_SONoCode) = RsSONo!Code
    FGrid.TextMatrix(FGrid.Row, Col_SONo) = RsSONo!Name
End If
End Sub

Private Sub TxtGridValid_TaxMRP()
FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
    FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(mPartyType, FGrid, CDate(Txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
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

Private Sub ChallanCreate(ByRef ChalDocID As String, MyGPNo As String)
Dim ChalDocIDhlp$, VoucherEditFlag2 As Boolean, TmpNo As Double
Dim Rst As ADODB.Recordset, I As Integer
ChalDocID = GetDocID(GCnFaS, SalChalType, Txt(VDate), VoucherEditFlag2, Txt(SerialNo2), LblVPrefix2)
If GCn.Execute("Select Count(*) From SP_Sale Where DocID='" & ChalDocID & "' ").Fields(0) > 0 Then
    'SetMax_VoucherPrefix "DocID", SalChalType, "Sp_Sale"
    ChalDocID = GetDocID(GCnFaS, SalChalType, Txt(VDate), VoucherEditFlag2, Txt(SerialNo2), LblVPrefix2)
End If
ChalDocIDhlp = Replace(ChalDocID, " ", "")
        GCn.Execute "Insert Into SP_Sale(" _
            & "DocID ,DocIDHelp ,V_Type ,V_No ,Site_Code ," _
            & "V_Date ,Cash_Credit ,Party_Code ,Party_Name ,Address ," _
            & "L_C ,Form_Code ,RoadPermit_FormCode ,RoadPermit_No ,GR_RR_No ," _
            & "GR_RR_Date ,CrAc ,Case_No ,Case_Mark ,Mode_Dispatch ," _
            & "Transport,Rep_Code ,Remarks ,Det_Tax ,SprAmt_MRP_TB ," _
            & "SprAmt_MRP_TP,OilAmt_MRP_TB,OilAmt_MRP_TP,SprAmt_TB ,SprAmt_TP ,OilAmt_TB ,OilAmt_TP ," _
            & "D_Per_TB ,D_Amt_TB ,D_Per_TP ,D_Amt_TP ,Addition ," _
            & "Packing ,Gen_Sur_Per ,Gen_Sur_Amt ,Trans_Amt ,Tax_Per ," _
            & "Tax_Amt ,Tax_Sur_Per ,Tax_Sur_Amt ,TOT_Per ,TOT_Amt ," _
            & "ReSalTax_Per, ReSalTax_Amt,Rounded ,Total_Amt,GP_No, GP_Date, Invoice_DocId ,U_Name ,U_EntDt ,U_AE, AddBy, AddDate, " _
            & "D_Per_MRP_TB,D_Amt_MRP_TB, D_Per_MRP_TP , D_Amt_MRP_TP, Tax_AmtMRP, TaxSur_AmtMRP, TOT_AmtMRP, SatAmt, Sat_Yn ) " _
            & "Values('" & ChalDocID & "','" & ChalDocIDhlp & "','" & SalChalType & "'," & Txt(SerialNo2) & ",'" & PubSiteCode & ForSiteCode & _
            "', " & ConvertDate(Txt(VDate)) & ",'" & Txt(DocType) & "','" & Txt(Party).Tag & "','" & Txt(Party) & "','" & Txt(Address1) & _
            "','" & left(Txt(LC), 1) & "','" & Txt(FormName).Tag & "','" & Txt(Form31Name).Tag & "','" & Txt(Form31No) & "','" & Txt(LRNo) & _
            "', " & ConvertDate(Txt(LRDate)) & ",'" & Txt(CrAc).Tag & "'," & Val(Txt(CaseNo)) & ",'" & Txt(CaseMark) & "','" & Txt(DispMode) & _
            "','" & Txt(Transport) & "','" & Txt(SPerson).Tag & "','" & Txt(Remark) & "'," & IIf(Txt(TaxDet) = "Yes", 1, 0) & "," & Val(Txt(MRPAmtTB)) & _
            " , " & Val(Txt(MRPAmtTP)) & "," & mMRPLubeTB & "," & mMRPLubeTP & "," & Val(Txt(SprAmtTB)) & "," & Val(Txt(SprAmtTP)) & "," & Val(Txt(OilAmtTB)) & "," & Val(Txt(OilAmtTP)) & _
            " , " & Val(Txt(DiscPerTB)) & "," & Val(Txt(DiscAmtTB)) & "," & Val(Txt(DiscPerTP)) & "," & Val(Txt(DiscAmtTP)) & "," & Val(Txt(Addition)) & _
            " , " & Val(Txt(PackCrg)) & "," & Val(Txt(GenSurPer)) & "," & Val(Txt(GenSurAmt)) & "," & Val(Txt(TransAmt)) & "," & Val(Txt(STaxPer)) & _
            " , " & Val(Txt(STaxAmt)) & "," & Val(Txt(TaxSurPer)) & "," & Val(Txt(TaxSurAmt)) & "," & Val(Txt(TurnOverPer)) & "," & Val(Txt(TurnOverAmt)) & _
            " , " & Val(Txt(ReSalTaxPer)) & "," & Val(Txt(ReSalTaxAmt)) & "," & Val(Txt(SROff)) & "," & Val(Txt(NetAmt)) & ",'" & MyGPNo & "', " & ConvertDate(IIf(MyGPNo = "", "", PubServerDate)) & ", '" & Txt(DocID) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & _
            ",'A', '" & pubUName & "', " & ConvertDateTime(PubServerDate) & "," & mMRevDisTBPer & "," & mTBDisAmtMRP & "," & mMRevDisTPPer & "," & mTPDisAmtMRP & ", " & _
            "" & mMRPTax & "," & mMRPTaxSur & ", " & mMRPTOT & ", " & Val(Txt(SatAmt)) & ", " & IIf(mSatYn, 1, 0) & " )"
            ' Update Sale Order
        UpdateSO ChalDocID
'        GCn.Execute ("Delete From SP_Stock Where DocID='" & Txt(DocID) & "'")
        ChallanItemAdd ChalDocID
        'Voucher Serial No. Updation LPS 21-05-03
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaS, ChalDocID, Txt(VDate)
End Sub

Private Sub ChallanEdit(ByRef mInvDocID As String)
Dim ChalVPrefix As String, ChalDocID$, I As Integer

        'For Challan Header Record
        GSQL = "Select DocId from SP_Sale where Invoice_DocId='" & mInvDocID & "'"
        ChalDocID = GCn.Execute(GSQL).Fields(0).Value
        GCn.Execute "Update SP_Sale Set " _
            & " Site_Code='" & PubSiteCode & ForSiteCode & _
            "', V_Date=" & ConvertDate(Txt(VDate).TEXT) & ",Party_Code='" & Txt(Party).Tag & "', Party_Name='" & Txt(Party).TEXT & "', Address='" & Txt(Address1).TEXT & _
            "', L_C='" & left(Txt(LC).TEXT, 1) & "', Form_Code='" & Txt(FormName).Tag & "', RoadPermit_FormCode='" & Txt(Form31Name).Tag & "', RoadPermit_No='" & Txt(Form31No).TEXT & _
            "', GR_RR_No='" & Txt(LRNo).TEXT & "', GR_RR_Date=" & ConvertDate(Txt(LRDate).TEXT) & ", CrAc=" & Val(Txt(CaseNo).TEXT) & ", Case_No=" & Val(Txt(CaseNo).TEXT) & _
            " , Case_Mark='" & Txt(CaseMark).TEXT & "', Mode_Dispatch='" & Txt(DispMode).TEXT & "',Transport='" & Txt(Transport).TEXT & "', Rep_Code='" & Txt(SPerson).Tag & _
            "', Remarks='" & Txt(Remark).TEXT & "', Det_Tax=" & IIf(Txt(TaxDet) = "Yes", 1, 0) & ", SprAmt_MRP_TB=" & Val(Txt(MRPAmtTB).TEXT) & _
            " , SprAmt_MRP_TP=" & Val(Txt(MRPAmtTP).TEXT) & ", SprAmt_TB=" & Val(Txt(SprAmtTB).TEXT) & ", SprAmt_TP=" & Val(Txt(SprAmtTP).TEXT) & _
            " , OilAmt_TB=" & Val(Txt(OilAmtTB).TEXT) & ", OilAmt_TP=" & Val(Txt(OilAmtTP).TEXT) & ", D_Per_TB=" & Val(Txt(DiscPerTB).TEXT) & _
            " , D_Amt_TB=" & Val(Txt(DiscAmtTB).TEXT) & ", D_Per_TP=" & Val(Txt(DiscPerTP).TEXT) & " , D_Amt_TP=" & Val(Txt(DiscAmtTP).TEXT) & _
            " , Addition=" & Val(Txt(Addition).TEXT) & " , Packing=" & Val(Txt(PackCrg).TEXT) & ", Gen_Sur_Per=" & Val(Txt(GenSurPer).TEXT) & _
            " , Gen_Sur_Amt=" & Val(Txt(GenSurAmt).TEXT) & " , Trans_Amt=" & Val(Txt(TransAmt).TEXT) & " ,Tax_Per=" & Val(Txt(STaxPer).TEXT) & _
            " , Tax_Amt=" & Val(Txt(STaxAmt).TEXT) & ", Tax_Sur_Per=" & Val(Txt(TaxSurPer).TEXT) & ",Tax_Sur_Amt=" & Val(Txt(TaxSurAmt).TEXT) & _
            " , TOT_Per=" & Val(Txt(TurnOverPer).TEXT) & ",TOT_Amt=" & Val(Txt(TurnOverAmt).TEXT) & " ,ReSalTax_Per = " & Val(Txt(ReSalTaxPer)) & _
            " , ReSalTax_Amt = " & Val(Txt(ReSalTaxAmt)) & ", Rounded=" & Val(Txt(SROff).TEXT) & ", Total_Amt=" & Val(Txt(NetAmt).TEXT) & _
            " , U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='A', D_Per_MRP_TB=" & mMRevDisTBPer & ",D_Amt_MRP_TB=" & mTBDisAmtMRP & _
            " , D_Per_MRP_TP=" & mMRevDisTPPer & ", D_Amt_MRP_TP=" & mTPDisAmtMRP & ", Tax_AmtMRP=" & mMRPTax & ", TaxSur_AmtMRP=" & mMRPTaxSur & ", TOT_AmtMRP= " & mMRPTOT & _
            " where DocID='" & ChalDocID & "'"
            ' Update Sale Order
        UpdateSO ChalDocID
        GCn.Execute ("Delete From SP_Stock Where DocID='" & Txt(DocID) & "'")
        GCn.Execute ("Delete From SP_Stock Where DocID='" & ChalDocID & "'")
        ChallanItemAdd ChalDocID
    End Sub

Private Sub ChallanItemAdd(ChalDocID As String)
Dim I As Integer
For I = 1 To FGrid.Rows - 1
    If FGrid.TextMatrix(I, Col_PNo) <> "" And Val(FGrid.TextMatrix(I, Col_Qty)) <> 0 Then
        If FGrid.TextMatrix(I, Col_SONo) <> "" Then
            GCn.Execute "Update SP_Order1 Set Sup_Qty=Sup_Qty+" & Val(FGrid.TextMatrix(I, Col_Qty)) & " Where OrderId='" & FGrid.TextMatrix(I, Col_SONoCode) & "' and Srl_No=" & Val(FGrid.TextMatrix(I, Col_SOSrNo)) & ""
        End If
        GCn.Execute "Insert Into SP_Stock(" _
            & "DocID,Srl_No,V_Type,V_No,V_Date,Site_Code," _
            & "Party_Code,L_C,Order_DocId,Order_Srl_No,Part_No," _
            & "Godown,Qty_Iss,Tax_YN,MRP_YN,Rate," _
            & "MRP_Rate,Disc_Per,Disc_Amt,Amount,Net_Amt," _
            & "Invoice_DocId,V_Date2,Rate2,MRP_Rate2,Disc_Per2," _
            & "Disc_Amt2,Amount2,Net_Amt2,Part_SrlNo,U_Name,U_EntDt,U_AE,TaxPer,TaxAmt, SatPer, SatAmt,PurDocNo,PurDocDate) " _
            & "Values('" & ChalDocID & "'," & I & ",'" & SalChalType & "'," & DeCodeDocID(ChalDocID, Document_No) & "," & ConvertDate(Format(Txt(VDate).TEXT, "dd/MMM/yyyy")) & ",'" & PubSiteCode & ForSiteCode & _
            "','" & Txt(Party).Tag & "','" & left(Txt(LC).TEXT, 1) & "','" & FGrid.TextMatrix(I, Col_SONoCode) & "'," & Val(FGrid.TextMatrix(I, Col_SOSrNo)) & ",'" & FGrid.TextMatrix(I, Col_PNo) & _
            "','" & FGrid.TextMatrix(I, Col_GodownCode) & "'," & Val(FGrid.TextMatrix(I, Col_Qty)) & "," & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & "," & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & "," & Val(FGrid.TextMatrix(I, Col_Rate)) & _
            " , " & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," & Val(FGrid.TextMatrix(I, Col_DiscPer)) & "," & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," & Val(FGrid.TextMatrix(I, Col_Amt)) & "," & Val(FGrid.TextMatrix(I, Col_ItemVal)) & _
            " ,'" & Txt(DocID).TEXT & "'," & ConvertDate(Txt(VDate).TEXT) & "," & Val(FGrid.TextMatrix(I, Col_Rate)) & "," & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," & Val(FGrid.TextMatrix(I, Col_DiscPer)) & _
            " , " & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," & Val(FGrid.TextMatrix(I, Col_Amt)) & "," & Val(FGrid.TextMatrix(I, Col_ItemVal)) & ",'" & FGrid.TextMatrix(I, Col_PartSrlNo) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A'," & Val(FGrid.TextMatrix(I, Col_TaxPer)) & "," & Val(FGrid.TextMatrix(I, Col_TaxAmt1)) & "," & Val(FGrid.TextMatrix(I, Col_SatPer)) & "," & Val(FGrid.TextMatrix(I, Col_SatAmt)) & ",'" & FGrid.TextMatrix(I, Col_PurDocId) & "'," & ConvertDate(FGrid.TextMatrix(I, Col_PurDate)) & ")"
        Call UpdStkGridToTable(FGrid.TextMatrix(I, Col_PNo), "-", FGrid.TextMatrix(I, Col_MRP), FGrid.TextMatrix(I, Col_Taxable), FGrid.TextMatrix(I, Col_Qty))
    End If
Next
End Sub

Private Sub ProcAcPost(rsCtrlAc As ADODB.Recordset)
'On Error GoTo lblExit
Dim xMRPSprTp As Double, xMRPOilTp As Double
Dim xSprTp As Double, xOilTp As Double
Dim mShare As Single, mShareAmt As Double, mShare2Amt As Double
Dim xNetAmt As Double, xRoundAmt As Double, xSprAmtMRPTB As Double, xSprAmtMRPTP As Double
Dim xOilAmtMRPTB As Double, xOilAmtMRPTP As Double
Dim xSprAmtTB  As Double, xSprAmtTP As Double, xOilAmtTB As Double, xOilAmtTP As Double
Dim xDisAmtTB As Double, xDisAmtTP As Double, xDisAmtMRPTB As Double, xDisAmtMRPTP As Double
Dim xGenSurAmt As Double, xTrans As Double, xTaxAmt As Double, xTaxAmtMRP As Double, xPack As Double
Dim xTurnOver, xTurnOverMrp As Double, xReSaleTaxAmt As Double, mFADocID$, mQry$, xSatAmt As Double
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
Dim mSId As Double

    ApplyConsolidatedPosting CDate(Txt(VDate))

    TaxSQL = "select TF.Tax_Ac_Code,TF.Sur_Ac_Code, Tf.AddTaxAc,sum(Tax_Amt+Tax_AmtMRP) as TaxAmt,sum(Tax_Sur_Amt+TaxSur_AmtMRP) as TaxSurAmt,Sum(SatAmt) As SatAmt " & _
        " from SP_Sale left join TaxFormsAc as TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code"
    
    If Txt(DocType) = "Cash" And IsConsolidatedPosting Then
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
            "where V_Date=" & ConvertDate(Txt(VDate)) & " and left(docid,8)='" & left(Txt(DocID), 8) & _
            "' group by TF.PurSal_Ac_Code"
            
        mQry = "select " & _
            "sum(Total_Amt) as NetAmt,sum(round(rounded,2)) as RoundAmt," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(SprAmt_MRP_TP) as SprAmtMRPTP, " & _
            "sum(OilAmt_MRP_TB) as OilAmtMRPTB, sum(OilAmt_MRP_TP) as OilAmtMRPTP, " & _
            "sum(SprAmt_TB) as SprAmtTB, sum(SprAmt_TP) as SprAmtTP, " & _
            "sum(OilAmt_TB) as OilAmtTB, sum(OilAmt_TP) as OilAmtTP, " & _
            "sum(D_Amt_TB) as DisAmtTB, sum(D_Amt_TP) as DisAmtTP, " & _
            "sum(D_Amt_MRP_TB) as DisAmtMRPTB, sum(D_Amt_MRP_TP) as DisAmtMRPTP," & _
            "sum(Gen_Sur_Amt) as GenSurAmt,sum(Trans_Amt) as Trans," & _
            "sum(Tax_Amt+Tax_Sur_Amt+Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmt," & _
            "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, " & cIIF(cUCase("left('" & PubComp_Name & "',3)") & "='JMK'", "sum(TOT_Amt)", "sum(TOT_Amt+TOT_AmtMrp)") & " as TurnOver, sum(TOT_AmtMrp) as  Tot_AmtMrp, Sum(SatAmt) as SatAmt, " & _
            "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
            "where V_Date=" & ConvertDate(Txt(VDate)) & " and left(docid,8)='" & left(Txt(DocID), 8) & "'"
        'for tax
        TaxSQL = TaxSQL & " where  V_Date= " & ConvertDate(Txt(VDate)) & " and left(docid,8)='" & left(Txt(DocID), 8) & _
            "' Group by TF.Tax_Ac_Code,TF.Sur_Ac_Code, Tf.AddTaxAc"
        mNarr = "Through Counter Cash Sale (Daily Posting)"
        mCommNarr = mNarr & " [Common]"
        mFADocID = left(Txt(DocID), 8) & "XXXXX" & "  " & Format(Txt(VDate), "yymmdd")
        PartyCode = PubSprCashAc
   Else
        PartyCode = Txt(Party).Tag
        mFADocID = Txt(DocID)
        mNarr = "Through Counter Cr Sale"
        mCommNarr = mNarr & " [Common]"
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "left join TaxFormsAc TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
            "where docid='" & Txt(DocID) & _
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
            "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, " & cIIF(cUCase("left('" & PubComp_Name & "',3)") & "='JMK'", "sum(TOT_Amt)", "sum(TOT_Amt+TOT_AmtMrp)") & " as TurnOver, sum(TOT_AmtMrp) as  Tot_AmtMrp, Sum(SatAmt) as SatAmt, " & _
            "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
            "from SP_Sale " & _
            "left join TaxFormsAc as TF on Sp_Sale.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code " & _
            "where docid='" & Txt(DocID) & "'"
        'for tax
        TaxSQL = TaxSQL & " where docid='" & Txt(DocID) & _
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
    xSatAmt = XNull(rsTemp1!SatAmt)
    xPack = IIf(IsNull(rsTemp1!Pack), 0, rsTemp1!Pack)
    If UCase(left(PubComp_Name, 3)) <> "JMK" Then
        xTurnOver = IIf(IsNull(rsTemp1!TurnOver), 0, rsTemp1!TurnOver) - IIf(IsNull(rsTemp1!Tot_AmtMrp), 0, rsTemp1!Tot_AmtMrp)
        xTurnOverMrp = IIf(IsNull(rsTemp1!Tot_AmtMrp), 0, rsTemp1!Tot_AmtMrp)
    Else
        xTurnOver = IIf(IsNull(rsTemp1!TurnOver), 0, rsTemp1!TurnOver)
    End If
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
                mShareAmt = mShareAmt + (xDisAmtMRPTB - mTot1ShareAmt)
                mShare2Amt = mShare2Amt + (xTaxAmtMRP - mTot2ShareAmt)
            End If
            If UCase(left(PubComp_Name, 3)) = "JMK" Then
                mSprAmtMRPTB = mSprAmtMRPTB - (mShareAmtSpr + mShare2AmtSpr)
            Else
                mSprAmtMRPTB = mSprAmtMRPTB - (mShareAmtSpr + mShare2AmtSpr + xTurnOverMrp)
            End If
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
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(mSprAmtMRPTB + mSprAmtTB, 2)
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

'   0.Party A/c or Cash A/c
'   1.Taxable Spr = MRP Spr TB + SPR TB
'   2.Taxpaid Spr = MRP Spr TP + SPR TP
'   3.Taxable Oil = MRP Oil TB + Oil TB
'   4.Taxable Oil = MRP Oil TP + Oil TP
'   5.xGenSurAmt
'   6.xPack
'   7.xTurnOver
'   8.xReSaleTaxAmt
    '*******
     'Sale Party A/c
    'I = 0
    LedgAry(0).SubCode = PartyCode
    LedgAry(0).AmtDr = Round(xNetAmt, 2)
    LedgAry(0).AmtCr = 0
    LedgAry(0).Narration = mNarr
    'Spare Sale A/c Taxpaid
    If xMRPSprTp + xSprTp <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!SprSalTP_Ac
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(xMRPSprTp + xSprTp, 2)
        LedgAry(I).Narration = mNarr ' & " Spare"
    End If
    'Oil Sale A/c Taxable
    If mTotMRPOilTB + mTotOilTB <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!OilSalTB_Ac
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mTotMRPOilTB + mTotOilTB, 2)
        LedgAry(I).Narration = mNarr ' & " Spare"
    End If
     'Oil Sale A/c Taxpaid
     If xMRPOilTp + xOilTp <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!OilSalTP_Ac
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(xMRPOilTp + xOilTp, 2)
        LedgAry(I).Narration = mNarr ' & " Spare"
     End If
      'GenSurAmt
     If xGenSurAmt <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!SprGenSur_Ac
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(xGenSurAmt, 2)
        LedgAry(I).Narration = mNarr ' & " Sale Tax"
     End If
    'Transportation
     If xTrans <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!Transportation_Ac
        If xTrans > 0 Then
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(xTrans, 2)
        Else
            LedgAry(I).AmtDr = Round(Abs(xTrans), 2)
           LedgAry(I).AmtCr = 0
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
                    LedgAry(I).AmtDr = 0
                    LedgAry(I).AmtCr = Round(RsTemp!TaxAmt, 2)
                Else
                    LedgAry(I).AmtDr = Round(Abs(RsTemp!TaxAmt), 2)
                    LedgAry(I).AmtCr = 0
                End If
                 LedgAry(I).Narration = mNarr '& " Sales Tax & Surcharge"
            End If
             If RsTemp!SatAmt <> 0 Then
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
                LedgAry(I).SubCode = RsTemp!AddTaxAc
                If RsTemp!TaxAmt > 0 Then
                    LedgAry(I).AmtDr = 0
                    LedgAry(I).AmtCr = Round(RsTemp!SatAmt, 2)
                Else
                    LedgAry(I).AmtDr = Round(Abs(RsTemp!SatAmt), 2)
                    LedgAry(I).AmtCr = 0
                End If
                 LedgAry(I).Narration = mNarr '& " Sales Tax & Surcharge"
            End If
            
            If RsTemp!TaxSurAmt <> 0 Then
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
                LedgAry(I).SubCode = RsTemp!Sur_Ac_Code
                If RsTemp!TaxSurAmt > 0 Then
                    LedgAry(I).AmtDr = 0
                    LedgAry(I).AmtCr = Round(RsTemp!TaxSurAmt, 2)
                Else
                    LedgAry(I).AmtDr = Round(Abs(RsTemp!TaxSurAmt), 2)
                    LedgAry(I).AmtCr = 0
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
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(xPack, 2)
        Else
            LedgAry(I).AmtDr = Round(Abs(xPack), 2)
            LedgAry(I).AmtCr = 0
        End If
        LedgAry(I).Narration = mNarr '& " Misc Charges"
    End If
    If xTurnOver <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!TOTax_Ac
        If xTurnOver > 0 Then
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(xTurnOver, 2)
        Else
            LedgAry(I).AmtDr = Round(Abs(xTurnOver), 2)
            LedgAry(I).AmtCr = 0
        End If
        LedgAry(I).Narration = mNarr '& " TurnOver Amt"
    End If
    If xReSaleTaxAmt <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!ReSaleTax_Ac
        If xReSaleTaxAmt > 0 Then
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(xReSaleTaxAmt, 2)
        Else
            LedgAry(I).AmtDr = Round(Abs(xReSaleTaxAmt), 2)
            LedgAry(I).AmtCr = 0
        End If
        LedgAry(I).Narration = mNarr '& " ReSale Tax Amount"
    End If
    If xRoundAmt <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!SprROff_Ac
        If xRoundAmt > 0 Then
            LedgAry(I).AmtDr = 0
            LedgAry(I).AmtCr = Round(xRoundAmt, 2)
        Else
            LedgAry(I).AmtDr = Round(Abs(xRoundAmt), 2)
            LedgAry(I).AmtCr = 0
        End If
        LedgAry(I).Narration = mNarr '& " Round Off"
    End If
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, mFADocID, CDate(Txt(VDate)), mCommNarr)
    If mResult <> 1 Then
        MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
    End If
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
                Txt(VDate).Tag = Txt(VDate).TEXT
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
If UCase(left(PubComp_Name, 3)) = "LMP" Then
    GSQL = "SELECT s.DocID,s.V_Type,s.V_No,s.V_Date,s.Cash_Credit,s.Job_DocID,s.Party_Code,s.Party_Name,s.Address,s.L_C,s.REP_CODE, s.Form_Code,s.RoadPermit_FormCode," & _
        "s.GR_RR_No, s.GR_RR_Date,s.CrAc, s.Case_No, s.Case_Mark, s.Mode_Dispatch, s.Transport, s.Remarks,s.SprAmt_MRP_TB,s.SprAmt_MRP_TP," & _
        "s.OilAmt_MRP_TB,s.OilAmt_MRP_TP,s.SprAmt_TB,s.SprAmt_TP, s.OilAmt_TB, s.OilAmt_TP, s.D_Per_TB, s.D_Amt_TB, s.D_Per_TP,s.D_Amt_TP," & _
        "s.D_Per_MRP_TB,s.D_Amt_MRP_TB,s.D_Per_MRP_TP,s.D_Amt_MRP_TP,s.Addition, s.Gen_Sur_Per, s.Gen_Sur_Amt, s.Trans_Amt, s.LineFileTaxSum," & _
        "s.Tax_Per, s.Tax_Amt, s.Tax_AmtMRP, s.Tax_Sur_Per,s.Tax_Sur_Amt,s.TaxSur_AmtMRP,s.Packing, s.TOT_Per, s.Tot_Amt, s.TOT_AmtMRP," & _
        "s.ReSalTax_Per, s.ReSalTax_Amt,s.Total_Amt, s.Rounded, s.Det_Tax,s.GP_No,s.GP_Date,s.Printed_YN,s.Invoice_DocId, s.U_Name, s.U_EntDt,S.CancelYN," & _
        "" & vIsNull("SPStk.Srl_No", "0") & " as Srl_No, " & xIsNull("SPStk.V_Date", "") & " as SPStk_V_Date, " & xIsNull("SPStk.Party_Code", "") & " as SPStkParty_Code,SPStk.L_C As SPStk_L_C," & _
        "" & xIsNull("SPStk.Job_DocID", "") & " as Job_DocID,SPStk.Mech_Code, SPStk.Order_DocId,SPStk.Order_Srl_No,SPStk.Part_No, SPStk.Lub_Category, SPStk.Godown," & _
        "" & vIsNull("SPStk.Qty_Doc", "0") & " as Qty_Doc," & vIsNull("SPStk.Qty_Rec", "0") & " as Qty_Rec, " & vIsNull("SPStk.Qty_Iss", "0") & " as Qty_Iss," & _
        "" & vIsNull("SPStk.Qty_Ret", "0") & " as Qty_Ret, " & vIsNull("SPStk.Tax_YN", "0") & " as Tax_YN, " & vIsNull("SPStk.MRP_YN", "0") & " as MRP_YN," & _
        "" & vIsNull("SPStk.Rate", "0") & " as Rate, " & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate, " & vIsNull("SPStk.Disc_Per", "0") & " as Disc_Per," & _
        " " & vIsNull("SPStk.Disc_Amt", "0") & " as Disc_Amt, " & vIsNull("SPStk.AMOUNT", "0") & " as AMOUNT," & vIsNull("SPStk.Ord_DiscPer", "0") & " as Ord_DiscPer," & _
        " " & vIsNull("SPStk.Ord_DiscAmt", "0") & " as Ord_DiscAmt, " & vIsNull("SPStk.Net_Amt", "0") & " as Net_Amt," & xIsNull("SPStk.Purpose", "") & " as Purpose," & _
        "SPStk.Part_SrlNo,SPStk.Remark,SPStk.Invoice_DocId as SPStk_Invoice_DocId, SPStk.V_Date2," & vIsNull("SPStk.Rate2", "0") & " as Rate2, " & vIsNull("SPStk.MRP_Rate2", "0") & " as MRP_Rate2," & _
        "" & vIsNull("SPStk.Disc_Per2", "0") & " as Disc_Per2, " & vIsNull("SPStk.Disc_Amt2", "0") & " as Disc_Amt2, " & vIsNull("SPStk.Amount2", "0") & " as Amount2," & _
        "" & vIsNull("SPStk.Ord_DiscPer2", "0") & " as Ord_DiscPer2, " & vIsNull("SPStk.Ord_DiscAmt2", "0") & " as Ord_DiscAmt2, " & vIsNull("SPStk.Net_Amt2", "0") & " as Net_Amt2,SPStk.Printed2, " & _
        "" & vIsNull("SPStk.TaxPer", "0") & " as TaxPer," & vIsNull("SPStk.TaxAmt", "0") & " as TaxAmt, sOrd.V_Date As OrderDate,Part.Part_Name," & xIsNull("Part.Unit", "") & " As PartUnit, " & _
        " SG.SiebelCode,SG.NamePrefix,SG.Name,SG.Add1, SG.Add2, SG.Add3,SG.PIN,SG.Phone,SG.LSTNo,SG.CSTNo,SG.RC_No,City.CityName,TF.Printing_Desc,Syctrl.GatePassOnSprInv,Syctrl.SprInvFooter,'" & CustOrdDet & "' as CustOrdDet,'' As X " & _
        " FROM ((((((SP_Sale as S left JOIN SP_Stock as SPStk ON S.DocID = SPStk.Invoice_DocId) " & _
        "left JOIN Part ON SPStk.Part_No = Part.PART_NO and Part.Div_Code = left(SPStk.Docid,1)) " & _
        "LEFT JOIN SubGroup as SG ON S.Party_Code = SG.SubCode) LEFT JOIN City ON SG.CityCode = City.CityCode)" & _
        "Left Join TaxForms TF on S.Form_Code=TF.Form_Code) " & _
        "Left Join Sp_Order sOrd On SPStk.Order_DocId = sOrd.OrderId) " & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable<>S.U_AE " & _
        "where S.DocId='" & Master!SearchCode & "' Order By SPStk.Srl_No"
Else
    GSQL = "SELECT s.DocID,s.V_Type,s.V_No,s.V_Date,s.Cash_Credit,s.Job_DocID,s.Party_Code,s.Party_Name,s.Address," & _
        "SG.NamePrefix,SG.Name,SG.Add1, SG.Add2, SG.Add3,City.CityName,SG.PIN,SG.Phone,SG.CSTNo,SG.RC_No,s.L_C,s.REP_CODE, s.Form_Code,s.RoadPermit_FormCode," & _
        "s.GR_RR_No, s.GR_RR_Date,s.CrAc, s.Case_No, s.Case_Mark, s.Mode_Dispatch, s.Transport, s.Remarks,s.SprAmt_MRP_TB,s.SprAmt_MRP_TP," & _
        "s.OilAmt_MRP_TB,s.OilAmt_MRP_TP,s.SprAmt_TB,s.SprAmt_TP, s.OilAmt_TB, s.OilAmt_TP, s.D_Per_TB, s.D_Amt_TB, s.D_Per_TP,s.D_Amt_TP," & _
        "s.D_Per_MRP_TB,s.D_Amt_MRP_TB,s.D_Per_MRP_TP,s.D_Amt_MRP_TP,s.Addition, s.Gen_Sur_Per, s.Gen_Sur_Amt, s.Trans_Amt, s.LineFileTaxSum," & _
        "s.Tax_Per, s.Tax_Amt, s.Tax_AmtMRP, s.Tax_Sur_Per,s.Tax_Sur_Amt,s.TaxSur_AmtMRP,s.Packing, s.TOT_Per, s.Tot_Amt, s.TOT_AmtMRP," & _
        "s.ReSalTax_Per, s.ReSalTax_Amt,s.Total_Amt, s.Rounded, s.Det_Tax,s.GP_No,s.GP_Date,s.Printed_YN,s.Invoice_DocId, s.U_Name, s.U_EntDt,S.CancelYN," & _
        "" & vIsNull("SPStk.Srl_No", "0") & " as Srl_No, " & xIsNull("SPStk.V_Date", "") & " as SPStk_V_Date, " & xIsNull("SPStk.Party_Code", "") & " as SPStkParty_Code,SPStk.L_C," & _
        "" & xIsNull("SPStk.Job_DocID", "") & " as Job_DocID,SPStk.Mech_Code, SPStk.Order_DocId,SPStk.Order_Srl_No,SPStk.Part_No,Part.Part_Name, SPStk.Lub_Category, SPStk.Godown," & _
        "" & vIsNull("SPStk.Qty_Doc", "0") & " as Qty_Doc," & vIsNull("SPStk.Qty_Rec", "0") & " as Qty_Rec, " & vIsNull("SPStk.Qty_Iss", "0") & " as Qty_Iss," & _
        "" & vIsNull("SPStk.Qty_Ret", "0") & " as Qty_Ret, " & vIsNull("SPStk.Tax_YN", "0") & " as Tax_YN, " & vIsNull("SPStk.MRP_YN", "0") & " as MRP_YN," & _
        "" & vIsNull("SPStk.Rate", "0") & " as Rate, " & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate, " & vIsNull("SPStk.Disc_Per", "0") & " as Disc_Per," & _
        " " & vIsNull("SPStk.Disc_Amt", "0") & " as Disc_Amt, " & vIsNull("SPStk.AMOUNT", "0") & " as AMOUNT," & vIsNull("SPStk.Ord_DiscPer", "0") & " as Ord_DiscPer," & _
        " " & vIsNull("SPStk.Ord_DiscAmt", "0") & " as Ord_DiscAmt, " & vIsNull("SPStk.Net_Amt", "0") & " as Net_Amt," & xIsNull("SPStk.Purpose", "") & " as Purpose," & _
        "SPStk.Part_SrlNo,SPStk.Remark,SPStk.Invoice_DocId as SPStk_Invoice_DocId, SPStk.V_Date2," & vIsNull("SPStk.Rate2", "0") & " as Rate2, " & vIsNull("SPStk.MRP_Rate2", "0") & " as MRP_Rate2,'' As X,Syctrl.SprInvFooter," & _
        "" & vIsNull("SPStk.Disc_Per2", "0") & " as Disc_Per2, " & vIsNull("SPStk.Disc_Amt2", "0") & " as Disc_Amt2, " & vIsNull("SPStk.Amount2", "0") & " as Amount2," & _
        "" & vIsNull("SPStk.Ord_DiscPer2", "0") & " as Ord_DiscPer2, " & vIsNull("SPStk.Ord_DiscAmt2", "0") & " as Ord_DiscAmt2, " & vIsNull("SPStk.Net_Amt2", "0") & " as Net_Amt2,SPStk.Printed2,TF.Printing_Desc,Syctrl.GatePassOnSprInv, " & _
        "'" & CustOrdDet & "' as CustOrdDet,SG.LSTNo, " & vIsNull("SPStk.TaxPer", "0") & " as TaxPer," & vIsNull("SPStk.TaxAmt", "0") & " as TaxAmt,S.Party_Code, SG.SiebelCode, sOrd.V_Date As OrderDate," & xIsNull("Part.Unit", "") & " As PartUnit, SpStk.SatPer, SpStk.SatAmt, S.SatAmt As SatAmt_H " & _
        " FROM (((((SP_Sale as S left JOIN SP_Stock as SPStk ON S.DocID = SPStk.Invoice_DocId) " & _
        "left JOIN Part ON SPStk.Part_No = Part.PART_NO and Part.Div_Code = left(SPStk.Docid,1)) " & _
        "LEFT JOIN (SubGroup as SG LEFT JOIN City ON SG.CityCode = City.CityCode) ON S.Party_Code = SG.SubCode) " & _
        "Left Join TaxForms TF on S.Form_Code=TF.Form_Code) " & _
        "Left Join Sp_Order sOrd On SPStk.Order_DocId = sOrd.OrderId) " & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable<>S.U_AE " & _
        "where S.DocId='" & Master!SearchCode & "' Order By SPStk.Srl_No"
End If
Select Case Index
     Case PScreen, PWindows
        If UCase(left(PubComp_Name, 3)) = "LMP" Then
            mRepName = IIf(OptPlain.Value = True, "SprSaleBillSiebel", "SprSaleBillSiebel")
        Else
            mRepName = IIf(mVatYn = 1, "SprSaleBillVat", "SprSaleBill")
        End If
        Call WindowsPrint(Index, GSQL)
        FrmPrn.Visible = False
    Case PDos
        If UCase(left(PubComp_Name, 3)) = "JMK" Then
            Call SpeedPrintJMK(GSQL, Optpre.Value)
            FrmPrn.Visible = False
        Else
            Call SpeedPrint(GSQL, Optpre.Value)
            FrmPrn.Visible = False
        End If
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "SprSaleBill", "SprSaleBill")
        Call PrinerSetUp
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        Txt(VDate).Tag = Txt(VDate).TEXT
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
Dim RstCompDet As ADODB.Recordset
Dim RstRep As ADODB.Recordset
Dim tmprs As ADODB.Recordset
Dim PartyLst$, mTitle$
Dim mJuriCity$, HlpLineNo$
Dim I As Integer, j As Integer

Set RstRep = GCn.Execute(mQry)
If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub

CreateFieldDefFile RstRep, PubRepoPath + "\" & mRepName & ".TTX", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
        

If mVatYn = 1 Then
    Set tmprs = GCn.Execute("Select LstNo From Subgroup Where SubCode='" & Txt(Party).Tag & "'")
    If tmprs.RecordCount > 0 Then
        PartyLst = XNull(tmprs!LstNo)
    End If
    If mVType = SalCrVType And PartyLst <> "" Then
        mTitle = "TAX INVOICE"
    ElseIf mVType = SalCrVType Then
        mTitle = "RETAIL INVOICE"
    Else
        mTitle = "RETAIL INVOICE [CASH]"
    End If
Else
    mTitle = IIf(mVType = SalCrVType, "INVOICE", "CASH MEMO")
End If


Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax, LstNoS, LstDateS from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("SrvTaxNo")
            rpt.FormulaFields(I).TEXT = "'" & PubServiceTaxNo & "'"
        Case UCase("Speciality")
            rpt.FormulaFields(I).TEXT = "'" & RstCompDet!S_SecSpeciality & "'"
        Case UCase("ActLst_No")
            rpt.FormulaFields(I).TEXT = "'" & RstCompDet!LstNoS & "'"
        Case UCase("ActLst_Date")
            rpt.FormulaFields(I).TEXT = "'" & RstCompDet!LstDateS & "'"
        Case UCase("LST")
            rpt.FormulaFields(I).TEXT = "'" & RstCompDet!S_SecLST & "'"
        Case UCase("LSTDate")
            rpt.FormulaFields(I).TEXT = "'" & RstCompDet!S_SecLST_Date & "'"
        Case UCase("CST")
            rpt.FormulaFields(I).TEXT = "'" & RstCompDet!S_SecCST & "'"
        Case UCase("CSTDate")
            rpt.FormulaFields(I).TEXT = "'" & RstCompDet!S_SecCST_Date & "'"
        Case UCase("Phone")
            rpt.FormulaFields(I).TEXT = "'" & RstCompDet!S_SecPhone & "'"
        Case UCase("Fax")
            rpt.FormulaFields(I).TEXT = "'" & RstCompDet!S_SecFax & "'"
        Case UCase("VouType")
            rpt.FormulaFields(I).TEXT = "'" & SalCashVType & "'"
        Case UCase("TOTCaption")
            rpt.FormulaFields(I).TEXT = "'" & pubTOTCaption & "'"
        Case UCase("GatePassOnSpareInvoice")
            rpt.FormulaFields(I).TEXT = "'" & PubGatePassOnSprInv & "'"
            
        Case UCase("HelpLine")
            Set tmprs = GCn.Execute("Select HelpLineNo from Syctrl")
            If tmprs.RecordCount > 0 Then
                HlpLineNo = IIf(IsNull(tmprs!HelpLineNo), "", Trim(tmprs!HelpLineNo))
                Set tmprs = Nothing
                If HlpLineNo <> "" Then
                    rpt.FormulaFields(I).TEXT = "'Help Line No : ' &  '" & HlpLineNo & "'"
                End If
            End If
                    
    End Select
Next
rpt.Database.SetDataSource RstRep
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
                    rpt.FormulaFields(I).TEXT = "'" & mTitle & "'"
            End Select
        Next
        rpt.PrintOut False
        If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
            GCn.Execute "update Sp_Sale set Printed_YN = 1  where Sp_Sale.docid='" & Master!SearchCode & "' "
        End If
    Case PScreen  'screen
         Call Report_View(rpt, mTitle, , True)
End Select
CmdPrint(PSetUp).Tag = ""
Set RstCompDet = Nothing
Set RstRep = Nothing
End Sub

Private Sub PrinerSetUp()
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.PrinterSetup (0)
CmdPrint(PSetUp).Tag = "1"
LblPrinter.CAPTION = rpt.PrinterName
End Sub

Private Sub SpeedPrint(mQry$, PrePrinted)
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
    Dim RstCompDet As ADODB.Recordset, RstRep As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, MRPTaxStr$
    
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select SprInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
        
    mHeader = 0   'Ideal 17
    mFooter = 19    'Line For Gate Pass =9 ,Line For NonTax Detail = 5
    mGatePass = 9
    mDetTax = 15
    mFooter = IIf(RstRep!Det_Tax = 1, mFooter, mDetTax)
    mFooter = mFooter + FooterCnt
    If (mVType = SalCashVType Or RstRep!GatePassOnSprInv = 1) And RstRep!Printed_YN = 0 And RstRep!CancelYN = 0 Then
        mFooter = mFooter + mGatePass
    End If
    'Sale Bill Header
    mDocStr = IIf(mVType = SalCrVType, "INVOICE", "CASH MEMO")
    mDupStr = IIf(RstRep!Printed_YN = 1, "(DUPLICATE)", "")
    If (mMRPTax + mMRPTaxSur + mMRPTOT) > 0 Then
        MRPTaxStr = "* Note:"
        If (mMRPTax + mMRPTaxSur) > 0 Then
            MRPTaxStr = MRPTaxStr & "Sales Tax Rs." & mMRPTax & ",Surcharge Rs. " & mMRPTaxSur
        End If
        If UCase(left(PubComp_Name, 3)) <> "JMK" Then
            If (mMRPTOT) > 0 Then
                MRPTaxStr = MRPTaxStr & " , " & PSTR(pubTOTCaption, 10, 0) & mMRPTOT
            End If
        End If
        MRPTaxStr = MRPTaxStr & " already added in MRP *'"
    End If
    If PrePrinted Then
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        mHeader = 8
    Else
        Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
        mHeader = mHeader + 1
        If XNull(RstCompDet!S_SecSpeciality) <> "" Then
            Print #1, PRN_TIT(RstCompDet!S_SecSpeciality, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
        mHeader = mHeader + 1
        If PubComp_Add2 <> "" Then
            Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If PubComp_City <> "" Then
            Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, PSTR(XNull(RstCompDet!S_SecLST) & IIf(XNull(RstCompDet!S_SecLST_Date) = "", "", " Dt. " & XNull(RstCompDet!S_SecLST_Date)), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!S_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!S_SecPhone)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstCompDet!S_SecCST) & IIf(XNull(RstCompDet!S_SecCST_Date) = "", "", " Dt. " & RstCompDet!S_SecCST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!S_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!S_SecFax)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
    End If
    If mVatYn = 1 Then
        If mVType = SalCashVType Then
            If RSOJPR = True Or left(PubComp_Name, 10) = "GANGANAGAR" Then
                Print #1, PRN_TIT("** VAT INVOICE " & mDupStr & " **", "A", PageWidth)
                mHeader = mHeader + 1
            Else
                If Txt(PType) = "Dealer" Then
                    Print #1, PRN_TIT("** TAX INVOICE" & mDupStr & " **", "A", PageWidth)
                    mHeader = mHeader + 1
                Else
                    If UCase(left(PubComp_Name, 5)) = "UJWAL" Then
                        Print #1, PRN_TIT("** TAX INVOICE " & mDupStr & " **", "A", PageWidth)
                    Else
                        Print #1, PRN_TIT("** RETAIL INVOICE " & mDupStr & " **", "A", PageWidth)
                    End If
                    mHeader = mHeader + 1
                End If

            End If
        Else
            Dim tmprs As ADODB.Recordset
            Set tmprs = GCn.Execute("Select Description from SubGroupType Left join Subgroup on Subgroup.Party_Type=SubgroupType.Party_Type where Subgroup.SubCode='" & RstRep!Party_code & "'")
            If tmprs.RecordCount > 0 Then
                If RSOJPR = True Or left(PubComp_Name, 10) = "GANGANAGAR" Then
                    Print #1, PRN_TIT("** VAT INVOICE" & mDupStr & " **", "A", PageWidth)
                    mHeader = mHeader + 1
                Else
                    If tmprs!Description = "Dealer" Then
                        Print #1, PRN_TIT("** TAX INVOICE" & mDupStr & " **", "A", PageWidth)
                        mHeader = mHeader + 1
                    Else
                        If UCase(left(PubComp_Name, 5)) = "UJWAL" Then
                            Print #1, PRN_TIT("** TAX INVOICE " & mDupStr & " **", "A", PageWidth)
                        Else
                            Print #1, PRN_TIT("** RETAIL INVOICE " & mDupStr & " **", "A", PageWidth)
                        End If
                        mHeader = mHeader + 1
                    End If
                End If
            End If
        End If
    Else
        Print #1, PRN_TIT("** " & IIf(mVType = SalCrVType, "CREDIT ", "") & mDocStr & mDupStr & " **", "A", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, mChr18 & "To," & mEmph
    mHeader = mHeader + 1
    Print #1, PSTR(RstRep!NamePrefix & " " & RstRep!Party_Name, 44) & Space(1) & PSTR(mDocStr & " NO.", Len(mDocStr) + 5) & " : " & PrinID(RstRep!DocID) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(IIf(mVType = SalCrVType, XNull(RstRep!Add1), XNull(RstRep!Address)), 44) & Space(1) & mEmph & PSTR(mDocStr & " DATE", Len(mDocStr) + 5) & " : " & PSTR(STR(RstRep!V_DATE), 14) & mEmph1
    mHeader = mHeader + 1
    Print #1, IIf(mVType = SalCrVType, XNull(RstRep!Add2), "") & Space(5) & mEmph & IIf(RstRep!CancelYN = 1, "** CANCELLED **", "") & mEmph1
    mHeader = mHeader + 1
    Print #1, IIf(mVType = SalCrVType, XNull(RstRep!Add3) & IIf(XNull(RstRep!Add3) <> "" And XNull(RstRep!CityName) <> "", ",", "") & XNull(RstRep!CityName) & " Phone : " & RstRep!Phone, "")
    mHeader = mHeader + 1
    Print #1, mDoub & "Remarks :" & XNull(RstRep!Remarks)
    mHeader = mHeader + 1
    If mVType = SalCrVType Then
        If mVatYn = 1 Then
            Print #1, mDoub & "TIN NO.:" & XNull(RstRep!LstNo) & Space(8) & "RC.No.:" & XNull(RstRep!RC_No) & mDoub1
            mHeader = mHeader + 1
        Else
            Print #1, mDoub & "CST NO.:" & XNull(RstRep!CstNo) & "   " & "LST NO.:" & XNull(RstRep!LstNo) & Space(5) & "RC.No.:" & XNull(RstRep!RC_No) & mDoub1
            mHeader = mHeader + 1
        End If
    End If
    Print #1, mDoub & "Customer Order Detail :" & XNull(RstRep!CustOrdDet) & mDoub1
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-") & mChr17 & mDoub
    mHeader = mHeader + 1
    If mVatYn = 1 Then
        If RstRep!Det_Tax = 1 Then
            Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("TAX %", 6, , AlignRight) & PSTR("TAX AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mDoub1 & mChr18
        Else
            Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mDoub1 & mChr18
        End If
        mHeader = mHeader + 1
    Else
        If RstRep!Det_Tax = 1 Then
            Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 35) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
            mHeader = mHeader + 1
            Print #1, Space(89) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mDoub1 & mChr18
            mHeader = mHeader + 1
        Else
            Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 28) & PSTR("DESCRIPTION", 40) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mDoub1 & mChr18
            mHeader = mHeader + 1
        End If
    End If
    Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
    mHeader = mHeader + 1
    mFix = PageLength - (mHeader + mFooter)
    Page = 1
    mLine = 1
    mSlNo = 1
    LAdd = VNull(RstRep!Gen_Sur_Amt) + VNull(RstRep!Trans_Amt) + VNull(RstRep!Tax_Amt) + VNull(RstRep!Tax_Sur_Amt) + VNull(RstRep!Packing) + VNull(RstRep!ReSalTax_Amt) + VNull(RstRep!Tot_Amt)
    SubTot = RstRep!SprAmt_MRP_TB + RstRep!OilAmt_MRP_TB + RstRep!SprAmt_MRP_TP + RstRep!OilAmt_MRP_TP _
        + RstRep!SprAmt_TB + RstRep!SprAmt_TP + RstRep!OilAmt_TB + RstRep!OilAmt_TP + Val(Txt(IWDiscTotTP).TEXT) + Val(Txt(IWDiscTotTB).TEXT)
    If RstRep.RecordCount > 0 Then
        I = 1
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
                If Not FirstPrint Then
                    Print #1, ""
                    FirstPrint = True
                End If
                If PrePrinted Then
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    mHeader = 8
                Else
                    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                    mHeader = mHeader + 1
                End If
                
                Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, mChr18 & "To," & mEmph
                mHeader = mHeader + 1
                Print #1, PSTR(RstRep!NamePrefix & " " & RstRep!Party_Name, 40) & Space(1) & PSTR(mDocStr & " NO.", Len(mDocStr) + 5) & " : " & PrinID(RstRep!DocID) & mEmph1
                mHeader = mHeader + 1
                Print #1, PSTR(IIf(mVType = SalCrVType, XNull(RstRep!Add1), XNull(RstRep!Address)), 40) & Space(1) & mEmph & PSTR(mDocStr & " DATE", Len(mDocStr) + 5) & " : " & PSTR(STR(RstRep!V_DATE), 14) & mEmph1
                mHeader = mHeader + 1
                Print #1, IIf(mVType = SalCrVType, XNull(RstRep!Add2), "") & Space(1) & mEmph & IIf(RstRep!CancelYN = 1, "** CANCELLED **", "") & mEmph1
                mHeader = mHeader + 1
                Print #1, IIf(mVType = SalCrVType, XNull(RstRep!Add3) & IIf(XNull(RstRep!Add3) <> "" And XNull(RstRep!CityName) <> "", ",", "") & XNull(RstRep!CityName), "")
                mHeader = mHeader + 1
                Print #1, mDoub & "Remarks :" & XNull(RstRep!Remarks)
                mHeader = mHeader + 1
                If mVType = SalCrVType Then
                    Print #1, mDoub & "CST NO.:" & XNull(RstRep!CstNo) & mDoub1
                    mHeader = mHeader + 1
                End If
                Print #1, Replace(Space(PageWidth), " ", "-") & mChr17 & mDoub
                mHeader = mHeader + 1
                If mVatYn = 1 Then
                    If RstRep!Det_Tax = 1 Then
                        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("TAX %", 6, , AlignRight) & PSTR("TAX AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mDoub1 & mChr18
                    Else
                        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mDoub1 & mChr18
                    End If
                    mHeader = mHeader + 1
                Else
                    If RstRep!Det_Tax = 1 Then
                        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 35) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
                        mHeader = mHeader + 1
                        Print #1, Space(89) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mDoub1 & mChr18
                        mHeader = mHeader + 1
                    Else
                        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 28) & PSTR("DESCRIPTION", 40) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mDoub1 & mChr18
                        mHeader = mHeader + 1
                    End If
                End If
                Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
                mHeader = mHeader + 1
                mFix = PageLength - (mHeader + mFooter)
                mLine = 1
            End If
            mRate = IIf(RstRep!MRP_YN = 1, RstRep!MRP_Rate2, RstRep!Rate2)
            If mVatYn = 1 Then
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstRep!Part_No, 22, , AlignLeft) & PSTR(RstRep!Part_Name, 30) & PSTR(RstRep!Qty_Iss, 12, 3)
                    If RstRep!Det_Tax = 1 Then
                        PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstRep!MRP_YN = 1, "M", "L") & _
                        PSTR(Format(RstRep!Disc_Per2, "0.00"), 8, 2) & " %" & PSTR(Format(RstRep!Disc_Amt2, "0.00"), 10, 2, AlignRight) & _
                        PSTR(Format(RstRep!TaxPer, "0.00"), 6, 2, AlignRight) & PSTR(Format(RstRep!TaxAmt, "0.00"), 10, 2, AlignRight) & _
                        PSTR(Format(RstRep!Net_Amt2, "0.00"), 12, 2, AlignRight)
                    Else
                        PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstRep!MRP_YN = 1, "M", "L") & _
                        PSTR(Format(RstRep!Disc_Per2, "0.00"), 8, 2) & " %" & PSTR(Format(RstRep!Disc_Amt2, "0.00"), 10, 2, AlignRight) & _
                        PSTR(Format(RstRep!Net_Amt2, "0.00"), 12, 2, AlignRight)
                    End If
            Else
                If RstRep!Det_Tax = 1 Then
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstRep!Part_No, 22, , AlignLeft) & PSTR(RstRep!Part_Name, 35) & PSTR(RstRep!Qty_Iss, 12, 3)
                    PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstRep!MRP_YN = 1, "M", "L") & _
                    PSTR(RstRep!Disc_Per2, 8, 2) & " %" & PSTR(RstRep!Disc_Amt2, 10, 2) & _
                    IIf(RstRep!Tax_YN = 0, PSTR(RstRep!Net_Amt2, 12, 2) & PSTR(0, 12, 2), PSTR(0, 12, 2) & PSTR(RstRep!Net_Amt2, 12, 2))
                Else
                    LAmtItem = RstRep!Net_Amt2 + RstRep!Disc_Amt2
                    LDAmt = LAmtItem + (LAmtItem * (LAdd / IIf(SubTot = 0, 1, SubTot)))
                    LAmtVal = LAmtVal + (LAmtItem * (LAdd / IIf(SubTot = 0, 1, SubTot)))
                    LdRate = LDAmt / IIf(RstRep!Qty_Iss = 0, 1, RstRep!Qty_Iss)
                     
                    If I = RstRep.RecordCount Then
                        If LAmtVal <> LAdd Then LDAmt = LDAmt + (LAdd - LAmtVal)
                        LdRate = LDAmt / IIf(RstRep!Qty_Iss = 0, 1, RstRep!Qty_Iss)
                    End If
                    mGrossAmt = mGrossAmt + (LDAmt - RstRep!Disc_Amt2)
                    I = I + 1
                    mAmount = Round(RstRep!Qty_Iss * RstRep!Rate, 2) - RstRep!Disc_Amt
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstRep!Part_No, 28, , AlignLeft) & PSTR(RstRep!Part_Name, 40) & PSTR(RstRep!Qty_Iss, 12, 3)
                    PrintStr = PrintStr & PSTR(LdRate, 11, 2) & " " & IIf(RstRep!MRP_YN = 1, "M", "L") & _
                    PSTR(RstRep!Disc_Per2, 8, 2) & " %" & PSTR(Format(RstRep!Disc_Amt2, "0.00"), 10, 2, AlignRight) & _
                    PSTR(Format(LDAmt - RstRep!Disc_Amt2, "0.00"), 12, 2, AlignRight)
                End If
            End If
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
    Print #1, mChr18 & "Customer's Signature"
    ' SALE FOOTER
    '22 space maintain between heading and :
    RstRep.MovePrevious
    If RstRep!Det_Tax = 1 Then
        Print #1, Replace(Space(21), " ", "-") & "TaxPaid" & Replace(Space(12), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")
        If mVatYn = 1 Then
            Print #1, PSTR("Item Disc.Amt", 16) & PSTR(Val(Txt(IWDiscTotTP)), 12, 2) & Space(8) & PSTR(Val(Txt(IWDiscTotTB)), 12, 2) _
            ; " | " & PSTR("V A T     ", 10, 0) & Space(6) & PSTR(RstRep!Tax_Amt, 12, 2) & mDoub
            
            Print #1, PSTR("MRP Items Amt", 16) & PSTR(RstRep!SprAmt_MRP_TP + RstRep!OilAmt_MRP_TP, 12, 2) & Space(8) & PSTR(RstRep!SprAmt_MRP_TB + RstRep!OilAmt_MRP_TB, 12, 2) & mDoub1 _
            ; " | " & IIf(VNull(RstRep!SatAmt_H) > 0, PSTR("S A T     ", 10, 0) & Space(6) & PSTR(RstRep!SatAmt_H, 12, 2), Space(10) & Space(6) & Space(12)) & mDoub
        
        Else
            Print #1, PSTR("Item Disc.Amt", 16) & PSTR(Val(Txt(IWDiscTotTP)), 12, 2) & Space(8) & PSTR(Val(Txt(IWDiscTotTB)), 12, 2) _
            ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstRep!Tax_Per, 5, 2) & "%" & PSTR(RstRep!Tax_Amt, 12, 2) & mDoub
            
            Print #1, PSTR("MRP Items Amt", 16) & PSTR(RstRep!SprAmt_MRP_TP + RstRep!OilAmt_MRP_TP, 12, 2) & Space(8) & PSTR(RstRep!SprAmt_MRP_TB + RstRep!OilAmt_MRP_TB, 12, 2) & mDoub1 _
            ; " | " & PSTR("Tax Surc. ", 10, 0) & PSTR(RstRep!Tax_Sur_Per, 5, 2) & "%" & PSTR(RstRep!Tax_Sur_Amt, 12, 2) & mDoub
            
        End If
        
        If UCase(left(PubComp_Name, 7)) = "JOHNSON" Then
            Print #1, PSTR("Spares Amount", 16) & PSTR(RstRep!SprAmt_TP, 12, 2) & Space(8) & PSTR(RstRep!SprAmt_TB, 12, 2) & mDoub1 _
            ; " | " & PSTR("Handling Charges", 16) & PSTR(RstRep!Packing, 12, 2) & mDoub
        Else
            Print #1, PSTR("Spares Amount", 16) & PSTR(RstRep!SprAmt_TP, 12, 2) & Space(8) & PSTR(RstRep!SprAmt_TB, 12, 2) & mDoub1 _
            ; " | " & PSTR("Misc. Charges", 16) & PSTR(RstRep!Packing, 12, 2) & mDoub
        End If

         
        Print #1, PSTR("Oil Amount ", 16) & PSTR(RstRep!OilAmt_TP, 12, 2) & Space(8) & PSTR(RstRep!OilAmt_TB, 12, 2) & mDoub1 _
        ; " | " & mEmph & PSTR("Sub Total[TP&TB]", 16) & PSTR(Val(Txt(STotB)), 12, 2) & mEmph1
        
        Print #1, PSTR("Discount ", 10, 0) & PSTR(RstRep!D_Per_TP, 5, 2) & "%" & PSTR(RstRep!D_Amt_TP, 12, 2) & PSTR(RstRep!D_Per_TB, 7, 2) & "%" & PSTR(RstRep!D_Amt_TB, 12, 2) _
        ; " | " & PSTR(pubTOTCaption, 10, 0) & PSTR(RstRep!TOT_Per, 5, 2) & "%" & PSTR(RstRep!Tot_Amt, 12, 2) & mEmph
        
        Print #1, PSTR("Sub Total [A]", 16) & PSTR(Val(Txt(STotATP)), 12, 2) & Space(8) & PSTR(Val(Txt(STotATB)), 12, 2) & mEmph1 _
        ; " | " & PSTR("ReSale Tax", 10, 0) & PSTR(RstRep!ReSalTax_Per, 5, 2) & "%" & PSTR(RstRep!ReSalTax_Amt, 12, 2)
        
        Print #1, PSTR("Gen Surch ", 10, 0) & PSTR(RstRep!Gen_Sur_Per, 5, 2) & "%" & PSTR(0, 12, 2) & PSTR(RstRep!Gen_Sur_Amt, 20, 2) _
        ; " | " & PSTR("Round Off", 16) & PSTR(RstRep!Rounded, 12, 2)
       
        Print #1, PSTR("Transportation", 16) & PSTR(0, 12, 2) & PSTR(RstRep!Trans_Amt, 20, 2) _
        ; " | " & mEmph & PSTR("Net Payble Rs.", 16) & PSTR(Val(Txt(NetAmt)), 12, 2) & mEmph1
    Else
        Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
        Print #1, Space(45) & PSTR("GOODS AMOUNT", 20) & " : " & PSTR(mGrossAmt, 12, 2) & mDoub1
        If RstRep!D_Amt_TP + RstRep!D_Amt_TB > 0 Then
            Print #1, Space(45) & PSTR("DISCOUNT", 20) & " : " & PSTR(RstRep!D_Amt_TP + RstRep!D_Amt_TB, 12, 2)
        Else
            Print #1, ""
        End If
        Print #1, Space(45) & PSTR("ROUNDED OFF", 20) & " : " & PSTR(Format(Val(Txt(NetAmt)) - (mGrossAmt - (RstRep!D_Amt_TP + RstRep!D_Amt_TB)), "0.00"), 12, 2, AlignRight) & mEmph
        Print #1, Space(45) & PSTR("Net Payble Rs.", 20) & " : " & PSTR(Val(Txt(NetAmt)), 12, 2) & mEmph1
    End If
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mDoub & ntow(Val(Txt(NetAmt)), "Rupees", "Paise") & mDoub1
    Print #1, Replace(Space(PageWidth), " ", "-")
    If mVatYn = 1 Then
        Print #1, ""
    Else
        Print #1, mChr17 & MRPTaxStr & mChr18 & Space(PageWidth - ((Len(MRPTaxStr) + 6) / 1.7)) & mChr17 & "E & OE" & mChr18
    End If
    Print #1, PSTR(RstRep!Printing_Desc, 25) & Space(PageWidth - (25 + Len("For " & PubComp_Name))) & "For " & mEmph & PubComp_Name & mEmph1
    Print #1, "Mode of Dispatch : " & PSTR(XNull(RstRep!Mode_Dispatch), 25) & mDoub
    Print #1, "Terms & Condition:" & mDoub1 & Replace(Space(PageWidth - 18), " ", "-") & mChr17
    Footer = Footer & vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    Print #1, Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
    ' Gate Pass Footer() SalCashVType
    If (mVType = SalCashVType Or RstRep!GatePassOnSprInv = 1) And VNull(RstRep!Printed_YN) = 0 And VNull(RstRep!CancelYN) = 0 Then
        Print #1, Replace(Space(PageWidth), " ", "-")
        Print #1, PRN_TIT("* " & mDocStr & " GATE PASS " & mDupStr & " *", "A", 80) & mEmph
        Print #1, "GATE PASS No. & DATE : " & XNull(RstRep!gp_no) & "  " & IIf(IsNull(RstRep!GP_Date), "", ConvertDate(RstRep!GP_Date)) & mEmph1
        Print #1, PSTR(RstRep!NamePrefix & " " & RstRep!Party_Name, 40) & Space(1) & PSTR(mDocStr & " NO.", Len(mDocStr) + 5) & " : " & PrinID(RstRep!DocID)
        Print #1, PSTR(XNull(RstRep!CityName), 40) & Space(1) & PSTR(mDocStr & " DATE", Len(mDocStr) + 5) & " : " & CDate(RstRep!V_DATE)
        Print #1, "Goods of " & mEmph & "Rs." & LTrim(PSTR(Val(Txt(NetAmt)), 9, 2)) & mEmph1 & " as per Document No. are being permitted for out."
        Print #1, "Mode of dispatch :" & XNull(RstRep!Mode_Dispatch)
        Print #1, ""
        Print #1, "Customer's Signature" & Space(50 - Len(PubComp_Name)) & "for " & mEmph & PubComp_Name & mEmph1
'        Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
        Print #1, mChr17 & "* a dataman software *" & mChr18
    End If
    Print #1, mEject
    Close #1
    FirstPrint = IIf(FirstPrint, FirstPrint, True)
    Set RstRep = Nothing
    Set RstCompDet = Nothing
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        'Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''        Print #1, "Type C:\RepPrint.Txt > Prn"
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
    
    
    
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
        GCn.Execute "update Sp_Sale set Printed_YN = 1  where Sp_Sale.docid='" & Master!SearchCode & "' "
    End If
    Exit Sub
ELoop:
    Set RstRep = Nothing
    Set RstCompDet = Nothing
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub
Private Sub FillOrderDetail()
Dim RsOrder As ADODB.Recordset, I As Double, SordNo As String, SordDocId$
GSQL = "Select Srl_No,Part_No,Tax_YN,MRP_YN,(Qty-Sup_Qty) as Qty,Rate,Disc_Per,Disc_Amt,Amount from SP_Order1 S1 Where OrderID='" & FGrid.TextMatrix(FGrid.Row, Col_SONoCode) & "'"
    Set RsOrder = New ADODB.Recordset
    RsOrder.CursorLocation = adUseClient
    RsOrder.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    If RsOrder.RecordCount > 0 Then
        SordNo = FGrid.TextMatrix(FGrid.Row, Col_SONo)
        SordDocId = FGrid.TextMatrix(FGrid.Row, Col_SONoCode)
        If FGrid.Rows < (FGrid.Row + RsOrder.RecordCount) Then
            FGrid.Rows = FGrid.Rows + RsOrder.RecordCount
        End If
        For I = FGrid.Row To (FGrid.Row + RsOrder.RecordCount) - 1
            FGrid.TextMatrix(I, Col_SONo) = SordNo
            FGrid.TextMatrix(I, Col_SONoCode) = SordDocId
            FGrid.TextMatrix(I, Col_SrNo) = RsOrder!Srl_No
            FGrid.TextMatrix(I, Col_PNo) = RsOrder!Part_No
            FGrid.TextMatrix(I, Col_Taxable) = IIf(RsOrder!Tax_YN = 1, "Yes", "No")
            FGrid.TextMatrix(I, Col_MRP) = IIf(RsOrder!MRP_YN = 1, "Yes", "No")
            FGrid.TextMatrix(I, Col_Qty) = Format(VNull(RsOrder!Qty), "0.00")
            FGrid.TextMatrix(I, Col_Rate) = Format(VNull(RsOrder!Rate), "0.0000")
            FGrid.TextMatrix(I, Col_MRPRate) = Format(IIf(RsOrder!MRP_YN = 1, VNull(RsOrder!Rate), 0), "0.0000")
            FGrid.TextMatrix(I, Col_DiscPer) = Format(RsOrder!Disc_Per, "0.0000")
            FGrid.TextMatrix(I, Col_DiscAmt) = RsOrder!Disc_Amt
            FGrid.TextMatrix(I, Col_Amt) = Format(RsOrder!Amount, "0.00")
            FGrid.TextMatrix(I, Col_ItemVal) = Format(RsOrder!Amount, "0.00")
            RsOrder.MoveNext
            If RsGodown.RecordCount > 0 And Trim(FGrid.TextMatrix(I, Col_Godown)) = "" Then
                RsGodown.MoveFirst
                RsGodown.FIND "Code ='" & PubSprCounterGodown & "'"
                FGrid.TextMatrix(I, Col_GodownCode) = RsGodown!Code
                FGrid.TextMatrix(I, Col_Godown) = RsGodown!Name
            End If
        Next
        Amt_Cal
    End If
End Sub
Private Sub SpeedPrintJMK(mQry$, PrePrinted)
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
    Dim mVatActualInvoiceNo As Long
    Dim mVatBookNo As Long
    Dim mVatInvoiceNo As Long
    
    Dim I As Integer, j As Integer
    Dim PrintStr$
    Dim RstCompDet As ADODB.Recordset, RstRep As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, MRPTaxStr$
    
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select SprInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
    PageLength = PubPageLength
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
        
    mHeader = 0   'Ideal 17
    mFooter = 19    'Line For Gate Pass =9 ,Line For NonTax Detail = 5
    mGatePass = 9
    mDetTax = 15
    mFooter = IIf(RstRep!Det_Tax = 1, mFooter, mDetTax)
    mFooter = mFooter + FooterCnt
    If (mVType = SalCashVType Or RstRep!GatePassOnSprInv = 1) And RstRep!Printed_YN = 0 And RstRep!CancelYN = 0 Then
        mFooter = mFooter + mGatePass
    End If
    'Sale Bill Header
    mDocStr = IIf(mVType = SalCrVType, "INVOICE", "CASH MEMO")
    If UCase(pubUName) <> "SA" Then
        mDupStr = IIf(RstRep!Printed_YN = 1, "(DUPLICATE)", "")
    End If
    If (mMRPTax + mMRPTaxSur + mMRPTOT) > 0 Then
        MRPTaxStr = "* Note:"
        If (mMRPTax + mMRPTaxSur) > 0 Then
            MRPTaxStr = MRPTaxStr & "Sales Tax Rs." & mMRPTax & ",Surcharge Rs. " & mMRPTaxSur
        End If
        If UCase(left(PubComp_Name, 3)) <> "JMK" Then
            If (mMRPTOT) > 0 Then
                MRPTaxStr = MRPTaxStr & " , " & PSTR(pubTOTCaption, 10, 0) & mMRPTOT
            End If
        End If
        MRPTaxStr = MRPTaxStr & " already added in MRP *'"
    End If
    If PrePrinted Then
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        mHeader = 8
    Else
        Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
        mHeader = mHeader + 1
        If XNull(RstCompDet!S_SecSpeciality) <> "" Then
            Print #1, PRN_TIT(RstCompDet!S_SecSpeciality, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
        mHeader = mHeader + 1
        If PubComp_Add2 <> "" Then
            Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        If PubComp_City <> "" Then
            Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
            mHeader = mHeader + 1
        End If
        Print #1, PSTR(XNull(RstCompDet!S_SecLST) & IIf(XNull(RstCompDet!S_SecLST_Date) = "", "", " Dt. " & XNull(RstCompDet!S_SecLST_Date)), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!S_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!S_SecPhone)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstCompDet!S_SecCST) & IIf(XNull(RstCompDet!S_SecCST_Date) = "", "", " Dt. " & RstCompDet!S_SecCST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!S_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!S_SecFax)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
    End If
    If mVatYn = 1 Then
        If mVType = SalCashVType Then
            If Txt(PType) = "Dealer" Then
                Print #1, PRN_TIT("** TAX INVOICE" & mDupStr & " **", "A", PageWidth)
                mHeader = mHeader + 1
            Else
                Print #1, PRN_TIT("** RETAIL INVOICE " & mDupStr & " **", "A", PageWidth)
                mHeader = mHeader + 1
            End If
        Else
            Dim tmprs As ADODB.Recordset
            Set tmprs = GCn.Execute("Select Description from SubGroupType Left join Subgroup on Subgroup.Party_Type=SubgroupType.Party_Type where Subgroup.SubCode='" & RstRep!Party_code & "'")
            If tmprs.RecordCount > 0 Then
                If tmprs!Description = "Dealer" Then
                    Print #1, PRN_TIT("** TAX INVOICE" & mDupStr & " **", "A", PageWidth)
                    mHeader = mHeader + 1
                Else
                    Print #1, PRN_TIT("** RETAIL INVOICE " & mDupStr & " **", "A", PageWidth)
                    mHeader = mHeader + 1
                End If
            End If
        End If
    Else
        Print #1, PRN_TIT("** " & IIf(mVType = SalCrVType, "CREDIT ", "") & mDocStr & mDupStr & " **", "A", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, mChr18 & "To," & mEmph
    mHeader = mHeader + 1
    If mVatYn = 1 And StrCmp(left(PubComp_Name, 3), "Jmk") Then
        mVatActualInvoiceNo = Val(Right(Replace(Right(RstRep!DocID, 8), " ", ""), Len(Replace(Right(RstRep!DocID, 8), " ", "")) - 1))
        mVatBookNo = Fix((mVatActualInvoiceNo - 1) / 50) + 1
        mVatInvoiceNo = ((mVatActualInvoiceNo - 1) Mod 50) + 1
        Print #1, PSTR(RstRep!NamePrefix & " " & RstRep!Party_Name, 44) & Space(1) & "Book No : " & mVatBookNo & "  C.M. No. : " & mVatInvoiceNo & mEmph1
    Else
        
        Print #1, PSTR(RstRep!NamePrefix & " " & RstRep!Party_Name, 44) & Space(1) & PSTR(mDocStr & " NO.", Len(mDocStr) + 5) & " : " & PrinID(RstRep!DocID) & mEmph1
    End If
    mHeader = mHeader + 1
    Print #1, PSTR(IIf(mVType = SalCrVType, XNull(RstRep!Add1), XNull(RstRep!Address)), 44) & Space(1) & mEmph & PSTR(mDocStr & " DATE", Len(mDocStr) + 5) & " : " & PSTR(STR(RstRep!V_DATE), 14) & mEmph1
    mHeader = mHeader + 1
    Print #1, IIf(mVType = SalCrVType, XNull(RstRep!Add2), "") & Space(5) & mEmph & IIf(RstRep!CancelYN = 1, "** CANCELLED **", "") & mEmph1
    mHeader = mHeader + 1
    Print #1, IIf(mVType = SalCrVType, XNull(RstRep!Add3) & IIf(XNull(RstRep!Add3) <> "" And XNull(RstRep!CityName) <> "", ",", "") & XNull(RstRep!CityName) & " Phone : " & RstRep!Phone, "")
    mHeader = mHeader + 1
    Print #1, mDoub & "Remarks :" & XNull(RstRep!Remarks)
    mHeader = mHeader + 1
    If mVType = SalCrVType Then
        If mVatYn = 1 Then
            Print #1, mDoub & "TIN NO.:" & XNull(RstRep!CstNo) & "   " & "LST NO.:" & XNull(RstRep!LstNo) & Space(5) & "RC.No.:" & XNull(RstRep!RC_No) & mDoub1
            mHeader = mHeader + 1
        Else
            Print #1, mDoub & "CST NO.:" & XNull(RstRep!CstNo) & "   " & "LST NO.:" & XNull(RstRep!LstNo) & Space(5) & "RC.No.:" & XNull(RstRep!RC_No) & mDoub1
            mHeader = mHeader + 1
        End If
    End If
    Print #1, mDoub & "Custmer Order Detail :" & XNull(RstRep!CustOrdDet) & mDoub1
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-") & mChr17 & mDoub
    mHeader = mHeader + 1
    If mVatYn = 1 Then
        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("TAX %", 6, , AlignRight) & PSTR("TAX AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mDoub1 & mChr18
        mHeader = mHeader + 1
    Else
        If RstRep!Det_Tax = 1 Then
            Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP RATE", 11, , AlignRight) & Space(2) & PSTR("RATE", 21, , AlignRight) & "<---------AMOUNT--------- >"
            mHeader = mHeader + 1
            Print #1, Space(7) & Space(22) & Space(30) & Space(12) & Space(11) & Space(21) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mDoub1 & mChr18
            mHeader = mHeader + 1
        Else
            Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP RATE", 11, , AlignRight) & Space(2) & PSTR("RATE", 21, , AlignRight) & PSTR("Amount", 27, , AlignRight)
            mHeader = mHeader + 1
        End If
    End If
    Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
    mHeader = mHeader + 1
    mFix = PageLength - (mHeader + mFooter)
    Page = 1
    mLine = 1
    mSlNo = 1
    LAdd = VNull(RstRep!Gen_Sur_Amt) + VNull(RstRep!Trans_Amt) + VNull(RstRep!Tax_Amt) + VNull(RstRep!Tax_Sur_Amt) + VNull(RstRep!Packing) + VNull(RstRep!ReSalTax_Amt) + VNull(RstRep!Tot_Amt)
    SubTot = RstRep!SprAmt_MRP_TB + RstRep!OilAmt_MRP_TB + RstRep!SprAmt_MRP_TP + RstRep!OilAmt_MRP_TP _
        + RstRep!SprAmt_TB + RstRep!SprAmt_TP + RstRep!OilAmt_TB + RstRep!OilAmt_TP + Val(Txt(IWDiscTotTP).TEXT) + Val(Txt(IWDiscTotTB).TEXT)
    If RstRep.RecordCount > 0 Then
        I = 1
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
                If Not FirstPrint Then
                    Print #1, ""
                    FirstPrint = True
                End If
                If PrePrinted Then
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    Print #1, "": Print #1, "": Print #1, "": Print #1, ""
                    mHeader = 8
                Else
                    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                    mHeader = mHeader + 1
                End If
                
                Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, mChr18 & "To," & mEmph
                mHeader = mHeader + 1
                Print #1, PSTR(RstRep!NamePrefix & " " & RstRep!Party_Name, 40) & Space(1) & PSTR(mDocStr & " NO.", Len(mDocStr) + 5) & " : " & PrinID(RstRep!DocID) & mEmph1
                mHeader = mHeader + 1
                Print #1, PSTR(IIf(mVType = SalCrVType, XNull(RstRep!Add1), XNull(RstRep!Address)), 40) & Space(1) & mEmph & PSTR(mDocStr & " DATE", Len(mDocStr) + 5) & " : " & PSTR(STR(RstRep!V_DATE), 14) & mEmph1
                mHeader = mHeader + 1
                Print #1, IIf(mVType = SalCrVType, XNull(RstRep!Add2), "") & Space(1) & mEmph & IIf(RstRep!CancelYN = 1, "** CANCELLED **", "") & mEmph1
                mHeader = mHeader + 1
                Print #1, IIf(mVType = SalCrVType, XNull(RstRep!Add3) & IIf(XNull(RstRep!Add3) <> "" And XNull(RstRep!CityName) <> "", ",", "") & XNull(RstRep!CityName), "")
                mHeader = mHeader + 1
                Print #1, mDoub & "Remarks :" & XNull(RstRep!Remarks)
                mHeader = mHeader + 1
                If mVType = SalCrVType Then
                    Print #1, mDoub & "CST NO.:" & XNull(RstRep!CstNo) & mDoub1
                    mHeader = mHeader + 1
                End If
                Print #1, Replace(Space(PageWidth), " ", "-") & mChr17 & mDoub
                mHeader = mHeader + 1
                If mVatYn = 1 Then
                    Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("TAX %", 6, , AlignRight) & PSTR("TAX AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mDoub1 & mChr18
                    mHeader = mHeader + 1
                Else
                    If RstRep!Det_Tax = 1 Then
                        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP RATE", 11, , AlignRight) & Space(2) & PSTR("RATE", 18, , AlignRight) & Space(3) & "<---------AMOUNT--------- >"
                        mHeader = mHeader + 1
                        Print #1, Space(7) & Space(22) & Space(30) & Space(12) & Space(11) & Space(2) & Space(18) & Space(3) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mDoub1 & mChr18
                        mHeader = mHeader + 1
                    Else
                        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("MRP RATE", 11, , AlignRight) & Space(2) & PSTR("RATE", 18, , AlignRight) & Space(3) & PSTR("Amount", 27, , AlignRight)
                        mHeader = mHeader + 1
                    End If
                End If
                Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
                mHeader = mHeader + 1
                mFix = PageLength - (mHeader + mFooter)
                mLine = 1
            End If
            mRate = IIf(RstRep!MRP_YN = 1, RstRep!MRP_Rate2, RstRep!Rate2)
            If mVatYn = 1 Then
            
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstRep!Part_No, 22, , AlignLeft) & PSTR(RstRep!Part_Name, 30) & PSTR(RstRep!Qty_Iss, 12, 3)
                    PrintStr = PrintStr & PSTR(mRate, 11, 2) & "  " & _
                    PSTR(Format(RstRep!Disc_Per2, "0.00"), 8, 2, AlignRight) & " %" & PSTR(Format(RstRep!Disc_Amt2, "0.00"), 10, 2, AlignRight) & _
                    PSTR(Format(RstRep!TaxPer, "0.00"), 6, 2, AlignRight) & PSTR(Format(RstRep!TaxAmt, "0.00"), 10, 2, AlignRight) & _
                    PSTR(Format(RstRep!Net_Amt2, "0.00"), 12, 2, AlignRight)
            
            
'                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstRep!Part_No, 22, , AlignLeft) & PSTR(RstRep!Part_Name, 30) & PSTR(RstRep!Qty_Iss, 12, 3)
'                    PrintStr = PrintStr & PSTR(IIf(RstRep!MRP_YN = 1, STR(mRate), 0), 11, 2, AlignRight) & " " & Space(2) & _
'                    PSTR(Format(STR(mRate), "0.00"), 18, 2, AlignRight) & Space(3) & _
'                    PSTR(Format(RstRep!Net_Amt2, "0.00"), 12, 2, AlignRight)
            Else
                If RstRep!Det_Tax = 1 Then
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstRep!Part_No, 22, , AlignLeft) & PSTR(RstRep!Part_Name, 30) & PSTR(RstRep!Qty_Iss, 12, 3)
                    PrintStr = PrintStr & PSTR(IIf(RstRep!MRP_YN = 1, STR(mRate), 0), 11, 2, AlignRight) & " " & Space(2) & _
                    PSTR(Format(STR(mRate), "0.00"), 18, 2, AlignRight) & Space(3) & _
                    PSTR(IIf(RstRep!Tax_YN = 0, Format(RstRep!Net_Amt2, "0.00"), "--"), 12, 2, AlignRight) & PSTR(IIf(RstRep!Tax_YN = 1, Format(RstRep!Net_Amt2, "0.00"), "--"), 12, 2, AlignRight)
                Else
                    PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstRep!Part_No, 22, , AlignLeft) & PSTR(RstRep!Part_Name, 30) & PSTR(RstRep!Qty_Iss, 12, 3)
                    PrintStr = PrintStr & PSTR(IIf(RstRep!MRP_YN = 1, STR(mRate), 0), 11, 2, AlignRight) & " " & Space(2) & _
                    PSTR(Format(STR(mRate), "0.00"), 18, 2, AlignRight) & Space(3) & _
                    PSTR(IIf(RstRep!Tax_YN = 0, Format(RstRep!Net_Amt2, "0.00"), "--"), 12, 2, AlignRight) & PSTR(IIf(RstRep!Tax_YN = 1, Format(RstRep!Net_Amt2, "0.00"), "--"), 12, 2, AlignRight)
                End If
            End If
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
    Print #1, mChr18 & "Customer's Signature"
    ' SALE FOOTER
    '22 space maintain between heading and :
    
    Dim mVat12 As Double
    Dim mVat4 As Double
    
    
    RstRep.MovePrevious
    mVat12 = GCn.Execute("Select " & vIsNull("Sum(TaxAmt)", "0") & " From Sp_Stock Where Invoice_DocId='" & RstRep!DocID & "' and TaxPer>=12.5").Fields(0).Value
    mVat4 = GCn.Execute("Select " & vIsNull("Sum(TaxAmt)", "0") & " From Sp_Stock Where Invoice_DocId='" & RstRep!DocID & "' and TaxPer<12").Fields(0).Value
    
    
    If RstRep!Det_Tax = 1 Then
        If mVatYn = 1 Then
            Print #1, Replace(Space(40), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")
            
'            Print #1, PSTR("Item Disc.Amt", 16) & Space(20) & PSTR(Val(Txt(IWDiscTotTB)), 12, 2) _
'            ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstRep!Tax_Per, 5, 2) & "%" & PSTR(RstRep!Tax_Amt, 12, 2) & mDoub
'
'            Print #1, PSTR("MRP Items Amt", 16) & Space(20) & PSTR(RstRep!SprAmt_MRP_TB + RstRep!OilAmt_MRP_TB, 12, 2) & mDoub1 _
'            ; " | " & PSTR("Tax Surc. ", 10, 0) & PSTR(RstRep!Tax_Sur_Per, 5, 2) & "%" & PSTR(RstRep!Tax_Sur_Amt, 12, 2) & mDoub
                
            Print #1, PSTR("Item Disc.Amt", 16) & Space(20) & PSTR(Val(Txt(IWDiscTotTB)), 12, 2) _
            ; " | " & PSTR("V A T  12.5%", 15, 0) & " " & PSTR(mVat12, 12, 2) & mDoub
            
            Print #1, PSTR("MRP Items Amt", 16) & Space(20) & PSTR(RstRep!SprAmt_MRP_TB + RstRep!OilAmt_MRP_TB, 12, 2) & mDoub1 _
            ; " | " & PSTR("V A T  4% ", 15, 0) & " " & PSTR(mVat4, 12, 2) & mDoub
 
            Print #1, PSTR("Oil Amount ", 16) & Space(20) & PSTR(RstRep!OilAmt_TB, 12, 2) & mDoub1 _
            ; " | " & mEmph & PSTR("Sub Total[TP&TB]", 16) & PSTR(Val(Txt(STotB)), 12, 2) & mEmph1
            
            Print #1, PSTR("Discount ", 10, 0) & Space(18) & PSTR(RstRep!D_Per_TB, 7, 2) & "%" & PSTR(RstRep!D_Amt_TB, 12, 2) _
            ; " | " & PSTR(pubTOTCaption, 10, 0) & PSTR(RstRep!TOT_Per, 5, 2) & "%" & PSTR(RstRep!Tot_Amt, 12, 2) & mEmph
            
            If Not mSatYn Then
                Print #1, PSTR("Sub Total [A]", 16) & Space(20) & PSTR(Val(Txt(STotATB)), 12, 2) & mEmph1 _
                ; " | " & PSTR("ReSale Tax", 10, 0) & PSTR(RstRep!ReSalTax_Per, 5, 2) & "%" & PSTR(RstRep!ReSalTax_Amt, 12, 2)
            Else
                Print #1, PSTR("Sub Total [A]", 16) & Space(20) & PSTR(Val(Txt(STotATB)), 12, 2) & mEmph1 _
                ; " | " & PSTR("S A T", 16, 0) & PSTR(Val(Txt(SatAmt)), 12, 2)
            End If
            
            Print #1, PSTR("Gen Surch ", 10, 0) & Space(18) & PSTR(RstRep!Gen_Sur_Amt, 20, 2) _
            ; " | " & PSTR("Round Off", 16) & PSTR(RstRep!Rounded, 12, 2)
           
            Print #1, PSTR("Transportation", 16) & Space(12) & PSTR(RstRep!Trans_Amt, 20, 2) _
            ; " | " & mEmph & PSTR("Net Payble Rs.", 16) & PSTR(Val(Txt(NetAmt)), 12, 2) & mEmph1
        Else
            Print #1, Replace(Space(21), " ", "-") & "TaxPaid" & Replace(Space(12), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")

            Print #1, PSTR("Item Disc.Amt", 16) & PSTR(Val(Txt(IWDiscTotTP)), 12, 2) & Space(8) & PSTR(Val(Txt(IWDiscTotTB)), 12, 2) _
            ; " | " & PSTR("V A T     ", 10, 0) & Space(6) & PSTR(RstRep!Tax_Amt, 12, 2) & mDoub
            
            Print #1, PSTR("MRP Items Amt", 16) & PSTR(RstRep!SprAmt_MRP_TP + RstRep!OilAmt_MRP_TP, 12, 2) & Space(8) & PSTR(RstRep!SprAmt_MRP_TB + RstRep!OilAmt_MRP_TB, 12, 2) & mDoub1 _
            ; " | " & Space(10) & Space(6) & Space(12) & mDoub
            
            Print #1, PSTR("Oil Amount ", 16) & PSTR(RstRep!OilAmt_TP, 12, 2) & Space(8) & PSTR(RstRep!OilAmt_TB, 12, 2) & mDoub1 _
            ; " | " & mEmph & PSTR("Sub Total[TP&TB]", 16) & PSTR(Val(Txt(STotB)), 12, 2) & mEmph1
            
            Print #1, PSTR("Discount ", 10, 0) & PSTR(RstRep!D_Per_TP, 5, 2) & "%" & PSTR(RstRep!D_Amt_TP, 12, 2) & PSTR(RstRep!D_Per_TB, 7, 2) & "%" & PSTR(RstRep!D_Amt_TB, 12, 2) _
            ; " | " & PSTR(pubTOTCaption, 10, 0) & PSTR(RstRep!TOT_Per, 5, 2) & "%" & PSTR(RstRep!Tot_Amt, 12, 2) & mEmph
            
            Print #1, PSTR("Sub Total [A]", 16) & PSTR(Val(Txt(STotATP)), 12, 2) & Space(8) & PSTR(Val(Txt(STotATB)), 12, 2) & mEmph1 _
            ; " | " & PSTR("ReSale Tax", 10, 0) & PSTR(RstRep!ReSalTax_Per, 5, 2) & "%" & PSTR(RstRep!ReSalTax_Amt, 12, 2)
            
            Print #1, PSTR("Gen Surch ", 10, 0) & PSTR(RstRep!Gen_Sur_Per, 5, 2) & "%" & PSTR(0, 12, 2) & PSTR(RstRep!Gen_Sur_Amt, 20, 2) _
            ; " | " & PSTR("Round Off", 16) & PSTR(RstRep!Rounded, 12, 2)
           
            Print #1, PSTR("Transportation", 16) & PSTR(0, 12, 2) & PSTR(RstRep!Trans_Amt, 20, 2) _
            ; " | " & mEmph & PSTR("Net Payble Rs.", 16) & PSTR(Val(Txt(NetAmt)), 12, 2) & mEmph1
        End If
    Else
        Print #1, Replace(Space(PageWidth), " ", "-") & mDoub
        Print #1, Space(45) & PSTR("GOODS AMOUNT", 20) & " : " & PSTR(mGrossAmt, 12, 2) & mDoub1
        If RstRep!D_Amt_TP + RstRep!D_Amt_TB > 0 Then
            Print #1, Space(45) & PSTR("DISCOUNT", 20) & " : " & PSTR(RstRep!D_Amt_TP + RstRep!D_Amt_TB, 12, 2)
        Else
            Print #1, ""
        End If
        Print #1, Space(45) & PSTR("ROUNDED OFF", 20) & " : " & PSTR(Format(Val(Txt(NetAmt)) - (mGrossAmt - (RstRep!D_Amt_TP + RstRep!D_Amt_TB)), "0.00"), 12, 2, AlignRight) & mEmph
        Print #1, Space(45) & PSTR("Net Payble Rs.", 20) & " : " & PSTR(Val(Txt(NetAmt)), 12, 2) & mEmph1
    End If
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, mDoub & ntow(Val(Txt(NetAmt)), "Rupees", "Paise") & mDoub1
    Print #1, Replace(Space(PageWidth), " ", "-")
    If mVatYn = 1 Then
        Print #1, ""
    Else
        Print #1, mChr17 & MRPTaxStr & mChr18 & Space(PageWidth - ((Len(MRPTaxStr) + 6) / 1.7)) & mChr17 & "E & OE" & mChr18
    End If
    Print #1, PSTR(RstRep!Printing_Desc, 25) & Space(PageWidth - (25 + Len("For " & PubComp_Name))) & "For " & mEmph & PubComp_Name & mEmph1
    Print #1, "Mode of Dispatch : " & PSTR(XNull(RstRep!Mode_Dispatch), 25) & mDoub
    Print #1, "Terms & Condition:" & mDoub1 & Replace(Space(PageWidth - 18), " ", "-") & mChr17
    Footer = Footer & vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
    Next
    Print #1, Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
    ' Gate Pass Footer() SalCashVType
    If (mVType = SalCashVType Or RstRep!GatePassOnSprInv = 1) And RstRep!Printed_YN = 0 And RstRep!CancelYN = 0 Then
        Print #1, Replace(Space(PageWidth), " ", "-")
        Print #1, PRN_TIT("* " & mDocStr & " GATE PASS " & mDupStr & " *", "A", 80) & mEmph
        Print #1, "GATE PASS No. & DATE : " & XNull(RstRep!gp_no) & "  " & IIf(IsNull(RstRep!GP_Date), "", ConvertDate(RstRep!GP_Date)) & mEmph1
        Print #1, PSTR(RstRep!NamePrefix & " " & RstRep!Party_Name, 40) & Space(1) & PSTR(mDocStr & " NO.", Len(mDocStr) + 5) & " : " & PrinID(RstRep!DocID)
        Print #1, PSTR(XNull(RstRep!CityName), 40) & Space(1) & PSTR(mDocStr & " DATE", Len(mDocStr) + 5) & " : " & CDate(RstRep!V_DATE)
        Print #1, "Goods of " & mEmph & "Rs." & LTrim(PSTR(Val(Txt(NetAmt)), 9, 2)) & mEmph1 & " as per Document No. are being permitted for out."
        Print #1, "Mode of dispatch :" & XNull(RstRep!Mode_Dispatch)
        Print #1, ""
        Print #1, "Customer's Signature" & Space(50 - Len(PubComp_Name)) & "for " & mEmph & PubComp_Name & mEmph1
'        Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
        Print #1, mChr17 & "* a dataman software *" & mChr18
    End If
    Print #1, mEject
    Close #1
    FirstPrint = IIf(FirstPrint, FirstPrint, True)
    Set RstRep = Nothing
    Set RstCompDet = Nothing
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        'Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''        Print #1, "Type C:\RepPrint.Txt > Prn"
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
    GCn.Execute "update Sp_Sale set Printed_YN = 1  where Sp_Sale.docid='" & Master!SearchCode & "' "
    Exit Sub
ELoop:
    Set RstRep = Nothing
    Set RstCompDet = Nothing
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub
Private Function FIFOStkIss(Part_No As String)
Dim TBRst As ADODB.Recordset, TPRst As ADODB.Recordset, TmpRst As ADODB.Recordset
Dim MRPTBRst As ADODB.Recordset, MRPTPRst As ADODB.Recordset
Dim TBCurrStk, TPCurrStk As Double, I As Double
Dim MRPTBCurrStk, MRPTPCurrStk As Double
Dim TBStkDate$, TBStkPurDocId$, TBStkPurDate$, TPStkDate$, TPStkPurDocId$, TPStkPurDate$
Dim MRPTBStkDate$, MRPTBStkPurDocId$, MRPTBStkPurDate$, MRPTPStkDate$, MRPTPStkPurDocId$, MRPTPStkPurDate$

Set TBRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec,SP_Purch.Party_Doc_No,SP_Purch.Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId where Qty_Rec > 0 and Tax_YN=1 and MRP_YN=0 and Part_No='" & Part_No & "'")
Set TPRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec,SP_Purch.Party_Doc_No,SP_Purch.Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId  where Qty_Rec > 0 and Tax_YN=0 and MRP_YN=0 and Part_No='" & Part_No & "'")
Set MRPTBRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec,SP_Purch.Party_Doc_No,SP_Purch.Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId  where Qty_Rec > 0 and Tax_YN=1 and MRP_YN=1 and Part_No='" & Part_No & "'")
Set MRPTPRst = GCn.Execute("Select SP_Stock.V_Date,SP_Stock.Qty_Rec,SP_Purch.Party_Doc_No,SP_Purch.Party_Doc_Date from SP_Stock Left Join SP_Purch on Sp_Stock.DocId=SP_Purch.DocId  where Qty_Rec > 0 and Tax_YN=0 and MRP_YN=1 and Part_No='" & Part_No & "'")

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
            MRPTPCurrStk = XNull(MRPTPCurrStk - MRPTPRst!Qty_Rec)
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
End If
End Function

'Sub CreateLog()
'    Dim RsTemp As ADODB.Recordset
'    Dim rsTemp1 As ADODB.Recordset
'    Dim mTotalQty As Double
'    Dim mTotalItem As Double
'    Dim mDiscount As Double
'    Dim mGoodsValue As Double
'
'
'
'    Set RsTemp = GCn.Execute("Select * From Sp_Sale Where DocId='" & Master!SearchCode & "'")
'    If RsTemp.RecordCount > 0 Then
'        Set rsTemp1 = GCn.Execute("Select Count(*) As Item, Sum(Qty_Iss) As Qty, Sum(Net_Amt2) As Amt From Sp_Stock Where Invoice_DocId='" & Master!SearchCode & "'")
'        If RsTemp.RecordCount > 0 Then
'            mTotalItem = VNull(rsTemp1!Item)
'            mTotalQty = VNull(rsTemp1!Qty)
'            mGoodsValue = VNull(rsTemp1!Amt)
'        End If
'        mDiscount = VNull(RsTemp!D_Amt_TB) + VNull(RsTemp!D_Amt_TP)
'        GCn.Execute "Insert Into DeleteLog(DocId, Type, VType, VDate, Total_Item, " & _
'                            "Total_Qty, GoodsValue, Bill_Amt, Discount, Addition, " & _
'                            "Deduction, AutoYn, User_Name, Del_Date, Del_Time, EditDate, EditTime) Values( " & _
'                            "'" & Master!SearchCode & "', 'Spare Sale Bill', '" & txt(DocType) & "', " & ConvertDate(XNull(RsTemp!V_DATE)) & ", " & mTotalItem & ", " & _
'                            "" & mTotalQty & ", " & mGoodsValue & ", " & VNull(RsTemp!Total_Amt) & ", " & mDiscount & ", " & RsTemp!Addition & ", " & _
'                            "0, '" & IIf(mReposting, "Y", "N") & "', '" & pubUName & "', " & ConvertDate(IIf(TopCtrl1.TopText2 = "Browse", PubLoginDate, "")) & ", '" & Time & "', " & ConvertDate(IIf(TopCtrl1.TopText2 = "Edit", PubLoginDate, "")) & ", '" & Time & "')"
'    End If
'End Sub


Sub Ini_Pub()
    Dim RsTemp As ADODB.Recordset
    
    Set RsTemp = GCn.Execute("Select CheckNegetiveStockSiteWise From Syctrl")
    If RsTemp.RecordCount > 0 Then
        mCheckNegetiveStockSiteWise = VNull(RsTemp!CheckNegetiveStockSiteWise)
    End If
    
    mVatYn = PubVATYN
End Sub

Sub DispTextVat()
        With FGrid
            If mVatYn = 1 Then
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
                
                .TextMatrix(0, Col_TaxAmt1) = ""
                .ColAlignmentFixed(Col_TaxAmt1) = flexAlignRightCenter
                .ColWidth(Col_TaxAmt1) = 0
                
                .ColWidth(Col_SatPer) = 0
                .ColWidth(Col_SatAmt) = 0
            End If
        End With
End Sub


