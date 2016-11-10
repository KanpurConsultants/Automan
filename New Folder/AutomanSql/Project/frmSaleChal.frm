VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmSaleChal 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Dispatch Challan Entry"
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
   LinkTopic       =   "form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   12195
   WhatsThisHelp   =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   5
      Left            =   930
      MaxLength       =   40
      TabIndex        =   4
      Top             =   405
      Width           =   3855
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   10
      Left            =   6150
      MaxLength       =   7
      TabIndex        =   189
      ToolTipText     =   "Press L-> Local or C-> Central"
      Top             =   405
      Width           =   1095
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
      Height          =   240
      Index           =   53
      Left            =   9540
      TabIndex        =   186
      Top             =   4980
      Width           =   1890
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Re-Post"
      Height          =   345
      Left            =   8475
      TabIndex        =   185
      Top             =   15
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DGOrdPart 
      Height          =   2625
      Left            =   390
      Negotiate       =   -1  'True
      TabIndex        =   161
      TabStop         =   0   'False
      Top             =   8580
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
      Left            =   180
      TabIndex        =   183
      Top             =   3150
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
      Left            =   5835
      TabIndex        =   171
      Top             =   2400
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
         Picture         =   "frmSaleChal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   178
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
         Picture         =   "frmSaleChal.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   177
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmSaleChal.frx":0678
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
         TabIndex        =   176
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmSaleChal.frx":0982
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
         TabIndex        =   175
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmSaleChal.frx":0C8C
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
         TabIndex        =   174
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
         TabIndex        =   173
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
         TabIndex        =   172
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
         Left            =   -45
         TabIndex        =   181
         Top             =   330
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
         TabIndex        =   180
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
         TabIndex        =   179
         Top             =   0
         Width           =   4695
      End
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
      Height          =   240
      Index           =   51
      Left            =   9540
      TabIndex        =   47
      ToolTipText     =   "Turn Over Tax %"
      Top             =   6000
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
      Height          =   240
      Index           =   52
      Left            =   10170
      TabIndex        =   48
      Top             =   6000
      Width           =   1260
   End
   Begin MSDataGridLib.DataGrid DGPerson 
      Height          =   3330
      Left            =   1665
      Negotiate       =   -1  'True
      TabIndex        =   169
      TabStop         =   0   'False
      Top             =   8295
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
      Left            =   390
      Negotiate       =   -1  'True
      TabIndex        =   164
      TabStop         =   0   'False
      Top             =   8550
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
   Begin MSDataGridLib.DataGrid DGSONo 
      Height          =   2775
      Left            =   -6345
      Negotiate       =   -1  'True
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   7545
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
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   3330
      Left            =   1485
      Negotiate       =   -1  'True
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   8625
      Visible         =   0   'False
      Width           =   9750
      _ExtentX        =   17198
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
            ColumnWidth     =   3839.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3869.858
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   5610
      Negotiate       =   -1  'True
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   8490
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
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
   Begin VB.Frame FrmDetail 
      BackColor       =   &H00CAF1FD&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   2205
      Left            =   -6090
      TabIndex        =   89
      Top             =   705
      Visible         =   0   'False
      Width           =   6285
      Begin VB.Line Line4 
         X1              =   3660
         X2              =   3885
         Y1              =   1035
         Y2              =   1035
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
         TabIndex        =   168
         Top             =   1395
         Width           =   360
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
         TabIndex        =   167
         Top             =   1410
         Width           =   1080
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
         TabIndex        =   166
         Top             =   465
         Width           =   1095
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
         TabIndex        =   165
         Top             =   465
         Width           =   885
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
         TabIndex        =   116
         Top             =   1635
         Width           =   645
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
         TabIndex        =   115
         Top             =   0
         Width           =   6285
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
         TabIndex        =   114
         Top             =   915
         Width           =   1110
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
         TabIndex        =   113
         Top             =   1185
         Width           =   1080
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
         TabIndex        =   112
         Top             =   1875
         Width           =   705
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
         TabIndex        =   111
         Top             =   930
         Width           =   810
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
         TabIndex        =   110
         Top             =   1650
         Width           =   345
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
         TabIndex        =   109
         Top             =   675
         Width           =   1005
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
         TabIndex        =   108
         Top             =   1185
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
         TabIndex        =   107
         Top             =   1410
         Width           =   390
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
         TabIndex        =   106
         Top             =   1410
         Width           =   765
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
         TabIndex        =   105
         Top             =   675
         Width           =   1590
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
         TabIndex        =   104
         Top             =   1170
         Width           =   360
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
         TabIndex        =   103
         Top             =   930
         Width           =   1095
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
         TabIndex        =   102
         Top             =   1657
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
         Index           =   10
         Left            =   3285
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
         Index           =   6
         Left            =   2115
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
         ForeColor       =   &H00C000C0&
         Height          =   210
         Index           =   14
         Left            =   5460
         TabIndex        =   98
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   3285
         TabIndex        =   97
         Top             =   1875
         Width           =   360
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
         TabIndex        =   96
         Top             =   1185
         Width           =   360
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
         TabIndex        =   95
         Top             =   1185
         Width           =   885
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
         TabIndex        =   94
         Top             =   930
         Width           =   660
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
         TabIndex        =   93
         Top             =   270
         Width           =   660
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
         TabIndex        =   92
         Top             =   255
         Width           =   825
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
         TabIndex        =   91
         Top             =   255
         Width           =   930
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
         TabIndex        =   90
         Top             =   255
         Width           =   1020
      End
      Begin VB.Line Line1 
         X1              =   1755
         X2              =   75
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line2 
         X1              =   2760
         X2              =   2475
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line3 
         X1              =   3750
         X2              =   3750
         Y1              =   1035
         Y2              =   2070
      End
   End
   Begin MSDataGridLib.DataGrid DGTrans 
      Height          =   3330
      Left            =   1380
      Negotiate       =   -1  'True
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   8535
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
   Begin MSDataGridLib.DataGrid DGGodown 
      Height          =   3330
      Left            =   10410
      Negotiate       =   -1  'True
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   8385
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
      Left            =   9915
      Negotiate       =   -1  'True
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   8415
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   50
      Left            =   3705
      MaxLength       =   10
      TabIndex        =   9
      Text            =   "0123456789"
      Top             =   1425
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   49
      Left            =   930
      MaxLength       =   40
      TabIndex        =   8
      Top             =   1425
      Width           =   1515
   End
   Begin MSDataGridLib.DataGrid DGCrAc 
      Height          =   3360
      Left            =   780
      Negotiate       =   -1  'True
      TabIndex        =   149
      TabStop         =   0   'False
      Top             =   8535
      Visible         =   0   'False
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   5927
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
            ColumnWidth     =   3225.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2489.953
         EndProperty
      EndProperty
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
      Height          =   240
      Index           =   0
      Left            =   9480
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   525
      Width           =   2295
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
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   15
      Left            =   6150
      MaxLength       =   3
      TabIndex        =   17
      Text            =   "WD"
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   7890
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   10095
      TabIndex        =   138
      Top             =   8400
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   60
         TabIndex        =   139
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   18
      Left            =   3285
      MaxLength       =   10
      TabIndex        =   12
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   8
      Left            =   930
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1170
      Width           =   3855
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
      Height          =   240
      Index           =   38
      Left            =   9540
      TabIndex        =   38
      Top             =   4470
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
      Height          =   240
      Index           =   47
      Left            =   9540
      TabIndex        =   50
      Top             =   6510
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
      Height          =   240
      Index           =   22
      Left            =   4485
      TabIndex        =   24
      Top             =   4725
      Width           =   1845
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
      Height          =   240
      Index           =   21
      Left            =   2295
      TabIndex        =   23
      Top             =   4725
      Width           =   1905
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
      Height          =   240
      Index           =   20
      Left            =   4485
      TabIndex        =   22
      Top             =   4470
      Width           =   1845
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
      Left            =   2505
      TabIndex        =   19
      Top             =   2850
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   6
      Left            =   930
      MaxLength       =   40
      TabIndex        =   5
      Top             =   660
      Width           =   3855
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
      Height          =   240
      Index           =   1
      Left            =   9840
      TabIndex        =   0
      ToolTipText     =   "Press S-> Sales Challan or T-> Transfer"
      Top             =   1005
      Width           =   1575
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   7
      Left            =   930
      MaxLength       =   40
      TabIndex        =   6
      Top             =   915
      Width           =   3855
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
      Height          =   240
      Index           =   4
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Cash/C"
      ToolTipText     =   "Press C-> Cash or R-> Credit"
      Top             =   1815
      Visible         =   0   'False
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   13
      Left            =   6150
      MaxLength       =   15
      TabIndex        =   15
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   14
      Left            =   6150
      MaxLength       =   11
      TabIndex        =   16
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   11
      Left            =   6150
      MaxLength       =   25
      TabIndex        =   13
      Top             =   660
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
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   16
      Left            =   6150
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   1680
      Width           =   2430
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   12
      Left            =   6150
      MaxLength       =   40
      TabIndex        =   14
      Top             =   915
      Width           =   2415
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
      Height          =   240
      Index           =   46
      Left            =   9540
      TabIndex        =   49
      Top             =   6255
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
      Height          =   240
      Index           =   45
      Left            =   10170
      TabIndex        =   46
      Top             =   5745
      Width           =   1260
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
      Height          =   240
      Index           =   44
      Left            =   9540
      TabIndex        =   45
      ToolTipText     =   "Turn Over Tax %"
      Top             =   5745
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
      Height          =   240
      Index           =   43
      Left            =   9540
      TabIndex        =   44
      Top             =   5490
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
      Height          =   255
      Index           =   42
      Left            =   10740
      TabIndex        =   42
      Top             =   7515
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
      Left            =   10065
      TabIndex        =   41
      ToolTipText     =   "Surcharge % on Local Sales Tax"
      Top             =   7515
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
      Height          =   240
      Index           =   40
      Left            =   10170
      TabIndex        =   40
      Top             =   4725
      Width           =   1260
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
      Height          =   240
      Index           =   39
      Left            =   9540
      TabIndex        =   39
      ToolTipText     =   "Local Sales Tax %"
      Top             =   4725
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
      Height          =   240
      Index           =   37
      Left            =   2295
      TabIndex        =   37
      Top             =   6255
      Width           =   1905
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
      Height          =   240
      Index           =   36
      Left            =   2925
      TabIndex        =   36
      Top             =   6000
      Width           =   1275
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
      Height          =   240
      Index           =   35
      Left            =   2295
      TabIndex        =   35
      ToolTipText     =   "General Surcharge %"
      Top             =   6000
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
      Height          =   240
      Index           =   19
      Left            =   2295
      TabIndex        =   21
      Top             =   4470
      Width           =   1905
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
      Height          =   240
      Index           =   34
      Left            =   9540
      TabIndex        =   43
      Top             =   5235
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
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   33
      Left            =   2295
      TabIndex        =   52
      Text            =   "WithDrawn"
      Top             =   6510
      Visible         =   0   'False
      Width           =   1905
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
      Height          =   240
      Index           =   32
      Left            =   4485
      TabIndex        =   34
      Top             =   5745
      Width           =   1845
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
      Height          =   240
      Index           =   31
      Left            =   2295
      TabIndex        =   33
      Text            =   "99999999.99"
      Top             =   5745
      Width           =   1905
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
      Height          =   240
      Index           =   30
      Left            =   5115
      TabIndex        =   32
      Top             =   5490
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
      Height          =   240
      Index           =   29
      Left            =   4485
      TabIndex        =   31
      ToolTipText     =   "Discount % Taxpaid"
      Top             =   5490
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
      Height          =   240
      Index           =   28
      Left            =   2925
      TabIndex        =   30
      Text            =   "99999999.99"
      Top             =   5490
      Width           =   1275
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
      Height          =   240
      Index           =   27
      Left            =   2295
      TabIndex        =   29
      Text            =   "99.99"
      ToolTipText     =   "Discount % Taxable"
      Top             =   5490
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
      Height          =   240
      Index           =   26
      Left            =   4485
      TabIndex        =   28
      Top             =   5235
      Width           =   1845
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
      Height          =   240
      Index           =   25
      Left            =   2295
      TabIndex        =   27
      Top             =   5235
      Width           =   1905
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
      Height          =   240
      Index           =   24
      Left            =   4485
      TabIndex        =   26
      Top             =   4980
      Width           =   1845
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
      Height          =   240
      Index           =   23
      Left            =   2295
      TabIndex        =   25
      Top             =   4980
      Width           =   1905
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   17
      Left            =   930
      TabIndex        =   11
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
      Height          =   240
      Index           =   2
      Left            =   9840
      MaxLength       =   11
      TabIndex        =   1
      Top             =   1275
      Width           =   1575
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   9
      Left            =   930
      MaxLength       =   40
      TabIndex        =   10
      Top             =   1680
      Width           =   3855
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
      Height          =   240
      Index           =   3
      Left            =   10500
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1545
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1935
      Left            =   30
      TabIndex        =   20
      Top             =   2250
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   0
      Cols            =   31
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   15595518
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      FocusRect       =   0
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "kkk"
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
      _Band(0).Cols   =   31
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   3330
      Left            =   7335
      Negotiate       =   -1  'True
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   8265
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
      Left            =   4470
      TabIndex        =   51
      Top             =   6345
      Width           =   1440
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   384
      Left            =   0
      TabIndex        =   184
      Top             =   0
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   688
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
      TabIndex        =   191
      Top             =   6750
      Width           =   45
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Despatch Type"
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
      Left            =   4830
      TabIndex        =   190
      ToolTipText     =   "Press L-> Local or C-> Central"
      Top             =   405
      Width           =   1230
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Tax"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   43
      Left            =   7575
      TabIndex        =   188
      Top             =   4980
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   17
      Left            =   9105
      TabIndex        =   187
      Top             =   4980
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
      Left            =   7560
      TabIndex        =   182
      Top             =   435
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   42
      Left            =   7575
      TabIndex        =   170
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   6015
      Width           =   1575
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
      Left            =   6420
      TabIndex        =   163
      Top             =   6345
      Width           =   660
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
      Left            =   6240
      TabIndex        =   162
      Top             =   6075
      Width           =   1200
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      Height          =   600
      Left            =   6195
      Shape           =   4  'Rounded Rectangle
      Top             =   6075
      Width           =   1290
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
      Index           =   16
      Left            =   9705
      TabIndex        =   160
      Top             =   1800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Credit"
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
      Left            =   8670
      TabIndex        =   159
      ToolTipText     =   "Press C-> Cash or R-> Credit"
      Top             =   1800
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   1665
      Left            =   8595
      Shape           =   4  'Rounded Rectangle
      Top             =   465
      Width           =   3240
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
      TabIndex        =   156
      Top             =   1545
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
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   36
      Left            =   2505
      TabIndex        =   152
      Top             =   1425
      Width           =   975
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
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   35
      Left            =   45
      TabIndex        =   151
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   25
      Left            =   9360
      TabIndex        =   147
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
      Left            =   8670
      TabIndex        =   146
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
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   9
      Left            =   45
      TabIndex        =   145
      Top             =   405
      Width           =   405
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   29
      Left            =   5100
      TabIndex        =   143
      ToolTipText     =   "Press Y-> Yes or N-> No"
      Top             =   7800
      Visible         =   0   'False
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   21
      Left            =   6030
      TabIndex        =   142
      Top             =   7890
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   10425
      TabIndex        =   141
      Top             =   780
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   8670
      TabIndex        =   140
      Top             =   780
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
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   40
      Left            =   2235
      TabIndex        =   137
      Top             =   1965
      Width           =   885
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   39
      Left            =   45
      TabIndex        =   136
      Top             =   1170
      Width           =   435
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
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   38
      Left            =   45
      TabIndex        =   135
      Top             =   915
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   1
      Left            =   8670
      TabIndex        =   134
      ToolTipText     =   "Press S-> Sales Challan or T-> Transfer"
      Top             =   1005
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   11
      Left            =   180
      TabIndex        =   133
      Top             =   4980
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   19
      Left            =   180
      TabIndex        =   132
      Top             =   4470
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
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   10
      Left            =   45
      TabIndex        =   131
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   33
      Left            =   180
      TabIndex        =   130
      Top             =   4725
      Width           =   1665
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
      Left            =   7575
      TabIndex        =   129
      Top             =   4485
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
      TabIndex        =   128
      Top             =   4470
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Despatch Mode"
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
      Index           =   34
      Left            =   4830
      TabIndex        =   127
      Top             =   660
      Width           =   1290
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
      TabIndex        =   126
      Top             =   6525
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   30
      Left            =   1965
      TabIndex        =   125
      Top             =   4725
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   7
      Left            =   7275
      TabIndex        =   124
      Top             =   4215
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   11100
      TabIndex        =   123
      Top             =   4215
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9090
      TabIndex        =   122
      Top             =   4215
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   25
      Left            =   9630
      TabIndex        =   121
      Top             =   4215
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   28
      Left            =   8850
      TabIndex        =   120
      Top             =   4215
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   27
      Left            =   10905
      TabIndex        =   119
      Top             =   4215
      Width           =   180
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
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   6
      Left            =   4830
      TabIndex        =   118
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
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   5
      Left            =   4830
      TabIndex        =   117
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   28
      Left            =   4830
      TabIndex        =   87
      Top             =   1680
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
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   27
      Left            =   4830
      TabIndex        =   86
      Top             =   915
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   18
      Left            =   9105
      TabIndex        =   85
      Top             =   6255
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   26
      Left            =   7575
      TabIndex        =   84
      Top             =   6255
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   25
      Left            =   7575
      TabIndex        =   83
      ToolTipText     =   "Turnover Tax on Taxable + Taxpaid Amount"
      Top             =   5760
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   24
      Left            =   7575
      TabIndex        =   82
      Top             =   5505
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   15
      Left            =   9630
      TabIndex        =   81
      Top             =   7515
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   23
      Left            =   8100
      TabIndex        =   80
      Top             =   7515
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   14
      Left            =   9105
      TabIndex        =   79
      Top             =   4725
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   22
      Left            =   7575
      TabIndex        =   78
      Top             =   4725
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   13
      Left            =   1965
      TabIndex        =   77
      Top             =   6255
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   21
      Left            =   180
      TabIndex        =   76
      Top             =   6255
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   12
      Left            =   1965
      TabIndex        =   75
      Top             =   6000
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   20
      Left            =   180
      TabIndex        =   74
      Top             =   6000
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   11
      Left            =   1965
      TabIndex        =   73
      Top             =   4470
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   10
      Left            =   9105
      TabIndex        =   72
      Top             =   5235
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   18
      Left            =   7575
      TabIndex        =   71
      Top             =   5250
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   9
      Left            =   1965
      TabIndex        =   70
      Top             =   6510
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Addition"
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
      Index           =   17
      Left            =   180
      TabIndex        =   69
      Top             =   6510
      Visible         =   0   'False
      Width           =   660
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   8
      Left            =   1965
      TabIndex        =   68
      Top             =   5745
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   16
      Left            =   180
      TabIndex        =   67
      Top             =   5745
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   7
      Left            =   1965
      TabIndex        =   66
      Top             =   5490
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   15
      Left            =   180
      TabIndex        =   65
      Top             =   5490
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   6
      Left            =   1965
      TabIndex        =   64
      Top             =   5235
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   14
      Left            =   180
      TabIndex        =   63
      Top             =   5235
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
      Top             =   4215
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
      Top             =   4215
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   5
      Left            =   1965
      TabIndex        =   60
      Top             =   4980
      Width           =   180
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
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   8
      Left            =   45
      TabIndex        =   59
      Top             =   1935
      Width           =   735
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salesman"
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
      Index           =   7
      Left            =   45
      TabIndex        =   58
      Top             =   1680
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   90
      Left            =   9705
      TabIndex        =   57
      Top             =   1005
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
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   92
      Left            =   9705
      TabIndex        =   56
      Top             =   1545
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
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   93
      Left            =   9705
      TabIndex        =   55
      Top             =   1275
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   0
      Left            =   8670
      TabIndex        =   54
      Top             =   1275
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   2
      Left            =   8670
      TabIndex        =   53
      Top             =   1545
      Width           =   690
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   600
      Left            =   4305
      Shape           =   4  'Rounded Rectangle
      Top             =   6075
      Width           =   1770
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
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   30
      Left            =   4395
      TabIndex        =   88
      Top             =   6090
      Width           =   1590
   End
End
Attribute VB_Name = "frmSaleChal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMRevDisTBPer As Double, mMRevDisTPPer As Double
Dim mTBDisAmtMRP As Double, mTPDisAmtMRP As Double
Dim mMRPTax As Double, mMRPTaxSur As Double, mMRPTOT As Double, mMRPReSales As Double
Dim mMRPLubeTB As Double, mMRPLubeTP  As Double

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
Dim mVType As String
Dim mVPrefix As String
Dim mSearchCode As String
Dim mPartyType As Byte         'Used to Detect Part Rate in GetRate Function
Dim ForSiteCode As String
Dim CutCol As Byte
Dim mCheckNegetiveStockSiteWise As Boolean
Private Const SalChalType As String = "SYSC"
Private Const TrfChalType As String = "SYSCT"
Private Const DocTypeChal As String = "Sale Challan"
Private Const DocTypeTrf As String = "Stock Transfer"

Private Const CellBackColLeave As String = &HEDF7FE
Private Const GridBackColorBkg As String = &HD7C6C8
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

' Under observation
Dim VoucherEditFlag As Boolean                  ' Used for whether we can edit voucher no or not
' End Under observation
Dim ListArray As Variant
Dim mListItem As ListItem

Dim mSatYn As Boolean


Private Const DocID As Byte = 0                 ' Doc.ID
Private Const DocType As Byte = 1               ' Document Type
Private Const VDate As Byte = 2                 ' Date
Private Const SerialNo As Byte = 3              ' Serial No.
Private Const CashCr As Byte = 4                ' Cash/Credit
Private Const Party As Byte = 5                 ' Party Name
Private Const Address1 As Byte = 6              ' Address1
Private Const CrAc As Byte = 7                  ' Debit A/c
Private Const FormName As Byte = 8              ' Form Name
Private Const Form31Name As Byte = 49           ' Form 31 Name
Private Const Form31No As Byte = 50             ' Form 31 No
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
Private Const Addition As Byte = 33             'withdrawn
Private Const PackCrg As Byte = 34              '
Private Const GenSurPer As Byte = 35            '
Private Const GenSurAmt As Byte = 36            '
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
Private Const ReSalTaxPer As Byte = 51          '
Private Const ReSalTaxAmt As Byte = 52          '
Private Const SatAmt As Byte = 53          '

'* Grid Column Declaration
Private Col_SrNo As Byte                 ' Serial No
Private Col_SONo As Byte                 ' Sale Order No Name
Private Col_SOSrNo As Byte              ' Sale Order Serial No
Private Col_PNo As Byte                  ' Part No
Private Col_SONoCode As Byte             ' Sale Order No Code
Private Col_Unit As Byte                 ' Unit
Private Col_MRP As Byte                  ' MRP Yes/No
Private Col_Taxable As Byte              ' Taxable Yes/No
Private Col_Qty As Byte                  ' Qty
Private Col_Rate As Byte                 ' Rate
Private Col_MRPRate As Byte             ' MRP Rate
Private Col_Amt As Byte                 ' Amt
Private Col_DiscPer As Byte             ' Disc. %
Private Col_DiscAmt As Byte             ' Disc. Amt.
Private Col_TaxPer As Byte              ' Tax Per.
Private Col_TaxAmt1 As Byte             ' Tax. Amt.
Private Col_SatPer As Byte              ' SAT Per.
Private Col_SatAmt1 As Byte             ' SAT Amt.
Private Col_ItemVal As Byte             ' Item Value
Private Col_GodownCode As Byte          ' Godown Code
Private Col_Godown As Byte              ' Godown
Private Col_PartSrlNo As Byte           ' Part Serial No
Private Col_PName As Byte               ' Part Name
Private Col_LName As Byte               ' Local Name
Private Col_MRPStkTP As Byte            ' MRP TP Qty 'Current Stock Qty
Private Col_MRPStkTB As Byte            ' MRP TB Qty
Private Col_TBStk As Byte               ' Taxbale Qty
Private Col_TPStk As Byte               ' Tax Paid Qty
Private Col_TBRate As Byte              ' Taxbale Rate
Private Col_TPRate As Byte              ' Tax Paid Rate
Private Col_Bin As Byte                 ' Bin
Private Col_LastRate As Byte            ' Last Purchase Rate
Private Col_HPRate As Byte              ' High Purchase Rate
Private Col_LPRate As Byte              ' Low Purchase Rate
Private Col_PartGrade As Byte           ' Part Grade (Used for Oil Item)
Private Col_EffectDate As Byte          ' MRP Effective Date/TB Effective Date

Private Const FromVno As Byte = 0
Private Const ToVno As Byte = 1
Private Const VType1 As Byte = 2

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim rsTaxPer As ADODB.Recordset
Dim mRepName As String

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To txt.Count - 1
        If I = DocID Or I = IWDiscTotTB Or I = IWDiscTotTP Or I = MRPAmtTB _
            Or I = MRPAmtTP Or I = SprAmtTB Or I = SprAmtTP Or I = OilAmtTB _
            Or I = OilAmtTP Or I = STotATB Or I = STotATP Or I = TaxableTot _
            Or I = STotB Or I = SROff Or I = NetSprAmt Or I = NetAmt Then
        Else
            txt(I).Enabled = Enb
        End If
    Next
    txt(STaxPer).Enabled = False
    txt(STaxAmt).Enabled = False
    txt(TaxSurPer).Enabled = False
    txt(TaxSurAmt).Enabled = False
 
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("SearchCode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select S.DocID As SearchCode From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' and S.V_Type In ('" & SalChalType & "', '" & TrfChalType & "') And S.DocID  = '" & MyValue & "' " _
            & "  Order by S.V_Date Desc,S.DocID desc")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
Exit Sub
ELoop:
    CheckError
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
    lblGatePass.CAPTION = ""
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End Sub
'* Used for intialize grid columns
Private Sub Grid_Ini()
'Serial No  | Part No | Part Name |Unit | MRP Yes/No | Taxable Yes/No  | Qty | Rate | Amt | Disc. % | Disc. Amt. | Item Value | Local Name
    With FGrid
        .left = Me.left '+ 60
        .width = Me.width - 90
        .top = 2250
        .BackColor = CellBackColLeave
        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 36

        .TextMatrix(0, Col_SrNo) = "S.No"
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 420
        
        .TextMatrix(0, Col_SONoCode) = "SO No.Code"
        .ColAlignment(Col_SONoCode) = flexAlignLeftCenter
        .ColWidth(Col_SONoCode) = 0

        .TextMatrix(0, Col_SONo) = "SO No."
        .ColAlignmentFixed(Col_SONo) = flexAlignLeftCenter
        .ColAlignment(Col_SONo) = flexAlignLeftCenter
        .ColWidth(Col_SONo) = 1150

        .TextMatrix(0, Col_SOSrNo) = "SO Srl No."
        .ColAlignment(Col_SOSrNo) = flexAlignLeftCenter
        .ColWidth(Col_SOSrNo) = 0
        
        .TextMatrix(0, Col_PNo) = "Part No."
        .ColAlignmentFixed(Col_PNo) = flexAlignLeftCenter
        .ColAlignment(Col_PNo) = flexAlignLeftCenter
        .ColWidth(Col_PNo) = 1500

        .TextMatrix(0, Col_Unit) = "Unit"
        .ColAlignmentFixed(Col_Unit) = flexAlignLeftCenter
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
        .ColAlignment(Col_Qty) = flexAlignRightCenter
        .ColWidth(Col_Qty) = 960

        .TextMatrix(0, Col_Rate) = "Rate"
        .ColAlignmentFixed(Col_Rate) = flexAlignRightCenter
        .ColAlignment(Col_Rate) = flexAlignRightCenter
        .ColWidth(Col_Rate) = 870

        .TextMatrix(0, Col_MRPRate) = "MRP Rate"
        .ColAlignmentFixed(Col_MRPRate) = flexAlignRightCenter
        .ColAlignment(Col_MRPRate) = flexAlignRightCenter
        .ColWidth(Col_MRPRate) = 870

        .TextMatrix(0, Col_Amt) = "Amount"
        .ColAlignmentFixed(Col_Amt) = flexAlignRightCenter
        .ColAlignment(Col_Amt) = flexAlignRightCenter
        .ColWidth(Col_Amt) = 1065

        .TextMatrix(0, Col_DiscPer) = "Disc%"
        .ColAlignmentFixed(Col_DiscPer) = flexAlignRightCenter
        .ColAlignment(Col_DiscPer) = flexAlignRightCenter
        .ColWidth(Col_DiscPer) = 555

        .TextMatrix(0, Col_DiscAmt) = "Disc.Amt"
        .ColAlignmentFixed(Col_DiscAmt) = flexAlignRightCenter
        .ColAlignment(Col_DiscAmt) = flexAlignRightCenter
        .ColWidth(Col_DiscAmt) = 840
        
        If PubVATYN = 1 Then
            .TextMatrix(0, Col_TaxPer) = "TaxPer"
            .ColAlignmentFixed(Col_TaxPer) = flexAlignRightCenter
            .ColWidth(Col_TaxPer) = 840
            
            .TextMatrix(0, Col_TaxAmt1) = "TaxAmt"
            .ColAlignmentFixed(Col_TaxAmt1) = flexAlignRightCenter
            .ColWidth(Col_TaxAmt1) = 840
        
            If PubSatYn = 1 Then
                .TextMatrix(0, Col_SatPer) = "Sat %"
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

        .TextMatrix(0, Col_ItemVal) = "Item Value"
        .ColAlignmentFixed(Col_ItemVal) = flexAlignRightCenter
        .ColAlignment(Col_ItemVal) = flexAlignRightCenter
        .ColWidth(Col_ItemVal) = 1095

        .TextMatrix(0, Col_GodownCode) = "Godown Code"
        .ColAlignment(Col_GodownCode) = flexAlignLeftCenter
        .ColWidth(Col_GodownCode) = 0

        .TextMatrix(0, Col_Godown) = "Godown"
        .ColAlignmentFixed(Col_Godown) = flexAlignLeftCenter
        .ColAlignment(Col_Godown) = flexAlignLeftCenter
        .ColWidth(Col_Godown) = 1200
        
        .TextMatrix(0, Col_PartSrlNo) = "Part SrlNo"
        .ColAlignmentFixed(Col_PartSrlNo) = flexAlignLeftCenter
        .ColAlignment(Col_PartSrlNo) = flexAlignLeftCenter
        .ColWidth(Col_PartSrlNo) = 1200
        
        .TextMatrix(0, Col_PName) = "Part Name"
        .ColAlignmentFixed(Col_PName) = flexAlignLeftCenter
        .ColAlignment(Col_PName) = flexAlignLeftCenter
        .ColWidth(Col_PName) = 2500

        .TextMatrix(0, Col_LName) = "Local Name"
        .ColAlignmentFixed(Col_LName) = flexAlignLeftCenter
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
        .ColAlignment(Col_Bin) = flexAlignLeftCenter
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
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    
    DGPart.width = FGrid.width: DGPart.left = FGrid.left: DGPart.top = FGrid.top + FGrid.height: DGPart.height = Me.height - (DGPart.top + mBotScale)
    DGSONo.left = FGrid.left: DGSONo.top = DGPart.top: DGSONo.height = Me.height - (FGrid.top + FGrid.height + mBotScale)
    DGGodown.left = Me.width - (DGGodown.width + mRtScale): DGGodown.top = DGPart.top: DGGodown.height = Me.height - (FGrid.top + FGrid.height + mBotScale)
    FrmDetail.width = 6285: FrmDetail.left = Me.width - (FrmDetail.width + mRtScale): FrmDetail.top = mTopScale: FrmDetail.height = 2130
    'DGParty.left = Me.width - (DGParty.width + mRtScale): DGParty.top = mTopScale
    DGParty.left = mRtScale: DGParty.top = mTopScale + FGrid.top
    DGCrAc.left = Me.width - (DGCrAc.width + mRtScale): DGCrAc.top = mTopScale
    DGForm.left = Me.width - (DGForm.width + mRtScale): DGForm.top = mTopScale
    DGForm31.left = Me.width - (DGForm31.width + mRtScale): DGForm31.top = mTopScale
    DGTrans.left = mLtScale: DGTrans.top = mTopScale
    
    FrmPrn.left = (Me.width - FrmPrn.width) / 2: FrmPrn.top = (Me.height - FrmPrn.height) / 2
    DGOrdPart.left = (Me.width - DGOrdPart.width) / 2: DGOrdPart.top = FGrid.top + FGrid.height: DGOrdPart.height = Me.height - (FGrid.top + FGrid.height + mBotScale)
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
    If DGPerson.Visible = True Then DGPerson.Visible = False
End Sub

Private Sub cmdPost_Click()
Dim I As Integer
Dim LedgAry(4) As LedgRec, mNarr$, mResult As Byte
Dim mAmount As Double, TaxAmt As Double, DisAmt As Double, OrdDisAmt1 As Double
Dim TTaxAmt As Double
        
    Master.MoveFirst
    Do Until Master.EOF
        Call MoveRec
        
        If CDate(txt(VDate).TEXT) < PubStartDate Or CDate(txt(VDate).TEXT) > PubEndDate Then GoTo MyNextRecord
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        Call TopCtrl1_eEdit
        Call Txt_Validate(STaxAmt, False)
        If txt(DocType).TEXT = DocTypeTrf Then
            mNarr = "Through Stock Transfer Issue"
            I = 0
            LedgAry(I).SubCode = txt(Party).Tag
            LedgAry(I).AmtDr = Val(txt(NetAmt))
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = txt(CrAc).Tag
            I = I + 1
            LedgAry(I).SubCode = "12008041" 'Txt(CrAc).Tag      '' Vishal Jain
            LedgAry(I).AmtCr = Val(txt(NetAmt))
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = txt(Party).Tag

            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, txt(DocID), CDate(txt(VDate)), mNarr & "[Common]")
            If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
        End If
MyNextRecord:
        Disp_Text SETS("INI", Me, Master)
        Master.MoveNext
    Loop

End Sub
Private Sub DGParty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
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
    FrmDetail.Visible = False
    FGrid.Tag = ""
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
        If InStr(Me.TopCtrl1.Tag, "D") <> 0 Then Me.TopCtrl1.tDel = True
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "Select S.*,SubGroup.Name As PartyName,SubGroup.Party_Type,SubGroupCr.Name As CreditAcName, " _
            & "TaxForms.Form_Desc As FormName,TaxForms31.Form_Desc As Form31Name,Emp.Emp_Name " _
            & "From (((((SP_Sale S Left Join SubGroup on S.Party_Code=SubGroup.SubCode) " _
            & "Left Join SubGroup SubGroupCr on S.CrAc=SubGroupCr.SubCode) " _
            & "Left Join TaxForms on S.Form_Code=TaxForms.Form_Code) " _
            & "Left Join TaxForms TaxForms31 on S.RoadPermit_FormCode=TaxForms31.Form_Code) " _
            & "Left Join Emp_Mast Emp on S.Rep_Code=Emp.Emp_Code) " _
            & "Where S.DocID = '" & Master!SearchCode & "' " _
            & "Order by S.V_Date,S.V_Type", GCn, adOpenStatic, adLockReadOnly
        If Master1!CancelYN = 1 Then
            LblCancel.Visible = True
            TopCtrl1.tEdit = False
        Else
            LblCancel.Visible = False
        End If
        If Master1!Invoice_DocID <> "" Then
            TopCtrl1.tEdit = False
            TopCtrl1.tDel = False
        End If
        txt(DocID).TEXT = Master1!DocID
        mSearchCode = txt(DocID)
        
        If PubBackEnd = "A" Then
            mSatYn = IIf(VNull(Master1!SAT_YN) = 1, True, False)
        Else
            mSatYn = IIf(VNull(Master1!SAT_YN) = True, True, False)
        End If
        DispTextVat
        
        LblDiv.CAPTION = "Division : " & left(Master1!DocID, 1)
        LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        LblUser = IIf(Not IsNull(Master1!AddDate), "Add By : " & XNull(Master1!AddBy) & "  Dated : " & XNull(Master1!AddDate), "") & IIf(Not IsNull(Master1!ModifyDate), "     Modify By : " & XNull(Master1!ModifyBy) & "  Dated : " & XNull(Master1!ModifyDate), "")
        lblGatePass.CAPTION = Master1!gp_no
        mVType = Master1!V_Type
        If mVType = SalChalType Then
            txt(DocType).TEXT = DocTypeChal
        ElseIf mVType = TrfChalType Then
            txt(DocType).TEXT = DocTypeTrf
        End If
        txt(VDate).TEXT = Master1!V_Date
        LblVPrefix.CAPTION = mID(Master1!DocID, 9, 5)
        txt(SerialNo).TEXT = Master1!V_NO
        txt(CashCr).TEXT = Master1!Cash_Credit
        txt(Party).Tag = Master1!Party_code
        RsParty.MoveFirst
        RsParty.FIND ("Code ='" & txt(Party).Tag & "'")

        If Master1!Cash_Credit = "Cash" Then
            txt(Party).TEXT = Master1!Party_Name
            mPartyType = 0
        Else
            txt(Party).TEXT = IIf(IsNull(Master1!PartyName), "", Master1!PartyName)
            mPartyType = VNull(Master1!Party_Type)
        End If
        txt(Address1).TEXT = XNull(Master1!Address)
        txt(CrAc).Tag = XNull(Master1!CrAc)
        txt(CrAc).TEXT = IIf(IsNull(Master1!CreditAcName), "", Master1!CreditAcName)
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
        txt(DispMode).TEXT = XNull(Master1!Mode_Dispatch)
        txt(Transport).TEXT = XNull(Master1!Transport)
        txt(LRNo).TEXT = XNull(Master1!GR_RR_No)
        txt(LRDate).TEXT = IIf(IsNull(Master1!GR_RR_Date), "", Master1!GR_RR_Date)
        If Master1!Det_Tax = 0 Then
            txt(TaxDet).TEXT = "No"
        ElseIf Master1!Det_Tax = 1 Then
            txt(TaxDet).TEXT = "Yes"
        End If
        txt(CaseNo).TEXT = XNull(Master1!Case_No)
        txt(CaseMark).TEXT = XNull(Master1!Case_Mark)
        
        txt(MRPAmtTB).TEXT = Format(Master1!SprAmt_MRP_TB + Master1!OilAmt_MRP_TB, "0.00")
        txt(MRPAmtTP).TEXT = Format(Master1!SprAmt_MRP_TP + Master1!OilAmt_MRP_TP, "0.00")
        mMRPLubeTB = Master1!OilAmt_MRP_TB
        mMRPLubeTP = Master1!OilAmt_MRP_TP
        txt(SprAmtTB).TEXT = Format(Master1!SprAmt_TB, "0.00")
        txt(SprAmtTP).TEXT = Format(Master1!SprAmt_TP, "0.00")
        txt(OilAmtTB).TEXT = Format(Master1!OilAmt_TB, "0.00")
        txt(OilAmtTP).TEXT = Format(Master1!OilAmt_TP, "0.00")
        txt(DiscPerTB).TEXT = Format(Master1!D_Per_TB, "0.0000")
        txt(DiscAmtTB).TEXT = Format(Master1!D_Amt_TB, "0.00")
        txt(DiscPerTP).TEXT = Format(Master1!D_Per_TP, "0.0000")
        txt(DiscAmtTP).TEXT = Format(Master1!D_Amt_TP, "0.00")
        txt(STotATB).TEXT = Format((Master1!SprAmt_MRP_TB + Master1!OilAmt_MRP_TB + Master1!SprAmt_TB + Master1!OilAmt_TB) - Master1!D_Amt_TB, "0.00")
        txt(STotATP).TEXT = Format((Master1!SprAmt_MRP_TP + Master1!OilAmt_MRP_TP + Master1!SprAmt_TP + Master1!OilAmt_TP) - Master1!D_Amt_TP, "0.00")
'        Txt(Addition).Text = Format(Master1!Addition, "0.00")
        txt(GenSurPer).TEXT = Format(Master1!Gen_Sur_Per, "0.00")
        txt(GenSurAmt).TEXT = Format(Master1!Gen_Sur_Amt, "0.00")
        txt(TransAmt).TEXT = Format(Master1!Trans_Amt, "0.00")
'        Txt(TaxableTot) = Format(Val(Txt(STotATB)) + Val(Txt(Addition)) + Val(Txt(PackCrg)) + Val(Txt(GenSurAmt)) + Val(Txt(TransAmt)), "0.00")
        txt(TaxableTot) = Format(Val(txt(STotATB)) + Val(txt(GenSurAmt)) + Val(txt(TransAmt)), "0.00")
        txt(STaxPer).TEXT = Format(Master1!Tax_Per, "0.00")
        txt(STaxAmt).TEXT = Format(Master1!Tax_Amt, "0.00")
        
        
        txt(SatAmt).TEXT = Format(Master1!SatAmt, "0.00")
        
        txt(TaxSurPer).TEXT = Format(Master1!Tax_Sur_Per, "0.00")
        txt(TaxSurAmt).TEXT = Format(Master1!Tax_Sur_Amt, "0.00")
        txt(PackCrg).TEXT = Format(Master1!Packing, "0.00")
'        Txt(STotB) = Format(Val(Txt(TaxableTot)) + Val(Txt(STaxAmt)) + Val(Txt(TaxSurAmt)), "0.00")
        txt(STotB) = Format(Val(txt(STotATP)) + Val(txt(TaxableTot)) + Val(txt(PackCrg)) + Val(txt(STaxAmt)) + Val(txt(TaxSurAmt)), "0.00")
        txt(TurnOverPer).TEXT = Format(Master1!TOT_Per, "0.00")
        txt(TurnOverAmt).TEXT = Format(Master1!Tot_Amt, "0.00")
        txt(ReSalTaxPer).TEXT = Format(Master1!ReSalTax_Per, "0.00")
        txt(ReSalTaxAmt).TEXT = Format(Master1!ReSalTax_Amt, "0.00")
        txt(SROff).TEXT = Format(Master1!Rounded, "0.00")
'        Txt(NetSprAmt) = Format(Val(Txt(STotB)) + Val(Txt(STotATP)) + Val(Txt(TurnOverAmt)) + Val(Txt(SROff)), "0.00")
        txt(NetSprAmt) = Format(Val(txt(STotB)) + Val(txt(TurnOverAmt)) + Val(txt(SROff)), "0.00")
        txt(NetAmt).TEXT = Format(Master1!Total_Amt, "0.00")
        
        mTBDisAmtMRP = Master1!D_Amt_MRP_TB
        mTPDisAmtMRP = Master1!D_Amt_MRP_TP
        mMRPTax = Master1!Tax_AmtMRP
        mMRPTaxSur = Master1!TaxSur_AmtMRP
        mMRPTOT = Master1!Tot_AmtMrp
        
        Set Rst = New ADODB.Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select P.Part_Name,P.Local_Name,P.Unit,P.MRP,P.MRP_Effect_Dt,P.TB_SRate ,P.TP_SRate ,P.TB_Effect_Dt ,P.Part_Grade, P.Cur_MRP_TBStk, P.Cur_MRP_TPStk, P.Cur_TB_Stk ,P.Cur_TP_Stk ,P.Bin_Loca ,P.High_Pur_Rate ,P.Low_Pur_Rate,Godown.God_Name, " & cTrim(cMID("SP_Stock.Order_DocID", "9", "5")) & " + " & cCStr(cTrim("Right(SP_Stock.Order_DocID,8)")) & " As OrderIDDisp,SP_Stock.* From (SP_Stock Left Join Part P On SP_Stock.Part_No=P.Part_No and P.Div_Code = left(SP_Stock.Docid,1)) Left Join Godown on SP_Stock.Godown=Godown.God_Code Where SP_Stock.DocID='" & Master1!DocID & "'", GCn, adOpenStatic, adLockReadOnly
        FGrid.Redraw = False
        FGrid.Rows = 1
        If Rst.RecordCount > 0 Then
            I = 1
            Do Until Rst.EOF
                FGrid.AddItem ""
                With FGrid
                    .TextMatrix(I, Col_SrNo) = I
                    .TextMatrix(I, Col_PNo) = Rst!Part_No
                    .TextMatrix(I, Col_SONoCode) = XNull(Rst!Order_DocId)
                    .TextMatrix(I, Col_SONo) = IIf(IsNull(Rst!OrderIDDisp), "", Rst!OrderIDDisp)
                    .TextMatrix(I, Col_SOSrNo) = XNull(Rst!Order_Srl_No)
                    .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                    .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 1, "Yes", "No")
                    .TextMatrix(I, Col_Qty) = Format(Rst!Qty_Iss, "0.000")
                    .TextMatrix(I, Col_Rate) = Format(Rst!Rate, "0.0000")
                    .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP_Rate, "0.0000")
                    If Rst!MRP_YN = 1 Then
                        .TextMatrix(I, Col_Amt) = Format((Rst!Qty_Iss * Rst!MRP_Rate), "0.00")
                    Else
                        .TextMatrix(I, Col_Amt) = Format((Rst!Qty_Iss * Rst!Rate), "0.00")
                    End If
                    .TextMatrix(I, Col_DiscPer) = Format(Rst!Disc_Per, "0.0000")
                    .TextMatrix(I, Col_DiscAmt) = Format(Rst!Disc_Amt, "0.00")
                       
                    If PubVATYN = 1 Then
                        .TextMatrix(I, Col_TaxPer) = Format(Rst!TaxPer, "0.0000")
                        .TextMatrix(I, Col_TaxAmt1) = Format(Rst!TaxAmt, "0.00")
                        
                        .TextMatrix(I, Col_SatPer) = Format(Rst!SatPer, "0.0000")
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
            txt(IWDiscTotTB).TEXT = ""
            txt(IWDiscTotTP).TEXT = ""
        End If
        FGrid.Redraw = True
    Else
        BlankText
    End If
    Grid_Hide
    If PubVATYN = 1 Then
        Amt_Cal
        txt(STaxPer).Enabled = False
        txt(STaxAmt).Enabled = False
        txt(TaxSurPer).Enabled = False
    End If
Set Rst = Nothing
Set Master1 = Nothing
txt(STaxPer).Enabled = False
txt(STaxAmt).Enabled = False
txt(TaxSurPer).Enabled = False
txt(TaxSurAmt).Enabled = False
txt(GenSurPer).Enabled = False
txt(GenSurAmt).Enabled = False
txt(DiscAmtTB).Enabled = False
txt(DiscAmtTP).Enabled = False
txt(DiscPerTB).Enabled = False
txt(DiscPerTP).Enabled = False
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
Dim TmpRst As ADODB.Recordset
    Select Case FGrid.Col
        Case Col_SONo
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            TxtGridValid_SONo
        
        Case Col_PNo, Col_PName, Col_LName
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            TxtGridValid_PNo
            'Nra Modi for NDP Stock transfer
            If txt(DocType) = "Stock Transfer" Then
            Set TmpRst = GCn.Execute("Select PurcDisc_Per from Part_DiscFactor Left Join Part on Part.Disc_Factor=Part_DiscFactor.DiscFac_Catg where Part.Part_No='" & RsPart!Code & "'")
            If TmpRst.RecordCount > 0 Then
                  If VNull(TmpRst!PurcDisc_Per) > 0 Then
                       FGrid.TextMatrix(FGrid.Row, Col_Rate) = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) - (Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) * VNull(TmpRst!PurcDisc_Per) / 100), "0.0000")
                  End If
            End If
            End If
            '***************************
        Case Col_Taxable, Col_MRP
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            TxtGridValid_TaxMRP
            
            'Nra Modi for NDP Stock transfer
            If txt(DocType) = "Stock Transfer" Then
            Set TmpRst = GCn.Execute("Select PurcDisc_Per from Part_DiscFactor Left Join Part on Part.Disc_Factor=Part_DiscFactor.DiscFac_Catg where Part.Part_No='" & RsPart!Code & "'")
            If TmpRst.RecordCount > 0 Then
                  If VNull(TmpRst!PurcDisc_Per) > 0 Then
                       FGrid.TextMatrix(FGrid.Row, Col_Rate) = Format(Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) - (Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) * VNull(TmpRst!PurcDisc_Per) / 100), "0.0000")
                  End If
            End If
            End If
            '***************************
        Case Col_Rate, Col_DiscPer, Col_TaxPer
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
                TxtGrid(0).SetFocus: TxtGridLeave = False:  Exit Function
            End If
            Amt_Cal
            If RsGodown.RecordCount > 0 And Trim(FGrid.TextMatrix(FGrid.Row, Col_Godown)) = "" Then
                RsGodown.MoveFirst
                RsGodown.FIND "Code ='" & PubSprCounterGodown & "'"
                FGrid.TextMatrix(FGrid.Row, Col_GodownCode) = RsGodown!Code
                FGrid.TextMatrix(FGrid.Row, Col_Godown) = RsGodown!Name
            End If
        Case Col_PartSrlNo
            FGrid.TextMatrix(FGrid.Row, Col_PartSrlNo) = TxtGrid(Index)
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

'* Used for Calculate the Amount in Grid
Private Sub Amt_Cal()
'Dim I As Integer
'Dim TotItDiscAmtTB As Double, TotItDiscAmtTP As Double
'Dim TotMRPAmtTB As Double, TotMRPAmtTP As Double
'Dim TotSprAmtTB As Double, TotSprAmtTP As Double
'Dim TotOilAmtTB As Double, TotOilAmtTP As Double
'Dim TotItDiscAmtTB As Double, TotItDiscAmtTP As Double
' To Change
'---
'nra modification
'end modi
Dim mAmount As Double, TaxAmt As Double, DisAmt As Double, OrdDisAmt1 As Double
Dim TTaxAmt As Double, mTaxableAmt As Double
If UCase(FGrid.TextMatrix(FGrid.Row, Col_MRP)) = "YES" Then
    FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty))), "0.00")
Else
    FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty))), "0.00")
End If
FGrid.TextMatrix(FGrid.Row, Col_DiscAmt) = Format(((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) * Val(FGrid.TextMatrix(FGrid.Row, Col_DiscPer))) / 100), "0.00")
FGrid.TextMatrix(FGrid.Row, Col_ItemVal) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Amt)) - Val(FGrid.TextMatrix(FGrid.Row, Col_DiscAmt))), "0.00")
 '******************** For Tax in Line File *************************
    If PubVATYN = 1 Then
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

' Used For Updation of Sale Order in case of Edit and Delete
Private Sub UpdateSO()
Dim Rst As ADODB.Recordset, I As Byte
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select order_docid,Order_Srl_No,Qty_iss From SP_Stock Where DocId='" & txt(DocID).TEXT & "'", GCn, adOpenDynamic, adLockOptimistic
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
        txt(Party).TEXT = RsParty!Name
        txt(Party).Tag = RsParty!Code
        txt(Address1).TEXT = RsParty!Add1
        If txt(Transport) = "" Then
            txt(Transport).TEXT = IIf(IsNull(RsParty!Transporter), "", RsParty!Transporter)
        End If
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

Private Sub DGTrans_Click()
    If rsTrans.RecordCount > 0 Then
        txt(Transport).TEXT = rsTrans!Name
    End If
    txt(Transport).SetFocus
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

Private Sub FGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Used in Grid col swapping
If TopCtrl1.TopText2 = "Browse" And FGrid.Col > 0 Then   'And Button = 2 And FGrid.Row = 1
    FGrid.MousePointer = 15
   CutCol = FGrid.Col
End If
End Sub

Private Sub FGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Used in Grid col swapping
If TopCtrl1.TopText2 = "Browse" Then  'And Button = 2 And FGrid.Row = 1
FGrid.MousePointer = 0
      If CutCol > 0 And CutCol <> FGrid.ColSel And FGrid.ColWidth(FGrid.ColSel) > 0 Then
            ResetGrid CutCol, FGrid.ColSel
            CutCol = 0
      End If
End If
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
If PubVATYN = 1 Then
    LBL(22).CAPTION = "V A T   "
    txt(39).Visible = False
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
TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Const: Grid_Ini
    Call Ini_Pub
    For I = 0 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg  '&HDFF4F2
        txt(I).ForeColor = CtrlFColOrg
    Next
    txt(VDate).Tag = PubLoginDate
    ForSiteCode = PubSiteCode
    mVType = DocTypeChal
    LBL(35) = PubForm31Caption
    LBL(36) = PubForm31Caption & " No."
    LBL(25) = pubTOTCaption
    If PubReSaleTaxPer = 0 Then
        LBL(42).Visible = False
        txt(ReSalTaxPer).Visible = False
        txt(ReSalTaxAmt).Visible = False
    End If
    mVType = SalChalType
    
    Set DGPart.DataSource = RsPart

    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
'    RsParty.Open "Select SubCode as Code,Name,Add1,Transporter,Party_Type From SubGroup Where FirmCode='" & PubFirmCode & "' and Nature='Customer' Order by Name", GCn, adOpenDynamic, adLockOptimistic
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Add1,Transporter,Party_Type,City.CityName From ((SubGroup " & _
        "left join City on City.CityCode=SubGroup.CityCode) " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode) " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,SubGroup.Add1,City.CityName from ((SubGroup " & _
        "Left Join City on City.CityCode=SubGroup.CityCode) " & _
        "left join " & FaTable("AcGroup") & " AcGroup on SubGroup.GroupCode=AcGroup.GroupCode) " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) not in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
    Set RsCrAc = New ADODB.Recordset
    RsCrAc.CursorLocation = adUseClient
    RsCrAc.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGCrAc.DataSource = RsCrAc
    
    Set rsForm = New ADODB.Recordset
    rsForm.CursorLocation = adUseClient
    rsForm.Open "Select Form_Code as Code,Form_Desc As Name,Tax_Per,Tax_Sur_Per From TaxForms Where Spare_YN=1 and Trn_Type='Sale' Order by Form_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGForm.DataSource = rsForm

    Set rsForm31 = New ADODB.Recordset
    rsForm31.CursorLocation = adUseClient
    rsForm31.Open "Select Form_Code as Code,Form_Desc As Name From TaxForms Where Spare_YN=1 and Trn_Type='Permit' Order by Form_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGForm31.DataSource = rsForm31

    Set rsTrans = New ADODB.Recordset
    rsTrans.CursorLocation = adUseClient
    rsTrans.Open "Select Distinct Transport as Name From SP_Sale Where Transport<>'' Order By Transport", GCn, adOpenDynamic, adLockOptimistic
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
    
    
     Dim SiteCond As String
     SiteCond = " And V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
         SiteCond = SiteCond & " and  " & cMID("S.DocID", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    If PubMoveRecYn Then
        Set Master = GCn.Execute("Select S.DocID As SearchCode From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' " & SiteCond & " and S.V_Type In ('" & SalChalType & "', '" & TrfChalType & "') " _
            & "  Order by S.V_Date Desc,S.DocID desc")
    Else
        Set Master = GCn.Execute("Select Top 1 S.DocID As SearchCode From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' " & SiteCond & " and S.V_Type In ('" & SalChalType & "', '" & TrfChalType & "') " _
            & "  Order by S.V_Date Desc,S.DocID desc")
    
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

Private Sub Form_Unload(Cancel As Integer)
    Set RsParty = Nothing
    Set RsCrAc = Nothing
    Set rsForm = Nothing
    Set rsForm31 = Nothing
    Set rsTrans = Nothing
    Set RsSONo = Nothing
    Set RsGodown = Nothing
    Set RsPerson = Nothing
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
    
    txt(VDate).TEXT = txt(VDate).Tag
    txt(DocType).TEXT = DocTypeChal
    txt(CashCr).TEXT = "Credit"
    txt(LC).TEXT = "Local"
    txt(TaxDet).TEXT = "Yes"
    txt(ReSalTaxPer) = IIf(PubReSaleTaxPer = 0, "", Format(PubReSaleTaxPer, "0.00"))
    mPartyType = 0
    txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
    txt(DocID).Tag = txt(DocID)
    txt(DocType).SetFocus
    txt(TurnOverPer) = MainLib.TOTCal()
    FGrid.Col = Col_SONo
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
Dim Rst As ADODB.Recordset
    'Check for existance of transactions
'    Set Rst = New ADODB.Recordset
'    Rst.CursorLocation = adUseClient
'    Rst.Open "Select DocId from SP_Stock Where DocId='" & Txt(DocId) & "' And Invoice_DocID<>''", GCn, adOpenDynamic, adLockOptimistic
'    If Rst.RecordCount  > 0 Then
'        MsgBox "Sale Bill Exists of this Dispatch Challan, " & vbCrLf & "Can't Edit the Reocord", vbInformation, "Validation"
'        Exit Sub
'    End If
    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    txt(DocType).Enabled = False
    txt(VDate).Enabled = False
    txt(SerialNo).Enabled = False
    txt(CashCr).Enabled = False
    FGrid.AddItem FGrid.Rows
    'Enable / Disable Text Box if values zero
    DisableEnableFooter txt(MRPAmtTB), txt(MRPAmtTP), txt(SprAmtTB), txt(SprAmtTP), _
            txt(OilAmtTB), txt(OilAmtTP), txt(DiscPerTB), txt(DiscPerTP), _
            txt(DiscAmtTB), txt(DiscAmtTP), txt(STotATB), _
            txt(GenSurPer), txt(GenSurAmt), txt(TaxableTot), _
            txt(STaxPer), txt(STaxAmt), txt(TaxSurPer), txt(TaxSurAmt)
    'EOF enable / disable section
    txt(Party).SetFocus
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim vBook As Variant, mTrans As Boolean, Rst As ADODB.Recordset
Dim LedgAry(1) As LedgRec, mResult As Byte, MsgStr$, mTitle$

    'Check for existance of transactions
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select DocId from SP_Stock Where DocId='" & txt(DocID) & "' And Invoice_DocID<>''", GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount > 0 Then
        MsgBox "Sale Bill Exists of this Dispatch Challan, " & vbCrLf & "Can't Delete the Reocord", vbInformation, "Validation"
        Exit Sub
    End If
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
            If txt(DocType).TEXT = DocTypeTrf Then
                'Unpost Ledger a/c
                mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, txt(DocID))
                If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
                'Unposting of Ledger completed
            End If
            UpdateSO
            UpdStkTableToTable txt(DocID), "+", "I"
            
            GCn.Execute ("Delete From SP_Stock Where DocID='" & txt(DocID) & "'")
            If GCn.Execute("Select CancelYN from SP_Sale where DocID='" & Master!SearchCode & "'").Fields(0).Value = 1 Then
                GCn.Execute ("Delete From SP_Sale Where DocID='" & txt(DocID) & "'")
            Else
                'New Cancel System
                GCn.Execute "Update SP_Sale Set " & _
                    " CancelYN=1, SprAmt_MRP_TB=0, SprAmt_MRP_TP=0,OilAmt_MRP_TB=0,OilAmt_MRP_TP=0,SprAmt_TB=0,SprAmt_TP=0 " & _
                    " ,OilAmt_TB=0,OilAmt_TP=0, D_Per_TB=0, D_Amt_TB=0,D_Per_TP=0,D_Amt_TP=0, Addition=0,Packing=0 " & _
                    " ,Gen_Sur_Per=0, Gen_Sur_Amt=0,Trans_Amt=0,Tax_Per=0, Tax_Amt=0,Tax_Sur_Per=0,Tax_Sur_Amt=0 " & _
                    " ,TOT_Per=0,TOT_Amt=0,Rounded=0, Total_Amt=0,U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & _
                    ", U_AE='D', D_Per_MRP_TB=0,D_Amt_MRP_TB=0, D_Per_MRP_TP =0, D_Amt_MRP_TP=0, Tax_AmtMRP=0 " & _
                    " , TaxSur_AmtMRP= 0, TOT_AmtMRP= 0, ReSalTax_Per=0, ReSalTax_Amt=0 " & _
                    "   Where DocID='" & txt(DocID) & "'"
            End If
            GCnFaS.CommitTrans
            GCn.CommitTrans
            mTrans = False
            Master.Requery
            If Master.RecordCount > 0 Then
                If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
            End If
            RsPart.Requery
            BUTTONS True, Me, Master, 0
            MoveRec
        End If
    Else
        MsgBox "No Records To Delete!", vbInformation, "Information"
    End If
Set Rst = Nothing
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
    SiteCond = " And V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and  " & cMID("S.DocID", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    If PubBackEnd = "A" Then
        GSQL = "Select S.DocId As SearchCode,S.Site_Code,Switch(S.V_Type='" & SalChalType & "','" & DocTypeChal & "',S.V_Type='" & TrfChalType & "','" & DocTypeTrf & "') As VType, " & cTrim(cMID("S.DocID", "9", "5")) & " As VPrefix, " & cCStr("S.V_No", 10) & " As V_No, " & cDt("S.V_Date") & " AS VDate,S.Cash_Credit, S.Party_Name as PartyName FROM SP_Sale S Where  left(S.DocID,1)='" & PubDivCode & "' " & SiteCond & " and S.V_Type In ('" & SalChalType & "','" & TrfChalType & "') Order by S.V_Date Desc,S.V_Type"
    ElseIf PubBackEnd = "S" Then
        GSQL = "Select S.DocId As SearchCode,S.Site_Code, Case When S.V_Type='" & SalChalType & "' Then '" & DocTypeChal & "' When S.V_Type='" & TrfChalType & "' Then '" & DocTypeTrf & "' End  As VType, " & cTrim(cMID("S.DocID", "9", "5")) & " As VPrefix, " & cCStr("S.V_No", 10) & " As V_No, " & cDt("S.V_Date") & " AS VDate,S.Cash_Credit, S.Party_Name as PartyName FROM SP_Sale S Where  left(S.DocID,1)='" & PubDivCode & "' " & SiteCond & " and S.V_Type In ('" & SalChalType & "','" & TrfChalType & "') Order by S.V_Date Desc,S.V_Type"
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
Dim MyGPNo$, MyGPDate$, mCrLimit As Double
Dim DocIdHlp$, mCurrBal As Double, mEditValue As Double
Dim LedgAry(1) As LedgRec, mNarr$, mResult As Byte, mGatePassOnSprInv$

On Error GoTo errlbl
    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    If IsValid(txt(DocType), "Document Type") = False Then Exit Sub
    If IsValid(txt(VDate), "Date") = False Then Exit Sub
    If IsValid(txt(SerialNo), "Serial Number") = False Then Exit Sub
    If IsValid(txt(CashCr), "Cash/Credit") = False Then Exit Sub
    If IsValid(txt(Party), "Party Name") = False Then Exit Sub
    If txt(DocType).TEXT = DocTypeTrf Then If IsValid(txt(CrAc), "Credit A/c") = False Then Exit Sub
    
    
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            If FGrid.TextMatrix(I, Col_MRP) = "" Then MsgBox "Please Specify MRP Yes/No in S.No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_MRP: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Col_Taxable) = "" Then MsgBox "Please Specify Taxable Yes/No in S.No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Taxable: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, Col_Qty)) = 0 Then MsgBox "Please Specify Quantity in S.No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Qty: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, Col_Godown) = "" Then MsgBox "Please Godown in S.No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Godown: FGrid.SetFocus: Exit Sub
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
    SprMrp FGrid, mMRevDisTBPer, mMRevDisTPPer, Col_PNo, Col_MRP, Col_Taxable, _
            Col_Qty, Col_Rate, Col_MRPRate, Col_DiscAmt, _
            Val(txt(DiscPerTB)), Val(txt(DiscPerTP)), _
            Val(txt(STaxPer)), Val(txt(TaxSurPer)), Val(txt(TurnOverPer))
    
   ' SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
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
    'Check Cr Limit for Challans
    If PubCrLimitCheck = 1 Then
        mCurrBal = 0
        mEditValue = 0
        mCurrBal = GCn.Execute("Select Curr_Bal from SubGroup where SubCode='" & txt(Party).Tag & "'").Fields(0).Value
        mCrLimit = GCn.Execute("Select CreditLimit from SubGroup where SubCode='" & txt(Party).Tag & "'").Fields(0).Value
        If TopCtrl1.TopText2 <> "Add" Then
            mEditValue = GCn.Execute("Select Total_Amt from SP_Sale S Where S.DocID = '" & Master!SearchCode & "'").Fields(0).Value
        End If
        mCurrBal = mCurrBal - mEditValue + Val(txt(NetAmt))
        If mCurrBal > 0 Then     'Dr Balance
            If mCurrBal > mCrLimit Then
                MsgBox "Cr Limit Rs." & mCrLimit & " Exceeds by Rs." & mCurrBal - mCrLimit & vbCrLf & "Add/Edit Denied !", vbInformation, "Cr Limit Checking"
                Me.ActiveControl.SetFocus: Exit Sub
            End If
        End If
    End If
    'EOF Cr Limit Checking
    GCn.BeginTrans
    GCnFaS.BeginTrans
    mTrans = True
    
    If TopCtrl1.TopText2 = "Add" Then
        'lp 12-03-03
        txt(DocID).Tag = txt(DocID)
        If GCn.Execute("Select Count(*) From SP_Sale Where Left(DocID,1)='" & PubDivCode & "' And V_Type = '" & mVType & "' And V_No = " & Val(txt(SerialNo)) & " ").Fields(0) > 0 Then
            If VoucherEditFlag Then
                MsgBox "Document No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                GoTo errlbl
            Else
                txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(txt(DocID).Tag, Document_No)) Then
                    MsgBox "Document No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo errlbl
                End If
            End If
        End If
        ' For Gate Pass
        mGatePassOnSprInv = GCn.Execute("Select GatePassOnSprInv from Syctrl").Fields(0).Value
        If mGatePassOnSprInv = 0 Then
            MyGPNo = "00000" & GCn.Execute("select " & vIsNull("max(" & cVal("Right(gp_no, 5)") & "))", "0") & "+1 from SP_Sale where left(gp_no,1)='" & PubDivCode & "' AND " & cMID("gp_no", "2", "1") & "='" & PubSiteCode & "' And (Isnull(Job_DocID) or Trim(Job_DocID)='')").Fields(0).Value
            MyGPNo = PubDivCode & PubSiteCode & ForSiteCode & Right(MyGPNo, 5)
            lblGatePass = MyGPNo
            MyGPDate = PubServerDate
        End If
        DocIdHlp = UCase(Replace(txt(DocID), " ", ""))
        '**********
        GCn.Execute "Insert Into SP_Sale(" _
            & "DocID ,DocIDHelp ,V_Type ,V_No ,Site_Code ," _
            & "V_Date,Cash_Credit ,Party_Code ,Party_Name ,Address ," _
            & "L_C ,Form_Code ,RoadPermit_FormCode ,RoadPermit_No ,GR_RR_No ," _
            & "GR_RR_Date ,CrAc ,Case_No ,Case_Mark ,Mode_Dispatch ," _
            & "Transport ,Rep_Code ,Remarks ,Det_Tax ,SprAmt_MRP_TB ," _
            & "SprAmt_MRP_TP ,OilAmt_MRP_TB,OilAmt_MRP_TP,SprAmt_TB ,SprAmt_TP ,OilAmt_TB ,OilAmt_TP ," _
            & "D_Per_TB ,D_Amt_TB ,D_Per_TP ,D_Amt_TP ,Addition ," _
            & "Packing ,Gen_Sur_Per ,Gen_Sur_Amt ,Trans_Amt ,Tax_Per ," _
            & "Tax_Amt ,Tax_Sur_Per ,Tax_Sur_Amt ,TOT_Per ,TOT_Amt ," _
            & "ReSalTax_Per, ReSalTax_Amt,Rounded ,Total_Amt ,U_Name ,U_EntDt ,U_AE, AddBy, AddDate, " _
            & "D_Per_MRP_TB,D_Amt_MRP_TB, D_Per_MRP_TP , D_Amt_MRP_TP, Tax_AmtMRP, TaxSur_AmtMRP, TOT_AmtMRP,GP_No,GP_Date, SatAmt, Sat_Yn ) " _
            & "Values('" & txt(DocID) & "','" & DocIdHlp & "','" & mVType & "'," & txt(SerialNo) & ",'" & PubSiteCode & PubSiteCode & _
            "'," & ConvertDate(Format(txt(VDate), "dd/MMM/yyyy")) & ",'" & txt(CashCr) & "','" & txt(Party).Tag & "','" & txt(Party) & "','" & txt(Address1) & _
            "','" & left(txt(LC), 1) & "','" & txt(FormName).Tag & "','" & txt(Form31Name).Tag & "','" & txt(Form31No) & "','" & txt(LRNo) & _
            "'," & ConvertDate(txt(LRDate)) & ",'" & txt(CrAc).Tag & "'," & Val(txt(CaseNo)) & ",'" & txt(CaseMark) & "','" & txt(DispMode) & _
            "','" & txt(Transport) & "','" & txt(SPerson).Tag & "','" & txt(Remark) & "'," & IIf(txt(TaxDet) = "Yes", 1, 0) & "," & Val(txt(MRPAmtTB)) - mMRPLubeTB & _
            " ," & Val(txt(MRPAmtTP)) - mMRPLubeTP & "," & mMRPLubeTB & "," & mMRPLubeTP & "," & Val(txt(SprAmtTB)) & "," & Val(txt(SprAmtTP)) & "," & Val(txt(OilAmtTB)) & "," & Val(txt(OilAmtTP)) & _
            " ," & Val(txt(DiscPerTB)) & "," & Val(txt(DiscAmtTB)) & "," & Val(txt(DiscPerTP)) & "," & Val(txt(DiscAmtTP)) & "," & Val(txt(Addition)) & _
            " ," & Val(txt(PackCrg)) & "," & Val(txt(GenSurPer)) & "," & Val(txt(GenSurAmt)) & "," & Val(txt(TransAmt)) & "," & Val(txt(STaxPer)) & _
            " ," & Val(txt(STaxAmt)) & "," & Val(txt(TaxSurPer)) & "," & Val(txt(TaxSurAmt)) & "," & Val(txt(TurnOverPer)) & "," & Val(txt(TurnOverAmt)) & _
            " ," & Val(txt(ReSalTaxPer)) & "," & Val(txt(ReSalTaxAmt)) & "," & Val(txt(SROff)) & "," & Val(txt(NetAmt)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & _
            ",'A', '" & pubUName & "', " & ConvertDateTime(PubServerDate) & "," & mMRevDisTBPer & "," & mTBDisAmtMRP & "," & mMRevDisTPPer & "," & mTPDisAmtMRP & "," & mMRPTax & "," & mMRPTaxSur & ", " & mMRPTOT & ", '" & MyGPNo & "'," & ConvertDate(MyGPDate) & ", " & Val(txt(SatAmt)) & ", " & IIf(mSatYn, 1, 0) & ")"
        'Voucher Serial No. Updation LPS 21-05-03
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaS, txt(DocID), txt(VDate)
    Else
        UpdateSO
        UpdStkTableToTable txt(DocID), "+", "I"
        GCn.Execute ("Delete From SP_Stock Where DocID='" & txt(DocID) & "'")

        GCn.Execute "Update SP_Sale Set " _
            & " Cash_Credit='" & txt(CashCr) & "',Party_Code='" & txt(Party).Tag & "', Party_Name='" & txt(Party) & "',Address='" & txt(Address1) & _
            "', L_C='" & left(txt(LC), 1) & "',Form_Code='" & txt(FormName).Tag & "',RoadPermit_FormCode='" & txt(Form31Name).Tag & "',RoadPermit_No='" & txt(Form31No) & _
            "', GR_RR_No='" & txt(LRNo) & "',GR_RR_Date=" & ConvertDate(txt(LRDate)) & ",CrAc='" & txt(CrAc).Tag & "',Case_No=" & Val(txt(CaseNo)) & _
            " , Case_Mark='" & txt(CaseMark) & "',Mode_Dispatch='" & txt(DispMode) & "',Transport='" & txt(Transport) & "',Rep_Code='" & txt(SPerson).Tag & _
            "', Remarks='" & txt(Remark) & "',Det_Tax=" & IIf(txt(TaxDet) = "Yes", 1, 0) & ",SprAmt_MRP_TB=" & Val(txt(MRPAmtTB)) - mMRPLubeTB & _
            " , SprAmt_MRP_TP=" & Val(txt(MRPAmtTP)) - mMRPLubeTP & ",OilAmt_MRP_TB=" & mMRPLubeTB & ",OilAmt_MRP_TP=" & mMRPLubeTP & ",SprAmt_TB=" & Val(txt(SprAmtTB)) & ",SprAmt_TP=" & Val(txt(SprAmtTP)) & _
            " , OilAmt_TB=" & Val(txt(OilAmtTB)) & ",OilAmt_TP=" & Val(txt(OilAmtTP)) & ", D_Per_TB=" & Val(txt(DiscPerTB)) & _
            " , D_Amt_TB=" & Val(txt(DiscAmtTB)) & ",D_Per_TP=" & Val(txt(DiscPerTP)) & ",D_Amt_TP=" & Val(txt(DiscAmtTP)) & _
            " , Addition=" & Val(txt(Addition)) & ",Packing=" & Val(txt(PackCrg)) & ", Gen_Sur_Per=" & Val(txt(GenSurPer)) & _
            " , Gen_Sur_Amt=" & Val(txt(GenSurAmt)) & ",Trans_Amt=" & Val(txt(TransAmt)) & ",Tax_Per=" & Val(txt(STaxPer)) & _
            " , Tax_Amt=" & Val(txt(STaxAmt)) & ",Tax_Sur_Per=" & Val(txt(TaxSurPer)) & ",Tax_Sur_Amt=" & Val(txt(TaxSurAmt)) & _
            " , TOT_Per=" & Val(txt(TurnOverPer)) & ",TOT_Amt=" & Val(txt(TurnOverAmt)) & ",Rounded=" & Val(txt(SROff)) & _
            " , Total_Amt=" & Val(txt(NetAmt)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & _
            ", U_AE='E', ModifyBy = '" & pubUName & "', ModifyDate = " & ConvertDateTime(PubServerDate) & ", D_Per_MRP_TB=" & mMRevDisTBPer & ",D_Amt_MRP_TB=" & mTBDisAmtMRP & ", D_Per_MRP_TP =" & mMRevDisTPPer & _
            " , D_Amt_MRP_TP=" & mTPDisAmtMRP & ", Tax_AmtMRP=" & mMRPTax & ", TaxSur_AmtMRP= " & mMRPTaxSur & ", TOT_AmtMRP= " & mMRPTOT & _
            " , ReSalTax_Per=" & Val(txt(ReSalTaxPer)) & ", SatAmt = " & Val(txt(SatAmt)) & ", Sat_Yn = " & IIf(mSatYn, 1, 0) & ", ReSalTax_Amt=" & Val(txt(ReSalTaxAmt)) & _
            "   Where DocID='" & txt(DocID) & "'"
    End If
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            If FGrid.TextMatrix(I, Col_SONo) <> "" Then
                GCn.Execute "Update SP_Order1 Set Sup_Qty=Sup_Qty+" & Val(FGrid.TextMatrix(I, Col_Qty)) & " Where OrderId='" & FGrid.TextMatrix(I, Col_SONoCode) & "' and Srl_No=" & FGrid.TextMatrix(I, Col_SOSrNo) & ""
            End If
            GCn.Execute "Insert Into SP_Stock(" _
                & "DocID,Srl_No,V_Type,V_No,V_Date,Site_Code," _
                & "Party_Code,L_C,Order_DocId,Order_Srl_No,Part_No," _
                & "Godown,Qty_Iss,Tax_YN,MRP_YN,Rate," _
                & "MRP_Rate,Disc_Per,Disc_Amt,Amount,Net_Amt,V_Rate," _
                & "Part_SrlNo,U_Name,U_EntDt,U_AE,TaxPer,TaxAmt, SatPer, SatAmt) " _
                & "Values('" & txt(DocID) & "'," & I & ",'" & mVType & "'," & txt(SerialNo) & "," & ConvertDate(Format(txt(VDate).TEXT, "dd/MMM/yyyy")) & ",'" & PubSiteCode & PubSiteCode & _
                "','" & txt(Party).Tag & "','" & left(txt(LC).TEXT, 1) & "','" & FGrid.TextMatrix(I, Col_SONoCode) & "'," & Val(FGrid.TextMatrix(I, Col_SOSrNo)) & ",'" & FGrid.TextMatrix(I, Col_PNo) & _
                "','" & FGrid.TextMatrix(I, Col_GodownCode) & "'," & Val(FGrid.TextMatrix(I, Col_Qty)) & "," & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & "," & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & "," & Val(FGrid.TextMatrix(I, Col_Rate)) & _
                " , " & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," & Val(FGrid.TextMatrix(I, Col_DiscPer)) & "," & Val(FGrid.TextMatrix(I, Col_DiscAmt)) & "," & Val(FGrid.TextMatrix(I, Col_Amt)) & "," & Val(FGrid.TextMatrix(I, Col_ItemVal)) & ", " & Val(FGrid.TextMatrix(I, Col_Amt)) & _
                " ,'" & FGrid.TextMatrix(I, Col_PartSrlNo) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "'," & Val(FGrid.TextMatrix(I, Col_TaxPer)) & "," & Val(FGrid.TextMatrix(I, Col_TaxAmt1)) & "," & Val(FGrid.TextMatrix(I, Col_SatPer)) & "," & Val(FGrid.TextMatrix(I, Col_SatAmt1)) & ")"
            Call UpdStkGridToTable(FGrid.TextMatrix(I, Col_PNo), "-", FGrid.TextMatrix(I, Col_MRP), FGrid.TextMatrix(I, Col_Taxable), FGrid.TextMatrix(I, Col_Qty))
        End If
    Next
        'A/c Posting for Transfer Case
        If txt(DocType).TEXT = DocTypeTrf Then
            mNarr = "Through Stock Transfer Issue"
            I = 0
            LedgAry(I).SubCode = txt(Party).Tag
            LedgAry(I).AmtDr = Val(txt(NetAmt))
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = txt(CrAc).Tag
            I = I + 1
            LedgAry(I).SubCode = txt(CrAc).Tag
            LedgAry(I).AmtCr = Val(txt(NetAmt))
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = txt(Party).Tag
            
            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, txt(DocID), CDate(txt(VDate)), mNarr & "[Common]")
            If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
        End If
        'EOF Posting
    GCnFaS.CommitTrans
    GCn.CommitTrans
    mTrans = False
    mSearchCode = txt(DocID)
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select S.DocID As SearchCode From SP_Sale S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' and S.V_Type In ('" & SalChalType & "', '" & TrfChalType & "') And S.DocID  = '" & mSearchCode & "' " _
            & "  Order by S.V_Date Desc,S.DocID desc")
    End If
    Master.FIND "SearchCode = '" & mSearchCode & "'"
    'lp 12-03-03
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(SerialNo)) > Val(DeCodeDocID(txt(DocID).Tag, Document_No)) Then
            MsgBox "Document No." & Trim(DeCodeDocID(txt(DocID).Tag, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
    End If
    TopCtrl1_ePrn
Exit Sub
errlbl:
    If mTrans Then GCnFaS.RollbackTrans: GCn.RollbackTrans
    CheckError
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
    Else
        Me.ActiveControl.SetFocus
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
            ListArray = Array(DocTypeChal, DocTypeTrf)
            Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
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
'LPS 24-04-02
        Case FormName
            If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).TEXT <> rsForm!Name Then
                rsForm.MoveFirst
                rsForm.FIND "Name ='" & txt(Index).TEXT & "'"
            End If
        Case Form31Name
            If rsForm31.RecordCount = 0 Or (rsForm31.EOF = True Or rsForm31.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).TEXT <> rsForm31!Name Then
                rsForm31.MoveFirst
                rsForm31.FIND "Name ='" & txt(Index).TEXT & "'"
            End If
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
            ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
        Case SerialNo
            NumDown txt(Index), KeyCode, 8, 0
        Case LC
            ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
        Case Party
            If txt(CashCr).TEXT = "Credit" Then
                DGridTxtKeyDown DGParty, txt, Party, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
            End If
        Case CrAc
            If txt(CashCr).TEXT = "Credit" Or txt(DocType).TEXT = DocTypeTrf Then
                DGridTxtKeyDown DGCrAc, txt, CrAc, RsCrAc, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
            End If
        Case SPerson
            DGridTxtKeyDown DGPerson, txt, SPerson, RsPerson, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
        Case FormName
            DGridTxtKeyDown DGForm, txt, FormName, rsForm, KeyCode, False, 1
        Case Form31Name
            DGridTxtKeyDown DGForm31, txt, Form31Name, rsForm31, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
        Case Transport
            DGridTxtKeyDown_Mast DGTrans, txt, Transport, rsTrans, KeyCode, False, 0
        Case CaseNo
            NumDown txt(Index), KeyCode, 8, 0
        Case DiscAmtTB, DiscAmtTP, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, PackCrg, TurnOverAmt, ReSalTaxAmt
            NumDown txt(Index), KeyCode, 8, 2
        Case GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ReSalTaxPer
            NumDown txt(Index), KeyCode, 2, 2
        Case DiscPerTB, DiscPerTP
            NumDown txt(Index), KeyCode, 2, 4
    End Select
    If FrmList.Visible = False And DGParty.Visible = False And DGCrAc.Visible = False And DGForm.Visible = False And DGForm31.Visible = False And DGPerson.Visible = False Then
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
        If txt(CashCr).TEXT = "Credit" Or txt(DocType).TEXT = DocTypeTrf Then
            If DGCrAc.Visible = True Then DGridTxtKeyPress txt, CrAc, RsCrAc, KeyAscii, "Name"
        End If
    Case SPerson
        If DGPerson.Visible = True Then DGridTxtKeyPress txt, SPerson, RsPerson, KeyAscii, "Name"
    Case FormName
        If DGForm.Visible = True Then DGridTxtKeyPress txt, FormName, rsForm, KeyAscii, "Name"
    Case Form31Name
        If DGForm31.Visible = True Then DGridTxtKeyPress txt, Form31Name, rsForm31, KeyAscii, "Name"
'    Case LC
'        If KeyAscii = 76 Or KeyAscii = 108 Or KeyAscii = 82 Or KeyAscii = 114 Then
'            If KeyAscii = 76 Or KeyAscii = 108 Then             ' L/l
'                txt(Index).Text = "Local"
'                KeyAscii = 0
'            ElseIf KeyAscii = 67 Or KeyAscii = 99 Then          ' C/c
'                txt(Index).Text = "Central"
'                KeyAscii = 0
'            End If
'        Else
'            KeyAscii = 0
'        End If
    Case CaseNo
        NumPress txt(Index), KeyAscii, 8, 0
    Case DiscAmtTB, DiscAmtTP, GenSurAmt, TransAmt, STaxAmt, TaxSurAmt, PackCrg, TurnOverAmt, ReSalTaxAmt
        NumPress txt(Index), KeyAscii, 8, 2
    Case GenSurPer, STaxPer, TaxSurPer, TurnOverPer, ReSalTaxPer
        NumPress txt(Index), KeyAscii, 2, 2
    Case DiscPerTB, DiscPerTP
        NumPress txt(Index), KeyAscii, 2, 4
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
        Case CashCr
            If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
        Case LC
            If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
        Case Transport
            If DGTrans.Visible = True Then DGridTxtKeyUp_Mast txt, Transport, rsTrans, KeyCode, "Name"
        Case DiscPerTB, DiscAmtTB, DiscPerTP, DiscAmtTP, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, PackCrg, TurnOverPer, TurnOverAmt, ReSalTaxPer, ReSalTaxAmt
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
            'MainLib.SprCalc NoLabour, FGrid, mMRevDisTBPer, mMRevDisTPPer, mTBDisAmtMRP, mTPDisAmtMRP, _
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
            If Not Trim(txt(Index).TEXT) <> DocTypeChal Or Trim(txt(Index).TEXT) <> DocTypeTrf Then
                txt(Index).TEXT = DocTypeChal
            End If
            If Trim(txt(Index).TEXT) = DocTypeChal Then
                txt(CrAc).Enabled = False
                txt(CrAc).TEXT = ""
                mVType = SalChalType
            ElseIf Trim(txt(Index).TEXT) = DocTypeTrf Then
                txt(CrAc).Enabled = True
                mVType = TrfChalType
            End If
            txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            txt(DocID).Tag = txt(DocID)
        Case VDate
            txt(Index).TEXT = RetDate(txt(Index))
            Cancel = Not CheckFinYear(txt(Index))
            If Cancel = False Then
                txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                txt(DocID).Tag = txt(DocID)
            End If
        Case SerialNo
            If VoucherEditFlag = True Then      ' Manual
                txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                txt(DocID).Tag = txt(DocID)
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select V_No From SP_Sale Where DocID='" & txt(DocID).TEXT & "'", GCn, adOpenStatic, adLockReadOnly
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Document No. Not Allowed", vbInformation, "Validation"
                    Cancel = True
                    txt(SerialNo).SetFocus
                End If
            End If
        Case LC ',CashCr
            txt(Index).TEXT = ListView.SelectedItem.TEXT
        Case Party
            If Trim(txt(Index).TEXT = "") Then
                MsgBox "Please Select Party", vbInformation, "Information"
                txt(Index).SetFocus
                Cancel = True
                Exit Sub
            End If
            ' To Populate Sale Orders Data Grid for selected Customer
            If txt(CashCr).TEXT = "Credit" Then
                If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then
                    txt(Index).TEXT = ""
                    txt(Index).Tag = ""
                    txt(Address1).TEXT = ""
                    mPartyType = 0
                Else
                    txt(Index).TEXT = RsParty!Name
                    txt(Index).Tag = RsParty!Code
                    txt(Address1).TEXT = RsParty!Add1
                    If txt(Transport) = "" Then
                        txt(Transport).TEXT = IIf(IsNull(RsParty!Transporter), "", RsParty!Transporter)
                    End If
                    mPartyType = VNull(RsParty!Party_Type)
                    GSQL = "Select OrderID as Code, " & cTrim(cMID("OrderID", "9", "5")) & "+" & cCStr(cTrim("Right(OrderID,8)")) & " as Name,V_Date From SP_Order Where left(OrderID,1)='" & PubDivCode & "' and Order_Type='S_SO' and Party_Code='" & txt(Party).Tag & "'  and V_Date<=" & ConvertDate(Format(txt(VDate), "dd-mmm-yyyy")) & " and OrdClosDate is null Order By OrderID"
                End If
            Else
                txt(Party).Tag = PubSprCashAc
                mPartyType = 0
                GSQL = "Select OrderID as Code, " & cTrim(cMID("OrderID", "9", "5")) & "+ " & cCStr(cTrim("Right(OrderID,8)")) & " as Name,V_Date From SP_Order Where left(OrderID,1)='" & PubDivCode & "' and Order_Type='S_SO' and V_Date<=" & ConvertDate(Format(txt(VDate), "dd-mmm-yyyy")) & " and OrdClosDate is null Order By OrderID"
            End If
            If Trim(txt(DocType).TEXT) = DocTypeChal And GSQL <> "" Then
                Set RsSONo = New ADODB.Recordset
                RsSONo.CursorLocation = adUseClient
                RsSONo.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
                Set DGSONo.DataSource = RsSONo
            End If
        Case FormName
            If rsForm.RecordCount > 0 Or (rsForm.EOF = False Or rsForm.BOF = False) Then
                If txt(Index).TEXT <> "" Then
                    txt(Index).TEXT = rsForm!Name
                    txt(Index).Tag = rsForm!Code
                    If TopCtrl1.TopText2.CAPTION = "Add" Then   ' To Assign Tax% in case of Add
                        txt(STaxPer).TEXT = rsForm!Tax_Per
                        txt(TaxSurPer).TEXT = rsForm!Tax_Sur_Per
                        Amt_Cal
                    End If
                End If
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        Case Form31Name
            If rsForm31.RecordCount > 0 Or (rsForm31.EOF = False Or rsForm31.BOF = False) Then
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
        Case LRDate
            txt(Index).TEXT = RetDate(txt(Index))
        Case CaseNo, DiscAmtTB, DiscAmtTP, GenSurPer, GenSurAmt, TransAmt, STaxPer, STaxAmt, TaxSurPer, TaxSurAmt, PackCrg, TurnOverPer, TurnOverAmt, ReSalTaxPer, ReSalTaxAmt, SROff
            If Index <> CaseNo Then
                txt(Index).TEXT = IIf(txt(Index).TEXT <> "", Format(txt(Index), "0.00"), "")
            End If
        Case DiscPerTB, DiscPerTP
            txt(Index).TEXT = IIf(txt(Index).TEXT <> "", Format(txt(Index), "0.0000"), "")
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
    If FrmDetail.Visible = False Then FrmDetail.Visible = True
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    TxtGrid(Index).MaxLength = 0
    Select Case FGrid.Col
        Case Col_SONo
            If Trim(txt(DocType).TEXT) = DocTypeChal Then
                If RsSONo.RecordCount = 0 Or (RsSONo.EOF = True Or RsSONo.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Col_SONo) = "" Then Exit Sub
                If FGrid.TextMatrix(FGrid.Row, Col_SONoCode) <> RsSONo!Code Then
                    RsSONo.MoveFirst
                    RsSONo.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, Col_SONo) & "'"
                End If
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
            If Trim(txt(DocType).TEXT) = DocTypeChal Then
                DGridTxtKeyDown DGSONo, TxtGrid, Index, RsSONo, KeyCode, True, 1
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave = True Then
                       GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, 1  ' 18, 2
                    End If
                End If
            End If
        Case Col_PNo
            If DGPart.Visible = False Then DGridColSwap DGPart, 0
            DGridTxtKeyDown DGPart, TxtGrid, 0, RsPart, KeyCode, True, 0, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, 2 ' 18, 1
                End If
            End If
        Case Col_MRP, Col_Taxable, Col_Qty, Col_PartSrlNo
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo ' 18
                End If
            End If
            
            
        Case Col_TaxPer
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, , Col_Godown ' 18
                End If
            End If
            
            
        Case Col_DiscPer
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, , Col_Godown
                End If
            End If
            
'        Case Col_Qty
'            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
'                If TxtGridLeave = True Then
'                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Godown ' 18
'                End If
'            End If
        Case Col_Rate
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, 2 ' 18, 2
                End If
            End If
        Case Col_DiscAmt
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                    If PubRestrict_Godown = 1 Then      ' Restrict Godown is "YES"
                        'Purpose not Clear, Redesign
                        GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Col_PartSrlNo
                    Else
                        GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo, , Col_Godown ' 18, 2
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
                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_PartSrlNo ' 18
                End If
            End If
        Case Col_LName
            If DGPart.Visible = False Then DGridColSwap DGPart, 6
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 6, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Godown, 1 ' 18, 1
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
    Case Col_SONo And Trim(txt(DocType).TEXT) = DocTypeChal
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
    Case Col_DiscPer, Col_TaxPer
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
    Select Case FGrid.Col
        Case Col_SONo And Trim(txt(DocType).TEXT) = DocTypeChal
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
        Case Col_MRP
            If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
                TxtGrid(Index) = ""
            ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Then
                TxtGrid(Index) = "Yes"
            Else
                TxtGrid(Index) = "No"
            End If
            
        Case Col_Taxable
            If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
                TxtGrid(Index) = ""
            ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Or Trim(TxtGrid(Index)) = "" Then
                TxtGrid(Index) = "Yes"
            Else
                TxtGrid(Index) = "No"
            End If
            
        Case Col_DiscPer, Col_TaxPer
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.0000")
        Case Col_DiscAmt
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        Case Col_Rate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.0000")
        Case Col_Qty
            FGrid.TextMatrix(FGrid.Row, Col_Qty) = Format(Val(TxtGrid(Index).TEXT), "0.000")
            CountItem
        Case Col_PartSrlNo
            FGrid.TextMatrix(FGrid.Row, Col_PartSrlNo) = TxtGrid(Index)
    End Select
    Amt_Cal
    If KeyCode = vbKeyEscape Then
        FGrid.SetFocus
        TxtGrid(0).Visible = False
        Grid_Hide
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
        Select Case FGrid.Col
            Case Col_SONo
                If txt(DocType) <> DocTypeTrf Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                    Amt_Cal
                End If
            Case Col_MRP, Col_Taxable, Col_PartSrlNo
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            Case Col_Qty, Col_Rate, Col_Amt, Col_DiscPer, Col_DiscAmt, Col_TaxPer
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                Amt_Cal
            Case Col_Godown
                If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                End If
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case FGrid.Col
            Case Col_SONo
                If txt(DocType) <> DocTypeTrf Then
                    GridDblClick Me, FGrid, TxtGrid, 0
                    TAddMode = False
                End If
            Case Col_PNo, Col_PName, Col_LName
                GridDblClick Me, FGrid, TxtGrid, 0
                TAddMode = False
            Case Col_MRP, Col_Taxable, Col_Qty, Col_Rate, Col_DiscPer, Col_DiscAmt, Col_PartSrlNo
                If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                    GridDblClick Me, FGrid, TxtGrid, 0
                    TAddMode = False
                End If
'by lps 30-04-02
'            Case Col_DiscAmt
'                GridDblClick Me, FGrid, TxtGrid, 0
            Case Col_Godown
                If FGrid.TextMatrix(FGrid.Row, Col_Qty) <> "" Then
                    If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
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
    Select Case FGrid.Col
        Case Col_SONo
            If txt(DocType) <> DocTypeTrf Then
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            End If
        Case Col_PNo, Col_PName, Col_LName
            Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        Case Col_Unit
            FGrid.Col = FGrid.Col + 1
            FGrid.SetFocus
        Case Col_PartSrlNo, Col_MRP, Col_Taxable
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            End If
        Case Col_Qty, Col_Rate, Col_DiscPer, Col_DiscAmt, Col_TaxPer
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
            MsgBox "No Entries To Delete!", vbCritical, "Delete Module"
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
    FGrid.TextMatrix(FGrid.Row, Col_LName) = IIf(IsNull(RsPart!LName), "", RsPart!LName)
    
    MainLib.Fill_Data mPartyType, LblFrm, FGrid, _
        RsPart!Code, RsPart!Name, IIf(IsNull(RsPart!LName), "", RsPart!LName), _
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
            FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(mPartyType, FGrid, CDate(txt(VDate)), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
'            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsPart!SalDisc_Per, "0.00")
        End If
    End If
End If
If PubVATYN = 1 Then
    Set rsTaxPer = GCn.Execute("Select Tax_Per, AddTaxPer, L_C from TaxForms where Form_Code='" & txt(FormName).Tag & "'")
    If rsTaxPer.RecordCount > 0 Then
        FGrid.TextMatrix(FGrid.Row, Col_TaxPer) = rsTaxPer!Tax_Per
        FGrid.TextMatrix(FGrid.Row, Col_SatPer) = VNull(rsTaxPer!AddTaxPer)
        
        If UTrim(XNull(rsTaxPer!L_C)) = "LOCAL" Then
           Set rsTaxPer = GCn.Execute("Select VatPer, AddTaxPer From Part_Grade Where PartGrade_Code = '" & FGrid.TextMatrix(FGrid.Row, Col_PartGrade) & "'")
           If rsTaxPer.RecordCount > 0 Then
               If VNull(rsTaxPer!VatPer) > 0 Then FGrid.TextMatrix(FGrid.Row, Col_TaxPer) = Format(rsTaxPer!VatPer, "0.00")
               If VNull(rsTaxPer(0)) > 0 Then FGrid.TextMatrix(FGrid.Row, Col_SatPer) = Format(VNull(rsTaxPer!AddTaxPer), "0.00")
           End If
        End If
        
    End If
End If
'SO Itme fill
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
'        If TopCtrl1.TopText2 = "Add" Or _
            TopCtrl1.TopText2 = "Edit" And Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) = 0 Then
            FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(mPartyType, FGrid, CDate(txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
'        End If
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

Private Sub Ini_Const()
  Col_SrNo = 0               ' Serial No
  Col_SONo = 1               ' Sale Order No Name
  Col_SOSrNo = 2             ' Sale Order Serial No
  Col_PNo = 3                ' Part No
  Col_SONoCode = 4           ' Sale Order No Code
  Col_Unit = 5               ' Unit
  Col_MRP = 6                ' MRP Yes/No
  Col_Taxable = 7            ' Taxable Yes/No
  Col_Qty = 8                ' Qty
  Col_Rate = 9               ' Rate
  Col_MRPRate = 10           ' MRP Rate
  Col_Amt = 11               ' Amt
  Col_DiscPer = 12           ' Disc. %
  Col_DiscAmt = 13           ' Disc. Amt.
  Col_TaxPer = 14            ' Tax Per.
  Col_TaxAmt1 = 15           ' Tax Amt.
  Col_SatPer = 16
  Col_SatAmt1 = 17
  Col_ItemVal = 18           ' Item Value
  Col_GodownCode = 19        ' Godown Code
  Col_Godown = 20            ' Godown
  Col_PartSrlNo = 21         ' Part Serial No
  Col_PName = 22             ' Part Name
  Col_LName = 23             ' Local Name
  Col_Bin = 24               ' Bin
  Col_MRPStkTP = 25          ' MRP TP Qty 'Current Stock Qty
  Col_MRPStkTB = 26          ' MRP TB Qty
  Col_TBStk = 27             ' Taxbale Qty
  Col_TPStk = 28             ' Tax Paid Qty
  Col_TBRate = 29            ' Taxbale Rate
  Col_TPRate = 30            ' Tax Paid Rate
  Col_LastRate = 31          ' Last Purchase Rate
  Col_HPRate = 32            ' High Purchase Rate
  Col_LPRate = 33            ' Low Purchase Rate
  Col_PartGrade = 34         ' Part Grade (Used for Oil Item)
  Col_EffectDate = 35        ' MRP Effective Date/TB Effective Date
End Sub

Private Sub ResetGrid(CutCol As Byte, PasteCol As Byte)
'Dim Sign1$, Sign2$, Sign3$, Sign4$
If CutCol < PasteCol Then
'    Sign1 = "-1": Sign2 = " >": Sign3 = "<": Sign4 = "-1"
    If Col_SONo = CutCol Then Col_SONo = PasteCol - 1 Else If Col_SONo > CutCol And Col_SONo < PasteCol Then Col_SONo = Col_SONo - 1
    If Col_SOSrNo = CutCol Then Col_SOSrNo = PasteCol - 1 Else If Col_SOSrNo > CutCol And Col_SOSrNo < PasteCol Then Col_SOSrNo = Col_SOSrNo - 1
    If Col_SONoCode = CutCol Then Col_SONoCode = PasteCol - 1 Else If Col_SONoCode > CutCol And Col_SONoCode < PasteCol Then Col_SONoCode = Col_SONoCode - 1
    If Col_PNo = CutCol Then Col_PNo = PasteCol - 1 Else If Col_PNo > CutCol And Col_PNo < PasteCol Then Col_PNo = Col_PNo - 1
    If Col_Unit = CutCol Then Col_Unit = PasteCol - 1 Else If Col_Unit > CutCol And Col_Unit < PasteCol Then Col_Unit = Col_Unit - 1
    If Col_MRP = CutCol Then Col_MRP = PasteCol - 1 Else If Col_MRP > CutCol And Col_MRP < PasteCol Then Col_MRP = Col_MRP - 1
    If Col_Taxable = CutCol Then Col_Taxable = PasteCol - 1 Else If Col_Taxable > CutCol And Col_Taxable < PasteCol Then Col_Taxable = Col_Taxable - 1
    If Col_Qty = CutCol Then Col_Qty = PasteCol - 1 Else If Col_Qty > CutCol And Col_Qty < PasteCol Then Col_Qty = Col_Qty - 1
    If Col_Rate = CutCol Then Col_Rate = PasteCol - 1 Else If Col_Rate > CutCol And Col_Rate < PasteCol Then Col_Rate = Col_Rate - 1
    If Col_MRPRate = CutCol Then Col_MRPRate = PasteCol - 1 Else If Col_MRPRate > CutCol And Col_MRPRate < PasteCol Then Col_MRPRate = Col_MRPRate - 1
    If Col_Amt = CutCol Then Col_Amt = PasteCol - 1 Else If Col_Amt > CutCol And Col_Amt < PasteCol Then Col_Amt = Col_Amt - 1
    If Col_DiscPer = CutCol Then Col_DiscPer = PasteCol - 1 Else If Col_DiscPer > CutCol And Col_DiscPer < PasteCol Then Col_DiscPer = Col_DiscPer - 1
    If Col_DiscAmt = CutCol Then Col_DiscAmt = PasteCol - 1 Else If Col_DiscAmt > CutCol And Col_DiscAmt < PasteCol Then Col_DiscAmt = Col_DiscAmt - 1
    If Col_ItemVal = CutCol Then Col_ItemVal = PasteCol - 1 Else If Col_ItemVal > CutCol And Col_ItemVal < PasteCol Then Col_ItemVal = Col_ItemVal - 1
    
    If Col_GodownCode = CutCol Then Col_GodownCode = PasteCol - 1 Else If Col_GodownCode > CutCol And Col_GodownCode < PasteCol Then Col_GodownCode = Col_GodownCode - 1
    If Col_Godown = CutCol Then Col_Godown = PasteCol - 1 Else If Col_Godown > CutCol And Col_Godown < PasteCol Then Col_Godown = Col_Godown - 1
    If Col_PName = CutCol Then Col_PName = PasteCol - 1 Else If Col_PName > CutCol And Col_PName < PasteCol Then Col_PName = Col_PName - 1
    If Col_LName = CutCol Then Col_LName = PasteCol - 1 Else If Col_LName > CutCol And Col_ItemVal < PasteCol Then Col_LName = Col_LName - 1
    
Else
'    Sign1 = "": Sign2 = "<": Sign3 = " >=": Sign4 = "+1"
    If Col_SONo = CutCol Then Col_SONo = PasteCol Else If Col_SONo < CutCol And Col_SONo >= PasteCol Then Col_SONo = Col_SONo + 1
    If Col_SOSrNo = CutCol Then Col_SOSrNo = PasteCol Else If Col_SOSrNo < CutCol And Col_SOSrNo >= PasteCol Then Col_SOSrNo = Col_SOSrNo + 1
    If Col_SONoCode = CutCol Then Col_SONoCode = PasteCol Else If Col_SONoCode < CutCol And Col_SONoCode >= PasteCol Then Col_SONoCode = Col_SONoCode + 1
    If Col_PNo = CutCol Then Col_PNo = PasteCol Else If Col_PNo < CutCol And Col_PNo >= PasteCol Then Col_PNo = Col_PNo + 1
    If Col_Unit = CutCol Then Col_Unit = PasteCol Else If Col_Unit < CutCol And Col_Unit >= PasteCol Then Col_Unit = Col_Unit + 1
    If Col_MRP = CutCol Then Col_MRP = PasteCol Else If Col_MRP < CutCol And Col_MRP >= PasteCol Then Col_MRP = Col_MRP + 1
    If Col_Taxable = CutCol Then Col_Taxable = PasteCol Else If Col_Taxable < CutCol And Col_Taxable >= PasteCol Then Col_Taxable = Col_Taxable + 1
    If Col_Qty = CutCol Then Col_Qty = PasteCol Else If Col_Qty < CutCol And Col_Qty >= PasteCol Then Col_Qty = Col_Qty + 1
    If Col_Rate = CutCol Then Col_Rate = PasteCol Else If Col_Rate < CutCol And Col_Rate >= PasteCol Then Col_Rate = Col_Rate + 1
    If Col_MRPRate = CutCol Then Col_MRPRate = PasteCol Else If Col_MRPRate < CutCol And Col_MRPRate >= PasteCol Then Col_MRPRate = Col_MRPRate + 1
    If Col_Amt = CutCol Then Col_Amt = PasteCol Else If Col_Amt < CutCol And Col_Amt >= PasteCol Then Col_Amt = Col_Amt + 1
    If Col_DiscPer = CutCol Then Col_DiscPer = PasteCol Else If Col_DiscPer < CutCol And Col_DiscPer >= PasteCol Then Col_DiscPer = Col_DiscPer + 1
    If Col_DiscAmt = CutCol Then Col_DiscAmt = PasteCol Else If Col_DiscAmt < CutCol And Col_DiscAmt >= PasteCol Then Col_DiscAmt = Col_DiscAmt + 1
    If Col_ItemVal = CutCol Then Col_ItemVal = PasteCol Else If Col_ItemVal < CutCol And Col_ItemVal >= PasteCol Then Col_ItemVal = Col_ItemVal + 1

    If Col_GodownCode = CutCol Then Col_GodownCode = PasteCol Else If Col_GodownCode < CutCol And Col_GodownCode >= PasteCol Then Col_GodownCode = Col_GodownCode + 1
    If Col_Godown = CutCol Then Col_Godown = PasteCol Else If Col_Godown < CutCol And Col_Godown >= PasteCol Then Col_Godown = Col_Godown + 1
    If Col_PName = CutCol Then Col_PName = PasteCol Else If Col_PName < CutCol And Col_PName >= PasteCol Then Col_PName = Col_PName + 1
    If Col_LName = CutCol Then Col_LName = PasteCol Else If Col_LName < CutCol And Col_ItemVal >= PasteCol Then Col_LName = Col_LName + 1
End If
Grid_Ini
MoveRec
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
GSQL = "SELECT Syctrl.SprInvFooter,S.DocID,S.V_Type,S.V_No,S.V_Date,S.Cash_Credit,S.Job_DocID,S.Party_Code,s.Party_Name,s.Address," & _
    "SG.NamePrefix,SG.Name,SG.Add1, SG.Add2, SG.Add3,City.CityName,SG.PIN,SG.Phone,SG.CSTNo,S.L_C, S.REP_CODE, S.Form_Code, S.RoadPermit_FormCode," & _
    "S.GR_RR_No, S.GR_RR_Date,S.CrAc, S.Case_No, S.Case_Mark, S.Mode_Dispatch, S.Transport, S.Remarks,S.SprAmt_MRP_TB, S.SprAmt_MRP_TP," & _
    "S.OilAmt_MRP_TB, S.OilAmt_MRP_TP,S.SprAmt_TB,S.SprAmt_TP, S.OilAmt_TB, S.OilAmt_TP, S.D_Per_TB, S.D_Amt_TB, S.D_Per_TP,S.D_Amt_TP," & _
    "S.D_Per_MRP_TB, S.D_Amt_MRP_TB, S.D_Per_MRP_TP, S.D_Amt_MRP_TP,S.Addition, S.Gen_Sur_Per, S.Gen_Sur_Amt, S.Trans_Amt, S.LineFileTaxSum," & _
    "S.Tax_Per, S.Tax_Amt, S.Tax_AmtMRP, S.Tax_Sur_Per,S.Tax_Sur_Amt,S.TaxSur_AmtMRP,S.Packing, S.TOT_Per, S.Tot_Amt, S.TOT_AmtMRP, " & _
    "S.ReSalTax_Per, S.ReSalTax_Amt,S.Total_Amt, S.Rounded, S.Det_Tax, S.GP_No, S.GP_Date, S.Printed_YN,S.Invoice_DocId, S.U_Name,S.U_EntDt,S.CancelYN, " & _
    "" & vIsNull("SPStk.Srl_No", "0") & " as Srl_No, " & xIsNull("SPStk.V_Date", "") & " as SPStk_V_Date, " & xIsNull("SPStk.Party_Code", "") & " as SPStkParty_Code,SPStk.L_C, " & xIsNull("SPStk.Job_DocID", "") & " as Job_DocID," & _
    "SPStk.Mech_Code, SPStk.Order_DocId,SPStk.Order_Srl_No,SPStk.Part_No,Part.Part_Name, SPStk.Lub_Category, SPStk.Godown," & _
    "" & vIsNull("SPStk.Qty_Doc", "0") & " as Qty_Doc, " & vIsNull("SPStk.Qty_Rec", "0") & " as Qty_Rec, " & vIsNull("SPStk.Qty_Iss", "0") & " as Qty_Iss," & _
    "" & vIsNull("SPStk.Qty_Ret", "0") & " as Qty_Ret, " & vIsNull("SPStk.Tax_YN", "0") & " as Tax_YN, " & vIsNull("SPStk.MRP_YN", "0") & " as MRP_YN," & _
    "" & vIsNull("SPStk.Rate", "0") & " as Rate, " & vIsNull("SPStk.MRP_Rate", "0") & " as MRP_Rate, " & vIsNull("SPStk.Disc_Per", "0") & " as Disc_Per," & _
    "" & vIsNull("SPStk.Disc_Amt", "0") & " as Disc_Amt, " & vIsNull("SPStk.AMOUNT", "0") & " as AMOUNT, " & vIsNull("SPStk.Ord_DiscPer", "0") & " as Ord_DiscPer," & _
    "" & vIsNull("SPStk.Ord_DiscAmt", "0") & " as Ord_DiscAmt, " & vIsNull("SPStk.Net_Amt", "0") & " as Net_Amt, " & xIsNull("SPStk.Purpose", "") & " as Purpose," & _
    "SPStk.Part_SrlNo,SPStk.Remark,SPStk.Invoice_DocId as SPStk_Invoice_DocId, SPStk.V_Date2, " & vIsNull("SPStk.Rate2", "0") & " as Rate2, " & vIsNull("SPStk.MRP_Rate2", "0") & " as MRP_Rate2," & _
    "" & vIsNull("SPStk.Disc_Per2", "0") & " as Disc_Per2, " & vIsNull("SPStk.Disc_Amt2", "0") & " as Disc_Amt2, " & vIsNull("SPStk.Amount2", "0") & " as Amount2," & _
    "" & vIsNull("SPStk.Ord_DiscPer2", "0") & " as Ord_DiscPer2, " & vIsNull("SPStk.Ord_DiscAmt2", "0") & " as Ord_DiscAmt2, " & vIsNull("SPStk.Net_Amt2", "0") & " as Net_Amt2,SPStk.Printed2,Syctrl.GatePassOnSprInv , " & vIsNull("SPStk.TaxPer", "0") & " as TaxPer, " & vIsNull("SPStk.TaxAmt", "0") & " as TaxAmt " & _
    "FROM (((SP_Sale as S left JOIN SP_Stock as SPStk ON S.DocID = SPStk.DocId) " & _
    "left JOIN Part ON SPStk.Part_No = Part.PART_NO and Part.Div_Code = left(SPStk.Docid,1)) " & _
    "LEFT JOIN (SubGroup as SG LEFT JOIN City ON SG.CityCode = City.CityCode) ON S.Party_Code = SG.SubCode) " & _
    "LEFT JOIN Syctrl ON Syctrl.LinkTable<>S.U_AE " & _
    "where S.DocId='" & Master!SearchCode & "'"
    
mRepName = IIf(OptPlain.Value = True, "SprSaleChal", "SprSaleChal")
Select Case Index
    Case PScreen, PWindows
        Call WindowsPrint(Index, GSQL)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint(GSQL, Optpre.Value)
        FrmPrn.Visible = False
    Case PSetUp
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
Dim RstCompDet As ADODB.Recordset
Dim RstRep As ADODB.Recordset
Dim I As Integer, j As Integer
On Error GoTo ERRORHANDLER
        
        Set RstRep = GCn.Execute(mQry)
        If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
        
        Set RstCompDet = GCn.Execute("select S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
        
        CreateFieldDefFile RstRep, PubRepoPath + "\" & mRepName & ".TTX", True
        If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
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
                    rpt.FormulaFields(I).TEXT = "'" & DeCodeDocID(Master!SearchCode, Document_Type) & "'"
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
                            rpt.FormulaFields(I).TEXT = "'" & IIf(mVType = SalChalType, "SALE CHALLAN", "TRANSFER CHALLAN") & "'"
                    End Select
                Next
                rpt.PrintOut False
                If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
                    GCn.Execute "update Sp_Sale set Printed_YN = 1  where Sp_Sale.docid='" & Master!SearchCode & "' "
                End If
            Case 1  'screen
                Call Report_View(rpt, IIf(mVType = SalChalType, "SALE CHALLAN", "TRANSFER CHALLAN"), , True)
        End Select
        Set RstCompDet = Nothing
        Set RstRep = Nothing
        CmdPrint(PSetUp).Tag = ""
ERRORHANDLER:
        CheckError
End Sub

Private Sub PrinerSetUp()
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.PrinterSetup (0)
CmdPrint(PSetUp).Tag = "1"
LblPrinter.CAPTION = rpt.PrinterName
End Sub

Private Sub SpeedPrint(mQry As String, PrePrinted)
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
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    
    Set RstRep = GCn.Execute(mQry)
    If RstRep.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    
    Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select SprInvFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
 
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mGatePass = 9
    mFooter = 19    'Line For Gate Pass =9 ,Line For NonTax Detail = 5
    mFooter = mFooter + FooterCnt
    mFooter = IIf(RstRep!Printed_YN = 0 And RstRep!CancelYN = 0, mFooter + mGatePass, mFooter)

    'Sale Bill Header
    mDocStr = IIf(mVType = SalChalType, "SALE CHALLAN", "STOCK TRANSFER")
    mDupStr = IIf(RstRep!Printed_YN = 1, "(DUPLICATE)", "")
    Close #1
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
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
        Print #1, PSTR(XNull(RstCompDet!S_SecLST) & IIf(XNull(RstCompDet!S_SecLST_Date) = "", "", " Dt. " & RstCompDet!S_SecLST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!S_SecPhone) = "", "", "PHONE : " & XNull(RstCompDet!S_SecPhone)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
        Print #1, PSTR(XNull(RstCompDet!S_SecCST) & IIf(XNull(RstCompDet!S_SecCST_Date) = "", "", " Dt. " & RstCompDet!S_SecCST_Date), 45) & Space(8) & PSTR(IIf(XNull(RstCompDet!S_SecFax) = "", "", "FAx   : " & XNull(RstCompDet!S_SecFax)), 27, , AlignRight, " ")
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth)
    mHeader = mHeader + 1
    Print #1, mChr18 & "To," & mEmph
    mHeader = mHeader + 1
    Print #1, PSTR(RstRep!NamePrefix & " " & RstRep!Party_Name, 40) & Space(1) & PSTR(mDocStr & " NO.", Len(mDocStr) + 5) & " : " & PrinID(RstRep!DocID) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstRep!Add1), 40) & Space(1) & mEmph & PSTR(mDocStr & " DATE", Len(mDocStr) + 5) & " : " & PSTR(STR(RstRep!V_Date), 14) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR(XNull(RstRep!Add2), 40) & Space(1) & mEmph & IIf(RstRep!CancelYN = 1, "** CANCELLED **", "") & mEmph1
    mHeader = mHeader + 1
    Print #1, XNull(RstRep!Add3) & IIf(XNull(RstRep!CityName) <> "", ",", "") & XNull(RstRep!CityName)
    mHeader = mHeader + 1
    Print #1, mDoub & "CST NO.:" & XNull(RstRep!CstNo) & mDoub1
    mHeader = mHeader + 1
    
    Print #1, Replace(Space(PageWidth), " ", "-") & mChr17 & mDoub
    mHeader = mHeader + 1
    If PubVATYN = 1 Then
        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 30) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & PSTR("DISC", 8, , AlignRight) & " %" & PSTR("DISC. AMT", 10, , AlignRight) & PSTR("TAX %", 6, , AlignRight) & PSTR("TAX AMT", 10, , AlignRight) & PSTR("AMOUNT", 12, , AlignRight) & mDoub1 & mChr18
        mHeader = mHeader + 1
    Else
        Print #1, PSTR("SRL.No", 7) & PSTR("PART NO.", 22) & PSTR("DESCRIPTION", 35) & PSTR("QUANTITY", 12, , AlignRight) & PSTR("RATE", 11, , AlignRight) & Space(2) & "<-----DISCOUNT----- >" & "<---------AMOUNT--------- >"
        mHeader = mHeader + 1
        Print #1, Space(89) & PSTR("%", 10, , AlignRight) & PSTR("AMOUNT", 10, , AlignRight) & PSTR("TAXPAID", 12, , AlignRight) & PSTR("TAXABLE", 12, , AlignRight) & mDoub1 & mChr18
        mHeader = mHeader + 1
    End If
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
                Print #1, PSTR(XNull(RstRep!Add1), 40) & Space(1) & mEmph & PSTR(mDocStr & " DATE", Len(mDocStr) + 5) & " : " & PSTR(STR(RstRep!V_Date), 14) & mEmph1
                mHeader = mHeader + 1
                Print #1, PSTR(XNull(RstRep!Add2), 40) & Space(1) & mEmph & IIf(RstRep!CancelYN = 1, "** CANCELLED **", "") & mEmph1
                mHeader = mHeader + 1
                Print #1, XNull(RstRep!Add3) & IIf(XNull(RstRep!CityName) <> "", ",", "") & XNull(RstRep!CityName)
                mHeader = mHeader + 1
                Print #1, mDoub & "CST NO.:" & XNull(RstRep!CstNo) & mDoub1
                
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
            If PubVATYN = 1 Then
                PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstRep!Part_No, 22, , AlignLeft) & PSTR(RstRep!Part_Name, 30) & PSTR(RstRep!Qty_Iss, 12, 3)
                PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstRep!MRP_YN = 1, "M", "L") & _
                PSTR(RstRep!Disc_Per, 8, 2) & " %" & PSTR(Format(RstRep!Disc_Amt, "0.00"), 10, 2, AlignRight) & _
                PSTR(Format(RstRep!TaxPer, "0.00"), 6, 2, AlignRight) & PSTR(Format(RstRep!TaxAmt, "0.00"), 10, 2, AlignRight) & _
                PSTR(Format(RstRep!Net_Amt, "0.00"), 12, 2, AlignRight)
            Else
                PrintStr = PSTR(Trim(STR(mSlNo)) & ".", 7) & PSTR(RstRep!Part_No, 22, , AlignLeft) & PSTR(RstRep!Part_Name, 35) & PSTR(RstRep!Qty_Iss, 12, 3)
                PrintStr = PrintStr & PSTR(mRate, 11, 2) & " " & IIf(RstRep!MRP_YN = 1, "M", "L") & _
                PSTR(RstRep!Disc_Per, 8, 2) & " %" & PSTR(RstRep!Disc_Amt, 10, 2) & _
                IIf(RstRep!Tax_YN = 0, PSTR(RstRep!Net_Amt, 12, 2) & PSTR(0, 12, 2), PSTR(0, 12, 2) & PSTR(RstRep!Net_Amt, 12, 2))
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
    RstRep.MovePrevious
    Print #1, mChr18 & "Customer's Signature"
' SALE FOOTER
    '22 space maintain between heading and :
    Print #1, Replace(Space(21), " ", "-") & "TaxPaid" & Replace(Space(12), " ", "-") & "Taxable" & Replace(Space(33), " ", "-")
    If PubVATYN = 1 Then
        Print #1, PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 12, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
        ; " | " & PSTR("V A T     ", 10, 0) & Space(6) & PSTR(RstRep!Tax_Amt, 12, 2) & mDoub
        
        Print #1, PSTR("MRP Items Amt", 16) & PSTR(RstRep!SprAmt_MRP_TP + RstRep!OilAmt_MRP_TP, 12, 2) & Space(8) & PSTR(RstRep!SprAmt_MRP_TB + RstRep!OilAmt_MRP_TB, 12, 2) & mDoub1 _
        ; " | " & PSTR("S A T     ", 10, 0) & Space(6) & PSTR(Val(txt(SatAmt)), 12, 2) & mDoub
    Else
        Print #1, PSTR("Item Disc.Amt", 16) & PSTR(Val(txt(IWDiscTotTP)), 12, 2) & Space(8) & PSTR(Val(txt(IWDiscTotTB)), 12, 2) _
        ; " | " & PSTR("Sales Tax ", 10, 0) & PSTR(RstRep!Tax_Per, 5, 2) & "%" & PSTR(RstRep!Tax_Amt, 12, 2) & mDoub
        
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
    'Print #1, PSTR(mTAXDESC, 25) & PSTR("E & OE", PageWidth - 25, , AlignRight)
    If XNull(RstRep!Remarks) <> "" Then
        Print #1, "Remarks : " & XNull(RstRep!Remarks)
        Print #1, Replace(Space(PageWidth), " ", "-")
    End If
    Print #1, mChr17 & "E & OE" & mChr18
    Print #1, ""
    Print #1, Space(PageWidth - Len("For " & PubComp_Name)) & "For " & mEmph & PubComp_Name & mEmph1 & mDoub
    Print #1, "Terms & Condition " & mDoub1 & Replace(Space(PageWidth - 18), " ", "-") & mChr17
    Footer = Footer & vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
     Next
    Print #1, Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
    ' Gate Pass Footer()
    If RstRep!GatePassOnSprInv = 0 And RstRep!Printed_YN = 0 And RstRep!CancelYN = 0 Then
        Print #1, Replace(Space(PageWidth), " ", "-")
        Print #1, PRN_TIT("* " & mDocStr & " GATE PASS " & mDupStr & " *", "A", 80) & mEmph
        Print #1, "GATE PASS No. & DATE : " & XNull(RstRep!gp_no) & "  " & IIf(IsNull(RstRep!GP_Date), "", ConvertDate(RstRep!GP_Date)) & mEmph1
        Print #1, PSTR(RstRep!NamePrefix & " " & RstRep!Party_Name, 40) & Space(1) & PSTR(mDocStr & " NO.", Len(mDocStr) + 5) & " : " & PrinID(RstRep!DocID)
        Print #1, PSTR(XNull(RstRep!CityName), 40) & Space(1) & PSTR(mDocStr & " DATE", Len(mDocStr) + 5) & " : " & CDate(RstRep!V_Date)
        Print #1, "Goods of " & mEmph & "Rs." & LTrim(PSTR(Val(txt(NetAmt)), 9, 2)) & mEmph1 & " as per Document No. are being permitted for out."
        Print #1, "Mode of dispatch :" & XNull(RstRep!Mode_Dispatch)
        Print #1, ""
        Print #1, "Customer's Signature" & Space(50 - Len(PubComp_Name)) & "for " & mEmph & PubComp_Name & mEmph1
        Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
    End If
    Print #1, mEject
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



