VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmChequePayment 
   Appearance      =   0  'Flat
   BackColor       =   &H00CFE0E0&
   Caption         =   "Cheque Payment"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10875
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
   ScaleHeight     =   7635
   ScaleWidth      =   10875
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.OptionButton Opt2 
      Caption         =   "Option1"
      Height          =   405
      Index           =   2
      Left            =   9060
      TabIndex        =   61
      Top             =   6300
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.OptionButton Opt2 
      Caption         =   "Option1"
      Height          =   405
      Index           =   1
      Left            =   8985
      TabIndex        =   60
      Top             =   6795
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSDataGridLib.DataGrid DgAcCode 
      Height          =   2175
      Left            =   2655
      Negotiate       =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   6135
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
         DataField       =   "Name"
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
            ColumnWidth     =   2865.26
         EndProperty
      EndProperty
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
      Left            =   90
      TabIndex        =   35
      Top             =   3675
      Visible         =   0   'False
      Width           =   5025
      Begin VB.OptionButton OptCheque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "Cheque Print"
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
         Left            =   840
         TabIndex        =   62
         Top             =   750
         Width           =   1410
      End
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
         Picture         =   "FrmChequePayment.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
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
         Picture         =   "FrmChequePayment.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "FrmChequePayment.frx":0678
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
         TabIndex        =   42
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "FrmChequePayment.frx":0982
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
         TabIndex        =   41
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "FrmChequePayment.frx":0C8C
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   300
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton OptVoucher 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
         Caption         =   "Voucher Print"
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
         Left            =   840
         TabIndex        =   36
         Top             =   435
         Width           =   1410
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
         Left            =   345
         TabIndex        =   46
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
         TabIndex        =   45
         Top             =   0
         Width           =   4695
      End
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
      Height          =   255
      Index           =   12
      Left            =   9015
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   2565
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
      Height          =   255
      Index           =   13
      Left            =   9015
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   2835
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   5
      Left            =   2025
      MaxLength       =   3
      TabIndex        =   15
      Top             =   2415
      Width           =   1245
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
      Height          =   255
      Index           =   16
      Left            =   3540
      MaxLength       =   255
      TabIndex        =   8
      Text            =   "PayTo2"
      Top             =   525
      Visible         =   0   'False
      Width           =   780
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
      Height          =   255
      Index           =   15
      Left            =   2025
      MaxLength       =   255
      TabIndex        =   7
      Top             =   1065
      Width           =   4305
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   14
      Left            =   5085
      MaxLength       =   11
      TabIndex        =   13
      Top             =   1875
      Width           =   1245
   End
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2175
      Left            =   1470
      Negotiate       =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5910
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Site Name"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   2310.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGVType 
      Height          =   2175
      Left            =   1050
      Negotiate       =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5700
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
         DataField       =   "Name"
         Caption         =   "Voucher Type"
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
   Begin VB.CommandButton cmdPost 
      Caption         =   "Re-Post"
      Height          =   330
      Left            =   7320
      TabIndex        =   49
      Top             =   15
      Visible         =   0   'False
      Width           =   1005
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
      Height          =   255
      Index           =   6
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1335
      Width           =   4305
   End
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   435
      Negotiate       =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5205
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
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   7185
      TabIndex        =   31
      Top             =   4365
      Visible         =   0   'False
      Width           =   1920
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   30
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   120
         Width           =   1860
         _ExtentX        =   3281
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
         BackColor       =   12640511
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
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   3825
      Left            =   -2235
      Negotiate       =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4890
      Visible         =   0   'False
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   6747
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
         DataField       =   "Father"
         Caption         =   "Father / Husband Name"
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
         BeginProperty Column01 
            ColumnWidth     =   5520.189
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   9
      Left            =   2025
      MaxLength       =   11
      TabIndex        =   12
      Top             =   1875
      Width           =   1245
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
      Height          =   255
      Index           =   8
      Left            =   5085
      MaxLength       =   6
      TabIndex        =   11
      Top             =   1605
      Width           =   1245
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
      Height          =   255
      Index           =   2
      Left            =   8235
      TabIndex        =   3
      Top             =   1680
      Width           =   2355
   End
   Begin VB.TextBox txt 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   2025
      MaxLength       =   12
      TabIndex        =   10
      Top             =   1605
      Width           =   1245
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
      Height          =   255
      Index           =   11
      Left            =   2025
      MaxLength       =   255
      TabIndex        =   14
      Top             =   2145
      Width           =   4305
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
      Height          =   255
      Index           =   1
      Left            =   8235
      TabIndex        =   2
      Top             =   1410
      Width           =   2355
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
      Height          =   255
      Index           =   4
      Left            =   9270
      TabIndex        =   5
      Top             =   2220
      Width           =   1320
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   661
   End
   Begin VB.TextBox txt 
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
      Height          =   255
      Index           =   10
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   6
      Top             =   795
      Width           =   4305
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
      Height          =   255
      Index           =   0
      Left            =   8235
      TabIndex        =   1
      Top             =   885
      Width           =   2355
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
      Height          =   255
      Index           =   3
      Left            =   8235
      TabIndex        =   4
      Top             =   1950
      Width           =   2355
   End
   Begin VB.Label LblAcPostBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Posting By"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   7065
      TabIndex        =   58
      Top             =   2595
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label LblAcPostDt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Posting Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   7065
      TabIndex        =   57
      Top             =   2865
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Payee (Yes/No)"
      BeginProperty Font 
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
      Left            =   300
      TabIndex        =   54
      Top             =   2430
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issued To"
      BeginProperty Font 
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
      Left            =   300
      TabIndex        =   53
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clearing Date"
      BeginProperty Font 
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
      Left            =   3795
      TabIndex        =   52
      Top             =   1890
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chq. Date"
      BeginProperty Font 
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
      Left            =   300
      TabIndex        =   51
      Top             =   1890
      Width           =   870
   End
   Begin VB.Label LblUser 
      Alignment       =   1  'Right Justify
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
      Left            =   4605
      TabIndex        =   50
      Top             =   3510
      Width           =   6060
   End
   Begin VB.Label LblPartyBal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bal."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6360
      TabIndex        =   48
      Top             =   795
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblAcBal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bal."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6345
      TabIndex        =   47
      Top             =   1620
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chq. No."
      BeginProperty Font 
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
      Left            =   3795
      TabIndex        =   29
      Top             =   1620
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party A/c"
      BeginProperty Font 
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
      Left            =   300
      TabIndex        =   28
      Top             =   810
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr Date"
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
      Index           =   27
      Left            =   7080
      TabIndex        =   27
      Top             =   1695
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount (Rs.)"
      BeginProperty Font 
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
      Index           =   34
      Left            =   300
      TabIndex        =   25
      Top             =   1620
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration"
      BeginProperty Font 
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
      Left            =   300
      TabIndex        =   24
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/c Name"
      BeginProperty Font 
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
      Left            =   300
      TabIndex        =   23
      Top             =   1350
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name"
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
      Index           =   15
      Left            =   7080
      TabIndex        =   22
      Top             =   1440
      Width           =   885
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1755
      Left            =   6990
      Top             =   795
      Width           =   3690
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Prefix"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8235
      TabIndex        =   21
      Top             =   2220
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr No."
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
      Index           =   1
      Left            =   7080
      TabIndex        =   20
      Top             =   2220
      Width           =   540
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division           "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7080
      TabIndex        =   19
      Top             =   1170
      Width           =   1440
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code    "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8895
      TabIndex        =   18
      Top             =   1170
      Width           =   1155
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   42
      Left            =   7080
      TabIndex        =   17
      Top             =   930
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr Type"
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
      Index           =   2
      Left            =   7080
      TabIndex        =   16
      Top             =   1965
      Width           =   675
   End
End
Attribute VB_Name = "FrmChequePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PubReceiptType As String
Dim RsSite As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim RsParty As ADODB.Recordset
Dim RsAcCode As ADODB.Recordset
Dim RsVType As ADODB.Recordset
Dim RsVno As ADODB.Recordset
Dim DocID As String
Public mVType As String, mNCat As String
Dim VoucherEditFlag As Boolean
Dim vPrefix As String, mqry As String, mSearchCode As String



Private Const TxtDocID          As Byte = 0
Private Const SiteCode          As Byte = 1
Private Const VDate             As Byte = 2
Private Const VType             As Byte = 3
Private Const SerialNo          As Byte = 4
Private Const AcPayeeCheque     As Byte = 5
Private Const AcHead            As Byte = 6
Private Const Amount            As Byte = 7
Private Const Chq_No            As Byte = 8
Private Const Chq_Date          As Byte = 9
Private Const Party             As Byte = 10
Private Const Narr              As Byte = 11
Private Const AcPostByName      As Byte = 12
Private Const AcPostDate        As Byte = 13
Private Const Clg_Date          As Byte = 14
Private Const PayTo1            As Byte = 15
Private Const PayTo2            As Byte = 16




Private Const SiteCode1 As Byte = 0
Private Const VType1 As Byte = 1
Private Const FromVno As Byte = 2
Private Const ToVno As Byte = 3

Dim ListArray As Variant
Dim mListItem As ListItem
Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String
Dim mTaxAcHead$, mTaxSurAcHead$, mTOTAcHead$

Private Const CnsPrinter As Byte = 0
Private Const CnsScreen As Byte = 1
Private Const CnsPrinterSetup As Byte = 2
Private Const CnsClose As Byte = 3

Private Sub cmdPost_Click()
On Error GoTo ErrLoop
Dim I As Integer
Dim LedgAry(4) As LedgRec, mNarr$, mResult As Byte

    Master.MoveFirst
    Do Until Master.EOF
        Call MoveRec
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        
        If CDate(Format(txt(VDate).TEXT, "dd/MMM/yyyy")) < PubStartDate Or CDate(Format(txt(VDate).TEXT, "dd/MMM/yyyy")) > PubEndDate Then GoTo MyNextRecord
        Call TopCtrl1_eEdit
        'A/c Posting
        If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
            Call AccountPosting
        End If
        'eof posting
MyNextRecord:
        Disp_Text SETS("INI", Me, Master)
        Master.MoveNext
    Loop
    Exit Sub
ErrLoop:
    Call CheckError
End Sub

Private Sub DGSite_Click()
If FrmPrn.Visible = False Then
    DGSite.Visible = False
    If RsSite.RecordCount > 0 Then
        txt(SiteCode).TEXT = RsSite!Name
        txt(SiteCode).Tag = RsSite!Code
    End If
    txt(SiteCode).SetFocus
Else
    DGSite.Visible = False
    If RsSite.RecordCount > 0 Then
        txtPrint(SiteCode1).TEXT = RsSite!Name
        txtPrint(SiteCode1).Tag = RsSite!Code
    End If
    txtPrint(SiteCode1).SetFocus
End If
End Sub

Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        txt(Party).TEXT = RsParty!Name
        txt(Party).Tag = RsParty!Code
    End If
    DGParty.Visible = False
    txt(Party).SetFocus
End Sub

Private Sub DgAcCode_Click()
    If RsAcCode.RecordCount > 0 Then
        txt(AcHead).TEXT = RsAcCode!Name
        txt(AcHead).Tag = RsAcCode!Code
    End If
    DgAcCode.Visible = False
    txt(AcHead).SetFocus
End Sub

Private Sub DGVno_Click()
Dim Index As Integer
If DGVno.Tag = "1" Then
    Index = ToVno
Else
    Index = FromVno
End If
    DGVno.Visible = False
    If RsVno.RecordCount > 0 Then
        txtPrint(Index).TEXT = RsVno!Code
    End If
    txtPrint(Index).SetFocus
End Sub

Private Sub DGVType_Click()
If FrmPrn.Visible = False Then
    If RsVType.RecordCount > 0 Then
        txt(VType).TEXT = RsVType!Name
        txt(VType).Tag = RsVType!Code
    End If
    DGVType.Visible = False
    txt(VType).SetFocus
Else
    If RsVType.RecordCount > 0 Then
        txtPrint(VType1).TEXT = RsVType!Name
        txtPrint(VType1).Tag = RsVType!Code
    End If
    DGVType.Visible = False
    txtPrint(VType1).SetFocus
End If
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
Dim I As Byte
    TopCtrl1.Tag = PubUParam: WinSetting Me, 4650, 11000, 850, 465
    TopCtrl1.Tag = "AEDP"
    
    DGParty.left = 0: DGParty.width = Me.width - 90: DGParty.top = txt(Amount).top: DGParty.height = Me.height - (DGParty.top + mBotScale)
    DGSite.left = 4500: DGSite.top = mTopScale
    DGVno.left = 4500: DGVno.top = mTopScale
    DGVType.left = 4500: DGVType.top = mTopScale
    FrmPrn.left = 525: FrmPrn.top = 2220
    
    '** Hide Vehicle Details if Only Vehicle Section is not activated
    If PubVCompCode = "" Then
        For I = 18 To 21
            Label3(I).Visible = False
        Next
        txt(9).Visible = False
        txt(11).Visible = False
        txt(12).Visible = False
        txt(15).Visible = False
        txt(16).Visible = False
    End If

    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "select DocId as searchcode " & _
                    " From Payment P " & _
                    " Left Join Voucher_Type Vt On P.V_Type =  Vt.V_Type " & _
                    " Where Vt.NCat In ('" & Voucher_NCat_BankPayment & "') And V_Date>=" & ConvertDate(PubStartDate) & " " & _
                    " order by V_Date desc,docid", GCn, adOpenDynamic, adLockOptimistic
    Else
        Set Master = GCn.Execute("select Top 1 DocId as SearchCode " & _
                    " from Payment P " & _
                    " Left Join Voucher_Type Vt On P.V_Type =  Vt.V_Type " & _
                    " Where Vt.NCat In ('" & Voucher_NCat_BankPayment & "') And V_Date>=" & ConvertDate(PubStartDate) & " " & _
                    " Where V_Date>=" & ConvertDate(PubStartDate) & " " & _
                    " order by V_Date desc,docid")
    End If
    Set RsSite = New ADODB.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select site_code as code,site_desc as name from site order by site_desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGSite.DataSource = RsSite
    
    Set RsVType = New ADODB.Recordset
    RsVType.CursorLocation = adUseClient
    RsVType.Open "select Voucher_Type.v_type as code,Description as name, NCat from Voucher_Type where NCat In ('" & Voucher_NCat_BankPayment & "') order by v_type ", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGVType.DataSource = RsVType
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "select SubGroup.SUBCODE as code,SubGroup.NAME, FPrefix + ' ' + FName as Father,Curr_Bal,ITWARD_NO,PANNO, Nature from SubGroup Where Nature Not In ('Cash', 'Bank') order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set RsAcCode = New ADODB.Recordset
    RsAcCode.CursorLocation = adUseClient
    RsAcCode.Open "select Sg.SUBCODE as code,Sg.NAME, Curr_Bal, Nature, ChequeReportName From SubGroup Sg Where " & xIsNull("Sg.ChequeReportName", "") & " <> '' order by Sg.name", GCn, adOpenDynamic, adLockOptimistic
    Set DgAcCode.DataSource = RsAcCode
    DgAcCode.left = txt(AcHead).left: DgAcCode.top = txt(AcHead).top + txt(AcHead).height + 20
    
    Set RsVno = New ADODB.Recordset
    RsVno.CursorLocation = adUseClient
    RsVno.Open "Select distinct v_no as code " & _
                " From Payment P " & _
                " Left Join Voucher_Type Vt On P.V_Type =  Vt.V_Type " & _
                " Where Vt.NCat In ('" & Voucher_NCat_BankPayment & "') ", GCn, adOpenDynamic, adLockOptimistic
    Set DGVno.DataSource = RsVno
    
    Call MoveRec
    Disp_Text SETS("INI", Me, Master)
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
Set Master = Nothing
Set RsParty = Nothing
Set RsAcCode = Nothing
Set RsVno = Nothing
Set RsSite = Nothing
Set RsVType = Nothing
End Sub

Private Sub ListView_Click()
txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
txt(Val(ListView.Tag)).SetFocus
FrmList.Visible = False
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    LblVPrefix.CAPTION = ""
    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
        txt(SiteCode).Tag = PubSiteCode
        txt(SiteCode) = PubSiteName
        txt(VDate).SetFocus
    Else
        txt(SiteCode).SetFocus
    End If
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant, mTrans As Boolean ',i As Integer
Dim LedgAry(1) As LedgRec, mResult As Byte, MsgStr$, mTitle$

If AcPostAuthorisation(txt(AcPostByName)) = False Then Exit Sub

If MsgBox(MsgStr, vbYesNo + vbCritical + vbDefaultButton2, mTitle) = vbYes Then
    vBook = Master.AbsolutePosition
    GCn.BeginTrans
    G_FaCn.BeginTrans
    mTrans = True
    'Unpost Ledger a/c
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, txt(TxtDocID))
    If mResult <> 1 Then err.Raise 1, , "Error in Ledger UnPosting"
    'Unposting of Ledger completed
    
    GCn.Execute ("delete from Payment where DocId='" & txt(TxtDocID) & "'")
    
    G_FaCn.CommitTrans
    GCn.CommitTrans
    mTrans = False
    Master.Requery
    If Master.RecordCount > 0 Then
        If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
    End If
    BUTTONS True, Me, Master, 0
    Call MoveRec
End If
Exit Sub

eloop1:
    If mTrans = True Then GCn.RollbackTrans: G_FaCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    If AcPostAuthorisation(txt(AcPostByName)) = False Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    txt(Party).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then CheckError
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
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
OptVoucher.Value = True
LblPrinter.CAPTION = Printer.DeviceName
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus
On Error GoTo ELoop

    'Call WindowsPrint(PScreen, "")
    Exit Sub
ELoop:
    Call CheckError
End Sub

Private Sub TopCtrl1_eRef()
    RsSite.Requery
    RsParty.Requery
    RsAcCode.Requery
    RsVType.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim Rst As ADODB.Recordset
    Dim mTrans As Boolean
    Dim mqry As String
    
    On Error GoTo errlbl
    Grid_Hide
    If IsValid(txt(SiteCode), "Site Name") = False Then Exit Sub
    If IsValid(txt(VDate), "Date") = False Then Exit Sub
    If IsValid(txt(VType), "Voucher Type") = False Then Exit Sub
    If txt(SerialNo).Enabled = True Then
        If txt(SerialNo).TEXT = "" Then MsgBox "SerialNo is required field", vbInformation, "Validation Check": txt(SerialNo).SetFocus: Exit Sub
    Else
        If txt(SerialNo).TEXT = "" Then MsgBox "SerialNo is required field", vbInformation, "Validation Check": txt(VType).SetFocus: Exit Sub
    End If
    
    If IsValid(txt(Party), "Party A/c") = False Then Exit Sub
    If IsValid(txt(AcHead), "Bank A/c") = False Then Exit Sub
    If IsValid(txt(Chq_No), "Cheque No.") = False Then Exit Sub
    If IsValid(txt(Chq_Date), "Cheque Date") = False Then Exit Sub
    
    If Trim(txt(Clg_Date)) <> "" Then
        If CDate(txt(Chq_Date)) > CDate(txt(Clg_Date)) Then
            MsgBox "Cheque Date Can't Be Greater Than From Clearing Date!...": txt(Clg_Date).SetFocus: Exit Sub
        End If
    End If
    
    If txt(Party).Tag = txt(AcHead).Tag Then
        MsgBox "Party A/c and Ledger A/c both same !" & vbCrLf & "Correct A/c Selection ", vbCritical, "A/c Checking"
        txt(AcHead).SetFocus: Exit Sub
    End If
    
    If Val(txt(Amount)) <= 0 Then
        MsgBox "Please Enter Amount", vbCritical, "Validation"
        txt(Amount).SetFocus: Exit Sub
    End If
    '********* cHECKING pOSTING cOTROLS
    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        txt(AcPostByName) = pubUName
        txt(AcPostDate) = PubServerDate
    End If
    '*********
    If TopCtrl1.TopText2.CAPTION = "Add" Then
    'lp 11-03-03
        DocID = txt(TxtDocID)
        If GCn.Execute("select count(*) from Payment where DocId='" & txt(TxtDocID) & "'").Fields(0) > 0 Then
            If VoucherEditFlag Then
                MsgBox "Document No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                Exit Sub
            Else
                txt(TxtDocID) = GetDocID(G_FaCn, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
                    MsgBox "Document No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    Exit Sub
                End If
            End If
        End If
        
        mSearchCode = txt(TxtDocID)
    End If
    
    

    GCn.BeginTrans
    G_FaCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2 = "Add" Then
        mqry = "INSERT INTO Payment ( DocId, Site_Code, V_Date, V_Type, V_No, PartyCode, Amount, AcCode, Chq_No, " & _
                " Chq_Date, Clg_Date, Narration, PayTo1, PayTo2, AcPayeeCheque, U_Name, U_EntDt, U_AE, AcPostByU_Name, AcPostByU_EntDt, AddBy, AddDate) " & _
                " VALUES  ( '" & txt(TxtDocID) & "', '" & PubSiteCode & txt(SiteCode).Tag & "', " & ConvertDate(txt(VDate)) & "," & _
                " '" & txt(VType).Tag & "', " & Val(txt(SerialNo).TEXT) & ", '" & txt(Party).Tag & "', " & Val(txt(Amount).TEXT) & "," & _
                " '" & txt(AcHead).Tag & "', '" & txt(Chq_No).TEXT & "', " & ConvertDate(txt(Chq_Date).TEXT) & ", " & ConvertDate(txt(Clg_Date).TEXT) & "," & _
                " '" & txt(Narr).TEXT & "', '" & txt(PayTo1).TEXT & "', '" & txt(PayTo2).TEXT & "', " & IIf(txt(AcPayeeCheque).TEXT = "Yes", 1, 0) & "," & _
                " '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A','" & txt(AcPostByName) & "'," & ConvertDate(txt(AcPostDate)) & "," & _
                " '" & pubUName & "'," & ConvertDate(PubServerDate) & ") "
        
        GCn.Execute (mqry)
        
        UpdVouSrlNo G_FaCn, txt(TxtDocID), txt(VDate)
    Else
            
        mqry = "Update Payment Set V_Date = " & ConvertDate(txt(VDate)) & ", PartyCode = '" & txt(Party).Tag & "', Amount = " & Val(txt(Amount).TEXT) & ", AcCode = '" & txt(AcHead).Tag & "', Chq_No = '" & txt(Chq_No).TEXT & "', " & _
                " Chq_Date = " & ConvertDate(txt(Chq_Date).TEXT) & ", Clg_Date = " & ConvertDate(txt(Clg_Date).TEXT) & ", Narration = '" & txt(Narr).TEXT & "', PayTo1 = '" & txt(PayTo1).TEXT & "', PayTo2 = '" & txt(PayTo2).TEXT & "'," & _
                " AcPayeeCheque = " & IIf(txt(AcPayeeCheque).TEXT = "Yes", 1, 0) & ", " & _
                " U_Name = '" & pubUName & "', U_EntDt = " & ConvertDate(PubServerDate) & ", U_AE = 'E', AcPostByU_Name = '" & txt(AcPostByName) & "', AcPostByU_EntDt = " & ConvertDate(txt(AcPostDate)) & ", ModifyBy = '" & pubUName & "', ModifyDate = " & ConvertDate(PubServerDate) & " " & _
                " Where DocId = '" & txt(TxtDocID) & "'"
            
        GCn.Execute (mqry)
    End If
    
    'A/c Posting
    Call AccountPosting
    
    G_FaCn.CommitTrans
    GCn.CommitTrans
    mTrans = False
    
        
    Set Rst = Nothing
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select DocId as searchcode " & _
                                    " From Payment P " & _
                                    " Left Join Voucher_Type Vt On P.V_Type =  Vt.V_Type " & _
                                    " Where Vt.NCat In ('" & Voucher_NCat_BankPayment & "') And V_Date>=" & ConvertDate(PubStartDate) & "  And DocId = '" & mSearchCode & "' " & _
                                    " order by V_Date desc,docid")
                    
                    
    End If
    Master.FIND "SearchCode = '" & mSearchCode & "'"
    
    Call TopCtrl1_ePrn
    
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    
    Exit Sub
errlbl:
    If mTrans = True Then
        GCn.RollbackTrans: G_FaCn.RollbackTrans
    End If
    CheckError
Exit Sub
End Sub

Private Sub AccountPosting()
    Dim LedgAry(4) As LedgRec, mNarr$, mResult As Byte
    Dim I As Integer

    mNarr = txt(Narr)

    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        I = 0
        LedgAry(I).SubCode = txt(Party).Tag
        LedgAry(I).ContraSub = txt(AcHead).Tag
        LedgAry(I).AmtDr = Val(txt(Amount))
        LedgAry(I).Narration = mNarr
        LedgAry(I).Chq_No = txt(Chq_No).TEXT
        LedgAry(I).Chq_Date = txt(Chq_Date).TEXT
        LedgAry(I).Clg_Date = txt(Clg_Date).TEXT
        
        I = I + 1
        LedgAry(I).SubCode = txt(AcHead).Tag
        LedgAry(I).ContraSub = txt(Party).Tag
        LedgAry(I).AmtCr = Val(txt(Amount))
        LedgAry(I).Narration = mNarr
        LedgAry(I).Chq_No = txt(Chq_No).TEXT
        LedgAry(I).Chq_Date = txt(Chq_Date).TEXT
        LedgAry(I).Clg_Date = txt(Clg_Date).TEXT
        
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, txt(TxtDocID), CDate(txt(VDate)), mNarr)
        If mResult <> 1 Then err.Raise 1, , "Error in Ledger Posting"
    End If
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "select P.DocId as searchcode, " & cDt("P.V_Date") & " As V_Date, P.V_Type, " & cCStr("P.V_No") & " As V_No, " & cCStr("P.Chq_No") & " As Chq_No, P.Chq_Date As Chq_Date, P.Narration " & _
            " From Payment P " & _
            " Left Join Voucher_Type Vt On P.V_Type =  Vt.V_Type " & _
            " Where Vt.NCat In ('" & Voucher_NCat_BankPayment & "') And V_Date>=" & ConvertDate(PubStartDate) & " "
            
            Set SearchForm = Me
    
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
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
        
        Set Master = GCn.Execute("select DocId as searchcode " & _
                                    " From Payment P " & _
                                    " Left Join Voucher_Type Vt On P.V_Type =  Vt.V_Type " & _
                                    " Where Vt.NCat In ('" & Voucher_NCat_BankPayment & "') And V_Date>=" & ConvertDate(PubStartDate) & "  And DocId = '" & MyValue & "' " & _
                                    " order by V_Date desc,docid")
                    
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
If txt(SiteCode).TEXT <> "" Then
    If txt(VDate).TEXT = "" Then txt(VDate).SetFocus: Ctrl_GetFocus txt(Index): Exit Sub
    If txt(VType).TEXT = "" Then txt(VType).SetFocus: Ctrl_GetFocus txt(Index): Exit Sub
End If
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case VType
        If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsVType!Name Then
            RsVType.MoveFirst
            RsVType.FIND "Name ='" & txt(Index).TEXT & "'"
        End If

    Case SiteCode
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Then Exit Sub
        If txt(Index).TEXT = "" Then
            RsSite.MoveFirst
            RsSite.FIND "code ='" & PubSiteCode & "'"
            txt(Index).Tag = RsSite!Code
            txt(Index).TEXT = RsSite!Name
        Else
            If txt(Index).TEXT <> RsSite!Name Then
                RsSite.MoveFirst
                RsSite.FIND "name ='" & txt(Index).TEXT & "'"
            End If
        End If
        
    Case Party
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).Tag <> RsParty!Code Then
            RsParty.MoveFirst
            RsParty.FIND "Code ='" & txt(Index).Tag & "'"
        End If
        
    Case AcHead
        If RsAcCode.RecordCount = 0 Or (RsAcCode.EOF = True Or RsAcCode.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).Tag <> RsAcCode!Code Then
            RsAcCode.MoveFirst
            RsAcCode.FIND "Code ='" & txt(Index).Tag & "'"
        End If
End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case VType
        DGridTxtKeyDown DGVType, txt, Index, RsVType, KeyCode, False, 1
    Case SiteCode
        DGridTxtKeyDown DGSite, txt, Index, RsSite, KeyCode, False, 1
    Case Party
        DGridTxtKeyDown DGParty, txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
    Case AcHead
        DGridTxtKeyDown DgAcCode, txt, Index, RsAcCode, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
End Select
If FrmList.Visible = False And DGVType.Visible = False And DGParty.Visible = False And DgAcCode.Visible = False And DGSite.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VType Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
        If Index <> AcPayeeCheque Then
            Ctrl_DownKeyDown KeyCode, Shift
        End If
        If Index = AcPayeeCheque Then
            Txt_Validate Index, False
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> SiteCode Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> Party Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Dim mVehYN As Boolean
Call CheckQuote(keyascii)

Select Case Index
    Case VType
        If DGVType.Visible = True Then DGridTxtKeyPress txt, Index, RsVType, keyascii, "name"
    Case SiteCode
        If DGSite.Visible = True Then DGridTxtKeyPress txt, Index, RsSite, keyascii, "Name"
    Case SerialNo
        Call NumPress(txt(Index), keyascii, 6, 0)
    Case Party
        If DGParty.Visible = True Then DGridTxtKeyPress txt, Index, RsParty, keyascii, "Name"
    Case AcHead
        If DgAcCode.Visible = True Then DGridTxtKeyPress txt, Index, RsAcCode, keyascii, "Name"
    Case Chq_No
        Call NumPress(txt(Index), keyascii, 6, 0)
        
    Case AcPayeeCheque
        If UCase(Chr(keyascii)) = "Y" Then
            txt(Index) = "Yes"
        ElseIf UCase(Chr(keyascii)) = "N" Or keyascii = vbKeyBack Or keyascii = vbKeyDelete Then
            txt(Index) = "No"
        End If
        
        keyascii = 0
    Case Amount
        Call NumPress(txt(Index), keyascii, 9, 2)
End Select
'KeyAscii = RetDGKeyAscii()
End Sub



Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ErrLoop

    Dim Rst As ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim I As Integer, mEnb As Boolean, mVehYN As Boolean
    Select Case Index
        Case VType
            If IsValid(txt(Index), "Voucher Type") = False Then Cancel = True: Exit Sub
            If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or txt(Index).TEXT = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
                mVType = "": mNCat = ""
            Else
                txt(Index).TEXT = RsVType!Name
                txt(Index).Tag = RsVType!Code
                mVType = txt(Index).Tag
                mNCat = RsVType!NCat
                'DocID
                txt(TxtDocID) = GetDocID(G_FaCn, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                DocID = txt(TxtDocID)
            End If
            
        Case SerialNo
            If IsValid(txt(SerialNo), "Serial No.") = False Then Cancel = True:   Exit Sub
            If VoucherEditFlag = True Then      ' Manual
                txt(TxtDocID) = GetDocID(G_FaCn, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                DocID = txt(TxtDocID)
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select * From Payment Where docid='" & txt(TxtDocID) & "'", GCn, adOpenDynamic, adLockOptimistic
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                    Cancel = True
                    txt(SerialNo).SetFocus
                End If
            End If
        
        Case SiteCode
            If IsValid(txt(Index), "Site Code") = False Then Cancel = True: Exit Sub
            If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or txt(Index).TEXT = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            Else
                txt(Index).TEXT = RsSite!Name
                txt(Index).Tag = RsSite!Code
            End If
            
        Case Party
            If IsValid(txt(Index), "Party Name") = False Then Cancel = True: Exit Sub
            If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
                txt(PayTo1).TEXT = ""
            Else
                txt(Index).TEXT = RsParty!Name
                txt(Index).Tag = RsParty!Code
                LblPartyBal = "Bal. " & Format(Abs(RsParty!Curr_Bal), "0.00")
                LblPartyBal = LblPartyBal & IIf(RsParty!Curr_Bal > 0, " Cr", IIf(RsParty!Curr_Bal < 0, " Dr", ""))
                If Trim(txt(PayTo1).TEXT) = "" Then txt(PayTo1).TEXT = RsParty!Name
            End If
            
        Case AcHead
            If IsValid(txt(Index), "A/C Head") = False Then Cancel = True: Exit Sub
            
            With RsAcCode
                If .RecordCount = 0 Or (.EOF = True Or .BOF = True) Or txt(Index).TEXT = "" Then
                    txt(Index).TEXT = ""
                    txt(Index).Tag = ""
                    mRepName = ""
                Else
                    txt(Index).TEXT = !Name
                    txt(Index).Tag = !Code
                    lblAcBal = "Bal. " & Format(Abs(!Curr_Bal), "0.00")
                    lblAcBal = lblAcBal & IIf(!Curr_Bal > 0, " Cr", IIf(!Curr_Bal < 0, " Dr", ""))
                    mRepName = XNull(!ChequeReportName)
                End If
            End With
            
        Case VDate
            If Len(Trim(txt(VDate).TEXT)) = 0 Then
                txt(VDate).TEXT = PubLoginDate
            Else
                txt(Index).TEXT = RetDate(txt(Index))
            End If
            Cancel = Not CheckFinYear(txt(VDate))
            If Cancel = False Then txt(VType).SetFocus
            
        Case Chq_Date, Clg_Date
            txt(Index).TEXT = RetDate(txt(Index))
            
        Case Chq_No
            If Val(txt(Index)) > 0 Then
                txt(Index).TEXT = Format(Val(txt(Index).TEXT), String(6, "0"))
            Else
                txt(Index).TEXT = ""
            End If
        Case Amount
            txt(Index).TEXT = Format(Val(txt(Index).TEXT), "0.00")
       
    End Select
    Ctrl_validate txt(Index)
    Set Rst = Nothing

    Exit Sub
ErrLoop:
    Call CheckError
End Sub

'*** Fuctions ********
Private Sub BlankText()
    Dim I As Byte
    For I = 0 To txt.Count - 1
        txt(I).TEXT = ""
        txt(I).Tag = ""
    Next I
    mVType = "": mNCat = "": mSearchCode = ""
    
    txt(AcPayeeCheque).TEXT = "No"
End Sub
Private Sub MoveRec()
Dim Rst As Recordset
Dim RsTemp As ADODB.Recordset
Dim I As Integer
On Error GoTo error1
    Call BlankText
    
    If Master.RecordCount > 0 Then
        mSearchCode = Master!SearchCode
        
        If InStr(Me.TopCtrl1.Tag, "E") <> 0 Then Me.TopCtrl1.tEdit = True
        mqry = "Select P.*, Vt.Description As VType_Description, Vt.NCat, Sg.Name As AcName " & _
                " From (((Payment As P  " & _
                " Left Join Voucher_Type As Vt On P.V_Type = Vt.V_Type " & _
                " Left Join SubGroup As Sg On P.AcCode = Sg.SubCode))) " & _
                " Where P.DocId = '" & Master!SearchCode & "'"
        Set RsTemp = GCn.Execute(mqry)
        If RsTemp.RecordCount > 0 Then
            DocID = RsTemp!DocID
            txt(TxtDocID).TEXT = RsTemp!DocID
            LblDiv.CAPTION = "Division : " & left(RsTemp!DocID, 1)
            LblSite.CAPTION = "Site Code : " & mID(RsTemp!Site_Code, 1, 1)
            txt(SiteCode).Tag = mID(RsTemp!Site_Code, 2, 1)
            txt(SiteCode).TEXT = GCn.Execute("select site_desc from site where site_code = '" & txt(SiteCode).Tag & "'").Fields(0).Value
            LblUser = IIf(Not IsNull(RsTemp!AddDate), "Add By : " & XNull(RsTemp!AddBy) & "  Dated : " & XNull(RsTemp!AddDate), "") & IIf(Not IsNull(RsTemp!ModifyDate), "     Modify By : " & XNull(RsTemp!ModifyBy) & "  Dated : " & XNull(RsTemp!ModifyDate), "")
            LblVPrefix.CAPTION = DeCodeDocID(RsTemp!DocID, Document_Prefix)
            txt(SerialNo).TEXT = RsTemp!V_NO
            txt(VDate).TEXT = RsTemp!V_DATE
            mVType = RsTemp!V_Type
            mNCat = RsTemp!NCat
            txt(VType).Tag = mVType
            txt(VType).TEXT = XNull(RsTemp!VType_Description)
            '*** A/c Posting Status
            txt(AcPostByName) = IIf(IsNull(RsTemp!AcPostByU_Name), "", RsTemp!AcPostByU_Name)
            txt(AcPostDate) = IIf(IsNull(RsTemp!AcPostByU_EntDt), "", RsTemp!AcPostByU_EntDt)
            '***
            txt(Party).Tag = RsTemp!PartyCode
            If txt(Party).Tag <> "" Then
                Set Rst = New Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "select NAME,Curr_Bal,PanNo,ITWARD_NO from SubGroup where Subcode = '" & txt(Party).Tag & "'", GCn, adOpenDynamic, adLockBatchOptimistic
                txt(Party) = Rst!Name
                LblPartyBal = "Bal. " & Format(Abs(Rst!Curr_Bal), "0.00")
                LblPartyBal = LblPartyBal & IIf(Rst!Curr_Bal > 0, " Cr", IIf(Rst!Curr_Bal < 0, " Dr", ""))
            Else
                txt(Party).TEXT = ""
            End If
            
            txt(AcHead).Tag = XNull(RsTemp!AcCode)
            txt(AcHead).TEXT = XNull(RsTemp!AcName)
            
            txt(Chq_No).TEXT = IIf(IsNull(RsTemp!Chq_No), "", RsTemp!Chq_No)
            txt(Chq_Date).TEXT = Format(XNull(RsTemp!Chq_Date), "dd/MMM/yyyy")
            txt(Clg_Date).TEXT = Format(XNull(RsTemp!Clg_Date), "dd/MMM/yyyy")
            txt(Narr).TEXT = IIf(IsNull(RsTemp!Narration), "", RsTemp!Narration)
            txt(Amount).TEXT = Format(RsTemp!Amount, "0.00")
            
            txt(PayTo1).TEXT = XNull(RsTemp!PayTo1)
            txt(PayTo2).TEXT = XNull(RsTemp!PayTo2)
            txt(AcPayeeCheque).TEXT = IIf(VNull(RsTemp!AcPayeeCheque), "Yes", "No")
        End If
    End If
    Grid_Hide
    Set Rst = Nothing
    Exit Sub
error1:
    CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
    Dim I As Integer
    For I = 0 To txt.Count - 1
        txt(I).Enabled = Enb
        txt(I).ForeColor = CtrlFColOrg
    Next
    
    txt(TxtDocID).Enabled = False
    
    If TopCtrl1.TopText2 = "Edit" Then
        txt(SiteCode).Enabled = False
        txt(VDate).Enabled = False
        txt(SerialNo).Enabled = False
        txt(VType).Enabled = False
    End If
End Sub

Private Sub Grid_Hide()
    If DGSite.Visible = True Then DGSite.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If DGVType.Visible = True Then DGVType.Visible = False
    If DGVno.Visible = True Then DGVno.Visible = False
    If DgAcCode.Visible = True Then DgAcCode.Visible = False
End Sub
 
'
'************ PRINTING CODE ******************
Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case FromVno, ToVno
        RsVno.Close
        RsVno.Open "Select v_no as code from Payment where right(Payment.Site_Code,1)='" & txtPrint(SiteCode1).Tag & "' and  Payment.V_Type='" & txtPrint(VType1).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        Set DGVno.DataSource = RsVno
        If txtPrint(Index).TEXT <> RsVno!Code Then
            RsVno.MoveFirst
            RsVno.FIND "code ='" & txtPrint(Index).TEXT & "'"
        End If
        If Index = ToVno Then DGVno.Tag = "1" Else DGVno.Tag = "2"
    Case VType1
        If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or txtPrint(Index).TEXT = "" Then Exit Sub
        If txtPrint(Index).TEXT <> RsVType!Name Then
            RsVType.MoveFirst
            RsVType.FIND "Name ='" & txtPrint(Index).TEXT & "'"
        End If
    Case SiteCode1
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Then Exit Sub
        If txtPrint(Index).TEXT = "" Then
            RsSite.MoveFirst
            RsSite.FIND "code ='" & PubSiteCode & "'"
            txtPrint(Index).Tag = RsSite!Code
            txtPrint(Index).TEXT = RsSite!Name
        Else
            If txtPrint(Index).TEXT <> RsSite!Name Then
                RsSite.MoveFirst
                RsSite.FIND "name ='" & txtPrint(Index).TEXT & "'"
            End If
        End If
End Select
End Sub

Private Sub TxtPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case FromVno, ToVno
        DGridTxtKeyDown DGVno, txtPrint, Index, RsVno, KeyCode, False, 0
    Case VType1
        DGridTxtKeyDown DGVType, txtPrint, Index, RsVType, KeyCode, False, 1
    Case SiteCode1
        DGridTxtKeyDown DGSite, txtPrint, Index, RsSite, KeyCode, False, 1
End Select
If DGVType.Visible = False And DGSite.Visible = False And DGVno.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> SiteCode1 Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TxtPrint_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(keyascii)
Select Case Index
    Case FromVno, ToVno
        If DGVno.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsVno, keyascii, "Code"
    Case VType1
        If DGVType.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsVType, keyascii, "name"
    Case SiteCode1
        If DGSite.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsSite, keyascii, "Name"
End Select

'KeyAscii = RetDGKeyAscii()
End Sub

Private Sub TxtPrint_LostFocus(Index As Integer)
  Ctrl_validate txtPrint(Index)
End Sub

Private Sub TxtPrint_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case ToVno, FromVno
        If RsVno.RecordCount = 0 Or (RsVno.EOF = True Or RsVno.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(Index).TEXT = ""
        Else
            txtPrint(Index).TEXT = RsVno!Code
        End If
    Case VType1
        If IsValid(txtPrint(Index), "Voucher Type") = False Then Cancel = True: Exit Sub
        If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(Index).TEXT = ""
            txtPrint(Index).Tag = ""
        Else
            txtPrint(Index).TEXT = RsVType!Name
            txtPrint(Index).Tag = RsVType!Code
        End If
    Case SiteCode1
        If IsValid(txtPrint(Index), "Site Code") = False Then Cancel = True: Exit Sub
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(Index).TEXT = ""
            txtPrint(Index).Tag = ""
        Else
            txtPrint(Index).TEXT = RsSite!Name
            txtPrint(Index).Tag = RsSite!Code
        End If
End Select
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
On Error GoTo ERRORHANDLER
'GSQL = "SELECT SG.NamePrefix, SG.Name as PartyName,SG.FPrefix,SG.FName,SG.Add1,SG.Add2,SG.Add3,SG.PANNo,SG.ITWARD_NO, City.CityName, SG1.Name as AcName,Voucher_Type.Description,  Payment.*, Syctrl.SprMoneyRectFooter,model_Grp.ModelGrp_Name as model, VO.Ord_No, VO.Ord_Date,VO.Chassis,Veh_Stock.EngineNo, CF.FinName " & _
    " FROM (((((((((Payment LEFT JOIN " & FaTable("Voucher_Type") & " ON Rect.V_Type = Voucher_Type.V_Type) " & _
    " LEFT JOIN SubGroup SG on Rect.PartyCode = SG.SubCode ) " & _
    " LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
    " LEFT JOIN SubGroup SG1 ON Rect.AcCode = SG1.SubCode) " & _
    " LEFT JOIN Veh_Order VO ON Rect.Ord_DocId = VO.OrdDocId) " & _
    " LEFT JOIN Veh_Stock ON Veh_Stock.Sal_DocId = VO.Inv_DocId) " & _
    " LEFT JOIN Model on Model.Model = VO.Model ) " & _
    " LEFT JOIN Model_Grp ON Model.Grp_Code = Model_Grp.ModelGrp_Code) " & _
    " LEFT JOIN ContractFinance CF ON VO.FB_CODE = CF.FinCode) " & _
    " LEFT JOIN Syctrl ON Syctrl.LinkTable  > Rect.U_AE   " & _
    " where Rect.Docid ='" & Master!SearchCode & "'"

Select Case Index
    Case PScreen
        If OptVoucher Then
            Set rpt = rdApp.OpenReport(PubFaReportPath + "\FaJVCHR.RPT")
            FaVoucherPrintingModule Me, rpt, mSearchCode
            Set rpt = Nothing
        Else
            Call WindowsPrint(Index, GSQL, True)
        End If
            FrmPrn.Visible = False
    Case PWindows
        If OptVoucher Then
            Set rpt = rdApp.OpenReport(PubFaReportPath + "\FaJVCHR.RPT")
            FaVoucherPrintingModule Me, rpt, mSearchCode
            Set rpt = Nothing
        Else
            Call WindowsPrint(Index, GSQL, False)
        End If
        FrmPrn.Visible = False
    Case PDos
        If OptVoucher Then
            PrintVoucherPlain mSearchCode
        Else
            MsgBox "Speed Print for Cheque is not Avaialable!"
        End If
        FrmPrn.Visible = False
    Case PSetUp
        'mRepName = IIf(OptPlain.Value = True, "SprCustRect", "SprCustRect")
        Call PrinerSetUp
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
End If
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub WindowsPrint(Index As Integer, mqry As String, ScreenPrint As Boolean)
On Error GoTo ERRORHANDLER
Dim I As Integer
Dim mDivSName$, RepTitle$
Dim Rst As ADODB.Recordset
Dim RST1 As ADODB.Recordset

    mDivSName = IIf(PubDivSName = "", "", "-" & PubDivSName & " ")
      
        
    mqry = " Select P.PayTo1, P.PayTo2, P.Amount, P.Chq_Date As ChqDate, P.Chq_No, P.AcPayeeCheque, " & _
            " Sg.Name As BankName, Sg.ChequeReportName " & _
            " From ((Payment P " & _
            " Left Join SubGroup Sg On P.AcCode = Sg.SubCode))" & _
            " Where P.DocId='" & mSearchCode & "'"
        
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mqry), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    mRepName = XNull(Rst!ChequeReportName)          '"Cheque_JalGawnBank"
    If Trim(mRepName) = "" Then MsgBox "Please Define Report Name for " & Rst!BankName & " in ""A/c Ledger Entry""!...": Exit Sub
    
    If VNull(Rst!AcPayeeCheque) Then
        RepTitle = "ACCOUNT PAYEE ONLY"
    Else
        RepTitle = ""
    End If
    
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords

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
                rpt.FormulaFields(I).TEXT = "'" & RepTitle & "'"
        End Select
    Next
    
    If ScreenPrint Then
        Call Report_View(rpt, RepTitle, , True)
    Else
        rpt.PrintOut False
    End If
    
'    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
'        GCn.Execute "update Payment set Printed = 1 where DocID='" & mSearchCode & "'"
'    End If
    Set RST1 = Nothing
    'Set rpt = Nothing

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

Private Sub PrintVoucherPlain(mDocId As String)
    Dim fob As New FileSystemObject

    Dim mCITYNAME As String, connectionId, mPAR_NAME As String
    Dim mPAR_ADDR1 As String, mPAR_ADDR2 As String, RST1 As ADODB.Recordset
    Dim connectionIdScr As String
    Dim RstVType As ADODB.Recordset
    Dim strAmount As String
    Dim I As Integer
    Dim DrAmt As Double
    Dim Remarks As String
    Dim v_Prefix As String
    Dim rstVoucherPrint As ADODB.Recordset
    Dim rstDebitAc As ADODB.Recordset
    Dim rstCreditAc As ADODB.Recordset
    Dim DebitAc As String
    Dim CreditAc As String
    Dim debitAmt As String
    Dim CreditAmt As String
    Dim Remark As String
    Dim TotCrAmt As Double
    
    Dim RowNo As Integer
    Dim RstDate As ADODB.Recordset
    Dim VDate As Date
    Dim PrintVno As String
    Dim RstEnviro As ADODB.Recordset
    Dim NarrPart1$, NarrPart2$, NarrPart3$, NarrPart4$, NarrPart5$
    On Error GoTo ERRORHANDLER


    v_Prefix = GCn.Execute("Select Max(Prefix) From Voucher_Prefix ").Fields(0).Value


    RowNo = 0
    Set rstVoucherPrint = GCn.Execute("Select Ledger.* ,LM.Narration as ComNarr, Vt.Description As VType_Desc " & _
                                      "From ((Ledger " & _
                                      "Left Join LedgerM LM on Ledger.DocID=LM.DocID)  " & _
                                      "Left Join Voucher_Type Vt on Ledger.V_Type=Vt.V_Type)  " & _
                                      "Where Ledger.DocID = '" & mDocId & "'")
    
    Set RstDate = GCn.Execute("Select Ledger.* ,LM.Narration as ComNarr from Ledger Left Join LedgerM LM on Ledger.DocID=LM.DocID Where Ledger.DocId='" & mDocId & "'")
    If RstDate.EOF = False Then
        VDate = XNull(RstDate!V_DATE)
    End If
    mPAR_NAME = ""
    mPAR_ADDR1 = ""
    mPAR_ADDR2 = ""
    mCITYNAME = ""
    
    
    If fob.FileExists("C:\Repprint.TXT") = False Then
        fob.CreateTextFile ("C:\RepPrint.TXT")
    End If
    Close #1
    Open "C:\RepPrint.TXT" For Output As #1
    
    Print #1, Chr(27) + Chr(67) + Chr(36): RowNo = RowNo + 1
    Print #1, PRN_TIT(PubComp_Name, "A", 80): RowNo = RowNo + 1
    Print #1, Chr(27) + Chr(69) + Space((68 - Len(PubComp_Add)) / 2) + PubComp_Add + ", " + PubComp_City + Chr(27) + Chr(70): RowNo = RowNo + 1
    Print #1, PRN_TIT(XNull(PubComp_Contact), "B", 80): RowNo = RowNo + 1
    Print #1, "": RowNo = RowNo + 1
    
    
    
    If rstVoucherPrint.EOF = True Then MsgBox "No Record to Print.": Exit Sub
    'If VType = "CN" Then
    Print #1, PRN_TIT(rstVoucherPrint!VType_Desc & " Voucher", "B", 80): RowNo = RowNo + 1
    
    Print #1, Space(50) + Chr(27) + Chr(69) + "Voucher No   : " + CStr(rstVoucherPrint!V_NO) + Chr(27) + Chr(70): RowNo = RowNo + 1
    
    Print #1, Space(50) + Chr(27) + Chr(69) + "Voucher Date : " + CStr(VDate) + Chr(27) + Chr(70): RowNo = RowNo + 1
    Print #1, "": RowNo = RowNo + 1
    Dim j As Integer
    
    
    
    
    Print #1, "-------------------------------------------------------------------------------": RowNo = RowNo + 1
    Print #1, "Srl." + "PARTICULARS                                        " + "       DR" + "            CR": RowNo = RowNo + 1
    Print #1, "-------------------------------------------------------------------------------": RowNo = RowNo + 1
    
I = 1
While rstVoucherPrint.EOF = False
    If rstVoucherPrint!AmtDr > 0 Then
        Set rstDebitAc = GCn.Execute("Select Name from SubGroup Where SubCode='" & rstVoucherPrint!SubCode & "'")
        DebitAc = XNull(rstDebitAc!Name)
        debitAmt = Format(VNull(rstVoucherPrint!AmtDr), "0.00")
        VDate = rstVoucherPrint!V_DATE

        NarrPart1 = "": NarrPart2 = "": NarrPart3 = "": NarrPart4 = "": NarrPart5 = ""
        Remark = XNull(rstVoucherPrint!ComNarr)
            Print #1, Space(1) + PSTR(CStr(I), 3) + SETW(DebitAc, 35) + Space(16) & SETN(debitAmt, 9): RowNo = RowNo + 1
            If XNull(rstVoucherPrint!Narration) <> "" Then
                NarrPart1 = mID(XNull(rstVoucherPrint!Narration), 1, 50)
                NarrPart2 = mID(XNull(rstVoucherPrint!Narration), 51, 50)
                NarrPart3 = mID(XNull(rstVoucherPrint!Narration), 101, 50)
                NarrPart4 = mID(XNull(rstVoucherPrint!Narration), 151, 50)
                NarrPart5 = mID(XNull(rstVoucherPrint!Narration), 201, 50)
                
                If NarrPart1 <> "" Then Print #1, Space(4) + NarrPart1: RowNo = RowNo + 1
                If NarrPart2 <> "" Then Print #1, Space(4) + NarrPart2: RowNo = RowNo + 1
                If NarrPart3 <> "" Then Print #1, Space(4) + NarrPart3: RowNo = RowNo + 1
                If NarrPart4 <> "" Then Print #1, Space(4) + NarrPart4: RowNo = RowNo + 1
                If NarrPart5 <> "" Then Print #1, Space(4) + NarrPart5: RowNo = RowNo + 1
            End If

    Else
        Set rstCreditAc = GCn.Execute("Select Name from SubGroup Where SubCode='" & rstVoucherPrint!SubCode & "'")
        CreditAc = XNull(rstCreditAc!Name)
        CreditAmt = Format(VNull(rstVoucherPrint!AmtCr), "0.00")
        VDate = rstVoucherPrint!V_DATE
        Remark = XNull(rstVoucherPrint!ComNarr)
        Print #1, "": RowNo = RowNo + 1
        Print #1, Space(1) + PSTR(CStr(I), 3) + SETW(CreditAc, 35) + Space(1) + Space(10) + Space(5) + Space(14) + SETN(CreditAmt, 9): RowNo = RowNo + 1
        'If rstVoucherPrint!Narration <> "" Then Print #1, Space(4) + XNull(rstVoucherPrint!Narration): RowNo = RowNo + 1
            If XNull(rstVoucherPrint!Narration) <> "" Then
                NarrPart1 = mID(XNull(rstVoucherPrint!Narration), 1, 50)
                NarrPart2 = mID(XNull(rstVoucherPrint!Narration), 51, 50)
                NarrPart3 = mID(XNull(rstVoucherPrint!Narration), 101, 50)
                NarrPart4 = mID(XNull(rstVoucherPrint!Narration), 151, 50)
                NarrPart5 = mID(XNull(rstVoucherPrint!Narration), 201, 50)
                
                If NarrPart1 <> "" Then Print #1, Space(4) + NarrPart1: RowNo = RowNo + 1
                If NarrPart2 <> "" Then Print #1, Space(4) + NarrPart2: RowNo = RowNo + 1
                If NarrPart3 <> "" Then Print #1, Space(4) + NarrPart3: RowNo = RowNo + 1
                If NarrPart4 <> "" Then Print #1, Space(4) + NarrPart4: RowNo = RowNo + 1
                If NarrPart5 <> "" Then Print #1, Space(4) + NarrPart5: RowNo = RowNo + 1
            End If
        
    End If
    I = I + 1
    TotCrAmt = TotCrAmt + Val(rstVoucherPrint!AmtCr)
    
    rstVoucherPrint.MoveNext
Wend
'    If VType <> "JV" Then
'        Print #1, Space(4) + "( " + Remark + " )"
'    Else
        If Len(Remark) > 0 Then
            Print #1, Space(4) + "( " + left(Remark, 35) + " )"
            If Len(mID(Remark, 36)) > 0 Then
                Print #1, Space(4) + "( " + mID(Remark, 36, 35) + " )"
            End If
            If Len(mID(Remark, 72)) > 0 Then
                Print #1, Space(4) + "( " + mID(Remark, 71, 35) + " )"
            End If
        End If
'    End If
    
    'Print #1, Space(4) + "( " + left(Remark, 50) + " )"
    Print #1, "": RowNo = RowNo + 1
    Print #1, "-------------------------------------------------------------------------------": RowNo = RowNo + 1
    Print #1, Space(20) + "TOTAL: " + Space(28) + SETN(CStr(Format(TotCrAmt, "0.00")), 9) + Space(5) + SETN(CStr(Format(TotCrAmt, "0.00")), 9): RowNo = RowNo + 1
    Print #1, "-------------------------------------------------------------------------------": RowNo = RowNo + 1
    
    strAmount = IIf(Val(TotCrAmt) > 0, "Rs. " + ntow(Val(TotCrAmt), "", "Ps."), "")
    Print #1, Space(8) + SETW(strAmount, 70): RowNo = RowNo + 1
    'Print #1, ""
    Print #1, "": RowNo = RowNo + 1
    Print #1, Space(5); "E O & E " + Space(37) + "FOR " + PubComp_Name: RowNo = RowNo + 1
    Print #1, "": RowNo = RowNo + 1
    Print #1, "": RowNo = RowNo + 1
    If StrCmp(left(PubComp_Name, 4), "comm") Then
        Print #1, "Recd By (Sign)" & Space(2) & "Prepared By : " & pubUName & Space(2) & "Cashier" & Space(2) & "Accountant" & Space(2) & "Partner/Manager/Director"
        RowNo = RowNo + 1
    End If
    Print #1,
    If RowNo < 34 Then
        For I = RowNo To 34
            Print #1, ""
        Next
        Chr (12)
    Else
        Print #1, Chr(12)
    End If
Close #1
'If MsgBox("Report on Screen ? ", vbYesNo + vbDefaultButton2, "Printing") = vbYes Then
'    Open "C:\Repprint.BAT" For Output As #1
'        Print #1, "Edit C:\Repprint.TXT"
'    Close #1
'    connectionIdScr = Shell("C:\reptmp\ScrBill.PIF", vbMaximizedFocus)
'Else
    Open "C:\RepPrint.BAT" For Output As #1
        Print #1, "TYPE C:\RepPrint.TXT>" & PubFaDosPort
    Close #1
    If PubRunPIF = "Y" Then
        connectionId = Shell("C:\reptmp\SBILL.pif", vbHide)
    Else
        connectionId = Shell("C:\RepPrint.bat", vbHide)
    End If
'End If


Set RST1 = Nothing
Set connectionId = Nothing
Exit Sub
ERRORHANDLER:     MsgBox err.Description, vbCritical, Me.CAPTION
End Sub



