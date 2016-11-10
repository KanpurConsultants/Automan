VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmJobBooking 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Service Booking"
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
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
      Height          =   210
      Index           =   0
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   9
      Top             =   2580
      Width           =   4995
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
      Height          =   210
      Index           =   23
      Left            =   8220
      MaxLength       =   15
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   3315
      Width           =   1575
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
      Left            =   3930
      TabIndex        =   48
      Top             =   4905
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
         Picture         =   "frmJobBooking.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   58
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
         Picture         =   "frmJobBooking.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmJobBooking.frx":0678
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
         TabIndex        =   56
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmJobBooking.frx":0982
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
         TabIndex        =   55
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmJobBooking.frx":0C8C
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
         Top             =   0
         Width           =   4695
      End
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
   Begin MSDataGridLib.DataGrid DGService 
      Height          =   2730
      Left            =   7635
      Negotiate       =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   7020
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
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
      Height          =   210
      Index           =   1
      Left            =   1680
      MaxLength       =   14
      TabIndex        =   1
      Top             =   660
      Width           =   1470
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
      Height          =   450
      Index           =   21
      Left            =   1680
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   3780
      Width           =   4995
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
      Height          =   210
      Index           =   3
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   2
      Top             =   900
      Width           =   3135
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
      Height          =   210
      Index           =   8
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   8
      Top             =   2340
      Width           =   4995
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
      Height          =   210
      Index           =   11
      Left            =   5310
      MaxLength       =   10
      TabIndex        =   12
      Top             =   2820
      Width           =   1365
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
      Height          =   210
      Index           =   10
      Left            =   3510
      MaxLength       =   25
      TabIndex        =   11
      Top             =   2820
      Width           =   1440
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
      Height          =   210
      Index           =   22
      Left            =   9225
      TabIndex        =   23
      Top             =   1020
      Width           =   1470
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
      Height          =   285
      Index           =   14
      Left            =   8625
      MaxLength       =   20
      TabIndex        =   15
      Top             =   2805
      Visible         =   0   'False
      Width           =   1140
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
      Height          =   210
      Index           =   7
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   7
      Top             =   2100
      Width           =   4995
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
      Height          =   210
      Index           =   4
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1380
      Width           =   4995
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
      Height          =   210
      Index           =   20
      Left            =   5310
      TabIndex        =   21
      Text            =   "28-APR-2002"
      Top             =   3540
      Width           =   1365
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
      Height          =   285
      Index           =   12
      Left            =   8625
      MaxLength       =   15
      TabIndex        =   13
      Top             =   2505
      Visible         =   0   'False
      Width           =   1140
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
      Height          =   285
      Index           =   13
      Left            =   10980
      MaxLength       =   10
      TabIndex        =   14
      Top             =   2505
      Visible         =   0   'False
      Width           =   870
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
      Height          =   210
      Index           =   19
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   20
      Top             =   3540
      Width           =   2130
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
      Height          =   210
      Index           =   9
      Left            =   1905
      MaxLength       =   25
      TabIndex        =   10
      Top             =   2820
      Width           =   1245
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
      Height          =   285
      Index           =   15
      Left            =   10980
      MaxLength       =   10
      TabIndex        =   16
      Top             =   2805
      Visible         =   0   'False
      Width           =   870
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
      Height          =   210
      Index           =   2
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1140
      Width           =   3135
   End
   Begin VB.TextBox txt 
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
      ForeColor       =   &H00C000C0&
      Height          =   210
      Index           =   18
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   19
      Top             =   3300
      Visible         =   0   'False
      Width           =   1470
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
      Height          =   210
      Index           =   6
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1860
      Width           =   4995
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
      Height          =   210
      Index           =   5
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1620
      Width           =   4995
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
      Height          =   210
      Index           =   16
      Left            =   1680
      TabIndex        =   17
      Top             =   3060
      Width           =   1470
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
      Height          =   210
      Index           =   17
      Left            =   5625
      MaxLength       =   8
      TabIndex        =   18
      Text            =   "99999999"
      Top             =   3060
      Width           =   1050
   End
   Begin MSDataGridLib.DataGrid DGModel 
      Height          =   1155
      Left            =   45
      Negotiate       =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   7260
      Visible         =   0   'False
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   2037
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
         DataField       =   "code"
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
         DataField       =   "ListName"
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
            DividerStyle    =   3
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   9000
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGBook 
      Height          =   2520
      Left            =   1200
      Negotiate       =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6975
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
         DataField       =   "PhoneResi"
         Caption         =   "Phone (R)"
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
            ColumnWidth     =   2715.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2459.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2234.835
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1230.236
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   2730
      Left            =   10230
      Negotiate       =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   5280
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
      ColumnCount     =   1
      BeginProperty Column00 
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
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Date*"
      BeginProperty Font 
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
      Left            =   150
      TabIndex        =   76
      Top             =   3075
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Deposit"
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
      Index           =   41
      Left            =   150
      TabIndex        =   75
      Top             =   3330
      Visible         =   0   'False
      Width           =   1440
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
      Left            =   150
      TabIndex        =   74
      Top             =   1875
      Width           =   690
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
      Left            =   150
      TabIndex        =   73
      Top             =   1410
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone No."
      BeginProperty Font 
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
      Left            =   150
      TabIndex        =   72
      Top             =   2850
      Width           =   1230
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
      Left            =   150
      TabIndex        =   71
      Top             =   1140
      Width           =   600
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
      Left            =   150
      TabIndex        =   70
      Top             =   3570
      Width           =   1230
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
      Left            =   150
      TabIndex        =   69
      Top             =   1650
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Regn No."
      BeginProperty Font 
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
      Left            =   150
      TabIndex        =   68
      Top             =   660
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No."
      BeginProperty Font 
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
      Left            =   150
      TabIndex        =   67
      Top             =   930
      Width           =   1005
   End
   Begin VB.Label Label3 
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
      Index           =   9
      Left            =   150
      TabIndex        =   66
      Top             =   3825
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City Name"
      BeginProperty Font 
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
      Left            =   150
      TabIndex        =   65
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job No. "
      BeginProperty Font 
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
      Left            =   7440
      TabIndex        =   63
      Top             =   3323
      Width           =   690
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
      Left            =   5040
      TabIndex        =   45
      Top             =   2828
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
      Left            =   3240
      TabIndex        =   44
      Top             =   2828
      Width           =   270
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
      Index           =   4
      Left            =   1635
      TabIndex        =   43
      Top             =   2828
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date && Time"
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
      Index           =   0
      Left            =   7515
      TabIndex        =   42
      Top             =   1020
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   1500
      Left            =   7350
      Top             =   450
      Width           =   4395
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
      Left            =   7815
      TabIndex        =   41
      Top             =   690
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
      Left            =   9540
      TabIndex        =   40
      Top             =   690
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Job No."
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
      Index           =   36
      Left            =   7425
      TabIndex        =   39
      Top             =   2535
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Date*"
      BeginProperty Font 
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
      Left            =   4035
      TabIndex        =   38
      Top             =   3570
      Width           =   1215
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   17
      Left            =   5190
      TabIndex        =   37
      Top             =   3570
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   16
      Left            =   8550
      TabIndex        =   36
      Top             =   2535
      Visible         =   0   'False
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   12
      Left            =   8550
      TabIndex        =   35
      Top             =   2835
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last KMs "
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
      Left            =   9885
      TabIndex        =   34
      Top             =   2835
      Visible         =   0   'False
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   14
      Left            =   10890
      TabIndex        =   33
      Top             =   2535
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Job Dt."
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
      Index           =   34
      Left            =   9885
      TabIndex        =   32
      Top             =   2535
      Visible         =   0   'False
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   9
      Left            =   10890
      TabIndex        =   31
      Top             =   2835
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Service"
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
      Index           =   27
      Left            =   7455
      TabIndex        =   30
      Top             =   2835
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label LblAmt 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   11220
      TabIndex        =   29
      Top             =   1605
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Advance Deposit For Date"
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
      Index           =   22
      Left            =   7515
      TabIndex        =   28
      Top             =   1605
      Width           =   2715
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
      Left            =   11490
      TabIndex        =   27
      Top             =   1335
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total No. of Vehicle for Service Date"
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
      Left            =   7515
      TabIndex        =   25
      Top             =   1335
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Srl No*"
      BeginProperty Font 
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
      Left            =   4050
      TabIndex        =   24
      Top             =   3105
      Width           =   1380
   End
End
Attribute VB_Name = "frmJobBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim MyIndex As Byte
Dim RSBook As ADODB.Recordset
Dim RsModel As ADODB.Recordset
Dim RsServ As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim RsCity As ADODB.Recordset
Dim DocID As String


Private Const City As Byte = 0
Private Const VehRegNo As Byte = 1
Private Const Model As Byte = 2
Private Const Chassis As Byte = 3
Private Const Engine As Byte = 4
Private Const OwnerName As Byte = 5
Private Const Address1 As Byte = 6
Private Const Address2 As Byte = 7
Private Const Address3 As Byte = 8
Private Const PhoneOff As Byte = 9
Private Const PhoneResi As Byte = 10
Private Const Mobile As Byte = 11
Private Const LastJobNo As Byte = 12
Private Const LastJobDt As Byte = 13
Private Const LastSrv As Byte = 14
Private Const LastKMS As Byte = 15
Private Const BookDate As Byte = 16
Private Const BookSrl As Byte = 17
Private Const AdvAmt As Byte = 18
Private Const SrvType As Byte = 19
Private Const SrvDate As Byte = 20
Private Const Remarks As Byte = 21
Private Const EntryDate As Byte = 22
Private Const JobDocID As Byte = 23

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String


Private Sub DGBook_Click()
If RSBook.RecordCount > 0 Then
    Select Case MyIndex
        Case VehRegNo, Chassis, OwnerName
            txt(VehRegNo).TEXT = XNull(RSBook!RegNo)
            txt(Model).Tag = XNull(RSBook!Model)
            txt(Model).TEXT = XNull(RSBook!Model)
            txt(Chassis).TEXT = XNull(RSBook!Chassis)
            txt(OwnerName).TEXT = XNull(RSBook!Name)
            txt(Address1).TEXT = XNull(RSBook!Add1)
            txt(Address2).TEXT = XNull(RSBook!Add2)
            txt(Address3).TEXT = XNull(RSBook!Add3)
            txt(PhoneResi).TEXT = XNull(RSBook!PhoneResi)
            txt(PhoneOff).TEXT = XNull(RSBook!PhoneOff)
            txt(Mobile).TEXT = XNull(RSBook!Mobile)
    End Select
End If
txt(MyIndex).SetFocus
DGBook.Visible = False
End Sub

Private Sub DGModel_Click()
If RsModel.RecordCount > 0 Then
    txt(Model).TEXT = RsModel!Code
    txt(Model).Tag = RsModel!Code
End If
txt(Model).SetFocus
DGModel.Visible = False
End Sub

Private Sub DGService_Click()
If RsServ.RecordCount > 0 Then
    txt(SrvType).Tag = RsServ!Code
    txt(SrvType).TEXT = RsServ!Name
End If
DGService.Visible = False
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
On Error GoTo ELoop
'Dim RstMain As Recordset
    
    '' pending points
    '' 1. Financial Year Date Checking with Booking Date-NR
    '' 2. Service Date should be higher or equal to booking date -- REq , done
    '' 3. Display of Last Job Card Details (label already exist on form  -- NR
    '' 4. Check any open job card in workshop for the same vehicle -- NR
    'concat(po_ycode,concat(Origin,PO_NO))
    WinSetting Me
    TopCtrl1.Tag = PubUParam:    Ini_Grid
    txt(BookDate).Tag = date
'         End If
    
    Dim SiteCond As String
    SiteCond = " Where  Book_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and left(Jb.site_code, 1) ='" & PubSiteCode & "'"
    End If
    
    
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "select JB.Div_Code+JB.Site_Code+right(space(8)+" & cCStr("JB.Book_no") & ",8) as SearchCode, JB.Div_Code,JB.Site_Code,JB.Book_No from Job_Booking as JB " & SiteCond & " order by JB.Book_Date desc, JB.Div_Code desc,JB.Site_Code desc,JB.Book_no desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 JB.Div_Code+JB.Site_Code+right(space(8)+" & cCStr("JB.Book_no") & ",8) as SearchCode, JB.Div_Code,JB.Site_Code,JB.Book_No from Job_Booking as JB  " & SiteCond & " order by JB.Book_Date desc, JB.Div_Code desc,JB.Site_Code desc,JB.Book_no desc", GCn, adOpenDynamic, adLockOptimistic
    End If
'    Master.Open "select JB.Div_Code+JB.Site_Code+right(space(8)+CStr(JB.Book_no),8) as SearchCode,JB.*, Service_type.Serv_type,Service_type.Serv_desc from Job_Booking as JB left join Service_type on JB.Service_type=Service_type.serv_type order by JB.Div_Code,JB.book_no,JB.Site_Code", GCn, adOpenDynamic, adLockOptimistic
    
    Set RSBook = New ADODB.Recordset
    GSQL = "SELECT MODEL AS Code, RegNo, Chassis, Name, Div_Code, Site_Code, Book_No, Book_Date, Add1, Add2,  Add3, '' As CityName, PhoneOff, PhoneResi, Mobile, Model, Engine, Advance, ForServiceDate, Remarks, Service_Type from Job_Booking where Chassis not in (Select Chassis from HisCaRD)" & _
           "Union " & _
           "Select MODEL AS Code, RegNo, Chassis, Name,' ' as Div_Code, '  ' as Site_Code, '' as Book_No, '' as Book_Date, Add1, Add2, Add3, CityName, PhoneOff, PhoneResi, Mobile, Model, Engine, 0 as Advance,  '' as ForServiceDate, '' as Remarks, '' as Service_Type from HisCard Left Join City C On C.CityCode=HisCard.CityCode " & _
           "ORDER BY RegNo"
    
    With RSBook
        .CursorLocation = adUseClient
'        .Open "SELECT JB.MODEL AS CODE,JB.regno,JB.chassis,JB.name,JB.Div_Code,JB.Site_Code,JB.Book_No,JB.Book_Date, JB.Add1, JB.Add2, JB.Add3, JB.PhoneOff, JB.PhoneResi, JB.Mobile, JB.Model, JB.Engine, JB.Advance, JB.ForServiceDate, JB.Remarks, JB.Service_Type from job_booking as JB order by Job_Booking.BOOK_NO", GCn, adOpenDynamic, adLockOptimistic
        .Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGBook.DataSource = RSBook
    RSBook.Sort = "BOOK_NO"
    RSBook.Sort = "Regno"
      
    Set RsModel = New ADODB.Recordset
    RsModel.CursorLocation = adUseClient
    RsModel.Open "Select MODEL as code,model as name ,model_desc as Listname FROM Model Order by MODEL", GCn, adOpenDynamic, adLockOptimistic
    Set DGModel.DataSource = RsModel
    RsModel.Sort = "code"
    
    
    Set RsCity = GCn.Execute("Select CityCode As Code, CityName As Name From City Order By CityName")
    Set DGCity.DataSource = RsCity
    
    
    Set RsServ = New ADODB.Recordset
    RsServ.CursorLocation = adUseClient
    RsServ.Open "Select Serv_type as code,serv_desc as name FROM service_type Order by Serv_DESC", GCn, adOpenDynamic, adLockOptimistic
    Set DGService.DataSource = RsServ
    RsServ.Sort = "name"
'    If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
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
MasterFormExit = False
    Set RSBook = Nothing
    Set Master = Nothing
    Set RsModel = Nothing
    Set RsServ = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    Call BlankText
    'Txt(BookSrl).TEXT = GCn.Execute("select iif(isnull(max(book_no)),0,max(book_no))+1 from job_booking").Fields(0)
    txt(BookSrl).TEXT = GCn.Execute("select " & vIsNull("Max(book_no)", "0") & "+1 from job_booking").Fields(0)
    txt(BookDate).TEXT = Format(txt(BookDate).Tag, "dd/MMM/yyyy")
    txt(SrvDate).TEXT = Format(IIf(UCase(left(PubComp_Name, 3)) = "LMP", txt(BookDate).Tag, date + 1), "dd/MMM/yyyy")
    txt(VehRegNo).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant
If Master.RecordCount > 0 Then
    If txt(JobDocID).Tag <> "" Then
        MsgBox "JobCard Made!" & vbCrLf & "Delete denied.", vbInformation, "Validation Check": Exit Sub
    End If
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        GCn.BeginTrans
        'GCn.Execute ("delete from Job_Booking where div_code='" & Right(LblDiv.CAPTION, 1) & "' and " & cTrim("Book_no") & " = " & Trim(Txt(BookSrl)) & " and Site_code='" & Right(LblSite.CAPTION, 1) & "'")
        GCn.Execute ("delete from Job_Booking where div_code='" & Right(LblDiv.CAPTION, 1) & "' and " & cTrim("Book_no") & " = " & Trim(txt(BookSrl)) & " ")
        GCn.CommitTrans
        Master.Requery
        RSBook.Requery
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
        End If
        BUTTONS True, Me, Master, 0
        MoveRec
    End If
Else
    MsgBox "No Records To Delete!", vbInformation, "Information"
End If
eloop1:
    If err.NUMBER <> 0 Then
       GCn.RollbackTrans
       MsgBox err.Description, vbCritical, " Deletion Message"
    End If
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    txt(VehRegNo).Enabled = False
    txt(Model).Enabled = False
    txt(Chassis).Enabled = False
    txt(Engine).Enabled = False
'    txt(AdvAmt).SetFocus
    txt(BookDate).SetFocus
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
    SiteCond = " Where  Book_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and left(Jb.site_code, 1) ='" & PubSiteCode & "'"
    End If
    
    GSQL = "Select JB.Div_Code+JB.Site_Code+right(space(8)+ " & cCStr("JB.Book_no") & ",8) as SearchCode,JB.REGNO,JB.CHASSIS,JB.NAME,JB.PHONEOFF,JB.PHONERESI FROM JOB_BOOKING JB " & SiteCond & " order by JB.Div_Code,JB.Site_Code,JB.Book_no"
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
        Set Master = GCn.Execute("select JB.Div_Code+JB.Site_Code+right(space(8)+" & cCStr("JB.Book_no") & ",8) as SearchCode, JB.Div_Code,JB.Site_Code,JB.Book_No from Job_Booking as JB  Where JB.Div_Code+JB.Site_Code+right(space(8)+" & cCStr("JB.Book_no") & ",8) = '" & MyValue & "' order by JB.Book_Date desc, JB.Div_Code desc,JB.Site_Code desc,JB.Book_no desc")
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
        If MasterFormExit Then Unload Me: Exit Sub
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
    RsServ.Requery
    RsModel.Requery
    RSBook.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim mTrans As Boolean
    Dim DocIdHlp As String
    Dim LedgAry(1) As LedgRec, mNarr$, mResult As Byte
    
    Grid_Hide
    
    If IsValid(txt(Model), "Model") = False Then Exit Sub
    If IsValid(txt(BookDate), "Booking Date") = False Then Exit Sub
    If IsValid(txt(SrvType), "Service Type") = False Then Exit Sub
    If IsValid(txt(SrvDate), "Service Date") = False Then Exit Sub
    If IsValid(txt(OwnerName), "Owner Name") = False Then Exit Sub
    If CDate(txt(BookDate)) > CDate(txt(SrvDate)) Then
        MsgBox "Booking Date is greater than Service Date", vbCritical, "Date Checking"
        txt(SrvDate).SetFocus: Exit Sub
    End If
    If txt(VehRegNo).TEXT = "" And txt(Chassis).TEXT = "" Then
        MsgBox "Regitration No. or Chassis No. should have some data"
        txt(VehRegNo).SetFocus
        Exit Sub
    End If
    
    If txt(VehRegNo).TEXT <> "" Then
        GSQL = "select Count(*) from job_booking where regno='" & txt(VehRegNo).TEXT & "' and forservicedate=" & ConvertDate(txt(SrvDate).TEXT) & " and book_no<>" & txt(BookSrl).TEXT
        If GCn.Execute(GSQL).Fields(0).Value > 0 Then
            MsgBox "Vehicle is already booked for same date", vbInformation, "Validation"
            'txt(VehRegNo).SetFocus
            Exit Sub
        End If
        GSQL = "select " & xIsNull("Chassis", "") & " as Chassis from job_booking where regno='" & txt(VehRegNo).TEXT & "'"
        Set GRs = New ADODB.Recordset
        GRs.CursorLocation = adUseClient
        GRs.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
        If GRs.RecordCount > 0 Then
            If txt(Chassis).TEXT <> GRs!Chassis And GRs!Chassis <> "" Then
                GSQL = "Registration No. is already allocated with Chassis No. " & GRs!Chassis & vbCrLf & "Continue ?"
                If MsgBox(GSQL, vbYesNo, "Chassis/Reg.No Validation") = vbNo Then
                    Set GRs = Nothing
                    txt(VehRegNo).SetFocus: Exit Sub
                End If
            End If
        End If
        Set GRs = Nothing
    End If
    
    If txt(Chassis).TEXT <> "" Then
        GSQL = "select Count(*) from job_booking where chassis='" & txt(Chassis).TEXT & "' and forservicedate=" & ConvertDate(txt(SrvDate).TEXT) & " and book_no<>" & txt(BookSrl).TEXT
        If GCn.Execute(GSQL).Fields(0).Value > 0 Then
            MsgBox "Vehicle is already booked for same date", vbInformation, "Validation"
            txt(Chassis).SetFocus
            Exit Sub
        End If
        GSQL = "select regno from job_booking where Chassis='" & txt(Chassis).TEXT & "'"
        Set GRs = New ADODB.Recordset
        GRs.CursorLocation = adUseClient
        GRs.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
        If GRs.RecordCount > 0 Then
            If txt(VehRegNo).TEXT <> GRs!RegNo And GRs!RegNo <> "" Then
                GSQL = "Chassis No. is already allocated with Registration No. " & GRs!RegNo & vbCrLf & "Continue ?"
                If MsgBox(GSQL, vbYesNo, "Chassis/Reg.No Validation") = vbNo Then
                    Set GRs = Nothing
                    'txt(Chassis).SetFocus
                    Exit Sub
                End If
            End If
        End If
        Set GRs = Nothing
    End If
    GCn.BeginTrans
    If Val(txt(AdvAmt)) > 0 Then GCnFaW.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2 = "Add" Then
        'srl no. get
        txt(BookSrl).TEXT = GCn.Execute("select " & vIsNull("max(book_no)", "0") & "+1 from job_booking").Fields(0)
        DocID = PubDivCode & PubSiteCode & Right(Space(8) + txt(BookSrl).TEXT, 8)  'Trim(TxtYear) + Trim(TxtVType) + Trim(CStr(TxtVNo))
        'insert rec
        GSQL = "insert into Job_Booking(Div_Code , Site_Code, Book_No, Book_Date, Name, Add1, Add2, Add3, CityCode, PhoneOff, PhoneResi, Mobile , Model, Chassis, Engine, Regno,Advance,ForServiceDate,Remarks,Service_Type, U_Name, U_EntDt, U_AE) " & _
            " values('" & PubDivCode & "','" & PubSiteCode & "'," & txt(BookSrl).TEXT & "," & ConvertDate(txt(BookDate).TEXT) & ",'" & txt(OwnerName).TEXT & "','" & txt(Address1).TEXT & "','" & txt(Address2).TEXT & "','" & txt(Address3).TEXT & "', '" & txt(City).Tag & "','" & txt(PhoneOff).TEXT & "'," & _
            " '" & txt(PhoneResi).TEXT & "','" & txt(Mobile).TEXT & "','" & txt(Model).Tag & "','" & txt(Chassis).TEXT & "','" & txt(Engine).TEXT & "','" & txt(VehRegNo).TEXT & "'," & Val(txt(AdvAmt).TEXT) & "," & ConvertDate(txt(SrvDate).TEXT) & ",'" & txt(Remarks).TEXT & "','" & txt(SrvType).Tag & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
    Else    'Edit
'        DocId = Master!Div_Code & Master!Site_Code & Right(Space(8) & Master!Book_no, 8) 'Trim(TxtYear) + Trim(TxtVType) + Trim(CStr(TxtVNo))
        DocID = Master!SearchCode
        GSQL = "update Job_Booking set Book_Date=" & ConvertDate(txt(BookDate).TEXT) & ", Name='" & txt(OwnerName).TEXT & _
            "', Add1='" & txt(Address1).TEXT & "', Add2='" & txt(Address2).TEXT & "', Add3='" & txt(Address3).TEXT & "', CityCode = '" & txt(City).Tag & _
            "', PhoneOff='" & txt(PhoneOff).TEXT & "', PhoneResi='" & txt(PhoneResi).TEXT & "', Mobile='" & txt(Mobile).TEXT & _
            "', Model='" & txt(Model).Tag & "', Chassis='" & txt(Chassis).TEXT & "', Engine='" & txt(Engine).TEXT & _
            "', Regno='" & txt(VehRegNo).TEXT & "',Advance=" & Val(txt(AdvAmt).TEXT) & ",ForServiceDate=" & ConvertDate(txt(SrvDate).TEXT) & _
            ",Remarks='" & txt(Remarks).TEXT & "',Service_Type='" & txt(SrvType).Tag & "', U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' " & _
            " where Div_code='" & Master!Div_Code & "' and book_no=" & txt(BookSrl) & " and Site_Code='" & Master!Site_Code & "'"
'        GCn.Execute ("delete from job_booking where Book_no=" & Val(txt(BookSrl).Text))
    End If
    GCn.Execute GSQL
    
    If Val(txt(AdvAmt)) > 0 Then GCnFaW.CommitTrans
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select JB.Div_Code+JB.Site_Code+right(space(8)+" & cCStr("JB.Book_no") & ",8) as SearchCode, JB.Div_Code,JB.Site_Code,JB.Book_No from Job_Booking as JB  Where JB.Div_Code+JB.Site_Code+right(space(8)+" & cCStr("JB.Book_no") & ",8) = '" & DocID & "' order by JB.Book_Date desc, JB.Div_Code desc,JB.Site_Code desc,JB.Book_no desc")
    End If
    RSBook.Requery
    Master.FIND "SearchCode = '" & DocID & "'"
'    If TopCtrl1.TopText2.Caption = "Add" Then
'        Txt(BookDate).Tag = Txt(BookDate).Text
'        TopCtrl1_eAdd
'        Exit Sub
'    End If
'    Disp_Text SETS("INI", Me, Master)
'    Call MoveRec
    TopCtrl1_ePrn
    Exit Sub

errlbl:
    If mTrans = True Then
        GCn.RollbackTrans
        If Val(txt(AdvAmt)) > 0 Then GCnFaW.CommitTrans
    End If
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus txt(Index)
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case VehRegNo
            DGridColSwap DGBook, 0
            RSBook.Sort = "RegNo"
            If RSBook.RecordCount = 0 Or txt(Index) = "" Then Exit Sub
            If UCase(txt(Index)) <> UCase(RSBook!RegNo) Then
                RSBook.MoveFirst
                RSBook.FIND "RegNo ='" & txt(Index) & "'"
            End If
            
        Case Chassis
            DGridColSwap DGBook, 1
            RSBook.Sort = "CHASSIS"
            If RSBook.RecordCount = 0 Or txt(Index) = "" Then Exit Sub
            If UCase(txt(Index)) <> UCase(RSBook!Chassis) Then
                RSBook.MoveFirst
                RSBook.FIND "Chassis ='" & txt(Index) & "'"
            End If
            
        Case OwnerName
            DGridColSwap DGBook, 3
            RSBook.Sort = "name"
            
        Case SrvType
            DGridColSwap DGService, 1
            If RsServ.RecordCount = 0 Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).TEXT <> RsServ!Code Then
                RsServ.MoveFirst
                RsServ.FIND "name ='" & txt(Index).TEXT & "'"
            End If
        Case Model
            If RsModel.RecordCount = 0 Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).TEXT <> RsModel!Code Then
                RsModel.MoveFirst
                RsModel.FIND "name ='" & txt(Index).TEXT & "'"
            End If
            
            
        Case City
            DGCity.Move txt(Index).left, txt(Index).top + txt(Index).height + 20
            If RsCity.RecordCount = 0 Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).TEXT <> RsCity!Code Then
                RsCity.MoveFirst
                RsCity.FIND "name ='" & txt(Index).TEXT & "'"
            End If
            
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
Dim mInd As Integer
    Select Case Index
        Case VehRegNo
            DGridTxtKeyDown_Mast DGBook, txt, Index, RSBook, KeyCode, False, 1
        Case Chassis
            'If Not IsEmpty(txt(Chassis).Text) Then RsBook.FIND ("CHASSIS='" & txt(Chassis) & "'")
            DGridTxtKeyDown_Mast DGBook, txt, Index, RSBook, KeyCode, False, 2
        Case OwnerName
            DGridTxtKeyDown_Mast DGBook, txt, Index, RSBook, KeyCode, False, 3
        Case Model
            'If Not IsEmpty(txt(MODEL).Tag) Then RsModel.FIND ("CODE= '" & txt(MODEL).Tag & "'")
            DGridTxtKeyDown DGModel, txt, Index, RsModel, KeyCode, False, 0, frmModel, "frmModel"
        
        Case City
            DGridTxtKeyDown DGCity, txt, Index, RsCity, KeyCode, False, 1, frmCity, "frmCity"
        
        Case SrvType
            DGridTxtKeyDown DGService, txt, Index, RsServ, KeyCode, False, 1, frmService, "frmService"
    End Select
    If DGBook.Visible = False And DGModel.Visible = False And DGService.Visible = False And DGCity.Visible = False Then
        ' KEY DOWN
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> Remarks Then
            Ctrl_DownKeyDown KeyCode, Shift
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = Remarks Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        ' Key UP
        If KeyCode = vbKeyUp And Index > 0 Then   ' <> VehRegNo Then
            For mInd = Index To 1 Step -1
                If mInd > 1 Then
                    If txt(mInd - 1).Enabled = True Then mInd = mInd - 1: Exit For
                Else
                    If txt(mInd).Enabled = True Then Exit For
                End If
            Next
            If mInd > 0 And mInd <> Index Then
                'txt(mInd).SetFocus ' SendKeys "+{Tab}"  'testing by lps
                Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
    Call CheckQuote(KeyAscii)
    Select Case Index
        Case BookSrl
            Call NumPress(txt(Index), KeyAscii, 8, 0)
        Case AdvAmt
            Call NumPress(txt(Index), KeyAscii, 8, 2)
        Case SrvType
            DGridTxtKeyPress txt, Index, RsServ, KeyAscii, "name"
        Case Model
            DGridTxtKeyPress txt, Index, RsModel, KeyAscii, "code"
        Case City
            DGridTxtKeyPress txt, Index, RsCity, KeyAscii, "Name"
    End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case VehRegNo
            DGridTxtKeyUp_Mast txt, Index, RSBook, KeyCode, "Regno"
        Case Chassis
            DGridTxtKeyUp_Mast txt, Index, RSBook, KeyCode, "Chassis"
        Case OwnerName
            DGridTxtKeyUp_Mast txt, Index, RSBook, KeyCode, "Name"
    End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim RsTemp As ADODB.Recordset
Select Case Index
    Case VehRegNo
        If RSBook.RecordCount > 0 Then
            If txt(VehRegNo) <> "" And UCase(Trim(txt(VehRegNo))) = UCase(Trim(RSBook!RegNo)) Then
                VehDetails
                txt(Model).Enabled = False
                txt(Chassis).Enabled = False
                txt(Engine).Enabled = False
            Else
                txt(Model).Enabled = True
                txt(Chassis).Enabled = True
                txt(Engine).Enabled = True
            End If
        End If
    Case Chassis
        If RSBook.RecordCount > 0 And RSBook.EOF = False And RSBook.BOF = False Then
            If txt(Chassis) <> "" And UCase(Trim(txt(Chassis))) = UCase(Trim(RSBook!Chassis)) Then
                VehDetails
                txt(Model).Enabled = False
                txt(VehRegNo).Enabled = False
                txt(Engine).Enabled = False
            Else
                Set RsTemp = GCn.Execute("Select Model From Model Where Chas_Type='" & left(Trim(txt(Chassis)), 6) & "'")
                If RsTemp.RecordCount > 0 Then
                    txt(Model) = XNull(RsTemp(0))
                End If
                
                txt(Model).Enabled = True
                txt(VehRegNo).Enabled = True
                txt(Engine).Enabled = True
            End If
        End If
    Case OwnerName
        If RSBook.RecordCount > 0 And RSBook.EOF = False And RSBook.BOF = False Then
            If txt(OwnerName).TEXT <> "" And UCase(Trim(txt(OwnerName).TEXT)) = UCase(Trim(RSBook!Name)) Then
                txt(Address1).TEXT = XNull(RSBook!Add1)
                txt(Address2).TEXT = XNull(RSBook!Add2)
                txt(Address3).TEXT = XNull(RSBook!Add3)
                txt(PhoneResi).TEXT = XNull(RSBook!PhoneResi)
                txt(PhoneOff).TEXT = XNull(RSBook!PhoneOff)
                txt(Mobile).TEXT = XNull(RSBook!Mobile)
            End If
        End If
        
    Case SrvType
        If txt(SrvType).TEXT <> "" And RsServ.EOF = False And RsServ.BOF = False Then
            txt(SrvType).TEXT = RsServ!Name
            txt(SrvType).Tag = RsServ!Code
        End If
        
    Case BookDate
        txt(BookDate).TEXT = RetDate(txt(BookDate))
        
    Case SrvDate
        txt(SrvDate).TEXT = RetDate(txt(SrvDate))
        Call veh_count
        
    Case Model
        If txt(Model).TEXT <> "" Then
            txt(Model).TEXT = RsModel!Code
        End If
        
    Case City
        If RsCity.RecordCount > 0 Then
            If txt(City) <> "" And UCase(Trim(txt(City))) = UCase(Trim(RsCity!Name)) Then
                txt(Index) = RsCity!Name
                txt(Index).Tag = RsCity!Code
            Else
                txt(Index) = ""
                txt(Index).Tag = ""
            End If
        End If
        
End Select
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    For I = 0 To 22
        txt(I).TEXT = ""
    Next I
End Sub

Private Sub MoveRec()
Dim Master1 As ADODB.Recordset
'Dim rs As Recordset
'Dim mVor As String
On Error GoTo error1

If Master.RecordCount > 0 Then
    Set Master1 = New Recordset
    Master1.CursorLocation = adUseClient
    Master1.Open "select JB.*, Service_type.Serv_type,Service_type.Serv_desc from Job_Booking as JB left join Service_type on JB.Service_type=Service_type.serv_type where JB.Div_Code+JB.Site_Code+right(space(8)+" & cCStr("JB.Book_no") & ",8)='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
    
    LblDiv.CAPTION = "Division : " & Master1!Div_Code
    LblSite.CAPTION = "Site Code : " & Master1!Site_Code
    txt(BookSrl).TEXT = Master1!Book_no
    txt(BookDate).TEXT = RetDate(Master1!Book_Date)
    txt(VehRegNo).TEXT = Master1!RegNo
    txt(Model).Tag = Master1!Model
    txt(Model).TEXT = Master1!Model
    txt(Chassis).TEXT = Master1!Chassis
    txt(Engine).TEXT = Master1!Engine
    txt(OwnerName).TEXT = XNull(Master1!Name)
    txt(Address1).TEXT = XNull(Master1!Add1)
    txt(Address2).TEXT = XNull(Master1!Add2)
    txt(Address3).TEXT = XNull(Master1!Add3)
    txt(PhoneOff).TEXT = XNull(Master1!PhoneOff)
    txt(PhoneResi).TEXT = XNull(Master1!PhoneResi)
    txt(Mobile).TEXT = XNull(Master1!Mobile)
'    txt(LastJobNo).Text = Master1!Book_Date
'    txt(LastJobDt).Text = Master1!Book_Date
'    txt(LastSrv).Text = Master1!Book_Date
'    txt(LastKMS).Text = Master1!Book_Date
    txt(AdvAmt).TEXT = XNull(Master1!Advance)
    txt(SrvType).Tag = XNull(Master1!Service_Type)
    txt(SrvType).TEXT = XNull(Master1!Serv_Desc)
    txt(SrvDate).TEXT = XNull(Master1!Forservicedate)
    txt(Remarks).TEXT = XNull(Master1!Remarks)
    txt(EntryDate).TEXT = XNull(Master1!U_EntDt)
    txt(JobDocID).Tag = IIf(IsNull(Master1!job_docid), "", Master1!job_docid)
    txt(JobDocID) = DeCodeDocID(txt(JobDocID).Tag, Document_No)
    Call veh_count
Else
    Call BlankText
End If
Set Master1 = Nothing
Grid_Hide
Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
    DGBook.width = Me.width - 90: DGBook.left = Me.left: DGBook.top = 4600: DGBook.height = 2500
    DGBook.Columns(1).width = 1890.142
    DGBook.Columns(2).width = 1920.189
    DGBook.Columns(3).width = 3344.882
    DGBook.Columns(4).width = 1214.929
    DGBook.Columns(5).width = 1214.929
    
    DGModel.width = DGBook.width: DGModel.left = DGBook.left: DGModel.top = DGBook.top: DGModel.height = DGBook.height
    DGService.width = 4440: DGService.left = 7365: DGService.top = 2430: DGService.height = 2730
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To txt.Count - 1
        txt(I).Enabled = Enb
    Next
    For I = 0 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg
        txt(I).ForeColor = CtrlFColOrg
    Next
    txt(LastJobDt).Enabled = False
    txt(LastJobNo).Enabled = False
    txt(LastSrv).Enabled = False
    txt(LastKMS).Enabled = False
    txt(BookSrl).Enabled = False
    txt(EntryDate).Enabled = False
End Sub

Private Sub Grid_Hide()
    If DGBook.Visible = True Then DGBook.Visible = False
    If DGModel.Visible = True Then DGModel.Visible = False
    If DGService.Visible = True Then DGService.Visible = False
    If DGCity.Visible = True Then DGCity.Visible = False
End Sub

Private Sub veh_count()
    LblTotVeh.CAPTION = GCn.Execute("select count(*) from job_booking where Div_Code='" & PubDivCode & "' and forservicedate=" & ConvertDate(txt(SrvDate).TEXT)).Fields(0)
    'LblAmt.CAPTION = GCn.Execute("select IIF(ISNULL(Sum(advance)),0,Sum(advance)) from job_booking where Div_Code='" & PubDivCode & "' and forservicedate=" & ConvertDate(Txt(SrvDate).TEXT)).Fields(0)
    LblAmt.CAPTION = GCn.Execute("select " & vIsNull("Sum(advance)", "0") & " from job_booking where Div_Code='" & PubDivCode & "' and forservicedate=" & ConvertDate(txt(SrvDate).TEXT)).Fields(0)
End Sub

Private Sub VehDetails()
    txt(VehRegNo).TEXT = RSBook!RegNo   'Not filled when VehRegNo call
    txt(Model).TEXT = RSBook!Model
    txt(Model).Tag = RSBook!Model
    txt(Chassis).TEXT = RSBook!Chassis  'Not filled when Chassis call
    txt(Engine).TEXT = RSBook!Engine
    txt(OwnerName).TEXT = RSBook!Name
    txt(Address1).TEXT = XNull(RSBook!Add1)
    txt(Address2).TEXT = XNull(RSBook!Add2)
    txt(Address3).TEXT = XNull(RSBook!Add3)
    txt(City) = XNull(RSBook!CityName)
    txt(PhoneResi).TEXT = XNull(RSBook!PhoneResi)
    txt(PhoneOff).TEXT = XNull(RSBook!PhoneOff)
    txt(Mobile).TEXT = XNull(RSBook!Mobile)
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
Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "JobBook", "JobBook")
        Call WindowsPrint(Index)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "JobBook", "JobBook")
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
Private Sub WindowsPrint(Index As Integer)
Dim mQry As String, RepTitle$
Dim Condstr$, mDocStr$
Dim RST1 As ADODB.Recordset
Dim Speciality$
Dim Rst As ADODB.Recordset
Dim I As Integer

On Error GoTo ERRORHANDLER
mQry = "SELECT JB.RegNo, JB.Model, JB.Chassis, JB.Engine, JB.Name, JB.Add1, JB.Add2, JB.Add3, JB.PhoneOff, JB.PhoneResi, JB.Mobile,JB.Service_Type, Service_Type.Serv_Desc, JB.Book_No, JB.Book_Date, JB.Advance, JB.ForServiceDate,JB.Remarks " & _
        "FROM Job_Booking JB LEFT JOIN Service_Type ON JB.Service_Type = Service_Type.Serv_Type " & _
        "where JB.Div_Code+JB.Site_Code+right(space(8)+" & cCStr("JB.Book_no") & ",8)='" & Master!SearchCode & "'"
Set Rst = New Recordset
Rst.CursorLocation = adUseClient
Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic

If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
Speciality = GCn.Execute("Select W_SecSpeciality from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
 
CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")

Set RST1 = New Recordset
RST1.CursorLocation = adUseClient
RST1.Open "select Div_SName,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
    mDocStr = "** Job Booking Slip" & IIf(RST1!Div_SName = "", "", " (" & RST1!Div_SName & ")") & "**"

For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("SubTitle")
            rpt.FormulaFields(I).TEXT = "'" & Speciality & "'"
        Case UCase("Comp_Phone")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecPhone & "'"
        Case UCase("Comp_Fax")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecFax & "'"
        Case UCase("RepTitle")
            rpt.FormulaFields(I).TEXT = "'" & mDocStr & "'"
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
        Call Report_View(rpt, "Job Booking Slip", , True)
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

Private Sub SpeedPrint()
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
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstBook As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mRate As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, RepTitle$, Speciality$, mTaxdesc$, mGoods_Amt As Double
    Dim Footer$, FooterCnt As Byte, mHeader As Byte, mFooter As Byte, mGatePass As Byte, mDetTax As Byte
    Dim SubTot As Double, FormulaStr1$, FormulaStr2$, PhoneStr$
    Dim fob As New FileSystemObject, SecondStr As Boolean
    Dim LdRate As Double, LAmtVal As Double
    Dim LDAmt As Double, LAmtItem As Double, mGrossAmt As Double, mQry$

    mQry = "SELECT JB.RegNo, JB.Model, JB.Chassis, JB.Engine, JB.Name, JB.Add1, JB.Add2, JB.Add3, JB.PhoneOff, JB.PhoneResi, JB.Mobile,JB.Service_Type, Service_Type.Serv_Desc, JB.Book_No, JB.Book_Date, JB.Advance, JB.ForServiceDate,JB.Remarks " & _
        "FROM Job_Booking JB LEFT JOIN Service_Type ON JB.Service_Type = Service_Type.Serv_Type " & _
        "where JB.Div_Code+JB.Site_Code+right(space(8)+" & cCStr("JB.Book_no") & ",8)='" & Master!SearchCode & "'"
    
    Set RstBook = GCn.Execute(mQry)
    If RstBook.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    PageLength = PubPageLengthHalf
    PageWidth = 80   '137 for chr15
    'chr 17 to chr 10 - > X * 0.56
    'chr 10 to chr 17 - > X * 1.7
        
    mHeader = 0   'Ideal 17
    mFooter = 2
          
    'Header
    RepTitle = GCn.Execute("Select Div_SName from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
    mDocStr = "** Job Booking Slip" & IIf(RepTitle = "", "", " (" & RepTitle & ")") & " **"
    
    Set RstCompDet = GCn.Execute("select W_SecSpeciality,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'")
    PhoneStr = IIf(XNull(RstCompDet!W_SecPhone) = "", "", "Phone : " & XNull(RstCompDet!W_SecPhone))
    PhoneStr = PhoneStr & IIf(XNull(RstCompDet!W_SecFax) = "", "", "Fax : " & XNull(RstCompDet!W_SecFax))

    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!W_SecSpeciality) <> "" Then
        Print #1, PRN_TIT(RstCompDet!W_SecSpeciality, "C", PageWidth, True)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    If PubComp_Add2 <> "" Then
        Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    If PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_City, "B", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, Space((PageWidth - Len(PhoneStr)) / 2) + PhoneStr
    mHeader = mHeader + 1
    
    Print #1, PRN_TIT(mDocStr, "B", PageWidth, True)
    mHeader = mHeader + 1
    Print #1, PSTR("Reg. No.", 12) & ": " & PSTR(XNull(RstBook!RegNo), 20) & Space(15) & mEmph & PSTR("Booking No.", 12) & ": " & STR(RstBook!Book_no) & mEmph1
    mHeader = mHeader + 1
    Print #1, PSTR("Chassis No.", 12) & ": " & PSTR(XNull(RstBook!Chassis), 20) & Space(15) & PSTR("Booking Date", 12) & ": " & Format(RstBook!Book_Date, "dd/MMM/YYYY")
    mHeader = mHeader + 1
    Print #1, PSTR("Engine No.", 12) & ": " & PSTR(XNull(RstBook!Engine), 20) & Space(15) & mEmph & PSTR("Model", 12) & ": " & XNull(RstBook!Model) & mEmph1
    mHeader = mHeader + 1
    Print #1, "Owner: " & mEmph & PSTR(RstBook!Name, 40) & mEmph1 & Space(2) & PSTR("Service Type", 12) & ": " & RstBook!Service_Type & " " & RstBook!Serv_Desc
    mHeader = mHeader + 1
    Print #1, Space(7) & PSTR(XNull(RstBook!Add1), 40) & Space(2) & mEmph & PSTR("Service Date", 12) & ": " & XNull(RstBook!Forservicedate) & mEmph1
    mHeader = mHeader + 1
    Print #1, Space(7) & PSTR(XNull(RstBook!Add2), 40) & Space(2) '& "Advance Rs. : " & Format(RstBook!Advance, "0.00")
    mHeader = mHeader + 1
    Print #1, Space(7) & PSTR(XNull(RstBook!Add3), 40)
    mHeader = mHeader + 1
    Print #1, "Phone No. : " & IIf(RstBook!PhoneOff = "", "", "(O)") & RstBook!PhoneOff & IIf(RstBook!PhoneResi = "", "", "(R)") & RstBook!PhoneResi & IIf(RstBook!PhoneResi = "", "", "(Mobile)") & RstBook!Mobile
    mHeader = mHeader + 1
    Print #1, "Remarks   : " & RstBook!Remarks
    mHeader = mHeader + 1
    Do Until mHeader >= PageLength - 2
        Print #1, ""
        mHeader = mHeader + 1
    Loop
    ' FOOTER
    Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("* a dataman software *")) / 2) & "* a dataman software *" & mChr18
'mODI SHEKHAR
'     Change PageLength = 34
     Print #1, Chr(27) + Chr(67) + Chr(PageLength) ' instead of Print #1,meject
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








