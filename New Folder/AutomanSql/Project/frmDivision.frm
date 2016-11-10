VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38DE742C-94E6-11D6-A3DA-080030001F87}#8.0#0"; "KEYNO.OCX"
Begin VB.Form frmDivision 
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Division"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9135
   KeyPreview      =   -1  'True
   LinkTopic       =   "form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame FrameSite 
      Height          =   3150
      Left            =   4395
      TabIndex        =   337
      Top             =   210
      Visible         =   0   'False
      Width           =   7410
      Begin MSDataGridLib.DataGrid DgSite 
         Height          =   2415
         Left            =   15
         TabIndex        =   338
         Top             =   705
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   -1  'True
         ColumnHeaders   =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
            DataField       =   "Name"
            Caption         =   "Site"
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
            BeginProperty Column00 
               ColumnWidth     =   6600.189
            EndProperty
         EndProperty
      End
      Begin VB.Label LblSiteList 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Site Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   15
         TabIndex        =   237
         Top             =   105
         Width           =   7365
      End
      Begin VB.Label LblSiteHead 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Site Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   15
         TabIndex        =   339
         Top             =   405
         Width           =   7365
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   2370
      TabIndex        =   319
      Top             =   5145
      Visible         =   0   'False
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame FrmFirm 
      BackColor       =   &H00CFE0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5190
      Left            =   2340
      TabIndex        =   238
      Top             =   4830
      Visible         =   0   'False
      Width           =   9105
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   102
         Left            =   1635
         MaxLength       =   30
         TabIndex        =   260
         Top             =   3960
         Width           =   1050
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   101
         Left            =   1635
         TabIndex        =   259
         Top             =   3690
         Width           =   2790
      End
      Begin VB.CommandButton Cmd 
         DisabledPicture =   "frmDivision.frx":0000
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4005
         Picture         =   "frmDivision.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   273
         ToolTipText     =   "Get Data Path"
         Top             =   7185
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Index           =   93
         Left            =   1635
         MaxLength       =   100
         TabIndex        =   272
         Top             =   4875
         Visible         =   0   'False
         Width           =   7215
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   92
         Left            =   1635
         MaxLength       =   80
         TabIndex        =   271
         Top             =   4230
         Width           =   7215
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   91
         Left            =   6165
         MaxLength       =   25
         TabIndex        =   270
         Top             =   3420
         Width           =   2670
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   90
         Left            =   6165
         MaxLength       =   25
         TabIndex        =   269
         Top             =   3150
         Width           =   2670
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   89
         Left            =   6165
         MaxLength       =   25
         TabIndex        =   268
         Top             =   2880
         Width           =   2670
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   88
         Left            =   6165
         MaxLength       =   40
         TabIndex        =   267
         Top             =   2610
         Width           =   2670
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   87
         Left            =   6165
         MaxLength       =   15
         TabIndex        =   266
         Top             =   2340
         Width           =   2670
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   86
         Left            =   6165
         MaxLength       =   15
         TabIndex        =   265
         Top             =   2070
         Width           =   2670
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   85
         Left            =   6165
         MaxLength       =   20
         TabIndex        =   264
         Top             =   1800
         Width           =   2670
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   84
         Left            =   6165
         MaxLength       =   10
         TabIndex        =   263
         Top             =   1530
         Width           =   1530
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   83
         Left            =   6165
         TabIndex        =   262
         Top             =   1260
         Width           =   1050
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   82
         Left            =   6165
         MaxLength       =   30
         TabIndex        =   261
         Top             =   990
         Width           =   2715
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   81
         Left            =   1635
         TabIndex        =   258
         Top             =   3420
         Width           =   1050
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   80
         Left            =   1635
         MaxLength       =   30
         TabIndex        =   257
         Top             =   3150
         Width           =   2790
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   79
         Left            =   1635
         MaxLength       =   6
         TabIndex        =   256
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   78
         Left            =   1635
         MaxLength       =   25
         TabIndex        =   255
         Top             =   2610
         Width           =   2790
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   77
         Left            =   1635
         MaxLength       =   40
         TabIndex        =   254
         Top             =   2340
         Width           =   2790
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   76
         Left            =   1635
         MaxLength       =   40
         TabIndex        =   253
         Top             =   2070
         Width           =   2790
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   75
         Left            =   1635
         MaxLength       =   40
         TabIndex        =   252
         Top             =   1800
         Width           =   2790
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   74
         Left            =   1635
         MaxLength       =   40
         TabIndex        =   251
         Top             =   1530
         Width           =   2790
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   73
         Left            =   1635
         MaxLength       =   15
         TabIndex        =   250
         Top             =   1260
         Width           =   1725
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Index           =   72
         Left            =   1635
         MaxLength       =   1
         TabIndex        =   249
         Top             =   990
         Width           =   400
      End
      Begin VB.CommandButton CmdFirm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   3270
         MousePointer    =   99  'Custom
         Picture         =   "frmDivision.frx":0E04
         Style           =   1  'Graphical
         TabIndex        =   248
         ToolTipText     =   "Cancel Record"
         Top             =   15
         Width           =   405
      End
      Begin VB.CommandButton CmdFirm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   2865
         MousePointer    =   99  'Custom
         Picture         =   "frmDivision.frx":118C
         Style           =   1  'Graphical
         TabIndex        =   247
         ToolTipText     =   "Save Record"
         Top             =   15
         Width           =   405
      End
      Begin VB.CommandButton CmdFirm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   3675
         MousePointer    =   99  'Custom
         Picture         =   "frmDivision.frx":16BE
         Style           =   1  'Graphical
         TabIndex        =   246
         ToolTipText     =   "Exit Form"
         Top             =   15
         Width           =   405
      End
      Begin VB.CommandButton CmdFirm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   9
         Left            =   2460
         MousePointer    =   99  'Custom
         Picture         =   "frmDivision.frx":1755
         Style           =   1  'Graphical
         TabIndex        =   245
         ToolTipText     =   "Move to Last Record"
         Top             =   15
         Width           =   405
      End
      Begin VB.CommandButton CmdFirm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   8
         Left            =   2055
         MousePointer    =   99  'Custom
         Picture         =   "frmDivision.frx":1A5F
         Style           =   1  'Graphical
         TabIndex        =   244
         ToolTipText     =   "Move to Next Record"
         Top             =   15
         Width           =   405
      End
      Begin VB.CommandButton CmdFirm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   1650
         MousePointer    =   99  'Custom
         Picture         =   "frmDivision.frx":1D69
         Style           =   1  'Graphical
         TabIndex        =   243
         ToolTipText     =   "Move to Previous Record"
         Top             =   15
         Width           =   405
      End
      Begin VB.CommandButton CmdFirm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   1245
         MaskColor       =   &H00FFFFFF&
         MousePointer    =   99  'Custom
         Picture         =   "frmDivision.frx":2073
         Style           =   1  'Graphical
         TabIndex        =   242
         ToolTipText     =   "Move to First Record"
         Top             =   15
         Width           =   405
      End
      Begin VB.CommandButton CmdFirm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   840
         MousePointer    =   99  'Custom
         Picture         =   "frmDivision.frx":237D
         Style           =   1  'Graphical
         TabIndex        =   241
         ToolTipText     =   "Delete Current Record"
         Top             =   15
         Width           =   405
      End
      Begin VB.CommandButton CmdFirm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   435
         MousePointer    =   99  'Custom
         Picture         =   "frmDivision.frx":24C7
         Style           =   1  'Graphical
         TabIndex        =   240
         ToolTipText     =   "Edit Current Record"
         Top             =   15
         Width           =   405
      End
      Begin VB.CommandButton CmdFirm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   30
         MousePointer    =   99  'Custom
         Picture         =   "frmDivision.frx":255A
         Style           =   1  'Graphical
         TabIndex        =   239
         ToolTipText     =   "Add new Record"
         Top             =   15
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LST Date"
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
         Index           =   97
         Left            =   180
         TabIndex        =   336
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LST No."
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
         Index           =   96
         Left            =   180
         TabIndex        =   335
         Top             =   3690
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   89
         Left            =   1485
         TabIndex        =   334
         Top             =   3690
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
         Index           =   88
         Left            =   1485
         TabIndex        =   333
         Top             =   3960
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
         Index           =   72
         Left            =   1485
         TabIndex        =   314
         Top             =   4500
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
         Index           =   73
         Left            =   1485
         TabIndex        =   313
         Top             =   4230
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
         Index           =   74
         Left            =   6030
         TabIndex        =   312
         Top             =   3420
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
         Index           =   75
         Left            =   6030
         TabIndex        =   311
         Top             =   3150
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
         Index           =   76
         Left            =   6030
         TabIndex        =   310
         Top             =   2880
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IT Ward No."
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
         Index           =   100
         Left            =   4725
         TabIndex        =   309
         Top             =   2880
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IT Ac No."
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
         Index           =   101
         Left            =   4725
         TabIndex        =   308
         Top             =   3150
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAN No."
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
         Index           =   102
         Left            =   4725
         TabIndex        =   307
         Top             =   3420
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speciality"
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
         Index           =   75
         Left            =   180
         TabIndex        =   306
         Top             =   4230
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FaDataPath"
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
         Index           =   93
         Left            =   180
         TabIndex        =   305
         Top             =   4500
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
         Index           =   77
         Left            =   6030
         TabIndex        =   304
         Top             =   2610
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
         Index           =   78
         Left            =   6030
         TabIndex        =   303
         Top             =   2340
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
         Index           =   79
         Left            =   6030
         TabIndex        =   302
         Top             =   2070
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
         Index           =   80
         Left            =   6030
         TabIndex        =   301
         Top             =   1800
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
         Index           =   81
         Left            =   6030
         TabIndex        =   300
         Top             =   1530
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
         Index           =   82
         Left            =   1485
         TabIndex        =   299
         Top             =   3420
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CST Date "
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
         Index           =   77
         Left            =   4725
         TabIndex        =   298
         Top             =   1275
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No."
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
         Index           =   78
         Left            =   4725
         TabIndex        =   297
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
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
         Index           =   79
         Left            =   4725
         TabIndex        =   296
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax No."
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
         Index           =   80
         Left            =   4725
         TabIndex        =   295
         Top             =   2070
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tele Gram No."
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
         Index           =   81
         Left            =   4710
         TabIndex        =   294
         Top             =   2340
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mail Id."
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
         Index           =   82
         Left            =   4710
         TabIndex        =   293
         Top             =   2610
         Width           =   570
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
         Index           =   83
         Left            =   6030
         TabIndex        =   292
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   84
         Left            =   6030
         TabIndex        =   291
         Top             =   1275
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
         Index           =   85
         Left            =   1485
         TabIndex        =   290
         Top             =   3150
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
         Index           =   86
         Left            =   1485
         TabIndex        =   289
         Top             =   2880
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
         Index           =   87
         Left            =   1485
         TabIndex        =   288
         Top             =   2610
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City Name"
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
         Index           =   83
         Left            =   180
         TabIndex        =   287
         Top             =   2610
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code"
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
         Index           =   84
         Left            =   180
         TabIndex        =   286
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tin No."
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
         Index           =   85
         Left            =   180
         TabIndex        =   285
         Top             =   3150
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tin Date"
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
         Index           =   86
         Left            =   180
         TabIndex        =   284
         Top             =   3420
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CST No. "
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
         Index           =   87
         Left            =   4725
         TabIndex        =   283
         Top             =   1005
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   90
         Left            =   1485
         TabIndex        =   282
         Top             =   1800
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
         Index           =   91
         Left            =   1485
         TabIndex        =   281
         Top             =   1530
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
         Index           =   92
         Left            =   1485
         TabIndex        =   280
         Top             =   1260
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
         Index           =   93
         Left            =   1485
         TabIndex        =   279
         Top             =   998
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Firm Code  "
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
         Index           =   72
         Left            =   180
         TabIndex        =   278
         Top             =   998
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Short Name  "
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
         Index           =   73
         Left            =   180
         TabIndex        =   277
         Top             =   1260
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Firm Name  "
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
         Index           =   74
         Left            =   180
         TabIndex        =   276
         Top             =   1530
         Width           =   1020
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   91
         Left            =   180
         TabIndex        =   275
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         Height          =   240
         Left            =   1635
         TabIndex        =   274
         Top             =   4515
         Width           =   7080
      End
   End
   Begin VB.Frame FrmList 
      BackColor       =   &H00CFE0E0&
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   -570
      TabIndex        =   235
      Top             =   5445
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   135
         TabIndex        =   236
         TabStop         =   0   'False
         Top             =   90
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
         BackColor       =   13623520
         BorderStyle     =   1
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
   Begin TabDlg.SSTab STab 
      Height          =   5010
      Left            =   8595
      TabIndex        =   10
      Top             =   4545
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   8837
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   12583104
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Division Information"
      TabPicture(0)   =   "frmDivision.frx":2798
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vehicle Section"
      TabPicture(1)   =   "frmDivision.frx":27B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(1)=   "Label3(10)"
      Tab(1).Control(2)=   "Label3(11)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Spare Section"
      TabPicture(2)   =   "frmDivision.frx":27D0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).Control(1)=   "Label3(33)"
      Tab(2).Control(2)=   "Label3(32)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Workshop Section"
      TabPicture(3)   =   "frmDivision.frx":27EC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(3)"
      Tab(3).Control(1)=   "Label3(55)"
      Tab(3).Control(2)=   "Label3(54)"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame1 
         BackColor       =   &H00CFE0E0&
         BorderStyle     =   0  'None
         Height          =   4230
         Index           =   0
         Left            =   120
         TabIndex        =   220
         Top             =   390
         Width           =   8760
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   3045
            MaxLength       =   5
            TabIndex        =   14
            Top             =   1995
            Width           =   975
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   3045
            MaxLength       =   40
            TabIndex        =   15
            Top             =   2280
            Width           =   5055
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   94
            Left            =   3045
            MaxLength       =   25
            TabIndex        =   13
            Top             =   1710
            Width           =   2280
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   3045
            MaxLength       =   1
            TabIndex        =   12
            Top             =   1440
            Width           =   400
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Short Name  "
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
            Index           =   76
            Left            =   1485
            TabIndex        =   316
            Top             =   1995
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
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   4
            Left            =   2880
            TabIndex        =   315
            Top             =   1995
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
            Index           =   1
            Left            =   2880
            TabIndex        =   226
            Top             =   2280
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
            Index           =   0
            Left            =   2880
            TabIndex        =   225
            Top             =   1710
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
            Index           =   20
            Left            =   2880
            TabIndex        =   224
            Top             =   1440
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Division Name  "
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
            Index           =   2
            Left            =   1485
            TabIndex        =   223
            Top             =   2280
            Width           =   1305
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Srl No."
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
            Left            =   1485
            TabIndex        =   222
            Top             =   1710
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Division Code"
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
            Left            =   1485
            TabIndex        =   221
            Top             =   1440
            Width           =   1155
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00CFE0E0&
         BorderStyle     =   0  'None
         Height          =   4245
         Index           =   1
         Left            =   -74955
         TabIndex        =   179
         Top             =   345
         Width           =   8760
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   96
            Left            =   1515
            MaxLength       =   30
            TabIndex        =   27
            Top             =   3210
            Width           =   975
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   95
            Left            =   1515
            MaxLength       =   30
            TabIndex        =   26
            Text            =   "31"
            Top             =   2940
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   27
            Left            =   510
            MaxLength       =   100
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   4155
            Width           =   7065
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   26
            Left            =   1515
            MaxLength       =   80
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   3480
            Width           =   7080
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   25
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   37
            Top             =   2670
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   24
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   36
            Top             =   2400
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   23
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   35
            Top             =   2130
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   22
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   34
            Top             =   1860
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   21
            Left            =   5895
            MaxLength       =   15
            TabIndex        =   33
            Top             =   1590
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   20
            Left            =   5895
            MaxLength       =   15
            TabIndex        =   32
            Top             =   1320
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   19
            Left            =   5895
            MaxLength       =   20
            TabIndex        =   31
            Top             =   1050
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   18
            Left            =   5895
            MaxLength       =   10
            TabIndex        =   30
            Top             =   780
            Width           =   1530
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   17
            Left            =   5895
            TabIndex        =   29
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   16
            Left            =   5895
            MaxLength       =   30
            TabIndex        =   28
            Top             =   240
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   15
            Left            =   1515
            TabIndex        =   25
            Top             =   2670
            Width           =   975
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   14
            Left            =   1515
            MaxLength       =   30
            TabIndex        =   24
            Top             =   2400
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   13
            Left            =   1515
            MaxLength       =   6
            TabIndex        =   23
            Top             =   2130
            Width           =   735
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   12
            Left            =   1515
            MaxLength       =   25
            TabIndex        =   22
            Top             =   1860
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   11
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   21
            Top             =   1590
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   10
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   20
            Top             =   1320
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   9
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   19
            Top             =   1050
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   8
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   18
            Top             =   780
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   7
            Left            =   1515
            MaxLength       =   15
            TabIndex        =   17
            Top             =   510
            Width           =   1650
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00800080&
            Height          =   240
            Index           =   6
            Left            =   1515
            MaxLength       =   1
            TabIndex        =   16
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LST Date"
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
            Height          =   240
            Index           =   89
            Left            =   75
            TabIndex        =   324
            Top             =   3210
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LST No."
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
            Height          =   240
            Index           =   88
            Left            =   75
            TabIndex        =   323
            Top             =   2940
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   32
            Left            =   1365
            TabIndex        =   322
            Top             =   2940
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
            Height          =   240
            Index           =   10
            Left            =   1365
            TabIndex        =   321
            Top             =   3210
            Width           =   45
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Label6"
            Height          =   255
            Left            =   1515
            TabIndex        =   39
            Top             =   3750
            Width           =   7065
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
            Index           =   27
            Left            =   1350
            TabIndex        =   219
            Top             =   3750
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
            Height          =   255
            Index           =   26
            Left            =   1350
            TabIndex        =   218
            Top             =   3480
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
            Height          =   240
            Index           =   25
            Left            =   5760
            TabIndex        =   217
            Top             =   2670
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
            Height          =   240
            Index           =   24
            Left            =   5760
            TabIndex        =   216
            Top             =   2400
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
            Height          =   240
            Index           =   23
            Left            =   5760
            TabIndex        =   215
            Top             =   2130
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IT Ward No."
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
            Height          =   240
            Index           =   23
            Left            =   4455
            TabIndex        =   214
            Top             =   2130
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IT Ac No."
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
            Height          =   240
            Index           =   24
            Left            =   4455
            TabIndex        =   213
            Top             =   2400
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PAN No."
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
            Height          =   240
            Index           =   25
            Left            =   4455
            TabIndex        =   212
            Top             =   2670
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Speciality"
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
            Index           =   26
            Left            =   60
            TabIndex        =   211
            Top             =   3480
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FaDataPath"
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
            Index           =   27
            Left            =   60
            TabIndex        =   210
            Top             =   3750
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
            Height          =   240
            Index           =   22
            Left            =   5760
            TabIndex        =   209
            Top             =   1860
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
            Height          =   240
            Index           =   21
            Left            =   5760
            TabIndex        =   208
            Top             =   1590
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
            Height          =   240
            Index           =   19
            Left            =   5760
            TabIndex        =   207
            Top             =   1320
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
            Height          =   240
            Index           =   18
            Left            =   5760
            TabIndex        =   206
            Top             =   1050
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
            Height          =   240
            Index           =   17
            Left            =   5760
            TabIndex        =   205
            Top             =   780
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
            Height          =   240
            Index           =   16
            Left            =   5760
            TabIndex        =   204
            Top             =   510
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CST Date "
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
            Height          =   240
            Index           =   17
            Left            =   4455
            TabIndex        =   203
            Top             =   510
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile No."
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
            Height          =   240
            Index           =   18
            Left            =   4455
            TabIndex        =   202
            Top             =   780
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No."
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
            Height          =   240
            Index           =   19
            Left            =   4455
            TabIndex        =   201
            Top             =   1050
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax No."
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
            Height          =   240
            Index           =   20
            Left            =   4455
            TabIndex        =   200
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tele Gram No."
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
            Height          =   240
            Index           =   21
            Left            =   4455
            TabIndex        =   199
            Top             =   1590
            Width           =   1200
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mail Id."
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
            Height          =   240
            Index           =   22
            Left            =   4455
            TabIndex        =   198
            Top             =   1860
            Width           =   570
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
            Height          =   240
            Index           =   15
            Left            =   5760
            TabIndex        =   197
            Top             =   240
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
            Height          =   240
            Index           =   14
            Left            =   1350
            TabIndex        =   196
            Top             =   2670
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
            Height          =   240
            Index           =   13
            Left            =   1350
            TabIndex        =   195
            Top             =   2400
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
            Height          =   240
            Index           =   12
            Left            =   1350
            TabIndex        =   194
            Top             =   2130
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
            Height          =   240
            Index           =   11
            Left            =   1350
            TabIndex        =   193
            Top             =   1860
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City Name"
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
            Height          =   240
            Index           =   12
            Left            =   60
            TabIndex        =   192
            Top             =   1860
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pin Code"
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
            Height          =   240
            Index           =   13
            Left            =   75
            TabIndex        =   191
            Top             =   2145
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tin No"
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
            Left            =   60
            TabIndex        =   190
            Top             =   2400
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tin Date"
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
            Left            =   60
            TabIndex        =   189
            Top             =   2670
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CST No. "
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
            Height          =   240
            Index           =   16
            Left            =   4455
            TabIndex        =   188
            Top             =   240
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   8
            Left            =   1350
            TabIndex        =   187
            Top             =   1050
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
            Height          =   240
            Index           =   7
            Left            =   1350
            TabIndex        =   186
            Top             =   780
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
            Height          =   240
            Index           =   6
            Left            =   1350
            TabIndex        =   185
            Top             =   510
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
            Height          =   240
            Index           =   5
            Left            =   1350
            TabIndex        =   184
            Top             =   240
            Width           =   45
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   9
            Left            =   60
            TabIndex        =   183
            Top             =   1050
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Section Name  "
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
            Height          =   240
            Index           =   8
            Left            =   60
            TabIndex        =   182
            Top             =   780
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Short Name  "
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
            Height          =   240
            Index           =   7
            Left            =   60
            TabIndex        =   181
            Top             =   510
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Firm Code  "
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
            Index           =   6
            Left            =   60
            TabIndex        =   180
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00CFE0E0&
         BorderStyle     =   0  'None
         Height          =   4245
         Index           =   2
         Left            =   -74970
         TabIndex        =   138
         Top             =   330
         Width           =   8760
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   98
            Left            =   1515
            MaxLength       =   30
            TabIndex        =   52
            Top             =   3210
            Width           =   975
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   97
            Left            =   1515
            TabIndex        =   51
            Top             =   2940
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Index           =   49
            Left            =   1065
            MaxLength       =   100
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   4095
            Width           =   7080
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   48
            Left            =   1515
            MaxLength       =   80
            TabIndex        =   63
            Top             =   3480
            Width           =   7080
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   47
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   62
            Top             =   2670
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   46
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   61
            Top             =   2400
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   45
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   60
            Top             =   2130
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   44
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   59
            Top             =   1860
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   43
            Left            =   5895
            MaxLength       =   15
            TabIndex        =   58
            Top             =   1590
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   42
            Left            =   5895
            MaxLength       =   15
            TabIndex        =   57
            Top             =   1320
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   41
            Left            =   5895
            MaxLength       =   20
            TabIndex        =   56
            Top             =   1050
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   40
            Left            =   5895
            MaxLength       =   10
            TabIndex        =   55
            Top             =   780
            Width           =   1530
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   39
            Left            =   5895
            TabIndex        =   54
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   38
            Left            =   5895
            MaxLength       =   30
            TabIndex        =   53
            Top             =   240
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   37
            Left            =   1515
            TabIndex        =   50
            Top             =   2670
            Width           =   975
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   36
            Left            =   1515
            MaxLength       =   30
            TabIndex        =   49
            Top             =   2400
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   35
            Left            =   1515
            MaxLength       =   6
            TabIndex        =   48
            Top             =   2130
            Width           =   735
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   34
            Left            =   1515
            MaxLength       =   25
            TabIndex        =   47
            Top             =   1860
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   33
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   46
            Top             =   1590
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   32
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   45
            Top             =   1320
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   31
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   44
            Top             =   1050
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   30
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   43
            Top             =   780
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   29
            Left            =   1515
            MaxLength       =   15
            TabIndex        =   42
            Top             =   510
            Width           =   1650
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Index           =   28
            Left            =   1515
            MaxLength       =   1
            TabIndex        =   41
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LST Date"
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
            Height          =   240
            Index           =   92
            Left            =   75
            TabIndex        =   328
            Top             =   3210
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LST No."
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
            Height          =   240
            Index           =   90
            Left            =   75
            TabIndex        =   327
            Top             =   2940
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   54
            Left            =   1365
            TabIndex        =   326
            Top             =   2940
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
            Height          =   240
            Index           =   33
            Left            =   1365
            TabIndex        =   325
            Top             =   3210
            Width           =   45
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Label6"
            Height          =   225
            Left            =   1515
            TabIndex        =   64
            Top             =   3855
            Width           =   7080
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
            Index           =   49
            Left            =   1350
            TabIndex        =   178
            Top             =   3855
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
            Height          =   240
            Index           =   48
            Left            =   1350
            TabIndex        =   177
            Top             =   3480
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
            Height          =   240
            Index           =   47
            Left            =   5760
            TabIndex        =   176
            Top             =   2670
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
            Height          =   240
            Index           =   46
            Left            =   5760
            TabIndex        =   175
            Top             =   2400
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
            Height          =   240
            Index           =   45
            Left            =   5760
            TabIndex        =   174
            Top             =   2130
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IT Ward No."
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
            Height          =   240
            Index           =   45
            Left            =   4455
            TabIndex        =   173
            Top             =   2130
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IT Ac No."
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
            Height          =   240
            Index           =   46
            Left            =   4455
            TabIndex        =   172
            Top             =   2400
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PAN No."
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
            Height          =   240
            Index           =   48
            Left            =   4455
            TabIndex        =   171
            Top             =   2670
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Speciality"
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
            Height          =   240
            Index           =   47
            Left            =   60
            TabIndex        =   170
            Top             =   3480
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FaDataPath"
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
            Index           =   49
            Left            =   60
            TabIndex        =   169
            Top             =   3855
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
            Height          =   240
            Index           =   44
            Left            =   5760
            TabIndex        =   168
            Top             =   1860
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
            Height          =   240
            Index           =   43
            Left            =   5760
            TabIndex        =   167
            Top             =   1590
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
            Height          =   240
            Index           =   42
            Left            =   5760
            TabIndex        =   166
            Top             =   1320
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
            Height          =   240
            Index           =   41
            Left            =   5760
            TabIndex        =   165
            Top             =   1050
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
            Height          =   240
            Index           =   40
            Left            =   5760
            TabIndex        =   164
            Top             =   780
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
            Height          =   240
            Index           =   39
            Left            =   5760
            TabIndex        =   163
            Top             =   510
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CST Date "
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
            Height          =   240
            Index           =   39
            Left            =   4455
            TabIndex        =   162
            Top             =   510
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile No."
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
            Height          =   240
            Index           =   40
            Left            =   4455
            TabIndex        =   161
            Top             =   780
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No."
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
            Height          =   240
            Index           =   41
            Left            =   4455
            TabIndex        =   160
            Top             =   1050
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax No."
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
            Height          =   240
            Index           =   42
            Left            =   4455
            TabIndex        =   159
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tele Gram No."
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
            Height          =   240
            Index           =   43
            Left            =   4455
            TabIndex        =   158
            Top             =   1590
            Width           =   1200
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mail Id."
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
            Height          =   240
            Index           =   44
            Left            =   4455
            TabIndex        =   157
            Top             =   1860
            Width           =   570
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
            Height          =   240
            Index           =   38
            Left            =   5760
            TabIndex        =   156
            Top             =   240
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
            Height          =   240
            Index           =   37
            Left            =   1350
            TabIndex        =   155
            Top             =   2670
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
            Height          =   240
            Index           =   36
            Left            =   1350
            TabIndex        =   154
            Top             =   2400
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
            Height          =   240
            Index           =   35
            Left            =   1350
            TabIndex        =   153
            Top             =   2130
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
            Height          =   240
            Index           =   34
            Left            =   1350
            TabIndex        =   152
            Top             =   1860
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City Name"
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
            Height          =   240
            Index           =   34
            Left            =   60
            TabIndex        =   151
            Top             =   1860
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pin Code"
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
            Height          =   240
            Index           =   35
            Left            =   60
            TabIndex        =   150
            Top             =   2130
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tin No"
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
            Left            =   60
            TabIndex        =   149
            Top             =   2400
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tin Date"
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
            Left            =   60
            TabIndex        =   148
            Top             =   2670
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CST No. "
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
            Height          =   240
            Index           =   38
            Left            =   4455
            TabIndex        =   147
            Top             =   240
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   31
            Left            =   1350
            TabIndex        =   146
            Top             =   1050
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
            Height          =   240
            Index           =   30
            Left            =   1350
            TabIndex        =   145
            Top             =   780
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
            Height          =   240
            Index           =   29
            Left            =   1350
            TabIndex        =   144
            Top             =   510
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
            Height          =   240
            Index           =   28
            Left            =   1350
            TabIndex        =   143
            Top             =   240
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Firm Code  "
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
            Index           =   28
            Left            =   60
            TabIndex        =   142
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Short Name  "
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
            Height          =   240
            Index           =   29
            Left            =   60
            TabIndex        =   141
            Top             =   510
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Section Name  "
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
            Height          =   240
            Index           =   30
            Left            =   60
            TabIndex        =   140
            Top             =   780
            Width           =   1260
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   31
            Left            =   60
            TabIndex        =   139
            Top             =   1050
            Width           =   690
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00CFE0E0&
         BorderStyle     =   0  'None
         Height          =   4245
         Index           =   3
         Left            =   -74955
         TabIndex        =   11
         Top             =   330
         Width           =   8760
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   100
            Left            =   1515
            MaxLength       =   30
            TabIndex        =   77
            Top             =   3210
            Width           =   975
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   99
            Left            =   1515
            TabIndex        =   76
            Top             =   2940
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H00C000C0&
            Height          =   240
            Index           =   3
            Left            =   1530
            TabIndex        =   90
            Top             =   3990
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H00C000C0&
            Height          =   240
            Index           =   4
            Left            =   4125
            TabIndex        =   91
            Top             =   3990
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H00C000C0&
            Height          =   240
            Index           =   5
            Left            =   7335
            TabIndex        =   92
            Top             =   3990
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Index           =   71
            Left            =   645
            MaxLength       =   100
            TabIndex        =   93
            Top             =   4260
            Width           =   7080
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   70
            Left            =   1515
            MaxLength       =   80
            TabIndex        =   88
            Top             =   3480
            Width           =   7080
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   69
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   87
            Top             =   2670
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   68
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   86
            Top             =   2400
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   67
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   85
            Top             =   2130
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   66
            Left            =   5895
            MaxLength       =   25
            TabIndex        =   84
            Top             =   1860
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   65
            Left            =   5895
            MaxLength       =   15
            TabIndex        =   83
            Top             =   1590
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   64
            Left            =   5895
            MaxLength       =   15
            TabIndex        =   82
            Top             =   1320
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   63
            Left            =   5895
            MaxLength       =   20
            TabIndex        =   81
            Top             =   1050
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   62
            Left            =   5895
            MaxLength       =   10
            TabIndex        =   80
            Top             =   780
            Width           =   1530
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   61
            Left            =   5895
            TabIndex        =   79
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   60
            Left            =   5895
            MaxLength       =   30
            TabIndex        =   78
            Top             =   240
            Width           =   2700
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   59
            Left            =   1515
            TabIndex        =   75
            Top             =   2670
            Width           =   975
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   58
            Left            =   1515
            MaxLength       =   30
            TabIndex        =   74
            Top             =   2400
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   57
            Left            =   1515
            MaxLength       =   6
            TabIndex        =   73
            Top             =   2130
            Width           =   735
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   56
            Left            =   1515
            MaxLength       =   25
            TabIndex        =   72
            Top             =   1860
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   55
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   71
            Top             =   1590
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   54
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   70
            Top             =   1320
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   53
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   69
            Top             =   1050
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   52
            Left            =   1515
            MaxLength       =   40
            TabIndex        =   68
            Top             =   780
            Width           =   2790
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   51
            Left            =   1515
            MaxLength       =   15
            TabIndex        =   67
            Top             =   510
            Width           =   1650
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Index           =   50
            Left            =   1515
            MaxLength       =   1
            TabIndex        =   66
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LST Date"
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
            Height          =   240
            Index           =   95
            Left            =   75
            TabIndex        =   332
            Top             =   3210
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LST No."
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
            Height          =   240
            Index           =   94
            Left            =   75
            TabIndex        =   331
            Top             =   2940
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   56
            Left            =   1365
            TabIndex        =   330
            Top             =   2940
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
            Height          =   240
            Index           =   55
            Left            =   1365
            TabIndex        =   329
            Top             =   3210
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
            Height          =   240
            Index           =   9
            Left            =   1350
            TabIndex        =   234
            Top             =   1860
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "JobCard Sr No"
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
            Index           =   3
            Left            =   60
            TabIndex        =   137
            Top             =   3990
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IPO Gen Sr No"
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
            Index           =   4
            Left            =   2790
            TabIndex        =   136
            Top             =   3990
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IPO  Warranti Sr No. :"
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
            Index           =   5
            Left            =   5505
            TabIndex        =   135
            Top             =   4005
            Visible         =   0   'False
            Width           =   1725
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
            Index           =   2
            Left            =   1350
            TabIndex        =   134
            Top             =   3990
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
            ForeColor       =   &H00C000C0&
            Height          =   240
            Index           =   3
            Left            =   4050
            TabIndex        =   133
            Top             =   3990
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
            Height          =   240
            Index           =   71
            Left            =   1350
            TabIndex        =   132
            Top             =   3750
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
            Height          =   240
            Index           =   70
            Left            =   1350
            TabIndex        =   131
            Top             =   3480
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
            Height          =   240
            Index           =   69
            Left            =   5760
            TabIndex        =   130
            Top             =   2670
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
            Height          =   240
            Index           =   68
            Left            =   5760
            TabIndex        =   129
            Top             =   2400
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
            Height          =   240
            Index           =   67
            Left            =   5760
            TabIndex        =   128
            Top             =   2130
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IT Ward No."
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
            Height          =   240
            Index           =   67
            Left            =   4455
            TabIndex        =   127
            Top             =   2130
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IT Ac No."
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
            Height          =   240
            Index           =   68
            Left            =   4455
            TabIndex        =   126
            Top             =   2400
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PAN No."
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
            Height          =   240
            Index           =   69
            Left            =   4455
            TabIndex        =   125
            Top             =   2670
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Speciality"
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
            Height          =   240
            Index           =   70
            Left            =   60
            TabIndex        =   124
            Top             =   3480
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FaDataPath"
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
            Height          =   240
            Index           =   71
            Left            =   60
            TabIndex        =   123
            Top             =   3750
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
            Height          =   240
            Index           =   66
            Left            =   5760
            TabIndex        =   122
            Top             =   1860
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
            Height          =   240
            Index           =   65
            Left            =   5760
            TabIndex        =   121
            Top             =   1590
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
            Height          =   240
            Index           =   64
            Left            =   5760
            TabIndex        =   120
            Top             =   1320
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
            Height          =   240
            Index           =   63
            Left            =   5760
            TabIndex        =   119
            Top             =   1050
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
            Height          =   240
            Index           =   62
            Left            =   5760
            TabIndex        =   118
            Top             =   780
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
            Height          =   240
            Index           =   61
            Left            =   5760
            TabIndex        =   117
            Top             =   510
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CST Date "
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
            Height          =   240
            Index           =   61
            Left            =   4455
            TabIndex        =   116
            Top             =   510
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile No."
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
            Height          =   240
            Index           =   62
            Left            =   4455
            TabIndex        =   115
            Top             =   780
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No."
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
            Height          =   240
            Index           =   63
            Left            =   4455
            TabIndex        =   114
            Top             =   1050
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax No."
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
            Height          =   240
            Index           =   64
            Left            =   4455
            TabIndex        =   113
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tele Gram No."
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
            Height          =   240
            Index           =   65
            Left            =   4455
            TabIndex        =   112
            Top             =   1590
            Width           =   1200
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mail Id."
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
            Height          =   240
            Index           =   66
            Left            =   4455
            TabIndex        =   111
            Top             =   1860
            Width           =   570
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
            Height          =   240
            Index           =   60
            Left            =   5760
            TabIndex        =   110
            Top             =   240
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
            Height          =   240
            Index           =   59
            Left            =   1350
            TabIndex        =   109
            Top             =   2670
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
            Height          =   240
            Index           =   58
            Left            =   1350
            TabIndex        =   108
            Top             =   2400
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
            Height          =   240
            Index           =   57
            Left            =   1350
            TabIndex        =   107
            Top             =   2130
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City Name"
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
            Height          =   240
            Index           =   56
            Left            =   60
            TabIndex        =   106
            Top             =   1860
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pin Code"
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
            Height          =   240
            Index           =   57
            Left            =   60
            TabIndex        =   105
            Top             =   2130
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tin No."
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
            Index           =   58
            Left            =   60
            TabIndex        =   104
            Top             =   2400
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tin Date"
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
            Index           =   59
            Left            =   60
            TabIndex        =   103
            Top             =   2670
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CST No. "
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
            Height          =   240
            Index           =   60
            Left            =   4455
            TabIndex        =   102
            Top             =   240
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   53
            Left            =   1350
            TabIndex        =   101
            Top             =   1050
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
            Height          =   240
            Index           =   52
            Left            =   1350
            TabIndex        =   100
            Top             =   780
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
            Height          =   240
            Index           =   51
            Left            =   1350
            TabIndex        =   99
            Top             =   510
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
            Height          =   240
            Index           =   50
            Left            =   1350
            TabIndex        =   98
            Top             =   240
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Firm Code  "
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
            Index           =   50
            Left            =   60
            TabIndex        =   97
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Short Name  "
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
            Height          =   240
            Index           =   51
            Left            =   60
            TabIndex        =   96
            Top             =   510
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Section Name  "
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
            Height          =   240
            Index           =   52
            Left            =   60
            TabIndex        =   95
            Top             =   780
            Width           =   1260
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   53
            Left            =   60
            TabIndex        =   94
            Top             =   1050
            Width           =   690
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Label6"
            Height          =   240
            Left            =   1500
            TabIndex        =   89
            Top             =   3750
            Width           =   7080
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   10
         Left            =   -74880
         TabIndex        =   232
         Top             =   1980
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   11
         Left            =   -74910
         TabIndex        =   231
         Top             =   2250
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   33
         Left            =   -74910
         TabIndex        =   230
         Top             =   2250
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   32
         Left            =   -74895
         TabIndex        =   229
         Top             =   1980
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   55
         Left            =   -74910
         TabIndex        =   228
         Top             =   2250
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   54
         Left            =   -74895
         TabIndex        =   227
         Top             =   1980
         Width           =   45
      End
   End
   Begin VB.Frame FrmHlp 
      BackColor       =   &H00CFE0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   1830
      TabIndex        =   3
      Top             =   870
      Visible         =   0   'False
      Width           =   4605
      Begin MSDataGridLib.DataGrid DgHlp 
         Height          =   1320
         Left            =   60
         TabIndex        =   4
         Top             =   360
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   2328
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777152
         ForeColor       =   8388736
         HeadLines       =   0
         RowHeight       =   15
         RowDividerStyle =   0
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
            DataField       =   "AssoComp_Code"
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
         BeginProperty Column01 
            DataField       =   "AssoComp_name"
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
            MarqueeStyle    =   1
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3449.764
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "     Code                     Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   4485
      End
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   810
      Top             =   3750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   14409168
      DialogTitle     =   "Define Central Data Path"
      FileName        =   "Automan"
   End
   Begin VB.Frame FrmLv 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4605
      Left            =   675
      TabIndex        =   1
      Top             =   -60
      Width           =   7380
      Begin VB.Frame FrameCmd 
         BackColor       =   &H00CFE0E0&
         Height          =   780
         Left            =   15
         TabIndex        =   340
         Top             =   3555
         Visible         =   0   'False
         Width           =   7320
         Begin VB.CommandButton cmdApply 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Exit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   5055
            Style           =   1  'Graphical
            TabIndex        =   345
            Top             =   195
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.CommandButton cmdApply 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Account"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   4050
            Style           =   1  'Graphical
            TabIndex        =   344
            Top             =   195
            Width           =   990
         End
         Begin VB.CommandButton cmdApply 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Work Shop"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   343
            Top             =   195
            Width           =   990
         End
         Begin VB.CommandButton cmdApply 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Spare"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2085
            Style           =   1  'Graphical
            TabIndex        =   342
            Top             =   195
            Width           =   990
         End
         Begin VB.CommandButton cmdApply 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Vehicle"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   1095
            Style           =   1  'Graphical
            TabIndex        =   341
            Top             =   195
            Width           =   990
         End
      End
      Begin VB.CommandButton cmdApply 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Set &Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   6315
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3210
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox Txtdate 
         Appearance      =   0  'Flat
         BackColor       =   &H00CFE0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2610
         TabIndex        =   7
         Text            =   "a"
         Top             =   3315
         Width           =   1710
      End
      Begin MSDataGridLib.DataGrid Grid1 
         Height          =   2490
         Left            =   5730
         TabIndex        =   0
         Top             =   570
         Visible         =   0   'False
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   4392
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ColumnHeaders   =   -1  'True
         ForeColor       =   12582912
         HeadLines       =   0
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Div_Code"
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
            DataField       =   "Div_Name"
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
            DataField       =   "Div_SName"
            Caption         =   "S Name"
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
            BeginProperty Column00 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4814.929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid Grid 
         Height          =   2400
         Left            =   45
         TabIndex        =   233
         Top             =   720
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   4233
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ColumnHeaders   =   -1  'True
         ForeColor       =   12582912
         HeadLines       =   0
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Div_Code"
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "Div_SName"
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
            BeginProperty Column00 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4800.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fin Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   4440
         TabIndex        =   317
         Top             =   3315
         Width           =   810
      End
      Begin VB.Label LblDivisionHead 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "     Code                      Division Name                                          Short Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   45
         TabIndex        =   6
         Top             =   420
         Width           =   7290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1260
         TabIndex        =   9
         Top             =   3315
         Width           =   1110
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "List of Division"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Index           =   3
         Left            =   45
         TabIndex        =   2
         Top             =   165
         Width           =   7290
      End
   End
   Begin KeyNo.DatamanKeyNo DatamanKeyNo1 
      Height          =   555
      Left            =   0
      TabIndex        =   346
      Top             =   0
      Visible         =   0   'False
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   979
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "dataman"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   7725
      TabIndex        =   320
      Top             =   5010
      Width           =   1080
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Database.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Left            =   165
      TabIndex        =   318
      Top             =   5070
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Menu POPUP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu MnuAdd 
         Caption         =   "&Add Division"
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "&Edit Division"
      End
      Begin VB.Menu MnuDel 
         Caption         =   "&Delete Division"
      End
      Begin VB.Menu MnuFirm 
         Caption         =   "Add/Edit Associated &Firm"
      End
      Begin VB.Menu MnuPer 
         Caption         =   "Permission"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Dash1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu MnuCancel 
         Caption         =   "&Cancel"
      End
   End
End
Attribute VB_Name = "frmDivision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstHlp As ADODB.Recordset, RstFrm As ADODB.Recordset, RsDiv As ADODB.Recordset
Dim FirmAddFlag As Byte, ADDFLAG As Integer, DivCode As String * 2, mAutoAdd As Boolean, xName As ListItem
Dim RsSite As ADODB.Recordset
Private Const ModuleVeh As Byte = 0, ModuleSpr As Byte = 1, ModuleWrk As Byte = 2
Private Const ModuleAc As Byte = 3, ModuleSetup As Byte = 4, VehPath As Byte = 27
Private Const SprPath As Byte = 49, WrkPath As Byte = 71
Private Const LstNoV    As Byte = 95
Private Const LstDateV  As Byte = 96
Private Const LstNoS    As Byte = 97
Private Const LstDateS  As Byte = 98
Private Const LstNoW    As Byte = 99
Private Const LstDateW  As Byte = 100
Private Const LstNo     As Byte = 101
Private Const LstDate   As Byte = 102


Private Sub cmdApply_Click(Index As Integer)
Label11.Visible = True
Label11.Refresh
If CDate(Txtdate) < PubStartDate Then
    MsgBox "Please enter valid login date", vbInformation, "Incorrect Login Date"
    Txtdate.SetFocus
Else
    Call ApplyModule(Index)
End If
End Sub

Private Sub CmdFirm_Click(Index As Integer)
Dim I As Byte, Firm_Code$, m_Trans As Byte, fob As New FileSystemObject
On Error GoTo ELoop
Select Case Index
Case 0 ' add
    For I = 72 To 102
         txt(I).TEXT = ""
         txt(I).Enabled = True
    Next
    Label5.CAPTION = ""
    txt(93).TEXT = PubCenDataPath
    txt(72).SetFocus
    FirmAddFlag = 1
    For I = 0 To 9
        If I = 3 Or I = 4 Then
            CmdFirm(I).Enabled = True
        Else
            CmdFirm(I).Enabled = False
        End If
    Next
Case 1 'edit
    If RstFrm.RecordCount = 0 Then MsgBox "No Record To edit", vbExclamation, "Massage": Exit Sub
    FirmAddFlag = 2
    For I = 73 To 102
        txt(I).Enabled = True
    Next
    txt(73).SetFocus
    For I = 0 To 9
        If I = 3 Or I = 4 Then
            CmdFirm(I).Enabled = True
        Else
            CmdFirm(I).Enabled = False
        End If
    Next
Case 2 'del
    If RstFrm.RecordCount = 0 Then
        MsgBox "No Records To Delete", vbExclamation, "Massage"
        Exit Sub
    End If
    If GCn.Execute("select count(*) from division where V_SecCompCode = '" & txt(72).TEXT & "' or S_SecCompCode = '" & txt(72).TEXT & "' or W_SecCompCode = '" & txt(72).TEXT & "'").Fields(0).Value > 0 Then
        MsgBox "Can't Delete Firm Is Being Used By Some Division ", vbExclamation, "Delete Error"
        Exit Sub
    End If
    If MsgBox("Are You Sure To Delete ?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.Execute "delete from AssociatedFirms where AssoComp_Code = '" & txt(72).TEXT & "'"
        RstFrm.Requery
        RstHlp.Requery
       If txt(93).TEXT <> "" Then
            Kill Pub_DataPath & "\" & txt(93).TEXT & "\*.*"
            RmDir Pub_DataPath & "\" & txt(93).TEXT
       End If
       Call Move_Frm
       Exit Sub
    End If
Case 3 'save
    If txt(72).TEXT = "" Then MsgBox Label3(72).CAPTION & "Is A Required Field", vbExclamation, "Input Error": txt(72).SetFocus: Exit Sub
    If txt(73).TEXT = "" Then MsgBox Label3(73).CAPTION & "Is A Required Field", vbExclamation, "Input Error": txt(73).SetFocus: Exit Sub
    If txt(74).TEXT = "" Then MsgBox Label3(74).CAPTION & "Is A Required Field", vbExclamation, "Input Error": txt(74).SetFocus: Exit Sub
    If FirmAddFlag = 1 Then
        If GCn.Execute("select count(*) from AssociatedFirms where AssoComp_Code = '" & txt(72) & "' and AssoComp_Name = '" & txt(73).TEXT & "'").Fields(0).Value > 0 Then _
            MsgBox "Duplicate Firm Code For This Company ", vbExclamation, "Input Error": Exit Sub
    End If
    GCn.BeginTrans
    m_Trans = 1
    Firm_Code = txt(72).TEXT
    If FirmAddFlag = 1 Then
        GCn.Execute "insert into AssociatedFirms(AssoComp_Code,AssoComp_SName,AssoComp_Name,Add1,Add2,Add3,City,PinCode,LST,LST_Date, CST, CST_Date, Mobile,Phone,Fax,Gram,MailID,IT_WardNo,IT_AcNo,PAN_No,Speciality,FADataPath, LstNo, LstDate,U_Name,U_EntDt,U_AE) " & _
        " values ('" & Firm_Code & "','" & txt(73).TEXT & "','" & XNull(txt(74).TEXT) & "','" & XNull(txt(75).TEXT) & "','" & XNull(txt(76).TEXT) & "','" & XNull(txt(77).TEXT) & "','" & XNull(txt(78).TEXT) & "','" & XNull(txt(79).TEXT) & "','" & _
        XNull(txt(80).TEXT) & "'," & ConvertDate(txt(81).TEXT) & ",'" & XNull(txt(82).TEXT) & "'," & ConvertDate(txt(83).TEXT) & ",'" & XNull(txt(84).TEXT) & "','" & XNull(txt(85).TEXT) & "','" & XNull(txt(86).TEXT) & "','" & XNull(txt(87).TEXT) & "','" & _
        XNull(txt(88).TEXT) & "','" & XNull(txt(89).TEXT) & "','" & XNull(txt(90).TEXT) & "','" & XNull(txt(91).TEXT) & "','" & (txt(92).TEXT) & "','" & XNull(txt(93).TEXT) & "', '" & txt(LstNo) & "', " & ConvertDate(txt(LstDate)) & ",'" & pubUName & "'," & ConvertDate(Txtdate.TEXT) & ",'A')"
        If txt(93).TEXT <> "" Then
            If fob.FolderExists(Pub_DataPath & "\" & txt(93).TEXT) Then
            Else
                MkDir Pub_DataPath & "\" & txt(93).TEXT
            End If
            FileCopy Pub_DataPath & "\BlankData\FAData.MDB", Pub_DataPath & "\" & txt(93).TEXT & "\FAData.mdb"
        End If
    Else
        GCn.Execute "update AssociatedFirms set  AssoComp_SName='" & txt(73).TEXT & "', AssoComp_Name='" & txt(74).TEXT & "', " & _
            "    Add1='" & txt(75).TEXT & "', Add2='" & txt(76).TEXT & "', Add3='" & txt(77).TEXT & "', " & _
            "    City='" & txt(78).TEXT & "' , PinCode='" & txt(79).TEXT & "', LST='" & txt(80).TEXT & "',LST_Date=" & ConvertDate(txt(81).TEXT) & ", LstNo='" & txt(LstNo) & "', LstDate=" & ConvertDate(LstDate) & ", CST='" & txt(82).TEXT & "', CST_Date=" & ConvertDate(txt(83).TEXT) & ", Mobile='" & txt(84).TEXT & "', Phone='" & txt(85).TEXT & "', " & _
            "    Fax='" & txt(86).TEXT & "' , Gram='" & txt(87).TEXT & "',MailID='" & txt(88).TEXT & "', IT_WardNo='" & txt(89).TEXT & "',IT_AcNo='" & txt(90).TEXT & "', PAN_No='" & txt(91).TEXT & "', Speciality='" & txt(92).TEXT & "', " & _
            "   U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(Txtdate.TEXT) & ",U_AE = 'E' where AssoComp_Code='" & Firm_Code & "'"
    End If
    GCn.CommitTrans
    m_Trans = 0
    RstFrm.Requery
    RstHlp.Requery
    If FirmAddFlag = 1 Then FirmAddFlag = 0: Call CmdFirm_Click(0): Exit Sub
    RstFrm.FIND "AssoComp_Code = '" & txt(72).TEXT & "'"
    Call Move_Frm
    For I = 72 To 102
        txt(I).Enabled = False
    Next
    For I = 0 To 9
        If I = 3 Or I = 4 Then
          CmdFirm(I).Enabled = False
        Else
          CmdFirm(I).Enabled = True
        End If
    Next
Case 4 'cancel
    For I = 72 To 102
        txt(I).Enabled = False
    Next
    For I = 0 To 9
        If I = 3 Or I = 4 Then
          CmdFirm(I).Enabled = False
        Else
          CmdFirm(I).Enabled = True
        End If
    Next
    Call Move_Frm
    FirmAddFlag = 0
Case 5 'exit
    FrmFirm.Visible = False
    If RsDiv.RecordCount <= 0 Then Call MnuAdd_Click
Case 6  'first
    If RstFrm.RecordCount = 0 Then Exit Sub
    If RstFrm.AbsolutePosition > 1 Then RstFrm.MoveFirst: Move_Frm
Case 7  'prev
    If RstFrm.RecordCount = 0 Then Exit Sub
    If RstFrm.AbsolutePosition > 1 Then RstFrm.MovePrevious: Move_Frm
Case 8  'next
    If RstFrm.RecordCount = 0 Then Exit Sub
    If RstFrm.AbsolutePosition < RstFrm.RecordCount Then RstFrm.MoveNext: Move_Frm
Case 9  'last
    If RstFrm.RecordCount = 0 Then Exit Sub
    If RstFrm.AbsolutePosition < RstFrm.RecordCount Then RstFrm.MoveLast: Move_Frm
End Select
Exit Sub
ELoop:
If m_Trans = 1 Then GCn.RollbackTrans
CheckError
End Sub
Private Sub DgHlp_Click()
    If STab.Tab = 1 Then
        txt(6).TEXT = RstHlp!AssoComp_Code
    ElseIf STab.Tab = 2 Then
        txt(28).TEXT = RstHlp!AssoComp_Code
    ElseIf STab.Tab = 3 Then
        txt(50).TEXT = RstHlp!AssoComp_Code
    End If
End Sub
Private Sub DGHlp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        FrmHlp.Visible = False
        If RstHlp.RecordCount = 0 Then Exit Sub
        Call FillData
'        Exit Sub       by lps 28-01-02
   If STab.Tab = 1 Then
        txt(7).SetFocus
    ElseIf STab.Tab = 2 Then
        txt(29).SetFocus
    ElseIf STab.Tab = 3 Then
        txt(51).SetFocus
    End If
End If
End Sub

Private Sub DgHlp_KeyUp(KeyCode As Integer, Shift As Integer)
    If STab.Tab = 1 Then
        txt(6).TEXT = RstHlp!AssoComp_Code
    ElseIf STab.Tab = 2 Then
        txt(28).TEXT = RstHlp!AssoComp_Code
    ElseIf STab.Tab = 3 Then
        txt(50).TEXT = RstHlp!AssoComp_Code
    End If
End Sub
Private Sub DgHlp_LostFocus()
    FrmHlp.Visible = False
End Sub

Private Sub DgSite_KeyPress(KeyAscii As Integer)
    RsSite.MoveFirst
    RsSite.FIND "Name Like '" & Chr(KeyAscii) & "*" & "'"
    If RsSite.EOF = True Then RsSite.MoveFirst
End Sub

Private Sub Form_Activate()
On Error GoTo ELoop
    If RsDiv.RecordCount = 0 Then Call MnuAdd_Click
    LblHelp(3).CAPTION = "Divisions of " & Me.CAPTION
    Grid.SetFocus
    
Exit Sub
ELoop:
    MsgBox err.Description
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ADDFLAG = 1 Or ADDFLAG = 2 Then Exit Sub
    If Button = 2 And PubULabel = "Y" Then
       MENUENABLE True
       PopupMenu popup
     End If
End Sub
                
Private Sub Frame1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
       MENUENABLE False
       PopupMenu popup
End If
End Sub
'    Div_Code , Div_Name, Div_Sname, JobCard_SrlNo, IPO_Gen_SrlNo, IPO_War_SrlNo,
'    V_SecCompCode, V_SecSName, V_SecName, V_SecAdd1, V_SecAdd2, V_SecAdd3,
'    V_SecCity , V_SecPinCode, V_SecLST, V_SecLST_Date, V_SecCST, V_SecCST_Date, V_SecMobile, V_SecPhone,
'    V_SecFax , V_SecGram, V_SecMailID, V_SecIT_WardNo, V_SecIT_AcNo, V_SecPAN_No, V_SecSpeciality, V_SecFADataPath,
'    S_SecCompCode, S_SecSName, S_SecName, S_SecAdd1,S_SecAdd2,S_SecAdd3,S_SecCity,S_SecPinCode, S_SecLST,  S_SecLST_Date,S_SecCST,S_SecCST_Date, S_SecMob,S_SecPhone,S_SecFax,S_SecGram,S_SecMailID,
'    S_SecIT_WardNo,S_SecIT_AcNo,S_SecPAN_No,S_SecSpeciality,S_SecFADataPath,
'    W_SecCompCode,W_SecSName,W_SecName,W_SecAdd1,W_SecAdd2,W_SecAdd3,W_SecCity, W_SecPinCode,  W_SecLST , W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecMobile,W_SecPhone,W_SecFax,W_SecGram,
'    W_SecMailID,W_SecIT_WardNo,W_SecIT_AcNo,W_SecPAN_No,W_SecSpeciality,W_SecFADataPath,
'    S_SecDefultSprGodown1 ,S_SecDefultSprGodown2,W_SecDefultSprGodown1 , U_Name, U_EntDt, U_AE, Trf_Date

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 10 Then
        LblSiteList = "Sites Of " & Me.CAPTION
        FrameSite.left = FrmLv.left
        FrameSite.top = FrmLv.top
        FrameSite.Visible = True
        FrameCmd.Visible = True
        DGSite.SetFocus
    End If
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ADDFLAG = 1 Or ADDFLAG = 2 Then Exit Sub
    If Button = 2 And PubULabel = "Y" Then
       MENUENABLE True
       PopupMenu popup
     End If
End Sub

Private Sub ListView_Click()
    SelectFAData
End Sub
Private Sub ListView_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SelectFAData
End Sub
Private Sub MnuAdd_Click()
On Error GoTo eloop1
Dim I As Byte
If RstFrm.RecordCount = 0 Then
    MsgBox "Open Associated Firm", vbExclamation, "Massage"
    mAutoAdd = True
    MnuFirm_Click
    Exit Sub
End If
ADDFLAG = 1
STab.Visible = True
STab.ZOrder 0
Disp_Text (True)
txt(27).Enabled = False
txt(49).Enabled = False
txt(71).Enabled = False
Label6.CAPTION = ""
Label7.CAPTION = ""
Label8.CAPTION = ""
STab.Tab = 0
BlankText
txt(0).SetFocus
MENUENABLE False
Exit Sub
eloop1:  Call CheckError
End Sub
Private Sub MnuEdit_Click()
On Error GoTo eloop1
Dim I As Integer
ADDFLAG = 2
Call MoveRec
STab.Visible = True
STab.ZOrder 0
Disp_Text (True)
txt(0).Enabled = False
txt(27).Enabled = False
txt(49).Enabled = False
txt(71).Enabled = False
STab.Tab = 0
txt(1).SetFocus
MENUENABLE False
eloop1:    Call CheckError
End Sub
Private Sub MnuDel_Click()
On Error GoTo eloop1
If RsDiv.RecordCount <= 0 Then
    MsgBox "No Records To Delete", vbExclamation, "Massage"
    Exit Sub
End If
If MsgBox("Are You Sure To Delete ?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
   GCn.Execute ("delete from division where Div_Code='" & RsDiv!Div_Code & "'")
   G_CompCn.Execute "update user1 set div_code ='',div_name ='' ,mod_veh=0,mod_spr=0,mod_wsp=0,mod_acc=0,mod_set=0  where comp_code='" & PubCenCompCode & "' and Div_Code = '" & RsDiv!Div_Code & "'"
   G_CompCn.Execute "delete from  user2 where comp_code='" & PubCenCompCode & "' and Div_Code = '" & RsDiv!Div_Code & "'"
   RsDiv.Requery
   Grid.Refresh
End If
eloop1:    CheckError
End Sub
Private Sub MNUCANCEL_Click()
If FrmHlp.Visible = True Then FrmHlp.Visible = False
ADDFLAG = 0
MENUENABLE True
STab.Visible = False
End Sub
Private Sub MnuFirm_Click()
Dim I As Byte
FrmFirm.Visible = True
FrmFirm.ZOrder 0
For I = 72 To 102
    txt(I).Enabled = False
Next
If mAutoAdd = True Then
    CmdFirm_Click (0)
    mAutoAdd = False
Else
    Move_Frm
End If

End Sub
Private Sub MnuPer_Click()
'        If Form_Chk("USER PERMISSIONS") = True Then Exit Sub
'    Load frmUser
'    frmUser.Show
'    Set frmUser = Nothing
End Sub
Private Sub MNUSAVE_Click()
Dim I As Byte, RstMod As Recordset
On Error GoTo ELoop
If FrmHlp.Visible = True Then FrmHlp.Visible = False
For I = 0 To 2
    If txt(I).TEXT = "" Then MsgBox Label3(I).CAPTION & "Is A Required Field", vbExclamation, "Input Error": STab.Tab = 0: txt(I).SetFocus: Exit Sub
Next
If ADDFLAG = 1 Then
    If GCn.Execute("select count(*) from division where Div_Code = '" & txt(0).TEXT & "'").Fields(0).Value > 0 Then _
        MsgBox "Duplicate Division Code For This Company ", vbExclamation, "Input Error": STab.Tab = 0: txt(0).SetFocus: Exit Sub
End If
GCn.BeginTrans
    DivCode = txt(0).TEXT
    If ADDFLAG = 1 Then
        GCn.Execute "insert into division(Div_Code,ProductSerial,Div_Sname, Div_Name ," & _
            "    V_SecCompCode, V_SecSName, V_SecName, V_SecAdd1, V_SecAdd2, V_SecAdd3, " & _
            "    V_SecCity , V_SecPinCode, V_SecLST, V_SecLST_Date, V_SecCST, V_SecCST_Date, V_SecMobile, V_SecPhone, " & _
            "    V_SecFax , V_SecGram, V_SecMailID, V_SecIT_WardNo, V_SecIT_AcNo, V_SecPAN_No, V_SecSpeciality, V_SecFADataPath, " & _
            "    S_SecCompCode, S_SecSName, S_SecName, S_SecAdd1,S_SecAdd2,S_SecAdd3,S_SecCity,S_SecPinCode, S_SecLST,  S_SecLST_Date,S_SecCST,S_SecCST_Date, S_SecMobile,S_SecPhone,S_SecFax,S_SecGram,S_SecMailID, " & _
            "    S_SecIT_WardNo,S_SecIT_AcNo,S_SecPAN_No,S_SecSpeciality,S_SecFADataPath, " & _
            "    W_SecCompCode,W_SecSName,W_SecName,W_SecAdd1,W_SecAdd2,W_SecAdd3,W_SecCity, W_SecPinCode,  W_SecLST , W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecMobile,W_SecPhone,W_SecFax,W_SecGram, " & _
            "    W_SecMailID,W_SecIT_WardNo,W_SecIT_AcNo,W_SecPAN_No,W_SecSpeciality,W_SecFADataPath, LstNoV, LstDateV, LstNoS, LstDateV, LstNoW, LstDateW, U_Name,U_EntDt,U_AE) " & _
            " values('" & txt(0).TEXT & "','" & txt(94).TEXT & "','" & XNull(txt(1).TEXT) & "','" & XNull(txt(2).TEXT) & "','" & XNull(txt(6).TEXT) & "','" & XNull(txt(7).TEXT) & "','" & XNull(txt(8).TEXT) & "','" & XNull(txt(9).TEXT) & "','" & XNull(txt(10).TEXT) & _
            "', '" & XNull(txt(11).TEXT) & "','" & XNull(txt(12).TEXT) & "','" & XNull(txt(13).TEXT) & "','" & XNull(txt(14).TEXT) & "'," & ConvertDate(txt(15).TEXT) & ",'" & XNull(txt(16).TEXT) & "'," & ConvertDate(txt(17).TEXT) & ",'" & XNull(txt(18).TEXT) & "','" & XNull(txt(19).TEXT) & "','" & XNull(txt(20).TEXT) & _
            "', '" & XNull(txt(21).TEXT) & "','" & XNull(txt(22).TEXT) & "','" & XNull(txt(23).TEXT) & "','" & XNull(txt(24).TEXT) & "','" & XNull(txt(25).TEXT) & "','" & XNull(txt(26).TEXT) & "','" & XNull(txt(27).TEXT) & "','" & XNull(txt(28).TEXT) & "','" & XNull(txt(29).TEXT) & "','" & XNull(txt(30).TEXT) & _
            "', '" & XNull(txt(31).TEXT) & "','" & XNull(txt(32).TEXT) & "','" & XNull(txt(33).TEXT) & "','" & XNull(txt(34).TEXT) & "','" & XNull(txt(35).TEXT) & "','" & XNull(txt(36).TEXT) & "'," & ConvertDate(txt(37).TEXT) & ",'" & XNull(txt(38).TEXT) & "'," & ConvertDate(txt(39).TEXT) & ",'" & XNull(txt(40).TEXT) & _
            "', '" & XNull(txt(41).TEXT) & "','" & XNull(txt(42).TEXT) & "','" & XNull(txt(43).TEXT) & "','" & XNull(txt(44).TEXT) & "','" & XNull(txt(45).TEXT) & "','" & XNull(txt(46).TEXT) & "','" & XNull(txt(47).TEXT) & "','" & XNull(txt(48).TEXT) & "','" & XNull(txt(49).TEXT) & "','" & XNull(txt(50).TEXT) & _
            "', '" & XNull(txt(51).TEXT) & "','" & XNull(txt(52).TEXT) & "','" & XNull(txt(53).TEXT) & "','" & XNull(txt(54).TEXT) & "','" & XNull(txt(55).TEXT) & "','" & XNull(txt(56).TEXT) & "','" & XNull(txt(57).TEXT) & "','" & XNull(txt(58).TEXT) & "'," & ConvertDate(txt(59).TEXT) & ",'" & XNull(txt(60).TEXT) & _
            "', " & ConvertDate(txt(61).TEXT) & ",'" & XNull(txt(62).TEXT) & "','" & XNull(txt(63).TEXT) & "','" & XNull(txt(64).TEXT) & "','" & XNull(txt(65).TEXT) & "','" & XNull(txt(66).TEXT) & "','" & XNull(txt(67).TEXT) & "','" & XNull(txt(68).TEXT) & "','" & XNull(txt(69).TEXT) & "','" & XNull(txt(70).TEXT) & _
            "','" & XNull(txt(71).TEXT) & "', '" & txt(LstNoV) & "', " & ConvertDate(txt(LstDateV)) & ", '" & txt(LstNoS) & "', " & ConvertDate(txt(LstDateS)) & ", '" & txt(LstNoW) & "', " & ConvertDate(txt(LstDateW)) & ", '" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
    Else
        GCn.Execute "update division set ProductSerial='" & txt(94).TEXT & "',Div_Name='" & XNull(txt(2).TEXT) & "', Div_Sname='" & XNull(txt(1).TEXT) & "', V_SeccompCode='" & XNull(txt(6).TEXT) & "', V_SecSName='" & XNull(txt(7).TEXT) & _
            "', V_SecName='" & XNull(txt(8).TEXT) & "', V_SecAdd1='" & XNull(txt(9).TEXT) & "', V_SecAdd2='" & XNull(txt(10).TEXT) & "', V_SecAdd3='" & XNull(txt(11).TEXT) & "', V_SecCity='" & XNull(txt(12).TEXT) & "' , V_SecPinCode='" & XNull(txt(13).TEXT) & "', V_SecLST='" & XNull(txt(14).TEXT) & "', V_SecLST_Date=" & ConvertDate(txt(15).TEXT) & _
            ",  V_SecCST='" & XNull(txt(16).TEXT) & "', V_SecCST_Date=" & ConvertDate(txt(17).TEXT) & ", V_SecMobile='" & XNull(txt(18).TEXT) & "', V_SecPhone='" & XNull(txt(19).TEXT) & "', V_SecFax='" & XNull(txt(20).TEXT) & "' , V_SecGram='" & XNull(txt(21).TEXT) & "', V_SecMailID='" & XNull(txt(22).TEXT) & "', V_SecIT_WardNo='" & XNull(txt(23).TEXT) & _
            "', V_SecIT_AcNo='" & XNull(txt(24).TEXT) & "', V_SecPAN_No='" & XNull(txt(25).TEXT) & "', V_SecSpeciality='" & XNull(txt(26).TEXT) & "',V_SecFADataPath = '" & XNull(txt(27).TEXT) & "', S_Seccompcode='" & XNull(txt(28).TEXT) & "',S_SecSName='" & XNull(txt(29).TEXT) & "', S_SecName='" & XNull(txt(30).TEXT) & "', S_SecAdd1='" & XNull(txt(31).TEXT) & _
            "', S_SecAdd2='" & XNull(txt(32).TEXT) & "',S_SecAdd3='" & XNull(txt(33).TEXT) & "',S_SecCity='" & XNull(txt(34).TEXT) & "',S_SecPinCode='" & XNull(txt(35).TEXT) & "', S_SecLST='" & XNull(txt(36).TEXT) & "',  S_SecLST_Date=" & ConvertDate(txt(37).TEXT) & ",S_SecCST='" & XNull(txt(38).TEXT) & "',S_SecCST_Date=" & ConvertDate(txt(39).TEXT) & _
            ",  S_SecMobile='" & XNull(txt(40).TEXT) & "',S_SecPhone='" & XNull(txt(41).TEXT) & "',S_SecFax='" & XNull(txt(42).TEXT) & "',S_SecGram='" & XNull(txt(43).TEXT) & "',S_SecMailID='" & XNull(txt(44).TEXT) & "', S_SecIT_WardNo='" & XNull(txt(45).TEXT) & "',S_SecIT_AcNo='" & XNull(txt(46).TEXT) & "',S_SecPAN_No='" & XNull(txt(47).TEXT) & _
            "', S_SecSpeciality='" & XNull(txt(48).TEXT) & "',S_SecFADataPath = '" & XNull(txt(49).TEXT) & "', W_Seccompcode='" & XNull(txt(50).TEXT) & "',W_SecSName='" & XNull(txt(51).TEXT) & "',W_SecName='" & XNull(txt(52).TEXT) & "',W_SecAdd1='" & XNull(txt(53).TEXT) & "',W_SecAdd2='" & XNull(txt(54).TEXT) & "',W_SecAdd3='" & XNull(txt(55).TEXT) & _
            "', W_SecCity='" & XNull(txt(56).TEXT) & "', W_SecPinCode='" & XNull(txt(57).TEXT) & "',  W_SecLST='" & XNull(txt(58).TEXT) & "' , W_SecLST_Date=" & ConvertDate(txt(59).TEXT) & ",W_SecCST='" & XNull(txt(60).TEXT) & "',W_SecCST_Date=" & ConvertDate(txt(61).TEXT) & ",W_SecMobile='" & XNull(txt(62).TEXT) & "',W_SecPhone='" & XNull(txt(63).TEXT) & _
            "', W_SecFax='" & XNull(txt(64).TEXT) & "',W_SecGram='" & XNull(txt(65).TEXT) & "', W_SecMailID='" & XNull(txt(66).TEXT) & "',W_SecIT_WardNo='" & XNull(txt(67).TEXT) & "',W_SecIT_AcNo='" & XNull(txt(68).TEXT) & "',W_SecPAN_No='" & XNull(txt(69).TEXT) & "',W_SecSpeciality='" & XNull(txt(70).TEXT) & "',W_SecFADataPath = '" & XNull(txt(71).TEXT) & _
            "', LstNoV='" & txt(LstNoV) & "', LstDateV=" & ConvertDate(txt(LstDateV)) & ", LstNoS='" & txt(LstNoS) & "', LstDateS=" & ConvertDate(txt(LstDateS)) & ", LstNoW='" & txt(LstNoW) & "', LstDateW=" & ConvertDate(txt(LstDateW)) & " " & _
            ", U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE= 'E' where Div_Code='" & txt(0).TEXT & "'"
    End If
    If G_CompCn.Execute("select count(*) from user1 where comp_code ='" & PubCenCompCode & "' and Div_Code = ''  and user_name= '" & pubUName & "'").Fields(0).Value = 0 Then
        G_CompCn.Execute "delete from user1 where comp_code = '" & PubCenCompCode & "' and div_code = '" & txt(0).TEXT & "' and user_name= '" & pubUName & "'"
        G_CompCn.Execute ("insert into user1(user_name,comp_code,div_code,div_name,mod_veh,mod_spr,mod_wsp,Mod_Acc) values('" & pubUName & "','" & PubCenCompCode & "','" & txt(0).TEXT & "','" & txt(1).TEXT & "'," & IIf(txt(6).TEXT <> "", 1, 0) & "," & IIf(txt(28).TEXT <> "", 1, 0) & "," & IIf(txt(50).TEXT <> "", 1, 0) & " ,1)")
    Else
        G_CompCn.Execute ("update user1 set div_code='" & txt(0).TEXT & "',div_name='" & txt(1).TEXT & "',mod_veh=" & IIf(txt(6).TEXT <> "", 1, 0) & ",mod_spr=" & IIf(txt(28).TEXT <> "", 1, 0) & ",mod_wsp=" & IIf(txt(50).TEXT <> "", 1, 0) & ",Mod_Acc=1 where comp_code ='" & PubCenCompCode & "' and Div_Code = ''  and user_name= '" & pubUName & "'")
    End If
    If pubUName <> "SA" Then
        If G_CompCn.Execute("select count(*) from user1 where comp_code ='" & PubCenCompCode & "' and Div_Code = ''  and user_name= 'SA'").Fields(0).Value = 0 Then
            G_CompCn.Execute "delete from user1 where comp_code = '" & PubCenCompCode & "' and div_code = '" & txt(0).TEXT & "' and user_name= 'SA'"
            G_CompCn.Execute ("insert into user1(user_name,comp_code,div_code,div_name,mod_veh,mod_spr,mod_wsp) values('SA','" & PubCenCompCode & "','" & txt(0).TEXT & "','" & txt(1).TEXT & "'," & IIf(txt(6).TEXT <> "", 1, 0) & "," & IIf(txt(28).TEXT <> "", 1, 0) & "," & IIf(txt(50).TEXT <> "", 1, 0) & " )")
        Else
            G_CompCn.Execute ("update user1 set div_code='" & txt(0).TEXT & "',div_name='" & txt(1).TEXT & "',mod_veh=" & IIf(txt(6).TEXT <> "", 1, 0) & ",mod_spr=" & IIf(txt(28).TEXT <> "", 1, 0) & ",mod_wsp=" & IIf(txt(50).TEXT <> "", 1, 0) & " where comp_code ='" & PubCenCompCode & "' and Div_Code = ''  and user_name= 'SA'")
        End If
    End If
        G_CompCn.Execute "delete from user2 where  comp_code = '" & PubCenCompCode & "' and div_code = '" & txt(0).TEXT & "' and user_name= '" & pubUName & "'"
        If txt(6).TEXT <> "" Then
            Set RstMod = New Recordset
            RstMod.CursorLocation = adUseClient
            
            RstMod.Open "select * from user_module where module_name='Vehicle' order by srno", G_CompCn, adOpenDynamic, adLockOptimistic
            Do Until RstMod.EOF
              G_CompCn.Execute ("insert into user2(comp_code,div_code,user_name,Module_Name,form_code,param_str) values('" & PubCenCompCode & "' ,'" & txt(0).TEXT & "','" & pubUName & "','Vehicle' ,'" & RstMod!Form_Code & "','ADEP')")
            RstMod.MoveNext
            Loop
        End If
        If txt(28).TEXT <> "" Then
            Set RstMod = New Recordset
            RstMod.CursorLocation = adUseClient
            RstMod.Open "select * from user_module  where Module_Name='Spare' order by srno", G_CompCn, adOpenStatic, adLockReadOnly
            Do Until RstMod.EOF
                G_CompCn.Execute ("insert into user2(comp_code,div_code,user_name,Module_Name,form_code,param_str) values('" & PubCenCompCode & "' ,'" & txt(0).TEXT & "','" & pubUName & "','Spare' ,'" & RstMod!Form_Code & "','ADEP')")
                RstMod.MoveNext
            Loop
        End If
        If txt(50).TEXT <> "" Then
            Set RstMod = New Recordset
            RstMod.CursorLocation = adUseClient
            RstMod.Open "select * from user_module  where Module_Name='Workshop' order by srno", G_CompCn, adOpenStatic, adLockReadOnly
            Do Until RstMod.EOF
                G_CompCn.Execute ("insert into user2(comp_code,div_code,user_name,Module_Name,form_code,param_str) values('" & PubCenCompCode & "' ,'" & txt(0).TEXT & "','" & pubUName & "','Workshop' ,'" & RstMod!Form_Code & "','ADEP')")
                RstMod.MoveNext
            Loop
       End If
       If pubUName <> "SA" Then
        G_CompCn.Execute "delete from user2 where  comp_code = '" & PubCenCompCode & "' and div_code = '" & txt(0).TEXT & "' and user_name= 'SA'"
        If txt(6).TEXT <> "" Then
            Set RstMod = New Recordset
            RstMod.CursorLocation = adUseClient
            RstMod.Open "select * from user_module  where Module_Name='Vehicle' order by srno", G_CompCn, adOpenStatic, adLockReadOnly
            Do Until RstMod.EOF
              G_CompCn.Execute ("insert into user2(comp_code,div_code,user_name,Module_Name,form_code,param_str) values('" & PubCenCompCode & "' ,'" & txt(0).TEXT & "','SA','Vehicle' ,'" & RstMod!Form_Code & "','ADEP')")
            RstMod.MoveNext
            Loop
       End If
        If txt(28).TEXT <> "" Then
            Set RstMod = New Recordset
            RstMod.CursorLocation = adUseClient
            RstMod.Open "select * from user_module  where Module_Name='Spare' order by srno", G_CompCn, adOpenStatic, adLockReadOnly
            Do Until RstMod.EOF
              G_CompCn.Execute ("insert into user2(comp_code,div_code,user_name,Module_Name,form_code,param_str) values('" & PubCenCompCode & "' ,'" & txt(0).TEXT & "','SA','Spare' ,'" & RstMod!Form_Code & "','ADEP')")
            RstMod.MoveNext
            Loop
       End If
        If txt(50).TEXT <> "" Then
            Set RstMod = New Recordset
            RstMod.CursorLocation = adUseClient
            RstMod.Open "select * from user_module  where Module_Name='Workshop' order by srno", G_CompCn, adOpenStatic, adLockReadOnly
            Do Until RstMod.EOF
              G_CompCn.Execute ("insert into user2(comp_code,div_code,user_name,Module_Name,form_code,param_str) values('" & PubCenCompCode & "' ,'" & txt(0).TEXT & "','SA','Workshop' ,'" & RstMod!Form_Code & "','ADEP')")
            RstMod.MoveNext
            Loop
       End If
       End If
    GCn.CommitTrans
    ADDFLAG = 0
    MENUENABLE True
    RsDiv.Requery
    Grid.Refresh
    STab.Visible = False
    Exit Sub
ELoop:
GCn.RollbackTrans
 CheckError
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     Call MoveRec
End Sub

Private Sub STab_Click(PreviousTab As Integer)
      If FrmHlp.Visible = True Then FrmHlp.Visible = False
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Call Ctrl_GetFocus(Index)
    FrmHlp.Visible = False
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Integer
If (KeyCode = 13 Or KeyCode = 40) And Index = 92 Then
    If MsgBox("Save Record?", vbYesNo, "Save Entry") = vbYes Then
        Call CmdFirm_Click(3)
        Exit Sub
    Else
        Call CmdFirm_Click(4)
        Exit Sub
    End If
End If

If (KeyCode = 13 Or KeyCode = 40) And Index = 5 Then
    If MsgBox("Save Record?", vbYesNo, "Save Entry") = vbYes Then
        Call MNUSAVE_Click
        Exit Sub
    Else
        Call MNUCANCEL_Click
        Exit Sub
    End If
End If
Select Case Index
    Case 6, 28, 50
    If KeyCode = 13 Then
            FrmHlp.Visible = False
            If Trim(txt(Index)) = "" Then
                Call FillData
                If STab.Tab = 1 Then
                    txt(7).SetFocus
                ElseIf STab.Tab = 2 Then
                    txt(29).SetFocus
                ElseIf STab.Tab = 3 Then
                    txt(51).SetFocus
                End If
                Exit Sub
           Else
            If RstHlp.RecordCount = 0 Then Exit Sub
                Call FillData
                If STab.Tab = 1 Then
                    txt(7).SetFocus
                ElseIf STab.Tab = 2 Then
                    txt(29).SetFocus
                ElseIf STab.Tab = 3 Then
                    txt(51).SetFocus
                End If
                Exit Sub
           End If
    End If
    If KeyCode = 16 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If FrmHlp.Visible = False Then GoTo NXT
    Else
        FrmHlp.Visible = True
        FrmHlp.ZOrder 0
    End If
    If FrmHlp.Visible = True Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
            Select Case KeyCode
                Case vbKeyUp
                    If RstHlp.AbsolutePosition > 1 Then
                        RstHlp.MovePrevious
                    Else
                     KeyCode = 0
                    End If
                Case vbKeyDown
                    If RstHlp.AbsolutePosition < RstHlp.RecordCount Then RstHlp.MoveNext
                Case vbKeyPageUp '33
                    For I = 1 To 10
                        If RstHlp.AbsolutePosition > 1 Then RstHlp.MovePrevious
                    Next
                Case vbKeyPageDown '34
                    For I = 1 To 10
                        If RstHlp.AbsolutePosition < RstHlp.RecordCount Then RstHlp.MoveNext
                    Next
            End Select
            If RstHlp.BOF = False And RstHlp.EOF = False Then
                If STab.Tab = 1 Then
                    txt(6).TEXT = RstHlp!AssoComp_Code
                ElseIf STab.Tab = 2 Then
                    txt(28).TEXT = RstHlp!AssoComp_Code
                ElseIf STab.Tab = 3 Then
                    txt(50).TEXT = RstHlp!AssoComp_Code
                End If
            End If
      End If
      Exit Sub
  End If
End Select
NXT:
If KeyCode = 40 Then   'keydown = 40
    If Index = 2 Then STab.Tab = 1: SendKeysA vbKeyTab, True: KeyCode = 0: Exit Sub
    If Index = 26 Then STab.Tab = 2: SendKeysA vbKeyTab, True: KeyCode = 0: Exit Sub
    If Index = 48 Then STab.Tab = 3: SendKeysA vbKeyTab, True: KeyCode = 0: Exit Sub
    If Index <> 92 And Index <> 5 Then SendKeysA vbKeyTab, True: KeyCode = 0: Exit Sub
ElseIf KeyCode = 38 And Index = 6 Then
    STab.Tab = 0: SendKeys "+{Tab}": KeyCode = 0: Exit Sub
ElseIf KeyCode = 38 And Index = 28 Then
     STab.Tab = 1: SendKeys "+{Tab}": KeyCode = 0: Exit Sub
ElseIf KeyCode = 38 And Index = 50 Then
     STab.Tab = 2: SendKeys "+{Tab}": KeyCode = 0: Exit Sub
ElseIf KeyCode = 38 And ADDFLAG = 1 And Index <> 0 Then    'keyup = 38
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = 38 And FirmAddFlag = 1 And Index <> 72 Then   'keyup = 38
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = 38 And ADDFLAG = 2 And Index <> 1 Then     'keyup = 38
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = 38 And FirmAddFlag = 2 And Index <> 73 Then    'keyup = 38
    SendKeys "+{Tab}"
    KeyCode = 0
End If

If KeyCode = 13 And Index = 2 Then STab.Tab = 1: SendKeysA vbKeyTab, True: Exit Sub
If KeyCode = 13 And Index = 26 Then STab.Tab = 2: SendKeysA vbKeyTab, True: Exit Sub
If KeyCode = 13 And Index = 48 Then STab.Tab = 3: SendKeysA vbKeyTab, True: Exit Sub
If KeyCode = 13 And Index <> 5 Then SendKeysA vbKeyTab, True: Exit Sub
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
CheckQuote KeyAscii
Select Case Index
    Case 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 3, 4, 5
        Call NumPress(txt(Index), KeyAscii, 8, 0)
    Case 72 'Firm code
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Integer
Select Case Index
  Case 6, 28, 50
    If FrmHlp.Visible = True Then
        If Trim(txt(Index)) = "" Then Exit Sub
           If RstHlp.RecordCount = 0 Then Exit Sub
           RstHlp.MoveFirst
           RstHlp.FIND "AssoComp_Code  >='" & FilterString(txt(Index)) & "'"
           If RstHlp.EOF = True Then RstHlp.MoveFirst
'          Exit Sub
    End If
End Select
KeyCode = 0
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    FrmHlp.Visible = False
    Call Ctrl_validate(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Dim I As Byte
Case 6, 28, 50
    If RstHlp.RecordCount = 0 Then Exit Sub
    If FrmHlp.Visible = True Then Call FillData
Case 72, 73, 74
    If Index = 72 And RstFrm.RecordCount > 0 Then
        If GCn.Execute("select count(*) from AssociatedFirms where AssoComp_Code = '" & txt(72).TEXT & "'").Fields(0).Value > 0 Then MsgBox "Duplicate Code", vbExclamation, "Validation Check": Cancel = True: txt(72).SetFocus: Exit Sub
    End If
    txt(93).TEXT = PubCenDataPath & "\" & "FaData" & txt(72).TEXT & Right(PubStartDate, 4)
    Label5.CAPTION = Pub_DataPath & "\" & txt(93).TEXT

'    If Txt(Index).Text = "" Then
'        MsgBox Label3(Index) & " Is Required", vbExclamation, "Validation Check"
'        Cancel = True
'    End If
Case 81, 83, 59, 61, 37, 39, 15, 17, LstDateV, LstDateS, LstDateW, LstDate
    txt(Index).TEXT = RetDate(txt(Index))
End Select
End Sub

'******* Fuctions **********

Private Sub Ctrl_validate(Index As Integer)
txt(Index).BackColor = CtrlBColOrg
txt(Index).ForeColor = CtrlFColOrg
End Sub

Private Sub Ctrl_GetFocus(Index As Integer)
txt(Index).BackColor = CtrlBCol
txt(Index).ForeColor = CtrlFCol
End Sub

Private Sub BlankText()
Dim I As Byte
For I = 0 To 93
    txt(I).TEXT = ""
Next I
End Sub

Private Sub ApplyModule(Index As Integer)
On Error Resume Next
ProgressBar1.Visible = True
Select Case Index
    Case 5    'Exit
         End
'         Unload Me
'         AddFlag = 0
'         FrmCompany.Show
'         Exit Sub
    Case ModuleVeh, ModuleSpr, ModuleWrk, ModuleAc ' 0, 1, 2,3
        PubDealerID = IIf(IsNull(RsDiv!Dealer_ID), "", RsDiv!Dealer_ID)
        pubLockDate = "31/Mar/2004"
        PubDivCode = RsDiv!Div_Code
        PubDivSName = IIf(IsNull(RsDiv!Div_SName), "", RsDiv!Div_SName)
        PubVCompCode = IIf(IsNull(RsDiv!v_SecCompCode), "", RsDiv!v_SecCompCode)
        PubSCompCode = IIf(IsNull(RsDiv!s_SecCompCode), "", RsDiv!s_SecCompCode)
        PubWCompCode = IIf(IsNull(RsDiv!w_SecCompCode), "", RsDiv!w_SecCompCode)
        PubWSecFaDataPath = IIf(IsNull(RsDiv!w_SecFaDataPath), "", RsDiv!w_SecFaDataPath)
'        If Index <> ModuleAc Then
            If txt(VehPath).TEXT <> "" Then
                PubVFADataPath = Pub_DataPath & "\" & txt(VehPath).TEXT & "\FaData.mdb"
                Set GCnFaV = New ADODB.Connection
                With GCnFaV
                    .CursorLocation = adUseClient
                    If PubBackEnd = "A" Then
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .ConnectionString = "Data Source=" & PubVFADataPath & ";Persist Security Info=False"
                    Else
                        .CommandTimeout = 1024
                        If PubDbUser <> "" Then
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        Else
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        End If
                    End If
                    .Open
                End With
            End If
            If txt(SprPath).TEXT <> "" Then
                PubSFADataPath = Pub_DataPath & "\" & txt(SprPath).TEXT & "\FaData.mdb"
                Set GCnFaS = New ADODB.Connection
                With GCnFaS
                    .CursorLocation = adUseClient
                    If PubBackEnd = "A" Then
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .ConnectionString = "Data Source=" & PubSFADataPath & ";Persist Security Info=False"
                    Else
                        .CommandTimeout = 1024
                        If PubDbUser <> "" Then
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        Else
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        End If
                    End If
                        
                    .Open
                End With
            End If
            If txt(WrkPath).TEXT <> "" Then
                 PubWFADataPath = Pub_DataPath & "\" & txt(WrkPath).TEXT & "\FaData.mdb"
                 Set GCnFaW = New ADODB.Connection
                 With GCnFaW
                     .CursorLocation = adUseClient
                     If PubBackEnd = "A" Then
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .ConnectionString = "Data Source=" & PubWFADataPath & ";Persist Security Info=False"
                    Else
                        .CommandTimeout = 1024
                        If PubDbUser <> "" Then
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        Else
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        End If
                    End If
                        
                     .Open
                 End With
            End If
'        End If

        Select Case Index
            Case ModuleVeh
                PubLoginModule = "Vehicle"
                If RsDiv!v_SecCompCode = "" Then MsgBox "Access Denied", vbExclamation, "Information": Exit Sub
                PubFirmCode = RsDiv!v_SecCompCode
                PubComp_Name = RsDiv!v_SecName
                PubComp_Add = RsDiv!v_SecAdd1
                PubComp_Add2 = XNull(RsDiv!v_SecAdd2)
                PubComp_City = XNull(RsDiv!v_SecCity)
                PubSecName = "Vehicle"
                PubComp_Contact = "PHONE : " & XNull(RsDiv!V_SecPhone) & " Fax   : " & XNull(RsDiv!V_SecFax)
                
                
                Set G_FaCn = New ADODB.Connection
'                G_FACN = GCnFaV:  G_FACN.Open
                With G_FaCn
                    .CursorLocation = adUseClient
                    If PubBackEnd = "A" Then
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .ConnectionString = "Data Source=" & PubVFADataPath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
                    Else
                        .CommandTimeout = 1024
                        If PubDbUser <> "" Then
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        Else
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        End If
                    End If
                    .Open
                End With
                PubFADataPath = PubVFADataPath
                'Licence checking
                
                
                DatamanKeyNo1.ModuleCode = "AMW"
                DatamanKeyNo1.FirmName = RsDiv!v_SecName
                'DatamanKeyNo1.CityName = RsDiv!v_SecCity
                DatamanKeyNo1.CityName = GCn.Execute("select City from AssociatedFirms where AssoComp_Code='" & PubFirmCode & "'").Fields(0).Value
                If DatamanKeyNo1.Validate(IIf(IsNull(RsDiv!ProductSerial), "", RsDiv!ProductSerial)) = False Then
                    If Not StrCmp(left(PubComp_Name, 6), "PRAYAG") Then
                        MsgBox "Dataman Demo Product ID : " & DatamanKeyNo1.ReturnID, vbInformation
                        If GCn.Execute("select count(*) from Veh_Stock").Fields(0).Value > 100 Then
                            TrialEnd.Show vbModal
                            End
                        ElseIf G_FaCn.Execute("select count(*) from Ledger").Fields(0).Value > 2000 Then
                            TrialEnd.Show vbModal
                            End
                        ElseIf GCn.Execute("select count(*) from SP_Stock").Fields(0).Value > 5000 Then
                            TrialEnd.Show vbModal
                        ElseIf GCn.Execute("select count(*) from Job_Card").Fields(0).Value > 100 Then
                            TrialEnd.Show vbModal
                        ElseIf GCn.Execute("select count(*) from SP_Sale").Fields(0).Value > 200 Then
                            TrialEnd.Show vbModal
                            End
                        End If
                    End If
                End If
                'eof licence
'                If pubUName <> "SA" Then
'                    MDIForm1.MnuSpr.Visible = False
'                    MDIForm1.MnuWorks.Visible = False
'                    MDIForm1.Fa.Visible = False
'                End If
            Case ModuleSpr
                PubLoginModule = "Spare"
                If RsDiv!s_SecCompCode = "" Then MsgBox "Access Denied", vbExclamation, "Information": Exit Sub
                PubFirmCode = XNull(RsDiv!s_SecCompCode)
                PubComp_Name = XNull(RsDiv!s_SecName)
                PubComp_Add = XNull(RsDiv!s_SecAdd1)
                PubComp_Add2 = XNull(RsDiv!s_SecAdd2)
                PubComp_City = XNull(RsDiv!s_SecCity)
                PubSecName = "Spare"
                PubComp_Contact = "PHONE : " & XNull(RsDiv!S_SecPhone) & " Fax   : " & XNull(RsDiv!S_SecFax)
                Set G_FaCn = New ADODB.Connection
'                G_FACN = GCnFaS:  G_FACN.Open
                 With G_FaCn
                    .CursorLocation = adUseClient
                    If PubBackEnd = "A" Then
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .ConnectionString = "Data Source=" & PubSFADataPath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
                    ElseIf PubBackEnd = "S" Then
                        .CommandTimeout = 1024
                        If PubDbUser <> "" Then
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        Else
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        End If
                    End If
                    .Open
                End With
                PubFADataPath = PubSFADataPath
                'Licence checking
                DatamanKeyNo1.ModuleCode = "AMW"
                DatamanKeyNo1.FirmName = RsDiv!s_SecName
                'DatamanKeyNo1.CityName = RsDiv!s_SecCity
                DatamanKeyNo1.CityName = GCn.Execute("select City from AssociatedFirms where AssoComp_Code='" & PubFirmCode & "'").Fields(0).Value
                If DatamanKeyNo1.Validate(IIf(IsNull(RsDiv!ProductSerial), "", RsDiv!ProductSerial)) = False Then
                    MsgBox "Dataman Demo Product ID : " & DatamanKeyNo1.ReturnID, vbInformation
                    If GCn.Execute("select count(*) from SP_Stock").Fields(0).Value > 500 Then End
                End If
                'eof licence
                
'                If pubUName <> "SA" Then
'                    MDIForm1.MnuVeh.Visible = False
'                    MDIForm1.MnuWorks.Visible = False
'                    MDIForm1.Fa.Visible = False
'                End If
            Case ModuleWrk
                PubLoginModule = "Workshop"
                If RsDiv!w_SecCompCode = "" Then MsgBox "Access Denied", vbExclamation, "Information": Exit Sub
                PubFirmCode = RsDiv!w_SecCompCode
                PubComp_Name = RsDiv!w_SecName
                PubComp_Add = XNull(RsDiv!w_SecAdd1)
                PubComp_Add2 = XNull(RsDiv!w_SecAdd2)
                PubComp_City = XNull(RsDiv!w_SecCity)
                PubSecName = "WorkShop"
                PubComp_Contact = "PHONE : " & XNull(RsDiv!W_SecPhone) & " Fax   : " & XNull(RsDiv!W_SecFax)
                Set G_FaCn = New ADODB.Connection
'                G_FACN = GCnFaW:  G_FACN.Open
                 With G_FaCn
                    .CursorLocation = adUseClient
                    If PubBackEnd = "A" Then
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .ConnectionString = "Data Source=" & PubWFADataPath & ";Persist Security Info=False"
                    ElseIf PubBackEnd = "S" Then
                        .CommandTimeout = 1024
                        If PubDbUser <> "" Then
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        Else
                    
                            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                        End If
                    End If
                    
                    .Open
                End With
                PubFADataPath = PubWFADataPath
                'Licence checking
                DatamanKeyNo1.ModuleCode = "AMW"
                DatamanKeyNo1.FirmName = RsDiv!w_SecName
                'DatamanKeyNo1.CityName = "RsDiv!w_SecCity
                DatamanKeyNo1.CityName = GCn.Execute("select City from AssociatedFirms where AssoComp_Code='" & PubFirmCode & "'").Fields(0).Value
                If DatamanKeyNo1.Validate(IIf(IsNull(RsDiv!ProductSerial), "", RsDiv!ProductSerial)) = False Then
                    MsgBox "Dataman Demo Product ID : " & DatamanKeyNo1.ReturnID, vbInformation
                    If GCn.Execute("select count(*) from Job_Card").Fields(0).Value > 100 Then End
                End If
'                'eof licence
''                If pubUName <> "SA" Then
''                    MDIForm1.MnuVeh.Visible = False
''                    MDIForm1.MnuSpr.Visible = False
''                    MDIForm1.Fa.Visible = False
'                End If
            Case ModuleAc
                PubLoginModule = "Accounts"
                PubVFADataPath = IIf(txt(VehPath).TEXT <> "", Pub_DataPath & "\" & txt(VehPath).TEXT & "\FaData.mdb", "")
                PubSFADataPath = IIf(txt(SprPath).TEXT <> "", Pub_DataPath & "\" & txt(SprPath).TEXT & "\FaData.mdb", "")
                PubWFADataPath = IIf(txt(WrkPath).TEXT <> "", Pub_DataPath & "\" & txt(WrkPath).TEXT & "\FaData.mdb", "")
                
 '               If (PubVFADataPath <> PubSFADataPath) Or (PubVFADataPath <> PubWFADataPath) Then
                        ListView.ListItems.Clear
                        Dim ix As Byte
                        If Trim(PubVFADataPath) <> "" Then
                            ix = ix + 1
                            Set xName = ListView.ListItems.Add(ix, , "Vehicle")
                        End If
                        If Trim(PubSFADataPath) <> "" Then
                            ix = ix + 1
                            Set xName = ListView.ListItems.Add(ix, , "Spare")
                        End If
                        If Trim(PubWFADataPath) <> "" Then
                            ix = ix + 1
                            Set xName = ListView.ListItems.Add(ix, , "Works")
                        End If
                        xName.EnsureVisible
                        xName.SELECTED = True
    
                    If ListView.ListItems.Count > 0 Then
                        FrmList.top = 1900
                        FrmList.left = 3405
                        ListView.top = 0
                        ListView.left = 0
                        ListView.width = FrmList.width
                        ListView.height = FrmList.height
                        ListView.ColumnHeaders(1).width = 1000
                        
                        FrmList.Visible = True
                        FrmList.ZOrder 0
                        ListView.SetFocus
                    End If
                    Exit Sub
'                Else
'                    SelectFAData
'                    Exit Sub
'                End If
                
                
'                PubFirmCode = RsDiv!W_SecCompCode
'                PubComp_Name = RsDiv!W_SecName
'                PubComp_Add = XNull(RsDiv!W_SecAdd1)
'                PubComp_Add2 = XNull(RsDiv!W_SecAdd2)
'                PubComp_City = XNull(RsDiv!W_SecCity) & IIf(RsDiv!W_SecPinCode = "", "", "-") & XNull(RsDiv!W_SecPinCode)
'
'                PubSecName = "Account"
'                If Txt(VehPath).Text <> "" Then
'                    Set G_FACN = New ADODB.Connection
'                    With G_FACN
'                        .CursorLocation = adUseClient
'                        .Provider = "Microsoft.Jet.OLEDB.4.0"
'                        .ConnectionString = "Data Source=" & PubVFADataPath & ";Persist Security Info=False"
'                        .Open
'                    End With
'                ElseIf Txt(SprPath).Text <> "" Then
'                    Set G_FACN = New ADODB.Connection
'                    With G_FACN
'                        .CursorLocation = adUseClient
'                        .Provider = "Microsoft.Jet.OLEDB.4.0"
'                        .ConnectionString = "Data Source=" & PubSFADataPath & ";Persist Security Info=False"
'                        .Open
'                    End With
'                ElseIf Txt(WrkPath).Text <> "" Then
'                    Set G_FACN = New ADODB.Connection
'                    With G_FACN
'                        .CursorLocation = adUseClient
'                        .Provider = "Microsoft.Jet.OLEDB.4.0"
'                        .ConnectionString = "Data Source=" & PubWFADataPath & ";Persist Security Info=False"
'                        .Open
'                    End With
'                End If
'                If pubUName <> "SA" Then
'                    MDIForm1.MnuVeh.Visible = False
'                    MDIForm1.MnuSpr.Visible = False
'                    MDIForm1.MnuWorks.Visible = False
'                End If
        
        End Select
                
        PubDefine
        Initialise_Pub
        'If PubBackEnd = "S" Then UpdateTableStructureSql
        ProgressBar1.Value = 30
        'If pubUName = "SA" Then
            MDIForm1.Disp_Menu
            MDIForm1.MnuVeh.Visible = cmdApply(ModuleVeh).Visible
            MDIForm1.MnuSpr.Visible = cmdApply(ModuleSpr).Visible
            MDIForm1.MnuWorks.Visible = cmdApply(ModuleWrk).Visible
            MDIForm1.fa.Visible = cmdApply(ModuleAc).Visible
            MDIForm1.DTools.Visible = True
        'End If
End Select

    
    

    MDIForm1.AllowModuleVeh = cmdApply(ModuleVeh).Visible
    MDIForm1.AllowModuleSpr = cmdApply(ModuleSpr).Visible
    MDIForm1.AllowModuleWrk = cmdApply(ModuleWrk).Visible
    MDIForm1.CAPTION = PubPackage & IIf(PubBackEnd = "S", "(SQL)", "") & " - [" & RsDiv!Div_SName & "] " & PubComp_Name
    ProgressBar1.Value = 45
                
    If StrCmp(left(PubComp_Name, 6), "Prayag") Then
        PubMoveRecYn = False
    Else
        PubMoveRecYn = True
    End If
                
    Set PubRsSyctrl = GCn.Execute("Select * from Syctrl")
    Set PubRsCompany = G_CompCn.Execute("Select * from Syctrl")
    
    'Unload frmDivision
    frmDivision.Hide
    
    MDIForm1.CAPTION = MDIForm1.CAPTION & " - " & PubSiteName
    MDIForm1.Show

'   MDIForm1.Caption = frmDivision.Caption & "/" & RsDiv!Div_Sname & "/" & PubComp_Name
''    Set rdApp = CreateObject("CrystalRuntime.Application")
    Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub MoveRec()
On Error GoTo err
Dim I As Byte, ButEnabled As Integer, CmdEnable As Integer, TotalWidth As Double, dispFirst As Boolean
Dim TotalEnabled As Integer
 For I = 0 To 3
   cmdApply(I).Visible = False
 Next
 For I = 4 To 4
   cmdApply(I).Enabled = True
 Next
 If PubULabel = "Y" Then
    MnuPer.Visible = True
Else
    MnuPer.Visible = False
End If

If RsDiv.RecordCount > 0 Then
    If G_CompCn.Execute("select div_code from user1 where user_name = '" & pubUName & "' and Div_Code = '" & RsDiv!Div_Code & "'").RecordCount = 0 Then Exit Sub
    If G_CompCn.Execute("select mod_veh from user1 where comp_code='" & PubCenCompCode & "' and Div_Code = '" & RsDiv!Div_Code & "' and user_name = '" & pubUName & "'").Fields(0).Value = 1 Then cmdApply(ModuleVeh).Visible = True
    If G_CompCn.Execute("select mod_spr from user1 where comp_code='" & PubCenCompCode & "' and Div_Code = '" & RsDiv!Div_Code & "' and user_name = '" & pubUName & "'").Fields(0).Value = 1 Then cmdApply(ModuleSpr).Visible = True
    If G_CompCn.Execute("select mod_wsp from user1 where comp_code='" & PubCenCompCode & "' and Div_Code = '" & RsDiv!Div_Code & "' and user_name = '" & pubUName & "'").Fields(0).Value = 1 Then cmdApply(ModuleWrk).Visible = True
    If G_CompCn.Execute("select Mod_Acc from user1 where comp_code='" & PubCenCompCode & "' and Div_Code = '" & RsDiv!Div_Code & "' and user_name = '" & pubUName & "'").Fields(0).Value = 1 Then cmdApply(ModuleAc).Visible = True
    
    txt(0).TEXT = RsDiv!Div_Code
    txt(94).TEXT = XNull(RsDiv!ProductSerial)
    txt(2).TEXT = RsDiv!Div_Name: txt(1).TEXT = RsDiv!Div_SName
    txt(6).TEXT = IIf(IsNull(RsDiv!v_SecCompCode), "", RsDiv!v_SecCompCode): txt(7).TEXT = IIf(IsNull(RsDiv!V_SecSName), "", RsDiv!V_SecSName): txt(8).TEXT = IIf(IsNull(RsDiv!v_SecName), "", RsDiv!v_SecName): txt(9).TEXT = IIf(IsNull(RsDiv!v_SecAdd1), "", RsDiv!v_SecAdd1): txt(10).TEXT = IIf(IsNull(RsDiv!v_SecAdd2), "", RsDiv!v_SecAdd2): txt(11).TEXT = IIf(IsNull(RsDiv!V_SecAdd3), "", RsDiv!V_SecAdd3)
    txt(12).TEXT = IIf(IsNull(RsDiv!v_SecCity), "", RsDiv!v_SecCity): txt(13).TEXT = IIf(IsNull(RsDiv!v_SecPinCode), "", RsDiv!v_SecPinCode): txt(14).TEXT = IIf(IsNull(RsDiv!V_SecLST), "", RsDiv!V_SecLST): txt(15).TEXT = XNull(RsDiv!V_SecLST_Date): txt(LstNoV).TEXT = IIf(IsNull(RsDiv!LstNoV), "", RsDiv!LstNoV): txt(LstDateV).TEXT = XNull(RsDiv!LstDateV): txt(16).TEXT = IIf(IsNull(RsDiv!V_SecCST), "", RsDiv!V_SecCST): txt(17).TEXT = XNull(RsDiv!V_SecCST_Date): txt(18).TEXT = IIf(IsNull(RsDiv!V_SecMobile), "", RsDiv!V_SecMobile): txt(19).TEXT = IIf(IsNull(RsDiv!V_SecPhone), "", RsDiv!V_SecPhone)
    txt(20).TEXT = IIf(IsNull(RsDiv!V_SecFax), "", RsDiv!V_SecFax): txt(21).TEXT = XNull(RsDiv!V_SecGram): txt(22).TEXT = XNull(RsDiv!V_SecMailID): txt(23).TEXT = XNull(RsDiv!V_SecIT_WardNo): txt(24).TEXT = XNull(RsDiv!V_SecIT_AcNo): txt(25).TEXT = XNull(RsDiv!V_SecPAN_No): txt(26).TEXT = XNull(RsDiv!V_SecSpeciality): txt(27).TEXT = XNull(RsDiv!V_SecFADataPath)
    txt(28).TEXT = XNull(RsDiv!s_SecCompCode): txt(29).TEXT = XNull(RsDiv!S_SecSName): txt(30).TEXT = XNull(RsDiv!s_SecName): txt(31).TEXT = XNull(RsDiv!s_SecAdd1): txt(32).TEXT = XNull(RsDiv!s_SecAdd2): txt(33).TEXT = XNull(RsDiv!S_SecAdd3): txt(34).TEXT = XNull(RsDiv!s_SecCity): txt(35).TEXT = XNull(RsDiv!s_SecPinCode): txt(36).TEXT = XNull(RsDiv!S_SecLST): txt(37).TEXT = XNull(RsDiv!S_SecLST_Date): txt(LstNoS).TEXT = XNull(RsDiv!LstNoS): txt(LstDateS).TEXT = XNull(RsDiv!LstDateS): txt(38).TEXT = XNull(RsDiv!S_SecCST): txt(39).TEXT = XNull(RsDiv!S_SecCST_Date): txt(40).TEXT = XNull(RsDiv!S_SecMobile): txt(41).TEXT = XNull(RsDiv!S_SecPhone): txt(42).TEXT = XNull(RsDiv!S_SecFax): txt(43).TEXT = XNull(RsDiv!S_SecGram): txt(44).TEXT = XNull(RsDiv!S_SecMailID)
    txt(45).TEXT = XNull(RsDiv!S_SecIT_WardNo): txt(46).TEXT = XNull(RsDiv!S_SecIT_AcNo): txt(47).TEXT = XNull(RsDiv!S_SecPAN_No): txt(48).TEXT = XNull(RsDiv!S_SecSpeciality): txt(49).TEXT = XNull(RsDiv!S_SecFADataPath)
    txt(50).TEXT = XNull(RsDiv!w_SecCompCode): txt(51).TEXT = XNull(RsDiv!W_SecSName): txt(52).TEXT = XNull(RsDiv!w_SecName): txt(53).TEXT = XNull(RsDiv!w_SecAdd1): txt(54).TEXT = XNull(RsDiv!w_SecAdd2): txt(55).TEXT = XNull(RsDiv!W_SecAdd3): txt(56).TEXT = XNull(RsDiv!w_SecCity): txt(57).TEXT = XNull(RsDiv!w_SecPinCode): txt(58).TEXT = XNull(RsDiv!W_SecLST): txt(59).TEXT = XNull(RsDiv!W_SecLST_Date): txt(LstNoW).TEXT = XNull(RsDiv!LstNoW): txt(LstDateW).TEXT = XNull(RsDiv!LstDateW): txt(60).TEXT = XNull(RsDiv!W_SecCST): txt(61).TEXT = XNull(RsDiv!W_SecCST_Date): txt(62).TEXT = XNull(RsDiv!W_SecMobile): txt(63).TEXT = XNull(RsDiv!W_SecPhone): txt(64).TEXT = XNull(RsDiv!W_SecFax): txt(65).TEXT = XNull(RsDiv!W_SecGram)
    txt(66).TEXT = XNull(RsDiv!W_SecMailID): txt(67).TEXT = XNull(RsDiv!W_SecIT_WardNo): txt(68).TEXT = XNull(RsDiv!W_SecIT_AcNo): txt(69).TEXT = XNull(RsDiv!W_SecPAN_No): txt(70).TEXT = XNull(RsDiv!W_SecSpeciality): txt(71).TEXT = XNull(RsDiv!w_SecFaDataPath)
    Label6.CAPTION = IIf(txt(27) <> "", Pub_DataPath & "\" & txt(27).TEXT, "")
    Label7.CAPTION = IIf(txt(49) <> "", Pub_DataPath & "\" & txt(49).TEXT, "")
    Label8.CAPTION = IIf(txt(71) <> "", Pub_DataPath & "\" & txt(71).TEXT, "")
End If
For I = 0 To 4
   If cmdApply(I).Enabled = True Then
      TotalEnabled = TotalEnabled + 1
      TotalWidth = TotalWidth + cmdApply(I).width
   End If
Next
For I = 0 To 4
   If cmdApply(I).Enabled = True Then
      If dispFirst = False Then
         CmdEnable = I
         cmdApply(I).left = (8200 / 2) - (TotalWidth / 2)
         dispFirst = True
      Else
        cmdApply(I).left = cmdApply(CmdEnable).left + cmdApply(CmdEnable).width
        CmdEnable = I
      End If
   End If
Next

Exit Sub
err:    CheckError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then
 If MsgBox("Are You Sure To Quit ?", vbYesNo + vbCritical + vbDefaultButton2, "Message Box !") = vbYes Then
    ADDFLAG = 0
    End
 End If
End If
Exit Sub
ELoop:      MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub MENUENABLE(Enb As Boolean)
    frmDivision.MnuAdd.Enabled = Enb
    frmDivision.MnuEdit.Enabled = Enb
    frmDivision.MnuDel.Enabled = Enb
    frmDivision.MnuFirm.Enabled = Enb
    frmDivision.MnuPer.Enabled = False 'Enb
    frmDivision.MnuSave.Enabled = Not Enb
    frmDivision.MnuCancel.Enabled = Not Enb
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim Condstr$
Dim I As Integer
For I = 0 To 93
    txt(I).BackColor = CtrlBColOrg
    txt(I).ForeColor = CtrlFColOrg
    txt(I).BackColor = CtrlBColOrg
    txt(I).ForeColor = CtrlFColOrg
Next
    MENUENABLE True
    FrmFirm.top = 0
    FrmFirm.left = 0
    STab.left = (Me.width - (STab.width + 90)) / 2    '   FrmAddComp.Left = 55
    STab.top = 100          '   FrmAddComp.Top = 100
    STab.Visible = False    '   FrmAddComp.Visible = False
    FrmLv.left = 885
    FrmLv.top = 180
    FrmHlp.left = 2700
    FrmHlp.top = 1110
    Grid.left = 45
    Grid.top = 705
    
    Call LST_Fields_Add
    
    Set RstFrm = New ADODB.Recordset
    With RstFrm
        If RstFrm.State <> 0 Then RstFrm.Close
        
        .ActiveConnection = GCn
        .LockType = adLockOptimistic
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .Open "select * from AssociatedFirms  order by AssoComp_Code"
    End With
    
'   Set RstFrm = GCn.Execute("select * from AssociatedFirms  order by AssoComp_Code")
    RstFrm.Requery
    FirmAddFlag = 0
    Set RstHlp = New ADODB.Recordset
    With RstHlp
        .ActiveConnection = GCn
        .LockType = adLockOptimistic
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .Open "Select * From AssociatedFirms Order by AssoComp_Code"
    End With
'   Set RstHlp = GCn.Execute("Select * From AssociatedFirms Order by AssoComp_Code")
    Set DgHlp.DataSource = RstHlp
    Set RsDiv = New Recordset
'   If UCase(PubUName) = "SA" Then
        RsDiv.CursorLocation = adUseClient
        RsDiv.Open "Select * from Division ", GCn, adOpenStatic, adLockReadOnly
'   Else
'       RsDiv.Open "Select * from Division where U_Name ='" & PubUName & "')", GCn, adOpenStatic, adLockReadOnly
'   End If
   Set Grid.DataSource = RsDiv
   Txtdate.TEXT = Format(date, "dd/mmm/yyyy")
   Label10 = Format(PubStartDate, "yyyy") & "-" & Format(PubEndDate, "yyyy")
   Call MoveRec
   
   

    
    LblSiteList = "Sites/Branches Of " & PubComp_Name & " - " & "[" & PubCenCompCode & "]"
    If pubUName <> "SA" Then
    
        Dim RsTemp As ADODB.Recordset
        Dim mSiteStr As String
        If StrCmp(left(PubComp_Name, 6), "Prayag") Then
            Set RsTemp = G_CompCn.Execute("DECLARE @temp NVARCHAR(1000) " & _
                                        "SET @temp='' " & _
                                        "select @temp= @temp + '`' +  User_site.site_code + '`,' from User_site  where User_Name ='" & pubUName & "'  " & _
                                        "IF LEN(@TEMP)>0 SET @temp=substring (@temp,1,len(@Temp)-1) " & _
                                        "SELECT @temp= 'SELECT ''' + @temp +'''' " & _
                                        "EXEC sys.sp_executesql @temp ")
            mSiteStr = Replace(IIf(XNull(RsTemp.Fields(0)) = "", "''", XNull(RsTemp.Fields(0))), "`", "'")
        Else
            mSiteStr = ""
            Set RsTemp = G_CompCn.Execute("select User_site.site_code from User_site  where User_Name ='" & pubUName & "'  ")
            For I = 0 To RsTemp.RecordCount - 1
                mSiteStr = mSiteStr & "'" & XNull(RsTemp!Site_Code) & "'" & IIf(I = RsTemp.RecordCount - 1, "", ",")
                RsTemp.MoveNext
            Next I
        End If
    
        If PubBackEnd = "A" Then
            Set RsSite = GCn.Execute("Select S.Site_Code As Code, S.Site_Desc As Name From Site S  Where Site_Code In (" & mSiteStr & ") Order By S.Site_Desc")
        Else
            
            Set RsSite = GCn.Execute("Select S.Site_Code As Code, S.Site_Desc As Name, S.Address1, S.Address2, S.Address3, S.City, S.Phone, S.Mobile,s.sitetype From Site S  Where Site_Code In (" & mSiteStr & ") Order By S.Site_Desc")
        End If
    Else
        If PubBackEnd = "A" Then
            Set RsSite = GCn.Execute("Select S.Site_Code As Code, S.Site_Desc As Name From Site S Order By S.Site_Desc")
        Else
            Set RsSite = GCn.Execute("Select S.Site_Code As Code, S.Site_Desc As Name, S.Address1, S.Address2, S.Address3, S.City, S.Phone, S.Mobile,s.sitetype From Site S Order By S.Site_Desc")
        End If
    End If
    Set DGSite.DataSource = RsSite
    If RsSite.RecordCount > 0 Then
        RsSite.FIND "Code='" & PubSiteCode & "'"
        If RsSite.EOF = True Then RsSite.MoveFirst
    Else
        MsgBox "Sorry! You Have No Permission Of Any Site to Login"
        End
    End If
   
   
   
   
   
   
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Txtdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeysA vbKeyTab, True
End Sub
Private Sub Txtdate_Validate(Cancel As Boolean)
    Txtdate.TEXT = RetDate(Txtdate)
    Cancel = Not CheckFinYear(Txtdate)
End Sub
Private Sub Move_Frm()
If RstFrm.RecordCount > 0 Then
    txt(72).TEXT = RstFrm!AssoComp_Code
    txt(73).TEXT = RstFrm!AssoComp_SName
    txt(74).TEXT = RstFrm!AssoComp_Name
    txt(75).TEXT = RstFrm!Add1
    txt(76).TEXT = RstFrm!Add2: txt(77).TEXT = RstFrm!Add3
    txt(78).TEXT = RstFrm!City: txt(79).TEXT = RstFrm!PinCode
    txt(80).TEXT = RstFrm!LST: txt(81).TEXT = XNull(RstFrm!LST_Date)
    txt(LstNo) = XNull(RstFrm!LstNo): txt(LstDate) = XNull(RstFrm!LstDate)
    txt(82).TEXT = RstFrm!CST: txt(83).TEXT = XNull(RstFrm!CST_Date)
    txt(84).TEXT = RstFrm!Mobile: txt(85).TEXT = RstFrm!Phone
    txt(86).TEXT = RstFrm!FAx:   txt(87).TEXT = RstFrm!Gram
    txt(88).TEXT = RstFrm!MailID: txt(89).TEXT = RstFrm!IT_WardNo
    txt(90).TEXT = RstFrm!IT_AcNo: txt(91).TEXT = RstFrm!PAN_No
    txt(92).TEXT = RstFrm!Speciality
    txt(93).TEXT = RstFrm!FADataPath
    Label5.CAPTION = Pub_DataPath & "\" & txt(93).TEXT
End If
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
For I = 0 To txt.Count - 1
 txt(I).Enabled = Enb
Next
FrameCmd.Visible = False
End Sub

Private Sub FillData()
  If RstHlp.RecordCount > 0 Then
  If RstHlp.BOF = True Or RstHlp.EOF = True Then Exit Sub
    If STab.Tab = 1 Then
        If txt(6) <> "" Then
            txt(6) = RstHlp!AssoComp_Code:  txt(7) = RstHlp!AssoComp_SName
            txt(8) = RstHlp!AssoComp_Name:  txt(9) = XNull(RstHlp!Add1)
            txt(10) = XNull(RstHlp!Add2):   txt(11) = XNull(RstHlp!Add3)
            txt(12) = XNull(RstHlp!City):   txt(13) = XNull(RstHlp!PinCode)
            txt(14) = XNull(RstHlp!LST):    txt(15) = XNull(RstHlp!LST_Date)
            txt(LstNoV) = XNull(RstHlp!LstNo):    txt(LstDateV) = XNull(RstHlp!LstDate)
            txt(16) = XNull(RstHlp!CST):    txt(17) = XNull(RstHlp!CST_Date)
            txt(18) = XNull(RstHlp!Mobile): txt(19) = XNull(RstHlp!Phone)
            txt(20) = XNull(RstHlp!FAx):    txt(21) = XNull(RstHlp!Gram)
            txt(22) = XNull(RstHlp!MailID): txt(23) = XNull(RstHlp!IT_WardNo)
            txt(24) = XNull(RstHlp!IT_AcNo): txt(25) = XNull(RstHlp!PAN_No)
            txt(26) = XNull(RstHlp!Speciality)
            txt(27) = XNull(RstHlp!FADataPath)
        Else
            txt(6) = "":    txt(7) = "":    txt(8) = "":    txt(9) = ""
            txt(10) = "":   txt(11) = "":   txt(12) = "":   txt(13) = ""
            txt(14) = "":   txt(15) = "":   txt(16) = "":   txt(17) = ""
            txt(18) = "":   txt(19) = "":   txt(20) = "":   txt(21) = ""
            txt(22) = "":   txt(23) = "":   txt(24) = "":   txt(25) = ""
            txt(26) = "":   txt(27) = "": txt(LstNoV) = "": txt(LstDateV) = ""
        End If
        Label6.CAPTION = IIf(txt(27) <> "", Pub_DataPath & "\" & txt(27), "")
    ElseIf STab.Tab = 2 Then
        If txt(28) <> "" Then
            txt(28) = RstHlp!AssoComp_Code:  txt(29) = RstHlp!AssoComp_SName
            txt(30) = RstHlp!AssoComp_Name:  txt(31) = XNull(RstHlp!Add1)
            txt(32) = XNull(RstHlp!Add2):    txt(33) = XNull(RstHlp!Add3)
            txt(34) = XNull(RstHlp!City):    txt(35) = XNull(RstHlp!PinCode)
            txt(36) = XNull(RstHlp!LST):     txt(37) = XNull(RstHlp!LST_Date)
            txt(LstNoS) = XNull(RstHlp!LstNo):    txt(LstDateS) = XNull(RstHlp!LstDate)
            txt(38) = XNull(RstHlp!CST):     txt(39) = XNull(RstHlp!CST_Date)
            txt(40) = XNull(RstHlp!Mobile):  txt(41) = XNull(RstHlp!Phone)
            txt(42) = XNull(RstHlp!FAx):     txt(43) = XNull(RstHlp!Gram)
            txt(44) = XNull(RstHlp!MailID):  txt(45) = XNull(RstHlp!IT_WardNo)
            txt(46) = XNull(RstHlp!IT_AcNo): txt(47) = XNull(RstHlp!PAN_No)
            txt(48) = XNull(RstHlp!Speciality)
            txt(49) = XNull(RstHlp!FADataPath)
        Else
            txt(28) = "":   txt(29) = "":   txt(30) = "":   txt(31) = ""
            txt(32) = "":   txt(33) = "":   txt(34) = "":   txt(35) = ""
            txt(36) = "":   txt(37) = "":   txt(38) = "":   txt(39) = ""
            txt(40) = "":   txt(41) = "":   txt(42) = "":   txt(43) = ""
            txt(44) = "":   txt(45) = "":   txt(46) = "":   txt(47) = ""
            txt(48) = "":   txt(49) = "": txt(LstNoS) = "": txt(LstNoS) = ""
        End If
        Label7.CAPTION = IIf(txt(49) <> "", Pub_DataPath & "\" & txt(49), "")
    ElseIf STab.Tab = 3 Then
        If txt(50) <> "" Then
            txt(50) = RstHlp!AssoComp_Code:  txt(51) = RstHlp!AssoComp_SName
            txt(52) = RstHlp!AssoComp_Name:  txt(53) = XNull(RstHlp!Add1)
            txt(54) = XNull(RstHlp!Add2):    txt(55) = XNull(RstHlp!Add3)
            txt(56) = XNull(RstHlp!City):    txt(57) = XNull(RstHlp!PinCode)
            txt(58) = XNull(RstHlp!LST):     txt(59) = XNull(RstHlp!LST_Date)
            txt(LstNoW) = XNull(RstHlp!LstNo):    txt(LstDateW) = XNull(RstHlp!LstDate)
            txt(60) = XNull(RstHlp!CST):     txt(61) = XNull(RstHlp!CST_Date)
            txt(62) = XNull(RstHlp!Mobile):  txt(63) = XNull(RstHlp!Phone)
            txt(64) = XNull(RstHlp!FAx):     txt(65) = XNull(RstHlp!Gram)
            txt(66) = XNull(RstHlp!MailID):  txt(67) = XNull(RstHlp!IT_WardNo)
            txt(68) = XNull(RstHlp!IT_AcNo): txt(69) = XNull(RstHlp!PAN_No)
            txt(70) = XNull(RstHlp!Speciality)
            txt(71) = XNull(RstHlp!FADataPath)
        Else
            txt(50) = "":   txt(51) = "":   txt(52) = "":   txt(53) = ""
            txt(54) = "":   txt(55) = "":   txt(56) = "":   txt(57) = ""
            txt(58) = "":   txt(59) = "":   txt(60) = "":   txt(61) = ""
            txt(62) = "":   txt(63) = "":   txt(64) = "":   txt(65) = ""
            txt(66) = "":   txt(67) = "":   txt(68) = "":   txt(69) = ""
            txt(70) = "":   txt(71) = "": txt(LstNoW) = "": txt(LstNoW) = ""
        End If
        Label8.CAPTION = IIf(txt(71) <> "", Pub_DataPath & "\" & txt(71), "")
    End If
    RstHlp.Requery
End If
End Sub
Public Sub PubDefine()
On Error Resume Next
Set RstFrm = New ADODB.Recordset
Dim RsTemp As ADODB.Recordset
With RstFrm
    .ActiveConnection = GCn
    .LockType = adLockReadOnly
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .Open "select Syctrl.* ,Site.site_desc,site.SiteType from Syctrl Left Join Site on Syctrl.site_code=site.site_code"
End With
'Speed Print Font declaration
mChr14 = Chr(14)
mChr10 = Chr(10)
'mChr15 = Chr(27) + Chr(103)
mChr17 = Chr(15)
mChr18 = Chr(18)
mChr20 = Chr(27) + Chr(77) + Chr(15)
mChr201 = Chr(27) + Chr(80) + Chr(18)
mEmph = Chr(27) + Chr(69)
mEmph1 = Chr(27) + Chr(70)
'Due
mDoub = mEmph  'Chr(27) + Chr(71) due to printer problem
mDoub1 = mEmph1 'Chr(27) + Chr(72)
mUnd = Chr(27) + Chr(45) + Chr(1)
mUnd1 = Chr(27) + Chr(45) + Chr(0)
mEject = Chr(12)
'eof declaration

'Common
PubLoginDate = Txtdate.TEXT
PubAcPostingByAllUser = IIf(RstFrm!AcPostingByAllUser = 1, True, False)
PubChqNoReq = IIf(RstFrm!ChqNoReq = 1, True, False)
PubLineFill = RstFrm!LineFill
PubAmountPrefix = RstFrm!AmountPrefix
PubPageLength = RstFrm!LinePerPage
    
PubSprTaxInvPrefix = RstFrm!SprTaxInvPrefix
PubVehTaxInvPrefix = RstFrm!VehTaxInvPrefix
PubTaxOnFreeLabYn = VNull(RstFrm!TaxOnFreeLabYN)
PubPageLengthHalf = RstFrm!LinePerHalfPage
PubSpeedPrint = RstFrm!SpeedPrint
PubCrLimitCheck = RstFrm!CrLimitCheck
PubForm31Caption = RstFrm!Form31Caption
PubServiceTaxNo = RstFrm!SrvTaxNo
PubSiteCode = RstFrm!Site_Code
PubSiteType = RstFrm!SiteType
PubSiteName = RstFrm!Site_Desc
PubRoundOffPosition = RstFrm!RoundOffPosition
PubRoundOffType = RstFrm!RoundOffType
PubLockFinancialYear = VNull(RstFrm!LockFinancialYear)
BiLanguage = False 'True
BiLanguageName = "Hindi"
BiLanguageFont = "Hitarth Hin Jalak"
PubEditLock = VNull(RstFrm!EditLock)
   
   
PubVehGodown = IIf(IsNull(RstFrm!VehGodown), "", RstFrm!VehGodown)
PubVehRateIncTaxYn = RstFrm!VehRateIncTax
PubSiteCodeDisplay = "('" & Trim(RstFrm!Site_Code) & "')"
If PubSiteWiseDisplayYn = 1 Then
        If PubSiteType = "H" Then
        PubFaSiteType = 1               '0-General,1-FromSite ForSite,2-2 Char Site
        Else
        PubFaSiteType = 0               '0-General,1-FromSite ForSite,2-2 Char Site
        End If
    PubFaSiteType = 1
Else
   PubFaSiteType = 0
End If
PubSiteCodeWiseMasterRst = False 'If You Want filter in Master Recordset
PubSiteCodeWiseHelp = False     'If You Want filter on Help
PubSiteCodeWidth = 2            'If in Ledger Site code as S then 1 else if SS then 2
                                'It will works only when PubFaSiteType=2
   
PubRSO_Code = XNull(RstFrm!RSO_Code)           'withdraw
PubOwnFinCode = XNull(RstFrm!OwnFinCode)       'withdraw
'    If Txt(VehPath).Text <> "" Then
'                Set GCnFaV = New ADODB.Connection

'Store Specific
pubTOT_On = RstFrm!TOT_On   '0-Sub Total (B) TB+TP, 1-Taxable+Taxpaid Total
If pubTOT_On = 0 Then
    pubTOTCaption = "TOT on Sub Total (B)"
Else
    pubTOTCaption = "TOT on SubTot(BefTax)"
End If
If RSOJPR = True Then
    PubSprIssOnNegStk = 0
Else
    PubSprIssOnNegStk = RstFrm!SprIssOnNegStk
End If

PubSprCounterGodown = RstFrm!SprCounterGodown
PubGatePassOnSprInv = RstFrm!GatePassOnSprInv
'** From Stores FAData.AcContorls table
If txt(SprPath).TEXT <> "" Then
    Set GRs = New Recordset
    With GRs
        .ActiveConnection = GCnFaS
        .LockType = adLockReadOnly
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Open "select * from AcControls"
    End With
    If GRs.RecordCount > 0 Then
        PubSprCashAc = XNull(GRs!SprCash_Ac) 'CPSprAc
        PubCreditCardAc = XNull(GRs!CreditCardAc)
        PubChqClrAc = XNull(GRs!ChqClrAc)
    Else
        MsgBox "Define Spare A/c Controls through Utility Menu", vbCritical, "Control Setting Required"
    End If
End If
 '** EOF Stores FAData
 
 'Workshop specific
 PubIPO_Separate = RstFrm!IPO_Separate
 PubSepLabourInv = RstFrm!SepLabourInv       'Withdraw
 PubLabRate_Chargable = RstFrm!LabRate_Chargable
 PubLabRate_Warranty = RstFrm!LabRate_Warranty
 PubSprWorksGodown = RstFrm!SprWorksGodown
 PubSrvGatePass = RstFrm!SrvGatePass         'Withdraw
 PubServiceZone = RstFrm!ServiceZone         'Withdraw
 pubLocalTaxFormSpr = RstFrm!LocalTaxFormSpr
 pubGovtTaxFormSpr = IIf(IsNull(RstFrm!GovtTaxFormSpr), "", RstFrm!GovtTaxFormSpr)
 '** From Workshop FAData.AcContorls table
 If txt(WrkPath).TEXT <> "" Then
     Set GRs = New Recordset
     With GRs
         .ActiveConnection = GCnFaS
         .LockType = adLockReadOnly
         .CursorType = adOpenStatic
         .CursorLocation = adUseClient
         .Open "select SrvCash_Ac,SrvLabour_Ac from AcControls"
     End With
     If GRs.RecordCount > 0 Then
         PubSrvLabAc = IIf(IsNull(GRs!SrvLabour_Ac), "", GRs!SrvLabour_Ac)
         PubSrvCashAc = IIf(IsNull(GRs!SrvCash_Ac), "", GRs!SrvCash_Ac)
     Else
         MsgBox "Define Workshop A/c Controls through Utility Menu", vbCritical, "Control Setting Required"
     End If
     Set GRs = Nothing
 End If
 '** EOF Stores FAData
' viaksh
 Dim I As Double
 Dim mFldUpd As Boolean, mFldUpd1 As Boolean
 mFldUpd = True
 mFldUpd1 = True
 I = 0
 For I = 0 To RstFrm.Fields.Count - 1
     If UCase(RstFrm.Fields(I).Name) = UCase("DiscOnLube") Then
          mFldUpd = False
     End If
     If UCase(RstFrm.Fields(I).Name) = UCase("TotOnLube") Then
          mFldUpd1 = False
     End If
 Next
If mFldUpd Or mFldUpd1 Then
      MDIForm1.AddNewFieldFAData
End If
' end of viaksh
'Store & Works
PubDiscOnLube = IIf(IsNull(RstFrm!DiscOnLube), 0, RstFrm!DiscOnLube)
PubTOTOnLube = IIf(IsNull(RstFrm!TotOnLube), 0, RstFrm!TotOnLube)
PubReSaleTaxPer = RstFrm!ReSaleTax_Per
PubTaxDetOnSprInv = RstFrm!TaxDetOnSprInv
PubRestrict_Godown = RstFrm!Restrict_Godown
PubGenSurChrgOnSpr = RstFrm!GenSurChrgOnSpr
PubTOT_YN = RstFrm!TOT_YN
PubTOT_Rate = RstFrm!TOT_Rate
PubTBR_to_TPR = RstFrm!TBR_to_TPR
PubPartGrade_Lub = RstFrm!PartGrade_Lub
PubPartGrade_Consum = RstFrm!PartGrade_Consum
PubPartGrade_Tool = RstFrm!PartGrade_Tool
PubMergeGenSur_TB_Sale = RstFrm!MergeGenSur_TB_Sale

'PubUParam = "AEDP"
PubSpeedPrint = IIf(RstFrm!SpeedPrint = 0, False, True)
PubVATYN = VNull(RstFrm!VAT_YN)
PubSDTYN = VNull(RstFrm!SDT_YN)
If PubBackEnd = "A" Then
    PubSatYn = VNull(RstFrm!SAT_YN)
Else
    PubSatYn = IIf(VNull(RstFrm!SAT_YN), 1, 0)
End If

If PubBackEnd = "A" Then
    PubSiteWiseDisplayYn = VNull(RstFrm!SiteWiseDisplaY_N)
Else
    PubSiteWiseDisplayYn = IIf(VNull(RstFrm!SiteWiseDisplaY_N), 1, 0)
    If PubSiteWiseDisplayYn = 1 Then PubFaSiteType = 1
End If


If XNull(GCn.Execute("Select TOTCaption from Syctrl").Fields(0).Value) = "" Then
    pubTOTCaption = "T O T"
Else
    pubTOTCaption = XNull(GCn.Execute("Select TOTCaption from Syctrl").Fields(0).Value)
End If
PubSiebelActiveYn = VNull(RstFrm!SiebelActiveYN)


Dim mVersion As Long
Dim mCurrentVersion As Long
mCurrentVersion = Val(App.Major & App.Minor & App.Revision)
Set RsTemp = G_CompCn.Execute("Select " & cVal("Max(Exe)") & " From Company")
mVersion = RsTemp(0)
If mVersion > mCurrentVersion Then
    MsgBox "You Can't Run Old Exe"
    End
ElseIf mVersion < mCurrentVersion Then
    G_CompCn.Execute "Update Company Set Exe = " & mCurrentVersion & " "
End If


Set RstFrm = Nothing
End Sub
Private Sub SelectFAData()

'        PubFirmCode = RsDiv!W_SecCompCode
'        PubComp_Name = RsDiv!W_SecName
'        PubComp_Add = XNull(RsDiv!W_SecAdd1)
'        PubComp_Add2 = XNull(RsDiv!W_SecAdd2)
'        PubComp_City = XNull(RsDiv!W_SecCity) & IIf(RsDiv!W_SecPinCode = "", "", "-") & XNull(RsDiv!W_SecPinCode)
        
    PubSecName = "Account"
    If ListView.SelectedItem.TEXT = "Vehicle" Then
        PubFirmCode = RsDiv!v_SecCompCode
        PubComp_Name = RsDiv!v_SecName
        PubComp_Add = XNull(RsDiv!v_SecAdd1)
        PubComp_Add2 = XNull(RsDiv!v_SecAdd2)
        PubComp_City = XNull(RsDiv!v_SecCity)
        PubComp_Contact = "PHONE : " & XNull(RsDiv!S_SecPhone) & " Fax   : " & XNull(RsDiv!S_SecFax)
        Set G_FaCn = New ADODB.Connection
        With G_FaCn
            .CursorLocation = adUseClient
            If PubBackEnd = "A" Then
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .ConnectionString = "Data Source=" & PubVFADataPath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
            Else
                If PubDbUser <> "" Then
                    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                Else
            
                    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                End If
            End If
            
            .Open
        End With
        PubFADataPath = PubVFADataPath
        'Licence checking
        DatamanKeyNo1.ModuleCode = "AMW"
        DatamanKeyNo1.FirmName = RsDiv!v_SecName
        DatamanKeyNo1.CityName = RsDiv!v_SecCity
        If Not StrCmp(left(PubComp_Name, 6), "PRAYAG") Then
            If DatamanKeyNo1.Validate(IIf(IsNull(RsDiv!ProductSerial), "", RsDiv!ProductSerial)) = False Then
                MsgBox "Dataman Demo Product ID : " & DatamanKeyNo1.ReturnID, vbInformation
                If G_FaCn.Execute("select count(*) from Ledger ").Fields(0).Value > 250 Then End
            End If
        End If
        'eof licence
    ElseIf ListView.SelectedItem.TEXT = "Spare" Then
        PubFirmCode = RsDiv!s_SecCompCode
        PubComp_Name = RsDiv!s_SecName
        PubComp_Add = XNull(RsDiv!s_SecAdd1)
        PubComp_Add2 = XNull(RsDiv!s_SecAdd2)
        PubComp_City = XNull(RsDiv!s_SecCity)
        Set G_FaCn = New ADODB.Connection
        With G_FaCn
            .CursorLocation = adUseClient
            If PubBackEnd = "A" Then
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .ConnectionString = "Data Source=" & PubSFADataPath & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
            Else
                If PubDbUser <> "" Then
                    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                Else
            
                    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                End If
            End If
            .Open
        End With
        PubFADataPath = PubSFADataPath
        'Licence checking
        DatamanKeyNo1.ModuleCode = "AMW"
        DatamanKeyNo1.FirmName = RsDiv!s_SecName
        DatamanKeyNo1.CityName = RsDiv!s_SecCity
        If Not StrCmp(left(PubComp_Name, 6), "PRAYAG") Then
            If DatamanKeyNo1.Validate(IIf(IsNull(RsDiv!ProductSerial), "", RsDiv!ProductSerial)) = False Then
                MsgBox "Dataman Demo Product ID : " & DatamanKeyNo1.ReturnID, vbInformation
                If G_FaCn.Execute("select count(*) from Ledger ").Fields(0).Value > 250 Then End
            End If
        End If
        'eof licence
    ElseIf ListView.SelectedItem.TEXT = "Works" Then
        PubFirmCode = RsDiv!w_SecCompCode
        PubComp_Name = RsDiv!w_SecName
        PubComp_Add = XNull(RsDiv!w_SecAdd1)
        PubComp_Add2 = XNull(RsDiv!w_SecAdd2)
        PubComp_City = XNull(RsDiv!w_SecCity)
        Set G_FaCn = New ADODB.Connection
        With G_FaCn
            .CursorLocation = adUseClient
            If PubBackEnd = "A" Then
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .ConnectionString = "Data Source=" & PubWFADataPath & ";Persist Security Info=False"
            Else
                If PubDbUser <> "" Then
                    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & PubDbUser & ";Password=" & PubDbPass & ";Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                Else
                    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & PubCenDataPath & ";Data Source=" & PubServerName
                End If
            End If
            .Open
        End With
        PubFADataPath = PubWFADataPath
        'Licence checking
        DatamanKeyNo1.ModuleCode = "AMW"
        DatamanKeyNo1.FirmName = RsDiv!w_SecName
        DatamanKeyNo1.CityName = RsDiv!w_SecCity
        If Not StrCmp(left(PubComp_Name, 6), "PRAYAG") Then
            If DatamanKeyNo1.Validate(IIf(IsNull(RsDiv!ProductSerial), "", RsDiv!ProductSerial)) = False Then
    '            MsgBox "Dataman Demo Product ID : " & DatamanKeyNo1.ReturnID, vbInformation
    '            If G_FaCn.Execute("select count(*) from Ledger ").Fields(0).Value > 250 Then End
            End If
        End If
        'eof licence
    End If
    FrmList.Visible = False
'    If pubUName <> "SA" Then
'        MDIForm1.MnuVeh.Visible = False
'        MDIForm1.MnuSpr.Visible = False
'        MDIForm1.MnuWorks.Visible = False
'    End If







    PubLoginDate = Txtdate.TEXT
    PubDefine
    MDIForm1.CAPTION = PubPackage & " - [" & RsDiv!Div_SName & "] " & PubComp_Name
    Unload frmDivision
    MDIForm1.Show
    
    
End Sub


Sub LST_Fields_Add()
On Error Resume Next

        GCn.Execute "Alter Table Division Add LstNo nVarChar(30) Default ''"
        GCn.Execute "Alter Table Division Add LstDate SmallDateTime"
        GCn.Execute "Alter Table Division Add LstNoV nVarChar(30) Default ''"
        GCn.Execute "Alter Table Division Add LstDateV SmallDateTime"
        GCn.Execute "Alter Table Division Add LstNoS nVarChar(30) Default ''"
        GCn.Execute "Alter Table Division Add LstDateS SmallDateTime"
        GCn.Execute "Alter Table Division Add LstNoW nVarChar(30) Default ''"
        GCn.Execute "Alter Table Division Add LstDateW SmallDateTime"
        GCn.Execute "Alter Table AssociatedFirms Add LstNo nVarChar(30) Default ''"
        GCn.Execute "Alter Table AssociatedFirms Add LstDate SmallDateTime"

End Sub

Sub UpdateTableStructureSql()
Dim RsUser As ADODB.Recordset
Dim TmpRst As ADODB.Recordset
On Error Resume Next



    CreateNewTableSQL

    'GCn.Execute "Update Division Set VBilRptName='VehBill-LMP'"
    
    GCn.Execute "Alter Table  LedgerM Add DmsSubCode nVarChar(15)   Default ''"
    GCn.Execute "Alter Table  Veh_Order Add DoNo nVarChar(20)   Default ''"
    GCn.Execute "Alter Table Veh_Order Alter Column DoNo DateTime "
    GCn.Execute "Alter Table  Veh_Order Add DoReciveDate DateTime   Default ''"
    GCn.Execute "Alter Table  Veh_Order Add DoIssueDate DateTime   Default ''"
    GCn.Execute "Alter Table  Veh_Transfer Add OrdDocID nVarChar(21)   Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add DefaultOilPartNo nVarChar(21)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add DefaultPartNo nVarChar(21)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add DefaultLabourHead nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add DefaultMechanic nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add DefaultSupervisor nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add VehicleCentralPurchaseTaxForm nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add VehicleLocalPurchaseTaxForm nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add VehiclePurchaseDiscountItem nVarChar(22)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add VehiclePurchaseTransportItem nVarChar(22)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add VehicleTaxOnDeliveryCharges nVarChar(1)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add SpareCentralPurchaseTaxForm nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add SpareLocalPurchaseTaxForm nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add SpareCentralSaleTaxForm nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add SpareLocalSaleTaxForm nVarChar(4)     Default ''"
    
    GCn.Execute "Alter Table  Godown Add DmsCode         nVarChar(30)     Default ''"
    GCn.Execute "Alter Table  Model_Cat Add DmsCode         nVarChar(50)     Default ''"
    GCn.Execute "Alter Table  Syctrl    Add OpenPartyInAllCompany    Numeric(1) Default 0"
    GCn.Execute "Alter Table  DmsEnviro Add VatInputAc         nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro Add Vat4InputAc        nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  Syctrl    Add LockFinancialYear    Numeric(1) Default 0"
    GCn.Execute "Alter Table  Job_Card  Add Created_AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Job_Card  Add Created_AddDate DateTime "
    GCn.Execute "Alter Table  Job_Card  Add Created_ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Job_Card  Add Created_ModifyDate DateTime "
    
    GCn.Execute "Alter Table  Job_Card  Add Closed_AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Job_Card  Add Closed_AddDate DateTime "
    GCn.Execute "Alter Table  Job_Card  Add Closed_ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Job_Card  Add Closed_ModifyDate DateTime "
    
    
    
    GCn.Execute "Alter Table  Job_Lab  Add AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Job_Lab  Add AddDate DateTime "
    GCn.Execute "Alter Table  Job_Lab  Add ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Job_Lab  Add ModifyDate DateTime "
    
    GCn.Execute "Alter Table  Ledger  Add AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Ledger  Add AddDate DateTime "
    GCn.Execute "Alter Table  Ledger  Add ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Ledger  Add ModifyDate DateTime "
    
    GCn.Execute "Alter Table  LedgerM  Add AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  LedgerM  Add AddDate DateTime "
    GCn.Execute "Alter Table  LedgerM  Add ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  LedgerM  Add ModifyDate DateTime "
    
    GCn.Execute "Alter Table  SP_Purch  Add AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  SP_Purch  Add AddDate DateTime "
    GCn.Execute "Alter Table  SP_Purch  Add ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  SP_Purch  Add ModifyDate DateTime "
    
    GCn.Execute "Alter Table  SP_Sale  Add AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  SP_Sale  Add AddDate DateTime "
    GCn.Execute "Alter Table  SP_Sale  Add ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  SP_Sale  Add ModifyDate DateTime "
    
    GCn.Execute "Alter Table  Veh_Order  Add Book_AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order  Add Book_AddDate DateTime "
    GCn.Execute "Alter Table  Veh_Order  Add Book_ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order  Add Book_ModifyDate DateTime "
    
    GCn.Execute "Alter Table  Veh_Order  Add DelCh_AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order  Add DelCh_AddDate DateTime "
    GCn.Execute "Alter Table  Veh_Order  Add DelCh_ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order  Add DelCh_ModifyDate DateTime "
    
    GCn.Execute "Alter Table  Veh_Order  Add Inv_AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order  Add Inv_AddDate DateTime "
    GCn.Execute "Alter Table  Veh_Order  Add Inv_ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order  Add Inv_ModifyDate DateTime "
    
    GCn.Execute "Alter Table  Veh_Order1  Add Book_AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order1  Add Book_AddDate DateTime "
    GCn.Execute "Alter Table  Veh_Order1  Add Book_ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order1  Add Book_ModifyDate DateTime "
    
    GCn.Execute "Alter Table  Veh_Order1  Add DelCh_AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order1  Add DelCh_AddDate DateTime "
    GCn.Execute "Alter Table  Veh_Order1  Add DelCh_ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order1  Add DelCh_ModifyDate DateTime "
    
    GCn.Execute "Alter Table  Veh_Order1  Add Inv_AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order1  Add Inv_AddDate DateTime "
    GCn.Execute "Alter Table  Veh_Order1  Add Inv_ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Order1  Add Inv_ModifyDate DateTime "
    
    GCn.Execute "Alter Table  Veh_Purch1  Add AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Purch1  Add AddDate DateTime "
    GCn.Execute "Alter Table  Veh_Purch1  Add ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Veh_Purch1  Add ModifyDate DateTime "
    
    GCn.Execute "Alter Table  Rect  Add AddBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Rect  Add AddDate DateTime "
    GCn.Execute "Alter Table  Rect  Add ModifyBy NVarChar(10) Default '' "
    GCn.Execute "Alter Table  Rect  Add ModifyDate DateTime "
    GCn.Execute "Alter Table  Ledger Add Chq_Favour NVarChar(255) Default '' "
    GCn.Execute "Alter Table  Ledger Add Chq_AcPayee NVarChar(255) Default '' "
    
    
    GCn.Execute "Alter Table  Sp_Stock  Add GatePassNo Numeric(18,0) Default 0 "
    GCn.Execute "Alter Table  Sp_Purch  Add GatePassNo Numeric(18,0) Default 0 "
    GCn.Execute "Alter Table  Estimate1  Add SatPer Float Default 0 "
    GCn.Execute "Alter Table  Estimate1  Add SatAmt Float Default 0 "
    GCn.Execute "Alter Table  Estimate   Add SatAmt Float Default 0 "
    GCn.Execute "Alter Table  TaxForms   Add AddTaxPer VarChar(8) Default '' "
    GCn.Execute "Alter Table  TaxFormsAc   Add AddTaxAc  VarChar(8) Default '' "
    GCn.Execute "Alter Table  Part_Grade   Add AddTaxPer Float Default 0 "
    GCn.Execute "Alter Table  Syctrl    Add Sat_Yn Bit   Default 0 "
    GCn.Execute "Alter Table  Sp_Sale   Add Sat_Yn Bit Default 0 "
    GCn.Execute "Alter Table  Sp_Sale   Add SatAmt Float Default 0 "
    GCn.Execute "Alter Table  Sp_Purch  Add Sat_Yn Bit Default 0 "
    GCn.Execute "Alter Table  Sp_Purch  Add SatAmt Float Default 0 "
    GCn.Execute "Alter Table  Sp_Stock  Add Sat_Yn Bit Default 0 "
    GCn.Execute "Alter Table  Sp_Stock  Add SatPer Float Default 0 "
    GCn.Execute "Alter Table  Sp_Stock  Add SatAmt Float Default 0 "
    GCn.Execute "Alter Table  Veh_Order  Add Sat_Yn Bit Default 0 "
    GCn.Execute "Alter Table  Veh_Order  Add SatPer Float Default 0 "
    GCn.Execute "Alter Table  Veh_Order  Add SatAmt Float Default 0 "
    GCn.Execute "Alter Table  Veh_Order1  Add Sat_Yn Bit Default 0 "
    GCn.Execute "Alter Table  Veh_Order1  Add SatPer Float Default 0 "
    GCn.Execute "Alter Table  Veh_Order1  Add SatAmt Float Default 0 "
    
    GCn.Execute "Alter Table  Veh_Purch1  Add Sat_Yn Bit Default 0 "
    GCn.Execute "Alter Table  Veh_Purch1 Add SatPer Float Default 0 "
    GCn.Execute "Alter Table  Veh_Purch1 Add SatAmt Float Default 0 "
        
    
    GCn.Execute "Alter Table  Estimate1 Add Item_Value Numeric(18,2) Default 0 "
    GCn.Execute "Alter Table  Syctrl Add ServiceTaxPer_Saperate Numeric(18,2) Default 0 "
    GCn.Execute "Alter Table  Syctrl Add HECessPer Numeric(18,2) Default 0 "
    GCn.Execute "Alter Table  Job_Card Add ServiceTaxPer_Saperate Numeric(18,2) Default 0 "
    GCn.Execute "Alter Table  Job_Card Add HECessPer Numeric(18,2) Default 0 "
    GCn.Execute "Alter Table  Job_Card Add ServiceTaxAmt_Saperate Numeric(18,2) Default 0 "
    GCn.Execute "Alter Table  Job_Card Add HECessAmt Numeric(18,2) Default 0 "
    
    
    GCn.Execute "Alter Table  LedgerM Add DmsRefNo nVarChar(40) Default '' "
    GCn.Execute "Alter Table  LedgerM Add DmsRefNo nVarChar(40) Default '' "
    GCn.Execute "Alter Table  Veh_Stock Add  BodyBuilder_BodyType nVarChar(5) Default '' "
    GCn.Execute "Alter Table  Veh_Stock Add  BodyBuilder nVarChar(5) Default ''"
    GCn.Execute "Alter Table  Veh_Stock Add  BodyBuilder_Remark nVarChar(50) Default ''"
    GCn.Execute "Alter Table  Veh_Stock Add  BodyBuilder_IssDate smallDateTime"
    GCn.Execute "Alter Table  Veh_Stock Add  BodyBuilder_RecDate smallDateTime"
    
    GCn.Execute "Alter Table  Exp_Emp1 Add  SubCode nVarChar(8) Default ''"
    GCn.Execute "Alter Table  Ledger Add  EmpDetailYn nVarChar(1) Default 'N'"
    GCn.Execute "Alter Table  SubGroup Add  EmpDetailYn nVarChar(1) Default 'N'"
    GCn.Execute "Alter Table  SubGroupAlias Add  EmpDetailYn nVarChar(1) Default 'N'"
    
    
    GCn.Execute "Alter Table  Syctrl Add  MakeDataBlank nVarChar(20) Default ''"
    GCn.Execute "Alter Table  Site Add  Address1 nVarChar(40) Default ''"
    GCn.Execute "Alter Table  Site Add  Address2 nVarChar(40) Default ''"
    GCn.Execute "Alter Table  Site Add  Address3 nVarChar(40) Default ''"
    GCn.Execute "Alter Table  Site Add  City nVarChar(40) Default ''"
    GCn.Execute "Alter Table  Site Add  PinCode nVarChar(10) Default ''"
    GCn.Execute "Alter Table  Site Add  Phone nVarChar(30) Default ''"
    GCn.Execute "Alter Table  Site Add  Mobile nVarChar(25) Default ''"
    GCn.Execute "Alter Table  Site Add  LstNo nVarChar(30) Default ''"
    GCn.Execute "Alter Table  Site Add  LstDate SmallDateTime"
    GCn.Execute "Alter Table  Site Add  CstNo nVarChar(30) Default ''"
    GCn.Execute "Alter Table  Site Add  CstDate SmallDateTime"
    GCn.Execute "Alter Table Sp_Stock Alter Column Rate2 Numeric(18,3)"
    GCn.Execute "Alter Table Sp_Stock Alter Column Mrp_Rate2 Numeric(18,3)"
    GCn.Execute "Alter table Model ALTER COLUMN Chas_Type nVarChar(9)"
    GCn.Execute "Alter table Veh_Order ALTER COLUMN Chas_Type nVarChar(9)"
    GCn.Execute "Alter table Veh_Order1 ALTER COLUMN Chas_Type nVarChar(9)"
    GCn.Execute "Alter table Model ALTER COLUMN Model_Type nVarChar(3)"
    
    GCn.Execute "Alter Table  Part Add  Dep_Item nVarChar(5) Default ''"
    GCn.Execute "Alter Table  Labour Add  Dep_Item nVarChar(5) Default ''"
   
    
    
    G_CompCn.Execute "Alter Table  User_Module Add  Menu_Name nVarChar(25)"
    G_CompCn.Execute "Alter Table Company Add Exe Numeric "
    GCn.Execute "Alter Table  Veh_Purch1  Add Sat_Yn Bit Default 0 "
    
    GCn.Execute "Alter Table  Part              Add  NDP                    Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Estimate1         Add  TaxPer                 Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Estimate1         Add  TaxAmt                 Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Chas_Mth          Add  Code                   nVarChar(2)     Default ''"
    GCn.Execute "Alter Table  AcControls        Add  SubventionClaimAc      nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  AcControls        Add  SubventionAc           nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  AcControls        Add  IndirectExpAc          nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  AcControls        Add  OctraiAc               nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  AcControls        Add  RegnFeeAc              nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  AcControls        Add  InsuranceFeeAc         nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  AcControls        Add  CreditCardAc           nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  AcControls        Add  ChqClrAc               nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  Emp_Mast          Add Supervisor              nVarChar(5)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro         Add VehCstPurchaseAc        nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  DmsEnviro         Add SprSaleVat4Ac        nVarChar(8)     Default ''"
    GCn.Execute "Alter Table  Veh_Order         Add SubventionScheme        nVarChar(20)    Default ''"
    GCn.Execute "Alter Table  Veh_Order         Add DeliveryFrom            nVarChar(10)    Default ''"
    GCn.Execute "Alter Table  Veh_Order         Add RTOFee                  Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Order         Add Insurance               Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Order         Add DealerContribution      Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Order         Add TataContribution        Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Order         Add HandlingCharges         Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Order         Add Subvention              Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Order1        Add SubventionScheme        nVarChar(20)    Default ''"
    GCn.Execute "Alter Table  Veh_Order1        Add DeliveryFrom            nVarChar(10)    Default ''"
    GCn.Execute "Alter Table  Veh_Order1        Add RTOFee                  Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Order1        Add Insurance               Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Order1        Add DealerContribution      Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Order1        Add TataContribution        Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Order1        Add HandlingCharges         Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Order1        Add Subvention              Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Job_Lab           Add ActualHrs               Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Purch1        Add SubventionCredit        Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Rate          Add Reg_FeeCom              Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Rate          Add HandlingCharges         Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Rate          Add GenExGodRate            Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Veh_Rate          Add GovtExGodRate           Numeric(18,3)   Default 0"
    GCn.Execute "Alter Table  Model_Cat         Add OldCode                 nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  Model_Grp         Add OldCode                 nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  Godown            Add OldCode                 nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  FinBank           Add OldCode                 nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  FinBank           Add xName                   nVarChar(40)    Default ''"
    GCn.Execute "Alter Table  Area              Add OldCode                 nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  ContractFinance   Add OldCode                 nVarChar(4)     Default ''"
    GCn.Execute "Alter Table  Job_Card          Add CreditCardNo            nVarChar(20)    Default ''"
    GCn.Execute "Alter Table  Rect              Add CreditCardNo            nVarChar(20)    Default ''"
    GCn.Execute "Alter Table  Syctrl            Add EditLock                Numeric(9,2) Default 0"
    GCn.Execute "Alter Table  Syctrl            Add CheckNegetiveStockSiteWise    Numeric(1) Default 0"
    GCn.Execute "Alter Table  Syctrl            Add LabDiscAfterTaxYn       Numeric(1)      Default 0"
    GCn.Execute "Alter Table  Syctrl            Add RtoInsInBill            Numeric(1)      Default 0"
    GCn.Execute "Alter Table  Syctrl            Add TaxOnFreeLabYN          Numeric(1)      Default 0"
    GCn.Execute "Alter Table  Syctrl            Add PostRegnFeeYn           Numeric(1) Default 0"
    GCn.Execute "Alter Table  Syctrl            Add PostInsuranceFeeYn      Numeric(1) Default 0"
    GCn.Execute "Alter Table  Syctrl            Add PostOctraiSaperatelyYN  Numeric(1) Default 0"
    GCn.Execute "Alter Table  Syctrl            Add SprTaxInvPrefix         nVarChar(5)"
    GCn.Execute "Alter Table  Syctrl            Add VehTaxInvPrefix         nVarChar(5)"
    GCn.Execute "Alter Table  Service_Type      Add RateEditableYN          Numeric(1) Default 0"
    
    GCn.Execute "Alter Table  SubGroup          Add ChequeReportName          nVarChar(50) "
    GCn.Execute "Alter Table  SubGroup          Add MulticityChequeReportName          nVarChar(50) "
    GCn.Execute "Alter Table  SubGroupAlias          Add ChequeReportName          nVarChar(50) "
    GCn.Execute "Alter Table  SubGroupAlias          Add MulticityChequeReportName          nVarChar(50) "
    GCn.Execute "Alter Table  HisCard          Add InsuranceCompany          nVarChar(5)"
    GCn.Execute "Alter Table  HisCard          Add InsuranceExpiry          SmallDateTime"
    GCn.Execute "Alter Table  HisCard          Add InsurancePolicyNo        nVarChar(20)"
    GCn.Execute "Alter Table  HisCard          ALTER COLUMN Supplier_BillNo   nVARCHAR(25)"

  
    
    '################   MODEL FIELD WIDTH INCREASE   ################
    
    GCn.Execute ("ALTER TABLE Veh_OrderM ALTER COLUMN Model VARCHAR(30)")
    
    GCn.Execute ("ALTER TABLE Veh_Rate Drop constraint DF_Veh_Rate_Model")
    GCn.Execute ("ALTER TABLE Veh_Rate Drop constraint PK_Veh_Rate")
    GCn.Execute ("ALTER TABLE Veh_Rate ALTER COLUMN MODEL VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Veh_Rate Add constraint DF_Veh_Rate_Model Default '' for Model")
    GCn.Execute ("ALTER TABLE Veh_Rate Add constraint PK_Veh_Rate Primary Key (Model, Effective_Date, Rso_Work, Taxable_Yn)")
    
    
    GCn.Execute ("ALTER TABLE CustInfo Drop constraint DF_CustInfo_Model")
    GCn.Execute ("ALTER TABLE CustInfo ALTER COLUMN Model VARCHAR(30)")
    GCn.Execute ("ALTER TABLE CustInfo Add constraint DF_CustInfo_Model Default '' for Model")
    GCn.Execute ("ALTER TABLE Veh_Stock ALTER COLUMN MODEL VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Subvention ALTER COLUMN Model VARCHAR(30)")
    
    
    
    GCn.Execute ("ALTER TABLE Veh_Forecast Drop constraint DF_Veh_Forecast_Model")
    GCn.Execute ("ALTER TABLE Veh_Forecast Drop constraint PK_Veh_Forecast")
    GCn.Execute ("ALTER TABLE Veh_Forecast ALTER COLUMN MODEL VARCHAR(30) Not Null")
    GCn.Execute ("ALTER TABLE Veh_Forecast Add constraint DF_Veh_Forecast_Model Default '' for Model")
    GCn.Execute ("ALTER TABLE Veh_Forecast Add constraint PK_Veh_Forecast Primary Key (For_Year,Model)")
    
    
    GCn.Execute ("ALTER TABLE Model_Grp Drop constraint DF_Model_Grp_ModelGrp_Name")
    GCn.Execute ("ALTER TABLE Model_Grp ALTER COLUMN ModelGrp_Name VARCHAR(40)")
    GCn.Execute ("ALTER TABLE Model_Grp Add constraint DF_Model_Grp_ModelGrp_Name Default '' for ModelGrp_Name")
    
    
    GCn.Execute ("ALTER TABLE Veh_Margin Drop constraint DF_Veh_Margin_Model")
    GCn.Execute ("ALTER TABLE VEH_Margin ALTER COLUMN Model VARCHAR(30)")
    GCn.Execute ("ALTER TABLE VEH_Margin Add constraint DF_VEH_Margin_Model Default '' for Model")
    
    
    GCn.Execute ("ALTER TABLE Job_Inspection Drop constraint DF_Job_Inspection_Model")
    GCn.Execute ("ALTER TABLE Job_Inspection ALTER COLUMN Model VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Job_Inspection Add constraint DF_Job_Inspection_Model Default '' for Model")
    
    
    GCn.Execute ("ALTER TABLE Labour_CheckList Drop constraint DF_Labour_CheckList_Model")
    GCn.Execute ("ALTER TABLE Labour_CheckList Drop constraint PK_Labour_CheckList")
    GCn.Execute ("ALTER TABLE Labour_CheckList ALTER COLUMN Model VARCHAR(30) Not Null")
    GCn.Execute ("ALTER TABLE Labour_CheckList Add constraint DF_Labour_CheckList_Model Default '' for Model")
    GCn.Execute ("ALTER TABLE Labour_CheckList Add constraint PK_Labour_CheckList Primary Key (Model, Serv_Type,Lab_Code)")
    
    
    GCn.Execute ("ALTER TABLE Veh_Order1 ALTER COLUMN MODEL VARCHAR(30)")
    
    
    
    GCn.Execute ("ALTER TABLE Veh_SubGroupQuot Drop constraint DF_Veh_SubGroupQuot_Model")
    GCn.Execute ("ALTER TABLE Veh_SubGroupQuot Drop constraint PK_Veh_SubGroupQuot")
    GCn.Execute ("ALTER TABLE Veh_SubGroupQuot ALTER COLUMN MODEL VARCHAR(30) Not Null")
    GCn.Execute ("ALTER TABLE Veh_SubGroupQuot Add constraint DF_Veh_SubGroupQuot_Model Default '' for Model")
    'GCn.Execute ("ALTER TABLE Veh_SubGroupQuot Add constraint PK_Veh_SubGroupQuot Primary Key (Model, Serv_Type,Lab_Code)")
    
    GCn.Execute ("ALTER TABLE ModelCheckList Drop constraint DF_ModelCheckList_Model")
    GCn.Execute ("ALTER TABLE ModelCheckList Drop constraint PK_ModelCheckList")
    GCn.Execute ("ALTER TABLE ModelCheckList ALTER COLUMN MODEL VARCHAR(30) Not Null")
    GCn.Execute ("ALTER TABLE ModelCheckList Add constraint DF_ModelCheckList_Model Default '' for Model")
    GCn.Execute ("ALTER TABLE ModelCheckList Add constraint PK_ModelCheckList Primary Key (Model, Item_Code)")
    
    
    GCn.Execute ("ALTER TABLE Veh_Target Drop constraint DF_Veh_Target_Model")
    GCn.Execute ("ALTER TABLE Veh_Target Drop constraint PK_Veh_Target")
    GCn.Execute ("ALTER TABLE Veh_Target ALTER COLUMN MODEL VARCHAR(30) Not Null")
    GCn.Execute ("ALTER TABLE Veh_Target Add constraint DF_Veh_Target_Model Default '' for Model")
    GCn.Execute ("ALTER TABLE Veh_Target Add constraint PK_Veh_Target Primary Key (Model, Rep_Code)")
    
    
    
    GCn.Execute ("ALTER TABLE Veh_Transfer Drop constraint DF_Veh_Transfer_Model")
    GCn.Execute ("ALTER TABLE Veh_Transfer ALTER COLUMN MODEL VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Veh_Transfer Add constraint DF_Veh_Transfer_Model Default '' for Model")
    
    GCn.Execute ("ALTER TABLE RTOModel Drop constraint DF_RTOModel_Model")
    GCn.Execute ("ALTER TABLE RTOModel ALTER COLUMN MODEL VARCHAR(30)")
    GCn.Execute ("ALTER TABLE RTOModel Add constraint DF_RTOModel_Model Default '' for Model")
    
    
    GCn.Execute ("ALTER TABLE Veh_Order Drop constraint DF_Veh_Order_Model")
    GCn.Execute ("ALTER TABLE Veh_Order ALTER COLUMN MODEL VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Veh_Order Add constraint DF_Veh_Order_Model Default '' for Model")
    
    GCn.Execute ("ALTER TABLE HisCard Drop constraint DF_HisCard_Model")
    GCn.Execute ("ALTER TABLE HisCard ALTER COLUMN Model VARCHAR(30)")
    GCn.Execute ("ALTER TABLE HisCard Add constraint DF_HisCard_Model Default '' for Model")
    
    GCn.Execute ("ALTER TABLE Job_Booking Drop constraint DF_Job_Booking_Model")
    GCn.Execute ("ALTER TABLE Job_Booking ALTER COLUMN Model VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Job_Booking Add constraint DF_Job_Booking_Model Default '' for Model")
    
    GCn.Execute ("ALTER TABLE Veh_CheckList Drop constraint DF_Veh_CheckList_Model")
    GCn.Execute ("ALTER TABLE Veh_CheckList ALTER COLUMN MODEL VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Veh_CheckList Add constraint DF_Veh_CheckList_Model Default '' for Model")
    
    
    GCn.Execute ("ALTER TABLE Model Drop constraint DF_Model_Model")
    GCn.Execute ("ALTER TABLE Model ALTER COLUMN MODEL VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Model Add constraint DF_Model_Model Default '' for Model")
    
    
    GCn.Execute ("ALTER TABLE Veh_InvCancel Drop constraint DF_Veh_InvCancel_Model")
    GCn.Execute ("ALTER TABLE Veh_InvCancel ALTER COLUMN MODEL VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Veh_InvCancel Add constraint DF_Veh_InvCancel_Model Default '' for Model")
    
    GCn.Execute ("ALTER TABLE Service_Rates Drop constraint DF_Service_Rates_Model")
    GCn.Execute ("ALTER TABLE Service_Rates ALTER COLUMN Model VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Service_Rates Add constraint DF_Service_Rates_Model Default '' for Model")
    
    GCn.Execute ("ALTER TABLE Veh_Quot1 Drop constraint DF_Veh_Quot1_Model")
    GCn.Execute ("ALTER TABLE Veh_Quot1 ALTER COLUMN MODEL VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Veh_Quot1 Add constraint DF_Veh_Quot1_Model Default '' for Model")
    
    
    GCn.Execute ("ALTER TABLE Estimate Drop constraint DF_Estimate_Model")
    GCn.Execute ("ALTER TABLE Estimate ALTER COLUMN Model VARCHAR(30)")
    GCn.Execute ("ALTER TABLE Estimate Add constraint DF_Estimate_Model Default '' for Model")
    
    
    GCn.Execute ("ALTER TABLE RTOData Drop constraint DF_RTOData_Model")
    GCn.Execute ("ALTER TABLE RTOData ALTER COLUMN MODEL VARCHAR(30)")
    GCn.Execute ("ALTER TABLE RTOData Add constraint DF_RTOData_Model Default '' for Model")
    
    GCn.Execute ("ALTER TABLE FinGroupSummary Drop constraint DF_FinGroupSummary_Model")
    GCn.Execute ("ALTER TABLE FinGroupSummary ALTER COLUMN MODEL VARCHAR(30)")
    GCn.Execute ("ALTER TABLE FinGroupSummary Add constraint DF_FinGroupSummary_Model Default '' for Model")
    
    
    
    '################ END OF  MODEL FIELD WIDTH INCREASE   ################
    
    
    
    GCn.Execute "Alter Table Model Alter Column Model_Desc  VarChar(80) "
    GCn.Execute "Alter Table Sp_Purch Alter Column Party_Doc_No  VarChar(25) "
    GCn.Execute "Alter Table Sp_Stock Alter Column PurDocNo  VarChar(25) "
    GCn.Execute "Alter Table Veh_Purch1 Alter Column PBill_No VarChar(25) "
    GCn.Execute "Alter Table Veh_Purch1 Alter Column OBNo VarChar(20) "
    GCn.Execute "Alter Table Veh_Stock Alter Column PBill_No VarChar(25) "
    GCn.Execute "Alter table FinBank ALTER COLUMN FinBankCode VarChar(5)"
    GCn.Execute "Alter table FinBank ALTER COLUMN FinBankName VarChar(50)"
    GCn.Execute "Alter table ContractFinance ALTER COLUMN FinBankCode VarChar(5)"
    GCn.Execute "Alter table ContractFinance ALTER COLUMN FinName NVarChar(50)"
            
    GCn.Execute "Alter Table  PART ALTER COLUMN MRP Numeric(18,3)"
    GCn.Execute "Alter Table  PART ALTER COLUMN TB_SRATE Numeric(18,3)"
    GCn.Execute "Alter Table  PART ALTER COLUMN TP_SRATE Numeric(18,3)"
    
    GCn.Execute "Alter table SubGroup ALTER COLUMN Nature TEXT(25)"
    GCn.Execute "Alter table SubGroupAlias ALTER COLUMN Nature TEXT(25)"
    GCn.Execute "Alter Table DeleteLog Alter Column VType VarChar(30) "
    GCn.Execute "Alter Table Voucher_Type Add SiteBaseNumber NVarChar(1) Default '' "
    GCn.Execute "Alter Table Voucher_Prefix Add Site_Code NVarChar(1) Default '' "
    
    GCn.Execute "Alter Table  Veh_Order  Add SpecialDiscount Float Default 0 "
    GCn.Execute "Alter Table  AcControls Add SpecialDiscountAc nVarchar(8) "
    
    
    
    
    
    
    '########### R E C O R D S     F O R     S Y N C H R O N I S A T I O N ###############

    If StrCmp(left(PubComp_Name, 6), "prayag") Then
        GCn.Execute "INSERT INTO dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('AMD_Dealer','D_Code','D_Code','U_EntDt','Trf_Date') "
        GCn.Execute "INSERT INTO dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Area','AreaCode','AreaCode','U_EntDt','Trf_Date') "
        GCn.Execute "Alter Table BodyBuilder ADD Trf_Date SMALLDATETIME "
        GCn.Execute "INSERT INTO dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('BodyBuilder','BodyBuilderCode','BodyBuilderCode','U_EntDt','Trf_Date') "
        GCn.Execute "Alter Table BodyType ADD Trf_Date SMALLDATETIME "
        GCn.Execute "INSERT INTO dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('BodyType','BodyTypeCode','BodyTypeCode','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('City','CityCode','CityCode','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('ColMast','Col_Code','Col_Code','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('ContractFinance','FinCode','FinCode','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Designation','Designation','Designation','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Emp_Mast','Emp_Code','Emp_Code','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('EmpMast','Emp_Code','Emp_Code','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('FinBank','FinBankCode','FinBankCode','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('FinGroup','FinGrpCode','FinGrpCode','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Godown','God_Code','God_Code','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('HisCard','CardNo','CardNo','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Ledger','DocId','DocId','U_EntDt','Trf_Date')"
        GCn.Execute "Alter Table LedgerRef ADD Trf_Date SMALLDATETIME "
        GCn.Execute "Alter Table LedgerAdj ADD Trf_Date SMALLDATETIME "
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('LedgerAdj','DocId1','DocId1','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('LedgerM','DocId','DocId','U_EntDt','Trf_Date')"
        GCn.Execute "Alter Table LedgerTDS  ADD Trf_Date SMALLDATETIME "
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('LedgerTds','DocId','DocId','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Model','Model','Model','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Model_Cat','ModelCat_Code','ModelCat_Code','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Model_Grp','ModelGrp_Code','ModelGrp_Code','U_EntDt','Trf_Date')"
        GCn.Execute "Alter Table OffTake ADD Trf_Date SMALLDATETIME "
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('OffTake','Code','Code','U_EntDt','Trf_Date')"
        GCn.Execute "Alter Table OffTake1 ADD Trf_Date SMALLDATETIME "
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('OffTake1','Code','Code','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('ProspectiveCust','Cust_Code','Cust_Code','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Site','Site_Code','Site_Code','U_EntDt','Trf_Date')"
        GCn.Execute "Alter Table SubGroup ADD Trf_Date SMALLDATETIME "
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Subgroup','SubCode','SubCode','U_EntDt','Trf_Date')"
        GCn.Execute "Alter Table SubGroupAlias ADD Trf_Date SMALLDATETIME "
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('SubgroupAlias','SubCode','SubCode','U_EntDt','Trf_Date')"
        GCn.Execute "Alter Table Subvention  ADD Trf_Date SMALLDATETIME "
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Subvention','SchemeNo','SchemeNo','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('TaxForms','Form_Code','Form_Code','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('TaxFormsAc','Form_Code','Form_Code','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('TaxFormStk','Form_Code+FormNo','Form_Code+FormNo','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_AddiService','AddSrvCode','AddSrvCode','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_AddiService','AddSrvCode','AddSrvCode','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_AmdModel','Prod_Code','Prod_Code','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_InvCancel','OrdDocId+Convert(nVarChar,Cancel_Date)','OrdDocId+Convert(nVarChar,Cancel_Date)','U_EntDt','Trf_Date')"
        GCn.Execute "Alter Table Veh_Margin  ADD Trf_Date SMALLDATETIME "
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_Margin','Convert(nVarChar,Sl_No)','Convert(nVarChar,Sl_No)','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_OfftakeIncentive','SrlNo','SrlNo','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_Order','OrdDocId','OrdDocId','U_EntDt','Trf_Date')"
        GCn.Execute "Alter Table Veh_Order1  ADD Trf_Date SMALLDATETIME "
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_Order1','OrdDocId','OrdDocId','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_Purch1','DocId','DocId','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_Purch2','DocId','DocId','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_Quot','DocId','DocId','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_Quot1','DocId','DocId','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_Quot2','DocId','DocId','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_Stock','ChassisNo','ChassisNo','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Veh_Transfer','DocId','DocId','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Visits','Convert(nVArChar,VisitDate)+Rep_Code+Convert(nVarChar,SrlNo)','Convert(nVArChar,VisitDate)+Rep_Code+Convert(nVarChar,SrlNo)','U_EntDt','Trf_Date')"
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Rect','DocId','DocId','U_EntDt','Trf_Date')"
        GCn.Execute "Alter Table dbo.Voucher_Prefix  ADD Trf_Date SMALLDATETIME "
        GCn.Execute "Insert Into dbo.Synchronisation_Fields (TableName,SearchKey,UniqueKey,UpdateDateField,UploadDateField) VALUES ('Voucher_Prefix','CONVERT(NVARCHAR,V_Type)+Convert(nVarChar,Div_Code)+Convert(nVarChar,Date_From)+Convert(nVarChar,Date_to)+Convert(nVarChar,Prefix)+Convert(nVarChar,Site_Code)','CONVERT(NVARCHAR,V_Type)+Convert(nVarChar,Div_Code)+Convert(nVarChar,Date_From)+Convert(nVarChar,Date_to)+Convert(nVarChar,Prefix)+Convert(nVarChar,Site_Code)','','Trf_Date')"
    End If
    
    
    
    
    
    
    
    
    
    
    If UCase(left(PubComp_Name, 3)) = "LMP" Then
        GCn.Execute "Update Syctrl Set TaxOnFreeLabYN=1"
    End If
    
        Set RsUser = G_CompCn.Execute("Select User_Name From UserMast")
        
        If RsUser.RecordCount > 0 Then
            Do Until RsUser.EOF
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FaGrEnt' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FaGrEnt', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='frmSubGroup' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'frmSubGroup', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FAVRENT' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FAVRENT', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FAVTYPE' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FAVTYPE', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FATDSCAT' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FATDSCAT', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                            
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FATDSCERTIFICATE' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FATDSCERTIFICATE', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                            
                If G_CompCn.Execute("Select * From User2 Where User_Name = '" & RsUser!user_name & "' And Module_Name='Account' And Form_Code='FATDSCHAL' And Comp_Code = '" & PubCenCompCode & "' And Div_Code = '" & PubDivCode & "'").RecordCount = 0 Then
                    G_CompCn.Execute "Insert Into User2 (User_Name, Module_Name, Form_Code, Param_Str, Comp_Code, Div_Code) " & _
                                   "Values('" & RsUser!user_name & "', 'Account', 'FATDSCHAL', 'AEDP', '" & PubCenCompCode & "', '" & PubDivCode & "')"
                End If
                            
                RsUser.MoveNext
            Loop
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
        If PubSiebelActiveYn Then
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='SBLCQ'")
            If TmpRst.RecordCount = 0 Then
                G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                              "('GENFA','','SBLCQ','Siebel Receipts - Cheque/Draft','SiebelReceiptsChequeDraft','SBLCD','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
                G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('SBLCQ'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'SBLCQ',700000,'" & PubDivCode & "')")
            End If
            
            
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='SBLCS'")
            If TmpRst.RecordCount = 0 Then
                G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                              "('GENFA','','SBLCS','Siebel Receipts - Cash','SiebelReceiptsCash','SBLCS','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
                G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('SBLCS'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'SBLCS',700000,'" & PubDivCode & "')")
            End If
            
            
            Set TmpRst = G_FaCn.Execute("Select * from Voucher_type where V_Type='SBLRO'")
            If TmpRst.RecordCount = 0 Then
                G_FaCn.Execute ("Insert Into Voucher_Type(Category,NCat,V_Type,Description,Description_Help,Short_Name,Number_Method,Separate_Narr,Common_Narr,Narration,ChqNo,ChqDt,ClgDt,Print_VNo,U_Name,U_EntDt,U_AE) Values " & _
                              "('GENFA','','SBLRO','Siebel Receipt - Release Order','SiebelReceiptReleaseOrder','SBLRO','Automatic','N','N','N','N','N','','N','" & pubUName & "'," & FaConvertDate(PubLoginDate) & ",'A'" & ")")
                G_FaCn.Execute ("Insert Into Voucher_Prefix(V_Type,Date_From,Date_To,Prefix,Start_Srl_No,Div_CODE) Values ('SBLRO'," & ConvertDate(PubStartDate) & "," & ConvertDate(PubEndDate) & ",'SBLRO',700000,'" & PubDivCode & "')")
            End If
            
        End If
        
        GCn.Execute "Alter Table Syctrl Add DebtorInSupplierHelp Numeric(1) Default 0"
        GCn.Execute "Alter Table Syctrl Add eCessPer Numeric(9,2) Default 0"
        GCn.Execute "Alter Table Part_Grade Add VatPer Numeric(9,2) Default 0"
        GCn.Execute "Alter Table Syctrl Add VatPerOnLube Numeric(9,2) Default 0"
        GCn.Execute "Alter Table Job_Card Add eCessPer Numeric(9,2) Default 0"
        GCn.Execute "Alter Table Job_Card Add eCessAmt Numeric(9,2) Default 0"
        GCn.Execute "Alter Table Job_Card Add FreeWarrLabAmt Numeric(9,2) Default 0"
        GCn.Execute "Alter Table Site Add ShortName nVarchar(5) "
        
        GCn.Execute "Alter Table Job_Demand Add Lab_Code VarChar(8) Default ''"
        GCn.Execute "Alter Table Godown Alter Column God_Name VarChar(30) "
        GCn.Execute "Alter Table Area Alter Column AreaName VarChar(30) "
        
        
        GCn.Execute "Alter Table DmsEnviro Add SprPurchase4Ac nVarChar(8) Default ''"
        GCn.Execute "Alter Table DmsEnviro Add Vat4Ac nVarChar(8) Default ''"
        GCn.Execute "Alter Table Division Add LstNo nVarChar(30) Default ''"
        GCn.Execute "Alter Table Division Add LstDate SmallDateTime"
        GCn.Execute "Alter Table Division Add LstNoV nVarChar(30) Default ''"
        GCn.Execute "Alter Table Division Add LstDateV SmallDateTime"
        GCn.Execute "Alter Table Division Add LstNoS nVarChar(30) Default ''"
        GCn.Execute "Alter Table Division Add LstDateS SmallDateTime"
        GCn.Execute "Alter Table Division Add LstNoW nVarChar(30) Default ''"
        GCn.Execute "Alter Table Division Add LstDateW SmallDateTime"
        GCn.Execute "Alter Table AssociatedFirms Add LstNo nVarChar(30) Default ''"
        GCn.Execute "Alter Table AssociatedFirms Add LstDate SmallDateTime"
        
        
        GCn.Execute "Alter Table Job_Demand Add Lab_Rate Numeric(9,2) Default 0"
        GCn.Execute "Alter Table Job_Demand Add Time_Req Numeric(9,2) Default 0"
        GCn.Execute "Alter Table Job_Demand Add Amount Numeric(9,2) Default 0"
        GCn.Execute "Alter Table Job_Card Add TempCloseDate SmallDateTime"
        
        
            
'            GCn.Execute "Alter Table Veh_Rate Drop Constraint PK_Veh_Rate"
'            GCn.Execute "Alter Table Veh_Rate Alter Column Model VarChar(24) Not Null"
'            GCn.Execute "Alter Table Veh_Rate Add Constraint PK_Veh_Rate Primary Key (Model,Effective_Date, Rso_Work, Taxable_Yn)"
'            GCn.Execute "Alter Table Part Alter Column MRP Numeric(9,3)"
'            GCn.Execute "Alter Table Part Alter Column Tp_SRate Numeric(9,3)"
'            GCn.Execute "Alter Table Part Alter Column Tb_SRate Numeric(9,3)"
            
        
        
        ''''''''''For Adding Records In SubGroupType Table'''''''''''
        If GCn.Execute("Select Count(*) From SubGroupType").Fields(0) = 0 Then
            GCn.Execute "Insert Into SubGroupType (Party_Type, Description, Mrp_Disc, TB_Disc, Tp_Disc, U_Name, U_EntDt, U_AE) " & _
                        "Values(0, 'General', 0, 0, 0, 'SA', " & ConvertDate(PubLoginDate) & ", 'A')"
            GCn.Execute "Insert Into SubGroupType (Party_Type, Description, Mrp_Disc, TB_Disc, Tp_Disc, U_Name, U_EntDt, U_AE) " & _
                        "Values(1, 'Dealer', 0, 0, 0, 'SA', " & ConvertDate(PubLoginDate) & ", 'A')"
            GCn.Execute "Insert Into SubGroupType (Party_Type, Description, Mrp_Disc, TB_Disc, Tp_Disc, U_Name, U_EntDt, U_AE) " & _
                        "Values(2, 'Retailer', 0, 0, 0, 'SA', " & ConvertDate(PubLoginDate) & ", 'A')"
            GCn.Execute "Insert Into SubGroupType (Party_Type, Description, Mrp_Disc, TB_Disc, Tp_Disc, U_Name, U_EntDt, U_AE) " & _
                        "Values(3, 'Auth. Spare', 0, 0, 0, 'SA', " & ConvertDate(PubLoginDate) & ", 'A')"
            GCn.Execute "Insert Into SubGroupType (Party_Type, Description, Mrp_Disc, TB_Disc, Tp_Disc, U_Name, U_EntDt, U_AE) " & _
                        "Values(4, 'Individual', 0, 0, 0, 'SA', " & ConvertDate(PubLoginDate) & ", 'A')"
            GCn.Execute "Insert Into SubGroupType (Party_Type, Description, Mrp_Disc, TB_Disc, Tp_Disc, U_Name, U_EntDt, U_AE) " & _
                        "Values(5, 'Telco', 0, 0, 0, 'SA', " & ConvertDate(PubLoginDate) & ", 'A')"
        End If
        
        If GCn.Execute("Select Count(*) From Purpose").Fields(0) = 0 Then
            GCn.Execute "Insert Into Purpose (PurposeCode, Site_Code, PurposeName, U_Name, U_EntDt, U_AE) " & _
                        "Values ('1', '" & PubSiteCode & "', 'Personal', 'SA', " & ConvertDate(PubLoginDate) & ", 'A')"
            GCn.Execute "Insert Into Purpose (PurposeCode, Site_Code, PurposeName, U_Name, U_EntDt, U_AE) " & _
                        "Values ('2', '" & PubSiteCode & "', 'Commercial', 'SA', " & ConvertDate(PubLoginDate) & ", 'A')"
            GCn.Execute "Insert Into Purpose (PurposeCode, Site_Code, PurposeName, U_Name, U_EntDt, U_AE) " & _
                        "Values ('3', '" & PubSiteCode & "', 'Other', 'SA', " & ConvertDate(PubLoginDate) & ", 'A')"
        End If
        
                
        If GCn.Execute("Select " & xIsNull("VBilRptName", "") & " From Division Where Div_Code='" & PubDivCode & "'").Fields(0).Value = "" Then
            GCn.Execute "Update Division  Set VBilRptName='VehBill-SBL' Where Div_Code='" & PubDivCode & "'"
        End If
        
        
        
        
        
    
    
    
        GCn.Execute "Alter Table Subvention Drop Constraint PK__Subvention__65CC03DF"
        GCn.Execute "Alter Table Subvention Add Constraint PK__Subvention__65CC03DF Primary Key (SchemeNo, ModelGroup)"
                
        G_CompCn.Execute "Alter Table Company Add  DatabaseName nVarChar(10) Default ''"
        
        
        
        
        GCn.Execute "Alter Table DeleteLog Add Type varchar(20)"
        GCn.Execute "Alter Table DeleteLog Add VType Varchar(10)"
        GCn.Execute "Alter Table DeleteLog Add VDate Varchar(15)"
        GCn.Execute "Alter Table DeleteLog Add Total_Item Numeric(18, 3)"
        GCn.Execute "Alter Table DeleteLog add Total_Qty Numeric(18, 3)"
        GCn.Execute "Alter Table DeleteLog add GoodsValue Numeric(18, 3)"
        GCn.Execute "Alter Table DeleteLog add Discount Numeric(18, 3)"
        GCn.Execute "Alter Table DeleteLog add Addition Numeric(18, 3)"
        GCn.Execute "Alter Table DeleteLog add Deduction Numeric(18, 3)"
        GCn.Execute "Alter Table DeleteLog add LabDiscount Numeric(18, 3)"
        GCn.Execute "Alter Table DeleteLog add LabAmount Numeric(18, 3)"
        GCn.Execute "Alter Table DeleteLog add AutoYn Varchar(1)"
        GCn.Execute "Alter Table DeleteLog add EditDate SmallDateTime"
        GCn.Execute "Alter Table DeleteLog add EditTime Varchar(20)"
        
        GCn.Execute "Alter table Job_Warr1 ALTER COLUMN Cmpl_Date nVarChar(20)"
        GCn.Execute "Alter table Job_Warr1 ALTER COLUMN Repair_Date nVarChar(20)"
        GCn.Execute "Alter table Job_Warr1 ALTER COLUMN PCR_Date nVarChar(20)"
        
        If GCn.Execute("Select Count(*) From INFORMATION_SCHEMA.columns WHERE TABLE_NAME= 'Veh_Stock' And Column_Name='ChassisNo' And Character_Maximum_Length=20 ").Fields(0).Value <= 0 Then
            GCn.Execute "ALTER TABLE Veh_OrderM ALTER COLUMN Chassis NVARCHAR(20)"
            GCn.Execute "ALTER TABLE CustInfo ALTER COLUMN Chassis NVARCHAR(20)"
            GCn.Execute "ALTER TABLE Job_Inspection ALTER COLUMN Chassis NVARCHAR(20)"
            GCn.Execute "ALTER TABLE Veh_Order1 ALTER COLUMN Chassis NVARCHAR(20)"
            GCn.Execute "ALTER TABLE Veh_Order ALTER COLUMN Chassis NVARCHAR(20)"
            GCn.Execute "ALTER TABLE HisCard ALTER COLUMN Chassis NVARCHAR(20)"
            GCn.Execute "ALTER TABLE Job_Booking ALTER COLUMN Chassis NVARCHAR(20)"
            GCn.Execute "ALTER TABLE Veh_InvCancel ALTER COLUMN Chassis NVARCHAR(20)"
            GCn.Execute "ALTER TABLE Estimate ALTER COLUMN Chassis NVARCHAR(20)"
            GCn.Execute "Alter Table RTOData Drop Constraint PK_RTOData "
            GCn.Execute "ALTER TABLE RTOData ALTER COLUMN CHASSIS_NO NVARCHAR(20) NOT Null"
            GCn.Execute "ALTER TABLE RtoData DROP CONSTRAINT DF_RtoData_Chassis_No"
            GCn.Execute "Alter Table RTOData Add Constraint DF_RtoData_Chassis_No DEFAULT '' For CHASSIS_NO"
            GCn.Execute "Alter Table RTOData Add Constraint PK_RTOData Primary Key (CHASSIS_NO,MODEL)"
            GCn.Execute "Alter Table Veh_Stock Drop Constraint PK_Veh_Stock"
            GCn.Execute "ALTER TABLE Veh_Stock ALTER COLUMN ChassisNo NVARCHAR(20) NOT Null"
            GCn.Execute "ALTER TABLE Veh_Stock DROP CONSTRAINT DF_Veh_Stock_ChassisNo"
            GCn.Execute "Alter Table Veh_Stock Add Constraint DF_Veh_Stock_ChassisNo DEFAULT '' For ChassisNo"
            GCn.Execute "Alter Table Veh_Stock Add Constraint PK_Veh_Stock Primary Key (CHASSISNO)"
            
            GCn.Execute "ALTER TABLE Veh_Transfer ALTER COLUMN ChassisNo NVARCHAR(20)"
            GCn.Execute "Alter Table Veh_CheckList Drop Constraint PK_Veh_CheckList"
            GCn.Execute "ALTER TABLE Veh_CheckList ALTER COLUMN ChassisNo NVARCHAR(20) NOT Null"
            GCn.Execute "ALTER TABLE Veh_CheckList DROP CONSTRAINT DF_Veh_CheckList_ChassisNo"
            GCn.Execute "Alter Table Veh_CheckList Add Constraint DF_Veh_CheckList_ChassisNo DEFAULT '' For ChassisNo"
            GCn.Execute "Alter Table Veh_CheckList Add Constraint PK_Veh_CheckList Primary Key (ChassisNo,Item_Code,MODEL)"
            
            
            GCn.Execute "ALTER TABLE Job_Inspection ALTER COLUMN Engine NVARCHAR(25)"
            GCn.Execute "ALTER TABLE HisCard ALTER COLUMN Engine NVARCHAR(25)"
            GCn.Execute "ALTER TABLE Job_Booking ALTER COLUMN Engine NVARCHAR(25)"
            GCn.Execute "ALTER TABLE Job_Warr1 ALTER COLUMN Engine NVARCHAR(25)"
            GCn.Execute "ALTER TABLE Estimate ALTER COLUMN Engine NVARCHAR(25)"
            
            GCn.Execute "ALTER TABLE RTOData ALTER COLUMN ENGINE_NO NVARCHAR(25)"
            
            GCn.Execute "ALTER TABLE Veh_Stock ALTER COLUMN EngineNo NVARCHAR(25)"
            GCn.Execute "ALTER TABLE Veh_Transfer ALTER COLUMN EngineNo NVARCHAR(25)"
        End If
        GCn.Execute "Alter table DmsData ALTER COLUMN Chassis nVarchar(20)"
        GCn.Execute "Alter Table Sp_Stock add RateType Varchar(5)"
        GCn.Execute "Alter Table Job_Lab add RateType Varchar(5)"
        GCn.Execute "Alter Table Job_Card add Variation_Spare float"
        GCn.Execute "Alter Table Job_Card add Variation_Labour float"
        
        
        If StrCmp(left(PubComp_Name, 6), "prayag") Then CreateSqlTriggers
        
        G_FaCn.Execute "Update Voucher_Type Set NCat=V_Type Where NCat = '' "
        
Exit Sub
DispErr:
    MsgBox err.Description
    Resume Next
End Sub

    Sub CreateSqlTriggers()
        Dim mQry$
        Dim RsTemp As Recordset
        Dim I As Integer

        On Error GoTo ELoop

        mQry = "SELECT * FROM Synchronisation_Fields WHERE IsNull(SearchKey,'')<>''"
        Set RsTemp = GCn.Execute(mQry)
        With RsTemp
            For I = 0 To RsTemp.RecordCount - 1
                mQry = "IF EXISTS (SELECT * FROM sys.Triggers WHERE name ='Tr_" & XNull(.Fields("TableName")) & "') "
                mQry = mQry & " Drop Trigger Tr_" & XNull(.Fields("TableName")) & " "
                GCn.Execute mQry



                mQry = "CREATE TRIGGER Tr_" & XNull(.Fields("TableName")) & " ON " & XNull(.Fields("TableName")) & "  " & _
                        "FOR INSERT, Update " & _
                        "AS " & _
                        "DECLARE @UniqueKey NVARCHAR(Max) " & _
                        "DECLARE @UDate DateTime " & _
                        "IF EXISTS (SELECT * FROM Inserted)      " & _
                        "   BEGIN  " & _
                        "       SET @UniqueKey=(SELECT Top 1 " & XNull(.Fields("UniqueKey")) & " FROM Inserted i) " & _
                        "       Set @UDate = (SELECT GetDate()) " & _
                        "       IF NOT UPDATE (" & .Fields("UploadDateField") & ") " & _
                        "       Begin        " & _
                        "           Update " & XNull(.Fields("TableName")) & " Set " & .Fields("UploadDateField") & "=Null Where " & .Fields("UniqueKey") & "=@UniqueKey    " & _
                        "       End        " & _
                        "   END "
                GCn.Execute mQry

                RsTemp.MoveNext
            Next
        End With
        
Exit Sub
ELoop:
    MsgBox err.Description
    Resume Next
    End Sub




Sub CreateNewTableSQL()
Dim mQry$
On Error Resume Next

    If G_CompCn.Execute("Select IsNull(Count(*),0) from sysColumns where id = object_id('UserGroup1') ").Fields(0).Value = 0 Then
        G_CompCn.Execute "SELECT User_Name, Module_Name, Form_Code, Param_Str INTO UserGroup1 FROM User2 WHERE 1=2"
    End If
    If G_CompCn.Execute("Select IsNull(Count(*),0) from sysColumns where id = object_id('UserGroup') ").Fields(0).Value = 0 Then
        G_CompCn.Execute "Create Table UserGroup (User_Name nVarChar(10))"
    End If
    
    
    If GCn.Execute("Select IsNull(Count(*),0) from sysColumns where id = object_id('DmsErrLog') ").Fields(0).Value = 0 Then
         GCn.Execute "CREATE TABLE dbo.DmsErrLog " & _
                        "    ( " & _
                        "    Cat VARCHAR (50) NULL, " & _
                        "    [Key] VARCHAR (50) NOT NULL, " & _
                        "    Narration  VarChar(200) Not NULL,U_EntDt         SMALLDATETIME NOT NULL " & _
                        "    )"
                        
                        
    End If
    
    
    If G_CompCn.Execute("Select IsNull(Count(*),0) from sysColumns where id = object_id('dbo.User_Site') ").Fields(0).Value = 0 Then
        G_CompCn.Execute "CREATE TABLE dbo.User_Site " & _
                        "    ( " & _
                        "    Site_Code VARCHAR (1) NULL, " & _
                        "    User_Name VARCHAR (10) NOT NULL, " & _
                        "    Comp_Code  VarChar(2) Not NULL, " & _
                        "    )"
        If StrCmp(left(PubComp_Name, 6), "PRAYAG") Then
            G_CompCn.Execute "INSERT INTO DBO.User_Site (Site_Code, User_Name, Comp_Code) VALUES ('1', 'ACCOUNT', '08')"
            G_CompCn.Execute "INSERT INTO DBO.User_Site (Site_Code, User_Name, Comp_Code) VALUES ('2', 'ACCOUNT', '08')"
            G_CompCn.Execute "INSERT INTO DBO.User_Site (Site_Code, User_Name, Comp_Code) VALUES ('2', 'MALAY', '08')"
            G_CompCn.Execute "INSERT INTO DBO.User_Site (Site_Code, User_Name, Comp_Code) VALUES ('1', 'RAHUL', '08')"
            G_CompCn.Execute "INSERT INTO DBO.User_Site (Site_Code, User_Name, Comp_Code) VALUES ('1', 'TARIQUE', '08')"
            G_CompCn.Execute "INSERT INTO DBO.User_Site (Site_Code, User_Name, Comp_Code) VALUES ('1', 'ANIL', '08')"
            G_CompCn.Execute "INSERT INTO DBO.User_Site (Site_Code, User_Name, Comp_Code) VALUES ('1', 'HARSH', '08')"
            G_CompCn.Execute "INSERT INTO DBO.User_Site (Site_Code, User_Name, Comp_Code) VALUES ('2', 'HARSH', '08')"
            G_CompCn.Execute "INSERT INTO DBO.User_Site (Site_Code, User_Name, Comp_Code) VALUES ('2', 'LOKESH', '08')"
            G_CompCn.Execute "INSERT INTO DBO.User_Site (Site_Code, User_Name, Comp_Code) VALUES ('1', 'SANJAY', '08')"
        End If
    End If
    
    If GCn.Execute("Select IsNull(Count(*),0) from sysColumns where id = object_id('dbo.Payment') ").Fields(0).Value = 0 Then
        GCn.Execute "CREATE TABLE dbo.Payment " & _
                    "( " & _
                    "DocId           NVARCHAR (21) CONSTRAINT DF_Payment_DocId DEFAULT ('') NOT NULL, " & _
                    "Site_Code       NVARCHAR (2) CONSTRAINT DF_Payment_Site_Code DEFAULT ('') NOT NULL, " & _
                    "V_Date          SMALLDATETIME NOT NULL, " & _
                    "V_Type          NVARCHAR (5) CONSTRAINT DF_Payment_V_Type DEFAULT ('') NOT NULL, " & _
                    "V_No            BIGINT CONSTRAINT DF_Payment_V_No DEFAULT ((0)) NOT NULL, " & _
                    "PartyCode       NVARCHAR (8) CONSTRAINT DF_Payment_PartyCode DEFAULT ('') NOT NULL, " & _
                    "Amount          FLOAT CONSTRAINT DF_Table_1_AMOUNT DEFAULT ((0)) NOT NULL, " & _
                    "AcCode          NVARCHAR (8) CONSTRAINT DF_Payment_AcCode DEFAULT ('') NOT NULL, " & _
                    "Chq_No          NVARCHAR (20) CONSTRAINT DF_Table_1_DDNo DEFAULT ('') NULL, Chq_Date        SMALLDATETIME NULL, Clg_Date        SMALLDATETIME NULL, " & _
                    "Narration       NVARCHAR (255) CONSTRAINT DF_Payment_Narration DEFAULT ('') NULL, " & _
                    "PayTo1          NVARCHAR (255) CONSTRAINT DF_Payment_Narration1 DEFAULT ('') NULL, PayTo2          NVARCHAR (255) CONSTRAINT DF_Payment_Narration2 DEFAULT ('') NULL, " & _
                    "Printed         TINYINT CONSTRAINT DF_Payment_Printed DEFAULT ((0)) NULL, " & _
                    "AcPayeeCheque   BIT CONSTRAINT DF_Payment_AcPayeeCheque DEFAULT ((0)) NULL, " & _
                    "AcPostByU_Name  NVARCHAR (10) CONSTRAINT DF_Payment_AcPostByU_Name DEFAULT ('') NULL, " & _
                    "AcPostByU_EntDt SMALLDATETIME NULL, " & _
                    "U_Name          NVARCHAR (10) CONSTRAINT DF_Payment_U_Name DEFAULT ('') NOT NULL, " & _
                    "U_EntDt         SMALLDATETIME NOT NULL, " & _
                    "U_AE            NVARCHAR (1) CONSTRAINT DF_Payment_U_AE DEFAULT ('') NOT NULL, " & _
                    "AddBy           NVARCHAR (10) CONSTRAINT DF_Payment_AddBy DEFAULT ('') NULL, AddDate         DATETIME NULL, " & _
                    "ModifyBy        NVARCHAR (10) CONSTRAINT DF_Payment_ModifyBy DEFAULT ('') NULL, " & _
                    "ModifyDate      DATETIME NULL, " & _
                    "CONSTRAINT PK_Payment PRIMARY KEY (DocId) " & _
                    ")"
    End If
    

        If GCn.Execute("Select IsNull(Count(*),0) from sysColumns where id = object_id('dbo.BodyBuilder') ").Fields(0).Value = 0 Then
            GCn.Execute "Create Table dbo.BodyBuilder(BodyBuilderCode nVarChar(5) Primary Key, " & _
                                               "BodyBuilderDesc nVarChar(50) Not Null, " & _
                                               "Add1 nVarChar(50), " & _
                                               "Add2 nVarChar(50), " & _
                                               "CityCode nVarChar(5), " & _
                                               "Contact nVarChar(50), " & _
                                               "Site_Code nVarChar(2), " & _
                                               "U_EntDt SmallDateTime, " & _
                                               "U_Name nVarChar(10), " & _
                                               "U_AE nVarChar(1)) "
        End If

        If GCn.Execute("Select IsNull(Count(*),0) from sysColumns where id = object_id('dbo.BodyType') ").Fields(0).Value = 0 Then
            GCn.Execute "Create Table dbo.BodyType(BodyTypeCode nVarChar(5) Primary Key, " & _
                                               "BodyTypeDesc nVarChar(50) Not Null, " & _
                                               "Site_Code nVarChar(2), " & _
                                               "U_EntDt SmallDateTime, " & _
                                               "U_Name nVarChar(10), " & _
                                               "U_AE nVarChar(1)) "
        End If


        If GCn.Execute("Select IsNull(Count(*),0) from sysColumns where id = object_id('dbo.Log_Structure') ").Fields(0).Value = 0 Then
            GCn.Execute "Create Table dbo.Log_Structure(Particular nVarChar(50) , " & _
                                               "Remark nVarChar(50) )"
        End If


        GCn.Execute "Create Table dbo.User_Site(" _
                        & " Site_Code   VarChar(1), " _
                        & " User_Name   Varchar(10)          Not Null, " _
                        & " Comp_Code   Varchar(2)           Not Null, " _
                        & ")"
                
                
        GCn.Execute "Create Table dbo.DmsSubGroup(" _
                        & " DmsSubCode          VarChar(15), " _
                        & " Name                VarChar(50), " _
                        & " Add1                VarChar(50), " _
                        & " Add2                VarChar(50), " _
                        & " City                VarChar(50), " _
                        & " PinCode             VarChar(6), " _
                        & " State               VarChar(2), " _
                        & " Phone               VarChar(35), " _
                        & " Fax                 VarChar(24), " _
                        & " Email               VarChar(50), " _
                        & " [Group]               VarChar(20), " _
                        & " Division            VarChar(40), " _
                        & " AutomanSubCode      Varchar(8)       " _
                        & ") "
                                
        
        GCn.Execute "Create Table dbo.Exp_Emp(" _
                        & " DocId         VarChar(21), " _
                        & " V_Type        VarChar(5), " _
                        & " V_Prefix      VarChar(5), " _
                        & " V_No          Numeric(18), " _
                        & " V_Date        SmallDateTime, " _
                        & " ExpAc         VarChar(8), " _
                        & " CashBankAc    VarChar(8), " _
                        & " Amount        Numeric(18,3), " _
                        & " Narration     VarChar(255)" _
                        & ")"
                        
        GCn.Execute "Create Table dbo.Exp_Emp1(" _
                        & " DocId         VarChar(21), " _
                        & " Srl           Numeric(18), " _
                        & " Emp_Code      VarChar(5), " _
                        & " Amount        Numeric(18,3)" _
                        & ")"
                        
        
        GCn.Execute "Create Table dbo.Budget_Exp(" _
                        & " ExpAc         VarChar(8), " _
                        & " Site_Code     VarChar(2), " _
                        & " VDate         SmallDateTime, " _
                        & " Month         VarChar(2), " _
                        & " Amount        Numeric(18,3), " _
                        & " U_Name        VarChar(10), " _
                        & " U_EntDt       SmallDateTime, " _
                        & " U_AE          VarChar(1)" _
                        & ")"


        GCn.Execute "Create Table dbo.Subvention (" _
                        & " SchemeNo            VarChar(20)              , " _
                        & " FromDate            SmallDateTime           , " _
                        & " ToDate              SmallDateTime           , " _
                        & " ModelGroup          VarChar(5)              , " _
                        & " Model               VarChar(24)             , " _
                        & " DealerContribution  Numeric(18,3)           , " _
                        & " TataContribution    Numeric(18,3)           , " _
                        & " TotalSubvention     Numeric(18,3)           , " _
                        & " U_Name              VarChar(10)             , " _
                        & " U_EntDt             SmallDateTime           , " _
                        & " U_AE                VarChar(1)             " _
                        & ")"



        GCn.Execute "Create Table dbo.OffTake(" _
                        & " Code                Numeric(5), " _
                        & " SchemeNo            VarChar(20), " _
                        & " FromDate            SmallDateTime, " _
                        & " ToDate              SmallDateTime, " _
                        & " Qty                 Numeric(18,3), " _
                        & " Amount              Numeric(18,3), " _
                        & " U_Name              VarChar(10), " _
                        & " U_EntDt             SmallDateTime, " _
                        & " U_AE                VarChar(1)" _
                        & ")"


        GCn.Execute "Create Table dbo.OffTake1(" _
                        & " Code                Numeric(5), " _
                        & " SrlNo               Numeric(5),   " _
                        & " ModelGrp            VarChar(5)" _
                        & ")"
                        
                        
        GCn.Execute "CREATE TABLE dbo.RateType " & _
                    "   ( " & _
                    "   Code         NVARCHAR (5) CONSTRAINT DF_Table_1_BodyBuilderCode DEFAULT ('') NOT NULL, " & _
                    "   Description  NVARCHAR (50) CONSTRAINT DF_Table_1_BodyBuilderDesc DEFAULT ('') NOT NULL, " & _
                    "   VariationPer FLOAT CONSTRAINT DF_Table_1_Add1 DEFAULT ((0)) NOT NULL, " & _
                    "   Site_Code    NVARCHAR (2) CONSTRAINT DF_RateType_Site_Code DEFAULT ('') NOT NULL, " & _
                    "   U_Name       NVARCHAR (15) CONSTRAINT DF_RateType_U_Name DEFAULT ('') NOT NULL, " & _
                    "   U_EntDt      SMALLDATETIME NOT NULL, " & _
                    "   U_AE         NVARCHAR (1) CONSTRAINT DF_RateType_U_AE DEFAULT ('') NOT NULL, " & _
                    "   CONSTRAINT PK_RateType PRIMARY KEY (Code), " & _
                    "   CONSTRAINT IX_RateType UNIQUE (Description) " & _
                    "   ) "


        GCn.Execute "CREATE TABLE dbo.Synchronisation_Fields " & _
                    "   ( " & _
                    "   RowId           BIGINT Identity NOT NULL, " & _
                    "   TableName       NVARCHAR (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
                    "   SearchKey       NVARCHAR (21) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
                    "   UniqueKey       NVARCHAR (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
                    "   UpdateDateField NVARCHAR (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
                    "   UploadDateField NVARCHAR (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
                    "   CONSTRAINT PK_Synchronisation_Fields PRIMARY KEY (RowId), " & _
                    "   CONSTRAINT IX_Synchronisation_Fields UNIQUE (TableName) " & _
                    "   ) "

        GCn.Execute "CREATE TABLE dbo.Synchronisation_Errors " & _
                    "   ( " & _
                    "   RowId     BIGINT IDENTITY NOT NULL, " & _
                    "   TableName NVARCHAR (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
                    "   SearchKey NVARCHAR (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
                    "   Message   NVARCHAR (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
                    "   CONSTRAINT PK_Synchronisation_Errors PRIMARY KEY (RowId) " & _
                    "   ) "



        GCn.Execute "CREATE TABLE dbo.Insurance " & _
                    "( " & _
                    "Code          NVARCHAR (5) COLLATE SQL_Latin1_General_CP1_CI_AS CONSTRAINT DF_Insurance_Code DEFAULT ('') NULL, " & _
                    "Name          NVARCHAR (50) COLLATE SQL_Latin1_General_CP1_CI_AS CONSTRAINT DF_Insurance_Name DEFAULT ('') NULL, " & _
                    "Add1          NVARCHAR (50) COLLATE SQL_Latin1_General_CP1_CI_AS CONSTRAINT DF_Insurance_Add1 DEFAULT ('') NULL, " & _
                    "Add2          NVARCHAR (50) COLLATE SQL_Latin1_General_CP1_CI_AS CONSTRAINT DF_Insurance_Add2 DEFAULT ('') NULL, " & _
                    "City          NVARCHAR (5) COLLATE SQL_Latin1_General_CP1_CI_AS CONSTRAINT DF_Insurance_City DEFAULT ('') NULL, " & _
                    "ContactPerson NVARCHAR (50) COLLATE SQL_Latin1_General_CP1_CI_AS CONSTRAINT DF_Insurance_ContactPerson DEFAULT ('') NULL, " & _
                    "U_Name        NVARCHAR (15) COLLATE SQL_Latin1_General_CP1_CI_AS CONSTRAINT DF_Insurance_U_Name DEFAULT ('') NULL, " & _
                    "U_EntDt       SMALLDATETIME NULL, " & _
                    "U_AE          NVARCHAR (1) COLLATE SQL_Latin1_General_CP1_CI_AS CONSTRAINT DF_Insurance_U_AE DEFAULT ('') NULL )"



        GCn.Execute "CREATE TABLE dbo.Rect1 " & _
                    "( " & _
                    "DocID   NVARCHAR (21) NOT NULL, " & _
                    "Sr      INT NOT NULL, " & _
                    "ChqNo   NVARCHAR (20) NOT NULL, " & _
                    "ChqDate SMALLDATETIME NULL, " & _
                    "ChqAmt  FLOAT CONSTRAINT DF_Rect1_ChqAmt DEFAULT ((0)) NOT NULL, " & _
                    "CONSTRAINT PK_Rect1 PRIMARY KEY (DocID,Sr), " & _
                    "CONSTRAINT FK_Rect1_Rect FOREIGN KEY (DocID) REFERENCES dbo.Rect (DocId) " & _
                    ") "



    mQry = "CREATE TABLE dbo.Body_Purch " & _
                "( " & _
                "DocID           NVARCHAR (21) CONSTRAINT DF_Body_Purch_DocID DEFAULT ('') NOT NULL, " & _
                "DocIDHelp       NVARCHAR (21) CONSTRAINT DF_Body_Purch_DocIDHelp DEFAULT ('') NOT NULL, " & _
                "Site_Code       NVARCHAR (2) CONSTRAINT DF_Body_Purch_Site_Code DEFAULT ('') NOT NULL, " & _
                "V_Type          NVARCHAR (5) CONSTRAINT DF_Body_Purch_V_Type DEFAULT ('') NOT NULL, " & _
                "V_NO            INT CONSTRAINT DF_Body_Purch_V_NO DEFAULT ((0)) NOT NULL, " & _
                "V_Date          SMALLDATETIME NOT NULL, " & _
                "PARTYCODE       NVARCHAR (8) CONSTRAINT DF_Body_Purch_PARTYCODE DEFAULT ('') NOT NULL, " & _
                "PBILL_NO        VARCHAR (25) CONSTRAINT DF_Body_Purch_PBILL_NO DEFAULT ('') NULL, " & _
                "PBILL_DATE      SMALLDATETIME NULL, "
    mQry = mQry + "Form_Code       NVARCHAR (4) CONSTRAINT DF_Body_Purch_Form_Code DEFAULT ('') NULL, " & _
                "AMOUNT          NUMERIC (18,2) CONSTRAINT DF_Body_Purch_AMOUNT DEFAULT ((0)) NULL, " & _
                "Addition        NUMERIC (18,2) CONSTRAINT DF_Body_Purch_Addition DEFAULT ((0)) NULL, " & _
                "Deduction       NUMERIC (18,2) CONSTRAINT DF_Body_Purch_Deduction DEFAULT ((0)) NULL, " & _
                "Exsice          NUMERIC (18,2) CONSTRAINT DF_Body_Purch_Exsice DEFAULT ((0)) NULL, " & _
                "TAX_PER         NUMERIC (18,2) CONSTRAINT DF_Body_Purch_TAX_PER DEFAULT ((0)) NULL, " & _
                "TAX_Amt         NUMERIC (18,2) CONSTRAINT DF_Body_Purch_TAX_Amt DEFAULT ((0)) NULL, " & _
                "TaxSur_Per      NUMERIC (18,2) CONSTRAINT DF_Body_Purch_TaxSur_Per DEFAULT ((0)) NULL, " & _
                "TaxSur_Amt      NUMERIC (18,2) CONSTRAINT DF_Body_Purch_TaxSur_Amt DEFAULT ((0)) NULL, " & _
                "MISC_AMT        NUMERIC (18,2) CONSTRAINT DF_Body_Purch_MISC_AMT DEFAULT ((0)) NULL, " & _
                "Tot_AMOUNT      NUMERIC (18,2) CONSTRAINT DF_Body_Purch_Tot_AMOUNT DEFAULT ((0)) NULL, " & _
                "P_AMOUNT        NUMERIC (18,2) CONSTRAINT DF_Body_Purch_P_AMOUNT DEFAULT ((0)) NULL, " & _
                "ADJ_AMT         NUMERIC (18,2) CONSTRAINT DF_Body_Purch_ADJ_AMT DEFAULT ((0)) NULL, " & _
                "U_Name          NVARCHAR (10) CONSTRAINT DF_Body_Purch_U_Name DEFAULT ('') NOT NULL, "
    mQry = mQry + "U_EntDt         SMALLDATETIME NOT NULL, " & _
                "U_AE            NVARCHAR (1) CONSTRAINT DF_Body_Purch_U_AE DEFAULT ('') NOT NULL, " & _
                "Trf_Date        SMALLDATETIME NULL, " & _
                "DrAcCode        NVARCHAR (8) CONSTRAINT DF_Body_Purch_DrAcCode DEFAULT ('') NOT NULL, " & _
                "AcPostByU_Name  NVARCHAR (10) CONSTRAINT DF_Body_Purch_AcPostByU_Name DEFAULT ('') NULL, " & _
                "AcPostByU_EntDt SMALLDATETIME NULL, " & _
                "AddBy           NVARCHAR (10) CONSTRAINT DF_Body_Purch_AddBy DEFAULT ('') NULL, " & _
                "AddDate         DATETIME NULL, " & _
                "ModifyBy        NVARCHAR (10) CONSTRAINT DF_Body_Purch_ModifyBy DEFAULT ('') NULL, " & _
                "ModifyDate      DATETIME NULL, " & _
                "Sat_Yn          BIT CONSTRAINT DF_Body_Purch_Sat_Yn DEFAULT ((0)) NULL, " & _
                "SatPer          FLOAT CONSTRAINT DF_Body_Purch_SatPer DEFAULT ((0)) NULL, " & _
                "SatAmt          FLOAT CONSTRAINT DF_Body_Purch_SatAmt DEFAULT ((0)) NULL, " & _
                "CONSTRAINT PK_Body_Purch PRIMARY KEY (DocID) " & _
                ")"
                
           GCn.Execute mQry
                
                
                
                mQry = "CREATE TABLE dbo.Body_PurchDetail " & _
                        "( " & _
                        "ChassisNo     NVARCHAR (20) CONSTRAINT DF_Body_PurchDetail_ChassisNo DEFAULT ('') NOT NULL, " & _
                        "Pur_DocId     NVARCHAR (21) CONSTRAINT DF_Body_PurchDetail_Pur_DocId DEFAULT ('') NOT NULL, " & _
                        "Pur_SrlNo     TINYINT CONSTRAINT DF_Body_PurchDetail_Pur_SrlNo DEFAULT ((0)) NOT NULL, " & _
                        "Pur_DocIDHelp NVARCHAR (21) CONSTRAINT DF_Body_PurchDetail_Pur_DocIDHelp DEFAULT ('') NULL, " & _
                        "Pur_SiteCode  NVARCHAR (2) CONSTRAINT DF_Body_PurchDetail_Pur_SiteCode DEFAULT ('') NULL, " & _
                        "Pur_VType     NVARCHAR (5) CONSTRAINT DF_Body_PurchDetail_Pur_VType DEFAULT ('') NULL, " & _
                        "Pur_VNO       INT CONSTRAINT DF_Body_PurchDetail_Pur_VNO DEFAULT ((0)) NULL, " & _
                        "Pur_VDate     SMALLDATETIME NULL, "
                mQry = mQry + "MODEL         VARCHAR (25) CONSTRAINT DF_Body_PurchDetail_MODEL DEFAULT ('') NULL, " & _
                            "Godown        NVARCHAR (3) CONSTRAINT DF_Body_PurchDetail_Godown DEFAULT ('') NULL, " & _
                            "EngineNo      NVARCHAR (25) CONSTRAINT DF_Body_PurchDetail_EngineNo DEFAULT ('') NULL, " & _
                            "VehSerialNo   NVARCHAR (20) CONSTRAINT DF_Body_PurchDetail_VehSerialNo DEFAULT ('') NULL, " & _
                            "RATE          NUMERIC (18,2) CONSTRAINT DF_Body_PurchDetail_RATE DEFAULT ((0)) NULL, " & _
                            "FIXED         NUMERIC (18,2) CONSTRAINT DF_Body_PurchDetail_FIXED DEFAULT ((0)) NULL, " & _
                            "VRATE         NUMERIC (18,2) CONSTRAINT DF_Body_PurchDetail_VRATE DEFAULT ((0)) NULL, " & _
                            "PBILL_NO      VARCHAR (25) CONSTRAINT DF_Body_PurchDetail_PBILL_NO DEFAULT ('') NULL, " & _
                            "PBILL_DATE    SMALLDATETIME NULL, " & _
                            "PartyCode     NVARCHAR (8) CONSTRAINT DF_Body_PurchDetail_PartyCode DEFAULT ('') NULL, " & _
                            "Remarks       NVARCHAR (40) CONSTRAINT DF_Body_PurchDetail_Remarks DEFAULT ('') NULL, " & _
                            "U_Name        NVARCHAR (10) CONSTRAINT DF_Body_PurchDetail_U_Name DEFAULT ('') NOT NULL, " & _
                            "U_EntDt       SMALLDATETIME NOT NULL, " & _
                            "U_AE          NVARCHAR (1) CONSTRAINT DF_Body_PurchDetail_U_AE DEFAULT ('') NOT NULL, " & _
                            "Trf_Date      SMALLDATETIME NULL, " & _
                            "TransAxlNo    NVARCHAR (20) CONSTRAINT DF_Body_PurchDetail_TransAxlNo DEFAULT ('') NULL, " & _
                            "CONSTRAINT PK_Body_PurchDetail PRIMARY KEY (Pur_DocId,Pur_SrlNo) " & _
                            ")"
            GCn.Execute mQry



        If G_CompCn.Execute("SELECT Count(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='SubGroupCounter'").Fields(0).Value = 0 Then
            mQry = "CREATE TABLE dbo.SubGroupCounter " & _
                   "( " & _
                   "SubGroupAcCode Float CONSTRAINT DF_SubGroupCounter_SubGroupAcCode DEFAULT (0) NOT NULL, " & _
                   "CityCode       SMALLINT CONSTRAINT DF_SubGroupCounter_CityCode DEFAULT (0) NULL, " & _
                   "U_Name         NVARCHAR (10) CONSTRAINT DF_SubGroupCounter_U_Name DEFAULT ('') NULL, " & _
                   "U_EntDt        SMALLDATETIME NULL, " & _
                   "U_AE           NVARCHAR (1) CONSTRAINT DF_SubGroupCounter_U_AE DEFAULT ('') NULL " & _
                   ")"
            G_CompCn.Execute mQry
            
            
            mQry = GCn.Execute("Select IsNull(SubGroupAcCode,0) from SubGroupCounter").Fields(0).Value
            mQry = "Insert Into SubGroupCounter(SubGroupAcCode, U_Name, U_EntDt, U_AE) Values (" & Val(mQry) & ", '" & pubUName & "', '" & PubLoginDate & "', 'A')"
            G_CompCn.Execute mQry
        End If




        If GCn.Execute("SELECT Count(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='DmsEnviro'").Fields(0).Value = 0 Then
            mQry = "    CREATE TABLE dbo.DmsEnviro " & _
            "    ( " & _
            "    WsDebtorGroupCode    NVARCHAR (8) CONSTRAINT DF_DmsEnviro_WsDebtorGroupCode DEFAULT ('') NULL, " & _
            "    SprDebtorGroupCode   NVARCHAR (8) CONSTRAINT DF_DmsEnviro_SprDebtorGroupCode DEFAULT ('') NULL, " & _
            "    VehDebtorGroupCode   NVARCHAR (8) CONSTRAINT DF_DmsEnviro_VehDebtorGroupCode DEFAULT ('') NULL, " & _
            "    SprCreditorGroupCode NVARCHAR (8) CONSTRAINT DF_DmsEnviro_SprCreditorGroupCode DEFAULT ('') NULL, " & _
            "    VehCreditorGroupCode NVARCHAR (8) CONSTRAINT DF_DmsEnviro_VehCreditorGroupCode DEFAULT ('') NULL, " & _
            "    SprSaleAc            NVARCHAR (8) CONSTRAINT DF_DmsEnviro_SprSaleAc DEFAULT ('') NULL, " & _
            "    VehSaleAc            NVARCHAR (8) CONSTRAINT DF_DmsEnviro_VehSaleAc DEFAULT ('') NULL, " & _
            "    LubeSaleAc           NVARCHAR (8) CONSTRAINT DF_DmsEnviro_LubeSaleAc DEFAULT ('') NULL, " & _
            "    VatAc                NVARCHAR (8) CONSTRAINT DF_DmsEnviro_VatAc DEFAULT ('') NULL, " & _
            "    WSCashAc             NVARCHAR (8) CONSTRAINT DF_DmsEnviro_WSCashAc DEFAULT ('') NULL, " & _
            "    SprCashAc            NVARCHAR (8) CONSTRAINT DF_DmsEnviro_SprCashAc DEFAULT ('') NULL, " & _
            "    VehCashAc            NVARCHAR (8) CONSTRAINT DF_DmsEnviro_VehCashAc DEFAULT ('') NULL, " & _
            "    LabourAc             NVARCHAR (8) CONSTRAINT DF_DmsEnviro_LabourAc DEFAULT ('') NULL, " & _
            "    ServTaxAc            NVARCHAR (8) CONSTRAINT DF_DmsEnviro_ServTaxAc DEFAULT ('') NULL, " & _
            "    LocalStateName       NVARCHAR (20) CONSTRAINT DF_DmsEnviro_LocalStateName DEFAULT ('') NULL, " & _
            "    SprBankAc            NVARCHAR (8) CONSTRAINT DF_DmsEnviro_SprBankAc DEFAULT ('') NULL, "
            mQry = mQry & "    VehBankAc            NVARCHAR (8) CONSTRAINT DF_DmsEnviro_VehBankAc DEFAULT ('') NULL, " & _
            "    SprPurchaseAc        NVARCHAR (8) CONSTRAINT DF_DmsEnviro_SprPurchaseAc DEFAULT ('') NULL, " & _
            "    VehPurchaseAc        NVARCHAR (8) CONSTRAINT DF_DmsEnviro_VehPurchaseAc DEFAULT ('') NULL, " & _
            "    CstAc                NVARCHAR (8) CONSTRAINT DF_DmsEnviro_CstAc DEFAULT ('') NULL, " & _
            "    ROffAc               NVARCHAR (8) CONSTRAINT DF_DmsEnviro_ROffAc DEFAULT ('') NULL, " & _
            "    WsBankAc             NVARCHAR (8) CONSTRAINT DF_DmsEnviro_WsBankAc DEFAULT ('') NULL, " & _
            "    SprCstPurchaseAc     NVARCHAR (8) CONSTRAINT DF_DmsEnviro_SprCstPurchaseAc DEFAULT ('') NULL, " & _
            "    OtherChargesAc       NVARCHAR (8) CONSTRAINT DF_DmsEnviro_OtherChargesAc DEFAULT ('') NULL, " & _
            "    DiscountAc           NVARCHAR (8) CONSTRAINT DF_DmsEnviro_DiscountAc DEFAULT ('') NULL, " & _
            "    VehPurGroupCode      NVARCHAR (8) CONSTRAINT DF_DmsEnviro_VehPurGroupCode DEFAULT ('') NULL, " & _
            "    VehSaleGroupCode     NVARCHAR (8) CONSTRAINT DF_DmsEnviro_VehSaleGroupCode DEFAULT ('') NULL, " & _
            "    SprPurGroupCode      NVARCHAR (8) CONSTRAINT DF_DmsEnviro_SprPurGroupCode DEFAULT ('') NULL, " & _
            "    SprSaleGroupCode     NVARCHAR (8) CONSTRAINT DF_DmsEnviro_SprSaleGroupCode DEFAULT ('') NULL, " & _
            "    VatGroupCode         NVARCHAR (8) CONSTRAINT DF_DmsEnviro_VatGroupCode DEFAULT ('') NULL, " & _
            "    ServiceTaxGroupCode  NVARCHAR (8) CONSTRAINT DF_DmsEnviro_ServiceTaxGroupCode DEFAULT ('') NULL, " & _
            "    VehCstPurchaseAc     NVARCHAR (8) DEFAULT ('') NULL, " & _
            "    SprSaleVat4Ac        NVARCHAR (8) DEFAULT ('') NULL, " & _
            "    Vat4Ac               NVARCHAR (8) DEFAULT ('') NULL, " & _
            "    SprPurchase4Ac       NVARCHAR (8) DEFAULT ('') NULL, " & _
            "    VatInputAc           NVARCHAR (8) DEFAULT ('') NULL, " & _
            "    Vat4InputAc          NVARCHAR (8) DEFAULT ('') NULL " & _
            "    ) "
            GCn.Execute mQry
            
            
        End If


        If GCn.Execute("SELECT Count(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='DmsBankAc'").Fields(0).Value = 0 Then
            mQry = "CREATE TABLE dbo.DmsBankAc " & _
                "( " & _
                "AutomanBankCode NVARCHAR (8) CONSTRAINT DF_DmsBankAc_AutomanBankCode DEFAULT ('') NULL, " & _
                "DmsBankCode     NVARCHAR (15) CONSTRAINT DF_DmsBankAc_DmsBankCode DEFAULT ('') NULL " & _
                ")"
            GCn.Execute mQry
        End If


        If GCn.Execute("SELECT Count(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME ='DmsSupplierAc'").Fields(0).Value = 0 Then
            mQry = "CREATE TABLE dbo.DmsSupplierAc " & _
                   "( " & _
                   "AutomanSupplierCode NVARCHAR (8) CONSTRAINT DF_DmsSupplierAc_AutomanSupplierCode DEFAULT ('') NULL, " & _
                   "DmsCode     NVARCHAR (15) CONSTRAINT DF_DmsSupplierAc_DmsCode DEFAULT ('') NULL " & _
                   ")"
            GCn.Execute mQry
        End If
        
     If GCn.Execute("Select IsNull(Count(*),0) from sysColumns where id = object_id('dbo.Deprecation_itemMaster') ").Fields(0).Value = 0 Then
           mQry = "CREATE TABLE dbo.Deprecation_itemMaster " & _
                  " (Code        NVARCHAR (5) NOT NULL," & _
                  " ShortName   NVARCHAR (2) NULL," & _
                  " Description NVARCHAR (50) NULL," & _
                  " Dep_per     FLOAT NULL," & _
                  " Site_Code   NVARCHAR (2) CONSTRAINT DF_Deprecation_itemMaster_Site_Code DEFAULT ('') NOT NULL," & _
                  " U_Name      NVARCHAR (15) CONSTRAINT DF_Deprecation_itemMaster_U_Name DEFAULT ('') NOT NULL," & _
                  " U_EntDt     SMALLDATETIME NOT NULL," & _
                  " U_AE        NVARCHAR (1) CONSTRAINT DF_Deprecation_itemMaster_U_AE DEFAULT ('') NOT NULL," & _
                  " CONSTRAINT PK_Deprecation_itemMaster PRIMARY KEY (Code)," & _
                  " CONSTRAINT IX_Deprecation_itemMaster UNIQUE (Description) )"
         GCn.Execute mQry
       End If
      If GCn.Execute("Select IsNull(Count(*),0) from sysColumns where id = object_id('dbo.Deprecation_Master') ").Fields(0).Value = 0 Then
            mQry = "CREATE TABLE dbo.Deprecation_Master " & _
             " (Code      NVARCHAR (5) NOT NULL," & _
             " Dep_Month FLOAT NULL," & _
             " Dep_per   FLOAT NULL," & _
             " Site_Code NVARCHAR (2) CONSTRAINT DF_Deprecation_Master_Site_Code DEFAULT ('') NOT NULL," & _
             " U_Name    NVARCHAR (15) CONSTRAINT DF_Deprecation_Master_U_Name DEFAULT ('') NOT NULL," & _
             " U_EntDt   SMALLDATETIME NOT NULL," & _
             " U_AE      NVARCHAR (1) CONSTRAINT DF_Deprecation_Master_U_AE DEFAULT ('') NOT NULL," & _
             " CONSTRAINT PK_Deprecation_Master PRIMARY KEY (Code)," & _
             " CONSTRAINT IX_Deprecation_Master UNIQUE (Dep_Month))"
            GCn.Execute mQry
     End If
End Sub


Sub Initialise_Pub()

    Dim mWhereCond$, mSQry$
    PubSiteType = RsSite!SiteType
        
    If PubSiteWiseDisplayYn = 1 Then
        If PubSiteType = "H" Then
            PubFaSiteType = 1               '0-General,1-FromSite ForSite,2-2 Char Site
        Else
            PubFaSiteType = 0               '0-General,1-FromSite ForSite,2-2 Char Site
        End If
    Else
        PubFaSiteType = 1
    End If
     
    PubSiteCode = RsSite!Code
    PubSiteName = RsSite!Name
    
   
    If XNull(RsSite!Address1) <> "" Then
        PubComp_Add = XNull(RsSite!Address1)
        PubComp_Add2 = XNull(RsSite!Address2)
        PubComp_Add3 = XNull(RsSite!Address3)
        PubComp_City = XNull(RsSite!City)
        If XNull(RsSite!Phone) = "" And XNull(RsSite!Mobile) = "" Then
            PubComp_Contact = ""
        ElseIf XNull(RsSite!Mobile) = "" Then
            PubComp_Contact = "Phone : " & XNull(RsSite!Phone)
        ElseIf XNull(RsSite!Phone) = "" Then
            PubComp_Contact = "Mobile : " & XNull(RsSite!Mobile)
        End If
        PubComp_TINNo = XNull(RsSite!LstNo)
        PubComp_CstNo = XNull(RsSite!CstNo)
        
    End If

    
                                     
End Sub



