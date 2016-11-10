VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmRepeteJob 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Repeat Job Analysis"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11835
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
      Left            =   1965
      TabIndex        =   61
      Top             =   3180
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
         Picture         =   "frmRepeteJob.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   71
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
         Picture         =   "frmRepeteJob.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmRepeteJob.frx":0678
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
         TabIndex        =   69
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmRepeteJob.frx":0982
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
         TabIndex        =   68
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmRepeteJob.frx":0C8C
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
         TabIndex        =   67
         ToolTipText     =   "Printer "
         Top             =   285
         Visible         =   0   'False
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
         Top             =   720
         Width           =   750
      End
      Begin VB.Line Line9 
         X1              =   1470
         X2              =   1470
         Y1              =   510
         Y2              =   600
      End
      Begin VB.Line Line10 
         X1              =   2820
         X2              =   2820
         Y1              =   630
         Y2              =   735
      End
      Begin VB.Line Line11 
         X1              =   360
         X2              =   360
         Y1              =   615
         Y2              =   720
      End
      Begin VB.Line Line12 
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
         Left            =   -165
         TabIndex        =   74
         Top             =   315
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
         TabIndex        =   73
         Top             =   1275
         Width           =   4650
      End
      Begin VB.Label Label39 
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
         TabIndex        =   72
         Top             =   0
         Width           =   4695
      End
   End
   Begin MSDataGridLib.DataGrid DGJob 
      Height          =   1365
      Left            =   210
      Negotiate       =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3435
      Visible         =   0   'False
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   2408
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   19
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Job_No"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3195.213
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00BAD3C9&
      Height          =   4605
      Left            =   165
      TabIndex        =   25
      Top             =   2295
      Width           =   11385
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3105
         TabIndex        =   15
         Top             =   4230
         Width           =   660
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   7230
         TabIndex        =   16
         Top             =   4185
         Width           =   660
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   10020
         TabIndex        =   17
         Top             =   4050
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   10020
         TabIndex        =   14
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   10005
         TabIndex        =   13
         Top             =   2745
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   10020
         TabIndex        =   12
         Top             =   2085
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   10020
         TabIndex        =   11
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   10035
         TabIndex        =   10
         Top             =   885
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   6060
         MaxLength       =   20
         TabIndex        =   8
         Top             =   750
         Width           =   1725
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   7
         Left            =   2685
         MaxLength       =   25
         TabIndex        =   9
         Top             =   1050
         Width           =   1965
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   2685
         MaxLength       =   40
         TabIndex        =   7
         Top             =   750
         Width           =   1965
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00808080&
         Height          =   615
         Left            =   8295
         TabIndex        =   28
         Top             =   90
         Width           =   0
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   60
         Left            =   30
         TabIndex        =   27
         Top             =   675
         Width           =   11310
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check Ordering System and regular Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   8025
         TabIndex        =   59
         Top             =   3855
         Width           =   1785
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Counselling proper final Inspection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   8040
         TabIndex        =   58
         Top             =   3225
         Width           =   1725
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "In-House Training"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8055
         TabIndex        =   57
         Top             =   2730
         Width           =   1590
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ensure road test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8055
         TabIndex        =   56
         Top             =   2115
         Width           =   1455
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Counselling"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8040
         TabIndex        =   55
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer informed after getting part"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4035
         TabIndex        =   54
         Top             =   4230
         Width           =   3150
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ii)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3795
         TabIndex        =   53
         Top             =   4230
         Width           =   150
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Was the order placed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1050
         TabIndex        =   52
         Top             =   4230
         Width           =   1935
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "i)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   900
         TabIndex        =   51
         Top             =   4230
         Width           =   105
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defect could not be attended as the part was not available."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   900
         TabIndex        =   50
         Top             =   3900
         Width           =   5190
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Technician did not do the proper job due to negligence / casualness."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   870
         TabIndex        =   49
         Top             =   3345
         Width           =   6090
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Technician does not know the correct procedure for doing the job."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   48
         Top             =   2715
         Width           =   5790
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defect was not diagnosed properly as road test was not done."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   47
         Top             =   2070
         Width           =   5490
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "During the previous visit the defect was not diagnosed properly by floor superwiser /service advisor."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   855
         TabIndex        =   46
         Top             =   1410
         Width           =   7050
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Send Field Trouble to Tata Eng."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   8010
         TabIndex        =   45
         Top             =   780
         Width           =   1770
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Batch Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   870
         TabIndex        =   44
         Top             =   1065
         Width           =   1050
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Make"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         TabIndex        =   43
         Top             =   750
         Width           =   510
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Component Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   42
         Top             =   750
         Width           =   1635
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   41
         Top             =   3990
         Width           =   105
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   40
         Top             =   3405
         Width           =   105
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   39
         Top             =   2745
         Width           =   105
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   38
         Top             =   2085
         Width           =   105
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   37
         Top             =   1485
         Width           =   105
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   36
         Top             =   915
         Width           =   105
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   15
         X2              =   11370
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   15
         X2              =   11370
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   15
         X2              =   11370
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   30
         X2              =   11385
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   15
         X2              =   11370
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10020
         TabIndex        =   35
         Top             =   390
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Implement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10005
         TabIndex        =   34
         Top             =   150
         Width           =   1080
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Measures"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8145
         TabIndex        =   33
         Top             =   405
         Width           =   1035
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Counter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8130
         TabIndex        =   32
         Top             =   150
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complaint Fitted in previous service / Repair Field"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   780
         TabIndex        =   31
         Top             =   405
         Width           =   5205
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason / Analysis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   780
         TabIndex        =   30
         Top             =   165
         Width           =   1905
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   195
         TabIndex        =   29
         Top             =   420
         Width           =   375
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         X1              =   9810
         X2              =   9810
         Y1              =   90
         Y2              =   4575
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   7935
         X2              =   7935
         Y1              =   105
         Y2              =   4575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   570
         X2              =   570
         Y1              =   105
         Y2              =   4575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sr."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   195
         TabIndex        =   26
         Top             =   165
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BAD3C9&
      Height          =   1635
      Left            =   150
      TabIndex        =   18
      Top             =   645
      Width           =   11385
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   8160
         MaxLength       =   25
         TabIndex        =   3
         Top             =   1110
         Width           =   1965
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   8160
         TabIndex        =   2
         Top             =   705
         Width           =   1965
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   8145
         TabIndex        =   1
         Top             =   315
         Width           =   1965
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2760
         TabIndex        =   6
         Top             =   1050
         Width           =   1965
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   5
         Top             =   675
         Width           =   1965
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   4
         Top             =   300
         Width           =   1965
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mechanic Attended"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6195
         TabIndex        =   24
         Top             =   1170
         Width           =   1725
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Present Job Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6195
         TabIndex        =   23
         Top             =   765
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Present Job No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6180
         TabIndex        =   22
         Top             =   315
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Job Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   810
         TabIndex        =   21
         Top             =   1095
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Job No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   795
         TabIndex        =   20
         Top             =   690
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Reg  No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   795
         TabIndex        =   19
         Top             =   300
         Width           =   1515
      End
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmRepeteJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const Reg_No As Integer = 0
Private Const PrvJob_No As Integer = 1
Private Const PrvJob_Date As Integer = 2
Private Const JobNo As Integer = 3
Private Const Job_Date As Integer = 4
Private Const Mechanic As Integer = 5
Private Const Component As Integer = 6
Private Const Batch_Code As Integer = 7
Private Const Make As Integer = 8
Private Const Ord_Placed As Integer = 9
Private Const Cust_Informed As Integer = 10
Private Const Imp_Date1 As Integer = 11
Private Const Imp_Date2 As Integer = 12
Private Const Imp_Date3 As Integer = 13
Private Const Imp_Date4 As Integer = 14
Private Const Imp_Date5 As Integer = 15
Private Const Imp_Date6 As Integer = 16
Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim MyIndex As Byte
Dim ADDFLAG$
Dim TAddMode As Boolean

Dim mRepName As String
Dim RsJob As ADODB.Recordset
Dim Master As ADODB.Recordset

Private Sub CmdPrint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
GSQL = "SELECT * from RepeatJob WHERE JobNo=" & txt(JobNo) & ""
'"
Select Case Index
    Case PScreen, PWindows
        If txt(JobNo) <> "" Then
            mRepName = IIf(OptPlain.Value = True, "RepeatJob", "RepeatJob")
        End If
        Call WindowsPrint(GSQL, Index)
        FrmPrn.Visible = False
        
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
If Index <> PSetUp And ADDFLAG <> "B" Then
    If ADDFLAG = "A" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
End If
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub WindowsPrint(mQRY As String, Index As Integer)
On Error GoTo ERRORHANDLER
Dim Rst As ADODB.Recordset
Dim mReportCount As Integer, i As Integer
 
Set Rst = GCn.Execute(mQRY)

CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
For i = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
        Case UCase("TITLE1")
            rpt.FormulaFields(i).TEXT = "'Repeat Job Analysis Form'"
    End Select
Next
     
rpt.Database.SetDataSource Rst
rpt.ReadRecords
Set Rst = Nothing
Select Case Index
    Case PWindows  'Printer
        For i = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
                Case UCase("comp_name")
                    rpt.FormulaFields(i).TEXT = "'" & PubComp_Name & "'"
                Case UCase("comp_add1")
                    rpt.FormulaFields(i).TEXT = "'" & PubComp_Add & "'"
                Case UCase("comp_add2")
                    rpt.FormulaFields(i).TEXT = "'" & PubComp_Add2 & "'"
                Case UCase("comp_city")
                    rpt.FormulaFields(i).TEXT = "'" & PubComp_City & "'"
            End Select
        Next
        rpt.PrintOut False
    Case PScreen  'screen
        Call Report_View(rpt, "Repeat Job Analysis Form", , True)
End Select

CmdPrint(PSetUp).Tag = ""
Set rpt = Nothing
Set RST1 = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Load()
    TopCtrl1.Tag = PubUParam: WinSetting Me
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select Job_DocId as SearchCode from RepeatJob where left(Job_DocID,1)='" & PubDivCode & "' Order by JobNo Desc", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsJob = New ADODB.Recordset
    With RsJob
        .CursorLocation = adUseClient
        .Open "select  J.DocId AS CODE, " & cCStr("J.Job_No") & " As FindJobNo,J.Job_No, HC.Model,HC.RegNo, HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name, J.DocId,J.Govt_YN, J.Job_Date, J.JobCloseDate,j.cardno, HC.Add1, HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc, City.CityName,Emp_Mast.Emp_Name,Emp_Mast.Emp_Code from (((job_card as J left Join Hiscard as HC on J.CardNo=HC.CardNo) left Join Service_Type as ST on J.Serv_Type=ST.Serv_Type) LEFT JOIN EMP_MAST ON J.RECBY_MECHANIC=EMP_MAST.EMP_CODE) Left Join City on HC.CityCode=City.CityCode  where left(j.DocId,1)='" & PubDivCode & "' Order by J.docID", GCn, adOpenDynamic, adLockOptimistic
    End With
    RsJob.Sort = "code"
    Set DGJob.DataSource = RsJob
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec

End Sub
Private Sub DGJob_Click()
Call History_Field
DGJob.Visible = False
End Sub
Private Sub UpdRequery()
    RsJob.Requery
End Sub
Private Sub TopCtrl1_eEdit()
Dim i As Integer
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    txt(JobNo).Enabled = False
    txt(Job_Date).Enabled = False
    txt(Reg_No).Enabled = False
    txt(PrvJob_No).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub
Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
    If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
                    
        GCn.Execute "Delete from RepeatJob  where job_Docid='" & txt(JobNo).Tag & "'"
    
        GCn.CommitTrans
        
        Master.Requery
        Call UpdRequery
        
        If Master.RecordCount > 0 Then
            Call MoveRec
        Else
            Call BlankText
        End If
        BUTTONS True, Me, Master, 0
    End If
    Exit Sub
eloop1:
    GCn.RollbackTrans
    MsgBox err.Description, vbCritical, " Deletion Message"
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
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("searchcode='" & MyValue & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub History_Field(Optional MakeBlank As Boolean)
If TopCtrl1.TopText2 = "Add" Then
    txt(JobNo).Tag = XNull(RsJob!Code)
    txt(JobNo).TEXT = XNull(RsJob!Job_No)
    txt(Job_Date).TEXT = RsJob!Job_Date
    txt(SrvType).TEXT = XNull(RsJob!Serv_Desc)
    txt(Reg_No).TEXT = XNull(RsJob!RegNo)
    txt(Mechanic).TEXT = XNull(RsJob!Emp_Name)
    txt(Mechanic).Tag = XNull(RsJob!Emp_Code)
    Call UpdLastJC
End If
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim i As Integer
    TopCtrl1.TopText2 = "Add"
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    Ini_Grid
    txt(JobNo).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub
Private Sub TopCtrl1_eCancel()
Dim i As Integer
On Error GoTo ErrorLoop
    If MsgBox("Cancel Entry ?", vbExclamation + vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        RsJob.Filter = ""
        Call BlankText
        Call Ini_Grid
        Call MoveRec
    Else
        Me.ActiveControl.SetFocus
    End If
    Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "select Job_DocId as SearchCode,JobNo,PrvJobNo,PrvJobDate from RepeatJob where left(Job_DocId,1)='" & PubDivCode & "' order by JobNo"
        Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub BlankText()
Dim i As Byte
    For i = 0 To txt.Count - 1
        txt(i).TEXT = ""
        If i <> GPDt Then txt(i).Tag = ""
    Next i
   
End Sub
Public Sub Disp_Text(Enb As Boolean)
Dim i As Integer
    'New Testing for Speed purpose
    ADDFLAG = left(TopCtrl1.TopText2, 1)
    'eof New Testing
    For i = 1 To txt.Count - 1
        txt(i).Enabled = Enb
    Next
    
    For i = 1 To txt.Count - 1
        txt(i).BackColor = CtrlBColOrg
        txt(i).ForeColor = CtrlFColOrg
    Next
    If TopCtrl1.TopText2 = "Edit" Then
        txt(JobNo).Enabled = False
        txt(Component).SetFocus
    End If
'    txt(Job_Date).Enabled = False
    txt(Reg_No).Enabled = False
'    txt(PrvJob_No).Enabled = False
'    txt(PrvJob_Date).Enabled = False
'
'    txt(Mechanic).Enabled = False
End Sub

Private Sub TopCtrl1_eRef()
    Call UpdRequery
End Sub

Private Sub TopCtrl1_eSave()
    Dim mTrans As Boolean
'    On Error GoTo err
    GCn.BeginTrans
    mTrans = True
    If ADDFLAG = "A" Then
        GSQL = "insert into RepeatJob(" _
            & "Job_DocId,JobNo,Job_Date,RegNo,Mech,PrvJobNo,PrvJobDate,Comp_Name,Batch_Code," _
            & "Make,Ord_Placed,Cust_Informed,Imp_date1,Imp_date2,Imp_date3,Imp_date4,Imp_date5,Imp_date6,U_Name, U_EntDt, U_AE) " _
            & " values(" _
            & "'" & txt(JobNo).Tag & "'," & Val(txt(JobNo)) & "," & ConvertDate(txt(Job_Date)) & ",'" & txt(Reg_No) & "','" & txt(Mechanic).TEXT & "'," _
            & "" & txt(PrvJob_No).TEXT & "," & ConvertDate(txt(PrvJob_Date)) & ",'" & txt(Component) & "','" & txt(Batch_Code) & "'" _
            & ",'" & txt(Make) & "'," & IIf(txt(Ord_Placed) = "Yes", 1, 0) & "," & IIf(txt(Cust_Informed) = "Yes", 1, 0) & "," & ConvertDate(txt(Imp_Date1)) & "," & ConvertDate(txt(Imp_Date2)) & "," & ConvertDate(txt(Imp_Date3)) & "," & ConvertDate(txt(Imp_Date4)) & "," & ConvertDate(txt(Imp_Date5)) & "," & ConvertDate(txt(Imp_Date6)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & ADDFLAG & "')"
    ElseIf ADDFLAG = "E" Then
        GSQL = "Update RepeatJob Set Job_DocId='" & txt(JobNo).Tag & "',RegNo = '" & txt(Reg_No) & "',JobNo = " & txt(JobNo) & ",Job_Date=" & ConvertDate(txt(Job_Date)) & ",Mech='" & txt(Mechanic) & "',PrvJobNo='" & txt(PrvJob_No) & "',PrvJobDate='" & txt(PrvJob_Date) & "',Comp_Name='" & txt(Component) & "',Batch_Code='" & txt(Batch_Code) & "'," _
            & "Make='" & txt(Make) & "',Ord_Placed=" & IIf(txt(Ord_Placed) = "Yes", 1, 0) & ",Cust_Informed=" & IIf(txt(Cust_Informed) = "Yes", 1, 0) & ",Imp_date1=" & ConvertDate(txt(Imp_Date1)) & ",Imp_date2=" & ConvertDate(txt(Imp_Date2)) & ",Imp_date3=" & ConvertDate(txt(Imp_Date3)) & ",Imp_date4=" & ConvertDate(txt(Imp_Date4)) & ",Imp_date5=" & ConvertDate(txt(Imp_Date5)) & ",Imp_date6=" & ConvertDate(txt(Imp_Date6)) & "" _
            & ",U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='" & ADDFLAG & "' where Job_DocId='" & txt(JobNo).Tag & "'"
    End If
    GCn.Execute GSQL
    GCn.CommitTrans
    mTrans = False
    If ADDFLAG = "A" Then TopCtrl1_ePrn
    Disp_Text SETS("INI", Me, Master)
    Master.Requery
    Call UpdRequery
'    If txt(JobNo) <> "" Then
'        Call MoveRec
'    End If
    Exit Sub

err:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub
Private Sub TopCtrl1_ePrn()
FrmPrn.top = (Me.height - FrmPrn.height) / 2
FrmPrn.left = (Me.width - FrmPrn.width) / 2
FrmPrn.Visible = True
FrmPrn.ZOrder 0
OptPlain.Value = True
LblPrinter.CAPTION = Printer.DeviceName
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True


End Sub

Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus txt(Index)
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case JobNo
            RsJob.Filter = ("Job_Date<=" & ConvertDate(PubLoginDate) & " and Jobclosedate = Null")
            If RsJob.RecordCount <= 0 Then Exit Sub
            DGridColSwap DGJob, 0
            RsJob.Sort = "JOB_NO"
            If txt(Index).Tag <> "" And txt(Index).Tag <> RsJob!Code Then
                RsJob.FIND ("JOB_NO='" & txt(Index).TEXT & "'")
            End If
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case JobNo
            DGridTxtKeyDown DGJob, txt, Index, RsJob, KeyCode, False, 1
            Call History_Field
    End Select
        If DGJob.Visible = False Then
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And ((ADDFLAG = "A") Or (ADDFLAG = "E")) Then
                Select Case Index
                    Case Job_Date
                        txt(Job_Date) = RetDate(txt(Job_Date))
                    Case PrvJob_Date
                        txt(PrvJob_Date) = RetDate(txt(PrvJob_Date))
                    Case Imp_Date1
                        txt(Imp_Date1) = RetDate(txt(Imp_Date1))
                    Case Imp_Date2
                        txt(Imp_Date2) = RetDate(txt(Imp_Date2))
                    Case Imp_Date3
                        txt(Imp_Date3) = RetDate(txt(Imp_Date3))
                    Case Imp_Date4
                        txt(Imp_Date4) = RetDate(txt(Imp_Date4))
                    Case Imp_Date5
                        txt(Imp_Date5) = RetDate(txt(Imp_Date5))
                    Case Imp_Date6
                        txt(Imp_Date6) = RetDate(txt(Imp_Date6))
               End Select
                
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And ((ADDFLAG = "A" And Index = Imp_Date6) Or (ADDFLAG = "E" And Index = Imp_Date6)) Then
                If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
            End If
                Ctrl_DownKeyDown KeyCode, Shift
                
            End If
           
            
            ' KEY UP
            If ADDFLAG = "A" Then
                If Index <> GPNo Then
                    If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
                End If
            ElseIf ADDFLAG = "E" Then
                If Index <> GPDt Then
                    If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
                End If
            End If
        End If

End Sub
Private Sub Ini_Grid()
    DGJob.left = Frame2.left: DGJob.width = Frame2.width: DGJob.top = Frame2.top: DGJob.height = Frame2.height
End Sub
Private Sub Grid_Hide()
    If DGJob.Visible = True Then DGJob.Visible = False
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
    Select Case Index
        Case JobNo
            DGridTxtKeyPress txt, Index, RsJob, KeyAscii, "Findjobno"
        Case Ord_Placed
            If KeyAscii = Asc("Y") Or KeyAscii = Asc("y") Then
                txt(Ord_Placed) = "Yes"
            ElseIf KeyAscii = Asc("N") Or KeyAscii = Asc("n") Then
                txt(Ord_Placed) = "No"
            End If
            KeyAscii = 0
        Case Cust_Informed
            If KeyAscii = Asc("Y") Or KeyAscii = Asc("y") Then
                txt(Cust_Informed) = "Yes"
            ElseIf KeyAscii = Asc("N") Or KeyAscii = Asc("n") Then
                txt(Cust_Informed) = "No"
            End If
            KeyAscii = 0
    End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate txt(Index)
End Sub
Private Sub UpdLastJC()
    Dim RsTemp As ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    RsTemp.CursorLocation = adUseClient
    RsTemp.Open "SELECT Top 1 DocId,JOB_NO,JOB_DATE,AtKMsHrs,Srv.Serv_SrlNo,Srv.Serv_Type,Srv.SERV_DESC AS SrvDesc,EMP_MAST.EMP_NAME AS MECH_NAME " & _
            " FROM ((JOB_CARD LEFT JOIN Service_Type Srv ON JOB_CARD.SERV_TYPE=Srv.SERV_TYPE) " & _
            " LEFT JOIN EMP_MAST ON JOB_CARD.RECBY_MECHANIC=EMP_MAST.EMP_CODE) " & _
            " LEFT JOIN HisCard ON JOB_CARD.CardNo=HisCard.CardNo " & _
            " WHERE HisCard.RegNo='" & txt(Reg_No).TEXT & _
            "' and Job_Date< " & ConvertDate(txt(Job_Date)) & _
            " ORDER BY JOB_DATE Desc ", GCn, adOpenStatic, adLockReadOnly
    If RsTemp.RecordCount > 0 Then
        txt(PrvJob_No).TEXT = RsTemp!Job_No
        txt(PrvJob_No).Tag = RsTemp!DocID
        txt(PrvJob_Date).TEXT = RsTemp!Job_Date
        
    Else
        txt(PrvJob_No).TEXT = ""
        txt(PrvJob_Date).TEXT = ""
    End If
    Set RsTemp = Nothing
End Sub
Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        FrmPrn.Visible = False
        If Index <> PSetUp And ADDFLAG <> "B" Then
            If ADDFLAG = "A" Then TopCtrl1_eAdd: Exit Sub
            Disp_Text SETS("INI", Me, Master)
            Call MoveRec
        End If
    End If
End Sub
Private Sub MoveRec()
Dim Master1 As Recordset, rs1 As Recordset
Dim i As Integer
On Error GoTo error1
    BlankText
    If Master.RecordCount > 0 Then
    '   Master.Open "select GP.GatePassNo as SearchCode,GP.*, Emp_Mast.Emp_Name as MechName from Job_GatePass as GP Left Join Emp_Mast on GP.Mech_Code=Emp_Mast.Emp_Code  where left(Job_DocId,1)='" & PubDivCode & "' order by gp.GatePassNo", GCn, adOpenDynamic, adLockOptimistic
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "Select RepeatJob.* from RepeatJob Where RepeatJob.Job_DocID='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
            
        txt(JobNo).Tag = VNull(Master1!job_docid)
        txt(JobNo).TEXT = VNull(Master1!JobNo)
        txt(Job_Date).TEXT = XNull(Master1!Job_Date)
        txt(Reg_No).TEXT = XNull(Master1!RegNo)
        txt(Mechanic).TEXT = XNull(Master1!Mech)
        txt(PrvJob_No).TEXT = XNull(Trim(Right(Master1!PrvJobNo, 8)))
        txt(PrvJob_Date).TEXT = XNull(Master1!PrvJobDate)
        txt(Component).TEXT = XNull(Master1!Comp_Name)
        
        txt(Batch_Code).TEXT = XNull(Master1!Batch_Code)
        txt(Make).TEXT = XNull(Master1!Make)
        txt(Ord_Placed).TEXT = IIf(Master1!Ord_Placed = 1, "Yes", "No")
        txt(Cust_Informed).TEXT = IIf(Master1!Cust_Informed = 1, "Yes", "No")
        txt(Imp_Date1).TEXT = XNull(Master1!Imp_Date1)
        txt(Imp_Date2).TEXT = XNull(Master1!Imp_Date2)
        txt(Imp_Date3).TEXT = XNull(Master1!Imp_Date3)
        txt(Imp_Date4).TEXT = XNull(Master1!Imp_Date4)
        txt(Imp_Date5).TEXT = XNull(Master1!Imp_Date5)
        txt(Imp_Date6).TEXT = XNull(Master1!Imp_Date6)
        
    Grid_Hide
    
    
    Set Master1 = Nothing
    Exit Sub
    End If
error1:
    CheckError
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
If Index = JobNo Then
    If GCn.Execute("Select Job_DocID from RepeatJob where Job_DocId='" & txt(JobNo).Tag & "'").RecordCount > 0 Then
        MsgBox "Record for this Job Already exists."
        BlankText
        txt(JobNo).SetFocus
        Cancel = True
    End If
End If
End Sub
