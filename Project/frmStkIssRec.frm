VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmStkIssRec 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Stock Adjustments Issue/ Receipt"
   ClientHeight    =   7770
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
   ScaleHeight     =   7770
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2745
      Top             =   1620
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   74
      Top             =   2280
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton CmdVatStock 
      Caption         =   "Convert to VAT Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   73
      Top             =   0
      Width           =   2175
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   4980
      TabIndex        =   22
      Top             =   645
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   75
         TabIndex        =   23
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
      Height          =   255
      Index           =   6
      Left            =   1935
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1200
      Width           =   2100
   End
   Begin MSDataGridLib.DataGrid DGGodown 
      Height          =   3330
      Left            =   -5490
      Negotiate       =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   -1050
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
   Begin VB.Frame FrmDetail 
      BackColor       =   &H00CAF1FD&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   2205
      Left            =   11295
      TabIndex        =   39
      Top             =   -1035
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         Index           =   2
         Left            =   75
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2865
      Left            =   -4725
      Negotiate       =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   -1200
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
      Height          =   255
      Index           =   4
      Left            =   1935
      MaxLength       =   40
      TabIndex        =   5
      Top             =   930
      Width           =   4830
   End
   Begin VB.TextBox Txt 
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
      Height          =   255
      Index           =   5
      Left            =   1110
      MaxLength       =   30
      TabIndex        =   9
      Top             =   6705
      Width           =   4905
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
      Height          =   255
      Index           =   0
      Left            =   9405
      TabIndex        =   1
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
      Left            =   2550
      TabIndex        =   7
      Top             =   3195
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9765
      TabIndex        =   2
      Top             =   1080
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
      Height          =   255
      Index           =   2
      Left            =   9765
      MaxLength       =   11
      TabIndex        =   3
      Top             =   1350
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
      Height          =   255
      Index           =   3
      Left            =   10425
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1620
      Width           =   900
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   4080
      Left            =   15
      TabIndex        =   8
      Top             =   2595
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   7197
      _Version        =   393216
      BackColor       =   14940925
      Cols            =   25
      BackColorFixed  =   15259902
      ForeColorFixed  =   8388736
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   16761087
      GridColorFixed  =   8421504
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
      _Band(0).Cols   =   25
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   2670
      Left            =   120
      Negotiate       =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   7110
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Converting Taxpaid Stock to Taxable Stock"
      Height          =   240
      Left            =   15
      TabIndex        =   75
      Top             =   1995
      Visible         =   0   'False
      Width           =   3975
      WordWrap        =   -1  'True
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
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   2
      Left            =   1815
      TabIndex        =   72
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Adj."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   4
      Left            =   780
      TabIndex        =   71
      Top             =   1200
      Width           =   930
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
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   0
      Left            =   8745
      TabIndex        =   37
      Top             =   6915
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Goods Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   0
      Left            =   7095
      TabIndex        =   36
      Top             =   6915
      Width           =   1530
   End
   Begin VB.Label LblGoodsValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   8970
      TabIndex        =   35
      Top             =   6945
      Width           =   360
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
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   6
      Left            =   1020
      TabIndex        =   33
      Top             =   6705
      Width           =   180
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   27
      Left            =   210
      TabIndex        =   32
      Top             =   6705
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   3
      Left            =   780
      TabIndex        =   31
      Top             =   930
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
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   1
      Left            =   1815
      TabIndex        =   30
      Top             =   930
      Width           =   45
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFFF&
      Height          =   1485
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   450
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
      Height          =   255
      Left            =   9765
      TabIndex        =   29
      Top             =   1620
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   25
      Left            =   9285
      TabIndex        =   27
      Top             =   525
      Width           =   90
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
      Height          =   255
      Index           =   31
      Left            =   8610
      TabIndex        =   26
      Top             =   525
      Width           =   585
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
      ForeColor       =   &H00FF00FF&
      Height          =   270
      Left            =   10350
      TabIndex        =   25
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
      ForeColor       =   &H00FF00FF&
      Height          =   270
      Left            =   8610
      TabIndex        =   24
      Top             =   825
      Width           =   660
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
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   1
      Left            =   8610
      TabIndex        =   21
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total No. of Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   7
      Left            =   7095
      TabIndex        =   20
      Top             =   6705
      Width           =   1440
   End
   Begin VB.Label LblQty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   10965
      TabIndex        =   19
      Top             =   6705
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   8970
      TabIndex        =   18
      Top             =   6705
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   25
      Left            =   9525
      TabIndex        =   17
      Top             =   6705
      Width           =   1110
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
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   28
      Left            =   8745
      TabIndex        =   16
      Top             =   6705
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
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   27
      Left            =   10755
      TabIndex        =   15
      Top             =   6705
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
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   90
      Left            =   9630
      TabIndex        =   14
      Top             =   1080
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
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   92
      Left            =   9630
      TabIndex        =   13
      Top             =   1620
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
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   93
      Left            =   9630
      TabIndex        =   12
      Top             =   1350
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   0
      Left            =   8610
      TabIndex        =   11
      Top             =   1350
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
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   2
      Left            =   8610
      TabIndex        =   10
      Top             =   1620
      Width           =   810
   End
End
Attribute VB_Name = "frmStkIssRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim RsParty As ADODB.Recordset
Dim RsGodown As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim mVType As String, mVPrefix As String
Dim mSearchCode As String
Dim ExitCtrl As Boolean
Dim mPartyType As Byte         'Used to Detect Part Rate in GetRate Function
Dim OldDocType As String
Dim mCheckNegetiveStockSiteWise As Boolean

'grid color scheme
Private Const CellBackColLeave As String = &HE3FAFD
'Private Const CellForeColLeave As String = &HFF00FF
'Private Const CellBackColEnter As String = &HCAF1FD
Private Const GridBackColorBkg As String = &HD7C6C8    ' me.backColor=&HB9D8EE
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const AdjIssVType As String = "SYIAD"
Private Const AdjRecVType As String = "SXRAD"

' Under observation
Dim VoucherEditFlag As Boolean                  ' Used for whether we can edit voucher no or not
' End Under observation
Dim ListArray As Variant
Dim mListItem As ListItem

Private Const DocID As Byte = 0                 ' Doc.ID
Private Const DocType As Byte = 1               ' Document Type
Private Const VDate As Byte = 2                 ' Date
Private Const SerialNo As Byte = 3              ' Serial No.
Private Const Party As Byte = 4                 ' A/c Name
Private Const Remark As Byte = 5                ' Remark
Private Const AdjType As Byte = 6                ' Adjustment Types Breakage,Loan etc

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_PNo As Byte = 1               ' Part No
Private Const Col_Unit As Byte = 2              ' Unit
Private Const Col_MRP As Byte = 3               ' MRP Yes/No
Private Const Col_Taxable As Byte = 4           ' Taxable Yes/No
Private Const Col_Qty As Byte = 5               ' Qty
Private Const Col_Rate As Byte = 6              ' Rate
Private Const Col_MRPRate As Byte = 7           ' MRP Rate
Private Const Col_Amt As Byte = 8               ' Amt
Private Const Col_DiscPer As Byte = 9           ' Disc. %
Private Const Col_GodownCode As Byte = 10       ' Godown Code
Private Const Col_Godown As Byte = 11           ' Godown
Private Const Col_PName As Byte = 12            ' Part Name
Private Const Col_LName As Byte = 13            ' Local Name
Private Const Col_MRPStkTP As Byte = 14         ' MRP Stk TP 'Current Stock Qty
Private Const Col_MRPStkTB As Byte = 15         ' MRP Stk TB
Private Const Col_TBStk As Byte = 16            ' Taxbale Qty
Private Const Col_TPStk As Byte = 17            ' Tax Paid Qty
Private Const Col_TBRate As Byte = 18           ' Taxbale Rate
Private Const Col_TPRate As Byte = 19           ' Tax Paid Rate
Private Const Col_Bin As Byte = 20              ' Bin
Private Const Col_LastRate As Byte = 21         ' Last Purchase Rate
Private Const Col_HPRate As Byte = 22           ' High Purchase Rate
Private Const Col_LPRate As Byte = 23           ' Low Purchase Rate
Private Const Col_PartGrade As Byte = 24        ' Part Grade (Used for Oil Item)
Private Const Col_EffectDate As Byte = 25       ' MRP Effective Date/TB Effective Date
Private Const Col_SrlNo As Byte = 26            ' SP_Stock SrlNo (DocID+SrlNo)


Dim mConverting As Boolean

Private Sub Disp_Text(Enb As Boolean)
    txt(DocType).Enabled = Enb
    txt(VDate).Enabled = Enb
    txt(SerialNo).Enabled = Enb
    txt(Party).Enabled = Enb
    txt(AdjType).Enabled = Enb
    txt(Remark).Enabled = Enb
    
    If Not StrCmp(left(PubComp_Name, 3), "JMK") Then CmdVatStock.Visible = False
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("SearchCode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select Distinct S.DocID As SearchCode,S.DocID,S.V_Date,S.V_Type " _
            & "From SP_Stock S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' and S.V_Type In ('" & AdjIssVType & "','" & AdjRecVType & "') And  S.DocID = '" & MyValue & "' " _
            & "Order by S.V_Date desc,S.V_Type")
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
        txt(I).TEXT = ""
    Next I
    txt(DocID).Tag = ""
    LblDiv.CAPTION = "Division : "
    LblSite.CAPTION = "Site Code : "
    LblVPrefix.CAPTION = ""
    LblIVal.CAPTION = ""
    LblQty.CAPTION = ""
    LblGoodsValue.CAPTION = ""

    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End Sub

'* Used for intialize grid columns
Private Sub Grid_Ini()
    With FGrid
        .left = Me.left '+ 60
        .width = Me.width - 90
        .top = 2595
        .BackColor = CellBackColLeave
        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 27

        .TextMatrix(0, Col_SrNo) = "S.No"
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 450

        .TextMatrix(0, Col_PNo) = "Part No"
        .ColAlignment(Col_PNo) = flexAlignLeftCenter
        .ColWidth(Col_PNo) = 1500

        .TextMatrix(0, Col_PName) = "Part Name"
        .ColAlignment(Col_PName) = flexAlignLeftCenter
        .ColWidth(Col_PName) = 2500

        .TextMatrix(0, Col_LName) = "Local Name"
        .ColAlignment(Col_LName) = flexAlignLeftCenter
        .ColWidth(Col_LName) = 2000

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
        .ColWidth(Col_MRPRate) = 0

        .TextMatrix(0, Col_Amt) = "Amount"
        .ColAlignmentFixed(Col_Amt) = flexAlignRightCenter
        .ColWidth(Col_Amt) = 1065

        .TextMatrix(0, Col_DiscPer) = "Disc Per"    'added to maintain std.
        .ColWidth(Col_DiscPer) = 0

        .TextMatrix(0, Col_GodownCode) = "Godown Code"
        .ColAlignment(Col_GodownCode) = flexAlignLeftCenter
        .ColWidth(Col_GodownCode) = 0

        .TextMatrix(0, Col_Godown) = "Godown"
        .ColAlignment(Col_Godown) = flexAlignLeftCenter
        .ColWidth(Col_Godown) = 1200

        .TextMatrix(0, Col_MRPStkTP) = "MRP Stock TP"
        .ColWidth(Col_MRPStkTP) = 0

        .TextMatrix(0, Col_MRPStkTB) = "MRP Stock TB"
        .ColWidth(Col_MRPStkTB) = 0

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

        .TextMatrix(0, Col_EffectDate) = "Rate Effective Date"
        .ColWidth(Col_EffectDate) = 0

        .TextMatrix(0, Col_SrlNo) = "Stock Srl No"
        .ColWidth(Col_SrlNo) = 0
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    
    DGPart.width = FGrid.width: DGPart.left = FGrid.left: DGPart.top = mTopScale:  DGPart.height = FGrid.top - mTopScale
    DGGodown.left = FGrid.left: DGGodown.top = DGPart.top: DGGodown.height = DGPart.height
    FrmDetail.width = 6285: FrmDetail.left = 5595: FrmDetail.top = mTopScale: FrmDetail.height = 2130
    DGParty.left = Me.width - (DGParty.width + mRtScale): DGParty.top = mTopScale
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
    If DGParty.Visible = True Then DGParty.Visible = False
    If DGGodown.Visible = True Then DGGodown.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGPart.Visible = True Then DGPart.Visible = False
End Sub

Private Sub MoveRec()
Dim Rst As ADODB.Recordset, I As Integer
On Error GoTo ELoop
    If Master.RecordCount > 0 Then
        FGrid.Redraw = False
        FGrid.Rows = 1
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "Select P.Part_Name ,P.Local_Name ,P.Unit ,P.MRP ,P.TB_SRate ,P.TP_SRate ," _
            & "P.MRP_Effect_Dt ,P.TB_Effect_Dt ," _
            & "P.Cur_MRP_TBStk, P.Cur_MRP_TPStk,P.Cur_TB_Stk ,P.Cur_TP_Stk," _
            & "P.Bin_Loca ,P.High_Pur_Rate ,P.Low_Pur_Rate ," _
            & "SubGroup.Name As AcName,Godown.God_Name,SP_Stock.* " _
            & "From ((SP_Stock Left Join Part P On SP_Stock.Part_No=P.Part_No and P.Div_Code = left(SP_Stock.Docid,1)) " _
            & "Left Join SubGroup on SP_Stock.Party_Code=SubGroup.SubCode) " _
            & "Left Join Godown on SP_Stock.Godown=Godown.God_Code " _
            & "Where SP_Stock.DocID='" & Master!DocID & "' Order By SP_Stock.Srl_No", GCn, adOpenStatic, adLockReadOnly
        
        If Rst.RecordCount > 0 Then
            txt(DocID).TEXT = Master!DocID
            mSearchCode = txt(DocID)
            LblDiv.CAPTION = "Division : " & left(Rst!DocID, 1)
            LblSite.CAPTION = "Site Code : " & Rst!Site_Code
            mVType = Rst!V_Type
            If mVType = AdjIssVType Then
                txt(DocType).TEXT = "Issue"
                OldDocType = "Issue"
            ElseIf mVType = AdjRecVType Then
                txt(DocType).TEXT = "Receipt"
                OldDocType = "Receipt"
            End If
            
            If Rst!Purpose = "A" Then
                txt(AdjType).TEXT = "Assemble"
            ElseIf Rst!Purpose = "B" Then
                txt(AdjType).TEXT = "Breakage"
            ElseIf Rst!Purpose = "D" Then
                txt(AdjType).TEXT = "Dismental"
            ElseIf Rst!Purpose = "L" Then
                txt(AdjType).TEXT = "Loan"
            ElseIf Rst!Purpose = "O" Then
                txt(AdjType).TEXT = "Others"
            Else
                txt(AdjType).TEXT = ""
            End If

            
            txt(VDate).TEXT = Rst!V_Date
            LblVPrefix.CAPTION = mID(Rst!DocID, 9, 5)
            txt(SerialNo).TEXT = Rst!V_NO
            txt(Party).Tag = Rst!Party_code
            txt(Party).TEXT = IIf(IsNull(Rst!AcName), "", Rst!AcName)
            txt(Remark).TEXT = IIf(IsNull(Rst!Remark), "", Rst!Remark)
            
            I = 1
            Do Until Rst.EOF
                FGrid.AddItem ""
                With FGrid
                    .TextMatrix(I, Col_SrNo) = I
                    .TextMatrix(I, Col_PNo) = Rst!Part_No
                    .TextMatrix(I, Col_Unit) = IIf(IsNull(Rst!Unit), "", Rst!Unit)
                    .TextMatrix(I, Col_MRP) = IIf(Rst!MRP_YN = 0, "No", "Yes")
                    .TextMatrix(I, Col_Taxable) = IIf(Rst!Tax_YN = 0, "No", "Yes")
                    If txt(DocType).TEXT = "Issue" Then
                        .TextMatrix(I, Col_Qty) = IIf(Rst!Qty_Iss = 0, "", Format(Rst!Qty_Iss, "0.000"))
                    Else
                        .TextMatrix(I, Col_Qty) = IIf(Rst!Qty_Rec = 0, "", Format(Rst!Qty_Rec, "0.000"))
                    End If
                    .TextMatrix(I, Col_Rate) = IIf(Rst!Rate = 0, "", Format(Rst!Rate, "0.000"))
                    .TextMatrix(I, Col_MRPRate) = Format(Rst!MRP, "0.00")
                    .TextMatrix(I, Col_Amt) = Format(Rst!Amount, "0.00")
                    .TextMatrix(I, Col_GodownCode) = Rst!Godown
                    .TextMatrix(I, Col_Godown) = IIf(IsNull(Rst!God_Name), "", Rst!God_Name)
                    .TextMatrix(I, Col_PName) = IIf(IsNull(Rst!Part_Name), "", Rst!Part_Name)
                    .TextMatrix(I, Col_LName) = IIf(IsNull(Rst!Local_Name), "", Rst!Local_Name)
                    .TextMatrix(I, Col_MRPStkTP) = IIf(IsNull(Rst!Cur_MRP_TPStk), 0, Rst!Cur_MRP_TPStk)
                    .TextMatrix(I, Col_MRPStkTB) = IIf(IsNull(Rst!Cur_MRP_TbStk), 0, Rst!Cur_MRP_TbStk)
                    .TextMatrix(I, Col_TBStk) = IIf(IsNull(Rst!Cur_TB_STk), 0, Rst!Cur_TB_STk)
                    .TextMatrix(I, Col_TPStk) = IIf(IsNull(Rst!Cur_TP_Stk), 0, Rst!Cur_TP_Stk)
                    .TextMatrix(I, Col_TBRate) = IIf(IsNull(Rst!TB_SRate), 0, Rst!TB_SRate)
                    .TextMatrix(I, Col_TPRate) = IIf(IsNull(Rst!TP_SRate), 0, Rst!TP_SRate)
                    .TextMatrix(I, Col_Bin) = IIf(IsNull(Rst!Bin_Loca), "", Rst!Bin_Loca)
                    .TextMatrix(I, Col_LastRate) = ""
                    .TextMatrix(I, Col_HPRate) = IIf(IsNull(Rst!high_pur_rate), 0, Rst!high_pur_rate)
                    .TextMatrix(I, Col_LPRate) = IIf(IsNull(Rst!low_pur_rate), 0, Rst!low_pur_rate)
                    .TextMatrix(I, Col_EffectDate) = Format(IIf(Rst!MRP_YN = 1, Rst!MRP_Effect_Dt, Rst!TB_Effect_Dt), "dd/MMM/yyyy")
                    .TextMatrix(I, Col_SrlNo) = Rst!Srl_No
                End With
                Rst.MoveNext
                I = I + 1
            Loop
            FGrid.FixedRows = 1
            CountItem
        Else
            FGrid.AddItem FGrid.Rows
            FGrid.FixedRows = 1
        End If
    Else
        BlankText
    End If
    FGrid.Redraw = True
    FrmDetail.Visible = False
    Grid_Hide
Set Rst = Nothing
'TopCtrl1.tPrn = False
Exit Sub
ELoop:
   FGrid.Redraw = True
   CheckError
End Sub

' Used For Checking Duplicate Items in the Grid
Private Function ChkDuplicate() As Boolean
Dim I As Integer, X As String, Y As String
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
    X = UCase(CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col1))) + CStr(Trim(FGrid.TextMatrix(FGrid.Row, Col2))) + CStr(Trim(TxtGrid(0).TEXT)))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        Y = UCase(CStr(Trim(FGrid.TextMatrix(I, Col1))) + CStr(Trim(FGrid.TextMatrix(I, Col2))) + CStr(Trim(FGrid.TextMatrix(I, Col3))))
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
Dim DiscPer As Integer, PartRate As Double
Dim rstDisc As ADODB.Recordset
Select Case FGrid.Col
    Case Col_PNo, Col_PName, Col_LName
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        TxtGridValid_PNo
        'Get Disc per for showing NDP Price
        Set rstDisc = GCn.Execute("Select PD.PurcDisc_Per From Part_DiscFactor PD Left Join Part on Part.Disc_Factor=PD.DiscFac_Catg where Part.Part_No='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "' and Part.Div_Code='" & PubDivCode & "'")
        If rstDisc.RecordCount > 0 Then
            DiscPer = rstDisc.Fields(0).Value
        Else
            DiscPer = 0
        End If
        PartRate = Val(FGrid.TextMatrix(FGrid.Row, Col_Rate))
        FGrid.TextMatrix(FGrid.Row, Col_Rate) = Round(PartRate - ((PartRate * DiscPer) / 100), 2)
    Case Col_Taxable, Col_MRP
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
'            If TopCtrl1.TopText2 = "Add" Or _
                TopCtrl1.TopText2 = "Edit" And Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(mPartyType, FGrid, CDate(txt(VDate).TEXT), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
                'Get Disc per for showing NDP Price
                Set rstDisc = GCn.Execute("Select PD.PurcDisc_Per From Part_DiscFactor PD Left Join Part on Part.Disc_Factor=PD.DiscFac_Catg where Part.Part_No='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "' and Part.Div_Code='" & PubDivCode & "'")
                If rstDisc.RecordCount > 0 Then
                    DiscPer = rstDisc.Fields(0).Value
                Else
                    DiscPer = 0
                End If
                PartRate = Val(FGrid.TextMatrix(FGrid.Row, Col_Rate))
                FGrid.TextMatrix(FGrid.Row, Col_Rate) = Round(PartRate - ((PartRate * DiscPer) / 100), 2)
'           End If
        End If
        Amt_Cal
    Case Col_Qty
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.000")
        If txt(DocType).TEXT = "Issue" Then
            If CheckSprStock(FGrid, FGrid.Row, Col_MRP, Col_Taxable, Col_Qty, Col_MRPStkTB, Col_MRPStkTP, Col_TBStk, Col_TPStk) = False Then TxtGrid(0).SetFocus: TxtGridLeave = False: Exit Function
        End If
        If RsGodown.RecordCount > 0 Or Trim(FGrid.TextMatrix(FGrid.Row, Col_Godown)) = "" Then
            RsGodown.MoveFirst
            RsGodown.FIND "Code ='" & PubSprCounterGodown & "'"
            FGrid.TextMatrix(FGrid.Row, Col_GodownCode) = RsGodown!Code
            FGrid.TextMatrix(FGrid.Row, Col_Godown) = RsGodown!Name
        End If
        Amt_Cal
    Case Col_Rate
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
        Amt_Cal
    Case Col_Godown
        If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or TxtGrid(0).TEXT = "" Then
            FGrid.TextMatrix(FGrid.Row, Col_GodownCode) = ""
            FGrid.TextMatrix(FGrid.Row, Col_Godown) = ""
        Else
            FGrid.TextMatrix(FGrid.Row, Col_GodownCode) = RsGodown!Code
            FGrid.TextMatrix(FGrid.Row, Col_Godown) = RsGodown!Name
        End If
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
Dim I As Integer, TotQty As Double, TotGoodsVal As Double
    FGrid.TextMatrix(FGrid.Row, Col_Amt) = Format((Val(FGrid.TextMatrix(FGrid.Row, Col_Rate)) * Val(FGrid.TextMatrix(FGrid.Row, Col_Qty))), "0.00")
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            TotQty = TotQty + Val(FGrid.TextMatrix(I, Col_Qty))
            TotGoodsVal = TotGoodsVal + Val(FGrid.TextMatrix(I, Col_Amt))
        End If
    Next I
    LblQty.CAPTION = Format(TotQty, "0.000")
    LblGoodsValue.CAPTION = Format(TotGoodsVal, "0.00")
End Sub

Private Sub CountItem()
Dim I As Integer, TotItems As Integer, TotQty As Double, TotGoodsVal As Double
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            TotQty = TotQty + Val(FGrid.TextMatrix(I, Col_Qty))
            TotGoodsVal = TotGoodsVal + Val(FGrid.TextMatrix(I, Col_Amt))
            TotItems = TotItems + 1
        End If
    Next I
    LblIVal.CAPTION = Format(TotItems, "0")
    LblQty.CAPTION = Format(TotQty, "0.000")
    LblGoodsValue.CAPTION = Format(TotGoodsVal, "0.00")
End Sub

Private Sub CmdVatStock_Click()
    On Error GoTo DispErr
    
    
    
    Dim Rst As ADODB.Recordset, RST1 As ADODB.Recordset
    Dim mQry As String, Condstr As String, SourcePath As String
    Dim GCN1 As ADODB.Connection
    Dim RsPart As ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim mCost As Double
    Dim TotQty As Double
    Dim mQty As Double
    Dim mReqQty As Double
    Dim mFifoCost As Double
    Dim mChkQty As Double
    Dim mTrans As Boolean



    Dim RsTpStk As ADODB.Recordset
    Dim RstRep As ADODB.Recordset
    Dim mVNoIssue       As Long
    Dim mVNoReceive     As Long
    Dim I               As Long
    Dim j               As Long
    Dim mSrlIssue       As Long
    Dim mSrlReceive     As Long
    Dim mTransDate As String
    
    Dim mCount  As Long
    Dim mDocIdIssue$, mDocIdReceive$, mVTypeIssue$, mVPrefixIssue$, mVDate$, mVTypeReceive$, mVPrefixReceive$
    
                        
    If MsgBox("It will Convert All TaxPaid Stock to Taxable Stock. Do u want to Continue? ", vbYesNo) = vbNo Then Exit Sub
    If MsgBox("Changes Can't be Undo. Do u want to Continue? ", vbYesNo) = vbNo Then Exit Sub
    mTransDate = InputBox("Enter Date (Stock will be calculated uptoSpecified date", "Conversion Date", PubLoginDate)
    If mTransDate = "" Then MsgBox "Transaction Date is Required. Porcess is Terminated.": Exit Sub
    mTransDate = MakeDate(mTransDate)


    
    GCn.Execute ("Drop Table SP_StockNew")
    GCn.Execute ("Select * into SP_StockNew from SP_Stock")

    Set RstRep = New ADODB.Recordset
    With RstRep
        .Fields.Append "Div_Code", adChar, 1, adFldIsNullable
        .Fields.Append "Part_No", adVarChar, 21, adFldIsNullable
        .Fields.Append "Part_Name", adVarChar, 50, adFldIsNullable
        .Fields.Append "Qty", adDouble, 12, adFldIsNullable
        .Fields.Append "Tax_Yn", adChar, 1, adFldIsNullable
        .Fields.Append "Mrp_Yn", adChar, 1, adFldIsNullable
        .Fields.Append "Rate", adDouble, 12, adFldIsNullable
        .Fields.Append "Amount", adDouble, 12, adFldIsNullable
        .Fields.Append "Days", adDouble, 12, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .Open
    End With
        
    mQry = "SELECT Left(S.DocId,1) As Div_Code, Trim(S.Part_No) As Part_No, " & _
           "sum(S.Qty_Rec)-Sum(S.Qty_Iss)+Sum(S.Qty_Ret) As mQty, Tax_Yn, 1 as Mrp_Yn, " & _
           "Max(Rate) As mRate, Max(Amount) As Amount, Max(S.V_Rate) As V_Rate " & _
           "From Sp_StockNew S " & _
           "WHERE iif(S.v_Date=" & ConvertDate(Format(DateAdd("D", -1, PubStartDate), "dd/MMM/yyyy")) & ",S.V_Type='SXAO',IIF(S.V_Date>= " & ConvertDate(PubStartDate) & " And S.V_Date<=" & ConvertDate(mTransDate) & ",S.V_Type<>'SXAO')) and Left(S.DocId,1)='" & PubDivCode & "' And S.Tax_Yn=0 " & Condstr & _
           "Group By Left(S.DocId,1), Part_No, Tax_Yn " & _
           "Having (sum(Qty_Rec)-Sum(Qty_Iss)+Sum(Qty_Ret))>0 "
    Set RsPart = GCn.Execute(mQry)
        
    With RsPart
        mConverting = True
        Label1.Visible = True
        If RsPart.RecordCount > 0 Then
            ProgressBar1.Visible = True
            Do While Not .EOF
                If VNull(!mQty) > 0 Then
                    Set RsTemp = GCn.Execute("Select Part_No, V_Date, Qty_Rec, Rate " & _
                                          "From Sp_StockNew S " & _
                                          "Where S.Part_No='" & !Part_No & "' And S.Tax_Yn = " & !Tax_YN & " " & _
                                          "And Left(S.DocId,1)='" & !Div_Code & "' And Qty_Rec>0  And S.V_Date<=" & ConvertDate(mTransDate) & "" & _
                                          "Order by V_Date Desc")
                                                                
                    
                    mQty = 0
                    mReqQty = 0
                    mFifoCost = 0
                    mChkQty = 0
                    Debug.Print RsTemp.RecordCount
                    
                    If RsTemp.RecordCount = 0 Then
                        With RstRep
                            .AddNew
                                !Div_Code = RsPart!Div_Code
                                !Part_No = RsPart!Part_No
                                !Qty = RsPart!mQty
                                !Tax_YN = RsPart!Tax_YN
                                !MRP_YN = RsPart!MRP_YN
                                !Rate = 0
                                !Amount = 0
                                !DAYS = 0
                                mChkQty = RsPart!mQty
                            .Update
                        End With
                    Else
                        Do Until RsTemp.EOF
                            If mQty < VNull(!mQty) Then
                                mReqQty = IIf((mQty + VNull(RsTemp!Qty_Rec)) > VNull(!mQty), VNull(!mQty) - mQty, RsTemp!Qty_Rec)
                                mQty = mQty + VNull(RsTemp!Qty_Rec)
                                
                                mCost = (mReqQty * VNull(RsTemp!Rate))
                                mFifoCost = mFifoCost + mCost
                                mChkQty = mChkQty + mReqQty
                                With RstRep
                                    .AddNew
                                        !Div_Code = RsPart!Div_Code
                                        !Part_No = left(RsPart!Part_No, 21)
                                        !Qty = mReqQty
                                        !Tax_YN = RsPart!Tax_YN
                                        !MRP_YN = RsPart!MRP_YN
                                        !Rate = VNull(RsTemp!Rate)
                                        !Amount = mCost
                                        !DAYS = Abs(DateDiff("D", PubLoginDate, RsTemp!V_Date))
                                    .Update
                                End With
                                                            
                                RsTemp.MoveNext
                            Else
                                Exit Do
                            End If
                        Loop
                    End If
                Else
                    With RstRep
                        .AddNew
                            !Div_Code = RsPart!Div_Code
                            !Part_No = RsPart!Part_No
                            !Qty = RsPart!mQty
                            !Tax_YN = RsPart!Tax_YN
                            !MRP_YN = RsPart!MRP_YN
                            !Rate = 0
                            !Amount = 0
                            !DAYS = 0
                            mChkQty = RsPart!mQty
                        .Update
                    End With
                End If
                If mChkQty <> RsPart!mQty Then
                    MsgBox "Diff RsPart = " & RsPart!mQty & " Calc = " & mChkQty
                End If
                If Round(ProgressBar1.Value) <> 100 Then ProgressBar1.Value = (.AbsolutePosition / .RecordCount) * 100
                
                For j = 1 To 999
                    Label1.Refresh
                Next j
                Set RsTemp = Nothing
                .MoveNext
            Loop
            ProgressBar1.Visible = False
        End If
        
    End With
    
    
    
    Set RsPart = Nothing
    Set RsTemp = Nothing
    
    
    
    RstRep.Filter = "Qty>0"
    If RstRep.RecordCount = 0 Then
        MsgBox "No Records in this Division for conversion. Process Terminated"
        Exit Sub
    End If
    
                        
                    
    mVTypeIssue = "SYIAD"
    mVTypeReceive = "SXRAD"
    mVDate = DateAdd("D", -1, PubStartDate)
    mVPrefixIssue = G_FaCn.Execute("Select IIF(IsNull(Prefix),'',Prefix) " & _
                              "From Voucher_Prefix Where V_Type = 'SYIAD'").Fields(0).Value
    mVPrefixReceive = G_FaCn.Execute("Select IIF(IsNull(Prefix),'',Prefix) " & _
                              "From Voucher_Prefix Where V_Type = 'SXRAD'").Fields(0).Value
                              
    mVNoIssue = G_FaCn.Execute("Select Val(Start_Srl_No) + 1 From Voucher_Prefix " & _
                          "Where V_Type='" & mVTypeIssue & "' And Date_From = #" & DateAdd("D", 1, CDate(mVDate)) & "#").Fields(0).Value
                                
    mVNoReceive = G_FaCn.Execute("Select Val(Start_Srl_No) + 1 From Voucher_Prefix " & _
                          "Where V_Type='" & mVTypeReceive & "' And Date_From = #" & DateAdd("D", 1, CDate(mVDate)) & "#").Fields(0).Value
                                
                
                
    mDocIdIssue = PubDivCode + PubSiteCode & PubSiteCode + Space(5 - Len(mVTypeIssue)) + _
                mVTypeIssue + Space(5 - Len(CStr(mVPrefixIssue))) + mVPrefixIssue + Space(8 - Len(CStr(mVNoIssue))) + CStr(mVNoIssue)
                
    mDocIdReceive = PubDivCode + PubSiteCode & PubSiteCode + Space(5 - Len(mVTypeReceive)) + _
                mVTypeReceive + Space(5 - Len(CStr(mVPrefixReceive))) + mVPrefixReceive + Space(8 - Len(CStr(mVNoReceive))) + CStr(mVNoReceive)
                
    
    
    
    G_FaCn.BeginTrans
    GCn.BeginTrans
    mTrans = True
    
            RstRep.Filter = "Qty>0"
            If RstRep.RecordCount > 0 Then
                mSrlIssue = 1
                mSrlReceive = 1
                mCount = 0
                ProgressBar1.Value = 0
                ProgressBar1.Visible = True
            Else
                MsgBox "No Records Found to Adjust"
                Exit Sub
            End If
            
        With RstRep
            Do Until RstRep.EOF
                If VNull(!Qty) > 0 Then
                    GCn.Execute "Insert Into SP_Stock(" _
                        & "DocID,Srl_No,V_Type,V_No,V_Date,Site_Code," _
                        & "Party_Code,Remark,Part_No, Qty_Iss,Tax_YN," _
                        & "MRP_YN,Rate,MRP_Rate,Amount,Net_Amt," _
                        & "Godown,Purpose, U_Name,U_EntDt,U_AE,V_Rate) " _
                        & "Values(" _
                        & "'" & mDocIdIssue & "'," & mSrlIssue & ",'" & mVTypeIssue & "'," & mVNoIssue & "," & ConvertDate(mTransDate) & ",'" & PubSiteCode & PubSiteCode & "'," _
                        & "'','','" & !Part_No & "'," & !Qty & ",0," _
                        & "0," & Val(!Rate) & ",0," & Val(!Amount) & ",0," _
                        & "'" & PubSprCounterGodown & "','','VatStk',#" & PubServerDate & "#,'A',0)"
                        
                        mSrlReceive = mSrlReceive + 1
                        
                    GCn.Execute "Insert Into SP_Stock(" _
                        & "DocID,Srl_No,V_Type,V_No,V_Date,Site_Code," _
                        & "Party_Code,Remark,Part_No, Qty_Rec,Tax_YN," _
                        & "MRP_YN,Rate,MRP_Rate,Amount,Net_Amt," _
                        & "Godown,Purpose, U_Name,U_EntDt,U_AE,V_Rate) " _
                        & "Values(" _
                        & "'" & mDocIdReceive & "'," & mSrlReceive & ",'" & mVTypeReceive & "'," & mVNoReceive & "," & ConvertDate(mTransDate) & ",'" & PubSiteCode & PubSiteCode & "'," _
                        & "'','','" & !Part_No & "'," & Abs(!Qty) & ",1," _
                        & "0," & Val(!Rate) & ",0," & Val(!Amount) & ",0," _
                        & "'" & PubSprCounterGodown & "','','VatStk',#" & PubServerDate & "#,'A',0)"
                        
                        mSrlIssue = mSrlIssue + 1
                End If
    
                
                mCount = mCount + 1
                If mCount < .RecordCount Then
                    ProgressBar1.Value = mCount * 100 / .RecordCount
                End If
                
                For j = 1 To 999
                    Label1.Refresh
                Next j
    
                .MoveNext
            Loop
        End With
            
                                    
        
        G_FaCn.Execute "UPDATE Voucher_Prefix SET Start_Srl_No = Start_Srl_No + 1 " & _
                       "WHERE V_Type = '" & mVTypeIssue & "'"
        G_FaCn.Execute "UPDATE Voucher_Prefix SET Start_Srl_No = Start_Srl_No + 1 " & _
                       "WHERE V_Type = '" & mVTypeReceive & "'"
                       
        
    G_FaCn.CommitTrans
    GCn.CommitTrans
    mTrans = False
    
    MsgBox " # Stock Conversion Done # "
    ProgressBar1.Visible = False
    mConverting = False
    Label1.Visible = False
    
    Set RsTemp = Nothing
    
Exit Sub
DispErr:
    If err.Description = "Table 'SP_StockNew' does not exist." Then Resume Next
    If mTrans = True Then G_FaCn.RollbackTrans:      GCn.RollbackTrans
    MsgBox err.Description
End Sub

Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        txt(Party).TEXT = RsParty!Name
        txt(Party).Tag = RsParty!Code
    End If
    txt(Party).SetFocus
    DGParty.Visible = False
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
    
    TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini
    Call Ini_Pub
    If RSOJPR = True Then
        txt(AdjType).Visible = False
    End If
    For I = 0 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg '&HDFF4F2
        txt(I).ForeColor = CtrlFColOrg
'        Txt(I).BorderStyle = 1
    Next
'    Hook TxtGrid(0).hWnd
    txt(VDate).Tag = PubLoginDate
    mVType = AdjIssVType

    Set DGPart.DataSource = RsPart

    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME from SubGroup " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        "Order by SubGroup.name"
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
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
        Set Master = GCn.Execute("Select Distinct S.DocID As SearchCode,S.DocID,S.V_Date,S.V_Type " _
            & "From SP_Stock S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' " & SiteCond & " and S.V_Type In ('" & AdjIssVType & "','" & AdjRecVType & "') " _
            & "Order by S.V_Date desc,S.V_Type")
    Else
        Set Master = GCn.Execute("Select Distinct Top 1  S.DocID As SearchCode,S.DocID,S.V_Date,S.V_Type " _
            & "From SP_Stock S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' " & SiteCond & " and S.V_Type In ('" & AdjIssVType & "','" & AdjRecVType & "') " _
            & "Order by S.V_Date desc,S.V_Type")
    
    End If

    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
   CheckError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsParty = Nothing
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
    txt(VDate).TEXT = txt(VDate).Tag
    txt(DocType).TEXT = "Issue"
    txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
    txt(DocID).Tag = txt(DocID)
    mPartyType = 0
    txt(DocType).SetFocus
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
    FGrid.AddItem FGrid.Rows
    txt(Party).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim vBook As Variant, mTrans As Boolean
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            vBook = Master.AbsolutePosition
            GCn.BeginTrans
            mTrans = True
            If txt(DocType).TEXT = "Issue" Then
                UpdStkTableToTable txt(DocID), "+", "I"
                'eof edit stock upd
            ElseIf txt(DocType).TEXT = "Receipt" Then
                'Stock Updation during edit
                UpdStkTableToTable txt(DocID), "-", "R"
                'eof edit stock upd
            End If

            GCn.Execute ("Delete From SP_Stock Where DocID='" & txt(DocID).TEXT & "'")
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
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    Dim SiteCond As String
    SiteCond = " And V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and  " & cMID("S.DocID", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    If PubBackEnd = "A" Then
        GSQL = "Select Distinct S.DocId As SearchCode," & cDt("S.V_Date") & " As V_Date ,Switch(S.V_Type='" & AdjIssVType & "','Issue',S.V_Type='" & AdjRecVType & "','Receipt') As DocType, " & cTrim(cMID("S.DocID", "9", "5")) & " As VPrefix, " & cCStr("S.V_No", 10) & " As V_No, S.V_Date AS VDate,SubGroup.Name As AcName,S.Remark From SP_Stock S Left Join SubGroup on S.Party_Code=SubGroup.SubCode Where left(S.DocID,1)='" & PubDivCode & "'  " & SiteCond & " and S.V_Type In ('" & AdjIssVType & "','" & AdjRecVType & "') Order by S.V_Date desc"
    ElseIf PubBackEnd = "S" Then
        GSQL = "Select Distinct S.DocId As SearchCode," & cDt("S.V_Date") & " As V_Date , Case When S.V_Type='" & AdjIssVType & "' Then 'Issue' When S.V_Type='" & AdjRecVType & "' Then 'Receipt' End As DocType, " & cTrim(cMID("S.DocID", "9", "5")) & " As VPrefix, " & cCStr("S.V_No", 10) & " As V_No, S.V_Date AS VDate,SubGroup.Name As AcName,S.Remark From SP_Stock S Left Join SubGroup on S.Party_Code=SubGroup.SubCode Where left(S.DocID,1)='" & PubDivCode & "' " & SiteCond & " and S.V_Type In ('" & AdjIssVType & "','" & AdjRecVType & "') Order by S.V_Date desc"
    End If
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    MoveRec
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_ePrn()
Dim X1
Dim mQry As String
Dim Rst As ADODB.Recordset
Dim I   As Integer
On Error GoTo ELoop
If Master.RecordCount <= 0 Then Exit Sub
  
    
    
mQry = "Select SS.DocID, SS.Srl_No, SS.V_Type, SS.V_No, SS.V_Date, SS.Site_Code," _
            & "SS.Party_Code, SS.Remark, SS.Part_No," & IIf(mVType = "SYIAD", "SS.Qty_ISS", "SS.Qty_Rec") & " as Qty, " _
            & "SS.Tax_YN, SS.MRP_YN, SS.Rate, SS.MRP_Rate, SS.Amount, SS.Net_Amt, SS.Godown, SS.Purpose, " _
            & "S.Name as Party_Name,P.Part_Name,G.God_Name, " _
            & " '" & txt(DocType) & "' as DocType, '" & txt(AdjType) & "' as AdjType, SS.GatePassNo, " & IIf(mVType = "SYIAD", "'Issue'", "'Receive'") & " As IssueReceive " _
            & "From (((SP_Stock SS " _
            & "Left Join SubGroup S On SS.Party_Code=S.SubCode) " _
            & "Left Join Part P On SS.Part_No=P.Part_No) " _
            & "Left Join Godown G On SS.Godown=G.God_Code) " _
            & "Where DocId='" & txt(DocID) & "' "

Set Rst = GCn.Execute(mQry)
X1 = CreateFieldDefFile(Rst, PubRepoPath + "\StkTrn.ttx", True)
Set rpt = rdApp.OpenReport(PubRepoPath + "\StkTrn.RPT")
rpt.Database.SetDataSource Rst
rpt.ReadRecords
Report_View rpt, "Stock Adjustment Slip", 0, False
Set Rst = Nothing
Exit Sub
ELoop:
    MsgBox err.Description
End Sub


Private Sub TopCtrl1_eRef()
On Error GoTo ELoop
    RsParty.Requery
    RsPart.Requery
    RsGodown.Requery
    'Master.Requery
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean, mGridFilled As Boolean, ReOrderQty As Double
Dim mGatePassNo As Long
Dim Rst As ADODB.Recordset, DocIdHlp$, TmpStr$
On Error GoTo ELoop
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    
    If IsValid(txt(DocType), LBL(1)) = False Then Exit Sub
    If IsValid(txt(VDate), LBL(0)) = False Then Exit Sub
    If IsValid(txt(SerialNo), LBL(2)) = False Then Exit Sub
   ' If IsValid(txt(AdjType), Label3(4)) = False Then Exit Sub
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_PNo) <> "" Then
            If FGrid.TextMatrix(I, Col_MRP) = "" Then MsgBox "Please Specify MRP Yes/No in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_MRP: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
            If FGrid.TextMatrix(I, Col_Taxable) = "" Then MsgBox "Please Specify Taxable Yes/No in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Taxable: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
            If Val(FGrid.TextMatrix(I, Col_Qty)) = 0 Then MsgBox "Please Specify Quantity in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Qty: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
            If FGrid.TextMatrix(I, Col_Godown) = "" Then MsgBox "Please Specify Godown in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Godown: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
            mGridFilled = True
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Item Detail", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = Col_PNo: FGrid.SetFocus: Exit Sub ': FGrid.CellBackColor = CellBackColEnter
    
    If TopCtrl1.TopText2 = "Add" Then
        txt(DocID).Tag = txt(DocID)
        If GCn.Execute("Select Count(*) From SP_Stock Where DocID='" & txt(DocID) & "'").Fields(0) > 0 Then
            If VoucherEditFlag Then
                MsgBox "Serial No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                Exit Sub
            Else
                txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(txt(DocID).Tag, Document_No)) Then
                    MsgBox "Serial No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    Exit Sub
                End If
            End If
        End If
        
        mGatePassNo = GCn.Execute("Select " & vIsNull("Max(GatePassNo)", "0") & " + 1 From Sp_Stock ").Fields(0).Value
        
    End If
    DocIdHlp = UCase(Replace(txt(DocID), " ", ""))
    
    GCn.BeginTrans
        mTrans = True
        If txt(DocType).TEXT = "Issue" Then
            TmpStr = "Qty_Iss"
            'Stock Updation during edit
            UpdStkTableToTable txt(DocID), "+", "I"
            'eof edit stock upd
        ElseIf txt(DocType).TEXT = "Receipt" Then
            TmpStr = "Qty_Rec"
            'Stock Updation during edit
            UpdStkTableToTable txt(DocID), "-", "R"
            'eof edit stock upd
        End If
        
        GCn.Execute ("Delete From SP_Stock Where DocID='" & txt(DocID).TEXT & "'")
        
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_PNo) <> "" And Val(FGrid.TextMatrix(I, Col_Qty)) <> 0 Then
                GCn.Execute "Insert Into SP_Stock(" _
                    & "DocID,Srl_No,V_Type,V_No,V_Date,Site_Code," _
                    & "Party_Code,Remark,Part_No," & TmpStr & ",Tax_YN," _
                    & "MRP_YN,Rate,MRP_Rate,Amount,Net_Amt," _
                    & "Godown,Purpose,U_Name,U_EntDt,U_AE,V_Rate, GatePassNo) " _
                    & "Values(" _
                    & "'" & txt(DocID).TEXT & "'," & I & ",'" & mVType & "'," & txt(SerialNo).TEXT & "," & ConvertDate(txt(VDate).TEXT) & ",'" & PubSiteCode & PubSiteCode & "'," _
                    & "'" & txt(Party).Tag & "','" & txt(Remark).TEXT & "','" & FGrid.TextMatrix(I, Col_PNo) & "'," & Val(FGrid.TextMatrix(I, Col_Qty)) & "," & IIf(FGrid.TextMatrix(I, Col_Taxable) = "Yes", 1, 0) & "," _
                    & "" & IIf(FGrid.TextMatrix(I, Col_MRP) = "Yes", 1, 0) & "," & Val(FGrid.TextMatrix(I, Col_Rate)) & "," & Val(FGrid.TextMatrix(I, Col_MRPRate)) & "," & Val(FGrid.TextMatrix(I, Col_Amt)) & "," & Val(FGrid.TextMatrix(I, Col_Amt)) & "," _
                    & "'" & FGrid.TextMatrix(I, Col_GodownCode) & "','" & left(txt(AdjType), 1) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(TopCtrl1.TopText2 = "Add", "A", "E") & "'," & Val(FGrid.TextMatrix(I, Col_Rate)) & ", " & mGatePassNo & ")"
                
                If txt(DocType).TEXT = "Issue" Then
                    Call UpdStkGridToTable(FGrid.TextMatrix(I, Col_PNo), "-", FGrid.TextMatrix(I, Col_MRP), FGrid.TextMatrix(I, Col_Taxable), Val(FGrid.TextMatrix(I, Col_Qty)))
                ElseIf txt(DocType).TEXT = "Receipt" Then
                    Call UpdStkGridToTable(FGrid.TextMatrix(I, Col_PNo), "+", FGrid.TextMatrix(I, Col_MRP), FGrid.TextMatrix(I, Col_Taxable), Val(FGrid.TextMatrix(I, Col_Qty)))
                End If
'modi lps stopped at Cuttack 02.09.03
'                ' Used For Creating a Indent When Stock is Less Than Re-Order Level
'                ReOrderQty = GCn.Execute("Select ReOrd_Lvl From Part Where Part_No='" & FGrid.TextMatrix(i, Col_PNo) & "' AND div_code ='" & PubDivCode & "'").Fields(0).Value
'                If (Val(FGrid.TextMatrix(i, Col_MRPStkTP)) - Val(FGrid.TextMatrix(i, Col_Qty))) < ReOrderQty Then
''                    CreateSprIndent
'                End If
            End If
        Next
        If TopCtrl1.TopText2 = "Add" Then
            'Voucher Serial No. Updation LPS 21-05-03
            'update Table only when DocSrlNo >Table.SerialNo
            UpdVouSrlNo GCnFaS, txt(DocID), txt(VDate)
        End If
    GCn.CommitTrans
    mTrans = False
    mSearchCode = txt(DocID)
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select Distinct S.DocID As SearchCode,S.DocID,S.V_Date,S.V_Type " _
            & "From SP_Stock S " _
            & "Where left(S.DocID,1)='" & PubDivCode & "' and S.V_Type In ('" & AdjIssVType & "','" & AdjRecVType & "') And  S.DocID = '" & mSearchCode & "' " _
            & "Order by S.V_Date desc,S.V_Type")
    End If
    Master.FIND "SearchCode = '" & mSearchCode & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(SerialNo)) > Val(DeCodeDocID(txt(DocID).Tag, Document_No)) Then
            MsgBox "Serial No." & Trim(DeCodeDocID(txt(DocID).Tag, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
        txt(VDate).Tag = txt(VDate).TEXT
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
        ListArray = Array("Issue", "Receipt")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
        OldDocType = txt(DocType).TEXT
    Case Party
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case AdjType
        ListArray = Array("Assemble", "Breakage", "Dismental", "Loan", "Others")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 5)
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
    Case Party
        DGridTxtKeyDown DGParty, txt, Party, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
    Case AdjType
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 1400
End Select
If FrmList.Visible = False And DGParty.Visible = False Then
    If Index = Remark And (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    Else
        If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
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
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case DocType, AdjType
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
Dim Rst As ADODB.Recordset
On Error GoTo ELoop
Select Case Index
    Case DocType
        txt(Index).TEXT = ListView.SelectedItem.TEXT
        If Not Trim(txt(Index).TEXT) <> "Issue" Or Trim(txt(Index).TEXT) <> "Receipt" Then
            txt(Index).TEXT = "Issue"
        End If
        If Trim(txt(Index).TEXT) = "Issue" Then
            mVType = AdjIssVType
        ElseIf Trim(txt(Index).TEXT) = "Receipt" Then
            mVType = AdjRecVType
        End If
        If txt(Index).TEXT <> OldDocType Then
            txt(Party).TEXT = ""
            txt(Party).Tag = ""
        End If
        txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
        txt(DocID).Tag = txt(DocID)
    Case VDate
        If Len(Trim(txt(VDate).TEXT)) = 0 Then
            MsgBox "Blank Date", vbOKOnly, "Validation Check"
            Cancel = True
        Else
            txt(Index) = RetDate(txt(Index))
            txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            txt(DocID).Tag = txt(DocID)
        End If
    Case SerialNo
        If VoucherEditFlag = True Then      ' Manual
            txt(DocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
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
    Case Party
        If RsParty.RecordCount > 0 Then
            If txt(Index).TEXT <> "" Then
                txt(Index).TEXT = RsParty!Name
                txt(Index).Tag = RsParty!Code
            End If
        Else
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        End If
    Case AdjType
        txt(Index).TEXT = ListView.SelectedItem.TEXT
End Select
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
'    Ctrl_GetFocus TxtGrid(Index)
    Grid_Hide
    If FrmDetail.Visible = False Then FrmDetail.Visible = True
'    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
        Case Col_PNo
            TxtGrid(0).MaxLength = 22
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "Code"
9            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "Code='" & FGrid.TextMatrix(FGrid.Row, Col_PNo) & "'"
6
 If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case Col_PName
            TxtGrid(0).MaxLength = 40
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "Name"
            If FGrid.TextMatrix(FGrid.Row, Col_PName) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, Col_PName) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case Col_LName
            TxtGrid(0).MaxLength = 40
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "LName"
            If FGrid.TextMatrix(FGrid.Row, Col_LName) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "LName ='" & FGrid.TextMatrix(FGrid.Row, Col_LName) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case Col_Godown
            TxtGrid(0).MaxLength = 20
            If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Col_Godown) = "" Then Exit Sub
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
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName, , Col_MRP
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
            If DGPart.Visible = False Then DGridColSwap DGPart, 2
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 2, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_LName, 1
                End If
            End If
        Case Col_MRP
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Godown
                End If
            End If
        Case Col_Taxable
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Godown
                End If
            End If
        Case Col_Qty
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Godown
                End If
            End If
        Case Col_Rate
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Godown, , Col_Godown
                End If
            End If
        Case Col_Godown
            DGridTxtKeyDown DGGodown, TxtGrid, Index, RsGodown, KeyCode, True, 1, frmGodown, "frmGodown"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Godown
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
        Case Col_Godown
            If DGGodown.Visible = True Then
                If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                    DGridTxtKeyPress TxtGrid, Index, RsGodown, KeyAscii, "Name"
                End If
            End If
        Case Col_Qty
            NumPress TxtGrid(Index), KeyAscii, 8, 3
        Case Col_Rate
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
        Case Col_Qty
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.000")
            CountItem
            Amt_Cal
        Case Col_Rate
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
            Amt_Cal
    End Select
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
'    FrmDetail.Visible = True
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.Col = Col_Unit Then Exit Sub
    Select Case FGrid.Col
        Case Col_PNo, Col_PName, Col_LName
            GridDblClick Me, FGrid, TxtGrid, 0
        Case Col_MRP, Col_Taxable, Col_Qty, Col_Rate
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                GridDblClick Me, FGrid, TxtGrid, 0
            End If
        Case Col_Godown
            If FGrid.TextMatrix(FGrid.Row, Col_Qty) <> "" Then
                If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                    GridDblClick Me, FGrid, TxtGrid, 0
                End If
            End If
    End Select
    TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
'    FGrid.CellBackColor = CellBackColEnter
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
'        FGrid.CellBackColor = CellBackColLeave
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
'        FGrid.CellBackColor = CellBackColLeave
        SendKeysA vbKeyTab, True
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGrid.Col
            Case Col_MRP, Col_Taxable
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            Case Col_Qty
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
            Case Col_PNo, Col_PName, Col_LName
                GridDblClick Me, FGrid, TxtGrid, 0
            Case Col_Unit
                GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, Col_LName, , Col_MRP
            Case Col_MRP, Col_Taxable, Col_Qty, Col_Rate
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
        Case Col_PNo, Col_PName, Col_LName
            Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
        Case Col_Unit
            FGrid.Col = FGrid.Col + 1
            FGrid.SetFocus
        Case Col_MRP, Col_Taxable
            If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
                Get_Text Me, FGrid, TxtGrid, 0, False, KeyAscii
            End If
        Case Col_Rate, Col_Qty
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
                Amt_Cal
                MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, Col_PNo), _
                    FGrid.TextMatrix(FGrid.Row, Col_PName), FGrid.TextMatrix(FGrid.Row, Col_LName), _
                    Col_MRPStkTB, Col_MRPStkTP, _
                    Col_TBStk, Col_TPStk, _
                    Col_MRPRate, Col_TBRate, _
                    Col_TPRate, Col_Bin, _
                    Col_LastRate, Col_HPRate, Col_LPRate, mCheckNegetiveStockSiteWise
            End If
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
    FGrid.TextMatrix(FGrid.Row, Col_PName) = IIf(IsNull(RsPart!Name), "", RsPart!Name)
    FGrid.TextMatrix(FGrid.Row, Col_LName) = IIf(IsNull(RsPart!LName), "", RsPart!LName)
    
        MainLib.Fill_Data mPartyType, LblFrm, FGrid, _
        RsPart!Code, IIf(IsNull(RsPart!Name), "", RsPart!Name), IIf(IsNull(RsPart!LName), "", RsPart!LName), _
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
           ' FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsPart!PurDisc_Per, "0.00")
        End If
    End If
End If
If FGrid.TextMatrix(FGrid.Rows - 1, Col_PNo) <> "" Then FGrid.AddItem FGrid.Rows
End Sub


Sub Ini_Pub()
    Dim RsTemp As ADODB.Recordset
    
    Set RsTemp = GCn.Execute("Select CheckNegetiveStockSiteWise From Syctrl")
    If RsTemp.RecordCount > 0 Then
        mCheckNegetiveStockSiteWise = VNull(RsTemp!CheckNegetiveStockSiteWise)
    End If
End Sub
