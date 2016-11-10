VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmSyCtrl 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "System Controls"
   ClientHeight    =   8595
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
   ScaleHeight     =   8595
   ScaleWidth      =   11820
   Visible         =   0   'False
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
      Height          =   225
      Index           =   18
      Left            =   7125
      MaxLength       =   40
      TabIndex        =   38
      Top             =   2535
      Width           =   435
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
      Left            =   9495
      MaxLength       =   40
      TabIndex        =   40
      Top             =   2295
      Width           =   435
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
      Index           =   16
      Left            =   7125
      MaxLength       =   40
      TabIndex        =   37
      Top             =   2295
      Width           =   435
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
      Index           =   15
      Left            =   9495
      MaxLength       =   40
      TabIndex        =   35
      Top             =   1275
      Width           =   435
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
      Left            =   2505
      MaxLength       =   40
      TabIndex        =   33
      Top             =   2805
      Width           =   435
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
      Left            =   2505
      MaxLength       =   40
      TabIndex        =   11
      Top             =   2550
      Width           =   435
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
      Index           =   12
      Left            =   2505
      MaxLength       =   40
      TabIndex        =   10
      Top             =   2295
      Width           =   435
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   720
      TabIndex        =   29
      Top             =   5685
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   60
         TabIndex        =   30
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
      DataField       =   "`"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Index           =   11
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   3525
      Width           =   11400
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
      Left            =   8940
      MaxLength       =   15
      TabIndex        =   13
      Top             =   2040
      Width           =   990
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
      Left            =   6045
      MaxLength       =   16
      TabIndex        =   12
      Top             =   2040
      Width           =   1515
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
      Left            =   2505
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1530
      Width           =   2280
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
      Index           =   7
      Left            =   2505
      MaxLength       =   12
      TabIndex        =   7
      Top             =   1785
      Width           =   2280
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
      Left            =   7125
      MaxLength       =   16
      TabIndex        =   8
      Top             =   1785
      Width           =   435
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
      Left            =   9495
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1785
      Width           =   435
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
      Index           =   4
      Left            =   7125
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1530
      Width           =   435
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
      Left            =   7125
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1275
      Width           =   435
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
      Index           =   3
      Left            =   9495
      MaxLength       =   16
      TabIndex        =   3
      Top             =   1530
      Width           =   435
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
      Index           =   1
      Left            =   2505
      MaxLength       =   40
      TabIndex        =   9
      Top             =   2040
      Width           =   435
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
      Index           =   0
      Left            =   2505
      TabIndex        =   1
      Top             =   1275
      Width           =   2280
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
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2355
      Left            =   10215
      Negotiate       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4154
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
            ColumnWidth     =   2789.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Wise Display Y/N  :"
      BeginProperty Font 
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
      Index           =   19
      Left            =   5250
      TabIndex        =   42
      Top             =   2535
      Width           =   1905
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lock Financial Year"
      BeginProperty Font 
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
      Index           =   18
      Left            =   7620
      TabIndex        =   41
      Top             =   2295
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SAT Applicable(Y/N)  :"
      BeginProperty Font 
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
      Index           =   17
      Left            =   5250
      TabIndex        =   39
      Top             =   2295
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Lock (In Days)"
      BeginProperty Font 
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
      Left            =   7830
      TabIndex        =   36
      Top             =   1275
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Siebel Active     (Y/N)  :"
      BeginProperty Font 
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
      Left            =   630
      TabIndex        =   34
      Top             =   2805
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SDT Applicable(Y/N)  :"
      BeginProperty Font 
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
      Left            =   630
      TabIndex        =   32
      Top             =   2550
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT Applicable(Y/N)   :"
      BeginProperty Font 
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
      Left            =   630
      TabIndex        =   31
      Top             =   2295
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Transaction's Footer Details :"
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
      Index           =   13
      Left            =   240
      TabIndex        =   28
      Top             =   3270
      Width           =   3360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Round Type :"
      BeginProperty Font 
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
      Left            =   7785
      TabIndex        =   27
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Round Off Position :"
      BeginProperty Font 
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
      Left            =   4290
      TabIndex        =   26
      Top             =   2040
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jurisdiction City           :"
      BeginProperty Font 
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
      Index           =   10
      Left            =   630
      TabIndex        =   25
      Top             =   1530
      Width           =   1830
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Road Permit Caption :"
      BeginProperty Font 
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
      Left            =   630
      TabIndex        =   24
      Top             =   1785
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cr Limit Checking :"
      BeginProperty Font 
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
      Index           =   3
      Left            =   5505
      TabIndex        =   23
      Top             =   1785
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Prefix :"
      BeginProperty Font 
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
      Left            =   8235
      TabIndex        =   22
      Top             =   1785
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line Character :"
      BeginProperty Font 
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
      Left            =   5760
      TabIndex        =   21
      Top             =   1530
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Restrict Godown         :"
      BeginProperty Font 
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
      Index           =   5
      Left            =   630
      TabIndex        =   20
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print Page No. :"
      BeginProperty Font 
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
      Left            =   8175
      TabIndex        =   19
      Top             =   1530
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Date on Report :"
      BeginProperty Font 
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
      Index           =   7
      Left            =   5055
      TabIndex        =   18
      Top             =   1275
      Width           =   2010
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name                    :"
      BeginProperty Font 
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
      Left            =   630
      TabIndex        =   17
      Top             =   1275
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RSO Code :"
      BeginProperty Font 
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
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmSyCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'FA Connection for Works FAData
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim RsSite As adodb.Recordset
Dim Syctrl As adodb.Recordset
Dim ExitCtrl As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

'grid color scheme
Private Const CellBackColLeave As String = &HC8E8DA
Private Const CellForeColLeave As String = &H0&
Private Const CellBackColEnter As String = &HC0E0FF
Private Const GridBackColorBkg As String = &HBAD3C9

Dim MyIndex As Byte
Private Const Site_Code             As Byte = 0
Private Const Restrict_Godown       As Byte = 1
Private Const PrintReportDate       As Byte = 2
Private Const PrintPageNo           As Byte = 3
Private Const LineFill              As Byte = 4
Private Const AmountPrefix          As Byte = 5
Private Const CrLimitCheck          As Byte = 6
Private Const Form31Caption         As Byte = 7
Private Const JuriCity              As Byte = 8
Private Const RoundOffPosition      As Byte = 9
Private Const RoundOffType          As Byte = 10
Private Const SprMoneyRectFooter    As Byte = 11
Private Const VAT_YN                As Byte = 12
Private Const SDT_YN                As Byte = 13
Private Const SiebelActiveYN        As Byte = 14
Private Const EditLock              As Byte = 15
Private Const SAT_YN                As Byte = 16
Private Const LockFinancialYear     As Byte = 17
Private Const SiteWiseDisplay_YN                As Byte = 18

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next
End Sub

'* Used for intialize grid columns
Private Sub Grid_Ini()
    DGSite.left = Me.width - (DGSite.width + mRtScale): DGSite.top = mTopScale
End Sub

Private Sub Grid_Hide()
    If DGSite.Visible = True Then DGSite.Visible = False
End Sub

Private Sub MoveRec()
Dim Syctrl1 As adodb.Recordset
On Error GoTo Eloop
'General Settings
If Syctrl!Site_Code <> "" And RsSite.EOF = False And RsSite.BOF = False Then
    RsSite.MoveFirst
    RsSite.FIND ("Code ='" & Syctrl!Site_Code & "'")
    Txt(Site_Code).Tag = Syctrl!Site_Code
    Txt(Site_Code) = IIf(RsSite.EOF, "", RsSite!Name)
Else
    Txt(Site_Code) = ""
    Txt(Site_Code).Tag = ""
End If
If RsSite.RecordCount > 0 And RsSite.EOF Then RsSite.MoveFirst
Txt(Restrict_Godown) = IIf(Syctrl!Restrict_Godown = 1, "Yes", "No")
Txt(PrintReportDate) = IIf(Syctrl!PrintReportDate = 1, "Yes", "No")
Txt(PrintPageNo) = IIf(Syctrl!PrintPageNo = 1, "Yes", "No")
Txt(LineFill) = IIf(IsNull(Syctrl!LineFill), "", Syctrl!LineFill)
Txt(AmountPrefix) = IIf(IsNull(Syctrl!AmountPrefix), "", Syctrl!AmountPrefix)
Txt(CrLimitCheck) = IIf(Syctrl!CrLimitCheck = 1, "Yes", "No")
Txt(Form31Caption) = IIf(IsNull(Syctrl!Form31Caption), "", Syctrl!Form31Caption)
Txt(JuriCity) = IIf(IsNull(Syctrl!Juri_CITY), "", Syctrl!Juri_CITY)
Txt(VAT_YN) = IIf(VNull(Syctrl!VAT_YN) = 0, "No", "Yes")
Txt(SDT_YN) = IIf(VNull(Syctrl!SDT_YN) = 0, "No", "Yes")
Txt(SAT_YN) = IIf(VNull(Syctrl!SAT_YN) = 0, "No", "Yes")
Txt(SiteWiseDisplay_YN) = IIf(VNull(Syctrl!SiteWiseDisplaY_N) = 0, "No", "Yes")
Txt(LockFinancialYear) = IIf(VNull(Syctrl!LockFinancialYear) = 0, "No", "Yes")
Txt(SiebelActiveYN) = IIf(VNull(Syctrl!SiebelActiveYN) = 0, "No", "Yes")
Txt(EditLock) = VNull(Syctrl!EditLock)
If Syctrl!RoundOffPosition = 0 Then
    Txt(RoundOffPosition) = "0- > Rupees"
ElseIf Syctrl!RoundOffPosition = 1 Then
    Txt(RoundOffPosition) = "1- >10 Paise"
ElseIf Syctrl!RoundOffPosition = 2 Then
    Txt(RoundOffPosition) = "2- >25 Paise"
ElseIf Syctrl!RoundOffPosition = 3 Then
    Txt(RoundOffPosition) = "3- > 50 Paise"
ElseIf Syctrl!RoundOffPosition = 4 Then
    Txt(RoundOffPosition) = "4- > No Round Off"
End If
If Syctrl!RoundOffType = "S" Then
    Txt(RoundOffType) = "Standard"
ElseIf Syctrl!RoundOffType = "U" Then
    Txt(RoundOffType) = "Upper"
ElseIf Syctrl!RoundOffType = "L" Then
    Txt(RoundOffType) = "Lower"
End If
Txt(SprMoneyRectFooter) = IIf(IsNull(Syctrl!SprMoneyRectFooter), "", Syctrl!SprMoneyRectFooter)

Set Syctrl1 = Nothing
TopCtrl1.tAdd = False
TopCtrl1.tDel = False
TopCtrl1.tFirst = False
TopCtrl1.tPrev = False
TopCtrl1.tNext = False
TopCtrl1.tLast = False
TopCtrl1.tFind = False
TopCtrl1.tPrn = False
Exit Sub
Eloop:
    CheckError
End Sub

Private Sub DGSite_Click()
On Error GoTo Eloop
    If RsSite.RecordCount > 0 Then
        Txt(MyIndex).TEXT = RsSite!Name
        Txt(MyIndex).Tag = RsSite!Code
    End If
    Txt(MyIndex).SetFocus
    DGSite.Visible = False
Exit Sub
Eloop:
    CheckError
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Eloop
    FormKeyDown Me, KeyCode, Shift
Exit Sub
Eloop:
    CheckError
End Sub

Private Sub Form_Load()
On Error GoTo Eloop
Dim I As Byte
    TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini

    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg '&HDFF4F2
        Txt(I).ForeColor = CtrlFColOrg
    Next
    Set RsSite = New adodb.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "Select Site_Code as Code, Site_Desc As Name From Site Order by Site_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGSite.DataSource = RsSite
    
    Set Syctrl = New adodb.Recordset
    Syctrl.LockType = adLockOptimistic
    Syctrl.CursorLocation = adUseClient
    Syctrl.CursorType = adOpenDynamic
    Set Syctrl = GCn.Execute("Select * from Syctrl")
    
    Disp_Text SETS("INI", Me, Syctrl)
    MoveRec
Exit Sub
Eloop:
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
    Set RsSite = Nothing
    Set Syctrl = Nothing
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo Eloop
    Disp_Text SETS("EDIT", Me, Syctrl)
    If Txt(Site_Code) = "" Then
        Txt(Site_Code).SetFocus
    Else
        'Txt(Site_Code).Enabled = False
        Txt(PrintReportDate).SetFocus
    End If
Exit Sub
Eloop:
    CheckError
End Sub

Private Sub TopCtrl1_eRef()
On Error GoTo Eloop
    RsSite.Requery
    Syctrl.Requery
Exit Sub
Eloop:
    CheckError
End Sub

Private Sub TopCtrl1_eSave()
Dim mTrans As Boolean, SyctrlSql$
On Error GoTo Eloop
'Apply necessary validations
If IsValid(Txt(Site_Code), "Site Name") = False Then Exit Sub
    GCn.BeginTrans
        mTrans = True
        '0- > Rupees ,  1- >10 Paise , 2- >25 Paise ,  3- > 50 Paise, 4- > No Round Off
        If TopCtrl1.TopText2 = "Edit" Then   'Edit Bill
            GSQL = "update Syctrl set Site_Code = '" & Txt(Site_Code).Tag & "', Restrict_Godown= " & IIf(Txt(Restrict_Godown) = "Yes", 1, 0) & _
                ",PrintReportDate=" & IIf(Txt(PrintReportDate) = "Yes", 1, 0) & " , PrintPageNo= " & IIf(Txt(PrintPageNo) = "Yes", 1, 0) & _
                ", LineFill = '" & Txt(LineFill) & "', AmountPrefix = '" & Txt(AmountPrefix) & _
                "',CrLimitCheck = " & IIf(Txt(CrLimitCheck) = "Yes", 1, 0) & ", Form31Caption = '" & Txt(Form31Caption) & _
                "',Juri_City = '" & Txt(JuriCity) & _
                "',RoundOffPosition = " & Val(left(Txt(RoundOffPosition), 1)) & ",RoundOffType = '" & left(Txt(RoundOffType), 1) & _
                "',SprMoneyRectFooter = '" & Txt(SprMoneyRectFooter) & "',VAT_YN = " & IIf(Txt(VAT_YN) = "No", 0, 1) & ", " & _
                "SDT_YN = " & IIf(Txt(SDT_YN) = "No", 0, 1) & ", SAT_YN = " & IIf(Txt(SAT_YN) = "No", 0, 1) & ", LockFinancialYear  = " & IIf(Txt(LockFinancialYear) = "No", 0, 1) & ",SiebelActiveYN = " & IIf(Txt(SiebelActiveYN) = "No", 0, 1) & ", EditLock=" & Val(Txt(EditLock)) & ", SiteWiseDisplaY_N = " & IIf(Txt(SiteWiseDisplay_YN) = "No", 0, 1) & ""
            GCn.Execute GSQL
        End If
        If Txt(SDT_YN) = "Yes" Then
            GCn.Execute ("Update Syctrl set TOTCaption = 'S D T'")
            GCn.Execute ("Update Syctrl set TOT_YN = 1")
        Else
            GCn.Execute ("Update Syctrl set TOTCaption = 'T O T'")
            GCn.Execute ("Update Syctrl set TOT_YN = 1")
        End If
    GCn.CommitTrans
    mTrans = False
    Syctrl.Requery
    Disp_Text SETS("INI", Me, Syctrl)
    MoveRec
Exit Sub
Eloop:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo Eloop
Dim I As Byte
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Syctrl)
        MoveRec
        For I = 0 To Txt.Count - 1
            Txt(I).BackColor = CtrlBColOrg
            Txt(I).ForeColor = CtrlFColOrg
        Next
    End If
Exit Sub
Eloop:
    CheckError
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
On Error GoTo Eloop
Ctrl_GetFocus Txt(Index)
Grid_Hide
MyIndex = Index
Select Case Index
    Case Site_Code
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsSite!Name Then
            RsSite.MoveFirst
            RsSite.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    Case RoundOffPosition
        ListArray = Array("0- > Rupees", "1- >10 Paise", "2- >25 Paise", "3- > 50 Paise", "4- > No Round Off")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 5)
    Case RoundOffType
        ListArray = Array("Standard", "Upper", "Lower")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 3)
End Select
Exit Sub
Eloop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Eloop
    Select Case Index
        Case Site_Code
            DGridTxtKeyDown DGSite, Txt, Index, RsSite, KeyCode, False, 1
        Case RoundOffPosition
            ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1600
        Case RoundOffType
            ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1000
    End Select
    If Txt(Index).MultiLine = False Then
        If DGSite.Visible = False And FrmList.Visible = False Then    'Arrow Key
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = SDT_YN Then
               If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
            End If
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> SDT_YN Then Ctrl_DownKeyDown KeyCode, Shift
            If TopCtrl1.TopText2.CAPTION = "Edit" Then
                If Index <> Site_Code And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
Exit Sub
Eloop:
    CheckError
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Eloop
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
Select Case Index
'    Case Valid_Day, RebDays
'        NumPress Txt(Index), KeyAscii, 3, 0
'    Case DelayInttRate
'        NumPress Txt(Index), KeyAscii, 2, 2
    Case EditLock
        NumPress Txt(Index), KeyAscii, 3, 0
    Case Restrict_Godown, PrintReportDate, PrintPageNo, CrLimitCheck, VAT_YN, SDT_YN, SAT_YN, SiebelActiveYN, LockFinancialYear, SiteWiseDisplay_YN
        If UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index).TEXT = "Yes"
            KeyAscii = 0
        Else    'If KeyAscii = 87 Or KeyAscii = 119 Then   ' W/w
            If KeyAscii <> vbKeyReturn Then
                Txt(Index).TEXT = "No"
                KeyAscii = 0
            End If
        End If
    Case Site_Code
        If DGSite.Visible = True Then DGridTxtKeyPress Txt, Index, RsSite, KeyAscii, "Name"
End Select
Exit Sub
Eloop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Eloop
Select Case Index
    Case RoundOffPosition, RoundOffType
        ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
End Select
Exit Sub
Eloop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo Eloop
    Select Case Index
        Case Site_Code
            If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or Txt(Index).TEXT = "" Then
                Txt(Index) = ""
                Txt(Index).Tag = ""
            Else
                Txt(Index).Tag = RsSite!Code
                Txt(Index) = RsSite!Name
            End If
    End Select
Exit Sub
Eloop:
    CheckError
End Sub


