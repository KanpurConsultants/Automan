VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehDel 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Vehicle Delivery Challan"
   ClientHeight    =   10980
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
   ScaleHeight     =   10980
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdPost 
      Caption         =   "Re-Post"
      Height          =   375
      Left            =   9375
      TabIndex        =   150
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   147
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Fill Details"
      Height          =   6735
      Left            =   -360
      TabIndex        =   127
      Top             =   8175
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox connecteddocu 
         Height          =   300
         Left            =   4080
         TabIndex        =   16
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox kmsreading 
         Height          =   300
         Left            =   4080
         TabIndex        =   17
         Top             =   5640
         Width           =   2655
      End
      Begin VB.TextBox pdion 
         Height          =   300
         Left            =   4080
         TabIndex        =   18
         Top             =   6000
         Width           =   2655
      End
      Begin VB.TextBox deliverytakenby 
         Height          =   300
         Left            =   4080
         TabIndex        =   19
         Top             =   6360
         Width           =   2655
      End
      Begin VB.TextBox toolsandequip 
         Height          =   300
         Left            =   4080
         TabIndex        =   15
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox tyremakeE 
         Height          =   300
         Left            =   1080
         TabIndex        =   14
         Top             =   4560
         Width           =   2415
      End
      Begin VB.TextBox tyremakeD 
         Height          =   345
         Left            =   4080
         TabIndex        =   13
         Top             =   4200
         Width           =   2655
      End
      Begin VB.TextBox tyremakeC 
         Height          =   300
         Left            =   1080
         TabIndex        =   12
         Top             =   4200
         Width           =   2415
      End
      Begin VB.TextBox tyremakeB 
         Height          =   330
         Left            =   4080
         TabIndex        =   11
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox Rearaxleno 
         Height          =   300
         Left            =   4080
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox fipno 
         Height          =   300
         Left            =   4080
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox alternatorno 
         Height          =   300
         Left            =   4080
         TabIndex        =   4
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox startingmono 
         Height          =   300
         Left            =   4080
         TabIndex        =   5
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox steeringboxno 
         Height          =   300
         Left            =   4080
         TabIndex        =   6
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox accompno 
         Height          =   300
         Left            =   4080
         TabIndex        =   7
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox batteryno 
         Height          =   300
         Left            =   4080
         TabIndex        =   8
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox tyremake 
         Height          =   300
         Left            =   4080
         TabIndex        =   9
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox tyremakeA 
         Height          =   300
         Left            =   1080
         TabIndex        =   10
         Top             =   3840
         Width           =   2415
      End
      Begin VB.TextBox Transaxleno 
         Height          =   300
         Left            =   4080
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox CabNo 
         Height          =   300
         Left            =   4080
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Battery Make and No."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   148
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Taken By."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   146
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Tools and Equipments"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   145
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Connected Documents"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   144
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Kms.Reading"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   143
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "PDI On"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   142
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "E."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   141
         Top             =   4560
         Width           =   375
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "A."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   140
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "B."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   139
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "C."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   138
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "D."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   137
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Tyre Make and No."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   136
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Motor Make and No."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   135
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Steering Box Make and No."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   134
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "AC Compressor Make and No."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   133
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Alternator Make and No."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   132
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "F.I.P. Make and No."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   131
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Rear Axle Make and No."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   130
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaxle No."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   129
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cab No."
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   128
         Top             =   240
         Width           =   2175
      End
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   126
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   661
   End
   Begin VB.CommandButton CmdPrintInfo 
      BackColor       =   &H00D7C6C8&
      Caption         =   "Print Customer Sales Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6555
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   -15
      Visible         =   0   'False
      Width           =   3570
   End
   Begin MSDataGridLib.DataGrid DGConFin 
      Height          =   3750
      Left            =   8865
      Negotiate       =   -1  'True
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   10485
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6615
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
      Caption         =   "Ins. Company Help"
      ColumnCount     =   2
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
            ColumnWidth     =   2865.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   0
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
      Index           =   43
      Left            =   5115
      TabIndex        =   60
      Top             =   4755
      Width           =   1965
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
      Index           =   41
      Left            =   1845
      TabIndex        =   30
      Top             =   1155
      Width           =   360
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
      Index           =   40
      Left            =   1845
      TabIndex        =   28
      Top             =   915
      Width           =   360
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
      Height          =   210
      Index           =   42
      Left            =   2220
      TabIndex        =   31
      Top             =   1155
      Width           =   4860
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
      Index           =   39
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   116
      TabStop         =   0   'False
      Text            =   "VFa"
      Top             =   2295
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
      Height          =   210
      Index           =   38
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   115
      TabStop         =   0   'False
      Text            =   "0123456789"
      Top             =   2055
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid DGInv 
      Height          =   3000
      Left            =   -480
      Negotiate       =   -1  'True
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   10575
      Visible         =   0   'False
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   5292
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
      Caption         =   "Invoice Help"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "InvoiceNo."
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
         DataField       =   "Inv_Date"
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
      BeginProperty Column02 
         DataField       =   "Name"
         Caption         =   "Party"
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
      BeginProperty Column05 
         DataField       =   "Ord_No"
         Caption         =   "BookingNo."
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
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4454.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1769.953
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1110.047
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
      Left            =   7350
      TabIndex        =   101
      Top             =   7425
      Visible         =   0   'False
      Width           =   5025
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
         TabIndex        =   111
         Top             =   720
         Width           =   750
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
         TabIndex        =   110
         Top             =   720
         Width           =   1200
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
         TabIndex        =   109
         Top             =   300
         Visible         =   0   'False
         Width           =   375
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
         TabIndex        =   108
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
         Index           =   2
         Left            =   7425
         TabIndex        =   107
         Top             =   555
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmVehDel.frx":0000
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
         TabIndex        =   106
         ToolTipText     =   "Printer "
         Top             =   285
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmVehDel.frx":030A
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
         TabIndex        =   105
         ToolTipText     =   "Screen"
         Top             =   615
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmVehDel.frx":0614
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
         TabIndex        =   104
         ToolTipText     =   "Printer "
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1590
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
         Picture         =   "frmVehDel.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
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
         Picture         =   "frmVehDel.frx":0E4C
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Delete Current Record"
         Top             =   0
         Width           =   315
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
         TabIndex        =   114
         Top             =   0
         Width           =   4695
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
         TabIndex        =   113
         Top             =   1275
         Width           =   4650
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
         Left            =   -75
         TabIndex        =   112
         Top             =   315
         Width           =   3315
      End
      Begin VB.Line Line6 
         X1              =   2820
         X2              =   345
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line5 
         X1              =   360
         X2              =   360
         Y1              =   615
         Y2              =   720
      End
      Begin VB.Line Line7 
         X1              =   2820
         X2              =   2820
         Y1              =   630
         Y2              =   735
      End
      Begin VB.Line Line8 
         X1              =   1470
         X2              =   1470
         Y1              =   510
         Y2              =   600
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
      Height          =   210
      Index           =   37
      Left            =   5820
      TabIndex        =   27
      Top             =   675
      Width           =   1260
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
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   36
      Left            =   1845
      TabIndex        =   59
      Top             =   4755
      Width           =   1425
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
      Left            =   1845
      TabIndex        =   23
      Top             =   435
      Width           =   1425
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
      Height          =   210
      Index           =   7
      Left            =   4530
      TabIndex        =   24
      Text            =   "23/DEC/2003"
      Top             =   435
      Width           =   1260
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
      Height          =   210
      Index           =   5
      Left            =   4530
      TabIndex        =   26
      Top             =   675
      Width           =   1260
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   210
      Index           =   21
      Left            =   5655
      TabIndex        =   43
      Top             =   3075
      Width           =   1425
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
      Index           =   23
      Left            =   1845
      TabIndex        =   46
      Top             =   3555
      Width           =   480
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
      Height          =   210
      Index           =   30
      Left            =   1845
      TabIndex        =   53
      Top             =   4035
      Width           =   1425
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
      Height          =   255
      Index           =   24
      Left            =   3675
      TabIndex        =   47
      Top             =   3555
      Width           =   435
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
      Index           =   18
      Left            =   1845
      TabIndex        =   41
      Top             =   2835
      Width           =   5235
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
      Height          =   210
      Index           =   20
      Left            =   1845
      TabIndex        =   42
      Top             =   3075
      Width           =   1425
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
      Index           =   17
      Left            =   1845
      TabIndex        =   36
      Top             =   2355
      Width           =   2010
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
      Index           =   15
      Left            =   4875
      MaxLength       =   25
      TabIndex        =   40
      Top             =   2595
      Width           =   2205
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
      Left            =   1845
      TabIndex        =   37
      Top             =   2595
      Width           =   2010
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
      Index           =   14
      Left            =   5190
      MaxLength       =   20
      TabIndex        =   39
      Top             =   2355
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
      Height          =   210
      Index           =   13
      Left            =   5190
      TabIndex        =   38
      Top             =   2115
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
      Height          =   210
      Index           =   4
      Left            =   1845
      TabIndex        =   25
      Top             =   675
      Width           =   1425
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
      Index           =   12
      Left            =   1845
      TabIndex        =   35
      Text            =   " "
      Top             =   2115
      Width           =   2010
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
      Index           =   11
      Left            =   1845
      TabIndex        =   34
      Text            =   " "
      Top             =   1875
      Width           =   5235
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
      Index           =   10
      Left            =   1845
      TabIndex        =   33
      Text            =   " "
      Top             =   1635
      Width           =   5235
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
      Index           =   9
      Left            =   1845
      TabIndex        =   32
      Top             =   1395
      Width           =   5235
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
      Height          =   210
      Index           =   25
      Left            =   5115
      TabIndex        =   48
      Top             =   3555
      Width           =   525
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
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   28
      Left            =   5115
      TabIndex        =   51
      Top             =   3795
      Width           =   525
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
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   29
      Left            =   5655
      TabIndex        =   52
      Text            =   "add"
      Top             =   3795
      Width           =   1425
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
      Index           =   33
      Left            =   5655
      MaxLength       =   10
      TabIndex        =   56
      Top             =   4275
      Width           =   1425
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
      Left            =   9270
      MaxLength       =   20
      TabIndex        =   20
      Top             =   1065
      Width           =   2085
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
      Height          =   210
      Index           =   34
      Left            =   1845
      TabIndex        =   57
      Top             =   4515
      Width           =   1425
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
      Height          =   210
      Index           =   32
      Left            =   1845
      TabIndex        =   55
      Top             =   4275
      Width           =   1425
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
      Height          =   210
      Index           =   26
      Left            =   5655
      TabIndex        =   49
      Text            =   "less"
      Top             =   3555
      Width           =   1425
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
      Index           =   19
      Left            =   5655
      TabIndex        =   45
      Top             =   3315
      Width           =   1425
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
      Height          =   210
      Index           =   27
      Left            =   1845
      TabIndex        =   50
      Top             =   3795
      Width           =   480
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
      Height          =   210
      Index           =   31
      Left            =   5655
      MaxLength       =   14
      TabIndex        =   54
      Top             =   4035
      Width           =   1425
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
      Index           =   3
      Left            =   10380
      MaxLength       =   8
      TabIndex        =   22
      Top             =   1545
      Width           =   975
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
      Height          =   210
      Index           =   8
      Left            =   2220
      TabIndex        =   29
      Top             =   915
      Width           =   4860
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
      Height          =   210
      Index           =   35
      Left            =   5655
      TabIndex        =   58
      Top             =   4515
      Width           =   1425
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
      Height          =   210
      Index           =   22
      Left            =   1845
      TabIndex        =   44
      Top             =   3315
      Width           =   1425
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
      Left            =   10095
      MaxLength       =   12
      TabIndex        =   21
      Top             =   1305
      Width           =   1260
   End
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   -2130
      Negotiate       =   -1  'True
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   10080
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
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
      Caption         =   "Voucher No"
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
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2175
      Left            =   -1155
      Negotiate       =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   10590
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
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
      Caption         =   "Site Help"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1305
      Left            =   150
      TabIndex        =   120
      Top             =   5025
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   2302
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   9
      BackColorFixed  =   12243913
      ForeColorFixed  =   128
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      FocusRect       =   0
      Appearance      =   0
      FormatString    =   "kkk"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   2100
      Left            =   7320
      TabIndex        =   121
      Top             =   2805
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   3704
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   4
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      GridColorFixed  =   192
      GridColorUnpopulated=   16761024
      FocusRect       =   0
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "SrNo."
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
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
      Index           =   0
      Left            =   8640
      MaxLength       =   21
      TabIndex        =   149
      Top             =   600
      Width           =   2805
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
      Left            =   180
      TabIndex        =   151
      Top             =   6630
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration By :"
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
      Index           =   20
      Left            =   3675
      TabIndex        =   123
      Top             =   4770
      Width           =   1440
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check List Items"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   18
      Left            =   7335
      TabIndex        =   122
      Top             =   2535
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name"
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
      Index           =   19
      Left            =   135
      TabIndex        =   119
      Top             =   1170
      Width           =   1230
   End
   Begin VB.Label LblAcPostDt 
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7740
      TabIndex        =   118
      Top             =   2310
      Width           =   405
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7740
      TabIndex        =   117
      Top             =   2070
      Width           =   1245
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
      Index           =   25
      Left            =   3915
      TabIndex        =   100
      Top             =   2370
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
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
      TabIndex        =   99
      Top             =   2580
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp. Del. Dt."
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
      Index           =   18
      Left            =   5820
      TabIndex        =   97
      Top             =   435
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Payable Amt"
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
      Index           =   4
      Left            =   150
      TabIndex        =   95
      Top             =   4770
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Y/N"
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
      TabIndex        =   93
      Top             =   3570
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Fee"
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
      Index           =   45
      Left            =   150
      TabIndex        =   92
      Top             =   4050
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rebate Days"
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
      Index           =   44
      Left            =   2400
      TabIndex        =   91
      Top             =   3570
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financier"
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
      Left            =   150
      TabIndex        =   90
      Top             =   2850
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financed Amount"
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
      TabIndex        =   89
      Top             =   3090
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RTO Office"
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
      Left            =   150
      TabIndex        =   88
      Top             =   2370
      Width           =   915
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
      Index           =   27
      Left            =   3915
      TabIndex        =   87
      Top             =   2595
      Width           =   915
   End
   Begin VB.Label Label3 
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
      Index           =   24
      Left            =   3915
      TabIndex        =   86
      Top             =   2130
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No."
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
      Index           =   23
      Left            =   150
      TabIndex        =   85
      Top             =   675
      Width           =   1035
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
      Index           =   21
      Left            =   150
      TabIndex        =   84
      Top             =   2130
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
      Index           =   14
      Left            =   150
      TabIndex        =   83
      Top             =   1410
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Bill No.*"
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
      Left            =   150
      TabIndex        =   82
      Top             =   435
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Bill Date "
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
      Left            =   3300
      TabIndex        =   81
      Top             =   435
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Charges"
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
      Left            =   150
      TabIndex        =   80
      Top             =   4530
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Amount"
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
      Left            =   3675
      TabIndex        =   79
      Top             =   3090
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less TDS @"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   17
      Left            =   3675
      TabIndex        =   78
      Top             =   3810
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cover Note No."
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
      Index           =   16
      Left            =   3675
      TabIndex        =   77
      Top             =   4290
      Width           =   1320
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   15
      Left            =   7890
      TabIndex        =   76
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Charges"
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
      Index           =   13
      Left            =   150
      TabIndex        =   75
      Top             =   4290
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest @"
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
      Index           =   12
      Left            =   4170
      TabIndex        =   74
      Top             =   3570
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Paid"
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
      Index           =   32
      Left            =   150
      TabIndex        =   73
      Top             =   3330
      Width           =   1155
   End
   Begin VB.Label LblPaybleAmt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payable Amount"
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
      Height          =   195
      Left            =   3675
      TabIndex        =   72
      Top             =   3330
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Y/N"
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
      Left            =   150
      TabIndex        =   71
      Top             =   3810
      Width           =   720
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
      Index           =   9
      Left            =   3675
      TabIndex        =   70
      Top             =   4050
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   1290
      Left            =   7680
      Top             =   555
      Width           =   3780
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   9675
      TabIndex        =   69
      Top             =   1575
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Doc.  No."
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
      Index           =   1
      Left            =   7890
      TabIndex        =   68
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division           "
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
      Left            =   7890
      TabIndex        =   67
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code    "
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
      Left            =   9945
      TabIndex        =   66
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Date"
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
      Index           =   43
      Left            =   3300
      TabIndex        =   65
      Top             =   675
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
      Left            =   7890
      TabIndex        =   64
      Top             =   608
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stamp / Duty Charges"
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
      Left            =   3675
      TabIndex        =   63
      Top             =   4530
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
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
      TabIndex        =   62
      Top             =   930
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
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
      Index           =   2
      Left            =   7890
      TabIndex        =   61
      Top             =   1320
      Width           =   1185
   End
End
Attribute VB_Name = "frmVehDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsVno As ADODB.Recordset
Dim RsSite As ADODB.Recordset
Dim RsInv As ADODB.Recordset
Dim RsConFin As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim DocID As String '* 21
Dim InvdocId As String '* 21
Public mVType As String
Dim VoucherEditFlag As Boolean
Dim vPrefix As String
Private Const TxtDocID As Byte = 0
Private Const SiteCode As Byte = 1
Private Const VDate As Byte = 2
Private Const SerialNo As Byte = 3
Private Const BookNo As Byte = 4
Private Const BookDate As Byte = 5
Private Const InvNo As Byte = 6
Private Const InvDate As Byte = 7
Private Const Party As Byte = 8
Private Const Add1 As Byte = 9
Private Const Add2 As Byte = 10
Private Const Add3 As Byte = 11
Private Const City As Byte = 12
Private Const Model As Byte = 13
Private Const ChassisNo As Byte = 14
Private Const EngineNo As Byte = 15
Private Const Colours As Byte = 16
Private Const RTO  As Byte = 17
Private Const FB_Code  As Byte = 18
Private Const PayAmt As Byte = 19
Private Const FinAmt  As Byte = 20
Private Const VehAmt As Byte = 21
Private Const AdvAmt As Byte = 22
Private Const IntYN As Byte = 23
Private Const RebDays As Byte = 24
Private Const IntPer As Byte = 25
Private Const IntAmt As Byte = 26
Private Const TDSYN As Byte = 27
Private Const TDSPer As Byte = 28
Private Const TDSAmt As Byte = 29
Private Const RegFee As Byte = 30
Private Const RegNo As Byte = 31
Private Const IncChg As Byte = 32
Private Const CoverNote As Byte = 33
Private Const SerChg As Byte = 34
Private Const StampChg As Byte = 35
Private Const NetPayAmt As Byte = 36
Private Const ExpDate As Byte = 37
Private Const AcPostByName As Byte = 38
Private Const AcPostDate As Byte = 39
Private Const NamePrefix As Byte = 40
Private Const FNamePrefix As Byte = 41
Private Const fname As Byte = 42
Private Const RegBy As Byte = 43

Private Const Col_DocID As Byte = 1
Private Const Col_VNo As Byte = 2
Private Const Col_Date As Byte = 3
Private Const Col_VType As Byte = 4
Private Const Col_Amt As Byte = 5
Private Const Col_DrCr As Byte = 6
Private Const Col_IntDays As Byte = 7
Private Const Col_IntAmt As Byte = 8

Private Const ItemCode As Byte = 1
Private Const Description As Byte = 2
Private Const DefVal As Byte = 3
Private Const PIndex As Byte = 4

Private Const SiteCode1 As Byte = 0
Private Const FromVno As Byte = 1
Private Const ToVno As Byte = 2

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String

Private Sub CmdFill_Click()
If Frame1.Visible = False Then
    Frame1.Visible = True
    CabNo.SetFocus
Else
    Frame1.Visible = False
End If
End Sub

Private Sub cmdPost_Click()
Dim I As Integer, mStartdate As Date, mEndDate As Date
    mStartdate = InputBox("Posting Required from which Date ?", "Start Date for Posting", PubLoginDate)
    mEndDate = InputBox("Posting Required upto which Date ?", "Last Date for Posting", PubLoginDate)

    If Master.RecordCount > 0 Then Call TopCtrl1_eFirst
    Do Until Master.EOF
        Call MoveRec
        
        If IsNull(Master!DelCh_UEntDt) Then GoTo MyNextRecord
        If Master!DelCh_UEntDt < CDate(mStartdate) Then GoTo MyNextRecord
        If Master!DelCh_UEntDt > CDate(mEndDate) Then GoTo MyNextRecord
        
        For I = 0 To Txt.Count - 1
            Txt(I).Refresh
            
        Next
        If AcPostAuthorisation(Txt(AcPostByName)) = False Then GoTo MyNextRecord
        Disp_Text SETS("EDIT", Me, Master)
        If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
            If CDate(Txt(VDate).TEXT) >= PubStartDate And CDate(Txt(VDate).TEXT) <= PubEndDate Then ProcAcPost
        End If
        Disp_Text SETS("INI", Me, Master)

MyNextRecord:
        Master.MoveNext
    Loop
End Sub
Private Sub CmdPrintInfo_Click()
Dim Rst As ADODB.Recordset, RstSub1 As ADODB.Recordset, RstSub2 As ADODB.Recordset
Dim RstSub3 As ADODB.Recordset, mQry As String
Dim I As Integer, Rst2, RST3 As ADODB.Recordset
On Error GoTo ERRORHANDLER
    'If IsValid(txtPrint(Model), "Model") = False Then Exit Sub
    'If IsValid(txtPrint(ChasNo), "Chassis No") = False Then Exit Sub
    mQry = "SELECT VO.STAMP_DUTY,VO.model,VO.REG_FEE, VO.INS_FEE, VO.S_CHARGE, VO.Net_AMOUNT, " & _
        " VO.MISC_INFO, Godown.God_Name, City.CityName, City_2.CityName, " & _
        " Model.Model_Desc, Model.Model_Desc1, CF.FinName,CF.Add1, CF.Add2, " & _
        " CF.PinCode,VP1.V_Date as PurDt,VP1.V_NO as PurVno, VP1.Tot_Amount, " & _
        " SubGroup.Name, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, SubGroup.PIN, VStk.Pur_DocId, " & _
        " VStk.ChassisNo, VStk.EngineNo, VStk.Chassis_RctDocNo, VStk.Chassis_RctDate, VStk.AL_Name, " & _
        " SubGroup_1.Name as Supplier, VP1.PBILL_NO, VP1.PBILL_DATE, VP1.V_NO, VP1.V_Date, " & _
        " VO.OrdDocId, VO.Ord_Date, Emp_Mast.Emp_Name, VO.Inv_DocId, VO.Inv_Date, VO.VRATE, VO.MARGINE, " & _
        " VO.Transport, VO.OtherChrg, VO.TAX_Amt, VO.Surcharge_Amt, VO.FIN_AMT, VO.Interest, VO.DelCh_DocId, VO.DelCh_DT,VO.AdvEMI,VO.OtherChrg,VO.Rebate,SubGroup.FName " & _
        " FROM (((((((((Veh_Stock VStk LEFT JOIN Veh_Purch1 VP1 ON VStk.Pur_DocId = VP1.DocID) " & _
        " LEFT JOIN Veh_Order VO on VStk.Sal_DocId=VO.Inv_DocId) " & _
        " LEFT JOIN ContractFinance CF on VO.FB_CODE=CF.FinCode) " & _
        " LEFT JOIN City AS City_2 ON CF.City = City_2.CityCode) " & _
        " LEFT JOIN Model ON VO.MODEL = Model.MODEL) " & _
        " LEFT JOIN SubGroup ON VO.PartyCode = SubGroup.SubCode) " & _
        " LEFT JOIN City ON SubGroup.CityCode = City.CityCode) " & _
        " LEFT JOIN Godown ON VStk.Godown = Godown.God_Code) " & _
        " LEFT JOIN SubGroup AS SubGroup_1 ON VP1.PARTYCODE = SubGroup_1.SubCode) " & _
        " LEFT JOIN Emp_Mast ON VO.REP_CODE = Emp_Mast.Emp_Code " & _
        " where VStk.chassisno ='" & Txt(ChassisNo).TEXT & "' and Vstk.model ='" & Txt(Model) & "'"

    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub




   




   'Recordset is made for subreport2

    mQry = "SELECT Rect.Prov_No,Rect.Ord_DocId, Rect.V_Type, Rect.V_No, Rect.V_Date, Rect.Site_Code, Rect.AMOUNT, Rect.DrCr, Rect.Narration " & _
        " FROM Veh_Order " & _
        " LEFT JOIN Rect ON Veh_Order.OrdDocId = Rect.Ord_DocId " & _
        " where Veh_Order.chassis ='" & Txt(ChassisNo).TEXT & "' and  Veh_Order.model ='" & Txt(Model) & "'"

   Set RstSub2 = New Recordset
   RstSub2.CursorLocation = adUseClient
   RstSub2.Open (mQry), GCn, adOpenDynamic, adLockOptimistic






    



    mRepName = "CustSalesInfor"

    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True

    CreateFieldDefFile RstSub2, PubRepoPath + "\" & mRepName & "2.ttx", True

    

    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    rpt.Database.SetDataSource Rst

    rpt.OpenSubreport("SUBREP2").Database.SetDataSource RstSub2

    

    Set Rst2 = New ADODB.Recordset
    Rst2.CursorLocation = adUseClient
    Rst2.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic

    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("SubTitle")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecSpeciality & "'"
            Case UCase("LST")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecLST & "'"
            Case UCase("LSTDate")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecLST_Date & "'"
            Case UCase("CST")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecCST & "'"
            Case UCase("CSTDate")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecCST_Date & "'"
            Case UCase("Phone")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecFax & "'"
            Case UCase("Gram")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecGram & "'"
        End Select
    Next

    
        For I = 1 To rpt.OpenSubreport("SUBREP2").FormulaFields.Count
            Select Case UCase(rpt.OpenSubreport("SUBREP2").FormulaFields(I).FormulaFieldName)
                Case UCase("SpeedPrint")
                    rpt.OpenSubreport("SUBREP1").FormulaFields(I).TEXT = "'1'"
            End Select
        Next
'        For I = 1 To rpt.OpenSubreport("SUBREP2").FormulaFields.Count
'            Select Case UCase(rpt.OpenSubreport("SUBREP2").FormulaFields(I).FormulaFieldName)
'                Case UCase("SpeedPrint")
'                    rpt.OpenSubreport("SUBREP2").FormulaFields(I).TEXT = "'1'"
'            End Select
'        Next
'        For I = 1 To rpt.OpenSubreport("SUBREP3").FormulaFields.Count
'            Select Case UCase(rpt.OpenSubreport("SUBREP3").FormulaFields(I).FormulaFieldName)
'                Case UCase("SpeedPrint")
'                    rpt.OpenSubreport("SUBREP3").FormulaFields(I).TEXT = "'1'"
'            End Select
'        Next
    

    rpt.ReadRecords

      'Printer
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
                    rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
            End Select
            Next
        '    rpt.PrintOut False
        'Case PScreen  'screen
            Call Report_View(rpt, Me.CAPTION, , True)
        'Case PDos
        '    Call Report_View(rpt, Me.CAPTION, 1)

Set Rst = Nothing
Set Rst2 = Nothing
Set RST3 = Nothing
Set rpt = Nothing
CmdPrint(PSetUp).Tag = ""
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub DGSite_Click()
If FrmPrn.Visible = False Then
    DGSite.Visible = False
    If RsSite.RecordCount > 0 Then
        Txt(SiteCode).TEXT = RsSite!Name
        Txt(SiteCode).Tag = RsSite!Code
    End If
    Txt(SiteCode).SetFocus
Else
    DGSite.Visible = False
    If RsSite.RecordCount > 0 Then
        txtPrint(SiteCode1).TEXT = RsSite!Name
        txtPrint(SiteCode1).Tag = RsSite!Code
    End If
    txtPrint(SiteCode1).SetFocus
End If
End Sub

Private Sub DGInv_Click()
    DGInv.Visible = False
    If RsInv.RecordCount > 0 Then
        Txt(InvNo).TEXT = RsInv!Code
        FillRecords RsInv
    End If
    Txt(InvNo).SetFocus
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
TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid
    DGInv.left = Me.left: DGInv.top = Txt(RTO).top: DGInv.width = Me.width - 90
    DGSite.left = 5145: DGSite.top = mTopScale
    DGVno.left = 5145: DGVno.top = mTopScale
    FrmPrn.left = 525: FrmPrn.top = 2220
    
    mVType = "V_DCL"
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
    Dim sitecond As String
    sitecond = " And DelCh_DT Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("Veh_Order.DelCh_DocId", "3", "1") & " ='" & PubSiteCode & "'"
    End If
    
    If PubMoveRecYn Then
        Master.Open "select DelCh_DocId as searchcode,Veh_Order.*,ContractFinance.FinName from Veh_Order Left Join ContractFinance on Veh_Order.RegBy=ContractFinance.FinCode where left(DelCh_DocId,1)='" & PubDivCode & "' and (DelCh_DocId <> '' Or DelCh_DocId Is Not Null) and  DelCh_VType = '" & mVType & "' " & sitecond & " order by DelCh_DocId desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Set Master = GCn.Execute("select Top 1 DelCh_DocId as searchcode,Veh_Order.*,ContractFinance.FinName from Veh_Order Left Join ContractFinance on Veh_Order.RegBy=ContractFinance.FinCode where left(DelCh_DocId,1)='" & PubDivCode & "' and (DelCh_DocId <> '' Or DelCh_DocId Is Not Null) and  DelCh_VType = '" & mVType & "' " & sitecond & "Order by DelCh_DocId desc")
    End If
    
    Set RsSite = New ADODB.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select site_code as code,site_desc as name from site order by site_desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGSite.DataSource = RsSite
    
'    Set RsInv = New ADODB.Recordset
'    RsInv.CursorLocation = adUseClient
'    RsInv.Open "SELECT trim(str(Veh_Order.Inv_No)) as code,Veh_Order.MODEL, Veh_Order.Chassis, Veh_Order.PartyCode, Veh_Order.EXP_DATE, Veh_Order.FB_CODE, Veh_Order.FIN_AMT, Veh_Order.Colour_Code, Veh_Order.Inv_DocId, Veh_Order.Ord_No, Veh_Order.Ord_Date, Veh_Order.Inv_SiteCode, Veh_Order.Inv_Date, Veh_Order.Inv_No ,Veh_Order.net_amount " & _
'        "FROM Veh_Order where Veh_Order.inv_DocId <> '' and Veh_Order.DelCh_DocId ='' order by Inv_No", GCn, adOpenDynamic, adLockOptimistic
'    Set DGInv.DataSource = RsInv

    'MODI SHEKHAR 23 Jan
    Set RsInv = GCn.Execute("SELECT " & cTrim(CStr("Veh_Order.Inv_No")) & " as code,Veh_Order.MODEL, Veh_Order.Chassis, Veh_Order.PartyCode, SubGroup.Name,Veh_Order.exp_DATE, Veh_Order.FB_CODE, Veh_Order.FIN_AMT, Veh_Order.Colour_Code, Veh_Order.Inv_DocId, Veh_Order.OrdDocId, Veh_Order.Ord_No, Veh_Order.Ord_Date, Veh_Order.Inv_SiteCode, Veh_Order.Inv_Date, Veh_Order.Inv_No ,Veh_Order.net_amount,Veh_Order.RTO " & _
        "FROM Veh_Order Left Join SubGroup on Veh_Order.PartyCode=SubGroup.SubCode " & _
        "where left(Veh_Order.inv_DocId,1)='" & PubDivCode & "' and Veh_Order.inv_DocId <> '' And Veh_Order.Inv_DocId Is Not Null  and (Veh_Order.DelCh_DocId ='' Or Veh_Order.DelCh_DocId Is Null) order by Inv_No")
    Set DGInv.DataSource = RsInv
    'END modi
    
    Set RsVno = New ADODB.Recordset
    RsVno.CursorLocation = adUseClient
    RsVno.Open "Select distinct DelCh_No as code from Veh_Order where left(Veh_Order.DelCh_DocId,1)='" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
    Set DGVno.DataSource = RsVno

    Set RsConFin = New ADODB.Recordset
    RsConFin.CursorLocation = adUseClient
    RsConFin.Open "select FinCode as code,FinName as Name from ContractFinance where FinCatg=3 order by FinName ", GCn, adOpenDynamic, adLockOptimistic
    Set DGConFin.DataSource = RsConFin
    
    Call MoveRec
    Disp_Text SETS("INI", Me, Master)
    If UCase(left(PubComp_Name, 7)) = "SOCIETY" Then
        CmdPrintInfo.Visible = True
    End If
    
    If UCase(left(PubComp_Name, 7)) = "JOHNSON" Then
        CmdFill.Visible = True
        'CabNo.SetFocus
    End If
    
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
Set RsInv = Nothing
Set Master = Nothing
Set RsVno = Nothing
Set RsSite = Nothing
End Sub

Private Sub Text1_Change()

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
        Txt(SiteCode).Tag = PubSiteCode
        Txt(SiteCode) = PubSiteName
        Txt(VDate).SetFocus
    Else
        Txt(SiteCode).SetFocus
    End If
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim I As Integer
Dim LedgAry(1) As LedgRec, mResult As Byte

If AcPostAuthorisation(Txt(AcPostByName)) = False Then Exit Sub

If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
    GCn.BeginTrans: GCnFaV.BeginTrans
    'Unpost Ledger a/c
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaV, Txt(TxtDocID))
    If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
    'Unposting of Ledger completed
    
    GCn.Execute "Update Veh_Stock set DelCh_DocId= '',DelCh_Date=Null " & _
        "where Sal_docid = '" & InvdocId & "'"
        
    GCn.Execute "update veh_order set DelCh_DocId = '' , DelCh_DocIDHelp='', DelCh_SiteCode='', DelCh_VType='', DelCh_No=0, DelCh_DT=Null, " & _
    "Interest_YN=0,RebDays=0,InterestPer=0,Interest=0,TDS_YN=0,TDS_Per=0,TDS_Amt=0,  " & _
    "REG_FEE=0,REG_NO='',INS_FEE=0,Ins_NOTE='',S_CHARGE=0,STAMP_DUTY=0,  " & _
    "DelCh_UName='',DelCh_UEntDt=null,DelCh_UAE='',DelCh_AcPostByUName='',DelCh_AcPostByUEntDt=Null " & _
    " where inv_docid = '" & InvdocId & "'"
    
    GCnFaV.CommitTrans
    GCn.CommitTrans
    
    Master.Requery
    RsInv.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
Exit Sub
eloop1:
    If err.NUMBER <> 0 Then GCn.RollbackTrans: GCnFaV.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
    If AcPostAuthorisation(Txt(AcPostByName)) = False Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    Txt(IntYN).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
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
OptPlain.Value = True
LblPrinter.CAPTION = Printer.DeviceName
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus
End Sub

Private Sub TopCtrl1_eRef()
    RsInv.Requery
    RsSite.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim Rst As ADODB.Recordset
    Dim mTrans As Boolean
    Dim DocIdHlp$, mNarr$
    Dim mFundSource As Byte
    Dim mTrntypeprn As Byte
   ' On Error GoTo errlbl


    If IsEditable(RetDate(Txt(VDate))) = False Then Exit Sub
    Grid_Hide
    If IsValid(Txt(SiteCode), "SiteCode") = False Then Exit Sub
    If IsValid(Txt(VDate), "Challan Date") = False Then Exit Sub
    If IsValid(Txt(SerialNo), "Challan Number") = False Then Exit Sub
    If IsValid(Txt(InvNo), "Invoice No.") = False Then Exit Sub
    Amt_Cal
    '********* cHECKING pOSTING cOTROLS
    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        If ProcAcPost(True) = False Then Me.ActiveControl.SetFocus: Exit Sub
        Txt(AcPostByName) = pubUName
        Txt(AcPostDate) = PubServerDate
    End If
    '**********
    GCn.BeginTrans
    GCnFaV.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2 = "Add" Then
        DocID = Txt(TxtDocID)
        If GCn.Execute("select count(*) from veh_order where Left(DelCh_DocId,1)='" & PubDivCode & "'And DelCh_VType = '" & mVType & "' And DelCh_No=" & Val(Txt(SerialNo)) & " ").Fields(0) > 0 Then
            If VoucherEditFlag Then 'And Txt(SerialNo).Visible Then
                MsgBox "Delivery Document No. already exists, Retry", vbCritical, "Validation Error"
                Txt(SerialNo).SetFocus
                GoTo errlbl
            Else
                Txt(TxtDocID) = GetDocID(GCnFaV, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                If Val(Txt(SerialNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
                    MsgBox "Delivery Document No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo errlbl
                End If
            End If
        End If
        DocIdHlp = Replace(Txt(TxtDocID), " ", "")
        
        GCn.Execute "update veh_order set DelCh_DocId = '" & Txt(TxtDocID) & "' , DelCh_DocIDHelp='" & DocIdHlp & "', DelCh_SiteCode='" & PubSiteCode & Txt(SiteCode).Tag & "', DelCh_VType='" & mVType & "', DelCh_No=" & Val(Txt(SerialNo).TEXT) & ", DelCh_DT=" & ConvertDate(Txt(VDate)) & ", " & _
            "Interest_YN=" & IIf(Txt(IntYN) = "Yes", 1, 0) & ",RebDays=" & Val(Txt(RebDays)) & ",InterestPer=" & Val(Txt(IntPer)) & ",Interest=" & Val(Txt(IntAmt)) & ",TDS_YN=" & IIf(Txt(TDSYN) = "Yes", 1, 0) & ",TDS_Per=" & Val(Txt(TDSPer)) & ",TDS_Amt=" & Val(Txt(TDSAmt)) & ",  " & _
            "REG_FEE=" & Val(Txt(RegFee)) & ",REG_NO='" & Txt(RegNo) & "',INS_FEE=" & Val(Txt(IncChg)) & ",Ins_NOTE='" & Txt(CoverNote) & "',S_CHARGE=" & Val(Txt(SerChg)) & ",STAMP_DUTY=" & Val(Txt(StampChg)) & ",  " & _
            "DelCh_UName='" & pubUName & "',DelCh_UEntDt=" & ConvertDate(PubServerDate) & ",DelCh_UAE='A', " & _
            "DelCh_AcPostByUName='" & Txt(AcPostByName) & "',DelCh_AcPostByUEntDt=" & ConvertDate(Txt(AcPostDate)) & _
            ", DelCh_AddBy = '" & pubUName & "', DelCh_AddDate = " & ConvertDateTime(PubServerDate) & ",RegBy='" & Txt(RegBy).Tag & "' where inv_docid = '" & InvdocId & "'"
        '17.07.03 LPS
        GCn.Execute "update Veh_Stock set DelCh_DocId = '" & Txt(TxtDocID) & "',DelCh_Date=" & ConvertDate(Txt(VDate)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & _
            " where Sal_DocId = '" & InvdocId & "'"
        'eof LPS
        'Voucher Serial No. Updation LPS 21-05-03
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaV, Txt(TxtDocID), Txt(VDate)
    Else
        GCn.Execute "update veh_order set  " & _
            "Interest_YN=" & IIf(Txt(IntYN) = "Yes", 1, 0) & ",RebDays=" & Val(Txt(RebDays)) & ",InterestPer=" & Val(Txt(IntPer)) & ",Interest=" & Val(Txt(IntAmt)) & ",TDS_YN=" & IIf(Txt(TDSYN) = "Yes", 1, 0) & ",TDS_Per=" & Val(Txt(TDSPer)) & ",TDS_Amt=" & Val(Txt(TDSAmt)) & ", " & _
            "REG_FEE=" & Val(Txt(RegFee)) & ",REG_NO='" & Txt(RegNo) & "',INS_FEE=" & Val(Txt(IncChg)) & ",Ins_NOTE='" & Txt(CoverNote) & "',S_CHARGE=" & Val(Txt(SerChg)) & ",STAMP_DUTY=" & Val(Txt(StampChg)) & ",  " & _
            "DelCh_UName='" & pubUName & "',DelCh_UEntDt=" & ConvertDate(PubServerDate) & ",DelCh_UAE='E', " & _
            "DelCh_AcPostByUName='" & Txt(AcPostByName) & "',DelCh_AcPostByUEntDt=" & ConvertDate(Txt(AcPostDate)) & _
            ",DelCh_ModifyBy = '" & pubUName & "', DelCh_ModifyDate = " & ConvertDateTime(PubServerDate) & ",RegBy='" & Txt(RegBy).Tag & "' where DelCh_DocId = '" & Txt(TxtDocID) & "'"
    End If
    For I = 1 To FGrid.Rows - 1
        GCn.Execute ("Update Rect set IntValue=" & Val(FGrid.TextMatrix(I, Col_IntAmt)) & ", IntDays=" & Val(FGrid.TextMatrix(I, Col_IntDays)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' where DocId= '" & FGrid.TextMatrix(I, Col_DocID) & "'")
    Next I
    
    'A/c Posting
    If PubAcPostingByAllUser Or (PubAcPostingByAllUser = False And pubUAcPosting = "Y") Then
        ProcAcPost
    End If
    'EOF of A/c Posting Section
GCnFaV.CommitTrans
GCn.CommitTrans
mTrans = False
Set Rst = Nothing
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select DelCh_DocId as searchcode,Veh_Order.*,ContractFinance.FinName from Veh_Order Left Join ContractFinance on Veh_Order.RegBy=ContractFinance.FinCode where left(DelCh_DocId,1)='" & PubDivCode & "' and (DelCh_DocId <> '' Or DelCh_DocId Is Not Null) and  DelCh_VType = '" & mVType & "' And DelCh_DocId = '" & Txt(TxtDocID) & "' Order by DelCh_DocId desc")
    End If
    RsInv.Requery
    Master.FIND "DelCh_DocId = '" & Txt(TxtDocID) & "'"
    'lp 11-03-03
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        'If Val(Txt(SerialNo)) > DeCodeDocID(DocId, Document_No) Then
        '    MsgBox "Delivery Document No." & Trim(DeCodeDocID(DocId, Document_No)) & " already exists ! " & vbCrLf & "New No. " & Txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        'End If
    End If
    TopCtrl1_ePrn
    Exit Sub
errlbl:
    If mTrans Then GCn.RollbackTrans: GCnFaV.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
     Dim sitecond As String
     sitecond = " And DelCh_DT Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and " & cMID("Veh_Order.DelCh_DocId", "3", "1") & " ='" & PubSiteCode & "'"
    End If
    
    GSQL = "select Inv_No as searchcode,MODEL," & cCStr("Inv_No", 10) & " As Inv_No," & cDt("Inv_Date") & "  as Inv_Date, " & cCStr("DelCh_No") & " As DelCh_No, " & cDt("DelCh_DT") & " As DelCh_Date,Chassis,Srv_BookNo,DelCh_DocId from Veh_Order where left(DelCh_DocId,1)='" & PubDivCode & "' " & sitecond & "" ' DelCh_DocId = '" & mVType & "'"
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
        Master.FIND ("Inv_No='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("select DelCh_DocId as searchcode,Veh_Order.*,ContractFinance.FinName from Veh_Order Left Join ContractFinance on Veh_Order.RegBy=ContractFinance.FinCode where left(DelCh_DocId,1)='" & PubDivCode & "' and (DelCh_DocId <> '' Or DelCh_DocId Is Not Null) and  DelCh_VType = '" & mVType & "' And DelCh_DocId = '" & MyValue & "' Order by DelCh_DocId desc")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub


Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus Txt(Index)
Grid_Hide
Select Case Index
    Case SiteCode
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Then Exit Sub
        If Txt(Index).TEXT = "" Then
            RsSite.MoveFirst
            RsSite.FIND "code ='" & PubSiteCode & "'"
            Txt(Index).Tag = RsSite!Code
            Txt(Index).TEXT = RsSite!Name
        Else
            If Txt(Index).TEXT <> RsSite!Name Then
                RsSite.MoveFirst
                RsSite.FIND "name ='" & Txt(Index).TEXT & "'"
            End If
        End If
    Case InvNo
        '************MIDI SHEKHAR 23 Jan
        If IsValid(Txt(VDate), "Challan Date") = False Then Exit Sub
        Set RsInv = GCn.Execute("SELECT " & cTrim(CStr("Veh_Order.Inv_No")) & " as code,Veh_Order.MODEL, Veh_Order.Chassis, Veh_Order.PartyCode, SubGroup.Name, Veh_Order.exp_DATE, Veh_Order.FB_CODE, Veh_Order.FIN_AMT, Veh_Order.Colour_Code, Veh_Order.Inv_DocId, Veh_Order.OrdDocId, Veh_Order.Ord_No, Veh_Order.Ord_Date, Veh_Order.Inv_SiteCode, Veh_Order.Inv_Date, Veh_Order.Inv_No ,Veh_Order.net_amount,Veh_Order.RTO " & _
        "FROM Veh_Order Left Join SubGroup on Veh_Order.PartyCode=SubGroup.SubCode " & _
        "where left(Veh_Order.inv_DocId,1)='" & PubDivCode & "' and Veh_Order.inv_DocId <> '' And Veh_Order.Inv_DocId Is Not Null And (Veh_Order.DelCh_DocId ='' Or Veh_Order.DelCh_DocId Is Null) and Veh_Order.Inv_Date <= " & ConvertDate(Txt(VDate)) & " order by Inv_No")
        Set DGInv.DataSource = RsInv
        RsInv.Sort = "Code"
        '**********MODI END
        If RsInv.RecordCount = 0 Or (RsInv.EOF = True Or RsInv.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
            If Txt(Index).TEXT <> RsInv!Code Then
                RsInv.MoveFirst
                RsInv.FIND "code ='" & Txt(Index).TEXT & "'"
            End If
'    Case IntPer, IntAmt, TDSPer, TDSAmt, RegFee, IncChg, SerChg, StampChg, RebDays
'         SendKeys "{HOME}+{END}"
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
    Case SiteCode
        DGridTxtKeyDown DGSite, Txt, Index, RsSite, KeyCode, False, 1
    Case InvNo
        DGridTxtKeyDown DGInv, Txt, Index, RsInv, KeyCode, False, 0
    Case RegBy
        DGridTxtKeyDown DGConFin, Txt, Index, RsConFin, KeyCode, False, 1
End Select
If DGInv.Visible = False And DGSite.Visible = False And DGConFin.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VDate Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> RegBy Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = RegBy Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> SiteCode Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> IntYN Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case SiteCode
        If DGSite.Visible = True Then DGridTxtKeyPress Txt, Index, RsSite, KeyAscii, "Name"
    Case InvNo
        If DGInv.Visible = True Then DGridTxtKeyPress Txt, Index, RsInv, KeyAscii, "Code"
    Case RegBy
        If DGConFin.Visible = True Then DGridTxtKeyPress Txt, Index, RsConFin, KeyAscii, "Name"
    Case SerialNo
        Call NumPress(Txt(Index), KeyAscii, 6, 0)
    Case IntPer, IntAmt, TDSPer, TDSAmt, RegFee, IncChg, SerChg, StampChg
        Call NumPress(Txt(Index), KeyAscii, 8, 2)
    Case RebDays
        Call NumPress(Txt(Index), KeyAscii, 3, 0)
    Case IntYN
        If UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index) = "Yes"
            EnableText True
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            Txt(Index) = "No"
            EnableText False
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            Txt(Index) = ""
        End If
        KeyAscii = 0
    Case TDSYN
        If UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index) = "Yes"
        ElseIf UCase(Chr(KeyAscii)) = "N" Then
            Txt(Index) = "No"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            Txt(Index) = ""
        End If
        KeyAscii = 0
End Select

'KeyAscii = RetDGKeyAscii()
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs

Select Case Index
    Case RebDays
        FillCustTrns Txt(BookNo).Tag, False, True, Val(Txt(IntPer)), Val(Txt(RebDays)), Txt(InvDate)
    Case IntPer
        FillCustTrns Txt(BookNo).Tag, False, True, Val(Txt(IntPer)), Val(Txt(RebDays)), Txt(InvDate)
        Amt_Cal
    Case TDSPer
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = 16 Then Exit Sub
'        If mInterest  >= mTDSExAmt Then
'    mTDSYN='Y'
'    mTDSRATE = mPTDSRATE
'    mTDS = Round(mInterest * mTDSRATE / 100, 0)
'Else
'    mTDSYN='N';mTDSRATEE=0;mTDS=0
'End If
        Txt(TDSAmt).TEXT = Format(Val(Txt(TDSPer).TEXT) * Val(Txt(IntAmt).TEXT) / 100, "0.00")
        Amt_Cal
    Case IntAmt, TDSAmt, RegFee, IncChg, SerChg, StampChg
         Amt_Cal
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim I As Integer
Select Case Index
     Case SiteCode
        If IsValid(Txt(Index), "Site Code") = False Then Cancel = True: Exit Sub
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsSite!Name
            Txt(Index).Tag = RsSite!Code
        End If
    Case InvNo
        If IsValid(Txt(Index), "Invoice No.") = False Then Cancel = True: Exit Sub
        If RsInv.RecordCount = 0 Or (RsInv.EOF = True Or RsInv.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
        Else
            Txt(Index).TEXT = RsInv!Code
        End If
        FillRecords RsInv
    Case RegNo
        If RsConFin.RecordCount = 0 Or (RsConFin.EOF = True Or RsConFin.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsConFin!Name
            Txt(Index).Tag = RsConFin!Code
        End If
    Case VDate
        If Len(Trim(Txt(VDate).TEXT)) = 0 Then
            Txt(VDate).TEXT = PubLoginDate
        Else
            Txt(Index).TEXT = RetDate(Txt(Index))
        End If
        If CheckFinYear(Txt(Index)) Then
            Txt(TxtDocID) = GetDocID(GCnFaV, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
            DocID = Txt(TxtDocID)
            Set RsInv = GCn.Execute("SELECT " & cTrim(CStr("Veh_Order.Inv_No")) & " as code,Veh_Order.MODEL, Veh_Order.Chassis, Veh_Order.PartyCode, SubGroup.Name,Veh_Order.exp_DATE, Veh_Order.FB_CODE, Veh_Order.FIN_AMT, Veh_Order.Colour_Code, Veh_Order.Inv_DocId, Veh_Order.OrdDocId, Veh_Order.Ord_No, Veh_Order.Ord_Date, Veh_Order.Inv_SiteCode, Veh_Order.Inv_Date, Veh_Order.Inv_No ,Veh_Order.net_amount,Veh_Order.RTO " & _
                "FROM Veh_Order Left Join SubGroup on Veh_Order.PartyCode=SubGroup.SubCode " & _
                "where left(Veh_Order.inv_DocId,1)='" & PubDivCode & "' and Veh_Order.inv_DocId <> '' and Veh_Order.DelCh_DocId ='' and Inv_Date <= " & ConvertDate(Txt(VDate)) & " order by Inv_No")
            Set DGInv.DataSource = RsInv
        Else
            Cancel = True
        End If
    Case SerialNo
        If IsValid(Txt(SerialNo), "Serial No.") = False Then Cancel = True:   Exit Sub
            If VoucherEditFlag Then      ' Manual
                Txt(TxtDocID) = GetDocID(GCnFaV, mVType, Txt(VDate), VoucherEditFlag, Txt(SerialNo), LblVPrefix)
                DocID = Txt(TxtDocID)
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select DelCh_DocId From veh_Order Where DelCh_DocId='" & Txt(TxtDocID) & "'", GCn, adOpenStatic, adLockReadOnly
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                    Cancel = True
                    Txt(SerialNo).SetFocus
                End If
            End If
    Case IntPer, IntAmt, TDSPer, TDSAmt, RegFee, IncChg, SerChg, StampChg
         Txt(Index).TEXT = Format(Txt(Index).TEXT, "0.00")
         Amt_Cal
End Select
Set Rst = Nothing
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
Next I
Txt(BookNo).Tag = ""
FGrid.Rows = 1
FGrid.AddItem ""
FGrid.FixedRows = 1
FGrid1.Rows = 1
FGrid1.AddItem ""
FGrid1.FixedRows = 1
End Sub

Private Sub MoveRec()
Dim Rst As Recordset
Dim I As Integer
On Error GoTo error1
If Master.RecordCount > 0 Then
    DocID = Master!DelCh_DocId
    Txt(TxtDocID).TEXT = Master!DelCh_DocId
    LblDiv.CAPTION = "Division : " & left(Master!DelCh_DocId, 1)
    LblSite.CAPTION = "Site Code : " & mID(Master!DelCh_SiteCode, 1, 1)
    Txt(SiteCode).Tag = mID(Master!DelCh_SiteCode, 2, 1)
    Txt(SiteCode).TEXT = GCn.Execute("select site_desc from site where site_code = '" & Txt(SiteCode).Tag & "'").Fields(0).Value
    LblVPrefix.CAPTION = mID(Master!DelCh_DocId, 8, 5)
    Txt(SerialNo).TEXT = Master!DelCh_No
    Txt(VDate).TEXT = Master!DelCh_DT
    '*** A/c Posting Status
    Txt(AcPostByName) = IIf(IsNull(Master!DelCh_AcPostByUName), "", Master!DelCh_AcPostByUName)
    Txt(AcPostDate) = IIf(IsNull(Master!DelCh_AcPostByUEntDt), "", Master!DelCh_AcPostByUEntDt)
    '***
    LblUser = IIf(Not IsNull(Master!Delch_AddDate), "Add By : " & XNull(Master!Delch_AddBy) & "  Dated : " & XNull(Master!Delch_AddDate), "") & IIf(Not IsNull(Master!Delch_ModifyDate), "     Modify By : " & XNull(Master!Delch_ModifyBy) & "  Dated : " & XNull(Master!Delch_ModifyDate), "")
    FillRecords Master
'    FillCheckListGrid Txt(Model), Txt(ChassisNo)
    '***
    Txt(RebDays).TEXT = IIf(IsNull(Master!RebDays) Or Master!RebDays = 0, "", Master!RebDays)
    Txt(IntYN).TEXT = IIf(Master!Interest_YN = 1, "Yes", "No")
    Txt(TDSYN).TEXT = IIf(Master!TDS_YN = 1, "Yes", "No")
    
    Txt(IntPer).TEXT = IIf(IsNull(Master!InterestPer) Or Master!InterestPer = 0, "", Format(Master!InterestPer, "0.00"))
    Txt(TDSPer).TEXT = IIf(IsNull(Master!TDS_Per) Or Master!TDS_Per = 0, "", Format(Master!TDS_Per, "0.00"))
    Txt(IntAmt).TEXT = IIf(IsNull(Master!Interest) Or Master!Interest = 0, "", Format(Master!Interest, "0.00"))
    Txt(TDSAmt).TEXT = IIf(IsNull(Master!TDS_Amt) Or Master!TDS_Amt = 0, "", Format(Master!TDS_Amt, "0.00"))
    
    Txt(RegNo).TEXT = IIf(IsNull(Master!Reg_No), "", Master!Reg_No)
    Txt(CoverNote).TEXT = IIf(IsNull(Master!INS_NOTE), "", Master!INS_NOTE)
    
    Txt(RegFee).TEXT = IIf(IsNull(Master!REG_FEE) Or Master!REG_FEE = 0, "", Format(Master!REG_FEE, "0.00"))
    Txt(IncChg).TEXT = IIf(IsNull(Master!INS_FEE) Or Master!INS_FEE = 0, "", Format(Master!INS_FEE, "0.00"))
    Txt(SerChg).TEXT = IIf(IsNull(Master!S_CHARGE) Or Master!S_CHARGE = 0, "", Format(Master!S_CHARGE, "0.00"))
    Txt(StampChg).TEXT = IIf(IsNull(Master!STAMP_DUTY) Or Master!STAMP_DUTY = 0, "", Format(Master!STAMP_DUTY, "0.00"))
    Txt(RegBy) = XNull(Master!FinName)
    Txt(RegBy).Tag = XNull(Master!RegBy)
Else
    Call BlankText
End If
Grid_Hide
Set Rst = Nothing
Amt_Cal
If UCase(left(PubComp_Name, 7)) = "JOHNSON" Then
    CLEARTEXT
End If
Exit Sub
error1:
    CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
    Txt(I).ForeColor = CtrlFColOrg
Next

If UCase(left(PubComp_Name, 4)) = "ENAR" Then Txt(SiteCode).Enabled = False

If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
    LblPaybleAmt = "Receivable Amt"
End If

If TopCtrl1.TopText2 = "Edit" Then
    Txt(SiteCode).Enabled = False
    Txt(VDate).Enabled = False
    Txt(SerialNo).Enabled = False
    Txt(BookNo).Enabled = False
    Txt(Model).Enabled = False
    Txt(InvNo).Enabled = False
End If

Txt(TxtDocID).Enabled = False
Txt(ExpDate).Enabled = False
Txt(ChassisNo).Enabled = False
Txt(NamePrefix).Enabled = False
Txt(FNamePrefix).Enabled = False
Txt(fname).Enabled = False
Txt(Party).Enabled = False
Txt(Add1).Enabled = False
Txt(Add2).Enabled = False
Txt(Add3).Enabled = False
Txt(City).Enabled = False
Txt(EngineNo).Enabled = False
Txt(Colours).Enabled = False
Txt(Model).Enabled = False
Txt(FinAmt).Enabled = False
Txt(FB_Code).Enabled = False
Txt(InvDate).Enabled = False
Txt(BookNo).Enabled = False
Txt(BookDate).Enabled = False
Txt(PayAmt).Enabled = False
Txt(AdvAmt).Enabled = False
Txt(TDSAmt).Enabled = False
Txt(NetPayAmt).Enabled = False
Txt(VehAmt).Enabled = False
Txt(RTO).Enabled = False


If GCn.Execute("Select " & vIsNull("RtoInsInBill", "0") & " From Syctrl").Fields(0) = 0 Then
    Txt(RegFee).Enabled = Enb
    Txt(IncChg).Enabled = Enb
Else
    Txt(RegFee).Enabled = False
    Txt(IncChg).Enabled = False
End If

   'BookNo, BookDate,InvNo,InvDate,Party,Add1,Add2,Add3,City,Model,ChassisNo,EngineNo,Colours,RTO 17
'FB_Code,PayAmt,FinAmt,VehAmt,AdvAmt
'IntYN,RebDays,IntPer,IntAmt,TDSYN,TDSper,TDSAmt , RegFee, RegNo, Incchg, CoverNote, SerChg, StampChg

txtDisabled_Color Me

End Sub
Private Sub Grid_Hide()
    If DGVno.Visible = True Then DGVno.Visible = False
    If DGInv.Visible = True Then DGInv.Visible = False
    If DGSite.Visible = True Then DGSite.Visible = False
    If DGConFin.Visible = True Then DGConFin.Visible = False
End Sub
Private Sub Amt_Cal()
 
Txt(PayAmt) = Format((Val(Txt(VehAmt)) - Val(Txt(FinAmt)) - Val(Txt(AdvAmt))), "0.00")
Txt(NetPayAmt) = Format((Val(Txt(PayAmt)) - Val(Txt(IntAmt)) + Val(Txt(TDSAmt)) + Val(Txt(TDSAmt)) + Val(Txt(RegFee)) + Val(Txt(IncChg)) + Val(Txt(SerChg)) + Val(Txt(StampChg))), "0.00")
If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
    If Txt(PayAmt) < 0 Then Txt(PayAmt) = 0
End If
End Sub
Private Sub FillRecords(RsFill As ADODB.Recordset)
Dim Rst As ADODB.Recordset
    If RsFill.RecordCount > 0 Then
        Txt(BookNo).TEXT = RsFill!Ord_No
        Txt(BookNo).Tag = RsFill!OrdDocId
        Txt(BookDate).TEXT = RsFill!Ord_Date
        Txt(ExpDate).TEXT = IIf(IsNull(RsFill!EXP_DATE), "", RsFill!EXP_DATE)
        Txt(InvNo).TEXT = RsFill!Inv_No
        Txt(InvDate).TEXT = IIf(IsNull(RsFill!Inv_Date), "", RsFill!Inv_Date)
        InvdocId = RsFill!Inv_DocId
        Txt(RTO) = RsFill!RTO
        Txt(Party).Tag = IIf(IsNull(RsFill!PartyCode), "", RsFill!PartyCode)
        If Txt(Party).Tag <> "" Then
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select NamePrefix,name,FPrefix,FName,add1,add2,add3,CityCode from SubGroup where Subcode = '" & Txt(Party).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
            Txt(NamePrefix).TEXT = IIf(IsNull(Rst!NamePrefix), "", Rst!NamePrefix)
            Txt(Party).TEXT = Rst!Name
            Txt(FNamePrefix).TEXT = IIf(IsNull(Rst!FPrefix), "", Rst!FPrefix)
            Txt(fname).TEXT = IIf(IsNull(Rst!fname), "", Rst!fname)
            Txt(Add1).TEXT = IIf(IsNull(Rst!Add1), "", Rst!Add1)
            Txt(Add2).TEXT = IIf(IsNull(Rst!Add2), "", Rst!Add2)
            Txt(Add3).TEXT = IIf(IsNull(Rst!Add3), "", Rst!Add3)
            Txt(City).Tag = IIf(IsNull(Rst!CityCode), "", Rst!CityCode)
            If Txt(City).Tag <> "" Then
                Txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").Fields(0).Value
            End If
        End If
        Txt(Model).TEXT = RsFill!Model
        Txt(Colours).Tag = IIf(IsNull(RsFill!Colour_Code), "", RsFill!Colour_Code)
        If Txt(Colours).Tag <> "" Then
            Txt(Colours).TEXT = GCn.Execute("select col_desc from colmast where col_code = '" & Txt(Colours).Tag & "'").Fields(0).Value
        End If
        Txt(FB_Code).Tag = IIf(IsNull(RsFill!FB_Code), "", RsFill!FB_Code)
        If Txt(FB_Code).Tag <> "" Then
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select fincode as code,finname + ',' + " & xIsNull("City.CityName", "") & " as name,AcCode,FinBankCode from ContractFinance " & _
                "left join city on left(ContractFinance.City,4)=City.CityCode where fincatg = 0  and  fincode = '" & Txt(FB_Code).Tag & "'", GCn, adOpenStatic, adLockReadOnly
            If Rst.RecordCount > 0 Then
                Txt(FB_Code).TEXT = Rst!Name
            Else
                Txt(FB_Code).TEXT = ""
            End If
        Else
            Txt(FB_Code).TEXT = ""
        End If
        Txt(ChassisNo).TEXT = IIf(IsNull(RsFill!Chassis), "", RsFill!Chassis)
        Set Rst = New Recordset
        Rst.Open "SELECT Veh_Stock.EngineNo,Veh_Stock.VehSerialNo,Veh_Stock.tax_yn,Veh_Stock.PBILL_NO,Veh_Stock.PBILL_DATE FROM Veh_Stock where Veh_Stock.MODEL  = '" & Txt(Model) & "' and Veh_Stock.ChassisNo = '" & Txt(ChassisNo) & "' and Veh_Stock.Sal_DocId= '" & RsFill!Inv_DocId & "'", GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount > 0 Then
            Txt(EngineNo).TEXT = IIf(IsNull(Rst!EngineNo), "", Rst!EngineNo)
        End If
        Txt(FinAmt).TEXT = Format(IIf(IsNull(RsFill!Fin_Amt), 0, RsFill!Fin_Amt), "0.00")
        Txt(VehAmt).TEXT = Format(IIf(IsNull(RsFill!Net_Amount), 0, RsFill!Net_Amount), "0.00")
        'modified for docid / invdate by lps
        Txt(AdvAmt) = Format(PartyAdvance(RsFill!OrdDocId, Txt(InvDate)), "0.00")
        '*****end modi
        If Txt(IntYN).TEXT = "No" Then EnableText False
        'Display Customer Trns
        FillCustTrns Txt(BookNo).Tag, True, False, Val(Txt(IntPer)), Val(Txt(RebDays)), Txt(InvDate)
        'Display Check List Items
        FillCheckListGrid Txt(Model), Txt(ChassisNo)
    Else
        Txt(BookNo).TEXT = ""
        Txt(ExpDate).TEXT = ""
        Txt(BookDate).TEXT = ""
        Txt(InvNo).TEXT = ""
        Txt(InvDate).TEXT = ""
        InvdocId = ""
        Txt(Party).Tag = ""
        Txt(Party).TEXT = ""
        Txt(Add1).TEXT = ""
        Txt(Add2).TEXT = ""
        Txt(Add3).TEXT = ""
        Txt(City).Tag = ""
        Txt(City).TEXT = ""
        Txt(Model).TEXT = ""
        Txt(Colours).Tag = ""
        Txt(Colours).TEXT = ""
        Txt(FB_Code).Tag = ""
        Txt(FB_Code).TEXT = ""
        Txt(ChassisNo).TEXT = ""
        Txt(EngineNo).TEXT = ""
        Txt(FinAmt).TEXT = ""
        Txt(VehAmt).TEXT = ""
        Txt(AdvAmt).TEXT = ""
    End If
    Set Rst = Nothing
    Amt_Cal
End Sub
Private Sub EnableText(Enb As Boolean)
Txt(IntAmt).Enabled = Enb
Txt(IntPer).Enabled = Enb
Txt(RebDays).Enabled = Enb
Txt(TDSYN).Enabled = Enb
Txt(TDSPer).Enabled = Enb
Txt(TDSAmt).Enabled = Enb
Txt(IntAmt).BackColor = CtrlBColOrg
Txt(IntPer).BackColor = CtrlBColOrg
Txt(RebDays).BackColor = CtrlBColOrg
Txt(TDSYN).BackColor = CtrlBColOrg
Txt(TDSPer).BackColor = CtrlBColOrg
Txt(TDSAmt).BackColor = CtrlBColOrg

If Enb = False Then
    Txt(IntAmt).TEXT = ""
    Txt(IntPer).TEXT = ""
    Txt(RebDays).TEXT = ""
    Txt(TDSYN).TEXT = ""
    Txt(TDSPer).TEXT = ""
    Txt(TDSAmt).TEXT = ""
    Txt(IntAmt).BackColor = CtrlBColDisabled
    Txt(IntPer).BackColor = CtrlBColDisabled
    Txt(RebDays).BackColor = CtrlBColDisabled
    Txt(TDSYN).BackColor = CtrlBColDisabled
    Txt(TDSPer).BackColor = CtrlBColDisabled
    Txt(TDSAmt).BackColor = CtrlBColDisabled
End If
End Sub
'************************ PRINTING CODE ******************


Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case FromVno, ToVno
            RsVno.Close
            RsVno.Open "Select DelCh_No as code from Veh_Order where right(veh_order.DelCh_SiteCode,1)='" & txtPrint(SiteCode1).Tag & "' and  veh_order.DelCh_VType='V_DCL'", GCn, adOpenDynamic, adLockOptimistic
            Set DGVno.DataSource = RsVno
            If txtPrint(Index).TEXT <> RsVno!Code Then
                RsVno.MoveFirst
                RsVno.FIND "code ='" & txtPrint(Index).TEXT & "'"
            End If
            If Index = ToVno Then DGVno.Tag = "1" Else DGVno.Tag = "2"
       
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
    Case SiteCode1
        DGridTxtKeyDown DGSite, txtPrint, Index, RsSite, KeyCode, False, 1
End Select
If DGSite.Visible = False And DGVno.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> SiteCode1 Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TxtPrint_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case FromVno, ToVno
        If DGVno.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsVno, KeyAscii, "Code"
    Case SiteCode1
        If DGSite.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsSite, KeyAscii, "Name"
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
Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "VehDel", "VehDel")
        Call WindowsPrint(Index)
        FrmPrn.Visible = False
    Case PDos
        Call SpeedPrint
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "VehDel", "VehDel")
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
Dim Rst As ADODB.Recordset, mQry As String
Dim RstSub1 As ADODB.Recordset
Dim I As Integer
Dim Rst2 As ADODB.Recordset
Dim mCurrBal, mCrLimit As Double
On Error GoTo ERRORHANDLER

'    mCurrBal = GCn.Execute("Select Curr_Bal from SubGroup where SubCode='" & txt(Party).Tag & "'").Fields(0).Value
'    mCrLimit = GCn.Execute("Select CreditLimit from SubGroup where SubCode='" & txt(Party).Tag & "'").Fields(0).Value
'    If mCurrBal  > 0 Then     'Dr Balance
'        If mCrLimit  > 0 Then
'            If mCurrBal  > mCrLimit Then
'                If MsgBox("Cr Limit Rs." & mCrLimit & " Exceeds by Rs." & mCurrBal - mCrLimit & vbCrLf & "Want To Print ?", vbYesNo, "Cr Limit Checking") = vbNo Then
'                    Me.ActiveControl.SetFocus: Exit Sub
'                End If
'            End If
'        Else
'            If MsgBox("Balance Rs. " & mCurrBal & vbCrLf & "Want To Print ?", vbYesNo, "Balance Checking") = vbNo Then
'                Me.ActiveControl.SetFocus: Exit Sub
'            End If
'        End If
'    End If
        If UCase(left(PubComp_Name, 7)) = "JOHNSON" Then
        
            mQry = "SELECT '" & Txt(InvNo).TEXT & "' as inv_No,'" & Txt(InvDate) & "' as inv_Date,Veh_Order.Inv_DocId," & _
                " SubGroup.FPrefix,subgroup.FName,veh_order.DelChPrn_YN,veh_order.FIN_AMT,Veh_Purch1.Tot_Amount, " & _
                " City_1.CityName AS fincity, ContractFinance.Add1 AS finadd1, ContractFinance.Add2 AS finadd2, FinBank.FinBankName, '" & Txt(FB_Code).TEXT & "' as FinName,  City.CityName, Veh_Order.DelCh_UName, Veh_Order.DelCh_UEntDt, Veh_Order.DelCh_No, Veh_Order.DelCh_DT, Veh_Order.Fund_Source,  Model.TYRES, '" & Txt(Model).TEXT & "' as MODEL, Veh_Order.DelCh_SiteCode, Model.RIMS,  '" & Txt(ChassisNo).TEXT & "' as ChassisNo,'" & Txt(EngineNo).TEXT & "' as EngineNo ," & _
                " '" & CabNo & "' AS CabNo,'" & Transaxleno & "' AS Transaxleno,'" & Rearaxleno & "' AS Rearaxleno ,'" & fipno & "' AS fipno,'" & alternatorno & "' AS alternatorno,'" & startingmono & "' AS startingmono,'" & accompno & "' AS accompno, " & _
                " '" & tyremake & "' AS tyremake ,'" & tyremakeA & "' AS tyremakeA ,'" & tyremakeB & "' AS tyremakeB  ,'" & tyremakeC & "' AS tyremakeC ,'" & tyremakeD & "' AS tyremakeD ,'" & tyremakeE & "' AS tyremakeE ,'" & toolsandequip & "' AS toolsandequip, " & _
                " '" & connecteddocu & "' AS connecteddocu ,'" & kmsreading & "' AS kmsreading ,'" & pdion & "' AS pdion  ,'" & batteryno & "' AS batteryno ,'" & deliverytakenby & "' AS deliverytakenby, " & _
                " Veh_Order.Ord_No, Veh_Order.Ord_Date, Model.Model_Desc, Model.Model_Desc1, ColMast.Col_Desc,'" & Txt(Party) & "' as Name,'" & Txt(Add1) & "' as add1,'" & Txt(Add2) & "' as add2,'" & Txt(Add3) & "' as add3, SubGroup.PIN, ContractFinance.PinCode,Veh_Order.DelCh_DocId,Veh_Order.OrdDocId,'" & steeringboxno & "' AS steeringboxno " & _
                " FROM (((((((((Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) LEFT JOIN TaxForms ON Veh_Order.Form_Code = TaxForms.Form_Code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code) " & _
                " LEFT JOIN Model ON Veh_Order.MODEL = Model.MODEL) LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) " & _
                " LEFT JOIN City ON SubGroup.CityCode = City.CityCode) LEFT JOIN ContractFinance ON Veh_Order.FB_CODE = ContractFinance.FinCode) " & _
                " LEFT JOIN FinBank ON ContractFinance.FinBankCode = FinBank.FinBankCode) LEFT JOIN City AS City_1 ON ContractFinance.City = City_1.CityCode) LEFT JOIN Veh_Purch1 ON Veh_Stock.Pur_DocId = Veh_Purch1.DocID " & _
                " where Veh_Order.DelCh_DocId = '" & Master!SearchCode & "'"
            
            mRepName = "VehDelJohnSon"
    Else
            mQry = "SELECT Veh_Order.inv_No,Veh_Order.inv_Date,Veh_Order.Inv_DocId," & _
                " SubGroup.FPrefix,subgroup.FName,veh_order.DelChPrn_YN,veh_order.FIN_AMT,Veh_Purch1.Tot_Amount, " & _
                " City_1.CityName AS fincity, ContractFinance.Add1 AS finadd1, ContractFinance.Add2 AS finadd2, FinBank.FinBankName, ContractFinance.FinName,  City.CityName, Veh_Order.DelCh_UName, Veh_Order.DelCh_UEntDt, Veh_Order.DelCh_No, Veh_Order.DelCh_DT, Veh_Order.Fund_Source,  Model.TYRES, Veh_Order.MODEL, Veh_Order.DelCh_SiteCode, Model.RIMS,  Veh_Stock.ChassisNo, Veh_Stock.EngineNo," & _
                " Veh_Order.Ord_No, Veh_Order.Ord_Date, Model.Model_Desc, Model.Model_Desc1, ColMast.Col_Desc, SubGroup.Name, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, SubGroup.PIN, ContractFinance.PinCode,Veh_Order.DelCh_DocId,Veh_Order.OrdDocId " & _
                " FROM (((((((((Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) LEFT JOIN TaxForms ON Veh_Order.Form_Code = TaxForms.Form_Code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code) " & _
                " LEFT JOIN Model ON Veh_Order.MODEL = Model.MODEL) LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) " & _
                " LEFT JOIN City ON SubGroup.CityCode = City.CityCode) LEFT JOIN ContractFinance ON Veh_Order.FB_CODE = ContractFinance.FinCode) " & _
                " LEFT JOIN FinBank ON ContractFinance.FinBankCode = FinBank.FinBankCode) LEFT JOIN City AS City_1 ON ContractFinance.City = City_1.CityCode) LEFT JOIN Veh_Purch1 ON Veh_Stock.Pur_DocId = Veh_Purch1.DocID " & _
                " where Veh_Order.DelCh_DocId = '" & Master!SearchCode & "'"
    End If



    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
        
    mQry = "SELECT VO.Inv_DocId,VP2.Trn_Type,VAM.Prod_Name,VP2.QTY " & _
        "FROM (Veh_Order as VO LEFT JOIN Veh_Stock as VS ON VO.Inv_DocId = VS.Sal_DocId) " & _
        "LEFT JOIN (Veh_Purch2 as VP2 LEFT JOIN Veh_AMDModel as VAM ON VP2.PROD_CODE = VAM.Prod_Code) ON VS.Pur_DocId = VP2.DocID " & _
        "where VO.DelCh_DocId = '" & Master!SearchCode & "'"
        
    Set RstSub1 = New Recordset
    RstSub1.CursorLocation = adUseClient
    RstSub1.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
        
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
   ' CreateFieldDefFile RstSub1, PubRepoPath + "\" & mRepName & "1.ttx", True

    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    If UCase(left(PubComp_Name, 7)) = "JOHNSON" Then
            rpt.Database.SetDataSource Rst
            Set Rst2 = New ADODB.Recordset
            Rst2.CursorLocation = adUseClient
            Rst2.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
            For I = 1 To rpt.FormulaFields.Count
                Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                    Case UCase("SubTitle")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecSpeciality & "'"
                    Case UCase("LST")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecLST & "'"
                    Case UCase("LSTDate")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecLST_Date & "'"
                    Case UCase("CST")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecCST & "'"
                    Case UCase("CSTDate")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecCST_Date & "'"
                    Case UCase("Phone")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecPhone & "'"
                    Case UCase("Fax")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecFax & "'"
                    Case UCase("Gram")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecGram & "'"
                End Select
            Next
    Else
           rpt.Database.SetDataSource Rst
           rpt.OpenSubreport("SUBREP!").Database.SetDataSource RstSub1

            Set Rst2 = New ADODB.Recordset
            Rst2.CursorLocation = adUseClient
            Rst2.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
   
            
            For I = 1 To rpt.FormulaFields.Count
                Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                    Case UCase("PrintShortage")
                        If RstSub1.RecordCount > 0 And Not IsNull(RstSub1!Prod_Name) Then
                            rpt.FormulaFields(I).TEXT = 1
                        Else
                            rpt.FormulaFields(I).TEXT = 0
                        End If
                    Case UCase("SubTitle")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecSpeciality & "'"
                    Case UCase("LST")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecLST & "'"
                    Case UCase("LSTDate")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecLST_Date & "'"
                    Case UCase("CST")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecCST & "'"
                    Case UCase("CSTDate")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecCST_Date & "'"
                    Case UCase("Phone")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecPhone & "'"
                    Case UCase("Fax")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecFax & "'"
                    Case UCase("Gram")
                        rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecGram & "'"
                    Case UCase("SubRep")
                        rpt.FormulaFields(I).TEXT = "" & IIf(RstSub1.RecordCount = 0, 0, 1) & ""
                End Select
            Next
        End If
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
                    rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
            End Select
            Next
            rpt.PrintOut False
            CmdPrint(0).Tag = ""
            If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
                GCn.Execute "update Veh_order set DelChPrn_YN = 1  where Veh_Order.DelCh_DocId = '" & Master!SearchCode & "'"

            End If

    Case PScreen  'screen
            Call Report_View(rpt, Me.CAPTION, , True)
End Select
CmdPrint(PSetUp).Tag = ""
Set Rst = Nothing
Set Rst2 = Nothing
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
On Error Resume Next
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, j As Integer, mQry As String
    Dim PrintStr As String
    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstDel As ADODB.Recordset
    Dim Page As Byte, mLine As Byte, mFix As Byte
    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$
    Dim FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double, RstInvDet As ADODB.Recordset
    Dim fob As New FileSystemObject
    Dim mJuriCity As String
    Dim Cnt As Byte, mAmt As Double, PrnStr As String, PrnStr1 As String
    Dim Left1 As String, Left2 As String, Left3 As String
    Dim Left4 As String, Left5 As String, Left6 As String, Left7 As String
    Dim Right1 As String, Right2 As String, Right3 As String
    Dim Right4 As String, Right5 As String, Right6 As String, Right7 As String
    Dim NetAmt As Double
    Dim mCurrBal, mCrLimit As Double, mInv_No$
       
    mCurrBal = GCn.Execute("Select Curr_Bal from SubGroup where SubCode='" & Txt(Party).Tag & "'").Fields(0).Value
    mCrLimit = GCn.Execute("Select CreditLimit from SubGroup where SubCode='" & Txt(Party).Tag & "'").Fields(0).Value
    
    If mCurrBal > 0 Then      'Dr Balance
        If mCrLimit > 0 Then
            If mCurrBal > mCrLimit Then
                If MsgBox("Cr Limit Rs." & mCrLimit & " Exceeds by Rs." & mCurrBal - mCrLimit & vbCrLf & "Want To Print ?", vbYesNo, "Cr Limit Checking") = vbNo Then
                    Me.ActiveControl.SetFocus: Exit Sub
                End If
            End If
        Else
            If MsgBox("Balance Rs. " & mCurrBal & vbCrLf & "Want To Print ?", vbYesNo, "Balance Checking") = vbNo Then
                Me.ActiveControl.SetFocus: Exit Sub
            End If
        End If
    End If

    Set RstDel = GCn.Execute("SELECT Veh_Order.inv_No,Veh_Order.inv_Date,Veh_Order.Inv_DocId," & _
        " subgroup.FPrefix,subgroup.FName,veh_order.DelChPrn_YN,veh_order.FIN_AMT,Veh_Purch1.Tot_Amount, " & _
        "City_1.CityName AS fincity, ContractFinance.Add1 AS finadd1, ContractFinance.Add2 AS finadd2, FinBank.FinBankName, ContractFinance.FinName,  City.CityName, Veh_Order.DelCh_UName, Veh_Order.DelCh_UEntDt, Veh_Order.DelCh_No, Veh_Order.DelCh_DT, Veh_Order.Fund_Source,  Model.TYRES, Veh_Order.MODEL, Veh_Order.DelCh_SiteCode, Model.RIMS,  Veh_Stock.ChassisNo, Veh_Stock.EngineNo," & _
        " Veh_Order.Ord_No, Veh_Order.Ord_Date, Model.Model_desc, Model.Model_Desc1, ColMast.Col_Desc, SubGroup.Name, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, SubGroup.PIN, ContractFinance.PinCode,Veh_Order.DelCh_DocId,Veh_Order.OrdDocId,Model.Model as Modl,Model.Sales_Desc " & _
        " FROM (((((((((Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) LEFT JOIN TaxForms ON Veh_Order.Form_Code = TaxForms.Form_Code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code) " & _
        " LEFT JOIN Model ON Veh_Order.MODEL = Model.MODEL) LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) " & _
        "LEFT JOIN City ON SubGroup.CityCode = City.CityCode) LEFT JOIN ContractFinance ON Veh_Order.FB_CODE = ContractFinance.FinCode) LEFT JOIN FinBank ON ContractFinance.FinBankCode = FinBank.FinBankCode) LEFT JOIN City AS City_1 ON ContractFinance.City = City_1.CityCode) LEFT JOIN Veh_Purch1 ON Veh_Stock.Pur_DocId = Veh_Purch1.DocID " & _
        " where Veh_Order.DelCh_DocId = '" & Master!SearchCode & "'")
      
    If RstDel.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
 
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 15
    
    ' Header
          
    mDocStr = IIf(RstDel!DelChPrn_YN = 0, "Vehicle Delivery Order", "Vehicle Delivery Order (Duplicate)")
    mDupStr = ""

      Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

         Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
         mHeader = mHeader + 1
         If XNull(RstCompDet!V_SecSpeciality) <> "" Then
             Print #1, PRN_TIT(RstCompDet!V_SecSpeciality, "C", PageWidth)
             mHeader = mHeader + 1
         End If
         Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
         mHeader = mHeader + 1
         
         If PubComp_Add2 <> "" Or PubComp_City <> "" Then
             Print #1, PRN_TIT(PubComp_Add2 & IIf(PubComp_Add2 = "" Or PubComp_City = "", "", ",") & PubComp_City, "C", PageWidth)
             mHeader = mHeader + 1
         End If
         Print #1, PRN_TIT(IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", " Fax   : ") & XNull(RstCompDet!V_SecFax), "C", PageWidth)
         mHeader = mHeader + 1
         If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
                Print #1, PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & XNull(RstCompDet!V_SecCST_Date)), 40) & PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & XNull(RstCompDet!V_SecLST_Date)), 40, , AlignRight)
         Else
                Print #1, PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & XNull(RstCompDet!V_SecLST_Date)), 40) & PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & XNull(RstCompDet!V_SecCST_Date)), 40, , AlignRight)
         End If
         mHeader = mHeader + 1

         Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth) & mChr18 & mEmph
         mHeader = mHeader + 1
        
 '0 -Hypothecation ,1- Hire purchase ,2 -Own Fund,3- Lease

    If RstDel!Fund_Source = 0 Then   'Hypothecation
        Left1 = "To,"
        Left2 = XNull(RstDel!Name)
        Left3 = XNull(RstDel!FPrefix) & " " & XNull(RstDel!fname)
        Left4 = XNull(RstDel!Add1)
        Left5 = XNull(RstDel!Add2)
        Left6 = XNull(RstDel!Add3) & IIf(XNull(RstDel!CityName) = "" Or XNull(RstDel!Add3) = "", "", ",") & XNull(RstDel!CityName)
        
        Right1 = "Under Hypothecation to  "
        Right2 = XNull(RstDel!finbankname)
        If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
            Right3 = XNull(RstDel!FinAdd1)
            Right4 = XNull(RstDel!FinAdd2)
            Right5 = XNull(RstDel!FinCity)
        Else
            Right3 = ""
            Right4 = ""
            Right5 = ""
        End If
        Right6 = "Finance Amount :" & Format(RstDel!Fin_Amt, "0.00")
        
    ElseIf RstDel!Fund_Source = 1 Then  'Hire Purchase
        Left1 = "Sold to under HPA with, "
        Left2 = left("U/F " & XNull(RstDel!finbankname) & Space(40), 40)
        If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
            Left3 = XNull(RstDel!FinAdd1)
            Left4 = XNull(RstDel!FinAdd2)
            Left5 = XNull(RstDel!FinCity)
        Else
            Left3 = ""
            Left4 = ""
            Left5 = ""
        End If
        Left6 = ""
           
        Right1 = "Delivered to Hirer, "
        Right2 = XNull(RstDel!Name)
        Right3 = XNull(RstDel!FPrefix) & " " & XNull(RstDel!fname)
        Right4 = XNull(RstDel!Add1)
        Right5 = XNull(RstDel!Add2)
        Right6 = XNull(RstDel!Add3) & IIf(XNull(RstDel!CityName) = "" Or XNull(RstDel!Add3) = "", "", ",") & XNull(RstDel!CityName)
    
    ElseIf RstDel!Fund_Source = 3 Then 'Lease
        Left1 = "To, "
        Left2 = XNull(RstDel!Name)
        Left3 = XNull(RstDel!FPrefix) & " " & XNull(RstDel!fname)
        Left4 = XNull(RstDel!Add1)
        Left5 = XNull(RstDel!Add2)
        Left6 = XNull(RstDel!Add3) & IIf(XNull(RstDel!CityName) = "" Or XNull(RstDel!Add3) = "", "", ",") & XNull(RstDel!CityName)
        
        Right1 = "Leaser  "
        Right2 = XNull(RstDel!finbankname)
        If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
            Right3 = XNull(RstDel!FinAdd1)
            Right4 = XNull(RstDel!FinAdd2)
            Right5 = XNull(RstDel!FinCity)
        Else
            Right3 = ""
            Right4 = ""
            Right5 = ""
        End If
        Right6 = "Lease Amount :" & RstDel!Fin_Amt
    Else
        Left1 = "Sold To,"
        Left2 = XNull(RstDel!Name)
        Left3 = XNull(RstDel!FPrefix) & " " & XNull(RstDel!fname)
        Left4 = XNull(RstDel!Add1)
        Left5 = XNull(RstDel!Add2)
        Left6 = XNull(RstDel!Add3) & IIf(XNull(RstDel!CityName) = "" Or XNull(RstDel!Add3) = "", "", ",") & XNull(RstDel!CityName)
    End If

        Print #1, mChr18 & mEmph & PSTR(Left1, 40) & PSTR(Right1, 40) & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR(Left2, 40) & PSTR(Right2, 40)
        mHeader = mHeader + 1
        Print #1, PSTR(Left3, 40) & PSTR(Right3, 40)
        mHeader = mHeader + 1
        Print #1, PSTR(Left4, 40) & PSTR(Right4, 40)
        mHeader = mHeader + 1
        Print #1, PSTR(Left5, 40) & PSTR(Right5, 40)
        mHeader = mHeader + 1
        Print #1, PSTR(Left6, 40) & PSTR(Right6, 40)
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        
        Set RstInvDet = GCn.Execute("select SupInvOnVehSaleInv , TaxDetOnVehInv, VehSaleInv_Prefix from syctrl")
        
        Print #1, mEmph & PSTR("Booking No.  : " & PrinID(RstDel!OrdDocId), 40) & "Delivery Order No. : " & " " & PrinID(RstDel!DelCh_DocId) & mEmph1
        mHeader = mHeader + 1
        Print #1, PSTR("Booking Date : " & STR(RstDel!Ord_Date), 40) & "Delivery Order Date: " & XNull(RstDel!DelCh_DT)
        mHeader = mHeader + 1
        
        mInv_No = Trim(DeCodeDocID(RstDel!Inv_DocId, Document_Prefix)) & " - " & Trim(DeCodeDocID(RstDel!Inv_DocId, Document_No))
        Print #1, mEmph & "Invoice No. & Date: " & PSTR(mInv_No, 17, , AlignLeft) & " " & XNull(RstDel!Inv_Date) & mEmph1
        mHeader = mHeader + 1
        
        Print #1, Replace(Space(PageWidth), " ", "-")
        mHeader = mHeader + 1
               
        Print #1, mEmph & PSTR("Model", 15) & " : " & RstDel!Modl & mEmph1
        mHeader = mHeader + 1
        If RstDel!Model_Desc <> "" Then
            Print #1, Space(15) & "   " & RstDel!Model_Desc
            mHeader = mHeader + 1
        End If
        Print #1, Space(15) & "   <" & RstDel!Tyres & " > Tyres And " & RstDel!Rims & " Rims  >"
        mHeader = mHeader + 1
        Print #1, PSTR("Colour", 15) & " : " & RstDel!Col_Desc
        mHeader = mHeader + 1
        Print #1, mEmph & PSTR("Chassis No.", 15) & " : " & RstDel!ChassisNo & mEmph1
        mHeader = mHeader + 1
        Print #1, mEmph & PSTR("Engine No.", 15) & " : " & RstDel!EngineNo & mEmph1
        mHeader = mHeader + 1
        Print #1, "Battery Perticulars : Fitted with 12 volt Battery : Make          No."
        mHeader = mHeader + 1
        Print #1, ""
        If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
            mHeader = mHeader + 1
            Print #1, PRN_TIT("List of Documents Supplied with the Chassis", "C", PageWidth)
            mHeader = mHeader + 1
            Print #1, Replace(Space(PageWidth), " ", "-")
            mHeader = mHeader + 1
            Print #1, PSTR("Description", 30) & PSTR("Qty", 10, , AlignRight) & PSTR("Description", 30) & PSTR("Qty", 10, , AlignRight)
            mHeader = mHeader + 1
            Print #1, Replace(Space(PageWidth), " ", "-")
            mHeader = mHeader + 1
            Print #1, PSTR("1.Vehicle Defect Report Form", 30) & Space(10) & PSTR("5.Ignition Key", 30)
            mHeader = mHeader + 1
            Print #1, PSTR("2.Operator's Service Book", 30) & Space(10) & PSTR("6.Wiper Motor Assy Set", 30)
            mHeader = mHeader + 1
            Print #1, PSTR("3.Battery Warranty Card", 30) & Space(10) & PSTR("7.Tool Kit", 30)
            mHeader = mHeader + 1
            Print #1, PSTR("4.Key Ring", 30) & Space(10) & PSTR("8.Jack & Tomy", 30)
            mHeader = mHeader + 1
        Else
            Print #1, PSTR("With The Fallowing Tools", 80)
            mHeader = mHeader + 1
            Print #1, Replace(Space(PageWidth), " ", "-")
            mHeader = mHeader + 1
            Print #1, PSTR("Owner | ToolKit | Stepny | Reflector | FirstAid | Coupans | Bulb | Jack | Remark", 80)
            mHeader = mHeader + 1
            Print #1, PSTR("Manual|         |        |           | Kit      |         |      |&Tomy |       ", 80)
            mHeader = mHeader + 1
            Print #1, ""
            Print #1, Replace(Space(PageWidth), " ", "-")
            mHeader = mHeader + 1
        End If
'        Set Rst = GCn.Execute("SELECT Veh_Order.Inv_DocId,Veh_Purch2.Trn_Type,Veh_Purch2.QTY, Veh_AMDModel.Prod_Name " & _
        "FROM (Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) " & _
        "LEFT JOIN (Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code) ON Veh_Stock.Pur_DocId = Veh_Purch2.DocID " & _
        "where Veh_Order.DelCh_DocId = '" & Master!SearchCode & "'")
        'If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
        '    Set Rst = GCn.Execute("SELECT VO.Inv_DocId,VP2.Trn_Type,VAM.Prod_Name,VP2.QTY " & _
        '        "FROM (Veh_Order as VO LEFT JOIN Veh_Stock as VS ON VO.Inv_DocId = VS.Sal_DocId) " & _
        '        "LEFT JOIN (Veh_Purch2 as VP2 LEFT JOIN Veh_AMDModel as VAM ON VP2.PROD_CODE = VAM.Prod_Code) ON VS.Pur_DocId = VP2.DocID " & _
        '        "where VO.DelCh_DocId = '" & Master!SearchCode & "'")
        '    If Rst.RecordCount > 0 And Not IsNull(Rst!Prod_Name) Then
        '        Print #1, mEmph & "Shortage :  " & mEmph1
        '        mHeader = mHeader + 1
        '        Print #1, mDoub & PSTR("ItemName", 52) & PSTR("Qty", 13, , AlignRight) & mDoub1
        '        mHeader = mHeader + 1
        '        Do Until Rst.EOF
        '            Print #1, PSTR(Rst!Prod_Name, 52) & PSTR(Rst!Qty, 13, 2)
        '            mHeader = mHeader + 1
        '            Rst.MoveNext
        '        Loop
        '        Print #1, Replace(Space(PageWidth), " ", "-")
        '        mHeader = mHeader + 1
        '    End If
        'End If
        Do Until mHeader >= PageLength - (mFooter + 5)
            Print #1, ""
            mHeader = mHeader + 1
        Loop
        Print #1, mChr17 & "E.& OE." & mChr18 & mEmph & PSTR("For " & PubComp_Name, PageWidth - 5, , AlignRight) & mEmph1
        Print #1, ""
        Print #1, ""
        Print #1, PSTR("Authorised Signatory", PageWidth, , AlignRight)
        Print #1, Replace(Space(PageWidth), " ", "-")
        
        If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
            Print #1, "Received Tata Diesel Chassis as detailed above in satisfactory order & good"
            Print #1, "condition."
            Print #1, ""
            Print #1, ""
            Print #1, "Signature of Customer"
            Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
        Else
            Print #1, PSTR("1.Delivery of the Vehicle, Physically taken from ExShowroom in Perfect Condition.", 80)
            mHeader = mHeader + 1
            Print #1, PSTR("2.The Vehicle is supplied Subject to normal warranty given by the manufacturer &", 80)
            mHeader = mHeader + 1
            Print #1, PSTR(" no warranty or guarantee other than that given by the manufacturer shall be ", 80)
            mHeader = mHeader + 1
            Print #1, PSTR(" stipulated as applicable to this purchase.", 80)
            mHeader = mHeader + 1
            Print #1, PSTR("3.The warranty or guarantee of the vehicle would not be entertained under any", 80)
            mHeader = mHeader + 1
            Print #1, PSTR(" circumtances in case of cutting/tapping of wires for fittment of any unbranded", 80)
            mHeader = mHeader + 1
            Print #1, PSTR(" accessories in the vehicle", 80)
            mHeader = mHeader + 1
            Print #1, ""
            Print #1, ""
            Print #1, "Signature of Customer"
            Print #1, Replace(Space(PageWidth), " ", "-") & mChr17
        End If
        
        Print #1, mChr17 & RstDel!DelCh_UName & " " & RstDel!DelCh_UEntDt & Space(((PageWidth * 1.7) - Len("") - Len(RstDel!DelCh_UName & " " & RstDel!DelCh_UEntDt)) / 2) & "" & mChr18
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
        GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
    End If
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Function ProcAcPost(Optional CheckCtrls As Boolean) As Boolean
On Error GoTo lblExit
        Dim MsgStr$, rsCtrlAc As ADODB.Recordset, RsTemp As ADODB.Recordset
        'A/c Posting related declarations
        Dim I As Integer, mBookDocID$
        Dim LedgAry(9) As LedgRec, mResult As Byte, mNarr$, mInvPrefix$
        
        Set rsCtrlAc = New ADODB.Recordset
        rsCtrlAc.CursorLocation = adUseClient
        rsCtrlAc.Open "Select Interest_Ac,TDS_Ac,StampDuty_Ac,ServiceChrg_Ac, RegnFeeAc, InsuranceFeeAc From AcControls", G_FaCn, adOpenStatic, adLockReadOnly
        If rsCtrlAc.RecordCount <= 0 Then
            MsgStr = "Please Add Records in A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        If (Val(Txt(IntAmt)) <> 0 And (IsNull(rsCtrlAc!Interest_Ac) Or rsCtrlAc!Interest_Ac = "")) Then
            MsgStr = "Please define Interest A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        If (Val(Txt(TDSAmt)) <> 0 And (IsNull(rsCtrlAc!TDS_Ac) Or rsCtrlAc!TDS_Ac = "")) Then
            MsgStr = "Please define TDS A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        If (Val(Txt(StampChg)) <> 0 And (IsNull(rsCtrlAc!StampDuty_Ac) Or rsCtrlAc!StampDuty_Ac = "")) Then
            MsgStr = "Please define Stamp Duty A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        If (Val(Txt(SerChg)) <> 0 And (IsNull(rsCtrlAc!ServiceChrg_Ac) Or rsCtrlAc!ServiceChrg_Ac = "")) Then
            MsgStr = "Please define Service Charge A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        
        If (Val(Txt(RegFee)) <> 0 And (IsNull(rsCtrlAc!RegnFeeAc) Or rsCtrlAc!RegnFeeAc = "")) Then
            MsgStr = "Please define Registration Fee A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        
        If (Val(Txt(IncChg)) <> 0 And (IsNull(rsCtrlAc!InsuranceFeeAc) Or rsCtrlAc!InsuranceFeeAc = "")) Then
            MsgStr = "Please define Insurance Fee A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !"
            ProcAcPost = False
            GoTo lblExit
        End If
        
        If CheckCtrls Then 'Control setting found Ok
            ProcAcPost = True: Exit Function
        End If

        'Sale Party A/c
        mBookDocID = GCn.Execute("select OrdDocId from Veh_Order where Inv_DocId='" & InvdocId & "'").Fields(0).Value
        mInvPrefix = GCn.Execute("select " & xIsNull("Inv_Prefix", "") & " from Veh_Order where Inv_DocId='" & InvdocId & "'").Fields(0).Value
'        If TopCtrl1.TopText2 = "Edit" Then
'            InvDocId = Master!Inv_DocId
'        Else
'            InvDocId = RsInv!Inv_DocId    'Master!Inv_DocId
'        End If
        mNarr = "By Vehicle Delivery " & Txt(ChassisNo)
        mNarr = mNarr & " Invoice No." & mInvPrefix & Trim(DeCodeDocID(InvdocId, Document_No))
        I = 0
        If Val(Txt(IntAmt)) > 0 Then
            LedgAry(I).SubCode = rsCtrlAc!Interest_Ac
            LedgAry(I).AmtDr = Round(Val(Txt(IntAmt)), 2)
            LedgAry(I).Narration = mNarr & " Interest" '& " Telco Inv. No." & Txt(TelcoInvNo)
            LedgAry(I).ContraSub = Txt(Party).Tag
            I = I + 1
            LedgAry(I).SubCode = Txt(Party).Tag
            LedgAry(I).AmtCr = Round(Val(Txt(IntAmt)), 2)
            LedgAry(I).Narration = mNarr & " Interest" '& " Telco Inv. No." & Txt(TelcoInvNo)
            LedgAry(I).ContraSub = rsCtrlAc!Interest_Ac
        End If
        'TDS A/c
        If Val(Txt(TDSAmt)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = Txt(Party).Tag
            LedgAry(I).AmtDr = Round(Val(Txt(TDSAmt)), 2)
            LedgAry(I).Narration = mNarr & " TDS" '& " Telco Inv. No." & Txt(TelcoInvNo)
            LedgAry(I).ContraSub = rsCtrlAc!TDS_Ac
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!TDS_Ac
            LedgAry(I).AmtCr = Round(Val(Txt(TDSAmt)), 2)
            LedgAry(I).Narration = mNarr & " TDS" '& " Telco Inv. No." & Txt(TelcoInvNo)
            LedgAry(I).ContraSub = Txt(Party).Tag
        End If
        'Service Charge
        If Val(Txt(SerChg)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = Txt(Party).Tag
            LedgAry(I).AmtDr = Round(Val(Txt(SerChg)), 2)
            LedgAry(I).Narration = mNarr & " Service Chrg." '& " Telco Inv. No." & Txt(TelcoInvNo)
            LedgAry(I).ContraSub = rsCtrlAc!ServiceChrg_Ac
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!ServiceChrg_Ac
            LedgAry(I).AmtCr = Round(Val(Txt(SerChg)), 2)
            LedgAry(I).Narration = mNarr & " Service Chrg."
            LedgAry(I).ContraSub = Txt(Party).Tag
        End If
        'Stamp & Duty Amt
        If Val(Txt(StampChg)) <> 0 Then
            I = I + 1
            LedgAry(I).SubCode = Txt(Party).Tag
            LedgAry(I).AmtDr = Round(Val(Txt(StampChg)), 2)
            LedgAry(I).Narration = mNarr & " Service Chrg." '& " Telco Inv. No." & Txt(TelcoInvNo)
            LedgAry(I).ContraSub = rsCtrlAc!StampDuty_Ac
            I = I + 1
            LedgAry(I).SubCode = rsCtrlAc!StampDuty_Ac
            LedgAry(I).AmtCr = Round(Val(Txt(StampChg)), 2)
            LedgAry(I).Narration = mNarr & " Service Chrg."
            LedgAry(I).ContraSub = Txt(Party).Tag
        End If
        
        
        If GCn.Execute("Select PostRegnFeeYn From Syctrl").Fields(0) = 1 Then
            If Val(Txt(RegFee)) <> 0 Then
                I = I + 1
                LedgAry(I).SubCode = Txt(Party).Tag
                LedgAry(I).AmtDr = Round(Val(Txt(RegFee)), 2)
                LedgAry(I).Narration = mNarr & " Registration Fee" '& " Telco Inv. No." & Txt(TelcoInvNo)
                LedgAry(I).ContraSub = rsCtrlAc!RegnFeeAc
                I = I + 1
                LedgAry(I).SubCode = rsCtrlAc!RegnFeeAc
                LedgAry(I).AmtCr = Round(Val(Txt(RegFee)), 2)
                LedgAry(I).Narration = mNarr & " Registration Fee"
                LedgAry(I).ContraSub = Txt(Party).Tag
            End If
        End If
        
        If GCn.Execute("Select PostInsuranceFeeYn From Syctrl").Fields(0) = 1 Then
            If Val(Txt(IncChg)) <> 0 Then
                I = I + 1
                LedgAry(I).SubCode = Txt(Party).Tag
                LedgAry(I).AmtDr = Round(Val(Txt(IncChg)), 2)
                LedgAry(I).Narration = mNarr & " Insurance Fee" '& " Telco Inv. No." & Txt(TelcoInvNo)
                LedgAry(I).ContraSub = rsCtrlAc!InsuranceFeeAc
                I = I + 1
                LedgAry(I).SubCode = rsCtrlAc!InsuranceFeeAc
                LedgAry(I).AmtCr = Round(Val(Txt(IncChg)), 2)
                LedgAry(I).Narration = mNarr & " Insurance Fee"
                LedgAry(I).ContraSub = Txt(Party).Tag
            End If
        End If
        
        
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaV, Txt(TxtDocID), CDate(Txt(VDate)))
        If mResult <> 1 Then
            MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
            ProcAcPost = False
        Else
            ProcAcPost = True
        End If
lblExit:
If MsgStr <> "" Then
    MsgBox MsgStr, vbCritical, "A/c Posting"
ElseIf err.NUMBER > 0 Then
    MsgBox err.Description, vbCritical, "A/c Posting"
End If
Set rsCtrlAc = Nothing
Set RsTemp = Nothing
End Function

Private Function FillCustTrns(OrderID As String, FillData As Boolean, CalcIntt As Boolean, Optional mIntRate As Single, Optional mRebDays As Integer, Optional mInvDate As Date) As Double
If OrderID = "" Then Exit Function
Dim I As Integer, mDays As Integer, mIntt As Double, mTotIntt As Double
Dim Rst As ADODB.Recordset
If FillData Then
    mTotIntt = 0
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    GSQL = "Select DocId,left(docid,2) + '/' + " & cCStr(cTrim("right(docid,8)")) & " as VNO,V_Date,VouType.Description, AMOUNT as Amt,DrCr, " & cIIF("IntDays=0", "''", "IntDays") & " as InttDays, " & cIIF("IntValue=0", "0", "IntValue") & " as IntAmt " & _
    " From Rect left join " & FaTable("Voucher_Type") & " VouType on " & _
    " Rect.V_Type=VouType.V_Type " & _
    " where Ord_DocId='" & OrderID & "' and V_Date<=" & ConvertDate(mInvDate) & ""
    Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
        Do While Rst.EOF = False
            mTotIntt = mTotIntt + Val(Rst!IntAmt)
            Rst.MoveNext
        Loop
    End If
    Set FGrid.DataSource = Rst
    Ini_Grid
End If
If CalcIntt Then
    'Calculate fresh interest & fill intt value in grid
    For I = 1 To FGrid.Rows - 1
        mDays = mInvDate - CDate(FGrid.TextMatrix(I, Col_Date)) - mRebDays
        If FGrid.TextMatrix(I, Col_DrCr) = "C" Then
            mIntt = Round((Val(FGrid.TextMatrix(I, Col_Amt)) * mIntRate * mDays) / 36500, 0)
        Else
            mIntt = -1 * Round((Val(FGrid.TextMatrix(I, Col_Amt)) * mIntRate * mDays) / 36500, 0)
        End If
        FGrid.TextMatrix(I, Col_IntDays) = IIf(mDays <> 0, mDays, "")
        FGrid.TextMatrix(I, Col_IntAmt) = IIf(mIntt <> 0, Format(mIntt, "0.00"), "")
        mTotIntt = mTotIntt + mIntt
    Next I
End If
Txt(IntAmt) = IIf(mTotIntt = 0, "", Format(mTotIntt, "0.00"))
'        Dim RstTemp As Recordset, mDays As Integer, CalDays As Integer
'        Set RstTemp = GCn.Execute("Select Max(V_Date) as LastDate from Rect where Ord_DocId = '" & txt(BookNo).Tag & "'  and V_Date <= " & ConvertDate(txt(Vdate)) & "")
'        If RstTemp.RecordCount  > 0 Then
'            If IsNull(RstTemp!LastDate) Then
'                mDays = 0
'            Else
'                mDays = DateDiff("D", RstTemp!LastDate, txt(Vdate))
'            End If
'        End If
'        If mDays  > Val(txt(RebDays)) Then
'           CalDays = mDays - Val(txt(RebDays))
'           txt(IntAmt).Text = Format((Val(txt(AdvAmt).Text) * Val(txt(IntPer).Text) / 100) * CalDays / 365, "0.00")
'        Else
'           txt(IntAmt).Text = ""
'        End If
'        txt(TDSAmt).Text = Format(Val(txt(TDSPer).Text) * Val(txt(IntAmt).Text) / 100, "0.00")
'        Set RstTemp = Nothing
  Set Rst = Nothing
End Function
Private Sub Ini_Grid()
    With FGrid
        .RowHeightMin = PubGridRowHeight
        .Cols = 9
        .height = (.RowHeight(0) * 6) + 15
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColWidth(0) = 400

        .TextMatrix(0, Col_DocID) = "DocID"
        .ColAlignment(Col_DocID) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_DocID) = flexAlignLeftCenter
        .ColWidth(Col_DocID) = 0

        .TextMatrix(0, Col_VNo) = "VoucherNo."
        .ColAlignment(Col_VNo) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_VNo) = flexAlignLeftCenter
        .ColWidth(Col_VNo) = 1500
        
        .TextMatrix(0, Col_Date) = "Date"
        .ColAlignment(Col_Date) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_Date) = flexAlignLeftCenter
        .ColWidth(Col_Date) = 2000

        .TextMatrix(0, Col_VType) = "Type"
        .ColAlignment(Col_VType) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_VType) = flexAlignLeftCenter
        .ColWidth(Col_VType) = 3000
        
        .TextMatrix(0, Col_Amt) = "Amount"
        .ColAlignment(Col_Amt) = flexAlignRightCenter
        .ColAlignmentFixed(Col_Amt) = flexAlignCenterCenter
        .ColWidth(Col_Amt) = 1500
        
        .TextMatrix(0, Col_DrCr) = "Dr/Cr"
        .ColAlignment(Col_DrCr) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_DrCr) = flexAlignCenterCenter
        .ColWidth(Col_DrCr) = 700

        .TextMatrix(0, Col_IntDays) = "InttDays"
        .ColAlignment(Col_IntDays) = flexAlignRightCenter
        .ColAlignmentFixed(Col_IntDays) = flexAlignCenterCenter
        .ColWidth(Col_IntDays) = 800

        .TextMatrix(0, Col_IntAmt) = "Intt.Amt"
        .ColAlignment(Col_IntAmt) = flexAlignRightCenter
        .ColAlignmentFixed(Col_IntAmt) = flexAlignCenterCenter
        .ColWidth(Col_IntAmt) = 1200
    End With
    
    With FGrid1
'        .left = 7455
'        .top = 3030
        .Cols = 4
        .RowHeightMin = PubGridRowHeight
        .height = (.RowHeight(0) * 8) + 15
        
        .TextMatrix(0, 0) = "S.No."
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 550

        .TextMatrix(0, ItemCode) = "ItemCode"
        .ColAlignment(ItemCode) = flexAlignLeftCenter
        .ColWidth(ItemCode) = 0
        
        .TextMatrix(0, Description) = "Description"
        .ColAlignment(Description) = flexAlignLeftCenter
        .ColWidth(Description) = 2500
                
        .TextMatrix(0, DefVal) = "Value"
        .ColAlignment(DefVal) = flexAlignLeftCenter
        .ColWidth(DefVal) = 1150
        .ColWidth(PIndex) = 0
    End With
    Frame1.left = 2625
    Frame1.top = 390
End Sub

Private Sub FillCheckListGrid(Model As String, ChassisNo As String)

Dim I As Integer
Dim Rst As ADODB.Recordset
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    GSQL = "Select VCL.MODEL,VCL.ChassisNo,VCL.Item_Code,MCLM.Item_Description,VCL.Default_Value " & _
        " From (Veh_CheckList VCL left join ModelCheckList MCL on VCL.Model + VCL.Item_Code=MCL.Model + MCL.Item_Code) " & _
        " Left Join ModelCheckListMast MCLM on VCL.Item_Code=MCLM.Item_Code " & _
        " where VCL.Model='" & Model & "' and VCL.ChassisNo='" & ChassisNo & "' order by MCLM.Report_Index"
    Rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    FGrid1.Rows = 1
    If Rst.RecordCount > 0 Then
        I = 1
        Do Until Rst.EOF
            FGrid1.AddItem ""
            With FGrid1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, ItemCode) = Rst!Item_Code
                .TextMatrix(I, Description) = Rst!Item_Description
                .TextMatrix(I, DefVal) = Rst!Default_Value
            End With
            Rst.MoveNext
            I = I + 1
        Loop
    Else
    
'        If GCn.Execute("Select Item_Code from ModelCheckList where Model=''").RecordCount  > 0 Then
'            MsgBox "Vehicle Check List Items not found, "
'        Set rst = New ADODB.Recordset
'        rst.CursorLocation = adUseClient
'        GSQL = "Select VCL.MODEL,VCL.ChassisNo,VCL.Item_Code,MCL.Description,VCL.Value " & _
'            " From Veh_CheckList VCL left join ModelCheckList MCL on " & _
'            " VCL.Item_Code=MCL.Item_Code " & _
'            " where VCL.Model='" & Model & "' and VCL.ChassisNo='" & ChassisNo & "'"
'        rst.Open GSQL, GCn, adOpenStatic, adLockReadOnly
        FGrid1.AddItem FGrid.Rows
    End If
    FGrid1.FixedRows = 1
'    FGrid1.Redraw = True
    'Set FGrid.DataSource = rst
    'Ini_Grid
Set Rst = Nothing
End Sub
Private Sub CLEARTEXT()
    CabNo = ""
    Transaxleno = ""
    Rearaxleno = ""
    fipno = ""
    alternatorno = ""
    startingmono = ""
    steeringboxno = ""
    accompno = ""
    tyremake = ""
    tyremakeA = ""
    tyremakeB = ""
    tyremakeC = ""
    tyremakeD = ""
    tyremakeE = ""
    toolsandequip = ""
    connecteddocu = ""
    kmsreading = ""
    pdion = ""
    batteryno = ""
    deliverytakenby = ""
End Sub

