VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmPurBill 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Spare Purchase Entry"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   13560
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10335
   ScaleWidth      =   13560
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt 
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
      Index           =   35
      Left            =   2115
      MaxLength       =   40
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   6615
      Width           =   1230
   End
   Begin VB.TextBox txt 
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
      Height          =   240
      Index           =   34
      Left            =   5265
      MaxLength       =   40
      TabIndex        =   149
      Top             =   5880
      Width           =   1230
   End
   Begin MSDataGridLib.DataGrid DGTrans 
      Height          =   4935
      Left            =   8460
      Negotiate       =   -1  'True
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   8370
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
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
      Caption         =   "Transport Help"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdTransPost 
      Caption         =   "Post Trans."
      Height          =   315
      Left            =   9345
      TabIndex        =   148
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Re-Update Transaction"
      Height          =   315
      Left            =   6735
      TabIndex        =   147
      Top             =   15
      Width           =   2535
   End
   Begin VB.TextBox txt 
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
      Height          =   240
      Index           =   20
      Left            =   5265
      MaxLength       =   40
      TabIndex        =   28
      Top             =   5625
      Width           =   1230
   End
   Begin VB.TextBox txt 
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
      Height          =   240
      Index           =   21
      Left            =   5265
      MaxLength       =   40
      TabIndex        =   29
      Top             =   6135
      Width           =   1230
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   240
      Index           =   17
      Left            =   2130
      MaxLength       =   40
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5595
      Width           =   1230
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   19
      Left            =   5265
      MaxLength       =   40
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5370
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   7665
      TabIndex        =   139
      Top             =   2460
      Visible         =   0   'False
      Width           =   345
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
      Left            =   705
      TabIndex        =   138
      Top             =   3405
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   1365
      TabIndex        =   71
      Top             =   8010
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   255
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   0
         Width           =   4125
         _ExtentX        =   7276
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
         BackColor       =   16777152
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
      Height          =   2955
      Left            =   2400
      Negotiate       =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   7695
      Visible         =   0   'False
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   5212
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
      Caption         =   "Party Help"
      ColumnCount     =   5
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
         DataField       =   "Add1"
         Caption         =   "Add1"
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
         DataField       =   "Add2"
         Caption         =   "Add2"
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
         DataField       =   "Add3"
         Caption         =   "Add3"
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
         DataField       =   "City"
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
            ColumnWidth     =   4545.071
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
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
      Height          =   240
      Index           =   33
      Left            =   10425
      TabIndex        =   32
      Top             =   5880
      Width           =   1230
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   23
      Left            =   6630
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5625
      Width           =   1725
   End
   Begin VB.TextBox txt 
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
      Height          =   240
      Index           =   28
      Left            =   2130
      MaxLength       =   40
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1230
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   32
      Left            =   10425
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6390
      Width           =   1230
   End
   Begin VB.TextBox txt 
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
      Height          =   240
      Index           =   30
      Left            =   9615
      TabIndex        =   33
      Top             =   6135
      Width           =   570
   End
   Begin VB.TextBox txt 
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
      Height          =   240
      Index           =   31
      Left            =   10425
      MaxLength       =   40
      TabIndex        =   34
      Top             =   6135
      Width           =   1230
   End
   Begin VB.TextBox txt 
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
      Index           =   29
      Left            =   7575
      MaxLength       =   8
      TabIndex        =   132
      Text            =   "ChalSrlN"
      Top             =   1695
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Frame FrmDetail 
      BackColor       =   &H00CAF1FD&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   2205
      Left            =   4800
      TabIndex        =   99
      Top             =   2820
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
         Left            =   2100
         TabIndex        =   130
         Top             =   1395
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MRP Taxpaid"
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
         Left            =   75
         TabIndex        =   129
         Top             =   1410
         Width           =   1095
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Part Name>"
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
         Index           =   1
         Left            =   1140
         TabIndex        =   128
         Top             =   465
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Name"
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
         Index           =   9
         Left            =   75
         TabIndex        =   127
         Top             =   465
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taxable"
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
         Index           =   10
         Left            =   75
         TabIndex        =   126
         Top             =   1635
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Index           =   11
         Left            =   60
         TabIndex        =   125
         Top             =   0
         Width           =   6180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Stock"
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
         Index           =   12
         Left            =   3930
         TabIndex        =   124
         Top             =   915
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MRP Taxable"
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
         Index           =   13
         Left            =   75
         TabIndex        =   123
         Top             =   1185
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Paid"
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
         Index           =   14
         Left            =   75
         TabIndex        =   122
         Top             =   1875
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   15
         Left            =   2805
         TabIndex        =   121
         Top             =   930
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
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
         Index           =   16
         Left            =   4800
         TabIndex        =   120
         Top             =   1650
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Name"
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
         Index           =   17
         Left            =   75
         TabIndex        =   119
         Top             =   675
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00BBDBB3&
         BackStyle       =   0  'Transparent
         Caption         =   "Pur. Rate"
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
         Index           =   18
         Left            =   3810
         TabIndex        =   118
         Top             =   1185
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High"
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
         Left            =   4800
         TabIndex        =   117
         Top             =   1410
         Width           =   375
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
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
         Left            =   5340
         TabIndex        =   116
         Top             =   1410
         Width           =   900
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Part Local Name>"
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
         Left            =   1140
         TabIndex        =   115
         Top             =   675
         Width           =   1665
      End
      Begin VB.Label LblFrm 
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
         Index           =   4
         Left            =   2085
         TabIndex        =   114
         Top             =   1170
         Width           =   375
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000.000"
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
         Index           =   8
         Left            =   5145
         TabIndex        =   113
         Top             =   930
         Width           =   1110
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
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
         Index           =   13
         Left            =   5340
         TabIndex        =   112
         Top             =   1650
         Width           =   900
      End
      Begin VB.Label LblFrm 
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
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   10
         Left            =   3270
         TabIndex        =   111
         Top             =   1635
         Width           =   375
      End
      Begin VB.Label LblFrm 
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
         Index           =   7
         Left            =   2100
         TabIndex        =   110
         Top             =   1875
         Width           =   375
      End
      Begin VB.Label LblFrm 
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
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   6
         Left            =   2100
         TabIndex        =   109
         Top             =   1635
         Width           =   375
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
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
         Index           =   14
         Left            =   5340
         TabIndex        =   108
         Top             =   1185
         Width           =   900
      End
      Begin VB.Label LblFrm 
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
         Index           =   11
         Left            =   3270
         TabIndex        =   107
         Top             =   1875
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last"
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
         Index           =   23
         Left            =   4815
         TabIndex        =   106
         Top             =   1185
         Width           =   345
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000.00"
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
         Index           =   9
         Left            =   2730
         TabIndex        =   105
         Top             =   1185
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   21
         Left            =   1800
         TabIndex        =   104
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
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
         Index           =   20
         Left            =   75
         TabIndex        =   103
         Top             =   270
         Width           =   690
      End
      Begin VB.Label LblFrm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<Part No>"
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
         Index           =   0
         Left            =   1140
         TabIndex        =   102
         Top             =   255
         Width           =   900
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Bin Loca>"
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
         Left            =   4920
         TabIndex        =   101
         Top             =   255
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bin Location"
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
         Index           =   47
         Left            =   3765
         TabIndex        =   100
         Top             =   255
         Width           =   1035
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
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   240
      Index           =   18
      Left            =   2130
      MaxLength       =   40
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5850
      Width           =   1230
   End
   Begin VB.TextBox txt 
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
      Height          =   240
      Index           =   22
      Left            =   5265
      TabIndex        =   30
      Top             =   6390
      Width           =   1230
   End
   Begin VB.TextBox txt 
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
      Height          =   240
      Index           =   27
      Left            =   2130
      MaxLength       =   40
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6105
      Width           =   1230
   End
   Begin MSDataGridLib.DataGrid DGOrdPart 
      Height          =   2625
      Left            =   1920
      Negotiate       =   -1  'True
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   7260
      Visible         =   0   'False
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   4630
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
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
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   9
      Left            =   6675
      MaxLength       =   20
      TabIndex        =   6
      Top             =   405
      Width           =   1425
   End
   Begin VB.Frame FrmSel 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   -8730
      TabIndex        =   84
      Top             =   2955
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CommandButton CmdSel 
         BackColor       =   &H00E2D5C0&
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   1800
         Width           =   1155
      End
      Begin VB.CommandButton CmdSel 
         BackColor       =   &H00E2D5C0&
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
         Left            =   3270
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   1800
         Width           =   1155
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGridSel 
         Height          =   1695
         Left            =   45
         TabIndex        =   85
         Top             =   30
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   15525079
         Cols            =   7
         BackColorFixed  =   14940925
         ForeColorFixed  =   8388608
         BackColorSel    =   16711680
         BackColorBkg    =   14737632
         BackColorUnpopulated=   14865856
         GridColor       =   14940925
         GridColorFixed  =   8421631
         GridLinesFixed  =   1
         AllowUserResizing=   3
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   "            |MR No.            |MR Date      |SupplierChl No|Supplier Chl Date |MR Value|Docid"
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
         _Band(0).Cols   =   7
      End
   End
   Begin VB.TextBox txt 
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
      Index           =   26
      Left            =   1425
      MaxLength       =   50
      TabIndex        =   19
      Top             =   2190
      Width           =   3735
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1395
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1425
      Width           =   3735
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   6675
      MaxLength       =   10
      TabIndex        =   12
      Text            =   "0123456789"
      Top             =   660
      Width           =   1425
   End
   Begin VB.TextBox txt 
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
      Index           =   1
      Left            =   10245
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1590
      Width           =   975
   End
   Begin VB.TextBox txt 
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
      Index           =   25
      Left            =   1425
      MaxLength       =   8
      TabIndex        =   9
      Top             =   915
      Width           =   1470
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   24
      Left            =   1425
      MaxLength       =   40
      TabIndex        =   10
      Text            =   "01234567890123456789012345"
      Top             =   1170
      Width           =   3735
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   661
   End
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   1995
      Left            =   1020
      Negotiate       =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   7830
      Visible         =   0   'False
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   3519
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Part No"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "MRP"
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
      BeginProperty Column05 
         DataField       =   "CurStk"
         Caption         =   "Cur Stock"
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
            DividerStyle    =   3
            ColumnWidth     =   2715.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3600
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1065.26
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FBFBFB&
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   1425
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "aaaa"
      Top             =   405
      Width           =   3720
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
      Height          =   240
      Index           =   3
      Left            =   9390
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1320
      Width           =   1275
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Left            =   660
      TabIndex        =   20
      Top             =   4200
      Visible         =   0   'False
      Width           =   690
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
      Height          =   240
      Index           =   0
      Left            =   9390
      MaxLength       =   21
      TabIndex        =   1
      Top             =   510
      Width           =   2250
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Index           =   16
      Left            =   2130
      MaxLength       =   40
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "99999999.99"
      Top             =   5340
      Width           =   1230
   End
   Begin VB.TextBox txt 
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
      Index           =   13
      Left            =   6675
      MaxLength       =   15
      TabIndex        =   15
      Text            =   "012345678901234"
      Top             =   915
      Width           =   1425
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   12
      Left            =   1425
      MaxLength       =   30
      TabIndex        =   16
      Top             =   1935
      Width           =   3735
   End
   Begin VB.TextBox txt 
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
      Index           =   11
      Left            =   4665
      MaxLength       =   4
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txt 
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
      Left            =   1425
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1680
      Width           =   1470
   End
   Begin VB.TextBox txt 
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
      Index           =   15
      Left            =   6675
      MaxLength       =   12
      TabIndex        =   18
      Top             =   1425
      Width           =   1425
   End
   Begin VB.TextBox txt 
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
      Index           =   14
      Left            =   6675
      MaxLength       =   15
      TabIndex        =   17
      Text            =   "012345678901234"
      Top             =   1170
      Width           =   1425
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   3900
      MaxLength       =   12
      TabIndex        =   8
      Text            =   "29-APR-2002"
      Top             =   660
      Width           =   1245
   End
   Begin VB.TextBox txt 
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
      Left            =   1425
      MaxLength       =   15
      TabIndex        =   7
      Top             =   660
      Width           =   1470
   End
   Begin VB.TextBox txt 
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
      Left            =   9390
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1050
      Width           =   1275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   2550
      Left            =   90
      TabIndex        =   21
      Top             =   2475
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   4498
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   39
      BackColorFixed  =   14940925
      ForeColorFixed  =   8388608
      BackColorSel    =   16777215
      ForeColorSel    =   12582912
      BackColorBkg    =   14737632
      BackColorUnpopulated=   14865856
      GridColor       =   0
      GridColorFixed  =   12632319
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
      FormatString    =   $"frmPurBill.frx":0000
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   39
   End
   Begin MSDataGridLib.DataGrid DGGod 
      Height          =   2145
      Left            =   6120
      Negotiate       =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   8445
      Visible         =   0   'False
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3784
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
      Caption         =   "Tax Form Help"
      ColumnCount     =   1
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
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   4935
      Left            =   6540
      Negotiate       =   -1  'True
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   8115
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
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
      Caption         =   "Tax Form Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Form Description"
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
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGPONo 
      Height          =   2790
      Left            =   300
      Negotiate       =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   7935
      Visible         =   0   'False
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   4921
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "PO Reg. No."
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
         DataField       =   "Order_Reg_Dt"
         Caption         =   "PO Reg.Date"
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
      BeginProperty Column03 
         DataField       =   "OurDocNo"
         Caption         =   "PO No. "
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
         DataField       =   "v_date"
         Caption         =   "PO Date"
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
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total SFC Amount"
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
      Index           =   54
      Left            =   225
      TabIndex        =   154
      Top             =   6615
      Width           =   1485
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
      Left            =   270
      TabIndex        =   153
      Top             =   6660
      Width           =   45
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   53
      Left            =   2970
      TabIndex        =   152
      Top             =   690
      Width           =   390
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
      Index           =   22
      Left            =   5100
      TabIndex        =   151
      Top             =   5880
      Width           =   45
   End
   Begin VB.Label Label3 
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
      Index           =   52
      Left            =   3465
      TabIndex        =   150
      Top             =   5880
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Expences Paid (Other than Bill Amount)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   630
      Index           =   51
      Left            =   8595
      TabIndex        =   146
      Top             =   5325
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      Height          =   1410
      Left            =   8505
      Top             =   5295
      Width           =   3225
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Amount"
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
      Index           =   38
      Left            =   3465
      TabIndex        =   145
      Top             =   5625
      Width           =   960
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
      Index           =   19
      Left            =   5100
      TabIndex        =   144
      Top             =   5625
      Width           =   45
   End
   Begin VB.Label Label3 
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
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   39
      Left            =   3465
      TabIndex        =   143
      Top             =   6135
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   20
      Left            =   5100
      TabIndex        =   142
      Top             =   6135
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
      Height          =   255
      Index           =   18
      Left            =   5100
      TabIndex        =   141
      Top             =   5370
      Width           =   45
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   37
      Left            =   3465
      TabIndex        =   140
      Top             =   5370
      Width           =   1530
   End
   Begin VB.Label Label3 
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
      Height          =   225
      Index           =   50
      Left            =   8595
      TabIndex        =   137
      Top             =   5880
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Bill Amount"
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
      Index           =   41
      Left            =   6915
      TabIndex        =   136
      Top             =   5355
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Amount"
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
      Index           =   49
      Left            =   8595
      TabIndex        =   135
      Top             =   6390
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Tax @"
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
      Index           =   48
      Left            =   8595
      TabIndex        =   134
      Top             =   6150
      Width           =   960
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
      Left            =   5340
      TabIndex        =   133
      Top             =   1815
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label lblVPrefix2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ChalVPrefix"
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
      Height          =   225
      Left            =   6525
      TabIndex        =   131
      Top             =   1695
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction"
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
      Index           =   40
      Left            =   3435
      TabIndex        =   98
      Top             =   6390
      Width           =   840
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
      Index           =   21
      Left            =   5115
      TabIndex        =   97
      Top             =   6390
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Oil Value"
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
      Index           =   6
      Left            =   225
      TabIndex        =   96
      Top             =   6375
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   1995
      TabIndex        =   95
      Top             =   6360
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Spare Value"
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
      Index           =   46
      Left            =   225
      TabIndex        =   94
      Top             =   6120
      Width           =   1455
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
      Left            =   1995
      TabIndex        =   93
      Top             =   6105
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goods Receipt"
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
      Index           =   31
      Left            =   5280
      TabIndex        =   87
      Top             =   405
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Remarks"
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
      Index           =   45
      Left            =   105
      TabIndex        =   83
      Top             =   2190
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Permit Form"
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
      Index           =   4
      Left            =   105
      TabIndex        =   81
      Top             =   1425
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Permit No."
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
      Index           =   24
      Left            =   5280
      TabIndex        =   80
      Top             =   660
      Width           =   930
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   1470
      Left            =   8250
      Top             =   435
      Width           =   3465
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VPrefix"
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
      Height          =   225
      Left            =   9540
      TabIndex        =   79
      Top             =   1605
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill  No."
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
      Left            =   8310
      TabIndex        =   78
      Top             =   1605
      Width           =   615
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
      Height          =   225
      Index           =   92
      Left            =   9315
      TabIndex        =   77
      Top             =   1605
      Width           =   45
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division        :"
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
      Left            =   8310
      TabIndex        =   76
      Top             =   780
      Width           =   1065
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code    :"
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
      Left            =   10140
      TabIndex        =   75
      Top             =   780
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   44
      Left            =   105
      TabIndex        =   74
      Top             =   915
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   43
      Left            =   105
      TabIndex        =   73
      Top             =   1170
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   23
      Left            =   9315
      TabIndex        =   69
      Top             =   510
      Width           =   45
   End
   Begin VB.Label Label3 
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
      Index           =   42
      Left            =   8295
      TabIndex        =   68
      Top             =   510
      Width           =   585
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
      Height          =   285
      Index           =   17
      Left            =   1995
      TabIndex        =   67
      Top             =   5835
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Order Discount"
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
      Index           =   36
      Left            =   210
      TabIndex        =   66
      Top             =   5865
      Width           =   1695
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
      Height          =   285
      Index           =   16
      Left            =   1995
      TabIndex        =   65
      Top             =   5580
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Discount"
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
      Index           =   35
      Left            =   210
      TabIndex        =   64
      Top             =   5610
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   15
      Left            =   1995
      TabIndex        =   63
      Top             =   5355
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Index           =   34
      Left            =   210
      TabIndex        =   62
      Top             =   5355
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
      Index           =   14
      Left            =   10785
      TabIndex        =   61
      Top             =   5040
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Goods Amount"
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
      Index           =   33
      Left            =   9030
      TabIndex        =   60
      Top             =   5040
      Width           =   1680
   End
   Begin VB.Label LblAmt 
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
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   10995
      TabIndex        =   59
      Top             =   5040
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supply Mode "
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
      Index           =   32
      Left            =   5280
      TabIndex        =   58
      Top             =   915
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transporter"
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
      Index           =   30
      Left            =   105
      TabIndex        =   57
      Top             =   1935
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Case"
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
      Index           =   29
      Left            =   3480
      TabIndex        =   56
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Case Marking"
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
      Index           =   28
      Left            =   105
      TabIndex        =   55
      Top             =   1680
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GR / Bilty Date"
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
      Index           =   27
      Left            =   5280
      TabIndex        =   54
      Top             =   1440
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GR / Bilty No."
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
      Index           =   26
      Left            =   5280
      TabIndex        =   53
      Top             =   1170
      Width           =   1065
   End
   Begin VB.Label LblPQty 
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
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   7815
      TabIndex        =   52
      Top             =   5040
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity(Phy)"
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
      Index           =   22
      Left            =   5985
      TabIndex        =   51
      Top             =   5040
      Width           =   1530
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
      Index           =   4
      Left            =   7605
      TabIndex        =   50
      Top             =   5040
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Inv No. "
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
      Index           =   5
      Left            =   105
      TabIndex        =   49
      Top             =   675
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   105
      TabIndex        =   48
      Top             =   405
      Width           =   705
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
      Index           =   3
      Left            =   4395
      TabIndex        =   47
      Top             =   5040
      Width           =   120
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
      Left            =   1785
      TabIndex        =   46
      Top             =   5040
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity(Doc)"
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
      Left            =   2730
      TabIndex        =   45
      Top             =   5040
      Width           =   1560
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
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2085
      TabIndex        =   44
      Top             =   5040
      Width           =   105
   End
   Begin VB.Label LblDQty 
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
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4575
      TabIndex        =   43
      Top             =   5040
      Width           =   465
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
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   7
      Left            =   210
      TabIndex        =   41
      Top             =   5040
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
      Height          =   225
      Index           =   1
      Left            =   1560
      TabIndex        =   40
      Top             =   5040
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   91
      Left            =   9315
      TabIndex        =   39
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   93
      Left            =   9315
      TabIndex        =   38
      Top             =   1335
      Width           =   45
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   0
      Left            =   8310
      TabIndex        =   37
      Top             =   1335
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
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
      Left            =   8310
      TabIndex        =   36
      Top             =   1050
      Width           =   690
   End
End
Attribute VB_Name = "frmPurBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BackColorSelEnter As String = &HF8D7FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const ChalVType As String = "SXGR"
Private Const PurCashVType As String = "SXPIC"
Private Const PurCrVType As String = "SXPIR"

Dim mCheckNegetiveStockSiteWise As Boolean
Dim RsParty As ADODB.Recordset
Dim rsPONo As ADODB.Recordset
Dim rsGod As ADODB.Recordset
Dim rsForm As ADODB.Recordset
Dim rsForm31 As ADODB.Recordset
Dim rsTrans As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim rsCtrlAc As ADODB.Recordset


Dim mReposting As Boolean

Dim FirmAddFlag As Byte
Dim GridKey As Integer
Dim DocID As String * 21
Dim mVType As String
Dim VoucherEditFlag As Boolean
Dim vPrefix As String
Dim mPartyType As Byte         'Used to Detect Part Rate in GetRate Function

Private Const TxtDocID As Byte = 0
Private Const OPtSel As Byte = 9
Private Const SerialNo As Byte = 1
Private Const VDate As Byte = 2
Private Const VType As Byte = 3
Private Const Party As Byte = 4
Private Const SuppChlNo As Byte = 5
Private Const SuppChlDate As Byte = 6
Private Const FormType As Byte = 24
Private Const PermitType As Byte = 7
Private Const FormNo As Byte = 8
'Private Const ChlType As Byte = 9
Private Const CaseMark          As Byte = 10
Private Const CaseNo            As Byte = 11
Private Const Transport         As Byte = 12
Private Const LC                As Byte = 25
Private Const Remark            As Byte = 26
Private Const SupplyMode        As Byte = 13
Private Const GrNo              As Byte = 14
Private Const GrDate            As Byte = 15
Private Const TOTAmt            As Byte = 16
Private Const TotDis            As Byte = 17
Private Const TotOrdDis         As Byte = 18
Private Const TotGoods          As Byte = 19
Private Const TaxAmt            As Byte = 20
Private Const Addition          As Byte = 21
Private Const Deduction         As Byte = 22
Private Const NetAmt            As Byte = 23
Private Const SprAmt            As Byte = 27
Private Const OilAmt            As Byte = 28
Private Const SerialNo2         As Byte = 29
Private Const EntryTaxPer       As Byte = 30
Private Const EntryTaxAmt       As Byte = 31
Private Const TotPurAmt         As Byte = 32
Private Const Transportation    As Byte = 33
Private Const SatAmt            As Byte = 34
Private Const SFCAmt            As Byte = 35
'TotPurAmt
' Col Declaration

Private Const PONo As Byte = 1
Private Const PNo As Byte = 2
Private Const Unit As Byte = 3
Private Const MRP As Byte = 4
Private Const Taxable As Byte = 5
Private Const DQty As Byte = 6
Private Const PQty As Byte = 7
Private Const FRate As Byte = 8 'NDP
Private Const Amt  As Byte = 9
Private Const DisPer  As Byte = 10
Private Const DisRs  As Byte = 11
Private Const DisOrd  As Byte = 12
Private Const DisOrdRs  As Byte = 13

Private Const SFCPer As Byte = 14
Private Const SFCAmt1 As Byte = 15

Private Const TaxPer As Byte = 16
Private Const TaxAmt1 As Byte = 17
Private Const SatPer As Byte = 18
Private Const SatAmt1 As Byte = 19

Private Const ItemVal As Byte = 20
Private Const Godown As Byte = 21
Private Const NDP  As Byte = 22
Private Const PartSrlNo As Byte = 23         ' Part Serial N
Private Const PName As Byte = 24
Private Const LName As Byte = 25
Private Const MRPStkTB As Byte = 26
Private Const MRPStkTP As Byte = 27
Private Const TBStk As Byte = 28
Private Const TPStk As Byte = 29
Private Const TBRate As Byte = 30
Private Const TPRate As Byte = 31
Private Const Bin As Byte = 32
Private Const LastRate As Byte = 33
Private Const HPRate As Byte = 34
Private Const LPRate As Byte = 35
Private Const God As Byte = 36
Private Const PONOCode As Byte = 37
Private Const POSrlNo As Byte = 38
Private Const ChlDocId As Byte = 39
Private Const ChlDocSr As Byte = 40
Private Const PartGrade As Byte = 41
Private Const EffectDate As Byte = 42
Private Const MRPRate As Byte = 43

Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem
Dim rsTaxPer As ADODB.Recordset
Dim mSatYn As Boolean

Private Sub cmdPost_Click()
Dim I As Integer, mStartdate As String, mEndDate As String
Dim DupMaster As ADODB.Recordset
If Master.RecordCount > 0 Then
    Set DupMaster = Master.Clone
    
    If DupMaster.RecordCount > 0 Then DupMaster.MoveFirst
    mStartdate = InputBox("Posting Required from which Date ?", "Start Date for Posting", PubLoginDate)
    mEndDate = InputBox("Posting Required upto which Date ?", "Last Date for Posting", PubLoginDate)
    
    If mStartdate = "" Or mEndDate = "" Then Exit Sub
    mStartdate = MakeDate(mStartdate)
    mEndDate = MakeDate(mEndDate)
    
    
    mReposting = True
    Do Until DupMaster.EOF
'        If Trim(mID(DupMaster!SearchCode, 3, 5)) <> TrfVType Then GoTo MyNextRecord
        If IsNull(DupMaster!V_DATE) Then GoTo MyNextRecord
        If DupMaster!V_DATE < CDate(mStartdate) Then GoTo MyNextRecord
        If DupMaster!V_DATE > CDate(mEndDate) Then GoTo MyNextRecord
        
        Me.SEARCHBACK (DupMaster!SearchCode)
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        FGrid.Refresh
        For I = 0 To txt.Count - 1
            txt(I).Refresh
        Next
        
        Call TopCtrl1_eEdit
        Call TopCtrl1_eSave
MyNextRecord:
        DupMaster.MoveNext
    Loop
    Master.MoveFirst
    mReposting = False
    MsgBox "Updation Complete", vbInformation, "Re-Updation"
    Set DupMaster = Nothing
End If
End Sub

Private Sub CmdTransPost_Click()
Master.MoveFirst
    Do Until Master.EOF
        Call MoveRec
        Disp_Text SETS("EDIT", Me, Master)
        'txt(Vdate).SetFocus
        FGrid.AddItem FGrid.Rows
        
        TopCtrl1_eSave
        'Prn.Visible = False
        Me.Refresh
MyNextRecord:
        Master.MoveNext
    Loop
End Sub

Private Sub DGOrdPart_Click()
FGrid.TextMatrix(FGrid.Row, POSrlNo) = GRs!Srl_No
Set GRs = Nothing
FGrid.SetFocus
DGOrdPart.Visible = False
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
    If FrmDetail.Visible = True Then FrmDetail.Visible = False
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
If rsCtrlAc!SprCash_Ac = "" Then
    MsgStr = "Please Fill Spare Purchase "
    UnLoadFrm = True
End If
'EOF Spare A/c control checking
If UnLoadFrm Then
    MsgBox "Spare Purchase Entry Loading Aborted !" & vbCrLf & MsgStr & " A/c Controls through Utility Menu", vbInformation, "Validation"
    Unload Me
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
    Call Ini_Pub
    Label3(4) = PubForm31Caption
    Label3(24) = PubForm31Caption & " No."
    mVType = PurCashVType
    txt(VDate).Tag = PubLoginDate
    
    'A/c Pstong Control Checking
    Set rsCtrlAc = New ADODB.Recordset
    rsCtrlAc.CursorLocation = adUseClient
    'CSSprAc=Temp Sale A/c
    rsCtrlAc.Open "Select SprPurTrans_Ac,EntryTax_Ac,SprCash_Ac From AcControls", GCnFaS, adOpenDynamic, adLockOptimistic
    'eof checking
    
        Dim sitecond As String
        sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("Docid", "3", "1") & "='" & PubSiteCode & "'"
    End If

    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "select DocID as SearchCode,DocID,U_EntDt, V_Date  from Sp_Purch  " & _
            "where left(DocID,1)='" & PubDivCode & "' " & sitecond & " and v_type in ('" & PurCashVType & "','" & PurCrVType & "') Order By V_Date Desc, DocID desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 DocID as SearchCode,DocID,U_EntDt, V_Date  from Sp_Purch  " & _
            "where left(DocID,1)='" & PubDivCode & "' " & sitecond & " and v_type in ('" & PurCashVType & "','" & PurCrVType & "') Order By V_Date Desc, DocID desc", GCn, adOpenDynamic, adLockOptimistic
    End If
    
    Set DGPart.DataSource = RsPart
    
    Set rsForm = New ADODB.Recordset
    With rsForm
        .CursorLocation = adUseClient
        .Open "SELECT TaxForms.Form_Code as code,TaxForms.form_Desc as name FROM TaxForms where Trn_Type='Purchase' and Spare_YN = 1 Order by TaxForms.form_Desc ", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGForm.DataSource = rsForm
        
    Set rsForm31 = New ADODB.Recordset
    With rsForm31
        .CursorLocation = adUseClient
        .Open "SELECT TaxForms.Form_Code as code ,TaxForms.form_Desc as name  FROM TaxForms where Spare_YN = 1 and trn_Type = 'Permit' order by  TaxForms.form_Desc", GCn, adOpenDynamic, adLockOptimistic
    End With

'    Set rsPONo = New ADODB.Recordset
'    With rsPONo
'        .CursorLocation = adUseClient
'        .Open "Select OrderID as Code,Order_Reg_No as Name,Order_Reg_Dt, Right(OrderID,13) as OurDocNo,V_Date From SP_Order Where left(Order_Type,4)='S_PO' Order By OrderID", GCn, adOpenDynamic, adLockOptimistic
'    End With
'    Set DGPONo.DataSource = rsPONo

    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
'    RsParty.Open "select SubGroup.SubCode as code,SubGroup.NAME,Party_Type from SubGroup  Where firmCode = '" & PubFirmCode & "' and Nature='Supplier'  order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    If GCn.Execute("Select " & vIsNull("DebtorInSupplierHelp", "0") & " From Syctrl").Fields(0) = 1 Then
        GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,SubGroup.Add1,SubGroup.Add2,SubGroup.Add3,City.CityName as City,Party_Type from ((SubGroup " & _
            "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode) Left Join City on SubGroup.CityCode=City.CityCode )" & _
            "Where left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
            " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
    Else
        GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,SubGroup.Add1,SubGroup.Add2,SubGroup.Add3,City.CityName as City,Party_Type from ((SubGroup " & _
            "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode) Left Join City on SubGroup.CityCode=City.CityCode )" & _
            "Where left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "') " & _
            " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
    End If
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set rsGod = New ADODB.Recordset
    rsGod.CursorLocation = adUseClient
    rsGod.Open "select god_code as code,god_name as name from godown where appli_for=0 order by god_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGGod.DataSource = rsGod
    
    Set rsTrans = New ADODB.Recordset
    rsTrans.CursorLocation = adUseClient
    rsTrans.Open "select distinct transport as name from  sp_Purch  where  transport <>   '' order by transport", GCn, adOpenDynamic, adLockOptimistic
    Set DGTrans.DataSource = rsTrans

    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsParty = Nothing
Set rsPONo = Nothing
Set rsGod = Nothing
Set rsForm = Nothing
Set rsTrans = Nothing
Set Master = Nothing
Set mListItem = Nothing
End Sub

Private Sub ListView_Click()
txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
txt(Val(ListView.Tag)).SetFocus
FrmList.Visible = False
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Double
    
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    If PubSatYn = 1 Then mSatYn = True
    DispText_Vat
    LblVPrefix.CAPTION = ""
    DocID = ""
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
    txt(TxtDocID).Enabled = False
    mPartyType = 0
    txt(VDate) = txt(VDate).Tag
    txt(VDate).SetFocus
    FGrid.Col = PONo
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim I As Double, mTrans As Boolean
Dim LedgAry(1) As LedgRec, mResult As Byte, MsgStr$, mTitle$


If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub

ApplyConsolidatedPosting CDate(txt(VDate))

If GCn.Execute("Select CancelYN from SP_Purch where DocID='" & Master!SearchCode & "'").Fields(0).Value = 1 Then
    MsgStr = "Are You Sure To Delete This ? "
    mTitle = "Delete Entry!"
Else
    MsgStr = "Are You Sure To Cancel This ? "
    mTitle = "Cancel Entry!"
End If
If MsgBox(MsgStr, vbYesNo + vbCritical + vbDefaultButton2, mTitle) = vbYes Then
    GCn.BeginTrans
    GCnFaS.BeginTrans
    mTrans = True
    '********
    For I = 1 To FGrid.Rows - 1
        GCn.Execute ("update sp_stock set  Rate2= 0,MRP_Rate2=0,Disc_Per2=0,Disc_Amt2=0," & _
            "Amount2=0,Ord_DiscPer2=0,Ord_DiscAmt2=0,Net_Amt2=0, " & _
            "v_date2=" & ConvertDate("") & ",Invoice_DocId = ''  " & _
            "where Invoice_Docid='" & Master!DocID & "'")
        GCn.Execute ("update sp_purch set Invoice_DocId='' " & _
            " where Invoice_Docid='" & Master!DocID & "'")
    Next
    
    If mTitle = "Delete Entry!" Then
        CreateLog Me, Master!SearchCode, mReposting
        GCn.Execute ("delete from Sp_Purch where docId = '" & Master!DocID & "'")
    Else
        GCn.Execute ("update sp_purch  set " & _
            " CancelYN=1,Tot_Amt=0,Tot_Disc_Amt= 0,Tot_Ord_DiscAmt=0,SprAmt=0,OilAmt=0,Tot_Goods_Value=0," & _
            " Tax_Amt=0,Addition =0,Deduction=0,NET_AMT =0,U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E', " & _
            " EntryTaxamt=0, Transportation=0 where DocId = '" & txt(TxtDocID) & "'")
    End If
    '*********
    'Unpost Ledger a/c
    If txt(VType).TEXT = "Cash" And IsConsolidatedPosting Then
        'A/c Posting
        ProcAcPost rsCtrlAc
        'EOF Posting
    Else
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, txt(TxtDocID))
        If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
    End If
    'Unposting of Ledger completed
    GCnFaS.CommitTrans
    GCn.CommitTrans
    mTrans = False
    Master.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
Exit Sub
eloop1:
    If mTrans Then GCnFaS.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
    
    Disp_Text SETS("EDIT", Me, Master)
    txt(OPtSel).TEXT = ""
    FGrid.AddItem FGrid.Rows
    txt(Party).SetFocus
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
Dim I As Double
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
On Error GoTo ELoop
Dim RstRep As ADODB.Recordset
Dim mQry As String, I As Integer, X11
'Sp_Purch->Spp.DocID,Spp.DocIDHelp   ,Spp.V_Type  ,Spp.V_No    ,Spp.Site_Code   ,Spp.V_Date  ,Spp.Cash_Credit ,Spp.Party_Code  ,Spp.Party_Name  ,Spp.L_C ,Spp.Form_Code   ,Spp.FormNo  ,Spp.FormIssRecDate  ,Spp.Party_Doc_No    ,Spp.Party_Doc_Date  ,Spp.RoadPermit_FormCode ,Spp.RoadPermit_No   ,Spp.GR_RR_No    ,Spp.GR_RR_Date  ,Spp.Tot_No_of_Items ,Spp.Tot_Doc_Qty ,Spp.Tot_Phy_Qty ,Spp.SprAmt_MRP_TB   ,Spp.SprAmt_MRP_TP   ,Spp.OilAmt_MRP_TB   ,Spp.OilAmt_MRP_TP   ,Spp.SprAmt_TB   ,Spp.SprAmt_TP   ,Spp.OilAmt_TB   ,Spp.OilAmt_TP   ,Spp.OilAmt  ,Spp.SprAmt  ,Spp.Tot_Amt ,Spp.Tot_Disc_Amt    ,Spp.Tot_Ord_DiscAmt ,Spp.Tot_Goods_Value ,Spp.Tax_Amt ,Spp.Addition    ,Spp.Deduction   ,Spp.NET_AMT ,Spp.EntryTaxPer ,Spp.EntryTaxAmt ,Spp.Case_No ,Spp.Case_Mark   ,Spp.Transport   ,Spp.Supply_Mode ,Spp.Remarks ,Spp.Invoice_DocId   ,Spp.AcPsoting_YN    ,Spp.DrAc_Code   ,Spp.Printed_YN  ,Spp.CancelYN    ,Spp.CancelRemark    ,Spp.U_Name  ,Spp.U_EntDt ,Spp.U_AE    ,Spp.Trf_Date    ,Spp.Transportation  ,Spp.SiebelDocID
'Sp_Stock->DocID   Srl_No  V_Type  V_No    V_Date  Party_Code  ,Sps.L_C ,Sps.Job_DocID   ,Sps.Job_DivCode ,Sps.Mech_Code   ,Sps.Order_DocId ,Sps.Order_Srl_No    ,Sps.Part_No ,Sps.Lub_Category    ,Sps.Godown  ,Sps.Qty_Doc ,Sps.Qty_Rec ,Sps.Qty_Iss ,Sps.Qty_Ret ,Sps.Tax_YN  ,Sps.MRP_YN  ,Sps.Rate    ,Sps.MRP_Rate    ,Sps.Disc_Per    ,Sps.Disc_Amt    ,Sps.Amount  ,Sps.Ord_DiscPer ,Sps.Ord_DiscAmt ,Sps.Net_Amt ,Sps.Purpose ,Sps.V_Rate  ,Sps.Part_SrlNo  ,Sps.Remark  ,Sps.Printed ,Sps.Invoice_DocId   ,Sps.V_Date2 ,Sps.Rate2   ,Sps.MRP_Rate2   ,Sps.Disc_Per2   ,Sps.Disc_Amt2   ,Sps.Amount2 ,Sps.Ord_DiscPer2    ,Sps.Ord_DiscAmt2    ,Sps.Net_Amt2    ,Sps.Printed2    ,Sps.TrnComplete_YN  ,Sps.Site_Code   ,Sps.U_Name  ,Sps.U_EntDt ,Sps.U_AE    Trf_Date    ,Sps.Claim_Div   ,Sps.Claim_Site  ,Sps.Claim_YearPrefix    ,Sps.Claim_Type  ,Sps.Claim_No    ,Sps.Claim_Date  ,Sps.ClaimId ,Sps.OldCodeMR   ,Sps.TaxPer  ,Sps.TaxAmt  ,Sps.PurDocNo    ,Sps.PurDocDate  ,Sps.SiebelDocID
    mQry = "Select Spp.DocID,Spp.DocIDHelp,Spp.V_Type,Spp.V_No ,Spp.Site_Code,Spp.V_Date,Spp.Cash_Credit ,Spp.Party_Code,Spp.Party_Name As Party_Name_SpPur, " & _
            " Spp.L_C ,Spp.Form_Code,Spp.FormNo,Spp.FormIssRecDate,Spp.Party_Doc_No ,Spp.Party_Doc_Date,Spp.RoadPermit_FormCode , " & _
            " Spp.RoadPermit_No,Spp.GR_RR_No ,Spp.GR_RR_Date,Spp.Tot_No_of_Items ,Spp.Tot_Doc_Qty ,Spp.Tot_Phy_Qty ,Spp.SprAmt_MRP_TB, " & _
            " Spp.SprAmt_MRP_TP,Spp.OilAmt_MRP_TB,Spp.OilAmt_MRP_TP,Spp.SprAmt_TB,Spp.SprAmt_TP,Spp.OilAmt_TB,Spp.OilAmt_TP,Spp.OilAmt, " & _
            " Spp.SprAmt,Spp.Tot_Amt ,Spp.Tot_Disc_Amt ,Spp.Tot_Ord_DiscAmt ,Spp.Tot_Goods_Value ,Spp.Tax_Amt ,Spp.Addition ,Spp.Deduction, " & _
            " Spp.NET_AMT ,Spp.EntryTaxPer ,Spp.EntryTaxAmt ,Spp.Case_No ,Spp.Case_Mark,Spp.Transport,Spp.Supply_Mode ,Spp.Remarks , " & _
            " Spp.Invoice_DocId,Spp.AcPsoting_YN ,Spp.DrAc_Code,Spp.Printed_YN,Spp.CancelYN ,Spp.CancelRemark ,Spp.U_Name,Spp.U_EntDt ,Spp.U_AE , " & _
            " Spp.Trf_Date ,Spp.Transportation,Spp.SiebelDocID, " & _
            " Sps.Job_DocID,Sps.Job_DivCode ,Sps.Mech_Code,Sps.Order_DocId,Sps.Order_Srl_No,Sps.Part_No ,Sps.Lub_Category,Sps.Godown,Sps.Qty_Doc, " & _
            " Sps.Qty_Rec,Sps.Qty_Iss,Sps.Qty_Ret,Sps.Tax_YN,Sps.MRP_YN,Sps.Rate,Sps.MRP_Rate,Sps.Disc_Per,Sps.Disc_Amt,Sps.Amount As AmountLine,Sps.Ord_DiscPer, " & _
            " Sps.Ord_DiscAmt AS Ord_DiscAmtLine,Sps.Net_Amt As Net_AmtLine,Sps.Purpose,Sps.V_Rate,Sps.Part_SrlNo,Sps.Remark As RemarkLine,Sps.Printed,Sps.V_Date2,Sps.Rate2, " & _
            " Sps.MRP_Rate2,Sps.Disc_Per2,Sps.Disc_Amt2,Sps.Amount2,Sps.Ord_DiscPer2,Sps.Ord_DiscAmt2,Sps.Net_Amt2,Sps.Printed2,Sps.TrnComplete_YN, " & _
            " Sps.Claim_Site,Sps.Claim_YearPrefix,Sps.Claim_Type,Sps.Claim_No,Sps.Claim_Date,Sps.ClaimId, " & _
            " Sps.TaxPer,Sps.TaxAmt,Sps.PurDocNo,Sps.PurDocDate,Sg.Name as PartyName,Sg.Add1,Sg.Add2,Sg.Add3,Sg.LstNo,Sg.CstNo,City.CityName,Part.Part_Name,TF.Form_Desc,TF.Printing_Desc,Part.UNIT,Sps.SFCPer,Sps.SFCAMT  " & _
            " From ((((Sp_Purch As Spp Left Join Sp_Stock As Sps On Spp.DocId=Sps.Invoice_DocId)    " & _
            "                        Left Join SubGroup As Sg  On Sg.SubCode   = Spp.Party_Code)    " & _
            "                        Left Join City            On Sg.CityCode  = City.CityCode)     " & _
            "                        Left Join Part            On Part.PART_NO = Sps.Part_No)       " & _
            "                        Left Join TaxForms As TF  On Tf.Form_Code = Spp.Form_Code      " & _
            " Where Spp.DocId='" & Master!SearchCode & "' "
        
    Set RstRep = GCn.Execute(mQry)
    
    If RstRep.RecordCount <= 0 Then MsgBox "** No Recors Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    X11 = CreateFieldDefFile(RstRep, PubRepoPath + "\Sp_PurchaseBill.ttx", True)
    Set rpt = rdApp.OpenReport(PubRepoPath + "\Sp_PurchaseBill.RPT")
    rpt.Database.SetDataSource RstRep
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
            Case UCase("title")
                rpt.FormulaFields(I).TEXT = "'" & "Spare Purchase Bill" & "'"
        End Select
    Next
    rpt.ReadRecords
    'rpt.PrintOut
    
    Call Report_View(rpt, Me.CAPTION, 0, True)
    
    Set RstRep = Nothing
    Exit Sub
ELoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eRef()
    RsParty.Requery
    RsPart.Requery
675567665    rsForm31.Requery
    rsForm.Requery
    rsTrans.Requery
'    rsPONo.Requery
    rsGod.Requery
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Double, SQLPBill$, PurAcCode$
Dim Rst As ADODB.Recordset, mTrans As Boolean, mGridFilled As Boolean
Dim DocIdHlp$, ChalVPrefix$, ChalVNo$, ChalDocID$, ChalDocIDhlp$, VoucherEditFlag2 As Boolean
Dim mItemVal As Double, mItemQty As Double, mTotDiffAmt As Double
Dim mDiffPerc As Single, mDiffAmt As Double
Dim mDiffPosted As Double, LastI As Integer, TotSpr As Double, TotOil As Double
'modishekhar
Dim OilAmtMrpTP As Double, OilAmtMrpTB As Double, OilAmtTP As Double, OilAmtTB As Double
Dim SprAmtMrpTP As Double, SprAmtMrpTB As Double, SprAmtTP As Double, SprAmtTB As Double

On Error GoTo errlbl
    
    If IsEditable(RetDate(txt(VDate))) = False Then Exit Sub
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

    If IsValid(txt(VDate), "Bill Date") = False Then Exit Sub
    If TopCtrl1.TopText2.CAPTION <> "Edit" Then If IsValid(txt(OPtSel), "Select/Create") = False Then Exit Sub
    If IsValid(txt(VType), "Cash/Credit") = False Then Exit Sub
    If IsValid(txt(SerialNo), "Bill Number") = False Then Exit Sub
    If IsValid(txt(LC), "Purchase Type") = False Then Exit Sub
    If IsValid(txt(Party), "Supplier Name") = False Then Exit Sub
    If IsValid(txt(FormType), "Form Type") = False Then Exit Sub
    If txt(SuppChlDate) <> "" Then
        If CDate(txt(SuppChlDate)) > CDate(txt(VDate)) Then
            MsgBox "Supplier Document Date  > Bill Date", vbOKOnly, "Validation": txt(SuppChlDate).SetFocus: Exit Sub
        End If
    End If
    
    
    
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, PNo) <> "" Then
            If FGrid.TextMatrix(I, Taxable) = "" Then MsgBox "Fill Taxable Yes/No in S.No. " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = Taxable: FGrid.SetFocus: Exit Sub
            If FGrid.TextMatrix(I, MRP) = "" Then MsgBox "Fill MRP Yes/No in S.No. " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = MRP: FGrid.SetFocus: Exit Sub
            If Val(FGrid.TextMatrix(I, PQty)) = 0 Then MsgBox "Fill Quantity in S.No. " & I, vbInformation, "Required data": FGrid.Row = I: FGrid.Col = PQty: FGrid.SetFocus: Exit Sub
            'Check removed by Nra for entering FOC Entry
'            If Val(FGrid.TextMatrix(I, FRate)) = 0 Then
''                If PubULabel <> "Y" Then
'                    MsgBox "Please Specify Rate in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = FRate: FGrid.SetFocus: Exit Sub
''                End If
'            End If
            If Val(FGrid.TextMatrix(I, DisRs)) + Val(FGrid.TextMatrix(I, DisOrd)) > Val(FGrid.TextMatrix(I, Amt)) Then
                MsgBox "Discount is greater than Item Value in Row No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = FRate: FGrid.SetFocus: Exit Sub
            End If
            If FGrid.TextMatrix(I, God) = "" Then MsgBox "Fill Godown in S.No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Godown: FGrid.SetFocus: Exit Sub
            mGridFilled = True
        End If
        'Total's
        If FGrid.TextMatrix(I, PartGrade) = PubPartGrade_Lub Then
            TotOil = TotOil + Val(FGrid.TextMatrix(I, ItemVal))
        Else
            TotSpr = TotSpr + Val(FGrid.TextMatrix(I, ItemVal))
        End If
    Next
    If mGridFilled = False Then MsgBox "Please Fill Item Detail", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = PNo: FGrid.SetFocus: Exit Sub
    If Val(txt(EntryTaxAmt)) <> 0 Then
        If IsNull(rsCtrlAc!EntryTax_Ac) Or rsCtrlAc!EntryTax_Ac = "" Then
            MsgBox "Please Fill Entry Tax A/c Code in Spare System Controls", vbInformation, "Control A/c Not Defined": txt(EntryTaxAmt).SetFocus: Exit Sub
        End If
    End If
    If Val(txt(Transportation)) <> 0 Then
        If IsNull(rsCtrlAc!SprPurTrans_Ac) Or rsCtrlAc!SprPurTrans_Ac = "" Then
            MsgBox "Please Fill Spare Purchase Transportation A/c in " & vbCrLf & "Spare System Controls", vbOKOnly, "Control A/c Not Defined"
            txt(Transportation).SetFocus: Exit Sub
        End If
    End If
    'Calculating Landed Rate for Each Part
    mTotDiffAmt = Val(txt(TaxAmt)) + Val(txt(Addition)) + Val(txt(Deduction))
    If TotSpr <> 0 Then
        mDiffPerc = Round((TotSpr * 100) / (TotSpr + TotOil), 2)
        mDiffAmt = Round(mTotDiffAmt * mDiffPerc / 100, 2)
    End If
    TotSpr = TotSpr + mDiffAmt
    TotOil = TotOil + (mTotDiffAmt - mDiffAmt)
    mDiffPosted = 0
    mDiffAmt = 0
    If mTotDiffAmt <> 0 Then
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, PNo) <> "" Then
                mItemVal = Val(FGrid.TextMatrix(I, ItemVal))
                mItemQty = Val(FGrid.TextMatrix(I, PQty))
                mDiffPerc = Round((mItemVal * 100) / Val(txt(TotGoods)), 2)
                mDiffAmt = Round(mTotDiffAmt * mDiffPerc / 100, 2)
                mDiffPosted = mDiffPosted + mDiffAmt
                FGrid.TextMatrix(I, NDP) = Round((mItemVal + mDiffAmt) / mItemQty, 2)
                LastI = I
            End If
        Next
    End If
    If mTotDiffAmt - mDiffPosted <> 0 Then
        mItemVal = Val(FGrid.TextMatrix(LastI, ItemVal))
        mItemQty = Val(FGrid.TextMatrix(LastI, PQty))
        FGrid.TextMatrix(LastI, NDP) = Round((mItemVal + mTotDiffAmt - mDiffPosted) / mItemQty, 2)
    End If
    'EOF Landed Rate Calculation
    '
    'modishekhar
    'calculation for distinguish spr and oil amt
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, PNo) <> "" Then
            If FGrid.TextMatrix(I, PartGrade) = PubPartGrade_Lub Then
                If FGrid.TextMatrix(I, MRP) = "Yes" Then
                    If FGrid.TextMatrix(I, Taxable) = "Yes" Then
                        OilAmtMrpTB = Val(FGrid.TextMatrix(I, ItemVal)) + OilAmtMrpTB
                    Else
                        OilAmtMrpTP = Val(FGrid.TextMatrix(I, ItemVal)) + OilAmtMrpTP
                    End If
                Else
                    If FGrid.TextMatrix(I, Taxable) = "Yes" Then
                        OilAmtTB = Val(FGrid.TextMatrix(I, ItemVal)) + OilAmtTB
                    Else
                        OilAmtTP = Val(FGrid.TextMatrix(I, ItemVal)) + OilAmtTP
                    End If
                End If
            Else
                If FGrid.TextMatrix(I, MRP) = "Yes" Then
                    If FGrid.TextMatrix(I, Taxable) = "Yes" Then
                        SprAmtMrpTB = Val(FGrid.TextMatrix(I, ItemVal)) + SprAmtMrpTB
                    Else
                        SprAmtMrpTP = Val(FGrid.TextMatrix(I, ItemVal)) + SprAmtMrpTP
                    End If
                Else
                    If FGrid.TextMatrix(I, Taxable) = "Yes" Then
                        SprAmtTB = Val(FGrid.TextMatrix(I, ItemVal)) + SprAmtTB
                    Else
                        SprAmtTP = Val(FGrid.TextMatrix(I, ItemVal)) + SprAmtTP
                    End If
                End If
            End If
        End If
    Next
    RemoveTxtNull
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        DocID = txt(TxtDocID)
        If GCn.Execute("select count(*) from sp_purch where Left(DocID,1)='" & PubDivCode & "' And V_Type = '" & mVType & "' And V_No = " & Val(txt(SerialNo)) & " ").Fields(0) > 0 Then
            If VoucherEditFlag Then
                MsgBox "Purchase Serial No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                Exit Sub
            Else
                txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(txt(TxtDocID).Tag, Document_No)) Then
                    MsgBox "Purchase Serial No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    Exit Sub
                End If
            End If
        End If
    End If
    DocIdHlp = Replace(txt(TxtDocID), " ", "")
    
    GCn.BeginTrans
    GCnFaS.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION <> "Add" Then   'Edit Case
        CreateLog Me, Master!SearchCode, mReposting
        GCn.Execute ("update sp_purch  set Cash_Credit = '" & txt(VType) & "', Party_Code = '" & txt(Party).Tag & "', Party_Name= '" & txt(Party) & "', Party_Doc_No ='" & txt(SuppChlNo) & "', " & _
            "  Party_Doc_Date =" & ConvertDate(txt(SuppChlDate)) & ",L_C = '" & left(txt(LC), 1) & "',form_code = '" & txt(FormType).Tag & "',Case_no=" & Val(txt(CaseNo)) & ",Case_Mark='" & txt(CaseMark) & "',Tot_No_of_Items = " & Val(LblIVal.CAPTION) & _
            ", Tot_Doc_Qty = " & Val(LblDQty.CAPTION) & ",Tot_Phy_Qty = " & Val(LblPQty.CAPTION) & ",Tot_Amt = " & Val(txt(TOTAmt)) & ",Tot_Disc_Amt= " & Val(txt(TotDis)) & ",Tot_Ord_DiscAmt=" & Val(txt(TotOrdDis)) & " ," & _
            "  remarks  = '" & txt(Remark) & "', SprAmt=" & Val(txt(SprAmt)) & ", OilAmt=" & Val(txt(OilAmt)) & ", Tot_Goods_Value=" & Val(txt(TotGoods)) & ",Tax_Amt=" & Val(txt(TaxAmt)) & _
            ", Addition =" & Val(txt(Addition)) & ",Deduction=" & Val(txt(Deduction)) & ",NET_AMT = " & Val(txt(NetAmt)) & ",Supply_Mode = '" & txt(SupplyMode) & _
            "',EntryTaxPer = " & Val(txt(EntryTaxPer)) & ",SFCAMT=" & Val(txt(SFCAmt)) & ",EntryTaxamt = " & Val(txt(EntryTaxAmt)) & ", SatAmt = " & Val(txt(SatAmt)) & ",U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E', ModifyBy = '" & pubUName & "', ModifyDate = " & ConvertDateTime(PubServerDate) & ",GR_RR_No='" & txt(GrNo) & "',GR_RR_Date=" & ConvertDate(txt(GrDate)) & ",Transportation=" & Val(txt(Transportation)) & _
            " where DocId = '" & txt(TxtDocID) & "'")
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, PNo) <> "" And Val(FGrid.TextMatrix(I, PQty)) <> 0 Then
                GCn.Execute ("update sp_stock set Qty_Doc = " & Val(FGrid.TextMatrix(I, DQty)) & ", Rate2=" & Val(FGrid.TextMatrix(I, FRate)) & ", V_Rate=" & Val(FGrid.TextMatrix(I, NDP)) & _
                    ", Disc_Per2=" & Val(FGrid.TextMatrix(I, DisPer)) & ",Disc_Amt2=" & Val(FGrid.TextMatrix(I, DisRs)) & " , Amount2=" & Val(FGrid.TextMatrix(I, Amt)) & _
                    ", Ord_DiscPer2=" & Val(FGrid.TextMatrix(I, DisOrd)) & ", Ord_DiscAmt2=" & Val(FGrid.TextMatrix(I, DisOrdRs)) & ", Net_Amt2=" & Val(FGrid.TextMatrix(I, ItemVal)) & _
                    ", PurDocNo='" & txt(SuppChlNo) & "',PurDocDate=" & ConvertDate(txt(SuppChlDate)) & ", U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E' " & _
                    ",TaxPer=" & Val(FGrid.TextMatrix(I, TaxPer)) & ",TaxAmt=" & Val(FGrid.TextMatrix(I, TaxAmt1)) & ",SFCPer=" & Val(FGrid.TextMatrix(I, SFCPer)) & ",SFCAmt=" & Val(FGrid.TextMatrix(I, SFCAmt1)) & ", SatPer=" & Val(FGrid.TextMatrix(I, SatPer)) & ", SatAmt=" & Val(FGrid.TextMatrix(I, SatAmt1)) & " where invoice_docid = '" & Master!DocID & "' and srl_no = " & FGrid.TextMatrix(I, ChlDocSr) & "")
            End If
        Next
    Else    'Add
        SQLPBill = "insert into sp_purch(DocID,DocIDHelp,V_Type,V_No,Site_Code," _
            & "V_Date,Cash_Credit,Party_Code,Party_Name,Party_Doc_No," _
            & "Party_Doc_Date,L_C,form_code,Tot_No_of_Items," _
            & "Tot_Doc_Qty,Tot_Phy_Qty,Tot_Amt,Tot_Disc_Amt,Tot_Ord_DiscAmt," _
            & "SprAmt,OilAmt,Tot_Goods_Value,Tax_Amt, SatAmt,Addition,Deduction," _
            & "NET_AMT,EntryTaxPer,EntryTaxAmt,Remarks,Supply_Mode,U_Name,U_EntDt,U_AE, AddBy, AddDate,Transportation,Case_no,Case_Mark,GR_RR_No,GR_RR_Date,RoadPermit_Formcode,RoadPermit_no, Sat_Yn,SFCAMT) values(" _
            & "'" & txt(TxtDocID) & "','" & DocIdHlp & "','" & mVType & "'," & Val(txt(SerialNo)) & ",'" & PubSiteCode & PubSiteCode & _
            "'," & ConvertDate(txt(VDate)) & ",'" & txt(VType) & "','" & txt(Party).Tag & "','" & txt(Party) & "','" & txt(SuppChlNo) & _
            "'," & ConvertDate(txt(SuppChlDate)) & ",'" & left(txt(LC), 1) & "','" & txt(FormType).Tag & "'," & Val(LblIVal.CAPTION) & _
            ", " & Val(LblDQty.CAPTION) & "," & Val(LblPQty.CAPTION) & "," & Val(txt(TOTAmt)) & "," & Val(txt(TotDis)) & "," & Val(txt(TotOrdDis)) & _
            ", " & Val(txt(SprAmt)) & "," & Val(txt(OilAmt)) & ", " & Val(txt(TotGoods)) & "," & Val(txt(TaxAmt)) & ", " & Val(txt(SatAmt)) & "," & Val(txt(Addition)) & "," & Val(txt(Deduction)) & _
            ", " & Val(txt(NetAmt)) & "," & Val(txt(EntryTaxPer)) & "," & Val(txt(EntryTaxAmt)) & ",'" & txt(Remark) & "','" & txt(SupplyMode) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A', '" & pubUName & "', " & ConvertDateTime(PubServerDate) & "," & Val(txt(Transportation)) & "," & Val(txt(CaseNo)) & ",'" & txt(CaseMark) & "','" & txt(GrNo) & "'," & ConvertDate(txt(GrDate)) & ",'" & txt(PermitType).Tag & "','" & txt(FormNo) & "', " & IIf(mSatYn, 1, 0) & "," & Val(txt(SFCAmt)) & ")"
        If txt(OPtSel) = "Create" Then
            'Create Challan / Material Receipt
            '********* < Start Rahul At U.N.Automobiles 10-04-2003  > *********************
            ChalDocID = GetDocID(GCnFaS, ChalVType, txt(VDate), VoucherEditFlag2, txt(SerialNo2), LblVPrefix2)
            '************ < End  > *******************************************************
            ChalDocIDhlp = Replace(ChalDocID, " ", "")
            GSQL = "insert into sp_purch(DocID,DocIDHelp,V_Type,V_No,Site_Code," _
                & "V_Date,Cash_Credit,Party_Code,Party_Name,Party_Doc_No," _
                & "Party_Doc_Date,RoadPermit_Formcode,RoadPermit_no,GR_RR_No,GR_RR_Date," _
                & "L_C,form_code,Tot_No_of_Items,Tot_Doc_Qty,Tot_Phy_Qty," _
                & "Tot_Amt,Tot_Disc_Amt,Tot_Ord_DiscAmt,Tot_Goods_Value,Tax_Amt, SatAmt," _
                & "Addition,Deduction,NET_AMT,Case_no,Case_Mark," _
                & "Transport,Supply_Mode,Invoice_DocId , Sat_Yn,U_Name,U_EntDt,U_AE, AddBy, AddDate,SFCAMT) values(" _
                & "'" & ChalDocID & "','" & ChalDocIDhlp & "','" & ChalVType & "'," & Val(txt(SerialNo2)) & ",'" & PubSiteCode & PubSiteCode & "'," _
                & "" & ConvertDate(txt(VDate)) & ",'" & txt(VType) & "','" & txt(Party).Tag & "','" & txt(Party) & "','" & txt(SuppChlNo) & "'," _
                & "" & ConvertDate(txt(SuppChlDate)) & ",'" & txt(PermitType).Tag & "','" & txt(FormNo) & "','" & txt(GrNo) & "'," & ConvertDate(txt(GrDate)) & "," _
                & "'" & left(txt(LC), 1) & "','" & txt(FormType).Tag & "'," & Val(LblIVal.CAPTION) & "," & Val(LblDQty.CAPTION) & "," & Val(LblPQty.CAPTION) & "," _
                & "" & Val(txt(TOTAmt)) & "," & Val(txt(TotDis)) & "," & Val(txt(TotOrdDis)) & "," & Val(txt(TotGoods)) & "," & Val(txt(TaxAmt)) & "," & Val(txt(SatAmt)) & "," _
                & "" & Val(txt(Addition)) & "," & Val(txt(Deduction)) & "," & Val(txt(NetAmt)) & "," & Val(txt(CaseNo)) & ",'" & txt(CaseMark) & "'," _
                & "'" & txt(Transport) & "','" & txt(SupplyMode) & "','" & txt(TxtDocID) & "', " & IIf(mSatYn, 1, 0) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A', '" & pubUName & "', " & ConvertDateTime(PubServerDate) & "," & Val(txt(SFCAmt)) & ")"
            GCn.Execute (GSQL)
            
            'Purchase Bill Add
            GCn.Execute (SQLPBill)
            For I = 1 To FGrid.Rows - 1
                If FGrid.TextMatrix(I, PNo) <> "" And Val(FGrid.TextMatrix(I, PQty)) <> 0 Then
                    GCn.Execute ("insert into sp_stock(DocID,Srl_No,V_Type,V_No,V_Date,Party_Code,L_C,Order_DocId, " & _
                        " Part_No, Godown, Qty_Doc, Qty_Rec, Tax_YN, MRP_YN, Rate, V_Rate, " & _
                        " Disc_Per,Disc_Amt, Amount, Ord_DiscPer, Ord_DiscAmt, Net_Amt, " & _
                        " Part_SrlNo, Site_Code,PurDocNo,PurDocDate, U_Name, U_EntDt, U_AE, " & _
                        " v_date2,Invoice_DocId,Rate2, Amount2, Disc_Per2, Disc_Amt2, Ord_DiscPer2, Ord_DiscAmt2, Net_Amt2, TaxPer, TaxAmt, SatPer, SatAmt , SFCPer, SFCAmt ) " & _
                        " values('" & ChalDocID & "'," & I & ",'" & ChalVType & "'," & Val(txt(SerialNo2)) & "," & ConvertDate(txt(VDate)) & ",'" & txt(Party).Tag & "','" & left(txt(LC), 1) & "','" & FGrid.TextMatrix(I, PONOCode) & "', " & _
                        " '" & FGrid.TextMatrix(I, PNo) & "','" & FGrid.TextMatrix(I, God) & "'," & Val(FGrid.TextMatrix(I, DQty)) & ", " & Val(FGrid.TextMatrix(I, PQty)) & "," & IIf(FGrid.TextMatrix(I, Taxable) = "Yes", 1, 0) & ", " & IIf(FGrid.TextMatrix(I, MRP) = "Yes", 1, 0) & "," & Val(FGrid.TextMatrix(I, FRate)) & " ," & Val(FGrid.TextMatrix(I, NDP)) & " , " & _
                        " " & Val(FGrid.TextMatrix(I, DisPer)) & "," & Val(FGrid.TextMatrix(I, DisRs)) & "," & Val(FGrid.TextMatrix(I, Amt)) & ", " & Val(FGrid.TextMatrix(I, DisOrd)) & "," & Val(FGrid.TextMatrix(I, DisOrdRs)) & "," & Val(FGrid.TextMatrix(I, ItemVal)) & _
                        ",'" & FGrid.TextMatrix(I, PartSrlNo) & "','" & PubSiteCode & PubSiteCode & "','" & txt(SuppChlNo) & "'," & ConvertDate(txt(SuppChlDate)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & _
                        ",'A'," & ConvertDate(txt(VDate)) & ",'" & txt(TxtDocID) & "', " & Val(FGrid.TextMatrix(I, FRate)) & ", " & Val(FGrid.TextMatrix(I, Amt)) & "," & Val(FGrid.TextMatrix(I, DisPer)) & "," & Val(FGrid.TextMatrix(I, DisRs)) & ", " & Val(FGrid.TextMatrix(I, DisOrd)) & "," & Val(FGrid.TextMatrix(I, DisOrdRs)) & "," & Val(FGrid.TextMatrix(I, ItemVal)) & "," & Val(FGrid.TextMatrix(I, TaxPer)) & " ," & Val(FGrid.TextMatrix(I, TaxAmt1)) & "," & Val(FGrid.TextMatrix(I, SatPer)) & " ," & Val(FGrid.TextMatrix(I, SatAmt1)) & "," & Val(FGrid.TextMatrix(I, SFCPer)) & " ," & Val(FGrid.TextMatrix(I, SFCAmt1)) & ")")
                    If FGrid.TextMatrix(I, PONo) <> "" And FGrid.TextMatrix(I, POSrlNo) <> "" Then
                        GCn.Execute "Update SP_Order1 Set Sup_Qty=Sup_Qty+" & Val(FGrid.TextMatrix(I, PQty)) & " Where OrderId='" & FGrid.TextMatrix(I, PONOCode) & "' and Srl_No=" & FGrid.TextMatrix(I, POSrlNo) & ""
                    End If
                    Call UpdStkGridToTable(FGrid.TextMatrix(I, PNo), "+", FGrid.TextMatrix(I, MRP), FGrid.TextMatrix(I, Taxable), FGrid.TextMatrix(I, PQty))
                End If
            Next
        Else    'If Txt(OPtSel) = "Select" Then
            GCn.Execute (SQLPBill)
            For I = 1 To FGridSel.Rows - 1
                If FGridSel.TextMatrix(I, 0) <> "" And FGridSel.TextMatrix(I, 6) <> "" Then
                    GCn.Execute ("update sp_purch set Invoice_DocId = '" & txt(TxtDocID) & "' where DocId = '" & FGridSel.TextMatrix(I, 6) & "'")
                End If
            Next
            For I = 1 To FGrid.Rows - 1
                If FGrid.TextMatrix(I, PNo) <> "" And Val(FGrid.TextMatrix(I, PQty)) <> 0 Then
                    GCn.Execute ("update sp_stock set Rate2=" & Val(FGrid.TextMatrix(I, FRate)) & ", V_Rate=" & Val(FGrid.TextMatrix(I, NDP)) & _
                        ", Disc_Per2=" & Val(FGrid.TextMatrix(I, DisPer)) & ",Disc_Amt2=" & Val(FGrid.TextMatrix(I, DisRs)) & " , Amount2=" & Val(FGrid.TextMatrix(I, Amt)) & _
                        ", Ord_DiscPer2=" & Val(FGrid.TextMatrix(I, DisOrd)) & ", Ord_DiscAmt2=" & Val(FGrid.TextMatrix(I, DisOrdRs)) & ", Net_Amt2=" & Val(FGrid.TextMatrix(I, ItemVal)) & _
                        ", PurDocNo='" & txt(SuppChlNo) & "',PurDocDate=" & ConvertDate(txt(SuppChlDate)) & ", U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E',v_date2 = " & ConvertDate(txt(VDate)) & ",Invoice_DocId = '" & txt(TxtDocID) & _
                        "' where docid = '" & FGrid.TextMatrix(I, ChlDocId) & "' and Srl_No = " & Val(FGrid.TextMatrix(I, ChlDocSr)) & "")
                End If
            Next
        End If
    End If
    'Update Last purch rate
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, PNo) <> "" And Val(FGrid.TextMatrix(I, PQty)) <> 0 Then
            GCn.Execute ("Update Part Set PurDocId = '" & ChalDocID & "',PurDate = " & ConvertDate(txt(VDate)) & ",PurRate=" & Val(FGrid.TextMatrix(I, NDP)) & " where Part_No='" & FGrid.TextMatrix(I, PNo) & "' and Div_Code='" & PubDivCode & "'")
        End If
    Next
    '**********************
    
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If txt(OPtSel) <> "Select" Then
            'Update Srl No. for Created challan No.
            'update Table only when DocSrlNo >Table.SerialNo
            UpdVouSrlNo GCnFaS, ChalDocID, txt(VDate)
        End If
        'Update Srl No. for Purchase No.
        UpdVouSrlNo GCnFaS, txt(TxtDocID), txt(VDate)
    End If
    'A/c Posting
    '************
    If Val(txt(NetAmt)) > 0 Then
        ProcAcPost rsCtrlAc
    End If
    'EOF Posting
    GCnFaS.CommitTrans
    GCn.CommitTrans
    mTrans = False
    Set Rst = Nothing
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select DocID as SearchCode,DocID,U_EntDt, V_Date  from Sp_Purch  " & _
        "where left(DocID,1)='" & PubDivCode & "' and v_type in ('" & PurCashVType & "','" & PurCrVType & "') And DocId = '" & txt(TxtDocID) & "' Order By V_Date Desc, DocID desc")
    End If
    rsTrans.Requery
    Master.FIND "SearchCode = '" & txt(TxtDocID) & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(SerialNo)) > Val(DeCodeDocID(DocID, Document_No)) Then
            MsgBox "Purchase Serial No." & Trim(DeCodeDocID(DocID, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
        txt(VDate).Tag = txt(VDate)
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans Then GCnFaS.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    
        Dim sitecond As String
        sitecond = " And  V_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = sitecond & " and  " & cMID("sp_purch.Docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    GSQL = "SELECT sp_purch.DocId as searchcode, sp_purch.V_Date AS VoucherDate,sp_purch.DocId, sp_purch.v_Type, sp_purch.v_No,SP_Purch.Party_Doc_No as PDocNo, sp_purch.Site_Code, SubGroup.Name as PartyName FROM sp_purch LEFT JOIN SubGroup ON sp_purch.Party_Code = SubGroup.SubCode where v_type in ('" & PurCashVType & "','" & PurCrVType & "') " & sitecond & " Order By V_Date Desc"
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
        Set Master = GCn.Execute("select DocID as SearchCode,DocID,U_EntDt, V_Date  from Sp_Purch  " & _
            "where left(DocID,1)='" & PubDivCode & "' and v_type in ('" & PurCashVType & "','" & PurCrVType & "') And DocId = '" & MyValue & "' Order By V_Date Desc, DocID desc")
    
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub Txt_GotFocus(Index As Integer)
On Error GoTo ELoop
If txt(VType).TEXT = "" And Index <> VDate Then txt(VType).SetFocus
TxtGrid(0).Visible = False
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case VType
        ListArray = Array("Cash", "Credit")
        Set mListItem = ListView_Items(ListView, txt, VType, ListArray, 2)
    Case OPtSel
        ListArray = Array("Create", "Select")
        Set mListItem = ListView_Items(ListView, txt, OPtSel, ListArray, 2)
    Case Party
        Set DGParty.DataSource = RsParty
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case LC
        ListArray = Array("Local", "Central")
        Set mListItem = ListView_Items(ListView, txt, LC, ListArray, 2)
    Case PermitType
        Set DGForm.DataSource = rsForm31
        If rsForm31.RecordCount = 0 Or (rsForm31.EOF = True Or rsForm31.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsForm31!Name Then
            rsForm31.MoveFirst
            rsForm31.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case FormType
        Set DGForm.DataSource = rsForm
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsForm!Name Then
            rsForm.MoveFirst
            rsForm.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case Addition, Deduction, TaxAmt, EntryTaxPer, EntryTaxAmt, SatAmt
        SendKeys "{HOME}+{END}"
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case VType
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case LC
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case OPtSel
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case Party
        If txt(VType).TEXT = "Credit" Then
            DGridTxtKeyDown DGParty, txt, Party, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
        End If
    Case Transport
        DGridTxtKeyDown_Mast DGTrans, txt, Transport, rsTrans, KeyCode, False, 0
    Case FormType
        DGridTxtKeyDown DGForm, txt, FormType, rsForm, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
    Case PermitType
        DGridTxtKeyDown DGForm, txt, PermitType, rsForm31, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
End Select
If FrmList.Visible = False And DGParty.Visible = False And DGTrans.Visible = False And DGGod.Visible = False And DGForm.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VType Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> EntryTaxAmt Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = EntryTaxAmt Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> VDate Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> Party Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case PermitType
        If DGForm.Visible = True Then DGridTxtKeyPress txt, Index, rsForm31, KeyAscii, "Name"
    Case Party
        If txt(VType).TEXT = "Credit" Then
            If DGParty.Visible = True Then DGridTxtKeyPress txt, Index, RsParty, KeyAscii, "Name"
            lblGroup.Visible = True: lblGroup.BackColor = vbBlack: lblGroup.Locked = True: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
        End If
    Case FormType
        If DGForm.Visible = True Then DGridTxtKeyPress txt, Index, rsForm, KeyAscii, "Name"
    Case SerialNo
        Call NumPress(txt(Index), KeyAscii, 7, 0)
    Case CaseNo
        Call NumPress(txt(Index), KeyAscii, 6, 0)
    Case EntryTaxPer
        Call NumPress(txt(Index), KeyAscii, 2, 2)
    Case Addition, Deduction, TaxAmt, EntryTaxAmt, Transportation, SatAmt
        Call NumPress(txt(Index), KeyAscii, 8, 2)
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case VType
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case LC
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case OPtSel
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case Transport
        If DGTrans.Visible = True Then DGridTxtKeyUp_Mast txt, Transport, rsTrans, KeyCode, "Name"
    Case Addition, Deduction, TaxAmt, EntryTaxPer, EntryTaxAmt, Transportation, SatAmt
        Amt_Cal Index
End Select
'Amt_Cal
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim I As Double
Select Case Index
    Case EntryTaxPer, Addition, Deduction, TaxAmt, EntryTaxAmt, Transportation, TaxAmt, SatAmt
        txt(Index) = Format(txt(Index), "0.00")

    Case VType
        If IsValid(txt(VType), "Cash Credit") = False Then Cancel = True:   Exit Sub
        If txt(VType).TEXT <> "" Then txt(VType).TEXT = ListView.SelectedItem.TEXT
        If txt(VType).TEXT = "Cash" Then
            txt(Party).TEXT = "Cash"
            txt(Party).Tag = PubSprCashAc
            mVType = PurCashVType
        Else
            txt(Party).TEXT = ""
            txt(Party).Tag = ""
            mVType = PurCrVType
        End If
        txt(VType).Tag = txt(VType).TEXT
        txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
        DocID = txt(TxtDocID)
    Case LC
        If txt(LC).TEXT <> "" Then txt(LC).TEXT = ListView.SelectedItem.TEXT
        If IsValid(txt(LC), "Purchase Type") = False Then Cancel = True:   Exit Sub
    Case OPtSel
        If IsValid(txt(OPtSel), "Select/Create") = False Then Cancel = True:   Exit Sub
        If txt(OPtSel).TEXT = "Select" Then
            FrmSel.left = (Me.width - FrmSel.width) / 2
            FrmSel.Visible = True
            FrmSel.ZOrder 0
            FGridSel.Rows = 1
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "SELECT DocID, " & cTrim(cMID("docID", "8", "5")) & "+ " & cCStr(cTrim("Right(docID,8)")) & " as DocIdcode, V_Date, Party_Doc_No, SP_Purch.Party_Doc_Date, SP_Purch.Tot_Goods_Value FROM SP_Purch where left(DocID,1)='" & PubDivCode & "' and v_type = '" & ChalVType & "' and  Party_Code = '" & txt(Party).Tag & "' and  (Invoice_DocId Is Null or Invoice_DocId  = '') order by SP_Purch.Party_Doc_No, SP_Purch.Party_Doc_Date", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                Do Until Rst.EOF
                    FGridSel.AddItem "" & Chr(9) & Rst!DocIDcode & Chr(9) & Rst!V_DATE & Chr(9) & Rst!Party_Doc_No & Chr(9) & Rst!Party_Doc_Date & Chr(9) & Rst!Tot_Goods_Value & Chr(9) & Rst!DocID
                    Rst.MoveNext
                Loop
            End If
            FGridSel.AddItem ""
            FGridSel.FixedRows = 1
            FGridSel.Col = 0
            For I = 1 To FGridSel.Rows - 1
                FGridSel.Row = I
                FGridSel.CellFontName = "wingdings"
                FGridSel.CellFontSize = 16
            Next
            FldEnabled False
            FGridSel.SetFocus
        Else
            FldEnabled True
        End If
        txt(OPtSel).Tag = txt(OPtSel).TEXT
    
    Case SuppChlNo
        If GCn.Execute("Select * from SP_Purch where Party_Code='" & txt(Party).Tag & "' and Party_Doc_No='" & txt(SuppChlNo).TEXT & "' and Sp_Purch.V_type='" & mVType & "' ").RecordCount > 0 And TopCtrl1.TopText2 = "Add" Then
            MsgBox " This Supplier Document No for the Party Already Exists.", vbInformation + vbOKOnly, "Validation": txt(SuppChlNo).SetFocus
            Cancel = True
            Exit Sub
        End If
    Case Party
        If IsValid(txt(Index), Label3(3)) = False Then Cancel = True: Exit Sub
        'by lps 25-06-02
        If txt(VType).TEXT = "Cash" Then
            mPartyType = 0
            txt(Index).Tag = PubSprCashAc
            GSQL = "Select OrderID as Code,Order_Reg_No as Name,Order_Reg_Dt, " & cTrim(cMID("OrderID", "8", "5")) & "+ " & cCStr(cTrim("Right(OrderID,8)")) & " as OurDocNo,V_Date From SP_Order Where left(OrderID,1)='" & PubDivCode & "' and left(Order_Type,4)='S_PO' and V_Date<=" & ConvertDate(Format(txt(VDate), "dd-mmm-yyyy")) & " and OrdClosDate is null Order By OrderID"
        ElseIf txt(VType).TEXT = "Credit" Then
            mPartyType = VNull(RsParty!Party_Type)
            txt(Index).TEXT = RsParty!Name
            txt(Index).Tag = RsParty!Code
            GSQL = "Select OrderID as Code,Order_Reg_No as Name,Order_Reg_Dt, " & cTrim(cMID("OrderID", "8", "5")) & " + " & cCStr(cTrim("Right(OrderID,8)")) & " as OurDocNo,V_Date From SP_Order Where left(OrderID,1)='" & PubDivCode & "' and left(Order_Type,4)='S_PO' and Party_Code='" & txt(Party).Tag & "' and V_Date<=" & ConvertDate(Format(txt(VDate), "dd-mmm-yyyy")) & " and OrdClosDate is null Order By OrderID"
        End If
        If txt(OPtSel).TEXT <> "Select" Then
            Set rsPONo = New ADODB.Recordset
            rsPONo.CursorLocation = adUseClient
            rsPONo.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
            Set DGPONo.DataSource = rsPONo
        End If
    Case PermitType
        If rsForm31.RecordCount = 0 Or (rsForm31.EOF = True Or rsForm31.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = rsForm31!Name
            txt(Index).Tag = rsForm31!Code
        End If
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = rsForm!Name
            txt(Index).Tag = rsForm!Code
        End If
    Case SuppChlDate, GrDate
        txt(Index).TEXT = RetDate(txt(Index))
    Case VDate
        If Len(Trim(txt(VDate).TEXT)) = 0 Then
            txt(VDate).TEXT = PubLoginDate
        Else
            txt(Index).TEXT = RetDate(txt(Index))
        End If
        Cancel = Not CheckFinYear(txt(Index))
        If Cancel = False Then
            If txt(VType).TEXT = "" Then txt(VType).SetFocus: Exit Sub
            txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            DocID = txt(TxtDocID)
        End If
    Case SerialNo
        If IsValid(txt(SerialNo), "Serial No.") = False Then Cancel = True:   Exit Sub
        If VoucherEditFlag Then      ' Manual
            txt(TxtDocID) = GetDocID(GCnFaS, mVType, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            DocID = txt(TxtDocID)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select DocID From sp_purch Where docid='" & txt(TxtDocID) & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Cancel = True
                txt(SerialNo).SetFocus
            End If
        End If
End Select
Set Rst = Nothing
End Sub

Private Sub DGPart_Click()
If RsPart.RecordCount > 0 Then
    Select Case FGrid.Col
        Case PNo
            TxtGrid(0).TEXT = RsPart!Code
        Case PName
            TxtGrid(0) = RsPart!Name
        Case LName
            TxtGrid(0) = RsPart!LName
    End Select
End If
TxtGridValid_PNo
If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
DGPart.Visible = False
End Sub

Private Sub DGForm_Click()
    If rsForm.RecordCount > 0 Then
        txt(FormType).TEXT = rsForm!Name
        txt(FormType).Tag = rsForm!Code
    End If
    txt(FormType).SetFocus
    DGForm.Visible = False
End Sub
Private Sub DGTrans_Click()
    If rsTrans.RecordCount > 0 Then
        txt(Transport).TEXT = rsTrans!Name
    End If
    txt(Transport).SetFocus
    DGTrans.Visible = False
End Sub
Private Sub DGGod_Click()
    If rsGod.RecordCount > 0 Then
        TxtGrid(0).TEXT = rsGod!Name
         FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
         FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
    End If
   TxtGrid(0).SetFocus
    DGGod.Visible = False
End Sub

Private Sub DGPONo_Click()
    If rsPONo.RecordCount > 0 Then
            TxtGrid(0).TEXT = rsPONo!Name
            FGrid.TextMatrix(FGrid.Row, PONOCode) = rsPONo!Code
            FGrid.TextMatrix(FGrid.Row, PONo) = rsPONo!Name
    End If
   TxtGrid(0).SetFocus
    DGPONo.Visible = False
End Sub

Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        txt(Party).TEXT = RsParty!Name
        txt(Party).Tag = RsParty!Code
        mPartyType = RsParty!Party_Type
    End If
    txt(Party).SetFocus
    DGParty.Visible = False
    lblGroup.Visible = False
End Sub
Private Sub FGridSel_Click()
    If FGridSel.TextMatrix(FGridSel.Row, 0) = "" Then
        FGridSel.TextMatrix(FGridSel.Row, 0) = "ü"
    Else
        FGridSel.TextMatrix(FGridSel.Row, 0) = ""
    End If
End Sub

Private Sub FGrid_Click()
TxtGrid(0).Visible = False
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub

End Sub

Private Sub FGrid_DblClick()
    FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    TxtGrid(0).Visible = False
    If TopCtrl1.TopText2 <> "Browse" Then
        MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, PNo), _
            FGrid.TextMatrix(FGrid.Row, PName), FGrid.TextMatrix(FGrid.Row, LName), _
            MRPStkTB, MRPStkTP, TBStk, TPStk, _
            MRPRate, TBRate, TPRate, Bin, _
            LastRate, HPRate, LPRate, mCheckNegetiveStockSiteWise
        FrmDetail.Visible = True
    End If
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)

'Leave Cell-- > Enter Cell-- >KeyDown
If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If txt(VType).TEXT = "" Then txt(VType).SetFocus: Exit Sub
If TopCtrl1.TopText2.CAPTION <> "Edit" Then If txt(OPtSel).TEXT = "" Then txt(OPtSel).SetFocus: Exit Sub
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
    If TopCtrl1.TopText2.CAPTION = "Add" And txt(OPtSel).TEXT = "Create" Then
        Select Case FGrid.Col
            Case PONo, PQty, DQty, PartSrlNo
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        End Select
    End If
    Select Case FGrid.Col
        Case FRate, DisPer, DisRs, DisOrd, DisOrdRs
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    End Select
    Amt_Cal1
    Amt_Cal
End If

If KeyCode = vbKeyReturn Then
    If TopCtrl1.TopText2.CAPTION = "Add" And txt(OPtSel).TEXT = "Create" Then
        Select Case FGrid.Col
            Case PONo, PNo, PName, LName
                Call GridDblClick(Me, FGrid, TxtGrid, 0)
                TAddMode = False
            Case PartSrlNo, Taxable, Godown, MRP, PQty, DQty
                If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
                    Call GridDblClick(Me, FGrid, TxtGrid, 0)
                    TAddMode = False
                End If
        End Select
    End If
    Select Case FGrid.Col
        Case FRate, DisPer, DisRs, DisOrd, DisOrdRs, TaxPer, SatPer
            If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
                Call GridDblClick(Me, FGrid, TxtGrid, 0)
                TAddMode = False
            End If
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Add" And txt(OPtSel).TEXT = "Create" Then
    Select Case FGrid.Col
        Case PONo, PNo, PName, LName
            Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
        Case PartSrlNo, Godown, MRP, Taxable
            If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
            End If
        Case PQty
            If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
            End If
    End Select
End If
Select Case FGrid.Col
    Case Unit, Amt, ItemVal
        FGrid.Col = FGrid.Col + 1
        FGrid.SetFocus
    Case FRate, DisPer, DisOrd, DisRs, DisOrdRs, TaxPer, SatPer, DQty
        If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
           Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
        End If
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
'If Txt(OPtSel).TEXT = "Select" Or TopCtrl1.TopText2.CAPTION = "Edit" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
Dim I As Double
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
         End If
         For I = 1 To FGrid.Rows - 1
            FGrid.TextMatrix(I, 0) = I
         Next
        Amt_Cal
        MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, PNo), _
            FGrid.TextMatrix(FGrid.Row, PName), FGrid.TextMatrix(FGrid.Row, LName), _
            MRPStkTB, MRPStkTP, TBStk, TPStk, _
            MRPRate, TBRate, TPRate, Bin, _
            LastRate, HPRate, LPRate, mCheckNegetiveStockSiteWise
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
    FGrid.SetFocus
End If
Exit Sub
End Sub

Private Sub FGrid_RowColChange()
    If TopCtrl1.TopText2.CAPTION <> "Browse" Then
        MainLib.Fill_Frame Me.LblFrm, FGrid, FGrid.TextMatrix(FGrid.Row, PNo), _
           FGrid.TextMatrix(FGrid.Row, PName), FGrid.TextMatrix(FGrid.Row, LName), _
           MRPStkTB, MRPStkTP, TBStk, TPStk, _
           MRPRate, TBRate, TPRate, Bin, _
           LastRate, HPRate, LPRate, mCheckNegetiveStockSiteWise
    End If
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

Private Sub CmdSel_Click(Index As Integer)
Dim Rst As ADODB.Recordset
Dim I As Double
Select Case Index
    Case 1
        Call FillGridData
        Call Amt_Cal
    Case 2
        FrmSel.Visible = False
End Select
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
Next I
DocID = ""
End Sub

Private Sub MoveRec()
Dim Rs As Recordset, Master1 As ADODB.Recordset, I As Double
On Error GoTo error1
If Master.RecordCount > 0 Then
    If InStr(Me.TopCtrl1.Tag, "E") <> 0 Then Me.TopCtrl1.tEdit = True
    Set Master1 = New Recordset
    Master1.CursorLocation = adUseClient
    Master1.Open "select SubGroup.Name,SubGroup.Party_Type,SP_Purch.* from SP_Purch " _
        & " left join SubGroup on SP_Purch.Party_Code=SubGroup.SubCode " _
        & " where DocID='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
    If Master1!CancelYN = 1 Then
        TopCtrl1.tEdit = False
        LblCancel.Visible = True
    Else
        LblCancel.Visible = False
    End If
    DocID = Master!SearchCode
    txt(TxtDocID) = Master1!DocID
    LblDiv.CAPTION = "Division : " & left(Master1!DocID, 1)
    LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        txt(SFCAmt) = Format(VNull(Master1!SFCAmt), "0.00")
    LblUser = IIf(Not IsNull(Master1!AddDate), "Add By : " & XNull(Master1!AddBy) & "  Dated : " & XNull(Master1!AddDate), "") & IIf(Not IsNull(Master1!ModifyDate), "     Modify By : " & XNull(Master1!ModifyBy) & "  Dated : " & XNull(Master1!ModifyDate), "")
    LblVPrefix.CAPTION = mID(Master1!DocID, 8, 5)
    txt(SerialNo) = Master1!V_NO
    txt(VDate) = Master1!V_DATE
    txt(VType) = Master1!Cash_Credit
    txt(Party).Tag = Master1!Party_code
    If txt(VType) = "Cash" Then
        mVType = PurCashVType
        mPartyType = 0
        txt(Party) = Master1!Party_Name
    ElseIf txt(VType) = "Credit" Then
        mVType = PurCrVType
        mPartyType = VNull(Master1!Party_Type)
        txt(Party) = Master1!Name
    End If
    If PubBackEnd = "A" Then
        mSatYn = IIf(VNull(Master1!SAT_YN) = 1, True, False)
    Else
        mSatYn = IIf(VNull(Master1!SAT_YN) = True, True, False)
    End If
    DispText_Vat
    txt(SuppChlNo) = IIf(IsNull(Master1!Party_Doc_No), "", Master1!Party_Doc_No)
    txt(SuppChlDate) = IIf(IsNull(Master1!Party_Doc_Date), "", Master1!Party_Doc_Date)
    txt(LC) = IIf(Master1!L_C = "L", "Local", "Central")
    txt(FormType).Tag = IIf(IsNull(Master1!Form_Code), "", Master1!Form_Code)
    If txt(FormType).Tag <> "" Then
        txt(FormType) = GCn.Execute("select form_desc from taxforms where form_code = '" & txt(FormType).Tag & "'").Fields(0).Value
    Else
        txt(FormType) = ""
    End If
    txt(PermitType).Tag = IIf(IsNull(Master1!RoadPermit_FormCode), "", Master1!RoadPermit_FormCode)
    If txt(PermitType).Tag <> "" Then
        txt(PermitType) = GCn.Execute("select form_desc from taxforms where form_code = '" & txt(PermitType).Tag & "'").Fields(0).Value
    Else
        txt(PermitType) = ""
    End If
    txt(FormNo) = IIf(IsNull(Master1!RoadPermit_No), "", Master1!RoadPermit_No)
    txt(GrNo) = IIf(IsNull(Master1!GR_RR_No), "", Master1!GR_RR_No)
    txt(GrDate) = IIf(IsNull(Master1!GR_RR_Date), "", Master1!GR_RR_Date)
    txt(Remark) = IIf(IsNull(Master1!Remarks), "", Master1!Remarks)
    txt(CaseMark) = XNull(Master1!Case_Mark)
    txt(CaseNo) = XNull(Master1!Case_No)
    txt(Transport) = XNull(Master1!Transport)
    txt(SupplyMode) = XNull(Master1!Supply_Mode)
    
    LblIVal.CAPTION = Format(IIf(IsNull(Master1!Tot_No_of_Items), 0, Master1!Tot_No_of_Items), "0")
    LblDQty.CAPTION = Format(IIf(IsNull(Master1!Tot_Doc_Qty), 0, Master1!Tot_Doc_Qty), "0.000")
    LblPQty.CAPTION = Format(IIf(IsNull(Master1!Tot_Phy_Qty), 0, Master1!Tot_Phy_Qty), "0.000")
    LblAmt.CAPTION = Format(IIf(IsNull(Master1!Tot_Goods_Value), 0, Master1!Tot_Goods_Value), "0.00")
    txt(TOTAmt) = Format(IIf(IsNull(Master1!Tot_Amt), 0, Master1!Tot_Amt), "0.00")
    txt(TotDis) = Format(IIf(IsNull(Master1!Tot_Disc_Amt), 0, Master1!Tot_Disc_Amt), "0.00")
    txt(TotOrdDis) = Format(IIf(IsNull(Master1!Tot_Ord_DiscAmt), 0, Master1!Tot_Ord_DiscAmt), "0.00")
    txt(TotGoods) = Format(IIf(IsNull(Master1!Tot_Goods_Value), 0, Master1!Tot_Goods_Value), "0.00")
    txt(TaxAmt) = Format(IIf(IsNull(Master1!Tax_Amt), 0, Master1!Tax_Amt), "0.00")
    txt(SatAmt) = Format(VNull(Master1!SatAmt), "0.00")
    txt(Addition) = Format(IIf(IsNull(Master1!Addition), 0, Master1!Addition), "0.00")
    txt(Deduction) = Format(IIf(IsNull(Master1!Deduction), 0, Master1!Deduction), "0.00")
    txt(SprAmt) = Format(IIf(IsNull(Master1!SprAmt), 0, Master1!SprAmt), "0.00")
    txt(OilAmt) = Format(IIf(IsNull(Master1!OilAmt), 0, Master1!OilAmt), "0.00")
    txt(NetAmt) = Format(IIf(IsNull(Master1!Net_Amt), 0, Master1!Net_Amt), "0.00")
    txt(Transportation) = Format(IIf(IsNull(Master1!Transportation), 0, Master1!Transportation), "0.00")
    txt(EntryTaxPer) = Format(IIf(IsNull(Master1!EntryTaxPer), 0, Master1!EntryTaxPer), "0.00")
    txt(EntryTaxAmt) = Format(IIf(IsNull(Master1!EntryTaxAmt), 0, Master1!EntryTaxAmt), "0.00")
    txt(TotPurAmt) = Format(Val(txt(NetAmt)) + Val(txt(EntryTaxAmt) + Val(txt(Transportation))), "0.00")
    Set Rs = New Recordset
    Set Rs = GCn.Execute("SELECT SPO.Order_Reg_No,P.Part_Name, " & cTrim(cMID("Stk.Order_DocID", "8", "5")) & " + " & cCStr(cTrim("Right(Stk.Order_DocID,8)")) & " As OrderIDDisp, " & _
            " P.Local_Name,P.Part_Grade,P.UNIT, P.MRP, P.Cur_MRP_TBStk,P.Cur_MRP_TPStk,P.Cur_TB_Stk,P.Cur_TP_Stk, " & _
            " P.Cur_TB_Stk, P.Cur_TP_Stk, P.TP_SRate, P.TB_SRate, P.Bin_Loca, P.High_Pur_Rate, P.Low_Pur_Rate, Stk.*, G.God_Name" & _
            " FROM ((Sp_Stock Stk LEFT JOIN Part P ON Stk.Part_No = P.PART_NO and P.Div_Code = left(STK.DocID,1)) LEFT JOIN Godown G ON Stk.Godown = G.God_Code ) " & _
            " Left Join SP_Order SPO on Stk.Order_DocID=SPO.OrderID " & _
            " where Stk.Invoice_DocId = '" & Master1!DocID & "'")
    I = 1
    FGrid.Rows = 1
    If Rs.RecordCount > 0 Then
        Do Until Rs.EOF
            FGrid.AddItem ""
            With FGrid
                .TextMatrix(I, ChlDocId) = IIf(IsNull(Rs!DocID), "", Rs!DocID)
                .TextMatrix(I, ChlDocSr) = Rs!Srl_No
                .TextMatrix(I, 0) = I
'                .TextMatrix(i, PONo) = IIf(IsNull(rs!OrderIDDisp), "", rs!OrderIDDisp)
                .TextMatrix(I, PONo) = IIf(IsNull(Rs!Order_Reg_No), "", Rs!Order_Reg_No)
                .TextMatrix(I, PNo) = Rs!Part_No
                .TextMatrix(I, PONOCode) = XNull(Rs!Order_DocId)
                .TextMatrix(I, POSrlNo) = XNull(Rs!Order_Srl_No)
                .TextMatrix(I, Unit) = IIf(IsNull(Rs!Unit), "", Rs!Unit)
                .TextMatrix(I, MRP) = IIf(Rs!MRP_YN = 1, "Yes", "No")
                .TextMatrix(I, Taxable) = IIf(Rs!Tax_YN = 1, "Yes", "No")
                .TextMatrix(I, DQty) = IIf(Rs!Qty_Doc = 0, "", Format(Rs!Qty_Doc, "0.000"))
                .TextMatrix(I, PQty) = IIf(Rs!Qty_Rec = 0, "", Format(Rs!Qty_Rec, "0.000"))
                .TextMatrix(I, FRate) = IIf(Rs!Rate2 = 0, "", Format(Rs!Rate2, "0.0000"))
                .TextMatrix(I, Amt) = IIf(Rs!Amount2 = 0, "", Format(Rs!Amount2, "0.00"))
                .TextMatrix(I, DisPer) = IIf(Rs!Disc_Per2 = 0, "", Format(Rs!Disc_Per2, "0.00"))
                .TextMatrix(I, DisRs) = IIf(Rs!Disc_Amt2 = 0, "", Format(Rs!Disc_Amt2, "0.00"))
                .TextMatrix(I, DisOrd) = IIf(Rs!ord_Discper2 = 0, "", Format(Rs!ord_Discper2, "0.00"))
                .TextMatrix(I, DisOrdRs) = IIf(Rs!ord_Discamt2 = 0, "", Format(Rs!ord_Discamt2, "0.00"))
                
                .TextMatrix(I, SFCPer) = VNull(Rs!SFCPer)
                .TextMatrix(I, SFCAmt1) = Format(VNull(Rs!SFCAmt), "0.00")
                    
                If PubVATYN = 1 Then
                    .TextMatrix(I, TaxPer) = VNull(Rs!TaxPer)
                    .TextMatrix(I, TaxAmt1) = Format(VNull(Rs!TaxAmt), "0.00")
                    If mSatYn Then
                        .TextMatrix(I, SatPer) = VNull(Rs!SatPer)
                        .TextMatrix(I, SatAmt1) = Format(VNull(Rs!SatAmt), "0.00")
                    End If
                End If
                .TextMatrix(I, NDP) = IIf(Rs!V_Rate = 0, "", Format(Rs!V_Rate, "0.00"))
                .TextMatrix(I, ItemVal) = IIf(Rs!Net_Amt2 = 0, "", Format(Rs!Net_Amt2, "0.00"))
                .TextMatrix(I, God) = Rs!Godown
                .TextMatrix(I, Godown) = IIf(IsNull(Rs!God_Name), "", Rs!God_Name)
                .TextMatrix(I, PName) = IIf(IsNull(Rs!Part_Name), "", Rs!Part_Name)
                .TextMatrix(I, LName) = IIf(IsNull(Rs!Local_Name), "", Rs!Local_Name)
                .TextMatrix(I, MRPStkTB) = IIf(IsNull(Rs!Cur_MRP_TbStk), "", Rs!Cur_MRP_TbStk)
                .TextMatrix(I, MRPStkTP) = IIf(IsNull(Rs!Cur_MRP_TPStk), "", Rs!Cur_MRP_TPStk)
                .TextMatrix(I, MRPRate) = IIf(Rs!MRP = 0, "", Format(Rs!MRP, "0.00"))
                .TextMatrix(I, TBStk) = IIf(IsNull(Rs!Cur_TB_STk), "", Rs!Cur_TB_STk)
                .TextMatrix(I, TPStk) = IIf(IsNull(Rs!Cur_TP_Stk), "", Rs!Cur_TP_Stk)
                .TextMatrix(I, TBRate) = IIf(IsNull(Rs!TB_SRate), "", Rs!TB_SRate)
                .TextMatrix(I, TPRate) = IIf(IsNull(Rs!TP_SRate), "", Rs!TP_SRate)
                .TextMatrix(I, Bin) = IIf(IsNull(Rs!Bin_Loca), "", Rs!Bin_Loca)
'                    .TextMatrix(i, LastRate) = ""
                .TextMatrix(I, HPRate) = IIf(IsNull(Rs!high_pur_rate), "", Rs!high_pur_rate)
                .TextMatrix(I, LPRate) = IIf(IsNull(Rs!low_pur_rate), "", Rs!low_pur_rate)
                .TextMatrix(I, PartGrade) = IIf(IsNull(Rs!Part_Grade), "", Rs!Part_Grade)
            End With
            Rs.MoveNext
            I = I + 1
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    Set Rs = Nothing
Else
    Call BlankText
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End If
Set Master1 = Nothing
Set Rs = Nothing
Grid_Hide
'Amt_Cal 'EntTaxAmt
'Me.TopCtrl1.tPrn = False
Exit Sub
error1:
        CheckError
End Sub

Private Sub Ini_Grid()
Dim I As Byte
  ' |Part No.1|Part Name2|Unit 3|PO No 4|Taxable 5|MRP6|Qty(Doc)7|Qty(Phy)8|NDP 9 |Amount 10
'  |Dis %11|Ord Dis %12|Amount 13|Loal Name 14|Curr Stk Qty 15|MRP Qty 16 |Taxable Qty 17|TaxPaid Qty 18|Taxable Rate 19|TaxPaid Rate 20|Bin Location 21|Last Purch Rate 22|High Purch Rate 23|Low Purch Rate 24

'    FGrid.FormatString = "SrNo.|Part No.            |Part Name             |Unit |Godown          |PO No.         |Tax Y/N|MRP Y/N| Qty(Doc)|Qty(Phy)|Rate     |Amount    |Dis %    |Dis Rs   |Ord Dis %  |Ord Dis Rs  |NDP     |ItemValue   |Local Name|Curr Stk Qty|MRP Qty|Taxable Qty|TaxPaid Qty|Taxable Rate|TaxPaid Rate|Bin Location|Last Purch Rate|High Purch Rate|Low Purch Rate"
    'SrNo.1|Part No.2|Part Name3|Unit 4|Godown5|PO No.6|Tax Y/N 7|MRP Y/N8| Qty(Doc)9|Qty(Phy)10|Rate 11|Amount12|Dis %13|Dis Rs14|Ord Dis %15|Ord Dis Rs16|NDP 17|ItemValue 18|Local Name19|Curr Stk Qty20|MRP Qty21|Taxable Qty22|TaxPaid Qty23|Taxable Rate24|TaxPaid Rate25|Bin Location26|Last Purch Rate27|High Purch Rate28|Low Purch Rate29"
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
    With FGrid
        .RowHeightMin = PubGridRowHeight
        .Cols = 44
        .TextMatrix(0, 0) = "S.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 450

        .TextMatrix(0, PONo) = "Order No"
        .ColAlignment(PONo) = flexAlignLeftCenter
        .ColWidth(PONo) = 1395
        
        .TextMatrix(0, PNo) = "Part No"
        .ColAlignment(PNo) = flexAlignLeftCenter
        .ColWidth(PNo) = 1500
               
        .TextMatrix(0, Unit) = "Unit"
        .ColAlignment(Unit) = flexAlignLeftCenter
        .ColWidth(Unit) = 550
        
        .TextMatrix(0, MRP) = "MRP"
        .ColAlignment(MRP) = flexAlignLeftCenter
        .ColWidth(MRP) = 450

        .TextMatrix(0, Taxable) = "Tax"
        .ColAlignment(Taxable) = flexAlignLeftCenter
        .ColWidth(Taxable) = 420

        .TextMatrix(0, DQty) = "Qty(Doc)"
        .ColAlignmentFixed(DQty) = flexAlignRightCenter
        .ColWidth(DQty) = 960
        
        .TextMatrix(0, PQty) = "Qty(Phy)"
        .ColAlignmentFixed(PQty) = flexAlignRightCenter
        .ColWidth(PQty) = 960
        
        .TextMatrix(0, FRate) = "Rate"
        .ColAlignmentFixed(FRate) = flexAlignRightCenter
        .ColWidth(FRate) = 870

        .TextMatrix(0, Amt) = "Amount"
        .ColAlignmentFixed(Amt) = flexAlignRightCenter
        .ColWidth(Amt) = 1065
        
        .TextMatrix(0, DisPer) = "Disc%"
        .ColAlignmentFixed(DisPer) = flexAlignRightCenter
        .ColWidth(DisPer) = 555

        .TextMatrix(0, DisRs) = "Disc.Amt"
        .ColAlignmentFixed(DisRs) = flexAlignRightCenter
        .ColWidth(DisRs) = 840
        
        .TextMatrix(0, DisOrd) = "ODis%"
        .ColAlignmentFixed(DisOrd) = flexAlignRightCenter
        .ColWidth(DisOrd) = 555

        .TextMatrix(0, DisOrdRs) = "OrdDisc"
        .ColAlignmentFixed(DisOrdRs) = flexAlignRightCenter
        .ColWidth(DisOrdRs) = 840
        
           .TextMatrix(0, SFCPer) = "SFCPer"
            .ColAlignmentFixed(SFCPer) = flexAlignRightCenter
            .ColWidth(SFCPer) = 840
            
            .TextMatrix(0, SFCAmt1) = "SFCAmt"
            .ColAlignmentFixed(SFCAmt1) = flexAlignRightCenter
            .ColWidth(SFCAmt1) = 840
            
            
        If PubVATYN = 1 Then
            .TextMatrix(0, TaxPer) = "TaxPer"
            .ColAlignmentFixed(TaxPer) = flexAlignRightCenter
            .ColWidth(TaxPer) = 840
            
            .TextMatrix(0, TaxAmt1) = "TaxAmt"
            .ColAlignmentFixed(TaxAmt1) = flexAlignRightCenter
            .ColWidth(TaxAmt1) = 840
            
            If PubSatYn = 1 Then
                .TextMatrix(0, SatPer) = "Sat %"
                .ColAlignmentFixed(SatPer) = flexAlignRightCenter
                .ColWidth(SatPer) = 840
                
                .TextMatrix(0, SatAmt1) = "Sat Amt"
                .ColAlignmentFixed(SatAmt1) = flexAlignRightCenter
                .ColWidth(SatAmt1) = 840
            Else
                .ColWidth(SatPer) = 0
                .ColWidth(SatAmt1) = 0
            End If
        Else
            .TextMatrix(0, TaxPer) = ""
            .ColAlignmentFixed(TaxPer) = flexAlignRightCenter
            .ColWidth(TaxPer) = 0
        
            .TextMatrix(0, TaxAmt1) = ""
            .ColAlignmentFixed(TaxAmt1) = flexAlignRightCenter
            .ColWidth(TaxAmt1) = 0
            
            .ColWidth(SatPer) = 0
            .ColWidth(SatAmt1) = 0
        End If
        
        .TextMatrix(0, ItemVal) = "Item Value"
        .ColAlignmentFixed(ItemVal) = flexAlignRightCenter
        .ColWidth(ItemVal) = 1095
        
        .TextMatrix(0, Godown) = "Godown"
        .ColAlignmentFixed(Godown) = flexAlignRightCenter
        .ColWidth(Godown) = 1095
        
        .TextMatrix(0, PartSrlNo) = "Part SrlNo"
        .ColAlignmentFixed(PartSrlNo) = flexAlignLeftCenter
        .ColAlignment(PartSrlNo) = flexAlignLeftCenter
        .ColWidth(PartSrlNo) = 1095
        
        .TextMatrix(0, NDP) = "NDP"
        .ColAlignmentFixed(NDP) = flexAlignRightCenter
        .ColWidth(NDP) = 870

        .TextMatrix(0, PName) = "Part Name"
        .ColAlignment(PName) = flexAlignLeftCenter
        .ColWidth(PName) = 2500
        
        .TextMatrix(0, LName) = "Local Name"
        .ColAlignment(LName) = flexAlignLeftCenter
        .ColWidth(LName) = 2000
    End With
   
    For I = 19 To 38
        FGrid.ColWidth(I) = 0
    Next
    FGrid.ColAlignment(18) = flexAlignLeftCenter
    
    With FGridSel
        .RowHeightMin = PubGridRowHeight
        .ColAlignmentFixed = flexAlignCenterCenter
        .TextMatrix(0, 1) = "MR No."
        .ColAlignment(1) = flexAlignLeftCenter
        .ColWidth(1) = 1230
        
        .TextMatrix(0, 2) = "MR Date"
        .ColAlignment(2) = flexAlignLeftCenter
        .ColWidth(2) = 1250
        
        .TextMatrix(0, 3) = "Supplier Doc.No."
        .ColAlignment(3) = flexAlignLeftCenter
        .ColWidth(3) = 1530
        
        .TextMatrix(0, 4) = "Doc. Date"
        .ColAlignment(4) = flexAlignLeftCenter
        .ColWidth(4) = 1250
        
        .TextMatrix(0, 5) = "Total Amt"
        .ColAlignment(5) = flexAlignRightCenter
        .ColWidth(5) = 1250
        
        .TextMatrix(0, 6) = "DocID"
        .ColAlignment(6) = flexAlignLeftCenter
        .ColWidth(6) = 1815
        .width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + 300
    End With
    FrmSel.top = 1000
    FrmSel.width = FGridSel.width + 90
    
    FrmDetail.width = 6285: FrmDetail.left = Me.width - (FrmDetail.width + mRtScale): FrmDetail.top = mTopScale: FrmDetail.height = 2130
    FGrid.left = Me.left: FGrid.width = Me.width - 90: FGrid.top = 2500 ': FGrid.height = 2895
    DGPart.width = FGrid.width: DGPart.left = FGrid.left: DGPart.top = FGrid.top + FGrid.height: DGPart.height = Me.height - (FGrid.top + FGrid.height + mBotScale)
    DGPONo.left = FGrid.left: DGPONo.top = DGPart.top: DGPONo.height = Me.height - (FGrid.top + FGrid.height + mBotScale)
    DGGod.left = Me.width - (DGGod.width + mRtScale): DGGod.top = mTopScale
        
    DGParty.width = 11535:   DGParty.left = 0: DGParty.top = FGrid.top  '390
    'DGParty.height = 5160
    DGTrans.width = 6000: DGTrans.left = FGrid.left: DGTrans.top = FGrid.top: DGTrans.height = FGrid.height
    DGForm.width = 6000: DGForm.left = FGrid.left: DGForm.top = FGrid.top: DGForm.height = FGrid.height
    DGOrdPart.left = (Me.width - DGOrdPart.width) / 2: DGOrdPart.top = FGrid.top + FGrid.height: DGOrdPart.height = Me.height - (FGrid.top + FGrid.height + mBotScale)
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Double
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
    txt(I).ForeColor = CtrlFColOrg
Next
txt(TxtDocID).Enabled = False
If TopCtrl1.TopText2 = "Edit" Then
    txt(VDate).Enabled = False
    txt(SerialNo).Enabled = False
    txt(VType).Enabled = False
    txt(OPtSel).Enabled = False
End If
FldEnabled False

txtDisabled_Color Me

TxtGrid(0).BackColor = CtrlBCol
TxtGrid(0).ForeColor = CtrlFCol

If PubSiebelActiveYn = 1 And pubUName = "SA" Then
    cmdPost.Visible = True
Else
    cmdPost.Visible = False
End If

End Sub
Private Sub Grid_Hide()
    If DGPart.Visible = True Then DGPart.Visible = False
    If DGTrans.Visible = True Then DGTrans.Visible = False
    If DGForm.Visible = True Then DGForm.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If lblGroup.Visible = True Then lblGroup.Visible = False
    If DGPONo.Visible = True Then DGPONo.Visible = False
    If DGGod.Visible = True Then DGGod.Visible = False
   ' If DGOrdPart.Visible = True Then DGOrdPart.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub
Private Sub DGParty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DGParty.Row >= 0 Then
    lblGroup.TEXT = G_FaCn.Execute("Select AcGroup.GroupName from (AcGroup Left Join SubGroup on SubGroup.GroupCode=AcGroup.GroupCode) where SubGroup.SubCode='" & RsParty!Code & "'").Fields(0).Value
    lblGroup.Refresh
End If
End Sub
Private Sub Amt_Cal1()
'      FGrid.TextMatrix(FGrid.Row, FRate) = Format((Val(FGrid.TextMatrix(FGrid.Row, NDP)) - Val(FGrid.TextMatrix(FGrid.Row, DisRs))), "0.00")
'      FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, NDP)) * Val(FGrid.TextMatrix(FGrid.Row, PQty)), "0.00")
'      FGrid.TextMatrix(FGrid.Row, ItemVal) = Format((Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, PQty))) - Val(FGrid.TextMatrix(FGrid.Row, DisOrdRs)), "0.00")
Dim mAmount As Double, TaxAmt As Double, DisAmt As Double, OrdDisAmt1 As Double
Dim mTaxableAmt As Double
    If FGrid.TextMatrix(FGrid.Row, DisPer) <> "" Then
        FGrid.TextMatrix(FGrid.Row, DisRs) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(FGrid.TextMatrix(FGrid.Row, DisPer)) / 100, "0.00")
    End If
    If FGrid.TextMatrix(FGrid.Row, DisOrd) <> "" Then
        FGrid.TextMatrix(FGrid.Row, DisOrdRs) = Format((Val(FGrid.TextMatrix(FGrid.Row, Amt)) - Val(FGrid.TextMatrix(FGrid.Row, DisRs))) * Val(FGrid.TextMatrix(FGrid.Row, DisOrd)) / 100, "0.00")
    End If
    FGrid.TextMatrix(FGrid.Row, ItemVal) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) - Val(FGrid.TextMatrix(FGrid.Row, DisRs)) - Val(FGrid.TextMatrix(FGrid.Row, DisOrdRs)), "0.00")
    If Val(FGrid.TextMatrix(FGrid.Row, PQty)) <> 0 Then
        FGrid.TextMatrix(FGrid.Row, NDP) = Format(Val(FGrid.TextMatrix(FGrid.Row, ItemVal)) / Val(FGrid.TextMatrix(FGrid.Row, PQty)), "0.00")
    Else
        FGrid.TextMatrix(FGrid.Row, NDP) = ""
    End If
    If PubVATYN = 1 Then
        If FGrid.TextMatrix(FGrid.Row, TaxPer) <> "" Then
            mAmount = Val(FGrid.TextMatrix(FGrid.Row, Amt))
            DisAmt = Val(FGrid.TextMatrix(FGrid.Row, DisRs))
            OrdDisAmt1 = Val(FGrid.TextMatrix(FGrid.Row, DisOrdRs))
            If FGrid.TextMatrix(FGrid.Row, MRP) = "Yes" And FGrid.TextMatrix(FGrid.Row, Taxable) = "Yes" Then
                If StrCmp(left(PubComp_Name, 3), "jmk") Then
                 FGrid.TextMatrix(FGrid.Row, SFCAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, SFCPer)) / 100, "0.00")
                 
                    FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1) + Val(FGrid.TextMatrix(FGrid.Row, SFCAmt1))) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
                    If mSatYn Then
                        FGrid.TextMatrix(FGrid.Row, SatAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1) + Val(FGrid.TextMatrix(FGrid.Row, SFCAmt1))) * Val(FGrid.TextMatrix(FGrid.Row, SatPer)) / 100, "0.00")
                    End If
                Else
                    If mSatYn Then
                        mTaxableAmt = Format((mAmount - (DisAmt + OrdDisAmt1)) * 100 / (100 + Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) + Val(FGrid.TextMatrix(FGrid.Row, SatPer))), "0.00")
                        FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
                        FGrid.TextMatrix(FGrid.Row, SatAmt1) = Format(mTaxableAmt * Val(FGrid.TextMatrix(FGrid.Row, SatPer)) / 100, "0.00")
                    Else
                        FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / (100 + Val(FGrid.TextMatrix(FGrid.Row, TaxPer))), "0.00")
                    End If
                    FGrid.TextMatrix(FGrid.Row, ItemVal) = Format(Val(FGrid.TextMatrix(FGrid.Row, ItemVal)) - Val(FGrid.TextMatrix(FGrid.Row, TaxAmt1)) - Val(FGrid.TextMatrix(FGrid.Row, SatAmt1)), "0.00")
                End If
            ElseIf FGrid.TextMatrix(FGrid.Row, MRP) = "No" And FGrid.TextMatrix(FGrid.Row, Taxable) = "Yes" Then
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, TaxPer)) / 100, "0.00")
                If mSatYn Then
                    FGrid.TextMatrix(FGrid.Row, SatAmt1) = Format((mAmount - (DisAmt + OrdDisAmt1)) * Val(FGrid.TextMatrix(FGrid.Row, SatPer)) / 100, "0.00")
                End If
            Else
                FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
                FGrid.TextMatrix(FGrid.Row, SatAmt1) = ""
            End If
        End If
    End If
End Sub
 
 Private Sub Amt_Cal(Optional Index As Integer)
 Dim I As Double
 Dim IQty As Double, DQty1 As Double, ICnt As Integer, IGAmt As Double
 Dim IDic As Double, IOrdDic As Double, IAmt As Double, TotSpr As Double, TotOil As Double
 Dim mNetAmt As Double
 Dim TaxPer As Double
 Dim mTaxPer As Double
 Dim NSFCAmt As Double
 Dim TaxAmount As Double, TaxAmountMRP As Double
 Dim SatAmount As Double
 Dim SurPer As Double
 For I = 1 To FGrid.Rows - 1
    If FGrid.TextMatrix(I, PNo) <> "" Then
        IQty = IQty + Val(FGrid.TextMatrix(I, PQty))
        DQty1 = DQty1 + Val(FGrid.TextMatrix(I, DQty))
        IAmt = IAmt + Val(FGrid.TextMatrix(I, Amt))
        IDic = IDic + Val(FGrid.TextMatrix(I, DisRs))
        IOrdDic = IOrdDic + Val(FGrid.TextMatrix(I, DisOrdRs))
        IGAmt = IGAmt + Val(FGrid.TextMatrix(I, ItemVal)) + Val(FGrid.TextMatrix(I, SFCAmt1))
        NSFCAmt = NSFCAmt + Val(FGrid.TextMatrix(I, SFCAmt1))
        If FGrid.TextMatrix(I, PartGrade) = PubPartGrade_Lub Then
            TotOil = TotOil + Val(FGrid.TextMatrix(I, ItemVal))
        Else
            TotSpr = TotSpr + Val(FGrid.TextMatrix(I, ItemVal))
        End If
        If PubVATYN = 1 Then
            If FGrid.TextMatrix(I, MRP) = "Yes" Then
                TaxAmountMRP = TaxAmountMRP + Val(FGrid.TextMatrix(I, TaxAmt1))
                TaxAmount = TaxAmount + Val(FGrid.TextMatrix(I, TaxAmt1))
            Else
                TaxAmount = TaxAmount + Val(FGrid.TextMatrix(I, TaxAmt1))
            End If
            SatAmount = SatAmount + Val(FGrid.TextMatrix(I, SatAmt1))
        Else
            If txt(FormType) <> "" Then
                TaxPer = GCn.Execute("Select Tax_Per from TaxForms where Form_Code='" & txt(FormType).Tag & "'").Fields(0).Value
                SurPer = GCn.Execute("Select Tax_Sur_Per from TaxForms where Form_Code='" & txt(FormType).Tag & "'").Fields(0).Value
                mTaxPer = TaxPer + (TaxPer * SurPer / 100)
                    If FGrid.TextMatrix(I, MRP) = "Yes" And FGrid.TextMatrix(I, Taxable) = "Yes" Then
                        TaxAmountMRP = TaxAmountMRP + Round(((Val(FGrid.TextMatrix(I, Amt)) - (Val(FGrid.TextMatrix(I, DisRs)) + Val(FGrid.TextMatrix(I, DisOrdRs)))) * mTaxPer) / (100 + mTaxPer), 2)
                        If FGrid.TextMatrix(I, Taxable) = "Yes" Then
                            TaxAmount = TaxAmount + Round(((Val(FGrid.TextMatrix(I, Amt)) - (Val(FGrid.TextMatrix(I, DisRs)) + Val(FGrid.TextMatrix(I, DisOrdRs)))) * mTaxPer) / (100 + mTaxPer), 2)
                        End If
                    Else
                        If FGrid.TextMatrix(I, Taxable) = "Yes" Then
                            TaxAmount = TaxAmount + Round(((Val(FGrid.TextMatrix(I, Amt)) - (Val(FGrid.TextMatrix(I, DisRs)) + Val(FGrid.TextMatrix(I, DisOrdRs)))) * mTaxPer) / 100, 2)
                        End If
                    End If
            End If
        End If
        ICnt = ICnt + 1
    End If
Next I
    LblIVal.CAPTION = Format(ICnt, "0")
    LblPQty.CAPTION = Format(IQty, "0.000")
    LblDQty.CAPTION = Format(DQty1, "0.000")
    LblAmt.CAPTION = Format(IGAmt, "0.00")
    txt(TOTAmt) = Format(IAmt, "0.00")
    txt(TotDis) = Format(IDic, "0.00")
     txt(SFCAmt).TEXT = Format(NSFCAmt, "0.00")
    txt(TotOrdDis) = Format(IOrdDic, "0.00")
    txt(TotGoods) = Format(IGAmt, "0.00")
    txt(SprAmt) = Format(TotSpr, "0.00")
    txt(OilAmt) = Format(TotOil, "0.00")
    If TaxAmount > 0 Then
        txt(TaxAmt) = Format(TaxAmount, "0.00")
    End If
    txt(SatAmt) = Format(SatAmount, "0.00")
    If PubVATYN = 1 Then
        If StrCmp(left(PubComp_Name, 3), "JMK") Then
        'kunal
           ' txt(NetAmt) = Format((IGAmt + Val(txt(TaxAmt)) + Val(txt(SatAmt)) - TaxAmountMRP + Val(txt(Addition)) - Val(txt(Deduction))), "0.00")
            txt(NetAmt) = Format((IGAmt + Val(txt(TaxAmt)) + Val(txt(SatAmt)) - Val(txt(Addition)) - Val(txt(Deduction))), "0.00")
        Else
            txt(NetAmt) = Format((IGAmt + Val(txt(TaxAmt)) + Val(txt(SatAmt)) + Val(txt(Addition)) - Val(txt(Deduction))), "0.00")
        End If
    Else
        txt(NetAmt) = Format((IGAmt + Val(txt(TaxAmt)) - TaxAmountMRP + Val(txt(Addition)) - Val(txt(Deduction))), "0.00")
    End If
    'For Entry Tax
    mNetAmt = Val(txt(NetAmt)) + Val(txt(Transportation))
    If Index <> EntryTaxAmt Then
        txt(EntryTaxAmt) = Format(mNetAmt * Val(txt(EntryTaxPer)) / 100, "0.00")
    End If
    txt(TotPurAmt) = Format(mNetAmt + Val(txt(EntryTaxAmt)), "0.00")
    'Eof Entry Tax
 End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    Grid_Hide
    If FrmDetail.Visible = False Then FrmDetail.Visible = True
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    TxtGrid(Index).MaxLength = 0
    Select Case FGrid.Col
        Case PONo
            If rsPONo.RecordCount = 0 Or (rsPONo.EOF = True Or rsPONo.BOF = True) Or FGrid.TextMatrix(FGrid.Row, PONo) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, PONOCode) <> rsPONo!Code Then
                rsPONo.MoveFirst
                rsPONo.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, PONOCode) & "'"
            End If
         Case PNo
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "CODE"
            If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "code ='" & FGrid.TextMatrix(FGrid.Row, PNo) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case PName
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "name"
            If FGrid.TextMatrix(FGrid.Row, PName) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, PName) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case LName
            If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Then Exit Sub
            RsPart.Sort = "lname"
            If FGrid.TextMatrix(FGrid.Row, LName) <> "" Then
                RsPart.MoveFirst
                RsPart.FIND "lname ='" & FGrid.TextMatrix(FGrid.Row, LName) & "'"
                If RsPart.EOF = True Then RsPart.MoveFirst
            End If
        Case Godown
            If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Godown) = "" Then Exit Sub
            If FGrid.TextMatrix(FGrid.Row, Godown) <> rsGod!Name Then
                rsGod.MoveFirst
                rsGod.FIND "name ='" & FGrid.TextMatrix(FGrid.Row, Godown) & "'"
            End If
        Case PartSrlNo
            TxtGrid(Index).MaxLength = 20
    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then TxtGrid(0).TEXT = TxtGrid(0).Tag: Exit Sub
    Select Case FGrid.Col
        Case PONo   '3
            DGridTxtKeyDown DGPONo, TxtGrid, Index, rsPONo, KeyCode, True, 1
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                   GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo
                End If
            End If
        Case PNo    '1
           ' If DGPart.Visible = False Then DGridColSwap DGPart, 0
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 0, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo, 1
                End If
            End If
        Case Godown
            DGridTxtKeyDown DGGod, TxtGrid, 0, rsGod, KeyCode, True, 1, frmGodown, "frmGodown"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, 0, KeyCode, TAddMode, PONo
                End If
            End If
        Case Taxable, MRP, DQty, PQty, NDP, DisRs, DisOrdRs, TaxAmt1, SatAmt1
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo
                End If
            End If
            
        Case FRate
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo, , DisPer
                End If
            End If
            
        Case DisPer
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo, , DisOrd
                End If
            End If
            
            
        Case DisOrd
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo, , Godown
                End If
            End If
            
        Case TaxPer, SatPer
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo, , Godown
                End If
            End If
            
        Case PName
            If DGPart.Visible = False Then DGridColSwap DGPart, 1
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 1, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo
                End If
            End If
        Case LName   '3
            If DGPart.Visible = False Then DGridColSwap DGPart, 2
            DGridTxtKeyDown DGPart, TxtGrid, Index, RsPart, KeyCode, True, 2, frmPartMast, "frmPartMast"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, PartSrlNo
                End If
            End If
    End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
If KeyAscii = vbKeyEscape Then Exit Sub
Call CheckQuote(KeyAscii)
Select Case FGrid.Col
    Case PNo
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsPart, KeyAscii, "CODE"
    Case PName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsPart, KeyAscii, "name"
    Case LName
        If DGPart.Visible = True Then DGridTxtKeyPress TxtGrid, 0, RsPart, KeyAscii, "Lname"
    Case PONo
        If DGPONo.Visible = True Then DGridTxtKeyPress TxtGrid, 0, rsPONo, KeyAscii, "name"
    Case PQty, DQty
        NumPress TxtGrid(Index), KeyAscii, 8, 3
    Case DisPer, DisOrd
        NumPress TxtGrid(Index), KeyAscii, 2, 2
    Case DisRs, DisOrdRs
        NumPress TxtGrid(Index), KeyAscii, 8, 2
    Case FRate
        NumPress TxtGrid(Index), KeyAscii, 8, 4
    Case Godown
        If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
            If DGGod.Visible = True Then DGridTxtKeyPress TxtGrid, 0, rsGod, KeyAscii, "Name"
        End If
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case FGrid.Col
    Case PONo
        If KeyCode <> 13 And DGPONo.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, rsPONo, KeyCode, "name", True
    Case PNo
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0:   DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "CODE", True
    Case PName
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "name", True
    Case LName
        If KeyCode <> 13 And DGPart.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Lname", True
    Case Godown
'        If KeyCode <> 13 And DGGod.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, 0, rsGod, KeyCode, "Name", True
        If KeyCode <> 13 And DGGod.Visible = False Then
            TxtGrid_KeyDown Index, GridKey, 0
            If PubRestrict_Godown = 0 Then      ' Restrict Godown is "NO"
                DGridTxtKeyPress TxtGrid, 0, rsGod, KeyCode, "Name", True
            End If
        End If
    Case Taxable
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            TxtGrid(Index) = ""
        ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Or TxtGrid(Index) = "" Then
            TxtGrid(Index) = "Yes"
        Else
            TxtGrid(Index) = "No"
        End If
        
    Case MRP
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            TxtGrid(Index) = ""
        ElseIf UCase(left$(TxtGrid(Index), 1)) = "Y" Then
            TxtGrid(Index) = "Yes"
        Else
            TxtGrid(Index) = "No"
        End If
        
    Case FRate  'ndp
        FGrid.TextMatrix(FGrid.Row, FRate) = Format(Val(TxtGrid(Index).TEXT), "0.0000")
        FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, DQty)), "0.00")
    Case PQty, DQty
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index).TEXT), "0.000")
        FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, DQty)), "0.00")
    Case DisPer
        If TxtGrid(Index) <> "" Then
            FGrid.TextMatrix(FGrid.Row, DisPer) = Format(TxtGrid(Index).TEXT, "0.00")
            FGrid.TextMatrix(FGrid.Row, DisRs) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) * Val(TxtGrid(Index).TEXT) / 100, "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, DisPer) = ""
            FGrid.TextMatrix(FGrid.Row, DisRs) = ""
        End If
    Case DisRs
        If TxtGrid(Index) <> "" Then
            FGrid.TextMatrix(FGrid.Row, DisRs) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, DisPer) = ""
            FGrid.TextMatrix(FGrid.Row, DisRs) = ""
        End If
    Case DisOrd
        If Val(TxtGrid(Index)) <> 0 Then
            FGrid.TextMatrix(FGrid.Row, DisOrd) = Format(TxtGrid(Index).TEXT, "0.00")
            FGrid.TextMatrix(FGrid.Row, DisOrdRs) = Format(Val(FGrid.TextMatrix(FGrid.Row, Amt)) - Val(FGrid.TextMatrix(FGrid.Row, DisRs)) * Val(TxtGrid(Index).TEXT) / 100, "0.00")
        Else
            FGrid.TextMatrix(FGrid.Row, DisOrd) = ""
            FGrid.TextMatrix(FGrid.Row, DisOrdRs) = ""
        End If
    Case DisOrdRs
        If Val(TxtGrid(Index)) <> 0 Then
            FGrid.TextMatrix(FGrid.Row, DisOrdRs) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        Else
           FGrid.TextMatrix(FGrid.Row, DisOrd) = ""
           FGrid.TextMatrix(FGrid.Row, DisOrdRs) = ""
        End If
    Case PartSrlNo
        FGrid.TextMatrix(FGrid.Row, PartSrlNo) = TxtGrid(Index)
End Select
Amt_Cal1
Amt_Cal
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
    Grid_Hide
End If
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid.Col
    Case PONo
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        TxtGridValid_PONo
    Case PNo, PName, LName
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        TxtGridValid_PNo
    Case Taxable, MRP
        If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
        TxtGridValid_TaxMRP
        Amt_Cal1
        Amt_Cal
    Case Godown
        TxtGridValid_Godown
    Case FRate
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(Index) = "", "", Format(Val(TxtGrid(Index).TEXT), "0.0000"))
        FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, DQty)), "0.00")
        If PubVATYN = 1 Then
           If txt(FormType).Tag <> "" Then
                Set rsTaxPer = GCn.Execute("Select Tax_Per,SFCPER from TaxForms where Form_Code='" & txt(FormType).Tag & "'")
                 If rsTaxPer.RecordCount > 0 Then
                    FGrid.TextMatrix(FGrid.Row, TaxPer) = rsTaxPer!Tax_Per
                    FGrid.TextMatrix(FGrid.Row, SFCPer) = VNull(rsTaxPer!SFCPer)
                 End If
           End If
        End If

        Amt_Cal1
        Amt_Cal
    Case DQty, PQty
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(Index) = "", "", Format(Val(TxtGrid(Index).TEXT), "0.000"))
        If FGrid.Col = DQty Then
            If Val(FGrid.TextMatrix(FGrid.Row, PQty)) = 0 Then
                FGrid.TextMatrix(FGrid.Row, PQty) = FGrid.TextMatrix(FGrid.Row, DQty)
            End If
        End If
        FGrid.TextMatrix(FGrid.Row, Amt) = Format(Val(FGrid.TextMatrix(FGrid.Row, FRate)) * Val(FGrid.TextMatrix(FGrid.Row, DQty)), "0.00")
        Amt_Cal1
        Amt_Cal
        If rsGod.RecordCount > 0 And Trim(FGrid.TextMatrix(FGrid.Row, Godown)) = "" Then
            rsGod.MoveFirst
            rsGod.FIND "Code ='" & PubSprCounterGodown & "'"
            FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
            FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
        End If
    Case DisPer
        FGrid.TextMatrix(FGrid.Row, DisPer) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0).TEXT, "0.00"))
        If FGrid.TextMatrix(FGrid.Row, DisPer) = "" Then FGrid.TextMatrix(FGrid.Row, DisRs) = ""
        Amt_Cal1
        Amt_Cal
    Case DisRs
        FGrid.TextMatrix(FGrid.Row, DisRs) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0).TEXT, "0.00"))
        If Val(FGrid.TextMatrix(FGrid.Row, DisRs)) + Val(FGrid.TextMatrix(FGrid.Row, DisOrd)) > Val(FGrid.TextMatrix(FGrid.Row, Amt)) Then
            TxtGridLeave = False: Exit Function
        End If
        Amt_Cal1
        Amt_Cal
    Case DisOrd
        FGrid.TextMatrix(FGrid.Row, DisOrd) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0).TEXT, "0.00"))
        If Val(FGrid.TextMatrix(FGrid.Row, DisOrd)) = 0 Then FGrid.TextMatrix(FGrid.Row, DisOrdRs) = 0
        Amt_Cal1
        Amt_Cal
    Case DisOrdRs
        FGrid.TextMatrix(FGrid.Row, DisOrdRs) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0).TEXT, "0.00"))
        If Val(FGrid.TextMatrix(FGrid.Row, DisRs)) + Val(FGrid.TextMatrix(FGrid.Row, DisOrd)) > Val(FGrid.TextMatrix(FGrid.Row, Amt)) Then
            TxtGridLeave = False: Exit Function
        End If
        Amt_Cal1
        Amt_Cal
     Case TaxPer
        FGrid.TextMatrix(FGrid.Row, TaxPer) = TxtGrid(0)
        If FGrid.TextMatrix(FGrid.Row, TaxPer) = "" Then FGrid.TextMatrix(FGrid.Row, TaxAmt1) = ""
        Amt_Cal1
        Amt_Cal
     Case SatPer
        FGrid.TextMatrix(FGrid.Row, SatPer) = TxtGrid(0)
        If FGrid.TextMatrix(FGrid.Row, SatPer) = "" Then FGrid.TextMatrix(FGrid.Row, SatAmt1) = ""
        Amt_Cal1
        Amt_Cal
        
End Select
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    TxtGrid(0).Visible = False
End If
End Function

Private Function ChkDuplicate() As Boolean
Dim I As Double
Dim X As String, Y As String
Dim Col1 As Byte, Col2 As Byte, Col3 As Byte, Col4 As Byte
    Select Case FGrid.Col
    Case PNo, PName, LName
        Col4 = FGrid.Col
        Col1 = PONo
        Col2 = Taxable
        Col3 = MRP
    Case MRP
        Col1 = PNo
        Col2 = PONo
        Col3 = Taxable
        Col4 = MRP
    Case Taxable
        Col1 = PNo
        Col2 = PONo
        Col4 = Taxable
        Col3 = MRP
    Case PONo
        Col1 = PNo
        Col4 = PONo
        Col2 = Taxable
        Col3 = MRP
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

Private Sub FillGridData()
Dim Rs As ADODB.Recordset
Dim I As Double

FGrid.Rows = 1
For I = 1 To FGridSel.Rows - 1
    If FGridSel.TextMatrix(I, 0) <> "" Then
        Set Rs = GCn.Execute("Select P.L_C, P.Form_Code, T.Form_Desc, P.Party_Doc_No, P.Party_Doc_Date, P.Transport, P.Case_Mark, P.Case_No  From (Sp_Purch P " & _
                            "Left Join TaxForms T On T.Form_Code = P.Form_Code) " & _
                            "Where DocId = '" & FGridSel.TextMatrix(I, 6) & "'")
        If Rs.RecordCount > 0 Then
            txt(LC) = IIf(Rs!L_C = "L", "Local", "Central")
            txt(FormType) = XNull(Rs!form_desc)
            txt(FormType).Tag = XNull(Rs!Form_Code)
            txt(SuppChlNo) = XNull(Rs!Party_Doc_No)
            txt(SuppChlDate) = XNull(Rs!Party_Doc_Date)
            txt(Transport) = XNull(Rs!Transport)
            txt(CaseMark) = XNull(Rs!Case_Mark)
            txt(CaseNo) = XNull(Rs!Case_No)
        End If
        Set Rs = New Recordset
         Set Rs = GCn.Execute("SELECT SPO.Order_Reg_No,P.Part_Name," & cTrim(cMID("Stk.Order_DocID", "8", "5")) & "+ " & cCStr(cTrim("Right(Stk.Order_DocID,8)")) & " As OrderIDDisp, " & _
            " P.Local_Name,P.Part_Grade, P.UNIT, P.MRP, P.Cur_MRP_TBStk, P.Cur_MRP_TPStk, P.Cur_TB_Stk, P.Cur_TP_Stk,P.TP_SRate, P.TB_SRate, P.Bin_Loca, P.High_Pur_Rate, P.Low_Pur_Rate, Stk.*, G.God_Name" & _
            " FROM ((Sp_Stock Stk LEFT JOIN Part P ON Stk.Part_No = P.PART_NO and P.Div_Code = left(STK.DocID,1)) " & _
            " LEFT JOIN Godown G ON Stk.Godown = G.God_Code) " & _
            " Left Join SP_Order SPO on Stk.Order_DocID=SPO.OrderID " & _
            " where Stk.docId = '" & FGridSel.TextMatrix(I, 6) & "'")
            
        If Rs.RecordCount > 0 Then
            Do Until Rs.EOF
                FGrid.AddItem ""
                With FGrid
                    .TextMatrix(.Rows - 1, ChlDocId) = IIf(IsNull(Rs!DocID), "", Rs!DocID)
                    .TextMatrix(.Rows - 1, ChlDocSr) = Rs!Srl_No
                    .TextMatrix(.Rows - 1, 0) = FGrid.Rows - 1
                    .TextMatrix(.Rows - 1, PONo) = IIf(IsNull(Rs!Order_Reg_No), "", Rs!Order_Reg_No)
                    .TextMatrix(.Rows - 1, PNo) = Rs!Part_No
                    .TextMatrix(.Rows - 1, PONOCode) = XNull(Rs!Order_DocId)
                    .TextMatrix(.Rows - 1, POSrlNo) = XNull(Rs!Order_Srl_No)
                    .TextMatrix(.Rows - 1, Unit) = IIf(IsNull(Rs!Unit), "", Rs!Unit)
                    .TextMatrix(.Rows - 1, MRP) = IIf(Rs!MRP_YN = 1, "Yes", "No")
                    .TextMatrix(.Rows - 1, Taxable) = IIf(Rs!Tax_YN = 1, "Yes", "No")
                    .TextMatrix(.Rows - 1, DQty) = IIf(Rs!Qty_Doc = 0, "", Format(Rs!Qty_Doc, "0.000"))
                    .TextMatrix(.Rows - 1, PQty) = IIf(Rs!Qty_Rec = 0, "", Format(Rs!Qty_Rec, "0.000"))
'                   .TextMatrix(.Rows - 1, FRate) = IIf(rs!MRP_RATE = 0, "", Format(rs!MRP_RATE, "0.00"))
                    .TextMatrix(.Rows - 1, FRate) = IIf(Rs!Rate = 0, "", Format(Rs!Rate, "0.0000"))
                    .TextMatrix(.Rows - 1, Amt) = IIf(Rs!Amount = 0, "", Format(Rs!Amount, "0.00"))
                    .TextMatrix(.Rows - 1, DisPer) = IIf(Rs!Disc_Per = 0, "", Format(Rs!Disc_Per, "0.00"))
                    .TextMatrix(.Rows - 1, DisRs) = IIf(Rs!Disc_Amt = 0, "", Format(Rs!Disc_Amt, "0.00"))
                    .TextMatrix(.Rows - 1, DisOrd) = IIf(Rs!ord_Discper = 0, "", Format(Rs!ord_Discper, "0.00"))
                    .TextMatrix(.Rows - 1, DisOrdRs) = IIf(Rs!ord_Discamt = 0, "", Format(Rs!ord_Discamt, "0.00"))
                    
                    If PubVATYN = 1 Then
                        If .TextMatrix(.Rows - 1, MRP) = "Yes" Then
                            .TextMatrix(.Rows - 1, SFCPer) = Format(VNull(Rs!SFCPer), "0.00")
                            .TextMatrix(.Rows - 1, SFCAmt1) = Format(VNull(Rs!SFCAmt), "0.00")
                            
                            .TextMatrix(.Rows - 1, TaxPer) = Format(VNull(Rs!TaxPer), "0.00")
                            .TextMatrix(.Rows - 1, TaxAmt1) = Format(VNull(Rs!TaxAmt), "0.00")
                            .TextMatrix(.Rows - 1, SatPer) = Format(VNull(Rs!SatPer), "0.00")
                            .TextMatrix(.Rows - 1, SatAmt1) = Format(VNull(Rs!SatAmt), "0.00")

                            '.TextMatrix(.Rows - 1, ItemVal) = Format(VNull(rs!AMOUNT) - Val(.TextMatrix(.Rows - 1, TaxAmt1)) - Val(.TextMatrix(.Rows - 1, DisRs)), "0.00")
                            .TextMatrix(.Rows - 1, ItemVal) = Format(VNull(Rs!Net_Amt), "0.00")
                        Else
                            .TextMatrix(.Rows - 1, TaxPer) = Format(VNull(Rs!TaxPer), "0.00")
                            .TextMatrix(.Rows - 1, TaxAmt1) = Format(VNull(Rs!TaxAmt), "0.00")
                            .TextMatrix(.Rows - 1, SatPer) = Format(VNull(Rs!SatPer), "0.00")
                            .TextMatrix(.Rows - 1, SatAmt1) = Format(VNull(Rs!SatAmt), "0.00")
                            
                            '.TextMatrix(.Rows - 1, ItemVal) = IIf(rs!AMOUNT = 0, "", Format(rs!AMOUNT - Val(.TextMatrix(.Rows - 1, DisRs)), "0.00"))
                            .TextMatrix(.Rows - 1, ItemVal) = Format(VNull(Rs!Net_Amt), "0.00")
                        End If
                    Else
                        .TextMatrix(.Rows - 1, ItemVal) = IIf(Rs!Amount = 0, "", Format(Rs!Amount, "0.00"))
                    End If

                    .TextMatrix(.Rows - 1, NDP) = IIf(Rs!Rate = 0, "", Format(Rs!Rate, "0.00"))
                    .TextMatrix(.Rows - 1, God) = Rs!Godown
                    .TextMatrix(.Rows - 1, Godown) = IIf(IsNull(Rs!God_Name), "", Rs!God_Name)
                    .TextMatrix(.Rows - 1, PName) = IIf(IsNull(Rs!Part_Name), "", Rs!Part_Name)
                    .TextMatrix(.Rows - 1, LName) = IIf(IsNull(Rs!Local_Name), "", Rs!Local_Name)
                    .TextMatrix(.Rows - 1, MRPStkTB) = IIf(IsNull(Rs!Cur_MRP_TbStk), "", Rs!Cur_MRP_TbStk)
                    .TextMatrix(.Rows - 1, MRPStkTP) = IIf(IsNull(Rs!Cur_MRP_TPStk), "", Rs!Cur_MRP_TPStk)
                    .TextMatrix(.Rows - 1, MRPRate) = IIf(Rs!MRP = 0, "", Format(Rs!MRP, "0.00"))
                    .TextMatrix(.Rows - 1, TBStk) = IIf(IsNull(Rs!Cur_TB_STk), "", Rs!Cur_TB_STk)
                    .TextMatrix(.Rows - 1, TPStk) = IIf(IsNull(Rs!Cur_TP_Stk), "", Rs!Cur_TP_Stk)
                    .TextMatrix(.Rows - 1, TBRate) = IIf(IsNull(Rs!TB_SRate), "", Rs!TB_SRate)
                    .TextMatrix(.Rows - 1, TPRate) = IIf(IsNull(Rs!TP_SRate), "", Rs!TP_SRate)
                    .TextMatrix(.Rows - 1, Bin) = IIf(IsNull(Rs!Bin_Loca), "", Rs!Bin_Loca)
    '                    .TextMatrix(cnt, LastRate) = ""
                    .TextMatrix(.Rows - 1, HPRate) = IIf(IsNull(Rs!high_pur_rate), "", Rs!high_pur_rate)
                    .TextMatrix(.Rows - 1, LPRate) = IIf(IsNull(Rs!low_pur_rate), "", Rs!low_pur_rate)
                    .TextMatrix(.Rows - 1, PartGrade) = IIf(IsNull(Rs!Part_Grade), "", Rs!Part_Grade)
                End With
                Rs.MoveNext
            Loop
        End If
    End If
Next
Set Rs = Nothing
FrmSel.Visible = False
FGrid.AddItem FGrid.Rows
FGrid.FixedRows = 1
End Sub

Private Sub FldEnabled(Enb As Boolean)
'    txt(PermitType).Enabled = Enb
'    txt(FormNo).Enabled = Enb
'    txt(CaseNo).Enabled = Enb
'    txt(CaseMark).Enabled = Enb
'    txt(Transport).Enabled = Enb
'    txt(GrNo).Enabled = Enb
'    txt(GrDate).Enabled = Enb
    
    txtDisabled_Color Me

    TxtGrid(0).BackColor = CtrlBCol
    TxtGrid(0).ForeColor = CtrlFCol
End Sub

Private Sub RemoveTxtNull()
Dim I As Double
For I = 0 To txt.Count - 1
    txt(I).TEXT = IIf(IsNull(txt(I).TEXT), "", txt(I).TEXT)
Next I
End Sub

Private Sub TxtGridValid_PONo()
If rsPONo.RecordCount = 0 Or (rsPONo.EOF = True Or rsPONo.BOF = True) Or TxtGrid(0).TEXT = "" Then
    FGrid.TextMatrix(FGrid.Row, PONo) = ""
    FGrid.TextMatrix(FGrid.Row, PONOCode) = ""
Else
    FGrid.TextMatrix(FGrid.Row, PONo) = rsPONo!Name
    FGrid.TextMatrix(FGrid.Row, PONOCode) = rsPONo!Code
End If

End Sub
Private Sub TxtGridValid_PNo()
Dim OldPNo$
If RsPart.RecordCount = 0 Or (RsPart.EOF = True Or RsPart.BOF = True) Or TxtGrid(0).TEXT = "" Then
    FGrid.TextMatrix(FGrid.Row, PNo) = ""
    FGrid.TextMatrix(FGrid.Row, PName) = ""
    FGrid.TextMatrix(FGrid.Row, LName) = ""
    MainLib.Fill_Data mPartyType, LblFrm, FGrid, _
        "", "", "", Unit, MRP, Taxable, _
        MRPStkTB, MRPStkTP, TBStk, TPStk, _
        MRPRate, TBRate, TPRate, Bin, _
        HPRate, LPRate, LastRate, PartGrade, _
        EffectDate, DisPer, mCheckNegetiveStockSiteWise, True
Else
    OldPNo = FGrid.TextMatrix(FGrid.Row, PNo)
    FGrid.TextMatrix(FGrid.Row, PNo) = RsPart!Code
    FGrid.TextMatrix(FGrid.Row, PName) = RsPart!Name
    FGrid.TextMatrix(FGrid.Row, LName) = RsPart!LName
    MainLib.Fill_Data mPartyType, LblFrm, FGrid, _
        RsPart!Code, RsPart!Name, RsPart!LName, _
        Unit, MRP, Taxable, _
        MRPStkTB, MRPStkTP, TBStk, TPStk, _
        MRPRate, TBRate, TPRate, Bin, _
        HPRate, LPRate, LastRate, PartGrade, _
        EffectDate, DisPer, mCheckNegetiveStockSiteWise, True
''by LPS 27-04-2K2
'    If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> "" Then
'        If FGrid.TextMatrix(FGrid.Row, Col_PNo) <> OldPNo Then
'            FGrid.TextMatrix(FGrid.Row, Col_Rate) = GetRate(FGrid, CDate(Txt(Vdate).Text), FGrid.TextMatrix(FGrid.Row, Col_PNo), Col_MRP, Val(FGrid.TextMatrix(FGrid.Row, Col_MRPRate)), Col_Taxable, Val(FGrid.TextMatrix(FGrid.Row, Col_TBRate)), Val(FGrid.TextMatrix(FGrid.Row, Col_TPRate)), Col_EffectDate, Col_MRPRate)
''            FGrid.TextMatrix(FGrid.Row, Col_DiscPer) = Format(RsPart!SalDisc_Per, "0.00")
'        End If
'    End If
'******************** For Tax in Line File *************************
If PubVATYN = 1 Then
   If txt(FormType).Tag <> "" Then
        Set rsTaxPer = GCn.Execute("Select Tax_Per, AddTaxPer,L_C,SFCPER  from TaxForms where Form_Code='" & txt(FormType).Tag & "'")
         If rsTaxPer.RecordCount > 0 Then
            FGrid.TextMatrix(FGrid.Row, TaxPer) = rsTaxPer!Tax_Per
            FGrid.TextMatrix(FGrid.Row, SatPer) = XNull(rsTaxPer!AddTaxPer)
            FGrid.TextMatrix(FGrid.Row, SFCPer) = VNull(rsTaxPer!SFCPer)
            If UTrim(XNull(rsTaxPer!L_C)) = "LOCAL" Then
               Set rsTaxPer = GCn.Execute("Select VatPer, AddTaxPer From Part_Grade Where PartGrade_Code='" & FGrid.TextMatrix(FGrid.Row, PartGrade) & "'")
               If rsTaxPer.RecordCount > 0 Then
                   If VNull(rsTaxPer!VatPer) > 0 Then FGrid.TextMatrix(FGrid.Row, TaxPer) = Format(rsTaxPer!VatPer, "0.00")
                   If VNull(rsTaxPer!AddTaxPer) > 0 Then FGrid.TextMatrix(FGrid.Row, SatPer) = Format(rsTaxPer!AddTaxPer, "0.00")
               End If
            End If
         End If
   End If
End If
'*******************************************************************
End If

If FGrid.TextMatrix(FGrid.Row, PONOCode) <> "" Then
    GSQL = "Select s1.Srl_No From SP_Order1 S1 Where OrderID='" & FGrid.TextMatrix(FGrid.Row, PONOCode) & "' and Part_No='" & FGrid.TextMatrix(FGrid.Row, PNo) & "'"
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    If GRs.RecordCount > 1 Or GRs.RecordCount <= 0 Then
        Set GRs = New ADODB.Recordset
        GRs.CursorLocation = adUseClient
        GRs.Open "Select s1.Srl_No,s1.PART_NO,P.Part_Name,s1.QTY,s1.Sup_Qty,(s1.Qty-s1.Sup_Qty) As PendQty From SP_Order1 S1 Left Join Part P on S1.Part_no=P.Part_No and P.Div_Code = left(S1.OrderId,1) Where OrderID='" & FGrid.TextMatrix(FGrid.Row, PONOCode) & "'", GCn, adOpenStatic, adLockReadOnly
        Set DGOrdPart.DataSource = GRs
        GRs.FIND ("Part_No='" & FGrid.TextMatrix(FGrid.Row, PNo) & "'")
        If GRs.EOF Then
            GRs.MoveFirst
        End If
        DGOrdPart.Visible = True
        DGOrdPart.ZOrder 0
        DGOrdPart.SetFocus
    Else
        FGrid.TextMatrix(FGrid.Row, POSrlNo) = GRs!Srl_No
        Set GRs = Nothing
        FGrid.SetFocus
        DGOrdPart.Visible = False
    End If
End If
If FGrid.TextMatrix(FGrid.Rows - 1, PNo) <> "" Then FGrid.AddItem FGrid.Rows
End Sub

Private Sub TxtGridValid_Godown()
    If rsGod.RecordCount = 0 Or (rsGod.EOF = True Or rsGod.BOF = True) Or TxtGrid(0).TEXT = "" Then
        FGrid.TextMatrix(FGrid.Row, Godown) = ""
        FGrid.TextMatrix(FGrid.Row, God) = ""
    Else
        FGrid.TextMatrix(FGrid.Row, Godown) = rsGod!Name
        FGrid.TextMatrix(FGrid.Row, God) = rsGod!Code
    End If
End Sub

Private Sub TxtGridValid_TaxMRP()
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0).TEXT
    If FGrid.TextMatrix(FGrid.Row, PNo) <> "" Then
        FGrid.TextMatrix(FGrid.Row, FRate) = Format(GetRate(mPartyType, FGrid, CDate(txt(VDate)), FGrid.TextMatrix(FGrid.Row, PNo), MRP, Val(FGrid.TextMatrix(FGrid.Row, MRPRate)), Taxable, Val(FGrid.TextMatrix(FGrid.Row, TBRate)), Val(FGrid.TextMatrix(FGrid.Row, TPRate)), EffectDate, MRPRate), "0.0000")
        If Val(FGrid.TextMatrix(FGrid.Row, FRate)) <> 0 Then
            FGrid.TextMatrix(FGrid.Row, FRate) = Format(FGrid.TextMatrix(FGrid.Row, FRate), "0.0000")
        End If
    End If
End Sub

Private Function ProcAcPost(rsCtrlAc As ADODB.Recordset) As Boolean
On Error GoTo lblExit
Dim xNetAmt As Double, xEntryTaxAmt As Double, xTransportation As Double, TransAc$
'A/c Posting related declarations
Dim LedgAry() As LedgRec, mCommNarr$, ContraCodeCr$
Dim mResult As Byte, mNarr$, TaxSQL$, I As Double, j As Double
Dim mSprPurPfx$, mFADocID$, PartyCode$


    ApplyConsolidatedPosting CDate(txt(VDate))

    mSprPurPfx = "PPPPP"
    If txt(VType) = "Cash" And IsConsolidatedPosting Then
        xTransportation = VNull(GCn.Execute("select sum(Transportation) from SP_Purch where V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(TxtDocID), 8) & "' and CancelYN=0 ").Fields(0).Value)
        xEntryTaxAmt = VNull(GCn.Execute("select sum(EntryTaxAmt) from SP_Purch where V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(TxtDocID), 8) & "' and CancelYN=0 ").Fields(0).Value)
        GSQL = "select TF.PurSal_Ac_Code,TF.Tax_Ac_Code, AddTaxAc,sum(NET_AMT+EntryTaxAmt+Transportation) as NetAmt,sum(Tax_Amt) as TaxAmt, Sum(SatAmt) As SatAmt,TaxForms.L_C " & _
            "from (SP_Purch " & _
            "left join TaxFormsAc as TF on SP_Purch.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code) Left Join TaxForms on SP_Purch.Form_Code=TaxForms.Form_Code  " & _
            "where V_Date=" & ConvertDate(txt(VDate)) & " and left(docid,8)='" & left(txt(TxtDocID), 8) & _
            "' and " & vIsNull("CancelYN", "0") & "=0 Group by TF.PurSal_Ac_Code,TF.Tax_Ac_Code, tf.AddTaxAc,TaxForms.L_C"
        mNarr = "Through Spare Cash Purchase (Daily Posting)"
        mCommNarr = mNarr & " [Common]"
        'Undelete old Posting (individual if any)
        'LedgerUnPost GCnFaS, Txt(TxtDocId)
        'Create FA DocID for Daily Posting
        mFADocID = left(txt(TxtDocID), 8) & mSprPurPfx & "  " & Format(PubStartDate, "yy") & Format(txt(VDate), "mmdd")
        PartyCode = PubSprCashAc
    Else
        PartyCode = txt(Party).Tag
        mFADocID = txt(TxtDocID)
        mNarr = "Cr Purchase "
        If txt(SuppChlNo) <> "" Then
            mNarr = mNarr & " Party Document No." & txt(SuppChlNo)
        End If
        If txt(SuppChlDate) <> "" Then
            mNarr = mNarr & " Date " & txt(SuppChlDate)
        End If
        mCommNarr = mNarr & " [Common]"
        xEntryTaxAmt = Val(txt(EntryTaxAmt)) ' VNull(GCn.Execute("select sum(EntryTaxAmt) from SP_Purch where docid='" & Txt(TxtDocId) & "' and CancelYN=0 ").Fields(0).Value)
        xTransportation = Val(txt(Transportation)) ' VNull(GCn.Execute("select sum(EntryTaxAmt) from SP_Purch where docid='" & Txt(TxtDocId) & "' and CancelYN=0 ").Fields(0).Value)
        GSQL = "select TF.PurSal_Ac_Code,TF.Tax_Ac_Code, Tf.AddTaxAc,sum(NET_AMT+EntryTaxAmt+Transportation) as NetAmt,sum(Tax_Amt) as TaxAmt, Sum(SatAmt) As SatAmt,TaxForms.L_C " & _
            "from (SP_Purch " & _
            "left join TaxFormsAc as TF on SP_Purch.Form_Code+'" & PubDivCode & "'=TF.Form_Code+TF.Div_Code ) Left Join TaxForms on SP_Purch.Form_Code=TaxForms.Form_Code  " & _
            "where docid='" & txt(TxtDocID) & _
            "' group by TF.PurSal_Ac_Code,TF.Tax_Ac_Code, Tf.AddTaxAc,TaxForms.L_C"
    End If
    
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    '*** pURCHASE Amount Row
'   0.Purchase A/c
'   1.Party A/c or Cash A/c
'    Dim LedgAry() As LedgRec
    
'*********
    I = -1
    If I = -1 Then
         ReDim Preserve LedgAry(1)
         I = 0
    Else
         I = UBound(LedgAry) + 1
         ReDim Preserve LedgAry(I)
    End If
    Do While GRs.EOF = False
        If PubVATYN = 1 And XNull(GRs!L_C) = "Local" Then
            If IsNull(GRs!PurSal_Ac_Code) Or GRs!PurSal_Ac_Code = "" Then
                MsgBox "Please Define Purchase A/c in Tax Forms " & GRs!PurSal_Ac_Code & vbCrLf & "A/c Psoting Aborted", vbCritical, "A/c Posting"
                GoTo lblExit
            End If
            If IsNull(GRs!Tax_Ac_Code) Or GRs!Tax_Ac_Code = "" Then
                MsgBox "Please Define Tax A/c in Tax Forms " & GRs!Tax_Ac_Code & vbCrLf & "A/c Psoting Aborted", vbCritical, "A/c Posting"
                GoTo lblExit
            End If
            If mSatYn And XNull(GRs!L_C) = "Local" Then
                If IsNull(GRs!AddTaxAc) Or GRs!AddTaxAc = "" Then
                    MsgBox "Please Define Additional Tax A/c in Tax Forms " & XNull(GRs!AddTaxAc) & vbCrLf & "A/c Psoting Aborted", vbCritical, "A/c Posting"
                    GoTo lblExit
                End If
            End If
            If I = -1 Then
                ReDim Preserve LedgAry(1)
                I = 0
            Else
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
            End If
            LedgAry(I).SubCode = GRs!PurSal_Ac_Code
            LedgAry(I).AmtDr = VNull(GRs!NetAmt) - VNull(GRs!TaxAmt) - VNull(GRs!SatAmt)
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = PartyCode
 
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            
            LedgAry(I).SubCode = GRs!Tax_Ac_Code
            LedgAry(I).AmtDr = VNull(GRs!TaxAmt)
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = PartyCode

            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            
            LedgAry(I).SubCode = XNull(GRs!AddTaxAc)
            LedgAry(I).AmtDr = VNull(GRs!SatAmt)
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = PartyCode

            xNetAmt = xNetAmt + IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
            
        Else
            If IsNull(GRs!PurSal_Ac_Code) Or GRs!PurSal_Ac_Code = "" Then
                MsgBox "Please Define Purchase A/c in Tax Forms " & GRs!PurSal_Ac_Code & vbCrLf & "A/c Psoting Aborted", vbCritical, "A/c Posting"
                GoTo lblExit
            End If
            If I = -1 Then
                ReDim Preserve LedgAry(1)
                I = 0
            Else
                I = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(I)
            End If
            LedgAry(I).SubCode = GRs!PurSal_Ac_Code
            LedgAry(I).AmtDr = IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
            LedgAry(I).Narration = mNarr
            LedgAry(I).ContraSub = PartyCode
            
            xNetAmt = xNetAmt + IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
        End If
        GRs.MoveNext
    Loop
    If xTransportation <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!SprPurTrans_Ac
        LedgAry(I).AmtCr = xTransportation
        LedgAry(I).Narration = mNarr
    End If
    If xEntryTaxAmt <> 0 Then
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = rsCtrlAc!EntryTax_Ac
        LedgAry(I).AmtCr = xEntryTaxAmt
        LedgAry(I).Narration = mNarr
        
    End If
    I = UBound(LedgAry) + 1
    ReDim Preserve LedgAry(I)
    LedgAry(I).SubCode = PartyCode
    LedgAry(I).AmtCr = (xNetAmt - (xEntryTaxAmt + xTransportation))
    LedgAry(I).Narration = mNarr

    
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, mFADocID, CDate(txt(VDate)), mCommNarr)
    If mResult <> 1 Then
        MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
        ProcAcPost = True
    End If
lblExit:
    Set GRs = Nothing
    If err.NUMBER <> 0 Then
        MsgBox err.Description & vbCrLf & "Ledger Posting Terminated!", vbCritical
        ProcAcPost = True
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
'    Set RsTemp = GCn.Execute("Select * From Sp_Purch Where DocId='" & Master!SearchCode & "'")
'    If RsTemp.RecordCount > 0 Then
'        Set rsTemp1 = GCn.Execute("Select Count(*) As Item, Sum(Qty_Rec) As Qty, Sum(Net_Amt2) As Amt From Sp_Stock Where Invoice_DocId='" & Master!SearchCode & "'")
'        If RsTemp.RecordCount > 0 Then
'            mTotalItem = VNull(rsTemp1!Item)
'            mTotalQty = VNull(rsTemp1!Qty)
'            mGoodsValue = VNull(rsTemp1!Amt)
'        End If
'        GCn.Execute "Insert Into DeleteLog(DocId, Type, VType, VDate, Total_Item, " & _
'                            "Total_Qty, GoodsValue, Bill_Amt, Discount, Addition, " & _
'                            "Deduction, AutoYn, User_Name, Del_Date, Del_Time, EditDate, EditTime) Values( " & _
'                            "'" & Master!SearchCode & "', 'Spare Purchase Bill', '" & txt(VType) & "', " & ConvertDate(XNull(RsTemp!V_DATE)) & ", " & mTotalItem & ", " & _
'                            "" & mTotalQty & ", " & mGoodsValue & ", " & VNull(RsTemp!Net_Amt) & ", " & VNull(RsTemp!Tot_Disc_Amt) & ", " & VNull(RsTemp!Addition) & ", " & _
'                            "" & VNull(RsTemp!Deduction) & ", '" & IIf(mReposting, "Y", "N") & "', '" & pubUName & "', " & ConvertDate(IIf(TopCtrl1.TopText2 = "Browse", PubLoginDate, "")) & ", '" & Time & "', " & ConvertDate(IIf(TopCtrl1.TopText2 = "Edit", PubLoginDate, "")) & ", '" & Time & "')"
'    End If
'End Sub
'
'




Sub Ini_Pub()
    Dim RsTemp As ADODB.Recordset
    
    Set RsTemp = GCn.Execute("Select CheckNegetiveStockSiteWise From Syctrl")
    If RsTemp.RecordCount > 0 Then
        mCheckNegetiveStockSiteWise = VNull(RsTemp!CheckNegetiveStockSiteWise)
    End If
End Sub

Sub DispText_Vat()
    With FGrid
        If PubVATYN = 1 Then
            .TextMatrix(0, TaxPer) = "TaxPer"
            .ColAlignmentFixed(TaxPer) = flexAlignRightCenter
            .ColWidth(TaxPer) = 840
            
            .TextMatrix(0, TaxAmt1) = "TaxAmt"
            .ColAlignmentFixed(TaxAmt1) = flexAlignRightCenter
            .ColWidth(TaxAmt1) = 840
            
            If mSatYn Then
                .TextMatrix(0, SatPer) = "Sat %"
                .ColAlignmentFixed(SatPer) = flexAlignRightCenter
                .ColWidth(SatPer) = 840
                
                .TextMatrix(0, SatAmt1) = "Sat Amt"
                .ColAlignmentFixed(SatAmt1) = flexAlignRightCenter
                .ColWidth(SatAmt1) = 840
            Else
                .ColWidth(SatPer) = 0
                .ColWidth(SatAmt1) = 0
            End If
        Else
            .TextMatrix(0, TaxPer) = ""
            .ColAlignmentFixed(TaxPer) = flexAlignRightCenter
            .ColWidth(TaxPer) = 0
        
            .TextMatrix(0, TaxAmt1) = ""
            .ColAlignmentFixed(TaxAmt1) = flexAlignRightCenter
            .ColWidth(TaxAmt1) = 0
            
            .ColWidth(SatPer) = 0
            .ColWidth(SatAmt1) = 0
        End If
    
    End With
End Sub
