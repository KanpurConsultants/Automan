VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmSubGroup 
   BackColor       =   &H00B7D2D9&
   Caption         =   "Ledger Accounts  Entry"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   47
      Left            =   1650
      MaxLength       =   50
      TabIndex        =   29
      Top             =   6240
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   46
      Left            =   1515
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "Mr."
      Top             =   660
      Width           =   480
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   45
      Left            =   1650
      MaxLength       =   40
      TabIndex        =   28
      Top             =   5970
      Width           =   4140
   End
   Begin MSDataGridLib.DataGrid DGTDSCat 
      Height          =   3330
      Left            =   11340
      Negotiate       =   -1  'True
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Category Name"
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
            DividerStyle    =   3
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   44
      Left            =   2040
      TabIndex        =   136
      Top             =   2160
      Width           =   1410
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   43
      Left            =   2040
      TabIndex        =   135
      Top             =   1890
      Width           =   1410
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   8640
      TabIndex        =   126
      Top             =   6945
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   330
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   75
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
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1515
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1200
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2145
      MaxLength       =   40
      TabIndex        =   11
      Top             =   2730
      Width           =   3645
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   42
      Left            =   7470
      MaxLength       =   40
      TabIndex        =   48
      Top             =   6825
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   41
      Left            =   7470
      MaxLength       =   40
      TabIndex        =   47
      Top             =   6555
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   40
      Left            =   7470
      MaxLength       =   40
      TabIndex        =   46
      Top             =   6285
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   39
      Left            =   2220
      MaxLength       =   40
      TabIndex        =   45
      Top             =   6900
      Width           =   3570
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   38
      Left            =   1650
      MaxLength       =   4
      TabIndex        =   44
      Text            =   "Mr."
      Top             =   6900
      Width           =   525
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   1650
      TabIndex        =   21
      Top             =   5160
      Width           =   1410
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   4500
      TabIndex        =   22
      Top             =   5160
      Width           =   1290
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   4500
      TabIndex        =   27
      Top             =   5700
      Width           =   1290
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   7470
      MaxLength       =   35
      TabIndex        =   43
      Top             =   5925
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   36
      Left            =   10245
      MaxLength       =   6
      TabIndex        =   42
      Top             =   5655
      Width           =   1365
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   34
      Left            =   7470
      MaxLength       =   40
      TabIndex        =   40
      Top             =   5385
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   7470
      MaxLength       =   40
      TabIndex        =   39
      Top             =   5115
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   7470
      MaxLength       =   40
      TabIndex        =   38
      Top             =   4845
      Width           =   4140
   End
   Begin VB.TextBox Txt 
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   35
      Left            =   7470
      TabIndex        =   41
      Top             =   5655
      Width           =   2085
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   7470
      MaxLength       =   4
      TabIndex        =   36
      Text            =   "Mr."
      Top             =   4575
      Width           =   525
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   8040
      MaxLength       =   40
      TabIndex        =   37
      Top             =   4575
      Width           =   3570
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   7470
      TabIndex        =   34
      Top             =   3810
      Width           =   2790
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1650
      MaxLength       =   4
      TabIndex        =   10
      Text            =   "Mr."
      Top             =   2730
      Width           =   480
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   7470
      MaxLength       =   20
      TabIndex        =   33
      Top             =   3540
      Width           =   2790
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   2985
      Left            =   10290
      Negotiate       =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   4845
      Visible         =   0   'False
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   5265
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
      BeginProperty Column01 
         DataField       =   "Code"
         Caption         =   "City Code"
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
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGUnderAc 
      Height          =   3330
      Left            =   10335
      Negotiate       =   -1  'True
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   4425
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Group Name"
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
         Caption         =   "Group Code"
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
         DataField       =   "GroupNature"
         Caption         =   "GroupNature"
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
         DataField       =   "MainGrCode"
         Caption         =   "MainGrCode"
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
         DataField       =   "GroupLevel"
         Caption         =   "GroupLevel"
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
         DataField       =   "CurrentCount"
         Caption         =   "CurrentCount"
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
      BeginProperty Column06 
         DataField       =   "CurrentBalance"
         Caption         =   "CurrentBalance"
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
      BeginProperty Column07 
         DataField       =   "SubLedYN"
         Caption         =   "SubLedYN"
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
         DataField       =   "AliasYN"
         Caption         =   "AliasYN"
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
      BeginProperty Column09 
         DataField       =   "GroupHelp"
         Caption         =   "GroupHelp"
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
      BeginProperty Column10 
         DataField       =   "Nature"
         Caption         =   "Nature"
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
            ColumnWidth     =   5040
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
         BeginProperty Column04 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGAcAlias 
      Height          =   3330
      Left            =   10350
      Negotiate       =   -1  'True
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   4470
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
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
      ColumnCount     =   5
      BeginProperty Column00 
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
      BeginProperty Column02 
         DataField       =   "AliasYN"
         Caption         =   "AliasYN"
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
         DataField       =   "NameHelp"
         Caption         =   "NameHelp"
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
         DataField       =   "GroupCode"
         Caption         =   "Under Group"
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
            ColumnWidth     =   5040
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
         BeginProperty Column04 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   2805
      MaxLength       =   3
      TabIndex        =   26
      Text            =   "Yes"
      Top             =   5700
      Width           =   510
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   6750
      TabIndex        =   8
      Top             =   2055
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1515
      MaxLength       =   8
      TabIndex        =   1
      Top             =   390
      Width           =   1200
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   7470
      MaxLength       =   200
      TabIndex        =   35
      Top             =   4080
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1515
      MaxLength       =   40
      TabIndex        =   5
      Top             =   930
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.TextBox Txt 
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   1650
      TabIndex        =   15
      Top             =   3810
      Width           =   2085
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2010
      MaxLength       =   40
      TabIndex        =   3
      Top             =   660
      Width           =   3645
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   1650
      MaxLength       =   40
      TabIndex        =   12
      Top             =   3000
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1650
      MaxLength       =   40
      TabIndex        =   13
      Top             =   3270
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   1650
      MaxLength       =   40
      TabIndex        =   14
      Top             =   3540
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4425
      MaxLength       =   6
      TabIndex        =   16
      Top             =   3810
      Width           =   1365
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   1650
      MaxLength       =   35
      TabIndex        =   17
      Top             =   4080
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   1650
      MaxLength       =   24
      TabIndex        =   18
      Top             =   4350
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   1650
      MaxLength       =   24
      TabIndex        =   19
      Top             =   4620
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   1650
      MaxLength       =   50
      TabIndex        =   20
      Top             =   4890
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   4500
      TabIndex        =   24
      Top             =   5430
      Width           =   1290
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   1650
      TabIndex        =   23
      Top             =   5430
      Width           =   1410
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   840
      MaxLength       =   3
      TabIndex        =   25
      Text            =   "Yes"
      Top             =   5700
      Width           =   510
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   7470
      MaxLength       =   20
      TabIndex        =   32
      Top             =   3270
      Width           =   2790
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   7470
      MaxLength       =   30
      TabIndex        =   31
      Top             =   3000
      Width           =   2790
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   7470
      MaxLength       =   30
      TabIndex        =   30
      Top             =   2730
      Width           =   2790
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7470
      MaxLength       =   40
      TabIndex        =   6
      Top             =   690
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      HelpContextID   =   2
      Index           =   2
      Left            =   7470
      MaxLength       =   40
      TabIndex        =   4
      Top             =   420
      Visible         =   0   'False
      Width           =   4140
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1155
      Left            =   5940
      TabIndex        =   9
      Top             =   1455
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   2037
      _Version        =   393216
      BackColor       =   12243913
      Cols            =   8
      BackColorFixed  =   16703741
      ForeColorFixed  =   16512
      BackColorSel    =   15259902
      BackColorBkg    =   13623520
      GridColor       =   12632319
      GridColorFixed  =   16761024
      GridLinesFixed  =   1
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSDataGridLib.DataGrid DGAcName 
      Height          =   3330
      Left            =   10920
      Negotiate       =   -1  'True
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   3900
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
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
      ColumnCount     =   5
      BeginProperty Column00 
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
      BeginProperty Column02 
         DataField       =   "AliasYN"
         Caption         =   "AliasYN"
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
         DataField       =   "NameHelp"
         Caption         =   "NameHelp"
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
         DataField       =   "GroupCode"
         Caption         =   "Under Group"
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
            ColumnWidth     =   5040
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
         BeginProperty Column04 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGPartyType 
      Height          =   2520
      Left            =   8655
      Negotiate       =   -1  'True
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   5070
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4445
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
         Caption         =   "Party Type"
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
            DividerStyle    =   3
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   24
      Left            =   1500
      TabIndex        =   142
      Top             =   6240
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Conct . No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   90
      TabIndex        =   141
      Top             =   6240
      Width           =   1395
   End
   Begin VB.Label LblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Edit Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   4785
      TabIndex        =   140
      Top             =   405
      Width           =   2520
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   1290
      X2              =   5925
      Y1              =   6765
      Y2              =   6765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transporter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   90
      TabIndex        =   139
      Top             =   5970
      Width           =   1050
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   37
      Left            =   1500
      TabIndex        =   138
      Top             =   5970
      Width           =   150
   End
   Begin VB.Line Shape3 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   5970
      X2              =   11775
      Y1              =   6225
      Y2              =   6225
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      X1              =   8190
      X2              =   11790
      Y1              =   4470
      Y2              =   4470
   End
   Begin VB.Label LblHindi 
      AutoSize        =   -1  'True
      BackColor       =   &H00CAF1FD&
      Caption         =   "Hindi Section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   60
      TabIndex        =   133
      Top             =   6630
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   5955
      TabIndex        =   137
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   36
      Left            =   7320
      TabIndex        =   132
      Top             =   6285
      Width           =   150
   End
   Begin VB.Label LblAddBiLang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   5940
      TabIndex        =   131
      Top             =   6630
      Width           =   765
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   35
      Left            =   1500
      TabIndex        =   130
      Top             =   6900
      Width           =   150
   End
   Begin VB.Label LblConPerBiLang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   90
      TabIndex        =   129
      Top             =   6900
      Width           =   1365
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   34
      Left            =   1905
      TabIndex        =   125
      Top             =   1650
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   35
      Left            =   240
      TabIndex        =   124
      Top             =   1650
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   885
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   1590
      Width           =   4170
   End
   Begin VB.Label LblOpBalType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dr/Cr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3510
      TabIndex        =   123
      Top             =   1890
      Width           =   555
   End
   Begin VB.Label LblCurBalType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dr/Cr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3510
      TabIndex        =   122
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   32
      Left            =   240
      TabIndex        =   121
      Top             =   1890
      Width           =   1560
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   33
      Left            =   1905
      TabIndex        =   120
      Top             =   1890
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   31
      Left            =   240
      TabIndex        =   119
      Top             =   2160
      Width           =   1425
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   31
      Left            =   1905
      TabIndex        =   118
      Top             =   2160
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00CAECF0&
      Caption         =   "Other Temporary Details "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   29
      Left            =   5940
      TabIndex        =   117
      Top             =   4365
      Width           =   2250
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   30
      Left            =   90
      TabIndex        =   116
      Top             =   5160
      Width           =   990
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   32
      Left            =   1500
      TabIndex        =   115
      Top             =   5160
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local/Central"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   28
      Left            =   3135
      TabIndex        =   114
      Top             =   5160
      Width           =   1185
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   30
      Left            =   4380
      TabIndex        =   113
      Top             =   5160
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   27
      Left            =   3570
      TabIndex        =   112
      Top             =   5700
      Width           =   750
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   29
      Left            =   4380
      TabIndex        =   111
      Top             =   5700
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   26
      Left            =   5940
      TabIndex        =   110
      Top             =   5925
      Width           =   1125
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   28
      Left            =   7320
      TabIndex        =   109
      Top             =   4845
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   25
      Left            =   5940
      TabIndex        =   108
      Top             =   4845
      Width           =   765
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   27
      Left            =   7320
      TabIndex        =   107
      Top             =   5655
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   24
      Left            =   5955
      TabIndex        =   106
      Top             =   5655
      Width           =   330
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   26
      Left            =   10095
      TabIndex        =   105
      Top             =   5655
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   23
      Left            =   9690
      TabIndex        =   104
      Top             =   5655
      Width           =   330
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   25
      Left            =   7320
      TabIndex        =   103
      Top             =   5925
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father/Guardian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   22
      Left            =   5940
      TabIndex        =   102
      Top             =   4575
      Width           =   1455
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   21
      Left            =   5940
      TabIndex        =   101
      Top             =   3810
      Width           =   1290
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   23
      Left            =   7320
      TabIndex        =   100
      Top             =   3810
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IT Ward No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   20
      Left            =   5940
      TabIndex        =   99
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   22
      Left            =   7320
      TabIndex        =   98
      Top             =   3540
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt. Party"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   19
      Left            =   1665
      TabIndex        =   93
      Top             =   5700
      Width           =   975
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   21
      Left            =   2655
      TabIndex        =   92
      Top             =   5700
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   18
      Left            =   90
      TabIndex        =   91
      Top             =   390
      Width           =   495
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   20
      Left            =   1365
      TabIndex        =   90
      Top             =   390
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remark"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   17
      Left            =   5940
      TabIndex        =   89
      Top             =   4080
      Width           =   720
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   19
      Left            =   7320
      TabIndex        =   88
      Top             =   4080
      Width           =   150
   End
   Begin VB.Label LblNature 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2040
      TabIndex        =   87
      Top             =   1650
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   45
      X2              =   11820
      Y1              =   2655
      Y2              =   2655
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   18
      Left            =   4380
      TabIndex        =   86
      Top             =   5430
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   16
      Left            =   3270
      TabIndex        =   85
      Top             =   5430
      Width           =   1050
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   17
      Left            =   1500
      TabIndex        =   84
      Top             =   5430
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   15
      Left            =   90
      TabIndex        =   83
      Top             =   5430
      Width           =   975
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   16
      Left            =   690
      TabIndex        =   82
      Top             =   5700
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   14
      Left            =   90
      TabIndex        =   81
      Top             =   5700
      Width           =   555
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   15
      Left            =   7320
      TabIndex        =   80
      Top             =   3270
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAN No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   13
      Left            =   5940
      TabIndex        =   79
      Top             =   3270
      Width           =   780
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   14
      Left            =   7320
      TabIndex        =   78
      Top             =   3000
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LST No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   12
      Left            =   5940
      TabIndex        =   77
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   13
      Left            =   7320
      TabIndex        =   76
      Top             =   2730
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CST No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   11
      Left            =   5940
      TabIndex        =   75
      Top             =   2730
      Width           =   765
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   10
      Left            =   90
      TabIndex        =   74
      Top             =   4890
      Width           =   570
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   12
      Left            =   1500
      TabIndex        =   73
      Top             =   4890
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FAX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   9
      Left            =   90
      TabIndex        =   72
      Top             =   4620
      Width           =   375
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   11
      Left            =   1500
      TabIndex        =   71
      Top             =   4620
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   8
      Left            =   90
      TabIndex        =   70
      Top             =   4350
      Width           =   615
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   10
      Left            =   1500
      TabIndex        =   69
      Top             =   4350
      Width           =   150
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   9
      Left            =   1500
      TabIndex        =   68
      Top             =   4080
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   6
      Left            =   3870
      TabIndex        =   67
      Top             =   3810
      Width           =   330
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   8
      Left            =   4275
      TabIndex        =   66
      Top             =   3810
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   5
      Left            =   90
      TabIndex        =   65
      Top             =   3810
      Width           =   330
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   7
      Left            =   1500
      TabIndex        =   64
      Top             =   3810
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   4
      Left            =   90
      TabIndex        =   63
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   6
      Left            =   1500
      TabIndex        =   62
      Top             =   3000
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   61
      Top             =   2730
      Width           =   1365
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   5
      Left            =   1500
      TabIndex        =   60
      Top             =   2730
      Width           =   150
   End
   Begin VB.Label LblAliasBiLang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Hindi)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6480
      TabIndex        =   59
      Top             =   675
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   3
      Left            =   7320
      TabIndex        =   58
      Top             =   690
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label LblNameBiLang 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Hindi)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6480
      TabIndex        =   57
      Top             =   405
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   56
      Top             =   420
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   4
      Left            =   1365
      TabIndex        =   55
      Top             =   1200
      Width           =   150
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   2
      Left            =   1365
      TabIndex        =   54
      Top             =   930
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   0
      Left            =   1365
      TabIndex        =   53
      Top             =   660
      Width           =   150
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   52
      Top             =   930
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   7
      Left            =   90
      TabIndex        =   51
      Top             =   4080
      Width           =   1125
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Under Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   2
      Left            =   90
      TabIndex        =   50
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   49
      Top             =   660
      Width           =   555
   End
End
Attribute VB_Name = "frmSubGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Don't Change Tag Property of (Txt(AcName)) Control as it is used in other activities
Option Explicit
Public MasterFormExit As Boolean
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim ExitCtrl As Boolean
Dim SysGroup As String                      ' For System Group Tracking(Y-> System Group,N-> Non System Group)
Private Const mVType As String = "F_AO"
Dim VNo As Long
Dim VPrefixUpdateFlag As Byte               ' Used For Whether We've to Update counter of Voucher_Prefix Table or not
Dim mSearchCode As String, mDocId As String
Dim OldName As String
Dim OldAlias As String
Dim OldMainGrCode As String                 ' Old Under Group MainGrCode
Dim OldCurBal As Double                     ' Old Current Balance
Dim OldCurBalType As String                 ' Old Current Balance Type
Dim DetailFlag As Byte                      ' Used for Detail will have to feed or not
Dim AliasName As String                     ' For Alias
Dim Master As ADODB.Recordset
Dim RsAcName As ADODB.Recordset
Dim RsAcNameHelp As ADODB.Recordset
Dim RsAcAlias As ADODB.Recordset
Dim RsUnderAc As ADODB.Recordset
Dim Rscity As ADODB.Recordset
Dim RsTDSCat As ADODB.Recordset
Dim rsPartyType As ADODB.Recordset
Dim TmpSQL As String
Dim ListArray As Variant
Dim mListItem As ListItem
'************* Constant defined for useing in place of Index
Private Const SubCode = 0                   ' SubCode
Private Const AcName = 1                    ' Name
Private Const AcNameBiLang = 2              ' Name Bi Lang
Private Const AcAlias = 3                   ' Alias
Private Const AcAliasBiLang = 4             ' Alias Bi Lang
Private Const UnderGroup = 5                ' Under Group
Private Const ConPersonPrefix = 6           ' Contact Person Prefix
Private Const ConPerson = 7                 ' Contact Person
Private Const Add1 = 8                      ' Address1
Private Const Add2 = 9                      ' Address2
Private Const Add3 = 10                     ' Address3
Private Const City = 11                     ' City
Private Const Pin = 12                      ' PIN
Private Const Phone = 13                    ' Phone
Private Const Mobile = 14                   ' Mobile
Private Const FAx = 15                      ' FAX
Private Const EMail = 16                    ' EMail
Private Const Religion = 17                 ' Religion
Private Const PartyType = 18                ' Party Type
Private Const LC = 19                       ' Local/Central
Private Const CrLimit = 20                  ' Credit Limit
Private Const CrDays = 21                   ' Credit Days
Private Const ActiveYN = 22                 ' Active Y/N
Private Const GovtPartyYN = 23              ' Govt.Party Y/N
Private Const CST = 24                      ' CST No
Private Const LST = 25                      ' LST No
Private Const PAN = 26                      ' PAN No
Private Const ITWardNo = 27                 ' IT Ward No
Private Const TDSCat = 28                   ' TDS Cat.
Private Const Remark = 29                   ' Remark
Private Const ConPersonPrefixB = 30         ' Contact Person Prefix Bussiness
Private Const ConPersonB = 31               ' Contact Person Bussiness
Private Const Add1B = 32                    ' Address1 Bussiness
Private Const Add2B = 33                    ' Address2 Bussiness
Private Const Add3B = 34                    ' Address3 Bussiness
Private Const CityB = 35                    ' City Bussiness
Private Const PinB = 36                     ' PIN Bussiness
Private Const PhoneB = 37                   ' Phone Bussiness
Private Const ConPersonPrefixBiLang = 38    ' Contact Person Prefix Bi Lang
Private Const ConPersonBiLang = 39          ' Contact Person Bi Lang
Private Const Add1BiLang = 40               ' Address1 Bi Lang
Private Const Add2BiLang = 41               ' Address2 Bi Lang
Private Const Add3BiLang = 42               ' Address3 Bi Lang
Private Const OpBal = 43                    ' Opening Balance
Private Const CurBal = 44                   ' Current Balance
Private Const Transporter = 45              ' Transporter
Private Const NamePrefix = 46              ' Name Prefix
Private Const RCNo = 47                    'Rate Contact No
'************* Constant defined for useing in place of Index in Line File
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_RefDate = 1                   ' Ref Date
Private Const Col_RefNo = 2                     ' Ref No
Private Const Col_Amount = 3                    ' Amount
Private Const Col_CrDr = 4                      ' Cr/Dr
Private Const Col_VNo = 5                       ' Voucher No

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    For I = 0 To txt.Count - 1
        If I = OpBal Or I = CurBal Then
        Else
            txt(I).Enabled = Enb
        End If
    Next
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    Master.MoveFirst
    Master.FIND ("SearchCode='" & MyValue & "'")
    MoveRec
    BUTTONS True, Me, Master, 0
Exit Sub
ELoop:
    CheckError
End Sub
'To Make Controls Blank
Private Sub BlankText()
Dim I As Byte
    For I = 0 To txt.Count - 1
        txt(I) = ""
        txt(I).Tag = ""
    Next
    LblNature.CAPTION = ""
    LblOpBalType.CAPTION = ""
    LblCurBalType.CAPTION = ""
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End Sub
'This Function is used to change the position of control when Bi language is True or false
Private Sub ConvBiLanguage(Enb As Boolean)
    If Enb = True Then
        LblNameBiLang.CAPTION = "(" & BiLanguageName & ")"
        LblNameBiLang.Visible = True
        LblColon(1).Visible = True
        txt(AcNameBiLang).Font = BiLanguageFont
        txt(AcNameBiLang).Visible = True

        LblAliasBiLang.CAPTION = "(" & BiLanguageName & ")"
        LblAliasBiLang.Visible = True
        LblColon(3).Visible = True
        txt(AcAliasBiLang).Font = BiLanguageFont
        txt(AcAliasBiLang).Visible = True

        LblConPerBiLang.Visible = True
        LblColon(35).Visible = True
        txt(ConPersonPrefixBiLang).Font = BiLanguageFont
        txt(ConPersonPrefixBiLang).Visible = True
        txt(ConPersonBiLang).Font = BiLanguageFont
        txt(ConPersonBiLang).Visible = True

        LblAddBiLang.Visible = True
        LblColon(36).Visible = True
        txt(Add1BiLang).Font = BiLanguageFont
        txt(Add1BiLang).Visible = True
        txt(Add2BiLang).Font = BiLanguageFont
        txt(Add2BiLang).Visible = True
        txt(Add3BiLang).Font = BiLanguageFont
        txt(Add3BiLang).Visible = True
        Line3.Visible = True
        Shape3.Visible = True
    Else
        Shape3.Visible = False: LblHindi.Visible = False
        LblNameBiLang.Visible = False: LblColon(1).Visible = False: txt(AcNameBiLang).Visible = False
        LblAliasBiLang.Visible = False: LblColon(3).Visible = False: txt(AcAliasBiLang).Visible = False
        LblConPerBiLang.Visible = False: LblColon(35).Visible = False: txt(ConPersonPrefixBiLang).Visible = False: txt(ConPersonBiLang).Visible = False
        LblAddBiLang.Visible = False: LblColon(36).Visible = False: txt(Add1BiLang).Visible = False: txt(Add2BiLang).Visible = False: txt(Add3BiLang).Visible = False
    End If
End Sub

Private Sub Grid_Hide()
    If DGAcName.Visible = True Then DGAcName.Visible = False
    If DGAcAlias.Visible = True Then DGAcAlias.Visible = False
    If DGUnderAc.Visible = True Then DGUnderAc.Visible = False
    If DGCity.Visible = True Then DGCity.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGTDSCat.Visible = True Then DGTDSCat.Visible = False
    If DGPartyType.Visible = True Then DGPartyType.Visible = False
End Sub
'* Used for intialize grid columns
Private Sub Grid_Ini()
    With FGrid
        .left = 5940 ' Me.left + 90
'        .width = 6555
        .top = 1455 '1575
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 6

        .TextMatrix(0, Col_SrNo) = "S.No"
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 450

        .TextMatrix(0, Col_VNo) = "Voucher No"
        .ColAlignment(Col_VNo) = flexAlignLeftCenter
        .ColWidth(Col_VNo) = 0

        .TextMatrix(0, Col_RefDate) = "Ref. Date"
        .ColAlignment(Col_RefDate) = flexAlignLeftCenter
        .ColWidth(Col_RefDate) = 1500

        .TextMatrix(0, Col_RefNo) = "Ref. No"
        .ColAlignment(Col_RefNo) = flexAlignLeftCenter
        .ColWidth(Col_RefNo) = 1600

        .TextMatrix(0, Col_Amount) = "Amount"
        .ColAlignmentFixed(Col_Amount) = flexAlignRightCenter
        .ColWidth(Col_Amount) = 1200

        .TextMatrix(0, Col_CrDr) = "Cr/Dr"
        .ColAlignment(Col_CrDr) = flexAlignLeftCenter
        .ColWidth(Col_CrDr) = 600
    End With
'    DGAcName.left = txt(AcName).left: DGAcName.top = txt(AcName).top + txt(AcName).height + 15
'    DGAcAlias.left = txt(AcAlias).left: DGAcAlias.top = txt(AcAlias).top + txt(AcAlias).height + 15
'    DGUnderAc.left = txt(UnderGroup).left: DGUnderAc.top = txt(UnderGroup).top + txt(UnderGroup).height + 15
'    DGCity.left = txt(City).left: DGCity.top = txt(City).top + txt(City).height + 15
'    DGTDSCat.left = txt(TDSCat).left: DGTDSCat.top = txt(TDSCat).top + txt(TDSCat).height + 15
'    DGPartyType.left = txt(PartyType).left: DGPartyType.top = txt(PartyType).top + txt(PartyType).height + 15
    DGAcName.left = 6105: DGAcName.top = mTopScale
    DGAcAlias.left = 6105: DGAcAlias.top = mTopScale
    DGUnderAc.left = 6105: DGUnderAc.top = mTopScale
    DGCity.left = 7635: DGCity.top = mTopScale
    DGTDSCat.left = 8055: DGTDSCat.top = mTopScale: DGTDSCat.height = 2200
    DGPartyType.left = Me.width - (DGPartyType.width + mRtScale): DGPartyType.top = mTopScale
End Sub

Private Sub DetailEnb(DetFlag As Byte)
Dim I As Byte
    If DetFlag = 1 Then
        For I = 6 To 42
            txt(I).Enabled = True
        Next
    Else
        For I = 6 To 42
            txt(I).Enabled = False
            txt(I).TEXT = ""
        Next
    End If
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid.Col
    Case Col_RefDate
        If txtGrid(0) <> "" Then
            If RetDate(txtGrid(0)) >= PubStartDate Then
                MsgBox "Date should be before financial year", vbCritical, "Automan"
                TxtGridLeave = False
                Exit Function
            Else
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(txtGrid(0))
            End If
        End If
    Case Col_RefNo ', Col_CrDr
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = txtGrid(0).TEXT
        If FGrid.Col = Col_CrDr Then
            CalOpBal
            If FGrid.TextMatrix(FGrid.Rows - 1, Col_CrDr) <> "" And Val(FGrid.TextMatrix(FGrid.Rows - 1, Col_Amount)) <> 0 Then FGrid.AddItem FGrid.Rows
        End If
    Case Col_Amount
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(txtGrid(0).TEXT), "0.00")
        CalOpBal
        If FGrid.TextMatrix(FGrid.Rows - 1, Col_CrDr) <> "" And Val(FGrid.TextMatrix(FGrid.Rows - 1, Col_Amount)) <> 0 Then FGrid.AddItem FGrid.Rows
End Select
ExitCtrl = True
TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    txtGrid(0).Visible = False
End If
End Function

Private Sub CalOpBal()
Dim I As Integer, cr As Double, dr As Double, FinalCrDr As Double
    cr = 0: dr = 0: FinalCrDr = 0
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_CrDr) = "Cr" Then
            cr = cr + Val(FGrid.TextMatrix(I, Col_Amount))
        ElseIf FGrid.TextMatrix(I, Col_CrDr) = "Dr" Then
            dr = dr + Val(FGrid.TextMatrix(I, Col_Amount))
        End If
    Next
    FinalCrDr = Format(cr - dr, "0.00")
    txt(OpBal) = Format(Abs(FinalCrDr), "0.00")
    If FinalCrDr > 0 Then
        LblOpBalType.CAPTION = "Cr"
    Else
        LblOpBalType.CAPTION = "Dr"
    End If
    
    txt(CurBal) = Format(Abs(Val(txt(CurBal).Tag) + FinalCrDr), "0.00")
    If Val(txt(CurBal).Tag) + FinalCrDr > 0 Then
        LblCurBalType.CAPTION = "Cr"
    Else
        LblCurBalType.CAPTION = "Dr"
    End If
End Sub

Private Sub MoveRec()
'On Error GoTo ELoop
Dim Rst As ADODB.Recordset, CrAmt As Double, DrAmt As Double
Dim CRDR As String, CrDrAmt As Double, I As Integer, SiteType$
Dim mTAdd As Boolean, mTDel As Boolean
    If Master.RecordCount > 0 Then
       Set Rst = G_FaCn.Execute("Select S.*,G.GroupName,G.MainGrCode,C.CityName,C1.CityName As CityNameB,T.TDS_Desc,ST.Description As PartyType From ((((SubGroup S Left Join AcGroup G on S.GroupCode=G.GroupCode) " _
            & "Left Join City C on S.CityCode=C.CityCode) Left Join City C1 on S.TCityCode=C1.CityCode) " _
            & "Left Join TDSCat T on S.TDS_Catg=T.TDS_Catg) Left Join SubGroupType ST on S.Party_Type=ST.Party_Type Where s.SubCode='" & Master!SearchCode & "'")
        txt(SubCode).TEXT = Rst!SubCode
        If Rst!Name = "Cash" Then
            SysGroup = "Y"
        Else
            SysGroup = "N"
        End If
        txt(SubCode) = Rst!SubCode
        txt(AcName) = Rst!Name
        txt(NamePrefix) = IIf(IsNull(Rst!NamePrefix), "", Rst!NamePrefix)
        OldName = txt(AcName)
        txt(AcNameBiLang) = Rst!NameBiLang
        txt(UnderGroup) = IIf(IsNull(Rst!GroupName), "", Rst!GroupName)
        txt(UnderGroup).Tag = Rst!GroupCode
        OldMainGrCode = Rst!MainGrCode
        LblNature.CAPTION = XNull(Rst!Nature)
        If Rst!Nature = "Customer" Or Rst!Nature = "Supplier" Then
            DetailFlag = 1
        Else
            DetailFlag = 0
        End If
        CrAmt = G_FaCn.Execute("Select iif(isnull(Sum(AmtCr)),0,Sum(AmtCr)) From Ledger Where V_Type='" & mVType & "' and SubCode='" & txt(SubCode) & "'").Fields(0).Value
        DrAmt = G_FaCn.Execute("Select iif(isnull(Sum(AmtDr)),0,Sum(AmtDr)) From Ledger Where V_Type='" & mVType & "' and SubCode='" & txt(SubCode) & "'").Fields(0).Value
        txt(OpBal).TEXT = Format(Abs(CrAmt - DrAmt), "0.00")
        LblOpBalType.CAPTION = IIf(CrAmt > DrAmt, "Cr", IIf(CrAmt < DrAmt, "Dr", ""))
        txt(CurBal).TEXT = Format(IIf(IsNull(Rst!Curr_Bal), 0, Abs(Rst!Curr_Bal)), "0.00")
        LblCurBalType.CAPTION = IIf(Rst!Curr_Bal > 0, "Cr", IIf(Rst!Curr_Bal < 0, "Dr", ""))
        OldCurBal = VNull(Rst!Curr_Bal)
        OldCurBalType = IIf(Rst!Curr_Bal > 0, "Cr", "Dr")
        txt(CurBal).Tag = VNull(Rst!Curr_Bal) - (CrAmt - DrAmt)
        txt(ConPersonPrefix) = IIf(IsNull(Rst!ConPrefix), "", Rst!ConPrefix)
        txt(ConPerson) = IIf(IsNull(Rst!ConPerson), "", Rst!ConPerson)
        txt(Add1) = IIf(IsNull(Rst!Add1), "", Rst!Add1)
        txt(Add2) = IIf(IsNull(Rst!Add2), "", Rst!Add2)
        txt(Add3) = IIf(IsNull(Rst!Add3), "", Rst!Add3)
        txt(City) = IIf(IsNull(Rst!CityName), "", Rst!CityName)
        txt(City).Tag = IIf(IsNull(Rst!CityCode), "", Rst!CityCode)
        txt(Pin) = IIf(IsNull(Rst!Pin), "", Rst!Pin)
        txt(Phone) = IIf(IsNull(Rst!Phone), "", Rst!Phone)
        txt(Mobile) = IIf(IsNull(Rst!Mobile), "", Rst!Mobile)
        txt(FAx) = IIf(IsNull(Rst!FAx), "", Rst!FAx)
        txt(EMail) = IIf(IsNull(Rst!EMail), "", Rst!EMail)
        If Rst!Religion = 0 Or IsNull(Rst!Religion) Then
            txt(Religion) = "N/A"
        ElseIf Rst!Religion = 1 Then
            txt(Religion) = "Hindu"
        ElseIf Rst!Religion = 2 Then
            txt(Religion) = "Muslim"
        ElseIf Rst!Religion = 3 Then
            txt(Religion) = "Sikh"
        ElseIf Rst!Religion = 4 Then
            txt(Religion) = "Christian"
        End If
        txt(PartyType) = IIf(IsNull(Rst!PartyType), "", Rst!PartyType)
        txt(PartyType).Tag = IIf(IsNull(Rst!Party_Type), 0, Rst!Party_Type)
        txt(LC).TEXT = IIf(Rst!L_C = "L", "Local", "Central")
        txt(CrLimit) = IIf(IsNull(Rst!CreditLimit), "", Rst!CreditLimit)
        txt(CrDays) = IIf(IsNull(Rst!CreditDays), "", Rst!CreditDays)
        txt(ActiveYN) = IIf(Rst!ActiveYN = 0, "No", "Yes")
        txt(GovtPartyYN) = IIf(Rst!Govt_YN = 0, "No", "Yes")
        txt(CST) = IIf(IsNull(Rst!CstNo), "", Rst!CstNo)
        txt(LST) = IIf(IsNull(Rst!LstNo), "", Rst!LstNo)
        txt(PAN) = IIf(IsNull(Rst!PanNo), "", Rst!PanNo)
        txt(ITWardNo) = IIf(IsNull(Rst!ITWARD_NO), "", Rst!ITWARD_NO)
        txt(TDSCat).TEXT = IIf(IsNull(Rst!TDS_Desc), "", Rst!TDS_Desc)
        txt(TDSCat).Tag = IIf(IsNull(Rst!TDS_Catg), "", Rst!TDS_Catg)
        txt(Remark) = IIf(IsNull(Rst!Remark), "", Rst!Remark)

        txt(ConPersonPrefixB) = IIf(IsNull(Rst!FPrefix), "", Rst!FPrefix)
        txt(ConPersonB) = IIf(IsNull(Rst!fname), "", Rst!fname)
        txt(Add1B) = IIf(IsNull(Rst!TAdd1), "", Rst!TAdd1)
        txt(Add2B) = IIf(IsNull(Rst!TAdd2), "", Rst!TAdd2)
        txt(Add3B) = IIf(IsNull(Rst!TAdd3), "", Rst!TAdd3)
        txt(CityB) = IIf(IsNull(Rst!CityNameB), "", Rst!CityNameB)
        txt(CityB).Tag = IIf(IsNull(Rst!TCityCode), "", Rst!TCityCode)
        txt(PinB) = IIf(IsNull(Rst!TPin), "", Rst!TPin)
        txt(PhoneB) = IIf(IsNull(Rst!TPhone), "", Rst!TPhone)
        txt(Transporter) = IIf(IsNull(Rst!Transporter), "", Rst!Transporter)
        txt(RCNo) = IIf(IsNull(Rst!RC_No), "", Rst!RC_No)

        txt(ConPersonPrefixBiLang) = IIf(IsNull(Rst!ConPrefixBiLang), "", Rst!ConPrefixBiLang)
        txt(ConPersonBiLang) = IIf(IsNull(Rst!ConPersonBiLang), "", Rst!ConPersonBiLang)
        txt(Add1BiLang) = IIf(IsNull(Rst!Add1BiLang), "", Rst!Add1BiLang)
        txt(Add2BiLang) = IIf(IsNull(Rst!Add2BiLang), "", Rst!Add2BiLang)
        txt(Add3BiLang) = IIf(IsNull(Rst!Add3BiLang), "", Rst!Add3BiLang)
        ' For Alias
        Set Rst = G_FaCn.Execute("Select SubCode,Name,NameBiLang,AliasYN From SubGroupAlias Where SubCode='" & txt(SubCode) & "' and AliasYN='Y'")
        If Rst.RecordCount > 0 Then
            txt(AcAlias) = Rst!Name
            OldAlias = Rst!Name
            txt(AcAliasBiLang) = IIf(IsNull(Rst!NameBiLang), "", Rst!NameBiLang)
        Else
            txt(AcAlias) = "": OldAlias = ""
            txt(AcAliasBiLang) = ""
        End If

        FGrid.Rows = 1
        Set Rst = G_FaCn.Execute("Select DocID,V_SNo,V_Type,V_No,V_Date,Site_Code,SubCode,iif(isnull(AmtCr),0,AmtCr) As AmtCr1,iif(isnull(AmtDr),0,AmtDr) As AmtDr1,Chq_No,Chq_Date From Ledger Where V_Type='" & mVType & "' and SubCode='" & txt(SubCode) & "' Order by V_SNo")
        If Rst.RecordCount > 0 Then
            I = 1
            mDocId = Rst!DocId
            VNo = Rst!V_NO
            VPrefixUpdateFlag = 1
            Do Until Rst.EOF
                If Rst!AmtCr1 = 0 Then
                    CRDR = "Dr"
                    CrDrAmt = Rst!AmtDr1
                ElseIf Rst!AmtDr1 = 0 Then
                    CRDR = "Cr"
                    CrDrAmt = Rst!AmtCr1
                End If
                FGrid.AddItem I & Chr(9) & Format(Rst!V_DATE, "dd/MMM/yyyy") & Chr(9) & Rst!Chq_No & Chr(9) & Format(CrDrAmt, "0.00") & Chr(9) & CRDR & Chr(9) & Rst!V_NO
                             '0                                1                                2                         3                         4                  5
                I = I + 1
                Rst.MoveNext
            Loop
            FGrid.FixedRows = 1
        Else
            mDocId = ""
            VNo = 0
            VPrefixUpdateFlag = 0
            FGrid.AddItem FGrid.Rows
            FGrid.FixedRows = 1
        End If
    Else
        BlankText
    End If
    Set Rst = Nothing
'    SiteType = GCn.Execute("Select SiteType from Site where Site_Code='" & PubSiteCode & "'").Fields(0).Value
'    If InStr(Me.TopCtrl1.Tag, "A") <> 0 Then mTAdd = True
'    If InStr(Me.TopCtrl1.Tag, "D") <> 0 Then mTDel = True
'    If SiteType = "H" Then  'HO
'        TopCtrl1.tAdd = mTAdd
'        TopCtrl1.tDel = mTAdd
'        LblStatus = "Add/Delete Enabled"
'    Else
'        TopCtrl1.tAdd = False
'        TopCtrl1.tDel = False
'        LblStatus = "Add/Delete disabled"
'    End If
'TopCtrl1.tDel = False ' At Cuttack
Exit Sub
'ELoop:
  '  CheckError
End Sub

Private Function GetReligion() As Byte
    If txt(Religion) = "N/A" Then
        GetReligion = 0
    ElseIf txt(Religion) = "Hindu" Then
        GetReligion = 1
    ElseIf txt(Religion) = "Muslim" Then
        GetReligion = 2
    ElseIf txt(Religion) = "Sikh" Then
        GetReligion = 3
    ElseIf txt(Religion) = "Christian" Then
        GetReligion = 4
    End If
End Function
'* Used for Generate Voucher No.
Private Function VoucherNo() As String
Dim Rst As ADODB.Recordset, VType As String, mVPrefix As String
    VType = "F_AO"
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & VType & "' Order By VP.Date_From DESC", G_FaCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount > 0 Then
        VNo = Rst!start_srl_no + 1
        mVPrefix = Rst!Prefix
    End If
    mDocId = PubDivCode + PubSiteCode + PubSiteCode + Space(5 - Len(CStr(VType))) + VType + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(VNo))) + CStr(VNo)
    VoucherNo = mDocId
Set Rst = Nothing
End Function

Private Function SubGroupUpdate(ByRef xType As String, ByRef xTableName As String, ByRef xAcID As String, ByRef xSubCode As String, ByRef xSubName As String, ByRef xSubNameBiLang As String, ByRef xAliasYN As String, xA_E As String) As String
Dim Nature As String, GroupNature As String
Dim MyCurrBal As Double, xNature$
'    xNature = G_FACN.Execute("Select IsNull(Nature,'',Nature) N From AcGroup Where GroupCode='" & Txt(UnderGroup).Tag & "'").Fields(0).Value
'
'    '), "Other", G_FACN.Execute("Select Nature From AcGroup Where GroupCode='" & Txt(UnderGroup).Tag & "'").Fields(0).Value)

    Nature = IIf(IsNull(G_FaCn.Execute("Select Nature From AcGroup Where GroupCode='" & txt(UnderGroup).Tag & "'").Fields(0).Value), "Other", G_FaCn.Execute("Select Nature From AcGroup Where GroupCode='" & txt(UnderGroup).Tag & "'").Fields(0).Value)
    GroupNature = G_FaCn.Execute("Select GroupNature From AcGroup Where GroupCode='" & txt(UnderGroup).Tag & "'").Fields(0).Value
    If LblCurBalType = "Cr" Then
        MyCurrBal = Val(txt(CurBal).TEXT) - OldCurBal
    Else
        MyCurrBal = -Val(txt(CurBal).TEXT) - OldCurBal
    End If
    If xType = "Add" Then
        SubGroupUpdate = "Insert Into " & xTableName & "(" _
            & "AcID,Site_Code,SubCode,FirmCode,NamePrefix,Name,NameBiLang," _
            & "NameHelp,GroupCode,GroupNature,Nature,AliasYN,ConPrefix," _
            & "ConPerson,Add1,Add2,Add3,CityCode," _
            & "PIN,Phone,Mobile,Fax,EMail," _
            & "Religion,Party_Type,L_C,CreditLimit,CreditDays," _
            & "ActiveYN,Govt_YN,CSTNo,LSTNo,PANNo," _
            & "ITWard_No,TDS_Catg,Remark,FPrefix,FName," _
            & "TAdd1,TAdd2,TAdd3,TCityCode,TPIN," _
            & "TPhone,ConPrefixBiLang,ConPersonBiLang,Add1BiLang,Add2BiLang," _
            & "Add3BiLang,Curr_Bal,Transporter,U_Name,U_EntDt,U_AE,RC_No) " _
            & "Values ('" & xAcID & "','" & PubSiteCode & "','" & xSubCode & "','" & PubFirmCode & "','" & txt(NamePrefix) & "','" & xSubName & "','" & xSubNameBiLang & _
            "','" & FilterString(xSubName) & "','" & txt(UnderGroup).Tag & "','" & GroupNature & "','" & Nature & "','" & xAliasYN & "','" & txt(ConPersonPrefix) & _
            "','" & txt(ConPerson) & "','" & txt(Add1) & "','" & txt(Add2) & "','" & txt(Add3) & "','" & txt(City).Tag & _
            "','" & txt(Pin) & "','" & txt(Phone) & "','" & txt(Mobile) & "','" & txt(FAx) & "','" & txt(EMail) & _
            "', " & GetReligion & "," & Val(txt(PartyType).Tag) & ",'" & IIf(txt(LC) = "Local", "L", "C") & "'," & Val(txt(CrLimit)) & "," & Val(txt(CrDays)) & _
            " , " & IIf(txt(ActiveYN) = "Yes", 1, 0) & "," & IIf(txt(GovtPartyYN) = "Yes", 1, 0) & ",'" & txt(CST) & "','" & txt(LST) & "','" & txt(PAN) & _
            "','" & txt(ITWardNo) & "','" & txt(TDSCat).Tag & "','" & txt(Remark) & "','" & txt(ConPersonPrefixB) & "','" & txt(ConPersonB) & _
            "','" & txt(Add1B) & "','" & txt(Add2B) & "','" & txt(Add3B) & "','" & txt(CityB).Tag & "','" & txt(PinB) & _
            "','" & txt(PhoneB) & "','" & txt(ConPersonPrefixBiLang) & "','" & txt(ConPersonBiLang) & "','" & txt(Add1BiLang) & "','" & txt(Add2BiLang) & _
            "','" & txt(Add3BiLang) & "'," & MyCurrBal & ",'" & txt(Transporter) & "','" & pubUName & "',#" & PubServerDate & "#,'" & xA_E & "','" & txt(RCNo) & "')"
    ElseIf xType = "Edit" Then
        '* For Getting New Values
        SubGroupUpdate = "Update " & xTableName & " Set " _
            & "NamePrefix='" & txt(NamePrefix) & "',Name='" & xSubName & "',NameBiLang='" & xSubNameBiLang & "'," _
            & "NameHelp='" & FilterString(xSubName) & "',GroupCode='" & txt(UnderGroup).Tag & "',GroupNature='" & GroupNature & "'," _
            & "Nature='" & Nature & "',AliasYN='" & xAliasYN & "'," _
            & "ConPrefix='" & txt(ConPersonPrefix) & "',ConPerson='" & txt(ConPerson) & "'," _
            & "Add1='" & txt(Add1) & "',Add2='" & txt(Add2) & "'," _
            & "Add3='" & txt(Add3) & "',CityCode='" & txt(City).Tag & "'," _
            & "PIN='" & txt(Pin) & "',Phone='" & txt(Phone) & "'," _
            & "Mobile='" & txt(Mobile) & "',Fax='" & txt(FAx) & "'," _
            & "EMail='" & txt(EMail) & "',Religion=" & GetReligion & "," _
            & "Party_Type=" & Val(txt(PartyType).Tag) & ",L_C='" & IIf(txt(LC) = "Local", "L", "C") & "'," _
            & "CreditLimit=" & Val(txt(CrLimit)) & ",CreditDays=" & Val(txt(CrDays)) & "," _
            & "ActiveYN=" & IIf(txt(ActiveYN) = "Yes", 1, 0) & ",Govt_YN=" & IIf(txt(GovtPartyYN) = "Yes", 1, 0) & "," _
            & "CSTNo='" & txt(CST) & "',LSTNo='" & txt(LST) & "'," _
            & "PANNo='" & txt(PAN) & "',ITWard_No='" & txt(ITWardNo) & "'," _
            & "TDS_Catg='" & txt(TDSCat).Tag & "',Remark='" & txt(Remark) & "'," _
            & "FPrefix='" & txt(ConPersonPrefixB) & "',FName='" & txt(ConPersonB) & "'," _
            & "TAdd1='" & txt(Add1B) & "',TAdd2='" & txt(Add2B) & "'," _
            & "TAdd3='" & txt(Add3B) & "',TCityCode='" & txt(CityB).Tag & "'," _
            & "TPIN='" & txt(PinB) & "',TPhone='" & txt(PhoneB) & "'," _
            & "ConPrefixBiLang='" & txt(ConPersonPrefixBiLang) & "',ConPersonBiLang='" & txt(ConPersonBiLang) & "'," _
            & "Add1BiLang='" & txt(Add1BiLang) & "',Add2BiLang='" & txt(Add2BiLang) & "'," _
            & "Add3BiLang='" & txt(Add3BiLang) & "',Transporter='" & txt(Transporter) & "',Curr_Bal=curr_bal+" & MyCurrBal & "," _
            & "U_Name='" & pubUName & "',U_EntDt=#" & Format(Now, "dd/MMM/yyyy HH:NN:SS") & "#,U_AE='" & xA_E & "',RC_No='" & txt(RCNo) & "'" _
            & "Where AcID='" & xAcID & "'"
    Else
        SubGroupUpdate = ""
    End If
End Function

'Database updation procedure For Addition
Private Sub UpdateDataBaseAdd()
On Error GoTo ELoop
Dim AcCodeAlias As String * 8   ' AcCode Alias
Dim ID As Integer               ' Code From SubGroupCounter Table
Dim mTrans As Boolean, iCount As Integer
'/////
Dim Rst As ADODB.Recordset, I As Integer, AmtCrDr As String
'/////
'* Database Updation
    If CodeEditFlag = True Then
        Set Rst = G_FaCn.Execute("Select SubCode From SubGroupAlias Where SubCode='" & txt(SubCode) & "'")
        If Rst.RecordCount > 0 Then
            MsgBox "Code Already Exists", vbInformation, "Validation"
            GSQL = "Select SubGroupAcCode From SubGroupCounter"
            'PubFirmCode is included in SubCode, Single Character
            txt(SubCode) = PubSiteCode & IIf(PubFirmCode = "", "0", PubFirmCode) & Format(G_FaCn.Execute(GSQL).Fields(0).Value, "000000")
            txt(SubCode).Tag = txt(SubCode)
            txt(SubCode).SetFocus
            Exit Sub
        End If
    End If
    G_FaCn.BeginTrans
    GCn.BeginTrans
'    For iCount = 1 To 1000 by lps for auto insert
        mTrans = True
'        ID = G_FACN.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0).Value
' by lps for auto insert
'        txt(SubCode) = left(txt(SubCode), 2) & Right("000000" & Val(Mid(txt(SubCode), 3, 6)) + 1, 6)
'        txt(AcName) = left(txt(AcName), 4) & iCount
        TmpSQL = SubGroupUpdate("Add", "SubGroup", txt(SubCode), txt(SubCode), txt(AcName), txt(AcNameBiLang), "N", "A")
        G_FaCn.Execute (TmpSQL)
        GCn.Execute (TmpSQL)
        TmpSQL = SubGroupUpdate("Add", "SubGroupAlias", txt(SubCode), txt(SubCode), txt(AcName), txt(AcNameBiLang), "N", "A")
        G_FaCn.Execute (TmpSQL)
        GCn.Execute (TmpSQL)
        ' Update Current Balance of Group
        If Val(txt(CurBal)) <> 0 Then
            CalBalAcGroup "SubGroup", G_FaCn, RsUnderAc!MainGrCode, Val(txt(CurBal)), IIf(LblCurBalType.CAPTION = "Cr", "+", "-")
        End If
        ' update SubGroupCounter table
        G_FaCn.Execute ("Update SubGroupCounter Set SubGroupAcCode=SubGroupAcCode+1")
        ' Used For Alias of Group
        If Trim(txt(AcAlias)) <> "" Then
            ID = G_FaCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0).Value
            AcCodeAlias = PubSiteCode & IIf(PubFirmCode = "", "0", PubFirmCode) & Format(CStr(ID), "000000")

            TmpSQL = SubGroupUpdate("Add", "SubGroupAlias", AcCodeAlias, txt(SubCode), txt(AcAlias), txt(AcAliasBiLang), "Y", "A")
            G_FaCn.Execute (TmpSQL)
            GCn.Execute (TmpSQL)

            G_FaCn.Execute ("Update SubGroupCounter Set SubGroupAcCode=SubGroupAcCode+1")
        End If
        '(Insert into line file)
        Dim AmtCr As Double, AmtDr As Double, mID As Double
        If Val(FGrid.TextMatrix(1, Col_Amount)) <> 0 Then
            mDocId = VoucherNo
            For I = 1 To FGrid.Rows - 1
                If Val(FGrid.TextMatrix(I, Col_Amount)) <> 0 Then
                    VPrefixUpdateFlag = 0
                    AmtCrDr = IIf(FGrid.TextMatrix(I, Col_CrDr) = "Cr", "AmtCr", "AmtDr")
                    G_FaCn.Execute "Delete From Ledger where DocId='" & mDocId & "'"
                    G_FaCn.Execute "Delete From LedgerRef where DocId='" & mDocId & "'"
                    G_FaCn.Execute "Insert Into Ledger(" _
                        & "DocId,V_SNo,V_Type,V_No,Site_Code," _
                        & "V_Date,SubCode," & AmtCrDr & ",Chq_No,Chq_Date," _
                        & "Narration,U_Name,U_EntDt,U_AE) " _
                        & "Values(" _
                        & "'" & mDocId & "'," & I & ",'" & mVType & "'," & VNo & ",'" & PubSiteCode & PubSiteCode & "'," _
                        & "" & ConvertDate(IIf(FGrid.TextMatrix(I, Col_RefDate) = "", PubStartDate - 1, FGrid.TextMatrix(I, Col_RefDate))) & ",'" & txt(SubCode) & "'," & Val(FGrid.TextMatrix(I, Col_Amount)) & ",'" & FGrid.TextMatrix(I, Col_RefNo) & "'," & ConvertDate(FGrid.TextMatrix(I, Col_RefDate)) & "," _
                        & "'Party Opening','" & pubUName & "',#" & PubServerDate & "#,'A')"
                End If
                AmtCr = IIf(AmtCrDr = "AmtCr", Val(FGrid.TextMatrix(I, Col_Amount)), 0)
                AmtDr = IIf(AmtCrDr = "AmtDr", Val(FGrid.TextMatrix(I, Col_Amount)), 0)
                mID = IIf(IsNull(G_FaCn.Execute("select Max(ID) from LedgerRef").Fields(0).Value), 0, G_FaCn.Execute("select Max(ID) from LedgerRef").Fields(0).Value) + 1
                GSQL = "INSERT INTO LedgerRef (ID,DOCID,V_SNO,DR,CR,SUBCODE,U_Name,U_EntDt,U_AE,AgRefType,AgRefNo,DueDate,V_dATE)" & _
                       "VALUES (" & mID & ",'" & mDocId & "'," & I & "," & AmtDr & "," & AmtCr & ",'" & txt(SubCode) & "','" & pubUName & "'," & FaConvertDate(Now) & ",'E','New Ref','" & FGrid.TextMatrix(I, Col_RefNo) & "'," & FaConvertDate(Now) & "," & ConvertDate(IIf(FGrid.TextMatrix(I, Col_RefDate) = "", PubStartDate - 1, FGrid.TextMatrix(I, Col_RefDate))) & ")"
                G_FaCn.Execute GSQL
            Next
           
            ' To Update Voucher_Prefix Serial No
            If VPrefixUpdateFlag = 0 Then
                If FGrid.Rows >= 2 Then
                    Set Rst = New ADODB.Recordset
                    Rst.CursorLocation = adUseClient
                    Rst.Open "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & mVType & "'", G_FaCn, adOpenDynamic, adLockOptimistic
                    If Rst.RecordCount > 0 Then
                        G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=Start_Srl_No+1 Where V_Type='" & Rst!V_tYPE & "'"
                    End If
                End If
            End If
        End If
'    Next by lps for auto insert
    G_FaCn.CommitTrans
    GCn.CommitTrans
    mTrans = False
    mSearchCode = txt(SubCode)
    RsAcName.Requery
    RsAcAlias.Requery
    Master.Requery
'    TopCtrl1_eRef

    Master.FIND "SearchCode = '" & mSearchCode & "'"
    TopCtrl1_eAdd
Set Rst = Nothing
Exit Sub
ELoop:
   If mTrans = True Then G_FaCn.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub
'Database updation procedure For Edit
Private Sub UpdateDataBaseEdit()
'On Error GoTo ELoop
Dim Rst As ADODB.Recordset, mTrans As Boolean
Dim I As Byte, ID As Integer
Dim AcID As String, AcIDAlias As String
Dim NewMainGrCode As String
'/////
Dim AmtCrDr As String
'/////
    '* For Getting New Values
    NewMainGrCode = G_FaCn.Execute("Select MainGrCode From AcGroup Where GroupCode='" & txt(UnderGroup).Tag & "'").Fields(0).Value
    G_FaCn.BeginTrans
    GCn.BeginTrans
        mTrans = True
' Ledger Updation
        G_FaCn.Execute ("Delete From Ledger Where DocId='" & mDocId & "' and V_Type='" & mVType & "' and SubCode= '" & txt(SubCode) & "'")
'       G_FaCn.Execute ("Delete From Ledger Where V_Type='" & mVType & "' and SubCode= '" & txt(SubCode) & "'")
        TmpSQL = SubGroupUpdate("Edit", "SubGroup", txt(SubCode), txt(SubCode), txt(AcName), txt(AcNameBiLang), "N", "E")
        G_FaCn.Execute (TmpSQL)
        GCn.Execute (TmpSQL)
        TmpSQL = SubGroupUpdate("Edit", "SubGroupAlias", txt(SubCode), txt(SubCode), txt(AcName), txt(AcNameBiLang), "N", "E")
        G_FaCn.Execute (TmpSQL)
        GCn.Execute (TmpSQL)
        ' Update Current Balance of Group
        If (OldMainGrCode <> NewMainGrCode Or OldCurBal <> Val(txt(CurBal).TEXT) Or OldCurBalType <> LblCurBalType.CAPTION) Then
            CalBalAcGroup "SubGroup", G_FaCn, OldMainGrCode, OldCurBal, IIf(OldCurBalType = "Cr", "-", "+")
            CalBalAcGroup "SubGroup", G_FaCn, NewMainGrCode, Val(txt(CurBal)), IIf(LblCurBalType.CAPTION = "Cr", "+", "-")
        End If
        '* Used For Alias
        '* If Previously alias exists and now it is blank
        If OldAlias <> "" And Trim(txt(AcAlias)) = "" Then
            AcID = G_FaCn.Execute("Select AcID From SubGroupAlias Where Name='" & OldAlias & "'").Fields(0).Value
            G_FaCn.Execute ("Delete * From SubGroupAlias Where AcID='" & AcID & "'")
            GCn.Execute ("Delete * From SubGroupAlias Where AcID='" & AcID & "'")
        '* If Previously Alias is Blank and now it Exists
        ElseIf OldAlias = "" And Trim(txt(AcAlias)) <> "" Then
            ID = G_FaCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0).Value
            AcIDAlias = PubSiteCode & IIf(PubFirmCode = "", "0", PubFirmCode) & Format(CStr(ID), "000000")

            TmpSQL = SubGroupUpdate("Add", "SubGroupAlias", AcIDAlias, txt(SubCode), txt(AcAlias), txt(AcAliasBiLang), "Y", "E")
            G_FaCn.Execute (TmpSQL)
            GCn.Execute (TmpSQL)
            G_FaCn.Execute ("Update SubGroupCounter Set SubGroupAcCode=SubGroupAcCode+1")
        ElseIf OldAlias <> "" Then
            AcID = G_FaCn.Execute("Select AcID From SubGroupAlias Where Name='" & OldAlias & "'").Fields(0).Value

            TmpSQL = SubGroupUpdate("Edit", "SubGroupAlias", AcID, txt(SubCode), txt(AcAlias), txt(AcAliasBiLang), "Y", "E")
            G_FaCn.Execute (TmpSQL)
            GCn.Execute (TmpSQL)
        End If
        Dim AmtCr As Double, AmtDr As Double, mID As Double
        ' Ledger Updation
        If mDocId = "" Then mDocId = VoucherNo
        G_FaCn.Execute "Delete From Ledger where DocId='" & mDocId & "'"
        G_FaCn.Execute "Delete From LedgerRef where DocId='" & mDocId & "'"
        For I = 1 To FGrid.Rows - 1
            If Val(FGrid.TextMatrix(I, Col_Amount)) <> 0 Then
                AmtCrDr = IIf(FGrid.TextMatrix(I, Col_CrDr) = "Cr", "AmtCr", "AmtDr")
                G_FaCn.Execute "Insert Into Ledger(" _
                & "DocId,V_SNo,V_Type,V_No,Site_Code," _
                & "V_Date,SubCode," & AmtCrDr & ",Chq_No,Chq_Date," _
                & "Narration,U_Name,U_EntDt,U_AE) " _
                & "Values(" _
                & "'" & mDocId & "'," & I & ",'" & mVType & "'," & VNo & ",'" & PubSiteCode & PubSiteCode & "'," _
                & "" & ConvertDate(IIf(FGrid.TextMatrix(I, Col_RefDate) = "", PubStartDate - 1, FGrid.TextMatrix(I, Col_RefDate))) & ",'" & txt(SubCode) & "'," & Val(FGrid.TextMatrix(I, Col_Amount)) & ",'" & FGrid.TextMatrix(I, Col_RefNo) & "'," & ConvertDate(FGrid.TextMatrix(I, Col_RefDate)) & "," _
                & "'Party Opening','" & pubUName & "',#" & PubServerDate & "#,'E')"
            End If
        AmtCr = Val(IIf(AmtCrDr = "AmtCr", Val(FGrid.TextMatrix(I, Col_Amount)), 0))
        AmtDr = Val(IIf(AmtCrDr = "AmtDr", Val(FGrid.TextMatrix(I, Col_Amount)), 0))
        mID = IIf(IsNull(G_FaCn.Execute("select Max(ID) from LedgerRef").Fields(0).Value), 0, G_FaCn.Execute("select Max(ID) from LedgerRef").Fields(0).Value) + 1
        GSQL = "INSERT INTO LedgerRef (ID,DOCID,V_SNO,DR,CR,SUBCODE,U_Name,U_EntDt,U_AE,AgRefType,AgRefNo,DueDate,V_dATE)" & _
                "VALUES (" & mID & ",'" & mDocId & "'," & I & "," & AmtDr & "," & AmtCr & ",'" & txt(SubCode) & "','" & pubUName & "'," & FaConvertDate(Now) & ",'E','New Ref','" & FGrid.TextMatrix(I, Col_RefNo) & "'," & FaConvertDate(Now) & "," & ConvertDate(IIf(FGrid.TextMatrix(I, Col_RefDate) = "", PubStartDate - 1, FGrid.TextMatrix(I, Col_RefDate))) & ")"
        G_FaCn.Execute GSQL
        Next
'         To Update Voucher_Prefix Serial No
        If VPrefixUpdateFlag = 0 Then
            If FGrid.Rows >= 2 Then
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & mVType & "'", G_FaCn, adOpenDynamic, adLockOptimistic
                If Rst.RecordCount > 0 Then
                    G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=Start_Srl_No+1 Where V_Type='" & Rst!V_tYPE & "'"
                End If
            End If
        End If
    G_FaCn.CommitTrans
    GCn.CommitTrans
    mTrans = False
    mSearchCode = txt(SubCode)
    RsAcName.Requery
    RsAcAlias.Requery
    Master.Requery
'    TopCtrl1_eRef

    Master.FIND "SearchCode = '" & mSearchCode & "'"
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Set Rst = Nothing
Exit Sub
ELoop:
    If mTrans = True Then
        G_FaCn.RollbackTrans: GCn.RollbackTrans
    End If
    CheckError
End Sub
'Database updation procedure For Delete
Private Sub UpdateDataBaseDelete()
On Error GoTo ELoop
Dim vBook As Variant, mTrans As Boolean
'    MsgBox "Delete disabled due to Data Security Problem", vbInformation, "Validation Check"
'    Exit Sub
    
    '* For FA Data Delete..
    If G_FaCn.Execute("Select * From Ledger Where SubCode='" & txt(SubCode) & "'").RecordCount > 0 Then
        MsgBox "Transactions Exists for this Party.You can not Delete It.", vbInformation, "Validation Check"
        Exit Sub
    End If
    '* For Vehicle Data Delete..
    If GCn.Execute("Select * From Veh_Purch1 Where PartyCode='" & txt(SubCode) & "'").RecordCount > 0 Then
        MsgBox "Transactions Exists for this Party.You can not Delete It.", vbInformation, "Validation Check"
        Exit Sub
    End If
    
    If GCn.Execute("Select * From Veh_Order Where PartyCode='" & txt(SubCode) & "'").RecordCount > 0 Then
        MsgBox "Transactions Exists for this Party.You can not Delete It.", vbInformation, "Validation Check"
        Exit Sub
    End If
    
    '* For Spare Data Delete..
    If GCn.Execute("Select * From SP_Sale Where Party_Code='" & txt(SubCode) & "'").RecordCount > 0 Then
        MsgBox "Transactions Exists for this Party.You can not Delete It.", vbInformation, "Validation Check"
        Exit Sub
    End If
    
    If GCn.Execute("Select * From SP_Purch Where Party_Code='" & txt(SubCode) & "'").RecordCount > 0 Then
        MsgBox "Transactions Exists for this Party.You can not Delete It.", vbInformation, "Validation Check"
        Exit Sub
    End If
    
    If GCn.Execute("Select * From SP_Order Where Party_Code='" & txt(SubCode) & "'").RecordCount > 0 Then
        MsgBox "Transactions Exists for this Party.You can not Delete It.", vbInformation, "Validation Check"
        Exit Sub
    End If
    'For System Group Account..
    If SysGroup = "Y" Then
        MsgBox "This is a System Group.You Can not Delete this A/c", vbInformation, "Validation Check"
        Exit Sub
    End If
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        G_FaCn.BeginTrans
        GCn.BeginTrans
            mTrans = True
            GCn.Execute ("Delete From SubGroupAlias Where SubCode='" & txt(SubCode) & "'")
            GCn.Execute ("Delete From SubGroup Where SubCode='" & txt(SubCode) & "'")

            G_FaCn.Execute ("Delete From SubGroupAlias Where SubCode='" & txt(SubCode) & "'")
            G_FaCn.Execute ("Delete From SubGroup Where SubCode='" & txt(SubCode) & "'")
        G_FaCn.CommitTrans
        GCn.CommitTrans
        mTrans = False
        TopCtrl1_eRef
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
        End If
        BUTTONS True, Me, Master, 0
        MoveRec
    End If
Exit Sub
ELoop:
    If mTrans = True Then
        G_FaCn.RollbackTrans: GCn.RollbackTrans
    End If
    CheckError
End Sub

Private Sub DGCity_Click()
On Error GoTo ELoop
    DGCity.Visible = False
    If Rscity.RecordCount > 0 Then
        txt(DGCity.Tag).TEXT = Rscity!Name
        txt(DGCity.Tag).Tag = Rscity!Code
    End If
    txt(DGCity.Tag).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGPartyType_Click()
On Error GoTo ELoop
    DGPartyType.Visible = False
    If rsPartyType.RecordCount > 0 Then
        txt(PartyType).TEXT = rsPartyType!Name
        txt(PartyType).Tag = rsPartyType!Code
    End If
    txt(PartyType).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGTDSCat_Click()
On Error GoTo ELoop
    DGTDSCat.Visible = False
    If RsTDSCat.RecordCount > 0 Then
        txt(TDSCat).TEXT = RsTDSCat!Name
        txt(TDSCat).Tag = RsTDSCat!Code
    End If
    txt(TDSCat).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub DGUnderAc_Click()
On Error GoTo ELoop
    DGUnderAc.Visible = False
    If RsUnderAc.RecordCount > 0 Then
        txt(UnderGroup).TEXT = RsUnderAc!Name
        txt(UnderGroup).Tag = RsUnderAc!Code
    End If
    txt(UnderGroup).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub ListView_Click()
On Error GoTo ELoop
    txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    txt(Val(ListView.Tag)).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Load()
Dim I As Byte
    DetailFlag = 0
    VPrefixUpdateFlag = 1
    TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini
    ConvBiLanguage BiLanguage
    For I = 0 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg
        txt(I).ForeColor = CtrlFColOrg
    Next

    CodeEditFlag = False 'False 'True
    If CodeEditFlag = True Then
        Lbl(18).Visible = True
        LblColon(20).Visible = True
        txt(SubCode).Visible = True
    Else
        Lbl(18).Visible = False
        LblColon(20).Visible = False
        txt(SubCode).Visible = False
    End If

    Set RsAcName = New ADODB.Recordset
    RsAcName.CursorLocation = adUseClient
    RsAcName.Open "Select SubCode As Code,Name,AliasYN,NameHelp,GroupCode From SubGroupAlias Order by Name", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGAcName.DataSource = RsAcName

    Set RsAcAlias = RsAcName
    Set DGAcAlias.DataSource = RsAcAlias

    Set RsUnderAc = New ADODB.Recordset
    RsUnderAc.CursorLocation = adUseClient
    RsUnderAc.Open "Select GroupCode As Code,GroupName As Name,GroupNature,MainGrCode,CurrentBalance,SubLedYN,AliasYN,GroupHelp,Nature From AcGroup Where MainGrCode<>'999' Order by GroupName", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGUnderAc.DataSource = RsUnderAc

    Set Rscity = New ADODB.Recordset
    Rscity.CursorLocation = adUseClient
    Rscity.Open "Select CityCode As Code,CityName As Name,CityHelp From City Order by CityName", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGCity.DataSource = Rscity

    Set RsTDSCat = New ADODB.Recordset
    RsTDSCat.CursorLocation = adUseClient
    RsTDSCat.Open "Select TDS_Catg As Code,TDS_Desc As Name From TDSCat Order by TDS_Desc", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGTDSCat.DataSource = RsTDSCat

    Set rsPartyType = New ADODB.Recordset
    rsPartyType.CursorLocation = adUseClient
    rsPartyType.Open "Select Party_Type As Code,Description As Name From SubGroupType Order by Description", G_FaCn, adOpenDynamic, adLockOptimistic
    Set DGPartyType.DataSource = rsPartyType

    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
     Set Master = G_FaCn.Execute("Select s.SubCode As SearchCode,S.SubCode " _
            & "From SubGroup S " _
            & "Where s.AliasYN<>'Y' order by Name")

    Disp_Text SETS("INI", Me, Master)
    MoveRec
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set Master = Nothing: Set RsAcName = Nothing
    Set RsAcAlias = Nothing: Set RsAcNameHelp = Nothing
    Set RsUnderAc = Nothing: Set Rscity = Nothing
    Set RsTDSCat = Nothing: Set rsPartyType = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    DetailEnb 0
    txt(ActiveYN).TEXT = "Yes"
    txt(GovtPartyYN).TEXT = "No"
    txt(LC).TEXT = "Local"
    txt(Religion).TEXT = "N/A"

    txt(SubCode) = PubSiteCode & IIf(PubFirmCode = "", "0", PubFirmCode) & Format(G_FaCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0).Value, "000000")
    txt(SubCode).Tag = txt(SubCode).TEXT

    If CodeEditFlag = True Then
        txt(SubCode).SetFocus
    Else
        txt(AcName).SetFocus
    End If
    OldCurBal = 0
    OldCurBalType = ""
    SysGroup = "N"
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    If Master.RecordCount > 0 Then
        Disp_Text SETS("EDIT", Me, Master)
        txt(SubCode).Enabled = False
        DetailEnb DetailFlag
        FGrid.AddItem FGrid.Rows
        If SysGroup = "Y" Then
            txt(AcName).Enabled = False
            txt(UnderGroup).Enabled = False
'            Txt(AcAlias).SetFocus
            FGrid.SetFocus
        Else
            txt(AcName).SetFocus
        End If
    Else
        MsgBox "There Is No Record To Edit.", vbInformation, "Information"
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
    If Master.RecordCount > 0 Then
        UpdateDataBaseDelete
    Else
        MsgBox "There Is No Record To Delete", vbInformation, "Information"
    End If
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
    GSQL = "Select SubCode As SearchCode,Name,ConPrefix,ConPerson," _
        & "Add1,Add2,Add3,C.CityName,PIN,Phone," _
        & "Mobile,FAX,EMail,CSTNo As CST,LSTNo As LST,PANNo As PAN," _
        & "Switch(ActiveYN=0,'No',ActiveYN=1,'Yes') As Active," _
        & "Switch(Govt_YN=0,'No',Govt_YN=1,'Yes') As GovtParty," _
        & "CreditLimit,CreditDays,Remark " _
        & "From SubGroup S Left Join City C on S.CityCode=C.CityCode " _
        & "Where S.AliasYN<>'Y' Order by Name"
    Set SearchForm = Me
    FIND2.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_ePrn()
Dim RstPrn As New ADODB.Recordset, X11 As Variant
On Error GoTo Errloop
    Set RstPrn = G_FaCn.Execute("SELECT AcGroup.GroupName, SubGroup.Name,1 as XNo " & _
                             "FROM AcGroup INNER JOIN SubGroup ON AcGroup.GroupCode = SubGroup.GroupCode")
    X11 = CreateFieldDefFile(RstPrn, PubRepoPath + "\SubGroupPrn.ttx", True)

    Set rpt = rdApp.OpenReport(PubRepoPath + "\SubGroupPrn.RPT")
    rpt.Database.SetDataSource RstPrn
    rpt.ReadRecords
    Report_View rpt, Me.CAPTION, 0, False
    Set RstPrn = Nothing
    Set rpt = Nothing
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eRef()
    RsAcName.Requery
    'RsAcNameHelp.Requery
    RsAcAlias.Requery
    RsUnderAc.Requery
    Rscity.Requery
    RsTDSCat.Requery
    rsPartyType.Requery
    Master.Requery
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer
    If txtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid_LostFocus 0
            txtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    If IsValid(txt(SubCode), "A/C Code") = False Then Exit Sub
    If IsValid(txt(AcName), "A/C Name") = False Then Exit Sub
    If IsValid(txt(UnderGroup), "Under Group") = False Then Exit Sub
    For I = 1 To FGrid.Rows - 1
        If Val(FGrid.TextMatrix(I, Col_Amount)) <> 0 Then
            If FGrid.TextMatrix(I, Col_CrDr) = "" Then MsgBox "Please Specify Cr/Dr in Row No " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_CrDr: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
        End If
        If FGrid.TextMatrix(I, Col_RefDate) <> "" Then
            If CDate(FGrid.TextMatrix(I, Col_RefDate)) >= PubStartDate Then
                MsgBox "Date should be before financial year", vbCritical, "Automan"
                FGrid.Row = I: FGrid.Col = Col_RefDate: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
            End If
        End If
    Next

    If TopCtrl1.TopText2 = "Add" Then
        UpdateDataBaseAdd
    Else
        UpdateDataBaseEdit
    End If
    If MasterFormExit Then Unload Me: Exit Sub
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    If MasterFormExit Then Unload Me: Exit Sub
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

Private Sub txt_GotFocus(Index As Integer)
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case NamePrefix
        ListArray = Array("    ", "Mr.", "Mrs.", "Miss", "Ms", "M/S")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 5)
    Case AcName
        If RsAcName.RecordCount = 0 Or (RsAcName.EOF = True Or RsAcName.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsAcName!Name Then
            RsAcName.MoveFirst
            RsAcName.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case AcAlias
        If RsAcAlias.RecordCount = 0 Or (RsAcAlias.EOF = True Or RsAcAlias.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsAcAlias!Name Then
            RsAcAlias.MoveFirst
            RsAcAlias.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case UnderGroup
        If RsUnderAc.RecordCount = 0 Or (RsUnderAc.EOF = True Or RsUnderAc.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsUnderAc!Name Then
            RsUnderAc.MoveFirst
            RsUnderAc.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case ConPersonPrefix
        ListArray = Array("Mr.", "Mrs.", "Miss") ', "M/S")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 3)
    Case ConPersonPrefixB
        ListArray = Array("S/O", "W/O", "D/O", "C/O", "And ", "U/C")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 6)
    Case City, CityB
        DGCity.Tag = Index
'        DGCity.left = txt(Index).left: DGCity.top = txt(Index).top + txt(Index).height + 15
        If Rscity.RecordCount = 0 Or (Rscity.EOF = True Or Rscity.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> Rscity!Name Then
            Rscity.MoveFirst
            Rscity.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case Religion
        ListArray = Array("N/A", "Hindu", "Muslim", "Sikh", "Christian")
        Set mListItem = ListView_Items(ListView, txt, Religion, ListArray, 5)
    Case PartyType
        If rsPartyType.RecordCount = 0 Or (rsPartyType.EOF = True Or rsPartyType.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsPartyType!Name Then
            rsPartyType.MoveFirst
            rsPartyType.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
    Case LC
        ListArray = Array("Local", "Central")
        Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 2)
    Case TDSCat
        If RsTDSCat.RecordCount = 0 Or (RsTDSCat.EOF = True Or RsTDSCat.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsTDSCat!Name Then
            RsTDSCat.MoveFirst
            RsTDSCat.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
End Select
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case NamePrefix
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 1800
    Case AcName
        DGridTxtKeyDown_Mast DGAcName, txt, Index, RsAcName, KeyCode, False, 1
    Case AcAlias
        DGridTxtKeyDown_Mast DGAcAlias, txt, Index, RsAcAlias, KeyCode, False, 1
    Case UnderGroup
        DGridTxtKeyDown DGUnderAc, txt, Index, RsUnderAc, KeyCode, False, 1, frmGrEnt, "frmGrEnt"
    Case ConPersonPrefix
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 1200
    Case ConPersonPrefixB
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 1800
    Case City, CityB
        DGridTxtKeyDown DGCity, txt, Index, Rscity, KeyCode, False, 1, frmCity, "frmCity"
    Case Religion
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 1500
    Case PartyType
        DGridTxtKeyDown DGPartyType, txt, Index, rsPartyType, KeyCode, False, 1
    Case LC
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case TDSCat
        DGridTxtKeyDown DGTDSCat, txt, Index, RsTDSCat, KeyCode, False, 1
    Case PhoneB
        If BiLanguage = False Then
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
                If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
                    TopCtrl1_eSave
                Else
                    txt(Index).SetFocus
                    Exit Sub
                End If
            End If
        End If
    Case Add3BiLang
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
                TopCtrl1_eSave
            Else
                txt(Index).SetFocus
                Exit Sub
            End If
        End If
End Select
If FrmList.Visible = False And DGAcName.Visible = False And DGAcAlias.Visible = False And DGUnderAc.Visible = False And DGCity.Visible = False And DGTDSCat.Visible = False And DGPartyType.Visible = False Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If CodeEditFlag = True Then
            If Index <> SubCode And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        Else
'            If Index <> AcName And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            If Index <> NamePrefix And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
        If SysGroup = "Y" Then
            If BiLanguage = True Then
                If Index <> AcNameBiLang And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            Else
                If Index <> AcAlias And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        Else
'            If Index <> AcName And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            If Index <> NamePrefix And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
    End If
End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
Select Case Index
    Case UnderGroup
        If DGUnderAc.Visible = True Then DGridTxtKeyPress txt, Index, RsUnderAc, KeyAscii, "Name"
    Case City, CityB
        If DGCity.Visible = True Then DGridTxtKeyPress txt, Index, Rscity, KeyAscii, "Name"
    Case PartyType
        If DGPartyType.Visible = True Then DGridTxtKeyPress txt, Index, rsPartyType, KeyAscii, "Name"
    Case TDSCat
        If DGTDSCat.Visible = True Then DGridTxtKeyPress txt, Index, RsTDSCat, KeyAscii, "Name"
    Case CrLimit
        NumPress txt(CrLimit), KeyAscii, 9, 2
    Case CrDays
        NumPress txt(CrDays), KeyAscii, 3, 0
    Case ActiveYN, GovtPartyYN
        If KeyAscii = 89 Or KeyAscii = 121 Or KeyAscii = 78 Or KeyAscii = 110 Then
            If KeyAscii = 89 Or KeyAscii = 121 Then         ' Y/y
                txt(Index).TEXT = "Yes"
                KeyAscii = 0
            ElseIf KeyAscii = 78 Or KeyAscii = 110 Then     ' N/n
                txt(Index).TEXT = "No"
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case AcName
        If DGAcName.Visible = True Then DGridTxtKeyUp_Mast txt, Index, RsAcName, KeyCode, "Name"
    Case AcAlias
        If DGAcAlias.Visible = True Then DGridTxtKeyUp_Mast txt, Index, RsAcAlias, KeyCode, "Name"
    Case NamePrefix, ConPersonPrefix, ConPersonPrefixB, Religion, LC
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Ctrl_validate txt(Index)
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, SameName As Byte, SameName1 As Byte
Select Case Index
    Case SubCode
        If txt(SubCode).TEXT = "" Then txt(SubCode) = txt(SubCode).Tag: txt(SubCode).SetFocus: Exit Sub
        If G_FaCn.Execute("Select SubCode From SubGroup Where SubCode='" & txt(SubCode) & "'").RecordCount > 0 Then
            MsgBox "A/c Code Already Exists", vbInformation, "Validation"
            txt(SubCode) = txt(SubCode).Tag
            txt(SubCode).SetFocus
            Exit Sub
        End If
    Case AcName
        If txt(Index).TEXT = "" Then Exit Sub
        If TopCtrl1.TopText2 = "Add" Then         ' For Add Mode
            Set Rst = G_FaCn.Execute("Select NameHelp From SubGroupAlias Where NameHelp='" & FilterString(txt(AcName).TEXT) & "'")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate A/c Name not Allowed", vbInformation, "Validation"
                txt(AcName).SetFocus
                Cancel = True
                Exit Sub
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then      ' For Edit Mode
            Set Rst = G_FaCn.Execute("Select NameHelp From SubGroupAlias Where NameHelp='" & FilterString(txt(AcName).TEXT) & "' and Name<>'" & OldName & "'")
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate A/c Name not Allowed", vbInformation, "Validation"
                txt(AcName) = OldName
                txt(AcName).SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
    Case AcAlias
        If txt(Index).TEXT = "" Then Exit Sub
        If TopCtrl1.TopText2 = "Add" Then         ' For Add Mode
            If UCase(Trim(txt(AcAlias).TEXT)) = UCase(Trim(txt(AcName).TEXT)) Then SameName = 1
            Set Rst = G_FaCn.Execute("Select NameHelp From SubGroupAlias Where NameHelp='" & FilterString(txt(AcAlias)) & "'")
            If Rst.RecordCount > 0 Then SameName1 = 1
            If SameName = 1 Or SameName1 = 1 Then
                MsgBox "Duplicate Alias not Allowed", vbInformation, "Validation"
                txt(AcAlias).SetFocus
                Cancel = True
                Exit Sub
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then     ' For Edit Mode
            If UCase(Trim(txt(AcAlias).TEXT)) = UCase(Trim(txt(AcName).TEXT)) Then SameName = 1
            Set Rst = G_FaCn.Execute("Select NameHelp From SubGroupAlias Where NameHelp='" & FilterString(txt(AcAlias)) & "' and NameHelp<>'" & FilterString(OldAlias) & "'")
            If Rst.RecordCount > 0 Then SameName1 = 1
            If SameName = 1 Or SameName1 = 1 Then
                MsgBox "Duplicate Alias not Allowed", vbInformation, "Validation"
                txt(AcAlias).SetFocus
                Cancel = True
                Exit Sub
            End If
        End If
    Case UnderGroup
        If txt(Index).TEXT = "" Then Exit Sub
        '**** For Alias Eqivalent Name Searching
        If RsUnderAc.RecordCount > 0 Or (RsUnderAc.EOF = False Or RsUnderAc.BOF = False) Or txt(Index).TEXT <> "" Then
            Set Rst = G_FaCn.Execute("Select ID,GroupCode,GroupName,Nature,AliasYN From AcGroup Where GroupCode='" & RsUnderAc!Code & "'")
            If Rst.RecordCount > 0 Then
                If Rst!Nature = "Supplier" Or Rst!Nature = "Customer" Then
                    DetailFlag = 1
                    DetailEnb DetailFlag
                Else
                    DetailFlag = 0
                    DetailEnb DetailFlag
                End If
                LblNature = IIf(IsNull(Rst!Nature), "", Rst!Nature)
'                While Not Rst.EOF
'                    If Rst!AliasYN = "N" Then
'                        Txt(UnderGroup) = Trim(Rst!GroupName)
'                        Txt(UnderGroup).Tag = Rst!GroupCode     'Rst!ID
'                    End If
'                    Rst.MoveNext
'                Wend
            End If
        End If
    Case NamePrefix, ConPersonPrefix, ConPersonPrefixB, Religion, LC
        If txt(Index).TEXT <> "" Then txt(Index).TEXT = ListView.SelectedItem.TEXT
    Case Add1, Add1B, Add2, Add2B, Add3, Add3B, Pin, PinB, Phone, PhoneB
        If txt(Index) <> "" Then
            If Index = Add1 And txt(Add1B) = "" Then
                txt(Add1B) = txt(Add1)
            ElseIf Index = Add2 And txt(Add2B) = "" Then
                txt(Add2B) = txt(Add2)
            ElseIf Index = Add3 And txt(Add3B) = "" Then
                txt(Add3B) = txt(Add3)
            ElseIf Index = Pin And txt(PinB) = "" Then
                txt(PinB) = txt(Pin)
            ElseIf Index = Phone And txt(PhoneB) = "" Then
                txt(PhoneB) = txt(Phone)
            End If
        Else
            txt(Index) = ""
        End If
    Case City, CityB
        If Rscity.RecordCount > 0 Or (Rscity.EOF = False Or Rscity.BOF = False) Then
            If txt(Index).TEXT <> "" Then
                txt(Index).TEXT = Rscity!Name
                txt(Index).Tag = Rscity!Code
                If Index = City And txt(CityB) = "" Then
                    txt(CityB).TEXT = Rscity!Name
                    txt(CityB).Tag = Rscity!Code
                End If
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        End If
    Case PartyType
        If rsPartyType.RecordCount > 0 Or (rsPartyType.EOF = False Or rsPartyType.BOF = False) Then
            If txt(Index).TEXT <> "" Then
                txt(Index).TEXT = rsPartyType!Name
                txt(Index).Tag = rsPartyType!Code
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        End If
    Case TDSCat
        If RsTDSCat.RecordCount > 0 Or (RsTDSCat.EOF = False Or RsTDSCat.BOF = False) Then
            If txt(Index).TEXT <> "" Then
                txt(Index).TEXT = RsTDSCat!Name
                txt(Index).Tag = RsTDSCat!Code
            Else
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            End If
        End If
End Select
Set Rst = Nothing
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
    Ctrl_GetFocus txtGrid(Index)
    Grid_Hide
    FGrid.CellBackColor = CellBackColLeave
    txtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Col
    Case Col_RefNo
        txtGrid(0).MaxLength = 15
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then txtGrid(0) = txtGrid(0).Tag: Exit Sub
Select Case FGrid.Col
    Case Col_RefDate, Col_RefNo, Col_Amount
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then
                GridTxtDown FGrid, txtGrid, Index, KeyCode, TAddMode, Col_CrDr
            Else
                TxtGrid_LostFocus 0
                txtGrid(0).SetFocus
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
    Case Col_Amount
        NumPress txtGrid(Index), KeyAscii, 10, 2
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case FGrid.Col
    Case Col_Amount
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(txtGrid(Index).TEXT), "0.00")
End Select
If KeyCode = vbKeyEscape Then
    FGrid.SetFocus
    txtGrid(0).Visible = False
    Grid_Hide
End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_LostFocus(Index As Integer)
On Error GoTo ELoop
    If ExitCtrl = False Then Exit Sub
    Ctrl_validate txtGrid(Index)
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
    txtGrid(0).Visible = False
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub

Private Sub FGrid_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
FGrid_KeyPress vbKeyReturn
End Sub

Private Sub FGrid_EnterCell()
    FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
    FGrid.CellBackColor = CellBackColEnter
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
        FGrid.CellBackColor = CellBackColLeave
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And FGrid.Row = FGrid.Rows - 1 Then 'Val(FGrid.Tag) = FGrid.Rows - 1 Then
        FGrid.CellBackColor = CellBackColLeave
        If DetailFlag = 0 Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
                TopCtrl1_eSave
            Else
            
                FGrid.SetFocus
                FGrid_EnterCell
'                Exit Sub
            End If
        Else
            SendKeys vbTab
        End If
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    Select Case FGrid.Col
    Case Col_RefDate, Col_RefNo, Col_CrDr, Col_Amount
        If KeyCode = vbKeyDelete And Shift = 0 Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        End If
    End Select
    If KeyCode = vbKeyReturn Then
        Select Case FGrid.Col
            Case Col_RefDate, Col_RefNo, Col_CrDr, Col_Amount
                GridDblClick Me, FGrid, txtGrid, 0
        End Select
        TAddMode = False
    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
On Error GoTo ELoop
If FGrid.Col = Col_RefDate Then
    Get_Text Me, FGrid, txtGrid, 0, False, KeyAscii
End If
If FGrid.TextMatrix(FGrid.Row, Col_RefDate) <> "" Then
    Select Case FGrid.Col
        Case Col_RefNo
            Get_Text Me, FGrid, txtGrid, 0, False, KeyAscii
        Case Col_Amount
            Get_Text Me, FGrid, txtGrid, 0, True, KeyAscii
        Case Col_CrDr
            If UCase(Chr(KeyAscii)) = "D" Then
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Dr"
            ElseIf UCase(Chr(KeyAscii)) = "C" Then
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "Cr"
            End If
            KeyAscii = 0
            CalOpBal
    End Select
End If
If KeyAscii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGrid.Row >= 1 Then
            If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                If FGrid.Rows > 2 Then
                    FGrid.RemoveItem (FGrid.Row)
                    CalOpBal
                Else
                    FGrid.Rows = 1
                    FGrid.AddItem FGrid.Rows
                    FGrid.FixedRows = 1
                End If
            End If
            For I = 1 To FGrid.Rows - 1
                FGrid.TextMatrix(I, Col_SrNo) = I
            Next
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_Scroll()
    txtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub

