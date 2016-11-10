VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmWorkVehiMast 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Workshop Vehicle Master"
   ClientHeight    =   9000
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
   ScaleHeight     =   9000
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Txt 
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
      Left            =   3000
      MaxLength       =   11
      TabIndex        =   167
      Text            =   "77/777/7777"
      Top             =   2055
      Width           =   1230
   End
   Begin VB.TextBox Txt 
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
      Index           =   1
      Left            =   3000
      TabIndex        =   166
      Text            =   "29-APR-2002"
      Top             =   615
      Width           =   1230
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   73
      Left            =   6870
      MaxLength       =   5
      TabIndex        =   9
      Top             =   1575
      Width           =   630
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   72
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   8
      Top             =   2295
      Width           =   2790
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   71
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1095
      Width           =   2790
   End
   Begin MSDataGridLib.DataGrid DGModel 
      Height          =   3810
      Left            =   -525
      Negotiate       =   -1  'True
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   8820
      Visible         =   0   'False
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   6720
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
      RowDividerStyle =   0
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Code"
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
         DataField       =   "Chas_Type"
         Caption         =   "Chassis Type"
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
         DataField       =   "CODE"
         Caption         =   "Model Type"
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
         Caption         =   "Model Name"
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
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   6900.095
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   70
      Left            =   11025
      MaxLength       =   6
      TabIndex        =   24
      Text            =   "999999"
      Top             =   2805
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   67
      Left            =   11025
      MaxLength       =   6
      TabIndex        =   21
      Text            =   "999999"
      Top             =   2535
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   69
      Left            =   10260
      MaxLength       =   6
      TabIndex        =   23
      Text            =   "999999"
      Top             =   2805
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   66
      Left            =   10260
      MaxLength       =   6
      TabIndex        =   20
      Text            =   "999999"
      Top             =   2535
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   68
      Left            =   9480
      MaxLength       =   6
      TabIndex        =   22
      Text            =   "999999"
      Top             =   2805
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   65
      Left            =   9480
      MaxLength       =   6
      TabIndex        =   19
      Text            =   "999999"
      Top             =   2535
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DGColor 
      Height          =   4935
      Left            =   1620
      Negotiate       =   -1  'True
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   7065
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
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
         Caption         =   "Color"
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGDeal 
      Height          =   4935
      Left            =   5835
      Negotiate       =   -1  'True
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   8070
      Visible         =   0   'False
      Width           =   6150
      _ExtentX        =   10848
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
         Caption         =   "Dealer "
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
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4500.284
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   64
      Left            =   6870
      MaxLength       =   50
      TabIndex        =   11
      Top             =   1815
      Width           =   2685
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   4980
      Left            =   2325
      Negotiate       =   -1  'True
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   6975
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   8784
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   63
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   4
      Text            =   "999999"
      Top             =   1335
      Width           =   990
   End
   Begin VB.TextBox Txt 
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
      Left            =   6870
      MaxLength       =   10
      TabIndex        =   13
      Top             =   855
      Width           =   675
   End
   Begin VB.TextBox Txt 
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
      Height          =   210
      Index           =   14
      Left            =   10005
      MaxLength       =   3
      TabIndex        =   16
      Top             =   1095
      Width           =   555
   End
   Begin VB.TextBox Txt 
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
      Left            =   8580
      MaxLength       =   11
      TabIndex        =   18
      Text            =   "28-APR-2003"
      Top             =   1335
      Width           =   1335
   End
   Begin VB.TextBox Txt 
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
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   10
      Top             =   2535
      Width           =   2790
   End
   Begin VB.TextBox Txt 
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
      Left            =   1440
      MaxLength       =   14
      TabIndex        =   7
      Text            =   "UP78-AE1234"
      Top             =   2055
      Width           =   1530
   End
   Begin VB.TextBox Txt 
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
      Left            =   6870
      MaxLength       =   40
      TabIndex        =   12
      Top             =   615
      Width           =   4770
   End
   Begin VB.TextBox Txt 
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
      Left            =   6870
      MaxLength       =   10
      TabIndex        =   17
      Top             =   1335
      Width           =   1680
   End
   Begin VB.TextBox Txt 
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
      Left            =   6870
      MaxLength       =   11
      TabIndex        =   15
      Top             =   1095
      Width           =   1680
   End
   Begin VB.TextBox Txt 
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
      Left            =   7575
      MaxLength       =   50
      TabIndex        =   14
      Top             =   855
      Width           =   4065
   End
   Begin VB.TextBox Txt 
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
      Left            =   1440
      MaxLength       =   25
      TabIndex        =   6
      Top             =   1815
      Width           =   2790
   End
   Begin VB.TextBox Txt 
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
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1575
      Width           =   2790
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   2
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   2
      Top             =   855
      Width           =   2790
   End
   Begin VB.TextBox Txt 
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
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "F999999"
      Top             =   615
      Width           =   1530
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4050
      Left            =   210
      TabIndex        =   73
      Top             =   3075
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   7144
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Technical Details"
      TabPicture(0)   =   "frmWorkVehiMast.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tyres Details"
      TabPicture(1)   =   "frmWorkVehiMast.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(0)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Owner Details"
      TabPicture(2)   =   "frmWorkVehiMast.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00BAD3C9&
         BorderStyle     =   0  'None
         Height          =   3885
         Index           =   0
         Left            =   -74970
         TabIndex        =   83
         Top             =   345
         Width           =   11520
         Begin VB.TextBox Txt 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   43
            Left            =   7710
            MaxLength       =   15
            TabIndex        =   54
            Top             =   1545
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   45
            Left            =   7695
            MaxLength       =   15
            TabIndex        =   56
            Top             =   2235
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   47
            Left            =   7695
            MaxLength       =   15
            TabIndex        =   58
            Top             =   2475
            Width           =   2385
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   48
            Left            =   4050
            MaxLength       =   15
            TabIndex        =   59
            Top             =   2940
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   7695
            MaxLength       =   15
            TabIndex        =   50
            Top             =   630
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   7710
            MaxLength       =   15
            TabIndex        =   52
            Top             =   1305
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   44
            Left            =   4035
            MaxLength       =   15
            TabIndex        =   55
            Top             =   2235
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   46
            Left            =   4035
            MaxLength       =   15
            TabIndex        =   57
            Top             =   2475
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   42
            Left            =   4050
            MaxLength       =   15
            TabIndex        =   53
            Top             =   1545
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   4050
            MaxLength       =   15
            TabIndex        =   51
            Top             =   1305
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   38
            Left            =   4065
            MaxLength       =   15
            TabIndex        =   49
            Top             =   630
            Width           =   2385
         End
         Begin VB.Shape Shape1 
            Height          =   720
            Index           =   2
            Left            =   2550
            Top             =   2130
            Width           =   8055
         End
         Begin VB.Shape Shape1 
            Height          =   720
            Index           =   1
            Left            =   2550
            Top             =   1200
            Width           =   8055
         End
         Begin VB.Shape Shape1 
            Height          =   465
            Index           =   0
            Left            =   2550
            Top             =   510
            Width           =   8055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Wheel"
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
            Index           =   56
            Left            =   795
            TabIndex        =   120
            Top             =   2955
            Width           =   1095
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
            Index           =   58
            Left            =   2190
            TabIndex        =   119
            Top             =   2955
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Right -1"
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
            Index           =   55
            Left            =   6690
            TabIndex        =   118
            Top             =   2265
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Right -2"
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
            Index           =   54
            Left            =   6690
            TabIndex        =   117
            Top             =   2505
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rear"
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
            Index           =   53
            Left            =   795
            TabIndex        =   116
            Top             =   1305
            Width           =   405
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
            Index           =   55
            Left            =   2190
            TabIndex        =   115
            Top             =   2250
            Width           =   75
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Right"
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
            Index           =   52
            Left            =   6690
            TabIndex        =   114
            Top             =   645
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Right -1"
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
            Index           =   51
            Left            =   6705
            TabIndex        =   113
            Top             =   1320
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Right -2"
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
            Index           =   50
            Left            =   6705
            TabIndex        =   112
            Top             =   1560
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle"
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
            Index           =   33
            Left            =   795
            TabIndex        =   111
            Top             =   2250
            Width           =   540
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
            Index           =   51
            Left            =   2190
            TabIndex        =   110
            Top             =   1305
            Width           =   75
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Left -1"
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
            Index           =   32
            Left            =   3105
            TabIndex        =   109
            Top             =   2250
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Left -2"
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
            Left            =   3105
            TabIndex        =   108
            Top             =   2490
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Front"
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
            Index           =   10
            Left            =   795
            TabIndex        =   107
            Top             =   630
            Width           =   435
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
            Index           =   3
            Left            =   2190
            TabIndex        =   106
            Top             =   630
            Width           =   75
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Left -2"
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
            Index           =   2
            Left            =   3120
            TabIndex        =   86
            Top             =   1545
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Left -1"
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
            Left            =   3120
            TabIndex        =   85
            Top             =   1305
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Left"
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
            Left            =   3105
            TabIndex        =   84
            Top             =   630
            Width           =   315
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00BAD3C9&
         BorderStyle     =   0  'None
         Height          =   3915
         Index           =   1
         Left            =   45
         TabIndex        =   87
         Top             =   285
         Width           =   11475
         Begin VB.TextBox Txt 
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
            Left            =   3030
            MaxLength       =   10
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   1470
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   27
            Left            =   8130
            MaxLength       =   10
            TabIndex        =   42
            Top             =   1455
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   8130
            MaxLength       =   20
            TabIndex        =   38
            Top             =   495
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   21
            Left            =   8130
            MaxLength       =   25
            TabIndex        =   37
            Top             =   255
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   19
            Left            =   3030
            MaxLength       =   20
            TabIndex        =   36
            Top             =   2910
            Width           =   2385
         End
         Begin VB.Frame FrmList 
            BorderStyle     =   0  'None
            Height          =   810
            Left            =   10740
            TabIndex        =   145
            Top             =   1950
            Visible         =   0   'False
            Width           =   1275
            Begin MSComctlLib.ListView ListView 
               Height          =   1815
               Left            =   30
               TabIndex        =   146
               TabStop         =   0   'False
               Top             =   75
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   3201
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
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   62
            Left            =   8130
            MaxLength       =   6
            TabIndex        =   46
            Top             =   2415
            Width           =   1275
         End
         Begin VB.TextBox Txt 
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
            Left            =   8130
            MaxLength       =   20
            TabIndex        =   48
            Top             =   2895
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   8130
            MaxLength       =   15
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Top             =   1215
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            ForeColor       =   &H00800080&
            Height          =   210
            Index           =   15
            Left            =   3030
            MaxLength       =   12
            TabIndex        =   25
            Top             =   270
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            ForeColor       =   &H00404040&
            Height          =   210
            Index           =   16
            Left            =   3030
            MaxLength       =   20
            TabIndex        =   26
            Top             =   510
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   3030
            MaxLength       =   20
            TabIndex        =   27
            Top             =   750
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   3030
            MaxLength       =   20
            TabIndex        =   28
            Top             =   990
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   8130
            MaxLength       =   20
            TabIndex        =   39
            Top             =   735
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   8130
            MaxLength       =   15
            TabIndex        =   40
            Top             =   975
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   24
            Left            =   3030
            MaxLength       =   15
            TabIndex        =   31
            Top             =   1710
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   28
            Left            =   8130
            MaxLength       =   15
            TabIndex        =   43
            Top             =   1695
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   29
            Left            =   3030
            MaxLength       =   10
            TabIndex        =   29
            Top             =   1230
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   3030
            MaxLength       =   15
            TabIndex        =   35
            Top             =   2670
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   31
            Left            =   3030
            MaxLength       =   15
            TabIndex        =   32
            Top             =   1950
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   3030
            MaxLength       =   15
            TabIndex        =   33
            Top             =   2190
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   3030
            MaxLength       =   15
            TabIndex        =   34
            Top             =   2430
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Height          =   210
            Index           =   34
            Left            =   8130
            MaxLength       =   15
            TabIndex        =   45
            Top             =   2175
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Left            =   8130
            MaxLength       =   10
            TabIndex        =   44
            Top             =   1935
            Width           =   2385
         End
         Begin VB.TextBox Txt 
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
            Index           =   36
            Left            =   8130
            MaxLength       =   6
            TabIndex        =   47
            Top             =   2655
            Width           =   2385
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Steering Gear Box No."
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
            Left            =   5940
            TabIndex        =   162
            Top             =   503
            Width           =   1935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Steering Type"
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
            Left            =   630
            TabIndex        =   161
            Top             =   1478
            Width           =   1200
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filler Cap Lock No."
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
            Left            =   5940
            TabIndex        =   160
            Top             =   1463
            Width           =   1605
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Paint Shade (Color)"
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
            Left            =   5940
            TabIndex        =   159
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cab No."
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
            Left            =   615
            TabIndex        =   158
            Top             =   2918
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Break Type"
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
            Index           =   41
            Left            =   5940
            TabIndex        =   144
            Top             =   2423
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuel Unit"
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
            Left            =   5940
            TabIndex        =   105
            Top             =   2663
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Additional Equipment"
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
            Left            =   5940
            TabIndex        =   104
            Top             =   2903
            Width           =   1800
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuel"
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
            Left            =   5940
            TabIndex        =   103
            Top             =   1943
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Coupon Book No."
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
            Height          =   210
            Index           =   6
            Left            =   630
            TabIndex        =   102
            Top             =   270
            Width           =   2100
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Front Axle No."
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
            Left            =   630
            TabIndex        =   101
            Top             =   758
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gear Box No."
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
            Left            =   630
            TabIndex        =   100
            Top             =   518
            Width           =   1155
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FIP No."
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
            Index           =   15
            Left            =   630
            TabIndex        =   99
            Top             =   1238
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Head Lamp Make"
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
            Left            =   5940
            TabIndex        =   98
            Top             =   1703
            Width           =   1470
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Steering Make"
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
            Left            =   5940
            TabIndex        =   97
            Top             =   1223
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Steering Lock No."
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
            Left            =   600
            TabIndex        =   96
            Top             =   1718
            Width           =   1515
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Door Lock No."
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
            Left            =   5940
            TabIndex        =   95
            Top             =   983
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Body No."
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
            Left            =   5940
            TabIndex        =   94
            Top             =   743
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rear Axle no."
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
            Left            =   630
            TabIndex        =   93
            Top             =   998
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Battery"
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
            Left            =   630
            TabIndex        =   92
            Top             =   2438
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alternator"
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
            Left            =   630
            TabIndex        =   91
            Top             =   1958
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Make Of Radiator"
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
            Left            =   5940
            TabIndex        =   90
            Top             =   2183
            Width           =   1485
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Starter Motor"
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
            Left            =   615
            TabIndex        =   89
            Top             =   2198
            Width           =   1140
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frame No."
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
            Index           =   23
            Left            =   630
            TabIndex        =   88
            Top             =   2678
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00BAD3C9&
         BorderStyle     =   0  'None
         Height          =   3870
         Index           =   2
         Left            =   -74985
         TabIndex        =   74
         Top             =   360
         Width           =   11535
         Begin TabDlg.SSTab SSTab2 
            Height          =   3600
            Left            =   7665
            TabIndex        =   137
            Top             =   240
            Visible         =   0   'False
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   6350
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Owner"
            TabPicture(0)   =   "frmWorkVehiMast.frx":0054
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Picture1(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Vehicle"
            TabPicture(1)   =   "frmWorkVehiMast.frx":0070
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Picture1(1)"
            Tab(1).ControlCount=   1
            Begin VB.Image Picture1 
               Height          =   3165
               Index           =   0
               Left            =   30
               Stretch         =   -1  'True
               Top             =   375
               Width           =   3645
            End
            Begin VB.Image Picture1 
               Height          =   3165
               Index           =   1
               Left            =   -74970
               Stretch         =   -1  'True
               Top             =   375
               Width           =   3645
            End
         End
         Begin MSComDlg.CommonDialog CDlg 
            Left            =   10560
            Top             =   60
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "Select Photograph"
            FileName        =   "*.*"
            Filter          =   "All Picture Files"
            InitDir         =   "C:"
         End
         Begin VB.TextBox Txt 
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
            Index           =   56
            Left            =   2550
            MaxLength       =   11
            TabIndex        =   68
            Top             =   2100
            Width           =   1425
         End
         Begin VB.TextBox Txt 
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
            Index           =   57
            Left            =   2550
            MaxLength       =   11
            TabIndex        =   69
            Top             =   2340
            Width           =   1425
         End
         Begin VB.TextBox Txt 
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
            Index           =   61
            Left            =   2550
            MaxLength       =   20
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   3060
            Width           =   2295
         End
         Begin VB.TextBox Txt 
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
            Index           =   53
            Left            =   2550
            MaxLength       =   25
            TabIndex        =   64
            Top             =   1140
            Width           =   2790
         End
         Begin VB.TextBox Txt 
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
            Index           =   60
            Left            =   2550
            MaxLength       =   20
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   2820
            Width           =   2295
         End
         Begin VB.TextBox Txt 
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
            Index           =   59
            Left            =   2550
            MaxLength       =   20
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   2580
            Width           =   2295
         End
         Begin VB.TextBox Txt 
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
            Index           =   58
            Left            =   2550
            MaxLength       =   40
            TabIndex        =   67
            Top             =   1860
            Width           =   4830
         End
         Begin VB.TextBox Txt 
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
            Index           =   55
            Left            =   2550
            MaxLength       =   25
            TabIndex        =   66
            Top             =   1620
            Width           =   4830
         End
         Begin VB.TextBox Txt 
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
            Index           =   54
            Left            =   2550
            MaxLength       =   12
            TabIndex        =   65
            Top             =   1380
            Width           =   2790
         End
         Begin VB.TextBox Txt 
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
            Index           =   51
            Left            =   2550
            MaxLength       =   40
            TabIndex        =   62
            Top             =   660
            Width           =   4830
         End
         Begin VB.TextBox Txt 
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
            Index           =   50
            Left            =   2550
            MaxLength       =   40
            TabIndex        =   61
            Top             =   420
            Width           =   4830
         End
         Begin VB.TextBox Txt 
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
            Index           =   49
            Left            =   2550
            MaxLength       =   10
            TabIndex        =   60
            Top             =   180
            Width           =   1425
         End
         Begin VB.TextBox Txt 
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
            Index           =   52
            Left            =   2550
            MaxLength       =   15
            TabIndex        =   63
            Top             =   900
            Width           =   4830
         End
         Begin VB.Label LblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(O)wner/(D)river"
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
            Index           =   0
            Left            =   4185
            TabIndex        =   143
            Top             =   165
            Width           =   1470
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Photograph"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Index           =   40
            Left            =   8625
            TabIndex        =   138
            Top             =   -15
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth"
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
            Left            =   750
            TabIndex        =   136
            Top             =   2108
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Marriage"
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
            Left            =   750
            TabIndex        =   135
            Top             =   2348
            Width           =   1440
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City "
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
            Index           =   29
            Left            =   750
            TabIndex        =   134
            Top             =   908
            Width           =   405
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
            Index           =   30
            Left            =   750
            TabIndex        =   133
            Top             =   1148
            Width           =   1230
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
            Index           =   31
            Left            =   750
            TabIndex        =   82
            Top             =   405
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Driven By"
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
            Index           =   28
            Left            =   750
            TabIndex        =   81
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Job No."
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
            Left            =   750
            TabIndex        =   80
            Top             =   2588
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person"
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
            Index           =   36
            Left            =   750
            TabIndex        =   79
            Top             =   1868
            Width           =   1305
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mail ID"
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
            Left            =   750
            TabIndex        =   78
            Top             =   1628
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile"
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
            Left            =   750
            TabIndex        =   77
            Top             =   1388
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Kilometers"
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
            Index           =   49
            Left            =   750
            TabIndex        =   76
            Top             =   3068
            Width           =   1320
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Job Date"
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
            Left            =   750
            TabIndex        =   75
            Top             =   2828
            Width           =   1155
         End
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Warr (Y/N) "
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
      Index           =   71
      Left            =   4335
      TabIndex        =   165
      Top             =   1583
      Width           =   1845
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Details "
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
      Index           =   69
      Left            =   105
      TabIndex        =   164
      Top             =   2310
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Variant"
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
      Index           =   68
      Left            =   105
      TabIndex        =   163
      Top             =   1110
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount % at the Time of Requisition Issue:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   67
      Left            =   5715
      TabIndex        =   157
      Top             =   2265
      Visible         =   0   'False
      Width           =   3645
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
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   66
      Left            =   10215
      TabIndex        =   156
      Top             =   2265
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MRP"
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
      Left            =   9645
      TabIndex        =   155
      Top             =   2265
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Taxpaid"
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
      Left            =   10995
      TabIndex        =   154
      Top             =   2265
      Visible         =   0   'False
      Width           =   645
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
      Left            =   9345
      TabIndex        =   153
      Top             =   2805
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2. Oil"
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
      Left            =   8595
      TabIndex        =   152
      Top             =   2805
      Visible         =   0   'False
      Width           =   420
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
      Left            =   9345
      TabIndex        =   151
      Top             =   2535
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1. Spare"
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
      Left            =   8595
      TabIndex        =   150
      Top             =   2535
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Locking Remarks"
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
      Index           =   43
      Left            =   4335
      TabIndex        =   149
      Top             =   1823
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis Type"
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
      Index           =   42
      Left            =   105
      TabIndex        =   147
      Top             =   1350
      Width           =   1140
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Y)es/(N)o"
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
      Left            =   10650
      TabIndex        =   142
      Top             =   1125
      Width           =   900
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
      Index           =   74
      Left            =   9885
      TabIndex        =   132
      Top             =   1155
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt Vehicle"
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
      Index           =   72
      Left            =   8820
      TabIndex        =   131
      Top             =   1125
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No."
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
      Index           =   70
      Left            =   105
      TabIndex        =   130
      Top             =   2550
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Dealer Code && Name"
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
      Index           =   65
      Left            =   4335
      TabIndex        =   129
      Top             =   863
      Width           =   2445
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery / Selling Date"
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
      Index           =   64
      Left            =   4335
      TabIndex        =   128
      Top             =   1103
      Width           =   1950
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mfg. Invoice No. && Date"
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
      Index           =   63
      Left            =   4335
      TabIndex        =   127
      Top             =   1343
      Width           =   2040
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name"
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
      Index           =   62
      Left            =   4335
      TabIndex        =   126
      Top             =   623
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RegNo.&&Date"
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
      Index           =   61
      Left            =   105
      TabIndex        =   125
      Top             =   2070
      Width           =   1140
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
      Index           =   60
      Left            =   105
      TabIndex        =   124
      Top             =   1830
      Width           =   915
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
      Index           =   59
      Left            =   105
      TabIndex        =   123
      Top             =   1590
      Width           =   1005
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
      Index           =   58
      Left            =   105
      TabIndex        =   122
      Top             =   870
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Srl No && Date"
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
      Index           =   57
      Left            =   105
      TabIndex        =   121
      Top             =   630
      Width           =   1170
   End
End
Attribute VB_Name = "frmWorkVehiMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long

Dim RsModel As ADODB.Recordset
Dim RsDeal As ADODB.Recordset
Dim RsColor As ADODB.Recordset
Dim RsCity As ADODB.Recordset
Dim Master As ADODB.Recordset
Private Const SNo As Byte = 0
Private Const SDate As Byte = 1
Private Const Model As Byte = 2
Private Const ChasNo As Byte = 3
Private Const Engno As Byte = 4
Private Const DNAME As Byte = 5
Private Const DCode As Byte = 6
Private Const SellDate As Byte = 7
Private Const InvNo As Byte = 8
Private Const OwName As Byte = 9
Private Const RegNo As Byte = 10
Private Const RegDate As Byte = 11
Private Const Serial As Byte = 12
Private Const InvDate As Byte = 13
Private Const GovtVeh As Byte = 14
Private Const ServBookNo As Byte = 15
Private Const GearBoxNo As Byte = 16
Private Const FAxleNo  As Byte = 17
Private Const RAxleNo As Byte = 18
Private Const CabNo As Byte = 19
Private Const BodyNo As Byte = 20
Private Const Paint As Byte = 21
Private Const DoorNo As Byte = 22
Private Const SGearBoxNo As Byte = 23
Private Const SteerNo As Byte = 24
Private Const SteerType As Byte = 25
Private Const SteerMake As Byte = 26
Private Const FillCapNo As Byte = 27
Private Const HeadLamp As Byte = 28
Private Const fipno As Byte = 29
Private Const FrameNo As Byte = 30
Private Const Alternator As Byte = 31
Private Const StartMotor As Byte = 32
Private Const Battery As Byte = 33
Private Const Radiator As Byte = 34
Private Const FUEL As Byte = 35
Private Const FuelUnit As Byte = 36
Private Const AddEquip As Byte = 37
Private Const FrontL As Byte = 38
Private Const FrontR As Byte = 39
Private Const RearL1 As Byte = 40
Private Const RearR1 As Byte = 41
Private Const RearL2 As Byte = 42
Private Const RearR2 As Byte = 43
Private Const MidL1 As Byte = 44
Private Const MidR1 As Byte = 45
Private Const MidL2 As Byte = 46
Private Const MidR2 As Byte = 47
Private Const SpWheel As Byte = 48
Private Const DrivBy As Byte = 49
Private Const Add1 As Byte = 50
Private Const Add2 As Byte = 51
Private Const City As Byte = 52
Private Const Telephone As Byte = 53
Private Const Mobile As Byte = 54
Private Const MailID As Byte = 55
Private Const DOB As Byte = 56
Private Const DOM As Byte = 57
Private Const ContPerson As Byte = 58
Private Const LastJobNo As Byte = 59
Private Const LastJobDt As Byte = 60
Private Const LastKiloM As Byte = 61
Private Const Brake As Byte = 62
Private Const ChasType As Byte = 63
Private Const LockedText As Byte = 64

Private Const SprMrp As Byte = 65
Private Const SprTB As Byte = 66
Private Const SprTP As Byte = 67
Private Const OilMRP As Byte = 68
Private Const OilTB As Byte = 69
Private Const OilTP As Byte = 70

Private Const Varient As Byte = 71
Private Const VehDet As Byte = 72
Private Const ExtendWar As Byte = 73

Dim ListArray As Variant
Dim mListItem As ListItem
Dim TAddMode As Boolean
Dim FirmAddFlag As Byte
Dim GridKey As Integer
Dim ExitCtrl As Boolean

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
Dim I As Byte
WinSetting Me: Ini_Grid: TopCtrl1.Tag = PubUParam
    For I = 0 To 25
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    If PubMoveRecYn Then
        Master.Open "select CardNO as SearchCode from Hiscard where div_code='" & PubDivCode & "' Order by CardNo", GCn, adOpenDynamic, adLockOptimistic
    Else
        Set Master = GCn.Execute("Select Top 1 CardNO as SearchCode from Hiscard where div_code='" & PubDivCode & "' Order by CardNo")
    End If
    
    Set RsModel = New ADODB.Recordset
    With RsModel
        .CursorLocation = adUseClient
        .Open "SELECT Model as code,Model_Desc as name ,chas_type from Model  where (div_code='" & PubDivCode & "' or Div_Code='') order by Model_DESC", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DgModel.DataSource = RsModel
    
    Set RsDeal = New ADODB.Recordset
    With RsDeal
        .CursorLocation = adUseClient
        .Open "SELECT D_CODE as code,D_NAME as name  from AMD_DEALER order by D_NAME", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGDeal.DataSource = RsDeal
    
    Set RsColor = New ADODB.Recordset
    With RsColor
        .CursorLocation = adUseClient
        .Open "SELECT COL_CODE as code,COL_Desc as name from COLMAST order by COL_DESC", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGColor.DataSource = RsColor
    
    Set RsCity = New ADODB.Recordset
    With RsCity
        .CursorLocation = adUseClient
        .Open "SELECT CITYCODE as code,CITYNAME as name from CITY order by CITYNAME", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGCity.DataSource = RsCity
    'Data1.Refresh
'    If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    MoveRec
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
Set RsModel = Nothing
Set RsColor = Nothing
Set RsCity = Nothing
Set RsDeal = Nothing
Set Master = Nothing
End Sub

Private Sub ListView_Click()
Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
FrmList.Visible = False
Txt(Val(ListView.Tag)).SetFocus
End Sub

Private Sub Picture1_dblClick(Index As Integer)
On Error GoTo ELoop
If TopCtrl1.TopText2 = "Browse" Then Exit Sub
Select Case Index
    Case 0
        CDlg.InitDir = Pub_DataPath & "\Pictures\Owner"
        If Trim(Picture1(Index).Tag) = Trim(Pub_DataPath & "\Pictures\Owner" & "\nopicture.bmp") Then
            CDlg.Action = 1
            Picture1(Index).Tag = Txt(SNo) & "O"
            FileCopy CDlg.FileName, Pub_DataPath & "\Pictures\Owner" & "\" & Picture1(Index).Tag
            Picture1(Index).Picture = LoadPicture(Pub_DataPath & "\Pictures\Owner" & "\" & Picture1(Index).Tag)
        Else
            If MsgBox("Do You Wnat To Change Photo,Previous Will Be Removed ?", vbYesNo, "Cancel") = vbYes Then
                CDlg.Action = 1
                Picture1(Index).Tag = Txt(SNo) & "O"
                FileCopy CDlg.FileName, Pub_DataPath & "\Pictures\Owner" & "\" & Picture1(Index).Tag
                Picture1(Index).Picture = LoadPicture(Pub_DataPath & "\Pictures\Owner" & "\" & Picture1(Index).Tag)
            End If
        End If
    Case 1
        CDlg.InitDir = Pub_DataPath & "\Pictures\Vehicle"
        If Trim(Picture1(Index).Tag) = Trim(Pub_DataPath & "\Pictures\Vehicle" & "\nopicture.bmp") Then
            CDlg.Action = 1
            Picture1(Index).Tag = Txt(SNo) & "V"
            FileCopy CDlg.FileName, Pub_DataPath & "\Pictures\Vehicle" & "\" & Picture1(Index).Tag
            Picture1(Index).Picture = LoadPicture(Pub_DataPath & "\Pictures\Owner" & "\" & Picture1(Index).Tag)
        Else
            If MsgBox("Do You Wnat To Change Photo,Previous Will Be Removed ?", vbYesNo, "Cancel") = vbYes Then
                CDlg.Action = 1
                Picture1(Index).Tag = Txt(SNo) & "V"
                FileCopy CDlg.FileName, Pub_DataPath & "\Pictures\Vehicle" & "\" & Picture1(Index).Tag
                Picture1(Index).Picture = LoadPicture(Pub_DataPath & "\Pictures\Vehicle" & "\" & Picture1(Index).Tag)
            End If
        End If
End Select
Exit Sub
ELoop:
    If err.NUMBER <> 0 Then MsgBox err.Description

End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim urs As Recordset
'Dim i As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    
    Set urs = GCn.Execute("select max(" & cVal(cMID("cardno", "3", "6")) & "),max(carddate) from hiscard")
    Txt(SNo).TEXT = PubSiteCode + Right("000000" & IIf(IsNull((urs.Fields(0).Value)), 1, (urs.Fields(0).Value) + 1), 6)
    Set urs = Nothing
    Txt(SNo).Enabled = False
    Txt(SDate) = PubLoginDate
    Txt(GovtVeh) = "No"
    Txt(Model).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
Dim XBM
On Error GoTo eloop1
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
        XBM = Master.Bookmark
        GCn.Execute ("delete from hiscard where cardno = '" & Master!SearchCode & "'")
        GCn.CommitTrans
        Master.Requery
        If Master.RecordCount >= XBM Then
            Master.Bookmark = XBM
        Else
            If Master.EOF = False Then Master.MoveLast
        End If
        BUTTONS True, Me, Master, 0
        Call MoveRec
    End If

eloop1:
    If err.NUMBER <> 0 Then GCn.RollbackTrans:      CheckError
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    If Txt(SteerType) = "Power" Then Txt(SteerMake).Enabled = False
    Txt(SNo).Enabled = False
    Txt(SDate).SetFocus
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
    GSQL = "select CardNo as SearchCode,CardNo, Chassis,RegNo,Engine,Model,Name,ConPerson as Contact_Person,Mobile,VehSerialNo, Supplier_BillNo, CouponNo from Hiscard  where div_code='" & PubDivCode & "' order by Chassis"
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
        Set Master = GCn.Execute("Select CardNO as SearchCode from Hiscard where div_code='" & PubDivCode & "' And CardNO  = '" & MyValue & "' Order by CardNo")
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
    For I = 0 To 25
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
End If
Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim I As Integer, mQry$, mRepName$, mAddRec As Boolean, ExitLoop As Boolean, RecCountFix As Integer
Dim Rst As ADODB.Recordset, RstRep As ADODB.Recordset, RST1 As ADODB.Recordset
Dim Rst2 As ADODB.Recordset, RST3 As ADODB.Recordset, Rst4 As ADODB.Recordset

On Error Resume Next

    mRepName = "VehHistory"
    mQry = "select H.CardNo,H.Site_Code,H.CardDate,H.Model,H.RegNo,H.RegDate," & _
        " H.Chas_Type,H.Chassis,H.Engine,H.VehSerialNo,H.Supplier_BillNo,H.Supplier_BillDate," & _
        " H.Dealer_Code,H.Delivery_Date,H.CouponNo,H.ColourCode," & _
        " H.Steer_Type,H.Steer_Make,H.Alternator,H.StarterMotor,H.Battery," & _
        " H.Name,H.ConPerson,H.Add1,H.Add2,H.Add3,H.CityCode,H.PhoneOff,H.PhoneResi,H.Mobile,H.Mail_ID," & _
        " H.DOB,H.DOM,H.OwnDrive,H.OwnerRemark,H.Next_JobDate,H.Ac_Code,H.Govt_YN,H.Inv_No,H.Locked_Text," & _
        " H.LJob_DocId,H.LJob_Date,H.LJob_AtKMsHrs,M.Model_Desc,M.Chas_Type AS ModelChasType," & _
        " col.Col_Desc, Amd_Dealer.D_Name,City.CityName, Emp.Emp_Name as LMechName " & _
        " from ((((((Hiscard H left join Model M on H.Model=M.Model) " & _
        " left join Colmast Col on H.ColourCode=Col.Col_Code) " & _
        " left join Amd_Dealer on H.Dealer_Code=Amd_Dealer.D_Code)" & _
        " left join City on H.CityCode=City.CityCode) " & _
        " Left Join Job_Card J on H.LJob_DocID=J.DocId) " & _
        " Left Join Emp_Mast Emp on J.RecBy_Mechanic=Emp.Emp_Code) " & _
        " where H.CardNo='" & Master!SearchCode & "'"
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    'Select Jobcard & other details
    
    mQry = "Select J.DocId,J.Job_Date,J.JobCloseDate,J.AtKMsHrs,J.Serv_Type,J.RecBy_Mechanic as RecBy_MechCode, Emp.Emp_Name as RecBy_MechName," & _
        " J.NetLab_Amt,S.Total_Amt as NetSpr_Amt,J.ObservBy_Super,J.ActionBy_Super" & _
        " from (Job_Card J Left join SP_Sale S on J.DocId_InvSpr=S.DocID) " & _
        " Left join Emp_Mast Emp on J.RecBy_Mechanic=Emp.Emp_Code " & _
        " where J.CardNo='" & Master!SearchCode & _
        "' Order by J.Job_Date Desc, j.Docid Desc"
    Set RST1 = New Recordset
    RST1.CursorLocation = adUseClient
    RST1.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    
        'Create temp table
        CreaTabRstRep RstRep, VehHisRstTmp
        'temp table created
    
    If RST1.RecordCount > 0 Then
        'Problem Reported Details
        mQry = "Select JD.Job_DocID,JD.S_No,JD.Code as Prob_Code,JD.Details as Prob_Reported " & _
            " from (Job_Card J Left join Job_Demand JD on J.DocId=JD.Job_DocID) " & _
            " where J.CardNo='" & Master!SearchCode & "' and JD.Job_DocID=J.DocID " & _
            " Order by JD.Job_DocID,JD.S_No"
        Set Rst2 = New Recordset
        Rst2.CursorLocation = adUseClient
        Rst2.Open (mQry), GCn, adOpenStatic, adLockReadOnly
        
        'Labour Done
        mQry = "Select JL.Job_DocID,JL.Lab_Code,L.Lab_Desc as Lab_Done " & _
            " from ((Job_Card J Left join Job_Lab JL on J.DocId=JL.Job_DocID) " & _
            " Left join Labour L on JL.Lab_Code=L.Lab_Code) " & _
            " where J.CardNo='" & Master!SearchCode & "' and JL.Job_DocID=J.DocID " & _
            " Order by JL.Job_DocId,JL.S_No"
        Set RST3 = New Recordset
        RST3.CursorLocation = adUseClient
        RST3.Open (mQry), GCn, adOpenStatic, adLockReadOnly

        'Requisition / Parts
        mQry = "Select distinct Stk.Job_DocID,Stk.DocId as DocIDReq, " & _
            " Stk.Part_No,Part.Part_Name,Stk.Purpose, Stk.Qty_Iss - Qty_Ret as Qty, Stk.Rate, Stk.Amount " & _
            " from ((Job_Card J Left join Sp_Stock as Stk on J.DocID=Stk.Job_DocId) " & _
            " left join Part on Stk.Part_No=Part.Part_No ) " & _
            " where J.CardNo='" & Master!SearchCode & "' and Stk.Job_DocID=J.DocID " & _
            " Order by Stk.Job_DocId,Stk.DocId"
        Set Rst4 = New Recordset
        Rst4.CursorLocation = adUseClient
        Rst4.Open (mQry), GCn, adOpenStatic, adLockReadOnly
        RST1.MoveFirst
        For I = 1 To RST1.RecordCount
            With RstRep
                .AddNew
                .Fields("JobDocID") = RST1!DocID
                .Fields("JobNo") = Replace(Right(RST1!DocID, 13), " ", "")
                .Fields("Job_Date") = RST1!Job_Date
                .Fields("JobCloseDate") = RST1!JobCloseDate
                .Fields("AtKMsHrs") = RST1!AtKMsHrs
                .Fields("Serv_Type") = RST1!Serv_Type
                .Fields("RecBy_MechCode") = RST1!RecBy_MechCode
                .Fields("RecBy_MechName") = RST1!RecBy_MechName
                .Fields("NetLab_Amt") = RST1!NetLab_Amt
                .Fields("NetSpr_Amt") = RST1!NetSpr_Amt
                .Fields("ObservBy_Super") = RST1!ObservBy_Super
                .Fields("ActionBy_Super") = RST1!ActionBy_Super
                .Update
            End With
            'Problem Reported Details
            
            If Rst2.RecordCount <= 0 Then GoTo lblLabDone
            Rst2.MoveFirst
            Rst2.FIND ("Job_DocID='" & RST1!DocID & "'")
            If Rst2.EOF Then GoTo lblLabDone
            ExitLoop = False
            mAddRec = False
            If I = 1 Then
               RstRep.MoveFirst
            Else
                RstRep.Move (2 - (RstRep.RecordCount - RecCountFix))
                If RstRep.EOF Then mAddRec = True
            End If
            Do While Not ExitLoop 'Rst2!Job_DocID = Rst1!DocId
                If Rst2!job_docid = RST1!DocID Then
                    With RstRep
                        If mAddRec Then
                            .AddNew
                            .Fields("JobDocID") = RST1!DocID
                            .Fields("JobNo") = Replace(Right(RST1!DocID, 13), " ", "")
                            .Fields("Job_Date") = RST1!Job_Date
                            .Fields("JobCloseDate") = RST1!JobCloseDate
                            .Fields("AtKMsHrs") = RST1!AtKMsHrs
                            .Fields("Serv_Type") = RST1!Serv_Type
                            .Fields("RecBy_MechCode") = RST1!RecBy_MechCode
                            .Fields("RecBy_MechName") = RST1!RecBy_MechName
                            .Fields("NetLab_Amt") = RST1!NetLab_Amt
                            .Fields("NetSpr_Amt") = RST1!NetSpr_Amt
                            .Fields("ObservBy_Super") = RST1!ObservBy_Super
                            .Fields("ActionBy_Super") = RST1!ActionBy_Super
                        End If
                        .Fields("JobDocID") = Rst2!job_docid
                        .Fields("Prob_Code") = Rst2!Prob_Code
                        .Fields("Prob_Reported") = Rst2!Prob_Reported
                        .Update
                    End With
                End If
                Rst2.MoveNext
                If Rst2.EOF Then
                    ExitLoop = True
                ElseIf Rst2!job_docid <> RST1!DocID Then
                    ExitLoop = True
                Else
                    RstRep.MoveNext
                    If RstRep.EOF Then mAddRec = True
                End If
            Loop
            
lblLabDone:
            'Labour Done
            If RST3.RecordCount <= 0 Then GoTo lblRequisition
            RST3.MoveFirst
            RST3.FIND ("Job_DocID='" & RST1!DocID & "'")
            If RST3.EOF Then GoTo lblRequisition
            ExitLoop = False
            mAddRec = False
            If I = 1 Then
               RstRep.MoveFirst
            Else
                RstRep.Move (2 - (RstRep.RecordCount - RecCountFix))
                If RstRep.EOF Then mAddRec = True
            End If
            Do While Not ExitLoop 'Rst3!Job_DocID = Rst1!DocId
                If RST3!job_docid = RST1!DocID Then
                    With RstRep
                        If mAddRec Then
                            .AddNew
                            .Fields("JobDocID") = RST1!DocID
                            .Fields("JobNo") = Replace(Right(RST1!DocID, 13), " ", "")
                            .Fields("Job_Date") = RST1!Job_Date
                            .Fields("JobCloseDate") = RST1!JobCloseDate
                            .Fields("AtKMsHrs") = RST1!AtKMsHrs
                            .Fields("Serv_Type") = RST1!Serv_Type
                            .Fields("RecBy_MechCode") = RST1!RecBy_MechCode
                            .Fields("RecBy_MechName") = RST1!RecBy_MechName
                            .Fields("NetLab_Amt") = RST1!NetLab_Amt
                            .Fields("NetSpr_Amt") = RST1!NetSpr_Amt
                            .Fields("ObservBy_Super") = RST1!ObservBy_Super
                            .Fields("ActionBy_Super") = RST1!ActionBy_Super
                        End If
                        .Fields("JobDocID") = RST1!DocID
                        .Fields("Lab_Code") = RST3!Lab_Code
                        .Fields("Lab_Done") = RST3!Lab_Done
                        .Update
                    End With
                End If
                RST3.MoveNext
                If RST3.EOF Then
                    ExitLoop = True
                ElseIf RST3!job_docid <> RST1!DocID Then
                    ExitLoop = True
                Else
                    RstRep.MoveNext
                    If RstRep.EOF Then mAddRec = True
                End If
            Loop
            
lblRequisition:
            'Requisition / Parts
            If Rst4.RecordCount <= 0 Then GoTo lblRst1MoveNext
            Rst4.MoveFirst
            Rst4.FIND ("Job_DocID='" & RST1!DocID & "'")
            If Rst4.EOF Then GoTo lblRst1MoveNext
            ExitLoop = False
            mAddRec = False
            If I = 1 Then
               RstRep.MoveFirst
            Else
                RstRep.Move (2 - (RstRep.RecordCount - RecCountFix))
                If RstRep.EOF Then mAddRec = True
            End If
            Do While Not ExitLoop 'Rst4!Job_DocID = Rst1!DocId
                If Rst4!job_docid = RST1!DocID Then
                    With RstRep
                        If mAddRec Then
                            .AddNew
                            .Fields("JobDocID") = RST1!DocID
                            .Fields("JobNo") = Replace(Right(RST1!DocID, 13), " ", "")
                            .Fields("Job_Date") = RST1!Job_Date
                            .Fields("JobCloseDate") = RST1!JobCloseDate
                            .Fields("AtKMsHrs") = RST1!AtKMsHrs
                            .Fields("Serv_Type") = RST1!Serv_Type
                            .Fields("RecBy_MechCode") = RST1!RecBy_MechCode
                            .Fields("RecBy_MechName") = RST1!RecBy_MechName
                            .Fields("NetLab_Amt") = RST1!NetLab_Amt
                            .Fields("NetSpr_Amt") = RST1!NetSpr_Amt
                            .Fields("ObservBy_Super") = RST1!ObservBy_Super
                            .Fields("ActionBy_Super") = RST1!ActionBy_Super
                        End If
                        .Fields("JobDocID") = RST1!DocID
                        .Fields("DocIDReq") = Rst4!DocIDReq
                        .Fields("Part_No") = Rst4!Part_No
                        .Fields("Part_Name") = Rst4!Part_Name
                        .Fields("Purpose") = Rst4!Purpose
                        .Fields("Rate") = Rst4!Rate
                        .Fields("Qty") = Rst4!Qty
                        .Fields("Amount") = Rst4!Amount
                        .Update
                    End With
                End If
                Rst4.MoveNext
                If Rst4.EOF Then
                    ExitLoop = True
                ElseIf Rst4!job_docid <> RST1!DocID Then
                    ExitLoop = True
                Else
                    RstRep.MoveNext
                    If RstRep.EOF Then mAddRec = True
                End If
            Loop

lblRst1MoveNext:
            RecCountFix = RstRep.RecordCount
            RST1.MoveNext
        Next
    End If
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    CreateFieldDefFile RstRep, PubRepoPath + "\" & mRepName & "1.TTX", True
    
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    rpt.OpenSubreport("VehHisDet1").Database.SetDataSource RstRep
    rpt.OpenSubreport("VehHisDet1").ReadRecords
    
    Set Rst = Nothing
    Set RST1 = Nothing
    Set Rst2 = Nothing
    Set RST3 = Nothing
    Set Rst4 = Nothing
    Set RstRep = Nothing
    
    Call Report_View(rpt, Me.CAPTION & "[History]", , False)
    Set rpt = Nothing
Exit Sub
ERRORHANDLER:
    MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub TopCtrl1_eRef()
    RsModel.Requery
    RsCity.Requery
    RsColor.Requery
    RsDeal.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer, Condstr$
    Dim mTrans As Boolean
    Dim DocIdHlp As String
    On Error GoTo errlbl
    
    If IsValid(Txt(Model), "Model") = False Then Exit Sub
    If IsValid(Txt(ChasNo), "Chassis Number") = False Then Exit Sub
    If IsValid(Txt(Engno), "Engine Number") = False Then
    End If
    If Txt(ContPerson) <> "" Then
        If Txt(DOB) <> "" Then
            If Txt(DOM) <> "" Then
                If CDate(Txt(DOB)) > CDate(Txt(DOM)) Then
                    MsgBox "Date of Marriage is less than Date of Birth", vbInformation, "Validation"
                    Txt(DOB).SetFocus: Exit Sub
                End If
            End If
        End If
    Else
        If Txt(DOB) <> "" Or Txt(DOM) <> "" Then
            MsgBox "Please Enter Contact Person", vbInformation, "Validation"
            Txt(ContPerson).SetFocus: Exit Sub
        End If
    End If
    'special case for duplicate cardno checking/generation
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        Dim urs As Recordset
        Set urs = GCn.Execute("select max( " & cVal(cMID("cardno", "3", "6")) & "),max(carddate) from hiscard")
        Txt(SNo).TEXT = PubSiteCode + Right("000000" & IIf(IsNull((urs.Fields(0).Value)), 1, (urs.Fields(0).Value) + 1), 6)
        Set urs = Nothing
    End If
    'eof of special case
    'Invoice date check
    If Txt(InvDate) <> "" Then
            If Txt(SellDate) <> "" Then
                If CDate(Txt(InvDate)) > CDate(Txt(SellDate)) Then
                    MsgBox "Date of Invoice is greater than Date of Sale", vbInformation, "Validation"
                    Txt(InvDate).SetFocus: Exit Sub
                End If
            End If
    End If
    
        'Duplicate Checking Condition for Edit Case
        If TopCtrl1.TopText2.CAPTION <> "Add" Then   'Edit Case
            Condstr = " and CardNo<>'" & Txt(SNo) & "'"
        End If
        'Check Duplicate RegNo
        If Txt(RegNo) <> "" Then
            GSQL = "Select Chassis from HisCard where RegNo='" & Txt(RegNo) & "'"
            GSQL = GSQL & Condstr
            Set GRs = New ADODB.Recordset
            GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
            If GRs.RecordCount > 0 Then
                MsgBox "Registration No. " & Txt(RegNo) & " already exists with Chassis " & GRs!Chassis, vbOKOnly, "Duplicate Checking"
                Set GRs = Nothing: Txt(RegNo).SetFocus: Exit Sub
            End If
            Set GRs = Nothing
        End If
        'Validate duplicate Chassis No.
        If Txt(ChasNo) <> "" Then
            GSQL = "Select RegNo, Engine from HisCard where Chassis='" & Txt(ChasNo) & "'"
            GSQL = GSQL & Condstr
            Set GRs = New ADODB.Recordset
            GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
            If GRs.RecordCount > 0 Then
                MsgBox "Chassis No. " & Txt(ChasNo) & " already exists with " & vbCrLf & _
                    "Reg.No. " & GRs!RegNo & " / Engine No. " & GRs!Engine, vbOKOnly, "Duplicate Checking"
                Set GRs = Nothing: Txt(ChasNo).SetFocus: Exit Sub
            End If
            Set GRs = Nothing
        End If
        'Validate duplicate Engine No.
        If Txt(Engno) <> "" Then
            GSQL = "Select RegNo, Chassis from HisCard where Engine='" & Txt(Engno) & "'"
            GSQL = GSQL & Condstr
            Set GRs = New ADODB.Recordset
            GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
            If GRs.RecordCount > 0 Then
                MsgBox "Engine No. " & Txt(Engno) & " already exists with " & vbCrLf & _
                    "Reg.No. " & GRs!RegNo & " / Chassis No. " & GRs!Chassis, vbOKOnly, "Duplicate Checking"
                Set GRs = Nothing: Txt(Engno).SetFocus: Exit Sub
            End If
            Set GRs = Nothing
        End If

    GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        GSQL = "insert into hiscard(CardNo, Site_Code, Div_Code,CardDate, Model,RegNo, RegDate, CHAS_TYPE, Chassis,Engine , VehSerialNo, " & _
            " Supplier_BillNo, Supplier_BillDate, Dealer_Code, Delivery_Date, CouponNo, GBoxNo, FAxelNo , RAxelNo , SteerGBNo ,CabinNo," & _
            " BodyNo, Colourcode, DoorLockNo , SteerLockNo, FillerCapLockNo , HeadLampMake, FIP_No , Frame_No , Steer_Type , Steer_Make," & _
            " Alternator,StarterMotor,Battery,Brake_Type,Radiator_Make,Tyre_FL,Tyre_FR,Tyre_ML1,Tyre_ML2,Tyre_MR1,Tyre_MR2,Tyre_RL1,Tyre_RL2," & _
            " Tyre_RR1,Tyre_RR2,Spare_Wheel,Fuel,Fuel_Unit,Name,ConPerson,Add1,Add2,CityCode,PhoneOff,Mobile,Mail_ID,DOB,DOM,OwnDrive,Govt_YN," & _
            " Addl_Equp,Locked_Text,VehicleImage,OwnerImage,U_Name,U_EntDt,U_AE," & _
            "DisSprMRP, DisSprTB, DisSprTP, DisOilMRP, DisOilTB, DisOilTP,Varient,VehDet,ExtendWar) " & _
            " values('" & Txt(SNo) & "','" & PubSiteCode & "','" & PubDivCode & "'," & ConvertDate(Txt(SDate)) & ",'" & Txt(Model).Tag & "','" & Txt(RegNo) & "'," & ConvertDate(Txt(RegDate)) & ",'" & Txt(ChasType) & "','" & Txt(ChasNo) & "','" & Txt(Engno) & "','" & Txt(Serial) & _
            "','" & Txt(InvNo) & "'," & ConvertDate(Txt(InvDate)) & ",'" & Txt(DCode) & "'," & ConvertDate(Txt(SellDate)) & ",'" & Txt(ServBookNo) & "','" & Txt(GearBoxNo) & "','" & Txt(FAxleNo) & "', '" & Txt(RAxleNo) & "','" & Txt(SGearBoxNo) & "', '" & Txt(CabNo) & _
            "','" & Txt(BodyNo) & "' , '" & Txt(Paint).Tag & "', '" & Txt(DoorNo) & "',  '" & Txt(SteerNo) & "', '" & Txt(FillCapNo) & "', '" & Txt(HeadLamp) & "', '" & Txt(fipno) & "','" & Txt(FrameNo) & "','" & Txt(SteerType) & "', '" & Txt(SteerMake) & _
            "','" & Txt(Alternator) & "', '" & Txt(StartMotor) & "', '" & Txt(Battery) & "','" & Txt(Brake) & "','" & Txt(Radiator) & "','" & Txt(FrontL) & "','" & Txt(FrontR) & "','" & Txt(MidL1) & "','" & Txt(MidL2) & "','" & Txt(MidR1) & "','" & Txt(MidR2) & "','" & Txt(RearL1) & "','" & Txt(RearL2) & _
            "','" & Txt(RearR1) & "','" & Txt(RearR2) & "','" & Txt(SpWheel) & "','" & Txt(FUEL) & "','" & Txt(FuelUnit) & "','" & Txt(OwName) & "','" & Txt(ContPerson) & "','" & Txt(Add1) & "','" & Txt(Add2) & "','" & Txt(City).Tag & "','" & Txt(Telephone) & "','" & Txt(Mobile) & "','" & Txt(MailID) & "'," & ConvertDate(Txt(DOB)) & "," & ConvertDate(Txt(DOM)) & "," & IIf(Txt(DrivBy) = "Owner", "1", "0") & "," & IIf(Txt(GovtVeh).TEXT = "Yes", "1", "0") & _
            " ,'" & Txt(AddEquip) & "','" & Txt(LockedText) & "','" & Picture1(0).Tag & "','" & Picture1(1).Tag & "','" & pubUName & "', " & ConvertDate(PubServerDate) & _
            ",'A', " & Val(Txt(SprMrp)) & ", " & Val(Txt(SprTB)) & ", " & Val(Txt(SprTP)) & ", " & Val(Txt(OilMRP)) & ", " & Val(Txt(OilTB)) & ", " & Val(Txt(OilTP)) & ",'" & Txt(Varient) & "','" & Txt(VehDet) & "'," & IIf(Txt(ExtendWar) = "Yes", 1, 0) & ")"
        GCn.Execute GSQL
    Else
        GCn.Execute ("update hiscard set Site_Code='" & PubSiteCode & "',CardDate = " & ConvertDate(Txt(SDate)) & ",Model='" & Txt(Model).Tag & "' , RegNo='" & Txt(RegNo) & "', RegDate=" & ConvertDate(Txt(RegDate)) & _
            " ,Chas_Type='" & Txt(ChasType) & "', Chassis='" & Txt(ChasNo) & "', Engine='" & Txt(Engno) & "', VehSerialNo='" & Txt(Serial) & "', Supplier_BillNo='" & Txt(InvNo) & "', Supplier_BillDate =" & ConvertDate(Txt(InvDate)) & ", Dealer_Code = '" & Txt(DCode) & _
            "',Delivery_Date=" & ConvertDate(Txt(SellDate)) & ", CouponNo='" & Txt(ServBookNo) & "', GBoxNo='" & Txt(GearBoxNo) & "', FAxelNo='" & Txt(FAxleNo) & "' , RAxelNo='" & Txt(RAxleNo) & "' , SteerGBNo ='" & Txt(SGearBoxNo) & "',CabinNo='" & Txt(CabNo) & _
            "',BodyNo='" & Txt(BodyNo) & "' , Colourcode='" & Txt(Paint).Tag & "', DoorLockNo='" & Txt(DoorNo) & "' , SteerLockNo='" & Txt(SteerNo) & "', FillerCapLockNo='" & Txt(FillCapNo) & "' , HeadLampMake= '" & Txt(HeadLamp) & "', FIP_No='" & Txt(fipno) & _
            "',Frame_No='" & Txt(FrameNo) & "' , Steer_Type='" & Txt(SteerType) & "' , Steer_Make='" & Txt(SteerMake) & "' ,Alternator='" & Txt(Alternator) & "',StarterMotor='" & Txt(StartMotor) & "',Battery='" & Txt(Battery) & "',Brake_Type='" & Txt(Brake) & _
            "',Radiator_Make='" & Txt(Radiator) & "',Tyre_FL='" & Txt(FrontL) & "',Tyre_FR='" & Txt(FrontR) & "',Tyre_ML1='" & Txt(MidL1) & "',Tyre_ML2='" & Txt(MidL2) & "',Tyre_MR1='" & Txt(MidR1) & "',Tyre_MR2='" & Txt(MidR2) & "',Tyre_RL1='" & Txt(RearL1) & _
            "',Tyre_RL2='" & Txt(RearL2) & "',Tyre_RR1='" & Txt(RearR1) & "',Tyre_RR2='" & Txt(RearR2) & "',Spare_Wheel='" & Txt(SpWheel) & "',Fuel='" & Txt(FUEL) & "',Fuel_Unit='" & Txt(FuelUnit) & "',Name='" & Txt(OwName) & "',ConPerson='" & Txt(ContPerson) & _
            "',Add1='" & Txt(Add1) & "',Add2='" & Txt(Add2) & "',CityCode='" & Txt(City).Tag & "',PhoneOff='" & Txt(Telephone) & "',Mobile='" & Txt(Mobile) & "',Mail_ID='" & Txt(MailID) & "',DOB=" & ConvertDate(Txt(DOB)) & ",DOM=" & ConvertDate(Txt(DOM)) & _
            " ,OwnDrive=" & IIf(Txt(DrivBy) = "Owner", "1", "0") & ",Govt_YN=" & IIf(Txt(GovtVeh).TEXT = "Yes", "1", "0") & ",Addl_Equp='" & Txt(AddEquip) & "',Locked_Text='" & Txt(LockedText) & "',VehicleImage='" & Picture1(1).Tag & "',OwnerImage='" & Picture1(0).Tag & _
            "',U_Name='" & pubUName & "',U_EntDt= " & ConvertDate(PubServerDate) & ",U_AE='E', DisSprMRP=" & Val(Txt(SprMrp)) & ",DisSprTB= " & Val(Txt(SprTB)) & ", DisSprTP=" & Val(Txt(SprTP)) & ", DisOilMRP=" & Val(Txt(OilMRP)) & ", DisOilTB=" & Val(Txt(OilTB)) & _
            ", DisOilTP=" & Val(Txt(OilTP)) & ",Varient ='" & Txt(Varient) & "',VehDet ='" & Txt(VehDet) & "',ExtendWar= " & IIf(Txt(ExtendWar) = "Yes", 1, 0) & " where CardNo='" & Txt(SNo) & "'")
    End If
    
    GCn.Execute "Update Model set Chas_Type='" & left(Trim(Txt(ChasNo)), 6) & "' Where Model='" & Txt(Model) & "' And Chas_Type=''"
    
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select CardNO as SearchCode from Hiscard where div_code='" & PubDivCode & "' And CardNO  = '" & Txt(SNo) & "' Order by CardNo")
    End If
    Master.FIND "SearchCode = '" & Txt(SNo) & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Select Case Index
    Case Model
        If RsModel.RecordCount = 0 Or (RsDeal.EOF Or RsDeal.BOF) Then Exit Sub
        RsModel.Sort = "Code"
        If Txt(Index) <> "" Then
            RsModel.MoveFirst
            RsModel.FIND "code ='" & Txt(Index) & "'"
            If RsModel.EOF Then RsModel.MoveFirst
        End If
    Case DCode
        If RsDeal.RecordCount = 0 Or (RsDeal.EOF = True Or RsDeal.BOF = True) Then Exit Sub
        RsDeal.Sort = "CODE"
        If Txt(Index) <> "" Then
            RsDeal.MoveFirst
            RsDeal.FIND "code ='" & Txt(Index) & "'"
            If RsDeal.EOF = True Then RsDeal.MoveFirst
        End If
    Case DNAME
        If RsDeal.RecordCount = 0 Or (RsDeal.EOF = True Or RsDeal.BOF = True) Then Exit Sub
        RsDeal.Sort = "Name"
        If Txt(Index) <> "" Then
            RsDeal.MoveFirst
            RsDeal.FIND "Name ='" & Txt(Index) & "'"
            If RsDeal.EOF = True Then RsDeal.MoveFirst
        End If
    Case 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37
       SSTab1.Tab = 0
    Case 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48
       SSTab1.Tab = 1
    Case 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61
       SSTab1.Tab = 2
    Case ChasNo
        Txt(Index).SelStart = Len(Txt(Index))
    Case Brake
        ListArray = Array("DAOH", "S - CAM", "OTHER")
        Set mListItem = ListView_Items(ListView, Txt, Brake, ListArray, 3)
    Case VehDet
        ListArray = Array("Personal", "Taxi", "Corporate")
        Set mListItem = ListView_Items(ListView, Txt, VehDet, ListArray, 3)
End Select
Ctrl_GetFocus Txt(Index)
Grid_Hide
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Byte
Dim Txtdate As Boolean
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case SNo
        'SiteCode Edit restricted
        KeyCode = RestrictKey(1, KeyCode, Txt(Index), Shift)
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        End If
    Case ChasNo
        KeyCode = RestrictKey(Len(Txt(ChasType)), KeyCode, Txt(Index), Shift)
    Case DCode, DNAME
        DGridTxtKeyDown DGDeal, Txt, Index, RsDeal, KeyCode, False, 1, frmDealer, "frmDealer"
    Case Brake
        ListView_KeyDown FrmList, ListView, Txt, Brake, KeyCode, Shift, Txt(Index).left, Txt(Index).top + Txt(Index).height, Txt(Index).width, FrmList.height
    Case Model
        DGridTxtKeyDown DgModel, Txt, Model, RsModel, KeyCode, False, 0, frmModel, "frmModel"
    Case Paint
        DGridTxtKeyDown DGColor, Txt, Paint, RsColor, KeyCode, False, 1, frmColor, "frmColor"
    Case City
        DGridTxtKeyDown DGCity, Txt, City, RsCity, KeyCode, False, 1, frmCity, "frmCity"
    Case VehDet
        ListView_KeyDown FrmList, ListView, Txt, VehDet, KeyCode, Shift, Txt(Index).left, FrmList.top, Txt(Index).width, FrmList.height
'    Case SprMrp, SprTB, SprTP, OilMRP, OilTB, OilTP
'            NumDown txt(Index), KeyCode, 5, 2
End Select
If FrmList.Visible = False And DgModel.Visible = False And DGCity.Visible = False And DGDeal.Visible = False And DGColor.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> DOM Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = DOM Then
        Txt_Validate Index, False
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> SNo Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> SNo Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If

End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case ChasNo
        KeyAscii = RestrictKey(Len(Txt(ChasType)), KeyAscii, Txt(Index), 0)
    Case SNo
        KeyAscii = RestrictKey(1, KeyAscii, Txt(Index), 0)
    Case Model
        If DgModel.Visible = True Then DGridTxtKeyPress Txt, Model, RsModel, KeyAscii, "CODE"
    Case DCode
        If DGDeal.Visible = True Then DGridTxtKeyPress Txt, Index, RsDeal, KeyAscii, "Code" ': Txt(DCode) = DGDeal.Text
    Case DNAME
        If DGDeal.Visible = True Then DGridTxtKeyPress Txt, Index, RsDeal, KeyAscii, "Name" ': Txt(DCode) = DGDeal.Text
    Case Paint
        If DGColor.Visible = True Then DGridTxtKeyPress Txt, Paint, RsColor, KeyAscii, "Name"
    Case City
        If DGCity.Visible = True Then DGridTxtKeyPress Txt, City, RsCity, KeyAscii, "Name"
    Case SprMrp, SprTB, SprTP, OilMRP, OilTB, OilTP
        Call NumPress(Txt(Index), KeyAscii, 2, 2)
    Case ExtendWar
        If KeyAscii = Asc("Y") Or KeyAscii = Asc("y") Then
            Txt(Index) = "Yes"
        ElseIf KeyAscii = Asc("N") Or KeyAscii = Asc("n") Then
            Txt(Index) = "No"
        End If
        KeyAscii = 0
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case Brake
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
'    Case Model
'        If DGModel.Visible = True Then DGridTxtKeyUp Txt, Model, RsModel, KeyCode, "Name"
'    Case DNAME
'        If DGDeal.Visible = True Then DGridTxtKeyUp Txt, DNAME, RsDeal, KeyCode, "Name": Txt(DCode) = DGDeal.Text
'    Case Paint
'        If DGColor.Visible = True Then DGridTxtKeyUp Txt, Paint, RsColor, KeyCode, "Name"
'    Case City
'        If DGCity.Visible = True Then DGridTxtKeyUp Txt, City, RsCity, KeyCode, "Name"
    Case GovtVeh
        If Len(Txt(Index)) = 0 Or UCase(mID(Txt(Index), 1, 1)) = "N" Then
            Txt(Index) = "No"
        ElseIf UCase(mID(Txt(Index), 1, 1)) = "Y" Then
            Txt(Index) = "Yes"
        Else
            Txt(Index) = "No"
        End If
    Case DrivBy
        If Len(Txt(Index)) = 0 Or UCase(mID(Txt(Index), 1, 1)) = "D" Then
            Txt(Index) = "Driver"
        ElseIf UCase(mID(Txt(Index), 1, 1)) = "O" Then
            Txt(Index) = "Owner"
        Else
            Txt(Index) = "Owner"
        End If
    Case SteerType
        If Len(Txt(Index)) = 0 Or UCase(mID(Txt(Index), 1, 1)) = "P" Then
            Txt(Index) = "Power"
        ElseIf UCase(mID(Txt(Index), 1, 1)) = "M" Then
            Txt(Index) = "Manual"
        Else
            Txt(Index) = "Manual"
        End If
        If Txt(SteerType) = "Power" Then
            Txt(SteerMake).Enabled = True
        Else
            Txt(SteerMake).Enabled = False
        End If
    Case VehDet
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim xValue$
Select Case Index
    Case Model
        If RsModel.RecordCount = 0 Or (RsModel.EOF = True Or RsModel.BOF = True) Or Txt(Index) = "" Then
            Txt(ChasType) = ""
            Txt(ChasNo) = ""
        Else
            Txt(ChasType) = IIf(IsNull(RsModel!Chas_Type), "", RsModel!Chas_Type)
            If Trim(Txt(ChasNo)) = "" Then
               Txt(ChasNo) = RsModel!Chas_Type
            ElseIf left(Txt(ChasNo), 6) <> RsModel!Chas_Type Then
               Txt(ChasNo) = RsModel!Chas_Type
            End If
        End If
    Case DCode
        If RsDeal.RecordCount = 0 Or (RsDeal.EOF = True Or RsDeal.BOF = True) Or Txt(Index) = "" Then
            Txt(DCode) = ""
            Txt(DNAME) = ""
            Txt(DNAME).Enabled = True
        Else
            Txt(DCode) = IIf(IsNull(RsDeal!Code), "", RsDeal!Code)
            Txt(DNAME) = IIf(IsNull(RsDeal!Name), "", RsDeal!Name)
            Txt(DNAME).Enabled = False
        End If
    Case DNAME
        If RsDeal.RecordCount = 0 Or (RsDeal.EOF = True Or RsDeal.BOF = True) Or Txt(Index) = "" Then
            Txt(DCode) = ""
            Txt(DNAME) = ""
        Else
            Txt(DCode) = IIf(IsNull(RsDeal!Code), "", RsDeal!Code)
            Txt(DNAME) = IIf(IsNull(RsDeal!Name), "", RsDeal!Name)
        End If
    Case SDate, RegDate, SellDate, InvDate, DOB, DOM, LastJobDt
        If Len(Trim(Txt(Index))) > 0 Then
            Txt(Index) = RetDate(Txt(Index))
        End If
End Select
End Sub

Private Sub DGCity_Click()
    DGCity.Visible = False
    If RsCity.RecordCount > 0 Then
        Txt(City).TEXT = RsCity!Name
        Txt(City).Tag = RsCity!Code
    End If
    Txt(City).SetFocus
End Sub

Private Sub DGColor_Click()
    DGColor.Visible = False
    If RsColor.RecordCount > 0 Then
        Txt(Paint).TEXT = RsColor!Name
        Txt(Paint).Tag = RsColor!Code
    End If
    Txt(Paint).SetFocus
End Sub

Private Sub DGDeal_Click()
    DGDeal.Visible = False
    If RsDeal.RecordCount > 0 Then
        Txt(DNAME).TEXT = RsDeal!Name
        Txt(DNAME).Tag = RsDeal!Code
        Txt(DCode).TEXT = RsDeal!Code
    End If
    If Txt(DNAME).Enabled Then
        Txt(DNAME).SetFocus
    Else
        Txt(DCode).SetFocus
    End If
End Sub

Private Sub DGModel_Click()
    DgModel.Visible = False
    If RsModel.RecordCount > 0 Then
        Txt(Model).TEXT = RsModel!Name
        Txt(Model).Tag = RsModel!Code
        Txt(ChasType) = RsModel!Chas_Type
        If Trim(Txt(ChasNo)) = "" Then
           Txt(ChasNo) = RsModel!Chas_Type
        ElseIf left(Txt(ChasNo), 6) <> RsModel!Chas_Type Then
           Txt(ChasNo) = RsModel!Chas_Type
        End If
    End If
    Txt(Model).SetFocus
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To 73
    Txt(I).TEXT = ""
Next I
Picture1(0) = Nothing
Picture1(1) = Nothing
End Sub
Private Sub MoveRec()
On Error GoTo error1
Dim Master1 As ADODB.Recordset
If Master.RecordCount > 0 Then
    Set Master1 = New Recordset
    Master1.CursorLocation = adUseClient
    Master1.Open "select hiscard.*,MODEL.CHAS_TYPE AS CHTYPE,colmast.col_desc as colorname, amd_dealer.d_code as dcode,amd_dealer.d_name as dname,city.cityname as cityname from (((hiscard left join model on hiscard.model=model.model)left join colmast on hiscard.colourcode=colmast.col_code)left join amd_dealer on hiscard.dealer_code=amd_dealer.d_code)left join city on hiscard.citycode=city.citycode where hiscard.CardNo='" & Master!SearchCode & "'", GCn, adOpenDynamic, adLockOptimistic

    Txt(SNo) = Master1!CardNo 'Mid(Master1!cardno, 2, 6)  '0
    Txt(SDate) = Master1!CardDATE
    Txt(Model) = IIf(IsNull(Master1!Model), "", Master1!Model)
    Txt(Model).Tag = IIf(IsNull(Master1!Model), "", Master1!Model)
    Txt(ChasType) = IIf(IsNull(Master1!CHTYPE), "", Master1!CHTYPE)
    Txt(ChasNo) = IIf(IsNull(Master1!Chassis), "", Master1!Chassis)
    Txt(Engno) = IIf(IsNull(Master1!Engine), "", Master1!Engine)
    Txt(DNAME) = IIf(IsNull(Master1!DNAME), "", Master1!DNAME)
    Txt(DCode) = IIf(IsNull(Master1!DCode), "", Master1!DCode)
    Txt(SellDate) = IIf(IsNull(Master1!Delivery_Date), "", Master1!Delivery_Date)
    Txt(InvNo) = IIf(IsNull(Master1!SUPPLIER_BILLNO), "", Master1!SUPPLIER_BILLNO)
    Txt(OwName) = IIf(IsNull(Master1!Name), "", Master1!Name)
    Txt(RegNo) = IIf(IsNull(Master1!RegNo), "", Master1!RegNo) '10
    Txt(RegDate) = IIf(IsNull(Master1!RegDate), "", Master1!RegDate)
    Txt(Serial) = IIf(IsNull(Master1!VehSerialNo), "", Master1!VehSerialNo)
    Txt(InvDate) = IIf(IsNull(Master1!Supplier_BillDate), "", Master1!Supplier_BillDate)
    Txt(GovtVeh) = IIf(Master1!Govt_YN = "1", "Yes", "No")
    Txt(ServBookNo) = IIf(IsNull(Master1!CouponNo), "", Master1!CouponNo)
    Txt(GearBoxNo) = IIf(IsNull(Master1!GBoxNo), "", Master1!GBoxNo)
    Txt(FAxleNo) = IIf(IsNull(Master1!FAxelNo), "", Master1!FAxelNo)
    Txt(RAxleNo) = IIf(IsNull(Master1!RAxelNo), "", Master1!RAxelNo)
    Txt(CabNo) = IIf(IsNull(Master1!cabinno), "", Master1!cabinno)
    Txt(BodyNo) = IIf(IsNull(Master1!BodyNo), "", Master1!BodyNo) '20
    Txt(Paint) = IIf(IsNull(Master1!Colourcode), "", Master1!Colourcode)
    Txt(Paint).Tag = IIf(IsNull(Master1!Colourcode), "", Master1!Colourcode)
    Txt(DoorNo) = IIf(IsNull(Master1!DoorLockNo), "", Master1!DoorLockNo)
    Txt(SGearBoxNo) = IIf(IsNull(Master1!SteerGBNo), "", Master1!SteerGBNo)
    Txt(SteerNo) = IIf(IsNull(Master1!SteerLockNo), "", Master1!SteerLockNo)
    Txt(SteerType) = IIf(IsNull(Master1!Steer_Type), "", Master1!Steer_Type)
    Txt(SteerMake) = IIf(IsNull(Master1!Steer_Make), "", Master1!Steer_Make)
    Txt(FillCapNo) = IIf(IsNull(Master1!FillerCapLockNo), "", Master1!FillerCapLockNo)
    Txt(HeadLamp) = IIf(IsNull(Master1!HeadLampMake), "", Master1!HeadLampMake)
    Txt(fipno) = IIf(IsNull(Master1!FIP_No), "", Master1!FIP_No)
    Txt(FrameNo) = IIf(IsNull(Master1!Frame_No), "", Master1!Frame_No) '30
    Txt(Alternator) = IIf(IsNull(Master1!Alternator), "", Master1!Alternator)
    Txt(StartMotor) = IIf(IsNull(Master1!StarterMotor), "", Master1!StarterMotor)
    Txt(Battery) = IIf(IsNull(Master1!Battery), "", Master1!Battery)
    Txt(Radiator) = IIf(IsNull(Master1!Radiator_Make), "", Master1!Radiator_Make)
    Txt(FUEL) = IIf(IsNull(Master1!FUEL), "", Master1!FUEL)
    Txt(FuelUnit) = IIf(IsNull(Master1!Fuel_Unit), "", Master1!Fuel_Unit)
    Txt(AddEquip) = IIf(IsNull(Master1!Addl_Equp), "", Master1!Addl_Equp)
    Txt(FrontL) = IIf(IsNull(Master1!Tyre_FL), "", Master1!Tyre_FL)
    Txt(FrontR) = IIf(IsNull(Master1!Tyre_FR), "", Master1!Tyre_FR)
    Txt(MidL1) = IIf(IsNull(Master1!Tyre_ML1), "", Master1!Tyre_ML1) '40
    Txt(MidR1) = IIf(IsNull(Master1!Tyre_MR1), "", Master1!Tyre_MR1)
    Txt(MidL2) = IIf(IsNull(Master1!Tyre_ML2), "", Master1!Tyre_ML2)
    Txt(MidR2) = IIf(IsNull(Master1!Tyre_MR2), "", Master1!Tyre_MR2)
    Txt(RearL1) = IIf(IsNull(Master1!Tyre_RL1), "", Master1!Tyre_RL1)
    Txt(RearR1) = IIf(IsNull(Master1!Tyre_RR1), "", Master1!Tyre_RR1)
    Txt(RearL2) = IIf(IsNull(Master1!Tyre_RL2), "", Master1!Tyre_RL2)
    Txt(RearR2) = IIf(IsNull(Master1!Tyre_RR2), "", Master1!Tyre_RR2)
    Txt(SpWheel) = IIf(IsNull(Master1!Spare_Wheel), "", Master1!Spare_Wheel)
    Txt(DrivBy) = IIf(Master1!OwnDrive = "1", "Owner", "Driver")
    Txt(Add1) = IIf(IsNull(Master1!Add1), "", Master1!Add1) '50
    Txt(Add2) = IIf(IsNull(Master1!Add2), "", Master1!Add2)
    Txt(City) = IIf(IsNull(Master1!CityName), "", Master1!CityName)
    Txt(City).Tag = IIf(IsNull(Master1!CityCode), "", Master1!CityCode)
    Txt(Telephone) = IIf(IsNull(Master1!PhoneOff), "", Master1!PhoneOff)
    Txt(Mobile) = IIf(IsNull(Master1!Mobile), "", Master1!Mobile)
    Txt(MailID) = IIf(IsNull(Master1!Mail_ID), "", Master1!Mail_ID)
    Txt(DOB) = IIf(IsNull(Master1!DOB), "", Master1!DOB)
    Txt(DOM) = IIf(IsNull(Master1!DOM), "", Master1!DOM)
    Txt(Brake) = IIf(IsNull(Master1!Brake_Type), "", Master1!Brake_Type)
    Txt(ContPerson) = IIf(IsNull(Master1!ConPerson), "", Master1!ConPerson)
    Txt(LastJobNo) = IIf(IsNull(Master1!ljob_docID), "", Master1!ljob_docID)
    Txt(LastJobDt) = IIf(IsNull(Master1!ljob_date), "", Master1!ljob_date) '60
    Txt(LastKiloM) = IIf(IsNull(Master1!ljob_atkmshrs), "", Master1!ljob_atkmshrs)
    Txt(LockedText) = IIf(IsNull(Master1!Locked_Text), "", Master1!Locked_Text)
    Txt(SprMrp) = IIf(IsNull(Master1!DisSprMRP), "", Format(Master1!DisSprMRP, "0.00"))
    Txt(SprTB) = IIf(IsNull(Master1!DisSprTB), "", Format(Master1!DisSprTB, "0.00"))
    Txt(SprTP) = IIf(IsNull(Master1!DisSprTP), "", Format(Master1!DisSprTP, "0.00"))
    
    Txt(OilMRP) = IIf(IsNull(Master1!DisOilMRP), "", Format(Master1!DisOilMRP, "0.00"))
    Txt(OilTB) = IIf(IsNull(Master1!DisOilTB), "", Format(Master1!DisOilTB, "0.00"))
    Txt(OilTP) = IIf(IsNull(Master1!DisOilTP), "", Format(Master1!DisOilTP, "0.00"))
    
    Txt(Varient) = XNull(Master1!Varient)
    Txt(VehDet) = XNull(Master1!VehDet)
    Txt(ExtendWar) = IIf(Master1!ExtendWar = 1, "Yes", "No")
    
'    If IsNull(Master1!OWNERIMAGE) Or Trim(Master1!OWNERIMAGE) = "" Then
'        Picture1(0).Picture = LoadPicture(Pub_DataPath & "\Pictures\Owner" & "\nopicture.bmp")
'        Picture1(0).Tag = "nopicture.bmp"
'    Else
'        Picture1(0).Picture = LoadPicture(Pub_DataPath & "\Pictures\Owner" & "\" & Master1!OWNERIMAGE)
'        Picture1(0).Tag = Master1!OWNERIMAGE
'    End If
'
'    If IsNull(Master1!VEHICLEIMAGE) Or Trim(Master1!VEHICLEIMAGE) = "" Then
'        Picture1(1).Picture = LoadPicture(Pub_DataPath & "\Pictures\Vehicle" & "\nopicture.bmp")
'        Picture1(1).Tag = "nopicture.bmp"
'    Else
'        Picture1(1).Picture = LoadPicture(Pub_DataPath & "\Pictures\Vehicle" & "\" & Master1!VEHICLEIMAGE)
'        Picture1(1).Tag = Master1!VEHICLEIMAGE
'    End If
End If
Set Master1 = Nothing
Grid_Hide
TopCtrl1.tDel = False
Exit Sub
error1:
    If err.NUMBER = 53 Then Picture1(0).Picture = LoadPicture(Pub_DataPath & "\Pictures\Owner" & "\nopicture.bmp"): Resume Next
        CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To 73
    Txt(I).Enabled = Enb
Next
'Txt(DCode).Enabled = False
Txt(SteerMake).Enabled = False
Txt(ChasType).Enabled = False
Txt(LastJobNo).Enabled = False
Txt(LastJobDt).Enabled = False
Txt(LastKiloM).Enabled = False

'For i = 13 To 25
'    Txt(i).Enabled = Enb
'Next
End Sub
Private Sub Grid_Hide()
    If DgModel.Visible = True Then DgModel.Visible = False
    If DGCity.Visible = True Then DGCity.Visible = False
    If DGDeal.Visible = True Then DGDeal.Visible = False
    If DGColor.Visible = True Then DGColor.Visible = False
End Sub


Private Sub Ini_Grid()
'    DGCity.top = mTopScale: DGCity.left = 7155
'    DGDeal.top = mTopScale: DGDeal.left = 7155
'    DGColor.top = mTopScale: DGColor.left = 7155
'    DGModel.top = 2805: DGModel.left =  165
    
    DGColor.left = Me.width - (DGColor.width + mRtScale): DGColor.top = mTopScale: DGColor.height = Me.height - (mTopScale + mBotScale)
    DGCity.left = Me.width - (DGCity.width + mRtScale): DGCity.top = mTopScale: DGCity.height = Me.height - (mTopScale + mBotScale)
    DGDeal.left = Me.width - (DGDeal.width + mRtScale): DGDeal.top = mTopScale: DGDeal.height = Me.height - (mTopScale + mBotScale)
    DgModel.top = SSTab1.top: DgModel.left = SSTab1.left: DgModel.width = SSTab1.width: DgModel.height = SSTab1.height
   SSTab1.Tab = 0

End Sub
