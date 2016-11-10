VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmEmpMast 
   Caption         =   "Employee Master"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8145
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
   ScaleHeight     =   8070
   ScaleWidth      =   8145
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DgSupervisor 
      Height          =   1845
      Left            =   3600
      Negotiate       =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   5550
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   3254
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
         Caption         =   "Supervisor"
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
         Caption         =   "Supervisor"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   25
      Left            =   2355
      MaxLength       =   20
      TabIndex        =   17
      Top             =   3675
      Width           =   5445
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   24
      Left            =   2355
      MaxLength       =   8
      TabIndex        =   54
      Top             =   5115
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DGDIV 
      Height          =   1845
      Left            =   -1095
      Negotiate       =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   6390
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   3254
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
         DataField       =   "Div_Code"
         Caption         =   "Div_Code"
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
         Caption         =   "Working Division"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   23
      Left            =   2355
      TabIndex        =   2
      Top             =   795
      Width           =   5445
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   4935
      Left            =   -1755
      Negotiate       =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7335
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGEmp 
      Height          =   3330
      Left            =   -8910
      Negotiate       =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   7455
      Visible         =   0   'False
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   5874
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Employee Name"
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
         DataField       =   "ncode"
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
         DataField       =   "fname"
         Caption         =   "Father Name"
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
            ColumnWidth     =   4440.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1665.071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   5054.74
         EndProperty
      EndProperty
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   22
      Left            =   5910
      MaxLength       =   11
      TabIndex        =   25
      Top             =   5115
      Width           =   1890
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   17
      Left            =   2355
      MaxLength       =   50
      TabIndex        =   20
      Top             =   4155
      Width           =   5445
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   18
      Left            =   2355
      MaxLength       =   50
      TabIndex        =   21
      Top             =   4395
      Width           =   5445
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   19
      Left            =   2355
      MaxLength       =   50
      TabIndex        =   22
      Top             =   4635
      Width           =   5445
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   20
      Left            =   2355
      MaxLength       =   11
      TabIndex        =   23
      Top             =   4875
      Width           =   1500
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   21
      Left            =   5910
      MaxLength       =   8
      TabIndex        =   24
      Top             =   4875
      Width           =   1890
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   4215
      TabIndex        =   26
      Top             =   6960
      Visible         =   0   'False
      Width           =   2010
      Begin MSComctlLib.ListView ListView 
         Height          =   1815
         Left            =   0
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   0
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   3201
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   16
      Left            =   5910
      MaxLength       =   7
      TabIndex        =   19
      Top             =   3915
      Width           =   1890
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   15
      Left            =   2355
      MaxLength       =   12
      TabIndex        =   18
      Top             =   3915
      Width           =   1890
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   14
      Left            =   2355
      MaxLength       =   20
      TabIndex        =   16
      Text            =   "16"
      Top             =   3435
      Width           =   5445
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   13
      Left            =   2355
      MaxLength       =   11
      TabIndex        =   15
      Top             =   3195
      Width           =   2265
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
      ForeColor       =   &H00000000&
      Height          =   210
      IMEMode         =   3  'DISABLE
      Index           =   12
      Left            =   6300
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   14
      Top             =   2955
      Width           =   1500
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
      ForeColor       =   &H00000000&
      Height          =   210
      IMEMode         =   3  'DISABLE
      Index           =   11
      Left            =   2355
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   2955
      Width           =   2265
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   10
      Left            =   7335
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2715
      Width           =   465
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   9
      Left            =   2355
      MaxLength       =   11
      TabIndex        =   11
      Top             =   2715
      Width           =   2265
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   8
      Left            =   2355
      MaxLength       =   50
      TabIndex        =   10
      Top             =   2475
      Width           =   5445
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   7
      Left            =   5445
      MaxLength       =   14
      TabIndex        =   9
      Top             =   2235
      Width           =   2355
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   661
      tAdd            =   0   'False
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   2355
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1515
      Width           =   5445
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   2355
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1035
      Width           =   5445
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   6
      Left            =   2355
      MaxLength       =   12
      TabIndex        =   8
      Top             =   2235
      Width           =   2265
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   5
      Left            =   2355
      MaxLength       =   14
      TabIndex        =   7
      Top             =   1995
      Width           =   2265
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   4
      Left            =   2355
      MaxLength       =   25
      TabIndex        =   6
      Top             =   1755
      Width           =   5445
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   2355
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1275
      Width           =   5445
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2355
      MaxLength       =   40
      TabIndex        =   1
      Top             =   555
      Width           =   5445
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation*"
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
      Left            =   255
      TabIndex        =   56
      Top             =   3465
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Over Time Rate/Hr."
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
      Left            =   240
      TabIndex        =   55
      Top             =   5130
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Working in Division*"
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
      Left            =   255
      TabIndex        =   52
      Top             =   810
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
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
      Left            =   4680
      TabIndex        =   50
      Top             =   2970
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name*"
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
      Left            =   255
      TabIndex        =   49
      Top             =   570
      Width           =   1500
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
      Index           =   58
      Left            =   255
      TabIndex        =   48
      Top             =   1290
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      Left            =   255
      TabIndex        =   47
      Top             =   2010
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pager"
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
      Left            =   4875
      TabIndex        =   46
      Top             =   2250
      Width           =   495
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
      Index           =   65
      Left            =   255
      TabIndex        =   45
      Top             =   2250
      Width           =   540
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
      Index           =   42
      Left            =   255
      TabIndex        =   44
      Top             =   1770
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S/O / W/O Name"
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
      Left            =   255
      TabIndex        =   43
      Top             =   1050
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Access Password"
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
      Left            =   255
      TabIndex        =   42
      Top             =   2970
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Status*"
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
      Left            =   255
      TabIndex        =   41
      Top             =   3930
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervisor (Team)"
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
      Left            =   255
      TabIndex        =   40
      Top             =   3720
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Type*"
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
      Left            =   255
      TabIndex        =   39
      Top             =   3210
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      Left            =   4680
      TabIndex        =   38
      Top             =   2730
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
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
      Left            =   255
      TabIndex        =   37
      Top             =   2490
      Width           =   870
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
      Index           =   7
      Left            =   255
      TabIndex        =   36
      Top             =   2730
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Course*"
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
      Left            =   4305
      TabIndex        =   35
      Top             =   3945
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly Salary"
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
      Left            =   4305
      TabIndex        =   34
      Top             =   4890
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details of Training"
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
      Left            =   255
      TabIndex        =   33
      Top             =   4410
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Academic Qualification "
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
      Left            =   255
      TabIndex        =   32
      Top             =   4650
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Joining Date"
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
      Left            =   255
      TabIndex        =   31
      Top             =   4890
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Working Experience"
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
      Left            =   255
      TabIndex        =   30
      Top             =   4170
      Width           =   1710
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leaving Date"
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
      Left            =   4305
      TabIndex        =   29
      Top             =   5115
      Width           =   1125
   End
End
Attribute VB_Name = "frmEmpMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean

Dim RsHelp As ADODB.Recordset
Dim RsCity As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim RsDiv As ADODB.Recordset
Dim RsSupervisor As ADODB.Recordset



Dim urec As Integer
Dim EName As String
Private Const EmpName As Byte = 0
Private Const FathName As Byte = 1
Private Const Add1 As Byte = 2
Private Const Add2 As Byte = 3
Private Const City As Byte = 4
Private Const Phone As Byte = 5
Private Const Pager As Byte = 6
Private Const Mobile As Byte = 7
Private Const Referen As Byte = 8
Private Const DOB As Byte = 9
Private Const Age As Byte = 10
Private Const AccPW As Byte = 11
Private Const ConPW As Byte = 12
Private Const EmpType As Byte = 13
Private Const Desig As Byte = 14
Private Const Status As Byte = 15
Private Const Course As Byte = 16
Private Const Experience As Byte = 17
Private Const Traning As Byte = 18
Private Const Qualific As Byte = 19
Private Const DOJ As Byte = 20
Private Const Salary As Byte = 21
Private Const DOL As Byte = 22
Private Const DIV As Byte = 23
Private Const OTRate As Byte = 24
Private Const Supervisor As Byte = 25
Dim PreName As String
Dim ListArray As Variant
Dim mListItem As ListItem
Dim TAddMode As Boolean
Dim FirmAddFlag As Byte
Dim GridKey As Integer
Dim ExitCtrl As Boolean
Private Sub DGCity_Click()
    DGCity.Visible = False
    If RsCity.RecordCount > 0 Then
        Txt(City).TEXT = RsCity!Name
        Txt(City).Tag = RsCity!Code
    End If
    Txt(City).SetFocus
End Sub
Private Sub DGSupervisor_Click()
    DgSupervisor.Visible = False
    If RsSupervisor.RecordCount > 0 Then
        Txt(Supervisor).TEXT = RsSupervisor!Name
        Txt(Supervisor).Tag = RsSupervisor!Code
    End If
    Txt(Supervisor).SetFocus
End Sub

Private Sub DGdiv_Click()
    DGDIV.Visible = False
    If RsDiv.RecordCount > 0 Then
        Txt(DIV).TEXT = RsDiv!Div_Name
        Txt(DIV).Tag = RsDiv!Div_Code
    End If
    Txt(DIV).SetFocus
End Sub
Private Sub DGEmp_Click()
    DGEmp.Visible = False
    If RsHelp.RecordCount > 0 Then
        Txt(EmpName).TEXT = RsHelp!Name
        'Txt(City).Tag = RsCity!Code
    End If
    Txt(EmpName).SetFocus
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
Dim I As Byte
    WinSetting Me, 6210, 8265
    TopCtrl1.Tag = PubUParam: TopCtrl1.TopText1 = Me.CAPTION
    For I = 0 To 22
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
    DGEmp.Columns(0).width = 4440.189
    DGEmp.Columns(1).width = 750.0473
    DGEmp.Columns(2).width = 5054.74
    DGCity.left = Me.width - (DGCity.width + mRtScale): DGCity.top = mTopScale
    DGEmp.width = Me.width - 90: DGEmp.left = Me.left: DGEmp.top = Me.height - (DGEmp.height + mBotScale)
    DGDIV.left = Txt(DIV).left
    DGDIV.top = Txt(DIV).height + Txt(DIV).top
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
'    Master.Open "select Emp_Mast.*, Emp_Mast.Emp_Code as SearchCode from Emp_Mast where site_code='" & PubSiteCode & "'" & " order by emp_name", GCn, adOpenDynamic, adLockOptimistic

    Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
     sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
    sitecond = ""
    End If
    
    If PubMoveRecYn Then
        Master.Open "select Emp_Mast.*, Emp_Mast.Emp_Code as SearchCode from Emp_Mast " & sitecond & " Order by Emp_Name", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 Emp_Mast.*, Emp_Mast.Emp_Code as SearchCode from Emp_Mast " & sitecond & " Order by Emp_Name", GCn, adOpenDynamic, adLockOptimistic
    End If
    Set RsHelp = New ADODB.Recordset
    RsHelp.CursorLocation = adUseClient
    RsHelp.Open "select emp_name as name, emp_code as ncode, FName from emp_mast order by emp_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGEmp.DataSource = RsHelp
    Set RsCity = New ADODB.Recordset
    With RsCity
        .CursorLocation = adUseClient
        .Open "SELECT CITYCODE as code,CITYNAME as name from CITY  order by CITYNAME", GCn, adOpenDynamic, adLockOptimistic
    End With
    'Working Division Active By Rahul 10-04-2003 At Udaipur (UN Automobiles)
    Set RsDiv = New ADODB.Recordset
    RsDiv.CursorLocation = adUseClient
    RsDiv.Open "select Div_Code, Div_Name from Division order by Div_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGDIV.DataSource = RsDiv

    Set RsSupervisor = GCn.Execute("Select Emp_Code as Code, Emp_Name As Name From Emp_Mast Order by Emp_Name")
    Set DgSupervisor.DataSource = RsSupervisor

    Set DGCity.DataSource = RsCity
'    If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub

ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form_Unload (-1)
End Sub
Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set Master = Nothing
    Set RsHelp = Nothing
    Set RsCity = Nothing
    Set RsDiv = Nothing
End Sub
Private Sub ListView_Click()
Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
FrmList.Visible = False
Txt(Val(ListView.Tag)).SetFocus

'If ListView.ListItems.count  > 0 Then txt(Val(ListView.Tag)).Text = ListView.SelectedItem.Text
'FrmList.Visible = False
'txt(Val(ListView.Tag)).SetFocus
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    Txt(EmpName).SetFocus
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
Exit Sub
Dim mo As Integer, XBM
On Error GoTo eloop1
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
        XBM = Master.Bookmark
        GCn.Execute ("delete from emp_mast where emp_code = '" & Txt(EmpName).Tag & "'")
        GCn.CommitTrans
        Master.Requery
        RsHelp.Requery
        If Master.RecordCount >= XBM Then
            Master.Bookmark = XBM
        Else
            If Master.EOF = False Then Master.MoveLast
        End If
        Call MoveRec
        BUTTONS True, Me, Master, 0
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
    Txt(EmpName).SetFocus
    EName = Txt(EmpName).TEXT
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
    
    Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
     sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
    sitecond = ""
    End If
    
    
    GSQL = "select Emp_Code as SearchCode,Emp_Name,FName,ADD1,ADD2,CityName from Emp_Mast " & sitecond & " Order by Emp_Name"
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
        Set Master = GCn.Execute("select Emp_Mast.*, Emp_Mast.Emp_Code as SearchCode from Emp_Mast Where Emp_Mast.Emp_Code = '" & MyValue & "' Order by Emp_Name")
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
    For I = 0 To 22
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
Dim I As Integer, mQry$, mRepName$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER


    mRepName = "EmpCard"
    If PubBackEnd = "A" Then
        mQry = "Select Emp_Code,Site_Code,Emp_Name,switch(Emp_Type=0,'Sales Staff',Emp_Type=1,'Workshop',Emp_Type=2,'Others') as EmpType," & _
            " FName,ADD1,ADD2,CityName,PHONE,PAGER,Mobile,Reference,DOB,AGE,Access_PWD,Designation," & _
            " Switch(ServStatus='0','Skilled',ServStatus='1','Semi-Skilled',ServStatus='2','Un-Skilled') as SrvStatus," & _
            " Switch(Course = 'A', 'Advance',Course = 'B', 'Basic',Course = 'N','None') as CourseName,Experience,Training," & _
            " Qualification,JoinDate,Salary,OT_Rate,LeftOn,U_Name,U_EntDt,U_AE " & _
            " from Emp_Mast where Emp_Code='" & Txt(EmpName).Tag & "'"
    ElseIf PubBackEnd = "S" Then
        mQry = "Select Emp_Code,Site_Code,Emp_Name,case Emp_Type WHEN 0 THEN 'Sales Staff' WHEN 1 THEN 'Workshop' WHEN 2 THEN 'Others' END as EmpType," & _
            " FName,ADD1,ADD2,CityName,PHONE,PAGER,Mobile,Reference,DOB,AGE,Access_PWD,Designation," & _
            " CASE ServStatus WHEN '0' THEN 'Skilled' WHEN '1' THEN 'Semi-Skilled' WHEN '2' THEN 'Un-Skilled' END as SrvStatus," & _
            " CASE Course WHEN 'A' THEN 'Advance' WHEN 'B' THEN 'Basic' WHEN 'N' THEN 'None' END as CourseName,Experience,Training," & _
            " Qualification,JoinDate,Salary,OT_Rate,LeftOn,U_Name,U_EntDt,U_AE " & _
            " from Emp_Mast where Emp_Code='" & Txt(EmpName).Tag & "'"
    End If
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION, , True)
    Set Rst = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION

End Sub
Private Sub TopCtrl1_eRef()
    RsCity.Requery
    RsHelp.Requery
    RsDiv.Requery
    RsSupervisor.Requery
End Sub
Private Sub TopCtrl1_eSave()
Dim I As Integer
Dim mTrans As Boolean
Dim DocIdHlp As String
Dim mSearchCode$
On Error GoTo errlbl
    If IsValid(Txt(EmpName), "Employee Name") = False Then Exit Sub
    If IsValid(Txt(DIV), Label3(16)) = False Then Exit Sub
    If IsValid(Txt(EmpType), Label3(4)) = False Then Exit Sub
    If IsValid(Txt(Desig), Label3(3)) = False Then Exit Sub
    If IsValid(Txt(Status), Label3(2)) = False Then Exit Sub
    If IsValid(Txt(Course), Label3(9)) = False Then Exit Sub
    
    If Txt(ConPW) <> Txt(AccPW) Then MsgBox "Confirm Password", vbCritical, "Validation": Txt(ConPW).SetFocus: Exit Sub
    
    Txt(DOL).TEXT = RetDate(Txt(DOL))
    Txt_Validate EmpName, True
    GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        'txt(EmpName).Tag = IIf(IsNull(GCn.Execute("select max(val(" & cMID("Emp_Code", 2, 3) & ")) from emp_mast where site_code='" & PubSiteCode & "'").Fields(0).Value), 1, GCn.Execute("select max(val(mid(emp_code,2,3))) from emp_mast where site_code='" & PubSiteCode & "'").Fields(0).Value + 1)
        Txt(EmpName).Tag = IIf(IsNull(GCn.Execute("select Max(" & cVal(cMID("Emp_Code", "2", "3")) & ") from emp_mast where Left(Emp_Code,1)='" & PubSiteCode & "'").Fields(0).Value), 1, GCn.Execute("select max(" & cVal(cMID("Emp_Code", "2", "3")) & ") from emp_mast where Left(Emp_Code,1)='" & PubSiteCode & "'").Fields(0).Value + 1)
        
        GCn.Execute ("insert into Emp_Mast(Emp_Code,Site_Code,Div_Code,Emp_Name,Emp_Type,FName,ADD1,ADD2,CityName,PHONE,PAGER,Mobile,Reference,DOB,AGE,Access_PWD,Designation,ServStatus,Course,Experience,Training,Qualification,JoinDate,Salary,LeftOn,Supervisor,U_Name,U_EntDt,U_AE,OT_Rate) " & _
          " values('" & PubSiteCode & Txt(EmpName).Tag & "','" & PubSiteCode & "','" & Txt(DIV).Tag & "','" & Txt(EmpName) & "'," & IIf(Txt(EmpType) = "Sales Staff", 0, IIf(Txt(EmpType) = "Workshop", 1, 2)) & ",'" & Txt(FathName) & "','" & Txt(Add1) & "','" & Txt(Add2) & "','" & Txt(City) & "','" & Txt(Phone) & "','" & Txt(Pager) & "','" & Txt(Mobile) & _
          "', '" & Txt(Referen) & "'," & ConvertDate(Txt(DOB)) & "," & Val(Txt(Age)) & ",'" & Txt(AccPW) & "','" & Txt(Desig) & "','" & IIf(Txt(Status).TEXT = "Skilled", "0", IIf(Txt(Status).TEXT = "Semi-Skilled", "1", "2")) & "','" & left(Txt(Course), 1) & "','" & Txt(Experience) & "','" & Txt(Traning) & "','" & Txt(Qualific) & _
          "', " & ConvertDate(Txt(DOJ)) & "," & Val(Txt(Salary)) & "," & ConvertDate(Txt(DOL)) & ", '" & Txt(Supervisor).Tag & "','" & pubUName & "'," & ConvertDate(Format(PubServerDate, "dd/MMM/yyyy HH:NN:SS")) & ",'" & left(TopCtrl1.TopText2, 1) & "'," & Val(Txt(OTRate)) & ")")
        mSearchCode = PubSiteCode & Txt(EmpName).Tag
    Else
        GCn.Execute ("update Emp_Mast set Site_Code='" & PubSiteCode & "',Emp_Name='" & Txt(EmpName) & "',Emp_Type=" & IIf(Txt(EmpType) = "Sales Staff", 0, IIf(Txt(EmpType) = "Workshop", 1, 2)) & ",FName='" & Txt(FathName) & "',ADD1='" & Txt(Add1) & "',ADD2='" & Txt(Add2) & _
            "',CityName='" & Txt(City) & "',PHONE='" & Txt(Phone) & "',PAGER='" & Txt(Pager) & "',Mobile='" & Txt(Mobile) & "', Reference='" & Txt(Referen) & "',DOB=" & ConvertDate(Txt(DOB)) & ",AGE=" & Val(Txt(Age)) & ",Access_PWD='" & Txt(AccPW) & _
            "',Designation='" & Txt(Desig) & "',ServStatus='" & IIf(Txt(Status).TEXT = "Skilled", "0", IIf(Txt(Status).TEXT = "Semi-Skilled", "1", "2")) & "',Course='" & left(Txt(Course), 1) & "',Experience='" & Txt(Experience) & "',Training='" & Txt(Traning) & _
            "',Qualification='" & Txt(Qualific) & "',JoinDate=" & ConvertDate(Txt(DOJ)) & ",Salary=" & Val(Txt(Salary)) & ",LeftOn=" & ConvertDate(Txt(DOL)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(Format(PubServerDate, "dd/MMM/yyyy HH:NN:SS")) & _
            ",Supervisor = '" & Txt(Supervisor).Tag & "', U_AE='" & left(TopCtrl1.TopText2, 1) & _
            "',Div_Code='" & Txt(DIV).Tag & "',OT_Rate=" & Val(Txt(OTRate)) & " where emp_code='" & Txt(EmpName).Tag & "'")
        mSearchCode = Txt(EmpName).Tag
    End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select Emp_Mast.*, Emp_Mast.Emp_Code as SearchCode from Emp_Mast Where Emp_Mast.Emp_Code = '" & mSearchCode & "' Order by Emp_Name")
    End If
    RsHelp.Requery
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        Master.FIND "emp_code = '" & PubSiteCode & Txt(EmpName).Tag & "'"
        TopCtrl1_eAdd
        Exit Sub
    Else
        Master.FIND "emp_code = '" & Txt(EmpName).Tag & "'"
    End If
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
Exit Sub
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim rsdes As ADODB.Recordset, I As Integer, XXA() As String
Select Case Index
    Case EmpType
        ListArray = Array("Sales Staff", "Workshop", "Others")
        Set mListItem = ListView_Items(ListView, Txt, EmpType, ListArray, 3)
    Case EmpName
        PreName = Txt(EmpName).TEXT
    Case Desig
        Set rsdes = New ADODB.Recordset
        With rsdes
             .CursorLocation = adUseClient
             .Open "SELECT designation from designation where emp_type=" & IIf(Txt(EmpType) = "Sales Staff", 0, IIf(Txt(EmpType) = "Workshop", 1, 2)), GCn, adOpenDynamic, adLockOptimistic
        End With
        Do While Not rsdes.EOF
            I = I
            ReDim Preserve XXA(I)
            XXA(I) = rsdes!Designation
            I = I + 1
            rsdes.MoveNext
        Loop
        urec = rsdes.RecordCount
        Set mListItem = ListView_Items(ListView, Txt, Desig, XXA, rsdes.RecordCount)
        Set rsdes = Nothing
    Case Status
        ListArray = Array("Skilled", "Semi-Skilled", "Un-Skilled")
        Set mListItem = ListView_Items(ListView, Txt, Status, ListArray, 3)
    Case Course
        ListArray = Array("Advance", "Basic", "None")
        Set mListItem = ListView_Items(ListView, Txt, Course, ListArray, 3)
    Case EmpName
        Set rsdes = New ADODB.Recordset
        With rsdes
             .CursorLocation = adUseClient
             .Open "SELECT designation  from designation where site_code='" & PubSiteCode & "'" & " and emp_type=" & IIf(Txt(EmpType) = "Sales Staff", 0, IIf(Txt(EmpType) = "Workshop", 1, 2)), GCn, adOpenDynamic, adLockOptimistic
        End With
        ListView_Items_RecordSet ListView, Txt, Index, rsdes
    Case Supervisor
        DgSupervisor.Move Txt(Index).left, Txt(Index).top + Txt(Index).height + 20
        If RsSupervisor.RecordCount > 0 Then
            RsSupervisor.MoveFirst
            RsSupervisor.FIND "Code = '" & Txt(Supervisor).Tag & "'"
        End If
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
    'f TopCtrl1.PrvKeyCode <> vbKeyEscape Then Txt(Index).Text = ""
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case EmpName
        DGridTxtKeyDown_Mast DGEmp, Txt, EmpName, RsHelp, KeyCode, True
    Case EmpType
        ListView_KeyDown FrmList, ListView, Txt, EmpType, KeyCode, Shift, Txt(EmpType).left, (Txt(EmpType).top + Txt(EmpType).height), Txt(EmpType).width, 260 * 3
    Case Desig
        If ListView.ListItems.Count > 0 Then ListView_KeyDown FrmList, ListView, Txt, Desig, KeyCode, Shift, Txt(Desig).left, (Txt(Desig).top + Txt(Desig).height), Txt(Desig).width, 260 * urec
'        ListView_KeyDown FrmList, ListView, txt, Desig, KeyCode, Shift, txt(Desig).left, (txt(Desig).top + txt(Desig).Height), txt(Desig).width, 260 * urec
    Case Status
        ListView_KeyDown FrmList, ListView, Txt, Status, KeyCode, Shift, Txt(Status).left, (Txt(Status).top + Txt(Status).height), Txt(Status).width, 260 * 3
    Case Course
        ListView_KeyDown FrmList, ListView, Txt, Course, KeyCode, Shift, Txt(Course).left, (Txt(Course).top + Txt(Course).height), Txt(Course).width, 260 * 3
    Case City
        DGridTxtKeyDown DGCity, Txt, City, RsCity, KeyCode, False, 1, frmCity, "frmCity"
    Case Supervisor
        DGridTxtKeyDown DgSupervisor, Txt, Index, RsSupervisor, KeyCode, False, 1
    Case DIV
        DGridTxtKeyDown DGDIV, Txt, DIV, RsDiv, KeyCode, False, 1
End Select

If DGDIV.Visible = False And DGCity.Visible = False And FrmList.Visible = False And DgSupervisor.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> DOL Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = DOL Then
            Txt_Validate Index, True
'            If CDate(Txt(DOB)) < CDate(Txt(DOJ)) Then Exit Sub
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> EmpName Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> EmpName Then
            If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub
Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case City
        If DGCity.Visible = True Then DGridTxtKeyPress Txt, City, RsCity, KeyAscii, "Name"
    Case Supervisor
        If DgSupervisor.Visible = True Then DGridTxtKeyPress Txt, Index, RsSupervisor, KeyAscii, "Name"
    Case DIV
        If DGDIV.Visible = True Then DGridTxtKeyPress Txt, DIV, RsDiv, KeyAscii, "Div_Name"
    Case Age
        Call NumPress(Txt(Index), KeyAscii, 3, 0)
    Case Salary
        Call NumPress(Txt(Index), KeyAscii, 6, 2)
    Case OTRate
        Call NumPress(Txt(Index), KeyAscii, 3, 2)
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case EmpName
        If DGEmp.Visible = True Then DGridTxtKeyUp_Mast Txt, EmpName, RsHelp, KeyCode, "Name"
    Case EmpType
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, EmpType, KeyCode, mListItem
    Case Desig
        If FrmList.Visible = True And ListView.ListItems.Count > 0 Then ListView_KeyUp ListView, Txt, Desig, KeyCode, mListItem
        If UCase(Txt(Desig)) = "MECHANIC" Then
            Txt(Supervisor).Enabled = True
        Else
            Txt(Supervisor).Enabled = False
        End If
    Case Status
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Status, KeyCode, mListItem
    Case Course
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Course, KeyCode, mListItem
    Case DOB
        If KeyCode = vbKeyDown Then SendKeysA vbKeyTab, True
        If Txt(Index) = "" Then Txt(Age).Enabled = True

End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim rsdes As ADODB.Recordset
Dim I As Integer
Dim XXA() As String
Dim mo As Integer
Select Case Index
    Case EmpName
'        If TopCtrl1.TopText2 = "Edit" And RsHelp!ncode = Txt(EmpName).Tag Then
'        'If RsHelp.RecordCount  > 0 Then
'        If UCase(DGEmp.Text) = UCase(Txt(Index).Text) Then Exit Sub
'        '    If UCase(DGEmp.Text) = UCase(Txt(Index).Text) Then
'           '     MsgBox "Duplicate Employee Name", vbCritical, "Validation": Cancel = True
'        '    End If
'        'End If
'        'End If
'        ElseIf RsHelp.RecordCount  > 0 Then
'            If UCase(DGEmp.Text) = UCase(Txt(Index).Text) Then
'                MsgBox "Duplicate Employee Name", vbCritical, "Validation": Cancel = True
'            End If
'        End If
        If RsHelp.RecordCount = 0 Then Exit Sub
        If TopCtrl1.TopText2 = "Add" Then
            If GCn.Execute("select count(*) from Emp_Mast where Emp_Name='" & Txt(Index).TEXT & "'").Fields(0) > 0 Then
                MsgBox "Duplicate Employee Name", vbCritical, "Validation Error"
                Txt(Index).TEXT = ""
                Cancel = True
                Exit Sub
            End If
        End If
        If TopCtrl1.TopText2 = "Edit" And Txt(EmpName).TEXT <> EName Then
            If GCn.Execute("select count(*) from Emp_Mast where Emp_Name='" & Txt(Index).TEXT & "'").Fields(0) > 0 Then
                MsgBox "Duplicate Employee Name", vbCritical, "Validation Error"
                Txt(Index).TEXT = ""
                Cancel = True
                Exit Sub
            End If
         End If
        
    Case DOB, DOJ, DOL
        If Len(Trim(Txt(Index).TEXT)) > 0 Then
            Txt(Index).TEXT = RetDate(Txt(Index))
        End If
        If Index = DOB And Txt(Index).TEXT <> "" And Txt(DOJ).TEXT <> "" Then
            If CDate(Txt(DOJ)) < CDate(Txt(DOB)) Then
                MsgBox "Date of Birth Can't be less than Date of Joining", vbInformation, "Validation": Cancel = True
            End If
        ElseIf Index = DOJ And Txt(Index).TEXT <> "" And Txt(DOB).TEXT <> "" Then
            
            If CDate(Txt(DOJ)) < CDate(Txt(DOB)) Then
                MsgBox "Date of Joining Can't be less than Date of Birth", vbInformation, "Validation": Cancel = True
            End If
        ElseIf Index = DOL And Txt(Index).TEXT <> "" And Txt(DOJ).TEXT <> "" Then
            If CDate(Txt(DOL)) < CDate(Txt(DOJ)) Then
                MsgBox "Date of Leaving Can't be less than Date of Joining", vbInformation, "Validation": Cancel = True
            End If
        End If
            If Index = DOB And Txt(Index) <> "" Then
                Txt(Age).TEXT = DateDiff("yyyy", Txt(DOB), date)
                If Txt(Age).Enabled = True Then Txt(Age).Enabled = False
            End If
    Case Supervisor
        If RsSupervisor.RecordCount > 0 And RsSupervisor.EOF = False And RsSupervisor.BOF = False And Txt(Supervisor) <> "" Then
            Txt(Supervisor) = RsSupervisor!Name
            Txt(Supervisor).Tag = RsSupervisor!Code
        Else
            Txt(Supervisor) = ""
            Txt(Supervisor).Tag = ""
        End If
    Case ConPW
        If Txt(ConPW) <> "" Then
        If Txt(ConPW) <> Txt(AccPW) Then MsgBox "Invalid Password", vbCritical, "Validation": Cancel = True
        End If
    Case AccPW
        Txt(ConPW).SetFocus
    Case EmpType
        If Txt(EmpType).TEXT <> "" Then Txt(EmpType).TEXT = ListView.SelectedItem.TEXT
'        Set rsdes = New ADODB.Recordset
'        With rsdes
'             .CursorLocation = adUseClient
'             .Open "SELECT designation  from designation where site_code='" & PubSiteCode & "'" & " and emp_type=" & IIf(txt(EmpType) = "Sales Staff", 0, IIf(txt(EmpType) = "Mechanic", 1, 2)), GCn, adOpenDynamic, adLockOptimistic
'        End With
'        ListView_Items_RecordSet
'        Do While Not rsdes.EOF
'            i = i
'            ReDim Preserve XXA(i)
'            XXA(i) = rsdes!Designation
'            i = i + 1
'            rsdes.MoveNext
'        Loop
'        If rsdes.RecordCount  > 0 Then rsdes.MoveFirst
'        If txt(EmpType).Text <> txt(EmpType).Tag Then
'            If rsdes.RecordCount  > 0 Then
'                txt(Desig).Text = rsdes!Designation
'            Else
'                txt(Desig).Text = ""
'            End If
'        End If
'        urec = rsdes.RecordCount
'        Set mListItem = ListView_Items(ListView, txt, Desig, XXA, rsdes.RecordCount)
'        Set rsdes = Nothing
End Select
End Sub


'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To 22
    Txt(I).TEXT = ""
Next I
End Sub

Private Sub MoveRec()
Dim mo As String
TopCtrl1.tDel = False
If Master.RecordCount > 0 Then
    Txt(EmpName) = Master!Emp_Name   '0
    Txt(EmpName).Tag = Master!Emp_Code '0
    Txt(FathName) = IIf(IsNull(Master!fname), "", Master!fname)
    Txt(Add1) = IIf(IsNull(Master!Add1), "", Master!Add1)
    Txt(Add2) = IIf(IsNull(Master!Add2), "", Master!Add2)
    Txt(City) = IIf(IsNull(Master!CityName), "", Master!CityName)
    Txt(Phone) = IIf(IsNull(Master!Phone), "", Master!Phone)
    Txt(Pager) = IIf(IsNull(Master!Pager), "", Master!Pager)
    Txt(Mobile) = IIf(IsNull(Master!Mobile), "", Master!Mobile)
    Txt(Referen) = IIf(IsNull(Master!Reference), "", Master!Reference)
    Txt(DOB) = IIf(IsNull(Master!DOB), "", Master!DOB)
    Txt(Age) = IIf(IsNull(Master!Age), "", Master!Age)
    Txt(AccPW) = IIf(IsNull(Master!Access_PWD), "", Master!Access_PWD)
    Txt(ConPW) = IIf(IsNull(Master!Access_PWD), "", Master!Access_PWD) '10
    Txt(EmpType) = IIf(Master!emp_type = 0, "Sales Staff", IIf(Master!emp_type = 1, "Workshop", "Others"))
    Txt(Desig) = IIf(IsNull(Master!Designation), "", Master!Designation)
    If UCase(Txt(Desig)) <> "MECHANIC" Then
        Txt(Supervisor).Enabled = False
    Else
        If TopCtrl1.TopText2 <> "Browse" Then
            Txt(Supervisor).Enabled = True
        Else
            Txt(Supervisor).Enabled = False
        End If
    End If
    Txt(Supervisor).Tag = XNull(Master!Supervisor)
    If Txt(Supervisor).Tag <> "" Then
        Txt(Supervisor) = GCn.Execute("Select Emp_Name From Emp_Mast Where Emp_Code='" & Txt(Supervisor).Tag & "'").Fields(0)
    Else
        Txt(Supervisor) = ""
    End If
    Txt(Status) = IIf(Master!ServStatus = "0", "Skilled", IIf(Master!ServStatus = "1", "Semi-Skilled", "Un-Skilled"))
    Txt(Course) = IIf(Master!Course = "A", "Advance", IIf(Master!Course = "B", "Basic", "None"))
    Txt(Experience) = IIf(IsNull(Master!Experience), "", Master!Experience)
    Txt(Traning) = IIf(IsNull(Master!Training), "", Master!Training)
    Txt(Qualific) = IIf(IsNull(Master!Qualification), "", Master!Qualification)
    Txt(DOJ) = IIf(IsNull(Master!JoinDate), "", Master!JoinDate)
    Txt(Salary) = IIf(IsNull(Master!Salary), 0, Master!Salary)
    Txt(OTRate) = IIf(IsNull(Master!OT_Rate), 0, Master!OT_Rate)
    Txt(DOL) = IIf(IsNull(Master!LeftOn), "", Master!LeftOn) '20
    If Not IsNull(Master!Div_Code) Then
        RsDiv.MoveFirst
        RsDiv.FIND ("Div_Code='" & Master!Div_Code & "'")
        Txt(DIV).TEXT = RsDiv!Div_Name
        Txt(DIV).Tag = Master!Div_Code
    Else
        Txt(DIV).TEXT = ""
        Txt(DIV).Tag = ""
    End If
End If
Grid_Hide
'TopCtrl1.tPrn = False
Exit Sub
error1:
        CheckError
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
Next


    If UCase(Txt(Desig)) <> "MECHANIC" Then
        Txt(Supervisor).Enabled = False
    Else
        If TopCtrl1.TopText2 <> "Browse" Then
            Txt(Supervisor).Enabled = True
        Else
            Txt(Supervisor).Enabled = False
        End If
    End If

If Txt(DOB) <> "" Then Txt(Age).Enabled = False
'Txt(JobType).Enabled = False
'Txt(JobTypeN).Enabled = False
End Sub
Private Sub Grid_Hide()
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGCity.Visible = True Then DGCity.Visible = False
    If DGEmp.Visible = True Then DGEmp.Visible = False
    If DGDIV.Visible = True Then DGDIV.Visible = False
    If DgSupervisor.Visible = True Then DgSupervisor.Visible = False
End Sub

