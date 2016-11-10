VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "topctl.ocx"
Begin VB.Form frmVehTripClose 
   Appearance      =   0  'Flat
   BackColor       =   &H00CFE0E0&
   Caption         =   "Trip Close Entry"
   ClientHeight    =   6615
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
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
      Height          =   255
      Index           =   33
      Left            =   9270
      TabIndex        =   19
      Top             =   2310
      Width           =   1470
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
      Height          =   255
      Index           =   32
      Left            =   9270
      MaxLength       =   8
      TabIndex        =   20
      Top             =   2580
      Width           =   1470
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
      Height          =   255
      Index           =   31
      Left            =   9270
      TabIndex        =   21
      Top             =   2850
      Width           =   1470
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
      Height          =   255
      Index           =   30
      Left            =   10035
      TabIndex        =   32
      Top             =   5550
      Width           =   705
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
      Height          =   255
      Index           =   3
      Left            =   7515
      TabIndex        =   31
      Top             =   5550
      Width           =   900
   End
   Begin MSDataGridLib.DataGrid DGCode 
      Height          =   4470
      Left            =   5235
      Negotiate       =   -1  'True
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   6435
      Visible         =   0   'False
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   7885
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
      Caption         =   "Machine No  Help"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Machine No"
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
         DataField       =   "Site_Code"
         Caption         =   "SiteCode"
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
            ColumnWidth     =   3254.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
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
      Height          =   255
      Index           =   4
      Left            =   3495
      TabIndex        =   92
      Top             =   1905
      Width           =   1470
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   3570
      Left            =   -135
      Negotiate       =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   6297
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
      Caption         =   "From To Place Help"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "FromCode"
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
         DataField       =   "FromPlace"
         Caption         =   "From"
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
         DataField       =   "ToPlace"
         Caption         =   "To"
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
         DataField       =   "ToCode"
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
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2580.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2984.882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   14.74
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
      Height          =   255
      Index           =   29
      Left            =   9255
      TabIndex        =   30
      Top             =   5280
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
      Height          =   255
      Index           =   28
      Left            =   9255
      TabIndex        =   29
      Top             =   5010
      Width           =   1470
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
      Height          =   255
      Index           =   27
      Left            =   9255
      TabIndex        =   28
      Top             =   4740
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
      Height          =   255
      Index           =   26
      Left            =   9255
      MaxLength       =   5
      TabIndex        =   27
      Top             =   4470
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
      Height          =   255
      Index           =   25
      Left            =   9255
      TabIndex        =   26
      Top             =   4200
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
      Height          =   255
      Index           =   24
      Left            =   9255
      MaxLength       =   5
      TabIndex        =   25
      Top             =   3930
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
      Height          =   255
      Index           =   23
      Left            =   9255
      TabIndex        =   24
      Top             =   3660
      Width           =   1470
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
      Height          =   255
      Index           =   22
      Left            =   9255
      TabIndex        =   23
      Top             =   3390
      Width           =   1470
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
      Height          =   255
      Index           =   21
      Left            =   9255
      TabIndex        =   22
      Top             =   3120
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
      Height          =   705
      Index           =   20
      Left            =   9270
      MaxLength       =   35
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   1590
      Width           =   2190
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
      Height          =   255
      Index           =   19
      Left            =   9270
      TabIndex        =   17
      Top             =   1320
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
      Height          =   255
      Index           =   18
      Left            =   9270
      TabIndex        =   16
      Top             =   1050
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
      Height          =   255
      Index           =   17
      Left            =   3495
      TabIndex        =   15
      Top             =   6060
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
      Height          =   255
      Index           =   16
      Left            =   3495
      MaxLength       =   5
      TabIndex        =   14
      Top             =   5790
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
      Height          =   255
      Index           =   15
      Left            =   3495
      TabIndex        =   13
      Top             =   5520
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
      Height          =   900
      Index           =   14
      Left            =   3495
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4605
      Width           =   2190
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
      Height          =   255
      Index           =   13
      Left            =   3495
      TabIndex        =   11
      Top             =   4335
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
      Height          =   255
      Index           =   12
      Left            =   3495
      TabIndex        =   10
      Top             =   4065
      Width           =   1470
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
      Height          =   255
      Index           =   11
      Left            =   3495
      TabIndex        =   9
      Top             =   3795
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
      Height          =   255
      Index           =   10
      Left            =   3495
      MaxLength       =   35
      TabIndex        =   8
      Top             =   3525
      Width           =   2190
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
      Height          =   255
      Index           =   9
      Left            =   3495
      TabIndex        =   7
      Top             =   3255
      Width           =   2190
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
      Height          =   255
      Index           =   6
      Left            =   3495
      MaxLength       =   5
      TabIndex        =   4
      Top             =   2445
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
      Height          =   255
      Index           =   7
      Left            =   3495
      TabIndex        =   5
      Top             =   2715
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
      Height          =   255
      Index           =   8
      Left            =   3495
      MaxLength       =   5
      TabIndex        =   6
      Top             =   2985
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
      Height          =   255
      Index           =   5
      Left            =   3495
      LinkTimeout     =   255
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2175
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
      Height          =   255
      Index           =   0
      Left            =   3495
      TabIndex        =   0
      Top             =   1095
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
      Height          =   255
      Index           =   1
      Left            =   3495
      TabIndex        =   1
      Top             =   1365
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
      Height          =   255
      Index           =   2
      Left            =   3495
      TabIndex        =   2
      Top             =   1635
      Width           =   1470
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   661
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
      Height          =   285
      Index           =   31
      Left            =   9075
      TabIndex        =   105
      Top             =   2340
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Freight Charge"
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
      Index           =   33
      Left            =   5895
      TabIndex        =   104
      Top             =   2310
      Width           =   1230
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
      Height          =   285
      Index           =   30
      Left            =   9075
      TabIndex        =   103
      Top             =   2610
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional KMS."
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
      Left            =   5895
      TabIndex        =   102
      Top             =   2595
      Width           =   1275
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
      Height          =   285
      Index           =   29
      Left            =   9075
      TabIndex        =   101
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Freight"
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
      Index           =   31
      Left            =   5895
      TabIndex        =   100
      Top             =   2865
      Width           =   1425
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
      Height          =   285
      Index           =   28
      Left            =   9960
      TabIndex        =   99
      Top             =   5550
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Desiel"
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
      Index           =   30
      Left            =   8520
      TabIndex        =   98
      Top             =   5565
      Width           =   1515
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
      Height          =   285
      Index           =   24
      Left            =   7380
      TabIndex        =   97
      Top             =   5550
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Desiel"
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
      Index           =   29
      Left            =   5880
      TabIndex        =   96
      Top             =   5550
      Width           =   1335
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
      Height          =   285
      Index           =   9
      Left            =   3330
      TabIndex        =   94
      Top             =   1905
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trip Start Date"
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
      Left            =   615
      TabIndex        =   93
      Top             =   1905
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Expenses"
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
      Left            =   5895
      TabIndex        =   91
      Top             =   5280
      Width           =   1680
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
      Height          =   285
      Index           =   27
      Left            =   9075
      TabIndex        =   90
      Top             =   5280
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Break Down Remarks"
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
      Index           =   26
      Left            =   5895
      TabIndex        =   89
      Top             =   5010
      Width           =   1815
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
      Height          =   285
      Index           =   26
      Left            =   9075
      TabIndex        =   88
      Top             =   5010
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Break Down Cost"
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
      Index           =   25
      Left            =   5895
      TabIndex        =   87
      Top             =   4740
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   25
      Left            =   9075
      TabIndex        =   86
      Top             =   4740
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Break Down UpTo Time"
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
      Index           =   24
      Left            =   5895
      TabIndex        =   85
      Top             =   4470
      Width           =   1965
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
      Height          =   285
      Index           =   23
      Left            =   9075
      TabIndex        =   84
      Top             =   4470
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Break Down UpTo Date"
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
      Index           =   23
      Left            =   5895
      TabIndex        =   83
      Top             =   4200
      Width           =   1935
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
      Height          =   285
      Index           =   22
      Left            =   9075
      TabIndex        =   82
      Top             =   4200
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Break Down From Time"
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
      Index           =   22
      Left            =   5895
      TabIndex        =   81
      Top             =   3930
      Width           =   1950
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
      Height          =   285
      Index           =   21
      Left            =   9075
      TabIndex        =   80
      Top             =   3930
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Break Down From Date"
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
      Index           =   21
      Left            =   5895
      TabIndex        =   79
      Top             =   3660
      Width           =   1920
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
      Height          =   285
      Index           =   20
      Left            =   9075
      TabIndex        =   78
      Top             =   3660
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Meter Reading"
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
      Left            =   5895
      TabIndex        =   77
      Top             =   3390
      Width           =   1200
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
      Height          =   285
      Index           =   19
      Left            =   9075
      TabIndex        =   76
      Top             =   3390
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parchi Expenses"
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
      Left            =   5895
      TabIndex        =   75
      Top             =   3105
      Width           =   1395
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
      Height          =   285
      Index           =   18
      Left            =   9075
      TabIndex        =   74
      Top             =   3120
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnLoading Weighing Slip Remarks"
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
      Left            =   5910
      TabIndex        =   73
      Top             =   1590
      Width           =   2910
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
      Height          =   285
      Index           =   17
      Left            =   9090
      TabIndex        =   72
      Top             =   1590
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnLoading Weighing Slip Date"
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
      Left            =   5910
      TabIndex        =   71
      Top             =   1320
      Width           =   2535
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
      Height          =   285
      Index           =   16
      Left            =   9090
      TabIndex        =   70
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnLoading Weighing Slip No"
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
      Left            =   5910
      TabIndex        =   69
      Top             =   1050
      Width           =   2385
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
      Height          =   285
      Index           =   15
      Left            =   9090
      TabIndex        =   68
      Top             =   1050
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unloading Place"
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
      Left            =   690
      TabIndex        =   67
      Top             =   6045
      Width           =   1365
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
      Height          =   285
      Index           =   13
      Left            =   3315
      TabIndex        =   66
      Top             =   6060
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unloading Time"
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
      Left            =   690
      TabIndex        =   65
      Top             =   5775
      Width           =   1320
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
      Height          =   285
      Index           =   12
      Left            =   3315
      TabIndex        =   64
      Top             =   5790
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unloading Date"
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
      Index           =   13
      Left            =   690
      TabIndex        =   63
      Top             =   5505
      Width           =   1290
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
      Height          =   285
      Index           =   11
      Left            =   3315
      TabIndex        =   62
      Top             =   5520
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Weighing Slip Remarks"
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
      Left            =   615
      TabIndex        =   61
      Top             =   4605
      Width           =   2670
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
      Height          =   285
      Index           =   10
      Left            =   3330
      TabIndex        =   60
      Top             =   4620
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Weighing Slip Date"
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
      Left            =   615
      TabIndex        =   59
      Top             =   4335
      Width           =   2295
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
      Height          =   285
      Index           =   8
      Left            =   3330
      TabIndex        =   58
      Top             =   4335
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Weighing Slip No"
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
      Left            =   615
      TabIndex        =   57
      Top             =   4065
      Width           =   2145
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
      Height          =   285
      Index           =   7
      Left            =   3330
      TabIndex        =   56
      Top             =   4065
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Qty."
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
      Left            =   615
      TabIndex        =   55
      Top             =   3795
      Width           =   1020
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
      Height          =   285
      Index           =   5
      Left            =   3330
      TabIndex        =   54
      Top             =   3795
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Item"
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
      Left            =   615
      TabIndex        =   53
      Top             =   3525
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
      Height          =   285
      Index           =   3
      Left            =   3330
      TabIndex        =   52
      Top             =   3525
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Place"
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
      Left            =   615
      TabIndex        =   51
      Top             =   3255
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   0
      Left            =   3330
      TabIndex        =   50
      Top             =   3255
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Trip Close Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   885
      TabIndex        =   49
      Top             =   450
      Width           =   2130
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   5775
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   11415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reaching Time"
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
      Left            =   615
      TabIndex        =   47
      Top             =   2445
      Width           =   1260
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
      Height          =   285
      Index           =   14
      Left            =   3330
      TabIndex        =   46
      Top             =   2445
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Date"
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
      Left            =   615
      TabIndex        =   45
      Top             =   2715
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   6
      Left            =   3330
      TabIndex        =   44
      Top             =   2715
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
      Height          =   285
      Index           =   4
      Left            =   3330
      TabIndex        =   43
      Top             =   2985
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Time"
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
      Left            =   615
      TabIndex        =   42
      Top             =   2985
      Width           =   1140
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
      Height          =   285
      Index           =   2
      Left            =   3330
      TabIndex        =   41
      Top             =   2175
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reaching Date"
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
      Left            =   615
      TabIndex        =   40
      Top             =   2175
      Width           =   1230
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
      Height          =   285
      Index           =   1
      Left            =   3330
      TabIndex        =   38
      Top             =   1095
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slip No."
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
      Left            =   615
      TabIndex        =   37
      Top             =   1095
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Truck No"
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
      Left            =   615
      TabIndex        =   36
      Top             =   1635
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date "
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
      Left            =   615
      TabIndex        =   35
      Top             =   1365
      Width           =   435
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
      Height          =   285
      Index           =   88
      Left            =   3330
      TabIndex        =   34
      Top             =   1635
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
      Height          =   285
      Index           =   90
      Left            =   3330
      TabIndex        =   33
      Top             =   1365
      Width           =   45
   End
End
Attribute VB_Name = "frmVehTripClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rscity As ADODB.Recordset
Dim RsCode As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim adddis As Double
Dim PurExp As Double
Private Const SlipNo As Byte = 0
Private Const SlipDate As Byte = 1
Private Const TruckNo As Byte = 2
Private Const SDesiel As Byte = 3
Private Const StDate As Byte = 4
Private Const ReachDate  As Byte = 5
Private Const ReachTime As Byte = 6
Private Const LoadDate As Byte = 7
Private Const LoadTime As Byte = 8

Private Const LoadPlace As Byte = 9
Private Const LoadItem As Byte = 10
Private Const LoadQty As Byte = 11
Private Const LoadWeighingSNo As Byte = 12
Private Const LoadWeighingSDate As Byte = 13
Private Const LoadWeighingSRemarks As Byte = 14
Private Const UnloadingDate As Byte = 15
Private Const UnloadingTime As Byte = 16
Private Const UnloadingPlace As Byte = 17
Private Const UnLoadWeighingSNo As Byte = 18
Private Const UnLoadWeighingSDate As Byte = 19
Private Const UnLoadWeighingSRemarks As Byte = 20
Private Const ParchiExp As Byte = 21
Private Const MeterReading As Byte = 22
Private Const BDownFromDate As Byte = 23
Private Const BDownFromTime As Byte = 24
Private Const BDownToDate As Byte = 25
Private Const BDownToTime As Byte = 26
Private Const BDownCost As Byte = 27
Private Const BDownRemarks As Byte = 28
Private Const AdditionalExp As Byte = 29
Private Const ADesiel As Byte = 30
Private Const AddFrt As Byte = 31
Private Const AddKMS As Byte = 32
Private Const Frt As Byte = 33

Private Const mVtype As String = "V_TSC"
Private Const mVPrefix As String = "V_TSC"
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_Load()
'On Error GoTo ELoop
Dim i As Byte
    TopCtrl1.Tag = "AEDP"
    WinSetting Me: Ini_Grid
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select V_No as searchcode,* from Veh_TripClose Where Div_Code='" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
    Set Rscity = New ADODB.Recordset
    Rscity.CursorLocation = adUseClient
    Rscity.Open "select DesiFrom as FromCode,City.CityName as FromPlace,DesiUpTo as ToCode,City1.CityName as ToPlace from (FrChartMast Left Join City on FrChartMast.DesiFrom=City.CityCode) Left Join City as City1 on FrChartMast.DesiUpTo=City1.CityCode", GCn, adOpenDynamic, adLockOptimistic
    Set DGCity.DataSource = Rscity
    
    Set RsCode = New ADODB.Recordset
    RsCode.CursorLocation = adUseClient
    RsCode.Open "select Site_Code,TruckNo as name,StDate from  Veh_Trip order by TruckNo and Div_Code='" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
    Set DGCode.DataSource = RsCode
    
    Call MoveRec
    Disp_Text SETS("INI", Me, Master)
    Exit Sub
ELoop: MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Rscity = Nothing
    Set Master = Nothing
    Set RsCode = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
'On Error GoTo ErrorLoop
Dim i As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    txt(SlipNo).TEXT = G_FaCn.Execute("select Start_Srl_No from Voucher_Prefix where V_Type='" & mVtype & "'").Fields(0).Value + 1
    txt(SlipDate).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub
Private Sub TopCtrl1_eDel()
Dim xDocID$
xDocID = PubDivCode & PubSiteCode & PubSiteCode & mVtype & mVPrefix & Space(8 - Len(CStr(txt(SlipNo)))) + Trim(txt(SlipNo))
On Error GoTo eloop1
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                GCn.BeginTrans
                GCn.Execute ("delete from veh_TripClose where V_No=" & txt(SlipNo).TEXT & " and Div_Code='" & PubDivCode & "'")
                GCn.Execute ("Update Veh_Trip Set CloseV_No=0,CloseV_Date=null Where TruckNo='" & txt(TruckNo) & "'")
                GCn.CommitTrans
                Master.Requery
                Rscity.Requery
                RsCode.Requery
                Call MoveRec
                BUTTONS True, Me, Master, 0
            End If
Call LedgerUnPost(G_FaCn, xDocID)
eloop1:
    If err.NUMBER <> 0 Then
        GCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
    End If
End Sub
Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    txt(SlipDate).SetFocus
    txt(TruckNo).Enabled = False
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
Dim i As Integer
'On Error GoTo ErrorLoop
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
Private Sub TopCtrl1_eRef()
    Rscity.Requery
    RsCode.Requery
End Sub
Private Sub TopCtrl1_eSave()
Dim i As Integer
Dim mtrans As Boolean
Dim Rent_YN As Byte
Dim xDocID$
'On Error GoTo errlbl
    Grid_Hide
    If IsValid(txt(SlipDate), "Slip Date") = False Then Exit Sub
    If IsValid(txt(TruckNo), "Truck No") = False Then Exit Sub
    GCn.BeginTrans
    mtrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
    xDocID = PubDivCode & PubSiteCode & PubSiteCode & mVtype & mVPrefix & Space(8 - Len(CStr(SlipNo))) + txt(SlipNo)
    
    
    GSQL = "insert into veh_TripClose(V_No,Div_Code,Site_Code,V_Date,TruckNo,StDate, " & _
        "ReachDate,ReachTime,LoadDate,LoadTime,LoadPlace,LoadItem,LoadQty,LoadWeighingSNo," & _
        "LoadWeighingSDate,LoadWeighingSRemarks,UnloadingDate,UnloadingTime,UnloadingPlace," & _
        "UnLoadWeighingSNo,UnLoadWeighingSDate,UnLoadWeighingSRemarks,ParchiExp,MeterReading," & _
        "BDownFromDate,BDownFromTime,BDownToDate,BDownToTime,BDownCost,BDownRemarks,AdditionalExp,SDesiel,ADesiel,Freight,AddKMS,AddFreight," & _
        "U_Name,U_EntDt,U_AE) values(" & _
        "" & txt(SlipNo).TEXT & ",'" & PubDivCode & "','" & txt(TruckNo).Tag & "'," & ConvertDate(txt(SlipDate)) & ",'" & txt(TruckNo).TEXT & "'," & ConvertDate(txt(StDate)) & "," & _
        "" & ConvertDate(txt(ReachDate)) & ",'" & txt(ReachTime) & "'," & ConvertDate(txt(LoadDate)) & ",'" & txt(LoadTime) & "','" & txt(LoadPlace).Tag & "','" & txt(LoadItem) & "'," & Val(txt(LoadQty)) & "," & Val(txt(LoadWeighingSNo)) & "," & _
        "" & ConvertDate(txt(LoadWeighingSDate)) & ",'" & txt(LoadWeighingSRemarks) & "'," & ConvertDate(txt(UnloadingDate)) & ",'" & txt(UnloadingTime) & "','" & txt(UnloadingPlace).Tag & "'," & _
        "" & Val(txt(UnLoadWeighingSNo)) & "," & ConvertDate(txt(UnLoadWeighingSDate)) & ",'" & txt(UnLoadWeighingSRemarks) & "'," & Val(txt(ParchiExp)) & "," & Val(txt(MeterReading)) & "," & _
        "" & ConvertDate(txt(BDownFromDate)) & ",'" & txt(BDownFromTime) & "'," & ConvertDate(txt(BDownToDate)) & ",'" & txt(BDownToTime) & "'," & Val(txt(BDownCost)) & ",'" & txt(BDownRemarks) & "'," & Val(txt(AdditionalExp)) & "," & Val(txt(SDesiel)) & "," & Val(txt(ADesiel)) & "," & Val(txt(Frt)) & "," & Val(txt(AddKMS)) & "," & Val(txt(AddFrt)) & "," & _
        "'" & pubUName & "',#" & PubLoginDate & "#,'A')"
    GCn.Execute GSQL
    GCn.Execute ("Update Veh_Trip Set CloseV_No=" & txt(SlipNo).TEXT & ",CloseV_Date=" & ConvertDate(txt(SlipDate).TEXT) & " Where TruckNo='" & txt(TruckNo) & "'")
    
    UpdVouSrlNo G_FaCn, xDocID, CDate(txt(SlipDate))
    Else
    GCn.Execute ("delete from veh_TripClose where V_No=" & txt(SlipNo).TEXT & " and Div_Code='" & PubDivCode & "' and Site_Code='" & txt(TruckNo).Tag & "'")
    GCn.Execute ("Update Veh_Trip Set CloseV_No=0,CloseV_Date=null Where TruckNo='" & txt(TruckNo) & "'")
    
    GSQL = "insert into veh_TripClose(V_No,Div_Code,Site_Code,V_Date,TruckNo,StDate, " & _
        "ReachDate,ReachTime,LoadDate,LoadTime,LoadPlace,LoadItem,LoadQty,LoadWeighingSNo," & _
        "LoadWeighingSDate,LoadWeighingSRemarks,UnloadingDate,UnloadingTime,UnloadingPlace," & _
        "UnLoadWeighingSNo,UnLoadWeighingSDate,UnLoadWeighingSRemarks,ParchiExp,MeterReading," & _
        "BDownFromDate,BDownFromTime,BDownToDate,BDownToTime,BDownCost,BDownRemarks,AdditionalExp,SDesiel,ADesiel,Freight,AddKMS,AddFreight," & _
        "U_Name,U_EntDt,U_AE) values(" & _
        "" & txt(SlipNo).TEXT & ",'" & PubDivCode & "','" & txt(TruckNo).Tag & "'," & ConvertDate(txt(SlipDate)) & ",'" & txt(TruckNo).TEXT & "'," & ConvertDate(txt(StDate)) & "," & _
        "" & ConvertDate(txt(ReachDate)) & ",'" & txt(ReachTime) & "'," & ConvertDate(txt(LoadDate)) & ",'" & txt(LoadTime) & "','" & txt(LoadPlace).Tag & "','" & txt(LoadItem) & "'," & Val(txt(LoadQty)) & "," & Val(txt(LoadWeighingSNo)) & "," & _
        "" & ConvertDate(txt(LoadWeighingSDate)) & ",'" & txt(LoadWeighingSRemarks) & "'," & ConvertDate(txt(UnloadingDate)) & ",'" & txt(UnloadingTime) & "','" & txt(UnloadingPlace).Tag & "'," & _
        "" & Val(txt(UnLoadWeighingSNo)) & "," & ConvertDate(txt(UnLoadWeighingSDate)) & ",'" & txt(UnLoadWeighingSRemarks) & "'," & Val(txt(ParchiExp)) & "," & Val(txt(MeterReading)) & "," & _
        "" & ConvertDate(txt(BDownFromDate)) & ",'" & txt(BDownFromTime) & "'," & ConvertDate(txt(BDownToDate)) & ",'" & txt(BDownToTime) & "'," & Val(txt(BDownCost)) & ",'" & txt(BDownRemarks) & "'," & Val(txt(AdditionalExp)) & "," & Val(txt(SDesiel)) & "," & Val(txt(ADesiel)) & "," & Val(txt(Frt)) & "," & Val(txt(AddKMS)) & "," & Val(txt(AddFrt)) & "," & _
        "'" & pubUName & "',#" & PubLoginDate & "#,'A')"
    GCn.Execute GSQL
    GCn.Execute ("Update Veh_Trip Set CloseV_No=" & txt(SlipNo).TEXT & ",CloseV_Date=" & ConvertDate(txt(SlipDate).TEXT) & " Where TruckNo='" & txt(TruckNo) & "'")
    
    End If
'*******A/C Posting**************
'If Val(txt(FuelValue)) > 0 Then ProcAcPost
'********************************
    GCn.CommitTrans
    mtrans = False
    Master.Requery
    Rscity.Requery
    RsCode.Requery
    Master.FIND "SearchCode = " & txt(SlipNo) & ""
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mtrans = True Then
        GCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
Exit Sub
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = Master.Source
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Public Sub SEARCHBACK(ByVal MYVALUE As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("searchcode='" & MYVALUE & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case LoadPlace
        If Rscity.RecordCount = 0 Or (Rscity.EOF = True Or Rscity.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> Rscity!FromPlace Then
            Rscity.MoveFirst
            Rscity.FIND "Fromcode ='" & txt(Index).Tag & "'"
        End If
    Case TruckNo
        If RsCode.RecordCount = 0 Or (RsCode.EOF = True Or RsCode.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsCode!Name Then
            RsCode.MoveFirst
            RsCode.FIND "name ='" & txt(Index).TEXT & "'"
        End If
End Select
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Byte
Dim Txtdate As Boolean
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case LoadPlace
         DGridTxtKeyDown DGCity, txt, Index, Rscity, KeyCode, False, 1
    Case TruckNo
         DGridTxtKeyDown DGCode, txt, Index, RsCode, KeyCode, False, 1
    Case AddKMS
         txt(AddFrt) = Val(txt(AddKMS)) * adddis
         txt(Index).TEXT = Format(txt(Index).TEXT, "0.00")
     Case ParchiExp
         txt(ParchiExp) = PurExp
         txt(Index).TEXT = Format(txt(Index).TEXT, "0.00")
         
End Select
If DGCode.Visible = False And DGCity.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> ADesiel Then Ctrl_DownKeyDown KeyCode, Shift: Exit Sub
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = ADesiel Then If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave: Exit Sub
        If KeyCode = vbKeyUp And Index <> SlipDate Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
 Call CheckQuote(KeyAscii)
Select Case Index
Case LoadPlace
    If DGCity.Visible = True Then DGridTxtKeyPress txt, Index, Rscity, KeyAscii, "FromPlace"
Case TruckNo
    If DGCode.Visible = True Then DGridTxtKeyPress txt, Index, RsCode, KeyAscii, "Name"
Case ParchiExp, AdditionalExp, LoadQty, BDownCost, Frt, AddKMS, AddFrt
    Call NumPress(txt(Index), KeyAscii, 8, 2)
Case MeterReading, LoadWeighingSNo, UnLoadWeighingSNo
    Call NumPress(txt(Index), KeyAscii, 8, 0)
Case ReachTime, LoadTime, UnloadingTime, BDownFromTime, BDownToTime
    Call NumPress(txt(Index), KeyAscii, 2, 2)
End Select
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
End Select
End Sub
Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim RsDsl As ADODB.Recordset
Dim RsDsl1 As ADODB.Recordset
Select Case Index
    Case SlipDate, StDate, ReachDate, LoadDate, LoadWeighingSDate, UnloadingDate, UnLoadWeighingSDate, BDownFromDate, BDownToDate
        txt(Index) = RetDate(txt(Index))
    Case LoadPlace
        If Rscity.RecordCount = 0 Or (Rscity.EOF = True Or Rscity.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = Rscity!FromPlace
            txt(Index).Tag = Rscity!Fromcode
            
            txt(UnloadingPlace).TEXT = Rscity!ToPlace
            txt(UnloadingPlace).Tag = Rscity!tocode
    Set RsDsl = New ADODB.Recordset
    Set RsDsl1 = New ADODB.Recordset
    RsDsl.CursorLocation = adUseClient
    RsDsl1.CursorLocation = adUseClient
    RsDsl.Open "select FrChartMast1.DiselQty from (FrChartMast1 left join Vehicle on Vehicle.Veh_group=FrChartMast1.VehCat) " & _
               "where Vehicle.LorryNo='" & txt(TruckNo) & "' and FrChartMast1.DesiFrom='" & txt(LoadPlace).Tag & "' and FrChartMast1.DesiupTo='" & txt(UnloadingPlace).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
    RsDsl1.Open "select FrChartMast.KMS * FrChartMast.TripFact as Fret,FrChartMast.Parchi,FrChartMast.DistExtra,FrChartMast.AddDistChrg " & _
               "from FrChartMast left join FrChartMast1 on (FrChartMast1.DesiFrom=FrChartMast.DesiFrom and FrChartMast1.DesiUpTo=FrChartMast.DesiUpTo) " & _
               "where FrChartMast.DesiFrom='" & txt(LoadPlace).Tag & "' and FrChartMast.DesiupTo='" & txt(UnloadingPlace).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        
        If RsDsl.RecordCount = 0 Or (RsDsl.EOF = True Or RsDsl.BOF = True) Then
        txt(SDesiel).TEXT = 0
        Else
        txt(SDesiel).TEXT = Format(RsDsl!DiselQty, "0.00")
        End If
        
        If RsDsl1.RecordCount = 0 Or (RsDsl1.EOF = True Or RsDsl1.BOF = True) Then
        txt(Frt).TEXT = 0
        adddis = 0
        Else
        txt(Frt).TEXT = Format(RsDsl1!Fret, "0.00")
        adddis = Format(RsDsl1!AddDistChrg, "0.00")
        PurExp = Format(RsDsl1!Parchi, "0.00")
        txt(ParchiExp) = PurExp
        End If
    
    End If
    Case TruckNo
        If RsCode.RecordCount = 0 Or (RsCode.EOF = True Or RsCode.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
        Else
            txt(Index).TEXT = RsCode!Name
            txt(StDate).TEXT = RsCode!StDate
        End If
    Case ParchiExp, AdditionalExp, Frt, SDesiel, ADesiel
        txt(Index).TEXT = Format(txt(Index).TEXT, "0.00")
    Case ReachTime, LoadTime, UnloadingTime, BDownFromTime, BDownToTime
        txt(Index).TEXT = Format(txt(Index).TEXT, "hh:mm")
    Case AddKMS
              txt(AddFrt) = Val(txt(AddKMS)) * adddis
              txt(Index).TEXT = Format(txt(Index).TEXT, "0.00")
End Select
End Sub
Private Sub DGCode_Click()
    If RsCode.RecordCount > 0 Then
        txt(TruckNo).TEXT = RsCode!Name
        txt(TruckNo).Tag = RsCode!Site_Code
    End If
    DGCode.Visible = False
    txt(TruckNo).SetFocus
End Sub
Private Sub DGCity_Click()
    If Rscity.RecordCount > 0 Then
        txt(LoadPlace).TEXT = Rscity!FromPlace
        txt(LoadPlace).Tag = Rscity!Fromcode
        txt(UnloadingPlace).TEXT = Rscity!ToPlace
        txt(UnloadingPlace).Tag = Rscity!tocode
    End If
    DGCity.Visible = False
    txt(LoadPlace).SetFocus
End Sub
'******* Fuctions **********
Private Sub BlankText()
Dim i As Byte
For i = 0 To txt.Count - 1
    If i <> 3 Then
        txt(i).TEXT = ""
    End If
Next i
End Sub
Private Sub MoveRec()
'On Error GoTo error1
If Master.RecordCount > 0 Then
txt(SlipNo) = Master!SearchCode
txt(SlipDate) = Master!V_DATE
txt(TruckNo) = XNull(Master!TruckNo)
txt(TruckNo).Tag = XNull(Master!Site_Code)
txt(StDate) = IIf(IsNull(Master!StDate), "", Master!StDate)
txt(ReachDate) = IIf(IsNull(Master!ReachDate), "", Master!ReachDate)
txt(ReachTime) = XNull(Master!ReachTime)
txt(LoadDate) = IIf(IsNull(Master!LoadDate), "", Master!LoadDate)
txt(LoadTime) = XNull(Master!LoadTime)
If XNull(Master!LoadPlace) <> "" Then
    txt(LoadPlace) = GCn.Execute("Select CityName from City where CityCode='" & Master!LoadPlace & "'").Fields(0).Value
    txt(LoadPlace).Tag = XNull(Master!LoadPlace)
End If
txt(LoadItem) = XNull(Master!LoadItem)
txt(LoadQty) = Format(VNull(Master!LoadQty), "0.00")
txt(LoadWeighingSNo) = VNull(Master!LoadWeighingSNo)
txt(LoadWeighingSDate) = IIf(IsNull(Master!LoadWeighingSDate), "", Master!LoadWeighingSDate)
txt(LoadWeighingSRemarks) = XNull(Master!LoadWeighingSRemarks)
txt(UnloadingDate) = IIf(IsNull(Master!UnloadingDate), "", Master!UnloadingDate)
txt(UnloadingTime) = XNull(Master!UnloadingTime)
If XNull(Master!UnloadingPlace) <> "" Then
    txt(UnloadingPlace) = GCn.Execute("Select CityName from City where CityCode='" & Master!UnloadingPlace & "'").Fields(0).Value
    txt(UnloadingPlace).Tag = XNull(Master!UnloadingPlace)
End If
txt(UnLoadWeighingSNo) = VNull(Master!UnLoadWeighingSNo)
txt(UnLoadWeighingSDate) = IIf(IsNull(Master!UnLoadWeighingSDate), "", Master!UnLoadWeighingSDate)
txt(UnLoadWeighingSRemarks) = XNull(Master!UnLoadWeighingSRemarks)
txt(ParchiExp) = Format(VNull(Master!ParchiExp), "0.00")
txt(MeterReading) = VNull(Master!MeterReading)
txt(BDownFromDate) = IIf(IsNull(Master!BDownFromDate), "", Master!BDownFromDate)
txt(BDownFromTime) = XNull(Master!BDownFromTime)
txt(BDownToDate) = IIf(IsNull(Master!BDownToDate), "", Master!BDownToDate)
txt(BDownToTime) = XNull(Master!BDownToTime)
txt(BDownCost) = Format(VNull(Master!BDownCost), "0.00")
txt(BDownRemarks) = XNull(Master!BDownRemarks)
txt(AdditionalExp) = Format(VNull(Master!AdditionalExp), "0.00")
txt(SDesiel) = Format(VNull(Master!SDesiel), "0.00")
txt(ADesiel) = Format(VNull(Master!ADesiel), "0.00")
txt(Frt) = Format(VNull(Master!Freight), "0.00")
txt(AddKMS) = Format(VNull(Master!AddKMS), "0.00")
txt(AddFrt) = Format(VNull(Master!AddFreight), "0.00")

Else
    Call BlankText
End If
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Ini_Grid()
Dim i As Byte
DGCode.left = txt(TruckNo).left: DGCode.top = txt(TruckNo).top + txt(TruckNo).height + 10
DGCity.left = txt(LoadPlace).left: DGCity.top = txt(LoadPlace).top + txt(LoadPlace).height + 10
End Sub
Private Sub Disp_Text(enb As Boolean)
Dim i As Integer
For i = 0 To txt.Count - 1
'    If i <> 3 Then
        txt(i).Enabled = enb
        txt(i).ForeColor = CtrlFColOrg
'    End If
Next
txt(SlipNo).Enabled = False
txt(StDate).Enabled = False
txt(UnloadingPlace).Enabled = False
txtDisabled_Color Me
End Sub
Private Sub Grid_Hide()
    If DGCity.Visible = True Then DGCity.Visible = False
    If DGCode.Visible = True Then DGCode.Visible = False
End Sub
'Private Function ProcAcPost(Optional CheckCtrls As Boolean) As Boolean
'On Error GoTo lblExit
'        Dim MsgStr$, rstSubCode As ADODB.Recordset, rsTemp As ADODB.Recordset
'        Dim mGTotAmt As Double, mNarr$, mDocId$, mSubCode
'
'        Set rstSubCode = New ADODB.Recordset
'        rstSubCode.CursorLocation = adUseClient
'        rstSubCode.Open "Select SubCode  From Vehicle where lorryNo='" & txt(TruckNo) & "'", GCn, adOpenStatic, adLockReadOnly
'        If XNull(rstSubCode!SubCode) = "" Then
'            MsgStr = "Please Define Vehicle A/C Name in Vehicle Master" & vbCrLf & "A/c Posting Aborted !"
'            ProcAcPost = False
'            GoTo lblExit
'        End If
'        'A/c Posting related declarations
'        Dim i As Integer, mBookDocID$
'        Dim LedgAry(7) As LedgRec, mResult As Byte
'
'        mDocId = PubDivCode & PubSiteCode & PubSiteCode & mVtype & mVPrefix & Space(8 - Len(txt(SlipNo))) & txt(SlipNo)
'
'        mNarr = "From Vehicle Trip Entry No. " & txt(SlipNo) & " Dt. " & txt(SlipDate) & " and Truck No " & txt(TruckNo)
'        i = 0
'        'Debit to Vehicle A/C
'        LedgAry(i).SubCode = rstSubCode!SubCode
'        LedgAry(i).AmtDr = Round(Val(txt(FuelValue)), 2)
'        LedgAry(i).Narration = mNarr
'        'Credit to Fuel Supplier A/C
'        i = i + 1
'        LedgAry(i).SubCode = txt(FuelParty).Tag
'        LedgAry(i).AmtCr = Round(Val(txt(FuelValue)), 2)
'        LedgAry(i).Narration = mNarr
'
'        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, mDocId, CDate(txt(SlipDate)), mNarr)
'        If mResult <> 1 Then
'            MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
'            ProcAcPost = False
'        Else
'            ProcAcPost = True
'        End If
'lblExit:
'If MsgStr <> "" Then
'    MsgBox MsgStr, vbCritical, "A/c Posting"
'ElseIf err.NUMBER > 0 Then
'    MsgBox err.Description, vbCritical, "A/c Posting"
'End If
'Set rsTemp = Nothing
'End Function
'

