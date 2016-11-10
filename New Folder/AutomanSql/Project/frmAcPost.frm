VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmAcPost 
   Appearance      =   0  'Flat
   BackColor       =   &H00CFE0E0&
   Caption         =   "Spare A/c Posting Module"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11460
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
   ScaleHeight     =   7185
   ScaleWidth      =   11460
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Cash"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   8
      Left            =   2835
      TabIndex        =   10
      Top             =   3495
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Credit"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   9
      Left            =   2835
      TabIndex        =   11
      Top             =   3735
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Cash"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   0
      Left            =   2835
      TabIndex        =   4
      Top             =   1740
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Credit"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   1
      Left            =   2835
      TabIndex        =   5
      Top             =   1980
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Cash"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   2
      Left            =   2835
      TabIndex        =   6
      Top             =   2310
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Credit"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   3
      Left            =   2835
      TabIndex        =   7
      Top             =   2565
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Cash"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   4
      Left            =   2835
      TabIndex        =   8
      Top             =   2925
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Credit"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   5
      Left            =   2835
      TabIndex        =   9
      Top             =   3165
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Issue"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   11
      Left            =   2835
      TabIndex        =   13
      Top             =   4320
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Receipts"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   10
      Left            =   2835
      TabIndex        =   12
      Top             =   4080
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Sale"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   13
      Left            =   2835
      TabIndex        =   17
      Top             =   6000
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Purchase"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   12
      Left            =   2835
      TabIndex        =   16
      Top             =   5745
      Width           =   1170
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Credit"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   7
      Left            =   2835
      TabIndex        =   15
      Top             =   5205
      Width           =   1125
   End
   Begin VB.CheckBox ChkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Cash"
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   6
      Left            =   2835
      TabIndex        =   14
      Top             =   4950
      Width           =   1125
   End
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2175
      Left            =   7650
      Negotiate       =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   -690
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
         MarqueeStyle    =   4
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   705.26
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   3660
      TabIndex        =   3
      Top             =   1110
      Width           =   2475
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Exit"
      Height          =   570
      Index           =   1
      Left            =   8970
      TabIndex        =   19
      Top             =   4365
      Width           =   1890
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Start Posting"
      Height          =   570
      Index           =   0
      Left            =   8970
      TabIndex        =   18
      Top             =   3780
      Width           =   1890
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
      Height          =   225
      Index           =   2
      Left            =   3660
      TabIndex        =   2
      Top             =   870
      Width           =   1230
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
      Height          =   225
      Index           =   1
      Left            =   3660
      TabIndex        =   1
      Text            =   "29/Dec/2003"
      Top             =   630
      Width           =   1230
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   661
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
      Height          =   270
      Index           =   0
      Left            =   6330
      MaxLength       =   8
      TabIndex        =   20
      Top             =   495
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4.Sale Returns"
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
      Index           =   30
      Left            =   840
      TabIndex        =   53
      Top             =   3510
      Width           =   1230
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   25
      Left            =   4185
      TabIndex        =   52
      Top             =   3510
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   26
      Left            =   4185
      TabIndex        =   51
      Top             =   3750
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3. Vehicle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   29
      Left            =   675
      TabIndex        =   50
      Top             =   5475
      Width           =   945
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2. Workshop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   28
      Left            =   675
      TabIndex        =   49
      Top             =   4680
      Width           =   1170
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.Spares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   27
      Left            =   675
      TabIndex        =   48
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   24
      Left            =   4185
      TabIndex        =   47
      Top             =   6015
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   23
      Left            =   4185
      TabIndex        =   46
      Top             =   5760
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   22
      Left            =   4185
      TabIndex        =   45
      Top             =   4335
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   21
      Left            =   4185
      TabIndex        =   44
      Top             =   4095
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   20
      Left            =   4185
      TabIndex        =   43
      Top             =   5220
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   19
      Left            =   4185
      TabIndex        =   42
      Top             =   4965
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   18
      Left            =   4185
      TabIndex        =   41
      Top             =   3180
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   17
      Left            =   4185
      TabIndex        =   40
      Top             =   2940
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   16
      Left            =   4185
      TabIndex        =   39
      Top             =   2580
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   15
      Left            =   4185
      TabIndex        =   38
      Top             =   2325
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   14
      Left            =   4185
      TabIndex        =   37
      Top             =   1995
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Not Completed*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   13
      Left            =   4185
      TabIndex        =   36
      Top             =   1755
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2. Purchase Returns"
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
      Index           =   9
      Left            =   840
      TabIndex        =   35
      Top             =   2325
      Width           =   1695
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5. Stock Transfer"
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
      Index           =   8
      Left            =   840
      TabIndex        =   34
      Top             =   4095
      Width           =   1380
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name :"
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
      Height          =   225
      Index           =   12
      Left            =   9810
      TabIndex        =   33
      Top             =   3030
      Width           =   960
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name :"
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
      Index           =   11
      Left            =   7680
      TabIndex        =   32
      Top             =   3030
      Width           =   960
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name :"
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
      Index           =   10
      Left            =   9810
      TabIndex        =   31
      Top             =   2340
      Width           =   960
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total No.of Records >>"
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
      Left            =   7680
      TabIndex        =   30
      Top             =   2340
      Width           =   1890
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   7695
      TabIndex        =   29
      Top             =   1920
      Width           =   960
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name :"
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
      Index           =   0
      Left            =   2580
      TabIndex        =   27
      Top             =   1110
      Width           =   960
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.Job Close"
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
      Left            =   840
      TabIndex        =   26
      Top             =   4965
      Width           =   990
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3.Counter Sale"
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
      Index           =   6
      Left            =   840
      TabIndex        =   25
      Top             =   2940
      Width           =   1230
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.Purchases"
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
      Index           =   5
      Left            =   840
      TabIndex        =   24
      Top             =   1755
      Width           =   1050
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
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
      Left            =   3240
      TabIndex        =   23
      Top             =   870
      Width           =   300
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date From :"
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
      Left            =   2580
      TabIndex        =   22
      Top             =   630
      Width           =   960
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
      Height          =   270
      Left            =   5670
      TabIndex        =   21
      Top             =   495
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "frmAcPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMRevDisTBPer As Double, mMRevDisTPPer As Double
Dim mTBDisAmtMRP As Double, mTPDisAmtMRP As Double
Dim mMRPTax As Double, mMRPTaxSur As Double, mMRPTOT As Double, mMRPReSales As Double
Dim mMRPLubeTB As Double, mMRPLubeTP  As Double

Dim RsSite As ADODB.Recordset
Dim rsCtrlAc As ADODB.Recordset
Dim rsCtrlAcLab As ADODB.Recordset

Dim mVType As String, mVPrefix As String
Dim ForSiteCode As String

Private Const ChalVType As String = "SXGR"
Private Const SprTrfRectVType As String = "SXGRT"
Private Const SprPurCashVType As String = "SXPIC"
Private Const SprPurCrVType As String = "SXPIR"
Private Const SprPRetCashVType As String = "SYPRC"
Private Const SprPRetCrVType As String = "SYPRR"
Private Const TrfRetRecVType As String = "SYPRT"

Private Const SprTrfChalType As String = "SYSCT"
Private Const SprSalCrVType As String = "SYSIR"
Private Const SprSalCashVType As String = "SYSIC"
Private Const RetCashSalVType As String = "SXSRC"
Private Const RetCrSalVType As String = "SXSRR"
Private Const RetTrfIssVType As String = "SXSRT"

Private Const LabourCashVtype As String = "W_LIC"
Private Const JobSalCashVType As String = "W_SIC"
Private Const LabourCrVtype As String = "W_LIR"
Private Const JobSalCrVType As String = "W_SIR"
Private Const VehPurVType As String = "V_PB"
Private Const VehSalVType As String = "V_SB"

'grid color scheme
Private Const CellBackColLeave As String = &HC8E8DA
Private Const GridBackColorBkg As String = &HBAD3C9
Private Const BackColorSelEnter As String = &HFEE0FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

' Under observation
Dim VoucherEditFlag As Boolean                  ' Used for whether we can edit voucher no or not
' End Under observation

Private Const DocID As Byte = 0                 ' Doc.ID
Private Const DateFrom As Byte = 1
Private Const DateTo As Byte = 2
Private Const SiteCode As Byte = 3

Private Const FromVno As Byte = 0
Private Const ToVno As Byte = 1
Private Const VType1 As Byte = 2

Private Const ChkSprPurCash As Byte = 0
Private Const ChkSprPurCr As Byte = 1
Private Const ChkSprPurRetCash As Byte = 2
Private Const ChkSprPurRetCr As Byte = 3
Private Const ChkSprCouCash As Byte = 4
Private Const ChkSprCouCr As Byte = 5
Private Const ChkJobCash As Byte = 6
Private Const ChkJobCr As Byte = 7
Private Const ChkSprCouRetCash As Byte = 8
Private Const ChkSprCouRetCr As Byte = 9
Private Const ChkSprTrfRect As Byte = 10
Private Const ChkSprTrfIss As Byte = 11
Private Const ChkVehPur As Byte = 12
Private Const ChkVehSal As Byte = 13

Private Const Checked As Byte = 1

Private Const StartPost As Byte = 0
Private Const ExitPost As Byte = 1

'* Used for clear all text boxes used in the form
Private Sub BlankText()
Dim i As Integer
    For i = 0 To txt.Count - 1
        txt(i).TEXT = ""
        txt(i).Tag = ""
    Next i
    txt(DocID).Tag = ""
    LblVPrefix.CAPTION = ""
End Sub

Private Sub Cmd_Click(Index As Integer)
If txt(DateFrom) = "" Or txt(DateTo) = "" Then
    MsgBox "Please Enter Valid Date!", vbInformation, "Invalid Date"
    txt(DateFrom).SetFocus
    Exit Sub
End If

If CDate(txt(DateFrom)) > CDate(txt(DateTo)) Then
    MsgBox "Please Enter Valid Date!", vbInformation, "Invalid Date"
    txt(DateFrom).SetFocus
    Exit Sub
End If
If Index = StartPost Then
    Cmd(StartPost).Enabled = False
    If ChkBox(ChkSprPurCash).Value = Checked Then
        PostSprPur "P", SprPurCashVType, "Spare Purchase (Cash)", 13
    End If
    If ChkBox(ChkSprPurCr).Value = Checked Then
        PostSprPur "P", SprPurCrVType, "Spare Purchase (Cr)", 14
    End If
    If ChkBox(ChkSprPurRetCash).Value = Checked Then
        PostSprPur "R", SprPRetCashVType, "Spare Purchase Return (Cash)", 15
    End If
    If ChkBox(ChkSprPurRetCr).Value = Checked Then
        PostSprPur "R", SprPRetCrVType, "Spare Purchase Return (Cr)", 16
    End If
    If ChkBox(ChkSprCouCash).Value = Checked Then
        PostSprSalCou "S", SprSalCashVType, "Spare Sale Counter(Cash)", 17
    End If
    If ChkBox(ChkSprCouCr).Value = Checked Then
        PostSprSalCou "S", SprSalCrVType, "Spare Sale Counter(Cr)", 18
    End If
    If ChkBox(ChkSprCouRetCash).Value = Checked Then
        PostSprSalCou "R", RetCashSalVType, "Spare Sale Ret Counter(Cash)", 25
    End If
    If ChkBox(ChkSprCouRetCr).Value = Checked Then
        PostSprSalCou "R", RetCrSalVType, "Spare Sale Ret Counter(Cr)", 26
    End If
    '***
    If ChkBox(ChkSprTrfRect).Value = Checked Then
        PostSprTrfRect "R", SprTrfRectVType, "Spare Transfer Rect", 21
    End If
    If ChkBox(ChkSprTrfIss).Value = Checked Then
        PostSprTrfIssue "I", SprTrfChalType, "Spare Transfer Issue", 22
    End If
    '***
    If ChkBox(ChkJobCash).Value = Checked Then
        PostJobClose "S", JobSalCashVType, "Workshop Sale (Cash)", 19
    End If
    If ChkBox(ChkJobCr).Value = Checked Then
        PostJobClose "S", JobSalCrVType, "Workshop Sale (Cr)", 20
    End If
    If ChkBox(ChkVehPur).Value = Checked Then
        PostVehPur "P", VehPurVType, "Vehicle Purchase", 23
    End If
    If ChkBox(ChkVehSal).Value = Checked Then
        PostVehSal "S", VehSalVType, "Vehicle Sale", 24
    End If
    Cmd(StartPost).Enabled = True
Else
    Unload Me
End If
End Sub

Private Sub DGSite_Click()
    If RsSite.RecordCount > 0 Then
        txt(SiteCode).TEXT = RsSite!Name
        txt(SiteCode).Tag = RsSite!Code
    End If
    txt(SiteCode).SetFocus
    DgSite.Visible = False
End Sub

Private Sub Form_Activate()
Dim UnLoadFrm As Boolean, MsgStr$
If rsCtrlAc.RecordCount <= 0 Then
    MsgStr = "No Records in Spare A/c Controls"
    UnLoadFrm = True
End If
If rsCtrlAc!SprCash_Ac = "" Then
    MsgStr = "Please Fill Spare Purchase "
    UnLoadFrm = True
End If
If rsCtrlAc!SprSalTP_Ac = "" Or _
    rsCtrlAc!OilSalTB_Ac = "" Or rsCtrlAc!OilSalTP_Ac = "" Or _
    rsCtrlAc!SprCash_Ac = "" Or rsCtrlAc!SprDiscTB_Ac = "" Or rsCtrlAc!SprGenSur_Ac = "" Or _
    rsCtrlAc!Transportation_Ac = "" Or rsCtrlAc!ReSaleTax_Ac = "" Or _
    rsCtrlAc!MiscChrg_Ac = "" Or rsCtrlAc!TOTax_Ac = "" Or rsCtrlAc!SprROff_Ac = "" Then
    MsgStr = "Please Fill Spare"
    UnLoadFrm = True
End If
'EOF Spare A/c control checking

If UnLoadFrm Then
    MsgBox "Spare A/c Posting Form Loading Aborted !" & vbCrLf & MsgStr & " A/c Controls through Utility Menu", vbInformation, "Validation"
    Unload Me
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
Dim i As Byte
TopCtrl1.Tag = PubUParam: WinSetting Me
    For i = 0 To txt.Count - 1
        txt(i).BackColor = CtrlBColOrg '&HDFF4F2
        txt(i).ForeColor = CtrlFColOrg
    Next
    DgSite.left = (Me.width - (DgSite.width - mRtScale)): DgSite.top = mTopScale
'    DGGod.left = 6630: DGGod.top = mTopScale: DGGod.Height = FGrid.top - mTopScale
lblRefresh
'    mVType = SalCashVType
    ForSiteCode = PubSiteCode
    txt(DateFrom) = PubLoginDate
    
    'A/c Pstong Control Checking
    Set rsCtrlAc = New ADODB.Recordset
    rsCtrlAc.CursorLocation = adUseClient
'    rsCtrlAc.Open "Select SprSalTP_Ac,OilSalTB_Ac,OilSalTP_Ac,CSSprAc,SprGenSur_Ac,ReSaleTax_Ac,SprCash_Ac,SprDiscTB_Ac,Transportation_Ac,MiscChrg_Ac,TOTax_Ac,SprROff_Ac From AcControls", G_FACN, adOpenDynamic, adLockOptimistic
    rsCtrlAc.Open "Select * From AcControls where Div_Code='" & PubDivCode & "'", G_FaCn, adOpenDynamic, adLockOptimistic
    'eof checking
    Set RsSite = New ADODB.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select site_code as code,site_desc as name from site order by site_desc", GCn, adOpenDynamic, adLockOptimistic
    Set DgSite.DataSource = RsSite

Exit Sub
ELoop:
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
    Set rsCtrlAc = Nothing
    Set rsCtrlAcLab = Nothing
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
On Error GoTo ELoop
Ctrl_GetFocus txt(Index)
Select Case Index
    Case SiteCode
        Set DgSite.DataSource = RsSite
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Then Exit Sub
        If txt(Index).TEXT = "" Then
            RsSite.MoveFirst
            RsSite.FIND "code ='" & PubSiteCode & "'"
            txt(Index).Tag = RsSite!Code
            txt(Index).TEXT = RsSite!Name
        Else
            If txt(Index).TEXT <> RsSite!Name Then
                RsSite.MoveFirst
                RsSite.FIND "name ='" & txt(Index).TEXT & "'"
            End If
        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case SiteCode
        DGridTxtKeyDown DgSite, txt, Index, RsSite, KeyCode, False, 1
End Select
    If DgSite.Visible = False Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
        If KeyCode = vbKeyUp Then
            If Index <> DateFrom Then Ctrl_UpKeyDown KeyCode, Shift
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
    Case SiteCode
        If DgSite.Visible = True Then DGridTxtKeyPress txt, Index, RsSite, KeyAscii, "Name"
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
    Select Case Index
        Case DateFrom, DateTo
            txt(Index).TEXT = RetDate(txt(Index))
            Cancel = Not CheckFinYear(txt(Index))
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub lblRefresh()
    Lbl(1) = ""
    Lbl(1).Refresh
    Lbl(10) = ""
    Lbl(10).Refresh
    Lbl(11) = ""
    Lbl(11).Refresh
    Lbl(12) = ""
    Lbl(12).Refresh
End Sub

Private Sub PostSprPur(TrnType$, mVType$, TitleStr$, lblIndex As Integer)
Dim xNetAmt As Double, xEntryTaxAmt As Double, xTransportation As Double
'A/c Posting related declarations
Dim LedgAry() As LedgRec
Dim mResult As Byte, mCommNarr$, mNarr$, TaxSQL$, i As Integer, j As Integer
Dim mSprPurPfx$, mFADocID$, mTxtDocID$, PartyAcCode$
Dim rsSPPurch As ADODB.Recordset, mTrans As Boolean

lblRefresh
Lbl(1) = TitleStr
Lbl(1).Refresh

Lbl(lblIndex) = "Posting in progress..."
Lbl(lblIndex).Visible = True
Lbl(lblIndex).Refresh

If TrnType = "P" Then
    mSprPurPfx = "PPPPP"
Else
    mSprPurPfx = "QQQQQ"
End If

If mVType = SprPurCashVType Or _
    mVType = SprPRetCashVType Then
    GSQL = "Select distinct V_Date,'' as DocID from SP_Purch where left(DocID,1)='" & PubDivCode & "' and trim(" & cMID("DocID", "4", "5") & ")='" & mVType & "' and V_Date >=#" & txt(DateFrom) & "# and V_Date<=#" & txt(DateTo) & "# and CancelYN=0 Order By V_Date"
Else
    GSQL = "Select distinct V_Date,DocID,Transportation,EntryTaxAmt,Party_Code,Party_Doc_No,Party_Doc_Date from SP_Purch where left(DocID,1)='" & PubDivCode & "' and trim(" & cMID("DocID", "4", "5") & ")='" & mVType & "' and V_Date >=#" & txt(DateFrom) & "# and V_Date<=#" & txt(DateTo) & "# and CancelYN=0 Order By V_Date,DocID"
End If
Set rsSPPurch = GCn.Execute(GSQL)
If rsSPPurch.RecordCount <= 0 Then
    Lbl(lblIndex) = "No Purchase Records for Posting!"
    MsgBox Lbl(lblIndex), vbInformation, "No Records!"
    GoTo lblExit
End If
Lbl(10) = rsSPPurch.RecordCount
Lbl(10).Refresh

mTrans = True
GCn.BeginTrans
GCnFaS.BeginTrans

'Start Ledger Posting
Do While rsSPPurch.EOF = False
    '**
    Erase LedgAry
    mCommNarr = ""
    mResult = 0
    mNarr = ""
    TaxSQL = ""
    i = 0
    j = 0
    xNetAmt = 0
    xEntryTaxAmt = 0
    xTransportation = 0
    '****
    Lbl(11) = rsSPPurch!V_DATE
    Lbl(11).Refresh
    
    If mVType = SprPurCashVType Or _
        mVType = SprPRetCashVType Then
        
        mTxtDocID = PubDivCode & PubSiteCode & txt(SiteCode).Tag & mVType
        xEntryTaxAmt = VNull(GCn.Execute("select sum(EntryTaxAmt) from SP_Purch " & _
                "where V_Date=#" & rsSPPurch!V_DATE & "# and left(DocID,8)='" & mTxtDocID & "'").Fields(0).Value)
        xTransportation = VNull(GCn.Execute("select sum(Transportation) from SP_Purch " & _
                "where V_Date=#" & rsSPPurch!V_DATE & "# and left(DocID,8)='" & mTxtDocID & "'").Fields(0).Value)
        GSQL = "select TF.PurSal_Ac_Code,sum(NET_AMT+EntryTaxAmt+Transportation) as NetAmt " & _
            "from SP_Purch " & _
            "left join TaxFormsAc as TF on SP_Purch.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code " & _
            "where V_Date=#" & rsSPPurch!V_DATE & "# and left(Docid,8)='" & mTxtDocID & _
            "' Group by TF.PurSal_Ac_Code"
        
        If TrnType = "P" Then
            mNarr = "Through Spare Cash Purchase (Daily Posting)"
        Else    'Purchase Return Cash
            mNarr = "Through Spare Purchase Return Cash (Daily Posting)"
        End If
        mCommNarr = mNarr & " [Common]"
        'Undelete old Posting (individual if any)
        'LedgerUnPost GCnFaS, Txt(TxtDocId)
        'Create FA DocID for Daily Posting
        mFADocID = mTxtDocID & mSprPurPfx & "  " & Format(PubStartDate, "yy") & Format(rsSPPurch!V_DATE, "mmdd")
        PartyAcCode = PubSprCashAc ', Txt(Party).Tag)
    Else
        PartyAcCode = rsSPPurch!Party_code
        mFADocID = rsSPPurch!DocID
        '**
        mNarr = "Cr Purchase" & IIf(TrnType = "P", " Return", "")
        '**
        If rsSPPurch!Party_Doc_No <> "" Then
            mNarr = mNarr & " Party Document No." & rsSPPurch!Party_Doc_No
        End If
        If rsSPPurch!Party_Doc_Date <> "" Then
            mNarr = mNarr & " Date " & rsSPPurch!Party_Doc_Date
        End If
        mCommNarr = mNarr & " [Common]"
        xEntryTaxAmt = VNull(rsSPPurch!EntryTaxAmt)
        xTransportation = VNull(rsSPPurch!Transportation)
        GSQL = "select TF.PurSal_Ac_Code,sum(NET_AMT+EntryTaxAmt+Transportation) as NetAmt " & _
            "from SP_Purch " & _
            "left join TaxFormsAc as TF on SP_Purch.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code " & _
            "where docid='" & mFADocID & _
            "' group by TF.PurSal_Ac_Code"
    End If
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    '*** pURCHASE Amount Row
'   0.Purchase A/c
'   1.Party A/c or Cash A/c
'    Dim LedgAry() As LedgRec
    
'*********
    i = -1
    Do While GRs.EOF = False
        If IsNull(GRs!PurSal_Ac_Code) Or GRs!PurSal_Ac_Code = "" Then
            MsgBox "Please Define Purchase A/c in Tax Forms " & GRs!PurSal_Ac_Code & vbCrLf & "A/c Psoting Aborted", vbCritical, "A/c Posting"
            GoTo lblExit
        End If
        If i = -1 Then
            ReDim Preserve LedgAry(1)
            i = 0
        Else
            i = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(i)
        End If
        LedgAry(i).SubCode = GRs!PurSal_Ac_Code
        LedgAry(i).ContraSub = PartyAcCode
        If TrnType = "P" Then
            LedgAry(i).AmtDr = IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
        Else
            LedgAry(i).AmtCr = IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
        End If
        LedgAry(i).Narration = mNarr
        
        xNetAmt = xNetAmt + IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
        GRs.MoveNext
    Loop
    If xTransportation <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!SprPurTrans_Ac
        If TrnType = "P" Then
            LedgAry(i).AmtCr = xTransportation
        Else
            LedgAry(i).AmtDr = xTransportation
        End If
        LedgAry(i).Narration = mNarr
    End If
    If xEntryTaxAmt <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!EntryTax_Ac
        If TrnType = "P" Then
            LedgAry(i).AmtCr = xEntryTaxAmt
        Else
            LedgAry(i).AmtDr = xEntryTaxAmt
        End If
        LedgAry(i).Narration = mNarr
    End If

    i = UBound(LedgAry) + 1
    ReDim Preserve LedgAry(i)
    LedgAry(i).SubCode = PartyAcCode
    If TrnType = "P" Then
        LedgAry(i).AmtCr = (xNetAmt - (xEntryTaxAmt + xTransportation))
    Else
        LedgAry(i).AmtDr = (xNetAmt - (xEntryTaxAmt + xTransportation))
    End If
    LedgAry(i).Narration = mNarr
    
    mResult = LedgerPost("A", LedgAry, GCnFaS, mFADocID, CDate(rsSPPurch!V_DATE), mCommNarr)
    If mResult <> 1 Then
        MsgBox "Error in Ledger Posting" & " : " & mFADocID, vbOKOnly, "Validation"
'        ProcAcPost = True
    End If
    Lbl(12) = rsSPPurch.AbsolutePosition
    Lbl(12).Refresh

    rsSPPurch.MoveNext
Loop
GCn.CommitTrans
GCnFaS.CommitTrans
mTrans = False
Lbl(lblIndex) = "Posting completed sucessfully !"
Lbl(lblIndex).Refresh
'MsgBox TitleStr & vbCrLf & "Posting completed sucessfully !", vbInformation, "Ledger Posting"

lblExit:
    Set rsSPPurch = Nothing
    Set GRs = Nothing
    If mTrans Then
        GCn.RollbackTrans
        GCnFaS.RollbackTrans
    End If
    If err.NUMBER <> 0 Then
        MsgBox err.Description & vbCrLf & "Ledger Posting Terminated!", vbCritical
'        ProcAcPost = True
    End If
End Sub

Private Sub PostSprSalCou(TrnType$, mVType$, TitleStr$, lblIndex As Integer)
On Error GoTo lblExit
Dim xMRPSprTp As Double, xMRPOilTp As Double
Dim xSprTp As Double, xOilTp As Double
Dim mShare As Single, mShareAmt As Double, mShare2Amt As Double
Dim xNetAmt As Double, xRoundAmt As Double, xSprAmtMRPTB As Double, xSprAmtMRPTP As Double
Dim xOilAmtMRPTB As Double, xOilAmtMRPTP As Double
Dim xSprAmtTB  As Double, xSprAmtTP As Double, xOilAmtTB As Double, xOilAmtTP As Double
Dim xDisAmtTB As Double, xDisAmtTP As Double, xDisAmtMRPTB As Double, xDisAmtMRPTP As Double
Dim xGenSurAmt As Double, xTrans As Double, xTaxAmt As Double, xTaxAmtMRP As Double, xPack As Double
Dim xTurnOver As Double, xReSaleTaxAmt As Double, mFADocID$, mQRY$, PartyCode$
Dim RsTemp As ADODB.Recordset, rsTemp1 As ADODB.Recordset
'A/c Posting related declarations
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, mNarr$, TaxSQL$, i As Integer, j As Integer
Dim mSprAmtMRPTB As Double, mSprAmtTB As Double
Dim mOilAmtMRPTB As Double, mOilAmtTB As Double
Dim mTotMRPOilTB As Double, mTotOilTB As Double, mTotShareAmt As Double
Dim mShareSpr As Single, mShareAmtSpr As Double, mShare2AmtSpr As Double
Dim mTot1ShareAmt As Double, mTot2ShareAmt As Double, mTot3ShareAmt As Double
Dim mPrefix$, rsSPSal As ADODB.Recordset, mTxtDocID$, mTrans As Boolean

lblRefresh
Lbl(1) = TitleStr
Lbl(1).Refresh

Lbl(lblIndex) = "Posting in progress..."
Lbl(lblIndex).Visible = True
Lbl(lblIndex).Refresh

If TrnType = "S" Then
    mPrefix = "XXXXX"
Else
    mPrefix = "AAAAA"
End If
If mVType = SprSalCashVType Or _
    mVType = RetCashSalVType Then
     GSQL = "Select distinct V_Date,'' as DocID from SP_Sale where left(DocID,1)='" & PubDivCode & "' and trim(" & cMID("DocID", "4", "5") & ")='" & mVType & "' and V_Date >=#" & txt(DateFrom) & "# and V_Date<=#" & txt(DateTo) & "# and CancelYN=0 Order By V_Date"
Else
    GSQL = "Select distinct V_Date,DocID,Party_Code from SP_Sale where left(DocID,1)='" & PubDivCode & "' and trim(" & cMID("DocID", "4", "5") & ")='" & mVType & "' and V_Date >=#" & txt(DateFrom) & "# and V_Date<=#" & txt(DateTo) & "# and CancelYN=0 Order By V_Date,DocID"
End If
Set rsSPSal = GCn.Execute(GSQL)
If rsSPSal.RecordCount <= 0 Then
    Lbl(lblIndex) = "No Records for Posting!"
    MsgBox Lbl(lblIndex), vbInformation, "No Records!"
    GoTo lblExit
End If
Lbl(10) = rsSPSal.RecordCount
Lbl(10).Refresh

mTxtDocID = PubDivCode & PubSiteCode & txt(SiteCode).Tag & mVType

mTrans = True
GCn.BeginTrans
GCnFaS.BeginTrans

'Start Ledger Posting
Do While rsSPSal.EOF = False
    '**
    Erase LedgAry
    xMRPSprTp = 0: xMRPOilTp = 0
    xSprTp = 0: xOilTp = 0
    mShare = 0: mShareAmt = 0: mShare2Amt = 0
    xNetAmt = 0: xRoundAmt = 0: xSprAmtMRPTB = 0: xSprAmtMRPTP = 0
    xOilAmtMRPTB = 0: xOilAmtMRPTP = 0
    xSprAmtTB = 0: xSprAmtTP = 0: xOilAmtTB = 0: xOilAmtTP = 0
    xDisAmtTB = 0: xDisAmtTP = 0: xDisAmtMRPTB = 0: xDisAmtMRPTP = 0
    xGenSurAmt = 0: xTrans = 0: xTaxAmt = 0: xTaxAmtMRP = 0: xPack = 0
    xTurnOver = 0: xReSaleTaxAmt = 0: mFADocID = "": mQRY = "": PartyCode = ""
    'A/c Posting related declarations
    mSprAmtMRPTB = 0: mSprAmtTB = 0
    mOilAmtMRPTB = 0: mOilAmtTB = 0
    mTotMRPOilTB = 0: mTotOilTB = 0: mTotShareAmt = 0
    mShareSpr = 0: mShareAmtSpr = 0: mShare2AmtSpr = 0
    mTot1ShareAmt = 0: mTot2ShareAmt = 0: mTot3ShareAmt = 0
    '**
    mCommNarr = ""
    mResult = 0
    mNarr = ""
    i = 0
    j = 0
    '****
    Lbl(11) = rsSPSal!V_DATE
    Lbl(11).Refresh
    
    TaxSQL = "select TF.Tax_Ac_Code,TF.Sur_Ac_Code,sum(Tax_Amt+Tax_AmtMRP) as TaxAmt,sum(Tax_Sur_Amt+TaxSur_AmtMRP) as TaxSurAmt " & _
        " from SP_Sale left join TaxFormsAc as TF on Sp_Sale.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code"
    
    If mVType = SprSalCashVType Or _
        mVType = RetCashSalVType Then
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "inner join TaxFormsAc TF on Sp_Sale.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code " & _
            "where V_Date=#" & rsSPSal!V_DATE & "# and left(docid,8)='" & mTxtDocID & _
            "' Group by TF.PurSal_Ac_Code"
            
        mQRY = "select " & _
            "sum(Total_Amt) as NetAmt,sum(rounded) as RoundAmt," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(SprAmt_MRP_TP) as SprAmtMRPTP, " & _
            "sum(OilAmt_MRP_TB) as OilAmtMRPTB, sum(OilAmt_MRP_TP) as OilAmtMRPTP, " & _
            "sum(SprAmt_TB) as SprAmtTB, sum(SprAmt_TP) as SprAmtTP, " & _
            "sum(OilAmt_TB) as OilAmtTB, sum(OilAmt_TP) as OilAmtTP, " & _
            "sum(D_Amt_TB) as DisAmtTB, sum(D_Amt_TP) as DisAmtTP, " & _
            "sum(D_Amt_MRP_TB) as DisAmtMRPTB, sum(D_Amt_MRP_TP) as DisAmtMRPTP," & _
            "sum(Gen_Sur_Amt) as GenSurAmt,sum(Trans_Amt) as Trans," & _
            "sum(Tax_Amt+Tax_Sur_Amt+Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmt," & _
            "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, sum(TOT_Amt) as TurnOver, " & _
            "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
            "from SP_Sale " & _
            "inner join TaxFormsAc TF on Sp_Sale.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code " & _
            "where V_Date=#" & rsSPSal!V_DATE & "# and left(docid,8)='" & mTxtDocID & "'"
        'for tax
        TaxSQL = TaxSQL & " where  V_Date=#" & rsSPSal!V_DATE & "# and left(docid,8)='" & mTxtDocID & _
            "' Group by TF.Tax_Ac_Code,TF.Sur_Ac_Code"
        mNarr = "Through Counter Cash Sale (Daily Posting)"
        mCommNarr = mNarr & " [Common]"
        mFADocID = mTxtDocID & mPrefix & "  " & Format(rsSPSal!V_DATE, "yymmdd")
        PartyCode = PubSprCashAc
    Else
        PartyCode = rsSPSal!Party_code
        mFADocID = rsSPSal!DocID
        mNarr = "Through Counter Cr Sale"
        mCommNarr = mNarr & " [Common]"
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "inner join TaxFormsAc TF on Sp_Sale.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code " & _
            " where docid='" & rsSPSal!DocID & _
            "' group by TF.PurSal_Ac_Code"

        mQRY = "select " & _
            "sum(Total_Amt) as NetAmt,sum(rounded) as RoundAmt," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(SprAmt_MRP_TP) as SprAmtMRPTP, " & _
            "sum(OilAmt_MRP_TB) as OilAmtMRPTB, sum(OilAmt_MRP_TP) as OilAmtMRPTP, " & _
            "sum(SprAmt_TB) as SprAmtTB, sum(SprAmt_TP) as SprAmtTP, " & _
            "sum(OilAmt_TB) as OilAmtTB, sum(OilAmt_TP) as OilAmtTP, " & _
            "sum(D_Amt_TB) as DisAmtTB, sum(D_Amt_TP) as DisAmtTP, " & _
            "sum(D_Amt_MRP_TB) as DisAmtMRPTB, sum(D_Amt_MRP_TP) as DisAmtMRPTP," & _
            "sum(Gen_Sur_Amt) as GenSurAmt,sum(Trans_Amt) as Trans," & _
            "sum(Tax_Amt+Tax_Sur_Amt+Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmt," & _
            "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, sum(TOT_Amt) as TurnOver, " & _
            "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
            "from SP_Sale " & _
            "inner join TaxFormsAc as TF on Sp_Sale.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code " & _
            "where docid='" & rsSPSal!DocID & "'"
        'for tax
        TaxSQL = TaxSQL & " where docid='" & rsSPSal!DocID & _
            "' Group by TF.Tax_Ac_Code,TF.Sur_Ac_Code"
    End If
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open mQRY, GCn, adOpenStatic, adLockReadOnly
    
    'for tax purpose
   
    
    Set RsTemp = New ADODB.Recordset
    RsTemp.CursorLocation = adUseClient
    RsTemp.Open TaxSQL, GCn, adOpenStatic, adLockReadOnly

'        1.MRP Spr TB = SprAmtMRPTB-OilAmtMRPTB - part of DisAmtMRPTB - part of TaxAmtMRP
'        3.MRP Oil TB = OilAmtMRPTB - part of DisAmtMRPTB - part of TaxAmtMRP

'        2.MRP Spr TP = SprAmtMRPTP-OilAmtMRPTP - part of DisAmtMRPTP
'        4.MRP Oil TP = OilAmtMRPTP - part of DisAmtMRPTP

'        1.Spr TB = SprAmtTB - part of DisAmtTB
'        3.Oil TB = OilAmtTB - part of DisAmtTB
'        2.Spr TP = SprAmtTP - part of DisAmtTP
'        4.Oil TP = OilAmtTP - part of DisAmtTP
    xNetAmt = IIf(IsNull(rsTemp1!NetAmt), 0, rsTemp1!NetAmt): xRoundAmt = IIf(IsNull(rsTemp1!RoundAmt), 0, rsTemp1!RoundAmt)
    xSprAmtMRPTB = IIf(IsNull(rsTemp1!SprAmtMrpTB), 0, rsTemp1!SprAmtMrpTB)
    xOilAmtMRPTB = IIf(IsNull(rsTemp1!OilAmtMrpTB), 0, rsTemp1!OilAmtMrpTB)
    xSprAmtMRPTP = IIf(IsNull(rsTemp1!SprAmtMrpTP), 0, rsTemp1!SprAmtMrpTP)
    xOilAmtMRPTP = IIf(IsNull(rsTemp1!OilAmtMrpTP), 0, rsTemp1!OilAmtMrpTP)
    xSprAmtTB = IIf(IsNull(rsTemp1!SprAmtTB), 0, rsTemp1!SprAmtTB)
    xOilAmtTB = IIf(IsNull(rsTemp1!OilAmtTB), 0, rsTemp1!OilAmtTB)
    xSprAmtTP = IIf(IsNull(rsTemp1!SprAmtTP), 0, rsTemp1!SprAmtTP)
    xOilAmtTP = IIf(IsNull(rsTemp1!OilAmtTP), 0, rsTemp1!OilAmtTP)
    xDisAmtTB = IIf(IsNull(rsTemp1!DisAmtTB), 0, rsTemp1!DisAmtTB)
    xDisAmtTP = IIf(IsNull(rsTemp1!DisAmtTP), 0, rsTemp1!DisAmtTP)
    xDisAmtMRPTB = IIf(IsNull(rsTemp1!DisAmtMRPTB), 0, rsTemp1!DisAmtMRPTB)
    xDisAmtMRPTP = IIf(IsNull(rsTemp1!DisAmtMRPTP), 0, rsTemp1!DisAmtMRPTP)
    xGenSurAmt = IIf(IsNull(rsTemp1!GenSurAmt), 0, rsTemp1!GenSurAmt)
    xTrans = IIf(IsNull(rsTemp1!Trans), 0, rsTemp1!Trans)
    xTaxAmt = IIf(IsNull(rsTemp1!TaxAmt), 0, rsTemp1!TaxAmt)
    xTaxAmtMRP = IIf(IsNull(rsTemp1!TaxAmtMRP), 0, rsTemp1!TaxAmtMRP)
    xPack = IIf(IsNull(rsTemp1!Pack), 0, rsTemp1!Pack)
    xTurnOver = IIf(IsNull(rsTemp1!TurnOver), 0, rsTemp1!TurnOver)
    xReSaleTaxAmt = IIf(IsNull(rsTemp1!ReSaleTaxAmt), 0, rsTemp1!ReSaleTaxAmt)
    '*** Sale Amount Row
    i = 1
    ReDim Preserve LedgAry(1)
    '**Taxable Spr / Oil Calculation
     Do While GRs.EOF = False
        mOilAmtMRPTB = IIf(IsNull(GRs!OilAmtMrpTB), 0, GRs!OilAmtMrpTB)
        mSprAmtMRPTB = IIf(IsNull(GRs!SprAmtMrpTB), 0, GRs!SprAmtMrpTB) ' - mOilAmtMRPTB
        mSprAmtTB = IIf(IsNull(GRs!SprAmtTB), 0, GRs!SprAmtTB)
        mOilAmtTB = IIf(IsNull(GRs!OilAmtTB), 0, GRs!OilAmtTB)
        'Allocate values in their proportions
        If (mSprAmtMRPTB + mOilAmtMRPTB) <> 0 Then
            mShare = Round((mSprAmtMRPTB + mOilAmtMRPTB) * 100 / (xSprAmtMRPTB + xOilAmtMRPTB), 2)
            mShareAmt = Round(xDisAmtMRPTB * mShare / 100, 2)
            mShare2Amt = Round(xTaxAmtMRP * mShare / 100, 2)
            mShareSpr = Round(mSprAmtMRPTB * 100 / (mSprAmtMRPTB + mOilAmtMRPTB), 2)
            mShareAmtSpr = Round(mShareAmt * mShareSpr / 100, 2)
            mShare2AmtSpr = Round(mShare2Amt * mShareSpr / 100, 2)
            mTot1ShareAmt = mTot1ShareAmt + mShareAmt
            mTot2ShareAmt = mTot2ShareAmt + mShare2Amt
            If GRs.AbsolutePosition = GRs.RecordCount Then
                mShareAmt = mShareAmt + ((xDisAmtMRPTB) - mTot1ShareAmt)
                mShare2Amt = mShare2Amt + ((xTaxAmtMRP) - mTot2ShareAmt)
            End If
            mSprAmtMRPTB = mSprAmtMRPTB - (mShareAmtSpr + mShare2AmtSpr)
            mOilAmtMRPTB = mOilAmtMRPTB - ((mShareAmt + mShare2Amt) - (mShareAmtSpr + mShare2AmtSpr))
        End If
        '*****
        If (mSprAmtTB + mOilAmtTB) <> 0 Then
            mShare = Round((mSprAmtTB + mOilAmtTB) * 100 / (xSprAmtTB + xOilAmtTB), 2)
            mShareAmt = Round((xDisAmtTB - xDisAmtMRPTB) * mShare / 100, 2)
            mShareSpr = Round(mSprAmtTB * 100 / (mSprAmtTB + mOilAmtTB), 2)
            mShareAmtSpr = Round(mShareAmt * mShareSpr / 100, 2)
            mTot3ShareAmt = mTot3ShareAmt + mShareAmt
            If GRs.AbsolutePosition = GRs.RecordCount Then
                mShareAmt = mShareAmt + ((xDisAmtTB - xDisAmtMRPTB) - mTot3ShareAmt)
            End If
            mSprAmtTB = mSprAmtTB - (mShareAmtSpr)
            mOilAmtTB = mOilAmtTB - (mShareAmt - mShareAmtSpr)
        End If
        'Spare Sale A/c Taxable
        mTotMRPOilTB = mTotMRPOilTB + mOilAmtMRPTB
        mTotOilTB = mTotOilTB + mOilAmtTB
        '*****
        If mSprAmtMRPTB + mSprAmtTB <> 0 Then
            i = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(i)
            LedgAry(i).SubCode = GRs!PurSal_Ac_Code
            If TrnType = "S" Then
                LedgAry(i).AmtDr = 0
                LedgAry(i).AmtCr = Round(mSprAmtMRPTB + mSprAmtTB, 2)
            Else
                LedgAry(i).AmtDr = Round(mSprAmtMRPTB + mSprAmtTB, 2)
                LedgAry(i).AmtCr = 0
            End If
            LedgAry(i).Narration = mNarr ' & " Spare"
        End If
        GRs.MoveNext
    Loop
    If (xSprAmtMRPTP + xOilAmtMRPTP) <> 0 Then
        xMRPSprTp = xSprAmtMRPTP '- xOilAmtMRPTP
        xMRPOilTp = xOilAmtMRPTP
        mShare = Round(xMRPSprTp * 100 / (xSprAmtMRPTP + xOilAmtMRPTP), 2)
        mShareAmt = Round(xDisAmtMRPTP * mShare / 100, 2)
        xMRPSprTp = xMRPSprTp - (mShareAmt)
        xMRPOilTp = xMRPOilTp - (xDisAmtMRPTP - (mShareAmt))
    End If
    If (xSprAmtTP + xOilAmtTP) <> 0 Then
        mShare = Round(xSprAmtTP * 100 / (xSprAmtTP + xOilAmtTP), 2)
        mShareAmt = Round((xDisAmtTP - xDisAmtMRPTP) * mShare / 100, 2)
        xSprTp = xSprAmtTP - (mShareAmt)
        xOilTp = xOilAmtTP - ((xDisAmtTP - xDisAmtMRPTP) - (mShareAmt))
    End If

'   0.Party A/c or Cash A/c
'   1.Taxable Spr = MRP Spr TB + SPR TB
'   2.Taxpaid Spr = MRP Spr TP + SPR TP
'   3.Taxable Oil = MRP Oil TB + Oil TB
'   4.Taxable Oil = MRP Oil TP + Oil TP
'   5.xGenSurAmt
'   6.xPack
'   7.xTurnOver
'   8.xReSaleTaxAmt
'    Dim LedgAry() As LedgRec
    
    'Sale Party A/c
    'I = 0
    LedgAry(0).SubCode = PartyCode
    If TrnType = "S" Then
        LedgAry(0).AmtDr = Round(xNetAmt, 2)
        LedgAry(0).AmtCr = 0
    Else
        LedgAry(0).AmtCr = Round(xNetAmt, 2)
        LedgAry(0).AmtDr = 0
    End If
    LedgAry(0).Narration = mNarr
    'Spare Sale A/c Taxpaid
    If xMRPSprTp + xSprTp <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!SprSalTP_Ac
        If TrnType = "S" Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(xMRPSprTp + xSprTp, 2)
        Else
            LedgAry(i).AmtDr = Round(xMRPSprTp + xSprTp, 2)
            LedgAry(i).AmtCr = 0
        End If
        LedgAry(i).Narration = mNarr ' & " Spare"
    End If
    'Oil Sale A/c Taxable
    If mTotMRPOilTB + mTotOilTB <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!OilSalTB_Ac
        If TrnType = "S" Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(mTotMRPOilTB + mTotOilTB, 2)
        Else
            LedgAry(i).AmtDr = Round(mTotMRPOilTB + mTotOilTB, 2)
            LedgAry(i).AmtCr = 0
        End If
        LedgAry(i).Narration = mNarr ' & " Spare"
    End If
     'Oil Sale A/c Taxpaid
     If xMRPOilTp + xOilTp <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!OilSalTP_Ac
        If TrnType = "S" Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(xMRPOilTp + xOilTp, 2)
        Else
            LedgAry(i).AmtCr = 0
            LedgAry(i).AmtDr = Round(xMRPOilTp + xOilTp, 2)
        End If
        LedgAry(i).Narration = mNarr ' & " Spare"
     End If
      'GenSurAmt
     If xGenSurAmt <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!SprGenSur_Ac
        If TrnType = "S" Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(xGenSurAmt, 2)
        Else
            LedgAry(i).AmtCr = 0
            LedgAry(i).AmtDr = Round(xGenSurAmt, 2)
        End If
        LedgAry(i).Narration = mNarr ' & " Sale Tax"
     End If
    'Transportation
     If xTrans <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!Transportation_Ac
        If (xTrans > 0 And TrnType = "S") Or (xTrans < 0 And TrnType <> "S") Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(Abs(xTrans), 2)
        Else
            LedgAry(i).AmtDr = Round(Abs(xTrans), 2)
           LedgAry(i).AmtCr = 0
        End If
        LedgAry(i).Narration = mNarr '& " Transportation"
     End If
     If RsTemp.RecordCount > 0 Then
         Do While RsTemp.EOF = False
             If RsTemp!TaxAmt <> 0 Then
                i = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(i)
                LedgAry(i).SubCode = RsTemp!Tax_Ac_Code
                If (RsTemp!TaxAmt > 0 And TrnType = "S") Or (RsTemp!TaxAmt < 0 And TrnType <> "S") Then
                    LedgAry(i).AmtDr = 0
                    LedgAry(i).AmtCr = Round(Abs(RsTemp!TaxAmt), 2)
                Else
                    LedgAry(i).AmtDr = Round(Abs(RsTemp!TaxAmt), 2)
                    LedgAry(i).AmtCr = 0
                End If
                 LedgAry(i).Narration = mNarr '& " Sales Tax & Surcharge"
            End If
            If RsTemp!TaxSurAmt <> 0 Then
                i = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(i)
                LedgAry(i).SubCode = RsTemp!Sur_Ac_Code
                If (RsTemp!TaxSurAmt > 0 And TrnType = "S") Or (RsTemp!TaxSurAmt < 0 And TrnType <> "S") Then
                    LedgAry(i).AmtDr = 0
                    LedgAry(i).AmtCr = Round(Abs(RsTemp!TaxSurAmt), 2)
                Else
                    LedgAry(i).AmtDr = Round(Abs(RsTemp!TaxSurAmt), 2)
                    LedgAry(i).AmtCr = 0
                End If
                 LedgAry(i).Narration = mNarr '& " Sales Tax & Surcharge"
             End If
             '***
             RsTemp.MoveNext
         Loop
     End If
    'Misc / Packing Chrg
    If xPack <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!MiscChrg_Ac
        If (xPack > 0 And TrnType = "S") Or (xPack < 0 And TrnType <> "S") Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(Abs(xPack), 2)
        Else
            LedgAry(i).AmtDr = Round(Abs(xPack), 2)
            LedgAry(i).AmtCr = 0
        End If
        LedgAry(i).Narration = mNarr '& " Misc Charges"
    End If
    If xTurnOver <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!TOTax_Ac
        If (xTurnOver > 0 And TrnType = "S") Or (xTurnOver < 0 And TrnType <> "S") Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(Abs(xTurnOver), 2)
        Else
            LedgAry(i).AmtDr = Round(Abs(xTurnOver), 2)
            LedgAry(i).AmtCr = 0
        End If
        LedgAry(i).Narration = mNarr '& " TurnOver Amt"
    End If
    If xReSaleTaxAmt <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!ReSaleTax_Ac
        If (xReSaleTaxAmt > 0 And TrnType = "S") Or (xReSaleTaxAmt < 0 And TrnType <> "S") Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(Abs(xReSaleTaxAmt), 2)
        Else
            LedgAry(i).AmtDr = Round(Abs(xReSaleTaxAmt), 2)
            LedgAry(i).AmtCr = 0
        End If
        LedgAry(i).Narration = mNarr '& " ReSale Tax Amount"
    End If
    If xRoundAmt <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!SprROff_Ac
        If TrnType = "S" Then
            If xRoundAmt > 0 Then
                LedgAry(i).AmtDr = 0
                LedgAry(i).AmtCr = Round(xRoundAmt, 2)
            Else
                LedgAry(i).AmtDr = Round(Abs(xRoundAmt), 2)
                LedgAry(i).AmtCr = 0
            End If
        Else
            If xRoundAmt > 0 Then
                LedgAry(i).AmtDr = Round(xRoundAmt, 2)
                LedgAry(i).AmtCr = 0
            Else
                LedgAry(i).AmtDr = 0
                LedgAry(i).AmtCr = Round(Abs(xRoundAmt), 2)
            End If
        End If
        LedgAry(i).Narration = mNarr '& " Round Off"
    End If
    
    mResult = LedgerPost("A", LedgAry, GCnFaS, mFADocID, CDate(rsSPSal!V_DATE), mCommNarr)
    If mResult <> 1 Then MsgBox "Error in Ledger Posting" & " : " & mFADocID, vbOKOnly, "Validation"
    '**
    Lbl(12) = rsSPSal.AbsolutePosition
    Lbl(12).Refresh
    rsSPSal.MoveNext
Loop
GCn.CommitTrans
GCnFaS.CommitTrans
mTrans = False

Lbl(lblIndex) = "Posting completed sucessfully !"
Lbl(lblIndex).Refresh
'MsgBox TitleStr & vbCrLf & "Posting completed sucessfully !", vbInformation, "Ledger Posting"

lblExit:
    Set RsTemp = Nothing
    Set rsSPSal = Nothing
    Set GRs = Nothing
    If mTrans Then
        GCn.RollbackTrans
        GCnFaS.RollbackTrans
    End If
    If err.NUMBER <> 0 Then
        MsgBox err.Description & vbCrLf & "Ledger Posting Terminated!", vbCritical
    End If
End Sub

Private Sub PostJobClose(TrnType$, mVType$, TitleStr$, lblIndex As Integer)
'Checking Labour A/c Controls
'Checking Labour A/c Controls
Set rsCtrlAcLab = New ADODB.Recordset
rsCtrlAcLab.CursorLocation = adUseClient
rsCtrlAcLab.Open "Select SrvCash_Ac,SrvLabourTB_Ac,SrvLabour_Ac,SrvTax_Ac,SrvROff_Ac From AcControls", GCnFaW, adOpenStatic, adLockOptimistic
'EOF Labour A/c control checking
If rsCtrlAcLab.RecordCount <= 0 Then
    Lbl(lblIndex) = "No Records in Labour A/c Controls"
    MsgBox Lbl(lblIndex), vbCritical, "Insert Labour A/c Control Rec"
    Exit Sub
End If
If rsCtrlAcLab!SrvCash_Ac = "" Or rsCtrlAcLab!SrvLabourTB_Ac = "" Or _
    rsCtrlAcLab!SrvLabour_Ac = "" Or rsCtrlAcLab!SrvTax_Ac = "" Or rsCtrlAcLab!SrvROff_Ac = "" Then
    Lbl(lblIndex) = "Please Fill Labour A/c Controls"
    MsgBox Lbl(lblIndex), vbCritical, "Fill Labour A/c Controls"
    Exit Sub
End If

On Error GoTo lblExit
Dim xMRPSprTp As Double, xMRPOilTp As Double
Dim xSprTp As Double, xOilTp As Double
Dim mShare As Single, mShareAmt As Double, mShare2Amt As Double
Dim xNetAmt As Double, xRoundAmt As Double, xSprAmtMRPTB As Double, xSprAmtMRPTP As Double
Dim xOilAmtMRPTB As Double, xOilAmtMRPTP As Double
Dim xSprAmtTB  As Double, xSprAmtTP As Double, xOilAmtTB As Double, xOilAmtTP As Double
Dim xDisAmtTB As Double, xDisAmtTP As Double, xDisAmtMRPTB As Double, xDisAmtMRPTP As Double
Dim xGenSurAmt As Double, xTrans As Double, xTaxAmt As Double, xTaxAmtMRP As Double, xPack As Double
Dim xTurnOver As Double, xReSaleTaxAmt As Double, mFADocidSpr$, mFADocidLab$, mQRY$
Dim xNetLabAmt As Double, xLabAmtTB As Double, xLabAmtTP As Double, xLabDisc As Double
Dim xServTaxAmt As Double, xLabROff As Single
Dim RsTemp As ADODB.Recordset, rsTemp1 As ADODB.Recordset
'A/c Posting related declarations
Dim LedgAry() As LedgRec, LedgAryLab() As LedgRec, mCommNarr$, mLabSQL$
Dim mResult As Byte, mNarr$, TaxSQL$, i As Integer, j As Integer
Dim mSprAmtMRPTB As Double, mSprAmtTB As Double
Dim mOilAmtMRPTB As Double, mOilAmtTB As Double
Dim mTotMRPOilTB As Double, mTotOilTB As Double, mTotShareAmt As Double
Dim mShareSpr As Single, mShareAmtSpr As Double, mShare2AmtSpr As Double
Dim mTot1ShareAmt As Double, mTot2ShareAmt As Double, mTot3ShareAmt As Double
Dim RsJob As ADODB.Recordset, SpareDocID$, LabourDocID$, PartyCode$, PartyCodeLab$
Dim SepLabPost As Boolean, mTrans As Boolean

If UCase(PubSFADataPath) <> UCase(PubWFADataPath) Then
    SepLabPost = True
End If

lblRefresh
Lbl(1) = TitleStr
Lbl(1).Refresh

Lbl(lblIndex) = "Posting in progress..."
Lbl(lblIndex).Visible = True
Lbl(lblIndex).Refresh

If mVType = JobSalCashVType Then
    If PubDealerID = "1109800" Then
        If CDate(txt(DateFrom)) <= CDate(pubLockDate) Then
            MsgBox "Start Date " & txt(DateFrom) & " is less than Lock Date " & pubLockDate, vbInformation, "Works Cash Posting Locked"
            Lbl(lblIndex) = "Posting Lock Date " & pubLockDate
            GoTo lblExit
        End If
    End If
    GSQL = "Select distinct JobCloseDate from Job_Card where left(DocId,1)='" & PubDivCode & "' and JobCloseDate >=#" & txt(DateFrom) & "# and JobCloseDate<=#" & txt(DateTo) & "# and CRMemo=0 Order by JobCloseDate Desc" ''<>Null
Else
    GSQL = "Select DocID,JobCloseDate,DocId_InvSpr,DocId_InvLab,DrSpr_AcCode,DrLab_AcCode from Job_Card where left(DocId,1)='" & PubDivCode & "' and JobCloseDate >=#" & txt(DateFrom) & "# and JobCloseDate<=#" & txt(DateTo) & "# and CRMemo=1 Order by JobCloseDate,DocID Desc"
End If
Set RsJob = GCn.Execute(GSQL)
If RsJob.RecordCount <= 0 Then
    Lbl(lblIndex) = "No Records for Posting!"
    MsgBox Lbl(lblIndex), vbInformation, "No Records!"
    GoTo lblExit
End If
Lbl(10) = RsJob.RecordCount
Lbl(10).Refresh

mTrans = True
GCn.BeginTrans
GCnFaS.BeginTrans
If SepLabPost Then
    GCnFaW.BeginTrans
End If

'Start Ledger Posting
Do While RsJob.EOF = False
    '**
    Erase LedgAry
    Erase LedgAryLab
    '*********
    xMRPSprTp = 0: xMRPOilTp = 0
    xSprTp = 0: xOilTp = 0
    mShare = 0: mShareAmt = 0: mShare2Amt = 0
    xNetAmt = 0: xRoundAmt = 0: xSprAmtMRPTB = 0: xSprAmtMRPTP = 0
    xOilAmtMRPTB = 0: xOilAmtMRPTP = 0
    xSprAmtTB = 0: xSprAmtTP = 0: xOilAmtTB = 0: xOilAmtTP = 0
    xDisAmtTB = 0: xDisAmtTP = 0: xDisAmtMRPTB = 0: xDisAmtMRPTP = 0
    xGenSurAmt = 0: xTrans = 0: xTaxAmt = 0: xTaxAmtMRP = 0: xPack = 0
    xTurnOver = 0: xReSaleTaxAmt = 0: mQRY = ""
    xNetLabAmt = 0: xLabAmtTB = 0: xLabAmtTP = 0: xLabDisc = 0
    xServTaxAmt = 0: xLabROff = 0
    mSprAmtMRPTB = 0: mSprAmtTB = 0
    mOilAmtMRPTB = 0: mOilAmtTB = 0
    mTotMRPOilTB = 0: mTotOilTB = 0: mTotShareAmt = 0
    mShareSpr = 0: mShareAmtSpr = 0: mShare2AmtSpr = 0
    mTot1ShareAmt = 0: mTot2ShareAmt = 0: mTot3ShareAmt = 0
    '*********
    mCommNarr = ""
    mResult = 0
    mNarr = ""
    i = 0
    j = 0
    '****
   ' If PubVATYN = 1 Then
   '     TaxSQL = "select TF.Tax_Ac_Code,TF.Sur_Ac_Code,sum(Tax_Amt) as TaxAmt,sum(Tax_Sur_Amt+TaxSur_AmtMRP) as TaxSurAmt " & _
            " from SP_Sale left join TaxFormsAc TF on Sp_Sale.Form_Code&'" & PubDivCode & "' =TF.Form_Code&TF.Div_Code "
   ' Else
        TaxSQL = "select TF.Tax_Ac_Code,TF.Sur_Ac_Code,sum(Tax_Amt+Tax_AmtMRP) as TaxAmt,sum(Tax_Sur_Amt+TaxSur_AmtMRP) as TaxSurAmt " & _
            " from SP_Sale left join TaxFormsAc TF on Sp_Sale.Form_Code&'" & PubDivCode & "' =TF.Form_Code&TF.Div_Code "
   'End If
    If mVType = JobSalCashVType Then
        
        SpareDocID = PubDivCode & PubSiteCode & txt(SiteCode).Tag & mVType
        'Arp Start - ENAR
        'LabourDocID = PubDivCode & PubSiteCode & Txt(SiteCode).Tag & mVType
        LabourDocID = PubDivCode & PubSiteCode & txt(SiteCode).Tag & "W_LIC"
        'Arp End
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "inner join TaxFormsAc TF on Sp_Sale.Form_Code&'" & PubDivCode & "' =TF.Form_Code&TF.Div_Code " & _
            "where V_Date=#" & RsJob!JobCloseDate & "# and left(docid,8)='" & left(SpareDocID, 8) & _
            "' Group by TF.PurSal_Ac_Code"
            
        mQRY = "select " & _
            "sum(Total_Amt) as NetAmt,sum(rounded) as RoundAmt," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(SprAmt_MRP_TP) as SprAmtMRPTP, " & _
            "sum(OilAmt_MRP_TB) as OilAmtMRPTB, sum(OilAmt_MRP_TP) as OilAmtMRPTP, " & _
            "sum(SprAmt_TB) as SprAmtTB, sum(SprAmt_TP) as SprAmtTP, " & _
            "sum(OilAmt_TB) as OilAmtTB, sum(OilAmt_TP) as OilAmtTP, " & _
            "sum(D_Amt_TB) as DisAmtTB, sum(D_Amt_TP) as DisAmtTP, " & _
            "sum(D_Amt_MRP_TB) as DisAmtMRPTB, sum(D_Amt_MRP_TP) as DisAmtMRPTP," & _
            "sum(Gen_Sur_Amt) as GenSurAmt,sum(Trans_Amt) as Trans," & _
            "sum(Tax_Amt+Tax_Sur_Amt+Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmt," & _
            "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, sum(TOT_Amt) as TurnOver, " & _
            "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
            "from SP_Sale " & _
            "inner join TaxForms TF on Sp_Sale.Form_Code=TF.Form_Code " & _
            "where V_Date=#" & RsJob!JobCloseDate & "# and left(docid,8)='" & left(SpareDocID, 8) & "'"
        'for tax
        TaxSQL = TaxSQL & " where  V_Date=#" & RsJob!JobCloseDate & "# and left(docid,8)='" & left(SpareDocID, 8) & _
            "' Group by TF.Tax_Ac_Code,TF.Sur_Ac_Code"
        '**Labour
        mLabSQL = "Select sum(LabAmt_TB) as LabAmt_TB,sum(LabAmt_TP) as LabAmt_TP,sum(Lab_D_Amt) as Lab_D_Amt" & _
            ",sum(Lab_TaxAmt) as Lab_TaxAmt,sum(Lab_RoundOff) as Lab_RoundOff,sum(NetLab_Amt) as NetLab_Amt " & _
            "from Job_Card where JobCloseDate=" & ConvertDate(RsJob!JobCloseDate) & _
            " and left(DocId_InvLab,8)='" & left(LabourDocID, 8) & "'"
        '***********
        mNarr = "Workshop Cash Sale (Daily Posting)"
        mCommNarr = mNarr & " [Common]"
        mFADocidSpr = left(SpareDocID, 8) & "YYYYY" & "  " & Format(RsJob!JobCloseDate, "yymmdd")
        mFADocidLab = left(LabourDocID, 8) & "ZZZZZ" & "  " & Format(RsJob!JobCloseDate, "yymmdd")
        PartyCode = PubSprCashAc
        PartyCodeLab = PubSrvCashAc
    Else
        PartyCode = RsJob!DrSpr_AcCode
        PartyCodeLab = RsJob!DrLab_AcCode
        SpareDocID = RsJob!DocId_InvSpr
        mFADocidSpr = SpareDocID
        mFADocidLab = RsJob!DocID_InvLab
       ' SpareDocID = PubDivCode & PubSiteCode & Txt(SiteCode).Tag & mVType
        LabourDocID = RsJob!DocID_InvLab
        
        
        GSQL = "select TF.PurSal_Ac_Code," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(OilAmt_MRP_TB) as OilAmtMRPTB," & _
            "sum(SprAmt_TB) as SprAmtTB, sum(OilAmt_TB) as OilAmtTB " & _
            "from SP_Sale " & _
            "inner join TaxFormsAc TF on Sp_Sale.Form_Code&'" & PubDivCode & "' =TF.Form_Code&TF.Div_Code " & _
            "where docid='" & SpareDocID & _
            "' group by TF.PurSal_Ac_Code"

        mQRY = "select " & _
            "sum(Total_Amt) as NetAmt,sum(rounded) as RoundAmt," & _
            "sum(SprAmt_MRP_TB) as SprAmtMRPTB, sum(SprAmt_MRP_TP) as SprAmtMRPTP, " & _
            "sum(OilAmt_MRP_TB) as OilAmtMRPTB, sum(OilAmt_MRP_TP) as OilAmtMRPTP, " & _
            "sum(SprAmt_TB) as SprAmtTB, sum(SprAmt_TP) as SprAmtTP, " & _
            "sum(OilAmt_TB) as OilAmtTB, sum(OilAmt_TP) as OilAmtTP, " & _
            "sum(D_Amt_TB) as DisAmtTB, sum(D_Amt_TP) as DisAmtTP, " & _
            "sum(D_Amt_MRP_TB) as DisAmtMRPTB, sum(D_Amt_MRP_TP) as DisAmtMRPTP," & _
            "sum(Gen_Sur_Amt) as GenSurAmt,sum(Trans_Amt) as Trans," & _
            "sum(Tax_Amt+Tax_Sur_Amt+Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmt," & _
            "sum(Tax_AmtMRP+TaxSur_AmtMRP) as TaxAmtMRP,sum(Packing) as Pack, sum(TOT_Amt) as TurnOver, " & _
            "sum(ReSalTax_Amt) as ReSaleTaxAmt " & _
            "from SP_Sale " & _
            "inner join TaxForms TF on Sp_Sale.Form_Code=TF.Form_Code " & _
            "where docid='" & SpareDocID & "'"
        'for tax
        TaxSQL = TaxSQL & " where docid='" & SpareDocID & _
            "' Group by TF.Tax_Ac_Code,TF.Sur_Ac_Code"
        '**Labour
        mLabSQL = "Select sum(LabAmt_TB) as LabAmt_TB,sum(LabAmt_TP) as LabAmt_TP,sum(Lab_D_Amt) as Lab_D_Amt" & _
            ",sum(Lab_TaxAmt) as Lab_TaxAmt,sum(Lab_RoundOff) as Lab_RoundOff,sum(NetLab_Amt) as NetLab_Amt " & _
            "from Job_Card where JobCloseDate=" & ConvertDate(RsJob!JobCloseDate) & _
            " and DocId_InvLab='" & LabourDocID & "'"
        '****
'        mNarr = "Works Job No. " & PrinID(RsJob!Job_DocID) & " Cr Spare Bill No. " & PrinID(SpareDocID) & " Dt." & RsJob!JobCloseDate
'        If xNetLabAmt <> 0 Then
'            If RsJob!DocID_InvLab <> "" Then
'                mNarr = mNarr & " Labour Bill No. " & PrinID(RsJob!DocID_InvLab) 'lblLabourBill
'            End If
'        End If
'        mCommNarr = mNarr & " [Common]"
    End If
    
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open mQRY, GCn, adOpenStatic, adLockReadOnly
    
    'for tax purpose
    Set RsTemp = New ADODB.Recordset
    RsTemp.CursorLocation = adUseClient
    RsTemp.Open TaxSQL, GCn, adOpenStatic, adLockReadOnly

'        1.MRP Spr TB = SprAmtMRPTB-OilAmtMRPTB - part of DisAmtMRPTB - part of TaxAmtMRP
'        3.MRP Oil TB = OilAmtMRPTB - part of DisAmtMRPTB - part of TaxAmtMRP

'        2.MRP Spr TP = SprAmtMRPTP-OilAmtMRPTP - part of DisAmtMRPTP
'        4.MRP Oil TP = OilAmtMRPTP - part of DisAmtMRPTP

'        1.Spr TB = SprAmtTB - part of DisAmtTB
'        3.Oil TB = OilAmtTB - part of DisAmtTB
'        2.Spr TP = SprAmtTP - part of DisAmtTP
'        4.Oil TP = OilAmtTP - part of DisAmtTP

    xNetAmt = IIf(IsNull(rsTemp1!NetAmt), 0, rsTemp1!NetAmt): xRoundAmt = IIf(IsNull(rsTemp1!RoundAmt), 0, rsTemp1!RoundAmt)
    xSprAmtMRPTB = IIf(IsNull(rsTemp1!SprAmtMrpTB), 0, rsTemp1!SprAmtMrpTB)
    xOilAmtMRPTB = IIf(IsNull(rsTemp1!OilAmtMrpTB), 0, rsTemp1!OilAmtMrpTB)
    xSprAmtMRPTP = IIf(IsNull(rsTemp1!SprAmtMrpTP), 0, rsTemp1!SprAmtMrpTP)
    xOilAmtMRPTP = IIf(IsNull(rsTemp1!OilAmtMrpTP), 0, rsTemp1!OilAmtMrpTP)
    xSprAmtTB = IIf(IsNull(rsTemp1!SprAmtTB), 0, rsTemp1!SprAmtTB)
    xOilAmtTB = IIf(IsNull(rsTemp1!OilAmtTB), 0, rsTemp1!OilAmtTB)
    xSprAmtTP = IIf(IsNull(rsTemp1!SprAmtTP), 0, rsTemp1!SprAmtTP)
    xOilAmtTP = IIf(IsNull(rsTemp1!OilAmtTP), 0, rsTemp1!OilAmtTP)
    xDisAmtTB = IIf(IsNull(rsTemp1!DisAmtTB), 0, rsTemp1!DisAmtTB)
    xDisAmtTP = IIf(IsNull(rsTemp1!DisAmtTP), 0, rsTemp1!DisAmtTP)
    xDisAmtMRPTB = IIf(IsNull(rsTemp1!DisAmtMRPTB), 0, rsTemp1!DisAmtMRPTB)
    xDisAmtMRPTP = IIf(IsNull(rsTemp1!DisAmtMRPTP), 0, rsTemp1!DisAmtMRPTP)
    xGenSurAmt = IIf(IsNull(rsTemp1!GenSurAmt), 0, rsTemp1!GenSurAmt)
    xTrans = IIf(IsNull(rsTemp1!Trans), 0, rsTemp1!Trans)
    xTaxAmt = IIf(IsNull(rsTemp1!TaxAmt), 0, rsTemp1!TaxAmt)
    xTaxAmtMRP = IIf(IsNull(rsTemp1!TaxAmtMRP), 0, rsTemp1!TaxAmtMRP)
    xPack = IIf(IsNull(rsTemp1!Pack), 0, rsTemp1!Pack)
    xTurnOver = IIf(IsNull(rsTemp1!TurnOver), 0, rsTemp1!TurnOver)
    xReSaleTaxAmt = IIf(IsNull(rsTemp1!ReSaleTaxAmt), 0, rsTemp1!ReSaleTaxAmt)
    '*** Sale Amount Row
    i = 1
    ReDim Preserve LedgAry(1)
    '**Taxable Spr / Oil Calculation
     Do While GRs.EOF = False
        mOilAmtMRPTB = IIf(IsNull(GRs!OilAmtMrpTB), 0, GRs!OilAmtMrpTB)
        mSprAmtMRPTB = IIf(IsNull(GRs!SprAmtMrpTB), 0, GRs!SprAmtMrpTB) ' - mOilAmtMRPTB
        mSprAmtTB = IIf(IsNull(GRs!SprAmtTB), 0, GRs!SprAmtTB)
        mOilAmtTB = IIf(IsNull(GRs!OilAmtTB), 0, GRs!OilAmtTB)
        'Allocate values in their proportions
        If (mSprAmtMRPTB + mOilAmtMRPTB) <> 0 Then
            mShare = Round((mSprAmtMRPTB + mOilAmtMRPTB) * 100 / (xSprAmtMRPTB + xOilAmtMRPTB), 2)
            mShareAmt = Round(xDisAmtMRPTB * mShare / 100, 2)
            mShare2Amt = Round(xTaxAmtMRP * mShare / 100, 2)
            mShareSpr = Round(mSprAmtMRPTB * 100 / (mSprAmtMRPTB + mOilAmtMRPTB), 2)
            mShareAmtSpr = Round(mShareAmt * mShareSpr / 100, 2)
            mShare2AmtSpr = Round(mShare2Amt * mShareSpr / 100, 2)
            mTot1ShareAmt = mTot1ShareAmt + mShareAmt
            mTot2ShareAmt = mTot2ShareAmt + mShare2Amt
            If GRs.AbsolutePosition = GRs.RecordCount Then
                mShareAmt = mShareAmt + ((xDisAmtMRPTB) - mTot1ShareAmt)
                mShare2Amt = mShare2Amt + ((xTaxAmtMRP) - mTot2ShareAmt)
            End If
   '        If PubVATYN = 1 Then
   '          mSprAmtMRPTB = mSprAmtMRPTB '- (mShareAmtSpr + mShare2AmtSpr)
   '        Else
             mSprAmtMRPTB = mSprAmtMRPTB - (mShareAmtSpr + mShare2AmtSpr)
   '        End If
            mOilAmtMRPTB = mOilAmtMRPTB - ((mShareAmt + mShare2Amt) - (mShareAmtSpr + mShare2AmtSpr))
        End If
        '*****
        If (mSprAmtTB + mOilAmtTB) <> 0 Then
            mShare = Round((mSprAmtTB + mOilAmtTB) * 100 / (xSprAmtTB + xOilAmtTB), 2)
            mShareAmt = Round((xDisAmtTB - xDisAmtMRPTB) * mShare / 100, 2)
            mShareSpr = Round(mSprAmtTB * 100 / (mSprAmtTB + mOilAmtTB), 2)
            mShareAmtSpr = Round(mShareAmt * mShareSpr / 100, 2)
            mTot3ShareAmt = mTot3ShareAmt + mShareAmt
            If GRs.AbsolutePosition = GRs.RecordCount Then
                mShareAmt = mShareAmt + ((xDisAmtTB - xDisAmtMRPTB) - mTot3ShareAmt)
            End If
            mSprAmtTB = mSprAmtTB - (mShareAmtSpr)
            mOilAmtTB = mOilAmtTB - (mShareAmt - mShareAmtSpr)
        End If
        'Spare Sale A/c Taxable
        mTotMRPOilTB = mTotMRPOilTB + mOilAmtMRPTB
        mTotOilTB = mTotOilTB + mOilAmtTB
        '*****
        If mSprAmtMRPTB + mSprAmtTB <> 0 Then
            i = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(i)
            LedgAry(i).SubCode = GRs!PurSal_Ac_Code
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(mSprAmtMRPTB + mSprAmtTB, 2)
            LedgAry(i).Narration = mNarr ' & " Spare"
        End If
        GRs.MoveNext
    Loop
    If (xSprAmtMRPTP + xOilAmtMRPTP) <> 0 Then
        xMRPSprTp = xSprAmtMRPTP '- xOilAmtMRPTP
        xMRPOilTp = xOilAmtMRPTP
        mShare = Round(xMRPSprTp * 100 / (xSprAmtMRPTP + xOilAmtMRPTP), 2)
        mShareAmt = Round(xDisAmtMRPTP * mShare / 100, 2)
        xMRPSprTp = xMRPSprTp - (mShareAmt)
        xMRPOilTp = xMRPOilTp - (xDisAmtMRPTP - (mShareAmt))
    End If
    If (xSprAmtTP + xOilAmtTP) <> 0 Then
        mShare = Round(xSprAmtTP * 100 / (xSprAmtTP + xOilAmtTP), 2)
        mShareAmt = Round((xDisAmtTP - xDisAmtMRPTP) * mShare / 100, 2)
        xSprTp = xSprAmtTP - (mShareAmt)
        xOilTp = xOilAmtTP - ((xDisAmtTP - xDisAmtMRPTP) - (mShareAmt))
    End If
    '*Labour
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open mLabSQL, GCn, adOpenStatic, adLockReadOnly
    xNetLabAmt = 0: xLabAmtTB = 0: xLabAmtTP = 0: xLabDisc = 0
    xServTaxAmt = 0: xLabROff = 0
    If rsTemp1.RecordCount > 0 Then
        xNetLabAmt = IIf(IsNull(rsTemp1!NetLab_Amt), 0, rsTemp1!NetLab_Amt)
        xLabAmtTB = IIf(IsNull(rsTemp1!LabAmt_TB), 0, rsTemp1!LabAmt_TB)
        xLabAmtTP = IIf(IsNull(rsTemp1!LabAmt_TP), 0, rsTemp1!LabAmt_TP)
        xLabDisc = IIf(IsNull(rsTemp1!Lab_D_Amt), 0, rsTemp1!Lab_D_Amt)
        xServTaxAmt = IIf(IsNull(rsTemp1!Lab_TaxAmt), 0, rsTemp1!Lab_TaxAmt)
        xLabROff = IIf(IsNull(rsTemp1!Lab_RoundOff), 0, rsTemp1!Lab_RoundOff)
    End If
    Set rsTemp1 = Nothing
    If (xLabAmtTB + xLabAmtTP) <> 0 Then
        mShare = Round(xLabAmtTB * 100 / (xLabAmtTB + xLabAmtTP), 2)
        mShareAmt = Round(xLabDisc * mShare / 100, 2)
        xLabAmtTB = xLabAmtTB - (mShareAmt)
        xLabAmtTP = xLabAmtTP - (xLabDisc - mShareAmt)
    End If
    '**
'0.Party A/c or Cash A/c
'1.Taxable Spr = MRP Spr TB + SPR TB
'2.Taxpaid Spr = MRP Spr TP + SPR TP
'3.Taxable Oil = MRP Oil TB + Oil TB
'4.Taxable Oil = MRP Oil TP + Oil TP
'5.xGenSurAmt
'6.xPack
'7.xTurnOver
'8.xReSaleTaxAmt
    '*******
    'Sale Party A/c
    'I = 0
    If mVType <> JobSalCashVType Then
        mNarr = "Works Job No. " & PrinID(RsJob!DocID) & " Cr Spare Bill No. " & PrinID(SpareDocID) '& " Dt." & Txt(JobCDt) & " Rs." & Format(xNetAmt, "0.00")
        If xNetLabAmt <> 0 Then
            If RsJob!DocID_InvLab <> "" Then
                mNarr = mNarr & " Labour Bill No. " & PrinID(RsJob!DocID_InvLab) '& " Rs." & Format(xNetLabAmt, "0.00")
            End If
        End If
        mCommNarr = mNarr & " [Common]"
    End If

    LedgAry(0).SubCode = PartyCode
    
    LedgAry(0).AmtDr = IIf(SepLabPost, xNetAmt, xNetAmt + xNetLabAmt)
    LedgAry(0).AmtCr = 0
    LedgAry(0).Narration = mNarr
    'Spare Sale A/c Taxpaid
    If xMRPSprTp + xSprTp <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!SprSalTP_Ac
        LedgAry(i).AmtDr = 0
        LedgAry(i).AmtCr = Round(xMRPSprTp + xSprTp, 2)
        LedgAry(i).Narration = mNarr ' & " Spare"
    End If
    'Oil Sale A/c Taxable
    If mTotMRPOilTB + mTotOilTB <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!OilSalTB_Ac
        LedgAry(i).AmtDr = 0
        LedgAry(i).AmtCr = Round(mTotMRPOilTB + mTotOilTB, 2)
        LedgAry(i).Narration = mNarr ' & " Spare"
    End If
     'Oil Sale A/c Taxpaid
     If xMRPOilTp + xOilTp <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!OilSalTP_Ac
        LedgAry(i).AmtDr = 0
        LedgAry(i).AmtCr = Round(xMRPOilTp + xOilTp, 2)
        LedgAry(i).Narration = mNarr ' & " Spare"
     End If
      'GenSurAmt
     If xGenSurAmt <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!SprGenSur_Ac
        LedgAry(i).AmtDr = 0
        LedgAry(i).AmtCr = Round(xGenSurAmt, 2)
        LedgAry(i).Narration = mNarr ' & " Sale Tax"
     End If
    'Transportation
     If xTrans <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!Transportation_Ac
        If xTrans > 0 Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(xTrans, 2)
        Else
            LedgAry(i).AmtDr = Round(Abs(xTrans), 2)
           LedgAry(i).AmtCr = 0
        End If
        LedgAry(i).Narration = mNarr '& " Transportation"
     End If
     If RsTemp.RecordCount > 0 Then
         Do While RsTemp.EOF = False
            If RsTemp!TaxAmt <> 0 Then
                i = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(i)
                LedgAry(i).SubCode = RsTemp!Tax_Ac_Code
                If RsTemp!TaxAmt > 0 Then
                    LedgAry(i).AmtDr = 0
                    LedgAry(i).AmtCr = Round(RsTemp!TaxAmt, 2)
                Else
                    LedgAry(i).AmtDr = Round(Abs(RsTemp!TaxAmt), 2)
                    LedgAry(i).AmtCr = 0
                End If
                LedgAry(i).Narration = mNarr '& " Sales Tax & Surcharge"
            End If
            If RsTemp!TaxSurAmt <> 0 Then
                i = UBound(LedgAry) + 1
                ReDim Preserve LedgAry(i)
                LedgAry(i).SubCode = RsTemp!Sur_Ac_Code
                If RsTemp!TaxSurAmt > 0 Then
                    LedgAry(i).AmtDr = 0
                    LedgAry(i).AmtCr = Round(RsTemp!TaxSurAmt, 2)
                Else
                    LedgAry(i).AmtDr = Round(Abs(RsTemp!TaxSurAmt), 2)
                    LedgAry(i).AmtCr = 0
                End If
                 LedgAry(i).Narration = mNarr '& " Sales Tax & Surcharge"
             End If
             RsTemp.MoveNext
         Loop
     End If
    'Misc / Packing Chrg
    If xPack <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!MiscChrg_Ac
        If Val(xPack) > 0 Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(xPack, 2)
        Else
            LedgAry(i).AmtDr = Round(Abs(xPack), 2)
            LedgAry(i).AmtCr = 0
        End If
        LedgAry(i).Narration = mNarr '& " Misc Charges"
    End If
    If xTurnOver <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!TOTax_Ac
        If xTurnOver > 0 Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(xTurnOver, 2)
        Else
            LedgAry(i).AmtDr = Round(Abs(xTurnOver), 2)
            LedgAry(i).AmtCr = 0
        End If
        LedgAry(i).Narration = mNarr '& " TurnOver Amt"
    End If
    If xReSaleTaxAmt <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!ReSaleTax_Ac
        If xReSaleTaxAmt > 0 Then
            LedgAry(i).AmtDr = 0
            LedgAry(i).AmtCr = Round(xReSaleTaxAmt, 2)
        Else
            LedgAry(i).AmtDr = Round(Abs(xReSaleTaxAmt), 2)
            LedgAry(i).AmtCr = 0
        End If
        LedgAry(i).Narration = mNarr '& " ReSale Tax Amount"
    End If
    '********
    If SepLabPost Then  'Separate Posting for Spr & Labour
        If xRoundAmt <> 0 Then
            i = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(i)
            LedgAry(i).SubCode = rsCtrlAc!SprROff_Ac
            If xRoundAmt > 0 Then
                LedgAry(i).AmtDr = 0
                LedgAry(i).AmtCr = Round(xRoundAmt, 2)
            Else
                LedgAry(i).AmtDr = Round(Abs(xRoundAmt), 2)
                LedgAry(i).AmtCr = 0
            End If
            LedgAry(i).Narration = mNarr '& " Round Off"
        End If
        'Second LedgAryLab
        'Labour Amt
        If xNetLabAmt <> 0 Then
            ReDim Preserve LedgAryLab(0)
            i = 0
            LedgAryLab(i).SubCode = PartyCodeLab
            LedgAryLab(i).AmtDr = xNetLabAmt
            LedgAryLab(i).Narration = mNarr & "Labour charges"
            If xLabAmtTB <> 0 Then
                i = UBound(LedgAryLab) + 1
                ReDim Preserve LedgAryLab(i)
                LedgAryLab(i).SubCode = rsCtrlAcLab!SrvLabourTB_Ac    'Labour A/c Code
                LedgAryLab(i).AmtCr = xLabAmtTB
                LedgAryLab(i).Narration = mNarr & "Labour charges"
            End If
            If xLabAmtTP <> 0 Then
                i = UBound(LedgAryLab) + 1
                ReDim Preserve LedgAryLab(i)
                LedgAryLab(i).SubCode = rsCtrlAcLab!SrvLabour_Ac    'Labour A/c Code
                LedgAryLab(i).AmtCr = xLabAmtTP
                LedgAryLab(i).Narration = mNarr & "Labour charges"
            End If
            'Service Tax
            If xServTaxAmt <> 0 Then
                i = UBound(LedgAryLab) + 1
                ReDim Preserve LedgAryLab(i)
                LedgAryLab(i).SubCode = rsCtrlAcLab!SrvTax_Ac    'Service Tax A/c Code
                LedgAryLab(i).AmtCr = xServTaxAmt
                LedgAryLab(i).Narration = mNarr & " Service Tax on Labour charges"
            End If
            'Labour Round Off
            If xLabROff <> 0 Then
                i = UBound(LedgAryLab) + 1
                ReDim Preserve LedgAryLab(i)
                LedgAryLab(i).SubCode = rsCtrlAcLab!SrvROff_Ac
                If xLabROff > 0 Then
                    LedgAryLab(i).AmtCr = xLabROff
                Else
                    LedgAryLab(i).AmtDr = Abs(xLabROff)
                End If
                LedgAryLab(i).Narration = mNarr & " Labour Round Diff."
            End If
        End If
    Else    'Combined Posting for Spr & Labour
        'Net Posting Amt = Spr + Labour Amt
        'Labour Taxable
        If xLabAmtTB <> 0 Then
            i = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(i)
            LedgAry(i).SubCode = rsCtrlAcLab!SrvLabourTB_Ac    'Taxable Labour A/c Code
            LedgAry(i).AmtCr = xLabAmtTB
            LedgAry(i).Narration = mNarr & "Labour charges"
        End If
        'Labour Taxpaid
        If xLabAmtTP <> 0 Then
            i = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(i)
            LedgAry(i).SubCode = rsCtrlAcLab!SrvLabour_Ac    'Taxpaid Labour A/c Code
            LedgAry(i).AmtCr = xLabAmtTP
            LedgAry(i).Narration = mNarr & "Labour charges"
        End If
        'Service Tax
        If xServTaxAmt <> 0 Then
            i = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(i)
            LedgAry(i).SubCode = rsCtrlAcLab!SrvTax_Ac    'Service Tax A/c Code
            LedgAry(i).AmtCr = xServTaxAmt
            LedgAry(i).Narration = mNarr & " Service Tax on Labour charges"
        End If
        'Round Off = Spare Round Off + Labour round Off
        If xRoundAmt + xLabROff <> 0 Then
            i = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(i)
            LedgAry(i).SubCode = rsCtrlAc!SprROff_Ac
            
            If xRoundAmt + xLabROff > 0 Then
                LedgAry(i).AmtCr = xRoundAmt + xLabROff
            Else
                LedgAry(i).AmtDr = Abs(xRoundAmt + xLabROff)
            End If
            LedgAry(i).Narration = mNarr & " Round Diff. Spare+Labour"
        End If
    End If
    mResult = LedgerPost("A", LedgAry, GCnFaS, mFADocidSpr, CDate(RsJob!JobCloseDate), mCommNarr)
    If mResult <> 1 Then MsgBox "Error in Ledger Posting : " & mFADocidSpr, vbOKOnly, "Validation"
    If SepLabPost Then
        mResult = LedgerPost("A", LedgAryLab, GCnFaW, mFADocidLab, CDate(RsJob!JobCloseDate), mCommNarr)
        If mResult <> 1 Then MsgBox "Error in Ledger Posting : " & mFADocidLab, vbOKOnly, "Validation"
    End If
    '**
    Lbl(12) = RsJob.AbsolutePosition
    Lbl(12).Refresh

    RsJob.MoveNext
Loop
GCn.CommitTrans
GCnFaS.CommitTrans
If SepLabPost Then
    GCnFaW.CommitTrans
End If
mTrans = False
Lbl(lblIndex) = "Posting completed sucessfully !"
Lbl(lblIndex).Refresh

lblExit:
    Set GRs = Nothing
    Set RsTemp = Nothing
    Set rsTemp1 = Nothing
    If mTrans Then
        GCn.RollbackTrans
        GCnFaS.RollbackTrans
        If SepLabPost Then
            GCnFaW.RollbackTrans
        End If
    End If
    If err.NUMBER <> 0 Then MsgBox err.Description, vbCritical, "Ledger Posting Failed!'"

End Sub

Private Sub PostVehPur(TrnType$, mVType$, TitleStr$, lblIndex As Integer)
Dim xNetAmt As Double, xEntryTaxAmt As Double
'A/c Posting related declarations
Dim mResult As Byte, mCommNarr$, mNarr$, TaxSQL$, i As Integer, j As Integer
Dim mSprPurPfx$, mFADocID$, mTxtDocID$, PartyAcCode$
Dim rsVehPurch As ADODB.Recordset, mTrans As Boolean
Dim rsVehStock As ADODB.Recordset

lblRefresh
Lbl(1) = TitleStr
Lbl(1).Refresh

Lbl(lblIndex) = "Posting in progress..."
Lbl(lblIndex).Visible = True
Lbl(lblIndex).Refresh

GSQL = "Select Distinct V_Date,DocID,PartyCode,PBILL_NO,PBILL_DATE,Tot_Amount,Tax_Amt, " & _
    "TF.PurSal_Ac_Code,Veh_Purch1.U_AE, TaxForms.L_C, TF.Tax_Ac_Code " & _
    " from (Veh_Purch1 " & _
    " left join TaxFormsAc as TF on Veh_Purch1.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code) " & _
    " left join TaxForms on Veh_Purch1.Form_Code=TaxForms.Form_Code " & _
    " where left(DocID,1)='" & PubDivCode & "' and trim(" & cMID("DocID", "4", "5") & ")='" & mVType & "' and V_Date >=#" & txt(DateFrom) & "# and V_Date<=#" & txt(DateTo) & _
    " # Order By V_Date,DocID"
    
Set rsVehPurch = GCn.Execute(GSQL)
If rsVehPurch.RecordCount <= 0 Then
    Lbl(lblIndex) = "No Purchase Records for Posting!"
    MsgBox Lbl(lblIndex), vbInformation, "No Records!"
    GoTo lblExit
End If
Lbl(10) = rsVehPurch.RecordCount
Lbl(10).Refresh

GSQL = "Select Pur_DocId,ChassisNo from Veh_Stock " & _
    " where left(Pur_DocID,1)='" & PubDivCode & "' and trim(" & cMID("Pur_DocID", "4", "5") & ")='" & mVType & _
    "' and Pur_VDate >=#" & txt(DateFrom) & "# and Pur_VDate<=#" & txt(DateTo) & _
    "#  Order By Pur_DocID,ChassisNo"
Set rsVehStock = GCn.Execute(GSQL)
If rsVehStock.RecordCount <= 0 Then
    MsgBox "Discrepency in Chassis Records!,Posting Aborted", vbInformation, "No Records in Veh_Stock!"
    GoTo lblExit
End If

mTrans = True
GCn.BeginTrans
GCnFaV.BeginTrans
'Start Ledger Posting
Do While rsVehPurch.EOF = False
    '**
    Dim LedgAry(2) As LedgRec
    mResult = 0
    TaxSQL = ""
    i = 0
    j = 0
    xNetAmt = 0
    xEntryTaxAmt = 0
    '****
    Lbl(11) = rsVehPurch!V_DATE
    Lbl(11).Refresh
    
    mNarr = "Through Vehicle Purchase Mfg Bill No." & rsVehPurch!PBILL_NO & " Date " & rsVehPurch!PBILL_DATE & " Chassis No."
    'Create Narration with Chassis No's
    If rsVehStock.RecordCount > 0 Then
        rsVehStock.MoveFirst
        rsVehStock.FIND ("Pur_DocID='" & rsVehPurch!DocID & "'")
    End If
    If rsVehStock.EOF = False Then
        Do While rsVehStock!Pur_DocId = rsVehPurch!DocID
            If rsVehStock!ChassisNo <> "" Then
                mNarr = mNarr & rsVehStock!ChassisNo & "."
            End If
            rsVehStock.MoveNext
            If rsVehStock.EOF Then
                Exit Do
            ElseIf rsVehStock!Pur_DocId = rsVehPurch!DocID Then
                Exit Do
            End If
        Loop
    End If
    mCommNarr = mNarr
    i = 0
    If rsVehPurch!Tot_Amount <> 0 Then
        If PubVATYN = 1 And rsVehPurch!L_C = "Local" Then
            'Purchase A/c
            LedgAry(i).SubCode = rsVehPurch!PurSal_Ac_Code
            LedgAry(i).AmtDr = VNull(rsVehPurch!Tot_Amount) - VNull(rsVehPurch!Tax_Amt)
            LedgAry(i).Narration = mNarr
            LedgAry(i).ContraSub = rsVehPurch!PartyCode
            i = i + 1
        
            'Tax A/c
            LedgAry(i).SubCode = rsVehPurch!Tax_Ac_Code
            LedgAry(i).AmtDr = rsVehPurch!Tax_Amt
            LedgAry(i).Narration = mNarr
            LedgAry(i).ContraSub = rsVehPurch!PartyCode
            i = i + 1
        
        Else
            'Purchase A/c
            LedgAry(i).SubCode = rsVehPurch!PurSal_Ac_Code
            LedgAry(i).AmtDr = rsVehPurch!Tot_Amount
            LedgAry(i).Narration = mNarr
            LedgAry(i).ContraSub = rsVehPurch!PartyCode
            i = i + 1
        End If
        'Party A/c
        LedgAry(i).SubCode = rsVehPurch!PartyCode
        LedgAry(i).AmtCr = rsVehPurch!Tot_Amount
        LedgAry(i).Narration = mNarr
        LedgAry(i).ContraSub = rsVehPurch!PurSal_Ac_Code
    End If
    mResult = LedgerPost(rsVehPurch!U_AE, LedgAry, GCnFaV, rsVehPurch!DocID, CDate(rsVehPurch!V_DATE), mCommNarr)
    Erase LedgAry
    If mResult <> 1 Then
        MsgBox "Error in Ledger Posting : " & rsVehPurch!DocID, vbOKOnly, "Validation"
    End If
    '**
    Lbl(12) = rsVehPurch.AbsolutePosition
    Lbl(12).Refresh
    rsVehPurch.MoveNext
Loop
GCn.CommitTrans
GCnFaV.CommitTrans
mTrans = False
Lbl(lblIndex) = "Posting completed sucessfully !"
Lbl(lblIndex).Refresh
'MsgBox TitleStr & vbCrLf & "Posting completed sucessfully !", vbInformation, "Ledger Posting"

lblExit:
    Set rsVehPurch = Nothing
    Set rsVehStock = Nothing
    Set GRs = Nothing
    If mTrans Then
        GCn.RollbackTrans
        GCnFaV.RollbackTrans
    End If
    If err.NUMBER <> 0 Then
        MsgBox err.Description & vbCrLf & "Ledger Posting Terminated!", vbCritical
'        ProcAcPost = True
    End If
End Sub

Private Sub PostVehSal(TrnType$, mVType$, TitleStr$, lblIndex As Integer)
Dim xNetAmt As Double, xEntryTaxAmt As Double
'A/c Posting related declarations
Dim mBookDocID$
Dim mResult As Byte, mCommNarr$, mNarr$, TaxSQL$, i As Integer, j As Integer
Dim mSprPurPfx$, mFADocID$, mTxtDocID$, PartyAcCode$
Dim rsVehSal As ADODB.Recordset, mTrans As Boolean
Dim SubTotA As Double, mPostFinAmt As Byte, mGTotAmt As Double, mTOT_Ac_Code$

lblRefresh
Lbl(1) = TitleStr
Lbl(1).Refresh

Lbl(lblIndex) = "Posting in progress..."
Lbl(lblIndex).Visible = True
Lbl(lblIndex).Refresh

If IsNull(rsCtrlAc!Fitment_Ac) Or rsCtrlAc!Fitment_Ac = "" Or _
    IsNull(rsCtrlAc!Fuel_Ac) Or rsCtrlAc!Fuel_Ac = "" Or _
    IsNull(rsCtrlAc!VehROff_Ac) Or rsCtrlAc!VehROff_Ac = "" Then
    MsgBox "Please define Fitment,Fuel and Round Off A/c's in Vehicle A/c Controls" & vbCrLf & "A/c Posting Aborted !", vbInformation, "Controls Not Filled!"
    GoTo lblExit
End If

mPostFinAmt = GCn.Execute("select " & vIsNull("PostFinAmt", "0") & " as PostFinAmt from Syctrl").Fields(0).Value

'**TurnOver Tax A/c
GSQL = "Select sum(TOT_Amt) from Veh_Order " & _
    " where left(Inv_DocID,1)='" & PubDivCode & "' and trim(" & cMID("Inv_DocID", "4", "5") & ")='" & mVType & _
    "' and Inv_Date >=#" & txt(DateFrom) & "# and Inv_Date<=#" & txt(DateTo) & "#"
If GCn.Execute(GSQL).Fields(0).Value > 0 Then
    mTOT_Ac_Code = G_FaCn.Execute("select " & xIsNull("TOTax_Ac", "") & " as TOT_Ac from AcControls where Div_Code='" & PubDivCode & "'").Fields(0).Value
    If mTOT_Ac_Code = "" Then
        MsgBox "Please define Turn Over Tax A/c in System Controls" & vbCrLf & "Posting Aborted!", vbInformation, "Controls Not Filled!"
        GoTo lblExit
    End If
End If
'***
GSQL = "Select OrdDocId,Inv_Date,Inv_Prefix,Inv_DocId,PartyCode,VRate,Margine,Rebate,InciChrg,Octroi,RegTemp,TransitInsu,MVT,Transport,OtherChrg,DieselAmt,TOT_Amt,Fin_Amt,Fund_Source,Net_Amount," & _
    " Fit_Amt,Tax_Amt,Surcharge_Amt,Fit_Tax,Tax_Amt,DieselAmt,Round_off,Chassis,TF.Tax_Ac_Code,TF.Sur_Ac_Code," & _
    " switch(CF.Ac_YN='1','Y',CF.Ac_YN<>'1','N') as ACYN,CF.AcCode as FinACCode,TF.PurSal_Ac_Code,Veh_Order.Inv_UAE " & _
    " from (Veh_Order " & _
    " left join TaxFormsAc as TF on Veh_Order.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code) " & _
    " Left Join ContractFinance as CF on Veh_Order.FB_CODE=CF.FinCode " & _
    " where left(Inv_DocID,1)='" & PubDivCode & "' and trim(" & cMID("Inv_DocID", "4", "5") & ")='" & mVType & _
    "' and Inv_Date >=#" & txt(DateFrom) & "# and Inv_Date<=#" & txt(DateTo) & _
    "# Order By Inv_Date,Inv_DocID"

Set rsVehSal = GCn.Execute(GSQL)
If rsVehSal.RecordCount <= 0 Then
    Lbl(lblIndex) = "No Sale Records for Posting!"
    MsgBox Lbl(lblIndex), vbInformation, "No Records!"
    GoTo lblExit
End If
Lbl(10) = rsVehSal.RecordCount
Lbl(10).Refresh

mTrans = True
GCn.BeginTrans
GCnFaV.BeginTrans
'Start Ledger Posting
Do While rsVehSal.EOF = False
    '**
    'A/c Posting related declarations
    Dim LedgAry(7) As LedgRec
    mResult = 0
    mNarr = ""
    TaxSQL = ""
    i = 0
    j = 0
    xNetAmt = 0
    xEntryTaxAmt = 0
    '****
    Lbl(11) = rsVehSal!Inv_Date
    Lbl(11).Refresh
    '***
    If mPostFinAmt = 1 And rsVehSal!Fin_Amt <> 0 Then
        If (rsVehSal!Fund_Source = 0 Or rsVehSal!Fund_Source = 1) And rsVehSal!AcYN = "Y" Then
            If rsVehSal!FinAcCode = "" Or IsNull(rsVehSal!FinAcCode) Then
                MsgBox "Please Define A/c Code in Financier Master" & vbCrLf & "A/c Posting Aborted !", vbInformation, "No Records in Veh_Stock!"
                GoTo lblExit
            End If
        End If
    End If
    '***
    mBookDocID = PrinID(rsVehSal!OrdDocId)
    mNarr = "By Sales Invoice No." & PrinID(rsVehSal!Inv_DocId) & " Dt. " & rsVehSal!Inv_Date & " Chassis " & rsVehSal!Chassis
    mCommNarr = mNarr
    i = 0
    LedgAry(i).SubCode = rsVehSal!PartyCode
    mGTotAmt = rsVehSal!Net_Amount
    If mPostFinAmt = 0 Then
        mGTotAmt = rsVehSal!Net_Amount + rsVehSal!Fin_Amt
    End If
    LedgAry(i).AmtDr = Round(rsVehSal!Net_Amount, 2)
    LedgAry(i).Narration = mNarr
    'Vehicle Sale A/c
    SubTotA = rsVehSal!vrate + rsVehSal!Margine - rsVehSal!Rebate + rsVehSal!InciChrg
    SubTotA = SubTotA + rsVehSal!Octroi + rsVehSal!RegTemp + rsVehSal!TransitInsu + rsVehSal!MVT + rsVehSal!Transport + rsVehSal!OtherChrg
    SubTotA = Round(SubTotA - rsVehSal!DieselAmt, 0)
    If SubTotA <> 0 Then
        i = i + 1
        LedgAry(i).SubCode = rsVehSal!PurSal_Ac_Code
        LedgAry(i).AmtCr = Round(SubTotA, 2)
        LedgAry(i).Narration = mNarr
    End If
    'Fitment Amount
    If rsVehSal!Fit_Amt <> 0 Then
        i = i + 1
        LedgAry(i).SubCode = rsCtrlAc!Fitment_Ac
        LedgAry(i).AmtCr = Round(rsVehSal!Fit_Amt, 2)
        LedgAry(i).Narration = mNarr & " Additional Fitments on Vehicle Sale Bill"
    End If
    'Tax Amt
    If rsVehSal!Tax_Amt + rsVehSal!Surcharge_Amt + rsVehSal!Fit_Tax <> 0 Then
        If rsVehSal!Tax_Ac_Code <> "" And rsVehSal!Sur_Ac_Code <> "" _
             And rsVehSal!Tax_Ac_Code <> rsVehSal!Sur_Ac_Code Then
            If rsVehSal!Tax_Amt <> 0 Then
                i = i + 1
                LedgAry(i).SubCode = rsVehSal!Tax_Ac_Code
                LedgAry(i).AmtCr = Round(rsVehSal!Tax_Amt + rsVehSal!Fit_Tax, 2)
                LedgAry(i).Narration = mNarr & " Sale Tax"
            End If
            If rsVehSal!Surcharge_Amt <> 0 Then
                i = i + 1
                LedgAry(i).SubCode = rsVehSal!Sur_Ac_Code
                LedgAry(i).AmtCr = Round(rsVehSal!Surcharge_Amt, 2)
                LedgAry(i).Narration = mNarr & " Surcharge on Sales Tax"
            End If
        Else
            i = i + 1
            LedgAry(i).SubCode = rsVehSal!Tax_Ac_Code
            LedgAry(i).AmtCr = Round(rsVehSal!Tax_Amt + rsVehSal!Surcharge_Amt + rsVehSal!Fit_Tax, 2)
            LedgAry(i).Narration = mNarr & " Sales Tax & Surcharge"
        End If
    End If
    If rsVehSal!Tot_Amt <> 0 Then
        i = i + 1
        LedgAry(i).SubCode = mTOT_Ac_Code
        LedgAry(i).AmtCr = rsVehSal!Tot_Amt
        LedgAry(i).Narration = mNarr & " TOT Amt"
    End If
    If rsVehSal!Round_off <> 0 Then
        i = i + 1
        LedgAry(i).SubCode = rsCtrlAc!VehROff_Ac
        If rsVehSal!Round_off > 0 Then
            LedgAry(i).AmtCr = Round(rsVehSal!Round_off, 2)
        Else
            LedgAry(i).AmtDr = Round(Abs(rsVehSal!Round_off), 2)
        End If
        LedgAry(i).Narration = mNarr & " Round Off"
    End If
    'Fuel Amount
    If rsVehSal!DieselAmt <> 0 Then
        i = i + 1
        LedgAry(i).SubCode = rsCtrlAc!Fuel_Ac
        LedgAry(i).AmtDr = Round(rsVehSal!DieselAmt, 2)
        LedgAry(i).Narration = mNarr & " Fuel Amount"
    End If
    If mPostFinAmt = 1 And rsVehSal!Fin_Amt <> 0 Then
        If (rsVehSal!Fund_Source = 0 Or rsVehSal!Fund_Source = 1) And rsVehSal!AcYN = "Y" Then
            If rsVehSal!AcCode = "" Or IsNull(rsVehSal!AcCode) Then
            Else
                i = i + 1
                LedgAry(i).SubCode = rsVehSal!FinAcCode
                LedgAry(i).AmtDr = Round(rsVehSal!Fin_Amt, 2)
                LedgAry(i).Narration = mNarr & " Finance Amount."
                i = i + 1
                LedgAry(i).SubCode = rsVehSal!PartyCode
                LedgAry(i).AmtCr = Round(rsVehSal!Fin_Amt, 2)
                LedgAry(i).Narration = mNarr & " Finance Amount."
            End If
        End If
    End If
    mResult = LedgerPost("A", LedgAry, GCnFaV, rsVehSal!Inv_DocId, CDate(rsVehSal!Inv_Date), mCommNarr)
    Erase LedgAry
    If mResult <> 1 Then
         MsgBox "Error in Ledger Posting : " & rsVehSal!Inv_DocId, vbOKOnly, "Validation"
    End If
    '**
    Lbl(12) = rsVehSal.AbsolutePosition
    Lbl(12).Refresh
    rsVehSal.MoveNext
Loop
GCn.CommitTrans
GCnFaV.CommitTrans
mTrans = False
Lbl(lblIndex) = "Posting completed sucessfully !"
Lbl(lblIndex).Refresh
'MsgBox TitleStr & vbCrLf & "Posting completed sucessfully !", vbInformation, "Ledger Posting"

lblExit:
    Set rsVehSal = Nothing
    Set GRs = Nothing
    If mTrans Then
        GCn.RollbackTrans
        GCnFaV.RollbackTrans
    End If
    If err.NUMBER <> 0 Then
        MsgBox err.Description & vbCrLf & "Ledger Posting Terminated!", vbCritical
'        ProcAcPost = True
    End If
End Sub

Private Sub PostSprTrfIssue(TrnType$, mVType$, TitleStr$, lblIndex As Integer)
On Error GoTo lblExit
'Dim xMRPSprTp As Double, xMRPOilTp As Double
'Dim xSprTp As Double, xOilTp As Double
'Dim mShare As Single, mShareAmt As Double, mShare2Amt As Double
'Dim xNetAmt As Double, xRoundAmt As Double, xSprAmtMRPTB As Double, xSprAmtMRPTP As Double
'Dim xOilAmtMRPTB As Double, xOilAmtMRPTP As Double
'Dim xSprAmtTB  As Double, xSprAmtTP As Double, xOilAmtTB As Double, xOilAmtTP As Double
'Dim xDisAmtTB As Double, xDisAmtTP As Double, xDisAmtMRPTB As Double, xDisAmtMRPTP As Double
'Dim xGenSurAmt As Double, xTrans As Double, xTaxAmt As Double, xTaxAmtMRP As Double, xPack As Double
Dim xTurnOver As Double, xReSaleTaxAmt As Double, mFADocID$, mQRY$, PartyCode$
Dim RsTemp As ADODB.Recordset, rsTemp1 As ADODB.Recordset
'A/c Posting related declarations
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, mNarr$, TaxSQL$, i As Integer, j As Integer
'Dim mSprAmtMRPTB As Double, mSprAmtTB As Double
'Dim mOilAmtMRPTB As Double, mOilAmtTB As Double
'Dim mTotMRPOilTB As Double, mTotOilTB As Double, mTotShareAmt As Double
'Dim mShareSpr As Single, mShareAmtSpr As Double, mShare2AmtSpr As Double
'Dim mTot1ShareAmt As Double, mTot2ShareAmt As Double, mTot3ShareAmt As Double
Dim mPrefix$, rsSPSal As ADODB.Recordset, mTxtDocID$, mTrans As Boolean

lblRefresh
Lbl(1) = TitleStr
Lbl(1).Refresh

Lbl(lblIndex) = "Posting in progress..."
Lbl(lblIndex).Visible = True
Lbl(lblIndex).Refresh
'
GSQL = "Select distinct V_Date,DocID,Party_Code,CrAc,Total_Amt from SP_Sale where left(DocID,1)='" & PubDivCode & "' and trim(" & cMID("DocID", "4", "5") & ")='" & mVType & "' and V_Date >=#" & txt(DateFrom) & "# and V_Date<=#" & txt(DateTo) & "# and CancelYN=0 Order By V_Date,DocID"
Set rsSPSal = GCn.Execute(GSQL)
If rsSPSal.RecordCount <= 0 Then
    Lbl(lblIndex) = "No Records for Posting!"
    MsgBox Lbl(lblIndex), vbInformation, "No Records!"
    GoTo lblExit
End If
Lbl(10) = rsSPSal.RecordCount
Lbl(10).Refresh

mTrans = True
GCn.BeginTrans
GCnFaS.BeginTrans

'Start Ledger Posting
Do While rsSPSal.EOF = False
    '**
    Erase LedgAry
    '**
    mResult = 0
    mNarr = "Through Stock Transfer Issue"
    mCommNarr = mNarr
    i = 1
    '****
    Lbl(11) = rsSPSal!V_DATE
    Lbl(11).Refresh
    
    mFADocID = rsSPSal!DocID
    ReDim Preserve LedgAry(1)
    LedgAry(i).SubCode = rsSPSal!Party_code
    LedgAry(i).AmtDr = rsSPSal!Total_Amt
    LedgAry(i).Narration = mNarr
    LedgAry(i).ContraSub = rsSPSal!CrAc
    i = i + 1
    i = UBound(LedgAry) + 1
    ReDim Preserve LedgAry(i)
    LedgAry(i).SubCode = rsSPSal!CrAc
    LedgAry(i).AmtCr = rsSPSal!Total_Amt
    LedgAry(i).Narration = mNarr
    LedgAry(i).ContraSub = rsSPSal!Party_code
       
    mResult = LedgerPost("A", LedgAry, GCnFaS, mFADocID, CDate(rsSPSal!V_DATE), mCommNarr)
    If mResult <> 1 Then MsgBox "Error in Ledger Posting : " & mFADocID, vbOKOnly, "Validation"
    '**
    Lbl(12) = rsSPSal.AbsolutePosition
    Lbl(12).Refresh
    rsSPSal.MoveNext
Loop
GCn.CommitTrans
GCnFaS.CommitTrans
mTrans = False

Lbl(lblIndex) = "Posting completed sucessfully !"
Lbl(lblIndex).Refresh

lblExit:
    Set RsTemp = Nothing
    Set rsSPSal = Nothing
    Set GRs = Nothing
    If mTrans Then
        GCn.RollbackTrans
        GCnFaS.RollbackTrans
    End If
    If err.NUMBER <> 0 Then
        MsgBox err.Description & vbCrLf & "Ledger Posting Terminated!", vbCritical
    End If
End Sub


Private Sub PostSprTrfRect(TrnType$, mVType$, TitleStr$, lblIndex As Integer)
Dim xNetAmt As Double, xEntryTaxAmt As Double
'A/c Posting related declarations
Dim LedgAry() As LedgRec
Dim mResult As Byte, mCommNarr$, mNarr$, TaxSQL$, i As Integer, j As Integer
Dim mSprPurPfx$, mFADocID$, mTxtDocID$, PartyAcCode$
Dim rsSPPurch As ADODB.Recordset, mTrans As Boolean

lblRefresh
Lbl(1) = TitleStr
Lbl(1).Refresh

Lbl(lblIndex) = "Posting in progress..."
Lbl(lblIndex).Visible = True
Lbl(lblIndex).Refresh

GSQL = "Select distinct V_Date,DocID,EntryTaxAmt,Party_Code,Party_Doc_No,Party_Doc_Date,NET_AMT+EntryTaxAmt as NetAmt,DrAc_Code " & _
    "from SP_Purch where left(DocID,1)='" & PubDivCode & "' and trim(" & cMID("DocID", "4", "5") & ")='" & mVType & "' and V_Date >=#" & txt(DateFrom) & "# and V_Date<=#" & txt(DateTo) & "# and CancelYN=0 Order By V_Date,DocID"

Set rsSPPurch = GCn.Execute(GSQL)
If rsSPPurch.RecordCount <= 0 Then
    Lbl(lblIndex) = "No Records for Posting!"
    MsgBox Lbl(lblIndex), vbInformation, "No Records!"
    GoTo lblExit
End If
Lbl(10) = rsSPPurch.RecordCount
Lbl(10).Refresh

mTrans = True
GCn.BeginTrans
GCnFaS.BeginTrans

'Start Ledger Posting
Do While rsSPPurch.EOF = False
    '**
    Erase LedgAry
    mCommNarr = ""
    mResult = 0
    mNarr = ""
    TaxSQL = ""
    i = 0
    xNetAmt = 0
    xEntryTaxAmt = 0
    '****
    Lbl(11) = rsSPPurch!V_DATE
    Lbl(11).Refresh
    mFADocID = rsSPPurch!DocID
    PartyAcCode = rsSPPurch!Party_code
    '**
    mNarr = "Spare Transfer Rect"
    '**
    If rsSPPurch!Party_Doc_No <> "" Then
        mNarr = mNarr & " Party Document No." & rsSPPurch!Party_Doc_No
    End If
    If rsSPPurch!Party_Doc_Date <> "" Then
        mNarr = mNarr & " Date " & rsSPPurch!Party_Doc_Date
    End If
    mCommNarr = mNarr & " [Common]"
    xEntryTaxAmt = VNull(rsSPPurch!EntryTaxAmt)
    '*** pURCHASE Amount Row
'   0.Purchase A/c
'   1.Party A/c or Cash A/c
'*********
    i = 1
    ReDim Preserve LedgAry(1)
    LedgAry(i).SubCode = rsSPPurch!DrAc_Code
    LedgAry(i).ContraSub = PartyAcCode
    LedgAry(i).AmtDr = IIf(IsNull(rsSPPurch!NetAmt), 0, rsSPPurch!NetAmt)
    LedgAry(i).Narration = mNarr
    xNetAmt = xNetAmt + IIf(IsNull(rsSPPurch!NetAmt), 0, rsSPPurch!NetAmt)
    If xEntryTaxAmt <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!EntryTax_Ac
        LedgAry(i).AmtCr = xEntryTaxAmt
        LedgAry(i).Narration = mNarr
    End If
    i = UBound(LedgAry) + 1
    ReDim Preserve LedgAry(i)
    LedgAry(i).SubCode = PartyAcCode
    LedgAry(i).AmtCr = (xNetAmt - xEntryTaxAmt)
    LedgAry(i).Narration = mNarr
    
    mResult = LedgerPost("A", LedgAry, GCnFaS, mFADocID, CDate(rsSPPurch!V_DATE), mCommNarr)
    If mResult <> 1 Then
        MsgBox "Error in Ledger Posting : " & mFADocID, vbOKOnly, "Validation"
    End If
    Lbl(12) = rsSPPurch.AbsolutePosition
    Lbl(12).Refresh

    rsSPPurch.MoveNext
Loop
GCn.CommitTrans
GCnFaS.CommitTrans
mTrans = False
Lbl(lblIndex) = "Posting completed sucessfully !"
Lbl(lblIndex).Refresh
'MsgBox TitleStr & vbCrLf & "Posting completed sucessfully !", vbInformation, "Ledger Posting"

lblExit:
    Set rsSPPurch = Nothing
    If mTrans Then
        GCn.RollbackTrans
        GCnFaS.RollbackTrans
    End If
    If err.NUMBER <> 0 Then
        MsgBox err.Description & vbCrLf & "Ledger Posting Terminated!", vbCritical
'        ProcAcPost = True
    End If
    End Sub

