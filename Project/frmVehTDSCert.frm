VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVehTDSCert 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TDS Certificate"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   9855
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
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2085
      Left            =   2235
      Negotiate       =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1365
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   3678
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
      Caption         =   "Party Help"
      ColumnCount     =   1
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
   Begin MSDataGridLib.DataGrid DGBook 
      Height          =   2085
      Left            =   -660
      Negotiate       =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2475
      Visible         =   0   'False
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   3678
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Booking No"
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
         Caption         =   "Booking DocId"
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
         DataField       =   "Tds_Per"
         Caption         =   "TDS%"
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
         DataField       =   "Tds_amt"
         Caption         =   "TDS Amt"
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
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2039.811
         EndProperty
      EndProperty
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
      Index           =   5
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1215
      Width           =   4260
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
      Index           =   7
      Left            =   2670
      TabIndex        =   7
      Top             =   1485
      Width           =   1950
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
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   2
      Top             =   675
      Width           =   975
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
      Index           =   6
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1485
      Width           =   975
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
      Index           =   3
      Left            =   2670
      TabIndex        =   3
      Top             =   675
      Width           =   1950
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
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Top             =   945
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
      Index           =   8
      Left            =   7350
      MaxLength       =   30
      TabIndex        =   8
      Top             =   90
      Width           =   2415
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
      Index           =   9
      Left            =   7350
      MaxLength       =   20
      TabIndex        =   9
      Top             =   360
      Width           =   2415
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
      Index           =   11
      Left            =   7350
      TabIndex        =   11
      Top             =   900
      Width           =   555
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
      Index           =   10
      Left            =   7350
      TabIndex        =   10
      Top             =   630
      Width           =   1485
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
      Index           =   12
      Left            =   7350
      TabIndex        =   12
      Top             =   1170
      Width           =   1485
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
      Left            =   1680
      TabIndex        =   0
      Top             =   135
      Width           =   2085
   End
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
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
      Index           =   3
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Printer "
      Top             =   4440
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
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   405
      Width           =   4260
   End
   Begin VB.OptionButton OptPlain 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D7C6C8&
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
      Left            =   2025
      TabIndex        =   13
      Top             =   2535
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton Optpre 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D7C6C8&
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
      Left            =   3390
      TabIndex        =   14
      Top             =   2535
      Width           =   1260
   End
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00F8D7FD&
      Caption         =   "Speed &Print"
      DisabledPicture =   "frmVehTDSCert.frx":0000
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
      Index           =   2
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Printer "
      Top             =   3540
      Width           =   1590
   End
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Screen"
      DisabledPicture =   "frmVehTDSCert.frx":030A
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
      Index           =   1
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Screen"
      Top             =   3840
      Width           =   1590
   End
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Windows Print"
      DisabledPicture =   "frmVehTDSCert.frx":0614
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
      Index           =   0
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Printer "
      Top             =   4140
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
      Left            =   435
      Picture         =   "frmVehTDSCert.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Screen"
      Top             =   4410
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Issue Dt."
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
      Left            =   75
      TabIndex        =   34
      Top             =   960
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Signature"
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
      Left            =   6060
      TabIndex        =   31
      Top             =   105
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.D.S.(%)"
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
      Left            =   6060
      TabIndex        =   30
      Top             =   915
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Value"
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
      Left            =   6060
      TabIndex        =   29
      Top             =   645
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
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
      Left            =   6060
      TabIndex        =   28
      Top             =   375
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.D.S. Amount"
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
      Left            =   6060
      TabIndex        =   27
      Top             =   1185
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No."
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
      Left            =   75
      TabIndex        =   26
      Top             =   150
      Width           =   1155
   End
   Begin VB.Line Line2 
      X1              =   4560
      X2              =   4560
      Y1              =   2445
      Y2              =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Challan No && Dt"
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
      Left            =   75
      TabIndex        =   25
      Top             =   1500
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party"
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
      Left            =   75
      TabIndex        =   24
      Top             =   420
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TDS Cert. No  && Dt."
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
      Left            =   75
      TabIndex        =   23
      Top             =   690
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name"
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
      Left            =   75
      TabIndex        =   22
      Top             =   1230
      Width           =   1125
   End
   Begin VB.Line Line1 
      X1              =   2085
      X2              =   2085
      Y1              =   2430
      Y2              =   2550
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
      Left            =   765
      TabIndex        =   21
      Top             =   4425
      Width           =   7275
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   225
      Index           =   41
      Left            =   2850
      TabIndex        =   20
      Top             =   2115
      Width           =   825
   End
   Begin VB.Line Line6 
      X1              =   4560
      X2              =   2085
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line8 
      X1              =   3210
      X2              =   3210
      Y1              =   2325
      Y2              =   2415
   End
End
Attribute VB_Name = "frmVehTDSCert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsParty As ADODB.Recordset
Dim RSBook As ADODB.Recordset
Dim PartyCode As String
Private Const BookNo As Byte = 0
Private Const Party As Byte = 1
Private Const TDSNo As Byte = 2
Private Const TDSDt As Byte = 3
Private Const TDSIssDt As Byte = 4
Private Const Bank As Byte = 5
Private Const ChalNo As Byte = 6
Private Const ChalDt As Byte = 7
Private Const Sign As Byte = 8
Private Const Desig As Byte = 9
Private Const Interest As Byte = 10
Private Const TDSPer As Byte = 11
Private Const TDSAmt As Byte = 12

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String

Private Sub DGBook_Click()
    DGBook.Visible = False
    If RSBook.RecordCount > 0 Then
            txtPrint(BookNo).Tag = RSBook!Code
            txtPrint(BookNo).TEXT = RSBook!Name
    End If
End Sub

Private Sub DGParty_Click()
    DGParty.Visible = False
    If RsParty.RecordCount > 0 Then
        txtPrint(Party).TEXT = RsParty!Name
        txtPrint(Party).Tag = RsParty!Code
    End If
    txtPrint(Party).SetFocus
End Sub

Private Sub DGSite_Click()

End Sub


Private Sub Form_Activate()
'If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim i As Byte
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "select SubGroup.Subcode as code,SubGroup.NAME from SubGroup Where  Nature='Supplier'  order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set RSBook = New ADODB.Recordset
    RSBook.CursorLocation = adUseClient
    RSBook.Open "SELECT " & cCStr("Veh_Order.Ord_No") & " as Name,Veh_Order.OrdDocId as Code,Veh_Order.partycode,SubGroup.Name as PartyName, Veh_Order.Interest,Veh_Order.TDS_Per,Veh_Order.TDS_Amt " & _
    "FROM Veh_Order LEFT JOIN Subgroup ON Veh_Order.PartyCode = Subgroup.subcode  " & _
    "where veh_order.OrdDocId <> '' and veh_order.DelCh_DocId <> ''", GCn, adOpenDynamic, adLockOptimistic
    Set DGBook.DataSource = RSBook
    
    DGBook.left = 945: DGBook.top = 2100
    DGParty.left = 945: DGParty.top = 2100
    
    CmdPrint(PDos).Enabled = False
LblPrinter.CAPTION = Printer.DeviceName
Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RSBook = Nothing
Set RsParty = Nothing
End Sub
Private Sub Grid_Hide()
    If DGBook.Visible = True Then DGBook.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
End Sub


'************************ PRINTING CODE ******************

Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case BookNo
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Then Exit Sub
            If txtPrint(Index).TEXT = "" Then
                txtPrint(Index).Tag = ""
                txtPrint(Index).TEXT = ""
            Else
                If txtPrint(Index).Tag <> RSBook!Code Then
                    RSBook.MoveFirst
                    RSBook.FIND "code ='" & txtPrint(BookNo).Tag & "'"
                End If
            End If
    Case Party
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Then Exit Sub
        If txtPrint(Index).TEXT = "" Then
            txtPrint(Index).Tag = ""
            txtPrint(Index).TEXT = ""
        Else
            If txtPrint(Index).TEXT <> RsParty!Name Then
                RsParty.MoveFirst
                RsParty.FIND "Code ='" & txtPrint(Index).Tag & "'"
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
    Case BookNo
        DGridTxtKeyDown DGBook, txtPrint, Index, RSBook, KeyCode, False, 0
    Case Party
        DGridTxtKeyDown DGParty, txtPrint, Index, RsParty, KeyCode, False, 1
End Select
If DGParty.Visible = False And DGBook.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> BookNo Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TxtPrint_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)
Select Case Index
    Case Party
        If DGParty.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsParty, keyascii, "Name"
    Case BookNo
        If DGBook.Visible = True Then DGridTxtKeyPress txtPrint, Index, RSBook, keyascii, "Name"
End Select
End Sub

Private Sub TxtPrint_LostFocus(Index As Integer)
    Ctrl_validate txtPrint(Index)
End Sub

Private Sub TxtPrint_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
   Case TDSDt, TDSIssDt, ChalDt
        txtPrint(Index) = RetDate(txtPrint(Index))
   Case BookNo
        If RSBook.RecordCount = 0 Or (RSBook.EOF = True Or RSBook.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(BookNo).TEXT = ""
            txtPrint(BookNo).Tag = ""
        Else
            txtPrint(BookNo).Tag = RSBook!Code
            txtPrint(BookNo).TEXT = RSBook!Name
        End If
        FillData
   Case Party
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(Index).TEXT = ""
            txtPrint(Index).Tag = ""
        Else
            txtPrint(Index).TEXT = RsParty!Name
            txtPrint(Index).Tag = RsParty!Code
        End If

End Select
End Sub

Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If
End Sub

Private Sub CmdPrint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
Select Case Index
    Case PScreen, PWindows, PDos
        mRepName = IIf(OptPlain.Value = True, "VehTDSCert", "VehTDSCert")
        Call WindowsPrint(Index)
'    Case PDos
'        Call SpeedPrint
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "VehTDSCert", "VehTDSCert")
        Call PrinerSetUp
    Case PClose 'Close Report Frame
        Unload Me
End Select
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub WindowsPrint(Index As Integer)
Dim Rst As ADODB.Recordset, RstSub1 As ADODB.Recordset, RstSub2 As ADODB.Recordset, mQRY As String
Dim i As Integer, Rst2 As ADODB.Recordset
On Error GoTo ERRORHANDLER
     
     If IsValid(txtPrint(Party), "Party") = False Then Exit Sub
     
     If txtPrint(BookNo).TEXT <> "" Then
     GCn.BeginTrans
        GCn.Execute ("UpDate veh_order set " & _
        "TDS_CNO='" & txtPrint(TDSNo) & "',TDS_CDATE=" & ConvertDate(txtPrint(TDSDt)) & ",TDS_IDATE=" & ConvertDate(txtPrint(TDSIssDt)) & ",TDS_SIGN='" & txtPrint(Sign) & "', " & _
        "TDS_DESIG='" & txtPrint(Desig) & "',TDS_BankName='" & txtPrint(Bank) & "',TDS_ChalNo='" & txtPrint(ChalNo) & "',TDS_ChalDate=" & ConvertDate(txtPrint(ChalDt)) & " " & _
        " where OrdDocId = '" & txtPrint(BookNo).Tag & "'")
     GCn.CommitTrans
     End If
               
     mQRY = "select subgroup.Add1,subgroup.Add2,subgroup.Add3,City.Cityname,subgroup.name as Party,Interest,TDS_Per,TDS_Amt,TDS_CNO,TDS_CDATE,TDS_IDATE,TDS_SIGN,TDS_DESIG,TDS_BankName,TDS_ChalDate,TDS_ChalNo " & _
        " FROM Veh_Order LEFT JOIN (SubGroup LEFT JOIN City ON SubGroup.CityCode = City.CityCode) ON Veh_Order.PartyCode = SubGroup.SubCode " & _
        " where OrdDocId = '" & txtPrint(BookNo).Tag & "'"
    
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
     
       
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
                     
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    rpt.Database.SetDataSource Rst
       
    Set Rst2 = New ADODB.Recordset
    Rst2.CursorLocation = adUseClient
    Rst2.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
    For i = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
            Case UCase("SubTitle")
                rpt.FormulaFields(i).TEXT = "'" & Rst2!V_SecSpeciality & "'"
            Case UCase("LST")
                rpt.FormulaFields(i).TEXT = "'" & Rst2!V_SecLST & "'"
            Case UCase("LSTDate")
                rpt.FormulaFields(i).TEXT = "'" & Rst2!V_SecLST_Date & "'"
            Case UCase("CST")
                rpt.FormulaFields(i).TEXT = "'" & Rst2!V_SecCST & "'"
            Case UCase("CSTDate")
                rpt.FormulaFields(i).TEXT = "'" & Rst2!V_SecCST_Date & "'"
            Case UCase("Phone")
                rpt.FormulaFields(i).TEXT = "'" & Rst2!V_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(i).TEXT = "'" & Rst2!V_SecFax & "'"
            Case UCase("Gram")
                rpt.FormulaFields(i).TEXT = "'" & Rst2!V_SecGram & "'"
        End Select
    Next
   
   
    rpt.ReadRecords

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
                Case UCase("Title")
                    rpt.FormulaFields(i).TEXT = "'" & Me.CAPTION & "'"
            End Select
            Next
            rpt.PrintOut False
        Case PScreen  'screen
            Call Report_View(rpt, Me.CAPTION, , True)
        Case PDos
            Call Report_View(rpt, Me.CAPTION, 1)
End Select
Set Rst = Nothing
Set Rst2 = Nothing
Set rpt = Nothing
CmdPrint(PSetUp).Tag = ""
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

Private Sub FillData()
Dim Rst As ADODB.Recordset
    If txtPrint(BookNo).TEXT <> "" And RSBook.EOF = False And RSBook.BOF = False And RSBook.RecordCount > 0 Then
        Set Rst = GCn.Execute("select Interest,TDS_Per,TDS_Amt,TDS_CNO,TDS_CDATE,TDS_IDATE,TDS_SIGN,TDS_DESIG,TDS_BankName,TDS_ChalNo,TDS_ChalDate from Veh_order where OrdDocId = '" & RSBook!Code & "'")
        
        txtPrint(Party).Tag = RSBook!PartyCode
        txtPrint(Party).TEXT = RSBook!PartyName
        
        txtPrint(TDSNo).TEXT = XNull(Rst!TDS_CNO)
        txtPrint(TDSDt).TEXT = IIf(IsNull(Rst!TDS_CDATE), "", Rst!TDS_CDATE)
        txtPrint(TDSIssDt).TEXT = IIf(IsNull(Rst!TDS_IDATE), "", Rst!TDS_IDATE)
        
        txtPrint(ChalNo).TEXT = XNull(Rst!TDS_ChalNo)
        txtPrint(ChalDt).TEXT = IIf(IsNull(Rst!TDS_ChalDate), "", Rst!TDS_ChalDate)
        txtPrint(Bank).TEXT = XNull(Rst!TDS_BankName)
        
        txtPrint(Sign).TEXT = XNull(Rst!TDS_SIGN)
        txtPrint(Desig).TEXT = XNull(Rst!TDS_DESIG)
        txtPrint(Interest).TEXT = Format(Rst!Interest, "0.00")
        txtPrint(TDSPer).TEXT = Format(Rst!TDS_Per, "0.00")
        txtPrint(TDSAmt).TEXT = Format(Rst!TDS_Amt, "0.00")
        
        txtPrint(Party).Enabled = False
        txtPrint(Interest).Enabled = False
        txtPrint(TDSPer).Enabled = False
        txtPrint(TDSAmt).Enabled = False
        
    Else
        txtPrint(Party).Tag = ""
        txtPrint(Party).TEXT = ""
        
        txtPrint(TDSNo).TEXT = ""
        txtPrint(TDSDt).TEXT = ""
        txtPrint(TDSIssDt).TEXT = ""
        
        txtPrint(ChalNo).TEXT = ""
        txtPrint(ChalDt).TEXT = ""
        txtPrint(Bank).TEXT = ""
        
        txtPrint(Sign).TEXT = ""
        txtPrint(Desig).TEXT = ""
        txtPrint(Interest).TEXT = Format(0, "0.00")
        txtPrint(TDSPer).TEXT = Format(0, "0.00")
        txtPrint(TDSAmt).TEXT = Format(0, "0.00")
        
        txtPrint(Party).Enabled = True
        txtPrint(Interest).Enabled = True
        txtPrint(TDSPer).Enabled = True
        txtPrint(TDSAmt).Enabled = True
    End If
End Sub
