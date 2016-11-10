VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmWarrantyCredit 
   BackColor       =   &H0095AFCC&
   Caption         =   "Warranty Credit Note Entry"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6045
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
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   6045
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DGAccount 
      Height          =   4755
      Left            =   10620
      Negotiate       =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   8387
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
            DividerStyle    =   3
            ColumnWidth     =   4004.788
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   794.835
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   3630
      MaxLength       =   12
      TabIndex        =   21
      Top             =   7290
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
      Left            =   5160
      MaxLength       =   12
      TabIndex        =   23
      Top             =   5940
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   5160
      MaxLength       =   12
      TabIndex        =   26
      Top             =   6750
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
      Left            =   5160
      MaxLength       =   12
      TabIndex        =   24
      Top             =   6210
      Width           =   1470
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
      Height          =   255
      Index           =   23
      Left            =   5160
      MaxLength       =   12
      TabIndex        =   25
      Top             =   6480
      Width           =   1470
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
      Height          =   255
      Index           =   20
      Left            =   5160
      MaxLength       =   12
      TabIndex        =   22
      Top             =   5670
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   3630
      MaxLength       =   12
      TabIndex        =   20
      Top             =   7020
      Width           =   1470
   End
   Begin VB.TextBox txtgrid1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
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
      Height          =   270
      Index           =   0
      Left            =   8085
      MaxLength       =   40
      TabIndex        =   8
      Top             =   3855
      Visible         =   0   'False
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
      Index           =   26
      Left            =   6885
      MaxLength       =   12
      TabIndex        =   28
      Top             =   5940
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   6885
      MaxLength       =   12
      TabIndex        =   31
      Top             =   6750
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
      Left            =   6885
      MaxLength       =   12
      TabIndex        =   29
      Top             =   6210
      Width           =   1470
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
      Height          =   255
      Index           =   28
      Left            =   6885
      MaxLength       =   12
      TabIndex        =   30
      Top             =   6480
      Width           =   1470
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
      Height          =   255
      Index           =   25
      Left            =   6885
      MaxLength       =   12
      TabIndex        =   27
      Top             =   5670
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
      Index           =   14
      Left            =   3630
      MaxLength       =   12
      TabIndex        =   16
      Top             =   5940
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   3630
      MaxLength       =   12
      TabIndex        =   19
      Top             =   6750
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
      Index           =   15
      Left            =   3630
      MaxLength       =   12
      TabIndex        =   17
      Top             =   6210
      Width           =   1470
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
      Height          =   255
      Index           =   16
      Left            =   3630
      MaxLength       =   12
      TabIndex        =   18
      Top             =   6480
      Width           =   1470
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
      Height          =   255
      Index           =   13
      Left            =   3630
      MaxLength       =   12
      TabIndex        =   15
      Top             =   5670
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2130
      MaxLength       =   8
      TabIndex        =   1
      Top             =   705
      Width           =   855
   End
   Begin VB.TextBox txt 
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
      Index           =   2
      Left            =   4935
      MaxLength       =   25
      TabIndex        =   2
      Top             =   705
      Width           =   1425
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
      Index           =   3
      Left            =   1515
      MaxLength       =   12
      TabIndex        =   3
      Top             =   975
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
      Left            =   1515
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1245
      Width           =   4845
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
      Left            =   1515
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1515
      Width           =   4845
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
      Left            =   1515
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1785
      Width           =   4845
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
      Height          =   255
      Index           =   8
      Left            =   2100
      MaxLength       =   12
      TabIndex        =   10
      Top             =   5670
      Width           =   1470
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
      Height          =   255
      Index           =   11
      Left            =   2100
      MaxLength       =   12
      TabIndex        =   13
      Top             =   6480
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
      Index           =   10
      Left            =   2100
      MaxLength       =   12
      TabIndex        =   12
      Top             =   6210
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   2100
      MaxLength       =   12
      TabIndex        =   14
      Top             =   6750
      Width           =   1470
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   661
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
      Index           =   4
      Left            =   4935
      MaxLength       =   25
      TabIndex        =   4
      Top             =   975
      Width           =   1425
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
      Index           =   9
      Left            =   2100
      MaxLength       =   12
      TabIndex        =   11
      Top             =   5940
      Width           =   1470
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   2865
      Left            =   135
      TabIndex        =   9
      Top             =   2070
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   5054
      _Version        =   393216
      BackColor       =   12243913
      Cols            =   6
      BackColorFixed  =   4210816
      ForeColorFixed  =   65535
      BackColorSel    =   16711680
      BackColorBkg    =   11189721
      GridColor       =   128
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      GridLineWidthFixed=   1
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
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Credit Note Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   7
      Left            =   150
      TabIndex        =   55
      Top             =   7305
      Width           =   1980
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Difference Amt"
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
      Index           =   6
      Left            =   5355
      TabIndex        =   54
      Top             =   5415
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less : TDS Deducted"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   5
      Left            =   150
      TabIndex        =   53
      Top             =   7035
      Width           =   1755
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rejected Amount"
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
      Index           =   4
      Left            =   6900
      TabIndex        =   52
      Top             =   5415
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Passed Amount"
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
      Left            =   3780
      TabIndex        =   51
      Top             =   5400
      Width           =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Claimed Amount"
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
      Left            =   2190
      TabIndex        =   50
      Top             =   5400
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Claim Diff A/c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   3
      Left            =   150
      TabIndex        =   49
      Top             =   1785
      Width           =   1095
   End
   Begin VB.Label lblPrefix 
      BackStyle       =   0  'Transparent
      Caption         =   "VPREFIX"
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
      Left            =   1515
      TabIndex        =   47
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblDocId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DocId"
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
      Left            =   8235
      TabIndex        =   46
      Top             =   1050
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   13
      Left            =   150
      TabIndex        =   45
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Note No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   44
      Top             =   990
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warranty A/c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   26
      Left            =   150
      TabIndex        =   43
      Top             =   1515
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recd. From"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   39
      Left            =   150
      TabIndex        =   42
      Top             =   1245
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Misc Amt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   23
      Left            =   150
      TabIndex        =   41
      Top             =   6225
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour+Spl Amt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   40
      Left            =   150
      TabIndex        =   40
      Top             =   5685
      Width           =   1320
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division :"
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
      Left            =   7215
      TabIndex        =   39
      Top             =   780
      Width           =   750
   End
   Begin VB.Label lblDocCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DocID     :"
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
      Left            =   7215
      TabIndex        =   38
      Top             =   1050
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Dt."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   11
      Left            =   3885
      TabIndex        =   37
      Top             =   720
      Width           =   945
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
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   29
      Left            =   150
      TabIndex        =   36
      Top             =   6495
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   20
      Left            =   150
      TabIndex        =   35
      Top             =   6765
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Note Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   8
      Left            =   3435
      TabIndex        =   34
      Top             =   975
      Width           =   1365
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   660
      Left            =   7035
      Top             =   705
      Width           =   4680
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code      :"
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
      Left            =   9585
      TabIndex        =   33
      Top             =   780
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NDP Amt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00471A71&
      Height          =   225
      Index           =   35
      Left            =   150
      TabIndex        =   32
      Top             =   5955
      Width           =   750
   End
End
Attribute VB_Name = "frmWarrantyCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TAddMode As Boolean
Dim ExitCtrl As Boolean
Dim GridKey As Integer

Dim VoucherEditFlag As Boolean
Dim ProfDocId As String
Dim ForSiteCode As String

Private Const VType As String = "W_WCR"

Dim MyIndex As Byte
Dim MyCardNo As String
Dim Rst As ADODB.Recordset

Dim Master As ADODB.Recordset
Dim Master1 As ADODB.Recordset
Dim RsPending As ADODB.Recordset
Dim RsParty As ADODB.Recordset

'Text Box (Form)
Private Const VNo As Byte = 1
Private Const VDate As Byte = 2
Private Const CrNoteNo As Byte = 3
Private Const CrNoteDt As Byte = 4
Private Const RecdFromAc As Byte = 5
Private Const WarrAc As Byte = 6
Private Const DiffAc As Byte = 7
Private Const ClmLabour As Byte = 8
Private Const ClmNDP As Byte = 9
Private Const ClmMisc As Byte = 10
Private Const ClmTax As Byte = 11
Private Const ClmTotal As Byte = 12
Private Const PassLabour As Byte = 13
Private Const PassNDP As Byte = 14
Private Const PassMisc As Byte = 15
Private Const PassTax As Byte = 16
Private Const PassTotal As Byte = 17
Private Const TDSAmt As Byte = 18
Private Const CNoteAmt As Byte = 19
Private Const DiffLabour As Byte = 20
Private Const DiffNDP As Byte = 21
Private Const DiffMisc As Byte = 22
Private Const DiffTax As Byte = 23
Private Const DiffTotal As Byte = 24
Private Const RejLabour As Byte = 25
Private Const RejNDP As Byte = 26
Private Const RejMisc As Byte = 27
Private Const RejTax As Byte = 28
Private Const RejTotal As Byte = 29

'Fgrid1 Columns
Private Const C_ClmNo As Byte = 1
Private Const C_ClmDate As Byte = 2
Private Const C_PartNo As Byte = 3

Private Const C_Reject As Byte = 4

Private Const C_ClmLab As Byte = 5
Private Const C_PassLab As Byte = 6
Private Const C_ClmSpl As Byte = 7
Private Const C_PassSpl As Byte = 8
Private Const C_ClmSrv As Byte = 9
Private Const C_PassSrv As Byte = 10

Private Const C_ClmQty As Byte = 11
Private Const C_PassQty As Byte = 12
Private Const C_ClmNDP As Byte = 13
Private Const C_PassNDP As Byte = 14
Private Const C_ClmMisc As Byte = 15
Private Const C_PassMisc As Byte = 16
Private Const C_ClmLST As Byte = 17
Private Const C_PassLST As Byte = 18
Private Const C_ClmSurc As Byte = 19
Private Const C_PassSurc As Byte = 20
Private Const C_ClmTot As Byte = 21
Private Const C_PassTOT As Byte = 22

Private Const C_ClmAmt As Byte = 23
Private Const C_PassAmt As Byte = 24

Private Const C_PartName As Byte = 25

Private Const C_SrlNo As Byte = 26
Private Const C_Div As Byte = 27
Private Const C_Site As Byte = 28
Private Const C_TaxYN As Byte = 29

Private Sub DGAccount_Click()
If RsParty.RecordCount > 0 Then
    Txt(MyIndex).Tag = RsParty!Code
    Txt(MyIndex).TEXT = RsParty!Name
End If
DGAccount.Visible = False
Txt(MyIndex).SetFocus
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
    
    '' pending Points :
    '' Calculation of Tax Claimed columns  (in Grid)
    
    TopCtrl1.Tag = UserPermission(Me.Name)
    ForSiteCode = PubSiteCode
    Call BlankText
        
    lblPrefix.CAPTION = ""
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select Distinct JW2.Div_code,JW2.site_code, JW2.V_docId as SearchCode,JW2.V_DocId,JW2.V_No,JW2.V_Date,JW2.CrNoteNo,JW2.CrNoteDate,JW2.RecdFromCode,JW2.WarrCode,JW2.TDS_Amt,JW2.DiffCode,Supp.Name as RecdFromName,Warr.Name as WarrName,Diff.Name as DiffName from ((Job_Warr2 as JW2 left join Subgroup as Supp on JW2.RecdFromCode=Supp.Subcode) left join Subgroup as Warr on JW2.WarrCode=Warr.Subcode) left join Subgroup as Diff on JW2.DiffCode=Diff.Subcode  WHERE JW2.Div_code = '" & PubDivCode & "' and (JW2.V_DocId<>'' and  not isnull(JW2.v_docid)) order by JW2.V_No", GCn, adOpenDynamic, adLockOptimistic
    
    Set Master1 = New ADODB.Recordset
    Master1.CursorLocation = adUseClient
    Master1.Open "select JW2.ProwNo,JW2.Prowdt,JW2.Div_code,JW2.site_code,JW2.docid,JW2.SrlNo,JW2.V_docId,JW2.Claim_Rejected,JW2.Part_No,JW2.TotQty,JW2.MRP_YN,JW2.Tax_YN,JW2.price,JW2.Labour_Amt,JW2.Misc_Chrg,JW2.Spl_amt,JW2.Qty_Pass,JW2.Labour_Pass,JW2.Spl_Pass,JW2.Spr_Pass,JW2.Misc_Pass,JW2.Lst_Pass,JW2.Surc_Pass,JW2.TOT_Pass,SrvTax_Pass,JW2.V_NO, JW2.V_DATE,JW2.CRNOTENO, JW2.CRNOTEDATE,part.part_name  from Job_Warr2 as JW2 left join part on JW2.part_no=part.part_no WHERE JW2.Div_code = '" & PubDivCode & "' and (JW2.V_DocId<>'' and not isnull(JW2.v_docid)) order by JW2.V_No", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsPending = New ADODB.Recordset
    RsPending.CursorLocation = adUseClient
    RsPending.Open "select JW2.ProwNo as ProwNo,JW2.Prowdt as ProwDt,JW2.Div_code,JW2.site_code,JW2.docid, JW2.SrlNo,JW2.V_docId,JW2.Claim_Rejected,Job_Warr1.ProwNo,Job_Warr1.prowdt,JW2.Part_No, JW2.totQty, JW2.MRP_YN, JW2.Tax_YN, JW2.price, JW2.Labour_Amt, JW2.Misc_Chrg, JW2.Spl_Amt, JW2.Qty_Pass, JW2.Labour_Pass, JW2.Spl_Pass, JW2.Spr_Pass,JW2.Misc_Pass,JW2.Lst_Pass,JW2.Surc_Pass,JW2.TOT_Pass,SrvTax_Pass,JW2.V_No,JW2.V_Date,JW2.CRNOTENO, JW2.CRNOTEDATE,part.part_name from (job_warr2 as JW2 LEFT JOIN Job_Warr1 ON JW2.DocID = Job_Warr1.DocID) left join part on JW2.part_no=part.part_no WHERE  JW2.Div_code = '" & PubDivCode & "' and (JW2.V_DocId='' or isnull(JW2.v_docid)) and Job_Warr1.WBill_DocId<>'' order by JW2.V_No", GCn, adOpenDynamic, adLockOptimistic
    
    
    Set RsParty = New ADODB.Recordset
    With RsParty
        .CursorLocation = adUseClient
        .Open "Select Subcode as code,Name  FROM SubGroup Order by Name", GCn, adOpenDynamic, adLockOptimistic
        .Sort = "Name"
    End With
    Set DGAccount.DataSource = RsParty
    
    Ini_Grid
    Call MoveRec
    Disp_Text SETS("INI", Me, Master)
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
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

Private Sub Form_Resize()
    If Me.WindowState <> vbMaximized Then
        Me.left = MDIForm1.left
    End If
    Ini_Grid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set Master1 = Nothing
    Set RsPending = Nothing
    Set RsParty = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    If RsPending.RecordCount = 0 Then
        MsgBox "No Claims are pending for Settlement", vbInformation, "Validation"
'        Exit Sub
    End If
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
     
    Txt(VDate).TEXT = Format(date, "dd/MMM/yyyy")
    Txt(VNo).TEXT = GCn.Execute("select " & vIsNull("max(V_no)", "0") & "+1 from Job_Warr2 where left(v_docid,1)='" & PubDivCode & "' and " & cMID("v_docid", "2", "2") & "='" & PubSiteCode + ForSiteCode & "' AND " & cMID("v_docid", "4", "5") & "='" & VType & "'").Fields(0)
    lblDocId = GetDocID(GCnFaS, VType, Txt(VDate).TEXT, VoucherEditFlag, Txt(VNo), lblPrefix, ForSiteCode)
    
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode

    Call Fill_GridNew
    
    If VoucherEditFlag = True Then
        Txt(VNo).Enabled = True
        Txt(VNo).SetFocus
    Else
        Txt(VDate).SetFocus
    End If
    Call txtDisabled_Color(Me)
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant
Dim LedgAry(1) As LedgRec, mResult As Byte
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            vBook = Master.AbsolutePosition
            GCn.BeginTrans
            GCnFaS.BeginTrans
            
            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, lblDocId)
            If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
    
            GSQL = "Update Job_Warr2 set " _
                    & "v_DocId='',v_No=0,V_date= Null,crnoteno=''," _
                    & "CrNoteDate=Null,recdfromCode='',WarrCode='',diffcode=''," _
                    & "claim_rejected=0,Qty_Pass=0,Labour_Pass=0,Spl_Pass=0," _
                    & "Spr_Pass=0,Misc_Pass=0,LST_Pass=0,Surc_Pass=0,tds_amt=0," _
                    & "tot_pass=0,SrvTax_Pass=0,U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E'" _
                    & "where V_docid='" & lblDocId & "'"
            GCn.Execute GSQL
            
            GCnFaS.CommitTrans
            GCn.CommitTrans
            
            
            Master.Requery
            Call UpdRequery
    
            If Master.RecordCount > 0 Then
                If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
            Else
                Call BlankText
            End If
            BUTTONS True, Me, Master, 0
        End If
    Else
        MsgBox "No Records To Delete!", vbInformation, "Information"
    End If
    Exit Sub
eloop1:
    GCn.RollbackTrans: GCnFaS.RollbackTrans
    MsgBox err.Description, vbCritical, " Deletion Message"
End Sub

Private Sub TopCtrl1_eEdit()
Dim I As Integer
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    Txt(VNo).Enabled = False
    Txt(VDate).SetFocus
    Call Fill_GridNew
    Call txtDisabled_Color(Me)
    Exit Sub
eloop1:
    MsgBox err.Description, vbExclamation, " Editing Message"
End Sub

Private Sub TopCtrl1_eExit()
    If TopCtrl1.TopText2 = "Browse" Then Unload Me
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
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("searchcode='" & MyValue & "'")
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
    If MsgBox("Cancel Entry ?", vbExclamation + vbYesNo, "Terminate Process") = vbYes Then
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
    Call UpdRequery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim AddFlg As String
    Dim LedgAry(4) As LedgRec, mNarr$, mResult As Byte
'    On Error GoTo errlbl

    If txtgrid1(0).Visible = True Then
        If TxtGrid1Leave = False Then
            TxtGrid1_LostFocus 0
            txtgrid1(0).SetFocus
            Exit Sub
        Else
            txtgrid1(0).Visible = False
        End If
    End If

    Grid_Hide
    
    If IsValid(Txt(VNo), "Voucher No.") = False Then Exit Sub
    If IsValid(Txt(VDate), "Voucher Date") = False Then Exit Sub
    If IsValid(Txt(CrNoteNo), "Credit Note No.") = False Then Exit Sub
    If IsValid(Txt(CrNoteDt), "Credit Note Date") = False Then Exit Sub
    If IsValid(Txt(RecdFromAc), "Recd. From Supplier Name") = False Then Exit Sub
    If IsValid(Txt(WarrAc), "Warranty A/c") = False Then Exit Sub
    If IsValid(Txt(DiffAc), "Diff A/c") = False Then Exit Sub
    
    GCn.BeginTrans
    mTrans = True

    Select Case TopCtrl1.TopText2
        Case "Add"
            AddFlg = "A"
            If VoucherEditFlag = True Then
                GSQL = "Select Count(*) From Job_Warr2 Where v_DocID='" & lblDocId & "'"
                If GCn.Execute(GSQL).Fields(0) > 0 Then
                    MsgBox "Voucher No. " & Txt(VNo).TEXT & " Already Exists", vbCritical, "Validation Ertror"
                    Txt(VNo).SetFocus
                    Exit Sub
                End If
            Else
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select VT.Number_Method,VP.V_Type,VP.Date_From,VP.Prefix,VP.Start_Srl_No From Voucher_Type VT Left Join Voucher_Prefix VP on VT.V_Type=VP.V_Type Where VP.V_Type='" & VType & "' And VP.Date_From<=" & ConvertDate(Format(Txt(VDate).TEXT, "dd/MMM/yyyy")) & " Order By VP.Date_From DESC", GCn, adOpenDynamic, adLockOptimistic
                If Val(Rst!start_srl_no) >= Val(Txt(VNo).TEXT) Then
                    lblDocId = GetDocID(GCnFaS, VType, Txt(VDate).TEXT, VoucherEditFlag, Txt(VNo), lblPrefix, ForSiteCode)
                End If
                If Rst.RecordCount > 0 Then
                    GSQL = "Update Voucher_Prefix Set Start_Srl_No=Start_Srl_No+1 Where V_Type='" & Rst!V_Type & "' and Date_From=" & ConvertDate(Format(Rst!Date_From, "dd/MMM/yyyy")) & ""
                    GCn.Execute GSQL
                End If
            End If
            
        Case "Edit"
            AddFlg = "E"
                GSQL = "Update Job_Warr2 set " _
                        & "v_DocId='',v_No=0,V_date= Null,crnoteno=''," _
                        & "CrNoteDate=Null,recdfromCode='',WarrCode='',diffcode=''," _
                        & "claim_rejected=0,Qty_Pass=0,Labour_Pass=0,Spl_Pass=0," _
                        & "Spr_Pass=0,Misc_Pass=0,LST_Pass=0,Surc_Pass=0,tds_amt=0," _
                        & "tot_pass=0,SrvTax_Pass=0,U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='" & AddFlg & "' " _
                        & "where V_docid='" & lblDocId & "'"
                GCn.Execute GSQL
    End Select
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, C_Reject) = "Yes" Or Val(FGrid1.TextMatrix(I, C_PassLab)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassSpl)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassNDP)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassMisc)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassLST)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassSurc)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassTOT)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassSrv)) > 0 Then
            GSQL = "Update Job_Warr2 set " _
                    & "v_DocId='" & lblDocId & "',v_No='" & Val(Txt(VNo)) & "',V_date=" & ConvertDate(Txt(VDate)) & ",crnoteno='" & Txt(CrNoteNo) & "'," _
                    & "CrNoteDate=" & ConvertDate(Txt(CrNoteDt)) & ",recdfromCode='" & Txt(RecdFromAc).Tag & "',WarrCode='" & Txt(WarrAc).Tag & "',diffcode='" & Txt(DiffAc).Tag & "'," _
                    & "claim_rejected=" & IIf(FGrid1.TextMatrix(I, C_Reject) = "Yes", 1, 0) & ",Qty_Pass=" & Val(FGrid1.TextMatrix(I, C_PassQty)) & ",Labour_Pass=" & Val(FGrid1.TextMatrix(I, C_PassLab)) & ",Spl_Pass=" & Val(FGrid1.TextMatrix(I, C_PassSpl)) & "," _
                    & "Spr_Pass=" & Val(FGrid1.TextMatrix(I, C_PassNDP)) & ",Misc_Pass=" & Val(FGrid1.TextMatrix(I, C_PassMisc)) & ",LST_Pass=" & Val(FGrid1.TextMatrix(I, C_PassLST)) & ",Surc_Pass=" & Val(FGrid1.TextMatrix(I, C_PassSurc)) & ",tds_amt=" & Val(Txt(TDSAmt)) & "," _
                    & "tot_pass=" & Val(FGrid1.TextMatrix(I, C_PassTOT)) & ",SrvTax_Pass=" & Val(FGrid1.TextMatrix(I, C_PassSrv)) & ",U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='" & AddFlg & "' " _
                    & "where Div_Code='" & FGrid1.TextMatrix(I, C_Div) & "' and site_code='" & FGrid1.TextMatrix(I, C_Site) & "' and prowno ='" & FGrid1.TextMatrix(I, C_ClmNo) & "' and ProwDt='" & FGrid1.TextMatrix(I, C_ClmDate) & "' and srlno=" & FGrid1.TextMatrix(I, C_SrlNo)
            GCn.Execute GSQL
        End If
    Next I
    GCnFaS.BeginTrans
    'Ac Posting
        Set GRs = New ADODB.Recordset
        Set GRs = G_FaCn.Execute("Select TDS_Ac from AcControls")
        If GRs.RecordCount > 0 Then
            If GRs!TDS_Ac = "" Or GRs!TDS_Ac = Null Then
                MsgBox "Please Define TDS AC In A/c Controls" & vbCrLf & "A/c Posting Aborted !"
                GoTo lblExit
            End If
        End If
    
        mNarr = "Through Warranti Credit Note"
        I = 0
        LedgAry(I).SubCode = Txt(RecdFromAc).Tag
        LedgAry(I).AmtCr = Val(Txt(CNoteAmt))
        LedgAry(I).Narration = mNarr
                
        I = I + 1
        LedgAry(I).SubCode = Txt(RecdFromAc).Tag
        LedgAry(I).AmtCr = Val(Txt(ClmTotal)) - Val(Txt(CNoteAmt))
        LedgAry(I).Narration = mNarr
        
        I = I + 1
        LedgAry(I).SubCode = Txt(WarrAc).Tag
        LedgAry(I).AmtDr = Val(Txt(CNoteAmt))
        LedgAry(I).Narration = mNarr
        
        I = I + 1
        LedgAry(I).SubCode = Txt(DiffAc).Tag
        LedgAry(I).AmtDr = Val(Txt(ClmTotal)) - Val(Txt(CNoteAmt)) - Val(Txt(TDSAmt))
        LedgAry(I).Narration = mNarr
        
        I = I + 1
        LedgAry(I).SubCode = GRs!TDS_Ac
        LedgAry(I).AmtDr = Val(Txt(TDSAmt))
        LedgAry(I).Narration = mNarr
        
        
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, lblDocId, CDate(Txt(VDate)))
        
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"

lblExit:
    GCnFaS.CommitTrans
    GCn.CommitTrans
    
    mTrans = False
    
    Master.Requery
    
    Call UpdRequery
    
    Master.FIND "searchcode = '" & lblDocId & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub

errlbl:
    If mTrans = True Then
        GCn.RollbackTrans: GCnFaS.RollbackTrans: CheckError
    Else
        CheckError
    End If
Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus Txt(Index)
    txtgrid1(0).Visible = False
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case RecdFromAc, WarrAc, DiffAc
            If RsParty.RecordCount = 0 Or Txt(Index).TEXT = "" Or RsParty.EOF = True Or RsParty.BOF = True Then Exit Sub
            If Txt(Index).Tag <> RsParty!Code Then
                RsParty.MoveFirst
                RsParty.FIND ("code='" & Txt(Index).Tag & "'")
            End If
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    
    Select Case Index
        Case RecdFromAc, WarrAc, DiffAc
            DGridTxtKeyDown DGAccount, Txt, Index, RsParty, KeyCode, False, 1, frmSubGroup
    End Select
    If DGAccount.Visible = False Then
        '' KEY DOWN and Enter Key
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
            If Index = TDSAmt Then
                If MsgBox("Save Entry ?", vbInformation + vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave: Exit Sub
            Else
                Ctrl_DownKeyDown KeyCode, Shift
            End If
        End If
        ' KEY UP
        If TopCtrl1.TopText2 = "Add" Then
            If (Txt(VNo).Enabled = False And Index <> VDate) Or (Txt(VNo).Enabled = True And Index <> VNo) Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then
            If Index <> VDate Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)
    Select Case Index
        Case VNo
            Call NumPress(Txt(Index), keyascii, 8, 0)
        Case TDSAmt
            Call NumPress(Txt(Index), keyascii, 8, 2)
        Case RecdFromAc, WarrAc, DiffAc
            DGridTxtKeyPress Txt, Index, RsParty, keyascii, "name"
    End Select
End Sub


Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case TDSAmt
            Txt(CNoteAmt).TEXT = Format(Val(Txt(PassTotal)) + Val(Txt(TDSAmt)), "0.00")
    End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case VNo
            lblDocId = GetDocID(GCnFaS, VType, Txt(VDate).TEXT, VoucherEditFlag, Txt(VNo), lblPrefix, ForSiteCode)
            If VoucherEditFlag = True Then    ' Manual
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select v_Docid From Job_Warr2 Where v_DocID='" & lblDocId & "'", GCn, adOpenDynamic, adLockOptimistic
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                    If Txt(VNo).Enabled = True Then Txt(VNo).SetFocus
                End If
            End If
        Case RecdFromAc, WarrAc, DiffAc
            If RsParty.RecordCount > 0 And RsParty.EOF = False And RsParty.BOF = False Then
                If Txt(Index).TEXT <> "" Then
                    Txt(Index).TEXT = RsParty!Name
                    Txt(Index).Tag = RsParty!Code
                End If
            Else
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            End If
        Case VDate, CrNoteDt
            Txt(Index).TEXT = RetDate(Txt(Index))
        Case TDSAmt, ClmTax, PassTax, RejTax
            Txt(Index).TEXT = Format(Txt(Index), "0.00")
            Txt(CNoteAmt).TEXT = Format(Val(Txt(PassTotal)) + Val(Txt(TDSAmt)), "0.00")
            Amt_Calc
    End Select
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    For I = 1 To Txt.Count
        Txt(I).TEXT = ""
        Txt(I).Tag = ""
    Next I
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
    
    lblDocId.CAPTION = ""
    lblDocId.Refresh
    lblPrefix.CAPTION = ""
    lblDocId.Refresh
End Sub

Private Sub MoveRec()
Dim Rs As Recordset
Dim mVor As String
Dim I As Integer
On Error GoTo error1
    If Master.RecordCount > 0 Then
        LblDiv.CAPTION = "Division : " & DeCodeDocID(Master!v_Docid, Division_Code)
        LblSite.CAPTION = "Site Code : " & DeCodeDocID(Master!v_Docid, Current_Site)
        lblDocId.CAPTION = Master!v_Docid
        lblPrefix.CAPTION = DeCodeDocID(Master!v_Docid, Document_Prefix)
        
        Txt(RecdFromAc).Tag = XNull(Master!RecdFromcode)
        Txt(WarrAc).Tag = XNull(Master!WarrCode)
        Txt(DiffAc).Tag = XNull(Master!DiffCode)

''RecdFromName,Warr.Name as WarrName,Diff.Name as DiffName

        Txt(RecdFromAc).TEXT = XNull(Master!RecdFromName)
        Txt(WarrAc).TEXT = XNull(Master!WarrName)
        Txt(DiffAc).TEXT = XNull(Master!DiffName)

        Txt(VNo).TEXT = Master!V_NO
        Txt(VDate).TEXT = Format(Master!V_DATE, "dd/MMM/yyyy")
        
        Txt(CrNoteNo).TEXT = Master!CrNoteNo
        Txt(CrNoteDt).TEXT = Format(Master!CrNoteDate, "dd/MMM/yyyy")
        
        Txt(TDSAmt).TEXT = Format(Master!TDS_Amt, "0.00")
        
        Call Fill_Grid
    Else
        Call BlankText
    End If
    Call Amt_Calc
    
    Grid_Hide
    Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
    With FGrid1
        .RowHeightMin = PubGridRowHeight '220
        .height = PubGridRowHeight * 15
        .Cols = 30
        
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 400
        
'        .TextMatrix(0, C_ClmYear) = "Year"
'        .ColAlignment(C_ClmYear) = flexAlignLeftCenter
'        .ColAlignmentFixed(C_ClmYear) = flexAlignLeftCenter
'        .ColWidth(C_ClmYear) = 600
'
'        .TextMatrix(0, C_ClmType) = "Type"
'        .ColAlignment(C_ClmType) = flexAlignLeftCenter
'        .ColAlignmentFixed(C_ClmType) = flexAlignLeftCenter
'        .ColWidth(C_ClmType) = 600
        
        .TextMatrix(0, C_ClmNo) = "Prow.No."
        .ColAlignment(C_ClmNo) = flexAlignLeftCenter
        .ColAlignmentFixed(C_ClmNo) = flexAlignLeftCenter
        .ColWidth(C_ClmNo) = 1100
        
        .TextMatrix(0, C_ClmDate) = "Prow.Date"
        .ColAlignment(C_ClmDate) = flexAlignLeftCenter
        .ColAlignmentFixed(C_ClmDate) = flexAlignLeftCenter
        .ColWidth(C_ClmDate) = 1100
        
        .TextMatrix(0, C_PartNo) = "Part No."
        .ColAlignment(C_PartNo) = flexAlignLeftCenter
        .ColAlignmentFixed(C_PartNo) = flexAlignLeftCenter
        .ColWidth(C_PartNo) = 1100
        
        .TextMatrix(0, C_Reject) = "Reject"
        .ColAlignment(C_Reject) = flexAlignLeftCenter
        .ColAlignmentFixed(C_Reject) = flexAlignLeftCenter
        .ColWidth(C_Reject) = 700
        
        .TextMatrix(0, C_ClmLab) = "Clm.Lab."
        .ColAlignment(C_ClmLab) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmLab) = flexAlignRightCenter
        .ColWidth(C_ClmLab) = 800

        .TextMatrix(0, C_ClmSpl) = "Clm.Spl."
        .ColAlignment(C_ClmSpl) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmSpl) = flexAlignRightCenter
        .ColWidth(C_ClmSpl) = 800
        
        .TextMatrix(0, C_ClmSrv) = "Clm. Srv."
        .ColAlignment(C_ClmSrv) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmSrv) = flexAlignRightCenter
        .ColWidth(C_ClmSrv) = 800
        
        .TextMatrix(0, C_PassLab) = "Pass Lab."
        .ColAlignment(C_PassLab) = flexAlignRightCenter
        .ColAlignmentFixed(C_PassLab) = flexAlignRightCenter
        .ColWidth(C_PassLab) = 800

        .TextMatrix(0, C_PassSpl) = "Pass Spl."
        .ColAlignment(C_PassSpl) = flexAlignRightCenter
        .ColAlignmentFixed(C_PassSpl) = flexAlignRightCenter
        .ColWidth(C_PassSpl) = 800

        .TextMatrix(0, C_PassSrv) = "Pass Srv."
        .ColAlignment(C_PassSrv) = flexAlignRightCenter
        .ColAlignmentFixed(C_PassSrv) = flexAlignRightCenter
        .ColWidth(C_PassSrv) = 800

        .TextMatrix(0, C_ClmQty) = "Clm.Qty."
        .ColAlignment(C_ClmQty) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmQty) = flexAlignRightCenter
        .ColWidth(C_ClmQty) = 800
        
        .TextMatrix(0, C_ClmNDP) = "Clm. NDP"
        .ColAlignment(C_ClmNDP) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmNDP) = flexAlignRightCenter
        .ColWidth(C_ClmNDP) = 900
        
        .TextMatrix(0, C_ClmMisc) = "Clm.Misc."
        .ColAlignment(C_ClmMisc) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmMisc) = flexAlignRightCenter
        .ColWidth(C_ClmMisc) = 900
        
        .TextMatrix(0, C_ClmLST) = "Clm. LST"
        .ColAlignment(C_ClmLST) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmLST) = flexAlignRightCenter
        .ColWidth(C_ClmLST) = 900
        
        .TextMatrix(0, C_ClmSurc) = "Clm.Surc."
        .ColAlignment(C_ClmSurc) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmSurc) = flexAlignRightCenter
        .ColWidth(C_ClmSurc) = 900
        
        .TextMatrix(0, C_ClmTot) = "Clm. TOT"
        .ColAlignment(C_ClmTot) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmTot) = flexAlignRightCenter
        .ColWidth(C_ClmTot) = 900
        
        
        .TextMatrix(0, C_PassQty) = "Pass Qty."
        .ColAlignment(C_PassQty) = flexAlignRightCenter
        .ColAlignmentFixed(C_PassQty) = flexAlignRightCenter
        .ColWidth(C_PassQty) = 800
        
        .TextMatrix(0, C_PassNDP) = "Pass NDP"
        .ColAlignment(C_PassNDP) = flexAlignRightCenter
        .ColAlignmentFixed(C_PassNDP) = flexAlignRightCenter
        .ColWidth(C_PassNDP) = 900
        
        .TextMatrix(0, C_PassMisc) = "Pass Misc."
        .ColAlignment(C_PassMisc) = flexAlignRightCenter
        .ColAlignmentFixed(C_PassMisc) = flexAlignRightCenter
        .ColWidth(C_PassMisc) = 900
        
        .TextMatrix(0, C_PassLST) = "Pass LST"
        .ColAlignment(C_PassLST) = flexAlignRightCenter
        .ColAlignmentFixed(C_PassLST) = flexAlignRightCenter
        .ColWidth(C_PassLST) = 900
        
        .TextMatrix(0, C_PassSurc) = "Pass Surc."
        .ColAlignment(C_PassSurc) = flexAlignRightCenter
        .ColAlignmentFixed(C_PassSurc) = flexAlignRightCenter
        .ColWidth(C_PassSurc) = 900
        
        .TextMatrix(0, C_PassTOT) = "Pass TOT"
        .ColAlignment(C_PassTOT) = flexAlignRightCenter
        .ColAlignmentFixed(C_PassTOT) = flexAlignRightCenter
        .ColWidth(C_PassTOT) = 900

        .TextMatrix(0, C_ClmAmt) = "Total Claim"
        .ColAlignment(C_ClmAmt) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmAmt) = flexAlignRightCenter
        .ColWidth(C_ClmAmt) = 900

        .TextMatrix(0, C_PassAmt) = "Total Pass"
        .ColAlignment(C_PassAmt) = flexAlignRightCenter
        .ColAlignmentFixed(C_PassAmt) = flexAlignRightCenter
        .ColWidth(C_PassAmt) = 900

        .TextMatrix(0, C_PartName) = "Part Name"
        .ColAlignment(C_PartName) = flexAlignLeftCenter
        .ColAlignmentFixed(C_PartName) = flexAlignLeftCenter
        .ColWidth(C_PartName) = 2100

        .TextMatrix(0, C_SrlNo) = "Part Srl"
        .ColAlignment(C_SrlNo) = flexAlignLeftCenter
        .ColAlignmentFixed(C_SrlNo) = flexAlignLeftCenter
        .ColWidth(C_SrlNo) = 0

        .TextMatrix(0, C_Div) = "Div"
        .ColAlignment(C_Div) = flexAlignLeftCenter
        .ColAlignmentFixed(C_Div) = flexAlignLeftCenter
        .ColWidth(C_Div) = 0

        .TextMatrix(0, C_Site) = "Site"
        .ColAlignment(C_Site) = flexAlignLeftCenter
        .ColAlignmentFixed(C_Site) = flexAlignLeftCenter
        .ColWidth(C_Site) = 0
'        .ColWidth(C_ClmLST) = 0
'        .ColWidth(C_ClmSurc) = 0
'        .ColWidth(C_ClmTot) = 0
        .ColWidth(C_TaxYN) = 0
        'C_ClmLST,C_ClmSurc,C_ClmTot
    End With
    DGAccount.width = 5475: DGAccount.left = 6360: DGAccount.top = 705: DGAccount.height = 4755
End Sub

Public Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 1 To Txt.Count
        Txt(I).Enabled = Enb
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
    
    For I = 8 To 29
        Txt(I).Enabled = False
    Next
    
    Txt(TDSAmt).Enabled = Enb
'    Txt(ClmTax).Enabled = Enb
'    Txt(PassTax).Enabled = Enb
'    Txt(RejTax).Enabled = Enb
    txtgrid1(0).BackColor = CtrlBCol
    txtgrid1(0).ForeColor = CtrlFCol
    txtgrid1(0).Enabled = Enb
    
    Call txtDisabled_Color(Me)
End Sub

Private Sub Grid_Hide()
    If DGAccount.Visible = True Then DGAccount.Visible = False
End Sub

Private Sub UpdRequery()
    Master1.Requery
    RsPending.Requery
    RsParty.Requery
End Sub

Private Sub FGrid1_Click()
    txtgrid1(0).Visible = False
End Sub

Private Sub FGrid1_DblClick()
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid1.Col
        Case C_Reject, C_PassSrv, C_PassLab, C_PassSpl, C_PassQty, C_PassNDP, C_PassMisc, C_PassLST, C_PassSurc, C_PassTOT
            GridDblClick Me, FGrid1, txtgrid1, 0
    End Select
    TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_EnterCell()
    FGrid1.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid1_GotFocus()
    FGrid1.CellBackColor = CellBackColEnter
    txtgrid1(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
        FGrid1.CellBackColor = CellBackColLeave
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
        FGrid1.CellBackColor = CellBackColLeave
        SendKeysA vbKeyTab, True
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid1.Tag = FGrid1.Row
    Select Case FGrid1.Col
        Case C_Reject
            If KeyCode = vbKeyDelete Then
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = "No"
            End If
        Case C_PassSrv, C_PassLab, C_PassSpl, C_PassQty, C_PassNDP, C_PassMisc, C_PassLST, C_PassSurc, C_PassTOT
            If KeyCode = vbKeyDelete Then
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
            End If
    End Select
    If KeyCode = vbKeyReturn Then
        Select Case FGrid1.Col
            Case C_Reject, C_PassSrv, C_PassLab, C_PassSpl, C_PassQty, C_PassNDP, C_PassMisc, C_PassLST, C_PassSurc, C_PassTOT
                GridDblClick Me, FGrid1, txtgrid1, 0
                TAddMode = False
        End Select
        TAddMode = False
    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyPress(keyascii As Integer)
On Error GoTo ELoop
    Select Case FGrid1.Col
        Case C_Reject
            Get_Text Me, FGrid1, txtgrid1, 0, False, keyascii
        Case C_PassSrv, C_PassLab, C_PassSpl, C_PassQty, C_PassNDP, C_PassMisc, C_PassLST, C_PassSurc, C_PassTOT
            Get_Text Me, FGrid1, txtgrid1, 0, True, keyascii
        Case Else
            FGrid1_LeaveCell
            If FGrid1.Col = C_ClmMisc Then
                FGrid1.Col = FGrid1.Col + 4
                FGrid1_EnterCell
                FGrid1.SetFocus
                If keyascii <> vbKeyReturn Then TAddMode = True
                Exit Sub
            End If
            
            If FGrid1.Col = FGrid1.Cols - 1 Or FGrid1.Col > C_PartName Then
                If FGrid1.Row <> FGrid1.Rows - 1 Then
                    FGrid1.Row = FGrid1.Row + 1
                    FGrid1.Col = 1
                End If
            Else
                If FGrid1.Col < C_PartName Then
                    FGrid1.Col = FGrid1.Col + 1
                End If
            End If
            FGrid1_EnterCell
            FGrid1.SetFocus
    End Select
    If keyascii <> vbKeyReturn Then TAddMode = True
    Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid1.ColSel = False Then Exit Sub
    '' Note : Deletion Facility not Required
''    If KeyCode = vbKeyD And Shift = 2 Then
''        If FGrid1.Row  >= 1 Then
''            If MsgBox("Are You Sure To Delete Entry ?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
''                If FGrid1.Rows  > 2 Then
''                    FGrid1.RemoveItem (FGrid1.Row)
''                Else
''                    FGrid1.Rows = 1
''                    FGrid1.AddItem FGrid1.Rows
''                    FGrid1.FixedRows = 1
''                End If
''            End If
''            For i = 1 To FGrid1.Rows - 1
''                FGrid1.TextMatrix(i, 0) = i
''            Next
''        Else
''            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
''        End If
''        FGrid1.SetFocus
''    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_LeaveCell()
    FGrid1.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid1_Scroll()
    txtgrid1(0).Visible = False
    DGAccount.Visible = False
End Sub

Private Sub FGrid1_Validate(Cancel As Boolean)
    FGrid1.CellBackColor = CellBackColLeave
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
On Error GoTo ELoop
If ExitCtrl = False Then Exit Sub
    Ctrl_GetFocus txtgrid1(0)
    Grid_Hide
    FGrid1.CellBackColor = CellBackColLeave
    txtgrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    txtgrid1(0).MaxLength = 12
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        txtgrid1(0).TEXT = txtgrid1(0).Tag
        TxtGrid1_KeyUp Index, KeyCode, Shift
        txtgrid1(0).Visible = False
        FGrid1.SetFocus
        Exit Sub
    End If
    Select Case FGrid1.Col
        Case C_Reject
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave = True Then
                     GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_PassTOT
                Else
                    TxtGrid1_LostFocus 0
                    txtgrid1(0).SetFocus
                End If
            End If
        Case C_PassSrv, C_PassLab, C_PassSpl, C_PassQty, C_PassNDP, C_PassMisc, C_PassLST, C_PassSurc
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave = True Then
                     GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_PassTOT
                Else
                    TxtGrid1_LostFocus 0
                    txtgrid1(0).SetFocus
                End If
            End If
        Case C_PassTOT
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave = True Then
                     GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_PassTOT, , C_Reject
                Else
                    TxtGrid1_LostFocus 0
                    txtgrid1(0).SetFocus
                End If
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, keyascii As Integer)
On Error GoTo ELoop
    CheckQuote keyascii
    Select Case FGrid1.Col
        Case C_PassSrv, C_PassLab, C_PassSpl, C_PassQty, C_PassNDP, C_PassMisc, C_PassLST, C_PassSurc, C_PassTOT
            If FGrid1.TextMatrix(FGrid1.Row, C_Reject) = "Yes" Then
                txtgrid1(Index).TEXT = ""
                Exit Sub
            End If
            NumPress txtgrid1(Index), keyascii, 8, 2
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
    Case 0
        Select Case FGrid1.Col
            Case C_Reject
                If Len(txtgrid1(Index)) = 0 Or UCase(mID(txtgrid1(Index), 1, 1)) = "N" Then
                    txtgrid1(Index) = "No"
                ElseIf UCase(mID(txtgrid1(Index), 1, 1)) = "Y" Then
                    txtgrid1(Index) = "Yes"
                Else
                    txtgrid1(Index) = "No"
                End If
        End Select
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_LostFocus(Index As Integer)
On Error GoTo ELoop
    If ExitCtrl = False Then Exit Sub
    Ctrl_validate txtgrid1(Index)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
Dim I As Integer
On Error GoTo ELoop
    Select Case FGrid1.Col
        Case C_Reject
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = txtgrid1(Index).TEXT
            If txtgrid1(0).TEXT = "Yes" Then
                FGrid1.TextMatrix(FGrid1.Row, C_PassSrv) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_PassLab) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_PassSpl) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_PassQty) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_PassNDP) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_PassMisc) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_PassLST) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_PassSurc) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_PassTOT) = ""
            End If
        Case C_PassSrv, C_PassLab, C_PassSpl, C_PassQty, C_PassNDP, C_PassMisc, C_PassLST, C_PassSurc, C_PassTOT
            If FGrid1.TextMatrix(FGrid1.Row, C_Reject) <> "Yes" Then
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = Format(txtgrid1(Index).TEXT, "0.00")
            Else
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
            End If
    End Select
    Call Amt_Calc
NXT:
    txtgrid1(0).MaxLength = 12
    Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid1Leave() As Boolean
Dim I As Integer
    Select Case FGrid1.Col
        Case C_Reject
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = txtgrid1(0).TEXT
        Case C_PassSrv, C_PassLab, C_PassSpl, C_PassQty, C_PassNDP, C_PassMisc, C_PassLST, C_PassSurc, C_PassTOT
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = Format(txtgrid1(0).TEXT, "0.00")
    End Select
    txtgrid1(0).Visible = False
    ExitCtrl = True
    TxtGrid1Leave = True
    FGrid1.SetFocus
End Function

Private Sub Fill_Grid()
Dim I As Integer
Dim TotNDP As Double, RecTaxAmt, RecTOTAmt, TaxAmt, SurAmt, TOTAmt, LabTaxAmt As Double
    If Master1.RecordCount = 0 Then
        FGrid1.Rows = FGrid1.Rows
        FGrid1.AddItem ""
        FGrid1.FixedRows = 1
        Exit Sub
    End If
    
    Master1.MoveFirst
    FGrid1.Rows = 1
    I = 1
        Do Until Master1.EOF
nxtrec:
            If Master1.EOF = True Then GoTo Myexit
            If Master!v_Docid <> Master1!v_Docid Then
                Master1.MoveNext
                GoTo nxtrec
            End If
            Set GRs = GCn.Execute("SELECT Job_WarBill.SrvTax_Per,Job_WarBill.RecdLST_TBPer, Job_WarBill.RecdLST_TPPer, Job_WarBill.RecdTOT_TBPer, Job_WarBill.RecdTOT_TPPer, Job_WarBill.Tax_Per, Job_WarBill.Tax_Sur_Per, Job_WarBill.TOT_Per " & _
            "FROM Job_Warr1 LEFT JOIN Job_WarBill ON Job_Warr1.WBill_DocId = Job_WarBill.DocID where Job_Warr1.Div_Code = '" & Master1!Div_Code & "' and Job_Warr1.Site_Code= '" & Master1!Site_Code & "' and Job_Warr1.prowno= '" & Master1!ProwNo & "' and Job_Warr1.prowdt= '" & Master1!ProwDt & "'")
            If GRs.RecordCount > 0 Then
                TotNDP = (Master1!TotQty * Master1!Price)
'               RecdLST_TBPer,RecdTOT_TBPer,Tax_Per,Tax_Sur_Per,TOT_Per
                If Master1!Tax_YN = 1 Then 'Taxable
                    RecTaxAmt = TotNDP * GRs!RecdLST_TBPer / 100
                    RecTOTAmt = (TotNDP + RecTaxAmt) * GRs!RecdTOT_TBPer / 100
                Else
                    RecTaxAmt = TotNDP * GRs!RecdLST_TPPer / 100
                    RecTOTAmt = (TotNDP + RecTaxAmt) * GRs!RecdTOT_TPPer / 100
                End If
                TaxAmt = (TotNDP + RecTaxAmt + RecTOTAmt) * GRs!Tax_Per / 100
                SurAmt = TaxAmt * GRs!Tax_Sur_Per / 100
                TOTAmt = (TotNDP + RecTaxAmt + RecTOTAmt + TaxAmt + SurAmt) * GRs!TOT_Per / 100
                LabTaxAmt = Master1!Labour_Amt * GRs!SrvTax_Per / 100
            End If
            
            FGrid1.AddItem ""
            
            With FGrid1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, C_ClmSrv) = Format(LabTaxAmt, "0.00")
                .TextMatrix(I, C_ClmLST) = Format(RecTaxAmt + TaxAmt, "0.00")
                .TextMatrix(I, C_ClmSurc) = Format(SurAmt, "0.00")
                .TextMatrix(I, C_ClmTot) = Format(RecTOTAmt + TOTAmt, "0.00")
                .TextMatrix(I, C_ClmNo) = Master1!ProwNo
                .TextMatrix(I, C_ClmDate) = Master1!ProwDt
                .TextMatrix(I, C_PartNo) = Master1!Part_No
                .TextMatrix(I, C_Reject) = IIf(Master1!Claim_Rejected = 1, "Yes", "No")
                .TextMatrix(I, C_ClmLab) = Format(Master1!Labour_Amt, "0.00")
                .TextMatrix(I, C_ClmSpl) = Format(Master1!Spl_Amt, "0.00")
                .TextMatrix(I, C_PassLab) = Format(Master1!Labour_Pass, "0.00")
                .TextMatrix(I, C_PassSpl) = Format(Master1!Spl_Pass, "0.00")
                .TextMatrix(I, C_PassSrv) = Format(Master1!SrvTax_Pass, "0.00")
                .TextMatrix(I, C_ClmQty) = Format(Master1!TotQty, "0.00")
                .TextMatrix(I, C_ClmNDP) = Format(Master1!TotQty * Master1!Price, "0.00")
                .TextMatrix(I, C_ClmMisc) = Format(Master1!Misc_Chrg, "0.00")
                .TextMatrix(I, C_PassQty) = Format(Master1!Qty_Pass, "0.00")
                .TextMatrix(I, C_PassNDP) = Format(Master1!Spr_Pass, "0.00")
                .TextMatrix(I, C_PassMisc) = Format(Master1!Misc_pass, "0.00")
                .TextMatrix(I, C_PassLST) = Format(Master1!Lst_Pass, "0.00")
                .TextMatrix(I, C_PassSurc) = Format(Master1!Surc_Pass, "0.00")
                .TextMatrix(I, C_PassTOT) = Format(Master1!TOT_Pass, "0.00")
                .TextMatrix(I, C_ClmAmt) = Val(.TextMatrix(I, C_ClmLab)) + Val(.TextMatrix(I, C_ClmSpl)) + Val(.TextMatrix(I, C_ClmNDP)) + Val(.TextMatrix(I, C_ClmMisc)) + Val(.TextMatrix(I, C_ClmLST)) + Val(.TextMatrix(I, C_ClmSurc)) + Val(.TextMatrix(I, C_ClmTot)) + Val(.TextMatrix(I, C_ClmSrv))
                .TextMatrix(I, C_PassAmt) = Val(.TextMatrix(I, C_PassLab)) + Val(.TextMatrix(I, C_PassSpl)) + Val(.TextMatrix(I, C_PassNDP)) + Val(.TextMatrix(I, C_PassMisc)) + Val(.TextMatrix(I, C_PassLST)) + Val(.TextMatrix(I, C_PassSurc)) + Val(.TextMatrix(I, C_PassTOT)) + Val(.TextMatrix(I, C_PassSrv))
                .TextMatrix(I, C_PartName) = XNull(Master1!Part_Name)
                .TextMatrix(I, C_SrlNo) = XNull(Master1!Srlno)
                .TextMatrix(I, C_Div) = XNull(Master1!Div_Code)
                .TextMatrix(I, C_Site) = XNull(Master1!Site_Code)
            End With
            I = I + 1
            Master1.MoveNext
            Set GRs = Nothing
        Loop
Myexit:
        FGrid1.FixedRows = 1
End Sub

Private Sub Amt_Calc()
Dim Mytot As Double
Dim I As Integer
Dim MyClmLab As Double, MyPassLab As Double, MyRejLab As Double
Dim MyClmNDP As Double, MyPassNDP As Double, MyRejNDP As Double
Dim MyClmMisc As Double, MyPassMisc As Double, MyRejMisc As Double
Dim MyClmTax As Double, MyPassTax As Double, MyRejTax As Double
Dim MyClmTotal As Double, MyPassTotal As Double, MyRejTotal As Double
        
    MyClmLab = 0: MyPassLab = 0: MyRejLab = 0
    MyClmNDP = 0: MyPassNDP = 0: MyRejNDP = 0
    MyClmMisc = 0: MyPassMisc = 0: MyRejMisc = 0
    MyClmTax = 0: MyPassTax = 0: MyRejTax = 0
    MyClmTotal = 0: MyPassTotal = 0: MyRejTotal = 0
        
    For I = 1 To FGrid1.Rows - 1
        FGrid1.TextMatrix(I, C_PassAmt) = Format(Val(FGrid1.TextMatrix(I, C_PassLab)) + Val(FGrid1.TextMatrix(I, C_PassSpl)) + Val(FGrid1.TextMatrix(I, C_PassNDP)) + Val(FGrid1.TextMatrix(I, C_PassMisc)) + Val(FGrid1.TextMatrix(I, C_PassLST)) + Val(FGrid1.TextMatrix(I, C_PassSurc)) + Val(FGrid1.TextMatrix(I, C_PassTOT)), "0.00")
        
        If FGrid1.TextMatrix(I, C_Reject) = "Yes" Then
            MyRejLab = MyRejLab + Val(FGrid1.TextMatrix(I, C_ClmLab)) + Val(FGrid1.TextMatrix(I, C_ClmSpl))
            MyRejNDP = MyRejNDP + Val(FGrid1.TextMatrix(I, C_ClmNDP))
            MyRejMisc = MyRejMisc + Val(FGrid1.TextMatrix(I, C_ClmMisc))
            MyRejTax = MyRejTax + Val(FGrid1.TextMatrix(I, C_ClmLST)) + Val(FGrid1.TextMatrix(I, C_ClmSurc)) + Val(FGrid1.TextMatrix(I, C_ClmTot)) + Val(FGrid1.TextMatrix(I, C_ClmSrv))
            MyRejTotal = MyRejLab + MyRejNDP + MyRejMisc + MyRejTax
        ElseIf Val(FGrid1.TextMatrix(I, C_PassLab)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassSpl)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassNDP)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassMisc)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassLST)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassSurc)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassTOT)) > 0 Or Val(FGrid1.TextMatrix(I, C_PassSrv)) > 0 Then
            MyClmLab = MyClmLab + Val(FGrid1.TextMatrix(I, C_ClmLab)) + Val(FGrid1.TextMatrix(I, C_ClmSpl))
            MyClmNDP = MyClmNDP + Val(FGrid1.TextMatrix(I, C_ClmNDP))
            MyClmMisc = MyClmMisc + Val(FGrid1.TextMatrix(I, C_ClmMisc))
            MyClmTax = MyClmTax + Val(FGrid1.TextMatrix(I, C_ClmLST)) + Val(FGrid1.TextMatrix(I, C_ClmSurc)) + Val(FGrid1.TextMatrix(I, C_ClmTot)) + Val(FGrid1.TextMatrix(I, C_ClmSrv))
            MyClmTotal = MyClmLab + MyClmNDP + MyClmMisc + MyClmTax
            
            MyPassLab = MyPassLab + Val(FGrid1.TextMatrix(I, C_PassLab)) + Val(FGrid1.TextMatrix(I, C_PassSpl))
            MyPassNDP = MyPassNDP + Val(FGrid1.TextMatrix(I, C_PassNDP))
            MyPassMisc = MyPassMisc + Val(FGrid1.TextMatrix(I, C_PassMisc))
            MyPassTax = MyPassTax + Val(FGrid1.TextMatrix(I, C_PassLST)) + Val(FGrid1.TextMatrix(I, C_PassSurc)) + Val(FGrid1.TextMatrix(I, C_PassTOT)) + Val(FGrid1.TextMatrix(I, C_PassSrv))
            MyPassTotal = MyPassLab + MyPassNDP + MyPassMisc + MyPassTax
        End If
    Next I
    
'    MyPassTotal = MyPassTotal + Val(Txt(PassTax))
'    MyClmTotal = MyClmTotal + Val(Txt(ClmTax))
'    MyRejTotal = MyRejTotal + Val(Txt(RejTax))
    
    
    Txt(RejLabour) = Format(MyRejLab, "0.00")
    Txt(RejNDP) = Format(MyRejNDP, "0.00")
    Txt(RejMisc) = Format(MyRejMisc, "0.00")
    Txt(RejTax) = Format(MyRejTax, "0.00")
    Txt(RejTotal) = Format(MyRejTotal, "0.00")

    Txt(ClmLabour) = Format(MyClmLab, "0.00")
    Txt(ClmNDP) = Format(MyClmNDP, "0.00")
    Txt(ClmMisc) = Format(MyClmMisc, "0.00")
    Txt(ClmTax) = Format(MyClmTax, "0.00")
    Txt(ClmTotal) = Format(MyClmTotal, "0.00")

    Txt(PassLabour) = Format(MyPassLab, "0.00")
    Txt(PassNDP) = Format(MyPassNDP, "0.00")
    Txt(PassMisc) = Format(MyPassMisc, "0.00")
    Txt(PassTax) = Format(MyPassTax, "0.00")
    Txt(PassTotal) = Format(MyPassTotal, "0.00")

    Txt(DiffLabour) = Format(MyClmLab - MyPassLab, "0.00")
    Txt(DiffNDP) = Format(MyClmNDP - MyPassNDP, "0.00")
    Txt(DiffMisc) = Format(MyClmMisc - MyPassMisc, "0.00")
    Txt(DiffTax) = Format(Val(Txt(ClmTax)) - Val(Txt(PassTax)), "0.00")
    Txt(DiffTotal) = Format(MyClmTotal - MyPassTotal, "0.00")

    Txt(CNoteAmt).TEXT = Format(MyPassTotal + Val(Txt(TDSAmt)), "0.00")
End Sub

Private Sub Fill_GridNew()
Dim I As Integer
Dim TotNDP As Double, RecTaxAmt, RecTOTAmt, TaxAmt, SurAmt, TOTAmt, LabTaxAmt As Double
    If TopCtrl1.TopText2 = "Add" Then
        If RsPending.RecordCount = 0 Then
            FGrid1.Rows = 1
            FGrid1.AddItem FGrid1.Rows
            FGrid1.FixedRows = 1
            Exit Sub
        End If
        FGrid1.Rows = 1
        I = 1
    Else    '' Edit Mode
        '' Only Add New Rows for Pending Claim
        If RsPending.RecordCount = 0 Then Exit Sub
        I = FGrid1.Rows
    End If
    RsPending.MoveFirst
        Do Until RsPending.EOF
            Set GRs = GCn.Execute("SELECT Job_WarBill.SrvTax_Per,Job_WarBill.RecdLST_TBPer, Job_WarBill.RecdLST_TPPer, Job_WarBill.RecdTOT_TBPer, Job_WarBill.RecdTOT_TPPer, Job_WarBill.Tax_Per, Job_WarBill.Tax_Sur_Per, Job_WarBill.TOT_Per " & _
            "FROM Job_Warr1 LEFT JOIN Job_WarBill ON Job_Warr1.WBill_DocId = Job_WarBill.DocID where left(Job_Warr1.WBill_DocId,1) = '" & RsPending!Div_Code & "' and Job_Warr1.Site_Code= '" & RsPending!Site_Code & "' and Job_Warr1.Prowno= '" & RsPending!ProwNo & "' and Job_Warr1.Prowdt= '" & RsPending!ProwDt & "' and Job_Warr1.WBill_DocId <> ''")
            If GRs.RecordCount > 0 Then
                TotNDP = (RsPending!TotQty * RsPending!Price)
'               RecdLST_TBPer,RecdTOT_TBPer,Tax_Per,Tax_Sur_Per,TOT_Per
                If RsPending!Tax_YN = 1 Then 'Taxable
                    RecTaxAmt = TotNDP * GRs!RecdLST_TBPer / 100
                    RecTOTAmt = (TotNDP + RecTaxAmt) * GRs!RecdTOT_TBPer / 100
                Else
                    RecTaxAmt = TotNDP * GRs!RecdLST_TPPer / 100
                    RecTOTAmt = (TotNDP + RecTaxAmt) * GRs!RecdTOT_TPPer / 100
                End If
                TaxAmt = (TotNDP + RecTaxAmt + RecTOTAmt) * GRs!Tax_Per / 100
                SurAmt = TaxAmt * GRs!Tax_Sur_Per / 100
                TOTAmt = (TotNDP + RecTaxAmt + RecTOTAmt + TaxAmt + SurAmt) * GRs!TOT_Per / 100
                LabTaxAmt = RsPending!Labour_Amt * GRs!SrvTax_Per / 100
            End If
        
            FGrid1.AddItem ""
            With FGrid1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, C_ClmSrv) = Format(LabTaxAmt, "0.00")
                .TextMatrix(I, C_ClmLST) = Format(RecTaxAmt + TaxAmt, "0.00")
                .TextMatrix(I, C_ClmSurc) = Format(SurAmt, "0.00")
                .TextMatrix(I, C_ClmTot) = Format(RecTOTAmt + TOTAmt, "0.00")
                .TextMatrix(I, C_ClmNo) = RsPending!ProwNo
                .TextMatrix(I, C_ClmDate) = RsPending!ProwDt
                .TextMatrix(I, C_PartNo) = RsPending!Part_No
                .TextMatrix(I, C_Reject) = IIf(RsPending!Claim_Rejected = 1, "Yes", "No")
                .TextMatrix(I, C_ClmLab) = Format(RsPending!Labour_Amt, "0.00")
                .TextMatrix(I, C_ClmSpl) = Format(RsPending!Spl_Amt, "0.00")
                .TextMatrix(I, C_PassLab) = Format(RsPending!Labour_Pass, "0.00")
                .TextMatrix(I, C_PassSpl) = Format(RsPending!Spl_Pass, "0.00")
                .TextMatrix(I, C_PassSrv) = Format(RsPending!SrvTax_Pass, "0.00")
                .TextMatrix(I, C_ClmQty) = Format(RsPending!TotQty, "0.00")
                .TextMatrix(I, C_ClmNDP) = Format(RsPending!TotQty * RsPending!Price, "0.00")
                .TextMatrix(I, C_ClmMisc) = Format(RsPending!Misc_Chrg, "0.00")
                .TextMatrix(I, C_PassQty) = Format(RsPending!Qty_Pass, "0.00")
                .TextMatrix(I, C_PassNDP) = Format(RsPending!Spr_Pass, "0.00")
                .TextMatrix(I, C_PassMisc) = Format(RsPending!Misc_pass, "0.00")
                .TextMatrix(I, C_PassLST) = Format(RsPending!Lst_Pass, "0.00")
                .TextMatrix(I, C_PassSurc) = Format(RsPending!Surc_Pass, "0.00")
                .TextMatrix(I, C_PassTOT) = Format(RsPending!TOT_Pass, "0.00")
                .TextMatrix(I, C_ClmAmt) = Format(Val(.TextMatrix(I, C_ClmLab)) + Val(.TextMatrix(I, C_ClmSpl)) + Val(.TextMatrix(I, C_ClmNDP)) + Val(.TextMatrix(I, C_ClmMisc)) + Val(.TextMatrix(I, C_ClmLST)) + Val(.TextMatrix(I, C_ClmSurc)) + Val(.TextMatrix(I, C_ClmTot)), "0.00")
                .TextMatrix(I, C_PassAmt) = Format(Val(.TextMatrix(I, C_PassLab)) + Val(.TextMatrix(I, C_PassSpl)) + Val(.TextMatrix(I, C_PassNDP)) + Val(.TextMatrix(I, C_PassMisc)) + Val(.TextMatrix(I, C_PassLST)) + Val(.TextMatrix(I, C_PassSurc)) + Val(.TextMatrix(I, C_PassTOT)), "0.00")
                .TextMatrix(I, C_PartName) = XNull(RsPending!Part_Name)
                .TextMatrix(I, C_TaxYN) = IIf(RsPending!Tax_YN = 0, "No", "Yes")
                .TextMatrix(I, C_SrlNo) = RsPending!Srlno
                .TextMatrix(I, C_Div) = RsPending!Div_Code
                .TextMatrix(I, C_Site) = RsPending!Site_Code
            End With
            I = I + 1
            RsPending.MoveNext
            Set GRs = Nothing
        Loop
        FGrid1.FixedRows = 1
        Amt_Calc
End Sub

