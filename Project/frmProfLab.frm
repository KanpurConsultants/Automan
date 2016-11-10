VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmProfLab 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Performa Labour Entry"
   ClientHeight    =   7230
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
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DGHist 
      Height          =   2520
      Left            =   225
      Negotiate       =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6810
      Visible         =   0   'False
      Width           =   11700
      _ExtentX        =   20638
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "RegNo"
         Caption         =   "Reg. No."
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Name"
         Caption         =   "Owner Name"
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
         DataField       =   "PhoneOff"
         Caption         =   "Phone (O)"
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
         DataField       =   "Govt"
         Caption         =   "Govt"
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
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4004.788
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   599.811
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGLabour 
      Height          =   2730
      Left            =   8445
      Negotiate       =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   3570
      Visible         =   0   'False
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   4815
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
      ColumnCount     =   3
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
         Caption         =   "Labour Description"
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
         DataField       =   "LabGrp_Desc"
         Caption         =   "Group"
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
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4320
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   23
      Left            =   7200
      TabIndex        =   69
      Top             =   1155
      Visible         =   0   'False
      Width           =   2100
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
      Index           =   17
      Left            =   1485
      MaxLength       =   5
      TabIndex        =   20
      Text            =   "99.99"
      Top             =   6195
      Width           =   600
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
      Index           =   22
      Left            =   5565
      MaxLength       =   25
      TabIndex        =   2
      Text            =   "28-APR-2002"
      Top             =   435
      Width           =   1290
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   21
      Left            =   2400
      MaxLength       =   8
      TabIndex        =   1
      Top             =   435
      Width           =   780
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   9690
      MaxLength       =   4
      TabIndex        =   24
      Top             =   1485
      Visible         =   0   'False
      Width           =   615
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
      Left            =   5565
      MaxLength       =   25
      TabIndex        =   4
      Top             =   705
      Width           =   1290
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
      Left            =   9000
      MaxLength       =   40
      TabIndex        =   17
      Top             =   4140
      Visible         =   0   'False
      Width           =   705
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
      Index           =   13
      Left            =   1755
      MaxLength       =   25
      TabIndex        =   14
      Top             =   2865
      Width           =   5100
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
      Left            =   1755
      MaxLength       =   40
      TabIndex        =   10
      Top             =   1785
      Width           =   5100
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
      Left            =   1755
      MaxLength       =   40
      TabIndex        =   11
      Top             =   2055
      Width           =   5100
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
      Index           =   11
      Left            =   1755
      MaxLength       =   40
      TabIndex        =   12
      Top             =   2325
      Width           =   5100
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
      Left            =   1755
      MaxLength       =   40
      TabIndex        =   13
      Top             =   2595
      Visible         =   0   'False
      Width           =   5100
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
      Left            =   1755
      MaxLength       =   14
      TabIndex        =   6
      Top             =   1245
      Width           =   1425
   End
   Begin VB.TextBox txt 
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
      Index           =   2
      Left            =   1755
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "Help"
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
      Index           =   4
      Left            =   5565
      MaxLength       =   25
      TabIndex        =   5
      Top             =   975
      Width           =   1290
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
      Index           =   15
      Left            =   2340
      MaxLength       =   10
      TabIndex        =   18
      Text            =   "999999.99"
      Top             =   5655
      Width           =   1110
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
      Index           =   19
      Left            =   2340
      MaxLength       =   10
      TabIndex        =   22
      Top             =   6465
      Width           =   1110
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
      Index           =   18
      Left            =   2340
      MaxLength       =   10
      TabIndex        =   21
      Top             =   6195
      Width           =   1110
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
      Index           =   20
      Left            =   2340
      MaxLength       =   10
      TabIndex        =   23
      Top             =   6735
      Width           =   1110
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
      Index           =   6
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1245
      Width           =   2055
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
      Index           =   8
      Left            =   4800
      MaxLength       =   25
      TabIndex        =   9
      Top             =   1515
      Width           =   2055
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
      Index           =   16
      Left            =   2340
      MaxLength       =   10
      TabIndex        =   19
      Top             =   5925
      Width           =   1110
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
      Index           =   14
      Left            =   1755
      MaxLength       =   25
      TabIndex        =   15
      Top             =   3135
      Width           =   5100
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
      Left            =   1755
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1515
      Width           =   1875
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   2085
      Left            =   135
      TabIndex        =   16
      Top             =   3465
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   3678
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   6
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   16777215
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      GridColorFixed  =   8421504
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   1
      Appearance      =   0
      GridLineWidthFixed=   1
      FormatString    =   "KKKK"
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
   Begin MSDataGridLib.DataGrid DGJob 
      Height          =   2520
      Left            =   3765
      Negotiate       =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   6465
      Visible         =   0   'False
      Width           =   11820
      _ExtentX        =   20849
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "FindJobNo"
         Caption         =   "Job No."
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
      BeginProperty Column02 
         DataField       =   "RegNo"
         Caption         =   "Reg. No"
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
         DataField       =   "VehSerialNo"
         Caption         =   "Veh.Srl No."
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
         DataField       =   "Name"
         Caption         =   "Owner Name"
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
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   3
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3195.213
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   2730
      Left            =   4275
      Negotiate       =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   5715
      Visible         =   0   'False
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   4815
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
      RowHeight       =   19
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "name"
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin VB.Label lblVPrefix 
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1755
      TabIndex        =   67
      Top             =   450
      Width           =   720
   End
   Begin VB.Label lblDocId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Performa Doc Id"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   225
      Left            =   8790
      TabIndex        =   66
      Top             =   795
      Width           =   1380
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
      Index           =   7
      Left            =   1650
      TabIndex        =   65
      Top             =   3135
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Performa Dt."
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
      Index           =   2
      Left            =   4365
      TabIndex        =   64
      Top             =   450
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   8
      Left            =   5460
      TabIndex        =   63
      Top             =   450
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   0
      Left            =   1650
      TabIndex        =   62
      Top             =   450
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Performa No."
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
      Index           =   0
      Left            =   165
      TabIndex        =   61
      Top             =   450
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
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   29
      Left            =   9585
      TabIndex        =   60
      Top             =   1485
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Against JobCard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   13
      Left            =   8070
      TabIndex        =   59
      Top             =   1500
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JC Open Dt."
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
      Left            =   4365
      TabIndex        =   58
      Top             =   720
      Width           =   990
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
      Height          =   225
      Index           =   3
      Left            =   5460
      TabIndex        =   57
      Top             =   720
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Index           =   26
      Left            =   165
      TabIndex        =   56
      Top             =   2055
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
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
      Index           =   31
      Left            =   165
      TabIndex        =   55
      Top             =   3135
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
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
      Index           =   39
      Left            =   165
      TabIndex        =   54
      Top             =   1785
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Tax"
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
      Left            =   300
      TabIndex        =   53
      Top             =   6210
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Index           =   40
      Left            =   300
      TabIndex        =   52
      Top             =   5670
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Index           =   10
      Left            =   165
      TabIndex        =   51
      Top             =   2865
      Width           =   300
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division            :"
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
      TabIndex        =   50
      Top             =   525
      Width           =   1245
   End
   Begin VB.Label lblDocCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Performa DocID :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   225
      Left            =   7215
      TabIndex        =   49
      Top             =   795
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
      Height          =   255
      Index           =   4
      Left            =   1650
      TabIndex        =   48
      Top             =   2865
      Width           =   45
   End
   Begin VB.Line Line1 
      X1              =   150
      X2              =   11745
      Y1              =   3420
      Y2              =   3420
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
      Height          =   225
      Index           =   34
      Left            =   5460
      TabIndex        =   47
      Top             =   990
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Dt."
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
      Left            =   4620
      TabIndex        =   46
      Top             =   990
      Width           =   765
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
      Index           =   28
      Left            =   1650
      TabIndex        =   45
      Top             =   720
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
      Height          =   225
      Index           =   27
      Left            =   2250
      TabIndex        =   44
      Top             =   5670
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
      Height          =   255
      Index           =   24
      Left            =   2250
      TabIndex        =   43
      Top             =   6465
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rounded off"
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
      Left            =   300
      TabIndex        =   42
      Top             =   6480
      Width           =   1005
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
      Height          =   270
      Index           =   22
      Left            =   2250
      TabIndex        =   41
      Top             =   6187
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   19
      Left            =   2250
      TabIndex        =   40
      Top             =   6735
      Width           =   45
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
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
      Index           =   20
      Left            =   300
      TabIndex        =   39
      Top             =   6750
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard No."
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
      Index           =   12
      Left            =   165
      TabIndex        =   38
      Top             =   720
      Width           =   1035
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
      Index           =   1
      Left            =   4680
      TabIndex        =   37
      Top             =   1245
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No."
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
      Index           =   8
      Left            =   3630
      TabIndex        =   36
      Top             =   1245
      Width           =   1035
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
      Index           =   90
      Left            =   1650
      TabIndex        =   35
      Top             =   1245
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration No."
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
      Index           =   3
      Left            =   165
      TabIndex        =   34
      Top             =   1245
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   660
      Left            =   7035
      Top             =   450
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
      Top             =   525
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Amount"
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
      Index           =   35
      Left            =   300
      TabIndex        =   32
      Top             =   5940
      Width           =   1410
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
      Index           =   17
      Left            =   2250
      TabIndex        =   31
      Top             =   5925
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
      Height          =   255
      Index           =   18
      Left            =   1650
      TabIndex        =   30
      Top             =   1515
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
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
      Index           =   38
      Left            =   165
      TabIndex        =   29
      Top             =   1515
      Width           =   495
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
      Index           =   26
      Left            =   1650
      TabIndex        =   28
      Top             =   1785
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
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   27
      Top             =   1515
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No."
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
      Index           =   33
      Left            =   3750
      TabIndex        =   26
      Top             =   1515
      Width           =   915
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
      Index           =   6
      Left            =   1650
      TabIndex        =   25
      Top             =   2055
      Width           =   45
   End
End
Attribute VB_Name = "frmProfLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BackColorSelEnter As String = &HF8D7FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Dim TAddMode As Boolean
Dim GridKey As Integer

Dim VoucherEditFlag As Boolean
Dim ProfDocId As String
Dim ForSiteCode As String

Private Const VType As String = "W_PL"

Dim MyIndex As Byte
Dim MyCardNo As String
Dim Rst As ADODB.Recordset

Dim Master As ADODB.Recordset
Dim RsJob As ADODB.Recordset
Dim RsLab As ADODB.Recordset
Dim RsCity As ADODB.Recordset
Dim RsHist As ADODB.Recordset

'Text Box (Form)
Private Const JCType As Byte = 1
Private Const JobNo As Byte = 2
Private Const JobDt As Byte = 3
Private Const JobCDt As Byte = 4
Private Const VehRegNo As Byte = 5
Private Const Chassis As Byte = 6
Private Const Model As Byte = 7
Private Const Engine As Byte = 8
Private Const Party As Byte = 9
Private Const Address1 As Byte = 10
Private Const Address2 As Byte = 11
Private Const Address3 As Byte = 12
Private Const City As Byte = 13
Private Const PhoneOff As Byte = 14
Private Const TOTAmt As Byte = 15
Private Const DiscAmt As Byte = 16
Private Const SrvPerc As Byte = 17
Private Const SrvAmt As Byte = 18
Private Const Roundedoff As Byte = 19
Private Const NetAmt As Byte = 20
Private Const ProfNo As Byte = 21
Private Const ProfDt As Byte = 22
Private Const TxtDocID As Byte = 23

'Text Box (Grid)
Private Const mTxtGrid1 As Byte = 1

'Fgrid1 Columns
Private Const C_LabCode As Byte = 1
Private Const C_LabName As Byte = 2
Private Const C_ChgHrs As Byte = 3
Private Const C_ChgAmt As Byte = 4
Private Const C_Remarks As Byte = 5

Private Sub DGHist_Click()
If RsHist.RecordCount > 0 Then
    txt(VehRegNo).TEXT = XNull(RsHist!RegNo)
    txt(Model) = XNull(RsHist!Model)
    txt(Chassis).TEXT = XNull(RsHist!Chassis)
    txt(Engine).TEXT = XNull(RsHist!Engine)
    txt(Party).TEXT = XNull(RsHist!Name)
    txt(Address1).TEXT = XNull(RsHist!Add1)
    txt(Address2).TEXT = XNull(RsHist!Add2)
    txt(City).Tag = XNull(RsHist!CityCode)
    txt(City).TEXT = XNull(RsHist!CityName)
    txt(PhoneOff).TEXT = XNull(RsHist!PhoneOff)
End If
DGHist.Visible = False
txt(VehRegNo).SetFocus
End Sub

Private Sub DGJob_Click()
If Master.RecordCount > 0 Then
    Call History_Field
End If
txt(MyIndex).SetFocus
DGJob.Visible = False
End Sub

Private Sub DGLabour_Click()
If RsLab.RecordCount > 0 Then
    txtgrid1(0).Tag = RsLab!Code
    txtgrid1(0).TEXT = RsLab!Name
End If
txtgrid1(0).SetFocus
DGLabour.Visible = False
End Sub

Private Sub DGCity_Click()
If RsCity.RecordCount > 0 Then
    txt(MyIndex).Tag = RsCity!Code
    txt(MyIndex).TEXT = RsCity!Name
End If
txt(MyIndex).SetFocus
DgCity.Visible = False
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
    
    WinSetting Me:    Ini_Grid
    TopCtrl1.Tag = PubUParam: ForSiteCode = PubSiteCode
    txt(ProfDt).Tag = PubLoginDate
    Call BlankText
    LblVPrefix.CAPTION = ""
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
     Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and  " & cMID("es.Docid", "3", "1") & "='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If



    If PubMoveRecYn Then
        Master.Open "select ES.docId as SearchCode,ES.docId,ES.JOB_DOCID,ES.SITE_CODE,ES.V_NO,ES.V_DATE,ES.STORES_WORKS,ES.PARTY_NAME,ES.ADDRESS,ES.ADDRESS2,ES.ADDRESS3,ES.PHONENO,ES.CARDNO,ES.CITYCODE,ES.LAB_AMT,ES.LAB_D_AMT,ES.LAB_TAXPER,ES.LAB_TAXAMT,ES.Lab_Rounded,ES.Lab_Total_Amt,JC.Job_No,JC.Job_Date,JC.JobCloseDate,ES.Model,ES.RegNo,ES.Chassis,ES.Engine,City.CityName " & _
                    "from ((ESTIMATE as ES left Join job_card as JC on ES.Job_DocId=JC.DocId) " & _
                    "left Join Hiscard as HC on ES.CardNo=HC.CardNo) " & _
                    "Left Join City on ES.CityCode=City.CityCode " & _
                    "WHERE LEFT(ES.DocId,1)='" & PubDivCode & "' " & sitecond & " AND ES.V_TYPE= '" & VType & "' order by ES.V_Date desc,ES.V_Type,ES.DocID desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 ES.docId as SearchCode,ES.docId,ES.JOB_DOCID,ES.SITE_CODE,ES.V_NO,ES.V_DATE,ES.STORES_WORKS,ES.PARTY_NAME,ES.ADDRESS,ES.ADDRESS2,ES.ADDRESS3,ES.PHONENO,ES.CARDNO,ES.CITYCODE,ES.LAB_AMT,ES.LAB_D_AMT,ES.LAB_TAXPER,ES.LAB_TAXAMT,ES.Lab_Rounded,ES.Lab_Total_Amt,JC.Job_No,JC.Job_Date,JC.JobCloseDate,ES.Model,ES.RegNo,ES.Chassis,ES.Engine,City.CityName " & _
                    "from ((ESTIMATE as ES left Join job_card as JC on ES.Job_DocId=JC.DocId) " & _
                    "left Join Hiscard as HC on ES.CardNo=HC.CardNo) " & _
                    "Left Join City on ES.CityCode=City.CityCode " & _
                    "WHERE LEFT(ES.DocId,1)='" & PubDivCode & "' " & sitecond & " AND ES.V_TYPE= '" & VType & "' order by ES.V_Date desc,ES.V_Type,ES.DocID desc", GCn, adOpenDynamic, adLockOptimistic
    End If
    
     If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and  " & cMID("j.Docid", "3", "1") & "='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    Set RsJob = New ADODB.Recordset
    With RsJob
        .CursorLocation = adUseClient
        .Open "select J.DocId AS CODE," & cCStr("J.Job_No") & " As FindJobNo,J.Job_No,HC.Model,HC.RegNo, HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name, J.DocId,J.CARDNO,J.Govt_YN, J.Job_Date, J.JobCloseDate, HC.Add1, HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc,HC.CityCode,City.CityName from ((job_card as J left Join Hiscard as HC on J.CardNo=HC.CardNo) left Join Service_Type as ST on J.Serv_Type=ST.Serv_Type) Left Join City on HC.CityCode=City.CityCode where left(J.DocId,1)='" & PubDivCode & "' " & sitecond & " and (right(j.DocId_InvSpr,8) <> 'Cancelld' Or J.DocId_InvSpr Is Null ) order by J.Job_No", GCn, adOpenDynamic, adLockOptimistic
    End With
'    RsJob.Sort = "Job_No"
    Set DGJob.DataSource = RsJob

    Set RsLab = New ADODB.Recordset
    RsLab.CursorLocation = adUseClient
    RsLab.Open "Select L.Lab_Code as code,L.Lab_Desc as name,LG.LabGrp_Desc,L.Lab_Rate,L.Time_Req " & _
        "FROM Labour  L left join Labour_Group LG on L.Lab_Group=LG.Lab_Group " & _
        "Order by Lab_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGLabour.DataSource = RsLab
    RsLab.Sort = "code"
    RsLab.Sort = "name"
    
    Set RsCity = New ADODB.Recordset
    RsCity.CursorLocation = adUseClient
    RsCity.Open "Select CityCode as code,CityName as name FROM City Order by CityName", GCn, adOpenDynamic, adLockOptimistic
    Set DgCity.DataSource = RsCity
    RsCity.Sort = "Name"
    'modi lps
    Set RsHist = New ADODB.Recordset
    RsHist.CursorLocation = adUseClient
    'Modify SQL for speed
    RsHist.Open "Select RegNo as Code,Chassis,RegNo,Model,Name,Engine,GOVT_YN," & _
            " " & cIIF("GOVT_YN=0", "'No'", "'Yes'") & " as Govt, Add1,Add2,PhoneOff,PhoneResi,Mobile," & _
            " VehSerialNo,HISCARD.CityCode,City.CityName " & _
            " FROM (Hiscard " & _
            " left join city on Hiscard.CityCode=City.CityCode) " & _
            " Where HISCARD.Div_Code='" & PubDivCode & "' Order by Regno", GCn, adOpenDynamic, adLockOptimistic
    Set DGHist.DataSource = RsHist
    RsHist.Sort = "Code"
    'eof
    
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

Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RsJob = Nothing
    Set RsLab = Nothing
    Set RsCity = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
     
    txt(ProfDt).TEXT = txt(ProfDt).Tag 'Format(Date, "dd/MMM/yyyy")
    txt(ProfNo).TEXT = Format(GCn.Execute("select " & vIsNull("max(V_no)", "0") & "+1 from ESTIMATE where left(docid,1)='" & PubDivCode & "' and " & cMID("docid", "2", "2") & "='" & PubSiteCode + ForSiteCode & "' AND V_TYPE='" & VType & "'").Fields(0), "000000")
    
    txt(TxtDocID) = GetDocID(GCnFaW, VType, txt(ProfDt).TEXT, VoucherEditFlag, txt(ProfNo), LblVPrefix, ForSiteCode)
    ProfDocId = txt(TxtDocID)
    lblDocId.CAPTION = "DocId : " & ProfDocId
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    If VoucherEditFlag Then
        txt(ProfNo).Enabled = True
        txt(ProfNo).SetFocus
    Else
        txt(ProfDt).SetFocus
    End If
    Exit Sub
'    Txt(JCType).Text = "No"
'    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim mTrans As Boolean
    If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
        mTrans = True
        GCn.Execute "Delete from Estimate1  where Docid='" & ProfDocId & "'"
        GCn.Execute "Delete from Estimate  where Docid='" & ProfDocId & "'"
    
        GCn.CommitTrans
        mTrans = False
        Master.Requery
        Call UpdRequery
        
        If Master.RecordCount > 0 Then
            Call MoveRec
        Else
            Call BlankText
        End If
        BUTTONS True, Me, Master, 0
    End If
    Exit Sub
eloop1:
    If mTrans Then GCn.RollbackTrans
    MsgBox err.Description, vbCritical, " Deletion Message"
End Sub

Private Sub TopCtrl1_eEdit()
Dim I As Integer
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    txt(ProfDt).SetFocus
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
  Dim sitecond As String
 
    
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
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("searchcode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("select ES.docId as SearchCode,ES.docId,ES.JOB_DOCID,ES.SITE_CODE,ES.V_NO,ES.V_DATE,ES.STORES_WORKS,ES.PARTY_NAME,ES.ADDRESS,ES.ADDRESS2,ES.ADDRESS3,ES.PHONENO,ES.CARDNO,ES.CITYCODE,ES.LAB_AMT,ES.LAB_D_AMT,ES.LAB_TAXPER,ES.LAB_TAXAMT,ES.Lab_Rounded,ES.Lab_Total_Amt,JC.Job_No,JC.Job_Date,JC.JobCloseDate,ES.Model,ES.RegNo,ES.Chassis,ES.Engine,City.CityName " & _
                    "from ((ESTIMATE as ES left Join job_card as JC on ES.Job_DocId=JC.DocId) " & _
                    "left Join Hiscard as HC on ES.CardNo=HC.CardNo) " & _
                    "Left Join City on ES.CityCode=City.CityCode " & _
                    "WHERE LEFT(ES.DocId,1)='" & PubDivCode & "' AND ES.V_TYPE= '" & VType & "' And ES.docId = '" & MyValue & "' order by ES.V_Date desc,ES.V_Type,ES.DocID desc")
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

Private Sub TopCtrl1_ePrn()
Dim mQry$, RepTitle$
Dim Condstr$, FormulaStr1$, FormulaStr2$, Speciality$
Dim Rst As ADODB.Recordset, RST1 As ADODB.Recordset
Dim I As Integer
Dim SecondStr As Boolean

On Error GoTo ERRORHANDLER
mQry = "select E.V_NO,E.V_DATE, " & _
    "E.PARTY_NAME,E.ADDRESS,E.ADDRESS2,E.ADDRESS3, " & _
    "E.PHONENO,E.CARDNO,E.CITYCODE,E.LAB_AMT, " & _
    "E.LAB_D_AMT,E.LAB_TAXPER,E.LAB_TAXAMT," & _
    "E.Lab_Rounded,E.Lab_Total_Amt," & _
    "E1.Lab_Code,E1.Lab_Desc,E1.Hrs_Taken,E1.lab_CHARGES, " & _
    "City.CityName, HC.Chassis, HC.Engine, E.RegNo, HC.RegDate,JC.Job_No " & _
    "from (((Estimate as E left Join job_card as JC on E.Job_DocId=JC.DocId) " & _
    "left Join Hiscard as HC on E.CardNo=HC.CardNo) " & _
    "Left Join City on E.CityCode=City.CityCode) " & _
    "Left Join Estimate1 as E1 on E.DocID=E1.DocID " & _
    "WHERE E.docId= '" & Master!SearchCode & "' "

Set Rst = New Recordset
Rst.CursorLocation = adUseClient
Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic

If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
RepTitle = GCn.Execute("Select Div_SName from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
Speciality = GCn.Execute("Select W_SecSpeciality from Division where Div_Code='" & PubDivCode & "'").Fields(0).Value
 
CreateFieldDefFile Rst, PubRepoPath + "\WksProLab.ttx", True
Set rpt = rdApp.OpenReport(PubRepoPath & "\WksProLab.RPT")

Set RST1 = New Recordset
RST1.CursorLocation = adUseClient
RST1.Open "select W_SecLST,W_SecLST_Date,W_SecCST,W_SecCST_Date,W_SecPhone,W_SecFax,W_SecGram from division where Div_Code='" & PubDivCode & "' and W_SecCompCode =  '" & PubWCompCode & "'", GCn, adOpenDynamic, adLockOptimistic

For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("SubTitle")
            rpt.FormulaFields(I).TEXT = "'" & Speciality & "'"
        Case UCase("FormulaStr1")
            rpt.FormulaFields(I).TEXT = "'" & FormulaStr1 & "'"
        Case UCase("FormulaStr2")
            If SecondStr = True Then
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr2 & "'"
            Else
                rpt.FormulaFields(I).TEXT = ""
            End If
        Case UCase("Phone")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecPhone & "'"
        Case UCase("Fax")
            rpt.FormulaFields(I).TEXT = "'" & RST1!W_SecFax & "'"
    End Select
Next
           
rpt.Database.SetDataSource Rst
rpt.ReadRecords
'        For i = 1 To rpt.FormulaFields.Count
'        Select Case UCase(rpt.FormulaFields(i).FormulaFieldName)
'            Case UCase("comp_name")
'                rpt.FormulaFields(i).Text = "'" & PubComp_Name & "'"
'            Case UCase("comp_add1")
'                rpt.FormulaFields(i).Text = "'" & PubComp_Add & "'"
'            Case UCase("comp_add2")
'                rpt.FormulaFields(i).Text = "'" & PubComp_Add2 & "'"
'            Case UCase("comp_city")
'                rpt.FormulaFields(i).Text = "'" & PubComp_City & "'"
'            Case UCase("Title")
'                rpt.FormulaFields(i).Text = "'Performa Labour Invoice'"
'        End Select
'        Next
'        rpt.PrintOut False
        
        Call Report_View(rpt, "Performa Labour Invoice", , False)


Set Rst = Nothing
Set RST1 = Nothing
Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub TopCtrl1_eRef()
    Call UpdRequery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mTrans As Boolean
    Dim SrNo As Integer
    On Error GoTo errlbl

    If txtgrid1(0).Visible = True Then
        If TxtGrid1Leave = False Then
            txtgrid1(0).SetFocus
            Exit Sub
        Else
            txtgrid1(0).Visible = False
        End If
    End If

    Grid_Hide
    
    If IsValid(txt(ProfNo), "Performa No.") = False Then Exit Sub
    If IsValid(txt(ProfDt), "Performa Date") = False Then Exit Sub
    If txt(JCType).TEXT = "Yes" Then
        If IsValid(txt(JobNo), "JobCard No") = False Then Exit Sub
    End If
    If IsValid(txt(Party), "Party Name") = False Then Exit Sub

    '' checking for data in fgrid1
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, C_LabCode) <> "" Then GoTo Mynxt
    Next I
    MsgBox "No Labour Details Feeded ", vbInformation
    Exit Sub

Mynxt:
    '' eof : checking of data in fgrid1
    
    GCn.BeginTrans
    mTrans = True
    
    Select Case TopCtrl1.TopText2
        Case "Add"
            'lp 11-03-03
            ProfDocId = txt(TxtDocID)
            If GCn.Execute("Select Count(*) From estimate Where DocID='" & txt(TxtDocID) & "'").Fields(0) > 0 Then
                If VoucherEditFlag Then 'Manual No. System
                    MsgBox "Performa No. " & txt(ProfNo) & " Already Exists", vbCritical, "Validation Error"
                    txt(ProfNo).SetFocus
                    GoTo errlbl
                Else
                    txt(TxtDocID) = GetDocID(GCnFaW, VType, txt(ProfDt), VoucherEditFlag, txt(ProfNo), LblVPrefix, ForSiteCode)
                    If Val(txt(ProfNo)) <= Val(DeCodeDocID(ProfDocId, Document_No)) Then
                        MsgBox "Performa No. " & txt(ProfNo) & " already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                        GoTo errlbl
                    End If
                End If
            End If
            'lp end
            GSQL = "insert into estimate(DocId,DocIdHelp,Site_Code,V_Type,V_No,V_Date,Stores_Works,Job_DocId,CardNo,Party_Name,Address,Address2,Address3,CityCode,PhoneNo,Model,RegNo,Chassis,Engine,Lab_Amt,Lab_D_Amt,Lab_TaxPer,Lab_TaxAmt,Lab_Rounded,Lab_Total_Amt,U_Name, U_EntDt, U_AE) " & _
                " values('" & txt(TxtDocID) & "','" & Replace(txt(TxtDocID), " ", "") & "','" & PubSiteCode & "','" & VType & "'," & txt(ProfNo).TEXT & "," & ConvertDate(txt(ProfDt).TEXT) & ",'Workshop','" & txt(JobNo).Tag & "','" & MyCardNo & "','" & txt(Party).TEXT & "','" & txt(Address1).TEXT & "','" & txt(Address2).TEXT & "','" & txt(Address3).TEXT & "','" & txt(City).Tag & "','" & txt(PhoneOff).TEXT & "'," & _
                " '" & txt(Model) & "','" & txt(VehRegNo) & "','" & txt(Chassis) & "','" & txt(Engine) & "'," & Val(txt(TOTAmt).TEXT) & "," & Val(txt(DiscAmt).TEXT) & "," & Val(txt(SrvPerc).TEXT) & "," & Val(txt(SrvAmt).TEXT) & "," & Val(txt(Roundedoff).TEXT) & "," & Val(txt(NetAmt).TEXT) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
            GCn.Execute GSQL
            'Voucher Serial No. Updation LPS 21-05-03
            'update Table only when DocSrlNo >Table.SerialNo
            UpdVouSrlNo GCnFaS, txt(TxtDocID), txt(ProfDt)
        Case "Edit"
            GCn.Execute "Delete from Estimate1  where Docid='" & txt(TxtDocID) & "'"
            GSQL = "update estimate set V_Date = " & ConvertDate(txt(ProfDt).TEXT) & ",Job_DocId = '" & txt(JobNo).Tag & "',CardNo='" & MyCardNo & "',Party_Name='" & txt(Party).TEXT & "',Address='" & txt(Address1).TEXT & "',Address2='" & txt(Address2).TEXT & "',Address3='" & txt(Address3).TEXT & "',phoneno='" & txt(PhoneOff).TEXT & "',CityCode='" & txt(City).Tag & "',Model='" & txt(Model) & "',RegNo='" & txt(VehRegNo) & "',Chassis='" & txt(Chassis) & "',Engine='" & txt(Engine) & "',Lab_Amt=" & txt(TOTAmt).TEXT & ",Lab_D_Amt=" & txt(DiscAmt).TEXT & ",Lab_TaxPer=" & txt(SrvPerc).TEXT & ",Lab_TaxAmt=" & txt(SrvAmt).TEXT & ",Lab_Rounded=" & txt(Roundedoff).TEXT & ",Lab_Total_Amt=" & txt(NetAmt).TEXT & ",U_Name='" & pubUName & "', U_EntDt=" & ConvertDate(PubServerDate) & ", U_AE='E' where docId='" & txt(TxtDocID) & "'"
            GCn.Execute GSQL
    End Select
    SrNo = 1
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, C_LabCode) <> "" And (Not IsNull(FGrid1.TextMatrix(I, C_LabCode))) Then
            GSQL = "insert into estimate1(" _
                & "DocId,Site_Code,Sr_No,V_type,lab_code," _
                & "hrs_taken,lab_charges,Lab_Desc,Remarks," _
                & "U_Name, U_EntDt, U_AE) " _
                & " values(" _
                & "'" & ProfDocId & "','" & PubSiteCode & "'," & SrNo & ",'" & VType & "','" & FGrid1.TextMatrix(I, C_LabCode) & "'," _
                & "" & Val(FGrid1.TextMatrix(I, C_ChgHrs)) & "," & Val(FGrid1.TextMatrix(I, C_ChgAmt)) & ", '" & FGrid1.TextMatrix(I, C_LabName) & "','" & FGrid1.TextMatrix(I, C_Remarks) & "'," _
                & "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')"
            GCn.Execute GSQL
            SrNo = SrNo + 1
        End If
    Next I
    GCn.CommitTrans
    mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("select ES.docId as SearchCode,ES.docId,ES.JOB_DOCID,ES.SITE_CODE,ES.V_NO,ES.V_DATE,ES.STORES_WORKS,ES.PARTY_NAME,ES.ADDRESS,ES.ADDRESS2,ES.ADDRESS3,ES.PHONENO,ES.CARDNO,ES.CITYCODE,ES.LAB_AMT,ES.LAB_D_AMT,ES.LAB_TAXPER,ES.LAB_TAXAMT,ES.Lab_Rounded,ES.Lab_Total_Amt,JC.Job_No,JC.Job_Date,JC.JobCloseDate,ES.Model,ES.RegNo,ES.Chassis,ES.Engine,City.CityName " & _
                    "from ((ESTIMATE as ES left Join job_card as JC on ES.Job_DocId=JC.DocId) " & _
                    "left Join Hiscard as HC on ES.CardNo=HC.CardNo) " & _
                    "Left Join City on ES.CityCode=City.CityCode " & _
                    "WHERE LEFT(ES.DocId,1)='" & PubDivCode & "' AND ES.V_TYPE= '" & VType & "' And ES.docId = '" & txt(TxtDocID) & "' order by ES.V_Date desc,ES.V_Type,ES.DocID desc")
    End If
    Call UpdRequery
    
    Master.FIND "searchcode = '" & txt(TxtDocID) & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(ProfNo)) <> Val(DeCodeDocID(ProfDocId, Document_No)) Then
            MsgBox "Performa No." & Trim(DeCodeDocID(ProfDocId, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(ProfNo) & " alloted", vbCritical, "Document No. Changed"
        End If
        txt(ProfDt).Tag = txt(ProfDt).TEXT
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub

errlbl:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus txt(Index)
    txtgrid1(0).Visible = False
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case VehRegNo   'if not job no, details from history
            If RsHist.RecordCount = 0 Or txt(Index) = "" Then Exit Sub
            If UCase(txt(Index)) <> UCase(RsHist!RegNo) Then
                RsHist.MoveFirst
                RsHist.FIND "RegNo ='" & txt(Index) & "'"
            End If
        Case JCType
            If txt(JCType).TEXT = "" Then
                txt(JCType).TEXT = "No"
            End If
        Case JobNo
            DGridColSwap DGJob, 0
            RsJob.Sort = "JOB_NO"
            If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).TEXT <> RsJob!Code Then
                RsJob.MoveFirst
                RsJob.FIND "JOB_NO ='" & txt(Index).TEXT & "'"
            End If
'modi lps 03.09.03
'        Case Chassis
'            DGridColSwap DGJob, 1
'            RsJob.Sort = "CHASSIS"
'            If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).Text = "" Then Exit Sub
'            If txt(Index).Text <> RsJob!Code Then
'                RsJob.MoveFirst
'                RsJob.FIND "CHASSIS ='" & txt(Index).Text & "'"
'            End If
'eof
'            If Txt(Index).Tag <> "" And Txt(Index).Tag <> RsJob!Code Then
'                RsJob.FIND ("CHASSIS='" & Txt(Index).Text & "'")
'            End If
'modi lps 03.09.03
'        Case VehRegNo
'            DGridColSwap DGJob, 2
'            RsJob.Sort = "REGNO"
'            If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).Text = "" Then Exit Sub
'            If txt(Index).Text <> RsJob!Code Then
'                RsJob.MoveFirst
'                RsJob.FIND "RegNo ='" & txt(Index).Text & "'"
'            End If
'eof
'            If Txt(Index).Tag <> "" And Txt(Index).Tag <> RsJob!Code Then
'                RsJob.FIND ("REGNO='" & Txt(Index).Text & "'")
'            End If
        Case City
            DGridColSwap DgCity, 1
            RsCity.Sort = "code"
            If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).TEXT <> RsJob!Code Then
                RsCity.MoveFirst
                RsCity.FIND ("code='" & txt(Index).Tag & "'")
            End If
'            If Txt(Index).Tag <> "" And Txt(Index).Tag <> Rscity!Code Then
'                Rscity.FIND ("code='" & Txt(Index).Tag & "'")
'            End If
        Case DiscAmt, SrvPerc
            Call Amt_Calc
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case JobNo
            DGridTxtKeyDown DGJob, txt, Index, RsJob, KeyCode, False, 1
        Case VehRegNo
            DGridTxtKeyDown DGHist, txt, Index, RsHist, KeyCode, False, 2
'modi lps 03.09.03
'        Case VehRegNo
'            DGridTxtKeyDown DGJob, txt, Index, RsJob, KeyCode, False, 3
'        Case Chassis
'            DGridTxtKeyDown DGJob, txt, Index, RsJob, KeyCode, False, 4
'eof
        Case City
            DGridTxtKeyDown DgCity, txt, Index, RsCity, KeyCode, False, 1, frmCity, "frmCity"
        Case DiscAmt, SrvPerc
            Call Amt_Calc
    End Select
    If DGHist.Visible = False And DGJob.Visible = False And DgCity.Visible = False And DGLabour.Visible = False Then
        '' KEY DOWN and Enter Key
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> SrvPerc Then
            Ctrl_DownKeyDown KeyCode, Shift
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = SrvPerc Then
            If MsgBox("Save Entry ?", vbInformation + vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        ' KEY UP
        If TopCtrl1.TopText2 = "Add" Then
            If (txt(ProfNo).Enabled = False And Index <> ProfDt) Or (txt(ProfNo).Enabled = True And Index <> ProfNo) Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then
            If Index <> ProfDt Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
    Select Case Index
        Case VehRegNo, City, Party
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
    
    Select Case Index
        Case ProfNo
            Call NumPress(txt(Index), KeyAscii, 8, 0)
        Case JobNo
            DGridTxtKeyPress txt, Index, RsJob, KeyAscii, "FindJobNo"
        Case VehRegNo
            If DGHist.Visible = True Then DGridTxtKeyPress txt, Index, RsHist, KeyAscii, "RegNo"
'modi lps 03.09.03
'            DGridTxtKeyPress txt, Index, RsJob, KeyAscii, "regno"
'        Case Chassis
'            DGridTxtKeyPress txt, Index, RsJob, KeyAscii, "chassis"
'eof
        Case City
            DGridTxtKeyPress txt, Index, RsCity, KeyAscii, "name"
        Case DiscAmt, SrvPerc
            Call Amt_Calc
        Case JCType
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
        Case DiscAmt, SrvPerc
            Call Amt_Calc
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
    Select Case Index
        Case ProfNo
            txt(TxtDocID) = GetDocID(GCnFaW, VType, txt(ProfDt).TEXT, VoucherEditFlag, txt(ProfNo), LblVPrefix, ForSiteCode)
            ProfDocId = txt(TxtDocID)
            lblDocId = "DocId : " & ProfDocId
            If VoucherEditFlag Then     ' Manual
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select Docid From estimate Where DocID='" & ProfDocId & "'", GCn, adOpenDynamic, adLockOptimistic
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                    If txt(ProfNo).Enabled = True Then txt(ProfNo).SetFocus
                End If
            End If
        Case JCType
            If txt(Index).TEXT = "Yes" Then
                txt(JobNo).Enabled = True
                txt(VehRegNo).Enabled = True
                txt(Chassis).Enabled = True
            Else
                txt(JCType).TEXT = "No"
                txt(JobNo).Enabled = False
                txt(VehRegNo).Enabled = False
                txt(Chassis).Enabled = False
                
                txt(JobNo).TEXT = ""
                txt(VehRegNo).TEXT = ""
                txt(Chassis).TEXT = ""
                
                txt(JobDt).TEXT = ""
                txt(JobCDt).TEXT = ""
                txt(Engine).TEXT = ""
                txt(Model).TEXT = ""
            End If
        Case VehRegNo
            If RsHist.EOF = True Or RsHist.BOF = True Then Exit Sub
            If RsHist.RecordCount > 0 Then
                txt(VehRegNo).TEXT = XNull(RsHist!RegNo)
                txt(Model) = XNull(RsHist!Model)
                txt(Chassis).TEXT = XNull(RsHist!Chassis)
                txt(Engine).TEXT = XNull(RsHist!Engine)
                txt(Party).TEXT = XNull(RsHist!Name)
                txt(Address1).TEXT = XNull(RsHist!Add1)
                txt(Address2).TEXT = XNull(RsHist!Add2)
                txt(City).Tag = XNull(RsHist!CityCode)
                txt(City).TEXT = XNull(RsHist!CityName)
                txt(PhoneOff).TEXT = XNull(RsHist!PhoneOff)
            End If
        Case JobNo  ', VehRegNo, Chassis    'modi lps 03.09.03
            If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or txt(Index).TEXT = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            Else
                If txt(Index).Tag <> "" Then
                    RsJob.Sort = "CODE"
                    RsJob.FIND ("CODE='" & txt(Index).Tag & "'")
                End If
                If RsJob.BOF = True Or RsJob.EOF = True Then Exit Sub
                Call History_Field
                txt(VehRegNo).Enabled = False
                txt(Chassis).Enabled = False
                txt(Party).Enabled = False
                txt(Address1).Enabled = False
                txt(Address2).Enabled = False
                txt(City).Enabled = False
                txt(PhoneOff).Enabled = False
            End If
        Case City
            If txt(Index).Tag <> "" Then
                RsCity.Sort = "CODE"
                RsCity.FIND ("CODE='" & txt(Index).Tag & "'")
                If RsCity.EOF = True Or RsCity.BOF = True Then Exit Sub
                txt(City).TEXT = RsCity!Name
            End If
        Case DiscAmt, SrvPerc
            txt(Index).TEXT = Format(txt(Index).TEXT, "0.00")
            Call Amt_Calc
    End Select
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    For I = 1 To txt.Count
        txt(I).TEXT = ""
        If I <> ProfDt Then
            txt(I).Tag = ""
        End If
    Next I
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
    ProfDocId = ""
    lblDocId.CAPTION = ""
    lblDocId.Refresh
End Sub

Private Sub MoveRec()
Dim Rs As Recordset
Dim mVor As String
Dim I As Integer
On Error GoTo error1
    If Master.RecordCount > 0 Then
        LblDiv.CAPTION = "Division : " & left(Master!DocID, 1)
        LblSite.CAPTION = "Site Code : " & Master!Site_Code
        lblDocId.CAPTION = Master!DocID
        ProfDocId = Master!DocID
        txt(TxtDocID) = Master!DocID
        
        txt(VehRegNo).Tag = XNull(Master!job_docid)
        txt(Chassis).Tag = XNull(Master!job_docid)
        txt(JobNo).Tag = XNull(Master!job_docid)

        LblVPrefix.CAPTION = mID(Master!DocID, 9, 5)
        txt(ProfNo).TEXT = Master!V_NO
        txt(ProfDt).TEXT = Format(Master!V_DATE, "dd/MMM/yyyy")

        If Master!job_docid <> "" Then
            txt(JCType).TEXT = "Yes"
            txt(JobNo).TEXT = XNull(Master!Job_No)
            txt(JobDt).TEXT = XNull(Master!Job_Date)
        Else
            txt(JCType).TEXT = "No"
            txt(JobNo).TEXT = ""
            txt(JobDt).TEXT = ""
        End If
        txt(Party).TEXT = XNull(Master!Party_Name)
        txt(Address1).TEXT = XNull(Master!Address)
        txt(Address2).TEXT = XNull(Master!Address2)
        txt(Address3).TEXT = XNull(Master!Address3)
        txt(PhoneOff).TEXT = XNull(Master!PhoneNO)
        txt(City).TEXT = XNull(Master!CityName)
        txt(City).Tag = XNull(Master!CityCode)
        txt(VehRegNo).TEXT = XNull(Master!RegNo)
        txt(Chassis).TEXT = XNull(Master!Chassis)
        txt(Model).TEXT = XNull(Master!Model)
        txt(Engine).TEXT = XNull(Master!Engine)
        txt(TOTAmt).TEXT = Format(Master!Lab_Amt, "0.00")
        txt(DiscAmt).TEXT = Format(Master!Lab_D_Amt, "0.00")
        txt(SrvPerc).TEXT = Format(Master!Lab_TaxPer, "0.00")
        txt(SrvAmt).TEXT = Format(Master!Lab_TaxAmt, "0.00")
        txt(Roundedoff).TEXT = Format(Master!Lab_Rounded, "0.00")
        txt(NetAmt).TEXT = Format(Master!Lab_Total_Amt, "0.00")
        
        MyCardNo = Master!CardNo
        Call Fill_Grid(Master!DocID)
    Else
        Call BlankText
    End If
    Grid_Hide
    FGrid1_GotFocus
    Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
    With FGrid1
        .left = Me.left
        .width = Me.width - 90
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 6
        
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 400
        
        .TextMatrix(0, C_LabCode) = "Labour Code"
        .ColAlignment(C_LabCode) = flexAlignLeftCenter
        .ColWidth(C_LabCode) = 1000
        
        .TextMatrix(0, C_LabName) = "Labour Description"
        .ColAlignment(C_LabName) = flexAlignLeftCenter
        .ColWidth(C_LabName) = 4200

        .TextMatrix(0, C_ChgHrs) = "Ch.Hrs."
        .ColAlignment(C_ChgHrs) = flexAlignRightCenter
        .ColWidth(C_ChgHrs) = 600

        .TextMatrix(0, C_ChgAmt) = "Ch.Amt."
        .ColAlignment(C_ChgAmt) = flexAlignRightCenter
        .ColWidth(C_ChgAmt) = 1000

        .TextMatrix(0, C_Remarks) = "Remarks"
        .ColAlignment(C_Remarks) = flexAlignLeftCenter
        .ColWidth(C_Remarks) = 4200
    End With
    BackColorSelLeave = FGrid1.BackColorSel
    ForeColorSelEnter = FGrid1.ForeColorSel
    
    DGJob.width = Me.width - 60: DGJob.left = FGrid1.left: DGJob.top = FGrid1.top: DGJob.height = Me.height - (DGJob.top + mBotScale)
    DGLabour.width = 7000:
    DGLabour.left = Me.width - (DGLabour.width + mRtScale): DGLabour.top = mTopScale: DGLabour.height = Me.height - (DGLabour.top + mBotScale)
    DGHist.width = FGrid1.width: DGHist.left = FGrid1.left: DGHist.top = FGrid1.top + FGrid1.height: DGHist.height = Me.height - (DGHist.top + mBotScale)
    DgCity.width = 4740: DgCity.left = Me.width - (DgCity.width + mRtScale): DgCity.top = mTopScale: DgCity.height = Me.height - (DgCity.top + mBotScale)
End Sub

Public Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 1 To txt.Count
        txt(I).Enabled = Enb
    Next
    
    For I = 1 To txt.Count
        txt(I).BackColor = CtrlBColOrg
        txt(I).ForeColor = CtrlFColOrg
    Next
    
    txtgrid1(0).BackColor = CtrlBCol
    txtgrid1(0).ForeColor = CtrlFCol
    txtgrid1(0).Enabled = Enb
    
    txt(JobCDt).Enabled = False
    txt(JobDt).Enabled = False
    txt(Engine).Enabled = False
    txt(Model).Enabled = False
    txt(TOTAmt).Enabled = False
    txt(SrvAmt).Enabled = False
    txt(Roundedoff).Enabled = False
    txt(NetAmt).Enabled = False
End Sub

Private Sub Grid_Hide()
    If DGJob.Visible = True Then DGJob.Visible = False
    If DGLabour.Visible = True Then DGLabour.Visible = False
    If DgCity.Visible = True Then DgCity.Visible = False
End Sub

Private Sub UpdRequery()
    RsJob.Requery
    RsLab.Requery
    RsCity.Requery
End Sub

Private Sub History_Field()
    txt(VehRegNo).Tag = XNull(RsJob!Code)
    txt(Chassis).Tag = XNull(RsJob!Code)
    txt(JobNo).Tag = XNull(RsJob!Code)
    
    txt(JobNo).TEXT = XNull(RsJob!Job_No)
    txt(JobDt).TEXT = RsJob!Job_Date
    txt(JobCDt).TEXT = IIf(RsJob!JobCloseDate = #1/1/1900# Or IsNull(RsJob!JobCloseDate), "", RsJob!JobCloseDate)
    txt(VehRegNo).TEXT = XNull(RsJob!RegNo)
    txt(Chassis).TEXT = XNull(RsJob!Chassis)
    txt(Model).TEXT = XNull(RsJob!Model)
    txt(Engine).TEXT = XNull(RsJob!Engine)
    txt(Party).TEXT = XNull(RsJob!Name)
    txt(Address1).TEXT = XNull(RsJob!Add1)
    txt(Address2).TEXT = XNull(RsJob!Add2)
'    txt(Address3).Text = XNull(RsJob!Add3)
    txt(City).TEXT = XNull(RsJob!CityName)
    txt(City).Tag = XNull(RsJob!CityCode)
    txt(PhoneOff).TEXT = XNull(RsJob!PhoneOff)
    MyCardNo = RsJob!CardNo
    RsCity.FIND ("code='" & txt(City).Tag & "'")
End Sub

Private Sub FGrid1_Click()
    txtgrid1(0).Visible = False
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    SetMaxLength
End Sub

Private Sub FGrid1_DblClick()
FGrid1_KeyPress vbKeyReturn
End Sub

Private Sub FGrid1_GotFocus()
    FGrid1.BackColorSel = BackColorSelEnter
    FGrid1.ForeColorSel = ForeColorSelEnter
    txtgrid1(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
        SendKeysA vbKeyTab, True
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid1.Tag = FGrid1.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGrid1.Col
            Case C_ChgHrs, C_ChgAmt
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
                Calc_GridAmt
            Case C_Remarks
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
        End Select
    End If
'    If KeyCode = vbKeyReturn Then
'        Select Case FGrid1.Col
'            Case C_LabName, C_ChgHrs
'                GridDblClick Me, FGrid1, TxtGrid1, 0
'                TAddMode = False
'            Case C_ChgAmt
'                GridDblClick Me, FGrid1, TxtGrid1, 0
'                TAddMode = False
'            Case C_Remarks
'                GridDblClick Me, FGrid1, TxtGrid1, 0
'                TAddMode = False
'        End Select
'    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
SetMaxLength
    Select Case FGrid1.Col
        Case C_LabCode
            Get_Text Me, FGrid1, txtgrid1, 0, False, KeyAscii
        Case C_LabName
            If FGrid1.TextMatrix(FGrid1.Row, C_LabCode) = "" Then
                Get_Text Me, FGrid1, txtgrid1, 0, False, KeyAscii
            Else
                FGrid1.Col = C_ChgAmt
            End If
        Case C_ChgHrs
            Get_Text Me, FGrid1, txtgrid1, 0, True, KeyAscii
        Case C_ChgAmt
            Get_Text Me, FGrid1, txtgrid1, 0, True, KeyAscii
        Case C_Remarks
            Get_Text Me, FGrid1, txtgrid1, 0, True, KeyAscii
    End Select
'    Select Case FGrid1.Col
'        Case C_LabName
'            Get_Text Me, FGrid1, TxtGrid1, 0, False, KeyAscii
'        Case C_ChgHrs
'            Get_Text Me, FGrid1, TxtGrid1, 0, True, KeyAscii
'        Case C_ChgAmt
'            Get_Text Me, FGrid1, TxtGrid1, 0, True, KeyAscii
'        Case C_Remarks
'            Get_Text Me, FGrid1, TxtGrid1, 0, True, KeyAscii
'        Case C_LabCode
'            FGrid1.Col = FGrid1.Col + 1
'            FGrid1.SetFocus
'    End Select
    If KeyAscii <> vbKeyReturn Then TAddMode = True
    Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid1.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGrid1.Row >= 1 Then
            If MsgBox("Are You Sure To Delete Entry ?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                If FGrid1.Rows > 2 Then
                    FGrid1.RemoveItem (FGrid1.Row)
                Else
                    FGrid1.Rows = 1
                    FGrid1.AddItem FGrid1.Rows
                    FGrid1.FixedRows = 1
                End If
            End If
            For I = 1 To FGrid1.Rows - 1
                FGrid1.TextMatrix(I, 0) = I
            Next
            Calc_GridAmt
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid1.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub FGrid1_LostFocus()
    FGrid1.BackColorSel = BackColorSelLeave
    FGrid1.ForeColorSel = FGrid1.ForeColor
    If TopCtrl1.TopText2 <> "Browse" Then Calc_GridAmt
End Sub

Private Sub FGrid1_Scroll()
    txtgrid1(0).Visible = False
    Grid_Hide
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
On Error GoTo ELoop
    Grid_Hide
    txtgrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    Select Case FGrid1.Col
        Case C_LabCode
            If RsLab.EOF = True Or RsLab.BOF = True Or txtgrid1(Index).TEXT = "" Then Exit Sub
            RsLab.MoveFirst
            RsLab.Sort = "Code"
            RsLab.FIND "Code='" & FGrid1.TextMatrix(FGrid1.Row, C_LabCode) & "'"
        Case C_LabName
            If RsLab.EOF = True Or RsLab.BOF = True Or txtgrid1(Index).TEXT = "" Then Exit Sub
            RsLab.MoveFirst
            RsLab.Sort = "name"
            RsLab.FIND "name='" & FGrid1.TextMatrix(FGrid1.Row, C_LabName) & "'"
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then txtgrid1(0).TEXT = txtgrid1(0).Tag: Exit Sub
    Select Case FGrid1.Col
        Case C_LabCode
            If DGLabour.Visible = False Then DGridColSwap DGLabour, 0
            DGridTxtKeyDown DGLabour, txtgrid1, 0, RsLab, KeyCode, True, 0, frmLabDesc, "frmLabDesc"
            If KeyCode = vbKeyReturn Then
                If TxtGrid1Leave Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks, 1
                End If
            End If
        Case C_LabName
            DGridColSwap DGLabour, 1
            DGridTxtKeyDown DGLabour, txtgrid1, 0, RsLab, KeyCode, True, 1, frmLabDesc, "frmLabDesc"
            If KeyCode = vbKeyReturn Then
                If TxtGrid1Leave Then GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks
            End If
        Case C_ChgHrs
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave Then GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks
                If FGrid1.Col = C_LabCode Then FGrid1.Col = C_LabName
            End If
        Case C_ChgAmt
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave Then GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks
                If FGrid1.Col = C_LabCode Then FGrid1.Col = C_LabName
            End If
        Case C_Remarks
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGrid1Leave Then GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_Remarks - 1
                If FGrid1.Col = C_LabCode Then FGrid1.Col = C_LabName
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
If KeyAscii = vbKeyEscape Then Exit Sub
    CheckQuote KeyAscii
    Select Case FGrid1.Col
        Case C_LabCode
            If DGLabour.Visible = True Then DGridTxtKeyPress txtgrid1, Index, RsLab, KeyAscii, "Code"
        Case C_ChgAmt
            NumPress txtgrid1(Index), KeyAscii, 6, 2
        Case C_ChgHrs
            NumPress txtgrid1(Index), KeyAscii, 2, 2
        Case C_LabName
            If DGLabour.Visible = True Then DGridTxtKeyPress txtgrid1, Index, RsLab, KeyAscii, "Name"
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
            Case C_LabCode
                If KeyCode <> 13 And DGLabour.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0: DGridTxtKeyPress txtgrid1, Index, RsLab, KeyCode, "Name", True
            Case C_LabName
                If KeyCode <> 13 And DGLabour.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0: DGridTxtKeyPress txtgrid1, Index, RsLab, KeyCode, "Name", True
        End Select
End Select
If KeyCode = vbKeyEscape Then
    FGrid1.SetFocus
    txtgrid1(0).Visible = False
End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGrid1Leave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid1Leave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim I As Integer
    Select Case FGrid1.Col
        Case C_LabCode, C_LabName
            If RsLab.EOF = True Or RsLab.BOF = True Or txtgrid1(Index).TEXT = "" Then
                FGrid1.TextMatrix(FGrid1.Row, C_LabCode) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_LabName) = ""

                FGrid1.TextMatrix(FGrid1.Row, C_ChgHrs) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_ChgAmt) = ""
                FGrid1.TextMatrix(FGrid1.Row, C_Remarks) = ""
            Else
                For I = 1 To FGrid1.Rows - 1
                    If FGrid1.TextMatrix(I, C_LabCode) = RsLab!Code And I <> FGrid1.Row Then
                        MsgBox "Duplicate Labour Not Allowed", vbInformation, "Validation"
                        GoTo NXT
                    End If
                Next I
                FGrid1.TextMatrix(FGrid1.Row, C_LabCode) = RsLab!Code
                FGrid1.TextMatrix(FGrid1.Row, C_LabName) = RsLab!Name

'                Set Rst = New ADODB.Recordset
'                Rst.CursorLocation = adUseClient
'                Rst.Open "Select lab_rate,time_req,wtime_req,fixed From labour_model Where model='" & txt(Model).Text & "' and lab_code='" & RsLab!Code & "'", GCn, adOpenDynamic, adLockOptimistic
'                If Rst.RecordCount  > 0 Then
'                    FGrid1.TextMatrix(FGrid1.Row, C_ChgHrs) = Rst!TIME_REQ
'                    If Rst!TIME_REQ  > 0 Then
'                        FGrid1.TextMatrix(FGrid1.Row, C_ChgAmt) = Rst!Lab_Rate
'                    Else
'                        FGrid1.TextMatrix(FGrid1.Row, C_ChgAmt) = "0.00"
'                    End If
'                Else
                    FGrid1.TextMatrix(FGrid1.Row, C_ChgHrs) = IIf(IsNull(RsLab!TIME_REQ), 0, RsLab!TIME_REQ)
'                    If RsLab!TIME_REQ  > 0 Then
                        FGrid1.TextMatrix(FGrid1.Row, C_ChgAmt) = IIf(IsNull(RsLab!Lab_Rate), 0, RsLab!Lab_Rate)
'                    Else
'                        FGrid1.TextMatrix(FGrid1.Row, C_ChgAmt) = "0.00"
'                    End If
'                End If
'                Set Rst = Nothing
            End If
        Case C_ChgHrs
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = Format(txtgrid1(0).TEXT, "0.00")
        Case C_ChgAmt
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = Format(txtgrid1(0).TEXT, "0.00")
        Case C_Remarks
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = txtgrid1(0).TEXT
    End Select
    TxtGrid1Leave = True
NXT:
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid1.SetFocus
    txtgrid1(Index).Visible = False
End If
End Function

Private Sub Fill_Grid(ByVal DocID As String)
Dim I As Integer
    FGrid1.Rows = 1
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    GSQL = "Select ES.* FROM ESTIMATE1 AS ES Where ES.DocId='" & DocID & "' order by ES.lab_code"
    Rst.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    I = 1
    If Rst.RecordCount > 0 Then
        Do Until Rst.EOF
            FGrid1.AddItem ""
            With FGrid1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, C_LabCode) = Rst!Lab_Code
                .TextMatrix(I, C_LabName) = XNull(Rst!Lab_Desc)
                .TextMatrix(I, C_ChgHrs) = Format(Rst!Hrs_Taken, "0.00")
                .TextMatrix(I, C_ChgAmt) = Format(Rst!lab_CHARGES, "0.00")
                .TextMatrix(I, C_Remarks) = XNull(Rst!Remarks)
            End With
            I = I + 1
            Rst.MoveNext
        Loop
        FGrid1.AddItem ""
        FGrid1.FixedRows = 1
    Else
        FGrid1.Rows = FGrid1.Rows
        FGrid1.AddItem ""
        FGrid1.FixedRows = 1
    End If
    Set Rst = Nothing
End Sub

Private Sub Amt_Calc()
Dim Mytot As Double
    txt(SrvAmt).TEXT = Format((Val(txt(TOTAmt).TEXT) - Val(txt(DiscAmt).TEXT)) * Val(txt(SrvPerc).TEXT) / 100, "0.00")
    Mytot = (Val(txt(TOTAmt).TEXT) - Val(txt(DiscAmt).TEXT)) + txt(SrvAmt).TEXT
    txt(Roundedoff).TEXT = Format(Round(Mytot, 0) - Mytot, "0.00")
    txt(NetAmt).TEXT = Format(Round(Mytot, 0), "0.00")
End Sub

Private Sub Calc_GridAmt()
Dim Mytot As Double, I As Integer
    Mytot = 0
    For I = 1 To FGrid1.Rows - 1
        Mytot = Mytot + Val(FGrid1.TextMatrix(I, C_ChgAmt))
    Next I
    txt(TOTAmt).TEXT = Format(Mytot, "0.00")
    Call Amt_Calc
End Sub

Private Sub SetMaxLength()
Select Case FGrid1.Col
        Case C_LabName
            txtgrid1(0).MaxLength = 40
            txtgrid1(0).Alignment = 0
        Case C_ChgHrs, C_ChgAmt
            txtgrid1(0).Alignment = 1
        Case C_Remarks
            txtgrid1(0).Alignment = 0
            txtgrid1(0).MaxLength = 20
        Case Else
            txtgrid1(0).MaxLength = 0
    End Select
End Sub
