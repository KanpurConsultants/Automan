VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmWarrantyBill 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Warranty Bill Entry"
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
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   6045
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DGClaim 
      Height          =   2655
      Left            =   -240
      Negotiate       =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   6180
      Visible         =   0   'False
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   4683
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Claim DocID"
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
         DataField       =   "ProwNo"
         Caption         =   "Prow No."
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
         DataField       =   "ProwDt"
         Caption         =   "Prow Dt."
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
         DataField       =   "EngNo"
         Caption         =   "Eng/Aggr.No."
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
         DataField       =   "JobNo"
         Caption         =   "Job Card"
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
         DataField       =   "PCR_Date"
         Caption         =   "PCR Date"
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
            DividerStyle    =   6
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   6
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            DividerStyle    =   6
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            DividerStyle    =   6
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            DividerStyle    =   6
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   2055.118
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2670
      MaxLength       =   5
      TabIndex        =   50
      Top             =   4860
      Width           =   345
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
      Left            =   4890
      MaxLength       =   5
      TabIndex        =   52
      Top             =   4860
      Width           =   345
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
      Left            =   4890
      MaxLength       =   5
      TabIndex        =   53
      Top             =   5130
      Width           =   345
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
      Index           =   29
      Left            =   10050
      MaxLength       =   8
      TabIndex        =   61
      Top             =   5325
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
      Height          =   255
      Index           =   21
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   16
      Top             =   6045
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
      Height          =   255
      Index           =   22
      Left            =   5295
      MaxLength       =   14
      TabIndex        =   17
      Top             =   6045
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
      Height          =   255
      Index           =   20
      Left            =   2670
      MaxLength       =   5
      TabIndex        =   56
      Top             =   6045
      Width           =   345
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
      Index           =   26
      Left            =   9630
      MaxLength       =   1
      TabIndex        =   57
      Top             =   4785
      Width           =   345
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
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   15
      Top             =   5775
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
      Height          =   255
      Index           =   17
      Left            =   3075
      MaxLength       =   14
      TabIndex        =   14
      Top             =   5505
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
      Height          =   255
      Index           =   18
      Left            =   2670
      MaxLength       =   5
      TabIndex        =   55
      Top             =   5775
      Width           =   345
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
      Index           =   16
      Left            =   2670
      MaxLength       =   5
      TabIndex        =   54
      Top             =   5505
      Width           =   345
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
      Index           =   6
      Left            =   5295
      MaxLength       =   20
      TabIndex        =   8
      Top             =   4230
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
      Index           =   5
      Left            =   3075
      MaxLength       =   14
      TabIndex        =   7
      Top             =   4230
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
      Index           =   7
      Left            =   5295
      MaxLength       =   20
      TabIndex        =   9
      Top             =   4500
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
      Index           =   25
      Left            =   10050
      MaxLength       =   10
      TabIndex        =   20
      Top             =   4515
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
      Index           =   31
      Left            =   10050
      MaxLength       =   20
      TabIndex        =   21
      Top             =   5865
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   10050
      MaxLength       =   8
      TabIndex        =   60
      Top             =   5055
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
      Index           =   1
      Left            =   2175
      MaxLength       =   2
      TabIndex        =   1
      Top             =   525
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
      Index           =   3
      Left            =   1545
      MaxLength       =   40
      TabIndex        =   3
      Top             =   795
      Width           =   5115
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
      Index           =   32
      Left            =   10050
      MaxLength       =   8
      TabIndex        =   22
      Top             =   6540
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   5295
      MaxLength       =   20
      TabIndex        =   19
      Top             =   6540
      Width           =   1425
   End
   Begin VB.TextBox txtgrid1 
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
      Height          =   270
      Index           =   0
      Left            =   6705
      MaxLength       =   40
      TabIndex        =   5
      Top             =   2415
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
      Index           =   4
      Left            =   1545
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1065
      Width           =   5115
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
      Left            =   5235
      MaxLength       =   20
      TabIndex        =   2
      Top             =   525
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
      Height          =   255
      Index           =   11
      Left            =   5295
      MaxLength       =   20
      TabIndex        =   11
      Top             =   4860
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
      Height          =   255
      Index           =   12
      Left            =   2670
      MaxLength       =   5
      TabIndex        =   51
      Top             =   5130
      Width           =   345
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
      Left            =   3075
      MaxLength       =   14
      TabIndex        =   10
      Top             =   4860
      Width           =   1425
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
      Index           =   15
      Left            =   5295
      MaxLength       =   15
      TabIndex        =   13
      Top             =   5130
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
      Height          =   255
      Index           =   13
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   12
      Top             =   5130
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   18
      Top             =   6540
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
      Index           =   27
      Left            =   10050
      MaxLength       =   8
      TabIndex        =   58
      Top             =   4785
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
      Height          =   255
      Index           =   30
      Left            =   10050
      MaxLength       =   15
      TabIndex        =   71
      Top             =   5595
      Width           =   1425
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   2445
      Left            =   75
      TabIndex        =   6
      Top             =   1425
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   4313
      _Version        =   393216
      BackColor       =   12243913
      Cols            =   3
      BackColorFixed  =   4210816
      ForeColorFixed  =   65535
      BackColorSel    =   16711680
      BackColorBkg    =   14667998
      GridColor       =   128
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
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2865
      Left            =   7335
      Negotiate       =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   7305
      Visible         =   0   'False
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   5054
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
            DividerStyle    =   6
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3495.118
         EndProperty
      EndProperty
   End
   Begin VB.Line Line5 
      X1              =   30
      X2              =   7065
      Y1              =   5445
      Y2              =   5445
   End
   Begin VB.Shape Shape4 
      Height          =   450
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   6435
      Width           =   7035
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FAD1D5&
      Caption         =   "%"
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
      Index           =   24
      Left            =   9345
      TabIndex        =   68
      Top             =   4800
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001860A7&
      Height          =   225
      Index           =   19
      Left            =   4620
      TabIndex        =   67
      Top             =   5145
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001860A7&
      Height          =   225
      Index           =   18
      Left            =   4620
      TabIndex        =   66
      Top             =   4875
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "%"
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
      Index           =   17
      Left            =   2445
      TabIndex        =   65
      Top             =   6060
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "%"
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
      Index           =   12
      Left            =   2445
      TabIndex        =   64
      Top             =   5790
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "%"
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
      Index           =   10
      Left            =   2445
      TabIndex        =   63
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001860A7&
      Height          =   225
      Index           =   4
      Left            =   2445
      TabIndex        =   62
      Top             =   5145
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001860A7&
      Height          =   225
      Index           =   2
      Left            =   2445
      TabIndex        =   59
      Top             =   4875
      Width           =   180
   End
   Begin VB.Shape Shape3 
      Height          =   420
      Left            =   7575
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   4095
   End
   Begin VB.Line Line6 
      X1              =   7035
      X2              =   7035
      Y1              =   3930
      Y2              =   6405
   End
   Begin VB.Line Line4 
      X1              =   45
      X2              =   7020
      Y1              =   4170
      Y2              =   4170
   End
   Begin VB.Line Line3 
      X1              =   4560
      X2              =   4560
      Y1              =   3945
      Y2              =   6405
   End
   Begin VB.Line Line2 
      X1              =   2310
      X2              =   2310
      Y1              =   3930
      Y2              =   6390
   End
   Begin VB.Line Line1 
      X1              =   15
      X2              =   7020
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000040C0&
      Height          =   2460
      Left            =   30
      Top             =   3930
      Width           =   11700
   End
   Begin VB.Label lblPrefix 
      BackStyle       =   0  'Transparent
      Caption         =   "VPrefix"
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
      Left            =   1575
      TabIndex        =   49
      Top             =   540
      Width           =   630
   End
   Begin VB.Label lblDocId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LblDocId"
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
      Left            =   8175
      TabIndex        =   46
      Top             =   945
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill DocID"
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
      Left            =   6990
      TabIndex        =   45
      Top             =   945
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Misc.Charges"
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
      Index           =   23
      Left            =   8130
      TabIndex        =   44
      Top             =   5070
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Amount"
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
      Index           =   22
      Left            =   75
      TabIndex        =   43
      Top             =   4245
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oil Amount"
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
      Index           =   21
      Left            =   75
      TabIndex        =   42
      Top             =   4515
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LST Paid && Recovered "
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
      TabIndex        =   41
      Top             =   4875
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount (A+B)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00426388&
      Height          =   225
      Index           =   14
      Left            =   7800
      TabIndex        =   40
      Top             =   6555
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour Amount"
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
      Index           =   11
      Left            =   7995
      TabIndex        =   38
      Top             =   4530
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TurnOver Tax Payable"
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
      Left            =   75
      TabIndex        =   37
      Top             =   6060
      Width           =   1785
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
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
      Left            =   75
      TabIndex        =   36
      Top             =   540
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name"
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
      Left            =   75
      TabIndex        =   35
      Top             =   810
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total (B)"
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
      Index           =   15
      Left            =   8220
      TabIndex        =   34
      Top             =   5880
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Tax on Labour"
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
      Left            =   7485
      TabIndex        =   33
      Top             =   4800
      Width           =   1815
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   9
      Left            =   75
      TabIndex        =   32
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Dt."
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
      Index           =   1
      Left            =   4620
      TabIndex        =   31
      Top             =   540
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surcharge On LST Payable"
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
      Index           =   39
      Left            =   75
      TabIndex        =   30
      Top             =   5790
      Width           =   2220
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
      Left            =   6990
      TabIndex        =   29
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rounded Off"
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
      Left            =   8265
      TabIndex        =   28
      Top             =   5610
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LST Payable"
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
      TabIndex        =   27
      Top             =   5520
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOT Paid && Recovered"
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
      Left            =   75
      TabIndex        =   26
      Top             =   5145
      Width           =   1860
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   780
      Left            =   6825
      Top             =   540
      Width           =   4830
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
      Left            =   9795
      TabIndex        =   25
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spl. Charges"
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
      Index           =   37
      Left            =   8220
      TabIndex        =   24
      Top             =   5340
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total (A)"
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
      Index           =   33
      Left            =   75
      TabIndex        =   23
      Top             =   6555
      Width           =   1080
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "                                                           Taxable Amount                        TaxPaid Amount"
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
      Index           =   13
      Left            =   60
      TabIndex        =   39
      Top             =   3960
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Height          =   585
      Left            =   60
      TabIndex        =   69
      Top             =   4200
      Width           =   6975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   45
      TabIndex        =   70
      Top             =   4815
      Width           =   6975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   30
      TabIndex        =   72
      Top             =   5460
      Width           =   6975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FAD1D5&
      Height          =   2430
      Left            =   7065
      TabIndex        =   73
      Top             =   3945
      Width           =   4650
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Height          =   420
      Left            =   60
      TabIndex        =   74
      Top             =   6450
      Width           =   6975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   7605
      TabIndex        =   75
      Top             =   6525
      Width           =   4035
   End
End
Attribute VB_Name = "frmWarrantyBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TAddMode As Boolean
Dim ExitCtrl As Boolean
Dim GridKey As Integer


Dim VoucherEditFlag As Boolean
Private Const VType As String = "W_WB"

Dim ForSiteCode As String

Dim MyIndex As Byte
Dim Rst As ADODB.Recordset

Dim Master As ADODB.Recordset
Dim RsClaim As ADODB.Recordset
Dim RsParty As ADODB.Recordset

'Text Box (Form)
Private Const BillNo As Byte = 1
Private Const BillDt As Byte = 2
Private Const Supplier As Byte = 3
Private Const WarrAc As Byte = 4
Private Const SprTB As Byte = 5
Private Const SprTP As Byte = 6
Private Const OilTP As Byte = 7

Private Const RecCST_Per As Byte = 8
Private Const RecCST_Amt As Byte = 9
Private Const RecLST_Per As Byte = 10
Private Const RecLST_Amt As Byte = 11
Private Const RecTOT_TBPer As Byte = 12
Private Const RecTOT_TBAmt As Byte = 13
Private Const RecTOT_TPPer As Byte = 14
Private Const RecTOT_TPAmt As Byte = 15
Private Const Lst_Per As Byte = 16
Private Const Lst_Amt As Byte = 17
Private Const Surc_Per As Byte = 18
Private Const Surc_Amt As Byte = 19

Private Const TOT_Per As Byte = 20
Private Const TOT_TBAmt As Byte = 21
Private Const TOT_TPAmt As Byte = 22

Private Const SubATB As Byte = 23
Private Const SubATP As Byte = 24

Private Const LabAmt As Byte = 25
Private Const Ser_Per As Byte = 26
Private Const Ser_Amt As Byte = 27
Private Const MiscAmt As Byte = 28
Private Const SplAmt As Byte = 29
Private Const RoundOff As Byte = 30
Private Const SubB As Byte = 31
Private Const NetAmt As Byte = 32

'Fgrid1 Columns
Private Const C_ClmNo As Byte = 1
Private Const C_ClmDt As Byte = 2
Private Const C_SprTB As Byte = 3
Private Const C_SprTP As Byte = 4
Private Const C_OilTP As Byte = 5
Private Const C_Misc As Byte = 6
Private Const C_Labour As Byte = 7
Private Const C_Spl As Byte = 8
Private Const C_ID As Byte = 9

Private Sub DGParty_Click()
If RsParty.RecordCount > 0 Then
    Txt(MyIndex).TEXT = RsParty!Name
    Txt(MyIndex).Tag = RsParty!Code
End If
DGParty.Visible = False
Txt(MyIndex).SetFocus
End Sub

Private Sub DGClaim_Click()
If RsClaim.RecordCount > 0 Then
    Select Case FGrid1.Col
        Case C_ClmNo
            txtgrid1(0).TEXT = RsClaim!ProwNo
    End Select
    txtgrid1(0).Tag = RsClaim!Code
End If
DGClaim.Visible = False
txtgrid1(0).SetFocus
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
Dim SrNo As Integer
    
    TopCtrl1.Tag = PubUParam ' UserPermission(Me.Name)
    ForSiteCode = PubSiteCode
    Call BlankText
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select Jw.v_no AS CODE,Jw.v_no AS SEARCHCODE,JW.*,Supp.Name as SuppName,Warr.Name as WarrName " _
                & "FROM (JOB_WARBILL AS JW left join Subgroup as Supp on jw.Party_Ac=Supp.Subcode) left join Subgroup as Warr on jw.Warranty_Ac=Warr.Subcode where left(JW.site_code,1)='" & PubSiteCode & "' order by JW.DocId", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsClaim = GCn.Execute("select Jw1.Docid AS CODE,ProwNo,ProwDt,EngNo,right(Jw1.job_DocId,8) as JobNo, " & cCStr("Pcr_Date") & " FROM JOB_WARR1 AS JW1 where JW1.WBill_DocId = '' and JW1.ProwNo <> ''")
    Set DGClaim.DataSource = RsClaim
    
    GSQL = "Select SG.SubCode as Code,SG.Name,Nature " & _
        " From (SubGroup SG " & FaTable("AcGroup") & " on SG.GroupCode=AcGroup.GroupCode) " & _
        " Where  " & _
        " left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        " Order by SG.Name"
    Set RsParty = New ADODB.Recordset
    With RsParty
        .CursorLocation = adUseClient
'        .Open "Select Subcode as Code, Name, Nature FROM Subgroup Order by Name", GCn, adOpenDynamic, adLockOptimistic
        .Open GSQL, GCn, adOpenDynamic, adLockOptimistic
        .Sort = "Name"
    End With
    Set DGParty.DataSource = RsParty
    
    Ini_Grid
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
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
    Set RsClaim = Nothing
    Set RsParty = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    Txt(Ser_Per) = Format(PubService_Tax, "0.00")
    Txt(BillDt).TEXT = Format(date, "dd/MMM/yyyy")
    If GCn.Execute("Select VT.Number_Method From Voucher_Type VT  Where VT.V_Type='" & VType & "'").Fields(0).Value = "Manual" Then
        Txt(BillNo).TEXT = GCn.Execute("select " & vIsNull("max(V_No)", "0") & "+1 from job_WBill where left(docid,1)='" & PubDivCode & "' and " & cMID("docid", "2", "2") & "='" & PubSiteCode + ForSiteCode & "'").Fields(0)
    End If
    lblDocId = GetDocID(GCnFaS, VType, Txt(BillDt).TEXT, VoucherEditFlag, Txt(BillNo), lblPrefix, ForSiteCode)
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    If VoucherEditFlag = True Then
        Txt(BillNo).Enabled = True
        Txt(BillNo).SetFocus
    Else
        Txt(BillDt).SetFocus
    End If
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant, I As Integer
Dim LedgAry(1) As LedgRec, mResult As Byte
     
    For I = 1 To FGrid1.Rows - 1
        Set GRs = GCn.Execute("select V_DocID from job_warr2 where docid='" & FGrid1.TextMatrix(I, C_ID) & "'")
        If GRs.RecordCount > 0 Then
            If GRs!v_Docid <> "" Or GRs!v_Docid <> Null Then
                MsgBox "Warranti Credit Note has been made against this bill Entry Can't Delete.", vbInformation, "Validation"
                Exit Sub
            End If
        End If
    Next
    Set GRs = Nothing
   
    If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        GCn.BeginTrans
        GCnFaS.BeginTrans
        
        
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, lblDocId)
        If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"

        For I = 1 To FGrid1.Rows - 1
            If FGrid1.TextMatrix(I, C_ID) <> "" Then
                GSQL = "Update Job_WARR1 set " _
                        & "WBill_DocId='',Tax_Per=0,Tax_Sur_Per=0,TOT_Per=0,U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' WHERE DocId = '" & FGrid1.TextMatrix(I, C_ID) & "'"
                GCn.Execute GSQL
            End If
        Next
        GCn.Execute ("Delete * From JOB_WarBill Where DocId='" & lblDocId & "'")
        
        GCnFaS.CommitTrans
        GCn.CommitTrans
        Master.Requery
        
        Call UpdRequery
        
        If Master.RecordCount > 0 Then
            If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
            Call MoveRec
        Else
            Call BlankText
        End If
        
        BUTTONS True, Me, Master, 0
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
    LblDiv.CAPTION = "Division : " & DeCodeDocID(Master!DocID, Division_Code)
    LblSite.CAPTION = "Site Code : " & Master!Site_Code
    lblDocId.CAPTION = Master!DocID
    
    Txt(BillNo).Enabled = False
    Txt(BillDt).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
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
    Master.FIND ("code='" & MyValue & "'")
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
        txtDisabled_Color Me
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
    Dim SrNo As Integer
    Dim mTrans As Boolean
    Dim AddFlg As String, MyDocId As String
    Dim LedgAry(1) As LedgRec, mNarr$, mResult As Byte

'    On Error GoTo errlbl

    Grid_Hide
    
    If IsValid(Txt(Supplier), "Supplier Name") = False Then Exit Sub
    If IsValid(Txt(WarrAc), "Warranty Claim A/c ") = False Then Exit Sub
    If IsValid(Txt(BillDt), "Warranty Bill Date") = False Then Exit Sub
    If TopCtrl1.TopText2 = "Add" And VoucherEditFlag = True Then
        If IsValid(Txt(BillNo), "Warranty Bill No") = False Then Exit Sub
    End If
    
    '' checking for data in fgrid1
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, C_ID) <> "" Then GoTo Mynxt
    Next I
    MsgBox "No Claims are selected", vbInformation
    FGrid1.SetFocus
    Exit Sub

Mynxt:
    '' eof : checking of data in fgrid1
    
    GCn.BeginTrans
    mTrans = True
    
    Select Case TopCtrl1.TopText2
        Case "Add"
            AddFlg = "A"
            If VoucherEditFlag = True Then
                GSQL = "Select Count(*) From job_WarBill Where DocID='" & lblDocId.CAPTION & "'"
                If GCn.Execute(GSQL).Fields(0) > 0 Then
                    MsgBox "Warranty Bill No. " & Txt(BillNo).TEXT & " Already Exists", vbCritical, "Validation Ertror"
                    Txt(BillNo).SetFocus
                    Exit Sub
                End If
            Else
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                GSQL = "Select VP.Start_Srl_No From Voucher_Prefix VP Where VP.V_Type='" & VType & "' And VP.Date_From<=" & ConvertDate(Format(Txt(BillDt), "dd/MMM/yyyy")) & " Order By VP.Date_From DESC"
                Rst.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
                If Val(Rst!start_srl_no) >= Val(Txt(BillNo).TEXT) Then
                    lblDocId = GetDocID(GCnFaS, VType, Txt(BillDt).TEXT, VoucherEditFlag, Txt(BillNo), lblPrefix, ForSiteCode)
                End If
                If Rst.RecordCount > 0 Then
                    GSQL = "Update Voucher_Prefix Set start_srl_no=start_srl_no+1 where V_Type='" & VType & "' And Date_From<=" & ConvertDate(Format(Txt(BillDt), "dd/MMM/yyyy")) & ""
                    GCn.Execute GSQL
                End If
            End If
            GSQL = "insert into Job_WarBill(" _
                    & "DocId,Site_Code,V_Type,V_no,Party_Ac, " _
                    & "V_Date,Warranty_Ac,SprAmt_TB,SprAmt_TP,OilAmt_TP," _
                    & "RecdLST_TBPer,RecdLST_TPPer,RecdTOT_TBPer,RecdTOT_TPPer,Tax_Per," _
                    & "Tax_Sur_Per,TOT_Per,SrvTax_Per,RecdLst_TBAmt,RecdLst_TPAmt," _
                    & "RecdTOT_TBAmt,RecdTOT_TPAmt,Tax_Amt,Tax_Sur_Amt,TOT_TBAmt," _
                    & "TOT_TPAmt,Labour_Amt,SrvTax_Amt,MiscAmt,SplAmt," _
                    & "Rounded,Total_Amt," _
                    & "U_Name,U_EntDt,U_AE)" _
                    & " values(" _
                    & "'" & lblDocId & "','" & PubSiteCode & "','" & VType & "'," & Txt(BillNo) & ",'" & Txt(Supplier).Tag & "'," _
                    & "" & ConvertDate(Txt(BillDt)) & ",'" & Txt(WarrAc).Tag & "'," & Val(Txt(SprTB)) & "," & Val(Txt(SprTP)) & "," & Val(Txt(OilTP)) & "," _
                    & "" & Val(Txt(RecCST_Per)) & "," & Val(Txt(RecLST_Per)) & "," & Val(Txt(RecTOT_TBPer)) & "," & Val(Txt(RecTOT_TPPer)) & "," & Val(Txt(Lst_Per)) & "," _
                    & "" & Val(Txt(Surc_Per)) & "," & Val(Txt(TOT_Per)) & "," & Val(Txt(Ser_Per)) & "," & Val(Txt(RecCST_Amt)) & "," & Val(Txt(RecLST_Amt)) & "," _
                    & "" & Val(Txt(RecTOT_TBAmt)) & "," & Val(Txt(RecTOT_TPAmt)) & "," & Val(Txt(Lst_Amt)) & "," & Val(Txt(Surc_Amt)) & "," & Val(Txt(TOT_TBAmt)) & "," _
                    & "" & Val(Txt(TOT_TPAmt)) & "," & Val(Txt(LabAmt)) & "," & Val(Txt(Ser_Amt)) & "," & Val(Txt(MiscAmt)) & "," & Val(Txt(SplAmt)) & "," _
                    & "" & Val(Txt(RoundOff)) & "," & Val(Txt(NetAmt)) & "," _
                    & "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & AddFlg & "')"
            GCn.Execute GSQL
        Case "Edit"
            AddFlg = "E"
            
            GSQL = "Update Job_WarBill set v_date=" & ConvertDate(Txt(BillDt)) & ",Party_Ac='" & Txt(Supplier).Tag & "'," _
                    & "Warranty_Ac='" & Txt(WarrAc).Tag & "',SprAmt_TB=" & Val(Txt(SprTB)) & ",SprAmt_TP=" & Val(Txt(SprTP)) & ",OilAmt_TP=" & Val(Txt(OilTP)) & "," _
                    & "RecdLST_TBPer=" & Val(Txt(RecCST_Per)) & ",RecdLST_TPPer=" & Val(Txt(RecLST_Per)) & ",RecdTOT_TBPer=" & Val(Txt(RecTOT_TBPer)) & ",RecdTOT_TPPer=" & Val(Txt(RecTOT_TPPer)) & ",Tax_Per=" & Val(Txt(Lst_Per)) & "," _
                    & "Tax_Sur_Per=" & Val(Txt(Surc_Per)) & ",TOT_Per=" & Val(Txt(TOT_Per)) & ",SrvTax_Per=" & Val(Txt(Ser_Per)) & ",RecdLst_TBAmt=" & Val(Txt(RecCST_Amt)) & ",RecdLst_TPAmt=" & Val(Txt(RecLST_Amt)) & "," _
                    & "RecdTOT_TBAmt=" & Val(Txt(RecTOT_TBAmt)) & ",RecdTOT_TPAmt=" & Val(Txt(RecTOT_TPAmt)) & ",Tax_Amt=" & Val(Txt(Lst_Amt)) & ",Tax_Sur_Amt=" & Val(Txt(Surc_Amt)) & ",TOT_TBAmt=" & Val(Txt(TOT_TBAmt)) & "," _
                    & "TOT_TPAmt=" & Val(Txt(TOT_TPAmt)) & ",Labour_Amt=" & Val(Txt(LabAmt)) & ",SrvTax_Amt=" & Val(Txt(Ser_Amt)) & ",MiscAmt=" & Val(Txt(MiscAmt)) & ",SplAmt=" & Val(Txt(SplAmt)) & "," _
                    & "Rounded=" & Val(Txt(RoundOff)) & ",Total_Amt=" & Val(Txt(NetAmt)) & "," _
                    & "U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & AddFlg & "'" _
                    & " where DocId='" & lblDocId & "'"
            GCn.Execute GSQL
    End Select

    For I = 1 To FGrid1.Rows - 1
    'Div_Code,Site_Code,Year_Prefix,Claim_No,Claim_Type,Srl_No,LST_PerClaim,Surc_PerClaim,TOT_PerClaim,SrvTax_PerClaim
    
        If FGrid1.TextMatrix(I, C_ID) <> "" Then
            GSQL = "Update Job_WARR1 set " _
                    & "WBill_DocId='" & lblDocId & "',Tax_Per=" & Val(Txt(BillNo)) & ",Tax_Sur_Per=" & Val(Txt(BillNo)) & ",TOT_Per= " & Val(Txt(BillNo)) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & AddFlg & "' WHERE DocId = '" & FGrid1.TextMatrix(I, C_ID) & "'"
            GCn.Execute GSQL
        End If
    Next I
    
    GCnFaS.BeginTrans
        'A/c Posting
    '************
        mNarr = "Through Warranti Bill"
        I = 0
        LedgAry(I).SubCode = Txt(Supplier).Tag
        LedgAry(I).AmtDr = Val(Txt(NetAmt))
        LedgAry(I).Narration = mNarr
        LedgAry(I).ContraSub = Txt(WarrAc).Tag
        I = I + 1
        LedgAry(I).SubCode = Txt(WarrAc).Tag
        LedgAry(I).AmtCr = Val(Txt(NetAmt))
        LedgAry(I).Narration = mNarr
        LedgAry(I).ContraSub = Txt(Supplier).Tag
        
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, lblDocId, CDate(Txt(BillDt)))
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
    'EOF Posting

    GCnFaS.CommitTrans
    GCn.CommitTrans
    
    mTrans = False
    
    Master.Requery
    Call UpdRequery
    
    Master.FIND "Docid = '" & lblDocId & "'"
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
        Case Supplier, WarrAc
            If Txt(Index).TEXT <> "" And Txt(Index).Tag <> RsParty!Code Then
                RsParty.MoveFirst
                RsParty.FIND ("Code='" & Txt(Index).Tag & "'")
            End If
        Case RecCST_Per, RecLST_Per, RecTOT_TBPer, RecTOT_TPPer, Lst_Per, Surc_Per, TOT_Per, Ser_Per
            SendKeys "{Home}+{End}"
        Case RecCST_Amt, RecLST_Amt, RecTOT_TBAmt, RecTOT_TPAmt, Lst_Amt, Surc_Amt, TOT_TBAmt, TOT_TPAmt, Ser_Amt
            SendKeys "{Home}+{End}"
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case Supplier
            DGridColSwap DGParty, 1
            DGridTxtKeyDown DGParty, Txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
        Case WarrAc
            DGridColSwap DGParty, 1
            DGridTxtKeyDown DGParty, Txt, Index, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
    End Select
    If DGParty.Visible = False And DGClaim.Visible = False Then
        '' KEY DOWN
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
            If Index <> Ser_Amt Then
                Ctrl_DownKeyDown KeyCode, Shift
            End If
            If Index = Ser_Amt Then
                If MsgBox("Save Entry ?", vbInformation + vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave: Exit Sub
            End If
        End If
        
        ' KEY UP
        If TopCtrl1.TopText2 = "Add" Then
            If (Txt(BillNo).Enabled = True And Index <> BillNo) Or (Txt(BillNo).Enabled = False And Index <> BillDt) Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then
            If Index <> BillDt Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
    Call CheckQuote(keyascii)
    Select Case Index
        Case BillNo
            Call NumPress(Txt(Index), keyascii, 8, 0)
        Case RecCST_Per, RecLST_Per, RecTOT_TBPer, RecTOT_TPPer, Lst_Per, Surc_Per, TOT_Per, Ser_Per
            Call NumPress(Txt(Index), keyascii, 2, 2)
            Call Txt_Validate(Index, True)
        Case RecCST_Amt, RecLST_Amt, RecTOT_TBAmt, RecTOT_TPAmt, Lst_Amt, Surc_Amt, TOT_TBAmt, TOT_TPAmt, Ser_Amt
            Call NumPress(Txt(Index), keyascii, 8, 2)
            Call Txt_Validate(Index, True)
        Case Supplier, WarrAc
            DGridTxtKeyPress Txt, Index, RsParty, keyascii, "name"
    End Select
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case RecCST_Per, RecLST_Per, RecTOT_TBPer, RecTOT_TPPer, Lst_Per, Surc_Per, TOT_Per, Ser_Per
            Call Txt_Validate(Index, True)
        Case RecCST_Amt, RecLST_Amt, RecTOT_TBAmt, RecTOT_TPAmt, Lst_Amt, Surc_Amt, TOT_TBAmt, TOT_TPAmt, Ser_Amt
            Call Txt_Validate(Index, True)
    End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case BillNo
            lblPrefix.CAPTION = XNull(lblPrefix.CAPTION)
            lblDocId = GetDocID(GCnFaS, VType, Txt(BillDt).TEXT, VoucherEditFlag, Txt(BillNo), lblPrefix, ForSiteCode)
            If VoucherEditFlag = True Then    ' Manual
                Set Rst = New ADODB.Recordset
                Rst.CursorLocation = adUseClient
                Rst.Open "Select Docid From Job_warbill Where DocID='" & lblDocId & "'", GCn, adOpenDynamic, adLockOptimistic
                If Rst.RecordCount > 0 Then
                    MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                    If Txt(BillNo).Enabled = True Then Txt(BillNo).SetFocus
                End If
            End If
        Case BillDt
            Txt(Index) = RetDate(Txt(Index))
        Case Supplier, WarrAc
            If RsParty.RecordCount = 0 Or RsParty.BOF = True Or RsParty.EOF = True Or Txt(Index).TEXT = "" Then Exit Sub
            If Txt(Index).Tag <> RsParty!Code Then
                RsParty.Sort = "CODE"
                RsParty.FIND ("CODE='" & Txt(Index).Tag & "'")
            End If
        Case RecCST_Per
            If Val(Txt(RecCST_Per)) <> 0 Then
                Txt(RecCST_Amt) = Format(Val(Txt(SprTB)) * Val(Txt(RecCST_Per)) / 100, "0.00")
            Else
                Txt(RecCST_Amt) = ""
            End If
        Case RecTOT_TBPer
            If Val(Txt(RecTOT_TBPer)) <> 0 Then
                Txt(RecTOT_TBAmt) = Format((Val(Txt(SprTB)) + Val(Txt(RecCST_Amt))) * Val(Txt(RecTOT_TBPer)) / 100, "0.00")
            Else
                Txt(RecTOT_TBAmt) = ""
            End If
        Case Lst_Per
            If Val(Txt(Lst_Per)) <> 0 Then
                Txt(Lst_Amt) = Format((Val(Txt(SprTB)) + Val(Txt(RecTOT_TBAmt)) + Val(Txt(RecCST_Amt))) * Val(Txt(Lst_Per)) / 100, "0.00")
            Else
                Txt(Lst_Amt) = ""
            End If
        Case Surc_Per
            If Val(Txt(Surc_Per)) <> 0 Then
                Txt(Surc_Amt) = Format(Val(Txt(Lst_Amt)) * Val(Txt(Surc_Per)) / 100, "0.00")
            Else
                Txt(Surc_Amt) = ""
            End If
        Case TOT_Per
            If Val(Txt(TOT_Per)) <> 0 Then
                Txt(TOT_TBAmt) = Format((Val(Txt(Lst_Amt)) + Val(Txt(Surc_Amt)) + Val(Txt(SprTB)) + Val(Txt(RecTOT_TBAmt)) + Val(Txt(RecCST_Amt))) * Val(Txt(TOT_Per)) / 100, "0.00")
                Txt(TOT_TPAmt) = Format((Val(Txt(SprTP)) + Val(Txt(OilTP)) + Val(Txt(RecLST_Amt)) + Val(Txt(RecTOT_TPAmt))) * Val(Txt(TOT_Per)) / 100, "0.00")
            Else
                Txt(TOT_TBAmt) = ""
                Txt(TOT_TPAmt) = ""
            End If
        Case RecLST_Per
            If Val(Txt(RecLST_Per)) <> 0 Then
                Txt(RecLST_Amt) = Format((Val(Txt(SprTP)) + Val(Txt(OilTP))) * Val(Txt(RecLST_Per)) / 100, "0.00")
            Else
                Txt(RecLST_Amt) = ""
            End If
        Case RecTOT_TPPer
            If Val(Txt(RecTOT_TPPer)) <> 0 Then
                Txt(RecTOT_TPAmt) = Format((Val(Txt(SprTP)) + Val(Txt(OilTP)) + Val(Txt(RecLST_Amt))) * Val(Txt(RecTOT_TPPer)) / 100, "0.00")
            Else
                Txt(RecTOT_TPAmt) = ""
            End If
        Case Ser_Per
            If Val(Txt(Ser_Per)) <> 0 Then
                Txt(Ser_Amt) = Format(Val(Txt(LabAmt)) * Val(Txt(Ser_Per)) / 100, "0.00")
            Else
                Txt(Ser_Amt) = ""
            End If
    End Select
    Call Calc_FooterAmt
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    For I = 1 To Txt.Count
        Txt(I).TEXT = ""
        Txt(I).Tag = ""
    Next I
    
    lblDocId.CAPTION = ""
    lblDocId.Refresh
    
    lblPrefix.CAPTION = ""
    lblPrefix.Refresh
    
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
    txtDisabled_Color Me
End Sub

Private Sub MoveRec()
Dim Rs As Recordset
Dim mVor As String
Dim I As Integer
On Error GoTo error1
'    TopCtrl1.tEdit = False
    If Master.RecordCount > 0 Then

        LblDiv.CAPTION = "Division : " & DeCodeDocID(Master!DocID, Division_Code)
        LblSite.CAPTION = "Site Code : " & DeCodeDocID(Master!DocID, For_Site_Code)
        lblDocId.CAPTION = Master!DocID
        lblPrefix.CAPTION = DeCodeDocID(Master!DocID, Document_Prefix)
        
        Txt(BillNo).TEXT = Master!V_NO
        Txt(BillDt).TEXT = Master!V_DATE
        
        Txt(Supplier).TEXT = Master!SuppName
        Txt(Supplier).Tag = Master!Party_Ac
        Txt(WarrAc).TEXT = Master!WarrName
        Txt(WarrAc).Tag = Master!Warranty_Ac
        
        Txt(SprTB).TEXT = Master!SprAmt_TB
        Txt(SprTP).TEXT = Master!SprAmt_TP
        Txt(OilTP).TEXT = Master!OilAmt_TP
        
        Txt(RecCST_Per).TEXT = Master!RecdLST_TBPer
        Txt(RecLST_Per).TEXT = Master!RecdLST_TPPer
        Txt(RecTOT_TBPer).TEXT = Master!RecdTOT_TBPer
        Txt(RecTOT_TPPer).TEXT = Master!RecdTOT_TPPer
        Txt(Lst_Per).TEXT = Master!Tax_Per
        Txt(Surc_Per).TEXT = Master!Tax_Sur_Per
        Txt(TOT_Per).TEXT = Master!TOT_Per
        Txt(Ser_Per).TEXT = Master!SrvTax_Per
        
        Txt(RecCST_Amt).TEXT = Master!recdLst_tbamt
        Txt(RecLST_Amt).TEXT = Master!recdLst_tPamt
        Txt(RecTOT_TBAmt).TEXT = Master!recdtot_tbamt
        Txt(RecTOT_TPAmt).TEXT = Master!recdtot_tpamt
        Txt(Lst_Amt).TEXT = Master!Tax_Amt
        Txt(Surc_Amt).TEXT = Master!Tax_Sur_Amt
        Txt(TOT_TBAmt).TEXT = Master!TOT_TBAmt
        Txt(TOT_TPAmt).TEXT = Master!TOT_TPAmt
        Txt(Ser_Amt).TEXT = Master!SrvTax_amt
        
        Txt(LabAmt).TEXT = Master!Labour_Amt
        Txt(MiscAmt).TEXT = Master!MiscAmt
        Txt(SplAmt).TEXT = Master!SplAmt
        
        Txt(RoundOff).TEXT = Master!Rounded
        Txt(NetAmt).TEXT = Master!Total_Amt
        
        Call Fill_Grid
        Call Calc_FooterAmt
    Else
        Call BlankText
    End If
    Grid_Hide
    Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
    With FGrid1
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 10
        
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 400
          
        .TextMatrix(0, C_ClmNo) = "Prow.No"
        .ColAlignment(C_ClmNo) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmNo) = flexAlignRightCenter
        .ColWidth(C_ClmNo) = 1100
    
        .TextMatrix(0, C_ClmDt) = "Prow.Date"
        .ColAlignment(C_ClmDt) = flexAlignRightCenter
        .ColAlignmentFixed(C_ClmDt) = flexAlignRightCenter
        .ColWidth(C_ClmDt) = 1100
    
        .TextMatrix(0, C_SprTB) = "TB Spare"
        .ColAlignment(C_SprTB) = flexAlignRightCenter
        .ColAlignmentFixed(C_SprTB) = flexAlignRightCenter
        .ColWidth(C_SprTB) = 1100
    
        .TextMatrix(0, C_SprTP) = "TP Spare"
        .ColAlignment(C_SprTP) = flexAlignRightCenter
        .ColAlignmentFixed(C_SprTP) = flexAlignRightCenter
        .ColWidth(C_SprTP) = 1100
    
        .TextMatrix(0, C_OilTP) = "TP Oil"
        .ColAlignment(C_OilTP) = flexAlignRightCenter
        .ColAlignmentFixed(C_OilTP) = flexAlignRightCenter
        .ColWidth(C_OilTP) = 1100
    
        .TextMatrix(0, C_Misc) = "Misc.Chrg."
        .ColAlignment(C_Misc) = flexAlignRightCenter
        .ColAlignmentFixed(C_Misc) = flexAlignRightCenter
        .ColWidth(C_Misc) = 1100
    
        .TextMatrix(0, C_Labour) = "Labour Amt"
        .ColAlignment(C_Labour) = flexAlignRightCenter
        .ColAlignmentFixed(C_Labour) = flexAlignRightCenter
        .ColWidth(C_Labour) = 1100
    
        .TextMatrix(0, C_Spl) = "Spl.Chrg."
        .ColAlignment(C_Spl) = flexAlignRightCenter
        .ColAlignmentFixed(C_Spl) = flexAlignRightCenter
        .ColWidth(C_Spl) = 1100
    
        .TextMatrix(0, C_ID) = "ClaimId"
        .ColAlignment(C_ID) = flexAlignRightCenter
        .ColAlignmentFixed(C_ID) = flexAlignRightCenter
        .ColWidth(C_ID) = 0
    End With
    DGParty.width = 4740: DGParty.left = Shape2.left: DGParty.top = Shape2.top: DGParty.height = 5000
    DGClaim.left = Me.left + 125: DGClaim.top = 3960: DGClaim.height = 2655
End Sub

Public Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    For I = 1 To Txt.Count
        Txt(I).Enabled = Enb
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
        'txt(i).MaxLength = 0
    Next
    Txt(SprTB).Enabled = False
    Txt(SprTP).Enabled = False
    Txt(OilTP).Enabled = False
    Txt(SubATB).Enabled = False
    Txt(SubATP).Enabled = False
    Txt(SubB).Enabled = False
    Txt(LabAmt).Enabled = False
    Txt(MiscAmt).Enabled = False
    Txt(SplAmt).Enabled = False
    Txt(RoundOff).Enabled = False
    Txt(NetAmt).Enabled = False
    Txt(RecLST_Amt).Enabled = False
    Txt(RecCST_Amt).Enabled = False
    Txt(RecTOT_TBAmt).Enabled = False
    Txt(RecTOT_TPAmt).Enabled = False
    Txt(Lst_Amt).Enabled = False
    Txt(Surc_Amt).Enabled = False
    Txt(TOT_TBAmt).Enabled = False
    Txt(TOT_TPAmt).Enabled = False
    Txt(Ser_Amt).Enabled = False
    txtDisabled_Color Me
End Sub

Private Sub Grid_Hide()
    If DGParty.Visible = True Then DGParty.Visible = False
    If DGClaim.Visible = True Then DGClaim.Visible = False
End Sub

Private Sub UpdRequery()
    RsClaim.Requery
    RsParty.Requery
End Sub

Private Sub FGrid1_Click()
    txtgrid1(0).Visible = False
End Sub

Private Sub FGrid1_DblClick()
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid1.Col
        Case C_ClmNo
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
        SendKeysA vbKeyTab, True
        FGrid1.SetFocus
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid1.Tag = FGrid1.Row
    Select Case FGrid1.Col
        Case C_ClmNo
            If KeyCode = vbKeyDelete And Shift = 0 Then
                FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
            End If
            If KeyCode = vbKeyReturn Then
                GridDblClick Me, FGrid1, txtgrid1, 0
            End If
    End Select
    TAddMode = False
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyPress(keyascii As Integer)
On Error GoTo ELoop
    Select Case FGrid1.Col
        Case C_ClmNo
            Get_Text Me, FGrid1, txtgrid1, 0, False, keyascii
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
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid1.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_LeaveCell()
    FGrid1.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid1_Scroll()
    txtgrid1(0).Visible = False
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
    Select Case FGrid1.Col
        Case C_ClmNo
            If RsClaim.RecordCount = 0 Then Exit Sub
            RsClaim.MoveFirst
            If txtgrid1(Index).TEXT = "" Then Exit Sub
            RsClaim.FIND "Code='" & FGrid1.TextMatrix(FGrid1.Row, C_ID) & "'"
    End Select
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
        Case C_ClmNo
            DGridTxtKeyDown DGClaim, txtgrid1, 0, RsClaim, KeyCode, True, 1
            If KeyCode = vbKeyReturn Then
                If TxtGrid1Leave = True Then
                     GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, C_ClmNo
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
    Call CheckQuote(keyascii)
    Select Case Index
    Case 0
        Select Case FGrid1.Col
            Case C_ClmNo
                If DGClaim.Visible = True Then DGridTxtKeyPress txtgrid1, Index, RsClaim, keyascii, "prowno"
        End Select
    End Select

End Sub

Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
    Case 0
        Select Case FGrid1.Col
            Case C_ClmNo
                If KeyCode <> 13 And DGClaim.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0: DGridTxtKeyPress txtgrid1, Index, RsClaim, KeyCode, "prowno", True
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
        Case C_ClmNo
            If RsClaim.EOF = True Or RsClaim.BOF = True Or txtgrid1(0).TEXT = "" Then
                Call Blank_Cells
            Else
                For I = 1 To FGrid1.Rows - 1
                    If FGrid1.TextMatrix(I, C_ID) = RsClaim!Code And I <> FGrid1.Row Then
                        MsgBox "Duplicate Claim No. Not Allowed", vbInformation, "Validation"
                        GoTo NXT
                    End If
                Next I
                FGrid1.TextMatrix(FGrid1.Row, C_ClmNo) = RsClaim!ProwNo
                FGrid1.TextMatrix(FGrid1.Row, C_ClmDt) = RsClaim!ProwDt
                FGrid1.TextMatrix(FGrid1.Row, C_ID) = RsClaim!Code
                Call Update_Cells(FGrid1.Row)
            End If
            Calc_FgridTotal
    End Select
NXT:
    txtgrid1(0).MaxLength = 10
    Exit Sub
ELoop:
    CheckError
End Sub
Private Function TxtGrid1Leave() As Boolean
Dim I As Integer
    Select Case FGrid1.Col
        Case C_ClmNo
            If RsClaim.EOF = True Or RsClaim.BOF = True Then
                Call Blank_Cells
            Else
                For I = 1 To FGrid1.Rows - 1
                    If FGrid1.TextMatrix(I, C_ID) = RsClaim!Code And I <> FGrid1.Row Then
                        MsgBox "Duplicate Claim No. Not Allowed", vbInformation, "Validation"
                        GoTo NXT
                    End If
                Next I
                FGrid1.TextMatrix(FGrid1.Row, C_ClmNo) = RsClaim!ProwNo
                FGrid1.TextMatrix(FGrid1.Row, C_ClmDt) = RsClaim!ProwDt
                FGrid1.TextMatrix(FGrid1.Row, C_ID) = RsClaim!Code
                Call Update_Cells(FGrid1.Row)
            End If
            Calc_FgridTotal
    End Select
NXT:
    txtgrid1(0).MaxLength = 10
    txtgrid1(0).Visible = False
    ExitCtrl = True
    TxtGrid1Leave = True
    FGrid1.SetFocus
End Function

Private Sub Fill_Grid()
Dim MyRst As ADODB.Recordset
Dim I As Integer
    FGrid1.Rows = 1
    Set MyRst = New ADODB.Recordset
    MyRst.CursorLocation = adUseClient
    
    GSQL = "Select JW1.* From Job_warr1 as jw1 Where Jw1.WBill_DocId='" & Master!DocID & "' order By DocID"
    MyRst.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    I = 1
    If MyRst.RecordCount > 0 Then
        Do Until MyRst.EOF
            FGrid1.AddItem ""
            With FGrid1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, C_ClmNo) = MyRst!ProwNo
                .TextMatrix(I, C_ClmDt) = MyRst!ProwDt
                .TextMatrix(I, C_ID) = MyRst!DocID
            End With
            Call Update_Cells(I)
            I = I + 1
            MyRst.MoveNext
        Loop
        FGrid1.AddItem FGrid1.Rows
        FGrid1.FixedRows = 1
    Else
        FGrid1.Rows = FGrid1.Rows
        FGrid1.AddItem FGrid1.Rows
        FGrid1.FixedRows = 1
    End If
    Set MyRst = Nothing
End Sub
Private Sub Blank_Cells()
    FGrid1.TextMatrix(FGrid1.Row, C_ClmNo) = ""
    FGrid1.TextMatrix(FGrid1.Row, C_ID) = ""
    FGrid1.TextMatrix(FGrid1.Row, C_ClmDt) = ""
    FGrid1.TextMatrix(FGrid1.Row, C_SprTB) = ""
    FGrid1.TextMatrix(FGrid1.Row, C_SprTP) = ""
    FGrid1.TextMatrix(FGrid1.Row, C_OilTP) = ""
    FGrid1.TextMatrix(FGrid1.Row, C_Labour) = ""
    FGrid1.TextMatrix(FGrid1.Row, C_Misc) = ""
    FGrid1.TextMatrix(FGrid1.Row, C_Spl) = ""
End Sub
Private Sub Update_Cells(ByVal RowNo As Integer)
Dim MyRst As ADODB.Recordset
Dim I As Integer
Dim Amt1 As Integer, Amt2 As Integer, Amt3 As Integer, Amt4 As Integer, Amt5 As Integer, Amt6 As Integer
''  TB Spr              TP Spr          OIL Amt         Misc Amt            Labour          Spl
    
    If FGrid1.TextMatrix(FGrid1.Row, C_ClmNo) = "" Then Exit Sub
    Set MyRst = New ADODB.Recordset
    MyRst.CursorLocation = adUseClient
    GSQL = "Select JW2.*,Part.part_Grade From Job_warr2 as jw2 left join part on jw2.part_no=part.part_no and Part.Div_Code = left(jw2.docid,1)  Where Jw2.docid='" & FGrid1.TextMatrix(RowNo, C_ID) & "'"
    MyRst.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    
    If MyRst.RecordCount > 0 Then
        Do Until MyRst.EOF
            If MyRst!Part_Grade <> PubPartGrade_Lub Then
                If MyRst!Tax_YN = 0 Then
                    Amt2 = Amt2 + (MyRst!Price * MyRst!TotQty)
                Else
                    Amt1 = Amt1 + (MyRst!Price * MyRst!TotQty)
                End If
            Else
                Amt3 = Amt3 + (MyRst!Price * MyRst!TotQty)
            End If
            Amt4 = Amt4 + MyRst!Misc_Chrg
            Amt5 = Amt5 + MyRst!Labour_Amt
            Amt6 = Amt6 + MyRst!Spl_Amt
            MyRst.MoveNext
        Loop
    Else
        Amt1 = 0
        Amt2 = 0
        Amt3 = 0
        Amt4 = 0
        Amt5 = 0
        Amt6 = 0
    End If

    With FGrid1
        .TextMatrix(RowNo, C_SprTB) = Format(Amt1, "0.00")
        .TextMatrix(RowNo, C_SprTP) = Format(Amt2, "0.00")
        .TextMatrix(RowNo, C_OilTP) = Format(Amt3, "0.00")
        .TextMatrix(RowNo, C_Misc) = Format(Amt4, "0.00")
        .TextMatrix(RowNo, C_Labour) = Format(Amt5, "0.00")
        .TextMatrix(RowNo, C_Spl) = Format(Amt6, "0.00")
    End With
    Set MyRst = Nothing
End Sub

Private Sub Calc_FgridTotal()
Dim I As Integer
    Txt(SprTB) = ""
    Txt(SprTP) = ""
    Txt(OilTP) = ""
    Txt(LabAmt) = ""
    Txt(MiscAmt) = ""
    Txt(SplAmt) = ""
    For I = 1 To FGrid1.Rows - 1
        Txt(SprTB) = Format(Val(Txt(SprTB)) + Val(FGrid1.TextMatrix(I, C_SprTB)), "0.00")
        Txt(SprTP) = Format(Val(Txt(SprTP)) + Val(FGrid1.TextMatrix(I, C_SprTP)), "0.00")
        Txt(OilTP) = Format(Val(Txt(OilTP)) + Val(FGrid1.TextMatrix(I, C_OilTP)), "0.00")
        Txt(LabAmt) = Format(Val(Txt(LabAmt)) + Val(FGrid1.TextMatrix(I, C_Labour)), "0.00")
        Txt(MiscAmt) = Format(Val(Txt(MiscAmt)) + Val(FGrid1.TextMatrix(I, C_Misc)), "0.00")
        Txt(SplAmt) = Format(Val(Txt(SplAmt)) + Val(FGrid1.TextMatrix(I, C_Spl)), "0.00")
    Next I
    Call Calc_FooterAmt
End Sub

Private Sub Calc_FooterAmt()
Dim MyVar As Double
    Txt(SubATB) = Format(Val(Txt(Lst_Amt)) + Val(Txt(Surc_Amt)) + Val(Txt(SprTB)) + Val(Txt(RecTOT_TBAmt)) + Val(Txt(RecCST_Amt)) + Val(Txt(TOT_TBAmt)), "0.00")
    Txt(SubATP) = Format(Val(Txt(SprTP)) + Val(Txt(OilTP)) + Val(Txt(RecLST_Amt)) + Val(Txt(RecTOT_TPAmt)) + Val(Txt(TOT_TPAmt)), "0.00")
    MyVar = Val(Txt(SubATB)) + Val(Txt(SubATP)) + Val(Txt(LabAmt)) + Val(Txt(Ser_Amt)) + Val(Txt(MiscAmt)) + Val(Txt(SplAmt))
    Txt(RoundOff) = dmRoundOff(MyVar, 0)
    Txt(SubB) = Format(Val(Txt(LabAmt)) + Val(Txt(Ser_Amt)) + Val(Txt(MiscAmt)) + Val(Txt(SplAmt)) + Val(Txt(RoundOff)), "0.00")
    Txt(NetAmt) = Format(MyVar + Val(Txt(RoundOff)), "0.00")
End Sub
