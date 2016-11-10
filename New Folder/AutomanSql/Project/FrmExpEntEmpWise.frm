VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmExpEntEmpWise 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Employee Wise Expence Entry"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11535
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   7785
      MaxLength       =   40
      TabIndex        =   29
      Top             =   4365
      Width           =   1365
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   6345
      MaxLength       =   40
      TabIndex        =   28
      Top             =   4365
      Width           =   1365
   End
   Begin VB.TextBox txtgrid 
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
      Index           =   0
      Left            =   4860
      MaxLength       =   40
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSDataGridLib.DataGrid DgVPrefix 
      Height          =   2730
      Left            =   2940
      Negotiate       =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7290
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4815
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
      RowDividerStyle =   1
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   "Voucher Prefix "
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
            ColumnWidth     =   3420.284
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgVType 
      Height          =   2730
      Left            =   1590
      Negotiate       =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4815
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
      RowDividerStyle =   1
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   "Voucher Type"
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
            ColumnWidth     =   3420.284
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
      Height          =   210
      Index           =   7
      Left            =   7995
      MaxLength       =   40
      TabIndex        =   8
      Top             =   6255
      Visible         =   0   'False
      Width           =   1755
   End
   Begin MSDataGridLib.DataGrid DgCashBankAc 
      Height          =   2730
      Left            =   1320
      Negotiate       =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6765
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4815
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
      RowDividerStyle =   1
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   "Cash/Bank A/c Name"
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
            ColumnWidth     =   3420.284
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
      Height          =   210
      Index           =   6
      Left            =   10170
      MaxLength       =   40
      TabIndex        =   3
      Top             =   735
      Width           =   1245
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
      Index           =   5
      Left            =   7740
      MaxLength       =   40
      TabIndex        =   2
      Top             =   735
      Width           =   1245
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
      Index           =   4
      Left            =   3285
      MaxLength       =   40
      TabIndex        =   1
      Top             =   735
      Width           =   2955
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
      Index           =   3
      Left            =   570
      MaxLength       =   40
      TabIndex        =   0
      Top             =   735
      Width           =   1365
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
      Height          =   945
      Index           =   2
      Left            =   1035
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   4710
      Width           =   4500
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
      Index           =   1
      Left            =   7995
      MaxLength       =   40
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   5205
   End
   Begin VB.TextBox txtgrid1 
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
      Index           =   0
      Left            =   9660
      MaxLength       =   40
      TabIndex        =   4
      Top             =   5070
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   -1965
      TabIndex        =   12
      Top             =   6855
      Visible         =   0   'False
      Width           =   2505
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   75
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   30
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3228
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
      Height          =   210
      Index           =   0
      Left            =   7995
      MaxLength       =   40
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   5205
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   661
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   1695
      Left            =   5670
      TabIndex        =   5
      Top             =   4710
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   12632319
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      GridColorFixed  =   0
      FocusRect       =   0
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "ddd"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
   Begin MSDataGridLib.DataGrid DgAcName 
      Height          =   2730
      Left            =   1110
      Negotiate       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6675
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4815
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
      RowDividerStyle =   1
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "name"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   3420.284
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgEmp 
      Height          =   2730
      Left            =   1515
      Negotiate       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6930
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4815
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
      RowDividerStyle =   1
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "name"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   3420.284
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3150
      Left            =   0
      TabIndex        =   7
      Top             =   1095
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5556
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   12632319
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   0
      GridColorFixed  =   0
      FocusRect       =   0
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "ddd"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
   Begin VB.Line Line1 
      X1              =   8190
      X2              =   8190
      Y1              =   1095
      Y2              =   2880
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount ......................"
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
      Left            =   6195
      TabIndex        =   25
      Top             =   6285
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher No. ........"
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
      Left            =   9105
      TabIndex        =   23
      Top             =   750
      Width           =   1590
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Prefix....."
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
      Left            =   6375
      TabIndex        =   22
      Top             =   750
      Width           =   1560
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Type....."
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
      Left            =   1980
      TabIndex        =   21
      Top             =   750
      Width           =   1485
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date................."
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
      Left            =   45
      TabIndex        =   20
      Top             =   750
      Width           =   1425
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration ....................."
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
      Left            =   30
      TabIndex        =   19
      Top             =   4680
      Width           =   2115
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash / Bank A/c ......."
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
      Left            =   6210
      TabIndex        =   18
      Top             =   6750
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expencence A/c .......... "
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
      Left            =   6210
      TabIndex        =   17
      Top             =   6510
      Visible         =   0   'False
      Width           =   2085
   End
End
Attribute VB_Name = "FrmExpEntEmpWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte

Dim mDocId$
Dim VoucherEditFlag As Boolean
Dim RsVType         As ADODB.Recordset
Dim RsVPrefix       As ADODB.Recordset
Dim RstMain         As ADODB.Recordset
Dim RsAcName         As ADODB.Recordset
Dim RsCashBankAc    As ADODB.Recordset
Dim RsEmp           As ADODB.Recordset

Dim mFlag As Byte
Dim GridKey As Integer
Dim RsTrb As ADODB.Recordset

Dim mGridBackColor As String
Dim mGridBackColorSel As String




Private Const Col_Emp_Desc  As Byte = 1
Private Const Col_Amount    As Byte = 2
Private Const Col_Emp_Code  As Byte = 3
Private Const Col_AcCode    As Byte = 4


''''''''Fgrid Constants'''''''''''''
Private Const F_DrCr As Byte = 1
Private Const F_AcName As Byte = 2
Private Const F_Balance As Byte = 3
Private Const F_AmtDr As Byte = 4
Private Const F_AmtCr As Byte = 5
Private Const F_Narration As Byte = 6
Private Const F_ChqNo As Byte = 7
Private Const F_ChqDate As Byte = 8
Private Const F_AcCode As Byte = 9
Private Const F_EmpDetailYn As Byte = 10
Private Const F_Nature As Byte = 11

Private Const T_ExpAc       As Byte = 0
Private Const T_CashBankAc  As Byte = 1
Private Const T_Narration   As Byte = 2
Private Const T_Date        As Byte = 3
Private Const T_VType       As Byte = 4
Private Const T_VPrefix     As Byte = 5
Private Const T_VNo         As Byte = 6
Private Const T_Amount      As Byte = 7
Private Const T_AmtDr       As Byte = 8
Private Const T_AmtCr       As Byte = 9


Private Sub FGrid_RowColChange()
    ShowEmpAcWise
End Sub
Sub ShowEmpAcWise()
Dim I As Integer
Dim mCnt As Integer
    I = 0: mCnt = 0
    For I = 1 To FGrid1.Rows - 1
        If StrCmp(FGrid.TextMatrix(FGrid.Row, F_AcCode), FGrid1.TextMatrix(I, Col_AcCode)) Or FGrid1.TextMatrix(I, Col_AcCode) = "" Then
            FGrid1.RowHeight(I) = 240
        Else
            FGrid1.RowHeight(I) = 0
        End If
    Next I
    If FGrid1.TextMatrix(FGrid1.Rows - 1, Col_Emp_Desc) <> "" Then
        FGrid1.AddItem ""
    End If

End Sub

Private Sub FGrid1_Click()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    txtgrid1(0).Visible = False
End Sub


Private Sub FGrid1_DblClick()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.TextMatrix(FGrid.Row, F_AcCode) = "" Or Not StrCmp(FGrid.TextMatrix(FGrid.Row, F_EmpDetailYn), "Y") Then Exit Sub
    Select Case FGrid1.Col
        Case Col_Emp_Desc, Col_Amount
            Call GridDblClick(Me, FGrid1, txtgrid1, 0)
    End Select
End Sub

Private Sub FGrid1_EnterCell()
    'FGrid1.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid1_GotFocus()
    FGrid1.BackColorSel = mGridBackColorSel   'FaBackColorSelEnter

    'FGrid1.Col = Col_Emp_Desc
    txtgrid1(0).Visible = False
    AmtCal
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid1.Tag) = FGrid1.Rows - 1 Then
    If MsgBox("Do You Want to Save?", vbYesNo) = vbYes Then TopCtrl1_eSave
'    SendKeysA vbKeyTab, True
'    KeyCode = 0
End If
GridKey = KeyCode
FGrid1.Tag = FGrid1.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid1.Col
        Case Col_Amount
            FGrid1 = ""
        Case Col_Emp_Desc
            FGrid1 = ""
            FGrid1.TextMatrix(FGrid1.Row, Col_Emp_Code) = ""
    End Select
End If
If KeyCode = vbKeyReturn Then
    If FGrid.TextMatrix(FGrid.Row, F_AcCode) = "" Or Not StrCmp(FGrid.TextMatrix(FGrid.Row, F_EmpDetailYn), "Y") Then Exit Sub
    Select Case FGrid1.Col
        Case Col_Emp_Code, Col_Amount
            Call GridDblClick(Me, FGrid1, txtgrid1, 0)
            
    End Select
End If
KeyCode = 0

End Sub

Private Sub FGrid1_KeyPress(keyascii As Integer)
    If FGrid.TextMatrix(FGrid.Row, F_AcCode) = "" Or Not StrCmp(FGrid.TextMatrix(FGrid.Row, F_EmpDetailYn), "Y") Then Exit Sub
    Select Case FGrid1.Col
        Case Col_Emp_Desc
           Call Get_Text(Me, FGrid1, txtgrid1, 0, False, keyascii)
        Case Col_Amount
            Call Get_Text(Me, FGrid1, txtgrid1, 0, True, keyascii)
    End Select
End Sub

Private Sub FGrid1_LostFocus()
FGrid1.BackColorSel = mGridBackColor   'FaCellBackColLeave1

FGrid1_Validate (True)
End Sub

Private Sub FGrid1_RowColChange()
    AmtCal
End Sub

Private Sub FGrid1_Scroll()
    txtgrid1(0).Visible = False

End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid1.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid1.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
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
AmtCal
Exit Sub
End Sub

Private Sub FGrid1_Validate(Cancel As Boolean)
'    FGrid1.CellBackColor = CellBackColLeave
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


WinSetting Me
TopCtrl1.Tag = PubUParam
Ini_Grid




Set RstMain = New ADODB.Recordset
RstMain.Open "Select Distinct EE.DocId as SearchCode, EE.V_Date From Ledger EE Left Join Voucher_Type vt On EE.V_Type = Vt.V_Type Where Category = 'EXP' Order by EE.V_Date, EE.DocId", G_FaCn, adOpenDynamic, adLockOptimistic

Set RsVType = GCn.Execute("Select V_Type As Code, Description As Name FROM Voucher_Type Where Category In ('EXP') Order by Description")
Set DGVType.DataSource = RsVType

Set RsVPrefix = GCn.Execute("Select Prefix As Code, Prefix As Name, VP.V_Type FROM Voucher_Prefix VP Left Join Voucher_Type VT On VP.V_Type = VT.V_Type Where VT.NCat In ('JV') And Date_From<= " & ConvertDate(txt(T_Date)) & " And Date_To>= " & ConvertDate(txt(T_Date)) & " Order by Prefix")
Set DgVPrefix.DataSource = RsVPrefix

Set RsAcName = GCn.Execute("Select SubCode As Code, Name, Curr_Bal, Nature, EmpDetailYn FROM SubGroup Order by Name")
Set DgAcName.DataSource = RsAcName

Set RsCashBankAc = GCn.Execute("Select SubCode As Code, Name FROM SubGroup Where Nature In ('Cash', 'Bank') Order by Name")
Set DgCashBankAc.DataSource = RsCashBankAc

Set RsEmp = GCn.Execute("Select Emp_Code As Code, Emp_Name As Name From Emp_Mast Order By Emp_Name")
Set DGEmp.DataSource = RsEmp

Disp_Text SETS("INI", Me, RstMain)
MoveRec
ADDFLAG = 0:    mFlag = 0




Grid_Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set RstMain = Nothing: Set RsAcName = Nothing
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo ErrLoop
BlankText
Disp_Text SETS("ADD", Me, RstMain)

ADDFLAG = 1


FGrid1.Rows = 1
FGrid1.AddItem ""
FGrid1.FixedRows = 1


txt(T_Date).SetFocus



Exit Sub
ErrLoop:    MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo ErrLoop
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    FGrid.SetFocus
    ADDFLAG = 2
    
    FGrid.AddItem ""
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
AmtCal
Exit Sub
ErrLoop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
On Error GoTo ErrLoop
Dim transFalg As Byte
Dim LedgAry(1) As LedgRec, mResult As Byte
transFalg = 0
Dim XBM
Dim Res As Integer
    If RstMain.RecordCount > 0 Then
        If MsgBox("Sure To Delete Record", vbYesNo) = vbYes Then
            XBM = RstMain.Bookmark
            
            GCn.Execute "Delete From Exp_Emp1 Where DocId = '" & mDocId & "'"
            mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, mDocId)
            If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
        
            RstMain.Requery
            
            If RstMain.RecordCount >= XBM Then
                RstMain.Bookmark = XBM
            Else
                If RstMain.EOF = False Then RstMain.MoveLast
            End If
            
            MoveRec
            BUTTONS True, Me, RstMain, 0
        End If
    Else
        MsgBox "No Records To Delete.", vbInformation, "Information"
    End If

Exit Sub
ErrLoop:
    MsgBox err.Description, vbExclamation, " Deletion Error "
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, RstMain, 1
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, RstMain, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, RstMain, 3
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, RstMain, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "Select Distinct EE.DocId as SearchCode, " & cCStr("EE.V_Date") & " As V_Date, V.Description, " & cCStr("EE.V_No") & " As V_No From Ledger EE Left Join Voucher_Type V On EE.V_Type = V.V_Type Where Category='EXP' Order by EE.V_Date, EE.DocId"
    Set SearchForm = Me
    FAFind.IsNonFaFind = False
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    RstMain.MoveFirst
    RstMain.FIND ("SearchCode='" & MyValue & "'")
    BUTTONS True, Me, RstMain, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
'Dim I As Integer, mQRY$, mRepName$
'Dim Rst As ADODB.Recordset
'On Error GoTo ERRORHANDLER
'
'    mRepName = "Exp_Emp"
'    mQRY = "Select EE.ExpAc, EE.Emp_Code, Sum(EE.Amount) As Amount, Max(S.Name) As ExpAcName, Max(M.Name) As Emp_name " & _
'         "from Exp_Emp EE Left Join SubGroup S On S.SubCode = EE.ExpAc " & _
'         "Left Join Emp_Mast M On M.Code = EE.Emp_Code Where EE.Site_Code = '" & PubSiteCode & "' " & _
'         "Group By EE.ExpAc, EE.Emp_Code Order By ExpAcName, EE.Emp_Code "
'
'
'    Set Rst = New Recordset
'    Rst.CursorLocation = adUseClient
'    Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
'    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
'    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
'    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
'    rpt.Database.SetDataSource Rst
'    rpt.ReadRecords
'    Call Report_View(rpt, Me.CAPTION, , False)
'    Set Rst = Nothing
'Exit Sub
'ERRORHANDLER:
'      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub TopCtrl1_eSave()
Dim mTrans As Byte
Dim mSearchCode$
Dim I As Integer
On Error GoTo ELoop

    Dim mDR As Double, mCR As Double
    mDR = 0: mCR = 0
    With FGrid
        For I = 1 To .Rows - 1
            If StrCmp(.TextMatrix(I, F_DrCr), "Dr") And .TextMatrix(I, F_AcName) <> "" Then
                mDR = mDR + Val(.TextMatrix(I, F_AmtDr))
            ElseIf StrCmp(.TextMatrix(I, F_DrCr), "Cr") And .TextMatrix(I, F_AcName) <> "" Then
                mCR = mCR + Val(.TextMatrix(I, F_AmtCr))
            End If
        Next I
    End With
    If mDR <> mCR Or mDR <= 0 Or mCR <= 0 Then
        MsgBox "Amount Dr. Is Not Equal to Amount Cr. Or is Zero"
        Exit Sub
    End If
    

    
        
    If TopCtrl1.TopText2 = "Add" Then
        If GCn.Execute("Select Count(*) From Ledger Where Left(DocId,1)='" & PubDivCode & "' And  V_Type='" & txt(T_VType).Tag & "' And V_Prefix='" & txt(T_VPrefix) & "' And  V_No = " & Val(txt(T_VNo)) & " ").Fields(0).Value > 0 Then
            If VoucherEditFlag Then
                MsgBox "Document No. already exists, Retry", vbCritical, "Validation Error"
                txt(T_VNo).SetFocus
                GoTo ELoop
            Else
                mDocId = GetDocID(GCnFaS, txt(T_VType).Tag, txt(T_Date), VoucherEditFlag, txt(T_VNo), txt(T_VPrefix))
                If Val(txt(T_VNo)) <= Val(DeCodeDocID(mDocId, Document_No)) Then
                    MsgBox "Document No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo ELoop
                End If
            End If
        End If
    End If
    
    
    mTrans = 1
    GCn.BeginTrans
        
        G_FaCn.Execute "Delete From Exp_Emp1 Where DocId = '" & mDocId & "'"
        
        With FGrid1
            For I = 1 To .Rows - 1
                If .TextMatrix(I, Col_Emp_Desc) <> "" Then
                    GCn.Execute "Insert Into Exp_Emp1 (DocId, Srl, SubCode, Emp_Code, Amount) " & _
                                "Values ('" & mDocId & "', " & I & ", '" & .TextMatrix(I, Col_AcCode) & "',  '" & .TextMatrix(I, Col_Emp_Code) & "', " & Val(.TextMatrix(I, Col_Amount)) & ") "
                End If
            Next I
        End With
        AcPost
        If TopCtrl1.TopText2 = "Add" Then
            UpdVouSrlNo G_FaCn, mDocId, CDate(txt(T_Date))
        End If
    
    GCn.CommitTrans
    mTrans = 0
                    
    RstMain.Requery
    RstMain.FIND ("SearchCode = '" & mDocId & "'")

    
    Disp_Text SETS("INI", Me, RstMain)
    MoveRec
    Grid_Hide

Exit Sub
ELoop:
    If mTrans = 1 Then GCn.RollbackTrans
    MsgBox err.Description, vbCritical
End Sub



Sub AcPost()
    Dim LedgAry() As LedgRec
    Dim mResult As Byte, I As Integer, j As Integer

    I = 0
    With FGrid
        For j = 1 To FGrid.Rows - 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = .TextMatrix(j, F_AcCode)
            LedgAry(I).AmtDr = IIf(StrCmp(.TextMatrix(j, F_DrCr), "Dr"), Val(.TextMatrix(j, F_AmtDr)), 0)
            LedgAry(I).AmtCr = IIf(StrCmp(.TextMatrix(j, F_DrCr), "Cr"), Val(.TextMatrix(j, F_AmtCr)), 0)
            LedgAry(I).Narration = .TextMatrix(j, F_Narration)
            LedgAry(I).Chq_No = .TextMatrix(j, F_ChqNo)
            If .TextMatrix(j, F_ChqDate) <> "" Then
                LedgAry(I).Chq_Date = CDate(.TextMatrix(j, F_ChqDate))
            End If
            LedgAry(I).EmpDetailYn = .TextMatrix(j, F_EmpDetailYn)
            
            I = I + 1
        Next j
        
        mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, G_FaCn, mDocId, CDate(txt(T_Date)), txt(T_Narration))
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation": Exit Sub
    End With
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ErrLoop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        Disp_Text SETS("INI", Me, RstMain)
        DoEvents
        MoveRec
    End If
Exit Sub
ErrLoop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub MoveRec()
Dim Rs As Recordset
Dim I As Integer
On Error GoTo ErrLoop


    BlankText
If RstMain.RecordCount > 0 Then
    
    
    
    Set Rs = G_FaCn.Execute("Select L.DocId, L.V_Sno, l.V_Type, L.V_Prefix, L.V_Date, L.V_No, L.SubCode, L.Chq_No, L.Chq_Date, L.EmpDetailYn, " & _
                           "L.Narration, Lm.Narration As NarrationMain, L.AmtDr, L.AmtCr, S.Name, S.Curr_Bal,  S.Nature, Vt.Description As VoucherType " & _
                           "From (((LedgerM Lm " & _
                           "Left Join Ledger L On L.DocId = LM.DocId) " & _
                           "Left Join SubGroup S On L.SubCode = S.SubCode) " & _
                           "Left Join Voucher_Type Vt On L.V_Type = Vt.V_Type) " & _
                           "Where L.DocId  = '" & RstMain!SearchCode & "' Order By L.DocId, L.V_SNo ")
    
    I = 1
    FGrid.Rows = 1
    If Rs.RecordCount > 0 Then
        mDocId = Rs!DocID
        txt(T_Date) = XNull(Rs!V_DATE)
        txt(T_VType).Tag = XNull(Rs!V_Type)
        txt(T_VType) = XNull(Rs!VoucherType)
        txt(T_VPrefix) = XNull(Rs!v_Prefix)
        txt(T_VNo) = XNull(Rs!V_NO)
        txt(T_Narration) = XNull(Rs!NarrationMain)
        
        Do Until Rs.EOF
            FGrid.AddItem ""
            FGrid.TextMatrix(I, 0) = I
            FGrid.TextMatrix(I, F_DrCr) = IIf(VNull(Rs!AmtDr) > 0, "Dr", "Cr")
            FGrid.TextMatrix(I, F_AcCode) = XNull(Rs!SubCode)
            FGrid.TextMatrix(I, F_AcName) = XNull(Rs!Name)
            FGrid.TextMatrix(I, F_Balance) = Abs(VNull(Rs!Curr_Bal)) & IIf(VNull(Rs!Curr_Bal) < 0, " Dr", " Cr")
            FGrid.TextMatrix(I, F_AmtDr) = Format(IIf(VNull(Rs!AmtDr) > 0, VNull(Rs!AmtDr), VNull(Rs!AmtDr)), "0.00")
            FGrid.TextMatrix(I, F_AmtCr) = Format(IIf(VNull(Rs!AmtCr) > 0, VNull(Rs!AmtCr), VNull(Rs!AmtCr)), "0.00")
            FGrid.TextMatrix(I, F_Narration) = XNull(Rs!Narration)
            FGrid.TextMatrix(I, F_Nature) = XNull(Rs!Nature)
            FGrid.TextMatrix(I, F_EmpDetailYn) = XNull(Rs!EmpDetailYn)
            FGrid.TextMatrix(I, F_ChqNo) = XNull(Rs!Chq_No)
            FGrid.TextMatrix(I, F_ChqDate) = XNull(Rs!Chq_Date)
            
            I = I + 1
            Rs.MoveNext
        Loop
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem ""
        FGrid.FixedRows = 1
    End If
    
            
    
    Set Rs = G_FaCn.Execute("Select E.*, Emp.Emp_Name " & _
                            "From Exp_Emp1 E " & _
                            "Left Join Emp_Mast Emp On E.Emp_Code = Emp.Emp_Code " & _
                            "Where E.DocId = '" & RstMain!SearchCode & "' " & _
                            "Order By E.Srl")
    I = 1
    FGrid1.Rows = 1
    If Rs.RecordCount > 0 Then
        Do Until Rs.EOF
            FGrid1.AddItem ""
            
            FGrid1.TextMatrix(I, 0) = I
            FGrid1.TextMatrix(I, Col_Emp_Code) = XNull(Rs!Emp_Code)
            FGrid1.TextMatrix(I, Col_Emp_Desc) = XNull(Rs!Emp_Name)
            FGrid1.TextMatrix(I, Col_AcCode) = XNull(Rs!SubCode)
            FGrid1.TextMatrix(I, Col_Amount) = Format(XNull(Rs!Amount), "0.00")
            
            I = I + 1
            Rs.MoveNext
        Loop
        
        FGrid1.FixedRows = 1
    Else
        FGrid1.AddItem ""
        FGrid1.FixedRows = 1
    End If
    
    ShowEmpAcWise
    AmtCal
End If
Exit Sub
ErrLoop:        MsgBox err.Description
End Sub
Private Sub TopCtrl1_eRef()
    RsAcName.Requery
    RsEmp.Requery
    RsCashBankAc.Requery
    RsVType.Requery
End Sub
Private Sub TopCtrl1_eExit()
    RstMain.Cancel
    Unload Me
End Sub


Private Sub Txt_GotFocus(Index As Integer)

txtgrid1(0).Visible = False
Ctrl_GetFocus txt(Index)
Grid_Hide

Select Case Index
    Case T_VType
        DGVType.Move txt(Index).left, txt(Index).top + txt(Index).height + 20
        If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsVType!Name Then
            RsVType.MoveFirst
            RsVType.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
        
    Case T_VPrefix
        RsVPrefix.Filter = adFilterNone
        RsVPrefix.Filter = "V_Type = '" & txt(T_VType).Tag & "'"
        DgVPrefix.Move txt(Index).left, txt(Index).top + txt(Index).height + 20
        If RsVPrefix.RecordCount = 0 Or (RsVPrefix.EOF = True Or RsVPrefix.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsVPrefix!Name Then
            RsVPrefix.MoveFirst
            RsVPrefix.FIND "Name ='" & txt(Index).TEXT & "'"
        End If

    Case T_ExpAc
        DgAcName.Move txt(Index).left, txt(Index).top + txt(Index).height + 20
        If RsAcName.RecordCount = 0 Or (RsAcName.EOF = True Or RsAcName.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsAcName!Name Then
            RsAcName.MoveFirst
            RsAcName.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
        
    Case T_CashBankAc
        DgCashBankAc.Move txt(Index).left, txt(Index).top + txt(Index).height + 20
        If RsCashBankAc.RecordCount = 0 Or (RsCashBankAc.EOF = True Or RsCashBankAc.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsCashBankAc!Name Then
            RsCashBankAc.MoveFirst
            RsCashBankAc.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
        
End Select
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If

Select Case Index
    Case T_VType
        DGridTxtKeyDown DGVType, txt, Index, RsVType, KeyCode, False, 1
    Case T_VPrefix
        DGridTxtKeyDown DgVPrefix, txt, Index, RsVPrefix, KeyCode, False, 1
    Case T_ExpAc
        DGridTxtKeyDown DgAcName, txt, Index, RsAcName, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
    Case T_CashBankAc
        DGridTxtKeyDown DgCashBankAc, txt, Index, RsCashBankAc, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
        
End Select

If DgAcName.Visible = False And DgVPrefix.Visible = False And DGVType.Visible = False And DgCashBankAc.Visible = False And DGEmp.Visible = False Then

        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> T_Narration Then
            Ctrl_DownKeyDown KeyCode, Shift
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = T_Narration Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> T_Narration Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> T_Narration Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If

End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(keyascii)
Select Case Index
    Case T_VPrefix
        If DgVPrefix.Visible = True Then DGridTxtKeyPress txt, Index, RsVPrefix, keyascii, "Name"
    Case T_VType
        If DGVType.Visible = True Then DGridTxtKeyPress txt, Index, RsVType, keyascii, "Name"
    Case T_ExpAc
        If DgAcName.Visible = True Then DGridTxtKeyPress txt, Index, RsAcName, keyascii, "Name"
    Case T_CashBankAc
        If DgCashBankAc.Visible = True Then DGridTxtKeyPress txt, Index, RsCashBankAc, keyascii, "Name"
End Select

'KeyAscii = RetDGKeyAscii()
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate txt(Index)
    Select Case Index
        Case T_VPrefix
            FGrid.SetFocus
    End Select
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim I As Integer
Dim mDays As Integer
Select Case Index
    Case T_Date
        If txt(Index) = "" Then
            txt(Index) = PubLoginDate
        Else
            txt(Index) = RetDate(txt(Index))
        End If
        Set RsVPrefix = GCn.Execute("Select Prefix As Code, Prefix As Name, VP.V_Type FROM Voucher_Prefix VP Left Join Voucher_Type VT On VP.V_Type = VT.V_Type Where VT.NCat In ('JV') And Date_From<= " & ConvertDate(txt(T_Date)) & " And Date_To>= " & ConvertDate(txt(T_Date)) & " Order by Prefix")
        Set DgVPrefix.DataSource = RsVPrefix
        
    Case T_VType
        If RsVType.RecordCount = 0 Or (RsVType.EOF = True Or RsVType.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
            txt(T_VPrefix) = ""
        Else
            txt(Index).TEXT = RsVType!Name
            txt(Index).Tag = RsVType!Code
            
            RsVPrefix.Filter = adFilterNone
            RsVPrefix.Filter = "V_Type = '" & txt(Index).Tag & "'"
            If RsVPrefix.RecordCount > 0 Then
                txt(T_VPrefix) = RsVPrefix!Name
            End If
        End If
        mDocId = GetDocID(GCnFaS, txt(T_VType).Tag, txt(T_Date), VoucherEditFlag, txt(T_VNo), txt(T_VPrefix))
    Case T_VPrefix
        If RsVPrefix.RecordCount = 0 Or (RsVPrefix.EOF = True Or RsVPrefix.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsVPrefix!Name
            txt(Index).Tag = RsVPrefix!Code
        End If
        mDocId = GetDocID(GCnFaS, txt(T_VType).Tag, txt(T_Date), VoucherEditFlag, txt(T_VNo), txt(T_VPrefix))
        
    Case T_ExpAc
        If RsAcName.RecordCount = 0 Or (RsAcName.EOF = True Or RsAcName.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsAcName!Name
            txt(Index).Tag = RsAcName!Code
        End If
        
    Case T_CashBankAc
        If RsCashBankAc.RecordCount = 0 Or (RsCashBankAc.EOF = True Or RsCashBankAc.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).TEXT = RsCashBankAc!Name
            txt(Index).Tag = RsCashBankAc!Code
        End If
        
    Case T_VNo
        If VoucherEditFlag = True Then      ' Manual
            mDocId = GetDocID(GCnFaS, txt(T_VType).Tag, txt(T_Date), VoucherEditFlag, txt(T_VNo), txt(T_VPrefix))
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            If GCn.Execute("Select Count(*) From Exp_Emp Where Left(DocId,1)='" & PubDivCode & "' And  V_Type='" & txt(T_VType).Tag & "' And V_Prefix='" & txt(T_VPrefix) & "' And  V_No = " & Val(txt(T_VNo)) & " ").Fields(0).Value > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Cancel = True
                txt(T_VNo).SetFocus
            End If
        End If
        
End Select
Set Rst = Nothing

End Sub



Private Sub BlankText()
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).TEXT = ""
    txt(I).Tag = ""
Next I


FGrid.Rows = 1
FGrid.AddItem ""
FGrid.FixedRows = 1


FGrid1.Rows = 1
FGrid1.AddItem ""
FGrid1.FixedRows = 1
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
For I = 0 To txt.Count - 1
    txt(I).Enabled = Enb
Next
txt(T_AmtCr).Enabled = False
txt(T_AmtDr).Enabled = False

End Sub

'Private Sub Ini_Grid()
'    FGrid.RowHeightMin = 250
'    FGrid.ColWidth(25) = 0
'End Sub

Sub Grid_Hide()
    DgAcName.Visible = False
    DGEmp.Visible = False
    DgCashBankAc.Visible = False
    DgVPrefix.Visible = False
    DGVType.Visible = False
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
Ctrl_GetFocus txtgrid1(Index)
    Grid_Hide
    txtgrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    
    Select Case FGrid1.Col
        Case Col_Emp_Desc
            DGEmp.Move txtgrid1(0).left, txtgrid1(0).top + txtgrid1(0).height + 20
            If RsEmp.RecordCount = 0 Or FGrid1.TextMatrix(FGrid1.Row, Col_Emp_Code) = "" Then Exit Sub
            RsEmp.MoveFirst
            RsEmp.FIND "Code ='" & FGrid1.TextMatrix(FGrid1.Row, Col_Emp_Code) & "'"
    End Select
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtgrid1(0).TEXT = txtgrid1(0).Tag
        TxtGrid1_KeyUp Index, KeyCode, Shift
        FGrid1.SetFocus
        txtgrid1(0).Visible = False
        Exit Sub
    End If
    Select Case FGrid1.Col
        Case Col_Emp_Desc
            If DGEmp.Visible = False Then DGridColSwap DGEmp, 0
            DGridTxtKeyDown DGEmp, txtgrid1, Index, RsEmp, KeyCode, False, 1
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DGEmp.Visible = False) Then
                If TxtGrid1Leave = True Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, True, Col_Amount, , Col_Amount
                End If
            End If
        Case Col_Amount
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DGEmp.Visible = False) Then
                If TxtGrid1Leave = True Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, True, Col_Amount, 1
                End If
            End If
                        
                        
    End Select
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, keyascii As Integer)
    Call CheckQuote(keyascii)
    Select Case FGrid1.Col
        Case Col_Emp_Desc
            DGridTxtKeyPress txtgrid1, Index, RsEmp, keyascii, "Name"
        Case Col_Amount
            NumPress txtgrid1(0), keyascii, 8, 2
    End Select
End Sub

Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        Select Case FGrid1.Col
            Case Col_Emp_Desc
                If KeyCode <> 13 And DgAcName.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0
                DGridTxtKeyUp_Mast txtgrid1, Index, RsEmp, KeyCode, "Name"
            Case Col_Amount
                If txtgrid1(0).Visible = True Then
                    FGrid1 = Format(txtgrid1(0), "0.00")
                    AmtCal
                End If
        End Select
End Sub

Private Sub TxtGrid1_LostFocus(Index As Integer)
    'If ExitCtrl = False Then Exit Sub
    txtgrid1(Index).Visible = False
End Sub

Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGrid1Leave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid1Leave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim Repeat$

Select Case FGrid1.Col
    Case Col_Emp_Desc
        With FGrid1
            If RsEmp.RecordCount = 0 Or txtgrid1(0).TEXT = "" Or RsEmp.EOF = True Or RsEmp.BOF = True Then
                .TextMatrix(.Row, Col_Emp_Code) = ""
                .TextMatrix(.Row, Col_Emp_Desc) = ""
                .TextMatrix(.Row, Col_AcCode) = ""
                
            Else
                .TextMatrix(.Row, Col_Emp_Code) = RsEmp!Code
                .TextMatrix(.Row, Col_Emp_Desc) = RsEmp!Name
                .TextMatrix(.Row, Col_AcCode) = FGrid.TextMatrix(FGrid.Row, F_AcCode)
            End If
        End With
    Case Col_Amount
        FGrid1 = Format(txtgrid1(0), "0.00")
End Select
TxtGrid1Leave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid1.SetFocus
    txtgrid1(0).Visible = False
End If
End Function


Sub Ini_Grid()
    mGridBackColor = FGrid.BackColor
    mGridBackColorSel = FGrid.BackColorSel


    With FGrid1
        .Cols = 5
        .TextMatrix(0, 0) = "Srl."
        .BackColorSel = mGridBackColor
                                
        .TextMatrix(0, Col_Emp_Desc) = "Employee Name"
        .ColAlignment(Col_Emp_Desc) = flexAlignLeftCenter
        .ColWidth(Col_Emp_Desc) = 3200
                
        .ColWidth(Col_Emp_Code) = 0
                
        .TextMatrix(0, Col_Amount) = "Amount"
        .ColAlignment(Col_Amount) = flexAlignRightCenter
        .ColWidth(Col_Amount) = 1300
                
        .ColWidth(Col_AcCode) = 0
    End With
    
    
    With FGrid
        .Cols = 12
        .width = 11600
        .BackColorSel = mGridBackColor
        
        .TextMatrix(0, 0) = ""
        
        
        .TextMatrix(0, F_DrCr) = "Dr/Cr"
        .ColAlignment(F_DrCr) = flexAlignLeftCenter
        .ColWidth(F_DrCr) = 600
        
        .TextMatrix(0, F_AcName) = "Particular"
        .ColAlignment(F_AcName) = flexAlignLeftCenter
        .ColWidth(F_AcName) = 3500
        
        .TextMatrix(0, F_Balance) = "Balance"
        .ColAlignment(F_Balance) = flexAlignLeftCenter
        .ColWidth(F_Balance) = 1500
        
        .TextMatrix(0, F_AmtDr) = "Amt. (Dr)"
        .ColAlignment(F_AmtDr) = flexAlignRightCenter
        .ColWidth(F_AmtDr) = 1200
        
        .TextMatrix(0, F_AmtCr) = "Amt. (Cr)"
        .ColAlignment(F_AmtCr) = flexAlignRightCenter
        .ColWidth(F_AmtCr) = 1200
        
        .TextMatrix(0, F_Narration) = "Narration"
        .ColAlignment(F_Narration) = flexAlignLeftCenter
        .ColWidth(F_Narration) = 2800
        
        .TextMatrix(0, F_ChqNo) = "Chq No"
        .ColAlignment(F_ChqNo) = flexAlignLeftCenter
        .ColWidth(F_ChqNo) = 1000
        
        .TextMatrix(0, F_ChqDate) = "Chq Date"
        .ColAlignment(F_ChqDate) = flexAlignLeftCenter
        .ColWidth(F_ChqDate) = 1000
        
        .ColWidth(F_AcCode) = 0
        .ColWidth(F_EmpDetailYn) = 0
        .ColWidth(F_Nature) = 0
    End With
    
    FGrid.Col = F_AmtDr
    txt(T_AmtDr).left = FGrid.CellLeft
    txt(T_AmtDr).width = FGrid.CellWidth
    FGrid.Col = F_AmtCr
    txt(T_AmtCr).left = FGrid.CellLeft
    txt(T_AmtCr).width = FGrid.CellWidth
End Sub



Sub AmtCal()
    Dim I As Integer
    Dim j As Integer
    Dim mAmt As Double
    Dim mDR As Double, mCR As Double
    
    
    mAmt = 0
    DoEvents
    For j = 1 To FGrid.Rows - 1
        If StrCmp(FGrid.TextMatrix(j, F_EmpDetailYn), "Y") Then
            mAmt = 0
            For I = 1 To FGrid1.Rows - 1
                If StrCmp(FGrid.TextMatrix(j, F_AcCode), FGrid1.TextMatrix(I, Col_AcCode)) Then
                    mAmt = mAmt + Val(FGrid1.TextMatrix(I, Col_Amount))
                End If
            Next I
            FGrid.TextMatrix(j, F_AmtDr) = Format(mAmt, "0.00")
        End If
    Next j
    

    mDR = 0: mCR = 0
    With FGrid
        For I = 1 To .Rows - 1
            If StrCmp(.TextMatrix(I, F_DrCr), "Dr") And .TextMatrix(I, F_AcName) <> "" Then
                mDR = mDR + Val(.TextMatrix(I, F_AmtDr))
            ElseIf StrCmp(.TextMatrix(I, F_DrCr), "Cr") And .TextMatrix(I, F_AcName) <> "" Then
                mCR = mCR + Val(.TextMatrix(I, F_AmtCr))
            End If
        Next I
    End With
    txt(T_AmtDr) = Format(mDR, "0.00")
    txt(T_AmtCr) = Format(mCR, "0.00")
End Sub

Private Sub FGrid_Click()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    txtgrid(0).Visible = False
End Sub


Private Sub FGrid_DblClick()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid.Col
        Case F_DrCr, F_AcName, F_Narration
            Call GridDblClick(Me, FGrid, txtgrid, 0)
        Case F_ChqNo, F_ChqDate
            If StrCmp(FGrid.TextMatrix(FGrid.Row, F_Nature), "Bank") Then
                Call GridDblClick(Me, FGrid, txtgrid, 0)
            End If
        
        Case F_AmtDr
            If StrCmp(FGrid.TextMatrix(FGrid.Row, F_DrCr), "Dr") And FGrid.TextMatrix(FGrid.Row, F_AcName) <> "" Then
                GridDblClick Me, FGrid, txtgrid, 0
            End If
        Case F_AmtCr
            If StrCmp(FGrid.TextMatrix(FGrid.Row, F_DrCr), "Cr") And FGrid.TextMatrix(FGrid.Row, F_AcName) <> "" Then
                GridDblClick Me, FGrid, txtgrid, 0
            End If
        
    End Select
End Sub

Private Sub FGrid_EnterCell()
    'FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = mGridBackColorSel  'FaBackColorSelEnter

    'FGrid.Col = Col_Emp_Desc
    txtgrid(0).Visible = False
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    If MsgBox("Do You Want to Save?", vbYesNo) = vbYes Then TopCtrl1_eSave
'    SendKeysA vbKeyTab, True
'    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    Select Case FGrid.Col
        Case F_AcName
            FGrid = ""
            FGrid.TextMatrix(FGrid.Row, F_AcCode) = ""
            FGrid.TextMatrix(FGrid.Row, F_Balance) = ""
        Case Else
            FGrid = ""
    End Select
End If
If KeyCode = vbKeyReturn Then
    Select Case FGrid.Col
        Case F_DrCr, F_AcName, F_Narration
            Call GridDblClick(Me, FGrid, txtgrid, 0)
        Case F_ChqNo, F_ChqDate
            If StrCmp(FGrid.TextMatrix(FGrid.Row, F_Nature), "Bank") Then
                Call GridDblClick(Me, FGrid, txtgrid, 0)
            End If
        Case F_AmtDr
            If StrCmp(FGrid.TextMatrix(FGrid.Row, F_DrCr), "Dr") And FGrid.TextMatrix(FGrid.Row, F_AcName) <> "" Then
                GridDblClick Me, FGrid, txtgrid, 0
            End If
        Case F_AmtCr
            If StrCmp(FGrid.TextMatrix(FGrid.Row, F_DrCr), "Cr") And FGrid.TextMatrix(FGrid.Row, F_AcName) <> "" Then
                GridDblClick Me, FGrid, txtgrid, 0
            End If
            
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)
    Select Case FGrid.Col
        Case F_AcName, F_Narration
            Call Get_Text(Me, FGrid, txtgrid, 0, False, keyascii)
        Case F_ChqNo, F_ChqDate
            'If StrCmp(FGrid.TextMatrix(FGrid.Row, F_Nature), "Bank") Then
                Call Get_Text(Me, FGrid, txtgrid, 0, False, keyascii)
            'End If
        Case F_AmtDr
            If StrCmp(FGrid.TextMatrix(FGrid.Row, F_DrCr), "Dr") And FGrid.TextMatrix(FGrid.Row, F_AcName) <> "" Then
                Call Get_Text(Me, FGrid, txtgrid, 0, True, keyascii)
            End If
        Case F_AmtCr
            If StrCmp(FGrid.TextMatrix(FGrid.Row, F_DrCr), "Cr") And FGrid.TextMatrix(FGrid.Row, F_AcName) <> "" Then
                Call Get_Text(Me, FGrid, txtgrid, 0, True, keyascii)
            End If

        Case F_DrCr
            If TopCtrl1.TopText2 <> "Browse" Then
                If keyascii = Asc("D") Or keyascii = Asc("d") Then
                    FGrid = "Dr"
                    FGrid.Col = F_AcName
                    RsAcName.Filter = adFilterNone
                    RsAcName.Filter = "Nature = 'Expenses'"
                    Set DgAcName.DataSource = RsAcName
                ElseIf keyascii = Asc("C") Or keyascii = Asc("c") Then
                    FGrid = "Cr"
                    FGrid.Col = F_AcName
                    RsAcName.Filter = adFilterNone
                    Set DgAcName.DataSource = RsAcName
                Else
                    FGrid = ""
                End If
            End If
    End Select
End Sub

Private Sub FGrid_LostFocus()
FGrid.BackColorSel = mGridBackColor   'FaCellBackColLeave1
FGrid_Validate (True)
End Sub

Private Sub FGrid_Scroll()
    txtgrid(0).Visible = False

End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid.ColSel = False Then Exit Sub
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
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
   
FGrid.SetFocus
End If
Exit Sub
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
'    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
Ctrl_GetFocus txtgrid(Index)
    Grid_Hide
    txtgrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    
    Select Case FGrid.Col
        Case F_AcName
            DgAcName.Move txtgrid(0).left, txtgrid(0).top + txtgrid(0).height + 20
            If RsAcName.RecordCount = 0 Or FGrid.TextMatrix(FGrid.Row, F_AcCode) = "" Then Exit Sub
            RsAcName.MoveFirst
            RsAcName.FIND "Code ='" & FGrid.TextMatrix(FGrid.Row, F_AcCode) & "'"
    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim mLastCol As Byte
    If KeyCode = vbKeyEscape Then
        txtgrid(0).TEXT = txtgrid(0).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        FGrid.SetFocus
        txtgrid(0).Visible = False
        Exit Sub
    End If
    mLastCol = F_Narration
    Select Case FGrid.Col
        Case F_DrCr, F_Narration, F_ChqNo, F_ChqDate
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DGEmp.Visible = False) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, txtgrid, Index, KeyCode, True, mLastCol, 1
                End If
            End If
        Case F_ChqNo
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DGEmp.Visible = False) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, txtgrid, Index, KeyCode, True, F_ChqDate, 1
                End If
            End If
        
        Case F_AmtDr, F_AmtCr
            If KeyCode = vbKeyReturn Or (KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, txtgrid, Index, KeyCode, True, mLastCol, , F_Narration
                End If
            End If
        
        Case F_AcName
            If DgAcName.Visible = False Then DGridColSwap DgAcName, 0
            DGridTxtKeyDown DgAcName, txtgrid, Index, RsAcName, KeyCode, False, 1
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DgAcName.Visible = False) Then
                If TxtGridLeave = True Then
                    If StrCmp(FGrid.TextMatrix(FGrid.Row, F_DrCr), "Dr") Then
                        GridTxtDown FGrid, txtgrid, Index, KeyCode, True, mLastCol, , F_AmtDr
                    Else
                        GridTxtDown FGrid, txtgrid, Index, KeyCode, True, mLastCol, , F_AmtCr
                    End If
                End If
            End If
                        
                        
    End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, keyascii As Integer)
    Call CheckQuote(keyascii)
    Select Case FGrid.Col
        Case F_AcName
            DGridTxtKeyPress txtgrid, Index, RsAcName, keyascii, "Name"
        Case F_AmtDr, F_AmtCr
            NumPress txtgrid(0), keyascii, 8, 2
    End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        Select Case FGrid.Col
            Case F_AcName
                If KeyCode <> 13 And DgAcName.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
                DGridTxtKeyUp_Mast txtgrid, Index, RsAcName, KeyCode, "Name"
            Case F_AmtDr, F_AmtCr
                FGrid = Format(txtgrid(0), "0.00")
                AmtCal
        End Select
End Sub

Private Sub TxtGrid_LostFocus(Index As Integer)
    'If ExitCtrl = False Then Exit Sub
    txtgrid(Index).Visible = False
    Select Case FGrid.Col
        Case F_AmtDr
            If StrCmp(FGrid.TextMatrix(FGrid.Row, F_EmpDetailYn), "Y") Then
                FGrid1.SetFocus
                FGrid1.Col = Col_Emp_Desc
                FGrid1.Row = FGrid1.Rows - 1
            End If
    End Select
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim Repeat$

Select Case FGrid.Col
    Case F_AcName
        With FGrid
            If RsAcName.RecordCount = 0 Or txtgrid(0).TEXT = "" Or RsAcName.EOF = True Or RsAcName.BOF = True Then
                .TextMatrix(.Row, F_AcName) = ""
                .TextMatrix(.Row, F_AcCode) = ""
                .TextMatrix(.Row, F_Balance) = ""
                .TextMatrix(.Row, F_EmpDetailYn) = ""
                .TextMatrix(.Row, F_Nature) = ""
            Else
                .TextMatrix(.Row, F_AcCode) = RsAcName!Code
                .TextMatrix(.Row, F_AcName) = RsAcName!Name
                .TextMatrix(.Row, F_Balance) = Abs(VNull(RsAcName!Curr_Bal)) & IIf(VNull(RsAcName!Curr_Bal) < 0, " Dr", " Cr")
                .TextMatrix(.Row, F_EmpDetailYn) = XNull(RsAcName!EmpDetailYn)
                .TextMatrix(.Row, F_Nature) = XNull(RsAcName!Nature)
                Fill_AgAmt
            End If
        End With
        ShowEmpAcWise
    Case F_AmtDr, F_AmtCr
        FGrid = Format(txtgrid(0), "0.00")
        ShowEmpAcWise
    Case F_Narration, F_ChqNo
        FGrid = txtgrid(0)
        ShowEmpAcWise
    Case F_ChqDate
        FGrid = RetDate(txtgrid(0))
        ShowEmpAcWise
End Select

TxtGridLeave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid.SetFocus
    txtgrid(0).Visible = False
End If
End Function

Sub Fill_AgAmt()
Dim I As Integer
Dim mDR As Double, mCR As Double
    With FGrid
        For I = 1 To FGrid.Rows - 1
            If StrCmp(.TextMatrix(I, F_DrCr), "Dr") Then
                mDR = mDR + Val(.TextMatrix(I, F_AmtDr))
            ElseIf StrCmp(.TextMatrix(I, F_DrCr), "Cr") Then
                mCR = mCR + Val(.TextMatrix(I, F_AmtCr))
            End If
        Next I
        
        If StrCmp(FGrid.TextMatrix(FGrid.Row, F_DrCr), "Dr") Then
            If Val(.TextMatrix(.Row, F_AmtDr)) = 0 Then
                If mCR > mDR Then
                    .TextMatrix(.Row, F_AmtDr) = Format(mCR - mDR, "0.00")
                End If
            End If
        End If
        
        If StrCmp(FGrid.TextMatrix(FGrid.Row, F_DrCr), "Cr") Then
            If Val(.TextMatrix(.Row, F_AmtCr)) = 0 Then
                If mDR > mCR Then
                    .TextMatrix(.Row, F_AmtCr) = Format(mDR - mCR, "0.00")
                End If
            End If
        End If
        
    End With
End Sub
