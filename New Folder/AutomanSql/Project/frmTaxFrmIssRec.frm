VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmTaxFrmIssRec 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Tax Form Issue / Receipt"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11610
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
   ScaleHeight     =   7335
   ScaleWidth      =   11610
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton CmdTrn 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Display Bill Details"
      Height          =   645
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   555
      Width           =   1335
   End
   Begin VB.Frame FrameTrn 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFCCD8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4920
      Left            =   1695
      TabIndex        =   33
      Top             =   1935
      Visible         =   0   'False
      Width           =   8055
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGridTrn 
         Height          =   4290
         Left            =   0
         TabIndex        =   34
         Top             =   735
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   7567
         _Version        =   393216
         BackColor       =   15525079
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   14940925
         ForeColorFixed  =   16576
         BackColorSel    =   16711680
         BackColorBkg    =   14737632
         BackColorUnpopulated=   14865856
         GridColor       =   14940925
         GridColorFixed  =   12632319
         FocusRect       =   0
         GridLinesFixed  =   1
         BorderStyle     =   0
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   "*Bill Details*"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Index           =   2
         Left            =   6330
         TabIndex        =   39
         Top             =   0
         Width           =   1680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party   :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   13
         Left            =   1455
         TabIndex        =   38
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SlNo.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   12
         Left            =   105
         TabIndex        =   37
         Top             =   45
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form No.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   11
         Left            =   1455
         TabIndex        =   36
         Top             =   45
         Width           =   1080
      End
   End
   Begin VB.Frame FrmDup 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Duplicate Form No's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5640
      Left            =   11475
      TabIndex        =   29
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid2 
         Height          =   4890
         Left            =   90
         TabIndex        =   31
         Top             =   315
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   8625
         _Version        =   393216
         ForeColor       =   4194304
         FixedCols       =   0
         BackColorFixed  =   16761024
         BackColorBkg    =   12632256
         GridColor       =   12640511
         TextStyleFixed  =   1
         FocusRect       =   0
         GridLinesFixed  =   0
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton CmdClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Close"
         Default         =   -1  'True
         Height          =   345
         Left            =   255
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   5250
         Width           =   2280
      End
   End
   Begin VB.TextBox Txt 
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
      Left            =   2355
      MaxLength       =   4
      TabIndex        =   6
      Text            =   "XXXX"
      Top             =   1470
      Width           =   525
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEE0FD&
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
      Left            =   4485
      MaxLength       =   6
      TabIndex        =   7
      Text            =   "999999"
      Top             =   1470
      Width           =   945
   End
   Begin VB.TextBox Txt 
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
      Left            =   5910
      MaxLength       =   6
      TabIndex        =   8
      Text            =   "999999"
      Top             =   1470
      Width           =   945
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBDAE9&
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
      Left            =   8460
      MaxLength       =   15
      TabIndex        =   9
      Top             =   1470
      Width           =   1230
   End
   Begin MSDataGridLib.DataGrid DGTaxAc 
      Height          =   2865
      Left            =   435
      Negotiate       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5205
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
            DividerStyle    =   3
            ColumnWidth     =   30.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4710.047
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8415
      TabIndex        =   5
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5250
      TabIndex        =   4
      Text            =   "Purchase/Sale/Road Permit"
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2490
      TabIndex        =   3
      Text            =   "Local"
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2490
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "0123"
      Top             =   600
      Width           =   675
   End
   Begin VB.TextBox Txt 
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
      Left            =   5250
      MaxLength       =   25
      TabIndex        =   2
      Text            =   "0123456789012345678901234"
      Top             =   600
      Width           =   4395
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   661
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   8100
      TabIndex        =   12
      Top             =   4995
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   90
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   15
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
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   2520
      TabIndex        =   11
      Top             =   4575
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   3075
      Left            =   90
      TabIndex        =   10
      Top             =   1935
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   5424
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   9
      BackColorFixed  =   12640511
      ForeColorFixed  =   128
      BackColorSel    =   16711680
      ForeColorSel    =   12648447
      BackColorBkg    =   13623520
      GridColor       =   8438015
      GridColorFixed  =   192
      GridColorUnpopulated=   16711935
      FocusRect       =   0
      GridLinesFixed  =   1
      MergeCells      =   1
      AllowUserResizing=   1
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
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label AppFor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ApplicableFor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   10200
      TabIndex        =   32
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Trn. Type :"
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
      Left            =   7005
      TabIndex        =   28
      Top             =   885
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   " Receipt From Deptt. Only "
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
      Index           =   9
      Left            =   960
      TabIndex        =   27
      Top             =   1200
      Width           =   2130
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFFF&
      Height          =   495
      Left            =   855
      Top             =   1335
      Width           =   9285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Series :"
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
      Index           =   8
      Left            =   1200
      TabIndex        =   26
      Top             =   1485
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No. From :"
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
      Index           =   7
      Left            =   3045
      TabIndex        =   25
      Top             =   1485
      Width           =   1380
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   6
      Left            =   5535
      TabIndex        =   24
      Top             =   1485
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Date :"
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
      Index           =   3
      Left            =   7245
      TabIndex        =   23
      Top             =   1485
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Index           =   4
      Left            =   4725
      TabIndex        =   21
      Top             =   885
      Width           =   390
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   4
      Left            =   5145
      TabIndex        =   20
      Top             =   885
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   8
      Left            =   2385
      TabIndex        =   19
      Top             =   885
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local / Central"
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
      Index           =   0
      Left            =   945
      TabIndex        =   18
      Top             =   885
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   5
      Left            =   2385
      TabIndex        =   17
      Top             =   615
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Code"
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
      Index           =   5
      Left            =   945
      TabIndex        =   16
      Top             =   615
      Width           =   930
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   9
      Left            =   5145
      TabIndex        =   15
      Top             =   615
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Index           =   1
      Left            =   4170
      TabIndex        =   14
      Top             =   615
      Width           =   945
   End
End
Attribute VB_Name = "frmTaxFrmIssRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const TrnHideColor As String = &HC0FFFF
Private Const TrnHideCaption As String = "Hide Bill Details"
Private Const TrnShowColor As String = &HEBDAE9
Private Const TrnShowCaption As String = "Display Bill Details"

Dim FGridTrnModified As Boolean
Dim TabName$
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim RsTaxAc As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim mSearchCode As String
Dim ListArray As Variant
Dim mListItem As ListItem
Dim GridRow1() As Integer
Dim mGridStartRow As Integer
Dim mGridEndRow As Integer

'grid color scheme
Private Const CellBackColLeave As String = &HEDF7FE
Private Const CellForeColLeave As String = &HFF00FF
Private Const CellBackColEnter As String = &HEBDAE9    '&HCAF1FD
'Private Const GridBackColorBkg As String = Me.BackColor
Private Const CellBackColLeave1 As String = &HECE4D7
Private Const CellBackColEnter1 As String = &HEBDAE9 '&HFFFFC0

Private Const FormCode As Byte = 0              ' Form Code
Private Const FormDesc As Byte = 1              ' Form Description
Private Const L_C As Byte = 2                   ' Local / Central
Private Const TrnType As Byte = 3               ' Form Type Purchase / Sale
Private Const FormTrnType As Byte = 4           ' Form Transaction Type

Private Const FormSeries As Byte = 5            ' Form Series if any
Private Const SrlNoFrom As Byte = 6             ' Form Serial No. From
Private Const SrlNoTo As Byte = 7               ' Form Serial No. To
Private Const RectDate As Byte = 8              ' Receipt Date

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_FormNo As Byte = 1            ' Form No.
Private Const Col_RecFrom As Byte = 2           ' Rec From Deptt / Party
Private Const Col_RecDate As Byte = 3           ' RecDate
Private Const Col_SubCode As Byte = 4           ' Issue to Party
Private Const Col_SubName As Byte = 5           ' Issue to Party
Private Const Col_IssDate As Byte = 6           ' Issue Date
Private Const Col_CurrentStatus As Byte = 7     ' Qty Return
Private Const Col_Remarks As Byte = 8           ' Remarks
Private Const Col_AddEdit As Byte = 9           ' Add / Edit /Delete Status

'* Grid Column Declaration for GridTrn
Private Const Trn_FormNo As Byte = 0            ' Form No.
Private Const Trn_IssRecDate As Byte = 1        ' RecDate
'Private Const Trn_SrNo As Byte = 0              ' Serial No
Private Const Trn_SubCode As Byte = 2           ' Party Code
Private Const Trn_DocID As Byte = 3             ' DocID
Private Const Trn_DocNo As Byte = 4             ' DocNo
Private Const Trn_VDate As Byte = 5             ' V_Date
Private Const Trn_BillAmt As Byte = 6           ' Bill Amount
Private Const Trn_OldFrmNo As Byte = 7            ' For Internal Use

Private Sub Disp_Text(Enb As Boolean)
    If TopCtrl1.TopText2 = "Edit" Then
        Enb = False
    End If
    Txt(FormCode).Enabled = False
    Txt(FormDesc).Enabled = False
    Txt(L_C).Enabled = False
    Txt(TrnType).Enabled = False
    Txt(FormTrnType).Enabled = False
    Txt(FormSeries).Enabled = Enb
    Txt(SrlNoFrom).Enabled = Enb
    Txt(SrlNoTo).Enabled = Enb
    Txt(RectDate).Enabled = Enb

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

Private Sub BlankText()
Dim I As Byte
'* Used for clear all text boxes used in the form
    For I = 5 To 8
        Txt(I).TEXT = ""
    Next I
'Used for Grid Coloumn Width Initialization on various Option Like (Workshop/Store/Return)
'    FGrid.Rows = 1
'    FGrid.AddItem FGrid.Rows
'    FGrid.FixedRows = 1
End Sub

'* Used for intialize grid columns
Private Sub Grid_Ini()
FrmDup.left = 4575
FrmDup.top = 1185
    
    With FGrid2
        .ColWidth(0) = 800
        .TextMatrix(0, 0) = "Srl.No."
        .ColWidth(1) = 1400
        .TextMatrix(0, 1) = "Form No."
    End With
    
    With FGrid
        .left = Me.left '+ 45
        .width = Me.width - 150
        .top = 1935
        .height = FGrid.RowHeight(0) * 15
        .RowHeightMin = 0 'PubGridRowHeight
        .BackColor = CellBackColLeave
        .BackColorBkg = Me.BackColor
        .Cols = 11
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        
        .TextMatrix(0, Col_SrNo) = "S.No."
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 480

        .TextMatrix(0, Col_FormNo) = "Form No."
        .ColAlignment(Col_FormNo) = flexAlignLeftCenter
        .ColWidth(Col_FormNo) = 1185

        .TextMatrix(0, Col_RecFrom) = "RecFrom"
        .ColAlignment(Col_RecFrom) = flexAlignLeftCenter
        .ColWidth(Col_RecFrom) = 840
        
        .TextMatrix(0, Col_RecDate) = "Rec.Date"
        .ColAlignment(Col_RecDate) = flexAlignLeftCenter
        .ColWidth(Col_RecDate) = 1200
        
        .TextMatrix(0, Col_SubCode) = "Issue To Party Code"
        .ColAlignment(Col_SubCode) = flexAlignLeftCenter
        .ColWidth(Col_SubCode) = 0

        .TextMatrix(0, Col_SubName) = "Issue To Party"
        .ColAlignment(Col_SubName) = flexAlignLeftCenter
        .ColWidth(Col_SubName) = 3500

        .TextMatrix(0, Col_IssDate) = "Issue Date"
        .ColAlignment(Col_IssDate) = flexAlignLeftCenter
        .ColWidth(Col_IssDate) = 1200
        
        .TextMatrix(0, Col_CurrentStatus) = "Status"
        .ColAlignment(Col_CurrentStatus) = flexAlignLeftCenter
        .ColWidth(Col_CurrentStatus) = 810

        .TextMatrix(0, Col_Remarks) = "Remarks"
        .ColAlignment(Col_Remarks) = flexAlignLeftCenter
        .ColWidth(Col_Remarks) = 2265
    End With
    DGTaxAc.width = 5130:   DGTaxAc.left = 6700
    DGTaxAc.top = mTopScale '390
    DGTaxAc.height = 4935
End Sub

Private Sub Grid_Hide()
    If DGTaxAc.Visible = True Then DGTaxAc.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Sub MoveRec()
Dim Rst As ADODB.Recordset, I As Integer
Dim mCurrStatus$
On Error GoTo ELoop
    If Master.RecordCount > 0 Then
        Txt(FormCode) = Master!Form_Code
        Txt(FormDesc) = Master!form_desc
        Txt(L_C) = Master!L_C
        Txt(TrnType) = Master!Trn_Type
        Txt(FormTrnType) = FxFormTrnType(Master!FormTrnType)
        AppFor.CAPTION = Master!AppFor
        Set Rst = GCn.Execute("Select TStk.FormNo,TStk.RecFrom,TStk.RecDate,TStk.IssDate, " _
            & "TStk.CurrentStatus,TStk.Remarks,TStk.SubCode,SubGroup.Name " _
            & "From TaxFormStk TStk Left Join SubGroup On TStk.SubCode=SubGroup.SubCode " _
            & "Where TStk.Form_Code='" & Master!Form_Code & "' Order By TStk.FormNo")
            
        FGrid.Redraw = False
        FGrid.Rows = 1
        If Rst.RecordCount > 0 Then
            I = 1
            Do Until Rst.EOF
                mCurrStatus = FormCurrStatus(Rst!CurrentStatus)
                FGrid.AddItem ""
                With FGrid
                    .TextMatrix(I, Col_SrNo) = I
                    .TextMatrix(I, Col_FormNo) = Rst!FormNo
                    .TextMatrix(I, Col_RecFrom) = IIf(Rst!RecFrom = "D", "Deptt.", "Party")
                    .TextMatrix(I, Col_RecDate) = Rst!RecDate
                    .TextMatrix(I, Col_IssDate) = IIf(IsNull(Rst!IssDate), "", Rst!IssDate)
                    .TextMatrix(I, Col_SubCode) = IIf(IsNull(Rst!SubCode), "", Rst!SubCode)
                    .TextMatrix(I, Col_SubName) = IIf(IsNull(Rst!Name), "", Rst!Name)
                    .TextMatrix(I, Col_CurrentStatus) = mCurrStatus
                    .TextMatrix(I, Col_Remarks) = IIf(IsNull(Rst!Remarks), "", Rst!Remarks)
                    .TextMatrix(I, Col_AddEdit) = "N"
                End With
                Rst.MoveNext
                I = I + 1
            Loop
            FGrid.FixedRows = 1
        Else
            FGrid.AddItem FGrid.Rows
            FGrid.FixedRows = 1
            TopCtrl1.tDel = False
        End If
        FGrid.Redraw = True
    End If
    BlankText
    Set Rst = Nothing
    Grid_Hide
    TopCtrl1.tPrn = False
    
Exit Sub
ELoop:
    CheckError
End Sub

' Used For Checking Duplicate Items in the Grid
Private Function ChkDuplicate(GridCol As Byte) As Boolean
Dim I As Integer, SearchString As String
'    SearchString = UCase(FGrid.TextMatrix(FGrid.Row, GridCol))
    SearchString = UCase(TxtGrid(0))
    For I = 1 To FGrid.Rows - 1
        If I = FGrid.Row Then GoTo nxt1
        If SearchString = UCase(FGrid.TextMatrix(I, GridCol)) And FGrid.TextMatrix(I, GridCol) <> "" Then
            MsgBox "Duplicate Form No. Not Allowed", vbInformation, "Validation"
            TxtGrid(0).SetFocus:  ChkDuplicate = False:   Exit Function
        End If
nxt1:
    Next
    ChkDuplicate = True
End Function

Private Sub CmdClose_Click()
FrmDup.Visible = False
End Sub

Private Sub CmdTrn_Click()
If CmdTrn.CAPTION = TrnShowCaption Then
    'Display
    If FGrid.TextMatrix(FGrid.Row, Col_SubName) <> "" Then
        CmdTrn.BackColor = TrnHideColor
        CmdTrn.CAPTION = TrnHideCaption
        FGrid.Col = Col_FormNo
        FGrid.CellFontBold = True
        DispTrn
    End If
Else
    'Hide DispTrn
    If FGridTrnModified Then
        If MsgBox("Save corrections ? ", vbYesNo + vbCritical + vbDefaultButton2, "Save Form") = vbYes Then
            FGridMainUpd
        End If
    End If
    FrameTrn.Visible = False
    CmdTrn.BackColor = TrnShowColor
    CmdTrn.CAPTION = TrnShowCaption
    FGrid.Col = Col_FormNo
    FGrid.CellFontBold = False
End If
End Sub

Private Sub DGTaxAc_Click()
    If RsTaxAc.RecordCount > 0 Then
        TxtGrid(0).Tag = RsTaxAc!Code
        TxtGrid(0) = RsTaxAc!Name
    End If
    If TxtGrid(0).Visible = True Then TxtGrid(0).SetFocus
    DGTaxAc.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub


Private Sub FGridTrn_EnterCell()
'FGridTrn.CellBackColor = CellBackColEnter1
End Sub

Private Sub FGridTrn_GotFocus()
'FGridTrn.CellBackColor = CellBackColEnter1

End Sub

Private Sub FGridTrn_KeyDown(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If KeyCode = 13 Then SendKeysA vbKeyTab, True
If FGridTrn.Rows < 1 Or TopCtrl1.TopText2 = "Browse" Then Exit Sub
If KeyCode = vbKeySpace And FGridTrn.Col = 0 Then
    If FGridTrn.TextMatrix(FGridTrn.Row, Trn_DocNo) <> "" Then
        If FGridTrn.TextMatrix(FGridTrn.Row, Trn_FormNo) = "" Then '"ü", " ", "ü")
            FGridTrn.CellFontSize = 11
            FGridTrn.CellFontBold = True
            FGridTrn.TextMatrix(FGridTrn.Row, Trn_FormNo) = FGrid.TextMatrix(FGrid.Row, Col_FormNo)
            If Txt(FormTrnType) = "Receipt" Then
                FGridTrn.TextMatrix(FGridTrn.Row, Trn_IssRecDate) = FGrid.TextMatrix(FGrid.Row, Col_RecDate)
            Else
                FGridTrn.TextMatrix(FGridTrn.Row, Trn_IssRecDate) = FGrid.TextMatrix(FGrid.Row, Col_IssDate)
            End If
            FGridTrn.Col = Trn_IssRecDate
            FGridTrn.CellFontSize = 11
            FGridTrn.CellFontBold = True
        Else
            FGridTrn.TextMatrix(FGridTrn.Row, Trn_FormNo) = ""
            FGridTrn.TextMatrix(FGridTrn.Row, Trn_IssRecDate) = ""
        End If
        FGridTrnModified = True
        I = UBound(GridRow1) + 1
        ReDim Preserve GridRow1(I)
        GridRow1(I) = FGridTrn.Row
    End If
End If

End Sub

Private Sub FGridTrn_KeyPress(keyascii As Integer)
If FGridTrn.Col = 0 Or FGridTrn.Row = 0 Then Exit Sub

End Sub

Private Sub FGridTrn_LeaveCell()
'FGridTrn.CellBackColor = CellBackColLeave1
End Sub

Private Sub FGridTrn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FGridTrn.Col <> 0 Then Exit Sub
mGridStartRow = FGridTrn.Row

End Sub

Private Sub FGridTrn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer
Dim j As Integer
If FGridTrn.Col <> 0 Or mGridStartRow = 0 Or TopCtrl1.TopText2 = "Browse" Then Exit Sub
mGridEndRow = FGridTrn.RowSel
For j = mGridStartRow To mGridEndRow

    FGridTrn.Row = j
    FGridTrn.Col = 0
    If FGridTrn.TextMatrix(j, Trn_DocNo) <> "" Then
        If FGridTrn.TextMatrix(j, Trn_FormNo) = "" Then '"ü", " ", "ü")
            FGridTrn.CellFontSize = 11
            FGridTrn.CellFontBold = True
            FGridTrn.TextMatrix(j, Trn_FormNo) = FGrid.TextMatrix(FGrid.Row, Col_FormNo)
            If Txt(FormTrnType) = "Receipt" Then
                FGridTrn.TextMatrix(j, Trn_IssRecDate) = FGrid.TextMatrix(FGrid.Row, Col_RecDate)
            Else
                FGridTrn.TextMatrix(j, Trn_IssRecDate) = FGrid.TextMatrix(FGrid.Row, Col_IssDate)
            End If
            FGridTrn.Col = Trn_IssRecDate
            FGridTrn.CellFontSize = 11
            FGridTrn.CellFontBold = True
        Else
            FGridTrn.TextMatrix(j, Trn_FormNo) = ""
            FGridTrn.TextMatrix(j, Trn_IssRecDate) = ""
        End If
        FGridTrnModified = True
        I = UBound(GridRow1) + 1
        ReDim Preserve GridRow1(I)
        GridRow1(I) = FGridTrn.Row
    End If
Next
FGridTrn.Row = mGridStartRow
FGridTrn.CellBackColor = CellBackColLeave1
FGridTrn.Row = mGridEndRow
mGridStartRow = 0

End Sub

Private Sub FGridTrn_Validate(Cancel As Boolean)
'FGridTrn.CellBackColor = CellBackColLeave1

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
    CheckError
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
    TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg '&HDFF4F2
        Txt(I).ForeColor = CtrlFColOrg
'        Txt(I).BorderStyle = 1
    Next
'    Hook TxtGrid(0).hWnd
    CmdTrn.BackColor = TrnShowColor
    CmdTrn.CAPTION = TrnShowCaption

    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME from SubGroup " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        "order by SubGroup.name"
    Set RsTaxAc = New ADODB.Recordset
    RsTaxAc.CursorLocation = adUseClient
    RsTaxAc.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGTaxAc.DataSource = RsTaxAc
    
    Set Master = New ADODB.Recordset
    With Master
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
    End With
    
    Set Master = GCn.Execute("Select " & cIIF("TF.Spare_YN=1", "'Spare'", "'Vehicle'") & " As AppFor, T.Form_Code,T.Form_Desc,T.Form_Code As SearchCode,T.L_C,t.Trn_Type,t.FormTrnType " & _
        " from TaxForms T left join TaxForms TF on T.Form_Code=TF.Form_Code " & _
        " where T.FormTrnType >0 " & _
        " Order by T.Form_Code,T.Trn_Type,TF.Spare_YN,TF.Vehicle_YN")
        
    MoveRec
    Disp_Text SETS("INI", Me, Master)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Resize()
    Grid_Ini
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (TopCtrl1.tSave = True And TopCtrl1.tCancel = False) Then
        MsgBox "Please Save the Updated Data !", vbOKOnly, "Save Updation's"
        Me.ActiveControl.SetFocus
        Cancel = 1
        Exit Sub
    End If
    Set RsTaxAc = Nothing
    Set Master = Nothing
End Sub

Private Sub ListView_Click()
On Error GoTo ELoop
    If TxtGrid(0).Visible = True Then
        TxtGrid(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
        FrmList.Visible = False
        TxtGrid(Val(ListView.Tag)).SetFocus
    Else
        Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
        FrmList.Visible = False
        Txt(Val(ListView.Tag)).SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    If MsgBox("Form Receipt from Deptt.? ", vbYesNo + vbCritical + vbDefaultButton2, "Deptt. Receipts") = vbYes Then
        Txt(FormSeries).SetFocus
    Else
        Disp_Text (False)
        FGrid.Row = FGrid.Rows - 1
        FGrid.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
'Check for existance of transactions
    Disp_Text SETS("EDIT", Me, Master)
    FGrid.Row = FGrid.Rows - 1
    FGrid.SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim vBook As Variant
'Check for existance of transactions
    If Master.RecordCount > 0 Then
        If FGrid.Rows > 1 Then
            If MsgBox("All Form Transactions will be Removed." & vbCrLf & "Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
    '            vBook = Master.AbsolutePosition
                GCn.BeginTrans
                GCn.Execute ("Delete From TaxFormStk Where Form_Code='" & Master!Form_Code & "'")
                GCn.CommitTrans
                
                BUTTONS True, Me, Master, 0
                MoveRec
            End If
        Else
            MsgBox "No Records To Delete!", vbInformation, "Information"
        End If
    Else
        MsgBox "No Records To Delete!", vbInformation, "Information"
    End If
Exit Sub
ELoop:
    GCn.RollbackTrans
    CheckError
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
    GSQL = "Select T.Form_Code As SearchCode,Form_Code,Form_Desc," & _
        "L_C as Local_Central,Trn_Type," & cCStr("Vehicle_YN") & " as Vehicle, " & cCStr("Spare_YN") & " as Spare " & _
        "From TaxForms T  where T.FormTrnType >0 Order by Form_Code,Trn_Type,Vehicle_YN,Spare_YN"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eRef()
On Error GoTo ELoop
    RsTaxAc.Requery
    Master.Requery
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean, MsgStr$, InValidCol As Byte
Dim Rst As ADODB.Recordset, TmpStr As String, mRecForm$
On Error GoTo ELoop
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave(FGrid.Col) = False Then
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

    If (FGrid.Rows = 2 And FGrid.TextMatrix(1, Col_FormNo) = "") Or FGrid.Rows < 2 Then MsgBox "Please Fill Form Receipt/Issue Details", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = Col_FormNo: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
    For I = 1 To FGrid.Rows - 1
        If Trim(FGrid.TextMatrix(I, Col_FormNo)) <> "" Then
'            MsgStr = "Form No.": InValidCol = Col_FormNo
            If Trim(FGrid.TextMatrix(I, Col_RecDate)) = "" Then
                MsgStr = "Receipt Date": InValidCol = Col_RecDate
    '        ElseIf FGrid.TextMatrix(I, Col_IssDate) <> "" Then
    '            MsgStr = "Issue Date": InValidCol = Col_issdate
    '        ElseIf FGrid.TextMatrix(I, Col_SubCode) <> "" Then
    '            MsgStr = "Issue To Party": InValidCol = Col_subcode
    '        ElseIf FGrid.TextMatrix(I, Col_CurrentStatus) <> "" Then
    '            MsgStr = "Current Status": InValidCol = Col_currentstatus
            End If
        End If
        If Len(MsgStr) > 0 Then
            MsgBox MsgStr & "is Reqiured in Serial No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = InValidCol: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
        End If
    Next
    GCn.BeginTrans
    mTrans = True
    For I = 1 To FGrid.Rows - 1
        GSQL = ""
        If FGrid.TextMatrix(I, Col_AddEdit) = "E" Then
            'Case of Edit
            GSQL = "update TaxFormStk set RecDate=" & ConvertDate(FGrid.TextMatrix(I, Col_RecDate)) & _
                ", IssDate=" & ConvertDate(FGrid.TextMatrix(I, Col_IssDate)) & _
                ", SubCode='" & FGrid.TextMatrix(I, Col_SubCode) & _
                "', CurrentStatus=" & FormCurrStatus(FGrid.TextMatrix(I, Col_CurrentStatus)) & _
                ", Remarks='" & FGrid.TextMatrix(I, Col_Remarks) & _
                "', Site_Code='" & PubSiteCode & "', U_Name='" & pubUName & _
                "',U_EntDt=# " & PubServerDate & "#,U_AE='E' " & _
                " where Form_Code='" & Txt(FormCode).TEXT & _
                "' and FormNo='" & FGrid.TextMatrix(I, Col_FormNo) & "'"
        ElseIf FGrid.TextMatrix(I, Col_AddEdit) = "D" Then
            'Case of Delete
            GSQL = "Delete from TaxFormStk " & _
                " where Form_Code='" & Txt(FormCode).TEXT & _
                "' and FormNo='" & Txt(Col_FormNo).TEXT & "'"
        ElseIf Trim(FGrid.TextMatrix(I, Col_FormNo)) <> "" And FGrid.TextMatrix(I, Col_AddEdit) = "" Then
            'Case of Insert
            mRecForm = IIf(FGrid.TextMatrix(I, Col_RecFrom) = "", "P", left(FGrid.TextMatrix(I, Col_RecFrom), 1))
            GSQL = "Insert into TaxFormStk (Form_Code,FormNo,RecFrom,RecDate," & _
                "IssDate,SubCode,CurrentStatus," & _
                "Remarks,Site_Code,U_Name,U_EntDt,U_AE)" & _
                " values('" & Txt(FormCode).TEXT & "','" & FGrid.TextMatrix(I, Col_FormNo) & "', '" & mRecForm & "', " & ConvertDate(FGrid.TextMatrix(I, Col_RecDate)) & _
                ", " & ConvertDate(FGrid.TextMatrix(I, Col_IssDate)) & ", '" & FGrid.TextMatrix(I, Col_SubCode) & "'," & FormCurrStatus(FGrid.TextMatrix(I, Col_CurrentStatus)) & _
                ", '" & FGrid.TextMatrix(I, Col_Remarks) & "', '" & PubSiteCode & "', '" & pubUName & "'," & ConvertDate(PubServerDate) & ", 'A' )"
        End If
        If GSQL <> "" Then
            GCn.Execute GSQL
        End If
    Next
    GCn.CommitTrans
    mTrans = False
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To Txt.Count - 1
            Txt(I).BackColor = CtrlBColOrg
            Txt(I).ForeColor = CtrlFColOrg
        Next
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
On Error GoTo ELoop
Ctrl_GetFocus Txt(Index)
TxtGrid(0).Visible = False
Grid_Hide
Select Case Index
    Case FormSeries
'        ListArray = Array("General", "Warranty")
'        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 2)
    Case SrlNoFrom
'        If RsJob.RecordCount = 0 Or (RsJob.EOF = True Or RsJob.BOF = True) Or Txt(Index).Text = "" Then Exit Sub
'        If Txt(Index).Text <> RsJob!Name Then
'            RsJob.MoveFirst
'            RsJob.FIND "Name ='" & Txt(Index).Text & "'"
'        End If
    Case SrlNoTo
'        If RsMech.RecordCount = 0 Or (RsMech.EOF = True Or RsMech.BOF = True) Or Txt(Index).Text = "" Then Exit Sub
'        If Txt(Index).Text <> RsMech!Name Then
'            RsMech.MoveFirst
'            RsMech.FIND "Name ='" & Txt(Index).Text & "'"
'        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Dim I As Integer, j As Long, mFormNo As String, RowsBeforAdd As Integer, mSrlNo As Integer
Dim mFormFind As Boolean

Select Case Index
    Case SrlNoFrom
        NumDown Txt(Index), KeyCode, 6, 0
    Case SrlNoTo
        NumDown Txt(Index), KeyCode, 6, 0
End Select
    If FrmList.Visible = False And DGTaxAc.Visible = False Then
         If TopCtrl1.TopText2.CAPTION = "Add" Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
                Ctrl_DownKeyDown KeyCode, Shift
                
                If Index = RectDate Then
                    Txt_LostFocus RectDate
                    Txt(FormSeries).Enabled = False
                    Txt(SrlNoFrom).Enabled = False
                    Txt(SrlNoTo).Enabled = False
                    Txt(RectDate).Enabled = False
                    If MsgBox("Insert Form No. into Grid ? ", vbYesNo + vbCritical + vbDefaultButton1, "Insert Form Nos.") = vbYes Then
                        'Insert Rec in Grid
                        FGrid2.Rows = 1:    FGrid.Redraw = False
                        If FGrid.TextMatrix(FGrid.Rows - 1, Col_FormNo) = "" Then
                            If FGrid.Rows = 2 Then
                                FGrid.Rows = 1
                            Else
                                FGrid.RemoveItem (FGrid.Rows - 1)
                            End If
                        End If
                        RowsBeforAdd = (FGrid.Rows - 1): mSrlNo = Val(FGrid.TextMatrix(FGrid.Rows - 1, Col_SrNo))
                        
                        For j = Val(Txt(SrlNoFrom)) To Val(Txt(SrlNoTo))
                            mFormFind = False
                            mFormNo = Trim(Txt(FormSeries)) & j
                            If RowsBeforAdd > 1 Then
                                For I = 1 To RowsBeforAdd
                                    'Checking for Already Existing form No
                                    If UCase(mFormNo) = UCase(FGrid.TextMatrix(I, Col_FormNo)) Then
                                        'Display Already Existing form No
                                        FGrid2.AddItem (FGrid.TextMatrix(I, Col_SrNo) & ". " & Chr(9) & mFormNo)
                                        mFormFind = True: Exit For
                                    End If
                                Next I
                            End If
                            If mFormFind = False Then
                                mSrlNo = mSrlNo + 1
                                FGrid.AddItem (mSrlNo & Chr(9) & mFormNo & Chr(9) & "Deptt." & Chr(9) & Txt(RectDate))
                            End If
                        Next j
                        
                        FGrid.FixedRows = 1:    FGrid.Redraw = True
                        If FGrid2.Rows > 1 Then
                            FGrid2.FixedRows = 1: FrmDup.Visible = True
                            FrmDup.ZOrder 0
                            FGrid2.SetFocus
                        End If
                    End If
                    If FrmDup.Visible = False Then FGrid.SetFocus: FGrid_GotFocus
                End If
            End If
        End If
    End If
Exit Sub
ELoop:
    FGrid.Redraw = True
    CheckError
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
On Error GoTo ELoop
If keyascii = 39 Then keyascii = 0: Exit Sub
Select Case Index
    Case SrlNoFrom
        NumPress Txt(Index), keyascii, 6, 0
    Case SrlNoTo
        NumPress Txt(Index), keyascii, 6, 0
End Select
Exit Sub

ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
'    Select Case Index
'    Case DocType
'        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
'    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
'Validate- >LostFocus
Dim Rst As ADODB.Recordset, I As Integer
On Error GoTo ELoop
Select Case Index
    Case SrlNoFrom
        If Len(Trim(Txt(SrlNoFrom).TEXT)) = 0 Then
            MsgBox "From Serial No. is blank", vbOKOnly, "Validation Check"
            Cancel = True
        End If
    Case SrlNoTo
        If Len(Trim(Txt(SrlNoTo).TEXT)) = 0 Then
            MsgBox "To Serial No. is blank", vbOKOnly, "Validation Check"
            Cancel = True
        Else
            If Val(Txt(SrlNoTo)) < Val(Txt(SrlNoFrom)) Then
                MsgBox "To Serial No. " & Txt(SrlNoTo) & " < " & Txt(SrlNoFrom), vbOKOnly, "Validation Check"
                Cancel = True
            End If
        End If
    Case RectDate
        If Val(Txt(SrlNoTo)) + Val(Txt(SrlNoFrom)) <> 0 Then
            If Len(Trim(Txt(RectDate).TEXT)) = 0 Then
                MsgBox "Blank Date", vbOKOnly, "Validation Check"
                Cancel = True
            End If
        End If
        Txt(Index).TEXT = RetDate(Txt(Index))
End Select
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
Ctrl_GetFocus TxtGrid(Index)
Grid_Hide
FGrid.CellBackColor = CellBackColLeave  'CellBackColorOrg
TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
Select Case FGrid.Col
    Case Col_FormNo
        TxtGrid(0).MaxLength = 10
    Case Col_RecDate
        TxtGrid(0).MaxLength = 15
    Case Col_SubName
        TxtGrid(0).MaxLength = 50
        If RsTaxAc.RecordCount = 0 Or (RsTaxAc.EOF = True Or RsTaxAc.BOF = True) Then Exit Sub
        RsTaxAc.Sort = "Name"
        If FGrid.TextMatrix(FGrid.Row, Col_SubCode) <> "" Then
            RsTaxAc.MoveFirst
            RsTaxAc.FIND "Code ='" & FGrid.TextMatrix(FGrid.Row, Col_SubCode) & "'"
            If RsTaxAc.EOF = True Then RsTaxAc.MoveFirst
        End If
    Case Col_IssDate
        TxtGrid(0).MaxLength = 15
    Case Col_CurrentStatus
        TxtGrid(0).MaxLength = 10
        ListArray = Array("N.A.", "Issue", "Returned", "Damaged", "Lost")
        Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 5)
    Case Col_Remarks
        TxtGrid(0).MaxLength = 25
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        TxtGrid(0).TEXT = TxtGrid(0).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        TxtGrid(0).Visible = False
        FGrid.SetFocus
        Exit Sub
    End If
    Select Case FGrid.Col
        Case Col_FormNo, Col_RecDate, Col_IssDate, Col_Remarks
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave(FGrid.Col) = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, FGrid.Cols
                End If
            End If
        Case Col_SubName
            If DGTaxAc.Visible = False Then DGridColSwap DGTaxAc, 0
            DGridTxtKeyDown DGTaxAc, TxtGrid, Index, RsTaxAc, KeyCode, True, 1, frmSubGroup, "frmSubGroup"
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave(FGrid.Col) = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, FGrid.Cols
                Else
                    TxtGrid(0).SetFocus
                End If
            End If
        Case Col_CurrentStatus
            ListView_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left + TxtGrid(0).width, TxtGrid(0).top, 1250, 1300
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave(FGrid.Col) = True Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, FGrid.Cols
                Else
'                    TxtGrid_LostFocus 0
                    TxtGrid(0).SetFocus
                End If
            End If
'        Case Col_Remarks
'            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
'                If TxtGridLeave(FGrid.Col) = True Then
'                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, FGrid.Cols
'                Else
'                    TxtGrid_LostFocus 0
'                    TxtGrid(0).SetFocus
'                End If
'            End If
        End Select
'    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, keyascii As Integer)
On Error GoTo ELoop
CheckQuote keyascii
Select Case FGrid.Col
    Case Col_FormNo
'        If DGTaxAc.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsPart, KeyAscii, "Code"
    Case Col_SubName
        If DGTaxAc.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsTaxAc, keyascii, "Name"
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case FGrid.Col
    Case Col_FormNo
'        If KeyCode <> 13 And DGTaxAc.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsPart, KeyCode, "Code", True
    Case Col_SubName
        If KeyCode <> 13 And DGTaxAc.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, RsTaxAc, KeyCode, "Name", True
    Case Col_CurrentStatus
        If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
        ListView_KeyUp ListView, TxtGrid, Index, KeyCode, mListItem
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Select Case FGrid.Col
    Case Col_FormNo
        'Check duplicate form no
    Case Col_SubName
        If FGrid.TextMatrix(FGrid.Rows - 1, Col_FormNo) <> "" Then FGrid.AddItem FGrid.Rows
    Case Col_CurrentStatus
        TxtGrid(Index).TEXT = ListView.SelectedItem.TEXT
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(Index).TEXT
    Case Col_Remarks
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(Index).TEXT
        If FGrid.TextMatrix(FGrid.Rows - 1, Col_FormNo) <> "" Then FGrid.AddItem FGrid.Rows
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGridLeave(FGridCol) As Boolean
Dim j As Integer
Select Case FGridCol
    Case Col_FormNo
        If ChkDuplicate(Col_FormNo) = False Then TxtGridLeave = False: Exit Function
        If FGrid.TextMatrix(FGrid.Row, FGridCol) <> TxtGrid(0) Then
            FGrid.TextMatrix(FGrid.Row, FGridCol) = TxtGrid(0): CellFontColor FGrid: FGridAddEditDel
        End If
        If FGrid.TextMatrix(FGrid.Rows - 1, 1) <> "" Then FGrid.AddItem FGrid.Rows
    Case Col_RecDate
        If FGrid.TextMatrix(FGrid.Row, FGridCol) <> TxtGrid(0) Then
            FGrid.TextMatrix(FGrid.Row, FGridCol) = RetDate(TxtGrid(0)): CellFontColor FGrid: FGridAddEditDel
        End If
    Case Col_SubName
        If RsTaxAc.RecordCount = 0 Or (RsTaxAc.EOF = True Or RsTaxAc.BOF = True) Or TxtGrid(0).TEXT = "" Then
            If FGrid.TextMatrix(FGrid.Row, FGridCol) <> TxtGrid(0) Then
                FGrid.TextMatrix(FGrid.Row, Col_SubCode) = ""
                FGrid.TextMatrix(FGrid.Row, Col_SubName) = ""
                CellFontColor FGrid: FGridAddEditDel
            End If
        Else
            If UCase(FGrid.TextMatrix(FGrid.Row, FGridCol)) <> UCase(RsTaxAc!Name) Then
                FGrid.TextMatrix(FGrid.Row, Col_SubCode) = RsTaxAc!Code
                FGrid.TextMatrix(FGrid.Row, Col_SubName) = RsTaxAc!Name
                CellFontColor FGrid: FGridAddEditDel
            End If
        End If
    Case Col_IssDate
        If RetDate(TxtGrid(0)) <> "" Then
            If CDate(RetDate(TxtGrid(0))) < CDate(FGrid.TextMatrix(FGrid.Row, Col_RecDate)) Then
                MsgBox "Issue Date < Receipt Date", vbOKOnly, "Validation"
                TxtGridLeave = False: Exit Function
            End If
        End If
        If FGrid.TextMatrix(FGrid.Row, FGridCol) = "" And TxtGrid(0) <> "" Then
            FGrid.TextMatrix(FGrid.Row, Col_CurrentStatus) = "Issue"
        ElseIf FGrid.TextMatrix(FGrid.Row, FGridCol) <> "" And TxtGrid(0) = "" Then
            FGrid.TextMatrix(FGrid.Row, Col_CurrentStatus) = "N.A."
        End If
        If FGrid.TextMatrix(FGrid.Row, FGridCol) <> TxtGrid(0) Then
            FGrid.TextMatrix(FGrid.Row, FGridCol) = RetDate(TxtGrid(0)): CellFontColor FGrid: FGridAddEditDel
        End If
    Case Col_CurrentStatus, Col_Remarks
        If FGrid.TextMatrix(FGrid.Row, FGridCol) <> TxtGrid(0) Then
            FGrid.TextMatrix(FGrid.Row, FGridCol) = TxtGrid(0): CellFontColor FGrid: FGridAddEditDel
        End If
End Select

TxtGridLeave = True
TxtGrid(0).Visible = False
FGrid.SetFocus
End Function

Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
On Error GoTo ELoop
'                    FGrid.CellFontName = "Roman"
'                    FGrid.CellFontSize = 10
'                                FGrid.RowHeight(2) = PubGridRowHeight + 200
'                                FGrid.RowHeight(3) = PubGridRowHeight + 400
'                                FGrid.RowHeight(4) = PubGridRowHeight + 600

If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Dim TempStatVal As Byte
    Select Case FGrid.Col
        Case Col_FormNo
            If FGrid.TextMatrix(FGrid.Row, Col_AddEdit) = "" Then
                If Trim(FGrid.TextMatrix(FGrid.Row, Col_SubName)) = "" Then GridDblClick Me, FGrid, TxtGrid, 0
            End If
        Case Col_RecDate
            If Trim(FGrid.TextMatrix(FGrid.Row, Col_SubName)) = "" Then GridDblClick Me, FGrid, TxtGrid, 0
        Case Col_SubName, Col_IssDate
            TempStatVal = FormCurrStatus(FGrid.TextMatrix(FGrid.Row, Col_CurrentStatus))
            If TempStatVal <= 2 Then GridDblClick Me, FGrid, TxtGrid, 0
        Case Col_CurrentStatus, Col_Remarks
            GridDblClick Me, FGrid, TxtGrid, 0
    End Select
    TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_EnterCell()
Dim I As Integer
    I = FGrid.Col
    FGrid.Col = 0
    FGrid.CellBackColor = CellBackColEnter
    FGrid.Col = I
    FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
    FGrid_EnterCell
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And FGrid.Row > 1 Then
        If Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
            FGrid.CellBackColor = CellBackColLeave
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
        FGrid.CellBackColor = CellBackColLeave
        SendKeysA vbKeyTab, True
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then 'Delete Key
        Select Case FGrid.Col
            Case Col_RecDate
                If FxAllowEdit(FGrid.Col) = True And FGrid.TextMatrix(FGrid.Row, Col_SubName) = "" Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                End If
            Case Col_SubName
                If FGrid.TextMatrix(FGrid.Row, Col_IssDate) = "" Then
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
                End If
            Case Col_IssDate, Col_CurrentStatus, Col_Remarks
                If FxAllowEdit(FGrid.Col) = True Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        End Select
    ElseIf KeyCode = vbKeyReturn Then
        TAddMode = False
        FGrid_KeyPress KeyCode
'        Select Case FGrid.Col
'            Case Col_FormNo, Col_RecDate
'                If FGrid.TextMatrix(FGrid.Row, Col_SubName) = "" Then GridDblClick Me, FGrid, TxtGrid, 0
'            Case Col_FormNo, Col_RecDate, Col_SubName, Col_IssDate, Col_CurrentStatus, Col_Remarks
'                GridDblClick Me, FGrid, TxtGrid, 0
'        End Select
'        TAddMode = False
    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)
On Error GoTo ELoop
Dim TempStatVal As Byte
Select Case FGrid.Col
    Case Col_FormNo
        If FGrid.TextMatrix(FGrid.Row, Col_AddEdit) = "" Then
            If Trim(FGrid.TextMatrix(FGrid.Row, Col_SubName)) = "" Then Get_Text Me, FGrid, TxtGrid, 0, False, keyascii
        End If
    Case Col_RecDate
        If Trim(FGrid.TextMatrix(FGrid.Row, Col_SubName)) = "" Then Get_Text Me, FGrid, TxtGrid, 0, False, keyascii
    Case Col_SubName, Col_IssDate
        TempStatVal = FormCurrStatus(FGrid.TextMatrix(FGrid.Row, Col_CurrentStatus))
        If TempStatVal <= 2 Then Get_Text Me, FGrid, TxtGrid, 0, False, keyascii
    Case Col_CurrentStatus, Col_Remarks
        Get_Text Me, FGrid, TxtGrid, 0, False, keyascii
End Select
If keyascii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Dim I As Integer, mRowVisible As Boolean, mSrlNo As Integer
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGrid.Row >= 1 Then
            If MsgBox("Are You Sure To Delete Current Row?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                FGrid.Redraw = False
                If FGrid.TextMatrix(FGrid.Row, Col_AddEdit) = "" Then
                    FGrid.RemoveItem (FGrid.Row)
                Else
                    FGrid.RowHeight(FGrid.Row) = 0
                End If
                For I = 1 To FGrid.Rows - 1
                    If FGrid.RowHeight(I) > 0 Then
                        mRowVisible = True: Exit For
                    End If
                Next
                If mRowVisible = False Then
                    FGrid.AddItem ""
                End If
                For I = 1 To FGrid.Rows - 1
                    If FGrid.RowHeight(I) > 0 Then
                        mSrlNo = mSrlNo + 1
                        FGrid.TextMatrix(I, Col_SrNo) = mSrlNo
                    End If
                Next
            End If
            FGrid.Redraw = True
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
Dim I As Integer
    I = FGrid.Col
    FGrid.Col = 0
    FGrid.CellBackColor = FGrid.BackColorFixed
    FGrid.Col = I
    FGrid.CellBackColor = CellBackColLeave 'CellBackColorOrg
End Sub

Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid_LeaveCell
End Sub

Private Function FormCurrStatus(ByRef CheckVal As Variant)
    If IsNumeric(CheckVal) Then
        If CheckVal = 1 Then
            FormCurrStatus = "Issue"
        ElseIf CheckVal = 2 Then
            FormCurrStatus = "Returned"
        ElseIf CheckVal = 3 Then
            FormCurrStatus = "Damaged"
        ElseIf CheckVal = 4 Then
            FormCurrStatus = "Lost"
        Else
            FormCurrStatus = "N.A."
        End If
    Else
        If CheckVal = "Issue" Then
            FormCurrStatus = 1
        ElseIf CheckVal = "Returned" Then
            FormCurrStatus = 2
        ElseIf CheckVal = "Damaged" Then
            FormCurrStatus = 3
        ElseIf CheckVal = "Lost" Then
            FormCurrStatus = 4
        Else
            FormCurrStatus = 0
        End If
    End If
End Function

Private Function FxAllowEdit(FGridCol)
    FxAllowEdit = False
    If TopCtrl1.TopText2 = "Add" Then
        If FGrid.TextMatrix(FGrid.Row, Col_AddEdit) = "" Then
            FxAllowEdit = True
        Else
            FxAllowEdit = False
        End If
    ElseIf TopCtrl1.TopText2 = "Edit" Then
        FxAllowEdit = True
    End If
End Function

Private Sub FGridAddEditDel()
If FGrid.TextMatrix(FGrid.Row, Col_AddEdit) <> "" Then FGrid.TextMatrix(FGrid.Row, Col_AddEdit) = "E"
End Sub

Private Sub CellFontColor(FG As MSHFlexGrid)
FG.CellForeColor = CellForeColLeave
End Sub

Private Sub DispTrn()
Dim RsTemp As ADODB.Recordset, TrnDate As Date, MsgStr$
    FrameTrn.left = (Me.width - FrameTrn.width) / 2
    FrameTrn.top = Shape1.top - 135 'FGrid.top
    FrameTrn.height = FrameTrn.height - 15
    ReDim Preserve GridRow1(0)
    GridRow1(0) = 0
    

If Txt(TrnType) <> "Form31" Then
    If Txt(FormTrnType) = "Receipt" Then '  "Issue"
        If FGrid.TextMatrix(FGrid.Row, Col_RecDate) = "" Then
            MsgStr = "Receipt Date is empty!"
        Else
            TrnDate = CDate(FGrid.TextMatrix(FGrid.Row, Col_RecDate))
        End If
    Else
        If FGrid.TextMatrix(FGrid.Row, Col_IssDate) = "" Then
            MsgStr = "Issue Date is empty!"
        Else
            TrnDate = CDate(FGrid.TextMatrix(FGrid.Row, Col_IssDate))
        End If
    End If
    If MsgStr <> "" Then MsgBox MsgStr, vbOKOnly, "Receipt / Issue Date": Exit Sub
    
    Label3(2).CAPTION = "*" & AppFor & " " & Txt(TrnType) & " Bill Details*"
    Label3(2).left = FrameTrn.width - Label3(2).width
    Label3(12).CAPTION = "Sl.No.: " & FGrid.TextMatrix(FGrid.Row, Col_SrNo)
    Label3(11).CAPTION = "Form No.   : " & FGrid.TextMatrix(FGrid.Row, Col_FormNo)
    Label3(13).CAPTION = "Party Name : " & FGrid.TextMatrix(FGrid.Row, Col_SubName)
    
    If AppFor.CAPTION = "Spare" Then
        TabName = IIf(Txt(TrnType) = "Purchase", "SP_Purch", "SP_Sale")
        GSQL = "select FormNo,FormIssRecDate,Party_Code,docid,right(docid,13) as DocNo,V_Date,format(Total_Amt,'0.00') as TotalAmt,FormNo as OldFrmNo " & _
            " from " & TabName & _
            "  where form_code='" & Txt(FormCode) & _
            "' and (formno='" & FGrid.TextMatrix(FGrid.Row, Col_FormNo) & "' or formno='' or isnull(formno)) " & _
            "  and v_date >= " & ConvertDate(PubStartDate) & _
            "  and v_date >= " & ConvertDate(TrnDate) & _
            "  Order by V_Date,DocId"
    Else    'Vehicle
        If Txt(TrnType) = "Purchase" Then
            TabName = "Veh_Purch1"
            GSQL = "select FormNo,FormIssRecDate,PARTYCODE,docid,right(docid,13) as DocNo, V_Date,format(Tot_Amount,'0.00') as TotalAmt,FormNo as OldFrmNo " & _
                " from " & TabName & _
                "  where form_code='" & Txt(FormCode) & _
                "' and (formno='" & FGrid.TextMatrix(FGrid.Row, Col_FormNo) & "' or formno='' or isnull(formno)) " & _
                "  and v_date >= " & ConvertDate(PubStartDate) & _
                "  and v_date >= " & ConvertDate(TrnDate) & _
                "  Order by V_Date,DocId"
        Else
            TabName = "Veh_Order"
            GSQL = "select FormNo,FormIssRecDate,PartyCode,Inv_DocId,right(Inv_DocId,13) as DocNo,Inv_Date,format(Net_Amount,'0.00') as TotalAmt,FormNo as OldFrmNo " & _
                " from " & TabName & _
                "  where form_code='" & Txt(FormCode) & _
                "  and (formno='" & FGrid.TextMatrix(FGrid.Row, Col_FormNo) & "' or formno='' or isnull(formno)) " & _
                "  and Inv_Date >= " & ConvertDate(PubStartDate) & _
                "  and Inv_Date >= " & ConvertDate(TrnDate) & _
                "  Order by Inv_Date,Inv_DocId"
        End If
    End If
    Set RsTemp = GCn.Execute(GSQL)
    If RsTemp.RecordCount > 0 Then
        Set FGridTrn.DataSource = RsTemp
    Else
        FGridTrn.Rows = 1
        FGridTrn.AddItem ""
        FGrid.FixedRows = 1
    End If
    Set RsTemp = Nothing
    FrameTrn.Visible = True
    FrameTrn.ZOrder 0
End If
    With FGridTrn
        .left = 0
        .width = FrameTrn.width
        .Cols = 8
        .FixedRows = 1
        .height = FGrid.RowHeight(0) * 16
        .TextMatrix(0, Trn_FormNo) = " Form No."
        .ColWidth(Trn_FormNo) = 1400
        .TextMatrix(0, Trn_IssRecDate) = " IssRec Date"
        .ColWidth(Trn_IssRecDate) = 1800
        .TextMatrix(0, Trn_SubCode) = "SubCode"
        .ColWidth(Trn_SubCode) = 0
        .TextMatrix(0, Trn_DocID) = "DocID"
        .ColWidth(Trn_DocID) = 0
        .TextMatrix(0, Trn_DocNo) = "Bill No."
        .ColWidth(Trn_DocNo) = 1530

        .TextMatrix(0, Trn_VDate) = "Bill Date"
        .ColWidth(Trn_VDate) = 1410

        .TextMatrix(0, Trn_BillAmt) = "  Bill Amount"
        .ColAlignment(Trn_BillAmt) = flexAlignRightCenter
        .ColWidth(Trn_BillAmt) = 1200
        
        .ColWidth(Trn_OldFrmNo) = 0
        .ColWidth(8) = 0
    End With

    FGridTrn.SetFocus
End Sub

Private Sub FGridMainUpd()
On Error GoTo ELoop
Dim I As Integer, mFormNo$, mIssRecDate$, mTrans As Boolean
'Save FGridTrn values directly in Table
GCn.BeginTrans
mTrans = True
For I = 1 To FGridTrn.Rows - 1
    If FGridTrn.TextMatrix(I, Trn_FormNo) <> FGridTrn.TextMatrix(I, Trn_OldFrmNo) Then
        mFormNo = FGridTrn.TextMatrix(I, Trn_FormNo)
        mIssRecDate = FGridTrn.TextMatrix(I, Trn_IssRecDate)
        
        If AppFor.CAPTION = "Spare" Then
            GSQL = "Update " & TabName & " set FormNo='" & mFormNo & _
                "', FormIssRecDate=" & ConvertDate(mIssRecDate) & " where DocID='" & FGridTrn.TextMatrix(I, Trn_DocID) & "'"
        Else    'Vehicle
            If Txt(TrnType) = "Purchase" Then
                GSQL = "Update " & TabName & " set FormNo=" & mFormNo & _
                    ", FormIssRecDate=" & ConvertDate(mIssRecDate) & " where DocID='" & FGridTrn.TextMatrix(I, Trn_DocID) & "'"
            Else
                GSQL = "Update " & TabName & " set FormNo=" & mFormNo & _
                    ", FormIssRecDate=" & ConvertDate(mIssRecDate) & " where Inv_DocID='" & FGridTrn.TextMatrix(I, Trn_DocID) & "'"
            End If
        End If
        GCn.Execute (GSQL)
        TopCtrl1.tCancel = False
    End If
Next
GCn.CommitTrans
mTrans = False
FGridTrnModified = False
FGridTrn.Rows = 1
FGridTrn.AddItem ""
FGridTrn.FixedRows = 1
Exit Sub

ELoop:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub

