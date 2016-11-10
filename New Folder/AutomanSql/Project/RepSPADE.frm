VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form RepSPADE 
   BackColor       =   &H00C8E8DA&
   Caption         =   "ReprtForm"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11820
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   11820
   Begin VB.Frame frmDetail 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Detail"
      Height          =   1380
      Left            =   120
      TabIndex        =   21
      Top             =   345
      Visible         =   0   'False
      Width           =   11565
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Rep. :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   225
         Index           =   10
         Left            =   165
         TabIndex        =   32
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   225
         Index           =   9
         Left            =   2865
         TabIndex        =   31
         Top             =   825
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address2"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   7
         Left            =   165
         TabIndex        =   30
         Top             =   570
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VNo && Date"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   0
         Left            =   5835
         TabIndex        =   29
         Top             =   90
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   225
         Index           =   8
         Left            =   165
         TabIndex        =   28
         Top             =   825
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   225
         Index           =   1
         Left            =   165
         TabIndex        =   27
         Top             =   90
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   2
         Left            =   165
         TabIndex        =   26
         Top             =   330
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   3
         Left            =   5835
         TabIndex        =   25
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Follow up:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   225
         Index           =   6
         Left            =   5835
         TabIndex        =   24
         Top             =   825
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp.Del.Date :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   5
         Left            =   8835
         TabIndex        =   23
         Top             =   570
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MAP :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   4
         Left            =   5835
         TabIndex        =   22
         Top             =   570
         Width           =   525
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   4410
      Left            =   585
      TabIndex        =   20
      Top             =   1455
      Visible         =   0   'False
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   7779
      _Version        =   393216
      BackColor       =   15525079
      Cols            =   8
      BackColorFixed  =   14940925
      ForeColorFixed  =   192
      BackColorSel    =   14667992
      ForeColorSel    =   12582912
      BackColorBkg    =   14737632
      BackColorUnpopulated=   14865856
      GridColor       =   14940925
      GridColorFixed  =   12632319
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "SrNo."
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox TxtGrid1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.TextBox TxtSearch1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   240
      HideSelection   =   0   'False
      Left            =   1125
      TabIndex        =   18
      Top             =   1035
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Report"
      DownPicture     =   "RepSPADE.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   6075
      Width           =   1290
   End
   Begin VB.CommandButton BTNEXIT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E&xit"
      DownPicture     =   "RepSPADE.frx":3132
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5910
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exit Form"
      Top             =   6075
      Width           =   1290
   End
   Begin VB.PictureBox Pic 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11820
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6045
      Width           =   11820
      Begin VB.Label LblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "LblTitle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   7230
         TabIndex        =   16
         Top             =   0
         Width           =   4470
      End
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   7290
      TabIndex        =   13
      Top             =   -195
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   405
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   150
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
         BackColor       =   16777152
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
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   375
      TabIndex        =   12
      Top             =   510
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   240
      HideSelection   =   0   'False
      Left            =   -135
      TabIndex        =   11
      Top             =   1500
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   5040
      TabIndex        =   7
      Top             =   3900
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   5
      Top             =   3885
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   2
      Left            =   5685
      TabIndex        =   3
      Top             =   1830
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1860
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   1785
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2910
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   12632256
      ForeColorSel    =   128
      BackColorBkg    =   14873572
      GridColor       =   8438015
      GridColorFixed  =   192
      GridColorUnpopulated=   16711935
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   2
      Left            =   4965
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   2910
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   12632256
      ForeColorSel    =   128
      BackColorBkg    =   14873572
      GridColor       =   8438015
      GridColorFixed  =   192
      GridColorUnpopulated=   16711935
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   4
      Left            =   4995
      TabIndex        =   8
      Top             =   3825
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   2910
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   12632256
      ForeColorSel    =   128
      BackColorBkg    =   14873572
      GridColor       =   8438015
      GridColorFixed  =   192
      GridColorUnpopulated=   16711935
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1650
      Index           =   3
      Left            =   165
      TabIndex        =   6
      Top             =   3825
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   2910
      _Version        =   393216
      BackColor       =   15595518
      ForeColor       =   128
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   12632256
      ForeColorSel    =   128
      BackColorBkg    =   14873572
      GridColor       =   8438015
      GridColorFixed  =   192
      GridColorUnpopulated=   16711935
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1560
      Left            =   1605
      TabIndex        =   0
      Top             =   525
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2752
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16512
      Rows            =   5
      Cols            =   3
      FixedRows       =   0
      BackColorFixed  =   13166810
      ForeColorFixed  =   16384
      BackColorSel    =   16711680
      ForeColorSel    =   12648447
      BackColorBkg    =   13166810
      GridColor       =   13166810
      GridColorFixed  =   13166810
      GridColorUnpopulated=   12648447
      GridLinesFixed  =   1
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   661
   End
End
Attribute VB_Name = "RepSPADE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CellBackColLeave As String = &HFFFFFF
Private Const CellBackColEnter As String = &HFFFFC0
Private Const CellBackColLeave1 As String = &HEDF7FE
Private Const CellBackColEnter1 As String = &HFFFFC0
Private Const BackColorSelEnter As String = &HF8D7FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$
Dim Condstr$, CondStr1$
Dim mAdd As Boolean, mEdit As Boolean, mDel As Boolean, mPrn As Boolean
Dim Master As ADODB.Recordset
Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset
Dim RsGrid3 As ADODB.Recordset
Dim RsGrid4 As ADODB.Recordset
Dim RepTitle As String, RepName As String
Dim RepPrint As Boolean
Dim RstRep As ADODB.Recordset
Dim RstRep1 As ADODB.Recordset
Dim SubRep1 As Boolean
Private Const GridRowHeight As Integer = 270
'////////********VEHICLE***********////////////////////*****
Private Const SPADE_Report As Byte = 1

'*******************Constants For Fgrid1****************************************
Private Const Col_PDtNo As Byte = 1
Private Const Col_RefBy As Byte = 2
Private Const Col_Name As Byte = 3
Private Const Col_City = 4
Private Const Col_Phone = 5
Private Const Col_Veh1st As Byte = 6
Private Const Col_Bank = 7
Private Const Col_Govt As Byte = 8
Private Const Col_Status As Byte = 9
Private Const Col_Model As Byte = 10
Private Const Col_VehQty As Byte = 11
Private Const Col_Map As Byte = 12
Private Const Col_DeliDate As Byte = 13
Private Const Col_Follow As Byte = 14
Private Const Col_Area As Byte = 15
Private Const Col_VNo As Byte = 16
Private Const Col_VDate As Byte = 17
Private Const Col_Add As Byte = 18
Private Const Col_Add2 As Byte = 19
Private Const Col_Add3 As Byte = 20
Private Const Col_ModelDesc As Byte = 21
Private Const Col_DocID As Byte = 22
Private Const Col_AreaName As Byte = 23
Private Const Col_RepName As Byte = 24

'**********************************************************************

Private Const Date1 As Byte = 0
Private Const Date2 As Byte = 1
Private Const List1 As Byte = 2
Private Const List2 As Byte = 3
Private Const List3 As Byte = 4

Private Const Cat1 As Byte = 5
Private Const Cat2 As Byte = 6
Private Const Cat3 As Byte = 7
Private Const Cat4 As Byte = 8
Private Const Cat5 As Byte = 9


Public GRepFormName As String
Dim mLastRow As Integer
Dim mFirstRow As Integer
Dim mHelpGridNo
Dim GridKey As Integer
Dim TAddMode As Boolean
Dim ListArray As Variant
Dim GridString1 As String
Dim GridString2 As String
Dim GridString3 As String
Dim GridString4 As String
Dim GridRow1() As Integer
Dim GridRow2() As Integer
Dim GridRow3() As Integer
Dim GridRow4() As Integer
Dim mGridStartRow As Integer
Dim mGridEndRow As Integer

Private Const SprMrRct$ = "SXGR"           'Material Receipt
Private Const SprMrTrf$ = "SXGRT"          'Material Rectipt Transfer
Private Const SprSlChal$ = "SYSC"           'Sale Challan       LPS 24-09
Private Const SprTrfChal$ = "SYSCT"         'Transfer Issue     LPS 24-09
Private Const SprSlCsh$ = "SYSIC"          'Cash Sale
Private Const SprSlCre$ = "SYSIR"          'Credit Sale
Private Const WksSlCsh$ = "W_SIC"          'Cash Sale
Private Const WksSlCre$ = "W_SIR"          'Credit Sale
Private Const SprSlRetCsh$ = "SXSRC"       'Cash Sale Return
Private Const SprSlRetCre$ = "SXSRR"       'Credit Sale Return
Private Const SprSlTrfRet$ = "SXSRT"       'Transfer Issue Return
Private Const SprPurCsh$ = "SXPIC"         'Cash Purchase
Private Const SprPurCre$ = "SXPIR"         'Credit Purchase
Private Const SprPrRetCsh$ = "SYPRC"       'Purchase Return Cash
Private Const SprPrRetCre$ = "SYPRR"       'Purchase Return Credit
Private Const SprPrTrfRet$ = "SYPRT"       'Transfer Receipt Return
Private Const SprQuotation$ = "S_QU"       'Spare Quotation
Private Const WksEst$ = "W_EST"       'Workshop Estimation
Private Const WksPro$ = "W_PL"       'Workshop Proforma Labour
Private Const WksGenReq$ = "W_RG"       'Workshop General Reqisition
Private Const WksReqWrt$ = "W_RW"       'Workshop Warranti Reqisition
Dim mListItem As ListItem

Private Sub btnexit_Click()
    Unload Me
End Sub

Private Sub BTNPRINT_Click()
On Error GoTo ELoop
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 1)) = False Then: Exit Sub
    If FGrid.TextMatrix(List1, 1) = "Status" Then
        If IsNotBlank(List2, FGrid.TextMatrix(List2, 1)) = False Then
            FGrid.TextMatrix(List2, 1) = "Cold"
        End If
    End If
    TopCtrl1.Tag = PubUParam
    TopCtrl1.TopText1 = ""
    TopCtrl1.TopText2 = "Browse": TopCtrl1.TopText2.ForeColor = RGB(0, 0, 255)
    If InStr(PubUParam, "A") <> 0 Then mAdd = True Else mAdd = False
    If InStr(PubUParam, "E") <> 0 Then TopCtrl1.tEdit = True Else TopCtrl1.tEdit = False
    If InStr(PubUParam, "D") <> 0 Then mDel = True Else mDel = False
    If InStr(PubUParam, "P") <> 0 Then mPrn = True Else mPrn = False
    
    Ini_Grid1
    Fill_Grid 100, "General"
    Disp_Text False
  Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Check1_Click(Index As Integer)
    If Check1(Index).Value = Unchecked Then
        GridSel(Index).Enabled = True
        If GridSel(Index).Rows > 1 Then
            GridSel(Index).Row = 1: GridSel(Index).Col = 1
        End If
    Else
        GridSel(Index).Enabled = False
        If GridSel(Index).Rows > 1 Then
            GridSel(Index).Row = 0: GridSel(Index).Col = 0
            GridSel(Index).RowSel = GridSel(Index).Rows - 1
        End If
    End If
End Sub

Private Sub Check1_GotFocus(Index As Integer)
    Check1(Index).BackColor = &HFF&
End Sub

Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub

Private Sub Check1_Validate(Index As Integer, Cancel As Boolean)
Check1(Index).BackColor = &H800000
End Sub

Private Sub FGrid1_RowColChange()
Label1(1).CAPTION = "Party Name: " & FGrid1.TextMatrix(FGrid1.Row, Col_Name)
Label1(2).CAPTION = "Address   : " & FGrid1.TextMatrix(FGrid1.Row, Col_Add)
Label1(7).CAPTION = "            " & FGrid1.TextMatrix(FGrid1.Row, Col_Add2)
Label1(8).CAPTION = "City      : " & FGrid1.TextMatrix(FGrid1.Row, Col_City)
Label1(9).CAPTION = "Area : " & FGrid1.TextMatrix(FGrid1.Row, Col_AreaName)
Label1(10).CAPTION = "Sales Rep.: " & FGrid1.TextMatrix(FGrid1.Row, Col_RepName)

Label1(0).CAPTION = "Quot.No. && Date : " & FGrid1.TextMatrix(FGrid1.Row, Col_PDtNo)
Label1(3).CAPTION = "Model     : " & FGrid1.TextMatrix(FGrid1.Row, Col_Model)
Label1(4).CAPTION = "Monthly AP: " & FGrid1.TextMatrix(FGrid1.Row, Col_Map)
Label1(6).CAPTION = "Follow Up : " & FGrid1.TextMatrix(FGrid1.Row, Col_Follow)
Label1(5).CAPTION = "Exp Del. Date: " & FGrid1.TextMatrix(FGrid1.Row, Col_DeliDate)
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte

WinSetting Me  ', 6885, 11500
FGrid.Visible = True
TopCtrl1.Visible = False

   Global_Grid
   TopCtrl1.TopText2 = "Add"
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub FGrid1_DblClick()
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
FGrid1_KeyPress vbKeyReturn
TAddMode = False
End Sub

Private Sub FGrid1_GotFocus()
    FGrid1.BackColorSel = BackColorSelEnter
    FGrid1.ForeColorSel = ForeColorSelEnter
    TxtSearch1.TEXT = ""
    txtgrid1(0).Visible = False
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
'Leave Cell-- > Enter Cell-- >KeyDown
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
        Case Col_Map, Col_DeliDate, Col_Follow
            FGrid1.CellForeColor = vbRed
            FGrid1.TextMatrix(FGrid1.Row, 0) = "¤"
    End Select
End If
KeyCode = 0
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
Select Case FGrid1.Col
    Case Col_Map
        txtgrid1(0).MaxLength = 15 '20
        Call Get_Text(Me, FGrid1, txtgrid1, 0, True, KeyAscii)
    Case Col_Follow
        txtgrid1(0).MaxLength = 20
        Call Get_Text(Me, FGrid1, txtgrid1, 0, True, KeyAscii)
    Case Col_DeliDate
        txtgrid1(0).MaxLength = 12
       Call Get_Text(Me, FGrid1, txtgrid1, 0, True, KeyAscii)
End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid1_LostFocus()
    FGrid1.BackColorSel = BackColorSelLeave
    FGrid1.ForeColorSel = FGrid1.ForeColor
End Sub

Private Sub FGrid1_Scroll()
txtgrid1(0).Visible = False
'Grid_Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
If GridSel(4).Visible = True Then Set RsGrid1 = Nothing
If GridSel(1).Visible = True Then Set RsGrid2 = Nothing
If GridSel(2).Visible = True Then Set RsGrid3 = Nothing
If GridSel(3).Visible = True Then Set RsGrid4 = Nothing
Set RstRep = Nothing
Set mListItem = Nothing
Set rpt = Nothing
Set Master = Nothing
End Sub

Private Sub GridSel_EnterCell(Index As Integer)
GridSel(Index).CellBackColor = CellBackColEnter1
End Sub

Private Sub GridSel_GotFocus(Index As Integer)
GridSel(Index).CellBackColor = CellBackColEnter1
End Sub

Private Sub GridSel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Integer
If KeyCode = 13 Then SendKeysA vbKeyTab, True
If GridSel(Index).Rows < 1 Then Exit Sub
If KeyCode = vbKeySpace And GridSel(Index).Col = 0 Then
    GridSel(Index).CellFontName = "WINGDINGS"
    GridSel(Index).CellFontSize = 14
    GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = IIf(GridSel(Index).TextMatrix(GridSel(Index).Row, 0) = "ü", " ", "ü")
    Select Case Index
        Case 1
            I = UBound(GridRow1) + 1
            ReDim Preserve GridRow1(I)
            GridRow1(I) = GridSel(Index).Row
        Case 2
            I = UBound(GridRow2) + 1
            ReDim Preserve GridRow2(I)
            GridRow2(I) = GridSel(Index).Row
        Case 3
            I = UBound(GridRow3) + 1
            ReDim Preserve GridRow3(I)
            GridRow3(I) = GridSel(Index).Row
        Case 4
            I = UBound(GridRow4) + 1
            ReDim Preserve GridRow4(I)
            GridRow4(I) = GridSel(Index).Row
    End Select
End If
End Sub

Private Sub GridSel_KeyPress(Index As Integer, KeyAscii As Integer)
If GridSel(Index).Col = 0 Or GridSel(Index).Row = 0 Then Exit Sub
Select Case Index
    Case 1
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid1, KeyAscii, RsGrid1.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 2
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid2, KeyAscii, RsGrid2.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 3
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid3, KeyAscii, RsGrid3.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 4
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid4, KeyAscii, RsGrid4.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
End Select
TxtSearch.Tag = Index
End Sub



Private Sub TopCtrl1_eCancel()
TopCtrl1.TopText2 = "Browse": TopCtrl1.TopText2.ForeColor = RGB(0, 0, 255)
Disp_Text False
End Sub

Private Sub TopCtrl1_eEdit()
TopCtrl1.TopText2 = "Edit": TopCtrl1.TopText2.ForeColor = RGB(255, 0, 0)
Disp_Text True
End Sub

Private Sub TopCtrl1_eExit()
BTNEXIT.SetFocus
FrmDetail.Visible = False
FGrid1.Visible = False

TopCtrl1.TopText2 = "Edit": TopCtrl1.TopText2.ForeColor = RGB(255, 0, 0)
TopCtrl1.Visible = False
'Disp_Text True  'vijay jain
End Sub

Private Sub TopCtrl1_ePrn()
On Error GoTo ERRORHANDLER
RepPrint = True
SPADE_ReportProc
If RepPrint = False Then Exit Sub
       
CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True
Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")
rpt.Database.SetDataSource RstRep
rpt.ReadRecords

Call Formulas
Call Report_View(rpt, RepTitle, , False)
Set RstRep = Nothing
Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION

End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer
Dim mTrans As Boolean
Dim mQry As String
On Error GoTo errlbl
    If txtgrid1(0).Visible = True Then
        If TxtGrid1Leave = False Then
            txtgrid1(0).SetFocus
            Exit Sub
        End If
    End If
    I = 1
    GCn.BeginTrans
    mTrans = True
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, 0) = "¤" Then
            GCn.Execute " update Veh_Quot set MAP='" & FGrid1.TextMatrix(I, Col_Map) & _
                "',FOLLOW_UP='" & (FGrid1.TextMatrix(I, Col_Follow)) & _
                "',DEL_DATE=" & ConvertDate(FGrid1.TextMatrix(I, Col_DeliDate)) & _
                ", U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E' " & _
                " where Veh_Quot.DocID= '" & FGrid1.TextMatrix(I, Col_DocID) & "'"
        End If
    Next
GCn.CommitTrans
mTrans = False

'RsDate.Requery
'Fill_Grid 100, "General"
TopCtrl1.TopText2 = "Browse": TopCtrl1.TopText2.ForeColor = RGB(0, 0, 255)
Disp_Text False
Exit Sub
errlbl:
If mTrans = True Then
    GCn.RollbackTrans: CheckError
Else
    CheckError
End If

End Sub

Private Sub TxtSearch_Click()
TxtSearch.TEXT = "": GridSel(Val(TxtSearch.Tag)).SetFocus: TxtSearch.Visible = False
End Sub

Private Sub TxtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If NavigationKey(KeyCode) = True Then GridSel(Val(TxtSearch.Tag)).SetFocus: TxtSearch.Visible = False
If KeyCode = vbKeyDelete Then TxtSearch.TEXT = ""
If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then GridSel(Val(TxtSearch.Tag)).SetFocus: TxtSearch.Visible = False
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
Select Case TxtSearch.Tag
    Case 1
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid1, KeyAscii, RsGrid1.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 2
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid2, KeyAscii, RsGrid2.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 3
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid3, KeyAscii, RsGrid3.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 4
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid4, KeyAscii, RsGrid4.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
End Select
End Sub

Private Sub TxtSearch_LostFocus()
    TxtSearch.TEXT = "": GridSel(Val(TxtSearch.Tag)).SetFocus: TxtSearch.Visible = False
End Sub

Private Sub GridSel_LeaveCell(Index As Integer)
GridSel(Index).CellBackColor = CellBackColLeave1
End Sub

Private Sub GridSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If GridSel(Index).Col <> 0 Then Exit Sub
mGridStartRow = GridSel(Index).Row
End Sub

Private Sub GridSel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer
Dim j As Integer
If GridSel(Index).Col <> 0 Or mGridStartRow = 0 Then Exit Sub
mGridEndRow = GridSel(Index).RowSel
For j = mGridStartRow To mGridEndRow
    GridSel(Index).Row = j
    GridSel(Index).Col = 0
    GridSel(Index).CellFontName = "WINGDINGS"
    GridSel(Index).CellFontSize = 14
    GridSel(Index).TextMatrix(j, 0) = IIf(GridSel(Index).TextMatrix(j, 0) = "ü", " ", "ü")
    Select Case Index
        Case 1
            I = UBound(GridRow1) + 1
            ReDim Preserve GridRow1(I)
            GridRow1(I) = GridSel(Index).Row
        Case 2
            I = UBound(GridRow2) + 1
            ReDim Preserve GridRow2(I)
            GridRow2(I) = GridSel(Index).Row
        Case 3
            I = UBound(GridRow3) + 1
            ReDim Preserve GridRow3(I)
            GridRow3(I) = GridSel(Index).Row
        Case 4
            I = UBound(GridRow4) + 1
            ReDim Preserve GridRow4(I)
            GridRow4(I) = GridSel(Index).Row
    End Select
Next
mGridStartRow = 0
End Sub

Private Sub GridSel_Validate(Index As Integer, Cancel As Boolean)
GridSel(Index).CellBackColor = CellBackColLeave1
End Sub

Private Sub ListView_Click()
    TxtGrid(0).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    TxtGrid(0).SetFocus
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
   Select Case FGrid.Row
    Case List1
        Select Case GRepFormName

            Case SPADE_Report
              ListArray = Array("Status", "Reffered By", "City", "Financier Type", "Area", "All")
              Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 6)

          End Select
    Case List2
        Select Case GRepFormName
            Case SPADE_Report
                If FGrid.TextMatrix(List1, 1) = "Status" Then
                    ListArray = Array("Cold", "Warm", "Hot")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
                Else
                    FGrid.Enabled = False
                End If
        End Select
    Case List3
        Select Case GRepFormName
'           Case VehSalereg
'               ListArray = Array("PartyWise", "CityWise", "FinancierGrp", "FinancierName", "FormType", "All")
'               Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 6)
        End Select
    Case Cat2
        Select Case GRepFormName
'           Case ModFWiseMicro
'               FGrid.TextMatrix(Cat2, 1) = Val(FGrid.TextMatrix(Cat1, 1)) + 1
'               TxtGridLeave
        End Select
End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Integer
If KeyCode = vbKeyEscape Then
    TxtGrid(0).TEXT = TxtGrid(0).Tag
    TxtGrid_KeyUp Index, KeyCode, Shift
    FGrid.SetFocus
    TxtGrid(0).Visible = False
    Grid_Hide
    Exit Sub
End If
Select Case FGrid.Row
    Case List1, List2, List3
        ListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), TxtGrid(0).width
        If KeyCode = vbKeyReturn Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
    Case Date1, Date2, Cat1, Cat2, Cat3, Cat4, Cat5
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGridLeave = True Then TxtKeyDown
        End If
End Select
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Dim KeyCode As Integer
 Call CheckQuote(KeyAscii)
 Select Case FGrid.Row
 Case Cat1
        Select Case GRepFormName
'            Case ModFWiseMicro
'                NumPress TxtGrid(Index), KeyAscii, 4, 0
        End Select
   
        'KeyCode = 0
'           TxtGrid(0).Enabled = False

'    Case Cat3
'        Select Case GRepFormName
'            Case SprStkAgeing
'                NumPress TxtGrid(Index), KeyAscii, 3, 0
'        End Select
'    Case Cat4
'        Select Case GRepFormName
'            Case SprStkAgeing
'                NumPress TxtGrid(Index), KeyAscii, 3, 0
'        End Select
'    Case Cat5
'        Select Case GRepFormName
'            Case SprStkAgeing
'                NumPress TxtGrid(Index), KeyAscii, 3, 0
'        End Select
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
    Select Case FGrid.Row
'        Case Cat1, Cat2
'             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = IIf(TxtGrid(0) = "", "", Format(TxtGrid(0), "0.00"))
'
'        Case Cat2
'            'If Val(FGrid.TextMatrix(Cat2, 1))  > Val(FGrid.TextMatrix(Cat2, 1)) Then
            
        Case List1, List2, List3
            If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
            ListView_KeyUp ListView, TxtGrid, 0, KeyCode, mListItem
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
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
Dim KeyCode As Integer
Select Case FGrid.Row
    Case Cat3, Cat1, Cat2, Cat4, Cat5
         FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
    Case List2, List3
        If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
    Case List1
        If TxtGrid(0).TEXT <> "" Then
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
            If TxtGrid(0).TEXT <> "Status" Then FGrid.TextMatrix(List2, 1) = ""
            Select Case TxtGrid(0).TEXT
                Case "Status"
                    GridSel(3).Visible = False
                    Check1(3).Visible = False
                Case "Reffered By"
                    Grid3Sql = "select '' as O,RefName as Refference_Name,RefCode  as code from Reffered order by RefName"
                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
                    mHelpGridNo = 3
                Case "City"
                    Grid3Sql = "select '' as O,CityName as City_Name,CityCode  as code from City order by CityName"
                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
                    mHelpGridNo = 3
                Case "Financier Type"
                    Grid3Sql = "select '' as O,FinGrpName as FinGrp_Name,FinGrpCode  as code from FinGroup order by FinGrpName"
                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
''''            Case "FinancierName"
''''               Grid3Sql = "select '' as O,FinName as Financer_Name,FinCode  as code from ContractFinance order by FinName"
''''               GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
                    mHelpGridNo = 3
                Case "Area"
                    Grid3Sql = "select '' as O,AreaName as Area_Name,AreaCode  as code from Area order by AreaName"
                    GridSel(3).Visible = True: Check1(3).Visible = True: GridInitialise 3, Grid3Sql
                    mHelpGridNo = 3
                Case "All"
                    GridSel(3).Visible = False: Check1(3).Visible = False
            End Select
        End If
    Case Date1, Date2
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
End Select
    TxtGridLeave = True
    If ValidateCall = False Then
        FGrid.SetFocus
        TxtGrid(0).Visible = False
    End If
End Function

'******* Fuctions **********

Private Sub Global_Grid()
Dim I As Integer, Cnt As Integer
Pic.top = Me.top - Pic.width - 10
BTNPRINT.left = (Pic.width - (BTNPRINT.width + BTNEXIT.width)) / 2: BTNPRINT.top = Pic.top + 10
BTNEXIT.left = BTNPRINT.left + BTNPRINT.width: BTNEXIT.top = Pic.top + 10
FGrid.left = (Me.width - FGrid.width) / 2: FGrid.top = 75

FGrid.Rows = 10  '5
FGrid.Cols = 3
FGrid.FixedCols = 1
FGrid.ColWidth(0) = 2200
FGrid.ColWidth(1) = 2000
FGrid.ColWidth(2) = 0
FGrid.ColAlignment(1) = flexAlignLeftCenter
For I = 0 To FGrid.Rows - 1
    FGrid.RowHeight(I) = 0
Next
Ini_Grid
For I = 1 To 4
    If GridSel(I).Visible = True Then Cnt = Cnt + 1
Next
'FGrid.Height = (((mLastRow - mFirstRow) + 1) * PubGridRowHeight) + 500
FGrid.height = (((mLastRow + 1) - mFirstRow) * PubGridRowHeight) + 500
Select Case mHelpGridNo
Case 0
    FGrid.top = 1000
Case 1
    GridSel(1).left = (Me.width - GridSel(1).width) / 2
    GridSel(1).top = FGrid.top + FGrid.height + 500
    GridSel(1).height = Me.height - FGrid.height - Pic.height - 1200
    Check1(1).top = GridSel(1).top + 20: Check1(1).left = GridSel(1).left + 40
Case 2
    GridSel(1).left = (Me.width / 2 - GridSel(1).width) / 2
    GridSel(1).top = FGrid.top + FGrid.height + 500
    GridSel(1).height = Me.height - FGrid.height - Pic.height - 1200
    Check1(1).top = GridSel(1).top + 20: Check1(1).left = GridSel(1).left + 40
    
    GridSel(2).left = Me.width / 2 + (Me.width / 2 - GridSel(1).width) / 2
    GridSel(2).top = FGrid.top + FGrid.height + 500
    GridSel(2).height = Me.height - FGrid.height - Pic.height - 1200
    Check1(2).top = GridSel(2).top + 20: Check1(2).left = GridSel(2).left + 40
    
Case 3
    GridSel(1).left = (Me.width / 2 - GridSel(1).width) / 2
    GridSel(1).top = FGrid.top + FGrid.height + 500
    Check1(1).top = GridSel(1).top + 20: Check1(1).left = GridSel(1).left + 40
   
    GridSel(3).left = GridSel(1).left
    GridSel(3).top = GridSel(1).top + GridSel(1).height + 500
    Check1(3).top = GridSel(3).top + 20: Check1(3).left = GridSel(3).left + 40
    
    GridSel(2).left = Me.width / 2 + (Me.width / 2 - GridSel(1).width) / 2
    GridSel(2).top = FGrid.top + FGrid.height + 500
    GridSel(2).height = GridSel(1).height + GridSel(2).height + 500
    Check1(2).top = GridSel(2).top + 20: Check1(2).left = GridSel(2).left + 40
    
Case 4
    GridSel(1).left = (Me.width / 2 - GridSel(1).width) / 2
    GridSel(1).top = FGrid.top + FGrid.height + 500
    Check1(1).top = GridSel(1).top + 20: Check1(1).left = GridSel(1).left + 40
    
    GridSel(2).left = Me.width / 2 + (Me.width / 2 - GridSel(1).width) / 2
    GridSel(2).top = FGrid.top + FGrid.height + 500
    Check1(2).top = GridSel(2).top + 20: Check1(2).left = GridSel(2).left + 40
    
    GridSel(3).left = GridSel(1).left
    GridSel(3).top = GridSel(1).top + GridSel(1).height + 500
    Check1(3).top = GridSel(3).top + 20: Check1(3).left = GridSel(3).left + 40
    
    GridSel(4).left = Me.width / 2 + (Me.width / 2 - GridSel(1).width) / 2
    GridSel(4).top = GridSel(1).top + GridSel(1).height + 500
    Check1(4).top = GridSel(4).top + 20: Check1(4).left = GridSel(4).left + 40

End Select
End Sub
Private Sub Grid_Hide()
If FrmList.Visible = True Then FrmList.Visible = False
End Sub
Private Sub FGrid_DblClick()
FGrid_KeyPress vbKeyReturn
'    Select Case FGrid.Row
'        Case Date1, Date2, List1, List2, List3, Cat1, Cat2, Cat3, Cat4, Cat5
'            Call GridDblClick(Me, FGrid, TxtGrid, 0)
'    End Select
TAddMode = False
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
Dim I As Integer
    Select Case FGrid.Row
        Case Cat1, Cat2, Cat3, Cat4, Cat5
            Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
        Case Date1, Date2, List1, List3
            Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
         Case List2
            If FGrid.TextMatrix(List1, 1) = "Status" Then
                Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
            End If
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell-- > Enter Cell-- >KeyDown
If KeyCode = vbKeyUp And Val(FGrid.Tag) = mFirstRow Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeys "+{Tab}"
    KeyCode = 0
ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = mLastRow Then
    FGrid.CellBackColor = CellBackColLeave
    SendKeysA vbKeyTab, True
    KeyCode = 0
End If
GridKey = KeyCode
FGrid.Tag = FGrid.Row
If KeyCode = vbKeyDelete And Shift = 0 Then
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
    FGrid.TextMatrix(FGrid.Row, 2) = ""
End If

'If KeyCode = vbKeyReturn Then
'    Select Case FGrid.Row
'        Case Date1, Date2, List1, List2, List3, Cat1, Cat2, Cat3, Cat4, Cat5
'            Call GridDblClick(Me, FGrid, TxtGrid, 0)
'            TAddMode = False
'    End Select
'End If
KeyCode = 0
End Sub

Private Sub FGrid_EnterCell()
FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
   FGrid.CellBackColor = CellBackColEnter
   Grid_Hide
   TxtGrid(0).Visible = False
End Sub
Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_Scroll()
TxtGrid(0).Visible = False
Grid_Hide
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
    
End Sub

Private Function FillString(GridArray As Variant, Gridindex As Integer, DataType As Byte) As String
On Error GoTo ELoop
Dim ac_str As String
Dim I As Integer
Dim GridRow As Integer
    ac_str = ""
    For I = 0 To UBound(GridArray)
        If GridArray(I) = 0 Then GoTo NXT:
        GridRow = GridArray(I)
        If GridSel(Gridindex).TextMatrix(GridRow, 0) = "ü" Then
                If DataType = 0 Then
                   ac_str = ac_str + IIf(ac_str = "", GridSel(Gridindex).TextMatrix(GridRow, 2), "," + GridSel(Gridindex).TextMatrix(GridRow, 2))
                ElseIf DataType = 1 Then
                   ac_str = ac_str + IIf(ac_str = "", "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'", "," + "'" + GridSel(Gridindex).TextMatrix(GridRow, 2) + "'")
                End If
            GridSel(Gridindex).TextMatrix(GridRow, 0) = ""
        Else
            GridArray(I) = 0
        End If
NXT:
    Next
    For I = 0 To UBound(GridArray)
        GridRow = GridArray(I)
        If GridArray(I) <> 0 Then
            GridSel(Gridindex).TextMatrix(GridRow, 0) = "ü"
        End If
    Next
'    Erase GridArray
'    ReDim Preserve GridArray(0)
'    GridArray(0) = 0
    If ac_str = "" Then
        MsgBox "Select " & GridSel(Gridindex).TextMatrix(0, 1), vbInformation
        GridSel(Gridindex).SetFocus
        RepPrint = False
        Exit Function
    End If
    FillString = ac_str
    Exit Function
ELoop:
    RepPrint = False
    MsgBox err.Description
End Function

Private Sub TxtKeyDown()
Dim I As Integer
    If FGrid.Row = mLastRow Then SendKeysA vbKeyTab, True: Exit Sub
    For I = FGrid.Row To FGrid.Rows - 1
         If FGrid.RowHeight(I + 1) <> 0 Then FGrid.Row = I + 1: Exit For
    Next
End Sub
Private Sub GridInitialise(Gridindex As Integer, GridSql As String)
Dim Index As Integer
Index = Gridindex
If Index = 1 Then
    Set RsGrid1 = New ADODB.Recordset: RsGrid1.CursorLocation = adUseClient
    RsGrid1.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid1
'    GridSel(Index).top = G1Top: GridSel(Index).left = G1left
    ReDim Preserve GridRow1(0)
    GridRow1(0) = 0
End If
If Index = 2 Then
    Set RsGrid2 = New ADODB.Recordset: RsGrid2.CursorLocation = adUseClient
    RsGrid2.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid2
'    GridSel(Index).top = G2Top: GridSel(Index).left = G2left
    ReDim Preserve GridRow2(0)
    GridRow2(0) = 0
End If
If Index = 3 Then
    Set RsGrid3 = New ADODB.Recordset: RsGrid3.CursorLocation = adUseClient
    RsGrid3.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid3
'    GridSel(Index).top = G3Top: GridSel(Index).left = G3left
        ReDim Preserve GridRow3(0)
        GridRow3(0) = 0
End If
If Index = 4 Then
    Set RsGrid4 = New ADODB.Recordset: RsGrid4.CursorLocation = adUseClient
    RsGrid4.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid4
'    GridSel(Index).top = G4Top: GridSel(Index).left = G4left
    ReDim Preserve GridRow4(0)
    GridRow4(0) = 0
End If
GridSel(Index).height = 1700
GridSel(Index).Visible = True: GridSel(Index).Enabled = False: Check1(Index).Visible = True
GridSel(Index).width = 5200: GridSel(Index).ColWidth(0) = 600: GridSel(Index).ColWidth(2) = 0: GridSel(Index).ColWidth(1) = 4000
'Check1(Index).top = GridSel(Index).top + 20: Check1(Index).left = GridSel(Index).left + 40
Check1(Index).width = 580: Check1(Index).height = GridSel(Index).RowHeight(0) + 20: Check1(Index).Value = Checked
End Sub

Private Sub Ini_Grid()
'Date1 , Date2, List1, List1, List2, List3
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where  site_code ='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If

Select Case GRepFormName
    Case SPADE_Report   'vijay Vehicle
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Based On": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List2, 0) = "Status": .RowHeight(List2) = GridRowHeight
            
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "Status"
             
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 3
        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
        GridInitialise 1, Grid1Sql
        Grid2Sql = "select '' as O,Model.Model_Desc As Model_Description,Model.Model As Code from Model order by Model.Model_Desc"
        GridInitialise 2, Grid2Sql
End Select
End Sub
Public Function IsNotBlank(FieldRow As Integer, FieldCaption As String) As Boolean
    If FGrid.TextMatrix(FieldRow, 1) = "" Then
        MsgBox FieldCaption & " Should not be Blank.", vbInformation, "Validation Check"
        FGrid.SetFocus
        FGrid.Row = FieldRow
        FGrid.Col = 1
        IsNotBlank = False
    Else
        IsNotBlank = True
    End If
End Function

Private Sub Formulas()
On Error GoTo ELoop
Dim I As Integer
Select Case GRepFormName


Case SPADE_Report
    For I = 1 To rpt.FormulaFields.Count
    Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
        Case UCase("DATEBETWEEN")
            rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"

    End Select
    Next
End Select
Exit Sub
ELoop:
     MsgBox err.Description
End Sub

Private Sub SPADE_ReportProc()
On Error GoTo ELoop
Dim mQry As String

  '*******Note: Condstr has been created in Fill_Grid and it's being defined at option explicit
  'vijay jain(14/12/02)******************
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (GSQL), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    RepName = "SpadeRep"
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub

Public Sub SelGridKeyPressLocal(txt As Object, SelGrid As Object, Index As Integer, Rst As ADODB.Recordset, ByRef KeyAscii As Integer, FindFldName As String, Optional CellBackColEnter As ColorConstants, Optional CellBackColLeave As ColorConstants)
Dim FindStr$    ' As String
Dim LPlace As Byte
'    If FilterKeyCode(KeyAscii) = True Then Exit Sub
    If SelGrid(Index).Rows < 1 Then Exit Sub
    If Rst.RecordCount <= 0 Then txt.TEXT = "": Exit Sub
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyDelete Then Exit Sub
        If KeyAscii = vbKeyBack Then
            If Len(txt.SelText) > 1 Then
                txt.SelLength = Len(txt.SelText) - 1
                FindStr = txt.SelText
            Else
                txt.TEXT = ""
                SelGrid(Index).SetFocus
                txt.Visible = False
                Exit Sub
            End If
        Else
            FindStr = txt.SelText + Chr(KeyAscii)
        End If
        Rst.MoveFirst
        If Rst.Fields(FindFldName).Type = adInteger Then    'Numeric Search
            Rst.FIND "" & FindFldName & "  >=" & Val(FindStr) & ""
        Else    'character serach
            Rst.FIND "" & FindFldName & " like '" & FindStr & "*'"
        End If
        KeyAscii = 0
       If Rst.AbsolutePosition <> adPosEOF And Rst.AbsolutePosition <> adPosBOF Then
            SelGrid(Index).CellBackColor = CellBackColLeave
            SelGrid(Index).Row = Rst.AbsolutePosition
            SelGrid(Index).CellBackColor = CellBackColEnter
            txt.TEXT = Rst.Fields(FindFldName).Value
            txt.SelLength = Len(FindStr)
            txt.left = SelGrid(Index).CellLeft + SelGrid(Index).left
            txt.top = SelGrid(Index).CellTop + SelGrid(Index).top
            If txt.Visible = False Then
                txt.Visible = True: txt.ZOrder 0: txt.SetFocus: txt.BackColor = SelGrid(Index).CellBackColor
                 txt.ForeColor = SelGrid(Index).CellForeColor: txt.width = SelGrid(Index).CellWidth: txt.height = SelGrid(Index).CellHeight
            End If
       End If
End Sub
                                                               
Private Function GetYear() As Integer
Dim mYear%
mYear = Val(TxtGrid(0).TEXT)
        If mYear = 0 Then mYear = Year(date)
        If mYear > 1999 Then mYear = Right(STR(mYear), 2)
        mYear = Val(mID(CStr(Year(date)), 1, 4 - Len(Trim(CStr(mYear)))) + Trim(CStr(mYear)))
GetYear = mYear
End Function
Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
'For i = 0 To txt.count - 1
'    txt(i).Enabled = Enb
'Next
TopCtrl1.tEdit = Not Enb
TopCtrl1.tExit = Not Enb
TopCtrl1.tPrn = Not Enb
TopCtrl1.tCancel = Enb
TopCtrl1.tSave = Enb

TopCtrl1.tRef = False
TopCtrl1.tAdd = False
TopCtrl1.tFirst = False
TopCtrl1.tNext = False
TopCtrl1.tPrev = False
TopCtrl1.tLast = False
TopCtrl1.tFind = False
TopCtrl1.tDel = False

txtgrid1(0).Visible = False
txtgrid1(0).Visible = False
TxtSearch1.Visible = False

 
End Sub
Private Sub TxtGrid1_GotFocus(Index As Integer)
    txtgrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    FGrid1.SetFocus
    txtgrid1(0).TEXT = txtgrid1(0).Tag
    txtgrid1(0).Visible = False
    Exit Sub
End If
Select Case FGrid1.Col
    Case Col_Map, Col_DeliDate, Col_Follow
        If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
            If TxtGrid1Leave = True Then
                 GridTxtDown FGrid1, txtgrid1, Index, KeyCode, TAddMode, 15, , , True, True
            End If
        End If
End Select
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
 Call CheckQuote(KeyAscii)
'Select Case FGrid.Col
'    Case Col_MRP, Col_Taxable, Col_TaxPaid
'        Call NumPress(TxtGrid(Index), KeyAscii, 8, 2)
'End Select
End Sub

Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGrid1Leave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid1Leave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid1.Col
    Case Col_Map, Col_Follow
       If FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) <> (txtgrid1(0).TEXT) Then
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = txtgrid1(0).TEXT
            FGrid1.CellForeColor = vbRed
            FGrid1.TextMatrix(FGrid1.Row, 0) = "¤"
        End If
    Case Col_DeliDate
    txtgrid1(0).MaxLength = 11
   
        If Format(FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col), "DD/MMM/YYYY") <> Format((txtgrid1(0).TEXT), "DD/MMM/YYYY") Then
            FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = RetDate(txtgrid1(0))
            FGrid1.CellForeColor = vbRed
            FGrid1.TextMatrix(FGrid1.Row, 0) = "¤"
        End If

End Select

TxtGrid1Leave = True
If ValidateCall = False Then
    FGrid1.SetFocus
    txtgrid1(0).Visible = False
End If
End Function

Private Sub Ini_Grid1()
'DGDate.left = txt(Vdate).left: DGDate.top = txt(Vdate).top + txt(Vdate).Height + 15
Dim I As Byte
    FrmDetail.top = TopCtrl1.height
    FrmDetail.left = Me.left
    FrmDetail.width = Me.width - 120
    With FGrid1
        .top = TopCtrl1.height + FrmDetail.height
        .left = Me.left
        .width = Me.width - 120
        .height = Me.height - (TopCtrl1.height + FrmDetail.height + mBotScale) ' 7635
        .RowHeightMin = PubGridRowHeight
        .Cols = 25
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColWidth(0) = 300

        .TextMatrix(0, Col_PDtNo) = "Quotation No. & Dt."
        .ColAlignment(Col_PDtNo) = flexAlignLeftCenter           'Performa Date and Inv No
        .ColAlignmentFixed(Col_PDtNo) = flexAlignLeftCenter
        .ColWidth(Col_PDtNo) = 2000

        .TextMatrix(0, Col_RefBy) = "Ref By"                    'Reffered By
        .ColAlignment(Col_RefBy) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_RefBy) = flexAlignLeftCenter
        .ColWidth(Col_RefBy) = 570
        
        .TextMatrix(0, Col_City) = "City"               'City Name
        .ColAlignment(Col_City) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_City) = flexAlignLeftCenter
        .ColWidth(Col_City) = 1300
        
        .TextMatrix(0, Col_Phone) = "Phone"             ' Phone No of Customer
        .ColAlignment(Col_Phone) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_Phone) = flexAlignLeftCenter
        .ColWidth(Col_Phone) = 0
        
        .TextMatrix(0, Col_Veh1st) = "Veh 1st"              'Vehicle !st Yes/No
        .ColAlignment(Col_Veh1st) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_Veh1st) = flexAlignLeftCenter
        .ColWidth(Col_Veh1st) = 600

        .TextMatrix(0, Col_Bank) = "Bank If Any"            'Bank If Any
        .ColAlignment(Col_Bank) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_Bank) = flexAlignLeftCenter
        .ColWidth(Col_Bank) = 2500

        .TextMatrix(0, Col_Govt) = "Govt"                       'Govt Yes/No
        .ColAlignment(Col_Govt) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_Govt) = flexAlignLeftCenter
        .ColWidth(Col_Govt) = 435

        .TextMatrix(0, Col_Status) = "Status"           'call StAtus Cold/Warm/Hot
        .ColAlignment(Col_Status) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_Status) = flexAlignLeftCenter
        .ColWidth(Col_Status) = 600

        .TextMatrix(0, Col_Model) = "Model"
        .ColAlignment(Col_Model) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_Model) = flexAlignLeftCenter
        .ColWidth(Col_Model) = 1815 ' 3000
        
        .TextMatrix(0, Col_VehQty) = "Qty"
        .ColAlignment(Col_VehQty) = flexAlignCenterCenter
        .ColAlignmentFixed(Col_VehQty) = flexAlignCenterCenter
        .ColWidth(Col_VehQty) = 315

        .TextMatrix(0, Col_Map) = "MAP"
        .ColAlignment(Col_Map) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_Map) = flexAlignLeftCenter
        .ColWidth(Col_Map) = 1680
        
        .TextMatrix(0, Col_DeliDate) = "Del. Date"
        .ColAlignment(Col_DeliDate) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_DeliDate) = flexAlignLeftCenter
        .ColWidth(Col_DeliDate) = 1335

        .TextMatrix(0, Col_Follow) = "Follow Up Results"
        .ColAlignment(Col_Follow) = flexAlignLeftCenter
        .ColAlignmentFixed(Col_Follow) = flexAlignLeftCenter
        .ColWidth(Col_Follow) = 2000
        
        .ColWidth(Col_Area) = 0
        .ColWidth(Col_VNo) = 0
        .ColWidth(Col_VDate) = 0
        .ColWidth(Col_Name) = 0
        .ColWidth(Col_Add) = 0
        .ColWidth(Col_Add2) = 0
        .ColWidth(Col_Add3) = 0
        .ColWidth(Col_ModelDesc) = 0
        .ColWidth(Col_DocID) = 0
        .ColWidth(Col_Phone) = 0
        .ColWidth(Col_AreaName) = 0
        .ColWidth(Col_RepName) = 0
    End With
    BackColorSelLeave = FGrid1.BackColorSel
    ForeColorSelEnter = FGrid1.ForeColorSel
End Sub

Private Sub Fill_Grid(Index As Integer, GridCaption As String)
Dim Rst As ADODB.Recordset
Dim MRPPer As Double, TBPer As Double, TPPer As Double
Dim mQry As String
Set Master = New Recordset
If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then: Exit Sub
If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then: Exit Sub
If IsNotBlank(List1, FGrid.TextMatrix(List1, 1)) = False Then: Exit Sub
If FGrid.TextMatrix(List1, 1) = "Status" Then
   If IsNotBlank(List2, FGrid.TextMatrix(List2, 1)) = False Then: Exit Sub
End If
Condstr = ""
CondStr1 = ""
If Check1(3).Value = Checked Then
   Select Case FGrid.TextMatrix(List1, 1)
        Case "Reffered By"
            CondStr1 = " Order By VQ.REF_CODE"
        Case "City"
            CondStr1 = " Order By City.CityName"
        Case "Financier Type"
            CondStr1 = " Order By CF.FinName"
        Case "Area"
            CondStr1 = " Order By VQ.Area"
        Case "Status"
            CondStr1 = " Order By VQ.Call_Status"
   End Select
End If

    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1)
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1)

    Condstr = " where VQ.V_Date  >= " & ConvertDate(FGrid.TextMatrix(Date1, 1)) & "  AND VQ.V_Date <= " & ConvertDate(FGrid.TextMatrix(Date2, 1)) & "  "

    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and left(VQ.Site_Code,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and left(VQ.Site_Code,1) ='" & PubSiteCode & "' "
    End If


    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and VQ1.Model in (" & GridString2 & ")"
   
    Select Case FGrid.TextMatrix(List1, 1)
        Case "Reffered By"
             If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
             If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VQ.REF_CODE in (" & GridString3 & ") Order By VQ.REF_CODE"
        Case "City"
             If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
             If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VQ.CityCode in (" & GridString3 & ") order By City.CityName"
        Case "Financier Type"
             If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
             If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VQ.FB_CODE in (Select CF.FinCode From ContractFinance where CF.UnderFinGrp in (" & GridString3 & ")) Order By CF.FinName"
        Case "Area"
             If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
             If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VQ.AREA in (" & GridString3 & ") Order By VQ.Area"
        Case "Status"
           Select Case FGrid.TextMatrix(List2, 1)
               Case "Cold"
                 Condstr = Condstr & " and VQ.Call_Status=0"
               Case "Warm"
                 Condstr = Condstr & " and VQ.Call_Status=1"
               Case "Hot"
               Condstr = Condstr & " and VQ.Call_Status=2"
          End Select
   End Select
   If PubBackEnd = "A" Then
        GSQL = "SELECT (" & cTrim(cMID("VQ.DocID", "9", "5")) & " + " & cTrim(cMID("VQ.DocID", "14", "8")) & " + ' '+ " & cCStr("VQ.V_Date") & ")as PerNoDt, VQ.REF_CODE, PC.Name,City.CityName, " & _
            " PC.PhoneOff, " & cIIF("VQ.FirstVeh_YN=0", "'No'", "'Yes'") & " AS 1stVeh," & _
            " CF.FinName As Bank_If_Any, Switch(VQ.GOVT_YN=0,'No',VQ.GOVT_YN=1,'Yes' ) AS Govt," & _
            " Switch(VQ.Call_Status=0,'Cold',VQ.Call_Status=1,'Warm' ,VQ.Call_Status=2,'Hot') AS Status," & _
            " VQ1.Model, VQ1.QTY, VQ.MAP,VQ.DEL_DATE,VQ.FOLLOW_UP,VQ.Area,VQ.V_No,VQ.V_Date, " & _
            " PC.Add1,PC.Add2,PC.Add3,Model.Model_Desc,VQ.DocID,Area.AreaName,Emp_Mast.Emp_Name,Reffered.RefName " & _
            " FROM (((((((Veh_Quot VQ Left Join ProspectiveCust PC ON VQ.Party_Code = PC.Cust_Code) " & _
            " LEFT JOIN City ON VQ.CityCode = City.CityCode) " & _
            " LEFT JOIN Veh_Quot1 VQ1 ON VQ.DocId = VQ1.DocId)" & _
            " LEFT JOIN Model On Model.Model=VQ1.Model)" & _
            " LEFT JOIN ContractFinance CF On CF.FinCode=VQ.FB_CODE) " & _
            " Left Join Area on VQ.AREA = Area.AreaCode) " & _
            " Left Join Reffered on VQ.REF_CODE=Reffered.RefCode) " & _
            " Left Join Emp_Mast on VQ.REP_CODE=Emp_Mast.Emp_Code"
    ElseIf PubBackEnd = "S" Then
        GSQL = "SELECT (" & cTrim(cMID("VQ.DocID", "9", "5")) & " + " & cTrim(cMID("VQ.DocID", "14", "8")) & " + ' '+ " & cCStr("VQ.V_Date") & ")as PerNoDt, VQ.REF_CODE, PC.Name,City.CityName, " & _
            " PC.PhoneOff, " & cIIF("VQ.FirstVeh_YN=0", "'No'", "'Yes'") & " AS FstVeh," & _
            " CF.FinName As Bank_If_Any, " & cIIF("VQ.GOVT_YN =0", "'No'", "'Yes'") & " AS Govt," & _
            " (Case VQ.Call_Status When 0 Then 'Cold' When 1 Then 'Warm' When 2 Then 'Hot' End) AS Status," & _
            " VQ1.Model, VQ1.QTY, VQ.MAP,VQ.DEL_DATE,VQ.FOLLOW_UP,VQ.Area,VQ.V_No,VQ.V_Date, " & _
            " PC.Add1,PC.Add2,PC.Add3,Model.Model_Desc,VQ.DocID,Area.AreaName,Emp_Mast.Emp_Name,Reffered.RefName " & _
            " FROM (((((((Veh_Quot VQ Left Join ProspectiveCust PC ON VQ.Party_Code = PC.Cust_Code) " & _
            " LEFT JOIN City ON VQ.CityCode = City.CityCode) " & _
            " LEFT JOIN Veh_Quot1 VQ1 ON VQ.DocId = VQ1.DocId)" & _
            " LEFT JOIN Model On Model.Model=VQ1.Model)" & _
            " LEFT JOIN ContractFinance CF On CF.FinCode=VQ.FB_CODE) " & _
            " Left Join Area on VQ.AREA = Area.AreaCode) " & _
            " Left Join Reffered on VQ.REF_CODE=Reffered.RefCode) " & _
            " Left Join Emp_Mast on VQ.REP_CODE=Emp_Mast.Emp_Code"
    
    End If
        
 GSQL = GSQL + Condstr + CondStr1

Set Master = GCn.Execute(GSQL)
If Master.RecordCount <= 0 Then
    MsgBox "No Records Found! Sorry", vbExclamation, Me.CAPTION: Exit Sub
    FGrid1.Rows = 1
    FGrid1.AddItem ""
    FGrid1.FixedRows = 1
Else
    Set FGrid1.DataSource = Master
    FrmDetail.Visible = True
    FGrid1.Visible = True
    TopCtrl1.Visible = True
    TopCtrl1.ZOrder 0
    FGrid1.ZOrder 0
    FrmDetail.ZOrder 0
End If
Ini_Grid1
FGrid1.SetFocus
FGrid1_RowColChange
End Sub

