VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form RepFormCommon 
   BackColor       =   &H00C8E8DA&
   Caption         =   "MIS Report"
   ClientHeight    =   8220
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11880
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11880
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
      Index           =   11
      Left            =   7620
      TabIndex        =   34
      Top             =   1485
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   915
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
      Index           =   10
      Left            =   7665
      TabIndex        =   33
      Top             =   1710
      Value           =   1  'Checked
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
      Index           =   9
      Left            =   8670
      TabIndex        =   32
      Top             =   1470
      Value           =   1  'Checked
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
      Index           =   8
      Left            =   8670
      TabIndex        =   31
      Top             =   1725
      Value           =   1  'Checked
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
      Index           =   7
      Left            =   9675
      TabIndex        =   30
      Top             =   1455
      Value           =   1  'Checked
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
      Index           =   12
      Left            =   9675
      TabIndex        =   29
      Top             =   1710
      Value           =   1  'Checked
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
      Index           =   6
      Left            =   9645
      TabIndex        =   22
      Top             =   1245
      Value           =   1  'Checked
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
      Index           =   5
      Left            =   9645
      TabIndex        =   20
      Top             =   990
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton CmdMin 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1065
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton BTNPRINT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Print"
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
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Print Report"
      Top             =   570
      Width           =   1620
   End
   Begin VB.CommandButton BTNEXIT 
      BackColor       =   &H00C0FFFF&
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
      Height          =   375
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exit Form"
      Top             =   570
      Width           =   1620
   End
   Begin VB.CommandButton BtnSpeed 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Speed Print"
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
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Print Report"
      Top             =   -105
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   -2265
      TabIndex        =   13
      Top             =   735
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   225
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   105
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3942
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
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
      Height          =   360
      Index           =   0
      Left            =   870
      TabIndex        =   12
      Top             =   285
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   240
      HideSelection   =   0   'False
      Left            =   1590
      TabIndex        =   11
      Top             =   5700
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
      Left            =   8640
      TabIndex        =   7
      Top             =   1260
      Value           =   1  'Checked
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
      Left            =   8640
      TabIndex        =   5
      Top             =   1005
      Value           =   1  'Checked
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
      Left            =   7650
      TabIndex        =   3
      Top             =   1245
      Value           =   1  'Checked
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
      Left            =   7590
      TabIndex        =   1
      Top             =   1020
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1185
      Index           =   1
      Left            =   75
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   2090
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
      Height          =   1185
      Index           =   2
      Left            =   4050
      TabIndex        =   4
      Top             =   2115
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2090
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
      Height          =   1185
      Index           =   4
      Left            =   75
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2090
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
      Height          =   1185
      Index           =   3
      Left            =   7410
      TabIndex        =   6
      Top             =   2115
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2090
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
      Left            =   615
      TabIndex        =   0
      Top             =   45
      Width           =   5580
      _ExtentX        =   9843
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
         Name            =   "Arial"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1185
      Index           =   5
      Left            =   4035
      TabIndex        =   9
      Top             =   3375
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2090
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
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   0
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   661
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSel 
      Height          =   1185
      Index           =   6
      Left            =   7410
      TabIndex        =   21
      Top             =   3390
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2090
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
      Height          =   1185
      Index           =   8
      Left            =   3765
      TabIndex        =   23
      Top             =   4590
      Visible         =   0   'False
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   2090
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
      Height          =   1185
      Index           =   7
      Left            =   0
      TabIndex        =   24
      Top             =   4605
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2090
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
      Height          =   1185
      Index           =   12
      Left            =   7440
      TabIndex        =   25
      Top             =   5715
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2090
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
      Height          =   1185
      Index           =   9
      Left            =   7380
      TabIndex        =   26
      Top             =   4515
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2090
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
      Height          =   1185
      Index           =   10
      Left            =   15
      TabIndex        =   27
      Top             =   5835
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2090
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
      Height          =   1185
      Index           =   11
      Left            =   3885
      TabIndex        =   28
      Top             =   5805
      Visible         =   0   'False
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   2090
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
   Begin VB.Shape Shape1 
      Height          =   795
      Left            =   6585
      Top             =   195
      Width           =   5070
   End
   Begin VB.Label LblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "LblTitle"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   6585
      TabIndex        =   15
      Top             =   195
      Width           =   5070
   End
End
Attribute VB_Name = "RepFormCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Grid1Sql As String, Grid2Sql As String, Grid3Sql As String, Grid4Sql As String
Dim Grid5Sql As String, Grid6Sql As String, Grid7Sql As String, Grid8Sql As String
Dim Grid9Sql As String, Grid10Sql As String, Grid11Sql As String, Grid12Sql As String
Dim Condstr As String


Dim ActiveGridTop As Double, ActiveGridLeft As Double
Dim ActiveGridHeight As Double, ActiveGridWidth As Double

Private Const CellBackColLeave As String = &HFFFFFF
Private Const CellBackColEnter As String = &HFFFFC0
Private Const CellBackColLeave1 As String = &HEDF7FE
Private Const CellBackColEnter1 As String = &HFFFFC0

Dim RsGrid1 As ADODB.Recordset
Dim RsGrid2 As ADODB.Recordset
Dim RsGrid3 As ADODB.Recordset
Dim RsGrid4 As ADODB.Recordset
Dim RsGrid5 As ADODB.Recordset
Dim RsGrid6 As ADODB.Recordset
Dim RsGrid7 As ADODB.Recordset
Dim RsGrid8 As ADODB.Recordset
Dim RsGrid9 As ADODB.Recordset
Dim RsGrid10 As ADODB.Recordset
Dim RsGrid11 As ADODB.Recordset
Dim RsGrid12 As ADODB.Recordset

Dim RepTitle As String, RepName As String
Dim RepPrint As Boolean

Dim RstRep As ADODB.Recordset
Dim RstRep1 As ADODB.Recordset
Dim SubRep1 As Boolean
Dim SpeedPrn As Boolean

Dim FormulaStr1 As String, FormulaStr2 As String, FormulaStr3 As String, FormulaStr4 As String
Dim FormulaStr5 As String, FormulaStr6 As String, FormulaStr7 As String, FormulaStr8 As String
Dim FormulaStr9 As String, FormulaStr10 As String, FormulaStr11 As String, FormulaStr12 As String
Dim aa As String

Dim TransLimitValue As Double

Private Const GridRowHeight As Integer = 270



Private Const VehReposesReg As Byte = 0
Private Const VehSaleReg As Byte = 1
Private Const PurSaleTaxSumm As Byte = 2
Private Const VehSumModel As Byte = 3
Private Const BodyBuilderChassis As Byte = 4
Private Const StockAtBodyBuilder As Byte = 5
Private Const VehOfftakeNRetail As Byte = 6


Private Const Date1 As Byte = 0
Private Const Date2 As Byte = 1
Private Const Date3 As Byte = 2
Private Const Date4 As Byte = 3
Private Const List1 As Byte = 4
Private Const List2 As Byte = 5
Private Const List3 As Byte = 6
Private Const List4 As Byte = 7
Private Const List5 As Byte = 8
Private Const FromVno As Byte = 9
Private Const ToVno As Byte = 10
Private Const Cat1 As Byte = 11
Private Const Cat2 As Byte = 12
Private Const Cat3 As Byte = 13
Private Const Cat4 As Byte = 14
Private Const Cat5 As Byte = 15
Private Const Cat6 As Byte = 16


Private Const DocCatg As String = "SALE"    ' Kapil
Private DocNCat As String

Public GRepFormName As String
Dim mLastRow As Integer
Dim mFirstRow As Integer
Dim mHelpGridNo
Dim GridKey As Integer
Dim TAddMode As Boolean
Dim ListArray As Variant
Dim ListArray1 As Variant   'up
Dim GridString1 As String
Dim GridString2 As String
Dim GridString3 As String
Dim GridString4 As String
Dim GridString5 As String
Dim GridString6 As String
Dim GridString7 As String
Dim GridString8 As String
Dim GridString9 As String
Dim GridString10 As String
Dim GridString11 As String
Dim GridString12 As String

Dim GridRow1() As Integer
Dim GridRow2() As Integer
Dim GridRow3() As Integer
Dim GridRow4() As Integer
Dim GridRow5() As Integer
Dim GridRow6() As Integer
Dim GridRow7() As Integer
Dim GridRow8() As Integer
Dim GridRow9() As Integer
Dim GridRow10() As Integer
Dim GridRow11() As Integer
Dim GridRow12() As Integer

Dim mGridStartRow As Integer
Dim mGridEndRow As Integer
Dim oBAL As Double
'Private Const WksReqWrt$ = "W_RW"
Dim mListItem As ListItem

Private Sub btnexit_Click()
    Unload Me
End Sub
Private Sub BTNPRINT_Click()
On Error GoTo ERRORHANDLER
SubRep1 = False
RepPrint = True

Select Case GRepFormName
    Case VehReposesReg
        ProcVehReposesReg
    Case VehSaleReg
        VehSaleRegProc
    Case PurSaleTaxSumm
        PurSaleTaxSummProc
    Case VehSumModel
        VehSumModelProc
    Case BodyBuilderChassis
        ProcBodyBuilderChassis
    Case StockAtBodyBuilder
        ProcStockAtBodyBuilder
    Case VehOfftakeNRetail
        VehOfftakeNRetailProc
End Select




If RepPrint = False Then Exit Sub
CreateFieldDefFile RstRep, PubRepoPath & "\" & RepName & ".ttx", True
If SubRep1 = True Then CreateFieldDefFile RstRep1, PubRepoPath & "\" & RepName & "1.ttx", True

Set rpt = rdApp.OpenReport(PubRepoPath & "\" & RepName & ".RPT")
rpt.Database.SetDataSource RstRep
If SubRep1 = True Then rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstRep1
rpt.ReadRecords

Set RstRep = Nothing
Call Formulas
Call Report_View(rpt, RepTitle, 0, False)

SpeedPrn = False
Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub


Private Sub ProcVehReposesReg()
 Dim mQry As String, mQRY1 As String
 Dim Rst As ADODB.Recordset, RstSysView As New ADODB.Recordset, rec As New ADODB.Recordset
    Condstr = ""
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    
    If CheckGridSel = False Then RepPrint = False: Exit Sub
    
    Condstr = " where REPOSES.V_date>=" & ConvertDate(FGrid.TextMatrix(Date1, 1)) & " and REPOSES.V_date<=" & ConvertDate(FGrid.TextMatrix(Date2, 1)) & ""
    
    If FGrid.TextMatrix(List1, 1) = "ReposesOnly" Then
        Condstr = Condstr & " and Reposes.V_Type='Repos'"
    ElseIf FGrid.TextMatrix(List1, 1) = "ReleasedOnly" Then
        Condstr = Condstr & " and Reposes.V_Type='Relea'"
    End If
    
    Call MakeSelection
    
    mQry = " SELECT REPOSES.Case_No,RegNo,iif(REPOSES.V_Type='Repos','Reposes','Release') as V_Type," & _
           " Subgroup.Name as Hirer,Subgroup.Add1,Subgroup.Add2,Subgroup.Add3,City.CityName,City.PinCode,0 as OD_Amt,0 as Due_Amt,NetAmt as Con_Value,REPOSES.V_Date as Rpos_Dt," & _
           " REPOSES.NARRATION,REPOSES.FIR_DATE,REPOSES.INV_LIST,REPOSES.AGENCY,Reposes.SeizingCharges,Reposes.GarageAmt,Model.Name " & _
           " From ((((((((REPOSES " & _
           " Left Join CaseData On REPOSES.Site_Code&REPOSES.Case_No=CaseData.Site_Code&CaseData.V_No) " & _
           " left join Subgroup on CaseData.HirerCode=Subgroup.SubCode) " & _
           " left join Model on CaseData.ModelCode=Model.Code) " & _
           " left join City on Subgroup.CityCode=City.CityCode) " & _
           " left join AreaMast on SubGroup.AreaCode=AreaMast.AreaCode) " & _
           " left join Subgroup introducer on CaseData.IntroducerCode=introducer.SubCode) " & _
           " left join Inspect on CaseData.InspectorCode=Inspect.Code) " & _
           " left join Site on Reposes.Site_Code=Site.Site_Code) " & _
           " left join Scheme on CaseData.SchemeCode=Scheme.Code "
 
     mQry = mQry & Condstr
      
     Set RstRep = New Recordset
     RstRep.CursorLocation = adUseClient
     RstRep.Open (mQry), G_FaCn, adOpenDynamic, adLockOptimistic
     If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
        RepName = "VehReposesReg"
        RepTitle = "Vehicle Reposes Register"
        Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub


Private Sub BtnSpeed_Click()
SpeedPrn = True
BTNPRINT_Click
End Sub
Private Sub Check1_Click(Index As Integer)
'    If Check1(Index).Value = Unchecked Then
'        GridSel(Index).Enabled = True
'        If GridSel(Index).Rows > 1 Then
'            GridSel(Index).Row = 1: GridSel(Index).Col = 1
'        End If
'    Else
'        GridSel(Index).Enabled = False
'        If GridSel(Index).Rows > 1 Then
'            GridSel(Index).Row = 0: GridSel(Index).Col = 0
'            GridSel(Index).RowSel = GridSel(Index).Rows - 1
'        End If
'    End If
    
'***********************************************
'modishekhar  251103
    If Check1(Index).Value = Unchecked Then
        GridSel(Index).Enabled = True
        If GridSel(Index).Rows > 1 Then
            GridSel(Index).Row = 1: GridSel(Index).Col = 1
        End If
        CmdMin.Tag = Index
        'GridSelMax True
    Else
        'GridSelMax False
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
Private Sub CmdMin_Click()
GridSelMax False
End Sub
Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
WinSettingGlobalForm Me   ', 6885, 11500
   Global_Grid
   TopCtrl1.TopText2 = "Add"
   'If Mid(UserPermission(Me.Caption), 4, 1) = "*" Then BTNPRINT.Enabled = False
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_Resize()
'    If GRepFormName <> SaleChalRep And GRepFormName <> SaleBillRep Then
'        If Me.WindowState <> vbMinimized Then
'            WinSetting Me ', 6885, 11500
'        End If
'    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If GridSel(4).Visible = True Then Set RsGrid1 = Nothing
If GridSel(1).Visible = True Then Set RsGrid2 = Nothing
If GridSel(2).Visible = True Then Set RsGrid3 = Nothing
If GridSel(3).Visible = True Then Set RsGrid4 = Nothing
If GridSel(4).Visible = True Then Set RsGrid5 = Nothing 'Up
If GridSel(5).Visible = True Then Set RsGrid6 = Nothing 'Up
Set RstRep = Nothing
Set mListItem = Nothing
Set rpt = Nothing
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
If GridSel(Index).Col = 0 Then
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
        Case 5
            I = UBound(GridRow5) + 1
            ReDim Preserve GridRow5(I)
            GridRow5(I) = GridSel(Index).Row
        Case 6
            I = UBound(GridRow6) + 1
            ReDim Preserve GridRow6(I)
            GridRow6(I) = GridSel(Index).Row
        Case 7
            I = UBound(GridRow7) + 1
            ReDim Preserve GridRow7(I)
            GridRow7(I) = GridSel(Index).Row
        Case 8
            I = UBound(GridRow8) + 1
            ReDim Preserve GridRow8(I)
            GridRow8(I) = GridSel(Index).Row
        Case 9
            I = UBound(GridRow9) + 1
            ReDim Preserve GridRow9(I)
            GridRow9(I) = GridSel(Index).Row
        Case 10
            I = UBound(GridRow10) + 1
            ReDim Preserve GridRow10(I)
            GridRow10(I) = GridSel(Index).Row
        Case 11
            I = UBound(GridRow11) + 1
            ReDim Preserve GridRow11(I)
            GridRow11(I) = GridSel(Index).Row
        Case 12
            I = UBound(GridRow12) + 1
            ReDim Preserve GridRow12(I)
            GridRow12(I) = GridSel(Index).Row

    End Select
  End If
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
    Case 5
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid5, KeyAscii, RsGrid5.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 6
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid6, KeyAscii, RsGrid6.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 7
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid7, KeyAscii, RsGrid7.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 8
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid8, KeyAscii, RsGrid8.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 9
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid9, KeyAscii, RsGrid9.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 10
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid10, KeyAscii, RsGrid10.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 11
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid11, KeyAscii, RsGrid11.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 12
       SelGridKeyPressLocal TxtSearch, GridSel, Index, RsGrid12, KeyAscii, RsGrid12.Fields(GridSel(Index).Col).Name, CellBackColEnter1, CellBackColLeave1
End Select
TxtSearch.Tag = Index
End Sub
Private Sub ListView1_Click()                   'UP
    TxtGrid(0).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    TxtGrid(0).SetFocus
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
    Case 5
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid5, KeyAscii, RsGrid5.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 6
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid6, KeyAscii, RsGrid6.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
     Case 7
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid7, KeyAscii, RsGrid7.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 8
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid8, KeyAscii, RsGrid8.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 9
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid9, KeyAscii, RsGrid9.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 10
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid10, KeyAscii, RsGrid10.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 11
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid11, KeyAscii, RsGrid11.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
    Case 12
       SelGridKeyPressLocal TxtSearch, GridSel, Val(TxtSearch.Tag), RsGrid12, KeyAscii, RsGrid12.Fields(GridSel(Val(TxtSearch.Tag)).Col).Name, CellBackColEnter1, CellBackColLeave1
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
        Case 5
            I = UBound(GridRow5) + 1
            ReDim Preserve GridRow5(I)
            GridRow5(I) = GridSel(Index).Row
        Case 6
            I = UBound(GridRow6) + 1
            ReDim Preserve GridRow6(I)
            GridRow6(I) = GridSel(Index).Row
        Case 7
            I = UBound(GridRow7) + 1
            ReDim Preserve GridRow7(I)
            GridRow7(I) = GridSel(Index).Row
        Case 8
            I = UBound(GridRow8) + 1
            ReDim Preserve GridRow8(I)
            GridRow8(I) = GridSel(Index).Row
        Case 9
            I = UBound(GridRow9) + 1
            ReDim Preserve GridRow9(I)
            GridRow9(I) = GridSel(Index).Row
        Case 10
            I = UBound(GridRow10) + 1
            ReDim Preserve GridRow10(I)
            GridRow10(I) = GridSel(Index).Row
        Case 11
            I = UBound(GridRow11) + 1
            ReDim Preserve GridRow11(I)
            GridRow11(I) = GridSel(Index).Row
        Case 12
            I = UBound(GridRow12) + 1
            ReDim Preserve GridRow12(I)
            GridRow12(I) = GridSel(Index).Row
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
Dim RsTemp As ADODB.Recordset, rsTemp1 As ADODB.Recordset
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    Select Case FGrid.Row
        Case List1
            Select Case GRepFormName
                Case VehReposesReg
                    ListArray = Array("All", "ReposesOnly", "ReleasedOnly")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
                Case VehSaleReg
                    ListArray = Array("Summary", "Detailed")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
                Case PurSaleTaxSumm
                    ListArray = Array("Spare Purchase Tax Summary", "Spare Sale Tax Summary", "Vehicle Purchase Tax Summary", "Vehicle Sale Tax Summary", "Service Tax Register", "Spare Sale & Labour Summary")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 6)
                Case VehSumModel
                    ListArray = Array("Qty Wise", "Value Wise", "Both")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 3)
            End Select
        Case List2
            Select Case GRepFormName
                Case VehSaleReg
                    ListArray = Array("SalesManWise", "PartyWise", "CityWise", "FinancierGrp", "FinancierName", "FormType", "Insu.Auth.", "Site Wise", "Model", "Model Category", "Model Group", "Godown", "All")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 13)
                Case VehSumModel
                    ListArray = Array("No", "Yes")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            End Select
        Case List3
            Select Case GRepFormName
                Case VehSumModel
                    ListArray = Array("Model Category", "Site")
                    Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray, 2)
            End Select
        Case List4
            ListArray1 = Array("No", "Yes")
            Set mListItem = ListView_Items(ListView, TxtGrid, Index, ListArray1, 2)
    Case List5
        Select Case GRepFormName
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
Case List1, List2, List3, List4, List5
        ListViewReport_KeyDown FrmList, ListView, TxtGrid, 0, KeyCode, Shift, TxtGrid(0).left, (TxtGrid(0).top + TxtGrid(0).height + 25), TxtGrid(0).width

        If KeyCode = vbKeyReturn Then
                If TxtGridLeave = True Then TxtKeyDown
            End If
Case Date1, Date2, Date3, Date4, Cat1, Cat2, Cat3, Cat4, Cat5, FromVno, ToVno
    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
        If TxtGridLeave = True Then TxtKeyDown
    End If
End Select
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
 Call CheckQuote(KeyAscii)
 Select Case GRepFormName
 End Select
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case FGrid.Row
        Case List1, List2, List3, List4, List5
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
Dim Grid1Sql As String: Dim FromVn As Integer, ToVn As Integer
Select Case FGrid.Row
        Case Date1
            Select Case GRepFormName
                Case VehOfftakeNRetail
                    FGrid = RetDate(TxtGrid(0))
                    FGrid = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(FGrid, 8))))
                Case Else
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
            End Select
        Case Cat1, Cat2, Cat3, Cat4, Cat5
             FGrid.TextMatrix(FGrid.Row, FGrid.Col) = TxtGrid(0)
        Case List1, List2, List3, List4, List5
            If TxtGrid(0).TEXT <> "" Then FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ListView.SelectedItem.TEXT
        Case Date1, Date2, Date3, Date4
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(0))
        Case FromVno
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Val(TxtGrid(0))
        Case ToVno
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Val(TxtGrid(0))
            FromVn = Val(FGrid.TextMatrix(FromVno, 1))
            ToVn = Val(FGrid.TextMatrix(ToVno, 1))
             If FromVn > ToVn Then
                MsgBox " 'From' Voucher No Must Be Greater Then 'TO' Voucher No.", vbInformation, "Renge"
             End If
End Select
    TxtGridLeave = True
    If ValidateCall = False Then
        FGrid.SetFocus
        TxtGrid(0).Visible = False
    End If
End Function
'******* Fuctions **********
Private Sub Global_Grid()
Dim I As Integer, Cnt As Integer, mHeight As Integer, mTop As Integer
FGrid.Rows = 10  '5
FGrid.Cols = 3
FGrid.FixedCols = 1
FGrid.ColWidth(0) = 2200
FGrid.ColWidth(1) = 3000
FGrid.ColWidth(2) = 0
FGrid.ColAlignment(1) = flexAlignLeftCenter
For I = 0 To FGrid.Rows - 1
    FGrid.RowHeight(I) = 0
Next
Ini_Grid
FGrid.height = (((mLastRow + 1) - mFirstRow) * PubGridRowHeight) + 800
FGrid.top = 75: FGrid.left = 100
mHeight = Me.height - FGrid.height - 1500
mTop = FGrid.top + FGrid.height + 500
If mHelpGridNo <= 3 Then
    mHeight = mHeight
ElseIf mHelpGridNo > 3 And mHelpGridNo <= 6 Then
    mHeight = (mHeight / 2)
ElseIf mHelpGridNo > 6 And mHelpGridNo <= 9 Then
    mHeight = (mHeight / 3)
ElseIf mHelpGridNo > 9 And mHelpGridNo <= 12 Then
    mHeight = (mHeight / 4)
End If
For I = 1 To mHelpGridNo
    GridSel(I).height = mHeight
    Select Case I
        Case 1, 4, 7, 10
            GridSel(I).left = FGrid.left
        Case 2, 5, 8, 11
            GridSel(I).left = FGrid.left + GridSel(I).width + 500
        Case 3, 6, 9, 12
            GridSel(I).left = FGrid.left + (2 * GridSel(I).width) + 1000
    End Select
    Select Case I
        Case 1, 2, 3
            GridSel(I).top = mTop
        Case 4, 5, 6
            GridSel(I).top = mTop + GridSel(I).height + 100
        Case 7, 8, 9
            GridSel(I).top = mTop + (2 * GridSel(I).height) + 200
        Case 10, 11, 12
            GridSel(I).top = mTop + (3 * GridSel(I).height) + 300
    End Select
    Check1(I).top = GridSel(I).top + 20: Check1(I).left = GridSel(I).left + 40
Next
'Nra Update for setting up the Form Size
'Select Case GRepFormName
'    Case SaleChalRep, SaleBillRep
'        GridSel(1).top = FGrid.top + FGrid.height + 100: GridSel(1).left = FGrid.left: GridSel(1).height = FGrid.height: GridSel(1).width = FGrid.width
'        BTNPRINT.left = 4500 / 2 - BTNPRINT.width: BTNEXIT.left = 4500 / 2 + 50: BTNPRINT.top = GridSel(1).top + GridSel(1).height + 500: BTNEXIT.top = BTNPRINT.top
'        Me.width = Me.width / 2: Me.left = 3000: Me.height = Me.height - 500
'        Check1(1).top = GridSel(1).top + 20: Check1(1).left = GridSel(1).left + 40
'        FGrid.Row = List1
'End Select
End Sub
Private Sub Grid_Hide()
If FrmList.Visible = True Then FrmList.Visible = False
End Sub
Private Sub FGrid_DblClick()
    Select Case FGrid.Row
        Case Date1, Date2, Date3, Date4, List1, List2, List3, List4, List5, Cat1, Cat2, Cat3, Cat4, Cat5, FromVno, ToVno
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
    End Select
TAddMode = False
End Sub
Private Sub FGrid_KeyPress(KeyAscii As Integer)
Dim I As Integer
    Select Case FGrid.Row
        Case Cat1, Cat2, Cat3, Cat4, Cat5
           Call Get_Text(Me, FGrid, TxtGrid, 0, True, KeyAscii)
        Case Date1, Date2, Date3, Date4, List1, List2, List3, List4, List5, ToVno, FromVno
            Call Get_Text(Me, FGrid, TxtGrid, 0, False, KeyAscii)
    End Select
If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'Leave Cell--> Enter Cell-->KeyDown
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
If KeyCode = vbKeyReturn Then
    Select Case FGrid.Row
        Case Date1, Date2, Date3, Date4, List1, List2, List3, List4, List5, Cat1, Cat2, Cat3, Cat4, Cat5
            Call GridDblClick(Me, FGrid, TxtGrid, 0)
            TAddMode = False
    End Select
End If
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
Dim formulastr As String   'Modishekhar 17 mar
formulastr = "" 'Modishekhar 17 mar
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
           'Modishekhar 17 mar
            If Len(formulastr + GridSel(Gridindex).TextMatrix(GridRow, 2)) < 255 Then
                formulastr = formulastr + IIf(formulastr = "", "For " & GridSel(Gridindex).TextMatrix(0, 1) & " : " & GridSel(Gridindex).TextMatrix(GridRow, 1), "," & GridSel(Gridindex).TextMatrix(GridRow, 1))
            End If
           'Modi End
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
'Modishekhar 17 mar
    Select Case Gridindex
        Case 1
            FormulaStr1 = mID(formulastr, 1, 254)
        Case 2
            FormulaStr2 = mID(formulastr, 1, 254)
        Case 3
            FormulaStr3 = mID(formulastr, 1, 254)
        Case 4
            FormulaStr4 = mID(formulastr, 1, 254)
        Case 5
            FormulaStr5 = mID(formulastr, 1, 254)
        Case 6
            FormulaStr6 = mID(formulastr, 1, 254)
        Case 7
            FormulaStr7 = mID(formulastr, 1, 254)
        Case 8
            FormulaStr8 = mID(formulastr, 1, 254)
        Case 9
            FormulaStr9 = mID(formulastr, 1, 254)
        Case 10
            FormulaStr10 = mID(formulastr, 1, 254)
        Case 11
            FormulaStr11 = mID(formulastr, 1, 254)
        Case 12
            FormulaStr12 = mID(formulastr, 1, 254)
    End Select
    'modi end
    
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
On Error Resume Next
Dim I As Integer
    If FGrid.Row = mLastRow Then SendKeysA vbKeyTab, True: Exit Sub
    For I = FGrid.Row To FGrid.Rows - 1
          If FGrid.RowHeight(I + 1) <> 0 Then FGrid.Row = I + 1: Exit For
    Next
End Sub
Private Sub GridInitialise(Gridindex As Integer, GridSql As String, Optional UseFaCn As Boolean)
Dim Index As Integer
Index = Gridindex

If UseFaCn Then
    If Index = 1 Then
        Set RsGrid1 = New ADODB.Recordset: RsGrid1.CursorLocation = adUseClient
        RsGrid1.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid1
        ReDim Preserve GridRow1(0)
        GridRow1(0) = 0
    End If
    If Index = 2 Then
        Set RsGrid2 = New ADODB.Recordset: RsGrid2.CursorLocation = adUseClient
        RsGrid2.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid2
        ReDim Preserve GridRow2(0)
        GridRow2(0) = 0
    End If
    If Index = 3 Then
        Set RsGrid3 = New ADODB.Recordset: RsGrid3.CursorLocation = adUseClient
        RsGrid3.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid3
        ReDim Preserve GridRow3(0)
        GridRow3(0) = 0
    End If
    If Index = 4 Then
        Set RsGrid4 = New ADODB.Recordset: RsGrid4.CursorLocation = adUseClient
        RsGrid4.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid4
        ReDim Preserve GridRow4(0)
        GridRow4(0) = 0
    End If
    If Index = 5 Then
        Set RsGrid5 = New ADODB.Recordset: RsGrid5.CursorLocation = adUseClient
        RsGrid5.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid5
        ReDim Preserve GridRow5(0)
        GridRow5(0) = 0
    End If
    
    If Index = 6 Then
        Set RsGrid6 = New ADODB.Recordset: RsGrid6.CursorLocation = adUseClient
        RsGrid6.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid6
        ReDim Preserve GridRow6(0)
        GridRow6(0) = 0
    End If
    If Index = 7 Then
        Set RsGrid7 = New ADODB.Recordset: RsGrid7.CursorLocation = adUseClient
        RsGrid7.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid7
            ReDim Preserve GridRow7(0)
            GridRow7(0) = 0
    End If
    
    If Index = 8 Then
        Set RsGrid8 = New ADODB.Recordset: RsGrid8.CursorLocation = adUseClient
        RsGrid8.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid8
        ReDim Preserve GridRow8(0)
        GridRow8(0) = 0
    End If
    
    If Index = 9 Then
        Set RsGrid9 = New ADODB.Recordset: RsGrid9.CursorLocation = adUseClient
        RsGrid9.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid9
        ReDim Preserve GridRow9(0)
        GridRow9(0) = 0
    End If
    If Index = 10 Then
        Set RsGrid10 = New ADODB.Recordset: RsGrid10.CursorLocation = adUseClient
        RsGrid10.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid10
        ReDim Preserve GridRow10(0)
        GridRow10(0) = 0
    End If
    If Index = 11 Then
        Set RsGrid11 = New ADODB.Recordset: RsGrid11.CursorLocation = adUseClient
        RsGrid11.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid11
            ReDim Preserve GridRow11(0)
            GridRow11(0) = 0
    End If
    If Index = 12 Then
        Set RsGrid12 = New ADODB.Recordset: RsGrid12.CursorLocation = adUseClient
        RsGrid12.Open GridSql, G_FaCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid12
        ReDim Preserve GridRow12(0)
        GridRow12(0) = 0
    End If
Else
    If Index = 1 Then
        Set RsGrid1 = New ADODB.Recordset: RsGrid1.CursorLocation = adUseClient
        RsGrid1.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid1
        ReDim Preserve GridRow1(0)
        GridRow1(0) = 0
    End If
    If Index = 2 Then
        Set RsGrid2 = New ADODB.Recordset: RsGrid2.CursorLocation = adUseClient
        RsGrid2.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid2
        ReDim Preserve GridRow2(0)
        GridRow2(0) = 0
    End If
    If Index = 3 Then
        Set RsGrid3 = New ADODB.Recordset: RsGrid3.CursorLocation = adUseClient
        RsGrid3.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid3
        ReDim Preserve GridRow3(0)
        GridRow3(0) = 0
    End If
    If Index = 4 Then
        Set RsGrid4 = New ADODB.Recordset: RsGrid4.CursorLocation = adUseClient
        RsGrid4.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid4
        ReDim Preserve GridRow4(0)
        GridRow4(0) = 0
    End If
    If Index = 5 Then
        Set RsGrid5 = New ADODB.Recordset: RsGrid5.CursorLocation = adUseClient
        RsGrid5.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid5
        ReDim Preserve GridRow5(0)
        GridRow5(0) = 0
    End If
    
    If Index = 6 Then
        Set RsGrid6 = New ADODB.Recordset: RsGrid6.CursorLocation = adUseClient
        RsGrid6.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid6
        ReDim Preserve GridRow6(0)
        GridRow6(0) = 0
    End If
    If Index = 7 Then
        Set RsGrid7 = New ADODB.Recordset: RsGrid7.CursorLocation = adUseClient
        RsGrid7.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid7
            ReDim Preserve GridRow7(0)
            GridRow7(0) = 0
    End If
    
    If Index = 8 Then
        Set RsGrid8 = New ADODB.Recordset: RsGrid8.CursorLocation = adUseClient
        RsGrid8.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid8
        ReDim Preserve GridRow8(0)
        GridRow8(0) = 0
    End If
    
    If Index = 9 Then
        Set RsGrid9 = New ADODB.Recordset: RsGrid9.CursorLocation = adUseClient
        RsGrid9.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid9
        ReDim Preserve GridRow9(0)
        GridRow9(0) = 0
    End If
    If Index = 10 Then
        Set RsGrid10 = New ADODB.Recordset: RsGrid10.CursorLocation = adUseClient
        RsGrid10.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid10
        ReDim Preserve GridRow10(0)
        GridRow10(0) = 0
    End If
    If Index = 11 Then
        Set RsGrid11 = New ADODB.Recordset: RsGrid11.CursorLocation = adUseClient
        RsGrid11.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid11
            ReDim Preserve GridRow11(0)
            GridRow11(0) = 0
    End If
    If Index = 12 Then
        Set RsGrid12 = New ADODB.Recordset: RsGrid12.CursorLocation = adUseClient
        RsGrid12.Open GridSql, GCn, adOpenStatic, adLockReadOnly: Set GridSel(Index).DataSource = RsGrid12
        ReDim Preserve GridRow12(0)
        GridRow12(0) = 0
    End If
End If
GridSel(Index).Visible = True: GridSel(Index).Enabled = False: Check1(Index).Visible = True
GridSel(Index).ColWidth(0) = 600: GridSel(Index).ColWidth(2) = 0: GridSel(Index).ColWidth(1) = 2000: GridSel(Index).ColWidth(3) = 1000
Check1(Index).width = 580: Check1(Index).height = GridSel(Index).RowHeight(0) + 20 ' modishekhar  251103: Check1(Index).Value = Checked
'GridSel(Index).width = 3000
End Sub

Private Sub Ini_Grid()
Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where  site_code='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If


Select Case GRepFormName
    Case VehReposesReg
        With FGrid
            .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
            .TextMatrix(Date2, 0) = "To Date": .RowHeight(Date2) = GridRowHeight
            .TextMatrix(List1, 0) = "Reposes Status": .RowHeight(List1) = GridRowHeight
            .TextMatrix(List5, 0) = "Case Status": .RowHeight(List5) = GridRowHeight
            .TextMatrix(List5, 1) = "Running"
            .TextMatrix(Date1, 1) = PubStartDate
            .TextMatrix(Date2, 1) = PubLoginDate
            .TextMatrix(List1, 1) = "All"
            .TextMatrix(List4, 0) = "With Address": .RowHeight(List4) = GridRowHeight
            .TextMatrix(List4, 1) = "No"
            
        End With
        mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 9
        Call CreateHelpGrid
    Case VehSaleReg    'vijay Vehicle 16/11/02
            With FGrid
                .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
                .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
                .TextMatrix(List1, 0) = "Summary/Detailed": .RowHeight(List1) = GridRowHeight
                .TextMatrix(List2, 0) = "Type": .RowHeight(List2) = GridRowHeight
    
                .TextMatrix(Date1, 1) = PubStartDate
                .TextMatrix(Date2, 1) = PubLoginDate
                .TextMatrix(List1, 1) = "Summary"
                .TextMatrix(List2, 1) = "PartyWise"
            End With
            mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 9
            Call CreateHelpGrid
        
    Case PurSaleTaxSumm
            With FGrid
                .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
                .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
                .TextMatrix(List1, 0) = "Type": .RowHeight(List1) = GridRowHeight
    
    
                .TextMatrix(Date1, 1) = PubStartDate
                .TextMatrix(Date2, 1) = PubLoginDate
                .TextMatrix(List1, 1) = "Spare Sale Tax Summary"
                
            End With
            mFirstRow = Date1: mLastRow = List1: mHelpGridNo = 2
            
            
            Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
            GridInitialise 1, Grid1Sql
            
            
            Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
            GridInitialise 2, Grid2Sql
    Case VehSumModel
            With FGrid
                .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
                .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
                .TextMatrix(List1, 0) = "Type": .RowHeight(List1) = GridRowHeight
                .TextMatrix(List2, 0) = "Model Group": .RowHeight(List2) = GridRowHeight
                .TextMatrix(List3, 0) = "Group By": .RowHeight(List3) = GridRowHeight
                
                .TextMatrix(Date1, 1) = IIf(FGrid.TextMatrix(Date1, 1) = "", PubStartDate, FGrid.TextMatrix(Date1, 1))
                .TextMatrix(Date2, 1) = IIf(FGrid.TextMatrix(Date2, 1) = "", PubLoginDate, FGrid.TextMatrix(Date2, 1))
                .TextMatrix(List1, 1) = IIf(FGrid.TextMatrix(List1, 1) = "", "Both", FGrid.TextMatrix(List1, 1))
                .TextMatrix(List2, 1) = IIf(FGrid.TextMatrix(List2, 1) <> "", TxtGrid(0), "No")
                .TextMatrix(List3, 1) = IIf(FGrid.TextMatrix(List3, 1) <> "", TxtGrid(0), "Model Category")
            End With
            
            mFirstRow = Date1: mLastRow = List2: mHelpGridNo = 9
            
            Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
            GridInitialise 1, Grid1Sql
                        
            Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
            GridInitialise 2, Grid2Sql
        
            Grid4Sql = "select '' as O,Model.Model As Model_Description, Model.Model As Code From Model order by Model.Model"
            GridInitialise 4, Grid4Sql
    
            Grid5Sql = "select '' as O, ModelGrp_Name As  Model_Group, ModelGrp_Code As Code From Model_Grp Order by ModelGrp_Name"
            GridInitialise 5, Grid5Sql
    
            Grid6Sql = "select '' as O, ModelCat_Name As  Model_Category, ModelCat_Code As Code From Model_Cat Order by ModelCat_Name"
            GridInitialise 6, Grid6Sql
    
            Grid7Sql = "select '' as O, God_Name As Godown, God_Code As Code From Godown Where Appli_For=1 Order by God_Name"
            GridInitialise 7, Grid7Sql
    
            Grid8Sql = "select '' as O,ContractFinance.FinName As FinancerName,ContractFinance.FinCode As Code from ContractFinance order by ContractFinance.FinName"
            GridInitialise 8, Grid8Sql
    
            Grid9Sql = "select '' as O,Emp_Name As SalesMan,Emp_Code As Code from Emp_Mast order by Emp_Name"
            GridInitialise 9, Grid9Sql
            
            
            Call CreateHelpGrid
        
        
        Case VehOfftakeNRetail
            With FGrid
                .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
                .TextMatrix(Date1, 1) = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(PubLoginDate, 8))))
            End With
            
            
            mFirstRow = Date1: mLastRow = Date1: mHelpGridNo = 9
            
            Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & "order by site_desc"
            GridInitialise 1, Grid1Sql
            
            
            Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
            GridInitialise 2, Grid2Sql
    
    
            Grid3Sql = "select '' AS O,SubGroup.NAME as Party_Name,SubGroup.SubCode as Code from SubGroup " & _
                "left join  " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
                "Where left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
                " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
            GridInitialise 3, Grid3Sql
    
    
            Grid4Sql = "select '' as O,Model.Model As Model_Description, Model.Model As Code From Model order by Model.Model"
            GridInitialise 4, Grid4Sql
    
    
            Grid5Sql = "select '' as O, ModelGrp_Name As  Model_Group, ModelGrp_Code As Code From Model_Grp Order by ModelGrp_Name"
            GridInitialise 5, Grid5Sql
    
    
            Grid6Sql = "select '' as O, ModelCat_Name As  Model_Category, ModelCat_Code As Code From Model_Cat Order by ModelCat_Name"
            GridInitialise 6, Grid6Sql
    
    
            Grid7Sql = "select '' as O, God_Name As  Godown, God_Code As Code From Godown Where Appli_For=1 Order by God_Name"
            GridInitialise 7, Grid7Sql

        
    Case BodyBuilderChassis, StockAtBodyBuilder
            With FGrid
                .TextMatrix(Date1, 0) = "From Date": .RowHeight(Date1) = GridRowHeight
                .TextMatrix(Date2, 0) = "UpTo Date": .RowHeight(Date2) = GridRowHeight
                
                .TextMatrix(Date1, 1) = IIf(FGrid.TextMatrix(Date1, 1) = "", PubStartDate, FGrid.TextMatrix(Date1, 1))
                .TextMatrix(Date2, 1) = IIf(FGrid.TextMatrix(Date2, 1) = "", PubLoginDate, FGrid.TextMatrix(Date2, 1))
            End With
            
            mFirstRow = Date1: mLastRow = Date2: mHelpGridNo = 9
            
            Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & " order by site_desc"
            GridInitialise 1, Grid1Sql
                        
            Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
            GridInitialise 2, Grid2Sql
        
            Grid4Sql = "select '' as O,Model.Model As Model_Description, Model.Model As Code From Model order by Model.Model"
            GridInitialise 4, Grid4Sql
    
            Grid5Sql = "select '' as O, ModelGrp_Name As  Model_Group, ModelGrp_Code As Code From Model_Grp Order by ModelGrp_Name"
            GridInitialise 5, Grid5Sql
    
            Grid6Sql = "select '' as O, ModelCat_Name As  Model_Category, ModelCat_Code As Code From Model_Cat Order by ModelCat_Name"
            GridInitialise 6, Grid6Sql
    
            Grid7Sql = "select '' as O, God_Name As Godown, God_Code As Code From Godown Where Appli_For=1 Order by God_Name"
            GridInitialise 7, Grid7Sql
    
            Grid8Sql = "select '' as O, BodyBuilderDesc As Body_Builder_Name, BodyBuilderCode As Code from BodyBuilder order by BodyBuilderDesc"
            GridInitialise 8, Grid8Sql
    
            Grid9Sql = "select '' as O, BodyTypeDesc As Body_Type,BodyTypeCode As Code from BodyType order by BodyTypeDesc"
            GridInitialise 9, Grid9Sql
            
            
            'Call CreateHelpGrid
End Select
End Sub

Private Function IsNotBlank(FieldRow As Integer, FieldCaption As String) As Boolean
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
'Modishekhar 17 mar
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("DATEBETWEEN")
                rpt.FormulaFields(I).TEXT = "'From :'+ '" & Format(FGrid.TextMatrix(Date1, 1), "dd/mmm/yyyy") & "' + ' To ' + '" & Format(FGrid.TextMatrix(Date2, 1), "dd/mmm/yyyy") & "'"
            Case UCase("Formulastr1")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr1 & "'"
            Case UCase("Formulastr2")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr2 & "'"
            Case UCase("Formulastr3")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr3 & "'"
            Case UCase("Formulastr4")
                rpt.FormulaFields(I).TEXT = "'" & FormulaStr4 & "'"
            Case UCase("RepTitle")
                rpt.FormulaFields(I).TEXT = "'Daily Sales Report-'+ '" & Format(FGrid.TextMatrix(Date1, 1), "mmm-yyyy") & "' "
            Case UCase("list1")
                rpt.FormulaFields(I).TEXT = " '" & FGrid.TextMatrix(List1, 1) & "'"
            Case UCase("List3")
                rpt.FormulaFields(I).TEXT = " '" & FGrid.TextMatrix(List3, 1) & "'"
                
        End Select
    Next
    FormulaStr1 = "": FormulaStr2 = "": FormulaStr3 = "": FormulaStr4 = ""
    

    
    
    
    
    Select Case GRepFormName
        Case VehSaleReg  'Vijay
            For I = 1 To rpt.FormulaFields.Count
                Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                    Case UCase("CAmt")
                        rpt.FormulaFields(I).TEXT = " '" & RstRep1!Cancel_Amt & "'"
                    Case UCase("CTax")
                        rpt.FormulaFields(I).TEXT = " '" & RstRep1!CancelTax_Amt & "'"
                    Case UCase("CTaxLocal")
                        rpt.FormulaFields(I).TEXT = " '" & VNull(RstRep1!TaxLocal) & "'"
                    Case UCase("CTaxCentral")
                        rpt.FormulaFields(I).TEXT = " '" & VNull(RstRep1!TaxCentral) & "'"
                    Case UCase("CTot")
                        rpt.FormulaFields(I).TEXT = " '" & RstRep1!CancelTOT_Amt & "'"
                    Case UCase("TOTCaption")
                        rpt.FormulaFields(I).TEXT = " '" & pubTOTCaption & "'"
                End Select
            Next
    End Select





Exit Sub
ELoop:
     MsgBox err.Description
End Sub


Public Sub SelGridKeyPressLocal(txt As Object, SelGrid As Object, Index As Integer, Rst As ADODB.Recordset, ByRef KeyAscii As Integer, FindFldName As String, Optional CellBackColEnter As ColorConstants, Optional CellBackColLeave As ColorConstants)
Dim FindStr$    ' As String
Dim LPlace As Byte
'    If FilterKeyCode(KeyAscii) = True Then Exit Sub
    If SelGrid(Index).Rows < 1 Then Exit Sub
    If Rst.RecordCount <= 0 Then txt.TEXT = "": Exit Sub
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then Exit Sub  ''modishekhar 251103
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
            Rst.FIND "" & FindFldName & " >=" & Val(FindStr) & ""
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

Public Function ListView_Items_RecordSet_Local(LV As Object, txt As Object, Index As Integer, Rst As ADODB.Recordset) As ListItem
    Dim xName As ListItem
    Dim I As Long
    LV.ListItems.Clear
        
    If Rst.RecordCount <= 0 Then Exit Function
    Set xName = LV.ListItems.Add(, , "All")
    xName.SubItems(1) = ""
    Do Until Rst.EOF
        Set xName = LV.ListItems.Add(, , Rst.Fields("Name").Value)
        If Not IsNull(Rst.Fields("Code").Value) Then
            xName.SubItems(1) = CStr(Rst.Fields("code").Value)
        End If
    Rst.MoveNext
    Loop
    Set xName = LV.FindItem(txt(Index), 0, , 1)
    If xName Is Nothing Then
        Exit Function
    Else
        xName.EnsureVisible
        xName.SELECTED = True
    End If
    Set ListView_Items_RecordSet_Local = xName
End Function


Private Sub GridSelMax(Enb As Boolean)
'modishekhar  251103
If CmdMin.Tag = "" Then Exit Sub
    If Enb = True Then
        ActiveGridTop = GridSel(Val(CmdMin.Tag)).top
        ActiveGridLeft = GridSel(Val(CmdMin.Tag)).left
        ActiveGridHeight = GridSel(Val(CmdMin.Tag)).height
        ActiveGridWidth = GridSel(Val(CmdMin.Tag)).width
        GridSel(Val(CmdMin.Tag)).top = GridSel(1).top
        GridSel(Val(CmdMin.Tag)).left = GridSel(1).left
        
        GridSel(Val(CmdMin.Tag)).width = 11500
        GridSel(Val(CmdMin.Tag)).ColWidth(1) = 10500
        GridSel(Val(CmdMin.Tag)).height = Me.height - FGrid.height - 1000
        Check1(Val(CmdMin.Tag)).left = GridSel(Val(CmdMin.Tag)).left + 20
        Check1(Val(CmdMin.Tag)).top = GridSel(Val(CmdMin.Tag)).top + 40
        GridSel(Val(CmdMin.Tag)).ZOrder 0
        Check1(Val(CmdMin.Tag)).ZOrder 0
        CmdMin.ZOrder 0
        CmdMin.top = GridSel(Val(CmdMin.Tag)).top
        CmdMin.left = GridSel(Val(CmdMin.Tag)).left + GridSel(Val(CmdMin.Tag)).width - CmdMin.width - 350
        CmdMin.Visible = True
    Else
        GridSel(Val(CmdMin.Tag)).top = ActiveGridTop
        GridSel(Val(CmdMin.Tag)).left = ActiveGridLeft
        GridSel(Val(CmdMin.Tag)).height = ActiveGridHeight
        GridSel(Val(CmdMin.Tag)).width = ActiveGridWidth
        GridSel(Val(CmdMin.Tag)).ColWidth(1) = 2000
        Check1(Val(CmdMin.Tag)).left = GridSel(Val(CmdMin.Tag)).left + 20
        Check1(Val(CmdMin.Tag)).top = GridSel(Val(CmdMin.Tag)).top + 40
        GridSel(Val(CmdMin.Tag)).ZOrder 0
        Check1(Val(CmdMin.Tag)).ZOrder 0
        CmdMin.Visible = False
        CmdMin.Tag = ""
    End If
End Sub
Public Sub WinSettingGlobalForm(ByRef FrmName As Form)
On Error Resume Next
With FrmName
    .height = 7585 '8000
    .width = 11920 '12000
    .top = 0
    .left = 0
End With
End Sub

Private Sub CreateHelpGrid()
 Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where  site_code='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If


        Grid1Sql = "select '' as O,site_desc as SiteName,site_code  as code from site " & sitecond & "order by site_desc"
        GridInitialise 1, Grid1Sql
        
        
        Grid2Sql = "select '' as O,Division.Div_Name As DivisionName,Division.Div_Code As Code from Division order by Division.Div_Name"
        GridInitialise 2, Grid2Sql

        Grid3Sql = "select '' AS O,SubGroup.NAME as Party_Name,SubGroup.SubCode as Code from SubGroup " & _
            "left join  " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
            "Where left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
            " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
        GridInitialise 3, Grid3Sql

        Grid4Sql = "select '' as O,Model.Model As Model_Description, Model.Model As Code From Model order by Model.Model"
        GridInitialise 4, Grid4Sql

        Grid5Sql = "select '' as O, ModelGrp_Name As  Model_Group, ModelGrp_Code As Code From Model_Grp Order by ModelGrp_Name"
        GridInitialise 5, Grid5Sql

        Grid6Sql = "select '' as O, ModelCat_Name As  Model_Category, ModelCat_Code As Code From Model_Cat Order by ModelCat_Name"
        GridInitialise 6, Grid6Sql

        Grid7Sql = "select '' as O, God_Name As  Godown, God_Code As Code From Godown Where Appli_For=1 Order by God_Name"
        GridInitialise 7, Grid7Sql

        Grid8Sql = "select '' as O,ContractFinance.FinName As FinancerName,ContractFinance.FinCode As Code from ContractFinance order by ContractFinance.FinName"
        GridInitialise 8, Grid8Sql

        Grid9Sql = "select '' as O,Emp_Name As SalesMan,Emp_Code As Code from Emp_Mast order by Emp_Name"
        GridInitialise 9, Grid9Sql

End Sub
Private Function CheckGridSel() As Boolean
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 0): If RepPrint = False Then Exit Function
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Function
    If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1): If RepPrint = False Then Exit Function
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Function
    If Check1(5).Value = Unchecked Then GridString5 = FillString(GridRow5, 5, 1): If RepPrint = False Then Exit Function
    If Check1(6).Value = Unchecked Then GridString6 = FillString(GridRow6, 6, 1): If RepPrint = False Then Exit Function
    If Check1(7).Value = Unchecked Then GridString7 = FillString(GridRow7, 7, 1): If RepPrint = False Then Exit Function
    If Check1(8).Value = Unchecked Then GridString8 = FillString(GridRow8, 8, 1): If RepPrint = False Then Exit Function
    If Check1(9).Value = Unchecked Then GridString9 = FillString(GridRow9, 9, 1): If RepPrint = False Then Exit Function
    CheckGridSel = True
End Function
Private Sub MakeSelection()
    If FGrid.TextMatrix(List5, 1) = "Running" Then
        Condstr = Condstr & " And (CaseFlag Is Null Or CaseFlag='R' Or Trim(CaseFlag)='')"
    ElseIf FGrid.TextMatrix(List5, 1) = "Closed" Then
        Condstr = Condstr & " And CaseFlag='C' "
    End If
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " And CaseData.V_No in (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " And CaseData.IntroducerCode in (" & GridString2 & ")"
    If Check1(3).Value = Unchecked Then Condstr = Condstr & " And CaseData.HirerCode in (" & GridString3 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " And Subgroup.CityCode in (" & GridString4 & ")"
    If Check1(5).Value = Unchecked Then Condstr = Condstr & " And Subgroup.AreaCode in (" & GridString5 & ")"
    If Check1(6).Value = Unchecked Then Condstr = Condstr & " And CaseData.ModelCode in (" & GridString6 & ")"
    If Check1(7).Value = Unchecked Then Condstr = Condstr & " And CaseData.InspectorCode in (" & GridString7 & ")"
    If Check1(8).Value = Unchecked Then Condstr = Condstr & " And CaseData.SchemeCode in (" & GridString8 & ")"
    If Check1(9).Value = Unchecked Then Condstr = Condstr & " And CaseData.Site_Code in (" & GridString9 & ")"
End Sub





Private Sub VehSaleRegProc()
On Error GoTo ELoop
Dim mQry As String, Condstr$
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
'    If IsNotBlank(List1, FGrid.TextMatrix(List2, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    If Check1(5).Value = Unchecked Then GridString5 = FillString(GridRow5, 5, 1): If RepPrint = False Then Exit Sub
    If Check1(6).Value = Unchecked Then GridString6 = FillString(GridRow6, 6, 1): If RepPrint = False Then Exit Sub
    If Check1(7).Value = Unchecked Then GridString7 = FillString(GridRow7, 7, 1): If RepPrint = False Then Exit Sub
    If Check1(8).Value = Unchecked Then GridString8 = FillString(GridRow8, 8, 1): If RepPrint = False Then Exit Sub
    If Check1(9).Value = Unchecked Then GridString9 = FillString(GridRow9, 9, 1): If RepPrint = False Then Exit Sub

    Condstr = " Where VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("VO.Inv_DocId", "3", "1") & " in (" & GridString1 & ")"
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(VO.Inv_DocId,1) in (" & GridString2 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and VO.Model in (" & GridString4 & ")"
    If Check1(5).Value = Unchecked Then Condstr = Condstr & " and Model.Grp_Code in (" & GridString5 & ")"
    If Check1(6).Value = Unchecked Then Condstr = Condstr & " and Model.Cat_Code in (" & GridString6 & ")"
    If Check1(7).Value = Unchecked Then Condstr = Condstr & " and VStk.Godown in (" & GridString7 & ")"
    If Check1(8).Value = Unchecked Then Condstr = Condstr & " and VO.FB_Code in (" & GridString8 & ")"
    If Check1(9).Value = Unchecked Then Condstr = Condstr & " and VO.Rep_Code in (" & GridString9 & ")"

    Select Case FGrid.TextMatrix(List2, 1)
        Case "PartyWise"
             If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
             If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.PartyCode in (" & GridString3 & ")"
        Case "CityWise"
             If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
             If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.PartyCode In (Select SG1.SubCode From SubGroup as SG1 where SG1.CityCode In (" & GridString3 & "))"
             
        Case "FinancierGrp"
             If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
             If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.FB_CODE in (Select ContractFinance.FinCode From ContractFinance where ContractFinance.UnderFinGrp in (" & GridString3 & "))"
            
        Case "FinancierName"
            If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.FB_CODE in (" & GridString3 & ")"
        Case "FormType"
            If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.Form_Code in (" & GridString3 & ")"
        Case "Insu.Auth."
            If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.RegBy in (" & GridString3 & ")"
        Case "SalesManWise"
            If Check1(3).Value = Unchecked Then GridString3 = FillString(GridRow3, 3, 1)
            If Check1(3).Value = Unchecked Then Condstr = Condstr & " and VO.Rep_Code in (" & GridString3 & ")"
             
            mQry = "SELECT VO.Inv_DocId, VO.Inv_Date, VO.OrdDocId,VO.Ord_Date,VO.Chassis as ChassisNo, VO.VRATE, " & _
                   "VO.MARGINE, VO.InciChrg, VO.Octroi, VO.Transport, VO.RegTemp, VO.TAX_Amt,VO.Surcharge_Amt, " & _
                   "VO.OtherChrg, VO.Net_AMOUNT, Vo.Insurance, VO.Form_Code,VO.Fin_Amt,VO.MISC_INFO,VO.TOT_Amt, " & _
                   "VO.SubTot, VO.RtoFee, SG.NamePrefix,SG.NAME, SG.FPrefix, SG.FName, SG.Add1, SG.Add2, SG.Add3, " & _
                   "City.CityName, Model.Model,VStk.EngineNo, VStk.PBILL_NO, VStk.PBILL_DATE,TF.Form_Desc, " & _
                   "FinGroup.FinGrpName,ContractFinance.FinName, Emp_Mast.Emp_Name, " & _
                   "" & cIIF("'" & UCase(left(PubComp_Name, 4)) & "'='ENAR'", "Model.Sales_Desc", "Mg.ModelGrp_Name") & " as Model_Group, " & _
                   "S.Site_Desc, G.God_Name As Godown_Name, MC.ModelCat_Name As ModelCategoryName, VStk.BodyBuilder_IssDate, BB.BodyBuilderDesc, Bt.BodyTypeDesc, VO.Rebate  " & _
                   " FROM (((((((((((((Veh_Order as VO LEFT JOIN Veh_Stock as VStk ON VO.chassis = VStk.ChassisNo) " & _
                   "Left Join BodyBuilder BB On BB.BodyBuilderCode =  Vstk.BodyBuilder) " & _
                   "Left Join BodyType Bt On Bt.BodyTypeCode =  Vstk.BodyBuilder_BodyType) " & _
                   "LEFT JOIN Model ON VO.MODEL = Model.MODEL) " & _
                   "Left Join Model_Grp Mg On Model.Grp_Code=Mg.ModelGrp_Code) " & _
                   "LEFT JOIN SubGroup as SG ON VO.PartyCode = SG.SubCode) " & _
                   "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
                   "LEFT JOIN ContractFinance ON VO.FB_CODE = ContractFinance.FinCode) " & _
                   "LEFT JOIN FinGroup ON ContractFinance.UnderFinGrp = FinGroup.FinGrpCode) " & _
                   "LEFT JOIN TaxForms as TF ON VO.Form_Code = TF.Form_Code) " & _
                   "LEFT JOIN Emp_Mast ON VO.Rep_Code = Emp_Mast.Emp_Code )" & _
                   "Left Join Site S On S.Site_Code= " & cMID("VO.Inv_DocId", "3", "1") & ") " & _
                   "Left Join Godown G On G.God_Code = VStk.Godown) " & _
                   "Left Join Model_Cat MC On MC.ModelCat_Code = Model.Cat_Code "
            GoTo NXT
        
    End Select
      
    mQry = "SELECT VO.Inv_DocId, VO.Inv_Date, VO.OrdDocId,VO.Ord_Date,SG.NamePrefix, SG.Name, SG.FPrefix, " & _
        "SG.FName, SG.Add1, SG.Add2, SG.Add3, City.CityName, Model.Model,VO.Chassis as ChassisNo, " & _
        "VStk.EngineNo,VO.VRATE, VStk.PBILL_NO, VStk.PBILL_DATE, VO.MARGINE, VO.InciChrg, VO.Octroi, " & _
        "VO.Transport, VO.RegTemp, VO.TAX_Amt,VO.Surcharge_Amt, VO.OtherChrg, VO.Net_AMOUNT,VO.Form_Code, " & _
        "TF.Form_Desc,FinGroup.FinGrpName, ContractFinance.FinName, VO.Fin_Amt, VO.MISC_INFO,VO.TOT_Amt, " & _
        "'" & FGrid.TextMatrix(List2, 1) & "' as ReportType, VO.SubTot, sg.phone, VO.RtoFee, VO.Insurance, " & _
        "MG.ModelGrp_Name, S.Site_Desc, G.God_Name As Godown_Name, MC.ModelCat_Name As ModelCategoryName, Model.Sales_Desc, Model.Model_Desc, 1 As Qty, VStk.BodyBuilder_IssDate, BB.BodyBuilderDesc, Bt.BodyTypeDesc, VO.Rebate, TF.L_C " & _
        "FROM ((((((((((((Veh_Order as VO LEFT JOIN Veh_Stock as VStk ON VO.chassis = VStk.ChassisNo) " & _
        "Left Join BodyBuilder BB On BB.BodyBuilderCode =  Vstk.BodyBuilder) " & _
        "Left Join BodyType Bt On Bt.BodyTypeCode =  Vstk.BodyBuilder_BodyType) " & _
        "LEFT JOIN Model ON VO.MODEL = Model.MODEL) " & _
        "Left Join Model_Grp Mg On Model.Grp_Code=Mg.ModelGrp_Code) " & _
        "LEFT JOIN SubGroup as SG ON VO.PartyCode = SG.SubCode) " & _
        "LEFT JOIN City ON SG.CityCode = City.CityCode) " & _
        "LEFT JOIN ContractFinance ON VO.FB_CODE = ContractFinance.FinCode) " & _
        "LEFT JOIN FinGroup ON ContractFinance.UnderFinGrp = FinGroup.FinGrpCode) " & _
        "LEFT JOIN TaxForms as TF ON VO.Form_Code = TF.Form_Code) " & _
        "Left Join Site S On S.Site_Code= " & cMID("VO.Inv_DocId", "3", "1") & ") " & _
        "Left Join Godown G On VStk.Godown = G.God_Code) " & _
        "Left Join Model_Cat MC On MC.ModelCat_Code = Model.Cat_Code "


NXT:
    mQry = mQry + Condstr & " and (Vstk.Sal_Vtype<>'V_TRF' Or VStk.Sal_VType Is Null) order by VO.Inv_Date,right(VO.Inv_DocId,8) "
    'mQRY = mQRY + Condstr & " and (Vstk.Sal_Vtype<>'V_TRF' Or VStk.Sal_VType Is Null)  "

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    Set RstRep1 = New ADODB.Recordset
    RstRep1.Open "Select Sum(VRATE+Margine) as Cancel_Amt,sum(Tax_Amt) as CancelTax_Amt,sum(TOT_Amt) as CancelTOT_Amt, Sum(" & cIIF("TF.L_C='Local'", "Tax_Amt", "0") & ") As TaxLocal, Sum(" & cIIF("TF.L_C='Central'", "Tax_Amt", "0") & ") As TaxCentral from Veh_Order1 As VO1 Left Join TaxForms TF On TF.Form_Code = VO1.Form_Code Where VO1.Ord_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO1.Ord_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " ", GCn, adOpenStatic, adLockReadOnly
    
    If FGrid.TextMatrix(List1, 1) = "Summary" Then
        If FGrid.TextMatrix(List3, 1) = "All" Then
            RepName = "VehSaleRegSumAll"
        Else
            If UCase(left(PubComp_Name, 4)) = "ENAR" Or UCase(left(PubComp_Name, 6)) = "J.M.A." Then
                RepName = "VehSaleRegSum_Enar"
            Else
                RepName = "VehSaleRegSum"
            End If
        End If
    ElseIf FGrid.TextMatrix(List1, 1) = "Detailed" Then
        If FGrid.TextMatrix(List3, 1) = "All" Then
            RepName = "VehSaleRegDetAll"
        Else
            If StrCmp(left(PubComp_Name, 4), "Enar") Or UCase(left(PubComp_Name, 6)) = "J.M.A." Then
                RepName = "VehSaleRegDet_Enar"
            Else
                RepName = "VehSaleRegDet"
            End If
        End If
    End If
    
    If FGrid.TextMatrix(List2, 1) = "SalesManWise" Then
        If FGrid.TextMatrix(List1, 1) = "Summary" Then
            RepName = "SalesManWiseSaleRepSum"
        Else
            RepName = "SalesManWiseSaleRep"
        End If
        Me.CAPTION = "Sales Man Wise Sale Report"
    End If
    
    RepTitle = UCase(Me.CAPTION)
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub




Private Sub SpeedPrintSumm()
    Dim PageWidth As Byte, PageLength As Integer, mHeader As Double, Counter As Double, mCounter As Double
    Dim isLast As Boolean, mRec As Integer, PageNo As Double
    Dim RstCompDet As ADODB.Recordset, TotalNetAmt As Double
    Dim fob As New FileSystemObject
    
    Set RstCompDet = GCn.Execute("select S_SecSpeciality,S_SecLST,S_SecLST_Date,S_SecCST,S_SecCST_Date,S_SecPhone,S_SecFax from division where Div_Code='" & PubDivCode & "' and S_SecCompCode =  '" & PubSCompCode & "'")
    
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    PageLength = PubPageLength
    PageWidth = 132
    mRec = 9
    'Header printing
    Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
    mHeader = mHeader + 1
    If XNull(RstCompDet!S_SecSpeciality) <> "" Then
        Print #1, PRN_TIT(RstCompDet!S_SecSpeciality, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
    mHeader = mHeader + 1
    If PubComp_Add2 <> "" Then
        Print #1, PRN_TIT(PubComp_Add2, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    If PubComp_City <> "" Then
        Print #1, PRN_TIT(PubComp_City, "C", PageWidth)
        mHeader = mHeader + 1
    End If
    
    Print #1, PRN_TIT("Vehicle Sale Register", "C", PageWidth)
    mHeader = mHeader + 1
    
    Print #1, "From : " & FGrid.TextMatrix(Date1, 1) & "  To : " & FGrid.TextMatrix(Date2, 1)
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, mChr17 & Space(5) & Space(10) & PSTR("Sales", 20) & PSTR("Order", 20) & PSTR("Name of Customer", 35) & PSTR("Model", 15) & PSTR("Telco", 15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(20)
    mHeader = mHeader + 1
    Print #1, PSTR("#", 5) & PSTR("Invoice", 10) & PSTR("Date", 20) & PSTR("Date", 20) & PSTR("Address", 35) & PSTR("Chassis-No", 15) & PSTR("Bill No", 15) & PSTR("Sale-Amt", 15) & PSTR("Financer Group", 35) & Space(15) & PSTR("Form", 5) & PSTR("Spl.Info", 20)
    mHeader = mHeader + 1
    Print #1, PSTR("#", 5) & PSTR("Prefix", 10) & PSTR("Inv-No", 20) & PSTR("No.", 20) & PSTR("Name Of City", 35) & PSTR("Engine No", 15) & PSTR("Date", 15) & Space(15) & PSTR("Financer Name", 35) & PSTR("Fin-Amt", 15) & PSTR("Type", 5) & Space(20) & mChr18
    mHeader = mHeader + 1
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    RstRep.MoveFirst
    mHeader = 1
    While Not RstRep.EOF = True
        If Counter <= mRec Then
            Counter = Counter + 1
            mCounter = mCounter + 1
            Print #1, mChr17 & PSTR(STR(mCounter), 5) & PSTR(mID(RstRep!Inv_DocId, 9, 5), 10) & PSTR("'" & RstRep!Inv_Date & "'", 20) & PSTR("'" & RstRep!Ord_Date & "'", 20) & PSTR(RstRep!Name, 33) & Space(2) & PSTR(RstRep!Model, 20) & PSTR("'" & RstRep!PBILL_NO & "'", 15) & PSTR(IIf(RstRep!Net_Amount = 0, "", STR(RstRep!Net_Amount)), 15) & PSTR(RstRep!FinGrpName, 35) & PSTR(IIf(RstRep!Fin_Amt = 0, "", STR(RstRep!Fin_Amt)), 15) & PSTR(RstRep!Form_Code, 5) & PSTR(RstRep!MISC_INFO, 20)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(10) & PSTR(PrinID(RstRep!Inv_DocId), 20) & PSTR(PrinID(RstRep!OrdDocId), 20) & PSTR(RstRep!Add1, 33) & Space(2) & PSTR(RstRep!ChassisNo, 20) & PSTR("'" & RstRep!PBILL_DATE & "'", 15) & Space(15) & PSTR(RstRep!FinName, 35) & Space(15) & Space(5) & Space(20)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(10) & Space(20) & Space(20) & PSTR(RstRep!Add2, 33) & Space(2) & PSTR(RstRep!EngineNo, 20) & Space(15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(20)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(10) & Space(20) & Space(20) & PSTR(RstRep!Add3, 33) & Space(2) & Space(15) & Space(15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(20)
            mHeader = mHeader + 1
            Print #1, Space(5) & Space(10) & Space(20) & Space(20) & PSTR(RstRep!CityName, 33) & Space(2) & Space(15) & Space(15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(20) & mChr18
            mHeader = mHeader + 1
            TotalNetAmt = TotalNetAmt + Val(RstRep!Net_Amount)
            If Counter = mRec Then isLast = True
            RstRep.MoveNext
        Else
            If isLast Then
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = 0: Counter = 0
                isLast = False
                Print #1, Space(PageWidth / 2) & "Page :" & PageNo + 1
                PageNo = PageNo + 1
                Print #1, mEject
                Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
                mHeader = mHeader + 1
                Print #1, PRN_TIT("Vehicle Sale Register", "C", PageWidth)
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
                Print #1, mChr17 & Space(5) & Space(10) & PSTR("Sales", 20) & PSTR("Order", 20) & PSTR("Name of Customer", 35) & PSTR("Model", 15) & PSTR("Telco", 15) & Space(15) & Space(35) & Space(15) & Space(5) & Space(20)
                mHeader = mHeader + 1
                Print #1, PSTR("#", 5) & PSTR("Invoice", 10) & PSTR("Date", 20) & PSTR("Date", 20) & PSTR("Address", 35) & PSTR("Chassis-No", 15) & PSTR("Bill No", 15) & PSTR("Sale-Amt", 15) & PSTR("Financer Group", 35) & Space(15) & PSTR("Form", 5) & PSTR("Spl.Info", 20)
                mHeader = mHeader + 1
                Print #1, PSTR("#", 5) & PSTR("Prefix", 10) & PSTR("Inv-No", 20) & PSTR("No.", 20) & PSTR("Name Of City", 35) & PSTR("Engine No", 15) & PSTR("Date", 15) & Space(15) & PSTR("Financer Name", 35) & PSTR("Fin-Amt", 15) & PSTR("Type", 5) & Space(20) & mChr18
                mHeader = mHeader + 1
                Print #1, Replace(Space(PageWidth), " ", "-")
                mHeader = mHeader + 1
            End If
        End If
    
    Wend
    Print #1, Replace(Space(PageWidth), " ", "-")
    mHeader = mHeader + 1
    Print #1, Space(5) & Space(10) & Space(20) & Space(20) & Space(33) & Space(2) & Space(20) & Space(15) & Space(15) & Space(35) & PSTR("Total --- >", 20) & PSTR(IIf(TotalNetAmt = 0, "", STR(TotalNetAmt)), 20, , AlignRight)
    Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        'Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''        Print #1, "Type C:\RepPrint.Txt > Prn"
''    Else
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
''    End If
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
    End If
End Sub

Private Sub PurSaleTaxSummProc()
On Error GoTo ELoop
Dim RsTemp As ADODB.Recordset
Dim mQry As String, Condstr$
Dim sQryParty$, sQryTaxableAmt$, sQryTaxAmt$
Dim mPartyField$, mTaxableField$, mTaxField$
Dim mVehPurGroupCode$, mVehSaleGroupCode$, mSprPurGroupCode$, mSprSaleGroupCode$, mVatGroupCode$, mServiceTaxGroupCode$
Dim mVehCashAc$, mSprCashAc$, mWsCashAc$, mServiceTaxAc$, mLabourAc$
Const PubDebtorMainGroupCode = "060004"
Const PubCreditorMainGroupCode = "030003"
Dim mPartyGroup$

    Set RsTemp = G_FaCn.Execute("Select GroupCode From AcGroup Where Left(MainGrCode,6) In ('" & PubDebtorMainGroupCode & "', '" & PubCreditorMainGroupCode & "')")
    If RsTemp.RecordCount > 0 Then
        mPartyGroup = ""
        Do Until RsTemp.EOF
            mPartyGroup = mPartyGroup & "'" & XNull(RsTemp(0)) & "', "
            RsTemp.MoveNext
        Loop
        mPartyGroup = left(mPartyGroup, Len(mPartyGroup) - 2)
    End If


    Set RsTemp = GCn.Execute("Select VehPurGroupCode, VehSaleGroupCode, SprPurGroupCode, " & _
                            "SprSaleGroupCode, VatGroupCode, ServiceTaxGroupCode, SprCashAc, " & _
                            "VehCashAc, WsCashAc, ServTaxAc, LabourAc From DmsEnviro")
    With RsTemp
        If .RecordCount > 0 Then
            mVehPurGroupCode = XNull(!VehPurGroupCode)
            mVehSaleGroupCode = XNull(!VehSaleGroupCode)
            mSprPurGroupCode = XNull(!SprPurGroupCode)
            mSprSaleGroupCode = XNull(!SprSaleGroupCode)
            mVatGroupCode = XNull(!VatGroupCode)
            mServiceTaxGroupCode = XNull(!ServiceTaxGroupCode)
            mSprCashAc = XNull(!SprCashAc)
            mVehCashAc = XNull(!VehCashAc)
            mWsCashAc = XNull(!WsCashAc)
            mServiceTaxAc = XNull(!ServTaxAc)
            mLabourAc = XNull(!LabourAc)
        End If
    End With
    
    
    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 0)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(List1, FGrid.TextMatrix(List1, 0)) = False Then RepPrint = False: Exit Sub
    
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub

    Condstr = " Where Lm.V_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Lm.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    
    
    If UCase(Trim(FGrid.TextMatrix(List1, 1))) = "SPARE PURCHASE TAX SUMMARY" Then
        mPartyField = "Ld.AmtCr"
        mTaxField = "Ld.AmtDr"
        
        sQryTaxableAmt = "Select Sum(Ld.AmtDr) " & _
                         "From (Ledger Ld " & _
                         "Left Join SubGroup S On Ld.SubCode=S.SubCode) " & _
                         "Where Ld.AmtDr>0 And Ld.DocId=Lm.DocId And S.GroupCode='" & mSprPurGroupCode & "'"


        Condstr = Condstr & " And Lm.V_type In ('" & PubDmsVTypeSprPurCredit & "', '" & PubDmsVTypeSprPurCash & "') "
    ElseIf UCase(Trim(FGrid.TextMatrix(List1, 1))) = "SPARE SALE TAX SUMMARY" Or StrCmp(FGrid.TextMatrix(List1, 1), "Spare Sale & Labour Summary") Then
        mPartyField = "Ld.AmtDr"
        mTaxField = "Ld.AmtCr"
        
        sQryTaxableAmt = "Select Sum(Ld.AmtCr) " & _
                         "From (Ledger Ld " & _
                         "Left Join SubGroup S On Ld.SubCode=S.SubCode) " & _
                         "Where Ld.AmtCr>0 And Ld.DocId=Lm.DocId And S.GroupCode='" & mSprSaleGroupCode & "'"

    
        Condstr = Condstr & " And Lm.V_type In ('" & PubDmsVTypeSprSaleCredit & "', '" & PubDmsVTypeSprSaleCash & "', '" & PubDmsVTypeWorkshopSaleCash & "', '" & PubDmsVTypeWorkshopSaleCredit & "') "
    ElseIf UCase(Trim(FGrid.TextMatrix(List1, 1))) = "VEHICLE SALE TAX SUMMARY" Then
        mPartyField = "Ld.AmtDr"
        mTaxField = "Ld.AmtCr"
        
        sQryTaxableAmt = "Select Sum(Ld.AmtCr) " & _
                         "From (Ledger Ld " & _
                         "Left Join SubGroup S On Ld.SubCode=S.SubCode) " & _
                         "Where Ld.AmtCr>0 And Ld.DocId=Lm.DocId And S.GroupCode='" & mVehSaleGroupCode & "'"

    
        Condstr = Condstr & " And Lm.V_type In ('" & PubDmsVTypeVehSale & "') "
    ElseIf UCase(Trim(FGrid.TextMatrix(List1, 1))) = "VEHICLE PURCHASE TAX SUMMARY" Then
        mPartyField = "Ld.AmtCr"
        mTaxField = "Ld.AmtDr"
        
        sQryTaxableAmt = "Select Sum(Ld.AmtDr) " & _
                         "From (Ledger Ld " & _
                         "Left Join SubGroup S On Ld.SubCode=S.SubCode) " & _
                         "Where Ld.AmtDr>0 And Ld.DocId=Lm.DocId And S.GroupCode='" & mVehPurGroupCode & "'"

    
        Condstr = Condstr & " And Lm.V_type In ('" & PubDmsVTypeVehPur & "') "
    End If
    

''    sQryParty = "Select Max(Name) As PartyName, Ld.DocId " & _
''                "From ((Ledger Ld " & _
''                "Left Join SubGroup S On S.SubCode=Ld.SubCode) " & _
''                "Left Join AcGroup A On A.GroupCode=S.GroupCode) " & _
''                "Where (Left(A.MainGrCode,6) in ('" & PubDebtorMainGroupCode & "', '" & PubCreditorMainGroupCode & "') Or Ld.SubCode In ('" & mVehCashAc & "', '" & mSprCashAc & "', '" & mWsCashAc & "') ) " & _
''                "And " & mPartyField & " > 0 Group By Ld.DocId"


    If UCase(Trim(FGrid.TextMatrix(List1, 1))) = "SPARE SALE TAX SUMMARY" Or StrCmp(FGrid.TextMatrix(List1, 1), "Spare Sale & Labour Summary") Then
        sQryParty = "Select Case When (Max(LD.SubCode)<>'" & mWsCashAc & "' And Max(LD.SubCode) <> '" & mVehCashAc & "' And Max(LD.SubCode) <> '" & mSprCashAc & "') Or IsNull(Max(LM.DmsSubCode),'')='' Then Max(S.Name) Else Max(Dsg.Name) End As PartyName, Ld.DocId " & _
                    "From (((Ledger Ld " & _
                    "Left Join LedgerM Lm on Ld.DocID = LM.DocID) " & _
                    "Left Join DmsSubGroup Dsg On Lm.DmsSubCode = Dsg.DmsSubCode) " & _
                    "Left Join SubGroup S On S.SubCode=Ld.SubCode) " & _
                    "Where (S.GroupCode in (" & mPartyGroup & ") Or Ld.SubCode In ('" & mVehCashAc & "', '" & mSprCashAc & "', '" & mWsCashAc & "') ) " & _
                    "And " & mPartyField & " > 0 Group By Ld.DocId"
    Else
        sQryParty = "Select Max(Name) As PartyName, Ld.DocId " & _
                    "From (Ledger Ld " & _
                    "Left Join SubGroup S On S.SubCode=Ld.SubCode) " & _
                    "Where (S.GroupCode in (" & mPartyGroup & ") Or Ld.SubCode In ('" & mVehCashAc & "', '" & mSprCashAc & "', '" & mWsCashAc & "') ) " & _
                    "And " & mPartyField & " > 0 Group By Ld.DocId"
    End If

    sQryTaxAmt = "Select Sum(" & mTaxField & ") " & _
                "From (Ledger Ld " & _
                "Left  Join SubGroup S On Ld.SubCode=S.SubCode) " & _
                "Where " & mTaxField & ">0 And Ld.DocId=Lm.DocId And S.GroupCode='" & mVatGroupCode & "' and Ld.SubCode Not In ('" & mServiceTaxAc & "') "


    If UCase(Trim(FGrid.TextMatrix(List1, 1))) = "SERVICE TAX REGISTER" Then
        Condstr = Condstr & " And " & cMID("Lm.DocId", "4", "5") & " In ('" & PubDmsVTypeWorkshopSaleCash & "', '" & PubDmsVTypeWorkshopSaleCredit & "') "
        
        sQryTaxableAmt = "Select Sum(Ld.AmtCr) " & _
                         "From (Ledger Ld " & _
                         "Left Join SubGroup S On Ld.SubCode=S.SubCode) " & _
                         "Where Ld.AmtCr>0 And Ld.DocId=Lm.DocId And Ld.subCode in ('" & mLabourAc & "')"
        
        
        sQryParty = "Select Max(Name) As PartyName, Ld.DocId " & _
                    "From ((Ledger Ld " & _
                    "Left Join SubGroup S On S.SubCode=Ld.SubCode) " & _
                    "Left Join AcGroup A On A.GroupCode=S.GroupCode) " & _
                    "Where (Left(A.MainGrCode,6) in ('" & PubDebtorMainGroupCode & "', '" & PubCreditorMainGroupCode & "') Or Ld.SubCode In ('" & mVehCashAc & "', '" & mSprCashAc & "', '" & mWsCashAc & "') ) " & _
                    "And  Ld.AmtDr  > 0 Group By Ld.DocId"
    
        sQryTaxAmt = "Select Sum(Ld.AmtCr) " & _
                    "From (Ledger Ld " & _
                    "Left  Join SubGroup S On Ld.SubCode=S.SubCode) " & _
                    "Where Ld.AmtCr >0 And Ld.DocId=Lm.DocId And Ld.SubCode In ('" & mServiceTaxAc & "') "
        
    End If



    If StrCmp(FGrid.TextMatrix(List1, 1), "Spare Sale & Labour Summary") Then
        
        RepName = "SaleLabourTaxSumm"
        
        mQry = "Select Lm.DocId, Lm.V_Date, Lm.V_Type, Lm.V_No, Lm.DmsRefNo, P.PartyName , " & _
               "(" & sQryTaxableAmt & ") As TaxableAmt, (" & sQryTaxAmt & ") As TaxAmt, Vt.Description As Voucher_Type, 0 as TaxableLabour, 0 as TaxLabour " & _
               "From ((LedgerM Lm  " & _
               "Left Join (" & sQryParty & ") As P On P.DocId = Lm.DocId) " & _
               "Left Join Voucher_Type Vt On Lm.V_Type=Vt.V_Type) "
    
        Condstr = Condstr & " And (" & sQryTaxableAmt & ")>0 "
    
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Lm.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(Lm.DocId,1) in (" & GridString2 & ")"
    
        mQry = mQry + Condstr
    
    
        Condstr = " Where Lm.V_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and Lm.V_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
        Condstr = Condstr & " And " & cMID("Lm.DocId", "4", "5") & " In ('" & PubDmsVTypeWorkshopSaleCash & "', '" & PubDmsVTypeWorkshopSaleCredit & "') "
        
        sQryTaxableAmt = "Select Sum(Ld.AmtCr) " & _
                         "From (Ledger Ld " & _
                         "Left Join SubGroup S On Ld.SubCode=S.SubCode) " & _
                         "Where Ld.AmtCr>0 And Ld.DocId=Lm.DocId And Ld.subCode in ('" & mLabourAc & "')"
        
        
        sQryParty = "Select Max(Name) As PartyName, Ld.DocId " & _
                    "From ((Ledger Ld " & _
                    "Left Join SubGroup S On S.SubCode=Ld.SubCode) " & _
                    "Left Join AcGroup A On A.GroupCode=S.GroupCode) " & _
                    "Where (Left(A.MainGrCode,6) in ('" & PubDebtorMainGroupCode & "', '" & PubCreditorMainGroupCode & "') Or Ld.SubCode In ('" & mVehCashAc & "', '" & mSprCashAc & "', '" & mWsCashAc & "') ) " & _
                    "And  Ld.AmtDr  > 0 Group By Ld.DocId"
    
        sQryTaxAmt = "Select Sum(Ld.AmtCr) " & _
                    "From (Ledger Ld " & _
                    "Left  Join SubGroup S On Ld.SubCode=S.SubCode) " & _
                    "Where Ld.AmtCr >0 And Ld.DocId=Lm.DocId And Ld.SubCode In ('" & mServiceTaxAc & "') "
    
        
        mQry = mQry + " Union All Select Lm.DocId, Lm.V_Date, Lm.V_Type, Lm.V_No, Lm.DmsRefNo, P.PartyName , " & _
               "0 As TaxableAmt, 0 As TaxAmt, Vt.Description As Voucher_Type, (" & sQryTaxableAmt & ") as TaxableLabour, (" & sQryTaxAmt & ") as TaxLabour " & _
               "From ((LedgerM Lm  " & _
               "Left Join (" & sQryParty & ") As P On P.DocId = Lm.DocId) " & _
               "Left Join Voucher_Type Vt On Lm.V_Type=Vt.V_Type) "
    
        
        Condstr = Condstr & " And (" & sQryTaxableAmt & ")>0 "
    
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Lm.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(Lm.DocId,1) in (" & GridString2 & ")"
    
        mQry = mQry + Condstr
    
    
    
    
        mQry = "SELECT Max(DocID) AS DocID, Max(V_Date) AS V_Date, Max(V_Type) AS V_Type, Max(V_No) as V_No, DmsRefNo, Max(PartyName) AS PartyName, Sum(TaxableAmt) AS TaxableAmt, Sum(TaxAmt) AS TaxAmt, Max(Voucher_Type) AS Voucher_Type, Sum(TaxableLabour) AS TaxableLabour, IsNull(Sum(TaxLabour),0) AS TaxLabour " & _
               "From " & _
               " ( " & mQry & ") As X " & _
               "Group By X.DmsRefNo Order By X.V_Date, X.DocId"
           
    
    Else
        RepName = "PurSaleTaxSumm"

        mQry = "Select Lm.DocId, Lm.V_Date, Lm.V_Type, Lm.V_No, Lm.DmsRefNo, P.PartyName , " & _
               "(" & sQryTaxableAmt & ") As TaxableAmt, (" & sQryTaxAmt & ") As TaxAmt, Vt.Description As Voucher_Type " & _
               "From ((LedgerM Lm  " & _
               "Left Join (" & sQryParty & ") As P On P.DocId = Lm.DocId) " & _
               "Left Join Voucher_Type Vt On Lm.V_Type=Vt.V_Type) "
    
        Condstr = Condstr & " And (" & sQryTaxableAmt & ")>0 "
        If Check1(1).Value = Unchecked Then Condstr = Condstr & " and " & cMID("Lm.DocId", "3", "1") & " in (" & GridString1 & ")"
        If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(Lm.DocId,1) in (" & GridString2 & ")"
    
          
    
        mQry = mQry + Condstr + " Order By Lm.V_Date, Lm.DocId"
    End If

    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), G_FaCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    

    RepTitle = Trim(FGrid.TextMatrix(List1, 1))
    
    
Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub



Private Sub VehSumModelProc()
On Error GoTo ELoop
Dim mQry$, mQRY1$, mQRY2$, mQRY3$, mQry4$, Condstr As String, I As Integer, Rst As ADODB.Recordset
Dim mGrpStr$
Dim mMainGrpStr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
        
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    If Check1(5).Value = Unchecked Then GridString5 = FillString(GridRow5, 5, 1): If RepPrint = False Then Exit Sub
    If Check1(6).Value = Unchecked Then GridString6 = FillString(GridRow6, 6, 1): If RepPrint = False Then Exit Sub
    If Check1(7).Value = Unchecked Then GridString7 = FillString(GridRow7, 7, 1): If RepPrint = False Then Exit Sub
    If Check1(8).Value = Unchecked Then GridString8 = FillString(GridRow8, 8, 1): If RepPrint = False Then Exit Sub
    If Check1(9).Value = Unchecked Then GridString9 = FillString(GridRow9, 9, 1): If RepPrint = False Then Exit Sub

    'Condstr = " Where VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & " "
    Condstr = ""
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Left(VS.Pur_SiteCode,1) in (" & GridString1 & ")"
    
 If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and Left(VS.Pur_SiteCode,1)  ='" & PubSiteCode & "' "
    End If
    
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(VS.Pur_DocID,1) in (" & GridString2 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and M.Model in (" & GridString4 & ")"
    If Check1(5).Value = Unchecked Then Condstr = Condstr & " and M.Grp_Code in (" & GridString5 & ")"
    If Check1(6).Value = Unchecked Then Condstr = Condstr & " and M.Cat_Code in (" & GridString6 & ")"
    If Check1(7).Value = Unchecked Then Condstr = Condstr & " and VS.Godown in (" & GridString7 & ")"
'    If Check1(8).Value = Unchecked Then Condstr = Condstr & " and VO.FB_Code in (" & GridString8 & ")"
'    If Check1(9).Value = Unchecked Then Condstr = Condstr & " and VO.Rep_Code in (" & GridString9 & ")"
        
          
          
          
          
    If FGrid.TextMatrix(List2, 1) = "Yes" Then
        mGrpStr = "MG.ModelGrp_Name"
    Else
        mGrpStr = "VS.Model"
    End If
      
    If FGrid.TextMatrix(List3, 1) = "Site" Then
        mMainGrpStr = "Site.Site_Desc"
    Else
        mMainGrpStr = "MC.ModelCat_Name"
    End If
      
          
          
              
          
    mQry = "SELECT " & mMainGrpStr & " as Vehicle_Type,M.Model_Desc,M.Sales_Desc, " & mGrpStr & " As  Model, " & vIsNull("VP.Amount+VP.Addition-VP.Deduction+VP.Misc_Amt", "VS.Rate") & " AS VRATE, 0 as VRATE1, 0 as VRATE2, 1 as Opening, 0 as Purchase, 0 as LastPurch, 0 as Sale, 0 as LastSale, MG.ModelGrp_Name, 0 as Cancelled,0 as ClStk,0 as ClStkVal " & _
        " from (((((Veh_Stock as VS LEFT JOIN Model as M ON VS.MODEL = M.MODEL) " & _
        " Left Join Model_Grp MG on MG.ModelGrp_Code=M.Grp_Code) " & _
        " Left Join Model_Cat MC on MG.ModelCat_Code=MC.ModelCat_Code) " & _
        " left join Veh_Purch1 as VP on VS.PUR_DOCiD=VP.DOCID) " & _
        " Left Join Site On Site.Site_Code = Left(VS.Pur_SiteCode,1)) " & _
        " WHERE VS.Pur_VDate < " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & _
        " And (VS.Sal_VDate Is Null or VS.Sal_VDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & ") "
        
    mQRY1 = "SELECT " & mMainGrpStr & " as Vehicle_Type,M.Model_Desc,M.Sales_Desc, " & mGrpStr & " As  Model,0 as VRATE, " & vIsNull("VP.Amount+VP.Addition-VP.Deduction+VP.Misc_Amt", "VS.Rate") & " as VRATE1,0 as VRATE2,0 as Opening,1 as Purchase," & _
        " " & cIIF("VS.Pur_VDate = " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "", "1", "0") & " as LastPurch,0 as Sale,0 as LastSale,MG.ModelGrp_Name, 0 as Cancelled,0 as ClStk,0 as ClStkVal " & _
        " from (((((Veh_Stock as VS LEFT JOIN Model as M ON VS.MODEL = M.MODEL)  " & _
        " Left Join Model_Grp MG on MG.ModelGrp_Code=M.Grp_Code) " & _
        " Left Join Model_Cat MC on MG.ModelCat_Code=MC.ModelCat_Code) " & _
        " left join Veh_Purch1 as VP on VS.PUR_DOCiD=VP.DOCID) " & _
        " Left Join Site On Site.Site_Code = Left(VS.Pur_SiteCode,1)) " & _
        " WHERE (VS.Pur_VDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & _
        " And VS.Pur_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ")"

'    mQRY2 = "SELECT MC.ModelCat_Name as Vehicle_Type,M.Model_Desc,M.Sales_Desc, " & mGrpStr & " As  Model,0 as VRATE,0 as VRATE1,VO.Net_Amount-Vo.Tax_Amt as VRATE2,0 as Opening,0 as Purchase,0 as LastPurch,1 as Sale, " & _
'        " " & cIIF("VO.Inv_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "", "1", "0") & " as LastSale,MG.ModelGrp_Name, 0 as Cancelled,0 as ClStk,0 as ClStkVal " & _
'        " from ((((Veh_Stock as VS LEFT JOIN Model as M ON VS.MODEL = M.MODEL)  Left Join Model_Grp MG on MG.ModelGrp_Code=M.Grp_Code)  Left Join Model_Cat MC on MG.ModelCat_Code=MC.ModelCat_Code) " & _
'        " left join Veh_Order as VO on VS.Sal_DocID=VO.Inv_DocID) " & _
'        " WHERE (VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & _
'        "  And VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ")"

    mQRY2 = "SELECT " & mMainGrpStr & " as Vehicle_Type,M.Model_Desc,M.Sales_Desc, " & mGrpStr & " As  Model,0 as VRATE,0 as VRATE1, " & IIf(StrCmp(left(PubComp_Name, 4), "Enar"), vIsNull("VP.Amount+VP.Addition-VP.Deduction+VP.Misc_Amt", "VS.Rate"), "Vo.Net_Amount - Vo.Tax_Amt") & " as VRATE2,0 as Opening,0 as Purchase,0 as LastPurch,1 as Sale, " & _
        " " & cIIF("VO.Inv_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "", "1", "0") & " as LastSale,MG.ModelGrp_Name, 0 as Cancelled,0 as ClStk,0 as ClStkVal " & _
        " from ((((((Veh_Stock as VS LEFT JOIN Model as M ON VS.MODEL = M.MODEL)  Left Join Model_Grp MG on MG.ModelGrp_Code=M.Grp_Code)  Left Join Model_Cat MC on MG.ModelCat_Code=MC.ModelCat_Code) " & _
        " left join Veh_Order as VO on VS.Sal_DocID=VO.Inv_DocID) " & _
        " Left Join Veh_Purch1 VP On VS.Pur_DocId=VP.DocId) " & _
        " Left Join Site On Site.Site_Code = Left(VS.Pur_SiteCode,1)) " & _
        " WHERE (VO.Inv_Date  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & _
        "  And VO.Inv_Date <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ")"

    mQRY3 = "SELECT " & mMainGrpStr & " as Vehicle_Type,M.Model_Desc,M.Sales_Desc, " & mGrpStr & " As  Model,0 as VRATE,0 as VRATE1,VO.Net_Amount-Vo.Tax_Amt as VRATE2,0 as Opening,0 as Purchase,0 as LastPurch,0 as Sale, " & _
        " " & cIIF("VO.Inv_Date = " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "", "1", "0") & " as LastSale,MG.ModelGrp_Name, 1 As Cancelled,0 as ClStk,0 as ClStkVal " & _
        " from (((((Veh_Stock as VS LEFT JOIN Model as M ON VS.MODEL = M.MODEL) " & _
        " Left Join Model_Grp MG on MG.ModelGrp_Code=M.Grp_Code)  " & _
        " Left Join Model_Cat MC on MG.ModelCat_Code=MC.ModelCat_Code) " & _
        " left join Veh_Order1 as VO on VS.ChassisNo=VO.Chassis) " & _
        " Left Join Site On Site.Site_Code = Left(VS.Pur_SiteCode,1)) " & _
       " WHERE (VO.Inv_UEntDT  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & _
        "  And VO.Inv_UEntDT <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ")"
    
    mQry4 = "SELECT " & mMainGrpStr & " as Vehicle_Type,M.Model_Desc,M.Sales_Desc, " & mGrpStr & " As  Model,0 as VRATE,0 as VRATE1,0 as VRATE2,0 as Opening,0 as Purchase,0 as LastPurch,0 as Sale, " & _
        " 0 as LastSale,MG.ModelGrp_Name, 0 as Cancelled,1 as ClStk,VP.TOT_AMOUNT - VP.TAX_AMT as ClStkVal " & _
        " from (((((Veh_Stock as VS LEFT JOIN Model as M ON VS.MODEL = M.MODEL)  " & _
        " Left Join Model_Grp MG on MG.ModelGrp_Code=M.Grp_Code) " & _
        " Left Join Model_Cat MC on MG.ModelCat_Code=MC.ModelCat_Code) " & _
        " left join Veh_Purch1 as VP on VS.PUR_DOCiD=VP.DOCID) " & _
        " Left Join Site On Site.Site_Code = Left(VS.Pur_SiteCode,1)) " & _
        " WHERE VS.Pur_VDate  <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "" & _
        " And (VS.Sal_VDate is null Or Sal_VDate > " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ")"


    mQry = mQry & Condstr
    mQRY1 = mQRY1 & Condstr
    mQRY2 = mQRY2 & Condstr
    mQRY3 = mQRY3 & Condstr
    mQry4 = mQry4 & Condstr
    
        'Create temp table
        Set RstRep = New ADODB.Recordset
        With RstRep
            .Fields.Append "Vehicle_Type", adChar, 20, adFldIsNullable
            .Fields.Append "Model", adChar, 20, adFldIsNullable
            .Fields.Append "MonthOpen", adInteger, 4, adFldIsNullable
            .Fields.Append "MonthOpenVal", adInteger, 14, adFldIsNullable
            .Fields.Append "MonthPur", adInteger, 4, adFldIsNullable
            .Fields.Append "MonthPurVal", adInteger, 14, adFldIsNullable
            .Fields.Append "PurDay", adInteger, 4, adFldIsNullable
            .Fields.Append "MonthSal", adInteger, 4, adFldIsNullable
            .Fields.Append "MonthSalVal", adInteger, 14, adFldIsNullable
            .Fields.Append "SalDay", adInteger, 4, adFldIsNullable
            .Fields.Append "Model_Desc", adChar, 80, adFldIsNullable
            .Fields.Append "Sales_Desc", adChar, 40, adFldIsNullable
            .Fields.Append "MonthCancelled", adInteger, 4, adFldIsNullable
            .Fields.Append "ClStk", adInteger, 4, adFldIsNullable
            .Fields.Append "ClStkVal", adInteger, 14, adFldIsNullable
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open
        End With
        'temp table created
        Dim j As Integer
        For I = 1 To 5
            If I = 2 Then
                mQry = mQRY1
            ElseIf I = 3 Then
                mQry = mQRY2
            ElseIf I = 4 Then
                mQry = mQRY3
            ElseIf I = 5 Then
                mQry = mQry4
            End If
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
            For j = 1 To Rst.RecordCount
                With RstRep
                    .AddNew
'                    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
                        .Fields("Vehicle_Type") = Rst!Vehicle_Type
'                    Else
'                        .Fields("Vehicle_Type") = Rst!ModelGrp_Name
'                    End If
                    .Fields("Model") = Rst!Model  'Grp_Name        '' Model
                    .Fields("MonthOpen") = Rst!Opening
                    .Fields("MonthOpenVal") = Rst!vrate
                    .Fields("MonthPur") = Rst!Purchase
                    If I = 2 Then
                        .Fields("MonthPurVal") = Rst!VRATE1
                    Else
                        .Fields("MonthPurVal") = 0
                    End If
                    .Fields("PurDay") = Rst!LastPurch
                    .Fields("MonthSal") = Rst!Sale
                    .Fields("ClStk") = Rst!ClStk
                    If I = 3 Then
                        .Fields("MonthSalVal") = Rst!VRATE2
                    Else
                        .Fields("MonthSalVal") = 0
                    End If
                    If I = 5 Then
                        .Fields("ClStkVal") = Rst!ClStkVal
                    Else
                        .Fields("ClStkVal") = 0
                    End If
                    .Fields("SalDay") = Rst!LastSale
                    .Fields("Model_Desc") = XNull(Rst!Model_Desc)
                    .Fields("Sales_Desc") = XNull(Rst!Sales_Desc)
                    .Fields("MonthCancelled") = VNull(Rst!Cancelled)
                    .Update
                End With
                Rst.MoveNext
            Next
            Set Rst = Nothing
        Next
    'Set RstRep = New Recordset
    'RstRep.CursorLocation = adUseClient
    'RstRep.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    If FGrid.TextMatrix(List1, 1) = "Qty Wise" Then
        If UCase(left(PubComp_Name, 4)) <> "ENAR" And UCase(left(PubComp_Name, 6)) <> "J.M.A." Then
            RepName = "VehSumModelQty"
        Else
            RepName = "VehSumModelQty_enar"
        End If
        RepTitle = UCase(Me.CAPTION)
    ElseIf FGrid.TextMatrix(List1, 1) = "Value Wise" Then
        RepName = "VehSumModelVal"
        RepTitle = UCase(Me.CAPTION)
    ElseIf FGrid.TextMatrix(List1, 1) = "Both" Then
        RepName = "VehSumModel"
        RepTitle = UCase(Me.CAPTION)
    End If
    
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub



Private Sub VehOfftakeNRetailProc()
On Error GoTo ELoop
Dim mQry$, mQRY1$, mQRY2$, mQRY3$, mQry4$, Condstr As String, I As Integer, Rst As ADODB.Recordset
Dim mGrpStr$, CondStr1$
Dim mMainGrpStr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    
        
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    If Check1(5).Value = Unchecked Then GridString5 = FillString(GridRow5, 5, 1): If RepPrint = False Then Exit Sub
    If Check1(6).Value = Unchecked Then GridString6 = FillString(GridRow6, 6, 1): If RepPrint = False Then Exit Sub
    If Check1(7).Value = Unchecked Then GridString7 = FillString(GridRow7, 7, 1): If RepPrint = False Then Exit Sub
    
    Condstr = "": CondStr1 = ""
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Left(S.Pur_SiteCode,1) in (" & GridString1 & ")"
    
    If Check1(1).Value = Checked Then
        If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and Left(S.Pur_SiteCode,1)  ='" & PubSiteCode & "' "
    End If
    
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(S.Pur_DocID,1) in (" & GridString2 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and M.Model in (" & GridString4 & ")"
    If Check1(4).Value = Unchecked Then CondStr1 = CondStr1 & " and M.Model in (" & GridString4 & ")"
    If Check1(5).Value = Unchecked Then Condstr = Condstr & " and M.Grp_Code in (" & GridString5 & ")"
    If Check1(5).Value = Unchecked Then CondStr1 = CondStr1 & " and M.Grp_Code in (" & GridString5 & ")"
    If Check1(6).Value = Unchecked Then Condstr = Condstr & " and M.Cat_Code in (" & GridString6 & ")"
    If Check1(6).Value = Unchecked Then Condstr = CondStr1 & " and M.Cat_Code1 in (" & GridString6 & ")"
    If Check1(7).Value = Unchecked Then Condstr = Condstr & " and S.Godown in (" & GridString7 & ")"
        
                    
          
          
      
    Dim bMonthStartDate          As Date
    Dim bMonthEndDate  As Date
    
    
    bMonthEndDate = RetDate(FGrid.TextMatrix(Date1, 1))
    bMonthStartDate = RetDate("01/" & Right(bMonthEndDate, 8))
              
              
              
              

    mQry = "SELECT m.MODEL, M.Grp_Code, Mg.ModelGrp_Name, M.Cat_Code, Mc.ModelCat_Name, " & _
           "IsNull(YrOp.Qty,0) AS YearOpening, IsNull(MthOp.Qty,0) AS MonthOpening, " & _
           "IsNull(YrPur.Qty,0) AS YrPur, IsNull(MthPur.Qty,0) AS MthPur, " & _
           "IsNull(YrPur.Qty,0) + IsNull(MthPur.Qty,0) AS TotalPur, IsNull(YrSal.Qty,0) AS YrSal, " & _
           "isNull(MthSal.Qty,0) AS MthSal,  IsNull(YrSal.Qty,0) + isNull(MthSal.Qty,0) AS Total_Sal, " & _
           "IsNull(MthOp.Qty,0) + IsNull(MthPur.Qty,0)-IsNull(MthSal.Qty,0) AS MthClosing, '" & bMonthStartDate & "' as  MonthStartDate,Mc.OldCode " & _
        "FROM Model M " & _
        "LEFT JOIN Model_Grp  Mg ON M.Grp_Code =mg.Modelgrp_Code " & _
        "LEFT JOIN Model_Cat  Mc ON M.Cat_Code =mc.ModelCat_Code " & _
        "Full JOIN (SELECT S.Model, Count(*) AS Qty FROM Veh_Stock S LEFT JOIN Model M ON M.MODEL = S.MODEL  WHERE S.Pur_Vdate<'" & PubStartDate & "' And (S.Sal_Vdate IS NULL Or S.Sal_VDate >='" & PubStartDate & "') " & Condstr & "  GROUP BY S.MODEL) AS YrOp ON M.MODEL = YrOp.Model " & _
        "Full JOIN (SELECT S.Model, Count(*) AS Qty FROM Veh_Stock S LEFT JOIN Model M ON M.MODEL = S.MODEL  WHERE S.Pur_Vdate<'" & bMonthStartDate & "' And (S.Sal_Vdate IS NULL Or S.Sal_VDate >='" & bMonthStartDate & "') " & Condstr & "  GROUP BY S.MODEL) AS MthOp ON M.MODEL = MthOp.Model " & _
        "Full JOIN (SELECT S.Model, Count(*) AS Qty FROM Veh_Stock S LEFT JOIN Model M ON M.MODEL = S.MODEL  WHERE S.Pur_Vdate>='" & PubStartDate & "' And S.Pur_VDate <'" & bMonthStartDate & "' " & Condstr & "  GROUP BY S.MODEL) AS YrPur ON M.MODEL = YrPur.Model " & _
        "Full JOIN (SELECT S.Model, Count(*) AS Qty FROM Veh_Stock S LEFT JOIN Model M ON M.MODEL = S.MODEL  WHERE S.Pur_Vdate>='" & bMonthStartDate & "' And S.Pur_VDate <='" & bMonthEndDate & "' " & Condstr & "  GROUP BY S.MODEL) AS MthPur ON M.MODEL = MthPur.Model " & _
        "Full JOIN (SELECT S.Model, Count(*) AS Qty FROM Veh_Stock S LEFT JOIN Model M ON M.MODEL = S.MODEL  WHERE S.Sal_Vdate>='" & PubStartDate & "' And S.Sal_VDate <'" & bMonthStartDate & "' " & Condstr & "  GROUP BY S.MODEL) AS YrSal ON M.MODEL = YrSal.Model " & _
        "Full JOIN (SELECT S.Model, Count(*) AS Qty FROM Veh_Stock S LEFT JOIN Model M ON M.MODEL = S.MODEL  WHERE S.Sal_Vdate>='" & bMonthStartDate & "' And S.Sal_VDate <='" & bMonthEndDate & "' " & Condstr & "  GROUP BY S.MODEL) AS MthSal ON M.MODEL = MthSal.Model " & _
        "WHERE YrOp.Qty Is Not Null " & _
        "OR MthOp.Qty IS NOT NULL " & _
        "OR  YrPur.Qty IS NOT NULL " & _
        "OR MthPur.Qty IS NOT NULL " & _
        "OR  YrSal.Qty IS NOT NULL " & _
        "OR MthSal.Qty IS NOT NULL " & CondStr1
    
              
              
              
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    
    RepName = "ModelWiseOfftakeAndSales"
    RepTitle = UCase(Me.CAPTION)
    
Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub




Private Sub ProcBodyBuilderChassis()
On Error GoTo ELoop
Dim mQry$, mQRY1$, mQRY2$, mQRY3$, mQry4$, Condstr As String, I As Integer, Rst As ADODB.Recordset
Dim mGrpStr$
Dim mMainGrpStr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
        
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    If Check1(5).Value = Unchecked Then GridString5 = FillString(GridRow5, 5, 1): If RepPrint = False Then Exit Sub
    If Check1(6).Value = Unchecked Then GridString6 = FillString(GridRow6, 6, 1): If RepPrint = False Then Exit Sub
    If Check1(7).Value = Unchecked Then GridString7 = FillString(GridRow7, 7, 1): If RepPrint = False Then Exit Sub
    If Check1(8).Value = Unchecked Then GridString8 = FillString(GridRow8, 8, 1): If RepPrint = False Then Exit Sub
    If Check1(9).Value = Unchecked Then GridString9 = FillString(GridRow9, 9, 1): If RepPrint = False Then Exit Sub

    Condstr = ""
    Condstr = " Where S.BodyBuilder_IssDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and S.BodyBuilder_IssDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  "
    
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Left(S.Pur_SiteCode,1) in (" & GridString1 & ")"
        If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and Left(S.Pur_SiteCode,1)  ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(S.Pur_DocID,1) in (" & GridString2 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and M.Model in (" & GridString4 & ")"
    If Check1(5).Value = Unchecked Then Condstr = Condstr & " and M.Grp_Code in (" & GridString5 & ")"
    If Check1(6).Value = Unchecked Then Condstr = Condstr & " and M.Cat_Code in (" & GridString6 & ")"
    If Check1(7).Value = Unchecked Then Condstr = Condstr & " and S.Godown in (" & GridString7 & ")"
    If Check1(8).Value = Unchecked Then Condstr = Condstr & " and S.BodyBuilder in (" & GridString8 & ")"
    If Check1(9).Value = Unchecked Then Condstr = Condstr & " and S.BodyBuilder_BodyType in (" & GridString9 & ")"
        
          
          
      
    mQry = "Select B.BodyBuilderDesc, Bt.BodyTypeDesc, S.ChassisNo, S.EngineNo, S.Model,  M.Sales_Desc As Model_SalesDesc, " & _
         "S.BodyBuilder_Remark, S.BodyBuilder_IssDate, S.BodyBuilder_RecDate " & _
         "From (((((BodyBuilder B " & _
         "Left Join Veh_Stock S On B.BodyBuilderCode=S.BodyBuilder) " & _
         "Left Join Model M On S.Model=M.Model) " & _
         "Left Join Model_Grp Mg On M.Grp_Code=Mg.ModelGrp_Code) " & _
         "Left Join Model_Cat Mc On M.Cat_Code=Mc.ModelCat_Code) " & _
         "Left Join BodyType Bt On S.BodyBuilder_BodyType = Bt.BodyTypeCode) "
    mQry = mQry & Condstr
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    

    RepName = "BodyBuilderWiseChassis"
    RepTitle = UCase(Me.CAPTION)
    
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub




Private Sub ProcStockAtBodyBuilder()
On Error GoTo ELoop
Dim mQry$, mQRY1$, mQRY2$, mQRY3$, mQry4$, Condstr As String, I As Integer, Rst As ADODB.Recordset
Dim mGrpStr$
Dim mMainGrpStr$

    If IsNotBlank(Date1, FGrid.TextMatrix(Date1, 1)) = False Then RepPrint = False: Exit Sub
    If IsNotBlank(Date2, FGrid.TextMatrix(Date2, 1)) = False Then RepPrint = False: Exit Sub
        
        
    If Check1(1).Value = Unchecked Then GridString1 = FillString(GridRow1, 1, 1): If RepPrint = False Then Exit Sub
    If Check1(2).Value = Unchecked Then GridString2 = FillString(GridRow2, 2, 1): If RepPrint = False Then Exit Sub
    If Check1(4).Value = Unchecked Then GridString4 = FillString(GridRow4, 4, 1): If RepPrint = False Then Exit Sub
    If Check1(5).Value = Unchecked Then GridString5 = FillString(GridRow5, 5, 1): If RepPrint = False Then Exit Sub
    If Check1(6).Value = Unchecked Then GridString6 = FillString(GridRow6, 6, 1): If RepPrint = False Then Exit Sub
    If Check1(7).Value = Unchecked Then GridString7 = FillString(GridRow7, 7, 1): If RepPrint = False Then Exit Sub
    If Check1(8).Value = Unchecked Then GridString8 = FillString(GridRow8, 8, 1): If RepPrint = False Then Exit Sub
    If Check1(9).Value = Unchecked Then GridString9 = FillString(GridRow9, 9, 1): If RepPrint = False Then Exit Sub

    Condstr = ""
    'Condstr = " Where S.BodyBuilder_IssDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and S.BodyBuilder_IssDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & "  "
    Condstr = " Where ((S.Pur_VDate  >= " & ConvertDate(Format(FGrid.TextMatrix(Date1, 1), "dd/MMM/yyyy")) & " and S.Pur_VDate <= " & ConvertDate(Format(FGrid.TextMatrix(Date2, 1), "dd/MMM/yyyy")) & ") Or S.Pur_VType='V_OST' ) "
    
    
    If Check1(1).Value = Unchecked Then Condstr = Condstr & " and Left(S.Pur_SiteCode,1) in (" & GridString1 & ")"
    If Check1(1).Value = Checked Then
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then Condstr = Condstr & " and Left(S.Pur_SiteCode,1) ='" & PubSiteCode & "' "
    End If
    
    If Check1(2).Value = Unchecked Then Condstr = Condstr & " and left(S.Pur_DocID,1) in (" & GridString2 & ")"
    If Check1(4).Value = Unchecked Then Condstr = Condstr & " and M.Model in (" & GridString4 & ")"
    If Check1(5).Value = Unchecked Then Condstr = Condstr & " and M.Grp_Code in (" & GridString5 & ")"
    If Check1(6).Value = Unchecked Then Condstr = Condstr & " and M.Cat_Code in (" & GridString6 & ")"
    If Check1(7).Value = Unchecked Then Condstr = Condstr & " and S.Godown in (" & GridString7 & ")"
    If Check1(8).Value = Unchecked Then Condstr = Condstr & " and S.BodyBuilder in (" & GridString8 & ")"
    If Check1(9).Value = Unchecked Then Condstr = Condstr & " and S.BodyBuilder_BodyType in (" & GridString9 & ")"
        
    mQry = "Select B.BodyBuilderDesc, Bt.BodyTypeDesc, S.ChassisNo, S.EngineNo, S.Model,  M.Sales_Desc As Model_SalesDesc, " & _
         "S.BodyBuilder_Remark, S.BodyBuilder_IssDate, S.BodyBuilder_RecDate, 1 As Pur_Qty, " & _
         "" & cIIF("S.BodyBuilder Is Not Null And S.BodyBuilder<>''", "1", "0") & "  As Build_Qty, " & cIIF("S.Sal_VDate Is Not Null", "1", "0") & "  As Sold_Qty " & _
         "From (((((Veh_Stock S " & _
         "Left Join BodyBuilder B On B.BodyBuilderCode=S.BodyBuilder) " & _
         "Left Join Model M On S.Model=M.Model) " & _
         "Left Join Model_Grp Mg On M.Grp_Code=Mg.ModelGrp_Code) " & _
         "Left Join Model_Cat Mc On M.Cat_Code=Mc.ModelCat_Code) " & _
         "Left Join BodyType Bt On S.BodyBuilder_BodyType = Bt.BodyTypeCode) "
    mQry = mQry & Condstr
    
    mQry = "Select x.BodyBuilderDesc, x.BodyTypeDesc, x.Model, Max(x.Model_SalesDesc) As Model_SalesDesc, " & _
           "Sum(x.Pur_Qty) As Pur_Qty, Sum(x.Build_Qty) As Build_Qty, Sum(x.Sold_Qty) As Sold_Qty " & _
           "From (" & mQry & ") As x Group By x.Model, x.BodyBuilderDesc, x.BodyTypeDesc " & _
           "Order By x.Model"
    
    Set RstRep = New Recordset
    RstRep.CursorLocation = adUseClient
    RstRep.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If RstRep.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: RepPrint = False: Exit Sub
    

    RepName = "BodyBuilderWiseStock"
    RepTitle = UCase(Me.CAPTION)
    
    Exit Sub
ELoop:
    RepPrint = False
    MsgBox err.Description
End Sub



