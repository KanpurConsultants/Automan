VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmJobObserAction 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Supervisor Observation Entry"
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
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   8715
      MaxLength       =   25
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   1455
      Width           =   1470
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
      Left            =   4125
      MaxLength       =   40
      TabIndex        =   21
      Top             =   6525
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   20
      Left            =   6000
      MaxLength       =   40
      TabIndex        =   20
      Top             =   4725
      Width           =   5745
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   900
      Index           =   18
      Left            =   165
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   3705
      Width           =   5790
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   19
      Left            =   6000
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   3705
      Width           =   5790
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   17
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   17
      Top             =   3030
      Width           =   5070
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   4575
      MaxLength       =   25
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   13
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   13
      Top             =   2520
      Width           =   5070
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   225
      Index           =   9
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   9
      Top             =   1500
      Width           =   5070
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   225
      Index           =   10
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   10
      Top             =   1755
      Width           =   5070
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   225
      Index           =   11
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   11
      Top             =   2010
      Width           =   5070
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   225
      Index           =   12
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   12
      Top             =   2265
      Width           =   5070
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   225
      Index           =   3
      Left            =   1560
      MaxLength       =   14
      TabIndex        =   3
      Top             =   735
      Width           =   1590
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   225
      Index           =   1
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "Help"
      Top             =   480
      Width           =   1590
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   4575
      MaxLength       =   20
      TabIndex        =   4
      Top             =   735
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   16
      Left            =   4980
      MaxLength       =   10
      TabIndex        =   16
      Top             =   2775
      Width           =   1650
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   15
      Left            =   3210
      MaxLength       =   25
      TabIndex        =   15
      Top             =   2775
      Width           =   1395
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   225
      Index           =   7
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1245
      Width           =   1590
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   4575
      MaxLength       =   25
      TabIndex        =   6
      Top             =   990
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   4575
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1245
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   14
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   14
      Top             =   2775
      Width           =   1260
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
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
      Height          =   225
      Index           =   5
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   5
      Top             =   990
      Width           =   1590
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   1950
      Left            =   180
      TabIndex        =   22
      Top             =   5040
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   3440
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      BackColorFixed  =   13623520
      ForeColorFixed  =   8388736
      BackColorSel    =   15595518
      ForeColorSel    =   8388608
      BackColorBkg    =   13623520
      GridColor       =   0
      GridColorFixed  =   33023
      FocusRect       =   0
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "MW"
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Close Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Index           =   7
      Left            =   7275
      TabIndex        =   50
      Top             =   1470
      Width           =   1425
   End
   Begin VB.Label lblDocCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard DocID :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   7080
      TabIndex        =   48
      Top             =   765
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observation by Technical Supervisor :"
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
      Index           =   14
      Left            =   180
      TabIndex        =   47
      Top             =   4740
      Width           =   3210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Person Name"
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
      Height          =   225
      Index           =   2
      Left            =   4770
      TabIndex        =   46
      Top             =   4740
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
      Height          =   255
      Index           =   7
      Left            =   5925
      TabIndex        =   45
      Top             =   4725
      Width           =   45
   End
   Begin VB.Line Line2 
      X1              =   180
      X2              =   11775
      Y1              =   4665
      Y2              =   4665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervisor Observation :"
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
      Index           =   13
      Left            =   165
      TabIndex        =   44
      Top             =   3465
      Width           =   2100
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Action Taken :"
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
      Index           =   11
      Left            =   6000
      TabIndex        =   43
      Top             =   3465
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Index           =   9
      Left            =   165
      TabIndex        =   42
      Top             =   3045
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JC Open Dt.*"
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
      Left            =   3405
      TabIndex        =   41
      Top             =   495
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Index           =   26
      Left            =   165
      TabIndex        =   40
      Top             =   1755
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
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
      Index           =   31
      Left            =   165
      TabIndex        =   39
      Top             =   2775
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name"
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
      Index           =   39
      Left            =   165
      TabIndex        =   38
      Top             =   1500
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Index           =   10
      Left            =   165
      TabIndex        =   37
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division            :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7080
      TabIndex        =   36
      Top             =   510
      Width           =   1470
   End
   Begin VB.Label lblDocId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DocID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8760
      TabIndex        =   35
      Top             =   765
      Width           =   540
   End
   Begin VB.Line Line1 
      X1              =   165
      X2              =   11760
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard No.*"
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
      Index           =   12
      Left            =   165
      TabIndex        =   34
      Top             =   495
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No."
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
      Index           =   8
      Left            =   3405
      TabIndex        =   33
      Top             =   735
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(M)"
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
      Left            =   4635
      TabIndex        =   32
      Top             =   2775
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(R)"
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
      Left            =   2880
      TabIndex        =   31
      Top             =   2775
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(O)"
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
      Left            =   1095
      TabIndex        =   30
      Top             =   2775
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Srl No."
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
      Left            =   165
      TabIndex        =   29
      Top             =   1245
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration No."
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
      Left            =   165
      TabIndex        =   28
      Top             =   735
      Width           =   1365
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      Height          =   615
      Left            =   6915
      Top             =   450
      Width           =   4830
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code      :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8910
      TabIndex        =   27
      Top             =   510
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Type"
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
      Index           =   37
      Left            =   3405
      TabIndex        =   26
      Top             =   1245
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
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
      Index           =   38
      Left            =   165
      TabIndex        =   25
      Top             =   990
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
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   8
      Left            =   4440
      TabIndex        =   24
      Top             =   1245
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No."
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
      Index           =   33
      Left            =   3405
      TabIndex        =   23
      Top             =   990
      Width           =   915
   End
End
Attribute VB_Name = "frmJobObserAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'&H00E8D8FE&
Dim TAddMode As Boolean
Dim ExitCtrl As Boolean
Dim GridKey As Integer

Dim ForSiteCode As String

Dim MyIndex As Byte
Dim Rst As ADODB.Recordset

Dim Master As ADODB.Recordset
Dim RsJob As ADODB.Recordset
Dim RsMech As ADODB.Recordset

'grid color scheme
Private Const BackColorSelEnter As String = &HF8D7FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

'Text Box (Form)
Private Const JobNo As Byte = 1
Private Const JobDt As Byte = 2
Private Const VehRegNo As Byte = 3
Private Const Chassis As Byte = 4
Private Const Model As Byte = 5
Private Const Engine As Byte = 6
Private Const VehSrlNo As Byte = 7
Private Const SrvType As Byte = 8
Private Const OwnerName As Byte = 9
Private Const Address1 As Byte = 10
Private Const Address2 As Byte = 11
Private Const Address3 As Byte = 12
Private Const City As Byte = 13
Private Const PhoneOff As Byte = 14
Private Const PhoneResi As Byte = 15
Private Const Mobile As Byte = 16
Private Const Remarks As Byte = 17

Private Const SObser As Byte = 18
Private Const SAction As Byte = 19
Private Const MechName As Byte = 20
Private Const JobClDt As Byte = 21

'Fgrid1 Columns
Private Const C_Desc As Byte = 1
Private Const C_Observ As Byte = 2
Private Const C_Make As Byte = 3

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
    
    WinSetting Me
    Ini_Grid
    TopCtrl1.Tag = PubUParam    '"*EDP"
    ForSiteCode = PubSiteCode
    Call BlankText
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
      Dim SiteCond As String
        SiteCond = " And  Job_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
        If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
            SiteCond = SiteCond & " and  " & cMID("Jc.DocId", "3", "1") & "='" & PubSiteCode & "'"
        End If
    
    If PubMoveRecYn Then
        Master.Open "select Jc.DocId AS SearchCode from Job_Card as JC where left(JC.DocID,1)='" & PubDivCode & "' and right(" & xIsNull("jc.DocId_InvSpr", "") & ",8) <> 'Cancelld' " & SiteCond & " order by JC.DocID", GCn, adOpenDynamic, adLockOptimistic
    Else
        Master.Open "select Top 1 Jc.DocId AS SearchCode from Job_Card as JC where left(JC.DocID,1)='" & PubDivCode & "' and right(" & xIsNull("jc.DocId_InvSpr", "") & ",8) <> 'Cancelld' " & SiteCond & " order by JC.DocID", GCn, adOpenDynamic, adLockOptimistic
    End If
'    Master.Open "select Jc.DocId AS CODE,Jc.DocId AS SearchCode,JC.Job_No,JC.Site_Code, HC.Model,HC.RegNo, HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name, Jc.DocId,JC.Govt_YN, JC.Job_Date, JC.JobCloseDate,jc.cardno,jc.ObservBy_Super,jc.ActionBy_Super,jc.ObservBy_Eng, HC.Add1, HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc, City.CityName from ((job_card as JC left Join Hiscard as HC on JC.CardNo=HC.CardNo) left Join Service_Type as ST on JC.Serv_Type=ST.Serv_Type) Left Join City on HC.CityCode=City.CityCode where left(JC.site_code,1)='" & PubSiteCode & "' order by JC.docID", GCn, adOpenDynamic, adLockOptimistic
       
       If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = "and  " & cMID("Jc.DocId", "3", "1") & "='" & PubSiteCode & "'"
    Else
      SiteCond = ""
    End If
    
    Set RsJob = New ADODB.Recordset
    With RsJob
        .CursorLocation = adUseClient
        .Open "select  Jc.DocId AS CODE,JC.Job_No, HC.Model,HC.RegNo, HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name, Jc.DocId,JC.Govt_YN,Jc.OpenRemarks, JC.Job_Date, JC.JobCloseDate,jc.cardno, HC.Add1, HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc, City.CityName from ((job_card as JC left Join Hiscard as HC on JC.CardNo=HC.CardNo) left Join Service_Type as ST on JC.Serv_Type=ST.Serv_Type) Left Join City on HC.CityCode=City.CityCode where left(JC.DocId,1)='" & PubDivCode & "' " & SiteCond & " order by JC.docID", GCn, adOpenDynamic, adLockOptimistic
    End With
    RsJob.Sort = "code"
    
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

Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set RsJob = Nothing
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
    If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
        
        GCn.Execute ("update job_card set observby_super='',actionby_super='',observBy_eng='' where job_card.docid='" & lblDocId & "'")
        GCn.Execute "Delete from Job_card2  where Docid='" & lblDocId.CAPTION & "'"
        
        GCn.CommitTrans
        
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
    GCn.RollbackTrans
    MsgBox err.Description, vbCritical, " Deletion Message"
End Sub

Private Sub TopCtrl1_eEdit()
Dim I As Integer
On Error GoTo eloop1
    If RsJob.EOF = True Or RsJob.BOF = True Then Exit Sub
    If Not IsNull(RsJob!JobCloseDate) Then
        MsgBox "JobCard is Closed,Editing not allowed", vbInformation, "Validation"
        Exit Sub
    End If
    Disp_Text SETS("EDIT", Me, Master)
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    lblDocId.CAPTION = RsJob!Code
    
    txt(JobNo).Enabled = False
    txt(Chassis).Enabled = False
    txt(OwnerName).Enabled = False
    txt(VehRegNo).Enabled = False
    
    txt(SObser).Locked = False
    txt(SAction).Locked = False
    
    txt(SObser).SetFocus
    txt(SObser).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
'    If TopCtrl1.TopText2 = "Browse" Then Unload Me
    Unload Me
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    
    
     Dim SiteCond As String
     SiteCond = " And  Job_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      SiteCond = SiteCond & " and " & cMID("jc.Docid", "3", "1") & "='" & PubSiteCode & "'"
    End If
    
    GSQL = "SELECT JC.DocId as searchcode,Jc.Job_no,JC.Job_date,JC.JobCloseDate, ServICE_type.SERV_Desc,Coupon,hiscard.regno,hiscard.chassis,hiscard.name  FROM (JOB_card as jc left join  servICE_type on jc.serv_type=servICE_type.serv_type) left join HISCARD ON jc.cardno=hiscard.cardno  where left(JC.DocID,1)='" & PubDivCode & "' " & SiteCond & " order by jc.JOB_NO"
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
        Set Master = GCn.Execute("Select Top 1 Jc.DocId AS SearchCode from Job_Card as JC where left(JC.DocID,1)='" & PubDivCode & "' and right(" & xIsNull("jc.DocId_InvSpr", "") & ",8) <> 'Cancelld' order by JC.DocID")
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

Private Sub TopCtrl1_eRef()
    Call UpdRequery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim SrNo As Integer
    Dim mTrans As Boolean
    On Error GoTo errlbl

    Grid_Hide
    If IsValid(txt(JobNo), "Job Card No.") = False Then Exit Sub
                
    For I = 1 To Fgrid1.Rows - 1
        If Fgrid1.TextMatrix(I, C_Observ) <> "" And txt(MechName).TEXT = "" Then MsgBox "Person Name is Required Information", vbInformation, "Validation": txt(MechName).SetFocus: Exit Sub
    Next I
    
    If txt(SObser).TEXT <> "" And IsValid(txt(SAction), "Action Taken") = False Then Exit Sub
    
    GCn.BeginTrans
    mTrans = True
    
    GCn.Execute "Delete from Job_card2  where Docid='" & lblDocId.CAPTION & "'"
    GSQL = "Update Job_Card set ObservBy_Super='" & txt(SObser).TEXT & "', ActionBy_Super='" & txt(SAction).TEXT & "',ObservBy_Eng='" & txt(MechName).TEXT & "' where DocId='" & lblDocId & "'"
    GCn.Execute GSQL
    
    SrNo = 1
    For I = 1 To Fgrid1.Rows - 1
        If Trim(Fgrid1.TextMatrix(I, C_Desc) + Fgrid1.TextMatrix(I, C_Make) + Fgrid1.TextMatrix(I, C_Observ)) <> "" Then
            GSQL = "insert into Job_Card2(" _
                & "DocId,S_No,Site_Code,Particulars,Detail," _
                & "Make, U_Name, U_EntDt, U_AE) " _
                & " values(" _
                & "'" & lblDocId.CAPTION & "'," & SrNo & ",'" & PubSiteCode & "','" & Fgrid1.TextMatrix(I, C_Desc) & "','" & Fgrid1.TextMatrix(I, C_Observ) & "'," _
                & "'" & Fgrid1.TextMatrix(I, C_Make) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
            GCn.Execute GSQL
            SrNo = SrNo + 1
        End If
    Next I
    
    GCn.CommitTrans
    mTrans = False
    
    Master.Requery
    Call UpdRequery
    
    Master.FIND "SearchCode = '" & lblDocId & "'"
'    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub

errlbl:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus txt(Index)
    Grid_Hide
    MyIndex = Index
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    '' KEY DOWN
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = MechName Then
        Ctrl_DownKeyDown KeyCode, Shift
    End If
    
    ' KEY UP
    If Index = MechName Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate txt(Index)
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
    For I = 1 To txt.Count
        txt(I).TEXT = ""
        txt(I).Tag = ""
    Next I
    lblDocId.CAPTION = ""
    lblDocId.Refresh
    
    Fgrid1.Rows = 2
    Fgrid1.FixedRows = 1
End Sub

Private Sub MoveRec()
'Dim rs As Recordset
Dim Master1 As Recordset
'Dim mVor As String
'Dim i As Integer
On Error GoTo error1
    TopCtrl1.tAdd = False
    If Master.RecordCount > 0 Then
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "select JC.Job_No,JC.Site_Code, HC.Model,HC.RegNo, HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name, Jc.DocId,JC.Govt_YN, JC.Job_Date, JC.JobCloseDate,jc.cardno,jc.ObservBy_Super,jc.ActionBy_Super,jc.ObservBy_Eng, HC.Add1, HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, HC.Mobile, ST.Serv_Desc, City.CityName from ((job_card as JC left Join Hiscard as HC on JC.CardNo=HC.CardNo) left Join Service_Type as ST on JC.Serv_Type=ST.Serv_Type) Left Join City on HC.CityCode=City.CityCode where JC.DocId='" & Master!SearchCode & "' order by JC.docID", GCn, adOpenStatic, adLockReadOnly
        
        txt(SObser).Enabled = True
        txt(SObser).Locked = True
        txt(SAction).Enabled = True
        txt(SAction).Locked = True
                
        LblDiv.CAPTION = "Division : " & left(Master1!DocID, 1)
        LblSite.CAPTION = "Site Code : " & Master1!Site_Code
        lblDocId.CAPTION = Master1!DocID
        '
        txt(SObser).TEXT = XNull(Master1!ObservBy_Super)
        txt(SAction).TEXT = XNull(Master1!ActionBy_Super)
        txt(MechName).TEXT = XNull(Master1!ObservBy_Eng)
        
        RsJob.Sort = "code"
        RsJob.FIND ("Code='" & Master1!DocID & "'")
        
        If RsJob.EOF = True Or RsJob.BOF = True Then
            txt(JobNo).TEXT = ""
            txt(JobDt).TEXT = ""
            txt(JobClDt) = ""
            txt(VehRegNo).TEXT = ""
            txt(Chassis).TEXT = ""
            txt(Model).TEXT = ""
            txt(Engine).TEXT = ""
            txt(VehSrlNo).TEXT = ""
            txt(OwnerName).TEXT = ""
            txt(Address1).TEXT = ""
            txt(Address2).TEXT = ""
            txt(Address3).TEXT = ""
            txt(PhoneOff).TEXT = ""
            txt(PhoneResi).TEXT = ""
            txt(Mobile).TEXT = ""
            txt(City).TEXT = ""
            txt(SrvType).TEXT = ""
            txt(Remarks).TEXT = ""
        Else
            txt(JobNo).TEXT = RsJob!Job_No
            txt(JobDt).TEXT = RsJob!Job_Date
            txt(JobClDt).TEXT = IIf(IsNull(RsJob!JobCloseDate), "", RsJob!JobCloseDate)
            txt(VehRegNo).TEXT = XNull(RsJob!RegNo)
            txt(Chassis).TEXT = XNull(RsJob!Chassis)
            txt(Model).TEXT = XNull(RsJob!Model)
            txt(Engine).TEXT = XNull(RsJob!Engine)
            txt(VehSrlNo).TEXT = XNull(RsJob!VehSerialNo)
            txt(OwnerName).TEXT = XNull(RsJob!Name)
            txt(Address1).TEXT = XNull(RsJob!Add1)
            txt(Address2).TEXT = XNull(RsJob!Add2)
            txt(Address3).TEXT = XNull(RsJob!Add3)
            txt(PhoneOff).TEXT = XNull(RsJob!PhoneOff)
            txt(PhoneResi).TEXT = XNull(RsJob!PhoneResi)
            txt(Mobile).TEXT = XNull(RsJob!Mobile)
            txt(City).TEXT = XNull(RsJob!CityName)
            txt(SrvType).TEXT = XNull(RsJob!Serv_Desc)
            txt(Remarks).TEXT = XNull(RsJob!OpenRemarks)
        End If
        Call Fill_Grid
    Else
        Call BlankText
    End If
    Grid_Hide
    Set Master1 = Nothing
    
    Exit Sub
error1:
    CheckError
End Sub

Private Sub Ini_Grid()
    
    With Fgrid1
'        .left = Me.left '+ 60
'        .width = Me.width - 90
'        .top = 2550
'        .BackColor = CellBackColLeave
'        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 4
        
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 400
        
        .TextMatrix(0, C_Desc) = "Item Particular"
        .ColAlignment(C_Desc) = flexAlignLeftCenter
        .ColWidth(C_Desc) = 2700
        
        .TextMatrix(0, C_Observ) = "Observation"
        .ColAlignment(C_Observ) = flexAlignLeftCenter
        .ColWidth(C_Observ) = 5000

        .TextMatrix(0, C_Make) = "Make"
        .ColAlignment(C_Make) = flexAlignLeftCenter
        .ColWidth(C_Make) = 2700
    End With
    BackColorSelLeave = Fgrid1.BackColorSel
    ForeColorSelEnter = Fgrid1.ForeColorSel
    
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
    
    txt(JobDt).Enabled = False
    txt(JobClDt).Enabled = False
    txt(Remarks).Enabled = False
    txt(Engine).Enabled = False
    txt(Model).Enabled = False
    txt(VehSrlNo).Enabled = False
    txt(SrvType).Enabled = False
    
    txt(Address1).Enabled = False
    txt(Address2).Enabled = False
    txt(Address3).Enabled = False
    txt(City).Enabled = False
    txt(PhoneOff).Enabled = False
    txt(PhoneResi).Enabled = False
    txt(Mobile).Enabled = False
End Sub

Private Sub Grid_Hide()
    ''''
End Sub

Private Sub UpdRequery()
    RsJob.Requery
End Sub

Private Sub History_Field()
    txt(VehRegNo).Tag = XNull(RsJob!Code)
    txt(Chassis).Tag = XNull(RsJob!Code)
    txt(OwnerName).Tag = XNull(RsJob!Code)
    txt(JobNo).Tag = XNull(RsJob!Code)
    
    txt(JobNo).TEXT = XNull(RsJob!Job_No)
    txt(JobDt).TEXT = RsJob!Job_Date
    txt(VehRegNo).TEXT = XNull(RsJob!RegNo)
    txt(Chassis).TEXT = XNull(RsJob!Chassis)
    txt(Model).TEXT = XNull(RsJob!Model)
    txt(Engine).TEXT = XNull(RsJob!Engine)
    txt(VehSrlNo).TEXT = XNull(RsJob!VehSerialNo)
    txt(OwnerName).TEXT = XNull(RsJob!Name)
    txt(Address1).TEXT = XNull(RsJob!Add1)
    txt(Address2).TEXT = XNull(RsJob!Add2)
    txt(Address3).TEXT = XNull(RsJob!Add3)
    txt(City).TEXT = XNull(RsJob!CityName)
    txt(PhoneOff).TEXT = XNull(RsJob!PhoneOff)
    txt(PhoneResi).TEXT = XNull(RsJob!PhoneResi)
    txt(Mobile).TEXT = XNull(RsJob!Mobile)
    txt(Remarks).TEXT = XNull(RsJob!OpenRemarks)
End Sub

Private Sub FGrid1_Click()
    txtgrid1(0).Visible = False
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
        Select Case Fgrid1.Col
        Case C_Desc
            txtgrid1(0).MaxLength = 20
            txtgrid1(0).Alignment = 0
        Case C_Observ
            txtgrid1(0).MaxLength = 40
            txtgrid1(0).Alignment = 0
        Case C_Make
            txtgrid1(0).MaxLength = 20
            txtgrid1(0).Alignment = 0
'        Case Else
'            txtgrid1(0).MaxLength = 0
    End Select

End Sub

Private Sub FGrid1_DblClick()
FGrid1_KeyPress vbKeyReturn
End Sub

Private Sub FGrid1_GotFocus()
    Fgrid1.BackColorSel = BackColorSelEnter
    Fgrid1.ForeColorSel = ForeColorSelEnter
    txtgrid1(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(Fgrid1.Tag) = (Fgrid1.Rows - (Fgrid1.Rows - 1)) Then
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(Fgrid1.Tag) = Fgrid1.Rows - 1 Then
        If MsgBox("Save Entry ?", vbInformation + vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave: Exit Sub
        Fgrid1.SetFocus
        KeyCode = 0
    End If
    GridKey = KeyCode
    Fgrid1.Tag = Fgrid1.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Fgrid1.TextMatrix(Fgrid1.Row, Fgrid1.Col) = ""
    End If
    If KeyCode = vbKeyReturn Then
        GridDblClick Me, Fgrid1, txtgrid1, 0
    End If
    TAddMode = False
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
    Select Case Fgrid1.Col
        Case C_Desc
            txtgrid1(0).MaxLength = 20
            txtgrid1(0).Alignment = 0
        Case C_Observ
            txtgrid1(0).MaxLength = 40
            txtgrid1(0).Alignment = 0
        Case C_Make
            txtgrid1(0).MaxLength = 20
            txtgrid1(0).Alignment = 0
'        Case Else
'            txtgrid1(0).MaxLength = 0
    End Select
    Get_Text Me, Fgrid1, txtgrid1, 0, False, KeyAscii
    If KeyAscii <> vbKeyReturn Then TAddMode = True
    Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If Fgrid1.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If Fgrid1.Row >= 1 Then
            If MsgBox("Are You Sure To Delete Entry ?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                If Fgrid1.Rows > 2 Then
                    Fgrid1.RemoveItem (Fgrid1.Row)
                Else
                    Fgrid1.Rows = 1
                    Fgrid1.AddItem Fgrid1.Rows
                    Fgrid1.FixedRows = 1
                End If
            End If
            For I = 1 To Fgrid1.Rows - 1
                Fgrid1.TextMatrix(I, 0) = I
            Next
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        Fgrid1.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_LostFocus()
    Fgrid1.BackColorSel = BackColorSelLeave
    Fgrid1.ForeColorSel = Fgrid1.ForeColor
End Sub

Private Sub FGrid1_Scroll()
    txtgrid1(0).Visible = False
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
On Error GoTo ELoop
    Grid_Hide
    txtgrid1(0).Tag = Fgrid1.TextMatrix(Fgrid1.Row, Fgrid1.Col)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then txtgrid1(0).TEXT = txtgrid1(0).Tag: Exit Sub
    If KeyCode = vbKeyReturn Then
        If TxtGrid1Leave Then GridTxtDown Fgrid1, txtgrid1, Index, KeyCode, TAddMode, 3
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Exit Sub
    Call CheckQuote(KeyAscii)
End Sub

Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then
    Fgrid1.SetFocus
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
    Fgrid1.TextMatrix(Fgrid1.Row, Fgrid1.Col) = txtgrid1(0).TEXT
'    txtgrid1(0).MaxLength = 40
    
If ValidateCall = False Then
    Fgrid1.SetFocus
    txtgrid1(Index).Visible = False
End If
End Function

Private Sub Fill_Grid()
Dim MyRst As ADODB.Recordset
Dim I As Integer
    Fgrid1.Rows = 1
    Set MyRst = New ADODB.Recordset
    MyRst.CursorLocation = adUseClient
    GSQL = "Select JC2.* From Job_card2 as jc2 Where Jc2.DocId='" & lblDocId & "' order by Jc2.S_no"
    MyRst.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    I = 1
    If MyRst.RecordCount > 0 Then
        Do Until MyRst.EOF
            Fgrid1.AddItem ""
            With Fgrid1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, C_Desc) = MyRst!Particulars
                .TextMatrix(I, C_Observ) = XNull(MyRst!detail)
                .TextMatrix(I, C_Make) = XNull(MyRst!Make)
            End With
            I = I + 1
            MyRst.MoveNext
        Loop
        Fgrid1.AddItem ""
        Fgrid1.FixedRows = 1
    Else
        Fgrid1.Rows = Fgrid1.Rows
        Fgrid1.AddItem Fgrid1.Rows
        Fgrid1.FixedRows = 1
    End If
    Set MyRst = Nothing
End Sub
