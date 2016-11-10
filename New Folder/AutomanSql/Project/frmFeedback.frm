VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmFeedback 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Customer FeedBack Entry"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   9165
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3135
      TabIndex        =   33
      Top             =   45
      Width           =   345
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   360
      TabIndex        =   31
      Top             =   3390
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   60
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   225
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
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   2850
      MaxLength       =   50
      TabIndex        =   15
      Top             =   6615
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   12
      Left            =   6870
      MaxLength       =   50
      TabIndex        =   14
      Top             =   6285
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   11
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6300
      Width           =   2700
   End
   Begin MSDataGridLib.DataGrid DGJob 
      Height          =   1365
      Left            =   210
      Negotiate       =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4785
      Visible         =   0   'False
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   2408
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Job_No"
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
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3195.213
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   2250
      TabIndex        =   1
      Top             =   990
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   10
      Left            =   7995
      TabIndex        =   10
      Top             =   2220
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   9
      Left            =   7995
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1590
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   8
      Left            =   7995
      TabIndex        =   8
      Top             =   1290
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   7995
      TabIndex        =   7
      Top             =   990
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   7995
      TabIndex        =   6
      Top             =   690
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   2250
      TabIndex        =   5
      Top             =   2190
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   2250
      TabIndex        =   4
      Top             =   1890
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   2250
      TabIndex        =   3
      Top             =   1590
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   2250
      TabIndex        =   2
      Top             =   1290
      Width           =   2700
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   2250
      TabIndex        =   0
      Top             =   690
      Width           =   2700
   End
   Begin MSFlexGridLib.MSFlexGrid FGrid 
      Height          =   3450
      Left            =   105
      TabIndex        =   12
      Top             =   2775
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   6085
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   12570311
      BackColorBkg    =   13623520
      GridLinesFixed  =   1
      Appearance      =   0
      FormatString    =   $"frmFeedback.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   661
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dissatisfied Complaint Nature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   13
      Left            =   60
      TabIndex        =   30
      Top             =   6675
      Width           =   2640
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expected Next Visit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   12
      Left            =   4890
      TabIndex        =   29
      Top             =   6360
      Width           =   1740
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status of Feedback"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   11
      Left            =   75
      TabIndex        =   28
      Top             =   6345
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Card Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   6
      Left            =   330
      TabIndex        =   25
      Top             =   990
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Close Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   5
      Left            =   330
      TabIndex        =   24
      Top             =   1305
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   4
      Left            =   330
      TabIndex        =   23
      Top             =   1620
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mechanic Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   330
      TabIndex        =   22
      Top             =   1935
      Width           =   1470
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Advisor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   330
      TabIndex        =   21
      Top             =   2235
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Card No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   330
      TabIndex        =   20
      Top             =   705
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   10
      Left            =   6060
      TabIndex        =   19
      Top             =   975
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   9
      Left            =   6060
      TabIndex        =   18
      Top             =   1290
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   8
      Left            =   6060
      TabIndex        =   17
      Top             =   1605
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   7
      Left            =   6060
      TabIndex        =   16
      Top             =   2235
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   6060
      TabIndex        =   11
      Top             =   675
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2145
      Left            =   120
      Top             =   525
      Width           =   11430
   End
End
Attribute VB_Name = "frmFeedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const JobNo = 0
Private Const JobDate = 1
Private Const JobCloseDate = 2
Private Const ServName = 3
Private Const MechName = 4
Private Const ServAdv = 5
Private Const RegNo = 6
Private Const CustName = 7
Private Const ConctPerson = 8
Private Const CustAdd = 9
Private Const CustPhone = 10
Private Const FeedbackStatus = 11
Private Const NxtVisit = 12
Private Const DisNature = 13
Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4

Dim ListArray As Variant
Dim mListItem As ListItem

Dim MyIndex As Byte
Dim ADDFLAG$
Dim TAddMode As Boolean
Dim mRepName As String
Dim RsJob As ADODB.Recordset
Dim Master As ADODB.Recordset
Private Sub Grid_Hide()
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGJob.Visible = True Then DGJob.Visible = False
End Sub
Private Sub Ini_Grid()
    DGJob.left = FGrid.left: DGJob.width = FGrid.width: DGJob.top = FGrid.top: DGJob.height = FGrid.height
End Sub
Private Sub FGrid_Click()
Dim I As Integer
If TopCtrl1.TopText2 = "Add" Or TopCtrl1.TopText2 = "Edit" And FGrid.Col <> 0 Then
    For I = 1 To 5
        FGrid.TextMatrix(FGrid.Row, I) = ""
    Next
    FGrid.CellFontName = "wingdings"
    FGrid.CellFontSize = 18
    FGrid.CellForeColor = vbBlue
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "ü"
End If
End Sub
Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And FGrid.Col <> 0 Then
        FGrid_Click
    End If
End Sub
Private Sub BlankText()
Dim I As Byte, j As Byte
    For I = 1 To txt.Count - 1
        txt(I).TEXT = ""
    Next I
    For I = 1 To 8
        For j = 1 To 5
            FGrid.TextMatrix(I, j) = ""
        Next
    Next
End Sub
Private Sub ListView_Click()
On Error GoTo ELoop
    txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    txt(Val(ListView.Tag)).SetFocus
    FrmList.Visible = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Activate()
    Call FillGrid
    Call MoveRec
End Sub

Private Sub Form_Load()
    TopCtrl1.Tag = PubUParam: WinSetting Me
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select Job_DocId as SearchCode from CustFeedback where left(Job_DocID,1)='" & PubDivCode & "' Order by Job_DocId Desc", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsJob = New ADODB.Recordset
    With RsJob
        .CursorLocation = adUseClient
        .Open "select  J.DocId AS CODE, " & cCStr("J.Job_No") & " As FindJobNo,J.Job_No, HC.Model,HC.RegNo," & _
              "HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name, J.DocId,J.Govt_YN, J.Job_Date," & _
              "J.JobCloseDate,j.cardno, HC.Add1, HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, " & _
              "HC.Mobile,HC.ConPerson,ST.Serv_Desc, City.CityName,Emp.Emp_Name as MechName,Emp.Emp_Code,Emp1.Emp_Name as Supervisor " & _
              "from ((((job_card as J left Join Hiscard as HC on J.CardNo=HC.CardNo) " & _
              "LEFT JOIN Service_Type as ST on J.Serv_Type=ST.Serv_Type) " & _
              "LEFT JOIN EMP_MAST as EMP ON J.RECBY_MECHANIC=EMP.EMP_CODE) " & _
              "LEFT JOIN EMP_MAST as EMP1 on J.RecBy_Supervisor=EMP1.EMP_CODE) " & _
              "LEFT JOIN City on HC.CityCode=City.CityCode  " & _
              "where left(j.DocId,1)='" & PubDivCode & "' and J.JobCloseDate is not null " & _
              "Order by J.docID", GCn, adOpenDynamic, adLockOptimistic
    End With
    RsJob.Sort = "code"
    Set DGJob.DataSource = RsJob
    Disp_Text SETS("INI", Me, Master)
    
End Sub
Private Sub DGJob_Click()
'Call History_Field
DGJob.Visible = False
End Sub
Private Sub History_Field(Optional MakeBlank As Boolean)
If TopCtrl1.TopText2 = "Add" Then
    txt(JobNo).Tag = XNull(RsJob!Code)
    txt(JobNo).TEXT = XNull(RsJob!Job_No)
    txt(JobDate).TEXT = RsJob!Job_Date
    txt(ServName).TEXT = XNull(RsJob!Serv_Desc)
    txt(RegNo).TEXT = XNull(RsJob!RegNo)
    
    txt(JobCloseDate).TEXT = XNull(RsJob!JobCloseDate)
    txt(ServAdv).TEXT = XNull(RsJob!Supervisor)
    txt(CustName).TEXT = XNull(RsJob!Name)
    txt(ConctPerson).TEXT = XNull(RsJob!ConPerson)
    txt(CustAdd).TEXT = XNull(RsJob!Add1) & " " & XNull(RsJob!Add2) & " " & XNull(RsJob!Add3) & " " & XNull(RsJob!CityName)
    txt(CustPhone).TEXT = XNull(RsJob!PhoneResi)
    
    txt(MechName).TEXT = XNull(RsJob!MechName)
    txt(MechName).Tag = XNull(RsJob!Emp_Code)
End If
End Sub
Public Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    'New Testing for Speed purpose
    ADDFLAG = left(TopCtrl1.TopText2, 1)
    'eof New Testing
    For I = 11 To txt.Count - 1
        txt(I).Enabled = Enb
    Next
    For I = 1 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg
        txt(I).ForeColor = CtrlFColOrg
    Next
    txt(JobNo).Enabled = True
    If TopCtrl1.TopText2 <> "Add" Then
        txt(JobNo).Enabled = False
    End If
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    TopCtrl1.TopText2 = "Add"
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    Ini_Grid
    txt(JobNo).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub
Private Sub TopCtrl1_eCancel()
Dim I As Integer
On Error GoTo ErrorLoop
    If MsgBox("Cancel Entry ?", vbExclamation + vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        RsJob.Filter = ""
        Call Ini_Grid
'        Call MoveRec
    Else
        Me.ActiveControl.SetFocus
    End If
    Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eRef()
    RsJob.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer, j As Integer, No As Integer, SelNo As Integer
    Dim parameterArr(8, 2) As Integer
    Dim mTrans As Boolean
'    On Error GoTo errlbl
    
    Grid_Hide
    If TopCtrl1.TopText2 = "Add" Then
        If IsValid(txt(JobNo), "JobCard No") = False Then Exit Sub
    End If
    GCn.BeginTrans
    mTrans = True
    For I = 1 To 8
        No = 5
        For j = 1 To 5
            If FGrid.TextMatrix(I, j) <> "" Then
               SelNo = j
               parameterArr(I, 1) = No
            End If
            No = No - 1
        Next
        parameterArr(I, 0) = SelNo
    Next
    Select Case TopCtrl1.TopText2
        Case "Add"
            GSQL = "insert into CustFeedback (Job_DocId,Parameter1, Parameter2, Parameter3,Parameter4,Parameter5,Parameter6,Parameter7,Parameter8," & _
                   "Point1,Point2,Point3,Point4,Point5,Point6,Point7,Point8,FeedbackStat,NxtVisit,CompNature, U_Name, U_EntDt, U_AE) " & _
                   " values('" & txt(JobNo).Tag & "'," & parameterArr(1, 0) & "," & parameterArr(2, 0) & "," & parameterArr(3, 0) & "," & parameterArr(4, 0) & "," & parameterArr(5, 0) & "," & parameterArr(6, 0) & "," & parameterArr(7, 0) & "," & parameterArr(8, 0) & "," & _
                   "" & parameterArr(1, 1) & "," & parameterArr(2, 1) & "," & parameterArr(3, 1) & "," & parameterArr(4, 1) & "," & parameterArr(5, 1) & "," & parameterArr(6, 1) & "," & parameterArr(7, 1) & "," & parameterArr(8, 1) & "," & _
                   "'" & txt(FeedbackStatus) & " ','" & txt(NxtVisit) & "','" & txt(DisNature) & "'," & _
                   "'" & pubUName & " '," & ConvertDate(PubServerDate) & ",'A')"
            GCn.Execute GSQL
        Case "Edit"
           GCn.Execute "Update CustFeedback Set Job_DocId='" & txt(JobNo).Tag & "',Parameter1=" & parameterArr(1, 0) & ",Parameter2 =" & parameterArr(2, 0) & ",Parameter3 =" & parameterArr(3, 0) & ",Parameter4 =" & parameterArr(4, 0) & ",Parameter5 =" & parameterArr(5, 0) & ",Parameter6 =" & parameterArr(6, 0) & ",Parameter7 =" & parameterArr(7, 0) & ",Parameter8 =" & parameterArr(8, 0) & "," & _
                "Point1 =" & parameterArr(1, 1) & ",Point2 =" & parameterArr(2, 1) & ",Point3 =" & parameterArr(3, 1) & ",Point4 =" & parameterArr(4, 1) & ",Point5 =" & parameterArr(5, 1) & ",Point6 =" & parameterArr(6, 1) & ",Point7 =" & parameterArr(7, 1) & ",Point8 =" & parameterArr(8, 1) & "," & _
                "FeedbackStat ='" & txt(FeedbackStatus) & "',NxtVisit ='" & txt(NxtVisit) & "',CompNature ='" & txt(DisNature) & "'," & _
                "U_Name ='" & pubUName & "',U_EntDt = " & ConvertDate(PubServerDate) & ",U_AE ='E' where Job_Docid ='" & txt(JobNo).Tag & "'"
    End Select
    GCn.CommitTrans
    mTrans = False
    Master.Requery
    Disp_Text SETS("INI", Me, Master)
    Exit Sub
errlbl:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "select  J.DocId AS SearchCode,J.Job_No,J.Job_Date,J.JobCloseDate " & _
              "from (job_card as J left Join CustFeedback as CF on J.DocId=CF.Job_DocID)" & _
              "where left(j.DocId,1)='" & PubDivCode & "' " & _
              "Order by J.docID"
        Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus txt(Index)
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case JobNo
            If RsJob.RecordCount <= 0 Then Exit Sub
            DGridColSwap DGJob, 0
            RsJob.Sort = "JOB_NO"
            If txt(Index).Tag <> "" And txt(Index).Tag <> RsJob!Code Then
                RsJob.FIND ("JOB_NO='" & txt(Index).TEXT & "'")
            End If
        Case FeedbackStatus
            ListArray = Array("WrongNo", "Waiting", "Matured", "Satisfied", "Dissatisfied")
            Set mListItem = ListView_Items(ListView, txt, Index, ListArray, 5)
    End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case JobNo
            DGridTxtKeyDown DGJob, txt, Index, RsJob, KeyCode, False, 1
        Case FeedbackStatus
            ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top - 1500), txt(Index).width, 1500
    End Select
    If DGJob.Visible = False And FrmList.Visible = False Then
        '' KEY DOWN
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And ((ADDFLAG = "A" And Index <> DisNature) Or (ADDFLAG = "E" And Index <> DisNature)) Then
            Ctrl_DownKeyDown KeyCode, Shift
            Call History_Field
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And ((ADDFLAG = "A" And Index = DisNature) Or (ADDFLAG = "E" And Index = DisNature)) Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        ' KEY UP
        If ADDFLAG = "A" Then
            If Index <> GPNo Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        ElseIf ADDFLAG = "E" Then
            If Index <> GPDt Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
End Sub
Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)
    Select Case Index
        Case JobNo
            DGridTxtKeyPress txt, Index, RsJob, keyascii, "Findjobno"
    End Select
End Sub
Private Sub TopCtrl1_eEdit()
Dim I As Integer
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    FGrid.SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub
Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
    If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
        GCn.Execute "Delete from CustFeedback  where job_Docid='" & txt(JobNo).Tag & "'"
        GCn.CommitTrans
        Master.Requery
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

Private Sub TopCtrl1_eExit()
    Unload Me
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
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("searchcode='" & MyValue & "'")
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case FeedbackStatus
        If FrmList.Visible = True Then ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Ctrl_validate txt(Index)
End Sub
Private Sub MoveRec()
Dim Master1 As Recordset, rs1 As Recordset
Dim SelArr(7) As Integer
Dim I As Integer
    Call BlankText
    If Master.RecordCount > 0 Then
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
          With Master1
            .Open "select CF.*, J.DocId AS CODE, " & cCStr("J.Job_No") & " As FindJobNo,J.Job_No,HC.RegNo," & _
                  "HC.Name, J.DocId,J.Job_Date," & _
                  "J.JobCloseDate,j.cardno, HC.Add1, HC.Add2, HC.add3, HC.PhoneOff, HC.PhoneResi, " & _
                  "HC.Mobile,HC.ConPerson,ST.Serv_Desc, City.CityName,Emp.Emp_Name as MechName,Emp.Emp_Code,Emp1.Emp_Name as Supervisor " & _
                  "from (((((job_card as J left Join Hiscard as HC on J.CardNo=HC.CardNo) " & _
                  "LEFT JOIN Service_Type as ST on J.Serv_Type=ST.Serv_Type) " & _
                  "LEFT JOIN EMP_MAST as EMP ON J.RECBY_MECHANIC=EMP.EMP_CODE) " & _
                  "LEFT JOIN EMP_MAST as EMP1 on J.RecBy_Supervisor=EMP1.EMP_CODE) " & _
                  "LEFT JOIN City on HC.CityCode=City.CityCode)  " & _
                  "LEFT JOIN CustFeedback CF on J.DocID=CF.Job_DocId  " & _
                  "Where J.DocID='" & Master!SearchCode & "'" & _
                  "Order by J.docID", GCn, adOpenDynamic, adLockOptimistic
        End With
            txt(JobNo).Tag = XNull(Master1!Code)
            txt(JobNo).TEXT = XNull(Master1!Job_No)
            txt(JobDate).TEXT = Master1!Job_Date
            txt(ServName).TEXT = XNull(Master1!Serv_Desc)
            txt(RegNo).TEXT = XNull(Master1!RegNo)
            
            txt(JobCloseDate).TEXT = XNull(Master1!JobCloseDate)
            txt(ServAdv).TEXT = XNull(Master1!Supervisor)
            txt(CustName).TEXT = XNull(Master1!Name)
            txt(ConctPerson).TEXT = XNull(Master1!ConPerson)
            txt(CustAdd).TEXT = XNull(Master1!Add1) & " " & XNull(Master1!Add2) & " " & XNull(Master1!Add3) & " " & XNull(Master1!CityName)
            txt(CustPhone).TEXT = XNull(Master1!PhoneResi)
            
            txt(MechName).TEXT = XNull(Master1!MechName)
            txt(MechName).Tag = XNull(Master1!Emp_Code)
            SelArr(0) = Master1!Parameter1: SelArr(1) = Master1!Parameter2: SelArr(2) = Master1!Parameter3
            SelArr(3) = Master1!Parameter4: SelArr(4) = Master1!Parameter5: SelArr(5) = Master1!Parameter6
            SelArr(6) = Master1!Parameter7: SelArr(7) = Master1!Parameter8
            For I = 0 To 7
                If SelArr(I) <> 0 Then
                    FGrid.Row = I + 1: FGrid.Col = SelArr(I)
                    FGrid.CellFontName = "wingdings"
                    FGrid.CellFontSize = 18
                    FGrid.CellForeColor = vbBlue
                    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = "ü"
                End If
            Next
            txt(FeedbackStatus).TEXT = XNull(Master1!FeedbackStat)
            txt(NxtVisit).TEXT = XNull(Master1!NxtVisit)
            txt(DisNature).TEXT = XNull(Master1!CompNature)
            Grid_Hide
        Set Master1 = Nothing
        Exit Sub
        End If
error1:
    CheckError
End Sub
Private Sub FillGrid()
    With FGrid
        .Rows = 9
        .TextMatrix(1, 0) = "Ease of obtaining appointment"
        .TextMatrix(2, 0) = "Time taken to open a job card"
        .TextMatrix(3, 0) = "Attitude of the service person"
        .TextMatrix(4, 0) = "Car delevered of time"
        .TextMatrix(5, 0) = "Solution to all problems reported by you"
        .TextMatrix(6, 0) = "Explanation of job done and the bill"
        .TextMatrix(7, 0) = "General appearance of the workshop"
        .TextMatrix(8, 0) = "Quality of washing"
    End With
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case FeedbackStatus
        txt(Index).TEXT = ListView.SelectedItem.TEXT
End Select
End Sub
