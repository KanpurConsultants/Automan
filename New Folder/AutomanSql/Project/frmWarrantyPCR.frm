VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmWarrantyPCR 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Product Complaint Report Entry"
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
      Index           =   20
      Left            =   5235
      MaxLength       =   10
      TabIndex        =   20
      Top             =   2955
      Width           =   930
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
      Index           =   14
      Left            =   5235
      MaxLength       =   20
      TabIndex        =   14
      Top             =   2145
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
      Index           =   12
      Left            =   5235
      MaxLength       =   8
      TabIndex        =   12
      Top             =   1875
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
      Index           =   15
      Left            =   1545
      MaxLength       =   1
      TabIndex        =   15
      ToolTipText     =   "Status Code : 1 - Drive Away, 2 - Sold"
      Top             =   2415
      Width           =   405
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
      Left            =   1545
      MaxLength       =   2
      TabIndex        =   1
      Top             =   525
      Width           =   405
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
      Left            =   1545
      MaxLength       =   2
      TabIndex        =   2
      Top             =   795
      Width           =   405
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
      Left            =   5235
      MaxLength       =   40
      TabIndex        =   18
      Top             =   2685
      Width           =   5370
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
      Left            =   1545
      MaxLength       =   8
      TabIndex        =   5
      Text            =   "Help"
      Top             =   1065
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5235
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1065
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
      Left            =   4275
      MaxLength       =   40
      TabIndex        =   25
      Top             =   5700
      Visible         =   0   'False
      Width           =   705
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
      Height          =   660
      Index           =   23
      Left            =   165
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   4125
      Width           =   5790
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
      Height          =   660
      Index           =   24
      Left            =   6015
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   4125
      Width           =   5790
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
      Index           =   21
      Left            =   6180
      MaxLength       =   40
      TabIndex        =   21
      Top             =   2955
      Width           =   4425
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
      Index           =   4
      Left            =   5235
      MaxLength       =   20
      TabIndex        =   4
      Top             =   795
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
      Index           =   22
      Left            =   1545
      MaxLength       =   1
      TabIndex        =   22
      ToolTipText     =   "Road Code : 1 - Plain Metalled, 2 - Plain Kutcha, 3 - Off Road, 4 - Hilly Metalled, 5 - Hilly Kutcha, 6 - Desert, 7 - Others"
      Top             =   3225
      Width           =   405
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
      Left            =   1545
      MaxLength       =   20
      TabIndex        =   13
      Top             =   2145
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
      Index           =   17
      Left            =   1545
      MaxLength       =   1
      TabIndex        =   17
      ToolTipText     =   "Failure Code : 1 - OE, 2 - REPEAT, 3 - SPARE PARTS"
      Top             =   2685
      Width           =   405
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
      Left            =   1545
      MaxLength       =   1
      TabIndex        =   19
      ToolTipText     =   "Operation Code : 1 - Drive Away, 2 - Long Route, 3 - City Route, 4 - Construction, 5 - Mining, 6 - Forest, 7 - Marine, 8 - Others"
      Top             =   2955
      Width           =   405
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
      Left            =   1545
      MaxLength       =   14
      TabIndex        =   7
      Top             =   1335
      Width           =   1470
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
      Index           =   3
      Left            =   1965
      MaxLength       =   10
      TabIndex        =   3
      Top             =   795
      Width           =   1050
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
      Index           =   8
      Left            =   5235
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1335
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
      Left            =   1545
      MaxLength       =   20
      TabIndex        =   11
      Top             =   1875
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
      Index           =   9
      Left            =   1545
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1605
      Width           =   2175
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
      Index           =   10
      Left            =   5235
      MaxLength       =   8
      TabIndex        =   10
      Top             =   1605
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
      Index           =   16
      Left            =   5235
      MaxLength       =   15
      TabIndex        =   16
      Top             =   2415
      Width           =   1425
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   2025
      Left            =   165
      TabIndex        =   26
      Top             =   4830
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   3572
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
   Begin MSDataGridLib.DataGrid DGJob 
      Height          =   2520
      Left            =   10230
      Negotiate       =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   7155
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
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   3195.213
         EndProperty
      EndProperty
   End
   Begin VB.Label LabelHelp 
      BackColor       =   &H00CFE0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   90
      TabIndex        =   58
      Top             =   3555
      Width           =   11670
   End
   Begin VB.Label lblRoad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "** Uknown **"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00426388&
      Height          =   210
      Left            =   2055
      TabIndex        =   56
      Top             =   3240
      Width           =   930
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "** Uknown **"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00426388&
      Height          =   210
      Left            =   2055
      TabIndex        =   55
      Top             =   2430
      Width           =   930
   End
   Begin VB.Label lblOperation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "** Uknown **"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00426388&
      Height          =   210
      Left            =   2055
      TabIndex        =   54
      Top             =   2970
      Width           =   930
   End
   Begin VB.Label lblFailure 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "** Uknown **"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00426388&
      Height          =   210
      Left            =   2055
      TabIndex        =   53
      Top             =   2700
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Repaired Date"
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
      Left            =   3765
      TabIndex        =   52
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Parts KMS"
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
      Left            =   3765
      TabIndex        =   51
      Top             =   1890
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status Code"
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
      Left            =   105
      TabIndex        =   50
      Top             =   2430
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year Prefix"
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
      Left            =   105
      TabIndex        =   49
      Top             =   540
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Owner Name"
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
      Left            =   3765
      TabIndex        =   48
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Failure"
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
      Index           =   18
      Left            =   105
      TabIndex        =   47
      Top             =   2700
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operation"
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
      Index           =   17
      Left            =   105
      TabIndex        =   46
      Top             =   2970
      Width           =   810
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
      Index           =   15
      Left            =   105
      TabIndex        =   45
      Top             =   1080
      Width           =   1035
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
      Index           =   7
      Left            =   3765
      TabIndex        =   44
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label lblDocCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCard DocID :"
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
      TabIndex        =   43
      Top             =   855
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature of Complaints"
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
      Index           =   13
      Left            =   165
      TabIndex        =   42
      Top             =   3900
      Width           =   1785
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cause of Failure"
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
      Left            =   6015
      TabIndex        =   41
      Top             =   3900
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Dealer "
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
      Left            =   3765
      TabIndex        =   40
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Claim Dt."
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
      Left            =   3765
      TabIndex        =   39
      Top             =   810
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Complaint Date"
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
      Left            =   105
      TabIndex        =   38
      Top             =   2160
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Road"
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
      Index           =   10
      Left            =   105
      TabIndex        =   37
      Top             =   3240
      Width           =   450
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
      TabIndex        =   36
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label lblDocId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DocID"
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
      Left            =   8670
      TabIndex        =   35
      Top             =   855
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Claim Type && No."
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
      Left            =   105
      TabIndex        =   34
      Top             =   810
      Width           =   1425
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
      Left            =   3765
      TabIndex        =   33
      Top             =   1335
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Fitted Date"
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
      Left            =   105
      TabIndex        =   32
      Top             =   1890
      Width           =   1425
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
      Left            =   105
      TabIndex        =   31
      Top             =   1335
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   615
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
      Left            =   8655
      TabIndex        =   30
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Mileage "
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
      Index           =   37
      Left            =   3765
      TabIndex        =   29
      Top             =   1620
      Width           =   1350
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
      Left            =   3765
      TabIndex        =   28
      Top             =   2415
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aggregate No."
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
      Left            =   105
      TabIndex        =   27
      Top             =   1620
      Width           =   1200
   End
End
Attribute VB_Name = "frmWarrantyPCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TAddMode As Boolean
Dim ExitCtrl As Boolean
Dim GridKey As Integer

Dim ForSiteCode As String

Dim MyIndex As Byte
Dim Rst As ADODB.Recordset

Dim Master As ADODB.Recordset
Dim RsJob As ADODB.Recordset

'Text Box (Form)
Private Const YPrefix As Byte = 1
Private Const ClmType As Byte = 2
Private Const ClmNo As Byte = 3
Private Const ClmDate As Byte = 4
Private Const JobNo As Byte = 5
Private Const JobDt As Byte = 6
Private Const VehRegNo As Byte = 7
Private Const Chassis As Byte = 8
Private Const Aggregate As Byte = 9
Private Const CurrKMS As Byte = 10
Private Const SprDate As Byte = 11
Private Const SprKMS As Byte = 12
Private Const CompDate As Byte = 13
Private Const RepairDate As Byte = 14
Private Const StatusCD As Byte = 15
Private Const Model As Byte = 16
Private Const FailureCD As Byte = 17
Private Const OwnerName As Byte = 18
Private Const Operation As Byte = 19
Private Const DCode As Byte = 20
Private Const DNAME As Byte = 21
Private Const Road As Byte = 22
Private Const CNature As Byte = 23
Private Const CCause As Byte = 24

'Fgrid1 Columns
Private Const C_Cust As Byte = 1
Private Const C_Item As Byte = 2
Private Const C_Observ As Byte = 3
Private Const C_Action As Byte = 4

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Dim dd As Byte
    If KeyCode = vbKeyF2 Or KeyCode = vbKeyF3 Or KeyCode = vbKeyF4 Or _
        (KeyCode = 70 And Shift = 2) Or (KeyCode = 80 And Shift = 2) Or _
        (KeyCode = 83 And Shift = 2) Or KeyCode = vbKeyEscape Or _
        KeyCode = vbKeyF5 Or KeyCode = vbKeyF10 Or KeyCode = vbKeyHome Or _
        KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyEnd Then
        TopCtrl1.TopKey_Down KeyCode, Shift
    End If
    If KeyCode <> vbKeyF10 Then
        If TopCtrl1.PrvKeyCode = vbKeyEscape Then
            TopCtrl1.PrvKeyCode = 0
        Else
            TopCtrl1.PrvKeyCode = KeyCode
        End If
    End If
    Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
Dim SrNo As Integer
    
    TopCtrl1.Tag = PubUParam 'UserPermission(Me.Name)
    ForSiteCode = PubSiteCode
    Call BlankText
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select JW1.Year_Prefix+RIGHT(SPACE(2)+JW1.Claim_Type,2)+RIGHT(SPACE(10)+JW1.Claim_No,10) AS SEARCHCODE,JW1.Year_Prefix+RIGHT(SPACE(2)+JW1.Claim_Type,2)+RIGHT(SPACE(10)+JW1.Claim_No,10) AS CODE, JW1.Div_Code,Jw1.Site_Code,Jw1.Year_Prefix,Jw1.Claim_No,Jw1.Claim_Type,jw1.Claim_date,jw1.job_DocId,jw1.cmpl_Date,jw1.Repair_Date,jw1.atkmsHrs,jw1.sprSoldDate,jw1.sprkms,jw1.Aggregate_No,jw1.failurecode,jw1.status_code,jw1.operationcode,jw1.roadcode,jw1.cmpl_detail,jw1.causeoffailure,jw1.wbill_no,jc.Job_no,jc.job_date,JC.cardno,hc.model,hc.name,hc.dealer_code,hc.chassis,hc.engine,hc.regno,amd_dealer.D_name " _
                & "FROM ((JOB_WARR1 AS JW1 left join Job_Card as Jc on Jw1.Job_DocId=Jc.DocId) left join Hiscard as Hc on Jc.CardNo=hc.CardNo) left join amd_dealer on hc.dealer_code=amd_dealer.d_code where left(JW1.site_code,1)='" & PubSiteCode & "' AND left(JW1.site_code,1)='" & PubSiteCode & "' order by JW1.CLAIM_TYPE,JW1.CLAIM_NO", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsJob = New ADODB.Recordset
    With RsJob
        .CursorLocation = adUseClient
        .Open "select Jc.DocId AS CODE,  " & cCStr("JC.Job_No") & " as Job_No, HC.Model, HC.RegNo, HC.Chassis, HC.Engine , HC.VehSerialNo, HC.Name, Jc.DocId,JC.Govt_YN,Jc.OpenRemarks, JC.Job_Date, JC.JobCloseDate,jc.cardno,Hc.Dealer_Code,amd_dealer.d_name from (job_card as JC left Join Hiscard as HC on JC.CardNo=HC.CardNo) left join amd_dealer on hc.dealer_code=amd_dealer.d_code where left(JC.site_code,1)='" & PubSiteCode & "' order by JC.docID", GCn, adOpenDynamic, adLockOptimistic
    End With
    Set DGJob.DataSource = RsJob
    RsJob.Sort = "code"
    
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
    Set RsJob = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    
    lblRoad = FxRoad(0)
    lblStatus = FxStatus(0)
    lblFailure = FxFailure(0)
    lblOperation = FxOperation(0)
    
    Txt(CNature).Locked = False
    Txt(CCause).Locked = False
    
    Txt(YPrefix).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim vBook As Variant

    If Master!wbill_no <> 0 Then
        MsgBox "Warranty Bill No. " & Master!wbill_no & " prepared against this Claim No.," & vbCrLf & " Can't Delete", vbInformation, "Validation"
        Exit Sub
    End If
    
    If GCn.Execute("select Claim_No from job_warr3 where Div_code='" & PubDivCode & "' and Site_code='" & PubSiteCode & "' and Claim_Type='" & Txt(ClmType).TEXT & "' and claim_no='" & Txt(ClmNo).TEXT & "' and Year_Prefix='" & Txt(YPrefix).TEXT & "'").RecordCount > 0 Then
        MsgBox "Parts Feeded against this Claim No.,Can't Delete", vbInformation, "Validation"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure To Delete Entry? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        vBook = Master.AbsolutePosition
        GCn.BeginTrans
        
        GCn.Execute ("DELETE * FROM JOB_WARR2 WHERE Div_code='" & Master!Div_Code & "' and Site_code='" & Master!Site_Code & "' and Claim_Type='" & Txt(ClmType).TEXT & "' and claim_no='" & Txt(ClmNo).TEXT & "' and Year_Prefix='" & Txt(YPrefix).TEXT & "'")
        GCn.Execute ("DELETE * FROM JOB_WARR1 WHERE Div_code='" & Master!Div_Code & "' and Site_code='" & Master!Site_Code & "' and Claim_Type='" & Txt(ClmType).TEXT & "' and claim_no='" & Txt(ClmNo).TEXT & "' and Year_Prefix='" & Txt(YPrefix).TEXT & "'")
        
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
    LblDiv.CAPTION = "Division : " & Master!Div_Code
    LblSite.CAPTION = "Site Code : " & Master!Site_Code
    lblDocId.CAPTION = Master!job_docid
    
    Txt(YPrefix).Enabled = False
    Txt(ClmType).Enabled = False
    Txt(ClmNo).Enabled = False
    
    Txt(JobNo).Enabled = False
    Txt(JobDt).Enabled = False
    Txt(VehRegNo).Enabled = False
    Txt(Chassis).Enabled = False
    Txt(OwnerName).Enabled = False
    Txt(Model).Enabled = False
    Txt(DCode).Enabled = False
    Txt(DNAME).Enabled = False
    
    Txt(CNature).Locked = False
    Txt(CCause).Locked = False
    Txt(ClmDate).SetFocus
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
    Dim AddFlg As String, MyDocId As String
    On Error GoTo errlbl

    Grid_Hide
    If Len(Txt(YPrefix).TEXT) < 2 Then
        MsgBox "Please specify Two digits for Year", vbInformation, "Validation"
        Txt(YPrefix).SetFocus
        Exit Sub
    End If
    If IsValid(Txt(ClmType), "Claim Type") = False Then Exit Sub
    If IsValid(Txt(ClmNo), "Claim No.") = False Then Exit Sub
    If IsValid(Txt(ClmDate), "Claim Date") = False Then Exit Sub
    If IsValid(Txt(JobNo), "Job Card No.") = False Then Exit Sub
    If IsValid(Txt(Aggregate), "Aggregate No.") = False Then Exit Sub
    If IsValid(Txt(CompDate), "Complaint Date") = False Then Exit Sub
    If IsValid(Txt(RepairDate), "Repaired Date") = False Then Exit Sub
    
    If IsValid(Txt(ClmDate), "Claim Date") = False Then Exit Sub
    
    GCn.BeginTrans
    mTrans = True
    
    MyDocId = Txt(YPrefix) & Right(Space(2) + Txt(ClmType), 2) & Right(Space(10) + Txt(ClmNo), 10)
    
    Select Case TopCtrl1.TopText2
        Case "Add"
            AddFlg = "A"
            If GCn.Execute("select Claim_No from job_warr1 where Div_code='" & PubDivCode & "' and Site_code='" & PubSiteCode & "' and Claim_Type='" & Txt(ClmType).TEXT & "' and claim_no='" & Txt(ClmNo).TEXT & "' and Year_Prefix='" & Txt(YPrefix).TEXT & "'").RecordCount > 0 Then
                MsgBox "Duplicate Claim No.", vbInformation, "Validation"
                Txt(ClmType).SetFocus
                Exit Sub
            End If
            GSQL = "insert into Job_Warr1(" _
                    & "Div_Code,Site_Code,Year_Prefix,Claim_no,Claim_Type, " _
                    & "Claim_Date,Job_DocId,Cmpl_Date,Repair_Date,AtKmsHrs," _
                    & "SprSoldDate,SprKMS,Aggregate_no,FailureCode,Status_Code,OperationCode," _
                    & "RoadCode,Cmpl_Detail,CauseofFailure,ClaimU_Name,ClaimU_EntDt," _
                    & "ClaimU_AE,U_Name,U_EntDt,U_AE)" _
                    & " values(" _
                    & "'" & PubDivCode & "','" & PubSiteCode & "','" & Txt(YPrefix) & "','" & Txt(ClmNo) & "','" & Txt(ClmType) & "'," _
                    & "" & ConvertDate(Txt(ClmDate)) & ",'" & Txt(JobNo).Tag & "'," & ConvertDate(Txt(CompDate)) & "," & ConvertDate(Txt(RepairDate)) & "," & Val(Txt(CurrKMS)) & "," _
                    & "" & ConvertDate(Txt(SprDate)) & "," & Val(Txt(SprKMS).TEXT) & ",'" & Txt(Aggregate) & "','" & Txt(FailureCD) & "','" & Txt(StatusCD) & "','" & Txt(Operation) & "'," _
                    & "'" & Txt(Road) & "','" & Txt(CNature).TEXT & "','" & Txt(CCause) & "','" & pubUName & "'," & ConvertDate(PubServerDate) & "," _
                    & "'A','" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
            GCn.Execute GSQL
        Case "Edit"
            AddFlg = "E"
            GCn.Execute ("DELETE * FROM JOB_WARR2 WHERE Div_code='" & Master!Div_Code & "' and Site_code='" & Master!Site_Code & "' and Claim_Type='" & Txt(ClmType).TEXT & "' and claim_no='" & Txt(ClmNo).TEXT & "' and Year_Prefix='" & Txt(YPrefix).TEXT & "'")
            
            GSQL = "Update Job_Warr1 set Claim_Date=" & ConvertDate(Txt(ClmDate)) & ",Cmpl_Date=" & ConvertDate(Txt(CompDate)) & ",Repair_Date=" & ConvertDate(Txt(RepairDate)) & ",AtKmsHrs=" & Val(Txt(CurrKMS)) & "," _
                    & "SprSoldDate=" & ConvertDate(Txt(SprDate)) & ",SprKMS=" & Val(Txt(SprKMS).TEXT) & ",Aggregate_no='" & Txt(Aggregate) & "',FailureCode='" & Txt(FailureCD) & "',Status_Code='" & Txt(StatusCD) & "'," _
                    & "RoadCode='" & Txt(Road) & "',OperationCode='" & Txt(Operation) & "',Cmpl_Detail='" & Txt(CNature).TEXT & "',CauseofFailure='" & Txt(CCause) & "',ClaimU_Name='" & pubUName & "',ClaimU_EntDt=" & ConvertDate(PubServerDate) & "," _
                    & "ClaimU_AE='E',U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='E'"
            
            GCn.Execute GSQL
    End Select

    SrNo = 1
    For I = 1 To FGrid1.Rows - 1
        If Trim(FGrid1.TextMatrix(I, C_Cust) + FGrid1.TextMatrix(I, C_Item) + FGrid1.TextMatrix(I, C_Observ) + FGrid1.TextMatrix(I, C_Action)) <> "" Then
            GSQL = "insert into Job_WARR2(" _
                    & "Div_code,Site_code,Year_Prefix,Claim_no,claim_type," _
                    & "Srl_no,Cust_Complaint,Items,Observation,Corrective_Action," _
                    & "U_Name, U_EntDt, U_AE) " _
                    & " values(" _
                    & "'" & PubDivCode & "','" & PubSiteCode & "','" & Txt(YPrefix).TEXT & "','" & Txt(ClmNo).TEXT & "','" & Txt(ClmType).TEXT & "'," _
                    & "" & SrNo & ",'" & FGrid1.TextMatrix(I, C_Cust) & "','" & FGrid1.TextMatrix(I, C_Item) & "','" & FGrid1.TextMatrix(I, C_Observ) & "','" & FGrid1.TextMatrix(I, C_Action) & "'," _
                    & "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & AddFlg & "')"
            GCn.Execute GSQL
            SrNo = SrNo + 1
        End If
    Next I
    
    GCn.CommitTrans
    mTrans = False
    
    Master.Requery
    Call UpdRequery
    
    Master.FIND "Code = '" & MyDocId & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub

errlbl:
    If mTrans = True Then
        GCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Ctrl_GetFocus Txt(Index)
    GetHelp Index, True
    TxtGrid1(0).Visible = False
    Grid_Hide
    MyIndex = Index
    Select Case MyIndex
        Case JobNo
            DGridColSwap DGJob, 0
            RsJob.Sort = "JOB_NO"
            If RsJob.RecordCount = 0 Or RsJob.EOF = True Or RsJob.BOF = True Or Txt(Index).TEXT = "" Then Exit Sub
            If Txt(Index).Tag <> "" And Txt(Index).Tag <> RsJob!Code Then
                RsJob.FIND ("JOB_NO='" & Txt(Index).TEXT & "'")
            End If
        Case Chassis
            DGridColSwap DGJob, 1
            RsJob.Sort = "CHASSIS"
            If RsJob.RecordCount = 0 Or RsJob.EOF = True Or RsJob.BOF = True Or Txt(Index).TEXT = "" Then Exit Sub
            If Txt(Index).Tag <> "" And Txt(Index).Tag <> RsJob!Code Then
                RsJob.FIND ("CHASSIS='" & Txt(Index).TEXT & "'")
            End If
        Case VehRegNo
            DGridColSwap DGJob, 2
            RsJob.Sort = "REGNO"
            If RsJob.RecordCount = 0 Or RsJob.EOF = True Or RsJob.BOF = True Or Txt(Index).TEXT = "" Then Exit Sub
            If Txt(Index).Tag <> "" And Txt(Index).Tag <> RsJob!Code Then
                RsJob.FIND ("REGNO='" & Txt(Index).TEXT & "'")
            End If
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
    End If
    Select Case Index
        Case JobNo
            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 1
        Case VehRegNo
            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 3
        Case Chassis
            DGridTxtKeyDown DGJob, Txt, Index, RsJob, KeyCode, False, 4
    End Select
    If DGJob.Visible = False Then
        '' KEY DOWN
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
            If Index <> CNature And Index <> CCause Then
                Ctrl_DownKeyDown KeyCode, Shift
            End If
        End If
        
        ' KEY UP
        If TopCtrl1.TopText2 = "Add" Then
            If Index <> YPrefix And Index <> CNature And Index <> CCause Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        ElseIf TopCtrl1.TopText2 = "Edit" Then
            If Index <> ClmDate And Index <> CNature And Index <> CCause Then
                If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
End Sub

Private Sub Txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
    Select Case Index
        Case YPrefix
            Call NumPress(Txt(Index), KeyAscii, 2, 0)
        Case CurrKMS, SprKMS
            Call NumPress(Txt(Index), KeyAscii, 8, 0)
        Case StatusCD, FailureCD, Operation, Road
            Call NumPress(Txt(Index), KeyAscii, 1, 0)
        Case JobNo
            Call NumPress(Txt(Index), KeyAscii, 8, 0)
            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "JOB_NO"
        Case VehRegNo
            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "Regno"
        Case Chassis
            DGridTxtKeyPress Txt, Index, RsJob, KeyAscii, "Chassis"
    End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
    GetHelp Index, True
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case JobNo, VehRegNo, Chassis
            If Txt(Index).Tag <> "" Then
                RsJob.Sort = "CODE"
                RsJob.FIND ("CODE='" & Txt(Index).Tag & "'")
            End If
            If RsJob.BOF = True Or RsJob.EOF = True Then Exit Sub
            lblDocId = RsJob!DocID
            Call History_Field
            
        Case YPrefix
            If Len(Txt(Index).TEXT) < 2 Then
                MsgBox "Year Prefix should have 2 digits of Year", vbInformation, "Validation"
                Txt(Index).SetFocus
                Exit Sub
            End If
        
        Case ClmDate, SprDate, CompDate, RepairDate
            Txt(Index).TEXT = RetDate(Txt(Index))
        
        Case StatusCD
            lblStatus = FxStatus(Val(Txt(StatusCD).TEXT))
            lblStatus.Refresh
        
        Case FailureCD
            lblFailure = FxFailure(Val(Txt(FailureCD).TEXT))
            lblFailure.Refresh
        
        Case Operation
            lblOperation = FxOperation(Val(Txt(Operation).TEXT))
            lblOperation.Refresh
        
        Case Road
            lblRoad = FxRoad(Val(Txt(Road).TEXT))
            lblRoad.Refresh
    End Select
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
    
    FGrid1.Rows = 1
    FGrid1.AddItem FGrid1.Rows
    FGrid1.FixedRows = 1
End Sub

Private Sub MoveRec()
Dim rs As Recordset
Dim mVor As String
Dim I As Integer
On Error GoTo error1
    If Master.RecordCount > 0 Then
        
        Txt(CNature).Enabled = True
        Txt(CNature).Locked = True
        Txt(CCause).Enabled = True
        Txt(CCause).Locked = True

        LblDiv.CAPTION = "Division : " & left(Master!Div_Code, 1)
        LblSite.CAPTION = "Site Code : " & Master!Site_Code
        lblDocId.CAPTION = Master!job_docid
        
        Txt(YPrefix).TEXT = Master!Year_Prefix
        Txt(ClmType).TEXT = Master!claim_type
        Txt(ClmNo).TEXT = Master!claim_no
        Txt(ClmDate).TEXT = Master!Claim_Date
        Txt(JobNo).TEXT = Master!Job_No
        Txt(JobDt).TEXT = Master!Job_Date
        Txt(VehRegNo).TEXT = IIf(IsNull(Master!RegNo), "", Master!RegNo)
        Txt(Chassis).TEXT = IIf(IsNull(Master!Chassis), "", Master!Chassis)
        Txt(Aggregate).TEXT = IIf(IsNull(Master!Aggregate_No), "", Master!Aggregate_No)
        Txt(CurrKMS).TEXT = IIf(IsNull(Master!AtKMsHrs), 0, Master!AtKMsHrs)
        Txt(SprDate).TEXT = IIf(IsNull(Master!sprsolddate), "", Master!sprsolddate)
        Txt(SprKMS).TEXT = Master!SprKMS
        Txt(CompDate).TEXT = IIf(IsNull(Master!Cmpl_Date), "", Master!Cmpl_Date)
        Txt(RepairDate).TEXT = IIf(IsNull(Master!Repair_Date), "", Master!Repair_Date)
        Txt(StatusCD).TEXT = Val(Master!Status_Code)
        Txt(FailureCD).TEXT = Val(Master!FailureCode)
        Txt(Operation).TEXT = Val(Master!OperationCode)
        Txt(Road).TEXT = Val(Master!RoadCode)
        Txt(Model).TEXT = Master!Model
        Txt(OwnerName).TEXT = Master!Name
        Txt(DCode).TEXT = Master!dealer_code
        Txt(DNAME).TEXT = XNull(Master!D_Name)
        
        Txt(CNature).TEXT = Master!Cmpl_Detail
        Txt(CCause).TEXT = Master!CauseOfFailure
        
        lblRoad = FxRoad(IIf(IsNull(Master!RoadCode), 0, Val(Master!RoadCode)))
        lblStatus = FxStatus(IIf(IsNull(Master!Status_Code), 0, Val(Master!Status_Code)))
        lblFailure = FxFailure(IIf(IsNull(Master!FailureCode), 0, Val(Master!FailureCode)))
        lblOperation = FxOperation(IIf(IsNull(Master!OperationCode), 0, Val(Master!OperationCode)))
        
        Call Fill_Grid
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
        .height = FGrid1.RowHeight(1) * 10
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 5
        
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 400
        
        .TextMatrix(0, C_Cust) = "Complaint Reported By Driver"
        .ColAlignment(C_Cust) = flexAlignLeftCenter
        .ColWidth(C_Cust) = 3100
        
        .TextMatrix(0, C_Item) = "Item Checked & Make"
        .ColAlignment(C_Item) = flexAlignLeftCenter
        .ColWidth(C_Item) = 2900

        .TextMatrix(0, C_Observ) = "Observation"
        .ColAlignment(C_Observ) = flexAlignLeftCenter
        .ColWidth(C_Observ) = 3700
    
        .TextMatrix(0, C_Action) = "Corrective Action"
        .ColAlignment(C_Action) = flexAlignLeftCenter
        .ColWidth(C_Action) = 2000
    End With
    DGJob.width = Me.width - 60: DGJob.left = Me.left + 30: DGJob.top = FGrid1.top: DGJob.height = FGrid1.height
End Sub

Public Sub Disp_Text(Enb As Boolean)
Dim I As Integer
    
    For I = 1 To Txt.Count
        Txt(I).Enabled = Enb
    Next
    
    For I = 1 To Txt.Count
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
    
    Txt(Model).Enabled = False
    Txt(OwnerName).Enabled = False
    Txt(DCode).Enabled = False
    Txt(DNAME).Enabled = False
    Txt(JobDt).Enabled = False
    
End Sub

Private Sub Grid_Hide()
    If DGJob.Visible = True Then DGJob.Visible = False
End Sub

Private Sub UpdRequery()
    RsJob.Requery
End Sub

Private Sub History_Field()
    Txt(VehRegNo).Tag = XNull(RsJob!Code)
    Txt(Chassis).Tag = XNull(RsJob!Code)
    Txt(JobNo).Tag = XNull(RsJob!Code)
    
    Txt(JobNo).TEXT = XNull(RsJob!Job_No)
    Txt(JobDt).TEXT = RsJob!Job_Date
    Txt(VehRegNo).TEXT = XNull(RsJob!RegNo)
    Txt(Chassis).TEXT = XNull(RsJob!Chassis)
    Txt(Model).TEXT = XNull(RsJob!Model)
    Txt(OwnerName).TEXT = XNull(RsJob!Name)
    Txt(CurrKMS).TEXT = XNull(RsJob!AtKMsHrs)
    Txt(CompDate).TEXT = XNull(RsJob!Job_Date)
    Txt(RepairDate).TEXT = XNull(RsJob!JobCloseDate)

    Txt(DCode).TEXT = XNull(RsJob!dealer_code)
    Txt(DNAME).TEXT = XNull(RsJob!dealer_name)
End Sub

Private Sub FGrid1_Click()
    TxtGrid1(0).Visible = False
End Sub

Private Sub FGrid1_DblClick()
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    GridDblClick Me, FGrid1, TxtGrid1, 0
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
    TxtGrid1(0).Visible = False
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
        If MsgBox("Save Entry ?", vbInformation + vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave: Exit Sub
        FGrid1.SetFocus
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid1.Tag = FGrid1.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = ""
    End If
    If KeyCode = vbKeyReturn Then
        GridDblClick Me, FGrid1, TxtGrid1, 0
    End If
    TAddMode = False
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
    Get_Text Me, FGrid1, TxtGrid1, 0, False, KeyAscii
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
    TxtGrid1(0).Visible = False
End Sub

Private Sub FGrid1_Validate(Cancel As Boolean)
    FGrid1.CellBackColor = CellBackColLeave
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
On Error GoTo ELoop
If ExitCtrl = False Then Exit Sub
    Ctrl_GetFocus TxtGrid1(0)
    Grid_Hide
    FGrid1.CellBackColor = CellBackColLeave
    TxtGrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    Select Case FGrid1.Col
        Case C_Cust
            TxtGrid1(0).MaxLength = 30
        Case C_Item
            TxtGrid1(0).MaxLength = 25
        Case C_Observ
            TxtGrid1(0).MaxLength = 60
        Case C_Action
            TxtGrid1(0).MaxLength = 25
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        TxtGrid1(0).TEXT = TxtGrid1(0).Tag
        TxtGrid1(0).Visible = False
        FGrid1.SetFocus
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Then
        If TxtGrid1Leave = True Then
             GridTxtDown FGrid1, TxtGrid1, Index, KeyCode, TAddMode, 4
        Else
            TxtGrid1_LostFocus 0
            TxtGrid1(0).SetFocus
        End If
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
End Sub

Private Sub TxtGrid1_LostFocus(Index As Integer)
On Error GoTo ELoop
    If ExitCtrl = False Then Exit Sub
    Ctrl_validate TxtGrid1(Index)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid1_Validate(Index As Integer, Cancel As Boolean)
Dim I As Integer
On Error GoTo ELoop
    FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = TxtGrid1(Index).TEXT
    TxtGrid1(0).MaxLength = 60
    Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid1Leave() As Boolean
Dim I As Integer
    FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col) = TxtGrid1(0).TEXT
    TxtGrid1(0).MaxLength = 60
    
    TxtGrid1(0).Visible = False
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
    
    GSQL = "Select JW2.* From Job_warr2 as jw2 Where Jw2.Claim_Type='" & Txt(ClmType) & "' and Jw2.Claim_no='" & Txt(ClmNo) & "' and Jw2.Year_Prefix='" & Txt(YPrefix) & "' order by Jw2.Srl_no"
    
    MyRst.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    I = 1
    If MyRst.RecordCount > 0 Then
        Do Until MyRst.EOF
            FGrid1.AddItem ""
            With FGrid1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, C_Cust) = MyRst!Cust_Complaint
                .TextMatrix(I, C_Item) = MyRst!Items
                .TextMatrix(I, C_Observ) = XNull(MyRst!Observation)
                .TextMatrix(I, C_Action) = XNull(MyRst!Corrective_Action)
            End With
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

Private Sub GetHelp(Index As Integer, Omit As Boolean)
If Omit = False Then LabelHelp.CAPTION = "": Exit Sub
    Select Case Index
        Case StatusCD, FailureCD, Operation, Road
            Select Case Index
                Case StatusCD
                    LabelHelp.CAPTION = "Press 0- >Unknown,1- >Drive Away,2- >Sold"
                Case FailureCD
                    LabelHelp.CAPTION = "Press 0- >Unknown,1- >OE,2- >Repeat,3- >Spare Parts"
                Case Operation
                    LabelHelp.CAPTION = "Press 0- >Unknown,1- >Drive Away,2- >Long Route,3- >City Route,4- >Construction,5- >Mining,6- >Forest,7- >Marine,8- >Others"
                Case Road
                    LabelHelp.CAPTION = "Press 0- >Unknown,1- >Plain Metalled,2- >Plain Kutcha,3- >Off Road,4- >Hilly Metalled,5- >Killy Kutcha,6- >Desert,7- >Others"
            End Select
        Case Else
            LabelHelp.CAPTION = ""
    End Select
End Sub

