VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmPurOth 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   Caption         =   "Other Purchase Entry"
   ClientHeight    =   8115
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
   ScaleHeight     =   8115
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   2205
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1680
      Width           =   4800
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
      Left            =   2205
      MaxLength       =   50
      TabIndex        =   10
      Top             =   2220
      Width           =   4800
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   2205
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1950
      Width           =   4800
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
      Index           =   11
      Left            =   2205
      TabIndex        =   11
      Top             =   2490
      Width           =   1320
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
      Index           =   1
      Left            =   10365
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1590
      Width           =   975
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
      Left            =   2205
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1410
      Width           =   1020
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   4620
      TabIndex        =   23
      Top             =   6930
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   300
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   30
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
      BackColor       =   &H00FBFBFB&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   1605
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "aaaa"
      Top             =   870
      Width           =   4560
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
      Left            =   9540
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1320
      Width           =   1245
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
      Index           =   0
      Left            =   9540
      MaxLength       =   21
      TabIndex        =   1
      Top             =   510
      Width           =   2100
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   4920
      MaxLength       =   12
      TabIndex        =   7
      Text            =   "29-APR-2002"
      Top             =   1140
      Width           =   1245
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
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1140
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
      Index           =   2
      Left            =   9540
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1050
      Width           =   1245
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   4935
      Left            =   3900
      Negotiate       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5235
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777152
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
      Caption         =   "Party Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGForm 
      Height          =   3150
      Left            =   -1425
      Negotiate       =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4980
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   5556
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777152
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
      Caption         =   "Tax Form Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Form Description"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGDrAc 
      Height          =   4935
      Left            =   7860
      Negotiate       =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4785
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777152
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
      ColumnCount     =   1
      BeginProperty Column00 
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
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
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   3
      Left            =   2085
      TabIndex        =   44
      Top             =   1665
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
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   4
      Left            =   705
      TabIndex        =   42
      Top             =   1680
      Width           =   870
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   25
      Left            =   2085
      TabIndex        =   40
      Top             =   1410
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   22
      Left            =   2085
      TabIndex        =   39
      Top             =   2490
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
      Height          =   255
      Index           =   0
      Left            =   2085
      TabIndex        =   38
      Top             =   1950
      Width           =   45
   End
   Begin VB.Label Lbl3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   705
      TabIndex        =   36
      Top             =   2220
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   2085
      TabIndex        =   35
      Top             =   2220
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dr A/c Name"
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
      Left            =   705
      TabIndex        =   34
      Top             =   1950
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Bill Amount"
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
      Index           =   41
      Left            =   705
      TabIndex        =   33
      Top             =   2520
      Width           =   1290
   End
   Begin VB.Label LblCancel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Cancelled*"
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
      Left            =   6390
      TabIndex        =   32
      Top             =   930
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   1470
      Left            =   8280
      Top             =   435
      Width           =   3465
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VPrefix"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   9540
      TabIndex        =   30
      Top             =   1605
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill  No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   1
      Left            =   8385
      TabIndex        =   29
      Top             =   1605
      Width           =   630
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   92
      Left            =   9405
      TabIndex        =   28
      Top             =   1605
      Width           =   45
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division        :"
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
      Left            =   8385
      TabIndex        =   27
      Top             =   780
      Width           =   1065
   End
   Begin VB.Label LblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code    :"
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
      Height          =   255
      Left            =   10140
      TabIndex        =   26
      Top             =   780
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   44
      Left            =   705
      TabIndex        =   25
      Top             =   1410
      Width           =   1230
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
      Index           =   23
      Left            =   9405
      TabIndex        =   21
      Top             =   510
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOC ID"
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
      Height          =   255
      Index           =   42
      Left            =   8370
      TabIndex        =   20
      Top             =   510
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suplier Invoice No. && Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   5
      Left            =   705
      TabIndex        =   19
      Top             =   1155
      Width           =   2130
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   3
      Left            =   465
      TabIndex        =   18
      Top             =   855
      Width           =   690
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   88
      Left            =   3075
      TabIndex        =   17
      Top             =   1140
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   90
      Left            =   1485
      TabIndex        =   16
      Top             =   870
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   91
      Left            =   9405
      TabIndex        =   15
      Top             =   1050
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
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   93
      Left            =   9405
      TabIndex        =   14
      Top             =   1335
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Credit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   0
      Left            =   8385
      TabIndex        =   13
      Top             =   1335
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   8385
      TabIndex        =   12
      Top             =   1050
      Width           =   690
   End
End
Attribute VB_Name = "frmPurOth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BackColorSelEnter As String = &HF8D7FD
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const PurCashVType As String = "SXPTC"
Private Const PurCrVType As String = "SXPTR"

Dim RsParty As ADODB.Recordset
Dim rsForm As ADODB.Recordset
Dim rsForm31 As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim rsCtrlAc As ADODB.Recordset

Dim FirmAddFlag As Byte
Dim DocID As String * 21
Dim mVtype As String
Dim VoucherEditFlag As Boolean
Dim vPrefix As String
Dim mPartyType As Byte         'Used to Detect Part Rate in GetRate Function

Private Const TxtDocID As Byte = 0
Private Const SerialNo As Byte = 1
Private Const VDate As Byte = 2
Private Const VType As Byte = 3
Private Const Party As Byte = 4
Private Const SuppChlNo As Byte = 5
Private Const SuppChlDate As Byte = 6
Private Const LC As Byte = 7
Private Const FormType As Byte = 8
Private Const DrAcCode As Byte = 9
Private Const Remark As Byte = 10
Private Const NetAmt As Byte = 11
'Private Const PermitType As Byte = 12

Dim TAddMode As Boolean
Dim ListArray As Variant
Dim mListItem As ListItem

Private Sub Form_Activate()
Dim UnLoadFrm As Boolean, MsgStr$
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
    Call TopCtrl1_eRef
End If
If rsCtrlAc.RecordCount <= 0 Then
    MsgStr = "No Records in Spare A/c Controls"
    UnLoadFrm = True
End If
If rsCtrlAc!SprCash_Ac = "" Then
    MsgStr = "Please Fill Spare Purchase "
    UnLoadFrm = True
End If
'EOF Spare A/c control checking
If UnLoadFrm Then
    MsgBox "Spare Purchase Entry Loading Aborted !" & vbCrLf & MsgStr & " A/c Controls through Utility Menu", vbInformation, "Validation"
    Unload Me
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
Dim i As Byte
    TopCtrl1.Tag = PubUParam: WinSetting Me: Ini_Grid
'    Label3(4) = PubForm31Caption
'    Label3(24) = PubForm31Caption & " No."
    mVtype = PurCashVType
    txt(VDate).Tag = PubLoginDate
    
    'A/c Pstong Control Checking
    Set rsCtrlAc = New ADODB.Recordset
    rsCtrlAc.CursorLocation = adUseClient
    'CSSprAc=Temp Sale A/c
    rsCtrlAc.Open "Select EntryTax_Ac,SprCash_Ac From AcControls where Div_Code='" & PubDivCode & "'", GCnFaS, adOpenDynamic, adLockOptimistic
    'eof checking
    
'    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Party_Type from SubGroup " & _
'        "left join [" & PubSFADataPath & "].AcGroup on SubGroup.GroupCode=AcGroup.GroupCode " & _
'        "Where FirmCode = '" & PubFirmCode & "' and " & _
'        "left(AcGroup.MainGrCode,6) not in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
'        "order by SubGroup.name"
'    Set RsDrAc = New ADODB.Recordset
'    RsDrAc.CursorLocation = adUseClient
'    RsDrAc.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
'    Set DGDrAc.DataSource = RsDrAc

    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select DocID as SearchCode,DocID  from Sp_Purch  " & _
        "where left(DocID,1)='" & PubDivCode & "' and v_type in ('" & PurCashVType & "','" & PurCrVType & "') Order By V_Date Desc, DocID desc", GCn, adOpenDynamic, adLockOptimistic
    
    Set rsForm = New ADODB.Recordset
    rsForm.CursorLocation = adUseClient
    rsForm.Open "Select T.Form_Code as Code,T.Form_Desc As Name,T.Tax_Per,T.Tax_Sur_Per,T1.Tax_Ac_Code,T1.Sur_Ac_Code,T1.PurSal_Ac_Code " & _
        "From TaxForms as T left join TaxFormsAc as T1 on T.Form_Code&'" & PubDivCode & "'=T1.Form_Code&T1.Div_Code " & _
        "Where Spare_YN=1 and Trn_Type='Purchase' Order by Form_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGForm.DataSource = rsForm
        
    Set rsForm31 = New ADODB.Recordset
    With rsForm31
        .CursorLocation = adUseClient
        .Open "SELECT TaxForms.Form_Code as code ,TaxForms.form_Desc as name  FROM TaxForms where Spare_YN = 1 and trn_Type = 'Permit' order by  TaxForms.form_Desc", GCn, adOpenDynamic, adLockOptimistic
    End With

    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
'    RsParty.Open "select SubGroup.SubCode as code,SubGroup.NAME,Party_Type from SubGroup  Where firmCode = '" & PubFirmCode & "' and Nature='Supplier'  order by SubGroup.name", GCn, adOpenDynamic, adLockOptimistic
    GSQL = "select SubGroup.SubCode as code,SubGroup.NAME,Party_Type from SubGroup " & _
        "left join " & FaTable("AcGroup") & " on SubGroup.GroupCode=AcGroup.GroupCode " & _
        "Where  " & _
        "left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        " and SubGroup.AliasYN<>'Y' and AcGroup.AliasYN='N' order by SubGroup.name"
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsParty = Nothing
Set rsForm = Nothing
Set Master = Nothing
Set mListItem = Nothing
End Sub

Private Sub ListView_Click()
txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
txt(Val(ListView.Tag)).SetFocus
FrmList.Visible = False
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim i As Integer
    
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    LblVPrefix.CAPTION = ""
    DocID = ""
    txt(TxtDocID).Enabled = False
    mPartyType = 0
    txt(VDate) = txt(VDate).Tag
    txt(VDate).SetFocus
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
Dim i As Integer, mTrans As Boolean
Dim LedgAry(1) As LedgRec, mResult As Byte, MsgStr$, mTitle$

If GCn.Execute("Select CancelYN from SP_Purch where DocID='" & Master!SearchCode & "'").Fields(0).Value = 1 Then
    MsgStr = "Are You Sure To Delete This ? "
    mTitle = "Delete Entry!"
Else
    MsgStr = "Are You Sure To Cancel This ? "
    mTitle = "Cancel Entry!"
End If
If MsgBox(MsgStr, vbYesNo + vbCritical + vbDefaultButton2, mTitle) = vbYes Then
    GCn.BeginTrans
    GCnFaS.BeginTrans
    mTrans = True
    If GCn.Execute("Select CancelYN from SP_Purch where DocID='" & Master!SearchCode & "'").Fields(0).Value = 1 Then
        GCn.Execute ("delete from Sp_Purch where docId = '" & Master!DocID & "'")
    Else
        GCn.Execute ("update sp_purch  set " & _
            " CancelYN=1,Tot_Amt=0,Tot_Disc_Amt= 0,Tot_Ord_DiscAmt=0,SprAmt=0,OilAmt=0,Tot_Goods_Value=0," & _
            " Tax_Amt=0,Addition =0,Deduction=0,NET_AMT =0,U_Name='" & pubUName & "', U_EntDt=#" & PubServerDate & "#, U_AE='E' " & _
            " where DocId = '" & txt(TxtDocID) & "'")
    End If
    '*********
    'Unpost Ledger a/c
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, txt(TxtDocID))
    If mResult <> 1 Then MsgBox "Error in Ledger UnPosting", vbOKOnly, "Validation"
    'Unposting of Ledger completed
    GCnFaS.CommitTrans
    GCn.CommitTrans
    mTrans = False
    Master.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
Exit Sub
eloop1:
    If mTrans Then GCnFaS.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    txt(Party).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
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

Private Sub TopCtrl1_eCancel()
Dim i As Integer
On Error GoTo ErrorLoop
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
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
'SELECT SP_Purch.Case_Mark, SP_Purch.RoadPermit_No, SP_Purch.Addition, SP_Purch.Deduction, SP_Purch.Tot_Disc_Amt, SP_Purch.Tot_Ord_DiscAmt, SP_Purch.Tax_Amt, SP_Purch.NET_AMT, SP_Purch.Supply_Mode, SP_Purch.Transport, SP_Purch_1.Party_Doc_No AS Expr1, SP_Purch_1.Party_Doc_Date AS Expr2, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, City.CityName, Part.Part_Name, SP_Stock.Part_No, SP_Stock.Qty_Rec, SP_Stock.Rate, SP_Stock.Net_Amt, SP_Purch.V_Date, SP_Purch.V_No, SP_Purch.Party_Doc_No, SP_Purch.Party_Doc_Date, SP_Purch.GR_RR_No, SP_Purch.GR_RR_Date, SP_Purch.Case_No
'FROM (((SP_Purch LEFT JOIN SP_Stock ON SP_Purch.DocID = SP_Stock.DocID) LEFT JOIN (SubGroup LEFT JOIN City ON SubGroup.CityCode = City.CityCode) ON SP_Purch.Party_Code = SubGroup.SubCode) LEFT JOIN Part ON SP_Stock.Part_No = Part.PART_NO) LEFT JOIN SP_Purch AS SP_Purch_1 ON SP_Stock.Invoice_DocId = SP_Purch_1.DocID
End Sub

Private Sub TopCtrl1_eRef()
'    RsDrAc.Requery
    RsParty.Requery
    rsForm31.Requery
    rsForm.Requery
End Sub

Private Sub TopCtrl1_eSave()
Dim i As Integer, SQLPBill$, PurAcCode$
Dim Rst As ADODB.Recordset, mTrans As Boolean
Dim DocIdHlp$, VoucherEditFlag2 As Boolean
Dim LastI As Integer

On Error GoTo errlbl
    Grid_Hide

    If IsValid(txt(VDate), "Bill Date") = False Then Exit Sub
    If IsValid(txt(VType), "Cash/Credit") = False Then Exit Sub
    If IsValid(txt(SerialNo), "Bill Number") = False Then Exit Sub
    If IsValid(txt(LC), "Purchase Type") = False Then Exit Sub
    If IsValid(txt(Party), "Supplier Name") = False Then Exit Sub
    If IsValid(txt(FormType), "Form Type") = False Then Exit Sub
    If txt(SuppChlDate) <> "" Then
        If CDate(txt(SuppChlDate)) > CDate(txt(VDate)) Then
            MsgBox "Supplier Document Date  > Bill Date", vbOKOnly, "Validation": txt(SuppChlDate).SetFocus: Exit Sub
        End If
    End If
    RemoveTxtNull
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        DocID = txt(TxtDocID)
        If GCn.Execute("select count(*) from sp_purch where DocID='" & txt(TxtDocID) & "'").Fields(0) > 0 Then
            If VoucherEditFlag Then
                MsgBox "Purchase Serial No. already exists, Retry", vbCritical, "Validation Error"
                txt(SerialNo).SetFocus
                GoTo errlbl
            Else
                txt(TxtDocID) = GetDocID(GCnFaS, mVtype, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
                If Val(txt(SerialNo)) <= Val(DeCodeDocID(txt(TxtDocID).Tag, Document_No)) Then
                    MsgBox "Purchase Serial No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                    GoTo errlbl
                End If
            End If
        End If
    End If
    DocIdHlp = Replace(txt(TxtDocID), " ", "")
    GCn.BeginTrans
    GCnFaS.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION <> "Add" Then   'Edit Case
        GCn.Execute ("update sp_purch  set Cash_Credit = '" & txt(VType) & "', Party_Code == '" & txt(Party).Tag & "', Party_Name= '" & txt(Party) & _
            "', Party_Doc_No ='" & txt(SuppChlNo) & "',Party_Doc_Date =" & ConvertDate(txt(SuppChlDate)) & ",L_C = '" & left(txt(LC), 1) & "',form_code = '" & txt(FormType).Tag & _
            "', remarks  = '" & txt(Remark) & "', Tot_Goods_Value=" & Val(txt(NetAmt)) & ",NET_AMT = " & Val(txt(NetAmt)) & _
            ",U_Name='" & pubUName & "', U_EntDt=#" & PubServerDate & "#, U_AE='E', DrAc_Code='" & txt(DrAcCode).Tag & _
            "' where DocId = '" & txt(TxtDocID) & "'")
    Else    'Add
        SQLPBill = "insert into sp_purch(DocID,DocIDHelp,V_Type,V_No,Site_Code," _
            & "V_Date,Cash_Credit,Party_Code,Party_Name,Party_Doc_No," _
            & "Party_Doc_Date,L_C,form_code,Tot_Goods_Value," _
            & "NET_AMT,Remarks,U_Name,U_EntDt,U_AE,DrAc_Code) values(" _
            & "'" & txt(TxtDocID) & "','" & DocIdHlp & "','" & mVtype & "'," & Val(txt(SerialNo)) & ",'" & PubSiteCode & PubSiteCode & _
            "'," & ConvertDate(txt(VDate)) & ",'" & txt(VType) & "','" & txt(Party).Tag & "','" & txt(Party) & "','" & txt(SuppChlNo) & _
            "'," & ConvertDate(txt(SuppChlDate)) & ",'" & left(txt(LC), 1) & "','" & txt(FormType).Tag & "', " & Val(txt(NetAmt)) & _
            ", " & Val(txt(NetAmt)) & ",'" & txt(Remark) & "','" & pubUName & "',#" & PubServerDate & "#,'A','" & txt(DrAcCode).Tag & "')"
        'Purchase Bill Add
        GCn.Execute (SQLPBill)
    End If
    
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        'Update Srl No. for Purchase No.
        UpdVouSrlNo GCnFaS, txt(TxtDocID), txt(VDate)
    End If
    'A/c Posting
    '************
    ProcAcPost rsCtrlAc
    'EOF Posting
    GCnFaS.CommitTrans
    GCn.CommitTrans
    mTrans = False
    Set Rst = Nothing
    Master.Requery
    Master.FIND "SearchCode = '" & txt(TxtDocID) & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(txt(SerialNo)) > Val(DeCodeDocID(DocID, Document_No)) Then
            MsgBox "Purchase Serial No." & Trim(DeCodeDocID(DocID, Document_No)) & " already exists ! " & vbCrLf & "New No. " & txt(SerialNo) & " alloted", vbCritical, "Document No. Changed"
        End If
        txt(VDate).Tag = txt(VDate)
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    Exit Sub
errlbl:
    If mTrans Then GCnFaS.RollbackTrans: GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "SELECT sp_purch.DocId as searchcode, sp_purch.V_Date AS VoucherDate,sp_purch.DocId, sp_purch.v_Type, sp_purch.v_No, sp_purch.Site_Code, SubGroup.Name as PartyName FROM sp_purch LEFT JOIN SubGroup ON sp_purch.Party_Code = SubGroup.SubCode where v_type in ('" & PurCashVType & "','" & PurCrVType & "') Order By V_Date Desc"
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
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub Txt_GotFocus(Index As Integer)
On Error GoTo ELoop
If txt(VType).TEXT = "" And Index <> VDate Then txt(VType).SetFocus
Ctrl_GetFocus txt(Index)
Grid_Hide
Select Case Index
    Case VType
        ListArray = Array("Cash", "Credit")
        Set mListItem = ListView_Items(ListView, txt, VType, ListArray, 2)
    Case Party
        Set DGParty.DataSource = RsParty
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "name ='" & txt(Index).TEXT & "'"
        End If
    Case LC
        ListArray = Array("Local", "Central")
        Set mListItem = ListView_Items(ListView, txt, LC, ListArray, 2)
'    Case PermitType
'        Set DGForm.DataSource = rsForm31
'        If rsForm31.RecordCount = 0 Or (rsForm31.EOF = True Or rsForm31.BOF = True) Or txt(Index).Text = "" Then Exit Sub
'        If txt(Index).Text <> rsForm31!Name Then
'            rsForm31.MoveFirst
'            rsForm31.FIND "name ='" & txt(Index).Text & "'"
'        End If
    Case FormType
        Set DGForm.DataSource = rsForm
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> rsForm!Name Then
            rsForm.MoveFirst
            rsForm.FIND "name ='" & txt(Index).TEXT & "'"
        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
On Error GoTo ELoop
Select Case Index
    Case VType
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case LC
        ListView_KeyDown FrmList, ListView, txt, Index, KeyCode, Shift, txt(Index).left, (txt(Index).top + txt(Index).height), txt(Index).width, 600
    Case Party
        If txt(VType).TEXT = "Credit" Then
            DGridTxtKeyDown DGParty, txt, Party, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
        End If
    Case FormType
        DGridTxtKeyDown DGForm, txt, FormType, rsForm, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
'    Case PermitType
'        DGridTxtKeyDown DGForm, txt, PermitType, rsForm31, KeyCode, False, 1, frmTaxForms, "frmTaxForms"
End Select
If FrmList.Visible = False And DGParty.Visible = False And DGForm.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VType Then Txt_Validate Index, True
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> NetAmt Then Ctrl_DownKeyDown KeyCode, Shift
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = NetAmt Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    If TopCtrl1.TopText2.CAPTION = "Add" And Index <> VDate Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> Party Then
        If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
    End If
End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(keyascii)
Select Case Index
'    Case PermitType
'        If DGForm.Visible = True Then DGridTxtKeyPress txt, Index, rsForm31, KeyAscii, "Name"
    Case Party
        If txt(VType).TEXT = "Credit" Then
            If DGParty.Visible = True Then DGridTxtKeyPress txt, Index, RsParty, keyascii, "Name"
        End If
    Case FormType
        If DGForm.Visible = True Then DGridTxtKeyPress txt, Index, rsForm, keyascii, "Name"
    Case SerialNo
        Call NumPress(txt(Index), keyascii, 6, 0)
    Case NetAmt
        Call NumPress(txt(Index), keyascii, 8, 2)
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case VType
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
    Case LC
        ListView_KeyUp ListView, txt, Index, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim i As Integer
Select Case Index
    Case VType
        If IsValid(txt(VType), "Cash Credit") = False Then Cancel = True:   Exit Sub
        If txt(VType).TEXT <> "" Then txt(VType).TEXT = ListView.SelectedItem.TEXT
        If txt(VType).TEXT = "Cash" Then
            txt(Party).TEXT = "Cash"
            txt(Party).Tag = PubSprCashAc
            mVtype = PurCashVType
        Else
            txt(Party).TEXT = ""
            txt(Party).Tag = ""
            mVtype = PurCrVType
        End If
        txt(VType).Tag = txt(VType).TEXT
        txt(TxtDocID) = GetDocID(GCnFaS, mVtype, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
        DocID = txt(TxtDocID)
    Case LC
        If txt(LC).TEXT <> "" Then txt(LC).TEXT = ListView.SelectedItem.TEXT
        If IsValid(txt(LC), "Purchase Type") = False Then Cancel = True:   Exit Sub
    Case Party
        If IsValid(txt(Index), Label3(3)) = False Then Cancel = True: Exit Sub
        'by lps 25-06-02
        If txt(VType).TEXT = "Cash" Then
            mPartyType = 0
            txt(Index).Tag = PubSprCashAc
            GSQL = "Select OrderID as Code,Order_Reg_No as Name,Order_Reg_Dt, " & cTrim(cMID("OrderID", "8", "5")) & "+CStr(Trim(Right(OrderID,8))) as OurDocNo,V_Date From SP_Order Where left(OrderID,1)='" & PubDivCode & "' and left(Order_Type,4)='S_PO' and V_Date<=#" & Format(txt(VDate), "dd-mmm-yyyy") & "# and OrdClosDate is null Order By OrderID"
        ElseIf txt(VType).TEXT = "Credit" Then
            mPartyType = RsParty!Party_Type
            txt(Index).TEXT = RsParty!Name
            txt(Index).Tag = RsParty!Code
            GSQL = "Select OrderID as Code,Order_Reg_No as Name,Order_Reg_Dt, Trim(" & cMID("OrderID", "8", "5") & ")+CStr(Trim(Right(OrderID,8))) as OurDocNo,V_Date From SP_Order Where left(OrderID,1)='" & PubDivCode & "' and left(Order_Type,4)='S_PO' and Party_Code='" & txt(Party).Tag & "' and V_Date<=#" & Format(txt(VDate), "dd-mmm-yyyy") & "# and OrdClosDate is null Order By OrderID"
        End If
'    Case PermitType
'        If rsForm31.RecordCount = 0 Or (rsForm31.EOF = True Or rsForm31.BOF = True) Or txt(Index).Text = "" Then
'            txt(Index).Text = ""
'            txt(Index).Tag = ""
'        Else
'            txt(Index).Text = rsForm31!Name
'            txt(Index).Tag = rsForm31!Code
'        End If
    Case FormType
        If rsForm.RecordCount = 0 Or (rsForm.EOF = True Or rsForm.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            If IsNull(rsForm!PurSal_Ac_Code) Then
                MsgBox "Please Define Purchase / Sale A/c in selected form", vbCritical, "Validation"
                Cancel = True
                Exit Sub
            End If
            txt(Index).TEXT = rsForm!Name
            txt(Index).Tag = rsForm!Code
            txt(DrAcCode).Tag = IIf(IsNull(rsForm!PurSal_Ac_Code), "", rsForm!PurSal_Ac_Code)
            txt(DrAcCode) = GCn.Execute("Select Name from subgroup where SubCode='" & txt(DrAcCode).Tag & "'").Fields(0).Value
        End If
    Case SuppChlDate
        txt(Index).TEXT = RetDate(txt(Index))
    Case VDate
        If Len(Trim(txt(VDate).TEXT)) = 0 Then
            txt(VDate).TEXT = PubLoginDate
        Else
            txt(Index).TEXT = RetDate(txt(Index))
        End If
        Cancel = Not CheckFinYear(txt(Index))
        If Cancel = False Then
            If txt(VType).TEXT = "" Then txt(VType).SetFocus: Exit Sub
            txt(TxtDocID) = GetDocID(GCnFaS, mVtype, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            DocID = txt(TxtDocID)
        End If
    Case SerialNo
        If IsValid(txt(SerialNo), "Serial No.") = False Then Cancel = True:   Exit Sub
        If VoucherEditFlag Then      ' Manual
            txt(TxtDocID) = GetDocID(GCnFaS, mVtype, txt(VDate), VoucherEditFlag, txt(SerialNo), LblVPrefix)
            DocID = txt(TxtDocID)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select DocID From sp_purch Where docid='" & txt(TxtDocID) & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Serial No. Not Allowed", vbInformation, "Validation"
                Cancel = True
                txt(SerialNo).SetFocus
            End If
        End If
End Select
Set Rst = Nothing
End Sub
Private Sub DGForm_Click()
    If rsForm.RecordCount > 0 Then
        txt(FormType).TEXT = rsForm!Name
        txt(FormType).Tag = rsForm!Code
    End If
    txt(FormType).SetFocus
    DGForm.Visible = False
End Sub

Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        txt(Party).TEXT = RsParty!Name
        txt(Party).Tag = RsParty!Code
        mPartyType = RsParty!Party_Type
    End If
    txt(Party).SetFocus
    DGParty.Visible = False
End Sub

'******* Fuctions **********
Private Sub BlankText()
Dim i As Byte
For i = 0 To txt.Count - 1
    txt(i).TEXT = ""
Next i
DocID = ""
End Sub

Private Sub MoveRec()
Dim Master1 As ADODB.Recordset, i As Integer
On Error GoTo error1
If Master.RecordCount > 0 Then
    If InStr(Me.TopCtrl1.Tag, "E") <> 0 Then Me.TopCtrl1.tEdit = True
    Set Master1 = New Recordset
    Master1.CursorLocation = adUseClient
    Master1.Open "select SubGroup.Name,SubGroup.Party_Type,SP_Purch.* from SP_Purch " _
        & " left join SubGroup on SP_Purch.Party_Code=SubGroup.SubCode " _
        & " where DocID='" & Master!SearchCode & "'", GCn, adOpenStatic, adLockReadOnly
    If Master1!CancelYN = 1 Then
        TopCtrl1.tEdit = False
        LblCancel.Visible = True
    Else
        LblCancel.Visible = False
    End If
    DocID = Master!SearchCode
    txt(TxtDocID) = Master1!DocID
    LblDiv.CAPTION = "Division : " & left(Master1!DocID, 1)
    LblSite.CAPTION = "Site Code : " & Master1!Site_Code
    LblVPrefix.CAPTION = mID(Master1!DocID, 9, 5)
    txt(SerialNo) = Master1!V_NO
    txt(VDate) = Master1!V_DATE
    txt(VType) = Master1!Cash_Credit
    txt(Party).Tag = Master1!Party_code
    If txt(VType) = "Cash" Then
        mVtype = PurCashVType
        mPartyType = 0
        txt(Party) = Master1!Party_Name
    ElseIf txt(VType) = "Credit" Then
        mVtype = PurCrVType
        mPartyType = Master1!Party_Type
        txt(Party) = Master1!Name
    End If
    txt(DrAcCode).Tag = IIf(IsNull(Master1!DrAc_Code), "", Master1!DrAc_Code)
    txt(DrAcCode) = GCn.Execute("Select Name from subgroup where SubCode='" & Master1!DrAc_Code & "'").Fields(0).Value

    txt(SuppChlNo) = IIf(IsNull(Master1!Party_Doc_No), "", Master1!Party_Doc_No)
    txt(SuppChlDate) = IIf(IsNull(Master1!Party_Doc_Date), "", Master1!Party_Doc_Date)
    txt(LC) = IIf(Master1!L_C = "L", "Local", "Central")
    txt(FormType).Tag = IIf(IsNull(Master1!Form_Code), "", Master1!Form_Code)
    If txt(FormType).Tag <> "" Then
        txt(FormType) = GCn.Execute("select form_desc from taxforms where form_code = '" & txt(FormType).Tag & "'").Fields(0).Value
    Else
        txt(FormType) = ""
    End If
'    txt(PermitType).Tag = IIf(IsNull(Master1!RoadPermit_FormCode), "", Master1!RoadPermit_FormCode)
'    If txt(PermitType).Tag <> "" Then
'        txt(PermitType) = GCn.Execute("select form_desc from taxforms where form_code = '" & txt(PermitType).Tag & "'").Fields(0).Value
'    Else
'        txt(PermitType) = ""
'    End If
'    txt(FormNo) = IIf(IsNull(Master1!RoadPermit_No), "", Master1!RoadPermit_No)
    txt(Remark) = IIf(IsNull(Master1!Remarks), "", Master1!Remarks)
    
    'txt(TotAmt) = Format(IIf(IsNull(Master1!Tot_Amt), 0, Master1!Tot_Amt), "0.00")
    'txt(TotGoods) = Format(IIf(IsNull(Master1!Tot_Goods_Value), 0, Master1!Tot_Goods_Value), "0.00")
    txt(NetAmt) = Format(IIf(IsNull(Master1!Net_Amt), 0, Master1!Net_Amt), "0.00")
    'txt(TotPurAmt) = Format(Val(txt(NetAmt)) + Val(txt(EntryTaxAmt)), "0.00")
Else
    Call BlankText
End If
Set Master1 = Nothing
Grid_Hide
Me.TopCtrl1.tPrn = False
Exit Sub

error1:
Set Master1 = Nothing
CheckError
End Sub

Private Sub Ini_Grid()
Dim i As Byte
  ' |Part No.1|Part Name2|Unit 3|PO No 4|Taxable 5|MRP6|Qty(Doc)7|Qty(Phy)8|NDP 9 |Amount 10
'  |Dis %11|Ord Dis %12|Amount 13|Loal Name 14|Curr Stk Qty 15|MRP Qty 16 |Taxable Qty 17|TaxPaid Qty 18|Taxable Rate 19|TaxPaid Rate 20|Bin Location 21|Last Purch Rate 22|High Purch Rate 23|Low Purch Rate 24
    DGDrAc.width = 5130:   DGDrAc.left = Me.width - (DGDrAc.width + mRtScale): DGDrAc.top = mTopScale '390
    DGDrAc.height = 4935

    DGParty.width = 5130:   DGParty.left = Me.width - (DGParty.width + mRtScale): DGParty.top = mTopScale '390
    DGParty.height = 4935
    DGForm.width = DGParty.width: DGForm.left = DGParty.left: DGForm.top = DGParty.top: DGForm.height = DGParty.height
End Sub
Private Sub Disp_Text(Enb As Boolean)
Dim i As Integer
For i = 0 To txt.Count - 1
    txt(i).Enabled = Enb
    txt(i).ForeColor = CtrlFColOrg
Next
txt(TxtDocID).Enabled = False
If TopCtrl1.TopText2 = "Edit" Then
    txt(VDate).Enabled = False
    txt(SerialNo).Enabled = False
    txt(VType).Enabled = False
End If
FldEnabled False

txtDisabled_Color Me

End Sub
Private Sub Grid_Hide()
    If DGForm.Visible = True Then DGForm.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Private Sub FldEnabled(Enb As Boolean)
'    txt(PermitType).Enabled = Enb
'    txt(FormNo).Enabled = Enb
    
    txtDisabled_Color Me

End Sub

Private Sub RemoveTxtNull()
Dim i As Integer
For i = 0 To txt.Count - 1
    txt(i).TEXT = IIf(IsNull(txt(i).TEXT), "", txt(i).TEXT)
Next i
End Sub


Private Function ProcAcPost(rsCtrlAc As ADODB.Recordset) As Boolean
On Error GoTo lblExit
Dim xNetAmt As Double, xEntryTaxAmt As Double
'A/c Posting related declarations
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, mNarr$, TaxSQL$, i As Integer, J As Integer
Dim mSprPurPfx$, mFADocID$, SupDocNo$
If txt(SuppChlNo) <> "" Then
    SupDocNo = "Supplier Document No." & txt(SuppChlNo)
End If
If txt(SuppChlDate) <> "" Then
    SupDocNo = SupDocNo & " Date " & txt(SuppChlDate)
End If

    If txt(VType) = "Cash" Then
        mNarr = "Through Other Purchase (Cash) " & SupDocNo & " " & txt(Remark)
    Else
        mNarr = "Through Other Purchase (Cr) " & SupDocNo & " " & txt(Remark)
    End If
    mCommNarr = mNarr & " [Common]"
    mFADocID = txt(TxtDocID)
    GSQL = "select TF.PurSal_Ac_Code,sum(NET_AMT) as NetAmt " & _
        "from SP_Purch " & _
        "left join TaxFormsAc as TF on SP_Purch.Form_Code&'" & PubDivCode & "'=TF.Form_Code&TF.Div_Code " & _
        "where docid='" & txt(TxtDocID) & _
        "' group by TF.PurSal_Ac_Code"
    
    Set GRs = New ADODB.Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open GSQL, GCn, adOpenStatic, adLockReadOnly
    
    '*** pURCHASE Amount Row
'   0.Purchase A/c
'   1.Party A/c or Cash A/c
'    Dim LedgAry() As LedgRec
    
'*********
    i = -1
    Do While GRs.EOF = False
        If IsNull(GRs!PurSal_Ac_Code) Or GRs!PurSal_Ac_Code = "" Then
            MsgBox "Please Define Purchase A/c in Tax Forms " & GRs!PurSal_Ac_Code & vbCrLf & "A/c Psoting Aborted", vbCritical, "A/c Posting"
            GoTo lblExit
        End If
        If i = -1 Then
            ReDim Preserve LedgAry(1)
            i = 0
        Else
            i = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(i)
        End If
        LedgAry(i).SubCode = GRs!PurSal_Ac_Code 'Purchase A/c
        LedgAry(i).AmtDr = IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
        LedgAry(i).Narration = mNarr
        LedgAry(i).ContraSub = IIf(txt(VType) = "Cash", PubSprCashAc, txt(Party).Tag)
        
        xNetAmt = xNetAmt + IIf(IsNull(GRs!NetAmt), 0, GRs!NetAmt)
        GRs.MoveNext
    Loop
    If xEntryTaxAmt <> 0 Then
        i = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(i)
        LedgAry(i).SubCode = rsCtrlAc!EntryTax_Ac
        LedgAry(i).AmtCr = xEntryTaxAmt
        LedgAry(i).Narration = mNarr
    End If
    i = UBound(LedgAry) + 1
    ReDim Preserve LedgAry(i)
    LedgAry(i).SubCode = IIf(txt(VType) = "Cash", PubSprCashAc, txt(Party).Tag)
    LedgAry(i).AmtCr = (xNetAmt - xEntryTaxAmt)
    LedgAry(i).Narration = mNarr
    
    mResult = LedgerPost(left(TopCtrl1.TopText2, 1), LedgAry, GCnFaS, mFADocID, CDate(txt(VDate)), mCommNarr)
    If mResult <> 1 Then
        MsgBox "Error in Ledger Posting", vbOKOnly, "Validation"
        ProcAcPost = True
    End If
lblExit:
    Set GRs = Nothing
    If err.NUMBER <> 0 Then
        MsgBox err.Description & vbCrLf & "Ledger Posting Terminated!", vbCritical
        ProcAcPost = True
    End If
End Function

