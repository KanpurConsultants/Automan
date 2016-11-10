VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmService 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Service Master"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   14
      Left            =   2685
      MaxLength       =   20
      TabIndex        =   9
      Top             =   2610
      Width           =   585
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
      Index           =   13
      Left            =   8355
      MaxLength       =   6
      TabIndex        =   11
      Top             =   1485
      Width           =   690
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
      Index           =   12
      Left            =   8355
      MaxLength       =   6
      TabIndex        =   15
      Top             =   2025
      Width           =   690
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
      Index           =   11
      Left            =   8355
      MaxLength       =   6
      TabIndex        =   13
      Top             =   1755
      Width           =   690
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
      Index           =   10
      Left            =   7440
      MaxLength       =   6
      TabIndex        =   10
      Top             =   1485
      Width           =   660
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
      Index           =   9
      Left            =   7440
      MaxLength       =   6
      TabIndex        =   14
      Top             =   2025
      Width           =   660
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
      Index           =   8
      Left            =   7440
      MaxLength       =   6
      TabIndex        =   12
      Top             =   1755
      Width           =   660
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   7
      Left            =   2685
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1650
      Width           =   585
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   2685
      MaxLength       =   20
      TabIndex        =   2
      Top             =   930
      Width           =   2490
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   6
      Left            =   2685
      MaxLength       =   12
      TabIndex        =   7
      Top             =   2130
      Width           =   2490
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   5
      Left            =   2685
      MaxLength       =   18
      TabIndex        =   3
      Top             =   1170
      Width           =   2490
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   7065
      TabIndex        =   22
      Top             =   4380
      Visible         =   0   'False
      Width           =   2505
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   0
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   0
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
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   4
      Left            =   2685
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2370
      Width           =   585
   End
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   2685
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1890
      Width           =   585
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   2685
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1410
      Width           =   585
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2685
      MaxLength       =   2
      TabIndex        =   1
      Top             =   690
      Width           =   585
   End
   Begin MSDataGridLib.DataGrid DGServ 
      Height          =   3225
      Left            =   1290
      TabIndex        =   41
      Top             =   3540
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   5689
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
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
      Caption         =   "List Of Service"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Serv_Type"
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
         DataField       =   "Serv_Desc"
         Caption         =   "Description"
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
         DataField       =   "Serv_Type"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   3089.764
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Editable (Y/N)"
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
      Left            =   900
      TabIndex        =   40
      Top             =   2610
      Width           =   1635
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "[100%]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   7
      Left            =   8355
      TabIndex        =   39
      Top             =   1230
      Width           =   660
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "[100%]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   6
      Left            =   7455
      TabIndex        =   38
      Top             =   1230
      Width           =   660
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001860A7&
      Height          =   225
      Index           =   5
      Left            =   8145
      TabIndex        =   36
      Top             =   1500
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001860A7&
      Height          =   225
      Index           =   4
      Left            =   7245
      TabIndex        =   35
      Top             =   1500
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001860A7&
      Height          =   225
      Index           =   3
      Left            =   8145
      TabIndex        =   34
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001860A7&
      Height          =   225
      Index           =   1
      Left            =   7245
      TabIndex        =   33
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001860A7&
      Height          =   225
      Index           =   0
      Left            =   8145
      TabIndex        =   32
      Top             =   1770
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001860A7&
      Height          =   225
      Index           =   2
      Left            =   7245
      TabIndex        =   31
      Top             =   1770
      Width           =   180
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   195
      Index           =   5
      Left            =   8355
      TabIndex        =   30
      Top             =   1005
      Width           =   585
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Spare"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   210
      Index           =   3
      Left            =   7470
      TabIndex        =   29
      Top             =   997
      Width           =   525
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Share"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   5670
      TabIndex        =   28
      Top             =   1485
      Width           =   1410
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Share"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   5670
      TabIndex        =   27
      Top             =   2025
      Width           =   1140
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telco Share"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   5670
      TabIndex        =   26
      Top             =   1755
      Width           =   1020
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Serial No."
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
      Left            =   915
      TabIndex        =   25
      Top             =   1650
      Width           =   1545
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Name"
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
      Left            =   915
      TabIndex        =   24
      Top             =   930
      Width           =   1200
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Annual Target"
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
      Left            =   915
      TabIndex        =   21
      Top             =   2370
      Width           =   1200
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chargeable From"
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
      Left            =   915
      TabIndex        =   20
      Top             =   2130
      Width           =   1485
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
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
      Left            =   915
      TabIndex        =   19
      Top             =   1890
      Width           =   435
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Free Service Index"
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
      Left            =   915
      TabIndex        =   18
      Top             =   1410
      Width           =   1635
   End
   Begin VB.Label LblName 
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
      Index           =   1
      Left            =   915
      TabIndex        =   17
      Top             =   1185
      Width           =   1125
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Code"
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
      Left            =   915
      TabIndex        =   16
      Top             =   690
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackColor       =   &H00BAD3C9&
      BorderStyle     =   1  'Fixed Single
      Height          =   1890
      Left            =   5550
      TabIndex        =   37
      Top             =   825
      Width           =   3645
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MasterFormExit As Boolean
Dim ADDFLAG As Byte
Dim RstMain As ADODB.Recordset, RstHelp As ADODB.Recordset
Dim mFlag As Byte
Private Const S_T_Code = 0, S_T_Desc = 1, F_S_Code = 2, DAYS = 3, ServType = 5
Private Const ChrgFrom = 6, An_Target = 4, Serv_Cat = 0, ServSrlNo = 7
Private Const SprTel = 8, SprDlr = 9, SprCust = 10
Private Const LabTel = 11, LabDlr = 12, LabCust = 13
Private Const RateEditableYN As Byte = 14


Dim ListArray As Variant
Dim mListItem As ListItem

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_Deactivate()
If MasterFormExit = True Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
FormKeyDown Me, KeyCode, Shift, MasterFormExit
Exit Sub
ELoop:
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
Me.top = 0: Me.left = 0
PubUParam = PubUParam  'UserPermission(Me.Name)
TopCtrl1.Tag = PubUParam
Set RstMain = New ADODB.Recordset
'RstMain.Open "Select Serv_Type as searchcode,Service_Type.* From Service_Type  where SITE_CODE=" & Chk_Text(PubSiteCode) & " Order by Serv_Type", GCn, adOpenDynamic, adLockOptimistic
If PubMoveRecYn Then
    RstMain.Open "Select Serv_Type as SearchCode,Service_Type.* From Service_Type Order by Serv_Type", GCn, adOpenDynamic, adLockOptimistic
Else
    Set RstMain = GCn.Execute("Select Top 1 Serv_Type as SearchCode,Service_Type.* From Service_Type Order by Serv_Type")
End If

Set RstHelp = New ADODB.Recordset
'RstHelp.Open "Select Serv_Type,Serv_Desc FROM Service_Type where SITE_CODE=" & Chk_Text(PubSiteCode) & "Order by Serv_Type", GCn, adOpenDynamic, adLockOptimistic
RstHelp.Open "Select Serv_Type,Serv_Desc FROM Service_Type Order by Serv_Type", GCn, adOpenDynamic, adLockOptimistic
CtrlClckCol
'If MasterFormExit Then TopCtrl1_eAdd: Exit Sub
Disp_Text SETS("INI", Me, RstMain)
MoveRec


ADDFLAG = 0:    mFlag = 0
Set DGServ.DataSource = RstHelp

End Sub

Private Sub Form_Unload(Cancel As Integer)
MasterFormExit = False
    Set RstMain = Nothing: Set RstHelp = Nothing
End Sub

Private Sub ListView_Click()
    Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    Txt(Val(ListView.Tag)).SetFocus
End Sub

Public Sub TopCtrl1_eAdd()
On Error GoTo Errloop
BlankText
Disp_Text SETS("ADD", Me, RstMain)
Txt(S_T_Code).Tag = Txt(S_T_Code)
Txt(SprCust) = "100.00"
Txt(LabCust) = "100.00"
Txt_GotFocus S_T_Code
ADDFLAG = 1
Txt(F_S_Code).Enabled = False
Txt(S_T_Code).SetFocus
Exit Sub
Errloop:    MsgBox err.Description, vbCritical
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
Dim I As Byte
If RstMain.RecordCount > 0 Then
    Disp_Text SETS("EDIT", Me, RstMain)
    Txt(S_T_Code).Enabled = False
    Txt(S_T_Desc).Tag = Txt(S_T_Desc)
    Txt_GotFocus S_T_Desc
    ADDFLAG = 2
    If Txt(ServType) = "AMC" Then
        For I = 8 To 13
            Txt(I).Enabled = True
        Next
    Else
        For I = 8 To 13
            Txt(I).Enabled = False
        Next
    End If
    If Txt(ServType) = "Free Service" Then
        Txt(F_S_Code).Enabled = True
    Else
        Txt(F_S_Code).Enabled = False
    End If
    Txt(S_T_Desc).SetFocus
Else
    MsgBox "There Is No Record To Edit.", vbInformation, "Information"
End If
Exit Sub
Errloop:    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
On Error GoTo Errloop
Dim transFalg As Byte
transFalg = 0
Dim XBM
Dim Res As Integer
    If RstMain.RecordCount > 0 Then
        Res = MsgBox("Do You Want to Delete Record ", 4 + vbQuestion, "Confirmation ")
        If Res = 6 Then
            GCn.BeginTrans
            XBM = RstMain.Bookmark
            transFalg = 1
            GCn.Execute ("delete * from Service_Type where Serv_Type= '" & Txt(S_T_Code)) & "'"
            GCn.CommitTrans
            transFalg = 0
            RstMain.Requery
            RstHelp.Requery
            If RstMain.RecordCount >= XBM Then
                RstMain.Bookmark = XBM
            Else
                If RstMain.EOF = False Then RstMain.MoveLast
            End If
            Call MoveRec
        End If
    Else
        MsgBox "No Records To Delete.", vbInformation, "Information"
    End If

Exit Sub
Errloop:    If transFalg = 1 Then GCn.RollbackTrans
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
'switch(Serv_catg='P','PDI,Serv_Catg='F','Free Service',Serv_Catg='C','Chargeable',Serv_Catg='A','Accidental',Serv_Catg='D','D & P',Serv_Catg='R','Repeat Job') as ServCat
    If PubBackEnd = "A" Then
        GSQL = "Select Serv_Type as SearchCode,Serv_Type,Serv_Desc," & _
        "switch(Serv_catg='P','PDI',Serv_Catg='F','Free Service',Serv_Catg='C','Chargeable',Serv_Catg='A','Accidental',Serv_Catg='D','D & P',Serv_Catg='R','Repeat Job') as ServCat," & _
        "Chrg_From From Service_Type Order by Serv_Type"
    Else
        GSQL = "Select Serv_Type as SearchCode,Serv_Type,Serv_Desc," & _
        "Case Serv_catg When 'P' Then 'PDI' When 'F' Then 'Free Service' When 'C' Then 'Chargeable' When 'A' Then 'Accidental' When 'D' Then 'D & P' When 'R' Then 'Repeat Job' End  as ServCat," & _
        "Chrg_From From Service_Type Order by Serv_Type"
    End If
    Set SearchForm = Me
    FIND.Show vbModal
    Exit Sub    'SubServ_Type
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RstMain.MoveFirst
        RstMain.FIND ("SEARCHCODE='" & MyValue & "'")
    Else
        Set RstMain = GCn.Execute("Select Serv_Type as SearchCode,Service_Type.* From Service_Type Where Serv_Type ='" & MyValue & "' Order by Serv_Type")
    End If
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
Dim I As Integer, mQRY$, mRepName$
Dim Rst As ADODB.Recordset
On Error GoTo ERRORHANDLER

    mRepName = "ServiceType"
    mQRY = "SELECT Service_Type.*,switch(Serv_catg='P','PDI',Serv_Catg='F','Free Service',Serv_Catg='C','Chargeable',Serv_Catg='A','Accidental',Serv_Catg='D','D & P',Serv_Catg='R','Repeat Job') as ServCat from Service_Type"

    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
     rpt.Database.SetDataSource Rst
     rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION, , True)
    Set Rst = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub TopCtrl1_eSave()
Dim transFlag As Byte
On Error GoTo Errloop
Dim mServType As String
    transFlag = 0
    If IsValid(Txt(S_T_Code), "Service Type Code") = False Then Txt_GotFocus S_T_Code: Exit Sub
    If IsValid(Txt(S_T_Desc), "Service Description") = False Then Txt_GotFocus S_T_Desc: Exit Sub
    If IsValid(Txt(ServType), "Service Type") = False Then Txt_GotFocus S_T_Desc: Exit Sub
    If Txt(ServType) = "Free Service" Then
        If IsValid(Txt(F_S_Code), "Free Service Code") = False Then Txt_GotFocus S_T_Desc: Exit Sub
    End If
    If IsValid(Txt(ChrgFrom), "Chargeable From") = False Then Txt_GotFocus S_T_Desc: Exit Sub
    If IsValid(Txt(ServSrlNo), "Service Serial No.") = False Then Txt_GotFocus ServSrlNo: Exit Sub
    'SprTel, SprDlr, SprCust, LabTel, LabDlr, LabCust
'    If Val(Txt(SprTel)) + Val(Txt(SprDlr)) + Val(Txt(SprCust)) <> 100 Then MsgBox "Fill 100% Cumulative Share for Spare", vbInformation, "Validation Check": Txt(SprCust).SetFocus: Exit Sub
'    If Val(Txt(LabTel)) + Val(Txt(LabDlr)) + Val(Txt(LabCust)) <> 100 Then MsgBox "Fill 100% Cumulative Share for Labour", vbInformation, "Validation Check": Txt(LabCust).SetFocus: Exit Sub
    If ADDFLAG = 1 Then If GCn.Execute("Select COUNT(*) From Service_Type Where Serv_Type= " & Chk_Text(PubSiteCode + Trim(Txt(S_T_Code))) & " AND SITE_CODE='" & PubSiteCode & "'").Fields(0) > 0 Then MsgBox "Service Type Code Already Exists", vbInformation, "Duplicate Checking": Txt_GotFocus S_T_Code: Txt(S_T_Code).SetFocus: Exit Sub
    GCn.BeginTrans
    transFlag = 1
    
    Select Case Txt(ServType)
        Case ""
            mServType = ""
        Case "PDI"
            mServType = "P"
        Case "Free Service"
            mServType = "F"
        Case "Chargeable"
            mServType = "C"
        Case "Accidental"
            mServType = "A"
        Case "Denting & Painting"
            mServType = "D"
        Case "Repeat Job"
           mServType = "R"
        Case "AMC"
            mServType = "M"
    End Select

    If ADDFLAG = 1 Then
        GCn.Execute ("DELETE From Service_Type Where Serv_Type= " & Chk_Text(Trim(Txt(S_T_Code))) & " AND SITE_CODE='" & PubSiteCode & "'")
        GCn.Execute ("Insert Into Service_Type(Serv_Type,Site_Code,SERV_DESC,Serv_Catg,FreeServCode,Days,Chrg_From,Serv_Target,Serv_SrlNo,SprTel, SprDlr, SprCust, LabTel, LabDlr, LabCust, RateEditableYn, U_Name, U_EntDt,U_AE) Values(" & Chk_Text(Trim(Txt(S_T_Code))) & ",'" & PubSiteCode & "'," & Chk_Text(Trim(Txt(S_T_Desc))) & ",'" & mServType & "','" & IIf(Txt(F_S_Code) <> "", Txt(F_S_Code), " ") & "'," & Val(Txt(DAYS)) & "," & IIf(Txt(ChrgFrom) = "Customer", 1, 0) & "," & Val(Txt(An_Target)) & "," & Val(Txt(ServSrlNo)) & "," & Val(Txt(SprTel)) & ", " & Val(Txt(SprDlr)) & ", " & Val(Txt(SprCust)) & ", " & Val(Txt(LabTel)) & "," & Val(Txt(LabDlr)) & " , " & Val(Txt(LabCust)) & ", " & IIf(Txt(RateEditableYN) = "Yes", 1, 0) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & IIf(ADDFLAG = 1, "A", "E") & "')")
    ElseIf ADDFLAG = 2 Then
        GCn.Execute ("UPDATE Service_Type SET Site_Code='" & PubSiteCode & "',SERV_DESC='" & Trim(Txt(S_T_Desc)) & "',SERV_CATG='" & mServType & "',FREESERVCODE='" & IIf(Txt(F_S_Code) <> "", Txt(F_S_Code), " ") & "',DAYS=" & IIf(Len(Txt(DAYS)) = 0, 0, Txt(DAYS)) & ",CHRG_FROM=" & IIf(Txt(ChrgFrom) = "Customer", 1, 0) & ",SERV_TARGET=" & IIf(Len(Txt(An_Target)) = 0, 0, Txt(An_Target)) & ",Serv_SrlNo=" & Val(Txt(ServSrlNo)) & ",SprTel = " & Val(Txt(SprTel)) & ", SprDlr = " & Val(Txt(SprDlr)) & ", SprCust = " & Val(Txt(SprCust)) & ",LabTel =  " & Val(Txt(LabTel)) & ",LabDlr = " & Val(Txt(LabDlr)) & " , LabCust = " & Val(Txt(LabCust)) & ", RateEditableYn=" & IIf(Txt(RateEditableYN) = "Yes", 1, 0) & ",U_Name='" & pubUName & "',U_EntDt=" & ConvertDate(PubServerDate) & ",U_AE='" & IIf(ADDFLAG = 1, "A", "E") & "' WHERE Serv_Type= '" & Trim(Txt(S_T_Code)) & "'" & "")
    End If
    GCn.CommitTrans
    If MasterFormExit Then Unload Me: Exit Sub
    transFlag = 0
    If PubMoveRecYn Then
        RstMain.Requery
    Else
        Set RstMain = GCn.Execute("Select Serv_Type as SearchCode,Service_Type.* From Service_Type Where Serv_Type ='" & Trim(Txt(S_T_Code)) & "' Order by Serv_Type")
    End If
    RstHelp.Requery
    RstMain.FIND ("Serv_Type='" & Trim(Txt(S_T_Code))) & "'"
    If ADDFLAG = 1 Then
        BlankText
        Txt_GotFocus S_T_Code
        Txt(S_T_Code).SetFocus
    Else
        Disp_Text SETS("INI", Me, RstMain)
        MoveRec
        CtrlClckCol
        ADDFLAG = 0
        DGServ.Visible = False
    End If
Exit Sub
Errloop:    If transFlag = 1 Then GCn.RollbackTrans
            MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        If MasterFormExit Then Unload Me: Exit Sub
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, RstMain)
        Me.ActiveControl.SetFocus
        MoveRec
        CtrlClckCol
        DGServ.Visible = False
    End If
Exit Sub
Errloop:
    MsgBox err.Description, vbCritical
End Sub

'**********Functions***********
Private Sub CtrlClckCol()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).BackColor = CtrlBColOrg:      Txt(I).ForeColor = CtrlFColOrg
Next I
'For i = 0 To Ser_Typ.Count - 1
'    Ser_Typ(i).BackColor = CtrlBColOrg:      Ser_Typ(i).ForeColor = CtrlFColOrg
'Next i
End Sub

Private Sub MoveRec()
On Error GoTo Errloop
Dim I As Byte
RST_BOF_EOF RstMain
TopCtrl1.tDel = False
If RstMain.RecordCount <= 0 Then
    BlankText
Else
    Txt(S_T_Code) = XNull(RstMain!Serv_Type)
    Txt(S_T_Desc) = XNull(RstMain!Serv_Desc)
    Select Case XNull(RstMain!serv_catg)
        Case ""
            Txt(ServType) = ""
        Case "P"
            Txt(ServType) = "PDI"
        Case "F"
            Txt(ServType) = "Free Service"
        Case "C"
            Txt(ServType) = "Chargeable"
        Case "A"
            Txt(ServType) = "Accidental"
        Case "D"
            Txt(ServType) = "Denting & Painting"
        Case "R"
            Txt(ServType) = "Repeat Job"
        Case "M"
            Txt(ServType) = "AMC"
    End Select
    Txt(F_S_Code) = XNull(RstMain!FREESERVCODE)
    Txt(ServSrlNo) = XNull(RstMain!Serv_SrlNo)
    Txt(DAYS) = XNull(RstMain!DAYS)
    Txt(SprTel) = RstMain!SprTel
    Txt(SprDlr) = RstMain!SprDlr
    Txt(SprCust) = RstMain!SprCust
    Txt(LabTel) = RstMain!LabTel
    Txt(LabDlr) = RstMain!LabDlr
    Txt(LabCust) = RstMain!LabCust
    Txt(RateEditableYN) = IIf(VNull(RstMain!RateEditableYN) = 0, "No", "Yes")
    
    
    Select Case XNull(RstMain!Chrg_From)
        Case ""
            Txt(ChrgFrom) = ""
        Case 0
            Txt(ChrgFrom) = "Customer"
        Case 1
            Txt(ChrgFrom) = "Self(Dealer)"
    End Select
'    Ser_Typ(Char_From).ListIndex = IIf(IsNull(RstMain!chrg_from), -1, RstMain!chrg_from)
    Txt(An_Target) = XNull(RstMain!Serv_Target)
End If '"Customer","Self(Dealer)"

Grid_Hide
Exit Sub
Errloop:        MsgBox err.Description
End Sub
Private Sub TopCtrl1_eRef()
    RstHelp.Requery
End Sub
Private Sub TopCtrl1_eExit()
    RstMain.Cancel
    Unload Me
End Sub

Private Sub ServCodeSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "serv_type >=" & Chk_Text(XNull(Trim(Txt(S_T_Code))))
End Sub
Private Sub ServNameSearch()
If RstHelp.RecordCount <= 0 Then Exit Sub
RstHelp.MoveFirst
RstHelp.FIND "serv_Desc >=" & Chk_Text(XNull(Txt(S_T_Desc)))
End Sub

Private Sub Txt_Change(Index As Integer)
If ADDFLAG <> 0 Then
    Select Case Index
        Case S_T_Code, S_T_Desc
            If RstHelp.RecordCount = 0 Then Exit Sub
            'If DGServ.Visible = True Then DGServ.Visible = False
            DGServ.Visible = True
            DGServ.top = Txt(Index).top + Txt(Index).height + 10
            DGServ.left = Txt(Index).left
            DGServ.ZOrder 0
    End Select
End If
End Sub
Private Sub Txt_GotFocus(Index As Integer)
DGServ.Columns(0).width = 1000.1: DGServ.Columns(1).width = 3535.024: DGServ.Columns(2).width = 1000.1
Grid_Hide
Dim mBookMark
    Ctrl_GetFocus Txt(Index)
mFlag = 0
    If DGServ.Visible = True Then DGServ.Visible = False
    RST_BOF_EOF RstHelp
    Txt(Index).Tag = Txt(Index)
    Select Case Index
        Case S_T_Code, S_T_Desc
            If RstHelp.BOF Or RstHelp.EOF Then Exit Sub
    End Select
    Select Case Index
        Case ServType
            ListArray = Array("PDI", "Free Service", "Chargeable", "Accidental", "Denting & Painting", "Repeat Job", "AMC")
            Set mListItem = ListView_Items(ListView, Txt, ServType, ListArray, 7)
        Case ChrgFrom
            ListArray = Array("Customer", "Self(Dealer)")
            Set mListItem = ListView_Items(ListView, Txt, ChrgFrom, ListArray, 2)
        Case S_T_Code
            DGServ.Columns(2).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "Serv_type ASC"
            RstHelp.Bookmark = mBookMark
            ServCodeSearch
        Case S_T_Desc
            DGServ.Columns(0).width = 0
            mBookMark = RstHelp.Bookmark
            RstHelp.Sort = "serv_Desc ASC"
            RstHelp.Bookmark = mBookMark
            ServNameSearch
    End Select
    If Txt(Index) = "" Then Txt_Change Index
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim result As Boolean, I As Integer
Dim Txtdate As Boolean
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case ChrgFrom
            If KeyCode <> vbKeyEscape Then
                ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 600
            End If
    
    Case ServType
            If KeyCode <> vbKeyEscape Then
                ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1850
            End If
                If Txt(ServType) = "Free Service" Then
                    Txt(F_S_Code).Enabled = True
                Else
                    Txt(F_S_Code).Enabled = False
                End If
        If FrmList.Visible = False Then
        If KeyCode = 13 Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Then
            SendKeysA vbKeyTab, True
            KeyCode = 0
        ElseIf KeyCode = vbKeyUp Then
            SendKeys "+{Tab}"
            KeyCode = 0
        End If
        End If
End Select
If FrmList.Visible = False And Index <> ServType Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> LabDlr Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = LabDlr Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> S_T_Code Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> S_T_Desc Then
            If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
Call CheckQuote(keyascii)
Select Case Index
    Case F_S_Code
        NumPress Txt(F_S_Code), keyascii, 1, 0
    Case DAYS
        NumPress Txt(DAYS), keyascii, 3, 0
    Case An_Target
        NumPress Txt(An_Target), keyascii, 4, 0
    Case SprTel, SprDlr, SprCust, LabTel, LabDlr, LabCust
        NumPress Txt(An_Target), keyascii, 3, 2
    Case RateEditableYN
        If keyascii <> 13 Then
            If keyascii = Asc("y") Or keyascii = Asc("Y") Then
                Txt(Index) = "Yes"
            Else
                Txt(Index) = "No"
            End If
            keyascii = 0
        End If
    End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
mFlag = 0
If KeyCode <> vbKeyUp Or KeyCode <> vbKeyDown Or KeyCode <> 33 Or KeyCode <> 34 Then
Select Case Index
    Case ChrgFrom, ServType
         ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
    Case S_T_Code
        ServCodeSearch
    Case S_T_Desc
        ServNameSearch
End Select
End If
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Select Case Index
Case ServType
                If Txt(ServType) = "Free Service" Then
                    Txt(F_S_Code).Enabled = True
                Else
                    Txt(F_S_Code).Enabled = False
                End If
End Select
Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim I As Byte
    Select Case Index
        Case ServType
                If Txt(ServType) = "Free Service" Then
                    Txt(F_S_Code).Enabled = True
                    Txt(F_S_Code).SetFocus
                    For I = 8 To 13
                        Txt(I).Enabled = False
                    Next
                    Txt(SprTel) = "0.00": Txt(SprDlr) = "0.00": Txt(SprCust) = "100.00"
                    Txt(LabTel) = "0.00": Txt(LabDlr) = "0.00": Txt(LabCust) = "100.00"
                ElseIf Txt(ServType) = "AMC" Then
                    Txt(F_S_Code).Enabled = False
                    For I = 8 To 13
                        Txt(I).Enabled = True
                    Next
                Else
                    Txt(F_S_Code).Enabled = False
                    For I = 8 To 13
                        Txt(I).Enabled = False
                    Next
                    Txt(SprTel) = "0.00": Txt(SprDlr) = "0.00": Txt(SprCust) = "100.00"
                    Txt(LabTel) = "0.00": Txt(LabDlr) = "0.00": Txt(LabCust) = "100.00"
                End If
        Case S_T_Code
            Set Rst = GCn.Execute("SELECT * FROM Service_Type WHERE serv_type=" & Chk_Text(PubSiteCode + Trim(Txt(S_T_Code))))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Service Type Code Already Exists", vbInformation, "Validation": Txt(S_T_Code) = Txt(S_T_Code).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Serv_Type <> RstMain!Serv_Type Then MsgBox "Service Type Code Already Exists", vbInformation, "Validation": Txt(S_T_Code) = Txt(S_T_Code).Tag: Cancel = True: Exit Sub
                End If
            End If
        Case S_T_Desc
            Set Rst = GCn.Execute("SELECT * FROM Service_Type WHERE Serv_Desc=" & Chk_Text(Txt(S_T_Desc)))
            If ADDFLAG = 1 Then
                If Not Rst.EOF Then MsgBox "Service Description Already Exists", vbInformation, "Validation": Txt(S_T_Desc) = Txt(S_T_Desc).Tag: Cancel = True: Exit Sub
            ElseIf ADDFLAG = 2 Then
                If Not Rst.EOF Then
                    If Rst!Serv_Desc <> RstMain!Serv_Desc Then MsgBox "Service Description Already Exists", vbInformation, "Validation": Txt(S_T_Desc) = Txt(S_T_Desc).Tag: Cancel = True: Exit Sub
                End If
            End If
    End Select
Set Rst = Nothing
End Sub

Private Sub BlankText()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
Next I
'For i = 0 To Ser_Typ.Count - 1
'    Ser_Typ(i).ListIndex = -1
'Next i
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
'CmbOrder.Enabled = IIf(AddFlag = 1, True, False)
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
Next
'For i = 0 To Ser_Typ.Count - 1
'    Ser_Typ(i).Enabled = Enb
'Next
End Sub

Private Sub Grid_Hide()
    If FrmList.Visible = True Then FrmList.Visible = False
    If DGServ.Visible = True Then DGServ.Visible = False
End Sub


