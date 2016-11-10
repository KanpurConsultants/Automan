VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmCrmDmsInventoryImport 
   Caption         =   "CRM Inventory Import"
   ClientHeight    =   6840
   ClientLeft      =   1500
   ClientTop       =   2400
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   11805
   Begin VB.Frame Frame4 
      BackColor       =   &H00CFE0E0&
      Caption         =   "CRM DMS Environment Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Left            =   75
      TabIndex        =   0
      Top             =   4005
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8280
         TabIndex        =   2
         Top             =   2325
         Width           =   2070
      End
      Begin VB.CommandButton CmdCancel1 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8280
         TabIndex        =   1
         Top             =   2790
         Width           =   2070
      End
      Begin MSDataGridLib.DataGrid DgHelp 
         Height          =   1845
         Left            =   5850
         Negotiate       =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   7860
         Visible         =   0   'False
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   3254
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Code"
            Caption         =   "Name"
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
            Caption         =   "Name"
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
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3195.213
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6105
         Left            =   465
         TabIndex        =   4
         Top             =   390
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   10769
         _Version        =   393216
         Tabs            =   7
         TabsPerRow      =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Group Parameters"
         TabPicture(0)   =   "FrmCrmDmsInventoryImport.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Lbl(30)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Lbl(29)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Lbl(28)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Lbl(27)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Lbl(26)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Lbl(25)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Lbl(4)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Lbl(3)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Lbl(2)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Lbl(1)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Lbl(0)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Txt(32)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Txt(31)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Txt(30)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Txt(29)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Txt(28)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Txt(27)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Txt(6)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Txt(5)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Txt(4)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Txt(3)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Txt(2)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).ControlCount=   22
         TabCaption(1)   =   "Account Parameters 1"
         TabPicture(1)   =   "FrmCrmDmsInventoryImport.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lbl(36)"
         Tab(1).Control(1)=   "Lbl(35)"
         Tab(1).Control(2)=   "Lbl(33)"
         Tab(1).Control(3)=   "Lbl(32)"
         Tab(1).Control(4)=   "Lbl(23)"
         Tab(1).Control(5)=   "Lbl(11)"
         Tab(1).Control(6)=   "Lbl(10)"
         Tab(1).Control(7)=   "Lbl(9)"
         Tab(1).Control(8)=   "Lbl(8)"
         Tab(1).Control(9)=   "Lbl(7)"
         Tab(1).Control(10)=   "Lbl(6)"
         Tab(1).Control(11)=   "Lbl(5)"
         Tab(1).Control(12)=   "Lbl(22)"
         Tab(1).Control(13)=   "Txt(38)"
         Tab(1).Control(14)=   "Txt(37)"
         Tab(1).Control(15)=   "Txt(35)"
         Tab(1).Control(16)=   "Txt(34)"
         Tab(1).Control(17)=   "Txt(25)"
         Tab(1).Control(18)=   "Txt(13)"
         Tab(1).Control(19)=   "Txt(12)"
         Tab(1).Control(20)=   "Txt(11)"
         Tab(1).Control(21)=   "Txt(10)"
         Tab(1).Control(22)=   "Txt(9)"
         Tab(1).Control(23)=   "Txt(8)"
         Tab(1).Control(24)=   "Txt(7)"
         Tab(1).Control(25)=   "Txt(24)"
         Tab(1).ControlCount=   26
         TabCaption(2)   =   "Account Parameters 2"
         TabPicture(2)   =   "FrmCrmDmsInventoryImport.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Lbl(34)"
         Tab(2).Control(1)=   "Lbl(31)"
         Tab(2).Control(2)=   "Lbl(24)"
         Tab(2).Control(3)=   "Lbl(17)"
         Tab(2).Control(4)=   "Lbl(16)"
         Tab(2).Control(5)=   "Lbl(15)"
         Tab(2).Control(6)=   "Lbl(21)"
         Tab(2).Control(7)=   "Lbl(20)"
         Tab(2).Control(8)=   "Lbl(19)"
         Tab(2).Control(9)=   "Lbl(18)"
         Tab(2).Control(10)=   "Lbl(14)"
         Tab(2).Control(11)=   "Lbl(13)"
         Tab(2).Control(12)=   "Lbl(12)"
         Tab(2).Control(13)=   "Txt(36)"
         Tab(2).Control(14)=   "Txt(33)"
         Tab(2).Control(15)=   "Txt(26)"
         Tab(2).Control(16)=   "Txt(19)"
         Tab(2).Control(17)=   "Txt(18)"
         Tab(2).Control(18)=   "Txt(17)"
         Tab(2).Control(19)=   "Txt(23)"
         Tab(2).Control(20)=   "Txt(22)"
         Tab(2).Control(21)=   "Txt(21)"
         Tab(2).Control(22)=   "Txt(20)"
         Tab(2).Control(23)=   "Txt(16)"
         Tab(2).Control(24)=   "Txt(15)"
         Tab(2).Control(25)=   "Txt(14)"
         Tab(2).ControlCount=   26
         TabCaption(3)   =   "Other Detail"
         TabPicture(3)   =   "FrmCrmDmsInventoryImport.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label3"
         Tab(3).Control(1)=   "Label5"
         Tab(3).Control(2)=   "FGrid2"
         Tab(3).Control(3)=   "FGrid1"
         Tab(3).Control(4)=   "txtgrid1(0)"
         Tab(3).Control(5)=   "txtgrid2(0)"
         Tab(3).ControlCount=   6
         TabCaption(4)   =   "Spare Inventory Parameters"
         TabPicture(4)   =   "FrmCrmDmsInventoryImport.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Lbl(43)"
         Tab(4).Control(1)=   "Lbl(44)"
         Tab(4).Control(2)=   "Lbl(45)"
         Tab(4).Control(3)=   "Lbl(46)"
         Tab(4).Control(4)=   "Lbl(50)"
         Tab(4).Control(5)=   "Lbl(51)"
         Tab(4).Control(6)=   "Txt(45)"
         Tab(4).Control(7)=   "Txt(46)"
         Tab(4).Control(8)=   "Txt(47)"
         Tab(4).Control(9)=   "Txt(48)"
         Tab(4).Control(10)=   "Txt(52)"
         Tab(4).Control(11)=   "Txt(53)"
         Tab(4).ControlCount=   12
         TabCaption(5)   =   "Vehicle Inventory Parameters"
         TabPicture(5)   =   "FrmCrmDmsInventoryImport.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Lbl(37)"
         Tab(5).Control(1)=   "Lbl(38)"
         Tab(5).Control(2)=   "Lbl(39)"
         Tab(5).Control(3)=   "Lbl(40)"
         Tab(5).Control(4)=   "Lbl(41)"
         Tab(5).Control(5)=   "Lbl(42)"
         Tab(5).Control(6)=   "Txt(39)"
         Tab(5).Control(7)=   "Txt(40)"
         Tab(5).Control(8)=   "Txt(41)"
         Tab(5).Control(9)=   "Txt(42)"
         Tab(5).Control(10)=   "Txt(43)"
         Tab(5).Control(11)=   "Txt(44)"
         Tab(5).ControlCount=   12
         TabCaption(6)   =   "Job Bill IParameter"
         TabPicture(6)   =   "FrmCrmDmsInventoryImport.frx":00A8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Lbl(47)"
         Tab(6).Control(1)=   "Lbl(48)"
         Tab(6).Control(2)=   "Lbl(49)"
         Tab(6).Control(3)=   "Txt(49)"
         Tab(6).Control(4)=   "Txt(50)"
         Tab(6).Control(5)=   "Txt(51)"
         Tab(6).ControlCount=   6
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   53
            Left            =   -71895
            TabIndex        =   151
            Text            =   "Text1"
            Top             =   3990
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   52
            Left            =   -71895
            TabIndex        =   149
            Text            =   "Text1"
            Top             =   3675
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   51
            Left            =   -72690
            TabIndex        =   145
            Text            =   "Text1"
            Top             =   3045
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   50
            Left            =   -72690
            TabIndex        =   142
            Text            =   "Text1"
            Top             =   2415
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   49
            Left            =   -72690
            TabIndex        =   144
            Text            =   "Text1"
            Top             =   2730
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   48
            Left            =   -71910
            TabIndex        =   138
            Text            =   "Text1"
            Top             =   3360
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   47
            Left            =   -71910
            TabIndex        =   134
            Text            =   "Text1"
            Top             =   3045
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   46
            Left            =   -71910
            TabIndex        =   133
            Text            =   "Text1"
            Top             =   2730
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   45
            Left            =   -71910
            TabIndex        =   132
            Text            =   "Text1"
            Top             =   2415
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   44
            Left            =   -72030
            TabIndex        =   130
            Text            =   "Text1"
            Top             =   4005
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   43
            Left            =   -72030
            TabIndex        =   128
            Text            =   "Text1"
            Top             =   3690
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   42
            Left            =   -72030
            TabIndex        =   126
            Text            =   "Text1"
            Top             =   3375
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   41
            Left            =   -72030
            TabIndex        =   124
            Text            =   "Text1"
            Top             =   3060
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   40
            Left            =   -72030
            TabIndex        =   122
            Text            =   "Text1"
            Top             =   2745
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   39
            Left            =   -72030
            TabIndex        =   120
            Text            =   "Text1"
            Top             =   2430
            Width           =   3210
         End
         Begin VB.TextBox txtgrid2 
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
            Height          =   510
            Index           =   0
            Left            =   -71535
            MaxLength       =   40
            TabIndex        =   115
            Top             =   3285
            Visible         =   0   'False
            Width           =   945
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
            Height          =   510
            Index           =   0
            Left            =   -71520
            MaxLength       =   40
            TabIndex        =   42
            Top             =   5535
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   14
            Left            =   -72135
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   4440
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   15
            Left            =   -72135
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   4980
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   16
            Left            =   -72135
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   5250
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   20
            Left            =   -72135
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   3630
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   21
            Left            =   -72135
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   3090
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   22
            Left            =   -72135
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   2820
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   23
            Left            =   -72135
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   2550
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   17
            Left            =   -72135
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   3360
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   18
            Left            =   -72135
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   5520
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   19
            Left            =   -72135
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   4170
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   26
            Left            =   -72135
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   5790
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   33
            Left            =   -72135
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   4710
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   36
            Left            =   -72135
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   3900
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   24
            Left            =   -72240
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   5610
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   7
            Left            =   -72240
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   2370
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   8
            Left            =   -72240
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   5070
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   9
            Left            =   -72240
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   4800
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   10
            Left            =   -72240
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   4530
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   11
            Left            =   -72240
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   3450
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   12
            Left            =   -72240
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   3180
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   13
            Left            =   -72240
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   2910
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   25
            Left            =   -72240
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   5340
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   34
            Left            =   -72240
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   3720
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   35
            Left            =   -72240
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   2640
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   2775
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   2760
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   2775
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   3030
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   2775
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   3300
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   2775
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   3570
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   2775
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   3840
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   27
            Left            =   2775
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   5190
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   28
            Left            =   2775
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   4920
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   29
            Left            =   2775
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   4650
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   30
            Left            =   2775
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   4380
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   31
            Left            =   2775
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   4110
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   32
            Left            =   2775
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   5460
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   37
            Left            =   -72240
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   4260
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   38
            Left            =   -72240
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   3990
            Width           =   3210
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
            Height          =   1605
            Left            =   -74460
            TabIndex        =   43
            Top             =   4800
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   2831
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   3
            BackColorFixed  =   13623520
            ForeColorFixed  =   0
            BackColorSel    =   15718112
            ForeColorSel    =   12582912
            BackColorBkg    =   13623520
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid2 
            Height          =   1605
            Left            =   -74475
            TabIndex        =   116
            Top             =   2550
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   2831
            _Version        =   393216
            BackColor       =   16777215
            Cols            =   3
            BackColorFixed  =   13623520
            ForeColorFixed  =   0
            BackColorSel    =   15718112
            ForeColorSel    =   12582912
            BackColorBkg    =   13623520
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
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Oil Part (Missing Detail)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   51
            Left            =   -74715
            TabIndex        =   152
            Top             =   4005
            Width           =   2730
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default PartNo (Missing Detail)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   50
            Left            =   -74715
            TabIndex        =   150
            Top             =   3690
            Width           =   2670
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Labour Head"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   49
            Left            =   -74730
            TabIndex        =   147
            Top             =   3060
            Width           =   1755
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Mechanic"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   48
            Left            =   -74745
            TabIndex        =   146
            Top             =   2430
            Width           =   1500
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Supervisor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   47
            Left            =   -74730
            TabIndex        =   143
            Top             =   2745
            Width           =   1560
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local Sale Tax Form"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   46
            Left            =   -74715
            TabIndex        =   139
            Top             =   3375
            Width           =   1800
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Central Purchase Tax Form"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   45
            Left            =   -74730
            TabIndex        =   137
            Top             =   2430
            Width           =   2400
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local Purchase Tax Form"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   44
            Left            =   -74745
            TabIndex        =   136
            Top             =   2745
            Width           =   2235
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Central Sale Tax Form"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   43
            Left            =   -74730
            TabIndex        =   135
            Top             =   3060
            Width           =   1950
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax on Delivery Chg. (Y/N)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   42
            Left            =   -74625
            TabIndex        =   131
            Top             =   3705
            Width           =   2295
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transport Item"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   41
            Left            =   -74640
            TabIndex        =   129
            Top             =   4005
            Width           =   1245
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transport Item"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   40
            Left            =   -74625
            TabIndex        =   127
            Top             =   3390
            Width           =   1245
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Discount Item"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   39
            Left            =   -74640
            TabIndex        =   125
            Top             =   3075
            Width           =   1200
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local Purchase Tax Form"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   38
            Left            =   -74640
            TabIndex        =   123
            Top             =   2760
            Width           =   2235
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Central Purchase Tax Form"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   37
            Left            =   -74640
            TabIndex        =   121
            Top             =   2445
            Width           =   2400
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H80000007&
            Caption         =   "Automan Bank A/c ## Dms Bank A/c"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   -74460
            TabIndex        =   119
            Top             =   4500
            Width           =   5505
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000007&
            Caption         =   "Automan Supplier ## Dms Supplier"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   -74475
            TabIndex        =   117
            Top             =   2250
            Width           =   5505
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Purchase A/c..............."
            Height          =   195
            Index           =   12
            Left            =   -74640
            TabIndex        =   80
            Top             =   4470
            Width           =   2700
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CST A/c................................"
            Height          =   195
            Index           =   13
            Left            =   -74625
            TabIndex        =   79
            Top             =   5010
            Width           =   2625
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Round Off A/c........................."
            Height          =   195
            Index           =   14
            Left            =   -74625
            TabIndex        =   78
            Top             =   5280
            Width           =   2700
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Purchase A/c 12.5 %........."
            Height          =   195
            Index           =   18
            Left            =   -74625
            TabIndex        =   77
            Top             =   3660
            Width           =   2910
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Bank A/c................."
            Height          =   195
            Index           =   19
            Left            =   -74625
            TabIndex        =   76
            Top             =   3120
            Width           =   2475
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Bank A/c................"
            Height          =   195
            Index           =   20
            Left            =   -74625
            TabIndex        =   75
            Top             =   2850
            Width           =   2310
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local State Name..............."
            Height          =   195
            Index           =   21
            Left            =   -74625
            TabIndex        =   74
            Top             =   2580
            Width           =   2400
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Workshop Bank A/c..........."
            Height          =   195
            Index           =   15
            Left            =   -74625
            TabIndex        =   73
            Top             =   3390
            Width           =   2355
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other Charges A/c.................."
            Height          =   195
            Index           =   16
            Left            =   -74625
            TabIndex        =   72
            Top             =   5550
            Width           =   2685
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Cst Purchase A/c...................."
            Height          =   195
            Index           =   17
            Left            =   -74640
            TabIndex        =   71
            Top             =   4200
            Width           =   3240
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Discount A/c..........................."
            Height          =   195
            Index           =   24
            Left            =   -74625
            TabIndex        =   70
            Top             =   5820
            Width           =   2700
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Cst Purchase A/c........."
            Height          =   195
            Index           =   31
            Left            =   -74625
            TabIndex        =   69
            Top             =   4740
            Width           =   2685
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Purchase A/c 4 %........."
            Height          =   195
            Index           =   34
            Left            =   -74625
            TabIndex        =   68
            Top             =   3930
            Width           =   2640
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Tax A/c................"
            Height          =   195
            Index           =   22
            Left            =   -74460
            TabIndex        =   67
            Top             =   5655
            Width           =   2325
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Sale A/c................."
            Height          =   195
            Index           =   5
            Left            =   -74445
            TabIndex        =   66
            Top             =   2415
            Width           =   2310
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Cash A/c..............."
            Height          =   195
            Index           =   6
            Left            =   -74460
            TabIndex        =   65
            Top             =   5115
            Width           =   2355
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Cash A/c................."
            Height          =   195
            Index           =   7
            Left            =   -74445
            TabIndex        =   64
            Top             =   4845
            Width           =   2370
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Workshop Cash A/c.........."
            Height          =   195
            Index           =   8
            Left            =   -74445
            TabIndex        =   63
            Top             =   4575
            Width           =   2295
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vat A/c (Sale)............................."
            Height          =   195
            Index           =   9
            Left            =   -74445
            TabIndex        =   62
            Top             =   3495
            Width           =   2955
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Sale A/c................."
            Height          =   195
            Index           =   10
            Left            =   -74445
            TabIndex        =   61
            Top             =   3225
            Width           =   2415
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lubricant Sale A/c..............."
            Height          =   195
            Index           =   11
            Left            =   -74445
            TabIndex        =   60
            Top             =   2955
            Width           =   2460
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Labour A/c......................."
            Height          =   195
            Index           =   23
            Left            =   -74460
            TabIndex        =   59
            Top             =   5385
            Width           =   2310
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vat 4 % A/c (Sale)............................."
            Height          =   195
            Index           =   32
            Left            =   -74445
            TabIndex        =   58
            Top             =   3765
            Width           =   3360
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Vat 4% Sale A/c................."
            Height          =   195
            Index           =   33
            Left            =   -74460
            TabIndex        =   57
            Top             =   2685
            Width           =   3000
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Debtor Group........."
            Height          =   195
            Index           =   0
            Left            =   570
            TabIndex        =   56
            Top             =   2805
            Width           =   2280
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Workshop Debtor Group...."
            Height          =   195
            Index           =   1
            Left            =   570
            TabIndex        =   55
            Top             =   3075
            Width           =   2325
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Debtor Group........"
            Height          =   195
            Index           =   2
            Left            =   570
            TabIndex        =   54
            Top             =   3345
            Width           =   2325
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Creditor Group........."
            Height          =   195
            Index           =   3
            Left            =   570
            TabIndex        =   53
            Top             =   3615
            Width           =   2400
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Creditor Group....."
            Height          =   195
            Index           =   4
            Left            =   570
            TabIndex        =   52
            Top             =   3885
            Width           =   2265
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VAT Group..........................."
            Height          =   195
            Index           =   25
            Left            =   570
            TabIndex        =   51
            Top             =   5235
            Width           =   2550
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Purchase Group........."
            Height          =   195
            Index           =   26
            Left            =   570
            TabIndex        =   50
            Top             =   4965
            Width           =   2580
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Sale Group.................."
            Height          =   195
            Index           =   27
            Left            =   570
            TabIndex        =   49
            Top             =   4695
            Width           =   2715
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Purchase Group...."
            Height          =   195
            Index           =   28
            Left            =   570
            TabIndex        =   48
            Top             =   4425
            Width           =   2175
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Sale Group..................."
            Height          =   195
            Index           =   29
            Left            =   570
            TabIndex        =   47
            Top             =   4155
            Width           =   2670
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Tax Group.................."
            Height          =   195
            Index           =   30
            Left            =   555
            TabIndex        =   46
            Top             =   5490
            Width           =   2685
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vat 4 % A/c (Purchase)............................."
            Height          =   195
            Index           =   35
            Left            =   -74445
            TabIndex        =   45
            Top             =   4305
            Width           =   3765
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vat A/c (Purchase)............................."
            Height          =   195
            Index           =   36
            Left            =   -74445
            TabIndex        =   44
            Top             =   4035
            Width           =   3360
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Date Criteria"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   8880
      TabIndex        =   108
      Top             =   5520
      Visible         =   0   'False
      Width           =   3150
      Begin VB.TextBox Txt 
         Height          =   285
         Index           =   0
         Left            =   1410
         TabIndex        =   110
         Text            =   "Text1"
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox Txt 
         Height          =   285
         Index           =   1
         Left            =   1410
         TabIndex        =   109
         Text            =   "Text1"
         Top             =   615
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         Height          =   195
         Left            =   390
         TabIndex        =   112
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         Height          =   195
         Left            =   375
         TabIndex        =   111
         Top             =   645
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Select Ms-Excel File..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   91
      Top             =   390
      Visible         =   0   'False
      Width           =   11595
      Begin VB.CommandButton CmdSpareSale 
         Caption         =   "Spare Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   8625
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   225
         Width           =   1425
      End
      Begin VB.CommandButton CmdJobBill 
         Caption         =   "Job Bill"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   7215
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   225
         Width           =   1425
      End
      Begin VB.CommandButton CmdSparePurchase 
         Caption         =   "Spare Purchase"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   5805
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   225
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Vehicle Money Receipt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   5
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   1740
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Spare Money Receipt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   6
         Left            =   8670
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   1740
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Vehicle Sale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   2
         Left            =   7260
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1740
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Vehicle Purchase"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   9
         Left            =   5850
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   1740
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "WorkShop Sale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   1
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   1740
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Spare Counter Sale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   0
         Left            =   3030
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   1740
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Spare Purchase"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   3
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   1740
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Part Master"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   13
         Left            =   4395
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   225
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Unit Master"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   12
         Left            =   2985
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   225
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   8
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   1740
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Supplier Payment"
         Height          =   525
         Index           =   4
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   4200
         Width           =   1320
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Spare Sale Return"
         Height          =   540
         Index           =   7
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   4200
         Width           =   1425
      End
      Begin VB.CommandButton CmdVehiclePurchase 
         Caption         =   "Vehicle Purchase (Inventory)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   225
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Model Import"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   11
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   225
         Width           =   1425
      End
      Begin MSComctlLib.ProgressBar Prg 
         Height          =   270
         Left            =   165
         TabIndex        =   101
         Top             =   1230
         Visible         =   0   'False
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label LblVPrefix 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V.Prefix"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   10560
         TabIndex        =   107
         Top             =   2805
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Maching Account Names"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   6660
      TabIndex        =   87
      Top             =   5940
      Visible         =   0   'False
      Width           =   8985
      Begin VB.CommandButton CmdOk 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2865
         TabIndex        =   89
         Top             =   1995
         Width           =   1110
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4095
         TabIndex        =   88
         Top             =   1995
         Width           =   1110
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
         Height          =   1605
         Left            =   120
         TabIndex        =   90
         Top             =   300
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   2831
         _Version        =   393216
         BackColorFixed  =   13623520
         BackColorBkg    =   13623520
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Error Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   195
      TabIndex        =   81
      Top             =   3390
      Visible         =   0   'False
      Width           =   11625
      Begin VB.CheckBox ChkAllErr 
         BackColor       =   &H00CFE0E0&
         Caption         =   "All Types"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6150
         TabIndex        =   85
         Top             =   15
         Width           =   1170
      End
      Begin VB.CommandButton CmdDelErr 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   10245
         TabIndex        =   84
         Top             =   2565
         Width           =   1185
      End
      Begin VB.CommandButton CmdDelErr 
         Caption         =   "Show All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   9045
         TabIndex        =   83
         Top             =   2565
         Width           =   1185
      End
      Begin VB.TextBox TxtShow 
         Appearance      =   0  'Flat
         Height          =   915
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Text            =   "FrmCrmDmsInventoryImport.frx":00C4
         Top             =   1995
         Width           =   8865
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FgridErr 
         Height          =   1620
         Left            =   120
         TabIndex        =   86
         Top             =   270
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   2858
         _Version        =   393216
         BackColorFixed  =   13623520
         BackColorBkg    =   13623520
         AllowUserResizing=   3
         Appearance      =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   165
      Top             =   6870
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   75
      Top             =   6735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   3  'Align Left
      Height          =   375
      Left            =   0
      TabIndex        =   113
      Top             =   0
      Visible         =   0   'False
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   661
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Automan Supplier ## Dms Supplier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   0
      TabIndex        =   118
      Top             =   0
      Width           =   5505
   End
   Begin VB.Label LblTimer 
      Caption         =   "Label3"
      Height          =   480
      Left            =   1800
      TabIndex        =   114
      Top             =   6180
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "FrmCrmDmsInventoryImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Master As ADODB.Recordset, RsNew1 As ADODB.Recordset, RsNew As ADODB.Recordset, RsTemp As ADODB.Recordset
Dim ExcelGcn1 As ADODB.Connection, ExcelGcn2 As ADODB.Connection
Dim DmsConn As ADODB.Connection
Dim RsDmsEnviro As ADODB.Recordset
Dim RsHelp As ADODB.Recordset
Dim RsCity As ADODB.Recordset
Dim RsState As ADODB.Recordset
Dim RsSubGroup As ADODB.Recordset
Dim RsTaxForm As ADODB.Recordset
Dim RsADItem As ADODB.Recordset
Dim RsAcGroup As ADODB.Recordset
Dim RsMechanic As ADODB.Recordset
Dim RsSupervisor As ADODB.Recordset
Dim RsLabour As ADODB.Recordset

Dim RsDms           As ADODB.Recordset
   
Dim CodeCntSpare As Long
Dim CodeCntLabour As Long
Public mFormType As Byte
Const ImportForm As Byte = 1
Const Enviro     As Byte = 2


Dim mFlag As Byte
Dim GridKey As Integer
Dim mIsAnySubCodeCreated As Boolean


'Fgrid1 Constants
Private Const F1_BankAc     As Byte = 1
Private Const F1_DmsCode    As Byte = 2
Private Const F1_BankAcCode As Byte = 3


'Fgrid1 Constants
Private Const F2_SupplierAc     As Byte = 1
Private Const F2_DmsCode    As Byte = 2
Private Const F2_SupplierAcCode As Byte = 3


'Txt Constants
Const FromDate              As Byte = 0
Const ToDate                As Byte = 1
Const SprDebtorGroupCode    As Byte = 2
Const WsDebtorGroupCode     As Byte = 3
Const VehDebtorGroupCode    As Byte = 4
Const SprCreditorGroupCode  As Byte = 5
Const VehCreditorGroupCode  As Byte = 6
Const SprSaleAc             As Byte = 7
Const VehCashAc             As Byte = 8
Const SprCashAc             As Byte = 9
Const WsCashAc              As Byte = 10
Const VatAc                 As Byte = 11
Const VehSaleAc             As Byte = 12
Const LubSaleAc             As Byte = 13
Const VehPurchaseAc         As Byte = 14
Const CstAc                 As Byte = 15
Const ROffAc                As Byte = 16
Const WsBankAc              As Byte = 17
Const OtherChargesAc        As Byte = 18
Const SprCstPurchaseAc      As Byte = 19
Const SprPurchaseAc         As Byte = 20
Const VehBankAc             As Byte = 21
Const SprBankAc             As Byte = 22
Const LocalStateName        As Byte = 23
Const ServTaxAc             As Byte = 24
Const LabourAc              As Byte = 25
Const DiscountAc            As Byte = 26
Const VatGroupCode          As Byte = 27
Const VehPurGroupCode       As Byte = 28
Const VehSaleGroupCode      As Byte = 29
Const SprPurGroupCode       As Byte = 30
Const SprSaleGroupCode      As Byte = 31
Const ServiceTaxGroupCode   As Byte = 32
Const VehCstPurchaseAc      As Byte = 33
Const Vat4Ac                As Byte = 34
Const SprSaleVat4Ac         As Byte = 35
Const SprPurchase4Ac        As Byte = 36
Const Vat4InputAc           As Byte = 37
Const VatInputAc            As Byte = 38
Const VehicleCentralPurchaseTaxForm            As Byte = 39
Const VehicleLocalPurchaseTaxForm              As Byte = 40
Const VehiclePurchaseDiscountItem              As Byte = 41
Const VehiclePurchaseTransportItem             As Byte = 42
Const VehicleTaxOnDeliveryCharges              As Byte = 43
Const SpareCentralPurchaseTaxForm              As Byte = 45
Const SpareLocalPurchaseTaxForm                As Byte = 46
Const SpareCentralSaleTaxForm                  As Byte = 47
Const SpareLocalSaleTaxForm                    As Byte = 48
Const DefaultSupervisor         As Byte = 49
Const DefaultMechanic           As Byte = 50
Const DefaultLabourHead         As Byte = 51
Const DefaultPartNo             As Byte = 52
Const DefaultOilPartNo            As Byte = 53




'Cmd Constants
Const ImpSprCounterSale     As Byte = 0
Const ImpWorkShopSale       As Byte = 1
Const ImpVehcleSale         As Byte = 2
Const ImpSparePurchase      As Byte = 3
Const ImpSupplierPayment    As Byte = 4
Const ImpMoneyRectVehicle   As Byte = 5
Const ImpMoneyRectSpare     As Byte = 6
Const ImpSprSaleReturn      As Byte = 7
Const ImpCustomer           As Byte = 8
Const ImpVehiclePurchase    As Byte = 9
Const ImpModelImport    As Byte = 11
Const ImpUnitImport    As Byte = 12
Const ImpPartImport    As Byte = 13

Dim GSQL$

Dim CopyCnt As Long, ErrorCnt As Long
'FGrid Constants
Const FSel      As Byte = 0
Const fname     As Byte = 1
Const FFName    As Byte = 2
Const FAdd1     As Byte = 3
Const FAdd2     As Byte = 4
Const FAdd3     As Byte = 5
Const FCity     As Byte = 6
Const FSubCode  As Byte = 7



'FGridErr Constants
Const FErr_Cat          As Byte = 1
Const FErr_DmsRef       As Byte = 2
Const FErr_Narration    As Byte = 3



Private Sub CmdCancel_Click()
Dim I As Integer

    For I = 1 To FGrid.Rows - 1
        FGrid.TextMatrix(I, FSel) = ""
    Next I
    
    Frame3.Visible = False
End Sub



Private Sub CmdDelErr_Click(Index As Integer)
Dim mCondStr$
Dim RsTemp As ADODB.Recordset
    Select Case UCase(CmdDelErr(Index).CAPTION)
        Case "DELETE"
            If ChkAllErr.Value = 0 Then mCondStr = " Where Cat='" & FgridErr.TextMatrix(1, FErr_Cat) & "'"
            GCn.Execute "Delete from DmsErrLog " & mCondStr & ""
        Case "SHOW ALL"
            Set RsTemp = GCn.Execute("Select Cat As Category, [Key] as Dms_Reference, Narration From DmsErrLog " & mCondStr)
            Set FgridErr.DataSource = RsTemp
            Ini_Grid FgridErr
    End Select
End Sub



Private Sub CmdImport_Click(Index As Integer)
    Dim X As Long
    Dim RsTemp          As ADODB.Recordset
'    Dim RsDms           As adodb.Recordset
    Dim mCnt            As ADODB.Recordset
    Dim mSubGroupCounter    As Long
    Dim mSubCode$, mDmsSubCode$, mQry$, mNarr$, mLocalCentral$, mCondStr$
    Dim mFileName$, mFileTitle$, MState$, mCashCredit$, mVouCat$, mInvoiceNo$
    Dim mNetLabour       As Double
    Dim mTaxableLabour   As Double
    Dim mServTaxLabour   As Double
    Dim mNetAmount       As Double
    Dim mSaleAmt         As Double
    Dim mSprSaleAmt      As Double
    Dim mSprSaleVat4Amt  As Double
    Dim mLubeSaleAmt     As Double
    Dim mVatAmt          As Double
    Dim mVat12           As Double
    Dim mVat4            As Double
    Dim mCstAmt          As Double
    Dim mDiff            As Double
    Dim mDiffLab         As Double
    Dim mPurchaseAmt     As Double
    Dim mPurchaseAmt12   As Double
    Dim mPurchaseAmt4    As Double
    Dim mBankAcCode      As String
    Dim mDiscount        As Double
    Dim mLabDiscount     As Double
    Dim mOtherCharges    As Double
    Dim mLabOtherCharges As Double
    Dim Rstcolor As ADODB.Recordset
    Dim rstmodel As ADODB.Recordset
    
    Dim mColorCode As String
    Dim mModelCode As String
    
      Set rstmodel = GCn.Execute("SELECT * FROM Model")
    Set Rstcolor = GCn.Execute("Select * from ColMast ")
'On Error GoTo DispErr
                    
    
    'If XNull(RsDmsEnviro!CashAc) = "" Then MsgBox "Plz Define CashAc In DmsEnviro": Exit Sub
    mSubGroupCounter = G_CompCn.Execute("Select SubGroupAcCode From SubGroupCounter").Fields(0)
    
    CD1.FileName = ""
    
    Call SelectFile
    mFileName = CD1.FileName
    mFileTitle = CD1.FileTitle
    If mFileName = "" Then Exit Sub
    mFileTitle = mID(mFileTitle, 1, Len(mFileTitle) - 4)
    Set DmsConn = New Connection
    DmsConn.CursorLocation = adUseClient
    DmsConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mFileName & ";Extended Properties=Excel 8.0"
    
    Set RsDms = DmsConn.Execute("Select * from [" & mFileTitle & "$]")
    

    
    
    For X = 1 To 9999
        Lbl(0).Refresh
    Next X
    
    
    With RsDms
    
                Select Case Index
                    Case ImpCustomer
                        If .RecordCount > 0 Then
                            Prg.Value = 0
                            Prg.Visible = True
                            
                            mVouCat = "SubGroup"
                            Do Until .EOF
                                GCn.BeginTrans
                                G_FaCn.BeginTrans
                                If GCn.Execute("Select Count(*) From DmsSubGroup Where DmsSubCode='" & XNull(.Fields("Customer_Code")) & "'").Fields(0).Value = 0 Then
                                        Set RsTemp = GCn.Execute("Select AutomanSite From DmsSite Where DmsDivision='" & XNull(.Fields("Division")) & "'")
                                        If RsTemp.RecordCount > 0 Then
                                            GCn.Execute "Delete From DmsErrLog Where [Key] = '" & XNull(.Fields("Customer_Code")) & "' "
                                            GCn.Execute "Insert Into DmsSubGroup(DmsSubCode, Name, Add1, Add2, City, " & _
                                                                "PinCode, State, Phone, Fax, Email, " & _
                                                                "[Group], Division) " & _
                                                        "Values ('" & XNull(.Fields("Customer_Code")) & "', '" & left(XNull(.Fields("Customer_Name")), 50) & "', '" & left(XNull(.Fields("Addr_L_1")), 50) & "', '" & left(XNull(.Fields("Addr_L_2")), 50) & "', '" & left(XNull(.Fields("City")), 50) & "',  " & _
                                                                "'" & left(XNull(.Fields("Pin_Code")), 6) & "', '" & left(XNull(.Fields("State")), 2) & "', '" & left(Trim(XNull(.Fields("Phone"))), 35) & "', '" & left(Trim(XNull(.Fields("Fax"))), 24) & "', '" & left(Trim(XNull(.Fields("Email"))), 50) & "', " & _
                                                                "'" & XNull(.Fields("Group")) & "', '" & XNull(.Fields("Division")) & "')"
                                        Else
                                            CreateErrLog mVouCat, XNull(.Fields("Customer_Code")), XNull(.Fields("Division")) & " Not Defined In DmsDivision Table"
                                        End If
                                End If
                                
                                
                                
                                If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                .MoveNext
                                GCn.CommitTrans
                                G_FaCn.CommitTrans
                            Loop
                        End If
        '                Set RsSubGroup = GCn.Execute("Select '' As Sel, S.Name, S.Add1 As Address1, S.Add2 As Address2, S.Add3 As Address3, S.CityName As City From SubGroup S Left Join City C On C.CityCode=S.CityCode")
        '                If .RecordCount > 0 Then
        '                    Do Until .EOF
        '                        Set RsSubGroup.Filter = adFilterNone
        '                        Set RsSubGroup.Filter = "Replace(Replace(Replace(Name,' ',''),'.',''),',','') Like '" & Replace(Replace(Replace(!Name, " ", ""), ".", ""), ",", "") & "*'"
        '                        If RsSubGroup.RecordCount > 0 Then
        '                            Set FGrid.DataSource = RsSubGroup
        '                            Ini_Grid
        '                            FGrid.Visible = True
        '                        Else
        '                        End If
        '                    Loop
        '                End If
            
                
                
                    Case ImpSprCounterSale
                        If XNull(RsDmsEnviro!SprDebtorGroupCode) = "" Then MsgBox "Plz Define SprDebtorGroupCode In DmsEnviro": Exit Sub
                        
                        
                        
                        If ChkFieldExist(RsDms, "Account_Code") And ChkFieldExist(RsDms, "Parts_Invoice_Amount") And _
                           ChkFieldExist(RsDms, "Parts Amount") And ChkFieldExist(RsDms, "Lubricant Amount") And _
                           ChkFieldExist(RsDms, "Vat") And ChkFieldExist(RsDms, "Invoice_No") And ChkFieldExist(RsDms, "Invoice_Date") And _
                           ChkFieldExist(RsDms, "Mode Of Payment") And ChkFieldExist(RsDms, "Division") And ChkFieldExist(RsDms, "Customer_Code") Then
                        
                                mVouCat = "Spare Sale"
                                                            
                                .Filter = adFilterNone
                                .Filter = "Invoice_Status='New'"
                                
                                If .RecordCount > 0 Then
                                    Prg.Value = 0
                                    Prg.Visible = True
                                    Do Until .EOF
                                        GCn.BeginTrans
                                        G_FaCn.BeginTrans
                                            mInvoiceNo = XNull(!Invoice_No)
                                            mDmsSubCode = IIf(XNull(!Account_Code) = "", XNull(!Customer_Code), XNull(!Account_Code))
                                            GCn.Execute "Delete From DmsErrLog Where [Key]='" & mInvoiceNo & "'"
                                            mSubCode = AutomanSubcode(mDmsSubCode, RsDmsEnviro!SprDebtorGroupCode, "Customer")
                                            If mSubCode = "" Then 'And .Fields("Mode Of Payment") = "Credit" Then
                                                Call CreateErrLog(mVouCat, !Invoice_No, "Account_Code - " & !Account_Code & " Not Found In Automan")
                                            Else
                                                mNetAmount = eVal(.Fields("Parts_Invoice_Amount"))
                                                mSprSaleAmt = eVal(.Fields("Parts Amount"))
                                                mLubeSaleAmt = eVal(.Fields("Lubricant Amount"))
                                                mVatAmt = eVal(.Fields("Total_Tax_Amount"))
                                                mOtherCharges = eVal(.Fields("Other Charges"))
                                                mNarr = "Counter Sale Against Invoice No " & mInvoiceNo
                                                
                                                If Format(mNetAmount, "0.0") = Format(mSprSaleAmt + mLubeSaleAmt + mVatAmt + mOtherCharges, "0.0") Then
                                                    mNetAmount = Round(mSprSaleAmt + mLubeSaleAmt + mVatAmt + mOtherCharges, 2)
                                                    If SprCounterSale(.Fields("Mode Of Payment"), mSubCode, mNetAmount, mSprSaleAmt, mLubeSaleAmt, mVatAmt, mNarr, !Invoice_Date, !Invoice_No, !division, mOtherCharges) = False Then
                                                        Call CreateErrLog(mVouCat, mInvoiceNo, "Error In Ledger Posting")
                                                    End If
                                                Else
                                                    Call CreateErrLog(mVouCat, mInvoiceNo, "Total Amount : " & mNetAmount & ", Not Match With Parts Amt : " & mSprSaleAmt & " + Lubricant Amt : " & mLubeSaleAmt & " + Vat Amt : " & mVatAmt & " + Other Charges : " & mOtherCharges)
                                                End If
                                            End If
                                            If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                            .MoveNext
                                        GCn.CommitTrans
                                        G_FaCn.CommitTrans
                                    Loop
                                End If
                        End If
                                                    
                    
                    
                    
                    
                    
                    Case ImpWorkShopSale
                            If XNull(RsDmsEnviro!WsDebtorGroupCode) = "" Then MsgBox "Plz Define WsDebtorGroupCode In DmsEnviro": Exit Sub
                            
                            If ChkFieldExist(RsDms, "Customer_Code") And ChkFieldExist(RsDms, "Total Labour Amount") And _
                               ChkFieldExist(RsDms, "Labour Invoice Amount") And ChkFieldExist(RsDms, "Service Tax") And _
                               ChkFieldExist(RsDms, "Spares Invoice Amount") And ChkFieldExist(RsDms, "Invoice_No") And ChkFieldExist(RsDms, "Invoice_Date") And _
                               ChkFieldExist(RsDms, "Mode Of Payment") And ChkFieldExist(RsDms, "Division") And ChkFieldExist(RsDms, "Customer_Code") And _
                               ChkFieldExist(RsDms, "Parts Amount") And ChkFieldExist(RsDms, "Vat") And ChkFieldExist(RsDms, "Output VAT @ 12#5%") And _
                               ChkFieldExist(RsDms, "Output VAT @ 4%") And ChkFieldExist(RsDms, "Narration") Then
                            

                            
                                    mVouCat = "Workshop Sale"
                                    
                                    
                                    .Filter = adFilterNone
                                    .Filter = "Invoice_Status='New'"
                                    
                                    If .RecordCount > 0 Then
                                        Prg.Value = 0
                                        Prg.Visible = True
                    
                                        Do Until .EOF
                                            GCn.BeginTrans
                                            G_FaCn.BeginTrans
                                            mInvoiceNo = XNull(!Invoice_No)
                                            If StrCmp(mInvoiceNo, "UjwaAu-DH-0809-01100") Then
                                                MsgBox ""
                                            End If
                                            mDmsSubCode = IIf(XNull(!Account_Code) = "", XNull(!Customer_Code), XNull(!Account_Code))
                                            GCn.Execute "Delete From DmsErrLog Where [Key]='" & mInvoiceNo & "'"
                                            mSubCode = AutomanSubcode(mDmsSubCode, RsDmsEnviro!SprDebtorGroupCode, "Customer")
                                            If mSubCode = "" Then
                                                Call CreateErrLog(mVouCat, mInvoiceNo, "Account_Code - " & !Account_Code & " Not Found In Automan")
                                            Else
                                                mNetLabour = eVal(.Fields("Total Labour Amount"))
                                                mTaxableLabour = eVal(.Fields("Labour Invoice Amount"))
                                                mServTaxLabour = eVal(.Fields("Service Tax")) + eVal(.Fields("S n HE")) + eVal(.Fields("Cess Tax"))
                                                mLabOtherCharges = eVal(.Fields("Other Charges Labour"))
                                                mLabDiscount = eVal(.Fields("Discount Labour"))
                                                
                                                mNetAmount = eVal(.Fields("Spares Invoice Amount"))
                                                mSprSaleAmt = eVal(.Fields("VAT Assessible Amt 1")) + eVal(.Fields("VAT Assessible Amt 2")) + eVal(.Fields("VAT Assessible Amt 3"))
                                                'mLubeSaleAmt = eVal(.Fields("Lubricant Amount"))
                                                mSprSaleVat4Amt = eVal(.Fields("VAT Assessible Amt 4"))
                                                mVatAmt = eVal(.Fields("Vat"))
                                                mVat12 = eVal(.Fields("Output VAT @ 12#5%"))
                                                mVat4 = eVal(.Fields("Output VAT @ 4%"))
                                                mDiscount = eVal(.Fields("Discount Job Parts"))
                                                mOtherCharges = eVal(.Fields("Other Charges"))
                                                mNarr = XNull(!Narration) & " Invoice No " & XNull(!Invoice_No)
                                                
                                                mDiff = Format(mNetAmount + mNetLabour, "0.0") - Format(mSprSaleAmt + mSprSaleVat4Amt + mVatAmt + mTaxableLabour + mServTaxLabour + mOtherCharges - (mDiscount + mLabDiscount), "0.0")
                                                mDiffLab = Format(mNetLabour, "0.0") - Format(mTaxableLabour + mServTaxLabour, "0.0")
                                                
                                                
                                                If Abs(mDiff) < 0.9 Then
                                                    mNetAmount = Round(mSprSaleAmt + mSprSaleVat4Amt + mVat4 + mVat12 + mOtherCharges - mDiscount, 2)
                                                    mNetLabour = Round(mTaxableLabour + mServTaxLabour + mLabOtherCharges - mLabDiscount, 2)
                                                    
                                                    If WorkShopSale(.Fields("Mode Of Payment"), mSubCode, mNetAmount, mSprSaleAmt, mSprSaleVat4Amt, mVat12, mNarr, !Invoice_Date, mNetLabour, mTaxableLabour, mServTaxLabour, !Invoice_No, !division, mOtherCharges, mLabOtherCharges, mDiscount, mLabDiscount, mVat4) = False Then
                                                        Call CreateErrLog(mVouCat, mInvoiceNo, "Error In Ledger Posting")
                                                    End If
                                                Else
                                                    Call CreateErrLog(mVouCat, mInvoiceNo, "Total Spare Amount : " & mNetAmount & " + Total Labour : " & mNetLabour & ", Not Match With Parts Amt : " & mSprSaleAmt & " + Lubricant Amt : " & mLubeSaleAmt & " + Vat Amt : " & mVatAmt & " + Taxable Labour : " & mTaxableLabour & "+ Service Tax : " & mServTaxLabour & " + OtherCharges : " & mOtherCharges & " - Discount : " & mDiscount & ", DIFFERENCE : " & mDiff)
                                                End If
                                            End If
                                            
                                            If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                            .MoveNext
                                            GCn.CommitTrans
                                            G_FaCn.CommitTrans
                                        Loop
                                        
                                    End If
                            End If
                    
                    Case ImpSparePurchase
                        If XNull(RsDmsEnviro!SprCreditorGroupCode) = "" Then MsgBox "Plz Define SprCreditorGroupCode In DmsEnviro": Exit Sub
                        If XNull(RsDmsEnviro!VatInputAc) = "" Then MsgBox "Plz Define Vat A/c Purchase In DmsEnviro": Exit Sub
                        If XNull(RsDmsEnviro!Vat4InputAc) = "" Then MsgBox "Plz Define Vat A/c Purchase In DmsEnviro": Exit Sub
                        
                        If ChkFieldExist(RsDms, "Vendor Name") And ChkFieldExist(RsDms, "Total_Invoice_Amount") And _
                           ChkFieldExist(RsDms, "Total_Tax_Amount") And ChkFieldExist(RsDms, "Net_Amount") And _
                           ChkFieldExist(RsDms, "Invoice #") And ChkFieldExist(RsDms, "Invoice_Date") And _
                           ChkFieldExist(RsDms, "Division") Then
                        
                                If .RecordCount > 0 Then
                                    Prg.Value = 0
                                    Prg.Visible = True
                                    Do Until .EOF
                                        GCn.BeginTrans
                                        G_FaCn.BeginTrans
                                            mInvoiceNo = .Fields("Invoice #")
                                            GCn.Execute "Delete From DmsErrLog Where [Key]='" & mInvoiceNo & "'"
                                            mSubCode = AutomanSubcode(XNull(.Fields("Vendor Name")), XNull(RsDmsEnviro!SprCreditorGroupCode), "Supplier")
                                            If mSubCode = "" Then
                                                Call CreateErrLog("Spare Purchase", .Fields("Invoice #"), "Account_Code - " & .Fields("Vendor Name") & " Not Found In Automan")
                                            Else
                                                mNetAmount = VNull(!Total_Invoice_Amount)
                                                mVatAmt = eVal(!Total_Tax_Amount)
                                                mVat12 = eVal(.Fields("Input VAT @ 12#5%"))
                                                mVat4 = eVal(.Fields("Input VAT @ 4%"))
                                                mPurchaseAmt = eVal(!Net_Amount)
                                                mPurchaseAmt12 = eVal(.Fields("VAT Assessible Amt 1"))
                                                mPurchaseAmt4 = eVal(.Fields("VAT Assessible Amt 4"))
                                            
                                                If eVal(!CST) > 0 Then
                                                    mLocalCentral = "Central"
                                                    mPurchaseAmt12 = mPurchaseAmt + eVal(!Total_Tax_Amount)
                                                Else
                                                    mLocalCentral = "Local"
                                                End If
                                                
                                                mNarr = "Spare Purchase Against Invoice No " & XNull(.Fields("Invoice #"))
                                                
                                                
                                                If Format(mNetAmount, "0.0") = Format(mVat12 + mVat4 + mPurchaseAmt + eVal(!CST), "0.0") Then
                                                    mNetAmount = Round(mVat12 + mVat4 + mPurchaseAmt12 + mPurchaseAmt4, 2)
                                                    If SparePurchase(mSubCode, mNetAmount, mPurchaseAmt12, mPurchaseAmt4, mVat12, mVat4, mNarr, !Invoice_Date, mLocalCentral, .Fields("Invoice #"), !division) = False Then
                                                        Call CreateErrLog("Spare Purchase", .Fields("Invoice #"), "Error In Ledger Posting")
                                                    End If
                                                Else
                                                    Call CreateErrLog("Spare Purchase", .Fields("Invoice #"), "Total Amount : " & mNetAmount & ", Not Match With Purchase Amt 12.5 % : " & mPurchaseAmt12 & " Purchase Amt 4 % : " & mPurchaseAmt4 & " + Tax Amt  : " & mVatAmt)
                                                End If
                                            End If
                                        
                                        If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                        .MoveNext
                                        GCn.CommitTrans
                                        G_FaCn.CommitTrans
                                    Loop
                                End If
                        End If
                    
                    
                    
                    
                    
                    
                    
                    Case ImpSupplierPayment
                        If .RecordCount > 0 Then
                            Prg.Value = 0
                            Prg.Visible = True
                            Do Until .EOF
                                GCn.BeginTrans
                                G_FaCn.BeginTrans
                                    mSubCode = AutomanSubcode(XNull(!Supp_Code), pubSundryCrSysMainGrCode, "Supplier")
                                    mNetAmount = VNull(!Tot_Amt)
                                    mNarr = "Payment Against Payemnt Ref. No " & XNull(!Payment_Ref_No)
                                    
                                    If !Payment_Mode = "C" Then
                                        mCashCredit = "Cash"
                                    Else
                                        mCashCredit = "Credit"
                                    End If
                                    
                                    If SupplierPayment(mSubCode, mNetAmount, mNarr, !Payment_Ref_Date, mCashCredit, XNull(!Cheque_DD_No), XNull(!Cheque_DD_Date), !Payment_Ref_No) = False Then
                                        MsgBox "Posting Failed"
                                        Exit Sub
                                    End If
                                
                                If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                .MoveNext
                                GCn.CommitTrans
                                G_FaCn.CommitTrans
                            Loop
                        End If
                    
                    
                    
                    
                    
                    Case ImpMoneyRectSpare, ImpMoneyRectVehicle
                        If XNull(RsDmsEnviro!SprDebtorGroupCode) = "" Then MsgBox "Plz Define SprDebtorGroupCode In DmsEnviro": Exit Sub
                        mVouCat = "Spare Money Receipt"
                        
                    
                        If ChkFieldExist(RsDms, "Account_Code") And ChkFieldExist(RsDms, "Account_Name") And ChkFieldExist(RsDms, "Full_Name") And ChkFieldExist(RsDms, "Amount") And _
                           ChkFieldExist(RsDms, "Payment_Method") And ChkFieldExist(RsDms, "Chq_DD_RO_No") And _
                           ChkFieldExist(RsDms, "Receipt_Date") And ChkFieldExist(RsDms, "Receipt No") And _
                           ChkFieldExist(RsDms, "Division") And ChkFieldExist(RsDms, "Account_Code") Then
                        
                                If .RecordCount > 0 Then
                                    Prg.Value = 0
                                    Prg.Visible = True
                                    Do Until .EOF
                                        GCn.BeginTrans
                                        G_FaCn.BeginTrans
                                            mInvoiceNo = XNull(.Fields("Receipt No"))
                                            mDmsSubCode = IIf(XNull(!Account_Code) = "", XNull(!Customer_Code), XNull(!Account_Code))
                                            GCn.Execute "Delete From DmsErrLog Where [Key]='" & mInvoiceNo & "'"
                                            
                                            Set RsTemp = GCn.Execute("Select AutomanBankCode From DmsBankAc Where DmsBankCode='" & XNull(!Deposited_On_Bank) & "'")
                                            If RsTemp.RecordCount > 0 Then
                                                mBankAcCode = XNull(RsTemp!AutomanBankCode)
                                            Else
                                                mBankAcCode = ""
                                            End If
                                            
                                            mSubCode = AutomanSubcode(mDmsSubCode, XNull(RsDmsEnviro!SprDebtorGroupCode), "Customer")
                                            If mSubCode = "" Then
                                                Call CreateErrLog(mVouCat, mInvoiceNo, "Account_Code - " & XNull(!Customer_Code) & " Not Found In Automan")
                                            Else
                                                mNetAmount = eVal(!Amount)
                                                If UCase(PubDivCode) = "P" Then
                                                    mNarr = "Payment Received Against Receipt No " & mInvoiceNo & " For Account Name : " & XNull(!First_Name) & " " & XNull(!Middle_Name) & " " & XNull(!Last_Name)
                                                Else
                                                    mNarr = "Payment Received Against Receipt No " & mInvoiceNo & " For Account Name : " & XNull(!Account_Name)
                                                End If
                                                
                                                If MoneyRect(.Fields("Payment_Method"), mSubCode, mNetAmount, mNarr, !Receipt_Date, XNull(!Chq_DD_RO_No), !Instr_Date, mInvoiceNo, !division, mBankAcCode, CmdImport(Index).CAPTION) = False Then
                                                    Call CreateErrLog(mVouCat, mInvoiceNo, "Error In Ledger Posting")
                                                End If
                                            End If
                                        
                                        If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                        .MoveNext
                                        GCn.CommitTrans
                                        G_FaCn.CommitTrans
                                    Loop
                                End If
                        End If
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    Case ImpSprSaleReturn
                        If XNull(RsDmsEnviro!SprDebtorGroupCode) = "" Then MsgBox "Plz Define SprDebtorGroupCode In DmsEnviro": Exit Sub

                        If .RecordCount > 0 Then
                            Prg.Value = 0
                            Prg.Visible = True
                            Do Until .EOF
                                GCn.BeginTrans
                                G_FaCn.BeginTrans
                                    mSubCode = AutomanSubcode(XNull(!Customer_Id), RsDmsEnviro!SprDebtorGroupCode, "Customer")
                                                                                    
                                    mNetAmount = VNull(!Tot_Part_Amt)
                                    mSprSaleAmt = VNull(!Part_Selling_Amt) - VNull(!Part_Level_Discount)
                                    mVatAmt = mNetAmount - mSprSaleAmt
                                    mNarr = "Sale Return Against Invoice No " & XNull(!Invoice_No)
                                    
                                    If SprSaleReturn("Credit", mSubCode, mNetAmount, mSprSaleAmt, mVatAmt, mNarr, !Invoice_Date, !Invoice_No) = False Then
                                        MsgBox "Posting Failed"
                                        Exit Sub
                                    End If
                                
                                If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                .MoveNext
                                GCn.CommitTrans
                                G_FaCn.CommitTrans
                            Loop
                        End If
                        
                        

                        If .RecordCount > 0 Then
                            Prg.Value = 0
                            Prg.Visible = True
                            Do Until .EOF
                                GCn.BeginTrans
                                G_FaCn.BeginTrans
                                    mSubCode = RsDmsEnviro!CashAc
                                                                    
                                    mNetAmount = VNull(!Tot_Part_Amt)
                                    mSprSaleAmt = VNull(!Part_Selling_Amt) - VNull(!Part_Level_Discount)
                                    mVatAmt = mNetAmount - mSprSaleAmt
                                    mNarr = "Sale Return Against Invoice No " & XNull(!Invoice_No)
                                    
                                    If SprSaleReturn("Cash", mSubCode, mNetAmount, mSprSaleAmt, mVatAmt, mNarr, !Invoice_Date, !Invoice_No) = False Then
                                        MsgBox "Posting Failed"
                                        Exit Sub
                                    End If
                                
                                If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                .MoveNext
                                GCn.CommitTrans
                                G_FaCn.CommitTrans
                            Loop
                        End If
                        
                        
        '                Set  = DmsConn.Execute("Select  Convert(VarChar,I.Invoice_Date,3) As Invoice_Date, Sum(ID.RTN_Qty*ID.Unit_Rate)-Sum(ID.Discount_Amt)+Sum(((ID.RTN_Qty*ID.Unit_Rate)-ID.Discount_Amt)*ID.SaleTax_Rate/100) As Tot_Part_Amt, Sum(ID.Rtn_Qty*ID.Unit_Rate)-Sum(ID.Discount_Amt) As Part_Selling_Amt  From HMSI_PTTB_INVOICE_Part_Details  ID Left Join HMSI_PTTB_INVOICE_MASTER I On I.Invoice_No=ID.Invoice_No  Where Payment_Mode='CS' And Rtn_Qty>0 And I.Invoice_Date>='" & Txt(FromDate) & "' And I.Invoice_Date<='" & Txt(ToDate) & "' Group By Convert(VarChar,I.Invoice_Date,3) Order By Convert(VarChar,I.Invoice_Date,3)")
        '                If .RecordCount > 0 Then
        '                    Do Until .EOF
        '                            mSubCode = RsDmsEnviro!CashAc
        '
        '                            mNetAmount = VNull(!Tot_Part_Amt)
        '                            mSprSaleAmt = VNull(!Part_Selling_Amt)
        '                            mVatAmt = mNetAmount - mSprSaleAmt
        '                            mNarr = "Cash Sale Return For Date " & XNull(!Invoice_Date)
        '
        '                            If SprSaleReturn("Cash", mSubCode, mNetAmount, mSprSaleAmt, mVatAmt, mNarr, !Invoice_Date, !Invoice_No) = False Then
        '                                MsgBox "Posting Failed"
        '                                Exit Sub
        '                            End If
        '
        '
        '                        If Round(Prg.value) < 100 Then Prg.value = (.AbsolutePosition / .RecordCount) * 100
        '                        .MoveNext
        '                    Loop
        '                    MsgBox "Posting Completed Successfully"
        '                End If
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    Case ImpVehcleSale
                        If XNull(RsDmsEnviro!VehDebtorGroupCode) = "" Then MsgBox "Plz Define SprDebtorGroupCode In DmsEnviro": Exit Sub
                        
                        If ChkFieldExist(RsDms, "Customer_Code") And ChkFieldExist(RsDms, "Total_Order_Value") And _
                           ChkFieldExist(RsDms, "VatTax") And ChkFieldExist(RsDms, "Chassis_No") And _
                           ChkFieldExist(RsDms, "Invoice_Date") And ChkFieldExist(RsDms, "Division") And _
                           ChkFieldExist(RsDms, "Account_Code") Then
                        
                            mVouCat = "Vehicle Sale"
                                                            
                            .Filter = adFilterNone
                            .Filter = "[Invoice Status]='New'"
                            
                            If .RecordCount > 0 Then
                                Prg.Value = 0
                                Prg.Visible = True
                                Do Until .EOF
                                    GCn.BeginTrans
                                    G_FaCn.BeginTrans
                                        mInvoiceNo = XNull(!Invoice_No)

                                        mDmsSubCode = IIf(XNull(!Account_Code) = "", XNull(!Customer_Code), XNull(!Account_Code))
                                        GCn.Execute "Delete From DmsErrLog Where [Key]='" & mInvoiceNo & "'"
                                        
                                        mSubCode = AutomanSubcode(mDmsSubCode, RsDmsEnviro!SprDebtorGroupCode, "Customer")
                                        If mSubCode = "" Then
                                            Call CreateErrLog(mVouCat, mInvoiceNo, "Account_Code - " & XNull(!Customer_Code) & " Not Found In Automan")
                                        Else
                                            mNetAmount = eVal(!Total_Order_Value)
                                            mVatAmt = eVal(!VATTAX)
                                            mSaleAmt = mNetAmount - mVatAmt
                                            
                                            mNarr = left(" Invoice No " & XNull(!Invoice_No) & XNull(!Narration), 255)
                                            
                                            If VehicleSale(mSubCode, mNetAmount, mSaleAmt, mVatAmt, mNarr, XNull(!Invoice_Date), XNull(!Invoice_No), XNull(!division), XNull(!Chassis_No)) = False Then
                                                Call CreateErrLog(mVouCat, mInvoiceNo, "Error In Ledger Posting")
                                            End If
                                        End If
                                        
                                    If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                    .MoveNext
                                    GCn.CommitTrans
                                    G_FaCn.CommitTrans
                                Loop
                            End If
                        End If
                   
    

     Case ImpUnitImport
             If ChkFieldExist(RsDms, "UOM") Then
     
           Call UnitMasterDataUpdate(ImpUnitImport)

     End If
          
    Case ImpPartImport
    If ChkFieldExist(RsDms, "Part Number") And ChkFieldExist(RsDms, "Description") And ChkFieldExist(RsDms, "UoM") And _
            ChkFieldExist(RsDms, "Vendor") And ChkFieldExist(RsDms, "Vendor Location") And ChkFieldExist(RsDms, "Lead Time") And _
            ChkFieldExist(RsDms, "Discount Code (CVBU)") And _
            ChkFieldExist(RsDms, "Product Category") Then
            
       Call PartMasterDataUpdate(ImpPartImport)
    End If
    
    
    Case ImpModelImport
    
            If ChkFieldExist(RsDms, "Id") And ChkFieldExist(RsDms, "Product Category") And ChkFieldExist(RsDms, "Business Unit") And ChkFieldExist(RsDms, "Catalytic Converter") And ChkFieldExist(RsDms, "Cubic Capacity") And ChkFieldExist(RsDms, "Front Axle Weight") And _
                    ChkFieldExist(RsDms, "Gross Vehicle Weight") And ChkFieldExist(RsDms, "Horse Power") And ChkFieldExist(RsDms, "Number & Description of Type") And ChkFieldExist(RsDms, "Number of Cylinders") And ChkFieldExist(RsDms, "Orderable Flag") And ChkFieldExist(RsDms, "LOB") And _
                    ChkFieldExist(RsDms, "Parent Product Line") And ChkFieldExist(RsDms, "Product Line") And ChkFieldExist(RsDms, "Rear Axle Weight") And ChkFieldExist(RsDms, "Regulatory Certification") And ChkFieldExist(RsDms, "Steering") And _
                    ChkFieldExist(RsDms, "Type of Body") And ChkFieldExist(RsDms, "Unladen Weight") And ChkFieldExist(RsDms, "Product Name") And ChkFieldExist(RsDms, "Product Name1") And ChkFieldExist(RsDms, "UoM") And ChkFieldExist(RsDms, "Product/VC#") And ChkFieldExist(RsDms, "Product Line1") And _
                    ChkFieldExist(RsDms, "Class") And ChkFieldExist(RsDms, "Colour") And ChkFieldExist(RsDms, "Lead Time") And ChkFieldExist(RsDms, "Vehicle Face") And ChkFieldExist(RsDms, "Type") And ChkFieldExist(RsDms, "Product Description") And ChkFieldExist(RsDms, "Wheel base") And ChkFieldExist(RsDms, "Air Ventilation System") And _
                    ChkFieldExist(RsDms, "Fuel Tank") And ChkFieldExist(RsDms, "Model") And ChkFieldExist(RsDms, "RHD/ LHD") And ChkFieldExist(RsDms, "Load Body") And ChkFieldExist(RsDms, "Chassis") And ChkFieldExist(RsDms, "Cab Cowl") And ChkFieldExist(RsDms, "Fuel") And ChkFieldExist(RsDms, "Rear Axle") And _
                    ChkFieldExist(RsDms, "Vehicle Drive") And ChkFieldExist(RsDms, "Seat") Then
             
                   Call ModelMasterUpdate(ImpModelImport)
        
             End If
                    
                    Case ImpVehiclePurchase
                        If XNull(RsDmsEnviro!VehCreditorGroupCode) = "" Then MsgBox "Plz Define Vehicle Creditor Group In DmsEnviro": Exit Sub
                        If XNull(RsDmsEnviro!VatInputAc) = "" Then MsgBox "Plz Define Vat A/c Purchase In DmsEnviro": Exit Sub
                        If XNull(RsDmsEnviro!Vat4InputAc) = "" Then MsgBox "Plz Define Vat A/c Purchase In DmsEnviro": Exit Sub
                        
                        
                        If ChkFieldExist(RsDms, "Supplier_Name") And ChkFieldExist(RsDms, "Value") And _
                           ChkFieldExist(RsDms, "VatTax") And ChkFieldExist(RsDms, "Chassis_No") And _
                           ChkFieldExist(RsDms, "Invoice_Date") And ChkFieldExist(RsDms, "Division") And _
                           ChkFieldExist(RsDms, "Taxable Amount") And ChkFieldExist(RsDms, "Narration") And ChkFieldExist(RsDms, "TAX CST") Then
                           
                        
                            mVouCat = "Vehicle Purchase"
                            If .RecordCount > 0 Then
                                Prg.Value = 0
                                Prg.Visible = True
                                Do Until .EOF
                                    GCn.BeginTrans
                                    G_FaCn.BeginTrans
                                        mInvoiceNo = XNull(!Invoice_No)
                                        GCn.Execute "Delete From DmsErrLog Where [Key]='" & mInvoiceNo & "'"
                                        mSubCode = AutomanSubcode(XNull(!Supplier_Name), RsDmsEnviro!SprDebtorGroupCode, "Supplier")
                                        If mSubCode = "" Then
                                            Call CreateErrLog(mVouCat, mInvoiceNo, "Account_Code - " & XNull(!Supplier_Name) & " Not Found In Automan")
                                        Else
                                            mNetAmount = eVal(!Value)
                                            mVatAmt = eVal(.Fields("VatTax"))
                                            mCstAmt = eVal(.Fields("TAX CST"))
                                            mPurchaseAmt = eVal(.Fields("Taxable Amount")) + eVal(.Fields("Delivery Charges"))
                                            mNarr = left(" Invoice No " & XNull(!Invoice_No) & " " & XNull(!Narration), 255)
                                            
                                            If Format(mNetAmount, "0.0") = Format(mVatAmt + mPurchaseAmt + mCstAmt, "0.0") Then
                                                mNetAmount = Round(mVatAmt + mPurchaseAmt, 2)
                                                If VehiclePurchase(mSubCode, mNetAmount, mPurchaseAmt, mVatAmt, mCstAmt, mNarr, XNull(!Invoice_Date), XNull(!Invoice_No), XNull(!division), XNull(!Chassis_No)) = False Then
                                                    Call CreateErrLog(mVouCat, mInvoiceNo, "Error In Ledger Posting")
                                                End If
                                            Else
                                                Call CreateErrLog(mVouCat, mInvoiceNo, "Total Amount : " & mNetAmount & ", Not Match With Purchase Amt : " & mPurchaseAmt & " + Tax Amt : " & mVatAmt & " + Tax Cst : " & mCstAmt)
                                            End If
                                        End If
                                        
                                    If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
                                    .MoveNext
                                    GCn.CommitTrans
                                    G_FaCn.CommitTrans
                                Loop
                            End If
                        End If
                End Select
                
    End With
        
    MsgBox "Import Process Completed"
        
    'If ChkAllErr.Value = 0 Then mCondStr = " Where U_EntDt = " & ConvertDate(PubLoginDate) & ""
    Set RsTemp = GCn.Execute("Select Cat As Category, [Key] as Dms_Reference, Narration From DmsErrLog " & mCondStr)
    Set FgridErr.DataSource = RsTemp
    Ini_Grid FgridErr
        
    Set RsTemp = Nothing
    Set RsDms = Nothing
    If DmsConn.State <> 0 Then DmsConn.Close
Exit Sub
DispErr:
    MsgBox err.Description
    G_FaCn.RollbackTrans
    GCn.RollbackTrans
    Set RsDms = Nothing
    Set RsTemp = Nothing
    If DmsConn.State <> 0 Then DmsConn.Close
End Sub



Public Sub VehicleInventory(RsDms As Recordset)
'    With RsDms
'        If XNull(RsDmsEnviro!VehCreditorGroupCode) = "" Then MsgBox "Plz Define Vehicle Creditor Group In DmsEnviro": Exit Sub
'        If XNull(RsDmsEnviro!VatInputAc) = "" Then MsgBox "Plz Define Vat A/c Purchase In DmsEnviro": Exit Sub
'        If XNull(RsDmsEnviro!Vat4InputAc) = "" Then MsgBox "Plz Define Vat A/c Purchase In DmsEnviro": Exit Sub
'
'
'        If ChkFieldExist(RsDms, "Supplier_Name") And ChkFieldExist(RsDms, "Value") And _
'           ChkFieldExist(RsDms, "VatTax") And ChkFieldExist(RsDms, "Chassis_No") And _
'           ChkFieldExist(RsDms, "Invoice_Date") And ChkFieldExist(RsDms, "Division") And _
'           ChkFieldExist(RsDms, "Taxable Amount") And ChkFieldExist(RsDms, "Narration") And _
'           ChkFieldExist(RsDms, "TAX CST") And ChkFieldExist(RsDms, "VC_Number") Then
'
'
'
'            mVouCat = "Vehicle Purchase"
'            If .RecordCount > 0 Then
'                Prg.Value = 0
'                Prg.Visible = True
'                Do Until .EOF
'                    GCn.BeginTrans
'                    G_FaCn.BeginTrans
'                        mInvoiceNo = XNull(!Invoice_No)
'                        GCn.Execute "Delete From DmsErrLog Where [Key]='" & mInvoiceNo & "'"
'                        mSubCode = AutomanSubcode(XNull(!Supplier_Name), RsDmsEnviro!SprDebtorGroupCode, "Supplier")
'                        If mSubCode = "" Then
'                            Call CreateErrLog(mVouCat, mInvoiceNo, "Account_Code - " & XNull(!Supplier_Name) & " Not Found In Automan")
'                        Else
'                            mNetAmount = eVal(!Value)
'                            mVatAmt = eVal(.Fields("VatTax"))
'                            mCstAmt = eVal(.Fields("TAX CST"))
'                            mPurchaseAmt = eVal(.Fields("Taxable Amount")) + eVal(.Fields("Delivery Charges"))
'                            mNarr = left(" Invoice No " & XNull(!Invoice_No) & " " & XNull(!Narration), 255)
'
'                            If Format(mNetAmount, "0.0") = Format(mVatAmt + mPurchaseAmt + mCstAmt, "0.0") Then
'                                mNetAmount = Round(mVatAmt + mPurchaseAmt, 2)
'                                If VehiclePurchase(mSubCode, mNetAmount, mPurchaseAmt, mVatAmt, mCstAmt, mNarr, XNull(!Invoice_Date), XNull(!Invoice_No), XNull(!division), XNull(!Chassis_No)) = False Then
'                                    Call CreateErrLog(mVouCat, mInvoiceNo, "Error In Ledger Posting")
'                                End If
'                            Else
'                                Call CreateErrLog(mVouCat, mInvoiceNo, "Total Amount : " & mNetAmount & ", Not Match With Purchase Amt : " & mPurchaseAmt & " + Tax Amt : " & mVatAmt & " + Tax Cst : " & mCstAmt)
'                            End If
'                        End If
'
'                    If Round(Prg.Value) < 100 Then Prg.Value = (.AbsolutePosition / .RecordCount) * 100
'                    .MoveNext
'                    GCn.CommitTrans
'                    G_FaCn.CommitTrans
'                Loop
'            End If
'        End If
'    End With

End Sub



Private Sub FImportJobBill()
Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer, mV_TypeSpare As String, mV_TypeLabour As String, mV_Type As String, mV_TypeReq As String, CodeCnt As Variant
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String, mOrderQty As Double, mPhysicalQty As Double
Dim mPrefix As String, mname As String, mLubType As String, mTrnType As String, mDebitAc As String, mFormCode As String
Dim mChallanNo As String, mHeaderParty As String
Dim mQty As Double, mCount As Integer, mAmount As Double, mTaxAmt As Double, mDiscount As Double
Dim mSpareMrpAmt As Double, mOilMrpAmt As Double
Dim mInvoiceNo As String, mChallanID As String
Dim mFileName As String, mLineFileName As String
Dim mFileTitle As String, mLineFileTitle As String
Dim I As Integer
Dim mVouCat As String
Dim Master1 As New ADODB.Recordset
Dim mCashCredit As String
Dim mGodown As String
Dim mQry As String
Dim mSrl As Integer
Dim mTrans As Boolean
Dim mVTypeGR As String
Dim mVNoGr As String
Dim mEditFlag As Boolean
Dim mReqDocID As String
Dim mSpareDocID As String
Dim mLabourDocID As String
Dim mGatePass As Variant
Dim mNarration As String
Dim mChassisNo As String
Dim mRegnNo As String
Dim mModel As String
Dim mMechCode As String
Dim mLubCat As String
Dim mPurpose As String
Dim mTaxPer As Double
Dim mCreditAc As String
Dim mCardNo As String
Dim mServiceType As String
Dim mMechanic As String
Dim mSupervisor As String
Dim mCnt As String
Dim mLabourTax As Double
Dim mLabourTaxPer As Double
Dim mLineDisPer As Double
Dim mLineDisAmt As Double
Dim IsLineDetailFound As Boolean

'On Error GoTo ELoop

    Call SelectFile
    mFileName = CD1.FileName
    mFileTitle = CD1.FileTitle
    If mFileName = "" Then Exit Sub
    Call SelectFile
    mLineFileName = CD1.FileName
    mLineFileTitle = CD1.FileTitle
    If mFileName = "" Then Exit Sub
    If mLineFileName = "" Then Exit Sub
    mFileTitle = mID(mFileTitle, 1, Len(mFileTitle) - 4)
    mLineFileTitle = mID(mLineFileTitle, 1, Len(mLineFileTitle) - 4)
    
    
    mGodown = GCn.Execute("Select SprWorksGodown from Syctrl").Fields(0).Value
    Set DmsConn = New Connection
    DmsConn.CursorLocation = adUseClient
    DmsConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mFileName & ";Extended Properties=Excel 8.0"
    

    Set RsDms = DmsConn.Execute("Select * from [" & mFileTitle & "$] Where Invoice_Status='New' ")

    Set ExcelGcn2 = New Connection
    ExcelGcn2.CursorLocation = adUseClient
    ExcelGcn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mLineFileName & ";Extended Properties=Excel 8.0"


    


    GCn.BeginTrans
    mTrans = True


    If RsDms.RecordCount > 0 Then RsDms.MoveFirst


    mVouCat = "Job Bill"




    mV_Type = "W_JC"
    mV_TypeReq = "W_RGO"
       


    Do Until RsDms.EOF
        IsLineDetailFound = True
        If XNull(RsDms.Fields("Invoice_No")) = "" Then
            CreateErrLog mVouCat, mInvoiceNo, "Invoice No. is blank in excel file"
            ErrorCnt = 1
        Else
            mInvoiceNo = XNull(RsDms.Fields("Invoice_No"))
        End If
    
    
        GCn.Execute ("Delete from DmsErrLog Where [Key] = '" & XNull(RsDms.Fields("Invoice_No")) & "'")
        
    
        If XNull(RsDms.Fields("Job Card No")) = "" Then
            CreateErrLog mVouCat, mInvoiceNo, "Invoice No. is blank in excel file"
            ErrorCnt = 1
        End If
    
    
        If XNull(StringPass(RsDms.Fields("Division"))) = "" Then
            Call CreateErrLog(mVouCat, mInvoiceNo, "Division Name Field is blank in Excel File")
            ErrorCnt = 1
        Else
            If GCn.Execute("select * from DmsSite where DmsDivision='" & StringPass(RsDms.Fields("Division")) & "'").RecordCount > 0 Then
                mRecordSite = GCn.Execute("select AutomanSite from DmsSite where DmsDivision='" & StringPass(RsDms.Fields("Division")) & "'").Fields(0).Value
                mRecordDiv = GCn.Execute("select AutomanDivision from DmsSite where DmsDivision='" & StringPass(RsDms.Fields("Division")) & "'").Fields(0).Value
            Else
                Call CreateErrLog(mVouCat, mInvoiceNo, "Division Name in not defined in Automan")
                ErrorCnt = 1
            End If
        End If
    
    
        If XNull(RsDms.Fields("SR Type")) = "" Then
            CreateErrLog mVouCat, mInvoiceNo, "SR Type (Service Type) is blank in excel file"
            ErrorCnt = 1
        Else
            If GCn.Execute("Select Count(*) FROM Service_Type WHERE Serv_Desc = '" & XNull(RsDms.Fields("SR Type")) & "' ").Fields(0).Value = 0 Then
                mCnt = GCn.Execute("SELECT IsNull(Max(Convert(NUMERIC,Serv_Type)),0)+1 FROM Service_Type WHERE ISNUMERIC (Serv_Type )=1").Fields(0)
                mQry = "INSERT INTO dbo.Service_Type (Serv_Type, Site_Code, Serv_Desc, Serv_Catg, FreeServCode, Serv_SrlNo, Days, Chrg_From, SprTel, SprDlr, SprCust, LabTel, LabDlr, LabCust, Serv_Target, U_Name, U_EntDt, U_AE, Trf_Date, RateEditableYN) " & _
                     "VALUES ('" & mCnt & "', '" & PubSiteCode & "', '" & XNull(RsDms.Fields("SR Type")) & "', 'C', Null, 10, 30, 1, 1, 0, 0, 100, 0, 0, 100, 'Siebel', '" & PubLoginDate & "', 'A', Null, 1) "
                GCn.Execute mQry
                mServiceType = mCnt
            Else
                mServiceType = GCn.Execute("SELECT Serv_Type FROM Service_Type WHERE Serv_Desc = '" & XNull(RsDms.Fields("SR Type")) & "' ").Fields(0).Value
            End If
        End If
    
        If XNull(RsDms.Fields("Order Date")) = "" Then
            CreateErrLog mVouCat, mInvoiceNo, "Order Date (Job Card Date) is blank in excel file"
            ErrorCnt = 1
        End If
    
        If XNull(RsDms.Fields("Invoice_Date")) = "" Then
            CreateErrLog mVouCat, mInvoiceNo, "Invoice Date is blank in excel file"
            ErrorCnt = 1
        End If
    
        If XNull(RsDms.Fields("Narration")) = "" Then
            CreateErrLog mVouCat, mInvoiceNo, "Narration is blank in excel file"
            ErrorCnt = 1
        Else
            mNarration = XNull(RsDms.Fields("Narration"))
        End If
    
            
        mChassisNo = Trim(mID(mNarration, InStr(1, mNarration, "Chassis Number -") + 16, InStr(1, mNarration, "Reg. Number -") - InStr(1, mNarration, "Chassis Number -") - 16))
        mRegnNo = Trim(mID(mNarration, InStr(1, mNarration, "Reg. Number -") + 13, InStr(1, mNarration, "Model -") - InStr(1, mNarration, "Reg. Number -") - 13))
        mModel = Trim(mID(mNarration, InStr(1, mNarration, "Model -") + 7, 50))
    
 
        If mChassisNo = "" Then
            mChassisNo = "NA"
            CreateErrLog mVouCat, mInvoiceNo, "Chassis No in Narration is blank in excel file"
            'ErrorCnt = 1
        End If
    
        If mModel = "" Then
            mModel = "NA"
            CreateErrLog mVouCat, mInvoiceNo, "Model in Narration is blank in excel file"
            'ErrorCnt = 1
        End If
    
    
    
    
    If XNull(RsDms.Fields("Mode Of Payment")) = "" Then
        CreateErrLog mVouCat, mInvoiceNo, "Invoice No  is blank is excel file"
        ErrorCnt = 1
    Else
        If StrCmp(XNull(RsDms.Fields("Mode Of Payment")), "Cash") Then
            mV_TypeSpare = "W_SIC"
            mV_TypeLabour = "W_LIC"
            mCashCredit = "Cash"
        Else
            mV_TypeSpare = "W_SIR"
            mV_TypeLabour = "W_LIR"
            mCashCredit = "Credit"
        End If
    End If


    If mCashCredit = "Credit" Then
        With RsDms
        If GCn.Execute("Select Count(*) From DmsSubGroup Where DmsSubCode='" & XNull(RsDms.Fields("Account_Code")) & "'").Fields(0).Value = 0 Then
                Set RsTemp = GCn.Execute("Select AutomanSite From DmsSite Where DmsDivision='" & XNull(.Fields("Division")) & "'")
                If RsTemp.RecordCount > 0 Then
                    GCn.Execute "Delete From DmsErrLog Where [Key] = '" & XNull(.Fields("Customer_Code")) & "' "
                    GCn.Execute "Insert Into DmsSubGroup(DmsSubCode, Name,[Group], Division) " & _
                                " Values ('" & IIf(XNull(.Fields("Account_Code")) <> "", XNull(.Fields("Account_Code")), XNull(.Fields("Customer_Code"))) & "', " & _
                                "'" & left(IIf(XNull(.Fields("Account_Name")) <> "", XNull(.Fields("Account_Name")), XNull(.Fields("Full Name"))), 50) & "','Sundry Debtors', '" & XNull(.Fields("Division")) & "')"
                Else
                    CreateErrLog mVouCat, XNull(.Fields("Account_Code")), XNull(.Fields("Division")) & " Not Defined In DmsDivision Table"
                End If
        End If
    
        mPartyCode = AutomanSubcode(XNull(.Fields("Account_Code")), RsDmsEnviro!SprDebtorGroupCode, "Customer")
        End With
    End If
    
    mPrefix = "     "
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & " + 1 from SP_Stock where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "1") & "='" & mRecordSite & "' and V_Type='" & mV_TypeReq & "'").Fields(0).Value
    If mEditFlag Then
        If GCn.Execute("Select count(*) From Sp_Stock where Job_DocID='" & mDocId & "' and v_Type='" & mV_Type & "'").Fields(0) > 0 Then
            mReqDocID = GCn.Execute("Select DocId From Sp_Stock where Job_DocID='" & mDocId & "' and v_Type='" & mV_Type & "'").Fields(0)
        Else
            mReqDocID = mRecordDiv & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
        End If
        
        If GCn.Execute("Select GP_No From Job_Card where DocID='" & mDocId & "'").Fields(0) > 0 Then
            mGatePass = GCn.Execute("Select GP_No From Job_Card where DocID='" & mDocId & "'").Fields(0)
        Else
            mGatePass = "00000" & GCn.Execute("select " & vIsNull("max(" & cVal("right(gp_no,5)") & ")", "0") & "+1 from job_card where left(gp_no,1)='" & mRecordDiv & "' AND " & cMID("gp_no", "2", "1") & "='" & left(mRecordSite, 1) & "'").Fields(0).Value
            mGatePass = mRecordDiv & mRecordSite & Right(mGatePass, 5)
        End If
    Else
        mReqDocID = mRecordDiv & mRecordSite & mRecordSite & mV_TypeReq & mPrefix & Right("00000000" & CodeCnt, 8)
        mGatePass = "00000" & GCn.Execute("select " & vIsNull("max(" & cVal("right(gp_no,5)") & ")", "0") & "+1 from job_card where left(gp_no,1)='" & mRecordDiv & "' AND " & cMID("gp_no", "2", "1") & "='" & left(mRecordSite, 1) & "'").Fields(0).Value
        mGatePass = mRecordDiv & mRecordSite & Right(mGatePass, 5)
    End If
    
    If XNull(RsDms.Fields("Job Card No")) <> "" Then
        mDocId = mRecordDiv & mRecordSite & mRecordSite & mV_Type & " " & mPrefix & Right("00000000" & Val(Right(RsDms.Fields("Job Card No"), 6)), 8)
        mSpareDocID = mRecordDiv & mRecordSite & mRecordSite & mV_TypeSpare & mPrefix & Right("00000000" & Val(Right(RsDms.Fields("Invoice_No"), 5)), 8)
        mLabourDocID = mRecordDiv & mRecordSite & mRecordSite & mV_TypeLabour & mPrefix & Right("00000000" & Val(Right(RsDms.Fields("Invoice_No"), 5)), 8)

    End If
    
    
    If GCn.Execute("Select Count(*) From Job_Card With (NoLock) Where DocID = '" & mDocId & "'").Fields(0).Value > 0 Then
        CreateErrLog mVouCat, mInvoiceNo, "Job No already exist in automan"
        ErrorCnt = 1
    End If
    
    If GCn.Execute("Select Count(*) From Job_Card With (NoLock) Where DocID_InvSpr = '" & mSpareDocID & "'").Fields(0).Value > 0 Then
        CreateErrLog mVouCat, mInvoiceNo, "Invoice No already exist in automan"
        ErrorCnt = 1
    End If
    

    Set Master1 = CreateObject("ADODB.Recordset")
    GSQL = "Select * FROM [" & mLineFileTitle & "$] where Invoice_No='" & mInvoiceNo & "' Order By  [Invoice_No]"
    Master1.Open GSQL, ExcelGcn2, adOpenStatic


    If Master1.RecordCount = 0 Then
        CreateErrLog mVouCat, mInvoiceNo, " Line detail not found for Invoice No : " & mInvoiceNo
        IsLineDetailFound = False
    End If


    mTaxAmt = 0: mAmount = 0: mDiscount = 0: mSpareMrpAmt = 0: mOilMrpAmt = 0
    If Master1.RecordCount > 0 Then
        Master1.MoveFirst
        For I = 0 To Master1.RecordCount - 1
            Set RsTemp = GCn.Execute("Select Part_No, Part_Grade, MRP From Part Where Part_No='" & Master1.Fields("Part #") & "'")
            If RsTemp.RecordCount = 0 Then
                mQry = "Insert Into Part (Part_No, Div_Code, Site_Code, Part_Name, Part_Grade, TB_SRate ) "
                mQry = mQry + " Values('" & Master1.Fields("Part #") & "', '" & mRecordDiv & "', '" & mRecordSite & "', '" & Master1.Fields("Part #") & "','S', " & VNull(Master1.Fields("Rate")) & ")"
                GCn.Execute mQry
                CreateErrLog mVouCat, mInvoiceNo, " Part # :  " & Master1.Fields("Part #") & " is added in automan"
            End If
        
            If UTrim(XNull(Master1.Fields("Job Card_No"))) = "" Then
                CreateErrLog mVouCat, mInvoiceNo, " Job Card_No Field is blank in Line Detail "
                ErrorCnt = 1
            End If
        
            If UTrim(XNull(Master1.Fields("Part #"))) = "" Then
                CreateErrLog mVouCat, mInvoiceNo, " Part No. Field is blank in Line Detail "
                ErrorCnt = 1
            End If
        
       
            mTaxAmt = mTaxAmt + Val(Master1.Fields("Tax Amount"))
            mAmount = mAmount + Val(Master1.Fields("Net_Amount")) - Val(Master1.Fields("Tax Amount"))
            mDiscount = mDiscount + Val(Master1.Fields("Discount"))
            
            If XNull(Master1.Fields("Product Category")) = "Lubricant" Then
                mOilMrpAmt = mOilMrpAmt + VNull(Master1.Fields("Net_Amount")) - VNull(Master1.Fields("Tax Amount"))
            Else
                mSpareMrpAmt = mSpareMrpAmt + VNull(Master1.Fields("Net_Amount")) - VNull(Master1.Fields("Tax Amount"))
            End If
            
            Master1.MoveNext
        Next
    Else
        mSpareMrpAmt = eVal(RsDms.Fields("Parts Amount"))
        mOilMrpAmt = eVal(RsDms.Fields("Lubricant Amount"))
    End If
    
    If Round(eVal(RsDms.Fields("Total Parts Amount"))) <> 0 And IsLineDetailFound = False Then
'        MsgBox ""
    End If
    If Round(eVal(RsDms.Fields("Total Parts Amount"))) <> Round(mAmount) And IsLineDetailFound Then
        CreateErrLog mVouCat, mInvoiceNo, " Header Total Parts Amount : " & eVal(RsDms.Fields("Total Parts Amount")) & " does not match with Line Amount : " & mAmount & " "
        IsLineDetailFound = False 'ErrorCnt = 1
    End If
    
    If Round(eVal(RsDms.Fields("VAT"))) <> Round(mTaxAmt) And IsLineDetailFound Then
        CreateErrLog mVouCat, mInvoiceNo, " Header VAT Amount : " & eVal(RsDms.Fields("VAT")) & " does not match with Line Tax Amount : " & mTaxAmt & " "
        IsLineDetailFound = False 'ErrorCnt = 1
    End If
    
'    If eVal(RsDms.Fields("Discount Job Parts")) <> mDiscount Then
'        CreateErrLog mVouCat, mInvoiceNo, " Header Discount Job Parts : " & eVal(RsDms.Fields("Discount Job Parts")) & " does not match with Line Discount : " & mDiscount & " "
'        ErrorCnt = 1
'    End If
    
    
    If ErrorCnt = 0 Then
        With RsDms
            mCardNo = GCn.Execute("Select IsNull(Max(CardNo),'') from Hiscard where Chassis = '" & mChassisNo & "'").Fields(0).Value
            If mCardNo = "" Then
                mCardNo = GCn.Execute("select " & vIsNull("max(" & cVal(cMID("CardNo", "2", "len(cardno)-1")) & ")", "0") & "+1 from Hiscard where Site_Code='" & mRecordSite & "'  AND IsNumeric(SubString(CardNo,2, len(cardno)-1))=1 ").Fields(0).Value
                mCardNo = mRecordSite & Right("0000000" & mCardNo, 7)
                
                mQry = "Insert Into HisCard( "
                mQry = mQry + "CardNo, Site_Code, Div_Code, Carddate, Model, "
                mQry = mQry + "RegNo, Chassis, Engine, Delivery_Date, Dealer_Code, "
                mQry = mQry + "CouponNo, Supplier_BillNo, Supplier_BillDate, Name, Add1, "
                mQry = mQry + "Add2, Add3, PhoneOff, PhoneResi, Mobile, "
                mQry = mQry + "Govt_Yn, U_Name, U_EntDt, U_AE "
                mQry = mQry + ") "
                mQry = mQry + "Values ( "
                mQry = mQry + "'" & mCardNo & "', '" & mRecordSite & "', '" & mRecordDiv & "', '" & MakeDate(.Fields("Order Date")) & "', '" & mModel & "', "
                mQry = mQry + "'" & mRegnNo & "', '" & mChassisNo & "', '', Null, '', "
                mQry = mQry + "'', '', Null, '" & XNull(.Fields("Full Name")) & "', '', "
                mQry = mQry + "'', '', '', '', '', "
                mQry = mQry + "0, 'Siebel', '" & PubLoginDate & "', 'A' "
                mQry = mQry + ") "
                GCn.Execute mQry
            End If
        
            
            Dim mDisPer As Double
            mDisPer = 0
            If (mSpareMrpAmt + mOilMrpAmt) > 0 And eVal(.Fields("Discount Job Parts")) > 0 Then
                mDisPer = Format(eVal(.Fields("Discount Job Parts")) * 100 / (mSpareMrpAmt + mOilMrpAmt), "0.0000")
            End If
            mQry = "Insert Into Sp_Sale ( "
            mQry = mQry + "DocID, DocIDHelp, Site_Code, V_Type, V_No, "
            mQry = mQry + "V_Date, Party_Code, Cash_Credit, Party_Name, "
            mQry = mQry + "L_C, Form_Code, CrAc, SiebelDocID, Job_DocID, "
            mQry = mQry + "PType, GP_No, GP_Date, SprAmt_MRP_TB, SprAmt_MRP_TP, "
            mQry = mQry + "OilAmt_Mrp_TB, OilAmt_Mrp_TP, D_Per_Mrp_TB, D_Per_Mrp_TP, D_Amt_Mrp_TB, "
            mQry = mQry + "D_Amt_Mrp_TP, SprAmt_TB, SprAmt_TP, OilAmt_TB, OilAmt_TP, "
            mQry = mQry + "D_Per_TB, D_Per_TP, D_Amt_TB, D_Amt_TP, Addition, "
            mQry = mQry + "Tax_Amt, Packing, Tot_Per, Tot_Amt, ReSalTax_Per, "
            mQry = mQry + "ReSalTax_Amt, Total_Amt, Rounded, Det_Tax, AcPosting_Yn, "
            mQry = mQry + "U_Name, U_EntDt, U_AE "
            mQry = mQry + ")"
            mQry = mQry + "Values "
            mQry = mQry + "("
            mQry = mQry + "'" & mSpareDocID & "', '" & mSpareDocID & "', '" & mRecordSite & "', '" & DeCodeDocID(mSpareDocID, Document_Type) & "', '" & DeCodeDocID(mSpareDocID, Document_No) & "', "
            mQry = mQry + "'" & MakeDate(.Fields("Invoice_Date")) & "', '" & mPartyCode & "', '" & .Fields("Mode Of Payment") & "', '" & IIf(Len(XNull(.Fields("Full Name"))) > 40, left(Replace(XNull(.Fields("Full Name")), ".", ""), 40), XNull(.Fields("Full Name"))) & "', "
            mQry = mQry + "'L', '" & mFormCode & "', '" & mCreditAc & "', '" & .Fields("Invoice_No") & "', '" & mDocId & "', "
            mQry = mQry + "'General', '" & mGatePass & "', " & ConvertDate(MakeDate(.Fields("Invoice_Date"))) & ", " & mSpareMrpAmt & ", 0, "
            mQry = mQry + "" & mOilMrpAmt & ", 0," & mDisPer & ", 0, " & eVal(.Fields("Discount Job Parts")) & ", "
            mQry = mQry + "0, 0, 0, 0, 0, "
            mQry = mQry + " " & mDisPer & ", 0, " & eVal(.Fields("Discount Job Parts")) & ", 0, 0, "
            mQry = mQry + "" & eVal(.Fields("VAT")) & ", " & eVal(.Fields("Other Charges")) & ", 0, 0, 0, "
            mQry = mQry + "0, " & eVal(.Fields("Spares Invoice Amount")) & ", 0, 0, 1, "
            mQry = mQry + "'Siebel', '" & PubLoginDate & "', 'A' "
            mQry = mQry + ")"
            GCn.Execute mQry
            
            
            If IsLineDetailFound = False And (mSpareMrpAmt <> 0 Or mOilMrpAmt <> 0) Then
                mPurpose = "C"
            
                'Dim mLineDisAmt As Double
                'Dim mTaxAmt As Double
            
                'If eVal(.Fields("Discount Job Parts")) <> 0 Then
                '    mLineDisPer = eVal(.Fields("Discount Job Parts")) * 100 / mSpareMrpAmt + mOilMrpAmt
                'Else
                    mLineDisPer = 0
                'End If
                'mLineDisAmt = eVal(.Fields("Discount Job Parts"))
                mLineDisAmt = 0
                mTaxPer = Format(eVal(.Fields("VAT")) * 100 / (mSpareMrpAmt + mOilMrpAmt), "0.00")
                mTaxAmt = eVal(.Fields("VAT"))
                
                If mSpareMrpAmt <> 0 Then
                    GCn.Execute "Insert Into Sp_Stock (DocId, Srl_No, V_No, Site_Code, V_Type, V_Date, Job_DocId, " & _
                                "Job_DivCode, Mech_Code, Part_No, Lub_Category, Godown, " & _
                                "Qty_Doc, Qty_Iss, Tax_Yn, Mrp_Yn, Rate, Mrp_Rate, " & _
                                "Disc_Per, Disc_Amt, Amount, Net_Amt, Purpose, V_Rate, " & _
                                "TaxPer, TaxAmt, SiebelDocId, U_EntDt, U_Name, U_AE) Values ( " & _
                                "'" & mReqDocID & "',  1 , " & DeCodeDocID(mReqDocID, Document_No) & ", '" & mRecordSite & mRecordSite & "', '" & DeCodeDocID(mReqDocID, Document_Type) & "', " & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", '" & mDocId & "', " & _
                                "'" & mRecordDiv & "', '" & mMechCode & "', '" & RsDmsEnviro.Fields("DefaultPartNo") & "', '" & mLubCat & "', '" & PubSprWorksGodown & "', " & _
                                "1,1, 1, 1, " & mSpareMrpAmt + mTaxAmt & ", " & mSpareMrpAmt + mTaxAmt & ", " & _
                                "" & mLineDisPer & ", " & mLineDisAmt & ", " & mSpareMrpAmt + mTaxAmt & ", " & mSpareMrpAmt & ",'" & mPurpose & "', " & mSpareMrpAmt & ", " & _
                                "" & mTaxPer & ", " & mTaxAmt & ", '" & mDocId & "', " & ConvertDate(PubLoginDate) & ", 'Siebel', 'A')"
                    mLineDisPer = 0
                    mLineDisAmt = 0
                    mTaxPer = 0
                    mTaxAmt = 0
                End If
                
                If mOilMrpAmt <> 0 Then
                    GCn.Execute "Insert Into Sp_Stock (DocId, Srl_No, V_No, Site_Code, V_Type, V_Date, Job_DocId, " & _
                                "Job_DivCode, Mech_Code, Part_No, Lub_Category, Godown, " & _
                                "Qty_Doc, Qty_Iss, Tax_Yn, Mrp_Yn, Rate, Mrp_Rate, " & _
                                "Disc_Per, Disc_Amt, Amount, Net_Amt, Purpose, V_Rate, " & _
                                "TaxPer, TaxAmt, SiebelDocId, U_EntDt, U_Name, U_AE) Values ( " & _
                                "'" & mReqDocID & "', 2, " & DeCodeDocID(mReqDocID, Document_No) & ", '" & mRecordSite & mRecordSite & "', '" & DeCodeDocID(mReqDocID, Document_Type) & "', " & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", '" & mDocId & "', " & _
                                "'" & mRecordDiv & "', '" & mMechCode & "', '" & RsDmsEnviro.Fields("DefaultOilPartNo") & "', '" & mLubCat & "', '" & PubSprWorksGodown & "', " & _
                                "1,1, 1, 1, " & mOilMrpAmt + mTaxAmt & ", " & mOilMrpAmt + mTaxAmt & ", " & _
                                "" & mLineDisPer & ", " & mLineDisAmt & ", " & mOilMrpAmt + mTaxAmt & ", " & mOilMrpAmt & ",'" & mPurpose & "', " & mOilMrpAmt & ", " & _
                                "" & mTaxPer & ", " & mTaxAmt & ", '" & mDocId & "', " & ConvertDate(PubLoginDate) & ", 'Siebel', 'A')"
                End If
            End If
            
            
            
            mQry = "Insert Into Job_Card("
            mQry = mQry + "DocID, Site_Code, Job_No, Job_Date, Job_BookDivCode, "
            mQry = mQry + "Job_BookNo, Job_BookSiteCode, CardNo, Govt_Yn, Serv_Type, "
            mQry = mQry + "AtKmsHrs, Fuel, Est_SpCost, Est_LabCost, ArrivalTime, "
            mQry = mQry + "Recp_Time, ExpDelDate, Body_Damage, OpenRemarks, KmsHrs, "
            mQry = mQry + "JobType, RecBy_Mechanic, RecBy_Supervisor, TempCloseDate, SiebelDocID, "
            mQry = mQry + "CreatedU_Name, CreatedU_EntDt, CreatedU_AE, U_Name, U_EntDt, U_AE "
            mQry = mQry + ")"
            mQry = mQry + "Values("
            mQry = mQry + "'" & mDocId & "', '" & mRecordSite & "', '" & DeCodeDocID(mDocId, Document_No) & "', '" & MakeDate(.Fields("Order Date")) & "', '', "
            mQry = mQry + "'', '', '" & mCardNo & "', 0, '" & mServiceType & "', "
            mQry = mQry + "0, 0, 0, 0, '" & MakeDate(.Fields("Order Date")) & "', "
            mQry = mQry + "'" & MakeDate(.Fields("Order Date")) & "', '" & MakeDate(.Fields("Invoice_Date")) & "', '', '', 'K', "
            mQry = mQry + "'', '" & XNull(RsDmsEnviro.Fields("DefaultMechanic")) & "', '" & XNull(RsDmsEnviro.Fields("DefaultSupervisor")) & "', '" & MakeDate(.Fields("Invoice_Date")) & "', '" & .Fields("Job Card No") & "', "
            mQry = mQry + "'Siebel', '" & PubLoginDate & "', 'A', 'Siebel', '" & PubLoginDate & "', 'A' "
            mQry = mQry + ")"
            GCn.Execute mQry
            
            
            mLabourTax = eVal(.Fields("Service Tax")) + eVal(.Fields("Cess Tax")) + eVal(.Fields("S n HE"))
            If mLabourTax > 0 Then
                mLabourTaxPer = Format(mLabourTax * 100 / eVal(.Fields("Labour Invoice Amount")), "0.00")
            Else
                mLabourTaxPer = 0
            End If
            
            mQry = "Update Job_Card set JobCloseDate=" & cIIF("TempCloseDate Is Null", ConvertDate(MakeDate(.Fields("Invoice_Date"))), "TempCloseDate") & ",JobComp_Dt_Time=" & cIIF("TempCloseDate Is Null", ConvertDate(MakeDate(.Fields("Invoice_Date"))), "TempCloseDate") & ""
            mQry = mQry + ",CrMemo=" & IIf(.Fields("Mode of Payment") = "CREDIT", 1, 0) & ",BillingName='" & IIf(Len(XNull(.Fields("Full Name"))) > 40, left(Replace(XNull(.Fields("Full Name")), ".", ""), 40), XNull(.Fields("Full Name"))) & "',DelBy=RecBy_Mechanic" & ""
            mQry = mQry + ",DrSpr_AcCode='" & mPartyCode & "',DrLab_AcCode='" & mPartyCode & "',DocId_InvSpr='" & mSpareDocID & "',DocId_InvLab='" & mLabourDocID & "',GP_NO='" & mGatePass & ""
            mQry = mQry + "',LabAmt_TB=" & eVal(.Fields("Labour Invoice Amount")) & ",LabAmt_TP=0,Lab_D_Amt= " & eVal(.Fields("Discount Labour")) & ",LabD_Per= 0,Lab_TaxPer=" & mLabourTaxPer & ",Lab_TaxAmt= " & mLabourTax & ""
            mQry = mQry + ",Lab_RoundOff= 0,NetLab_Amt= " & eVal(.Fields("Total Labour Amount")) & ""
            mQry = mQry + ",ClosedU_Name='Siebel',ClosedU_EntDt=" & ConvertDate(PubLoginDate) & ",ClosedU_AE='A' where Job_Card.DocId='" & mDocId & "'"
            GCn.Execute mQry
            
            
            mQry = "INSERT INTO dbo.Job_Lab (Job_DocID, S_No, Site_Code, Lab_Code, Tax_YN, LabourAmt, Chrg_From, External_YN, U_Name, U_EntDt, U_AE, Chrg_Type) "
            mQry = mQry + "VALUES ('" & mDocId & "', 1, '" & mRecordSite & "', '" & XNull(RsDmsEnviro.Fields("DefaultLabourHead")) & "', " & IIf(mLabourTax > 0, 1, 0) & ", " & eVal(.Fields("Labour Invoice Amount")) & ", 'C', '0', 'Siebel', '" & PubLoginDate & "', 'A','C') "
            GCn.Execute mQry


            mQry = "INSERT INTO dbo.Job_Lab2 (Job_DocID, S_No, Mech_Code, Lab_Code, Site_Code, U_Name, U_EntDt, U_AE) "
            mQry = mQry + "VALUES ('" & mDocId & "', 1, '" & XNull(RsDmsEnviro.Fields("DefaultMechanic")) & "', '" & XNull(RsDmsEnviro.Fields("DefaultLabourHead")) & "', '" & mRecordSite & "', 'Siebel', '" & PubLoginDate & "', 'A')"
            GCn.Execute mQry
            
        End With
    
                    
        If IsLineDetailFound = True Then
            With Master1
                If Master1.RecordCount > 0 Then
                    Master1.MoveFirst
                    For I = 0 To Master1.RecordCount - 1
                        mPurpose = "C"
                        
                        If Val(.Fields("Tax Amount")) > 0 Then
                        
    '                        If Round(Val(.Fields("Net_Amount")) + Val(.Fields("Discount")) - Val(.Fields("Tax Amount"))) = Round(Val(.Fields("NTA")) * Val(.Fields("Qty"))) Then
                                mTaxPer = Format(Val(.Fields("Tax Amount")) * 100 / (Val(.Fields("Net_Amount")) + Val(.Fields("Discount")) - Val(.Fields("Tax Amount"))), "0.00")
                                mLineDisAmt = 0
                                mLineDisPer = 0
    '                        Else
    '                            mTaxPer = Format(Val(.Fields("Tax Amount")) * 100 / (Val(.Fields("NTA")) * Val(.Fields("Qty"))), "0.00")
    '                            mLineDisAmt = Val(.Fields("Discount"))
    '                            mLineDisPer = Val(.Fields("Discount")) * 100 / (Val(.Fields("Net_Amount")) + Val(.Fields("Discount")))
    '                        End If
                        Else
                            mTaxPer = 0
                        End If
                        
                        GCn.Execute "Insert Into Sp_Stock (DocId, Srl_No, V_No, Site_Code, V_Type, V_Date, Job_DocId, " & _
                                    "Job_DivCode, Mech_Code, Part_No, Lub_Category, Godown, " & _
                                    "Qty_Doc, Qty_Iss, Tax_Yn, Mrp_Yn, Rate, Mrp_Rate, " & _
                                    "Disc_Per, Disc_Amt, Amount, Net_Amt, Purpose, V_Rate, " & _
                                    "TaxPer, TaxAmt, SiebelDocId, U_EntDt, U_Name, U_AE) Values ( " & _
                                    "'" & mReqDocID & "', " & I & ", " & DeCodeDocID(mReqDocID, Document_No) & ", '" & mRecordSite & mRecordSite & "', '" & DeCodeDocID(mReqDocID, Document_Type) & "', " & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", '" & mDocId & "', " & _
                                    "'" & mRecordDiv & "', '" & mMechCode & "', '" & .Fields("Part #") & "', '" & mLubCat & "', '" & PubSprWorksGodown & "', " & _
                                    "" & Val(.Fields("Qty")) & "," & Val(.Fields("Qty")) & ", 1, 1, " & Val(.Fields("Rate")) & ", " & Val(.Fields("Rate")) & ", " & _
                                    "" & mLineDisPer & ", " & mLineDisAmt & ", " & Val(.Fields("Net_Amount")) + Val(.Fields("Discount")) & ", " & Val(.Fields("Net_Amount")) - Val(.Fields("Tax Amount")) & ",'" & mPurpose & "', " & Val(.Fields("Rate")) & ", " & _
                                    "" & mTaxPer & ", " & Val(.Fields("Tax Amount")) & ", '" & mDocId & "', " & ConvertDate(PubLoginDate) & ", 'Siebel', 'A')"
                        Master1.MoveNext
                    Next
                End If
            End With
        End If
    End If



        CodeCnt = CodeCnt + 1
        
        RsDms.MoveNext
        ErrorCnt = 0
    Loop
    GCn.CommitTrans
    mTrans = False
    MsgBox "Spare Detail Imported Successfully"

lblExit:
    Set RsNew = Nothing
    Exit Sub
ELoop:
    MsgBox err.Description
    If mTrans Then GCn.RollbackTrans
End Sub





Private Sub CmdJobBill_Click()
    FImportJobBill
End Sub

Private Sub CmdOk_Click()
Dim I As Integer
Dim Cnt As Integer
    For I = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, FSel) <> "" Then
            Cnt = Cnt + 1
        End If
    Next I
    If Cnt = 0 Then
        MsgBox "No Account Name Is Selected"
    Else
        Frame3.Visible = False
    End If
End Sub

Private Sub Cmdsave_Click()
Dim I As Integer
    If GCn.Execute("Select * from DmsEnviro").RecordCount > 0 Then
        GCn.Execute "Update DmsEnviro Set SprDebtorGroupCode='" & Txt(SprDebtorGroupCode).Tag & "', VehDebtorGroupCode = '" & Txt(VehDebtorGroupCode).Tag & "', " & _
                            "WsDebtorGroupCode='" & Txt(WsDebtorGroupCode).Tag & "', SprCreditorGroupCode='" & Txt(SprCreditorGroupCode).Tag & "', SprSaleAc='" & Txt(SprSaleAc).Tag & "', " & _
                            "LubeSaleAc='" & Txt(LubSaleAc).Tag & "', SprSaleVat4Ac = '" & Txt(SprSaleVat4Ac).Tag & "', VehSaleAc='" & Txt(VehSaleAc).Tag & "', SprPurchaseAc='" & Txt(SprPurchaseAc).Tag & "', " & _
                            "VehPurchaseAc='" & Txt(VehPurchaseAc).Tag & "', SprCashAc='" & Txt(SprCashAc).Tag & "', VehCashAc='" & Txt(VehCashAc).Tag & "', " & _
                            "WsCashAc='" & Txt(WsCashAc).Tag & "', SprBankAc='" & Txt(SprBankAc).Tag & "', VehBankAc='" & Txt(VehBankAc).Tag & "', WsBankAc='" & Txt(WsBankAc).Tag & "', " & _
                            "LocalStateName='" & Txt(LocalStateName) & "', LabourAc='" & Txt(LabourAc).Tag & "', ServTaxAc='" & Txt(ServTaxAc).Tag & "', " & _
                            "VehCreditorGroupCode='" & Txt(VehCreditorGroupCode).Tag & "', CstAc = '" & Txt(CstAc).Tag & "', VatAc='" & Txt(VatAc).Tag & "', " & _
                            "ROffAc='" & Txt(ROffAc).Tag & "', SprCstPurchaseAc='" & Txt(SprCstPurchaseAc).Tag & "', DiscountAc='" & Txt(DiscountAc).Tag & "', " & _
                            "OtherChargesAc='" & Txt(OtherChargesAc).Tag & "', VehPurGroupCode='" & Txt(VehPurGroupCode).Tag & "', VehSaleGroupCode='" & Txt(VehSaleGroupCode).Tag & "', " & _
                            "SprPurGroupCode='" & Txt(SprPurGroupCode).Tag & "', SprSaleGroupCode = '" & Txt(SprSaleGroupCode).Tag & "', VatGroupCode='" & Txt(VatGroupCode).Tag & "', " & _
                            "ServiceTaxGroupCode='" & Txt(ServiceTaxGroupCode).Tag & "', VehCstPurchaseAc = '" & Txt(VehCstPurchaseAc).Tag & "', VatInputAc = '" & Txt(VatInputAc).Tag & "', Vat4InputAc = '" & Txt(Vat4InputAc).Tag & "', Vat4Ac = '" & Txt(Vat4Ac).Tag & "', SprPurchase4Ac = '" & Txt(SprPurchase4Ac).Tag & "', " & _
                            "VehicleCentralPurchaseTaxForm = '" & Txt(VehicleCentralPurchaseTaxForm).Tag & "', VehicleLocalPurchaseTaxForm = '" & Txt(VehicleLocalPurchaseTaxForm).Tag & "', VehiclePurchaseDiscountItem = '" & Txt(VehiclePurchaseDiscountItem).Tag & "' " & _
                            ", VehiclePurchaseTransportItem = '" & Txt(VehiclePurchaseTransportItem).Tag & "'  , VehicleTaxOnDeliveryCharges = '" & left(Txt(VehicleTaxOnDeliveryCharges), 1) & "', SpareCentralPurchaseTaxForm = '" & Txt(SpareCentralPurchaseTaxForm).Tag & "' , SpareLocalPurchaseTaxForm = '" & Txt(SpareLocalPurchaseTaxForm).Tag & "' " & _
                            ", SpareCentralSaleTaxForm = '" & Txt(SpareCentralSaleTaxForm).Tag & "', SpareLocalSaleTaxForm = '" & Txt(SpareLocalSaleTaxForm).Tag & "', DefaultSupervisor ='" & Txt(DefaultSupervisor).Tag & "', DefaultMechanic = '" & Txt(DefaultMechanic).Tag & "', DefaultLabourHead = '" & Txt(DefaultLabourHead).Tag & "', DefaultPartNo = '" & Txt(DefaultPartNo).Tag & "', DefaultOilPartNo = '" & Txt(DefaultOilPartNo).Tag & "'"
    Else
        GCn.Execute "Insert Into DmsEnviro(SprDebtorGroupCode, VehDebtorGroupCode,WsDebtorGroupCode, SprCreditorGroupCode, VehCreditorGroupCode, " & _
                            "SprSaleAc, SprSaleVat4Ac, LubeSaleAc, VehSaleAc, SprPurchaseAc, VehPurchaseAc, " & _
                            "SprCashAc, VehCashAc, WsCashAc, SprBankAc, VehBankAc, " & _
                            "WsBankAc, LocalStateName, LabourAc, ServTaxAc, CstAc, " & _
                            "VatAc, ROffAc, SprCstPurchaseAc, OtherChargesAc, DiscountAc, " & _
                            "VehPurGroupCode, VehSaleGroupCode, SprPurGroupCode, SprSaleGroupCode, VatGroupCode, ServiceTaxGroupCode, VehCstPurchaseAc, Vat4Ac, SprPurchase4Ac, VatInputAc, Vat4InputAc, " & _
                            "VehicleCentralPurchaseTaxForm, VehicleLocalPurchaseTaxForm, VehiclePurchaseDiscountItem, VehiclePurchaseTransportItem, VehicleTaxOnDeliveryCharges, SpareCentralPurchaseTaxForm, SpareLocalPurchaseTaxForm, SpareCentralSaleTaxForm, SpareLocalSaleTaxForm, DefaultSupervisor, DefaultMechanic, DefaultLabourHead, DefaultPartNo, DefaultOilPartNo  ) " & _
                            "Values('" & Txt(SprDebtorGroupCode).Tag & "', '" & Txt(VehDebtorGroupCode).Tag & "', '" & Txt(WsDebtorGroupCode).Tag & "', '" & Txt(SprCreditorGroupCode).Tag & "', '" & Txt(VehCreditorGroupCode).Tag & "', " & _
                            "'" & Txt(SprSaleAc).Tag & "', '" & Txt(SprSaleVat4Ac).Tag & "', '" & Txt(LubSaleAc).Tag & "', '" & Txt(VehSaleAc).Tag & "', '" & Txt(SprPurchaseAc).Tag & "', '" & Txt(VehPurchaseAc).Tag & "', " & _
                            "'" & Txt(SprCashAc).Tag & "', '" & Txt(VehCashAc).Tag & "', '" & Txt(WsCashAc).Tag & "', '" & Txt(SprBankAc).Tag & "', '" & Txt(VehBankAc).Tag & "', " & _
                            "'" & Txt(WsBankAc).Tag & "', '" & Txt(LocalStateName) & "', '" & Txt(LabourAc).Tag & "', '" & Txt(ServTaxAc).Tag & "', '" & Txt(CstAc).Tag & "', " & _
                            "'" & Txt(VatAc).Tag & "', '" & Txt(ROffAc).Tag & "', '" & Txt(SprCstPurchaseAc).Tag & "','" & Txt(OtherChargesAc).Tag & "', '" & Txt(DiscountAc).Tag & "', " & _
                            "'" & Txt(VehPurGroupCode).Tag & "', '" & Txt(VehSaleGroupCode).Tag & "', '" & Txt(SprPurGroupCode).Tag & "', '" & Txt(SprSaleGroupCode).Tag & "', '" & Txt(VatGroupCode).Tag & "', '" & Txt(ServiceTaxGroupCode).Tag & "', '" & Txt(VehCstPurchaseAc).Tag & "', '" & Txt(Vat4Ac).Tag & "', '" & Txt(SprPurchase4Ac).Tag & "', '" & Txt(VatInputAc).Tag & "', '" & Txt(Vat4InputAc).Tag & "', " & _
                            "'" & Txt(VehicleCentralPurchaseTaxForm).Tag & "', '" & Txt(VehicleLocalPurchaseTaxForm).Tag & "', '" & Txt(VehiclePurchaseDiscountItem).Tag & "', '" & Txt(VehiclePurchaseTransportItem).Tag & "', '" & left(Txt(VehicleTaxOnDeliveryCharges), 1) & "', '" & Txt(SpareCentralPurchaseTaxForm).Tag & "', '" & Txt(SpareLocalPurchaseTaxForm).Tag & "', '" & Txt(SpareCentralSaleTaxForm).Tag & "', '" & Txt(SpareLocalSaleTaxForm).Tag & "', '" & Txt(DefaultSupervisor).Tag & "', '" & Txt(DefaultMechanic).Tag & "', '" & Txt(DefaultLabourHead).Tag & "', '" & Txt(DefaultPartNo).Tag & "', '" & Txt(DefaultOilPartNo).Tag & "')"
    End If
    
    GCn.Execute "Delete from DmsBankAc"
    
    With FGrid1
        For I = 1 To .Rows - 1
            If .TextMatrix(I, F1_BankAc) <> "" Then
                GCn.Execute "Insert Into DmsBankAc(AutomanBankCode, DmsBankCode) Values('" & .TextMatrix(I, F1_BankAcCode) & "', '" & .TextMatrix(I, F1_DmsCode) & "')"
            End If
        Next I
    End With
    
    GCn.Execute "Delete from DmsSupplierAc"
    
    With FGrid2
        For I = 1 To .Rows - 1
            If .TextMatrix(I, F2_SupplierAc) <> "" Then
                GCn.Execute "Insert Into DmsSupplierAc(AutomanSupplierCode, DmsCode) Values('" & .TextMatrix(I, F2_SupplierAcCode) & "', '" & .TextMatrix(I, F1_DmsCode) & "')"
            End If
        Next I
    End With
    
    
    Unload Me
End Sub

Private Sub CmdSparePurchase_Click()
    FImportSparePurchase
End Sub

Private Sub CmdSpareSale_Click()
    FImportSpareSale
End Sub

Private Sub CmdVehiclePurchase_Click()
    Dim mFileName As String
    Dim mFileTitle As String
    
    
    CD1.FileName = ""
    
    Call SelectFile
    mFileName = CD1.FileName
    mFileTitle = CD1.FileTitle
    If mFileName = "" Then Exit Sub
    mFileTitle = mID(mFileTitle, 1, Len(mFileTitle) - 4)
    Set DmsConn = New Connection
    DmsConn.CursorLocation = adUseClient
    DmsConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mFileName & ";Extended Properties=Excel 8.0"
    
    Set RsDms = DmsConn.Execute("Select * from [" & mFileTitle & "$]")

        If ChkFieldExist(RsDms, "Customer Invoice") And ChkFieldExist(RsDms, "Transporter Name") And _
           ChkFieldExist(RsDms, "LoB") And ChkFieldExist(RsDms, "Discount2") And ChkFieldExist(RsDms, "Chassis_No") And ChkFieldExist(RsDms, "Color") And _
           ChkFieldExist(RsDms, "Division") And ChkFieldExist(RsDms, "Engine_Number") And ChkFieldExist(RsDms, "Key_Number") And ChkFieldExist(RsDms, "Parent_Product_Line") And _
           ChkFieldExist(RsDms, "Product Line") And ChkFieldExist(RsDms, "VATTAX") And ChkFieldExist(RsDms, "VAT Classification 1") And ChkFieldExist(RsDms, "VAT Classification 2") And _
           ChkFieldExist(RsDms, "VAT Classification 3") And ChkFieldExist(RsDms, "VAT Classification 4") And _
           ChkFieldExist(RsDms, "TM VAT Rate") And ChkFieldExist(RsDms, "VAT Assessible Amt 1") And _
           ChkFieldExist(RsDms, "VAT Assessible Amt 2") And ChkFieldExist(RsDms, "VAT Assessible Amt 3") And _
           ChkFieldExist(RsDms, "VAT Assessible Amt 4") And ChkFieldExist(RsDms, "Narration") And _
           ChkFieldExist(RsDms, "VC_Number") And ChkFieldExist(RsDms, "Invoice_Date") And _
           ChkFieldExist(RsDms, "Invoice_No") And ChkFieldExist(RsDms, "VC_Description") And _
           ChkFieldExist(RsDms, "Quantiy") And ChkFieldExist(RsDms, "Supplier_Name") And _
           ChkFieldExist(RsDms, "Order_Number") And ChkFieldExist(RsDms, "Total Payble") And _
           ChkFieldExist(RsDms, "Physical Status") And ChkFieldExist(RsDms, "Status") And _
           ChkFieldExist(RsDms, "Value") And ChkFieldExist(RsDms, "Tax CST") And _
           ChkFieldExist(RsDms, "Tax CST for VAT") And ChkFieldExist(RsDms, "Delivery Charges") And _
           ChkFieldExist(RsDms, "Entry Tax") And ChkFieldExist(RsDms, "Tax LST") And _
           ChkFieldExist(RsDms, "Octroi") And ChkFieldExist(RsDms, "Rate") And _
           ChkFieldExist(RsDms, "ST Surcharge") And ChkFieldExist(RsDms, "Tax TOT") And _
           ChkFieldExist(RsDms, "Taxable Amount") And ChkFieldExist(RsDms, "Toll Tax") And _
           ChkFieldExist(RsDms, "Total Discount") And ChkFieldExist(RsDms, "Godown") Then

            Call VehiclePurchaseDataUpdate
        End If
End Sub

Private Sub FgridErr_Click()
    TxtShow = FgridErr
End Sub

Private Sub FgridErr_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Or KeyAscii = 13 Then SendKeysA vbKeyTab, True
End Sub

Private Sub Form_Load()

'    WinSetting Me, 5640, 3550


    BlankAll
    PubImportData = True
    Call AlignCtrls
    Call Ini_Grid(FgridErr)
    Call Ini_Grid(FGrid)
    Call Ini_Grid(FGrid1)
    Call Ini_Grid(FGrid2)
    
    Set RsDmsEnviro = GCn.Execute("Select * From DmsEnviro")
    
  
    Set RsSubGroup = GCn.Execute("Select SubCode As Code, Name From Subgroup Order By Name")
    Set RsAcGroup = G_FaCn.Execute("Select GroupCode As Code, GroupName As Name From AcGroup Order By GroupName")
      
    Set RsTaxForm = GCn.Execute("SELECT T.Form_Code AS Code, T.Form_Desc AS Name, T.L_C, T.Trn_Type, T.Vehicle_YN    FROM TaxForms T ORDER BY T.Form_Desc ")
    Set RsADItem = GCn.Execute("SELECT Prod_Code AS Code, Prod_Name AS Name FROM Veh_AMDModel ORDER BY Prod_Code ")
    
    Set RsMechanic = GCn.Execute("SELECT E.Emp_Code AS Code, E.Emp_Name AS Name FROM Emp_Mast E WHERE E.Designation = 'Mechanic'")
    Set RsSupervisor = GCn.Execute("SELECT E.Emp_Code AS Code, E.Emp_Name AS Name FROM Emp_Mast E WHERE E.Designation = 'Supervisor'")
    Set RsLabour = GCn.Execute("SELECT L.Lab_Code AS code, L.Lab_Desc AS Name FROM Labour L ORDER BY L.Lab_Desc ")
    If RsDmsEnviro.RecordCount = 0 Then MsgBox "Plz Define Settings In DmsEnviro": Exit Sub
    Call MoveRec


    Set RsCity = G_FaCn.Execute("Select * From City Order By CityName")
    Set RsState = G_FaCn.Execute("Select * From State Order by StateName")
End Sub

Sub AlignCtrls()
Dim mDistance As Integer
    mDistance = 90
    
    TopCtrl1.TopText2 = "Edit"
    If mFormType <> ImportForm Then
        Frame4.Move mDistance, mDistance
        Frame4.Visible = True
        'WinSetting Me, 8145, 11000
        WinSetting Me
        
    Else
        
        Frame2.Move mDistance, mDistance
        Frame2.Visible = True
        Frame5.Move mDistance, Frame2.top + Frame2.height + mDistance
        Frame5.Visible = True
        WinSetting Me, 9495, 11940
        
    End If
End Sub
Function SprCounterSale(mBillType As String, mPartyCode As String, mNetAmt As Double, mSprSaleAmt As Double, mLubeSaleAmt As Double, mVatAmt As Double, mNarr As String, mDate As Date, mInvoice_No As String, mDmsDivision As String, mOtherCharges As Double) As Boolean
On Error GoTo lblExit
'A/c Posting related declarations
Dim mVType$, mVPrefix
Dim mVNo As String
Dim mDocId As String
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, I As Integer, j As Integer
Dim RsTemp As ADODB.Recordset
Dim RsDmsDiv As ADODB.Recordset
Dim mROff As Single, mDivCode, mSiteCode
    GCn.CommandTimeout = 20
    G_FaCn.CommandTimeout = 20
    
    If UCase(mBillType) = "CASH" Then
        mVType = PubDmsVTypeSprSaleCash
        mPartyCode = XNull(RsDmsEnviro!SprCashAc)
        If mPartyCode = "" Then
            MsgBox "Please Define Spare Cash A/c In DmsEnviro"
            Exit Function
        End If
    Else
        mVType = PubDmsVTypeSprSaleCredit
    End If
    mVPrefix = "DMS"

    Set RsDmsDiv = GCn.Execute("Select AutomanSite, AutomanDivision From DmsSite Where DmsDivision='" & mDmsDivision & "'")
    If RsDmsDiv.RecordCount > 0 Then
        mDivCode = RsDmsDiv!AutomanDivision
        mSiteCode = RsDmsDiv!AutomanSite

        Set RsTemp = G_FaCn.Execute("Select DocId,V_No From LedgerM With (NOLOCK) Where DmsRefNo='" & mInvoice_No & "'")
        If RsTemp.RecordCount > 0 Then
            mDocId = RsTemp!DocID
            mVNo = RsTemp!V_NO
        Else
            mVNo = G_FaCn.Execute("Select IsNull(Max(V_No)," & Right(date, 1) & "00000" & ")+1 From Ledger  With (NOLOCK)  Where V_Type='" & mVType & "' And RTrim(ltrim(Substring(DocId,9,5)))='DMS' ").Fields(0)
            mDocId = mDivCode + mSiteCode & mSiteCode + Space(5 - Len(CStr(mVType))) + mVType + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(mVNo))) + CStr(mVNo)
            If G_FaCn.Execute("Select Count(*) from Ledger Where DocID='" & mDocId & "'").Fields(0).Value > 0 Then
                MsgBox "DocID Created Already Exist!"
                Exit Function
                Debug.Print mDocId
            End If
            
        End If
    
        
        
        If XNull(RsDmsEnviro!SprSaleAc) = "" Or XNull(RsDmsEnviro!LubeSaleAc) = "" Or XNull(RsDmsEnviro!VatAc) = "" Or XNull(RsDmsEnviro!ROffAc) = "" Then
            MsgBox "Please Define SprSaleAc, LubeSaleAc, VATAc In DMS Enviro"
            Exit Function
        End If
     
        
        mROff = Round(Round(mNetAmt) - mNetAmt, 2)
        
    
                
        ReDim LedgAry(I)
        LedgAry(0).SubCode = mPartyCode
        LedgAry(0).AmtDr = Round(mNetAmt + mROff, 2)
        LedgAry(0).AmtCr = 0
        LedgAry(0).Narration = mNarr
            
            
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!ROffAc
        LedgAry(I).AmtDr = IIf(mROff < 0, Abs(mROff), 0)
        LedgAry(I).AmtCr = IIf(mROff > 0, Abs(mROff), 0)
        LedgAry(I).Narration = mNarr
        
                    
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!SprSaleAc
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mSprSaleAmt, 2)
        LedgAry(I).Narration = mNarr
                
                
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!LubeSaleAc
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mLubeSaleAmt, 2)
        LedgAry(I).Narration = mNarr
                
        
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!VatAc
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mVatAmt, 2)
        LedgAry(I).Narration = mNarr
                
                
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!OtherChargesAc
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mOtherCharges, 2)
        LedgAry(I).Narration = mNarr
                
                
        mResult = LedgerPost("A", LedgAry, G_FaCn, mDocId, CDate(mDate), mNarr)
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation": Exit Function
        
        G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=" & mVNo & " Where V_Type='" & mVType & "'  And Div_Code='" & mDivCode & "' And Prefix ='" & mVPrefix & "'"
        G_FaCn.Execute "Update LedgerM Set DmsRefNo='" & mInvoice_No & "' Where DocId ='" & mDocId & "'"
        
        GCn.Execute "Delete From DmsData Where DmsRefNo='" & mInvoice_No & "'"
        GCn.Execute "Insert Into DmsData (DocId, VType, VDate, VNo, " & _
                    "SubCode, Amount, TaxableAmt, SprAmt, LubeAmt, TaxAmt, DmsRefNo, OtherCharges) " & _
                    "Values('" & mDocId & "', '" & mVType & "', " & ConvertDate(mDate) & ", " & mVNo & ", " & _
                    "'" & mPartyCode & "', " & mNetAmt & ", " & mSprSaleAmt + mLubeSaleAmt & ", " & mSprSaleAmt & ", " & mLubeSaleAmt & ", " & mVatAmt & ", '" & mInvoice_No & "', " & mOtherCharges & ")"
    Else
        CreateErrLog "Spare Sale", mInvoice_No, mDmsDivision & " Not Defined In DmsDivision Table"
    End If
    
    SprCounterSale = True
Exit Function
lblExit:
    Set GRs = Nothing
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function


Function SprSaleReturn(mBillType As String, mPartyCode As String, mNetAmt As Double, mSprSaleAmt As Double, mVatAmt As Double, mNarr As String, mDate As Date, mInvoice_No As String) As Boolean

On Error GoTo lblExit
'A/c Posting related declarations
Dim mVType$, mVPrefix
Dim mVNo As String
Dim mDocId As String
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, I As Integer, j As Integer
Dim RsTemp As ADODB.Recordset
Dim mROff As Single
    
    If UCase(mBillType) = "CASH" Then
        
        mVType = "SXSRC"
    Else
        mVType = "SXSRR"
    End If
    mVPrefix = "DMS"
    
    Set RsTemp = G_FaCn.Execute("Select DocId,V_No From LedgerM  With (NOLOCK)  Where DmsRefNo='" & mInvoice_No & "'")
    If RsTemp.RecordCount > 0 Then
        mDocId = RsTemp!DocID
        mVNo = RsTemp!V_NO
    Else
        mVNo = G_FaCn.Execute("Select IsNull(Max(V_No)," & Right(date, 1) & "00000" & ")+1 From Ledger  With (NOLOCK)  Where V_Type='" & mVType & "' And RTrim(ltrim(Substring(DocId,9,5)))='DMS' ").Fields(0)
        mDocId = PubDivCode + PubSiteCode & PubSiteCode + Space(5 - Len(CStr(mVType))) + mVType + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(mVNo))) + CStr(mVNo)
            If G_FaCn.Execute("Select Count(*) from Ledger Where DocID='" & mDocId & "'").Fields(0).Value > 0 Then
                MsgBox "DocID Created Already Exist!"
                Exit Function
                Debug.Print mDocId
            End If
        
    End If
        
    If XNull(RsDmsEnviro!SprSaleAc) = "" Or XNull(RsDmsEnviro!VatAc) = "" Or XNull(RsDmsEnviro!ROffAc) = "" Then
        MsgBox "Please Define SprSaleAc, VATAc In DMS Enviro"
        Exit Function
    End If

    mROff = Round(Round(mNetAmt) - mNetAmt, 2)
    
                
            
    
    ReDim LedgAry(I)
    LedgAry(0).SubCode = mPartyCode
    LedgAry(0).AmtDr = 0
    LedgAry(0).AmtCr = Round(mNetAmt, 2) + mROff
    LedgAry(0).Narration = mNarr
        
        
    I = UBound(LedgAry) + 1
    ReDim Preserve LedgAry(I)
    LedgAry(I).SubCode = RsDmsEnviro!ROffAc
    LedgAry(I).AmtDr = IIf(mROff > 0, Abs(mROff), 0)
    LedgAry(I).AmtCr = IIf(mROff < 0, Abs(mROff), 0)
    LedgAry(I).Narration = mNarr
        
    I = UBound(LedgAry) + 1
    ReDim Preserve LedgAry(I)
    LedgAry(I).SubCode = RsDmsEnviro!SprSaleAc
    LedgAry(I).AmtDr = Round(mSprSaleAmt, 2)
    LedgAry(I).AmtCr = 0
    LedgAry(I).Narration = mNarr
            
            
    
    I = UBound(LedgAry) + 1
    ReDim Preserve LedgAry(I)
    LedgAry(I).SubCode = RsDmsEnviro!VatAc
    LedgAry(I).AmtDr = Round(mVatAmt, 2)
    LedgAry(I).AmtCr = 0
    LedgAry(I).Narration = mNarr
            
            
    mResult = LedgerPost("A", LedgAry, G_FaCn, mDocId, CDate(mDate), mNarr)
    If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation": Exit Function
    
    
    G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=" & mVNo & " Where V_Type='" & mVType & "'"
    G_FaCn.Execute "Update LedgerM Set DmsRefNo='" & mInvoice_No & "' Where DocId ='" & mDocId & "'"
    
    
    GCn.Execute "Delete From DMS Where DmsRefNo='" & mInvoice_No & "'"
    GCn.Execute "Insert Into DMS (DocId, VType, VDate, VNo, " & _
                "SubCode, Amount, TaxAmt, DmsRefNo) " & _
                "Values('" & mDocId & "', '" & mVType & "', " & ConvertDate(mDate) & ", " & mVNo & ", " & _
                "'" & mPartyCode & "', " & mNetAmt & ", " & mVatAmt & ", '" & mInvoice_No & "')"
    
    
    SprSaleReturn = True
lblExit:
    Set GRs = Nothing
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function


Function WorkShopSale(mBillType As String, mPartyCode As String, mNetAmt As Double, mSprSaleAmt As Double, mSprSaleVat4Amt As Double, mVatAmt As Double, mNarr As String, mDate As Date, mTotalLabour As Double, mTaxableLabour As Double, mServiceTax As Double, mInvoice_No As String, mDmsDivision As String, mOtherCharges As Double, mLabOtherCharges As Double, mDiscount As Double, mLabDiscount As Double, mVat4 As Double) As Boolean
On Error GoTo lblExit
'A/c Posting related declarations
Dim mVTypeSpr$, mVTypeLab$, mVPrefix$
Dim mVnoSpr As String
Dim mVnoLab As String
Dim mDocIdSpr As String
Dim mDocIdLab As String
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, I As Integer, j As Integer
Dim mROff As Single, mSiteCode$, mDivCode$
Dim RsDmsDiv As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
          
    If UCase(mBillType) = "CASH" Then
        
        mVTypeSpr = PubDmsVTypeWorkshopSaleCash
        mVTypeLab = PubDmsVTypeWorkshopSaleCash
        mPartyCode = RsDmsEnviro!WsCashAc
        If mPartyCode = "" Then
            MsgBox "Workshop Cash A/c Is Not Defined In DmsEnviro"
            Exit Function
        End If
    Else
        mVTypeSpr = PubDmsVTypeWorkshopSaleCredit
        mVTypeLab = PubDmsVTypeWorkshopSaleCredit
    End If
        
    mVPrefix = "DMS"
    
    
    Set RsDmsDiv = GCn.Execute("Select AutomanSite, AutomanDivision From DmsSite  With (NOLOCK)  Where DmsDivision='" & mDmsDivision & "'")
    If RsDmsDiv.RecordCount > 0 Then
        mDivCode = RsDmsDiv!AutomanDivision
        mSiteCode = RsDmsDiv!AutomanSite
    
        Set RsTemp = G_FaCn.Execute("Select DocId, V_No From LedgerM  With (NOLOCK)  Where DmsRefNo='" & mInvoice_No & "' And V_Type='" & mVTypeSpr & "'")
        If RsTemp.RecordCount > 0 Then
            mDocIdSpr = RsTemp!DocID
            mVnoSpr = RsTemp!V_NO
        Else
            mVnoSpr = G_FaCn.Execute("Select IsNull(Max(V_No)," & Right(date, 1) & "00000" & ")+1 From Ledger With (NOLOCK)  Where V_Type='" & mVTypeSpr & "' And  RTrim(ltrim(Substring(DocId,9,5)))='DMS' ").Fields(0)
            mDocIdSpr = mDivCode + mSiteCode & mSiteCode + Space(5 - Len(CStr(mVTypeSpr))) + mVTypeSpr + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(mVnoSpr))) + CStr(mVnoSpr)
            
            If G_FaCn.Execute("Select Count(*) from Ledger Where DocID='" & mDocIdSpr & "'").Fields(0).Value > 0 Then
                MsgBox "DocID Created Already Exist!"
                Exit Function
                Debug.Print mDocIdSpr
            End If
            
        End If
        
        
'        Set RsTemp = G_FaCn.Execute("Select DocId,V_No From LedgerM Where DmsRefNo='" & mInvoice_No & "' And V_Type='" & mVTypeLab & "'")
'        If RsTemp.RecordCount > 0 Then
'            mDocIdLab = RsTemp!DocID
'            mVnoLab = RsTemp!V_NO
'        Else
'            mVnoLab = G_FaCn.Execute("Select IIF(IsNull(Max(V_No))," & Right(date, 1) & "00000" & ",Max(V_No))+1 From Ledger Where V_Type='" & mVTypeLab & "'").Fields(0)
'            mDocIdLab = mDivCode + mSiteCode & mSiteCode + Space(5 - Len(CStr(mVTypeLab))) + mVTypeLab + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(mVnoLab))) + CStr(mVnoLab)
'        End If
        mVnoLab = mVnoSpr
        mDocIdLab = mDocIdSpr
        
        
        If XNull(RsDmsEnviro!SprSaleAc) = "" Or XNull(RsDmsEnviro!SprSaleVat4Ac) = "" Or XNull(RsDmsEnviro!VatAc) = "" Or XNull(RsDmsEnviro!Vat4Ac) = "" Or XNull(RsDmsEnviro!LabourAc) = "" Or XNull(RsDmsEnviro!ServTaxAc) = "" Or XNull(RsDmsEnviro!ROffAc) = "" Then
            MsgBox "Please Define SprSaleAc, LubeSaleAc, VATAc, LabourAc, ServTaxAc, ROffAc In DMS Enviro"
            Exit Function
        End If
    
        mROff = Round(Round(mNetAmt + mTotalLabour) - (mNetAmt + mTotalLabour), 2)
        
    
        
        ReDim LedgAry(I)
        LedgAry(0).SubCode = mPartyCode
        LedgAry(0).AmtDr = Round(mNetAmt + mTotalLabour + mROff, 2)
        LedgAry(0).AmtCr = 0
        LedgAry(0).Narration = mNarr
            
        
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!DiscountAc
        LedgAry(I).AmtDr = Round(mDiscount + mLabDiscount, 2) 'Round(mDiscount + mLabDiscount, 2)
        LedgAry(I).AmtCr = 0
        LedgAry(I).Narration = mNarr
        
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!ROffAc
        LedgAry(I).AmtDr = Round(IIf(mROff < 0, Abs(mROff), 0), 2)
        LedgAry(I).AmtCr = Round(IIf(mROff > 0, Abs(mROff), 0), 2)
        LedgAry(I).Narration = mNarr
            
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!SprSaleAc
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mSprSaleAmt, 2)
        LedgAry(I).Narration = mNarr
                
                
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!SprSaleVat4Ac
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mSprSaleVat4Amt, 2)
        LedgAry(I).Narration = mNarr
                
        
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!OtherChargesAc
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mOtherCharges + mLabOtherCharges, 2)
        LedgAry(I).Narration = mNarr
        
        
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!VatAc
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mVatAmt, 2)
        LedgAry(I).Narration = mNarr
                
                
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!Vat4Ac
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mVat4, 2)
        LedgAry(I).Narration = mNarr
                
                
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!LabourAc
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mTaxableLabour, 2)
        LedgAry(I).Narration = mNarr
                
                
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!ServTaxAc
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mServiceTax, 2)
        LedgAry(I).Narration = mNarr
                
                
        mResult = LedgerPost("A", LedgAry, G_FaCn, mDocIdSpr, CDate(mDate), mNarr)
        If mResult <> 1 Then Exit Function
        G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=" & mVnoSpr & " Where V_Type='" & mVTypeSpr & "'"
        G_FaCn.Execute "Update LedgerM Set DmsRefNo='" & mInvoice_No & "' Where DocId ='" & mDocIdSpr & "'"
        
        
'        ReDim LedgAry(I)
'        LedgAry(0).SubCode = mPartyCode
'        LedgAry(0).AmtDr = Round(mTotalLabour, 2)
'        LedgAry(0).AmtCr = 0
'        LedgAry(0).Narration = mNarr
'
'        I = UBound(LedgAry) + 1
'        ReDim Preserve LedgAry(I)
'        LedgAry(I).SubCode = RsDmsEnviro!LabourAc
'        LedgAry(I).AmtDr = 0
'        LedgAry(I).AmtCr = Round(mTaxableLabour, 2)
'        LedgAry(I).Narration = mNarr
'
'
'        I = UBound(LedgAry) + 1
'        ReDim Preserve LedgAry(I)
'        LedgAry(I).SubCode = RsDmsEnviro!ServTaxAc
'        LedgAry(I).AmtDr = 0
'        LedgAry(I).AmtCr = Round(mServiceTax, 2)
'        LedgAry(I).Narration = mNarr
'
'
'        mResult = LedgerPost("A", LedgAry, G_FaCn, mDocIdLab, CDate(mDate), mNarr)
'        If mResult <> 1 Then Exit Function
        
        
        G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=" & mVnoLab & " Where V_Type='" & mVTypeLab & "' And Div_Code='" & mDivCode & "' And Prefix ='" & mVPrefix & "'"
        G_FaCn.Execute "Update LedgerM Set DmsRefNo='" & mInvoice_No & "' Where DocId ='" & mDocIdLab & "'"
        
        
        GCn.Execute "Delete From DmsData Where DmsRefNo='" & mInvoice_No & "'"
        GCn.Execute "Insert Into DmsData (DocId, VType, VDate, VNo, " & _
                    "SubCode, Amount, TaxableAmt, SprAmt, LubeAmt, TaxAmt, DmsRefNo, " & _
                    "Lab_DocId, LabAmount, LabTaxableAmt, SrvTax, OtherCharges, LabOtherCharges, Discount, LabDiscount) " & _
                    "Values('" & mDocIdSpr & "', '" & mVTypeSpr & "', " & ConvertDate(mDate) & ", " & mVnoSpr & ", " & _
                    "'" & mPartyCode & "', " & mNetAmt & ", " & mSprSaleAmt + mSprSaleVat4Amt & ", " & mSprSaleAmt & ", " & mSprSaleVat4Amt & ", " & mVatAmt & ", '" & mInvoice_No & "', " & _
                    "'" & mDocIdLab & "', " & mTotalLabour & ", " & mTaxableLabour & ", " & mServiceTax & ", " & mOtherCharges & ", " & mLabOtherCharges & ", " & mDiscount & ", " & mLabDiscount & ")"
    Else
        CreateErrLog "WorkShop Sale", mInvoice_No, mDmsDivision & " Not Defined In DmsDivision Table"
    End If
    
    
    WorkShopSale = True
Exit Function
lblExit:
    Set GRs = Nothing
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function



Function VehicleSale(mPartyCode As String, mNetAmt As Double, mSaleAmt As Double, mVatAmt As Double, mNarr As String, mDate As Date, mInvoice_No As String, mDmsDivision As String, mChassis As String) As Boolean


On Error GoTo lblExit
'A/c Posting related declarations
Dim mVType$, mVPrefix
Dim mVNo As String
Dim mDocId As String
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, I As Integer, j As Integer
Dim RsTemp As ADODB.Recordset
Dim RsDmsDiv As ADODB.Recordset
Dim mSiteCode$, mDivCode$
Dim mROff As Single
    

    mVType = PubDmsVTypeVehSale
    mVPrefix = "DMS"
    
    Set RsDmsDiv = GCn.Execute("Select AutomanSite, AutomanDivision From DmsSite  With (NOLOCK)  Where DmsDivision='" & mDmsDivision & "'")
    If RsDmsDiv.RecordCount > 0 Then
        mDivCode = RsDmsDiv!AutomanDivision
        mSiteCode = RsDmsDiv!AutomanSite

        Set RsTemp = G_FaCn.Execute("Select DocId,V_No From LedgerM With (NOLOCK) Where DmsRefNo='" & mInvoice_No & "'")
        If RsTemp.RecordCount > 0 Then
            mDocId = RsTemp!DocID
            mVNo = RsTemp!V_NO
        Else
            mVNo = G_FaCn.Execute("Select IsNull(Max(V_No),700000)+1 From Ledger With (NOLOCK)  Where V_Type='" & mVType & "' And RTrim(ltrim(Substring(DocId,9,5)))='DMS' ").Fields(0)
            mDocId = mDivCode + mSiteCode & mSiteCode + Space(5 - Len(CStr(mVType))) + mVType + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(mVNo))) + CStr(mVNo)
            If G_FaCn.Execute("Select Count(*) from Ledger Where DocID='" & mDocId & "'").Fields(0).Value > 0 Then
                MsgBox "DocID Created Already Exist!"
                Exit Function
                Debug.Print mDocId
            End If
        End If
        
        
    
        
        If XNull(RsDmsEnviro!VehPurchaseAc) = "" Or XNull(RsDmsEnviro!VehCstPurchaseAc) = "" Or XNull(RsDmsEnviro!VatAc) = "" Or XNull(RsDmsEnviro!ROffAc) = "" Then
            MsgBox "Please Define SprSaleAc, LubeSaleAc, VATAc In DMS Enviro"
            Exit Function
        End If
    
        
        mROff = Round(Round(mNetAmt) - mNetAmt, 2)
        
                    
        
        ReDim LedgAry(I)
        LedgAry(0).SubCode = mPartyCode
        LedgAry(0).AmtDr = Round(mNetAmt + mROff, 2)
        LedgAry(0).AmtCr = 0
        LedgAry(0).Narration = mNarr
            
        
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!ROffAc
        LedgAry(I).AmtDr = Round(IIf(mROff < 0, Abs(mROff), 0), 2)
        LedgAry(I).AmtCr = Round(IIf(mROff > 0, Abs(mROff), 0), 2)
        LedgAry(I).Narration = mNarr
        
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!VehSaleAc
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mSaleAmt, 2)
        LedgAry(I).Narration = mNarr
        
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = RsDmsEnviro!VatAc
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mVatAmt, 2)
        LedgAry(I).Narration = mNarr
                
                
        mResult = LedgerPost("A", LedgAry, G_FaCn, mDocId, CDate(mDate), mNarr)
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation": Exit Function
        
        
        G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=" & mVNo & " Where V_Type='" & mVType & "'  And Div_Code='" & mDivCode & "' And Prefix ='" & mVPrefix & "'"
        G_FaCn.Execute "Update LedgerM Set DmsRefNo='" & mInvoice_No & "' Where DocId ='" & mDocId & "'"
        
        
        GCn.Execute "Delete From DMSData Where DmsRefNo='" & mInvoice_No & "'"
        GCn.Execute "Insert Into DMSData (DocId, VType, VDate, VNo,  " & _
                    "SubCode, Amount, TaxableAmt, TaxAmt, DmsRefNo, Chassis) " & _
                    "Values('" & mDocId & "', '" & mVType & "', " & ConvertDate(mDate) & ", " & mVNo & ", " & _
                    "'" & mPartyCode & "', " & mNetAmt & ", " & mSaleAmt & ", " & mVatAmt & ", '" & mInvoice_No & "', '" & mChassis & "')"
    Else
        CreateErrLog "Vehicle Sale", mInvoice_No, mDmsDivision & " Not Defined In DmsDivision Table"
    End If
    
    VehicleSale = True
    
Exit Function
lblExit:
    Set GRs = Nothing
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function


Function VehiclePurchase(mPartyCode As String, mNetAmt As Double, mPurchaseAmt As Double, mVatAmt As Double, mCstAmt As Double, mNarr As String, mDate As Date, mInvoice_No As String, mDmsDivision As String, mChassis As String) As Boolean


On Error GoTo lblExit
'A/c Posting related declarations
Dim mVType$, mVPrefix
Dim mVNo As String
Dim mDocId As String
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, I As Integer, j As Integer
Dim RsTemp As ADODB.Recordset
Dim RsDmsDiv As ADODB.Recordset
Dim mSiteCode$, mDivCode$
Dim mROff As Single
    

    mVType = PubDmsVTypeVehPur
    mVPrefix = "DMS"
    
    Set RsDmsDiv = GCn.Execute("Select AutomanSite, AutomanDivision From DmsSite with (NOLOCK) Where DmsDivision='" & mDmsDivision & "'")
    If RsDmsDiv.RecordCount > 0 Then
        mDivCode = RsDmsDiv!AutomanDivision
        mSiteCode = RsDmsDiv!AutomanSite

        Set RsTemp = G_FaCn.Execute("Select DocId,V_No From LedgerM WITH (NOLOCK) Where DmsRefNo='" & mInvoice_No & "'")
        If RsTemp.RecordCount > 0 Then
            mDocId = RsTemp!DocID
            mVNo = RsTemp!V_NO
        Else
            mVNo = G_FaCn.Execute("Select IsNull(Max(V_No),700000)+1 From Ledger WITH (NOLOCK) Where V_Type='" & mVType & "' And  RTrim(ltrim(Substring(DocId,9,5)))='DMS' ").Fields(0)
            mDocId = mDivCode + mSiteCode & mSiteCode + Space(5 - Len(CStr(mVType))) + mVType + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(mVNo))) + CStr(mVNo)
            If G_FaCn.Execute("Select Count(*) from Ledger Where DocID='" & mDocId & "'").Fields(0).Value > 0 Then
                MsgBox "DocID Created Already Exist!"
                Exit Function
                Debug.Print mDocId
            End If
            
        End If
        
        
    
        
        If XNull(RsDmsEnviro!VehPurchaseAc) = "" Or XNull(RsDmsEnviro!VehCstPurchaseAc) = "" Or XNull(RsDmsEnviro!VatAc) = "" Or XNull(RsDmsEnviro!ROffAc) = "" Then
            MsgBox "Please Define Vehicle Purchase A/c,  VAT A/c, Round Off A/c In DMS Enviro"
            Exit Function
        End If
    
        
        'mROff = Round(Round(mNetAmt) - mNetAmt, 2)
        
                    
        
        ReDim LedgAry(I)
        LedgAry(0).SubCode = mPartyCode
        LedgAry(0).AmtDr = 0
        LedgAry(0).AmtCr = Round(mNetAmt + mCstAmt, 2) '+ mROff
        LedgAry(0).Narration = mNarr
            
        
'        I = UBound(LedgAry) + 1
'        ReDim Preserve LedgAry(I)
'        LedgAry(I).SubCode = RsDmsEnviro!ROffAc
'        LedgAry(I).AmtDr = IIf(mROff > 0, Abs(mROff), 0)
'        LedgAry(I).AmtCr = IIf(mROff < 0, Abs(mROff), 0)
'        LedgAry(I).Narration = mNarr
        
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = IIf(mCstAmt > 0, RsDmsEnviro!VehCstPurchaseAc, RsDmsEnviro!VehPurchaseAc)
        LedgAry(I).AmtDr = Round(mPurchaseAmt + mCstAmt, 2)
        LedgAry(I).AmtCr = 0
        LedgAry(I).Narration = mNarr
        
        If mVatAmt > 0 Then
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = RsDmsEnviro!VatInputAc
            LedgAry(I).AmtDr = Round(mVatAmt, 2)
            LedgAry(I).AmtCr = 0
            LedgAry(I).Narration = mNarr
        End If
                
                
        mResult = LedgerPost("A", LedgAry, G_FaCn, mDocId, CDate(mDate), mNarr)
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation": Exit Function
        
        
        G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=" & mVNo & " Where V_Type='" & mVType & "'  And Div_Code='" & mDivCode & "' And Prefix ='" & mVPrefix & "'"
        G_FaCn.Execute "Update LedgerM Set DmsRefNo='" & mInvoice_No & "' Where DocId ='" & mDocId & "'"
        
        
        GCn.Execute "Delete From DMSData Where DmsRefNo='" & mInvoice_No & "'"
        GCn.Execute "Insert Into DMSData (DocId, VType, VDate, VNo,  " & _
                    "SubCode, Amount, TaxableAmt, TaxAmt, DmsRefNo, Chassis) " & _
                    "Values('" & mDocId & "', '" & mVType & "', " & ConvertDate(mDate) & ", " & mVNo & ", " & _
                    "'" & mPartyCode & "', " & mNetAmt & ", " & mPurchaseAmt & ", " & mVatAmt & ", '" & mInvoice_No & "', '" & mChassis & "')"
    Else
        CreateErrLog "Vehicle Purchase", mInvoice_No, mDmsDivision & " Not Defined In DmsDivision Table"
    End If
    
    VehiclePurchase = True
Exit Function
lblExit:
    Set GRs = Nothing
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function


Function SparePurchase(mPartyCode As String, mNetAmt As Double, mPurchase12Amt As Double, mPurchase4Amt As Double, mVat12Amt As Double, mVat4Amt, mNarr As String, mDate As Date, mLocalCentral As String, mInvoice_No As String, mDmsDivision As String) As Boolean

'A/c Posting related declarations
Dim mVType$, mVPrefix
Dim mVNo As String
Dim mDocId As String
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, I As Integer, j As Integer
Dim RsTemp As ADODB.Recordset
Dim RsDmsDiv As ADODB.Recordset
Dim mROff As Single
Dim mSiteCode$, mDivCode$
Dim mVatAmt As Double, mPurchaseAmt As Double

    mVType = PubDmsVTypeSprPurCredit
    mVPrefix = "DMS"
    
    mVatAmt = IIf(UCase(mLocalCentral) = "LOCAL", mVat4Amt, mVat12Amt)
    mPurchaseAmt = IIf(UCase(mLocalCentral) = "LOCAL", mPurchase4Amt, mPurchase12Amt)
    
    Set RsDmsDiv = GCn.Execute("Select AutomanSite, AutomanDivision From DmsSite WITH (NOLOCK) Where DmsDivision='" & mDmsDivision & "'")
    If RsDmsDiv.RecordCount > 0 Then
        mDivCode = RsDmsDiv!AutomanDivision
        mSiteCode = RsDmsDiv!AutomanSite
    
    
        Set RsTemp = G_FaCn.Execute("Select DocId, V_No From LedgerM WITH (NOLOCK) Where DmsRefNo='" & mInvoice_No & "'")
        If RsTemp.RecordCount > 0 Then
            mDocId = RsTemp!DocID
            mVNo = RsTemp!V_NO
        Else
            mVNo = G_FaCn.Execute("Select IsNull(Max(V_No)," & Right(date, 1) & "00000" & ")+1 From Ledger With (NoLock) Where V_Type='" & mVType & "'  And RTrim(ltrim(Substring(DocId,9,5)))='DMS' ").Fields(0)
            mDocId = mDivCode + mSiteCode & mSiteCode + Space(5 - Len(CStr(mVType))) + mVType + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(mVNo))) + CStr(mVNo)
            If G_FaCn.Execute("Select Count(*) from Ledger Where DocID='" & mDocId & "'").Fields(0).Value > 0 Then
                MsgBox "DocID Created Already Exist!"
                Exit Function
                Debug.Print mDocId
            End If
            
        End If
            
        
        If XNull(RsDmsEnviro!SprPurchaseAc) = "" Or XNull(RsDmsEnviro!SprCstPurchaseAc) = "" Or XNull(RsDmsEnviro!VatAc) = "" Or XNull(RsDmsEnviro!CstAc) = "" Or XNull(RsDmsEnviro!ROffAc) = "" Then
            MsgBox "Please Define SprPurchaseAc, SprCstPurchaseAc, VATAc, CstAc In DMS Enviro"
            Exit Function
        End If
    
        
        'mROff = Round(Round(mNetAmt) - mNetAmt, 2)
        
            
        ReDim LedgAry(I)
        LedgAry(0).SubCode = IIf(UCase(mLocalCentral) = "LOCAL", RsDmsEnviro!SprPurchaseAc, RsDmsEnviro!SprCstPurchaseAc)
        LedgAry(0).AmtDr = Round(mPurchase12Amt, 2)
        LedgAry(0).AmtCr = 0
        LedgAry(0).Narration = mNarr
        
        
        If UCase(mLocalCentral) = "LOCAL" Then
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = RsDmsEnviro!SprPurchase4Ac
            LedgAry(I).AmtDr = Round(mPurchase4Amt, 2)
            LedgAry(I).AmtCr = 0
            LedgAry(I).Narration = mNarr
        End If
        
    
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = IIf(UCase(mLocalCentral) = "LOCAL", RsDmsEnviro!VatInputAc, RsDmsEnviro!CstAc)
        LedgAry(I).AmtDr = Round(mVat12Amt, 2)
        LedgAry(I).AmtCr = 0
        LedgAry(I).Narration = mNarr
        
        
        If UCase(mLocalCentral) = "LOCAL" Then
            I = UBound(LedgAry) + 1
            ReDim Preserve LedgAry(I)
            LedgAry(I).SubCode = RsDmsEnviro!Vat4InputAc
            LedgAry(I).AmtDr = Round(mVat4Amt, 2)
            LedgAry(I).AmtCr = 0
            LedgAry(I).Narration = mNarr
        End If
        
        
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = mPartyCode
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mNetAmt, 2) '+ mROff
        LedgAry(I).Narration = mNarr
                                                                             
        mResult = LedgerPost("A", LedgAry, G_FaCn, mDocId, CDate(mDate), mNarr)
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation": Exit Function
                
        G_FaCn.Execute "Update LedgerM Set DmsRefNo='" & mInvoice_No & "' Where DocId ='" & mDocId & "'"
        G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=" & mVNo & " Where V_Type='" & mVType & "'  And Div_Code='" & mDivCode & "' And Prefix ='" & mVPrefix & "'"
        
        GCn.Execute "Delete From DmsData Where DmsRefNo='" & mInvoice_No & "'"
        GCn.Execute "Insert Into DmsData (DocId, VType, VDate, VNo, L_C, " & _
                    "SubCode, Amount,TaxableAmt, SprAmt, TaxAmt, DmsRefNo) " & _
                    "Values('" & mDocId & "', '" & mVType & "', " & ConvertDate(mDate) & ", " & mVNo & ", '" & mLocalCentral & "', " & _
                    "'" & mPartyCode & "', " & mNetAmt & ", " & mPurchaseAmt & ", " & mPurchaseAmt & ", " & mVatAmt & ", '" & mInvoice_No & "')"
    Else
        CreateErrLog "Spare Purchase", mInvoice_No, mDmsDivision & " Not Defined In DmsDivision Table"
    End If
    
    SparePurchase = True
    
End Function


Function SupplierPayment(mPartyCode As String, mNetAmt As Double, mNarr As String, mDate As Date, mCashCredit As String, mChqNo As String, mChqDate As String, mInvoice_No As String) As Boolean
On Error GoTo lblExit
'A/c Posting related declarations
Dim mVType$, mVPrefix
Dim mVNo As String
Dim mDocId As String
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, I As Integer, j As Integer
Dim RsTemp As ADODB.Recordset
    
    If UCase(mCashCredit) = "CASH" Then
        mVType = "G_BCP"
    Else
        mVType = "G_BBP"
    End If
        
    mVType = IIf(UCase(mCashCredit) = "CASH", "G_BCP", "G_BBP")
    mVPrefix = "DMS"
    
    
    Set RsTemp = G_FaCn.Execute("Select DocId,V_No From LedgerM WITH (NOLOCK) Where DmsRefNo='" & mInvoice_No & "'")
    If RsTemp.RecordCount > 0 Then
        mDocId = RsTemp!DocID
        mVNo = RsTemp!V_NO
    Else
        mVNo = G_FaCn.Execute("Select IsNull(Max(V_No)," & Right(date, 1) & "00000" & ")+1 From Ledger WITH (NOLOCK) Where V_Type='" & mVType & "' And RTrim(ltrim(Substring(DocId,9,5)))='DMS' ").Fields(0)
        mDocId = PubDivCode + PubSiteCode & PubSiteCode + Space(5 - Len(CStr(mVType))) + mVType + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(mVNo))) + CStr(mVNo)
            If G_FaCn.Execute("Select Count(*) from Ledger Where DocID='" & mDocId & "'").Fields(0).Value > 0 Then
                MsgBox "DocID Created Already Exist!"
                Exit Function
                Debug.Print mDocId
            End If
        
    End If
    
    
    
    If XNull(RsDmsEnviro!SprCashAc) = "" Or XNull(RsDmsEnviro!SprBankAc) = "" Then
        MsgBox "Please Define SprCashAc, SprBankAc In DMS Enviro"
        Exit Function
    End If

            
    ReDim LedgAry(I)
    LedgAry(0).SubCode = mPartyCode
    LedgAry(0).AmtDr = Round(mNetAmt, 2)
    LedgAry(0).AmtCr = 0
    LedgAry(0).Narration = mNarr
    
    
    I = UBound(LedgAry) + 1
    ReDim Preserve LedgAry(I)
    LedgAry(I).SubCode = IIf(UCase(mCashCredit) = "CASH", RsDmsEnviro!SprCashAc, RsDmsEnviro!SprBankAc)
    LedgAry(I).AmtDr = 0
    LedgAry(I).AmtCr = Round(mNetAmt, 2)
    LedgAry(I).Narration = mNarr
    
            
    mResult = LedgerPost("A", LedgAry, G_FaCn, mDocId, CDate(mDate), mNarr)
    If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation": Exit Function
    G_FaCn.Execute "Update Ledger Set Chq_No='" & mChqNo & "', Chq_Date=" & ConvertDate(mChqDate) & " Where DocId='" & mDocId & "'"
    G_FaCn.Execute "Update LedgerM Set DmsRefNo='" & mInvoice_No & "' Where DocId ='" & mDocId & "'"
    G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=" & mVNo & " Where V_Type='" & mVType & "'  "
    SupplierPayment = True
lblExit:
    Set GRs = Nothing
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function



Function MoneyRect(mPaymentMode As String, mPartyCode As String, mNetAmt As Double, mNarr As String, mDate As Date, mChqNo As String, mChqDate As String, mInvoice_No As String, mDmsDivision As String, mDepositedBank As String, mType As String) As Boolean
On Error GoTo lblExit
'A/c Posting related declarations
Dim mCashBank$
Dim mVType$, mVPrefix
Dim mVNo As String
Dim mDocId As String
Dim LedgAry() As LedgRec, mCommNarr$
Dim mResult As Byte, I As Integer, j As Integer
Dim RsTemp As ADODB.Recordset
Dim RsDmsDiv As ADODB.Recordset
Dim mSiteCode$, mDivCode$
Dim mCashAc$

    If mType = "Vehicle Money Receipt" Then
        If XNull(RsDmsEnviro!VehCashAc) = "" Or XNull(RsDmsEnviro!VehBankAc) = "" Then
            MsgBox "Please Define Vehicle Cash A/c, Vehicle Bank A/c In DMS Enviro"
            Exit Function
        End If
    Else
        If XNull(RsDmsEnviro!SprCashAc) = "" Or XNull(RsDmsEnviro!SprBankAc) = "" Then
            MsgBox "Please Define Spare Cash A/c, Spare Bank A/c In DMS Enviro"
            Exit Function
        End If
    End If

    If mPaymentMode = "Cash" Then
        mCashBank = "Cash"
        
        If mType = "Vehicle Money Receipt" Then
            mDepositedBank = RsDmsEnviro!VehCashAc
        Else
            mDepositedBank = RsDmsEnviro!SprCashAc
        End If
        
    Else
        mCashBank = "Bank"
        If mDepositedBank <> "" Then
        Else
            If mType = "Vehicle Money Receipt" Then
                mDepositedBank = RsDmsEnviro!VehBankAc
            Else
                mDepositedBank = RsDmsEnviro!SprBankAc
            End If
        End If
    End If

    mVType = IIf(UCase(mCashBank) = "CASH", PubDmsVTypeMoneyRectCash, PubDmsVTypeMoneyRectBank)
    mVPrefix = "DMS"
    
    
    Set RsDmsDiv = GCn.Execute("Select AutomanSite, AutomanDivision From DmsSite WITH (NOLOCK) Where DmsDivision='" & mDmsDivision & "'")
    If RsDmsDiv.RecordCount > 0 Then
        mDivCode = RsDmsDiv!AutomanDivision
        mSiteCode = RsDmsDiv!AutomanSite
    
    
        Set RsTemp = G_FaCn.Execute("Select DocId,V_No From LedgerM WITH (NOLOCK) Where DmsRefNo='" & mInvoice_No & "' And V_Type='" & mVType & "'")
        If RsTemp.RecordCount > 0 Then
            mDocId = RsTemp!DocID
            mVNo = RsTemp!V_NO
        Else
            mVNo = G_FaCn.Execute("Select IsNull(Max(V_No)," & Right(date, 1) & "00000" & ")+1 From Ledger WITH (NOLOCK) Where V_Type='" & mVType & "' And RTrim(ltrim(Substring(DocId,9,5)))='DMS' ").Fields(0)
            mDocId = mDivCode + mSiteCode & mSiteCode + Space(5 - Len(CStr(mVType))) + mVType + Space(5 - Len(CStr(mVPrefix))) + mVPrefix + Space(8 - Len(CStr(mVNo))) + CStr(mVNo)
            If G_FaCn.Execute("Select Count(*) from Ledger Where DocID='" & mDocId & "'").Fields(0).Value > 0 Then
                MsgBox "DocID Created Already Exist!"
                Exit Function
                Debug.Print mDocId
            End If
            
        End If
                
        
    
                
        ReDim LedgAry(I)
        LedgAry(0).SubCode = IIf(UCase(mCashBank) = "CASH", RsDmsEnviro!VehCashAc, mDepositedBank)
        LedgAry(0).AmtDr = Round(mNetAmt, 2)
        LedgAry(0).AmtCr = 0
        LedgAry(0).Narration = mNarr
        
        
        I = UBound(LedgAry) + 1
        ReDim Preserve LedgAry(I)
        LedgAry(I).SubCode = mPartyCode
        LedgAry(I).AmtDr = 0
        LedgAry(I).AmtCr = Round(mNetAmt, 2)
        LedgAry(I).Narration = mNarr
        
                
        mResult = LedgerPost("A", LedgAry, G_FaCn, mDocId, CDate(mDate), mNarr)
        If mResult <> 1 Then MsgBox "Error in Ledger Posting", vbOKOnly, "Validation": Exit Function
        G_FaCn.Execute "Update Ledger Set Chq_No='" & mChqNo & "', Chq_Date=" & ConvertDate(mChqDate) & " Where DocId='" & mDocId & "'"
        G_FaCn.Execute "Update LedgerM Set DmsRefNo='" & mInvoice_No & "' Where DocId ='" & mDocId & "'"
        G_FaCn.Execute "Update Voucher_Prefix Set Start_Srl_No=" & mVNo & " Where V_Type='" & mVType & "'  And Div_Code='" & mDivCode & "' And Prefix ='" & mVPrefix & "'"
    Else
    End If
    
    
    MoneyRect = True
    
Exit Function
lblExit:
    Set GRs = Nothing
    If err.NUMBER <> 0 Then MsgBox err.Description
End Function






Private Sub Form_Unload(Cancel As Integer)
    
    Dim mSubGroupCounter As Double
    PubImportData = False
    'If mIsAnySubCodeCreated Then
        mSubGroupCounter = G_FaCn.Execute("Select " & vIsNull("Max(" & cVal("Right(SubCode,6)") & ")", "0") & "+1 From SubGroup WITH (NOLOCK)").Fields(0)
        G_CompCn.Execute "Update SubGroupCounter Set SubGroupAcCode=" & mSubGroupCounter & ""
        G_FaCn.Execute "Drop Table SubGroupAlias"
        G_FaCn.Execute "Select * Into SubGroupAlias From SubGroup WITH (NOLOCK)"
        If PubBackEnd = "A" Then
            GCn.Execute "Drop Table SubGroupAlias"
            GCn.Execute "Select * Into SubGroupAlias  From SubGroup WITH (NOLOCK)"
        End If
    'End If
    
    
    Set RsAcGroup = Nothing
    Set RsDmsEnviro = Nothing
    Set RsHelp = Nothing
    Set RsState = Nothing
    Set RsSubGroup = Nothing
End Sub



'Private Sub Timer1_Timer()
'Dim i As Integer
'    For i = 1 To 999
'        LblTimer.Refresh
'    Next i
'End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Select Case Index
        Case VehicleCentralPurchaseTaxForm
            Set RsHelp = RsTaxForm.Clone
            DgHelp.Columns(1).CAPTION = "Tax Form"
            RsHelp.Filter = adFilterNone
            RsHelp.Filter = " L_C = 'Central' And Trn_Type = 'Purchase' And Vehicle_Yn=1 "
        Case VehicleLocalPurchaseTaxForm
            Set RsHelp = RsTaxForm.Clone
            DgHelp.Columns(1).CAPTION = "Tax Form"
            RsHelp.Filter = adFilterNone
            RsHelp.Filter = " L_C = 'Local' And Trn_Type = 'Purchase' And Vehicle_Yn=1 "
        
        Case SpareCentralPurchaseTaxForm
            Set RsHelp = RsTaxForm.Clone
            DgHelp.Columns(1).CAPTION = "Tax Form"
            RsHelp.Filter = adFilterNone
            RsHelp.Filter = " L_C = 'Central' And Trn_Type = 'Purchase' And Vehicle_Yn=0 "
        
        Case SpareLocalPurchaseTaxForm
            Set RsHelp = RsTaxForm.Clone
            DgHelp.Columns(1).CAPTION = "Tax Form"
            RsHelp.Filter = adFilterNone
            RsHelp.Filter = " L_C = 'Local' And Trn_Type = 'Purchase' And Vehicle_Yn=0 "
        
        Case SpareCentralSaleTaxForm
            Set RsHelp = RsTaxForm.Clone
            DgHelp.Columns(1).CAPTION = "Tax Form"
            RsHelp.Filter = adFilterNone
            RsHelp.Filter = " L_C = 'Central' And Trn_Type = 'Sale' And Vehicle_Yn=0 "
        
        Case SpareLocalSaleTaxForm
            Set RsHelp = RsTaxForm.Clone
            DgHelp.Columns(1).CAPTION = "Tax Form"
            RsHelp.Filter = adFilterNone
            RsHelp.Filter = " L_C = 'Local' And Trn_Type = 'Sale' And Vehicle_Yn=0 "
            
        Case DefaultMechanic
            Set RsHelp = RsMechanic.Clone
            DgHelp.Columns(1).CAPTION = "Mechanic"
            
        Case DefaultSupervisor
            Set RsHelp = RsSupervisor.Clone
            DgHelp.Columns(1).CAPTION = "Supervisor"
            
        Case DefaultLabourHead
            Set RsHelp = RsLabour.Clone
            DgHelp.Columns(1).CAPTION = "Labour Head"
            
        Case DefaultPartNo
            Set RsHelp = RsPart.Clone
            RsHelp.Filter = adFilterNone
            RsHelp.Filter = " Part_Grade <> '" & PubPartGrade_Lub & "' "
            DgHelp.Columns(1).CAPTION = "Part No"
            Set DgHelp.DataSource = RsHelp
        Case DefaultOilPartNo
            Set RsHelp = RsPart.Clone
            RsHelp.Filter = adFilterNone
            RsHelp.Filter = " Part_Grade = '" & PubPartGrade_Lub & "' "
            DgHelp.Columns(1).CAPTION = "Part No"
            Set DgHelp.DataSource = RsHelp
            
            
        Case VehiclePurchaseDiscountItem, VehiclePurchaseTransportItem
            Set RsHelp = RsADItem.Clone
        
        Case SprDebtorGroupCode, WsDebtorGroupCode, VehDebtorGroupCode, SprCreditorGroupCode, VehCreditorGroupCode, VehSaleGroupCode, VehPurGroupCode, SprSaleGroupCode, SprPurGroupCode, VatGroupCode, ServiceTaxGroupCode
            Set RsHelp = G_FaCn.Execute("Select GroupCode As Code, GroupName As Name From AcGroup WITH (NOLOCK) Order By GroupName")
            DgHelp.Columns(1).CAPTION = "Group Name"
        Case SprPurchaseAc, VehPurchaseAc, SprCstPurchaseAc, VehCstPurchaseAc, SprPurchase4Ac
            CreateAcHelp "Purchase"
        Case SprSaleAc, LubSaleAc, VehSaleAc, SprSaleVat4Ac
            CreateAcHelp "Sale"
        Case SprCashAc, WsCashAc, VehCashAc
            CreateAcHelp "Cash"
        Case SprBankAc, WsBankAc, VehBankAc
            CreateAcHelp "Bank"
        Case VatAc, CstAc, ServTaxAc, LabourAc, ROffAc, OtherChargesAc, DiscountAc, Vat4Ac, VatInputAc, Vat4InputAc
            CreateAcHelp
    End Select
    
    
    
    Select Case Index
        Case SprDebtorGroupCode, WsDebtorGroupCode, VehDebtorGroupCode, SprCreditorGroupCode, _
                    VehSaleGroupCode, VehPurGroupCode, SprSaleGroupCode, SprPurGroupCode, _
                    VatGroupCode, ServiceTaxGroupCode, VehCreditorGroupCode, VehCashAc, _
                    SprCashAc, WsCashAc, VehBankAc, SprBankAc, WsBankAc, VehSaleAc, SprSaleAc, _
                    LubSaleAc, VehPurchaseAc, SprPurchaseAc, LabourAc, ServTaxAc, CstAc, VatAc, Vat4Ac, SprSaleVat4Ac, _
                    ROffAc, SprCstPurchaseAc, OtherChargesAc, DiscountAc, VehCstPurchaseAc, SprPurchase4Ac, VatInputAc, Vat4InputAc, _
                    VehicleCentralPurchaseTaxForm, VehicleLocalPurchaseTaxForm, SpareCentralPurchaseTaxForm, SpareLocalPurchaseTaxForm, _
                    SpareCentralSaleTaxForm, SpareLocalSaleTaxForm, VehiclePurchaseDiscountItem, VehiclePurchaseTransportItem, DefaultMechanic, DefaultSupervisor, DefaultLabourHead, DefaultPartNo, DefaultOilPartNo
             
            Set DgHelp.DataSource = RsHelp
            DgHelp.Move Txt(Index).left, Txt(Index).top + Txt(Index).height + 20
            If RsHelp.RecordCount > 0 Then
                RsHelp.FIND "Code = '" & Txt(Index).Tag & "'"
                If RsHelp.EOF = False Then
                    Txt(Index) = XNull(RsHelp.Fields("Name"))
                End If
            End If
    End Select
    
End Sub

Private Sub CreateAcHelp(Optional mNature As String)
    Dim mCondStr$
    
    
    If mNature <> "" Then mCondStr = "Where AG.Nature = '" & mNature & "'"
    
    Set RsHelp = G_FaCn.Execute("Select SubCode As Code, Name As Name From SubGroup SG Left Join AcGroup AG On AG.GroupCode=SG.GroupCode " & mCondStr & " Order By Name")
    DgHelp.Columns(1).CAPTION = mNature & " Account Name"
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case SprDebtorGroupCode, WsDebtorGroupCode, VehDebtorGroupCode, SprCreditorGroupCode, _
                    VehCreditorGroupCode, VehCashAc, SprCashAc, WsCashAc, VehBankAc, SprBankAc, _
                    WsBankAc, VehSaleAc, SprSaleAc, LubSaleAc, VehPurchaseAc, SprPurchaseAc, SprSaleVat4Ac, _
                    LabourAc, ServTaxAc, CstAc, VatAc, Vat4Ac, ROffAc, SprCstPurchaseAc, OtherChargesAc, DiscountAc, _
                    VehSaleGroupCode, VehPurGroupCode, SprSaleGroupCode, SprPurGroupCode, _
                    VatGroupCode, ServiceTaxGroupCode, VehCstPurchaseAc, SprPurchase4Ac, VatInputAc, Vat4InputAc, _
                    VehicleCentralPurchaseTaxForm, VehicleLocalPurchaseTaxForm, SpareCentralPurchaseTaxForm, SpareLocalPurchaseTaxForm, _
                    SpareCentralSaleTaxForm, SpareLocalSaleTaxForm, VehiclePurchaseDiscountItem, VehiclePurchaseTransportItem, DefaultMechanic, DefaultSupervisor, DefaultLabourHead, DefaultPartNo, DefaultOilPartNo
                    
            DGridTxtKeyDown DgHelp, Txt, Index, RsHelp, KeyCode, False, 1
    End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case SprDebtorGroupCode, WsDebtorGroupCode, VehDebtorGroupCode, SprCreditorGroupCode, _
                    VehCreditorGroupCode, VehCashAc, SprCashAc, WsCashAc, VehBankAc, SprBankAc, _
                    WsBankAc, VehSaleAc, SprSaleAc, LubSaleAc, VehPurchaseAc, SprPurchaseAc, SprSaleVat4Ac, _
                    LabourAc, ServTaxAc, CstAc, VatAc, Vat4Ac, ROffAc, SprCstPurchaseAc, OtherChargesAc, DiscountAc, _
                    VehSaleGroupCode, VehPurGroupCode, SprSaleGroupCode, SprPurGroupCode, _
                    VatGroupCode, ServiceTaxGroupCode, VehCstPurchaseAc, SprPurchase4Ac, VatInputAc, Vat4InputAc, _
                    VehicleCentralPurchaseTaxForm, VehicleLocalPurchaseTaxForm, SpareCentralPurchaseTaxForm, SpareLocalPurchaseTaxForm, _
                    SpareCentralSaleTaxForm, SpareLocalSaleTaxForm, VehiclePurchaseDiscountItem, VehiclePurchaseTransportItem, DefaultMechanic, DefaultSupervisor, DefaultLabourHead, DefaultPartNo, DefaultOilPartNo
                    
            If DgHelp.Visible = True Then DGridTxtKeyPress Txt, Index, RsHelp, KeyAscii, "Name"
    End Select
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case SprDebtorGroupCode, WsDebtorGroupCode, VehDebtorGroupCode, SprCreditorGroupCode, _
                    VehCreditorGroupCode, VehCashAc, SprCashAc, WsCashAc, VehBankAc, SprBankAc, _
                    WsBankAc, VehSaleAc, SprSaleAc, LubSaleAc, VehPurchaseAc, SprPurchaseAc, SprSaleVat4Ac, _
                    LabourAc, ServTaxAc, CstAc, VatAc, Vat4Ac, ROffAc, SprCstPurchaseAc, OtherChargesAc, _
                    DiscountAc, VehSaleGroupCode, VehPurGroupCode, SprSaleGroupCode, _
                    SprPurGroupCode, VatGroupCode, ServiceTaxGroupCode, VehCstPurchaseAc, SprPurchase4Ac, VatInputAc, Vat4InputAc, _
                    VehicleCentralPurchaseTaxForm, VehicleLocalPurchaseTaxForm, SpareCentralPurchaseTaxForm, SpareLocalPurchaseTaxForm, _
                    SpareCentralSaleTaxForm, SpareLocalSaleTaxForm, VehiclePurchaseDiscountItem, VehiclePurchaseTransportItem, DefaultMechanic, DefaultSupervisor, DefaultLabourHead, DefaultPartNo, DefaultOilPartNo
                    
            If RsHelp.RecordCount = 0 Or (RsHelp.EOF = True Or RsHelp.BOF = True) Or Txt(Index).TEXT = "" Then
                Txt(Index).TEXT = ""
                Txt(Index).Tag = ""
            Else
                Txt(Index).TEXT = RsHelp!Name
                Txt(Index).Tag = RsHelp!Code
            End If
        Case ToDate, FromDate
            Txt(Index) = RetDate(Txt(Index))
    End Select
End Sub



Function AutomanSubcode(mDmsSubCode As String, mAutomanGroupCode As String, mNature As String) As String
    Dim mConn As New ADODB.Connection
    Dim RsTemp As ADODB.Recordset
    Dim rsTemp1 As ADODB.Recordset
    Dim RsTempCity As ADODB.Recordset
    Dim mSubGroupCounter As Long
    Dim mSubCode$, mQry$, mname$, mCityCode$, mStateCode$
    Dim mLocalCentral$
    
    
    mConn.CursorLocation = adUseClient
    mConn.ConnectionString = G_FaCn.ConnectionString
    mConn.Open
    If mNature = "Supplier" Then
        Set RsTemp = mConn.Execute("Select AutomanSupplierCode From DmsSupplier With (NOLOCK) Where DmsSupplierCode = '" & mDmsSubCode & "' And AutomanSupplierCode Is Not Null And AutomanSupplierCode<>''  And AutomanSupplierCode<>'0'")
        If RsTemp.RecordCount > 0 Then
            AutomanSubcode = RsTemp!AutomanSupplierCode
            Exit Function
        End If
    End If
    
    Set RsTemp = GCn.Execute("Select AutomanSubCode From DmsSubGroup With (NOLOCK) Where DmsSubCode = '" & mDmsSubCode & "' And AutomanSubCode Is Not Null And AutomanSubCode<>''   And AutomanSubCode<>'0'")
    If RsTemp.RecordCount > 0 Then
        AutomanSubcode = XNull(RsTemp!AutomanSubcode)
    End If
    
    'If AutomanSubcode = "" Then
        Set RsTemp = mConn.Execute("Select SubCode From SubGroup With(NOLOCK) Where SiebelCode = '" & mDmsSubCode & "' And SiebelCode <> '' And SiebelCode Is Not Null")
        
        If RsTemp.RecordCount > 0 Then
            AutomanSubcode = RsTemp!SubCode
        Else
            If mSubGroupCounter = 0 Then
                mSubGroupCounter = G_FaCn.Execute("Select SubGroupAcCode From SubGroupCounter With (NOLOCK)").Fields(0)
            Else
                mSubGroupCounter = mSubGroupCounter + 1
            End If
                        
            Set RsTemp = GCn.Execute("Select * From DmsSubGroup WITH (NOLOCK) Where DmsSubCode='" & mDmsSubCode & "'")
            If RsTemp.RecordCount > 0 Then
            
                Set rsTemp1 = GCn.Execute("Select AutomanSite From DmsSite WITH (NOLOCK) Where DmsDivision='" & RsTemp!division & "'")
                If rsTemp1.RecordCount > 0 Then
                    If AutomanSubcode = "" Then
                        mSubCode = XNull(rsTemp1!AutomanSite) & PubFirmCode & Format(mSubGroupCounter, "000000")
                    Else
                        mSubCode = AutomanSubcode
                    End If
                
                    
                    '-----Commented For Maching With Old DataImport----------
                    'If GCn.Execute("Select Count(*) From SubGroup Where Name='" & left(XNull(RsTemp!Name), 40) & "'").RecordCount > 0 Then
                        mname = left(XNull(RsTemp!Name) & " [" & mDmsSubCode & "]", 40)
                    'Else
                    '    mname = left(XNull(RsTemp!Name), 40)
                    'End If
                                    
                                    
                    If mLocalCentral = "" Then mLocalCentral = "L"
                    If XNull(RsTemp!City) <> "" Then
                        Set RsTempCity = GCn.Execute("Select CityCode From City WITH (NOLOCK) Where CityName='" & RsTemp!City & "' Or CityHelp='" & FilterString(RsTemp!City) & "'")
                        If RsTempCity.RecordCount > 0 Then
                            mCityCode = XNull(RsTempCity!CityCode)
                        Else
                            RsCity.MoveFirst
                            mCityCode = RsCity(0)
'                            If StrCmp(left(PubComp_Name, 5), "Ujwal") Then
'                                mCityCode = GCn.Execute("Select Max(" & cVal("CityCode") & ")+1 From City Where InStr(CityCode,'D')=0 And InStr(CityCode,'E')=0").Fields(0)
'                            Else
'                                mCityCode = GCn.Execute("Select Max(CityCode)+1 From City").Fields(0)
'                            End If
'                            mStateCode = AutomanStateCode(XNull(RsTemp!State))
'                            GCn.Execute "Insert Into City (CityCode, Site_Code, CityName, CityHelp, StateCode, " & _
'                                                          " LocalCentral, U_Name, U_EntDt, U_AE) " & _
'                                        " Values ('" & mCityCode & "', '" & PubSiteCode & "', '" & XNull(RsTemp!City) & "', '" & FilterString(RsTemp!City) & "', " & Val(mStateCode) & ", " & _
'                                                 "'" & mLocalCentral & "', 'CrmDms', " & ConvertDate(PubLoginDate) & ", 'A')"
'                            RsCity.Requery
                        End If
                    End If
                    
                    
                    
                    
                    mQry = "Insert Into SubGroup (AcId, Site_Code, SubCode, FirmCode, NamePrefix, " & _
                                                "Name, NameHelp, GroupCode, Nature, Add1, " & _
                                                "Add2,  CityCode, Phone, Mobile, Email, " & _
                                                "CstNo, LstNo, ActiveYn, U_Name, " & _
                                                "U_EntDt, U_AE, GroupNature, AliasYn, SiebelCode ) " & _
                         " Values ('" & mSubCode & "', " & PubSiteCode & ", '" & mSubCode & "', " & PubFirmCode & ", '', " & _
                         "'" & mname & "', '" & mname & "', '" & mAutomanGroupCode & "', '" & mNature & "', '" & left(XNull(RsTemp!Add1), 40) & "', " & _
                         "'" & left(XNull(RsTemp!Add2), 40) & "', '" & mCityCode & "', '" & XNull(RsTemp!Phone) & "', '', '" & XNull(RsTemp!EMail) & "', " & _
                         "'', '', 1, 'CrmDms', " & _
                         "" & ConvertDate(PubLoginDate) & ", 'A', 'A', 'N', '" & mDmsSubCode & "')"
                         
                         
                         
                    GCn.Execute mQry
                    If PubBackEnd = "A" Then G_FaCn.Execute mQry
                    
                    G_FaCn.Execute ("Update  SubGroupCounter Set SubGroupAcCode=" & mSubGroupCounter + 1 & " ")
                    GCn.Execute "Update DmsSubGroup Set AutomanSubCode='" & mSubCode & "' Where DmsSubCode='" & mDmsSubCode & "'"
                    
                    mIsAnySubCodeCreated = True
                    AutomanSubcode = mSubCode
                Else
                    CreateErrLog "Ledger Account", XNull(RsTemp!division), XNull(RsTemp!division) & " Site Not Find In DmsSite Table"
                End If
            End If
        End If
    'End If



    Set RsTemp = Nothing
    Set rsTemp1 = Nothing
End Function





Private Sub SelectFile()
    
    CD1.CancelError = False
    CD1.DialogTitle = "Select CrmDms Excel Files"
    CD1.Filter = "Excel Files (*.xls)|*.xls"
    CD1.FilterIndex = 1
    CD1.Flags = cdlOFNHideReadOnly
    CD1.ShowOpen
    
End Sub




Private Sub Ini_Grid(mFGrid As Control)
    Select Case UCase(mFGrid.Name)
        Case "FGRID"
            With FGrid
                .Cols = 9
                
                .TextMatrix(0, FSel) = "Select"
                .ColWidth(FSel) = 600
                .ColAlignment(FSel) = flexAlignCenterCenter
                
                .TextMatrix(0, fname) = "Name"
                .ColWidth(FSel) = 2000
                .ColAlignment(FSel) = flexAlignLeftCenter
                
                .TextMatrix(0, FFName) = "Fathers Name"
                .ColWidth(FSel) = 2000
                .ColAlignment(FSel) = flexAlignLeftCenter
                
                .TextMatrix(0, FAdd1) = "Address1"
                .ColWidth(FAdd1) = 2000
                .ColAlignment(FAdd1) = flexAlignLeftCenter
                
                .TextMatrix(0, FAdd2) = "Address2"
                .ColWidth(FAdd2) = 2000
                .ColAlignment(FAdd2) = flexAlignLeftCenter
                
                .TextMatrix(0, FAdd3) = "Address3"
                .ColWidth(FAdd3) = 2000
                .ColAlignment(FAdd3) = flexAlignLeftCenter
                
                .TextMatrix(0, FCity) = "City"
                .ColWidth(FCity) = 2000
                .ColAlignment(FCity) = flexAlignLeftCenter
                
                .ColWidth(FSubCode) = 0
            End With
            
        Case "FGRIDERR"
            With FgridErr
                .Cols = 4
                
                .ColWidth(0) = 400
                
                .TextMatrix(0, FErr_Cat) = "Category"
                .ColAlignment(FErr_Cat) = flexAlignLeftCenter
                .ColWidth(FErr_Cat) = 2000
                
                .TextMatrix(0, FErr_DmsRef) = "Reference"
                .ColAlignment(FErr_DmsRef) = flexAlignLeftCenter
                .ColWidth(FErr_DmsRef) = 2500
                
                .TextMatrix(0, FErr_Narration) = "Narration"
                .ColAlignment(FErr_Narration) = flexAlignLeftCenter
                .ColWidth(FErr_Narration) = 10000
            End With
            
        Case "FGRID1"
            With FGrid1
                .Cols = 4
                
                .TextMatrix(0, 0) = ""
                .ColWidth(0) = 400
                
                .TextMatrix(0, F1_BankAc) = "Bank A/c Name"
                .ColAlignment(F1_BankAc) = flexAlignLeftCenter
                .ColWidth(F1_BankAc) = 3000
                
                .TextMatrix(0, F1_DmsCode) = "Dms A/c Code"
                .ColAlignment(FErr_DmsRef) = flexAlignLeftCenter
                .ColWidth(FErr_DmsRef) = 1500
                                                
                .ColWidth(F1_BankAcCode) = 0
            End With
            
        Case "FGRID2"
            With FGrid2
                .Cols = 4
                
                .TextMatrix(0, 0) = ""
                .ColWidth(0) = 400
                
                .TextMatrix(0, F2_SupplierAc) = "Supplier A/c Name"
                .ColAlignment(F2_SupplierAc) = flexAlignLeftCenter
                .ColWidth(F2_SupplierAc) = 3000
                
                .TextMatrix(0, F1_DmsCode) = "Dms A/c Code"
                .ColAlignment(FErr_DmsRef) = flexAlignLeftCenter
                .ColWidth(FErr_DmsRef) = 1500
                                                
                .ColWidth(F2_SupplierAcCode) = 0
            End With
            
    End Select
End Sub



Private Function AutomanStateCode(DmsStateCode As String) As String
    If DmsStateCode = "MH" Then
        If RsState.RecordCount = 0 Then Exit Function
        RsState.MoveFirst
        RsState.FIND "StateName Like 'Maharastra'"
        If RsState.EOF = False Then
            AutomanStateCode = RsState!StateCode
        Else
            AutomanStateCode = ""
        End If
    End If
End Function



Sub MoveRec()
Dim RsTemp As ADODB.Recordset
Dim I As Integer
        
        With RsDmsEnviro
            GetTaxFormName XNull(!VehicleCentralPurchaseTaxForm), Txt(VehicleCentralPurchaseTaxForm)
            GetTaxFormName XNull(!VehicleLocalPurchaseTaxForm), Txt(VehicleLocalPurchaseTaxForm)
            GetTaxFormName XNull(!SpareCentralPurchaseTaxForm), Txt(SpareCentralPurchaseTaxForm)
            GetTaxFormName XNull(!SpareLocalPurchaseTaxForm), Txt(SpareLocalPurchaseTaxForm)
            GetTaxFormName XNull(!SpareCentralSaleTaxForm), Txt(SpareCentralSaleTaxForm)
            GetTaxFormName XNull(!SpareLocalSaleTaxForm), Txt(SpareLocalSaleTaxForm)
            GetMechanicName XNull(!DefaultMechanic), Txt(DefaultMechanic)
            GetLabourHeadName XNull(!DefaultLabourHead), Txt(DefaultLabourHead)
            GetPartNoHeadName XNull(!DefaultPartNo), Txt(DefaultPartNo)
            GetPartNoHeadName XNull(!DefaultOilPartNo), Txt(DefaultOilPartNo)
            GetSupervisorName XNull(!DefaultSupervisor), Txt(DefaultSupervisor)
        
            GetAdItemName XNull(!VehiclePurchaseDiscountItem), Txt(VehiclePurchaseDiscountItem)
            GetAdItemName XNull(!VehiclePurchaseTransportItem), Txt(VehiclePurchaseTransportItem)
        
            If UCase(XNull(!VehicleTaxOnDeliveryCharges)) = "Y" Then
                Txt(VehicleTaxOnDeliveryCharges) = "YES"
            Else
                Txt(VehicleTaxOnDeliveryCharges) = "NO"
            End If
        
            GetAcGroupName XNull(!SprDebtorGroupCode), Txt(SprDebtorGroupCode)
            GetAcGroupName XNull(!VehDebtorGroupCode), Txt(SprDebtorGroupCode)
            GetAcGroupName XNull(!VehDebtorGroupCode), Txt(VehDebtorGroupCode)
            GetAcGroupName XNull(!WsDebtorGroupCode), Txt(WsDebtorGroupCode)
            GetAcGroupName XNull(!SprCreditorGroupCode), Txt(SprCreditorGroupCode)
            GetAcGroupName XNull(!VehCreditorGroupCode), Txt(VehCreditorGroupCode)
            GetAcGroupName XNull(!VehPurGroupCode), Txt(VehPurGroupCode)
            GetAcGroupName XNull(!VehSaleGroupCode), Txt(VehSaleGroupCode)
            GetAcGroupName XNull(!SprPurGroupCode), Txt(SprPurGroupCode)
            GetAcGroupName XNull(!SprSaleGroupCode), Txt(SprSaleGroupCode)
            GetAcGroupName XNull(!VatGroupCode), Txt(VatGroupCode)
            GetAcGroupName XNull(!ServiceTaxGroupCode), Txt(ServiceTaxGroupCode)
            
            GetSubName XNull(!SprSaleAc), Txt(SprSaleAc)
            GetSubName XNull(!LubeSaleAc), Txt(LubSaleAc)
            GetSubName XNull(!SprSaleVat4Ac), Txt(SprSaleVat4Ac)
            GetSubName XNull(!SprPurchase4Ac), Txt(SprPurchase4Ac)
            GetSubName XNull(!VehSaleAc), Txt(VehSaleAc)
            GetSubName XNull(!SprCashAc), Txt(SprCashAc)
            GetSubName XNull(!VehCashAc), Txt(VehCashAc)
            GetSubName XNull(!WsCashAc), Txt(WsCashAc)
            GetSubName XNull(!SprBankAc), Txt(SprBankAc)
            GetSubName XNull(!VehBankAc), Txt(VehBankAc)
            GetSubName XNull(!WsBankAc), Txt(WsBankAc)
            Txt(LocalStateName) = XNull(!LocalStateName)
            GetSubName XNull(!SprPurchaseAc), Txt(SprPurchaseAc)
            GetSubName XNull(!SprCstPurchaseAc), Txt(SprCstPurchaseAc)
            GetSubName XNull(!VehPurchaseAc), Txt(VehPurchaseAc)
            GetSubName XNull(!VehCstPurchaseAc), Txt(VehCstPurchaseAc)
            GetSubName XNull(!LabourAc), Txt(LabourAc)
            GetSubName XNull(!ServTaxAc), Txt(ServTaxAc)
            GetSubName XNull(!CstAc), Txt(CstAc)
            GetSubName XNull(!VatAc), Txt(VatAc)
            GetSubName XNull(!Vat4Ac), Txt(Vat4Ac)
            GetSubName XNull(!VatInputAc), Txt(VatInputAc)
            GetSubName XNull(!Vat4InputAc), Txt(Vat4InputAc)
            GetSubName XNull(!ROffAc), Txt(ROffAc)
        End With
    
    Set RsTemp = GCn.Execute("Select * From DmsbankAc ")
    With FGrid1
        .Rows = 1
        If RsTemp.RecordCount > 0 Then
            For I = 1 To RsTemp.RecordCount
                .AddItem ""
                .TextMatrix(I, F1_BankAcCode) = XNull(RsTemp!AutomanBankCode)
                .TextMatrix(I, F1_DmsCode) = XNull(RsTemp!DmsBankCode)


                RsSubGroup.MoveFirst
                RsSubGroup.FIND "Code = '" & XNull(RsTemp!AutomanBankCode) & "'"
                If RsSubGroup.EOF = False Then .TextMatrix(I, F1_BankAc) = XNull(RsSubGroup!Name)


                RsTemp.MoveNext
            Next I
            .FixedRows = 1
        Else
            .AddItem ""
            .FixedRows = 1
        End If
    End With
    
    
    
    
    Set RsTemp = GCn.Execute("Select * From DmsSupplierAc ")
    With FGrid2
        .Rows = 1
        If RsTemp.RecordCount > 0 Then
            For I = 1 To RsTemp.RecordCount
                .AddItem ""
                .TextMatrix(I, F2_SupplierAcCode) = XNull(RsTemp!AutomanSupplierCode)
                .TextMatrix(I, F1_DmsCode) = XNull(RsTemp!DmsCode)


                RsSubGroup.MoveFirst
                RsSubGroup.FIND "Code = '" & XNull(RsTemp!AutomanSupplierCode) & "'"
                If RsSubGroup.EOF = False Then .TextMatrix(I, F2_SupplierAc) = XNull(RsSubGroup!Name)


                RsTemp.MoveNext
            Next I
            .FixedRows = 1
        Else
            .AddItem ""
            .FixedRows = 1
        End If
    End With
    
        
    Set RsTemp = GCn.Execute("Select Cat As Category, [Key] as Dms_Reference, Narration From DmsErrLog ")
    Set FgridErr.DataSource = RsTemp
    Ini_Grid FgridErr
    
End Sub



Sub GetSubName(mSubCode As String, mTxt As TextBox)
    With RsSubGroup
        .MoveFirst
        .FIND "Code = '" & mSubCode & "'"
        If .EOF = False Then
            mTxt = XNull(!Name)
            mTxt.Tag = XNull(!Code)
        End If
    End With
End Sub


Sub GetAcGroupName(mGrpCode As String, mTxt As TextBox)
    With RsAcGroup
        .MoveFirst
        .FIND "Code = '" & mGrpCode & "'"
        If .EOF = False Then
            mTxt = XNull(!Name)
            mTxt.Tag = XNull(!Code)
        End If
    End With
End Sub

Sub GetTaxFormName(mTaxFormCode As String, mTxt As TextBox)
    With RsTaxForm
        .MoveFirst
        .FIND "Code = '" & mTaxFormCode & "'"
        If .EOF = False Then
            mTxt = XNull(!Name)
            mTxt.Tag = XNull(!Code)
        End If
    End With
End Sub
Sub GetLabourHeadName(mCode As String, mTxt As TextBox)
    With RsLabour
        .MoveFirst
        .FIND "Code = '" & mCode & "'"
        If .EOF = False Then
            mTxt = XNull(!Name)
            mTxt.Tag = XNull(!Code)
        End If
    End With
End Sub
Sub GetPartNoHeadName(mCode As String, mTxt As TextBox)
    With RsLabour
        .MoveFirst
        .FIND "Code = '" & mCode & "'"
        If .EOF = False Then
            mTxt = XNull(!Name)
            mTxt.Tag = XNull(!Code)
        End If
    End With
End Sub


Sub GetMechanicName(mTaxFormCode As String, mTxt As TextBox)
    With RsMechanic
        .MoveFirst
        .FIND "Code = '" & mTaxFormCode & "'"
        If .EOF = False Then
            mTxt = XNull(!Name)
            mTxt.Tag = XNull(!Code)
        End If
    End With
End Sub

Sub GetSupervisorName(mTaxFormCode As String, mTxt As TextBox)
    With RsSupervisor
        .MoveFirst
        .FIND "Code = '" & mTaxFormCode & "'"
        If .EOF = False Then
            mTxt = XNull(!Name)
            mTxt.Tag = XNull(!Code)
        End If
    End With
End Sub


Sub GetAdItemName(mAdItemCode As String, mTxt As TextBox)
    With RsADItem
        .MoveFirst
        .FIND "Code = '" & mAdItemCode & "'"
        If .EOF = False Then
            mTxt = XNull(!Name)
            mTxt.Tag = XNull(!Code)
        End If
    End With
End Sub


Sub CreateErrLog(mCategory As String, mKeyValue As String, mNarration As String)
    GCn.Execute "Insert Into DmsErrLog(Cat, [Key], Narration, U_EntDt) Values('" & mCategory & "', '" & mKeyValue & "', '" & mNarration & "', " & ConvertDate(PubLoginDate) & ")"
End Sub



Private Function ChkFieldExist(Rs As ADODB.Recordset, mFieldName As String) As Boolean
Dim I As Integer
    For I = 0 To Rs.Fields.Count - 1
        If UCase(Trim(Rs.Fields(I).Name)) = UCase(Trim(mFieldName)) Or UCase(Trim(Rs.Fields(I).Name)) = UCase(Trim(Replace(mFieldName, ".", "#"))) Then
            ChkFieldExist = True
            Exit For
        End If
    Next I
    If ChkFieldExist = False Then MsgBox "<" & mFieldName & "> Field Not Found In Selected Excel File"
    
End Function





Private Sub FGrid1_Click()
    txtgrid1(0).Visible = False
End Sub

Private Sub FGrid2_Click()
    txtgrid2(0).Visible = False
End Sub


Private Sub FGrid1_DblClick()
    Select Case FGrid1.Col
        Case F1_BankAc, F1_DmsCode
            Call GridDblClick(Me, FGrid1, txtgrid1, 0)
    End Select
End Sub


Private Sub FGrid1_GotFocus()
    FGrid1.BackColorSel = FaBackColorSelEnter

    'FGrid1.Col = F1_BankAc
    CreateAcHelp "Bank"
    Set DgHelp.DataSource = RsHelp
    txtgrid1(0).Visible = False
End Sub

Private Sub FGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp And Val(FGrid1.Tag) = (FGrid1.Rows - (FGrid1.Rows - 1)) Then
        SendKeys "+{Tab}"
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid1.Tag = FGrid1.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGrid1.Col
            Case F1_DmsCode
                FGrid1 = ""
            Case F1_BankAc
                FGrid1 = ""
                FGrid1.TextMatrix(FGrid1.Row, F1_BankAcCode) = ""
                
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case FGrid1.Col
            Case F1_BankAc, F1_DmsCode
                Call GridDblClick(Me, FGrid1, txtgrid1, 0)
        End Select
    End If
    KeyCode = 0
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
Select Case FGrid1.Col
    Case F1_BankAc
       Call Get_Text(Me, FGrid1, txtgrid1, 0, False, KeyAscii)
    Case F1_DmsCode
        Call Get_Text(Me, FGrid1, txtgrid1, 0, False, KeyAscii)
End Select

End Sub

Private Sub FGrid1_LostFocus()
FGrid1.BackColorSel = FaCellBackColLeave1

FGrid1_Validate (True)
End Sub

Private Sub FGrid1_Scroll()
txtgrid1(0).Visible = False
DgHelp.Visible = False
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer

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
Exit Sub
End Sub

Private Sub FGrid1_Validate(Cancel As Boolean)
'    FGrid1.CellBackColor = CellBackColLeave
End Sub

Private Sub TxtGrid1_GotFocus(Index As Integer)
Ctrl_GetFocus txtgrid1(Index)
    Grid_Hide
    txtgrid1(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    
    Select Case FGrid1.Col
        Case F1_BankAc
            'CreateAcHelp "Bank"
            DgHelp.Move FGrid1.left, txtgrid1(0).top + txtgrid1(0).height + 20
            If RsHelp.RecordCount = 0 Or FGrid1.TextMatrix(FGrid1.Row, F1_BankAc) = "" Then Exit Sub
            RsHelp.MoveFirst
            RsHelp.FIND "Code ='" & FGrid1.TextMatrix(FGrid1.Row, F1_BankAcCode) & "'"
            If RsHelp.EOF = True Then RsHelp.MoveFirst
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
        Case F1_BankAc
            If DgHelp.Visible = False Then DGridColSwap DgHelp, 1
            DGridTxtKeyDown DgHelp, txtgrid1, Index, RsHelp, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DgHelp.Visible = False) Then
                If TxtGrid1Leave = True Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, True, F1_DmsCode, 1
                End If
            End If
        Case F1_DmsCode
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DgHelp.Visible = False) Then
                If TxtGrid1Leave = True Then
                    GridTxtDown FGrid1, txtgrid1, Index, KeyCode, True, F1_BankAc, 1
                End If
            End If
            
    End Select
End Sub

Private Sub txtgrid1_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
    Select Case FGrid1.Col
        Case F1_BankAc
            DGridTxtKeyPress txtgrid1, Index, RsHelp, KeyAscii, "Name"
    End Select
End Sub

Private Sub TxtGrid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        Select Case FGrid1.Col
            Case F1_BankAc
                If KeyCode <> 13 And DgHelp.Visible = False Then TxtGrid1_KeyDown Index, GridKey, 0
                DGridTxtKeyUp_Mast txtgrid1, Index, RsHelp, KeyCode, "Name"
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
    Case F1_BankAc
        If RsHelp.RecordCount = 0 Or txtgrid1(0).TEXT = "" Or RsHelp.EOF = True Or RsHelp.BOF = True Then
            FGrid1.TextMatrix(FGrid1.Row, F1_BankAc) = ""
            FGrid1.TextMatrix(FGrid1.Row, F1_BankAcCode) = ""
        Else
            FGrid1.TextMatrix(FGrid1.Row, 0) = FGrid1.Row
            FGrid1.TextMatrix(FGrid1.Row, F1_BankAc) = RsHelp!Name
            FGrid1.TextMatrix(FGrid1.Row, F1_BankAcCode) = RsHelp!Code
        End If
    Case F1_DmsCode
        FGrid1 = txtgrid1(0)
End Select
TxtGrid1Leave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid1.SetFocus
    txtgrid1(0).Visible = False
End If
End Function




'##########################################
Private Sub FGrid2_DblClick()
    Select Case FGrid2.Col
        Case F2_SupplierAc, F1_DmsCode
            Call GridDblClick(Me, FGrid2, txtgrid2, 0)
    End Select
End Sub


Private Sub FGrid2_GotFocus()
    FGrid2.BackColorSel = FaBackColorSelEnter

    'FGrid2.Col = F2_SupplierAc
    CreateAcHelp "Supplier"
    Set DgHelp.DataSource = RsHelp
    txtgrid2(0).Visible = False
End Sub

Private Sub FGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp And Val(FGrid2.Tag) = (FGrid2.Rows - (FGrid2.Rows - 1)) Then
        SendKeys "+{Tab}"
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid2.Tag = FGrid2.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGrid2.Col
            Case F1_DmsCode
                FGrid2 = ""
            Case F2_SupplierAc
                FGrid2 = ""
                FGrid2.TextMatrix(FGrid2.Row, F2_SupplierAcCode) = ""
                
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case FGrid2.Col
            Case F2_SupplierAc, F1_DmsCode
                Call GridDblClick(Me, FGrid2, txtgrid2, 0)
        End Select
    End If
    KeyCode = 0
End Sub

Private Sub FGrid2_KeyPress(KeyAscii As Integer)
Select Case FGrid2.Col
    Case F2_SupplierAc
       Call Get_Text(Me, FGrid2, txtgrid2, 0, False, KeyAscii)
    Case F1_DmsCode
        Call Get_Text(Me, FGrid2, txtgrid2, 0, False, KeyAscii)
End Select

End Sub

Private Sub FGrid2_LostFocus()
FGrid2.BackColorSel = FaCellBackColLeave1

FGrid2_Validate (True)
End Sub

Private Sub FGrid2_Scroll()
txtgrid2(0).Visible = False
DgHelp.Visible = False
End Sub

Private Sub FGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer

If FGrid2.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid2.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid2.Rows > 2 Then
                FGrid2.RemoveItem (FGrid2.Row)
            Else
                FGrid2.Rows = 1
                FGrid2.AddItem FGrid2.Rows
                FGrid2.FixedRows = 1
            End If
         End If
         For I = 1 To FGrid2.Rows - 1
            FGrid2.TextMatrix(I, 0) = I
         Next
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
   
FGrid2.SetFocus
End If
Exit Sub
End Sub

Private Sub FGrid2_Validate(Cancel As Boolean)
'    FGrid2.CellBackColor = CellBackColLeave
End Sub

Private Sub TxtGrid2_GotFocus(Index As Integer)
Ctrl_GetFocus txtgrid2(Index)
    Grid_Hide
    txtgrid2(0).Tag = FGrid2.TextMatrix(FGrid2.Row, FGrid2.Col)
    
    Select Case FGrid2.Col
        Case F2_SupplierAc
            'CreateAcHelp "Bank"
            DgHelp.Move FGrid2.left, txtgrid2(0).top + txtgrid2(0).height + 20
            If RsHelp.RecordCount = 0 Or FGrid2.TextMatrix(FGrid2.Row, F2_SupplierAc) = "" Then Exit Sub
            RsHelp.MoveFirst
            RsHelp.FIND "Code ='" & FGrid2.TextMatrix(FGrid2.Row, F2_SupplierAcCode) & "'"
            If RsHelp.EOF = True Then RsHelp.MoveFirst
    End Select
End Sub

Private Sub TxtGrid2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtgrid2(0).TEXT = txtgrid2(0).Tag
        TxtGrid2_KeyUp Index, KeyCode, Shift
        FGrid2.SetFocus
        txtgrid2(0).Visible = False
        Exit Sub
    End If
    Select Case FGrid2.Col
        Case F2_SupplierAc
            If DgHelp.Visible = False Then DGridColSwap DgHelp, 1
            DGridTxtKeyDown DgHelp, txtgrid2, Index, RsHelp, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DgHelp.Visible = False) Then
                If TxtGrid2Leave = True Then
                    GridTxtDown FGrid2, txtgrid2, Index, KeyCode, True, F1_DmsCode, 1
                End If
            End If
        Case F1_DmsCode
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DgHelp.Visible = False) Then
                If TxtGrid2Leave = True Then
                    GridTxtDown FGrid2, txtgrid2, Index, KeyCode, True, F2_SupplierAc, 1
                End If
            End If
            
    End Select
End Sub

Private Sub txtGrid2_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
    Select Case FGrid2.Col
        Case F2_SupplierAc
            DGridTxtKeyPress txtgrid2, Index, RsHelp, KeyAscii, "Name"
    End Select
End Sub

Private Sub TxtGrid2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        Select Case FGrid2.Col
            Case F2_SupplierAc
                If KeyCode <> 13 And DgHelp.Visible = False Then TxtGrid2_KeyDown Index, GridKey, 0
                DGridTxtKeyUp_Mast txtgrid2, Index, RsHelp, KeyCode, "Name"
        End Select
End Sub

Private Sub TxtGrid2_LostFocus(Index As Integer)
    'If ExitCtrl = False Then Exit Sub
    txtgrid2(Index).Visible = False
End Sub

Private Sub TxtGrid2_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGrid2Leave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Function TxtGrid2Leave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Dim Repeat$
Select Case FGrid2.Col
    Case F2_SupplierAc
        If RsHelp.RecordCount = 0 Or txtgrid2(0).TEXT = "" Or RsHelp.EOF = True Or RsHelp.BOF = True Then
            FGrid2.TextMatrix(FGrid2.Row, F2_SupplierAc) = ""
            FGrid2.TextMatrix(FGrid2.Row, F2_SupplierAcCode) = ""
        Else
            FGrid2.TextMatrix(FGrid2.Row, 0) = FGrid2.Row
            FGrid2.TextMatrix(FGrid2.Row, F2_SupplierAc) = RsHelp!Name
            FGrid2.TextMatrix(FGrid2.Row, F2_SupplierAcCode) = RsHelp!Code
        End If
    Case F1_DmsCode
        FGrid2 = txtgrid2(0)
End Select
TxtGrid2Leave = True
'Important at the time of validating  a control if you are making the visibility of
'control false forcefully it will generate error
If ValidateCall = False Then
    FGrid2.SetFocus
    txtgrid2(0).Visible = False
End If
End Function



Sub Grid_Hide()
    If DgHelp.Visible = True Then DgHelp.Visible = False
End Sub


Sub BlankAll()
    BlankText Me
End Sub

Private Sub FImportSparePurchase()
Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer, mV_Type As String
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String, mOrderQty As Double, mPhysicalQty As Double
Dim mPrefix As String, mname As String, mLubType As String, mTrnType As String, mDebitAc As String, mFormCode As String
Dim mChallanNo As String, mHeaderParty As String
Dim mQty As Double, mCount As Integer, mAmount As Double
Dim mInvoiceNo As String, mChallanID As String
Dim mTax_Amt As Double, mLocal As String
Dim mFileName As String, mLineFileName As String
Dim mFileTitle As String, mLineFileTitle As String
Dim mVouCat As String
Dim Master1 As New ADODB.Recordset
Dim mCashCredit As String
Dim mGodown As String
Dim mQry As String
Dim mSrl As Integer
Dim mTrans As Boolean
Dim mVTypeGR As String
Dim mVNoGr As String
Dim CodeCnt As Long

'On Error GoTo ELoop
    
    Call SelectFile
    mFileName = CD1.FileName
    mFileTitle = CD1.FileTitle
    If mFileName = "" Then Exit Sub
    mFileTitle = mID(mFileTitle, 1, Len(mFileTitle) - 4)
    mGodown = GCn.Execute("Select SprWorksGodown from Syctrl").Fields(0).Value
    
    Set DmsConn = New Connection
    DmsConn.CursorLocation = adUseClient
    DmsConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mFileName & ";Extended Properties=Excel 8.0"
    
    Set RsDms = DmsConn.Execute("Select * from [" & mFileTitle & "$]")
    
    
    If ChkFieldExist(RsDms, "Invoice #") And ChkFieldExist(RsDms, "Invoice_Date") And _
       ChkFieldExist(RsDms, "Vendor Name") And ChkFieldExist(RsDms, "Division") Then
    End If
    
    
    GCn.BeginTrans
    mTrans = True
    
    
    If RsDms.RecordCount > 0 Then RsDms.MoveFirst
    
    
    mVTypeGR = "SXGR"
    mV_Type = "SXPIR"
    mVouCat = "Spare Purchase"
    
    If mV_Type = "SXPIR" Then
        mCashCredit = "Credit"
    Else
        mCashCredit = "Cash"
    End If
    
    
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from Sp_Purch where Left(DocID,1)='" & PubDivCode & "' and " & cMID("DocID", "2", "1") & "='" & PubSiteCode & "' and V_Type='" & mV_Type & "'").Fields(0).Value
    mVNoGr = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from Sp_Purch where Left(DocID,1)='" & PubDivCode & "' and " & cMID("DocID", "2", "1") & "='" & PubSiteCode & "' and V_Type='" & mVTypeGR & "'").Fields(0).Value
 
    
    Do Until RsDms.EOF
        If XNull(StringPass(RsDms.Fields("Invoice #"))) = "" Then ErrorCnt = 1
                
        mInvoiceNo = StringPass(RsDms.Fields("Invoice #").Value)
        GCn.Execute ("Delete from DmsErrLog Where [Key] = '" & XNull(RsDms.Fields("Invoice #")) & "'")
     
             
        If XNull(RsDms.Fields("Invoice_Date")) = "" Then
            CreateErrLog mVouCat, mInvoiceNo, " Invoice Date is Blank "
            ErrorCnt = 1
        End If
     
                                     
        mHeaderParty = XNull(StringPass(RsDms.Fields("Vendor Name")))
        If XNull(StringPass(RsDms.Fields("Vendor Name"))) = "" Then
            CreateErrLog mVouCat, mInvoiceNo, " Vendor Name not found for Invoice No : " & mInvoiceNo
            ErrorCnt = 1
        Else
            If GCn.Execute("Select AutomanSupplierCode as SubCode From DmsSupplierAc With(NOLOCK) Where DmsCode ='" & StringPass(RsDms.Fields("Vendor Name")) & "'").RecordCount > 0 Then
                mPartyCode = GCn.Execute("Select AutomanSupplierCode as SubCode From DmsSupplierAc With(NOLOCK) Where DmsCode ='" & StringPass(RsDms.Fields("Vendor Name")) & "'").Fields(0).Value
            Else
                Call CreateErrLog(mVouCat, mInvoiceNo, "Vendor Name - " & XNull(RsDms.Fields("Vendor Name")) & " Not Found In Automan")
                ErrorCnt = 1
            End If
            
        End If
           
               
        If GCn.Execute("Select V_no from SP_Purch where Party_Doc_No='" & StringPass(RsDms.Fields("Invoice #")) & "' and v_Type='" & mV_Type & "' and Party_Code='" & mPartyCode & "'").RecordCount > 0 Then
            Call CreateErrLog(mVouCat, mInvoiceNo, " Invoice No : " & mInvoiceNo & " Already exist in Automan ")
            ErrorCnt = 1
        End If
        
        If XNull(StringPass(RsDms.Fields("Division"))) = "" Then
            Call CreateErrLog(mVouCat, mInvoiceNo, "Division Name Field is blank in Excel File")
            ErrorCnt = 1
        Else
            If GCn.Execute("select * from DmsSite where DmsDivision='" & StringPass(RsDms.Fields("Division")) & "'").RecordCount > 0 Then
                mRecordSite = GCn.Execute("select AutomanSite from DmsSite where DmsDivision='" & StringPass(RsDms.Fields("Division")) & "'").Fields(0).Value
                mRecordDiv = GCn.Execute("select AutomanDivision from DmsSite where DmsDivision='" & StringPass(RsDms.Fields("Division")) & "'").Fields(0).Value
            Else
                Call CreateErrLog(mVouCat, mInvoiceNo, "Division Name in not defined in Automan")
                ErrorCnt = 1
            End If
        End If
            
        
        
        mPrefix = "SBL" & Format(RsDms.Fields("Invoice_Date"), "yy")
        CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from SP_Purch With (NoLock) where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mV_Type & "'").Fields(0).Value
        mVNoGr = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from SP_Purch  With (NoLock) where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mVTypeGR & "'").Fields(0).Value
        mDocId = mRecordDiv & mRecordSite & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
        mChallanID = mRecordDiv & mRecordSite & mRecordSite & mVTypeGR & mPrefix & Right("00000000" & mVNoGr, 8)
        mLubType = GCn.Execute("Select PartGrade_Lub from Syctrl").Fields(0).Value
        
        
        If mTrnType = mLubType Then
            mFormCode = GCn.Execute("Select SparePurchFormLubs from Enviro").Fields(0).Value
            mTax_Amt = Val(RsDms.Fields("Total_Tax_Amount"))
            mLocal = "L"
        Else
            If Not IsNull(RsDms.Fields("Total_Tax_Amount")) Then
                If Val(RsDms.Fields("Total_Tax_Amount")) > 0 Then
                    mFormCode = GCn.Execute("Select SpareLocalPurchaseTaxForm from DmsEnviro").Fields(0).Value
                    mTax_Amt = Val(RsDms.Fields("Total_Tax_Amount"))
                    mLocal = "L"
                End If
            End If
            
            If Not IsNull(RsDms.Fields("LST")) Then
                If Val(RsDms.Fields("LST")) > 0 Then
                    mFormCode = GCn.Execute("Select SpareLocalPurchaseTaxForm from DmsEnviro").Fields(0).Value
                    mTax_Amt = Val(RsDms.Fields("Total_Tax_Amount"))
                    mLocal = "L"
                End If
            End If
            
            If Not IsNull(RsDms.Fields("CST")) Then
                If Val(RsDms.Fields("CST")) > 0 Then
                    mFormCode = GCn.Execute("Select SpareCentralPurchaseTaxForm from DmsEnviro").Fields(0).Value
                    mTax_Amt = Val(RsDms.Fields("Total_Tax_Amount"))
                    mLocal = "C"
                End If
            End If
            
        End If
        If mFormCode = "" Then
            mFormCode = GCn.Execute("Select SpareLocalPurchaseTaxForm from DmsEnviro").Fields(0).Value
            mLocal = "L"
        End If
        
        If mRecordDiv <> "" Then
            mDebitAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc where Div_Code='" & mRecordDiv & "' and Form_Code='" & mFormCode & "'").Fields(0).Value
        End If
        
        mQty = 0
        mCount = 0
        mAmount = 0
        mTax_Amt = 0
        
        
        mAmount = VNull(RsDms.Fields("Net_Amount"))
        mChallanNo = XNull(RsDms.Fields("Invoice #"))
        
        
        If ErrorCnt = 0 Then
        
            mQry = "INSERT INTO dbo.SP_Purch "
            mQry = mQry & "(DocID, DocIDHelp, V_Type, V_No, Site_Code, "
            mQry = mQry & "V_Date, Cash_Credit, Party_Code, Party_Name, L_C, "
            mQry = mQry & "Form_Code,Party_Doc_No, Party_Doc_Date, Tot_No_of_Items, Tot_Doc_Qty, "
            mQry = mQry & "Tot_Phy_Qty, SprAmt_MRP_TB, SprAmt_MRP_TP, OilAmt_MRP_TB, OilAmt_MRP_TP, "
            mQry = mQry & "SprAmt_TB, SprAmt_TP, OilAmt_TB, OilAmt_TP, OilAmt, "
            mQry = mQry & "SprAmt, Tot_Amt, Tot_Disc_Amt, Tot_Ord_DiscAmt, Tot_Goods_Value, "
            mQry = mQry & "Tax_Amt, Addition, Deduction, NET_AMT, EntryTaxPer, "
            mQry = mQry & "EntryTaxAmt, Remarks, AcPsoting_YN, DrAc_Code, U_Name, "
            mQry = mQry & "U_EntDt, U_AE, Transportation, SiebelDocID, Sat_Yn, Invoice_DocID, "
            mQry = mQry & "SatAmt, AddBy, AddDate) "
            mQry = mQry & "VALUES ("
            mQry = mQry & "'" & mChallanID & "', '" & mChallanID & "', '" & Trim(mVTypeGR) & "', '" & mVNoGr & "', '" & mRecordSite & mRecordSite & "', "
            mQry = mQry & "" & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", '" & mCashCredit & "', '" & mPartyCode & "', '" & XNull(RsDms.Fields("Vendor Name")) & "', '" & mLocal & "', "
            mQry = mQry & "'" & mFormCode & "','" & XNull(RsDms.Fields("Invoice #")) & "', " & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", " & mCount & ", " & mQty & ", "
            mQry = mQry & "" & mQty & ", 0, 0, 0, 0, "
            mQry = mQry & "0, 0, 0, 0, " & mAmount & ", "
            mQry = mQry & "" & mAmount & ", " & mAmount & ", 0, 0, " & mAmount & ", "
            mQry = mQry & "" & VNull(RsDms.Fields("Total_Tax_Amount")) & ", 0, 0, " & VNull(RsDms.Fields("Total_Tax_Amount")) + mAmount & " , 0, "
            mQry = mQry & "0, '" & XNull(RsDms.Fields("Narration")) & "', 0, '" & mDebitAc & "', '" & pubUName & "', "
            mQry = mQry & "'" & PubLoginDate & "', 'A', 0, '" & XNull(RsDms.Fields("Invoice #")) & "', 0, '" & mDocId & "', "
            mQry = mQry & "0, '" & pubUName & "','" & PubLoginDate & "') "
    
            GCn.Execute mQry
        
        
            mQry = "INSERT INTO dbo.SP_Purch "
            mQry = mQry & "(DocID, DocIDHelp, V_Type, V_No, Site_Code, "
            mQry = mQry & "V_Date, Cash_Credit, Party_Code, Party_Name, L_C, "
            mQry = mQry & "Form_Code,Party_Doc_No, Party_Doc_Date, Tot_No_of_Items, Tot_Doc_Qty, "
            mQry = mQry & "Tot_Phy_Qty, SprAmt_MRP_TB, SprAmt_MRP_TP, OilAmt_MRP_TB, OilAmt_MRP_TP, "
            mQry = mQry & "SprAmt_TB, SprAmt_TP, OilAmt_TB, OilAmt_TP, OilAmt, "
            mQry = mQry & "SprAmt, Tot_Amt, Tot_Disc_Amt, Tot_Ord_DiscAmt, Tot_Goods_Value, "
            mQry = mQry & "Tax_Amt, Addition, Deduction, NET_AMT, EntryTaxPer, "
            mQry = mQry & "EntryTaxAmt, Remarks, AcPsoting_YN, DrAc_Code, U_Name, "
            mQry = mQry & "U_EntDt, U_AE, Transportation, SiebelDocID, Sat_Yn, "
            mQry = mQry & "SatAmt, AddBy, AddDate) "
            mQry = mQry & "VALUES ("
            mQry = mQry & "'" & mDocId & "', '" & mDocId & "', '" & Trim(mV_Type) & "', '" & CodeCnt & "', '" & mRecordSite & mRecordSite & "', "
            mQry = mQry & "" & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", '" & mCashCredit & "', '" & mPartyCode & "', '" & XNull(RsDms.Fields("Vendor Name")) & "', '" & mLocal & "', "
            mQry = mQry & "'" & mFormCode & "','" & XNull(RsDms.Fields("Invoice #")) & "', " & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", " & mCount & ", " & mQty & ", "
            mQry = mQry & "" & mQty & ", 0, 0, 0, 0, "
            mQry = mQry & "0, 0, 0, 0, " & mAmount & ", "
            mQry = mQry & "" & mAmount & ", " & mAmount & ", 0, 0, " & mAmount & ", "
            mQry = mQry & "" & VNull(RsDms.Fields("Total_Tax_Amount")) & ", 0, 0, " & VNull(RsDms.Fields("Total_Tax_Amount")) + mAmount & " , 0, "
            mQry = mQry & "0, '" & XNull(RsDms.Fields("Narration")) & "', 0, '" & mDebitAc & "', '" & pubUName & "', "
            mQry = mQry & "'" & PubLoginDate & "', 'A', 0, '" & XNull(RsDms.Fields("Invoice #")) & "', 0, "
            mQry = mQry & "0, '" & pubUName & "','" & PubLoginDate & "') "
    
            GCn.Execute mQry


                        
            mQry = "Insert Into Sp_Stock (DocID, Invoice_DocID, Site_Code, V_Type, V_No, V_Date, "
            mQry = mQry & "Party_Code, Srl_No, L_C, Remark, Part_No, "
            mQry = mQry & "Godown, Qty_Doc, Qty_Rec, Tax_Yn, MRP_YN, "
            mQry = mQry & "TaxAmt, TaxPer, Amount, Net_Amt, Rate, "
            mQry = mQry & "V_Rate, Part_SrlNo, Disc_Per, Disc_Amt, Ord_DiscPer, "
            mQry = mQry & "Ord_DiscAmt, U_Name, U_EntDt, U_AE) "
            mQry = mQry & "Values ( "
            mQry = mQry & "'" & mChallanID & "','" & mDocId & "', '" & mRecordSite & mRecordSite & "', '" & Trim(mVTypeGR) & "', " & mVNoGr & ", " & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", "
            mQry = mQry & "'" & mPartyCode & "', " & mSrl & ", '" & mLocal & " ', '" & mChallanNo & "', '" & XNull(RsDmsEnviro.Fields("DefaultPartNo")) & "', "
            mQry = mQry & "'" & mGodown & "', " & 1 & ", " & 1 & ", 1, " & IIf(mRecordDiv = "C", 1, 0) & ", "
            mQry = mQry & "" & VNull(RsDms.Fields("Total_Tax_Amount")) & ", " & VNull(RsDms.Fields("Total_Tax_Amount")) * 100 / mAmount & ", " & VNull(RsDms.Fields("Total_Tax_Amount")) + mAmount & ", " & mAmount & ", " & (VNull(RsDms.Fields("Total_Tax_Amount")) + mAmount) & ", "
            mQry = mQry & "" & mAmount & ", " & mSrl & ", 0, 0, 0, "
            mQry = mQry & "0, '" & pubUName & "', " & ConvertDate(PubLoginDate) & ", 'A') "
            GCn.Execute mQry
            
            If GCn.Execute("Select DocID From Sp_Purch With (NoLock) where Party_Doc_No='" & mChallanNo & "' and Party_Code='" & mPartyCode & "' and V_Type='SXGR'").RecordCount > 0 Then     '
                mChallanID = GCn.Execute("Select DocID From Sp_Purch where Party_Doc_No='" & mChallanNo & "' and Party_Code='" & mPartyCode & "' and V_Type='SXGR'").Fields(0).Value
                GCn.Execute ("Update Sp_Purch set Invoice_DocID='" & mDocId & "' where DocID='" & mChallanID & "'")
                GCn.Execute ("Update Sp_Stock set Invoice_DocID='" & mDocId & "'," & _
                             "v_Date2=" & ConvertDate(Format(left(RsDms.Fields("Invoice_Date"), 10), "dd/MMM/yyyy")) & _
                             ", Rate2=Rate, Amount2 =Amount,Net_Amt2=Net_Amt where DocID='" & mChallanID & "'")
            End If
            
                
        CodeCnt = CodeCnt + 1
        End If
        RsDms.MoveNext
        ErrorCnt = 0
    Loop
    GCn.CommitTrans
    mTrans = False
    MsgBox "Spare Detail Imported Successfully"
    
lblExit:
    Set RsNew = Nothing
    Exit Sub
ELoop:
    MsgBox err.Description
    If mTrans Then GCn.RollbackTrans
End Sub



Private Sub FImportSpareSale()
Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer, mV_Type As String, mChallanType As String, mHeaderAcCode  As String
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String, mOrderQty As Double, mPhysicalQty As Double
Dim mPrefix As String, mname As String, mLubType As String, mTrnType As String, mDebitAc As String, mFormCode As String
Dim mChallanNo As String, mHeaderParty As String
Dim mQty As Double, mCount As Integer, mAmount As Double
Dim mInvoiceNo As String, mChallanID As String
Dim mTax_Amt As Double, mTax_Amt1 As Double, mLocal As String
Dim mFileName As String, mLineFileName As String
Dim mFileTitle As String, mLineFileTitle As String
Dim mVouCat As String, mGatePassID As String
Dim Master1 As New ADODB.Recordset
Dim mCashCredit As String
Dim mGodown As String
Dim mQry As String
Dim mSrl As Integer
Dim mTrans As Boolean
Dim mVTypeGR As String
Dim mVNoGr As String
Dim CodeCnt As Long
Dim mSpareGodown As String
Dim mTaxPer As Double
Dim mRate As Double
Dim mMRP As Double
Dim mEditFlag As Boolean
Dim IsLineFileFound As Boolean
Dim mTaxAmt As Double
Dim mDisAmt As Double
Dim mDisPer As Double


Dim mSprAmt_MRP_TB As Double, mSprAmt_MRP_TP As Double, mOilAmt_MRP_TB As Double, mOilAmt_MRP_TP As Double
Dim mD_Per_MRP_TB As Double, mD_Per_MRP_TP As Double, mD_Amt_MRP_TB As Double, mD_Amt_MRP_TP As Double
        
Dim mSprAmt_TB As Double, mSprAmt_TP As Double, mOilAmt_TB As Double, mOilAmt_TP As Double
Dim mD_Per_TB As Double, mD_Per_TP As Double, mD_Amt_TB As Double, mD_Amt_TP As Double


'On Error GoTo ELoop
    
    Call SelectFile
    mFileName = CD1.FileName
    mFileTitle = CD1.FileTitle
    If mFileName = "" Then Exit Sub
    Call SelectFile
    mLineFileName = CD1.FileName
    mLineFileTitle = CD1.FileTitle
    If mLineFileName = "" Then Exit Sub
    mFileTitle = mID(mFileTitle, 1, Len(mFileTitle) - 4)
    mLineFileTitle = mID(mLineFileTitle, 1, Len(mLineFileTitle) - 4)
    mGodown = GCn.Execute("Select SprWorksGodown from Syctrl").Fields(0).Value
    Set DmsConn = New Connection
    DmsConn.CursorLocation = adUseClient
    DmsConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mFileName & ";Extended Properties=Excel 8.0"
    
    Set RsDms = DmsConn.Execute("Select * from [" & mFileTitle & "$]  Where Invoice_Status='New' ")
    
    Set ExcelGcn2 = New Connection
    ExcelGcn2.CursorLocation = adUseClient
    ExcelGcn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mLineFileName & ";Extended Properties=Excel 8.0"
    
    
    
    
    
    If RsDms.RecordCount > 0 Then RsDms.MoveFirst
    
    
    
    mVouCat = "Spare Sales"
    
    
    
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from Sp_Purch where Left(DocID,1)='" & PubDivCode & "' and " & cMID("DocID", "2", "1") & "='" & PubSiteCode & "' and V_Type='" & mV_Type & "'").Fields(0).Value
    mVNoGr = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from Sp_Purch where Left(DocID,1)='" & PubDivCode & "' and " & cMID("DocID", "2", "1") & "='" & PubSiteCode & "' and V_Type='" & mVTypeGR & "'").Fields(0).Value
 
    
    GCn.BeginTrans
    mTrans = True
    
    Do Until RsDms.EOF
        IsLineFileFound = True
        If RsDms!Invoice_Status <> "New" Then ErrorCnt = 1
        
        If IsNull(StringPass(RsDms.Fields("Invoice_No"))) Or StringPass(RsDms.Fields("Invoice_No")) = "" Then ErrorCnt = 1
        
        mInvoiceNo = StringPass(RsDms.Fields("Invoice_No").Value)
        mChallanNo = StringPass(RsDms.Fields("Order_No").Value)
                
        If IsNull(StringPass(RsDms.Fields("Division"))) Or StringPass(RsDms.Fields("Division")) = "" Then
            CreateErrLog mVouCat, mInvoiceNo, " Division Name field is Empty "
            ErrorCnt = 1
        Else
            If GCn.Execute("select * from DmsSite where DmsDivision='" & StringPass(RsDms.Fields("Division")) & "'").RecordCount > 0 Then
                mRecordSite = GCn.Execute("select AutomanSite from DmsSite where DmsDivision='" & StringPass(RsDms.Fields("Division")) & "'").Fields(0).Value
                mRecordDiv = GCn.Execute("select AutomanDivision from DmsSite where DmsDivision='" & StringPass(RsDms.Fields("Division")) & "'").Fields(0).Value
            Else
                Call CreateErrLog(mVouCat, mInvoiceNo, "Division Name in not defined in Automan")
                ErrorCnt = 1
            End If
        End If
                
        mChallanType = "SYSC"
        If RsDms.Fields("Mode Of Payment") = "CASH" Then
            mV_Type = "SYSIC"
        Else
            mV_Type = "SYSIR"
        End If
        
        
        
        
        
        If IsNull(StringPass(RsDms.Fields("Account_Code"))) Or StringPass(RsDms.Fields("Account_Code")) = "" Then
            If IsNull(StringPass(RsDms.Fields("Customer_Code"))) Or StringPass(RsDms.Fields("Customer_Code")) = "" Then
                Call CreateErrLog(mVouCat, mInvoiceNo, "Account/Customer Code field is Empty")
                ErrorCnt = 1
            Else
                mHeaderParty = RsDms.Fields("Full Name")
                mHeaderAcCode = RsDms.Fields("Customer_Code")
            End If
        Else
            mHeaderAcCode = RsDms.Fields("Account_Code")
        End If
        
        If RsDms.Fields("Mode Of Payment") = "CASH" Then
             mPartyCode = XNull(RsDmsEnviro.Fields("SprCashAc"))
        Else
        
            With RsDms
            If GCn.Execute("Select Count(*) From DmsSubGroup Where DmsSubCode='" & XNull(RsDms.Fields("Account_Code")) & "'").Fields(0).Value = 0 Then
                    Set RsTemp = GCn.Execute("Select AutomanSite From DmsSite Where DmsDivision='" & XNull(.Fields("Division")) & "'")
                    If RsTemp.RecordCount > 0 Then
                        GCn.Execute "Delete From DmsErrLog Where [Key] = '" & XNull(.Fields("Customer_Code")) & "' "
                        GCn.Execute "Insert Into DmsSubGroup(DmsSubCode, Name,[Group], Division) " & _
                                    " Values ('" & IIf(XNull(.Fields("Account_Code")) <> "", XNull(.Fields("Account_Code")), XNull(.Fields("Customer_Code"))) & "', " & _
                                    "'" & left(IIf(XNull(.Fields("Account_Name")) <> "", XNull(.Fields("Account_Name")), XNull(.Fields("Full Name"))), 50) & "','Sundry Debtors', '" & XNull(.Fields("Division")) & "')"
                    Else
                        CreateErrLog mVouCat, XNull(.Fields("Account_Code")), XNull(.Fields("Division")) & " Not Defined In DmsDivision Table"
                    End If
            End If
        
            mPartyCode = AutomanSubcode(XNull(.Fields("Account_Code")), RsDmsEnviro!SprDebtorGroupCode, "Customer")
            End With

        
        
'            If GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & mHeaderAcCode & "'").RecordCount = 0 Then
'                Call CreateErrLog(mVouCat, mInvoiceNo, "Account/Customer Code not found in Ledger Account RsDms of Automan")
'                ErrorCnt = 1
'            Else
'                mPartyCode = GCn.Execute("Select SubCode from SubGroup where siebelCode='" & mHeaderAcCode & "'").Fields(0).Value
'            End If
        End If
        
        If GCn.Execute("Select V_no from SP_Sale where SiebelDocID='" & mInvoiceNo & "' and V_Type='" & mV_Type & "'").RecordCount > 0 Then
            'GoTo DuplicateSkipped
            mEditFlag = True
        End If
            
       
        If IsNull(RsDms.Fields("Invoice_Date").Value) Or StringPass(RsDms.Fields("Invoice_Date").Value) = "" Then
            Call CreateErrLog(mVouCat, mInvoiceNo, "Invoice Date field is Empty")
            ErrorCnt = 1
        End If
        
        Dim mShortYear As String
        If Month(RsDms.Fields("Invoice_Date")) > 3 Then
            mShortYear = Right(Format(RsDms.Fields("Invoice_Date"), "yy"), 1) & Right(Val(Format(RsDms.Fields("Invoice_Date"), "yy")) + 1, 1)
        Else
            mShortYear = Right(Val(Format(RsDms.Fields("Invoice_Date"), "yy")) - 1, 1) & Right(Format(RsDms.Fields("Invoice_Date"), "yy"), 1)
        End If
        mPrefix = "SBL" & mShortYear 'Format(RsDms.Fields("Receipt_Date"), "yy")
        
        '' For Invoice Details :
        CodeCnt = Right(mInvoiceNo, 5) ''GCn.Execute("Select iif(isnull(Max(V_No)),0,Max(V_no))+1 from SP_Sale where Left(DocID,1)='" & mRecordDiv & "' and mid(DocID,2,2)='" & mRecordSite & mRecordSite & "' and V_Type='" & mV_Type & "'").Fields(0).Value
        If mEditFlag Then
            mDocId = GCn.Execute("Select DocId from SP_Sale where SiebelDocID='" & mInvoiceNo & "' and V_Type='" & mV_Type & "'").Fields(0)
        Else
            mDocId = mRecordDiv & mRecordSite & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
        End If
        
        '' For Challan Details :
        CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from SP_Sale With (NoLock) where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mChallanType & "'").Fields(0).Value
        mChallanNo = CodeCnt
        If mEditFlag Then
            mChallanID = GCn.Execute("Select DocId From Sp_Sale Where Invoice_DocID='" & mDocId & "'").Fields(0).Value
            mChallanNo = Val(Right(mChallanID, 8))
        Else
            mChallanID = mRecordDiv & mRecordSite & mRecordSite & " " & mChallanType & mPrefix & Right("00000000" & CodeCnt, 8)
        End If
        
        '' For GatePass Details :
        'CodeCnt = GCn.Execute("Select iif(isnull(Max(val(Left(GP_No,5)))),0,Max(val(Left(GP_no,5))))+1 from SP_Sale where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mChallanType & "'").Fields(0).Value
        CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal("Right(GP_No,5)") & ")", "0") & "+1 from SP_Sale where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mChallanType & "' and gp_no <>''").Fields(0).Value
        mGatePassID = mRecordDiv & mRecordSite & mRecordSite & Right("00000" & CodeCnt, 5)
        
        mLubType = GCn.Execute("Select PartGrade_Lub from Syctrl").Fields(0).Value
        
        
        If Not IsNull(RsDms.Fields("Total_Tax_Amount")) Then
            If eVal(RsDms.Fields("Total_Tax_Amount")) > 0 Then
                mFormCode = XNull(RsDmsEnviro.Fields("SpareLocalSaleTaxForm"))
                mTax_Amt = eVal(RsDms.Fields("Total_Tax_Amount"))
                mLocal = "L"
            End If
        End If
            If Not IsNull(RsDms.Fields("LST")) Then
                If eVal(RsDms.Fields("LST")) > 0 Then
                    mFormCode = XNull(RsDmsEnviro.Fields("SpareLocalSaleTaxForm"))
                    mTax_Amt = eVal(RsDms.Fields("Total_Tax_Amount"))
                    mLocal = "L"
                End If
        End If
        
        If mFormCode = "" Then
            mFormCode = XNull(RsDmsEnviro.Fields("SpareLocalSaleTaxForm"))
            mTax_Amt = eVal(RsDms.Fields("Total_Tax_Amount"))
            mLocal = "L"
        End If
        
        mDebitAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc where Div_Code='" & mRecordDiv & "' and Form_Code='" & mFormCode & "'").Fields(0).Value
        
                
        Set Master1 = CreateObject("ADODB.Recordset")
        GSQL = "Select * FROM [" & mLineFileTitle & "$] where Invoice_No='" & mInvoiceNo & "' And Invoice_Status='New' Order By  Invoice_No"
        Set Master1 = ExcelGcn2.Execute(GSQL)
        
        'GSQL = "Select * FROM [" & mLineFileName & "$] "
'        Master1.Open GSQL, ExcelGcn2, adOpenStatic
        GCn.Execute "Delete From Sp_Stock Where DocId='" & mChallanID & "'"
        
        If Master1.RecordCount = 0 Then
            Call CreateErrLog(mVouCat, mInvoiceNo, "Line detail not found in excel file")
            IsLineFileFound = False
            'ErrorCnt = 1
        End If
        
        mSrl = 1
        mAmount = 0
        mTax_Amt1 = 0
                        
        mSprAmt_MRP_TB = 0: mSprAmt_MRP_TP = 0: mOilAmt_MRP_TB = 0: mOilAmt_MRP_TP = 0
        mD_Per_MRP_TB = 0: mD_Per_MRP_TP = 0: mD_Amt_MRP_TB = 0: mD_Amt_MRP_TP = 0
        
        mSprAmt_TB = 0: mSprAmt_TP = 0: mOilAmt_TB = 0: mOilAmt_TP = 0
        mD_Per_TB = 0: mD_Per_TP = 0: mD_Amt_TB = 0: mD_Amt_TP = 0
        
        
        
                        
        If Not IsLineFileFound Then
            mSprAmt_MRP_TB = eVal(RsDms.Fields("Parts Amount"))
            mOilAmt_MRP_TB = eVal(RsDms.Fields("Lubricant Amount"))

            If (mSprAmt_MRP_TB + mOilAmt_MRP_TB) > 0 Then
                
                mDisAmt = eVal(RsDms.Fields("Discount Parts"))
        
                

                
                If mSprAmt_MRP_TB > 0 Then
                    
                    mTaxPer = Format(eVal(RsDms.Fields("VAT")) * 100 / (mSprAmt_MRP_TB), "0.000")
                    mTaxAmt = eVal(RsDms.Fields("VAT"))


'                    If eVal(RsDms.Fields("Discount Parts")) > 0 Then
'                        mDisPer = Format(eVal(RsDms.Fields("Discount Parts")) * 100 / (mSprAmt_MRP_TB + mTaxAmt), "0.00")
'                    Else
                        mDisPer = 0
'                    End If


                    
                    
                    
                    mQry = "Insert Into Sp_Stock ( "
                    mQry = mQry + " DocID, Site_Code, V_Type, V_No, V_Date, "
                    mQry = mQry + " Party_Code, Srl_No, L_C, Remark, Part_No, "
                    mQry = mQry + " Godown, Qty_Iss, Tax_Yn, Mrp_Yn, Amount, "
                    mQry = mQry + " TaxAmt, Disc_Amt, Disc_Per, Net_Amt, TaxPer, "
                    mQry = mQry + " Rate, Mrp_Rate, Part_SrlNo, Ord_DiscPer, Ord_DiscAmt, "
                    mQry = mQry + " Invoice_DocID, V_Date2, Rate2, Mrp_Rate2, Amount2, "
                    mQry = mQry + " Disc_Per2, Disc_Amt2, Net_Amt2, U_Name, U_EntDt, U_AE "
                    mQry = mQry + " ) "
                    mQry = mQry + " Values ( "
                    mQry = mQry + " '" & mChallanID & "', '" & mRecordSite & mRecordSite & "', '" & Trim(mChallanType) & "', " & Val(mChallanNo) & ", " & ConvertDate(MakeDate(left(RsDms.Fields("Order Date"), 10))) & ", "
                    mQry = mQry + " '" & mPartyCode & "', " & Val(mSrl) & ", '" & mLocal & "', '" & XNull(RsDms.Fields("Order_No")) & "', '" & XNull(RsDmsEnviro.Fields("DefaultPartNo")) & "', "
                    mQry = mQry + " '" & PubSprCounterGodown & "', 1, 1,1, " & mSprAmt_MRP_TB + mTax_Amt & ", "
                    mQry = mQry + " " & mTaxAmt & ", 0,0, " & mSprAmt_MRP_TB & ", " & mTaxPer & ",  "
                    mQry = mQry + " " & mSprAmt_MRP_TB + mTaxAmt & ", " & mSprAmt_MRP_TB + mTaxAmt & ", " & mSrl & ", 0,0, "
                    mQry = mQry + " '" & mDocId & "', " & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", " & mSprAmt_MRP_TB + mTaxAmt & ", " & mSprAmt_MRP_TB + mTaxAmt & ", " & mSprAmt_MRP_TB + mTaxAmt & ", "
                    mQry = mQry + " 0, 0, " & mSprAmt_MRP_TB & ", 'SIEBEL', '" & PubLoginDate & "', 'A'  "
                    mQry = mQry + " ) "
XNull (RsDmsEnviro.Fields("DefaultPartNo"))
                    
                    GCn.Execute mQry
                    
                    mDisPer = 0: mDisAmt = 0: mTaxPer = 0: mTaxAmt = 0
                End If
                
                
                If mOilAmt_MRP_TB > 0 Then
                    'If mDisAmt > 0 Then
                    '    mDisPer = Format(mDisAmt * 100 / (mOilAmt_MRP_TB + mTaxAmt), "0.00")
                    'Else
                        mDisPer = 0
                    'End If

'                    mQry = "Insert Into Sp_Stock ( "
'                    mQry = mQry + " DocID, Site_Code, V_Type, V_No, V_Date, "
'                    mQry = mQry + " Party_Code, Srl_No, L_C, Remark, Part_No, "
'                    mQry = mQry + " Godown, Qty_Iss, Tax_Yn, Mrp_Yn, Amount, "
'                    mQry = mQry + " TaxAmt, Disc_Amt, Disc_Per, Net_Amt, TaxPer, "
'                    mQry = mQry + " Rate, Mrp_Rate, Part_SrlNo, Ord_DiscPer, Ord_DiscAmt, "
'                    mQry = mQry + " Invoice_DocID, V_Date2, Rate2, Mrp_Rate2, Amount2, "
'                    mQry = mQry + " Disc_Per2, Disc_Amt2, Net_Amt2, U_Name, U_EntDt, U_AE "
'                    mQry = mQry + " ) "
'                    mQry = mQry + " Values ( "
'                    mQry = mQry + " '" & mChallanID & "', '" & mRecordSite & mRecordSite & "', '" & Trim(mChallanType) & "', " & Val(mChallanNo) & ", " & ConvertDate(MakeDate(left(RsDms.Fields("Order Date"), 10))) & ", "
'                    mQry = mQry + " '" & mPartyCode & "', 2, '" & mLocal & "', '" & XNull(RsDms.Fields("Order_No")) & "', '" & XNull(RsDmsEnviro.Fields("DefaultOilPartNo")) & "', "
'                    mQry = mQry + " '" & PubSprCounterGodown & "', 1, 1,1, " & mOilAmt_MRP_TB + mTax_Amt & ", "
'                    mQry = mQry + " " & mTaxAmt & ", " & mDisAmt & "," & mDisPer & ", " & mOilAmt_MRP_TB & ", " & mTaxPer & ",  "
'                    mQry = mQry + " " & mOilAmt_MRP_TB + mTaxAmt & ", " & mOilAmt_MRP_TB + mTaxAmt & ", " & mSrl & ", 0,0, "
'                    mQry = mQry + " '" & mDocId & "', " & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", " & mOilAmt_MRP_TB + mTaxAmt & ", " & mOilAmt_MRP_TB + mTaxAmt & ", " & mOilAmt_MRP_TB + mTaxAmt & ", "
'                    mQry = mQry + " " & mDisPer & ", " & mDisAmt & ", " & mOilAmt_MRP_TB & ", 'SIEBEL', '" & PubLoginDate & "', 'A'  "
'                    mQry = mQry + " ) "
                    
                    If mSprAmt_MRP_TB = 0 Then
                        mTaxPer = Format(eVal(RsDms.Fields("VAT")) * 100 / (mOilAmt_MRP_TB), "0.000")
                        mTaxAmt = eVal(RsDms.Fields("VAT"))
                    Else
                        mTaxPer = 0
                        mTaxAmt = 0
                    End If
                    
                    mQry = "Insert Into Sp_Stock ( "
                    mQry = mQry + " DocID, Site_Code, V_Type, V_No, V_Date, "
                    mQry = mQry + " Party_Code, Srl_No, L_C, Remark, Part_No, "
                    mQry = mQry + " Godown, Qty_Iss, Tax_Yn, Mrp_Yn, Amount, "
                    mQry = mQry + " TaxAmt, Disc_Amt, Disc_Per, Net_Amt, TaxPer, "
                    mQry = mQry + " Rate, Mrp_Rate, Part_SrlNo, Ord_DiscPer, Ord_DiscAmt, "
                    mQry = mQry + " Invoice_DocID, V_Date2, Rate2, Mrp_Rate2, Amount2, "
                    mQry = mQry + " Disc_Per2, Disc_Amt2, Net_Amt2, U_Name, U_EntDt, U_AE "
                    mQry = mQry + " ) "
                    mQry = mQry + " Values ( "
                    mQry = mQry + " '" & mChallanID & "', '" & mRecordSite & mRecordSite & "', '" & Trim(mChallanType) & "', " & Val(mChallanNo) & ", " & ConvertDate(MakeDate(left(RsDms.Fields("Order Date"), 10))) & ", "
                    mQry = mQry + " '" & mPartyCode & "', 2, '" & mLocal & "', '" & XNull(RsDms.Fields("Order_No")) & "', '" & XNull(RsDmsEnviro.Fields("DefaultOilPartNo")) & "', "
                    mQry = mQry + " '" & PubSprCounterGodown & "', 1, 1,1, " & mOilAmt_MRP_TB + mTax_Amt & ", "
                    mQry = mQry + " " & mTaxAmt & ", 0,0, " & mOilAmt_MRP_TB & ", " & mTaxPer & ",  "
                    mQry = mQry + " " & mOilAmt_MRP_TB + mTaxAmt & ", " & mOilAmt_MRP_TB + mTaxAmt & ", " & mSrl & ", 0,0, "
                    mQry = mQry + " '" & mDocId & "', " & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", " & mOilAmt_MRP_TB + mTaxAmt & ", " & mOilAmt_MRP_TB + mTaxAmt & ", " & mOilAmt_MRP_TB + mTaxAmt & ", "
                    mQry = mQry + " 0, 0, " & mOilAmt_MRP_TB & ", 'SIEBEL', '" & PubLoginDate & "', 'A'  "
                    mQry = mQry + " ) "
                    
                    
                    GCn.Execute mQry
                                        
                End If
            End If
        Else
               
            Master1.MoveFirst
            Do Until Master1.EOF
                
                mTaxPer = Round(VNull(Master1.Fields("Tax Amount")) * 100 / (VNull(Master1.Fields("Net_Amount")) - VNull(Master1.Fields("Tax Amount")) + VNull(Master1.Fields("Discount"))), 2)
                mRate = VNull(Master1.Fields("Net_Amount")) / VNull(Master1.Fields("Quantity"))
                mMRP = VNull(Master1.Fields("Net_Amount")) / VNull(Master1.Fields("Quantity"))
                
                mQry = "Insert Into Sp_Stock ( "
                mQry = mQry + " DocID, Site_Code, V_Type, V_No, V_Date, "
                mQry = mQry + " Party_Code, Srl_No, L_C, Remark, Part_No, "
                mQry = mQry + " Godown, Qty_Iss, Tax_Yn, Mrp_Yn, Amount, "
                mQry = mQry + " TaxAmt, Disc_Amt, Disc_Per, Net_Amt, TaxPer, "
                mQry = mQry + " Rate, Mrp_Rate, Part_SrlNo, Ord_DiscPer, Ord_DiscAmt, "
                mQry = mQry + " Invoice_DocID, V_Date2, Rate2, Mrp_Rate2, Amount2, "
                mQry = mQry + " Disc_Per2, Disc_Amt2, Net_Amt2, U_Name, U_EntDt, U_AE "
                mQry = mQry + " ) "
                mQry = mQry + " Values ( "
                mQry = mQry + " '" & mChallanID & "', '" & mRecordSite & mRecordSite & "', '" & Trim(mChallanType) & "', " & Val(mChallanNo) & ", " & ConvertDate(MakeDate(left(RsDms.Fields("Order Date"), 10))) & ", "
                mQry = mQry + " '" & mPartyCode & "', " & Val(mSrl) & ", '" & mLocal & "', '" & XNull(Master1.Fields("Order_No")) & "', '" & XNull(Master1.Fields("Part #")) & "', "
                mQry = mQry + " '" & PubSprCounterGodown & "', " & Val(Master1!Quantity) & ", 1,1, " & VNull(Master1.Fields("Net_Amount")) + VNull(Master1.Fields("Discount")) & ", "
                mQry = mQry + " " & VNull(Master1.Fields("Tax Amount")) & ", " & VNull(Master1.Fields("Discount")) & ",0, " & VNull(Master1.Fields("Net_Amount")) - VNull(Master1.Fields("Tax Amount")) & ", " & mTaxPer & ",  "
                mQry = mQry + " " & mRate & ", " & mMRP & ", " & mSrl & ", 0,0, "
                mQry = mQry + " '" & mDocId & "', " & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", " & mRate & ", " & mMRP & ", " & VNull(Master1.Fields("Net_Amount")) + VNull(Master1.Fields("Discount")) & ", "
                mQry = mQry + " 0, " & VNull(Master1.Fields("Discount")) & ", " & VNull(Master1.Fields("Net_Amount")) - VNull(Master1.Fields("Tax Amount")) & ", 'SIEBEL', '" & PubLoginDate & "', 'A'  "
                mQry = mQry + " ) "
                
                GCn.Execute mQry
                            
                If GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("part #") & "'").RecordCount > 0 Then
                    mTrnType = GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("part #") & "'").Fields(0).Value
                Else
                    mTrnType = "S"
                End If
                
                If mLubType = mTrnType Then
                    mOilAmt_MRP_TB = mOilAmt_MRP_TB + (IIf(IsNull(Master1.Fields("NTA")), 0, Master1.Fields("NTA")) * IIf(IsNull(Master1.Fields("Quantity")), 0, Master1.Fields("Quantity"))) - IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
                Else
                    mSprAmt_MRP_TB = mSprAmt_MRP_TB + (IIf(IsNull(Master1.Fields("NTA")), 0, Master1.Fields("NTA")) * IIf(IsNull(Master1.Fields("Quantity")), 0, Master1.Fields("Quantity"))) - IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
                End If
                
                mTax_Amt1 = mTax_Amt1 + VNull(Master1.Fields("Tax Amount"))
                mAmount = mAmount + IIf(IsNull(Master1.Fields("Net_Amount")), 0, Master1.Fields("Net_Amount"))
                mSrl = mSrl + 1
    
                Master1.MoveNext
            Loop
            
            
            If Round(mAmount - mTax_Amt1, 1) <> Round(IIf(IsNull(RsDms.Fields("Total_Parts_Amount")), 0, RsDms.Fields("Total_Parts_Amount")), 1) Then
                Call CreateErrLog(mVouCat, mInvoiceNo, "Line File Goods Value Total is not matched with Header file Goods Value (But Entry Posted in Automan")
                'ErrorCnt = 1
            End If
            
            
            If Round(mTax_Amt1, 1) <> Round(eVal(RsDms.Fields("Total_Tax_Amount")), 1) Then
                Call CreateErrLog(mVouCat, mInvoiceNo, "Line File Tax Value Total is not matched with Header file Tax Value (But Entry Posted in Automan)")
                'ErrorCnt = 1
            End If
        End If


        If mSprAmt_MRP_TB + mOilAmt_MRP_TB > 0 Then
            mD_Amt_MRP_TB = eVal(RsDms.Fields("Discount Parts"))
            mD_Amt_TB = eVal(RsDms.Fields("Discount Parts"))
        Else
            mD_Amt_MRP_TP = eVal(RsDms.Fields("Discount Parts"))
            mD_Amt_TP = eVal(RsDms.Fields("Discount Parts"))
        End If
        
        If mEditFlag Then
            GCn.Execute "Update Sp_Sale Set SprAmt_Mrp_TB=" & Round(mSprAmt_MRP_TB, 2) & ", SprAmt_MRP_TP = " & Round(mSprAmt_MRP_TP, 2) & ", " & _
                                        "OilAmt_MRP_TB = " & Round(mOilAmt_MRP_TB, 2) & ", OilAmt_MRP_TP = " & Round(mOilAmt_MRP_TP, 2) & ", " & _
                                        "D_Per_MRP_TB = " & Round(mD_Per_MRP_TB, 2) & ", D_Per_MRP_TP = " & Round(mD_Per_MRP_TP, 2) & ", " & _
                                        "D_Amt_MRP_TB = " & Round(mD_Amt_MRP_TB, 2) & ", D_Amt_MRP_TP = " & Round(mD_Amt_MRP_TP, 2) & ", " & _
                                        "SprAmt_TB = " & Round(mSprAmt_TB, 2) & ", SprAmt_TP = " & Round(mSprAmt_TP, 2) & ", " & _
                                        "OilAmt_TB = " & Round(mOilAmt_TB, 2) & ", OilAmt_TP = " & Round(mOilAmt_TP, 2) & ", " & _
                                        "D_Per_TB = " & Round(mD_Per_TB, 2) & ", D_Per_TP = " & Round(mD_Per_TP, 2) & ", " & _
                                        "D_Amt_TB = " & Round(mD_Amt_TB, 2) & ", D_Amt_TP = " & Round(mD_Amt_TP, 2) & ", " & _
                                        "Addition = 0, Tax_Amt = " & Round(IIf(IsNull(RsDms.Fields("Total_Tax_Amount")), 0, IIf(IsNull(RsDms.Fields("Discount Parts")), 0, Val(Format(mID(RsDms.Fields("Total_Tax_Amount"), 4, Len(RsDms.Fields("Total_Tax_Amount")) - 3), "0.00")))), 2) & ", " & _
                                        "Packing = " & Round(IIf(IsNull(RsDms.Fields("Other Charges")), 0, IIf(IsNull(RsDms.Fields("Other Charges")), 0, Val(mID(RsDms.Fields("Other Charges"), 4, Len(RsDms.Fields("Other Charges")) - 3)))), 2) & ", " & _
                                        "TOT_Per = 0, TOT_Amt = 0, ReSalTax_Per = 0, ReSalTax_Amt = 0,total_amt = " & Round(Val(Format(mID(RsDms.Fields("Parts_Invoice_Amount"), 4, Len(RsDms.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) & ", " & _
                                        "Rounded = " & Round(Val(Format(mID(RsDms.Fields("Parts_Invoice_Amount"), 4, Len(RsDms.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) - Round(Val(Format(mID(RsDms.Fields("Parts_Invoice_Amount"), 4, Len(RsDms.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 2) & ", " & _
                                        "Det_Tax = 1, AcPosting_YN = 1, U_Name = 'Siebel', U_EntDt = " & ConvertDate(Format(PubLoginDate, "Short Date")) & ", U_AE = 'E' " & _
                                        "Where DocId='" & mChallanID & "'"
            
            GCn.Execute "Update Sp_Sale Set SprAmt_MRP_TB = " & mSprAmt_MRP_TB & ", SprAmt_MRP_TP = " & mSprAmt_MRP_TP & ", OilAmt_MRP_TB = " & mOilAmt_MRP_TB & ", " & _
                                        "OilAmt_MRP_TP = " & mOilAmt_MRP_TP & ", D_Per_MRP_TB = " & mD_Per_MRP_TB & ", D_Per_MRP_TP = " & mD_Per_MRP_TP & ", " & _
                                        "D_Amt_MRP_TB = " & mD_Amt_MRP_TB & ", D_Amt_MRP_TP = " & mD_Amt_MRP_TP & ", SprAmt_TB = " & mSprAmt_TB & " , " & _
                                        "SprAmt_TP = " & mSprAmt_TP & ", OilAmt_TB = " & mOilAmt_TB & ", OilAmt_TP = " & mOilAmt_TP & ", " & _
                                        "D_Per_TB = " & mD_Per_TB & ", D_Per_TP = " & mD_Per_TP & ", D_Amt_TB = " & mD_Amt_TB & ", D_Amt_TP = " & mD_Amt_TP & ", " & _
                                        "Addition = 0, Tax_Amt = " & Round(IIf(IsNull(RsDms.Fields("Total_Tax_Amount")), 0, IIf(IsNull(RsDms.Fields("Discount Parts")), 0, Val(Format(mID(RsDms.Fields("Total_Tax_Amount"), 4, Len(RsDms.Fields("Total_Tax_Amount")) - 3), "0.00")))), 2) & ", " & _
                                        "Packing = " & Round(IIf(IsNull(RsDms.Fields("Other Charges")), 0, IIf(IsNull(RsDms.Fields("Other Charges")), 0, Val(mID(RsDms.Fields("Other Charges"), 4, Len(RsDms.Fields("Other Charges")) - 3)))), 2) & ", " & _
                                        "TOT_Per = 0, TOT_Amt = 0, ReSalTax_Per = 0, ReSalTax_Amt = 0, " & _
                                        "total_amt = " & Round(Val(Format(mID(RsDms.Fields("Parts_Invoice_Amount"), 4, Len(RsDms.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) & ", " & _
                                        "Rounded = " & Round(Val(Format(mID(RsDms.Fields("Parts_Invoice_Amount"), 4, Len(RsDms.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) - Round(Val(Format(mID(RsDms.Fields("Parts_Invoice_Amount"), 4, Len(RsDms.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 2) & ", " & _
                                        "Det_Tax = 1, AcPosting_YN = 1, U_Name = 'Siebel', U_EntDt = " & ConvertDate(Format(PubLoginDate, "Short Date")) & ", U_AE = 'E' " & _
                                        "Where DocId='" & mDocId & "'"
            mEditFlag = False
        Else
            'Insert New Rec for Challan
            mQry = "Insert Into Sp_Sale ( "
            mQry = mQry + "DocID, DocIDHelp, Site_Code, V_Type, V_No, "
            mQry = mQry + "V_Date, Party_Code, Cash_Credit, Party_Name, L_C, "
            mQry = mQry + "Form_Code, CrAc, SiebelDocID, Invoice_DocID, PType, "
            mQry = mQry + "GP_No, GP_Date, SprAmt_Mrp_TB, SprAmt_Mrp_TP, OilAmt_Mrp_TB, "
            mQry = mQry + "OilAmt_Mrp_TP, D_Per_Mrp_TB, D_Per_Mrp_TP, D_Amt_Mrp_TB, D_Amt_Mrp_TP, "
            mQry = mQry + "SprAmt_TB, SprAmt_TP, OilAmt_TB, OilAmt_TP, D_Per_TB, "
            mQry = mQry + "D_Per_TP, D_Amt_TB, D_Amt_TP, Addition, Tax_Amt, "
            mQry = mQry + "Packing, Tot_Per, Tot_Amt, ReSalTax_Per, ReSalTax_Amt, "
            mQry = mQry + "Total_amt, Rounded, Det_Tax, AcPosting_Yn, U_Name, "
            mQry = mQry + "U_EntDt, U_AE"
            mQry = mQry + ")"
            mQry = mQry + "Values ( "
            mQry = mQry + "'" & mChallanID & "', '" & Replace(mChallanID, " ", "") & "', '" & mRecordSite & mRecordSite & "', '" & Trim(mChallanType) & "', " & mChallanNo & ", "
            mQry = mQry + " " & ConvertDate(MakeDate(left(RsDms.Fields("Order Date"), 10))) & ", '" & mPartyCode & "', '" & RsDms.Fields("Mode Of Payment") & "', '" & left(mHeaderParty, 40) & "', '" & mLocal & "',   "
            mQry = mQry + " '" & mFormCode & "', '" & mDebitAc & "', '" & RsDms.Fields("Order_No") & "', '" & mDocId & "', 'General', "
            mQry = mQry + " '" & mGatePassID & "', " & ConvertDate(MakeDate(left(RsDms.Fields("Order Date"), 10))) & ", " & Round(mSprAmt_MRP_TB, 2) & ", " & Round(mSprAmt_MRP_TP, 2) & ", " & Round(mOilAmt_MRP_TB, 2) & ", "
            mQry = mQry + " " & Round(mOilAmt_MRP_TP, 2) & ", " & Round(mD_Per_MRP_TB, 2) & ", " & Round(mD_Per_MRP_TP, 2) & ", " & Round(mD_Amt_MRP_TB, 2) & ", " & Round(mD_Amt_MRP_TP, 2) & ", "
            mQry = mQry + " " & Round(mSprAmt_TB, 2) & ", " & Round(mSprAmt_TP, 2) & ", " & Round(mOilAmt_TB, 2) & ", " & Round(mOilAmt_TP, 2) & ", " & Round(mD_Per_TB, 2) & ", "
            mQry = mQry + " " & Round(mD_Per_TP, 2) & ", " & Round(mD_Amt_TB, 2) & ", " & Round(mD_Amt_TP, 2) & ", " & 0 & ", " & eVal(RsDms.Fields("Total_Tax_Amount")) & ", "
            mQry = mQry + " " & eVal(RsDms.Fields("Other Charges")) & ", 0, 0, 0, 0, "
            mQry = mQry + " " & eVal(RsDms.Fields("Parts_Invoice_Amount")) & ", 0, 1, 1, 'SIEBEL', " & ConvertDate(PubLoginDate) & ", 'A' "
            mQry = mQry + " )"
            GCn.Execute mQry
            
            mQry = "Insert Into Sp_Sale( "
            mQry = mQry + " DocID, DocIDHelp, Site_Code, V_Type, V_No, "
            mQry = mQry + " V_Date, Party_Code, Cash_Credit, Party_Name, L_C, "
            mQry = mQry + " Form_Code, CrAc, SiebelDocID, Invoice_DocID, PType, "
            mQry = mQry + " GP_No, GP_Date, SprAmt_Mrp_TB, SprAmt_Mrp_TP, OilAmt_Mrp_TB, "
            mQry = mQry + " OilAmt_Mrp_TP, D_Per_Mrp_TB, D_Per_Mrp_TP, D_Amt_Mrp_tB, D_Amt_Mrp_TP, "
            mQry = mQry + " SprAmt_TB, SprAmt_TP, OilAmt_TB, OilAmt_TP, D_Per_TB, "
            mQry = mQry + " D_Per_TP, D_Amt_TB, D_Amt_TP, Addition, Tax_Amt, "
            mQry = mQry + " Packing, Tot_Per, Tot_Amt, ReSalTax_Per, ReSalTax_Amt, "
            mQry = mQry + " Total_Amt, Rounded, Det_Tax, AcPosting_Yn, U_Name, "
            mQry = mQry + " U_EntDt, U_AE "
            mQry = mQry + " ) "
            mQry = mQry + " Values   ("
            mQry = mQry + " '" & mDocId & "', '" & Replace(mDocId, " ", "") & "', '" & mRecordSite & mRecordSite & "', '" & Trim(mV_Type) & "', " & Val(Right(mInvoiceNo, 5)) & ", "
            mQry = mQry + " " & ConvertDate(MakeDate(RsDms.Fields("Invoice_Date"))) & ", '" & mPartyCode & "', '" & RsDms.Fields("Mode Of Payment") & "', '" & left(mHeaderParty, 40) & "', '" & mLocal & "', "
            mQry = mQry + " '" & mFormCode & "', '" & mDebitAc & "', '" & mInvoiceNo & "', '', 'General', "
            mQry = mQry + " '" & mGatePassID & "', " & ConvertDate(MakeDate(left(RsDms.Fields("Order Date"), 10))) & ", " & mSprAmt_MRP_TB & ", " & mSprAmt_MRP_TP & ", " & mOilAmt_MRP_TB & ",  "
            mQry = mQry + " " & mOilAmt_MRP_TP & ", " & mD_Per_MRP_TB & ", " & mD_Per_MRP_TP & ", " & mD_Amt_MRP_TB & ", " & mD_Amt_MRP_TP & ", "
            mQry = mQry + " " & mSprAmt_TB & ", " & mSprAmt_TP & ", " & mOilAmt_TB & ", " & mOilAmt_TP & ", " & mD_Per_TB & ", "
            mQry = mQry + " " & mD_Per_TP & ", " & mD_Amt_TB & ", " & mD_Amt_TP & ", 0, " & eVal(RsDms.Fields("Total_Tax_Amount")) & ", "
            mQry = mQry + " " & eVal(RsDms.Fields("Other Charges")) & ", 0,0,0,0, "
            mQry = mQry + " " & eVal(RsDms.Fields("Parts_Invoice_Amount")) & ", 0, 1, 1, 'SIEBEL', "
            mQry = mQry + " " & ConvertDate(PubLoginDate) & ", 'A' "
            mQry = mQry + ")"
            
            GCn.Execute mQry
            
        End If
        RsDms.MoveNext
    Loop

    
    GCn.CommitTrans
    mTrans = False
    MsgBox "Spare Detail Imported Successfully"
    
lblExit:
    Set RsNew = Nothing
    Exit Sub
ELoop:
    MsgBox err.Description
    If mTrans Then GCn.RollbackTrans
End Sub



Private Sub VehiclePurchaseDataUpdate()
 '' On Error GoTo Eloop
Dim MasterCode As String, DocID As String, mV_Type As String, mPartyCode As String, mForm_Code As String
Dim mDebitAc As String, mMfgMonth As String, mMfgYear As String, mColourCode As String, mColourName As String, mGodownCode As String
Dim mTaxPer As Double, mDeductionCode As String, mAdditionCode As String
Dim mLength1 As Integer, mLength2 As Integer, mTaxOnDelivery As Boolean
Dim EditFlag As Boolean
Dim RsX As ADODB.Recordset
Dim xDocId$
Dim CodeCnt As Long

     
    GCn.BeginTrans
   
    CopyCnt = 0
    ErrorCnt = 0
    
    Dim mVouCat As String
    mV_Type = "V_PB"
    mVouCat = "Vehicle Purchase"
    
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from Veh_Purch1 where Left(DocID,1)='" & PubDivCode & "' and " & cMID("DocID", "2", "1") & "='" & PubSiteCode & "' and V_Type='" & mV_Type & "'").Fields(0).Value
    Do Until RsDms.EOF
        GCn.Execute "Delete From DmsErrLog Where [Key]='" & RsDms!Invoice_No & "'"
        ErrorCnt = 0
        If IsNull(StringPass(RsDms.Fields("Invoice_No"))) Or StringPass(RsDms.Fields("Invoice_No")) = "" Then ErrorCnt = 1
        EditFlag = False
        If GCn.Execute("Select PBill_No from Veh_Purch1 where Pbill_No='" & left(StringPass(RsDms.Fields("Invoice_no")), 10) & "'").RecordCount > 0 Then
            ErrorCnt = 1
        End If

                                                    
        If StringPass(RsDms.Fields("Supplier_Name")) = "" Then
            Call CreateErrLog(mVouCat, RsDms!Invoice_No, "Supplier Name is blank in excel file")
            ErrorCnt = 1
        Else
            If GCn.Execute("Select AutomanSupplierCode as SubCode From DmsSupplierAc With(NOLOCK) Where DmsCode ='" & StringPass(RsDms.Fields("Supplier_Name")) & "'").RecordCount > 0 Then
                mPartyCode = GCn.Execute("Select AutomanSupplierCode as SubCode From DmsSupplierAc With(NOLOCK) Where DmsCode ='" & StringPass(RsDms.Fields("Supplier_Name")) & "'").Fields(0).Value
            Else
                Call CreateErrLog(mVouCat, RsDms!Invoice_No, "Supplier Name - " & XNull(RsDms!Supplier_Name) & " Not Found In Automan")
                ErrorCnt = 1
            End If
        End If
       
        
        
        
        If XNull(RsDms!Invoice_Date) = "" Then
            Call CreateErrLog(mVouCat, RsDms!Invoice_No, " " & XNull(RsDms!Invoice_Date) & " Telco Invoice Date is Empty")
            ErrorCnt = 1
        End If
        
        If IsNull(StringPass(RsDms!Godown)) Or StringPass(RsDms!Godown) = "" Then
            Call CreateErrLog(mVouCat, RsDms!Invoice_No, " Godown Name " & XNull(RsDms!Godown) & " Not Found ")
            ErrorCnt = 1
        End If
        
        If IsNull(StringPass(RsDms!VC_Number)) Or StringPass(RsDms!VC_Number) = "" Then
            Call CreateErrLog(mVouCat, RsDms!Invoice_No, "VC_Number - " & XNull(RsDms!VC_Number) & "VC_Number Empty")
            ErrorCnt = 1
        Else
            If GCn.Execute("Select Model from Model where Model='" & StringPass(RsDms!VC_Number) & "'").RecordCount = 0 Then
                Call CreateErrLog(mVouCat, RsDms!Invoice_No, "VC_Number - " & XNull(RsDms!VC_Number) & " Not Found In Automan")
                ErrorCnt = 1
            End If
        End If
        
        If IsNull(StringPass(RsDms!Chassis_No)) Or StringPass(RsDms!Chassis_No) = "" Then
            Call CreateErrLog(mVouCat, RsDms!Invoice_No, "Chassis No  - " & XNull(RsDms!Chassis_No) & " Chassis Number is Empty")
            ErrorCnt = 1
        End If
        
        If GCn.Execute("Select ChassisNo from Veh_Stock where ChassisNo='" & StringPass(RsDms.Fields("Chassis_No")) & "'").RecordCount > 0 Then
            EditFlag = True
        End If
        
        If IsNull(StringPass(RsDms!Narration)) Or StringPass(RsDms!Narration) = "" Then
            Call CreateErrLog(mVouCat, RsDms!Invoice_No, "Chassis No - " & XNull(RsDms!Narration) & " Narration Field is Empty (Engine Number)")
            ErrorCnt = 1
        End If
            
        
        If Len(StringPass(RsDms.Fields("Chassis_No"))) = 17 Then
            If GCn.Execute("Select Name from Chas_Mth where Month_CD='" & mID(StringPass(RsDms!Chassis_No), 12, 1) & "'").RecordCount = 0 Then
                Call CreateErrLog(mVouCat, RsDms!Invoice_No, " " & XNull(RsDms!Chassis_No) & " Chassis Mfg. Month Name is not defined in Chas_Mth Table")
                ErrorCnt = 1
            Else
                mMfgMonth = GCn.Execute("Select Name from Chas_Mth where Month_CD='" & mID(StringPass(RsDms!Chassis_No), 12, 1) & "'").Fields(0).Value
            End If
        ElseIf Len(StringPass(RsDms.Fields("Chassis_No"))) > 17 Then
            mMfgMonth = Format(RsDms.Fields("Invoice_Date"), "MMMM")
        Else
            If GCn.Execute("Select Name from Chas_Mth where Month_CD='" & mID(StringPass(RsDms!Chassis_No), 7, 1) & "'").RecordCount = 0 Then
                Call CreateErrLog(mVouCat, RsDms!Invoice_No, " " & XNull(RsDms!Chassis_No) & " Chassis Mfg. Month Name is not defined in Chas_Mth Table")
                ErrorCnt = 1
            Else
                mMfgMonth = GCn.Execute("Select Name from Chas_Mth where Month_CD='" & mID(StringPass(RsDms!Chassis_No), 7, 1) & "'").Fields(0).Value
            End If
        End If
        
        If Len(StringPass(RsDms.Fields("Chassis_No"))) = 17 Then
            Select Case (mID(StringPass(RsDms!Chassis_No), 10, 1))
                Case "9"
                    mMfgYear = "2009"
                Case "0"
                    mMfgYear = "2010"
                Case "B"
                    mMfgYear = "2011"
                Case "C"
                    mMfgYear = "2012"
                Case "D"
                    mMfgYear = "2013"
                Case "E"
                    mMfgYear = "2014"
                Case "F"
                    mMfgYear = "2015"
                Case "G"
                    mMfgYear = "2016"
                Case "H"
                    mMfgYear = "2017"
                Case "I"
                    mMfgYear = "2018"
            End Select
        ElseIf Len(StringPass(RsDms.Fields("Chassis_No"))) > 17 Then
            mMfgYear = Format(RsDms.Fields("Invoice_Date"), "YYYY")
        Else
            If GCn.Execute("Select Name from Chas_Yr where Year_Cd='" & mID(StringPass(RsDms!Chassis_No), 8, 2) & "'").RecordCount = 0 Then
                Call CreateErrLog(mVouCat, RsDms!Invoice_No, " " & XNull(RsDms!Chassis_No) & " Chassis Mfg. Year Name is not defined in Chas_YR Table")
                ErrorCnt = 1
            Else
                mMfgYear = GCn.Execute("Select Name from Chas_Yr where Year_Cd='" & mID(StringPass(RsDms!Chassis_No), 8, 2) & "'").Fields(0).Value
            End If
        End If
                
        mGodownCode = GetGodownCode(left(StringPass(RsDms.Fields("Godown")), 20), "1")
        
        If GCn.Execute("Select count(*) from Model where Model='" & StringPass(RsDms.Fields("VC_Number")) & "'").Fields(0) > 0 Then
            mColourCode = GCn.Execute("Select Col_Code from Model where Model='" & StringPass(RsDms.Fields("VC_Number")) & "'").Fields(0).Value
        End If

        If GCn.Execute("Select Col_Code from ColMast where Col_Code='" & mColourCode & "'").RecordCount > 0 Then
            mColourName = GCn.Execute("Select Col_Desc from ColMast where Col_Code='" & mColourCode & "'").Fields(0).Value
        End If
        
        
        
        Dim mShortYear As String
        If XNull(RsDms.Fields("Invoice_Date")) <> "" Then
            If Month(RsDms.Fields("Invoice_Date")) > 3 Then
                mShortYear = Right(Format(RsDms.Fields("Invoice_Date"), "yy"), 1) & Right(Val(Format(RsDms.Fields("Invoice_Date"), "yy")) + 1, 1)
            Else
                mShortYear = Right(Val(Format(RsDms.Fields("Invoice_Date"), "yy")) - 1, 1) & Right(Format(RsDms.Fields("Invoice_Date"), "yy"), 1)
            End If
        End If
        DocID = PubDivCode & PubSiteCode & PubSiteCode & " " & mV_Type & "SBL" & mShortYear & Right("00000000" & CodeCnt, 8)
        
        
        Dim mTot_Amt As Double, mTax_Amt As Double, mMisc_Amt As Double
        Dim mDeduction As Double, mAddition As Double, mAmount As Double
        
        mTot_Amt = 0: mTax_Amt = 0: mMisc_Amt = 0
        mDeduction = 0: mAddition = 0: mAmount = 0
        
        
        mTot_Amt = RsDms!Value
        
        
        If UCase(XNull(RsDmsEnviro!VehicleTaxOnDeliveryCharges)) = "Y" Then
            mTaxOnDelivery = True
        Else
            mTaxOnDelivery = False
        End If
        
        If mTaxOnDelivery Then
            mMisc_Amt = 0
        Else
            mMisc_Amt = VNull(RsDms.Fields("Delivery Charges"))
        End If
        
        If eVal(RsDms.Fields("Tax Cst")) > 0 Then
            mForm_Code = XNull(RsDmsEnviro!VehicleCentralPurchaseTaxForm)
             mTaxPer = GCn.Execute("Select Tax_Per From TaxForms Where Form_Code = '" & mForm_Code & "'").Fields(0).Value
            mTax_Amt = eVal(RsDms.Fields("Tax Cst"))
        Else
            mForm_Code = XNull(RsDmsEnviro!VehicleLocalPurchaseTaxForm)
            mTaxPer = GCn.Execute("Select Tax_Per From TaxForms Where Form_Code = '" & mForm_Code & "'").Fields(0).Value
            If IsNull(RsDms.Fields("VatTax")) Or RsDms.Fields("VatTax") = "" Then
                mTax_Amt = Round((mTot_Amt) * mTaxPer / (100 + mTaxPer), 2)     ''- mMisc_Amt
            Else
                mTax_Amt = Val(RsDms.Fields("VatTax"))
            End If
        
        End If
            
        If mTaxOnDelivery Then
            mAddition = VNull(RsDms.Fields("Delivery Charges"))
        Else
            mAddition = 0
        End If
        mDeduction = VNull(RsDms.Fields("Total Discount"))
        mAmount = mTot_Amt + mDeduction - (mMisc_Amt + mTax_Amt + mAddition)
        
        
        If EditFlag = True Then
            'ArpitStart
            GCn.Execute "Update Veh_Purch1 Set Amount = " & mAmount & ", Tot_Amount = " & mTot_Amt & ", " & _
                        "Tax_Per = " & mTaxPer & ", Tax_Amt = " & mTax_Amt & ", Addition = " & mAddition & ", " & _
                        "Deduction = " & mDeduction & ", Misc_Amt = " & mMisc_Amt & ", U_EntDt = " & ConvertDate(date) & " " & _
                        "Where DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(RsDms!Chassis_No) & "' )"
                        
            
            
            Set RsX = GCn.Execute("Select Pur_DocId From Veh_Stock Where ChassisNo = '" & RsDms!Chassis_No & "'")
            If RsX.RecordCount > 0 Then
                GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & XNull(RsX(0)) & "' And Trn_Type='D'"
                GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & XNull(RsX(0)) & "' And Trn_Type='A'"
            End If
            
           
            If mDeduction > 0 Then
                Set RsX = GCn.Execute("Select Pur_DocId From Veh_Stock Where ChassisNo = '" & RsDms!Chassis_No & "'")
                 If RsX.RecordCount > 0 Then xDocId = XNull(RsX!Pur_DocId)
                
                    If GCn.Execute("Select DocId From Veh_Purch2 Where DocId = '" & xDocId & "'").RecordCount > 0 Then
                        'GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & xDocId & "' And Trn_Type='D'"
                        'GCn.Execute "Update Veh_Purch2  Set Rate = " & mDeduction & " " & _
                                    "Where DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(Master!Chassis_No) & "') And Trn_Type='D'"
                    End If
                GCn.Execute ("INSERT INTO dbo.Veh_Purch2  (DocId,Srl_No,Site_code,v_type,v_no,trn_type,Prod_code,qty,Rate, " & _
                                            " U_Name ,U_EntDt,U_AE   ) " & _
                                     "VALUES  ('" & xDocId & "',1,'" & PubSiteCode & PubSiteCode & "' ,'" & mV_Type & "', " & _
                                     " " & CodeCnt & ",'D','" & mDeductionCode & "', " & _
                                    " 1," & mDeduction & ",'Siebel','" & Format(PubLoginDate, "Short Date") & "','A')")
    
    
            End If
           
            
            If mAddition > 0 Then
                Set RsX = GCn.Execute("Select Pur_DocId From Veh_Stock Where ChassisNo = '" & RsDms!Chassis_No & "'")
                If RsX.RecordCount > 0 Then xDocId = XNull(RsX!Pur_DocId)
                
                If GCn.Execute("Select DocId From Veh_Purch2 Where DocId = '" & xDocId & "'").RecordCount > 0 Then
                    'GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & xDocId & "' And Trn_Type='A'"
                    'GCn.Execute "Update Veh_Purch2 Set Rate = " & mAddition & " " & _
                                "Where DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(Master!Chassis_No) & "') And Trn_Type='A'"
                End If
                
                  
                GCn.Execute ("INSERT INTO dbo.Veh_Purch2  (DocId,Srl_No,Site_code,v_type,v_no,trn_type,Prod_code,qty,Rate, " & _
                                            " U_Name ,U_EntDt,U_AE   ) " & _
                                     "VALUES  ('" & xDocId & "',2,'" & PubSiteCode & PubSiteCode & "' ,'" & mV_Type & "', " & _
                                     " " & CodeCnt & ",'A','" & mAdditionCode & "', " & _
                                    " 1," & mAddition & ",'Siebel','" & Format(PubLoginDate, "Short Date") & "','A')")
                  
    
            End If
                     If ErrorCnt = 0 Then
                        GCn.Execute "Update Veh_Stock Set Rate = " & mAmount & ", VRate = " & mTot_Amt & " " & _
                                    "Where Pur_DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(RsDms!Chassis_No) & "' )"
                    End If
                        
            EditFlag = False
            'ArpitEnd
        Else
           If ErrorCnt = 0 Then
               GCn.Execute ("INSERT INTO dbo.Veh_Purch1  (DocId,DocIDHelp,Site_Code ,V_Type ,V_NO ,V_DATE,PartyCode,PBill_No,Pbill_Date,BMS_Category " & _
                                             " ,DueDate ,Gate,GateDate ,Form_Code ,Amount ,Addition ,Deduction ,Exsice ,Tax_Per ,Tax_Amt ,Misc_Amt " & _
                                             " ,Tot_AMOUNT ,DrAcCode,U_Name ,U_EntDt,U_AE   ) " & _
                                             "VALUES  ('" & DocID & "','" & Replace(DocID, " ", "") & "','" & PubSiteCode & PubSiteCode & "' ,'" & mV_Type & "', " & _
                                             " " & CodeCnt & ",'" & MakeDate(RsDms!Invoice_Date) & "','" & mPartyCode & "', " & _
                                            " '" & RsDms!Invoice_No & "','" & MakeDate(RsDms!Invoice_Date) & "','','" & MakeDate(RsDms!Invoice_Date) & "', '', " & _
                                            " '" & MakeDate(RsDms!Invoice_Date) & "','" & mForm_Code & "' ," & mAmount & " ," & mAddition & " ," & mDeduction & ",0," & mTaxPer & " , " & _
                                            " " & mTax_Amt & " ," & mMisc_Amt & "," & mTot_Amt & ",'" & mDebitAc & "','Siebel','" & Format(PubLoginDate, "Short Date") & "','A')")
         If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
              mDeductionCode = "Dis"
          End If
        
               GCn.Execute ("INSERT INTO dbo.Veh_Purch2  (DocId,Srl_No,Site_code,v_type,v_no,trn_type,Prod_code,qty,Rate, " & _
                                                    " U_Name ,U_EntDt,U_AE   ) " & _
                                             "VALUES  ('" & DocID & "',1,'" & PubSiteCode & PubSiteCode & "' ,'" & mV_Type & "', " & _
                                             " " & CodeCnt & ",'D','" & mDeductionCode & "', " & _
                                            " 1," & mDeduction & ",'Siebel','" & Format(PubLoginDate, "Short Date") & "','A')")
        
               GCn.Execute ("INSERT INTO dbo.Veh_Purch2  (DocId,Srl_No,Site_code,v_type,v_no,trn_type,Prod_code,qty,Rate, " & _
                                                    " U_Name ,U_EntDt,U_AE   ) " & _
                                             "VALUES  ('" & DocID & "',2,'" & PubSiteCode & PubSiteCode & "' ,'" & mV_Type & "', " & _
                                             " " & CodeCnt & ",'A','" & mAdditionCode & "', " & _
                                            " 1," & mAddition & ",'Siebel','" & Format(PubLoginDate, "Short Date") & "','A')")
                    
                        mLength1 = InStr(1, StringPass(RsDms!Narration), "Engine") + Len("Engine Number - ")
                        mLength2 = InStr(1, StringPass(RsDms!Narration), "Chassis")
                        mLength2 = (mLength2 - mLength1)
                        
           
           
                 GCn.Execute ("INSERT INTO dbo.Veh_Stock(ChassisNo,Pur_DocId,pur_SrlNo,Pur_DocIDHelp,Pur_SiteCode,Pur_VType,Pur_VNo " & _
                      ",Pur_VDate,Mfg_Month,Mfg_Yr,InDate,Model,Chas_Type,Godown   ,EngineNo,Rate,Fixed,vrate,Colour_Code,Colours,Tax_YN,PBill_No,Pbill_Date,PartyCode,U_Name,U_EntDt,U_AE   ) " & _
                      "VALUES  ('" & StringPass(RsDms!Chassis_No) & "' ,'" & DocID & "',1,'" & Replace(DocID, " ", "") & "','" & PubSiteCode & PubSiteCode & "' " & _
                      " ,'" & mV_Type & "'," & CodeCnt & ", '" & MakeDate(RsDms!Invoice_Date) & "','" & mMfgMonth & "','" & mMfgYear & "' " & _
                      " ,'" & MakeDate(RsDms!Invoice_Date) & "','" & StringPass(RsDms!VC_Number) & "','" & left(StringPass(RsDms!Chassis_No), 6) & "' " & _
                      " , '" & mGodownCode & "','" & Replace(Trim(mID(StringPass(RsDms!Narration), mLength1, mLength2)), ".", "") & "' " & _
                      " ," & mAmount & ",0," & mTot_Amt & ",'" & mColourCode & "','" & mColourName & "',1, '" & StringPass(RsDms!Invoice_No) & "' " & _
                      " ,'" & MakeDate(RsDms!Invoice_Date) & "','" & mPartyCode & "', 'Siebel','" & Format(PubLoginDate, "Short Date") & "', 'A')")
            End If
                              
        End If
       
        CodeCnt = CodeCnt + 1
        RsDms.MoveNext
    Loop
    GCn.CommitTrans
   
lblExit:
    Set RsNew = Nothing
    Exit Sub
ELoop:
  
    Resume Next
End Sub
Function GetGodownCode(DmsCode As String, Apply_For As String) As String
    Dim mMaxID As String, mQry As String
    If GCn.Execute("Select Count(*) From Godown Where DmsCode = '" & DmsCode & "'").Fields(0).Value = 0 Then
        mMaxID = GCn.Execute("Select IsNull(Max(Convert(Float,God_Code)),0)+1 From Godown Where  IsNUMERIC(God_Code)=1 ").Fields(0).Value
        mQry = "INSERT INTO dbo.Godown (God_Code, Site_Code, God_Name, appli_for, u_name, u_entdt, u_ae, trf_date, oldcode, dmscode) " & _
               "VALUES ('" & mMaxID & "', '" & PubSiteCode & "', '" & DmsCode & "', '" & Apply_For & "', '" & pubUName & "', '" & PubLoginDate & "', 'A', Null, '', '" & DmsCode & "')"
        GCn.Execute mQry
        GetGodownCode = mMaxID
    Else
        GetGodownCode = GCn.Execute("Select IsNull(God_Code,'') from Godown Where DmsCode = '" & DmsCode & "'").Fields(0).Value
    End If
End Function


Function GetModelCatCode(DmsCode As String) As String
    Dim mMaxID As String, mQry As String
    If GCn.Execute("Select Count(*) From Model_Cat Where DmsCode = '" & DmsCode & "' Or ModelCat_Name = '" & DmsCode & "' ").Fields(0).Value = 0 Then
        mMaxID = GCn.Execute("Select IsNull(Max(Convert(Float,ModelCat_Code)),0)+1 From Model_Cat Where  IsNUMERIC(ModelCat_Code)=1 ").Fields(0).Value
        mQry = "INSERT INTO dbo.Model_Cat (ModelCat_Code, Site_Code, ModelCat_Name, u_name, u_entdt, u_ae, trf_date, oldcode, dmscode) " & _
               "VALUES ('" & mMaxID & "', '" & PubSiteCode & "', '" & DmsCode & "', '" & pubUName & "', '" & PubLoginDate & "', 'A', Null, '', '" & DmsCode & "')"
        GCn.Execute mQry
        GetModelCatCode = mMaxID
    Else
        GetModelCatCode = GCn.Execute("Select IsNull(ModelCat_Code,'') from Model_Cat Where DmsCode = '" & DmsCode & "' Or ModelCat_Name = '" & DmsCode & "' ").Fields(0).Value
    End If
End Function



Private Function StringPass(ByVal Temp As Variant) As String
    Temp = XNull(Temp)
    Temp = Replace(Temp, "'", "`")
    StringPass = Temp
End Function
Private Sub InsSkipRecMessage(ByVal Index As Integer, ByVal RecordNo As Long, ByVal ValueDetails, ByVal ColoumnDetail, ByVal ErrorDescription)
    ErrorCnt = ErrorCnt + 1
'    lblRecError(Index).CAPTION = ErrorCnt: lblRecError(Index).Refresh
    'NIKHIL
'    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & RecordNo & ",'" & left(StringPass(ValueDetails), 50) & "','" & ColoumnDetail & "','" & ErrorDescription & "')")
End Sub

Private Sub ModelMasterUpdate(ByVal Index As Long)
'' On Error GoTo Eloop
Dim MasterCode As String, mCatCode As String, mGrpCode As String, mColCode As String
Dim Chas_Type As String, Model_Desc As String, Model_Desc1 As String, Model_desc2 As String
Dim Sales_Desc As String, FMSN As String
Dim CodeCnt As Variant
    
  GCn.Execute "Delete From DmsErrLog "
    GCn.BeginTrans
    

    CopyCnt = 0
    ErrorCnt = 0

    CodeCnt = GCn.Execute("Select " & vIsNull("Max(Right(ModelCat_Code,2))", "0") & " from Model_Cat Where IsNumeric(Right(ModelCat_Code,2))=1").Fields(0).Value
    If IsNumeric(CodeCnt) Then
        CodeCnt = CodeCnt + 1
    Else
        CodeCnt = 30
    End If
    ErrorCnt = 0
    Do Until RsDms.EOF
        If IsNull(StringPass(RsDms.Fields("Parent Product Line"))) Or StringPass(RsDms.Fields("Parent Product Line")) = "" Then ErrorCnt = 1
        
        If GCn.Execute("Select ModelCat_Name from Model_Cat where ModelCat_Name='" & left(StringPass(RsDms.Fields("Parent Product Line")), 20) & "'").RecordCount > 0 Then ErrorCnt = 1
                        
        MasterCode = PubDivCode & CodeCnt
      If ErrorCnt = 0 Then
        GCn.Execute ("INSERT INTO dbo.Model_Cat  (ModelCat_Code,ModelCat_NAME,Site_code,OldCode, " & _
                                            " U_Name ,U_EntDt,U_AE   ) " & _
                                     "VALUES  ('" & MasterCode & "','" & left(StringPass(RsDms.Fields("Parent Product Line")), 20) & "','" & PubSiteCode & "' ,'', " & _
                                     " 'Siebel','" & Format(PubLoginDate, "Short Date") & "','A')")
      End If

        
        RsDms.MoveNext
    Loop
    
    '' Model Group Updation
    CopyCnt = 0
    ErrorCnt = 0
   
   
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(Right(ModelGrp_Code,4))", "0") & " from Model_Grp Where IsNumeric(Right(ModelGrp_Code,4))=1 ").Fields(0).Value
    If IsNumeric(CodeCnt) Then
        CodeCnt = CodeCnt + 1
    Else
        CodeCnt = 2000
    End If
    If RsDms.RecordCount > 0 Then RsDms.MoveFirst
    Do Until RsDms.EOF
        If IsNull(StringPass(RsDms.Fields("Product Line"))) Or StringPass(RsDms.Fields("Product Line")) = "" Then GoTo MyNextRecord1
        If GCn.Execute("Select ModelGrp_Name from Model_Grp where ModelGrp_Name='" & left(StringPass(RsDms.Fields("Product Line")), 20) & "'").RecordCount > 0 Then GoTo MyNextRecord1
                
        MasterCode = PubDivCode & Right("0000" & CodeCnt, 4)
        mCatCode = GetModelCatCode(left(StringPass(RsDms.Fields("Parent Product Line")), 20))
        mCatCode = GCn.Execute("Select ModelCat_Code from Model_Cat where ModelCat_Name='" & left(StringPass(RsDms.Fields("Parent Product Line")), 20) & "'").Fields(0).Value
        
        'Insert New Rec
       
        
          GCn.Execute ("INSERT INTO dbo.Model_Grp  (ModelGrp_Code,ModelGrp_Name,Wheel_Catg,ModelCat_Code,Site_Code,OldCode, " & _
                                            " U_Name ,U_EntDt,U_AE   ) " & _
                                     "VALUES  ('" & MasterCode & "','" & left(StringPass(RsDms.Fields("Product Line")), 20) & "','Four','" & mCatCode & "','" & PubSiteCode & "' ,'', " & _
                                     " 'Siebel','" & Format(PubLoginDate, "Short Date") & "','A')")


        CodeCnt = CodeCnt + 1
MyNextRecord1:
        CopyCnt = CopyCnt + 1
        
        
        RsDms.MoveNext
    Loop

    CopyCnt = 0
    ErrorCnt = 0
    If RsDms.RecordCount > 0 Then RsDms.MoveFirst
    Do Until RsDms.EOF
        If IsNull(StringPass(RsDms.Fields("Product/VC#"))) Or StringPass(RsDms.Fields("Product/VC#")) = "" Then GoTo MyNextRecord2
        
        If GCn.Execute("Select Model from Model where Model='" & left(StringPass(RsDms.Fields("Product/VC#")), 20) & "' and Div_Code <> '" & PubDivCode & "' ").RecordCount > 0 Then
            GCn.Execute "Update Model Set Div_Code = '' Where Model='" & left(StringPass(RsDms.Fields("Product/VC#")), 20) & "' "
            GoTo MyNextRecord2
        End If
        
        GCn.Execute ("Update Model set UNLADEN_WT='" & left(StringPass(RsDms.Fields("Unladen Weight")), 15) & "' Where Model='" & left(StringPass(RsDms.Fields("Product/VC#")), 20) & "' ")
        If GCn.Execute("Select Model from Model where Model='" & left(StringPass(RsDms.Fields("Product/VC#")), 20) & "'").RecordCount > 0 Then GoTo MyNextRecord2
        
        mColCode = ""
        If IsNull(StringPass(RsDms.Fields("Parent Product Line"))) Or StringPass(RsDms.Fields("Parent Product Line")) = "" Then
            mCatCode = PubDivCode & "XX"
        Else
            mCatCode = GCn.Execute("Select ModelCat_Code from Model_Cat where ModelCat_Name='" & left(StringPass(RsDms.Fields("Parent Product Line")), 20) & "'").Fields(0).Value
        End If
        If IsNull(StringPass(RsDms.Fields("Product Line"))) Or StringPass(RsDms.Fields("Product Line")) = "" Then
            mGrpCode = PubDivCode & "XX"
        Else
            mGrpCode = GCn.Execute("Select ModelGrp_Code from Model_Grp where ModelGrp_Name='" & left(StringPass(RsDms.Fields("Product Line")), 20) & "'").Fields(0).Value
        End If
        If Not IsNull(StringPass(RsDms.Fields("Colour"))) And StringPass(RsDms.Fields("Colour")) <> "" Then
            If GCn.Execute("Select Col_Desc from ColMast where Col_Desc='" & left(Replace(StringPass(RsDms!Colour), "_", " "), 20) & "'").RecordCount > 0 Then
                mColCode = GCn.Execute("Select Col_Code from ColMast where Col_Desc='" & left(Replace(StringPass(RsDms!Colour), "_", " "), 20) & "'").Fields(0).Value
            Else
                 Call CreateErrLog("Model Master", left(StringPass(RsDms.Fields("Product Name")), 40), " " & XNull(RsDms!Colour) & " Colour Name not found in Colour Master during Model Master Creation ")
            End If
        End If
        'Insert New Rec
        
        If IsNull(StringPass(RsDms.Fields("Product Line"))) Or StringPass(RsDms.Fields("Product Line")) = "" Then
          Chas_Type = "."
        Else
            Chas_Type = left(StringPass(RsDms.Fields("Product Line")), 6)
        End If
  
        If IsNull(StringPass(RsDms.Fields("Product Name"))) Or StringPass(RsDms.Fields("Product Name")) = "" Then
           Sales_Desc = StringPass(RsDms.Fields("Product Line"))
        Else
        Sales_Desc = left(StringPass(RsDms.Fields("Product Name")), 40)
        End If
            
        If IsNull(StringPass(RsDms.Fields("Product Description"))) Or StringPass(RsDms.Fields("Product Description")) = "" Then
            Model_Desc = StringPass(RsDms.Fields("Product/VC#"))
            Model_Desc1 = ""
            Model_desc2 = ""
        Else
            Model_Desc = left(StringPass(RsDms.Fields("Product Description")), 50)
            Model_Desc1 = mID(StringPass(RsDms.Fields("Product Description")), 51, 50)
            Model_desc2 = mID(StringPass(RsDms.Fields("Product Description")), 101, 50)
        End If
        
        
          GCn.Execute ("INSERT INTO Model  (Model,Vehicle_Type,Sales_Desc,Chas_Type,Model_desc,Model_desc1,Model_desc2,Grp_Code " & _
                       " ,Cat_Code,Active_YN,TyreDetails,HorsePower,Front_A_Wt,Rear_A_Wt,Unladen_Wt,Gross_Wt " & _
                       " ,WHEELBASE,FuelTankCapacity,RearAxleMake,Cylinder,FUEL,Manufacturer,ServiceTax_YN " & _
                       " ,Col_Code,RegulatoryCertificate,SteeringType,Vehicle_Drive,CubicCapacity,BodyType,Model_Type,Wheel_Catg,RLW,FMSN,Site_Code,Div_Code,U_Name,U_EntDt,U_AE  ) " & _
                       "VALUES  ('" & left(StringPass(RsDms.Fields("Product/VC#")), 20) & "','" & left(StringPass(RsDms!LOB), 5) & "','" & Sales_Desc & "' " & _
                       " ,'" & Chas_Type & " ','" & Model_Desc & "','" & Model_Desc1 & "','" & Model_desc2 & "','" & mGrpCode & "','" & mCatCode & "' " & _
                       " , 1,'" & left(StringPass(RsDms.Fields("Number & Description of Type")), 30) & "' " & _
                       " , '" & left(StringPass(RsDms.Fields("Horse Power")), 10) & "','" & left(StringPass(RsDms.Fields("Front Axle Weight")), 15) & "' " & _
                       " , '" & left(StringPass(RsDms.Fields("Front Axle Weight")), 15) & "','" & left(StringPass(RsDms.Fields("Unladen Weight")), 15) & "' " & _
                       " , '" & left(StringPass(RsDms.Fields("Gross Vehicle Weight")), 15) & "','" & RsDms.Fields("Wheel Base") & "' " & _
                       " , '" & Val(VNull(RsDms.Fields("Fuel Tank"))) & "','" & left(StringPass(RsDms.Fields("Rear Axle")), 30) & "' " & _
                       " , '" & RsDms.Fields("Number of Cylinders") & "', '" & left(StringPass(RsDms.Fields("Fuel")), 10) & "', 'Tata Motors Ltd.' " & _
                       " , 1,'" & mColCode & "', '" & left(StringPass(RsDms.Fields("Regulatory Certification")), 15) & "' " & _
                       " , '" & left(StringPass(RsDms.Fields("Steering")), 20) & "','" & left(StringPass(RsDms.Fields("Vehicle Drive")), 6) & "' " & _
                       " , '" & left(StringPass(RsDms.Fields("Cubic Capacity")), 10) & "','" & left(StringPass(RsDms.Fields("Type of Body")), 25) & "' " & _
                       " , '" & left(StringPass(Chas_Type), 2) & "', 'Four' " & _
                       " , 'XXXX', '" & FMSN & "','" & PubSiteCode & "', '" & PubDivCode & "', 'Siebel', '" & Format(PubLoginDate, "Short Date") & "', 'A')")

        
        
        CodeCnt = CodeCnt + 1
MyNextRecord2:
        CopyCnt = CopyCnt + 1

        
        RsDms.MoveNext
    Loop
    
    GCn.CommitTrans
    

lblExit:
    Set RsNew = Nothing
    Exit Sub
ELoop:
    ErrorCnt = ErrorCnt + 1
    Resume Next
End Sub
Private Sub UnitMasterDataUpdate(Index)
    '' On Error GoTo Eloop

    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    
    Do Until RsDms.EOF
       ErrorCnt = 0
        If IsNull(StringPass(RsDms.Fields("UoM"))) Or StringPass(RsDms.Fields("UoM")) = "" Then ErrorCnt = 1
        
        If GCn.Execute("Select Unit_Name from Unit where Unit_Name='" & StringPass(RsDms.Fields("UoM")) & "'").RecordCount > 0 Then
            Else
               If ErrorCnt = 0 Then
                GCn.Execute ("INSERT INTO dbo.Unit  (Unit_name,Site_Code, U_Name ,U_EntDt,U_AE   ) " & _
                                        "VALUES  ('" & StringPass(RsDms.Fields("UoM")) & "','" & PubSiteCode & "' , " & _
                                        " 'Siebel','" & Format(PubLoginDate, "Short Date") & "','A')")
              End If
        End If

        RsDms.MoveNext
    Loop
    GCn.CommitTrans
    
    
lblExit:
    Exit Sub
ELoop:
    Resume Next
End Sub

Private Sub PartMasterDataUpdate(Index)
'' On Error GoTo Eloop
Dim PartGrade As String
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Do Until RsDms.EOF
    ErrorCnt = 0
        If IsNull(StringPass(RsDms.Fields("Part Number"))) Or StringPass(RsDms.Fields("Part Number")) = "" Then ErrorCnt = 1
        If GCn.Execute("Select Part_No from Part where Part_No='" & StringPass(RsDms.Fields("Part Number")) & "' and Div_Code='" & PubDivCode & "'").RecordCount > 0 Then
            ErrorCnt = 1
        End If
        
        If IsNull(StringPass(RsDms.Fields("Description"))) Or StringPass(RsDms.Fields("Description")) = "" Then
           Call CreateErrLog("Part Master", RsDms.Fields("Part Number"), "  Part Name is Empty ")
            ErrorCnt = 1
        End If
        
        If IsNull(StringPass(RsDms.Fields("UoM"))) Or StringPass(RsDms.Fields("UoM")) = "" Then
            Call CreateErrLog("Part Master", RsDms.Fields("Part Number"), "  Unit is Empty is Empty ")
            ErrorCnt = 1
        End If
        
        If IsNull(StringPass(RsDms.Fields("Vendor"))) Or StringPass(RsDms.Fields("Vendor")) = "" Then
            Call CreateErrLog("Part Master", RsDms.Fields("Part Number"), "  Vender Name is Empty ")
            ErrorCnt = 1
        End If
        Select Case UCase(StringPass(RsDms.Fields("Product Category")))
                Case UCase("Lubricant")
                    PartGrade = "L"
                Case Else
                    PartGrade = "S"
            End Select
           If ErrorCnt = 0 Then
            GCn.Execute ("INSERT INTO dbo.Part  (Part_No,Part_NoHelp,Site_Code,Div_Code,Part_Name,Local_Name,Part_NameHelp,Unit,MARK_YN,Part_OEM,Supl_Loca,Value_Method,Active_YN,Security_Grade,Lead_Time,Disc_Factor,Bin_Loca,Min_Lvl,Max_Lvl,ReOrd_Lvl,Part_Grade,U_Name,U_EntDt,U_AE  ) " & _
                                         "VALUES  ('" & left(RsDms.Fields("Part Number"), 22) & "','" & Replace(left(RsDms.Fields("Part Number"), 22), " ", "") & "','" & PubSiteCode & "' ,'" & PubDivCode & "', " & _
                                         " '" & left(RsDms.Fields("Description"), 40) & "','" & left(RsDms.Fields("Description"), 40) & "', '" & Replace(left(RsDms.Fields("Description"), 40), " ", "") & "' " & _
                                         " , '" & left(RsDms.Fields("UoM"), 6) & "' ,'N','" & RsDms.Fields("Vendor") & "' ,'" & RsDms.Fields("Vendor Location") & "','FIFO' " & _
                                         " ,1,'A','" & Val(StringPass(RsDms.Fields("Lead Time"))) & "' ,'" & StringPass(RsDms.Fields("Discount Code (CVBU)")) & "','' " & _
                                         " ,0,0,0,'" & PartGrade & "','Siebel','" & Format(PubLoginDate, "Short Date") & "','A')")
          End If

        RsDms.MoveNext
    Loop
    GCn.CommitTrans
    
lblExit:
    Set RsNew = Nothing
    Exit Sub
ELoop:
    Resume Next
End Sub




