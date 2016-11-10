VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form FrmCrmDmsImport 
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import From CRM DMS"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   165
      Top             =   6870
   End
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
      Height          =   6465
      Left            =   330
      TabIndex        =   20
      Top             =   1515
      Visible         =   0   'False
      Width           =   11655
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
         Left            =   7695
         TabIndex        =   23
         Top             =   2835
         Width           =   2595
      End
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
         Left            =   7695
         TabIndex        =   22
         Top             =   2370
         Width           =   2595
      End
      Begin MSDataGridLib.DataGrid DgHelp 
         Height          =   1845
         Left            =   5850
         Negotiate       =   -1  'True
         TabIndex        =   21
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
         Height          =   5775
         Left            =   510
         TabIndex        =   33
         Top             =   420
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   10186
         _Version        =   393216
         Tabs            =   4
         Tab             =   2
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Group Parameters"
         TabPicture(0)   =   "FrmHondaDmsImport.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Txt(32)"
         Tab(0).Control(1)=   "Txt(31)"
         Tab(0).Control(2)=   "Txt(30)"
         Tab(0).Control(3)=   "Txt(29)"
         Tab(0).Control(4)=   "Txt(28)"
         Tab(0).Control(5)=   "Txt(27)"
         Tab(0).Control(6)=   "Txt(6)"
         Tab(0).Control(7)=   "Txt(5)"
         Tab(0).Control(8)=   "Txt(4)"
         Tab(0).Control(9)=   "Txt(3)"
         Tab(0).Control(10)=   "Txt(2)"
         Tab(0).Control(11)=   "Lbl(30)"
         Tab(0).Control(12)=   "Lbl(29)"
         Tab(0).Control(13)=   "Lbl(28)"
         Tab(0).Control(14)=   "Lbl(27)"
         Tab(0).Control(15)=   "Lbl(26)"
         Tab(0).Control(16)=   "Lbl(25)"
         Tab(0).Control(17)=   "Lbl(4)"
         Tab(0).Control(18)=   "Lbl(3)"
         Tab(0).Control(19)=   "Lbl(2)"
         Tab(0).Control(20)=   "Lbl(1)"
         Tab(0).Control(21)=   "Lbl(0)"
         Tab(0).ControlCount=   22
         TabCaption(1)   =   "Account Parameters 1"
         TabPicture(1)   =   "FrmHondaDmsImport.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lbl(22)"
         Tab(1).Control(1)=   "Lbl(5)"
         Tab(1).Control(2)=   "Lbl(6)"
         Tab(1).Control(3)=   "Lbl(7)"
         Tab(1).Control(4)=   "Lbl(8)"
         Tab(1).Control(5)=   "Lbl(9)"
         Tab(1).Control(6)=   "Lbl(10)"
         Tab(1).Control(7)=   "Lbl(11)"
         Tab(1).Control(8)=   "Lbl(23)"
         Tab(1).Control(9)=   "Lbl(32)"
         Tab(1).Control(10)=   "Lbl(33)"
         Tab(1).Control(11)=   "Lbl(35)"
         Tab(1).Control(12)=   "Lbl(36)"
         Tab(1).Control(13)=   "Txt(24)"
         Tab(1).Control(14)=   "Txt(7)"
         Tab(1).Control(15)=   "Txt(8)"
         Tab(1).Control(16)=   "Txt(9)"
         Tab(1).Control(17)=   "Txt(10)"
         Tab(1).Control(18)=   "Txt(11)"
         Tab(1).Control(19)=   "Txt(12)"
         Tab(1).Control(20)=   "Txt(13)"
         Tab(1).Control(21)=   "Txt(25)"
         Tab(1).Control(22)=   "Txt(34)"
         Tab(1).Control(23)=   "Txt(35)"
         Tab(1).Control(24)=   "Txt(37)"
         Tab(1).Control(25)=   "Txt(38)"
         Tab(1).ControlCount=   26
         TabCaption(2)   =   "Account Parameters 2"
         TabPicture(2)   =   "FrmHondaDmsImport.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Lbl(12)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Lbl(13)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Lbl(14)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Lbl(18)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Lbl(19)"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Lbl(20)"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "Lbl(21)"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "Lbl(15)"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "Lbl(16)"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "Lbl(17)"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "Lbl(24)"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "Lbl(31)"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "Lbl(34)"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "Txt(14)"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "Txt(15)"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).Control(15)=   "Txt(16)"
         Tab(2).Control(15).Enabled=   0   'False
         Tab(2).Control(16)=   "Txt(20)"
         Tab(2).Control(16).Enabled=   0   'False
         Tab(2).Control(17)=   "Txt(21)"
         Tab(2).Control(17).Enabled=   0   'False
         Tab(2).Control(18)=   "Txt(22)"
         Tab(2).Control(18).Enabled=   0   'False
         Tab(2).Control(19)=   "Txt(23)"
         Tab(2).Control(19).Enabled=   0   'False
         Tab(2).Control(20)=   "Txt(17)"
         Tab(2).Control(20).Enabled=   0   'False
         Tab(2).Control(21)=   "Txt(18)"
         Tab(2).Control(21).Enabled=   0   'False
         Tab(2).Control(22)=   "Txt(19)"
         Tab(2).Control(22).Enabled=   0   'False
         Tab(2).Control(23)=   "Txt(26)"
         Tab(2).Control(23).Enabled=   0   'False
         Tab(2).Control(24)=   "Txt(33)"
         Tab(2).Control(24).Enabled=   0   'False
         Tab(2).Control(25)=   "Txt(36)"
         Tab(2).Control(25).Enabled=   0   'False
         Tab(2).ControlCount=   26
         TabCaption(3)   =   "Other Detail"
         TabPicture(3)   =   "FrmHondaDmsImport.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtgrid1(0)"
         Tab(3).Control(1)=   "FGrid1"
         Tab(3).ControlCount=   2
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   38
            Left            =   -72240
            TabIndex        =   107
            Text            =   "Text1"
            Top             =   3090
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   37
            Left            =   -72240
            TabIndex        =   106
            Text            =   "Text1"
            Top             =   3360
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   32
            Left            =   -72225
            TabIndex        =   69
            Text            =   "Text1"
            Top             =   4560
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   31
            Left            =   -72225
            TabIndex        =   68
            Text            =   "Text1"
            Top             =   3210
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   30
            Left            =   -72225
            TabIndex        =   67
            Text            =   "Text1"
            Top             =   3480
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   29
            Left            =   -72225
            TabIndex        =   66
            Text            =   "Text1"
            Top             =   3750
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   28
            Left            =   -72225
            TabIndex        =   65
            Text            =   "Text1"
            Top             =   4020
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   27
            Left            =   -72225
            TabIndex        =   64
            Text            =   "Text1"
            Top             =   4290
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   -72225
            TabIndex        =   63
            Text            =   "Text1"
            Top             =   2940
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   -72225
            TabIndex        =   62
            Text            =   "Text1"
            Top             =   2670
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   -72225
            TabIndex        =   61
            Text            =   "Text1"
            Top             =   2400
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   -72225
            TabIndex        =   60
            Text            =   "Text1"
            Top             =   2130
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   -72225
            TabIndex        =   59
            Text            =   "Text1"
            Top             =   1860
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   35
            Left            =   -72240
            TabIndex        =   58
            Text            =   "Text1"
            Top             =   1740
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   34
            Left            =   -72240
            TabIndex        =   57
            Text            =   "Text1"
            Top             =   2820
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   25
            Left            =   -72240
            TabIndex        =   56
            Text            =   "Text1"
            Top             =   4440
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   13
            Left            =   -72240
            TabIndex        =   55
            Text            =   "Text1"
            Top             =   2010
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   12
            Left            =   -72240
            TabIndex        =   54
            Text            =   "Text1"
            Top             =   2280
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   11
            Left            =   -72240
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   2550
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   10
            Left            =   -72240
            TabIndex        =   52
            Text            =   "Text1"
            Top             =   3630
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   9
            Left            =   -72240
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   3900
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   8
            Left            =   -72240
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   4170
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   7
            Left            =   -72240
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   1470
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   24
            Left            =   -72240
            TabIndex        =   48
            Text            =   "Text1"
            Top             =   4710
            Width           =   3210
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   36
            Left            =   2865
            TabIndex        =   47
            Text            =   "Text1"
            Top             =   3000
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   33
            Left            =   2865
            TabIndex        =   46
            Text            =   "Text1"
            Top             =   3810
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   26
            Left            =   2865
            TabIndex        =   45
            Text            =   "Text1"
            Top             =   4890
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   19
            Left            =   2865
            TabIndex        =   44
            Text            =   "Text1"
            Top             =   3270
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   18
            Left            =   2865
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   4620
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   17
            Left            =   2865
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   2460
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   23
            Left            =   2865
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   1650
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   22
            Left            =   2865
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   1920
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   21
            Left            =   2865
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   2190
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   20
            Left            =   2865
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   2730
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   16
            Left            =   2865
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   4350
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   15
            Left            =   2865
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   4080
            Width           =   3375
         End
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   14
            Left            =   2865
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   3540
            Width           =   3375
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
            Left            =   -71460
            MaxLength       =   40
            TabIndex        =   34
            Top             =   3315
            Visible         =   0   'False
            Width           =   945
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
            Height          =   1605
            Left            =   -74520
            TabIndex        =   70
            Top             =   2595
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
            Caption         =   "Vat A/c (Purchase)............................."
            Height          =   195
            Index           =   36
            Left            =   -74445
            TabIndex        =   109
            Top             =   3135
            Width           =   3360
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vat 4 % A/c (Purchase)............................."
            Height          =   195
            Index           =   35
            Left            =   -74445
            TabIndex        =   108
            Top             =   3405
            Width           =   3765
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Tax Group.................."
            Height          =   195
            Index           =   30
            Left            =   -74445
            TabIndex        =   105
            Top             =   4590
            Width           =   2685
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Sale Group..................."
            Height          =   195
            Index           =   29
            Left            =   -74430
            TabIndex        =   104
            Top             =   3255
            Width           =   2670
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Purchase Group...."
            Height          =   195
            Index           =   28
            Left            =   -74430
            TabIndex        =   103
            Top             =   3525
            Width           =   2175
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Sale Group.................."
            Height          =   195
            Index           =   27
            Left            =   -74430
            TabIndex        =   102
            Top             =   3795
            Width           =   2715
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Purchase Group........."
            Height          =   195
            Index           =   26
            Left            =   -74430
            TabIndex        =   101
            Top             =   4065
            Width           =   2580
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VAT Group..........................."
            Height          =   195
            Index           =   25
            Left            =   -74430
            TabIndex        =   100
            Top             =   4335
            Width           =   2550
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Creditor Group....."
            Height          =   195
            Index           =   4
            Left            =   -74430
            TabIndex        =   99
            Top             =   2985
            Width           =   2265
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Creditor Group........."
            Height          =   195
            Index           =   3
            Left            =   -74430
            TabIndex        =   98
            Top             =   2715
            Width           =   2400
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Debtor Group........"
            Height          =   195
            Index           =   2
            Left            =   -74430
            TabIndex        =   97
            Top             =   2445
            Width           =   2325
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Workshop Debtor Group...."
            Height          =   195
            Index           =   1
            Left            =   -74430
            TabIndex        =   96
            Top             =   2175
            Width           =   2325
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Debtor Group........."
            Height          =   195
            Index           =   0
            Left            =   -74430
            TabIndex        =   95
            Top             =   1905
            Width           =   2280
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Vat 4% Sale A/c................."
            Height          =   195
            Index           =   33
            Left            =   -74460
            TabIndex        =   94
            Top             =   1785
            Width           =   3000
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vat 4 % A/c (Sale)............................."
            Height          =   195
            Index           =   32
            Left            =   -74445
            TabIndex        =   93
            Top             =   2865
            Width           =   3360
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Labour A/c......................."
            Height          =   195
            Index           =   23
            Left            =   -74460
            TabIndex        =   92
            Top             =   4485
            Width           =   2310
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lubricant Sale A/c..............."
            Height          =   195
            Index           =   11
            Left            =   -74445
            TabIndex        =   91
            Top             =   2055
            Width           =   2460
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Sale A/c................."
            Height          =   195
            Index           =   10
            Left            =   -74445
            TabIndex        =   90
            Top             =   2325
            Width           =   2415
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vat A/c (Sale)............................."
            Height          =   195
            Index           =   9
            Left            =   -74445
            TabIndex        =   89
            Top             =   2595
            Width           =   2955
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Workshop Cash A/c.........."
            Height          =   195
            Index           =   8
            Left            =   -74445
            TabIndex        =   88
            Top             =   3675
            Width           =   2295
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Cash A/c................."
            Height          =   195
            Index           =   7
            Left            =   -74445
            TabIndex        =   87
            Top             =   3945
            Width           =   2370
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Cash A/c..............."
            Height          =   195
            Index           =   6
            Left            =   -74460
            TabIndex        =   86
            Top             =   4215
            Width           =   2355
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Sale A/c................."
            Height          =   195
            Index           =   5
            Left            =   -74445
            TabIndex        =   85
            Top             =   1515
            Width           =   2310
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Tax A/c................"
            Height          =   195
            Index           =   22
            Left            =   -74460
            TabIndex        =   84
            Top             =   4755
            Width           =   2325
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Purchase A/c 4 %........."
            Height          =   195
            Index           =   34
            Left            =   375
            TabIndex        =   83
            Top             =   3030
            Width           =   2640
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Cst Purchase A/c........."
            Height          =   195
            Index           =   31
            Left            =   375
            TabIndex        =   82
            Top             =   3840
            Width           =   2685
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Discount A/c..........................."
            Height          =   195
            Index           =   24
            Left            =   375
            TabIndex        =   81
            Top             =   4920
            Width           =   2700
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Cst Purchase A/c...................."
            Height          =   195
            Index           =   17
            Left            =   360
            TabIndex        =   80
            Top             =   3300
            Width           =   3240
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other Charges A/c.................."
            Height          =   195
            Index           =   16
            Left            =   375
            TabIndex        =   79
            Top             =   4650
            Width           =   2685
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Workshop Bank A/c..........."
            Height          =   195
            Index           =   15
            Left            =   375
            TabIndex        =   78
            Top             =   2490
            Width           =   2355
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local State Name..............."
            Height          =   195
            Index           =   21
            Left            =   375
            TabIndex        =   77
            Top             =   1680
            Width           =   2400
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Bank A/c................"
            Height          =   195
            Index           =   20
            Left            =   375
            TabIndex        =   76
            Top             =   1950
            Width           =   2310
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Bank A/c................."
            Height          =   195
            Index           =   19
            Left            =   375
            TabIndex        =   75
            Top             =   2220
            Width           =   2475
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spare Purchase A/c 12.5 %........."
            Height          =   195
            Index           =   18
            Left            =   375
            TabIndex        =   74
            Top             =   2760
            Width           =   2910
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Round Off A/c........................."
            Height          =   195
            Index           =   14
            Left            =   375
            TabIndex        =   73
            Top             =   4380
            Width           =   2700
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CST A/c................................"
            Height          =   195
            Index           =   13
            Left            =   375
            TabIndex        =   72
            Top             =   4110
            Width           =   2625
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Purchase A/c..............."
            Height          =   195
            Index           =   12
            Left            =   360
            TabIndex        =   71
            Top             =   3570
            Width           =   2700
         End
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
      Left            =   120
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   11625
      Begin VB.TextBox TxtShow 
         Appearance      =   0  'Flat
         Height          =   915
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "FrmHondaDmsImport.frx":0070
         Top             =   1995
         Width           =   8865
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
         TabIndex        =   29
         Top             =   2565
         Width           =   1185
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
         TabIndex        =   28
         Top             =   2565
         Width           =   1185
      End
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
         TabIndex        =   27
         Top             =   15
         Width           =   1170
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FgridErr 
         Height          =   1620
         Left            =   120
         TabIndex        =   26
         Top             =   300
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
      Left            =   10665
      TabIndex        =   16
      Top             =   6135
      Visible         =   0   'False
      Width           =   8985
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
         TabIndex        =   19
         Top             =   1995
         Width           =   1110
      End
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
         TabIndex        =   18
         Top             =   1995
         Width           =   1110
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
         Height          =   1605
         Left            =   120
         TabIndex        =   17
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
   Begin MSComDlg.CommonDialog CD1 
      Left            =   75
      Top             =   6735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   1770
      Left            =   120
      TabIndex        =   5
      Top             =   210
      Visible         =   0   'False
      Width           =   11595
      Begin VB.CommandButton CmdImport 
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
         Index           =   10
         Left            =   5790
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   1200
         Visible         =   0   'False
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
         Left            =   10020
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   285
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
         Left            =   8610
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   285
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
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   285
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Spare Sale Return"
         Height          =   540
         Index           =   7
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4200
         Width           =   1425
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Supplier Payment"
         Height          =   525
         Index           =   4
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4200
         Width           =   1320
      End
      Begin MSComctlLib.ProgressBar Prg 
         Height          =   270
         Left            =   165
         TabIndex        =   12
         Top             =   1350
         Visible         =   0   'False
         Width           =   11280
         _ExtentX        =   19897
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
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
         Left            =   5790
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   285
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
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   285
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
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   285
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   285
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
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   285
         Width           =   1425
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
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   3150
      Begin VB.TextBox Txt 
         Height          =   285
         Index           =   1
         Left            =   1410
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   615
         Width           =   1215
      End
      Begin VB.TextBox Txt 
         Height          =   285
         Index           =   0
         Left            =   1410
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         Height          =   195
         Left            =   375
         TabIndex        =   2
         Top             =   645
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         Height          =   195
         Left            =   390
         TabIndex        =   1
         Top             =   330
         Width           =   900
      End
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   3  'Align Left
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   661
   End
   Begin VB.Label LblTimer 
      Caption         =   "Label3"
      Height          =   480
      Left            =   1800
      TabIndex        =   32
      Top             =   6180
      Width           =   1485
   End
End
Attribute VB_Name = "FrmCrmDmsImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DmsConn As ADODB.Connection
Dim RsDmsEnviro As ADODB.Recordset
Dim RsHelp As ADODB.Recordset
Dim RsCity As ADODB.Recordset
Dim RsState As ADODB.Recordset
Dim RsSubGroup As ADODB.Recordset
Dim RsAcGroup As ADODB.Recordset


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
    Dim RsDms           As ADODB.Recordset
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
                                                    If SprCounterSale(.Fields("Mode Of Payment"), mSubCode, mNetAmount, mSprSaleAmt, mLubeSaleAmt, mVatAmt, mNarr, !Invoice_Date, !Invoice_No, !division, mOtherCharges, mDmsSubCode) = False Then
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
                                                    
                                                    If WorkShopSale(.Fields("Mode Of Payment"), mSubCode, mNetAmount, mSprSaleAmt, mSprSaleVat4Amt, mVat12, mNarr, !Invoice_Date, mNetLabour, mTaxableLabour, mServTaxLabour, !Invoice_No, !division, mOtherCharges, mLabOtherCharges, mDiscount, mLabDiscount, mVat4, mDmsSubCode) = False Then
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
        GCn.Execute "Update DmsEnviro Set SprDebtorGroupCode='" & txt(SprDebtorGroupCode).Tag & "', VehDebtorGroupCode = '" & txt(VehDebtorGroupCode).Tag & "', " & _
                            "WsDebtorGroupCode='" & txt(WsDebtorGroupCode).Tag & "', SprCreditorGroupCode='" & txt(SprCreditorGroupCode).Tag & "', SprSaleAc='" & txt(SprSaleAc).Tag & "', " & _
                            "LubeSaleAc='" & txt(LubSaleAc).Tag & "', SprSaleVat4Ac = '" & txt(SprSaleVat4Ac).Tag & "', VehSaleAc='" & txt(VehSaleAc).Tag & "', SprPurchaseAc='" & txt(SprPurchaseAc).Tag & "', " & _
                            "VehPurchaseAc='" & txt(VehPurchaseAc).Tag & "', SprCashAc='" & txt(SprCashAc).Tag & "', VehCashAc='" & txt(VehCashAc).Tag & "', " & _
                            "WsCashAc='" & txt(WsCashAc).Tag & "', SprBankAc='" & txt(SprBankAc).Tag & "', VehBankAc='" & txt(VehBankAc).Tag & "', WsBankAc='" & txt(WsBankAc).Tag & "', " & _
                            "LocalStateName='" & txt(LocalStateName) & "', LabourAc='" & txt(LabourAc).Tag & "', ServTaxAc='" & txt(ServTaxAc).Tag & "', " & _
                            "VehCreditorGroupCode='" & txt(VehCreditorGroupCode).Tag & "', CstAc = '" & txt(CstAc).Tag & "', VatAc='" & txt(VatAc).Tag & "', " & _
                            "ROffAc='" & txt(ROffAc).Tag & "', SprCstPurchaseAc='" & txt(SprCstPurchaseAc).Tag & "', DiscountAc='" & txt(DiscountAc).Tag & "', " & _
                            "OtherChargesAc='" & txt(OtherChargesAc).Tag & "', VehPurGroupCode='" & txt(VehPurGroupCode).Tag & "', VehSaleGroupCode='" & txt(VehSaleGroupCode).Tag & "', " & _
                            "SprPurGroupCode='" & txt(SprPurGroupCode).Tag & "', SprSaleGroupCode = '" & txt(SprSaleGroupCode).Tag & "', VatGroupCode='" & txt(VatGroupCode).Tag & "', " & _
                            "ServiceTaxGroupCode='" & txt(ServiceTaxGroupCode).Tag & "', VehCstPurchaseAc = '" & txt(VehCstPurchaseAc).Tag & "', VatInputAc = '" & txt(VatInputAc).Tag & "', Vat4InputAc = '" & txt(Vat4InputAc).Tag & "', Vat4Ac = '" & txt(Vat4Ac).Tag & "', SprPurchase4Ac = '" & txt(SprPurchase4Ac).Tag & "' "
    Else
        GCn.Execute "Insert Into DmsEnviro(SprDebtorGroupCode, VehDebtorGroupCode,WsDebtorGroupCode, SprCreditorGroupCode, VehCreditorGroupCode, " & _
                            "SprSaleAc, SprSaleVat4Ac, LubeSaleAc, VehSaleAc, SprPurchaseAc, VehPurchaseAc, " & _
                            "SprCashAc, VehCashAc, WsCashAc, SprBankAc, VehBankAc, " & _
                            "WsBankAc, LocalStateName, LabourAc, ServTaxAc, CstAc, " & _
                            "VatAc, ROffAc, SprCstPurchaseAc, OtherChargesAc, DiscountAc, " & _
                            "VehPurGroupCode, VehSaleGroupCode, SprPurGroupCode, SprSaleGroupCode, VatGroupCode, ServiceTaxGroupCode, VehCstPurchaseAc, Vat4Ac, SprPurchase4Ac, VatInputAc, Vat4InputAc) " & _
                            "Values('" & txt(SprDebtorGroupCode).Tag & "', '" & txt(VehDebtorGroupCode).Tag & "', '" & txt(WsDebtorGroupCode).Tag & "', '" & txt(SprCreditorGroupCode).Tag & "', '" & txt(VehCreditorGroupCode).Tag & "', " & _
                            "'" & txt(SprSaleAc).Tag & "', '" & txt(SprSaleVat4Ac).Tag & "', '" & txt(LubSaleAc).Tag & "', '" & txt(VehSaleAc).Tag & "', '" & txt(SprPurchaseAc).Tag & "', '" & txt(VehPurchaseAc).Tag & "', " & _
                            "'" & txt(SprCashAc).Tag & "', '" & txt(VehCashAc).Tag & "', '" & txt(WsCashAc).Tag & "', '" & txt(SprBankAc).Tag & "', '" & txt(VehBankAc).Tag & "', " & _
                            "'" & txt(WsBankAc).Tag & "', '" & txt(LocalStateName) & "', '" & txt(LabourAc).Tag & "', '" & txt(ServTaxAc).Tag & "', '" & txt(CstAc).Tag & "', " & _
                            "'" & txt(VatAc).Tag & "', '" & txt(ROffAc).Tag & "', '" & txt(SprCstPurchaseAc).Tag & "','" & txt(OtherChargesAc).Tag & "', '" & txt(DiscountAc).Tag & "', " & _
                            "'" & txt(VehPurGroupCode).Tag & "', '" & txt(VehSaleGroupCode).Tag & "', '" & txt(SprPurGroupCode).Tag & "', '" & txt(SprSaleGroupCode).Tag & "', '" & txt(VatGroupCode).Tag & "', '" & txt(ServiceTaxGroupCode).Tag & "', '" & txt(VehCstPurchaseAc).Tag & "', '" & txt(Vat4Ac).Tag & "', '" & txt(SprPurchase4Ac).Tag & "', '" & txt(VatInputAc).Tag & "', '" & txt(Vat4InputAc).Tag & "')"
    End If
    
    GCn.Execute "Delete from DmsBankAc"
    
    With FGrid1
        For I = 1 To .Rows - 1
            If .TextMatrix(I, F1_BankAc) <> "" Then
                GCn.Execute "Insert Into DmsBankAc(AutomanBankCode, DmsBankCode) Values('" & .TextMatrix(I, F1_BankAcCode) & "', '" & .TextMatrix(I, F1_DmsCode) & "')"
            End If
        Next I
    End With
    Unload Me
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
'

    BlankAll
    PubImportData = True
    Call AlignCtrls
    Call Ini_Grid(FgridErr)
    Call Ini_Grid(FGrid)
    Call Ini_Grid(FGrid1)
    
    Set RsDmsEnviro = GCn.Execute("Select * From DmsEnviro")
    If RsDmsEnviro.RecordCount = 0 Then MsgBox "Plz Define Settings In DmsEnviro": Exit Sub
    
    Set RsSubGroup = GCn.Execute("Select SubCode As Code, Name From Subgroup Order By Name")
    Set RsAcGroup = G_FaCn.Execute("Select GroupCode As Code, GroupName As Name From AcGroup Order By GroupName")
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
        WinSetting Me, 5800, 11940
        
    End If
End Sub
Function SprCounterSale(mBillType As String, mPartyCode As String, mNetAmt As Double, mSprSaleAmt As Double, mLubeSaleAmt As Double, mVatAmt As Double, mNarr As String, mDate As Date, mInvoice_No As String, mDmsDivision As String, mOtherCharges As Double, mDmsSubCode As String) As Boolean
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
        G_FaCn.Execute "Update LedgerM Set DmsRefNo='" & mInvoice_No & "', DmsSubCode = '" & mDmsSubCode & "' Where DocId ='" & mDocId & "'"
        
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


Function WorkShopSale(mBillType As String, mPartyCode As String, mNetAmt As Double, mSprSaleAmt As Double, mSprSaleVat4Amt As Double, mVatAmt As Double, mNarr As String, mDate As Date, mTotalLabour As Double, mTaxableLabour As Double, mServiceTax As Double, mInvoice_No As String, mDmsDivision As String, mOtherCharges As Double, mLabOtherCharges As Double, mDiscount As Double, mLabDiscount As Double, mVat4 As Double, mDmsSubCode As String) As Boolean
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
        G_FaCn.Execute "Update LedgerM Set DmsRefNo='" & mInvoice_No & "', DmsSubCode = '" & mDmsSubCode & "' Where DocId ='" & mDocIdSpr & "'"
        
        
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
        'mSubGroupCounter = G_CompCn.Execute("Select " & vIsNull("Max(" & cVal("Right(SubCode,6)") & ")", "0") & "+1 From SubGroup WITH (NOLOCK)").Fields(0)
        'G_FaCn.Execute "Update SubGroupCounter Set SubGroupAcCode=" & mSubGroupCounter & ""
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
                    ROffAc, SprCstPurchaseAc, OtherChargesAc, DiscountAc, VehCstPurchaseAc, SprPurchase4Ac, VatInputAc, Vat4InputAc
             
            Set DgHelp.DataSource = RsHelp
            DgHelp.Move txt(Index).left, txt(Index).top + txt(Index).height + 20
            If RsHelp.RecordCount > 0 Then
                RsHelp.FIND "Code = '" & txt(Index).Tag & "'"
                If RsHelp.EOF = False Then
                    txt(Index) = XNull(RsHelp.Fields("Name"))
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
                    VatGroupCode, ServiceTaxGroupCode, VehCstPurchaseAc, SprPurchase4Ac, VatInputAc, Vat4InputAc
            DGridTxtKeyDown DgHelp, txt, Index, RsHelp, KeyCode, False, 1
    End Select
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case SprDebtorGroupCode, WsDebtorGroupCode, VehDebtorGroupCode, SprCreditorGroupCode, _
                    VehCreditorGroupCode, VehCashAc, SprCashAc, WsCashAc, VehBankAc, SprBankAc, _
                    WsBankAc, VehSaleAc, SprSaleAc, LubSaleAc, VehPurchaseAc, SprPurchaseAc, SprSaleVat4Ac, _
                    LabourAc, ServTaxAc, CstAc, VatAc, Vat4Ac, ROffAc, SprCstPurchaseAc, OtherChargesAc, DiscountAc, _
                    VehSaleGroupCode, VehPurGroupCode, SprSaleGroupCode, SprPurGroupCode, _
                    VatGroupCode, ServiceTaxGroupCode, VehCstPurchaseAc, SprPurchase4Ac, VatInputAc, Vat4InputAc
            If DgHelp.Visible = True Then DGridTxtKeyPress txt, Index, RsHelp, KeyAscii, "Name"
    End Select
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case SprDebtorGroupCode, WsDebtorGroupCode, VehDebtorGroupCode, SprCreditorGroupCode, _
                    VehCreditorGroupCode, VehCashAc, SprCashAc, WsCashAc, VehBankAc, SprBankAc, _
                    WsBankAc, VehSaleAc, SprSaleAc, LubSaleAc, VehPurchaseAc, SprPurchaseAc, SprSaleVat4Ac, _
                    LabourAc, ServTaxAc, CstAc, VatAc, Vat4Ac, ROffAc, SprCstPurchaseAc, OtherChargesAc, _
                    DiscountAc, VehSaleGroupCode, VehPurGroupCode, SprSaleGroupCode, _
                    SprPurGroupCode, VatGroupCode, ServiceTaxGroupCode, VehCstPurchaseAc, SprPurchase4Ac, VatInputAc, Vat4InputAc
            If RsHelp.RecordCount = 0 Or (RsHelp.EOF = True Or RsHelp.BOF = True) Or txt(Index).TEXT = "" Then
                txt(Index).TEXT = ""
                txt(Index).Tag = ""
            Else
                txt(Index).TEXT = RsHelp!Name
                txt(Index).Tag = RsHelp!Code
            End If
        Case ToDate, FromDate
            txt(Index) = RetDate(txt(Index))
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
            AutomanSubcode = ""
            mSubGroupCounter = G_CompCn.Execute("Select SubGroupAcCode From SubGroupCounter With (NOLOCK)").Fields(0)
                        
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
                    
                    G_CompCn.Execute ("Update  SubGroupCounter Set SubGroupAcCode=" & mSubGroupCounter + 1 & " ")
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
            GetAcGroupName XNull(!SprDebtorGroupCode), txt(SprDebtorGroupCode)
            GetAcGroupName XNull(!VehDebtorGroupCode), txt(SprDebtorGroupCode)
            GetAcGroupName XNull(!VehDebtorGroupCode), txt(VehDebtorGroupCode)
            GetAcGroupName XNull(!WsDebtorGroupCode), txt(WsDebtorGroupCode)
            GetAcGroupName XNull(!SprCreditorGroupCode), txt(SprCreditorGroupCode)
            GetAcGroupName XNull(!VehCreditorGroupCode), txt(VehCreditorGroupCode)
            GetAcGroupName XNull(!VehPurGroupCode), txt(VehPurGroupCode)
            GetAcGroupName XNull(!VehSaleGroupCode), txt(VehSaleGroupCode)
            GetAcGroupName XNull(!SprPurGroupCode), txt(SprPurGroupCode)
            GetAcGroupName XNull(!SprSaleGroupCode), txt(SprSaleGroupCode)
            GetAcGroupName XNull(!VatGroupCode), txt(VatGroupCode)
            GetAcGroupName XNull(!ServiceTaxGroupCode), txt(ServiceTaxGroupCode)
            
            
            GetSubName XNull(!SprSaleAc), txt(SprSaleAc)
            GetSubName XNull(!LubeSaleAc), txt(LubSaleAc)
            GetSubName XNull(!SprSaleVat4Ac), txt(SprSaleVat4Ac)
            GetSubName XNull(!SprPurchase4Ac), txt(SprPurchase4Ac)
            GetSubName XNull(!VehSaleAc), txt(VehSaleAc)
            GetSubName XNull(!SprCashAc), txt(SprCashAc)
            GetSubName XNull(!VehCashAc), txt(VehCashAc)
            GetSubName XNull(!WsCashAc), txt(WsCashAc)
            GetSubName XNull(!SprBankAc), txt(SprBankAc)
            GetSubName XNull(!VehBankAc), txt(VehBankAc)
            GetSubName XNull(!WsBankAc), txt(WsBankAc)
            txt(LocalStateName) = XNull(!LocalStateName)
            GetSubName XNull(!SprPurchaseAc), txt(SprPurchaseAc)
            GetSubName XNull(!SprCstPurchaseAc), txt(SprCstPurchaseAc)
            GetSubName XNull(!VehPurchaseAc), txt(VehPurchaseAc)
            GetSubName XNull(!VehCstPurchaseAc), txt(VehCstPurchaseAc)
            GetSubName XNull(!LabourAc), txt(LabourAc)
            GetSubName XNull(!ServTaxAc), txt(ServTaxAc)
            GetSubName XNull(!CstAc), txt(CstAc)
            GetSubName XNull(!VatAc), txt(VatAc)
            GetSubName XNull(!Vat4Ac), txt(Vat4Ac)
            GetSubName XNull(!VatInputAc), txt(VatInputAc)
            GetSubName XNull(!Vat4InputAc), txt(Vat4InputAc)
            GetSubName XNull(!ROffAc), txt(ROffAc)
            GetSubName XNull(!OtherChargesAc), txt(OtherChargesAc)
            GetSubName XNull(!DiscountAc), txt(DiscountAc)
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



Sub Grid_Hide()
    If DgHelp.Visible = True Then DgHelp.Visible = False
End Sub


Sub BlankAll()
    BlankText Me
End Sub
