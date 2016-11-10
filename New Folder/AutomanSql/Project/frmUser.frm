VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "USER PERMISSIONS"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DgUserGroup 
      Height          =   1830
      Left            =   6090
      Negotiate       =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3825
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   3228
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1.25
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "User Group"
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
            ColumnWidth     =   1665.071
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "No"
      Top             =   1890
      Width           =   615
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   661
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFE0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Index           =   1
      Left            =   45
      TabIndex        =   22
      Top             =   2895
      Width           =   11760
      Begin VB.Frame Frame2 
         BackColor       =   &H00CFE0E0&
         Caption         =   "Permissions"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   90
         TabIndex        =   27
         Top             =   1935
         Width           =   4455
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
            Height          =   240
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   1305
            MaxLength       =   10
            TabIndex        =   44
            Top             =   255
            Width           =   2925
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00CFE0E0&
            Caption         =   "Cancel Permission"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3060
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   1590
            Width           =   1200
         End
         Begin VB.CheckBox opt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00CFE0E0&
            Caption         =   "Account"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   3
            Left            =   225
            TabIndex        =   39
            Top             =   1545
            Width           =   1185
         End
         Begin VB.CheckBox opt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00CFE0E0&
            Caption         =   "Setup"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   4
            Left            =   225
            TabIndex        =   38
            Top             =   1845
            Width           =   1185
         End
         Begin VB.CheckBox opt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00CFE0E0&
            Caption         =   "WorkShop"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   2
            Left            =   225
            TabIndex        =   37
            Top             =   1245
            Width           =   1185
         End
         Begin VB.CheckBox opt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00CFE0E0&
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   225
            TabIndex        =   36
            Top             =   960
            Width           =   1185
         End
         Begin VB.CheckBox opt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00CFE0E0&
            Caption         =   "Vehicle"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   225
            TabIndex        =   35
            Top             =   660
            Width           =   1185
         End
         Begin VB.CommandButton CmdAllow 
            BackColor       =   &H00CFE0E0&
            Caption         =   "Add All"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1575
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   660
            Width           =   1035
         End
         Begin VB.CommandButton Cmdrevoke 
            BackColor       =   &H00CFE0E0&
            Caption         =   "Revoke All"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1575
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1845
            Width           =   1035
         End
         Begin VB.CommandButton CmdDel 
            BackColor       =   &H00CFE0E0&
            Caption         =   "View All"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1575
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   1545
            Width           =   1035
         End
         Begin VB.CommandButton CmdEdit 
            BackColor       =   &H00CFE0E0&
            Caption         =   "Delete All"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1575
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1245
            Width           =   1035
         End
         Begin VB.CommandButton Cmdadd 
            BackColor       =   &H00CFE0E0&
            Caption         =   "Edit All"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1575
            MaskColor       =   &H00C0C0FF&
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   1035
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00CFE0E0&
            Caption         =   "Save  Permission"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3060
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1110
            Width           =   1200
         End
         Begin VB.CommandButton CmdFillGroupDetail 
            BackColor       =   &H00CFE0E0&
            Caption         =   "Fill Group"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   3060
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   645
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Group"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   5
            Left            =   225
            TabIndex        =   45
            Top             =   270
            Width           =   975
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FGridPer 
         Height          =   3705
         Left            =   4755
         TabIndex        =   23
         Top             =   435
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   6535
         _Version        =   393216
         Cols            =   9
         BackColorFixed  =   12243913
         BackColorBkg    =   13623520
         Appearance      =   0
         FormatString    =   "||Module           | Options                                                   |Add   |Edit    |Delete |View| "
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Fgrid 
         Height          =   1455
         Left            =   75
         TabIndex        =   26
         Top             =   435
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2566
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13623520
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483646
         BackColorBkg    =   13623520
         FocusRect       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Module Wise Permissions"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4755
         TabIndex        =   41
         Top             =   180
         Width           =   2145
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Site Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   40
         Top             =   195
         Width           =   885
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FGridDupli 
      Height          =   6030
      Left            =   9345
      TabIndex        =   21
      Top             =   7245
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   10636
      _Version        =   393216
      Cols            =   6
      FormatString    =   "|comp|div|form      | Module | Param"
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Setup"
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   4
      Left            =   9975
      TabIndex        =   19
      Top             =   2535
      Width           =   810
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Account"
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   3
      Left            =   8385
      TabIndex        =   18
      Top             =   2520
      Width           =   960
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Workshop"
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   2
      Left            =   6630
      TabIndex        =   17
      Top             =   2520
      Width           =   1140
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Spare"
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   1
      Left            =   5190
      TabIndex        =   16
      Top             =   2520
      Width           =   825
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Vehicle"
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   0
      Left            =   3675
      TabIndex        =   15
      Top             =   2520
      Width           =   900
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "No"
      Top             =   1620
      Width           =   615
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   1
      Top             =   540
      Width           =   1350
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1800
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   810
      Width           =   1350
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1800
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   1350
   End
   Begin MSFlexGridLib.MSFlexGrid FGridComp 
      Height          =   1920
      Left            =   3285
      TabIndex        =   14
      Top             =   420
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   3387
      _Version        =   393216
      Cols            =   17
      BackColorFixed  =   12243913
      BackColorBkg    =   13623520
      GridLinesFixed  =   1
      Appearance      =   0
      FormatString    =   $"frmUser.frx":0000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/c Posting"
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
      Left            =   150
      TabIndex        =   25
      Top             =   1905
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   4
      Left            =   1695
      TabIndex        =   24
      Top             =   1890
      Width           =   45
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Modules / Sections -------------->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   210
      TabIndex        =   20
      Top             =   2565
      Width           =   2670
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Left            =   75
      Top             =   2430
      Width           =   11070
   End
   Begin VB.Shape Shape3 
      Height          =   1920
      Left            =   75
      Top             =   420
      Width           =   3165
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
      Height          =   225
      Index           =   3
      Left            =   1695
      TabIndex        =   13
      Top             =   1620
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   2
      Left            =   1695
      TabIndex        =   12
      Top             =   1095
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   0
      Left            =   1695
      TabIndex        =   11
      Top             =   795
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   1
      Left            =   1695
      TabIndex        =   10
      Top             =   555
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator"
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
      Index           =   3
      Left            =   135
      TabIndex        =   9
      Top             =   1635
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Left            =   120
      TabIndex        =   8
      Top             =   555
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Index           =   1
      Left            =   150
      TabIndex        =   7
      Top             =   825
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
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
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1095
      Width           =   1545
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''FGrid1 Constants'''''''''''''
Private Const F1SelEquip    As Byte = 0
Private Const FSiteName As Byte = 1
Private Const FSiteCode      As Byte = 2


Private Const UserGroups As Byte = 5



Private ADDFLAG As Integer, DNAME As String, rs1 As Recordset, con As String, uname1 As String, setflag As Boolean
'Private Const CtrlBColOrg = &HCFE0E0                       'Orginal BackColour
'Private Const CtrlFColOrg = &H80000008                   'Orginal ForeColour
'Private Const CtrlBCol = &H0&                              'Changed BackColour
'Private Const CtrlFCol = &HFFFF&                              'Changed ForeColour
Dim FillRec As Integer
Dim RsUser As ADODB.Recordset
Dim RsComp As ADODB.Recordset
Dim RsUserGroup As ADODB.Recordset
Dim RsDiv As ADODB.Recordset
Dim RsModule As ADODB.Recordset



Sub Ini_Grid()
    With FGrid
        .Cols = 3
        
        .TextMatrix(0, 0) = "Sel"
        .ColAlignment(0) = flexAlignCenterCenter
        .ColWidth(0) = 500
        
        .TextMatrix(0, FSiteName) = "Site Name"
        .ColAlignment(FSiteName) = flexAlignLeftCenter
        .ColWidth(FSiteName) = 3500
        
        .ColWidth(FSiteCode) = 0
    End With
End Sub



Private Sub CmdFillGroupDetail_Click()
    Dim GRs As ADODB.Recordset
    
    FGridDupli.Rows = 1
    FGridDupli.AddItem ""
    FGridDupli.FixedRows = 1
    
    Set GRs = G_CompCn.Execute("Select  UserGroup1.*,User_Module.name as FormName,User_MODULE.Module_Name as ModuleName " & _
                              "From User_Module  " & _
                              "Left join UserGroup1 on user_module.form_code +user_module.Module_Name=UserGroup1.form_code + UserGroup1.Module_Name " & _
                              "where (UserGroup1.user_name='" & txt(UserGroups) & "' Or UserGroup1.user_name Is Null)  order by user_MODULE.Module_Name,user_MODULE.name")
    Do Until GRs.EOF
        FGridDupli.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & GRs!Form_Code & Chr(9) & GRs!ModuleName & Chr(9) & GRs!param_str
        GRs.MoveNext
    Loop

    GRs.MoveFirst
'    FGridPer.Rows = 1
'    FGridPer.AddItem ""
'    FGridPer.FixedRows = 1
'    FGridPer.Redraw = False
    Do Until GRs.EOF
'        If NAME1 <> GRs!Name Then
'         If XNull(GRs!ModuleName) <> "" Then
'            FGridPer.AddItem "" & Chr(9) & GRs!Form_Code & Chr(9) & GRs!ModuleName & Chr(9) & GRs!FormName
'         End If
'        NAME1 = XNull(GRs!Name)
        param = ""
        For I = 1 To FGridDupli.Rows - 1
            If UCase(FGridDupli.TextMatrix(I, 4)) = UCase(GRs!ModuleName) And UCase(FGridDupli.TextMatrix(I, 3)) = UCase(GRs!Form_Code) Then
                paramval = FGridDupli.TextMatrix(I, 5)
                Exit For
            Else
                paramval = "****"
            End If
        Next
        If paramval <> "" Then
            For I = 1 To FGridPer.Rows - 1
                If UCase(FGridPer.TextMatrix(I, 2)) = UCase(GRs!ModuleName) And UCase(FGridPer.TextMatrix(I, 1)) = UCase(GRs!Form_Code) Then
                    FGridPer.Row = I
                    
                    If paramval <> "****" Then
                        Select Case UCase(FGridPer.TextMatrix(I, 2))
                            Case "ACCOUNT"
                               chk(3).Value = Checked
                            Case "VEHICLE"
                                chk(0).Value = Checked
                            Case "SPARE"
                                chk(1).Value = Checked
                            Case "WORKSHOP"
                                chk(2).Value = Checked
                        End Select
                    End If
                    FGridPer.Col = 4
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, FGridPer.Col) = IIf(mID(paramval, 1, 1) = "*", "", "ü")
                    FGridPer.Col = 5
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, FGridPer.Col) = IIf(mID(paramval, 2, 1) = "*", "", "ü")
                    FGridPer.Col = 6
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, FGridPer.Col) = IIf(mID(paramval, 3, 1) = "*", "", "ü")
                    FGridPer.Col = 7
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, FGridPer.Col) = IIf(mID(paramval, 4, 1) = "*", "", "ü")
                    
                    Exit For
                End If
            Next I
        End If
        GRs.MoveNext
    Loop

'    FGridPer.Redraw = True
End Sub

Private Sub DgUserGroup_Click()
    If RsUserGroup.RecordCount = 0 Or (RsUserGroup.EOF = True Or RsUserGroup.BOF = True) Or txt(Index).TEXT = "" Then
        txt(Index).TEXT = ""
        txt(Index).Tag = ""
    Else
        txt(Index).Tag = XNull(RsUserGroup!Code)
        txt(Index).TEXT = XNull(RsUserGroup!Name)
    End If
    DgUserGroup.Visible = False
End Sub

Private Sub FGrid_Click()

    Dim I       As Integer


    If TopCtrl1.TopText2 <> "Browse" Then
        FGrid.Col = 0
        FGrid.CellFontName = "WINGDINGS"
        FGrid.CellFontSize = 14
        If FGrid.TextMatrix(FGrid.Row, 0) = "" Then
            FGrid.TextMatrix(FGrid.Row, 0) = "ü"
        Else
            FGrid.TextMatrix(FGrid.Row, 0) = ""
        End If
    End If
End Sub

Sub Fill_Site()
On Error GoTo ErrDisp
    Dim TempRs  As ADODB.Recordset
    Dim I       As Integer
        
    If G_CompCn.Execute("Select Count(*) From User_Site Where User_Name='" & txt(0) & "'").Fields(0).Value > 0 Then
        Set TempRs = GCn.Execute _
                  ("Select S.Site_Code, S.Site_Desc, '' As mUserSite From Site S  " & _
                  " ")
    Else
        Set TempRs = GCn.Execute _
                  ("Select S.Site_Code, S.Site_Desc, '' As  mUserSite From Site S  ")
    End If
    I = 1
    FGrid.Rows = 1
    
    If TempRs.RecordCount > 0 Then
        Do Until TempRs.EOF
            FGrid.AddItem ""
            
            FGrid.Col = 0
            FGrid.Row = I
            FGrid.CellFontName = "WINGDINGS"
            FGrid.CellFontSize = 14
            FGrid.TextMatrix(I, 0) = IIf(XNull(G_CompCn.Execute("Select IsNull(Max(Site_Code),'') from user_site where user_Name='" & txt(0) & "' and site_code='" & XNull(TempRs!Site_Code) & "'").Fields(0)) = "", "", "ü")
        
            FGrid.TextMatrix(I, FSiteName) = XNull(TempRs!Site_Desc)
            FGrid.TextMatrix(I, FSiteCode) = XNull(TempRs!Site_Code)
            
            TempRs.MoveNext
            I = I + 1
        Loop
        
        FGrid.FixedRows = 1
    Else
        FGrid.AddItem ""
        FGrid.FixedRows = 1
    End If
    
    Set TempRs = Nothing
    Exit Sub
ErrDisp:
    MsgBox err.Description
End Sub




Private Sub Chk_Click(Index As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION <> "Browse" Then
If chk(Index).Value = Checked Then
    opt(Index).Enabled = True
Else
    For I = 1 To FGridPer.Rows - 1
     If FGridPer.TextMatrix(I, 2) = chk(Index).CAPTION Then
        FGridPer.TextMatrix(I, 4) = ""
        FGridPer.TextMatrix(I, 5) = ""
        FGridPer.TextMatrix(I, 6) = ""
        FGridPer.TextMatrix(I, 7) = ""
     End If
    Next
    opt(Index).Enabled = False
End If
Call Cmd_Enb(True)
End If
End Sub


Private Sub Chk_Validate(Index As Integer, Cancel As Boolean)
Call Cmd_Enb(True)
End Sub

Private Sub Command1_Click()
Call Cmd_Enb(False)
End Sub

Private Sub Command3_Click()
Dim I As Integer
For I = 1 To FGridDupli.Rows - 1
    If I <= FGridDupli.Rows - 1 Then
        If UCase(FGridDupli.TextMatrix(I, 2)) = UCase(FGridComp.TextMatrix(FGridComp.Row, 2)) And UCase(FGridDupli.TextMatrix(I, 1)) = UCase(FGridComp.TextMatrix(FGridComp.Row, 1)) Then
            If FGridDupli.Rows = 2 Then
                FGridDupli.Rows = 1
                Exit For
            Else
                FGridDupli.RemoveItem (I)
                I = I - 1
            End If
        End If
    End If
Next I
For I = 1 To FGridPer.Rows - 1
           FGridDupli.AddItem "" & Chr(9) & FGridComp.TextMatrix(FGridComp.Row, 1) & Chr(9) & FGridComp.TextMatrix(FGridComp.Row, 2) & Chr(9) & FGridPer.TextMatrix(I, 1) & Chr(9) & FGridPer.TextMatrix(I, 2) & Chr(9) & IIf(FGridPer.TextMatrix(I, 4) = "", "*", "A") & IIf(FGridPer.TextMatrix(I, 5) = "", "*", "E") & IIf(FGridPer.TextMatrix(I, 6) = "", "*", "D") & IIf(FGridPer.TextMatrix(I, 7) = "", "*", "P")
Next
FGridComp.TextMatrix(FGridComp.Row, 12) = IIf(chk(0).Value = Checked, 1, 0)
FGridComp.TextMatrix(FGridComp.Row, 13) = IIf(chk(1).Value = Checked, 1, 0)
FGridComp.TextMatrix(FGridComp.Row, 14) = IIf(chk(2).Value = Checked, 1, 0)
FGridComp.TextMatrix(FGridComp.Row, 15) = IIf(chk(3).Value = Checked, 1, 0)
FGridComp.TextMatrix(FGridComp.Row, 16) = IIf(chk(4).Value = Checked, 1, 0)
Call Cmd_Enb(False)
End Sub

Private Sub FGridComp_RowColChange()
Dim I As Integer
If FillRec = 0 Then
    For I = 1 To FGridComp.Rows - 1
        FGridComp.TextMatrix(I, 0) = ""
    Next
     FGridComp.TextMatrix(FGridComp.Row, 0) = "Ü"
     Call Fill_Line(FGridComp.Row)
     Call Set_Dupli(FGridComp.Row)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or (KeyCode = 70 And Shift = 2) Or (KeyCode = 80 And Shift = 2) Or (KeyCode = 83 And Shift = 2) Or KeyCode = vbKeyEscape Or KeyCode = vbKeyF5 Or KeyCode = vbKeyF10 Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyEnd Then TopCtrl1.TopKey_Down KeyCode, Shift
If KeyCode = 27 Then
    Unload Me
End If
Exit Sub
ELoop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
'0|1 Section |2 Module Name|3 Permission
'0|1 Code |2 Division | 3 Vehicle | 4 Spare | 5 WorkShop| 6 Account| 7 SetUp| 8 Permissiion
'0|1 Code |2 Name of Group Company | 3 Start Date |4 Permission
Me.left = 0: Me.top = 0
Dim I As Byte
On Error GoTo err
Set RsUser = New ADODB.Recordset
RsUser.LockType = adLockOptimistic
RsUser.CursorType = adOpenDynamic
Set RsUser = G_CompCn.Execute("select * from UserMast order by user_name")

Set RsUserGroup = G_CompCn.Execute("Select User_Name As Code, User_Name As Name From UserGroup Order By User_Name")
Set DgUserGroup.DataSource = RsUserGroup

Ini_Grid

For I = 0 To 3
    txt(I).BackColor = CtrlBColOrg
    txt(I).ForeColor = CtrlFColOrg
    txt(I).BackColor = CtrlBColOrg
    txt(I).ForeColor = CtrlFColOrg
Next

'    Frame1(1).Top = 975
'    Frame1(1).Left = 600
'
'FGridComp.Left = 30
'FGridComp.Top = 120
'FGridComp.Height = 1140
'FGridComp.Width = 7860
'
'FGridPer.Left = 30
'FGridPer.Top = 1290
'FGridPer.Height = 2175
'FGridPer.Width = 7860

'FGridPer.height = Frame1(1).height - 90
'FGridPer.top = 90
'FGridPer.left = (Frame1(1).left + Frame1(1).width) - (FGridPer.width + 105)
'Me.Top = 0
'Me.Left = 0
Me.height = 7635
Me.width = 11940
'FGridPer.FormatString = "||Section Name |Module/Form Name                                                    |Allow  |Add    |Edit    |Delete|"
'FGridComp.FormatString = "|||Name of Group Company                             |Division                                |Start Date       |Permission"

    'If PubULabel = "Y" Then
    TopCtrl1.Tag = PubUParam
    FGridPer.ColWidth(0) = 200
    FGridPer.ColWidth(1) = 0
    FGridPer.ColWidth(8) = 0
        
    FGridComp.ColWidth(0) = 400
    FGridComp.ColWidth(1) = 0
    FGridComp.ColWidth(2) = 0
    For I = 7 To 16
     FGridComp.ColWidth(I) = 0
    Next
    Disp_Text SETS("INI", Me, RsUser)
    Call MoveRec
    Exit Sub
err:
  CheckError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Opt_Click(Index As Integer)
Call Cmd_Enb(True)
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If RsUser.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "SELECT user_Name as SearchCode, User_Name FROM UserMast  order by User_Name"
    Set SearchForm = Me
    MultiComp = True
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    MultiComp = False
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub

End Sub

Private Sub Txt_GotFocus(Index As Integer)
    Call Ctrl_GetFocus(Index)
    Select Case Index
        Case UserGroups
            DgUserGroup.Move txt(UserGroups).left, txt(UserGroups).top + txt(UserGroups).height + 30
            If RsUserGroup.RecordCount = 0 Or (RsUserGroup.EOF = True Or RsUserGroup.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
            If txt(Index).TEXT <> RsUserGroup!Name Then
                RsUserGroup.MoveFirst
                RsUserGroup.FIND "Name ='" & txt(Index).TEXT & "'"
            End If
        
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case UserGroups
        DGridTxtKeyDown DgUserGroup, txt, Index, RsUserGroup, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
End Select

If DgUserGroup.Visible = False Then
    If KeyCode = 13 Then SendKeysA vbKeyTab, True
    If KeyCode = 40 And Index <> 4 Then   'keydown = 40
        SendKeysA vbKeyTab, True
    ElseIf KeyCode = 38 And ADDFLAG = 1 And Index <> 0 Then    'keyup = 38
        SendKeys "+{Tab}"
    ElseIf KeyCode = 38 And ADDFLAG = 2 And Index <> 1 Then    'keyup = 38
        SendKeys "+{Tab}"
    End If
    If KeyCode = 40 And Index = 4 Then
        FGridComp.SetFocus
    End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Call CheckQuote(KeyAscii)
    
    Select Case Index
        Case UserGroups
            If DgUserGroup.Visible = True Then DGridTxtKeyPress txt, Index, RsUserGroup, KeyAscii, "Name"
    End Select
End Sub
Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case 3, 4
        If Len(txt(Index)) = 0 Or UCase(mID(txt(Index), 1, 1)) = "N" Then
            txt(Index) = "No"
        ElseIf UCase(mID(txt(Index), 1, 1)) = "Y" Then
            txt(Index) = "Yes"
        Else
            txt(Index) = "No"
        End If
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Call Ctrl_validate(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
If txt(Index).TEXT = "" Then
    MsgBox Label1(Index) & " Is Required", vbExclamation, "Validation Check"
'    txt(Index).SetFocus
    Cancel = True
    Exit Sub
End If
Select Case Index
    Case 2
        If txt(2).TEXT <> txt(1).TEXT Then
            MsgBox "Please Retype Password For Confirmation", vbExclamation, "Validation Check"
            txt(2).TEXT = ""
            Cancel = True
            Exit Sub
        End If
    Case 4
        If txt(3) <> "Yes" Then
            txt(Index) = "No"
        End If
    Case UserGroups
        If RsUserGroup.RecordCount = 0 Or (RsUserGroup.EOF = True Or RsUserGroup.BOF = True) Or txt(Index).TEXT = "" Then
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
        Else
            txt(Index).Tag = XNull(RsUserGroup!Code)
            txt(Index).TEXT = XNull(RsUserGroup!Name)
        End If
        DgUserGroup.Visible = False
End Select
End Sub
Private Sub TopCtrl1_eFirst()
  BUTTONS True, Me, RsUser, 1
  Call MoveRec
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo eloop1
    Dim I As Integer
    Dim GRs As Recordset
    Dim CName As String
    Disp_Text SETS("ADD", Me, RsUser)
    ADDFLAG = 1
    txt(0) = ""
    txt(1) = ""
    txt(2) = ""
    txt(3) = "No"
    txt(4) = "No"
    FillRec = 1
    FGridComp.Rows = 1
    Set GRs = New Recordset
    GRs.Open "SELECT user1.param_str,Company.start_date ,Company.Comp_Name, User1.comp_code,User1.Div_Name,User1.div_code,User1.mod_veh,User1.mod_spr,User1.mod_wsp,User1.mod_acc,User1.mod_set FROM User1 LEFT JOIN Company ON User1.Comp_Code = Company.Comp_Code where user1.User_name = 'SA' and user1.comp_code='" & PubCenCompCode & "' Order by Company.Comp_Name,User1.Div_Code ", G_CompCn, adOpenStatic, adLockReadOnly
    Do Until GRs.EOF
        FGridComp.AddItem "" & Chr(9) & GRs!Comp_Code & Chr(9) & GRs!Div_Code & Chr(9) & GRs!Comp_Name & Chr(9) & GRs!Div_Name & Chr(9) & GRs!Start_Date & Chr(9) & "" & Chr(9) & GRs!mod_veh & Chr(9) & GRs!mod_spr & Chr(9) & GRs!mod_wsp & Chr(9) & GRs!mod_acc & Chr(9) & GRs!mod_set
        GRs.MoveNext
    Loop
    For I = 1 To FGridComp.Rows - 1
        FGridComp.Col = 0
        FGridComp.Row = I
        FGridComp.CellFontName = "wingdings"
        FGridComp.CellFontSize = 16
        FGridComp.CellForeColor = vbBlue
    Next
    FGridComp.Row = 1
    FGridComp.TextMatrix(1, 0) = "Ü"
    FGridDupli.Rows = 1
    chk(0).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 7)) = 1, True, False)
    chk(1).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 8)) = 1, True, False)
    chk(2).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 9)) = 1, True, False)
    chk(3).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 10)) = 1, True, False)
    chk(4).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 11)) = 1, True, False)
    chk(0).Value = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 12)) = 1, Checked, Unchecked)
    chk(1).Value = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 13)) = 1, Checked, Unchecked)
    chk(2).Value = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 14)) = 1, Checked, Unchecked)
    chk(3).Value = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 15)) = 1, Checked, Unchecked)
    chk(4).Value = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 16)) = 1, Checked, Unchecked)
    opt(0).Enabled = IIf(chk(0).Value = Checked, True, False)
    opt(1).Enabled = IIf(chk(1).Value = Checked, True, False)
    opt(2).Enabled = IIf(chk(2).Value = Checked, True, False)
    opt(3).Enabled = IIf(chk(3).Value = Checked, True, False)
    opt(4).Enabled = IIf(chk(4).Value = Checked, True, False)
    FGridPer.Rows = 1
    Set GRs = New Recordset
    'GRs.Open "select user2.*,User_Module.name as FormName,User_MODULE.Module_Name as ModuleName from user2 left join user_module on user_module.form_code +user_module.Module_Name=user2.form_code + user2.Module_Name where  user2.user_name='SA' and (user2.comp_code='" & PubCenCompCode & "' Or user2.comp_code Is Null) and user2.div_code='" & FGridComp.TextMatrix(FGridComp.Row, 2) & "' order by user_MODULE.Module_Name,user_MODULE.name", G_CompCn, adOpenStatic, adLockReadOnly
    GRs.Open "select user_module.*,User_Module.name as FormName,User_MODULE.Module_Name as ModuleName from user_Module order by user_MODULE.Module_Name,user_MODULE.name", G_CompCn, adOpenStatic, adLockReadOnly
    Do Until GRs.EOF
        If XNull(GRs!ModuleName) <> "" Then
            FGridPer.AddItem "" & Chr(9) & GRs!Form_Code & Chr(9) & GRs!ModuleName & Chr(9) & GRs!FormName
        End If
        GRs.MoveNext
    Loop
    txt(0).SetFocus
    setflag = False
    FillRec = 0
    Call Fill_Site
    Exit Sub
eloop1:
     If err.NUMBER <> 0 Then
        MsgBox err.Description, vbInformation, "Information"
    End If
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo eloop1
        ADDFLAG = 0
        If UCase(txt(0)) = "SA" Then MsgBox "SA Cannot Be Deleted.", vbInformation, "Information": Exit Sub
        If PubULabel = "Y" Then
            If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                G_CompCn.BeginTrans
                G_CompCn.Execute ("delete from user2 where user_name='" & RsUser!user_name & "' ")
                G_CompCn.Execute ("delete from user1 where user_name='" & RsUser!user_name & "'")
                G_CompCn.Execute ("delete from UserMast where user_name='" & RsUser!user_name & "'")
                G_CompCn.CommitTrans
                RsUser.Requery
                Call MoveRec
                BUTTONS True, Me, RsUser, 0
            End If
        Else
            MsgBox "Only SA Can Delete Any User.", vbInformation, "Information"
            Exit Sub
        End If
eloop1:
    If err.NUMBER <> 0 Then
       GCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
    End If
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    ADDFLAG = 2
    Disp_Text SETS("EDIT", Me, RsUser)
'    FGridComp.Row = 1
    chk(0).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 7)) = 1, True, False)
    chk(1).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 8)) = 1, True, False)
    chk(2).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 9)) = 1, True, False)
    chk(3).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 10)) = 1, True, False)
    chk(4).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 11)) = 1, True, False)
    opt(0).Enabled = IIf(chk(0).Value = Checked, True, False)
    opt(1).Enabled = IIf(chk(1).Value = Checked, True, False)
    opt(2).Enabled = IIf(chk(2).Value = Checked, True, False)
    opt(3).Enabled = IIf(chk(3).Value = Checked, True, False)
    opt(4).Enabled = IIf(chk(4).Value = Checked, True, False)
    txt(0).Enabled = False
    DNAME = txt(0)
    setflag = False
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    RsUser.Cancel
    Unload Me
End Sub

Private Sub TopCtrl1_eLast()
 BUTTONS True, Me, RsUser, 4
 Call MoveRec
End Sub

Private Sub TopCtrl1_eNext()
 BUTTONS True, Me, RsUser, 3
 Call MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
 BUTTONS True, Me, RsUser, 2
 Call MoveRec
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ErrorLoop
    If TopCtrl1.TopText2.CAPTION = "Add" Then Call MoveRec
    Call SETS("INI", Me, RsUser)
    Call Cmd_Enb(False)
    Call MoveRec
    ADDFLAG = 0
    Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_eSave()
    Dim I As Boolean, j As Integer, mTrans As Boolean
    Dim K As Integer
'    On Error GoTo errlbl
    If ADDFLAG = 1 And UCase(txt(0)) = "SA" Then MsgBox "You can't create User Name SA !!", vbInformation, "Validation Check": Exit Sub
    If Command3.Enabled = True Then MsgBox "First Save/Cancel Permission", vbInformation, "Validation Check ": Command3.SetFocus: Exit Sub
    If IsValid(txt(0), "User Name") = False Then Exit Sub
    If txt(1) <> txt(2) Then MsgBox "Password Not Confirmed", vbCritical, "Validation Message": Exit Sub
    If ADDFLAG = 2 And UCase(txt(0)) = "SA" Then
        G_CompCn.BeginTrans
        mTrans = True
        G_CompCn.Execute "update UserMast set PASSWD='" & CODIFY(RTrim(txt(1))) & "' where user_name = 'SA'"
        G_CompCn.CommitTrans
        mTrans = False
        DNAME = txt(0)
        RsUser.Requery
        uname1 = txt(0).TEXT
        setflag = True
        ADDFLAG = 0
        RsUser.FIND "user_name = 'SA'"
        Disp_Text SETS("INI", Me, RsUser)
        Call MoveRec
        Call Cmd_Enb(False)
        Exit Sub
    End If
    
    If ADDFLAG = 1 Then
        If G_CompCn.Execute("select count(*) from UserMast where user_name='" & txt(0) & "'").Fields(0) > 0 Then MsgBox "Duplicate User Name", vbCritical, "Validation Error": Exit Sub
    Else
        If txt(0) <> DNAME Then
            If G_CompCn.Execute("select count(*) from UserMast where user_name='" & txt(0) & "'").Fields(0) > 0 And DNAME <> RTrim(txt(0)) Then MsgBox "Duplicate User Name", vbCritical, "Validation Error": Exit Sub
        End If
    End If
    G_CompCn.BeginTrans
    mTrans = True
    G_CompCn.Execute ("delete from user1 where user_name='" & txt(0) & "' and comp_code='" & PubCenCompCode & "'")
    G_CompCn.Execute ("delete from user2 where user_name='" & txt(0) & "' and comp_code='" & PubCenCompCode & "'")
    If ADDFLAG = 1 Then
        G_CompCn.Execute "insert into UserMast(user_name,PASSWD,Label,AcPosting) values('" & txt(0) & "','" & CODIFY(RTrim(txt(1))) & "','" & IIf(txt(3).TEXT = "Yes", "Y", "N") & "','" & IIf(txt(4).TEXT = "Yes", "Y", "N") & "')"
    Else
        G_CompCn.Execute "update UserMast set PASSWD='" & CODIFY(RTrim(txt(1))) & "',Label='" & IIf(txt(3).TEXT = "Yes", "Y", "N") & "',AcPosting='" & IIf(txt(4).TEXT = "Yes", "Y", "N") & "' where user_name = '" & txt(0) & "'"
    End If
    For j = 1 To FGridComp.Rows - 1
        If Trim(FGridComp.TextMatrix(j, 6)) <> "" Then G_CompCn.Execute ("insert into user1(user_name,comp_code,div_code,div_name,mod_veh,mod_spr,mod_wsp,mod_acc,mod_set,param_str) values('" & txt(0) & "','" & FGridComp.TextMatrix(j, 1) & "','" & FGridComp.TextMatrix(j, 2) & "','" & FGridComp.TextMatrix(j, 4) & "'," & Val(FGridComp.TextMatrix(j, 12)) & "," & Val(FGridComp.TextMatrix(j, 13)) & "," & Val(FGridComp.TextMatrix(j, 14)) & "," & Val(FGridComp.TextMatrix(j, 15)) & "," & Val(FGridComp.TextMatrix(j, 16)) & ",'" & IIf(FGridComp.TextMatrix(j, 6) = "", "", "*") & "')")
    Next
    For j = 1 To FGridDupli.Rows - 1
        If FGridDupli.TextMatrix(j, 1) <> "" Then G_CompCn.Execute ("insert into user2(user_name,comp_code,div_code,form_code,Module_Name,param_str) values('" & txt(0) & "','" & FGridDupli.TextMatrix(j, 1) & "','" & FGridDupli.TextMatrix(j, 2) & "','" & FGridDupli.TextMatrix(j, 3) & "','" & FGridDupli.TextMatrix(j, 4) & "','" & IIf(mID(FGridDupli.TextMatrix(j, 5), 1, 1) = "*", "*", "A") & IIf(mID(FGridDupli.TextMatrix(j, 5), 2, 1) = "*", "*", "E") & IIf(mID(FGridDupli.TextMatrix(j, 5), 3, 1) = "*", "*", "D") & IIf(mID(FGridDupli.TextMatrix(j, 5), 4, 1) = "*", "*", "P") & "')")
    Next
    
    G_CompCn.Execute "Delete From User_Site Where User_Name='" & txt(0) & "'"
    For K = 1 To FGrid.Rows - 1
        If FGrid.TextMatrix(K, 0) <> "" Then
            G_CompCn.Execute "Insert Into User_Site (Site_Code, User_Name, Comp_Code) Values ('" & FGrid.TextMatrix(K, FSiteCode) & "', '" & txt(0) & "', '" & PubCenCompCode & "')"
        End If
    Next K
    G_CompCn.CommitTrans
    mTrans = False

    DNAME = txt(0)
    RsUser.Requery
    uname1 = txt(0).TEXT
    RsUser.Requery
    setflag = True
    ADDFLAG = 0
    RsUser.FIND "user_name = '" & uname1 & "'"
    Disp_Text SETS("INI", Me, RsUser)
    Call MoveRec
    Call Cmd_Enb(False)
    Exit Sub
errlbl:
    If mTrans Then G_CompCn.RollbackTrans
    MsgBox CStr(err.NUMBER) & " : " & err.Description, vbCritical, "User Creation Failed"
    Exit Sub
End Sub
Private Sub CmdEdit_Click()
   If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    For I = 1 To FGridPer.Rows - 1
        If FGridPer.TextMatrix(I, 3) <> "" Then
            If opt(0).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Vehicle" Then
                    FGridPer.Col = 6
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 6) = "ü"
            End If
            End If
            If opt(1).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Spare" Then
                    FGridPer.Col = 6
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 6) = "ü"
                End If
            End If
            If opt(2).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Workshop" Then
                    FGridPer.Col = 6
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 6) = "ü"
                End If
            End If
            If opt(3).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Account" Then
                    FGridPer.Col = 6
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 6) = "ü"
                End If
            End If
            If opt(4).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Setup" Then
                    FGridPer.Col = 6
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 6) = "ü"
                End If
            End If
        End If
    Next
End Sub
Private Sub Cmddel_Click()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    For I = 1 To FGridPer.Rows - 1
        If FGridPer.TextMatrix(I, 3) <> "" Then
            If opt(0).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Vehicle" Then
                    FGridPer.Col = 7
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 7) = "ü"
                End If
            End If
            If opt(1).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Spare" Then
                    FGridPer.Col = 7
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 7) = "ü"
                End If
            End If
            If opt(2).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Workshop" Then
                    FGridPer.Col = 7
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 7) = "ü"
                End If
            End If
            If opt(3).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Account" Then
                    FGridPer.Col = 7
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 7) = "ü"
                End If
            End If
            If opt(4).Value = Checked Then
                If FGridPer.TextMatrix(I, 2) = "Setup" Then
                    FGridPer.Col = 7
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 7) = "ü"
                End If
            End If
        End If
    Next
End Sub

Private Sub CmdAllow_Click()
     If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    For I = 1 To FGridPer.Rows - 1
        If opt(0).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Vehicle" Then
                FGridPer.Col = 4
                FGridPer.Row = I
                FGridPer.CellFontName = "wingdings"
                FGridPer.CellFontSize = 18
                FGridPer.CellForeColor = vbBlue
                FGridPer.TextMatrix(I, 4) = "ü"
            End If
        End If
        If opt(1).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Spare" Then
                FGridPer.Col = 4
                FGridPer.Row = I
                FGridPer.CellFontName = "wingdings"
                FGridPer.CellFontSize = 18
                FGridPer.CellForeColor = vbBlue
                FGridPer.TextMatrix(I, 4) = "ü"
            End If
        End If
       If opt(2).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Workshop" Then
                FGridPer.Col = 4
                FGridPer.Row = I
                FGridPer.CellFontName = "wingdings"
                FGridPer.CellFontSize = 18
                FGridPer.CellForeColor = vbBlue
                FGridPer.TextMatrix(I, 4) = "ü"
            End If
        End If
       If opt(3).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Account" Then
                FGridPer.Col = 4
                FGridPer.Row = I
                FGridPer.CellFontName = "wingdings"
                FGridPer.CellFontSize = 18
                FGridPer.CellForeColor = vbBlue
                FGridPer.TextMatrix(I, 4) = "ü"
            End If
        End If
        If opt(4).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Setup" Then
                FGridPer.Col = 4
                FGridPer.Row = I
                FGridPer.CellFontName = "wingdings"
                FGridPer.CellFontSize = 18
                FGridPer.CellForeColor = vbBlue
                FGridPer.TextMatrix(I, 4) = "ü"
            End If
        End If
    Next
End Sub
Private Sub Cmdadd_Click()
Dim I As Integer
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    For I = 1 To FGridPer.Rows - 1
        If FGridPer.TextMatrix(I, 3) <> "" Then
        If opt(0).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Vehicle" Then
                    FGridPer.Col = 5
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 5) = "ü"
            End If
        End If
        If opt(1).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Spare" Then
                    FGridPer.Col = 5
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 5) = "ü"
            End If
        End If
        If opt(2).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Workshop" Then
                    FGridPer.Col = 5
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 5) = "ü"
            End If
        End If
        If opt(3).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Account" Then
                    FGridPer.Col = 5
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 5) = "ü"
            End If
        End If
        If opt(4).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Setup" Then
                    FGridPer.Col = 5
                    FGridPer.Row = I
                    FGridPer.CellFontName = "wingdings"
                    FGridPer.CellFontSize = 18
                    FGridPer.CellForeColor = vbBlue
                    FGridPer.TextMatrix(I, 5) = "ü"
            End If
        End If
        End If
    Next
End Sub

Private Sub CmdRevoke_Click()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    For I = 1 To FGridPer.Rows - 1
        If opt(0).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Vehicle" Then
               FGridPer.TextMatrix(I, 4) = ""
               FGridPer.TextMatrix(I, 5) = ""
               FGridPer.TextMatrix(I, 6) = ""
               FGridPer.TextMatrix(I, 7) = ""
            End If
        End If
        If opt(1).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Spare" Then
               FGridPer.TextMatrix(I, 4) = ""
               FGridPer.TextMatrix(I, 5) = ""
               FGridPer.TextMatrix(I, 6) = ""
               FGridPer.TextMatrix(I, 7) = ""
            End If
        End If
       If opt(2).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Workshop" Then
               FGridPer.TextMatrix(I, 4) = ""
               FGridPer.TextMatrix(I, 5) = ""
               FGridPer.TextMatrix(I, 6) = ""
               FGridPer.TextMatrix(I, 7) = ""
            End If
        End If
        If opt(3).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Account" Then
               FGridPer.TextMatrix(I, 4) = ""
               FGridPer.TextMatrix(I, 5) = ""
               FGridPer.TextMatrix(I, 6) = ""
               FGridPer.TextMatrix(I, 7) = ""
            End If
        End If
        If opt(4).Value = Checked Then
            If FGridPer.TextMatrix(I, 2) = "Setup" Then
               FGridPer.TextMatrix(I, 4) = ""
               FGridPer.TextMatrix(I, 5) = ""
               FGridPer.TextMatrix(I, 6) = ""
               FGridPer.TextMatrix(I, 7) = ""
            End If
       End If
    Next
End Sub

Private Sub FGridPer_Click()
    If FGridPer.Col = 0 Or FGridPer.Col = 1 Or FGridPer.Col = 2 Or FGridPer.Col = 3 Or TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGridPer.Col = 0 Or FGridPer.Col = 1 Or FGridPer.Col = 2 Or FGridPer.Col = 3 Then Exit Sub
    If FGridPer.TextMatrix(FGridPer.Row, 3) = "" Then Exit Sub
    If FGridPer.TextMatrix(FGridPer.Row, 3) = "" Then Exit Sub
    If FGridPer.TextMatrix(FGridPer.Row, 2) = "Vehicle" And chk(0).Value = Unchecked Then Exit Sub
    If FGridPer.TextMatrix(FGridPer.Row, 2) = "Spare" And chk(1).Value = Unchecked Then Exit Sub
    If FGridPer.TextMatrix(FGridPer.Row, 2) = "Workshop" And chk(2).Value = Unchecked Then Exit Sub
    If FGridPer.TextMatrix(FGridPer.Row, 2) = "Account" And chk(3).Value = Unchecked Then Exit Sub
    If FGridPer.TextMatrix(FGridPer.Row, 2) = "Setup" And chk(4).Value = Unchecked Then Exit Sub
    
    If FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = "" Then
        FGridPer.Col = FGridPer.Col
        FGridPer.CellFontName = "wingdings"
        FGridPer.CellFontSize = 18
        If FGridPer.Col = 4 Then
            FGridPer.CellForeColor = vbBlue
        Else
            FGridPer.CellForeColor = vbBlue
        End If
        FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = "ü"
    Else
        FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = ""
        If FGridPer.Col = 4 Then
            FGridPer.TextMatrix(FGridPer.Row, 5) = ""
            FGridPer.TextMatrix(FGridPer.Row, 6) = ""
            FGridPer.TextMatrix(FGridPer.Row, 7) = ""
        End If
    End If
Call Cmd_Enb(True)
End Sub

Private Sub FGridPer_KeyPress(KeyAscii As Integer)
    If FGridPer.Col = 0 Or FGridPer.Col = 1 Or TopCtrl1.TopText2.CAPTION = "Browse" Or FGridPer.TextMatrix(FGridPer.Row, 0) = "" Or KeyAscii <> 32 Then Exit Sub
    If FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = "" Then
        FGridPer.Col = FGridPer.Col
        FGridPer.CellFontName = "wingdings"
        FGridPer.CellFontSize = 18
        FGridPer.CellForeColor = vbBlue
        FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = "ü"
    Else
        FGridPer.TextMatrix(FGridPer.Row, FGridPer.Col) = ""
    End If
End Sub

Private Sub FGridComp_Click()
Dim I As Integer
    If FGridComp.Col = 1 Or FGridComp.Col = 5 Or FGridComp.Col = 4 Or FGridComp.Col = 2 Or FGridComp.Col = 3 Or TopCtrl1.TopText2.CAPTION = "Browse" Or FGridComp.TextMatrix(FGridComp.Row, 1) = "" Or UCase(txt(0)) = "SA" Then Exit Sub
    If FGridComp.Col = 1 Or FGridComp.Col = 5 Or FGridComp.Col = 4 Or FGridComp.Col = 2 Or FGridComp.Col = 3 Or FGridComp.TextMatrix(FGridComp.Row, 1) = "" Or UCase(txt(0)) = "SA" Then Exit Sub
    If FGridComp.TextMatrix(FGridComp.Row, FGridComp.Col) = "" Then
        FGridComp.CellFontName = "wingdings"
        FGridComp.CellFontSize = 18
        FGridComp.CellForeColor = vbBlue
        FGridComp.TextMatrix(FGridComp.Row, FGridComp.Col) = "ü"
    Else
        If MsgBox("Are You Sure To Remove Permission ?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            FGridComp.TextMatrix(FGridComp.Row, FGridComp.Col) = ""
            FGridComp.TextMatrix(FGridComp.Row, 12) = 0
            FGridComp.TextMatrix(FGridComp.Row, 13) = 0
            FGridComp.TextMatrix(FGridComp.Row, 14) = 0
            FGridComp.TextMatrix(FGridComp.Row, 15) = 0
            FGridComp.TextMatrix(FGridComp.Row, 16) = 0
    
            For I = 0 To 4
                chk(I).Value = Unchecked
            Next
            For I = 1 To FGridPer.Rows - 1
                FGridPer.TextMatrix(I, 4) = ""
                FGridPer.TextMatrix(I, 5) = ""
                FGridPer.TextMatrix(I, 6) = ""
                FGridPer.TextMatrix(I, 7) = ""
            Next
            For I = 1 To FGridDupli.Rows - 1
                If I <= FGridDupli.Rows - 1 Then
                    If UCase(FGridDupli.TextMatrix(I, 2)) = UCase(FGridComp.TextMatrix(FGridComp.Row, 2)) And UCase(FGridDupli.TextMatrix(I, 1)) = UCase(FGridComp.TextMatrix(FGridComp.Row, 1)) Then
                        If FGridDupli.Rows = 2 Then
                            FGridDupli.Rows = 1
                            Exit For
                        Else
                            FGridDupli.RemoveItem (I)
                            I = I - 1
                        End If
                    End If
                End If
            Next I
        End If
    End If
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        RsUser.MoveFirst
        RsUser.FIND ("User_Name='" & MyValue & "'")
    Else
        Set RsUser = G_CompCn.Execute("select * from UserMast Where User_Name = '" & MyValue & "' order by user_name")
     End If
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub MoveRec()
On Error GoTo ELoop
Dim GRs As ADODB.Recordset
Dim Rs As Recordset, rs1 As Recordset, Name1 As String
FillRec = 1
'FGridPer.Redraw = False
'FGridComp.Redraw = False
FGridDupli.Rows = 1
FGridPer.Rows = 1
FGridComp.Rows = 1
If RsUser.RecordCount > 0 Then
    For I = 0 To 4
        chk(I).Value = Unchecked
    Next
    For I = 0 To 4
        opt(I).Enabled = False
    Next
    txt(0) = RsUser!user_name
    txt(1) = IIf(IsNull(RsUser!PASSWD), "", DCODIFY(RTrim(IIf(IsNull(RsUser!PASSWD), "", RsUser!PASSWD))))
    txt(2) = IIf(IsNull(RsUser!PASSWD), "", DCODIFY(RTrim(IIf(IsNull(RsUser!PASSWD), "", RsUser!PASSWD))))
    txt(3) = IIf(RsUser!Label = "Y", "Yes", "No")
    txt(4) = IIf(RsUser!AcPosting = "Y", "Yes", "No")
    Set GRs = New Recordset
    GRs.CursorLocation = adUseClient
    If PubBackEnd = "A" Then
        GRs.Open "select  User2.User_Name , User_Module.Form_Code, user2.Param_Str, '" & PubCenCompCode & "' as Comp_Code, User2.Div_Code  ,User_Module.name as FormName,User_MODULE.Module_Name as ModuleName from user_module  left join user2 on user_module.form_code +user_module.Module_Name=user2.form_code + user2.Module_Name where  (user2.user_name='" & txt(0) & "' Or User2.User_Name is Null) and (user2.comp_code='" & PubCenCompCode & "' Or User2.Comp_Code Is Null) order by user_MODULE.Module_Name,user_MODULE.name", G_CompCn, adOpenDynamic, adLockOptimistic
    Else
        GRs.Open "select  User2.User_Name , User_Module.Form_Code, IsNull(user2.Param_Str,'****') as Param_Str, '" & PubCenCompCode & "' as  Comp_Code, User2.Div_Code  ,User_Module.name as FormName,User_MODULE.Module_Name as ModuleName from user_module  left join user2 on user_module.form_code +user_module.Module_Name+'" & PubCenCompCode & "'+'" & txt(0) & "'=user2.form_code + user2.Module_Name+User2.Comp_Code +User2.User_Name where  (user2.user_name='" & txt(0) & "' Or User2.User_Name is Null) and (user2.comp_code='" & PubCenCompCode & "' Or User2.Comp_Code Is Null) order by user_MODULE.Module_Name,user_MODULE.name", G_CompCn, adOpenDynamic, adLockOptimistic
    End If
    Do Until GRs.EOF
        FGridDupli.AddItem "" & Chr(9) & GRs!Comp_Code & Chr(9) & GRs!Div_Code & Chr(9) & GRs!Form_Code & Chr(9) & GRs!ModuleName & Chr(9) & GRs!param_str
        GRs.MoveNext
    Loop
    Set GRs = New Recordset
    GRs.CursorLocation = adUseClient
    GRs.Open "SELECT user1.param_str,Company.start_date,Company.Comp_Name, User1.comp_code,User1.Div_Name,User1.div_code,User1.mod_veh,User1.mod_spr,User1.mod_wsp,User1.mod_acc,User1.mod_set FROM User1 LEFT JOIN Company ON User1.Comp_Code = Company.Comp_Code where user1.User_name = 'SA' and user1.comp_code='" & PubCenCompCode & "' Order by Company.Comp_Name,User1.Div_Code", G_CompCn, adOpenStatic, adLockReadOnly
    Do Until GRs.EOF
        FGridComp.AddItem "" & Chr(9) & GRs!Comp_Code & Chr(9) & GRs!Div_Code & Chr(9) & GRs!Comp_Name & Chr(9) & GRs!Div_Name & Chr(9) & GRs!Start_Date & Chr(9) & "" & Chr(9) & GRs!mod_veh & Chr(9) & GRs!mod_spr & Chr(9) & GRs!mod_wsp & Chr(9) & GRs!mod_acc & Chr(9) & GRs!mod_set
        Set rs1 = New Recordset
        rs1.CursorLocation = adUseClient
        rs1.Open "select param_str,mod_veh,mod_spr,mod_wsp,mod_set,mod_acc from user1 where user1.user_name='" & txt(0) & "' and user1.comp_code = '" & GRs!Comp_Code & "' and div_code = '" & GRs!Div_Code & "'", G_CompCn, adOpenStatic, adLockReadOnly
        If rs1.RecordCount > 0 Then
            FGridComp.Row = FGridComp.Rows - 1
            FGridComp.Col = 6
            FGridComp.CellFontName = "wingdings"
            FGridComp.CellFontSize = 18
            FGridComp.CellForeColor = vbBlue
            FGridComp.TextMatrix(FGridComp.Rows - 1, FGridComp.Col) = "ü"
            FGridComp.TextMatrix(FGridComp.Rows - 1, 12) = IIf(rs1!mod_veh = 1, 1, 0)
            FGridComp.TextMatrix(FGridComp.Rows - 1, 13) = IIf(rs1!mod_spr = 1, 1, 0)
            FGridComp.TextMatrix(FGridComp.Rows - 1, 14) = IIf(rs1!mod_wsp = 1, 1, 0)
            FGridComp.TextMatrix(FGridComp.Rows - 1, 15) = IIf(rs1!mod_acc = 1, 1, 0)
            FGridComp.TextMatrix(FGridComp.Rows - 1, 16) = IIf(rs1!mod_set = 1, 1, 0)
        End If
        GRs.MoveNext
    Loop
    FGridComp.Col = 0
    For I = 1 To FGridComp.Rows - 1
        FGridComp.Row = I
        FGridComp.CellFontName = "wingdings"
        FGridComp.CellFontSize = 16
        FGridComp.CellForeColor = vbBlue
    Next
    FGridComp.Row = 1
    FGridComp.TextMatrix(1, 0) = "Ü"
    'FGridPer.Redraw = True
    'FGridComp.Redraw = True
End If
FillRec = 0
Call Fill_Line(1)
Call Fill_Site

If FGridPer.Rows = 1 Then FGridPer.AddItem ""
If FGridComp.Rows = 1 Then FGridComp.AddItem ""
'TopCtrl1.tFind = False
TopCtrl1.tRef = False
TopCtrl1.tPrn = False

ELoop:
End Sub

Private Function CODIFY(txt As String) As String
    Dim xxx As String
    Dim xx As Integer, MyVal As Integer
    Randomize
    MyVal = Int((99 * Rnd) + 1)
    xxx = Chr(MyVal + 27)
    For xx = 1 To Len(txt)
        xxx = xxx + Chr(Asc(mID(txt, xx, 1)) + 27 + MyVal)
    Next
    CODIFY = xxx
End Function

Private Function DCODIFY(txt As String) As String
    Dim xxx As String
    Dim xx As Integer, MyVal As Integer
    If txt = "" Then DCODIFY = "": Exit Function
    MyVal = Asc(left(txt, 1)) - 27
    xxx = ""
    For xx = 1 To Len(txt) - 1
        xxx = xxx + Chr(Asc(mID(txt, xx + 1, 1)) - 27 - MyVal)
    Next
    DCODIFY = xxx
End Function

Private Sub Disp_Text(Enb As Boolean)
    txt(0).Enabled = Enb
    txt(1).Enabled = Enb
    txt(2).Enabled = Enb
    txt(3).Enabled = Enb
    txt(4).Enabled = Enb
    
    opt(0).Enabled = Enb
    opt(1).Enabled = Enb
    opt(2).Enabled = Enb
    opt(3).Enabled = Enb
    opt(4).Enabled = Enb
    
    chk(0).Enabled = Enb
    chk(1).Enabled = Enb
    chk(2).Enabled = Enb
    chk(3).Enabled = Enb
    chk(4).Enabled = Enb
    Cmdadd.Enabled = Enb
    CmdEdit.Enabled = Enb
    CmdDel.Enabled = Enb
    Cmdrevoke.Enabled = Enb
    CmdAllow.Enabled = Enb
    Command3.Enabled = Enb
    Command1.Enabled = Enb
End Sub
Private Sub Ctrl_validate(Index As Integer)
txt(Index).BackColor = CtrlBColOrg
txt(Index).ForeColor = CtrlFColOrg
End Sub
Private Sub Ctrl_GetFocus(Index As Integer)
txt(Index).BackColor = CtrlBCol
txt(Index).ForeColor = CtrlFCol
End Sub
Private Sub BlankText()
Dim I As Byte
For I = 0 To 3
    txt(I).TEXT = ""
Next I
End Sub
Private Sub Fill_Line(rowval)
Dim GRs As Recordset
Dim param_val As String
If FGridComp.Rows > 1 Then
    FGridComp.Row = rowval
    'FGridPer.Redraw = False
    For I = 0 To 4
        chk(I).Enabled = False
        chk(I).Value = Unchecked
    Next
    chk(0).Value = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 12)) = 1, Checked, Unchecked)
    chk(1).Value = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 13)) = 1, Checked, Unchecked)
    chk(2).Value = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 14)) = 1, Checked, Unchecked)
    chk(3).Value = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 15)) = 1, Checked, Unchecked)
    chk(4).Value = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 16)) = 1, Checked, Unchecked)
    If TopCtrl1.TopText2.CAPTION <> "Browse" Then
        chk(0).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 7)) = 1, True, False)
        chk(1).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 8)) = 1, True, False)
        chk(2).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 9)) = 1, True, False)
        chk(3).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 10)) = 1, True, False)
        chk(4).Enabled = IIf(Val(FGridComp.TextMatrix(FGridComp.Row, 11)) = 1, True, False)
        opt(0).Enabled = IIf(chk(0).Value = Checked, True, False)
        opt(1).Enabled = IIf(chk(1).Value = Checked, True, False)
        opt(2).Enabled = IIf(chk(2).Value = Checked, True, False)
        opt(3).Enabled = IIf(chk(3).Value = Checked, True, False)
        opt(4).Enabled = IIf(chk(4).Value = Checked, True, False)
        opt(0).Value = Unchecked
        opt(1).Value = Unchecked
        opt(2).Value = Unchecked
        opt(3).Value = Unchecked
        opt(4).Value = Unchecked
    End If
    FGridPer.Rows = 1
    Set GRs = New Recordset
    GRs.CursorLocation = adUseClient
    If PubBackEnd = "A" Then
        GRs.Open "select Distinct User2.User_Name , User_Module.Form_Code, user2.Param_Str, User2.Comp_Code, User2.Div_Code  ,User_Module.name as FormName,User_MODULE.Module_Name as ModuleName from user_module  left join user2 on user_module.form_code +user_module.Module_Name=user2.form_code + user2.Module_Name where (user2.user_name='SA' Or User2.User_Name Is Null) and (user2.comp_code='" & PubCenCompCode & "' Or User2.Comp_code Is Null) and (user2.div_code='" & FGridComp.TextMatrix(FGridComp.Row, 2) & "' Or User2.Div_Code Is Null) And IsNull(User_Module.Module_Name,'')<>'' order by user_MODULE.Module_Name,user_MODULE.name", G_CompCn, adOpenStatic, adLockReadOnly
    Else
        GRs.Open "select Distinct User2.User_Name , User_Module.Form_Code, IsNull(user2.Param_Str,'****') as Param_Str, '" & PubCenCompCode & "' as Comp_Code, User2.Div_Code  ,User_Module.name as FormName,User_MODULE.Module_Name as ModuleName from user_module  left join user2 on user_module.form_code +user_module.Module_Name+'" & PubCenCompCode & "'+'" & txt(0) & "'=user2.form_code + user2.Module_Name+User2.Comp_Code+User2.User_Name where (user2.user_name='" & txt(0) & "' Or User2.User_Name Is Null) and (user2.comp_code='" & PubCenCompCode & "' Or User2.Comp_code Is Null) and (user2.div_code='" & FGridComp.TextMatrix(FGridComp.Row, 2) & "' Or User2.Div_Code Is Null) And IsNull(User_Module.Module_Name,'')<>'' order by user_MODULE.Module_Name,user_MODULE.name", G_CompCn, adOpenStatic, adLockReadOnly
    End If
    GRs.MoveFirst
    Do Until GRs.EOF
'        If NAME1 <> GRs!Name Then
         If XNull(GRs!ModuleName) <> "" Then
            FGridPer.AddItem "" & Chr(9) & GRs!Form_Code & Chr(9) & GRs!ModuleName & Chr(9) & GRs!FormName
         End If
'        NAME1 = XNull(GRs!Name)
        param = ""
        For I = 1 To FGridDupli.Rows - 1
            If UCase(FGridDupli.TextMatrix(I, 1)) = UCase(GRs!Comp_Code) And UCase(FGridDupli.TextMatrix(I, 2)) = UCase(GRs!Div_Code) And UCase(FGridDupli.TextMatrix(I, 4)) = UCase(GRs!ModuleName) And UCase(FGridDupli.TextMatrix(I, 3)) = UCase(GRs!Form_Code) Then
                paramval = FGridDupli.TextMatrix(I, 5)
                Exit For
            Else
                paramval = "****"
            End If
        Next
        If paramval <> "" Then
            FGridPer.Row = FGridPer.Rows - 1
            FGridPer.Col = 4
            FGridPer.CellFontName = "wingdings"
            FGridPer.CellFontSize = 18
            FGridPer.CellForeColor = vbBlue
            FGridPer.TextMatrix(FGridPer.Rows - 1, FGridPer.Col) = IIf(mID(paramval, 1, 1) = "*", "", "ü")
            FGridPer.Col = 5
            FGridPer.CellFontName = "wingdings"
            FGridPer.CellFontSize = 18
            FGridPer.CellForeColor = vbBlue
            FGridPer.TextMatrix(FGridPer.Rows - 1, FGridPer.Col) = IIf(mID(paramval, 2, 1) = "*", "", "ü")
            FGridPer.Col = 6
            FGridPer.CellFontName = "wingdings"
            FGridPer.CellFontSize = 18
            FGridPer.CellForeColor = vbBlue
            FGridPer.TextMatrix(FGridPer.Rows - 1, FGridPer.Col) = IIf(mID(paramval, 3, 1) = "*", "", "ü")
            FGridPer.Col = 7
            FGridPer.CellFontName = "wingdings"
            FGridPer.CellFontSize = 18
            FGridPer.CellForeColor = vbBlue
            FGridPer.TextMatrix(FGridPer.Rows - 1, FGridPer.Col) = IIf(mID(paramval, 4, 1) = "*", "", "ü")
        End If
        GRs.MoveNext
    Loop
End If
'FGridPer.Redraw = True
End Sub
Private Sub Set_Dupli(GridRow As Integer)
'Dim i As Integer
'FGridComp.Row = GridRow
'For i = 1 To FGridDupli.Rows - 1
'    If i <= FGridDupli.Rows - 1 Then
'        If UCase(FGridDupli.TextMatrix(i, 2)) = UCase(FGridComp.TextMatrix(FGridComp.Row, 2)) And UCase(FGridDupli.TextMatrix(i, 1)) = UCase(FGridComp.TextMatrix(FGridComp.Row, 1)) Then
'            If FGridDupli.Rows = 2 Then
'                FGridDupli.Rows = 1
'                Exit For
'            Else
'                FGridDupli.RemoveItem (i)
'                i = i - 1
'            End If
'        End If
'    End If
'Next i
'For i = 1 To FGridPer.Rows - 1
'           FGridDupli.AddItem "" & Chr(9) & FGridComp.TextMatrix(FGridComp.Row, 1) & Chr(9) & FGridComp.TextMatrix(FGridComp.Row, 2) & Chr(9) & FGridPer.TextMatrix(i, 1) & Chr(9) & FGridPer.TextMatrix(i, 2) & Chr(9) & IIf(FGridPer.TextMatrix(i, 4) = "", "*", "A") & IIf(FGridPer.TextMatrix(i, 5) = "", "*", "E") & IIf(FGridPer.TextMatrix(i, 6) = "", "*", "D") & IIf(FGridPer.TextMatrix(i, 7) = "", "*", "P")
'Next
'FGridComp.TextMatrix(FGridComp.Row, 12) = IIf(Chk(0).Value = Checked, 1, 0)
'FGridComp.TextMatrix(FGridComp.Row, 13) = IIf(Chk(1).Value = Checked, 1, 0)
'FGridComp.TextMatrix(FGridComp.Row, 14) = IIf(Chk(2).Value = Checked, 1, 0)
'FGridComp.TextMatrix(FGridComp.Row, 15) = IIf(Chk(3).Value = Checked, 1, 0)
'FGridComp.TextMatrix(FGridComp.Row, 16) = IIf(Chk(4).Value = Checked, 1, 0)
End Sub
Private Sub Cmd_Enb(flag As Boolean)
    Command3.Enabled = flag
    Command1.Enabled = flag
'    FGridComp.Enabled = Not Flag
End Sub

