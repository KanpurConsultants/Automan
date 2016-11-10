VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "topctl.ocx"
Begin VB.Form FaVrEnt 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CBBE9E&
   Caption         =   "Voucher Entry"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   Icon            =   "FaVrEnt.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.Frame Frame1 
      BackColor       =   &H00BAD3D3&
      Caption         =   "Voucher Printing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4200
      Index           =   1
      Left            =   4875
      TabIndex        =   106
      Top             =   2625
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox ChkReport 
         BackColor       =   &H00BAD3D3&
         Caption         =   "Reciept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4065
         TabIndex        =   200
         Top             =   375
         Width           =   1725
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00BAD3D3&
         Caption         =   "VNo Selection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2130
         TabIndex        =   115
         Top             =   690
         Width           =   1800
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00BAD3D3&
         Caption         =   "Print Current"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   345
         TabIndex        =   114
         Top             =   690
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.Frame Frame1 
         Height          =   540
         Index           =   4
         Left            =   1815
         TabIndex        =   111
         Top             =   3360
         Width           =   2460
         Begin VB.CommandButton btnPrint1 
            Caption         =   "Print"
            Height          =   360
            Left            =   75
            TabIndex        =   113
            Top             =   135
            Width           =   1170
         End
         Begin VB.CommandButton BTNCLOSE 
            Caption         =   "Close"
            Height          =   360
            Left            =   1245
            TabIndex        =   112
            Top             =   135
            Width           =   1170
         End
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00BAD3D3&
         Caption         =   "VDate Selection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   4065
         TabIndex        =   110
         Top             =   690
         Width           =   1800
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00BAD3D3&
         Height          =   1500
         Index           =   3
         Left            =   345
         TabIndex        =   107
         Top             =   1275
         Visible         =   0   'False
         Width           =   5400
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   2040
            TabIndex        =   108
            Top             =   720
            Width           =   1605
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "For Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   1230
            TabIndex        =   109
            Top             =   720
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00BAD3D3&
         Enabled         =   0   'False
         Height          =   1500
         Index           =   2
         Left            =   345
         TabIndex        =   116
         Top             =   1275
         Width           =   5400
         Begin MSDataListLib.DataCombo DataCombo3 
            Height          =   315
            Left            =   1800
            TabIndex        =   117
            Top             =   435
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   315
            Index           =   0
            Left            =   1425
            TabIndex        =   118
            Top             =   1005
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   315
            Index           =   1
            Left            =   3705
            TabIndex        =   119
            Top             =   1005
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From V. No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   255
            TabIndex        =   122
            Top             =   1020
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To V. No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   2745
            TabIndex        =   121
            Top             =   1035
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "V. Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   840
            TabIndex        =   120
            Top             =   480
            Width           =   750
         End
      End
   End
   Begin VB.Frame FrameTDS 
      BackColor       =   &H00BFD0B7&
      Caption         =   "T.D.S."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   11010
      TabIndex        =   168
      Top             =   855
      Visible         =   0   'False
      Width           =   6450
      Begin VB.CommandButton TDSDelete 
         DisabledPicture =   "FaVrEnt.frx":030A
         Height          =   495
         Left            =   2790
         Picture         =   "FaVrEnt.frx":044C
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Delete "
         Top             =   1995
         Width           =   585
      End
      Begin VB.TextBox TxtTDSAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1170
         TabIndex        =   98
         Top             =   2310
         Width           =   1365
      End
      Begin VB.TextBox TxtTDS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1170
         TabIndex        =   97
         Top             =   2055
         Width           =   1365
      End
      Begin VB.TextBox TxtONAMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1170
         TabIndex        =   96
         Top             =   1800
         Width           =   1365
      End
      Begin VB.TextBox TxtTDSNarration 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   1170
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   95
         Top             =   585
         Width           =   3930
      End
      Begin VB.TextBox TxtTDSCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   1170
         MaxLength       =   50
         TabIndex        =   94
         Top             =   330
         Width           =   2700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.D.S. Amt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   165
         TabIndex        =   173
         Top             =   2325
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.D.S. %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   165
         TabIndex        =   172
         Top             =   2070
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   18
         Left            =   165
         TabIndex        =   171
         Top             =   1815
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Narration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   165
         TabIndex        =   170
         Top             =   585
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TDS A/C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   165
         TabIndex        =   169
         Top             =   345
         Width           =   780
      End
   End
   Begin VB.Frame FRAMEADJUST 
      BackColor       =   &H00E6AC86&
      Caption         =   "Adjustment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3270
      Left            =   11115
      TabIndex        =   153
      Top             =   105
      Visible         =   0   'False
      Width           =   11685
      Begin VB.TextBox TXTNARRATION 
         BackColor       =   &H00EBDBC7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   795
         Left            =   90
         TabIndex        =   156
         Text            =   "Text1"
         Top             =   2430
         Width           =   10920
      End
      Begin VB.CommandButton ADJ_CANCLE 
         DisabledPicture =   "FaVrEnt.frx":088E
         Height          =   495
         Left            =   11055
         Picture         =   "FaVrEnt.frx":09D0
         Style           =   1  'Graphical
         TabIndex        =   159
         ToolTipText     =   "Cancel Changes"
         Top             =   1935
         Width           =   585
      End
      Begin VB.CommandButton ADJ_OK 
         DisabledPicture =   "FaVrEnt.frx":0B12
         Height          =   495
         Left            =   11055
         Picture         =   "FaVrEnt.frx":0C14
         Style           =   1  'Graphical
         TabIndex        =   158
         ToolTipText     =   "Ok"
         Top             =   1440
         Width           =   585
      End
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   11055
         Picture         =   "FaVrEnt.frx":0D56
         Style           =   1  'Graphical
         TabIndex        =   157
         ToolTipText     =   "Full Adjuatments"
         Top             =   945
         Width           =   585
      End
      Begin VB.TextBox TXTADJ_AMT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F7F0DF&
         Height          =   285
         Left            =   4905
         TabIndex        =   155
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton BTS_AUTO_ADJ 
         Caption         =   "Auto"
         DisabledPicture =   "FaVrEnt.frx":1198
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11055
         Style           =   1  'Graphical
         TabIndex        =   154
         ToolTipText     =   "Ok"
         Top             =   450
         Width           =   585
      End
      Begin MSFlexGridLib.MSFlexGrid FgridAdjust 
         Height          =   1965
         Left            =   45
         TabIndex        =   175
         Top             =   450
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   3466
         _Version        =   393216
         Cols            =   14
         BackColor       =   12648447
         ForeColor       =   64
         BackColorFixed  =   12440489
         ForeColorFixed  =   128
         BackColorSel    =   8388608
         ForeColorSel    =   65535
         BackColorBkg    =   15117446
         GridColor       =   255
         FocusRect       =   0
         HighLight       =   2
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"FaVrEnt.frx":129A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "MMMM"
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
         Height          =   240
         Left            =   1035
         TabIndex        =   166
         Top             =   225
         Width           =   4425
      End
      Begin VB.Label Label1 
         BackColor       =   &H005EB0AC&
         BackStyle       =   0  'Transparent
         Caption         =   "For A/C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   13
         Left            =   135
         TabIndex        =   167
         Top             =   225
         Width           =   840
      End
      Begin VB.Label ADJ_LAB8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   10545
         TabIndex        =   165
         Top             =   225
         Width           =   480
      End
      Begin VB.Label ADJ_LAB5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   7590
         TabIndex        =   164
         Top             =   225
         Width           =   510
      End
      Begin VB.Label ADJ_LAB4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "MMMM"
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
         Height          =   240
         Left            =   6285
         TabIndex        =   163
         Top             =   225
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackColor       =   &H005EB0AC&
         BackStyle       =   0  'Transparent
         Caption         =   "Tr.Amt."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   14
         Left            =   5610
         TabIndex        =   162
         Top             =   225
         Width           =   675
      End
      Begin VB.Label Label1 
         BackColor       =   &H005EB0AC&
         BackStyle       =   0  'Transparent
         Caption         =   "Adjusted Amt."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   15
         Left            =   8205
         TabIndex        =   161
         Top             =   225
         Width           =   1290
      End
      Begin VB.Label ADJ_LAB7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "KK"
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
         Height          =   240
         Left            =   9495
         TabIndex        =   160
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.Frame FRAMEVLIST 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Voucher Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   4200
      Left            =   2475
      TabIndex        =   88
      Top             =   1590
      Visible         =   0   'False
      Width           =   8520
      Begin MSFlexGridLib.MSFlexGrid FGVLIST 
         Height          =   3090
         Left            =   60
         TabIndex        =   89
         Top             =   1050
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   5450
         _Version        =   393216
         Cols            =   9
         BackColor       =   13487355
         BackColorFixed  =   8421504
         ForeColorFixed  =   65535
         BackColorSel    =   13487355
         ForeColorSel    =   12582912
         BackColorBkg    =   13487355
         GridColor       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   $"FaVrEnt.frx":134D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4395
         MaxLength       =   12
         TabIndex        =   103
         Top             =   615
         Width           =   1185
      End
      Begin VB.TextBox TXTVDATE2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2565
         MaxLength       =   12
         TabIndex        =   100
         Top             =   315
         Width           =   1170
      End
      Begin VB.TextBox TXTVDATE1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   660
         MaxLength       =   12
         TabIndex        =   99
         Top             =   315
         Width           =   1170
      End
      Begin VB.CommandButton BTNVLCLOSE 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7365
         TabIndex        =   105
         Top             =   615
         Width           =   1050
      End
      Begin VB.CommandButton BTNVLOK 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7365
         TabIndex        =   104
         Top             =   255
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo Dcparty 
         Height          =   315
         Left            =   660
         TabIndex        =   101
         Top             =   600
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   4395
         TabIndex        =   102
         Top             =   300
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vr.Type"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   3795
         TabIndex        =   151
         Top             =   360
         Width           =   555
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   2250
         TabIndex        =   93
         Top             =   360
         Width           =   270
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fr.Dt."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   165
         TabIndex        =   92
         Top             =   360
         Width           =   390
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vr.No."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   3795
         TabIndex        =   91
         Top             =   660
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   165
         TabIndex        =   90
         Top             =   660
         Width           =   465
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame FrameRef 
      BackColor       =   &H00D9B0DD&
      Caption         =   "Referance Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4395
      Left            =   7485
      TabIndex        =   174
      Top             =   7815
      Visible         =   0   'False
      Width           =   8805
      Begin VB.CommandButton BtnRefAdjOK 
         DisabledPicture =   "FaVrEnt.frx":13F0
         Height          =   495
         Left            =   7860
         Picture         =   "FaVrEnt.frx":14F2
         Style           =   1  'Graphical
         TabIndex        =   198
         ToolTipText     =   "Save"
         Top             =   270
         Width           =   585
      End
      Begin VB.Frame FrmList 
         BorderStyle     =   0  'None
         Height          =   2505
         Left            =   6345
         TabIndex        =   186
         Top             =   1065
         Visible         =   0   'False
         Width           =   2010
         Begin MSComctlLib.ListView ListView 
            Height          =   2490
            Left            =   0
            TabIndex        =   180
            TabStop         =   0   'False
            Top             =   0
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   4392
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
            BackColor       =   12648447
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
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   240
         Index           =   0
         Left            =   5895
         MaxLength       =   150
         TabIndex        =   185
         Top             =   1680
         Visible         =   0   'False
         Width           =   690
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGridRef 
         Height          =   3555
         Left            =   45
         TabIndex        =   181
         Top             =   780
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   6271
         _Version        =   393216
         BackColor       =   12648447
         Cols            =   16
         BackColorFixed  =   15718825
         ForeColorFixed  =   128
         BackColorSel    =   16777215
         ForeColorSel    =   12582912
         BackColorBkg    =   12768697
         GridColor       =   255
         GridColorFixed  =   32896
         WordWrap        =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   3
         BorderStyle     =   0
         Appearance      =   0
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   16
      End
      Begin MSDataGridLib.DataGrid DGRefNo 
         Height          =   2670
         Left            =   4080
         TabIndex        =   179
         Top             =   1560
         Visible         =   0   'False
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   4710
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   12648447
         BorderStyle     =   0
         Enabled         =   -1  'True
         ColumnHeaders   =   -1  'True
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   19
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "RefNo"
            Caption         =   "Ref.No."
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
            DataField       =   "VDate"
            Caption         =   "Vr.Date"
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
            DataField       =   "Balance"
            Caption         =   "Balance"
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
            SizeMode        =   1
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               DividerStyle    =   1
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   1
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H005EB0AC&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
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
         Height          =   240
         Index           =   24
         Left            =   4995
         TabIndex        =   197
         Top             =   540
         Width           =   690
      End
      Begin VB.Label LblRefAdjBal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "999999999.99"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   5730
         TabIndex        =   196
         Top             =   540
         Width           =   1230
      End
      Begin VB.Label LblRefAdjDrCrBal 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "KK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   6975
         TabIndex        =   195
         Top             =   540
         Width           =   270
      End
      Begin VB.Label LblRefAdjDrCr 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "KK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   4635
         TabIndex        =   194
         Top             =   540
         Width           =   270
      End
      Begin VB.Label LblRefAmtDrCr 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cr."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2235
         TabIndex        =   191
         Top             =   540
         Width           =   270
      End
      Begin VB.Label LblRefAdj 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "999999999.99"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3405
         TabIndex        =   193
         Top             =   540
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H005EB0AC&
         BackStyle       =   0  'Transparent
         Caption         =   "Adjusted"
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
         Height          =   240
         Index           =   23
         Left            =   2550
         TabIndex        =   192
         Top             =   540
         Width           =   765
      End
      Begin VB.Label Label1 
         BackColor       =   &H005EB0AC&
         BackStyle       =   0  'Transparent
         Caption         =   "Tr.Amt."
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
         Height          =   240
         Index           =   22
         Left            =   90
         TabIndex        =   189
         Top             =   540
         Width           =   675
      End
      Begin VB.Label LblRefAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "999999999.99"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   990
         TabIndex        =   190
         Top             =   540
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackColor       =   &H005EB0AC&
         BackStyle       =   0  'Transparent
         Caption         =   "For A/C"
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
         Height          =   240
         Index           =   21
         Left            =   90
         TabIndex        =   187
         Top             =   300
         Width           =   840
      End
      Begin VB.Label LblRefName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "MMMM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   990
         TabIndex        =   188
         Top             =   300
         Width           =   3330
      End
   End
   Begin VB.PictureBox PicDN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7005
      Picture         =   "FaVrEnt.frx":1634
      ScaleHeight     =   255
      ScaleWidth      =   300
      TabIndex        =   184
      Top             =   4980
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicUP 
      Appearance      =   0  'Flat
      BackColor       =   &H00937B73&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      Picture         =   "FaVrEnt.frx":1A76
      ScaleHeight     =   225
      ScaleWidth      =   300
      TabIndex        =   183
      Top             =   780
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox TxtGlb 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   0
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   64
      Top             =   5220
      Width           =   7020
   End
   Begin MSDataGridLib.DataGrid DGTDSCODE 
      Height          =   3330
      Left            =   11040
      Negotiate       =   -1  'True
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   1395
      Visible         =   0   'False
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12648447
      BorderStyle     =   0
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
         Weight          =   700
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "SubCode"
         Caption         =   "SubCode"
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
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGAcHlp 
      Height          =   1725
      Left            =   7080
      TabIndex        =   124
      Top             =   1320
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3043
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   13825259
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   13504523
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      WrapCellPointer =   -1  'True
      RowDividerStyle =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "SUBCODE"
         Caption         =   "AcCode"
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
         Caption         =   "A/C Name"
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
         DataField       =   "NameHelp"
         Caption         =   "NameHelp"
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
         DataField       =   "Nature"
         Caption         =   "Nature"
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
         DataField       =   "GROUPNAME"
         Caption         =   "GroupName"
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
         SizeMode        =   1
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            DividerStyle    =   0
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   3555.213
         EndProperty
         BeginProperty Column02 
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column03 
            DividerStyle    =   1
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column04 
            DividerStyle    =   1
            ColumnWidth     =   3795.024
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtDetailS 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   6075
      Locked          =   -1  'True
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   152
      Top             =   5955
      Visible         =   0   'False
      Width           =   4980
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   11
      Left            =   330
      MaxLength       =   35
      TabIndex        =   60
      Top             =   7890
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   11
      Left            =   990
      MaxLength       =   255
      TabIndex        =   63
      Top             =   8130
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   11
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   59
      Text            =   "Cr"
      Top             =   7890
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   11
      Left            =   9015
      TabIndex        =   62
      Top             =   7890
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   11
      Left            =   7620
      TabIndex        =   61
      Top             =   7890
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   10
      Left            =   7620
      TabIndex        =   56
      Top             =   7425
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   10
      Left            =   9015
      TabIndex        =   57
      Top             =   7425
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   10
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   54
      Text            =   "Cr"
      Top             =   7425
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   10
      Left            =   990
      MaxLength       =   255
      TabIndex        =   58
      Top             =   7665
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   10
      Left            =   330
      MaxLength       =   35
      TabIndex        =   55
      Top             =   7425
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   9
      Left            =   7620
      TabIndex        =   51
      Top             =   6945
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   9
      Left            =   9015
      TabIndex        =   52
      Top             =   6945
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   9
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   49
      Text            =   "Cr"
      Top             =   6945
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   9
      Left            =   990
      MaxLength       =   255
      TabIndex        =   53
      Top             =   7185
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   9
      Left            =   330
      MaxLength       =   35
      TabIndex        =   50
      Top             =   6945
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   8
      Left            =   330
      MaxLength       =   35
      TabIndex        =   45
      Top             =   6480
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   8
      Left            =   990
      MaxLength       =   255
      TabIndex        =   48
      Top             =   6720
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   8
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   44
      Text            =   "Cr"
      Top             =   6480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   8
      Left            =   9015
      TabIndex        =   47
      Top             =   6480
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   8
      Left            =   7620
      TabIndex        =   46
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   7
      Left            =   330
      MaxLength       =   35
      TabIndex        =   40
      Top             =   5985
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   7
      Left            =   990
      MaxLength       =   255
      TabIndex        =   43
      Top             =   6225
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   7
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   39
      Text            =   "Cr"
      Top             =   5985
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   7
      Left            =   9015
      TabIndex        =   42
      Top             =   5985
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   7
      Left            =   7620
      TabIndex        =   41
      Top             =   5985
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   6
      Left            =   330
      MaxLength       =   35
      TabIndex        =   35
      Top             =   5535
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   6
      Left            =   990
      MaxLength       =   255
      TabIndex        =   38
      Top             =   5760
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   6
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   34
      Text            =   "Cr"
      Top             =   5535
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   6
      Left            =   9015
      TabIndex        =   37
      Top             =   5535
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   6
      Left            =   7620
      TabIndex        =   36
      Top             =   5535
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   5
      Left            =   7620
      TabIndex        =   31
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   5
      Left            =   9015
      TabIndex        =   32
      Top             =   4800
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   5
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   29
      Text            =   "Cr"
      Top             =   4800
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   5
      Left            =   990
      MaxLength       =   255
      TabIndex        =   33
      Top             =   5280
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   5
      Left            =   330
      MaxLength       =   35
      TabIndex        =   30
      Top             =   4725
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   4
      Left            =   330
      MaxLength       =   35
      TabIndex        =   25
      Top             =   4060
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   4
      Left            =   990
      MaxLength       =   255
      TabIndex        =   28
      Top             =   4540
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   0
      Left            =   990
      MaxLength       =   255
      TabIndex        =   8
      Top             =   1500
      Width           =   6540
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   1
      Left            =   990
      MaxLength       =   255
      TabIndex        =   13
      Top             =   2260
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   2
      Left            =   990
      MaxLength       =   255
      TabIndex        =   18
      Top             =   3020
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.TextBox TxtNar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   3
      Left            =   990
      MaxLength       =   255
      TabIndex        =   23
      Top             =   3780
      Visible         =   0   'False
      Width           =   6540
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   0
      Left            =   330
      MaxLength       =   35
      TabIndex        =   5
      Top             =   1020
      Width           =   7200
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   3
      Left            =   330
      MaxLength       =   35
      TabIndex        =   20
      Top             =   3300
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   2
      Left            =   330
      MaxLength       =   35
      TabIndex        =   15
      Top             =   2540
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.TextBox TxtAcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   1
      Left            =   330
      MaxLength       =   35
      TabIndex        =   10
      Top             =   1780
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   4
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   24
      Text            =   "Cr"
      Top             =   4060
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   3
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   19
      Text            =   "Cr"
      Top             =   3300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   2
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   14
      Text            =   "Cr"
      Top             =   2540
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   1
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "Cr"
      Top             =   1780
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox TxtCrDr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   0
      Left            =   15
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "Cr"
      Top             =   1020
      Width           =   285
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Height          =   375
      Left            =   0
      TabIndex        =   127
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   661
   End
   Begin MSDataGridLib.DataGrid DGVchrHlp 
      Height          =   4320
      Left            =   9105
      TabIndex        =   126
      Top             =   7680
      Visible         =   0   'False
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   7620
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   13825259
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      ForeColor       =   13504523
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      WrapCellPointer =   -1  'True
      RowDividerStyle =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "V_tYPE"
         Caption         =   "V_tYPE"
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
         DataField       =   "Description"
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
         DataField       =   "Prefix"
         Caption         =   "Prefix"
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
         SizeMode        =   1
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            DividerStyle    =   0
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   1
            Locked          =   -1  'True
            ColumnWidth     =   2819.906
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TXTClrDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   5580
      MaxLength       =   12
      TabIndex        =   67
      Top             =   6135
      Width           =   1125
   End
   Begin VB.TextBox TXTChDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   3765
      MaxLength       =   12
      TabIndex        =   66
      Top             =   6135
      Width           =   1125
   End
   Begin VB.TextBox TxtCHno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   765
      MaxLength       =   20
      TabIndex        =   65
      Top             =   6135
      Width           =   2520
   End
   Begin MSFlexGridLib.MSFlexGrid FGrid1 
      Height          =   570
      Left            =   165
      TabIndex        =   83
      Top             =   8100
      Visible         =   0   'False
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   1005
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      FixedRows       =   0
   End
   Begin VB.TextBox TxtVtYpe 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   3345
      TabIndex        =   1
      Top             =   450
      Width           =   2085
   End
   Begin VB.TextBox VchDt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   480
      MaxLength       =   12
      TabIndex        =   0
      Top             =   450
      Width           =   1095
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   4
      Left            =   9015
      TabIndex        =   27
      Top             =   4060
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   4
      Left            =   7620
      TabIndex        =   26
      Top             =   4060
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00937B73&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5520
      Index           =   0
      Left            =   10425
      TabIndex        =   73
      Top             =   0
      Width           =   1230
      Begin VB.Label LblShort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "F6 Contra"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   77
         Tag             =   "vbKeyF6"
         Top             =   525
         Width           =   750
      End
      Begin VB.Label LblShort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "F7 Payment"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   76
         Tag             =   "vbKeyF7"
         Top             =   885
         Width           =   870
      End
      Begin VB.Label LblShort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "F8 Receipt"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   75
         Tag             =   "vbKeyF8"
         Top             =   1245
         Width           =   825
      End
      Begin VB.Label LblShort 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "F9 Journal"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   74
         Tag             =   "vbKeyF9"
         Top             =   1620
         Width           =   825
      End
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   3
      Left            =   7620
      TabIndex        =   21
      Top             =   3300
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   3
      Left            =   9015
      TabIndex        =   22
      Top             =   3300
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   2
      Left            =   7620
      TabIndex        =   16
      Top             =   2540
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   2
      Left            =   9015
      TabIndex        =   17
      Top             =   2540
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   1
      Left            =   7620
      TabIndex        =   11
      Top             =   1780
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   1
      Left            =   9015
      TabIndex        =   12
      Top             =   1780
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox TxtCr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   9015
      TabIndex        =   7
      Top             =   1020
      Width           =   1365
   End
   Begin VB.TextBox TxtDr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   7620
      TabIndex        =   6
      Top             =   1020
      Width           =   1335
   End
   Begin VB.TextBox TxtVno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   6690
      MaxLength       =   8
      TabIndex        =   2
      Top             =   450
      Width           =   1065
   End
   Begin MSDataListLib.DataCombo TxtSite 
      DataField       =   "PARTY"
      DataSource      =   "master"
      Height          =   285
      Left            =   8520
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   503
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16249055
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblAmtRs 
      BackColor       =   &H005EB0AC&
      BackStyle       =   0  'Transparent
      Caption         =   "LblAmtRs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   7125
      TabIndex        =   199
      Top             =   5505
      Width           =   3285
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Narration                                                               "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   68
      Top             =   4980
      Width           =   7005
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   5
      Left            =   0
      TabIndex        =   182
      Top             =   5040
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "For Site"
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
      Left            =   7830
      TabIndex        =   178
      Top             =   450
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label LblHelp 
      AutoSize        =   -1  'True
      BackColor       =   &H00223966&
      Caption         =   "Press <Ins> Bill Wise Adjustment ,<Alt-R> Against Reference , <Alt-T> T.D.S.Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   15
      TabIndex        =   177
      Top             =   2760
      Visible         =   0   'False
      Width           =   10365
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   11
      Left            =   0
      TabIndex        =   150
      Top             =   8130
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   11
      Left            =   10425
      TabIndex        =   149
      Top             =   7920
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   10
      Left            =   10425
      TabIndex        =   148
      Top             =   7455
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   10
      Left            =   0
      TabIndex        =   147
      Top             =   7665
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   9
      Left            =   10425
      TabIndex        =   146
      Top             =   6960
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   9
      Left            =   0
      TabIndex        =   145
      Top             =   7185
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   8
      Left            =   0
      TabIndex        =   144
      Top             =   6720
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   8
      Left            =   10455
      TabIndex        =   143
      Top             =   6480
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   7
      Left            =   0
      TabIndex        =   142
      Top             =   6240
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   7
      Left            =   10470
      TabIndex        =   141
      Top             =   5993
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   6
      Left            =   0
      TabIndex        =   140
      Top             =   5760
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   6
      Left            =   10440
      TabIndex        =   139
      Top             =   5550
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   5
      Left            =   0
      TabIndex        =   138
      Top             =   5280
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   4
      Left            =   0
      TabIndex        =   137
      Top             =   4545
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   3
      Left            =   0
      TabIndex        =   136
      Top             =   3780
      Width           =   945
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   135
      Top             =   3015
      Width           =   945
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   134
      Top             =   2265
      Width           =   945
   End
   Begin VB.Label LblNar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   133
      Top             =   1500
      Width           =   945
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   4
      Left            =   0
      TabIndex        =   132
      Top             =   4320
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   3
      Left            =   0
      TabIndex        =   131
      Top             =   3540
      Width           =   7500
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   130
      Top             =   2780
      Width           =   7500
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   129
      Top             =   2020
      Width           =   7500
   End
   Begin VB.Label LblCb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   128
      Top             =   1260
      Width           =   7500
   End
   Begin VB.Label LblDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1680
      TabIndex        =   123
      Top             =   465
      Width           =   75
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clr.Dt."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5010
      TabIndex        =   86
      Top             =   6150
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3285
      TabIndex        =   85
      Top             =   6150
      Width           =   420
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ch. No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   84
      Top             =   6150
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dIFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6720
      TabIndex        =   82
      Top             =   5175
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LblVPrefix"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   5445
      TabIndex        =   81
      Top             =   450
      Width           =   705
   End
   Begin VB.Label LblVtype 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr.type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2700
      TabIndex        =   80
      Top             =   450
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00937B73&
      Caption         =   "       Particulars                                                                        "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BEFDFE&
      Height          =   225
      Index           =   0
      Left            =   15
      TabIndex        =   87
      Top             =   780
      Width           =   7545
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00937B73&
      Caption         =   "Credit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BEFDFE&
      Height          =   225
      Index           =   2
      Left            =   8985
      TabIndex        =   79
      Top             =   780
      Width           =   1395
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00937B73&
      Caption         =   " Debit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BEFDFE&
      Height          =   225
      Index           =   1
      Left            =   7575
      TabIndex        =   78
      Top             =   780
      Width           =   1395
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblCrAmt 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   10170
      TabIndex        =   72
      Top             =   5145
      Width           =   210
   End
   Begin VB.Label LblDrAmt 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   8730
      TabIndex        =   71
      Top             =   5145
      Width           =   225
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   7125
      X2              =   10395
      Y1              =   5490
      Y2              =   5490
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   7125
      X2              =   10395
      Y1              =   5445
      Y2              =   5445
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   7140
      X2              =   10410
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   8970
      X2              =   8970
      Y1              =   795
      Y2              =   5100
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   7560
      X2              =   7560
      Y1              =   795
      Y2              =   5100
   End
   Begin VB.Label LblDt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   15
      TabIndex        =   70
      Top             =   450
      Width           =   420
   End
   Begin VB.Label LblVno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr.No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6120
      TabIndex        =   69
      Top             =   450
      Width           =   555
   End
End
Attribute VB_Name = "FaVrEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BackColorSelLeave As String
Private Const BackColorSelEnter As String = &HF8D7FD, CtrlBCol = &H80000008, CtrlFCol = &H8000000E
Private CtrlBColOrg As Long, CtrlFColOrg As Long, CurrObj As TextBox
Private RstMain As ADODB.Recordset, RstMainAdj As ADODB.Recordset, RstVchrHlp As ADODB.Recordset, RstEnviro As ADODB.Recordset, RstRefHelp As ADODB.Recordset
Private RstAcHlpDr As ADODB.Recordset, RstAcHlpCr As ADODB.Recordset, RstTds As ADODB.Recordset, RstRef As ADODB.Recordset, RstTDSHlp As ADODB.Recordset
Private ScrolIndex As Integer, ADDFLAG As Byte, mSepNar As String, mCommNar As String, mNCat As String, mLastVrType As String, mLastPrefix As String
Dim mMoveFlag As Boolean, mCurrObjectShow As Boolean, FixRow As Integer, mxLastVrType As String
Dim ListArray As Variant, mListItem As ListItem, GridKey As Integer, mDefaultCrAc As String, mDefaultDrAc As String
Private PubDatamanFa As New DMFa.ClsFa, CURR_ROW As Integer, OLD_AMT1 As Double, TAddMode As Boolean
Private Const FAgRefType As Byte = 1, FAgRefNo As Byte = 2, FCr As Byte = 3, FDr As Byte = 4
Private Const FDocId As Byte = 5, FV_Sno As Byte = 6, FSubCode As Byte = 7, FDueDate As Byte = 8

Private Sub VRLIST()
Dim Rst As ADODB.Recordset, mQRY As String
On Error GoTo erroloop
    FGVLIST.Rows = 1
    mQRY = " WHERE NCAT<>'OPBAL'"
    If TXTVDATE1 <> "" And TXTVDATE2 <> "" Then mQRY = " WHERE v_date BETWEEN " & FaConvertDate(TXTVDATE1) & " and " & FaConvertDate(TXTVDATE2)
    If Dcparty <> "" Then mQRY = mQRY + IIf(Len(mQRY) = 0, " WHERE ", " AND ") + " LEDGER.SUBCODE='" & Dcparty.BoundText & "'"
    If DataCombo1 <> "" Then mQRY = mQRY + IIf(Len(mQRY) = 0, " WHERE ", " AND ") + " LEDGER.V_TYPE=" & FaChk_Text(DataCombo1.BoundText)
    If Text1 <> "" Then mQRY = mQRY + IIf(Len(mQRY) = 0, " WHERE ", " AND ") + " V_No=" & Val(Text1)
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    If PubSiteCodeWiseMasterRst = False Then
        Rst.Open "SELECT LEDGER.*,SUBGROUP.NAME FROM (LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE) LEFT JOIN VOUCHER_tYPE ON VOUCHER_tYPE.V_TYPE=LEDGER.V_tYPE " & mQRY & " ORDER BY V_DATE,LEDGER.V_TYPE,V_NO,V_SNo", G_FaCn, adOpenDynamic, adLockOptimistic
    Else
        If (PubFaSiteType = 1 And PubSeparateVrNoForSite = 1) Then
            Rst.Open "SELECT LEDGER.*,SUBGROUP.NAME FROM (LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE) LEFT JOIN VOUCHER_tYPE ON VOUCHER_tYPE.V_TYPE=LEDGER.V_tYPE " & mQRY & " AND RIGHT(LEDGER.SITE_CODE,1)='" & Trim(PubSeparateLogSite) & "' ORDER BY V_DATE,LEDGER.V_TYPE,V_NO,V_SNo", G_FaCn, adOpenDynamic, adLockOptimistic
        ElseIf PubFaSiteType = 1 Then
            Rst.Open "SELECT LEDGER.*,SUBGROUP.NAME FROM (LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE) LEFT JOIN VOUCHER_tYPE ON VOUCHER_tYPE.V_TYPE=LEDGER.V_tYPE " & mQRY & " AND LEFT(LEDGER.SITE_CODE,1)='" & Trim(PubSiteCode) & "' ORDER BY V_DATE,LEDGER.V_TYPE,V_NO,V_SNo", G_FaCn, adOpenDynamic, adLockOptimistic
        ElseIf PubFaSiteType = 2 Then
            Rst.Open "SELECT LEDGER.*,SUBGROUP.NAME FROM (LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE) LEFT JOIN VOUCHER_tYPE ON VOUCHER_tYPE.V_TYPE=LEDGER.V_tYPE " & mQRY & "  AND LEDGER.SITE_CODE='" & Trim(PubSiteCode) & "'  ORDER BY V_DATE,LEDGER.V_TYPE,V_NO,V_SNo", G_FaCn, adOpenDynamic, adLockOptimistic
        Else
            Rst.Open "SELECT LEDGER.*,SUBGROUP.NAME FROM (LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE) LEFT JOIN VOUCHER_tYPE ON VOUCHER_tYPE.V_TYPE=LEDGER.V_tYPE " & mQRY & " ORDER BY V_DATE,LEDGER.V_TYPE,V_NO,V_SNo", G_FaCn, adOpenDynamic, adLockOptimistic
        End If
    End If
    Do Until Rst.EOF
        FGVLIST.AddItem ("" & Chr(9) & Rst!V_Date & Chr(9) & Rst!V_tYPE & Chr(9) & Rst!V_NO & Chr(9) & FaXNull(Rst!V_Prefix) & Chr(9) & Rst!Name & Chr(9) & Rst!AmtDr & Chr(9) & Rst!AmtCr & Chr(9) & Rst!DocID)
        Rst.MoveNext
    Loop
Set Rst = Nothing
Exit Sub
erroloop:   MsgBox err.Description, vbInformation, Me.CAPTION: Exit Sub
End Sub
Private Sub BTNVLOK_Click()
If Trim(TXTVDATE1) = "" And Trim(TXTVDATE2) = "" And Trim(Text1) = "" And DataCombo1.BoundText = "" And Dcparty.BoundText = "" Then
Else
    VRLIST
End If
End Sub
Private Sub DGAcHlp_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyPageDown, vbKeyPageUp
        If TxtCrDr(Val(DGAcHlp.Tag)) = "Cr" Then
            If ADDFLAG <= 2 Then
                FaDGridTxtKeyDown DGAcHlp, TxtAcName, Val(DGAcHlp.Tag), RstAcHlpCr, KeyCode, False, 1
                If RstEnviro!AddressHelp = "Yes" And DGAcHlp.Visible = True Then
                    TxtDetailS.Visible = True
                    TxtDetailS.left = DGAcHlp.left
                    TxtDetailS.top = DGAcHlp.top + DGAcHlp.height
                    TxtDetailS.width = DGAcHlp.width
                Else
                    TxtDetailS.Visible = False
                End If
                If RstAcHlpCr.EOF = False And RstAcHlpCr.BOF = False Then
                    TxtDetailS = IIf(Trim(RstAcHlpCr!FatherName) = "", "", RstAcHlpCr!FatherName + vbCrLf) + RstAcHlpCr!GroupName + vbCrLf + RstAcHlpCr!NameWithADDR
                Else
                    TxtDetailS = ""
                End If
            End If
        Else
            If ADDFLAG <= 2 Then
                FaDGridTxtKeyDown DGAcHlp, TxtAcName, Val(DGAcHlp.Tag), RstAcHlpDr, KeyCode, False, 1
                If RstEnviro!AddressHelp = "Yes" And DGAcHlp.Visible = True Then
                    TxtDetailS.Visible = True
                    TxtDetailS.left = DGAcHlp.left
                    TxtDetailS.top = DGAcHlp.top + DGAcHlp.height
                    TxtDetailS.width = DGAcHlp.width
                Else
                    TxtDetailS.Visible = False
                End If
                If RstAcHlpDr.EOF = False And RstAcHlpDr.BOF = False Then
                    TxtDetailS = IIf(Trim(RstAcHlpDr!FatherName) = "", "", RstAcHlpDr!FatherName + vbCrLf) + RstAcHlpDr!GroupName + vbCrLf + RstAcHlpDr!NameWithADDR
                Else
                    TxtDetailS = ""
                End If
            End If
        End If
End Select
End Sub
Private Sub FGridRef_RowColChange()
'    If TxtGridLeave = True Then
'    '
'    End If
End Sub
Private Sub FGVLIST_Click()
If FGVLIST.TextMatrix(FGVLIST.Row, 8) <> "" Then
    RstMain.MoveFirst
    RstMain.Find ("DOCID='" & FGVLIST.TextMatrix(FGVLIST.Row, 8) & "'")
    If RstMain.EOF = False Then BUTTONS True, Me, RstMain, 0: MoveRec: FRAMEVLIST.Visible = False
End If
End Sub
Public Sub FindMove(mDocId As String)
    RstMain.MoveFirst
    RstMain.Find "DOCID='" & mDocId & "'"
    If RstMain.EOF = False Then MoveRec
End Sub
Private Sub Form_Deactivate()
    If TypeOf Me.ActiveControl Is TextBox Then
        Set CurrObj = Me.ActiveControl
        mCurrObjectShow = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set RstMain = Nothing
    Set RstVchrHlp = Nothing
    Set RstEnviro = Nothing
    Set PubDatamanFa = Nothing
    Set RstAcHlpDr = Nothing
    Set RstAcHlpCr = Nothing
    Set RstMainAdj = Nothing
    Set RstRef = Nothing
    Set RstTds = Nothing
    Set RstTDSHlp = Nothing
    Set RstRefHelp = Nothing
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    Text2 = PubDatamanFa.FaRetDateFunc(Text2)
End Sub
Private Sub TopCtrl1_eRef()
Dim mSiteHlp As String
mSiteHlp = ""
If PubSiteCodeWiseHelp = True Then
    mSiteHlp = "Where Site_Code='" & PubSiteCode & "'"
End If
    If PubBackEnd = "A" Then
        If RstEnviro!FilterAC = "No" Then
            Set RstAcHlpDr = New ADODB.Recordset
            If RstEnviro!ShowCityName = "Yes" Then
                RstAcHlpDr.Open ("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3))+RTRIM(IIF(ISNULL(CityName),'',','+CityName)) AS NameWithADDR,IIF(ISNULL(FNAME),'',FNAME) AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
            Else
                RstAcHlpDr.Open ("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3))+RTRIM(IIF(ISNULL(CityName),'',','+CityName)) AS NameWithADDR,IIF(ISNULL(FNAME),'',FNAME) AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
            End If
            Set RstAcHlpCr = New ADODB.Recordset
            If RstEnviro!ShowCityName = "Yes" Then
                RstAcHlpCr.Open ("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3))+RTRIM(IIF(ISNULL(CityName),'',','+CityName)) AS NameWithADDR,IIF(ISNULL(FNAME),'',FNAME) AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
            Else
                RstAcHlpCr.Open ("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3))+RTRIM(IIF(ISNULL(CityName),'',','+CityName)) AS NameWithADDR,IIF(ISNULL(FNAME),'',FNAME) AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
            End If
        End If
    Else
        If RstEnviro!FilterAC = "No" Then
            Set RstAcHlpDr = New ADODB.Recordset
            If RstEnviro!ShowCityName = "Yes" Then
                RstAcHlpDr.Open ("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,''))+RTRIM(ISNULL(CityName,'')) AS NameWithADDR,ISNULL(FNAME,'') AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
            Else
                RstAcHlpDr.Open ("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,''))+RTRIM(ISNULL(CityName,'')) AS NameWithADDR,ISNULL(FNAME,'') AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
            End If
            Set RstAcHlpCr = New ADODB.Recordset
            If RstEnviro!ShowCityName = "Yes" Then
                RstAcHlpCr.Open ("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,''))+RTRIM(ISNULL(CityName,'')) AS NameWithADDR,ISNULL(FNAME,'') AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
            Else
                RstAcHlpCr.Open ("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,''))+RTRIM(ISNULL(CityName,'')) AS NameWithADDR,ISNULL(FNAME,'') AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
            End If
        End If
    End If
    If Me.ActiveControl.Name = "TxtAcName" Then TxtAcName_GotFocus (Me.ActiveControl.Index)
    If RstAcHlpCr.RecordCount > 0 Then RstAcHlpCr.MoveFirst
    If RstAcHlpDr.RecordCount > 0 Then RstAcHlpDr.MoveFirst
    Grid_Hide
    Set RstEnviro = G_FaCn.Execute("SELECT * FROM FAENVIRO")
End Sub
Private Sub Form_Activate()
    If mCurrObjectShow = True Then
        If CurrObj.Visible = True And CurrObj.Enabled = True > 0 Then
            CurrObj.SetFocus
            mCurrObjectShow = False
        End If
        Set RstEnviro = G_FaCn.Execute("SELECT * FROM FAENVIRO")
        If TxtVtYpe(0).Tag <> "" Then
            FaVrTypeSetting TxtVtYpe(0).Tag
        Else
            FaVrTypeSetting
        End If
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mCancel As Boolean, mCurrentBalance As Double, RstFormKeyDown As ADODB.Recordset
If FrameRef.Visible = True Then
    Select Case KeyCode
        Case vbKeyEscape
            If DGRefNo.Visible = True Then Exit Sub
            Grid_Hide
            KeyCode = 0
            GoTo ExitHere
    End Select
    Exit Sub
End If
If KeyCode = vbKeyR And Shift = 4 Then Exit Sub
If KeyCode = vbKeyT And Shift = 4 Then Exit Sub
If KeyCode = vbKeyDelete And Shift = 2 Then RowDelete Me.ActiveControl.Index: KeyCode = 0: GoTo ExitHere
If KeyCode = vbKeyD And Shift = 2 Then RowDelete Me.ActiveControl.Index: KeyCode = 0: GoTo ExitHere
mCancel = False
If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("TXTADJ_AMT") Then GoTo ExitHere
Select Case KeyCode
    Case vbKeyEscape
        If DGAcHlp.Visible = True Or DGVchrHlp.Visible = True Then
            TxtAcName_Validate Me.ActiveControl.Index, mCancel
            Grid_Hide
            KeyCode = 0
            GoTo ExitHere
        ElseIf DGTDSCODE.Visible = True Or FrameTDS.Visible = True Then
            Grid_Hide
            KeyCode = 0
            GoTo ExitHere
        Else
            FrmHotKey KeyCode, 0
        End If
    Case vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9
            FrmHotKey KeyCode, 0
    Case vbKeyReturn, vbKeyDown, vbKeyUp
        Select Case KeyCode
            Case vbKeyDown, vbKeyUp
                If DGTDSCODE.Visible = True Or DGAcHlp.Visible = True Or DGVchrHlp.Visible = True Then GoTo ExitHere
        End Select
        If ADDFLAG <= 2 Then
            mCancel = False
            If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("VchDt") Then VchDt_Validate Me.ActiveControl.Index, mCancel
            If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("TxtCrDr") Then TxtCrDr_Validate Me.ActiveControl.Index, mCancel
            If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("TxtAcName") Then TxtAcName_Validate Me.ActiveControl.Index, mCancel
            If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("TxtCr") Then TxtCr_Validate Me.ActiveControl.Index, mCancel
            If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("TxtDr") Then Txtdr_Validate Me.ActiveControl.Index, mCancel
            If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("TxtNar") Then TxtNar_Validate Me.ActiveControl.Index, mCancel
            If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("TxtVtYpe") Then TxtVtYpe_Validate Me.ActiveControl.Index, mCancel
            If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("TXTChDate") Then TXTChDate_Validate Me.ActiveControl.Index, mCancel
            If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("TXTClrDate") Then TXTClrDate_Validate Me.ActiveControl.Index, mCancel
            If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("TXTREF") Then TxtRef_Validate (mCancel)
            If TypeOf Me.ActiveControl Is TextBox Then If UCase(Me.ActiveControl.Name) = UCase("TxtTDSCODE") Then DGTDSCODE.Visible = False: TxtTDSCode_Validate Me.ActiveControl.Index, mCancel
            If mCancel = True Then GoTo ExitHere
        End If
        Select Case KeyCode
            Case vbKeyDown, vbKeyReturn
                If TypeOf Me.ActiveControl Is TextBox Then
                    Select Case UCase(Me.ActiveControl.Name)
                        Case UCase("Text1")
                            BTNVLOK_Click
                            GoTo ExitHere
                        Case UCase("TxtTDSAMT")
                            TxtTDSAMT_Validate mCancel
                            If TxtCrDr(Val(FrameTDS.Tag)) = "Dr" Then
                                If RstTds.RecordCount > 0 Then
                                    TxtDr(Val(FrameTDS.Tag)) = Val(TxtDr(Val(FrameTDS.Tag))) + Val(TxtDr(Val(FrameTDS.Tag)).Tag)
                                    RstTds.Sort = "V_SNo ASC"
                                    RstTds.Find "V_SNO=" & Val(FGrid1.TextMatrix(Val(FrameTDS.Tag) + ScrolIndex, 1))
                                    If RstTds.EOF = False Then
                                        Do While RstTds!V_SNo = Val(FGrid1.TextMatrix(Val(FrameTDS.Tag) + ScrolIndex, 1))
                                            TxtDr(Val(FrameTDS.Tag)) = Val(TxtDr(Val(FrameTDS.Tag))) - Val(RstTds!TDSAmt)
                                            TxtDr(Val(FrameTDS.Tag)).Tag = ""
                                            RstTds.MoveNext
                                            If RstTds.EOF = True Then Exit Do
                                        Loop
                                    End If
                                End If
                                TxtDr(Val(FrameTDS.Tag)).SetFocus: GoTo ExitHere
                            Else
                                TxtCr(Val(FrameTDS.Tag)).SetFocus: GoTo ExitHere
                            End If
                            FrameTDS.Visible = False
                        Case UCase("TXTREF")
                            If TxtCrDr(Val(FrameRef.Tag)) = "Cr" Then
                                TxtCr(Val(FrameRef.Tag)).SetFocus: GoTo ExitHere
                            Else
                                TxtDr(Val(FrameRef.Tag)).SetFocus: GoTo ExitHere
                            End If
                            FrameRef.Visible = False
                        Case UCase("TxtCr")
                            If ADDFLAG <= 2 Then
                                If ADDFLAG = 2 Then
                                    mCurrentBalance = PubDatamanFa.FaCalculateOpeningBalance(Me, TxtAcName(Me.ActiveControl.Index).Tag, FixRow, False, RstMain!DocID)
                                Else
                                    mCurrentBalance = PubDatamanFa.FaCalculateOpeningBalance(Me, TxtAcName(Me.ActiveControl.Index).Tag, FixRow, False)
                                End If
                                Set RstFormKeyDown = G_FaCn.Execute("SELECT CreditLimit,NATURE FROM SUBGROUP WHERE SUBCODE=" & FaChk_Text(TxtAcName(Me.ActiveControl.Index).Tag))
                                If RstFormKeyDown.RecordCount > 0 Then
                                    If FaVNull(RstFormKeyDown!CreditLimit) > 0 Then If FaXNull(RstFormKeyDown!Nature) = "Supplier" And mCurrentBalance > 0 And FaXNull(RstEnviro!CreditLimit) = "Yes" Then If FaVNull(RstFormKeyDown!CreditLimit) < Abs(mCurrentBalance) Then MsgBox "Limit Exceed by " + Trim(STR(Abs(mCurrentBalance) - FaVNull(RstFormKeyDown!CreditLimit))), vbExclamation, "Warning"
                                    If FaXNull(RstFormKeyDown!Nature) = "Cash" And mCurrentBalance > 0 And FaXNull(RstEnviro!NegativeCashBalance) = "Yes" Then MsgBox "Negative Cash balance " + Chr(13) + Trim(STR(Abs(mCurrentBalance))) + " " + "Cr", vbExclamation, "Warning"
                                End If
                            End If
                            If mSepNar = "N" Then
                                If ADDFLAG <= 2 Then If askForSave(Me.ActiveControl.Index) = 1 Then GoTo ExitHere
                                NextRow Me.ActiveControl.Index, TxtCr, Me.ActiveControl.Name
                                If ADDFLAG <= 2 Then If Me.ActiveControl.Index = FixRow Then mCancel = PubDatamanFa.FaManageKeysControl(Me, vbKeyUp, 0)
                            End If
                        Case UCase("TxtDr")
                            If ADDFLAG <= 2 Then
                                If ADDFLAG = 2 Then
                                    mCurrentBalance = PubDatamanFa.FaCalculateOpeningBalance(Me, TxtAcName(Me.ActiveControl.Index).Tag, FixRow, False, RstMain!DocID)
                                Else
                                    mCurrentBalance = PubDatamanFa.FaCalculateOpeningBalance(Me, TxtAcName(Me.ActiveControl.Index).Tag, FixRow, False)
                                End If
                                Set RstFormKeyDown = G_FaCn.Execute("SELECT CreditLimit,NATURE FROM SUBGROUP WHERE SUBCODE=" & FaChk_Text(TxtAcName(Me.ActiveControl.Index).Tag))
                                If RstFormKeyDown.RecordCount > 0 Then
                                    If FaVNull(RstFormKeyDown!CreditLimit) > 0 Then
                                        If FaXNull(RstFormKeyDown!Nature) = "Customer" And mCurrentBalance < 0 And FaXNull(RstEnviro!DebitLimit) = "Yes" Then If FaVNull(RstFormKeyDown!CreditLimit) < Abs(mCurrentBalance) Then MsgBox "Limit Exceed by " + Trim(STR(Abs(mCurrentBalance) - FaVNull(RstFormKeyDown!CreditLimit))), vbExclamation, "Warning"
                                    End If
                                End If
                            End If
                            If mSepNar = "N" Then
                                If ADDFLAG <= 2 Then If askForSave(Me.ActiveControl.Index) = 1 Then GoTo ExitHere
                                NextRow Me.ActiveControl.Index, TxtCr, Me.ActiveControl.Name
                                If ADDFLAG <= 2 Then If Me.ActiveControl.Index = FixRow Then mCancel = PubDatamanFa.FaManageKeysControl(Me, vbKeyUp, 0)
                            End If
                        Case UCase("TxtNar")
                            If ADDFLAG <= 2 Then If askForSave(Me.ActiveControl.Index) = 1 Then GoTo ExitHere
                            NextRow Me.ActiveControl.Index, TxtNar, Me.ActiveControl.Name
                            If ADDFLAG <= 2 Then If Me.ActiveControl.Index = FixRow Then mCancel = PubDatamanFa.FaManageKeysControl(Me, vbKeyUp, 0)
                    End Select
                End If
            Case vbKeyUp
                If TypeOf Me.ActiveControl Is TextBox Then
                    Select Case UCase(Me.ActiveControl.Name)
                        Case UCase("TxtCrDr")
                            If ScrolIndex > 0 Then previousRow Me.ActiveControl.Index: GoTo ExitHere
                    End Select
                End If
        End Select
        mMoveFlag = True
        Select Case KeyCode
            Case vbKeyReturn, vbKeyDown, vbKeyUp
                If PubDatamanFa.FaManageKeysControl(Me, KeyCode, Shift) = True Then
                    If ADDFLAG <= 2 Then
                        If Val(LblDrAmt(0).CAPTION) + Val(LblCrAmt(0).CAPTION) > 0 Then
                            If Val(LblDrAmt(0).CAPTION) = Val(LblCrAmt(0).CAPTION) Then
                                If MsgBox("Save Yes/No", vbYesNo + vbInformation + vbDefaultButton1, Me.CAPTION) = vbYes Then
                                    mMoveFlag = False
                                    KeyCode = 0
                                    TopCtrl1_eSave
                                    GoTo ExitHere
                                Else
                                    mMoveFlag = False
                                    KeyCode = 0
                                    TxtCrDr(0).SetFocus
                                    GoTo ExitHere
                                End If
                            End If
                        End If
                    End If
                End If
        End Select
        mMoveFlag = False
        KeyCode = 0
ExitHere:
        If TypeOf Me.ActiveControl Is TextBox Then
            Select Case UCase(Me.ActiveControl.Name)
                Case UCase("TxtCrDr"), UCase("TxtAcName"), UCase("TxtCr"), UCase("TxtDr"), UCase("TxtNar")
                    PicDisp Me.ActiveControl.Index
                Case Else
                    PicFalse
            End Select
        End If
    Case Else
        PicFalse
        FaFormKeyDown Me, KeyCode, Shift
End Select
Set RstFormKeyDown = Nothing
End Sub
Private Sub Form_Load()
Dim I As Integer, RST1 As ADODB.Recordset, mSiteHlp As String, mSiteHlpAnd As String
mSiteHlp = ""
mSiteHlpAnd = ""
If PubSiteCodeWiseHelp = True Then
    mSiteHlp = "Where Site_Code='" & PubSiteCode & "'"
    mSiteHlpAnd = "And Site_Code='" & PubSiteCode & "'"
End If
    TopCtrl1.Tag = "AEDP": TopCtrl1.TopText1 = Me.CAPTION
    If PubSec = "SANJEEV" Then
        If rsUserPerm.RecordCount > 0 Then
            rsUserPerm.MoveFirst
            rsUserPerm.Find ("FORM_NAME='" & Me.CAPTION & "'")
            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
        End If
    ElseIf PubSec = "RAHUL" Then
        If rsUserPerm.RecordCount > 0 Then
            rsUserPerm.MoveFirst
            rsUserPerm.Find ("FORM_CODE='" & Me.Name & "'")
            If Not rsUserPerm.EOF Then TopCtrl1.Tag = rsUserPerm!param_str Else TopCtrl1.Tag = "****"
        End If
    End If
    FixRow = 4
    mMoveFlag = False
    Me.height = 7065
    Me.width = 11900
    Me.top = 0
    Me.left = 0
    ADDFLAG = 3
    ''''''''''''''''''''''
    PubDatamanFa.FaBackEnd = PubBackEnd
    PubDatamanFa.FaPubLoginDate = PubLoginDate
    PubDatamanFa.FaPubDivCode = PubDivCode
    PubDatamanFa.FaPubSiteCode = PubSiteCode
    PubDatamanFa.FaPubSiteCodeDisplay = PubSiteCodeDisplay
    PubDatamanFa.FaPubSiteName = PubSiteName
    PubDatamanFa.FapubUName = pubUName
    PubDatamanFa.FaDosPort = PubFaDosPort
    PubDatamanFa.FaRunPIF = PubRunPIF
    PubDatamanFa.FaPubSiteType = PubFaSiteType
    Set PubDatamanFa.SetG_FaCn = G_FaCn
    Set PubDatamanFa.SetG_CompCn = G_CompCn
    Set PubDatamanFa.SetrsUserPerm = rsUserPerm.Clone
    Set PubDatamanFa.SetMasterRst = FaMasterRst.Clone
    ''''''''''''''''''''''
    If PubFaSiteType <> 0 Then
        Set TxtSite.RowSource = G_FaCn.Execute("SELECT SITE_CODE,SITE_DESC FROM SITE ORDER BY SITE_DESC")
        TxtSite.ListField = "SITE_DESC"
        TxtSite.BoundColumn = "SITE_cODE"
        TxtSite.Tag = "SELECT SITE_CODE,SITE_DESC FROM SITE ORDER BY SITE_DESC"
        TxtSite.BoundText = PubSiteCode
    End If
    Set RstEnviro = G_FaCn.Execute("SELECT V_TYPE,DESCRIPTION FROM VOUCHER_TYPE WHERE Category='FA' ORDER BY DESCRIPTION")
    Set DataCombo3.RowSource = RstEnviro
    DataCombo3.ListField = "DESCRIPTION"
    DataCombo3.BoundColumn = "V_TYPE"
    DataCombo3.Tag = "SELECT V_TYPE,DESCRIPTION FROM VOUCHER_TYPE WHERE Category='FA' ORDER BY DESCRIPTION"
    
    Set DataCombo1.RowSource = RstEnviro
    DataCombo1.ListField = "DESCRIPTION"
    DataCombo1.BoundColumn = "V_TYPE"
    DataCombo1.Tag = "SELECT V_TYPE,DESCRIPTION FROM VOUCHER_TYPE WHERE Category='FA' ORDER BY DESCRIPTION"

    Set RstEnviro = G_FaCn.Execute("SELECT * FROM FAENVIRO")
    If RstEnviro.RecordCount <= 0 Then MsgBox "Parameter Not Set": Exit Sub
   
    TopCtrl1.TopText1 = "Voucher Entry"
    If PubBackEnd = "A" Then
        If RstEnviro!FilterAC = "No" Then
            Set RstAcHlpDr = New ADODB.Recordset
            Set RstAcHlpCr = New ADODB.Recordset
            If RstEnviro!ShowCityName = "Yes" Then
                RstAcHlpDr.Open ("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3))+RTRIM(IIF(ISNULL(CityName),'',','+CityName)) AS NameWithADDR,IIF(ISNULL(FNAME),'',FNAME) AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
                Set RstAcHlpCr = RstAcHlpDr.Clone
            Else
                RstAcHlpDr.Open ("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES,RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3))+RTRIM(IIF(ISNULL(CityName),'',','+CityName)) AS NameWithADDR,IIF(ISNULL(FNAME),'',FNAME) AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
                Set RstAcHlpCr = RstAcHlpDr.Clone
            End If
        End If
    Else
        If RstEnviro!FilterAC = "No" Then
            Set RstAcHlpDr = New ADODB.Recordset
            Set RstAcHlpCr = New ADODB.Recordset
            If RstEnviro!ShowCityName = "Yes" Then
                RstAcHlpDr.Open ("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,''))+RTRIM(ISNULL(CityName,'')) AS NameWithADDR,ISNULL(FNAME,'') AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
                Set RstAcHlpCr = RstAcHlpDr.Clone
            Else
                RstAcHlpDr.Open ("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,''))+RTRIM(ISNULL(CityName,'')) AS NameWithADDR,ISNULL(FNAME,'') AS FatherName From ViewSubgroup " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
                Set RstAcHlpCr = RstAcHlpDr.Clone
            End If
        End If
    End If
    
    Set RstTDSHlp = New ADODB.Recordset
    RstTDSHlp.Open ("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES,'' AS NameWithADDR From ViewSubgroup WHERE NATURE='T.D.S.' " & mSiteHlpAnd & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
    Set RstMain = New ADODB.Recordset
    
    If PubSiteCodeWiseMasterRst = False Then
        If PubFaSiteType = 1 Then
            Label18.Visible = True
            TxtSite.Visible = True
            RstMain.Open "SELECT DOCID FROM LEDGERM WHERE V_TYPE IN (SELECT V_TYPE FROM VOUCHER_tYPE WHERE Category='FA') AND V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(PubEndDate) & "  ORDER BY V_Date,DOCID", G_FaCn, adOpenDynamic, adLockOptimistic
        Else
            Label18.Visible = False
            TxtSite.Visible = False
            RstMain.Open "SELECT DOCID FROM LEDGERM WHERE V_TYPE IN (SELECT V_TYPE FROM VOUCHER_tYPE WHERE Category='FA') AND V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(PubEndDate) & "  ORDER BY V_Date,DOCID", G_FaCn, adOpenDynamic, adLockOptimistic
        End If
    Else
        If PubFaSiteType = 1 Then
            If PubSeparateVrNoForSite = 1 Then
                Label18.Visible = False
                TxtSite.Visible = True
                TxtSite.Enabled = False
                TxtSite.BoundText = PubSeparateLogSite
                RstMain.Open "SELECT DOCID FROM LEDGERM WHERE V_TYPE IN (SELECT V_TYPE FROM VOUCHER_tYPE WHERE Category='FA') AND V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(PubEndDate) & " AND RIGHT(SITE_CODE,1)='" & Trim(PubSeparateLogSite) & "' ORDER BY V_Date,DOCID", G_FaCn, adOpenDynamic, adLockOptimistic
            Else
                Label18.Visible = True
                TxtSite.Visible = True
                RstMain.Open "SELECT DOCID FROM LEDGERM WHERE V_TYPE IN (SELECT V_TYPE FROM VOUCHER_tYPE WHERE Category='FA') AND V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(PubEndDate) & "  ORDER BY V_Date,DOCID", G_FaCn, adOpenDynamic, adLockOptimistic
            End If
        ElseIf PubFaSiteType = 2 Then
            Label18.Visible = True
            TxtSite.Visible = True
            RstMain.Open "SELECT DOCID FROM LEDGERM WHERE V_TYPE IN (SELECT V_TYPE FROM VOUCHER_tYPE WHERE Category='FA') AND V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(PubEndDate) & "  AND SITE_CODE='" & Trim(PubSiteCode) & "' ORDER BY V_Date,DOCID", G_FaCn, adOpenDynamic, adLockOptimistic
        Else
            Label18.Visible = False
            TxtSite.Visible = False
            RstMain.Open "SELECT DOCID FROM LEDGERM WHERE V_TYPE IN (SELECT V_TYPE FROM VOUCHER_tYPE WHERE Category='FA') AND V_DATE BETWEEN " & FaConvertDate(PubStartDate) & " AND " & FaConvertDate(PubEndDate) & "  ORDER BY V_Date,DOCID", G_FaCn, adOpenDynamic, adLockOptimistic
        End If
    End If
    TXTADJ_AMT.Visible = False
    FRAMEADJUST.Visible = False
    Me.BackColor = &HE0E0E0
    SETS "INI", Me, RstMain
    If RstMain.RecordCount > 0 Then
        If PubFaSiteType = 2 Then
            Set RST1 = G_FaCn.Execute("SELECT DOCID FROM LastVoucher where U_Name='" & pubUName & "' AND SITE_cODE='" & PubSiteCode & "' ORDER BY U_EntDt DESC")
            If RST1.RecordCount > 0 Then
                If FaXNull(RST1!DocID) <> "" Then
                    RstMain.Find "DocId='" & RST1!DocID & "'"
                End If
            End If
        Else
            Set RST1 = G_FaCn.Execute("SELECT DOCID FROM LastVoucher where U_Name='" & pubUName & "' ORDER BY U_EntDt DESC")
            If RST1.RecordCount > 0 Then
                If FaXNull(RST1!DocID) <> "" Then
                    RstMain.Find "DocId='" & RST1!DocID & "'"
                End If
            End If
        End If
    End If
    Ini_Grid
    MoveRec
    LblAmtRs = ""
    Set RST1 = Nothing
End Sub
Private Sub MoveRec()
Dim MyRs As ADODB.Recordset, K As Integer, J As Integer, RST1 As ADODB.Recordset, RstX As ADODB.Recordset
'On Error GoTo Errloop
    mMoveFlag = True
    For J = 0 To FixRow
        MakeVisible J, False
    Next
    LblDrAmt(0).CAPTION = 0
    LblCrAmt(0).CAPTION = 0
    MakeEmpty
    ScrolIndex = 0
    LockFields True
    If RstMain.RecordCount <= 0 Then
        mNCat = "JV"
        VchDt(0) = PubLoginDate
        FaVrTypeSetting ""
        Set RstMainAdj = New ADODB.Recordset
        Set RstMainAdj = PubDatamanFa.FaAdjustRst(RstMainAdj)
        Set RstRef = New ADODB.Recordset
        Set RstRef = PubDatamanFa.FaRefRst(RstRef)
        Set RstTds = New ADODB.Recordset
        Set RstTds = PubDatamanFa.FaTDSRst(RstTds)
    Else
        If RstMain.EOF = True Or RstMain.BOF = True Then
            If RstMain.RecordCount > 0 Then
                If RstMain.EOF = True Then RstMain.MoveLast
                If RstMain.BOF = True Then RstMain.MoveFirst
            Else
                mMoveFlag = False: GoTo ExitHere
            End If
        End If
        J = 0
        Set RstMainAdj = New ADODB.Recordset
        Set RstMainAdj = PubDatamanFa.FaAdjustRst(RstMainAdj)
        Set RstRef = New ADODB.Recordset
        Set RstRef = PubDatamanFa.FaRefRst(RstRef)
        Set RstTds = New ADODB.Recordset
        Set RstTds = PubDatamanFa.FaTDSRst(RstTds)
        Set MyRs = New Recordset
        If RstEnviro!ShowCityName = "Yes" Then
            MyRs.Open "SELECT LEDGER.*,NameWithCity AS Name,ViewSubgroup.CURR_BAL FROM (LEDGER LEFT JOIN ViewSubgroup ON ViewSubgroup.SUBCODE=LEDGER.SUBCODE) LEFT JOIN VOUCHER_TYPE ON VOUCHER_TYPE.V_tYPE=LEDGER.V_TYPE WHERE LEDGER.DOCID='" & RstMain!DocID & "' ORDER BY LEDGER.V_SNO", G_FaCn, adOpenForwardOnly, adLockReadOnly
        Else
            MyRs.Open "SELECT LEDGER.*,SUBGROUP.NAME,SUBGROUP.CURR_BAL FROM (LEDGER LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGER.SUBCODE) LEFT JOIN VOUCHER_TYPE ON VOUCHER_TYPE.V_tYPE=LEDGER.V_TYPE WHERE LEDGER.DOCID='" & RstMain!DocID & "' ORDER BY LEDGER.V_SNO", G_FaCn, adOpenForwardOnly, adLockReadOnly
        End If
        If MyRs.RecordCount > 0 Then
            Set RstX = New ADODB.Recordset
            RstX.Open "SELECT LEDGERM.*,VOUCHER_TYPE.NCAT FROM LEDGERM LEFT JOIN VOUCHER_TYPE ON VOUCHER_TYPE.V_tYPE=LEDGERM.V_TYPE WHERE DOCID='" & RstMain!DocID & "' ORDER BY V_Date,DOCID", G_FaCn, adOpenForwardOnly, adLockReadOnly
            If RstX.RecordCount > 0 Then
                TxtVno(0).TEXT = RstX!V_NO
                VchDt(0) = RstX!V_Date
                LblDay = PubDatamanFa.FaRetDayFunc(VchDt(0))
                TxtGlb(0).TEXT = FaXNull(RstX!Narration)
                LblVPrefix.CAPTION = RstX!V_Prefix
                mLastPrefix = RstX!V_Prefix
                mNCat = RstX!NCat
            End If
            If Trim(mNCat) = "" Then MsgBox "Nature Category Not set"
            FaVrTypeSetting MyRs!V_tYPE, LblVPrefix
            TxtCHno(0) = FaXNull(MyRs!Chq_No)
            TXTChDate(0) = FaXNull(MyRs!Chq_Date)
            TXTClrDate(0) = FaXNull(MyRs!clg_date)
            If PubFaSiteType = 1 Then
                TxtSite.BoundText = Right(MyRs!SITE_CODE, 1)
            Else
                TxtSite.BoundText = Trim(MyRs!SITE_CODE)
            End If
            Do Until MyRs.EOF
                If MyRs!AmtCr > 0 Then
                    Set RST1 = New ADODB.Recordset
                    RST1.Open "SELECT * FROM LEDGERADJ WHERE DOCID1='" & RstMain!DocID & "' AND V_SNo1=" & MyRs!V_SNo, G_FaCn, adOpenForwardOnly, adLockReadOnly
                    Do Until RST1.EOF
                        With RstMainAdj
                            .AddNew
                            .Fields("DocId") = RST1!DocID2
                            .Fields("V_SNo") = RST1!V_SNo2
                            .Fields("VSNo") = RST1!V_SNo1
                            .Fields("DR") = RST1!cr
                            .Fields("SUBCODE") = RST1!SubCode
                            .Fields("AgRefNo") = RST1!AgRefNo
                            .Update
                        End With
                        RST1.MoveNext
                    Loop
                Else
                    Set RST1 = New ADODB.Recordset
                    RST1.Open "SELECT * FROM LEDGERADJ WHERE DOCID2='" & RstMain!DocID & "' AND V_SNo2=" & MyRs!V_SNo, G_FaCn, adOpenForwardOnly, adLockReadOnly
                    Do Until RST1.EOF
                        With RstMainAdj
                            .AddNew
                            .Fields("DocId") = RST1!DocID1
                            .Fields("V_SNo") = RST1!V_SNo1
                            .Fields("VSNo") = RST1!V_SNo2
                            .Fields("CR") = RST1!cr
                            .Fields("SUBCODE") = RST1!SubCode
                            .Fields("AgRefNo") = RST1!AgRefNo
                            .Update
                        End With
                        RST1.MoveNext
                    Loop
                End If
                Set RST1 = New ADODB.Recordset
                RST1.Open "SELECT * FROM LedgerRef WHERE DOCID='" & RstMain!DocID & "' AND V_SNo=" & MyRs!V_SNo, G_FaCn, adOpenForwardOnly, adLockReadOnly
                Do Until RST1.EOF
                    With RstRef
                        .AddNew
                        .Fields("DocId") = RST1!DocID
                        .Fields("V_SNo") = RST1!V_SNo
                        .Fields("DR") = RST1!dr
                        .Fields("CR") = RST1!cr
                        .Fields("SUBCODE") = RST1!SubCode
                        .Fields("AgRefNo") = RST1!AgRefNo
                        .Fields("AgRefType") = RST1!AgRefType
                        If Not IsNull(RST1!DUEDATE) Then
                            .Fields("DueDate") = RST1!DUEDATE
                        End If
                        .Update
                    End With
                    RST1.MoveNext
                Loop
                Set RST1 = New ADODB.Recordset
                RST1.Open "SELECT LEDGERTDS.*,SUBGROUP.NAME AS TDSNAME FROM LEDGERTDS LEFT JOIN SUBGROUP ON SUBGROUP.SUBCODE=LEDGERTDS.TDSCODE WHERE DOCID='" & RstMain!DocID & "' AND V_SNo=" & MyRs!V_SNo, G_FaCn, adOpenForwardOnly, adLockReadOnly
                Do Until RST1.EOF
                    With RstTds
                        .AddNew
                        .Fields("DocId") = RST1!DocID
                        .Fields("V_SNo") = RST1!V_SNo
                        .Fields("TDSCode") = RST1!TDSCODE
                        .Fields("TDSNAME") = RST1!TDSNAME
                        .Fields("TDSDrCode") = RST1!TDSDRCODE
                        .Fields("TDSYN") = RST1!TDSYN
                        .Fields("ONAMT") = RST1!ONAmt
                        .Fields("TDS") = RST1!TDS
                        .Fields("TDSAMT") = RST1!TDSAmt
                        .Fields("TDSPOST") = RST1!TDSPOST
                        .Fields("TDSDocId") = RST1!TDSDocId
                        .Fields("TDSV_SNo") = RST1!TDSV_Sno
                        Set RstX = New ADODB.Recordset
                        RstX.Open "SELECT * FROM LEDGER WHERE DOCID='" & RST1!TDSDocId & "' AND V_SNo=" & RST1!TDSV_Sno, G_FaCn, adOpenForwardOnly, adLockOptimistic
                        If RstX.RecordCount > 0 Then
                            .Fields("NARRATION") = RstX!Narration
                        End If
                        .Update
                    End With
                    RST1.MoveNext
                Loop
                FGrid1.AddItem "" & Chr(9) & MyRs!V_SNo & Chr(9) & MyRs!SubCode & Chr(9) & MyRs!Name & Chr(9) & "" & Chr(9) & Format(MyRs!AmtDr, "0.00") & Chr(9) & Format(MyRs!AmtCr, "0.00") & Chr(9) & IIf(MyRs!AmtCr > 0, "Cr", "Dr") & Chr(9) & MyRs!Narration & Chr(9) & Format(MyRs!Curr_Bal, "0.00")
                LblDrAmt(0).CAPTION = Val(LblDrAmt(0).CAPTION) + Val(MyRs!AmtDr)
                LblCrAmt(0).CAPTION = Val(LblCrAmt(0).CAPTION) + Val(MyRs!AmtCr)
                If J <= FixRow Then
                    MakeVisible J, True
                    TxtCrDr(J).TEXT = IIf(MyRs!AmtCr > 0, "Cr", "Dr")
                    TxtCrDr(J).Tag = MyRs!V_SNo
                    If TxtCrDr(J).TEXT = "Cr" Then
                        TxtCr(J).TEXT = Format(MyRs!AmtCr, "0.00")
                        TxtCr(J).Visible = True
                        TxtDr(J).TEXT = ""
                        TxtDr(J).Visible = False
                    Else
                        TxtDr(J).TEXT = Format(MyRs!AmtDr, "0.00")
                        TxtDr(J).Visible = True
                        TxtCr(J).TEXT = ""
                        TxtCr(J).Visible = False
                    End If
                    TxtAcName(J).Tag = MyRs!SubCode
                    TxtAcName(J).TEXT = FaXNull(MyRs!Name)
                    TxtNar(J).TEXT = FaXNull(MyRs!Narration)
                    If MyRs!Curr_Bal = 0 Then
                        LblCb(J) = ""
                    Else
                        LblCb(J) = "Current Balance :" + Format(Abs(FaVNull(MyRs!Curr_Bal)), "0.00") + " " + IIf(MyRs!Curr_Bal > 0, "Cr", "Dr")
                    End If
                    J = J + 1
                    K = K + 1
                End If
                MyRs.MoveNext
            Loop
        End If
        If Me.Visible = True Then If TxtCrDr(0).Visible = True Then TxtCrDr(0).SetFocus
    End If
    LblDrAmt(0).CAPTION = Format(Val(LblDrAmt(0).CAPTION), "0.00")
    LblCrAmt(0).CAPTION = Format(Val(LblCrAmt(0).CAPTION), "0.00")
    mMoveFlag = False
ExitHere:
    Set MyRs = Nothing
    Set RST1 = Nothing
    Set RstX = Nothing
    Exit Sub
Errloop:            MsgBox err.Description, vbInformation, Me.CAPTION
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
If FRAMEVLIST.Visible = True Then Exit Sub
If RstMain.RecordCount > 0 Then
    ADDFLAG = 2
    SETS "EDIT", Me, RstMain
    LockFields False
    If TypeOf Me.ActiveControl Is TextBox Then
    Else
        TxtCrDr(0).SetFocus
    End If
End If
Exit Sub
Errloop:        MsgBox err.Description, vbInformation, Me.CAPTION
End Sub
Private Sub TopCtrl1_eDel()
Dim I As Integer, XBM, RST1 As ADODB.Recordset
On Error GoTo Errloop
    If FRAMEVLIST.Visible = True Then GoTo ExitHere
    Set RST1 = G_FaCn.Execute("SELECT * FROM LEDGERTDS WHERE DOCID='" & RstMain!DocID & "' AND TDSPOST='Y'")
    If RST1.RecordCount > 0 Then
        MsgBox "T.D.S.Challan Already Made,Can't delete it"
        GoTo ExitHere
    End If
    XBM = RstMain.Bookmark
    If RstMain.RecordCount > 0 Then
        If MsgBox("Are sure to delete it", vbYesNo + vbCritical + vbDefaultButton1, Me.CAPTION) = vbYes Then
            G_FaCn.BeginTrans
            On Error Resume Next
            FaDeleteTrack G_FaCn, RstMain!DocID
            On Error GoTo Errloop
            Set RST1 = G_FaCn.Execute("SELECT * FROM LEDGERTDS WHERE DOCID='" & RstMain!DocID & "'")
            Do Until RST1.EOF
                FaCalCurrBal G_FaCn, RST1!TDSCODE, RST1!TDSAmt, 0
                FaCalCurrBal G_FaCn, RST1!TDSDRCODE, 0, RST1!TDSAmt
                G_FaCn.Execute "DELETE FROM LEDGER WHERE DOCID='" & RST1!TDSDocId & "'"
                RST1.MoveNext
            Loop
            If FGrid1.Rows > 0 Then
                For I = 0 To FGrid1.Rows - 1
                    If FGrid1.TextMatrix(I, 7) = "Cr" Then
                        FaCalCurrBal G_FaCn, FGrid1.TextMatrix(I, 2), Val(FGrid1.TextMatrix(I, 6)), 0
                    ElseIf FGrid1.TextMatrix(I, 7) = "Dr" Then
                        FaCalCurrBal G_FaCn, FGrid1.TextMatrix(I, 2), 0, Val(FGrid1.TextMatrix(I, 5))
                    End If
                Next
            End If
            G_FaCn.Execute "DELETE FROM LEDGER  WHERE DOCID='" & RstMain!DocID & "'"
            G_FaCn.Execute "DELETE FROM LEDGERM WHERE DocID='" & RstMain!DocID & "'"
            G_FaCn.Execute "DELETE FROM LEDGERADJ WHERE DocID1='" & RstMain!DocID & "' OR DocID2='" & RstMain!DocID & "'"
            G_FaCn.Execute "DELETE FROM LEDGERREF WHERE DocID='" & RstMain!DocID & "'"
            G_FaCn.Execute "DELETE FROM LEDGERTDS WHERE DocID='" & RstMain!DocID & "'"
            G_FaCn.CommitTrans
            RstMain.Requery
            If RstMain.RecordCount >= XBM Then
                RstMain.Bookmark = XBM
            Else
                If RstMain.EOF = False Then RstMain.MoveLast
            End If
            MoveRec
            BUTTONS True, Me, RstMain, 0
        End If
    Else
        MsgBox "There Is No Record To Delete.", vbInformation, Me.CAPTION
    End If
ExitHere:
    Set RST1 = Nothing
    Exit Sub
Errloop:            G_FaCn.RollbackTrans
                    MsgBox err.Description, vbCritical, Me.CAPTION
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
On Error GoTo ELoop
    If RstMain.RecordCount <= 0 Then MsgBox "No Records To Search", vbInformation, Me.CAPTION: Exit Sub
    If (PubFaSiteType = 1 And PubSeparateVrNoForSite = 1) Then
        If RstEnviro!ShowCityName = "Yes" Then
            Set Dcparty.RowSource = G_FaCn.Execute("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES From ViewSubgroup WHERE SITE_CODE='" & PubSeparateLogSite & "' order by Name")
        Else
            Set Dcparty.RowSource = G_FaCn.Execute("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES From ViewSubgroup WHERE SITE_CODE='" & PubSeparateLogSite & "' order by Name")
        End If
    ElseIf PubSiteCodeWiseHelp = True Then
        If RstEnviro!ShowCityName = "Yes" Then
            Set Dcparty.RowSource = G_FaCn.Execute("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES From ViewSubgroup WHERE SITE_CODE='" & PubSiteCode & "' order by Name")
        Else
            Set Dcparty.RowSource = G_FaCn.Execute("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES From ViewSubgroup WHERE SITE_CODE='" & PubSiteCode & "' order by Name")
        End If
    Else
        If RstEnviro!ShowCityName = "Yes" Then
            Set Dcparty.RowSource = G_FaCn.Execute("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES From ViewSubgroup order by Name")
        Else
            Set Dcparty.RowSource = G_FaCn.Execute("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES From ViewSubgroup order by Name")
        End If
    End If
    Dcparty.ListField = "Name"
    Dcparty.BoundColumn = "Subcode"
    TXTVDATE1 = ""
    TXTVDATE2 = ""
    Dcparty.BoundText = ""
    DataCombo1.BoundText = ""
    Text1 = ""
    FRAMEVLIST.Visible = True
    FRAMEVLIST.ZOrder 0
    TXTVDATE1.SetFocus
Exit Sub
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Sub
Private Sub TopCtrl1_ePrn()
Frame1(1).Visible = True
Frame1(1).left = 1845
Frame1(1).top = 675
Frame1(1).ZOrder 0
End Sub
Private Sub TopCtrl1_eSave()
Dim MyCheck As Byte, I As Integer
On Error GoTo Errloop
    If FaIsValid(VchDt(0), "Date") = False Then Exit Sub
    If PubDatamanFa.FaCheckFinYearFunc(VchDt(0)) = False Then Exit Sub
    If FaIsValid(TxtVno(0), "Voucher No") = False Then Exit Sub
    If PubFaSiteType = 1 Then If FaIsValid(TxtSite, "For Site") = False Then Exit Sub
    TxtVno(0) = Val(TxtVno(0))
    MyCheck = LedPost(IIf(ADDFLAG = 1, "A", "E"))
    If MyCheck = 2 Then
        If MyCheck = 2 Then If MsgBox("Vr.No.Already Exist,Generate New Vr.No.", vbDefaultButton1 + vbYesNo) = vbNo Then Exit Sub
        Do While True
            TxtVno(0) = Val(TxtVno(0)) + 1
            MyCheck = LedPost(IIf(ADDFLAG = 1, "A", "E"))
            If MyCheck <> 2 Then MsgBox "Vr.No.Already Exist,New Vr.No. is " + Trim(TxtVno(0)): Exit Do
        Loop
        mLastVrType = TxtVtYpe(0).Tag
        TopCtrl1_eAdd
    ElseIf MyCheck = 0 Then
        If ADDFLAG = 2 Then
            ADDFLAG = 3
            SETS "INI", Me, RstMain
        Else
            mLastVrType = TxtVtYpe(0).Tag
            TopCtrl1_eAdd
        End If
    End If
Exit Sub
Errloop:    MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub TopCtrl1_eCancel()
Dim I As Integer
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo + vbCritical + vbDefaultButton1, Me.CAPTION) = vbYes Then
        ADDFLAG = 3
        Grid_Hide
        SETS "INI", Me, RstMain
        For I = 0 To FixRow
            MakeVisible I, False
        Next
        MoveRec
    End If
Exit Sub
Errloop:        MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub LblShort_Click(Index As Integer)
Dim I As Integer
    If ADDFLAG = 3 Then Exit Sub
    For I = 1 To 4
        LblShort(I).FontBold = False
        LblShort(I).ForeColor = &HFFFF&
    Next
    LblShort(Index).ForeColor = &HFFFF00
    LblShort(Index).FontBold = True
    FrmHotKey Index, Index
End Sub
Public Sub Disp_Text(Enb As Boolean)
    '''
End Sub
Private Function VchTrnSaveLst(ByRef Index As Integer) As Byte
If ADDFLAG = 3 Then Exit Function
VchTrnSaveLst = PubDatamanFa.FaVchTrnSaveLst(Me, FixRow, Index, ScrolIndex)
End Function
Private Sub previousRow(Index As Integer)
If Index >= 1 And Index <= FixRow Then
    If mSepNar = "Y" Then
        If TxtNar(Index - 1).Visible = False Then TxtNar(Index - 1).Visible = True
        TxtNar(Index - 1).SetFocus
    ElseIf mSepNar = "N" Then
        Select Case Index
            Case Is > 0
                If TxtCrDr(Index - 1).TEXT = "Dr" Then
                    If TxtDr(Index - 1).Visible = False Then TxtDr(Index - 1).Visible = True
                    TxtDr(Index - 1).SetFocus
                ElseIf TxtCrDr(Index - 1).TEXT = "Cr" Then
                    If TxtCr(Index - 1).Visible = False Then TxtCr(Index - 1).Visible = True
                    TxtCr(Index - 1).SetFocus
                End If
        End Select
    End If
ElseIf Index = 0 Then
    If ScrolIndex = 0 Then Exit Sub
    ScrolIndex = ScrolIndex - 1
    ScrollFiller "Previous", 0
    If mSepNar = "Y" Then
        If TxtNar(Index).Visible = False Then TxtNar(Index).Visible = True
        TxtNar(Index).SetFocus
    ElseIf mSepNar = "N" Then
        If TxtCrDr(Index).TEXT = "Dr" Then
            If TxtDr(Index).Visible = False Then TxtDr(Index).Visible = True
            TxtDr(Index).SetFocus
        ElseIf TxtCrDr(Index).TEXT = "Cr" Then
            If TxtCr(Index).Visible = False Then TxtCr(Index).Visible = True
            TxtCr(Index).SetFocus
        End If
    End If
End If
End Sub
Private Function askForSave(Index As Integer) As Byte
askForSave = 0
If Val(LblDrAmt(0).CAPTION) + Val(LblCrAmt(0).CAPTION) > 0 Then
    If Val(LblDrAmt(0).CAPTION) = Val(LblCrAmt(0).CAPTION) Then
        If mCommNar = "Y" Then
            If Index + ScrolIndex = FGrid1.Rows - 1 Then
                TxtGlb(0).SetFocus
                askForSave = 1
                Exit Function
            End If
        Else
            If Index + ScrolIndex = FGrid1.Rows - 1 Then
                If ADDFLAG <= 2 Then
                    If MsgBox("Save Yes/No", vbYesNo + vbDefaultButton1 + vbInformation, Me.CAPTION) = vbYes Then TopCtrl1_eSave
                End If
                askForSave = 1
                Exit Function
            End If
        End If
    End If
End If
End Function
Private Sub NextRow(Index As Integer, Cntrl As Object, CntrlName As String)
Dim MyCheck As Byte
mMoveFlag = True
If TxtAcName(Index).Tag = "" And TxtAcName(Index).TEXT = "" And (Val(TxtCr(Index)) = 0 Or Val(TxtDr(Index)) = 0) Then Exit Sub
If ADDFLAG = 3 Then If Index + ScrolIndex = FGrid1.Rows - 1 Then Exit Sub
If Index = FixRow Then
    ScrolIndex = ScrolIndex + 1
    If Val(LblDrAmt(0).CAPTION) + Val(LblCrAmt(0).CAPTION) > 0 Then ScrollFiller "Next", Index
    Index = Index - 1
End If
If Val(LblDrAmt(0).CAPTION) = Val(LblCrAmt(0).CAPTION) Then
    If TxtCrDr(Index + 1) = "Dr" Then TxtCr(Index + 1).Visible = False
    If TxtCrDr(Index + 1) = "Cr" Then TxtDr(Index + 1).Visible = False
Else
    If TxtDr(Index + 1) = "" And TxtCr(Index + 1) = "" Then
        If Val(LblCrAmt(0).CAPTION) > Val(LblDrAmt(0).CAPTION) Then TxtCrDr(Index + 1).TEXT = "Dr": TxtCr(Index + 1).Visible = False
        If Val(LblDrAmt(0).CAPTION) > Val(LblCrAmt(0).CAPTION) Then TxtCrDr(Index + 1).TEXT = "Cr": TxtDr(Index + 1).Visible = False
    End If
End If
MakeVisible Index + 1, True
If Index + 1 = FixRow Then TxtCrDr(Index + 1).SetFocus
End Sub
Private Sub ScrollFiller(ScrollType As String, Index As Integer)
Dim K As Integer, mFixRow As Integer
mFixRow = FixRow
Select Case ScrollType
    Case "NextAdd"
        mFixRow = FixRow - 1
    Case "Previous"
    Case "Current"
End Select
For K = 0 To mFixRow
    If ScrolIndex + K > FGrid1.Rows - 1 Then
        MakeVisible K, False
        TxtAcName(K).Tag = ""
        TxtAcName(K).TEXT = ""
        TxtCr(K).TEXT = ""
        TxtDr(K).TEXT = ""
        TxtCrDr(K).TEXT = ""
        LblCb(K) = ""
    Else
        MakeVisible K, True
        TxtCrDr(K).TEXT = FGrid1.TextMatrix(ScrolIndex + K, 7)
        If TxtCrDr(K).TEXT = "Cr" Then
            RstAcHlpCr.MoveFirst
            RstAcHlpCr.Find "SUBCODE='" & FGrid1.TextMatrix(ScrolIndex + K, 2) & "'"
        Else
            RstAcHlpDr.MoveFirst
            RstAcHlpDr.Find "SUBCODE='" & FGrid1.TextMatrix(ScrolIndex + K, 2) & "'"
        End If
        TxtAcName(K).Tag = FGrid1.TextMatrix(ScrolIndex + K, 2)
        TxtAcName(K).TEXT = FGrid1.TextMatrix(ScrolIndex + K, 3)
        TxtCr(K).TEXT = FGrid1.TextMatrix(ScrolIndex + K, 6)
        TxtDr(K).TEXT = FGrid1.TextMatrix(ScrolIndex + K, 5)
        TxtNar(K).TEXT = FGrid1.TextMatrix(ScrolIndex + K, 8)
        If Val(FGrid1.TextMatrix(ScrolIndex + K, 9)) = 0 Then
            LblCb(K).CAPTION = ""
        Else
            LblCb(K).CAPTION = "Current Balance :" + Format(Abs(FGrid1.TextMatrix(ScrolIndex + K, 9)), "0.00") + " " + IIf(Val(FGrid1.TextMatrix(ScrolIndex + K, 9)) > 0, "Cr", "Dr")
        End If
        If FGrid1.TextMatrix(ScrolIndex + K, 7) = "Cr" Then
            TxtCr(K).Visible = True
            TxtDr(K).Visible = False
        Else
            TxtDr(K).Visible = True
            TxtCr(K).Visible = False
        End If
    End If
Next
End Sub
Private Sub RowDelete(Index As Integer)
Dim MyCheck As Byte, I As Integer
If ADDFLAG = 3 Then Exit Sub
mMoveFlag = True
If FGrid1.Rows > 0 Then
    LblDrAmt(0).CAPTION = Format(0, "0.00")
    LblCrAmt(0).CAPTION = Format(0, "0.00")
    If ScrolIndex + Index = 0 And FGrid1.Rows = 1 Then
        FGrid1.Rows = 0
        Index = 0
        ScrolIndex = 0
    Else
        FGrid1.RemoveItem ScrolIndex + Index
        If ScrolIndex > 0 Then
            ScrolIndex = ScrolIndex - 1
        Else
'            Index = Index - 1
        End If
    End If
    ScrollFiller "Next", Index
    LblDrAmt(0).CAPTION = Format(0, "0.00")
    LblCrAmt(0).CAPTION = Format(0, "0.00")
    If FGrid1.Rows > 0 Then
        For I = 0 To FGrid1.Rows - 1
            LblDrAmt(0).CAPTION = Format(Val(LblDrAmt(0).CAPTION) + Val(FGrid1.TextMatrix(I, 5)), "0.00")
            LblCrAmt(0).CAPTION = Format(Val(LblCrAmt(0).CAPTION) + Val(FGrid1.TextMatrix(I, 6)), "0.00")
        Next
    End If
    LblDrAmt(0).Refresh
    LblCrAmt(0).Refresh
    If TxtCrDr(Index).Visible = True Then
        TxtCrDr(Index).SetFocus
    Else
        If Index = 0 Then
            MakeVisible 0, True
            TxtCrDr(Index).SetFocus
        Else
            TxtCrDr(Index - 1).SetFocus
        End If
    End If
End If
mMoveFlag = False
End Sub
Private Sub FrmHotKey(ByRef KeyCode As Integer, ByRef Index As Integer)
Dim RST1CrDr As ADODB.Recordset
If Index = 0 And ADDFLAG = 3 Then Exit Sub
If Index > 0 Then
    Select Case Index
        Case 1
            KeyCode = vbKeyF6
        Case 2
            KeyCode = vbKeyF7
        Case 3
            KeyCode = vbKeyF8
        Case 4
            KeyCode = vbKeyF9
    End Select
End If
Grid_Hide
Select Case KeyCode
    Case vbKeyEscape 'Terminate Form
        TopCtrl1_eCancel
    Case vbKeyF6    'Contra
        mNCat = "CNT"
        mLastVrType = ""
        FaVrTypeSetting ""
        If Trim(TxtAcName(0)) = "" Then TxtCrDr(0) = "Cr"
    Case vbKeyF7    'Payment
        mNCat = "PMT"
        mLastVrType = ""
        FaVrTypeSetting ""
        If Trim(TxtAcName(0)) = "" Then TxtCrDr(0) = "Dr"
    Case vbKeyF8    'Receipt
        mNCat = "RCT"
        mLastVrType = ""
        FaVrTypeSetting ""
        If Trim(TxtAcName(0)) = "" Then TxtCrDr(0) = "Cr"
    Case vbKeyF9    'Journal
        mNCat = "JV"
        mLastVrType = ""
        FaVrTypeSetting ""
        If Trim(TxtAcName(0)) = "" Then TxtCrDr(0) = "Dr"
End Select
If TxtVtYpe(0).Tag <> "" Then
    Set RST1CrDr = G_FaCn.Execute("SELECT * FROM VOUCHER_TYPE WHERE V_TYPE='" & TxtVtYpe(0).Tag & "'")
    If RST1CrDr.RecordCount > 0 Then
        If Trim(FaXNull(RST1CrDr!FirstDrCr)) <> "" Then
            TxtCrDr(0) = FaXNull(RST1CrDr!FirstDrCr)
        End If
    End If
End If
Set RST1CrDr = Nothing
End Sub
Private Function LedPost(AED As String) As Byte
Dim I As Integer, mDR As Double, mCR As Double, J As Integer, mDocId As String, Begin As Byte
Dim RST1 As ADODB.Recordset, Rst2 As ADODB.Recordset, mSumCr As Double, mSumDr As Double
Dim mDrInclude As Boolean, mCrInclude As Boolean, mMULTI As Integer, mTime As String
Dim mRefNo As String, TDSVtype As String, TDSVNo As Long, mTDSDocID As String, mSanSite As String
On Error GoTo err
    Select Case PubFaSiteType
        Case 0
            mSanSite = PubSiteCode + PubSiteCode
        Case 1
            If AED = "E" Then
                Set RST1 = G_FaCn.Execute("SELECT * FROM LEDGER WHERE DocId='" & RstMain!DocID & "'")
                If RST1.RecordCount > 0 Then
                    mSanSite = left(RST1!SITE_CODE, 1) + Trim(TxtSite.BoundText)
                Else
                    mSanSite = left(PubSiteCode, 1) + Trim(TxtSite.BoundText)
                End If
            Else
                mSanSite = left(PubSiteCode, 1) + Trim(TxtSite.BoundText)
            End If
        Case 2
            mSanSite = FaSetW(TxtSite.BoundText, 2)
    End Select
    
    mDocId = PubDivCode + mSanSite + FaSetW(TxtVtYpe(0).Tag, 5) + FaSetW(LblVPrefix, 5) + FaSetN(TxtVno(0), 8)
    If Len(Trim(TxtGlb(0))) > 0 Then
        If Asc(Right(Trim(TxtGlb(0)), 1)) = 10 Then TxtGlb(0) = left(Trim(TxtGlb(0)), Len(Trim(TxtGlb(0))) - 1)
        If Asc(Right(Trim(TxtGlb(0)), 1)) = 13 Then TxtGlb(0) = left(Trim(TxtGlb(0)), Len(Trim(TxtGlb(0))) - 1)
    End If
    If AED = "A" Then If G_FaCn.Execute("SELECT COUNT(*) FROM LEDGERM WHERE DocId='" & mDocId & "'").Fields(0) > 0 Then LedPost = 2: Exit Function
    If AED = "E" Then If mDocId <> RstMain!DocID Then If G_FaCn.Execute("SELECT COUNT(*) FROM LEDGERM WHERE DocId='" & mDocId & "'").Fields(0) > 0 Then LedPost = 2: Exit Function
    Begin = 0
    mDrInclude = False
    mCrInclude = False
    LedPost = 0
    mDR = 0
    mCR = 0
    TDSVNo = 0
    TDSVtype = ""
    mRefNo = ""
    Set RST1 = G_FaCn.Execute("SELECT * FROM Voucher_Include WHERE V_TYPE='" & TxtVtYpe(0).Tag & "' ORDER BY GROUPCODE")
    If RST1.RecordCount <= 0 Then
        mDrInclude = True
        mCrInclude = True
    End If
    For I = 0 To FGrid1.Rows - 1
        If Val(FGrid1.TextMatrix(I, 5)) > 0 Then
            mDR = mDR + 1
            TDSVtype = FGrid1.TextMatrix(I, 2)
        ElseIf Val(FGrid1.TextMatrix(I, 6)) > 0 Then
            mCR = mCR + 1
            mRefNo = FGrid1.TextMatrix(I, 2)
        End If
    Next
    If mCR = 1 Or mDR = 1 Then
        For I = 0 To FGrid1.Rows - 1
            If Val(FGrid1.TextMatrix(I, 5)) > 0 Then
                If mCR = 1 Then FGrid1.TextMatrix(I, 4) = mRefNo
            ElseIf Val(FGrid1.TextMatrix(I, 6)) > 0 Then
                If mDR = 1 Then FGrid1.TextMatrix(I, 4) = TDSVtype
            End If
        Next
    End If
    TDSVtype = ""
    mRefNo = ""
    mDR = 0
    mCR = 0
    For I = 0 To FGrid1.Rows - 1
        mDR = mDR + Val(FGrid1.TextMatrix(I, 5))
        mCR = mCR + Val(FGrid1.TextMatrix(I, 6))
        If RST1.RecordCount > 0 Then
            Set Rst2 = G_FaCn.Execute("SELECT GROUPCODE FROM VIEWSUBGROUP WHERE SUBCODE=" & FaChk_Text(Trim(FGrid1.TextMatrix(I, 2))))
            If Rst2.RecordCount > 0 Then
                RST1.MoveFirst
                RST1.Find "GROUPCODE=" & FaChk_Text(Trim(Rst2!GroupCode))
                If Val(FGrid1.TextMatrix(I, 5)) > 0 Then If Not RST1.EOF Then If RST1!dr = "Y" Then mDrInclude = True
                If Val(FGrid1.TextMatrix(I, 6)) > 0 Then If Not RST1.EOF Then If RST1!cr = "Y" Then mCrInclude = True
            End If
        End If
        If FGrid1.TextMatrix(I, 4) = "" And Val(FGrid1.TextMatrix(I, 5)) > 0 Then
            For J = 0 To FGrid1.Rows - 1
                If FGrid1.TextMatrix(J, 4) = "" And Val(FGrid1.TextMatrix(J, 6)) = Val(FGrid1.TextMatrix(I, 5)) Then
                    FGrid1.TextMatrix(J, 4) = FGrid1.TextMatrix(I, 2)
                    FGrid1.TextMatrix(I, 4) = FGrid1.TextMatrix(J, 2)
                End If
            Next
        End If
    Next
    If AED = "A" Then If G_FaCn.Execute("SELECT COUNT(*) FROM LEDGERM WHERE DocId='" & mDocId & "'").Fields(0) > 0 Then LedPost = 2: Exit Function
    If AED = "E" Then If mDocId <> RstMain!DocID Then If G_FaCn.Execute("SELECT COUNT(*) FROM LEDGERM WHERE DocId='" & mDocId & "'").Fields(0) > 0 Then LedPost = 2: Exit Function
    If mDrInclude = False Then LedPost = 5: MsgBox "Entry Does not Include Required Debit Account", vbCritical, Me.CAPTION: Exit Function
    If mCrInclude = False Then LedPost = 5: MsgBox "Entry Does not Include Required Credit Account", vbCritical, Me.CAPTION: Exit Function
    If mDR = 0 Or Val(mDR) <> Val(mCR) Then LedPost = 4: MsgBox "Entry Mismatched", vbCritical, Me.CAPTION: Exit Function
    If RstTds.RecordCount > 0 Then
        Set RST1 = G_FaCn.Execute("SELECT * FROM VOUCHER_TYPE WHERE NCAT='TDS'")
        If RST1.RecordCount > 0 Then
            TDSVtype = RST1!V_tYPE
        Else
            MsgBox "There Must be a TDS Vr.Type", vbCritical, Me.CAPTION: Exit Function
        End If
    End If
    If RstTds.RecordCount > 0 Then
        Do Until RstTds.EOF
            If FaXNull(RstTds!TDSDocId) <> "" Then
                mTDSDocID = FaXNull(RstTds!TDSDocId)
                TDSVNo = Mid(FaXNull(RstTds!TDSDocId), 14, 8)
            End If
            RstTds.MoveNext
        Loop
        If mTDSDocID = "" Then
            Set RST1 = G_FaCn.Execute("SELECT Max(V_NO) AS VNO FROM LEDGER WHERE V_TYPE='" & TDSVtype & "' AND V_PREFIX='" & LblVPrefix & "'")
            If RST1.RecordCount > 0 Then
                TDSVNo = FaVNull(RST1!VNo) + 1
            Else
                TDSVNo = 1
            End If
            mTDSDocID = PubDivCode + mSanSite + FaSetW(TDSVtype, 5) + FaSetW(LblVPrefix, 5) + FaSetN(Trim(STR(TDSVNo)), 8)
        End If
    End If
    G_FaCn.BeginTrans
    Begin = 1
    If AED = "E" Or AED = "D" Then
        Set RST1 = G_FaCn.Execute("SELECT * FROM LEDGER WHERE DocId='" & RstMain!DocID & "'")
        Do Until RST1.EOF
            FaCalCurrBal G_FaCn, RST1!SubCode, IIf(RST1!AmtCr > 0, RST1!AmtCr, 0), IIf(RST1!AmtDr > 0, RST1!AmtDr, 0)
            RST1.MoveNext
        Loop
        G_FaCn.Execute "DELETE FROM LEDGER WHERE DocId='" & RstMain!DocID & "'"
        Set RST1 = G_FaCn.Execute("SELECT * FROM LEDGERTDS WHERE DOCID='" & RstMain!DocID & "'")
        If RST1.RecordCount > 0 Then
            Do Until RST1.EOF
                FaCalCurrBal G_FaCn, RST1!TDSCODE, RST1!TDSAmt, 0
                FaCalCurrBal G_FaCn, RST1!TDSDRCODE, 0, RST1!TDSAmt
                G_FaCn.Execute "DELETE FROM LEDGER WHERE DOCID='" & RST1!TDSDocId & "'"
                RST1.MoveNext
            Loop
            G_FaCn.Execute "DELETE FROM LEDGERTDS WHERE DocId='" & RstMain!DocID & "'"
        End If
    End If
    If AED = "A" Or AED = "E" Then
        J = 0
        For I = 0 To FGrid1.Rows - 1
            If Val(FGrid1.TextMatrix(I, 6)) + Val(FGrid1.TextMatrix(I, 5)) > 0 Then
                mRefNo = ""
                J = J + 1
                G_FaCn.Execute "INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,ContraSub,Narration,U_Name,U_EntDt,U_AE,Chq_No,Chq_Date,Clg_Date,AgRefNo) VALUES ('" & mDocId & "'," & Val(FGrid1.TextMatrix(I, 1)) & ",'" & TxtVtYpe(0).Tag & "'," & Val(TxtVno(0).TEXT) & "," & FaChk_Text(LblVPrefix) & ",'" & mSanSite & "'," & FaConvertDate(VchDt(0)) & ",'" & FGrid1.TextMatrix(I, 2) & "'," & Val(FGrid1.TextMatrix(I, 6)) & "," & Val(FGrid1.TextMatrix(I, 5)) & "," & FaChk_Text(FGrid1.TextMatrix(I, 4)) & "," & FaChk_Text(FGrid1.TextMatrix(I, 8)) & ",'" & pubUName & "'," & FaConvertDate(Now) & ",'" & AED & "'," & FaChk_Text(Trim(TxtCHno(0))) & "," & FaConvertDate(TXTChDate(0)) & "," & FaConvertDate(TXTClrDate(0)) & "," & FaChk_Text(mRefNo) & ")"
                If FGrid1.TextMatrix(I, 7) = "Cr" Then
                    FaCalCurrBal G_FaCn, FGrid1.TextMatrix(I, 2), 0, Val(FGrid1.TextMatrix(I, 6))
                ElseIf FGrid1.TextMatrix(I, 7) = "Dr" Then
                    FaCalCurrBal G_FaCn, FGrid1.TextMatrix(I, 2), Val(FGrid1.TextMatrix(I, 5)), 0
                End If
                If RstMainAdj.RecordCount > 0 Then
                    RstMainAdj.Sort = "VSNo ASC"
                    RstMainAdj.Find "VSNO=" & Val(FGrid1.TextMatrix(I, 1))
                    If RstMainAdj.EOF = False Then
                        Do While RstMainAdj!VSNO = Val(FGrid1.TextMatrix(I, 1))
                            If RstMainAdj!SubCode <> FGrid1.TextMatrix(I, 2) Then
                                RstMainAdj.Delete
                            End If
                            RstMainAdj.MoveNext
                            If RstMainAdj.EOF = True Then Exit Do
                        Loop
                    End If
                End If
                If RstRef.RecordCount > 0 Then
                    RstRef.Sort = "V_SNo ASC"
                    RstRef.Find "V_SNO=" & Val(FGrid1.TextMatrix(I, 1))
                    If RstRef.EOF = False Then
                        Do While RstRef!V_SNo = Val(FGrid1.TextMatrix(I, 1))
                            If RstRef!SubCode <> FGrid1.TextMatrix(I, 2) Then
                                RstRef.Delete
                            End If
                            RstRef.MoveNext
                            If RstRef.EOF = True Then Exit Do
                        Loop
                    End If
                End If
            End If
        Next
        
        If PubFaSiteType = 1 And PubSeparateVrNoForSite = 1 Then
            If AED = "A" Then
                G_FaCn.Execute "INSERT INTO LedgerM (DocId,V_Type,v_Prefix,V_No,Site_Code,V_Date,Narration,U_Name,U_EntDt,U_AE) VALUES ('" & mDocId & "','" & TxtVtYpe(0).Tag & "'," & FaChk_Text(LblVPrefix) & "," & Val(TxtVno(0).TEXT) & ",'" & mSanSite & "'," & FaConvertDate(VchDt(0)) & "," & FaChk_Text(Replace(TxtGlb(0), vbCrLf, "")) & ",'" & pubUName & "','" & Format(Now, "dd/MMM/yyyy") & "'," & FaChk_Text(AED) & ")"
                G_FaCn.Execute "UPDATE VOUCHER_Prefix SET Start_Srl_No=" & Val(TxtVno(0).TEXT) & " WHERE V_Type='" & TxtVtYpe(0).Tag & "' and Prefix='" & LblVPrefix & "' AND SITE_cODE='" & Trim(TxtSite.BoundText) & "'"
                mxLastVrType = TxtVtYpe(0).Tag
                mLastPrefix = LblVPrefix
            Else
                G_FaCn.Execute "UPDATE LedgerM SET DocId='" & mDocId & "',V_Type='" & TxtVtYpe(0).Tag & "',v_Prefix=" & FaChk_Text(LblVPrefix) & ",V_No=" & Val(TxtVno(0).TEXT) & ",Site_Code='" & mSanSite & "',V_Date=" & FaConvertDate(VchDt(0)) & ",Narration=" & FaChk_Text(Replace(TxtGlb(0), vbCrLf, "")) & ",U_Name='" & pubUName & "',U_EntDt='" & Format(Now, "dd/MMM/yyyy") & "',U_AE='" & AED & "' WHERE DocId='" & RstMain!DocID & "'"
            End If
            mTime = Format(VchDt(0), "dd/MMM/yyyy") + " " + Format(Time, "hh:nn:ss")
            
            If G_FaCn.Execute("SELECT COUNT(*) FROM LastVoucher WHERE User_Name='" & pubUName & "' AND V_Type='" & TxtVtYpe(0).Tag & "'").Fields(0).Value > 0 Then
                G_FaCn.Execute "UPDATE LastVoucher Set DOCID='" & mDocId & "',Last_Ent_Date=" & FaConvertDateTime(mTime) & ",V_Type='" & TxtVtYpe(0).Tag & "',U_EntDt=" & FaConvertDateTime(Now) & ",U_NAME='" & pubUName & "' where User_Name='" & pubUName & "' AND V_Type='" & TxtVtYpe(0).Tag & "'"
            Else
                G_FaCn.Execute "INSERT INTO LastVoucher (User_Name,V_Type,Last_Ent_Date,DOCID,U_Name,U_EntDt) Values ('" & pubUName & "','" & TxtVtYpe(0).Tag & "'," & FaConvertDateTime(mTime) & ",'" & mDocId & "','" & pubUName & "'," & FaConvertDateTime(Now) & ")"
            End If
        ElseIf PubFaSiteType = 2 Then
            If AED = "A" Then
                G_FaCn.Execute "INSERT INTO LedgerM (DocId,V_Type,v_Prefix,V_No,Site_Code,V_Date,Narration,U_Name,U_EntDt,U_AE) VALUES ('" & mDocId & "','" & TxtVtYpe(0).Tag & "'," & FaChk_Text(LblVPrefix) & "," & Val(TxtVno(0).TEXT) & ",'" & mSanSite & "'," & FaConvertDate(VchDt(0)) & "," & FaChk_Text(Replace(TxtGlb(0), vbCrLf, "")) & ",'" & pubUName & "','" & Format(Now, "dd/MMM/yyyy") & "'," & FaChk_Text(AED) & ")"
                G_FaCn.Execute "UPDATE VOUCHER_Prefix SET Start_Srl_No=" & Val(TxtVno(0).TEXT) & " WHERE V_Type='" & TxtVtYpe(0).Tag & "' and Prefix='" & LblVPrefix & "' AND SITE_cODE='" & mSanSite & "'"
                mxLastVrType = TxtVtYpe(0).Tag
                mLastPrefix = LblVPrefix
            Else
                G_FaCn.Execute "UPDATE LedgerM SET DocId='" & mDocId & "',V_Type='" & TxtVtYpe(0).Tag & "',v_Prefix=" & FaChk_Text(LblVPrefix) & ",V_No=" & Val(TxtVno(0).TEXT) & ",Site_Code='" & mSanSite & "',V_Date=" & FaConvertDate(VchDt(0)) & ",Narration=" & FaChk_Text(Replace(TxtGlb(0), vbCrLf, "")) & ",U_Name='" & pubUName & "',U_EntDt='" & Format(Now, "dd/MMM/yyyy") & "',U_AE='" & AED & "' WHERE DocId='" & RstMain!DocID & "'"
            End If
            mTime = Format(VchDt(0), "dd/MMM/yyyy") + " " + Format(Time, "hh:nn:ss")
            If G_FaCn.Execute("SELECT COUNT(*) FROM LastVoucher WHERE User_Name='" & pubUName & "' AND V_Type='" & TxtVtYpe(0).Tag & "' AND SITE_cODE='" & mSanSite & "'").Fields(0).Value > 0 Then
                G_FaCn.Execute "UPDATE LastVoucher Set DOCID='" & mDocId & "',Last_Ent_Date=" & FaConvertDateTime(mTime) & ",V_Type='" & TxtVtYpe(0).Tag & "',U_EntDt=" & FaConvertDateTime(Now) & ",U_NAME='" & pubUName & "' where User_Name='" & pubUName & "' AND V_Type='" & TxtVtYpe(0).Tag & "' AND SITE_cODE='" & mSanSite & "'"
            Else
                G_FaCn.Execute "INSERT INTO LastVoucher (User_Name,V_Type,Last_Ent_Date,DOCID,U_Name,U_EntDt,SITE_CODE) Values ('" & pubUName & "','" & TxtVtYpe(0).Tag & "'," & FaConvertDateTime(mTime) & ",'" & mDocId & "','" & pubUName & "'," & FaConvertDateTime(Now) & ",'" & PubSiteCode & "')"
            End If
        Else
            If AED = "A" Then
                G_FaCn.Execute "INSERT INTO LedgerM (DocId,V_Type,v_Prefix,V_No,Site_Code,V_Date,Narration,U_Name,U_EntDt,U_AE) VALUES ('" & mDocId & "','" & TxtVtYpe(0).Tag & "'," & FaChk_Text(LblVPrefix) & "," & Val(TxtVno(0).TEXT) & ",'" & mSanSite & "'," & FaConvertDate(VchDt(0)) & "," & FaChk_Text(Replace(TxtGlb(0), vbCrLf, "")) & ",'" & pubUName & "','" & Format(Now, "dd/MMM/yyyy") & "'," & FaChk_Text(AED) & ")"
                G_FaCn.Execute "UPDATE VOUCHER_Prefix SET Start_Srl_No=" & Val(TxtVno(0).TEXT) & " WHERE V_Type='" & TxtVtYpe(0).Tag & "' and Prefix='" & LblVPrefix & "'"
                mxLastVrType = TxtVtYpe(0).Tag
                mLastPrefix = LblVPrefix
            Else
                G_FaCn.Execute "UPDATE LedgerM SET DocId='" & mDocId & "',V_Type='" & TxtVtYpe(0).Tag & "',v_Prefix=" & FaChk_Text(LblVPrefix) & ",V_No=" & Val(TxtVno(0).TEXT) & ",Site_Code='" & mSanSite & "',V_Date=" & FaConvertDate(VchDt(0)) & ",Narration=" & FaChk_Text(Replace(TxtGlb(0), vbCrLf, "")) & ",U_Name='" & pubUName & "',U_EntDt='" & Format(Now, "dd/MMM/yyyy") & "',U_AE='" & AED & "' WHERE DocId='" & RstMain!DocID & "'"
            End If
            mTime = Format(VchDt(0), "dd/MMM/yyyy") + " " + Format(Time, "hh:nn:ss")
            If G_FaCn.Execute("SELECT COUNT(*) FROM LastVoucher WHERE User_Name='" & pubUName & "' AND V_Type='" & TxtVtYpe(0).Tag & "'").Fields(0).Value > 0 Then
                G_FaCn.Execute "UPDATE LastVoucher Set DOCID='" & mDocId & "',Last_Ent_Date=" & FaConvertDateTime(mTime) & ",V_Type='" & TxtVtYpe(0).Tag & "',U_EntDt=" & FaConvertDateTime(Now) & ",U_NAME='" & pubUName & "' where User_Name='" & pubUName & "' AND V_Type='" & TxtVtYpe(0).Tag & "'"
            Else
                G_FaCn.Execute "INSERT INTO LastVoucher (User_Name,V_Type,Last_Ent_Date,DOCID,U_Name,U_EntDt) Values ('" & pubUName & "','" & TxtVtYpe(0).Tag & "'," & FaConvertDateTime(mTime) & ",'" & mDocId & "','" & pubUName & "'," & FaConvertDateTime(Now) & ")"
            End If
        End If
        
        G_FaCn.Execute "DELETE FROM LEDGERAdj WHERE DocId1='" & mDocId & "' OR DocId2='" & mDocId & "'"
        If RstMainAdj.RecordCount > 0 Then
            RstMainAdj.MoveFirst
            Do Until RstMainAdj.EOF
                If FaVNull(RstMainAdj!cr) > 0 Then
                    G_FaCn.Execute ("INSERT INTO LEDGERADJ (DOCID1,V_SNO1,DOCID2,V_SNO2,CR,SUBCODE,U_Name,U_EntDt,U_AE) VALUES ('" & RstMainAdj!DocID & "'," & FaVNull(RstMainAdj!V_SNo) & ",'" & mDocId & "'," & FaVNull(RstMainAdj!VSNO) & "," & FaVNull(RstMainAdj!cr) & ",'" & RstMainAdj!SubCode & "','" & pubUName & "'," & FaConvertDate(Now) & "," & FaChk_Text(AED) & ")")
                Else
                    G_FaCn.Execute ("INSERT INTO LEDGERADJ (DOCID1,V_SNO1,DOCID2,V_SNO2,CR,SUBCODE,U_Name,U_EntDt,U_AE) VALUES ('" & mDocId & "'," & FaVNull(RstMainAdj!VSNO) & ",'" & RstMainAdj!DocID & "'," & FaVNull(RstMainAdj!V_SNo) & "," & FaVNull(RstMainAdj!dr) & ",'" & RstMainAdj!SubCode & "','" & pubUName & "'," & FaConvertDate(Now) & "," & FaChk_Text(AED) & ")")
                End If
                RstMainAdj.MoveNext
            Loop
        End If
        G_FaCn.Execute "DELETE FROM LedgerRef WHERE DocId='" & mDocId & "'"
        If RstRef.RecordCount > 0 Then
            J = 0
            Set RST1 = G_FaCn.Execute("SELECT MAX(ID) AS MaxId FROM LedgerRef")
            If RST1.RecordCount > 0 Then
                J = FaVNull(RST1!MaxId)
            End If
            RstRef.MoveFirst
            Do Until RstRef.EOF
                J = J + 1
                G_FaCn.Execute ("INSERT INTO LedgerRef (ID,DOCID,V_SNO,DR,CR,SUBCODE,U_Name,U_EntDt,U_AE,AgRefType,AgRefNo,DueDate,V_dATE) VALUES (" & J & ",'" & mDocId & "'," & FaVNull(RstRef!V_SNo) & "," & FaVNull(RstRef!dr) & "," & FaVNull(RstRef!cr) & ",'" & RstRef!SubCode & "','" & pubUName & "'," & FaConvertDate(Now) & "," & FaChk_Text(AED) & "," & FaChk_Text(RstRef!AgRefType) & "," & FaChk_Text(RstRef!AgRefNo) & "," & FaConvertDate(RstRef!DUEDATE) & "," & FaConvertDate(VchDt(0)) & ")")
                RstRef.MoveNext
            Loop
        End If
        J = 0
        If RstTds.RecordCount > 0 Then
            RstTds.MoveFirst
            Do Until RstTds.EOF
                J = J + 1
                G_FaCn.Execute "INSERT INTO LEDGERTDS (DocId,V_SNo,TDSCode,TDSDrCode,TDSYN,ONAMT,TDS,TDSAMT,TDSPOST,TDSDocId,TDSV_SNo,V_dATE,v_Prefix,Site_Code,U_Name,U_EntDt,U_AE) VALUES ('" & mDocId & "'," & RstTds!V_SNo & ",'" & RstTds!TDSCODE & "','" & RstTds!TDSDRCODE & "','" & RstTds!TDSYN & "'," & RstTds!ONAmt & "," & RstTds!TDS & "," & RstTds!TDSAmt & ",'" & RstTds!TDSPOST & "','" & mTDSDocID & "'," & J & "," & FaConvertDate(VchDt(0)) & "," & FaChk_Text(LblVPrefix) & ",'" & mSanSite & "','" & pubUName & "'," & FaConvertDate(Now) & ",'" & AED & "')"
                G_FaCn.Execute "INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,ContraSub,Narration,U_Name,U_EntDt,U_AE) VALUES ('" & mTDSDocID & "'," & J & ",'" & TDSVtype & "'," & TDSVNo & "," & FaChk_Text(LblVPrefix) & ",'" & mSanSite & "'," & FaConvertDate(VchDt(0)) & ",'" & RstTds!TDSCODE & "'," & RstTds!TDSAmt & ",0,'" & RstTds!TDSDRCODE & "'," & FaChk_Text(Replace(RstTds!Narration, vbCrLf, "")) & ",'" & pubUName & "'," & FaConvertDate(Now) & ",'" & AED & "')"
                J = J + 1
                G_FaCn.Execute "INSERT INTO LEDGER (DocId,V_SNo,V_Type,V_No,v_Prefix,Site_Code,V_Date,SubCode,AmtCr,AmtDr,ContraSub,Narration,U_Name,U_EntDt,U_AE) VALUES ('" & mTDSDocID & "'," & J & ",'" & TDSVtype & "'," & TDSVNo & "," & FaChk_Text(LblVPrefix) & ",'" & mSanSite & "'," & FaConvertDate(VchDt(0)) & ",'" & RstTds!TDSDRCODE & "',0," & RstTds!TDSAmt & ",'" & RstTds!TDSCODE & "'," & FaChk_Text(Replace(RstTds!Narration, vbCrLf, "")) & ",'" & pubUName & "'," & FaConvertDate(Now) & ",'" & AED & "')"
                FaCalCurrBal G_FaCn, RstTds!TDSCODE, 0, RstTds!TDSAmt
                FaCalCurrBal G_FaCn, RstTds!TDSDRCODE, RstTds!TDSAmt, 0
                RstTds.MoveNext
            Loop
        End If
        mLastVrType = TxtVtYpe(0).Tag
        G_FaCn.CommitTrans
        Begin = 0
        RstMain.Requery
        RstMain.Find "DocId='" & mDocId & "'"
        If ADDFLAG = 2 Then
            MoveRec
        End If
    End If
    Set RST1 = Nothing
    Set Rst2 = Nothing
    Exit Function
err:    If Begin = 1 Then G_FaCn.RollbackTrans
        MsgBox err.Description, vbCritical, Me.CAPTION
        LedPost = 1
End Function
Private Sub BTNCLOSE_Click()
Frame1(1).Visible = False
End Sub
Private Sub btnPrint1_Click()
    If ChkReport = 1 Then
        Set rpt = rdApp.OpenReport(PubFaReportPath + "\FaJVCHR.RPT")
        FaRectPrintingModule Me, rpt
    Else
        Set rpt = rdApp.OpenReport(PubFaReportPath + "\FaJVCHR.RPT")
        FaVoucherPrintingModule Me, rpt
    End If
    Set rpt = Nothing
End Sub
Private Sub BTNVLCLOSE_Click()
    FRAMEVLIST.Visible = False
End Sub
Private Sub DataCombo3_Validate(Cancel As Boolean)
If DataCombo3.TEXT = "" Then
    MsgBox "Select Voucher Type", vbInformation, Me.CAPTION
    DataCombo3.SetFocus
Else
    FaIniCombo "SELECT distinct v_no FROM LEDGER where v_type='" & DataCombo3.BoundText & "' ORDER BY v_no", DataCombo2(0), "v_no", "v_no"
    FaIniCombo "SELECT distinct v_no FROM LEDGER where v_type='" & DataCombo3.BoundText & "' ORDER BY v_no", DataCombo2(1), "v_no", "v_no"
End If
End Sub
Private Sub Opt2_Click(Index As Integer)
    Frame1(2).Enabled = IIf(Index = 0 Or Index = 2, False, True)
    Frame1(3).Visible = IIf(Index = 0 Or Index = 1, False, True)
    ChkReport.Visible = IIf(Index = 0, True, False)
    If Index = 2 Then
        Frame1(3).ZOrder 0
        Text2.Visible = True
        Text2.Enabled = True
        Text2.SetFocus
    Else
        Frame1(2).ZOrder 0
    End If
End Sub
Private Sub TXTClrDate_GotFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_GetFocus TXTClrDate(Index)
    Grid_Hide
End Sub
Private Sub TXTClrDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mMoveFlag = True Then Exit Sub
    If KeyCode = vbKeyEscape Then Grid_Hide
End Sub
Private Sub TXTClrDate_LostFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate TXTClrDate(Index)
End Sub
Private Sub TXTClrDate_Validate(Index As Integer, Cancel As Boolean)
    If mMoveFlag = True Then Exit Sub
    If ADDFLAG = 3 Then Exit Sub
    TXTClrDate(0) = PubDatamanFa.FaRetDateFunc(TXTClrDate(0))
End Sub
Private Sub TXTChDate_GotFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_GetFocus TXTChDate(Index)
    Grid_Hide
End Sub
Private Sub TXTChDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mMoveFlag = True Then Exit Sub
    If KeyCode = vbKeyEscape Then Grid_Hide
End Sub
Private Sub TXTChDate_LostFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate TXTChDate(Index)
End Sub
Private Sub TXTChDate_Validate(Index As Integer, Cancel As Boolean)
    If mMoveFlag = True Then Exit Sub
    If ADDFLAG = 3 Then Exit Sub
    TXTChDate(Index) = PubDatamanFa.FaRetDateFunc(TXTChDate(Index))
End Sub
Private Sub TxtCHno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mMoveFlag = True Then Exit Sub
    If KeyCode = vbKeyEscape Then Grid_Hide
End Sub
Private Sub TxtCHno_GotFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_GetFocus TxtCHno(Index)
    Grid_Hide
End Sub
Private Sub TxtCHno_LostFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate TxtCHno(Index)
    If Trim(TxtCHno(Index)) <> "" Then
        If ADDFLAG = 1 Then
            If G_FaCn.Execute("SELECT COUNT(*) FROM LEDGER WHERE Chq_No='" & TxtCHno(Index) & "'").Fields(0).Value > 0 Then
                MsgBox "Duplicate Cheque No.", vbInformation + vbDefaultButton2, "Cheque No.Validation"
            End If
        Else
            If G_FaCn.Execute("SELECT COUNT(*) FROM LEDGER WHERE Chq_No='" & TxtCHno(Index) & "' AND DOCID<>'" & RstMain!DocID & "'").Fields(0).Value > 0 Then
                MsgBox "Duplicate Cheque No.", vbInformation + vbDefaultButton2, "Cheque No.Validation"
            End If
        End If
    End If
End Sub
Private Sub TxtCr_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Val(TxtCr(Index)) <> 0 Then
        LblAmtRs = FaNToW(Val(TxtCr(Index)), "Rs", " Paise")
    Else
        LblAmtRs = ""
    End If
End Sub
Private Sub TxtCrDr_Validate(Index As Integer, Cancel As Boolean)
Dim RstAcType As ADODB.Recordset
    If mMoveFlag = True Then Exit Sub
    If ADDFLAG = 3 Then Exit Sub
    If ADDFLAG <= 2 Then If Trim(TxtCrDr(Index)) <> "Cr" And Trim(TxtCrDr(Index)) <> "Dr" And TxtCrDr(Index).Visible = True Then Cancel = True: Exit Sub
    If ADDFLAG <= 2 Then
        If Trim(TxtAcName(Index)) = "" Then
            If Trim(TxtCrDr(Index)) = "Cr" Then
                Set RstAcType = G_FaCn.Execute("SELECT SubGroup.Name,VOUCHER_TYPE.* FROM VOUCHER_TYPE LEFT JOIN SubGroup ON VOUCHER_TYPE.DefaultCrAC = SubGroup.SubCode WHERE VOUCHER_TYPE.V_TYPE='" & TxtVtYpe(0).Tag & "'")
                If RstAcType.RecordCount > 0 Then
                    TxtAcName(Index).Tag = FaXNull(RstAcType!DefaultCrAC)
                    TxtAcName(Index) = FaXNull(RstAcType!Name)
                End If
            End If
            If Trim(TxtCrDr(Index)) = "Dr" Then
                Set RstAcType = G_FaCn.Execute("SELECT SubGroup.Name,VOUCHER_TYPE.* FROM VOUCHER_TYPE LEFT JOIN SubGroup ON VOUCHER_TYPE.DefaultDrAC = SubGroup.SubCode WHERE VOUCHER_TYPE.V_TYPE='" & TxtVtYpe(0).Tag & "'")
                If RstAcType.RecordCount > 0 Then
                    TxtAcName(Index).Tag = FaXNull(RstAcType!DefaultDrAC)
                    TxtAcName(Index) = FaXNull(RstAcType!Name)
                End If
            End If
        End If
    End If
Set RstAcType = Nothing
End Sub
Private Sub TxtDr_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Val(TxtDr(Index)) <> 0 Then
        LblAmtRs = FaNToW(Val(TxtDr(Index)), "Rs", " Paise")
    Else
        LblAmtRs = ""
    End If
End Sub
Private Sub TxtGlb_GotFocus(Index As Integer)
Dim RST1 As ADODB.Recordset
    If mMoveFlag = True Then Exit Sub
    Ctrl_GetFocus TxtGlb(Index)
    Grid_Hide
    If ADDFLAG <= 2 And mCommNar = "Y" And Len(TxtGlb(Index)) = 0 Then
        Set RST1 = G_FaCn.Execute("SELECT Narration FROM VOUCHER_tYPE WHERE V_tYPE=" & FaChk_Text(TxtVtYpe(0).Tag))
        If RST1.RecordCount > 0 Then
            TxtGlb(0) = FaXNull(RST1!Narration)
        End If
    End If
Set RST1 = Nothing
End Sub
Private Sub TxtGlb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mMoveFlag = True Then Exit Sub
    If KeyCode = vbKeyEscape Then Grid_Hide
    If ADDFLAG <= 2 Then If KeyCode = vbKeyInsert Then TxtGlb(Index) = TxtGlb(Index) + FaNarr: TxtGlb(Index).SelStart = Len(TxtGlb(Index)): Exit Sub
End Sub
Private Sub TxtGlb_LostFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate TxtGlb(Index)
End Sub
Private Sub TxtCrDr_GotFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_GetFocus TxtCrDr(Index)
    Grid_Hide
End Sub
Private Sub TxtCrDr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mMoveFlag = True Then Exit Sub
    If KeyCode = vbKeyEscape Then Grid_Hide
End Sub
Private Sub TxtCrDr_KeyPress(Index As Integer, KeyAscii As Integer)
Dim MyCheck As Byte
    If mMoveFlag = True Then Exit Sub
    If ADDFLAG = 3 Then Exit Sub
    If KeyAscii = 68 Or KeyAscii = 100 Or KeyAscii = 67 Or KeyAscii = 99 Then
        If KeyAscii = 68 Or KeyAscii = 100 Then ' D/d
            If TxtCrDr(Index) = "Cr" Then
                TxtDr(Index) = TxtCr(Index)
                TxtDr(Index).Visible = True
                TxtCr(Index) = ""
                If Val(TxtDr(Index)) > 0 Then TxtDr(Index).SetFocus
            End If
            TxtCrDr(Index) = "Dr"
            TxtCrDr_LostFocus Index
            KeyAscii = 0
        ElseIf KeyAscii = 67 Or KeyAscii = 99 Then ' C/c
            If TxtCrDr(Index) = "Dr" Then
                TxtCr(Index) = TxtDr(Index)
                TxtCr(Index).Visible = True
                TxtDr(Index) = ""
                If Val(TxtCr(Index)) > 0 Then TxtCr(Index).SetFocus
            End If
            TxtCrDr(Index) = "Cr"
            TxtCrDr_LostFocus Index
            KeyAscii = 0
        End If
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub TxtCrDr_LostFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate TxtCrDr(Index)
    If ADDFLAG = 3 Then Exit Sub
    If TxtCrDr(Index).TEXT = "Cr" Then
        TxtDr(Index) = ""
        TxtDr(Index).Visible = False
        TxtCr(Index).Visible = True
    Else
        TxtCr(Index) = ""
        TxtDr(Index).Visible = True
        TxtCr(Index).Visible = False
    End If
End Sub
Private Sub TxtCr_GotFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_GetFocus TxtCr(Index)
    Grid_Hide
    If TxtCr(Index) = "" And TxtAcName(Index) <> "" Then If Val(LblDrAmt(0).CAPTION) > Val(LblCrAmt(0).CAPTION) Then TxtCr(Index) = Format(Val(LblDrAmt(0).CAPTION) - Val(LblCrAmt(0).CAPTION), "0.00")
    TxtCr(Index).SelLength = Len(TxtCr(Index).TEXT)
    If ADDFLAG <= 2 Then
        LblHelp.Visible = True
        LblHelp.left = 0
        LblHelp.top = TxtCHno(0).top + TxtCHno(0).height + 25
    End If
End Sub
Private Sub TxtCr_KeyPress(Index As Integer, KeyAscii As Integer)
    If mMoveFlag = True Then Exit Sub
    If ADDFLAG <= 2 Then FaNumPress TxtCr(Index), KeyAscii, 9, 2
End Sub
Private Sub TxtCr_LostFocus(Index As Integer)
    LblAmtRs = ""
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate TxtCr(Index)
End Sub
Private Sub TxtCr_Validate(Index As Integer, Cancel As Boolean)
Dim MyCheck As Byte, mAdj As Double
    If mMoveFlag = True Then Exit Sub
    If ADDFLAG = 3 Then Exit Sub
    If ADDFLAG <= 2 Then
        If Val(TxtCr(Index)) = 0 And TxtCr(Index).Visible = True Then Cancel = True: Exit Sub  'TxtCr(Index).SetFocus: Exit Sub
        TxtCr(Index).TEXT = Format(FaValidate_Numeric(TxtCr(Index)), "0.00")
        MyCheck = VchTrnSaveLst(Index)
        If RstEnviro!OnLineAdjustment = "Yes" Then
            mAdj = 0
            If RstMainAdj.RecordCount > 0 Then
                RstMainAdj.Sort = "VSNo ASC"
                RstMainAdj.Find "VSNO=" & Val(FGrid1.TextMatrix(Index + ScrolIndex, 1))
                If RstMainAdj.EOF = False Then
                    Do While RstMainAdj!VSNO = Val(FGrid1.TextMatrix(Index + ScrolIndex, 1))
                        If RstMainAdj!SubCode <> FGrid1.TextMatrix(Index + ScrolIndex, 2) Then
                            RstMainAdj.Delete
                        Else
                            mAdj = mAdj + FaVNull(RstMainAdj!dr)
                        End If
                        RstMainAdj.MoveNext
                        If RstMainAdj.EOF = True Then Exit Do
                    Loop
                End If
            End If
            If mAdj > Val(TxtCr(Index)) Then
                MsgBox "Can't Adjust More Amount"
                Cancel = True
            End If
        End If
    End If
    LblHelp.Visible = False
    LblAmtRs = ""
End Sub
Private Sub TxtDr_GotFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_GetFocus TxtDr(Index)
    Grid_Hide
    If TxtDr(Index) = "" And TxtAcName(Index) <> "" Then If Val(LblCrAmt(0).CAPTION) > Val(LblDrAmt(0).CAPTION) Then TxtDr(Index).TEXT = Format(Val(LblCrAmt(0).CAPTION) - Val(LblDrAmt(0).CAPTION), "0.00")
    TxtDr(Index).SelLength = Len(TxtDr(Index).TEXT)
    If ADDFLAG <= 2 Then
        LblHelp.Visible = True
        LblHelp.left = 0
        LblHelp.top = TxtCHno(0).top + TxtCHno(0).height + 25
    End If
End Sub
Private Sub TxtDr_KeyPress(Index As Integer, KeyAscii As Integer)
    If mMoveFlag = True Then Exit Sub
    If ADDFLAG <= 2 Then FaNumPress TxtDr(Index), KeyAscii, 9, 2
End Sub
Private Sub Txtdr_Validate(Index As Integer, Cancel As Boolean)
Dim MyCheck As Byte, mAdj As Double
    If mMoveFlag = True Then Exit Sub
    If ADDFLAG = 3 Then Exit Sub
    If ADDFLAG <= 2 Then
        If Val(TxtDr(Index)) = 0 And TxtDr(Index).Visible = True Then Cancel = True: Exit Sub  'TxtDr(Index).SetFocus: Exit Sub
        TxtDr(Index).TEXT = Format(FaValidate_Numeric(TxtDr(Index)), "0.00")
        MyCheck = VchTrnSaveLst(Index)
        If RstEnviro!OnLineAdjustment = "Yes" Then
            mAdj = 0
            If RstMainAdj.RecordCount > 0 Then
                RstMainAdj.Sort = "VSNo ASC"
                RstMainAdj.Find "VSNO=" & Val(FGrid1.TextMatrix(Index + ScrolIndex, 1))
                If RstMainAdj.EOF = False Then
                    Do While RstMainAdj!VSNO = Val(FGrid1.TextMatrix(Index + ScrolIndex, 1))
                        If RstMainAdj!SubCode <> FGrid1.TextMatrix(Index + ScrolIndex, 2) Then
                            RstMainAdj.Delete
                        Else
                            mAdj = mAdj + FaVNull(RstMainAdj!cr)
                        End If
                        RstMainAdj.MoveNext
                        If RstMainAdj.EOF = True Then Exit Do
                    Loop
                End If
            End If
            If mAdj > Val(TxtDr(Index)) Then
                MsgBox "CanX't Adjust More Amount"
                Cancel = True
            End If
        End If
    End If
    LblHelp.Visible = False
    LblAmtRs = ""
End Sub
Private Sub Txtdr_LostFocus(Index As Integer)
    LblAmtRs = ""
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate TxtDr(Index)
End Sub
Private Sub TxtGrid_LostFocus(Index As Integer)
'    TxtGrid_Validate Index, False
End Sub
Private Sub TxtNar_GotFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_GetFocus TxtNar(Index)
    Grid_Hide
End Sub
Private Sub TxtNar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mMoveFlag = True Then Exit Sub
    If KeyCode = vbKeyEscape Then Grid_Hide
    If ADDFLAG <= 2 Then If KeyCode = vbKeyInsert Then TxtNar(Index) = TxtNar(Index) + FaNarr: TxtNar(Index).SelStart = Len(TxtNar(Index)): Exit Sub
End Sub
Private Sub TxtNar_LostFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate TxtNar(Index)
    If ADDFLAG = 3 Then Exit Sub
    TxtNar_Validate Index, False
End Sub
Private Sub TxtNar_Validate(Index As Integer, Cancel As Boolean)
Dim MyCheck As Byte
    If mMoveFlag = True Then Exit Sub
    If ADDFLAG = 3 Then Exit Sub
    If ADDFLAG <= 2 Then MyCheck = VchTrnSaveLst(Index)
End Sub
Private Sub TxtONAMT_KeyPress(KeyAscii As Integer)
    If ADDFLAG <= 2 Then FaNumPress TxtONAMT, KeyAscii, 9, 2
End Sub
Private Sub TxtONAMT_KeyUp(KeyCode As Integer, Shift As Integer)
    If ADDFLAG <= 2 Then TxtTDSAMT = (Val(TxtONAMT) * Val(TxtTDS)) / 100
End Sub
Private Sub TxtTDS_KeyPress(KeyAscii As Integer)
    If ADDFLAG <= 2 Then FaNumPress TxtTDS, KeyAscii, 3, 4
End Sub
Private Sub TxtTDS_KeyUp(KeyCode As Integer, Shift As Integer)
    If ADDFLAG <= 2 Then TxtTDSAMT = (Val(TxtONAMT) * Val(TxtTDS)) / 100
End Sub
Private Sub TxtTDSAMT_KeyPress(KeyAscii As Integer)
    If ADDFLAG <= 2 Then FaNumPress TxtTDSAMT, KeyAscii, 9, 2
End Sub
Private Sub TXTVDATE1_GotFocus()
    Ctrl_GetFocus TXTVDATE1
    Grid_Hide
End Sub
Private Sub TXTVDATE1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Grid_Hide
End Sub
Private Sub TXTVDATE1_LOSTFOCUS()
    TXTVDATE1 = PubDatamanFa.FaRetDateFunc(TXTVDATE1)
    Ctrl_validate TXTVDATE1
End Sub
Private Sub TXTVDATE1_Validate(Cancel As Boolean)
    TXTVDATE1 = PubDatamanFa.FaRetDateFunc(TXTVDATE1)
End Sub
Private Sub TXTVDATE2_GotFocus()
    Ctrl_GetFocus TXTVDATE2
    Grid_Hide
End Sub
Private Sub TXTVDATE2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Grid_Hide
End Sub
Private Sub TXTVDATE2_LOSTFOCUS()
    TXTVDATE2 = PubDatamanFa.FaRetDateFunc(TXTVDATE2)
    Ctrl_validate TXTVDATE2
End Sub
Private Sub TXTVDATE2_Validate(Cancel As Boolean)
    TXTVDATE2 = PubDatamanFa.FaRetDateFunc(TXTVDATE2)
End Sub
Private Sub TxtVno_GotFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_GetFocus TxtVno(Index)
    Grid_Hide
End Sub
Private Sub TxtVno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mMoveFlag = True Then Exit Sub
    If KeyCode = vbKeyEscape Then Grid_Hide
End Sub
Private Sub TxtVno_KeyPress(Index As Integer, KeyAscii As Integer)
    If mMoveFlag = True Then Exit Sub
    FaNumPress TxtVno(Index), KeyAscii, 8, 0
End Sub
Private Sub TxtVno_LostFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate TxtVno(Index)
End Sub
Private Sub VchDt_GotFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_GetFocus VchDt(0)
    Grid_Hide
    VchDt(0).SelStart = 0
    VchDt(0).SelLength = Len(VchDt(0).TEXT)
End Sub
Private Sub VchDt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mMoveFlag = True Then Exit Sub
    If KeyCode = vbKeyEscape Then Grid_Hide
End Sub
Private Sub VchDt_LostFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate VchDt(0)
End Sub
Private Sub VchDt_Validate(Index As Integer, Cancel As Boolean)
Dim RST1 As ADODB.Recordset
    If mMoveFlag = True Then Exit Sub
    If ADDFLAG = 3 Then Exit Sub
    VchDt(0) = PubDatamanFa.FaRetDateFunc(VchDt(0))
    LblDay = PubDatamanFa.FaRetDayFunc(VchDt(0))
    If RstEnviro!DateLock = "Y" Then
        Set RST1 = G_FaCn.Execute("SELECT EDATE,SDATE FROM DateLock WHERE CODE='" & Me.Name & "'")
        If VchDt(0) = "" Then MsgBox "Date is Required": Cancel = True: Exit Sub
        If VchDt(0) <> "" Then If RST1.EOF <> True Or RST1.BOF <> True Then If CDate(VchDt(0)) > RST1!EDate Or CDate(VchDt(0)) < RST1!Sdate Then MsgBox "Date Not Permitted", vbCritical: Cancel = True: Exit Sub
    Else
        If VchDt(0) = "" Then VchDt(0) = PubLoginDate
        VchDt(0) = PubDatamanFa.FaRetDateFunc(VchDt(0))
        LblDay = PubDatamanFa.FaRetDayFunc(VchDt(0))
    End If
    FaVrTypeSetting mLastVrType, mLastPrefix
    Set RST1 = Nothing
End Sub
Private Sub TxtAcName_GotFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_GetFocus TxtAcName(Index)
    Grid_Hide
    If ADDFLAG <= 2 Then
        If TxtCrDr(Index) = "Cr" Then
            If RstAcHlpCr.RecordCount > 0 Then
                Set DGAcHlp.DataSource = Nothing
                RstAcHlpCr.MoveFirst
                Set DGAcHlp.DataSource = RstAcHlpCr
            Else
                Set DGAcHlp.DataSource = Nothing
            End If
            If RstAcHlpCr.RecordCount = 0 Or TxtAcName(Index).TEXT = "" Then Exit Sub
        Else
            If RstAcHlpDr.RecordCount > 0 Then
                Set DGAcHlp.DataSource = Nothing
                RstAcHlpDr.MoveFirst
                Set DGAcHlp.DataSource = RstAcHlpDr
            Else
                Set DGAcHlp.DataSource = Nothing
            End If
            If RstAcHlpDr.RecordCount = 0 Or TxtAcName(Index).TEXT = "" Then Exit Sub
        End If
        DGAcHlp.Tag = Index
        If TxtAcName(Index).TEXT <> "" Then
            If TxtCrDr(Index) = "Cr" Then
                RstAcHlpCr.MoveFirst
                RstAcHlpCr.Find "Name='" & TxtAcName(Index).TEXT & "'"
            Else
                RstAcHlpDr.MoveFirst
                RstAcHlpDr.Find "Name='" & TxtAcName(Index).TEXT & "'"
            End If
            DGAcHlp.ReBind
        Else
            If TxtCrDr(Index) = "Cr" Then
                RstAcHlpCr.MoveFirst
            Else
                RstAcHlpDr.MoveFirst
            End If
        End If
    End If
End Sub
Private Sub TxtAcName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mMoveFlag = True Then Exit Sub
    If KeyCode = vbKeyEscape Then Grid_Hide
    If TxtCrDr(Index) = "Cr" Then
        If ADDFLAG <= 2 Then
            FaDGridTxtKeyDown DGAcHlp, TxtAcName, Index, RstAcHlpCr, KeyCode, False, 1
            DGAcHlp.Tag = Index
            If RstEnviro!AddressHelp = "Yes" And DGAcHlp.Visible = True Then
                TxtDetailS.Visible = True
                TxtDetailS.left = DGAcHlp.left
                TxtDetailS.top = DGAcHlp.top + DGAcHlp.height
                TxtDetailS.width = DGAcHlp.width
            Else
                TxtDetailS.Visible = False
            End If
            If RstAcHlpCr.EOF = False And RstAcHlpCr.BOF = False Then
                TxtDetailS = IIf(Trim(RstAcHlpCr!FatherName) = "", "", RstAcHlpCr!FatherName + vbCrLf) + RstAcHlpCr!GroupName + vbCrLf + RstAcHlpCr!NameWithADDR
            Else
                TxtDetailS = ""
            End If
        End If
    Else
        If ADDFLAG <= 2 Then
            FaDGridTxtKeyDown DGAcHlp, TxtAcName, Index, RstAcHlpDr, KeyCode, False, 1
            DGAcHlp.Tag = Index
            If RstEnviro!AddressHelp = "Yes" And DGAcHlp.Visible = True Then
                TxtDetailS.Visible = True
                TxtDetailS.left = DGAcHlp.left
                TxtDetailS.top = DGAcHlp.top + DGAcHlp.height
                TxtDetailS.width = DGAcHlp.width
            Else
                TxtDetailS.Visible = False
            End If
            If RstAcHlpDr.EOF = False And RstAcHlpDr.BOF = False Then
                TxtDetailS = IIf(Trim(RstAcHlpDr!FatherName) = "", "", RstAcHlpDr!FatherName + vbCrLf) + RstAcHlpDr!GroupName + vbCrLf + RstAcHlpDr!NameWithADDR
            Else
                TxtDetailS = ""
            End If
        End If
    End If
End Sub
Private Sub TxtAcName_KeyPress(Index As Integer, KeyAscii As Integer)
    If mMoveFlag = True Then Exit Sub
    FaCheckQuote KeyAscii
    If TxtCrDr(Index) = "Cr" Then
        If DGAcHlp.Visible = True Then
            FaDGridTxtKeyPress TxtAcName, Index, RstAcHlpCr, KeyAscii, "Name"
            If RstEnviro!AddressHelp = "Yes" And DGAcHlp.Visible = True Then
                TxtDetailS.Visible = True
                TxtDetailS.left = DGAcHlp.left
                TxtDetailS.top = DGAcHlp.top + DGAcHlp.height
                TxtDetailS.width = DGAcHlp.width
            Else
                TxtDetailS.Visible = False
            End If
            If RstAcHlpCr.EOF = False And RstAcHlpCr.BOF = False Then
                TxtDetailS = IIf(Trim(RstAcHlpCr!FatherName) = "", "", RstAcHlpCr!FatherName + vbCrLf) + RstAcHlpCr!GroupName + vbCrLf + RstAcHlpCr!NameWithADDR
            Else
                TxtDetailS = ""
            End If
        End If
    Else
        If DGAcHlp.Visible = True Then
            FaDGridTxtKeyPress TxtAcName, Index, RstAcHlpDr, KeyAscii, "Name"
            If RstEnviro!AddressHelp = "Yes" And DGAcHlp.Visible = True Then
                TxtDetailS.Visible = True
                TxtDetailS.left = DGAcHlp.left
                TxtDetailS.top = DGAcHlp.top + DGAcHlp.height
                TxtDetailS.width = DGAcHlp.width
            Else
                TxtDetailS.Visible = False
            End If
            If RstAcHlpDr.EOF = False And RstAcHlpDr.BOF = False Then
                TxtDetailS = IIf(Trim(RstAcHlpDr!FatherName) = "", "", RstAcHlpDr!FatherName + vbCrLf) + RstAcHlpDr!GroupName + vbCrLf + RstAcHlpDr!NameWithADDR
            Else
                TxtDetailS = ""
            End If
        End If
    End If
End Sub
Private Sub TxtAcName_LostFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate TxtAcName(Index)
End Sub
Private Sub TxtAcName_Validate(Index As Integer, Cancel As Boolean)
Dim MyCheck As Byte
If mMoveFlag = True Then Exit Sub
If ADDFLAG = 3 Then Exit Sub
If TxtCrDr(Index) = "" Then Exit Sub
If DGAcHlp.Visible = True Then
    If TxtCrDr(Index) = "Cr" Then
        If RstAcHlpCr.RecordCount = 0 Or (RstAcHlpCr.EOF = True Or RstAcHlpCr.BOF = True) Then
            TxtAcName(Index).TEXT = ""
            TxtAcName(Index).Tag = ""
        Else
            TxtAcName(Index).TEXT = RstAcHlpCr!Name
            TxtAcName(Index).Tag = RstAcHlpCr!SubCode
        End If
    Else
        If RstAcHlpDr.RecordCount = 0 Or (RstAcHlpDr.EOF = True Or RstAcHlpDr.BOF = True) Then
            TxtAcName(Index).TEXT = ""
            TxtAcName(Index).Tag = ""
        Else
            TxtAcName(Index).TEXT = RstAcHlpDr!Name
            TxtAcName(Index).Tag = RstAcHlpDr!SubCode
        End If
    End If
End If
If ADDFLAG <= 2 Then
    If Trim(TxtAcName(Index).TEXT) = "" And TxtAcName(Index).Visible = True Then Cancel = True: Exit Sub
    MyCheck = VchTrnSaveLst(Index)
End If
End Sub
Private Sub DGAcHlp_Click()
    If DGAcHlp.Tag = "" Then Exit Sub
    DGAcHlp.Visible = False
    If TxtCrDr(DGAcHlp.Tag) = "Cr" Then
        If RstAcHlpCr.RecordCount > 0 Then
            TxtAcName(Val(DGAcHlp.Tag)).Tag = RstAcHlpCr!SubCode
            TxtAcName(Val(DGAcHlp.Tag)).TEXT = RstAcHlpCr!Name
            If RstEnviro!AddressHelp = "Yes" And DGAcHlp.Visible = True Then
                TxtDetailS.Visible = True
                TxtDetailS.left = DGAcHlp.left
                TxtDetailS.top = DGAcHlp.top + DGAcHlp.height
                TxtDetailS.width = DGAcHlp.width
            Else
                TxtDetailS.Visible = False
            End If
            If RstAcHlpCr.EOF = False And RstAcHlpCr.BOF = False Then
                TxtDetailS = IIf(Trim(RstAcHlpCr!FatherName) = "", "", RstAcHlpCr!FatherName + vbCrLf) + RstAcHlpCr!GroupName + vbCrLf + RstAcHlpCr!NameWithADDR
            Else
                TxtDetailS = ""
            End If
        End If
    Else
        If RstAcHlpDr.RecordCount > 0 Then
            TxtAcName(Val(DGAcHlp.Tag)).Tag = RstAcHlpDr!SubCode
            TxtAcName(Val(DGAcHlp.Tag)).TEXT = RstAcHlpDr!Name
            If RstEnviro!AddressHelp = "Yes" And DGAcHlp.Visible = True Then
                TxtDetailS.Visible = True
                TxtDetailS.left = DGAcHlp.left
                TxtDetailS.top = DGAcHlp.top + DGAcHlp.height
                TxtDetailS.width = DGAcHlp.width
            Else
                TxtDetailS.Visible = False
            End If
            If RstAcHlpDr.EOF = False And RstAcHlpDr.BOF = False Then
                TxtDetailS = IIf(Trim(RstAcHlpDr!FatherName) = "", "", RstAcHlpDr!FatherName + vbCrLf) + RstAcHlpDr!GroupName + vbCrLf + RstAcHlpDr!NameWithADDR
            Else
                TxtDetailS = ""
            End If
        End If
    End If
    TxtAcName(Val(DGAcHlp.Tag)).SetFocus
End Sub
Private Sub Grid_Hide()
    If DGAcHlp.Visible = True Then DGAcHlp.Visible = False
    If DGVchrHlp.Visible = True Then DGVchrHlp.Visible = False
    If TxtDetailS.Visible = True Then TxtDetailS.Visible = False
    If FrameRef.Visible = True Then FrameRef.Visible = False
    If FrameTDS.Visible = True Then FrameTDS.Visible = False
    If DGTDSCODE.Visible = True Then DGTDSCODE.Visible = False
End Sub
Private Sub Ctrl_validate(Ctrl As Object)
    Ctrl.BackColor = CtrlBColOrg
    Ctrl.ForeColor = CtrlFColOrg
End Sub
Private Sub Ctrl_GetFocus(Ctrl As Object)
    Ctrl.BackColor = CtrlBCol
    Ctrl.ForeColor = CtrlFCol
End Sub
Private Sub TxtVtYpe_GotFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Grid_Hide
    If TxtVtYpe(Index).Tag <> "" Then
        FaVrTypeSetting TxtVtYpe(Index).Tag, LblVPrefix
    Else
        FaVrTypeSetting ""
    End If
    Ctrl_GetFocus TxtVtYpe(Index)
End Sub
Private Sub TxtVtYpe_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mMoveFlag = True Then Exit Sub
    If KeyCode = vbKeyEscape Then Grid_Hide
    If ADDFLAG <= 2 Then
        DGVchrHlp.top = TxtVtYpe(Index).top + TxtVtYpe(Index).height + 25
        DGVchrHlp.left = TxtVtYpe(Index).left
        FaDGridTxtKeyDown DGVchrHlp, TxtVtYpe, Index, RstVchrHlp, KeyCode, True, 1
        If RstVchrHlp.EOF = False And RstVchrHlp.BOF = False Then
            LblVPrefix = FaXNull(RstVchrHlp!prefix)
        End If
    End If
End Sub
Private Sub TxtVtYpe_KeyPress(Index As Integer, KeyAscii As Integer)
    If mMoveFlag = True Then Exit Sub
    FaCheckQuote KeyAscii
    If DGVchrHlp.Visible = True Then FaDGridTxtKeyPress TxtVtYpe, Index, RstVchrHlp, KeyAscii, "Description"
    If RstVchrHlp.EOF = False And RstVchrHlp.BOF = False Then
        LblVPrefix = FaXNull(RstVchrHlp!prefix)
    End If
End Sub
Private Sub TxtVtYpe_LostFocus(Index As Integer)
    If mMoveFlag = True Then Exit Sub
    Ctrl_validate TxtVtYpe(Index)
End Sub
Private Sub TxtVtYpe_Validate(Index As Integer, Cancel As Boolean)
Dim MyRs As New ADODB.Recordset, RST1CrDr As ADODB.Recordset
If mMoveFlag = True Then Exit Sub
If ADDFLAG = 3 Then Exit Sub
If RstVchrHlp.RecordCount = 0 Or (RstVchrHlp.EOF = True Or RstVchrHlp.BOF = True) Or TxtVtYpe(Index).TEXT = "" Then
    TxtVtYpe(Index).TEXT = ""
    TxtVtYpe(Index).Tag = ""
    LblVPrefix = ""
Else
    TxtVtYpe(Index).TEXT = RstVchrHlp!Description
    TxtVtYpe(Index).Tag = RstVchrHlp!V_tYPE
    LblVPrefix = FaXNull(RstVchrHlp!prefix)
End If
If ADDFLAG <> 3 Then
    FaVrTypeSetting TxtVtYpe(0).Tag, LblVPrefix
    If TxtVtYpe(0).Tag <> "" Then
        Set RST1CrDr = G_FaCn.Execute("SELECT * FROM VOUCHER_TYPE WHERE V_TYPE='" & TxtVtYpe(0).Tag & "'")
        If RST1CrDr.RecordCount > 0 Then
            TxtCrDr(0) = FaXNull(RST1CrDr!FirstDrCr)
        End If
    End If
    If Trim(TxtCrDr(0)) = "" Then
        Select Case mNCat
            Case "CNT", "RCT"
                If Val(TxtDr(0)) = 0 Then TxtCrDr(0) = "Cr"
            Case "PMT", "JV"
                If Val(TxtCr(0)) = 0 Then TxtCrDr(0) = "Dr"
        End Select
    End If
End If
Set MyRs = Nothing
Set RST1CrDr = Nothing
End Sub
Private Sub DGVchrHlp_KeyDown(KeyCode As Integer, Shift As Integer)
DGVchrHlp.Visible = False
If RstVchrHlp.RecordCount > 0 Then
    TxtVtYpe(0).Tag = RstVchrHlp!V_tYPE
    TxtVtYpe(0).TEXT = RstVchrHlp!Description
End If
End Sub
Private Sub TopCtrl1_eAdd()
On Error GoTo Errloop
Dim RST1 As ADODB.Recordset, I As Integer
    If FRAMEVLIST.Visible = True Then Exit Sub
    PicFalse
    ADDFLAG = 1
    Set RstRef = New ADODB.Recordset
    Set RstRef = PubDatamanFa.FaRefRst(RstRef)
    Set RstTds = New ADODB.Recordset
    Set RstTds = PubDatamanFa.FaTDSRst(RstTds)
    Set RstMainAdj = New ADODB.Recordset
    Set RstMainAdj = PubDatamanFa.FaAdjustRst(RstMainAdj)
    MakeEmpty
    LockFields False
    ScrolIndex = 0
    FGrid1.Rows = 0
    SETS "ADD", Me, RstMain
    If mxLastVrType <> "" Then
        If PubFaSiteType = 2 Then
            Set RST1 = G_FaCn.Execute("SELECT LastVoucher.*,DESCRIPTION,NCAT FROM LastVoucher LEFT JOIN VOUCHER_tYPE ON VOUCHER_tYPE.V_TYPE=LASTVOUCHER.V_TYPE where User_Name=" & FaChk_Text(pubUName) & "  AND LastVoucher.V_tYPE='" & mxLastVrType & "' AND SITE_CODE='" & PubSiteCode & "' ORDER BY LastVoucher.LAST_ENT_DATE DESC")
        Else
            Set RST1 = G_FaCn.Execute("SELECT LastVoucher.*,DESCRIPTION,NCAT FROM LastVoucher LEFT JOIN VOUCHER_tYPE ON VOUCHER_tYPE.V_TYPE=LASTVOUCHER.V_TYPE where User_Name=" & FaChk_Text(pubUName) & "  AND LastVoucher.V_tYPE='" & mxLastVrType & "' ORDER BY LastVoucher.LAST_ENT_DATE DESC")
        End If
        If RST1.RecordCount > 0 Then
            mNCat = RST1!NCat
            mLastVrType = RST1!V_tYPE
            VchDt(0) = Format(RST1!Last_Ent_Date, "dd/MMM/yyyy")
            FaVrTypeSetting mLastVrType, mLastPrefix
            TxtVtYpe_Validate 0, False
        Else
            VchDt(0) = Format(Now(), "dd/MMM/yyyy")
            mNCat = "JV"
            FaVrTypeSetting ""
        End If
    Else
        If PubFaSiteType = 2 Then
            Set RST1 = G_FaCn.Execute("SELECT LastVoucher.*,DESCRIPTION,NCAT FROM LastVoucher LEFT JOIN VOUCHER_tYPE ON VOUCHER_tYPE.V_TYPE=LASTVOUCHER.V_TYPE where User_Name=" & FaChk_Text(pubUName) & " AND SITE_CODE='" & PubSiteCode & "' ORDER BY LastVoucher.LAST_ENT_DATE DESC")
        Else
            Set RST1 = G_FaCn.Execute("SELECT LastVoucher.*,DESCRIPTION,NCAT FROM LastVoucher LEFT JOIN VOUCHER_tYPE ON VOUCHER_tYPE.V_TYPE=LASTVOUCHER.V_TYPE where User_Name=" & FaChk_Text(pubUName) & "  ORDER BY LastVoucher.LAST_ENT_DATE DESC")
        End If
        If RST1.RecordCount > 0 Then
            mNCat = RST1!NCat
            mLastVrType = RST1!V_tYPE
            VchDt(0) = Format(RST1!Last_Ent_Date, "dd/MMM/yyyy")
            FaVrTypeSetting mLastVrType, mLastPrefix
            TxtVtYpe_Validate 0, False
        Else
            VchDt(0) = Format(PubLoginDate, "dd/MMM/yyyy")
            mNCat = "JV"
            TxtCrDr(0) = "Dr"
            FaVrTypeSetting ""
        End If
    End If
    MakeVisible 0, True
    For I = 1 To FixRow
        MakeVisible I, False
    Next
    LblDrAmt(0).CAPTION = Format(0, "0.00")
    LblCrAmt(0).CAPTION = Format(0, "0.00")
    VchDt(0).SetFocus
    Set RST1 = Nothing
    Exit Sub
Errloop:        MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub TxtDr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RstLedger As ADODB.Recordset, RstAdj As ADODB.Recordset, THIS_VR As Double
Dim THIS_OT_VR As Double, AmtAdjusted As Double, mTDSPer As Double, I As Integer
If mMoveFlag = True Then Exit Sub
If KeyCode = vbKeyEscape Then Grid_Hide
If ADDFLAG <= 2 Then
    If KeyCode = vbKeyInsert Then
        Txtdr_Validate Index, False
        If RstEnviro!OnLineAdjustment = "Yes" Then
            If Val(TxtDr(Index)) <= 0 Then Exit Sub
            FRAMEADJUST.Tag = Index
            Label4 = TxtAcName(Index)
            ADJ_LAB4 = Format(Val(TxtDr(Index)), "0.00")
            ADJ_LAB5 = "Dr"
            FgridAdjust.Rows = 1
            AmtAdjusted = 0
            If PubBackEnd = "A" Then
                Set RstLedger = G_FaCn.Execute("Select max(L.DocId) AS DI,MAX(V_SNo) AS VS,MAX(L.V_Type) AS VT,MAX(L.V_Prefix) AS VP,MAX(L.V_No) AS VN,MAX(L.V_Date) AS VD,MAX(AmtCr) As Amount,MAX(L.Narration) AS NARR,MAX(SG.Name) As PartyName,L.SubCode,IIF(IsNull(Sum(CR)),0,Sum(CR)) AS TSum,MAX(LEDGERM.Narration) AS MNARR,MAX(L.AgRefNo) AS AgRefNom FROM ((Ledger L Left Join SubGroup SG on SG.SubCode=L.CONTRASub) LEFT JOIN LEDGERAdj ON (L.SUBCODE=LedgerAdj.SUBCODE) AND (L.V_SNo=LedgerAdj.V_SNo1) AND (L.DocId=LedgerAdj.DocId1)) LEFT JOIN LEDGERM ON LEDGERM.DOCID=L.DOCID Where L.SubCode='" & TxtAcName(Index).Tag & "' GROUP BY L.SUBCODE,L.DOCID,L.V_SNO HAVING MAX(AmtCr)>0")
            ElseIf PubBackEnd = "S" Then
                Set RstLedger = G_FaCn.Execute("Select max(L.DocId) AS DI,MAX(V_SNo) AS VS,MAX(L.V_Type) AS VT,MAX(L.V_Prefix) AS VP,MAX(L.V_No) AS VN,MAX(L.V_Date) AS VD,MAX(AmtCr) As Amount,MAX(L.Narration) AS NARR,MAX(SG.Name) As PartyName,L.SubCode,IsNull(Sum(CR),0) AS TSum,MAX(LEDGERM.Narration) AS MNARR,MAX(L.AgRefNo) AS AgRefNom FROM ((Ledger L Left Join SubGroup SG on SG.SubCode=L.CONTRASub) LEFT JOIN LEDGERAdj ON (L.SUBCODE=LedgerAdj.SUBCODE) AND (L.V_SNo=LedgerAdj.V_SNo1) AND (L.DocId=LedgerAdj.DocId1)) LEFT JOIN LEDGERM ON LEDGERM.DOCID=L.DOCID Where L.SubCode='" & TxtAcName(Index).Tag & "' GROUP BY L.SUBCODE,L.DOCID,L.V_SNO HAVING MAX(AmtCr)>0")
            End If
            Do Until RstLedger.EOF
                If ADDFLAG = 2 Then
                    If PubBackEnd = "A" Then
                        Set RstAdj = G_FaCn.Execute("Select IIF(IsNull(Sum(CR)),0,Sum(CR)) AS TSum FROM LEDGERAdj WHERE SUBCODE='" & TxtAcName(Index).Tag & "' AND DocID2='" & RstMain!DocID & "' And V_SNo2=" & Val(FGrid1.TextMatrix(Index + ScrolIndex, 1)) & " AND DocID1='" & RstLedger!DI & "' And V_SNo1=" & Val(RstLedger!VS))
                    ElseIf PubBackEnd = "S" Then
                        Set RstAdj = G_FaCn.Execute("Select IsNull(Sum(CR),0) AS TSum FROM LEDGERAdj WHERE SUBCODE='" & TxtAcName(Index).Tag & "' AND DocID2='" & RstMain!DocID & "' And V_SNo2=" & Val(FGrid1.TextMatrix(Index + ScrolIndex, 1)) & " AND DocID1='" & RstLedger!DI & "' And V_SNo1=" & Val(RstLedger!VS))
                    End If
                    If RstAdj.RecordCount > 0 Then
                        AmtAdjusted = RstAdj!TSUM
                    End If
                End If
                THIS_VR = 0
                THIS_OT_VR = 0
                If (Index + ScrolIndex = FGrid1.Rows) Or FGrid1.Rows = 0 Then
                Else
                    If RstMainAdj.RecordCount > 0 Then
                        RstMainAdj.Sort = "VSNo ASC"
                        RstMainAdj.MoveFirst
                        Do Until RstMainAdj.EOF
                            If RstMainAdj!DocID = RstLedger!DI And RstMainAdj!V_SNo = RstLedger!VS Then
                                If RstMainAdj!VSNO = Val(FGrid1.TextMatrix(Index + ScrolIndex, 1)) Then
                                    THIS_VR = THIS_VR + FaVNull(RstMainAdj!cr)
                                Else
                                    THIS_OT_VR = THIS_OT_VR + FaVNull(RstMainAdj!cr)
                                End If
                            End If
                            RstMainAdj.MoveNext
                        Loop
                    End If
                End If
                If RstLedger!Amount - RstLedger!TSUM + AmtAdjusted - THIS_OT_VR > 0 Then
                    FgridAdjust.AddItem "" & Chr(9) & RstLedger!VT & Chr(9) & RstLedger!VN & Chr(9) & RstLedger!VP & Chr(9) & RstLedger!VS & Chr(9) & Format(RstLedger!VD, "Short Date") & Chr(9) & RstLedger!PartyName & Chr(9) & RstLedger!AgRefNom & Chr(9) & Format(RstLedger!Amount, "0.00") & Chr(9) & Format(RstLedger!Amount - RstLedger!TSUM + AmtAdjusted - THIS_OT_VR, "0.00") & Chr(9) & Format(THIS_VR, "0.00") & Chr(9) & FaXNull(RstLedger!mNarr) + " " + FaXNull(RstLedger!Narr) & Chr(9) & "" & Chr(9) & RstLedger!DI
                End If
                RstLedger.MoveNext
            Loop
            TXTNARRATION = ""
            CAL_ADJ_TOT
            If FgridAdjust.Rows > 1 Then FRAMEADJUST.Visible = True: FRAMEADJUST.ZOrder 0
            Exit Sub
        End If
    ElseIf KeyCode = vbKeyR And Shift = 4 Then
        Txtdr_Validate Index, False
        If RstEnviro!OnLineAdjustment = "Yes" Then
            If Val(TxtDr(Index)) <= 0 Then Exit Sub
            FrameRef.Tag = Index
            LblRefName = TxtAcName(Index)
            LblRefAmt = Format(Val(TxtDr(Index)), "0.00")
            LblRefAmtDrCr = "Dr"
            FGridRef.Rows = 1
            FGridRef.AddItem ""
            FGridRef.FixedRows = 1
            I = 1
            If RstRef.RecordCount > 0 Then
                RstRef.Sort = "V_SNo ASC"
                RstRef.MoveFirst
                RstRef.Find "V_SNO=" & Val(FGrid1.TextMatrix(Index + ScrolIndex, 1))
                Do Until RstRef.EOF
                    If RstRef!V_SNo = Val(FGrid1.TextMatrix(Index + ScrolIndex, 1)) Then
                        FGridRef.AddItem ""
                        FGridRef.TextMatrix(I, FDocId) = RstRef!DocID
                        FGridRef.TextMatrix(I, FV_Sno) = RstRef!V_SNo
                        FGridRef.TextMatrix(I, FDr) = FaBNull(RstRef!dr)
                        FGridRef.TextMatrix(I, FCr) = FaBNull(RstRef!cr)
                        FGridRef.TextMatrix(I, FSubCode) = RstRef!SubCode
                        FGridRef.TextMatrix(I, FAgRefNo) = RstRef!AgRefNo
                        FGridRef.TextMatrix(I, FAgRefType) = RstRef!AgRefType
                        FGridRef.TextMatrix(I, FDueDate) = RstRef!DUEDATE
                    End If
                    RstRef.MoveNext
                    I = I + 1
                Loop
            End If
            CalAmountLF
            FrameRef.Visible = True: FrameRef.ZOrder 0: FGridRef.Col = 1: FGridRef.SetFocus
            Exit Sub
        End If
    ElseIf KeyCode = vbKeyT And Shift = 4 Then
        Txtdr_Validate Index, False
        If Val(TxtDr(Index)) <= 0 Then Exit Sub
        If PubFaSiteType = 1 Then
            Set RstLedger = G_FaCn.Execute("SELECT TDSCAT.TDS_Catg AS CODE,TDSCAT.TDS_Desc AS NAME,TDSCAT.TDS_Percentage AS TDSPercentage,TDSCAT.TDS_Limit AS TDSLIMIT FROM TDSCAT LEFT JOIN SUBGROUP ON SUBGROUP.TDS_CATG=TDSCAT.TDS_Catg WHERE SUBGROUP.SUBCODE='" & TxtAcName(Index).Tag & "'")
        Else
            Set RstLedger = G_FaCn.Execute("SELECT TDSCAT.* FROM TDSCAT LEFT JOIN SUBGROUP ON SUBGROUP.TDS_CATG=TDSCAT.CODE WHERE SUBGROUP.SUBCODE='" & TxtAcName(Index).Tag & "'")
        End If
        If RstLedger.RecordCount > 0 Then
            mTDSPer = FaVNull(RstLedger!TDSPercentage)
        Else
            mTDSPer = 0
        End If
'        If RstLedger!TDSLimit > Val(TxtDr(Index)) Then Exit Sub
        FrameTDS.Tag = Index
        If RstTds.RecordCount > 0 Then
            RstTds.Sort = "V_SNo ASC"
            RstTds.MoveFirst
            RstTds.Find "V_SNO=" & Val(FGrid1.TextMatrix(Index + ScrolIndex, 1))
            If RstTds.EOF = False Then
                TxtTDSCode(0) = RstTds!TDSNAME
                TxtTDSCode(0).Tag = RstTds!TDSCODE
                TxtTDSNarration = FaXNull(RstTds!Narration)
                TxtONAMT = RstTds!ONAmt
                TxtTDS = RstTds!TDS
                TxtTDSAMT = RstTds!TDSAmt
                TxtDr(Index).Tag = RstTds!TDSAmt
            Else
                TxtTDSCode(0) = ""
                TxtTDSCode(0).Tag = ""
                TxtTDSNarration = ""
                TxtONAMT = Val(TxtDr(Index))
                TxtTDS = mTDSPer
                TxtTDSAMT = (Val(TxtONAMT) * Val(TxtTDS)) / 100
            End If
        Else
            TxtTDSCode(0) = ""
            TxtTDSCode(0).Tag = ""
            TxtTDSNarration = ""
            TxtONAMT = Val(TxtDr(Index))
            TxtTDS = mTDSPer
            TxtTDSAMT = ""
        End If
        FrameTDS.Visible = True
        FrameTDS.ZOrder 0
        TxtTDSCode(0).SetFocus
        DGTDSCODE.left = FrameTDS.left + TxtTDSCode(0).left
        DGTDSCODE.top = FrameTDS.top + TxtTDSCode(0).top + TxtTDSCode(0).height + 50
        Exit Sub
    End If
End If
Set RstLedger = Nothing
Set RstAdj = Nothing
End Sub
Private Sub TxtCr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RstLedger As ADODB.Recordset, RstAdj As ADODB.Recordset, I As Integer
Dim THIS_VR As Double, THIS_OT_VR As Double, AmtAdjusted As Double, mTDSPer As Double
If mMoveFlag = True Then Exit Sub
If KeyCode = vbKeyEscape Then Grid_Hide
If ADDFLAG <= 2 Then
    If KeyCode = vbKeyInsert Then
        TxtCr_Validate Index, False
        If RstEnviro!OnLineAdjustment = "Yes" Then
            If Val(TxtCr(Index)) <= 0 Then Exit Sub
            FRAMEADJUST.Tag = Index
            Label4 = TxtAcName(Index)
            ADJ_LAB4 = Format(Val(TxtCr(Index)), "0.00")
            ADJ_LAB5 = "Cr"
            FgridAdjust.Rows = 1
            AmtAdjusted = 0
            If PubBackEnd = "A" Then
                Set RstLedger = G_FaCn.Execute("Select max(L.DocId) AS DI,MAX(V_SNo) AS VS,MAX(L.V_Type) AS VT,MAX(L.V_Prefix) AS VP,MAX(L.V_No) AS VN,MAX(L.V_Date) AS VD,MAX(AmtDr) As Amount,MAX(L.Narration) AS NARR,MAX(SG.Name) As PartyName,L.SubCode,IIF(IsNull(Sum(CR)),0,Sum(CR)) AS TSum,MAX(LEDGERM.Narration) AS MNARR,MAX(L.AgRefNo) AS AgRefNom FROM ((Ledger L Left Join SubGroup SG on SG.SubCode=L.CONTRASub) LEFT JOIN LEDGERAdj ON  (L.SUBCODE = LedgerAdj.SUBCODE) AND (L.V_SNo = LedgerAdj.V_SNo2) AND (L.DocId = LedgerAdj.DocId2)) LEFT JOIN LEDGERM ON LEDGERM.DOCID=L.DOCID Where L.SubCode='" & TxtAcName(Index).Tag & "' GROUP BY L.SUBCODE,L.DOCID,L.V_SNO HAVING MAX(AmtDr)>0")
            ElseIf PubBackEnd = "S" Then
                Set RstLedger = G_FaCn.Execute("Select max(L.DocId) AS DI,MAX(V_SNo) AS VS,MAX(L.V_Type) AS VT,MAX(L.V_Prefix) AS VP,MAX(L.V_No) AS VN,MAX(L.V_Date) AS VD,MAX(AmtDr) As Amount,MAX(L.Narration) AS NARR,MAX(SG.Name) As PartyName,L.SubCode,IsNull(Sum(CR),0) AS TSum,MAX(LEDGERM.Narration) AS MNARR,MAX(L.AgRefNo) AS AgRefNom FROM ((Ledger L Left Join SubGroup SG on SG.SubCode=L.CONTRASub) LEFT JOIN LEDGERAdj ON  (L.SUBCODE = LedgerAdj.SUBCODE) AND (L.V_SNo = LedgerAdj.V_SNo2) AND (L.DocId = LedgerAdj.DocId2)) LEFT JOIN LEDGERM ON LEDGERM.DOCID=L.DOCID Where L.SubCode='" & TxtAcName(Index).Tag & "' GROUP BY L.SUBCODE,L.DOCID,L.V_SNO HAVING MAX(AmtDr)>0")
            End If
            Do Until RstLedger.EOF
                If ADDFLAG = 2 Then
                    If PubBackEnd = "A" Then
                        Set RstAdj = G_FaCn.Execute("Select IIF(IsNull(Sum(CR)),0,Sum(CR)) AS TSum FROM LEDGERAdj WHERE SUBCODE='" & TxtAcName(Index).Tag & "' AND DocID1='" & RstMain!DocID & "' And V_SNo1=" & Val(FGrid1.TextMatrix(Index + ScrolIndex, 1)) & " AND DocID2='" & RstLedger!DI & "' And V_SNo2=" & Val(RstLedger!VS))
                    ElseIf PubBackEnd = "S" Then
                        Set RstAdj = G_FaCn.Execute("Select IsNull(Sum(CR),0) AS TSum FROM LEDGERAdj WHERE SUBCODE='" & TxtAcName(Index).Tag & "' AND DocID1='" & RstMain!DocID & "' And V_SNo1=" & Val(FGrid1.TextMatrix(Index + ScrolIndex, 1)) & " AND DocID2='" & RstLedger!DI & "' And V_SNo2=" & Val(RstLedger!VS))
                        End If
                    If RstAdj.RecordCount > 0 Then
                        AmtAdjusted = RstAdj!TSUM
                    End If
                End If
                THIS_VR = 0
                THIS_OT_VR = 0
                If (Index + ScrolIndex = FGrid1.Rows) Or FGrid1.Rows = 0 Then
                Else
                    If RstMainAdj.RecordCount > 0 Then
                        RstMainAdj.Sort = "VSNo ASC"
                        RstMainAdj.MoveFirst
                        Do Until RstMainAdj.EOF
                            If RstMainAdj!DocID = RstLedger!DI And RstMainAdj!V_SNo = RstLedger!VS Then
                                If RstMainAdj!VSNO = Val(FGrid1.TextMatrix(Index + ScrolIndex, 1)) Then
                                    THIS_VR = THIS_VR + FaVNull(RstMainAdj!dr)
                                Else
                                    THIS_OT_VR = THIS_OT_VR + FaVNull(RstMainAdj!dr)
                                End If
                            End If
                            RstMainAdj.MoveNext
                        Loop
                    End If
                End If
                If RstLedger!Amount - RstLedger!TSUM + AmtAdjusted - THIS_OT_VR > 0 Then
                    FgridAdjust.AddItem "" & Chr(9) & RstLedger!VT & Chr(9) & RstLedger!VN & Chr(9) & RstLedger!VP & Chr(9) & RstLedger!VS & Chr(9) & Format(RstLedger!VD, "Short Date") & Chr(9) & RstLedger!PartyName & Chr(9) & RstLedger!AgRefNom & Chr(9) & Format(RstLedger!Amount, "0.00") & Chr(9) & Format(RstLedger!Amount - RstLedger!TSUM + AmtAdjusted - THIS_OT_VR, "0.00") & Chr(9) & Format(THIS_VR, "0.00") & Chr(9) & FaXNull(RstLedger!mNarr) + " " + FaXNull(RstLedger!Narr) & Chr(9) & "" & Chr(9) & RstLedger!DI
                End If
                RstLedger.MoveNext
            Loop
            TXTNARRATION = ""
            CAL_ADJ_TOT
            If FgridAdjust.Rows > 1 Then FRAMEADJUST.Visible = True: FRAMEADJUST.ZOrder 0
            Exit Sub
        End If
    ElseIf KeyCode = vbKeyR And Shift = 4 Then
        TxtCr_Validate Index, False
        If RstEnviro!OnLineAdjustment = "Yes" Then
            If Val(TxtCr(Index)) <= 0 Then Exit Sub
            FrameRef.Tag = Index
            LblRefName = TxtAcName(Index)
            LblRefAmt = Format(Val(TxtCr(Index)), "0.00")
            LblRefAmtDrCr = "Cr"
            FGridRef.Rows = 1
            FGridRef.AddItem ""
            FGridRef.FixedRows = 1
            I = 1
            If RstRef.RecordCount > 0 Then
                RstRef.Sort = "V_SNo ASC"
                RstRef.MoveFirst
                RstRef.Find "V_SNO=" & Val(FGrid1.TextMatrix(Index + ScrolIndex, 1))
                Do Until RstRef.EOF
                    If RstRef!V_SNo = Val(FGrid1.TextMatrix(Index + ScrolIndex, 1)) Then
                        FGridRef.AddItem ""
                        FGridRef.TextMatrix(I, FDocId) = RstRef!DocID
                        FGridRef.TextMatrix(I, FV_Sno) = RstRef!V_SNo
                        FGridRef.TextMatrix(I, FDr) = FaBNull(RstRef!dr)
                        FGridRef.TextMatrix(I, FCr) = FaBNull(RstRef!cr)
                        FGridRef.TextMatrix(I, FSubCode) = RstRef!SubCode
                        FGridRef.TextMatrix(I, FAgRefNo) = RstRef!AgRefNo
                        FGridRef.TextMatrix(I, FAgRefType) = RstRef!AgRefType
                        FGridRef.TextMatrix(I, FDueDate) = RstRef!DUEDATE
                    End If
                    RstRef.MoveNext
                    I = I + 1
                Loop
            End If
            CalAmountLF
            FrameRef.Visible = True: FrameRef.ZOrder 0: FGridRef.SetFocus
            Exit Sub
        End If
    ElseIf KeyCode = vbKeyT And Shift = 4 Then
        TxtCr_Validate Index, False
        If Val(TxtCr(Index)) <= 0 Then Exit Sub
        If PubFaSiteType = 1 Then
            Set RstLedger = G_FaCn.Execute("SELECT TDSCAT.TDS_Catg AS CODE,TDSCAT.TDS_Desc AS NAME,TDSCAT.TDS_Percentage AS TDSPercentage,TDSCAT.TDS_Limit AS TDSLIMIT FROM TDSCAT LEFT JOIN SUBGROUP ON SUBGROUP.TDS_CATG=TDSCAT.TDS_Catg WHERE SUBGROUP.SUBCODE='" & TxtAcName(Index).Tag & "'")
        Else
            Set RstLedger = G_FaCn.Execute("SELECT TDSCAT.* FROM TDSCAT LEFT JOIN SUBGROUP ON SUBGROUP.TDS_CATG=TDSCAT.CODE WHERE SUBGROUP.SUBCODE='" & TxtAcName(Index).Tag & "'")
        End If
        If RstLedger.RecordCount > 0 Then
            mTDSPer = FaVNull(RstLedger!TDSPercentage)
        Else
            mTDSPer = 0
        End If
'        If RstLedger!TDSLimit > Val(TxtDr(Index)) Then Exit Sub
        FrameTDS.Tag = Index
        If RstTds.RecordCount > 0 Then
            RstTds.Sort = "V_SNo ASC"
            RstTds.MoveFirst
            RstTds.Find "V_SNO=" & Val(FGrid1.TextMatrix(Index + ScrolIndex, 1))
            If RstTds.EOF = False Then
                TxtTDSCode(0) = RstTds!TDSNAME
                TxtTDSCode(0).Tag = RstTds!TDSCODE
                TxtTDSNarration = FaXNull(RstTds!Narration)
                TxtONAMT = RstTds!ONAmt
                TxtTDS = RstTds!TDS
                TxtTDSAMT = RstTds!TDSAmt
                TxtDr(Index).Tag = RstTds!TDSAmt
            Else
                TxtTDSCode(0) = ""
                TxtTDSCode(0).Tag = ""
                TxtTDSNarration = ""
                TxtONAMT = Val(TxtCr(Index))
                TxtTDS = mTDSPer
                TxtTDSAMT = (Val(TxtONAMT) * Val(TxtTDS)) / 100
            End If
        Else
            TxtTDSCode(0) = ""
            TxtTDSCode(0).Tag = ""
            TxtTDSNarration = ""
            TxtONAMT = Val(TxtCr(Index))
            TxtTDS = mTDSPer
            TxtTDSAMT = ""
        End If
        FrameTDS.Visible = True
        FrameTDS.ZOrder 0
        TxtTDSCode(0).SetFocus
        DGTDSCODE.left = FrameTDS.left + TxtTDSCode(0).left
        DGTDSCODE.top = FrameTDS.top + TxtTDSCode(0).top + TxtTDSCode(0).height + 50
        Exit Sub
    End If
End If
Set RstLedger = Nothing
Set RstAdj = Nothing
End Sub
Private Sub BTS_AUTO_ADJ_Click()
Dim K As Integer, OLD_AMT_ADJ As Long
For K = FgridAdjust.Row To FgridAdjust.Rows - 1
    If Val(ADJ_LAB4) > Val(ADJ_LAB7) Then
        If Val(FgridAdjust.TextMatrix(K, 10)) < Val(FgridAdjust.TextMatrix(K, 9)) And Val(ADJ_LAB7) <= ADJ_LAB4 Then
            OLD_AMT_ADJ = Val(FgridAdjust.TextMatrix(K, 10))
            If (Val(ADJ_LAB4) - (Val(ADJ_LAB7) - OLD_AMT_ADJ)) >= Val(FgridAdjust.TextMatrix(K, 9)) Then
                FgridAdjust.TextMatrix(K, 10) = Val(FgridAdjust.TextMatrix(K, 9))
            Else
                FgridAdjust.TextMatrix(K, 10) = (Val(ADJ_LAB4) - (Val(ADJ_LAB7) - OLD_AMT_ADJ))
            End If
        End If
    End If
    UPD_ADJ
Next
End Sub
Private Sub Command2_Click()
Dim OLD_AMT_ADJ As Long
If Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 10)) < Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 9)) And Val(ADJ_LAB7) < ADJ_LAB4 Then
    OLD_AMT_ADJ = Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 10))
    If (Val(ADJ_LAB4) - (Val(ADJ_LAB7) - OLD_AMT_ADJ)) >= Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 9)) Then
        FgridAdjust.TextMatrix(FgridAdjust.Row, 10) = Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 9))
    Else
        FgridAdjust.TextMatrix(FgridAdjust.Row, 10) = (Val(ADJ_LAB4) - (Val(ADJ_LAB7) - OLD_AMT_ADJ))
    End If
    UPD_ADJ
End If
End Sub
Private Sub UPD_ADJ()
CURR_ROW = FgridAdjust.Row
If FgridAdjust.Col = 10 Then
    TXTADJ_AMT.Visible = True
    TXTADJ_AMT.ZOrder 0
    TXTADJ_AMT.width = FgridAdjust.ColWidth(9)
    TXTADJ_AMT.top = FgridAdjust.top + FgridAdjust.CellTop
    TXTADJ_AMT.left = FgridAdjust.left + FgridAdjust.CellLeft
    TXTADJ_AMT.TEXT = Val(FgridAdjust.TextMatrix(CURR_ROW, 10))
    TXTADJ_AMT.SetFocus
End If
CAL_ADJ_TOT
End Sub
Private Sub CAL_ADJ_TOT()
Dim I As Integer, AMT_ADJ As Double
AMT_ADJ = 0
For I = 1 To FgridAdjust.Rows - 1
    AMT_ADJ = AMT_ADJ + Val(FgridAdjust.TextMatrix(I, 10))
Next
ADJ_LAB7 = Format(AMT_ADJ, "0.00")
ADJ_LAB8 = IIf(ADJ_LAB5 = "Cr", "Dr", "Cr")
End Sub
Private Sub ADJ_CANCLE_Click()
    FRAMEADJUST.Visible = False
    If TxtCrDr(Val(FRAMEADJUST.Tag)) = "Cr" Then
        TxtCr(Val(FRAMEADJUST.Tag)).SetFocus
    Else
        TxtDr(Val(FRAMEADJUST.Tag)).SetFocus
    End If
End Sub
Private Sub ADJ_OK_Click()
Dim I As Integer
    FRAMEADJUST.Visible = False
    VchTrnSaveLst Val(FRAMEADJUST.Tag)
    If RstMainAdj.RecordCount > 0 Then
        RstMainAdj.Sort = "VSNo ASC"
        RstMainAdj.Find "VSNO=" & Val(FGrid1.TextMatrix(Val(FRAMEADJUST.Tag) + ScrolIndex, 1))
        If RstMainAdj.EOF = False Then
            Do While RstMainAdj!VSNO = Val(FGrid1.TextMatrix(Val(FRAMEADJUST.Tag) + ScrolIndex, 1))
                RstMainAdj.Delete
                RstMainAdj.MoveNext
                If RstMainAdj.EOF = True Then Exit Do
            Loop
        End If
    End If
    For I = 1 To FgridAdjust.Rows - 1
        If Val(FgridAdjust.TextMatrix(I, 10)) > 0 Then
            With RstMainAdj
                .AddNew
                .Fields("DocId") = FgridAdjust.TextMatrix(I, 13)
                .Fields("V_SNo") = Val(FgridAdjust.TextMatrix(I, 4))
                .Fields("VSNo") = Val(FGrid1.TextMatrix(Val(FRAMEADJUST.Tag) + ScrolIndex, 1))
                If ADJ_LAB8 = "Cr" Then
                    .Fields("CR") = Val(FgridAdjust.TextMatrix(I, 10))
                Else
                    .Fields("DR") = Val(FgridAdjust.TextMatrix(I, 10))
                End If
                .Fields("SUBCODE") = FGrid1.TextMatrix(Val(FRAMEADJUST.Tag) + ScrolIndex, 2)
                .Fields("AgRefNo") = FgridAdjust.TextMatrix(I, 7)
                .Update
            End With
        End If
    Next
    If TxtCrDr(Val(FRAMEADJUST.Tag)) = "Cr" Then
        TxtCr(Val(FRAMEADJUST.Tag)).SetFocus
    Else
        TxtDr(Val(FRAMEADJUST.Tag)).SetFocus
    End If
End Sub
Private Sub FGRIDADJUST_Click()
    UPD_ADJ
End Sub
Private Sub FGRIDADJUST_KeyUp(KeyCode As Integer, Shift As Integer)
    UPD_ADJ
End Sub
Private Sub FGRIDADJUST_Scroll()
    UPD_ADJ
End Sub
Private Sub TXTADJ_AMT_GotFocus()
    OLD_AMT1 = Val(FgridAdjust.TextMatrix(FgridAdjust.Row, 10))
    SendKeys "{Home}+{End}"
End Sub
Private Sub TXTADJ_AMT_KeyPress(KeyAscii As Integer)
    FaNumPress TXTADJ_AMT, KeyAscii, 10, 2
End Sub
Private Sub TXTADJ_AMT_KeyDown(KeyCode As Integer, Shift As Integer)
    FaNumDown TXTADJ_AMT, KeyCode, 10, 2
    If KeyCode = vbKeyReturn Then TXTADJ_AMT_Validate False
End Sub
Private Sub TXTADJ_AMT_Validate(Cancel As Boolean)
    TXTADJ_AMT = FaValidate_Numeric(TXTADJ_AMT)
    If Val(TXTADJ_AMT) > Abs(Val(ADJ_LAB4) - Val(ADJ_LAB7) + OLD_AMT1) Or Val(FgridAdjust.TextMatrix(CURR_ROW, 9)) < Val(TXTADJ_AMT.TEXT) Then
        MsgBox " Amount is Greater Then Pendng Adj.Amt. You Can Adjust Only " + LTrim(RTrim(IIf(Val(FgridAdjust.TextMatrix(CURR_ROW, 9)) > Val(ADJ_LAB4) - Val(ADJ_LAB7) + OLD_AMT1, Val(ADJ_LAB4) - Val(ADJ_LAB7) + OLD_AMT1, FgridAdjust.TextMatrix(CURR_ROW, 9)))) + " Here", vbCritical, Me.CAPTION
        FgridAdjust.Row = CURR_ROW
        Cancel = True
    Else
        FgridAdjust.TextMatrix(CURR_ROW, 10) = Format(TXTADJ_AMT, "0.00")
        CAL_ADJ_TOT
        TXTADJ_AMT.Visible = False
    End If
End Sub
Private Sub FgridAdjust_RowColChange()
    TXTNARRATION = FgridAdjust.TextMatrix(FgridAdjust.Row, 11)
End Sub
Private Sub TxtRef_Validate(Cancel As Boolean)
'''''    If FrameRef.Visible = True Then FrameRef.Visible = False
'''''    If RstRef.RecordCount > 0 Then
'''''        RstRef.Sort = "V_SNo ASC"
'''''        RstRef.Find "V_SNO=" & Val(FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 1))
'''''        If RstRef.EOF = False Then
'''''            Do While RstRef!V_SNo = Val(FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 1))
'''''                RstRef.Delete
'''''                RstRef.MoveNext
'''''                If RstRef.EOF = True Then Exit Do
'''''            Loop
'''''        End If
'''''    End If
'''''    If Trim(TxtRef) <> "" Then
'''''        With RstRef
'''''            .AddNew
'''''            .Fields("V_SNo") = Val(FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 1))
'''''            .Fields("AgRefNo") = Trim(TxtRef)
'''''            .Update
'''''        End With
'''''    End If
'''''    If TxtCrDr(Val(FrameRef.Tag)) = "Cr" Then
'''''        TxtCr(Val(FrameRef.Tag)).SetFocus
'''''    Else
'''''        TxtDr(Val(FrameRef.Tag)).SetFocus
'''''    End If
End Sub
Private Sub TxtTDSAMT_Validate(Cancel As Boolean)
    If FrameTDS.Visible = True Then FrameTDS.Visible = False
    If RstTds.RecordCount > 0 Then
        RstTds.Sort = "V_SNo ASC"
        RstTds.Find "V_SNO=" & Val(FGrid1.TextMatrix(Val(FrameTDS.Tag) + ScrolIndex, 1))
        If RstTds.EOF = False Then
            Do While RstTds!V_SNo = Val(FGrid1.TextMatrix(Val(FrameTDS.Tag) + ScrolIndex, 1))
                RstTds.Delete
                RstTds.MoveNext
                If RstTds.EOF = True Then Exit Do
            Loop
        End If
    End If
    With RstTds
        .AddNew
        .Fields("V_SNo") = Val(FGrid1.TextMatrix(Val(FrameTDS.Tag) + ScrolIndex, 1))
        .Fields("TDSCode") = TxtTDSCode(0).Tag
        .Fields("TDSName") = TxtTDSCode(0)
        .Fields("TDSDrCode") = TxtAcName(Val(FrameTDS.Tag)).Tag
        .Fields("ONAMT") = Val(TxtONAMT)
        .Fields("TDS") = Val(TxtTDS)
        .Fields("TDSAMT") = Val(TxtTDSAMT)
        .Fields("NARRATION") = TxtTDSNarration
        .Update
    End With
    If TxtCrDr(Val(FrameTDS.Tag)) = "Cr" Then
        TxtCr(Val(FrameTDS.Tag)).SetFocus
    Else
        TxtDr(Val(FrameTDS.Tag)).SetFocus
    End If
End Sub
Private Sub TxtTDSCode_GotFocus(Index As Integer)
If DGTDSCODE.Visible = True Then DGTDSCODE.Visible = False
If ADDFLAG <= 2 Then
    If RstTDSHlp.RecordCount > 0 Then
        Set DGTDSCODE.DataSource = Nothing
        RstTDSHlp.MoveFirst
        Set DGTDSCODE.DataSource = RstTDSHlp
    Else
        Set DGTDSCODE.DataSource = Nothing
    End If
    If RstTDSHlp.RecordCount = 0 Or TxtTDSCode(0).TEXT = "" Then Exit Sub
    DGTDSCODE.Tag = Index
    If TxtTDSCode(0).TEXT <> "" Then
        RstTDSHlp.MoveFirst
        RstTDSHlp.Find "Name='" & TxtTDSCode(0).TEXT & "'"
        DGTDSCODE.ReBind
    Else
        RstTDSHlp.MoveFirst
    End If
End If
End Sub
Private Sub TxtTDSCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then DGTDSCODE.Visible = False: Exit Sub
If ADDFLAG <= 2 Then FaDGridTxtKeyDown DGTDSCODE, TxtTDSCode, 0, RstTDSHlp, KeyCode, True, 1
End Sub
Private Sub TxtTDSCode_KeyPress(Index As Integer, KeyAscii As Integer)
FaCheckQuote KeyAscii
If DGTDSCODE.Visible = True Then FaDGridTxtKeyPress TxtTDSCode, 0, RstTDSHlp, KeyAscii, "Name"
End Sub
Private Sub TxtTDSCode_Validate(Index As Integer, Cancel As Boolean)
Dim MyCheck As Byte
If ADDFLAG = 3 Then Exit Sub
If RstTDSHlp.RecordCount = 0 Or (RstTDSHlp.EOF = True Or RstTDSHlp.BOF = True) Then
    TxtTDSCode(0).TEXT = ""
    TxtTDSCode(0).Tag = ""
Else
    TxtTDSCode(0).TEXT = RstTDSHlp!Name
    TxtTDSCode(0).Tag = RstTDSHlp!SubCode
End If
If ADDFLAG <= 2 Then
    If Trim(TxtTDSCode(0).TEXT) = "" And TxtTDSCode(0).Visible = True Then Cancel = True: Exit Sub
End If
End Sub
Private Sub DGTDSCODE_Click()
    If DGTDSCODE.Visible = True Then DGTDSCODE.Visible = False
    If RstTDSHlp.RecordCount > 0 Then
        TxtTDSCode(0).Tag = RstTDSHlp!SubCode
        TxtTDSCode(0).TEXT = RstTDSHlp!Name
    End If
    TxtTDSCode(0).SetFocus
End Sub
Private Sub TDSDelete_Click()
Dim RST1 As ADODB.Recordset
If ADDFLAG = 2 Then
    Set RST1 = G_FaCn.Execute("SELECT * FROM LEDGERTDS WHERE DOCID='" & RstMain!DocID & "' AND TDSPOST='Y'")
    If RST1.RecordCount > 0 Then
        MsgBox "T.D.S.Challan Already Made,Can't delete it"
        Exit Sub
    End If
Else
    If RstTds.RecordCount > 0 Then
        RstTds.Sort = "V_SNo ASC"
        RstTds.Find "V_SNO=" & Val(FGrid1.TextMatrix(Val(FrameTDS.Tag) + ScrolIndex, 1))
        If RstTds.EOF = False Then
            Do While RstTds!V_SNo = Val(FGrid1.TextMatrix(Val(FrameTDS.Tag) + ScrolIndex, 1))
                If FaXNull(RstTds!TDSPOST) = "Y" Then
                    MsgBox "T.D.S.Challan Already Made,Can't delete it"
                    Exit Sub
                End If
                RstTds.Delete
                RstTds.MoveNext
                If RstTds.EOF = True Then Exit Do
            Loop
        End If
    End If
End If
TxtTDSCode(0).Tag = ""
TxtTDSCode(0) = ""
TxtONAMT = ""
TxtTDS = ""
TxtTDSAMT = ""
TxtTDSNarration = ""
FrameTDS.Visible = False
Set RST1 = Nothing
End Sub
Private Function FaNarr() As String
On Error GoTo ELoop
    PubDatamanFa.FaGlobeNarrForm.Show vbModal
    FaNarr = PubDatamanFa.FaGNarrFunc
Exit Function
ELoop:    If err.NUMBER <> 0 Then MsgBox err.Description, vbInformation, "Validation"
End Function
Private Sub PicDisp(Index As Integer)
    If ScrolIndex > 0 Then
        PicUP.Visible = True
    Else
        PicUP.Visible = False
    End If
    If (FGrid1.Rows - 1 - ScrolIndex - FixRow) > 0 Then
        PicDN.Visible = True
    Else
        PicDN.Visible = False
    End If
End Sub
Private Sub PicFalse()
    PicUP.Visible = False
    PicDN.Visible = False
End Sub
Private Sub Ini_Grid()
    Me.TopCtrl1.TopText1.left = 5800
    FRAMEVLIST.top = 630
    FRAMEVLIST.left = 2010
    Frame1(1).top = 630
    Frame1(1).left = 2010
    DGAcHlp.left = Label1(1).left
    DGAcHlp.top = Label1(1).top
    DGAcHlp.height = Line2(0).Y1 - DGAcHlp.top
    FrameTDS.left = 2000: FrameTDS.top = 1035
    FRAMEADJUST.left = 0: FRAMEADJUST.top = 1035
    FgridAdjust.left = 90: FgridAdjust.top = 465
    FRAMEADJUST.width = 11680: FRAMEADJUST.height = 5300
    FgridAdjust.width = 10700: FgridAdjust.height = 4000
    TXTNARRATION.top = FgridAdjust.top + FgridAdjust.height
    TXTNARRATION.width = FgridAdjust.width
    BTS_AUTO_ADJ.left = 10800
    Command2.left = 10800
    ADJ_OK.left = 10800
    ADJ_CANCLE.left = 10800
    FgridAdjust.ColAlignment(2) = flexAlignLeftCenter
    FgridAdjust.ColAlignment(5) = flexAlignLeftCenter
    FgridAdjust.ColAlignment(7) = flexAlignLeftCenter
    FgridAdjust.ColAlignment(8) = flexAlignRightCenter
    FgridAdjust.ColAlignment(9) = flexAlignRightCenter
    FgridAdjust.ColAlignment(10) = flexAlignRightCenter
    FgridAdjust.ColWidth(3) = 0
    FgridAdjust.ColWidth(11) = 0
    FgridAdjust.ColWidth(12) = 0
    FgridAdjust.ColWidth(13) = 0
    FrameRef.left = Me.left: FrameRef.top = TxtCrDr(0).top: FrameRef.width = TxtCr(0).left + TxtCr(0).width: FrameRef.height = Frame1(0).height - FrameRef.top
    FGridRef.left = FrameRef.left + 50: FGridRef.width = FrameRef.width - 100: FGridRef.height = FrameRef.height - 100
    With FGridRef
        BackColorSelLeave = .BackColor
        .Cols = 9
        .ColWidth(0) = 350
        .ColWidth(FDocId) = 0
        .ColWidth(FV_Sno) = 0
        .ColWidth(FSubCode) = 0
        .TextMatrix(0, FAgRefType) = "Adj.Type"
        .ColAlignment(FAgRefType) = flexAlignLeftCenter
        .ColWidth(FAgRefType) = 2000
        .TextMatrix(0, FAgRefNo) = "Ref."
        .ColAlignmentFixed(FAgRefNo) = flexAlignLeftCenter
        .ColAlignment(FAgRefNo) = flexAlignRightCenter
        .ColWidth(FAgRefNo) = 2000
        .TextMatrix(0, FDr) = "Debit"
        .ColAlignmentFixed(FDr) = flexAlignRightCenter
        .ColAlignment(FDr) = flexAlignRightCenter
        .ColWidth(FDr) = 1700
        .TextMatrix(0, FCr) = "Credit"
        .ColAlignmentFixed(FCr) = flexAlignRightCenter
        .ColAlignment(FCr) = flexAlignRightCenter
        .ColWidth(FCr) = 1700
        .TextMatrix(0, FDueDate) = "DueDate"
        .ColAlignmentFixed(FDueDate) = flexAlignLeftCenter
        .ColAlignment(FDueDate) = flexAlignLeftCenter
        .ColWidth(FDueDate) = 2000
    End With
End Sub
Private Sub FGridClick()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
End Sub
Private Sub FGridRef_Click()
    FGridClick
End Sub
Private Sub FGridRef_DblClick()
    FGridRef_KeyPress vbKeyReturn
    TAddMode = False
End Sub
Private Sub FGridRef_GotFocus()
    If FGridRef.BackColorSel = BackColorSelLeave Then FGridRef.Col = 1
    FGridRef.BackColorSel = BackColorSelEnter
    TxtGrid(0).Visible = False
End Sub
Private Sub FGridRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If TopCtrl1.TopText2.CAPTION = "Browse" And (KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then Form_KeyDown KeyCode, Shift
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGridRef.Tag) = (FGridRef.Rows - (FGridRef.Rows - 1)) Then
        FGridRef.CellBackColor = CellBackColLeave
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGridRef.Tag) = FGridRef.Rows - 1 Then
        FGridRef.CellBackColor = CellBackColLeave
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGridRef.Tag = FGridRef.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Select Case FGridRef.Col
            Case FV_Sno, FDocId, FSubCode, FAgRefNo, FCr, FDr, FDueDate
                FGridRef.TextMatrix(FGridRef.Row, FGridRef.Col) = ""
            Case FAgRefType
                FGridRef.TextMatrix(FGridRef.Row, FV_Sno) = ""
                FGridRef.TextMatrix(FGridRef.Row, FDocId) = ""
                FGridRef.TextMatrix(FGridRef.Row, FSubCode) = ""
                FGridRef.TextMatrix(FGridRef.Row, FAgRefNo) = ""
                FGridRef.TextMatrix(FGridRef.Row, FCr) = ""
                FGridRef.TextMatrix(FGridRef.Row, FDr) = ""
        End Select
    End If
    KeyCode = 0
End Sub
Private Sub FGridRef_KeyPress(KeyAscii As Integer)
    If TopCtrl1.TopText2 = "Browse" Then Exit Sub
    Select Case FGridRef.Col
        Case FAgRefType, FAgRefNo
            FaGet_Text Me, FGridRef, TxtGrid, 0, False, Asc(UCase(Chr(KeyAscii)))
            If Len(TxtGrid(0).TEXT) <= 1 Then TxtGrid(0).TEXT = TxtGrid(0).TEXT
            TxtGrid(0).SelStart = Len(TxtGrid(0).TEXT)
        Case FCr, FDr, FDueDate
            FaGet_Text Me, FGridRef, TxtGrid, 0, True, KeyAscii
    End Select
    If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub
Private Sub FGridRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGridRef.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGridRef.Row >= 1 Then
             If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                If FGridRef.Rows > 2 Then
                    FGridRef.RemoveItem FGridRef.Row
                Else
                    FGridRef.Rows = 1
                    FGridRef.AddItem ""
                    FGridRef.FixedRows = 1
                End If
                CalAmountLF
             End If
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGridRef.SetFocus
    End If
Exit Sub
End Sub
Private Sub FGridRef_Scroll()
    TxtGrid(0).Visible = False
End Sub
Private Sub FGridRef_LostFocus()
    If TxtGrid(0).Visible = False Then FGridRef.BackColorSel = BackColorSelLeave
End Sub
Private Sub TxtGrid_GotFocus(Index As Integer)
Select Case Index
    Case 0
        FGridRef.CellBackColor = CellBackColLeave
        TxtGrid(Index).Tag = FGridRef.TextMatrix(FGridRef.Row, FGridRef.Col)
        Select Case FGridRef.Col
            Case FAgRefType
                TxtGrid(Index).MaxLength = 10
                ListArray = Array("Advance", "Ag.Ref.", "New Ref", "On Account")
                Set mListItem = FaListView_Items(ListView, TxtGrid, Index, ListArray, 4)
                TxtGrid(Index).Tag = TxtGrid(Index).TEXT
            Case FAgRefNo
                TxtGrid(Index).MaxLength = 20
                If PubBackEnd = "A" Then
                    If Val(FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 5)) > 0 Then
                        Set RstRefHelp = G_FaCn.Execute("SELECT AgRefNo AS RefNo,MIN(V_dATE) as VDate,TRIM(CSTR(ABS(IIF(ISNULL(SUM(Cr)),0,SUM(Cr))-IIF(ISNULL(SUM(Dr)),0,SUM(Dr)))))+' '+IIF(IIF(ISNULL(SUM(Cr)),0,SUM(Cr))-IIF(ISNULL(SUM(Dr)),0,SUM(Dr))>0,'Cr','Dr') AS Balance FROM ledgerRef WHERE AGREFTYPE IN ('Advance','New Ref','Ag.Ref.') AND SUBCODE='" & FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 2) & "' GROUP BY SubCode,AgRefNo HAVING IIF(ISNULL(SUM(Cr)),0,SUM(Cr))-IIF(ISNULL(SUM(Dr)),0,SUM(Dr))>0 ORDER BY MIN(V_dATE),AGREFNO")
                    Else
                        Set RstRefHelp = G_FaCn.Execute("SELECT AgRefNo AS RefNo,MIN(V_dATE) as VDate,TRIM(CSTR(ABS(IIF(ISNULL(SUM(Cr)),0,SUM(Cr))-IIF(ISNULL(SUM(Dr)),0,SUM(Dr)))))+' '+IIF(IIF(ISNULL(SUM(Cr)),0,SUM(Cr))-IIF(ISNULL(SUM(Dr)),0,SUM(Dr))>0,'Cr','Dr') AS Balance FROM ledgerRef WHERE AGREFTYPE IN ('Advance','New Ref','Ag.Ref.') AND SUBCODE='" & FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 2) & "' GROUP BY SubCode,AgRefNo HAVING IIF(ISNULL(SUM(Cr)),0,SUM(Cr))-IIF(ISNULL(SUM(Dr)),0,SUM(Dr))<0 ORDER BY MIN(V_dATE),AGREFNO")
                    End If
                ElseIf PubBackEnd = "S" Then
                    If Val(FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 5)) > 0 Then
                        Set RstRefHelp = G_FaCn.Execute("SELECT AgRefNo AS RefNo,MIN(V_dATE) as VDate,TRIM(CSTR(ABS(ISNULL(SUM(Cr),0)-ISNULL(SUM(Dr),0))))+' '+Switch(ISNULL(SUM(Cr),0)-ISNULL(SUM(Dr),0)>0,'Cr',ISNULL(SUM(Cr),0)-ISNULL(SUM(Dr),0)<0,'Dr') AS Balance FROM ledgerRef WHERE AGREFTYPE IN ('Advance','New Ref','Ag.Ref.') AND SUBCODE='" & FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 2) & "' GROUP BY SubCode,AgRefNo HAVING ISNULL(SUM(Cr),0)-ISNULL(SUM(Dr),0)>0 ORDER BY MIN(V_dATE),AGREFNO")
                    Else
                        Set RstRefHelp = G_FaCn.Execute("SELECT AgRefNo AS RefNo,MIN(V_dATE) as VDate,TRIM(CSTR(ABS(ISNULL(SUM(Cr),0)-ISNULL(SUM(Dr),0))))+' '+Switch(ISNULL(SUM(Cr),0)-ISNULL(SUM(Dr),0)>0,'Cr',ISNULL(SUM(Cr),0)-ISNULL(SUM(Dr),0)<0,'Dr') AS Balance FROM ledgerRef WHERE AGREFTYPE IN ('Advance','New Ref','Ag.Ref.') AND SUBCODE='" & FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 2) & "' GROUP BY SubCode,AgRefNo HAVING ISNULL(SUM(Cr),0)-ISNULL(SUM(Dr),0)<0 ORDER BY MIN(V_dATE),AGREFNO")
                    End If
                End If
                Set DGRefNo.DataSource = RstRefHelp
                DGRefNo.left = TxtGrid(Index).left
                DGRefNo.top = TxtGrid(Index).top + TxtGrid(Index).height
                DGRefNo.Tag = Index
                If RstRefHelp.RecordCount = 0 Or (RstRefHelp.EOF = True Or RstRefHelp.BOF = True) Then Exit Sub
            Case FCr, FDr
                TxtGrid(Index).MaxLength = 0
                If TAddMode = False Then SendKeys "{Home}+{End}"
            Case FDueDate
                TxtGrid(Index).MaxLength = 0
                If TAddMode = False Then SendKeys "{Home}+{End}"
            End Select
    End Select
End Sub
Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case 0
        If KeyCode = vbKeyEscape Then
            TxtGrid(Index).TEXT = TxtGrid(Index).Tag
            TxtGrid_KeyUp Index, KeyCode, Shift
            FGridRef.SetFocus
            TxtGrid(Index).Visible = False
            If DGRefNo.Visible = True Then DGRefNo.Visible = False
            Exit Sub
        End If
        Select Case FGridRef.Col
            Case FAgRefType
                FaListView_KeyDown FrmList, ListView, TxtGrid, Index, KeyCode, Shift, TxtGrid(Index).left, (TxtGrid(Index).top + TxtGrid(Index).height + 15), TxtGrid(Index).width, 260 * 5
                If KeyCode = vbKeyReturn Then
                    If TxtGridLeave(Index) = True Then
                         FaGridTxtDown FGridRef, TxtGrid, Index, KeyCode, TAddMode, FDueDate
                    End If
                End If
            Case FAgRefNo
                If FGridRef.TextMatrix(FGridRef.Row, FAgRefType) = "Ag.Ref." Then
                    FaDGridTxtKeyDown DGRefNo, TxtGrid, Index, RstRefHelp, KeyCode, True, 0
                    If KeyCode = vbKeyReturn Then
                        If TxtGridLeave(Index) = True Then
                             FaGridTxtDown FGridRef, TxtGrid, Index, KeyCode, TAddMode, FDueDate
                        End If
                    End If
                Else
                    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                        If TxtGridLeave(Index) = True Then
                             FaGridTxtDown FGridRef, TxtGrid, Index, KeyCode, TAddMode, FDueDate
                        End If
                    End If
                End If
            Case FDr
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave(Index) = True Then
                        If Val(FGridRef.TextMatrix(FGridRef.Row, FGridRef.Col)) > 0 Then
                            FaGridTxtDown FGridRef, TxtGrid, Index, KeyCode, TAddMode, FDueDate, , FDueDate
                        Else
                            FaGridTxtDown FGridRef, TxtGrid, Index, KeyCode, TAddMode, FDueDate, , FDueDate
                        End If
                    End If
                End If
            Case FCr
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave(Index) = True Then
                         FaGridTxtDown FGridRef, TxtGrid, Index, KeyCode, TAddMode, FDueDate, , FDueDate
                         FGridRef.TextMatrix(FGridRef.Row, 0) = ""
                    End If
                End If
            Case FDueDate
                If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                    If TxtGridLeave(Index) = True Then
                        If Val(FGridRef.TextMatrix(FGridRef.Row, FDr)) = 0 And Val(FGridRef.TextMatrix(FGridRef.Row, FCr)) = 0 Then
                            If LblRefAdjDrCrBal = "Cr" Then
                                FGridRef.TextMatrix(FGridRef.Row, FCr) = Format(LblRefAdjBal, "0.00")
                            ElseIf LblRefAdjDrCrBal = "Dr" Then
                                FGridRef.TextMatrix(FGridRef.Row, FDr) = Format(LblRefAdjBal, "0.00")
                            End If
                        End If
                        If FGridRef.TextMatrix(FGridRef.Row, FAgRefNo) = "" And FGridRef.TextMatrix(FGridRef.Row, FAgRefType) <> "On Account" Then
                            FaGridTxtDown FGridRef, TxtGrid, Index, KeyCode, TAddMode, FDueDate, , FAgRefNo
                            Exit Sub
                        End If
                        If Val(LblRefAdjBal) = 0 Then
                            If MsgBox("Save Entries", vbQuestion + vbDefaultButton1 + vbYesNo, "Adjustment") = vbYes Then
                                RefAdjOk
                                Exit Sub
                            Else
                                FGridRef.SetFocus
                            End If
                            FaGridTxtDown FGridRef, TxtGrid, Index, KeyCode, TAddMode, FDueDate
                            FGridRef.TextMatrix(FGridRef.Row, 0) = ""
                        Else
                            FaGridTxtDown FGridRef, TxtGrid, Index, KeyCode, TAddMode, FDueDate
                            FGridRef.TextMatrix(FGridRef.Row, 0) = ""
                            FGridRef.TextMatrix(FGridRef.Row, FAgRefType) = "On Account"
                            FGridRef.TextMatrix(FGridRef.Row, FAgRefNo) = ""
                            If Trim(FGridRef.TextMatrix(FGridRef.Row, FDueDate)) = "" Then FGridRef.TextMatrix(FGridRef.Row, FDueDate) = VchDt(0)
                            If LblRefAdjDrCrBal = "Dr" Then
                                If Trim(FGridRef.TextMatrix(FGridRef.Row, FCr)) = "" And Trim(FGridRef.TextMatrix(FGridRef.Row, FDr)) = "" Then FGridRef.TextMatrix(FGridRef.Row, FDr) = Format(Val(LblRefAdjBal), "0.00")
                            ElseIf LblRefAdjDrCrBal = "Cr" Then
                                If Trim(FGridRef.TextMatrix(FGridRef.Row, FCr)) = "" And Trim(FGridRef.TextMatrix(FGridRef.Row, FDr)) = "" Then FGridRef.TextMatrix(FGridRef.Row, FCr) = Format(Val(LblRefAdjBal), "0.00")
                            End If
                            CalAmountLF
                        End If
                    End If
                End If
        End Select
End Select
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
FaCheckQuote KeyAscii
Select Case Index
    Case 0
        Case FAgRefNo
            If FGridRef.TextMatrix(FGridRef.Row, FAgRefType) = "Ag.Ref." Then
                If DGRefNo.Visible = True Then FaDGridTxtKeyPress TxtGrid, Index, RstRefHelp, KeyAscii, "RefNo"
            End If
        Case FCr, FDr
            FaNumPress TxtGrid(Index), KeyAscii, 8, 2
End Select
End Sub
Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case 0
        Select Case FGridRef.Col
            Case FAgRefType
                If KeyCode <> 13 And FrmList.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0
                FaListView_KeyUp ListView, TxtGrid, Index, KeyCode, mListItem
            Case FAgRefNo
                If FGridRef.TextMatrix(FGridRef.Row, FAgRefType) = "Ag.Ref." Then
                    If KeyCode <> 13 And DGRefNo.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: FaDGridTxtKeyPress TxtGrid, Index, RstRefHelp, KeyCode, "RefNo", True
                End If
        End Select
End Select
Exit Sub
ELoop:   MsgBox err.Description, vbCritical
End Sub
Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            Cancel = Not TxtGridLeave(Index)
    End Select
End Sub
Private Sub CalAmountLF()
Dim I As Integer, mTotal As Double, mBALANCE As Double
mTotal = 0
mBALANCE = 0
For I = 1 To FGridRef.Rows - 1
    mTotal = mTotal + Val(FGridRef.TextMatrix(I, FCr)) - Val(FGridRef.TextMatrix(I, FDr))
Next
LblRefAdj = Format(Abs(mTotal), "0.00")
LblRefAdjDrCr = IIf(mTotal = 0, "", IIf(mTotal > 0, "Cr", "Dr"))
mBALANCE = IIf(LblRefAmtDrCr = "Cr", Val(LblRefAmt), -Val(LblRefAmt)) - (mTotal)
LblRefAdjBal = Format(Abs(mBALANCE), "0.00")
LblRefAdjDrCrBal = IIf(mBALANCE > 0, "Cr", "Dr")
End Sub
Private Function TxtGridLeave(Optional Index As Integer) As Boolean
CalAmountLF
Select Case Index
    Case 0
        Select Case FGridRef.Col
            Case FAgRefType
                FGridRef.TextMatrix(FGridRef.Row, FAgRefType) = TxtGrid(Index).TEXT
                If FGridRef.TextMatrix(FGridRef.Row, FAgRefType) <> "Advance" And FGridRef.TextMatrix(FGridRef.Row, FAgRefType) <> "Ag.Ref." And FGridRef.TextMatrix(FGridRef.Row, FAgRefType) <> "New Ref" Then FGridRef.TextMatrix(FGridRef.Row, FAgRefType) = "On Account"
                If Trim(FGridRef.TextMatrix(FGridRef.Row, FAgRefNo)) = "" And FGridRef.TextMatrix(FGridRef.Row, FAgRefType) <> "Ag.Ref." And FGridRef.TextMatrix(FGridRef.Row, FAgRefType) <> "On Account" Then FGridRef.TextMatrix(FGridRef.Row, FAgRefNo) = Trim(TxtVtYpe(0).Tag) + "-" + Trim(TxtVno(0))
                If Trim(FGridRef.TextMatrix(FGridRef.Row, FDueDate)) = "" Then FGridRef.TextMatrix(FGridRef.Row, FDueDate) = VchDt(0)
                If LblRefAdjDrCrBal = "Dr" Then
                    If Trim(FGridRef.TextMatrix(FGridRef.Row, FCr)) = "" And Trim(FGridRef.TextMatrix(FGridRef.Row, FDr)) = "" Then FGridRef.TextMatrix(FGridRef.Row, FDr) = Format(Val(LblRefAdjBal), "0.00")
                ElseIf LblRefAdjDrCrBal = "Cr" Then
                    If Trim(FGridRef.TextMatrix(FGridRef.Row, FCr)) = "" And Trim(FGridRef.TextMatrix(FGridRef.Row, FDr)) = "" Then FGridRef.TextMatrix(FGridRef.Row, FCr) = Format(Val(LblRefAdjBal), "0.00")
                End If
            Case FAgRefNo
                If FGridRef.TextMatrix(FGridRef.Row, FAgRefType) = "Ag.Ref." Then
                   If RstRefHelp.RecordCount = 0 Or (RstRefHelp.EOF = True Or RstRefHelp.BOF = True) Or TxtGrid(Index).TEXT = "" Then
                        FGridRef.TextMatrix(FGridRef.Row, FAgRefNo) = ""
                    Else
                        FGridRef.TextMatrix(FGridRef.Row, FAgRefNo) = RstRefHelp!RefNo
                    End If
                    DGRefNo.Visible = False
                Else
                    FGridRef.TextMatrix(FGridRef.Row, FAgRefNo) = TxtGrid(Index).TEXT
                End If
            Case FCr
                FGridRef.TextMatrix(FGridRef.Row, FCr) = Format(TxtGrid(Index).TEXT, "0.00")
                If Val(FGridRef.TextMatrix(FGridRef.Row, FCr)) > 0 Then FGridRef.TextMatrix(FGridRef.Row, FDr) = "0.00"
            Case FDr
                FGridRef.TextMatrix(FGridRef.Row, FDr) = Format(TxtGrid(Index).TEXT, "0.00")
                If Val(FGridRef.TextMatrix(FGridRef.Row, FDr)) > 0 Then FGridRef.TextMatrix(FGridRef.Row, FCr) = "0.00"
            Case FDueDate
                FGridRef.TextMatrix(FGridRef.Row, FDueDate) = PubDatamanFa.FaRetDateFunc(TxtGrid(Index))
        End Select
End Select
CalAmountLF
TxtGridLeave = True
End Function
Private Sub ListView_Click()
    TxtGrid(0).TEXT = ListView.SelectedItem.TEXT
    FrmList.Visible = False
    If TxtGridLeave(0) = True Then
        FaGridTxtDown FGridRef, TxtGrid, 0, vbKeyReturn, TAddMode, FDueDate
    End If
    FGridRef.SetFocus
End Sub
Private Sub DGRefNo_Click()
    If RstRefHelp.RecordCount > 0 Then
        TxtGrid(0).TEXT = RstRefHelp!RefNo
    End If
    FGridRef.ZOrder 0
    TxtGrid(0).SetFocus
    DGRefNo.Visible = False
    If TxtGridLeave(0) = True Then
        FaGridTxtDown FGridRef, TxtGrid, 0, vbKeyReturn, TAddMode, FDueDate
    End If
End Sub
Private Sub RefAdjOk()
Dim I As Integer
    FrameRef.Visible = False
    VchTrnSaveLst Val(FrameRef.Tag)
    For I = 1 To FGridRef.Rows - 1
        If Val(FGridRef.TextMatrix(I, FCr)) + Val(FGridRef.TextMatrix(I, FDr)) > 0 Then
            If FGridRef.TextMatrix(I, FAgRefType) <> "On Account" Then
                If Trim(FGridRef.TextMatrix(I, FAgRefNo)) = "" Then
                    MsgBox "Ref.No.Needed"
                    Exit Sub
                End If
            End If
        End If
    Next
    If RstRef.RecordCount > 0 Then
        RstRef.Sort = "V_SNo ASC"
        RstRef.Find "V_SNO=" & Val(FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 1))
        If RstRef.EOF = False Then
            Do While RstRef!V_SNo = Val(FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 1))
                RstRef.Delete
                RstRef.MoveNext
                If RstRef.EOF = True Then Exit Do
            Loop
        End If
    End If
    For I = 1 To FGridRef.Rows - 1
        If Val(FGridRef.TextMatrix(I, FCr)) > 0 Or Val(FGridRef.TextMatrix(I, FDr)) > 0 Then
            With RstRef
                .AddNew
                .Fields("V_SNo") = Val(FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 1))
                .Fields("DocId") = FGridRef.TextMatrix(I, FDocId)
                .Fields("SUBCODE") = FGrid1.TextMatrix(Val(FrameRef.Tag) + ScrolIndex, 2)
                .Fields("AgRefType") = FGridRef.TextMatrix(I, FAgRefType)
                .Fields("AgRefNo") = FGridRef.TextMatrix(I, FAgRefNo)
                .Fields("CR") = Val(FGridRef.TextMatrix(I, FCr))
                .Fields("DR") = Val(FGridRef.TextMatrix(I, FDr))
                .Fields("DueDate") = FGridRef.TextMatrix(I, FDueDate)
                .Update
            End With
        End If
    Next
    If TxtCrDr(Val(FrameRef.Tag)) = "Cr" Then
        TxtCr(Val(FrameRef.Tag)).SetFocus
    Else
        TxtDr(Val(FrameRef.Tag)).SetFocus
    End If
End Sub
Private Sub BtnRefAdjOK_Click()
    If MsgBox("Save Entries", vbQuestion + vbDefaultButton1 + vbYesNo, "Adjustment") = vbYes Then
        RefAdjOk
        Exit Sub
    Else
        FGridRef.SetFocus
    End If
End Sub
Private Sub MakeEmpty()
Dim I As Integer
ScrolIndex = 0
FGrid1.Rows = 0
TxtVtYpe(0) = ""
TxtVtYpe(0).Tag = ""
TxtVno(0) = ""
LblVPrefix = ""
For I = 0 To TxtAcName.Count - 1
    TxtAcName(I).TEXT = ""
    TxtAcName(I).Tag = ""
    TxtCrDr(I).TEXT = ""
    TxtCrDr(I).Tag = ""
    TxtCr(I).TEXT = ""
    TxtDr(I).TEXT = ""
    TxtNar(I).TEXT = ""
    LblCb(I) = ""
    LblNar(I) = ""
Next
TxtGlb(0).TEXT = ""
TXTChDate(0) = ""
TxtCHno(0) = ""
TXTClrDate(0) = ""
LblDrAmt(0).CAPTION = ""
LblCrAmt(0).CAPTION = ""
End Sub
Private Sub MakeVisible(mRowToVisible As Integer, mLock As Boolean)
TxtCrDr(mRowToVisible).Visible = mLock
TxtAcName(mRowToVisible).Visible = mLock
TxtCr(mRowToVisible).Visible = mLock
TxtDr(mRowToVisible).Visible = mLock
TxtNar(mRowToVisible).Visible = mLock
If mLock = False Then
    LblCb(mRowToVisible) = ""
    TxtCrDr(mRowToVisible) = ""
    TxtAcName(mRowToVisible) = ""
    TxtCr(mRowToVisible) = ""
    TxtDr(mRowToVisible) = ""
    TxtNar(mRowToVisible) = ""
Else
    If TxtCrDr(mRowToVisible) = "Cr" Then TxtDr(mRowToVisible).Visible = False
    If TxtCrDr(mRowToVisible) = "Dr" Then TxtCr(mRowToVisible).Visible = False
End If
If mSepNar = "Y" And mLock = True Then
    LblNar(mRowToVisible).Visible = True
    LblNar(mRowToVisible) = "Narration"
    TxtNar(mRowToVisible).Visible = mLock
Else
    LblNar(mRowToVisible) = ""
    LblNar(mRowToVisible).Visible = False
    TxtNar(mRowToVisible).Visible = False
End If
If RstEnviro!ShowCurrentBalance = "Yes" Then
    LblCb(mRowToVisible).Visible = mLock
Else
    LblCb(mRowToVisible).Visible = False
End If
End Sub
Private Sub LockFields(mLock As Boolean)
Dim I As Integer
For I = 0 To TxtAcName.Count - 1
    TxtAcName(I).Locked = mLock
    TxtCrDr(I).Locked = mLock
    TxtCr(I).Locked = mLock
    TxtDr(I).Locked = mLock
    TxtNar(I).Locked = mLock
Next
TxtVtYpe(0).Locked = mLock
TxtVno(0).Locked = mLock
VchDt(0).Locked = mLock
TxtGlb(0).Locked = mLock
TXTChDate(0).Locked = mLock
TxtCHno(0).Locked = mLock
TXTClrDate(0).Locked = mLock
If TxtSite.Visible = True Then
    If PubFaSiteType = 1 And PubSeparateVrNoForSite = 1 Then
        TxtSite.Visible = True
        TxtSite.Enabled = False
    Else
        TxtSite.Enabled = Not mLock
    End If
End If
End Sub
Private Sub MakeVrTypeVisible(Rst As ADODB.Recordset)
Dim I As Integer, mTop As Integer, RST1 As ADODB.Recordset
If FaXNull(Rst!DefaultCrAC) <> "" Then
    mDefaultCrAc = FaXNull(Rst!DefaultCrAC)
Else
    mDefaultCrAc = ""
End If
If FaXNull(Rst!DefaultDrAC) <> "" Then
    mDefaultDrAc = FaXNull(Rst!DefaultDrAC)
Else
    mDefaultDrAc = ""
End If
If FaXNull(Rst!FirstDrCr) <> "" Then
    If Trim(TxtAcName(0)) = "" Then TxtCrDr(0) = FaXNull(Rst!FirstDrCr)
End If
mNCat = Rst!NCat
mLastVrType = Rst!V_tYPE
mSepNar = FaXNull(Rst!Separate_Narr)
mCommNar = FaXNull(Rst!Common_Narr)
TxtVtYpe(0) = Rst!Description
TxtVtYpe(0).Tag = Rst!V_tYPE
LblVPrefix = FaXNull(Rst!prefix)
TxtVno(0).Enabled = IIf(Rst!Number_Method = "Automatic", False, True)
If ADDFLAG = 1 Then
    If Rst!Number_Method = "Automatic" Then
        TxtVno(0).TEXT = Rst!start_srl_no + 1
    ElseIf Rst!Number_Method = "SemiAuto" Then
        Set RST1 = G_FaCn.Execute("SELECT MAX(V_NO) as vno from ledger WHERE V_TYPE='" & Rst!V_tYPE & "' AND V_PREFIX='" & LblVPrefix & "'")
        If RST1.RecordCount > 0 Then
            TxtVno(0).TEXT = FaVNull(RST1!VNo) + 1
        Else
            TxtVno(0).TEXT = 1
        End If
    End If
End If
TxtGlb(0).Visible = IIf(Rst!Common_Narr = "Y", True, False)
Label1(3).Visible = IIf(Rst!Common_Narr = "Y", True, False)
Label6.Visible = IIf(Rst!ChqNo = "Y", True, False)
TxtCHno(0).Visible = IIf(Rst!ChqNo = "Y", True, False)
Label5.Visible = IIf(Rst!ChqDT = "Y", True, False)
TXTChDate(0).Visible = IIf(Rst!ChqDT = "Y", True, False)
Label7.Visible = IIf(Rst!CLGDT = "Y", True, False)
TXTClrDate(0).Visible = IIf(Rst!CLGDT = "Y", True, False)
mTop = 1020
If RstEnviro!ShowCurrentBalance = "Yes" And mSepNar = "Y" Then
    FixRow = 4
    If ADDFLAG <= 2 Then
        If TypeOf Me.ActiveControl Is TextBox Then If Me.ActiveControl.Index > FixRow Then TxtCrDr(0).SetFocus
    End If
ElseIf (RstEnviro!ShowCurrentBalance = "Yes" And mSepNar = "N") Or (RstEnviro!ShowCurrentBalance = "No" And mSepNar = "Y") Then
    FixRow = 6
    If ADDFLAG <= 2 Then
        If TypeOf Me.ActiveControl Is TextBox Then If Me.ActiveControl.Index > FixRow Then TxtCrDr(0).SetFocus
    End If
Else
    FixRow = 11
End If
'''''Color Changing & Positioning
For I = 1 To 4
    LblShort(I).FontBold = False
    LblShort(I).ForeColor = &HFFFF&
Next
Select Case mNCat
    Case "CNT"
        Me.BackColor = &HC9DBB3
        CtrlBColOrg = &HC9DBB3
        CtrlFColOrg = &H80000012
        LblShort(1).ForeColor = &HFFFF00
        LblShort(1).FontBold = True
        PicDN.BackColor = Me.BackColor
    Case "PMT"
        Me.BackColor = &HC8D5EC
        CtrlBColOrg = &HC8D5EC
        CtrlFColOrg = &H80000012
        LblShort(2).ForeColor = &HFFFF00
        LblShort(2).FontBold = True
        PicDN.BackColor = Me.BackColor
    Case "RCT"
        Me.BackColor = &HDCD5BE
        CtrlBColOrg = &HDCD5BE
        CtrlFColOrg = &H80000012
        LblShort(3).ForeColor = &HFFFF00
        LblShort(3).FontBold = True
        PicDN.BackColor = Me.BackColor
    Case "JV"
        Me.BackColor = &HD0E6E4
        CtrlBColOrg = &HD0E6E4
        CtrlFColOrg = &H80000012
        LblShort(4).ForeColor = &HFFFF00
        LblShort(4).FontBold = True
        PicDN.BackColor = Me.BackColor
End Select
For I = 0 To 11
    TxtCrDr(I).top = mTop
    TxtCrDr(I).left = 15
    TxtAcName(I).top = mTop
    TxtAcName(I).left = 330
    TxtCr(I).top = mTop
    TxtCr(I).left = 9015
    TxtDr(I).top = mTop
    TxtDr(I).left = 7620
    If RstEnviro!ShowCurrentBalance = "Yes" Then
        mTop = mTop + 240
        LblCb(I).top = mTop
        LblCb(I).left = 15
    End If
    If mSepNar = "Y" Then
        mTop = mTop + 240
        LblNar(I).top = mTop
        LblNar(I).left = 15
        TxtNar(I).top = mTop
        TxtNar(I).left = 990
    End If
    mTop = mTop + 260
    If FixRow >= I And TxtCrDr(I) <> "" Then
        MakeVisible I, True
    Else
        MakeVisible I, False
    End If
    TxtAcName(I).BackColor = CtrlBColOrg
    TxtAcName(I).ForeColor = CtrlFColOrg
    TxtCr(I).BackColor = CtrlBColOrg
    TxtCr(I).ForeColor = CtrlFColOrg
    TxtCrDr(I).BackColor = CtrlBColOrg
    TxtCrDr(I).ForeColor = CtrlFColOrg
    TxtDr(I).BackColor = CtrlBColOrg
    TxtDr(I).ForeColor = CtrlFColOrg
    TxtNar(I).BackColor = CtrlBColOrg
    TxtNar(I).ForeColor = CtrlFColOrg
Next
TxtGlb(0).BackColor = CtrlBColOrg
TxtGlb(0).ForeColor = CtrlFColOrg
TxtVtYpe(0).BackColor = CtrlBColOrg
TxtVtYpe(0).ForeColor = CtrlFColOrg
TxtVno(0).ForeColor = CtrlFColOrg
TxtVno(0).BackColor = CtrlBColOrg
VchDt(0).BackColor = CtrlBColOrg
VchDt(0).ForeColor = CtrlFColOrg
TxtCHno(0).BackColor = CtrlBColOrg
TxtCHno(0).ForeColor = CtrlFColOrg
TXTChDate(0).BackColor = CtrlBColOrg
TXTChDate(0).ForeColor = CtrlFColOrg
TXTClrDate(0).BackColor = CtrlBColOrg
TXTClrDate(0).ForeColor = CtrlFColOrg
Set RST1 = Nothing
End Sub
Private Sub FaVrTypeSetting(Optional mVType As String, Optional mPrefix As String)
Dim MyRs As New ADODB.Recordset, mCondStr As String, mDate As Date, mSiteHlp As String
mSiteHlp = ""
If PubSiteCodeWiseHelp = True Then
    mSiteHlp = "And Site_Code='" & PubSiteCode & "'"
End If
mDate = IIf(VchDt(0) <> "", VchDt(0), PubLoginDate)

If (PubFaSiteType = 1 And PubSeparateVrNoForSite = 1) Then
    If PubBackEnd = "A" Then
        mCondStr = " AND Date_From<=" & FaConvertDate(mDate) & " AND IIF(ISNULL(DATE_TO)," & FaConvertDate(mDate) & ",DATE_TO)>=" & FaConvertDate(mDate) & " AND SITE_CODE='" & PubSeparateLogSite & "'"
    ElseIf PubBackEnd = "S" Then
        mCondStr = " AND Date_From<=" & FaConvertDate(mDate) & " AND ISNULL(DATE_TO," & FaConvertDate(mDate) & " )>=" & FaConvertDate(mDate) & " AND SITE_CODE='" & PubSeparateLogSite & "'"
    End If
ElseIf PubFaSiteType = 2 Then
    If PubBackEnd = "A" Then
        mCondStr = " AND Date_From<=" & FaConvertDate(mDate) & " AND IIF(ISNULL(DATE_TO)," & FaConvertDate(mDate) & ",DATE_TO)>=" & FaConvertDate(mDate) & " AND SITE_CODE='" & PubSiteCode & "'"
    ElseIf PubBackEnd = "S" Then
        mCondStr = " AND Date_From<=" & FaConvertDate(mDate) & " AND ISNULL(DATE_TO," & FaConvertDate(mDate) & " )>=" & FaConvertDate(mDate) & " AND SITE_CODE='" & PubSiteCode & "'"
    End If
Else
    If PubBackEnd = "A" Then
        mCondStr = " AND Date_From<=" & FaConvertDate(mDate) & " AND IIF(ISNULL(DATE_TO)," & FaConvertDate(mDate) & ",DATE_TO)>=" & FaConvertDate(mDate)
    ElseIf PubBackEnd = "S" Then
        mCondStr = " AND Date_From<=" & FaConvertDate(mDate) & " AND ISNULL(DATE_TO," & FaConvertDate(mDate) & " )>=" & FaConvertDate(mDate)
    End If
End If

'If mVType <> "" Then mCondStr = mCondStr + " AND VOUCHER_Type.V_tYPE='" & mVType & "' AND VOUCHER_Prefix.V_tYPE='" & mVType & "' "
'If Not IsMissing(mPrefix) And mPrefix <> "" Then mCondStr = mCondStr + " AND VOUCHER_PREFIX.PREFIX='" & mPrefix & "'"
Set RstVchrHlp = G_FaCn.Execute("Select Voucher_Type.V_Type,Voucher_Type.DESCRIPTION,Voucher_Prefix.Prefix,Voucher_Prefix.Date_From,Voucher_Prefix.Start_Srl_No,Number_Method,Separate_Narr,Common_Narr,NCAT,ChqNo,ChqDt,ClgDt,LTRIM(RTRIM(Voucher_Type.V_Type))+'-'+LTRIM(RTRIM(Voucher_Prefix.Prefix)) as SearchCode,Voucher_Type.DefaultCrAc,Voucher_Type.DefaultDrAC,Voucher_Type.FirstDrCr From Voucher_Prefix right join Voucher_Type on Voucher_Type.V_Type=Voucher_Prefix.V_Type Where Voucher_Type.NCAT='" & mNCat & "' " & mCondStr & " order by Description,Prefix ")
Set DGVchrHlp.DataSource = Nothing
Set DGVchrHlp.DataSource = RstVchrHlp
If RstVchrHlp.RecordCount > 0 Then
    RstVchrHlp.MoveFirst
    If RstVchrHlp.RecordCount = 1 Then
        TxtVtYpe(0).Enabled = False
    Else
        If mVType <> "" And mPrefix <> "" Then
            RstVchrHlp.MoveFirst
            Do Until RstVchrHlp.EOF
                If RstVchrHlp!SearchCode = Trim(mVType) + "-" + Trim(mPrefix) Then Exit Do
                RstVchrHlp.MoveNext
            Loop
        ElseIf mVType <> "" Then
            RstVchrHlp.MoveFirst
            Do Until RstVchrHlp.EOF
                If RstVchrHlp!V_tYPE = Trim(mVType) Then Exit Do
                RstVchrHlp.MoveNext
            Loop
        End If
        If RstVchrHlp.EOF = True Then RstVchrHlp.MoveFirst
        TxtVtYpe(0).Enabled = True
    End If
    MakeVrTypeVisible RstVchrHlp
End If
If PubBackEnd = "A" Then
    If RstEnviro!FilterAC = "Yes" Then
        Set RstAcHlpDr = New ADODB.Recordset
        If RstEnviro!ShowCityName = "Yes" Then
            RstAcHlpDr.Open ("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3))+RTRIM(IIF(ISNULL(CityName),'',','+CityName)) AS NameWithADDR,IIF(ISNULL(FNAME),'',FNAME) AS FatherName  From ViewSubgroup WHERE GroupCode NOT IN (SELECT GROUPCODE FROM Voucher_Exclude WHERE DR='Y' AND V_tYPE=" & FaChk_Text(TxtVtYpe(0).Tag) & ") " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
        Else
            RstAcHlpDr.Open ("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3))+RTRIM(IIF(ISNULL(CityName),'',','+CityName)) AS NameWithADDR,IIF(ISNULL(FNAME),'',FNAME) AS FatherName  From ViewSubgroup WHERE GroupCode NOT IN (SELECT GROUPCODE FROM Voucher_Exclude WHERE DR='Y' AND V_tYPE=" & FaChk_Text(TxtVtYpe(0).Tag) & ") " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
        End If
        Set RstAcHlpCr = New ADODB.Recordset
        If RstEnviro!ShowCityName = "Yes" Then
            RstAcHlpCr.Open ("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3))+RTRIM(IIF(ISNULL(CityName),'',','+CityName)) AS NameWithADDR,IIF(ISNULL(FNAME),'',FNAME) AS FatherName From ViewSubgroup WHERE GroupCode NOT IN (SELECT GROUPCODE FROM Voucher_Exclude WHERE CR='Y' AND V_tYPE=" & FaChk_Text(TxtVtYpe(0).Tag) & ") " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
        Else
            RstAcHlpCr.Open ("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(IIF(ISNULL(ADD1) OR LEN(ADD1)=0,'',ADD1))+RTRIM(IIF(ISNULL(ADD2) OR LEN(ADD2)=0,'',','+ADD2))+RTRIM(IIF(ISNULL(ADD3) OR LEN(ADD3)=0,'',','+ADD3))+RTRIM(IIF(ISNULL(CityName),'',','+CityName)) AS NameWithADDR,IIF(ISNULL(FNAME),'',FNAME) AS FatherName From ViewSubgroup WHERE GroupCode NOT IN (SELECT GROUPCODE FROM Voucher_Exclude WHERE CR='Y' AND V_tYPE=" & FaChk_Text(TxtVtYpe(0).Tag) & ") " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
        End If
    End If
Else
    If RstEnviro!FilterAC = "Yes" Then
        Set RstAcHlpDr = New ADODB.Recordset
        If RstEnviro!ShowCityName = "Yes" Then
            RstAcHlpDr.Open ("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,''))+RTRIM(ISNULL(CityName,'')) AS NameWithADDR,ISNULL(FNAME,'') AS FatherName From ViewSubgroup WHERE GroupCode NOT IN (SELECT GROUPCODE FROM Voucher_Exclude WHERE DR='Y' AND V_tYPE=" & FaChk_Text(TxtVtYpe(0).Tag) & ") " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
        Else
            RstAcHlpDr.Open ("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,''))+RTRIM(ISNULL(CityName,'')) AS NameWithADDR,ISNULL(FNAME,'') AS FatherName From ViewSubgroup WHERE GroupCode NOT IN (SELECT GROUPCODE FROM Voucher_Exclude WHERE DR='Y' AND V_tYPE=" & FaChk_Text(TxtVtYpe(0).Tag) & ") " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
        End If
        Set RstAcHlpCr = New ADODB.Recordset
        If RstEnviro!ShowCityName = "Yes" Then
            RstAcHlpCr.Open ("Select SubCode,NameWithCity AS Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,''))+RTRIM(ISNULL(CityName,'')) AS NameWithADDR,ISNULL(FNAME,'') AS FatherName From ViewSubgroup WHERE GroupCode NOT IN (SELECT GROUPCODE FROM Voucher_Exclude WHERE CR='Y' AND V_tYPE=" & FaChk_Text(TxtVtYpe(0).Tag) & ") " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
        Else
            RstAcHlpCr.Open ("Select SubCode,Name,GNAME AS GROUPNAME,GroupCode,MAINGRCODES, RTRIM(ISNULL(ADD1,''))+','+RTRIM(ISNULL(ADD2,''))+RTRIM(ISNULL(ADD3,''))+RTRIM(ISNULL(CityName,'')) AS NameWithADDR From ViewSubgroup,ISNULL(FNAME,'') AS FatherName WHERE GroupCode NOT IN (SELECT GROUPCODE FROM Voucher_Exclude WHERE CR='Y' AND V_tYPE=" & FaChk_Text(TxtVtYpe(0).Tag) & ") " & mSiteHlp & " order by Name"), G_FaCn, adOpenKeyset, adLockOptimistic
        End If
    End If
End If
Set MyRs = Nothing
End Sub

'Fgrid1
'0|1 V_SNo |2 SubCode |3 Name |4 "" |5 AmtDr |6 AmtCr |7 AmtCr |8 Narration |9 Curr_Bal

'FgridAdjust
'0|1 V.Type|2 V.No|3 Vr.Prefix|4 Sr.No|5 Date |6 A/C Name|7 Vr.Amt|8 Pend.Adj.|9 THIS Vr. Adjustments|10 NARRATION |11  |12 DocId

'FV_Sno As Byte = 1, FDocID As Byte = 2, FVSNo As Byte = 3, FSubCode As Byte = 4
'FAgRefType As Byte = 5, FAgRefNo As Byte = 6, FCr As Byte = 7, FDr As Byte = 8

