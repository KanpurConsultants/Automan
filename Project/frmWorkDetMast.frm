VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TopCtl.ocx"
Begin VB.Form frmWorkDetMast 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Workshop Details Master"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11370
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
   ScaleHeight     =   8010
   ScaleWidth      =   11370
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   64
      Left            =   4065
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1845
      Width           =   1965
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   63
      Left            =   5625
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1575
      Width           =   405
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   62
      Left            =   4065
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1575
      Width           =   405
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   1605
      TabIndex        =   102
      Top             =   3615
      Visible         =   0   'False
      Width           =   2010
      Begin MSComctlLib.ListView ListView 
         Height          =   1815
         Left            =   0
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   0
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   3201
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   61
      Left            =   4350
      MaxLength       =   10
      TabIndex        =   101
      Top             =   765
      Width           =   1995
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   60
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   63
      Top             =   7185
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   59
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   62
      Top             =   7185
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   58
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   61
      Top             =   7185
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   57
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   60
      Top             =   6915
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   56
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   59
      Top             =   6915
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   55
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   58
      Top             =   6915
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   54
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   57
      Top             =   6645
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   53
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   56
      Top             =   6645
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   52
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   55
      Top             =   6645
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   51
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   54
      Top             =   6375
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   50
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   53
      Top             =   6375
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   49
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   52
      Top             =   6375
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   48
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   51
      Top             =   6105
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   47
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   50
      Top             =   6105
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   46
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   49
      Top             =   6105
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   45
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   48
      Top             =   5835
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   44
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   47
      Top             =   5835
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   43
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   46
      Top             =   5835
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   42
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   45
      Top             =   5565
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   41
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   44
      Top             =   5565
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   40
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   43
      Top             =   5565
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   39
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   42
      Top             =   5295
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   38
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   41
      Top             =   5295
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   37
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   40
      Top             =   5295
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   36
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   39
      Top             =   5025
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   35
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   38
      Top             =   5025
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   34
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   37
      Top             =   5025
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   33
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   36
      Top             =   4755
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   32
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   35
      Top             =   4755
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   31
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   34
      Top             =   4755
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   30
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   33
      Top             =   4485
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   29
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   32
      Top             =   4485
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   28
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   31
      Top             =   4485
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   27
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   30
      Top             =   4215
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   26
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   29
      Top             =   4215
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   25
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   28
      Top             =   4215
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   24
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   27
      Top             =   3945
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   23
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   26
      Top             =   3945
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   22
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   25
      Top             =   3945
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   21
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   24
      Top             =   3675
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   20
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   23
      Top             =   3675
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   19
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   22
      Top             =   3675
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   18
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   21
      Top             =   3405
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   17
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   20
      Top             =   3405
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   16
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   19
      Top             =   3405
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   15
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   18
      Top             =   3135
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   14
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   17
      Top             =   3135
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   13
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   16
      Top             =   3135
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   12
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   15
      Top             =   2865
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   11
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   14
      Top             =   2865
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   10
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   13
      Top             =   2865
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   9
      Left            =   9165
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2595
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   240
      Index           =   8
      Left            =   7140
      MaxLength       =   2
      TabIndex        =   11
      Top             =   2595
      Width           =   360
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   7
      Left            =   5115
      MaxLength       =   2
      TabIndex        =   10
      Top             =   2595
      Width           =   360
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   661
      tAdd            =   0   'False
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   9540
      MaxLength       =   3
      TabIndex        =   6
      Top             =   705
      Width           =   525
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   4065
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1035
      Width           =   765
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   6
      Left            =   9540
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1515
      Width           =   525
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   5
      Left            =   9540
      MaxLength       =   3
      TabIndex        =   8
      Top             =   1245
      Width           =   525
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   4
      Left            =   9540
      MaxLength       =   3
      TabIndex        =   7
      Top             =   975
      Width           =   525
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   4065
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1305
      Width           =   405
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   0
      Left            =   4065
      MaxLength       =   1
      TabIndex        =   64
      Top             =   765
      Width           =   270
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
      Height          =   210
      Index           =   3
      Left            =   3930
      TabIndex        =   109
      Top             =   1860
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Workshop Running Cost"
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
      Index           =   25
      Left            =   1515
      TabIndex        =   108
      Top             =   1860
      Width           =   2025
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total ATMs"
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
      Height          =   210
      Index           =   24
      Left            =   4500
      TabIndex        =   107
      Top             =   1590
      Width           =   900
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
      Height          =   210
      Index           =   2
      Left            =   5490
      TabIndex        =   106
      Top             =   1605
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Working Hours Per Day"
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
      Height          =   210
      Index           =   23
      Left            =   1515
      TabIndex        =   105
      Top             =   1590
      Width           =   2370
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
      Height          =   210
      Index           =   1
      Left            =   3930
      TabIndex        =   104
      Top             =   1590
      Width           =   45
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skilled"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Index           =   22
      Left            =   4965
      TabIndex        =   100
      Top             =   2235
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Semi - Skilled"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Index           =   21
      Left            =   6660
      TabIndex        =   99
      Top             =   2250
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Un - Skilled"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Index           =   20
      Left            =   8775
      TabIndex        =   98
      Top             =   2265
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Index           =   19
      Left            =   2280
      TabIndex        =   97
      Top             =   2220
      Width           =   1035
   End
   Begin VB.Line Line2 
      X1              =   1470
      X2              =   10350
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   8310
      X2              =   8325
      Y1              =   2160
      Y2              =   7515
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   6300
      X2              =   6300
      Y1              =   2145
      Y2              =   7530
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4260
      X2              =   4260
      Y1              =   2145
      Y2              =   7500
   End
   Begin VB.Shape Shape1 
      Height          =   5370
      Left            =   1485
      Top             =   2160
      Width           =   8880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guards"
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
      Left            =   1830
      TabIndex        =   96
      Top             =   7245
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Advisor"
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
      Left            =   1830
      TabIndex        =   95
      Top             =   6945
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receptionist"
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
      Left            =   1830
      TabIndex        =   94
      Top             =   6675
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Computer Operator"
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
      Index           =   15
      Left            =   1830
      TabIndex        =   93
      Top             =   6405
      Width           =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Machinist"
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
      Index           =   14
      Left            =   1830
      TabIndex        =   92
      Top             =   5310
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Typist"
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
      Index           =   13
      Left            =   1830
      TabIndex        =   91
      Top             =   6135
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clerks Staff"
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
      Index           =   12
      Left            =   1830
      TabIndex        =   90
      Top             =   5865
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cashier"
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
      Index           =   11
      Left            =   1830
      TabIndex        =   89
      Top             =   5595
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apprentices"
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
      Left            =   1830
      TabIndex        =   88
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Electrician"
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
      Left            =   1830
      TabIndex        =   87
      Top             =   4500
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Painters"
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
      Index           =   8
      Left            =   1830
      TabIndex        =   86
      Top             =   4770
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asst. Works Manager"
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
      Left            =   1830
      TabIndex        =   85
      Top             =   2850
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Works Manager"
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
      Index           =   6
      Left            =   1830
      TabIndex        =   84
      Top             =   2580
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervisor"
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
      Left            =   1830
      TabIndex        =   83
      Top             =   3120
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asst. Mechanic"
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
      Left            =   1830
      TabIndex        =   82
      Top             =   3675
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Helper"
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
      Left            =   1830
      TabIndex        =   81
      Top             =   3945
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Washer"
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
      Left            =   1830
      TabIndex        =   80
      Top             =   4215
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mechanic"
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
      Left            =   1830
      TabIndex        =   79
      Top             =   3390
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Month"
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
      Height          =   210
      Index           =   0
      Left            =   1515
      TabIndex        =   78
      Top             =   1035
      Width           =   810
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
      Height          =   210
      Index           =   0
      Left            =   3930
      TabIndex        =   77
      Top             =   1035
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Washing Bay"
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
      Height          =   210
      Index           =   42
      Left            =   7785
      TabIndex        =   76
      Top             =   705
      Width           =   1080
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
      Height          =   210
      Index           =   20
      Left            =   9405
      TabIndex        =   75
      Top             =   705
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accident Repair Bay"
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
      Height          =   210
      Index           =   65
      Left            =   7785
      TabIndex        =   74
      Top             =   1515
      Width           =   1635
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
      Height          =   210
      Index           =   63
      Left            =   9405
      TabIndex        =   73
      Top             =   1515
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
      Height          =   210
      Index           =   62
      Left            =   9405
      TabIndex        =   72
      Top             =   1245
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
      Height          =   210
      Index           =   61
      Left            =   9405
      TabIndex        =   71
      Top             =   975
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
      Height          =   210
      Index           =   60
      Left            =   3930
      TabIndex        =   70
      Top             =   1320
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
      Height          =   210
      Index           =   59
      Left            =   3930
      TabIndex        =   69
      Top             =   765
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Repairing Bay"
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
      Height          =   210
      Index           =   60
      Left            =   7785
      TabIndex        =   68
      Top             =   1245
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Bay"
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
      Height          =   210
      Index           =   59
      Left            =   7785
      TabIndex        =   67
      Top             =   975
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Working Days Of Week"
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
      Height          =   210
      Index           =   58
      Left            =   1515
      TabIndex        =   66
      Top             =   1320
      Width           =   2325
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Card Type"
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
      Height          =   210
      Index           =   57
      Left            =   1515
      TabIndex        =   65
      Top             =   750
      Width           =   1185
   End
End
Attribute VB_Name = "frmWorkDetMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsDiv As ADODB.Recordset
Dim Master As ADODB.Recordset
Private Const JobType As Byte = 0
Private Const JobTypeN As Byte = 61
Private Const Month As Byte = 1
Private Const WorkDays As Byte = 2
Private Const WashBay As Byte = 3
Private Const ServBay As Byte = 4
Private Const RepairBay As Byte = 5
Private Const Accbay As Byte = 6
Private Const Manag1 As Byte = 7, Manag2 As Byte = 8, Manag3 As Byte = 9
Private Const AssManag1 As Byte = 10, AssManag2 As Byte = 11, AssManag3 As Byte = 12
Private Const SupVisor1 As Byte = 13, SupVisor2 As Byte = 14, SupVisor3 As Byte = 15
Private Const Mech1 As Byte = 16, Mech2 As Byte = 17, Mech3 As Byte = 18
Private Const AssMech1 As Byte = 19, AssMech2 As Byte = 20, AssMech3 As Byte = 21
Private Const Helper1 As Byte = 22, Helper2 As Byte = 23, Helper3 As Byte = 24
Private Const Wash1 As Byte = 25, Wash2 As Byte = 26, Wash3 As Byte = 27
Private Const Elect1 As Byte = 28, Elect2 As Byte = 29, Elect3 As Byte = 30
Private Const Paint1 As Byte = 31, Paint2 As Byte = 32, Paint3 As Byte = 33
Private Const Appren1 As Byte = 34, Appren2 As Byte = 35, Appren3 As Byte = 36
Private Const Mach1  As Byte = 37, Mach2  As Byte = 38, Mach3 As Byte = 39
Private Const Cash1 As Byte = 40, Cash2 As Byte = 41, Cash3 As Byte = 42
Private Const Clerk1 As Byte = 43, Clerk2 As Byte = 44, Clerk3 As Byte = 45
Private Const Typist1 As Byte = 46, Typist2 As Byte = 47, Typist3 As Byte = 48
Private Const Operate1 As Byte = 49, Operate2 As Byte = 50, Operate3 As Byte = 51
Private Const Recep1 As Byte = 52, Recep2 As Byte = 53, Recep3 As Byte = 54
Private Const ServAdv1 As Byte = 55, ServAdv2 As Byte = 56, ServAdv3 As Byte = 57
Private Const Guard1 As Byte = 58, Guard2 As Byte = 59, Guard3 As Byte = 60
Private Const HrsPerDay As Byte = 62, TotAtms As Byte = 63, WorkRunCost As Byte = 64
Dim ListArray As Variant
Dim mListItem As ListItem
Dim TAddMode As Boolean
Dim FirmAddFlag As Byte
Dim GridKey As Integer
Dim ExitCtrl As Boolean

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
TopCtrl1.Tag = PubUParam: TopCtrl1.TopText1 = "Workshop Details Master"
For I = 0 To 60
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
Next
         ListArray = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        Set mListItem = ListView_Items(ListView, Txt, Month, ListArray, 12)
   Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    Master.Open "select Wrk_Details.monthno as searchcode,Wrk_Details.*,Division.Div_sname as Divname from (Wrk_Details left join division on Wrk_Details.Job_type=division.div_code) where job_type='" & PubDivCode & "'" & " order by " & cVal("monthno") & "", GCn, adOpenDynamic, adLockOptimistic
    Disp_Text SETS("INI", Me, Master)
    MoveRec
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub Form_Resize()
If Me.WindowState <> vbMaximized Then
    Me.left = MDIForm1.left
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Master = Nothing
End Sub

Private Sub ListView_Click()
Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
FrmList.Visible = False
Txt(Val(ListView.Tag)).SetFocus
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    Txt(JobType) = PubDivCode
    Txt(JobTypeN) = GCn.Execute("select div_sname from division where div_code='" & PubDivCode & "'").Fields(0).Value
    Txt(Month).SetFocus
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
    Dim mo As Integer
Dim XBM
On Error GoTo eloop1
If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
    mo = IIf(Txt(Month) = "Jan", 1, IIf(Txt(Month) = "Fab", 2, IIf(Txt(Month) = "Mar", 3, IIf(Txt(Month) = "Apr", 4, IIf(Txt(Month) = "May", 5, IIf(Txt(Month) = "Jun", 6, IIf(Txt(Month) = "Jul", 7, IIf(Txt(Month) = "Aug", 8, IIf(Txt(Month) = "Sep", 9, IIf(Txt(Month) = "Oct", 10, IIf(Txt(Month) = "Nov", 11, 12)))))))))))
    GCn.BeginTrans
    GCn.Execute ("delete from Wrk_Details where job_type = '" & PubDivCode & "'" & " and monthno=" & mo)
    GCn.CommitTrans
    Master.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
eloop1:
    If err.NUMBER <> 0 Then
       GCn.RollbackTrans
        MsgBox err.Description, vbCritical, " Deletion Message"
    End If
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    Disp_Text SETS("EDIT", Me, Master)
    Txt(Month).Enabled = False
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "SELECT monthno as searchcode, MONTHNO FROM Wrk_Details where job_type='" & PubDivCode & "'"
    Set SearchForm = Me
    FIND.Show vbModal
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
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
    For I = 0 To 61
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
End If
Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_ePrn()
Dim mQRY$
Dim Rst As ADODB.Recordset
    mRepName = "WrkDetails"
    mQRY = "Select * from Wrk_Details where Job_type='" & PubDivCode & "' and Monthno=" & Master!SearchCode & ""
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenStatic, adLockReadOnly
    
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
    
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".TTX", True
    
    Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    rpt.Database.SetDataSource Rst
    rpt.ReadRecords
    Call Report_View(rpt, Me.CAPTION & "", , False)
    Set rpt = Nothing
    Set Rst = Nothing
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim mo As Integer
    Dim mTrans As Boolean
    Dim DocIdHlp As String
    On Error GoTo errlbl
        If IsValid(Txt(Month), "Month") = False Then: Exit Sub
        If IsValid(Txt(WorkDays), "Working Days") = False Then: Exit Sub
    mo = IIf(Txt(Month) = "Jan", 1, IIf(Txt(Month) = "Fab", 2, IIf(Txt(Month) = "Mar", 3, IIf(Txt(Month) = "Apr", 4, IIf(Txt(Month) = "May", 5, IIf(Txt(Month) = "Jun", 6, IIf(Txt(Month) = "Jul", 7, IIf(Txt(Month) = "Aug", 8, IIf(Txt(Month) = "Sep", 9, IIf(Txt(Month) = "Oct", 10, IIf(Txt(Month) = "Nov", 11, 12)))))))))))
    GCn.BeginTrans
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        GCn.Execute "insert into Wrk_Details(Job_Type,Monthno,WDays,Bays_Wash,Bays_Serv,Bays_Repair,Bays_AcRepair,WrkMgr_S,WrkMgr_SS,WrkMgr_US,AsstWrkMgr_S,AsstWrkMgr_SS,AsstWrkMgr_US,Supervisor_S,Supervisor_SS,Supervisor_US,Mechanic_S,Mechanic_SS,Mechanic_US,AsstMechanic_S,AsstMechanic_SS,AsstMechanic_US,Helper_S,Helper_SS,Helper_US,Washer_S,Washer_SS,Washer_US,Elect_S,Elect_SS,Elect_US,Painter_S,Painter_SS,Painter_US,Apprentice_S,Apprentice_SS,Apprentice_US,Machinist_S,Machinist_SS,Machinist_US,Cashier_S,Cashier_SS,Cashier_US,Clerk_S,Clerk_SS,Clerk_US,Typist_S,Typist_SS,Typist_US,Operator_S,Operator_SS,Operator_US,Reception_S,Reception_SS,Reception_US,SrvAdvisor_S,SrvAdvisor_SS,SrvAdvisor_US,Guards_S,Guards_SS,Guards_US,HrsPerDay,TotAtms,WorkRunCost) " & _
          " values('" & PubDivCode & "'," & mo & "," & Val(Txt(WorkDays)) & "," & Val(Txt(WashBay)) & "," & Val(Txt(ServBay)) & "," & Val(Txt(RepairBay)) & "," & Val(Txt(Accbay)) & "," & Val(Txt(Manag1)) & "," & Val(Txt(Manag2)) & "," & Val(Txt(Manag3)) & ", " & _
            " " & Val(Txt(AssManag1)) & "," & Val(Txt(AssManag2)) & "," & Val(Txt(AssManag3)) & "," & Val(Txt(SupVisor1)) & "," & Val(Txt(SupVisor2)) & "," & Val(Txt(SupVisor3)) & "," & Val(Txt(Mech1)) & ", " & Val(Txt(Mech2)) & "," & Val(Txt(Mech3)) & ", " & Val(Txt(AssMech1)) & " ," & Val(Txt(AssMech2)) & " , " & Val(Txt(AssMech3)) & ", " & Val(Txt(Helper1)) & ",  " & Val(Txt(Helper2)) & ", " & Val(Txt(Helper3)) & ", " & Val(Txt(Wash1)) & ", " & Val(Txt(Wash2)) & "," & Val(Txt(Wash3)) & "," & Val(Txt(Elect1)) & ", " & Val(Txt(Elect2)) & ", " & Val(Txt(Elect3)) & ", " & Val(Txt(Paint1)) & ", " & Val(Txt(Paint2)) & "," & Val(Txt(Paint3)) & "," & Val(Txt(Appren1)) & "," & Val(Txt(Appren2)) & "," & Val(Txt(Appren3)) & "," & Val(Txt(Mach1)) & "," & Val(Txt(Mach2)) & "," & Val(Txt(Mach3)) & "," & Val(Txt(Cash1)) & "," & Val(Txt(Cash2)) & "," & Val(Txt(Cash3)) & "," & Val(Txt(Clerk1)) & "," & Val(Txt(Clerk2)) & "," & Val(Txt(Clerk3)) & ", " & _
            " " & Val(Txt(Typist1)) & "," & Val(Txt(Typist2)) & "," & Val(Txt(Typist3)) & "," & Val(Txt(Operate1)) & "," & Val(Txt(Operate2)) & "," & Val(Txt(Operate3)) & "," & Val(Txt(Recep1)) & "," & Val(Txt(Recep2)) & "," & Val(Txt(Recep3)) & "," & Val(Txt(ServAdv1)) & "," & Val(Txt(ServAdv2)) & "," & Val(Txt(ServAdv3)) & "," & Val(Txt(Guard1)) & "," & Val(Txt(Guard2)) & "," & Val(Txt(Guard3)) & "," & Val(Txt(HrsPerDay)) & "," & Val(Txt(TotAtms)) & "," & Val(Txt(WorkRunCost)) & ")"
    Else
        GCn.Execute ("update Wrk_Details set WDays=" & Val(Txt(WorkDays)) & ",Bays_Wash=" & Val(Txt(WashBay)) & ",Bays_Serv=" & Val(Txt(ServBay)) & ",Bays_Repair=" & Val(Txt(RepairBay)) & ",Bays_AcRepair=" & Val(Txt(Accbay)) & ",WrkMgr_S=" & Val(Txt(Manag1)) & ",WrkMgr_SS=" & Val(Txt(Manag2)) & ",WrkMgr_US=" & Val(Txt(Manag3)) & ",AsstWrkMgr_S=" & Val(Txt(AssManag1)) & ",AsstWrkMgr_SS=" & Val(Txt(AssManag2)) & ",AsstWrkMgr_US=" & Val(Txt(AssManag3)) & ", " & _
        " Supervisor_S=" & Val(Txt(SupVisor1)) & ",Supervisor_SS=" & Val(Txt(SupVisor2)) & ",Supervisor_US=" & Val(Txt(SupVisor3)) & ",Mechanic_S=" & Val(Txt(Mech1)) & ",Mechanic_SS= " & Val(Txt(Mech2)) & ",Mechanic_US=" & Val(Txt(Mech3)) & ",AsstMechanic_S= " & Val(Txt(AssMech1)) & ",AsstMechanic_SS=" & Val(Txt(AssMech2)) & " ,AsstMechanic_US= " & Val(Txt(AssMech3)) & ",Helper_S=" & Val(Txt(Helper1)) & ",Helper_SS=" & Val(Txt(Helper2)) & ",Helper_US= " & Val(Txt(Helper3)) & ",Washer_S=" & Val(Txt(Wash1)) & ",Washer_SS=" & Val(Txt(Wash2)) & ",Washer_US=" & Val(Txt(Wash3)) & ",Elect_S=" & Val(Txt(Elect1)) & ",Elect_SS=" & Val(Txt(Elect2)) & ",Elect_US=" & Val(Txt(Elect3)) & ",Painter_S=" & Val(Txt(Paint1)) & ",Painter_SS=" & Val(Txt(Paint2)) & ",Painter_US=" & Val(Txt(Paint3)) & ",Apprentice_S=" & Val(Txt(Appren1)) & ",Apprentice_SS=" & Val(Txt(Appren2)) & ",Apprentice_US=" & Val(Txt(Appren3)) & ", " & _
        " Machinist_S=" & Val(Txt(Mach1)) & ",Machinist_SS=" & Val(Txt(Mach2)) & ",Machinist_US=" & Val(Txt(Mach3)) & ",Cashier_S=" & Val(Txt(Cash1)) & ",Cashier_SS=" & Val(Txt(Cash2)) & ",Cashier_US=" & Val(Txt(Cash3)) & ",Clerk_S=" & Val(Txt(Clerk1)) & ",Clerk_SS=" & Val(Txt(Clerk2)) & ",Clerk_US=" & Val(Txt(Clerk3)) & ",Typist_S=" & Val(Txt(Typist1)) & ",Typist_SS=" & Val(Txt(Typist2)) & ", " & _
        " Typist_US=" & Val(Txt(Typist3)) & ",Operator_S=" & Val(Txt(Operate1)) & ",Operator_SS=" & Val(Txt(Operate2)) & ",Operator_US=" & Val(Txt(Operate3)) & ",Reception_S=" & Val(Txt(Recep1)) & ",Reception_SS=" & Val(Txt(Recep2)) & ",Reception_US=" & Val(Txt(Recep3)) & ",SrvAdvisor_S=" & Val(Txt(ServAdv1)) & ",SrvAdvisor_SS=" & Val(Txt(ServAdv2)) & ",SrvAdvisor_US=" & Val(Txt(ServAdv3)) & ",Guards_S=" & Val(Txt(Guard1)) & ",Guards_SS=" & Val(Txt(Guard2)) & ",Guards_US=" & Val(Txt(Guard3)) & ",HrsPerDay=" & Val(Txt(HrsPerDay)) & ",TotAtms=" & Val(Txt(TotAtms)) & ",WorkRunCost=" & Val(Txt(WorkRunCost)) & " where Job_Type='" & PubDivCode & "'" & " and monthno=" & mo)
    End If
    GCn.CommitTrans
    mTrans = False
    Master.Requery
    Master.FIND "monthno = '" & mo & "'"
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
    Grid_Hide
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Byte
Dim Txtdate As Boolean
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case Month
        ListView_KeyDown FrmList, ListView, Txt, Month, KeyCode, Shift, Txt(Month).left, (Txt(Month).top + Txt(Month).height), Txt(Month).width, 3200
End Select
If FrmList.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index <> Guard3 Then Ctrl_DownKeyDown KeyCode, Shift
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = Guard3 Then
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
        End If
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> Month Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> Month Then
            If KeyCode = vbKeyUp Or KeyCode = vbKeyReturn Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub Txt_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
 Call CheckQuote(KeyAscii)
 If Index = Month Then Exit Sub
 Select Case Index
    Case WorkRunCost
        Call NumPress(Txt(Index), KeyAscii, 10, 2)
    Case Else
        Call NumPress(Txt(Index), KeyAscii, 3, 0)
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case Month
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Month, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim mo As Integer
Select Case Index
    Case WorkDays
         If Txt(WorkDays).TEXT = "" Then Exit Sub
         If Val(Txt(WorkDays).TEXT) > 7 Then MsgBox " Days Should Not be Greater Than 7  ", vbInformation, "Validation": Cancel = True
    Case Month
        mo = IIf(Txt(Month) = "Jan", 1, IIf(Txt(Month) = "Feb", 2, IIf(Txt(Month) = "Mar", 3, IIf(Txt(Month) = "Apr", 4, IIf(Txt(Month) = "May", 5, IIf(Txt(Month) = "Jun", 6, IIf(Txt(Month) = "Jul", 7, IIf(Txt(Month) = "Aug", 8, IIf(Txt(Month) = "Sep", 9, IIf(Txt(Month) = "Oct", 10, IIf(Txt(Month) = "Nov", 11, 12)))))))))))
            If GCn.Execute("select count(*) FROM Wrk_Details WHERE Job_Type= '" & PubDivCode & "'" & " and MonthNo=" & mo).Fields(0).Value > 0 Then
                MsgBox "Duplicate Order No ", vbInformation, "Validation Check"
                Cancel = True
                Txt(Month).SetFocus
                Exit Sub
            End If
End Select
End Sub


'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To 61
    Txt(I).TEXT = ""
Next I
End Sub

Private Sub MoveRec()
Dim mo As String
If Master.RecordCount > 0 Then
    mo = IIf(Trim(STR(Master!MonthNo)) = "1", "Jan", IIf(Trim(STR(Master!MonthNo)) = "2", "Fab", IIf(Trim(STR(Master!MonthNo)) = "3", "Mar", IIf(Trim(STR(Master!MonthNo)) = "4", "Apr", IIf(Trim(STR(Master!MonthNo)) = "5", "May", IIf(Trim(STR(Master!MonthNo)) = "6", "Jun", IIf(Trim(STR(Master!MonthNo)) = "7", "Jul", IIf(Trim(STR(Master!MonthNo)) = "8", "Aug", IIf(Trim(STR(Master!MonthNo)) = "9", "Sep", IIf(Trim(STR(Master!MonthNo)) = "10", "Oct", IIf(Trim(STR(Master!MonthNo)) = "11", "Nov", "Dec")))))))))))
    Txt(JobType) = Master!job_type  '0
    Txt(JobTypeN) = Master!DivName
    Txt(Month) = mo
    Txt(WorkDays) = VNull(Master!WDays)
    Txt(HrsPerDay) = VNull(Master!HrsPerDay)
    Txt(TotAtms) = VNull(Master!TotAtms)
    Txt(WorkRunCost) = Format(VNull(Master!WorkRunCost), "0.00")
    
    Txt(WashBay) = VNull(Master!Bays_Wash)
    Txt(ServBay) = IIf(IsNull(Master!Bays_Serv), 0, Master!Bays_Serv)
    Txt(RepairBay) = IIf(IsNull(Master!Bays_Repair), 0, Master!Bays_Repair)
    Txt(Accbay) = IIf(IsNull(Master!Bays_AcRepair), 0, Master!Bays_AcRepair)
    Txt(Manag1) = VNull(Master!WrkMgr_S)
    Txt(Manag2) = IIf(IsNull(Master!WrkMgr_SS), 0, Master!WrkMgr_SS)
    Txt(Manag3) = IIf(IsNull(Master!WrkMgr_US), 0, Master!WrkMgr_US)
    Txt(AssManag1) = IIf(IsNull(Master!AsstWrkMgr_S), 0, Master!AsstWrkMgr_S)
    Txt(AssManag2) = IIf(IsNull(Master!AsstWrkMgr_SS), 0, Master!AsstWrkMgr_SS) '10
    Txt(AssManag3) = IIf(IsNull(Master!AsstWrkMgr_US), 0, Master!AsstWrkMgr_US)
    Txt(SupVisor1) = IIf(IsNull(Master!Supervisor_S), 0, Master!Supervisor_S)
    Txt(SupVisor2) = IIf(IsNull(Master!Supervisor_SS), 0, Master!Supervisor_SS)
    Txt(SupVisor3) = IIf(IsNull(Master!Supervisor_US), 0, Master!Supervisor_US)
    Txt(Mech1) = IIf(IsNull(Master!Mechanic_S), 0, Master!Mechanic_S)
    Txt(Mech2) = IIf(IsNull(Master!Mechanic_SS), 0, Master!Mechanic_SS)
    Txt(Mech3) = IIf(IsNull(Master!Mechanic_US), 0, Master!Mechanic_US)
    Txt(AssMech1) = IIf(IsNull(Master!AsstMechanic_S), 0, Master!AsstMechanic_S)
    Txt(AssMech2) = IIf(IsNull(Master!AsstMechanic_SS), 0, Master!AsstMechanic_SS)
    Txt(AssMech3) = IIf(IsNull(Master!AsstMechanic_US), 0, Master!AsstMechanic_US) '20
    Txt(Helper1) = IIf(IsNull(Master!Helper_S), 0, Master!Helper_S)
    Txt(Helper2) = IIf(IsNull(Master!Helper_SS), 0, Master!Helper_SS)
    Txt(Helper3) = IIf(IsNull(Master!Helper_US), 0, Master!Helper_US)
    Txt(Wash1) = IIf(IsNull(Master!Washer_S), 0, Master!Washer_S)
    Txt(Wash2) = IIf(IsNull(Master!Washer_SS), 0, Master!Washer_SS)
    Txt(Wash3) = IIf(IsNull(Master!Washer_US), 0, Master!Washer_US)
    Txt(Elect1) = IIf(IsNull(Master!Elect_S), 0, Master!Elect_S)
    Txt(Elect2) = IIf(IsNull(Master!Elect_SS), 0, Master!Elect_SS)
    Txt(Elect3) = IIf(IsNull(Master!Elect_US), 0, Master!Elect_US)
    Txt(Paint1) = IIf(IsNull(Master!Painter_S), 0, Master!Painter_S)
    Txt(Paint2) = IIf(IsNull(Master!Painter_SS), 0, Master!Painter_SS) '30
    Txt(Paint3) = IIf(IsNull(Master!Painter_US), 0, Master!Painter_US)
    Txt(Appren1) = IIf(IsNull(Master!Apprentice_S), 0, Master!Apprentice_S)
    Txt(Appren2) = IIf(IsNull(Master!Apprentice_SS), 0, Master!Apprentice_SS)
    Txt(Appren3) = IIf(IsNull(Master!Apprentice_US), 0, Master!Apprentice_US)
    Txt(Mach1) = IIf(IsNull(Master!Machinist_S), 0, Master!Machinist_S)
    Txt(Mach2) = IIf(IsNull(Master!Machinist_SS), 0, Master!Machinist_SS)
    Txt(Mach3) = IIf(IsNull(Master!Machinist_US), 0, Master!Machinist_US)
    Txt(Cash1) = IIf(IsNull(Master!Cashier_S), 0, Master!Cashier_S)
    Txt(Cash2) = IIf(IsNull(Master!Cashier_SS), 0, Master!Cashier_SS)
    Txt(Cash3) = IIf(IsNull(Master!Cashier_US), 0, Master!Cashier_US) '40
    Txt(Clerk1) = IIf(IsNull(Master!Clerk_S), 0, Master!Clerk_S)
    Txt(Clerk2) = IIf(IsNull(Master!Clerk_SS), 0, Master!Clerk_SS)
    Txt(Clerk3) = IIf(IsNull(Master!Clerk_US), 0, Master!Clerk_US)
    Txt(Typist1) = IIf(IsNull(Master!Typist_S), 0, Master!Typist_S)
    Txt(Typist2) = IIf(IsNull(Master!Typist_SS), 0, Master!Typist_SS)
    Txt(Typist3) = IIf(IsNull(Master!Typist_US), 0, Master!Typist_US)
    Txt(Operate1) = IIf(IsNull(Master!Operator_S), 0, Master!Operator_S)
    Txt(Operate2) = IIf(IsNull(Master!Operator_SS), 0, Master!Operator_SS)
    Txt(Operate3) = IIf(IsNull(Master!Operator_US), 0, Master!Operator_US)
    Txt(Recep1) = IIf(IsNull(Master!Reception_S), 0, Master!Reception_S)
    Txt(Recep2) = IIf(IsNull(Master!Reception_SS), 0, Master!Reception_SS)
    Txt(Recep3) = IIf(IsNull(Master!Reception_US), 0, Master!Reception_US)
    Txt(ServAdv1) = IIf(IsNull(Master!SrvAdvisor_S), 0, Master!SrvAdvisor_S)
    Txt(ServAdv2) = IIf(IsNull(Master!SrvAdvisor_SS), 0, Master!SrvAdvisor_SS)
    Txt(ServAdv3) = IIf(IsNull(Master!SrvAdvisor_US), 0, Master!SrvAdvisor_US)
    Txt(Guard1) = IIf(IsNull(Master!Guards_S), 0, Master!Guards_S)
    Txt(Guard2) = IIf(IsNull(Master!Guards_SS), 0, Master!Guards_SS)
    Txt(Guard3) = IIf(IsNull(Master!Guards_US), 0, Master!Guards_US)
End If
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To 64
    Txt(I).Enabled = Enb
Next
Txt(JobType).Enabled = False
Txt(JobTypeN).Enabled = False
End Sub
Private Sub Grid_Hide()
    If FrmList.Visible = True Then FrmList.Visible = False
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("SEARCHCODE='" & MyValue & "'")
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

