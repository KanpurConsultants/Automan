VERSION 5.00
Begin VB.Form FrmBackup 
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6840
   Begin VB.Image Image1 
      Height          =   210
      Index           =   1
      Left            =   1680
      Picture         =   "FrmBackup.frx":0000
      Stretch         =   -1  'True
      Top             =   1335
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image Image2 
      Height          =   225
      Index           =   1
      Left            =   1800
      Picture         =   "FrmBackup.frx":015E
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   210
      Index           =   0
      Left            =   1680
      Picture         =   "FrmBackup.frx":0470
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image Image2 
      Height          =   225
      Index           =   0
      Left            =   1800
      Picture         =   "FrmBackup.frx":05CE
      Top             =   960
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Balance (Stock && A/c) Updation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Database Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "FrmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
