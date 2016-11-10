VERSION 5.00
Begin VB.Form TrialEnd 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "TrialEnd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Thanks !"
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
      Left            =   2145
      TabIndex        =   0
      Top             =   735
      Width           =   1305
   End
   Begin VB.Label Label6 
      Caption         =   "E_Mail : sales@datamannet.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   195
      TabIndex        =   6
      Top             =   2370
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   "Web : www.datamannet.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   5
      Top             =   2145
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Phones : 0512-2317191 , 2316505 , 2309410 , 3092334"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   195
      TabIndex        =   4
      Top             =   1935
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "25/16, Karachi Khana, Kanpur-208001"
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
      Left            =   195
      TabIndex        =   3
      Top             =   1710
      Width           =   5385
   End
   Begin VB.Label Label2 
      Caption         =   "Dataman Computer Systems (P) Ltd."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   195
      TabIndex        =   2
      Top             =   1275
      Width           =   4830
   End
   Begin VB.Line Line1 
      X1              =   -15
      X2              =   6090
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   270
      Shape           =   3  'Circle
      Top             =   90
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "You have crossed the trial limit.Please enter the valid Product Serial No."
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
      Left            =   1110
      TabIndex        =   1
      Top             =   135
      Width           =   4395
   End
End
Attribute VB_Name = "TrialEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub
