VERSION 5.00
Begin VB.Form FrmMdiBack 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10770
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   8340
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13665
   End
End
Attribute VB_Name = "FrmMdiBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Me.Picture = LoadPicture(PubRepoPath & "\Wallpaper.JPG")
    WinSetting Me, 7400, 11800
    Image1.Move 0, 0, Me.width, Me.height
    'SetResolutionFormLoad Me
End Sub
