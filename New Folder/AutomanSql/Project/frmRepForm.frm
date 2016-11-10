VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRepForm 
   Caption         =   "Report Display Form"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   8490
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrysReport1 
      Left            =   1590
      Top             =   1215
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "frmRepForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.Show
    CrysReport1.WindowParentHandle = Me.hwnd
    CrysReport1.Destination = crptToWindow
End Sub

