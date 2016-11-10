VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FaRepView 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9135
   Icon            =   "FaRepView.frx":0000
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   9135
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8565
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1830
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000A&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   5820
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000A&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   6210
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000A&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   2880
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000A&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   3270
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   390
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6870
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11835
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FaRepView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tReport

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyHome
        Command1.SetFocus
        Command2_Click (0)
    Case vbKeyPageUp
        Command1.SetFocus
        Command2_Click (1)
    Case vbKeyPageDown
        Command1.SetFocus
        Command2_Click (2)
    Case vbKeyEnd
        Command1.SetFocus
        Command2_Click (3)
    Case vbKeyEscape
        Unload Me
End Select
End Sub
Private Sub CRViewer1_DownloadFinished(ByVal loadingType As CRVIEWERLibCtl.CRLoadingType)
    Command1.SetFocus
End Sub
Private Sub Command2_Click(Index As Integer)
On Error Resume Next
Select Case Index
    Case 0
        CRViewer1.ShowFirstPage
    Case 1
        CRViewer1.ShowPreviousPage
    Case 2
        CRViewer1.ShowNextPage
    Case 3
        CRViewer1.ShowLastPage
End Select
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub CRViewer1_RefreshButtonClicked(UseDefault As Boolean)
    tReport.PrinterSetup (0)
End Sub
Private Sub Form_Load()
    Me.left = 0
    Me.top = 0
    Me.width = 11900
    Me.height = 7935
    rpt.ExportOptions.FormatType = crEFTPaginatedText
    Me.Show
End Sub
Private Sub Form_Resize()
    CRViewer1.top = 0
    CRViewer1.left = 0
    CRViewer1.height = ScaleHeight
    CRViewer1.width = ScaleWidth
End Sub
Property Let Rep_Set(mREPORT)
    Set tReport = mREPORT
    CRViewer1.ReportSource = mREPORT
    CRViewer1.ViewReport
End Property
