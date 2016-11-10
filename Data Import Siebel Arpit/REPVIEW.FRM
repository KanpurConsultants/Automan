VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form REPVIEW 
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7410
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   7455
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "REPVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tReport

Private Sub CRViewer1_RefreshButtonClicked(UseDefault As Boolean)
    tReport.PrinterSetup (0)
End Sub
Private Sub Form_Load()
    Call WinSetting(Me)
    CRViewer1.top = 0
    CRViewer1.left = 0
    CRViewer1.height = ScaleHeight
    CRViewer1.width = ScaleWidth
    Me.Show
End Sub
Property Let Rep_Set(mREPORT)
    Set tReport = mREPORT
    CRViewer1.ReportSource = mREPORT
    CRViewer1.ViewReport
End Property
