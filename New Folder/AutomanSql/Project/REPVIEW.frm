VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form REPVIEW 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   7365
   Begin MSComDlg.CommonDialog cDialogRepView 
      Left            =   8295
      Top             =   1860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7410
      Left            =   15
      TabIndex        =   0
      Top             =   -45
      Width           =   7455
      lastProp        =   500
      _cx             =   5080
      _cy             =   5080
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
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "REPVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tReport
Private Sub CRViewer1_ExportButtonClicked(UseDefault As Boolean)
'Disable CrViewer default working
If MsgBox("Use Crystel Reports Export Module ? ", vbYesNo + vbInformation) = vbYes Then
    Exit Sub
End If
UseDefault = False
cDialogRepView.CancelError = True
On Error GoTo ErrHandler
Dim oExportOptions As CRAXDRT.ExportOptions, nFilterIndex As Byte
cDialogRepView.Filter = "Excel Files(*.xls)|*.xls|"
cDialogRepView.DialogTitle = "Select File Name"
'cDialogRepView.FileName = mRepName
cDialogRepView.ShowSave

If cDialogRepView.FileName <> "" Then
    Set oExportOptions = tReport.ExportOptions
    With oExportOptions
        .DestinationType = crEDTDiskFile
        Select Case UCase(Right(cDialogRepView.FileName, 4))
            Case ".DOC"
                .FormatType = crEFTWordForWindows
            Case Else
                .FormatType = crEFTExcel80
        End Select
        
        .DiskFileName = cDialogRepView.FileName
    End With
    tReport.Export False
End If
'eof export code
'
ErrHandler:
  'User pressed the Cancel button
  Exit Sub
End Sub
Private Function FileFormatType(nFilterIndex As Byte) As Byte
'By     : LP Singh 13-08-2005
'Purpose: Function used to detect export file format selected by user
If nFilterIndex = 1 Then
    FileFormatType = crEFTCommaSeparatedValues
ElseIf nFilterIndex = 2 Then
    FileFormatType = crEFTExcel80
ElseIf nFilterIndex = 3 Then
    FileFormatType = crEFTText
ElseIf nFilterIndex = 4 Then
    FileFormatType = crEFTWordForWindows
End If
End Function

Private Sub CRViewer1_RefreshButtonClicked(UseDefault As Boolean)
    tReport.PrinterSetup (0)
End Sub

Private Sub Form_Load()
    Call WinSetting(Me)
    Me.WindowState = 2
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
Private Sub Form_Resize()
    CRViewer1.top = 0
    CRViewer1.left = 0
    CRViewer1.height = ScaleHeight
    CRViewer1.width = ScaleWidth
End Sub

