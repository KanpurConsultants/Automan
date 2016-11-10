VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVehChasDCard 
   Appearance      =   0  'Flat
   BackColor       =   &H00BAD3C9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vehicle Chassis Card"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   9855
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Printer "
      Top             =   4410
      Width           =   1590
   End
   Begin VB.TextBox txtPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7455
      TabIndex        =   6
      Top             =   555
      Width           =   1980
   End
   Begin VB.TextBox txtPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4065
      MaxLength       =   25
      TabIndex        =   5
      Top             =   555
      Width           =   1980
   End
   Begin VB.TextBox txtPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7455
      MaxLength       =   20
      TabIndex        =   4
      Top             =   285
      Width           =   1980
   End
   Begin VB.TextBox txtPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4065
      TabIndex        =   3
      Top             =   285
      Width           =   1980
   End
   Begin MSDataGridLib.DataGrid DGChass 
      Height          =   1980
      Left            =   750
      Negotiate       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1125
      Visible         =   0   'False
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   3493
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Chassis Help"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Code1"
         Caption         =   "Chassis No"
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
         DataField       =   "Code2"
         Caption         =   "Engine No"
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
         DataField       =   "Code3"
         Caption         =   "Telco Inv No"
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
         DataField       =   "Code4"
         Caption         =   "Date"
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
            ColumnWidth     =   2129.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2310.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1635.024
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton OptPlain 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00BAD3C9&
      Caption         =   "Plain"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   540
      Value           =   -1  'True
      Width           =   750
   End
   Begin VB.OptionButton Optpre 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00BAD3C9&
      Caption         =   "PrePrinted "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   1590
      TabIndex        =   2
      Top             =   540
      Width           =   1200
   End
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00F8D7FD&
      Caption         =   "Speed &Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Printer "
      Top             =   3420
      Width           =   1590
   End
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Screen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Screen"
      Top             =   3750
      Width           =   1590
   End
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Windows Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Printer "
      Top             =   4080
      Width           =   1590
   End
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   435
      Picture         =   "frmVehChasDCard.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Screen"
      Top             =   4425
      Width           =   315
   End
   Begin MSDataGridLib.DataGrid DGModel 
      Height          =   1980
      Left            =   2280
      Negotiate       =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2145
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3493
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Model Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Model"
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
            ColumnWidth     =   3014.929
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   2685
      X2              =   2685
      Y1              =   435
      Y2              =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telco Bill No."
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
      Left            =   6270
      TabIndex        =   18
      Top             =   570
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Name"
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
      Index           =   19
      Left            =   2910
      TabIndex        =   17
      Top             =   300
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine No."
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
      Index           =   20
      Left            =   2925
      TabIndex        =   16
      Top             =   570
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chassis No."
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
      Index           =   22
      Left            =   6270
      TabIndex        =   15
      Top             =   300
      Width           =   1035
   End
   Begin VB.Line Line1 
      X1              =   195
      X2              =   195
      Y1              =   435
      Y2              =   555
   End
   Begin VB.Label LblPrinter 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Current Active Printer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   765
      TabIndex        =   14
      Top             =   4425
      Width           =   7275
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Stationary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   41
      Left            =   960
      TabIndex        =   13
      Top             =   120
      Width           =   825
   End
   Begin VB.Line Line6 
      X1              =   2670
      X2              =   195
      Y1              =   435
      Y2              =   435
   End
   Begin VB.Line Line8 
      X1              =   1320
      X2              =   1320
      Y1              =   330
      Y2              =   420
   End
End
Attribute VB_Name = "frmVehChasDCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsModel As ADODB.Recordset
Dim RsChass As ADODB.Recordset
Dim ModelCode As String
Private Const Model As Byte = 0
Private Const ChasNo As Byte = 1
Private Const Engno As Byte = 2
Private Const TelcoBill As Byte = 3

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String

Private Sub DGChass_Click()
    DGChass.Visible = False
    If RsChass.RecordCount > 0 Then
        txtPrint(Model).TEXT = RsChass!Model
        txtPrint(ChasNo).TEXT = RsChass!CODE1
        txtPrint(Engno).TEXT = RsChass!code2
        txtPrint(TelcoBill).TEXT = RsChass!code3
    Else
        txtPrint(Model) = ""
        txtPrint(ChasNo).TEXT = ""
        txtPrint(Engno).TEXT = ""
        txtPrint(TelcoBill).TEXT = ""
    End If
End Sub

Private Sub DGModel_Click()
    DgModel.Visible = False
    If RsModel.RecordCount > 0 Then
        txtPrint(Model).TEXT = RsModel!Name
        txtPrint(Model).Tag = RsModel!Code
    End If
    txtPrint(Model).SetFocus
End Sub

Private Sub DGSite_Click()

End Sub


Private Sub Form_Activate()
'If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
    
    Set RsModel = New ADODB.Recordset
    RsModel.CursorLocation = adUseClient
    RsModel.Open "select Model as code,model as name from Model where Model in (select distinct Model from Veh_Stock) order by model", GCn, adOpenDynamic, adLockOptimistic
    Set DgModel.DataSource = RsModel
    
    Set RsChass = New ADODB.Recordset
    RsChass.CursorLocation = adUseClient
  '  RsChass.Open "SELECT distinct Veh_Stock.ChassisNo as code1, Veh_Stock.EngineNo as code2, Veh_Stock.PBILL_NO as code3, Veh_Stock.PBILL_DATE as code4,Veh_Stock.Model FROM Veh_Stock order by ChassisNo,engineno,PBILL_NO", GCn, adOpenDynamic, adLockOptimistic
     RsChass.Open "SELECT right(Veh_Stock.ChassisNo,5) as code1,Veh_Stock.ChassisNo as ChassisNo, Veh_Stock.EngineNo as code2, Veh_Stock.PBILL_NO as code3, Veh_Stock.PBILL_DATE as code4,Veh_Stock.Model FROM Veh_Stock order by ChassisNo,engineno,PBILL_NO", GCn, adOpenDynamic, adLockOptimistic
    RsChass.Sort = "CODE1"
    RsChass.Sort = "CODE2"
    RsChass.Sort = "CODE3"
    Set DGChass.DataSource = RsChass

LblPrinter.CAPTION = Printer.DeviceName
Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsChass = Nothing
Set RsModel = Nothing
End Sub
Private Sub Grid_Hide()
    If DGChass.Visible = True Then DGChass.Visible = False
    If DgModel.Visible = True Then DgModel.Visible = False
End Sub

'************************ PRINTING CODE ******************
Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case ChasNo
        If RsChass.RecordCount = 0 Or (RsChass.EOF = True Or RsChass.BOF = True) Then Exit Sub
        If txtPrint(ChasNo) = "" Then
            RsChass.Sort = "CODE1"
        Else
            RsChass.Sort = "CODE1"
            RsChass.MoveFirst
            RsChass.FIND "ChassisNo ='" & txtPrint(ChasNo) & "'"
            If RsChass.EOF = True Then RsChass.MoveFirst
        End If
    Case Engno
        If RsChass.RecordCount = 0 Or (RsChass.EOF = True Or RsChass.BOF = True) Then Exit Sub
        If txtPrint(Engno) = "" Then
            RsChass.Sort = "CODE2"
        Else
            RsChass.Sort = "CODE2"
            RsChass.MoveFirst
            RsChass.FIND "code2 ='" & txtPrint(Engno) & "'"
            If RsChass.EOF = True Then RsChass.MoveFirst
        End If
    Case TelcoBill
        If RsChass.RecordCount = 0 Or (RsChass.EOF = True Or RsChass.BOF = True) Then Exit Sub
        If txtPrint(TelcoBill) = "" Then
            RsChass.Sort = "CODE3"
        Else
            RsChass.Sort = "CODE3"
            RsChass.MoveFirst
            RsChass.FIND "code3 ='" & txtPrint(TelcoBill) & "'"
            If RsChass.EOF = True Then RsChass.MoveFirst
        End If
    Case Model
        If RsModel.RecordCount = 0 Or (RsModel.EOF = True Or RsModel.BOF = True) Then Exit Sub
        If txtPrint(Index).TEXT = "" Then
            txtPrint(Index).Tag = ""
            txtPrint(Index).TEXT = ""
        Else
            If txtPrint(Index).TEXT <> RsModel!Name Then
                RsModel.MoveFirst
                RsModel.FIND "name ='" & txtPrint(Index).TEXT & "'"
            End If
        End If
    
End Select
End Sub

Private Sub TxtPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case ChasNo
        If DGChass.Visible = False Then DGridColSwap DGChass, 0
        DGridTxtKeyDown DGChass, txtPrint, Index, RsChass, KeyCode, False, 0
    Case Engno
        If DGChass.Visible = False Then DGridColSwap DGChass, 1
        DGridTxtKeyDown DGChass, txtPrint, Index, RsChass, KeyCode, False, 1
    Case TelcoBill
        If DGChass.Visible = False Then DGridColSwap DGChass, 2
        DGridTxtKeyDown DGChass, txtPrint, Index, RsChass, KeyCode, False, 2
    Case Model
        DGridTxtKeyDown DgModel, txtPrint, Index, RsModel, KeyCode, False, 1
End Select
If DgModel.Visible = False And DGChass.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> Model Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TxtPrint_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case Index
     Case Model
        If DgModel.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsModel, KeyAscii, "Name"
    Case ChasNo
        If DGChass.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsChass, KeyAscii, "code1"
    Case Engno
        If DGChass.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsChass, KeyAscii, "code2"
    Case TelcoBill
        If DGChass.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsChass, KeyAscii, "code3"
End Select
End Sub

Private Sub TxtPrint_LostFocus(Index As Integer)
  Ctrl_validate txtPrint(Index)
End Sub

Private Sub TxtPrint_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
   Case ChasNo, Engno
        If RsChass.RecordCount = 0 Or (RsChass.EOF = True Or RsChass.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(ChasNo).TEXT = ""
            txtPrint(Engno).TEXT = ""
            txtPrint(TelcoBill).TEXT = ""
        Else
            txtPrint(Model).TEXT = RsChass!Model
            'ChassisNo
           ' txtPrint(ChasNo).TEXT = RsChass!CODE1
            txtPrint(ChasNo).TEXT = RsChass!ChassisNo
            txtPrint(Engno).TEXT = RsChass!code2
            txtPrint(TelcoBill).TEXT = RsChass!code3
        End If
    Case TelcoBill
        If RsChass.RecordCount = 0 Or (RsChass.EOF = True Or RsChass.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(TelcoBill).TEXT = ""
        Else
            txtPrint(ChasNo).TEXT = RsChass!CODE1
            txtPrint(Engno).TEXT = RsChass!code2
            txtPrint(TelcoBill).TEXT = RsChass!code3
        End If
   Case Model
        If IsValid(txtPrint(Index), "Model") = False Then Exit Sub
        If RsModel.RecordCount = 0 Or (RsModel.EOF = True Or RsModel.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(Index).TEXT = ""
            txtPrint(Index).Tag = ""
            ModelCode = ""
        Else
            If ModelCode <> RsModel!Code Then
                RsChass.Close
                RsChass.Open "SELECT distinct Veh_Stock.ChassisNo as code1, Veh_Stock.EngineNo as code2, Veh_Stock.PBILL_NO as code3, Veh_Stock.PBILL_DATE as code4,Model FROM Veh_Stock where model = '" & RsModel!Code & "'  order by ChassisNo,engineno,PBILL_NO", GCn, adOpenDynamic, adLockOptimistic
                Set DGChass.DataSource = RsChass
                txtPrint(ChasNo).TEXT = ""
                txtPrint(Engno).TEXT = ""
                txtPrint(TelcoBill).TEXT = ""
            End If
            txtPrint(Index).TEXT = RsModel!Name
            txtPrint(Index).Tag = RsModel!Code
            ModelCode = RsModel!Code
        End If

End Select
End Sub

Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If
End Sub

Private Sub CmdPrint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
Select Case Index
    Case PScreen, PWindows, PDos
        mRepName = IIf(OptPlain.Value = True, "VehChassDet", "VehChassDet")
        Call WindowsPrint(Index)
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "VehChassDet", "VehChassDet")
        Call PrinerSetUp
    Case PClose 'Close Report Frame
        Unload Me
End Select
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub WindowsPrint(Index As Integer)
Dim Rst As ADODB.Recordset, RstSub1 As ADODB.Recordset, RstSub2 As ADODB.Recordset
Dim RstSub3 As ADODB.Recordset, mQry As String
Dim I As Integer, Rst2 As ADODB.Recordset
On Error GoTo ERRORHANDLER
     
     If IsValid(txtPrint(Model), "Model") = False Then Exit Sub
     If IsValid(txtPrint(ChasNo), "Chassis No") = False Then Exit Sub
    
     mQry = "SELECT VO.STAMP_DUTY,VO.model,VO.REG_FEE, VO.INS_FEE, VO.S_CHARGE, VO.Net_AMOUNT, " & _
        " VO.MISC_INFO, Godown.God_Name, City.CityName, City_2.CityName, " & _
        " Model.Model_Desc, Model.Model_Desc1, CF.FinName,CF.Add1, CF.Add2, " & _
        " CF.PinCode,VP1.V_Date as PurDt,VP1.V_NO as PurVno, VP1.Tot_Amount, " & _
        " SubGroup.Name, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, SubGroup.PIN, VStk.Pur_DocId, " & _
        " VStk.ChassisNo, VStk.EngineNo, VStk.Chassis_RctDocNo, VStk.Chassis_RctDate, VStk.AL_Name, " & _
        " SubGroup_1.Name as Supplier, VP1.PBILL_NO, VP1.PBILL_DATE, VP1.V_NO, VP1.V_Date, " & _
        " VO.OrdDocId, VO.Ord_Date, Emp_Mast.Emp_Name, VO.Inv_DocId, VO.Inv_Date, VO.VRATE, VO.MARGINE, " & _
        " VO.Transport, VO.OtherChrg, VO.TAX_Amt, VO.Surcharge_Amt, VO.FIN_AMT, VO.Interest, VO.DelCh_DocId, VO.DelCh_DT, vo.REBATE , vo.SpecialDiscount " & _
        " FROM (((((((((Veh_Stock VStk LEFT JOIN Veh_Purch1 VP1 ON VStk.Pur_DocId = VP1.DocID) " & _
        " LEFT JOIN Veh_Order VO on VStk.ChassisNo=VO.Chassis) " & _
        " LEFT JOIN ContractFinance CF on VO.FB_CODE=CF.FinCode) " & _
        " LEFT JOIN City AS City_2 ON CF.City = City_2.CityCode) " & _
        " LEFT JOIN Model ON VO.MODEL = Model.MODEL) " & _
        " LEFT JOIN SubGroup ON VO.PartyCode = SubGroup.SubCode) " & _
        " LEFT JOIN City ON SubGroup.CityCode = City.CityCode) " & _
        " LEFT JOIN Godown ON VStk.Godown = Godown.God_Code) " & _
        " LEFT JOIN SubGroup AS SubGroup_1 ON VP1.PARTYCODE = SubGroup_1.SubCode) " & _
        " LEFT JOIN Emp_Mast ON VO.REP_CODE = Emp_Mast.Emp_Code " & _
        " where VStk.chassisno ='" & txtPrint(ChasNo).TEXT & "' and Vstk.model ='" & txtPrint(Model) & "'"
    
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
   
   mQry = "SELECT Veh_AMDModel.prod_name, Veh_Purch2.Trn_Type, Veh_Purch2.docid,((Veh_Purch2.QTY *Veh_Purch2.RATE) + Veh_Purch2.TAX_AMT + Veh_Purch2.TaxSur_AMT) as Amt  " & _
   "FROM (Veh_Stock LEFT JOIN Veh_Purch2 ON Veh_Stock.Pur_DocId = Veh_Purch2.DocID) LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code " & _
   "where veh_stock.chassisno ='" & txtPrint(ChasNo).TEXT & "' and  Veh_stock.model ='" & txtPrint(Model) & "'"
   
   Set RstSub1 = New Recordset
   RstSub1.CursorLocation = adUseClient
   RstSub1.Open (mQry), GCn, adOpenDynamic, adLockOptimistic

   'Recordset is made for subreport2
   
    mQry = "SELECT Rect.Prov_No,Rect.Ord_DocId, Rect.V_Type, Rect.V_No, Rect.V_Date, Rect.Site_Code, Rect.AMOUNT, Rect.DrCr, Rect.Narration, Veh_Order.REBATE , Veh_Order.SpecialDiscount,Veh_Order.Net_AMOUNT " & _
    "FROM Veh_Order LEFT JOIN Rect ON Veh_Order.OrdDocId = Rect.Ord_DocId " & _
    "where Veh_Order.chassis ='" & txtPrint(ChasNo).TEXT & "' and  Veh_Order.model ='" & txtPrint(Model) & "'"
    
   Set RstSub2 = New Recordset
   RstSub2.CursorLocation = adUseClient
   RstSub2.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
   
    mQry = "SELECT Job_Card.Job_No, Job_Card.Job_Date, Job_Card.FreeSpr_Amt, Job_Card.NetLab_Amt, SP_Sale.Total_Amt " & _
    "FROM (Job_Card " & _
    "LEFT JOIN SP_Sale ON Job_Card.DocId_InvSpr = SP_Sale.DocID) " & _
    "LEFT JOIN HisCard ON Job_Card.CardNo = HisCard.CardNo " & _
    "where HisCard.Chassis ='" & txtPrint(ChasNo).TEXT & "'"
   
    Set RstSub3 = New Recordset
    RstSub3.CursorLocation = adUseClient
    RstSub3.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
     
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    CreateFieldDefFile RstSub1, PubRepoPath + "\" & mRepName & "1.ttx", True
    CreateFieldDefFile RstSub2, PubRepoPath + "\" & mRepName & "2.ttx", True
    CreateFieldDefFile RstSub3, PubRepoPath + "\" & mRepName & "3.ttx", True
                  
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    rpt.Database.SetDataSource Rst
    rpt.OpenSubreport("SUBREP1").Database.SetDataSource RstSub1
    rpt.OpenSubreport("SUBREP2").Database.SetDataSource RstSub2
    rpt.OpenSubreport("SUBREP3").Database.SetDataSource RstSub3
    
    Set Rst2 = New ADODB.Recordset
    Rst2.CursorLocation = adUseClient
    Rst2.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic
    
    For I = 1 To rpt.FormulaFields.Count
        Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
            Case UCase("SubTitle")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecSpeciality & "'"
            Case UCase("LST")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecLST & "'"
            Case UCase("LSTDate")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecLST_Date & "'"
            Case UCase("CST")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecCST & "'"
            Case UCase("CSTDate")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecCST_Date & "'"
            Case UCase("Phone")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecPhone & "'"
            Case UCase("Fax")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecFax & "'"
            Case UCase("Gram")
                rpt.FormulaFields(I).TEXT = "'" & Rst2!V_SecGram & "'"
        End Select
    Next
    
    If Index = PDos Then
        For I = 1 To rpt.OpenSubreport("SUBREP1").FormulaFields.Count
            Select Case UCase(rpt.OpenSubreport("SUBREP1").FormulaFields(I).FormulaFieldName)
                Case UCase("SpeedPrint")
                    rpt.OpenSubreport("SUBREP1").FormulaFields(I).TEXT = "'1'"
            End Select
        Next
        For I = 1 To rpt.OpenSubreport("SUBREP2").FormulaFields.Count
            Select Case UCase(rpt.OpenSubreport("SUBREP2").FormulaFields(I).FormulaFieldName)
                Case UCase("SpeedPrint")
                    rpt.OpenSubreport("SUBREP2").FormulaFields(I).TEXT = "'1'"
            End Select
        Next
        For I = 1 To rpt.OpenSubreport("SUBREP3").FormulaFields.Count
            Select Case UCase(rpt.OpenSubreport("SUBREP3").FormulaFields(I).FormulaFieldName)
                Case UCase("SpeedPrint")
                    rpt.OpenSubreport("SUBREP3").FormulaFields(I).TEXT = "'1'"
            End Select
        Next
    End If
        
    rpt.ReadRecords

    Select Case Index
        Case PWindows  'Printer
            For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("comp_name")
                    rpt.FormulaFields(I).TEXT = "'" & PubComp_Name & "'"
                Case UCase("comp_add1")
                    rpt.FormulaFields(I).TEXT = "'" & PubComp_Add & "'"
                Case UCase("comp_add2")
                    rpt.FormulaFields(I).TEXT = "'" & PubComp_Add2 & "'"
                Case UCase("comp_city")
                    rpt.FormulaFields(I).TEXT = "'" & PubComp_City & "'"
                Case UCase("Title")
                    rpt.FormulaFields(I).TEXT = "'" & Me.CAPTION & "'"
            End Select
            Next
            rpt.PrintOut False
        Case PScreen  'screen
            Call Report_View(rpt, Me.CAPTION, , True)
        Case PDos
            Call Report_View(rpt, Me.CAPTION, 1)
End Select
Set Rst = Nothing
Set Rst2 = Nothing
Set RST3 = Nothing
Set rpt = Nothing
CmdPrint(PSetUp).Tag = ""
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub PrinerSetUp()
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.PrinterSetup (0)
CmdPrint(PSetUp).Tag = "1"
LblPrinter.CAPTION = rpt.PrinterName
End Sub

Private Sub SpeedPrint()
'On Error GoTo ELoop
''Paper Size 8.5*12
''Total Lines Per PAge 72
''Top Margin  3 Lines  (For 1/2 Inch)
''Header 15 Lines
''Footer 23 Lines
''Bottom Margin  3 Lines  (For 1/2 Inch)
''Contd. Remarks 2 Lines
''Gate Pass Detail 8 Lines
''Print Area 18
'    Dim i As Integer, j As Integer, mQRY As String
'    Dim PrintStr As String
'    Dim Rst As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstDel As ADODB.Recordset
'    Dim Page As Byte, mLine As Byte, mFix As Byte
'    Dim mSlNo As Integer, PageWidth As Byte, PageLength As Integer
'    Dim mDocStr$, mDupStr$, Speciality$
'    Dim FooterCnt As Byte, mHeader As Byte, mFooter As Byte
'    Dim SubTot As Double, RstInvDet As ADODB.Recordset
'    Dim Fob As New FileSystemObject
'    Dim mJuriCity As String
'    Dim Cnt As Byte, mAmt As Double, PrnStr As String, PrnStr1 As String
'    Dim Left1 As String, Left2 As String, Left3 As String
'    Dim Left4 As String, Left5 As String, Left6 As String, Left7 As String
'    Dim Right1 As String, Right2 As String, Right3 As String
'    Dim Right4 As String, Right5 As String, Right6 As String, Right7 As String
'    Dim NetAmt As Double
'
'    Set RstDel = GCn.Execute("SELECT Veh_Order.Inv_DocId," & _
'        " subgroup.FPrefix,subgroup.FName,veh_order.DelChPrn_YN,veh_order.FIN_AMT,Veh_Purch1.Tot_Amount, " & _
'        "City_1.CityName AS fincity, ContractFinance.Add1 AS finadd1, ContractFinance.Add2 AS finadd2, FinBank.FinBankName, ContractFinance.FinName,  City.CityName, Veh_Order.DelCh_UName, Veh_Order.DelCh_UEntDt, Veh_Order.DelCh_No, Veh_Order.DelCh_DT, Veh_Order.Fund_Source,  Model.TYRES, Veh_Order.MODEL, Veh_Order.DelCh_SiteCode, Model.RIMS,  Veh_Stock.ChassisNo, Veh_Stock.EngineNo," & _
'        " Veh_Order.Ord_No, Veh_Order.Ord_Date, Model.Model_Desc, Model.Model_Desc1, ColMast.Col_Desc, SubGroup.Name, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, SubGroup.PIN, ContractFinance.PinCode " & _
'        " FROM (((((((((Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) LEFT JOIN TaxForms ON Veh_Order.Form_Code = TaxForms.Form_Code) LEFT JOIN ColMast ON Veh_Stock.Colour_Code = ColMast.Col_Code) " & _
'        " LEFT JOIN Model ON Veh_Order.MODEL = Model.MODEL) LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) " & _
'        "LEFT JOIN City ON SubGroup.CityCode = City.CityCode) LEFT JOIN ContractFinance ON Veh_Order.FB_CODE = ContractFinance.FinCode) LEFT JOIN FinBank ON ContractFinance.FinBankCode = FinBank.FinBankCode) LEFT JOIN City AS City_1 ON ContractFinance.City = City_1.CityCode) LEFT JOIN Veh_Purch1 ON Veh_Stock.Pur_DocId = Veh_Purch1.DocID " & _
'        " where Veh_Order.DelCh_DocId = '" & Master!SearchCode & "'")
'
'    If RstDel.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.Caption: Exit Sub
'    If Fob.FileExists("C:\RepPrint.Txt") = False Then
'        Fob.CreateTextFile ("C:\RepPrint.Txt")
'    End If
'    If Fob.FileExists("C:\RepPrint.Bat") = False Then
'        Fob.CreateTextFile ("C:\RepPrint.Bat")
'    End If
'    Open "C:\RepPrint.Txt" For Output As #1
'
'    PageLength = PubPageLength
'    PageWidth = 80
'    mHeader = 0   'Ideal 17
'    mFooter = 9
'
'    ' Header
'
'    mDocStr = IIf(RstDel!DelChPrn_YN = 0, "Vehicle Delivery Order", "Vehicle Delivery Order (Duplicate)")
'    mDupStr = ""
'
'      Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")
'
'        Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
'        mHeader = mHeader + 1
'         If XNull(RstCompDet!V_SecSpeciality) <> "" Then
'             Print #1, PRN_TIT(RstCompDet!V_SecSpeciality, "C", PageWidth)
'             mHeader = mHeader + 1
'         End If
'         Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
'         mHeader = mHeader + 1
'
'         If PubComp_Add2 <> "" Or PubComp_City <> "" Then
'             Print #1, PRN_TIT(PubComp_Add2 & IIf(PubComp_Add2 = "" Or PubComp_City = "", "", ",") & PubComp_City, "C", PageWidth)
'             mHeader = mHeader + 1
'         End If
'         Print #1, PRN_TIT(IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") + XNull(RstCompDet!V_SecPhone) + IIf(XNull(RstCompDet!V_SecFax) = "", "", " Fax   : ") + XNull(RstCompDet!V_SecFax), "C", PageWidth)
'         mHeader = mHeader + 1
'         Print #1, PSTR(XNull(RstCompDet!V_SecCST) + IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " + str(RstCompDet!V_SecCST_Date)), 40) + PSTR(XNull(RstCompDet!V_SecLST) + IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " + str(RstCompDet!V_SecLST_Date)), 40, , AlignRight)
'         mHeader = mHeader + 1
'
'        Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth) + mChr18 + mEmph
'        mHeader = mHeader + 1
'
' '0 -Hypothecation ,1- Hire purchase ,2 -Own Fund,3- Lease
'
'    If RstDel!Fund_Source = 0 Then   'Hypothecation
'        Left1 = "To,"
'        Left2 = RstDel!Name
'        Left3 = XNull(RstDel!FPrefix) + " " + XNull(RstDel!FName)
'        Left4 = XNull(RstDel!Add1)
'        Left5 = XNull(RstDel!Add2)
'        Left6 = XNull(RstDel!Add3) + IIf(XNull(RstDel!CityName) = "" Or XNull(RstDel!Add3) = "", "", ",") + XNull(RstDel!CityName)
'
'        Right1 = "Under Hypothecation to  "
'        Right2 = RstDel!FinBankName
'        Right3 = XNull(RstDel!FinAdd1)
'        Right4 = XNull(RstDel!FinAdd2)
'        Right5 = XNull(RstDel!FinCity)
'        Right6 = "Finance Amount :" + Format(RstDel!FIN_AMT, "0.00")
'
'    ElseIf RstDel!Fund_Source = 1 Then  'Hire Purchase
'        Left1 = "Sold to under HPA with, "
'        Left2 = " U/F " & RstDel!FinBankName
'        Left3 = XNull(RstDel!FinAdd1)
'        Left4 = XNull(RstDel!FinAdd2)
'        Left5 = XNull(RstDel!City)
'        Left6 = ""
'
'        Right1 = "Delivered to Hirer, "
'        Right2 = RstDel!Name
'        Right3 = XNull(RstDel!FPrefix) + " " + XNull(RstDel!FName)
'        Right4 = XNull(RstDel!Add1)
'        Right5 = XNull(RstDel!Add2)
'        Right6 = XNull(RstDel!Add3) + IIf(XNull(RstDel!CityName) = "" Or XNull(RstDel!Add3), "", ",") + XNull(RstDel!CityName)
'
'    ElseIf RstDel!Fund_Source = 3 Then 'Lease
'        Left1 = "To, "
'        Left2 = XNull(RstDel!Name)
'        Left3 = XNull(RstDel!FPrefix) + " " + XNull(RstDel!FName)
'        Left4 = XNull(RstDel!Add1)
'        Left5 = XNull(RstDel!Add2)
'        Left6 = XNull(RstDel!Add3) + IIf(XNull(RstDel!CityName) = "" Or XNull(RstDel!Add3), "", ",") + XNull(RstDel!CityName)
'
'        Right1 = "Leaser  "
'        Right2 = XNull(RstDel!FinBankName)
'        Right3 = XNull(RstDel!FinAdd1)
'        Right4 = XNull(RstDel!FinAdd2)
'        Right5 = XNull(RstDel!City)
'        Right6 = "Lease Amount :" + RstDel!FIN_AMT
'    Else
'        Left1 = "Sold To,"
'        Left2 = XNull(RstDel!Name)
'        Left3 = XNull(RstDel!FPrefix) + " " + XNull(RstDel!FName)
'        Left4 = XNull(RstDel!Add1)
'        Left5 = XNull(RstDel!Add2)
'        Left6 = XNull(RstDel!Add3) + IIf(XNull(RstDel!CityName) = "" Or XNull(RstDel!Add3), "", ",") + XNull(RstDel!CityName)
'    End If
'
'        Print #1, mChr18 + mEmph + PSTR(Left1, 40) + PSTR(Right1, 40) + mEmph1
'        mHeader = mHeader + 1
'        Print #1, PSTR(Left2, 40) + PSTR(Right2, 40)
'        mHeader = mHeader + 1
'        Print #1, PSTR(Left3, 40) + PSTR(Right3, 40)
'        mHeader = mHeader + 1
'        Print #1, PSTR(Left4, 40) + PSTR(Right4, 40)
'        mHeader = mHeader + 1
'        Print #1, PSTR(Left5, 40) + PSTR(Right5, 40)
'        mHeader = mHeader + 1
'        Print #1, PSTR(Left6, 40) + PSTR(Right6, 40)
'        mHeader = mHeader + 1
'
'        Set RstInvDet = GCn.Execute("select SupInvOnVehSaleInv , TaxDetOnVehInv, VehSaleInv_Prefix from syctrl")
'
'        Print #1, PSTR("Booking No.  : " + str(RstDel!ord_no), 40) + mEmph + "Delivery Order No. : " + " " + PSTR(str(RstDel!DelCh_No), 8, , AlignLeft) + mEmph1
'        mHeader = mHeader + 1
'        Print #1, PSTR("Booking Date : " + str(RstDel!ord_date), 40) + mEmph + "Delivery Order Date : " + str(RstDel!DelCh_Dt) + mEmph1
'        mHeader = mHeader + 1
'
'        Print #1, Replace(Space(PageWidth), " ", "-")
'        mHeader = mHeader + 1
'
'        Print #1, PSTR("Model", 15) + " : " + RstDel!Model_Desc
'        mHeader = mHeader + 1
'        Print #1, RstDel!Model_Desc1
'        mHeader = mHeader + 1
'        Print #1, PSTR("Colour", 15) + " : " + RstDel!Col_Desc
'        mHeader = mHeader + 1
'        Print #1, PSTR("Chassis No.", 15) + " : " + RstDel!ChassisNo
'        mHeader = mHeader + 1
'        Print #1, PSTR("Engine No.", 15) + " : " + RstDel!EngineNo
'        mHeader = mHeader + 1
'
'        Print #1, "Battery Perticulars : Fitted with 12 volt Battery : Make          No."
'        mHeader = mHeader + 1
'        Print #1, PRN_TIT("List Of Documents Supplied With The Chassis", "C", PageWidth)
'        mHeader = mHeader + 1
'
'        Print #1, Replace(Space(PageWidth), " ", "-")
'        mHeader = mHeader + 1
'
'        Print #1, PSTR("Description", 30) + PSTR("Quantity", 10, , AlignRight) + PSTR("Description", 30) + PSTR("Quantity", 10, , AlignRight)
'        mHeader = mHeader + 1
'        Print #1, PSTR("1.Vehicle Defect Report Form", 30) + Space(10) + PSTR("5.Swich Key", 30)
'        mHeader = mHeader + 1
'        Print #1, PSTR("2.Operator's Service Book", 30) + Space(10) + PSTR("6.Wiper Motor Assy Set", 30)
'        mHeader = mHeader + 1
'        Print #1, PSTR("3.Battery Warranty Card", 30)
'        mHeader = mHeader + 1
'        Print #1, PSTR("4.Key Ring", 30)
'        mHeader = mHeader + 1
'
'        Set Rst = GCn.Execute("SELECT Veh_Order.Inv_DocId,Veh_Purch2.Trn_Type,Veh_Purch2.QTY, Veh_AMDModel.Prod_Name " & _
'        "FROM (Veh_Order LEFT JOIN Veh_Stock ON Veh_Order.Inv_DocId = Veh_Stock.Sal_DocId) LEFT JOIN (Veh_Purch2 LEFT JOIN Veh_AMDModel ON Veh_Purch2.PROD_CODE = Veh_AMDModel.Prod_Code) ON Veh_Stock.Pur_DocId = Veh_Purch2.DocID " & _
'        "where Veh_Order.DelCh_DocId = '" & Master!SearchCode & "'")
'
'        If Rst.RecordCount  > 0 Then
'            Print #1, mEmph + "Shortage :  " + mEmph1
'            mHeader = mHeader + 1
'            Print #1, mDoub + PSTR("ItemName", 52) + PSTR("Qty", 13, , AlignRight) + mDoub1
'            mHeader = mHeader + 1
'            Do Until Rst.EOF
'                Print #1, PSTR(Rst!Prod_Name, 52) + PSTR(Rst!Qty, 13, 2)
'                mHeader = mHeader + 1
'                Rst.MoveNext
'            Loop
'
'            Print #1, Replace(Space(PageWidth), " ", "-")
'            mHeader = mHeader + 1
'        End If
'
'        Do Until mHeader  >= PageLength - mFooter
'            Print #1, ""
'            mHeader = mHeader + 1
'        Loop
'
'        Print #1, "E. & OE." + mEmph + PSTR("For " + PubComp_Name, PageWidth - 8, , AlignRight) + mEmph1
'        Print #1, ""
'        Print #1, ""
'        Print #1, PSTR("Authorised Signatory", PageWidth, , AlignRight)
'
'        Print #1, "Received Tata Diesel Chassis as detailed above in satisfactory order & good condition."
'        Print #1, ""
'        Print #1, "Signature Of Customer"
'        Print #1, Replace(Space(PageWidth), " ", "-") + mChr17
'
'        Print #1, mChr17 + RstDel!DelCh_UName + " " + str(RstDel!DelCh_UEntDt) + Space(((PageWidth * 1.7) - Len("") - Len(RstDel!DelCh_UName + " " + str(RstDel!DelCh_UEntDt))) / 2) + "" + mChr18
'    Print #1, mEject
'    Close #1
'    Open "C:\RepPrint.Bat" For Output As #1
'    If Fob.FolderExists("c:\WinNt") Then
'        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
'    Else
'        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
'    End If
'    Close #1
'    If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
'        GCn.Execute "update veh_order set BillPrn_YN = 1  where veh_order.Inv_DocId = '" & Master!SearchCode & "'"
'    End If
'
''    Shell "C:\RepPrint.Bat", vbHide
'    Exit Sub
'ELoop:
'    Close #1: CheckError
'    'EOF Speed Printing Section
End Sub












