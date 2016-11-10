VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "topctl.ocx"
Begin VB.Form FrmFrMast 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Freight Chart Master"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame FrVCat 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   5280
      TabIndex        =   22
      Top             =   5760
      Visible         =   0   'False
      Width           =   4995
      Begin MSDataGridLib.DataGrid DGVCat 
         Height          =   3225
         Left            =   30
         TabIndex        =   23
         Top             =   330
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   5689
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   12632256
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   13504523
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
            DataField       =   "Code"
            Caption         =   ""
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
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               DividerStyle    =   0
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   3435.024
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Vehicle Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Index           =   0
         Left            =   30
         TabIndex        =   24
         Top             =   30
         Width           =   4935
      End
   End
   Begin VB.Frame FrCity 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3600
      Left            =   1065
      TabIndex        =   19
      Top             =   5130
      Visible         =   0   'False
      Width           =   4995
      Begin MSDataGridLib.DataGrid DGCity 
         Height          =   3225
         Left            =   45
         TabIndex        =   20
         Top             =   330
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   5689
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   12632256
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   13504523
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   0
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
            DataField       =   "CityCode"
            Caption         =   ""
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
            DataField       =   "CityName"
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               DividerStyle    =   0
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   3435.024
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   0
               Locked          =   -1  'True
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label LblHelp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "List of City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Index           =   1
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Width           =   4935
      End
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5025
      MaxLength       =   11
      TabIndex        =   6
      Top             =   2415
      Width           =   1365
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5025
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1065
      Width           =   3510
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00B7DBC8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   6135
      MaxLength       =   25
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4770
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5025
      MaxLength       =   12
      TabIndex        =   7
      Top             =   2685
      Width           =   1365
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   8235
      MaxLength       =   3
      TabIndex        =   5
      Top             =   2160
      Width           =   675
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5025
      MaxLength       =   11
      TabIndex        =   4
      Top             =   2145
      Width           =   1365
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5025
      MaxLength       =   11
      TabIndex        =   3
      Top             =   1875
      Width           =   1365
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5025
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1605
      Width           =   1365
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5025
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1335
      Width           =   3510
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGMain 
      Height          =   2280
      Left            =   2430
      TabIndex        =   8
      Top             =   3495
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   4022
      _Version        =   393216
      BackColor       =   14085097
      Cols            =   4
      BackColorFixed  =   128
      ForeColorFixed  =   65535
      BackColorSel    =   16711680
      BackColorBkg    =   13623520
      GridColor       =   16512
      AllowUserResizing=   3
      Appearance      =   0
      FormatString    =   "|Vechile Category  |Diesel Qty | Trip Time"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Shape Shape1 
      Height          =   2445
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   810
      Width           =   7275
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   5
      Left            =   3225
      TabIndex        =   18
      Top             =   2745
      Width           =   1035
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trip Factor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   3225
      TabIndex        =   17
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distance Charges"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   10
      Left            =   3225
      TabIndex        =   16
      Top             =   2475
      Width           =   1605
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distance Extra (Y/N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   6
      Left            =   6405
      TabIndex        =   15
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parchi Exp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   4
      Left            =   3225
      TabIndex        =   14
      Top             =   1935
      Width           =   960
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KMS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   3
      Left            =   3225
      TabIndex        =   13
      Top             =   1650
      Width           =   420
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation UpTo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   3225
      TabIndex        =   12
      Top             =   1380
      Width           =   1935
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   3195
      TabIndex        =   11
      Top             =   1110
      Width           =   1875
   End
End
Attribute VB_Name = "FrmFrMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Grid Variables
Dim GridKey As Integer
Dim ExitCtrl As Boolean
Dim TAddMode  As Boolean
Dim MyIndex As Integer
Dim ADDFLAG As Byte, mFlag As Byte
Dim Master As ADODB.Recordset, rstCityHelp As ADODB.Recordset, rstVCatHelp As ADODB.Recordset
Dim rsUserPerm As ADODB.Recordset
Private Const GVCat As Byte = 1
Private Const GVehCat As Byte = 2
Private Const GQty As Byte = 3
Private Const GTime As Byte = 4

'*************************
Private Const DesFrom As Byte = 0
Private Const DesUpTo As Byte = 1
Private Const Kms As Byte = 2
Private Const PExp As Byte = 3
Private Const TFact As Byte = 4
Private Const DisExt As Byte = 5
Private Const AddDist As Byte = 6
Private Const ATime As Byte = 7
Private IntNo As String
Private mAllDiv As Integer
Private Intcode As Long, mROW As Integer, mCol As Integer, mPROW As Integer
Private MyActCtrl As Object, CtrlFlag As Boolean, mType As String

Private Sub DGCITY_GotFocus()
   If rstCityHelp.RecordCount > 0 Then
        txt(MyIndex).TEXT = rstCityHelp!CityName
        txt(MyIndex).Tag = rstCityHelp!CityCode
        If FrCity.Visible = True Then FrCity.Visible = False
    txt(MyIndex).SetFocus
   End If
End Sub
Private Sub DGVCat_GotFocus()
   If rstVCatHelp.RecordCount > 0 Then
         FGMain.TextMatrix(FGMain.Row, GVCat) = rstVCatHelp!code
         FGMain.TextMatrix(FGMain.Row, GVehCat) = rstVCatHelp!Name
    If FrVCat.Visible = True Then FrVCat.Visible = False
   FGMain.SetFocus
   End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Me.WindowState = 2
    If CtrlFlag = True Then
        MyActCtrl.SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Errloop
    TopCtrl1.PrvKeyCode = KeyCode
    If KeyCode = vbKeyF2 Or KeyCode = vbKeyF3 Or KeyCode = vbKeyF4 Or _
        (KeyCode = 70 And Shift = 2) Or (KeyCode = 80 And Shift = 2) Or _
        (KeyCode = 83 And Shift = 2) Or KeyCode = vbKeyEscape Or _
        KeyCode = vbKeyF5 Or KeyCode = vbKeyF10 Or KeyCode = vbKeyHome Or _
        KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyEnd Then
        TopCtrl1.TopKey_Down KeyCode, Shift
    End If
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub Form_Load()
Dim i As Byte
'On Error GoTo Errloop
DGCity.Columns(0).width = 0
DGVCat.Columns(0).width = 0
    TopCtrl1.Tag = "AEDP"
    Call GridIni
    
    '******************** Colour Setting
    
    For i = 0 To txt.Count - 1
        txt(i).BackColor = CtrlBColOrg
        txt(i).ForeColor = CtrlFColOrg
    Next
    FGMain.BackColor = CtrlBColOrg
    'FGMain.BackColorBkg = FrmBackCol
    TxtGrid(0).BackColor = CtrlBCol
    '************************************
    Set Master = CreateObject("ADODB.Recordset")
    Set rstCityHelp = CreateObject("ADODB.Recordset")
    Set rstVCatHelp = CreateObject("ADODB.Recordset")
    '//**********Open Master
    With Master
        .ActiveConnection = GCn
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Open "select Distinct FC.DesiFrom+FC.DesiUpTo as SearchCode,FC.*,City.CityName as FromCity,City1.CityName as UpToCity from ((FrChartMast as FC left join City on (City.CityCode=FC.DesiFrom)) left join City as City1 on (City1.CityCode=FC.DesiUpTo)) order by FC.DesiFrom,FC.DesiUpTo"
    End With

    With rstCityHelp
        .ActiveConnection = GCn
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Open "SELECT CityCode,CityName FROM City ORDER BY CityName"
    End With
    
    With rstVCatHelp
        .ActiveConnection = GCn
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Open "SELECT Code,Name FROM Veh_Grp ORDER BY Name"
    End With
    Disp_Text SETS("INI", Me, Master)
    MoveRec
    ADDFLAG = 0
    mFlag = 0
    Set DGCity.DataSource = rstCityHelp
    Set DGVCat.DataSource = rstVCatHelp
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Master = Nothing
    Set rstCityHelp = Nothing
    Set rstVCatHelp = Nothing
End Sub

Private Sub Disp_Text(enb As Boolean)
Dim i As Byte
         For i = 0 To txt.Count - 1
         txt(i).Enabled = enb
         Next
End Sub
Private Sub MakeBlank()
Dim i As Integer
    For i = 0 To txt.Count - 1
        txt(i).TEXT = ""
    Next
    FGMain.Rows = 1
    FGMain.AddItem FGMain.Rows
    FGMain.FixedRows = 1
End Sub
Private Sub MoveRec()
Dim Rst As ADODB.Recordset, i As Integer
'On Error GoTo Errloop
    Grid_Hide
    RST_BOF_EOF Master
    If Master.RecordCount <= 0 Then
        MakeBlank
    Else
        txt(DesFrom).Tag = XNull(Master!DesiFrom)
        txt(DesFrom).TEXT = XNull(Master!FromCity)
        txt(DesUpTo).Tag = XNull(Master!DesiUpTo)
        txt(DesUpTo).TEXT = XNull(Master!UpToCity)
        txt(Kms) = Format(VNull(Master!Kms), "0.00")
        txt(PExp) = Format(VNull(Master!Parchi), "0.00")
        txt(TFact) = Format(VNull(Master!TripFact), "0.00")
        txt(DisExt).TEXT = IIf(Master!DistExtra = 0, "No", "Yes")
        txt(DisExt).Tag = IIf(Master!DistExtra = 1, 1, 0)
        txt(AddDist) = Format(VNull(Master!AddDistChrg), "0.00")
        txt(ATime).TEXT = Format(XNull(Master!AlarmTime), "0")
    End If
        
    Set Rst = CreateObject("ADODB.Recordset")
        With Rst
            .ActiveConnection = GCn
            .CursorType = adOpenStatic
            .CursorLocation = adUseClient
            .Open "select FrChartMast1.VehCat,Veh_Grp.Name,FrChartMast1.DiselQty,FrChartMast1.TripTime from ((FrChartMast1 left join FrChartMast on (FrChartMast.DesiFrom=FrChartMast1.DesiFrom and FrChartMast.DesiUpTo=FrChartMast1.DesiUpTo)) left join Veh_Grp on Veh_Grp.Code=FrChartMast1.VehCat) where FrChartMast1.DesiFrom='" & txt(DesFrom).Tag & "' and FrChartMast1.DesiUpTo='" & txt(DesUpTo).Tag & "'"
        End With
        FGMain.Rows = 1
           If Rst.RecordCount > 0 Then
              i = 1
              While Not Rst.EOF
                FGMain.AddItem (i & Chr(9) & Rst!VehCat & Chr(9) & Rst!Name & Chr(9) & Format(VNull(Rst!DiselQty), "0.00") & Chr(9) & Rst!TripTime)
                i = i + 1
                Rst.MoveNext
              Wend
              FGMain.FixedRows = 1
            Else
              FGMain.AddItem FGMain.Rows
              FGMain.FixedRows = 1
           End If
           Set Rst = Nothing

    If TopCtrl1.Visible = True And TopCtrl1.Enabled = True Then TopCtrl1.SetFocus
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo Errloop
    MakeBlank
    ADDFLAG = 1
    Disp_Text SETS("ADD", Me, Master)
    'txt(POrdNo) = GCn.Execute("Select iif(IsNull(Max(val(Ord_No))),1,Max(val(Ord_No))+1)AS MyCode From POrd ").Fields(0)
    FGMain.Col = 2
    txt(DesFrom).SetFocus
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information"
End Sub
Private Sub TopCtrl1_eEdit()
On Error GoTo Errloop
    If Master.RecordCount > 0 Then
        ADDFLAG = 2
        Disp_Text SETS("EDIT", Me, Master)
        FGMain.Col = 2
        txt(DesFrom).SetFocus
    Else
        MsgBox "There Is No Record To Edit.", vbInformation, "Information"
    End If
    Exit Sub
Errloop:
    MsgBox err.Description, vbExclamation, " Editing Error "
End Sub
Private Sub TopCtrl1_eDel()
Dim transFalg As Byte
On Error GoTo Errloop
    If Master.RecordCount > 0 Then
          If MsgBox("Are You Sure to Delete This Record", vbYesNo, "Confirmation") = vbYes Then
            transFalg = 1
            GCn.BeginTrans
                GCn.Execute ("Delete From FrChartMast where DesiFrom ='" & txt(DesFrom).Tag & "' and DesiUpTo='" & txt(DesUpTo).Tag & "'")
                GCn.Execute ("Delete From FrChartMast1 where DesiFrom ='" & txt(DesFrom).Tag & "' and DesiUpTo='" & txt(DesUpTo).Tag & "'")
            GCn.CommitTrans
            transFalg = 0
            Master.Requery
            Disp_Text SETS("INI", Me, Master)
            MoveRec
        End If
    Else
        MsgBox "There Is No Record To Delete.", vbInformation, "Information"
    End If
    Exit Sub
Errloop:
    If transFalg = 1 Then
        GCn.RollbackTrans
        MsgBox err.Description, vbExclamation
    End If
End Sub
Private Sub TopCtrl1_eFirst()
    BUTTONS True, Me, Master, 1
    MoveRec
End Sub
Private Sub TopCtrl1_ePrev()
    BUTTONS True, Me, Master, 2
    MoveRec
End Sub
Private Sub TopCtrl1_eNext()
    BUTTONS True, Me, Master, 3
    MoveRec
End Sub
Private Sub TopCtrl1_eLast()
    BUTTONS True, Me, Master, 4
    MoveRec
End Sub
Private Sub TopCtrl1_eFind()
On Error GoTo ELoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    GSQL = "select Distinct FC.DesiFrom+FC.DesiUpTo as SearchCode,City.CityName as FromCity,City1.CityName as UpToCity from ((FrChartMast as FC left join City on City.CityCode=FC.DesiFrom) left join City as City1 on City1.CityCode=FC.DesiUpTo) "
    Set SearchForm = Me
    FIND.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub
Public Sub SEARCHBACK(ByVal MYVALUE As String)
On Error GoTo ErrorLoop
    Master.MoveFirst
    Master.FIND ("SearchCode='" & MYVALUE & "'")
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TopCtrl1_eSave()
Dim transFlag As Byte, MySql$, MySql1$, str As String, mPREFIX As String, i As Integer
Dim X, Y, Z, X1, Y1, Z1, mCOUNT, J
'On Error GoTo Errloop
    Ctrl_BckColor
    transFlag = 0
    If IsValid(txt(DesFrom), "Desitination From") = False Then Txt_GotFocus DesFrom: Exit Sub
    If IsValid(txt(DesUpTo), "Desitination UpTo") = False Then Txt_GotFocus DesUpTo: Exit Sub
    If ADDFLAG = 1 Then
        If GCn.Execute("Select DesiFrom&DesiUpto from FrChartMast where DesiFrom='" & txt(DesFrom).Tag & "' and DesiUpTo='" & txt(DesUpTo).Tag & "'").RecordCount > 0 Then
            MsgBox "Freight Already Added for These Locations.", vbInformation + vbOKOnly
            txt(DesFrom).SetFocus: Exit Sub
        End If
    End If
    For i = 1 To FGMain.Rows - 1
        X = UCase(Trim(FGMain.TextMatrix(i, GVCat)))
        Z = UCase(Trim(FGMain.TextMatrix(i, GTime)))
        Y = UCase(Trim(Val(FGMain.TextMatrix(i, GQty))))
        mCOUNT = 0
        For J = 1 To FGMain.Rows - 1
            If Len(FGMain.TextMatrix(i, GVehCat)) = 0 And Len(FGMain.TextMatrix(i, GTime)) = 0 And Val(FGMain.TextMatrix(i, GQty)) > 0 Then
                    MsgBox "Detail Required", vbInformation, "Validation"
                    FGMain.Row = i
                     FGMain.Col = GTime
                     FGMain.SetFocus
            Exit Sub
            End If
        X1 = UCase(Trim(FGMain.TextMatrix(J, GVehCat)))
        Y1 = UCase(Trim(Val(FGMain.TextMatrix(J, GQty))))
        Z1 = UCase(Trim(FGMain.TextMatrix(J, GTime)))
        If X = X1 And Y = Y1 And Z = Z1 Then mCOUNT = mCOUNT + 1
            If mCOUNT > 1 Then
                MsgBox "Duplicate Item ", vbInformation, "Validation At Save"
                FGMain.Row = i
                FGMain.Col = GTime
                FGMain.SetFocus
                Exit Sub
            End If
        Next
    Next

    GCn.BeginTrans
        transFlag = 1
        If ADDFLAG = 1 Then
               MySql = "Insert Into FrChartMast(DesiFrom,DesiUpTo,KMS,Parchi,TripFact,DistExtra,AddDistChrg,AlarmTime," & _
                        "U_Name,U_EntDt,U_AE)" & _
                        " values ('" & txt(DesFrom).Tag & "','" & txt(DesUpTo).Tag & "'," & Val(txt(Kms)) & "," & Val(txt(PExp)) & "," & Val(txt(TFact)) & "," & Val(txt(DisExt).Tag) & "," & Val(txt(AddDist)) & "," & _
                        "'" & txt(ATime) & "' ," & _
                        "'SA'," & ConvertDate(PubLoginDate) & ",'A')"
       Else
                MySql = "Update FrChartMast set DesiFrom='" & txt(DesFrom).Tag & "',DesiUpTo='" & txt(DesUpTo).Tag & "',KMS=" & Val(txt(Kms)) & ",Parchi=" & Val(txt(PExp)) & ",TripFact=" & Val(txt(TFact)) & ",DistExtra=" & txt(DisExt).Tag & ",AddDistChrg=" & Val(txt(AddDist)) & "," & _
                        "AlarmTime='" & txt(ATime) & "' ," & _
                        "U_Name='SA',U_EntDt=" & ConvertDate(PubLoginDate) & ",U_AE='E' where DesiFrom='" & txt(DesFrom).Tag & "' and DesiUpTo='" & txt(DesUpTo).Tag & "' "
    
        End If
        GCn.Execute (MySql)
        GCn.Execute ("Delete * from FrChartMast1 where DesiFrom='" & txt(DesFrom).Tag & "' and DesiUpTo='" & txt(DesUpTo).Tag & "'")
        For i = 1 To FGMain.Rows - 1
             If Val(FGMain.TextMatrix(i, GVCat)) <> 0 Then
                    MySql1 = "INSERT INTO FrChartMast1(DesiFrom,DesiUpTo,VehCat,DiselQty,TripTime,U_Name,U_EntDt,U_AE) " & _
                    " VALUES ('" & txt(DesFrom).Tag & "','" & txt(DesUpTo).Tag & "'," & Val(FGMain.TextMatrix(i, GVCat)) & "," & Val(FGMain.TextMatrix(i, GQty)) & ",'" & FGMain.TextMatrix(i, GTime) & "','" & pubUName & "'," & ConvertDate(PubLoginDate) & ",'A')"
                    GCn.Execute (MySql1)
            End If
        Next
        GCn.CommitTrans
        transFlag = 0
        Master.Requery
        Master.MoveFirst
        Master.FIND ("SearchCode='" & txt(DesFrom) & txt(DesUpTo) & "'")
        If ADDFLAG = 1 Then
            Call TopCtrl1_eAdd
        Else
            Disp_Text SETS("INI", Me, Master)
            MoveRec
            ADDFLAG = 0
            FrCity.Visible = False
        End If
    Exit Sub
Errloop:
    If transFlag = 1 Then GCn.RollbackTrans
        MsgBox err.Description, vbCritical
End Sub
Private Sub TopCtrl1_eCancel()
On Error GoTo Errloop
    If MsgBox("Are You Sure To Cancel Changes", vbYesNo, "Confirmation") = vbYes Then
        Ctrl_BckColor
        ADDFLAG = 0
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        FrCity.Visible = False
    Else
        Me.ActiveControl.SetFocus
    End If
    Exit Sub
Errloop:
    MsgBox err.Description, vbCritical, "Information"
End Sub
Private Sub TopCtrl1_eRef()
    rstCityHelp.Requery
End Sub
Private Sub TopCtrl1_eExit()
    Unload Me
End Sub
Private Sub Txt_GotFocus(Index As Integer)
Dim mBookMark
On Error GoTo Errloop
    Grid_Hide
    Call Ctrl_GetFocus(Index)
    RST_BOF_EOF rstCityHelp
    MyIndex = Index
    Select Case Index
        Case DesFrom
         MyIndex = Index
            If txt(Index).TEXT = "" Then Exit Sub
            If rstCityHelp.RecordCount = 0 Then Exit Sub
            rstCityHelp.MoveFirst
            rstCityHelp.FIND "CityName='" & txt(Index).TEXT & "'"
            If rstCityHelp.BOF Or rstCityHelp.EOF Then Exit Sub
        Case DesUpTo
         MyIndex = Index
            If txt(Index).TEXT = "" Then Exit Sub
            If rstCityHelp.RecordCount = 0 Then Exit Sub
            rstCityHelp.MoveFirst
            rstCityHelp.FIND "CityName='" & txt(Index).TEXT & "'"
            If rstCityHelp.BOF Or rstCityHelp.EOF Then Exit Sub
    End Select
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Txt_Click(Index As Integer)
    txt(Index).ForeColor = CtrlFCol: txt(Index).BackColor = CtrlBCol
End Sub
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Integer
On Error GoTo Errloop
    If TopCtrl1.TopText2 = "Browse" Then Exit Sub
        If KeyCode = vbKeyEscape Then
        Grid_Hide
        Exit Sub
       End If
       With FrCity
            .left = txt(Index).left
            .top = txt(Index).top + txt(Index).height + 10
       End With
          Select Case Index
            Case DesFrom
                  Set MyActCtrl = Me.ActiveControl: CtrlFlag = True
                  DGridTxtKeyDown FrCity, txt, Index, rstCityHelp, KeyCode, False, 1
                  If KeyCode = vbKeyReturn Then
                  Ctrl_DownKeyDown KeyCode, Shift
                   End If
                   
            Case DesUpTo
                  Set MyActCtrl = Me.ActiveControl: CtrlFlag = True
                  DGridTxtKeyDown FrCity, txt, Index, rstCityHelp, KeyCode, False, 1
                  If KeyCode = vbKeyReturn Then
                  Ctrl_DownKeyDown KeyCode, Shift
                   End If
                  
           End Select
           
     If FrCity.Visible = False Then
               If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then
                  Ctrl_DownKeyDown KeyCode, Shift
                End If
        If KeyCode = vbKeyUp Then
           If (TopCtrl1.TopText2 = "Add" And Index <> DesFrom) Or (TopCtrl1.TopText2 = "Edit" And Index <> DesFrom) Then
                Ctrl_UpKeyDown KeyCode, Shift
            End If
        End If
    End If
        Exit Sub
           
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CheckQuote(KeyAscii)
    Select Case Index
        Case DesFrom
            If DGCity.Visible = True Then DGridTxtKeyPress txt, Index, rstCityHelp, KeyAscii, "CityName"
        Case DesUpTo
            If DGCity.Visible = True Then DGridTxtKeyPress txt, Index, rstCityHelp, KeyAscii, "CityName"
        Case Kms
            NumPress txt(Index), KeyAscii, 8, 2
        Case PExp, AddDist
            NumPress txt(Index), KeyAscii, 8, 2
        Case TFact, ATime
            NumPress txt(Index), KeyAscii, 8, 2
        Case DisExt
              If KeyAscii <> vbKeyReturn Then
                  txt(AddDist).Enabled = True
                  If Asc("Y") = KeyAscii Or Asc("y") = KeyAscii Then     ' Y/y
                        txt(Index).TEXT = "Yes"
                        txt(Index).Tag = 1
                        KeyAscii = 0
                        txt(AddDist).Enabled = True
                  ElseIf Asc("N") = KeyAscii Or Asc("n") = KeyAscii Then     ' N/n
                        txt(Index).TEXT = "No"
                        txt(Index).Tag = 0
                        KeyAscii = 0
                        txt(AddDist).Enabled = False
                  End If
               ElseIf txt(DisExt) = "" Then
                    txt(AddDist).Enabled = False
               End If
    End Select
End Sub
Private Sub Txt_LostFocus(Index As Integer)
    Call Ctrl_validate(Index)
End Sub
Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Lrs As ADODB.Recordset
On Error GoTo Errloop
    Select Case Index
        Case DesFrom
            If txt(Index).TEXT <> "" And rstCityHelp.EOF = False And rstCityHelp.BOF = False Then
                txt(Index).Tag = rstCityHelp!CityCode
                txt(Index).TEXT = rstCityHelp!CityName
            Else
                txt(Index).Tag = ""
                txt(Index).TEXT = ""
            End If
        Case ATime
           FGMain.SetFocus
    End Select
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub Ctrl_GetFocus(Index As Integer)
    txt(Index).BackColor = CtrlBCol
    txt(Index).ForeColor = CtrlFCol
    txt(Index).BorderStyle = 1
End Sub
Private Sub Ctrl_validate(Index As Integer)
    txt(Index).BackColor = CtrlBColOrg
    txt(Index).ForeColor = CtrlFColOrg
    txt(Index).BorderStyle = 0
End Sub
Private Sub Ctrl_BckColor()
Dim i As Integer
    For i = 0 To txt.Count - 1
        txt(i).BackColor = CtrlBColOrg
        txt(i).BorderStyle = 0
    Next
End Sub
Private Sub Grid_Hide()
    If FrCity.Visible = True Then FrCity.Visible = False
End Sub

Private Sub GridIni()
On Error GoTo err1
    With FGMain
        .RowHeightMin = PubGridRowHeight
        .Cols = 5
        .ColWidth(0) = 200                           ' marker
        
        .TextMatrix(0, GVCat) = "Vehicle Category"     ' Product group name
        .ColWidth(GVCat) = 0

        .TextMatrix(0, GVehCat) = "Vehicle Category"     ' Product group name
        .ColAlignmentFixed(GVehCat) = flexAlignLeftCenter
        .ColWidth(GVehCat) = 2600

        .TextMatrix(0, GQty) = "Diesel Qty"          'Qty
        .ColAlignment(GQty) = flexAlignRightCenter
        .ColAlignmentFixed(GQty) = flexAlignCenterCenter
        .ColWidth(GQty) = 1000

         .TextMatrix(0, GTime) = "Trip Time(Hrs)"
         .ColAlignment(GTime) = flexAlignRightCenter
        .ColAlignmentFixed(GTime) = flexAlignRightCenter
        .ColAlignment(GTime) = flexAlignLeftCenter
        .ColWidth(GTime) = 2400

    End With
Exit Sub
err1:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TxtGrid_LostFocus(Index As Integer)
On Error GoTo Errloop
    If ExitCtrl = False Then Exit Sub
'       Ctrl_validate 0
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo Errloop
 Grid_Hide
 If Index = 0 Then
    Select Case FGMain.Col
        Case GVehCat
            TxtGrid(0).MaxLength = 35
        Case GQty
            TxtGrid(0).MaxLength = 8
        Case GTime
            TxtGrid(0).MaxLength = 8
    End Select
    FGMain.CellBackColor = CellBackColLeave
 End If
 Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Errloop
        
     
     With FrVCat
            .left = TxtGrid(Index).left
            .top = TxtGrid(Index).top + TxtGrid(Index).height + 10
       End With
 
    If Index = 0 Then
        If KeyCode = vbKeyEscape Then
            TxtGrid(Index).TEXT = TxtGrid(Index).Tag
            TxtGrid_KeyUp Index, KeyCode, Shift
            TxtGrid(Index).Visible = False
            FGMain.SetFocus
            Exit Sub
        End If
     Select Case FGMain.Col
        
        Case GVehCat
        Set MyActCtrl = Me.ActiveControl: CtrlFlag = True
            DGridTxtKeyDown FrVCat, TxtGrid, Index, rstVCatHelp, KeyCode, False, 1
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave(Index) = True Then
                    GridTxtDown FGMain, TxtGrid, Index, KeyCode, TAddMode, 3, 0
                Else
                    TxtGrid_LostFocus 0
                    TxtGrid(0).SetFocus
                End If
             End If
        
        Case GQty
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave(Index) = True Then
                    GridTxtDown FGMain, TxtGrid, Index, KeyCode, TAddMode, 3, 0
                Else
                    TxtGrid_LostFocus 0
                    TxtGrid(0).SetFocus
                End If
             End If
        Case GTime
            If KeyCode = vbKeyReturn Then
                If TxtGridLeave(Index) = True Then
                     GridTxtDown FGMain, TxtGrid, Index, KeyCode, TAddMode, 2
                Else
                    TxtGrid_LostFocus 0
                    TxtGrid(0).SetFocus
                End If
            End If
     End Select
    End If
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Errloop
Call CheckQuote(KeyAscii)
  If Index = 0 Then
    Select Case FGMain.Col
     Case GVehCat
            If DGVCat.Visible = True Then DGridTxtKeyPress TxtGrid, Index, rstVCatHelp, KeyAscii, "Name"
    Case GQty
            Call NumPress(TxtGrid(Index), KeyAscii, 8, 2)
    End Select
  End If
  Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Errloop
    Select Case FGMain.Col
         Case GVehCat
            If KeyCode <> 13 And DGVCat.Visible = False Then TxtGrid_KeyDown Index, GridKey, 0: DGridTxtKeyPress TxtGrid, Index, rstVCatHelp, KeyCode, "Name", True
        Case GQty
            FGMain.TextMatrix(FGMain.Row, FGMain.Col) = Format(Val(TxtGrid(Index).TEXT), "0.00")
        Case GTime
            FGMain.TextMatrix(FGMain.Row, FGMain.Col) = Format(TxtGrid(Index).TEXT, "0")
'        Case GVehCat
'            FGMain.TextMatrix(FGMain.Row, FGMain.Col) = TxtGrid(Index).TEXT
        End Select
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
On Error GoTo errorbox
Dim J As Integer
If Index = 0 Then
    Select Case FGMain.Col
        Case GVehCat
            If rstVCatHelp.RecordCount <= 0 Or rstVCatHelp.EOF = True Or rstVCatHelp.BOF = True Then TxtGrid(Index).Visible = False: Exit Sub
            FGMain.TextMatrix(FGMain.Row, GVCat) = rstVCatHelp!code
            FGMain.TextMatrix(FGMain.Row, GVehCat) = rstVCatHelp!Name
            If FGMain.TextMatrix(FGMain.Rows - 1, 1) <> "" Then FGMain.AddItem FGMain.Rows
           
        Case GQty
            FGMain.TextMatrix(FGMain.Row, FGMain.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
'        Case GVehCat
'            FGMain.TextMatrix(FGMain.Row, FGMain.Col) = TxtGrid(0).TEXT
        Case GTime
            FGMain.TextMatrix(FGMain.Row, FGMain.Col) = Format(TxtGrid(0).TEXT, "0")
    End Select
    FGMain.SetFocus
    TxtGrid(Index).Visible = False
End If
errorbox:
    If err.NUMBER > 0 Then
        MsgBox err.Description, vbInformation
    End If
End Sub

Private Function TxtGridLeave(Index As Integer) As Boolean
Dim J As Integer
On Error GoTo Errloop
If Index = 0 Then
    Select Case FGMain.Col
    Case GVehCat
            If rstVCatHelp.RecordCount = 0 Or rstVCatHelp.EOF = True Or rstVCatHelp.BOF = True Then
                TxtGridLeave = False: ExitCtrl = False: Exit Function
            End If
            FGMain.TextMatrix(FGMain.Row, GVCat) = rstVCatHelp!code
            FGMain.TextMatrix(FGMain.Row, GVehCat) = rstVCatHelp!Name
                        '********************************************
            '********************************************
            If FGMain.TextMatrix(FGMain.Rows - 1, 1) <> "" Then FGMain.AddItem FGMain.Rows
     
        Case GQty
            FGMain.TextMatrix(FGMain.Row, FGMain.Col) = Format(Val(TxtGrid(0).TEXT), "0.00")
'        Case GVehCat
'            FGMain.TextMatrix(FGMain.Row, FGMain.Col) = TxtGrid(0).TEXT
        Case GTime
            FGMain.TextMatrix(FGMain.Row, FGMain.Col) = Format(TxtGrid(0).TEXT, "0")
    End Select
    ExitCtrl = True
    TxtGridLeave = True
    FGMain.SetFocus
    TxtGrid(Index).Visible = False
End If
Exit Function
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Function
End Function
Private Sub FGMain_Click()
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If TxtGrid(0).Visible = True Then TxtGrid(0).Visible = False
End Sub
Private Sub FGMain_EnterCell()
    FGMain.CellBackColor = CellBackColEnter
End Sub
Private Sub FGMain_GotFocus()
    FGMain.CellBackColor = CellBackColEnter
End Sub
Private Sub FGMain_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Errloop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGMain.Tag) = (FGMain.Rows - (FGMain.Rows - 1)) Then
            FGMain.CellBackColor = CellBackColLeave
            SendKeys "+{Tab}"
            KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGMain.Tag) = FGMain.Rows - 1 Then
        FGMain.CellBackColor = CellBackColLeave
            If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
                TopCtrl1_eSave
            Else
                FGMain.CellBackColor = CellBackColEnter
                Me.ActiveControl.SetFocus
            End If
    End If
    GridKey = KeyCode
    FGMain.Tag = FGMain.Row
    KeyCode = 0
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub FGMain_KeyPress(KeyAscii As Integer)
On Error GoTo Errloop
    Select Case FGMain.Col
        Case GVehCat, GTime
         ' *********  if text
           Call Get_Text(Me, FGMain, TxtGrid, 0, False, KeyAscii)
        Case GQty
         ' *********  if number
           Call Get_Text(Me, FGMain, TxtGrid, 0, True, KeyAscii)
    End Select
    If KeyAscii <> vbKeyReturn Then TAddMode = True
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub FGMain_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer
On Error GoTo Errloop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGMain.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGMain.Row >= 1 Then
            '*****************
                If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                    If FGMain.Rows > 2 Then
                        FGMain.RemoveItem (FGMain.Row)
                    Else
                        FGMain.Rows = 1
                        FGMain.AddItem FGMain.Rows
                        FGMain.FixedRows = 1
                    End If
                End If
            '*****************
            For i = 1 To FGMain.Rows - 1
                FGMain.TextMatrix(i, 0) = i
            Next
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGMain.SetFocus
    End If
    Exit Sub
Errloop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub
Private Sub FGMain_Scroll()
    TxtGrid(0).Visible = False
 End Sub
Private Sub FGMain_LeaveCell()
    FGMain.CellBackColor = CellBackColLeave
End Sub
Private Sub FGMain_Validate(Cancel As Boolean)
    FGMain.CellBackColor = CellBackColLeave
End Sub

