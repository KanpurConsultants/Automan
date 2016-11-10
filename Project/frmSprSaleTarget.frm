VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmSprSaleTarget 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Spare Sale Target"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11820
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11820
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDF4B5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   6840
      TabIndex        =   29
      Top             =   390
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Txt 
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
      Index           =   7
      Left            =   4425
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2805
      Width           =   1665
   End
   Begin VB.TextBox Txt 
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
      Index           =   6
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2805
      Width           =   1665
   End
   Begin VB.TextBox Txt 
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
      Index           =   5
      Left            =   4425
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.TextBox Txt 
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
      Index           =   4
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2535
      Width           =   1665
   End
   Begin VB.TextBox Txt 
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
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2265
      Width           =   4830
   End
   Begin VB.TextBox Txt 
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
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1995
      Width           =   4830
   End
   Begin VB.TextBox Txt 
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
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1725
      Width           =   4830
   End
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   2865
      Left            =   450
      Negotiate       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3810
      Visible         =   0   'False
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   5054
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Code"
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
         Caption         =   "A/c Name"
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
            DividerStyle    =   3
            ColumnWidth     =   30.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4710.047
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
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
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1455
      Width           =   4830
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   661
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   4455
      Left            =   6285
      TabIndex        =   9
      Top             =   1455
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   7858
      _Version        =   393216
      BackColor       =   14940925
      Rows            =   3
      Cols            =   5
      FixedRows       =   2
      BackColorFixed  =   15259902
      ForeColorFixed  =   8388736
      BackColorSel    =   15261111
      BackColorBkg    =   12243913
      GridColor       =   16761087
      GridColorFixed  =   8421504
      FocusRect       =   0
      GridLinesFixed  =   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "S.No. |    Date From   |    Date To   | Target Amount |   Achievements"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cr Limit"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   22
      Left            =   3150
      TabIndex        =   28
      Top             =   2820
      Width           =   645
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
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   7
      Left            =   4305
      TabIndex        =   27
      Top             =   2820
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cr Days"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   15
      Left            =   120
      TabIndex        =   26
      Top             =   2820
      Width           =   660
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
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   5
      Left            =   1155
      TabIndex        =   25
      Top             =   2820
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curr. Balance"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   6
      Left            =   3150
      TabIndex        =   24
      Top             =   2550
      Visible         =   0   'False
      Width           =   1125
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
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   4
      Left            =   4305
      TabIndex        =   23
      Top             =   2550
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   22
      Top             =   2550
      Width           =   840
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
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   3
      Left            =   1140
      TabIndex        =   21
      Top             =   2550
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
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   2
      Left            =   1140
      TabIndex        =   20
      Top             =   1725
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   1725
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Achievements Rs. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   0
      Left            =   2730
      TabIndex        =   18
      Top             =   3525
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label LblAchieve 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   4620
      TabIndex        =   17
      Top             =   3540
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1455
      Width           =   960
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
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   1
      Left            =   1140
      TabIndex        =   14
      Top             =   1455
      Width           =   45
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFFF&
      Height          =   540
      Left            =   8445
      Shape           =   4  'Rounded Rectangle
      Top             =   645
      Width           =   3240
   End
   Begin VB.Label LblSite 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   270
      Left            =   10350
      TabIndex        =   13
      Top             =   825
      Width           =   810
   End
   Begin VB.Label LblDiv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   270
      Left            =   8610
      TabIndex        =   12
      Top             =   825
      Width           =   660
   End
   Begin VB.Label LblTgtAmt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   4620
      TabIndex        =   11
      Top             =   3240
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total >>    Target Amount :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Index           =   25
      Left            =   2220
      TabIndex        =   10
      Top             =   3225
      Width           =   2205
   End
End
Attribute VB_Name = "frmSprSaleTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim RsParty As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim mSearchCode As String

'grid color scheme
Private Const CellBackColLeave As String = &HE3FAFD
Private Const CellForeColLeave As String = &HFF00FF
Private Const CellBackColEnter As String = &HCAF1FD
Private Const GridBackColorBkg As String = &HD7C6C8    ' me.backColor=&HB9D8EE

Private Const Party As Byte = 0                 ' A/c Name
Private Const Add1 As Byte = 1              ' Address1
Private Const Add2 As Byte = 2              ' Address2
Private Const Add3 As Byte = 3              ' Address2
Private Const PartyType As Byte = 4             ' Party Type
Private Const CurrBal As Byte = 5               ' Current Balance
Private Const CrDays As Byte = 6                ' Cr Days
Private Const CrLimit As Byte = 7               ' Cr Limit

'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_DateFrom As Byte = 1          ' DateFrom
Private Const Col_DateTo As Byte = 2            ' DateTo
Private Const Col_Target As Byte = 3            ' Target
Private Const Col_Achievements As Byte = 4      ' Achievements

Private Sub Disp_Text(Enb As Boolean)
    txt(Party).Enabled = Enb
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ELoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("SearchCode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select Distinct Party_Code As SearchCode,SG.Name " _
            & "From SaleTarget Left Join SubGroup SG on SaleTarget.Party_Code=SG.SubCode " & _
            " Where Party_Code  = '" & MyValue & "' " _
            & "Order by SG.Name")
    End If
    MoveRec
    BUTTONS True, Me, Master, 0
Exit Sub
ELoop:
    CheckError
End Sub
'* Used for clear all text boxes used in the form
Private Sub BlankText()
Dim I As Integer
    For I = 0 To txt.Count - 1
        txt(I).TEXT = ""
    Next I
    LblDiv.CAPTION = "Division : "
    LblSite.CAPTION = "Site Code : "
    LblTgtAmt.CAPTION = ""
    LblAchieve.CAPTION = ""

    FGrid.Rows = 2
    FGrid.AddItem FGrid.Rows - 1
    FGrid.FixedRows = 2
End Sub

'* Used for intialize grid columns
Private Sub Grid_Ini()
    With FGrid
        .left = 6285    'Me.left '+ 60
'        .width = Me.width - 90
        .top = 1455
        .BackColor = CellBackColLeave
        .BackColorBkg = GridBackColorBkg
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 5
        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        .MergeRow(0) = True

        .TextMatrix(0, Col_SrNo) = "S.No."
        .TextMatrix(1, Col_SrNo) = "S.No."
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 570 '450

'        .MergeCol(Col_DateFrom) = True
        .TextMatrix(0, Col_DateFrom) = "Period"
        .TextMatrix(1, Col_DateFrom) = "From"
        .ColWidth(Col_DateFrom) = 1470 '1260  '1095
        .ColAlignmentFixed(Col_DateFrom) = flexAlignCenterCenter
        
        .TextMatrix(0, Col_DateTo) = .TextMatrix(0, Col_DateFrom)
        .TextMatrix(1, Col_DateTo) = "To"
        .ColWidth(Col_DateTo) = 1470 '1260    '1095
        .ColAlignmentFixed(Col_DateTo) = flexAlignCenterCenter
        
        .MergeCol(Col_Target) = True
        .TextMatrix(0, Col_Target) = "Target"
        .TextMatrix(1, Col_Target) = "Target"
        .ColWidth(Col_Target) = 1530 '1275    '960
        .ColAlignmentFixed(Col_Target) = flexAlignRightCenter
        .ColAlignment(Col_Target) = flexAlignRightCenter
        
        .TextMatrix(0, Col_Achievements) = "Achieved"
        .TextMatrix(1, Col_Achievements) = "Amount"
        .ColWidth(Col_Achievements) = 0 '900
        .ColAlignmentFixed(Col_Achievements) = flexAlignRightCenter
        .ColAlignment(Col_Achievements) = flexAlignRightCenter
        .Rows = 3
        .FixedRows = 2
        .AddItem FGrid.Rows - 2
    End With
    DGParty.height = Me.height - (mTopScale + mBotScale)
    DGParty.left = Me.width - (DGParty.width + mRtScale): DGParty.top = mTopScale
End Sub

Private Sub Grid_Hide()
    If DGParty.Visible = True Then DGParty.Visible = False
End Sub

Private Sub MoveRec()
Dim Master1 As ADODB.Recordset, I As Integer, TmpStr As String
On Error GoTo ELoop
    If Master.RecordCount > 0 Then
        Set Master1 = New Recordset
        Master1.CursorLocation = adUseClient
        Master1.Open "Select SG.Name, Sg.Add1, Sg.Add2, Sg.Add3,City.CityName," _
            & "SG.Curr_Bal,SG.CreditDays,SG.CreditLimit,ST.Description,SaleTarget.* " _
            & "From (((SaleTarget Left Join SubGroup SG on SaleTarget.Party_Code=SG.SubCode) " _
            & "Left Join City on SG.CityCode=City.CityCode) " _
            & "Left Join SubGroupType ST on SG.Party_Type=ST.Party_Type) " _
            & "Where SaleTarget.Party_Code='" & Master!SearchCode & "' Order By SaleTarget.Srl_No", GCn, adOpenStatic, adLockReadOnly
        FGrid.Rows = 2
        If Master1.RecordCount > 0 Then
            txt(Party).TEXT = Master1!Name
            txt(Party).Tag = Master!SearchCode
            mSearchCode = Master!SearchCode
            LblDiv.CAPTION = "Division : " & Master1!Div_Code
            LblSite.CAPTION = "Site Code : " & Master1!Site_Code
            txt(Add1) = Master1!Add1
            txt(Add2) = Master1!Add2
            txt(Add3) = Master1!Add3
            txt(PartyType) = XNull(Master1!Description)
            txt(CurrBal) = Master1!Curr_Bal
            txt(CrDays) = Master1!CreditDays
            txt(CrLimit) = Master1!CreditLimit
            
            I = 1
            Do Until Master1.EOF
                FGrid.AddItem ""
                With FGrid
                    .TextMatrix(I + 1, Col_SrNo) = Master1!Srl_No
                    .TextMatrix(I + 1, Col_DateFrom) = Format(Master1!DateFrom, "dd/mmm/yyyy")
                    .TextMatrix(I + 1, Col_DateTo) = Format(Master1!DateTo, "dd/mmm/yyyy")
                    .TextMatrix(I + 1, Col_Target) = Format(Master1!TargetAmt, "0.00") 'Achievments
                    .TextMatrix(I + 1, Col_Achievements) = Format(Master1!Achievments, "0.00") 'Achievments
                End With
                Master1.MoveNext
                I = I + 1
            Loop
            
            FGrid.FixedRows = 2
        Else
            FGrid.AddItem FGrid.Rows - 1
            FGrid.FixedRows = 2
        End If
        Amt_Cal
    Else
        BlankText
    End If
    Grid_Hide
Set Master1 = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Function ChkDuplicate() As Boolean
Dim I As Integer, X As Date, Y As Date
    If txtgrid(0) <> "" Then
        X = RetDate(txtgrid(0))
        For I = 2 To FGrid.Rows - 1
            If I = FGrid.Row Then GoTo nxt1
            If FGrid.TextMatrix(I, Col_DateFrom) <> "" And FGrid.TextMatrix(I, Col_DateTo) <> "" Then
'            If FGrid.TextMatrix(i, FGrid.Col) <> "" Then
                Y = CDate(FGrid.TextMatrix(I, FGrid.Col))
                If X = Y Or _
                    (X >= CDate(FGrid.TextMatrix(I, Col_DateFrom)) And X <= CDate(FGrid.TextMatrix(I, Col_DateTo))) Then
                    MsgBox "Duplicate Period Date Not Allowed", vbInformation, "Validation"
                    txtgrid(0).SetFocus
                    ChkDuplicate = False
                    Exit Function
                End If
            End If
nxt1:
        Next
    End If
    ChkDuplicate = True
End Function

Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
    Select Case FGrid.Col
        Case Col_DateFrom
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            If txtgrid(0) <> "" And FGrid.TextMatrix(FGrid.Row, Col_DateTo) <> "" Then
                If CDate(RetDate(txtgrid(0))) > CDate(FGrid.TextMatrix(FGrid.Row, Col_DateTo)) Then
                    MsgBox "Date From is greater than Date To", vbOKOnly, "Validation"
                    TxtGridLeave = False: Exit Function
                End If
            End If
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(txtgrid(0))
        Case Col_DateTo
            If ChkDuplicate = False Then TxtGridLeave = False: Exit Function
            If txtgrid(0) <> "" And FGrid.TextMatrix(FGrid.Row, Col_DateFrom) <> "" Then
                If CDate(RetDate(txtgrid(0))) < CDate(FGrid.TextMatrix(FGrid.Row, Col_DateFrom)) Then
                    MsgBox "Date From is less than Date To", vbOKOnly, "Validation"
                    TxtGridLeave = False: Exit Function
                End If
            End If
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(txtgrid(0))
        Case Col_Target
            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(txtgrid(0).TEXT), "0.00")
            Amt_Cal
    End Select
    TxtGridLeave = True
    'Important at the time of validating  a control if you are making the visibility of
    'control false forcefully it will generate error
    If ValidateCall = False Then
        txtgrid(0).Visible = False
        FGrid.SetFocus
    End If
End Function
'* Used for Calculate the Amount
Private Sub Amt_Cal()
Dim I As Integer, TotGoodsVal As Double
    For I = 2 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_Target) <> "" Then
            TotGoodsVal = TotGoodsVal + Val(FGrid.TextMatrix(I, Col_Target))
        End If
    Next I
    LblTgtAmt.CAPTION = Format(TotGoodsVal, "0.00")
End Sub

Private Sub DGParty_Click()
    DGParty.Visible = False
    If RsParty.RecordCount > 0 Then
        txt(Party).TEXT = RsParty!Name
        txt(Party).Tag = RsParty!Code
        txt(Add1) = RsParty!Add1
        txt(Add2) = RsParty!Add2
        txt(Add3) = RsParty!Add3
        txt(PartyType) = RsParty!Description
        txt(CurrBal) = RsParty!Curr_Bal
        txt(CrDays) = RsParty!CreditDays
        txt(CrLimit) = RsParty!CreditLimit
    End If
    txt(Party).SetFocus
End Sub

Private Sub Form_Activate()
If TopCtrl1.PrvKeyCode = vbKeyInsert Then
        Call TopCtrl1_eRef
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    FormKeyDown Me, KeyCode, Shift
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Load()
On Error GoTo ELoop
Dim I As Byte
' To Change
'    PubRestrict_Godown = 0
'    PubSprWorksGodown = "RES"
'---
'to modify
    TopCtrl1.Tag = PubUParam: WinSetting Me: Grid_Ini
    For I = 0 To txt.Count - 1
        txt(I).BackColor = CtrlBColOrg '&HDFF4F2
        txt(I).ForeColor = CtrlFColOrg
'        Txt(I).BorderStyle = 1
    Next
    GSQL = "Select SG.SubCode as Code,SG.Name,Sg.Add1 , Sg.Add2, Sg.Add3, City.CityName, " & _
        " SG.Curr_Bal,SG.CreditDays,SG.CreditLimit,ST.Description " & _
        " From ((SubGroup SG left join " & FaTable("AcGroup") & " on SG.GroupCode=AcGroup.GroupCode) " & _
        " Left Join City on SG.CityCode=City.CityCode) " & _
        " Left Join SubGroupType ST on SG.Party_Type=ST.Party_Type " & _
        " Where  " & _
        " left(AcGroup.MainGrCode,6) in ('" & pubSundryCrSysMainGrCode & "','" & pubSundryDrSysMainGrCode & "') " & _
        " Order by SG.Name"
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open GSQL, GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
    If PubMoveRecYn Then
        Set Master = GCn.Execute("Select Distinct Party_Code As SearchCode,SG.Name " _
            & "From SaleTarget Left Join SubGroup SG on SaleTarget.Party_Code=SG.SubCode " _
            & "Order by SG.Name")
    Else
        Set Master = GCn.Execute("Select  Distinct Top 1 Party_Code As SearchCode,SG.Name " _
            & "From SaleTarget Left Join SubGroup SG on SaleTarget.Party_Code=SG.SubCode " _
            & "Order by SG.Name")
    End If

    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsParty = Nothing
    Set Master = Nothing
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ELoop
    BlankText
    Disp_Text SETS("ADD", Me, Master)
    LblDiv.CAPTION = "Division : " & PubDivCode
    LblSite.CAPTION = "Site Code : " & PubSiteCode
    txt(Party).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
    txt(Party).Enabled = False
    FGrid.AddItem FGrid.Rows - 1
    FGrid.SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
Dim vBook As Variant
    If Master.RecordCount > 0 Then
        If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            vBook = Master.AbsolutePosition
            GCn.BeginTrans
            GCn.Execute ("Delete From SaleTarget Where Party_Code='" & txt(Party).Tag & "'")
            GCn.CommitTrans
            Master.Requery
            If Master.RecordCount > 0 Then
                If Master.RecordCount >= vBook Then Master.AbsolutePosition = vBook Else Master.MoveLast
            End If
            BUTTONS True, Me, Master, 0
            MoveRec
        End If
    Else
        MsgBox "No Records To Delete!", vbInformation, "Information"
    End If
Exit Sub
ELoop:
    GCn.RollbackTrans
    CheckError
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
    GSQL = "Select Distinct Party_Code As SearchCode, Party_Code, SG.Name " _
        & "From SaleTarget Left Join SubGroup SG on SaleTarget.Party_Code=SG.SubCode " _
        & "Order by SG.Name"
    Set SearchForm = Me
    FIND.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eRef()
On Error GoTo ELoop
    RsParty.Requery
    Master.Requery
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean
Dim Rst As ADODB.Recordset, TmpStr As String
On Error GoTo ELoop
    If txtgrid(0).Visible = True Then
        If TxtGridLeave = False Then
            txtgrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide
    
    If IsValid(txt(Party), Label3(3)) = False Then Exit Sub
    If FGrid.Rows = 3 And FGrid.TextMatrix(2, Col_DateFrom) = "" Then MsgBox "Please Fill Target Details", vbInformation, "Validation": FGrid.Row = 1: FGrid.Col = Col_DateFrom: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
    For I = 2 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, Col_DateFrom) <> "" Then
            If FGrid.TextMatrix(I, Col_DateTo) = "" Then MsgBox "Please Specify To Date in S.No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_DateTo: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
            If FGrid.TextMatrix(I, Col_Target) = "" Then MsgBox "Please Specify Target in S.No. " & I, vbInformation, "Validation": FGrid.Row = I: FGrid.Col = Col_Target: FGrid.SetFocus: FGrid.CellBackColor = CellBackColEnter: Exit Sub
        End If
    Next
    
    GCn.BeginTrans
        mTrans = True
        If TopCtrl1.TopText2 = "Edit" Then
            GCn.Execute ("Delete From SaleTarget Where Party_Code='" & txt(Party).Tag & "'")
        End If
        For I = 2 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_DateFrom) <> "" And _
                FGrid.TextMatrix(I, Col_DateTo) <> "" And _
                Val(FGrid.TextMatrix(I, Col_Target)) <> 0 Then
                GCn.Execute "Insert Into SaleTarget(Party_Code,Div_Code,Site_Code," & _
                    " Srl_No,DateFrom,DateTo,TargetAmt,U_Name,U_EntDt,U_AE) " & _
                    " Values('" & txt(Party).Tag & "','" & PubDivCode & "','" & PubSiteCode & PubSiteCode & _
                    "'," & I - 1 & "," & ConvertDate(FGrid.TextMatrix(I, Col_DateFrom)) & "," & ConvertDate(FGrid.TextMatrix(I, Col_DateTo)) & _
                    "," & Val(FGrid.TextMatrix(I, Col_Target)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'" & left(TopCtrl1.TopText2, 1) & "')"
            End If
        Next
    GCn.CommitTrans
    mTrans = False
    mSearchCode = txt(Party).Tag
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select Distinct Party_Code As SearchCode,SG.Name " _
            & "From SaleTarget Left Join SubGroup SG on SaleTarget.Party_Code=SG.SubCode " & _
            " Where Party_Code  = '" & mSearchCode & "' " _
            & "Order by SG.Name")
    End If
    Master.FIND "SearchCode = '" & mSearchCode & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    If mTrans = True Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    Grid_Hide
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To txt.Count - 1
            txt(I).BackColor = CtrlBColOrg
            txt(I).ForeColor = CtrlFColOrg
        Next
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub Txt_GotFocus(Index As Integer)
On Error GoTo ELoop
Ctrl_GetFocus txt(Index)
txtgrid(0).Visible = False
Grid_Hide
Select Case Index
    Case Party
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or txt(Index).TEXT = "" Then Exit Sub
        If txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "Name ='" & txt(Index).TEXT & "'"
        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Select Case Index
    Case Party
        DGridTxtKeyDown DGParty, txt, Party, RsParty, KeyCode, False, 1, frmSubGroup, "frmSubGroup"
    End Select
    If DGParty.Visible = False Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
        If TopCtrl1.TopText2.CAPTION = "Add" Then
            If Index <> Party And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" Then
            If Index <> Party And KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
On Error GoTo ELoop
If keyascii = 39 Then keyascii = 0: Exit Sub
Select Case Index
    Case Party
        If DGParty.Visible = True Then DGridTxtKeyPress txt, Party, RsParty, keyascii, "Name"
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo ELoop
'Select Case Index
'    Case DocType, AdjType
'        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
'End Select
'Exit Sub
'ELoop:
'    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Select Case Index
    Case Party
        If RsParty.RecordCount > 0 Then
            If txt(Index).TEXT <> "" Then
                txt(Index).TEXT = RsParty!Name
                txt(Index).Tag = RsParty!Code
                txt(Add1) = RsParty!Add1
                txt(Add2) = RsParty!Add2
                txt(Add3) = RsParty!Add3
                txt(PartyType) = XNull(RsParty!Description)
                txt(CurrBal) = RsParty!Curr_Bal
                txt(CrDays) = RsParty!CreditDays
                txt(CrLimit) = RsParty!CreditLimit
            End If
        Else
            txt(Index).TEXT = ""
            txt(Index).Tag = ""
            txt(Add1) = ""
            txt(Add2) = ""
            txt(Add3) = ""
            txt(PartyType) = ""
            txt(CurrBal) = ""
            txt(CrDays) = ""
            txt(CrLimit) = ""
        End If
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
'    Ctrl_GetFocus TxtGrid(Index)
    Grid_Hide
    FGrid.CellBackColor = CellBackColLeave
    txtgrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
    txtgrid(0).Font = FGrid.Font
    Select Case FGrid.Col
'        Case Col_Godown
'            TxtGrid(0).MaxLength = 20
'            If RsGodown.RecordCount = 0 Or (RsGodown.EOF = True Or RsGodown.BOF = True) Or FGrid.TextMatrix(FGrid.Row, Col_Godown) = "" Then Exit Sub
'            If FGrid.TextMatrix(FGrid.Row, Col_Godown) <> RsGodown!Name Then
'                RsGodown.MoveFirst
'                RsGodown.FIND "Name ='" & FGrid.TextMatrix(FGrid.Row, Col_Godown) & "'"
'            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        txtgrid(0).TEXT = txtgrid(0).Tag
        TxtGrid_KeyUp Index, KeyCode, Shift
        txtgrid(0).Visible = False
        FGrid.SetFocus
        Exit Sub
    End If
    Select Case FGrid.Col
        Case Col_DateFrom, Col_DateTo, Col_Target
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid, txtgrid, Index, KeyCode, TAddMode, Col_Target '+ 1
                End If
            End If
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, keyascii As Integer)
On Error GoTo ELoop
    CheckQuote keyascii
    Select Case FGrid.Col
        Case Col_Target
            NumPress txtgrid(Index), keyascii, 8, 2
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case FGrid.Col
        Case Col_DateFrom, Col_DateTo
'            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = RetDate(TxtGrid(Index))
'            Amt_Cal
        Case Col_Target
'            FGrid.TextMatrix(FGrid.Row, FGrid.Col) = Format(Val(TxtGrid(Index)), "0.00")
'            Amt_Cal
    End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_Click()
    txtgrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    Select Case FGrid.Col
        Case Col_DateFrom, Col_DateTo, Col_Target
            GridDblClick Me, FGrid, txtgrid, 0
    End Select
    TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_EnterCell()
    FGrid.CellBackColor = CellBackColEnter
End Sub

Private Sub FGrid_GotFocus()
    txtgrid(0).Visible = False
    FGrid.CellBackColor = CellBackColEnter
    If FGrid.Row = FGrid.Rows - FGrid.FixedRows + 1 Then
        If FGrid.TextMatrix(FGrid.Row - 1, Col_DateTo) <> "" Then
            If FGrid.Row > 2 Then
                FGrid.TextMatrix(FGrid.Row, Col_DateFrom) = Format(CDate(FGrid.TextMatrix(FGrid.Row - 1, Col_DateTo)) + 1, "dd/mmm/yyyy")
            End If
        End If
    End If
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
        FGrid.CellBackColor = CellBackColLeave
        SendKeys "+{Tab}"
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown And Val(FGrid.Tag) = FGrid.Rows - 1 Then
'        FGrid.CellBackColor = CellBackColLeave
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
            FGrid.CellBackColor = CellBackColLeave
            TopCtrl1_eSave
        End If
        Exit Sub
'        SendKeysA vbKeyTab, True
'        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    Select Case FGrid.Col
        Case Col_DateFrom, Col_DateTo, Col_Target
            If KeyCode = vbKeyDelete And Shift = 0 Then
                FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
            End If
    End Select
    If KeyCode = vbKeyReturn Then
        Select Case FGrid.Col
            Case Col_DateFrom, Col_DateTo, Col_Target
                GridDblClick Me, FGrid, txtgrid, 0
        End Select
        TAddMode = False
    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyPress(keyascii As Integer)
On Error GoTo ELoop
    Select Case FGrid.Col
        Case Col_DateFrom, Col_DateTo
            Get_Text Me, FGrid, txtgrid, 0, False, keyascii
        Case Col_Target
            Get_Text Me, FGrid, txtgrid, 0, True, keyascii
    End Select
    If keyascii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
Dim I As Integer
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.ColSel = False Then Exit Sub
    If KeyCode = vbKeyD And Shift = 2 Then
        If FGrid.Row >= 1 Then
            If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
                If FGrid.Rows > 2 Then
                    FGrid.RemoveItem (FGrid.Row)
                Else
                    FGrid.Rows = 2
                    FGrid.AddItem FGrid.Rows - 1
                    FGrid.FixedRows = 2
                End If
            End If
            For I = 2 To FGrid.Rows - 1
               FGrid.TextMatrix(I, Col_SrNo) = I - 1
            Next
            Amt_Cal
        Else
            MsgBox "No Entries To Delete", vbCritical, "Delete Module"
        End If
        FGrid.SetFocus
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CellBackColLeave
End Sub

Private Sub FGrid_Scroll()
    txtgrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CellBackColLeave
End Sub

