VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmRepModelTgt 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Rep-wise / Model-wise Target (12 Month) Entry"
   ClientHeight    =   7230
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
   LinkTopic       =   " "
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11820
   Visible         =   0   'False
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   12
      Left            =   10905
      MaxLength       =   40
      TabIndex        =   47
      Top             =   6810
      Width           =   660
   End
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   11
      Left            =   8835
      MaxLength       =   40
      TabIndex        =   45
      Top             =   6810
      Width           =   660
   End
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   10
      Left            =   6930
      MaxLength       =   40
      TabIndex        =   43
      Top             =   6810
      Width           =   660
   End
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   9
      Left            =   5040
      MaxLength       =   40
      TabIndex        =   41
      Top             =   6810
      Width           =   660
   End
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   8
      Left            =   3000
      MaxLength       =   40
      TabIndex        =   39
      Top             =   6810
      Width           =   660
   End
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   7
      Left            =   945
      MaxLength       =   40
      TabIndex        =   37
      Top             =   6810
      Width           =   660
   End
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   6
      Left            =   10905
      MaxLength       =   40
      TabIndex        =   35
      Top             =   6540
      Width           =   660
   End
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   5
      Left            =   8835
      MaxLength       =   40
      TabIndex        =   33
      Top             =   6540
      Width           =   660
   End
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   4
      Left            =   6930
      MaxLength       =   40
      TabIndex        =   31
      Top             =   6540
      Width           =   660
   End
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   3
      Left            =   5040
      MaxLength       =   40
      TabIndex        =   29
      Top             =   6540
      Width           =   660
   End
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   2
      Left            =   3000
      MaxLength       =   40
      TabIndex        =   27
      Top             =   6540
      Width           =   660
   End
   Begin VB.TextBox TxtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   1
      Left            =   945
      MaxLength       =   40
      TabIndex        =   25
      Text            =   "99999"
      Top             =   6540
      Width           =   660
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
   Begin MSDataGridLib.DataGrid DGParty 
      Height          =   3330
      Left            =   5355
      Negotiate       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7335
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5874
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
      RowHeight       =   19
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Representative Name"
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
         DataField       =   "AcCode"
         Caption         =   "Ac.Code"
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
            ColumnWidth     =   5595.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   240
      Index           =   0
      Left            =   1905
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1980
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEE0FD&
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
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   1
      Top             =   495
      Width           =   3900
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   5190
      Left            =   30
      TabIndex        =   2
      Top             =   1080
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   9155
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   16
      FixedCols       =   2
      BackColorFixed  =   12243913
      ForeColorFixed  =   0
      BackColorSel    =   16777215
      ForeColorSel    =   8388608
      BackColorBkg    =   12243913
      GridColor       =   0
      GridColorFixed  =   8421504
      FocusRect       =   0
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   16
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Description"
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
      Left            =   315
      TabIndex        =   50
      Top             =   810
      Width           =   1485
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month Wise Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   18
      Left            =   90
      TabIndex        =   49
      Top             =   6270
      Width           =   1470
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   645
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   11745
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   12
      Left            =   10815
      TabIndex        =   48
      Top             =   6810
      Width           =   60
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   11
      Left            =   8745
      TabIndex        =   46
      Top             =   6810
      Width           =   60
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   10
      Left            =   6825
      TabIndex        =   44
      Top             =   6810
      Width           =   60
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   9
      Left            =   4950
      TabIndex        =   42
      Top             =   6810
      Width           =   60
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   8
      Left            =   2910
      TabIndex        =   40
      Top             =   6810
      Width           =   60
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   7
      Left            =   840
      TabIndex        =   38
      Top             =   6810
      Width           =   60
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   6
      Left            =   10815
      TabIndex        =   36
      Top             =   6540
      Width           =   60
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   5
      Left            =   8745
      TabIndex        =   34
      Top             =   6540
      Width           =   60
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   4
      Left            =   6825
      TabIndex        =   32
      Top             =   6540
      Width           =   60
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   2
      Left            =   4950
      TabIndex        =   30
      Top             =   6540
      Width           =   60
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   1
      Left            =   2910
      TabIndex        =   28
      Top             =   6540
      Width           =   60
   End
   Begin VB.Label LblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   0
      Left            =   840
      TabIndex        =   26
      Top             =   6540
      Width           =   60
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "April"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   17
      Left            =   8325
      TabIndex        =   24
      Top             =   0
      Width           =   390
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "April"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   16
      Left            =   6330
      TabIndex        =   23
      Top             =   30
      Width           =   390
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "April"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   15
      Left            =   5265
      TabIndex        =   22
      Top             =   15
      Width           =   390
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "April"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   14
      Left            =   3270
      TabIndex        =   21
      Top             =   45
      Width           =   390
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "April"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   13
      Left            =   1995
      TabIndex        =   20
      Top             =   75
      Width           =   390
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "April"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   12
      Left            =   0
      TabIndex        =   19
      Top             =   105
      Width           =   390
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "March"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   11
      Left            =   9840
      TabIndex        =   18
      Top             =   6810
      Width           =   540
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "February"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   10
      Left            =   7965
      TabIndex        =   17
      Top             =   6810
      Width           =   765
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "January"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   9
      Left            =   6075
      TabIndex        =   16
      Top             =   6810
      Width           =   675
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "December"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   8
      Left            =   4035
      TabIndex        =   15
      Top             =   6810
      Width           =   885
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "November"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   7
      Left            =   2010
      TabIndex        =   14
      Top             =   6810
      Width           =   855
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "October"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   6
      Left            =   90
      TabIndex        =   13
      Top             =   6810
      Width           =   690
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "September"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   5
      Left            =   9840
      TabIndex        =   12
      Top             =   6540
      Width           =   945
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "August"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   4
      Left            =   7965
      TabIndex        =   11
      Top             =   6540
      Width           =   615
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "July"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   3
      Left            =   6075
      TabIndex        =   10
      Top             =   6540
      Width           =   345
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "June"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   2
      Left            =   4035
      TabIndex        =   9
      Top             =   6540
      Width           =   405
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "May"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   1
      Left            =   2010
      TabIndex        =   8
      Top             =   6540
      Width           =   375
   End
   Begin VB.Label LblTot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "April"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   6540
      Width           =   390
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rep.Name"
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
      Height          =   255
      Index           =   9
      Left            =   315
      TabIndex        =   5
      Top             =   495
      Width           =   900
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   1515
      TabIndex        =   4
      Top             =   495
      Width           =   195
   End
End
Attribute VB_Name = "frmRepModelTgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TAddMode As Boolean
Dim GridKey As Integer
Dim RsParty As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim ExitCtrl As Boolean
Dim mSearchCode As String
Dim ForeColorSelEnter$
Dim BackColorSelLeave$

Private Const Party As Byte = 0                 ' Rep.Name
'* Grid Column Declaration
Private Const Col_SrNo As Byte = 0              ' Serial No
Private Const Col_ModelCode As Byte = 1         ' Model Code
Private Const Col_ModelDesc As Byte = 2         ' Model Desc
Private Const Col_Qty4 As Byte = 3              ' Qty Apr.
Private Const Col_Qty5 As Byte = 4              ' Qty May.
Private Const Col_Qty6 As Byte = 5              ' Qty Jun.
Private Const Col_Qty7 As Byte = 6              ' Qty Jul.
Private Const Col_Qty8 As Byte = 7              ' Qty Aug.
Private Const Col_Qty9 As Byte = 8              ' Qty Sep.
Private Const Col_Qty10 As Byte = 9             ' Qty Oct.
Private Const Col_Qty11 As Byte = 10            ' Qty Nov.
Private Const Col_Qty12 As Byte = 11            ' Qty Dec.
Private Const Col_Qty1 As Byte = 12             ' Qty Jan.
Private Const Col_Qty2 As Byte = 13             ' Qty Feb.
Private Const Col_Qty3 As Byte = 14             ' Qty Mar.
Private Const Col_Tot As Byte = 15              ' Total

Private Sub Disp_Text(Enb As Boolean)
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).Enabled = Enb
    Next
End Sub

Public Sub SEARCHBACK(ByVal MyValue As String)
'On Error GoTo ELoop
'    Master.MoveFirst
'    Master.FIND ("SearchCode='" & MYVALUE & "'")
'    MoveRec
'    BUTTONS True, Me, Master, 0
'Exit Sub
'eloop:
'    CheckError
On Error Resume Next
Dim I As Integer
    For I = 1 To FGrid.Rows - 1
'        If MyValue = Trim(txt(Party).Tag) + Trim(FGrid.TextMatrix(i, Col_ModelCode)) Then
        If MyValue = FGrid.TextMatrix(I, Col_ModelCode) Then
            FGrid.Row = I
            FGrid.Col = 3
            TxtGrid(0).left = FGrid.CellLeft
            FGrid.SetFocus
            Exit For
        End If
    Next
End Sub
'* Used for clear all text boxes used in the form
Private Sub BlankText()
Dim I As Byte
    For I = 0 To Txt.Count - 1
        Txt(I).TEXT = ""
        Txt(I).Tag = ""
    Next I
    For I = 1 To TxtTot.Count - 1
        TxtTot(I).TEXT = ""
    Next I
    FGrid.Rows = 1
    FGrid.AddItem FGrid.Rows
    FGrid.FixedRows = 1
End Sub

'* Used for intialize grid columns
Private Sub Grid_Ini()
Dim ColW As Integer

    DGParty.left = 5655: DGParty.top = 500
    ColW = 600
    With FGrid
        .top = 1080
        .left = Me.left ' + 45
        .width = Me.width - 90
        .height = 5190 'Me.height - 2200
        .RowHeightMin = PubGridRowHeight '220
        .Cols = 16
        
        .TextMatrix(0, Col_SrNo) = "S.No."
        .ColAlignment(Col_SrNo) = flexAlignLeftCenter
        .ColWidth(Col_SrNo) = 450

        .TextMatrix(0, Col_ModelCode) = "Model Code"
        .ColAlignmentFixed = flexAlignLeftCenter
        .ColAlignment(Col_ModelCode) = flexAlignLeftCenter
        .ColWidth(Col_ModelCode) = 3000
        
        .TextMatrix(0, Col_ModelDesc) = "Model Description"
        .ColWidth(Col_ModelDesc) = 0

        .TextMatrix(0, Col_Qty4) = "Apr"
        .ColAlignmentFixed(Col_Qty4) = flexAlignCenterBottom
        .ColWidth(Col_Qty4) = ColW

        .TextMatrix(0, Col_Qty5) = "May"
        .ColAlignmentFixed(Col_Qty5) = flexAlignCenterBottom
        .ColWidth(Col_Qty5) = ColW

        .TextMatrix(0, Col_Qty6) = "Jun"
        .ColAlignmentFixed(Col_Qty6) = flexAlignCenterBottom
        .ColWidth(Col_Qty6) = ColW

        .TextMatrix(0, Col_Qty7) = "Jul"
        .ColAlignmentFixed(Col_Qty7) = flexAlignCenterBottom
        .ColWidth(Col_Qty7) = ColW

        .TextMatrix(0, Col_Qty8) = "Aug"
        .ColAlignmentFixed(Col_Qty8) = flexAlignCenterBottom
        .ColWidth(Col_Qty8) = ColW

        .TextMatrix(0, Col_Qty9) = "Sep"
        .ColAlignmentFixed(Col_Qty9) = flexAlignCenterBottom
        .ColWidth(Col_Qty9) = ColW

        .TextMatrix(0, Col_Qty10) = "Oct"
        .ColAlignmentFixed(Col_Qty10) = flexAlignCenterBottom
        .ColWidth(Col_Qty10) = ColW

        .TextMatrix(0, Col_Qty11) = "Nov"
        .ColAlignmentFixed(Col_Qty11) = flexAlignCenterBottom
        .ColWidth(Col_Qty11) = ColW

        .TextMatrix(0, Col_Qty12) = "Dec"
        .ColAlignmentFixed(Col_Qty12) = flexAlignCenterBottom
        .ColWidth(Col_Qty12) = ColW

        .TextMatrix(0, Col_Qty1) = "Jan"
        .ColAlignmentFixed(Col_Qty1) = flexAlignCenterBottom
        .ColWidth(Col_Qty1) = ColW

        .TextMatrix(0, Col_Qty2) = "Feb"
        .ColAlignmentFixed(Col_Qty2) = flexAlignCenterBottom
        .ColWidth(Col_Qty2) = ColW

        .TextMatrix(0, Col_Qty3) = "Mar"
        .ColAlignmentFixed(Col_Qty3) = flexAlignCenterBottom
        .ColWidth(Col_Qty3) = ColW

        .TextMatrix(0, Col_Tot) = "Total"
        .ColAlignmentFixed(Col_Tot) = flexAlignCenterBottom
        .ColWidth(Col_Tot) = 750
    End With
    BackColorSelLeave = FGrid.BackColorSel
    ForeColorSelEnter = FGrid.ForeColorSel
End Sub

Private Sub Grid_Hide()
    If DGParty.Visible = True Then DGParty.Visible = False
End Sub

Private Function MakeBlank(Temp As Double) As String
    MakeBlank = IIf(Temp = 0, "", Temp)
End Function
Private Sub MoveRec()
Dim Rst As ADODB.Recordset, I As Integer, Tot As Double
On Error GoTo ELoop

If Master.RecordCount > 0 Then
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "SELECT V.REP_CODE As RepCode, Emp_Mast.Emp_Name, V.MODEL As ModelCode, Model.Model_Desc, V.* " & _
        "FROM (Veh_Target V Left JOIN Emp_Mast ON V.REP_CODE = Emp_Mast.Emp_Code) " & _
        "Left JOIN Model ON V.MODEL = Model.MODEL " & _
        "Where V.Rep_Code='" & Master!REPCODE & "'", GCn, adOpenDynamic, adLockOptimistic
    FGrid.Redraw = False
    FGrid.Rows = 1
    If Rst.RecordCount > 0 Then
        Txt(Party).TEXT = XNull(Rst!Emp_Name)
        Txt(Party).Tag = Rst!REPCODE
        mSearchCode = Txt(Party).Tag
        Dim mCellBackColor As String
        mCellBackColor = &HE3FAFD
        Do Until Rst.EOF
            Tot = 0
            Tot = Rst!TApr + Rst!TMay + Rst!TJun + Rst!TJul + Rst!TAug + Rst!TSep + Rst!TOct + Rst!TNov + Rst!TDec + Rst!TJan + Rst!TFeb + Rst!TMar
            I = I + 1
            With FGrid
                .AddItem ""
                .TextMatrix(I, Col_SrNo) = I
                .TextMatrix(I, Col_ModelCode) = Rst!ModelCode
                .TextMatrix(I, Col_ModelDesc) = IIf(IsNull(Rst!ModelCode), "", Rst!ModelCode)
                .TextMatrix(I, Col_Qty4) = MakeBlank(Rst!TApr)
                .TextMatrix(I, Col_Qty5) = MakeBlank(Rst!TMay)
                .TextMatrix(I, Col_Qty6) = MakeBlank(Rst!TJun)
                .TextMatrix(I, Col_Qty7) = MakeBlank(Rst!TJul)
                .TextMatrix(I, Col_Qty8) = MakeBlank(Rst!TAug)
                .TextMatrix(I, Col_Qty9) = MakeBlank(Rst!TSep)
                .TextMatrix(I, Col_Qty10) = MakeBlank(Rst!TOct)
                .TextMatrix(I, Col_Qty11) = MakeBlank(Rst!TNov)
                .TextMatrix(I, Col_Qty12) = MakeBlank(Rst!TDec)
                .TextMatrix(I, Col_Qty1) = MakeBlank(Rst!TJan)
                .TextMatrix(I, Col_Qty2) = MakeBlank(Rst!TFeb)
                .TextMatrix(I, Col_Qty3) = MakeBlank(Rst!TMar)
                .TextMatrix(I, Col_Tot) = MakeBlank(Tot)
            End With
            Rst.MoveNext
        Loop
        FGrid.Row = 1
        FGrid.FixedRows = 1
        CalTotal
    Else
        FGrid.AddItem FGrid.Rows
        FGrid.FixedRows = 1
    End If
    FGrid.Redraw = True
Else
    BlankText
End If
Grid_Hide
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FillModel()
Dim Rst As ADODB.Recordset, I As Integer
    Dim sitecond As String
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = " LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    If TopCtrl1.TopText2 = "Add" Then
        FGrid.Rows = 1
'        Set Rst = GCn.Execute("Select Model.Model As ModelCode,Model.Model_Desc " _
'          & "From Model where " & sitecond & " Order by Model.Model_Desc")
       Set Rst = GCn.Execute("Select Model.Model_Desc,Model.Model As ModelCode " _
          & "From Model Order by Model.Model_Desc")


    ElseIf TopCtrl1.TopText2 = "Edit" Then
'        Set Rst = GCn.Execute("Select Model.Model As ModelCode,Model.Model_Desc " _
'            & "From Model Where Model.Model Not in (Select Model From Veh_Target) and " & sitecond & " " _
'            & "Order by Model.Model_Desc")

     Set Rst = GCn.Execute("Select Model.Model As ModelCode,Model.Model_Desc " _
            & "From Model Where Model.Model Not in (Select Model From Veh_Target)  " _
            & "Order by Model.Model_Desc")
    
    End If
    If Rst.RecordCount > 0 Then
        I = 1
        Do Until Rst.EOF
                        '|0 Col_SrNo |1 Col_ModelCode         |2 Col_ModelDesc
            FGrid.AddItem FGrid.Rows & Chr(9) & Rst!ModelCode & Chr(9) & Rst!Model_Desc
            Rst.MoveNext '0                    1                    2
            I = I + 1
        Loop
        FGrid.FixedRows = 1
    End If
Set Rst = Nothing
End Sub

Private Function TxtGridLeave() As Boolean
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = MakeBlank(Val(TxtGrid(0).TEXT))
    ExitCtrl = True
    TxtGridLeave = True
    TxtGrid(0).Visible = False
    CalTotal
    FGrid.SetFocus
End Function

'* Used for Calculate the Total
Private Sub CalTotal()
Dim RowTot As Double, I As Integer
Dim TApr As Double, TMay As Double, TJun As Double, TJul As Double
Dim TAug As Double, TSep As Double, TOct As Double, TNov As Double
Dim TDec As Double, TJan As Double, TFeb As Double, TMar As Double
    With FGrid
        RowTot = Val(.TextMatrix(.Row, Col_Qty4)) + Val(.TextMatrix(.Row, Col_Qty5)) + _
              Val(.TextMatrix(.Row, Col_Qty6)) + Val(.TextMatrix(.Row, Col_Qty7)) + _
              Val(.TextMatrix(.Row, Col_Qty8)) + Val(.TextMatrix(.Row, Col_Qty9)) + _
              Val(.TextMatrix(.Row, Col_Qty10)) + Val(.TextMatrix(.Row, Col_Qty11)) + _
              Val(.TextMatrix(.Row, Col_Qty12)) + Val(.TextMatrix(.Row, Col_Qty1)) + _
              Val(.TextMatrix(.Row, Col_Qty2)) + Val(.TextMatrix(.Row, Col_Qty3))
        .TextMatrix(.Row, Col_Tot) = MakeBlank(RowTot)
        For I = 1 To .Rows - 1
            TApr = TApr + Val(.TextMatrix(I, Col_Qty4))
            TMay = TMay + Val(.TextMatrix(I, Col_Qty5))
            TJun = TJun + Val(.TextMatrix(I, Col_Qty6))
            TJul = TJul + Val(.TextMatrix(I, Col_Qty7))
            TAug = TAug + Val(.TextMatrix(I, Col_Qty8))
            TSep = TSep + Val(.TextMatrix(I, Col_Qty9))
            TOct = TOct + Val(.TextMatrix(I, Col_Qty10))
            TNov = TNov + Val(.TextMatrix(I, Col_Qty11))
            TDec = TDec + Val(.TextMatrix(I, Col_Qty12))
            TJan = TJan + Val(.TextMatrix(I, Col_Qty1))
            TFeb = TFeb + Val(.TextMatrix(I, Col_Qty2))
            TMar = TMar + Val(.TextMatrix(I, Col_Qty3))
        Next
    End With
    TxtTot(1).TEXT = TApr
    TxtTot(2).TEXT = TMay
    TxtTot(3).TEXT = TJun
    TxtTot(4).TEXT = TJul
    TxtTot(5).TEXT = TAug
    TxtTot(6).TEXT = TSep
    TxtTot(7).TEXT = TOct
    TxtTot(8).TEXT = TNov
    TxtTot(9).TEXT = TDec
    TxtTot(10).TEXT = TJan
    TxtTot(11).TEXT = TFeb
    TxtTot(12).TEXT = TMar
End Sub

Private Sub DGParty_Click()
    DGParty.Visible = False
    If RsParty.RecordCount > 0 Then
        Txt(Party).TEXT = RsParty!Name
        Txt(Party).Tag = RsParty!Code
    End If
    Txt(Party).SetFocus
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
'to modify
    WinSetting Me:    Grid_Ini
    TopCtrl1.Tag = PubUParam
    For I = 0 To Txt.Count - 1
        Txt(I).BackColor = CtrlBColOrg
        Txt(I).ForeColor = CtrlFColOrg
    Next
 Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = " and left(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "Select Emp_Code as Code,Emp_Name As Name From Emp_Mast Where Emp_Type=0 " & sitecond & " Order by Emp_Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty

    Set Master = New ADODB.Recordset
    Master.LockType = adLockOptimistic
    Master.CursorLocation = adUseClient
    Master.CursorType = adOpenDynamic
      If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where left(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    Set Master = GCn.Execute("Select Distinct V.Rep_Code As SearchCode,Rep_Code As RepCode " _
            & "From Veh_Target V " & sitecond & " Order by V.Rep_Code")

    MoveRec
    Disp_Text SETS("INI", Me, Master)
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
    FillModel
    Txt(Party).SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
On Error GoTo ELoop
    Disp_Text SETS("EDIT", Me, Master)
    Txt(Party).Enabled = False
    FillModel
    FGrid.SetFocus
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
On Error GoTo ELoop
    If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
        GCn.BeginTrans
        GCn.Execute ("Delete From Veh_Target Where Rep_Code='" & Txt(Party).Tag & "'")
        GCn.CommitTrans
        Master.Requery
        MoveRec
        BUTTONS True, Me, Master, 0
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
'    GSQL = "Select (V.Rep_Code+V.Model) As SearchCode,Emp_Mast.Emp_Name as RepName,Model.Model_Desc As ModelName FROM (Veh_Target V Left Join Emp_Mast On V.Rep_Code = Emp_Mast.Emp_Code) Left Join Model on V.Model=Model.Model Order by Emp_Mast.Emp_Name,Model.Model"
        Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If


    GSQL = "Select Model.Model As SearchCode,Model.Model As ModelName FROM Model " & sitecond & " Order by Model.Model"

    Set SearchForm = Me
    FIND.Show vbModal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TopCtrl1_ePrn()
Dim Rst As ADODB.Recordset
Dim mQry As String
    Dim sitecond As String
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "and LEFT(v.site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If




mQry = "SELECT V.REP_CODE As RepCode, Emp_Mast.Emp_Name, V.MODEL As ModelCode, Model.Model_Desc, " & _
    "V.REP_CODE,V.MODEL,V.Site_Code,V.tAPR,V.tMAY,V.tJUN,V.tJUL,V.tAUG,V.tSEP,V.tOCT,V.tNOV,V.tDEC,V.tJAN,V.tFEB,V.tMAR,V.U_Name,V.U_EntDt,V.U_AE " & _
    "FROM (Veh_Target V Left JOIN Emp_Mast ON V.REP_CODE = Emp_Mast.Emp_Code) " & _
    "Left JOIN Model ON V.MODEL = Model.MODEL Where V.Rep_Code='" & Master!REPCODE & "' " & sitecond & " and " & _
    "(V.TApr + V.TMay + V.TJun + V.TJul + V.TAug + V.TSep + V.TOct + V.TNov + V.TDec + V.TJan + V.TFeb + V.TMar) <> 0"
       
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open (mQry), GCn, adOpenStatic, adLockReadOnly
        If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub

        CreateFieldDefFile Rst, PubRepoPath + "\VehRepTgt.TTX", True
        Set rpt = rdApp.OpenReport(PubRepoPath & "\VehRepTgt.RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
        
        Call Report_View(rpt, "Modelwise Target")
End Sub

Private Sub TopCtrl1_eRef()
    RsParty.Requery
    Master.Requery
End Sub

Private Sub TopCtrl1_eSave()
Dim I As Integer, mTrans As Boolean
Dim Rst As ADODB.Recordset
On Error GoTo ELoop
    If TxtGrid(0).Visible = True Then
        If TxtGridLeave = False Then
            TxtGrid_LostFocus 0
            TxtGrid(0).SetFocus
            Exit Sub
        End If
    End If
    Grid_Hide

    If IsValid(Txt(Party), "Rep. Name") = False Then Exit Sub
    If TopCtrl1.TopText2 = "Add" Then
        If GCn.Execute("Select Count(*) From Veh_Target Where Rep_Code='" & Txt(Party).Tag & "'").Fields(0) > 0 Then
            MsgBox "Target for " & Txt(Party).TEXT & " Already Exists", vbCritical, "Validation Error"
            Txt(Party).SetFocus
        End If
    End If

    GCn.BeginTrans
        mTrans = True
        GCn.Execute ("Delete From Veh_Target Where Rep_Code='" & Txt(Party).Tag & "'")
        For I = 1 To FGrid.Rows - 1
            If FGrid.TextMatrix(I, Col_ModelDesc) <> "" Then
                GCn.Execute "Insert Into Veh_Target(" _
                & "Rep_Code,Model,Site_Code," _
                & "tApr,tMay,tJun,tJul," _
                & "tAug,tSep,tOct,tNov," _
                & "tDec,tJan,tFeb,tMar," _
                & "U_Name,U_EntDt,U_AE) " _
                & "Values(" _
                & "'" & Txt(Party).Tag & "','" & FGrid.TextMatrix(I, Col_ModelCode) & "','" & PubSiteCode & "'," _
                & "" & Val(FGrid.TextMatrix(I, Col_Qty4)) & "," & Val(FGrid.TextMatrix(I, Col_Qty5)) & "," & Val(FGrid.TextMatrix(I, Col_Qty6)) & "," & Val(FGrid.TextMatrix(I, Col_Qty7)) & "," _
                & "" & Val(FGrid.TextMatrix(I, Col_Qty8)) & "," & Val(FGrid.TextMatrix(I, Col_Qty9)) & "," & Val(FGrid.TextMatrix(I, Col_Qty10)) & "," & Val(FGrid.TextMatrix(I, Col_Qty11)) & "," _
                & "" & Val(FGrid.TextMatrix(I, Col_Qty12)) & "," & Val(FGrid.TextMatrix(I, Col_Qty1)) & "," & Val(FGrid.TextMatrix(I, Col_Qty2)) & "," & Val(FGrid.TextMatrix(I, Col_Qty3)) & "," _
                & "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"
            End If
        Next
    GCn.CommitTrans
    mTrans = False
    mSearchCode = Txt(Party).Tag
    Master.Requery
    Master.FIND "SearchCode = '" & mSearchCode & "'"
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        TopCtrl1_eAdd
        Exit Sub
    End If
    Disp_Text SETS("INI", Me, Master)
    MoveRec
Exit Sub
ELoop:
    If mTrans = True Then
        GCn.RollbackTrans: CheckError
    Else
        CheckError
    End If
End Sub

Private Sub TopCtrl1_eCancel()
On Error GoTo ELoop
Dim I As Byte
    If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
        Disp_Text SETS("INI", Me, Master)
        MoveRec
        For I = 0 To Txt.Count - 1
            Txt(I).BackColor = CtrlBColOrg
            Txt(I).ForeColor = CtrlFColOrg
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
    Ctrl_GetFocus Txt(Index)
    TxtGrid(0).Visible = False
    Grid_Hide
    Select Case Index
    Case Party
        If RsParty.RecordCount = 0 Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "Name ='" & Txt(Index).TEXT & "'"
        End If
    End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    Select Case Index
        Case Party
            If RsParty.RecordCount > 0 Then
                DGridTxtKeyDown DGParty, Txt, Party, RsParty, KeyCode, False, 1
            Else
                Txt_Validate Index, True
            End If
    End Select
    If DGParty.Visible = False Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then Ctrl_DownKeyDown KeyCode, Shift
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
Select Case Index
    Case Party
        If DGParty.Visible = True Then DGridTxtKeyPress Txt, Party, RsParty, KeyAscii, "Name"
End Select
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub Txt_LostFocus(Index As Integer)
    Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset, I As Byte
On Error GoTo ELoop
    Select Case Index
        Case Party
            If Txt(Party).TEXT <> "" Then
                If GCn.Execute("Select Count(*) From Veh_Target Where Rep_Code='" & Txt(Party).Tag & "'").Fields(0) > 0 Then
                    MsgBox "Target for " & Txt(Party).TEXT & " Already Exists", vbCritical, "Validation Error"
                    Txt(Party).SetFocus
                    Cancel = True
                    Exit Sub
                End If
            End If
            If RsParty.RecordCount > 0 And Txt(Index).TEXT <> "" Then
                Txt(Party).Tag = RsParty!Code
                Txt(Party).TEXT = RsParty!Name
            End If
    End Select
Set Rst = Nothing
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_GotFocus(Index As Integer)
On Error GoTo ELoop
    If ExitCtrl = False Then Exit Sub
    Ctrl_GetFocus TxtGrid(Index)
    FGrid.CellBackColor = CellBackColLeave
    TxtGrid(0).Tag = FGrid.TextMatrix(FGrid.Row, FGrid.Col)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If KeyCode = vbKeyEscape Then
        TxtGrid(0).TEXT = TxtGrid(0).Tag
        TxtGrid(0).Visible = False
        FGrid.SetFocus
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
        If TxtGridLeave = True Then
            If FGrid.Col = Col_Qty3 And KeyCode = vbKeyReturn Then
                If FGrid.Row <> FGrid.Rows - 1 Then
                    FGrid.Row = FGrid.Row + 1
                    FGrid.Col = Col_Qty4
                    FGrid.SetFocus
                End If
            Else
                If FGrid.Col <> Col_Tot Then
                    GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, 14
                End If
            End If
        Else
            TxtGrid_LostFocus 0
            TxtGrid(0).SetFocus
        End If
    End If
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ELoop
    CheckQuote KeyAscii
    NumPress TxtGrid(Index), KeyAscii, 4, 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_LostFocus(Index As Integer)
On Error GoTo ELoop
    If ExitCtrl = False Then Exit Sub
    Ctrl_validate TxtGrid(Index)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
    FGrid.TextMatrix(FGrid.Row, FGrid.Col) = MakeBlank(Val(TxtGrid(Index).TEXT))
    CalTotal
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_LostFocus()
    FGrid.BackColorSel = BackColorSelLeave
    FGrid.ForeColorSel = FGrid.ForeColor
End Sub

Private Sub FGrid_RowColChange()
Lbl(0).CAPTION = "Model Description : " & FGrid.TextMatrix(FGrid.Row, Col_ModelDesc)
End Sub

Private Sub FGrid_Click()
    TxtGrid(0).Visible = False
End Sub

Private Sub FGrid_DblClick()
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If FGrid.Col = Col_ModelDesc Or FGrid.Col = Col_Tot Then
    Else
        GridDblClick Me, FGrid, TxtGrid, 0
    End If
    TAddMode = False
Exit Sub
ELoop:
    CheckError
End Sub
Private Sub FGrid_GotFocus()
    FGrid.BackColorSel = BackColorSelEnter
    FGrid.ForeColorSel = ForeColorSelEnter
    Grid_Hide
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
    If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
    If KeyCode = vbKeyUp And Val(FGrid.Tag) = (FGrid.Rows - (FGrid.Rows - 1)) Then
        SendKeys "+{Tab}"
        KeyCode = 0
    End If
    GridKey = KeyCode
    FGrid.Tag = FGrid.Row
    If KeyCode = vbKeyDelete And Shift = 0 Then
        FGrid.TextMatrix(FGrid.Row, FGrid.Col) = ""
        CalTotal
    End If
    If KeyCode = vbKeyReturn Then
        If FGrid.Col = Col_ModelDesc Or FGrid.Col = Col_Tot Then
        Else
            GridDblClick Me, FGrid, TxtGrid, 0
        End If
        TAddMode = False
    End If
    If KeyCode = vbKeyDown And FGrid.Row = FGrid.Rows - 1 Then
        If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then TopCtrl1_eSave
    End If
    KeyCode = 0
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
On Error GoTo ELoop
    If FGrid.Col = Col_ModelDesc Or FGrid.Col = Col_Tot Then
    Else
        Get_Text Me, FGrid, TxtGrid, 0, True, KeyAscii
    End If
    If KeyAscii <> vbKeyReturn Then TAddMode = True
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub FGrid_Scroll()
    TxtGrid(0).Visible = False
    Grid_Hide
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
'    FGrid.CellBackColor = CellBackColLeave
End Sub
