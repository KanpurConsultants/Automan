VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmFindPart 
   BackColor       =   &H00CFE0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find Part"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGPart 
      Height          =   4335
      Left            =   465
      Negotiate       =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1725
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Part No."
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
         Caption         =   "Part Name"
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
         DataField       =   "MRP"
         Caption         =   "MRP"
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
         DataField       =   "CurrStk"
         Caption         =   "CurrStk"
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
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2654.929
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Search On Part Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   3690
      TabIndex        =   1
      Top             =   255
      Width           =   2355
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00CFE0E0&
      Caption         =   "Search On Part No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   0
      Left            =   1095
      TabIndex        =   0
      Top             =   255
      Value           =   -1  'True
      Width           =   2355
   End
   Begin VB.Frame FrmDetail 
      BackColor       =   &H00CAF1FD&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   2385
      Left            =   825
      TabIndex        =   6
      Top             =   1950
      Width           =   6285
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   36
         Top             =   1635
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MRP TaxPaid"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   35
         Top             =   1635
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taxable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   8
         Left            =   315
         TabIndex        =   34
         Top             =   1860
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item Detail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   270
         Index           =   9
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   6285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Stock"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   195
         Index           =   10
         Left            =   315
         TabIndex        =   32
         Top             =   1155
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MRP Taxable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   315
         TabIndex        =   31
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Paid"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   13
         Left            =   315
         TabIndex        =   30
         Top             =   2085
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   14
         Left            =   3915
         TabIndex        =   29
         Top             =   945
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   4830
         TabIndex        =   28
         Top             =   720
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   75
         TabIndex        =   27
         Top             =   510
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Rates >>"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   18
         Left            =   60
         TabIndex        =   26
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   3285
         TabIndex        =   25
         Top             =   720
         Width           =   375
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   60
         Left            =   3750
         TabIndex        =   24
         Top             =   735
         Width           =   1020
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Currentt Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   51
         Left            =   1215
         TabIndex        =   23
         Top             =   510
         Width           =   1665
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   53
         Left            =   2400
         TabIndex        =   22
         Top             =   1380
         Width           =   375
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   195
         Index           =   52
         Left            =   2400
         TabIndex        =   21
         Top             =   1155
         Width           =   375
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   61
         Left            =   5235
         TabIndex        =   20
         Top             =   735
         Width           =   1020
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   56
         Left            =   3900
         TabIndex        =   19
         Top             =   1860
         Width           =   375
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   55
         Left            =   2400
         TabIndex        =   18
         Top             =   2085
         Width           =   375
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   54
         Left            =   2400
         TabIndex        =   17
         Top             =   1860
         Width           =   375
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999999.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   59
         Left            =   2190
         TabIndex        =   16
         Top             =   735
         Width           =   1020
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   57
         Left            =   3900
         TabIndex        =   15
         Top             =   2085
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   1755
         TabIndex        =   14
         Top             =   720
         Width           =   345
      End
      Begin VB.Label LblFrm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   3900
         TabIndex        =   13
         Top             =   1380
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   21
         Left            =   2145
         TabIndex        =   12
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   20
         Left            =   75
         TabIndex        =   11
         Top             =   315
         Width           =   690
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Currentt Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   62
         Left            =   1215
         TabIndex        =   10
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label LblFrm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   58
         Left            =   4920
         TabIndex        =   9
         Top             =   300
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bin Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   3765
         TabIndex        =   8
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   195
         Index           =   15
         Left            =   3885
         TabIndex        =   7
         Top             =   1155
         Width           =   315
      End
      Begin VB.Line Line1 
         X1              =   2130
         X2              =   45
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line2 
         X1              =   3855
         X2              =   2925
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Line Line3 
         X1              =   4350
         X2              =   6240
         Y1              =   1035
         Y2              =   1035
      End
   End
   Begin VB.CommandButton CmdSearch 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1215
      Width           =   1155
   End
   Begin VB.CommandButton CmdSearch 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2835
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1215
      Width           =   1155
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1860
      MaxLength       =   40
      TabIndex        =   2
      Top             =   885
      Width           =   4905
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   885
      TabIndex        =   5
      Top             =   930
      Width           =   630
   End
End
Attribute VB_Name = "FrmFindPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsPart1 As ADODB.Recordset
Private Const F_LName As Byte = 51
Private Const F_CurStkQty As Byte = 52
Private Const F_MRPQty As Byte = 53
Private Const F_TBQty As Byte = 54
Private Const F_TPQty As Byte = 55
Private Const F_Mrp As Byte = 0
Private Const F_TBRate As Byte = 56
Private Const F_TPRate As Byte = 57
Private Const F_Bin As Byte = 58
Private Const F_LastRate As Byte = 59
Private Const F_HPRate As Byte = 60
Private Const F_LPRate As Byte = 61
Private Const F_PNo As Byte = 62
Dim RstHelp As ADODB.Recordset

Private Sub FillData()
Dim mStkMrpTb   As Double
Dim mStkMrpTbOp   As Double
Dim mStkMrpTp   As Double
Dim mStkMrpTpOp   As Double
Dim mStkTb      As Double
Dim mStkTbOp      As Double
Dim mStkTp      As Double
Dim mStkTpOp      As Double
If RsPart1.RecordCount = 0 Or (RsPart1.EOF = True Or RsPart1.BOF = True) Then
    LblFrm(F_PNo).CAPTION = "No Current Search"
    LblFrm(F_LName).CAPTION = "No Current Search"
    LblFrm(F_CurStkQty).CAPTION = "0.00"
    LblFrm(F_MRPQty).CAPTION = "0.00"
    LblFrm(F_TBQty).CAPTION = "0.00"
    LblFrm(F_TPQty).CAPTION = "0.00"
    LblFrm(F_Mrp).CAPTION = "0.00"
    LblFrm(F_TBRate).CAPTION = "0.00"
    LblFrm(F_TPRate).CAPTION = "0.00"
    LblFrm(F_Bin).CAPTION = ""
    LblFrm(F_LastRate).CAPTION = "N/A"
    LblFrm(F_HPRate).CAPTION = "0.00"
    LblFrm(F_LPRate).CAPTION = "0.00"
    LblFrm(1).CAPTION = "0.00"
   
Else
    LblFrm(F_PNo).CAPTION = RsPart1!Code
    LblFrm(F_LName).CAPTION = RsPart1!Name
    LblFrm(F_CurStkQty).CAPTION = Format(IIf(IsNull(RsPart1!Curstk), 0, RsPart1!Curstk), "0.000")
    
    mStkMrpTb = VNull(GCn.Execute("Select sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=1 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & "" & _
        " and Part_No='" & RsPart1!Code & "'").Fields(0).Value)
        
    mStkMrpTbOp = VNull(GCn.Execute("Select sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=1 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO'" & _
        " and Part_No='" & RsPart1!Code & "'").Fields(0).Value)
    
    mStkMrpTp = VNull(GCn.Execute("Select sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=0 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & "" & _
        " and Part_No='" & RsPart1!Code & "'").Fields(0).Value)
        
    mStkMrpTpOp = VNull(GCn.Execute("Select sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=0 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO'" & _
        " and Part_No='" & RsPart1!Code & "'").Fields(0).Value)
    
    mStkTb = VNull(GCn.Execute("Select sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=1 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & "" & _
        " and Part_No='" & RsPart1!Code & "'").Fields(0).Value)
        
    mStkTbOp = VNull(GCn.Execute("Select sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=1 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO'" & _
        " and Part_No='" & RsPart1!Code & "'").Fields(0).Value)
    
    mStkTp = VNull(GCn.Execute("Select sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=0 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & "" & _
        " and Part_No='" & RsPart1!Code & "'").Fields(0).Value)
        
    mStkTpOp = VNull(GCn.Execute("Select sum((Qty_Rec)-(Qty_Iss-Qty_Ret)) as BalQty " & _
        " from Sp_Stock where Tax_YN=0 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO'" & _
        " and Part_No='" & RsPart1!Code & "'").Fields(0).Value)
    
''    mstkmprtb = GCn.Execute("Select IIF(IsNUll(sum((IIf(IsNUll(Qty_Rec),0,Qty_Rec))-(IIF(IsNull(Qty_Iss),0,Qty_Iss)-IIF(IsNUll(Qty_Ret),0,Qty_Ret)))),0,sum((IIf(IsNUll(Qty_Rec),0,Qty_Rec))-(IIF(IsNull(Qty_Iss),0,Qty_Iss)-IIF(IsNUll(Qty_Ret),0,Qty_Ret)))) + (Select IIF(IsNull((sum((IIf(IsNUll(Qty_Rec),0,Qty_Rec))-(IIF(IsNull(Qty_Iss),0,Qty_Iss)-IIF(IsNUll(Qty_Ret),0,Qty_Ret)))),0,(sum((IIf(IsNUll(Qty_Rec),0,Qty_Rec))-(IIF(IsNull(Qty_Iss),0,Qty_Iss)-IIF(IsNUll(Qty_Ret),0,Qty_Ret)))) from Sp_Stock where Tax_YN=1 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO' AND Part_No='" & RsPart1!Code & "') From Sp_Stock where Tax_YN=1 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & " AND Part_No='" & RsPart1!Code & "'").Fields(0).Value
''    mstkmprtp = GCn.Execute("Select sum((IIf(IsNUll(Qty_Rec),0,Qty_Rec))-(IIF(IsNull(Qty_Iss),0,Qty_Iss)-IIF(IsNUll(Qty_Ret),0,Qty_Ret))) + (Select sum((IIf(IsNUll(Qty_Rec),0,Qty_Rec))-(IIF(IsNull(Qty_Iss),0,Qty_Iss)-IIF(IsNUll(Qty_Ret),0,Qty_Ret))) from Sp_Stock where Tax_YN=0 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO' AND Part_No='" & RsPart1!Code & "') From Sp_Stock where Tax_YN=0 and MRP_YN=1 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & " AND Part_No='" & RsPart1!Code & "'").Fields(0).Value
''    mStkTb = GCn.Execute("Select sum((IIf(IsNUll(Qty_Rec),0,Qty_Rec))-(IIF(IsNull(Qty_Iss),0,Qty_Iss)-IIF(IsNUll(Qty_Ret),0,Qty_Ret))) + (Select sum((IIf(IsNUll(Qty_Rec),0,Qty_Rec))-(IIF(IsNull(Qty_Iss),0,Qty_Iss)-IIF(IsNUll(Qty_Ret),0,Qty_Ret))) from Sp_Stock where Tax_YN=1 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO' AND Part_No='" & RsPart1!Code & "') From Sp_Stock where Tax_YN=1 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & " AND Part_No='" & RsPart1!Code & "'").Fields(0).Value
''    mStkTp = VNull(GCn.Execute("Select sum((IIf(IsNUll(Qty_Rec),0,Qty_Rec))-(IIF(IsNull(Qty_Iss),0,Qty_Iss)-IIF(IsNUll(Qty_Ret),0,Qty_Ret))) + (Select sum((IIf(IsNUll(Qty_Rec),0,Qty_Rec))-(IIF(IsNull(Qty_Iss),0,Qty_Iss)-IIF(IsNUll(Qty_Ret),0,Qty_Ret))) from Sp_Stock where Tax_YN=0 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date = " & ConvertDate(DateAdd("d", -1, PubStartDate)) & " and V_Type='SXAO' AND Part_No='" & RsPart1!Code & "') From Sp_Stock where Tax_YN=0 and MRP_YN=0 and Left(DocId,1)='" & PubDivCode & "' and V_Date >= " & ConvertDate(PubStartDate) & " AND Part_No='" & RsPart1!Code & "'").Fields(0).Value)
    
    If UCase(left(PubComp_Name, 3)) = "JMK" Then
        LblFrm(F_MRPQty) = Format(VNull(RsPart1!Cur_MRP_TbStk) + VNull(RsPart1!Cur_TB_STk), "0.000")
        LblFrm(1).CAPTION = Format(VNull(RsPart1!Cur_MRP_TPStk) + VNull(RsPart1!Cur_TP_Stk), "0.000")
        LblFrm(F_TBQty).Visible = False
        LblFrm(F_TPQty).Visible = False
    Else
        LblFrm(F_MRPQty).CAPTION = Format(mStkMrpTb + mStkMrpTbOp, "0.00") 'Format(IIf(IsNull(RsPart1!Cur_MRP_TbStk), 0, RsPart1!Cur_MRP_TbStk), "0.000")
        LblFrm(1).CAPTION = Format(mStkMrpTp + mStkMrpTpOp, "0.00") 'Format(IIf(IsNull(RsPart1!Cur_MRP_TPStk), 0, RsPart1!Cur_MRP_TPStk), "0.000")
        LblFrm(F_TBQty).CAPTION = Format(mStkTb + mStkTbOp, "0.00") 'Format(IIf(IsNull(RsPart1!Cur_TB_STk), 0, RsPart1!Cur_TB_STk), "0.000")
        LblFrm(F_TPQty).CAPTION = Format(mStkTp + mStkTpOp, "0.00") 'Format(IIf(IsNull(RsPart1!Cur_TP_Stk), 0, RsPart1!Cur_TP_Stk), "0.000")
    End If
            
    LblFrm(F_Mrp).CAPTION = Format(IIf(IsNull(RsPart1!MRP), 0, RsPart1!MRP), "0.00")
    LblFrm(F_TBRate).CAPTION = Format(IIf(IsNull(RsPart1!TB_SRate), 0, RsPart1!TB_SRate), "0.00")
    LblFrm(F_TPRate).CAPTION = Format(IIf(IsNull(RsPart1!TP_SRate), 0, RsPart1!TP_SRate), "0.00")
    LblFrm(F_Bin).CAPTION = IIf(IsNull(RsPart1!Bin_Loca), "", RsPart1!Bin_Loca)
    LblFrm(F_LastRate).CAPTION = Format(VNull(GCn.Execute("Select PurRate From Part where Part_No='" & RsPart1!Code & "' and Div_Code='" & PubDivCode & "'").Fields(0).Value), "0.00")
    LblFrm(F_HPRate).CAPTION = Format(VNull(GCn.Execute("Select Max(Rate) From SP_Stock where Part_No='" & RsPart1!Code & "' and v_Type='SXGR'").Fields(0).Value), "0.00")
    LblFrm(F_LPRate).CAPTION = Format(VNull(GCn.Execute("Select Min(Rate) From SP_Stock where Part_No='" & RsPart1!Code & "' and v_Type='SXGR'").Fields(0).Value), "0.00")
End If
End Sub

Private Sub cmdsearch_Click(Index As Integer)
Select Case Index
Case 0
    StkUpd txt(0).TEXT
    If Option1(0).Value = True Then
        If IsValid(txt(0), "Part No") = False Then Exit Sub
        Set RsPart1 = GCn.Execute("SELECT P.PART_NO AS code,P.Part_Name AS name, P.Local_Name AS LName, P.UNIT, P.Part_Grade, P.MRP,P.Cur_MRP_TBStk+P.Cur_MRP_TPStk as MRPQty,Cur_MRP_TBStk,Cur_MRP_TPStk, P.Cur_MRP_TBStk+P.Cur_MRP_TPStk+P.Cur_TB_Stk+P.Cur_TP_Stk AS CurStk, P.Cur_TB_Stk, P.Cur_TP_Stk, P.Cur_TB_Stk, P.MRP, P.TB_SRate, P.TP_SRate, P.Bin_Loca, P.High_Pur_Rate, P.Low_Pur_Rate, PD.PurcDisc_Per FROM Part P left JOIN Part_DiscFactor PD ON P.Disc_Factor = PD.DiscFac_Catg where P.PART_NO = '" & txt(0).TEXT & "' AND P.div_code ='" & PubDivCode & "'")
    Else
        If IsValid(txt(0), "Part Name") = False Then Exit Sub
        Set RsPart1 = GCn.Execute("SELECT P.PART_NO AS code, P.Part_Name AS name, P.Local_Name AS LName, P.UNIT, P.Part_Grade, P.MRP,P.Cur_MRP_TBStk+P.Cur_MRP_TPStk as MRPQty,Cur_MRP_TBStk,Cur_MRP_TPStk, P.Cur_MRP_TBStk+P.Cur_MRP_TPStk+P.Cur_TB_Stk+P.Cur_TP_Stk AS CurStk, P.Cur_TB_Stk, P.Cur_TP_Stk, P.Cur_TB_Stk, P.MRP, P.TB_SRate, P.TP_SRate, P.Bin_Loca, P.High_Pur_Rate, P.Low_Pur_Rate, PD.PurcDisc_Per FROM Part P left JOIN Part_DiscFactor PD ON P.Disc_Factor = PD.DiscFac_Catg where P.PART_Name = '" & txt(0).TEXT & "' AND P.div_code ='" & PubDivCode & "'")
    End If
    If RsPart1.RecordCount = 0 Then
        MsgBox "No Record Found Of Given Part", vbInformation, "Search Result"
    End If
    FillData
Case 1
    Unload Me
End Select
End Sub


Private Sub Form_Load()
Set RstHelp = New ADODB.Recordset
RstHelp.CursorLocation = adUseClient
Set RstHelp = RsPart  'GCn.Execute("select Part_No as Code,Part_Name as Name from Part where Div_code='" & PubDivCode & "'")
Set DGPart.DataSource = RstHelp
Option1_Click (0)
End Sub
Private Sub Option1_Click(Index As Integer)
Select Case Index
    Case 0
        LBL(0).CAPTION = "Part No"
        RstHelp.Sort = "Code"
    Case 1
        LBL(0).CAPTION = "Part Name"
        RstHelp.Sort = "Name"
End Select
txt(0).TEXT = ""
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
If Option1(0).Value = True Then
    If DGPart.Visible = False Then DGridColSwap DGPart, 0
    'DGridTxtKeyDown_Mast DGPart, Txt, Index, RstHelp, KeyCode, True, 0
    DGridTxtKeyDown DGPart, txt, Index, RstHelp, KeyCode, True, 0
Else
    If DGPart.Visible = False Then DGridColSwap DGPart, 1
    'DGridTxtKeyDown_Mast DGPart, Txt, Index, RstHelp, KeyCode, True, 1
    DGridTxtKeyDown DGPart, txt, Index, RstHelp, KeyCode, True, 1
End If
End Sub

Private Sub txt_KeyPress(Index As Integer, keyascii As Integer)
If Option1(0).Value = True Then
    'If Len(txt(0).TEXT) > 22 Then KeyAscii = 0
    'DGridTxtKeyUp_Mast Txt, Index, RstHelp, KeyAscii, "Code"
    DGridTxtKeyPress txt, Index, RstHelp, keyascii, "Code"
    'DGridTxtKeyUp_Mast Txt, Index, RstHelp, KeyAscii, "Code"
Else
    'If Len(txt(0).TEXT) > 40 Then KeyAscii = 0
    DGridTxtKeyPress txt, Index, RstHelp, keyascii, "Name"
    'If DGPart.Visible = True Then DGridTxtKeyUp_Mast Txt, Index, RstHelp, KeyAscii, "Name"
End If

If keyascii = 13 Then cmdsearch_Click (0)
End Sub
Private Sub DGPart_Click()
    DGPart.Visible = False
End Sub
Private Sub Grid_Hide()
    If DGPart.Visible = True Then DGPart.Visible = False
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    txt(0).Visible = False
    Grid_Hide
End If
End Sub
Private Sub StkUpd(PNo As String)
Dim I As Integer
Dim mSQry$, mQRY$
    Dim Rst As ADODB.Recordset
    GCn.BeginTrans
        mSQry = "Select Sum(Qty_Rec-Qty_Iss+Qty_Ret) From Sp_Stock " & _
                "WHERE (V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " " & _
                "Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & ") " & _
                "And Part_No='" & PNo & "' "
    
        
        If PubBackEnd = "S" Then
            mQRY = "Select P.Part_No as Code, P.Part_No, P.Part_Name As Name, P.Local_Name as LName, P.Unit , P.MRP, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, " & _
                            "(Select PurcDisc_Per From Part_DiscFactor Where DiscFac_Catg=P.Disc_Factor) As PurcDisc_Per, P.ReOrd_Lvl, " & _
                            "(" & mSQry & " And Mrp_Yn=1 And Tax_Yn=1) As Cur_MRP_TBStk, (" & mSQry & " And Mrp_Yn=1 And Tax_Yn=0) As Cur_MRP_TpStk, " & _
                            "(" & mSQry & " And Mrp_Yn=0 And Tax_Yn=1) As Cur_TBStk, (" & mSQry & " And Mrp_Yn=0 And Tax_Yn=0) As Cur_TpStk, " & _
                            "(" & mSQry & ") As CurrStk, P.Min_Lvl, P.Disc_Factor " & _
                            "From Part P Left Join Sp_Stock Stk On P.Part_No=Stk.Part_No " & _
                            "WHERE (V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & " Or Stk.Part_No Is Null)  And Div_Code='" & PubDivCode & "' and P.Part_No='" & PNo & "' " & _
                            "Group By P.Part_No, P.Part_Name, P.Local_Name, P.Unit, P.Mrp, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Disc_Factor, P.ReOrd_Lvl, P.Min_Lvl"
        Else
            mQRY = "Select P.Part_No as Code, P.Part_No, P.Part_Name As Name, P.Local_Name as LName, P.Unit , Format(P.MRP,'0.00') As Mrp, Format(P.TB_SRate,'0.00') As TB_SRate, Format(P.Tp_SRate,'0.00') As Tp_SRate, P.Bin_Loca, " & _
                            "(Select PurcDisc_Per From Part_DiscFactor Where DiscFac_Catg=P.Disc_Factor) As PurcDisc_Per, P.ReOrd_Lvl, " & _
                            "(" & mSQry & " And Mrp_Yn=1 And Tax_Yn=1) As Cur_MRP_TBStk, (" & mSQry & " And Mrp_Yn=1 And Tax_Yn=0) As Cur_MRP_TpStk, " & _
                            "(" & mSQry & " And Mrp_Yn=0 And Tax_Yn=1) As Cur_TBStk, (" & mSQry & " And Mrp_Yn=0 And Tax_Yn=0) As Cur_TpStk, " & _
                            "(" & mSQry & ") As CurrStk, P.Min_Lvl, P.Disc_Factor " & _
                            "From Part P Left Join Sp_Stock Stk On P.Part_No=Stk.Part_No " & _
                            "WHERE (V_Type=" & cIIF("V_Date=" & ConvertDate(DateAdd("D", -1, PubStartDate)), "'SXAO'") & " Or V_Type<>" & cIIF("V_Date>=" & ConvertDate(PubStartDate) & " And V_Date<=" & ConvertDate(PubLoginDate), "'SXAO'") & " Or Stk.Part_No Is Null)  And Div_Code='" & PubDivCode & "' And P.Part_No = '" & PNo & "' " & _
                            "Group By P.Part_No, P.Part_Name, P.Local_Name, P.Unit, P.Mrp, P.TB_SRate, P.Tp_SRate, P.Bin_Loca, P.Disc_Factor, P.ReOrd_Lvl, P.Min_Lvl"
        End If
        
        
        Set Rst = GCn.Execute(mQRY)
        If Rst.RecordCount > 0 Then
            With Rst
                GCn.Execute ("Update Part Set Part.Cur_TP_Stk=" & VNull(!Cur_TPStk) & ", " & _
                             "Part.Cur_TB_Stk=" & VNull(!Cur_TBStk) & ", Part.Cur_Mrp_TpStk=" & VNull(!Cur_MRP_TPStk) & ", " & _
                             "Part.Cur_Mrp_TBStk=" & VNull(!Cur_MRP_TbStk) & " where Part.Part_No='" & !Part_No & "' and Part.Div_Code='" & PubDivCode & "'")
            End With
        End If
    
    GCn.CommitTrans

Set Rst = Nothing
Exit Sub

End Sub

