VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A5C6D027-CC53-4DEC-A683-845894FE1E7D}#6.0#0"; "TOPCTL.OCX"
Begin VB.Form frmVehBook 
   BackColor       =   &H00BAD3C9&
   Caption         =   "Vehicle Booking"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12195
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   12195
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txt 
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
      Height          =   240
      Index           =   52
      Left            =   8865
      MaxLength       =   20
      TabIndex        =   138
      Text            =   "012345678901"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   51
      Left            =   8865
      MaxLength       =   20
      TabIndex        =   133
      Text            =   "012345678901"
      Top             =   2655
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   50
      Left            =   5100
      MaxLength       =   20
      TabIndex        =   5
      Top             =   870
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   49
      Left            =   1665
      TabIndex        =   130
      Top             =   4695
      Width           =   5055
   End
   Begin MSDataGridLib.DataGrid DgModelGroup 
      Height          =   2835
      Left            =   6795
      Negotiate       =   -1  'True
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   9180
      Visible         =   0   'False
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   5001
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Model Group"
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
            ColumnWidth     =   3209.953
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   48
      Left            =   1665
      MaxLength       =   20
      TabIndex        =   22
      Top             =   4185
      Width           =   5055
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   47
      Left            =   5715
      MaxLength       =   6
      TabIndex        =   14
      Text            =   "012345678901234"
      Top             =   2400
      Width           =   1005
   End
   Begin MSDataGridLib.DataGrid DGRef 
      Height          =   4935
      Left            =   5070
      Negotiate       =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   8670
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
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
      Caption         =   "Refer Person Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Reference Person"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGQuot 
      Height          =   3360
      Left            =   4275
      Negotiate       =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   7695
      Visible         =   0   'False
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   5927
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
      Caption         =   "Quotation Help"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Quot. No"
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
         DataField       =   "Prefix"
         Caption         =   "Prefix"
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
         DataField       =   "Party"
         Caption         =   "Party"
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
         DataField       =   "Model"
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
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   5369.953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2369.764
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgAcces 
      Height          =   2175
      Left            =   5460
      Negotiate       =   -1  'True
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   8445
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
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
      Caption         =   "Accessories"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Accessory Name"
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
            ColumnWidth     =   2865.26
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGMod 
      Height          =   2865
      Left            =   300
      Negotiate       =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   8475
      Visible         =   0   'False
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   5054
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
      Caption         =   "Model Help"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Model Code"
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
         DataField       =   "ModelGroup"
         Caption         =   "Model Group"
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
         DataField       =   "Sale_Rate"
         Caption         =   "Rate"
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
         DataField       =   "Chas_Type"
         Caption         =   "Chassis Type"
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
      BeginProperty Column04 
         DataField       =   "Name"
         Caption         =   "Model Name"
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
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   6075.213
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BAD3C9&
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
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   180
      TabIndex        =   112
      Top             =   6255
      Visible         =   0   'False
      Width           =   6360
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   40
         Left            =   1860
         MaxLength       =   12
         TabIndex        =   123
         Top             =   30
         Width           =   1425
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   41
         Left            =   1860
         MaxLength       =   12
         TabIndex        =   122
         Top             =   285
         Width           =   1425
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   44
         Left            =   1860
         MaxLength       =   12
         TabIndex        =   121
         Top             =   540
         Width           =   1425
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   42
         Left            =   4815
         MaxLength       =   12
         TabIndex        =   116
         Top             =   255
         Width           =   1425
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   43
         Left            =   4815
         MaxLength       =   12
         TabIndex        =   115
         Top             =   510
         Width           =   1425
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   45
         Left            =   4815
         MaxLength       =   12
         TabIndex        =   114
         Top             =   765
         Width           =   1425
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   46
         Left            =   4815
         MaxLength       =   12
         TabIndex        =   113
         Top             =   0
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance EMI"
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
         Index           =   21
         Left            =   0
         TabIndex        =   126
         Top             =   60
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration."
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
         Index           =   22
         Left            =   -15
         TabIndex        =   125
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Less Discount"
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
         Index           =   25
         Left            =   0
         TabIndex        =   124
         Top             =   555
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance"
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
         Height          =   255
         Index           =   23
         Left            =   3405
         TabIndex        =   120
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Charges"
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
         Height          =   255
         Index           =   24
         Left            =   3405
         TabIndex        =   119
         Top             =   525
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brokrage"
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
         Height          =   255
         Index           =   29
         Left            =   3405
         TabIndex        =   118
         Top             =   750
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subvention"
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
         Height          =   255
         Index           =   30
         Left            =   3405
         TabIndex        =   117
         Top             =   60
         Width           =   960
      End
   End
   Begin VB.TextBox txtgrid 
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
      Height          =   270
      Index           =   0
      Left            =   11520
      MaxLength       =   40
      TabIndex        =   110
      Top             =   6540
      Visible         =   0   'False
      Width           =   705
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid1 
      Height          =   1695
      Left            =   210
      TabIndex        =   108
      Top             =   8175
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      BackColorFixed  =   12632319
      ForeColorFixed  =   128
      BackColorSel    =   15718112
      ForeColorSel    =   12582912
      BackColorBkg    =   12243913
      GridColor       =   8438015
      GridColorFixed  =   32896
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "ddd"
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
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton CmdPrintInfo 
      BackColor       =   &H00D7C6C8&
      Caption         =   "Print Customer Sales Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   0
      Visible         =   0   'False
      Width           =   3570
   End
   Begin MSDataGridLib.DataGrid DGArea 
      Height          =   4935
      Left            =   4230
      Negotiate       =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   8235
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Area Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
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
         MarqueeStyle    =   5
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGRep 
      Height          =   4515
      Left            =   7515
      Negotiate       =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   7785
      Visible         =   0   'False
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   7964
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
      ColumnCount     =   1
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   39
      Left            =   1665
      MaxLength       =   35
      TabIndex        =   13
      Top             =   2400
      Width           =   3345
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   5
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   10
      Top             =   1635
      Width           =   5055
   End
   Begin MSDataGridLib.DataGrid DGFin 
      Height          =   2970
      Left            =   6405
      Negotiate       =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   7695
      Visible         =   0   'False
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   5239
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
      Caption         =   "Financier Help"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Financier"
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
         DataField       =   "Add1"
         Caption         =   "Address1"
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
         DataField       =   "Add2"
         Caption         =   "Address2"
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
            ColumnWidth     =   4259.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3240
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3089.764
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   38
      Left            =   3135
      TabIndex        =   104
      Top             =   525
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   4635
      TabIndex        =   103
      Top             =   555
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   36
      Left            =   6660
      TabIndex        =   101
      Top             =   435
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   225
      Index           =   33
      Left            =   10155
      MaxLength       =   12
      TabIndex        =   100
      TabStop         =   0   'False
      Text            =   "29/DEC/2003"
      Top             =   1740
      Width           =   1200
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   225
      Index           =   35
      Left            =   10155
      MaxLength       =   12
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   1995
      Width           =   1200
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   225
      Index           =   34
      Left            =   8880
      MaxLength       =   8
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   1995
      Width           =   1230
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   225
      Index           =   32
      Left            =   8880
      MaxLength       =   8
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   1740
      Width           =   1230
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   31
      Left            =   1665
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1125
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   30
      Left            =   1665
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1380
      Width           =   495
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   225
      Index           =   29
      Left            =   8880
      MaxLength       =   20
      TabIndex        =   30
      Top             =   2895
      Width           =   2475
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   28
      Left            =   2175
      MaxLength       =   40
      TabIndex        =   9
      Top             =   1380
      Width           =   4545
   End
   Begin MSDataGridLib.DataGrid DGCity 
      Height          =   4935
      Left            =   7710
      Negotiate       =   -1  'True
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   8025
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8705
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "City Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
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
         MarqueeStyle    =   5
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrmPrn 
      BackColor       =   &H00CAECF0&
      Caption         =   "Printing Option"
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
      Height          =   1605
      Left            =   180
      TabIndex        =   76
      Top             =   7380
      Visible         =   0   'False
      Width           =   5025
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H00CAECF0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4695
         MousePointer    =   99  'Custom
         Picture         =   "frmVehBook.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Delete Current Record"
         Top             =   15
         Width           =   315
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
         Left            =   15
         Picture         =   "frmVehBook.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Screen"
         Top             =   1275
         Width           =   315
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Windows Print"
         DisabledPicture =   "frmVehBook.frx":0678
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
         Left            =   3420
         MaskColor       =   &H00FFC0FF&
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Printer "
         Top             =   990
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Screen"
         DisabledPicture =   "frmVehBook.frx":0982
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
         Left            =   3420
         MaskColor       =   &H00EFD5B8&
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Screen"
         Top             =   660
         Width           =   1590
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8D7FD&
         Caption         =   "Speed &Print"
         DisabledPicture =   "frmVehBook.frx":0C8C
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
         Left            =   3420
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Printer "
         Top             =   330
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
         Index           =   2
         Left            =   7425
         TabIndex        =   86
         Top             =   555
         Visible         =   0   'False
         Width           =   300
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
         Left            =   7365
         TabIndex        =   85
         Top             =   2535
         Visible         =   0   'False
         Width           =   330
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
         Left            =   7470
         TabIndex        =   84
         Top             =   300
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton Optpre 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
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
         Left            =   1725
         TabIndex        =   78
         Top             =   720
         Width           =   1200
      End
      Begin VB.OptionButton OptPlain 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00CAECF0&
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
         Left            =   300
         TabIndex        =   77
         Top             =   720
         Width           =   750
      End
      Begin VB.Line Line8 
         X1              =   1470
         X2              =   1470
         Y1              =   510
         Y2              =   600
      End
      Begin VB.Line Line7 
         X1              =   2820
         X2              =   2820
         Y1              =   630
         Y2              =   735
      End
      Begin VB.Line Line5 
         X1              =   360
         X2              =   360
         Y1              =   615
         Y2              =   720
      End
      Begin VB.Line Line6 
         X1              =   2820
         X2              =   345
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   210
         Index           =   41
         Left            =   -105
         TabIndex        =   89
         Top             =   300
         Width           =   3315
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
         Left            =   330
         TabIndex        =   88
         Top             =   1275
         Width           =   4650
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Printer Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   15
         TabIndex        =   87
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   26
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   12
      Top             =   2145
      Width           =   5055
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   25
      Left            =   1665
      TabIndex        =   15
      Text            =   "0123456789012345678901234"
      Top             =   2655
      Width           =   2760
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   8
      Left            =   8880
      TabIndex        =   31
      Top             =   3150
      Width           =   645
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   7
      Left            =   5010
      MaxLength       =   15
      TabIndex        =   16
      Text            =   "012345678901234"
      Top             =   2655
      Width           =   1710
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   225
      Index           =   21
      Left            =   1665
      MaxLength       =   12
      TabIndex        =   29
      Top             =   5970
      Width           =   1425
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   24
      Left            =   1665
      TabIndex        =   28
      Top             =   5715
      Width           =   5055
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   23
      Left            =   1665
      TabIndex        =   27
      Top             =   5460
      Width           =   5055
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   22
      Left            =   5295
      MaxLength       =   12
      TabIndex        =   26
      Text            =   "012345678901"
      Top             =   5205
      Width           =   1425
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   20
      Left            =   8880
      TabIndex        =   34
      Top             =   3915
      Width           =   645
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
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
      Height          =   225
      Index           =   19
      Left            =   1665
      MaxLength       =   10
      TabIndex        =   25
      Top             =   5205
      Width           =   1425
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   18
      Left            =   1665
      MaxLength       =   20
      TabIndex        =   24
      Top             =   4950
      Width           =   5055
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   17
      Left            =   1665
      MaxLength       =   20
      TabIndex        =   23
      Top             =   4440
      Width           =   5055
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   11
      Left            =   8880
      TabIndex        =   35
      Top             =   4170
      Width           =   1470
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
      Height          =   225
      Index           =   0
      Left            =   8880
      MaxLength       =   12
      TabIndex        =   1
      Top             =   870
      Width           =   2130
   End
   Begin TopCtl.TopCtrl TopCtrl1 
      Align           =   1  'Align Top
      Height          =   384
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   688
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
      Height          =   225
      Index           =   2
      Left            =   9780
      MaxLength       =   9
      TabIndex        =   3
      Top             =   1365
      Width           =   1230
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
      Height          =   210
      Index           =   1
      Left            =   8880
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1125
      Width           =   2130
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   16
      Left            =   1665
      MaxLength       =   20
      TabIndex        =   21
      Top             =   3930
      Width           =   5055
   End
   Begin VB.Frame FrmList 
      BorderStyle     =   0  'None
      Height          =   1830
      Left            =   9135
      TabIndex        =   44
      Top             =   7365
      Visible         =   0   'False
      Width           =   2520
      Begin MSComctlLib.ListView ListView 
         Height          =   1830
         Left            =   0
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   60
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3228
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   4210752
         BackColor       =   16379351
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   4
      Left            =   2175
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1125
      Width           =   4545
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   9
      Left            =   8880
      TabIndex        =   32
      Top             =   3405
      Width           =   645
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   12
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   19
      Top             =   3420
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   14
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   17
      Top             =   2910
      Width           =   5055
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
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   10
      Left            =   8880
      TabIndex        =   33
      Top             =   3660
      Width           =   645
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   13
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   18
      Top             =   3165
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   15
      Left            =   1665
      MaxLength       =   25
      TabIndex        =   20
      Top             =   3675
      Width           =   5055
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   6
      Left            =   1665
      MaxLength       =   40
      TabIndex        =   11
      Top             =   1890
      Width           =   5055
   End
   Begin MSDataGridLib.DataGrid DGProf 
      Height          =   2835
      Left            =   5265
      Negotiate       =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   7755
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   5001
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Profession"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCol 
      Height          =   2445
      Left            =   5715
      Negotiate       =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   8085
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   4313
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
      Caption         =   "Colour Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "name"
         Caption         =   "Colors"
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
            ColumnWidth     =   4380.095
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGSite 
      Height          =   2175
      Left            =   5460
      Negotiate       =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   8445
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
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
      Caption         =   "Site Help"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Site Name"
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
         DataField       =   "code"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         SizeMode        =   1
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   2310.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGPurpose 
      Height          =   2835
      Left            =   4695
      Negotiate       =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   7770
      Visible         =   0   'False
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   5001
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
      Caption         =   "Purpose Help"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Purpose"
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
            ColumnWidth     =   4545.071
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGrid 
      Height          =   1950
      Left            =   7110
      TabIndex        =   36
      Top             =   5715
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   3440
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   4210752
      FixedRows       =   0
      RowHeightMin    =   255
      BackColorFixed  =   13623520
      ForeColorFixed  =   0
      BackColorSel    =   15261111
      BackColorBkg    =   13300221
      GridColor       =   13623520
      GridColorFixed  =   13623520
      GridColorUnpopulated=   13623520
      GridLinesFixed  =   1
      BorderStyle     =   0
      Appearance      =   0
      GridLineWidthFixed=   2
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLineWidthBand=   1
   End
   Begin MSDataGridLib.DataGrid DGVno 
      Height          =   2175
      Left            =   5625
      Negotiate       =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   9090
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Code"
         Caption         =   "Voucher No"
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
            ColumnWidth     =   2865.26
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DgParty 
      Height          =   2955
      Left            =   3750
      Negotiate       =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   8985
      Visible         =   0   'False
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   5212
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
      Caption         =   "Party Help"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Party Name"
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
         DataField       =   "FName"
         Caption         =   "Father Name"
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
         DataField       =   "Add1"
         Caption         =   "Address"
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
         DataField       =   "CityName"
         Caption         =   "City"
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
            ColumnWidth     =   4545.071
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   27
      Left            =   1665
      MaxLength       =   8
      TabIndex        =   92
      Top             =   870
      Width           =   750
   End
   Begin VB.TextBox txt 
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
      Height          =   225
      Index           =   3
      Left            =   2430
      MaxLength       =   8
      TabIndex        =   4
      Top             =   870
      Width           =   1170
   End
   Begin VB.TextBox lblGroup 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   1065
      TabIndex        =   105
      Top             =   360
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label LblDoNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DoNo"
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
      Index           =   4
      Left            =   4485
      TabIndex        =   137
      Top             =   885
      Width           =   465
   End
   Begin VB.Label LblDoNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DO Issue Date."
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
      Height          =   540
      Index           =   3
      Left            =   0
      TabIndex        =   136
      Top             =   0
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label LblDoNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DO Recive Date."
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
      Index           =   2
      Left            =   7170
      TabIndex        =   135
      Top             =   2655
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label LblDoNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DO Issue Date."
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
      Left            =   0
      TabIndex        =   134
      Top             =   15
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LblDoNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DO Issue Date."
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
      Left            =   7155
      TabIndex        =   132
      Top             =   2430
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Desc"
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
      Index           =   36
      Left            =   105
      TabIndex        =   131
      Top             =   4710
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Group"
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
      Index           =   35
      Left            =   105
      TabIndex        =   128
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pin No"
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
      Index           =   34
      Left            =   5100
      TabIndex        =   127
      Top             =   2415
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accessories Commited"
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
      Index           =   28
      Left            =   10290
      TabIndex        =   109
      Top             =   6585
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      Index           =   20
      Left            =   105
      TabIndex        =   106
      Top             =   2415
      Width           =   525
   End
   Begin VB.Label lblVPrefix2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QuotPrefix"
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
      Left            =   0
      TabIndex        =   102
      Top             =   420
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Doc No."
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
      Left            =   7170
      TabIndex        =   98
      Top             =   2010
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No."
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
      Index           =   18
      Left            =   7170
      TabIndex        =   96
      Top             =   1755
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAN"
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
      Left            =   7170
      TabIndex        =   94
      Top             =   2895
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name"
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
      Left            =   105
      TabIndex        =   93
      Top             =   1395
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City*"
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
      Index           =   15
      Left            =   90
      TabIndex        =   74
      Top             =   2670
      Width           =   450
   End
   Begin VB.Label LblVPrefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Prefix"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   8895
      TabIndex        =   72
      Top             =   1395
      Width           =   675
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   810
      Left            =   6945
      Top             =   825
      Width           =   4215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
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
      Index           =   38
      Left            =   4455
      TabIndex        =   71
      Top             =   2670
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Govt YN"
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
      Index           =   32
      Left            =   7170
      TabIndex        =   70
      Top             =   3165
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finance Amount"
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
      Index           =   14
      Left            =   105
      TabIndex        =   62
      Top             =   5985
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financier"
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
      Left            =   105
      TabIndex        =   61
      Top             =   5730
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Source"
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
      Left            =   105
      TabIndex        =   60
      Top             =   5475
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expected Delv. Date"
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
      Index           =   10
      Left            =   3435
      TabIndex        =   59
      Top             =   5220
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Permit Req YN"
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
      Index           =   9
      Left            =   7170
      TabIndex        =   58
      Top             =   3930
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Rate"
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
      Index           =   8
      Left            =   105
      TabIndex        =   57
      Top             =   5220
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
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
      Index           =   7
      Left            =   105
      TabIndex        =   56
      Top             =   4965
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model*"
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
      Index           =   6
      Left            =   105
      TabIndex        =   55
      Top             =   4455
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(N)ational/(Z)onal"
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
      Index           =   5
      Left            =   7170
      TabIndex        =   54
      Top             =   4185
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Intended Use"
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
      Index           =   4
      Left            =   105
      TabIndex        =   53
      Top             =   3945
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profession"
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
      Index           =   2
      Left            =   105
      TabIndex        =   52
      Top             =   3435
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No."
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
      Index           =   3
      Left            =   7170
      TabIndex        =   51
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name"
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
      Index           =   26
      Left            =   7185
      TabIndex        =   48
      Top             =   885
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Veh YN"
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
      Left            =   7170
      TabIndex        =   47
      Top             =   3420
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address Verif YN"
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
      Index           =   13
      Left            =   7170
      TabIndex        =   46
      Top             =   3675
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose"
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
      Index           =   37
      Left            =   105
      TabIndex        =   43
      Top             =   3690
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quotation No."
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
      Index           =   45
      Left            =   105
      TabIndex        =   42
      Top             =   885
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name*"
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
      Index           =   46
      Left            =   105
      TabIndex        =   41
      Top             =   1140
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Executive*"
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
      Index           =   31
      Left            =   105
      TabIndex        =   40
      Top             =   2925
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reffered By"
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
      Index           =   33
      Left            =   105
      TabIndex        =   39
      Top             =   3180
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Index           =   27
      Left            =   7185
      TabIndex        =   38
      Top             =   1133
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   105
      TabIndex        =   37
      Top             =   1650
      Width           =   690
   End
End
Attribute VB_Name = "frmVehBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CustDocId As String
Dim RsVno As ADODB.Recordset
Dim RsParty As ADODB.Recordset
Dim rsFin As ADODB.Recordset
Dim RsMod As ADODB.Recordset
Dim RsModelGroup As ADODB.Recordset
Dim RsPurpose As ADODB.Recordset
Dim RsCol As ADODB.Recordset
Dim RsCity As ADODB.Recordset
Dim RsQuot As ADODB.Recordset
Dim RsRef As ADODB.Recordset
Dim RsRep As ADODB.Recordset
Dim RsProf As ADODB.Recordset
Dim RsArea As ADODB.Recordset
Dim RsSite As ADODB.Recordset
Dim Master As ADODB.Recordset
Dim RsAcces As ADODB.Recordset

Dim DocID As String * 21
Dim QSiteCode As String
Dim QDocSrlNo As Byte
Dim QDocId As String * 21
Dim mVType As String
Dim VoucherEditFlag As Boolean
Dim vPrefix As String
Dim FinAcCode As String

Private Const Site          As Byte = 0
Private Const BookNo        As Byte = 2
Private Const VDate         As Byte = 1
Private Const QuotNo        As Byte = 3
Private Const Party_code    As Byte = 4
Private Const Add1          As Byte = 5
Private Const Add2          As Byte = 6
Private Const Area          As Byte = 7
Private Const GovtYn        As Byte = 8
Private Const FirstVeh      As Byte = 9
Private Const AddVerYN      As Byte = 10
Private Const NZ            As Byte = 11
Private Const Profession    As Byte = 12
Private Const REF_CODE      As Byte = 13
Private Const REP_CODE      As Byte = 14
Private Const Purpose       As Byte = 15
Private Const IndUse        As Byte = 16
Private Const Model         As Byte = 17
Private Const Colours       As Byte = 18
Private Const Rate          As Byte = 19
Private Const ReqPermit     As Byte = 20
Private Const Amt           As Byte = 21
Private Const ExpDelDt      As Byte = 22
Private Const FundSource    As Byte = 23
Private Const FB_Code       As Byte = 24
Private Const Add3          As Byte = 26
Private Const City          As Byte = 25
Private Const QuotPrefix    As Byte = 27
Private Const fname         As Byte = 28
Private Const PAN           As Byte = 29
Private Const FNamePrefix   As Byte = 30
Private Const NamePrefix    As Byte = 31
Private Const InvNo         As Byte = 32
Private Const InvDate       As Byte = 33
Private Const DelChNo       As Byte = 34
Private Const DelChDate     As Byte = 35
Private Const TxtDocID      As Byte = 36
Private Const QuotDocId     As Byte = 37
Private Const SerialNo2     As Byte = 38
Private Const Phone         As Byte = 39
Private Const AdvEMI        As Byte = 40
Private Const RegAmt        As Byte = 41
Private Const InsuAmt       As Byte = 42
Private Const OthAmt        As Byte = 43
Private Const DiscAmt       As Byte = 44
Private Const Brok          As Byte = 45
Private Const SubVen        As Byte = 46
Private Const PinCode       As Byte = 47
Private Const ModelGroup    As Byte = 48
Private Const ModelDesc     As Byte = 49
Private Const DONo          As Byte = 50
Private Const DOReciveDate  As Byte = 51
Private Const DOIssueDate   As Byte = 52

Dim ForeColorSelEnter$
Dim BackColorSelLeave$
Dim TAddMode As Boolean

'Fgrid1 Columns
Private Const Col_AccCode       As Byte = 1
Private Const Col_Accessory     As Byte = 2
Private Const Col_Qty           As Byte = 3

Private Const SiteCode1     As Byte = 0
Private Const FromVno       As Byte = 1
Private Const ToVno         As Byte = 2

Dim ListArray As Variant
Dim mListItem As ListItem

Private Const PWindows As Byte = 0
Private Const PScreen As Byte = 1
Private Const PDos As Byte = 2
Private Const PClose As Byte = 3
Private Const PSetUp As Byte = 4
Dim mRepName As String

Private Sub CmdPrintInfo_Click()
Dim Rst As ADODB.Recordset, RstSub1 As ADODB.Recordset, RstSub2 As ADODB.Recordset
Dim RstSub3 As ADODB.Recordset, mQry, Accessories As String
Dim I As Integer, Rst2, RST3 As ADODB.Recordset
On Error GoTo ERRORHANDLER

     'If IsValid(txtPrint(Model), "Model") = False Then Exit Sub
     'If IsValid(txtPrint(ChasNo), "Chassis No") = False Then Exit Sub

     mQry = "SELECT VO.STAMP_DUTY,VO.model,VO.REG_FEE, VO.INS_FEE, VO.S_CHARGE, VO.Rate as Net_AMOUNT, " & _
        " VO.MISC_INFO, City.CityName,  " & _
        " Model.Model_Desc, Model.Model_Desc1, " & _
        " SubGroup.Name, SubGroup.Add1, SubGroup.Add2, SubGroup.Add3, SubGroup.PIN,SubGroup.Phone,SubGroup.Mobile,  " & _
        " VO.OrdDocId, VO.Ord_Date,  " & _
        " VO.OtherChrg,  VO.FIN_AMT, VO.Interest, VO.AdvEMI,VO.OtherChrg,VO.Rebate,SubGroup.FName,VO.DelCh_Dt as DelDt,Vo.Chassis,CF.FinName,VO.DelCh_No as DelNo,VO.Brokrage,VO.SubVention,VO.AmtRecd " & _
        " FROM (((Veh_Order VO LEFT JOIN Model ON VO.MODEL = Model.MODEL) " & _
        " LEFT JOIN SubGroup ON VO.PartyCode = SubGroup.SubCode) " & _
        " LEFT JOIN City ON SubGroup.CityCode = City.CityCode) " & _
        " LEFT JOIN ContractFinance CF on VO.FB_CODE=CF.FinCode " & _
        " where VO.OrdDocId = '" & Master!SearchCode & "'"

    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub



   'Recordset is made for subreport2

    mQry = "SELECT Rect.Prov_No,Rect.Ord_DocId, Rect.V_Type, Rect.V_No, Rect.V_Date, Rect.Site_Code, Rect.AMOUNT, Rect.DrCr, Rect.Narration " & _
        " FROM Veh_Order " & _
        " LEFT JOIN Rect ON Veh_Order.OrdDocId = Rect.Ord_DocId " & _
        " where Veh_Order.OrdDocId = '" & Master!SearchCode & "'"

   Set RstSub2 = New Recordset
   RstSub2.CursorLocation = adUseClient
   RstSub2.Open (mQry), GCn, adOpenDynamic, adLockOptimistic

    mRepName = "CustSalesInfor"
    
    If GCn.Execute("Select * from Rect  Where Ord_DocId='" & Master!SearchCode & "'").RecordCount > 0 Then
        Rst!AmtRecd = GCn.Execute("Select Sum(Amount) from Rect  Where Ord_DocId='" & Master!SearchCode & "'").Fields(0).Value
        Rst.Update
    Else
        Rst!AmtRecd = 0
        Rst.Update
    End If
    

    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True

    CreateFieldDefFile RstSub2, PubRepoPath + "\" & mRepName & "2.ttx", True

    

    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
    rpt.Database.SetDataSource Rst

    rpt.OpenSubreport("SUBREP2").Database.SetDataSource RstSub2

    

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

    
        For I = 1 To rpt.OpenSubreport("SUBREP2").FormulaFields.Count
            Select Case UCase(rpt.OpenSubreport("SUBREP2").FormulaFields(I).FormulaFieldName)
                Case UCase("SpeedPrint")
                    rpt.OpenSubreport("SUBREP1").FormulaFields(I).TEXT = "'1'"
            End Select
        Next

    rpt.ReadRecords
    
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, Col_Accessory) <> "" Then
            If Accessories = "" Then
                Accessories = Accessories & GCn.Execute("Select Prod_Name from Veh_Amdmodel where Prod_Code='" & FGrid1.TextMatrix(I, Col_AccCode) & "'").Fields(0).Value & " ( " & FGrid1.TextMatrix(I, Col_Qty) & " ) "
            Else
                Accessories = Accessories & " , " & GCn.Execute("Select Prod_Name from Veh_Amdmodel where Prod_Code='" & FGrid1.TextMatrix(I, Col_AccCode) & "'").Fields(0).Value & " ( " & FGrid1.TextMatrix(I, Col_Qty) & " ) "
            End If
        End If
    Next

      'Printer
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
                Case UCase("AccComit")
                    rpt.FormulaFields(I).TEXT = "'" & Accessories & "'"
            End Select
            Next
        '    rpt.PrintOut False
        'Case PScreen  'screen
            Call Report_View(rpt, Me.CAPTION, , False)
        'Case PDos
        '    Call Report_View(rpt, Me.CAPTION, 1)

Set Rst = Nothing
Set Rst2 = Nothing
Set RST3 = Nothing
Set rpt = Nothing
CmdPrint(PSetUp).Tag = ""
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub

Private Sub DgAcces_Click()
DgAcces.Visible = False
TxtGrid(0).Visible = False
    If RsAcces.RecordCount > 0 Then
        With FGrid1
            .TextMatrix(.Row, Col_Accessory) = RsAcces!Name
            .TextMatrix(.Row, Col_AccCode) = RsAcces!Code
            .Col = Col_Qty
            .SetFocus
        End With
    End If
End Sub

Private Sub DGParty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'lblGroup.TEXT = G_FaCn.Execute("Select AcGroup.GroupName from (AcGroup Left Join SubGroup on SubGroup.GroupCode=AcGroup.GroupCode) where SubGroup.SubCode='" & RsParty!Code & "'").Fields(0).Value
'lblGroup.Refresh
End Sub

Private Sub DGPurpose_Click()
If RsPurpose.RecordCount > 0 Then
    Txt(Purpose).TEXT = RsPurpose!Name
    Txt(Purpose).Tag = RsPurpose!Code
End If
Txt(Purpose).SetFocus
DGPurpose.Visible = False
End Sub

Private Sub DGVno_Click()
Dim Index As Integer
If DGVno.Tag = "1" Then
    Index = ToVno
Else
    Index = FromVno
End If
    DGVno.Visible = False
    If RsVno.RecordCount > 0 Then
        txtPrint(Index).TEXT = RsVno!Code
    End If
    txtPrint(Index).SetFocus
End Sub

Private Sub DGSite_Click()
If FrmPrn.Visible = False Then
    DGSite.Visible = False
    If RsSite.RecordCount > 0 Then
        Txt(Site).TEXT = RsSite!Name
        Txt(Site).Tag = RsSite!Code
    End If
    Txt(Site).SetFocus
Else
    DGSite.Visible = False
    If RsSite.RecordCount > 0 Then
        txtPrint(SiteCode1).TEXT = RsSite!Name
        txtPrint(SiteCode1).Tag = RsSite!Code
    End If
    txtPrint(SiteCode1).SetFocus
End If
End Sub
Private Sub FGrid_Click()
    If FGrid.Col = 1 And FGrid.RowHeight(FGrid.Row) <> 0 Then
        FGrid.Col = 1
        FGrid.CellFontName = "WINGDINGS"
        FGrid.CellFontSize = 14
        FGrid.TextMatrix(FGrid.Row, 1) = IIf(FGrid.TextMatrix(FGrid.Row, 1) = "ü", " ", "ü")
    End If
End Sub

Private Sub FGrid_EnterCell()
FGrid.CellBackColor = CtrlBCol
End Sub

Private Sub FGrid_GotFocus()
Grid_Hide
FGrid.CellBackColor = CtrlBCol
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Val(FGrid.Tag) <> FGrid.Rows - 1 Then
     FGrid.CellBackColor = CtrlBColOrg
    FGrid.Row = FGrid.Row + 1
     FGrid.CellBackColor = CtrlBCol
    End If
End If
If KeyCode = vbKeyUp And Val(FGrid.Tag) = 0 Then
    FGrid.CellBackColor = CtrlBColOrg
    SendKeys "+{Tab}"
ElseIf (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Val(FGrid.Tag) = FGrid.Rows - 1 Then
    FGrid.CellBackColor = CtrlBColOrg
    If MsgBox("Save Record ?", vbYesNo, "Save Data") = vbYes Then
        FGrid.CellBackColor = CtrlBColOrg
        TopCtrl1_eSave
        Exit Sub
    Else
        FGrid.SetFocus
        FGrid.CellBackColor = CtrlBCol
    End If
    
End If
FGrid.Tag = FGrid.Row
End Sub
Private Sub FGrid_LeaveCell()
    FGrid.CellBackColor = CtrlBColOrg
End Sub

Private Sub FGrid_Validate(Cancel As Boolean)
    FGrid.CellBackColor = CtrlBColOrg
End Sub

Private Sub FGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then
    FGrid.Col = 1
    FGrid.CellFontName = "WINGDINGS"
    FGrid.CellFontSize = 14
    FGrid.TextMatrix(FGrid.Row, 1) = IIf(FGrid.TextMatrix(FGrid.Row, 1) = "ü", " ", "ü")
End If
End Sub

Private Sub FGrid1_KeyPress(KeyAscii As Integer)
Select Case FGrid1.Col
    Case Col_Accessory
       Call Get_Text(Me, FGrid1, TxtGrid, 0, False, KeyAscii)
    Case Col_Qty
       Call Get_Text(Me, FGrid1, TxtGrid, 0, True, KeyAscii)
End Select
'If KeyAscii <> vbKeyReturn Then TAddMode = True
End Sub

Private Sub FGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
If TopCtrl1.TopText2.CAPTION = "Browse" Then Exit Sub
If FGrid1.ColSel = False Then Exit Sub
If KeyCode = vbKeyD And Shift = 2 Then
    If FGrid1.Row >= 1 Then
        If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
            If FGrid1.Rows > 2 Then
                FGrid1.RemoveItem (FGrid1.Row)
            Else
                FGrid1.Rows = 1
                FGrid1.AddItem FGrid1.Rows
                FGrid1.FixedRows = 1
            End If
         End If
         For I = 1 To FGrid1.Rows - 1
            FGrid1.TextMatrix(I, 0) = I
         Next
    Else
        MsgBox "No Entries To Delete", vbCritical, "Delete Module"
    End If
   
FGrid1.SetFocus
End If
Exit Sub
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
MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Load()
On Error GoTo ELoop

    TopCtrl1.Tag = PubUParam: WinSetting Me
    
    Set Master = New ADODB.Recordset
    Master.CursorLocation = adUseClient
    
    
    If UCase(left(PubComp_Name, 6)) = "J.M.A." Then
        LblDoNo(0).Visible = True
        Txt(DONo).Visible = True
        LblDoNo(2).Visible = True
        Txt(51).Visible = True
        Txt(52).Visible = True
    End If
    
    
    Dim sitecond As String
    sitecond = " And Ord_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
        sitecond = sitecond & " and " & cMID("Veh_Order.OrdDocId", "3", "1") & " ='" & PubSiteCode & "'"
    End If
    
    
    If PubMoveRecYn Then
        Master.Open "Select OrdDocId as searchcode, OrdDocId from Veh_Order where Ord_VType = 'V_BK' and left(OrdDocId,1)='" & PubDivCode & "' and (Ord_Date>=" & ConvertDate(PubStartDate) & " Or Inv_Date is Null) " & sitecond & " order by Right(OrdDocId,8), Ord_Date desc", GCn, adOpenDynamic, adLockOptimistic
    Else
        Set Master = GCn.Execute("Select Top 1 OrdDocId as searchcode, OrdDocId from Veh_Order where Ord_VType = 'V_BK' and left(OrdDocId,1)='" & PubDivCode & "' and (Ord_Date>=" & ConvertDate(PubStartDate) & " Or Inv_Date is Null)  " & sitecond & " order by Right(OrdDocId,8), Ord_Date desc")
    End If
    
   
   
    Set RsRef = New ADODB.Recordset
    RsRef.CursorLocation = adUseClient
    RsRef.Open "select RefCode as code,RefName as name from reffered order by Refname", GCn, adOpenDynamic, adLockOptimistic
    Set DGRef.DataSource = RsRef
    
    
    Set RsQuot = New ADODB.Recordset
    RsQuot.CursorLocation = adUseClient
    RsQuot.Open "SELECT Veh_Quot1.DocId AS code, " & cCStr("Veh_Quot1.V_No") & " AS name, " & cMID("Veh_Quot1.docid", "9", "13") & " AS Prefix, Veh_Quot1.MODEL, (ProspectiveCust.Name + ' ' +  ProspectiveCust.NSuffix) as party " & _
    "FROM (Veh_Quot1 LEFT JOIN Veh_Quot ON Veh_Quot1.DocId = Veh_Quot.DocId) LEFT JOIN ProspectiveCust ON Veh_Quot.Party_Code = ProspectiveCust.Cust_Code " & _
    "WHERE left(Veh_Quot1.DocId,1)='" & PubDivCode & "' and (((Veh_Quot1.Book_DocId)='')) ORDER BY Veh_Quot1.V_No", GCn, adOpenDynamic, adLockOptimistic
    
    
    
    Set DGQuot.DataSource = RsQuot
    
    Set RsCity = New ADODB.Recordset
    RsCity.CursorLocation = adUseClient
    RsCity.Open "select citycode as code ,cityname as name from city order by cityname", GCn, adOpenDynamic, adLockOptimistic
    Set DGCity.DataSource = RsCity
    
    
    
    Set RsVno = New ADODB.Recordset
    RsVno.CursorLocation = adUseClient
    RsVno.Open "Select distinct Ord_No as code from Veh_Order where Ord_VType = 'V_BK' and left(OrdDocId,1)='" & PubDivCode & "'", GCn, adOpenDynamic, adLockOptimistic
    Set DGVno.DataSource = RsVno
    
    
    

    Set RsSite = New ADODB.Recordset
    RsSite.CursorLocation = adUseClient
    RsSite.Open "select site_code as code,site_desc as name from site order by site_desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGSite.DataSource = RsSite
    
    
    
    Set RsPurpose = New ADODB.Recordset
    RsPurpose.CursorLocation = adUseClient
    RsPurpose.Open "select PurposeCode as code,Purposename as name from Purpose order by PurposeName", GCn, adOpenDynamic, adLockOptimistic
    Set DGPurpose.DataSource = RsPurpose
    
    
    
    Set rsFin = New ADODB.Recordset
    rsFin.CursorLocation = adUseClient
    rsFin.Open "select fincode as code,finname + ',' + " & xIsNull("City.CityName", "") & " as name,Add1,Add2,AcCode,FinBankCode from ContractFinance " & _
    "left join city on left(ContractFinance.City,4)=City.CityCode where fincatg = 0  order by finname", GCn, adOpenDynamic, adLockOptimistic
    Set DGFin.DataSource = rsFin
  
    
  
  
    Set RsArea = New ADODB.Recordset
    RsArea.CursorLocation = adUseClient
    RsArea.Open "select AreaCode as code,AreaName as name from Area order by AreaName", GCn, adOpenDynamic, adLockOptimistic
    Set DGArea.DataSource = RsArea
   
    
   
    Set RsProf = New ADODB.Recordset
    RsProf.CursorLocation = adUseClient
    RsProf.Open "select ProfessionCode as code,Professionname as name from Profession order by Professionname", GCn, adOpenDynamic, adLockOptimistic
    Set DGProf.DataSource = RsProf
  
  
    
  
    Set RsRep = New ADODB.Recordset
    RsRep.CursorLocation = adUseClient
    RsRep.Open "select Emp_code as code,emp_name as name from emp_mast where emp_type = 0  order by Emp_name", GCn, adOpenDynamic, adLockOptimistic
    Set DGRep.DataSource = RsRep
    
    
    

    Set RsParty = New ADODB.Recordset
    RsParty.CursorLocation = adUseClient
    RsParty.Open "select S.SubCode as code,S.NAME,S.add1,S.add2,S.add3,S.Phone,S.CityCode,S.PanNo,S.FPrefix,S.FName,S.NamePrefix,City.CityName from SubGroup as S left join City on S.CityCode=City.CityCode Where S.Nature='Customer' Order by S.Name", GCn, adOpenDynamic, adLockOptimistic
    Set DGParty.DataSource = RsParty
    
    
    
    Set RsModelGroup = GCn.Execute("Select ModelGrp_Code As Code, ModelGrp_Name As Name From Model_Grp Order By ModelGrp_Name")
    Set DgModelGroup.DataSource = RsModelGroup
    
    
    
    If PubSiebelActiveYn = 1 Then
        Set RsMod = GCn.Execute("select Model as code,ModelGrp_Name as ModelGroup, Grp_Code,Col_desc  as Colour,Model_Desc as NAME, Chas_Type, Sale_Rate from (Model Left join Model_Grp on model.Grp_Code=Model_Grp.ModelGrp_Code) Left Join ColMast on Model.Col_Code=ColMast.Col_Code where (Model.Div_Code='" & PubDivCode & "' Or IsNull(Model.Div_Code,'') ='') order by Model")
        Set DGMod.DataSource = RsMod
    Else
        Set RsMod = GCn.Execute("select Model as code,Model_Desc as NAME, ModelGrp_Name As ModelGroup, Grp_Code, Chas_Type,ColMast.Col_Desc, Sale_Rate  from (model Left Join ColMast on Model.Col_Code=ColMast.Col_Code) Left Join Model_Grp On Model.Grp_Code=Model_Grp.ModelGrp_Code  where (Model.div_code='" & PubDivCode & "' Or IsNull(Model.Div_Code,'') ='') order by model")
        Set DGMod.DataSource = RsMod
        DGMod.Columns(1).width = 0
        DGMod.Columns(2).width = 0
    End If
    
    
    
    
    Set RsCol = New ADODB.Recordset
    RsCol.CursorLocation = adUseClient
    RsCol.Open "select Col_code as code,col_Desc  as name from colmast order by col_Desc", GCn, adOpenDynamic, adLockOptimistic
    Set DGCol.DataSource = RsCol
    
    
    
    Set RsAcces = New ADODB.Recordset
    RsAcces.CursorLocation = adUseClient
    RsAcces.Open "Select Prod_code as code,Prod_Name as Name FROM Veh_AmdModel Order by Prod_name", GCn, adOpenDynamic, adLockOptimistic
    Set DgAcces.DataSource = RsAcces
    
    
    
    mVType = "V_BK"
    Ini_Grid
    MoveRec
    Disp_Text SETS("INI", Me, Master)
    If UCase(left(PubComp_Name, 7)) = "SOCIETY" Then
        CmdPrintInfo.Visible = True
        FGrid1.Visible = True
        Label3(28).Visible = True
        Frame1.Visible = True
    End If
   Exit Sub
ELoop:    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If TopCtrl1.TopText2 <> "Browse" Then
        If MsgBox("Do you want to exit", vbExclamation + vbYesNo) = vbYes Then
            Exit Sub
        Else
            Cancel = 1
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsSite = Nothing
Set RsParty = Nothing
Set rsFin = Nothing
Set RsMod = Nothing
Set RsProf = Nothing
Set RsRep = Nothing
Set RsRef = Nothing
Set RsArea = Nothing
Set RsQuot = Nothing
Set RsCol = Nothing
Set RsVno = Nothing
Set Master = Nothing
Set RsPurpose = Nothing
Set mListItem = Nothing
End Sub

Private Sub ListView_Click()
Txt(Val(ListView.Tag)).TEXT = ListView.SelectedItem.TEXT
FrmList.Visible = False
Txt(Val(ListView.Tag)).SetFocus
End Sub

Private Sub TopCtrl1_eAdd()
On Error GoTo ErrorLoop
Dim I As Integer
    Disp_Text SETS("ADD", Me, Master)
    Call BlankText
    For I = 0 To FGrid.Rows - 1
        FGrid.TextMatrix(I, 1) = ""
    Next
    CustDocId = ""
    Txt(GovtYn) = "No"
    If UCase(left(PubComp_Name, 4)) = "ENAR" Then
        Txt(Site).Tag = PubSiteCode
        Txt(Site) = PubSiteName
        Txt(VDate).SetFocus
    Else
        Txt(Site).SetFocus
    End If
    Exit Sub
ErrorLoop:
    CheckError
End Sub

Private Sub TopCtrl1_eDel()
Dim mTrans As Boolean
On Error GoTo eloop1
    If GCn.Execute("Select  Inv_DocId from  veh_order where OrdDocId = '" & Master!SearchCode & "'").Fields(0).Value <> "" Then
        MsgBox "Invoice made against this Booking, Delete denied", vbInformation, "Deletion Denied": Exit Sub
    End If
    If GCn.Execute("Select Ord_DocId from  Rect where Ord_DocId = '" & Master!SearchCode & "'").RecordCount > 0 Then
        MsgBox "Receipt has been made against this Booking, Delete denied", vbInformation, "Deletion Denied": Exit Sub
    End If
    
If MsgBox("Are You Sure To Delete This ? ", vbYesNo + vbCritical + vbDefaultButton2, "Delete Entry !") = vbYes Then
    GCn.BeginTrans
    mTrans = True
    GCn.Execute "update veh_quot1 set Book_SiteCode= '',Book_docid='' where docid = '" & QDocId & "' and srl_no = " & QDocSrlNo & ""
    GCn.Execute ("delete from veh_order where  OrdDocId = '" & Master!OrdDocId & "'")
    GCn.CommitTrans
    mTrans = False
    RsQuot.Requery
    Master.Requery
    Call MoveRec
    BUTTONS True, Me, Master, 0
End If
Exit Sub
eloop1:
    If mTrans Then GCn.RollbackTrans
    CheckError
End Sub

Private Sub TopCtrl1_eEdit()
 On Error GoTo eloop1
    If Txt(InvNo) <> "" Then MsgBox "Invoice Made, Edit denied !", vbInformation, "Validation": Exit Sub
    If IsEditable(RetDate(Txt(VDate))) = False Then Exit Sub
    Disp_Text SETS("EDIT", Me, Master)
    Txt(Party_code).SetFocus
    Exit Sub
eloop1:
    If err.NUMBER <> 0 Then
        MsgBox err.Description, vbExclamation, " Editing Message"
    End If
End Sub

Private Sub TopCtrl1_eExit()
    Unload Me
End Sub

Private Sub TopCtrl1_eFind()
On Error GoTo ErrorLoop
    If Master.RecordCount <= 0 Then MsgBox "No Records To Search.", vbInformation, "Information": Exit Sub
    
    Dim sitecond As String
    sitecond = " And Ord_Date Between " & ConvertDate(PubStartDate) & " and " & ConvertDate(PubEndDate) & " "
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
        sitecond = sitecond & " and " & cMID("Veh_Order.OrdDocId", "3", "1") & " ='" & PubSiteCode & "'"
    End If


    GSQL = "Select OrdDocId as searchcode,OrdDocId, Site.Site_Desc as [Site_Name], " & cCStr("Ord_No", 8) & " As Ord_No, " & cDt("Ord_Date") & " As Ord_Date,subgroup.name,  " & _
    " PURPOSE,Veh_Order.MODEL, M.Model_Desc as [Model_Name], " & cDt("EXP_DATE") & " As Exp_Date,Veh_Order.GOVT_YN " & _
    " from Veh_Order left join subgroup on subgroup.subcode = veh_order.partycode " & _
    " LEFT JOIN Site ON Site.Site_Code = SubString(OrdDocID,3,1) " & _
    " LEFT JOIN Model M ON Veh_Order.MODEL = M.MODEL " & _
    " where Ord_VType = 'V_BK' and left(OrdDocId,1)='" & PubDivCode & "' " & sitecond & " order by Ord_Date"
    Set SearchForm = Me
    FAFind.IsNonFaFind = True
    FAFind.Show vbModal
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub
Public Sub SEARCHBACK(ByVal MyValue As String)
On Error GoTo ErrorLoop
    If PubMoveRecYn Then
        Master.MoveFirst
        Master.FIND ("searchcode='" & MyValue & "'")
    Else
        Set Master = GCn.Execute("Select OrdDocId as searchcode, OrdDocId from Veh_Order where Ord_VType = 'V_BK' and left(OrdDocId,1)='" & PubDivCode & "' and (Ord_Date>=" & ConvertDate(PubStartDate) & " Or Inv_Date is Null)  And OrdDocId = '" & MyValue & "'  order by Right(OrdDocId,8), Ord_Date desc")
    End If
    BUTTONS True, Me, Master, 0
    MoveRec
    Exit Sub
ErrorLoop:
    MsgBox err.Name, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_eFirst()
  BUTTONS True, Me, Master, 1
  Call MoveRec
End Sub

Private Sub TopCtrl1_eLast()
 BUTTONS True, Me, Master, 4
 Call MoveRec
End Sub

Private Sub TopCtrl1_eNext()
 BUTTONS True, Me, Master, 3
 Call MoveRec
End Sub

Private Sub TopCtrl1_ePrev()
 BUTTONS True, Me, Master, 2
 Call MoveRec
End Sub

Private Sub TopCtrl1_eCancel()
Dim I As Integer
On Error GoTo ErrorLoop
If MsgBox("Cancel ?", vbYesNo, "Terminate Process") = vbYes Then
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
Else
    Me.ActiveControl.SetFocus
End If
Exit Sub
ErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub TopCtrl1_ePrn()
FrmPrn.top = 2220
FrmPrn.left = (Me.width - FrmPrn.width) / 2
FrmPrn.Visible = True
FrmPrn.ZOrder 0
OptPlain.Value = True
LblPrinter.CAPTION = Printer.DeviceName
If TopCtrl1.TopText2 <> "Browse" Then CmdPrint(PScreen).Enabled = False Else CmdPrint(PScreen).Enabled = True
If PubSpeedPrint = True Then CmdPrint(PDos).SetFocus Else CmdPrint(PWindows).SetFocus

End Sub

Private Sub TopCtrl1_eRef()
    RsParty.Requery
    RsCity.Requery
    rsFin.Requery
    RsMod.Requery
    RsRep.Requery
    RsRef.Requery
    RsProf.Requery
    RsArea.Requery
    RsQuot.Requery
    RsCol.Requery
    RsPurpose.Requery
End Sub

Private Sub TopCtrl1_eSave()
    Dim I As Integer
    Dim Rst As Recordset
    Dim mTrans As Boolean
    Dim TSQL$
    Dim DocIdHlp$
    Dim ac_str$
    Dim AccCode$
    Dim AccQty$
    Dim mFundSource As Integer
On Error GoTo errlbl
    If IsEditable(RetDate(Txt(VDate))) = False Then Exit Sub
    If IsValid(Txt(Site), "Site Name") = False Then Exit Sub
    If IsValid(Txt(Party_code), "Party Name") = False Then Exit Sub
    If IsValid(Txt(VDate), "Date") = False Then Exit Sub
    If IsValid(Txt(BookNo), "Booking No.") = False Then Exit Sub
    'If IsValid(txt(Area), "Area Name") = False Then Exit Sub
    If IsValid(Txt(REP_CODE), "Sales Executive") = False Then Exit Sub
    If IsValid(Txt(Model), "MODEL") = False Then Exit Sub
    'If IsValid(txt(REF_CODE), "Refferd By") = False Then Exit Sub
    If Txt(City).Tag = "" Then
        MsgBox "Please enter City", vbOKOnly, "Validation"
        Txt(City).Enabled = True
        Txt(City) = ""
        Txt(City).SetFocus
        Exit Sub
    End If
    'If IsValid(txt(Profession), "Profession") = False Then Exit Sub
    'If IsValid(txt(Purpose), "Purpuse") = False Then Exit Sub
    'If IsValid(txt(Model), "Model No.") = False Then Exit Sub
    If Txt(ExpDelDt) <> "" Then
        If CDate(Txt(VDate)) > CDate(Txt(ExpDelDt)) Then
            MsgBox "Expected Delivery Date is less than Booking Date", vbOKOnly, "Date Validation"
            Txt(ExpDelDt).SetFocus
            Exit Sub
        End If
    End If
'    If IsValid(txt(FundSource), "Fund Source") = False Then Exit Sub
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, Col_Accessory) <> "" Then
            If Val(FGrid1.TextMatrix(I, Col_Qty)) = 0 Then MsgBox "Please Specify Quantity in Row No. " & I, vbInformation, "Validation": FGrid1.Row = I: FGrid1.Col = Col_Qty: FGrid1.SetFocus: Exit Sub
        End If
    Next
    
    AccCode = ""
    AccQty = ""
    For I = 1 To FGrid1.Rows - 1
        If FGrid1.TextMatrix(I, Col_Accessory) <> "" Then
            If AccCode = "" Then
                AccCode = AccCode + FGrid1.TextMatrix(I, Col_AccCode)
                AccQty = AccQty + FGrid1.TextMatrix(I, Col_Qty)
            Else
                AccCode = AccCode + "," + FGrid1.TextMatrix(I, Col_AccCode)
                AccQty = AccQty + "," + FGrid1.TextMatrix(I, Col_Qty)
            End If
        End If
    Next
    
    Grid_Hide
    ac_str = ""
    For I = 0 To FGrid.Rows - 1
        If FGrid.TextMatrix(I, 1) = "ü" Then
            ac_str = ac_str + IIf(ac_str = "", FGrid.TextMatrix(I, 2), "," + FGrid.TextMatrix(I, 2))
        End If
    Next
    
    
    
    
    Select Case Txt(FundSource).TEXT
        Case "Hypothecation"
            mFundSource = 0
        Case "Hire Purchase"
            mFundSource = 1
        Case "Own Fund"
            mFundSource = 2
        Case "Lease"
            mFundSource = 3
        Case "Agreement"
            mFundSource = 4
        Case "Lease & Agreement"
            mFundSource = 5
         Case "Loan Cum Hypt."
            mFundSource = 6
        Case Else
            mFundSource = 7
    End Select
GCn.BeginTrans
GCnFaV.BeginTrans    'If New Ledger A/c
    mTrans = True
    If TopCtrl1.TopText2.CAPTION = "Add" Then
    If UCase(left(PubComp_Name, 7)) = "SHANKAR" Or UCase(left(PubComp_Name, 6)) = "MAURYA" Then
        If GCn.Execute("select count(*) from subgroup where name&fname  = '" & Txt(Party_code).TEXT & Txt(fname).TEXT & "'").Fields(0) = 0 Then
            If MsgBox("Ledger A/c doesn't exist for this customer!!!" & vbCrLf & "Want to Create new Ledger AC ?", vbYesNo, "New Party  Creation") = vbYes Then
                 If AddSubGroup = False Then GoTo errlbl
            End If
        End If
    Else
        If Txt(Party_code).Tag = "" Then
            If MsgBox("Ledger A/c doesn't exist for this customer!!!" & vbCrLf & "Want to Create new Ledger AC ?", vbYesNo, "New Party  Creation") = vbYes Then
                If GCn.Execute("select count(*) from subgroup where name  = '" & Txt(Party_code).TEXT & "' ").Fields(0) = 0 Then
                    If AddSubGroup = False Then GoTo errlbl
                Else
                    MsgBox "Ledger A/c exists with same Name add some suffix", vbInformation, "Duplicate Ledger Ac"
                    Txt(Party_code).SetFocus
                    GoTo errlbl
                End If
            Else
                MsgBox "First Create Ledger A/c from Ledger Entry!!!", vbInformation, "Party not found"
                GoTo errlbl
            End If
        End If
    End If
        If GCn.Execute("select count(*) from ProspectiveCust where " & cUCase("name+add1+add2+add3+citycode") & " = " & cUCase("'" & Txt(Party_code).TEXT & Txt(Add1).TEXT & Txt(Add2).TEXT & Txt(Add3).TEXT & Txt(City).Tag & "'") & "").Fields(0) = 0 Then
            Call CreateCustomer
        Else
            CustDocId = GCn.Execute("select cust_code from ProspectiveCust where " & cUCase("name+add1+add2+add3+citycode") & " = " & cUCase("'" & Txt(Party_code).TEXT & Txt(Add1).TEXT & Txt(Add2).TEXT & Txt(Add3).TEXT & Txt(City).Tag & "'") & "").Fields(0).Value
        End If
        If Txt(QuotNo).TEXT = "" Then
            If CreateQuotation = False Then GoTo errlbl
        End If
        'determine booking serial no
        DocID = Txt(TxtDocID)
        
        
' Stopped Beacuse Taking Same Srl No of Different Sites
'        If GCn.Execute("select count(*) from veh_order where ordDocId='" & Txt(TxtDocID) & "'").Fields(0) > 0 Then
'            If VoucherEditFlag Then 'And txt(BookNo).Visible Then
'                MsgBox "Order No. already exists, Retry", vbCritical, "Validation Error"
'                Txt(BookNo).SetFocus
'                GoTo errlbl
'            Else
'                Txt(TxtDocID) = GetDocID(GCnFaV, mVtype, Txt(VDate), VoucherEditFlag, Txt(BookNo), LblVPrefix)
'                If Val(Txt(BookNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
'                    MsgBox "Order No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
'                    GoTo errlbl
'                End If
'            End If
'        End If


        If GCn.Execute("select count(*) from veh_order where Ord_No=" & Val(Txt(BookNo)) & " And Ord_VType = '" & mVType & "' And Left(ordDocId,1)='" & PubDivCode & "'").Fields(0) > 0 Then
            If VoucherEditFlag Then 'And txt(BookNo).Visible Then
                MsgBox "Order No. already exists, Retry", vbCritical, "Validation Error"
                Txt(BookNo).SetFocus
                GoTo errlbl
            Else
                Txt(TxtDocID) = GetDocID(GCnFaV, mVType, Txt(VDate), VoucherEditFlag, Txt(BookNo), LblVPrefix)
                If Val(Txt(BookNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
                    SetMax_VoucherPrefix "OrdDocID", mVType, "Veh_Order", "Ord_Date", GCnFaV
                    Txt(TxtDocID) = GetDocID(GCnFaV, mVType, Txt(VDate), VoucherEditFlag, Txt(BookNo), LblVPrefix)
                    If Val(Txt(BookNo)) <= Val(DeCodeDocID(DocID, Document_No)) Then
                        MsgBox "Order No. already exists ! " & vbCrLf & "Contact System Administrator", vbCritical, "Validation Error"
                        GoTo errlbl
                    End If
                End If
            End If
        End If

        DocIdHlp = Replace(Txt(TxtDocID), " ", "")
        'eof booking serial no.
        GCn.Execute ("insert into veh_order(OrdDocId,OrdDocIDHelp,Ord_SiteCode, " & _
            "Ord_VType,Ord_No,Ord_Date,Quot_SiteCode, " & _
            "Quot_DocId,QuotSrl_No,PartyCode,AREA,REF_CODE, " & _
            "REP_CODE,Profession,PURPOSE,MODEL, " & _
            "EXP_DATE,RATE,INTD_USE,GOVT_YN, " & _
            "AddVeri_YN,PermitReq_YN,Permit_N_Z,Fund_Source, " & _
            "FIN_AcCode,FB_CODE,FIN_AMT,Other_Facilities,Colour_Code,AdvEMI,Reg_Fee,Ins_Fee,OtherChrg,Rebate,AccCode,AccQty,Brokrage,Subvention, " & _
            "Book_UName,Book_UEntDt,Book_UAE, Book_AddBy, Book_AddDate,DoNo,DoReciveDate,DOIssueDate) " & _
            " values('" & Txt(TxtDocID) & "', '" & DocIdHlp & "' , '" & PubSiteCode & Txt(Site).Tag & "',  " & _
            "'" & mVType & "'," & Val(Txt(BookNo).TEXT) & ", " & ConvertDate(Txt(VDate).TEXT) & " ,'" & QSiteCode & "', " & _
            "'" & QDocId & "'," & QDocSrlNo & ", '" & Txt(Party_code).Tag & "' ,'" & Txt(Area).Tag & "','" & Txt(REF_CODE).Tag & "', " & _
            "'" & Txt(REP_CODE).Tag & "','" & Txt(Profession).Tag & "', '" & Txt(Purpose).Tag & "', '" & Txt(Model).TEXT & "', " & _
            "" & ConvertDate(Txt(ExpDelDt).TEXT) & ", " & Val(Txt(Rate)) & ", '" & Txt(IndUse) & "'," & IIf(Txt(GovtYn).TEXT = "Yes", 1, 0) & " , " & _
            "" & IIf(Txt(AddVerYN).TEXT = "Yes", 1, 0) & "," & IIf(Txt(ReqPermit).TEXT = "Yes", 1, 0) & "," & IIf(Txt(NZ).TEXT = "National", 0, IIf(Txt(NZ).TEXT = "Zonal", 1, 0)) & _
            ", " & mFundSource & " , " & _
            "'" & FinAcCode & "','" & Txt(FB_Code).Tag & "'," & Val(Txt(Amt).TEXT) & " ,'" & ac_str & "' , " & _
            "'" & Txt(Colours).Tag & "'," & Val(Txt(AdvEMI)) & "," & Val(Txt(RegAmt)) & "," & Val(Txt(InsuAmt)) & _
            "," & Val(Txt(OthAmt)) & "," & Val(Txt(DiscAmt)) & ",'" & AccCode & "','" & AccQty & "'," & Val(Txt(Brok)) & _
            "," & Val(Txt(SubVen)) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A', '" & pubUName & "', " & ConvertDateTime(PubServerDate) & " ,'" & Txt(50).TEXT & "'," & ConvertDate(Txt(51).TEXT) & "," & ConvertDate(Txt(52).TEXT) & ")")
        'Voucher Serial No. Updation LPS 21-05-03
        'update Table only when DocSrlNo >Table.SerialNo
        UpdVouSrlNo GCnFaV, Txt(TxtDocID), Txt(VDate)
        If Txt(QuotNo).TEXT <> "" Then GCn.Execute "update veh_quot1 set Book_SiteCode='" & Txt(Site).Tag & "',Book_docid='" & Txt(TxtDocID) & "' where docid = '" & QDocId & "'and srl_no = " & QDocSrlNo & ""
    Else    'Edit
        'Update not allowed
        'Quot_SiteCode='" & QSiteCode & "', Quot_DocId='" & QDocId & "',
        GCn.Execute "update veh_order set Ord_Date=" & ConvertDate(Txt(VDate)) & ", PartyCode='" & Txt(Party_code).Tag & "',AREA='" & Txt(Area).Tag & "',REF_CODE='" & Txt(REF_CODE).Tag & "', " & _
            "REP_CODE='" & Txt(REP_CODE).Tag & "',Profession='" & Txt(Profession).Tag & "',PURPOSE='" & Txt(Purpose).Tag & "',MODEL='" & Txt(Model).TEXT & "', " & _
            "EXP_DATE=" & ConvertDate(Txt(ExpDelDt).TEXT) & ",RATE=" & Val(Txt(Rate)) & ",INTD_USE='" & Txt(IndUse) & "',GOVT_YN=" & IIf(Txt(GovtYn).TEXT = "Yes", 1, 0) & ", " & _
            "AddVeri_YN=" & IIf(Txt(AddVerYN).TEXT = "Yes", 1, 0) & ",PermitReq_YN=" & IIf(Txt(ReqPermit).TEXT = "Yes", 1, 0) & ",Permit_N_Z=" & IIf(Txt(NZ).TEXT = "National", 0, IIf(Txt(NZ).TEXT = "Zonal", 1, 0)) & ",Fund_Source=" & mFundSource & ", " & _
            "FIN_AcCode='" & FinAcCode & "',FB_CODE='" & Txt(FB_Code).Tag & "',FIN_AMT=" & Val(Txt(Amt).TEXT) & ",Other_Facilities='" & ac_str & "', " & _
            "Colour_Code='" & Txt(Colours).Tag & "',AdvEMI=" & Val(Txt(AdvEMI)) & ",Reg_Fee=" & Txt(RegAmt) & ",Ins_Fee=" & Txt(InsuAmt) & ",OtherChrg=" & Txt(OthAmt) & ",Rebate=" & Txt(DiscAmt) & ",AccCode='" & AccCode & "',AccQty='" & AccQty & "',Brokrage=" & Val(Txt(Brok)) & ",Subvention=" & Val(Txt(SubVen)) & ",Book_UName='" & pubUName & _
            "',Book_UEntDt=" & ConvertDate(PubServerDate) & ",Book_UAE='E', Book_ModifyBy = '" & pubUName & "', Book_ModifyDate = " & ConvertDateTime(PubServerDate) & ",DoNo='" & Txt(50).TEXT & "',DoReciveDate=" & ConvertDate(Txt(51).TEXT) & ",DOIssueDate=" & ConvertDate(Txt(52).TEXT) & "  " & _
            "where OrdDocId = '" & Txt(TxtDocID) & "'"
        If GCn.Execute("Select " & xIsNull("CityCode", "") & " as CityCd from subGroup where SubCode='" & Txt(Party_code).Tag & "'").Fields(0).Value = "" Then
            GSQL = "Update SubGroup set CityCode='" & Txt(City).Tag & "' where SubCode='" & Txt(Party_code).Tag & "'"
            TSQL = "Update SubGroupAlias set CityCode='" & Txt(City).Tag & "' where SubCode='" & Txt(Party_code).Tag & "'"
            GCn.Execute GSQL
            GCn.Execute TSQL
            GCnFaV.Execute GSQL
            GCnFaV.Execute TSQL
        End If
    End If

    GCnFaV.CommitTrans
    GCn.CommitTrans
    mTrans = False
    If PubMoveRecYn Then
        Master.Requery
    Else
        Set Master = GCn.Execute("Select OrdDocId as searchcode, OrdDocId from Veh_Order where Ord_VType = 'V_BK' and left(OrdDocId,1)='" & PubDivCode & "' and (Ord_Date>=" & ConvertDate(PubStartDate) & " Or Inv_Date is Null)  And OrdDocId = '" & Txt(TxtDocID) & "'  order by Right(OrdDocId,8), Ord_Date desc")
    End If
    RsQuot.Requery
    RsParty.Requery
    Master.FIND "SearchCode = '" & Txt(TxtDocID) & "'"
    'lp 11-03-03
    If TopCtrl1.TopText2.CAPTION = "Add" Then
        If Val(Txt(BookNo)) > DeCodeDocID(DocID, Document_No) Then
            MsgBox "Order No." & Trim(DeCodeDocID(DocID, Document_No)) & " already exists ! " & vbCrLf & "New No. " & Txt(BookNo) & " alloted", vbCritical, "Document No. Changed"
        End If
    End If
    TopCtrl1_ePrn
    Exit Sub
errlbl:
    If mTrans Then GCn.RollbackTrans: GCnFaV.RollbackTrans
    CheckError
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Ctrl_GetFocus Txt(Index)
Grid_Hide
Select Case Index
    Case NamePrefix
        ListArray = Array("    ", "Mr.", "Mrs.", "Miss", "Ms", "M/S", "Dr.")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 7)
    Case FNamePrefix
        ListArray = Array("S/O", "W/O", "D/O", "C/O", "And ", "U/C")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 6)
    Case Site
        If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
            RsSite.Filter = adFilterNone
            RsSite.Filter = " Code = '" & PubSiteCode & "' "
        Else
            RsSite.Filter = adFilterNone
        End If
    
        Set DGSite.DataSource = RsSite
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Then Exit Sub
        If Txt(Index).TEXT = "" Then
            RsSite.MoveFirst
            RsSite.FIND "code ='" & PubSiteCode & "'"
            Txt(Index).Tag = RsSite!Code
            Txt(Index).TEXT = RsSite!Name
        Else
            If Txt(Index).TEXT <> RsSite!Name Then
                RsSite.MoveFirst
                RsSite.FIND "name ='" & Txt(Index).TEXT & "'"
            End If
        End If
    Case FundSource
        ListArray = Array("Hypothecation", "Hire Purchase", "Own Fund", "Lease", "Agreement", "Lease & Agreement", "Loan Cum Hypt.")
        Set mListItem = ListView_Items(ListView, Txt, Index, ListArray, 7)
    Case QuotNo
        If RsQuot.RecordCount = 0 Or (RsQuot.EOF = True Or RsQuot.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).Tag <> RsQuot!Code Then
            RsQuot.MoveFirst
            RsQuot.FIND "Code ='" & Txt(QuotNo).Tag & "'"
        End If
    Case Purpose
        If RsPurpose.RecordCount = 0 Or (RsPurpose.EOF = True Or RsPurpose.BOF = True) Or Txt(Purpose).TEXT = "" Then Exit Sub
        If Txt(Purpose).TEXT <> RsPurpose!Name Then
            RsPurpose.MoveFirst
            RsPurpose.FIND "name ='" & Txt(Purpose).TEXT & "'"
        End If
    Case City
        If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or Txt(City).TEXT = "" Then Exit Sub
        If Txt(City).TEXT <> RsCity!Name Then
            RsCity.MoveFirst
            RsCity.FIND "name ='" & Txt(City).TEXT & "'"
        End If
    Case FB_Code
        If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or Txt(FB_Code).TEXT = "" Then Exit Sub
        If Txt(FB_Code).TEXT <> rsFin!Name Then
            rsFin.MoveFirst
            rsFin.FIND "Code ='" & Txt(FB_Code).Tag & "'"
        End If
    Case Party_code
        Set DGParty.DataSource = RsParty
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then Exit Sub
        If Txt(Index).TEXT <> RsParty!Name Then
            RsParty.MoveFirst
            RsParty.FIND "name ='" & Txt(Index).TEXT & "'"
        End If
    Case Area
        If RsArea.RecordCount = 0 Or (RsArea.EOF = True Or RsArea.BOF = True) Or Txt(Area).TEXT = "" Then Exit Sub
        If Txt(Area).TEXT <> RsArea!Name Then
            RsArea.MoveFirst
            RsArea.FIND "name ='" & Txt(Area).TEXT & "'"
        End If

    Case REF_CODE
        If RsRef.RecordCount = 0 Or (RsRef.EOF = True Or RsRef.BOF = True) Or Txt(REF_CODE).TEXT = "" Then Exit Sub
        If Txt(REF_CODE).TEXT <> RsRef!Name Then
            RsRef.MoveFirst
            RsRef.FIND "name ='" & Txt(REF_CODE).TEXT & "'"
        End If
    Case REP_CODE
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or Txt(REP_CODE).TEXT = "" Then Exit Sub
        If Txt(REP_CODE).TEXT <> RsRep!Name Then
            RsRep.MoveFirst
            RsRep.FIND "name ='" & Txt(REP_CODE).TEXT & "'"
        End If
    Case Profession
        If RsProf.RecordCount = 0 Or (RsProf.EOF = True Or RsProf.BOF = True) Or Txt(Profession).TEXT = "" Then Exit Sub
        If Txt(Profession).TEXT <> RsProf!Name Then
            RsProf.MoveFirst
            RsProf.FIND "name ='" & Txt(Profession).TEXT & "'"
        End If
    Case Colours
        If RsCol.RecordCount = 0 Or (RsCol.EOF = True Or RsCol.BOF = True) Or Txt(Colours).TEXT = "" Then Exit Sub
        If Txt(Colours).TEXT <> RsCol!Name Then
            RsCol.MoveFirst
            RsCol.FIND "name ='" & Txt(Colours).TEXT & "'"
        End If
    Case Model
        RsMod.Filter = adFilterNone
        If Txt(ModelGroup) <> "" Then
            RsMod.Filter = "Grp_Code = '" & Txt(ModelGroup).Tag & "'"
        End If
        Set DGMod.DataSource = RsMod
        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or Txt(Model).TEXT = "" Then Exit Sub
        If Txt(Model).TEXT <> RsMod!Code Then
            RsMod.MoveFirst
            RsMod.FIND "code ='" & Txt(Model).TEXT & "'"
        End If
    Case ModelGroup
        DgModelGroup.Move Txt(Index).left, Txt(Index).top + Txt(Index).height + 30
        If RsModelGroup.RecordCount = 0 Or (RsModelGroup.EOF = True Or RsModelGroup.BOF = True) Or Txt(ModelGroup).TEXT = "" Then Exit Sub
        If Txt(ModelGroup).Tag <> RsModelGroup!Code Then
            RsModelGroup.MoveFirst
            RsModelGroup.FIND "code ='" & Txt(ModelGroup).Tag & "'"
        End If
    
    Case Rate, Amt
        SendKeys "{HOME}+{END}"
End Select
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim I As Byte
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
'38 =vbKeyUp : 40 = vbKeyDown
If KeyCode = vbKeyEscape Then
    Grid_Hide
    Exit Sub
End If
Select Case Index
    Case NamePrefix
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1800
    Case FNamePrefix
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width, 1800
    Case City
        DGridTxtKeyDown DGCity, Txt, Index, RsCity, KeyCode, False, 1, frmCity, "frmCity"
    Case Site
        DGridTxtKeyDown DGSite, Txt, Index, RsSite, KeyCode, False, 1
    Case FundSource
        ListView_KeyDown FrmList, ListView, Txt, Index, KeyCode, Shift, Txt(Index).left, (Txt(Index).top + Txt(Index).height), Txt(Index).width + 100, 1800
    Case QuotNo
        DGridTxtKeyDown_Mast DGQuot, Txt, QuotNo, RsQuot, KeyCode, False, 1
    Case Party_code
       DGridTxtKeyDown_Mast DGParty, Txt, Party_code, RsParty, KeyCode, False, 1
       'lblGroup.Visible = True: lblGroup.Locked = True: lblGroup.ZOrder 0: lblGroup.left = DGParty.left: lblGroup.top = DGParty.top - lblGroup.height: lblGroup.width = DGParty.width
    Case FB_Code
        DGridTxtKeyDown DGFin, Txt, Index, rsFin, KeyCode, False, 1, frmFinMast, "frmFinMast"
    Case Model
        DGridTxtKeyDown DGMod, Txt, Index, RsMod, KeyCode, False, 0, frmModel, "frmModel"
    Case ModelGroup
        DGridTxtKeyDown DgModelGroup, Txt, Index, RsModelGroup, KeyCode, False, 1, frmModelGrp, "frmModelGrp"
    Case Colours
        DGridTxtKeyDown DGCol, Txt, Index, RsCol, KeyCode, False, 1, frmColor, "frmColor"
    Case Area
        DGridTxtKeyDown DGArea, Txt, Index, RsArea, KeyCode, False, 1, frmArea, "frmArea"
    Case REF_CODE
        DGridTxtKeyDown DGRef, Txt, Index, RsRef, KeyCode, False, 1
    Case REP_CODE
        DGridTxtKeyDown DGRep, Txt, Index, RsRep, KeyCode, False, 1, frmEmpMast, "frmEmpMast"
    Case Profession
        DGridTxtKeyDown DGProf, Txt, Index, RsProf, KeyCode, False, 1
    Case Purpose
      DGridTxtKeyDown DGPurpose, Txt, Index, RsPurpose, KeyCode, False, 1
End Select
If FrmList.Visible = False And DGCity.Visible = False And DgModelGroup.Visible = False And DGQuot.Visible = False And DGPurpose.Visible = False And DGSite.Visible = False And DGFin.Visible = False And DGCol.Visible = False And DGMod.Visible = False And DGRep.Visible = False And DGRef.Visible = False And DGProf.Visible = False And DGArea.Visible = False And DGParty.Visible = False Then
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyUp) And Index = Party_code Then
            Txt_Validate Index, True
        End If
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) And Index = VDate Then Txt_Validate Index, True
        If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
        If TopCtrl1.TopText2.CAPTION = "Add" And Index <> Site Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        ElseIf TopCtrl1.TopText2.CAPTION = "Edit" And Index <> Party_code Then
            If KeyCode = vbKeyUp Then Ctrl_UpKeyDown KeyCode, Shift
        End If
End If
End Sub

Private Sub TXT_KeyPress(Index As Integer, KeyAscii As Integer)
Call CheckQuote(KeyAscii)
Select Case Index
    Case City
        If DGCity.Visible = True Then DGridTxtKeyPress Txt, Index, RsCity, KeyAscii, "Name"
    Case Site
        If DGSite.Visible = True Then DGridTxtKeyPress Txt, Index, RsSite, KeyAscii, "Name"
    Case QuotNo
        If DGQuot.Visible = True Then DGridTxtKeyPress Txt, Index, RsQuot, KeyAscii, "Name"
    Case Purpose
        If DGPurpose.Visible = True Then DGridTxtKeyPress Txt, Index, RsPurpose, KeyAscii, "Name"
    Case Model
        If DGMod.Visible = True Then DGridTxtKeyPress Txt, Index, RsMod, KeyAscii, "code"
    Case ModelGroup
        If DgModelGroup.Visible = True Then DGridTxtKeyPress Txt, Index, RsModelGroup, KeyAscii, "Name"
    'Case Party_code
    '    If DGParty.Visible = True Then DGridTxtKeyPress Txt, Index, RsParty, KeyAscii, "Name"
    Case FB_Code
        If DGFin.Visible = True Then DGridTxtKeyPress Txt, Index, rsFin, KeyAscii, "Name"
    Case Colours
        If DGCol.Visible = True Then DGridTxtKeyPress Txt, Index, RsCol, KeyAscii, "Name"
    Case Area
        If DGArea.Visible = True Then DGridTxtKeyPress Txt, Index, RsArea, KeyAscii, "Name"
    Case REP_CODE
        If DGRep.Visible = True Then DGridTxtKeyPress Txt, Index, RsRep, KeyAscii, "Name"
    Case REF_CODE
        If DGRef.Visible = True Then DGridTxtKeyPress Txt, Index, RsRef, KeyAscii, "Name"
    Case Profession
        If DGProf.Visible = True Then DGridTxtKeyPress Txt, Index, RsProf, KeyAscii, "Name"
    Case GovtYn, FirstVeh, AddVerYN, ReqPermit
        If UCase(Chr(KeyAscii)) = "N" Then
            Txt(Index) = "No"
        ElseIf UCase(Chr(KeyAscii)) = "Y" Then
            Txt(Index) = "Yes"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            Txt(Index) = ""
        End If
        KeyAscii = 0
    Case NZ
        If UCase(Chr(KeyAscii)) = "N" Then
            Txt(Index) = "National"
        ElseIf UCase(Chr(KeyAscii)) = "Z" Then
            Txt(Index) = "Zonal"
        ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            Txt(Index) = ""
        End If
        KeyAscii = 0
    Case QuotNo
        Call NumPress(Txt(Index), KeyAscii, 8, 0)
    Case BookNo
        Call NumPress(Txt(Index), KeyAscii, 8, 0)
    Case Rate, AdvEMI, InsuAmt, DiscAmt, OthAmt, RegAmt, Brok, SubVen
        Call NumPress(Txt(Index), KeyAscii, 7, 2)
    Case Amt
        Call NumPress(Txt(Index), KeyAscii, 7, 2)
End Select
End Sub

Private Sub Txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Select Case Index
    Case Party_code
        If DGParty.Visible = True Then DGridTxtKeyUp_Mast Txt, Index, RsParty, KeyCode, "Name"
    Case NamePrefix, FNamePrefix, FundSource
        If FrmList.Visible = True Then ListView_KeyUp ListView, Txt, Index, KeyCode, mListItem
End Select
End Sub

Private Sub Txt_LostFocus(Index As Integer)
  Ctrl_validate Txt(Index)
End Sub

Private Sub Txt_Validate(Index As Integer, Cancel As Boolean)
Dim Rst As ADODB.Recordset
Dim TempRs As ADODB.Recordset
Select Case Index
    Case City
        If RsCity.RecordCount = 0 Or (RsCity.EOF = True Or RsCity.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsCity!Name
            Txt(Index).Tag = RsCity!Code
        End If
    Case Site
         If IsValid(Txt(Index), "Site Code") = False Then Cancel = True: Exit Sub
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsSite!Name
            Txt(Index).Tag = RsSite!Code
        End If
    Case QuotNo
        If RsQuot.RecordCount = 0 Or (RsQuot.EOF = True Or RsQuot.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).Tag = RsQuot!Code
            Txt(Index).TEXT = RsQuot!Name
        End If
        Call Fill_Data
       
    Case NamePrefix, FundSource, FNamePrefix
        If Txt(Index).TEXT <> "" Then Txt(Index).TEXT = ListView.SelectedItem.TEXT
    Case Party_code
        If RsParty.RecordCount = 0 Or (RsParty.EOF = True Or RsParty.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Add1).Enabled = True
            Txt(Add2).Enabled = True
            Txt(Add3).Enabled = True
            Txt(Phone).Enabled = True
            Txt(City).Enabled = True
            Txt(PinCode).Enabled = True
'            Txt(City).Tag = ""
'            Txt(Index).Tag = ""
        Else
            If Txt(Party_code) = RsParty!Name Then
                FillPartyDetails
            Else
                Txt(Add1).Enabled = True
                Txt(Add2).Enabled = True
                Txt(Add3).Enabled = True
                Txt(Phone).Enabled = True
                Txt(City).Enabled = True
                Txt(PinCode).Enabled = True
'                Txt(City).Tag = ""
'                Txt(Index).Tag = ""
                Txt(PAN).Enabled = True
                Txt(fname).Enabled = True
                Txt(FNamePrefix).Enabled = True
            End If
        End If
    Case ModelGroup
        If RsModelGroup.RecordCount = 0 Or (RsModelGroup.EOF = True Or RsModelGroup.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsModelGroup!Name
            Txt(Index).Tag = RsModelGroup!Code
        End If
        
    Case Model
        If RsMod.RecordCount = 0 Or (RsMod.EOF = True Or RsMod.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
            Txt(ModelDesc).Tag = ""
        Else
            Txt(Index).TEXT = RsMod!Code
            Txt(Index).Tag = RsMod!Code
            Txt(ModelDesc) = RsMod!Name
        End If
        If Txt(VDate) <> "" Then
            'txt(Rate) = Format(VehSRate(txt(VDate), txt(Model), "No", "Y"), "0.00")
            Txt(Rate) = Format(VNull(RsMod!Sale_Rate), "0.00")
        End If
        If Txt(Model) <> "" Then
            Set TempRs = GCn.Execute("SELECT " & vIsNull("Sale_Rate", "0") & " as Sale_Rate from Model where model='" & Txt(Model) & "' and div_code='" & PubDivCode & "'")
            If TempRs.RecordCount > 0 Then Txt(Rate) = Format(TempRs!Sale_Rate, "0.00")
        End If
    Case Purpose
        If RsPurpose.RecordCount = 0 Or (RsPurpose.EOF = True Or RsPurpose.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsPurpose!Name
            Txt(Index).Tag = RsPurpose!Code
        End If

    Case FB_Code
        If rsFin.RecordCount = 0 Or (rsFin.EOF = True Or rsFin.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
            FinAcCode = ""
        Else
            Txt(Index).TEXT = rsFin!Name
            Txt(Index).Tag = rsFin!Code
            FinAcCode = rsFin!AcCode
        End If
    Case Colours
        If RsCol.RecordCount = 0 Or (RsCol.EOF = True Or RsCol.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsCol!Name
            Txt(Index).Tag = RsCol!Code
        End If
    Case Area
        If RsArea.RecordCount = 0 Or (RsArea.EOF = True Or RsArea.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsArea!Name
            Txt(Index).Tag = RsArea!Code
        End If
    Case REP_CODE
        If RsRep.RecordCount = 0 Or (RsRep.EOF = True Or RsRep.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsRep!Name
            Txt(Index).Tag = RsRep!Code
        End If
    Case Profession
        If RsProf.RecordCount = 0 Or (RsProf.EOF = True Or RsProf.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsProf!Name
            Txt(Index).Tag = RsProf!Code
        End If
    
    Case REF_CODE
        If RsRef.RecordCount = 0 Or (RsRef.EOF = True Or RsRef.BOF = True) Or Txt(Index).TEXT = "" Then
            Txt(Index).TEXT = ""
            Txt(Index).Tag = ""
        Else
            Txt(Index).TEXT = RsRef!Name
            Txt(Index).Tag = RsRef!Code
        End If
    Case ExpDelDt
        Txt(Index).TEXT = RetDate(Txt(Index))
    
    Case VDate
        If Len(Trim(Txt(VDate).TEXT)) = 0 Then
             Txt(VDate).TEXT = PubLoginDate
        Else
            Txt(Index).TEXT = RetDate(Txt(Index))
        End If
        If TopCtrl1.TopText2 = "Add" Then
            If CheckFinYear(Txt(Index)) Then
                Txt(TxtDocID) = GetDocID(GCnFaV, mVType, Txt(VDate), VoucherEditFlag, Txt(BookNo), LblVPrefix)
                DocID = Txt(TxtDocID)
            Else
                Cancel = True
            End If
        End If
    Case BookNo
        If IsValid(Txt(BookNo), "Booking No.") = False Then Cancel = True:   Exit Sub
        If VoucherEditFlag = True Then      ' Manual
            Txt(TxtDocID) = GetDocID(GCnFaV, mVType, Txt(VDate), VoucherEditFlag, Txt(BookNo), LblVPrefix)
            Set Rst = New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "Select * From veh_order Where OrdDocId='" & Txt(TxtDocID) & "'", GCn, adOpenDynamic, adLockOptimistic
            If Rst.RecordCount > 0 Then
                MsgBox "Duplicate Booking No. Not Allowed", vbInformation, "Validation"
                Cancel = True
                Txt(BookNo).SetFocus
            End If
        End If
        Case DOReciveDate, DOIssueDate
            Txt(Index) = RetDate(Txt(Index))
      End Select

      
Set Rst = Nothing
End Sub
Private Sub DGMod_Click()
    If RsMod.RecordCount > 0 Then
        Txt(Model).TEXT = RsMod!Code
        Txt(Model).Tag = RsMod!Code
    End If
    Txt(Model).SetFocus
    DGMod.Visible = False
End Sub
Private Sub DGParty_Click()
    If RsParty.RecordCount > 0 Then
        FillPartyDetails
    End If
    DGParty.Visible = False
    lblGroup.Visible = False
    Txt(Party_code).SetFocus
End Sub
Private Sub DGCol_Click()
    DGCol.Visible = False
    If RsCol.RecordCount > 0 Then
        Txt(Colours).TEXT = RsCol!Name
        Txt(Colours).Tag = RsCol!Code
    End If
    Txt(Colours).SetFocus
End Sub
Private Sub DGFin_Click()
DGFin.Visible = False
If rsFin.RecordCount > 0 Then
            Txt(FB_Code).TEXT = rsFin!Name
            Txt(FB_Code).Tag = rsFin!Code
            FinAcCode = rsFin!AcCode
End If
   Txt(FB_Code).SetFocus
End Sub
Private Sub DGArea_Click()
    DGArea.Visible = False
    If RsArea.RecordCount > 0 Then
        Txt(Area).TEXT = RsArea!Name
        Txt(Area).Tag = RsArea!Code
    End If
    Txt(Area).SetFocus
End Sub
Private Sub DGCity_Click()
    DGCity.Visible = False
    If RsCity.RecordCount > 0 Then
        Txt(City).TEXT = RsCity!Name
        Txt(City).Tag = RsCity!Code
    End If
    Txt(City).SetFocus
End Sub

Private Sub DGQuot_Click()
    DGQuot.Visible = False
    If RsQuot.RecordCount > 0 Then
        Txt(QuotNo).TEXT = RsQuot!Name
        Txt(QuotNo).Tag = RsQuot!Code
    End If
    Fill_Data
    Txt(QuotNo).SetFocus
End Sub

Private Sub DGProf_Click()
    DGProf.Visible = False
    If RsProf.RecordCount > 0 Then
        Txt(Profession).TEXT = RsProf!Name
        Txt(Profession).Tag = RsProf!Code
    End If
    Txt(Profession).SetFocus
    

End Sub

Private Sub DGRef_Click()
    DGRef.Visible = False
    If RsRef.RecordCount > 0 Then
        Txt(REF_CODE).TEXT = RsRef!Name
        Txt(REF_CODE).Tag = RsRef!Code
    End If
    Txt(REF_CODE).SetFocus

End Sub

Private Sub DGRep_Click()
    DGRep.Visible = False
        If RsRep.RecordCount > 0 Then
            Txt(REP_CODE).TEXT = RsRep!Name
            Txt(REP_CODE).Tag = RsRep!Code
        End If
        Txt(REP_CODE).SetFocus
End Sub



'******* Fuctions **********
Private Sub BlankText()
Dim I As Byte
For I = 0 To Txt.Count - 1
    Txt(I).TEXT = ""
    Txt(I).Tag = ""
Next I
End Sub

Private Sub MoveRec()
Dim Rst As ADODB.Recordset, Master1 As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim I As Integer
Dim j As Integer
Dim K As Integer
Dim OthFac As String
Dim ConOthFac As String
Dim ProdCode, ProdQty As String
On Error GoTo error1
If Master.RecordCount > 0 Then
    Set Master1 = New ADODB.Recordset
    Master1.CursorLocation = adUseClient
    Master1.Open "Select OrdDocId as searchcode, " & _
        "OrdDocId,OrdDocIDHelp,Ord_SiteCode,Ord_VType,Ord_No,Ord_Date,Quot_SiteCode, " & _
        "Quot_DocId,QuotSrl_No,PartyCode,AREA,REF_CODE,REP_CODE,Profession,PURPOSE,MODEL, " & _
        "EXP_DATE,RATE,FirstVeh_YN,INTD_USE,GOVT_YN,AddVeri_YN,PermitReq_YN,Permit_N_Z,Fund_Source, " & _
        "FIN_AcCode,FB_CODE,FIN_AMT,Other_Facilities,Colour_Code,Book_UName,Book_UEntDt,Book_UAE, " & _
        "Inv_DocId, Inv_Date,DelCh_DocId, DelCh_DT,AdvEMI,Ins_Fee,Reg_Fee,Rebate,OtherChrg,AccCode,AccQty,Brokrage,Subvention, DoNo,DoReciveDate,DOIssueDate " & _
        "from Veh_Order  where OrdDocId = '" & Master!SearchCode & "'", GCn, adOpenDynamic, adLockOptimistic

    Txt(InvNo) = DeCodeDocID(XNull(Master1!Inv_DocId), Document_No)
    Txt(InvDate) = IIf(IsNull(Master1!Inv_Date), "", Master1!Inv_Date)
    Txt(DelChNo) = DeCodeDocID(XNull(Master1!DelCh_DocId), Document_No)
    Txt(DelChDate) = IIf(IsNull(Master1!DelCh_DT), "", Master1!DelCh_DT)

    DocID = Master1!OrdDocId
    Txt(TxtDocID) = Master1!OrdDocId
    mVType = Master1!Ord_VType
    LblVPrefix.CAPTION = DeCodeDocID(Master1!OrdDocId, Document_Prefix)
    Txt(VDate).TEXT = Master1!Ord_Date
    Txt(BookNo).TEXT = Master1!Ord_No
    Txt(Site).Tag = mID(Master1!Ord_SiteCode, 2, 1)
    Txt(Site).TEXT = GCn.Execute("select site_desc from site where site_code = '" & Txt(Site).Tag & "'").Fields(0).Value
    QDocId = IIf(IsNull(Master1!Quot_DocId), "", Master1!Quot_DocId)
    Txt(QuotPrefix) = XNull(mID(Master1!Quot_DocId, 9, 13))
    QDocSrlNo = IIf(IsNull(Master1!QuotSrl_No), 0, Master1!QuotSrl_No)
    QSiteCode = IIf(IsNull(Master1!Quot_SiteCode), "", Master1!Quot_SiteCode)
    Txt(QuotNo).Tag = QDocId
    Txt(QuotNo).TEXT = Trim(mID(QDocId, 14, 8))
    Txt(Party_code).Tag = IIf(IsNull(Master1!PartyCode), "", Master1!PartyCode)
    If Txt(Party_code).Tag <> "" Then
        Set Rst = New Recordset
        Rst.CursorLocation = adUseClient
        Rst.Open "select NamePrefix,name,add1,add2,add3,Phone,CityCode,PANNO,FPrefix,FName, Pin from SubGroup where SubCode = '" & Txt(Party_code).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
        Txt(NamePrefix) = IIf(IsNull(Rst!NamePrefix), "", Rst!NamePrefix)
        Txt(Party_code).TEXT = XNull(Rst!Name)
        Txt(Add1).TEXT = IIf(IsNull(Rst!Add1), "", Rst!Add1)
        Txt(Add2).TEXT = IIf(IsNull(Rst!Add2), "", Rst!Add2)
        Txt(Add3).TEXT = IIf(IsNull(Rst!Add3), "", Rst!Add3)
        Txt(Phone).TEXT = IIf(IsNull(Rst!Phone), "", Rst!Phone)
        Txt(City).Tag = IIf(IsNull(Rst!CityCode), "", Rst!CityCode)
        If Txt(City).Tag <> "" Then
            RsCity.MoveFirst
            RsCity.FIND "Code='" & Txt(City).Tag & "'"
            If RsCity.EOF = False Then
                Txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").Fields(0).Value
            Else
                Txt(City).TEXT = ""
            End If
        Else
            Txt(City).TEXT = ""
        End If
        Txt(PAN) = IIf(IsNull(Rst!PanNo), "", Rst!PanNo)
        Txt(PinCode) = XNull(Rst!Pin)
        Txt(FNamePrefix) = IIf(IsNull(Rst!FPrefix), "", Rst!FPrefix)
        Txt(fname) = IIf(IsNull(Rst!fname), "", Rst!fname)
    Else
        Txt(Party_code).TEXT = ""
        Txt(Add1).TEXT = ""
        Txt(Add2).TEXT = ""
        Txt(Add3).TEXT = ""
        Txt(Phone).TEXT = ""
        Txt(City).TEXT = ""
        Txt(PAN) = ""
        Txt(FNamePrefix) = ""
        Txt(fname) = ""
    End If

        If Not IsNull(Master1!Fund_Source) Then
            Select Case Master1!Fund_Source
                Case 0
                    Txt(FundSource).TEXT = "Hypothecation"
                Case 1
                    Txt(FundSource).TEXT = "Hire Purchase"
                Case 2
                    Txt(FundSource).TEXT = "Own Fund"
                Case 3
                    Txt(FundSource).TEXT = "Lease"
                Case 4
                    Txt(FundSource).TEXT = "Agreement"
                Case 5
                    Txt(FundSource).TEXT = "Lease & Agreement"
                Case 6
                    Txt(FundSource).TEXT = "Loan Cum Hypt."
            End Select
        Else
            Txt(FundSource).TEXT = ""
        End If
        Txt(Profession).Tag = IIf(IsNull(Master1!Profession), "", Master1!Profession)
        If Txt(Profession).Tag <> "" And GCn.Execute("select Professionname from Profession where Professioncode = '" & Txt(Profession).Tag & "'").RecordCount > 0 Then
            Txt(Profession).TEXT = GCn.Execute("select Professionname from Profession where Professioncode = '" & Txt(Profession).Tag & "'").Fields(0).Value
        Else
            Txt(Profession).TEXT = ""
        End If
        Txt(Purpose).Tag = IIf(IsNull(Master1!Purpose), "", Master1!Purpose)
        If Txt(Purpose).Tag <> "" Then
            Set RsTemp = GCn.Execute("select PurposeName from Purpose where Purposecode = '" & Txt(Purpose).Tag & "'")
            If RsTemp.RecordCount > 0 Then
                Txt(Purpose).TEXT = XNull(RsTemp(0))
            Else
                Txt(Purpose).TEXT = ""
            End If
        Else
            Txt(Purpose).TEXT = ""
        End If
        Txt(DONo) = XNull(Master1!DONo)
        Txt(51) = XNull(Master1!DOReciveDate)
        Txt(52) = XNull(Master1!DOIssueDate)
        Txt(IndUse) = IIf(IsNull(Master1!INTD_USE), "", Master1!INTD_USE)
        Txt(Model).Tag = IIf(IsNull(Master1!Model), "", Master1!Model)
        Txt(Model).TEXT = IIf(IsNull(Master1!Model), "", Master1!Model)
        RsMod.MoveFirst
        RsMod.FIND "Code = '" & Txt(Model).Tag & "'"
        If RsMod.EOF = False Then
            Txt(ModelGroup) = XNull(RsMod!ModelGroup)
            Txt(ModelGroup).Tag = XNull(RsMod!Grp_Code)
            Txt(ModelDesc) = XNull(RsMod!Name)
        Else
            Txt(ModelDesc) = ""
            Txt(ModelGroup) = ""
            Txt(ModelGroup).Tag = ""
        End If
        Txt(Area).Tag = IIf(IsNull(Master1!Area), "", Master1!Area)
        If Txt(Area).Tag <> "" And GCn.Execute("select AREAname from AREA where AREAcode = '" & Txt(Area).Tag & "'").RecordCount > 0 Then
            Txt(Area).TEXT = GCn.Execute("select AREAname from AREA where AREAcode = '" & Txt(Area).Tag & "'").Fields(0).Value
        Else
            Txt(Area).TEXT = ""
        End If
        Txt(Colours).Tag = IIf(IsNull(Master1!Colour_Code), "", Master1!Colour_Code)
        If Txt(Colours).Tag <> "" Then
            Txt(Colours).TEXT = GCn.Execute("select col_desc from Colmast where col_code = '" & Txt(Colours).Tag & "'").Fields(0).Value
        Else
            Txt(Colours).TEXT = ""
        End If
        Txt(REP_CODE).Tag = IIf(IsNull(Master1!REP_CODE), "", Master1!REP_CODE)
        If Txt(REP_CODE).Tag <> "" And GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(REP_CODE).Tag & "'").RecordCount > 0 Then
            Txt(REP_CODE).TEXT = GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(REP_CODE).Tag & "'").Fields(0).Value
        Else
            Txt(REP_CODE).TEXT = ""
        End If
        Txt(REF_CODE).Tag = IIf(IsNull(Master1!REF_CODE), "", Master1!REF_CODE)
        If Txt(REF_CODE).Tag <> "" And GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(REF_CODE).Tag & "'").RecordCount > 0 Then
            Txt(REF_CODE).TEXT = GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(REF_CODE).Tag & "'").Fields(0).Value
        Else
            Txt(REF_CODE).TEXT = ""
        End If
        Txt(FB_Code).Tag = IIf(IsNull(Master1!FB_Code), "", Master1!FB_Code)
        If Txt(FB_Code).Tag <> "" Then
            Set Rst = New Recordset
            Rst.CursorLocation = adUseClient
            Rst.Open "select fincode as code,finname + ',' + " & xIsNull("City.CityName", "") & " as name,AcCode " & _
            " from ContractFinance left join city on left(ContractFinance.City,4)=City.CityCode " & _
            " where fincatg = 0 and fincode = '" & Txt(FB_Code).Tag & "'", GCn, adOpenDynamic, adLockOptimistic
            Txt(FB_Code).TEXT = XNull(Rst!Name)
            FinAcCode = IIf(IsNull(Rst!AcCode), "", Rst!AcCode)
        Else
            Txt(FB_Code).TEXT = ""
            FinAcCode = ""
        End If
        OthFac = IIf(IsNull(Master1!Other_Facilities), "", Master1!Other_Facilities)
        For I = 0 To FGrid.Rows - 1
            FGrid.TextMatrix(I, 1) = ""
        Next
        If OthFac <> "" Then
            For I = 0 To FGrid.Rows - 1
                ConOthFac = ""
                For j = 1 To Len(OthFac)
                   If mID(OthFac, j, 1) = "," Then
                         ConOthFac = mID(OthFac, 1, j - 1)
                         OthFac = mID(OthFac, j + 1, Len(OthFac))
                         Exit For
                   ElseIf j = Len(OthFac) And ConOthFac = "" Then
                         ConOthFac = OthFac
                         OthFac = ""
                   End If
                Next
                    For K = 0 To FGrid.Rows - 1
                        If FGrid.TextMatrix(K, 2) = ConOthFac Then
                            FGrid.Col = 1
                            FGrid.Row = K
                            FGrid.CellFontName = "WINGDINGS"
                            FGrid.CellFontSize = 14
                            FGrid.TextMatrix(K, 1) = "ü"
                            Exit For
                        End If
                    Next
            If OthFac = "" Then Exit For
            Next
        End If
        
        Txt(NZ) = IIf(Master1!Permit_N_Z = 0, "National", IIf(Master1!Permit_N_Z = 1, "Zonal", ""))
        Txt(ReqPermit) = IIf(Master1!PermitReq_YN = 1, "Yes", "No")
        Txt(FirstVeh) = IIf(Master1!FirstVeh_YN = 1, "Yes", "No")
        Txt(GovtYn) = IIf(Master1!Govt_YN = 1, "Yes", "No")
        Txt(AddVerYN) = IIf(Master1!AddVeri_YN = 1, "Yes", "No")
        Txt(Rate) = Format(Master1!Rate, "0.00")
        Txt(Amt) = Format(IIf(IsNull(Master1!Fin_Amt), 0, Master1!Fin_Amt), "0.00")
        Txt(ExpDelDt) = IIf(IsNull(Master1!EXP_DATE), "", Master1!EXP_DATE)
        Txt(AdvEMI) = Format(VNull(Master1!AdvEMI), "0.00")
        Txt(InsuAmt) = Format(VNull(Master1!INS_FEE), "0.00")
        Txt(RegAmt) = Format(VNull(Master1!REG_FEE), "0.00")
        Txt(OthAmt) = Format(VNull(Master1!OtherChrg), "0.00")
        Txt(DiscAmt) = Format(VNull(Master1!Rebate), "0.00")
        Txt(Brok) = Format(VNull(Master1!Brokrage), "0.00")
        Txt(SubVen) = Format(VNull(Master1!Subvention), "0.00")
        
        FGrid1.Rows = 1
        I = 1
        For j = 1 To Len(XNull(Master1!AccCode))
            If mID(Master1!AccCode, j, 1) <> "," Then
                ProdCode = ProdCode & mID(Master1!AccCode, j, 1)
            Else
                RsAcces.Sort = "name"
                RsAcces.MoveFirst
                RsAcces.FIND "Code ='" & ProdCode & "'"
                If RsAcces.EOF = True Then RsAcces.MoveFirst
                FGrid1.AddItem ""
                FGrid1.TextMatrix(I, Col_Accessory) = RsAcces!Name
                FGrid1.TextMatrix(I, Col_AccCode) = RsAcces!Code
                I = I + 1
                ProdCode = ""
            End If
        Next
        If ProdCode <> "" Then
            RsAcces.Sort = "name"
                RsAcces.MoveFirst
                RsAcces.FIND "Code ='" & ProdCode & "'"
                If RsAcces.EOF = True Then RsAcces.MoveFirst
                FGrid1.AddItem ""
                FGrid1.TextMatrix(I, Col_Accessory) = RsAcces!Name
                FGrid1.TextMatrix(I, Col_AccCode) = RsAcces!Code
                I = I + 1
                ProdCode = ""
        End If
        
        I = 1
        For j = 1 To Len(XNull(Master1!AccQty))
            If mID(Master1!AccQty, j, 1) <> "," Then
                ProdQty = ProdQty & mID(Master1!AccQty, j, 1)
            Else
                FGrid1.TextMatrix(I, Col_Qty) = ProdQty
                ProdQty = ""
                I = I + 1
            End If
        Next
'        If ProdQty <> "" Then
'                FGrid1.TextMatrix(I, Col_Qty) = ProdQty
'                ProdQty = ""
'        End If
        If FGrid1.Rows > 1 Then
            FGrid1.FixedRows = 1
        Else
            FGrid1.AddItem ""
            FGrid1.FixedRows = 1
        End If
        
        For K = 1 To FGrid1.Rows - 1
            FGrid1.TextMatrix(K, 0) = K
        Next
        
        
Else
    Call BlankText
End If
Set Rst = Nothing
Grid_Hide
Exit Sub
error1:
        CheckError
End Sub

Private Sub Disp_Text(Enb As Boolean)
Dim I As Integer
For I = 0 To Txt.Count - 1
    Txt(I).Enabled = Enb
    Txt(I).ForeColor = CtrlFColOrg
Next
If UCase(left(PubComp_Name, 4)) = "ENAR" Then Txt(Site).Enabled = False
FGrid.Enabled = Enb
If TopCtrl1.TopText2 = "Edit" Then
    Txt(Site).Enabled = False
    Txt(BookNo).Enabled = False
    'txt(VDate).Enabled = False
    Txt(QuotNo).Enabled = False
 
    
End If
    Txt(QuotPrefix).Enabled = False
    'txt(NamePrefix).Enabled = False
    Txt(FNamePrefix).Enabled = False
    Txt(fname).Enabled = False
    Txt(Add1).Enabled = False
    Txt(Add2).Enabled = False
    Txt(Add3).Enabled = False
    Txt(Phone).Enabled = False
    Txt(City).Enabled = False
    Txt(PinCode).Enabled = False
    Txt(InvNo).Enabled = False
    Txt(InvDate).Enabled = False
    Txt(DelChNo).Enabled = False
    Txt(DelChDate).Enabled = False

    If FGrid.Enabled = False Then
        FGrid.BackColor = CtrlBColDisabled
    Else
        FGrid.BackColor = CtrlBColOrg
    End If

Txt(ModelDesc).Enabled = False

End Sub
Private Sub Grid_Hide()
    If DGCity.Visible = True Then DGCity.Visible = False
    If DGSite.Visible = True Then DGSite.Visible = False
    If DGParty.Visible = True Then DGParty.Visible = False
    If lblGroup.Visible = True Then lblGroup.Visible = False
    If DGCol.Visible = True Then DGCol.Visible = False
    If DGMod.Visible = True Then DGMod.Visible = False
    If DGFin.Visible = True Then DGFin.Visible = False
    If DGRep.Visible = True Then DGRep.Visible = False
    If DGRef.Visible = True Then DGRef.Visible = False
    If DGProf.Visible = True Then DGProf.Visible = False
    If DGArea.Visible = True Then DGArea.Visible = False
    If DGVno.Visible = True Then DGVno.Visible = False
    If DGPurpose.Visible = True Then DGPurpose.Visible = False
    If DGQuot.Visible = True Then DGQuot.Visible = False
    If FrmList.Visible = True Then FrmList.Visible = False
    If DgAcces.Visible = True Then DgAcces.Visible = False
End Sub


Private Sub Ini_Grid()
Dim Rs As Recordset, MeWidth As Long
MeWidth = Me.width

 DGParty.Columns(3).width = 1544.882
 DGParty.Columns(2).width = 3330.142
 DGParty.Columns(1).width = 3254.74
 DGParty.Columns(0).width = 4275.213


DGParty.left = 15: DGParty.width = MeWidth - 90: DGParty.top = Me.height - (DGParty.height + mBotScale)
DGCity.left = MeWidth - (DGCity.width + mRtScale): DGCity.top = mTopScale: DGCity.height = 4935
DGRep.left = MeWidth - (DGRep.width + mRtScale): DGRep.top = mTopScale
DGProf.left = MeWidth - (DGProf.width + mRtScale): DGProf.top = mTopScale
DGRef.left = MeWidth - (DGRef.width + mRtScale): DGRef.top = mTopScale
DGArea.left = MeWidth - (DGArea.width + mRtScale): DGArea.top = mTopScale
DGMod.left = MeWidth - (DGMod.width + mRtScale): DGMod.top = mTopScale

DGPurpose.left = MeWidth - (DGPurpose.width + mRtScale):  DGPurpose.top = mTopScale
DGCol.left = MeWidth - (DGCol.width + mRtScale): DGCol.top = mTopScale
DGFin.left = 190 'MeWidth - (DGFin.width + mRtScale): DGFin.top = mTopScale
DGFin.top = 1250
DGQuot.left = MeWidth - (DGQuot.width + mRtScale): DGQuot.top = Me.height - (DGQuot.height + mBotScale)
DGSite.left = 3000: DGSite.top = mTopScale
DGVno.left = 3000: DGVno.top = mTopScale
DgAcces.left = DGRep.left: DgAcces.top = DGRep.top


With FGrid
    .left = Label3(5).left ' 170
    .top = Label3(5).top + Label3(5).height + 45
    .Cols = 3
    .width = 2200
    .BackColorBkg = Me.BackColor
    .BackColorFixed = Me.BackColor
    .GridColorFixed = Me.BackColor
    .ColWidth(0) = 1710
    .ColWidth(1) = 300
    .ColWidth(2) = 0
End With
    Set Rs = New Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open "Select AddSrvName,AddSrvcode from Veh_AddiService order by AddSrvName", GCn, adOpenDynamic, adLockOptimistic
    FGrid.Rows = 0
    If Rs.RecordCount > 0 Then
        Do Until Rs.EOF
            FGrid.AddItem Rs.Fields(0).Value & Chr(9) & "" & Chr(9) & Rs.Fields(1).Value
            Rs.MoveNext
        Loop
    FGrid.height = (FGrid.RowHeightMin * Rs.RecordCount) + 100
    End If
    Set Rs = Nothing
    
    With FGrid1
        .RowHeightMin = PubGridRowHeight '220
        .height = FGrid1.RowHeight(0) * 7
        .Cols = 4
        .left = 7125
        .top = 5220
        .TextMatrix(0, 0) = "Sr.No"
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 500
        .TextMatrix(1, 0) = "1"
        
        .TextMatrix(0, Col_AccCode) = ""
        .ColAlignment(Col_AccCode) = flexAlignLeftCenter
        .ColWidth(Col_AccCode) = 0
        
    
        .TextMatrix(0, Col_Accessory) = "Accessories"
        .ColAlignment(Col_Accessory) = flexAlignLeftCenter
        .ColWidth(Col_Accessory) = 3000
        
        .TextMatrix(0, Col_Qty) = "Qty"
        .ColAlignment(Col_Qty) = flexAlignRightCenter
        .ColWidth(Col_Qty) = 700
        
    End With
    BackColorSelLeave = FGrid1.BackColorSel
    ForeColorSelEnter = FGrid1.ForeColorSel
    
    
End Sub

Private Sub Fill_Data()
Dim Rst As ADODB.Recordset
Dim RstSubGrp As ADODB.Recordset
Dim RstBook As ADODB.Recordset
If GCn.Execute("select count(*) from veh_Quot where veh_quot.docid = '" & Txt(QuotNo).Tag & "'").Fields(0).Value > 0 Then
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open "select veh_Quot.*,veh_quot1.srl_no, veh_quot1.model,veh_quot1.rate,veh_quot1.Book_DocId from veh_Quot left join veh_quot1 on veh_quot1.docid = veh_quot.docid where veh_quot1.docid = '" & Txt(QuotNo).Tag & "' and model = '" & RsQuot!Model & "'", GCn, adOpenDynamic, adLockBatchOptimistic
    
    Txt(Party_code).TEXT = "": Txt(Party_code).Enabled = True
    Txt(NamePrefix).TEXT = "": Txt(NamePrefix).Enabled = True
    Txt(FNamePrefix).TEXT = "": Txt(FNamePrefix).Enabled = True
    Txt(fname).TEXT = "": Txt(fname).Enabled = True
    Txt(Add1).TEXT = "": Txt(Add1).Enabled = True
    Txt(Add2).TEXT = "": Txt(Add2).Enabled = True
    Txt(Add3).TEXT = "": Txt(Add3).Enabled = True
    Txt(Phone).TEXT = "": Txt(Phone).Enabled = True
    Txt(City).TEXT = "": Txt(City).Enabled = True
    Txt(PinCode).TEXT = "": Txt(PinCode).Enabled = True
    Txt(City).Tag = "": Txt(Party_code).Tag = ""
    Txt(Purpose).TEXT = "": Txt(Purpose).Enabled = True
    Txt(IndUse).Tag = "": Txt(IndUse).Tag = ""
Else
    GoTo FillBlank
End If

If Rst!Book_DocId <> "" Then
    MsgBox "Booking is already done against this Quotation!!! ", vbInformation, "Duplicate Quotation!"
    GoTo FillBlank
Else
    Set RstBook = GCn.Execute("select NPrefix,Name,NSuffix,FPrefix,FName,Add1,Add2,Add3,PhoneOff,CityCode from ProspectiveCust where Cust_Code = '" & Rst!Party_code & "'")
    
    Set RstSubGrp = GCn.Execute("select SubCode,NamePrefix,Name,FPrefix,FName,Add1,Add2,Add3,Phone,CityCode,PANNo from subgroup " & _
    "where name  = '" & RstBook!Name & " " & RstBook!NSuffix & _
    "' and add1  = '" & RstBook!Add1 & _
    "' and add2  = '" & RstBook!Add2 & _
    "' and add3  = '" & RstBook!Add3 & _
    "' and citycode  = '" & RstBook!CityCode & "'")
    If RstSubGrp.RecordCount > 0 Then
        Txt(Party_code).TEXT = RstSubGrp!Name
        Txt(Party_code).Tag = RstSubGrp!SubCode
        Txt(NamePrefix).TEXT = RstSubGrp!NamePrefix
        Txt(FNamePrefix).TEXT = RstSubGrp!FPrefix
        Txt(fname).TEXT = RstSubGrp!fname
        Txt(Add1).TEXT = RstSubGrp!Add1
        Txt(Add2).TEXT = RstSubGrp!Add2
        Txt(Add3).TEXT = RstSubGrp!Add3
        Txt(Phone).TEXT = RstSubGrp!Phone
        Txt(City).Tag = RstSubGrp!CityCode
        Txt(PAN) = RstSubGrp!PanNo
        
    Else
        Txt(Party_code).TEXT = RstBook!Name & " " & RstBook!NSuffix
        Txt(Party_code).Tag = ""
        Txt(NamePrefix).TEXT = RstBook!NPrefix
        Txt(FNamePrefix).TEXT = RstBook!FPrefix
        Txt(fname).TEXT = RstBook!fname
        Txt(Add1).TEXT = RstBook!Add1
        Txt(Add2).TEXT = RstBook!Add2
        Txt(Add3).TEXT = RstBook!Add3
        Txt(Phone).TEXT = RstBook!PhoneOff
        Txt(City).Tag = RstBook!CityCode
    End If
    If Txt(City).Tag <> "" And GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").RecordCount > 0 Then
        Txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").Fields(0).Value
    Else
        Txt(City).TEXT = ""
    End If
    
    Txt(Profession).Tag = IIf(IsNull(Rst!Profession), "", Rst!Profession)
    If Txt(Profession).Tag <> "" And GCn.Execute("select Professionname from Profession where Professioncode = '" & Txt(Profession).Tag & "'").RecordCount > 0 Then
        Txt(Profession).TEXT = GCn.Execute("select Professionname from Profession where Professioncode = '" & Txt(Profession).Tag & "'").Fields(0).Value
    Else
        Txt(Profession).TEXT = ""
    End If
    Txt(Area).Tag = IIf(IsNull(Rst!Area), "", Rst!Area)
    If Txt(Area).Tag <> "" And GCn.Execute("select AREAname from AREA where AREAcode = '" & Txt(Area).Tag & "'").RecordCount > 0 Then
        Txt(Area).TEXT = GCn.Execute("select AREAname from AREA where AREAcode = '" & Txt(Area).Tag & "'").Fields(0).Value
    Else
        Txt(Area).TEXT = ""
    End If
    Txt(REP_CODE).Tag = IIf(IsNull(Rst!REP_CODE), "", Rst!REP_CODE)
    If Txt(REP_CODE).Tag <> "" And GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(REP_CODE).Tag & "'").RecordCount > 0 Then
        Txt(REP_CODE).TEXT = GCn.Execute("select Emp_name from Emp_mast where Emp_Code = '" & Txt(REP_CODE).Tag & "'").Fields(0).Value
    Else
        Txt(REP_CODE).TEXT = ""
    End If
    Txt(REF_CODE).Tag = IIf(IsNull(Rst!REF_CODE), "", Rst!REF_CODE)
    If Txt(REF_CODE).Tag <> "" And GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(REF_CODE).Tag & "'").RecordCount > 0 Then
        Txt(REF_CODE).TEXT = GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(REF_CODE).Tag & "'").Fields(0).Value
    Else
        Txt(REF_CODE).TEXT = ""
    End If
    Txt(Model).TEXT = Rst!Model
    Txt(Rate).TEXT = Format(Rst!Rate, "0.00")
    
    If Txt(REF_CODE).Tag <> "" And GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(REF_CODE).Tag & "'").RecordCount > 0 Then
        Txt(REF_CODE).TEXT = GCn.Execute("select Refname from Reffered where RefCode = '" & Txt(REF_CODE).Tag & "'").Fields(0).Value
    Else
        Txt(REF_CODE).TEXT = ""
    End If
    Txt(Purpose).Tag = IIf(IsNull(Rst!Purpose), "", Rst!Purpose)
    If Txt(Purpose).Tag <> "" And GCn.Execute("select PurposeName from Purpose where PurposeCode = '" & Txt(Purpose).Tag & "'").RecordCount > 0 Then
        Txt(Purpose).TEXT = GCn.Execute("select PurposeName from Purpose where PurposeCode = '" & Txt(Purpose).Tag & "'").Fields(0).Value
    Else
        Txt(Purpose).TEXT = ""
    End If
    Txt(IndUse).TEXT = IIf(IsNull(Rst!INTD_USE), "", Rst!INTD_USE)
    
    QDocId = Rst!DocID
    Txt(QuotPrefix) = mID(QDocId, 9, 13)
    QSiteCode = Rst!Site_Code
    QDocSrlNo = Rst!Srl_No
    Txt(GovtYn) = IIf(Rst!Govt_YN = 1, "Yes", "No")
    Txt(FirstVeh) = IIf(Rst!FirstVeh_YN = 1, "Yes", "No")
    Txt(ExpDelDt) = IIf(IsNull(Rst!DEL_DATE), "", Rst!DEL_DATE)
    Txt(FB_Code).Tag = IIf(IsNull(Rst!FB_Code), "", Rst!FB_Code)
    If Txt(FB_Code).Tag <> "" Then
        Txt(FB_Code).TEXT = GCn.Execute("select finname as name from ContractFinance where fincatg = 0 and  fincode = '" & Txt(FB_Code).Tag & "'").Fields(0).Value
    Else
        Txt(FB_Code).TEXT = ""
    End If
End If
Set RstBook = Nothing
'txtDisabled_Color Me
GoTo ExitFillData
FillBlank:
    Txt(Party_code).TEXT = "": Txt(Party_code).Tag = ""
    Txt(Add1).TEXT = "": Txt(Add1).Enabled = False
    Txt(Add2).TEXT = "": Txt(Add2).Enabled = False
    Txt(Add3).TEXT = "": Txt(Add3).Enabled = False
    Txt(Phone).TEXT = "": Txt(Phone).Enabled = False
    Txt(City).TEXT = "": Txt(City).Enabled = False
    Txt(City).Tag = ""
    Txt(PinCode).TEXT = "": Txt(PinCode).Enabled = False
    Txt(QuotNo).TEXT = "":  Txt(Profession).Tag = "":   Txt(Profession).TEXT = ""
    Txt(Area).Tag = "":     Txt(Area).TEXT = "":        Txt(REP_CODE).Tag = ""
    Txt(REP_CODE).TEXT = "": Txt(REF_CODE).Tag = "":    Txt(REF_CODE).TEXT = ""
    QDocId = "": QDocSrlNo = 0:     QSiteCode = "":     Txt(GovtYn) = "No"
    Txt(FirstVeh) = "":     Txt(ExpDelDt) = "":         Txt(FB_Code).Tag = ""
    Txt(Model) = "":     Txt(Rate) = ""
    Txt(FB_Code).TEXT = ""
'    txtDisabled_Color Me
ExitFillData:
Set Rst = Nothing
End Sub

Private Sub CreateCustomer()
Dim VSrNo As Integer
    If GCn.Execute("select count(*) from ProspectiveCust where site_code = '" & Txt(Site).Tag & "'").Fields(0).Value > 0 Then
        VSrNo = GCn.Execute("select MAX(right(cust_CODE,7)) from ProspectiveCust  where site_code = '" & PubSiteCode & "'").Fields(0).Value + 1
    Else
        VSrNo = 1
    End If
    CustDocId = PubSiteCode + Space(7 - Len(CStr(VSrNo))) + CStr(VSrNo)
    GCn.Execute ("delete from ProspectiveCust where cust_Code='" & CustDocId & "'")
    GCn.Execute ("insert into ProspectiveCust(cust_Code,Site_Code,NPrefix,Name,FPrefix,FName,Govt_YN,Add1,Add2,Add3,PhoneOff,CityCode, " & _
    " AREA,REF_CODE,REP_CODE,Profession,FirstVeh_YN,U_Name, U_EntDt, U_AE ) " & _
      " values('" & CustDocId & "','" & Txt(Site).Tag & "','" & Txt(NamePrefix) & "','" & Txt(Party_code) & "','" & Txt(FNamePrefix) & "','" & Txt(fname) & "'," & IIf(Txt(GovtYn).TEXT = "Yes", 1, 0) & ",'" & Txt(Add1).TEXT & "','" & Txt(Add2).TEXT & "','" & Txt(Add3).TEXT & "','" & Txt(Phone) & "','" & Txt(City).Tag & "', " & _
    " '" & Txt(Area).Tag & "','" & Txt(REF_CODE).Tag & "','" & Txt(REP_CODE).Tag & "','" & Txt(Profession).Tag & "'," & IIf(Txt(FirstVeh).TEXT = "Yes", 1, 0) & ",'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
End Sub

Private Function CreateQuotation() As Boolean
Dim mVType2$, DocIdHlp$
Dim VoucherEditFlag2 As Boolean
Dim Rst As ADODB.Recordset, VNo As Long, ExpDate$, ValidQuot As Integer
If CustDocId = "" Then CustDocId = Txt(Party_code).Tag
    mVType2 = "V_QOT"
    QDocId = GetDocID(GCnFaV, mVType2, Txt(VDate), VoucherEditFlag2, Txt(SerialNo2), LblVPrefix2)
    DocIdHlp = Replace(QDocId, " ", "")
    QDocSrlNo = 1
    QSiteCode = PubSiteCode & Txt(Site).Tag
    
    ValidQuot = GCn.Execute("select Valid_Day from syctrl").Fields(0).Value
    ExpDate = Format(DateAdd("D", ValidQuot, Txt(VDate).TEXT), "dd/mmm/yyyy")
    GCn.Execute "insert into Veh_Quot(DocId,DocIDHelp,V_Type,V_No,Site_Code, " & _
        "V_Date,Party_Code,CityCode, " & _
        "AREA,REF_CODE,REP_CODE,Profession, " & _
        "fin_yn,FB_CODE , Govt_YN, FirstVeh_YN, " & _
        "AMOUNT,DEL_DATE,EXP_DATE, RoundOff_YN,U_Name, U_EntDt, U_AE ) " & _
        "values('" & QDocId & "','" & DocIdHlp & "','" & mVType2 & "'," & VNo & ",'" & PubSiteCode & Txt(Site).Tag & "'," & _
        "" & ConvertDate(Txt(VDate).TEXT) & ",'" & CustDocId & "','" & Txt(City).Tag & "'," & _
        "'" & Txt(Area).Tag & "','" & Txt(REF_CODE).Tag & "','" & Txt(REP_CODE).Tag & "','" & Txt(Profession).Tag & "'," & _
        "" & IIf(Txt(FundSource) = "Hypothecation" Or Txt(FundSource) = "Hire Purchase", 1, 0) & ",'" & Txt(FB_Code).Tag & "'," & IIf(Txt(GovtYn) = "Yes", 1, 0) & "," & IIf(Txt(FirstVeh) = "Yes", 1, 0) & "," & _
        "" & Val(Txt(Rate)) & "," & ConvertDate(Txt(ExpDelDt).TEXT) & "," & ConvertDate(ExpDate) & ", 1,'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')"

    GCn.Execute ("delete from Veh_Quot1 where docid='" & QDocId & "'")
    GCn.Execute ("insert into Veh_Quot1(DocId,srl_no,DocIDHelp,V_Type,V_No,Site_Code, " & _
        "MODEL,QTY,RATE,amount,Book_SiteCode,Book_docid,  " & _
        "U_Name, U_EntDt, U_AE ) " & _
        " values('" & QDocId & "',1,'" & DocIdHlp & "','" & mVType2 & "'," & VNo & ",'" & Txt(Site).Tag & "', " & _
        " '" & Txt(Model).TEXT & "',1," & Val(Txt(Rate).TEXT) & "," & Val(Txt(Rate).TEXT) & ",'','" & DocID & "', " & _
        "'" & pubUName & "'," & ConvertDate(PubServerDate) & ",'A')")
    
    CreateQuotation = UpdVouSrlNo(GCnFaV, QDocId, Txt(VDate))
End Function
'************************ PRINTING CODE ******************


Private Sub TxtGrid_GotFocus(Index As Integer)
Ctrl_GetFocus TxtGrid(Index)
    Grid_Hide
    TxtGrid(0).Tag = FGrid1.TextMatrix(FGrid1.Row, FGrid1.Col)
    Select Case FGrid1.Col
        Case Col_Accessory
            If RsAcces.RecordCount = 0 Or FGrid1.TextMatrix(FGrid1.Row, Col_Accessory) = "" Then Exit Sub
            RsAcces.Sort = "name"
            RsAcces.MoveFirst
            RsAcces.FIND "name ='" & FGrid1.TextMatrix(FGrid1.Row, Col_Accessory) & "'"
            If RsAcces.EOF = True Then RsAcces.MoveFirst
    End Select
End Sub

Private Sub TxtGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ELoop
If KeyCode = vbKeyEscape Then
        TxtGrid(0).Visible = False
        Exit Sub
End If
    Select Case FGrid1.Col
        Case Col_Accessory
            DGridTxtKeyDown DgAcces, TxtGrid, Index, RsAcces, KeyCode, False, 1, frmVehAMDMast, "frmVehAMDMast"
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And DgAcces.Visible = False) Then
                If TxtGridLeave = True Then
                    GridTxtDown FGrid1, TxtGrid, Index, KeyCode, TAddMode, Col_Qty, , Col_Qty
                End If
            End If
        Case Col_Qty
            If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown) And TAddMode = True) Then
                If TxtGridLeave = True Then
                     GridTxtDown FGrid, TxtGrid, Index, KeyCode, TAddMode, Col_Qty, 5
                     FGrid1.Row = FGrid1.Row + 1
                     FGrid1.Col = Col_Accessory
                     FGrid1.SetFocus
                End If
            End If
            
    End Select
Exit Sub
ELoop:
MsgBox err.Description
End Sub

Private Sub TxtGrid_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Exit Sub
Call CheckQuote(KeyAscii)
Select Case FGrid1.Col
    Case Col_Accessory
        If DgAcces.Visible = True Then DGridTxtKeyPress TxtGrid, Index, RsAcces, KeyAscii, "Name"
    Case Col_Qty
        Call NumPress(TxtGrid(Index), KeyAscii, 6, 0)
End Select
End Sub

Private Sub TxtGrid_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case FGrid1.Col
        Case Col_Accessory
            If KeyCode <> 13 And DgAcces.Visible = False Then TxtGrid_KeyDown Index, KeyCode, 0: DGridTxtKeyPress TxtGrid, Index, RsAcces, KeyCode, "name", True
        Case Col_Qty
            FGrid1.TextMatrix(FGrid1.Row, Col_Qty) = Format(Val(TxtGrid(Index).TEXT), "0")
    End Select
    If KeyCode = vbKeyEscape Then
        FGrid1.SetFocus
        TxtGrid(0).Visible = False
        Grid_Hide
    End If
End Sub

Private Sub TxtGrid_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ELoop
Cancel = Not TxtGridLeave(Index, True)
Exit Sub
ELoop:
    CheckError
End Sub

Private Sub TxtPrint_GotFocus(Index As Integer)
Ctrl_GetFocus txtPrint(Index)
Grid_Hide
Select Case Index
    Case FromVno, ToVno
            RsVno.Close
            RsVno.Open "Select Ord_No as code from Veh_Order where right(veh_order.ord_SiteCode,1)='" & txtPrint(SiteCode1).Tag & "' and  veh_order.ord_VType='V_BK'", GCn, adOpenDynamic, adLockOptimistic
            Set DGVno.DataSource = RsVno
            If txtPrint(Index).TEXT <> RsVno!Code Then
                RsVno.MoveFirst
                RsVno.FIND "code ='" & txtPrint(Index).TEXT & "'"
            End If
            If Index = ToVno Then DGVno.Tag = "1" Else DGVno.Tag = "2"
    
    Case SiteCode1
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Then Exit Sub
        If txtPrint(Index).TEXT = "" Then
            RsSite.MoveFirst
            RsSite.FIND "code ='" & PubSiteCode & "'"
            txtPrint(Index).Tag = RsSite!Code
            txtPrint(Index).TEXT = RsSite!Name
        Else
            If txtPrint(Index).TEXT <> RsSite!Name Then
                RsSite.MoveFirst
                RsSite.FIND "name ='" & txtPrint(Index).TEXT & "'"
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
    Case FromVno, ToVno
        DGridTxtKeyDown DGVno, txtPrint, Index, RsVno, KeyCode, False, 0
    Case SiteCode1
        DGridTxtKeyDown DGSite, txtPrint, Index, RsSite, KeyCode, False, 1
End Select
If DGSite.Visible = False And DGVno.Visible = False Then
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyReturn) Then Ctrl_DownKeyDown KeyCode, Shift
    If KeyCode = vbKeyUp And Index <> SiteCode1 Then Ctrl_UpKeyDown KeyCode, Shift
End If
End Sub

Private Sub TxtPrint_KeyPress(Index As Integer, KeyAscii As Integer)
'Sequence : KeyDown- >KeyPress- >KeyUp
'Validate- >LostFoucs
Call CheckQuote(KeyAscii)
Select Case Index
    Case FromVno, ToVno
        If DGVno.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsVno, KeyAscii, "Code"
    Case SiteCode1
        If DGSite.Visible = True Then DGridTxtKeyPress txtPrint, Index, RsSite, KeyAscii, "Name"
End Select

'KeyAscii = RetDGKeyAscii()
End Sub

Private Sub TxtPrint_LostFocus(Index As Integer)
  Ctrl_validate txtPrint(Index)
End Sub

Private Sub TxtPrint_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
    Case ToVno, FromVno
        If RsVno.RecordCount = 0 Or (RsVno.EOF = True Or RsVno.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(Index).TEXT = ""
        Else
            txtPrint(Index).TEXT = RsVno!Code
        End If
    Case SiteCode1
        If IsValid(txtPrint(Index), "Site Code") = False Then Cancel = True: Exit Sub
        If RsSite.RecordCount = 0 Or (RsSite.EOF = True Or RsSite.BOF = True) Or txtPrint(Index).TEXT = "" Then
            txtPrint(Index).TEXT = ""
            txtPrint(Index).Tag = ""
        Else
            txtPrint(Index).TEXT = RsSite!Name
            txtPrint(Index).Tag = RsSite!Code
        End If

End Select
End Sub
Private Sub CmdPrint_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    FrmPrn.Visible = False
    If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
        If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
        Disp_Text SETS("INI", Me, Master)
        Call MoveRec
    End If
End If
End Sub
Private Sub CmdPrint_Click(Index As Integer)
On Error GoTo ERRORHANDLER
GSQL = "SELECT Syctrl.VehBookFooter,Veh_Order.Ord_No,Veh_Order.exp_date,Veh_Order.rate, Veh_Order.Ord_Date, Model_Grp.MODELGrp_Name as Model, " & _
        "Veh_Order.INTD_USE,Veh_Order.Permit_N_Z, Veh_Order.AddVeri_YN, Veh_Order.FIN_AMT, Veh_Order.Fund_Source, ContractFinance.FinName,ContractFinance.Add1 as FinAdd1,ContractFinance.add2 as FinAdd2,c1.cityname as fincity,SubGroup.NamePrefix,SubGroup.Name," & _
        "SubGroup.add1,SubGroup.add2,SubGroup.add3,SubGroup.phone,SubGroupType.Description, Profession.ProfessionName, Purpose.PurposeName, City.CityName, Site.Site_Desc,Book_UName,Book_UEntDt,SubGroup.FPrefix,SubGroup.FName, SubGroup.Pin, State.StateName, Model.Model_Desc, Veh_Order.DoNo,DoReciveDate,DOIssueDate " & _
        "FROM (((((((((((Veh_Order LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) LEFT JOIN ContractFinance " & _
        "ON Veh_Order.FB_CODE = ContractFinance.FinCode) left join city c1 on ContractFinance.city = c1.citycode) LEFT JOIN Profession " & _
        "ON Veh_Order.Profession = Profession.ProfessionCode)" & _
        "Left Join Model on Veh_Order.Model=Model.Model) " & _
        "Left Join Model_Grp on Model.Grp_Code=Model_Grp.ModelGrp_Code) " & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable  >= " & xIsNull("Veh_Order.Inv_UAE", "A") & ")" & _
        "LEFT JOIN Purpose ON Veh_Order.PURPOSE = Purpose.PurposeCode) " & _
        "LEFT JOIN City ON SubGroup.CityCode = City.CityCode) " & _
        "Left Join State On State.StateCode = City.StateCode)  " & _
        "LEFT JOIN Site ON right(Veh_Order.Ord_SiteCode,1) = Site.Site_Code) " & _
        "LEFT JOIN SubGroupType ON SubGroup.Party_Type = SubGroupType.Party_Type " & _
        "where Veh_Order.OrdDocId ='" & Master!SearchCode & "'"

Select Case Index
    Case PScreen, PWindows
        mRepName = IIf(OptPlain.Value = True, "VehBooking", "VehBooking")
        WindowsPrint Index, GSQL
        FrmPrn.Visible = False
    Case PDos
        SpeedPrint GSQL
        FrmPrn.Visible = False
    Case PSetUp
        mRepName = IIf(OptPlain.Value = True, "VehBooking", "VehBooking")
        PrinerSetUp
    Case PClose 'Close Report Frame
        FrmPrn.Visible = False
        CmdPrint(PSetUp).Tag = ""
End Select
If Index <> PSetUp And TopCtrl1.TopText2.CAPTION <> "Browse" Then
    If TopCtrl1.TopText2.CAPTION = "Add" Then TopCtrl1_eAdd: Exit Sub
    Disp_Text SETS("INI", Me, Master)
    Call MoveRec
End If
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.CAPTION
End Sub
Private Sub WindowsPrint(Index As Integer, mQry As String)
Dim Rst As ADODB.Recordset ', mQry As String
Dim I As Integer
Dim RST1 As ADODB.Recordset
On Error GoTo ERRORHANDLER
'    mQry = "SELECT Syctrl.VehBookFooter,Veh_Order.Ord_No,Veh_Order.exp_date,Veh_Order.rate, Veh_Order.Ord_Date, Veh_Order.MODEL, " & _
        "Veh_Order.INTD_USE,Veh_Order.Permit_N_Z, Veh_Order.AddVeri_YN, Veh_Order.FIN_AMT, Veh_Order.Fund_Source, ContractFinance.FinName,ContractFinance.Add1 as FinAdd1,ContractFinance.add2 as FinAdd2,c1.cityname as fincity,SubGroup.NamePrefix,SubGroup.Name," & _
        "SubGroup.add1,SubGroup.add2,SubGroup.add3,SubGroup.phone,SubGroupType.Description, Profession.ProfessionName, Purpose.PurposeName, City.CityName, Site.Site_Desc,Book_UName,Book_UEntDt,SubGroup.FPrefix,SubGroup.FName " & _
        "FROM ((((((((Veh_Order LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) LEFT JOIN ContractFinance " & _
        "ON Veh_Order.FB_CODE = ContractFinance.FinCode) left join city c1 on ContractFinance.city = c1.citycode) LEFT JOIN Profession " & _
        "ON Veh_Order.Profession = Profession.ProfessionCode)" & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable  >=Veh_Order.Inv_UAE)" & _
        "LEFT JOIN Purpose ON Veh_Order.PURPOSE = Purpose.PurposeCode) LEFT JOIN City ON SubGroup.CityCode = City.CityCode) LEFT JOIN Site ON right(Veh_Order.Ord_SiteCode,1) = Site.Site_Code) LEFT JOIN SubGroupType ON SubGroup.Party_Type = SubGroupType.Party_Type where Veh_Order.OrdDocId ='" & Master!SearchCode & "'"
    
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQry), GCn, adOpenDynamic, adLockOptimistic
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.CAPTION: Exit Sub
   
    CreateFieldDefFile Rst, PubRepoPath + "\" & mRepName & ".ttx", True
    If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
            
        Set RST1 = New Recordset
        RST1.CursorLocation = adUseClient
        RST1.Open "select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax,V_SecGram from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubVCompCode & "'", GCn, adOpenDynamic, adLockOptimistic

        For I = 1 To rpt.FormulaFields.Count
            Select Case UCase(rpt.FormulaFields(I).FormulaFieldName)
                Case UCase("SubTitle")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!V_SecSpeciality & "'"
                Case UCase("Phone")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!V_SecPhone & "'"
                Case UCase("Fax")
                    rpt.FormulaFields(I).TEXT = "'" & RST1!V_SecFax & "'"
                Case UCase("Title")
            End Select
        Next
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
    Select Case Index
        Case PWindows   'Printer
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
'            If MsgBox("Printed Properly ? ", vbYesNo + vbCritical + vbDefaultButton1, "Printed Document !") = vbYes Then
'                GCn.Execute "update Sp_order set Printed = 1  where Sp_order.orderid='" & Master!OrderId & "'"
'            End If
        Case PScreen  'screen
            Call Report_View(rpt, Me.CAPTION, , True)
    End Select
CmdPrint(PSetUp).Tag = ""
Exit Sub
ERRORHANDLER:
        CheckError
End Sub

Private Sub PrinerSetUp()
If CmdPrint(PSetUp).Tag = "" Then Set rpt = rdApp.OpenReport(PubRepoPath + "\" & mRepName & ".RPT")
rpt.PrinterSetup (0)
CmdPrint(PSetUp).Tag = "1"
LblPrinter.CAPTION = rpt.PrinterName
End Sub

Private Sub SpeedPrint(mQry As String)
On Error GoTo ELoop
'Paper Size 8.5*12
'Total Lines Per PAge 72
'Top Margin  3 Lines  (For 1/2 Inch)
'Header 15 Lines
'Footer 23 Lines
'Bottom Margin  3 Lines  (For 1/2 Inch)
'Contd. Remarks 2 Lines
'Gate Pass Detail 8 Lines
'Print Area 18
    Dim I As Integer, j As Integer
    Dim PrintStr As String
    Dim Rs As ADODB.Recordset, RstCompDet As ADODB.Recordset, RstBook As ADODB.Recordset
    Dim LAdd As Double, Page As Byte, mLine As Byte, mFix As Byte
    Dim mQty As Double, mAmount As Double, mSlNo As Integer, PageWidth As Byte, PageLength As Integer
    Dim mDocStr$, mDupStr$, Speciality$, mTaxdesc As String, mGoods_Amt As Double
    Dim Footer As String, FooterCnt As Byte, mHeader As Byte, mFooter As Byte
    Dim SubTot As Double
    Dim fob As New FileSystemObject
    
 
'    GSQL = "SELECT Syctrl.VehBookFooter,Veh_Order.Ord_No, Veh_Order.Ord_Date, Veh_Order.MODEL,Veh_Order.EXP_DATE,Veh_Order.rate, " & _
        "Veh_Order.INTD_USE,Veh_Order.Permit_N_Z, Veh_Order.AddVeri_YN, Veh_Order.FIN_AMT, Veh_Order.Fund_Source, ContractFinance.FinName,ContractFinance.Add1 as FinAdd1,ContractFinance.add2 as FinAdd2,c1.cityname as FinCity, SubGroup.Name," & _
        "SubGroup.add1,SubGroup.add2,SubGroup.add3,SubGroup.phone,SubGroupType.Description, Profession.ProfessionName, Purpose.PurposeName, City.CityName, Site.Site_Desc,Book_UName,Book_UEntDt,SubGroup.FPrefix,SubGroup.FName " & _
        "FROM ((((((((Veh_Order LEFT JOIN SubGroup ON Veh_Order.PartyCode = SubGroup.SubCode) LEFT JOIN ContractFinance " & _
        "ON Veh_Order.FB_CODE = ContractFinance.FinCode) left join city c1 on ContractFinance.city = c1.citycode) LEFT JOIN Profession " & _
        "ON Veh_Order.Profession = Profession.ProfessionCode)" & _
        "LEFT JOIN Syctrl ON Syctrl.LinkTable  >=Veh_Order.Inv_UAE)" & _
        "LEFT JOIN Purpose ON Veh_Order.PURPOSE = Purpose.PurposeCode) LEFT JOIN City ON SubGroup.CityCode = City.CityCode) LEFT JOIN Site ON right(Veh_Order.Ord_SiteCode,1) = Site.Site_Code) LEFT JOIN SubGroupType ON SubGroup.Party_Type = SubGroupType.Party_Type where Veh_Order.OrdDocId ='" & Master!SearchCode & "'"
        
    Set RstBook = GCn.Execute(mQry)
    If RstBook.RecordCount <= 0 Then MsgBox "No Records To Print....", vbInformation, Me.CAPTION: Exit Sub
    If fob.FileExists("C:\RepPrint.Txt") = False Then
        fob.CreateTextFile ("C:\RepPrint.Txt")
    End If
    If fob.FileExists("C:\RepPrint.Bat") = False Then
        fob.CreateTextFile ("C:\RepPrint.Bat")
    End If
    Close #1
    Open "C:\RepPrint.Txt" For Output As #1
    
    FooterCnt = 1
    Footer = XNull(GCn.Execute("select VehBookFooter from Syctrl").Fields(0).Value)
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            FooterCnt = FooterCnt + 1
        End If
    Next
 
    PageLength = PubPageLength
    PageWidth = 80
    mHeader = 0   'Ideal 17
    mFooter = 8
    mFooter = mFooter + FooterCnt
      
    ' Header
      
      mDocStr = "Vehicle Booking"
      mDupStr = ""

        Set RstCompDet = GCn.Execute("select V_SecSpeciality,V_SecLST,V_SecLST_Date,V_SecCST,V_SecCST_Date,V_SecPhone,V_SecFax from division where Div_Code='" & PubDivCode & "' and V_SecCompCode =  '" & PubSCompCode & "'")

         Print #1, PRN_TIT(PubComp_Name, "A", PageWidth)
         mHeader = mHeader + 1
         If XNull(RstCompDet!V_SecSpeciality) <> "" Then
             Print #1, PRN_TIT(RstCompDet!V_SecSpeciality, "C", PageWidth)
             mHeader = mHeader + 1
         End If
         Print #1, PRN_TIT(PubComp_Add, "C", PageWidth)
         mHeader = mHeader + 1
         
         If PubComp_Add2 <> "" Or PubComp_City <> "" Then
             Print #1, PRN_TIT(PubComp_Add2 & IIf(PubComp_Add2 = "" Or PubComp_City = "", "", ",") & PubComp_City, "C", PageWidth)
             mHeader = mHeader + 1
         End If
         'Print #1, PRN_TIT((IIf(XNull(RstCompDet!V_SecPhone) = "", "", "PHONE : ") & XNull(RstCompDet!V_SecPhone) & IIf(XNull(RstCompDet!V_SecFax) = "", "", " Fax   : ") & XNull(RstCompDet!V_SecFax)), "C", PageWidth)
        If PubComp_Contact <> "" Then
            Print #1, PRN_TIT(PubComp_Contact, "C", PageWidth)
            mHeader = mHeader + 1
        End If
'         Print #1, PSTR(XNull(RstCompDet!V_SecCST) & IIf(XNull(RstCompDet!V_SecCST_Date) = "", "", " Dt. " & RstCompDet!V_SecCST_Date), 40) & PSTR(XNull(RstCompDet!V_SecLST) & IIf(XNull(RstCompDet!V_SecLST_Date) = "", "", " Dt. " & RstCompDet!V_SecLST_Date), 40, , AlignRight)
'         mHeader = mHeader + 1
         Print #1, ""
         mHeader = mHeader + 1
         Print #1, PRN_TIT("** " & mDocStr & mDupStr & " **", "A", PageWidth) & mChr18 & mEmph
         mHeader = mHeader + 1
         Print #1, ""
         mHeader = mHeader + 1
        
        Print #1, mChr18 & "I/we Wish to book one " & mEmph & RstBook!Model & mEmph1 & " vehicle for my/our operation under"
        mHeader = mHeader + 1
        Print #1, IIf(RstBook!Permit_N_Z = 0, "National", "Zonal") & "permit.I have read the terms & conditions as printed below"
        mHeader = mHeader + 1
        Print #1, "and I/we shall abide with the same:" & mEmph
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PSTR("Booking No.", 20) & " : " & PSTR(STR(RstBook!Ord_No), 20) & PSTR("Booking Date", 15) & " : " & PSTR(STR(RstBook!Ord_Date), 14) & mEmph1
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PSTR("Customer Name", 20) & " : " & RstBook!Name
        mHeader = mHeader + 1
        Print #1, Space(23) & RstBook!FPrefix & " " & RstBook!fname
        mHeader = mHeader + 1
        Print #1, Space(23) & XNull(RstBook!Add1) & IIf(XNull(RstBook!Add1) <> "" And XNull(RstBook!Add2) <> "", ",", "") & XNull(RstBook!Add2)
        mHeader = mHeader + 1
        If XNull(RstBook!Add3) <> "" Then
            Print #1, Space(23) & XNull(RstBook!Add3) & IIf(XNull(RstBook!Add3) <> "" And XNull(RstBook!CityName) <> "", ",", "") & XNull(RstBook!CityName) & "-" & XNull(RstBook!Pin)
            mHeader = mHeader + 1
        Else
            Print #1, Space(23) & IIf(XNull(RstBook!CityName) <> "", XNull(RstBook!CityName) & "-", "Pin : ") & XNull(RstBook!Pin)
            mHeader = mHeader + 1
        End If
        If XNull(RstBook!StateName) <> "" Then
            Print #1, Space(23) & "State : " & XNull(RstBook!StateName)
            mHeader = mHeader + 1
        End If
        
        Print #1, PSTR("Phone No.", 20) & " : " & RstBook!Phone
        mHeader = mHeader + 1
        Print #1, PSTR("Profession", 20) & " : " & RstBook!ProfessionName
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PSTR("Model of Vehicle", 20) & " : " & RstBook!Model
        mHeader = mHeader + 1
        Print #1, PSTR("Rate of Vehicle", 20) & " : " & Format(RstBook!Rate, "0.00")
        mHeader = mHeader + 1
        Print #1, PSTR("Purpose of Use", 20) & " : " & RstBook!PurposeName
        mHeader = mHeader + 1
        Print #1, PSTR("Intended Use", 20) & " : " & RstBook!INTD_USE
        mHeader = mHeader + 1
        Print #1, PSTR("Customer Catg", 20) & " : " & RstBook!Description
        mHeader = mHeader + 1
        Print #1, PSTR("Address Verify", 20) & " : " & IIf(RstBook!AddVeri_YN = 0, "No", "yes")
        mHeader = mHeader + 1
        Print #1, ""
        mHeader = mHeader + 1
        Print #1, PSTR("Finance Catg", 20) & " : " & Txt(FundSource).TEXT
        mHeader = mHeader + 1
        Print #1, PSTR("Financier's Name", 20) & " : " & RstBook!FinName
        mHeader = mHeader + 1
        If UCase(left(PubComp_Name, 5)) <> "SOCIE" Then
            Print #1, Space(23) & RstBook!FinAdd1 & IIf(RstBook!FinAdd2 = "", "", ",") & RstBook!FinAdd2
             mHeader = mHeader + 1
            If RstBook!FinCity <> "" Then
                Print #1, Space(23) & RstBook!FinCity
                mHeader = mHeader + 1
            End If
        End If
        Print #1, PSTR("Financed Amount", 20) & " : " & Format(RstBook!Fin_Amt, "0.00")
        mHeader = mHeader + 1
        Print #1, PSTR("Expected Delv. Date", 20) & " : " & RstBook!EXP_DATE
        mHeader = mHeader + 1
        
        Do Until mHeader >= PageLength - mFooter
            Print #1, ""
            mHeader = mHeader + 1
        Loop
    Print #1, mChr17 & "I/We undertake to pay the price which will be ruling  on the date of delivery of the vehicle." & mChr18
    Print #1, ""
    Print #1, "Date : " & PSTR("Customer Signature", PageWidth - 7, , AlignRight)
    Print #1, ""
    Print #1, Replace(Space(PageWidth), " ", "-")
    Print #1, PRN_TIT("Terms & Conditions", "C", PageWidth, True) & mChr17
    Footer = Footer + vbLf
    j = 1
    For I = 1 To Len(Footer)
        If mID(Footer, I, 1) = vbLf Then
            Print #1, RTrim(mID(Footer, j, I - j))
            j = I + 1
        End If
     Next
    Print #1, mChr18 & RstBook!Book_UName & " " & STR(RstBook!Book_UEntDt)
    Print #1, mChr17 & Space(((PageWidth * 1.7) - Len("")) / 2) & "" & mChr18
Print #1, mEject
    Close #1
    Open "C:\RepPrint.Bat" For Output As #1
'    If fob.FolderExists("c:\WinNt") Then
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.DeviceName, ":", "") & "\Prn"
''    Else
''        Print #1, "Type C:\RepPrint.Txt >" & Replace(Printer.Port, ":", "") & "\Prn"
''    End If
'        If Len(Printer.DeviceName) > 0 Then
'            mPrinterName = "Prn"
'            If left(Printer.DeviceName, 2) = "\\" Then
'                mPrinterName = Replace(Printer.DeviceName, ":", "") & "\Prn"
'            End If
'        Else
'            MsgBox "Invalid Printer Name", vbCritical, "Printer Error"
'        End If
'    Else
'        'mPrinterName = Replace(Printer.Port, ":", "") & "\Prn"
'        mPrinterName = PubFaDosPort
'    End If
'    Print #1, "Type C:\RepPrint.Txt >" & mPrinterName
    Print #1, "Type C:\RepPrint.Txt >" & PubFaDosPort
    Close #1
    Shell "C:\RepPrint.Bat", vbHide
    Exit Sub
ELoop:
    Close #1: CheckError
    'EOF Speed Printing Section
End Sub

Private Function AddSubGroup() As Boolean
On Error GoTo ELoop
Dim SubCode$, LableName$, VehDebGrp$, GroupNature$
Dim Ad1B$, Ad2B$, Ad3B$, mCityB$, mPinB$, mPhoneB$
VehDebGrp = GCnFaV.Execute("select VehDeb_Grp from accontrols").Fields(0).Value
If VehDebGrp = "" Then
    MsgBox "Define Debtors Group in System Cotrol", vbInformation, "Validation Check"
    AddSubGroup = False
    Exit Function
End If
Ad1B = Txt(Add1)
Ad2B = Txt(Add2)
Ad3B = Txt(Add3)
mCityB = Txt(City).Tag
'mPinB = txt(Pin)
'mPhoneB = txt(Phone)
GroupNature = GCnFaV.Execute("select GroupNature from acgroup WHERE GROUPCODE = '" & VehDebGrp & "'").Fields(0).Value

    GSQL = "Select SubGroupAcCode From SubGroupCounter"
    SubCode = PubSiteCode & IIf(PubFirmCode = "", "0", PubFirmCode) & Format(G_CompCn.Execute(GSQL).Fields(0).Value, "000000")
    GSQL = "(AcID,Site_Code,SubCode,FirmCode,NamePrefix,Name,NameBiLang," & _
           "NameHelp,GroupCode,GroupNature,Nature,AliasYn," & _
           "Add1,Add2,Add3,Phone,CityCode,PANNO,FPrefix,FName," & _
           "TAdd1,TAdd2,TAdd3,TCityCode, Pin," & _
           "U_Name,U_EntDt,U_AE)" & _
           "Values ('" & SubCode & "','" & PubSiteCode & "','" & SubCode & "','" & PubFirmCode & "','" & Txt(NamePrefix) & "','" & Txt(Party_code) & "',''," & _
           "'" & Txt(Party_code) & "','" & VehDebGrp & "','" & GroupNature & "','Customer','N'," & _
           "'" & Txt(Add1) & "','" & Txt(Add2) & "','" & Txt(Add3) & "','" & Txt(Phone) & "','" & Txt(City).Tag & "','" & Txt(PAN) & "','" & Txt(FNamePrefix) & "','" & Txt(fname) & "'," & _
           "'" & Ad1B & "','" & Ad2B & "','" & Ad3B & "','" & mCityB & "', '" & Txt(PinCode) & "'," & _
           "'" & pubUName & "'," & ConvertDate(Format(Txt(VDate), "dd/MMM/yyyy HH:NN:SS")) & ",'A')"
'G_FACN.BeginTrans
GCnFaV.Execute ("insert into Subgroup " & GSQL)
GCnFaV.Execute ("insert into SubgroupAlias " & GSQL)
If PubBackEnd = "A" Then
    GCn.Execute ("insert into Subgroup " & GSQL)
    GCn.Execute ("insert into SubgroupAlias " & GSQL)
End If

G_CompCn.Execute ("Update SubGroupCounter Set SubGroupAcCode=SubGroupAcCode+1")
Txt(Party_code).Tag = SubCode
GSQL = ""
'G_FACN.CommitTrans
AddSubGroup = True
Exit Function
ELoop:
'G_FACN.RollbackTrans
AddSubGroup = False
End Function

Private Sub FillPartyDetails()  'LPS 11-03-03
    Txt(NamePrefix).TEXT = XNull(RsParty!NamePrefix)
    Txt(Party_code).TEXT = XNull(RsParty!Name)
    Txt(Party_code).Tag = XNull(RsParty!Code)
    Txt(Add1).TEXT = IIf(IsNull(RsParty!Add1), "", RsParty!Add1)
    Txt(Add2).TEXT = IIf(IsNull(RsParty!Add2), "", RsParty!Add2)
    Txt(Add3).TEXT = IIf(IsNull(RsParty!Add3), "", RsParty!Add3)
    Txt(Phone).TEXT = IIf(IsNull(RsParty!Phone), "", RsParty!Phone)
    Txt(City).Tag = IIf(IsNull(RsParty!CityCode), "", RsParty!CityCode)
    If Txt(City).Tag <> "" Then
        Txt(City).TEXT = GCn.Execute("select cityname from city where citycode = '" & Txt(City).Tag & "'").Fields(0).Value
    End If
    Txt(PAN) = IIf(IsNull(RsParty!PanNo), "", RsParty!PanNo)
    Txt(fname) = IIf(IsNull(RsParty!fname), "", RsParty!fname)
    Txt(FNamePrefix) = IIf(IsNull(RsParty!FPrefix), "", RsParty!FPrefix)

    Txt(Add1).Enabled = False
    Txt(Add2).Enabled = False
    Txt(Add3).Enabled = False
    Txt(Phone).Enabled = False
    Txt(City).Enabled = False
    Txt(PAN).Enabled = False
    Txt(fname).Enabled = False
    Txt(FNamePrefix).Enabled = False
End Sub
Private Function TxtGridLeave(Optional Index As Integer, Optional ValidateCall As Boolean) As Boolean
Select Case FGrid1.Col
Dim I%
    Case Col_Accessory
        If RsAcces.RecordCount = 0 Or (RsAcces.EOF = True Or RsAcces.BOF = True) Or TxtGrid(0).TEXT = "" Then
            FGrid1.TextMatrix(FGrid1.Row, Col_Accessory) = ""
            FGrid1.TextMatrix(FGrid1.Row, Col_AccCode) = ""
        Else
            For I = 1 To FGrid1.Rows - 1
                If FGrid1.Row <> I Then
                    If FGrid1.TextMatrix(I, Col_AccCode) = RsAcces!Code Then
                        MsgBox "Duplicate Item Not Allowed", vbInformation, App.Title: TxtGridLeave = False: Exit Function
                    End If
                End If
            Next
            FGrid1.TextMatrix(FGrid1.Row, Col_Accessory) = RsAcces!Name
            FGrid1.TextMatrix(FGrid1.Row, Col_AccCode) = RsAcces!Code
        End If
        If FGrid1.TextMatrix(FGrid1.Rows - 1, 1) <> "" Then FGrid1.AddItem FGrid1.Rows
        
    Case Col_Qty
        FGrid1.TextMatrix(FGrid1.Row, Col_Qty) = Format(Val(TxtGrid(0).TEXT), "0.00")
        If FGrid1.TextMatrix(FGrid1.Rows - 1, 1) <> "" Then FGrid1.AddItem FGrid1.Rows
End Select
TxtGridLeave = True
If ValidateCall = False Then
    FGrid1.SetFocus
    TxtGrid(0).Visible = False
End If
End Function
