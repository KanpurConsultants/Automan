VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConvertData2 
   BackColor       =   &H00CFE2D9&
   Caption         =   "Import Data (From Siebel to Automan)"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14265
   Icon            =   "frmConvertData2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   14265
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   31
      Left            =   1530
      TabIndex        =   197
      Text            =   "ACAE_Spr"
      Top             =   6855
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Ledger A/c (Ws/Spr)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   31
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   196
      Top             =   6840
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   29
      Left            =   8340
      TabIndex        =   182
      Text            =   "CLOSE"
      Top             =   6510
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "JobCard Close"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   29
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   181
      Top             =   6495
      Width           =   1485
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "JobCard Requisition"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   30
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   192
      Top             =   8265
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   30
      Left            =   8205
      TabIndex        =   191
      Text            =   "Req"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   345
      Left            =   10665
      TabIndex        =   190
      Top             =   1050
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Update History Name"
      Height          =   405
      Left            =   8790
      TabIndex        =   189
      Top             =   1035
      Width           =   1815
   End
   Begin VB.TextBox TxtCentralData 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1545
      TabIndex        =   186
      Top             =   330
      Width           =   1380
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "JobCard Labour"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   28
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   177
      Top             =   6195
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   28
      Left            =   8340
      TabIndex        =   176
      Text            =   "LAB"
      Top             =   6210
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   27
      Left            =   8340
      TabIndex        =   172
      Text            =   "JOB"
      Top             =   5910
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "JobCard Entry"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   27
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   171
      Top             =   5895
      Width           =   1485
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Service Booking"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   26
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   5595
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   26
      Left            =   8340
      TabIndex        =   166
      Text            =   "SR"
      Top             =   5610
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   25
      Left            =   8340
      TabIndex        =   162
      Text            =   "CCCODE"
      Top             =   5310
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Trouble Master"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   25
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   161
      Top             =   5295
      Width           =   1485
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Labour Master"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   24
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   4995
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   24
      Left            =   8340
      TabIndex        =   153
      Text            =   "CJCODE"
      Top             =   5010
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   23
      Left            =   8340
      TabIndex        =   149
      Text            =   "SAE"
      Top             =   4710
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Stock Adjustment"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   23
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   4695
      Width           =   1485
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Stock Trf. (Outward)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   22
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   144
      Top             =   4395
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   22
      Left            =   8340
      TabIndex        =   143
      Text            =   "STE"
      Top             =   4410
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   21
      Left            =   8340
      TabIndex        =   139
      Text            =   "OTC"
      Top             =   4110
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Spare Sales Bill"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   21
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   4095
      Width           =   1485
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Stock Trf. (Inward)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   20
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   134
      Top             =   3795
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   20
      Left            =   8340
      TabIndex        =   133
      Text            =   "STE"
      Top             =   3810
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   19
      Left            =   8340
      TabIndex        =   129
      Text            =   "TM"
      Top             =   3510
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "PurchBill - Telco"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   19
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   3495
      Width           =   1485
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "MRN - Telco"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   18
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   3195
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   18
      Left            =   8340
      TabIndex        =   123
      Text            =   "TM"
      Top             =   3210
      Width           =   1740
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Repair Veh. Stock"
      Height          =   405
      Left            =   8805
      TabIndex        =   122
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Vehicle Booking -PCD"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   17
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   5925
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   17
      Left            =   1530
      TabIndex        =   117
      Text            =   "ORDMIS"
      Top             =   5940
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   16
      Left            =   8340
      TabIndex        =   113
      Text            =   "Local"
      Top             =   2910
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "MRN - Local Purch."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   16
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   2895
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   15
      Left            =   8340
      TabIndex        =   107
      Text            =   "PART"
      Top             =   2310
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Part Master"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   14
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   2595
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   14
      Left            =   8340
      TabIndex        =   102
      Text            =   "PART"
      Top             =   2610
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   13
      Left            =   1530
      TabIndex        =   98
      Text            =   "RCAE"
      Top             =   6540
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Customer Receipt"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   13
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   6525
      Width           =   1485
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Vehicle Sales/Del."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   6225
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   12
      Left            =   1530
      TabIndex        =   92
      Text            =   "INVAE"
      Top             =   6240
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   11
      Left            =   1530
      TabIndex        =   88
      Text            =   "ORDTRN"
      Top             =   5640
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Vehicle Booking -CVD"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   5625
      Width           =   1485
   End
   Begin VB.TextBox txtFirm 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1890
      TabIndex        =   83
      Text            =   "4"
      Top             =   1245
      Width           =   435
   End
   Begin VB.TextBox txtDiv 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1890
      TabIndex        =   82
      Top             =   990
      Width           =   435
   End
   Begin VB.TextBox txtSite 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1890
      TabIndex        =   81
      Text            =   "2"
      Top             =   720
      Width           =   435
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Ledger A/c"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   5325
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   1530
      TabIndex        =   72
      Text            =   "ACAE"
      Top             =   5340
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   1530
      TabIndex        =   68
      Text            =   "PURAE"
      Top             =   5040
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Vehicle Purchase"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   5025
      Width           =   1485
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Financer/Bank"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   4725
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   1530
      TabIndex        =   62
      Text            =   "Order.xls"
      Top             =   4740
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   1530
      TabIndex        =   58
      Text            =   "PURAE"
      Top             =   4440
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Godown/Location"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   4425
      Width           =   1485
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Model"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   4125
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   1530
      TabIndex        =   52
      Text            =   "MODEL"
      Top             =   4140
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   1530
      TabIndex        =   48
      Text            =   "INVMIS"
      Top             =   3840
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Purpose/Inteded"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3825
      Width           =   1485
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Inquiry/Ref.  Source"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3525
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   1530
      TabIndex        =   42
      Text            =   "OPPTRN"
      Top             =   3540
      Width           =   1740
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   1530
      TabIndex        =   38
      Text            =   "EmpMast"
      Top             =   3240
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Sales Person"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3225
      Width           =   1485
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2925
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1530
      TabIndex        =   32
      Text            =   "ACTRN"
      Top             =   2940
      Width           =   1740
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   21
      Left            =   13380
      TabIndex        =   31
      Text            =   "Text2"
      Top             =   15
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   6330
      TabIndex        =   24
      Top             =   1530
      Width           =   2475
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   6330
      TabIndex        =   23
      Top             =   75
      Width           =   2475
   End
   Begin VB.OptionButton optAuto 
      Caption         =   "All Import"
      Height          =   225
      Index           =   1
      Left            =   4515
      TabIndex        =   22
      Top             =   930
      Width           =   1785
   End
   Begin VB.OptionButton optAuto 
      Caption         =   "Selective Import"
      Height          =   225
      Index           =   0
      Left            =   4515
      TabIndex        =   21
      Top             =   690
      Value           =   -1  'True
      Width           =   1785
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1530
      TabIndex        =   20
      Text            =   "MODEL"
      Top             =   2640
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Colour "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2625
      Width           =   1485
   End
   Begin VB.TextBox ImportTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1530
      TabIndex        =   18
      Text            =   "ACAE"
      Top             =   2340
      Width           =   1740
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2325
      Width           =   1485
   End
   Begin VB.CommandButton CmdConvert 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Start All Import"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4515
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1170
      Width           =   1785
   End
   Begin VB.CommandButton CmdBlank 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Make Blank Print File"
      Height          =   255
      Left            =   12045
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   690
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.CommandButton CmdPrn 
      BackColor       =   &H00D7F0EE&
      Caption         =   "Print"
      Height          =   345
      Left            =   12300
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   315
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.CommandButton CmdPath 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Apply Parameter"
      Height          =   345
      Left            =   4515
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   330
      Width           =   1785
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Exit"
      Height          =   345
      Left            =   8805
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Width           =   1560
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   240
      Left            =   1545
      TabIndex        =   0
      Text            =   "D:\VishalJain\DOSData\PCD0506"
      Top             =   75
      Width           =   4755
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11790
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton ImportBtn 
      BackColor       =   &H00CAFDFD&
      Caption         =   "Unit Master"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   15
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   2295
      Width           =   1485
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   31
      Left            =   5580
      TabIndex        =   200
      Top             =   6855
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   31
      Left            =   4425
      TabIndex        =   199
      Top             =   6855
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   31
      Left            =   3270
      TabIndex        =   198
      Top             =   6855
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   29
      Left            =   12405
      TabIndex        =   185
      Top             =   6510
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   29
      Left            =   11250
      TabIndex        =   184
      Top             =   6510
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   10095
      TabIndex        =   183
      Top             =   6510
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   9960
      TabIndex        =   195
      Top             =   8280
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   30
      Left            =   11115
      TabIndex        =   194
      Top             =   8280
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   30
      Left            =   12270
      TabIndex        =   193
      Top             =   8280
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(e.g. Auto_01)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   19
      Left            =   3000
      TabIndex        =   188
      Top             =   330
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Central Data Dir :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   187
      Top             =   330
      Width           =   1215
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   10095
      TabIndex        =   180
      Top             =   6195
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   28
      Left            =   11250
      TabIndex        =   179
      Top             =   6195
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   28
      Left            =   12405
      TabIndex        =   178
      Top             =   6195
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   27
      Left            =   12405
      TabIndex        =   175
      Top             =   5895
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   27
      Left            =   11250
      TabIndex        =   174
      Top             =   5895
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   10095
      TabIndex        =   173
      Top             =   5895
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   10095
      TabIndex        =   170
      Top             =   5595
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   26
      Left            =   11250
      TabIndex        =   169
      Top             =   5595
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   26
      Left            =   12405
      TabIndex        =   168
      Top             =   5595
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   25
      Left            =   12405
      TabIndex        =   165
      Top             =   5295
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   25
      Left            =   11250
      TabIndex        =   164
      Top             =   5310
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   10095
      TabIndex        =   163
      Top             =   5310
      Width           =   1170
   End
   Begin VB.Label LBLCNT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Processed Rec"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   18
      Left            =   11250
      TabIndex        =   160
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Label LBLCNT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Error Rec"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   17
      Left            =   12405
      TabIndex        =   159
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Label LBLCNT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rec"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   16
      Left            =   10095
      TabIndex        =   158
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   10110
      X2              =   13575
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   10095
      TabIndex        =   157
      Top             =   5010
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   24
      Left            =   11250
      TabIndex        =   156
      Top             =   5010
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   24
      Left            =   12405
      TabIndex        =   155
      Top             =   5010
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   23
      Left            =   12405
      TabIndex        =   152
      Top             =   4710
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   23
      Left            =   11250
      TabIndex        =   151
      Top             =   4710
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   10095
      TabIndex        =   150
      Top             =   4710
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   10095
      TabIndex        =   147
      Top             =   4395
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   22
      Left            =   11250
      TabIndex        =   146
      Top             =   4410
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   22
      Left            =   12405
      TabIndex        =   145
      Top             =   4410
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   21
      Left            =   12405
      TabIndex        =   142
      Top             =   4110
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   21
      Left            =   11250
      TabIndex        =   141
      Top             =   4110
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   10095
      TabIndex        =   140
      Top             =   4110
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   10095
      TabIndex        =   137
      Top             =   3810
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   20
      Left            =   11250
      TabIndex        =   136
      Top             =   3810
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   20
      Left            =   12405
      TabIndex        =   135
      Top             =   3810
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   19
      Left            =   12405
      TabIndex        =   132
      Top             =   3510
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   19
      Left            =   11250
      TabIndex        =   131
      Top             =   3510
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   10095
      TabIndex        =   130
      Top             =   3510
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   10095
      TabIndex        =   127
      Top             =   3210
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   18
      Left            =   11250
      TabIndex        =   126
      Top             =   3210
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   18
      Left            =   12405
      TabIndex        =   125
      Top             =   3210
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   3285
      TabIndex        =   121
      Top             =   5940
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   17
      Left            =   4440
      TabIndex        =   120
      Top             =   5940
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   17
      Left            =   5595
      TabIndex        =   119
      Top             =   5940
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   16
      Left            =   12405
      TabIndex        =   116
      Top             =   2910
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   16
      Left            =   11250
      TabIndex        =   115
      Top             =   2910
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   10095
      TabIndex        =   114
      Top             =   2910
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   10095
      TabIndex        =   111
      Top             =   2310
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   15
      Left            =   11250
      TabIndex        =   110
      Top             =   2310
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   15
      Left            =   12420
      TabIndex        =   109
      Top             =   2310
      Width           =   1140
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   10095
      TabIndex        =   106
      Top             =   2610
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   14
      Left            =   11250
      TabIndex        =   105
      Top             =   2610
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   14
      Left            =   12405
      TabIndex        =   104
      Top             =   2610
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   13
      Left            =   5595
      TabIndex        =   101
      Top             =   6555
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   13
      Left            =   4440
      TabIndex        =   100
      Top             =   6540
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   3285
      TabIndex        =   99
      Top             =   6540
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   3285
      TabIndex        =   96
      Top             =   6240
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   12
      Left            =   4440
      TabIndex        =   95
      Top             =   6240
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   12
      Left            =   5595
      TabIndex        =   94
      Top             =   6240
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   11
      Left            =   5595
      TabIndex        =   91
      Top             =   5640
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   11
      Left            =   4440
      TabIndex        =   90
      Top             =   5640
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   3285
      TabIndex        =   89
      Top             =   5640
      Width           =   1170
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(C-CVD/P-PCD)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   15
      Left            =   2355
      TabIndex        =   86
      Top             =   1020
      Width           =   1260
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(1-Dhule/2-Jalgaon)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   14
      Left            =   2355
      TabIndex        =   85
      Top             =   750
      Width           =   1575
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(2-Ujwal Auto Pvt. Ltd.)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   7
      Left            =   2355
      TabIndex        =   84
      Top             =   1275
      Width           =   1905
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Division Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   5
      Left            =   45
      TabIndex        =   80
      Top             =   1020
      Width           =   1770
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Site Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   4
      Left            =   45
      TabIndex        =   79
      Top             =   750
      Width           =   1770
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Division"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   3
      Left            =   -240
      TabIndex        =   78
      Top             =   645
      Width           =   135
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Firm Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   2
      Left            =   45
      TabIndex        =   77
      Top             =   1275
      Width           =   1770
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   3285
      TabIndex        =   76
      Top             =   5340
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   10
      Left            =   4440
      TabIndex        =   75
      Top             =   5340
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   10
      Left            =   5595
      TabIndex        =   74
      Top             =   5340
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   9
      Left            =   5595
      TabIndex        =   71
      Top             =   5040
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   9
      Left            =   4440
      TabIndex        =   70
      Top             =   5040
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   3285
      TabIndex        =   69
      Top             =   5040
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   3285
      TabIndex        =   66
      Top             =   4740
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   8
      Left            =   4440
      TabIndex        =   65
      Top             =   4740
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   8
      Left            =   5595
      TabIndex        =   64
      Top             =   4755
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   7
      Left            =   5595
      TabIndex        =   61
      Top             =   4440
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   7
      Left            =   4440
      TabIndex        =   60
      Top             =   4440
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   3285
      TabIndex        =   59
      Top             =   4440
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3285
      TabIndex        =   56
      Top             =   4140
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   6
      Left            =   4440
      TabIndex        =   55
      Top             =   4140
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   6
      Left            =   5595
      TabIndex        =   54
      Top             =   4140
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   5
      Left            =   5595
      TabIndex        =   51
      Top             =   3840
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   50
      Top             =   3840
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3285
      TabIndex        =   49
      Top             =   3840
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3285
      TabIndex        =   46
      Top             =   3540
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   45
      Top             =   3540
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   4
      Left            =   5595
      TabIndex        =   44
      Top             =   3540
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   3
      Left            =   5595
      TabIndex        =   41
      Top             =   3240
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   40
      Top             =   3240
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3285
      TabIndex        =   39
      Top             =   3240
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3285
      TabIndex        =   36
      Top             =   2940
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   2
      Left            =   4440
      TabIndex        =   35
      Top             =   2940
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   2
      Left            =   5595
      TabIndex        =   34
      Top             =   2940
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   5595
      TabIndex        =   30
      Top             =   2640
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   29
      Top             =   2640
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3285
      TabIndex        =   28
      Top             =   2640
      Width           =   1170
   End
   Begin VB.Label lblRecError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   5595
      TabIndex        =   27
      Top             =   2340
      Width           =   1170
   End
   Begin VB.Label lblRecCopy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   26
      Top             =   2340
      Width           =   1170
   End
   Begin VB.Label lblRecTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3285
      TabIndex        =   25
      Top             =   2340
      Width           =   1170
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3300
      X2              =   6765
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Import Sucessfull"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   13
      Left            =   510
      TabIndex        =   16
      Top             =   2055
      Width           =   1395
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Import In-Process/Failed"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   12
      Left            =   510
      TabIndex        =   15
      Top             =   1815
      Width           =   1965
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Import not started"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   11
      Left            =   510
      TabIndex        =   14
      Top             =   1575
      Width           =   1440
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   10
      Left            =   75
      TabIndex        =   13
      Top             =   1815
      Width           =   390
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   9
      Left            =   75
      TabIndex        =   12
      Top             =   2055
      Width           =   390
   End
   Begin VB.Label LBLCNT 
      Appearance      =   0  'Flat
      BackColor       =   &H00CAFDFD&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   8
      Left            =   75
      TabIndex        =   11
      Top             =   1575
      Width           =   390
   End
   Begin VB.Label LBLCNT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rec"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   3285
      TabIndex        =   10
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excel File Folder :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   9
      Top             =   75
      Width           =   1245
   End
   Begin VB.Label LBLCNT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Error Rec"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   6
      Left            =   5595
      TabIndex        =   5
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Label LBLCNT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Processed Rec"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   4
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Label LblPrefix 
      Caption         =   "LblPrefix"
      Height          =   300
      Left            =   12600
      TabIndex        =   8
      Top             =   15
      Visible         =   0   'False
      Width           =   765
   End
End
Attribute VB_Name = "frmConvertData2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ErrorGCN As adodb.Connection, ExcelGcn1 As adodb.Connection, ExcelGcn2 As adodb.Connection

Dim OldGCn As adodb.Connection

Dim varScrptObj As New Scripting.FileSystemObject, varTxtstrm As Scripting.TextStream

Dim Master As adodb.Recordset, Master1 As adodb.Recordset
Dim RsNew As adodb.Recordset, RsNew1 As adodb.Recordset
Dim ChkDup As adodb.Recordset, rsTemp As adodb.Recordset

Dim Fob As New FileSystemObject, OpenCon As Boolean, CodeCnt As Variant

Private Const mPubDivision As Byte = 0
Private Const mSite As Byte = 1
Private Const mB_Code As Byte = 2
Private Const JobType As Byte = 3
Private Const JobNo As Byte = 21
Private Const SerialNo As Byte = 21

Private Const ReqTypeG As String = "W_RG"
Private Const ReqTypeW As String = "W_RW"

Dim DataPath$, FADataPath$, PubCenCompCode$, GSQL$
Dim tmpDb As DAO.Database
Dim db As DAO.Database
Private Const FinishColor As String = &HC0C0FF
Private Const ProcessColor As String = &HFFC0FF

Dim CopyCnt As Long, ErrorCnt As Long
Dim Comp_Path$, App_Path$, TxtVehPath$, mVType$
Dim VoucherEditFlag As Boolean, ConvertAll As Boolean
Private Const MsgNoRecToImport As String = "There is no records in import Database table !! "
Private Const TitleNoRec As String = "No Records!"
Private Const MsgUpdDone As String = "Updation Completed !!"
Private Const TitleUpdDone As String = "Job Finished!"

Private Sub CmdBlank_Click()
GCn.Execute ("Delete * from PrnMissRec")
MsgBox "Job Completed", vbInformation
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub



Private Sub ConvertModel()
'MODEL_CAT --- > ModelCat_Code   Site_Code   ModelCat_Name   U_Name  U_EntDt U_AE    Trf_Date
'MODEL_CTG-->GSLNO Name


'MODEL_GRP --- >ModelGrp_Code   Site_Code   ModelGrp_Name   ModelCat_Code   ModelDiv_Code   Wheel_Catg  U_Name  U_EntDt U_AE    Trf_Date
'MODELGRP --- >CODE    NAME    MODTEMP

'Vehicle_Type-->Vehicle_Type

'AMDMODL--> MODEL,Site_Code***,CHAS_TYPE,VEH_TYPE,MODEL_TYPE,
'MODEL_IND,Sales_Desc***,MODEL_DESC,MODEL_DES1,Model_Desc2**,
'GRP_CODE,CTG_SLNO,Div_Code***,RIMS**,Active_YN***
'STAT_IND,TYRES,TYRE_F,TYRE_M,TYRE_R
'TYRE_FS,TYRE_MS,TYRE_rsDos,RIMS,SEAT
'RLW,HORSEPOWER,FRONT_A_WT,REAR_A_WT,UNLADEN_WT,
'GROSS_WT,WHEELBASE,CYLINDER,FUEL,TRADE_NO,MANUFACTUR

'Model , Site_Code, Chas_Type, Vehicle_Type, Model_Type
'Model_Ind,Sales_Desc,Model_Desc,Model_Desc1,Model_Desc2
'Grp_Code,Cat_Code,Div_Code,Wheel_Catg,Active_YN,
'STAT_IND,TYRES,TYRE_F,TYRE_M,TYRE_R,
'TYRE_FS,TYRE_MS,TYRE_rsDos,RIMS,SEAT
'RLW,HORSEPOWER,FRONT_A_WT,REAR_A_WT,UNLADEN_WT
'GROSS_WT,WHEELBASE,CYLINDER,FUEL,TRADE_NO,Manufacturer
'U_Name  U_EntDt U_AE

'' On Error GoTo Eloop
Dim rsDos As adodb.Recordset
Dim Data1 As String
Dim Data2 As String
Dim Data3 As String
Dim Data4 As String
Dim Data5 As String
Dim mRims As String
Dim Cnt As Long
Dim Cnt1 As Long
Dim ColField As String
Dim ColName As String
'If OpenCon = False Then MsgBox "Give DataPath for importing database", vbInformation, "Blank Path": Text1.SetFocus: Exit Sub

'***** MODEL CATG
Cnt = 0
Cnt1 = 0
Set rsDos = New adodb.Recordset
rsDos.CursorLocation = adUseClient
rsDos.Open "select * from MODELCTG", OldGCn, adOpenDynamic, adLockOptimistic
If rsDos.RecordCount = 0 Then GoTo nxt1
LBLCNT(2).Caption = rsDos.RecordCount
LBLCNT(2).Refresh
LBLCNT(5).Caption = "Model Catg"
LBLCNT(5).Refresh
GCn.BeginTrans
GCn.Execute ("Delete * from MODEL_CAT")
Do Until rsDos.EOF
ColField = Replace(IIf(IsNull(rsDos!SLNo), "", rsDos!SLNo), "'", "`")
ColName = "Model Ctag"
Data1 = Replace(IIf(IsNull(rsDos!Name), "", rsDos!Name), "'", "`")
GCn.Execute "insert into MODEL_CAT( " & _
"ModelCat_Code,Site_Code,ModelCat_Name,U_Name,U_EntDt,U_AE) " & _
"values(" & _
"'" & rsDos!SLNo & "','" & PubSiteCode & "','" & Data1 & "','SA'," & ConvertDate(PubLoginDate) & ",'A')"
rsDos.MoveNext
LBLCNT(3).Caption = Cnt
LBLCNT(3).Refresh
Cnt = Cnt + 1
Loop
GCn.CommitTrans
nxt1:
'***** MODEL GRP
Cnt = 0
Cnt1 = 0
Set rsDos = New adodb.Recordset
rsDos.CursorLocation = adUseClient
rsDos.Open "select * from MODELGRP", OldGCn, adOpenDynamic, adLockOptimistic
If rsDos.RecordCount = 0 Then GoTo nxt2
    LBLCNT(2).Caption = rsDos.RecordCount
    LBLCNT(2).Refresh
    LBLCNT(5).Caption = "Model Group"
    LBLCNT(5).Refresh
    GCn.BeginTrans
    GCn.Execute ("Delete * from MODEL_GRP")
Do Until rsDos.EOF
ColField = Replace(IIf(IsNull(rsDos!Code), "", rsDos!Code), "'", "`")
ColName = "ModelGrp_Code"
Data2 = GCn.Execute("select modelcat_code from model_cat").Fields(0).Value
Data1 = Replace(IIf(IsNull(rsDos!Name), "", rsDos!Name), "'", "`")
GCn.Execute "insert into MODEL_GRP( " & _
"ModelGrp_Code,Site_Code,ModelGrp_Name,ModelCat_Code,ModelDiv_Code," & _
"Wheel_Catg,U_Name,U_EntDt,U_AE) " & _
"values(" & _
"'" & rsDos!Code & "','" & PubSiteCode & "','" & Data1 & "','" & Data2 & "','" & PubDivCode & "'," & _
"'Six','SA'," & ConvertDate(PubLoginDate) & ",'A')"
rsDos.MoveNext
LBLCNT(3).Caption = Cnt
LBLCNT(3).Refresh
Cnt = Cnt + 1
Loop
GCn.CommitTrans
nxt2:

'***** Vehicle Type
Cnt = 0
Cnt1 = 0
Set rsDos = New adodb.Recordset
rsDos.CursorLocation = adUseClient
rsDos.Open "select distinct VEH_TYPE from AMDMODL", OldGCn, adOpenDynamic, adLockOptimistic
If rsDos.RecordCount = 0 Then GoTo nxt3
    LBLCNT(2).Caption = rsDos.RecordCount
    LBLCNT(2).Refresh
    LBLCNT(5).Caption = "Vehicle Type"
    LBLCNT(5).Refresh
    GCn.BeginTrans
    GCn.Execute ("Delete * from Vehicle_Type")
Do Until rsDos.EOF
Data1 = Replace(IIf(IsNull(rsDos!Veh_Type), "", rsDos!Veh_Type), "'", "`")
GCn.Execute "insert into Vehicle_Type(Vehicle_Type) values('" & Data1 & "')"
rsDos.MoveNext
LBLCNT(3).Caption = Cnt
LBLCNT(3).Refresh
Cnt = Cnt + 1
Loop
GCn.CommitTrans
nxt3:
'***** Model master
Cnt = 0
Cnt1 = 0
Set rsDos = New adodb.Recordset
rsDos.CursorLocation = adUseClient
rsDos.Open "select * from AMDMODL", OldGCn, adOpenDynamic, adLockOptimistic
If rsDos.RecordCount = 0 Then MsgBox MsgNoRecToImport & "!! Model !!", vbInformation, TitleNoRec: Exit Sub
If GCn.Execute("select count(MODEL) from MODEL").Fields(0).Value > 0 Then
    If MsgBox("There are records in Model Master !! Do you want to continue ? ", vbYesNo + vbCritical + vbDefaultButton2, "Update Database !") = vbNo Then Exit Sub
End If
LBLCNT(2).Caption = rsDos.RecordCount
LBLCNT(2).Refresh
LBLCNT(5).Caption = "Model"
LBLCNT(5).Refresh
GCn.BeginTrans
GCn.Execute ("Delete * from Model")
Do Until rsDos.EOF
ColField = Replace(IIf(IsNull(rsDos!Model), "", rsDos!Model), "'", "`")
ColName = "Model"

Data1 = Replace(IIf(IsNull(rsDos!Model), "", rsDos!Model), "'", "`")
Data2 = Replace(IIf(IsNull(rsDos!Model_Desc), "", rsDos!Model_Desc), "'", "`")
Data3 = Replace(IIf(IsNull(rsDos!Model_Des1), "", rsDos!Model_Des1), "'", "`")
mRims = "Above Ten"
Select Case IIf(IsNull(rsDos!Rims), 0, rsDos!Rims)
    Case 3: mRims = "Two": Case 5: mRims = "Four": Case 7: mRims = "Six": Case 9: mRims = "Eight": Case 11: mRims = "Ten"
End Select
GCn.Execute "insert into Model( " & _
"Model , Site_Code, Chas_Type, Vehicle_Type, Model_Type," & _
"Model_Ind,Sales_Desc,Model_Desc,Model_Desc1,Model_Desc2," & _
"Grp_Code,Cat_Code,Div_Code,Wheel_Catg,Active_YN," & _
"STAT_IND,TYRES,TYRE_F,TYRE_M,TYRE_R," & _
"TYRE_FS,TYRE_MS,TYRE_rsDos,RIMS,SEAT," & _
"RLW,HORSEPOWER,FRONT_A_WT,REAR_A_WT,UNLADEN_WT," & _
"GROSS_WT,WHEELBASE,CYLINDER,FUEL,TRADE_NO," & _
"Manufacturer,U_Name,U_EntDt,U_AE)" & _
"values( " & _
"'" & Data1 & "','" & PubSiteCode & " ','" & IIf(IsNull(rsDos!Chas_Type), "", rsDos!Chas_Type) & "','" & IIf(IsNull(rsDos!Veh_Type), "", rsDos!Veh_Type) & "','" & IIf(IsNull(rsDos!Model_Type), "XX", rsDos!Model_Type) & "'," & _
"" & IIf(IsNull(rsDos!Model_Ind), 0, rsDos!Model_Ind) & ",'" & left(Data2, 40) & "','" & Data2 & "','" & Data3 & "','" & Data3 & "'," & _
"'" & IIf(IsNull(rsDos!Grp_Code), "1", rsDos!Grp_Code) & "' ,'" & IIf(IsNull(rsDos!CTG_SLNO), "", rsDos!CTG_SLNO) & "','" & PubDivCode & "','" & mRims & "',0," & _
"" & IIf(IsNull(rsDos!STAT_IND), 0, rsDos!STAT_IND) & "," & IIf(IsNull(rsDos!Tyres), 0, rsDos!Tyres) & "," & IIf(IsNull(rsDos!Tyre_F), 0, rsDos!Tyre_F) & "," & IIf(IsNull(rsDos!Tyre_M), 0, rsDos!Tyre_M) & "," & IIf(IsNull(rsDos!Tyre_R), 0, rsDos!Tyre_R) & "," & _
"'" & IIf(IsNull(rsDos!Tyre_FS), "", rsDos!Tyre_FS) & "','" & IIf(IsNull(rsDos!Tyre_MS), "", rsDos!Tyre_MS) & "','" & IIf(IsNull(rsDos!Tyre_rsDos), "", rsDos!Tyre_rsDos) & "'," & rsDos!Rims & "," & rsDos!Seat & "," & _
"'" & IIf(IsNull(rsDos!RLW), "", rsDos!RLW) & "','" & IIf(IsNull(rsDos!HorsePower), "", rsDos!HorsePower) & "','" & IIf(IsNull(rsDos!Front_A_Wt), "", rsDos!Front_A_Wt) & "','" & IIf(IsNull(rsDos!Rear_A_Wt), "", rsDos!Rear_A_Wt) & "','" & IIf(IsNull(rsDos!Unladen_Wt), "", rsDos!Unladen_Wt) & "'," & _
"'" & IIf(IsNull(rsDos!Gross_Wt), "", rsDos!Gross_Wt) & "'," & IIf(IsNull(rsDos!WheelBase), 0, rsDos!WheelBase) & "," & IIf(IsNull(rsDos!Cylinder), 0, rsDos!Cylinder) & ",'" & IIf(IsNull(rsDos!Fuel), "", rsDos!Fuel) & "','" & IIf(IsNull(rsDos!Trade_NO), "", rsDos!Trade_NO) & "'," & _
"'" & IIf(IsNull(rsDos!MANUFACTUR), "", rsDos!MANUFACTUR) & "','SA'," & ConvertDate(PubLoginDate) & ",'A')"

rsDos.MoveNext
LBLCNT(3).Caption = Cnt
LBLCNT(3).Refresh
Cnt = Cnt + 1
Loop
GCn.CommitTrans
Set rsDos = Nothing
MsgBox "Updation Complete For !! Model  !!", vbInformation, TitleUpdDone
Exit Sub
Eloop:
If err.NUMBER = -2147467259 Then
    Cnt1 = Cnt1 + 1: LBLCNT(7).Caption = Cnt1: LBLCNT(7).Refresh
    GCn.Execute ("insert into prnmissrec(code,colname,details) values('" & ColField & "','" & ColName & "','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
Else
    GCn.Execute ("insert into prnmissrec(code,colname,details) values('" & ColField & "','" & ColName & "','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End If
End Sub

Private Sub CmdDesig_Click()
'Conversion not required
End Sub

Private Sub CmdDiscFactor_Click()
'Conversion Not Required
End Sub

Private Sub CmdInspCatg_Click()
'Conversion Not Required
End Sub

Private Sub CmdInspElem_Click()
'Conversion not required
End Sub


Private Sub CmdPath_Click()
'' On Error GoTo Eloop
Dim mWinCN As Boolean
Dim mLine As String
Dim mCompany$
        Text1.Text = Dir1.Path
        
        If Text1 = "" Then MsgBox "Give Excel files Path", vbInformation, "Blank Path": Dir1.SetFocus: Exit Sub
        If TxtCentralData.Text = "" Then MsgBox "Central Data Directory Path is empty", vbCritical, "Blank Path": TxtCentralData.SetFocus: Exit Sub
        'Windows Data
        mWinCN = True
        
        
        
        
        

        
        
        
        
        If Fob.FileExists("c:\Automan.ini") = False Then MsgBox "S/W ini File Missing, Contact to System Administrator", vbCritical: End
        Set varTxtstrm = varScrptObj.OpenTextFile("c:\Automan.ini")
        
        varTxtstrm.SkipLine
        varTxtstrm.Skip 2: Pub_DataPath = varTxtstrm.ReadLine
        varTxtstrm.Skip 2: PubRepoPath = varTxtstrm.ReadLine
        varTxtstrm.Skip 2: PubBkpPath = varTxtstrm.ReadLine
        
        
        Set varTxtstrm = varScrptObj.OpenTextFile("c:\Automan.ini")
        
        Do Until varTxtstrm.AtEndOfLine
            mLine = varTxtstrm.ReadLine
                            
            If UTrim(left(mLine, 10)) = "SQLSERVER=" Then
                PubServerName = Mid(mLine, 11, Len(mLine) - 10)
            ElseIf UTrim(left(mLine, 8)) = "COMPANY=" Then
                mCompany = Mid(mLine, 9, Len(mLine) - 8)
            ElseIf UTrim(left(mLine, 5)) = "DATA=" Then
                Pub_DataPath = Mid(mLine, 6, Len(mLine) - 5)
            ElseIf UTrim(left(mLine, 8)) = "REPORTS=" Then
                PubRepoPath = Mid(mLine, 9, Len(mLine) - 8)
            ElseIf UTrim(left(mLine, 7)) = "BACKUP=" Then
                PubBkpPath = Mid(mLine, 7, Len(mLine) - 6)
            End If
            
        Loop
                                           
                    
        PubBackEnd = IIf(Trim(PubServerName) = "", "A", "S")
        
        
                               
        
                                                       
'        If Fob.FileExists("c:\Automan.ini") = False Then MsgBox "Automan.INI File is Missing, Contact to System Administrator", vbCritical: End
'        Set varTxtstrm = varScrptObj.OpenTextFile("c:\Automan.ini")
'        varTxtstrm.SkipLine
'        varTxtstrm.Skip 2: Pub_DataPath = varTxtstrm.ReadLine
'        varTxtstrm.Skip 2: PubRepoPath = varTxtstrm.ReadLine
'        varTxtstrm.Skip 2: PubBkpPath = varTxtstrm.ReadLine
         
         
         
         
         
         
        If PubBackEnd = "A" Then
            Comp_Path = Pub_DataPath & "\Company.mdb"
            Set GCnComp = New Connection
            With GCnComp
                .CursorLocation = adUseClient
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .ConnectionString = "Data Source=" & Comp_Path & ";Persist Security Info=False"
                .Open
            End With
        ElseIf PubBackEnd = "S" Then
            Comp_Path = "Company"
            Set GCnComp = New Connection
            With GCnComp
                .CursorLocation = adUseClient
                If mCompany = "" Then
                    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & Comp_Path & ";Data Source=" & PubServerName
                Else
                    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & mCompany & ";Data Source=" & PubServerName
                End If
                .Open
            End With
        End If
        'TxtCentralData = GCnComp.Execute("SELECT CentralData_Path FROM Company").Fields(0).Value
        
        
        
        
        
        If PubBackEnd = "A" Then
            DataPath = Pub_DataPath & "\" & TxtCentralData & "\Automan.mdb"
            App_Path = Pub_DataPath & "\" & TxtCentralData & "\Automan.mdb"
        ElseIf PubBackEnd = "S" Then
            DataPath = TxtCentralData
            App_Path = TxtCentralData
        End If
        
        
        If PubBackEnd = "A" Then
            If Fob.FileExists(DataPath) = False Then MsgBox "Automan.MDB File is Missing or Central Data Directory Path is invalid", vbCritical: End
            
            Set GCn = New adodb.Connection
            With GCn
                .CursorLocation = adUseClient
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .ConnectionString = "Data Source=" & App_Path & ";Persist Security Info=False;Jet OLEDB:Database Password=dtman"
                .Open
            End With
        ElseIf PubBackEnd = "S" Then
            Set GCn = New adodb.Connection
            With GCn
                .CursorLocation = adUseClient
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & App_Path & ";Data Source=" & PubServerName
                .Open
            End With
        End If
        
        TxtVehPath = GCn.Execute("SELECT V_SecFADataPath FROM Division").Fields(0).Value
        PubFADataPath = Pub_DataPath & "\" & TxtVehPath & "\FAData.mdb"
        FADataPath = Pub_DataPath & "\" & TxtVehPath & "\FAData.mdb"
        
        
        If PubBackEnd = "A" Then
            Set GCnFa = New adodb.Connection
            With GCnFa
                .CursorLocation = adUseClient
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .ConnectionString = "Data Source=" & PubFADataPath & ";Persist Security Info=False"
                .Open
            End With
        ElseIf PubBackEnd = "S" Then
            Set GCnFa = New adodb.Connection
            With GCnFa
                .CursorLocation = adUseClient
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & App_Path & ";Data Source=" & PubServerName
                .Open
            End With
        End If
        
        mWinCN = False


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '' Pending : Following values should be updated on the basis of each row of Transaction/master
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        PubDivCode = txtDiv.Text
        PubSiteCode = txtSite.Text
        ForSiteCode = PubSiteCode
        PubLoginDate = PubServerDate
        PubFirmCode = txtFirm.Text
        
        
        PubComp_Name = GCnComp.Execute("Select Comp_Name From Company").Fields(0)
        Exit Sub
        
         'GCn.Execute("Select V_SecCompCode from Division").Fields(0).Value
        PubStartDate = Format(GCnComp.Execute("SELECT Start_Date FROM Company").Fields(0).Value, "dd/MMM/yyyy")

        PubEndDate = DateAdd("YYYY", 1, PubStartDate) - 1
        LblPrefix = "Trf" & Format(PubStartDate, "yy")
        
        UpdateTableStructure
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
Eloop:
    If err.NUMBER = -2147467259 Then
        MsgBox "Database Directory not found", vbExclamation, "Unrecognised database"
        Text1.SetFocus
    End If
    MsgBox err.Description
    If mWinCN Then
        MsgBox "Windows Database not connected", vbExclamation, "Error in Windows Database"
    End If
End Sub

Private Sub CmdPrn_Click()
Dim mQRY As String, RepFileName$, RepTitle$
Dim Rst As adodb.Recordset
Dim i As Integer
'' On Error GoTo ERRORHANDLER
    mQRY = "SELECT * from PrnMissRec"
    Set Rst = New Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open (mQRY), GCn, adOpenDynamic, adLockOptimistic
    
    If Rst.RecordCount <= 0 Then MsgBox "** No Records Found to Print **", vbInformation, Me.Caption: Exit Sub
    RepFileName = "PrnMissRec"
         
        CreateFieldDefFile Rst, PubRepoPath + "\" & RepFileName & ".ttx", True
        Set rpt = rdApp.OpenReport(PubRepoPath + "\" & RepFileName & ".RPT")
        rpt.Database.SetDataSource Rst
        rpt.ReadRecords
        Call Report_View(rpt, "Missing Data List Printing", , False)
        Set Rst = Nothing
        Set rpt = Nothing
Exit Sub
ERRORHANDLER:
      MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub Command1_Click()
    
    Set RsNew1 = GCn.Execute("Select * From Veh_Order where Inv_DocID<>'' and Inv_Docid is not null")
    If RsNew1.RecordCount > 0 Then RsNew1.MoveFirst
    Do Until RsNew1.EOF
        GCn.Execute ("Update Veh_Stock set " & _
                    "Sal_DocID='" & RsNew1!Inv_DocID & "'," & _
                    "Sal_DocIDHelp='" & RsNew1!Inv_DocIDHelp & "'," & _
                    "Sal_Site_Code='" & RsNew1!Inv_SiteCode & "'," & _
                    "Sal_VType='" & RsNew1!Inv_VType & "'," & _
                    "Sal_VNo=" & RsNew1!Inv_No & "," & _
                    "Sal_VDate=" & ConvertDate(RsNew1!Inv_Date) & "," & _
                    "Sal_Rate=" & RsNew1!Net_Amount & "," & _
                    "Ord_SiteCode='" & RsNew1!Ord_SiteCode & "'," & _
                    "Ord_DocID='" & RsNew1!OrdDocId & "'," & _
                    "DelCh_DocID='" & RsNew1!DelCh_DocID & "'," & _
                    "DelCh_Date=" & ConvertDate(RsNew1!DelCh_Dt) & " where ChassisNo='" & RsNew1!Chassis & "'")
        RsNew1.MoveNext
    Loop
End Sub

Private Sub Command2_Click()
'' On Error GoTo Eloop
Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String
Dim mDocNumber As String, mSrvCode As String, mPrefix As String
Dim mBookSite As String, mBookDiv As String, mBookNo As String
Dim mNewCard As Boolean, mCardNo As String
Dim mSupervisor As String, mMechanic As String, mChassis As String, mRegNo As String
Dim mModel As String, mSellingDealer As String, mSrl As Integer
Dim mname As String
    
    CopyCnt = 0
    ErrorCnt = 0
  
    Set ExcelGcn1 = New Connection
    
    If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(Text2.Text) & "1.xls") = False Then MsgBox "Excel File (Header) not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
    
    ExcelGcn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(Text2.Text) & "1.xls;Extended Properties=Excel 8.0"
    Set Master = CreateObject("ADODB.Recordset")
    GSQL = "Select * FROM [" & Text2.Text & "$] Order By [Chassis No]"
    Master.Open GSQL, ExcelGcn1, adOpenStatic

    If Master.RecordCount = 0 Then Exit Sub
  
   
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Chassis No"))) Or StringPass(Master.Fields("Chassis No")) = "" Then
            mChassis = "" 'GCn.Execute("Select Chassis from Job_booking where SiebelDocID='" & Master.Fields("Service Request No") & "'").Fields(0).Value
        Else
            mChassis = XNull(Master.Fields("Chassis No"))
        End If
        
        If mChassis <> "" Then
            If PubDivCode = "C" Then
                mname = left(XNull(Master.Fields("Account")), 40)
            Else
                If Trim(left(Trim(XNull(Master.Fields("First Name"))) & " " & Trim(XNull(Master.Fields("Last Name"))), 40)) = "" Then
                    mname = left(XNull(Master.Fields("Account")), 40)
                Else
                    mname = left(Trim(XNull(Master.Fields("First Name"))) & " " & Trim(XNull(Master.Fields("Last Name"))), 40)
                End If
            End If
            GCn.BeginTrans
            GCn.Execute ("Update Hiscard set Name='" & mname & "' where Chassis='" & mChassis & "'")
            GCn.CommitTrans
        End If
NextRecord:
        Master.MoveNext
    Loop
End Sub

Private Sub Dir1_Change()
    Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
' On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub
Private Sub Form_Load()
'' On Error GoTo Eloop
    WinSetting Me, 11070, 15315             ''8250,  10875
    Drive1.Drive = left(App.Path, 2)
    Dir1.Path = App.Path
    Text1.Text = Dir1.Path
    
            
    Set ErrorGCN = New adodb.Connection
    With ErrorGCN
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & App.Path & "\AutomanSiebel.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password="
        .Open
    End With
    
    'Call UpdateTableStructure
    
    TxtCentralData = ErrorGCN.Execute("select CentralDataDirectory from SystemParameter").Fields(0).Value
    txtSite = ErrorGCN.Execute("select DefaultSite from SystemParameter").Fields(0).Value
    txtDiv = ErrorGCN.Execute("select DefaultDivision from SystemParameter").Fields(0).Value
    txtFirm = ErrorGCN.Execute("select DefaultFirmCode from SystemParameter").Fields(0).Value
    optAuto(0).BackColor = Me.BackColor
    optAuto(1).BackColor = Me.BackColor
    
    'Call CmdPath_Click
    Exit Sub
Eloop:
    MsgBox err.Description, vbInformation, "Information"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ExcelGcn1 = Nothing
    Set ExcelGcn2 = Nothing
    Set GCn = Nothing
End Sub

Public Sub AddFieldTable(db As DAO.Database, TableName As String, FieldName As String, _
        FieldType As Variant, Optional FieldSize As Integer, Optional RequiredYesNo As Boolean, _
        Optional AllowZero As Boolean, Optional DefValue As Variant)
'TableName$, FieldName$, FIELDTYPE As Variant, Optional FieldSize As Integer,
'Optional RequiredYesNo As Boolean, Optional AllowZero As Boolean, Optional DefValue As Variant
Dim tmpRs As DAO.Recordset
Dim n As Integer, TDF As TableDef, FLD As DAO.Field
    Set tmpRs = db.OpenRecordset("select * from " & TableName)
    For n = 0 To tmpRs.Fields.Count - 1
        If UCase(tmpRs.Fields(n).Name) = UCase(FieldName) Then
            GoTo myexit
        End If
    Next
    Set tmpRs = Nothing
    Set TDF = db.TableDefs(TableName)
    Set FLD = TDF.CreateField(FieldName)
    FLD.Type = FieldType
    
    If FieldType = 10 Then      '' Text Field
        FLD.Size = FieldSize
        If Not IsMissing(AllowZero) Then FLD.AllowZeroLength = AllowZero
    End If
    If Not IsMissing(DefValue) Then FLD.DefaultValue = DefValue
    If Not IsMissing(RequiredYesNo) Then FLD.Required = RequiredYesNo
    TDF.Fields.Append FLD
myexit:
    Set tmpRs = Nothing
End Sub

Private Function GetVType(OldVType As String)
'FAVTypeStrWin = "'G_ACR','G_ABR','G_BCP','G_BBP','G_CRN','G_DRN','G_TLR'"
'FAVTypeStrDOS = "'VU'   ,'VV'   ,'VW'   ,'VX'   ,'VC'   ,'VD'   ,'VY','VZ'"
'Telco Receipts : 'VY','VZ'
'Unknown Type in Customer Receipts: 'CP','CR','FM',

Select Case OldVType
    Case "CP", "BP"
        GetVType = "F_BP"
    Case "CR", "BR"
        GetVType = "F_AR"
    Case "CN"
        GetVType = "F_CRN"
    Case "DN"
        GetVType = "F_DRN"
    Case "J"
        GetVType = "F_JV"
    
    Case "VU"
        GetVType = "G_ACR"
    Case "VV"
        GetVType = "G_ABR"
    Case "VW"
        GetVType = "G_BCP"
    Case "VX"
        GetVType = "G_BBP"
    Case "VC"
        GetVType = "G_CRN"
    Case "VD"
        GetVType = "G_DRN"
    
    Case "VY", "VZ"
        GetVType = "G_TLR"
End Select
End Function


Private Function ConvPartyType(SalCatg As Byte)
Select Case SalCatg
    Case 1, 2, 4 '"TASS","TRADER","INSTITUTIONAL"
        ConvPartyType = 1
    Case 3  '"GOVT"
        ConvPartyType = 99
    Case 5  '"GENERAL"
        ConvPartyType = 0
End Select
End Function

Public Function FilterString(Str As String) As String
Dim Str1$, LEN1%, x%, Str2$
    FilterString = Replace(Str, " ", "")
    LEN1 = Len(FilterString)
    x = 1
    While LEN1 > 0
        Str1 = Mid(FilterString, x, 1)
        If (Str1 >= Chr(65) And Str1 <= Chr(90)) Or (Str1 >= Chr(97) And Str1 <= Chr(122)) Or (Str1 >= Chr(48) And Str1 <= Chr(57)) Then
            Str2 = Str2 & Str1
        End If
        x = x + 1
        LEN1 = LEN1 - 1
    Wend
    FilterString = UCase(Str2)
End Function

Private Function GetGovtYN(GovtYN As String) As Byte
If IsNull(GovtYN) Then
    GetGovtYN = 0
ElseIf GovtYN = "N" Then
    GetGovtYN = 0
Else
    GetGovtYN = 1
End If
End Function

Private Function GetFoundSource(DosFType$) As Byte
    Select Case DosFType
        Case "T"    'Hypothication
            GetFoundSource = 0
        Case "H"    'Hire Purchase
            GetFoundSource = 1
        Case "N"    'Own Fund
            GetFoundSource = 2
        Case "L"    'Lease
            GetFoundSource = 3
        Case Else   'Loan
            GetFoundSource = 4
    End Select
End Function

Private Function MatchGrpCode(rs As adodb.Recordset, DosCode$, mGrpCode$, mGrpNature$, mNature$) As Boolean
rs.MoveFirst
rs.Find ("Purc_Code='" & DosCode & "'")
If rs.EOF = False Then
    mGrpCode = "0024"
    mGrpNature = "E"
    mNature = "Purchase"
    MatchGrpCode = True
Else
    rs.MoveFirst
    rs.Find ("Sale_Code='" & DosCode & "'")
    If rs.EOF = False Then
        mGrpCode = "0023"
        mGrpNature = "R"
        mNature = "Sale"
        MatchGrpCode = True
    Else
        rs.MoveFirst
        rs.Find ("TaxCode='" & DosCode & "'")
        If rs.EOF = False Then
            mGrpCode = "0014"
            mGrpNature = "L"
            mNature = "Others"
            MatchGrpCode = True
        End If
    End If
End If
End Function

Private Function FxDesig(ByVal Desg As Integer) As String
If Desg = 1 Then
    FxDesig = "WORKS MANAGER"
ElseIf Desg = 2 Then
    FxDesig = "SUPERVISOR"
ElseIf Desg = 3 Then
    FxDesig = "MECHANIC"
ElseIf Desg = 4 Then
    FxDesig = "HELPER"
Else
    FxDesig = "OTHERS"
End If
End Function


Private Sub ImportBtn_Click(Index As Integer)
'' On Error GoTo Eloop
    If ImportTxt(Index).Text = "" Then MsgBox "File name is not defined", vbInformation, "Validation": ImportTxt(Index).SetFocus: Exit Sub
    
    If PubSiteCode = "" Then MsgBox "Site Code is Empty or Apply button not clicked", vbInformation, "Validation": Exit Sub
    If PubDivCode = "" Then MsgBox "Division Code is Empty or Apply button not clicked", vbInformation, "Validation": Exit Sub
    If MsgBox("Start Data Import Process for " & ImportBtn(Index).Caption & " ? ", vbQuestion + vbYesNo, "Confirmation") = vbNo Then Exit Sub
    
    '' Excel File Recordset Creation
    Select Case Index
        Case 16 '' Material Receipt - Local Purchase (Spares)
            Set ExcelGcn1 = New Connection
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & ".xls") = False Then MsgBox "Excel File not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            ExcelGcn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & ".xls;Extended Properties=Excel 8.0"
            
            Set Master1 = CreateObject("ADODB.Recordset")
            GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] Order By [Order #],[Challan #]"
            Master1.Open GSQL, ExcelGcn1, adOpenStatic
        
            lblRecTotal(Index).Caption = Master1.RecordCount
            lblRecTotal(Index).Refresh
        
            If Master1.RecordCount = 0 Then Exit Sub
        
        Case 18 '' Material Receipt - Tata Motors (Spares)
            '' Header File Details
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "1.xls") = False Then MsgBox "Excel File for Header Data not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            Set ExcelGcn1 = New Connection
            ExcelGcn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "1.xls;Extended Properties=Excel 8.0"
        
            '' Line File Details
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "2.xls") = False Then MsgBox "Excel File for Line Data not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            Set ExcelGcn2 = New Connection
            ExcelGcn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "2.xls;Extended Properties=Excel 8.0"
            
            Set Master1 = CreateObject("ADODB.Recordset")
            GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] Order By [Invoice #],[Challan #],[Order #]"
            Master1.Open GSQL, ExcelGcn2, adOpenStatic
            
            lblRecTotal(Index).Caption = Master1.RecordCount
            lblRecTotal(Index).Refresh
        
            If Master1.RecordCount = 0 Then Exit Sub
        
        Case 19 '' Purchase Bill - Tata Motors (Spares)
            Set ExcelGcn1 = New Connection
            Set ExcelGcn2 = New Connection
            
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "1.xls") = False Then MsgBox "Excel File (Header) not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "2.xls") = False Then MsgBox "Excel File (Line) not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            
            ExcelGcn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "1.xls;Extended Properties=Excel 8.0"
            Set Master = CreateObject("ADODB.Recordset")
            GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] Order By [Invoice #]"
            Master.Open GSQL, ExcelGcn1, adOpenStatic
        
            ExcelGcn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "2.xls;Extended Properties=Excel 8.0"
            
            lblRecTotal(Index).Caption = Master.RecordCount
            lblRecTotal(Index).Refresh
        
            If Master.RecordCount = 0 Then Exit Sub
        
        Case 21 '' Sales Challan Cum Bill (Spares)
            Set ExcelGcn1 = New Connection
            Set ExcelGcn2 = New Connection
            
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "1.xls") = False Then MsgBox "Excel File (Header) not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "2.xls") = False Then MsgBox "Excel File (Line) not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            
            ExcelGcn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "1.xls;Extended Properties=Excel 8.0"
            Set Master = CreateObject("ADODB.Recordset")
            GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] Order By Invoice_No"
            Master.Open GSQL, ExcelGcn1, adOpenStatic
'            ErrorGCN.Execute "SELECT * Into OTC1 FROM [" & ImportTxt(Index).Text & "$] IN '" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "1" & "' 'EXCEL 8.0;' "
            ExcelGcn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "2.xls;Extended Properties=Excel 8.0"
'            ErrorGCN.Execute "SELECT * Into OTC2 FROM [" & ImportTxt(Index).Text & "$] IN '" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "2" & "' 'EXCEL 8.0;' "
            
            lblRecTotal(Index).Caption = Master.RecordCount
            lblRecTotal(Index).Refresh
        
            If Master.RecordCount = 0 Then Exit Sub
        
        Case 20, 22         'Stock Transfer (Indward)/(Outward)
            Set ExcelGcn1 = New Connection
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & ".xls") = False Then MsgBox "Excel File not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            ExcelGcn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & ".xls;Extended Properties=Excel 8.0"
            
            Set Master = CreateObject("ADODB.Recordset")
            GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] Order By Narration"
            Master.Open GSQL, ExcelGcn1, adOpenStatic
    
            lblRecTotal(Index).Caption = Master.RecordCount
            lblRecTotal(Index).Refresh
    
            If Master.RecordCount = 0 Then Exit Sub
        Case 27             ' Job card entry
            Set ExcelGcn1 = New Connection
            Set ExcelGcn2 = New Connection
            
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "1.xls") = False Then MsgBox "Excel File (Header) not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "2.xls") = False Then MsgBox "Excel File (Line) not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            
            ExcelGcn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "1.xls;Extended Properties=Excel 8.0"
            Set Master = CreateObject("ADODB.Recordset")
            GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] Order By [Job Card #]"
            Master.Open GSQL, ExcelGcn1, adOpenStatic
        
            ExcelGcn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "2.xls;Extended Properties=Excel 8.0"
            
            lblRecTotal(Index).Caption = Master.RecordCount
            lblRecTotal(Index).Refresh
        
            If Master.RecordCount = 0 Then Exit Sub
        
        Case 28             ' Job card Labour 1entry
            Set ExcelGcn1 = New Connection
            
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & ".xls") = False Then MsgBox "Excel File (Header) not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            
            ExcelGcn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & ".xls;Extended Properties=Excel 8.0"
            Set Master = CreateObject("ADODB.Recordset")
            GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] Order By [Job Card #]"
            Master.Open GSQL, ExcelGcn1, adOpenStatic
            
            lblRecTotal(Index).Caption = Master.RecordCount
            lblRecTotal(Index).Refresh
        
            If Master.RecordCount = 0 Then Exit Sub
        
        Case 29             ' Job card Close Entry
            Set ExcelGcn1 = New Connection
            Set ExcelGcn2 = New Connection
            
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "1.xls") = False Then MsgBox "Excel File (Header) not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "2.xls") = False Then MsgBox "Excel File (Line) not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            
            ExcelGcn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "1.xls;Extended Properties=Excel 8.0"
            Set Master = CreateObject("ADODB.Recordset")
            GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] Order By [Job Card No]"
            Master.Open GSQL, ExcelGcn1, adOpenStatic
            
            ExcelGcn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & "2.xls;Extended Properties=Excel 8.0"
            
            lblRecTotal(Index).Caption = Master.RecordCount
            lblRecTotal(Index).Refresh
        
            If Master.RecordCount = 0 Then Exit Sub
        
        Case Else
            Set ExcelGcn1 = New Connection
            If Fob.FileExists(Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & ".xls") = False Then MsgBox "Excel File not found, Contact to System Administrator", vbCritical, "Excel File Name/Location Error": Exit Sub
            ExcelGcn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(Text1.Text) & "\" & Trim(ImportTxt(Index).Text) & ".xls;Extended Properties=Excel 8.0"
            
            Set Master = CreateObject("ADODB.Recordset")
            GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$]"
            Master.Open GSQL, ExcelGcn1, adOpenStatic
    
            lblRecTotal(Index).Caption = Master.RecordCount
            lblRecTotal(Index).Refresh
    
            If Master.RecordCount = 0 Then Exit Sub
    End Select
    
    Select Case Index
        Case 0 '' City Master
            Call CityMasterUpdate(Index)
        Case 1 '' Colour Master
            Call ColourMasterUpdate(Index)
        Case 2 '' Area Master
            Call AreaMasterUpdate(Index)
        Case 3 '' Sales Person
            Call SalesRepMasterUpdate(Index)
        Case 4 '' Ref. By/Inquiry Source
            Call RefByMasterUpdate(Index)
        Case 5 '' Purpose/Intended
            Call PurposeMasterUpdate(Index)
        Case 6 '' Model Master (Model Category/Model Group/Model)
            Call ModelMasterUpdate(Index)
        Case 7 '' Godown master
            Call GodownMasterUpdate(Index)
        Case 8 '' Financer master
            Call FinancerMasterUpdate(Index)
        Case 9 '' Vehicle Purchase Data
            Call VehiclePurchaseDataUpdate(Index)
        Case 10 '' Ledger A/c
            Call LedgerAccountDataUpdate(Index, "Vehicle")
        Case 11 '' Vehicle Booking Data (CVD)
            Call VehicleBookingDataUpdate(Index, "CVD")
        Case 12 '' Vehicle Sales Data/Booking
            Call VehicleSalesDataUpdate(Index)
        Case 13 '' Vehicle Sales Money Receipts
            Call MoneyReceiptDataUpdate(Index)
        
        '' Spare Part Department
        Case 15 '' Unit Master
            Call UnitMasterDataUpdate(Index)
        Case 14 '' Part Master
            Call PartMasterDataUpdate(Index)
        Case 16 '' Material Receipt  - Local Purchase
            Call MRNDataUpdate(Index, "Local")
        Case 17 '' Vehicle Booking Data (PCD)
            Call VehicleBookingDataUpdate(Index, "PCD")
        Case 18 '' Material Receipt - From Tata Motors
            Call MRNDataUpdate(Index, "Tata")
        Case 19 '' Purchase Bill - From Tata Motors
            Call PurchBillDataUpdate(Index, "Tata")
        Case 20 '' Stock Received (Inward)
            Call StkTrfUpdate(Index, "Inward")
        Case 21 '' Spare Sales Bill
            Call SprSaleUpdate(Index)
        Case 22 '' Stock Transfer (Outward)
            Call StkTrfUpdate(Index, "Outward")
        Case 23 '' Stock Adjustment (Issue/Recd.)
            Call StkAdjUpdate(Index)
        
        
        Case 24 '' Labour Master
            Call LabourMasterUpdate(Index)
        Case 25 '' Trouble Master
            Call TroubleMasterUpdate(Index)
        Case 26 '' Service Booking
            Call ServiceBookingUpdate(Index)
        Case 27 '' JobCard Entry Data
            Call JobCardEntryUpdate(Index)
        
        Case 28 '' JobCard Labour Entry
            Call JobCardLabourUpdate(Index)
        Case 29 '' JobCard Close Entry
            Call JobCardCloseUpdate(Index)
        Case 30 ''Job Requisition Entry
            Call Job_Req_Update(Index)
        Case 31 '' Ledger A/c
            Call LedgerAccountDataUpdate(Index, "Workshop")
    
    End Select
    Exit Sub
Eloop:
    MsgBox err.NUMBER & " " & err.Description, vbInformation + vbCritical, "Error Message"
    Exit Sub
End Sub
'Private Sub JobCardCloseUpdate(Index)
''' On Error GoTo Eloop
'Dim MasterCode As String, mReqDocId As String, mPartyCode As String, mPartyName As String
'
'Dim mLength As Integer, mCashCode As String, mTaxDetail As Boolean
'Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String
'Dim mDocNumber As String, mPrefix As String
'Dim mSupervisor As String, mMechanic As String
'Dim mSrl As Integer, mV_Type As String, mAmount As Double, mTax_Amt1 As Double
'Dim mJobCardID As String, mPurpose As String, mTrnType As String
'Dim SkipJobcard As Boolean, mGodown As String
'Dim mFormCode As String, mVATApplicable As Boolean, mLubDiscountAllow As Boolean
'Dim mLabAmtTP As Double, mLabAmtTB As Double, mLabTaxPer As Double, mLabTaxAmt As Double, mLabRounded As Double
'Dim mLabDiscPer As Double, mLabDiscAmt As Double, mLabNetAmt As Double
'Dim mLubCategory As String, mLubType As String, mCreditAc As String
'Dim mSpareType As String, mLabourType As String, mSpareDocID As String, mLabourDocID As String, mGatePass As String
'Dim mTempGoods As Double, mTempTaxAmt As Double, mTempDiscAmt As Double, mTempDiscPer As Double, mTempTaxPer As Double
'Dim mTempLub As Double, mTempSpare As Double
'
'Dim mSprAmt_MRP_TB As Double, mSprAmt_MRP_TP As Double, mOilAmt_MRP_TB As Double, mOilAmt_MRP_TP As Double
'Dim mSprAmt_TB As Double, mSprAmt_TP As Double, mOilAmt_TB As Double, mOilAmt_TP As Double
'Dim mD_Per_TB As Double, mD_Per_TP As Double, mD_Amt_TB As Double, mD_Amt_TP As Double
'Dim mD_Per_MRP_TB As Double, mD_Per_MRP_TP As Double, mD_Amt_MRP_TB As Double, mD_Amt_MRP_TP As Double
'Dim mFieldRename As Boolean, mTaxAmtShare As Double
'
'    ImportBtn(Index).BackColor = ProcessColor
'
'
''    mFieldRename = False
''    For i = 0 To Master.Fields.Count - 1
''        If UCase(Master.Fields(i).Name) = UCase("VATTAX") Then
''            mFieldRename = True
''            Exit For
''        End If
''    Next
''
''    If mFieldRename = False Then
''        MsgBox "Field not found  : VatTax", vbCritical, "Field Name not changed"
''        Exit Sub
''    End If
'
'    GCn.BeginTrans
'    CopyCnt = 0
'    ErrorCnt = 0
'
'    Set RsNew = New adodb.Recordset
'    RsNew.CursorLocation = adUseClient
'    RsNew.Open "Select * from SP_Sale", GCn, adOpenDynamic, adLockOptimistic
'
'    Set RsNew1 = New adodb.Recordset
'    RsNew1.CursorLocation = adUseClient
'    RsNew1.Open "Select * from SP_Stock", GCn, adOpenDynamic, adLockOptimistic
'
'    Set rsTemp = New adodb.Recordset
'    rsTemp.CursorLocation = adUseClient
'    rsTemp.Open "Select * from Job_Card", GCn, adOpenDynamic, adLockOptimistic
'
'    If Master.RecordCount > 0 Then Master.MoveFirst
'
'    mV_Type = "W_RGO"
'    mGodown = GCn.Execute("Select SprWorksGodown from Syctrl").Fields(0).Value
'    mTaxDetail = GCn.Execute("Select TaxDetOnSprInv from Syctrl").Fields(0).Value
'    mVATApplicable = ErrorGCN.Execute("Select VATApplicable from Enviro").Fields(0).Value
'    mLubDiscountAllow = GCn.Execute("Select DiscOnLube from Syctrl").Fields(0).Value
'    If mVATApplicable = True Then
'        mFormCode = ErrorGCN.Execute("Select SpareSaleFormVAT from enviro").Fields(0).Value
'    Else
'        mFormCode = ErrorGCN.Execute("Select SpareSaleFormLocal from enviro").Fields(0).Value
'    End If
'
'    Do Until Master.EOF
'        mDocNumber = StringPass(Master.Fields("Job Card No"))
'
'        '' Checking required for Jobcard is already closed or not (manually or earlier from siebel)
'
'        If Trim(mDocNumber) = "" Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("JC #"), "Jobcard Close Entry", "JC # Field is empty")
'            GoTo MyNextRecord
'        End If
'
'        If StringPass(Master.Fields("Invoice_Status")) <> "New" Then
'            GoTo DuplicateSkipped
'        End If
'
'        If GCn.Execute("Select DocID from Job_Card where SiebelDocID='" & mDocNumber & "'").RecordCount = 0 Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Job Card not found in Automan")
'            GoTo MyNextRecord
'        Else
'            mJobCardID = GCn.Execute("Select DocID from Job_Card where SiebelDocID='" & mDocNumber & "'").Fields(0).Value
'        End If
'
'        If GCn.Execute("Select V_No from Sp_Stock where Job_DocID='" & mJobCardID & "' and v_Type='" & mV_Type & "'").RecordCount > 0 Then
'            GoTo DuplicateSkipped
'        End If
'
'        If IsNull(StringPass(Master.Fields("Invoice_Date"))) Or StringPass(Master.Fields("Invoice_Date")) = "" Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Invoice_Date field is blank")
'            GoTo MyNextRecord
'        End If
'
'        mRecordDiv = left(mJobCardID, 1)
'        mRecordSite = Mid(mJobCardID, 2, 2)
'
'        mCashCode = ErrorGCN.Execute("Select CashAccountCode from SiteDivision where AutomanDiv='" & mRecordDiv & "' and AutomanSite='" & left(mRecordSite, 1) & "'").Fields(0).Value
'
'        If StringPass(Master.Fields("Mode of Payment")) = "CREDIT" Then
'            mSpareType = "W_SIR"
'            mLabourType = "W_LIR"
'            mPartyCode = ""
'        Else
'            mSpareType = "W_SIC"
'            mLabourType = "W_LIC"
'            mPartyCode = mCashCode
'        End If
'
'        mLubCategory = "N"
'        mLubType = GCn.Execute("Select PartGrade_Lub from Syctrl").Fields(0).Value
'        mMechanic = ErrorGCN.Execute("Select UnknownMechanic from Enviro").Fields(0).Value
'        mCreditAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc where Div_Code='" & mRecordDiv & "' and Form_Code='" & mFormCode & "'").Fields(0).Value
'
'        '' Document Serial Number
'        Dim mShortYear As String
'        If Month(Master.Fields("INVOICE_Date")) > 3 Then
'            mShortYear = Right(Format(Master.Fields("INVOICE_Date"), "yy"), 1) & Right(Val(Format(Master.Fields("INVOICE_Date"), "yy")) + 1, 1)
'        Else
'            mShortYear = Right(Val(Format(Master.Fields("INVOICE_Date"), "yy")) - 1, 1) & Right(Format(Master.Fields("INVOICE_Date"), "yy"), 1)
'        End If
'        mPrefix = "SBL" & mShortYear 'Format(Master.Fields("INVOICE_Date"), "yy")
'
'        'mPrefix = "SBL"
'
'
'        CodeCnt = GCn.Execute("Select iif(isnull(Max(V_No)),0,Max(V_no))+1 from SP_Stock where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & "' and V_Type='" & mV_Type & "'").Fields(0).Value
'        mReqDocId = mRecordDiv & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
'        mSpareDocID = mRecordDiv & mRecordSite & mSpareType & mPrefix & Right("00000000" & Val(Right(Master.Fields("Invoice_No"), 5)), 8)
'        mLabourDocID = mRecordDiv & mRecordSite & mLabourType & mPrefix & Right("00000000" & Val(Right(Master.Fields("Invoice_No"), 5)), 8)
'        mGatePass = "00000" & GCn.Execute("select iif(isnull(max(val(right(gp_no,5)))),0,max(val(right(gp_no,5))))+1 from job_card where left(gp_no,1)='" & mRecordDiv & "' AND " & cMID("gp_no", "2", "1") & "='" & left(mRecordSite, 1) & "'").Fields(0).Value
'        mGatePass = mRecordDiv & mRecordSite & Right(mGatePass, 5)
'
'        If GCn.Execute("Select DocId_InvSpr from Job_Card where DocId_InvSpr='" & mSpareDocID & "'").RecordCount > 0 Then
'            GoTo DuplicateSkipped
'        End If
'
'
'        If StringPass(Master.Fields("Mode of Payment")) = "CREDIT" Then
'            If IsNull(StringPass(Master.Fields("Account_Code"))) Or StringPass(Master.Fields("Account_Code")) = "" Then
'                If IsNull(StringPass(Master.Fields("Customer_Code"))) Or StringPass(Master.Fields("Customer_Code")) = "" Then
'                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Account/Customer Code field is blank")
'                    GoTo MyNextRecord
'                Else
'                    If GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & StringPass(Master.Fields("Customer_Code")) & "'").RecordCount = 0 Then
'                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Customer Code not found in Automan Software")
'                        GoTo MyNextRecord
'                    Else
'                        mPartyCode = GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & StringPass(Master.Fields("Customer_Code")) & "'").Fields(0).Value
'                        mPartyName = StringPass(Master.Fields("Full Name"))
'                    End If
'                End If
'            Else
'                If GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & StringPass(Master.Fields("Account_Code")) & "'").RecordCount = 0 Then
'                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Account Code not found in Automan Software")
'                    GoTo MyNextRecord
'                Else
'                    mPartyCode = GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & StringPass(Master.Fields("Account_Code")) & "'").Fields(0).Value
'                    mPartyName = StringPass(Master.Fields("Account_Name"))
'                End If
'            End If
'        End If
'
'        If mPartyCode = "" Then
'            MsgBox "PartyCode is blank for Jobcard No " & mDocNumber & ", Entry skipped", vbInformation, "Data Import : JobClose"
'            GoTo MyNextRecord
'        End If
'
'        mLabDiscAmt = VNull(Format(Mid(Master.Fields("Discount Labour"), 4, Len(Master.Fields("Discount Labour")) - 3), "0.00"))
'        mLabNetAmt = VNull(Format(Mid(Master.Fields("Total Labour Amount"), 4, Len(Master.Fields("Total Labour Amount")) - 3), "0.00"))
'        mLabRounded = 0
'        If mRecordDiv = "C" Then
'            mLabAmtTP = VNull(Format(Mid(Master.Fields("Labour Invoice Amount"), 4, Len(Master.Fields("Labour Invoice Amount")) - 3), "0.00"))
'            mLabAmtTB = 0
'            mLabTaxPer = 0
'            mLabTaxAmt = 0
'            If mLabAmtTP > 0 Then
'                mLabDiscPer = Round(mLabDiscAmt * 100 / mLabAmtTP, 4)
'            Else
'                mLabDiscPer = 0
'            End If
'        Else
'            mLabAmtTP = 0
'            mLabAmtTB = VNull(Format(Mid(Master.Fields("Labour Invoice Amount"), 4, Len(Master.Fields("Labour Invoice Amount")) - 3), "0.00"))
'            mLabTaxAmt = VNull(Format(Mid(Master.Fields("Service Tax"), 4, Len(Master.Fields("Service Tax")) - 3), "0.00"))
'            If mLabAmtTB > 0 Then
'                mLabTaxPer = Round(mLabTaxAmt * 100 / mLabAmtTB, 4)
'            Else
'                mLabTaxPer = 0
'            End If
'            If mLabAmtTB > 0 Then
'                mLabDiscPer = Round(mLabDiscAmt * 100 / mLabAmtTB, 4)
'            Else
'                mLabDiscPer = 0
'            End If
'        End If
'
'
'
'        '' Recordset Spares Details for current Jobcard
'        Set Master1 = CreateObject("ADODB.Recordset")
'        GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] where [JC #]='" & mDocNumber & "' and status in ('Invoiced','Cancelled','Partially Shipped','Shipped') Order By [JC #]"
'        Master1.Open GSQL, ExcelGcn2, adOpenStatic
'
'        If Master1.RecordCount > 0 Then Master1.MoveFirst
'        mSrl = 1
'        mAmount = 0
'        mTax_Amt1 = 0
'
'        mSprAmt_MRP_TB = 0: mSprAmt_MRP_TP = 0: mOilAmt_MRP_TB = 0: mOilAmt_MRP_TP = 0
'        mD_Per_MRP_TB = 0: mD_Per_MRP_TP = 0: mD_Amt_MRP_TB = 0: mD_Amt_MRP_TP = 0
'
'        mSprAmt_TB = 0: mSprAmt_TP = 0: mOilAmt_TB = 0: mOilAmt_TP = 0
'        mD_Per_TB = 0: mD_Per_TP = 0: mD_Amt_TB = 0: mD_Amt_TP = 0
'
'        mTempGoods = Val(Format(Master.Fields("Total Parts Amount"), "0.00"))
'        mTempLub = Val(Format(Mid(Master.Fields("Lubricant Amount"), 4, Len(Master.Fields("Lubricant Amount")) - 3), "0.00"))
'        mTempSpare = Val(Format(Mid(Master.Fields("Parts Amount"), 4, Len(Master.Fields("Parts Amount")) - 3), "0.00"))
'        If mTempGoods = 0 Then
'            mTempTaxAmt = 0: mTempDiscAmt = 0: mTempDiscPer = 0: mTempTaxPer = 0
'        Else
'            mTempTaxAmt = Val(Format(Mid(Master.Fields("VAT"), 4, Len(Master.Fields("VAT")) - 3), "0.00"))
'            mTempDiscAmt = Val(Format(Mid(Master.Fields("Discount Job Parts"), 4, Len(Master.Fields("Discount Job Parts")) - 3), "0.00"))
'            If mLubDiscountAllow = False Then
'                If mTempSpare = 0 Then
'                    mTaxAmtShare = Round(mTempLub * mTempTaxAmt / mTempGoods, 2)
'                    mTempDiscPer = Round(mTempDiscAmt * 100 / (mTempLub), 4)            ''+ mTaxAmtShare
'                Else
'                    mTaxAmtShare = Round(mTempSpare * mTempTaxAmt / mTempGoods, 2)
'                    mTempDiscPer = Round(mTempDiscAmt * 100 / (mTempSpare), 4)          ''+ mTaxAmtShare
'                End If
'            Else
'                mTempDiscPer = Round(mTempDiscAmt * 100 / (mTempGoods), 4)     ''+  mTempTaxAmt
'            End If
'            mTempTaxPer = Round(mTempTaxAmt * 100 / mTempGoods, 4)
'        End If
'        If Master1.RecordCount > 0 Then Master1.MoveFirst
'        Do Until Master1.EOF
''            If GCn.Execute("Select Job_DocID from SP_Stock where Job_DocID='" & mJobCardID & "' and V_Type='" & mV_Type & "'").RecordCount > 0 Then
''                GoTo DuplicateSkipped
''            End If
'
'            If IsNull(StringPass(Master1.Fields("Part #"))) Or StringPass(Master1.Fields("Part #")) = "" Then
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Part # field is blank")
'                GoTo MyNextRecord
'            End If
'
'
'            If IsNull(StringPass(Master1.Fields("Billing Type"))) Or StringPass(Master1.Fields("Billing Type")) = "" Then
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Billing Type field is blank in Line File [Default Billing Type assumed as PAID]")
'            Else
'                mPurpose = Master1.Fields("Billing Type")
'            End If
'            mTrnType = GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("part #") & "'").Fields(0).Value
'
'
'            'Insert JobCard Spare in SP_Stock
'            With RsNew1
'                .AddNew
'                !DocId = mReqDocId
'                !Site_Code = mRecordSite
'                !V_Type = mV_Type
'                !V_No = CodeCnt
'                !V_DATE = Format(Master.Fields("Invoice_Date"), "dd/MMM/yyyy")        '' Order Date = Challan Date
'                !Party_Code = mPartyCode
'                !L_C = "L"
'                !Remark = ""
'                !Part_SrlNo = mSrl
'                !Srl_No = mSrl
'                !Part_No = VNull(Master1.Fields("Part #"))
'                !godown = mGodown
'                !Job_DocId = mJobCardID
'                !Job_divCode = mRecordDiv
'                !Mech_Code = mMechanic
'                !TrnComplete_YN = 1
'                If mPurpose = "Free Service" Then
'                    If Master1.Fields("SR Type") = "PDI" Then
'                        !Purpose = "P"
'                    Else
'                        !Purpose = "F"
'                    End If
'                ElseIf mPurpose = "Paid" Then
'                    !Purpose = "C"
'                ElseIf mPurpose = "Warranty" Then
'                    !Purpose = "W"
'                Else
'                    If Master1.Fields("SR Type") = "PDI" Then
'                        !Purpose = "P"
'                    Else
'                        !Purpose = "L"
'                    End If
'                End If
'                If mTrnType = mLubType Then
'                    !Lub_Category = mLubCategory
'                Else
'                    !Lub_Category = ""
'                End If
'                !Qty_Doc = VNull(Master1!Quantity)
'                !Qty_Iss = VNull(Master1.Fields("Qty Shipped"))
'                If Master1.Fields("Status") = "Cancelled" Then
'                    !Qty_Ret = VNull(Master1.Fields("Qty Shipped"))
'                Else
'                    !Qty_Ret = 0
'                End If
'
'                If mVATApplicable Then
'                    !Tax_YN = 1             '' if VAT is applicable in State
'                Else
'                    !Tax_YN = 0     'Question to be Asked IIf(mLocal = "L", 0, 1)
'                End If
'                !MRP_YN = IIf(mRecordDiv = "C", 1, 0)
'
'                '' Goods Value
'                !Amount = IIf(Master1.Fields("Status") = "Cancelled", 0, Val(Format(Master1.Fields("Line Total"), "0.00")))
'                If mRecordDiv = "C" Then
'                    !Disc_Amt = 0       '' IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
'                    If !Amount > 0 Then
'                        !Disc_Per = 0            ''Round(!Disc_Amt * 100 / !Amount, 4)
'                    Else
'                        !Disc_Per = 0
'                    End If
'
'                    '' Tax Value
'                    If mTempTaxAmt > 0 Then
'                        If !Amount > 0 Then
'                            !TaxAmt = Round(!Amount * mTempTaxPer / (100 + mTempTaxPer), 2)
'                            !TaxPer = mTempTaxPer
'                        Else
'                            !TaxAmt = 0
'                            !TaxPer = 0
'                        End If
'                        mTax_Amt1 = mTax_Amt1 + !TaxAmt
'                    Else
'                        !TaxAmt = 0
'                        !TaxPer = 0
'                    End If
'                    !Net_Amt = !Amount - !TaxAmt
'                Else
'                    !Disc_Amt = 0       ''IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
'                    If !Amount > 0 Then
'                        !Disc_Per = 0   ''Round(!Disc_Amt * 100 / !Amount, 4)
'                    Else
'                        !Disc_Per = 0
'                    End If
'
'                    '' Tax Value
'                    If mTempTaxAmt > 0 Then
'                        If !Amount > 0 Then
'                            !TaxAmt = Round(!Amount * mTempTaxPer / 100, 2)
'                            !TaxPer = mTempTaxPer
'                        Else
'                            !TaxAmt = 0
'                            !TaxPer = 0
'                        End If
'                        mTax_Amt1 = mTax_Amt1 + !TaxAmt
'                    Else
'                        !TaxAmt = 0
'                        !TaxPer = 0
'                    End If
'                    !Net_Amt = !Amount
'                End If
'                mTrnType = GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("part #") & "'").Fields(0).Value
'                If mRecordDiv = "C" Then
'                    If !Tax_YN = 1 Then
'                        If mLubType = mTrnType Then
'                            mOilAmt_MRP_TB = mOilAmt_MRP_TB + (!Net_Amt - !Disc_Amt)
'                        Else
'                            mSprAmt_MRP_TB = mSprAmt_MRP_TB + (!Net_Amt - !Disc_Amt)
'                        End If
'                    Else
'                        If mLubType = mTrnType Then
'                            mOilAmt_MRP_TP = mOilAmt_MRP_TP + (!Net_Amt - !Disc_Amt)
'                        Else
'                            mSprAmt_MRP_TP = mSprAmt_MRP_TP + (!Net_Amt - !Disc_Amt)
'                        End If
'                    End If
'                Else
'                    If !Tax_YN = 1 Then
'                        If mLubType = mTrnType Then
'                            mOilAmt_TB = mOilAmt_TB + (!Net_Amt - !Disc_Amt)
'                        Else
'                            mSprAmt_TB = mSprAmt_TB + (!Net_Amt - !Disc_Amt)
'                        End If
'                    Else
'                        If mLubType = mTrnType Then
'                            mOilAmt_TP = mOilAmt_TP + (!Net_Amt - !Disc_Amt)
'                        Else
'                            mSprAmt_TP = mSprAmt_TP + (!Net_Amt - !Disc_Amt)
'                        End If
'                    End If
'                End If
'                If VNull(Master1.Fields("Qty Shipped")) = 0 Then
'                    !Rate = 0
'                    !MRP_Rate = !Rate
'                    '!V_Rate = !Rate
'                Else
'                    If mRecordDiv = "C" Then
'                        !Rate = Round((!Amount - !TaxAmt) / !Qty_Iss, 5)
'                        !MRP_Rate = Round(!Amount / !Qty_Iss, 5)
'                    Else
'                        !Rate = Round(!Amount / !Qty_Iss, 5)
'                        !MRP_Rate = !Rate
'                    End If
'                    '!V_Rate = !Rate     'Round(Val(Replace(Mid(IIf(IsNull(Master1.Fields("Net Amount")), "Rs.0", Master1.Fields("Net Amount")), 4, 15), ",", "")) / Master1!Qty, 4)
'                End If
'
'                !Ord_Discper = 0
'                !Ord_DiscAmt = 0
'
'                '' Invoice Details Updation
'                !Invoice_DocID = mSpareDocID
'                !V_Date2 = Master.Fields("Invoice_Date")
'                !Rate2 = !Rate
'                !MRP_Rate2 = !MRP_Rate
'                !Amount2 = !Amount
'                !Disc_Per2 = !Disc_Per
'                !Disc_Amt2 = !Disc_Amt
'                !Net_Amt2 = !Net_Amt
'
'                !U_Name = "Siebel"
'                !U_EntDt = Format(PubLoginDate, "Short Date")
'                !U_AE = "A"
'                .Update
'                mAmount = mAmount + (!Net_Amt - !Disc_Amt)     'Val(Replace(Mid(IIf(IsNull(Master1.Fields("Net_Amount")), "Rs.0", Master1.Fields("Net_Amount")), 4, 15), ",", ""))
'            End With
'            mSrl = mSrl + 1
'LineFileNextrecord:
'            Master1.MoveNext
'        Loop
'        If mRecordDiv = "C" Then
'            If Round(mAmount, 1) <> Round(mTempGoods, 1) Then
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Line File Goods Value Total is not matched with Header file Goods Value (But Entry Posted in Automan)")
'                'GoTo MyNextRecord
'            End If
'        Else
'            If Round(mAmount, 2) <> mTempGoods Then
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Line File Goods Value Total is not matched with Header file Goods Value (But Entry Posted in Automan)")
'                'GoTo MyNextRecord
'            End If
'        End If
'        If Round(mTax_Amt1, 1) <> Round(mTempTaxAmt, 1) Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Line File Tax Value Total is not matched with Header file Tax Value (But Entry Posted in Automan)")
'            'GoTo MyNextRecord
'        End If
''        If Master.Fields("Discount Parts") <> "Rs.0.00" Then
''            MsgBox ""
''        End If
'        If mRecordDiv = "C" Then
'            If mSprAmt_MRP_TB + mOilAmt_MRP_TB > 0 Then
'                mD_Amt_MRP_TB = mTempDiscAmt    '' IIf(IsNull(Master.Fields("Discount Job Parts")), 0, Val(Format(Mid(Master.Fields("Discount Job Parts"), 4, Len(Master.Fields("Discount Job Parts")) - 3), "0.00")))
'                mD_Amt_TB = mTempDiscAmt        ''IIf(IsNull(Master.Fields("Discount Job Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
'                'If mD_Amt_MRP_TB > 0 Then
'                    mD_Per_MRP_TB = mTempDiscPer    '' Round(mD_Amt_MRP_TB * 100 / (mSprAmt_MRP_TB + mOilAmt_MRP_TB), 4)
'                    mD_Per_TB = mTempDiscPer        ''Round(mD_Amt_MRP_TB * 100 / (mSprAmt_MRP_TB + mOilAmt_MRP_TB), 4)
'                'End If
'            Else
'                mD_Amt_MRP_TP = mTempDiscAmt    '' IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
'                mD_Amt_TP = mTempDiscAmt        '' IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
''                If mD_Amt_MRP_TP > 0 Then
'                    mD_Per_MRP_TP = mTempDiscPer ''Round(mD_Amt_MRP_TP * 100 / (mSprAmt_MRP_TP + mOilAmt_MRP_TP), 4)
'                    mD_Per_TP = mTempDiscPer     '' Round(mD_Amt_MRP_TP * 100 / (mSprAmt_MRP_TP + mOilAmt_MRP_TP), 4)
' '               End If
'            End If
'        Else
'            If mSprAmt_TB + mOilAmt_TB > 0 Then
'                mD_Amt_TB = mTempDiscAmt    '' IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
'                'If mD_Amt_TB > 0 Then
'                    mD_Per_TB = mTempDiscPer    ''Round(mD_Amt_TB * 100 / (mSprAmt_TB + mOilAmt_TB), 4)
'                'End If
'            Else
'                mD_Amt_TP = mTempDiscAmt    '' IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
'                'If mD_Amt_TP > 0 Then
'                    mD_Per_TP = mTempDiscPer ''Round(mD_Amt_TP * 100 / (mSprAmt_TP + mOilAmt_TP), 4)
'                'End If
'            End If
'        End If
'
'        'Insert JobCard Info. in Sp_Sale Table
'        With RsNew
'            .AddNew
'            !DocId = mSpareDocID
'            !DocIDHelp = Replace(mSpareDocID, " ", "")
'            !Site_Code = mRecordSite
'            !V_Type = Trim(mSpareType)
'            !V_No = Right("00000000" & Val(Right(Master.Fields("Invoice_No"), 5)), 8)
'            !V_DATE = Format(Master.Fields("Invoice_Date"), "dd/MMM/yyyy")
'            !Party_Code = mPartyCode
'            !Cash_Credit = Master.Fields("Mode Of Payment")
'            !Party_Name = IIf(Len(mPartyName) > 40, left(Replace(mPartyName, ".", ""), 40), mPartyName)
'            !L_C = "L"
'            !Form_Code = mFormCode
'            !CrAc = mCreditAc
'            !SiebelDocID = Master.Fields("Invoice_No")
'            !Job_DocId = mJobCardID
'            !PType = "General"
'            !GP_No = mGatePass
'            !GP_Date = Format(Master.Fields("Invoice_Date"), "dd/MMM/yyyy")
'
'            !SprAmt_MRP_TB = mSprAmt_MRP_TB
'            !SprAmt_MRP_TP = mSprAmt_MRP_TP
'            !OilAmt_MRP_TB = mOilAmt_MRP_TB
'            !OilAmt_MRP_TP = mOilAmt_MRP_TP
'            !D_Per_MRP_TB = mD_Per_MRP_TB
'            !D_Per_MRP_TP = mD_Per_MRP_TP
'            !D_Amt_MRP_TB = mD_Amt_MRP_TB
'            !D_Amt_MRP_TP = mD_Amt_MRP_TP
'
'            !SprAmt_TB = mSprAmt_TB
'            !SprAmt_TP = mSprAmt_TP
'            !OilAmt_TB = mOilAmt_TB
'            !OilAmt_TP = mOilAmt_TP
'            !D_Per_TB = mD_Per_TB
'            !D_Per_TP = mD_Per_TP
'            !D_Amt_TB = mD_Amt_TB
'            !D_Amt_TP = mD_Amt_TP
'
'            !Addition = 0
'
'            !Tax_Amt = mTempTaxAmt
'            !Packing = Round(IIf(IsNull(Master.Fields("Other Charges")), 0, IIf(IsNull(Master.Fields("Other Charges")), 0, Val(Mid(Master.Fields("Other Charges"), 4, Len(Master.Fields("Other Charges")) - 3)))), 2)
'
'            !TOT_Per = 0
'            !TOT_Amt = 0
'
'            !ReSalTax_Per = 0
'            !ReSalTax_Amt = 0
'
'            !total_amt = Round(Val(Format(Mid(Master.Fields("Spares Invoice Amount"), 4, Len(Master.Fields("Spares Invoice Amount")) - 3), "0.00")), 0)
'
'            Dim tt As Double
'            tt = mSprAmt_MRP_TB + mSprAmt_MRP_TP + mOilAmt_MRP_TB + mOilAmt_MRP_TP + mSprAmt_TB + mSprAmt_TP + mOilAmt_TB + mOilAmt_TP + mTempTaxAmt + Round(IIf(IsNull(Master.Fields("Other Charges")), 0, IIf(IsNull(Master.Fields("Other Charges")), 0, Val(Mid(Master.Fields("Other Charges"), 4, Len(Master.Fields("Other Charges")) - 3)))), 2) - (mD_Amt_TP + mD_Amt_TB)
''            If !total_amt <> Round(tt, 0) Then
''                MsgBox ""
''            End If
'
'            !Rounded = Round(Round(Val(Format(Mid(Master.Fields("Spares Invoice Amount"), 4, Len(Master.Fields("Spares Invoice Amount")) - 3), "0.00")), 0) - tt, 2)
'
'            !Det_Tax = mTaxDetail
'            !AcPosting_YN = 1
'
'            !U_Name = "Siebel"
'            !U_EntDt = Format(PubLoginDate, "Short Date")
'            !U_AE = "A"
'            .Update
'
'        End With
'
'        'Update JobCard for JobClose Information
'
'        GCn.Execute ("Update Job_Card set JobCloseDate=" & ConvertDate(Master.Fields("Invoice_Date")) & ",JobComp_Dt_Time=" & ConvertDate(Master.Fields("Invoice_Date")) & _
'            ",CrMemo=" & IIf(Master.Fields("Mode of Payment") = "CREDIT", 1, 0) & ",BillingName='" & IIf(Len(mPartyName) > 40, left(Replace(mPartyName, ".", ""), 40), mPartyName) & "',DelBy=RecBy_Mechanic" & _
'            ",DrSpr_AcCode='" & mPartyCode & "',DrLab_AcCode='" & mPartyCode & "',DocId_InvSpr='" & mSpareDocID & "',DocId_InvLab='" & mLabourDocID & "',GP_NO='" & mGatePass & _
'            "',LabAmt_TB=" & mLabAmtTB & ",LabAmt_TP=" & mLabAmtTP & ",Lab_D_Amt= " & mLabDiscAmt & ",LabD_Per= " & mLabDiscPer & ",Lab_TaxPer=" & mLabTaxPer & ",Lab_TaxAmt= " & mLabTaxAmt & _
'            ",Lab_RoundOff= " & mLabRounded & ",NetLab_Amt= " & mLabNetAmt & _
'            ",ClosedU_Name='Siebel',ClosedU_EntDt=" & ConvertDate(PubLoginDate) & ",ClosedU_AE='A' where Job_Card.DocId='" & mJobCardID & "'")
'
'DuplicateSkipped:
'        CopyCnt = CopyCnt + 1
'        lblRecCopy(Index).Caption = CopyCnt
'        lblRecCopy(Index).Refresh
'
'MyNextRecord:
'        Master.MoveNext
'    Loop
'    GCn.CommitTrans
'
'    ImportBtn(Index).BackColor = FinishColor
'    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
'
'lblExit:
'    Set RsNew = Nothing
'    Exit Sub
'Eloop:
'    mSrl = 0
'    ErrorCnt = ErrorCnt + mSrl
'    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
'    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Stock Transfer (Inward)','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
'    Resume Next
'End Sub


Private Sub JobCardCloseUpdate(Index)
'' On Error GoTo Eloop
Dim MasterCode As String, mReqDocId As String, mPartyCode As String, mPartyName As String
Dim rsTemp As adodb.Recordset
Dim mEditFlag As Boolean
Dim mLength As Integer, mCashCode As String, mTaxDetail As Boolean
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String
Dim mDocNumber As String, mPrefix As String
Dim mSupervisor As String, mMechanic As String
Dim mSrl As Integer, mV_Type As String, mAmount As Double, mTax_Amt1 As Double
Dim mJobCardID As String, mPurpose As String, mTrnType As String
Dim SkipJobcard As Boolean, mGodown As String
Dim mFormCode As String, mVATApplicable As Boolean, mLubDiscountAllow As Boolean
Dim mLabAmtTP As Double, mLabAmtTB As Double, mLabTaxPer As Double, mLabTaxAmt As Double, mLabRounded As Double
Dim mLabDiscPer As Double, mLabDiscAmt As Double, mLabNetAmt As Double
Dim mLubCategory As String, mLubType As String, mCreditAc As String
Dim mSpareType As String, mLabourType As String, mSpareDocID As String, mLabourDocID As String, mGatePass As String
Dim mTempGoods As Double, mTempTaxAmt As Double, mTempDiscAmt As Double, mTempDiscPer As Double, mTempTaxPer As Double
Dim mTempLub As Double, mTempSpare As Double
Dim mDisPer As Double
Dim mDisAmt As Double
Dim tt As Double
Dim mSprAmt_MRP_TB As Double, mSprAmt_MRP_TP As Double, mOilAmt_MRP_TB As Double, mOilAmt_MRP_TP As Double
Dim mSprAmt_TB As Double, mSprAmt_TP As Double, mOilAmt_TB As Double, mOilAmt_TP As Double
Dim mD_Per_TB As Double, mD_Per_TP As Double, mD_Amt_TB As Double, mD_Amt_TP As Double
Dim mD_Per_MRP_TB As Double, mD_Per_MRP_TP As Double, mD_Amt_MRP_TB As Double, mD_Amt_MRP_TP As Double
Dim mFieldRename As Boolean, mTaxAmtShare As Double
  
    ImportBtn(Index).BackColor = ProcessColor
    

'    mFieldRename = False
'    For i = 0 To Master.Fields.Count - 1
'        If UCase(Master.Fields(i).Name) = UCase("VATTAX") Then
'            mFieldRename = True
'            Exit For
'        End If
'    Next
'
'    If mFieldRename = False Then
'        MsgBox "Field not found  : VatTax", vbCritical, "Field Name not changed"
'        Exit Sub
'    End If
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from SP_Sale", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsNew1 = New adodb.Recordset
    RsNew1.CursorLocation = adUseClient
    RsNew1.Open "Select * from SP_Stock", GCn, adOpenDynamic, adLockOptimistic
    
    Set rsTemp = New adodb.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open "Select * from Job_Card", GCn, adOpenDynamic, adLockOptimistic
   
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    mV_Type = "W_RGO"
    mGodown = GCn.Execute("Select SprWorksGodown from Syctrl").Fields(0).Value
    mTaxDetail = GCn.Execute("Select TaxDetOnSprInv from Syctrl").Fields(0).Value
    mVATApplicable = ErrorGCN.Execute("Select VATApplicable from Enviro").Fields(0).Value
    mLubDiscountAllow = GCn.Execute("Select DiscOnLube from Syctrl").Fields(0).Value
    If mVATApplicable = True Then
        mFormCode = ErrorGCN.Execute("Select SpareSaleFormVAT from enviro").Fields(0).Value
    Else
        mFormCode = ErrorGCN.Execute("Select SpareSaleFormLocal from enviro").Fields(0).Value
    End If
    
    Do Until Master.EOF
        mDocNumber = StringPass(Master.Fields("Job Card No"))
            
        '' Checking required for Jobcard is already closed or not (manually or earlier from siebel)
            
        If Trim(mDocNumber) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("JC #"), "Jobcard Close Entry", "JC # Field is empty")
            GoTo MyNextRecord
        End If
            
        If StringPass(Master.Fields("Invoice_Status")) <> "New" Then
            GoTo DuplicateSkipped
        End If
        
        If GCn.Execute("Select DocID from Job_Card where SiebelDocID='" & mDocNumber & "'").RecordCount = 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Job Card not found in Automan")
            GoTo MyNextRecord
        Else
            mJobCardID = GCn.Execute("Select DocID from Job_Card where SiebelDocID='" & mDocNumber & "'").Fields(0).Value
        End If
        
        If GCn.Execute("Select V_No from Sp_Stock where Job_DocID='" & mJobCardID & "' and v_Type='" & mV_Type & "'").RecordCount > 0 Then
            GCn.Execute ("Update Job_Card set JobCloseDate=" & cIIF("TempCloseDate Is Null", ConvertDate(MakeDate(Master.Fields("Invoice_Date"))), "TempCloseDate") & ",JobComp_Dt_Time=" & cIIF("TempCloseDate Is Null", ConvertDate(MakeDate(Master.Fields("Invoice_Date"))), "TempCloseDate") & _
                " where Job_Card.DocId='" & mJobCardID & "'")
            'GoTo DuplicateSkipped
            mEditFlag = True
        End If
        
        If IsNull(StringPass(Master.Fields("Invoice_Date"))) Or StringPass(Master.Fields("Invoice_Date")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Invoice_Date field is blank")
            GoTo MyNextRecord
        End If
        
        mRecordDiv = left(mJobCardID, 1)
        mRecordSite = Mid(mJobCardID, 2, 2)
        
        mCashCode = ErrorGCN.Execute("Select CashAccountCode from SiteDivision where AutomanDiv='" & mRecordDiv & "' and AutomanSite='" & left(mRecordSite, 1) & "'").Fields(0).Value
        
        If StringPass(Master.Fields("Mode of Payment")) = "CREDIT" Then
            mSpareType = "W_SIR"
            mLabourType = "W_LIR"
            mPartyCode = ""
        Else
            mSpareType = "W_SIC"
            mLabourType = "W_LIC"
            mPartyCode = mCashCode
        End If
        
        mLubCategory = "N"
        mLubType = GCn.Execute("Select PartGrade_Lub from Syctrl").Fields(0).Value
        mMechanic = ErrorGCN.Execute("Select UnknownMechanic from Enviro").Fields(0).Value
        mCreditAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc where Div_Code='" & mRecordDiv & "' and Form_Code='" & mFormCode & "'").Fields(0).Value
        
        '' Document Serial Number
        Dim mShortYear As String
        If Month(Master.Fields("INVOICE_Date")) > 3 Then
            mShortYear = Right(Format(Master.Fields("INVOICE_Date"), "yy"), 1) & Right(Val(Format(Master.Fields("INVOICE_Date"), "yy")) + 1, 1)
        Else
            mShortYear = Right(Val(Format(Master.Fields("INVOICE_Date"), "yy")) - 1, 1) & Right(Format(Master.Fields("INVOICE_Date"), "yy"), 1)
        End If
        mPrefix = "SBL" & mShortYear 'Format(Master.Fields("INVOICE_Date"), "yy")
                
        'mPrefix = "SBL"
        
        
        CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & " + 1 from SP_Stock where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & "' and V_Type='" & mV_Type & "'").Fields(0).Value
        If mEditFlag Then
            If GCn.Execute("Select count(*) From Sp_Stock where Job_DocID='" & mJobCardID & "' and v_Type='" & mV_Type & "'").Fields(0) > 0 Then
                mReqDocId = GCn.Execute("Select DocId From Sp_Stock where Job_DocID='" & mJobCardID & "' and v_Type='" & mV_Type & "'").Fields(0)
            Else
                mReqDocId = mRecordDiv & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
            End If
            
            If GCn.Execute("Select GP_No From Job_Card where DocID='" & mJobCardID & "'").Fields(0) > 0 Then
                mGatePass = GCn.Execute("Select GP_No From Job_Card where DocID='" & mJobCardID & "'").Fields(0)
            Else
                mGatePass = "00000" & GCn.Execute("select " & vIsNull("max(" & cVal("right(gp_no,5)") & ")", "0") & "+1 from job_card where left(gp_no,1)='" & mRecordDiv & "' AND " & cMID("gp_no", "2", "1") & "='" & left(mRecordSite, 1) & "'").Fields(0).Value
                mGatePass = mRecordDiv & mRecordSite & Right(mGatePass, 5)
            End If
        Else
            mReqDocId = mRecordDiv & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
            mGatePass = "00000" & GCn.Execute("select " & vIsNull("max(" & cVal("right(gp_no,5)") & ")", "0") & "+1 from job_card where left(gp_no,1)='" & mRecordDiv & "' AND " & cMID("gp_no", "2", "1") & "='" & left(mRecordSite, 1) & "'").Fields(0).Value
            mGatePass = mRecordDiv & mRecordSite & Right(mGatePass, 5)
        End If
        mSpareDocID = mRecordDiv & mRecordSite & mSpareType & mPrefix & Right("00000000" & Val(Right(Master.Fields("Invoice_No"), 5)), 8)
        mLabourDocID = mRecordDiv & mRecordSite & mLabourType & mPrefix & Right("00000000" & Val(Right(Master.Fields("Invoice_No"), 5)), 8)
            
        If GCn.Execute("Select DocId_InvSpr from Job_Card where DocId_InvSpr='" & mSpareDocID & "'").RecordCount > 0 Then
            GCn.Execute ("Update Job_Card set JobCloseDate=" & cIIF("TempCloseDate Is Null", ConvertDate(MakeDate(Master.Fields("Invoice_Date"))), "TempCloseDate") & ",JobComp_Dt_Time=" & cIIF("TempCloseDate Is Null", ConvertDate(MakeDate(Master.Fields("Invoice_Date"))), "TempCloseDate") & _
                " where Job_Card.DocId='" & mJobCardID & "'")
            'GoTo DuplicateSkipped
            mEditFlag = True
        End If
            
            
        If StringPass(Master.Fields("Mode of Payment")) = "CREDIT" Then
            If IsNull(StringPass(Master.Fields("Account_Code"))) Or StringPass(Master.Fields("Account_Code")) = "" Then
                If IsNull(StringPass(Master.Fields("Customer_Code"))) Or StringPass(Master.Fields("Customer_Code")) = "" Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Account/Customer Code field is blank")
                    GoTo MyNextRecord
                Else
                    If GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & StringPass(Master.Fields("Customer_Code")) & "'").RecordCount = 0 Then
                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Customer Code not found in Automan Software")
                        GoTo MyNextRecord
                    Else
                        mPartyCode = GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & StringPass(Master.Fields("Customer_Code")) & "'").Fields(0).Value
                        mPartyName = StringPass(Master.Fields("Full Name"))
                    End If
                End If
            Else
                If GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & StringPass(Master.Fields("Account_Code")) & "'").RecordCount = 0 Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Account Code not found in Automan Software")
                    GoTo MyNextRecord
                Else
                    mPartyCode = GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & StringPass(Master.Fields("Account_Code")) & "'").Fields(0).Value
                    mPartyName = StringPass(Master.Fields("Account_Name"))
                End If
            End If
        End If

        If mPartyCode = "" Then
            MsgBox "PartyCode is blank for Jobcard No " & mDocNumber & ", Entry skipped", vbInformation, "Data Import : JobClose"
            GoTo MyNextRecord
        End If
        
        mLabDiscAmt = VNull(Format(Mid(Master.Fields("Discount Labour"), 4, Len(Master.Fields("Discount Labour")) - 3), "0.00"))
        mLabNetAmt = VNull(Format(Mid(Master.Fields("Total Labour Amount"), 4, Len(Master.Fields("Total Labour Amount")) - 3), "0.00"))
        mLabRounded = 0
        If mRecordDiv = "C" Then
            mLabAmtTP = VNull(Format(Mid(Master.Fields("Labour Invoice Amount"), 4, Len(Master.Fields("Labour Invoice Amount")) - 3), "0.00"))
            mLabAmtTB = 0
            mLabTaxPer = 0
            mLabTaxAmt = 0
            If mLabAmtTP > 0 Then
                mLabDiscPer = Round(mLabDiscAmt * 100 / mLabAmtTP, 4)
            Else
                mLabDiscPer = 0
            End If
        Else
            mLabAmtTP = 0
            mLabAmtTB = VNull(Format(Mid(Master.Fields("Labour Invoice Amount"), 4, Len(Master.Fields("Labour Invoice Amount")) - 3), "0.00"))
            mLabTaxAmt = Format(Val(Mid(Replace(Master.Fields("Service Tax"), ",", ""), 4, Len(Replace(Master.Fields("Service Tax"), ",", "")) - 3)) + Val(Mid(Replace(Master.Fields("Cess Tax"), ",", ""), 4, Len(Replace(Master.Fields("Cess Tax"), ",", "")) - 3)), "0.00")
            
            If mLabAmtTB > 0 Then
                mLabTaxPer = Round(mLabTaxAmt * 100 / mLabAmtTB, 4)
            Else
                mLabTaxPer = 0
            End If
            If mLabAmtTB > 0 Then
                mLabDiscPer = Round(mLabDiscAmt * 100 / mLabAmtTB, 4)
            Else
                mLabDiscPer = 0
            End If
        End If
        
        
        
        '' Recordset Spares Details for current Jobcard
        Set Master1 = CreateObject("ADODB.Recordset")
        GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] where [Order Number]='" & mDocNumber & "' and [Invoice Status] in ('New','Invoiced','Cancelled','Partially Shipped','Shipped') Order By [Order Number]"
        Master1.Open GSQL, ExcelGcn2, adOpenStatic
        If mEditFlag Then
            GCn.Execute "Delete From Sp_Stock Where DocId='" & mReqDocId & "'"
        End If
        
        If Master1.RecordCount > 0 Then Master1.MoveFirst
        mSrl = 1
        mAmount = 0
        mTax_Amt1 = 0
        
        mSprAmt_MRP_TB = 0: mSprAmt_MRP_TP = 0: mOilAmt_MRP_TB = 0: mOilAmt_MRP_TP = 0
        mD_Per_MRP_TB = 0: mD_Per_MRP_TP = 0: mD_Amt_MRP_TB = 0: mD_Amt_MRP_TP = 0
        
        mSprAmt_TB = 0: mSprAmt_TP = 0: mOilAmt_TB = 0: mOilAmt_TP = 0
        mD_Per_TB = 0: mD_Per_TP = 0: mD_Amt_TB = 0: mD_Amt_TP = 0
        
        mTempGoods = Val(Format(Master.Fields("Total Parts Amount"), "0.00"))
        mTempLub = Val(Format(Mid(Master.Fields("Lubricant Amount"), 4, Len(Master.Fields("Lubricant Amount")) - 3), "0.00"))
        mTempSpare = Val(Format(Mid(Master.Fields("Parts Amount"), 4, Len(Master.Fields("Parts Amount")) - 3), "0.00"))
        If mTempGoods = 0 Then
            mTempTaxAmt = 0: mTempDiscAmt = 0: mTempDiscPer = 0: mTempTaxPer = 0
        Else
            mTempTaxAmt = Val(Format(Mid(Master.Fields("VAT"), 4, Len(Master.Fields("VAT")) - 3), "0.00"))
            mTempDiscAmt = Val(Format(Mid(Master.Fields("Discount Job Parts"), 4, Len(Master.Fields("Discount Job Parts")) - 3), "0.00"))
            If mLubDiscountAllow = False Then
                If mTempSpare = 0 Then
                    mTaxAmtShare = Round(mTempLub * mTempTaxAmt / mTempGoods, 2)
                    mTempDiscPer = Round(mTempDiscAmt * 100 / (mTempLub), 4)            ''+ mTaxAmtShare
                Else
                    mTaxAmtShare = Round(mTempSpare * mTempTaxAmt / mTempGoods, 2)
                    mTempDiscPer = Round(mTempDiscAmt * 100 / (mTempSpare), 4)          ''+ mTaxAmtShare
                End If
            Else
                mTempDiscPer = Round(mTempDiscAmt * 100 / (mTempGoods), 4)     ''+  mTempTaxAmt
            End If
            mTempTaxPer = Round(mTempTaxAmt * 100 / mTempGoods, 4)
        End If
        If Master1.RecordCount > 0 Then Master1.MoveFirst
        Do Until Master1.EOF
'            If GCn.Execute("Select Job_DocID from SP_Stock where Job_DocID='" & mJobCardID & "' and V_Type='" & mV_Type & "'").RecordCount > 0 Then
'                GoTo DuplicateSkipped
'            End If
        
            If IsNull(StringPass(Master1.Fields("Part No"))) Or StringPass(Master1.Fields("Part No")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Part No field is blank")
                GoTo MyNextRecord
            End If
        
            
            If IsNull(StringPass(Master1.Fields("Billing Type"))) Or StringPass(Master1.Fields("Billing Type")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Billing Type field is blank in Line File [Default Billing Type assumed as PAID]")
            Else
                mPurpose = Master1.Fields("Billing Type")
            End If
            
            Set rsTemp = GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("Part No") & "'")
            If rsTemp.RecordCount > 0 Then
                mTrnType = XNull(rsTemp!Part_Grade)
            Else
                mTrnType = "S"
            End If
            
            'Insert JobCard Spare in SP_Stock
            With RsNew1
                .AddNew
                !DocId = mReqDocId
                !Site_Code = mRecordSite
                !V_Type = mV_Type
                !V_No = CodeCnt
                !V_DATE = Format(Master1.Fields("Date"), "dd/MMM/yyyy")        '' Order Date = Challan Date
                !Party_Code = mPartyCode
                !L_C = "L"
                !Remark = ""
                !Part_SrlNo = mSrl
                !Srl_No = mSrl
                !Part_No = VNull(Master1.Fields("Part No"))
                !godown = mGodown
                !Job_DocId = mJobCardID
                !Job_divCode = mRecordDiv
                !Mech_Code = mMechanic
                !TrnComplete_YN = 1
                If mPurpose = "Free Service" Then
                    !Purpose = "F"
                ElseIf mPurpose = "Paid" Then
                    !Purpose = "C"
                ElseIf mPurpose = "Warranty" Then
                    !Purpose = "W"
                Else
                    !Purpose = "L"
                End If
                
                If mTrnType = mLubType Then
                    !Lub_Category = mLubCategory
                Else
                    !Lub_Category = ""
                End If
                !Qty_Doc = VNull(Master1.Fields("Sold Qty"))
                !Qty_Iss = VNull(Master1.Fields("Sold Qty"))
                If Master1.Fields("Invoice Status") = "Cancelled" Then
                    !Qty_Ret = VNull(Master1.Fields("Sold Qty"))
                Else
                    !Qty_Ret = 0
                End If
                
                If mVATApplicable Then
                    !Tax_YN = 1             '' if VAT is applicable in State
                Else
                    !Tax_YN = 0     'Question to be Asked IIf(mLocal = "L", 0, 1)
                End If
                !MRP_YN = IIf(mRecordDiv = "C", 1, 0)
                
                '' Goods Value
                !Amount = IIf(Master1.Fields("Invoice Status") = "Cancelled", 0, Val(Format(Master1.Fields("Value"), "0.00")))
                If mRecordDiv = "C" Then
                    !Disc_Amt = 0       '' IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
                    If !Amount > 0 Then
                        !Disc_Per = 0            ''Round(!Disc_Amt * 100 / !Amount, 4)
                    Else
                        !Disc_Per = 0
                    End If
                    
                    '' Tax Value
                    If mTempTaxAmt > 0 Then
                        If !Amount > 0 Then
                            !TaxAmt = Round(!Amount * mTempTaxPer / (100 + mTempTaxPer), 2)
                            !TaxPer = mTempTaxPer
                        Else
                            !TaxAmt = 0
                            !TaxPer = 0
                        End If
                        If !Purpose = "C" Then
                            mTax_Amt1 = mTax_Amt1 + !TaxAmt
                        End If
                    Else
                        !TaxAmt = 0
                        !TaxPer = 0
                    End If
                    !Net_Amt = !Amount - !TaxAmt
                Else
'                    !Disc_Amt = 0       ''IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
'                    If !Amount > 0 Then
'                        !Disc_Per = 0   ''Round(!Disc_Amt * 100 / !Amount, 4)
'                    Else
'                        !Disc_Per = 0
'                    End If
                    
                    
                    
                    !Disc_Amt = ((VNull(Master1.Fields("Rate")) * VNull(Master1.Fields("Sold Qty"))) - VNull(Master1.Fields("Value")))
                    If VNull(Master1.Fields("Value")) <> 0 Then
                        !Disc_Per = !Disc_Amt * 100 / VNull(Master1.Fields("Value"))
                    Else
                        !Disc_Per = 0
                        !Disc_Amt = 0
                    End If
                    
                    '' Tax Value
                    If mTempTaxAmt > 0 Then
                        If !Amount > 0 Then
'                            !TaxAmt = Round(!Amount * mTempTaxPer / 100, 2)
'                            !TaxPer = mTempTaxPer
                            !TaxAmt = VNull(Master1.Fields("Tax Amount After Discount"))
                            !TaxPer = VNull(Master1.Fields("Tax Amount After Discount")) * 100 / VNull(Master1.Fields("Value"))
                        Else
                            !TaxAmt = 0
                            !TaxPer = 0
                        End If
                        If !Purpose = "C" Then
                            mTax_Amt1 = mTax_Amt1 + !TaxAmt
                        End If
                    Else
                        !TaxAmt = 0
                        !TaxPer = 0
                    End If
                    !Net_Amt = !Amount
                End If
                
                Set rsTemp = GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("Part No") & "'")
                If rsTemp.RecordCount > 0 Then
                    mTrnType = XNull(rsTemp!Part_Grade)
                Else
                    mTrnType = "S"
                End If
                If mRecordDiv = "C" Then
                    If !Purpose = "C" Then
                        If !Tax_YN = 1 Then
                            If mLubType = mTrnType Then
                                mOilAmt_MRP_TB = mOilAmt_MRP_TB + (!Net_Amt - !Disc_Amt)
                            Else
                                mSprAmt_MRP_TB = mSprAmt_MRP_TB + (!Net_Amt - !Disc_Amt)
                            End If
                        Else
                            If mLubType = mTrnType Then
                                mOilAmt_MRP_TP = mOilAmt_MRP_TP + (!Net_Amt - !Disc_Amt)
                            Else
                                mSprAmt_MRP_TP = mSprAmt_MRP_TP + (!Net_Amt - !Disc_Amt)
                            End If
                        End If
                    End If
                Else
                    If !Purpose = "C" Then
                        If !Tax_YN = 1 Then
                            If mLubType = mTrnType Then
                                mOilAmt_TB = mOilAmt_TB + (!Net_Amt - !Disc_Amt)
                            Else
                                mSprAmt_TB = mSprAmt_TB + (!Net_Amt - !Disc_Amt)
                            End If
                        Else
                            If mLubType = mTrnType Then
                                mOilAmt_TP = mOilAmt_TP + (!Net_Amt - !Disc_Amt)
                            Else
                                mSprAmt_TP = mSprAmt_TP + (!Net_Amt - !Disc_Amt)
                            End If
                        End If
                    End If
                End If
                If VNull(Master1.Fields("Sold Qty")) = 0 Then
                    !Rate = 0
                    !Mrp_Rate = !Rate
                    '!V_Rate = !Rate
                Else
                    If mRecordDiv = "C" Then
                        !Rate = Round((!Amount - !TaxAmt) / !Qty_Iss, 5)
                        !Mrp_Rate = Round(!Amount / !Qty_Iss, 5)
                    Else
                        !Rate = VNull(Master1.Fields("Rate"))   'Round(!Amount / !Qty_Iss, 5)
                        !Mrp_Rate = !Rate
                    End If
                    '!V_Rate = !Rate     'Round(Val(Replace(Mid(IIf(IsNull(Master1.Fields("Net Amount")), "Rs.0", Master1.Fields("Net Amount")), 4, 15), ",", "")) / Master1!Qty, 4)
                End If
                
                !Ord_Discper = 0
                !Ord_DiscAmt = 0
    
                '' Invoice Details Updation
                !Invoice_DocID = mSpareDocID
                !V_Date2 = Format(Master1.Fields("Date"), "DD/MMM/YYYY")
                !Rate2 = !Rate
                !MRP_Rate2 = !Mrp_Rate
                !Amount2 = !Amount
                !Disc_Per2 = !Disc_Per
                !Disc_Amt2 = !Disc_Amt
                !Net_Amt2 = !Net_Amt
    
                !U_Name = "Siebel"
                !U_EntDt = Format(PubLoginDate, "Short Date")
                !U_AE = IIf(mEditFlag, "E", "A")
                .Update
                If !Purpose = "C" Then
                    mAmount = mAmount + !Net_Amt
                End If
            End With
            mSrl = mSrl + 1
LineFileNextrecord:
            Master1.MoveNext
        Loop
        
        
        
        If mRecordDiv = "C" Then
            If Int(mAmount) <> Int(mTempGoods) Then
                If Round(mAmount) <> Round(mTempGoods) Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Line File Goods Value Total(" & Round(mAmount, 1) & ") is not matched with Header file Goods Value(" & Round(mTempGoods, 1) & ")")
                    GoTo MyNextRecord
                End If
            End If
        Else
            If Int(mAmount) <> Int(mTempGoods) Then
                If Round(mAmount) <> Round(mTempGoods) Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Line File Goods Value Total(" & mAmount & ") is not matched with Header file Goods Value(" & mTempGoods & ") (But Entry Posted in Automan)")
                    'GoTo MyNextRecord
                End If
            End If
    
        End If
        If Format(mTax_Amt1, "0") <> Format(mTempTaxAmt, "0") Then
            If Round(mTax_Amt1) <> Round(mTempTaxAmt) Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Close Entry", "Line File Tax(" & Round(mTax_Amt1, 1) & ") Value Total is not matched with Header file Tax Value(" & Round(mTempTaxAmt, 1) & ") (But Entry Posted in Automan)")
            End If
        End If
        
        
'        If Master.Fields("Discount Parts") <> "Rs.0.00" Then
'            MsgBox ""
'        End If
        If mRecordDiv = "C" Then
            If mSprAmt_MRP_TB + mOilAmt_MRP_TB > 0 Then
                mD_Amt_MRP_TB = mTempDiscAmt    '' IIf(IsNull(Master.Fields("Discount Job Parts")), 0, Val(Format(Mid(Master.Fields("Discount Job Parts"), 4, Len(Master.Fields("Discount Job Parts")) - 3), "0.00")))
                mD_Amt_TB = mTempDiscAmt        ''IIf(IsNull(Master.Fields("Discount Job Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
                'If mD_Amt_MRP_TB > 0 Then
                    mD_Per_MRP_TB = mTempDiscPer    '' Round(mD_Amt_MRP_TB * 100 / (mSprAmt_MRP_TB + mOilAmt_MRP_TB), 4)
                    mD_Per_TB = mTempDiscPer        ''Round(mD_Amt_MRP_TB * 100 / (mSprAmt_MRP_TB + mOilAmt_MRP_TB), 4)
                'End If
            Else
                mD_Amt_MRP_TP = mTempDiscAmt    '' IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
                mD_Amt_TP = mTempDiscAmt        '' IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
'                If mD_Amt_MRP_TP > 0 Then
                    mD_Per_MRP_TP = mTempDiscPer ''Round(mD_Amt_MRP_TP * 100 / (mSprAmt_MRP_TP + mOilAmt_MRP_TP), 4)
                    mD_Per_TP = mTempDiscPer     '' Round(mD_Amt_MRP_TP * 100 / (mSprAmt_MRP_TP + mOilAmt_MRP_TP), 4)
 '               End If
            End If
        Else
            If mSprAmt_TB + mOilAmt_TB > 0 Then
                mD_Amt_TB = mTempDiscAmt    '' IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
                'If mD_Amt_TB > 0 Then
                    mD_Per_TB = mTempDiscPer    ''Round(mD_Amt_TB * 100 / (mSprAmt_TB + mOilAmt_TB), 4)
                'End If
            Else
                mD_Amt_TP = mTempDiscAmt    '' IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
                'If mD_Amt_TP > 0 Then
                    mD_Per_TP = mTempDiscPer ''Round(mD_Amt_TP * 100 / (mSprAmt_TP + mOilAmt_TP), 4)
                'End If
            End If
        End If
        
        'Insert JobCard Info. in Sp_Sale Table
        If mEditFlag And GCn.Execute("Select Count(*) From Sp_Sale Where Job_DocId='" & mJobCardID & "'").Fields(0) > 0 Then
            tt = mSprAmt_MRP_TB + mSprAmt_MRP_TP + mOilAmt_MRP_TB + mOilAmt_MRP_TP + mSprAmt_TB + mSprAmt_TP + mOilAmt_TB + mOilAmt_TP + mTempTaxAmt + Round(IIf(IsNull(Master.Fields("Other Charges")), 0, IIf(IsNull(Master.Fields("Other Charges")), 0, Val(Mid(Master.Fields("Other Charges"), 4, Len(Master.Fields("Other Charges")) - 3)))), 2) - (mD_Amt_TP + mD_Amt_TB)
        
            GCn.Execute "Update Sp_Sale Set DocId = '" & mSpareDocID & "', DocIDHelp = '" & Replace(mSpareDocID, " ", "") & "', " & _
                                        "Site_Code = " & mRecordSite & ", V_Type = '" & Trim(mSpareType) & "', V_No = " & Right("00000000" & Val(Right(Master.Fields("Invoice_No"), 5)), 8) & ", " & _
                                        "V_DATE = " & ConvertDate(MakeDate(Master.Fields("Invoice_Date"))) & ", SiebelDocID = ' " & Master.Fields("Invoice_No") & "', GP_Date = " & ConvertDate(MakeDate(Master.Fields("Invoice_Date"))) & ", " & _
                                        "SprAmt_MRP_TB = " & mSprAmt_MRP_TB & ", SprAmt_MRP_TP = " & mSprAmt_MRP_TP & ", OilAmt_MRP_TB = " & mOilAmt_MRP_TB & ", OilAmt_MRP_TP = " & mOilAmt_MRP_TP & ", " & _
                                        "D_Per_MRP_TB = " & mD_Per_MRP_TB & ", D_Per_MRP_TP = " & mD_Per_MRP_TP & ", D_Amt_MRP_TB = " & mD_Amt_MRP_TB & ", D_Amt_MRP_TP = " & mD_Amt_MRP_TP & ", " & _
                                        "SprAmt_TB = " & mSprAmt_TB & ", SprAmt_TP = " & mSprAmt_TP & ", OilAmt_TB = " & mOilAmt_TB & ", OilAmt_TP = " & mOilAmt_TP & ", D_Per_TB = " & mD_Per_TB & ", " & _
                                        "D_Per_TP = " & mD_Per_TP & ", D_Amt_TB = " & mD_Amt_TB & ", D_Amt_TP = " & mD_Amt_TP & ", Addition = 0, Tax_Amt = " & mTempTaxAmt & ", Packing = " & Round(IIf(IsNull(Master.Fields("Other Charges")), 0, IIf(IsNull(Master.Fields("Other Charges")), 0, Val(Mid(Master.Fields("Other Charges"), 4, Len(Master.Fields("Other Charges")) - 3)))), 2) & ", " & _
                                        "TOT_Per = 0, TOT_Amt = 0, ReSalTax_Per = 0, ReSalTax_Amt = 0, total_amt = " & Round(Val(Format(Mid(Master.Fields("Spares Invoice Amount"), 4, Len(Master.Fields("Spares Invoice Amount")) - 3), "0.00")), 0) & ", " & _
                                        "Rounded = " & Round(Round(Val(Format(Mid(Master.Fields("Spares Invoice Amount"), 4, Len(Master.Fields("Spares Invoice Amount")) - 3), "0.00")), 0) - tt, 2) & ", " & _
                                        "Det_Tax = " & IIf(PubBackEnd = "S", Abs(CInt(mTaxDetail)), mTaxDetail) & ", AcPosting_YN = 1, U_Name = 'Siebel', U_EntDt = " & ConvertDate(Format(PubLoginDate, "Short Date")) & ", U_AE = 'E' " & _
                                        "WHERE Job_DocId='" & mJobCardID & "'"
                                                                                                    
            GCn.Execute ("Update Job_Card set JobCloseDate=" & cIIF("TempCloseDate Is Null", ConvertDate(MakeDate(Master.Fields("Invoice_Date"))), "TempCloseDate") & ",JobComp_Dt_Time=" & cIIF("TempCloseDate Is Null", ConvertDate(MakeDate(Master.Fields("Invoice_Date"))), "TempCloseDate") & _
                ",CrMemo=" & IIf(Master.Fields("Mode of Payment") = "CREDIT", 1, 0) & ",BillingName='" & IIf(Len(mPartyName) > 40, left(Replace(mPartyName, ".", ""), 40), mPartyName) & "',DelBy=RecBy_Mechanic" & _
                ",DocId_InvSpr='" & mSpareDocID & "',DocId_InvLab='" & mLabourDocID & "',GP_NO='" & mGatePass & _
                "',LabAmt_TB=" & mLabAmtTB & ",LabAmt_TP=" & mLabAmtTP & ",Lab_D_Amt= " & mLabDiscAmt & ",LabD_Per= " & mLabDiscPer & ",Lab_TaxPer=" & mLabTaxPer & ",Lab_TaxAmt= " & mLabTaxAmt & _
                ",Lab_RoundOff= " & mLabRounded & ",NetLab_Amt= " & mLabNetAmt & _
                ",ClosedU_Name='Siebel',ClosedU_EntDt=" & ConvertDate(PubLoginDate) & ",ClosedU_AE='E' where Job_Card.DocId='" & mJobCardID & "'")
            
            mEditFlag = False
        Else
            With RsNew
                .AddNew
                !DocId = mSpareDocID
                !DocIDHelp = Replace(mSpareDocID, " ", "")
                !Site_Code = mRecordSite
                !V_Type = Trim(mSpareType)
                !V_No = Right("00000000" & Val(Right(Master.Fields("Invoice_No"), 5)), 8)
                !V_DATE = MakeDate(Master.Fields("Invoice_Date"))
                !Party_Code = mPartyCode
                !Cash_Credit = Master.Fields("Mode Of Payment")
                !Party_Name = IIf(Len(mPartyName) > 40, left(Replace(mPartyName, ".", ""), 40), mPartyName)
                !L_C = "L"
                !Form_Code = mFormCode
                !CrAc = mCreditAc
                !SiebelDocID = Master.Fields("Invoice_No")
                !Job_DocId = mJobCardID
                !PType = "General"
                !GP_No = mGatePass
                !GP_Date = MakeDate(Master.Fields("Invoice_Date"))
                
                !SprAmt_MRP_TB = mSprAmt_MRP_TB
                !SprAmt_MRP_TP = mSprAmt_MRP_TP
                !OilAmt_MRP_TB = mOilAmt_MRP_TB
                !OilAmt_MRP_TP = mOilAmt_MRP_TP
                !D_Per_MRP_TB = mD_Per_MRP_TB
                !D_Per_MRP_TP = mD_Per_MRP_TP
                !D_Amt_MRP_TB = mD_Amt_MRP_TB
                !D_Amt_MRP_TP = mD_Amt_MRP_TP
                
                !SprAmt_TB = mSprAmt_TB
                !SprAmt_TP = mSprAmt_TP
                !OilAmt_TB = mOilAmt_TB
                !OilAmt_TP = mOilAmt_TP
                !D_Per_TB = mD_Per_TB
                !D_Per_TP = mD_Per_TP
                !D_Amt_TB = mD_Amt_TB
                !D_Amt_TP = mD_Amt_TP
                
                !Addition = 0
                
                !Tax_Amt = mTempTaxAmt
                !Packing = Round(IIf(IsNull(Master.Fields("Other Charges")), 0, IIf(IsNull(Master.Fields("Other Charges")), 0, Val(Mid(Master.Fields("Other Charges"), 4, Len(Master.Fields("Other Charges")) - 3)))), 2)
                
                !TOT_Per = 0
                !TOT_Amt = 0
                
                !ReSalTax_Per = 0
                !ReSalTax_Amt = 0
                
                !total_amt = Round(Val(Format(Mid(Master.Fields("Spares Invoice Amount"), 4, Len(Master.Fields("Spares Invoice Amount")) - 3), "0.00")), 0)
                
                
                tt = mSprAmt_MRP_TB + mSprAmt_MRP_TP + mOilAmt_MRP_TB + mOilAmt_MRP_TP + mSprAmt_TB + mSprAmt_TP + mOilAmt_TB + mOilAmt_TP + mTempTaxAmt + Round(IIf(IsNull(Master.Fields("Other Charges")), 0, IIf(IsNull(Master.Fields("Other Charges")), 0, Val(Mid(Master.Fields("Other Charges"), 4, Len(Master.Fields("Other Charges")) - 3)))), 2) - (mD_Amt_TP + mD_Amt_TB)
                
                !Rounded = Round(Round(Val(Format(Mid(Master.Fields("Spares Invoice Amount"), 4, Len(Master.Fields("Spares Invoice Amount")) - 3), "0.00")), 0) - tt, 2)
                
                !Det_Tax = IIf(PubBackEnd = "S", Abs(CInt(mTaxDetail)), mTaxDetail)
                !AcPosting_YN = 1
                
                !U_Name = "Siebel"
                !U_EntDt = Format(PubLoginDate, "Short Date")
                !U_AE = "A"
                .Update
                
            End With
            
            'Update JobCard for JobClose Information
                    
            GCn.Execute ("Update Job_Card set JobCloseDate=" & cIIF("TempCloseDate Is Null", ConvertDate(MakeDate(Master.Fields("Invoice_Date"))), "TempCloseDate") & ",JobComp_Dt_Time=" & cIIF("TempCloseDate Is Null", ConvertDate(MakeDate(Master.Fields("Invoice_Date"))), "TempCloseDate") & _
                ",CrMemo=" & IIf(Master.Fields("Mode of Payment") = "CREDIT", 1, 0) & ",BillingName='" & IIf(Len(mPartyName) > 40, left(Replace(mPartyName, ".", ""), 40), mPartyName) & "',DelBy=RecBy_Mechanic" & _
                ",DrSpr_AcCode='" & mPartyCode & "',DrLab_AcCode='" & mPartyCode & "',DocId_InvSpr='" & mSpareDocID & "',DocId_InvLab='" & mLabourDocID & "',GP_NO='" & mGatePass & _
                "',LabAmt_TB=" & mLabAmtTB & ",LabAmt_TP=" & mLabAmtTP & ",Lab_D_Amt= " & mLabDiscAmt & ",LabD_Per= " & mLabDiscPer & ",Lab_TaxPer=" & mLabTaxPer & ",Lab_TaxAmt= " & mLabTaxAmt & _
                ",Lab_RoundOff= " & mLabRounded & ",NetLab_Amt= " & mLabNetAmt & _
                ",ClosedU_Name='Siebel',ClosedU_EntDt=" & ConvertDate(PubLoginDate) & ",ClosedU_AE='A' where Job_Card.DocId='" & mJobCardID & "'")
        End If
        GCn.Execute ("Update Job_Card Set DrLab_AcCode=DrSpr_AcCode Where DrLab_AcCode Is Null Or DrLab_AcCode=''")
'        If PubBackEnd = "A" Then
'            GCn.Execute ("Update Job_Card, HisCard Set BillingName=" & xIsNull("H.Name", "Cash") & " Where BillingName='' Or BillingName Is Null")
'        Else
'            GCn.Execute ("Update Job_Card Set BillingName=" & xIsNull("H.Name", "Cash") & " From HisCard Where HisCard.CardNo=Job_Card.CardNo And (BillingName='' Or BillingName Is Null)")
'        End If
DuplicateSkipped:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
    
MyNextRecord:
        Master.MoveNext
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone

lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    mSrl = 0
    ErrorCnt = ErrorCnt + mSrl
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Stock Transfer (Inward)','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub


Private Sub JobCardLabourUpdate(Index)
'' On Error GoTo Eloop
Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String
Dim mDocNumber As String, mSrvCode As String, mPrefix As String
Dim mBookSite As String, mBookDiv As String, mBookNo As String
Dim mNewCard As Boolean, mCardNo As String, mLabDesc As String
Dim mSupervisor As String, mMechanic As String, mLabRate As Double
Dim mModel As String, mSellingDealer As String, mSrl As Integer
Dim mJobCardID As String, mBillTo As String, mLabNature As String, mLabType As String
Dim SkipJobcard As Boolean
Dim RsEmp As adodb.Recordset
Dim mDiscount As Double
Dim mOtherCharge As Double
Dim mResponce
Dim mOverWrite

Dim mEmpCode$
    
    mOverWrite = MsgBox("Do You Want To OverWrite If Record Exist?", vbYesNo)
    
    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Job_Lab", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsNew1 = New adodb.Recordset
    RsNew1.CursorLocation = adUseClient
    RsNew1.Open "Select * from Job_Lab2", GCn, adOpenDynamic, adLockOptimistic
   
    Set rsTemp = New adodb.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open "Select * from Job_Card", GCn, adOpenDynamic, adLockOptimistic
   
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    mLabRate = ErrorGCN.Execute("Select LabourRate from Enviro").Fields(0).Value
    Do Until Master.EOF
    
        mDocNumber = StringPass(Master.Fields("Job Card #"))
        
        If mDocNumber = "JC-UjwaAu-DH-0708-000530" Then
            MsgBox ""
        End If
        
        If Trim(mDocNumber) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Job Card #"), "Jobcard Labour Entry", "Job Card # Field is empty")
            GoTo MyNextRecord
        End If
        
        If GCn.Execute("Select DocID from Job_Card where SiebelDocID='" & mDocNumber & "'").RecordCount = 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Labour Entry", "Job Card # not found in Automan")
            GoTo MyNextRecord
        Else
            mJobCardID = GCn.Execute("Select DocID from Job_Card where SiebelDocID='" & mDocNumber & "'").Fields(0).Value
        End If

        
        mRecordDiv = left(mJobCardID, 1)
        mRecordSite = Mid(mJobCardID, 2, 2)
        
        If GCn.Execute("Select Job_DocID from Job_lab where Job_DocID='" & mJobCardID & "'").RecordCount > 0 Then
            If mOverWrite = vbYes Then
                GCn.Execute "Delete From Job_Lab Where Job_DocId='" & mJobCardID & "'"
                GCn.Execute "Delete From Job_Lab2 Where Job_DocId='" & mJobCardID & "'"
            Else
                GoTo DuplicateSkipped
            End If
        End If
        
        mSrl = 1
        Do While mDocNumber = Master.Fields("Job Card #")
            
            If IsNull(StringPass(Master.Fields("Job Code"))) Or StringPass(Master.Fields("Job Code")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Labour Entry", "Labour Code field is blank")
                GoTo LineRecordSkip
            End If
        
'            mMechanic = ""
'            If IsNull(StringPass(Master.Fields("Performed By"))) Or StringPass(Master.Fields("Performed By")) = "" Then
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Labour Entry", "Performed By Field is Blank [But Labour added to Automan Data]")
'                mMechanic = ErrorGCN.Execute("Select UnknownMechanic from Enviro").Fields(0).Value
'            Else
'                If GCn.Execute("Select Emp_Code from Emp_Mast where Emp_Name='" & Master.Fields("Performed By") & "'").RecordCount = 0 Then
'                    If GCn.Execute("Select Emp_Code from Emp_Mast where Reference='" & Master.Fields("Performed By") & "'").RecordCount = 0 Then
'                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Labour Entry", "Performed By Name not found in Staff/Employee Master [But Labour added to Automan Data]")
'                        mMechanic = ErrorGCN.Execute("Select UnknownMechanic from Enviro").Fields(0).Value
'                    Else
'                        mMechanic = GCn.Execute("Select Emp_Code from Emp_Mast where Reference='" & Master.Fields("Performed By") & "'").Fields(0).Value
'                    End If
'                Else
'                    mMechanic = GCn.Execute("Select Emp_Code from Emp_Mast where Emp_Name='" & Master.Fields("Performed By") & "'").Fields(0).Value
'                End If
'            End If
        
        
            mMechanic = ""
            If IsNull(StringPass(Master.Fields("Performed By"))) Or StringPass(Master.Fields("Performed By")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Labour Entry", XNull(Master.Fields("Performed By")) & " Performed By Field is Blank [But Labour added to Automan Data]")
                mMechanic = ErrorGCN.Execute("Select UnknownMechanic from Enviro").Fields(0).Value
            Else
                Set RsEmp = GCn.Execute("Select Emp_Code,Emp_Name, Reference from Emp_Mast ")
                If RsEmp.RecordCount > 0 Then
                    RsEmp.Find "Emp_Name Like '" & Master.Fields("Performed By") & "*" & "'"
                    If RsEmp.EOF = False And RsEmp.BOF = False Then
                        Debug.Print RsEmp!emp_name
                        mMechanic = XNull(RsEmp!Emp_Code)
                    Else
                        RsEmp.MoveFirst
                        RsEmp.Find "Reference Like '" & Master.Fields("Performed By") & "*" & "'"
                        If RsEmp.EOF = False And RsEmp.BOF = False Then
                            Debug.Print RsEmp!emp_name
                            mMechanic = XNull(RsEmp!Emp_Code)
                        Else
                            mResponce = MsgBox(" " & XNull(Master.Fields("Performed By")) & "  Employee Does Not Exist In AutoMan Do You Want to Create It Now?", vbYesNoCancel)
                            If mResponce = vbYes Then
                                mEmpCode = GCn.Execute("Select " & cVal(xIsNull("Max(" & cVal("Emp_Code") & ")", "0")) & " +1  From Emp_Mast").Fields(0)
                                GCn.Execute "Insert Into Emp_Mast (Emp_Code, Site_Code, Emp_Name, Emp_Type, Designation, ServStatus, Course, U_Name, U_EntDt, U_AE) " & _
                                            "Values ('" & mEmpCode & "', '" & PubSiteCode & "', '" & XNull(Master.Fields("Performed By")) & "',1,'MECHANIC','1', 'B', 'Siebel', " & ConvertDate(PubLoginDate) & ", 'A')"
                                mMechanic = mEmpCode
                            ElseIf mResponce = vbNo Then
                                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Labour Entry", XNull(Master.Fields("Performed By")) & " Performed By Name not found in Staff/Employee Master [But Labour added to Automan Data]")
                                mMechanic = ErrorGCN.Execute("Select UnknownMechanic from Enviro").Fields(0).Value
                            Else
                                GCn.RollbackTrans
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        
        
        
            If IsNull(StringPass(Master.Fields("Billing Type"))) Or StringPass(Master.Fields("Billing Type")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Labour Entry", "Billing Type field is blank [Defailt type PAID assumed]")
                mLabType = "PAID"
            Else
                mLabType = left(Master.Fields("Billing Type"), InStr(1, StringPass(Master.Fields("Billing Type")), "-") - 1)
            End If
            
            If IsNull(StringPass(Master.Fields("Remarks"))) Or StringPass(Master.Fields("Remarks")) = "" Then
                If GCn.Execute("select Lab_Desc from Labour where Lab_Code='" & StringPass(Master.Fields("Job Code")) & "'").RecordCount > 0 Then
                    mLabDesc = GCn.Execute("select Lab_Desc from Labour where Lab_Code='" & StringPass(Master.Fields("Job Code")) & "'").Fields(0).Value
                Else
                    mLabDesc = ""
                End If
            Else
                mLabDesc = StringPass(Master.Fields("Remarks"))
            End If
            
            If mLabType = "FREESERVICE" Then
                mLabNature = "F"
                mBillTo = IIf(mRecordDiv = "P", "O", "M")
            ElseIf mLabType = "WARRANTY" Then
                mLabNature = "W"
                mBillTo = "M"
            ElseIf mLabType = "PAID" Then
                mLabNature = "C"
                mBillTo = "C"
            ElseIf mLabType = "PDI" Then
                mLabNature = "P"
                mBillTo = IIf(mRecordDiv = "P", "S", "M")
            ElseIf mLabType = "FOC" Then
                mLabNature = "C"
                mBillTo = "S"
            ElseIf mLabType = "AMC" Then
                mLabNature = "A"
                mBillTo = "M"
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Labour Entry", "Billing Type is not known : " & Master.Fields("Billing Type"))
                GoTo LineRecordSkip
            End If
        
                    
            'Insert JobCard Labour
            With RsNew
                .AddNew
                !Job_DocId = mJobCardID
                !Site_Code = mRecordSite
                !S_No = mSrl
                !Lab_Code = Master.Fields("Job Code")
                !Mech_Voice = left(mLabDesc, 40)
                !Tax_YN = IIf(mRecordDiv = "C", 0, 1)
                !Major_YN = 0
                !Chrg_Type = mLabNature
                !Chrg_From = mBillTo
                
                If UTrim(Master.Fields("Discount")) <> "" Then
                    mDiscount = Val(Right(Replace(Master.Fields("Discount"), ",", ""), Len(Replace(Master.Fields("Discount"), ",", "")) - 3))
                Else
                    mDiscount = 0
                End If
                
                If UTrim(Master.Fields("Other Charge")) <> "" Then
                    mOtherCharge = Val(Right(Replace(Master.Fields("Other Charge"), ",", ""), Len(Replace(Master.Fields("Other Charge"), ",", "")) - 3))
                Else
                    mOtherCharge = 0
                End If
                
                
                If mLabNature = "W" Then
                    !Hrs_Taken = 0
                    !Lab_Rate = 0
                    !Hrs_War = VNull(Master.Fields("Billing Hours"))
                    If VNull(Master.Fields("Billing Hours")) > 0 Then
                        !War_Lab_Rate = Round((VNull(Master.Fields("Job Value")) - mDiscount) / Master.Fields("Billing Hours"), 2)
                    Else
                        !War_Lab_Rate = 0
                    End If
                    !LabourAmt = VNull(Master.Fields("Job Value")) - mDiscount
                Else
                    !Hrs_Taken = Master.Fields("Billing Hours")
                    If VNull(Master.Fields("Billing Hours")) > 0 Then
                        !Lab_Rate = Round((VNull(Master.Fields("Job Value")) - mDiscount) / Master.Fields("Billing Hours"), 2)
                    Else
                        !Lab_Rate = 0
                    End If
                    !Hrs_War = 0
                    !War_Lab_Rate = 0
                    !LabourAmt = VNull(Master.Fields("Job Value")) - mDiscount + mOtherCharge
                End If
                
                
                
                !U_Name = "Siebel"
                !U_EntDt = Format(PubLoginDate, "Short Date")
                !U_AE = "A"
                .Update
                '.CancelUpdate
            End With

            'Insert Jobcard Labour Mechanic Data
            With RsNew1
                .AddNew
                !Job_DocId = mJobCardID
                !Site_Code = mRecordSite
                !S_No = mSrl
                !Lab_Code = Master.Fields("Job Code")
                !Mech_Code = mMechanic
                !U_Name = "Siebel"
                !U_EntDt = Format(PubLoginDate, "Short Date")
                !U_AE = "A"
                .Update
            End With
            mSrl = mSrl + 1
            
            CopyCnt = CopyCnt + 1
            lblRecCopy(Index).Caption = CopyCnt
            lblRecCopy(Index).Refresh
    
LineRecordSkip:
            Master.MoveNext
            If Master.EOF = True Then GoTo SendToLOOP
        Loop
        GoTo SendToLOOP
DuplicateSkipped:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh

MyNextRecord:
        Master.MoveNext

SendToLOOP:
        If Master.EOF = True Then Exit Do
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone

lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    mSrl = 0
    ErrorCnt = ErrorCnt + mSrl
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Stock Transfer (Inward)','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub


Sub Job_Req_Update(Index)

    '' On Error GoTo DispError
    
    Dim mV_No As Double
    Dim mV_Type$, mV_Prefix$, mDocId$
    Dim mJobDocId$, mJobDocId_Sbl$, mDiv_Code$, mSite_Code$, mMechCode$
    Dim rsTemp As adodb.Recordset
    Dim i As Integer
    Dim mLubCat$, mPurpose$
    Dim mMrp As Double
    Dim mTaxPer As Double
    Dim mDisPer As Double
    Dim mDisAmt As Double
    Dim mDataOk As Boolean
    
    
    
    mV_Type = "W_RG"
    mV_Prefix = "SBL"
    mV_No = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from Sp_Stock where Left(DocID,1)='" & PubDivCode & "' and " & cMID("DocID", "2", "1") & "='" & PubSiteCode & "' and V_Type='" & mV_Type & "'").Fields(0).Value
    
    GCn.BeginTrans
    With Master
        If Master.RecordCount > 0 Then
            Do Until Master.EOF
                               
                
                If UTrim(XNull(.Fields("Order Number"))) = "" Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Order Number"), "Requisition Entry", "Order Number Is Blank In Excel File")
                    GoTo NextRecord
                Else
                    mJobDocId_Sbl = Master.Fields("Order Number")
                End If
                
                
                Set rsTemp = GCn.Execute("Select DocId From Job_Card Where SiebelDocId='" & mJobDocId_Sbl & "'")
                If rsTemp.RecordCount > 0 Then
                    mJobDocId = rsTemp!DocId
                Else
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Order Number"), "Requisition Entry", "Job Card Not Found In Automan")
                    GoTo NextRecord
                End If
                
                                                                
                Set rsTemp = GCn.Execute("Select Job_DocId From Sp_Stock Where SiebelDocId = '" & mJobDocId_Sbl & "'")
                If rsTemp.RecordCount > 0 Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, XNull(rsTemp!Job_DocId), "Requisition Entry", "Requision Already Found In AutoMan")
                    GoTo NextRecord
                End If
                                                                
                Set rsTemp = GCn.Execute("SELECT JobCloseDate From Job_Card Where DocId='" & mJobDocId & "' And JobCloseDate Is Not Null")
                If rsTemp.RecordCount > 0 Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Order Number"), "Requisition Entry", "Job Card Is Already Closed")
                    GoTo NextRecord
                End If
                
                
                Set rsTemp = ErrorGCN.Execute("Select AutoManDiv, AutoManSite From SiteDivision Where SiebelDiv='" & XNull(.Fields("Division")) & "'")
                If rsTemp.RecordCount > 0 Then
                    mDiv_Code = rsTemp!AutomanDiv
                    mSite_Code = rsTemp!AutomanSite
                Else
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, XNull(.Fields("Division")), "Requisition Entry", "Division Not Difined In SiteDivision Table Of AutoManSieble.MDB")
                    GoTo NextRecord
                End If
                
                
                Set rsTemp = ErrorGCN.Execute("Select UnknownMechanic From Enviro")
                If rsTemp.RecordCount > 0 Then
                    mMechCode = XNull(rsTemp!UnknownMechanic)
                End If
                
                
                
                mDocId = mDiv_Code & mSite_Code & mSite_Code & " " & mV_Type & "SBL" & Format(Master!Date, "yy") & Right("00000000" & mV_No, 8)
                i = 1
                    
                Do While .Fields("Order Number") = mJobDocId_Sbl
                    
                    Set rsTemp = GCn.Execute("Select Part_No, Part_Grade, MRP From Part Where Part_No='" & .Fields("Part No") & "'")
                    If rsTemp.RecordCount = 0 Then
                        GCn.Execute "Insert Into Part Values(Part_No, Div_Code, Site_Code, Part_Name, Part_Grade, TB_SRate )" & _
                                    "Values('" & .Fields("Part No") & "', '" & PubDivCode & "', '" & PubSiteCode & "', '" & .Fields("Part Desc") & "','S', " & VNull(.Fields("Rate")) & ")"
                        mLubCat = ""
                        mMrp = 0
                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, XNull(.Fields("Part No")), "Requisition Entry", "Part No Not Found In Automan. Part Opened In Part Master and Data Imported")
                    Else
                        If rsTemp!Part_Grade = "L" Then
                            mLubCat = "N"
                        Else
                            mLubCat = ""
                        End If
                        
                        mMrp = rsTemp!MRP
                    End If
                    
                    
                    Select Case UTrim(.Fields("Billing Type"))
                        Case "PAID"
                            mPurpose = "C"
                        Case "WARRANTY"
                            mPurpose = "W"
                        Case "SELF"
                            mPurpose = "F"
                    End Select
                    
                    If VNull(.Fields("Value")) <> 0 Then
                        mTaxPer = VNull(.Fields("Tax Amount After Discount")) * 100 / VNull(.Fields("Value"))
                    Else
                        mTaxPer = 0
                    End If
                    
                    
                    
                    mDisAmt = ((VNull(.Fields("Rate")) * VNull(.Fields("Sold Qty"))) - VNull(.Fields("Value")))
                    If VNull(.Fields("Value")) <> 0 Then
                        mDisPer = mDisAmt * 100 / VNull(.Fields("Value"))
                    Else
                        mDisPer = 0
                    End If
                    
                    
                    If Round(mTaxPer) > 3 And Round(mTaxPer) < 5 Then
                        mTaxPer = 4
                    ElseIf mTaxPer = 0 Then
                        mTaxPer = 0
                    ElseIf Round(mTaxPer) > 14 Then
                        mTaxPer = 15
                    Else
                        mTaxPer = 12.5
                    End If
                                        
                                        
                    
                    GCn.Execute "Insert Into Sp_Stock (DocId, Srl_No, V_No, Site_Code, V_Type, V_Date, Job_DocId, " & _
                                "Job_DivCode, Mech_Code, Part_No, Lub_Category, Godown, " & _
                                "Qty_Doc, Qty_Iss, Tax_Yn, Mrp_Yn, Rate, Mrp_Rate, " & _
                                "Disc_Per, Disc_Amt, Amount, Net_Amt, Purpose, V_Rate, " & _
                                "TaxPer, TaxAmt, SiebelDocId, U_EntDt, U_Name, U_AE) Values ( " & _
                                "'" & mDocId & "', " & i & ", " & mV_No & ", '" & mSite_Code & "', '" & mV_Type & "', " & ConvertDate(.Fields("Date")) & ", '" & mJobDocId & "', " & _
                                "'" & mDiv_Code & "', '" & mMechCode & "', '" & .Fields("Part No") & "', '" & mLubCat & "', '" & PubSprWorksGodown & "', " & _
                                "" & Val(.Fields("Sold Qty")) & "," & Val(.Fields("Sold Qty")) & ", 1, 0, " & Val(.Fields("Rate")) & ", " & mMrp & ", " & _
                                "" & mDisPer & ", " & mDisAmt & ", " & Val(.Fields("Value")) & ", " & Val(.Fields("Value")) & ",'" & mPurpose & "', " & Val(.Fields("Rate")) & ", " & _
                                "" & mTaxPer & ", " & Val(.Fields("Tax Amount After Discount")) & ", '" & mJobDocId_Sbl & "', " & ConvertDate(PubLoginDate) & ", 'Siebel', 'A')"
                    .MoveNext
                    i = i + 1
                Loop
                
                mV_No = mV_No + 1
                
                
NextRecord:
                .MoveNext
            Loop
        End If
    End With
    GCn.CommitTrans
    MsgBox "Updation Completed"
DispError:
    MsgBox err.Description
    Resume Next
End Sub




Private Sub JobCardEntryUpdate(Index)
'' On Error GoTo Eloop
Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String
Dim mDocNumber As String, mSrvCode As String, mPrefix As String
Dim mBookSite As String, mBookDiv As String, mBookNo As String
Dim mNewCard As Boolean, mCardNo As String
Dim mSupervisor As String, mMechanic As String, mChassis As String, mRegNo As String
Dim mModel As String, mSellingDealer As String, mSrl As Integer
    
    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Job_Card", GCn, adOpenDynamic, adLockOptimistic
    
    
    Set RsNew1 = New adodb.Recordset
    RsNew1.CursorLocation = adUseClient
    RsNew1.Open "Select * from Job_Demand", GCn, adOpenDynamic, adLockOptimistic
   
    Set rsTemp = New adodb.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open "Select * from Hiscard", GCn, adOpenDynamic, adLockOptimistic
   
   
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    Do Until Master.EOF
        mDocNumber = StringPass(Master.Fields("Job Card #"))
        
        If Trim(mDocNumber) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Job Card #"), "JobCard Entry", "Job Card # Field is empty")
            GoTo MyNextRecord
        End If
            
        
        If GCn.Execute("Select Job_no from Job_Card where SiebelDocID='" & mDocNumber & "'").RecordCount > 0 Then
            GCn.Execute "Update Job_Card Set Job_Date=" & ConvertDate(MakeDate(left(Master.Fields("Created Date Time"), 10))) & ", " & _
                            "ArrivalTime=" & IIf(PubBackEnd = "A", "#", "'") & MakeDate(left(Master.Fields("Created Date Time"), 10)) & " " & Format(Master.Fields("Created Date Time"), "hh:mm:ss") & IIf(PubBackEnd = "A", "#", "'") & ", " & _
                            "Recp_Time=" & IIf(PubBackEnd = "A", "#", "'") & MakeDate(left(Master.Fields("Created Date Time"), 10)) & " " & Format(Master.Fields("Created Date Time"), "hh:mm:ss") & IIf(PubBackEnd = "A", "#", "'") & ", " & _
                            "BillingName = '" & left(Master.Fields("First Name") & Master.Fields("Last Name"), 40) & "' " & _
                            "Where SiebelDocId='" & mDocNumber & "'"
            If IsNull(Master.Fields("Closed Date Time")) = False Then
                GCn.Execute "Update Job_Card Set JobCloseDate=" & IIf(PubBackEnd = "A", "#", "'") & MakeDate(left(Master.Fields("Closed Date Time"), 10)) & " " & Format(Master.Fields("Closed Date Time"), "hh:mm:ss") & IIf(PubBackEnd = "A", "#", "'") & ", " & _
                                "JobComp_Dt_Time=" & IIf(PubBackEnd = "A", "#", "'") & MakeDate(left(Master.Fields("Closed Date Time"), 10)) & " " & Format(Master.Fields("Closed Date Time"), "hh:mm:ss") & IIf(PubBackEnd = "A", "#", "'") & " " & _
                                "Where SiebelDocId='" & mDocNumber & "' And JobCloseDate Is Not Null"
            End If
            
            GoTo DuplicateSkipped
        End If
        
        If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").RecordCount > 0 Then
            mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").Fields(0).Value
            mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").Fields(0).Value
            mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").Fields(0).Value
        Else
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Automan Site/Division is not Defined in SiteDivision Table for this Group Value")
            GoTo MyNextRecord
        End If
        
        If IsNull(StringPass(Master.Fields("Created Date Time"))) Or StringPass(Master.Fields("Created Date Time")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Created Date Time is Empty")
            GoTo MyNextRecord
        End If
        
        mBookSite = ""
        mBookDiv = ""
        mBookNo = ""
        mChassis = ""
        mRegNo = ""
        If GCn.Execute("Select Book_No from Job_booking where SiebelDocID='" & Master.Fields("Service Request No") & "'").RecordCount > 0 Then
            mBookSite = GCn.Execute("Select Site_Code from Job_booking where SiebelDocID='" & Master.Fields("Service Request No") & "'").Fields(0).Value
            mBookDiv = GCn.Execute("Select Div_Code from Job_booking where SiebelDocID='" & Master.Fields("Service Request No") & "'").Fields(0).Value
            mBookNo = GCn.Execute("Select Book_No from Job_booking where SiebelDocID='" & Master.Fields("Service Request No") & "'").Fields(0).Value
            mModel = GCn.Execute("Select Model from Job_booking where SiebelDocID='" & Master.Fields("Service Request No") & "'").Fields(0).Value
            mSellingDealer = XNull(GCn.Execute("Select SellingDealerCode from Job_booking where SiebelDocID='" & Master.Fields("Service Request No") & "'").Fields(0).Value)
        Else
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Service Request (Job Booking) Data not found in Automan for this Jobcard")
            GoTo MyNextRecord
        End If
        
        mNewCard = False
        mCardNo = ""
        
        If IsNull(StringPass(Master.Fields("Chassis No"))) Or StringPass(Master.Fields("Chassis No")) = "" Then
            mChassis = GCn.Execute("Select Chassis from Job_booking where SiebelDocID='" & Master.Fields("Service Request No") & "'").Fields(0).Value
        Else
            mChassis = XNull(Master.Fields("Chassis No"))
        End If
        
        If IsNull(StringPass(Master.Fields("Vehicle Registration Number"))) Or StringPass(Master.Fields("Vehicle Registration Number")) = "" Then
            mRegNo = GCn.Execute("Select RegNo from Job_booking where SiebelDocID='" & Master.Fields("Service Request No") & "'").Fields(0).Value
        Else
            mRegNo = XNull(Master.Fields("Vehicle Registration Number"))
        End If
        
        If IsNull(mRegNo) Or StringPass(mRegNo) = "" Then
            If IsNull(mChassis) Or StringPass(mChassis) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Vehicle Chassis No & Registration Number is Blank")
                GoTo MyNextRecord
            Else
                If GCn.Execute("Select CardNo from Hiscard where chassis='" & mChassis & "'").RecordCount = 0 Then
                    mNewCard = True
                Else
                    If GCn.Execute("Select Div_Code from Hiscard where Chassis='" & mChassis & "'").Fields(0).Value <> mRecordDiv Then
                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Division Code is Diff. in Hiscard for Chassis No. " & mChassis)
                        GoTo MyNextRecord
                    Else
                        mNewCard = False
                        mCardNo = GCn.Execute("Select CardNo from Hiscard where Chassis='" & mChassis & "'").Fields(0).Value
                    End If
                End If
            End If
        Else
            If GCn.Execute("Select CardNo from Hiscard where RegNo='" & mRegNo & "'").RecordCount = 0 Then
                If GCn.Execute("Select CardNo from Hiscard where Chassis='" & mChassis & "'").RecordCount = 0 Or mChassis = "" Then
                    mNewCard = True
                Else
                    If GCn.Execute("Select Div_Code from Hiscard where Chassis='" & mChassis & "'").Fields(0).Value <> mRecordDiv Then
                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Division Code is Diff. in Hiscard for chassis No. " & mChassis)
                        GoTo MyNextRecord
                    Else
                        mNewCard = False
                        mCardNo = GCn.Execute("Select CardNo from Hiscard where chassis='" & mChassis & "'").Fields(0).Value
                    End If
                End If
            Else
                If GCn.Execute("Select Div_Code from Hiscard where RegNo='" & mRegNo & "'").Fields(0).Value <> mRecordDiv Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Division Code is Diff. in Hiscard for Registration No. " & mRegNo)
                    GoTo MyNextRecord
                Else
                    mNewCard = False
                    mCardNo = GCn.Execute("Select CardNo from Hiscard where RegNo='" & mRegNo & "'").Fields(0).Value
                End If
            End If
        End If
        
'        If IsNull(mChassis) Or StringPass(mChassis) = "" Then
'            If IsNull(mRegNo) Or StringPass(mRegNo) = "" Then
'                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Vehicle Chassis No & Registration Number is Blank")
'                GoTo MyNextRecord
'            Else
'                If GCn.Execute("Select CardNo from Hiscard where RegNo='" & mRegNo & "'").RecordCount = 0 Then
'                    mNewCard = True
'                Else
'                    If GCn.Execute("Select Div_Code from Hiscard where RegNo='" & mRegNo & "'").Fields(0).Value <> mRecordDiv Then
'                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Division Code is Diff. in Hiscard for Registration No. " & mRegNo)
'                        GoTo MyNextRecord
'                    Else
'                        mNewCard = False
'                        mCardNo = GCn.Execute("Select CardNo from Hiscard where RegNo='" & mRegNo & "'").Fields(0).Value
'                    End If
'                End If
'            End If
'        Else
'            If GCn.Execute("Select CardNo from Hiscard where Chassis='" & mChassis & "'").RecordCount = 0 Then
'                If GCn.Execute("Select CardNo from Hiscard where RegNo='" & mRegNo & "'").RecordCount = 0 Or mRegNo = "" Then
'                    mNewCard = True
'                Else
'                    If GCn.Execute("Select Div_Code from Hiscard where RegNo='" & mRegNo & "'").Fields(0).Value <> mRecordDiv Then
'                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Division Code is Diff. in Hiscard for Registration No. " & mRegNo)
'                        GoTo MyNextRecord
'                    Else
'                        mNewCard = False
'                        mCardNo = GCn.Execute("Select CardNo from Hiscard where RegNo='" & mRegNo & "'").Fields(0).Value
'                    End If
'                End If
'            Else
'                If GCn.Execute("Select Div_Code from Hiscard where Chassis='" & mChassis & "'").Fields(0).Value <> mRecordDiv Then
'                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Division Code is Diff. in Hiscard for Chassis No. " & mChassis)
'                    GoTo MyNextRecord
'                Else
'                    mNewCard = False
'                    mCardNo = GCn.Execute("Select CardNo from Hiscard where Chassis='" & mChassis & "'").Fields(0).Value
'                End If
'            End If
'        End If
        
        mSrvCode = ""
        If IsNull(StringPass(Master.Fields("SR Type"))) Or StringPass(Master.Fields("SR Type")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "SR Type (Service Type) is Blank")
            GoTo MyNextRecord
        Else
            If GCn.Execute("Select Serv_Type from Service_Type where Serv_Desc='" & Master.Fields("SR Type") & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Service Type not defined in Automan")
                GoTo MyNextRecord
            Else
                mSrvCode = GCn.Execute("Select Serv_Type from Service_Type where Serv_Desc='" & Master.Fields("SR Type") & "'").Fields(0).Value
            End If
        End If
            
        mMechanic = ""
        If IsNull(StringPass(Master.Fields("SR Assigned To"))) Or StringPass(Master.Fields("SR Assigned To")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "'SR Assigned To' Field is Blank")
            GoTo MyNextRecord
        Else
            If GCn.Execute("Select Emp_Code from Emp_Mast where Reference='" & Master.Fields("SR Assigned To") & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "SR Assigned To => Code not defined in Automan (In Referred Field of Emp_Mast Table)")
                GoTo MyNextRecord
            Else
                mMechanic = GCn.Execute("Select Emp_Code from Emp_Mast where Reference='" & Master.Fields("SR Assigned To") & "'").Fields(0).Value
            End If
        End If


'        mMechanic = ""
'        If IsNull(StringPass(Master.Fields("SR Assigned To"))) Or StringPass(Master.Fields("SR Assigned To")) = "" Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Entry", "SR Assigned To Field is Blank (Mechanic Name)")
'            GoTo MyNextRecord
'        Else
'            Set RsEmp = GCn.Execute("Select Emp_Code,Emp_Name, Reference from Emp_Mast ")
'            If RsEmp.RecordCount > 0 Then
'                RsEmp.Find "Emp_Name Like '" & Master.Fields("SR Assigned To") & "*" & "'"
'                If RsEmp.EOF = False And RsEmp.BOF = False Then
'                    Debug.Print RsEmp!emp_name
'                    mMechanic = XNull(RsEmp!Emp_Code)
'                Else
'                    RsEmp.MoveFirst
'                    RsEmp.Find "Reference Like '" & Master.Fields("SR Assigned To") & "*" & "'"
'                    If RsEmp.EOF = False And RsEmp.BOF = False Then
'                        Debug.Print RsEmp!emp_name
'                        mMechanic = XNull(RsEmp!Emp_Code)
'                    Else
'                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Entry", "SR Assigned To Name not found in Staff/Employee Master [Mechanic Name]")
'                        GoTo MyNextRecord
'                    End If
'                End If
'            End If
'        End If

        
        mSupervisor = ""
        If IsNull(StringPass(Master.Fields("Supervisor"))) Or StringPass(Master.Fields("Supervisor")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Supervisor Field is Blank")
            GoTo MyNextRecord
        Else
            If GCn.Execute("Select Emp_Code from Emp_Mast where Reference='" & Master.Fields("Supervisor") & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "JobCard Entry", "Supervisor Code not defined in Automan (In Referred Field of Emp_Mast Table)")
                GoTo MyNextRecord
            Else
                mSupervisor = GCn.Execute("Select Emp_Code from Emp_Mast where Reference='" & Master.Fields("Supervisor") & "'").Fields(0).Value
            End If
        End If
        
        
        
'        mSupervisor = ""
'        If IsNull(StringPass(Master.Fields("mSupervisor"))) Or StringPass(Master.Fields("mSupervisor")) = "" Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Entry", "Supervisor Field is Blank ")
'            GoTo MyNextRecord
'        Else
'            Set RsEmp = GCn.Execute("Select Emp_Code,Emp_Name, Reference from Emp_Mast ")
'            If RsEmp.RecordCount > 0 Then
'                RsEmp.Find "Emp_Name Like '" & Master.Fields("mSupervisor") & "*" & "'"
'                If RsEmp.EOF = False And RsEmp.BOF = False Then
'                    Debug.Print RsEmp!emp_name
'                    mSupervisor = XNull(RsEmp!Emp_Code)
'                Else
'                    RsEmp.MoveFirst
'                    RsEmp.Find "Reference Like '" & Master.Fields("mSupervisor") & "*" & "'"
'                    If RsEmp.EOF = False And RsEmp.BOF = False Then
'                        Debug.Print RsEmp!emp_name
'                        mSupervisor = XNull(RsEmp!Emp_Code)
'                    Else
'                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, mDocNumber, "Jobcard Entry", "Supervisor Name not found in Staff/Employee Master")
'                        GoTo MyNextRecord
'                    End If
'                End If
'            End If
'        End If
        
        
        
        '' Checking of Fields for History Card Creation Purpose
        If mNewCard Then
            'mCardNo = GCn.Execute("select iif(isnull(max(val(mid(CardNo,2 ,len(cardno)-1)))),0,Max(val(mid(CardNo,2,len(cardno)-1))))+1 from Hiscard where Site_Code='" & mRecordSite & "'").Fields(0).Value
            mCardNo = GCn.Execute("select " & vIsNull("max(" & cVal(cMID("CardNo", "2", "len(cardno)-1")) & ")", "0") & "+1 from Hiscard where Site_Code='" & mRecordSite & "'").Fields(0).Value
            mCardNo = mRecordSite & Right("0000000" & mCardNo, 7)
            With rsTemp
                .AddNew
                !CardNo = mCardNo
                !Site_Code = mRecordSite
                !Div_Code = mRecordDiv
                !CardDate = MakeDate(left(Master.Fields("Created Date Time"), 10))
                !Model = mModel
                !RegNo = left(mRegNo, 12)
                !Chassis = mChassis
                !ENGINE = ""
                !Delivery_Date = IIf(IsNull(Master.Fields("Vehicle Sale Date (Dealer)")), Master.Fields("Vehicle Sale Date (Dealer)"), MakeDate(XNull(Master.Fields("Vehicle Sale Date (Dealer)"))))
                !Dealer_Code = mSellingDealer
                !CouponNo = ""
                !Supplier_BillNo = ""
                !Supplier_BillDate = IIf(IsNull(Master.Fields("TM Invoice Date")), Master.Fields("TM Invoice Date"), MakeDate(XNull(Master.Fields("TM Invoice Date"))))
                If mRecordDiv = "C" Then
                    !Name = left(XNull(Master.Fields("Account")), 40)
                Else
                    If Trim(left(Trim(XNull(Master.Fields("First Name"))) & " " & Trim(XNull(Master.Fields("Last Name"))), 40)) = "" Then
                        !Name = left(XNull(Master.Fields("Account")), 40)
                    Else
                        !Name = left(Trim(XNull(Master.Fields("First Name"))) & " " & Trim(XNull(Master.Fields("Last Name"))), 40)
                    End If
                End If
                !Add1 = XNull(Master.Fields("Address"))
                !Add2 = XNull(Master.Fields("Site"))
                !Add3 = ""
                !PhoneOff = XNull(Master.Fields("Account Fax #"))
                !PhoneResi = XNull(Master.Fields("Account Phone #"))
                !Mobile = left(XNull(Master.Fields("Contact Phones (Res, Off, Mob)")), 10)
                !Govt_YN = 0
                
                !U_Name = "Siebel"
                !U_EntDt = Format(PubLoginDate, "Short Date")
                !U_AE = "A"
                
                .Update
                
            End With
        End If
        
        CodeCnt = Right(Master.Fields("Job Card #"), 6)
        
        Dim mShortYear As String
        If Month(Master.Fields("Created Date Time")) > 3 Then
            mShortYear = Right(Format(Master.Fields("Created Date Time"), "yy"), 1) & Right(Val(Format(Master.Fields("Created Date Time"), "yy")) + 1, 1)
        Else
            mShortYear = Right(Val(Format(Master.Fields("Created Date Time"), "yy")) - 1, 1) & Right(Format(Master.Fields("Created Date Time"), "yy"), 1)
        End If
        mPrefix = "SBL" & mShortYear 'Format(Master.Fields("Created Date Time"), "yy")

        mDocId = mRecordDiv & mRecordSite & mRecordSite & " W_JC" & mPrefix & Right("00000000" & CodeCnt, 8)
        
        'Insert JobCard
        With RsNew
            .AddNew
            !DocId = mDocId
            '!Div_Code = mRecordDiv
            !Site_Code = mRecordSite
            !Job_No = CodeCnt
            !job_Date = MakeDate(left(Master.Fields("Created Date Time"), 10))
            !Job_BookDivCode = mBookDiv
            !Job_BookNo = mBookNo
            !Job_BookSiteCode = mBookSite
            !CardNo = mCardNo
            !Govt_YN = 0
            !Serv_Type = mSrvCode
            !AtKMsHrs = VNull(Master.Fields("KMS"))
            !Fuel = ""
            !Est_SpCost = VNull(Format(Mid(Master.Fields("Effective Parts Estimate"), 4, Len(Master.Fields("Effective Parts Estimate")) - 3), "0.00"))
            !Est_LabCost = VNull(Format(Mid(Master.Fields("Effective Labour Estimate"), 4, Len(Master.Fields("Effective Labour Estimate")) - 3), "0.00"))
            
            !ArrivalTime = MakeDate(left(Master.Fields("Created Date Time"), 10)) & " " & Format(Master.Fields("Created Date Time"), "hh:mm:ss")
            !Recp_Time = MakeDate(left(Master.Fields("Created Date Time"), 10)) & " " & Format(Master.Fields("Created Date Time"), "hh:mm:ss")
            
            If Not IsNull(Master.Fields("Effective Final Delivery Estimate Date")) Then
                !ExpDelDate = MakeDate(left(Master.Fields("Effective Final Delivery Estimate Date"), 10)) & " " & Format(Master.Fields("Effective Final Delivery Estimate Date"), "hh:mm:ss")
            End If
            !Body_Damage = ""
            !OpenRemarks = ""
            !KMsHrs = "K"
            '!HrMeter = VNull(Master.Fields("KMS"))
            
            !JobType = ""
            !RecBy_Mechanic = mMechanic
            !RecBy_Supervisor = mSupervisor
            
            If IsNull(Master.Fields("Closed Date Time")) = False Then
                !TempCloseDate = MakeDate(left(Master.Fields("Closed Date Time"), 10)) & " " & Format(Master.Fields("Closed Date Time"), "hh:mm:ss")
            End If
            
            !SiebelDocID = mDocNumber

                        
            !CreatedU_Name = "Siebel"
            !CreatedU_EntDt = Format(PubLoginDate, "Short Date")
            !CreatedU_AE = "A"
            !U_Name = "Siebel"
            !U_EntDt = Format(PubLoginDate, "Short Date")
            !U_AE = "A"
            .Update
        End With

        If mNewCard = False Then
            GCn.Execute ("Update Hiscard set Name='" & left(XNull(Master.Fields("Account")), 40) & "' where CardNo='" & mCardNo & "'")
        End If
        
        '' Updation of Jobcard Information in Job_Booking Table
        GCn.Execute ("Update Job_Booking set Job_DocID='" & mDocId & "' where Div_Code='" & mBookDiv & "' and Book_No=" & mBookNo & " and Site_Code='" & mBookSite & "'")

        '' Record set for Driver complaints/Demand
        Set Master1 = CreateObject("ADODB.Recordset")
        GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] where [Job Card #]='" & Master.Fields("Job Card #") & "' Order By [Job Card #]"
        Master1.Open GSQL, ExcelGcn2, adOpenStatic

        If Master1.RecordCount > 0 Then Master1.MoveFirst
        mSrl = 1
        Do Until Master1.EOF
            'Insert Jobcard Demand
            With RsNew1
                .AddNew
                !Job_DocId = mDocId
                !S_No = mSrl
                !Code = XNull(Master1.Fields("Customer Complaint Code"))
                                
                If PubBackEnd = "S" Then
                    If IsNull(StringPass(Master1.Fields("Customer Voice"))) Or StringPass(Master1.Fields("Customer Voice")) = "" Then
                        If IsNull(StringPass(Master1.Fields("Customer Complaint Description"))) Or StringPass(Master1.Fields("Customer Complaint Description")) = "" Then
                            !Details = "- Unknown Trouble from Driver/Customer -"
                        Else
                            !Details = left(XNull(Master1.Fields("Customer Complaint Description")), 40)
                        End If
                    Else
                        !Details = left(XNull(Master1.Fields("Customer Voice")), 40)
                    End If
                Else
                    If IsNull(StringPass(Master1.Fields("Customer Complaint Description"))) Or StringPass(Master1.Fields("Customer Complaint Description")) = "" Then
                        If IsNull(StringPass(Master1.Fields("Customer Voice"))) Or StringPass(Master1.Fields("Customer Voice")) = "" Then
                            !Details = "* Unknown Trouble from Driver/Customer *"
                        Else
                            !Details = left(XNull(Master1.Fields("Customer Voice")), 40)
                        End If
                    Else
                        !Details = left(XNull(Master1.Fields("Customer Complaint Description")), 40)
                    End If
                End If
                !Site_Code = mRecordSite
                !Complaint_YN = 1
                !Repeat_YN = IIf(Master1.Fields("Repeat Complaint") = "Y", 1, 0)
                
                !U_Name = "Siebel"
                !U_EntDt = Format(PubLoginDate, "Short Date")
                !U_AE = "A"
                .Update
            End With
            mSrl = mSrl + 1
            Master1.MoveNext
        Loop

DuplicateSkipped:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh

MyNextRecord:
        Master.MoveNext
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone

lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    mSrl = 0
    ErrorCnt = ErrorCnt + mSrl
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','JobCard Entry','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub
Private Sub ServiceBookingUpdate(Index)
'' On Error GoTo Eloop
Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String
Dim mAdjSlipNo As String
Dim mSrvCode As String, mSellingDealer As String
Dim i As Integer, mFieldRename As Boolean
Dim mChecking As Boolean
    ImportBtn(Index).BackColor = ProcessColor
    
    
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Job_Booking", GCn, adOpenDynamic, adLockOptimistic
    
    mFieldRename = False
    For i = 0 To Master.Fields.Count - 1
        If UCase(Master.Fields(i).Name) = UCase("RegNo") Then
            mFieldRename = True
            Exit For
        End If
    Next
    If mFieldRename = False Then
        MsgBox "Field not found  : RegNo", vbCritical, "Field Name not changed"
        Exit Sub
    End If
    
    mFieldRename = False
    For i = 0 To Master.Fields.Count - 1
        If UCase(Master.Fields(i).Name) = UCase("Chassis No") Then
            mFieldRename = True
            Exit For
        End If
    Next
    If mFieldRename = False Then
        MsgBox "Field not found  : Chassis No", vbCritical, "Field Name not changed"
        Exit Sub
    End If
    
    mFieldRename = False
    For i = 0 To Master.Fields.Count - 1
        If UCase(Master.Fields(i).Name) = UCase("Temporary Chassis No") Then
            mFieldRename = True
            Exit For
        End If
    Next
    If mFieldRename = False Then
        MsgBox "Field not found  : Temporary Chassis No", vbCritical, "Field Name not changed"
        Exit Sub
    End If
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    Do Until Master.EOF
        mChecking = False
        mAdjSlipNo = StringPass(Master.Fields("SR #"))
        
        If Trim(mAdjSlipNo) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("SR #"), "Service Booking", "SR # Field is empty")
            GoTo MyNextRecord
        End If
        
        If GCn.Execute("Select Book_no from Job_Booking where SiebelDocID='" & mAdjSlipNo & "'").RecordCount > 0 Then
            GoTo DuplicateSkipped
        End If
        
        If IsNull(StringPass(Master.Fields("SR Created Date/Time"))) Or StringPass(Master.Fields("SR Created Date/Time")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Service Booking", "SR Created Date/Time is Empty")
            GoTo MyNextRecord
        End If
       
        If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Group")) & "'").RecordCount > 0 Then
            mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Group")) & "'").Fields(0).Value
            mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Group")) & "'").Fields(0).Value
            mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Group")) & "'").Fields(0).Value
        Else
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Service Booking", "Automan Site/Division is not Defined in SiteDivision Table for this Group Value")
            GoTo MyNextRecord
        End If
            
        mSrvCode = ""
        If IsNull(StringPass(Master.Fields("Service Type"))) Or StringPass(Master.Fields("Service Type")) = "" Then
        Else
            If GCn.Execute("Select Serv_Type from Service_Type where Serv_Desc='" & Master.Fields("Service Type") & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Service Booking", "Service Type not defined in Automan")
                GoTo MyNextRecord
            Else
                mSrvCode = GCn.Execute("Select Serv_Type from Service_Type where Serv_Desc='" & Master.Fields("Service Type") & "'").Fields(0).Value
            End If
        End If
            
        If IsNull(StringPass(Master.Fields("Product").Value)) Or StringPass(Master.Fields("Product").Value) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Service Booking", "Vehicle Product/Model field is Empty [Model Name * Unknown * Passed]")
            'GoTo MyNextRecord
        Else
            If GCn.Execute("Select Model from Model where Model='" & Master.Fields("Product") & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Product"), "Service Booking", "Vehicle Product/Model not exist in AUTOMAN Model Master [But data imported]")
                'GoTo MyNextRecord
            End If
        End If
        
        If IsNull(StringPass(Master.Fields("Selling Dealer").Value)) Or StringPass(Master.Fields("Selling Dealer").Value) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Service Booking", "Selling Dealer field is Empty [But Jobcard created in Automan]")
            mSellingDealer = ErrorGCN.Execute("Select UnknownSellingDealer From Enviro").Fields(0).Value
        Else
            If GCn.Execute("Select D_Code from AMD_Dealer where D_Name='" & left(Master.Fields("Selling Dealer"), 40) & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Selling Dealer"), "Service Booking", "Selling Dealer not exist in AUTOMAN Dealer Master")
                GoTo MyNextRecord
            Else
                mSellingDealer = GCn.Execute("Select D_Code from AMD_Dealer where D_Name='" & left(Master.Fields("Selling Dealer"), 40) & "'").Fields(0).Value
            End If
        End If
        
        CodeCnt = Val(Right(Master.Fields("SR #"), 8)) & Right(Master.Fields("SR #"), 5)
                
        'Insert New Rec
        With RsNew
            .AddNew
            !Div_Code = mRecordDiv
            !Site_Code = mRecordSite & mRecordSite
            !Book_No = CodeCnt
            !Book_Date = MakeDate(left(Master.Fields("SR Created Date/Time"), 10))
            !Name = left(StringPass(Master.Fields("Owner Account")), 40)
            '!Add1 = left(StringPass(Master.Fields("Account Street Address")), 40)
            '!Add2 = left(StringPass(Master.Fields("Account Site, City, Taluka, Dist, State, Country, PIN")), 40)
            !Add3 = StringPass(Master.Fields("Account Fax #"))
            !PhoneOff = left(StringPass(Master.Fields("Contact Phones (Res, Off, Mob)")), 25)
            !PhoneResi = StringPass(Master.Fields("Account Phone #"))
            !Mobile = left(StringPass(Master.Fields("Contact Work Phone #")), 10)
            If IsNull(StringPass(Master.Fields("Product").Value)) Or StringPass(Master.Fields("Product").Value) = "" Then
                !Model = "* Unknown *"
            Else
                !Model = StringPass(Master.Fields("Product"))
            End If
            If IsNull(Master.Fields("Chassis No")) Or StringPass(Master.Fields("Chassis No")) = "" Then
                !Chassis = left(Replace(StringPass(Master.Fields("Temporary Chassis No")), " ", ""), 15)
            Else
                !Chassis = left(StringPass(Master.Fields("Chassis No")), 15)
            End If
            !ENGINE = ""
            !RegNo = left(StringPass(Master.Fields("RegNo")), 12)
            !Advance = 0
            !ForServiceDate = IIf(Master.Fields("Booked For Date/Time") = "" Or IsNull(Master.Fields("Booked For Date/Time")), MakeDate(left(Master.Fields("SR Created Date/Time"), 10)) & " " & Format(Master.Fields("SR Created Date/Time"), "HH:MM"), MakeDate(left(XNull(Master.Fields("Booked For Date/Time")), 10)) & " " & Format(Master.Fields("Booked For Date/Time"), "HH:MM"))
            !Remarks = StringPass(Master.Fields("Summary"))
            !Service_Type = mSrvCode
            !SiebelDocID = mAdjSlipNo
            !SellingDealerCode = mSellingDealer

            !U_Name = "Siebel"
            !U_EntDt = IIf(Master.Fields("Booked For Date/Time") = "" Or IsNull(Master.Fields("Booked For Date/Time")), MakeDate(left(Master.Fields("SR Created Date/Time"), 10)) & " " & Format(Master.Fields("SR Created Date/Time"), "HH:MM"), MakeDate(left(XNull(Master.Fields("Booked For Date/Time")), 10)) & " " & Format(Master.Fields("Booked For Date/Time"), "HH:MM"))
            !U_AE = "A"
            .Update
            mChecking = True
        End With

DuplicateSkipped:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh

MyNextRecord:
        Master.MoveNext
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone

lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Service Booking','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub


Private Sub TroubleMasterUpdate(Index)
'' On Error GoTo Eloop
    
    ImportBtn(Index).BackColor = ProcessColor

    GCn.BeginTrans
    
    CopyCnt = 0
    ErrorCnt = 0
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Trouble", GCn, adOpenDynamic, adLockOptimistic
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Complaint Code"))) Or StringPass(Master.Fields("Complaint Code")) = "" Then GoTo SkipDuplicate
        
        If IsNull(StringPass(Master.Fields("Complaint Description"))) Or StringPass(Master.Fields("Complaint Description")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Complaint Code"), "Trouble Master", "Complaint/Trouble Description is Blank")
            GoTo MyNextRecord2
        End If
        
        If RsNew.RecordCount > 0 Then RsNew.MoveFirst
        Do Until RsNew.EOF
            If UCase(ReturnString(RsNew!Trouble_Code)) = UCase(ReturnString(left(StringPass(Master.Fields("Complaint Code")), 6))) Then
                GoTo SkipDuplicate
            End If
            RsNew.MoveNext
        Loop
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!Trouble_Code = Master.Fields("Complaint Code")
        RsNew!Trouble_Name = left(StringPass(Master.Fields("Complaint Description")), 40)
        RsNew!Site_Code = PubSiteCode
        RsNew!Div_Code = PubDivCode
        RsNew!TType = "Complaint"
        RsNew!Major = 0
        
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        CodeCnt = CodeCnt + 1

SkipDuplicate:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh

MyNextRecord2:
        Master.MoveNext
    Loop
    
ClearBuffer:
    GCn.CommitTrans
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone

lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & StringPass(Master.Fields("Job Code")) & "','Trouble Master','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub


Private Sub LabourMasterUpdate(Index)
'' On Error GoTo Eloop
Dim MasterCode As String
Dim mLabGrpCode As String, mLabTypeCode As String, mLabCode As String
Dim mLabHrsRate As Double, mASC As Integer, mFound As Boolean, mASC1 As Integer
    ImportBtn(Index).BackColor = ProcessColor

    GCn.BeginTrans
    mLabHrsRate = ErrorGCN.Execute("Select LabourRate from Enviro").Fields(0).Value
    
    '' Labour Group Updation
    
    CopyCnt = 0
    ErrorCnt = 0
    
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Labour_Group", GCn, adOpenDynamic, adLockOptimistic
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal("Lab_Group") & ")", "0") & "+1 from Labour_Group").Fields(0).Value
    
    
    Do Until Master.EOF
        
        If IsNull(StringPass(Master.Fields("Area"))) Or StringPass(Master.Fields("Area")) = "" Then GoTo MyNextRecord
        
        If RsNew.RecordCount > 0 Then RsNew.MoveFirst
        mFound = False
        Do Until RsNew.EOF
            If UCase(StringPass(RsNew!LabGrp_Desc)) = UCase(StringPass(left(StringPass(Master.Fields("Area")), 20))) Then
                mFound = True
                Exit Do
            End If
            RsNew.MoveNext
        Loop
        If mFound = True Then GoTo MyNextRecord
        
        
        mFound = False
        If RsNew.RecordCount > 0 Then RsNew.MoveFirst
        mASC = 48
        Do Until RsNew.EOF
            If RsNew!Lab_Group = Chr(mASC) Then
                RsNew.MoveFirst
                mASC = mASC + 1
                If mASC > 57 And mASC < 65 Then
                    mASC = 65
                ElseIf mASC > 90 Then
                    MsgBox "Code Limit for Labour Group Master has been Over, Please Contact to Administrator/Dataman", vbCritical, "Warning !!!"
                    GoTo ClearBuffer
                End If
            Else
                RsNew.MoveNext
            End If
        Loop
        MasterCode = Chr(mASC)
        'Insert New Rec
        RsNew.AddNew
        RsNew!Lab_Group = MasterCode
        RsNew!LabGrp_Desc = left(StringPass(Master.Fields("Area")), 20)
        RsNew!Site_Code = PubSiteCode
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    
    
    '' Labour Type Updation
    CopyCnt = 0
    ErrorCnt = 0
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Labour_Type", GCn, adOpenDynamic, adLockOptimistic
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal("Lab_Type") & ")", "0") & "+1 from Labour_Type").Fields(0).Value
    Do Until Master.EOF
        
        If IsNull(StringPass(Master.Fields("Sub-Area"))) Or StringPass(Master.Fields("Sub-Area")) = "" Then GoTo MyNextRecord1
        
        If RsNew.RecordCount > 0 Then RsNew.MoveFirst
        Do Until RsNew.EOF
            If UCase(StringPass(RsNew!Lab_Desc)) = UCase(StringPass(left(StringPass(Master.Fields("Sub-Area")), 20))) Then
                GoTo MyNextRecord1
            End If
            RsNew.MoveNext
        Loop
        
        If CodeCnt > 99 Then
            mASC = 48
            mASC1 = 32
            If RsNew.RecordCount > 0 Then RsNew.MoveFirst
            Do Until RsNew.EOF
                If RsNew!Lab_Type = Trim(Chr(mASC1) & Chr(mASC)) Then
                    RsNew.MoveFirst
                    mASC = mASC + 1
                    If mASC > 57 And mASC < 65 Then
                        mASC = 65
                    ElseIf mASC > 90 Then
                        If mASC1 = 32 Then mASC1 = 48 Else mASC1 = mASC1 + 1
                        mASC = 48
                        If mASC1 > 57 And mASC1 < 65 Then
                            mASC1 = 65
                        ElseIf mASC1 > 90 Then
                            MsgBox "Code Limit for Labour Type Master has been Over, Please Contact to Administrator/Dataman", vbCritical, "Warning !!!"
                            GoTo ClearBuffer
                        End If
                    End If
                Else
                    RsNew.MoveNext
                End If
            Loop
            MasterCode = Trim(Chr(mASC1) & Chr(mASC))
        Else
            MasterCode = Right("00" & CodeCnt, 2)
        End If
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!Lab_Type = MasterCode
        RsNew!Lab_Desc = left(StringPass(Master.Fields("Sub-Area")), 20)
        RsNew!Site_Code = PubSiteCode
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord1:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    
    '' Labour Master Updation
    CopyCnt = 0
    ErrorCnt = 0
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Labour", GCn, adOpenDynamic, adLockOptimistic
    Do Until Master.EOF
        'If IsNull(StringPass(Master.Fields("Job Code"))) Or StringPass(Master.Fields("Job Code")) = "" Then GoTo MyNextRecord2
        If IsNull(StringPass(Master.Fields("Job Code"))) Or StringPass(Master.Fields("Job Code")) = "" Then GoTo MyNextRecord2
        If IsNull(StringPass(Master.Fields("Job Code Desc"))) Or StringPass(Master.Fields("Job Code Desc")) = "" Then GoTo MyNextRecord2
        
        If StringPass(left(Master.Fields("Area"), 20)) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Job Code"), "Labour Master", "Area field is empty for Job Code")
            GoTo MyNextRecord2
        End If
        If StringPass(left(Master.Fields("Sub-Area"), 20)) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Job Code"), "Labour Master", "Sub-Area field is empty for Job Code")
            GoTo MyNextRecord2
        End If
        
        If RsNew.RecordCount > 0 Then RsNew.MoveFirst
        
        Do Until RsNew.EOF
            If UTrim(XNull(RsNew!Lab_Code)) = UTrim(XNull(Master.Fields("Job Code"))) Then
                GoTo MyNextRecord2
            End If
            RsNew.MoveNext
        Loop
        
        'Insert New Rec
        RsNew.AddNew
        
        RsNew!Lab_Code = Master.Fields("Job Code")
        RsNew!Lab_Desc = left(StringPass(Master.Fields("Job Code Desc")), 40)
        RsNew!Site_Code = PubSiteCode
        RsNew!Div_Code = PubDivCode
        RsNew!External_YN = 0
        RsNew!Major_YN = 0
        RsNew!Chrg_From = "C"
                
        If StringPass(left(Master.Fields("Area"), 20)) <> "" Then
            RsNew!Lab_Group = XNull(GCn.Execute("Select Lab_Group from Labour_Group where " & cUCase("LabGrp_Desc") & "='" & UCase(StringPass(left(Master.Fields("Area"), 20))) & "'").Fields(0).Value)
        End If
        If StringPass(left(Master.Fields("Sub-Area"), 20)) <> "" Then
            RsNew!Lab_Type = XNull(GCn.Execute("Select Lab_Type from Labour_Type where " & cUCase("Lab_Desc") & "='" & UCase(StringPass(left(Master.Fields("Sub-Area"), 20))) & "'").Fields(0).Value)
        End If
        RsNew!ModelBased = 0
        RsNew!WTime_Req = Master.Fields("Standard Labour Hrs")
        RsNew!Time_Req = Master.Fields("Standard Labour Hrs")
        RsNew!Lab_Rate = Master.Fields("Standard Labour Hrs") * mLabHrsRate
        
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord2:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    
ClearBuffer:
    GCn.CommitTrans
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone

lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & StringPass(Master.Fields("Job Code")) & "','Labour Group/Type/Desc master','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub
Private Sub StkAdjUpdate(Index)
'' On Error GoTo Eloop
Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer, mV_Type As String
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String
Dim mPrefix As String
Dim mAdjSlipNo As String
Dim mSrl As Integer, mQty As Double, mAmount As Double
Dim mTax_Amt As Double, mTax_Amt1 As Double, mTaxable As Boolean
Dim mVATApplicable As Boolean, mTrnType As String
    
    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from SP_Stock", GCn, adOpenDynamic, adLockOptimistic
    
    mVATApplicable = ErrorGCN.Execute("Select VatApplicable from Enviro").Fields(0).Value
    
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    Do Until Master.EOF
        mAdjSlipNo = StringPass(Master.Fields("Transaction #"))
        
        If Master!Type <> "Adjustment" Then GoTo DuplicateSkipped
        
        If Trim(mAdjSlipNo) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Transaction #"), "Stock Adjustment", "Transaction # Field is empty")
            GoTo MyNextRecord
        End If
        
        If StringPass(Master.Fields("Destination Location")) = "" And StringPass(Master.Fields("Source Location")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Stock Adjustment", "Source Location & Destination Location are empty")
            GoTo MyNextRecord
        End If
        
        If StringPass(Master.Fields("Destination Location")) = "" Then  '' Issue Type
            mTrnType = "I"
            mV_Type = "SYIAD"
        Else                                                            '' Recd. Type
            mTrnType = "R"
            mV_Type = "SXRAD"
        End If
        
        If GCn.Execute("Select V_no from SP_Purch where SiebelDocID='" & mAdjSlipNo & "' and V_Type='" & mV_Type & "'").RecordCount > 0 Then
            GoTo DuplicateSkipped
        End If
        
        If IsNull(StringPass(Master.Fields("Transaction Date/Time"))) Or StringPass(Master.Fields("Transaction Date/Time")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Stock Adjustment", "Transaction Date/Time is Empty")
            GoTo MyNextRecord
        End If
            
        If IsNull(StringPass(Master.Fields("Part #").Value)) Or StringPass(Master.Fields("Part #").Value) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Stock Adjustment", "Part Number field is Empty")
            GoTo MyNextRecord
        Else
            If GCn.Execute("Select Part_No from Part where Part_No='" & Master.Fields("Part #") & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Part #"), "Stock Adjustment", "Part Number not exist in AUTOMAN Part Master")
                'GoTo MyNextRecord
            End If
        End If
            
        If IsNull(StringPass(Master.Fields("Qty"))) Or StringPass(Master.Fields("Qty")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Stock Adjustment", "Qty field is Empty")
            GoTo MyNextRecord
        End If
        
        If mTrnType = "R" Then           '' Recd
            If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").RecordCount > 0 Then
                mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").Fields(0).Value
                mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").Fields(0).Value
                mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Stock Adjustment", "Automan Site/Division is not Defined in SiteDivision Table for this Destination Division")
                GoTo MyNextRecord
            End If
        Else
            If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Source Division")) & "'").RecordCount > 0 Then
                mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Source Division")) & "'").Fields(0).Value
                mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Source Division")) & "'").Fields(0).Value
                mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Source Division")) & "'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Stock Adjustment", "Automan Site/Division is not Defined in SiteDivision Table for this Source Division")
                GoTo MyNextRecord
            End If
        End If
        
        If mTrnType = "I" Then
            If ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Source Division")) & "' and Type='Adjustment Received'").RecordCount > 0 Then
                mPartyCode = ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Source Division")) & "' and Type='Adjustment Received'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Source Division"), "Stock Adjustment", "This Source Division is not defined in AccountConversion Table (As Account Code) : Transaction No." & mAdjSlipNo)
                GoTo MyNextRecord
            End If
        Else
            If ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Destination Division")) & "' and Type='Adjustment Issued'").RecordCount > 0 Then
                mPartyCode = ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Destination Division")) & "' and Type='Adjustment Issued'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Destination Division"), "Stock Adjustment", "This Destination Division is not defined in AccountConversion Table (As Account Code) : Transaction No." & mAdjSlipNo)
                GoTo MyNextRecord
            End If
        End If
        
        If mPartyCode = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mAdjSlipNo, "Stock Adjustment", "Account not found in Automan")
            GoTo MyNextRecord
        End If
        
        mPrefix = "SBL" & Format(Master.Fields("Transaction Date/Time"), "yy")
        CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from SP_STOCK where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mV_Type & "'").Fields(0).Value
        
        mDocId = mRecordDiv & mRecordSite & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
            
        mSrl = 1
                
        'Insert New Rec
        With RsNew
            .AddNew
            !DocId = mDocId
            !Site_Code = mRecordSite & mRecordSite
            !V_Type = Trim(mV_Type)
            !V_No = CodeCnt
            !V_DATE = MakeDate(left(Master.Fields("Transaction Date/Time"), 10))
            !Party_Code = mPartyCode
            !Srl_No = mSrl
            !L_C = "L"
            !Part_No = VNull(Master.Fields("part #"))
            !godown = GCn.Execute("Select SprWorksGodown from Syctrl").Fields(0).Value
            If mTrnType = "R" Then
                !Qty_Doc = Val(Master!Qty)
                !Qty_Rec = Val(Master!Qty)
            Else
                !Qty_Iss = Val(Master!Qty)
            End If
            If mVATApplicable Then
                !Tax_YN = 1             '' if VAT is applicable in State
            Else
                !Tax_YN = 0             ''IIf(mLocal = "L", 0, 1)
            End If
            !MRP_YN = IIf(mRecordDiv = "C", 1, 0)
            !Amount = Val(Master!Value)
            !Net_Amt = Val(Master!Value)
            !Rate = Round(Master!Value / Master!Qty, 2) ' Master!Rate
            !Part_SrlNo = mSrl
            !Purpose = "O"
            !Remark = mAdjSlipNo
            !SiebelDocID = mAdjSlipNo
            
            !TaxAmt = 0
            !TaxPer = 0
            !Disc_Per = 0
            !Disc_Amt = 0
            !Ord_Discper = 0
            !Ord_DiscAmt = 0

            !U_Name = "Siebel"
            !U_EntDt = Format(PubLoginDate, "Short Date")
            !U_AE = "A"
            .Update
        End With

DuplicateSkipped:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh

MyNextRecord:
        Master.MoveNext
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone

lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    mSrl = 0
    ErrorCnt = ErrorCnt + mSrl
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Stock Adjustment','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub


Private Sub SprSaleUpdate(Index)
'On Error GoTo Eloop
Dim mPartyCode As String
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String, mOrderQty As Double, mPhysicalQty As Double
Dim mname As String, mLubType As String, mTrnType As String, mDebitAc As String, mFormCode As String
Dim mInvoiceID As String, mV_Type As String, mInvoiceNo As String, mPrefix As String
Dim mChallanID As String, mChallanType As String, mChallanNo As String
Dim mGatePassID As String, mGatePassNo As String
Dim mHeaderParty As String, mHeaderAcCode As String, mSrl As Integer
Dim mQty As Double, mCount As Integer, mAmount As Double, mVATApplicable As Boolean
Dim mTax_Amt As Double, mTax_Amt1 As Double, mLocal As String, mTaxable As Boolean, mLength As Integer
Dim mEditFlag As Boolean
Dim mLineAmount As Double
Dim mLineRate As Double
Dim mLineMrpRate As Double
Dim mLineDisPer As Double
Dim mLineDiscount As Double
Dim mLineNetAmount As Double
Dim mLineTaxAmount As Double
  
Dim mSprAmt_MRP_TB As Double, mSprAmt_MRP_TP As Double, mOilAmt_MRP_TB As Double, mOilAmt_MRP_TP As Double
Dim mSprAmt_TB As Double, mSprAmt_TP As Double, mOilAmt_TB As Double, mOilAmt_TP As Double
Dim mD_Per_TB As Double, mD_Per_TP As Double, mD_Amt_TB As Double, mD_Amt_TP As Double
Dim mD_Per_MRP_TB As Double, mD_Per_MRP_TP As Double, mD_Amt_MRP_TB As Double, mD_Amt_MRP_TP As Double
    
    
    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from SP_Sale", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsNew1 = New adodb.Recordset
    RsNew1.CursorLocation = adUseClient
    RsNew1.Open "Select * from SP_Stock", GCn, adOpenDynamic, adLockOptimistic
    
    mVATApplicable = ErrorGCN.Execute("Select VatApplicable from Enviro").Fields(0).Value
    
    '' Provision of Cancelled Invoice not taken
    
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    Do Until Master.EOF
        If Master!Invoice_Status <> "New" Then GoTo MyNextRecord
        
        If IsNull(StringPass(Master.Fields("Invoice_No"))) Or StringPass(Master.Fields("Invoice_No")) = "" Then GoTo MyNextRecord
        
        mInvoiceNo = StringPass(Master.Fields("Invoice_No").Value)
        mChallanNo = StringPass(Master.Fields("Order_No").Value)
                
        If IsNull(StringPass(Master.Fields("Division"))) Or StringPass(Master.Fields("Division")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Division Name field is Empty")
            GoTo MyNextRecord
        Else
            If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").RecordCount > 0 Then
                mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").Fields(0).Value
                mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").Fields(0).Value
                mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master.Fields("Division")), "Spare Sale Bill", "Automan Site/Division is not Defined in SiteDivision Table for this Siebel Division , Invoice No. " & mInvoiceNo)
                GoTo MyNextRecord
            End If
        End If
                
        mChallanType = "SYSC"
        If Master.Fields("Mode Of Payment") = "CASH" Then
            mV_Type = "SYSIC"
        Else
            mV_Type = "SYSIR"
        End If
        
        If IsNull(StringPass(Master.Fields("Account_Code"))) Or StringPass(Master.Fields("Account_Code")) = "" Then
            If IsNull(StringPass(Master.Fields("Customer_Code"))) Or StringPass(Master.Fields("Customer_Code")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Account/Customer Code field is Empty")
                GoTo MyNextRecord
            Else
                mHeaderParty = Master.Fields("Full Name")
                mHeaderAcCode = Master.Fields("Customer_Code")
            End If
        Else
''''            mHeaderParty = Master.Fields("Account_Name")
            mHeaderAcCode = Master.Fields("Account_Code")
        End If
        
        If Master.Fields("Mode Of Payment") = "CASH" Then
             mPartyCode = ErrorGCN.Execute("select CashAccountCode from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").Fields(0).Value
        Else
            If GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & mHeaderAcCode & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Account/Customer Code not found in Ledger Account Master of Automan")
                GoTo MyNextRecord
            Else
                mPartyCode = GCn.Execute("Select SubCode from SubGroup where siebelCode='" & mHeaderAcCode & "'").Fields(0).Value
            End If
        End If
        If GCn.Execute("Select V_no from SP_Sale where SiebelDocID='" & mInvoiceNo & "' and V_Type='" & mV_Type & "'").RecordCount > 0 Then           ' and Party_Code='" & mPartyCode & "'
            'GoTo DuplicateSkipped
            mEditFlag = True
        End If
            
        
        If IsNull(Master.Fields("Invoice_Date").Value) Or StringPass(Master.Fields("Invoice_Date").Value) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Invoice Date field is Empty")
            GoTo MyNextRecord
        End If
        
        Dim mShortYear As String
        If Month(Master.Fields("Invoice_Date")) > 3 Then
            mShortYear = Right(Format(Master.Fields("Invoice_Date"), "yy"), 1) & Right(Val(Format(Master.Fields("Invoice_Date"), "yy")) + 1, 1)
        Else
            mShortYear = Right(Val(Format(Master.Fields("Invoice_Date"), "yy")) - 1, 1) & Right(Format(Master.Fields("Invoice_Date"), "yy"), 1)
        End If
        mPrefix = "SBL" & mShortYear 'Format(Master.Fields("Receipt_Date"), "yy")
        
        '' For Invoice Details :
        CodeCnt = Right(mInvoiceNo, 5) ''GCn.Execute("Select iif(isnull(Max(V_No)),0,Max(V_no))+1 from SP_Sale where Left(DocID,1)='" & mRecordDiv & "' and mid(DocID,2,2)='" & mRecordSite & mRecordSite & "' and V_Type='" & mV_Type & "'").Fields(0).Value
        If mEditFlag Then
            mInvoiceID = GCn.Execute("Select DocId from SP_Sale where SiebelDocID='" & mInvoiceNo & "' and V_Type='" & mV_Type & "'").Fields(0)
        Else
            mInvoiceID = mRecordDiv & mRecordSite & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
        End If
'
'        If CodeCnt = "05611" Then
'            MsgBox ""
'        End If
        '' For Challan Details :
        CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from SP_Sale where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mChallanType & "'").Fields(0).Value
        mChallanNo = CodeCnt
        If mEditFlag Then
            mChallanID = GCn.Execute("Select DocId From Sp_Sale Where Invoice_DocID='" & mInvoiceID & "'").Fields(0).Value
            mChallanNo = Val(Right(mChallanID, 8))
        Else
            mChallanID = mRecordDiv & mRecordSite & mRecordSite & " " & mChallanType & mPrefix & Right("00000000" & CodeCnt, 8)
        End If
        
        '' For GatePass Details :
        'CodeCnt = GCn.Execute("Select iif(isnull(Max(val(Left(GP_No,5)))),0,Max(val(Left(GP_no,5))))+1 from SP_Sale where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mChallanType & "'").Fields(0).Value
        CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal("Right(GP_No,5)") & ")", "0") & "+1 from SP_Sale where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mChallanType & "' and gp_no <>''").Fields(0).Value
        mGatePassID = mRecordDiv & mRecordSite & mRecordSite & Right("00000" & CodeCnt, 5)
        
        mLubType = GCn.Execute("Select PartGrade_Lub from Syctrl").Fields(0).Value
        
'        If mTrnType = mLubType Then
'            mFormCode = ErrorGCN.Execute("Select SpareSaleFormLubs from Enviro").Fields(0).Value
'            mTax_Amt = Val(Master.Fields("Total_Tax_Amount"))
'            mLocal = "L"
'        Else
            If Not IsNull(Master.Fields("Total_Tax_Amount")) Then
                If Val(Mid(Master.Fields("Total_Tax_Amount"), 4, Len(Master.Fields("Total_Tax_Amount")) - 3)) > 0 Then
                    mFormCode = ErrorGCN.Execute("Select SpareSaleFormVAT from Enviro").Fields(0).Value
                    mTax_Amt = Val(Mid(Master.Fields("Total_Tax_Amount"), 4, Len(Master.Fields("Total_Tax_Amount")) - 3))
                    mLocal = "L"
                End If
            End If
            If Not IsNull(Master.Fields("LST")) Then
                If Val(Mid(Master.Fields("LST"), 4, Len(Master.Fields("LST")) - 3)) > 0 Then
                    mFormCode = ErrorGCN.Execute("Select SpareSaleFormLocal from Enviro").Fields(0).Value
                    mTax_Amt = Val(Mid(Master.Fields("Total_Tax_Amount"), 4, Len(Master.Fields("Total_Tax_Amount")) - 3))
                    mLocal = "L"
                End If
'            End If
'            If Not IsNull(Master.Fields("CST")) Then
'                If Val(Master.Fields("CST")) > 0 Then
'                    mFormCode = ErrorGCN.Execute("Select SparePurchFormCST from Enviro").Fields(0).Value
'                    mTax_Amt = Val(Master.Fields("Total_Tax_Amount"))
'                    mLocal = "C"
'                End If
'            End If
        End If
        If mFormCode = "" Then
            mFormCode = ErrorGCN.Execute("Select SpareSaleFormVAT from Enviro").Fields(0).Value
            mTax_Amt = Val(Mid(Master.Fields("Total_Tax_Amount"), 4, Len(Master.Fields("Total_Tax_Amount")) - 3))
            mLocal = "L"
        End If
        
        mDebitAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc where Div_Code='" & mRecordDiv & "' and Form_Code='" & mFormCode & "'").Fields(0).Value
        
        
        Set Master1 = CreateObject("ADODB.Recordset")
        GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] where Invoice_No='" & mInvoiceNo & "' Order By  Invoice_No"
        Master1.Open GSQL, ExcelGcn2, adOpenStatic
        GCn.Execute "Delete From Sp_Stock Where DocId='" & mChallanID & "'"
        
        If Master1.RecordCount = 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Details of Line File Item not found for Invoice No. " & mInvoiceNo)
            GoTo MyNextRecord
        End If
        
        mSrl = 1
        mAmount = 0
        mTax_Amt1 = 0
                        
        mSprAmt_MRP_TB = 0: mSprAmt_MRP_TP = 0: mOilAmt_MRP_TB = 0: mOilAmt_MRP_TP = 0
        mD_Per_MRP_TB = 0: mD_Per_MRP_TP = 0: mD_Amt_MRP_TB = 0: mD_Amt_MRP_TP = 0
        
        mSprAmt_TB = 0: mSprAmt_TP = 0: mOilAmt_TB = 0: mOilAmt_TP = 0
        mD_Per_TB = 0: mD_Per_TP = 0: mD_Amt_TB = 0: mD_Amt_TP = 0
        
        Master1.MoveFirst
        Do Until Master1.EOF
            mLineAmount = 0
            mLineRate = 0
            mLineDiscount = 0
            mLineDisPer = 0
            mLineMrpRate = 0
            mLineNetAmount = 0
            mLineTaxAmount = 0
            
            
            'Insert New Rec
            With RsNew1
                .AddNew
                !DocId = mChallanID
                !Site_Code = mRecordSite & mRecordSite
                !V_Type = Trim(mChallanType)
                !V_No = mChallanNo
                !V_DATE = MakeDate(left(Master.Fields("Order Date"), 10))        '' Order Date = Challan Date
                !Party_Code = mPartyCode
                !Srl_No = mSrl
                !L_C = mLocal
                !Remark = XNull(Master1.Fields("Order_No"))
                !Part_No = VNull(Master1.Fields("Part #"))
                !godown = GCn.Execute("Select SprWorksGodown from Syctrl").Fields(0).Value
'                !Qty_Doc = Val(Master1!Quantity)
'                !Qty_Rec = Val(Master1!Quantity)
                !Qty_Iss = Val(Master1!Quantity)
                If mVATApplicable Then
                    !Tax_YN = 1             '' if VAT is applicable in State
                Else
                    !Tax_YN = IIf(mLocal = "L", 0, 1)
                End If
                !MRP_YN = IIf(mRecordDiv = "C", 1, 0)
                
                
                If mRecordDiv = "C" Then
                    '' Goods Value
                    mLineAmount = VNull(Master1.Fields("Net_Amount")) + IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))     ' (Round(IIf(IsNull(Master1.Fields("NTA")), 0, Master1.Fields("NTA")) * IIf(IsNull(Master1.Fields("Quantity")), 0, Master1.Fields("Quantity")), 2)) + IIf(IsNull(Master1.Fields("Tax Amount")), 0, Val(Master1.Fields("Tax Amount")))
                    If Not IsNull(Master1.Fields("Tax Amount")) Then
                        If Val(Master1.Fields("Tax Amount")) > 0 Then
                            mLineTaxAmount = IIf(IsNull(Master1.Fields("Tax Amount")), 0, Val(Master1.Fields("Tax Amount")))
                        End If
                    End If
                    mLineDiscount = IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
                    mLineNetAmount = VNull(Master1.Fields("Net_Amount")) - VNull(Master1.Fields("Tax Amount"))

                Else
                    '' Goods Value
                    mLineAmount = IIf(IsNull(Master1.Fields("Net_Amount")), 0, Master1.Fields("Net_Amount")) + IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
                    If Not IsNull(Master1.Fields("Tax Amount")) Then
                        If Val(Master1.Fields("Tax Amount")) > 0 Then
                            mLineTaxAmount = IIf(IsNull(Master1.Fields("Tax Amount")), 0, Val(Master1.Fields("Tax Amount")))
                        End If
                    End If
                    mLineDiscount = IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
                    mLineNetAmount = IIf(IsNull(Master1.Fields("Net_Amount")), 0, Master1.Fields("Net_Amount")) '+ IIf(IsNull(Master1.Fields("Tax Amount")), 0, Val(Master1.Fields("Tax Amount")))
                End If
                
                
                !Amount = mLineAmount
                !TaxAmt = mLineTaxAmount
                mTax_Amt1 = mTax_Amt1 + mLineTaxAmount
                !Disc_Amt = mLineDiscount
                If mLineAmount > 0 Then
                    mLineDisPer = Round(mLineDiscount * 100 / mLineAmount, 4)
                Else
                    mLineDisPer = 0
                End If
                !Disc_Per = mLineDisPer
                !Net_Amt = mLineNetAmount
                If mLineNetAmount = 0 Then
                    !TaxPer = 0
                Else
                    !TaxPer = Round(mLineTaxAmount * 100 / mLineNetAmount, 3)
                End If
                
                
                
                
                
                
                If GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("part #") & "'").RecordCount > 0 Then
                    mTrnType = GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("part #") & "'").Fields(0).Value
                Else
                    mTrnType = "S"
                End If
                If mRecordDiv = "C" Then
                    If !Tax_YN = 1 Then
                        If mLubType = mTrnType Then
                            mOilAmt_MRP_TB = mOilAmt_MRP_TB + (IIf(IsNull(Master1.Fields("NTA")), 0, Master1.Fields("NTA")) * IIf(IsNull(Master1.Fields("Quantity")), 0, Master1.Fields("Quantity"))) - IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
                        Else
                            mSprAmt_MRP_TB = mSprAmt_MRP_TB + (IIf(IsNull(Master1.Fields("NTA")), 0, Master1.Fields("NTA")) * IIf(IsNull(Master1.Fields("Quantity")), 0, Master1.Fields("Quantity"))) - IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
                        End If
                    Else
                        If mLubType = mTrnType Then
                            mOilAmt_MRP_TP = mOilAmt_MRP_TP + (IIf(IsNull(Master1.Fields("NTA")), 0, Master1.Fields("NTA")) * IIf(IsNull(Master1.Fields("Quantity")), 0, Master1.Fields("Quantity"))) - IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
                        Else
                            mSprAmt_MRP_TP = mSprAmt_MRP_TP + (IIf(IsNull(Master1.Fields("NTA")), 0, Master1.Fields("NTA")) * IIf(IsNull(Master1.Fields("Quantity")), 0, Master1.Fields("Quantity"))) - IIf(IsNull(Master1.Fields("Discount")), 0, Master1.Fields("Discount"))
                        End If
                    End If
                Else
                    If !Tax_YN = 1 Then
                        If mLubType = mTrnType Then
                            mOilAmt_TB = mOilAmt_TB + IIf(IsNull(Master1.Fields("Net_Amount")), 0, Master1.Fields("Net_Amount"))
                        Else
                            mSprAmt_TB = mSprAmt_TB + IIf(IsNull(Master1.Fields("Net_Amount")), 0, Master1.Fields("Net_Amount"))
                        End If
                    Else
                        If mLubType = mTrnType Then
                            mOilAmt_TP = mOilAmt_TP + IIf(IsNull(Master1.Fields("Net_Amount")), 0, Master1.Fields("Net_Amount"))
                        Else
                            mSprAmt_TP = mSprAmt_TP + IIf(IsNull(Master1.Fields("Net_Amount")), 0, Master1.Fields("Net_Amount"))
                        End If
                    End If
                End If
                
                
                
                
                If IsNull(Master1!Quantity) Or Master1!Quantity = 0 Then
                    mLineRate = 0
                    mLineMrpRate = 0
                Else
                    If mRecordDiv = "C" Then
                        mLineRate = Round((!Amount - !TaxAmt) / !Qty_Iss, 5)
                        mLineMrpRate = Round(!Amount / !Qty_Iss, 5)
                    Else
                        mLineRate = Round(!Amount / !Qty_Iss, 5)
                        mLineMrpRate = mLineRate
                    End If
                    '!V_Rate = !Rate     'Round(Val(Replace(Mid(IIf(IsNull(Master1.Fields("Net Amount")), "Rs.0", Master1.Fields("Net Amount")), 4, 15), ",", "")) / Master1!Qty, 4)
                End If
                
                
                !Rate = mLineRate
                !Mrp_Rate = mLineMrpRate
                
                
                
                
                !Part_SrlNo = mSrl
                !Ord_Discper = 0
                !Ord_DiscAmt = 0
    
                '' Invoice Details Updation
                !Invoice_DocID = mInvoiceID
                !V_Date2 = MakeDate(Master.Fields("Invoice_Date"))
                !Rate2 = mLineRate
                !MRP_Rate2 = mLineMrpRate
                !Amount2 = mLineAmount
                !Disc_Per2 = mLineDisPer
                !Disc_Amt2 = mLineDiscount
                !Net_Amt2 = mLineNetAmount
    
                !U_Name = "Siebel"
                !U_EntDt = Format(PubLoginDate, "Short Date")
                !U_AE = IIf(mEditFlag, "E", "A")
                .Update
            End With
            mAmount = mAmount + IIf(IsNull(Master1.Fields("Net_Amount")), 0, Master1.Fields("Net_Amount"))    'Val(Replace(Mid(IIf(IsNull(Master1.Fields("Net_Amount")), "Rs.0", Master1.Fields("Net_Amount")), 4, 15), ",", ""))
            mSrl = mSrl + 1
LineFileNextrecord:
            Master1.MoveNext
        Loop
        If mRecordDiv = "C" Then
            If Round(mAmount - mTax_Amt1, 1) <> Round(IIf(IsNull(Master.Fields("Total_Parts_Amount")), 0, Master.Fields("Total_Parts_Amount")), 1) Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Line File Goods Value Total is not matched with Header file Goods Value (But Entry Posted in Automan)")
                'GoTo MyNextRecord
            End If
        Else
            If Format(mAmount, "0") <> Format(IIf(IsNull(Master.Fields("Total_Parts_Amount")), 0, Master.Fields("Total_Parts_Amount")), "0") Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Line File Goods Value Total is not matched with Header file Goods Value (But Entry Posted in Automan)")
                'GoTo MyNextRecord
            End If
        End If
        If Round(mTax_Amt1, 1) <> Round(Val(Format(Mid(Master.Fields("Total_Tax_Amount"), 4, Len(Master.Fields("Total_Tax_Amount")) - 3), "0.00")), 1) Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Line File Tax Value Total is not matched with Header file Tax Value (But Entry Posted in Automan)")
            'GoTo MyNextRecord
        End If
'        If Master.Fields("Discount Parts") <> "Rs.0.00" Then
'            MsgBox ""
'        End If
        If mRecordDiv = "C" Then
            If mSprAmt_MRP_TB + mOilAmt_MRP_TB > 0 Then
                mD_Amt_MRP_TB = IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
                mD_Amt_TB = IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
                If mD_Amt_MRP_TB > 0 Then
                    mD_Per_MRP_TB = Round(mD_Amt_MRP_TB * 100 / (mSprAmt_MRP_TB + mOilAmt_MRP_TB), 4)
                    mD_Per_TB = Round(mD_Amt_MRP_TB * 100 / (mSprAmt_MRP_TB + mOilAmt_MRP_TB), 4)
                End If
            Else
                mD_Amt_MRP_TP = IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
                mD_Amt_TP = IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
                If mD_Amt_MRP_TP > 0 Then
                    mD_Per_MRP_TP = Round(mD_Amt_MRP_TP * 100 / (mSprAmt_MRP_TP + mOilAmt_MRP_TP), 4)
                    mD_Per_TP = Round(mD_Amt_MRP_TP * 100 / (mSprAmt_MRP_TP + mOilAmt_MRP_TP), 4)
                End If
            End If
        Else
            If mSprAmt_TB + mOilAmt_TB > 0 Then
                mD_Amt_TB = IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
                If mD_Amt_TB > 0 Then
                    mD_Per_TB = Round(mD_Amt_TB * 100 / (mSprAmt_TB + mOilAmt_TB), 4)
                End If
            Else
                mD_Amt_TP = IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Discount Parts"), 4, Len(Master.Fields("Discount Parts")) - 3), "0.00")))
                If mD_Amt_TP > 0 Then
                    mD_Per_TP = Round(mD_Amt_TP * 100 / (mSprAmt_TP + mOilAmt_TP), 4)
                End If
            End If
        End If
        
        If mEditFlag Then
            'Insert New Rec for Challan
            GCn.Execute "Update Sp_Sale Set SprAmt_Mrp_TB=" & Round(mSprAmt_MRP_TB, 2) & ", SprAmt_MRP_TP = " & Round(mSprAmt_MRP_TP, 2) & ", " & _
                                        "OilAmt_MRP_TB = " & Round(mOilAmt_MRP_TB, 2) & ", OilAmt_MRP_TP = " & Round(mOilAmt_MRP_TP, 2) & ", " & _
                                        "D_Per_MRP_TB = " & Round(mD_Per_MRP_TB, 2) & ", D_Per_MRP_TP = " & Round(mD_Per_MRP_TP, 2) & ", " & _
                                        "D_Amt_MRP_TB = " & Round(mD_Amt_MRP_TB, 2) & ", D_Amt_MRP_TP = " & Round(mD_Amt_MRP_TP, 2) & ", " & _
                                        "SprAmt_TB = " & Round(mSprAmt_TB, 2) & ", SprAmt_TP = " & Round(mSprAmt_TP, 2) & ", " & _
                                        "OilAmt_TB = " & Round(mOilAmt_TB, 2) & ", OilAmt_TP = " & Round(mOilAmt_TP, 2) & ", " & _
                                        "D_Per_TB = " & Round(mD_Per_TB, 2) & ", D_Per_TP = " & Round(mD_Per_TP, 2) & ", " & _
                                        "D_Amt_TB = " & Round(mD_Amt_TB, 2) & ", D_Amt_TP = " & Round(mD_Amt_TP, 2) & ", " & _
                                        "Addition = 0, Tax_Amt = " & Round(IIf(IsNull(Master.Fields("Total_Tax_Amount")), 0, IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Total_Tax_Amount"), 4, Len(Master.Fields("Total_Tax_Amount")) - 3), "0.00")))), 2) & ", " & _
                                        "Packing = " & Round(IIf(IsNull(Master.Fields("Other Charges")), 0, IIf(IsNull(Master.Fields("Other Charges")), 0, Val(Mid(Master.Fields("Other Charges"), 4, Len(Master.Fields("Other Charges")) - 3)))), 2) & ", " & _
                                        "TOT_Per = 0, TOT_Amt = 0, ReSalTax_Per = 0, ReSalTax_Amt = 0,total_amt = " & Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) & ", " & _
                                        "Rounded = " & Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) - Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 2) & ", " & _
                                        "Det_Tax = 1, AcPosting_YN = 1, U_Name = 'Siebel', U_EntDt = " & ConvertDate(Format(PubLoginDate, "Short Date")) & ", U_AE = 'E' " & _
                                        "Where DocId='" & mChallanID & "'"
            
            GCn.Execute "Update Sp_Sale Set SprAmt_MRP_TB = " & mSprAmt_MRP_TB & ", SprAmt_MRP_TP = " & mSprAmt_MRP_TP & ", OilAmt_MRP_TB = " & mOilAmt_MRP_TB & ", " & _
                                        "OilAmt_MRP_TP = " & mOilAmt_MRP_TP & ", D_Per_MRP_TB = " & mD_Per_MRP_TB & ", D_Per_MRP_TP = " & mD_Per_MRP_TP & ", " & _
                                        "D_Amt_MRP_TB = " & mD_Amt_MRP_TB & ", D_Amt_MRP_TP = " & mD_Amt_MRP_TP & ", SprAmt_TB = " & mSprAmt_TB & " , " & _
                                        "SprAmt_TP = " & mSprAmt_TP & ", OilAmt_TB = " & mOilAmt_TB & ", OilAmt_TP = " & mOilAmt_TP & ", " & _
                                        "D_Per_TB = " & mD_Per_TB & ", D_Per_TP = " & mD_Per_TP & ", D_Amt_TB = " & mD_Amt_TB & ", D_Amt_TP = " & mD_Amt_TP & ", " & _
                                        "Addition = 0, Tax_Amt = " & Round(IIf(IsNull(Master.Fields("Total_Tax_Amount")), 0, IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Total_Tax_Amount"), 4, Len(Master.Fields("Total_Tax_Amount")) - 3), "0.00")))), 2) & ", " & _
                                        "Packing = " & Round(IIf(IsNull(Master.Fields("Other Charges")), 0, IIf(IsNull(Master.Fields("Other Charges")), 0, Val(Mid(Master.Fields("Other Charges"), 4, Len(Master.Fields("Other Charges")) - 3)))), 2) & ", " & _
                                        "TOT_Per = 0, TOT_Amt = 0, ReSalTax_Per = 0, ReSalTax_Amt = 0, " & _
                                        "total_amt = " & Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) & ", " & _
                                        "Rounded = " & Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) - Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 2) & ", " & _
                                        "Det_Tax = 1, AcPosting_YN = 1, U_Name = 'Siebel', U_EntDt = " & ConvertDate(Format(PubLoginDate, "Short Date")) & ", U_AE = 'E' " & _
                                        "Where DocId='" & mInvoiceID & "'"
            mEditFlag = False
        Else
            'Insert New Rec for Challan
            With RsNew
                .AddNew
                !DocId = mChallanID
                !DocIDHelp = Replace(mChallanID, " ", "")
                !Site_Code = mRecordSite & mRecordSite
                !V_Type = Trim(mChallanType)
                !V_No = mChallanNo
                !V_DATE = MakeDate(left(Master.Fields("Order Date"), 10))
                !Party_Code = mPartyCode
                !Cash_Credit = Master.Fields("Mode Of Payment")
                !Party_Name = left(mHeaderParty, 40)
                !L_C = mLocal
                !Form_Code = mFormCode
                !CrAc = mDebitAc
                !SiebelDocID = Master.Fields("Order_No")
                !Invoice_DocID = mInvoiceID
                !PType = "General"
                !GP_No = mGatePassID
                !GP_Date = MakeDate(left(Master.Fields("Order Date"), 10))
                
                !SprAmt_MRP_TB = Round(mSprAmt_MRP_TB, 2)
                !SprAmt_MRP_TP = Round(mSprAmt_MRP_TP, 2)
                !OilAmt_MRP_TB = Round(mOilAmt_MRP_TB, 2)
                !OilAmt_MRP_TP = Round(mOilAmt_MRP_TP, 2)
                !D_Per_MRP_TB = Round(mD_Per_MRP_TB, 2)
                !D_Per_MRP_TP = Round(mD_Per_MRP_TP, 2)
                !D_Amt_MRP_TB = Round(mD_Amt_MRP_TB, 2)
                !D_Amt_MRP_TP = Round(mD_Amt_MRP_TP, 2)
                
                !SprAmt_TB = Round(mSprAmt_TB, 2)
                !SprAmt_TP = Round(mSprAmt_TP, 2)
                !OilAmt_TB = Round(mOilAmt_TB, 2)
                !OilAmt_TP = Round(mOilAmt_TP, 2)
                !D_Per_TB = Round(mD_Per_TB, 2)
                !D_Per_TP = Round(mD_Per_TP, 2)
                !D_Amt_TB = Round(mD_Amt_TB, 2)
                !D_Amt_TP = Round(mD_Amt_TP, 2)
                
                !Addition = 0
                
                !Tax_Amt = Round(IIf(IsNull(Master.Fields("Total_Tax_Amount")), 0, IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Total_Tax_Amount"), 4, Len(Master.Fields("Total_Tax_Amount")) - 3), "0.00")))), 2)
                !Packing = Round(IIf(IsNull(Master.Fields("Other Charges")), 0, IIf(IsNull(Master.Fields("Other Charges")), 0, Val(Mid(Master.Fields("Other Charges"), 4, Len(Master.Fields("Other Charges")) - 3)))), 2)
                
                !TOT_Per = 0
                !TOT_Amt = 0
                
                !ReSalTax_Per = 0
                !ReSalTax_Amt = 0
                
                !total_amt = Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0)
                !Rounded = Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) - Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 2)
                
                !Det_Tax = 1
                !AcPosting_YN = 1
                
                !U_Name = "Siebel"
                !U_EntDt = Format(PubLoginDate, "Short Date")
                !U_AE = "A"
                .Update
            End With
            
            'Insert New Rec for Invoice
            With RsNew
                .AddNew
                !DocId = mInvoiceID
                !DocIDHelp = Replace(mInvoiceID, " ", "")
                !Site_Code = mRecordSite & mRecordSite
                !V_Type = Trim(mV_Type)
                !V_No = Right(mInvoiceNo, 5)
                !V_DATE = MakeDate(Master.Fields("Invoice_Date"))
                !Party_Code = mPartyCode
                !Cash_Credit = Master.Fields("Mode Of Payment")
                !Party_Name = left(mHeaderParty, 40)
                !L_C = mLocal
                !Form_Code = mFormCode
                !CrAc = mDebitAc
                !SiebelDocID = mInvoiceNo
                !Invoice_DocID = ""
                !PType = "General"
                !GP_No = mGatePassID
                !GP_Date = MakeDate(left(Master.Fields("Order Date"), 10))
                
                !SprAmt_MRP_TB = mSprAmt_MRP_TB
                !SprAmt_MRP_TP = mSprAmt_MRP_TP
                !OilAmt_MRP_TB = mOilAmt_MRP_TB
                !OilAmt_MRP_TP = mOilAmt_MRP_TP
                !D_Per_MRP_TB = mD_Per_MRP_TB
                !D_Per_MRP_TP = mD_Per_MRP_TP
                !D_Amt_MRP_TB = mD_Amt_MRP_TB
                !D_Amt_MRP_TP = mD_Amt_MRP_TP
                
                !SprAmt_TB = mSprAmt_TB
                !SprAmt_TP = mSprAmt_TP
                !OilAmt_TB = mOilAmt_TB
                !OilAmt_TP = mOilAmt_TP
                !D_Per_TB = mD_Per_TB
                !D_Per_TP = mD_Per_TP
                !D_Amt_TB = mD_Amt_TB
                !D_Amt_TP = mD_Amt_TP
                
                !Addition = 0
                
                !Tax_Amt = Round(IIf(IsNull(Master.Fields("Total_Tax_Amount")), 0, IIf(IsNull(Master.Fields("Discount Parts")), 0, Val(Format(Mid(Master.Fields("Total_Tax_Amount"), 4, Len(Master.Fields("Total_Tax_Amount")) - 3), "0.00")))), 2)
                !Packing = Round(IIf(IsNull(Master.Fields("Other Charges")), 0, IIf(IsNull(Master.Fields("Other Charges")), 0, Val(Mid(Master.Fields("Other Charges"), 4, Len(Master.Fields("Other Charges")) - 3)))), 2)
                
                !TOT_Per = 0
                !TOT_Amt = 0
                
                !ReSalTax_Per = 0
                !ReSalTax_Amt = 0
                
                !total_amt = Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0)
                !Rounded = Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) - Round(Val(Format(Mid(Master.Fields("Parts_Invoice_Amount"), 4, Len(Master.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 2)
                
                !Det_Tax = 1
                !AcPosting_YN = 1
                
                !U_Name = "Siebel"
                !U_EntDt = Format(PubLoginDate, "Short Date")
                !U_AE = "A"
                .Update
            End With
        End If
DuplicateSkipped:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
MyNextRecord:
        Master.MoveNext
NextLoop:
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Spare Sale Bill','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub

''Private Sub SprSaleUpdate(Index)
'''' On Error GoTo Eloop
''Dim mPartyCode As String
''Dim mRecordFirm As String, mOrderQty As Double, mPhysicalQty As Double
''Dim mName As String, mLubType As String, mTrnType As String, mDebitAc As String, mFormCode As String
''Dim mInvoiceID As String, mV_Type As String, mInvoiceNo As String, mPrefix As String
''Dim mChallanID As String, mChallanType As String, mChallanNo As String
''Dim mGatePassID As String, mGatePassNo As String
''Dim mHeaderParty As String, mHeaderAcCode As String, mSrl As Integer
''Dim mQty As Double, mCount As Integer, mAmount As Double, mVATApplicable As Boolean
''Dim mTax_Amt As Double, mTax_Amt1 As Double, mLocal As String, mTaxable As Boolean, mLength As Integer
''
''Dim mSprAmt_MRP_TB As Double, mSprAmt_MRP_TP As Double, mOilAmt_MRP_TB As Double, mOilAmt_MRP_TP As Double
''Dim mSprAmt_TB As Double, mSprAmt_TP As Double, mOilAmt_TB As Double, mOilAmt_TP As Double
''Dim mD_Per_TB As Double, mD_Per_TP As Double, mD_Amt_TB As Double, mD_Amt_TP As Double
''Dim mD_Per_MRP_TB As Double, mD_Per_MRP_TP As Double, mD_Amt_MRP_TB As Double, mD_Amt_MRP_TP As Double
''Dim mOverWriteYn
''Dim RsMast As adodb.Recordset
''Dim RsMast1 As adodb.Recordset
''Dim rsTemp As adodb.Recordset
''Dim i As Integer
''Dim mSprWorksGodown As String
''Dim mTax_Yn As Byte
''Dim mMrp_Yn As Byte
''
''    ImportBtn(Index).BackColor = ProcessColor
''
''    mOverWriteYn = MsgBox("Do you want to OverWrite if Record Exist", vbYesNo)
''
''    GCn.BeginTrans
''    CopyCnt = 0
''    ErrorCnt = 0
''
''    mVATApplicable = ErrorGCN.Execute("Select VatApplicable from Enviro").Fields(0).Value
''    mSprWorksGodown = GCn.Execute("Select SprWorksGodown from Syctrl").Fields(0)
''    '' Provision of Cancelled Invoice not taken
''
''
''    Set RsMast = ErrorGCN.Execute("Select O1.*, SD.AutomanSite, SD.AutomanDiv, SD.AutomanFirm, SD.CashAccountCode " & _
''                                  "From ((OTC1 O1 " & _
''                                  "Left Join OTC2 O2 On O1.Order_No=O2.Order_No) " & _
''                                  "Left Join SiteDivision SD On O1.Division=SD.SiebelDiv) " & _
''                                  "Where O1.Invoice_No Is Not Null And O1.Invoice_No <> '' " & _
''                                  "And  O1.Division Is Not Null Or O1.Division <> '' " & _
''                                  "And ((O1.Account_Code Is Not Null And O1.Account_Code <> '') Or (O1.Customer_Code Is Not Null And O1.Customer_Code <> '')) " & _
''                                  "And O1.Invoice_Date Is Not Null")
''
''    If RsMast.RecordCount > 0 Then
''        Do Until RsMast.EOF
''
''
''            If XNull(RsMast.Fields("Account_Code")) = "" Then
''                If XNull(RsMast.Fields("Customer_Code")) <> "" Then
''                    mHeaderParty = RsMast.Fields("Full Name")
''                    mHeaderAcCode = RsMast.Fields("Customer_Code")
''                End If
''            Else
''                mHeaderAcCode = RsMast.Fields("Account_Code")
''            End If
''
''
''
''            mChallanType = "SYSC"
''            If RsMast.Fields("Mode Of Payment") = "CASH" Then
''                mV_Type = "SYSIC"
''                mPartyCode = XNull(RsMast!CashAccountCode)
''            Else
''                mV_Type = "SYSIR"
''                Set rsTemp = GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & mHeaderAcCode & "'")
''                If rsTemp.RecordCount > 0 Then
''                    mPartyCode = XNull(rsTemp(0))
''                Else
''                    Call InsSkipRecMessage(Index, RsMast.AbsolutePosition, mInvoiceNo, "Account Code", "Account Code does not exist In Automan")
''                    GoTo NextRecord
''                End If
''            End If
''
''
''
''            If GCn.Execute("Select V_no from SP_Sale where SiebelDocID='" & mInvoiceNo & "' and V_Type='" & mV_Type & "'").RecordCount > 0 Then           ' and Party_Code='" & mPartyCode & "'
''                If mOverWriteYn = vbYes Then
''                    GCn.Execute "Delete From Sp_Stock Where DocId In (Select DocId From Sp_Sale Where SiebelDocId='" & mInvoiceNo & "')"
''                    GCn.Execute "Delete From " & FaTable("Ledger") & " Where DocId In (Select DocId From Sp_Sale Where SiebelDocId='" & mInvoiceNo & "')"
''                    GCn.Execute "Delete From " & FaTable("LedgerM") & " Where DocId In (Select DocId From Sp_Sale Where SiebelDocId='" & mInvoiceNo & "')"
''                    GCn.Execute "Delete From Sp_Sale Where SiebelDocId='" & mInvoiceNo & "'"
''                Else
''                    GoTo NextRecord
''                End If
''            End If
''
''
''
''
''
''            Dim mShortYear As String
''
''            If Month(RsMast!Invoice_Date) > 3 Then
''                mShortYear = Right(Format(RsMast!Invoice_Date, "yy"), 1) & Right(Val(Format(RsMast!Invoice_Date, "yy")) + 1, 1)
''            Else
''                mShortYear = Right(Val(Format(RsMast!Invoice_Date, "yy")) - 1, 1) & Right(Format(RsMast!Invoice_Date, "yy"), 1)
''            End If
''            mPrefix = "SBL" & mShortYear
''
''
''
''
''
''            CodeCnt = Right(mInvoiceNo, 5)
''            mInvoiceID = RsMast!AutomanDiv & RsMast!AutomanSite & RsMast!AutomanSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
''
''            CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from SP_Sale where Left(DocID,1)='" & RsMast!AutomanDiv & "' and " & cMID("DocID", "2", "2") & "='" & RsMast!AutomanSite & RsMast!AutomanSite & "' and V_Type='" & mChallanType & "'").Fields(0).Value
''            mChallanNo = CodeCnt
''            mChallanID = RsMast!AutomanDiv & RsMast!AutomanSite & RsMast!AutomanSite & " " & mChallanType & mPrefix & Right("00000000" & CodeCnt, 8)
''
''
''            CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal("Right(GP_No,5)") & ")", "0") & "+1 from SP_Sale where Left(DocID,1)='" & RsMast!AutomanDiv & "' and " & cMID("DocID", "2", "2") & "='" & RsMast!AutomanSite & RsMast!AutomanSite & "' and V_Type='" & mChallanType & "' and gp_no <>''").Fields(0).Value
''            mGatePassID = RsMast!AutomanDiv & RsMast!AutomanSite & RsMast!AutomanSite & Right("00000" & CodeCnt, 5)
''
''            mLubType = GCn.Execute("Select PartGrade_Lub from Syctrl").Fields(0).Value
''
''
''
''
''            If Not IsNull(RsMast.Fields("Total_Tax_Amount")) Then
''                If Val(Mid(RsMast.Fields("Total_Tax_Amount"), 4, Len(RsMast.Fields("Total_Tax_Amount")) - 3)) > 0 Then
''                    mFormCode = ErrorGCN.Execute("Select SpareSaleFormVAT from Enviro").Fields(0).Value
''                    mTax_Amt = Val(Mid(RsMast.Fields("Total_Tax_Amount"), 4, Len(RsMast.Fields("Total_Tax_Amount")) - 3))
''                    mLocal = "L"
''                End If
''            End If
''
''            If Not IsNull(RsMast.Fields("LST")) Then
''                If Val(Mid(RsMast.Fields("LST"), 4, Len(RsMast.Fields("LST")) - 3)) > 0 Then
''                    mFormCode = ErrorGCN.Execute("Select SpareSaleFormLocal from Enviro").Fields(0).Value
''                    mTax_Amt = Val(Mid(RsMast.Fields("Total_Tax_Amount"), 4, Len(RsMast.Fields("Total_Tax_Amount")) - 3))
''                    mLocal = "L"
''                End If
''            End If
''
''            If mFormCode = "" Then
''                mFormCode = ErrorGCN.Execute("Select SpareSaleFormVAT from Enviro").Fields(0).Value
''                mTax_Amt = Val(Mid(RsMast.Fields("Total_Tax_Amount"), 4, Len(RsMast.Fields("Total_Tax_Amount")) - 3))
''                mLocal = "L"
''            End If
''
''            mDebitAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc where Div_Code='" & RsMast!AutomanDiv & "' and Form_Code='" & mFormCode & "'").Fields(0).Value
''
''
''
''
''            mSrl = 1
''            mAmount = 0
''            mTax_Amt1 = 0
''
''            mSprAmt_MRP_TB = 0: mSprAmt_MRP_TP = 0: mOilAmt_MRP_TB = 0: mOilAmt_MRP_TP = 0
''            mD_Per_MRP_TB = 0: mD_Per_MRP_TP = 0: mD_Amt_MRP_TB = 0: mD_Amt_MRP_TP = 0
''
''            mSprAmt_TB = 0: mSprAmt_TP = 0: mOilAmt_TB = 0: mOilAmt_TP = 0
''            mD_Per_TB = 0: mD_Per_TP = 0: mD_Amt_TB = 0: mD_Amt_TP = 0
''
''
''            Set RsMast1 = ErrorGCN.Execute("Select *, " & IIf(RsMast!AutomanDiv = "C", "((Net_Amount+Discount)-TaxAmt/Quantity)", "(Net_Amount+Discount)/Quantity") & " As Rate, (Net_Amount+Discount)/Quantity As Mrp_Rate, Otc2.Net_Amount+Otc2.Discount as Amount, DisAmt*100/Amount As Dis_Per From OTC2 Where Invoice_No='" & RsMast!Invoice_No & "'")
''            With RsMast1
''                If RsMast1.RecordCount > 0 Then
''                    i = 1
''                    Do Until RsMast1.EOF
''
''
''                        If GCn.Execute("Select Part_Grade from Part where part_no='" & RsMast1.Fields("part #") & "'").RecordCount > 0 Then
''                            mTrnType = GCn.Execute("Select Part_Grade from Part where part_no='" & RsMast1.Fields("part #") & "'").Fields(0).Value
''                        Else
''                            mTrnType = "S"
''                        End If
''
''
''                        mTax_Yn = IIf(mVATApplicable, 1, IIf(mLocal = "L", 0, 1))
''                        mTax_Yn = IIf(RsMast!AutomanDiv = "C", 1, 0)
''
''                        If RsMast!AutomanDiv = "C" Then
''                            If mTax_Yn = 1 Then
''                                If mLubType = mTrnType Then
''                                    mOilAmt_MRP_TB = mOilAmt_MRP_TB + (IIf(IsNull(RsMast1.Fields("NTA")), 0, RsMast1.Fields("NTA")) * IIf(IsNull(RsMast1.Fields("Quantity")), 0, RsMast1.Fields("Quantity"))) - IIf(IsNull(RsMast1.Fields("Discount")), 0, RsMast1.Fields("Discount"))
''                                Else
''                                    mSprAmt_MRP_TB = mSprAmt_MRP_TB + (IIf(IsNull(RsMast1.Fields("NTA")), 0, RsMast1.Fields("NTA")) * IIf(IsNull(RsMast1.Fields("Quantity")), 0, RsMast1.Fields("Quantity"))) - IIf(IsNull(RsMast1.Fields("Discount")), 0, RsMast1.Fields("Discount"))
''                                End If
''                            Else
''                                If mLubType = mTrnType Then
''                                    mOilAmt_MRP_TP = mOilAmt_MRP_TP + (IIf(IsNull(RsMast1.Fields("NTA")), 0, RsMast1.Fields("NTA")) * IIf(IsNull(RsMast1.Fields("Quantity")), 0, RsMast1.Fields("Quantity"))) - IIf(IsNull(RsMast1.Fields("Discount")), 0, RsMast1.Fields("Discount"))
''                                Else
''                                    mSprAmt_MRP_TP = mSprAmt_MRP_TP + (IIf(IsNull(RsMast1.Fields("NTA")), 0, RsMast1.Fields("NTA")) * IIf(IsNull(RsMast1.Fields("Quantity")), 0, RsMast1.Fields("Quantity"))) - IIf(IsNull(RsMast1.Fields("Discount")), 0, RsMast1.Fields("Discount"))
''                                End If
''                            End If
''                        Else
''                            If mTax_Yn = 1 Then
''                                If mLubType = mTrnType Then
''                                    mOilAmt_TB = mOilAmt_TB + IIf(IsNull(RsMast1.Fields("Net_Amount")), 0, RsMast1.Fields("Net_Amount"))
''                                Else
''                                    mSprAmt_TB = mSprAmt_TB + IIf(IsNull(RsMast1.Fields("Net_Amount")), 0, RsMast1.Fields("Net_Amount"))
''                                End If
''                            Else
''                                If mLubType = mTrnType Then
''                                    mOilAmt_TP = mOilAmt_TP + IIf(IsNull(RsMast1.Fields("Net_Amount")), 0, RsMast1.Fields("Net_Amount"))
''                                Else
''                                    mSprAmt_TP = mSprAmt_TP + IIf(IsNull(RsMast1.Fields("Net_Amount")), 0, RsMast1.Fields("Net_Amount"))
''                                End If
''                            End If
''                        End If
''
''
''                        GCn.Execute "Insert Into Sp_Stock (DocId, Site_Code, V_Type, V_No, V_Date, " & _
''                                    "Party_Code, Srl_No, L_C, Remark, Part_No, " & _
''                                    "Godown, Qty_Iss, Tax_Yn, Mrp_Yn, Amount, " & _
''                                    "Tax_Amt, Disc_Amt, Disc_Per, Net_Amt, TaxPer, " & _
''                                    "Rate, Mrp_Rate, Part_SrlNo, Ord_DiscPer, Ord_DiscAmt, " & _
''                                    "Invoice_DocId, V_Date2, Rate2, Mrp_Rate2, Amount2, " & _
''                                    "Disc_Per2, Disc_Amt2, Net_Amt2, U_Name, U_EntDt, U_AE) " & _
''                                    "Values('" & mChallanID & "', " & RsMast!AutomanSite & ", '" & mChallanType & "', " & mChallanNo & ", " & ConvertDate(MakeDate(RsMast!Invoice_Date)) & ", " & _
''                                    "'" & mPartyCode & "', " & i & ", '" & mLocal & "', '" & RsMast!Order_No & "', '" & RsMast1.Fields("Part #") & "', " & _
''                                    "'" & mSprWorksGodown & "', " & Val(RsMast1!Quantity) & ", " & IIf(mVATApplicable, 1, IIf(mLocal = "L", 0, 1)) & ", " & IIf(RsMast!AutomanDiv = "C", 1, 0) & ", " & Val(RsMast1!Amount) & ", " & _
''                                    "" & Val(RsMast1.Fields("Tax Amount")) & ", " & Val(RsMast1!Discount) & ", " & Val(RsMast1!DisPer) & ", " & Val(RsMast1!Net_Amount) & ", " & Val(RsMast1!TaxPer) & ", " & _
''                                    "" & Val(RsMast1!Rate) & ", " & Val(RsMast1!Mrp_Rate) & ", " & i & ", 0, 0, " & _
''                                    "'" & mInvoiceID & "', " & ConvertDate(MakeDate(RsMast!Invoice_Date)) & ", " & Val(RsMast1!Rate) & ", " & Val(RsMast1!Mrp_Rate) & ", " & Val(RsMast1!Net_Amount) + Val(RsMast1!Discount) & ", " & _
''                                    "" & Val(RsMast1!Discount) & ", " & Val(RsMast1!DisPer) & ", " & Val(RsMast1!Net_Amount) & ", 'Siebel', " & ConvertDate(PubLoginDate) & ", 'A')"
''
''
''
''                        i = i + 1
''                        RsMast1.MoveNext
''                    Loop
''                Else
''                    Call InsSkipRecMessage(Index, RsMast.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Details of Line File Item not found for Invoice No. " & mInvoiceNo)
''                    GoTo NextRecord
''                End If
''            End With
''
''
''
''
''
''                If RsMast!AutomanDiv = "C" Then
''                    If Round(mAmount - mTax_Amt1, 1) <> Round(IIf(IsNull(RsMast.Fields("Total_Parts_Amount")), 0, RsMast.Fields("Total_Parts_Amount")), 1) Then
''                        Call InsSkipRecMessage(Index, RsMast.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Line File Goods Value Total is not matched with Header file Goods Value (But Entry Posted in Automan)")
''                    End If
''                Else
''                    If Format(mAmount, "0") <> Format(IIf(IsNull(RsMast.Fields("Total_Parts_Amount")), 0, RsMast.Fields("Total_Parts_Amount")), "0") Then
''                        Call InsSkipRecMessage(Index, RsMast.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Line File Goods Value Total is not matched with Header file Goods Value (But Entry Posted in Automan)")
''                        'GoTo MyNextRecord
''                    End If
''                End If
''                If Round(mTax_Amt1, 1) <> Round(Val(Format(Mid(RsMast.Fields("Total_Tax_Amount"), 4, Len(RsMast.Fields("Total_Tax_Amount")) - 3), "0.00")), 1) Then
''                    Call InsSkipRecMessage(Index, RsMast.AbsolutePosition, mInvoiceNo, "Spare Sale Bill", "Line File Tax Value Total is not matched with Header file Tax Value (But Entry Posted in Automan)")
''                    'GoTo MyNextRecord
''                End If
''
''                If RsMast!AutomanDiv = "C" Then
''                    If mSprAmt_MRP_TB + mOilAmt_MRP_TB > 0 Then
''                        mD_Amt_MRP_TB = IIf(IsNull(RsMast.Fields("Discount Parts")), 0, Val(Format(Mid(RsMast.Fields("Discount Parts"), 4, Len(RsMast.Fields("Discount Parts")) - 3), "0.00")))
''                        mD_Amt_TB = IIf(IsNull(RsMast.Fields("Discount Parts")), 0, Val(Format(Mid(RsMast.Fields("Discount Parts"), 4, Len(RsMast.Fields("Discount Parts")) - 3), "0.00")))
''                        If mD_Amt_MRP_TB > 0 Then
''                            mD_Per_MRP_TB = Round(mD_Amt_MRP_TB * 100 / (mSprAmt_MRP_TB + mOilAmt_MRP_TB), 4)
''                            mD_Per_TB = Round(mD_Amt_MRP_TB * 100 / (mSprAmt_MRP_TB + mOilAmt_MRP_TB), 4)
''                        End If
''                    Else
''                        mD_Amt_MRP_TP = IIf(IsNull(RsMast.Fields("Discount Parts")), 0, Val(Format(Mid(RsMast.Fields("Discount Parts"), 4, Len(RsMast.Fields("Discount Parts")) - 3), "0.00")))
''                        mD_Amt_TP = IIf(IsNull(RsMast.Fields("Discount Parts")), 0, Val(Format(Mid(RsMast.Fields("Discount Parts"), 4, Len(RsMast.Fields("Discount Parts")) - 3), "0.00")))
''                        If mD_Amt_MRP_TP > 0 Then
''                            mD_Per_MRP_TP = Round(mD_Amt_MRP_TP * 100 / (mSprAmt_MRP_TP + mOilAmt_MRP_TP), 4)
''                            mD_Per_TP = Round(mD_Amt_MRP_TP * 100 / (mSprAmt_MRP_TP + mOilAmt_MRP_TP), 4)
''                        End If
''                    End If
''                Else
''                    If mSprAmt_TB + mOilAmt_TB > 0 Then
''                        mD_Amt_TB = IIf(IsNull(RsMast.Fields("Discount Parts")), 0, Val(Format(Mid(RsMast.Fields("Discount Parts"), 4, Len(RsMast.Fields("Discount Parts")) - 3), "0.00")))
''                        If mD_Amt_TB > 0 Then
''                            mD_Per_TB = Round(mD_Amt_TB * 100 / (mSprAmt_TB + mOilAmt_TB), 4)
''                        End If
''                    Else
''                        mD_Amt_TP = IIf(IsNull(RsMast.Fields("Discount Parts")), 0, Val(Format(Mid(RsMast.Fields("Discount Parts"), 4, Len(RsMast.Fields("Discount Parts")) - 3), "0.00")))
''                        If mD_Amt_TP > 0 Then
''                            mD_Per_TP = Round(mD_Amt_TP * 100 / (mSprAmt_TP + mOilAmt_TP), 4)
''                        End If
''                    End If
''                End If
''
''
''                GCn.Execute "Insert Into Sp_Sale (DocId, DocIdHelp, Site_Code, V_Type, V_No, " & _
''                            "V_Date, Party_Code, Cash_Credit, Party_Name, L_C, " & _
''                            "Form_Code, CrAc, SiebelDocId, Invoice_DocId, PType, " & _
''                            "Gp_No, Gp_Date, SprAmt_Mrp_TB, SprAmt_Mrp_TP, OilAmt_MrpTB, " & _
''                            "OilAmt_MrpTP, D_Per_Mrp_TB, D_Per_Mrp_Tp, D_Amt_Mrp_TB, D_Amt_Mrp_TP, " & _
''                            "SprAmt_TB, SprAmt_TP, OilAmt_TB, OilAmt_TP, D_Per_TB, " & _
''                            "D_Per_TP, D_Amt_TB, D_Amt_TB, Addition, Tax_Amt, " & _
''                            "Packing, Tot_Tax, Tot_Amt,  ReSalTax_Per, ReSalTax_Amt, " & _
''                            "Total_Amt, Rounded, Det_Tax, AcPostingYn, U_Name, " & _
''                            "U_EndDt, U_AE) " & _
''                            "Values ('" & mChallanID & "', '" & Replace(mChallanID, " ", "") & "', " & RsMast!AutomanSite & ", '" & Trim(mChallanType) & "', " & RsMast!Order_No & ", " & _
''                            "" & ConvertDate(MakeDate(left(RsMast.Fields("Order Date"), 10))) & ", '" & mPartyCode & "', '" & RsMast.Fields("Mode Of Payment") & "', '" & left(mHeaderParty, 40) & "', '" & mLocal & "', " & _
''                            "'" & mFormCode & "', '" & mDebitAc & "', '" & RsMast!Order_No & "', '" & mInvoiceID & "', 'General', " & _
''                            "'" & mGatePassID & "', " & ConvertDate(MakeDate(left(RsMast.Fields("Order Date"), 10))) & ", " & Round(mSprAmt_MRP_TB, 2) & ", " & Round(mSprAmt_MRP_TP, 2) & ", " & Round(mOilAmt_MRP_TB, 2) & ", " & _
''                            "" & Round(mOilAmt_MRP_TP, 2) & ", " & Round(mD_Per_MRP_TB, 2) & ", " & Round(mD_Per_MRP_TP, 2) & ", " & Round(mD_Amt_MRP_TB, 2) & ", " & Round(mD_Amt_MRP_TP, 2) & ", " & _
''                            "" & Round(mSprAmt_TB, 2) & ", " & Round(mSprAmt_TP, 2) & ", " & Round(mOilAmt_TB, 2) & ", " & Round(mOilAmt_TP, 2) & ", " & Round(mD_Per_TB, 2) & ", " & _
''                            "" & Round(mD_Per_TP, 2) & ", " & Round(mD_Amt_TB, 2) & ", " & Round(mD_Amt_TP, 2) & ", 0, " & Round(IIf(IsNull(RsMast.Fields("Total_Tax_Amount")), 0, IIf(IsNull(RsMast.Fields("Discount Parts")), 0, Val(Format(Mid(RsMast.Fields("Total_Tax_Amount"), 4, Len(RsMast.Fields("Total_Tax_Amount")) - 3), "0.00")))), 2) & ", " & _
''                            "" & Round(IIf(IsNull(RsMast.Fields("Other Charges")), 0, IIf(IsNull(RsMast.Fields("Other Charges")), 0, Val(Mid(RsMast.Fields("Other Charges"), 4, Len(RsMast.Fields("Other Charges")) - 3)))), 2) & ", 0, 0, 0, 0, " & _
''                            "" & Round(Val(Format(Mid(RsMast.Fields("Parts_Invoice_Amount"), 4, Len(RsMast.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) & ", " & Round(Val(Format(Mid(RsMast.Fields("Parts_Invoice_Amount"), 4, Len(RsMast.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 0) - Round(Val(Format(Mid(RsMast.Fields("Parts_Invoice_Amount"), 4, Len(RsMast.Fields("Parts_Invoice_Amount")) - 3), "0.00")), 2) & ", 1, 1, 'Siebel', " & _
''                            "" & ConvertDate(PubLoginDate) & ", 'A')"
''
''
''
''
''
''                GCn.Execute "Insert Into Sp_Sale (DocId, DocIdHelp, Site_Code, V_Type, V_No, " & _
''                            "V_Date, Party_Code, Cash_Credit, Party_Name, L_C, " & _
''                            "Form_Code, CrAc, SiebelDocId, Invoice_DocId, PType, " & _
''                            "Gp_No, Gp_Date, SprAmt_Mrp_TB, SprAmt_Mrp_TP, OilAmt_MrpTB, " & _
''                            "OilAmt_MrpTP, D_Per_Mrp_TB, D_Per_Mrp_Tp, D_Amt_Mrp_TB, D_Amt_Mrp_TP, " & _
''                            "SprAmt_TB, SprAmt_TP, OilAmt_TB, OilAmt_TP, D_Per_TB, " & _
''                            "D_Per_TP, D_Amt_TB, D_Amt_TB, Addition, Tax_Amt, " & _
''                            "Packing, Tot_Tax, Tot_Amt,  ReSalTax_Per, ReSalTax_Amt, " & _
''                            "Total_Amt, Rounded, Det_Tax, AcPostingYn, U_Name, " & _
''                            "U_EndDt, U_AE) " & _
''                            "Values ('" & mInvoiceID & "', '" & Replace(mInvoiceID, " ", "") & "', " & RsMast!AutomanSite & ", '" & Trim(mV_Type) & "', " & RsMast!Order_No & ", " & _
''                            "" & ConvertDate(MakeDate(left(RsMast!Invoice_Date, 10))) & ", '" & mPartyCode & "', '" & RsMast.Fields("Mode Of Payment") & "', '" & left(mHeaderParty, 40) & "', '" & mLocal & "', " & _
''                            "'" & mFormCode & "', '" & mDebitAc & "', '" & RsMast!Invoice_No & "', '', 'General', " & _
''                            "'" & mGatePassID & "', " & ConvertDate(MakeDate(left(RsMast.Fields("Order Date"), 10))) & ", " & Round(mSprAmt_MRP_TB, 2) & ", " & Round(mSprAmt_MRP_TP, 2) & ", " & Round(mOilAmt_MRP_TB, 2) & ", " & _
''                            "" & Round(mOilAmt_MRP_TP, 2) & ", " & Round(mD_Per_MRP_TB, 2) & ", " & Round(mD_Per_MRP_TP, 2) & ", " & Round(mD_Amt_MRP_TB, 2) & ", " & Round(mD_Amt_MRP_TP, 2) & ", " & _
''                            "" & Round(mSprAmt_TB, 2) & ", " & Round(mSprAmt_TP, 2) & ", " & Round(mOilAmt_TB, 2) & ", " & Round(mOilAmt_TP, 2) & ", " & Round(mD_Per_TB, 2) & ", " & _
''                            "" & Round(mD_Per_TP, 2) & ", " & Round(mD_Amt_TB, 2) & ", " & Round(mD_Amt_TP, 2) & ", 0, " & Round(IIf(IsNull(RsMast.Fields("Total_Tax_Amount")), 0, IIf(IsNull(RsMast.Fields("Discount Parts")), 0, Val(Format(Mid(RsMast.Fields("Total_Tax_Amount"), 4, Len(RsMast.Fields("Total_Tax_Amount")) - 3), "0.00")))), 2) & ", " & _
''                            "" & Round(IIf(IsNull(RsMast.Fields("Other Charges")), 0, IIf(IsNull(RsMast.Fields("Other Charges")), 0, Val(Mid(RsMast.Fields("Other Charges"), 4, Len(RsMast.Fields("Other Charges")) - 3)))), 2) & ", 0, 0, 0, 0, " & _
''                            "" & Round(Val(Format(Mid(RsMast!Parts_Invoice_Amount, 4, Len(RsMast!Parts_Invoice_Amount) - 3), "0.00")), 0) & ", " & Round(Val(Format(Mid(RsMast!Parts_Invoice_Amount, 4, Len(RsMast!Parts_Invoice_Amount) - 3), "0.00")), 0) - Round(Val(Format(Mid(RsMast!Parts_Invoice_Amount, 4, Len(RsMast!Parts_Invoice_Amount) - 3), "0.00")), 2) & ", 1, 1, 'Siebel', " & _
''                            "" & ConvertDate(PubLoginDate) & ", 'A')"
''
''
''
'''NextRecord:
''                CopyCnt = CopyCnt + 1
''                lblRecCopy(Index).Caption = CopyCnt
''                lblRecCopy(Index).Refresh
''
''                RsMast.MoveNext
''            Loop
''        End If
''
''
''
''
''NextRecord:
''    ImportBtn(Index).BackColor = FinishColor
''    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
''lblExit:
''    Set RsNew = Nothing
''    Exit Sub
''Eloop:
''    ErrorCnt = ErrorCnt + 1
''    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
''    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Spare Sale Bill','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''    Resume Next
''End Sub



Private Sub StkTrfUpdate(Index, TrnType)
'' On Error GoTo Eloop
Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer, mV_Type As String
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String, mOrderQty As Double, mPhysicalQty As Double
Dim mPrefix As String, mname As String, mDebitAc As String, mFormCode As String
Dim mOrderNo As String, mChallanNo As String, mGatePassID As String
Dim mSrl As Integer, mQty As Double, mCount As Integer, mAmount As Double
Dim mLocal As String
Dim mTax_Amt As Double, mTax_Amt1 As Double, mTaxable As Boolean
Dim mVATApplicable As Boolean, mTrnType As String, mLubType As String

Dim mSprAmt_MRP_TB As Double, mSprAmt_MRP_TP As Double, mOilAmt_MRP_TB As Double, mOilAmt_MRP_TP As Double
Dim mSprAmt_TB As Double, mSprAmt_TP As Double, mOilAmt_TB As Double, mOilAmt_TP As Double
Dim mD_Per_TB As Double, mD_Per_TP As Double, mD_Amt_TB As Double, mD_Amt_TP As Double
Dim mD_Per_MRP_TB As Double, mD_Per_MRP_TP As Double, mD_Amt_MRP_TB As Double, mD_Amt_MRP_TP As Double


    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    
    If TrnType = "Inward" Then
        mV_Type = "SXGRT"
        Set RsNew = New adodb.Recordset
        RsNew.CursorLocation = adUseClient
        RsNew.Open "Select * from SP_Purch", GCn, adOpenDynamic, adLockOptimistic
    Else
        mV_Type = "SYSCT"
        Set RsNew = New adodb.Recordset
        RsNew.CursorLocation = adUseClient
        RsNew.Open "Select * from SP_Sale", GCn, adOpenDynamic, adLockOptimistic
    End If
    
    Set RsNew1 = New adodb.Recordset
    RsNew1.CursorLocation = adUseClient
    RsNew1.Open "Select * from SP_Stock", GCn, adOpenDynamic, adLockOptimistic
    
    mVATApplicable = ErrorGCN.Execute("Select VatApplicable from Enviro").Fields(0).Value
    
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    Do Until Master.EOF
        'mOrderNo = XNull(Master.Fields("Order ID").Value)
        
        'mChallanNo = XNull(Master.Fields("Transaction #").Value)
        mChallanNo = Trim(Mid(Master!Narration, 33, InStr(1, Master!Narration, "Dated") - 33))
        
        
        If TrnType = "Inward" Then
            If Master!Type <> "Receive Internal" Then GoTo DuplicateSkipped
        Else
            If Master!Type <> "Ship Internal" Then GoTo DuplicateSkipped
        End If
        
        'If IsNull(StringPass(Master.Fields("Transaction #"))) Or StringPass(Master.Fields("Transaction #")) = "" Then
        If Trim(mChallanNo) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Transaction #"), "Stock Transfer (" & TrnType & ")", "Narration Field  is Empty")
            GoTo MyNextRecord
        End If
        
        If TrnType = "Inward" Then
            If GCn.Execute("Select V_no from SP_Purch where SiebelDocID='" & mChallanNo & "' and V_Type='" & mV_Type & "'").RecordCount > 0 Then
                GoTo DuplicateSkipped
            End If
        
        Else
            If GCn.Execute("Select V_no from SP_Sale where SiebelDocID='" & mChallanNo & "' and V_Type='" & mV_Type & "'").RecordCount > 0 Then
                GoTo DuplicateSkipped
            End If
        End If
        
        
        If IsNull(StringPass(Master.Fields("Transaction Date/Time"))) Or StringPass(Master.Fields("Transaction Date/Time")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (" & TrnType & ")", "Transaction Date/Time is Empty")
            GoTo MyNextRecord
        End If
            
        If IsNull(StringPass(Master.Fields("Transaction #"))) Or StringPass(Master.Fields("Transaction #")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (" & TrnType & ")", "Transaction # is Empty")
            GoTo MyNextRecord
        End If
            
        If IsNull(StringPass(Master.Fields("Part #").Value)) Or StringPass(Master.Fields("Part #").Value) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (" & TrnType & ")", "Part Number field is Empty")
            GoTo MyNextRecord
        Else
            If GCn.Execute("Select Part_No from Part where Part_No='" & Master.Fields("Part #") & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Part #"), "Stock Transfer (" & TrnType & ")", "Part Number not exist in AUTOMAN Part Master")
                'GoTo MyNextRecord
            End If
        End If
            
        If IsNull(StringPass(Master.Fields("Qty"))) Or StringPass(Master.Fields("Qty")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (" & TrnType & ")", "Qty field is Empty")
            GoTo MyNextRecord
        End If
        
                
        If IsNull(StringPass(Master.Fields("Destination Division"))) Or StringPass(Master.Fields("Destination Division")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (" & TrnType & ")", "Destination Division field is Empty")
            GoTo MyNextRecord
        Else
            If TrnType = "Inward" Then
                If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").RecordCount > 0 Then
                    mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").Fields(0).Value
                    mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").Fields(0).Value
                    mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").Fields(0).Value
                Else
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master.Fields("Destination Division")), "Stock Transfer (" & TrnType & ")", "Automan Site/Division is not Defined in SiteDivision Table for this Destination Division")
                    GoTo MyNextRecord
                End If
            End If
        End If
            
        If IsNull(StringPass(Master.Fields("Source Division"))) Or StringPass(Master.Fields("Source Division")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (" & TrnType & ")", "Source Division (Vendor Name/Supplied from) field is Empty")
            GoTo MyNextRecord
        Else
            If TrnType = "Outward" Then
                If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Source Division")) & "'").RecordCount > 0 Then
                    mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Source Division")) & "'").Fields(0).Value
                    mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Source Division")) & "'").Fields(0).Value
                    mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Source Division")) & "'").Fields(0).Value
                Else
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master.Fields("Source Division")), "Stock Transfer (" & TrnType & ")", "Automan Site/Division is not Defined in SiteDivision Table for this Source Division")
                    GoTo MyNextRecord
                End If
            End If
        End If
        
        If TrnType = "Inward" Then
            If ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Source Division")) & "' and Type='" & Master.Fields("Type") & "'").RecordCount > 0 Then
                mPartyCode = ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Source Division")) & "' and Type='" & Master.Fields("Type") & "'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Source Division"), "Stock Transfer (" & TrnType & ")", "This Source Division (Vendor Name/Supplied from) is not defined in AccountConversion Table (As Account Code)")
                GoTo MyNextRecord
            End If
        Else
            If ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Destination Division")) & "' and Type='" & Master.Fields("Type") & "'").RecordCount > 0 Then
                mPartyCode = ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Destination Division")) & "' and Type='" & Master.Fields("Type") & "'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Destination Division"), "Stock Transfer (" & TrnType & ")", "This Destination Division (Customer Name/Supplied To) is not defined in AccountConversion Table (As Account Code)")
                GoTo MyNextRecord
            End If
        End If
        
        mPrefix = "SBL" & Format(Master.Fields("Transaction Date/Time"), "yy")
        If TrnType = "Inward" Then
            mFormCode = ErrorGCN.Execute("Select SparePurchFormStockTrfIn from Enviro").Fields(0).Value
            CodeCnt = GCn.Execute("Select iif(isnull(Max(V_No)),0,Max(V_no))+1 from SP_Purch where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mV_Type & "'").Fields(0).Value
        Else
            '' For GatePass Details :
            CodeCnt = GCn.Execute("Select iif(isnull(Max(val(Left(GP_No,5)))),0,Max(val(Left(GP_no,5))))+1 from SP_Sale where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mV_Type & "'").Fields(0).Value
            mGatePassID = mRecordDiv & mRecordSite & mRecordSite & Right("00000" & CodeCnt, 5)
            
            mFormCode = ErrorGCN.Execute("Select SpareSaleFormStockTrfOut from Enviro").Fields(0).Value
            CodeCnt = GCn.Execute("Select iif(isnull(Max(V_No)),0,Max(V_no))+1 from SP_Sale where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mV_Type & "'").Fields(0).Value
        End If
        
        mDocId = mRecordDiv & mRecordSite & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
        mDebitAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc where Div_Code='" & mRecordDiv & "' and Form_Code='" & mFormCode & "'").Fields(0).Value
        mLocal = "L"
        
        'Insert New Rec
        With RsNew
            .AddNew
            !DocId = mDocId
            !DocIDHelp = Replace(mDocId, " ", "")
            !Site_Code = mRecordSite & mRecordSite
            !V_Type = Trim(mV_Type)
            !V_No = CodeCnt
            '!V_DATE = Format(Master.Fields("Transaction Date/Time"), "dd/MMM/yyyy")
            !V_DATE = MakeDate(left(Master.Fields("Transaction Date/Time"), 10))
            !Party_Code = mPartyCode
            !Cash_Credit = "Credit"
            !L_C = mLocal
            !Form_Code = mFormCode
            If TrnType = "Inward" Then
                !Party_Name = Master.Fields("Source Division")
                !Party_Doc_No = left(StringPass(Master.Fields("Transaction #")), 10)
                '!Party_Doc_Date = Format(Master.Fields("Transaction Date/Time"), "dd/MMM/yyyy")
                !Party_Doc_Date = MakeDate(left(Master.Fields("Transaction Date/Time"), 10))
                !DrAc_Code = mDebitAc
            Else
                !Party_Name = Master.Fields("Destination Division")
                !CrAc = mDebitAc
            End If
            !SiebelDocID = mChallanNo
            !U_Name = "Siebel"
            !U_EntDt = Format(PubLoginDate, "Short Date")
            !U_AE = "A"
        End With
            
        mSprAmt_MRP_TB = 0: mSprAmt_MRP_TP = 0: mOilAmt_MRP_TB = 0: mOilAmt_MRP_TP = 0
        mD_Per_MRP_TB = 0: mD_Per_MRP_TP = 0: mD_Amt_MRP_TB = 0: mD_Amt_MRP_TP = 0
        
        mSprAmt_TB = 0: mSprAmt_TP = 0: mOilAmt_TB = 0: mOilAmt_TP = 0
        mD_Per_TB = 0: mD_Per_TP = 0: mD_Amt_TB = 0: mD_Amt_TP = 0
            
        mSrl = 1
        mQty = 0
        mCount = 0
        mAmount = 0
        mTax_Amt = 0
        mTax_Amt1 = 0
        If (Master.EOF = False And Master.BOF = False) Then
            Do While mChallanNo = Trim(Mid(Master!Narration, 33, InStr(1, Master!Narration, "Dated") - 33))   '' XNull(Master.Fields("Transaction #"))
                If TrnType = "Inward" Then
                    If Master!Type <> "Receive Internal" Then GoTo LineFileNextrecord
                Else
                    If Master!Type <> "Ship Internal" Then GoTo LineFileNextrecord
                End If
                If mSrl <> 1 Then
                    If IsNull(StringPass(Master.Fields("Part #").Value)) Or StringPass(Master.Fields("Part #").Value) = "" Then
                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (" & TrnType & ")", "Part Number field is Empty")
                        GoTo LineFileNextrecord
                    Else
                        If GCn.Execute("Select Part_No from Part where Part_No='" & Master.Fields("Part #") & "'").RecordCount = 0 Then
                            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Part #"), "Stock Transfer (" & TrnType & ")", "Part Number not exist in AUTOMAN Part Master")
                            'GoTo MyNextRecord      '' skipping of record not required at this position
                        End If
                    End If
                End If
                
                'Insert New Rec
                With RsNew1
                    .AddNew
                    !DocId = mDocId
                    !Site_Code = mRecordSite & mRecordSite
                    !V_Type = Trim(mV_Type)
                    !V_No = CodeCnt
                    !V_DATE = MakeDate(left(Master.Fields("Transaction Date/Time"), 10))
                    !Party_Code = mPartyCode
                    !Srl_No = mSrl
                    !L_C = mLocal
                    !Part_No = VNull(Master.Fields("part #"))
                    !godown = GCn.Execute("Select SprWorksGodown from Syctrl").Fields(0).Value
                    If TrnType = "Inward" Then
                        !Qty_Doc = Val(Master!Qty)
                        !Qty_Rec = Val(Master!Qty)
                    Else
                        !Qty_Iss = Val(Master!Qty)
                    End If
                    If mVATApplicable Then
                        !Tax_YN = 1             '' if VAT is applicable in State
                    Else
                        !Tax_YN = IIf(mLocal = "L", 0, 1)
                    End If
                    !MRP_YN = IIf(mRecordDiv = "C", 1, 0)
                    !Amount = Val(Master!Value)
                    !Net_Amt = Val(Master!Value)
                    !Rate = Round(Master!Value / Master!Qty, 2) ' Master!Rate
                    !Part_SrlNo = mSrl
                    
                    mTrnType = GCn.Execute("Select Part_Grade from Part where part_no='" & Master.Fields("part #") & "'").Fields(0).Value
                    If mRecordDiv = "C" Then
                        If !Tax_YN = 1 Then
                            If mLubType = mTrnType Then
                                mOilAmt_MRP_TB = mOilAmt_MRP_TB + IIf(IsNull(Master.Fields("Value")), 0, Master.Fields("Value"))
                            Else
                                mSprAmt_MRP_TB = mSprAmt_MRP_TB + IIf(IsNull(Master.Fields("Value")), 0, Master.Fields("Value"))
                            End If
                        Else
                            If mLubType = mTrnType Then
                                mOilAmt_MRP_TP = mOilAmt_MRP_TP + IIf(IsNull(Master.Fields("Value")), 0, Master.Fields("Value"))
                            Else
                                mSprAmt_MRP_TP = mSprAmt_MRP_TP + IIf(IsNull(Master.Fields("Value")), 0, Master.Fields("Value"))
                            End If
                        End If
                    Else
                        If !Tax_YN = 1 Then
                            If mLubType = mTrnType Then
                                mOilAmt_TB = mOilAmt_TB + IIf(IsNull(Master.Fields("Value")), 0, Master.Fields("Value"))
                            Else
                                mSprAmt_TB = mSprAmt_TB + IIf(IsNull(Master.Fields("Value")), 0, Master.Fields("Value"))
                            End If
                        Else
                            If mLubType = mTrnType Then
                                mOilAmt_TP = mOilAmt_TP + IIf(IsNull(Master.Fields("Value")), 0, Master.Fields("Value"))
                            Else
                                mSprAmt_TP = mSprAmt_TP + IIf(IsNull(Master.Fields("Value")), 0, Master.Fields("Value"))
                            End If
                        End If
                    End If
                    
                    !TaxAmt = 0
                    !TaxPer = 0
                    !Disc_Per = 0
                    !Disc_Amt = 0
                    !Ord_Discper = 0
                    !Ord_DiscAmt = 0
        
                    !U_Name = "Siebel"
                    !U_EntDt = Format(PubLoginDate, "Short Date")
                    !U_AE = "A"
                    .Update
                End With
                mQty = mQty + Master!Qty
                mCount = mCount + 1
                mAmount = mAmount + Val(Master!Value)
                mSrl = mSrl + 1
LineFileNextrecord:
                Master.MoveNext
                CopyCnt = CopyCnt + 1
                lblRecCopy(Index).Caption = CopyCnt
                lblRecCopy(Index).Refresh
                If Master.EOF = True Then Exit Do
            Loop
        End If
        If Master.EOF = True Then Master.MovePrevious
        With RsNew
            If TrnType = "Inward" Then
                !Tot_No_Of_Items = mCount
                !Tot_Doc_Qty = mQty
                !Tot_Phy_Qty = mQty
                !TOT_Amt = mAmount
                !Tot_Disc_Amt = 0
                !Tot_Ord_DiscAmt = 0
                !Tot_Goods_value = mAmount
                !Tax_Amt = 0
                !Addition = 0
                !Deduction = 0
                !Net_Amt = mAmount
            Else
                !GP_No = mGatePassID
                If Not IsNull(Master.Fields("Transaction Date/Time")) Then
                !GP_Date = Format(Master.Fields("Transaction Date/Time"), "dd/MMM/yyyy")
                End If
                
                !SprAmt_MRP_TB = mSprAmt_MRP_TB
                !SprAmt_MRP_TP = mSprAmt_MRP_TP
                !OilAmt_MRP_TB = mOilAmt_MRP_TB
                !OilAmt_MRP_TP = mOilAmt_MRP_TP
                !D_Per_MRP_TB = mD_Per_MRP_TB
                !D_Per_MRP_TP = mD_Per_MRP_TP
                !D_Amt_MRP_TB = mD_Amt_MRP_TB
                !D_Amt_MRP_TP = mD_Amt_MRP_TP
                
                !SprAmt_TB = mSprAmt_TB
                !SprAmt_TP = mSprAmt_TP
                !OilAmt_TB = mOilAmt_TB
                !OilAmt_TP = mOilAmt_TP
                !D_Per_TB = mD_Per_TB
                !D_Per_TP = mD_Per_TP
                !D_Amt_TB = mD_Amt_TB
                !D_Amt_TP = mD_Amt_TP
                
                !total_amt = mAmount
                !Rounded = 0
                
                !Det_Tax = 1
                !AcPosting_YN = 1
            End If
            .Update
        End With
        If Master.AbsolutePosition = Master.RecordCount Then Master.MoveNext
        CodeCnt = CodeCnt + 1
        GoTo NextLoop
        
DuplicateSkipped:

MyNextRecord:
        mSrl = 0
        If Master.EOF = False And Master.BOF = False Then
                Do While (Master.EOF = False And Master.BOF = False) And mChallanNo = Trim(Mid(Master!Narration, 33, InStr(1, Master!Narration, "Dated") - 33))   '' XNull(Master.Fields("Transaction #"))
                mSrl = mSrl + 1
                Master.MoveNext
                CopyCnt = CopyCnt + 1
                lblRecCopy(Index).Caption = CopyCnt
                lblRecCopy(Index).Refresh
                If Master.EOF = True Then Exit Do
            Loop
        End If
'        CopyCnt = CopyCnt + mSrl - 1
'        lblRecCopy(Index).Caption = CopyCnt
'        lblRecCopy(Index).Refresh
NextLoop:
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    mSrl = 0
    Do Until mChallanNo = Mid(Master!Narration, 33, InStr(1, Master!Narration, "Dated") - 33)   ''Master.Fields("Transaction #")            ''mOrderNo = Master.Fields("Order ID") And
        mSrl = mSrl + 1
        Master.MoveNext
    Loop
    
    ErrorCnt = ErrorCnt + mSrl
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Stock Transfer (" & TrnType & ")','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub

Private Sub PurchBillDataUpdate(Index, TrnType)
Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer, mV_Type As String
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String, mOrderQty As Double, mPhysicalQty As Double
Dim mPrefix As String, mname As String, mLubType As String, mTrnType As String, mDebitAc As String, mFormCode As String
Dim mChallanNo As String, mHeaderParty As String
Dim mQty As Double, mCount As Integer, mAmount As Double
Dim mInvoiceNo As String, mChallanID As String
Dim mTax_Amt As Double, mTax_Amt1 As Double, mLocal As String
    
    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from SP_Purch", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsNew1 = New adodb.Recordset
    RsNew1.CursorLocation = adUseClient
    RsNew1.Open "Select * from SP_Stock", GCn, adOpenDynamic, adLockOptimistic
    
    mV_Type = "SXPIR"
    
    If Master.RecordCount > 0 Then Master.MoveFirst
    
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Invoice #"))) Or StringPass(Master.Fields("Invoice #")) = "" Then GoTo MyNextRecord
                
        mInvoiceNo = StringPass(Master.Fields("Invoice #").Value)
        
        
        Set Master1 = CreateObject("ADODB.Recordset")
        GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] where [Invoice #]=" & mInvoiceNo & " Order By  [Invoice #]"
        Master1.Open GSQL, ExcelGcn2, adOpenStatic
                        
                
        If Master1.RecordCount = 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Invoice #"), "Telco Purchase Bill Data", "Details of Line File Item not found for Invoice No. " & Master.Fields("Invoice #"))
            GoTo MyNextRecord
        End If
                
        If IsNull(StringPass(Master.Fields("Vendor Name"))) Or StringPass(Master.Fields("Vendor Name")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Telco Purchase Bill Data", "Vendor Name field is Empty")
            GoTo MyNextRecord
        End If
        
        mHeaderParty = Master.Fields("Vendor Name")
        
        If ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Vendor Name")) & "' and Type='Tata Spare Purchase'").RecordCount > 0 Then
            mPartyCode = ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Vendor Name")) & "' and Type='Tata Spare Purchase'").Fields(0).Value
        Else
            Call InsSkipRecMessage(Index, mInvoiceNo, Master.Fields("Vendor Name"), "Tata Purchase Bill Data", "This Vendor Account Code is not defined in AccountConversion Table for Inv No. " & mInvoiceNo)
            GoTo MyNextRecord
        End If
                
        If GCn.Execute("Select V_no from SP_Purch where Party_Doc_No='" & StringPass(Master.Fields("Invoice #")) & "' and v_Type='" & mV_Type & "' and Party_Code='" & mPartyCode & "'").RecordCount > 0 Then
            GoTo DuplicateSkipped
        End If
        
        If IsNull(StringPass(Master.Fields("Division"))) Or StringPass(Master.Fields("Division")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Telco Purchase Bill Data", "Division Name field is Empty")
            GoTo MyNextRecord
        Else
            If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").RecordCount > 0 Then
                mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").Fields(0).Value
                mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").Fields(0).Value
                mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Division")) & "'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master.Fields("Division")), "Telco Purchase Bill Data", "Automan Site/Division is not Defined in SiteDivision Table for this Siebel Division")
                GoTo MyNextRecord
            End If
        End If
            
        
        If IsNull(StringPass(Master1.Fields("Part #").Value)) Or StringPass(Master1.Fields("Part #").Value) = "" Then
            Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Telco Purchase Bill Data", "Part Number field is Empty")
            GoTo MyNextRecord
        Else
            If GCn.Execute("Select Part_No from Part where Part_No='" & Master1.Fields("Part #") & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master1.AbsolutePosition, Master1.Fields("Part #"), "Telco Purchase Bill Data", "Part Number not exist in AUTOMAN Part Master")
                'GoTo MyNextRecord
            End If
        End If
            
        If IsNull(StringPass(Master1.Fields("UoM"))) Or StringPass(Master1.Fields("UoM")) = "" Then
            Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Telco Purchase Bill Data", "Part Unit field is Empty")
            GoTo MyNextRecord
        End If
            
        If IsNull(StringPass(Master1.Fields("Qty"))) Or StringPass(Master1.Fields("Qty")) = "" Then
            Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Telco Purchase Bill Data", "Qty field is Empty")
            GoTo MyNextRecord
        End If
        
        mPrefix = "SBL" & Format(Master.Fields("Invoice_Date"), "yy")
        CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from SP_Purch where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and V_Type='" & mV_Type & "'").Fields(0).Value
        mDocId = mRecordDiv & mRecordSite & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
        mLubType = GCn.Execute("Select PartGrade_Lub from Syctrl").Fields(0).Value
        If GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("part #") & "'").RecordCount > 0 Then
            mTrnType = GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("part #") & "'").Fields(0).Value
        Else
            mTrnType = "S"
        End If
        
        
        If mTrnType = mLubType Then
            mFormCode = ErrorGCN.Execute("Select SparePurchFormLubs from Enviro").Fields(0).Value
            mTax_Amt = Val(Master.Fields("Total_Tax_Amount"))
            mLocal = "L"
        Else
            If Not IsNull(Master.Fields("Total_Tax_Amount")) Then
                If Val(Master.Fields("Total_Tax_Amount")) > 0 Then
                    mFormCode = ErrorGCN.Execute("Select SparePurchFormVAT from Enviro").Fields(0).Value
                    mTax_Amt = Val(Master.Fields("Total_Tax_Amount"))
                    mLocal = "L"
                End If
            End If
            If Not IsNull(Master.Fields("LST")) Then
                If Val(Master.Fields("LST")) > 0 Then
                    mFormCode = ErrorGCN.Execute("Select SparePurchFormLocal from Enviro").Fields(0).Value
                    mTax_Amt = Val(Master.Fields("Total_Tax_Amount"))
                    mLocal = "L"
                End If
            End If
            If Not IsNull(Master.Fields("CST")) Then
                If Val(Master.Fields("CST")) > 0 Then
                    mFormCode = ErrorGCN.Execute("Select SparePurchFormCST from Enviro").Fields(0).Value
                    mTax_Amt = Val(Master.Fields("Total_Tax_Amount"))
                    mLocal = "C"
                End If
            End If
            
            If Not IsNull(Master1.Fields("CST ON VAT")) Then
                If Val(Replace(Mid(Master1.Fields("CST ON VAT"), 4, Len(Master1.Fields("CST ON VAT")) - 3), "'", "")) > 0 Then
                    mFormCode = ErrorGCN.Execute("Select SparePurchFormCST from Enviro").Fields(0).Value
                    mLocal = "C"
                End If
            End If
        End If
        If mFormCode = "" Then
            mFormCode = ErrorGCN.Execute("Select SparePurchFormVAT from Enviro").Fields(0).Value
            mTax_Amt = Val(Master1.Fields("Total_Tax_Amount"))
            mLocal = "L"
        End If
        
        mDebitAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc where Div_Code='" & mRecordDiv & "' and Form_Code='" & mFormCode & "'").Fields(0).Value
        
        mQty = 0
        mCount = 0
        mAmount = 0
        mTax_Amt = 0
        mTax_Amt1 = 0
        
        'Do While mInvoiceNo = XNull(Master1.Fields("Invoice #")) And Not Master1.EOF
        Master1.MoveFirst
        Do Until Master1.EOF
            If mHeaderParty <> Master1.Fields("Vendor Name") Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Invoice #"), "Telco Purchase Bill Data", "Party Name is different in Header & Line File")
                GoTo MyNextRecord
            End If
            ' Tax Value
            If UCase(left(PubComp_Name, 3)) = "LMP" Then
                If Not IsNull(Master1.Fields("LST")) Then
                    If Val(Mid(Master1.Fields("LST"), 4, Len(Master1.Fields("LST")) - 3)) > 0 Then
                        mTax_Amt1 = mTax_Amt1 + IIf(IsNull(Master1.Fields("LST")), 0, Mid(Master1.Fields("LST"), 4, Len(Master1.Fields("LST")) - 3))
                    End If
                End If
                If Not IsNull(Master1.Fields("CST")) Then
                    If Val(Mid(Master1.Fields("CST"), 4, Len(Master1.Fields("CST")) - 3)) > 0 Then
                        mTax_Amt1 = mTax_Amt1 + IIf(IsNull(Master1.Fields("CST")), 0, Mid(Master1.Fields("CST"), 4, Len(Master1.Fields("CST")) - 3))
                    End If
                End If
                If Not IsNull(Master1.Fields("VAT")) Then
                    If Val(Mid(Master1.Fields("VAT"), 4, Len(Master1.Fields("VAT")) - 3)) > 0 Then
                        mTax_Amt1 = mTax_Amt1 + IIf(IsNull(Master1.Fields("VAT")), 0, Mid(Master1.Fields("VAT"), 4, Len(Master1.Fields("VAT")) - 3))
                    End If
                End If
            Else
                If Not IsNull(Master1.Fields("Total_Tax_Amount")) Then
                    If Val(Mid(Master1.Fields("Total_Tax_Amount"), 4, Len(Master1.Fields("Total_Tax_Amount")) - 3)) > 0 Then
                        mTax_Amt1 = mTax_Amt1 + IIf(IsNull(Master1.Fields("Total_Tax_Amount")), 0, Mid(Master1.Fields("Total_Tax_Amount"), 4, Len(Master1.Fields("Total_Tax_Amount")) - 3))
                    End If
                End If
            End If
            mAmount = mAmount + Round(Val(Format(Mid(Master1.Fields("Net Amount"), 4, Len(Master1.Fields("Net Amount")) - 3), "0.00")), 2)
            
            mQty = mQty + Master1!Qty
            mCount = mCount + 1
            
            Master1.MoveNext
        Loop
        
        If Round(mAmount, 1) <> Round(IIf(IsNull(Master.Fields("Net_Amount")), 0, Master.Fields("Net_Amount")), 1) Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Invoice #"), "Telco Purchase Bill Data", "Line File Goods Value Total(" & Val(mAmount) & ") is not matched with Header file Goods Value(" & VNull(Master.Fields("Net_Amount")) & ")")
            GoTo MyNextRecord
        End If
        If Round(mTax_Amt1, 1) <> Round(IIf(IsNull(Master.Fields("Total_Tax_Amount")), 0, Master.Fields("Total_Tax_Amount")), 1) Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Invoice #"), "Telco Purchase Bill Data", "Line File Tax Value Total is not matched with Header file Tax Value")
            GoTo MyNextRecord
        End If
        
        If GCn.Execute("Select DocID From Sp_Purch where Party_Doc_No='" & mInvoiceNo & "' and Party_Code='" & mPartyCode & "' and V_Type='SXGR'").RecordCount = 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Invoice #"), "Telco Purchase Bill Data", "Spare MRN Details not found in Automan against This Purchase Invoice/Challan")
            GoTo MyNextRecord
        End If
        
        'Insert New Rec
        With RsNew
            .AddNew
            !DocId = mDocId
            !DocIDHelp = Replace(mDocId, " ", "")
            !Site_Code = mRecordSite & mRecordSite
            !V_Type = Trim(mV_Type)
            !V_No = CodeCnt
            '!V_DATE = Format(Master.Fields("Invoice_Date"), "dd/MMM/yyyy")
            !V_DATE = MakeDate(Master.Fields("Invoice_Date"))
            !Party_Code = mPartyCode
            !Cash_Credit = "Credit"
            !Party_Name = Master.Fields("Vendor Name")
            !L_C = mLocal
            !Form_Code = mFormCode
            !Party_Doc_No = StringPass(Master.Fields("Invoice #"))
            '!Party_Doc_Date = Format(Master.Fields("Invoice_Date"), "dd/MMM/yyyy")
            !Party_Doc_Date = MakeDate(Master.Fields("Invoice_Date"))
            !DrAc_Code = mDebitAc
            !SiebelDocID = Master.Fields("Invoice #")
            
            !Tot_No_Of_Items = mCount
            !Tot_Doc_Qty = mQty
            !Tot_Phy_Qty = mQty
            If mTrnType = mLubType Then
                !OilAmt = mAmount
            Else
                !SprAmt = mAmount
            End If
            
            !TOT_Amt = mAmount
            !Tot_Disc_Amt = 0
            !Tot_Ord_DiscAmt = 0
            !Tot_Goods_value = mAmount
            !Tax_Amt = IIf(IsNull(Master.Fields("Total_Tax_Amount")), 0, Master.Fields("Total_Tax_Amount"))
            !Addition = 0
            !Deduction = 0
            !Net_Amt = IIf(IsNull(Master.Fields("Total_Tax_Amount")), 0, Master.Fields("Total_Tax_Amount")) + mAmount
            !U_Name = "Siebel"
            !U_EntDt = Format(PubLoginDate, "Short Date")
            !U_AE = "A"
            .Update
        End With
        
        '' Updation of MRN Header & LIne file for Purch Bill
'        Master1.MoveFirst
'        Master1.Find ("[Invoice #]='" & Master.Fields("Invoice #").Value & "'")
'        Do While mInvoiceNo = XNull(Master1.Fields("Invoice #")) And Not Master1.EOF
        Master1.MoveFirst
        Do Until Master1.EOF
            mChallanNo = Master1.Fields("Invoice #")
            If GCn.Execute("Select DocID From Sp_Purch where Party_Doc_No='" & mChallanNo & "' and Party_Code='" & mPartyCode & "' and V_Type='SXGR'").RecordCount > 0 Then     '
                mChallanID = GCn.Execute("Select DocID From Sp_Purch where Party_Doc_No='" & mChallanNo & "' and Party_Code='" & mPartyCode & "' and V_Type='SXGR'").Fields(0).Value
                GCn.Execute ("Update Sp_Purch set Invoice_DocID='" & mDocId & "' where DocID='" & mChallanID & "'")
                GCn.Execute ("Update Sp_Stock set Invoice_DocID='" & mDocId & "'," & _
                             "v_Date2=" & ConvertDate(Format(left(Master.Fields("Invoice_Date"), 10), "dd/MMM/yyyy")) & _
                             ", Rate2=Rate, Amount2 =Amount,Net_Amt2=Net_Amt where DocID='" & mChallanID & "'")
            End If
            Do While mChallanNo = Master1.Fields("Invoice #") 'And Master1.EOF = False
                Master1.MoveNext
                If Master1.EOF = True Then Exit Do
                If mChallanNo <> XNull(Master1.Fields("Invoice #")) Then Exit Do
            Loop
        Loop
        CodeCnt = CodeCnt + 1
DuplicateSkipped:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
MyNextRecord:
        Master.MoveNext
NextLoop:
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Telco Purchase Bill Data','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub
Private Sub MRNDataUpdate(Index, TrnType)
'' On Error GoTo Eloop
Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer, mV_Type As String
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String, mOrderQty As Double, mPhysicalQty As Double
Dim mPrefix As String, mname As String, mLubType As String, mTrnType As String, mDebitAc As String, mFormCode As String
Dim mOrderNo As String, mChallanNo As String
Dim mSrl As Integer, mQty As Double, mCount As Integer, mAmount As Double
Dim TranFalg As Boolean, TranFlag1 As Boolean, mLocal As String
Dim mTax_Amt As Double, mTax_Amt1 As Double, mTaxable As Boolean
Dim mVATApplicable As Boolean

    ImportBtn(Index).BackColor = ProcessColor
    'Master1.Sort = Master1.Fields("Order #").Name & Master1.Fields("Challan #").Name
    
    GCn.BeginTrans
    
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from SP_Purch", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsNew1 = New adodb.Recordset
    RsNew1.CursorLocation = adUseClient
    RsNew1.Open "Select * from SP_Stock", GCn, adOpenDynamic, adLockOptimistic
    
    mVATApplicable = ErrorGCN.Execute("Select VatApplicable from Enviro").Fields(0).Value
    
    TranFalg = False
    TranFlag1 = False
    mV_Type = " SXGR"
    
    If Master1.RecordCount > 0 Then Master1.MoveFirst
    
    
    Do Until Master1.EOF
        If TrnType = "Local" Then
            mOrderNo = XNull(Master1.Fields("Order #").Value)
            mChallanNo = XNull(Master1.Fields("Challan #").Value)
            If IsNull(StringPass(Master1.Fields("Order #"))) Or StringPass(Master1.Fields("Order #")) = "" Then GoTo MyNextRecord
        Else
            If IsNull(StringPass(Master1.Fields("Invoice #"))) Or StringPass(Master1.Fields("Invoice #")) = "" Then GoTo MyNextRecord
            
            mOrderNo = XNull(Master1.Fields("Invoice #").Value)
            mChallanNo = XNull(Master1.Fields("Invoice #").Value)
            
            Set Master = CreateObject("ADODB.Recordset")
            GSQL = "Select * FROM [" & ImportTxt(Index).Text & "$] where [Invoice #]=" & XNull(Master1.Fields("Invoice #").Value) & " Order By [Invoice #],[Order #]"
            Master.Open GSQL, ExcelGcn1, adOpenStatic
            
            If Master.RecordCount > 0 Then
                Master.MoveFirst
            Else
                Call InsSkipRecMessage(Index, Master1.AbsolutePosition, Master1.Fields("Invoice #"), "Spare Challan", "Invoice Details not found in Header File")
                GoTo MyNextRecord
            End If
        End If
        
        If GCn.Execute("Select V_no from SP_Purch where SiebelDocID='" & mOrderNo & "' and v_Type='" & Trim(mV_Type) & "'").RecordCount > 0 Then
            GoTo DuplicateSkipped
        End If
        
        If TrnType = "Local" Then
            If IsNull(StringPass(Master1.Fields("Challan Date"))) Or StringPass(Master1.Fields("Challan Date")) = "" Then
                If IsNull(StringPass(Master1.Fields("Last Updated Date"))) Or StringPass(Master1.Fields("Last Updated Date")) = "" Then
                    Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Spare Challan", "Challan Date is Empty")
                    GoTo MyNextRecord
                End If
            End If
        Else
'            If IsNull(StringPass(Master.Fields("Invoice_Date"))) Or StringPass(Master.Fields("Invoice_Date")) = "" Then
'                Call InsSkipRecMessage(Index, Master1.Fields("Invoice #"), "", "Spare Challan", "Invoice Date is Empty")
'                GoTo MyNextRecord
'            End If
        End If
        
        
        If IsNull(StringPass(Master1.Fields("Last Updated Date"))) Or StringPass(Master1.Fields("Last Updated Date")) = "" Then
            Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Spare Challan", "Last Updated Date is Empty")
            GoTo MyNextRecord
        End If
            
        If IsNull(StringPass(Master1.Fields("Part #").Value)) Or StringPass(Master1.Fields("Part #").Value) = "" Then
            Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Spare Challan", "Part Number field is Empty")
            GoTo MyNextRecord
        Else
            If GCn.Execute("Select Part_No from Part where Part_No='" & Master1.Fields("Part #") & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master1.AbsolutePosition, Master1.Fields("Part #"), "Spare Challan", "Part Number not exist in AUTOMAN Part Master")
                'GoTo MyNextRecord
            End If
        End If
            
        If IsNull(StringPass(Master1.Fields("UoM"))) Or StringPass(Master1.Fields("UoM")) = "" Then
            Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Spare Challan", "Part Unit field is Empty")
            GoTo MyNextRecord
        End If
            
        If IsNull(StringPass(Master1.Fields("Qty"))) Or StringPass(Master1.Fields("Qty")) = "" Then
            Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Spare Challan", "Qty field is Empty")
            GoTo MyNextRecord
        End If
        
        If TrnType = "Local" Then
            If IsNull(StringPass(Master1.Fields("Division"))) Or StringPass(Master1.Fields("Division")) = "" Then
                Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Spare Challan", "Division Name field is Empty")
                GoTo MyNextRecord
            Else
                If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master1!division) & "'").RecordCount > 0 Then
                    mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master1!division) & "'").Fields(0).Value
                    mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master1!division) & "'").Fields(0).Value
                    mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master1!division) & "'").Fields(0).Value
                Else
                    Call InsSkipRecMessage(Index, Master1.AbsolutePosition, StringPass(Master1!division), "Spare Challan", "Automan Site/Division is not Defined in SiteDivision Table for this Siebel Division")
                    GoTo MyNextRecord
                End If
            End If
        Else
            If IsNull(StringPass(Master1.Fields("Division Name"))) Or StringPass(Master1.Fields("Division Name")) = "" Then
                Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Spare Challan", "Division Name field is Empty")
                GoTo MyNextRecord
            Else
                If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master1.Fields("Division Name")) & "'").RecordCount > 0 Then
                    mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master1.Fields("Division Name")) & "'").Fields(0).Value
                    mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master1.Fields("Division Name")) & "'").Fields(0).Value
                    mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master1.Fields("Division Name")) & "'").Fields(0).Value
                Else
                    Call InsSkipRecMessage(Index, Master1.AbsolutePosition, StringPass(Master1.Fields("Division Name")), "Spare Challan", "Automan Site/Division is not Defined in SiteDivision Table for this Siebel Division")
                    GoTo MyNextRecord
                End If
            End If
        End If
            
        If IsNull(StringPass(Master1.Fields("Vendor Name"))) Or StringPass(Master1.Fields("Vendor Name")) = "" Then
            Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Spare Challan", "Vendor Name field is Empty")
            GoTo MyNextRecord
        End If
        
        If TrnType = "Local" Then
            mLength = Len(left(Trim(StringPass(Master1.Fields("Vendor Name"))), 40))
            If GCn.Execute("Select Name from Subgroup where left(name, " & mLength & ")='" & left(Trim(StringPass(Master1.Fields("Vendor Name"))), 40) & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master1.AbsolutePosition, StringPass(Master1.Fields("Vendor Name")), "Spare Challan", "Vendor Name not found in Ledger Account Master")
                GoTo MyNextRecord
            Else
                mPartyCode = GCn.Execute("Select SubCode from Subgroup where left(name, " & mLength & ")='" & left(Trim(StringPass(Master1.Fields("Vendor Name"))), 40) & "'").Fields(0).Value
            End If
        Else
            If ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Vendor Name")) & "' and Type='Tata Spare Purchase'").RecordCount > 0 Then
                mPartyCode = ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Vendor Name")) & "' and Type='Tata Spare Purchase'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master1.AbsolutePosition, Master.Fields("Vendor Name"), "Tata Spare Purchase", "This Vendor Account Code (In Header File) is not defined in AccountConversion Table")
                GoTo MyNextRecord
            End If
        End If
        
        CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from SP_Purch where Left(DocID,1)='" & mRecordDiv & "' and " & cMID("DocID", "2", "2") & "='" & mRecordSite & mRecordSite & "' and v_type='" & Trim(mV_Type) & "'").Fields(0).Value
        mPrefix = "SBL" & Format(Master1.Fields("Last Updated Date"), "yy")
        mDocId = mRecordDiv & mRecordSite & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
        mLubType = GCn.Execute("Select PartGrade_Lub from Syctrl").Fields(0).Value
        If GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("part #") & "'").RecordCount > 0 Then
            mTrnType = GCn.Execute("Select Part_Grade from Part where part_no='" & Master1.Fields("part #") & "'").Fields(0).Value
        Else
            mTrnType = "S"
        End If
        
        
        If TrnType = "Local" Then
            If mTrnType = mLubType Then
                mFormCode = ErrorGCN.Execute("Select SparePurchFormLubs from Enviro").Fields(0).Value
            Else
                mFormCode = ErrorGCN.Execute("Select SparePurchFormLocal from Enviro").Fields(0).Value
            End If
            mTax_Amt = 0
            mLocal = "L"
        Else
            If mTrnType = mLubType Then
                mFormCode = ErrorGCN.Execute("Select SparePurchFormLubs from Enviro").Fields(0).Value
                mLocal = "L"
            Else
                If Not IsNull(Master1.Fields("LST")) Then
                    If Val(Mid(Master1.Fields("LST"), 4, Len(Master1.Fields("LST")) - 3)) > 0 Then
                        mFormCode = ErrorGCN.Execute("Select SparePurchFormLocal from Enviro").Fields(0).Value
                        mLocal = "L"
                    End If
                End If
                If Not IsNull(Master1.Fields("CST")) Then
                    If Val(Mid(Master1.Fields("CST"), 4, Len(Master1.Fields("CST")) - 3)) > 0 Then
                        mFormCode = ErrorGCN.Execute("Select SparePurchFormCST from Enviro").Fields(0).Value
                        mLocal = "C"
                    End If
                End If
                If Not IsNull(Master1.Fields("VAT")) Then
                    If Val(Mid(Master1.Fields("VAT"), 4, Len(Master1.Fields("VAT")) - 3)) > 0 Then
                        mFormCode = ErrorGCN.Execute("Select SparePurchFormVAT from Enviro").Fields(0).Value
                        mLocal = "L"
                    End If
                End If
                If Not IsNull(Master1.Fields("CST ON VAT")) Then
                    If Val(Replace(Mid(Master1.Fields("CST ON VAT"), 4, Len(Master1.Fields("CST ON VAT")) - 3), "'", "")) > 0 Then
                        mFormCode = ErrorGCN.Execute("Select SparePurchFormCST from Enviro").Fields(0).Value
                        mLocal = "C"
                    End If
                End If
                
            End If
            If mFormCode = "" Then
                mFormCode = ErrorGCN.Execute("Select SparePurchFormVAT from Enviro").Fields(0).Value
                mLocal = "L"
            End If
        
        End If
        mDebitAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc where Div_Code='" & mRecordDiv & "' and Form_Code='" & mFormCode & "'").Fields(0).Value
        'Master1.MovePrevious
        
        'Insert New Rec
    
        
        With RsNew
            .AddNew
            TranFalg = True
            !DocId = mDocId
            !DocIDHelp = Replace(mDocId, " ", "")
            !Site_Code = mRecordSite & mRecordSite
            !V_Type = Trim(mV_Type)
            !V_No = CodeCnt
            
            '!V_DATE = Format(Master1.Fields("Last Updated Date"), "dd/MMM/yyyy")
            !V_DATE = MakeDate(left(Master1.Fields("Last Updated Date"), 10))
            !Party_Code = mPartyCode
            !Cash_Credit = "Credit"
            !Party_Name = Master1.Fields("Vendor Name")
            !L_C = mLocal
            !Form_Code = mFormCode
            !Party_Doc_No = left(mChallanNo, 10) 'left(StringPass(Master1.Fields("Challan #")), 10)
            If IsNull(Master1.Fields("Challan Date")) Or Master1.Fields("Challan Date") = "" Then
                !Party_Doc_Date = MakeDate(left(Master1.Fields("Last Updated Date"), 10))
            Else
                !Party_Doc_Date = MakeDate(IIf(IsNull(Master1.Fields("Challan Date")), Master1.Fields("Challan Date"), Format(Master1.Fields("Challan Date"), "dd/MMM/yyyy")))
            End If
            !DrAc_Code = mDebitAc
            !SiebelDocID = mOrderNo
            
            !U_Name = "Siebel"
            !U_EntDt = Format(PubLoginDate, "Short Date")
            !U_AE = "A"
        End With
            
        mSrl = 1
        mQty = 0
        mCount = 0
        mAmount = 0
        mTax_Amt = 0
        mTax_Amt1 = 0
        If (Master1.EOF = False And Master1.BOF = False) Then
            Do While IIf(TrnType = "Local", (mOrderNo = Master1.Fields("Order #") And mChallanNo = XNull(Master1.Fields("Challan #"))), mChallanNo = XNull(Master1.Fields("Invoice #")))
                If mSrl <> 1 Then
                    If IsNull(StringPass(Master1.Fields("Part #").Value)) Or StringPass(Master1.Fields("Part #").Value) = "" Then
                        Call InsSkipRecMessage(Index, Master1.AbsolutePosition, "", "Spare Challan", "Part Number field is Empty")
                        GoTo LineFileNextrecord
                    Else
                        If GCn.Execute("Select Part_No from Part where Part_No='" & Master1.Fields("Part #") & "'").RecordCount = 0 Then
                            Call InsSkipRecMessage(Index, Master1.AbsolutePosition, Master1.Fields("Part #"), "Spare Challan", "Part Number not exist in AUTOMAN Part Master")
                            'GoTo MyNextRecord      '' skipping of record not required at this position
                        End If
                    End If
                End If
                
                'Insert New Rec
                With RsNew1
                    .AddNew
                    TranFlag1 = True
                    !DocId = mDocId
                    !Site_Code = mRecordSite & mRecordSite
                    !V_Type = Trim(mV_Type)
                    !V_No = CodeCnt
                    !V_DATE = MakeDate(left(IIf(IsNull(Master1.Fields("Last Updated Date")), Master1.Fields("Last Updated Date"), Master1.Fields("Last Updated Date")), 10))
                    !Party_Code = mPartyCode
                    !Srl_No = mSrl
                    !L_C = mLocal
                    !Remark = XNull(Master1.Fields("Challan #"))
                    !Part_No = VNull(Master1.Fields("part #"))
                    !godown = GCn.Execute("Select SprWorksGodown from Syctrl").Fields(0).Value
                    !Qty_Doc = Val(Master1!Qty)
                    !Qty_Rec = Val(Master1!Qty)
                    If mVATApplicable Then
                        !Tax_YN = 1             '' if VAT is applicable in State
                    Else
                        !Tax_YN = IIf(mLocal = "L", 0, 1)
                    End If
                    !MRP_YN = IIf(mRecordDiv = "C", 1, 0)
                    If TrnType = "Local" Then
                        !TaxAmt = 0
                        !TaxPer = 0
                        !Amount = Val(Replace(Mid(IIf(IsNull(Master1!Amount), "Rs.0", Master1!Amount), 4, 15), ",", ""))
                        !Net_Amt = Val(Replace(Mid(IIf(IsNull(Master1!Amount), "Rs.0", Master1!Amount), 4, 15), ",", ""))
                        If IsNull(Master1!Qty) Or Master1!Qty = 0 Then
                            !Rate = 0
                            !V_Rate = 0
                        Else
                            !Rate = Round(Val(Replace(Mid(IIf(IsNull(Master1!Amount), "Rs.0", Master1!Amount), 4, 15), ",", "")) / Master1!Qty, 4)
                            !V_Rate = Round(Val(Replace(Mid(IIf(IsNull(Master1!Amount), "Rs.0", Master1!Amount), 4, 15), ",", "")) / Master1!Qty, 4)
                        End If
                    Else
                        '' Goods Value
                        !Amount = Val(Replace(Mid(IIf(IsNull(Master1.Fields("Net Amount")), "Rs.0", Master1.Fields("Net Amount")), 4, 15), ",", ""))
                        
                        '' Tax Value
                        If UCase(left(PubComp_Name, 3)) = "LMP" Then
                            If Not IsNull(Master1.Fields("Total_Tax_Amount")) Then
                                If Val(Mid(Master1.Fields("Total_Tax_Amount"), 4, Len(Master1.Fields("Total_Tax_Amount")) - 3)) > 0 Then
                                     !TaxAmt = Val(Replace(IIf(IsNull(Master1.Fields("Total_Tax_Amount")), 0, Mid(Master1.Fields("Total_Tax_Amount"), 4, Len(Master1.Fields("Total_Tax_Amount")) - 3)), ",", ""))
                                    mTax_Amt1 = mTax_Amt1 + Val(Replace(IIf(IsNull(Master1.Fields("Total_Tax_Amount")), 0, Mid(Master1.Fields("Total_Tax_Amount"), 4, Len(Master1.Fields("Total_Tax_Amount")) - 3)), ",", ""))
                                End If
                            End If
                        Else
                            If Not IsNull(Master1.Fields("LST")) Then
                                If Val(Mid(Master1.Fields("LST"), 4, Len(Master1.Fields("LST")) - 3)) > 0 Then
                                    !TaxAmt = IIf(IsNull(Master1.Fields("LST")), 0, Mid(Master1.Fields("LST"), 4, Len(Master1.Fields("LST")) - 3))
                                    mTax_Amt1 = mTax_Amt1 + IIf(IsNull(Master1.Fields("LST")), 0, Mid(Master1.Fields("LST"), 4, Len(Master1.Fields("LST")) - 3))
                                End If
                            End If
                            If Not IsNull(Master1.Fields("CST")) Then
                                If Val(Mid(Master1.Fields("CST"), 4, Len(Master1.Fields("CST")) - 3)) > 0 Then
                                    !TaxAmt = IIf(IsNull(Master1.Fields("CST")), 0, Mid(Master1.Fields("CST"), 4, Len(Master1.Fields("CST")) - 3))
                                    mTax_Amt1 = mTax_Amt1 + IIf(IsNull(Master1.Fields("CST")), 0, Mid(Master1.Fields("CST"), 4, Len(Master1.Fields("CST")) - 3))
                                End If
                            End If
                            If Not IsNull(Master1.Fields("VAT")) Then
                                If Val(Mid(Master1.Fields("VAT"), 4, Len(Master1.Fields("VAT")) - 3)) > 0 Then
                                    !TaxAmt = IIf(IsNull(Master1.Fields("VAT")), 0, Mid(Master1.Fields("VAT"), 4, Len(Master1.Fields("VAT")) - 3))
                                    mTax_Amt1 = mTax_Amt1 + IIf(IsNull(Master1.Fields("VAT")), 0, Mid(Master1.Fields("VAT"), 4, Len(Master1.Fields("VAT")) - 3))
                                End If
                            End If
                        End If
                                            
                        !Net_Amt = Val(Replace(Mid(IIf(IsNull(Master1.Fields("Net Amount")), "Rs.0", Master1.Fields("Net Amount")), 4, 15), ",", ""))
                        If mRecordDiv = "C" Then
                            !Amount = !Amount + !TaxAmt
                        End If
                        
                        
                        '' Tax Percentage
                        If !TaxAmt = 0 Then
                            !TaxPer = 0
                        Else
                            !TaxPer = Round(!TaxAmt * 100 / !Net_Amt, 1)
                        End If
                        
                        If IsNull(Master1!Qty) Or Master1!Qty = 0 Then
                            !Rate = 0
                            !V_Rate = 0
                        Else
                            !Rate = Round(!Amount / Master1!Qty, 2)
                            !V_Rate = !Rate     'Round(Val(Replace(Mid(IIf(IsNull(Master1.Fields("Net Amount")), "Rs.0", Master1.Fields("Net Amount")), 4, 15), ",", "")) / Master1!Qty, 4)
                        End If
                    End If
                    !Part_SrlNo = mSrl
                    !Disc_Per = 0
                    !Disc_Amt = 0
                    !Ord_Discper = 0
                    !Ord_DiscAmt = 0
                    
        
                    !U_Name = "Siebel"
                    !U_EntDt = Format(PubLoginDate, "Short Date")
                    !U_AE = "A"
                    .Update
                    TranFlag1 = False
                End With
                mQty = mQty + Master1!Qty
                mCount = mCount + 1
                If TrnType = "Local" Then
                    mAmount = mAmount + Val(Replace(Mid(IIf(IsNull(Master1!Amount), "Rs.0", Master1!Amount), 4, 15), ",", ""))
                Else
                    mAmount = mAmount + Val(Replace(Mid(IIf(IsNull(Master1.Fields("Net Amount")), "Rs.0", Master1.Fields("Net Amount")), 4, 15), ",", ""))
                End If
                mSrl = mSrl + 1
LineFileNextrecord:
                Master1.MoveNext
                CopyCnt = CopyCnt + 1
                lblRecCopy(Index).Caption = CopyCnt
                lblRecCopy(Index).Refresh
                If Master1.EOF = True Then Exit Do
            Loop
        End If
        If Master1.EOF = True Then Master1.MovePrevious
        With RsNew
            !Tot_No_Of_Items = mCount
            !Tot_Doc_Qty = mQty
            !Tot_Phy_Qty = mQty
            !TOT_Amt = mAmount
            !Tot_Disc_Amt = 0
            !Tot_Ord_DiscAmt = 0
            !Tot_Goods_value = mAmount
            !Tax_Amt = mTax_Amt1
            !Addition = 0
            !Deduction = 0
            !Net_Amt = mTax_Amt1 + mAmount
            .Update
            TranFalg = False
        End With
        If Master1.AbsolutePosition = Master1.RecordCount Then Master1.MoveNext
        CodeCnt = CodeCnt + 1
        GoTo NextLoop
        
DuplicateSkipped:

MyNextRecord:
        mSrl = 0
        If Master1.EOF = False And Master1.BOF = False Then
            Do While (Master1.EOF = False And Master1.BOF = False) And IIf(TrnType = "Local", mOrderNo = Master1.Fields("Order #") And mChallanNo = XNull(Master1.Fields("Challan #")), mChallanNo = XNull(Master1.Fields("Invoice #")))
                mSrl = mSrl + 1
                Master1.MoveNext
                CopyCnt = CopyCnt + 1
                lblRecCopy(Index).Caption = CopyCnt
                lblRecCopy(Index).Refresh
                If Master1.EOF = True Then Exit Do
            Loop
        End If
        CopyCnt = CopyCnt + mSrl - 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
NextLoop:
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    mSrl = 0
    Do Until IIf(TrnType = "Local", mOrderNo = Master1.Fields("Order #") And mChallanNo = Master1.Fields("Challan #"), mChallanNo = Master1.Fields("Invoice #")) And Master1.EOF
        mSrl = mSrl + 1
        Master1.MoveNext
    Loop
    
    ErrorCnt = ErrorCnt + mSrl
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master1.AbsolutePosition & ",'" & "" & "','Spare Challan','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub

Private Sub UnitMasterDataUpdate(Index)
    '' On Error GoTo Eloop

    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Unit", GCn, adOpenDynamic, adLockOptimistic
    
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("UoM"))) Or StringPass(Master.Fields("UoM")) = "" Then GoTo MyNextRecord
        
        If GCn.Execute("Select Unit_Name from Unit where Unit_Name='" & StringPass(Master.Fields("UoM")) & "'").RecordCount > 0 Then
            GoTo DuplicateSkipped
        End If
        
        With RsNew
            .AddNew
            !Unit_Name = left(Master.Fields("UoM"), 6)
            !Site_Code = PubSiteCode
            !U_Name = "Siebel"
            !U_EntDt = Format(PubLoginDate, "Short Date")
            !U_AE = "A"
            
            .Update
        End With

DuplicateSkipped:
        CodeCnt = CodeCnt + 1
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
MyNextRecord:
        Master.MoveNext
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Unit Master','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub

Private Sub PartMasterDataUpdate(Index)
'' On Error GoTo Eloop

    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Part", GCn, adOpenDynamic, adLockOptimistic
    
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Part Number"))) Or StringPass(Master.Fields("Part Number")) = "" Then GoTo MyNextRecord
        
        If GCn.Execute("Select Part_No from Part where Part_No='" & StringPass(Master.Fields("Part Number")) & "' and Div_Code='" & PubDivCode & "'").RecordCount > 0 Then
            GoTo DuplicateSkipped
        End If
        
        If IsNull(StringPass(Master.Fields("Description"))) Or StringPass(Master.Fields("Description")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Part Number"), "Part Master", "Part Name is Empty")
            GoTo MyNextRecord
        End If
        
        If IsNull(StringPass(Master.Fields("UoM"))) Or StringPass(Master.Fields("UoM")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Part Number"), "Part Master", "Unit is Empty")
            GoTo MyNextRecord
        End If
        
        If IsNull(StringPass(Master.Fields("Vendor"))) Or StringPass(Master.Fields("Vendor")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Part Number"), "Part Master", "Vendor Name is Empty")
            GoTo MyNextRecord
        End If
        
        With RsNew
            .AddNew
            !Part_No = left(Master.Fields("Part Number"), 22)
            !Part_NoHelp = Replace(left(Master.Fields("Part Number"), 22), " ", "")
            !Site_Code = PubSiteCode
            !Div_Code = PubDivCode
            !Part_Name = left(Master.Fields("Description"), 40)
            !Local_Name = left(Master.Fields("Description"), 40)
            !Part_NameHelp = Replace(left(Master.Fields("Description"), 40), " ", "")
            !Unit = left(Master.Fields("UoM"), 6)
            !Mark_Yn = "N"
            !Part_OEM = Master.Fields("Vendor")
            !Supl_Loca = Master.Fields("Vendor Location")
            !Value_Method = "FIFO"
            !Active_YN = 1
            !Security_Grade = "A"
            !Lead_Time = Val(StringPass(Master.Fields("Lead Time")))
            If txtDiv.Text = "C" Then
                !Disc_Factor = StringPass(Master.Fields("Discount Code (CVBU)"))
            Else
                !Disc_Factor = StringPass(Master.Fields("Discount Code"))
            End If
            !Bin_Loca = ""
            !Min_Lvl = 0
            !Max_lvl = 0
            !ReOrd_lvl = 0
            Select Case UCase(StringPass(Master.Fields("Product Category")))
                Case UCase("Lubricant")
                    !Part_Grade = "L"
                Case Else
                    !Part_Grade = "S"
            End Select
            
            !U_Name = "Siebel"
            !U_EntDt = Format(PubLoginDate, "Short Date")
            !U_AE = "A"

            .Update
        End With

DuplicateSkipped:
        CodeCnt = CodeCnt + 1
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
MyNextRecord:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Part Master','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub
Private Sub MoneyReceiptDataUpdate(Index)
'' On Error GoTo Eloop
Dim MasterCode As String, DocId As String, mPartyCode As String, mV_Type As String
Dim mPrefix As String, mOrdDocID As String, mSiebelRectType, mDebitAc As String, mOrdSite As String
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String, mDrCr As String, mSiebelCode As String
Dim mVehicleRect As Boolean
    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Rect", GCn, adOpenDynamic, adLockOptimistic
    
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Receipt No"))) Or StringPass(Master.Fields("Receipt No")) = "" Then GoTo MyNextRecord
        
        If GCn.Execute("Select SiebelRectNo from Rect where SiebelRectNo='" & StringPass(Master.Fields("Receipt No")) & "'").RecordCount > 0 Then
            GoTo DuplicateSkipped
        End If
        
        If IsNull(StringPass(Master.Fields("Division"))) Or StringPass(Master.Fields("Division")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Money Receipt (Sales)", "Division name is Empty")
            GoTo MyNextRecord
        Else
            If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").RecordCount > 0 Then
                mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").Fields(0).Value
                mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").Fields(0).Value
                mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").Fields(0).Value
                mSiebelRectType = Trim(ErrorGCN.Execute("Select RectTypeSiebelSales from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").Fields(0).Value)
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master!division), "Money Receipt (Sales)", "Automan Site/Division is not Defined in SiteDivision Table for this Siebel Division")
                GoTo MyNextRecord
            End If
        End If
        
        If Mid(Master!division, 9, 5) = "Sales" Then
            mVehicleRect = True
            If left(StringPass(Master.Fields("Receipt No")), Len(mSiebelRectType)) = mSiebelRectType Then      '' Receipt is not for Workshop Jobcard/OTC Sale
                If GCn.Execute("Select Ord_No from Veh_Order where SiebelOrderNo='" & StringPass(Master.Fields("Order_No")) & "'").RecordCount = 0 Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master.Fields("Order_No")), "Money Receipt (Sales)", "Booking/Order Not found in Automan")
                    GoTo MyNextRecord
                End If
            End If
        Else
            mVehicleRect = False
            'GoTo DuplicateSkipped       '' Record skipped if Receipt is made for Jobcard
        End If
        
        If IsNull(StringPass(Master.Fields("Receipt_Date"))) Or StringPass(Master.Fields("Receipt_Date")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Money Receipt (Sales)", "Receipt Date is Empty")
            GoTo MyNextRecord
        End If
        
        If IsNull(StringPass(Master.Fields("Payment_Method"))) Or StringPass(Master.Fields("Payment_Method")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Money Receipt (Sales)", "Payment Method is Empty")
            GoTo MyNextRecord
        End If
        
        If IsNull(StringPass(Master.Fields("Amount"))) Or StringPass(Master.Fields("Amount")) = "" Or Val(VNull(Master.Fields("Amount"))) = 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Money Receipt (Sales)", "Amount is ZERO")
            GoTo MyNextRecord
        End If
        
        mDrCr = ""
        If IsNull(StringPass(Master.Fields("Bill_Adjustment_TYPE"))) Or StringPass(Master.Fields("Bill_Adjustment_TYPE")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Money Receipt (Sales)", "Bill Adjustment Type is Empty")
            GoTo MyNextRecord
        Else
            If StringPass(Master.Fields("Bill_Adjustment_TYPE")) <> "Received" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master.Fields("Bill_Adjustment_TYPE")), "Money Receipt (Sales)", "Bill Adjustment Type should be only 'Received'")
                GoTo MyNextRecord
            Else
                mDrCr = "C"     '' here we will decide the nature for Party
            End If
        End If
        
        mDebitAc = ""
'        If IsNull(StringPass(Master.Fields("Deposited_On_Bank"))) Or StringPass(Master.Fields("Deposited_On_Bank")) = "" Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Money Receipt (Sales)", "Deposited in Bank is Empty")
'            GoTo MyNextRecord
'        Else
        If UCase(Master!Payment_Method) = UCase("Cash") Then
            mDebitAc = ErrorGCN.Execute("select CashAccountCode from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").Fields(0).Value
        Else
            If ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Deposited_On_Bank")) & "' and Type='Money Receipt'").RecordCount > 0 Then
                mDebitAc = ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Deposited_On_Bank")) & "' and Type='Money Receipt'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Deposited_On_Bank"), "Money Receipt (Sales)", "This Deposited in Bank Account Code is not defined in AccountConversion Table")
                GoTo MyNextRecord
            End If
        End If
                
'        End If
        
        
        mPartyCode = ""
        If PubDivCode = "C" Then
            mSiebelCode = Master!Account_Code
        Else
            mSiebelCode = XNull(Master!Customer_Code)
        End If
        
        If StringPass(mSiebelCode) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Money Receipt (Sales)", "Account Code (for CVD)/CustomerCode (for PCD) is Empty")
            GoTo MyNextRecord
        End If
        
        If GCn.Execute("Select SubCode from SubGroup where SiebelCode='" & StringPass(mSiebelCode) & "'").RecordCount = 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(mSiebelCode), "Money Receipt (Sales)", "Account Code not found in Ledger Account Master of Automan")
            GoTo MyNextRecord
        Else
            mPartyCode = GCn.Execute("Select SubCode from SubGroup where siebelCode='" & StringPass(mSiebelCode) & "'").Fields(0).Value
        End If
        Dim mShortYear As String
        If Month(Master.Fields("Receipt_Date")) > 3 Then
            mShortYear = Right(Format(Master.Fields("Receipt_Date"), "yy"), 1) & Right(Val(Format(Master.Fields("Receipt_Date"), "yy")) + 1, 1)
        Else
            mShortYear = Right(Val(Format(Master.Fields("Receipt_Date"), "yy")) - 1, 1) & Right(Format(Master.Fields("Receipt_Date"), "yy"), 1)
        End If
        If mVehicleRect Then
            mPrefix = "SBL" & mShortYear 'Format(Master.Fields("Receipt_Date"), "yy")
        Else
            mPrefix = "WRK" & mShortYear 'Format(Master.Fields("Receipt_Date"), "yy")
        End If
        Select Case UCase(Master!Payment_Method)
            Case UCase("Cash")
                mV_Type = ErrorGCN.Execute("Select RectType_Cash From Enviro").Fields(0).Value
            Case UCase("Cheque"), UCase("Demand Draft")
                mV_Type = ErrorGCN.Execute("Select RectType_Cheque From Enviro").Fields(0).Value
            Case UCase("Release Order")
                mV_Type = ErrorGCN.Execute("Select RectType_ReleaseOrder From Enviro").Fields(0).Value
            Case Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master!Payment_Method, "Money Receipt (Sales)", "Payment Method is New (This Receipt Type is not recognized to Import Data in Automan)")
                GoTo MyNextRecord
        End Select
        CodeCnt = Right("00000000" & Right(Master.Fields("Receipt No"), 6), 8)
        
        DocId = mRecordDiv & mRecordSite & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
        If mVehicleRect = True Then
            mOrdDocID = GCn.Execute("Select OrdDocID from Veh_Order where SiebelOrderNo='" & Master!Order_No & "'").Fields(0).Value
            mOrdSite = GCn.Execute("Select Ord_SiteCode from Veh_Order where SiebelOrderNo='" & Master!Order_No & "'").Fields(0).Value
        Else
            mOrdDocID = ""
            mOrdSite = ""
        End If
        'Checking for Blank Invoice No. for currend Order No. (because Single Order of Sieble can be for Multiple Invoice)
        With RsNew
            .AddNew
            !DocId = DocId
            !DocIDHelp = Replace(DocId, " ", "")
            !Site_Code = mRecordSite & mRecordSite
            !V_Type = mV_Type
            !V_No = CodeCnt
            !V_DATE = MakeDate(Master!Receipt_Date)
            !Prov_Date = IIf(IsNull(Master!Instr_Date), Master!Instr_Date, MakeDate(XNull(Master!Instr_Date)))
            !PartyCode = mPartyCode
            !Amount = VNull(Master!Amount)
            !DrCr = mDrCr
            !Narration = left(Trim(StringPass(Master!CHq_DD_RO_No) & " " & StringPass(Master!Drawn_on_Bank) & " " & StringPass(Master!Branch)), 100)
            !Narration1 = Mid(Trim(StringPass(Master!CHq_DD_RO_No) & " " & StringPass(Master!Drawn_on_Bank) & " " & StringPass(Master!Branch)), 101, 40)
            !AcCode = mDebitAc
            !DDNo = left(Trim(StringPass(Master!CHq_DD_RO_No)), 10)
            !PrintParty_YN = 0
            !Printed = 0
            !IForm = 0
            
            
            ''' Conditional Values
            !Vehicle_YN = mVehicleRect
            !Ord_SiteCode = mOrdSite
            !Ord_DocId = mOrdDocID
            !SiebelRectNo = Master.Fields("Receipt No")
            
            !U_Name = "Siebel"
            !U_EntDt = Format(PubLoginDate, "Short Date")
            !U_AE = "A"

            
            .Update
        End With
        If mVehicleRect Then
            If Master!Payment_Method = "Release Order" Then
                GCn.Execute ("Update Veh_Order set Fin_Amt=" & Master!Amount & " where OrdDocID='" & mOrdDocID & "'")
            End If
        End If

        'Insert Into Ledger
        GCnFa.Execute "INSERT INTO LEDGERM (DocId,V_Type,v_Prefix,V_No,Site_Code,V_Date,U_Name,U_EntDt,U_AE) values ('" & _
                      "" & DocId & "','" & mV_Type & "','" & mPrefix & "'," & CodeCnt & ",'" & mRecordSite & mRecordSite & "'," & ConvertDate(MakeDate(Master!Receipt_Date)) & ",'SA'," & ConvertDate(PubLoginDate) & ",'A') "
        Select Case mV_Type
            Case "SBLCS", "SBLRO", "SBLCQ"   'Receipt
                GCnFa.Execute "INSERT INTO LEDGER (DocId,V_Type,V_No,Site_Code,V_Date," & _
                    "V_SNo,Narration,SubCode,ContraSub,AmtCr,AmtDr, " & _
                    "U_Name,U_EntDt,U_AE) " & _
                    "values ('" & DocId & "','" & mV_Type & "'," & CodeCnt & ",'" & mRecordSite & mRecordSite & "'," & ConvertDate(MakeDate(Master!Receipt_Date)) & _
                    ",1,'" & Trim(StringPass(Master!CHq_DD_RO_No) & " " & StringPass(Master!Drawn_on_Bank) & " " & StringPass(Master!Branch)) & _
                    "','" & mPartyCode & "', '" & mDebitAc & "'," & Val(Master!Amount) & ",0 " & _
                    " ,'SA'," & ConvertDate(PubLoginDate) & ",'A') "
                
                GCnFa.Execute "INSERT INTO LEDGER (DocId,V_Type,V_No,Site_Code,V_Date," & _
                    "V_SNo,Narration,SubCode,ContraSub,AmtCr,AmtDr, " & _
                    "U_Name,U_EntDt,U_AE) " & _
                    "values ('" & DocId & "','" & mV_Type & "'," & CodeCnt & ",'" & mRecordSite & mRecordSite & "'," & ConvertDate(MakeDate(Master!Receipt_Date)) & _
                    ",2,'" & Trim(StringPass(Master!CHq_DD_RO_No) & " " & StringPass(Master!Drawn_on_Bank) & " " & StringPass(Master!Branch)) & _
                    "','" & mDebitAc & "', '" & mPartyCode & "',0," & Val(Master!Amount) & _
                    " ,'SA'," & ConvertDate(PubLoginDate) & ",'A') "
        End Select

DuplicateSkipped:
        CodeCnt = CodeCnt + 1
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
MyNextRecord:
        Master.MoveNext
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Money Receipt (Sales)','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub
Private Sub VehicleSalesDataUpdate(Index)
'' On Error GoTo Eloop
Dim MasterCode As String, DocId As String, mPartyCode As String, mV_Type As String
Dim mNetAmt As Double, mRoundAmt As Double, mVatTax As Double, mSalePrice As Double, mVRate As Double, mRate As Double
Dim mForm_Code As String, mVatPer As Double, mRTOName As String, mOrdDocID As String
Dim mRepCode As String, mColCode As String, mPrefix As String, ChalDocId As String
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String, mFinCode As String
    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Veh_Order order by SiebelOrderNo,OrdDocID", GCn, adOpenDynamic, adLockOptimistic
    
    Set rsTemp = GCn.Execute("Select * From ContractFinance where FinCatg=0")
    
    mV_Type = "V_SB"
    
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Invoice_No"))) Or StringPass(Master.Fields("Invoice_No")) = "" Then
            GoTo MyNextRecord
        End If
        
        If GCn.Execute("Select Ord_No from Veh_Order where SiebelInvoiceNo='" & StringPass(Master.Fields("Invoice_No")) & "'").RecordCount > 0 Then
            GoTo DuplicateSkipped
        End If
        
        If IsNull(StringPass(Master.Fields("Order_No"))) Or StringPass(Master.Fields("Order_No")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Order No. is Empty")
            GoTo MyNextRecord
        End If
        
        If GCn.Execute("Select Ord_No from Veh_Order where SiebelOrderNo='" & StringPass(Master.Fields("Order_No")) & "'").RecordCount = 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Booking/Order Not found in Automan")
            GoTo MyNextRecord
        End If
        
        If GCn.Execute("Select ChassisNo from Veh_stock where ChassisNo='" & Master!Chassis_No & "'").RecordCount = 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master!Chassis_No, "Vehicle Sales", "Chassis Not found In Purchase")
            GoTo MyNextRecord
        Else
            If GCn.Execute("Select ChassisNo from Veh_stock where (Sal_DocId='' or Sal_DocID is Null) and ChassisNo='" & Master!Chassis_No & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master!Chassis_No, "Vehicle Sales", "Chassis No is already sold")
                GoTo MyNextRecord
            End If
        End If
        
        If IsNull(StringPass(Master.Fields("Invoice_Date"))) Or StringPass(Master.Fields("Invoice_Date")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Invoice Date is Empty")
            GoTo MyNextRecord
        End If
                
        If IsNull(StringPass(Master.Fields("VC_No"))) Or StringPass(Master.Fields("VC_No")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Product/Model is Empty")
            GoTo MyNextRecord
        Else
            If GCn.Execute("Select Model from Veh_Order where SiebelOrderNo='" & StringPass(Master!Order_No) & "'").Fields(0).Value <> StringPass(Master.Fields("VC_No")) Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master!VC_No, "Vehicle Sales", "Siebel Sales VC code is mismatch with Automan Order VC Code")
                GoTo MyNextRecord
            End If
        End If
        
        If PubDivCode = "C" Then
            If IsNull(StringPass(Master.Fields("Account_Code"))) Or StringPass(Master.Fields("Account_Code")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Party Account Code is Empty")
                GoTo MyNextRecord
            End If
        Else
            If IsNull(StringPass(Master.Fields("Customer_Code"))) Or StringPass(Master.Fields("Customer_Code")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Customer Account Code is Empty")
                GoTo MyNextRecord
            End If
        End If
        
        If IsNull(StringPass(Master.Fields("Chassis_No"))) Or StringPass(Master.Fields("Chassis_No")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Chassis No. is Empty")
            GoTo MyNextRecord
        End If
        
        If IsNull(StringPass(Master.Fields("SALES_Ledger_name"))) Or StringPass(Master.Fields("SALES_Ledger_name")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Sales Ledger name is Empty")
            GoTo MyNextRecord
        End If
        
        If IsNull(StringPass(Master.Fields("Godown"))) Or StringPass(Master.Fields("Godown")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Godown name is Empty")
            GoTo MyNextRecord
        End If
        
        If IsNull(StringPass(Master.Fields("Division"))) Or StringPass(Master.Fields("Division")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Division name is Empty")
            GoTo MyNextRecord
        Else
            If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").RecordCount > 0 Then
                mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").Fields(0).Value
                mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").Fields(0).Value
                mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master!division), "Vehicle Sales", "Automan Site/Division is not Defined in SiteDivision Table for this Siebel Division")
                GoTo MyNextRecord
            End If
        End If
        
        If IsNull(StringPass(Master.Fields("SalesPerson_Name"))) Or StringPass(Master.Fields("SalesPerson_Name")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Sales Person name is Empty")
            GoTo MyNextRecord
        End If
        
        mPartyCode = GCn.Execute("Select Partycode from veh_order where siebelOrderNo='" & StringPass(Master.Fields("Order_No")) & "'").Fields(0).Value
        If GCn.Execute("Select SiebelCode from Subgroup where SubCode='" & mPartyCode & "'").RecordCount = 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Sales", "Siebel PartyCode not found in Ledger Account Master")
            GoTo MyNextRecord
        Else
            If GCn.Execute("Select SiebelCode from Subgroup where SubCode='" & mPartyCode & "'").Fields(0).Value <> StringPass(IIf(PubDivCode = "C", Master!Account_Code, Master!Customer_Code)) Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "Sales Party : " & IIf(PubDivCode = "C", Master!Account_Code, Master!Customer_Code) & " Order Party : " & mPartyCode, "Vehicle Sales", "Siebel Sales Party code is mismatch with Automan Order Party")
                GoTo MyNextRecord
            End If
        End If
'        mRepCode = ""
'        If GCn.Execute("Select Emp_Code from Emp_Mast where Reference='" & Master!SalesPerson_Name & "'").RecordCount = 0 Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master!SalesPerson_Name, "Vehicle Sales", "Sales Rep Code not found in Automan")
'            GoTo MyNextRecord
'        Else
'            mRepCode = GCn.Execute("Select Emp_Code from Emp_Mast where Reference='" & Master!SalesPerson_Name & "'").Fields(0).Value
'        End If

''        '' Searching for Bank/Financer Name Code
''        mFinCode = ""
''        If StringPass(Master.Fields("Financed By")) = "" Then
''            mFinCode = GCn.Execute("Select FinCode from ContractFinance where UnderFinGrp='CASH'").Fields(0).Value
''            mFinAcCode = GCn.Execute("Select iif(isnull(AcCode),'',AcCode) from ContractFinance where UnderFinGrp='CASH'").Fields(0).Value
''        Else
''            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
''            Do Until rsTemp.EOF
''                If ReturnString(rsTemp!FinName) = ReturnString(left(StringPass(Master.Fields("Financed By")), 40)) Then
''                    mFinCode = rsTemp!FinCode
''                    mFinAcCode = XNull(rsTemp!AcCode)
''                    Exit Do
''                End If
''                rsTemp.MoveNext
''            Loop
''            If mFinCode = "" Then
''                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Financed By"), "Vehicle Booking", "Financer name not found in Master")
''                GoTo MyNextRecord
''            End If
''        End If

        mColCode = GCn.Execute("Select " & xIsNull("Colour_Code", "") & " from Veh_Stock where ChassisNo='" & Master!Chassis_No & "'").Fields(0).Value
        mPrefix = "SBL" & Format(Master.Fields("Invoice_Date"), "yy")
        CodeCnt = Right("00000000" & Right(Master!Invoice_No, 5), 8)
        
        DocId = mRecordDiv & mRecordSite & mRecordSite & " " & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
        ChalDocId = mRecordDiv & mRecordSite & mRecordSite & "V_DCL" & mPrefix & Right("00000000" & CodeCnt, 8)
        
        mForm_Code = ErrorGCN.Execute("Select VehicleSaleFormCode from Enviro").Fields(0).Value
        mVatPer = ErrorGCN.Execute("Select VehicleSaleTaxPer from Enviro").Fields(0).Value
        mRTOName = GCn.Execute("Select City.CityName from SubGroup Left Join City on SubGroup.CityCode=City.CityCode where SubCode='" & mPartyCode & "'").Fields(0).Value
       
        mRate = GCn.Execute("Select Rate from Veh_Stock where ChassisNo='" & Master!Chassis_No & "' and (Sal_DocId='' or Sal_DocID is Null)").Fields(0).Value
        mVRate = GCn.Execute("Select VRate from Veh_Stock where ChassisNo='" & Master!Chassis_No & "' and (Sal_DocId='' or Sal_DocID is Null)").Fields(0).Value
        
        mNetAmt = Round(VNull(Master.Fields("Exshowroom_Price")), 2)
        
        'If IsNull(Master!VatTax) Or Master!VatTax = 0 Then
            mVatTax = Round(mNetAmt * mVatPer / (100 + mVatPer), 2)
            mSalePrice = mNetAmt - mVatTax
'        Else
'            mVatTax = Master!VatTax
'            If IsNull(Master.Fields("VAT Assessible Amt 1")) Or Master.Fields("VAT Assessible Amt 1") = 0 Then
'                mSalePrice = mNetAmt - mVatTax
'            Else
'                mSalePrice = Master.Fields("VAT Assessible Amt 1")
'            End If
'        End If
        mRoundAmt = Round(mNetAmt, 0) - mNetAmt
        mNetAmt = Round(mNetAmt, 0)
        
        'Checking for Blank Invoice No. for currend Order No. (because Single Order of Sieble can be for Multiple Invoice)
        If RsNew.RecordCount > 0 Then RsNew.MoveFirst
        RsNew.Find ("SiebelOrderNo='" & Master!Order_No & "'")
        If RsNew.EOF = True Or RsNew.BOF = True Then MsgBox " Order No. not found to updated sales information": GoTo MyNextRecord
        While True
            If RsNew.EOF = True Or RsNew!siebelOrderNo <> Master!Order_No Then GoTo DuplicateSkipped
            If IsNull(RsNew!SiebelInvoiceNo) Or RsNew!SiebelInvoiceNo = "" Then GoTo UpdateSales Else RsNew.MoveNext
        Wend
UpdateSales:
        mOrdDocID = RsNew!OrdDocId
        With RsNew
            !Inv_DocID = DocId
            !Inv_DocIDHelp = Replace(DocId, " ", "")
            !Inv_SiteCode = mRecordSite & mRecordSite
            !Inv_VType = mV_Type
            !Inv_No = CodeCnt
            !Inv_Date = Format(Master!Invoice_Date, "dd/MMM/yyyy")
            
            !DelCh_DocID = ChalDocId
            !DelCh_DocIDHelp = Replace(ChalDocId, " ", "")
            !DelCh_SiteCode = mRecordSite & mRecordSite
            !DelCh_VType = "V_DCL"
            !DelCh_No = CodeCnt
            !DelCh_Dt = Format(Master!Invoice_Date, "dd/MMM/yyyy")
            
            !DelChPrn_YN = 0
            !DelCh_UName = "Siebel"
            !DelCh_UEntDt = Format(PubLoginDate, "dd/MMM/yyyy")
            !DelCh_UAE = "A"
            
            !Inv_UName = "Siebel"
            !Inv_UEntDt = Format(PubLoginDate, "dd/MMM/yyyy")
            !Inv_UAE = "A"
            
            !TrnType_Prn = 0
            !RoundOff_YN = 1
            !Interest_YN = 0
            !TDS_YN = 0
            !Certi = CodeCnt
            !CertiPrn_YN = 0
            !TCertiPrn_YN = 0
            !BillPrn_YN = 0
            !Inv_Prefix = mPrefix
            !Chas_Type = left(Master!Chassis_No, 6)
            !Chassis = Master!Chassis_No
            !Colour_Code = mColCode
            !Form_Code = mForm_Code
            !RTO = mRTOName
            
            !Rate = mRate
            !VRate = mRate   '' mVRate
            !Margine = mSalePrice - mRate
            !Subtot = mSalePrice
            !Rebate = 0
            !InciChrg = 0
            !Octroi = 0
            !RegTemp = 0
            !TransitInsu = 0
            !Transport = 0
            !MVT = 0
            !Ins_Fee = VNull(Master.Fields("Total_Order_Value")) - VNull(Master.Fields("Exshowroom_Price"))
            !Tax_Per = mVatPer
            !Tax_Amt = mVatTax
            !OtherChrg = 0
            !Round_off = mRoundAmt
            !Net_Amount = mNetAmt
            !SiebelInvoiceNo = Master!Invoice_No
            
            .Update
        End With
        
        GCn.Execute ("Update Veh_Stock set Sal_DocID='" & DocId & "'," & _
                                        " Sal_DocIDHelp='" & Replace(DocId, " ", "") & "'," & _
                                        " Sal_Site_Code='" & mRecordSite & mRecordSite & "'," & _
                                        " Sal_VType='" & mV_Type & "'," & _
                                        " Sal_VNo=" & CodeCnt & "," & _
                                        " Sal_VDate=" & ConvertDate(Master!Invoice_Date) & "," & _
                                        " Sal_Rate=" & mNetAmt & "," & _
                                        " Ord_SiteCode='" & mRecordSite & mRecordSite & "'," & _
                                        " Ord_DocID='" & mOrdDocID & "'," & _
                                        " DelCh_DocID='" & ChalDocId & "'," & _
                                        " DelCh_Date=" & ConvertDate(Master!Invoice_Date) & "," & _
                                        " RTO_Name='" & mRTOName & "' where ChassisNo='" & Master!Chassis_No & "' and (Sal_DocID='' or Sal_DocID is Null)")

DuplicateSkipped:
        CodeCnt = CodeCnt + 1
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
MyNextRecord:
        Master.MoveNext
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Vehicle Sales','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub

Private Sub VehicleBookingDataUpdate(Index, DivisionType)
'' On Error GoTo Eloop
Dim MasterCode As String, DocId As String, mPartyCode As String, mLength As Integer, mV_Type As String
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String, mCityCode As String, mOrderQty As Integer
Dim mFinAcCode As String, mFinCode As String, mPrefix As String, mname As String
Dim mDmsSubCode As String
    
    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Veh_Order", GCn, adOpenDynamic, adLockOptimistic
    
    Set rsTemp = GCn.Execute("Select * From ContractFinance where FinCatg=0")
    
    
    
    mV_Type = "V_BK"
    Do Until Master.EOF
    
    
'        mDmsSubCode = IIf(XNull(!Account_Code) = "", XNull(!Customer_Code), XNull(!Account_Code))
'        mSubCode = AutomanSubcode(mDmsSubCode, RsDmsEnviro!SprDebtorGroupCode, "Customer")
    
        If IsNull(StringPass(Master.Fields("Order #"))) Or StringPass(Master.Fields("Order #")) = "" Then GoTo MyNextRecord
        
        If GCn.Execute("Select Ord_No from Veh_Order where SiebelOrderNo='" & StringPass(Master.Fields("Order #")) & "'").RecordCount > 0 Then
            
            GoTo DuplicateSkipped
        End If
                
        If IsNull(StringPass(Master.Fields("Order Date"))) Or StringPass(Master.Fields("Order Date")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Booking", "Order Date is Empty")
            GoTo MyNextRecord
        End If
        If DivisionType = "CVD" Then
            If IsNull(StringPass(Master.Fields("Product"))) Or StringPass(Master.Fields("Product")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Booking", "Product/Model is Empty")
                GoTo MyNextRecord
            End If
        Else
            If IsNull(StringPass(Master.Fields("Product/VC#"))) Or StringPass(Master.Fields("Product/VC#")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Booking", "Product/Model is Empty")
                GoTo MyNextRecord
            End If
        End If
        If IsNull(StringPass(Master.Fields("Quantity Requested"))) Or StringPass(Master.Fields("Quantity Requested")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Booking", "Quantity Requested field is Empty")
            GoTo MyNextRecord
        End If
        
        If DivisionType = "CVD" Then
            If IsNull(StringPass(Master.Fields("Account Street Address1"))) Or StringPass(Master.Fields("Account Street Address1")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Booking", "Account Street Address1 field is Empty")
                GoTo MyNextRecord
            End If
            
            mCityCode = ""
            If IsNull(StringPass(Master.Fields("Account City"))) Or StringPass(Master.Fields("Account City")) = "" Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Booking", "Account City field is Empty")
                GoTo MyNextRecord
            Else
                If GCn.Execute("Select CityName from City where CityHelp='" & left(Replace(Trim(StringPass(Master.Fields("Account City"))), " ", ""), 25) & "'").RecordCount > 0 Then
                    mCityCode = GCn.Execute("Select CityCode from City where CityHelp='" & left(Replace(Trim(StringPass(Master.Fields("Account City"))), " ", ""), 25) & "'").Fields(0).Value
                Else
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master.Fields("Account CIty")), "Vehicle Booking", "City Code not found in City Master")
                    GoTo MyNextRecord
                End If
            End If
            mLength = Len(left(Trim(StringPass(Master!Account)), 40))
            If GCn.Execute("Select Name from Subgroup where left(name, " & mLength & ")='" & left(Trim(StringPass(Master!Account)), 40) & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master.Fields("Account")), "Vehicle Booking", "Party Name not found in Ledger Account Master")
                GoTo MyNextRecord
            Else
                'If GCn.Execute("Select SubCode from Subgroup where left(name, " & mLength & ")='" & left(Trim(StringPass(Master!Account)), 40) & "' and trim(Add1)='" & left(Trim(StringPass(Master.Fields("Account Street Address1"))), 40) & "' and (trim(Add2)='" & left(Trim(StringPass(Master.Fields("Account Street Address2"))), 40) & "' or add2 is null) and CityCode='" & mCityCode & "'").RecordCount = 0 Then
                If GCn.Execute("Select SubCode from Subgroup where left(name, " & mLength & ")='" & left(Trim(StringPass(Master!Account)), 40) & "' and " & cTrim("Add1") & "='" & left(Trim(StringPass(Master.Fields("Account Street Address1"))), 40) & "' and CityCode='" & mCityCode & "'").RecordCount = 0 Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Account"), "Vehicle Booking", "Ledger A/c not found in Ledger Account Master")
                    GoTo MyNextRecord
                Else
                    mPartyCode = GCn.Execute("Select SubCode from Subgroup where left(name, " & mLength & ")='" & left(Trim(StringPass(Master!Account)), 40) & "' and " & cTrim("Add1") & "='" & left(Trim(StringPass(Master.Fields("Account Street Address1"))), 40) & "'  and CityCode='" & mCityCode & "'").Fields(0).Value
                End If
            End If
        Else
            mCityCode = ""
            mPartyCode = ""
            mname = Trim(StringPass(Master.Fields("Contact Name")))
            If left(Master.Fields("Contact Name"), 4) = "Mr. " Then
                mname = Mid(Trim(StringPass(Master.Fields("Contact Name"))), 5, 40)
            ElseIf left(Master.Fields("Contact Name"), 4) = "Dr. " Then
                mname = Mid(Trim(StringPass(Master.Fields("Contact Name"))), 5, 40)
            ElseIf left(Master.Fields("Contact Name"), 6) = "Miss. " Then
                mname = Mid(Trim(StringPass(Master.Fields("Contact Name"))), 7, 40)
            ElseIf left(Master.Fields("Contact Name"), 5) = "Mrs. " Then
                mname = Mid(Trim(StringPass(Master.Fields("Contact Name"))), 6, 40)
            End If
            mLength = Len(left(Trim(mname), 40))
            
            If GCn.Execute("Select Name from Subgroup where left(name, " & mLength & ")='" & left(Trim(mname), 40) & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, left(Trim(mname), 40), "Vehicle Booking", "Party Name not found in Ledger Account Master")
                GoTo MyNextRecord
            Else
                If GCn.Execute("Select SubCode from Subgroup where left(name, " & mLength & ")='" & left(Trim(mname), 40) & "'").RecordCount = 0 Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, left(Trim(mname), 40), "Vehicle Booking", "Ledger A/c not found in Ledger Account Master")
                    GoTo MyNextRecord
                Else
                    mPartyCode = GCn.Execute("Select SubCode from Subgroup where left(name, " & mLength & ")='" & left(Trim(mname), 40) & "'").Fields(0).Value
                End If
            End If
            
        End If
        
        '' Searching for Bank/Financer Name Code
        If DivisionType = "CVD" Then
            mFinCode = ""
            If StringPass(Master.Fields("Financed By")) = "" Then
                mFinCode = GCn.Execute("Select FinCode from ContractFinance where UnderFinGrp='CASH'").Fields(0).Value
                mFinAcCode = GCn.Execute("Select iif(isnull(AcCode),'',AcCode) from ContractFinance where UnderFinGrp='CASH'").Fields(0).Value
            Else
                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                Do Until rsTemp.EOF
                    If ReturnString(rsTemp!FinName) = ReturnString(left(StringPass(Master.Fields("Financed By")), 40)) Then
                        mFinCode = rsTemp!FinCode
                        mFinAcCode = XNull(rsTemp!AcCode)
                        Exit Do
                    End If
                    rsTemp.MoveNext
                Loop
                If mFinCode = "" Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Financed By"), "Vehicle Booking", "Financer name not found in Master")
                    GoTo MyNextRecord
                End If
            End If
        Else
            mFinCode = ""
            If StringPass(Master.Fields("Hypothecation")) = "" Then
                mFinCode = GCn.Execute("Select FinCode from ContractFinance where UnderFinGrp='CASH'").Fields(0).Value
                mFinAcCode = GCn.Execute("Select iif(isnull(AcCode),'',AcCode) from ContractFinance where UnderFinGrp='CASH'").Fields(0).Value
            Else
                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                Do Until rsTemp.EOF
                    If ReturnString(rsTemp!FinName) = ReturnString(left(StringPass(Master.Fields("Hypothecation")), 40)) Then
                        mFinCode = rsTemp!FinCode
                        mFinAcCode = XNull(rsTemp!AcCode)
                        Exit Do
                    End If
                    rsTemp.MoveNext
                Loop
                If mFinCode = "" Then
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Hypothecation"), "Vehicle Booking", "Financer name not found in Master")
                    GoTo MyNextRecord
                End If
            End If
        End If
        mOrderQty = IIf(VNull(Master.Fields("Quantity Requested")) = 0, 1, Master.Fields("Quantity Requested"))
        Do Until mOrderQty = 0
            
            mPrefix = "SBL" & Format(Master.Fields("Order Date"), "yy")
            CodeCnt = GCn.Execute("Select " & vIsNull("Max(ord_No)", "0") & "+1 from Veh_Order where Left(OrdDocID,1)='" & PubDivCode & "' and " & cMID("OrdDocID", "2", "1") & "='" & PubSiteCode & "'").Fields(0).Value
            DocId = PubDivCode & PubSiteCode & PubSiteCode & " " & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
           
            'Insert New Rec
            With RsNew
                .AddNew
                !OrdDocId = DocId
                !OrdDocIDHelp = Replace(DocId, " ", "")
                !Ord_SiteCode = PubSiteCode & PubSiteCode
                !Ord_VType = mV_Type
                !Ord_No = CodeCnt
                !Ord_Date = MakeDate(left(Master.Fields("Order Date"), 10))
                !Quot_SiteCode = PubSiteCode
                !PartyCode = mPartyCode
                If DivisionType = "CVD" Then
                    !Model = Master.Fields("Product")
                    !Fund_Source = IIf(StringPass(Master.Fields("Financed By")) = "", 2, 1)
                    !Fin_YN = IIf(StringPass(Master.Fields("Financed By")) = "", 0, 1)
                Else
                    !Model = Master.Fields("Product/VC#")
                    !Fund_Source = IIf(StringPass(Master.Fields("Hypothecation")) = "", 2, 1)
                    !Fin_YN = IIf(StringPass(Master.Fields("Hypothecation")) = "", 0, 1)
                End If
                !Qty = 1
                !Rate = 0
                !Intd_Use = ""
                !Permit_N_Z = 0     ' National Permit
                !Govt_YN = 0        ' No Govt.
                !Addveri_YN = 1
                !PermitReq_YN = 1
                !FB_Code = mFinCode
                !Fin_AcCode = mFinAcCode
                
                !siebelOrderNo = Master.Fields("Order #")
    
                !Book_UName = "Siebel"
                !Book_UEntDt = Format(PubLoginDate, "Short Date")
                !Book_UAE = "A"
                .Update
            End With
            mOrderQty = mOrderQty - 1
        Loop
DuplicateSkipped:
        CodeCnt = CodeCnt + 1
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
MyNextRecord:
        Master.MoveNext
    Loop
    GCn.CommitTrans
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Vehicle Booking','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub
Private Sub LedgerAccountDataUpdate(Index As Integer, mType As String)
'' On Error GoTo Eloop
Dim MasterCode As String, DocId As String, mGroupCode As String, mPartyCode As String
Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String
Dim mGroupNature As String, mNature As String, mCityCode As String, mCityName As String, mPartyType As String
Dim mName1 As String
    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from SubGroup", GCn, adOpenDynamic, adLockOptimistic
    
    Set rsTemp = New adodb.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open "Select * from SubGroupAlias", GCn, adOpenDynamic, adLockOptimistic
    
    Do Until Master.EOF
        
        If IsNull(StringPass(Master.Fields("Customer_Code"))) Or StringPass(Master.Fields("Customer_Code")) = "" Then GoTo MyNextRecord
        
        If GCn.Execute("Select SiebelCode from SubGroup where SiebelCode='" & StringPass(Master!Customer_Code) & "'").RecordCount > 0 Then
            GoTo DuplicateSkipped
        End If
                
        If IsNull(StringPass(Master.Fields("Customer_Name"))) Or StringPass(Master.Fields("Customer_Name")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Ledger Account", "Customer Name is Empty")
            GoTo MyNextRecord
        End If
                
        If IsNull(StringPass(Master.Fields("Group"))) Or StringPass(Master.Fields("Group")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Ledger Account", "A/c Group is Empty")
            GoTo MyNextRecord
        Else
            If mType = "Vehicle" Then
                If ErrorGCN.Execute("select AutomanGroupCode from AccountGroup where SiebelGroupName='" & StringPass(Master!Group) & "'").RecordCount > 0 Then
                    mGroupCode = ErrorGCN.Execute("select AutomanGroupCode from AccountGroup where SiebelGroupName='" & StringPass(Master!Group) & "'").Fields(0).Value
                Else
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master!Group), "Ledger Account", "Automan A/c Group Code is not Defined in AccountTable for this Siebel Group")
                    GoTo MyNextRecord
                End If
            Else
                If ErrorGCN.Execute("select AutomanGroupCodeWorkshop from AccountGroup where SiebelGroupName='" & StringPass(Master!Group) & "'").RecordCount > 0 Then
                    mGroupCode = ErrorGCN.Execute("select AutomanGroupCodeWorkShop from AccountGroup where SiebelGroupName='" & StringPass(Master!Group) & "'").Fields(0).Value
                Else
                    Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master!Group), "Ledger Account", "Automan A/c Group Code is not Defined in AccountTable for this Siebel Group")
                    GoTo MyNextRecord
                End If
            End If
        End If
        
        
        If IsNull(StringPass(Master.Fields("Division"))) Or StringPass(Master.Fields("Division")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Ledger Account", "Division is empty in Excel file")
            GoTo MyNextRecord
        Else
            If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").RecordCount > 0 Then
                mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").Fields(0).Value
                mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").Fields(0).Value
                mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master!division) & "'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master!division), "Ledger Account", "Automan Site/Division is not Defined in SiteDivision Table for this Siebel Division")
                GoTo MyNextRecord
            End If
        End If
        
        If IsNull(StringPass(Master!City)) Or StringPass(Master!City) = "" Then
            mCityCode = ""
        Else
            If GCn.Execute("Select CityName from City where CityHelp='" & Replace(Trim(StringPass(Master!City)), " ", "") & "'").RecordCount > 0 Then
                mCityCode = GCn.Execute("Select CityCode from City where CityHelp='" & Replace(Trim(StringPass(Master!City)), " ", "") & "'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master!City), "Ledger Account", "City Code not found in City Master")
                GoTo MyNextRecord
            End If
        End If
        
        mName1 = left(Trim(StringPass(Master!Customer_Name)) & " [" & Trim(StringPass(Master!Customer_Code)) & "]", 40)
        
        If GCn.Execute("Select Name from Subgroup where Name='" & mName1 & "'").RecordCount > 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, mName1, "Ledger Account", "Duplicate Ledger Name having Diff. Siebel Code")
            GoTo MyNextRecord
        End If
        
        CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal("right(SubCode,6)") & ")", "0") & "+1 from SubGroup where Left(SubCode,1)='" & mRecordSite & "' and  " & cMID("SubCode", "2", "1") & "='" & mRecordFirm & "'").Fields(0).Value
        mPartyCode = mRecordSite & mRecordFirm & Right("000000" & CodeCnt, 6)
        mGroupNature = GCnFa.Execute("Select GroupNature from AcGroup where GroupCode='" & mGroupCode & "'").Fields(0).Value
        mNature = GCnFa.Execute("Select Nature from AcGroup where GroupCode='" & mGroupCode & "'").Fields(0).Value
       
        
        'Insert New Rec
        With RsNew
            .AddNew
            !AcID = mPartyCode
            !SubCode = mPartyCode
            !Site_Code = mRecordSite
            !FirmCode = mRecordFirm
            !Name = left(Trim(StringPass(Master!Customer_Name)) & " [" & Trim(StringPass(Master!Customer_Code)) & "]", 40)
            !NameHelp = ReturnString(left(Trim(StringPass(Master!Customer_Name)) & " [" & Trim(StringPass(Master!Customer_Code)) & "]", 40))
            !GroupCode = mGroupCode
            !GroupNature = mGroupNature
            !Nature = mNature
            !AliasYN = "N"
            !ConPerson = left(Trim(StringPass(Master!First_Name) & " " & StringPass(Master!Middle_Name) & " " & StringPass(Master!Last_Name)), 40)
            !Add1 = left(Trim(StringPass(Master!Addr_L_1)), 40)
            !Add2 = left(Trim(StringPass(Master!Addr_L_2)), 40)
            !Add3 = ""
            !CityCode = mCityCode
            !Pin = left(StringPass(Master!Pin_Code), 6)
            !Phone = StringPass(Master!Phone)
            !Fax = StringPass(Master!Fax)
            !EMail = StringPass(Master!EMail)
            !ActiveYN = 1
            !Govt_YN = 0
            !CreditLimit = 0
            !CreditDays = 0
            !L_C = "L"
            !Party_Type = 0         ' For General Parties
            !SiebelCode = StringPass(Master!Customer_Code)
        
            !U_Name = "Siebel"
            !U_EntDt = Format(PubLoginDate, "Short Date")
            !U_AE = "A"
            .Update
        End With
        
        With rsTemp
            .AddNew
            !AcID = mPartyCode
            !SubCode = mPartyCode
            !Site_Code = mRecordSite
            !FirmCode = mRecordFirm
            !Name = left(Trim(StringPass(Master!Customer_Name)) & " [" & Trim(StringPass(Master!Customer_Code)) & "]", 40)
            !NameHelp = ReturnString(left(Trim(StringPass(Master!Customer_Name)) & " [" & Trim(StringPass(Master!Customer_Code)) & "]", 40))
            !GroupCode = mGroupCode
            !GroupNature = mGroupNature
            !Nature = mNature
            !AliasYN = "N"
            !ConPerson = left(Trim(StringPass(Master!First_Name) & " " & StringPass(Master!Middle_Name) & " " & StringPass(Master!Last_Name)), 40)
            !Add1 = left(Trim(StringPass(Master!Addr_L_1)), 40)
            !Add2 = left(Trim(StringPass(Master!Addr_L_2)), 40)
            !Add3 = ""
            !CityCode = mCityCode
            !Pin = StringPass(left(Master!Pin_Code, 6))
            !Phone = StringPass(Master!Phone)
            !Fax = StringPass(Master!Fax)
            !EMail = StringPass(Master!EMail)
            !ActiveYN = 1
            !Govt_YN = 0
            !CreditLimit = 0
            !CreditDays = 0
            !L_C = "L"
            !Party_Type = 0         ' For General Parties
            !SiebelCode = StringPass(Master!Customer_Code)
                        
            !U_Name = "Siebel"
            !U_EntDt = Format(PubLoginDate, "Short Date")
            !U_AE = "A"
            .Update
        End With
DuplicateSkipped:
        CodeCnt = CodeCnt + 1
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
MyNextRecord:
        Master.MoveNext
    Loop
        
    If PubBackEnd = "S" Then
        GCn.Execute ("Update SubGroupCounter Set SubGroupAcCode=" & CodeCnt & "")
    End If
    
    GCn.CommitTrans
    If PubBackEnd = "A" Then
        GCnFa.Execute ("Drop Table SubGroup")
        GCnFa.Execute ("Drop Table SubGroupAlias")
        GCnFa.BeginTrans
        
        GCnFa.Execute ("Select SubGroup.* into SubGroup from [" & DataPath & "].SubGroup")
        GCnFa.Execute ("Select SubGroup.* into .SubGroupAlias from [" & DataPath & "].SubGroup")
        CodeCnt = GCn.Execute("Select iif(isnull(Max(Val(right(SubCode,6)))),0,Max(Val(right(SubCode,6))))+1 from SubGroup").Fields(0).Value
        GCnFa.Execute ("Update SubGroupCounter Set SubGroupAcCode=" & CodeCnt & "")
        GCnFa.CommitTrans
    End If
    
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Ledger Account','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub
Private Sub VehiclePurchaseDataUpdate(Index)
'' On Error GoTo Eloop
Dim MasterCode As String, DocId As String, mV_Type As String, mPartyCode As String, mForm_Code As String
Dim mDebitAc As String, mMfgMonth As String, mMfgYear As String, mColourCode As String, mColourName As String, mGodownCode As String
Dim mTaxPer As Double, mDeductionCode As String, mAdditionCode As String
Dim mLength1 As Integer, mLength2 As Integer, mTaxOnDelivery As Boolean
Dim EditFlag As Boolean
Dim RsX As adodb.Recordset
Dim xDocId$

    ImportBtn(Index).BackColor = ProcessColor
    
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Veh_Purch1", GCn, adOpenDynamic, adLockOptimistic
    
    Set RsNew1 = New adodb.Recordset
    RsNew1.CursorLocation = adUseClient
    RsNew1.Open "Select * from Veh_Purch2", GCn, adOpenDynamic, adLockOptimistic
    
    Set rsTemp = New adodb.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open "Select * from Veh_Stock", GCn, adOpenDynamic, adLockOptimistic
    
    
    mV_Type = "V_PB"
    
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(V_No)", "0") & "+1 from Veh_Purch1 where Left(DocID,1)='" & PubDivCode & "' and " & cMID("DocID", "2", "1") & "='" & PubSiteCode & "' and V_Type='" & mV_Type & "'").Fields(0).Value
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Invoice_No"))) Or StringPass(Master.Fields("Invoice_No")) = "" Then GoTo MyNextRecord
        EditFlag = False
        
'Modi Arpit Because Telco Apply Vat After Some Days
'        If GCn.Execute("Select PBill_No from Veh_Purch1 where Pbill_No='" & left(StringPass(Master.Fields("Invoice_no")), 10) & "'").RecordCount > 0 Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Telco Invoice No. Already Exist in Automan")
'            GoTo MyNextRecord
'        End If
        
        If GCn.Execute("Select PBill_No from Veh_Purch1 where Pbill_No='" & left(StringPass(Master.Fields("Invoice_no")), 10) & "'").RecordCount > 0 Then
            EditFlag = True
        End If
'Modi End
                
        If StringPass(Master.Fields("Supplier_Name")) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Supplier Name is Empty")
            GoTo MyNextRecord
        Else
            If ErrorGCN.Execute("select AutomanAcCode from AccountConversion where Type='Vehicle Purchase' and SiebelAc='" & StringPass(Master.Fields("Supplier_Name")) & "'").RecordCount > 0 Then
                mPartyCode = ErrorGCN.Execute("select AutomanAcCode from AccountConversion where Type='Vehicle Purchase' and SiebelAc='" & StringPass(Master.Fields("Supplier_Name")) & "'").Fields(0).Value
            Else
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master!Supplier_Name), "Vehicle Purchase", "Automan A/c Code is not Defined in AccountConversionTable for This Supplier")
                GoTo MyNextRecord
            End If
        End If
        
        If IsNull(Master!Invoice_Date) Or Master!Invoice_Date = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Telco Invoice Date is Empty")
            GoTo MyNextRecord
        End If
        
        If IsNull(StringPass(Master!godown)) Or StringPass(Master!godown) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Godown Name is Empty")
            GoTo MyNextRecord
        End If
        
        If IsNull(StringPass(Master!VC_Number)) Or StringPass(Master!VC_Number) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "VC_Number is Empty")
            GoTo MyNextRecord
        Else
            If GCn.Execute("Select Model from Model where Model='" & StringPass(Master!VC_Number) & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "VC_Number is not Exist in Model Master")
                GoTo MyNextRecord
            End If
        End If
        
        If IsNull(StringPass(Master!Chassis_No)) Or StringPass(Master!Chassis_No) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Chassis Number is Empty")
            GoTo MyNextRecord
        End If
        
        If GCn.Execute("Select ChassisNo from Veh_Stock where ChassisNo='" & StringPass(Master.Fields("Chassis_No")) & "'").RecordCount > 0 Then
            If EditFlag = False Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master!Chassis_No, "Vehicle Purchase", "This Chassis No. Already Exist in Automan")
                GoTo MyNextRecord
            End If
        End If
        
        
        If IsNull(StringPass(Master!Narration)) Or StringPass(Master!Narration) = "" Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Narration Field is Empty (Engine Number)")
            GoTo MyNextRecord
        End If
            
        
        If Len(StringPass(Master.Fields("Chassis_No"))) = 17 Then
            If GCn.Execute("Select Name from Chas_Mth where Month_CD='" & Mid(StringPass(Master!Chassis_No), 12, 1) & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Chassis Mfg. Month Name is not defined in Chas_Mth Table")
                GoTo MyNextRecord
            Else
                mMfgMonth = GCn.Execute("Select Name from Chas_Mth where Month_CD='" & Mid(StringPass(Master!Chassis_No), 12, 1) & "'").Fields(0).Value
            End If
        ElseIf Len(StringPass(Master.Fields("Chassis_No"))) > 17 Then
            mMfgMonth = Format(Master.Fields("Invoice_Date"), "MMMM")
        Else
            If GCn.Execute("Select Name from Chas_Mth where Month_CD='" & Mid(StringPass(Master!Chassis_No), 7, 1) & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Chassis Mfg. Month Name is not defined in Chas_Mth Table")
                GoTo MyNextRecord
            Else
                mMfgMonth = GCn.Execute("Select Name from Chas_Mth where Month_CD='" & Mid(StringPass(Master!Chassis_No), 7, 1) & "'").Fields(0).Value
            End If
        End If
        
        If Len(StringPass(Master.Fields("Chassis_No"))) = 17 Then
            Select Case Val(Mid(StringPass(Master!Chassis_No), 10, 1))
                Case 9
                    mMfgYear = "2009"
                Case 0
                    mMfgYear = "2010"
                Case 1
                    mMfgYear = "2011"
                Case 2
                    mMfgYear = "2012"
                Case 3
                    mMfgYear = "2013"
                Case 4
                    mMfgYear = "2014"
                Case 5
                    mMfgYear = "2015"
                Case 6
                    mMfgYear = "2016"
                Case 7
                    mMfgYear = "2017"
                Case 8
                    mMfgYear = "2018"
            End Select
        ElseIf Len(StringPass(Master.Fields("Chassis_No"))) > 17 Then
            mMfgYear = Format(Master.Fields("Invoice_Date"), "YYYY")
        Else
            If GCn.Execute("Select Name from Chas_Yr where Year_Cd='" & Mid(StringPass(Master!Chassis_No), 8, 2) & "'").RecordCount = 0 Then
                Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Chassis Mfg. Year Name is not defined in Chas_YR Table")
                GoTo MyNextRecord
            Else
                mMfgYear = GCn.Execute("Select Name from Chas_Yr where Year_Cd='" & Mid(StringPass(Master!Chassis_No), 8, 2) & "'").Fields(0).Value
            End If
        End If
                
        If GCn.Execute("Select God_Code from Godown where Left(God_Name,20)='" & left(StringPass(Master.Fields("Godown")), 20) & "' and Appli_For=1").RecordCount = 0 Then
            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Godown Name not found Godown Master of Automan")
            GoTo MyNextRecord
        Else
            mGodownCode = GCn.Execute("Select God_Code from Godown where Left(God_Name,20)='" & left(StringPass(Master.Fields("Godown")), 20) & "' and Appli_For=1").Fields(0).Value
        End If
        
        mColourCode = GCn.Execute("Select Col_Code from Model where Model='" & StringPass(Master.Fields("VC_Number")) & "'").Fields(0).Value
        If mColourCode = "" Then
            mColourCode = ErrorGCN.Execute("Select DefaultColourCode from Enviro").Fields(0).Value
        End If
        mColourName = ""
'        If GCn.Execute("Select Col_Code from ColMast where Col_Code='" & mColourCode & "'").RecordCount = 0 Then
'            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Vehicle Purchase", "Colour Name not found Colour Master of Automan")
'            GoTo MyNextRecord
'        Else
'            mColourName = GCn.Execute("Select Col_Desc from ColMast where Col_Code='" & mColourCode & "'").Fields(0).Value
'        End If

        If GCn.Execute("Select Col_Code from ColMast where Col_Code='" & mColourCode & "'").RecordCount > 0 Then
            mColourName = GCn.Execute("Select Col_Desc from ColMast where Col_Code='" & mColourCode & "'").Fields(0).Value
        End If
        
        If eVal(Master.Fields("TAX CST")) > 0 Then
            mForm_Code = ErrorGCN.Execute("Select VehicleCstPurchFormCode from Enviro").Fields(0).Value
            mTaxPer = GCn.Execute("Select Tax_Per from TaxForms Where Form_Code='" & mForm_Code & "'").Fields(0).Value
        Else
            mForm_Code = ErrorGCN.Execute("Select VehiclePurchFormCode from Enviro").Fields(0).Value
            mTaxPer = ErrorGCN.Execute("Select VehiclePurchTaxPer from Enviro").Fields(0).Value
        End If
        
        mDeductionCode = ErrorGCN.Execute("Select ChassisDiscountItemCode from Enviro").Fields(0).Value
        mAdditionCode = ErrorGCN.Execute("Select ChassisTransportItemCode from Enviro").Fields(0).Value
        mTaxOnDelivery = ErrorGCN.Execute("Select TaxOnDeliveryCharges from Enviro").Fields(0).Value
        
        mDebitAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc Where Form_Code='" & mForm_Code & "' ").Fields(0).Value
        Dim mShortYear As String
        If Month(Master.Fields("Invoice_Date")) > 3 Then
            mShortYear = Right(Format(Master.Fields("Invoice_Date"), "yy"), 1) & Right(Val(Format(Master.Fields("Invoice_Date"), "yy")) + 1, 1)
        Else
            mShortYear = Right(Val(Format(Master.Fields("Invoice_Date"), "yy")) - 1, 1) & Right(Format(Master.Fields("Invoice_Date"), "yy"), 1)
        End If
        
        'DocId = PubDivCode & PubSiteCode & PubSiteCode & " " & mV_Type & "SBL" & Format(Master!Invoice_Date, "yy") & Right("00000000" & CodeCnt, 8)
        DocId = PubDivCode & PubSiteCode & PubSiteCode & " " & mV_Type & "SBL" & mShortYear & Right("00000000" & CodeCnt, 8)
        
        
        '' Calculation of Amount
        Dim mTot_Amt As Double, mTax_Amt As Double, mMisc_Amt As Double
        Dim mDeduction As Double, mAddition As Double, mAmount As Double
        
        mTot_Amt = 0: mTax_Amt = 0: mMisc_Amt = 0
        mDeduction = 0: mAddition = 0: mAmount = 0
        
        If Master.Fields("Chassis_No") = "445051HRZY00517" Then
            MsgBox ""
        End If
        
        If UCase(left(PubComp_Name, 3)) = "LMP" Then
            mTot_Amt = Master!Value
            
            If mTaxOnDelivery Then
                mMisc_Amt = 0
            Else
                mMisc_Amt = VNull(Master.Fields("Delivery Charges"))
            End If

            If IsNull(Master.Fields("VatTax")) Or Master.Fields("VatTax") = "" Then
                mTax_Amt = Round((mTot_Amt) * mTaxPer / (100 + mTaxPer), 2)     ''- mMisc_Amt
            Else
                mTax_Amt = Val(Master.Fields("VatTax"))
            End If

            If mTaxOnDelivery Then
                mAddition = VNull(Master.Fields("Delivery Charges"))
            Else
                mAddition = 0
            End If
            
            mAmount = mTot_Amt + mDeduction - (mMisc_Amt + mTax_Amt + mAddition)
        Else
            mTot_Amt = Master!Value
            
            If mTaxOnDelivery Then
                mMisc_Amt = 0
            Else
                mMisc_Amt = VNull(Master.Fields("Delivery Charges"))
            End If
            If eVal(Master.Fields("Tax Cst")) > 0 Then
                mTax_Amt = eVal(Master.Fields("Tax Cst"))
            Else
                If IsNull(Master.Fields("VatTax")) Or Master.Fields("VatTax") = "" Then
                    mTax_Amt = Round((mTot_Amt) * mTaxPer / (100 + mTaxPer), 2)     ''- mMisc_Amt
                Else
                    mTax_Amt = Val(Master.Fields("VatTax"))
                End If
            End If
            
            If mTaxOnDelivery Then
                mAddition = VNull(Master.Fields("Delivery Charges"))
            Else
                mAddition = 0
            End If
            mDeduction = VNull(Master.Fields("Total Discount"))
            mAmount = mTot_Amt + mDeduction - (mMisc_Amt + mTax_Amt + mAddition)
        End If
        
        
        If EditFlag = True Then
            'ArpitStart
            GCn.Execute "Update Veh_Purch1 Set Amount = " & mAmount & ", Tot_Amount = " & mTot_Amt & ", " & _
                        "Tax_Per = " & mTaxPer & ", Tax_Amt = " & mTax_Amt & ", Addition = " & mAddition & ", " & _
                        "Deduction = " & mDeduction & ", Misc_Amt = " & mMisc_Amt & ", U_EntDt = " & ConvertDate(Date) & " " & _
                        "Where DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(Master!Chassis_No) & "' )"
                        
            
            
            Set RsX = GCn.Execute("Select Pur_DocId From Veh_Stock Where ChassisNo = '" & Master!Chassis_No & "'")
            If RsX.RecordCount > 0 Then
                GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & XNull(RsX(0)) & "' And Trn_Type='D'"
                GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & XNull(RsX(0)) & "' And Trn_Type='A'"
            End If
            
            If mDeduction > 0 Then
                Set RsX = GCn.Execute("Select Pur_DocId From Veh_Stock Where ChassisNo = '" & Master!Chassis_No & "'")
                If RsX.RecordCount > 0 Then xDocId = XNull(RsX!Pur_DocId)
                
                If GCn.Execute("Select DocId From Veh_Purch2 Where DocId = '" & xDocId & "'").RecordCount > 0 Then
                    'GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & xDocId & "' And Trn_Type='D'"
                    'GCn.Execute "Update Veh_Purch2  Set Rate = " & mDeduction & " " & _
                                "Where DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(Master!Chassis_No) & "') And Trn_Type='D'"
                End If
                    With RsNew1
                        .AddNew
                        !DocId = xDocId
                        !Srl_No = 1
                        !Site_Code = PubSiteCode & PubSiteCode
                        !V_Type = mV_Type
                        !V_No = CodeCnt
                        !Trn_Type = "D"
                        !PROD_CODE = mDeductionCode
                        !Qty = 1
                        !Rate = mDeduction
                        
                        !U_Name = "Siebel"
                        !U_EntDt = Format(PubLoginDate, "Short Date")
                        !U_AE = "A"
                        .Update
                    End With
                'End If
            End If
            
            If mAddition > 0 Then
                Set RsX = GCn.Execute("Select Pur_DocId From Veh_Stock Where ChassisNo = '" & Master!Chassis_No & "'")
                If RsX.RecordCount > 0 Then xDocId = XNull(RsX!Pur_DocId)
                
                If GCn.Execute("Select DocId From Veh_Purch2 Where DocId = '" & xDocId & "'").RecordCount > 0 Then
                    'GCn.Execute "Delete From Veh_Purch2 Where DocId = '" & xDocId & "' And Trn_Type='A'"
                    'GCn.Execute "Update Veh_Purch2 Set Rate = " & mAddition & " " & _
                                "Where DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(Master!Chassis_No) & "') And Trn_Type='A'"
                End If
                    With RsNew1
                        .AddNew
                        !DocId = xDocId
                        !Srl_No = 2
                        !Site_Code = PubSiteCode & PubSiteCode
                        !V_Type = mV_Type
                        !V_No = CodeCnt
                        !Trn_Type = "A"
                        !PROD_CODE = mAdditionCode
                        !Qty = 1
                        !Rate = mAddition
                        
                        !U_Name = "Siebel"
                        !U_EntDt = Format(PubLoginDate, "Short Date")
                        !U_AE = "A"
                        .Update
                    End With
                'End If
            End If
                        
            GCn.Execute "Update Veh_Stock Set Rate = " & mAmount & ", VRate = " & mTot_Amt & " " & _
                        "Where Pur_DocId = (Select Pur_DocId From Veh_Stock Where ChassisNo = '" & StringPass(Master!Chassis_No) & "' )"
                        
            EditFlag = False
            'ArpitEnd
        Else
            'Insert New Rec
            With RsNew
                .AddNew
                !DocId = DocId
                !DocIDHelp = Replace(DocId, " ", "")
                !Site_Code = PubSiteCode & PubSiteCode
                !V_Type = mV_Type
                !V_No = CodeCnt
                !V_DATE = MakeDate(Master!Invoice_Date)
                !PartyCode = mPartyCode
                !PBill_No = Master!Invoice_No
                !Pbill_Date = MakeDate(Master!Invoice_Date)
                !BMS_Category = ErrorGCN.Execute("Select DefaultBMSCategory from Enviro").Fields(0).Value
                !DueDate = MakeDate(Master!Invoice_Date)
                !Gate = ""
                !GateDate = MakeDate(Master!Invoice_Date)
                !Form_Code = mForm_Code
                !Amount = mAmount
                !Addition = mAddition
                !Deduction = mDeduction
                !Exsice = 0
                !Tax_Per = mTaxPer
                !Tax_Amt = mTax_Amt
                !Misc_Amt = mMisc_Amt
                !Tot_AMOUNT = mTot_Amt
                !DrAcCode = mDebitAc
                
                !U_Name = "Siebel"
                !U_EntDt = Format(PubLoginDate, "Short Date")
                !U_AE = "A"
                .Update
            End With
            
            If mDeduction > 0 Then
                With RsNew1
                    .AddNew
                    !DocId = DocId
                    !Srl_No = 1
                    !Site_Code = PubSiteCode & PubSiteCode
                    !V_Type = mV_Type
                    !V_No = CodeCnt
                    !Trn_Type = "D"
                    !PROD_CODE = mDeductionCode
                    !Qty = 1
                    !Rate = mDeduction
                    
                    !U_Name = "Siebel"
                    !U_EntDt = Format(PubLoginDate, "Short Date")
                    !U_AE = "A"
                    .Update
                End With
            End If
            
            If mAddition > 0 Then
                With RsNew1
                    .AddNew
                    !DocId = DocId
                    !Srl_No = 2
                    !Site_Code = PubSiteCode & PubSiteCode
                    !V_Type = mV_Type
                    !V_No = CodeCnt
                    !Trn_Type = "A"
                    !PROD_CODE = mAdditionCode
                    !Qty = 1
                    !Rate = mAddition
                    
                    !U_Name = "Siebel"
                    !U_EntDt = Format(PubLoginDate, "Short Date")
                    !U_AE = "A"
                    .Update
                End With
            End If
            
            With rsTemp
                .AddNew
                !ChassisNo = StringPass(Master!Chassis_No)
                !Pur_DocId = DocId
                !pur_SrlNo = 1
                !Pur_DocIDHelp = Replace(DocId, " ", "")
                !Pur_SiteCode = PubSiteCode & PubSiteCode
                !Pur_VType = mV_Type
                !Pur_VNo = CodeCnt
                !Pur_VDate = MakeDate(Master!Invoice_Date)
                !Mfg_Month = mMfgMonth
                !Mfg_Yr = mMfgYear
                !InDate = MakeDate(Master!Invoice_Date)
                !Model = StringPass(Master!VC_Number)
                !Chas_Type = left(StringPass(Master!Chassis_No), 6)
                !godown = mGodownCode
                
                mLength1 = InStr(1, StringPass(Master!Narration), "Engine") + Len("Engine Number - ")
                mLength2 = InStr(1, StringPass(Master!Narration), "Chassis")
                mLength2 = (mLength2 - mLength1)
                !EngineNo = Replace(Trim(Mid(StringPass(Master!Narration), mLength1, mLength2)), ".", "")
                !Rate = mAmount
                !Fixed = 0
                !VRate = mTot_Amt
                !Colour_Code = mColourCode
                !Colours = mColourName
                !Tax_YN = 1
                !PBill_No = StringPass(Master!Invoice_No)
                !Pbill_Date = MakeDate(Master!Invoice_Date)
                !PartyCode = mPartyCode
                
                !U_Name = "Siebel"
                !U_EntDt = Format(PubLoginDate, "Short Date")
                !U_AE = "A"
                
                .Update
            End With
        End If
        CodeCnt = CodeCnt + 1
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
MyNextRecord:
        Master.MoveNext
    Loop
    GCn.CommitTrans
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Vehicle Purchase','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub


Private Sub ModelMasterUpdate(ByVal Index As Long)
'' On Error GoTo Eloop
Dim MasterCode As String, mCatCode As String, mGrpCode As String, mColCode As String

    ImportBtn(Index).BackColor = ProcessColor

    GCn.BeginTrans
    
    '' Model Category Updation
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Model_Cat", GCn, adOpenDynamic, adLockOptimistic
    'CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("ModelCat_Code", "2", "2")) & ")", "0") & " + 1 from Model_Cat").Fields(0).Value
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(Right(ModelCat_Code,2))", "0") & " from Model_Cat Where IsNumeric(Right(ModelCat_Code,2))=1").Fields(0).Value
    If IsNumeric(CodeCnt) Then
        CodeCnt = CodeCnt + 1
    Else
        CodeCnt = 30
    End If
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Parent Product Line"))) Or StringPass(Master.Fields("Parent Product Line")) = "" Then GoTo MyNextRecord
        
        If GCn.Execute("Select ModelCat_Name from Model_Cat where ModelCat_Name='" & left(StringPass(Master.Fields("Parent Product Line")), 20) & "'").RecordCount > 0 Then GoTo MyNextRecord
                
        'MasterCode = PubDivCode & Right("00" & CodeCnt, 2)
        MasterCode = PubDivCode & CodeCnt
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!ModelCat_Code = MasterCode
        RsNew!ModelCat_Name = left(StringPass(Master.Fields("Parent Product Line")), 20)
        
        RsNew!Site_Code = PubSiteCode
        RsNew!OldCode = ""
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    
    '' Model Group Updation
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Model_Grp", GCn, adOpenDynamic, adLockOptimistic
    'CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("ModelGrp_Code", "2", "4")) & ")", "0") & " + 1 from Model_Grp").Fields(0).Value
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(Right(ModelGrp_Code,4))", "0") & " from Model_Grp Where IsNumeric(Right(ModelGrp_Code,4))=1 ").Fields(0).Value
    If IsNumeric(CodeCnt) Then
        CodeCnt = CodeCnt + 1
    Else
        CodeCnt = 2000
    End If
    If Master.RecordCount > 0 Then Master.MoveFirst
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Product Line"))) Or StringPass(Master.Fields("Product Line")) = "" Then GoTo MyNextRecord1
        If GCn.Execute("Select ModelGrp_Name from Model_Grp where ModelGrp_Name='" & left(StringPass(Master.Fields("Product Line")), 20) & "'").RecordCount > 0 Then GoTo MyNextRecord1
                
        MasterCode = PubDivCode & Right("0000" & CodeCnt, 4)
        mCatCode = GCn.Execute("Select ModelCat_Code from Model_Cat where ModelCat_Name='" & left(StringPass(Master.Fields("Parent Product Line")), 20) & "'").Fields(0).Value
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!ModelGrp_Code = MasterCode
        RsNew!ModelGrp_Name = left(StringPass(Master.Fields("Product Line")), 20)
        RsNew!Wheel_Catg = "Four"
        RsNew!ModelCat_Code = mCatCode
        
        RsNew!Site_Code = PubSiteCode
        RsNew!OldCode = ""
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord1:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    
    '' Model Updation
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Model", GCn, adOpenDynamic, adLockOptimistic
    If Master.RecordCount > 0 Then Master.MoveFirst
    Do Until Master.EOF
        'Modi Arpit After Discussing With Salve
        'If IsNull(StringPass(Master.Fields("Product/VC#"))) Or StringPass(Master.Fields("Product/VC#")) = "" Or left(StringPass(Master.Fields("Product/VC#")), 1) <> "2" Then GoTo MyNextRecord2
        If IsNull(StringPass(Master.Fields("Product/VC#"))) Or StringPass(Master.Fields("Product/VC#")) = "" Then GoTo MyNextRecord2
        'Modi End
        If GCn.Execute("Select Model from Model where Model='" & left(StringPass(Master.Fields("Product/VC#")), 20) & "'").RecordCount > 0 Then GoTo MyNextRecord2
        
        mColCode = ""
        If IsNull(StringPass(Master.Fields("Parent Product Line"))) Or StringPass(Master.Fields("Parent Product Line")) = "" Then
            mCatCode = PubDivCode & "XX"
        Else
            mCatCode = GCn.Execute("Select ModelCat_Code from Model_Cat where ModelCat_Name='" & left(StringPass(Master.Fields("Parent Product Line")), 20) & "'").Fields(0).Value
        End If
        If IsNull(StringPass(Master.Fields("Product Line"))) Or StringPass(Master.Fields("Product Line")) = "" Then
            mGrpCode = PubDivCode & "XX"
        Else
            mGrpCode = GCn.Execute("Select ModelGrp_Code from Model_Grp where ModelGrp_Name='" & left(StringPass(Master.Fields("Product Line")), 20) & "'").Fields(0).Value
        End If
        If Not IsNull(StringPass(Master.Fields("Colour"))) And StringPass(Master.Fields("Colour")) <> "" Then
            If GCn.Execute("Select Col_Desc from ColMast where Col_Desc='" & left(Replace(StringPass(Master!Colour), "_", " "), 20) & "'").RecordCount > 0 Then
                mColCode = GCn.Execute("Select Col_Code from ColMast where Col_Desc='" & left(Replace(StringPass(Master!Colour), "_", " "), 20) & "'").Fields(0).Value
            Else
                Call InsertErrorMessage(Index, Master.AbsolutePosition, StringPass(Master!Colour), "Model Master", "Colour Name not found in Colour Master during Model Master Creation")
            End If
        End If
        'Insert New Rec
        RsNew.AddNew
        RsNew!Model = left(StringPass(Master.Fields("Product/VC#")), 20)
        If IsNull(StringPass(Master.Fields("Product Line"))) Or StringPass(Master.Fields("Product Line")) = "" Then
            RsNew!Chas_Type = "."
        Else
            RsNew!Chas_Type = left(StringPass(Master.Fields("Product Line")), 6)
        End If
        RsNew!Vehicle_Type = left(StringPass(Master!LOB), 5)
        If IsNull(StringPass(Master.Fields("Product Name"))) Or StringPass(Master.Fields("Product Name")) = "" Then
            RsNew!Sales_Desc = StringPass(Master.Fields("Product Line"))
        Else
            RsNew!Sales_Desc = left(StringPass(Master.Fields("Product Name")), 40)
        End If
            
        If IsNull(StringPass(Master.Fields("Product Description"))) Or StringPass(Master.Fields("Product Description")) = "" Then
            RsNew!Model_Desc = StringPass(Master.Fields("Product/VC#"))
            RsNew!Model_Desc1 = ""
            RsNew!Model_Desc2 = ""
        Else
            RsNew!Model_Desc = left(StringPass(Master.Fields("Product Description")), 50)
            RsNew!Model_Desc1 = Mid(StringPass(Master.Fields("Product Description")), 51, 50)
            RsNew!Model_Desc2 = Mid(StringPass(Master.Fields("Product Description")), 101, 50)
        End If
        RsNew!Grp_Code = mGrpCode
        RsNew!Cat_Code = mCatCode
        RsNew!Active_YN = 1
        RsNew!TyreDetails = left(StringPass(Master.Fields("Number & Description of Type")), 30)
        RsNew!HorsePower = left(StringPass(Master.Fields("Horse Power")), 10)
        RsNew!Front_A_Wt = left(StringPass(Master.Fields("Front Axle Weight")), 15)
        RsNew!Rear_A_Wt = left(StringPass(Master.Fields("Front Axle Weight")), 15)
        RsNew!Unladen_Wt = left(StringPass(Master.Fields("Unladen Weight")), 15)
        RsNew!Gross_Wt = left(StringPass(Master.Fields("Gross Vehicle Weight")), 15)
        If PubDivCode = "C" Then
            RsNew!WheelBase = Master.Fields("Wheel Base")
            RsNew!FuelTankCapacity = Val(VNull(Master.Fields("Fuel Tank")))
            RsNew!RearAxleMake = left(StringPass(Master.Fields("Rear Axle")), 30)
        End If
        RsNew!Cylinder = Master.Fields("Number of Cylinders")
        RsNew!Fuel = left(StringPass(Master.Fields("Fuel")), 10)
        RsNew!Manufacturer = "Tata Motors Ltd."
        RsNew!ServiceTax_YN = 1
        
        RsNew!Col_Code = mColCode
        RsNew!RegulatoryCertificate = left(StringPass(Master.Fields("Regulatory Certification")), 15)
        RsNew!SteeringType = left(StringPass(Master.Fields("Steering")), 20)
        RsNew!Vehicle_Drive = left(StringPass(Master.Fields("Vehicle Drive")), 6)
        RsNew!CubicCapacity = left(StringPass(Master.Fields("Cubic Capacity")), 10)
        RsNew!BodyType = left(StringPass(Master.Fields("Type of Body")), 25)
        RsNew!Model_Type = left(StringPass(RsNew!Chas_Type), 2)
        RsNew!Wheel_Catg = "Four"
        RsNew!RLW = "XXXX"
        
        Select Case StringPass(Master.Fields("Product Category"))
            Case "Fast Moving"
                RsNew!FMSN = "F"
            Case "Standard"
                RsNew!FMSN = "M"
            Case "Slow Moving"
                RsNew!FMSN = "S"
            Case Else
                RsNew!FMSN = "N"
        End Select
        
        RsNew!Site_Code = PubSiteCode
        RsNew!Div_Code = PubDivCode
        'RsNew!OldCode = ""
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord2:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    
    GCn.CommitTrans
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & StringPass(Master.Fields("Product Line")) & "','Model Master','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub


Private Sub FinancerMasterUpdate(ByVal Index As Long)
'' On Error GoTo Eloop
Dim MasterCode As String
Dim mFinGrpCode As String, mFinBankCode As String, mFieldName As String
    ImportBtn(Index).BackColor = ProcessColor

    GCn.BeginTrans
    
    '' Fin Bank Updation
    mFieldName = IIf(PubDivCode = "C", "Financed By", "Hypothecation")
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from FinBank", GCn, adOpenDynamic, adLockOptimistic
    
    'CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("FinBankCode", "1", "3")) & ")", "0") & "+1 from FinBank").Fields(0).Value
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(FinBankCode)", "0") & " from FinBank").Fields(0).Value
    If IsNumeric(CodeCnt) Then
        CodeCnt = CodeCnt + 1
    Else
        CodeCnt = 2000
    End If
    Do Until Master.EOF
        
        If IsNull(StringPass(Master.Fields(mFieldName))) Or StringPass(Master.Fields(mFieldName)) = "" Then GoTo MyNextRecord
        
        If RsNew.RecordCount > 0 Then RsNew.MoveFirst
        Do Until RsNew.EOF
            If ReturnString(RsNew!FinBankName) = ReturnString(left(StringPass(Master.Fields(mFieldName)), 35)) Then
                GoTo MyNextRecord
            End If
            RsNew.MoveNext
        Loop
                
        If CodeCnt > 999 Then
            'CodeCnt = GCn.Execute("Select iif(isnull(Max(Val(Mid(FinBankCode,2,2)))),0,Max(Val(Mid(FinBankCode,2,2))))+1 from FinBank where left(FinBankCode,1)='" & left(ReturnString(StringPass(Master.Fields(mFieldName))), 1) & "'").Fields(0).Value
            CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("FinBankCode", "2", "2")) & ")", "0") & " + 1 from FinBank where left(FinBankCode,1)='" & left(ReturnString(StringPass(Master.Fields(mFieldName))), 1) & "'").Fields(0).Value
            MasterCode = left(ReturnString(StringPass(Master.Fields(mFieldName))), 1) & Right("00" & CodeCnt, 2)
        Else
            MasterCode = Right("000" & CodeCnt, 3)
        End If
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!FinBankCode = MasterCode
        RsNew!FinBankName = left(StringPass(Master.Fields(mFieldName)), 35)
        RsNew!FinGrpCode = "PVTF"
        RsNew!Inv_Prefix = ""
        RsNew!Site_Code = PubSiteCode
        RsNew!xName = left(ReturnString(StringPass(Master.Fields(mFieldName))), 35)
        RsNew!OldCode = ""
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    
    
    '' Fin Bank Branch Updation
    CopyCnt = 0
    ErrorCnt = 0
    
    Set rsTemp = GCn.Execute("Select * From FinBank")
    
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from ContractFinance where FinCatg=0", GCn, adOpenDynamic, adLockOptimistic
    
    'CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("FinCode", "3", "4")) & ")", "0") & " + 1 from ContractFinance where Fincatg=0").Fields(0).Value
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(FinCode)", "0") & " from ContractFinance where Fincatg=0").Fields(0).Value
    If IsNumeric(CodeCnt) Then
        CodeCnt = CodeCnt + 1
    Else
        CodeCnt = 2000
    End If
    If Master.RecordCount > 0 Then Master.MoveFirst
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields(mFieldName))) Or StringPass(Master.Fields(mFieldName)) = "" Then GoTo MyNextRecord1
        If RsNew.RecordCount > 0 Then RsNew.MoveFirst
        Do Until RsNew.EOF
            If ReturnString(RsNew!FinName) = ReturnString(left(StringPass(Master.Fields(mFieldName)), 40)) Then
                GoTo MyNextRecord1
            End If
            RsNew.MoveNext
        Loop
        
        mFinGrpCode = ""
        mFinBankCode = ""
        '' Searching for Bank/Financer Name Code
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
        Do Until rsTemp.EOF
            If ReturnString(rsTemp!FinBankName) = ReturnString(left(StringPass(Master.Fields(mFieldName)), 35)) Then
                mFinGrpCode = rsTemp!FinGrpCode
                mFinBankCode = rsTemp!FinBankCode
                Exit Do
            End If
            rsTemp.MoveNext
        Loop
                
        If CodeCnt > 9999 Then
            'CodeCnt = GCn.Execute("Select iif(isnull(Max(Val(Mid(FinCode,3,3)))),0,Max(Val(Mid(FinCode,3,3))))+1 from ContractFinance where mid(FinCode,3,1)='" & left(ReturnString(StringPass(Master.Fields(mFieldName))), 1) & "'").Fields(0).Value
            CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("FinCode", "3", "3")) & ")", "0") & " + 1 from ContractFinance where " & cMID("FinCode", "3", "1") & "='" & left(ReturnString(StringPass(Master.Fields(mFieldName))), 1) & "'").Fields(0).Value
            MasterCode = PubSiteCode & "F" & left(ReturnString(StringPass(Master.Fields(mFieldName))), 1) & Right("000" & CodeCnt, 3)
        Else
            MasterCode = PubSiteCode & "F" & Right("0000" & CodeCnt, 4)
        End If
        
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!FinCode = MasterCode
        RsNew!FinName = left(StringPass(Master.Fields(mFieldName)), 40)
        RsNew!UnderFinGrp = mFinGrpCode
        RsNew!FinBankCode = mFinBankCode
        RsNew!FinCatg = 0
        RsNew!Add1 = ""
        RsNew!Add2 = ""
        RsNew!ContactPerson = ""
        RsNew!City = ""
        RsNew!PinCode = ""
        RsNew!Phone = ""
        RsNew!Fax = ""
        RsNew!Ac_YN = "N"
        RsNew!AcCode = ""
        
        RsNew!Site_Code = PubSiteCode
        RsNew!OldCode = ""
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord1:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    
    
    
    
    GCn.CommitTrans
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & StringPass(Master.Fields(mFieldName)) & "','Financer Master','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub
Private Sub InsertErrorMessage(ByVal Index As Integer, ByVal RecordNo As Long, ByVal ValueDetails, ByVal ColoumnDetail, ByVal ErrorDescription)
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & RecordNo & ",'" & ValueDetails & "','" & ColoumnDetail & "','" & ErrorDescription & "')")
End Sub

Private Sub InsSkipRecMessage(ByVal Index As Integer, ByVal RecordNo As Long, ByVal ValueDetails, ByVal ColoumnDetail, ByVal ErrorDescription)
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & RecordNo & ",'" & left(StringPass(ValueDetails), 50) & "','" & ColoumnDetail & "','" & ErrorDescription & "')")
End Sub


Private Function ReturnString(ByVal temp As String) As String
    temp = Replace(temp, " ", "")
    temp = Replace(temp, "_", "")
    temp = Replace(temp, "-", "")
    temp = Replace(temp, "/", "")
    temp = Replace(temp, "&", "")
    temp = Replace(temp, "(", "")
    temp = Replace(temp, ")", "")
    temp = Replace(temp, ",", "")
    temp = Replace(temp, "'", "")
    temp = Replace(temp, ".", "")
    ReturnString = temp
End Function

Private Sub GodownMasterUpdate(ByVal Index As Long)
'' On Error GoTo Eloop
Dim MasterCode As String
    ImportBtn(Index).BackColor = ProcessColor
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Godown", GCn, adOpenDynamic, adLockOptimistic
    'CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("God_Code", "2", "2")) & ")", "0") & " + 1 from Godown").Fields(0).Value
    CodeCnt = GCn.Execute("Select " & xIsNull("Max(God_Code)", "0") & "  from Godown").Fields(0).Value
    If IsNumeric(CodeCnt) Then
        CodeCnt = CodeCnt + 1
    Else
        CodeCnt = 2000
    End If
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Godown"))) Or StringPass(Master.Fields("Godown")) = "" Then GoTo MyNextRecord
        
        If GCn.Execute("Select God_Name from Godown where Left(God_Name,20)='" & left(StringPass(Master.Fields("Godown")), 20) & "'").RecordCount > 0 Then GoTo MyNextRecord
                
        MasterCode = PubSiteCode & Right("00" & CodeCnt, 2)
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!God_Code = MasterCode
        RsNew!God_Name = left(StringPass(Master.Fields("Godown")), 30)
        RsNew!Appli_For = 1
        RsNew!Site_Code = PubSiteCode
        
        RsNew!OldCode = ""
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    GCn.CommitTrans
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & StringPass(Master.Fields("Godown")) & "','Godown Master','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
End Sub
Private Sub PurposeMasterUpdate(ByVal Index As Long)
'' On Error GoTo Eloop
Dim MasterCode As String
    ImportBtn(Index).BackColor = ProcessColor

    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Purpose", GCn, adOpenDynamic, adLockOptimistic
    
    'CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("PurposeCode", "1", "2")) & ")", "0") & "+1 from Purpose").Fields(0).Value
    CodeCnt = GCn.Execute("Select " & xIsNull("Max(PurposeCode)", "0") & " from Purpose").Fields(0).Value
    If IsNumeric(CodeCnt) Then
        CodeCnt = CodeCnt + 1
    Else
        CodeCnt = 2000
    End If
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Intended Use"))) Or StringPass(Master.Fields("Intended Use")) = "" Then GoTo MyNextRecord
        
        If GCn.Execute("Select PurposeName from Purpose where PurposeName='" & left(StringPass(Master.Fields("Intended Use")), 25) & "'").RecordCount > 0 Then GoTo MyNextRecord
                
        MasterCode = Right("00" & CodeCnt, 2)
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!PurposeCode = MasterCode
        RsNew!PurposeName = left(StringPass(Master.Fields("Intended Use")), 25)
        RsNew!Site_Code = PubSiteCode
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    GCn.CommitTrans
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & StringPass(Master.Fields("Intended Use")) & "','Purpose Master','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
''    MsgBox err.NUMBER & " " & err.Description, vbCritical, "Error in Updation"
''    If err.NUMBER = -2147467259 Then
''        ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    ElseIf err.NUMBER <> 3356 Then
''        ErrorGCN.Execute ("insert into prnmissrec(code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    End If
End Sub

Private Sub RefByMasterUpdate(ByVal Index As Long)
' On Error GoTo Eloop
Dim MasterCode As String
    ImportBtn(Index).BackColor = ProcessColor

    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Reffered", GCn, adOpenDynamic, adLockOptimistic
    'CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("RefCode", "1", "4")) & ")", "0") & " + 1 from Reffered").Fields(0).Value
    CodeCnt = GCn.Execute("Select " & xIsNull("Max(RefCode)", "0") & "  from Reffered").Fields(0).Value
    If IsNumeric(CodeCnt) Then
        CodeCnt = CodeCnt + 1
    Else
        CodeCnt = 2000
    End If
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("Source of Prospect"))) Or StringPass(Master.Fields("Source of Prospect")) = "" Then GoTo MyNextRecord
        
        If GCn.Execute("Select RefName from Reffered where RefName='" & left(StringPass(Master.Fields("Source of Prospect")), 25) & "'").RecordCount > 0 Then GoTo MyNextRecord
                
        MasterCode = Right("0000" & CodeCnt, 4)
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!RefCode = MasterCode
        RsNew!RefName = left(StringPass(Master.Fields("Source of Prospect")), 25)
        RsNew!Site_Code = PubSiteCode
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    GCn.CommitTrans
'    GCnFa.Execute ("Delete from City")
'    GCnFa.Execute ("Insert into City Select * from [" & App_Path & "].City")
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & StringPass(Master.Fields("Source of Prospect")) & "','Ref. By Master','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
''    MsgBox err.NUMBER & " " & err.Description, vbCritical, "Error in Updation"
''    If err.NUMBER = -2147467259 Then
''        ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    ElseIf err.NUMBER <> 3356 Then
''        ErrorGCN.Execute ("insert into prnmissrec(code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    End If
End Sub
Private Sub SalesRepMasterUpdate(ByVal Index As Long)
' On Error GoTo Eloop
Dim MasterCode As String, EmpName As String
    ImportBtn(Index).BackColor = ProcessColor

    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Emp_Mast", GCn, adOpenDynamic, adLockOptimistic
    'CodeCnt = GCn.Execute("Select iif(isnull(Max(Val(Mid(Emp_Code,2,3)))),0,Max(Val(Mid(Emp_Code,2,3))))+1 from Emp_Mast").Fields(0).Value
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("Emp_Code", "2", "3")) & ")", "0") & " + 1 from Emp_Mast").Fields(0).Value
    Do Until Master.EOF
        If IsNull(StringPass(Master.Fields("First Name"))) Or StringPass(Master.Fields("First Name")) = "" Then GoTo MyNextRecord
        If StringPass(Master.Fields("Middle Initial")) = "" Then
            EmpName = left(Trim(StringPass(Master.Fields("First Name")) & " " & StringPass(Master.Fields("Last Name"))), 40)
        Else
            EmpName = left(Trim(StringPass(Master.Fields("First Name")) & " " & StringPass(Master.Fields("Middle Initial")) & " " & StringPass(Master.Fields("Last Name"))), 40)
        End If
        
        If GCn.Execute("Select Emp_Name from Emp_Mast where Emp_Name='" & EmpName & "'").RecordCount > 0 Then GoTo MyNextRecord
                
        MasterCode = PubSiteCode & Right("000" & CodeCnt, 3)
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!Emp_Code = MasterCode
        RsNew!emp_name = EmpName
        RsNew!xName = Replace(EmpName, " ", "")
        RsNew!Emp_Type = 0
        RsNew!FName = ""
        RsNew!Add1 = left(StringPass(Master!Address), 40)
        RsNew!Add2 = Mid(StringPass(Master!Address), 41, 40)
        RsNew!CityName = left(StringPass(Master!City), 25)
        RsNew!Phone = left(StringPass(Master.Fields("Home Phone #")), 14)
        RsNew!Pager = ""
        RsNew!Mobile = left(StringPass(Master.Fields("Work Phone #")), 14)
        RsNew!Reference = StringPass(Master.Fields("User ID"))
        RsNew!Designation = StringPass(Master!Responsibility)
        
        RsNew!Site_Code = PubSiteCode
        RsNew!Div_Code = PubDivCode
        RsNew!OldCode = ""
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = Format(PubLoginDate, "Short Date")
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    GCn.CommitTrans
'    GCnFa.Execute ("Delete from City")
'    GCnFa.Execute ("Insert into City Select * from [" & App_Path & "].City")
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & EmpName & "','Sales Rep.','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
''    MsgBox err.NUMBER & " " & err.Description, vbCritical, "Error in Updation"
''    If err.NUMBER = -2147467259 Then
''        ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    ElseIf err.NUMBER <> 3356 Then
''        ErrorGCN.Execute ("insert into prnmissrec(code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    End If
End Sub
Private Sub AreaMasterUpdate(ByVal Index As Long)
' On Error GoTo Eloop
Dim CityCode As String
    ImportBtn(Index).BackColor = ProcessColor
   
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from Area", GCn, adOpenDynamic, adLockOptimistic
    'CodeCnt = GCn.Execute("Select iif(isnull(Max(Val(Mid(AreaCode,1,3)))),0,Max(Val(Mid(AreaCode,1,3))))+1 from Area").Fields(0).Value
    'CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("AreaCode", "1", "3")) & ")", "0") & "+1 from Area").Fields(0).Value
    CodeCnt = GCn.Execute("Select " & xIsNull("Max(AreaCode)", "0") & " from Area").Fields(0).Value
    If IsNumeric(CodeCnt) Then
        CodeCnt = CodeCnt + 1
    Else
        CodeCnt = 2000
    End If
    Do Until Master.EOF
        If IsNull(StringPass(Master!Site)) Or StringPass(Master!Site) = "" Then GoTo MyNextRecord
        If GCn.Execute("Select AreaName from Area where AreaName='" & Replace(left(StringPass(Master!Site), 15), "A/P ", "") & "'").RecordCount > 0 Then GoTo MyNextRecord
                
        CityCode = Right("000" & CodeCnt, 3)
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!AreaCode = CityCode
        RsNew!AreaName = Replace(left(StringPass(Master!Site), 15), "A/P ", "")
        RsNew!Site_Code = PubSiteCode
        RsNew!OldCode = ""
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = PubLoginDate
        RsNew!U_AE = "A"
        RsNew.Update
        
        CodeCnt = CodeCnt + 1
MyNextRecord:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    GCn.CommitTrans
'    GCnFa.Execute ("Delete from City")
'    GCnFa.Execute ("Insert into City Select * from [" & App_Path & "].City")
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & StringPass(Master!Site) & "','Area/Site Name','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
''    MsgBox err.NUMBER & " " & err.Description, vbCritical, "Error in Updation"
''    If err.NUMBER = -2147467259 Then
''        ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    ElseIf err.NUMBER <> 3356 Then
''        ErrorGCN.Execute ("insert into prnmissrec(code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    End If
End Sub
Private Sub ColourMasterUpdate(ByVal Index As Long)
' On Error GoTo Eloop
Dim CityCode As String
    ImportBtn(Index).BackColor = ProcessColor
   
    GCn.BeginTrans
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from ColMast", GCn, adOpenDynamic, adLockOptimistic
    'CodeCnt = GCn.Execute("Select iif(isnull(Max(Val(Mid(Col_Code,2,3)))),0,Max(Val(Mid(Col_code,2,3))))+1 from ColMast").Fields(0).Value
    CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("Col_Code", "2", "3")) & ")", "0") & " + 1 from ColMast").Fields(0).Value
    Do Until Master.EOF
        If IsNull(StringPass(Master!Colour)) Or StringPass(Master!Colour) = "" Then GoTo MyNextRecord
        If GCn.Execute("Select Col_Desc from ColMast where Col_Desc='" & Replace(StringPass(Master!Colour), "_", " ") & "'").RecordCount > 0 Then GoTo MyNextRecord
                
        CityCode = PubSiteCode & Right("000" & CodeCnt, 3)
        
        'Insert New Rec
        RsNew.AddNew
        RsNew!Col_Code = CityCode
        RsNew!Site_Code = PubSiteCode
        RsNew!Col_Desc = Replace(StringPass(Master!Colour), "_", " ")
        RsNew!OldCode = ""
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = PubLoginDate
        RsNew!U_AE = "A"
        RsNew.Update
        CodeCnt = CodeCnt + 1
MyNextRecord:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    GCn.CommitTrans
'    GCnFa.Execute ("Delete from City")
'    GCnFa.Execute ("Insert into City Select * from [" & App_Path & "].City")
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:
    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & StringPass(Master!Colour) & "','Colour Name','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
''    MsgBox err.NUMBER & " " & err.Description, vbCritical, "Error in Updation"
''    If err.NUMBER = -2147467259 Then
''        ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    ElseIf err.NUMBER <> 3356 Then
''        ErrorGCN.Execute ("insert into prnmissrec(code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    End If
End Sub
Private Sub CityMasterUpdate(ByVal Index As Long)
' On Error GoTo Eloop
Dim mASC As Integer, mASC1 As Integer, mASC2 As Integer
Dim mDefState As Byte
Dim CityCode As String
    ImportBtn(Index).BackColor = ProcessColor


    mDefState = Val(ErrorGCN.Execute("Select IIF(IsNull(State),1,State) From Enviro").Fields(0))
    
    Set rsTemp = New adodb.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open "select * from City", GCn, adOpenDynamic, adLockOptimistic
   
   
   
    CopyCnt = 0
    ErrorCnt = 0
    Set RsNew = New adodb.Recordset
    RsNew.CursorLocation = adUseClient
    RsNew.Open "Select * from City order by citycode", GCn, adOpenDynamic, adLockOptimistic
    'RsNew.Sort = "XName"
    'Stopped because there are alfanumeric values in CityCode
    'If StrCmp(left(PubComp_Name, 4), "Enar") Then
        CodeCnt = GCn.Execute("Select " & vIsNull("Max(Citycode)", "6000") & "  from City").Fields(0).Value
        If IsNumeric(CodeCnt) Then
            CodeCnt = CodeCnt + 1
        Else
            CodeCnt = 6000
        End If
    'Else
    '    CodeCnt = GCn.Execute("Select " & vIsNull("Max(" & cVal(cMID("Citycode", "2", "3")) & ")", "0") & " + 1 from City").Fields(0).Value
    'End If
    
    Do Until Master.EOF
        If IsNull(StringPass(Master!City)) Or StringPass(Master!City) = "" Then GoTo MyNextRecord
        If GCn.Execute("Select CityName from City where CityHelp='" & left(Replace(Trim(StringPass(Master!City)), " ", ""), 25) & "'").RecordCount > 0 Then GoTo MyNextRecord
                
        If CodeCnt > 999 Then
            mASC = 48
            mASC1 = 32
            mASC2 = 32
            If RsNew.RecordCount > 0 Then RsNew.MoveFirst
            Do Until RsNew.EOF
                If RsNew!CityCode = PubSiteCode & Chr(mASC2) & Chr(mASC1) & Chr(mASC) Then
                    RsNew.MoveFirst
                    mASC = mASC + 1
                    If mASC > 57 And mASC < 65 Then
                        mASC = 65
                    ElseIf mASC > 90 Then
                        If mASC1 = 32 Then mASC1 = 48 Else mASC1 = mASC1 + 1
                        mASC = 48
                        If mASC1 > 57 And mASC1 < 65 Then
                            mASC1 = 65
                        ElseIf mASC1 > 90 Then
                            If mASC2 = 32 Then mASC2 = 48 Else mASC2 = mASC2 + 1
                                mASC = 48
                                mASC1 = 48
                            If mASC2 > 57 And mASC2 < 65 Then
                                mASC2 = 65
                            ElseIf mASC2 > 90 Then
                                MsgBox "Code Limit for City Master has been Over, Please Contact to Administrator/Dataman", vbCritical, "Warning !!!"
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    RsNew.MoveNext
                End If
            Loop
            CityCode = PubSiteCode & Chr(mASC2) & Chr(mASC1) & Chr(mASC)    ''PubSiteCode & Right("000" & CodeCnt, 3)
        Else
            CityCode = PubSiteCode & Right("000" & CodeCnt, 3)
        End If

        GCn.BeginTrans
        'Insert New Rec
        RsNew.AddNew
        RsNew!xName = left(Replace(Trim(StringPass(Master!City)), " ", ""), 25)
        RsNew!CityCode = CityCode
        RsNew!Site_Code = PubSiteCode
        RsNew!CityName = left(Trim(StringPass(Master!City)), 25)
        RsNew!CityHelp = left(Replace(Trim(StringPass(Master!City)), " ", ""), 25)
        RsNew!LocalCentral = "L"
        RsNew!StateCode = mDefState
        RsNew!OldCode = ""
        RsNew!U_Name = "Siebel"
        RsNew!U_EntDt = PubLoginDate
        RsNew!U_AE = "A"
        RsNew.Update
        GCn.CommitTrans
        
        CodeCnt = CodeCnt + 1
MyNextRecord:
        CopyCnt = CopyCnt + 1
        lblRecCopy(Index).Caption = CopyCnt
        lblRecCopy(Index).Refresh
        Master.MoveNext
    Loop
    If PubBackEnd = "A" Then
        GCnFa.Execute ("Delete from City")
        GCnFa.Execute ("Insert into City Select * from [" & App_Path & "].City")
    End If
    ImportBtn(Index).BackColor = FinishColor
    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
    
lblExit:
    Set RsNew = Nothing
    Exit Sub
Eloop:

    ErrorCnt = ErrorCnt + 1
    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & StringPass(Master!City) & "','CityName','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
    Resume Next
''    MsgBox err.NUMBER & " " & err.Description, vbCritical, "Error in Updation"
''    If err.NUMBER = -2147467259 Then
''        ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    ElseIf err.NUMBER <> 3356 Then
''        ErrorGCN.Execute ("insert into prnmissrec(code,colname,details) values(" & Master.AbsolutePosition & ",'" & CityCode & "','CityCode','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''        Resume Next
''    End If
End Sub

Private Sub optAuto_Click(Index As Integer)
    Select Case Index
        Case 0
            CmdConvert.Enabled = False
        Case 1
            CmdConvert.Enabled = True
    End Select
End Sub

Private Sub Text1_Click()
Dim mFileName As String, filepath As Byte
' On Error GoTo ErrHandler
    mFileName = ""
    CommonDialog1.InitDir = Pub_DataPath
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Select XLS File for Stock Updation"
    'CommonDialog1.
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "Excel Files (*.xls)|*.xls" '|Text Files (*.txt)|*.Txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
    Txt(filepath) = CommonDialog1.FileName
    mFileName = CommonDialog1.FileTitle
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub
Private Function StringPass(ByVal temp As Variant) As String
    temp = XNull(temp)
    temp = Replace(temp, "'", "`")
    StringPass = temp
End Function

Private Sub lblRefresh()

End Sub


'''' On Error GoTo Eloop
''Dim MasterCode As String, mDocId As String, mPartyCode As String, mLength As Integer, mV_Type As String
''Dim mRecordSite As String, mRecordDiv As String, mRecordFirm As String, mOrderQty As Double, mPhysicalQty As Double
''Dim mPrefix As String, mName As String, mDebitAc As String, mFormCode As String
''Dim mOrderNo As String, mChallanNo As String
''Dim mSrl As Integer, mQty As Double, mCount As Integer, mAmount As Double
''Dim TranFalg As Boolean, TranFlag1 As Boolean, mLocal As String
''Dim mTax_Amt As Double, mTax_Amt1 As Double, mTaxable As Boolean
''Dim mVATApplicable As Boolean
''
''    ImportBtn(Index).BackColor = ProcessColor
''
''    GCn.BeginTrans
''    CopyCnt = 0
''    ErrorCnt = 0
''    Set RsNew = New adodb.Recordset
''    RsNew.CursorLocation = adUseClient
''    RsNew.Open "Select * from SP_Purch", GCn, adOpenDynamic, adLockOptimistic
''
''    Set RsNew1 = New adodb.Recordset
''    RsNew1.CursorLocation = adUseClient
''    RsNew1.Open "Select * from SP_Stock", GCn, adOpenDynamic, adLockOptimistic
''
''    mVATApplicable = ErrorGCN.Execute("Select VatApplicable from Enviro").Fields(0).Value
''
''    TranFalg = False
''    TranFlag1 = False
''    mV_Type = "SXGRT"
''    If Master.RecordCount > 0 Then Master.MoveFirst
''
''    Do Until Master.EOF
''        'mOrderNo = XNull(Master.Fields("Order ID").Value)
''        mChallanNo = XNull(Master.Fields("Transaction #").Value)
''
''        If IsNull(StringPass(Master.Fields("Transaction #"))) Or StringPass(Master.Fields("Transaction #")) = "" Then GoTo MyNextRecord
''
''        If Master!Type <> "Receive Internal" Then GoTo DuplicateSkipped
''
''        If GCn.Execute("Select V_no from SP_Purch where SiebelDocID&Party_Doc_No='" & StringPass(Master.Fields("Transaction #")) & StringPass(Master.Fields("Transaction #")) & "' and V_Type='" & mV_Type & "'").RecordCount > 0 Then
''            GoTo DuplicateSkipped
''        End If
''
''        If IsNull(StringPass(Master.Fields("Transaction Date/Time"))) Or StringPass(Master.Fields("Transaction Date/Time")) = "" Then
''            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (Inward)", "Transaction Date/Time is Empty")
''            GoTo MyNextRecord
''        End If
''
''        If IsNull(StringPass(Master.Fields("Transaction #"))) Or StringPass(Master.Fields("Transaction #")) = "" Then
''            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (Inward)", "Transaction # is Empty")
''            GoTo MyNextRecord
''        End If
''
''        If IsNull(StringPass(Master.Fields("Part #").Value)) Or StringPass(Master.Fields("Part #").Value) = "" Then
''            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (Inward)", "Part Number field is Empty")
''            GoTo MyNextRecord
''        Else
''            If GCn.Execute("Select Part_No from Part where Part_No='" & Master.Fields("Part #") & "'").RecordCount = 0 Then
''                Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Part #"), "Stock Transfer (Inward)", "Part Number not exist in AUTOMAN Part Master")
''                'GoTo MyNextRecord
''            End If
''        End If
''
''        If IsNull(StringPass(Master.Fields("Qty"))) Or StringPass(Master.Fields("Qty")) = "" Then
''            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (Inward)", "Qty field is Empty")
''            GoTo MyNextRecord
''        End If
''
''        If IsNull(StringPass(Master.Fields("Destination Division"))) Or StringPass(Master.Fields("Destination Division")) = "" Then
''            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (Inward)", "Division Name field is Empty")
''            GoTo MyNextRecord
''        Else
''            If ErrorGCN.Execute("select * from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").RecordCount > 0 Then
''                mRecordSite = ErrorGCN.Execute("select AutomanSite from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").Fields(0).Value
''                mRecordDiv = ErrorGCN.Execute("select AutomanDiv from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").Fields(0).Value
''                mRecordFirm = ErrorGCN.Execute("select AutomanFirm from SiteDivision where SiebelDiv='" & StringPass(Master.Fields("Destination Division")) & "'").Fields(0).Value
''            Else
''                Call InsSkipRecMessage(Index, Master.AbsolutePosition, StringPass(Master.Fields("Destination Division")), "Stock Transfer (Inward)", "Automan Site/Division is not Defined in SiteDivision Table for this Siebel Division")
''                GoTo MyNextRecord
''            End If
''        End If
''
''        If IsNull(StringPass(Master.Fields("Source Division"))) Or StringPass(Master.Fields("Source Division")) = "" Then
''            Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (Inward)", "Source Division (Vendor Name/Supplied from) field is Empty")
''            GoTo MyNextRecord
''        End If
''
''        If ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Source Division")) & "' and Type='" & Master.Fields("Type") & "'").RecordCount > 0 Then
''            mPartyCode = ErrorGCN.Execute("Select AutomanAcCode from AccountConversion where SiebelAc='" & StringPass(Master.Fields("Source Division")) & "' and Type='" & Master.Fields("Type") & "'").Fields(0).Value
''        Else
''            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Source Division"), "Stock Transfer (Inward)", "This Source Division (Vendor Name/Supplied from) is not defined in AccountConversion Table (As Account Code)")
''            GoTo MyNextRecord
''        End If
''
''        mPrefix = "SBL" & Format(Master.Fields("Transaction Date/Time"), "yy")
''        CodeCnt = GCn.Execute("Select iif(isnull(Max(V_No)),0,Max(V_no))+1 from SP_Purch where Left(DocID,1)='" & mRecordDiv & "' and mid(DocID,2,2)='" & mRecordSite & mRecordSite & "' and V_Type='" & mV_Type & "'").Fields(0).Value
''        mDocId = mRecordDiv & mRecordSite & mRecordSite & mV_Type & mPrefix & Right("00000000" & CodeCnt, 8)
''        mFormCode = ErrorGCN.Execute("Select SparePurchFormStockTrfIn from Enviro").Fields(0).Value
''        mDebitAc = GCn.Execute("Select PurSal_Ac_Code from TaxFormsAc where Div_Code='" & mRecordDiv & "' and Form_Code='" & mFormCode & "'").Fields(0).Value
''        mLocal = "L"
''        'Insert New Rec
''        With RsNew
''            .AddNew
''            TranFalg = True
''            !DocId = mDocId
''            !DocIDHelp = Replace(mDocId, " ", "")
''            !Site_Code = mRecordSite & mRecordSite
''            !V_Type = Trim(mV_Type)
''            !V_No = CodeCnt
''            !V_Date = Format(Master.Fields("Transaction Date/Time"), "dd/MMM/yyyy")
''            !Party_Code = mPartyCode
''            !Cash_Credit = "Credit"
''            !Party_Name = Master.Fields("Source Division")
''            !L_C = mLocal
''            !Form_Code = mFormCode
''            !Party_Doc_No = left(StringPass(Master.Fields("Transaction #")), 10)
''            !Party_Doc_Date = Format(Master.Fields("Transaction Date/Time"), "dd/MMM/yyyy")
''            !DrAc_Code = mDebitAc
''            !SiebelDocID = Master.Fields("Transaction #")
''
''            !U_Name = "Siebel"
''            !U_EntDt = Format(PubLoginDate, "Short Date")
''            !U_AE = "A"
''        End With
''
''        mSrl = 1
''        mQty = 0
''        mCount = 0
''        mAmount = 0
''        mTax_Amt = 0
''        mTax_Amt1 = 0
''        If (Master.EOF = False And Master.BOF = False) Then
''            Do While mChallanNo = XNull(Master.Fields("Transaction #"))
''                If mSrl <> 1 Then
''                    If IsNull(StringPass(Master.Fields("Part #").Value)) Or StringPass(Master.Fields("Part #").Value) = "" Then
''                        Call InsSkipRecMessage(Index, Master.AbsolutePosition, "", "Stock Transfer (Inward)", "Part Number field is Empty")
''                        GoTo LineFileNextrecord
''                    Else
''                        If GCn.Execute("Select Part_No from Part where Part_No='" & Master.Fields("Part #") & "'").RecordCount = 0 Then
''                            Call InsSkipRecMessage(Index, Master.AbsolutePosition, Master.Fields("Part #"), "Stock Transfer (Inward)", "Part Number not exist in AUTOMAN Part Master")
''                            'GoTo MyNextRecord      '' skipping of record not required at this position
''                        End If
''                    End If
''                End If
''
''                'Insert New Rec
''                With RsNew1
''                    .AddNew
''                    TranFlag1 = True
''                    !DocId = mDocId
''                    !Site_Code = mRecordSite & mRecordSite
''                    !V_Type = Trim(mV_Type)
''                    !V_No = CodeCnt
''                    !V_Date = Format(Master.Fields("Transaction Date/Time"), "dd/MMM/yyyy")
''                    !Party_Code = mPartyCode
''                    !Srl_No = mSrl
''                    !L_C = mLocal
''                    !Part_No = VNull(Master.Fields("part #"))
''                    !godown = GCn.Execute("Select SprWorksGodown from Syctrl").Fields(0).Value
''                    !Qty_Doc = Val(Master!Qty)
''                    !Qty_Rec = Val(Master!Qty)
''                    If mVATApplicable Then
''                        !Tax_YN = 1             '' if VAT is applicable in State
''                    Else
''                        !Tax_YN = IIf(mLocal = "L", 0, 1)
''                    End If
''                    !MRP_YN = IIf(mRecordDiv = "C", 1, 0)
''
''                    !Amount = Val(Master!Value)
''                    !Net_Amt = Val(Master!Value)
''                    !Rate = Master!Rate
''
''                    !Part_SrlNo = mSrl
''                    !TaxAmt = 0
''                    !TaxPer = 0
''                    !Disc_Per = 0
''                    !Disc_Amt = 0
''                    !Ord_Discper = 0
''                    !Ord_DiscAmt = 0
''
''                    !U_Name = "Siebel"
''                    !U_EntDt = Format(PubLoginDate, "Short Date")
''                    !U_AE = "A"
''                    .Update
''                    TranFlag1 = False
''                End With
''LineFileNextrecord:
''                mQty = mQty + Master!Qty
''                mCount = mCount + 1
''                mAmount = mAmount + Val(Master!Value)
''
''                mSrl = mSrl + 1
''                Master.MoveNext
''                CopyCnt = CopyCnt + 1
''                lblRecCopy(Index).Caption = CopyCnt
''                lblRecCopy(Index).Refresh
''                If Master.EOF = True Then Exit Do
''            Loop
''        End If
''        If Master.EOF = True Then Master.MovePrevious
''        With RsNew
''            !Tot_No_Of_Items = mCount
''            !Tot_Doc_Qty = mQty
''            !Tot_Phy_Qty = mQty
''            !TOT_Amt = mAmount
''            !Tot_Disc_Amt = 0
''            !Tot_Ord_DiscAmt = 0
''            !Tot_Goods_value = mAmount
''            !Tax_Amt = 0
''            !Addition = 0
''            !Deduction = 0
''            !Net_Amt = mAmount
''            .Update
''            TranFalg = False
''        End With
''        If Master.AbsolutePosition = Master.RecordCount Then Master.MoveNext
''        CodeCnt = CodeCnt + 1
''        GoTo NextLoop
''
''DuplicateSkipped:
''
''MyNextRecord:
''        mSrl = 0
''        If Master.EOF = False And Master.BOF = False Then
''            Do While (Master.EOF = False And Master.BOF = False) And mChallanNo = XNull(Master.Fields("Transaction #"))
''                mSrl = mSrl + 1
''                Master.MoveNext
''                CopyCnt = CopyCnt + 1
''                lblRecCopy(Index).Caption = CopyCnt
''                lblRecCopy(Index).Refresh
''                If Master.EOF = True Then Exit Do
''            Loop
''        End If
''        CopyCnt = CopyCnt + mSrl - 1
''        lblRecCopy(Index).Caption = CopyCnt
''        lblRecCopy(Index).Refresh
''NextLoop:
''    Loop
''    GCn.CommitTrans
''
''    ImportBtn(Index).BackColor = FinishColor
''    If Not ConvertAll Then MsgBox ImportBtn(Index).Caption & vbCrLf & MsgUpdDone, vbInformation, TitleUpdDone
''lblExit:
''    Set RsNew = Nothing
''    Exit Sub
''Eloop:
''    mSrl = 0
''    Do Until mOrderNo = Master.Fields("Order ID") And mChallanNo = Master.Fields("Transaction #")
''        mSrl = mSrl + 1
''        Master.MoveNext
''    Loop
''
''    ErrorCnt = ErrorCnt + mSrl
''    lblRecError(Index).Caption = ErrorCnt: lblRecError(Index).Refresh
''    ErrorGCN.Execute ("insert into prnmissrec(ID,code,colname,details) values(" & Master.AbsolutePosition & ",'" & "" & "','Stock Transfer (Inward)','" & Mid(Replace(err.Description, "'", "`"), 1, 250) & "')")
''    Resume Next

Sub UpdateTableStructure()
     'On Error Resume Next
        
    ErrorGCN.Execute "Alter Table Enviro Add State VarChar(1)"
    ErrorGCN.Execute "Alter Table Enviro Add VehicleCstPurchFormCode VarChar(8)"
    ErrorGCN.Execute "Alter Table AccountGroup Add AutomanGroupCodeWorkShop VarChar(15)"
    GCn.Execute "Alter Table Model Alter Column Cat_Code VarChar(5)"
    GCn.Execute "Alter Table Model_Cat Alter Column ModelCat_Code VarChar(5)"
End Sub


