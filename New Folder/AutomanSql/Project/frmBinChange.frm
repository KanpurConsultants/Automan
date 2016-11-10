VERSION 5.00
Begin VB.Form frmBinChange 
   BackColor       =   &H00CFE0E0&
   Caption         =   "Change Bin Location"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFE0E0&
      Height          =   2895
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   5265
      Begin VB.TextBox Part_Name 
         Appearance      =   0  'Flat
         BackColor       =   &H00CFE0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1680
         MaxLength       =   45
         TabIndex        =   1
         Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         Top             =   855
         Width           =   3480
      End
      Begin VB.CommandButton CmdChange 
         Caption         =   "Change"
         Height          =   360
         Left            =   1695
         TabIndex        =   4
         Top             =   2280
         Width           =   1860
      End
      Begin VB.TextBox New_Bin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1695
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1710
         Width           =   2325
      End
      Begin VB.TextBox Curr_Bin 
         Appearance      =   0  'Flat
         BackColor       =   &H00CFE0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1695
         MaxLength       =   25
         TabIndex        =   2
         Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         Top             =   1245
         Width           =   3435
      End
      Begin VB.TextBox Part_No 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         MaxLength       =   22
         TabIndex        =   0
         Top             =   390
         Width           =   3435
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Part Name"
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
         Left            =   165
         TabIndex        =   9
         Top             =   870
         Width           =   1260
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "New Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   165
         TabIndex        =   8
         Top             =   1740
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Location"
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
         Left            =   135
         TabIndex        =   7
         Top             =   1275
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Part No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   165
         TabIndex        =   6
         Top             =   420
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmBinChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPart As ADODB.Recordset
Private Sub CmdChange_Click()
    If IsValid(Part_No, "Part No") = False Then Exit Sub
    GCn.Execute ("Update Part set Bin_Loca='" & New_Bin & "' where PART_No = '" & Part_No & "' AND div_code ='" & PubDivCode & "'")
    Part_No = ""
    Part_Name = "No Name Selected"
    Curr_Bin = "No Location Selected"
    New_Bin = ""
    Part_No.SetFocus
End Sub

Private Sub Form_Activate()
    Part_Name = "No Name Selected"
    Curr_Bin = "No Location Selected"
    Part_No.SetFocus
End Sub
Private Sub New_Bin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    If MsgBox("Save Data ? ", vbInformation + vbYesNo, "Save Message") = vbYes Then
        CmdChange_Click
    Else
        New_Bin.SetFocus
        Exit Sub
    End If
ElseIf KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Part_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    If IsValid(Part_No, "Part No") = False Then Exit Sub
    Set RsPart = GCn.Execute("SELECT Part_Name,Bin_Loca FROM Part where PART_NO = '" & Part_No & "' AND div_code ='" & PubDivCode & "'")
    
    If RsPart.RecordCount = 0 Then
        MsgBox "No Record Found  Of Given Part", vbInformation, "Search Result"
        Part_No = ""
        Part_No.SetFocus
        Exit Sub
    Else
        Part_Name = RsPart!Part_Name
        Curr_Bin = RsPart!Bin_Loca
        New_Bin.SetFocus
    End If
ElseIf KeyCode = vbKeyEscape Then
    Unload Me
End If

End Sub

