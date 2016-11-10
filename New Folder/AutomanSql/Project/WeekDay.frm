VERSION 5.00
Begin VB.Form WeekDay 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtDay 
      Height          =   420
      Left            =   2040
      TabIndex        =   3
      Top             =   1470
      Width           =   1890
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   1215
      TabIndex        =   2
      Top             =   2340
      Width           =   1965
   End
   Begin VB.TextBox TxtWeek 
      Height          =   420
      Left            =   2115
      TabIndex        =   1
      Top             =   930
      Width           =   1890
   End
   Begin VB.TextBox TxtMth 
      Height          =   285
      Left            =   2100
      TabIndex        =   0
      Top             =   360
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "Day"
      Height          =   480
      Left            =   225
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Week"
      Height          =   285
      Left            =   105
      TabIndex        =   5
      Top             =   1020
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Month"
      Height          =   330
      Left            =   240
      TabIndex        =   4
      Top             =   405
      Width           =   1245
   End
End
Attribute VB_Name = "WeekDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim mDate$, mDay As Byte, i As Byte
mDay = Val(TxtWeek) * 7
mDate = Format(mDay & "/" & Val(TxtMth) & "/" & Year(Date), "dd/mm/yyyy")
For i = 1 To 7
    If WeekDay(CDate(mDate)) = Val(TxtDay) Then
        MsgBox " New Date " & mDate & " Week Day " & WeekDay(CDate(mDate))
    Else
        mDate = CDate(mDate) - 1
    End If
Next
End Sub
