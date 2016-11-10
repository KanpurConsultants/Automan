VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSiteSelection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Site Selection"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSDataGridLib.DataGrid DgSite 
      Height          =   2505
      Left            =   225
      TabIndex        =   0
      Top             =   900
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4419
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
         Caption         =   "Site"
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
         BeginProperty Column00 
            ColumnWidth     =   3960
         EndProperty
      EndProperty
   End
   Begin VB.Label LblSiteHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      Site Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   225
      TabIndex        =   1
      Top             =   615
      Width           =   4695
   End
End
Attribute VB_Name = "FrmSiteSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsMast As ADODB.Recordset



Sub Initialise_Pub()
    PubSiteCode = RsMast!Code
    PubSiteName = RsMast!Name
    If XNull(RsMast!Address1) <> "" Then
        PubComp_Add = XNull(RsMast!Address1)
        PubComp_Add2 = XNull(RsMast!Address2)
        PubComp_Add3 = XNull(RsMast!Address3)
        PubComp_City = XNull(RsMast!City)
        PubComp_Contact = "PHONE : " & XNull(RsMast!Phone) & " MOBILE   : " & XNull(RsMast!Mobile)
    End If
End Sub


Private Sub DgSite_DblClick()
    Call Initialise_Pub
    Me.Hide
    MDIForm1.Show
End Sub

Private Sub DgSite_KeyPress(keyascii As Integer)
    If keyascii = 10 Or keyascii = 13 Then
        Call Initialise_Pub
        Me.Hide
        MDIForm1.Show
        MDIForm1.ZOrder 0
    End If
End Sub

Private Sub Form_Load()
Dim Condstr$
    FrmSiteSelection.CAPTION = PubPackage & "-[" & PubCenCompCode & "]" & PubComp_Name
    LblCompany = "Sites/Branches Of " & PubComp_Name & " - " & "[" & PubCenCompCode & "]"
    If pubUName <> "SA" Then
        Set RsMast = GCn.Execute("Select S.Site_Code As Code, S.Site_Desc As Name, S.Address1, S.Address2, S.Address3, S.City From User_Site US Left Join Site S On US.Site_Code=S.Site_Code Where US.User_Name = '" & pubUName & "' Order By S.Site_Desc")
    Else
        Set RsMast = GCn.Execute("Select S.Site_Code As Code, S.Site_Desc As Name, S.Address1, S.Address2, S.Address3, S.City From Site S Order By S.Site_Desc")
    End If
    Set DgSite.DataSource = RsMast
    If RsMast.RecordCount > 0 Then
        RsMast.FIND "Code='" & PubSiteCode & "'"
        If RsMast.EOF = True Then RsMast.MoveFirst
    Else
        MsgBox "Sorry! You Have No Permission Of Any Site to Login"
        End
    End If
End Sub




Private Sub Form_Unload(Cancel As Integer)
    Set RsMast = Nothing
End Sub
