VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMastList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CFE0E0&
   Caption         =   "Master List"
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   11610
   DrawStyle       =   6  'Inside Solid
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11610
   Begin VB.Frame Frame2 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   840
      Left            =   1725
      TabIndex        =   22
      Top             =   2580
      Visible         =   0   'False
      Width           =   6180
      Begin VB.OptionButton Opt1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CFE0E0&
         Caption         =   "Particular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   270
         Width           =   1170
      End
      Begin VB.OptionButton Opt1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CFE0E0&
         Caption         =   "Wheel Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   1
         Left            =   4035
         TabIndex        =   8
         Top             =   240
         Width           =   1860
      End
      Begin VB.OptionButton Opt1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CFE0E0&
         Caption         =   "Model Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Index           =   0
         Left            =   285
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1860
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Scope"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   975
      Index           =   1
      Left            =   1725
      TabIndex        =   21
      Top             =   870
      Width           =   6180
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CFE0E0&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   210
         TabIndex        =   2
         Top             =   615
         Width           =   2115
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CFE0E0&
         Caption         =   "Modified Records"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Index           =   1
         Left            =   3495
         TabIndex        =   1
         Top             =   300
         Width           =   2115
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CFE0E0&
         Caption         =   "All Records"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   2115
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   11610
      TabIndex        =   19
      Top             =   6060
      Width           =   11610
      Begin VB.CommandButton BtnExit 
         BackColor       =   &H00D3BEC9&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4665
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exit Form"
         Top             =   30
         Width           =   2190
      End
      Begin VB.CommandButton BtnPrint 
         BackColor       =   &H00D3BEC9&
         Caption         =   "&Print"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2490
         MaskColor       =   &H00800080&
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Print Reports"
         Top             =   30
         Width           =   2190
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFE0E0&
      Caption         =   "Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   705
      Index           =   0
      Left            =   1725
      TabIndex        =   15
      Top             =   1875
      Width           =   6180
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   3
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   37018
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   315
         Index           =   1
         Left            =   4590
         TabIndex        =   5
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   37018
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Index           =   1
         Left            =   3570
         TabIndex        =   17
         Top             =   277
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   16
         Top             =   277
         Width           =   945
      End
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAAAAAAAAAAA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   -1140
      TabIndex        =   18
      Top             =   15
      Width           =   11760
   End
   Begin VB.Shape Shape1 
      Height          =   360
      Left            =   60
      Top             =   7665
      Width           =   11775
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Orientation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8265
      TabIndex        =   14
      Top             =   7740
      Width           =   960
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Portrait"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   9765
      TabIndex        =   13
      Top             =   7755
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4695
      TabIndex        =   12
      Top             =   7740
      Width           =   990
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "80/132 Columns"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   6135
      TabIndex        =   11
      Top             =   7755
      Width           =   1380
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Printer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2940
      TabIndex        =   10
      Top             =   7710
      Width           =   2595
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Default Printer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   1410
      TabIndex        =   4
      Top             =   7740
      Width           =   1245
   End
End
Attribute VB_Name = "frmMastList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents FGrid1 As MSFlexGridLib.MSFlexGrid
Attribute FGrid1.VB_VarHelpID = -1
Dim rsType As Recordset
Public g_FormID             As Byte
'Report Index
Private Const V_StateList       As Byte = 0          'State List
Private Const V_CityList        As Byte = 1          'City List
Private Const V_EmployeeList    As Byte = 2          'Employee List
Private Const V_VehicleCategory As Byte = 3          'vehicle model category master
Private Const V_VehicleModel    As Byte = 4          'vehicle model group master
Private Const V_Aggregate       As Byte = 5          'Aggregate master
Private Const V_Finance         As Byte = 6          'contract/finance
Private Const V_ContractOEM     As Byte = 66         'Contract/OEM
Private Const V_Discount        As Byte = 7          'discount factor master
Private Const V_Propart         As Byte = 8          'proprietory part grade master
Private Const V_Godown          As Byte = 9          'godown master
Private Const V_Unit            As Byte = 10         'unit master
Private Const V_Part            As Byte = 11         'part master
Private Const V_Model           As Byte = 12         'vehicle model master
Private Const V_Dealer          As Byte = 13         'dealer master
Private Const V_Colour          As Byte = 14         'colour master

'Object Index
'for Option Button
Private Const Rec_All       As Byte = 0 'All Records
Private Const Rec_Modified  As Byte = 1 'Modified Records
Private Const rec_option1   As Byte = 0 'option1(0)
Private Const rec_option2   As Byte = 1 'option1(1)
Private Const rec_option3   As Byte = 2 'option1(2)

Private Sub btnexit_Click()
'    If frmMain.Report1.Count  > 1 Then Unload frmMain.Report1(1)
    Set rsType = Nothing
    Unload Me
End Sub

Private Sub BTNPRINT_Click()
On Error GoTo lblErrorBox
Dim I As Integer, RepForm As New frmRepForm, NotModify As Boolean
Dim Rst As Recordset, SqlQry$, SelStr$, ForItem$
Dim RepFileName$, RepTitle$
 Dim sitecond As String
'Dim CrysRep As CrystalReport
Dim ac_str As String, ac_str1$, j As Integer
ac_str = ""
If DTP1(0).Value > DTP1(1).Value Then
    MsgBox " The Starting Date Should be less then End Date", vbInformation, "Date"
    Exit Sub
End If

Select Case g_FormID
    Case V_EmployeeList
        'Emp_Type = 0- >Sales Staff/ 1- >Mechanic /2- >Others
        'Designation
            If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
              sitecond = "where LEFT(e.site_code,1)='" & PubSiteCode & "'"
            Else
              sitecond = ""
            End If
    
        If PubBackEnd = "A" Then
            SqlQry = "select E.*, Switch(Emp_Type = 0,'Sales Staff',Emp_Type = 1,'Mechanic',Emp_Type = 2,'Others') As EmpType from Emp_Mast E " & sitecond & " Order by E.Emp_Name"
        ElseIf PubBackEnd = "S" Then
            SqlQry = "select E.*, CASE Emp_Type When  0 Then 'Sales Staff' When 1 Then 'Mechanic' When 2 Then 'Others' End As EmpType from Emp_Mast E " & sitecond & " Order by E.Emp_Name"
        End If
        
        If opt(Rec_Modified).Value = True Then
            NotModify = False
            SqlQry = SqlQry & "  where E.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and part_discfactor.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and part_discfactor.U_AE='E'"
            RepTitle = Me.CAPTION & "(Modified)"
        Else
            NotModify = True
            RepTitle = Me.CAPTION
        End If
        RepFileName = "EmpReg"
    
    Case V_StateList
        RepTitle = Me.CAPTION
        SqlQry = "SELECT StateCode, Site_Code,StateName,U_EntDt FROM State "
        If opt(Rec_All).Value = True Then
                NotModify = True
                SqlQry = SqlQry & " order by statename"
        Else
                SqlQry = SqlQry & " where U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and U_AE='E' order by statename "
                RepTitle = Me.CAPTION & "(Modified)"
                NotModify = False
        End If
        RepFileName = "StateList"
        
    Case V_CityList
    
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(city.site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
        If PubBackEnd = "A" Then
            SqlQry = "select citycode,cityname,switch(localcentral='C','Central',localcentral='L','Local') as LocalCentral,city.U_EntDt,state.statename " & _
                    "from City left join State on City.StateCode=State.StateCode " & sitecond & ""
        ElseIf PubBackEnd = "S" Then
            SqlQry = "select citycode,cityname,Case LocalCentral When 'C' Then 'Central' When 'L' Then 'Local' End as LocalCentral,city.U_EntDt,state.statename " & _
                    "from City left join State on City.StateCode=State.StateCode " & sitecond & ""
        End If
        RepTitle = Me.CAPTION
        If Opt1(rec_option2).Value = True Then
            For I = 1 To FGrid1.Rows - 1
                If FGrid1.TextMatrix(I, 0) = "ü" Then
                    ac_str = ac_str + IIf(ac_str = "", "'" + FGrid1.TextMatrix(I, 1) + "'", "," + "'" + FGrid1.TextMatrix(I, 1) + "'")
                End If
            Next I
            If ac_str = "" Then
                MsgBox " Select States ", vbInformation, "States"
                Exit Sub
            End If
            SqlQry = SqlQry & " where state.statename in (" & ac_str & ")"
        End If
        If opt(Rec_Modified).Value = True Then
            If Opt1(rec_option1).Value = True Then
                  SqlQry = SqlQry & " where city.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and city.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and city.U_AE='E' order by cityname "
            ElseIf Opt1(rec_option2).Value = True Then
                  SqlQry = SqlQry & " and  city.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and city.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and city.U_AE='E' order by cityname "
            End If
            RepTitle = Me.CAPTION & "(Modified)"
            NotModify = False
        Else
            SqlQry = SqlQry
            NotModify = True
        End If
        RepFileName = "CityList"
        
    Case V_VehicleCategory
        SqlQry = " select site_code,modelcat_code,modelcat_name,model_cat.u_entdt from model_cat"
        RepTitle = Me.CAPTION
        RepFileName = "vehiclecategory"
        If opt(Rec_Modified).Value = True Then
            SqlQry = SqlQry & " where model_cat.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and model_cat.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and model_cat.U_AE='E' order by modelcat_name "
            RepTitle = Me.CAPTION & "(Modified)"
            NotModify = False
        Else
            SqlQry = SqlQry & " order by modelcat_name "
            NotModify = True
        End If
        
    Case V_VehicleModel
        For I = 1 To FGrid1.Rows - 1
            If FGrid1.TextMatrix(I, 0) = "ü" Then
                ac_str = ac_str + IIf(ac_str = "", "'" + FGrid1.TextMatrix(I, 1) + "'", "," + "'" + FGrid1.TextMatrix(I, 1) + "'")
            End If
        Next I
        
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = " LEFT(model_grp.site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
        If Opt1(rec_option1).Value = True And ac_str = "" Then
                SqlQry = " select modelgrp_code, model_grp.modelcat_code,model_grp.u_entdt,modelgrp_name,wheel_catg,model_cat.modelcat_name,division.div_name from ((model_grp left join model_cat on model_grp.modelcat_code=model_cat.modelcat_code)left join division on left(model_grp.modelgrp_code,1)=division.div_code) where " & sitecond & " "
        ElseIf Opt1(rec_option1).Value = True Then
                SqlQry = " select modelgrp_code, model_grp.modelcat_code,model_grp.u_entdt,modelgrp_name,wheel_catg,model_cat.modelcat_name,division.div_name from ((model_grp left join model_cat on model_grp.modelcat_code=model_cat.modelcat_code)left join division on left(model_grp.modelgrp_code,1)=division.div_code) where model_cat.modelcat_name in (" & ac_str & ") and " & sitecond & ""
        ElseIf Opt1(rec_option2).Value = True And ac_str = "" Then
                SqlQry = "select modelgrp_code, model_grp.modelcat_code,modelgrp_name,model_grp.u_entdt,wheel_catg,model_cat.modelcat_name,division.div_name from ((model_grp left join model_cat on model_grp.modelcat_code=model_cat.modelcat_code)left join division on left(model_grp.modelgrp_code,1)=division.div_code) where " & sitecond & ""
        ElseIf Opt1(rec_option2).Value = True Then
                SqlQry = "select modelgrp_code, model_grp.modelcat_code,modelgrp_name,model_grp.u_entdt,wheel_catg,model_cat.modelcat_name,division.div_name from ((model_grp left join model_cat on model_grp.modelcat_code=model_cat.modelcat_code)left join division on left(model_grp.modelgrp_code,1)=division.div_code) where model_grp.wheel_catg in (" & ac_str & ") and  " & sitecond & " "
        End If
        
        If opt(Rec_Modified).Value = True Then
            NotModify = False
            If Opt1(rec_option1).Value = True And ac_str = "" Then
                    SqlQry = SqlQry & " where model_grp.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and model_grp.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and model_grp.U_AE='E' "
                    RepTitle = Me.CAPTION & "(Modified)"
                    RepFileName = "vehiclemodel"
            ElseIf Opt1(rec_option1).Value = True Then
                    SqlQry = SqlQry & " and model_grp.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and model_grp.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and model_grp.U_AE='E' "
                    RepTitle = Me.CAPTION & "(Modified)"
                    RepFileName = "vehiclemodel"
            ElseIf Opt1(rec_option2).Value = True And ac_str = "" Then
                    SqlQry = SqlQry & " where model_grp.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and model_grp.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and model_grp.U_AE='E'"
                    RepTitle = Me.CAPTION & "(Modified)"
                    RepFileName = "vehiclemodel1"
            ElseIf Opt1(rec_option2).Value = True Then
                    SqlQry = SqlQry & " and model_grp.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and model_grp.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and model_grp.U_AE='E'"
                    RepTitle = Me.CAPTION & "(Modified)"
                    RepFileName = "vehiclemodel1"
            End If
        Else
            NotModify = True
            If Opt1(rec_option1).Value = True Then
                    RepTitle = Me.CAPTION
                    RepFileName = "vehiclemodel"
            ElseIf Opt1(rec_option2).Value = True Then
                    RepTitle = Me.CAPTION
                    RepFileName = "vehiclemodel1"
            End If
        End If
        
    Case V_Aggregate
        If opt(Rec_Modified).Value = True And chk.Value = 1 Then
                SqlQry = " select aggre_code,aggre_name,site_code,aggreengine,aggregate.U_EntDt from aggregate where aggreengine='Y'"
                SqlQry = SqlQry & " and  aggregate.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and aggregate.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and aggregate.U_AE='E' order by aggre_name "
                RepTitle = Me.CAPTION & "(Modified)"
                NotModify = False
        ElseIf opt(Rec_All).Value = True And chk.Value = 1 Then
                SqlQry = " select aggre_code,aggre_name,site_code,aggreengine,aggregate.U_EntDt from aggregate where aggreengine='y' order by aggre_name "
                RepTitle = Me.CAPTION
                NotModify = True
        ElseIf opt(Rec_Modified).Value = True Then
                SqlQry = "select aggre_code,aggre_name,site_code,aggreengine,aggregate.U_EntDt from aggregate where aggregate.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and aggregate.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and aggregate.U_AE='E' order by aggre_name "
                RepTitle = Me.CAPTION & "(Modified)"
                NotModify = False
        ElseIf opt(Rec_All).Value = True Then
                SqlQry = " select aggre_code,aggre_name,site_code,aggreengine,aggregate.U_EntDt from aggregate order by aggre_name "
                RepTitle = Me.CAPTION
                NotModify = True
        End If
        RepFileName = "Aggregate"

    Case V_Finance, V_ContractOEM
        If Opt1(rec_option1).Value = False And Opt1(rec_option2).Value = False And Opt1(rec_option3).Value = False Then
            MsgBox "Select Contractor/Financier Category", vbInformation, " Contractor/Financier"
            Exit Sub
        End If
         If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
              sitecond = "and LEFT(cf.site_code,1)='" & PubSiteCode & "'"
            Else
              sitecond = ""
            End If
            
        GSQL = " select CF.U_EntDt,CF.FinCode,FinCatg,FinName,CF.Add1,CF.Add2,contactperson,city,CF.AcCode,ac_yn,CF.pincode,CF.phone,CF.fax,subgroup.name " & _
            " from (ContractFinance CF left join SubGroup on CF.accode=subgroup.SubCode) "
        If Opt1(rec_option2).Visible = True Then
            If Opt1(rec_option1).Value = True Then
                SqlQry = " where CF.FinCatg= 0"
            ElseIf Opt1(rec_option2).Value = True Then
                SqlQry = " where CF.FinCatg=1"
            ElseIf Opt1(rec_option3).Value = True Then
                SqlQry = " where CF.FinCatg=2"
            End If
        Else
            If Opt1(rec_option1).Value = True Then
                SqlQry = " where CF.FinCatg=1"
            ElseIf Opt1(rec_option3).Value = True Then
                SqlQry = " where CF.FinCatg=2"
            Else
                SqlQry = " where CF.FinCatg=1 or CF.FinCatg=2"
            End If
        End If
        SqlQry = GSQL & SqlQry & sitecond
        If opt(Rec_Modified).Value = True Then
            NotModify = False
            SqlQry = SqlQry & "and CF.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and CF.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and CF.U_AE='E'"
            If Opt1(rec_option1).Value = True Then
                RepTitle = " Contractor Register (Modified)"
            ElseIf Opt1(rec_option2).Value = True Then
                RepTitle = " Financier Register (Modified)"
            ElseIf Opt1(rec_option3).Value = True Then
                RepTitle = " OEM Register (Modified)"
            End If
        Else
            NotModify = True
            If Opt1(rec_option1).Value = True Then
                RepTitle = "Contractor Register"
            ElseIf Opt1(rec_option2).Value = True Then
                RepTitle = "Financier Register"
            ElseIf Opt1(rec_option3).Value = True Then
                RepTitle = "OEM Register"
            End If
        End If
        RepFileName = "Finance"
        
    Case V_Discount
     If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
              sitecond = "where LEFT(part_discfactor.site_code,1)='" & PubSiteCode & "'"
            Else
              sitecond = ""
            End If
    
    
        SqlQry = " select site_code,purcdisc_per,saldisc_per,discfac_catg,part_discfactor.U_EntDt from part_discfactor " & sitecond & " "
        If opt(Rec_Modified).Value = True Then
            NotModify = False
            SqlQry = SqlQry & "  and part_discfactor.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and part_discfactor.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and part_discfactor.U_AE='E'"
            RepTitle = Me.CAPTION & "(Modified)"
        Else
            NotModify = True
            RepTitle = Me.CAPTION
        End If
        RepFileName = "Discount"
        
    Case V_Propart
        SqlQry = " select partgrade_code,partgrade_name,part_grade.U_EntDt,CF.FinName from (part_grade left join ContractFinance CF on part_grade.oem_code=CF.FinCode)"
        If chk.Value = 1 Then
            If opt(Rec_All).Value = True Then
                NotModify = True
                SqlQry = SqlQry & " order by partgrade_name "
                RepFileName = "Propart"
                RepTitle = Me.CAPTION
            ElseIf opt(Rec_Modified).Value = True Then
                NotModify = False
                SqlQry = SqlQry & " where part_grade.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and part_grade.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and part_grade.U_AE='E'  order by partgrade_name "
                RepTitle = Me.CAPTION & "(Modified)"
                RepFileName = "Propart"
            End If
        Else
            If opt(Rec_All).Value = True Then
                NotModify = True
                SqlQry = SqlQry
                RepTitle = Me.CAPTION
                RepFileName = "Propart1"
            ElseIf opt(Rec_Modified).Value = True Then
                NotModify = False
                SqlQry = SqlQry & " where part_grade.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and part_grade.U_EntDt  <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and part_grade.U_AE='E'  order by partgrade_name "
                RepTitle = Me.CAPTION & "(Modified)"
                RepFileName = "Propart1"
            End If
        End If
    Case V_Godown
        SqlQry = " select god_code,site_code,god_name,godown.U_EntDt from godown"
        If opt(Rec_Modified).Value = True Then
            NotModify = False
            SqlQry = SqlQry & " where godown.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and godown.U_EntDt <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and godown.U_AE='E' order by god_name "
            RepTitle = Me.CAPTION & "(Modified)"
        Else
            SqlQry = SqlQry & " order by god_name "
            NotModify = True
            RepTitle = Me.CAPTION
        End If
        RepFileName = "Godown"
        
    Case V_Unit
        SqlQry = " select site_code,Unit_name,unit.U_EntDt from Unit"
        If opt(Rec_Modified).Value = True Then
            SqlQry = SqlQry & " where unit.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and unit.U_EntDt <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and unit.U_AE='E' order by unit_name "
            RepTitle = Me.CAPTION & "(Modified)"
            NotModify = False
        Else
            SqlQry = SqlQry & " order by unit_name "
            NotModify = True
            RepTitle = Me.CAPTION
        End If
        RepFileName = "Unit"
        
    Case V_Part
        SqlQry = " select Part.site_code,part.part_no,part.part_name,part.U_EntDt " & _
                 "from (Part left join Part_Grade on Part.Part_Grade=Part_grade.PartGrade_Code) Where div_code = '" & PubDivCode & "'"
         If Opt1(rec_option2).Value = True Then
            For I = 1 To FGrid1.Rows - 1
                If FGrid1.TextMatrix(I, 0) = "ü" Then
                    ac_str = ac_str + IIf(ac_str = "", "'" + FGrid1.TextMatrix(I, 1) + "'", "," + "'" + FGrid1.TextMatrix(I, 1) + "'")
                    ac_str1 = ac_str1 + IIf(ac_str1 = "", "" + FGrid1.TextMatrix(I, 1) + "", "," + "" + FGrid1.TextMatrix(I, 1) + "")
                End If
            Next I
            If ac_str = "" Then
                MsgBox " Select Grades ", vbInformation, "PartsGrades"
                Exit Sub
            End If
            SqlQry = SqlQry & " and  Part_Grade.PartGrade_Name in (" & ac_str & ")"
        End If
        If opt(Rec_Modified).Value = True Then
            NotModify = False
            SqlQry = SqlQry & "  and part.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and part.U_EntDt <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and part.U_AE='E' order by part_name "
            RepTitle = Me.CAPTION & "(Modified)"
        Else
            SqlQry = SqlQry & " order by part_name "
            NotModify = True
            RepTitle = Me.CAPTION
        End If
        
        RepFileName = "Part"
        
    Case V_Model
    
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
        SqlQry = "select * from model " & sitecond & " "
        RepTitle = Me.CAPTION
        If opt(Rec_All).Value = True Then
            SqlQry = SqlQry & " order by model "
            NotModify = True
        ElseIf opt(Rec_Modified).Value = True Then
            NotModify = False
            SqlQry = SqlQry & "  where model.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and model.U_EntDt <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and model.U_AE='E' order by model "
            RepTitle = Me.CAPTION & "(Modified)"
        End If
        RepFileName = "Model"
    
    Case V_Dealer
    If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
      sitecond = "where LEFT(AMD_DEALER.site_code,1)='" & PubSiteCode & "'"
    Else
      sitecond = ""
    End If
    
    
        SqlQry = " select D_code,D_NAME,D_ADD1,D_ADD2,D_ADD3,D_ADD4,D_CITY,D_PIN_CODE,D_DIST,D_CST_NO,D_RST_NO,U_EntDt from AMD_DEALER " & sitecond & " "
        RepTitle = Me.CAPTION
        If opt(Rec_All).Value = True Then
                SqlQry = SqlQry & " order by D_NAME "
                NotModify = True
        ElseIf opt(Rec_Modified).Value = True Then
                NotModify = False
                SqlQry = SqlQry & " where AMD_DEALER.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and AMD_DEALER.U_EntDt <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and AMD_DEALER.U_AE='E' order by D_NAME "
                RepTitle = Me.CAPTION & "(Modified)"
        End If
        RepFileName = "Dealer"
    
    Case V_Colour
          If PubSiteWiseDisplayYn = 1 And UCase(PubSiteType) <> "H" Then
              sitecond = "where LEFT(ColMast.site_code,1)='" & PubSiteCode & "'"
            Else
              sitecond = ""
            End If
            
        SqlQry = " Select Col_Code,Col_Desc,U_EntDt From ColMast " & sitecond & " "
        RepTitle = Me.CAPTION
        If opt(Rec_All).Value = True Then
                SqlQry = SqlQry & " Order By Col_Desc "
                NotModify = True
        ElseIf opt(Rec_Modified).Value = True Then
                NotModify = False
                SqlQry = SqlQry & " Where ColMast.U_EntDt >= " & ConvertDate(Format(DTP1(0).Value, "dd/mmm/yyyy")) & "  and colmast.U_EntDt <= " & ConvertDate(Format(DTP1(1).Value, "dd/mmm/yyyy")) & " and colmast.U_AE='E' order by col_desc "
                RepTitle = Me.CAPTION & "(Modified)"
        End If
        RepFileName = "color"
End Select
    
'******Check whether Report File Exists or not
'if file
'******Create Data Recordset
Set Rst = GCn.Execute(SqlQry)
If Rst.BOF = False Or Rst.EOF = False Then
    'Databse Connectivity Process
    CreateFieldDefFile Rst, PubRepoPath + "\" & RepFileName & ".TTX", True
    RepForm.Tag = RepTitle
    RepForm.CAPTION = "* " + RepTitle + " *"
    With RepForm.CrysReport1
        .Connect = ConnectStr
        .ReportFileName = PubRepoPath + "\" & RepFileName & ".RPT"
         Call Formula_Title(RepForm, RepTitle)
        If opt(Rec_Modified).Value = True Then
            .Formulas(4) = "DateBetween ='Modified during '+ '" & Format(DTP1(0).Value, "dd/mmm/yyyy") & "' + ' To ' + '" & Format(DTP1(1).Value, "dd/mmm/yyyy") & "'"
        Else
            .Formulas(4) = "DateBetween ='From :'+ '" & Format(DTP1(0).Value, "dd/mmm/yyyy") & "' + ' To ' + '" & Format(DTP1(1).Value, "dd/mmm/yyyy") & "'"
        End If
        .Formulas(5) = "Modify = " & NotModify & ""
'        .Formulas(6) = "ListOf = '" & ac_str1 & "'"
        .SetTablePrivateData 0, 3, Rst
        .Action = 1
    End With
Else
    MsgBox "No Records to Print", vbInformation, "Information"
    Exit Sub
End If
Set RepForm = Nothing
Set Rst = Nothing
Exit Sub

lblErrorBox:
    Set RepForm = Nothing
    Set Rst = Nothing
    ProcErrorMsg
End Sub

Private Sub FGrid1_Click()
Call FGrid_Click(FGrid1)
End Sub

Private Sub Opt_Click(Index As Integer)
Dim Rs As Recordset
Dim I As Byte
Select Case Index
    Case Rec_All 'All Records
'        For I = 0 To FGrid1.Rows - 1
'           FGrid1.TextMatrix(I, 0) = ""
'        Next
'         FGrid1.Enabled = False
            Frame1(0).Enabled = False
    Case Rec_Modified   'Modified Records
'         FGrid1.Enabled = True
          Frame1(0).Enabled = True
End Select
End Sub

Private Sub Opt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeysA vbKeyTab, True
End Sub

Private Sub Form_Load()
Dim I As Byte
Dim Rst As ADODB.Recordset
On Error GoTo lblErrorLoop
    Call WinSetting(Me)
    
    Set FGrid1 = Me.Controls.Add("MSFlexGridLib.MSFlexGrid", "FGrid1")
    Call Ini_Fgrid(FGrid1, 105, 2945, 2570, 3900, 0)
    IniGrid
    For I = 0 To Frame1.Count - 1
        Frame1(I).BackColor = Me.BackColor
    Next
    For I = 0 To opt.Count - 1
        opt(I).BackColor = Me.BackColor
    Next
    DTP1(0).Value = PubStartDate
    DTP1(1).Value = PubLoginDate
    opt(Rec_All).Value = True
'    Label6.Caption = Prin.DeviceName
    If opt(Rec_All).Value = True Then
        Frame1(0).Enabled = False
    End If
    With Frame1(1)
        .width = 7916.117
        .left = (Me.width - .width) / 2 ' 2228.81
        .top = 870 '2094.618
        .height = 1059.347
    End With
    With Frame1(0)
        .width = Frame1(1).width '7916.117
        .left = Frame1(1).left ' 2228.81
        .top = Frame1(1).top + Frame1(1).height + 90 ' 3226.194
        .height = 1131.575
    End With
    With Frame2
        .width = Frame1(1).width  ' 7916.117
        .left = Frame1(1).left  ' 2228.81
        .top = Frame1(0).top + Frame1(0).height + 90 '4405.921
        .height = 1000 '1348.26
    End With
    Set Rst = New ADODB.Recordset
    Select Case g_FormID
    
        Case V_StateList
                FGrid1.Visible = False
                chk.Visible = False
                
        Case V_CityList
'                Frame2.Height = 3772.757
                Frame2.Visible = True
                chk.Visible = False
                FGrid1.FormatString = "     |State Name                              |"
                Call FGridPosition
                Frame2.CAPTION = "Scope For States"
                Opt1(rec_option1).CAPTION = " All"
                Opt1(rec_option2).CAPTION = " Selected "
                Opt1(rec_option2).Visible = True
                Opt1(rec_option3).Visible = False
                Call FillGrid(FGrid1, "select statename,statecode from state order by statename")
                FGrid1.Enabled = False
                FGrid1.Refresh
                
        Case V_VehicleCategory
                FGrid1.Visible = False
                Frame2.Visible = False
                chk.Visible = False
        
        Case V_VehicleModel
                FGrid1.FormatString = "         |Model Category                                                        |"
                'Frame2.Height = 3772.757
                Call FGridPosition
                chk.Visible = False
                Frame2.Visible = True
                Frame2.CAPTION = " Scope For Vehicle Model"
                Opt1(rec_option1).CAPTION = "Model Category"
                Opt1(rec_option2).CAPTION = "Wheel Category"
                Opt1(rec_option3).Visible = False
                Opt1(rec_option1).Value = True
                Call FillGrid(FGrid1, "select modelcat_name,modelcat_code from model_cat order by modelcat_name")
                FGrid1.Refresh
                
        Case V_Aggregate
                chk.CAPTION = "Engine(Y/N)"
                Frame2.Visible = False
                FGrid1.Visible = False
                chk.Visible = True
'                Frame1(1).Height = 1564.945
'                Frame1(0).top = 3731.792
                
        Case V_Finance, V_ContractOEM
                chk.Visible = False
                Frame2.Visible = True
                Frame2.CAPTION = " Scope"
                Opt1(rec_option1).CAPTION = "Contractor"
                Opt1(rec_option2).CAPTION = "Financier"
                If g_FormID = V_ContractOEM Then
                    Opt1(rec_option2).Visible = False
                End If
                Opt1(rec_option3).CAPTION = "OEM"
                FGrid1.Visible = False
                Opt1(rec_option1).Value = False
                
        Case V_Discount, V_EmployeeList
                chk.Visible = False
                Frame2.Visible = False
                FGrid1.Visible = False
                
        Case V_Propart
                FGrid1.Visible = False
                chk.CAPTION = "With Supplier (Y/N)"
                chk.Visible = True
                Frame2.Visible = False
                Frame2.CAPTION = " Scope For Propietary Part"
                Opt1(rec_option1).CAPTION = "Grade Code"
                Opt1(rec_option2).CAPTION = "Grade Name "
                Opt1(rec_option3).Visible = False
'                Frame1(1).Height = 1564.945
'                Frame1(0).top = 3731.792

        Case V_Godown
                chk.Visible = False
                Frame2.Visible = False
                FGrid1.Visible = False
        
        Case V_Unit
                chk.Visible = False
                Frame2.Visible = False
                FGrid1.Visible = False
                       
        Case V_Part
'                chk.Visible = False
'                Frame2.Visible = False
'                FGrid1.Visible = False
                
'                Frame2.height = 3772.757
                Frame2.height = 700.757
                Frame2.Visible = True
                chk.Visible = False
                FGrid1.FormatString = "     |Grade Name                              |"
                Call FGridPosition
                Frame2.CAPTION = "Scope For Grades"
                Opt1(rec_option1).CAPTION = " All"
                Opt1(rec_option2).CAPTION = " Selected "
                Opt1(rec_option2).Visible = True
                Opt1(rec_option3).Visible = False
                Call FillGrid(FGrid1, "select partGrade_Name,PartGrade_Code from Part_grade order by PartGrade_Name")
                FGrid1.Enabled = False
                FGrid1.Refresh
         
        Case V_Model
                chk.Visible = False
                Frame2.Visible = False
                Frame2.CAPTION = " Scope For Parts"
                FGrid1.Visible = False
                
        Case V_Dealer
                FGrid1.Visible = False
                chk.Visible = False
                Frame2.Visible = False
                
        Case V_Colour
                FGrid1.Visible = False
                chk.Visible = False
                Frame2.Visible = False
               
    End Select
    Exit Sub
lblErrorLoop:
    MsgBox err.Description, vbInformation, "Information": Exit Sub
End Sub

Private Sub Opt1_Click(Index As Integer)
    
    Dim I As Integer
    Select Case g_FormID
        Case V_CityList
            If Opt1(rec_option1).Value = True Then
                FGrid1.Enabled = False
                For I = 1 To FGrid1.Rows - 1
                    FGrid1.TextMatrix(I, 0) = ""
                Next I
            ElseIf Opt1(rec_option2).Value = True Then
                FGrid1.Enabled = True
            End If
         Case V_Part
            If Opt1(rec_option1).Value = True Then
                FGrid1.Enabled = False
                For I = 1 To FGrid1.Rows - 1
                    FGrid1.TextMatrix(I, 0) = ""
                Next I
            ElseIf Opt1(rec_option2).Value = True Then
                FGrid1.Enabled = True
            End If
        Case V_VehicleModel
           If Opt1(rec_option1).Value = True Then
                FGrid1.FormatString = "         |Model Category                                                        |"
                Call FillGrid(FGrid1, "SELECT model_cat.modelcat_name,model_cat.modelcat_code from model_cat order by model_cat.modelcat_name")
                FGrid1.Refresh
                Call FGridPosition
            ElseIf Opt1(rec_option2).Value = True Then
                FGrid1.FormatString = "         |Wheel Category                                                                            |"
'                Call FillGrid(FGrid1, "select distinct wheel_catg,modelgrp_name from model_grp order by modelgrp_name")
                Call FillGrid(FGrid1, "select distinct wheel_catg,wheel_catg as xx  from model_grp")
                FGrid1.Refresh
                Call FGridPosition
            End If
    End Select
End Sub
Private Sub IniGrid()
    With FGrid1
        .ColWidth(2) = 0
        .BackColor = &H80000005
        .BackColorBkg = &HCFE0E0
        .BackColorFixed = &HBFCEC7
        .BackColorSel = &H8000000D
        .ForeColor = &H80000008
        .ForeColorFixed = &H80000012
        .ForeColorSel = &H8000000E
        .GridColor = &HC0C0C0
        .GridColorFixed = &H0&
        .GridLineWidth = 1
        .FocusRect = flexFocusNone
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .BorderStyle = flexBorderSingle
        .TextStyle = flexTextFlat
        .TextStyleFixed = flexTextFlat
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
    End With
End Sub
Public Sub FGridPosition()
    FGrid1.top = Frame2.top + Frame2.height
    FGrid1.ColWidth(0) = 500
    FGrid1.ColWidth(1) = 5400
    FGrid1.ColWidth(2) = 0
    FGrid1.height = Me.height - (FGrid1.top + Picture1.height + 460)
    FGrid1.left = Frame2.left + 20
    FGrid1.width = Frame2.width - 30
End Sub
