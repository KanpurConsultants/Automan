Attribute VB_Name = "MyReso"
Public Xtwips As Integer, Ytwips As Integer
Public Xpixels As Integer, Ypixels As Integer

Type FRMSIZE
   height As Long
   width As Long
End Type

Public RePosForm As Boolean
Public DoResize As Boolean


Dim myForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer




Sub Resize_For_Resolution(ByVal SFX As Single, ByVal SFY As Single, myForm As Form)
    Dim I As Integer
    Dim j As Integer
    Dim SFFont As Single
    Dim FormCtrl As Boolean
    SFFont = (SFX + SFY) / 2
    
    On Error Resume Next
       
    
    With myForm
     For I = 0 To .Count - 1
     
     
       If TypeOf .Controls(I) Is ComboBox Then   ' cannot change Height
         .Controls(I).left = .Controls(I).left * SFX
         .Controls(I).top = .Controls(I).top * SFY
         .Controls(I).width = .Controls(I).width * SFX
         .Controls(I).GridLineWidth = .Controls(I).GridLineWidth * SFX
       ElseIf TypeOf .Controls(I) Is MSHFlexGrid Then
            .Controls(I).FontFixed.Size = .Controls(I).FontFixed.Size * SFFont
         .Controls(I).Move .Controls(I).left * SFX, _
          .Controls(I).top * SFY, _
          .Controls(I).width * SFX, _
          .Controls(I).height * SFY
       ElseIf TypeOf .Controls(I) Is Line Then
        .Controls(I).X1 = .Controls(I).X1 * SFX
        .Controls(I).X2 = .Controls(I).X2 * SFX
        .Controls(I).Y1 = .Controls(I).Y1 * SFY
        .Controls(I).Y2 = .Controls(I).Y2 * SFY
       Else
         .Controls(I).Move .Controls(I).left * SFX, _
          .Controls(I).top * SFY, _
          .Controls(I).width * SFX, _
          .Controls(I).height * SFY
          
          If TypeOf .Controls(I) Is DataGrid Then
                For j = 0 To .Controls(I).Columns.Count - 1
                    .Controls(I).Columns(j).width = .Controls(I).Columns(j).width * SFFont
                Next j
          End If
          
       End If
         'If Not TypeOf .Controls(I) Is MSHFlexGrid Then
            .Controls(I).Font.Size = .Controls(I).Font.Size * SFFont
         'End If
         
         If TypeOf .Controls(I) Is DataGrid Then
            .Controls(I).HeadFont.Size = .Controls(I).HeadFont.Size * SFFont
         End If
      Next I
      
      If RePosForm Then
        .Move .left * SFX, .top * SFY, .width * SFX, .height * SFY
      End If
      
    End With
End Sub

Public Function SetResolutionFormLoad(Form1 As Object) As Single

Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
' Size of Form in Pixels at design resolution
DesignX = 800
DesignY = 600
RePosForm = True   ' Flag for positioning Form
DoResize = False   ' Flag for Resize Event
' Set up the screen values
Xtwips = Screen.TwipsPerPixelX
Ytwips = Screen.TwipsPerPixelY

Ypixels = Screen.height / Ytwips ' Y Pixel Resolution
Xpixels = Screen.width / Xtwips  ' X Pixel Resolution

' Determine scaling factors
ScaleFactorX = (Xpixels / DesignX)
ScaleFactorY = (Ypixels / DesignY)
ScaleMode = 1  ' twips
'Exit Sub  ' uncomForm1nt to see how Form1 looks without resizing
Resize_For_Resolution ScaleFactorX, ScaleFactorY, Form1
myForm.height = Form1.height ' ReForm1mber the current size
myForm.width = Form1.width
SetResolutionFormLoad = (ScaleFactorX + ScaleFactorY) / 2
End Function
