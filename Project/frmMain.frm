VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Special Folder Location"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.FontBold = True
    Me.Print "Desktop"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_DESKTOP)
    
    Me.FontBold = True
    Me.Print "Favorites"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_FAVORITES)
    
    Me.FontBold = True
    Me.Print "Fonts"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_FONTS)
    
    Me.FontBold = True
    Me.Print "History"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_HISTORY)
    
    Me.FontBold = True
    Me.Print "My Documents"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_PERSONAL)
    
    Me.FontBold = True
    Me.Print "Program Files"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_PROGRAM_FILES)
    
    Me.FontBold = True
    Me.Print "Programs"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_PROGRAMS)
    
    Me.FontBold = True
    Me.Print "Recent"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_RECENT)
    
    Me.FontBold = True
    Me.Print "SendTo"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_SENDTO)
    
    Me.FontBold = True
    Me.Print "Start Menu"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_STARTMENU)
    
    Me.FontBold = True
    Me.Print "System"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_SYSTEM)
    
    Me.FontBold = True
    Me.Print "Templates"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_TEMPLATES)
    
    Me.FontBold = True
    Me.Print "Windows"
    Me.FontBold = False
    Me.Print vbTab & GetFolderPath(CSIDL_WINDOWS)
    
    ' try some them...
End Sub

