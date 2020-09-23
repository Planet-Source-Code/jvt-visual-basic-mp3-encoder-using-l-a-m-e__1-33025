VERSION 5.00
Begin VB.Form frmHelp 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    WriteHelp
End Sub
Sub drawback()
    Dim i As Integer
    Dim c As Integer
    c = 50
    For i = 0 To 15
        Me.Line (0, i)-(Me.ScaleWidth, i), RGB(c, c, c)
        c = c + (i * 5 / 2)
    Next i
    c = 0
    For i = Me.ScaleHeight To Me.ScaleHeight - 15 Step -1
        Me.Line (0, i)-(Me.ScaleWidth, i), RGB(c, c, c)
        Debug.Print Me.ScaleHeight - i
        c = c + ((Me.ScaleHeight - i) * 5 / 2)
    Next i
End Sub
Sub WriteHelp()
    Me.Cls
    drawback
    Me.CurrentY = 15
    Me.CurrentX = 15
    Me.FontBold = True
    Me.Font = "verdana"
    Me.FontSize = 12
    Me.ForeColor = QBColor(4)
    Me.Print "Quick Reference"
    Me.FontSize = 8
    Me.ForeColor = QBColor(0)
    Me.Line (15, 35)-(Me.ScaleWidth - 15, 35)
    Me.Print
    Me.FontBold = True
    '48
    Me.CurrentX = 15: Me.Print "File"
    Me.CurrentX = 15: Me.Print "Output"
    Me.CurrentX = 15: Me.Print "Bitrate"
    Me.CurrentX = 15: Me.Print "Frequency"
    Me.CurrentX = 15: Me.Print "Mode"
    Me.Print
    Me.CurrentX = 15: Me.Print "Use VBR"
    Me.CurrentX = 25: Me.Print "Max Bitrate"
    Me.CurrentX = 25: Me.Print "Quality"
    Me.Print
    Me.CurrentX = 15: Me.Print "Insert ID3 Tag"
    '97
    Me.FontBold = False
    Me.CurrentY = 48
    Me.CurrentX = 120: Me.Print "WAVE file to be encoded."
    Me.CurrentX = 120: Me.Print "MP3 file to be saved."
    Me.CurrentX = 120: Me.Print "Quality of the saved MP3 file."
    Me.CurrentX = 120: Me.Print "Sampling rate of the saved MP3 file."
    Me.CurrentX = 120: Me.Print "Sets encoding mode. (Stereo, Mono, etc)"
    Me.Print
    Me.CurrentX = 120: Me.Print "Enables/Disables Variable Bitrate Encoding."
    Me.CurrentX = 130: Me.Print "The Maximum Bitrate when using VBR."
    Me.CurrentX = 130: Me.Print "The quality of VBR encoding."
    Me.Print
    Me.CurrentX = 120: Me.Print "Enables/Disables saving of ID3 tag."
    Me.Print
    Me.Print
    Me.Print
    Dim CT As String
    Me.FontSize = 7
    CT = App.Title & " Version " & App.Major & "." & App.Minor & "." & App.Revision & " by Jayvee Tensuan © 2002"
    Me.CurrentX = (Me.ScaleWidth / 2) - (Me.TextWidth(CT) / 2)
    Me.Print CT
       
    CT = "This program uses L.A.M.E. MP3 encoding technology."
    Me.CurrentX = (Me.ScaleWidth / 2) - (Me.TextWidth(CT) / 2)
    Me.Print CT
    
    Me.ForeColor = QBColor(1)
    CT = "(http://www.mp3dev.org/)"
    Me.CurrentX = (Me.ScaleWidth / 2) - (Me.TextWidth(CT) / 2)
    Me.Print CT
End Sub
