VERSION 5.00
Begin VB.Form frmID3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit ID3 Tag"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmID3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTrackNo 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox txtAlbum 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtComments 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
   End
   Begin VB.ComboBox cmbGenre 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Track #:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2560
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Album:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   765
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Comments:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2205
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Genre:"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   1725
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Year:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1725
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Artist:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1245
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Title:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   280
      Width           =   495
   End
End
Attribute VB_Name = "frmID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    ID3.Artist = txtArtist.Text
    ID3.Album = txtAlbum.Text
    ID3.Name = txtTitle.Text
    ID3.Year = txtYear.Text
    ID3.Comment = txtComments.Text
    ID3.Genre = cmbGenre.ListIndex
    ID3.SongNumber = txtTrackNo.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To 125
        ID3.Genre = i
        cmbGenre.AddItem ID3.Genre_str
    Next i
    cmbGenre.ListIndex = 0
End Sub
