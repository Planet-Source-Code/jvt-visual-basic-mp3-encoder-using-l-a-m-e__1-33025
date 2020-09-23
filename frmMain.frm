VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3 Encoder"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmbEncode 
      Caption         =   "Start Encoding"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   6000
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   5520
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   930
      Left            =   5160
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   870
      ScaleWidth      =   675
      TabIndex        =   20
      Top             =   5400
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Encoding Parameters"
      Height          =   3135
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   5655
      Begin VB.CommandButton cmdEditID3 
         Caption         =   "Edit ID3"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   27
         Top             =   2560
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Insert ID3 Tag"
         Height          =   195
         Left            =   480
         TabIndex        =   26
         Top             =   2640
         Width           =   1335
      End
      Begin MSComctlLib.Slider vbrQual 
         Height          =   495
         Left            =   3720
         TabIndex        =   9
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         Max             =   9
         SelStart        =   5
         Value           =   5
      End
      Begin VB.ComboBox cmbMaxBitRate 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Sets the bitrate of the output file. The higher the bitrate, the better the quality, the larger the size"
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use VBR"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cmbMode 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Sets the encoding mode (Stereo, Mono, etc)"
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cmbFreq 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Sets the frequency of the output file. Set to Auto if you do not know the original frequency of the input file"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cmbBitrate 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Sets the bitrate of the output file. The higher the bitrate, the better the quality, the larger the size"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Quality:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   1965
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Max. Bitrate:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   1965
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "kbits/sec"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   1965
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Mode:"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   1000
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Hz"
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   400
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Frequency:"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   405
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "kbits/sec"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   400
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Bitrate:"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   405
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Settings"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File:"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Output:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1245
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5400
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      Height          =   255
      Left            =   4680
      TabIndex        =   22
      Top             =   5535
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Label8.Enabled = True
        Label9.Enabled = True
        cmbMaxBitRate.Enabled = True
        Label10.Enabled = True
        vbrQual.Enabled = True
    Else
        Label8.Enabled = False
        Label9.Enabled = False
        cmbMaxBitRate.Enabled = False
        Label10.Enabled = False
        vbrQual.Enabled = False
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        cmdEditID3.Enabled = True
    Else
        cmdEditID3.Enabled = False
    End If
End Sub

Private Sub cmbEncode_Click()
    On Local Error Resume Next
    Dim Mode As Byte
    Dim ret As EncodingErrors
    Dim SampleRate As Long
    If cmbMode.Text = "Stereo" Then Mode = 0
    If cmbMode.Text = "Joint Stereo" Then Mode = 1
    If cmbMode.Text = "Dual Channel" Then Mode = 2
    If cmbMode.Text = "Mono" Then Mode = 3
    
    cmbEncode.Enabled = False
    cmdCancel.Enabled = True
    SampleRate = Val(cmbFreq.Text)
    Err.Clear
    If cmbFreq.ListIndex = 0 Then
        Open txtFile.Text For Binary Access Read As #1
            If Err Then
                MsgBox "Could not open input file!", vbExclamation + vbOKOnly, App.Title
                Exit Sub
            End If
            Get #1, 25, SampleRate
        Close #1
        If SampleRate <> 32000 And SampleRate <> 44100 And SampleRate <> 48000 Then
            If cmbBitrate.Text <> "128" Then
                MsgBox "The detected sample rate (" & (SampleRate / 1000) & " KHz) is not fully supported by the encoder. For higher compatibility, the bitrate will be automatically set to 128 kbits/sec. There is a possibility that encoding will not succeed.", vbOKOnly + vbInformation
                cmbBitrate.ListIndex = 8
            End If
        End If
        Debug.Print SampleRate
    End If
    If Check1.Value = 1 Then
        
        If SampleRate <> 32000 And SampleRate <> 44100 And SampleRate <> 48000 Then
            MsgBox "The detected sample rate (" & (SampleRate / 1000) & " KHz) is not supported by VBR! Please choose between 32000, 44100, or 48000.", vbOKOnly + vbExclamation
            cmbEncode.Enabled = True
            cmdCancel.Enabled = False
            Exit Sub
        End If
        
        ret = SetVBR(1, vbrQual.Value, VBR_METHOD_DEFAULT, Val(cmbMaxBitRate.Text))
    End If
    
    ret = EncodeMp3(txtFile.Text, txtOutput.Text, Val(cmbBitrate.Text), SampleRate, Mode, AddressOf EnumEncoding)
    If ret = ENC_ERR_ENCODING_FAILED Then
        MsgBox "MP3 encoding failed", vbExclamation + vbOKOnly
        cmbEncode.Enabled = True
        cmdCancel.Enabled = False
    ElseIf ret = ENC_ERR_ENCODING_CANCELLED Then
        MsgBox "Encoding cancelled by user", vbExclamation, App.Title
        pb1.Value = 0
        lblPercent.Caption = "0%"
    ElseIf ret = ENC_ERR_NO_API Then
        MsgBox "Could not get encoding API!", vbExclamation + vbOKOnly, App.Title
        cmbEncode.Enabled = True
        cmdCancel.Enabled = False
    ElseIf ret = ENC_ERR_INPUT Then
        MsgBox "Could not open input file!", vbExclamation + vbOKOnly, App.Title
        cmbEncode.Enabled = True
        cmdCancel.Enabled = False
    ElseIf ret = ENC_ERR_OUTPUT Then
        MsgBox "Could not open output file!", vbExclamation + vbOKOnly, App.Title
        cmbEncode.Enabled = True
        cmdCancel.Enabled = False
    ElseIf ret = ENC_ERR_INVALID_PARAMS Then
        MsgBox "Invalid Encoding Parameters", vbExclamation + vbOKOnly, App.Title
        cmbEncode.Enabled = True
        cmdCancel.Enabled = False
    ElseIf ret = ENC_ERR_ENCODING_SUCCESS Then
        MsgBox "MP3 encoding done", vbInformation + vbOKOnly
        cmbEncode.Enabled = True
        cmdCancel.Enabled = False
        pb1.Value = 0
        lblPercent.Caption = "0%"
        If Check2.Value = 1 Then
            TID3.FileName = txtOutput.Text
            TID3.Load
            TID3.Album = ID3.Album
            TID3.Name = ID3.Name
            TID3.Artist = ID3.Artist
            TID3.Genre = ID3.Genre
            TID3.Comment = ID3.Comment
            TID3.Year = ID3.Year
            TID3.SongNumber = ID3.SongNumber
            TID3.Save
        End If
    Else
        MsgBox "Encoding failed", vbExclamation, App.Title
        cmbEncode.Enabled = True
        cmdCancel.Enabled = False
    End If
End Sub

Private Sub cmdBrowse_Click()
    With cd1
        .Filter = "Wave Files|*.wav|All Files|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
    End With
    txtFile.Text = cd1.FileName
    txtOutput.Text = Left(cd1.FileName, Len(cd1.FileName) - 4) + ".mp3"
End Sub

Private Sub cmdCancel_Click()
    EncCancel = True
    cmdCancel.Enabled = False
    cmbEncode.Enabled = True
End Sub

Private Sub cmdEditID3_Click()
    frmID3.Show 1, Me
End Sub

Private Sub cmdHelp_Click()
    frmHelp.Show 0, Me
    
End Sub

Private Sub Form_Load()
    cmbBitrate.Clear
    cmbBitrate.AddItem "32"
    cmbBitrate.AddItem "40"
    cmbBitrate.AddItem "48"
    cmbBitrate.AddItem "56"
    cmbBitrate.AddItem "64"
    cmbBitrate.AddItem "80"
    cmbBitrate.AddItem "96"
    cmbBitrate.AddItem "112"
    cmbBitrate.AddItem "128"
    cmbBitrate.AddItem "160"
    cmbBitrate.AddItem "192"
    cmbBitrate.AddItem "224"
    cmbBitrate.AddItem "256"
    cmbBitrate.AddItem "320"
    cmbBitrate.ListIndex = 8
    
    cmbMaxBitRate.Clear
    cmbMaxBitRate.AddItem "32"
    cmbMaxBitRate.AddItem "40"
    cmbMaxBitRate.AddItem "48"
    cmbMaxBitRate.AddItem "56"
    cmbMaxBitRate.AddItem "64"
    cmbMaxBitRate.AddItem "80"
    cmbMaxBitRate.AddItem "96"
    cmbMaxBitRate.AddItem "112"
    cmbMaxBitRate.AddItem "128"
    cmbMaxBitRate.AddItem "160"
    cmbMaxBitRate.AddItem "192"
    cmbMaxBitRate.AddItem "224"
    cmbMaxBitRate.AddItem "256"
    cmbMaxBitRate.AddItem "320"
    cmbMaxBitRate.ListIndex = 13
    
    cmbFreq.AddItem "Auto"
    cmbFreq.AddItem "32000"
    cmbFreq.AddItem "44100"
    cmbFreq.AddItem "48000"
    cmbFreq.ListIndex = 0
    
    cmbMode.AddItem "Stereo"
    cmbMode.AddItem "Joint Stereo"
    cmbMode.AddItem "Dual Channel"
    cmbMode.AddItem "Mono"
    cmbMode.ListIndex = 0
    
    'SetEncoder ENC_LAME
    
End Sub

Private Sub Picture1_Click()
    Dim Txt As String
    Dim el As String
    el = Chr(13) + Chr(10)
    Txt = "This Program uses technology from: " + el + el
    Txt = Txt + "Lame MP3 Encoder" + el
    Txt = Txt + "Lame Ain't an MP3 Encoder" + el
    Txt = Txt + "Site: http://www.mp3dev.org/" + el + el

    MsgBox Txt, vbInformation, "About MP3 Lame Encoder"
End Sub
