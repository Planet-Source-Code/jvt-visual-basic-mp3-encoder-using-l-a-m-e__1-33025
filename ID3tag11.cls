VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ID3tag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Author: Alvaro Redondo <merlin@sevillaonline.com>
' http://www.sevillaonline.com/ActiveX
' Version: 1.0
' Last modification date: 30.09.2000
' Description: ID3 tags are used to store information in multimedia files, commonly MP3 files.
'               This class manages version 1.0 and 1.1 ID3 tags.

Option Explicit

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Const OPEN_EXISTING = 3

Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_BEGIN = 0

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
    (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, _
    ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, _
    ByVal dwMoveMethod As Long) As Long

Private Declare Function SetEndOfFile Lib "kernel32" _
    (ByVal hFile As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Enum enGenres
    [Blues] = 0
    [Classic Rock] = 1
    [Country] = 2
    [Dance] = 3
    [Disco] = 4
    [Funk] = 5
    [Grunge] = 6
    [Hip-Hop] = 7
    [Jazz] = 8
    [Metal] = 9
    [New Age] = 10
    [Oldies] = 11
    [Other] = 12
    [Pop] = 13
    [R&B] = 14
    [Rap] = 15
    [Reggae] = 16
    [Rock] = 17
    [Techno] = 18
    [Industrial] = 19
    [Alternative] = 20
    [Ska] = 21
    [Death Metal] = 22
    [Pranks] = 23
    [Soundtrack] = 24
    [Euro-Techno] = 25
    [Ambient] = 26
    [Trip-Hop] = 27
    [Vocal] = 28
    [Jazz+Funk] = 29
    [Fusion] = 30
    [Trance] = 31
    [Classical] = 32
    [Instrumental] = 33
    [Acid] = 34
    [House] = 35
    [Game] = 36
    [Sound Clip] = 37
    [Gospel] = 38
    [Noise] = 39
    [AlternRock] = 40
    [Bass] = 41
    [Soul] = 42
    [Punk] = 43
    [Space] = 44
    [Meditative] = 45
    [Instrumental Pop] = 46
    [Instrumental Rock] = 47
    [Ethnic] = 48
    [Gothic] = 49
    [Darkwave] = 50
    [Techno-Industrial] = 51
    [Electronic] = 52
    [Pop-Folk] = 53
    [Eurodance] = 54
    [Dream] = 55
    [Southern Rock] = 56
    [Comedy] = 57
    [Cult] = 58
    [Gangsta] = 59
    [Top 40] = 60
    [Christian Rap] = 61
    [Pop/Funk] = 62
    [Jungle] = 63
    [Native American] = 64
    [Cabaret] = 65
    [New Wave] = 66
    [Psychadelic] = 67
    [Rave] = 68
    [Showtunes] = 69
    [Trailer] = 70
    [Lo-Fi] = 71
    [Tribal] = 72
    [Acid Punk] = 73
    [Acid Jazz] = 74
    [Polka] = 75
    [Retro] = 76
    [Musical] = 77
    [Rock & Roll] = 78
    [Hard Rock] = 79
    [Folk] = 80
    [Folk-Rock] = 81
    [National Folk] = 82
    [Swing] = 83
    [Fast Fusion] = 84
    [Bebob] = 85
    [Latin] = 86
    [Revival] = 87
    [Celtic] = 88
    [Bluegrass] = 89
    [Avantgarde] = 90
    [Gothic Rock] = 91
    [Progressive Rock] = 92
    [Psychedelic Rock] = 93
    [Symphonic Rock] = 94
    [Slow Rock] = 95
    [Big Band] = 96
    [Chorus] = 97
    [Easy Listening] = 98
    [Acoustic] = 99
    [Humour] = 100
    [Speech] = 101
    [Chanson] = 102
    [Opera] = 103
    [Chamber Music] = 104
    [Sonata] = 105
    [Symphony] = 106
    [Booty Bass] = 107
    [Primus] = 108
    [Porn Groove] = 109
    [Satire] = 110
    [Slow Jam] = 111
    [Club] = 112
    [Tango] = 113
    [Samba] = 114
    [Folklore] = 115
    [Ballad] = 116
    [Power Ballad] = 117
    [Rhythmic Soul] = 118
    [Freestyle] = 119
    [Duet] = 120
    [Punk Rock] = 121
    [Drum Solo] = 122
    [A capella] = 123
    [Euro-House] = 124
    [Dance Hall] = 125
End Enum

Private Type tipoTag
    Tag As String * 3
    Name As String * 30
    Artist As String * 30
    Album As String * 30
    Year As String * 4
    Comment As String * 30
    Genre As String * 1
End Type

Private m_sName As String, m_sAlbum As String, m_sArtist As String, _
    m_iYear As Integer, m_iGenre As Integer, m_sComment As String, _
    m_iSongNumber As Integer

Private m_sPath As String, m_bFileExists As Boolean, m_bReadOnly As Boolean, _
    m_bHasID3tag As Boolean

Property Get Album() As String
Attribute Album.VB_Description = "Sets / returns album name."
    Album = m_sAlbum
End Property

Property Let Album(sAlbum As String)
    m_sAlbum = IIf(Len(Trim(sAlbum)) > 30, Left(Trim(sAlbum), 30), Trim(sAlbum))
End Property

Property Get Artist() As String
Attribute Artist.VB_Description = "Sets / returns artist name."
    Artist = m_sArtist
End Property

Property Let Artist(sArtist As String)
    m_sArtist = IIf(Len(Trim(sArtist)) > 30, Left(Trim(sArtist), 30), Trim(sArtist))
End Property

Property Get Comment() As String
Attribute Comment.VB_Description = "Sets / returns comment field."
    Comment = m_sComment
End Property

Property Let Comment(sComment As String)
    m_sComment = IIf(Len(Trim(sComment)) > 30, Left(Trim(sComment), 30), Trim(sComment))
End Property

Property Get FileExists() As Boolean
Attribute FileExists.VB_Description = "Returns True if file exists."
    FileExists = m_bFileExists
End Property

Property Get FileName() As String
Attribute FileName.VB_Description = "Sets / returns path to the file we want to work with."
    FileName = m_sPath
End Property

Property Let FileName(ByVal sPath As String)
    Dim FF As Integer
    sPath = Trim(sPath)
    If sPath = "" Then
        m_sPath = ""
        m_bFileExists = False
        m_bReadOnly = False
        Exit Property
    End If
    On Error Resume Next
    FF = FreeFile
    Open sPath For Input As #FF
    If Err.Number <> 0 Then
        Err.Clear
        m_sPath = ""
        m_bFileExists = False
        m_bReadOnly = False
        Exit Property
    End If
    Close #FF
    m_sPath = sPath
    m_bFileExists = True
    Open sPath For Append As #FF
    If Err.Number <> 0 Then
        Err.Clear
        m_bReadOnly = True
    Else
        Close #FF
        m_bReadOnly = False
    End If
    On Error GoTo 0
    Load
End Property

Property Get Genre() As enGenres
Attribute Genre.VB_Description = "Sets / returns genre."
    Genre = CLng(m_iGenre)
End Property

Property Let Genre(lGenre As enGenres)
    m_iGenre = CInt(lGenre)
End Property

Property Get Genre_str() As String
Attribute Genre_str.VB_Description = "Returns genre as string."
    Select Case m_iGenre
        Case 0
            Genre_str = "Blues"
        Case 1
            Genre_str = "Classic Rock"
        Case 2
            Genre_str = "Country"
        Case 3
            Genre_str = "Dance"
        Case 4
            Genre_str = "Disco"
        Case 5
            Genre_str = "Funk"
        Case 6
            Genre_str = "Grunge"
        Case 7
            Genre_str = "Hip-Hop"
        Case 8
            Genre_str = "Jazz"
        Case 9
            Genre_str = "Metal"
        Case 10
            Genre_str = "New Age"
        Case 11
            Genre_str = "Oldies"
        Case 12
            Genre_str = "Other"
        Case 13
            Genre_str = "Pop"
        Case 14
            Genre_str = "R&B"
        Case 15
            Genre_str = "Rap"
        Case 16
            Genre_str = "Reggae"
        Case 17
            Genre_str = "Rock"
        Case 18
            Genre_str = "Techno"
        Case 19
            Genre_str = "Industrial"
        Case 20
            Genre_str = "Alternative"
        Case 21
            Genre_str = "Ska"
        Case 22
            Genre_str = "Death Metal"
        Case 23
            Genre_str = "Pranks"
        Case 24
            Genre_str = "Soundtrack"
        Case 25
            Genre_str = "Euro-Techno"
        Case 26
            Genre_str = "Ambient"
        Case 27
            Genre_str = "Trip-Hop"
        Case 28
            Genre_str = "Vocal"
        Case 29
            Genre_str = "Jazz+Funk"
        Case 30
            Genre_str = "Fusion"
        Case 31
            Genre_str = "Trance"
        Case 32
            Genre_str = "Classical"
        Case 33
            Genre_str = "Instrumental"
        Case 34
            Genre_str = "Acid"
        Case 35
            Genre_str = "House"
        Case 36
            Genre_str = "Game"
        Case 37
            Genre_str = "Sound Clip"
        Case 38
            Genre_str = "Gospel"
        Case 39
            Genre_str = "Noise"
        Case 40
            Genre_str = "AlternRock"
        Case 41
            Genre_str = "Bass"
        Case 42
            Genre_str = "Soul"
        Case 43
            Genre_str = "Punk"
        Case 44
            Genre_str = "Space"
        Case 45
            Genre_str = "Meditative"
        Case 46
            Genre_str = "Instrumental Pop"
        Case 47
            Genre_str = "Instrumental Rock"
        Case 48
            Genre_str = "Ethnic"
        Case 49
            Genre_str = "Gothic"
        Case 50
            Genre_str = "Darkwave"
        Case 51
            Genre_str = "Techno-Industrial"
        Case 52
            Genre_str = "Electronic"
        Case 53
            Genre_str = "Pop-Folk"
        Case 54
            Genre_str = "Eurodance"
        Case 55
            Genre_str = "Dream"
        Case 56
            Genre_str = "Southern Rock"
        Case 57
            Genre_str = "Comedy"
        Case 58
            Genre_str = "Cult"
        Case 59
            Genre_str = "Gangsta"
        Case 60
            Genre_str = "Top 40"
        Case 61
            Genre_str = "Christian Rap"
        Case 62
            Genre_str = "Pop/Funk"
        Case 63
            Genre_str = "Jungle"
        Case 64
            Genre_str = "Native American"
        Case 65
            Genre_str = "Cabaret"
        Case 66
            Genre_str = "New Wave"
        Case 67
            Genre_str = "Psychadelic"
        Case 68
            Genre_str = "Rave"
        Case 69
            Genre_str = "Showtunes"
        Case 70
            Genre_str = "Trailer"
        Case 71
            Genre_str = "Lo-Fi"
        Case 72
            Genre_str = "Tribal"
        Case 73
            Genre_str = "Acid Punk"
        Case 74
            Genre_str = "Acid Jazz"
        Case 75
            Genre_str = "Polka"
        Case 76
            Genre_str = "Retro"
        Case 77
            Genre_str = "Musical"
        Case 78
            Genre_str = "Rock & Roll"
        Case 79
            Genre_str = "Hard Rock"
        Case 80
            Genre_str = "Folk"
        Case 81
            Genre_str = "Folk-Rock"
        Case 82
            Genre_str = "National Folk"
        Case 83
            Genre_str = "Swing"
        Case 84
            Genre_str = "Fast Fusion"
        Case 85
            Genre_str = "Bebob"
        Case 86
            Genre_str = "Latin"
        Case 87
            Genre_str = "Revival"
        Case 88
            Genre_str = "Celtic"
        Case 89
            Genre_str = "Bluegrass"
        Case 90
            Genre_str = "Avantgarde"
        Case 91
            Genre_str = "Gothic Rock"
        Case 92
            Genre_str = "Progressive Rock"
        Case 93
            Genre_str = "Psychedelic Rock"
        Case 94
            Genre_str = "Symphonic Rock"
        Case 95
            Genre_str = "Slow Rock"
        Case 96
            Genre_str = "Big Band"
        Case 97
            Genre_str = "Chorus"
        Case 98
            Genre_str = "Easy Listening"
        Case 99
            Genre_str = "Acoustic"
        Case 100
            Genre_str = "Humour"
        Case 101
            Genre_str = "Speech"
        Case 102
            Genre_str = "Chanson"
        Case 103
            Genre_str = "Opera"
        Case 104
            Genre_str = "Chamber Music"
        Case 105
            Genre_str = "Sonata"
        Case 106
            Genre_str = "Symphony"
        Case 107
            Genre_str = "Booty Bass"
        Case 108
            Genre_str = "Primus"
        Case 109
            Genre_str = "Porn Groove"
        Case 110
            Genre_str = "Satire"
        Case 111
            Genre_str = "Slow Jam"
        Case 112
            Genre_str = "Club"
        Case 113
            Genre_str = "Tango"
        Case 114
            Genre_str = "Samba"
        Case 115
            Genre_str = "Folklore"
        Case 116
            Genre_str = "Ballad"
        Case 117
            Genre_str = "Power Ballad"
        Case 118
            Genre_str = "Rhythmic Soul"
        Case 119
            Genre_str = "Freestyle"
        Case 120
            Genre_str = "Duet"
        Case 121
            Genre_str = "Punk Rock"
        Case 122
            Genre_str = "Drum Solo"
        Case 123
            Genre_str = "A capella"
        Case 124
            Genre_str = "Euro-House"
        Case 125
            Genre_str = "Dance Hall"
        Case Else
            Genre_str = ""
    End Select
End Property

Property Get HasID3tag() As Boolean
Attribute HasID3tag.VB_Description = "Returns True if file has an ID3 1.0 or 1.1 tag."
    HasID3tag = m_bHasID3tag
End Property

Property Get Name() As String
Attribute Name.VB_Description = "Sets / returns track name field."
    Name = m_sName
End Property

Property Let Name(sName As String)
    m_sName = IIf(Len(Trim(sName)) > 30, Left(Trim(sName), 30), Trim(sName))
End Property

Property Get ReadOnly() As Boolean
Attribute ReadOnly.VB_Description = "Returns True if the file can't be modified."
    ReadOnly = m_bReadOnly
End Property

Property Get SongNumber() As Integer
Attribute SongNumber.VB_Description = "Sets / returns song number in album. If set to 0, tag version will be 1.0 and song number won't be saved. If set to a valid value, tag version will be 1.1."
    SongNumber = m_iSongNumber
End Property

Property Let SongNumber(iSongNumber As Integer)
    If iSongNumber >= 0 And iSongNumber <= 255 Then
        m_iSongNumber = iSongNumber
    Else
        m_iSongNumber = 0
    End If
End Property

Property Get Year() As Integer
Attribute Year.VB_Description = "Sets / returns year field."
    Year = m_iYear
End Property

Property Let Year(iYear As Integer)
    If iYear >= 0 And iYear <= 9999 Then
        m_iYear = iYear
    Else
        m_iYear = 0
    End If
End Property

Function Load() As Boolean
Attribute Load.VB_Description = "Loads tag information from file. Automatically called after setting the FileName property."
    Dim lLen As Long, lPos As Long, FF As Integer, tTag As tipoTag
    Clear
    m_bHasID3tag = False
    If Not m_bFileExists Then
        Load = False
        Exit Function
    End If
    On Error Resume Next
    FF = FreeFile
    Open m_sPath For Binary As #FF
    If Err.Number <> 0 Then
        Close #FF
        Load = False
        Exit Function
    End If
    On Error GoTo 0
    lLen = LOF(FF)
    lPos = lLen - 128 + 1
    If Not lPos > 0 Then
        Close #FF
        Load = False
        Exit Function
    End If
    Seek #FF, lPos
    Get #FF, , tTag
    Close #FF
    With tTag
        If .Tag = "TAG" Then
            m_bHasID3tag = True
            m_sName = RTrim(.Name)
            If InStr(m_sName, Chr(0)) > 0 Then _
                m_sName = Left(m_sName, InStr(m_sName, Chr(0)) - 1)
            m_sAlbum = RTrim(.Album)
            If InStr(m_sAlbum, Chr(0)) > 0 Then _
                m_sAlbum = Left(m_sAlbum, InStr(m_sAlbum, Chr(0)) - 1)
            m_sArtist = RTrim(.Artist)
            If InStr(m_sArtist, Chr(0)) > 0 Then _
                m_sArtist = Left(m_sArtist, InStr(m_sArtist, Chr(0)) - 1)
            If Trim(.Year) <> "" Then
                If IsNumeric(.Year) Then
                    m_iYear = CInt(.Year)
                Else
                    m_iYear = 0
                End If
            Else
                m_iYear = 0
            End If
            If Mid(.Comment, 29, 1) = Chr(0) Then
                If Right(.Comment, 1) <> Chr(0) Then
                    m_iSongNumber = Asc(Right(.Comment, 1))
                    m_sComment = RTrim(Left(.Comment, 28))
                    If InStr(m_sComment, Chr(0)) Then _
                        m_sComment = Left(m_sComment, InStr(m_sComment, Chr(0)) - 1)
                Else
                    m_iSongNumber = 0
                    m_sComment = RTrim(.Comment)
                    If InStr(m_sComment, Chr(0)) Then _
                        m_sComment = Left(m_sComment, InStr(m_sComment, Chr(0)) - 1)
                End If
            Else
                m_iSongNumber = 0
                m_sComment = RTrim(.Comment)
                If InStr(m_sComment, Chr(0)) Then _
                    m_sComment = Left(m_sComment, InStr(m_sComment, Chr(0)) - 1)
            End If
            m_iGenre = Asc(.Genre)
            Load = True
        Else
            Load = False
        End If
    End With
End Function

Function Save() As Boolean
Attribute Save.VB_Description = "Saves ID3 tag information to file."
    Dim tTag As tipoTag, FF As Integer, lLen As Long, _
        lPos As Long
    If Not m_bFileExists Or m_bReadOnly Then
        Save = False
        Exit Function
    End If
    FF = FreeFile
    On Error Resume Next
    Open m_sPath For Binary As #FF
    If Err.Number <> 0 Then
        Close #FF
        Save = False
        Exit Function
    End If
    lLen = LOF(FF)
    lPos = lLen - 128 + 1
    If Not lPos > 0 Then
        Close #FF
        Save = False
        Exit Function
    End If
    Seek #FF, lPos
    Get #FF, , tTag
    If Not tTag.Tag = "TAG" Then lPos = lLen + 1
    With tTag
        .Album = m_sAlbum & String(30 - Len(m_sAlbum), Chr(0))
        .Artist = m_sArtist & String(30 - Len(m_sArtist), Chr(0))
        If m_iSongNumber > 0 Then
            If Len(m_sComment) > 28 Then _
                m_sComment = Left(m_sComment, 28)
            .Comment = m_sComment & String(28 - Len(m_sComment), Chr(0)) & Chr(0) & Chr(m_iSongNumber)
        Else
            .Comment = m_sComment & String(30 - Len(m_sComment), Chr(0))
        End If
        .Genre = Chr(m_iGenre)
        .Name = m_sName & String(30 - Len(m_sName), Chr(0))
        .Tag = "TAG"
        .Year = IIf(m_iYear > 0, Format(m_iYear, "0000"), String(4, Chr(0)))
    End With
    Seek #FF, lPos
    Put #FF, , tTag
    Close #FF
    m_bHasID3tag = True
    Save = True
End Function

Function RemoveTag() As Boolean
Attribute RemoveTag.VB_Description = "Removes ID3 tag information from file."
    Dim FF As Integer, tTag As tipoTag, _
        lLen As Long, lPos As Long, lHnd As Long, SA As SECURITY_ATTRIBUTES, _
        lRes As Long
    If Not m_bFileExists Or m_bReadOnly Or Not m_bHasID3tag Then
        RemoveTag = False
        Exit Function
    End If
    FF = FreeFile
    On Error Resume Next
    Open m_sPath For Binary As #FF
    If Err.Number <> 0 Then
        RemoveTag = False
        Exit Function
    End If
    On Error GoTo 0
    lLen = LOF(FF)
    lPos = lLen - 128 + 1
    If Not lPos > 0 Then
        Close #FF
        RemoveTag = False
        Exit Function
    End If
    Seek #FF, lPos
    Get #FF, , tTag
    If Not tTag.Tag = "TAG" Then
        Close #FF
        RemoveTag = False
        Exit Function
    End If
    Close #FF

    lHnd = CreateFile(m_sPath & Chr(0), GENERIC_READ Or GENERIC_WRITE, 0&, _
        SA, OPEN_EXISTING, 0&, 0&)
    If lHnd = INVALID_HANDLE_VALUE Then
        RemoveTag = False
        Exit Function
    End If
    lPos = lPos - 1
    lRes = SetFilePointer(lHnd, lPos, 0&, FILE_BEGIN)
    lRes = SetEndOfFile(lHnd)
    If lRes = 0 Then
        RemoveTag = False
    Else
        RemoveTag = True
        m_bHasID3tag = False
    End If
    CloseHandle lHnd
End Function

Sub Clear()
Attribute Clear.VB_Description = "Clears all ID3 tag fields."
    m_sName = ""
    m_sAlbum = ""
    m_sArtist = ""
    m_iYear = 0
    m_iGenre = 255
    m_sComment = ""
    m_iSongNumber = 0
End Sub

Private Sub Class_Initialize()
    Clear
End Sub
