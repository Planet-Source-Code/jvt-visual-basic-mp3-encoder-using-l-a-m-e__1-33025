Attribute VB_Name = "Mp3Encoder"
Global ID3 As New ID3tag
Global TID3 As New ID3tag
Public Enum VBRMETHOD
    VBR_METHOD_NONE = -1
    VBR_METHOD_DEFAULT = 0
    VBR_METHOD_OLD = 1
    VBR_METHOD_NEW = 2
    VBR_METHOD_MTRH = 3
    VBR_METHOD_ABR = 4
End Enum

Public Enum EncodingErrors
    ENC_ERR_ENCODING_SUCCESS = 0
    ENC_ERR_ENCODING_FAILED = -1
    ENC_ERR_ENCODING_CANCELLED = -2
    ENC_ERR_NO_API = -3
    ENC_ERR_INPUT = -4
    ENC_ERR_OUTPUT = -5
    ENC_ERR_INVALID_PARAMS = -6
End Enum

Public Enum EncodeMode
    BE_MP3_MODE_STEREO = 0
    BE_MP3_MODE_JSTEREO = 1
    BE_MP3_MODE_DUALCHANNEL = 2
    BE_MP3_MODE_MONO = 3
End Enum

Global EncCancel As Boolean

'API declarations for encoding wrapper
Public Declare Function SetVBR Lib "MP3Enc.dll" (ByVal Enable As Long, ByVal Quality As Long, ByVal Method As VBRMETHOD, ByVal MaxBitRate As Long) As Long
Public Declare Function EncodeMp3 Lib "MP3Enc.dll" (ByVal lpszWavFile As String, ByVal lpszOutFile As String, ByVal BitRate As Long, ByVal SampleRate As Long, ByVal EncMode As EncodeMode, lpCallback As Any) As Long
Public Function EnumEncoding(ByVal nStatus As Integer) As Boolean
        
    If frmMain.pb1.Value <> nStatus Then
        frmMain.lblPercent.Caption = nStatus & "%"
        frmMain.pb1.Value = nStatus
    End If
    
    If EncCancel Then
        EnumEncoding = False
        EncCancel = False
    Else
        EnumEncoding = True
    End If
    
    DoEvents
End Function
