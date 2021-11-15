VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZipFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eZLib
   Z_OK = 0
   Z_STREAM_ERROR = -2 'Invalid compression level parameter
   Z_DATA_ERROR = -3 'Input data corrupted
   Z_MEM_ERROR = -4 'Not Enough Memory
   Z_BUF_ERROR = -5 'Not enough space in output buffer
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal Length As Long)

Private Declare Function UnCompress Lib "zlib.dll" Alias "uncompress" (Dest As Any, destLen As Any, Src As Any, ByVal srcLen As Long) As Long
Private Declare Function Crc32 Lib "zlib.dll" Alias "crc32" (ByVal crc As Long, Buffer As Any, ByVal Length As Long) As Long

Private Type typCentralFileHeader
   CentralFileHeaderSigniature As Long
   VersionMadeBy As Integer
   VersionNeededToExtract As Integer
   GeneralPurposeBitFlag As Integer
   CompressionMethod As Integer
   LastModFileTime As Integer
   LastModFileDate As Integer
   Crc32 As Long
   CompressedSize As Long
   UnCompressedSize As Long
   FileNameLength As Integer
   ExtraFieldLength As Integer
   FileCommentLength As Integer
   DiskNumberStart As Integer
   InternalFileAttributes As Integer
   ExternalFileAttributes As Long
   RelativeOffsetOfLocalHeader As Long
End Type

Private Type typLocalFileHeader
   LocalFileHeaderSignature As Long
   VersionNeededToExtract As Integer
   GeneralPurposeBitFlag As Integer
   CompressionMethod As Integer
   LastModFileTime As Integer
   LastModFileDate As Integer
   Crc32 As Long
   CompressedSize As Long
   UnCompressedSize As Long
   FileNameLength As Integer
   ExtraFieldLength As Integer
End Type

Private FileName As String
Private ExtraField As String
Private FileComment As String
Private FileData() As Byte

Private LocalFileHeader As typLocalFileHeader
Private CentralFileHeader As typCentralFileHeader

Private Sub Class_Initialize()

   With LocalFileHeader
      .LocalFileHeaderSignature = &H4034B50
      .VersionNeededToExtract = 20
      .GeneralPurposeBitFlag = 2
      .CompressionMethod = 8
   End With

   With CentralFileHeader
      .CentralFileHeaderSigniature = &H2014B50
      .VersionMadeBy = 20
      .VersionNeededToExtract = 20
      .GeneralPurposeBitFlag = 2
      .CompressionMethod = 8
      .InternalFileAttributes = 32
   End With
   
End Sub

Friend Property Let FilePath(s As String)
   
   Dim ModDate As Integer
   Dim ModTime As Integer

   Dim d As Date

   d = FileDateTime(s)
   
   ModDate = GetDOSDate(d)
   ModTime = GetDOSTime(d)

   FileName = GetFileName(s)
   With LocalFileHeader
      .FileNameLength = Len(FileName)
      .LastModFileDate = ModDate
      .LastModFileTime = ModTime
      CompressBytes s, FileData, .CompressedSize, .UnCompressedSize
   End With
   With CentralFileHeader
      .FileNameLength = Len(FileName)
      .LastModFileDate = ModDate
      .LastModFileTime = ModTime
      .CompressedSize = LocalFileHeader.CompressedSize
      .UnCompressedSize = LocalFileHeader.UnCompressedSize
   End With
   
End Property

Friend Property Let NewFileName(s As String)

   FileName = s
   
   LocalFileHeader.FileNameLength = Len(FileName)
   CentralFileHeader.FileNameLength = Len(FileName)

End Property

Friend Property Let Comment(s As String)

   With CentralFileHeader
      FileComment = s
      .FileCommentLength = Len(s)
   End With

End Property

Private Function GetFileName(FilePath As String) As String

   Dim a() As String
   
   a = Split(FilePath, "\")
   GetFileName = a(UBound(a))

End Function

Friend Sub WriteLocalFileHeader(fh As Long)

   CentralFileHeader.RelativeOffsetOfLocalHeader = Loc(fh)

   Put #fh, , LocalFileHeader
   Put #fh, , FileName
   Put #fh, , ExtraField
   Put #fh, , FileData
   
End Sub

Friend Sub WriteCentralFileHeader(fh As Long)

   Put #fh, , CentralFileHeader
   Put #fh, , FileName
   Put #fh, , ExtraField
   Put #fh, , FileComment

End Sub

Private Function GetDOSDate(ModDate As Date) As Integer

   Dim Day As Long
   Dim Month As Long
   Dim Year As Long
   
   Dim b(1) As Byte
   
   'There's a Windows API Function FileTimeToDosDateTime
   'but I couldn't get it to work so I did this the hard way

   'Bits   Contents
   '0–4    Day of the month (1–31)
   '5–8    Month (1 = January, 2 = February, etc.)
   '9–15   Year offset from 1980 (add 1980 to get actual year)
   
   Day = DatePart("d", ModDate)
   Month = DatePart("m", ModDate)
   Year = DatePart("yyyy", ModDate) - 1980
   
   If GetBit(Day, 0) Then SetBit b(0), 0
   If GetBit(Day, 1) Then SetBit b(0), 1
   If GetBit(Day, 2) Then SetBit b(0), 2
   If GetBit(Day, 3) Then SetBit b(0), 3
   If GetBit(Day, 4) Then SetBit b(0), 4
   If GetBit(Month, 0) Then SetBit b(0), 5
   If GetBit(Month, 1) Then SetBit b(0), 6
   If GetBit(Month, 2) Then SetBit b(0), 7
   If GetBit(Month, 3) Then SetBit b(1), 0
   If GetBit(Year, 0) Then SetBit b(1), 1
   If GetBit(Year, 1) Then SetBit b(1), 2
   If GetBit(Year, 2) Then SetBit b(1), 3
   If GetBit(Year, 3) Then SetBit b(1), 4
   If GetBit(Year, 4) Then SetBit b(1), 5
   If GetBit(Year, 5) Then SetBit b(1), 6
   If GetBit(Year, 6) Then SetBit b(1), 7
   
   CopyMemory GetDOSDate, b(0), 2
   
End Function

Private Function GetDOSTime(ModDate As Date) As Integer

   Dim Second As Long
   Dim Minute As Long
   Dim Hour As Long
   
   Dim b(1) As Byte
   
   'Bits   Contents
   '0–4    Second divided by 2
   '5–10   Minute (0–59)
   '11–15  Hour (0–23 on a 24-hour clock)
   
   Second = DatePart("s", ModDate) \ 2
   Minute = DatePart("n", ModDate)
   Hour = DatePart("h", ModDate)
   
   If GetBit(Second, 0) Then SetBit b(0), 0
   If GetBit(Second, 1) Then SetBit b(0), 1
   If GetBit(Second, 2) Then SetBit b(0), 2
   If GetBit(Second, 3) Then SetBit b(0), 3
   If GetBit(Second, 4) Then SetBit b(0), 4
   If GetBit(Minute, 0) Then SetBit b(0), 5
   If GetBit(Minute, 1) Then SetBit b(0), 6
   If GetBit(Minute, 2) Then SetBit b(0), 7
   If GetBit(Minute, 3) Then SetBit b(1), 0
   If GetBit(Minute, 4) Then SetBit b(1), 1
   If GetBit(Minute, 5) Then SetBit b(1), 2
   If GetBit(Hour, 0) Then SetBit b(1), 3
   If GetBit(Hour, 1) Then SetBit b(1), 4
   If GetBit(Hour, 2) Then SetBit b(1), 5
   If GetBit(Hour, 3) Then SetBit b(1), 6
   If GetBit(Hour, 4) Then SetBit b(1), 7
   
   CopyMemory GetDOSTime, b(0), 2
   
End Function

Private Sub SetBit(b As Byte, Bit As Long)

   b = b Or (2 ^ Bit)

End Sub

Private Function GetBit(l As Long, Bit As Long) As Boolean

   GetBit = ((l And 2 ^ Bit) > 0)

End Function
