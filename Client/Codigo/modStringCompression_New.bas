Attribute VB_Name = "modStringCompression"
'*******************************************************************************
' MODULE:       MZlib
' FILENAME:     C:\My Code\vb\zlib\MZlib.bas
' AUTHOR:       Phil Fresle
' CREATED:      20-Feb-2000
' COPYRIGHT:    Copyright 2000 Frez Systems Limited. All Rights Reserved.
'
' DESCRIPTION:
' This module wraps the Zlib DLL compress and uncompress routines that can be
' be used to compress and uncompress strings.  It is only really useful for long
' strings that contain a certain amount of repeating text. One such use might be
' to send XML over a network.
'
' The compress wrapper adds the original string size onto the front of the
' compressed string so when it comes to uncompressing you know the size of
' string you need to allocate.
'
' Zlib is a free lossless data-compression library available for a number of
' platforms, the homepage for it is http://www.cdrom.com/pub/infozip/zlib/
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on the add-in provided
' here, or any consultancy work, please do not hesitate to contact us.
'
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk
'
' MODIFICATION HISTORY:
' 1.0       20-Feb-2000
'           Phil Fresle
'           Initial Version
'*******************************************************************************
Option Explicit

Private Declare Function Compress Lib "ZLIB.DLL" _
        Alias "compress" (ByVal compr As String, comprLen As _
        Any, ByVal buf As String, ByVal buflen As Long) As Long
        
Private Declare Function Uncompress Lib "ZLIB.DLL" _
        Alias "uncompress" (ByVal uncompr As String, uncomprLen As _
        Any, ByVal compr As String, ByVal lcompr As Long) As Long

Private Const Z_OK              As Long = 0
Private Const Z_STREAM_END      As Long = 1
Private Const Z_NEED_DICT       As Long = 2
Private Const Z_ERRNO           As Long = -1
Private Const Z_STREAM_ERROR    As Long = -2
Private Const Z_DATA_ERROR      As Long = -3
Private Const Z_MEM_ERROR       As Long = -4
Private Const Z_BUF_ERROR       As Long = -5
Private Const Z_VERSION_ERROR   As Long = -6

'*******************************************************************************
' CompressString (FUNCTION)
'
' PARAMETERS:
' (In) - StringToCompress - Variant - String to compress
'
' RETURN VALUE:
' String - Compressed string
'
' DESCRIPTION:
' Compresses a string with the Zlib compress routine. Sticks the number of
' characters on the front of the string to aid with decompressing the data
' later.
'*******************************************************************************
Public Function CompressString(ByVal StringToCompress As String) As String
    Dim sCompressed     As String
    Dim lCompressedLen  As Integer
    Dim intStringLen    As Integer
    Dim lReturn         As Long
    
    intStringLen = Len(StringToCompress)
    lCompressedLen = (intStringLen * 1.01) + 13
    sCompressed = Space$(lCompressedLen)
    
    lReturn = Compress(sCompressed, lCompressedLen, StringToCompress, CLng(intStringLen))
    
    Select Case lReturn
        Case Z_OK
            sCompressed = left$(sCompressed, lCompressedLen)
            CompressString = Integer_To_String(intStringLen) & sCompressed
        Case Z_MEM_ERROR
            Err.Raise vbObjectError + Abs(lReturn), "CompressString", _
                "Insufficient memory to compress string"
        Case Z_BUF_ERROR
            Err.Raise vbObjectError + Abs(lReturn), "CompressString", _
                "Insufficient space in output buffer to compress string"
        Case Else
            Err.Raise vbObjectError + Abs(lReturn), "CompressString", _
                "Unknown error during compress operation"
    End Select
End Function

'*******************************************************************************
' UncompressString (FUNCTION)
'
' PARAMETERS:
' (In) - CompressedString - Variant -
'
' RETURN VALUE:
' String -
'
' DESCRIPTION:
' Uncompresses a string with the Zlib uncompress routine that has been previously
' compressed with the CompressString function in this module as it relies on the
' string starting with the number of characters required to output the string.
'*******************************************************************************
Public Function UncompressString(ByVal CompressedString As String) As String
    Dim sUncompressedString As String
    Dim intUncompressedLen    As Integer
    Dim sBuffer             As String
    Dim lBufferLen          As Long
    Dim lReturn             As Long
    
    intUncompressedLen = String_To_Integer(CompressedString, 1)
    sUncompressedString = Space$(intUncompressedLen)
    
    sBuffer = mid$(CompressedString, 3)
    lBufferLen = Len(sBuffer)
    
    lReturn = Uncompress(sUncompressedString, intUncompressedLen, sBuffer, lBufferLen)
    
    Select Case lReturn
        Case Z_OK
            UncompressString = sUncompressedString
        Case Z_MEM_ERROR
            Err.Raise vbObjectError + Abs(lReturn), "UncompressString", _
                "Insufficient memory to uncompress string"
        Case Z_BUF_ERROR
            Err.Raise vbObjectError + Abs(lReturn), "UncompressString", _
                "Insufficient space in output buffer to uncompress string"
        Case Z_DATA_ERROR
            Err.Raise vbObjectError + Abs(lReturn), "UncompressString", _
                "Cannot uncompress corrupt data"
        Case Else
            Err.Raise vbObjectError + Abs(lReturn), "UncompressString", _
                "Unknown error during uncompress operation"
    End Select
End Function
