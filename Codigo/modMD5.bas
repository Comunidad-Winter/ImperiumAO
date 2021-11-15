Attribute VB_Name = "modMD5"
'*****************************************************************
'modMD5 - ImperiumAO - v1.3.0
'
'MD5 calculations.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

' MD5.bas - wrapper for RSA MD5 DLL
'   derived from the RSA Data Security, Inc. MD5 Message-Digest Algorithm
' Functions:
'   MD5String (some string) -> MD5 digest of the given string as 32 bytes string
'   MD5File (some filename) -> MD5 digest of the file's content as a 32 bytes string
'      returns a null terminated "FILE NOT FOUND" if unable to open the
'      given filename for input
' Bugs, complaints, etc:
'   Francisco Carlos Piragibe de Almeida
'   piragibe@esquadro.com.br
' History
'       Apr, 17 1999 - fixed the null byte problem
' Contains public domain RSA C-code for MD5 digest (see MD5-original.txt file)
' The aamd532.dll DLL MUST be somewhere in your search path
'   for this to work

Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)

Public Function MD5String(ByVal p As String) As String
' compute MD5 digest on a given string, returning the result
    Dim r As String * 32, t As Long
    r = Space(32)
    t = Len(p)
    MDStringFix p, t, r
    MD5String = r
End Function

Public Function MD5File(ByVal f As String) As String
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function

Public Function hexMd52Asc(ByVal MD5 As String) As String
    Dim i As Integer, l As String
    
    MD5 = UCase$(MD5)
    If Len(MD5) Mod 2 = 1 Then MD5 = "0" & MD5
    
    For i = 1 To Len(MD5) \ 2
        l = mid(MD5, (2 * i) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr(hexHex2Dec(l))
    Next i
End Function

Public Function hexHex2Dec(ByVal hex As String) As Long
    
If Len(hex) > 1 Then
    If left$(hex, 2) <> "&H" Then
        hex = "&H" & hex
    End If
Else
    hex = "&H" & hex
End If

hexHex2Dec = Val("&H" & hex)

End Function

Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String
    Dim i As Integer, l As String
    For i = 1 To Len(Text)
        l = mid(Text, i, 1)
        txtOffset = txtOffset & Chr((Asc(l) + off) Mod 256)
    Next i
End Function
