Attribute VB_Name = "modGeneral"
'*****************************************************************
'modGeneral.bas - Generic Functions - v0.5.0
'
'Generic helper functions
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

'*****************************************************************
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'Augusto José Rando (barrin@imperiumao.com.ar) - 6/11/2005
'   - Add: General_Get_Temp_Dir method
'   - Add: Methods to get free RAM and total RAM, and pagefile
'
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 8/25/2004
'   - Add: General_Convert_Radians_To_Degrees method
'   - Add: General_Drive_Get_Free_Bytes method
'   - Add: General_Write_To_TextBox method
'   - Fix: A spelling mistake in General_Convert_Degrees_To_Radians (used to say Covert)
'
'Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) - 2/19/2003
'   - Add: General_Long_Color_To_RGB method
'
'Aaron Perkins(aaron@baronsoft.com) - 5/14/2002
'   - Still in Test Phase
'*****************************************************************
Option Explicit

'***************************
'Constants
'***************************
Public Const PI As Single = 3.14159265358979

'***************************
'External Functions
'***************************
'General_Var_Get and General_Var_Write
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Added by Augusto José Rando
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_LENGTH = 512

Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

'For making a form always on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

'To get OS version
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      ' Maintenance string for PSS usage
End Type

Private Const VER_PLATFORM_WIN32s As Long = 0&
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1&
Private Const VER_PLATFORM_WIN32_NT As Long = 2&

Private Declare Function GetOSVersion Lib "kernel32" _
Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private OSInfo As OSVERSIONINFO

'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, BytesTotal As Currency, FreeBytesTotal As Currency) As Long

'To get free bytes in RAM

Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Function General_File_Exists(ByVal file_path As String, ByVal file_type As VbFileAttribute) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a file exists
'*****************************************************************
    If Dir(file_path, file_type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function

Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = ""
    
    sSpaces = Space$(5000)
    
    GetPrivateProfileString Main, var, szReturn, sSpaces, Len(sSpaces), file
    
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function

Public Sub General_Var_Write(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal Value As String)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, Value, file
End Sub

Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As String) As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    
    LastPos = 0
    CurrentPos = 0
    
    For i = 1 To field_pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        General_Field_Read = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        General_Field_Read = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

Public Function General_Field_Count(ByVal Text As String, ByVal delimiter As Byte) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Count the number of fields in a delimited string
'*****************************************************************
    'If string is empty there aren't any fields
    If Len(Text) = 0 Then
        Exit Function
    End If

    Dim i As Long
    Dim FieldNum As Long
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
        End If
    Next i
    General_Field_Count = FieldNum + 1
End Function

Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
'*****************************************************************
'Author: Aaron Perkins
'Find a Random number between a range
'*****************************************************************
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Sub General_Sleep(ByVal Length As Double)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sleep for a given number a seconds
'*****************************************************************
    Dim curFreq As Currency
    Dim curStart As Currency
    Dim curEnd As Currency
    Dim dblResult As Double
    
    QueryPerformanceFrequency curFreq 'Get the timer frequency
    QueryPerformanceCounter curStart 'Get the start time
    
    Do Until dblResult >= Length
        QueryPerformanceCounter curEnd 'Get the end time
        dblResult = (curEnd - curStart) / curFreq 'Calculate the duration (in seconds)
        DoEvents
    Loop
End Sub

Public Function General_Get_Elapsed_Time() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If

    'Get current time
    QueryPerformanceCounter start_time
    
    'Calculate elapsed time
    General_Get_Elapsed_Time = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    QueryPerformanceCounter end_time
End Function

Public Function General_String_Is_Alphanumeric(ByVal test_string As String) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/09/2003
'Alows no special characters. Only letters, numbers, and spaces
'**************************************************************
    'Check Name
    Dim loopc As Long
    Dim ts As Byte
    For loopc = 1 To Len(test_string)
        ts = Asc(mid(test_string, loopc))
        If ts = 32 Or (ts >= 48 And ts <= 57) Or (ts >= 65 And ts <= 90) Or (ts >= 96 And ts <= 122) Then
            General_String_Is_Alphanumeric = True
        Else
            General_String_Is_Alphanumeric = False
            Exit Function
        End If
    Next loopc
End Function

Public Sub General_Form_On_Top_Set(ByRef s_form As Form, Optional ByVal s_on_top As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/09/2003
'Set a form on always on top or not
'**************************************************************
    Dim hwnd_flag As Long
    
    If s_on_top Then
        hwnd_flag = HWND_TOPMOST
    Else
        hwnd_flag = HWND_NOTOPMOST
    End If
    SetWindowPos s_form.hwnd, hwnd_flag, s_form.left / Screen.TwipsPerPixelX, s_form.top / Screen.TwipsPerPixelY, s_form.Width / Screen.TwipsPerPixelX, s_form.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Public Function General_Convert_Degrees_To_Radians(ByRef s_degree As Single) As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/12/2003
'Converts a degree to a radian
'**************************************************************
    General_Convert_Degrees_To_Radians = (s_degree * PI) / 180
End Function

Public Function General_Convert_Radians_To_Degrees(ByRef s_radians As Single) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/25/2004
'Converts a radian to degrees
'**************************************************************
    General_Convert_Radians_To_Degrees = (s_radians * 180) / PI
End Function

Public Function General_Distance(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Finds the distance between two points
'**************************************************************
    General_Distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function

Public Function General_RGB_Color_to_Long(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
        
    Dim C As Long
        
    If a > 127 Then
        a = a - 128
        C = a * 2 ^ 24 Or &H80000000
        C = C Or r * 2 ^ 16
        C = C Or g * 2 ^ 8
        C = C Or b
    Else
        C = a * 2 ^ 24
        C = C Or r * 2 ^ 16
        C = C Or g * 2 ^ 8
        C = C Or b
    End If
    
    General_RGB_Color_to_Long = C

End Function

Public Sub General_Long_Color_to_RGB(ByVal long_color As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
'***********************************
'Coded by Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 2/19/03
'Takes a long value and separates RGB values to the given variables
'***********************************
    Dim temp_color As String
    
    temp_color = hex(long_color)
    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String(6 - Len(temp_color), "0") + temp_color
    End If
    
    red = CLng("&H" + mid(temp_color, 1, 2))
    green = CLng("&H" + mid(temp_color, 3, 2))
    blue = CLng("&H" + mid(temp_color, 5, 2))

End Sub

Public Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/07/2004
'
'**************************************************************
    Dim RetVal As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    
    RetVal = GetDiskFreeSpace(left(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function

Public Sub General_Write_To_TextBox(ByVal TextBox As TextBox, ByVal Text As String)
'**************************************************************
'Author: juan Martín Sotuyo Dodero
'Last Modify Date: 8/14/2004
'
'**************************************************************
    TextBox.Text = TextBox.Text & vbNewLine & Text
End Sub

Public Sub General_Quick_Sort(ByRef SortArray As Variant, ByVal first As Long, ByVal last As Long)
'**************************************************************
'Author: juan Martín Sotuyo Dodero
'Last Modify Date: 3/03/2005
'Good old QuickSort algorithm :)
'**************************************************************
    Dim Low As Long, High As Long
    Dim temp As Variant
    Dim List_Separator As Variant
    
    Low = first
    High = last
    List_Separator = SortArray((first + last) / 2)
    Do While (Low <= High)
        Do While SortArray(Low) < List_Separator
            Low = Low + 1
        Loop
        Do While SortArray(High) > List_Separator
            High = High - 1
        Loop
        If Low <= High Then
            temp = SortArray(Low)
            SortArray(Low) = SortArray(High)
            SortArray(High) = temp
            Low = Low + 1
            High = High - 1
        End If
    Loop
    If first < High Then General_Quick_Sort SortArray, first, High
    If Low < last Then General_Quick_Sort SortArray, Low, last
End Sub

Public Function General_Get_Temp_Dir() As String
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'Gets windows temporary directory
'**************************************************************
   Dim s As String
   Dim C As Long
   s = Space$(MAX_LENGTH)
   C = GetTempPath(MAX_LENGTH, s)
   If C > 0 Then
       If C > Len(s) Then
           s = Space$(C + 1)
           C = GetTempPath(MAX_LENGTH, s)
       End If
   End If
   General_Get_Temp_Dir = IIf(C > 0, left$(s, C), "")
End Function

Public Function General_Load_Picture_From_Resource(ByVal picture_file_name As String) As IPicture
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'Loads a picture from a resource file and returns it
'**************************************************************

'On Error GoTo ErrorHandler

If Extract_File(Interface, App.Path & "\Recursos\", picture_file_name, Windows_Temp_Dir, False) Then
    Set General_Load_Picture_From_Resource = LoadPicture(Windows_Temp_Dir & picture_file_name)
    Call Delete_File(Windows_Temp_Dir & picture_file_name)
Else
    Set General_Load_Picture_From_Resource = Nothing
End If

Exit Function

ErrorHandler:
    If General_File_Exists(Windows_Temp_Dir & picture_file_name, vbNormal) Then
        Call Delete_File(Windows_Temp_Dir & picture_file_name)
    End If

End Function

Public Function General_Bytes_To_Megabytes(Bytes As Double) As Double
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'
'**************************************************************

Dim dblAns As Double
dblAns = (Bytes / 1024) / 1024
General_Bytes_To_Megabytes = format(dblAns, "###,###,##0.00")
End Function

Public Function General_Get_Total_Ram() As Double
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'
'**************************************************************
    
    'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPhys
    General_Get_Total_Ram = General_Bytes_To_Megabytes(dblAns)
End Function

Public Function General_Get_Free_Ram() As Double
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'
'**************************************************************
        
    'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    General_Get_Free_Ram = General_Bytes_To_Megabytes(dblAns)
End Function

Public Function General_Get_Free_Ram_Bytes() As Long
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'
'**************************************************************
    GlobalMemoryStatus pUdtMemStatus
    General_Get_Free_Ram_Bytes = pUdtMemStatus.dwAvailPhys
End Function

Public Function General_Get_Page_File_Size() As Double
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'
'**************************************************************
        
    'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPageFile
    General_Get_Page_File_Size = General_Bytes_To_Megabytes(dblAns)
End Function

Public Function General_Is_Valid_Email(ByVal AddressString As String) As Boolean

Dim sHost As String
Dim iPos As Integer
Dim sInvalidChars As String

If Len(Trim(AddressString)) = 0 Then
    General_Is_Valid_Email = False
    Exit Function
End If

sInvalidChars = "!#$%^&*()=+{}[]|\;:'/?>,< "

    For iPos = 1 To Len(AddressString)
        If InStr(sInvalidChars, _
            mid(AddressString, iPos, 1)) > 0 Then
            
            General_Is_Valid_Email = False
            Exit Function
        End If
    Next


iPos = InStr(AddressString, "@")

If iPos = 0 Or left(AddressString, 1) = "@" Then
    General_Is_Valid_Email = False
    Exit Function
End If

sHost = mid(AddressString, iPos + 1)
'can't have multiple "@" chars in the string
If InStr(sHost, "@") > 0 Then
    General_Is_Valid_Email = False
    Exit Function
End If

General_Is_Valid_Email = General_Is_Valid_IP_Host(UCase(sHost))

End Function

Private Function General_Is_Valid_IP_Host(ByVal HostString As String) As Boolean

Dim sHost As String
Dim bDottedQuad As Boolean
Dim sSplit() As String
Dim ictr As Integer
Dim bAns As Boolean
Dim sTopLevelDomains() As String

sHost = HostString

If InStr(sHost, ".") = 0 Then
    General_Is_Valid_IP_Host = False
    Exit Function
End If

sSplit = Split(sHost, ".")

If UBound(sSplit) = 3 Then
    bDottedQuad = True
    For ictr = 0 To 3
        If Not IsNumeric(sSplit(ictr)) Then
            bDottedQuad = False
            Exit For
        End If
    Next
    
    If bDottedQuad Then
        bAns = True
        For ictr = 0 To 3
            If ictr = 0 Then
            bAns = Val(sSplit(ictr)) <= 239
                If bAns = False Then Exit For
            Else
                bAns = Val(sSplit(ictr)) <= 255
                If bAns = False Then Exit For
            End If
        Next
        
        General_Is_Valid_IP_Host = bAns

        
        Exit Function
    End If
End If 'ubound(ssplit) = 3

    General_Is_Valid_IP_Host = General_Is_TopLevel_Domain(sSplit(UBound(sSplit)))

End Function

Private Function General_Is_TopLevel_Domain(ByVal DomainString As String) As Boolean

Dim asTopLevels() As String
Dim ictr As Integer
Dim iNumDomains As Integer
Dim bAns As Boolean

iNumDomains = 251
ReDim asTopLevels(iNumDomains - 1) As String


'Obtained from www.IANA.com.  Can and will change
asTopLevels(0) = "COM"
asTopLevels(1) = "ORG"
asTopLevels(2) = "NET"
asTopLevels(3) = "EDU"
asTopLevels(4) = "GOV"
asTopLevels(5) = "MIL"
asTopLevels(6) = "INT"
asTopLevels(7) = "AF"
asTopLevels(8) = "AL"
asTopLevels(9) = "DZ"
asTopLevels(10) = "AS"
asTopLevels(11) = "AD"
asTopLevels(12) = "AO"
asTopLevels(13) = "AI"
asTopLevels(14) = "AQ"
asTopLevels(15) = "AG"
asTopLevels(16) = "AR"
asTopLevels(17) = "AM"
asTopLevels(18) = "AW"
asTopLevels(19) = "AC"
asTopLevels(20) = "AU"
asTopLevels(21) = "AT"
asTopLevels(22) = "AZ"
asTopLevels(23) = "BS"
asTopLevels(24) = "BH"
asTopLevels(25) = "BD"
asTopLevels(26) = "BB"
asTopLevels(27) = "BY"
asTopLevels(28) = "BZ"
asTopLevels(29) = "BT"
asTopLevels(30) = "BJ"
asTopLevels(31) = "BE"
asTopLevels(32) = "BM"
asTopLevels(33) = "BO"
asTopLevels(34) = "BA"
asTopLevels(35) = "BW"
asTopLevels(36) = "BV"
asTopLevels(37) = "BR"
asTopLevels(38) = "IO"
asTopLevels(39) = "BN"
asTopLevels(40) = "BG"
asTopLevels(41) = "BF"
asTopLevels(42) = "BI"
asTopLevels(43) = "KH"
asTopLevels(44) = "CM"
asTopLevels(45) = "CA"
asTopLevels(46) = "CV"
asTopLevels(47) = "KY"
asTopLevels(48) = "CF"
asTopLevels(49) = "TD"
asTopLevels(50) = "CL"
asTopLevels(51) = "CN"
asTopLevels(52) = "CX"
asTopLevels(53) = "CC"
asTopLevels(54) = "CO"
asTopLevels(55) = "KM"
asTopLevels(56) = "CD"
asTopLevels(57) = "CG"
asTopLevels(58) = "CK"
asTopLevels(59) = "CR"
asTopLevels(60) = "CI"
asTopLevels(61) = "HR"
asTopLevels(62) = "CU"
asTopLevels(63) = "CY"
asTopLevels(64) = "CZ"
asTopLevels(65) = "DK"
asTopLevels(66) = "DJ"
asTopLevels(67) = "DM"
asTopLevels(68) = "DO"
asTopLevels(69) = "TP"
asTopLevels(70) = "EC"
asTopLevels(71) = "EG"
asTopLevels(72) = "SV"
asTopLevels(73) = "GQ"
asTopLevels(74) = "ER"
asTopLevels(75) = "EE"
asTopLevels(76) = "ET"
asTopLevels(77) = "FK"
asTopLevels(78) = "FO"
asTopLevels(79) = "FJ"
asTopLevels(80) = "FI"
asTopLevels(81) = "FR"
asTopLevels(82) = "GF"
asTopLevels(83) = "PF"
asTopLevels(84) = "TF"
asTopLevels(85) = "GA"
asTopLevels(86) = "GM"
asTopLevels(87) = "GE"
asTopLevels(88) = "DE"
asTopLevels(89) = "GH"
asTopLevels(90) = "GI"
asTopLevels(91) = "GR"
asTopLevels(92) = "GL"
asTopLevels(93) = "GD"
asTopLevels(94) = "GP"
asTopLevels(95) = "GU"
asTopLevels(96) = "GT"
asTopLevels(97) = "GG"
asTopLevels(98) = "GN"
asTopLevels(99) = "GW"
asTopLevels(100) = "GY"
asTopLevels(101) = "HT"
asTopLevels(102) = "HM"
asTopLevels(103) = "VA"
asTopLevels(104) = "HN"
asTopLevels(105) = "HK"
asTopLevels(106) = "HU"
asTopLevels(107) = "IS"
asTopLevels(108) = "IN"
asTopLevels(109) = "ID"
asTopLevels(110) = "IR"
asTopLevels(111) = "IQ"
asTopLevels(112) = "IE"
asTopLevels(113) = "IM"
asTopLevels(114) = "IL"
asTopLevels(115) = "IT"
asTopLevels(116) = "JM"
asTopLevels(117) = "JP"
asTopLevels(118) = "JE"
asTopLevels(119) = "JO"
asTopLevels(120) = "KZ"
asTopLevels(121) = "KE"
asTopLevels(122) = "KI"
asTopLevels(123) = "KP"
asTopLevels(124) = "KR"
asTopLevels(125) = "KW"
asTopLevels(126) = "KG"
asTopLevels(127) = "LA"
asTopLevels(128) = "LV"
asTopLevels(129) = "LB"
asTopLevels(130) = "LS"
asTopLevels(131) = "LR"
asTopLevels(132) = "LY"
asTopLevels(133) = "LI"
asTopLevels(134) = "LT"
asTopLevels(135) = "LU"
asTopLevels(136) = "MO"
asTopLevels(137) = "MK"
asTopLevels(138) = "MG"
asTopLevels(139) = "MW"
asTopLevels(140) = "MY"
asTopLevels(141) = "MV"
asTopLevels(142) = "ML"
asTopLevels(143) = "MT"
asTopLevels(144) = "MH"
asTopLevels(145) = "MQ"
asTopLevels(146) = "MR"
asTopLevels(147) = "MU"
asTopLevels(148) = "YT"
asTopLevels(149) = "MX"
asTopLevels(150) = "FM"
asTopLevels(151) = "MD"
asTopLevels(152) = "MC"
asTopLevels(153) = "MN"
asTopLevels(154) = "MS"
asTopLevels(155) = "MA"
asTopLevels(156) = "MZ"
asTopLevels(157) = "MM"
asTopLevels(158) = "NA"
asTopLevels(159) = "NR"
asTopLevels(160) = "NP"
asTopLevels(161) = "NL"
asTopLevels(162) = "AN"
asTopLevels(163) = "NC"
asTopLevels(164) = "NZ"
asTopLevels(165) = "NI"
asTopLevels(166) = "NE"
asTopLevels(167) = "NG"
asTopLevels(168) = "NU"
asTopLevels(169) = "NF"
asTopLevels(170) = "MP"
asTopLevels(171) = "NO"
asTopLevels(172) = "OM"
asTopLevels(173) = "PK"
asTopLevels(174) = "PW"
asTopLevels(175) = "PA"
asTopLevels(176) = "PG"
asTopLevels(177) = "PY"
asTopLevels(178) = "PE"
asTopLevels(179) = "PH"
asTopLevels(180) = "PN"
asTopLevels(181) = "PL"
asTopLevels(182) = "PT"
asTopLevels(183) = "PR"
asTopLevels(184) = "QA"
asTopLevels(185) = "RE"
asTopLevels(186) = "RO"
asTopLevels(187) = "RU"
asTopLevels(188) = "RW"
asTopLevels(189) = "KN"
asTopLevels(190) = "LC"
asTopLevels(191) = "VC"
asTopLevels(192) = "WS"
asTopLevels(193) = "SM"
asTopLevels(194) = "ST"
asTopLevels(195) = "SA"
asTopLevels(196) = "SN"
asTopLevels(197) = "SC"
asTopLevels(198) = "SL"
asTopLevels(199) = "SG"
asTopLevels(200) = "SK"
asTopLevels(201) = "SI"
asTopLevels(202) = "SB"
asTopLevels(203) = "SO"
asTopLevels(204) = "ZA"
asTopLevels(205) = "GS"
asTopLevels(206) = "ES"
asTopLevels(207) = "LK"
asTopLevels(208) = "SH"
asTopLevels(209) = "PM"
asTopLevels(210) = "SD"
asTopLevels(211) = "SR"
asTopLevels(212) = "SJ"
asTopLevels(213) = "SZ"
asTopLevels(214) = "SE"
asTopLevels(215) = "CH"
asTopLevels(216) = "SY"
asTopLevels(217) = "TW"
asTopLevels(218) = "TJ"
asTopLevels(219) = "TZ"
asTopLevels(220) = "TH"
asTopLevels(221) = "TG"
asTopLevels(222) = "TK"
asTopLevels(223) = "TO"
asTopLevels(224) = "TT"
asTopLevels(225) = "TN"
asTopLevels(226) = "TR"
asTopLevels(227) = "TM"
asTopLevels(228) = "TC"
asTopLevels(229) = "TV"
asTopLevels(230) = "UG"
asTopLevels(231) = "UA"
asTopLevels(232) = "AE"
asTopLevels(233) = "GB"
asTopLevels(234) = "US"
asTopLevels(235) = "UM"
asTopLevels(236) = "UY"
asTopLevels(237) = "UZ"
asTopLevels(238) = "VU"
asTopLevels(239) = "VE"
asTopLevels(240) = "VN"
asTopLevels(241) = "VG"
asTopLevels(242) = "VI"
asTopLevels(243) = "WF"
asTopLevels(244) = "EH"
asTopLevels(245) = "YE"
asTopLevels(246) = "YU"
asTopLevels(247) = "ZR"
asTopLevels(248) = "ZM"
asTopLevels(249) = "ZW"
asTopLevels(250) = "UK"

For ictr = 0 To iNumDomains - 1
    If asTopLevels(ictr) = DomainString Then
        bAns = True
        Exit For
    End If
Next

General_Is_TopLevel_Domain = bAns

End Function

Public Function General_Distance_Get(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer) As Integer
General_Distance_Get = Abs(x1 - x2) + Abs(y1 - y2)
End Function

Public Function General_Windows_Is_2000XP() As Boolean

On Error GoTo ErrorHandler

Dim RetVal As Long

OSInfo.dwOSVersionInfoSize = Len(OSInfo)
RetVal = GetOSVersion(OSInfo)

If OSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT And OSInfo.dwMajorVersion >= 5 Then
    General_Windows_Is_2000XP = True
Else
    General_Windows_Is_2000XP = False
End If

Exit Function

ErrorHandler:
    General_Windows_Is_2000XP = False

End Function
