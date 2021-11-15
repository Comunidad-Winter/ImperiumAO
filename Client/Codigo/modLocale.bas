Attribute VB_Name = "modLocale"
Option Explicit

Private Type tNPCLocale
    strName As String
    strDialog As String
End Type

Private Type tSpellLocale
    strName As String
    strDesc As String
    strHechizeroMsg As String
    strTargetMsg As String
    strOwnMsg As String
End Type

Private Type tItemLocale
    strName As String
    strDesc As String
End Type

Private arrLocale_SMG() As String
Private arrLocale_GUI() As String
Private arrLocale_FACC() As String
Private arrLocale_Error() As String
Private arrLocale_ITEM() As tItemLocale
Private arrLocale_CMD() As String
Private arrLocale_NPC() As tNPCLocale
Private arrLocale_SPL() As tSpellLocale

Public Function Load_Locales(ByVal strLang As String) As Boolean

On Error GoTo ErrorHandler

Dim strFile As String
Dim tmpStr As String
Dim intFile As Integer
Dim i As Long

strFile = "locale_smg_" & strLang & ".ind"

If Extract_File(Scripts, App.Path & "\Recursos", strFile, Windows_Temp_Dir) Then

    strFile = Windows_Temp_Dir & strFile

    ReDim arrLocale_SMG(1 To General_Get_Line_Count(strFile)) As String
    
    intFile = FreeFile
    Open strFile For Input As #intFile
    
    Do While Not EOF(intFile)
        i = i + 1
        Line Input #intFile, arrLocale_SMG(i)
    Loop
    
    Close #intFile
    Delete_File strFile
Else
    Exit Function
End If
    
strFile = "locale_gui_" & strLang & ".ind"

If Extract_File(Scripts, App.Path & "\Recursos", strFile, Windows_Temp_Dir) Then

    strFile = Windows_Temp_Dir & strFile

    ReDim arrLocale_GUI(1 To General_Get_Line_Count(strFile)) As String
    
    intFile = FreeFile
    Open strFile For Input As #intFile
    
    i = 0
    
    Do While Not EOF(intFile)
        i = i + 1
        Line Input #intFile, arrLocale_GUI(i)
    Loop
    
    Close #intFile
    Delete_File strFile
Else
    Exit Function
End If

strFile = "locale_error_" & strLang & ".ind"

If Extract_File(Scripts, App.Path & "\Recursos", strFile, Windows_Temp_Dir) Then

    strFile = Windows_Temp_Dir & strFile

    ReDim arrLocale_Error(1 To General_Get_Line_Count(strFile)) As String
    
    intFile = FreeFile
    Open strFile For Input As #intFile
    
    i = 0
    
    Do While Not EOF(intFile)
        i = i + 1
        Line Input #intFile, arrLocale_Error(i)
    Loop
    
    Close #intFile
    Delete_File strFile
Else
    Exit Function
End If

strFile = "locale_facc_" & strLang & ".ind"

If Extract_File(Scripts, App.Path & "\Recursos", strFile, Windows_Temp_Dir) Then

    strFile = Windows_Temp_Dir & strFile

    ReDim arrLocale_FACC(1 To General_Get_Line_Count(strFile)) As String
    
    intFile = FreeFile
    Open strFile For Input As #intFile
    
    i = 0
    
    Do While Not EOF(intFile)
        i = i + 1
        Line Input #intFile, arrLocale_FACC(i)
    Loop
    
    Close #intFile
    Delete_File strFile
Else
    Exit Function
End If

strFile = "locale_obj_" & strLang & ".ind"

If Extract_File(Scripts, App.Path & "\Recursos", strFile, Windows_Temp_Dir) Then

    strFile = Windows_Temp_Dir & strFile

    ReDim arrLocale_ITEM(1 To General_Get_Line_Count(strFile)) As tItemLocale
    
    intFile = FreeFile
    Open strFile For Input As #intFile
    
    i = 0
    
    Do While Not EOF(intFile)
        i = i + 1
        Line Input #intFile, tmpStr
        arrLocale_ITEM(i).strName = General_Field_Read(1, tmpStr, "|")
        arrLocale_ITEM(i).strDesc = General_Field_Read(2, tmpStr, "|")
    Loop
    
    Close #intFile
    Delete_File strFile
Else
    Exit Function
End If

strFile = "locale_cmd_" & strLang & ".ind"

If Extract_File(Scripts, App.Path & "\Recursos", strFile, Windows_Temp_Dir) Then

    strFile = Windows_Temp_Dir & strFile

    ReDim arrLocale_CMD(1 To General_Get_Line_Count(strFile)) As String
    
    intFile = FreeFile
    Open strFile For Input As #intFile
    
    i = 0
    
    Do While Not EOF(intFile)
        i = i + 1
        Line Input #intFile, arrLocale_CMD(i)
    Loop
    
    Close #intFile
    Delete_File strFile
Else
    Exit Function
End If

strFile = "locale_npc_" & strLang & ".ind"

If Extract_File(Scripts, App.Path & "\Recursos", strFile, Windows_Temp_Dir) Then

    strFile = Windows_Temp_Dir & strFile

    ReDim arrLocale_NPC(1 To General_Get_Line_Count(strFile)) As tNPCLocale
    
    intFile = FreeFile
    Open strFile For Input As #intFile
    
    i = 0
    
    Do While Not EOF(intFile)
        i = i + 1
        Line Input #intFile, tmpStr
        arrLocale_NPC(i).strDialog = General_Field_Read(2, tmpStr, "|")
        arrLocale_NPC(i).strName = General_Field_Read(1, tmpStr, "|")
    Loop
    
    Close #intFile
    Delete_File strFile
Else
    Exit Function
End If

strFile = "locale_spl_" & strLang & ".ind"

If Extract_File(Scripts, App.Path & "\Recursos", strFile, Windows_Temp_Dir) Then

    strFile = Windows_Temp_Dir & strFile

    ReDim arrLocale_SPL(1 To General_Get_Line_Count(strFile)) As tSpellLocale
    
    intFile = FreeFile
    Open strFile For Input As #intFile
    
    i = 0
    
    Do While Not EOF(intFile)
        i = i + 1
        Line Input #intFile, tmpStr
        arrLocale_SPL(i).strName = General_Field_Read(1, tmpStr, "|")
        arrLocale_SPL(i).strDesc = General_Field_Read(2, tmpStr, "|")
        arrLocale_SPL(i).strHechizeroMsg = General_Field_Read(3, tmpStr, "|")
        arrLocale_SPL(i).strTargetMsg = General_Field_Read(4, tmpStr, "|")
        arrLocale_SPL(i).strOwnMsg = General_Field_Read(5, tmpStr, "|")
    Loop
    
    Close #intFile
    Delete_File strFile
Else
    Exit Function
End If

Load_Locales = True

Exit Function

ErrorHandler:

End Function

Public Function Locale_Facc_Frase(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

Locale_Facc_Frase = arrLocale_FACC(intInd)

Exit Function

ErrorHandler:

End Function

Public Function Locale_Error(ByVal btInd As Byte) As String

On Error GoTo ErrorHandler

Locale_Error = arrLocale_Error(btInd)

Exit Function

ErrorHandler:

End Function

Public Function Locale_GUI_Frase(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

Locale_GUI_Frase = arrLocale_GUI(intInd)

Exit Function

ErrorHandler:

End Function

Public Function Locale_UserItem(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

Locale_UserItem = arrLocale_ITEM(intInd).strName

Exit Function

ErrorHandler:

End Function

Public Function Locale_UserItem_Desc(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

Locale_UserItem_Desc = arrLocale_ITEM(intInd).strDesc

Exit Function

ErrorHandler:

End Function

Public Function Locale_NPC_Name(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

Locale_NPC_Name = arrLocale_NPC(intInd).strName

Exit Function

ErrorHandler:

End Function

Public Function Locale_Spell_Owner(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

Locale_Spell_Owner = arrLocale_SPL(intInd).strOwnMsg

Exit Function

ErrorHandler:

End Function

Public Function Locale_Spell_Target(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

Locale_Spell_Target = arrLocale_SPL(intInd).strTargetMsg

Exit Function

ErrorHandler:

End Function

Public Function Locale_Spell_Name(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

Locale_Spell_Name = arrLocale_SPL(intInd).strName

Exit Function

ErrorHandler:

End Function

Public Function Locale_Spell_Caster(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

Locale_Spell_Caster = arrLocale_SPL(intInd).strHechizeroMsg

Exit Function

ErrorHandler:

End Function

Public Function Locale_NPC_Dialog(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

Locale_NPC_Dialog = arrLocale_NPC(intInd).strDialog

Exit Function

ErrorHandler:

End Function

Public Function Locale_CMD_Get(ByVal intInd As Integer) As String

On Error GoTo ErrorHandler

Locale_CMD_Get = arrLocale_CMD(intInd)

Exit Function

ErrorHandler:

End Function

Public Function Locale_Parse_GUI(ByVal strParse As String) As String

On Error GoTo ErrorHandler

Dim lngPosFirst As Long
Dim strTemp As String

lngPosFirst = InStr(1, strParse, "$")

If lngPosFirst <= 0 Then
    Locale_Parse_GUI = strParse
    Exit Function
End If

If InStr(1, strParse, " ") Then

Else
    Locale_Parse_GUI = arrLocale_GUI(Val(mid$(strParse, lngPosFirst + 1)))
End If

Exit Function

ErrorHandler:
    Locale_Parse_GUI = strParse

End Function

Public Function Locale_Parse_ServerMessage(ByVal bytHeader As Byte, Optional ByVal strExtra As String = vbNullString) As String

On Error GoTo ErrorHandler

Dim strLocale As String
Dim lngPos As Long

If LenB(strExtra) = 0 Then
    Locale_Parse_ServerMessage = arrLocale_SMG(bytHeader)
    Exit Function
End If

strLocale = arrLocale_SMG(bytHeader)
lngPos = InStr(1, strLocale, "%N")

If lngPos > 0 Then
    Locale_Parse_ServerMessage = Replace$(strLocale, "%N", strExtra)
    Exit Function
End If

lngPos = InStr(1, strLocale, "¬")

Do While lngPos > 0
    strLocale = Replace$(strLocale, mid$(strLocale, lngPos, 2), String_To_Byte(strExtra, CByte(mid$(strLocale, lngPos + 1, 1))))
    lngPos = InStr(lngPos + 1, strLocale, "¬")
Loop

lngPos = InStr(1, strLocale, "#")

Do While lngPos > 0
    strLocale = Replace$(strLocale, mid$(strLocale, lngPos, 2), String_To_Long(strExtra, CByte(mid$(strLocale, lngPos + 1, 1))))
    lngPos = InStr(lngPos + 1, strLocale, "#")
Loop

lngPos = InStr(1, strLocale, "&")

Do While lngPos > 0
    strLocale = Replace$(strLocale, mid$(strLocale, lngPos, 2), Locale_UserItem(String_To_Integer(strExtra, CByte(mid$(strLocale, lngPos + 1, 1)))))
    lngPos = InStr(lngPos + 1, strLocale, "&")
Loop

lngPos = InStr(1, strLocale, "%")

If lngPos > 0 Then
    strLocale = Replace$(strLocale, mid$(strLocale, lngPos, 2), mid$(strExtra, CByte(mid$(strLocale, lngPos + 1, 1))))
End If

ErrorHandler:
    Locale_Parse_ServerMessage = strLocale

End Function
