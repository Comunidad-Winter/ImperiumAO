Attribute VB_Name = "modConversiones"
Public Function Integer_To_String(ByVal Var As Integer) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    Dim temp As String
        
    'Convertimos a hexa
    temp = hex$(Var)
    
    'Nos aseguramos tenga 4 bytes de largo
    While Len(temp) < 4
        temp = "0" & temp
    Wend
    
    'Convertimos a string
    Integer_To_String = Chr$(Val("&H" & left$(temp, 2))) & Chr$(Val("&H" & Right$(temp, 2)))
Exit Function

errhandler:
End Function

Public Function String_To_Integer(ByVal str As String, ByVal start As Byte) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    On Error GoTo Error_Handler
    
    Dim temp_str As String
    
    'Asergurarse sea válido
    If Len(str) < start - 1 Then Exit Function
    
    'Convertimos a hexa el valor ascii del segundo byte
    temp_str = hex$(Asc(mid$(str, start + 1, 1)))
    
    'Nos aseguramos tenga 2 bytes (los ceros a la izquierda cuentan por ser el segundo byte)
    While Len(temp_str) < 2
        temp_str = "0" & temp_str
    Wend
    
    'Convertimos a integer
    String_To_Integer = Val("&H" & hex$(Asc(mid$(str, start, 1))) & temp_str)
            
    Exit Function
        
Error_Handler:
        
End Function

Public Function Byte_To_String(ByVal Var As Byte) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'Convierte un byte a string
'**************************************************************
    Byte_To_String = Chr$(Val("&H" & hex$(Var)))
Exit Function

errhandler:
End Function

Public Function String_To_Byte(ByVal str As String, ByVal start As Byte) As Byte
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    On Error GoTo Error_Handler
    
    If Len(str) < start Then Exit Function
    
    String_To_Byte = Asc(mid$(str, start, 1))
    
    Exit Function
        
Error_Handler:

End Function

Public Function Long_To_String(ByVal Var As Long) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    Dim temp As String
        
    'Convertimos a hexa
    temp = hex$(Var)
    
    'Nos aseguramos tenga 8 bytes de largo
    While Len(temp) < 8
        temp = "0" & temp
    Wend
    
    'Convertimos a string
    Long_To_String = Chr$(Val("&H" & left$(temp, 2))) & Chr$(Val("&H" & mid$(temp, 3, 2))) & Chr$(Val("&H" & mid$(temp, 5, 2))) & Chr$(Val("&H" & mid$(temp, 7, 2)))
Exit Function

errhandler:
End Function

Public Function String_To_Long(ByVal str As String, ByVal start As Byte) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    On Error GoTo Error_Handler
    
    If Len(str) < start - 3 Then Exit Function
    
    Dim temp_str As String
    Dim temp_str2 As String
    Dim temp_str3 As String
    
    'Tomamos los últimos 3 bytes y convertimos sus valroes ASCII a hexa
    temp_str = hex$(Asc(mid$(str, start + 1, 1)))
    temp_str2 = hex$(Asc(mid$(str, start + 2, 1)))
    temp_str3 = hex$(Asc(mid$(str, start + 3, 1)))
    
    'Nos aseguramos todos midan 2 bytes (los ceros a la izquierda cuentan por ser bytes 2, 3 y 4)
    While Len(temp_str) < 2
        temp_str = "0" & temp_str
    Wend
    
    While Len(temp_str2) < 2
        temp_str2 = "0" & temp_str2
    Wend
    
    While Len(temp_str3) < 2
        temp_str3 = "0" & temp_str3
    Wend
    
    'Convertimos a una única cadena hexa
    String_To_Long = Val("&H" & hex$(Asc(mid(str, start, 1))) & temp_str & temp_str2 & temp_str3)
            
    Exit Function
        
Error_Handler:

End Function
