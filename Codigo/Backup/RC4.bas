Attribute VB_Name = "RC4_Calc"
#If UsarCrypto = 1 Then

'***************************************************************
'**    Funcion
Public Function RC4(inp As String, key As String) As String
Dim S(0 To 255) As Byte, k(0 To 255) As Byte, I As Long
Dim j As Long, temp As Byte, Y As Byte, t As Long, X As Long
Dim Outp As String

For I = 0 To 255
    S(I) = I
Next

j = 1
For I = 0 To 255
    If j > Len(key) Then j = 1
    k(I) = Asc(Mid$(key, j, 1))
    j = j + 1
Next I

j = 0
For I = 0 To 255
    j = (j + S(I) + k(I)) Mod 256
    temp = S(I)
    S(I) = S(j)
    S(j) = temp
Next I

I = 0
j = 0
For X = 1 To Len(inp)
    I = (I + 1) Mod 256
    j = (j + S(I)) Mod 256
    temp = S(I)
    S(I) = S(j)
    S(j) = temp
    t = (S(I) + (S(j) Mod 256)) Mod 256
    Y = S(t)
    
    Outp = Outp & Chr(Asc(Mid(inp, X, 1)) Xor Y)
Next
RC4 = Outp
End Function

#End If
