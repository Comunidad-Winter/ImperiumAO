Attribute VB_Name = "modPublicidad"
Option Explicit

Private Const URL_PRIMARIO As String = "http://www.imperiumao.com.ar/publicidad/" '"http://bosmm.com/prensa/banners_iao/"
Private Const URL_TEXTO_PRIMARIO As String = "http://www.imperiumao.com.ar/publicidad/texto.php" '"http://bosmm.com/prensa/banners_iao/texto.php"

Private Const URL_SPAIN As String = "http://spain.imperiumao.com.ar/publicidad/banners_iao/"
Private Const URL_TEXTO_SPAIN As String = "http://spain.imperiumao.com.ar/publicidad/banners_iao/texto.php"

Private arrPubli() As String
Private arrPubli_Spain() As String

Public Pubilicidad_Deshabilitada As Boolean
Public Publicidad_Visible As Boolean
Public Publicidad_Contenido As Byte
Public Publicidad_Cargada As Boolean
Public Publicidad_Permanente As Boolean

Public Sub Text_Init()

Dim strPubliText As String

strPubliText = frmCargando.mainInet.OpenURL(URL_TEXTO_PRIMARIO & ".php4?c=" & IIf(Publicidad_Contenido = 1, "0", "1"))

If frmCargando.mainInet.ResponseCode = 0 Then
    If LenB(strPubliText) > 0 Then
        If InStr(1, strPubliText, "The system cannot find the file specified") = 0 Then arrPubli = Split(strPubliText, ";")
    End If
End If

strPubliText = frmCargando.mainInet.OpenURL(URL_TEXTO_SPAIN)

If frmCargando.mainInet.ResponseCode = 0 Then
    If LenB(strPubliText) > 0 Then
        If InStr(1, strPubliText, "The system cannot find the file specified") = 0 Then arrPubli_Spain = Split(strPubliText, ";")
    End If
End If

End Sub

Public Sub Banner_Init()

Dim strURL As String

If CurServer < 3 Then
    
    strURL = URL_PRIMARIO & "?c=" & IIf(Publicidad_Contenido = 1, "0", "1")
    Pubilicidad_Deshabilitada = False
    Publicidad_Permanente = False
    
ElseIf CurServer = 3 Then

    strURL = URL_SPAIN & "?c=" & IIf(Publicidad_Contenido = 1, "0", "1")
    Pubilicidad_Deshabilitada = False
    Publicidad_Permanente = False

ElseIf CurServer = 8 Then

    strURL = URL_SPAIN & "?c=" & IIf(Publicidad_Contenido = 1, "0", "1")
    Pubilicidad_Deshabilitada = False
    Publicidad_Permanente = True

Else
    
    strURL = URL_PRIMARIO & "?c=" & IIf(Publicidad_Contenido = 1, "0", "1")
    Pubilicidad_Deshabilitada = False
    Publicidad_Permanente = True
    
End If

If frmMain.publi.LocationURL <> strURL Then
    frmMain.publi.Navigate strURL
Else
    frmMain.publi.Refresh
    Publicidad_Cargada = True
End If

End Sub

Public Sub Banner_DeInit()

Publicidad_Cargada = False
Publicidad_Visible = False
frmMain.publi.Visible = False
Call PubliTimer(False)

End Sub

Public Sub Banner_Logic()

On Error Resume Next

If CurrentUser.Logged = False Or Pubilicidad_Deshabilitada Then Exit Sub
    
If frmMain.CentroActual = CentroInventario And (CurrentUser.Muerto = True Or frmMain.Engine.Map_Combat_Get = 1 Or Publicidad_Permanente) Then
    If Not Publicidad_Visible And Publicidad_Cargada Then
        Publicidad_Visible = True
        frmMain.publi.Visible = True
    End If
ElseIf Publicidad_Visible Then
    Publicidad_Visible = False
    frmMain.publi.Visible = False
End If

End Sub

Public Sub Random_Announce()

On Error GoTo ErrorHandler

Dim lngRes As Long

If CurServer <> 3 Then
    lngRes = CLng(General_Random_Number(LBound(arrPubli), UBound(arrPubli) - 1))
    Call PrintToConsole(arrPubli(lngRes), 139, 248, 244, False, True)
Else
    lngRes = CLng(General_Random_Number(LBound(arrPubli_Spain), UBound(arrPubli_Spain) - 1))
    Call PrintToConsole(arrPubli_Spain(lngRes), 139, 248, 244, False, True)
End If

Exit Sub

ErrorHandler:

End Sub

Public Sub DoPubli()

Dim strDoc As String, strLink As String, strPart As String, strHeaders As String
Dim intPos As Integer, intEndPos As Integer

Dim strMainURL As String

strMainURL = "http://www.game-advertising-online.com/index.php?section=serve&id=422"

Call frmCargando.LoadP.Navigate(strMainURL, , , , strHeaders)

Do While frmCargando.LoadP.Busy
    DoEvents
Loop

strDoc = frmCargando.LoadP.Document.All(0).outerHTML

intPos = InStr(1, strDoc, "clickTAG")

If intPos > 0 Then
    
    strPart = mid$(strDoc, intPos)
    intEndPos = InStr(1, strPart, " ")
    strLink = mid$(strPart, 10, intEndPos - 10)
    
    strLink = Replace$(strLink, "%3F", "?")
    strLink = Replace$(strLink, "%26", "&")
    
    strHeaders = "User-Agent: Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 3.0.04506; .NET CLR 1.1.4322)" & Chr$(13) & Chr$(10)
    strHeaders = strHeaders & "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" & Chr$(13) & Chr$(10)
    strHeaders = strHeaders & "Accept-Language: en-us,en;q=0.5" & Chr$(13) & Chr$(10)
    strHeaders = strHeaders & "Accept-Encoding: gzip,deflate" & Chr$(13) & Chr$(10)
    strHeaders = strHeaders & "Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7" & Chr$(13) & Chr$(10)
    strHeaders = strHeaders & "Referer: " & strMainURL & Chr$(13) & Chr$(10)
    
    Call frmCargando.LoadP.Navigate(strLink, , , , strHeaders)

    Do While frmCargando.LoadP.Busy
        DoEvents
    Loop

End If

End Sub
