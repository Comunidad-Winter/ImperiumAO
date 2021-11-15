VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmUpdater 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ImperiumAO - actualizaciónes automáticas"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5475
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdater.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   90
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   1
      Top             =   1020
      Width           =   5325
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buscando actualizaciónes..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   645
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   5235
   End
End
Attribute VB_Name = "frmUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CancelSearch As Boolean

Public Function FormatFileSize(ByVal dblFileSize As Double) As String

Select Case dblFileSize
    Case 0 To 999   ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1000 To 1023999    ' KB
        FormatFileSize = Format(dblFileSize / 1024, "##0.0") & " KB"
    Case 1024000 To (1024 * 10 ^ 6) - 1 ' MB
        FormatFileSize = Format(dblFileSize / (1024 ^ 2), "##0.0#") & " MB"
    Case Is > (1024 * 10 ^ 6)
        FormatFileSize = Format(dblFileSize / (1024 ^ 3), "##0.0#") & " GB"
End Select

End Function

Public Function FormatTime(ByVal sglTime As Single) As String
                           
Select Case sglTime
    Case 0 To 59
        FormatTime = Format(sglTime, "0") & " sec"
    Case 60 To 3599
        FormatTime = Format(Int(sglTime / 60), "#0") & _
                     " min " & _
                     Format(sglTime Mod 60, "0") & " sec"
    Case Else
        FormatTime = Format(Int(sglTime / 3600), "#0") & _
                     " hr " & _
                     Format(sglTime / 60 Mod 60, "0") & " min"
End Select

End Function

Public Function ReturnFileOrFolder(FullPath As String, _
                                   ReturnFile As Boolean, _
                                   Optional IsURL As Boolean = False) _
                                   As String

Dim intDelimiterIndex As Integer

intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
ReturnFileOrFolder = IIf(ReturnFile, _
                         Right(FullPath, Len(FullPath) - intDelimiterIndex), _
                         Left(FullPath, intDelimiterIndex))

End Function


Public Function DownloadFile(strURL As String, strDestination As String, Optional UserName As String = "", Optional Password As String = "", Optional TACCESO As AccessConstants = icUseDefault, Optional PROXY As String = "") As Boolean

Dim bData() As Byte         ' Data var
Dim intFile As Integer      ' FreeFile var
Dim a As Variant            ' Temp var
Dim intKReceived As Integer ' KB received so far
Dim intKFileLength As Long  ' KB total length of file
Dim lastTime As Single      ' time last chunk received
Dim sglRate As Single       ' var to hold transfer rate
Dim sglTime As Single       ' var to hold time remaining
Dim strFile As String       ' temp filename var
Dim strHeader As String     ' HTTP header store
Dim strHost As String       ' HTTP Host
Dim bz As Long
Dim bzd As Integer
On Local Error GoTo InternetErrorHandler

CancelSearch = False
BotonCancel = False
strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, False, True)

Me.Show
'StatusLabel2 = "Reciviendo la información del archivo..."
DoEvents
With Inet1
    .AccessType = TACCESO
    .PROXY = PROXY
    .Url = strURL
    .UserName = UserName
    .Password = Password
    .Execute , "GET"
End With

'StatusLabel = "Guardando:" & vbCr & vbCr & strFile & " desde " _
              & IIf(Len(strHost) < 33, strHost, "..." & Left(strHost, 30))

lastTime = Timer

While Inet1.StillExecuting
    DoEvents
    If CancelSearch Then
        GoTo ExitDownload
    End If
Wend

strHeader = Inet1.GetHeader

Select Case Mid(strHeader, 10, 3)
    Case "200"  ' OK!

    Case "401"  ' Not authorized
        AddToText2 "No hay autorización para la descarga del archivo"
        GoTo ExitDownload
    
    Case "404"  ' File Not Found
        AddToText2 "El archivo, " & _
               Inet1.Url & _
               " no pudo ser encontrado!"
        GoTo ExitDownload
        
    Case vbCrLf
        AddToText2 "No pude establecer conexión"
        GoTo ExitDownload
        
    Case Else
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        AddToText2 "Respuesta del server:" & vbCr & vbCr & _
               strHeader
        GoTo ExitDownload
End Select

strHeader = Inet1.GetHeader("Content-Length")
intKFileLength = CLng(Val(strHeader))
If intKFileLength = 0 Then
    GoTo ExitDownload
End If

'hay que hacere esto porque si el archivo es menor que la primer llamada me caga el codigo :P
If intKFileLength > 1024 Then 'velocidad normal
    bz = 1024
    bzd = 1
ElseIf intKFileLength > 256 Then 'lento
    bz = 256
    bzd = 4
Else 'MUY lento
    bz = 16
    bzd = 64
End If
intKFileLength = CInt(Val(strHeader) / 1024)

DoEvents

intKReceived = 0

On Local Error GoTo FileErrorHandler

If Inet1.ResponseCode = 0 Then
    intFile = FreeFile
    Open strDestination For Binary Access Write As #intFile
    bData = Inet1.GetChunk(bz, icByteArray)
    a = bData
        Do While LenB(a) > 0
        Put #intFile, , bData
        bData = Inet1.GetChunk(bz, icByteArray)
        a = bData
        If CancelSearch Then
            Close #intFile
            Kill strDestination
            GoTo ExitDownload
        End If
        intKReceived = intKReceived + 1
        If (intKReceived / bzd) < intKFileLength Then
            sglRate = (intKReceived / bzd) / (Timer - lastTime)
            sglTime = (intKFileLength - (intKReceived / bzd)) / sglRate
            'StatusLabel2 = "Tiempo restante estimado: " & _
                           FormatTime(sglTime) & _
                           " (" & _
                           FormatFileSize((intKReceived / bzd) * 1024#) & _
                           " de " & _
                           FormatFileSize(intKFileLength * 1024#) & _
                           " copiado)" & vbCr & vbCr & _
                           "Velocidad: " & _
                           Format(sglRate, "###,##0.0") & " KB/Sec"
            'ProgressBar1.value = intKReceived
            Caption = Format(((intKReceived / bzd) / intKFileLength), "##0%") & _
                      " de " & strFile & " completado"
        End If
    DoEvents
    Loop
    Put #intFile, , bData
    Close #intFile
End If

DoEvents

ExitDownload:

If bzd = 0 Then
'no se encontro archivo
    DownloadSuccess = False
ElseIf (intKReceived / bzd) >= intKFileLength Then
    DownloadSuccess = True
Else ' SINUHE ARREGLAME ESTO!!! PUTO!!! ME CAGA TODO EL UPDATE
    DownloadSuccess = False
    If Not Dir(strDestination) = Empty Then Kill strDestination
    If Not CancelSearch Then
        'StatusLabel = "Descarga fallada!"
        MsgBox "Descarga fallada!", _
        vbCritical, _
        "Error descargando el archivo"
    End If
End If

Inet1.Cancel
DoEvents
Unload Me
DoEvents

On Local Error GoTo 0

Exit Function

InternetErrorHandler:
    CancelSearch = True
    Inet1.Cancel
    AddToText "Error: " & Err.Description & " ocurrido."
    DoEvents
    DownloadSuccess = False
    BotonCancel = True
    Resume Next
    
FileErrorHandler:
    MsgBox "No pude escribir el archivo!", _
           vbCritical, _
           "Error de escritura"
    DownloadSuccess = False
    BotonCancel = True
    Resume Next
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    CancelSearch = True
End Sub
