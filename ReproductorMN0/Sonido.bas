Attribute VB_Name = "Sonido"
Public Duracion As Long

Dim Playing As Boolean
'Public Duracion As Single

Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Dim ElAlias As String

Sub Abrir(Archivo As String)
ElAlias = "1233"

Dim Ret As String * 128
Ret = mciSendString("Close  " + ElAlias + " ", 0, 0, 0)

Dim NombreCorto As String
Dim TMP As String * 255

NombreCorto = GetShortPathName(Archivo, TMP, 255)

NombreCorto = Left(TMP, NombreCorto)


'Cargar Video
Dim cmd As String
cmd = "open " + NombreCorto + " type MPEGVideo Alias " + ElAlias

Ret = mciSendString(cmd, 0&, 0&, 0&)

'Configura el Video en Milisegundos
r = mciSendString("set  " + ElAlias + "  time format milliseconds", 0, 0, 0)

Static D As String * 30
r = mciSendString("status " + ElAlias + " length", D, Len(D), 0)

Duracion = (Val(D)) - 1
End Sub

Sub vCerrar()
    Dim Ret As String * 128
    Ret = mciSendString("Close  " + ElAlias + " ", 0, 0, 0)
    Playing = False
End Sub

Sub vPlay()
    r = mciSendString("play  " + ElAlias, 0, 0, 0)
    Playing = True
End Sub

Sub vPausa()
    r = mciSendString("pause  " + ElAlias, 0, 0, 0)
    Playing = False
End Sub

Sub IrA(Segundo As Single)
    If Playing = False Then
        r = mciSendString("seek  " + ElAlias + "  to " + CStr(Segundo * 1000), 0, 0, 0)
    Else
        r = mciSendString("play  " + ElAlias + "  from " + CStr(Segundo * 1000), 0, 0, 0)
    End If
End Sub
