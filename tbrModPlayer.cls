VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim RET As String * 128 'lo uso en todos lados

Private WithEvents Reloj0 As tbrTimer.clsTimer
Attribute Reloj0.VB_VarHelpID = -1
Private WithEvents Reloj1 As tbrTimer.clsTimer
Attribute Reloj1.VB_VarHelpID = -1
Private WithEvents Reloj2 As tbrTimer.clsTimer
Attribute Reloj2.VB_VarHelpID = -1
Private WithEvents Reloj3 As tbrTimer.clsTimer
Attribute Reloj3.VB_VarHelpID = -1
Private WithEvents RelojKar As tbrTimer.clsTimer
Attribute RelojKar.VB_VarHelpID = -1

Private Terr As New tbrErrores.clsTbrERR

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long

Private Type RECT
        left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const WS_CHILD = &H40000000

'Property Variables:
Dim m_FileName(4) As String 'ahora son 4 (para kundera 6.8.200)
Dim m_Volumen(4) As Long 'ahora son 4 (para kundera 6.8.200)
Dim dwReturn As Long

Private Alias(4) As String
'lista de alias para usar en enganches o similares

Public Event Played(SecondsPlayed As Long, iAlias As Long, MS As Long)
Public Event BeginPlay(iAlias As Long)
Public Event EndPlay(iAlias As Long)
Public Event mmError(txtMasHist As String)
Public Event FaltaNextEvKAR(dMiliSec As Double)

Private TMPs As String
Private mTotalSec(3) As Single
Private mTotalMiliSec(3) As Single 'nuevo mayo 07 para mejor movento de volumen
Private mTotalFrames(3) As Single
Private mFramePerSecond(3) As Single
Private mFramePerMiliSecond(3) As Single 'nuevo mayo 07 para mejor movento de volumen
Private mSecActual(3) As Single
Private mMiliSecActual(3) As Single 'nuevo mayo 07 para mejor movento de volumen
Private mHastaTema(3) As Long 'al abrir a DLL el total tema(3) que decia _
    cuanto duraba (y lo cortaba si se cancelaba una cancion u otro) o sea _
    en general es la duracion real en segundos pero puede ser menos para _
    cortarlo antes.

Private mEsGratis(3) As Boolean

Private FSO As New Scripting.FileSystemObject

Private KAR As New tbrPlayMN0
Private TEM As New tbrSoftEncrMan.clstbrENC

Private MilliSecKarIni As Double
Private MilliSecKar As Double
Private mIntervalKar As Long 'intervalo para leer eventos (muy bajo jode en PCs viejas)

Public Sub doTem(i As Boolean, cl As String, ai As String, ao As String, Optional px As String = "") 'encripta o desencripta un archivo
    TEM.Encriptar i, cl, ai, ao, px
    'aparentemente graba si o si el archivo hay que ver si sirve despues
End Sub

Public Function GetPrefixKar(AR As String, nLEN As Long) As String
    'devuelve los primeros caracteres de un karaoke encriptado (AR) para saber
    'con clave abrirlo
    Dim Aux As String
    
    If Dir(AR) = "" Then Exit Function
    Aux = Space(nLEN)
    Open AR For Binary As #1
        Get #1, 1, Aux
    Close #1
    
    GetPrefixKar = Aux
End Function

'no se puede usar matriz en WithEvents !!!
Private Sub Reloj0_Timer(): Reloj_Timer 0: End Sub
Private Sub Reloj1_Timer(): Reloj_Timer 1: End Sub
Private Sub Reloj2_Timer(): Reloj_Timer 2: End Sub
Private Sub Reloj3_Timer(): Reloj_Timer 3: End Sub

Public Function GetCurrentMultimediaPos(iAlias As Long) As Long
    'se refiere a frame actual
    
    Dim pos As String * 128
    dwReturn = mciSendString("status " & Alias(iAlias) & " position", pos, 128, 0&)
    Terr.Anotar "002-0001s", iAlias, Alias(iAlias), dwReturn
    If Not dwReturn = 0 Then  'not success
        GetCurrentMultimediaPos = -1
        Exit Function
    End If
    
    'Success
    GetCurrentMultimediaPos = CLng(pos)
End Function

Public Function GetTotalframes(iAlias As Long) As Long
    
    Dim Total As String * 128
    
    dwReturn = mciSendString("status " & Alias(iAlias) & " length", Total, 128, 0&)
    
    If Not dwReturn = 0 Then  'not success
        mTotalFrames(iAlias) = -1
        GetTotalframes = -1
        Exit Function
    End If
    
    'Success
    mTotalFrames(iAlias) = CLng(Total)
    GetTotalframes = mTotalFrames(iAlias)
End Function

Public Property Get Volumen(iAlias As Long) As Long
    Terr.Anotar "002-0011", iAlias, m_Volumen(iAlias)
    Volumen = m_Volumen(iAlias) / 10
End Property

Public Property Let Volumen(iAlias As Long, ByVal New_Volumen As Long)
    On Error GoTo ERmp3
    'en mi máquina anda del 0 al 1000 (en todas)
    Terr.Anotar "002-0013"
    
    'asegurarse que no pida ni negativos ni mayores a 100
    If New_Volumen < 0 Then New_Volumen = 0
    If New_Volumen > 100 Then New_Volumen = 100
    
    m_Volumen(iAlias) = New_Volumen * 10
    TMPs = "SetAudio " + Alias(iAlias) + " Volume To " + CStr(m_Volumen(iAlias))
    Terr.Anotar "002-0014", iAlias, TMPs
    RET = mciSendString(TMPs, 0&, 0&, 0&)
    Terr.Anotar "002-0015", RET
    If RET <> 0 Then
        LogErrorMCI RET
        'no se pudo modificar el volumen
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoVolumen" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Terr.Anotar "002-0018"
    Exit Property
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Property

Public Property Get FileName(iAlias As Long) As String
    FileName = m_FileName(iAlias)
End Property

Public Property Let FileName(iAlias As Long, ByVal New_FileName As String)
    m_FileName(iAlias) = New_FileName
End Property

Public Function isPlayingAny() As Boolean
    Dim TmpB As Boolean
    TmpB = False
    Dim A44 As Long
    For A44 = 0 To 3
        If IsPlaying(A44) Then
            TmpB = True
            Exit For
        End If
    Next A44
    isPlayingAny = TmpB
End Function

Public Function isPlayingClock(iAlias As Long) As Boolean
    isPlayingClock = (RELOJ(iAlias).Interval > 0)
End Function

Public Function IsPlaying(iAlias As Long) As Boolean
    Dim S As String * 128
    On Error GoTo ERmp3
    Terr.Anotar "002-0027"
    If m_FileName(iAlias) = "" Then
        Terr.Anotar "002-0028", iAlias
        IsPlaying = False
    Else
        Terr.Anotar "002-0030"
        RET = mciSendString("status " + Alias(iAlias) + " mode", S, 128, 0&)
        Terr.Anotar "002-0031", RET, iAlias
        If RET = 263 Then '263 es cuando no ha abierto nada
            IsPlaying = False
            Exit Function
        End If
        If RET <> 0 Then
            LogErrorMCI RET
            'no se pudo modificar el volumen
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "IsplayStat" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
            'WriteLog "No se pudo definir el estado de ejecucion." + ". Tema: " + m_FileName + " Function IsPlaying", False
        End If
        'EN ESTE CASO ES NULO O ALGO ASI
        'YA QUE MCI NO TIENE LA CAPACIDAD DE STATUS!!!
        If RET = 274 Then
            Terr.Anotar "002-0034b", iAlias
            IsPlaying = True
        Else
            Terr.Anotar "002-0034", S
            IsPlaying = (Mid(S, 1, 7) = "playing")
        End If
        
        IsPlaying = (Mid(S, 1, 7) = "playing")
    End If
    
    Exit Function
ERmp3:
    Terr.Anotar "002-0034b"
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoOpen(iAlias As Long) As Long
    
    'devuelve
        '0: ok
        '1: no existe el archivo
        '2: WINDOWS NO REPRODUCE MP3 !
        '3 al mandar el Mci Open fallo !
    
    On Error GoTo ERmp3
    MiniCerrar iAlias
    'ver si esta el archivo
    Terr.Anotar "002-0041"
    If FSO.FileExists(m_FileName(iAlias)) = False Then
        RaiseEvent mmError("****ERRmmPlayer 1009****" + vbCrLf + _
                        "NoFile" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        DoOpen = 1
        Exit Function
    End If
    
    Dim Puto277 As Boolean
    Puto277 = False
Otro277:
    Terr.Anotar "002-0040", Puto277
    
    Dim lenShort As Long, TMP As String * 255
    lenShort = GetShortPathName(m_FileName(iAlias), TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacios el reproductor no anda.
    'JOYA JOYA JOYA
    
    'parece que los nombres lasrgos de carpetas o de archivos o del path completo _
    esta limitado por en cantidad por una variable MAX_PATH que vale distinto segun _
    el sistema operativo o de archivos
    'el texto dice que esto se puede omitir poniendo al inicio
    '\\?\
    
    'File Name Conventions
    '.....
    'Process a path as a null-terminated string. The maximum length for a path, _
    including a trailing backslash, is given by MAX_PATH.
    'The Unicode versions of several functions permit paths that exceed the _
    MAX_PATH length if the path has the "\\?\" prefix. The "\\?\" tells the _
    function to turn off path parsing. However, each component in the path _
    cannot be more than MAX_PATH characters long. Use the "\\?\" prefix with _
    paths for local storage devices and the "\\?\UNC\" prefix with paths having _
    the Universal Naming Convention (UNC) format. The "\\?\" is ignored as part _
    of the path. For example, "\\?\C:\myworld\private" is seen as _
    "C:\myworld\private", and "\\?\UNC\bill_g_1\hotstuff\coolapps" is seen _
    as "\\bill_g_1\hotstuff\coolapps".

    'devido a que a algunos clientes le da el error mci 377 que se define como error_interno
    'o sea no tiene descripcion clara. Es posible que seas por como corta el nombre
    'por ejemplo la funcion para
    'C:\Archivos de programa\3PM\discos\RICARDO ARJORA - Quien Dijo Ayer (2007)\CD 1 - 10 - Mujeres.mp3
    'devuelve
    'C:\ARCHIV~1\3PM\discos\RI8C71~1\CD1-10~1.MP3
    'ademas fallo al hacele el open
    'que no tiene nada que ver, me gustaria que hubeira devuelto
    'C:\ARCHIV~1\3PM\discos\RICAR~1\CD1-10~1.MP3
    
    'viendo esto podria analizar la respuesta
    'y si me deja dudas copiar el archivo a c:\bosta.mp3 o la extencion que sea
    'y mierda que no va a reproducir
    
    Terr.Anotar "002-0045", m_FileName(iAlias)
    
    Dim FileNameSHORT As String
    FileNameSHORT = left$(TMP, lenShort)
    
    Terr.Anotar "002-0046", FileNameSHORT, "Verifica igual 11-2-2017 ver 3.6.223"
    
    If m_FileName(iAlias) = FileNameSHORT Then
        Terr.Anotar "002-0046gg", "Parece que no se acorto la URL!!!"
        ' probar agregando comillas al PATH
        FileNameSHORT = Chr(34) & m_FileName(iAlias) & Chr(34)
    End If
    
    Dim cmdToDo As String * 255
    
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias " + _
        Alias(iAlias) + " style " & WS_CHILD
    
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0049", dwReturn, cmdToDo
    
    If dwReturn <> 0 Then
        If dwReturn = 263 Then
            'si da el error 263 es probable que la máquina no tenga MCI, lo que le paso a Mauro con W98 PE y a efren con ME
            RaiseEvent mmError("****ERRmmPlayer 1001****" + vbCrLf + _
                        "WINDOWS NO REPRODUCE MP3!!!" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
            DoOpen = 2
            Exit Function
        End If
        
        If dwReturn = 277 Then
            'es un error que ni gates sabe que es por eso lo llaman "error interno"
            'uno de los motivos puede ser que el getShoertPath joda.
            'por eso si llego aqui hago una prueba
            If Puto277 = False Then
                Puto277 = True
                Dim WA As String
                WA = App.Path + "\Musica.mpx" 'antes era "C:\" y no se si existe !!!
                If FSO.FileExists(WA) Then FSO.DeleteFile WA, True
                FSO.CopyFile m_FileName(iAlias), WA, True
                m_FileName(iAlias) = WA
                GoTo Otro277
            End If
        
        End If
        LogErrorMCI dwReturn
        'no se puedo abrir!!!
        RaiseEvent mmError("****ERRmmPlayer 1002****" + vbCrLf + _
                        "NoAbre" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        DoOpen = 3
        Exit Function
    Else
        'esta todo ok
        DoOpen = 0
        If Puto277 Then
            'quiere decir que lo solucione la cocha de la lora que error putooooo!!!!
            Terr.Anotar "PMQLP1212"
            Terr.AppendLog "PMQLP1212" 'PutaMadreQueLoPario
        End If
    End If
    
    Terr.Anotar "002-0054"
    
    MiniInicVarMci iAlias
    
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer 1003****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoOpenVideo(Style As String, HWind As Long, _
    X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, iAlias As Long)
    
    'devuelve
        '0: ok
        '1: no existe el archivo
        
        '3 al mandar el Mci Open fallo !
        '4: MCIERR_NO_WINDOW
        '5: otros errores <> 4 que se presentan al pegar el video a una ventana
    
    On Error GoTo ERmp3
    
    Terr.Anotar "002-0058"
    
    MiniCerrar iAlias
    
    Dim cmdToDo As String * 255
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    If Dir(m_FileName(iAlias)) = "" Then
        Terr.Anotar "002-0065b"
        RaiseEvent mmError("****ERRmmPlayer 1004****" + vbCrLf + _
                        "NoFile" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        DoOpenVideo = 1
        Exit Function
    End If
    Terr.Anotar "002-0068"
    lenShort = GetShortPathName(m_FileName(iAlias), TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    Terr.Anotar "002-0069", m_FileName(iAlias)
    'volu = mciGetDeviceID(lenShort)
    FileNameSHORT = left$(TMP, lenShort)
    Terr.Anotar "002-0070", FileNameSHORT
    
    If m_FileName(iAlias) = FileNameSHORT Then
        Terr.Anotar "002-0070gag", "Parece que no se acorto la URL!!!"
        ' probar agregando comillas al PATH
        FileNameSHORT = Chr(34) & m_FileName(iAlias) & Chr(34)
    End If
    
    Terr.Anotar "002-0071", HWind
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias " + _
        Alias(iAlias) + " parent " + CStr(HWind) + " style " & WS_CHILD 'Style
        
    Terr.Anotar "002-0072", cmdToDo '                         xxxxxx
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0073", dwReturn, Style, HWind
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer 1005****" + vbCrLf + _
                        "NoAbreVid" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        DoOpenVideo = 2
        Exit Function 'agregado 15/4/08 ??? antes no salia!!
    End If
    Terr.Anotar "002-0076", X1, X2, Y1, Y2
    
    'por si mando ancho o alto en cero!!
    If X2 = 0 Or Y2 = 0 Then
        'Get Window Size
        Dim rec As RECT
        Call GetWindowRect(HWind, rec)
        X2 = rec.Right - rec.left
        Y2 = rec.Bottom - rec.top
    End If

    cmdToDo = "put " + Alias(iAlias) + " window at " + CStr(X1) + " " + CStr(Y1) + _
        " " + CStr(X2) + " " + CStr(Y2)
    
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0077", Style, HWind, dwReturn
    
    If dwReturn <> 0 Then
        '********************
        'si es 346 es que no tiene ventana de presentacion (base+90=MCIERR_NO_WINDOW)
        '¿¿¿¿¿????????
        'probe con style popup y overlapped y no sirven ni solucionan
        If dwReturn = 346 Then
            RaiseEvent mmError("****ERRmmPlayer 1006****" + vbCrLf + _
                        "MCIERR_NO_WINDOW=346" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
            'pasa con videos con codecs nuevos que al cargarse si funciona en
            'WMP pero no en 3PM!!!!!!!!!!!!!!!!!!!
            DoOpenVideo = 4
            Exit Function
        '********************
        Else
            LogErrorMCI dwReturn
            
            RaiseEvent mmError("****ERRmmPlayer 1007****" + vbCrLf + _
                        "NoAbreVid2" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
            DoOpenVideo = 5
            'no salgo por que no se si es grabe o no
        End If
    Else
        DoOpenVideo = 0
    End If
    
    
    MiniInicVarMci iAlias
    
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer 1008****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Private Function GetFramesPerSecond(iAlias As Long) As Long
    mFramePerSecond(iAlias) = mTotalFrames(iAlias) / (mTotalSec(iAlias))
    mFramePerMiliSecond(iAlias) = mTotalFrames(iAlias) / (mTotalMiliSec(iAlias))
    GetFramesPerSecond = mFramePerSecond(iAlias)
End Function

Public Function DoPlay(iAlias As Long, Optional FullScreen As Boolean = False) As Long
    '0 ok
    '1 error mci
    
    Terr.AppendSinHist "py_purple"
    
    On Error GoTo ERmp3
    Terr.Anotar "002-0082", CStr(FullScreen), iAlias
    If FullScreen Then
        'dwReturn = mciSendString("play " + Alias(iAlias) + " fullscreen from 0 to " + CStr(mTotalFrames(iAlias)), 0&, 0&, 0&)
        'dwReturn = mciSendString("play " + Alias(iAlias) + " fullscreen from 0", 0&, 0&, 0&)
        dwReturn = mciSendString("play " + Alias(iAlias) + " fullscreen", 0&, 0&, 0&)
    Else
        'dwReturn = mciSendString("play " + Alias(iAlias) + " from 0 to " + CStr(mTotalFrames(iAlias)), 0&, 0&, 0&)
        'dwReturn = mciSendString("play " + Alias(iAlias) + " from 0", 0&, 0&, 0&)
        dwReturn = mciSendString("play " + Alias(iAlias), 0&, 0&, 0&)
    End If
    If dwReturn <> 0 Then
        DoPlay = 1
        LogErrorMCI dwReturn
        
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoPlay" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        Exit Function
    End If
    Terr.Anotar "002-0086", dwReturn
    RELOJ(iAlias).Interval = 60 'puede variar
    RELOJ(iAlias).Enabled = True
    Terr.Anotar "002-0087"
    DoPlay = 0
    RaiseEvent BeginPlay(iAlias)
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoPause(iAlias As Long)
    On Error GoTo ERmp3
    Terr.Anotar "002-0088"
    dwReturn = mciSendString("pause " + Alias(iAlias), 0&, 0&, 0&)
    Terr.Anotar "002-0089"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoPause" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Terr.Anotar "002-0092"
    RELOJ(iAlias).Interval = 0
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoStop(iAlias As Long) As String
    On Error GoTo ERmp3
    Terr.Anotar "002-0093", iAlias
    dwReturn = mciSendString("stop " + Alias(iAlias), 0&, 0&, 0&)
    If dwReturn <> 0 And dwReturn <> 263 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoStop" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Terr.Anotar "002-0097"
    RELOJ(iAlias).Interval = 0
    Terr.Anotar "002-0098"
    RaiseEvent EndPlay(iAlias)
    
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoClose(iAlias As Long) As String
    'iAlias es el que hay que cerrar o:
        '99 cierra todos
        
    On Error GoTo ERmp3
    Terr.Anotar "002-0099", iAlias
    If iAlias = 99 Then 'cierra todos
        If Terr.ModoGrabaTodo Then Terr.StopGrabaTodo 'dejo de grabar!!!
        
        Dim F11 As Long
        For F11 = 0 To 3
            Terr.Anotar "002-0099za", F11
            dwReturn = mciSendString("stop " + Alias(F11), 0&, 0&, 0&)
            RELOJ(F11).Interval = 0
        Next F11
        
        'limpiar el kar !
        RelojKar.Interval = 0: RelojKar.Enabled = False: KAR.BorraTemp
        dwReturn = mciSendString("close all", 0&, 0&, 0&)
        
        If dwReturn <> 0 And dwReturn <> 263 Then '263 ES CUANDO NO HAY NADA ABIERTO
            LogErrorMCI dwReturn
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoClose" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        End If
        
        
        
    Else 'o solo el elegido
        dwReturn = mciSendString("close " + Alias(iAlias), 0&, 0&, 0&)
        If dwReturn <> 0 And dwReturn <> 263 Then '263 ES CUANDO NO HAY NADA ABIERTO
            LogErrorMCI dwReturn
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoClose2" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        End If
        Terr.Anotar "002-0103"
        RELOJ(iAlias).Interval = 0
        
        'reloj especial del karaoke!
        If iAlias = 2 Then
            RelojKar.Interval = 0: RelojKar.Enabled = False
            KAR.BorraTemp
        End If
            
    End If
    'SI SIGUE EL RELOJ SE MARCAN 1000 errores!!!!!!!!!!
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function PercentPlay(iAlias As Long)
    On Error GoTo ERmp3
    Terr.Anotar "002-0104"
    PercentPlay = PositionInSec(iAlias) / mTotalSec(iAlias) * 100
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function PositionInSec(iAlias As Long) As Long
    PositionInSec = mSecActual(iAlias)
End Function

Public Function FaltaInSec(iAlias As Long)
    On Error GoTo ERmp3
    Terr.Anotar "002-0130"
    FaltaInSec = mTotalSec(iAlias) - mSecActual(iAlias)
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function Position(iAlias As Long) As String
    On Error GoTo ERmp3
    Dim Mins As Long, Sec As Long
    Terr.Anotar "002-0131"
    Sec = mSecActual(iAlias)
    Terr.Anotar "002-0132r", Sec, iAlias
    If Sec < 60 Then Position = "00:" & Format(Sec, "00")
    If Sec > 59 Then
        Terr.Anotar "002-0134r"
        Mins = Int(Sec / 60)
        Terr.Anotar "002-0135r", Mins
        Sec = Sec - (Mins * 60)
        Terr.Anotar "002-0136r", Sec
        Position = Format(Mins, "00") & ":" & Format(Sec, "00")
    End If
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function Falta(iAlias As Long) As String
    On Error GoTo ERmp3
    Dim Mins As Long, Sec As Long
    Terr.Anotar "002-0131"
    Sec = FaltaInSec(iAlias)
    Terr.Anotar "002-0132", Sec, iAlias
    If Sec < 60 Then Falta = "00:" & Format(Sec, "00")
    If Sec > 59 Then
        Terr.Anotar "002-0134"
        Mins = Int(Sec / 60)
        Terr.Anotar "002-0135", Mins
        Sec = Sec - (Mins * 60)
        Terr.Anotar "002-0136", Sec
        Falta = Format(Mins, "00") & ":" & Format(Sec, "00")
    End If
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function LengthInSec(iAlias As Long) As Long
    Terr.Anotar "002-0137"
    LengthInSec = CLng(mTotalSec(iAlias))
End Function

Public Function Length(iAlias As Long) As String
    Dim Sec As Long, Mins As Long
    On Error GoTo ERmp3
    Terr.Anotar "002-0148"
    Sec = mTotalSec(iAlias)
    Terr.Anotar "002-0149"
    If Sec < 60 Then Length = "0:" & Format(Sec, "00")
    Terr.Anotar "002-0150"
    If Sec > 59 Then
        Terr.Anotar "002-0151"
        Mins = Int(Sec / 60)
        Terr.Anotar "002-0152"
        Sec = Sec - (Mins * 60)
        Terr.Anotar "002-0153"
        Length = Format(Mins, "00") & ":" & Format(Sec, "00")
    End If
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function SeekTo(Second, iAlias As Long)
    On Error GoTo ERmp3
    Terr.Anotar "002-0154", iAlias
    
    'transformar segundos a frames
    Dim STF As Long
    STF = (Second / 1000) * mFramePerSecond(iAlias)
    
    If IsPlaying(iAlias) Then
        Terr.Anotar "002-0155"
        dwReturn = mciSendString("play " + Alias(iAlias) + " from " & STF, 0&, 0&, 0&)
        Terr.Anotar "002-0156"
        If dwReturn <> 0 And dwReturn <> 282 Then '282 es que pide un lugar de tiempo que no existe!
            LogErrorMCI dwReturn
            Terr.Anotar "002-0158"
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoSeek" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        End If
    Else
        Terr.Anotar "002-0159"
        dwReturn = mciSendString("seek " + Alias(iAlias) + " to " & STF, 0&, 0&, 0&)
        If dwReturn <> 0 And dwReturn <> 263 And dwReturn <> 282 Then
            LogErrorMCI dwReturn
            RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoSeek2" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        End If
    End If
    
    If iAlias = 2 Then ResincroKar
    
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Function Record() 'no se como hara con los alias, estimo que graba todo ¿¿??
    On Error GoTo ERmp3
    Terr.Anotar "002-0162"
    dwReturn = mciSendString("Close MP3rec", 0&, 0&, 0&)
    Terr.Anotar "002-0163"
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        Terr.Anotar "002-0164"
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        Terr.Anotar "002-0165"
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoRecord" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Terr.Anotar "002-0166"
    Dim cmdToDo As String * 255
    Terr.Anotar "002-0167"
    'abrir nuevo
    Terr.Anotar "002-0168"
    cmdToDo = "open new type WaveAudio Alias MP3rec"
    Terr.Anotar "002-0169"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoRecord2" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        Exit Function
    End If
    'iniciar grabacion
    Terr.Anotar "002-0174"
    cmdToDo = "record MP3rec"
    Terr.Anotar "002-0175"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoRec3" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Function StopRecord()
    On Error GoTo ERmp3
    Dim cmdToDo As String * 255
    Terr.Anotar "002-0178"
    'parar nuevo
    cmdToDo = "stop MP3rec"
    Terr.Anotar "002-0179"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0180"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoRec4" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    'grabar grabacion
    Terr.Anotar "002-0182"
    cmdToDo = "save MP3rec c:\3pm.wav"
    Terr.Anotar "002-0183"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0184"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NosREC" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    
    'cerrra grabacion
    Terr.Anotar "002-0185"
    cmdToDo = "Close MP3rec "
    Terr.Anotar "002-0186"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    Terr.Anotar "002-0187"
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        LogErrorMCI dwReturn
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NosRec2" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Sub LogErrorMCI(CodeErrMCI)
    On Error GoTo ERmp3
    
    Dim Buffer As String, Largo As Integer
    Buffer = Space$(512)
    
    Largo = mciGetErrorString(CodeErrMCI, Buffer, Len(Buffer))
    
    Dim ErrTEXT As String
    
    ErrTEXT = Trim(left(Buffer, Len(Buffer)))
    'en este writelog pongo la fecha y hora
    
    Dim txErr As String
    txErr = "****ERRmmPlayer -M-C-I- ****" + vbCrLf + _
                        "N:" + Trim(CStr(CodeErrMCI)) + vbCrLf + _
                        ErrTEXT + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado
    
    Terr.AppendLog txErr
    RaiseEvent mmError(txErr)

    Exit Sub
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Sub

Public Function QuickLargoDeTema(TemaQuick As String) As String
    
    On Local Error GoTo ErrFunc
    Terr.Anotar "002-0192", dwReturn
    
    QuickLargoDeTema = "-:-"
    
'***************************************
'Ya estaba sacado parece que se usaba para karaokes
'***************************************
'    Increible negrada esta y asi lo estan usando algunos
'    FileName(2) = TemaQuick
'    DoOpen 2
'    MiniInicVarMci 2

'***************************************
'sacado el 29/02/2008 no se para que mierda es
'***************************************
'    If dwReturn = 264 Then 'no hay memoria sufuciente!!!
'        LogErrorMCI dwReturn
'        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
'                        "NoMEM" + vbCrLf + _
'                        "HIST:" + vbCrLf + _
'                        Terr.LogAcumulado)
'        'En este caso queda tildado y no puede volver a mostrar hasta que se
'        'cierre el MCI original (el que reproduce)
'        Exit Function
'    End If
'***************************************

'***************************************
'sacado el 29/02/2008 no se para que mierda es
'parece que se usaba en karaokes

'***************************************
'    '------------ver el largo--------------
'    Dim Sec As Long, Mins As Long
'    Sec = mTotalFrames(2) / mFramePerSecond(2)
'    Terr.Anotar "002-0210", Sec
'    If Sec < 60 Then
'        QuickLargoDeTema = "00:" & Format(Sec, "00")
'    Else
'        Mins = Int(Sec / 60)
'        Sec = Sec - (Mins * 60)
'        QuickLargoDeTema = Format(Mins, "00") & ":" & Format(Sec, "00")
'        Terr.Anotar "002-0213", Mins, Sec
'    End If
'
'    MiniCerrar 2
'
'    Exit Function
'
ErrFunc:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
 
End Function

Private Function SoloNumeros(TXT As String) As String
    Dim Largo As Long
    Largo = Len(TXT)
    Dim TmpNumber As String
    TmpNumber = ""
    Dim Letra As String, A As Long
    For A = 1 To Largo
        Letra = Mid(TXT, A, 1)
        If IsNumeric(Letra) Then
            TmpNumber = TmpNumber + Letra
        End If
    Next
    If TmpNumber = "" Then TmpNumber = "0"
    SoloNumeros = TmpNumber
End Function


Public Sub SetDefaultDevice(typeDevice As String, drvDefaultDevice As String)
    'this sub is very important to set the default MCI device
    'maybe xing mpeg installed in your computer and it not support
    'all multimedia files
    'because of this you can rest the default device of MCI to
    'drivers microsft
    'which came with windows or you when install Microsft media player
    'ok any way the default device Following:
    'Device Type        Driver
    'MPEGVideo          mciqtz.drv          this is the most important
    'sequencer          mciseq.drv
    'avivideo           mciavi.drv
    'waveaudio          mciwave.drv
    'videodisc          mcipionr.drv
    'cdaudio            mcicda.drv
    
    'the following for ATI all in Wonder 128 VGA card
    'DvdVideo           MciCinem.drv DVD
    'ATIMPEGVIDEO       mciatim1.drv
    
    'e.g. :
    'SetDefaultDevice "MPEGVideo", "mciqtz.drv" ' this the most
    'improtant device and it will receives calls mci
    'Some programs change this device like xing mpeg
    'and if this occur you can not play all mutimedia files
    'and will occur unexpected errors
    'because of this write this line when your program loaded
    'SetDefaultDevice "MPEGVideo", "mciqtz.drv"
    'to set the strongest default device
    
    'Note: Windows 2000 not use system.ini to set drivers.it use registry.
    
    Dim Res As String
    Dim TMP As String * 255
    Dim Windir As String
    Res = GetWindowsDirectory(TMP, 255)
    Windir = left$(TMP, Res)
    Res = WritePrivateProfileString("MCI", typeDevice, drvDefaultDevice, Windir & "\" & "system.ini")
End Sub

Public Function GetDefaultDevice(typeDevice As String) As String
    'this Function help you if you want to know the default device
    'the parameter must be the device type like:
    'MPEGVideo
    'sequencer
    'avivideo
    'waveaudio
    'videodisc
    'cdaudio
    'and the returned value is a string for the default device
    'Please read the description of sub SetDefaultDevice
    
    Dim TMP As String * 255
    Dim Res As String
    Dim Windir As String
    Res = GetWindowsDirectory(TMP, 255)
    Windir = left$(TMP, Res)
    Res = GetPrivateProfileString("MCI", typeDevice, "None", TMP, 255, Windir & "\" & "system.ini")
    GetDefaultDevice = left$(TMP, Res)
End Function

Private Function MiniCerrar(iAlias As Long) As Long
    
    Terr.Anotar "002-0036", iAlias, Alias(iAlias)
    RET = mciSendString("Close " + Alias(iAlias), 0&, 0&, 0&)
    Terr.Anotar "002-0037"
    RET = CLng(RET)
    If RET <> 0 And RET <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI RET
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NoCierra" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    End If
End Function

Public Sub MiniInicVarMci(iAlias As Long)
    'usar formato de frames que es mmucho mas rapido
    Terr.Anotar "iniMCI001", iAlias
    
    Dim Total As String * 128
    Dim coMh As String
    coMh = "set " + Alias(iAlias) + " time format frames"
    dwReturn = mciSendString(coMh, Total, 128, 0&)
    
'    'agregado en feb 2010, despues de años de funcionar se murio ...
'    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no ha abierto nada
'        LogErrorMCI dwReturn
'        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
'                        "Comando:" + coMh + vbCrLf + _
'                        "HIST:" + vbCrLf + _
'                        Terr.LogAcumulado)
'    End If
    
    'llenar las variables basicas para no usar de nuevo mciSend
    GetTotalframes iAlias 'aqui se carga mTotalFrames
    GetTotalTimeByMS iAlias 'aqui se carga mtotalSec
    GetFramesPerSecond iAlias 'aquis e dividen los dos anteriores y se obtiene mFrames per sec
    
    'siempre asi, puede ser cambiado si quiero que termine antes desde 3PM o mpRock
    mHastaTema(iAlias) = mTotalSec(iAlias)
End Sub

Public Function GetTotalTimeByMS(iAlias As Long) As Long
    
    Dim TotalTime As String * 128

    dwReturn = mciSendString("set " & Alias(iAlias) & " time format ms", TotalTime, 128, 0&)
    dwReturn = mciSendString("status " & Alias(iAlias) & " length", TotalTime, 128, 0&)
    
    mciSendString "set " & Alias(iAlias) & " time format frames", 0&, 0&, 0& ' return focus to frames not to time
    
    If Not dwReturn = 0 Then  'not success
        mTotalSec(iAlias) = -1
        mTotalMiliSec(iAlias) = -1
        GetTotalTimeByMS = -1
        Exit Function
    End If
    
    'Success
    mTotalSec(iAlias) = Val(TotalTime) / 1000
    mTotalMiliSec(iAlias) = Val(TotalTime)
    GetTotalTimeByMS = Val(TotalTime)
End Function

Private Sub Class_Initialize()

    Set Reloj0 = New tbrTimer.clsTimer
    Set Reloj1 = New tbrTimer.clsTimer
    Set Reloj2 = New tbrTimer.clsTimer
    Set Reloj3 = New tbrTimer.clsTimer
    Set RelojKar = New tbrTimer.clsTimer
    
    RelojKar.Interval = 0
    RelojKar.Enabled = False
    mIntervalKar = 40 'para cuando se cargue
    
    Reloj0.Interval = 0
    Reloj1.Interval = 0
    Reloj2.Interval = 0
    Reloj3.Interval = 0
    Reloj0.Enabled = False
    Reloj1.Enabled = False
    Reloj2.Enabled = False
    Reloj3.Enabled = False
    
    Alias(0) = "MP3Play0"
    Alias(1) = "Mp3Play1"
    Alias(2) = "Mp3Play2"
    Alias(3) = "Mp3Play3"
    
End Sub

Public Sub SetDriversVideo()
    'asegurarse que se reproduzca de manera predeterminada con lo que corresponde
    'algunos tienen problemas con los videos
    'imagino que ponen codecs que cambian esto para bien (no lo creo)
    'por las dudas lo hago opcional
    
    'this Function help you if you want to know the default device
    'the parameter must be the device type like:
    'MPEGVideo
    'sequencer
    'avivideo
    'waveaudio
    'videodisc
    If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
        'if Driver "mciqtz.drv" not the default device for type
        '"MpegVideo" then set mciqtz.drv as a default device
        
        SetDefaultDevice "MPEGVideo", "mciqtz.drv"
        'this mciqtz.drv most improtant driver and it will receives calls mci for MPEG types
        'Some programs change this device like xing mpeg
        'and if this occur you can not play all mutimedia files
        'and will occur unexpected errors
        
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + "mciqtz.drv NO ERA PREDETERMINADO")
    End If
    
    If Not GetDefaultDevice("sequencer") = "mciseq.drv" Then
        'if Driver"mciseq.drv" not the default device for type
        '"sequencer" then set mciqtz.drv as a default device
        SetDefaultDevice "sequencer", "mciseq.drv"
    End If
    
    If Not GetDefaultDevice("avivideo") = "mciavi.drv" Then
        'if Driver"mciavi.drv" not the default device for type
        '"avivideo" then set avivideo as a default device
        SetDefaultDevice "avivideo", "mciavi.drv"
        
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + "mciavi.drv NO ERA PREDETERMINADO")
    End If
    
End Sub

Public Sub DefinePathLogs(P1 As String, P2 As String)
    Terr.FileLog = P1
    'en realidad nunca uso el append ya que le paso el detalle al programa usuario de la DLL para que haga lo que quiera!
    Terr.LargoAcumula = 300
    KAR.DefinePathLog P2
    Terr.Anotar "MP3001", P1, P2
End Sub

Public Sub ActivaFulLog(P1 As String)
    Terr.FileLogGrabaTodo = P1
    Terr.ModoGrabaTodo = True
    Terr.StartGrabaTodo
End Sub

Private Function RELOJ(i) As tbrTimer.clsTimer
    If i = 0 Then Set RELOJ = Reloj0
    If i = 1 Then Set RELOJ = Reloj1
    If i = 2 Then Set RELOJ = Reloj2
    If i = 3 Then Set RELOJ = Reloj3
End Function

Public Property Let HastaTema(i As Long, lHasta As Long)
    mHastaTema(i) = lHasta
End Property

Public Property Get HastaTema(i As Long) As Long
    HastaTema = mHastaTema(i)
End Property

Public Property Get EsGratis(i As Long) As Boolean
    EsGratis = mEsGratis(i)
End Property

Public Property Let EsGratis(i As Long, NewEsGratis As Boolean)
    mEsGratis(i) = NewEsGratis
End Property

Public Function DoOpenKar(ArchMN0 As String, ToPicture As Object, ToPelota As Object) As Long
    
    On Error GoTo ERmp3
    DoOpenKar = 0
    If FSO.FileExists(ArchMN0) = False Then
        RaiseEvent mmError("****ERRmmPlayerK****" + vbCrLf + _
                        "NoFile" + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
        DoOpenKar = 1
        Exit Function
    End If
    
    MiniCerrar 2
    ToPicture.AutoRedraw = True
    Terr.Anotar "kar001"
    'limpia el picture y esconde la pelota
    KAR.CleanPic
    Terr.Anotar "Kar002"
    'definir donde se muestra
    KAR.DefinePictureBox ToPicture
    KAR.DefinePelota ToPelota
    Terr.Anotar "Kar003"
    Dim RET As Long
    RET = KAR.DoOpenKar(ArchMN0) 'lo separa y fabrica el mp3 en getPathMP3
    If RET <> 0 Then
        RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        "NO PUDO ABRIR EL KARAOKE" + vbCrLf + _
                        ArchMN0 + vbCrLf + _
                        "ERR=" + CStr(RET) + vbCrLf + _
                        "LOGK:" + KAR.GetLOG + vbCrLf + _
                        "HISTMM:" + vbCrLf + _
                        Terr.LogAcumulado)
        
        DoOpenKar = 2
        Exit Function
    End If
    Terr.Anotar "Kar004"
    m_FileName(2) = KAR.GetPathMP3
    DoOpen 2
    
    Exit Function
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    Resume Next
End Function

Public Function DoPlayKar()
    'arrancar la musica
    DoPlay 2
    
    'descifrar en que milisegundo empezo la reproduccion!
    '              = ahora - tiempo que lleva reproducido
    
    ResincroKar
    
    RelojKar.Enabled = True
    RelojKar.Interval = mIntervalKar
    
End Function

Public Sub SetIntervalEventosKAR(NewInterval As Long)
    mIntervalKar = NewInterval
    If RelojKar.Enabled Then RelojKar.Interval = mIntervalKar
End Sub

Public Sub DoStopKar()
    RelojKar.Interval = 0: RelojKar.Enabled = False
    KAR.CleanPic
    DoStop 2
    KAR.BorraTemp
End Sub

Private Sub RelojKar_Timer()
    MilliSecKar = (CDbl(Timer) * 1000) - MilliSecKarIni
    RaiseEvent FaltaNextEvKAR(KAR.EjecutarEvento(MilliSecKar))
End Sub

Public Sub ResincroKar()
    Dim EmpezoHace As Double
    EmpezoHace = GetCurrentMultimediaPos(2) / mFramePerSecond(2)
    MilliSecKarIni = (CDbl(Timer) - EmpezoHace) * 1000
End Sub

Private Sub Reloj_Timer(Index As Integer)
    On Error GoTo ERmp3
    Terr.Anotar "002-0001", Index, mFramePerSecond(CLng(Index))
    'primero ver si termina el tema
    
    Dim CurrPos As Single, lFPS As Long
    CurrPos = GetCurrentMultimediaPos(CLng(Index))
    lFPS = mFramePerSecond(CLng(Index))
    
    'no dividir por cero!
    If lFPS = 0 Then lFPS = 1
    mSecActual(CLng(Index)) = CurrPos / lFPS
    
    If mFramePerMiliSecond(Index) = 0 Then mFramePerMiliSecond(Index) = 1
    mMiliSecActual(CLng(Index)) = CurrPos / mFramePerMiliSecond(Index)
    
    Terr.Anotar "002-0001b", CurrPos, mSecActual(Index)
    
    Dim Termino As Boolean
    Termino = False
    
    'se me terminan los karaokes cuando estan empezando
    If CurrPos = -1 Then
        Termino = True
        Terr.Anotar "cur1", CurrPos, mTotalFrames(CLng(Index))
        GoTo EndSong
    End If
    
    If mTotalFrames(CLng(Index)) = -1 Then
        Termino = True
        Terr.Anotar "cur2", CurrPos, mTotalFrames(CLng(Index))
        GoTo EndSong
    End If
    
    If CurrPos >= (mTotalFrames(CLng(Index)) - 1) Then
        Terr.Anotar "cur3", CurrPos, mTotalFrames(CLng(Index))
        Termino = True
        GoTo EndSong
    End If
    
    'If IsPlaying(CLng(Index)) = False Then Termino = True: GoTo EndSong
    
    If mSecActual(CLng(Index)) >= mHastaTema(Index) Then
        Termino = True
        Terr.Anotar "cur4", CurrPos, mTotalFrames(CLng(Index))
        GoTo EndSong
    End If
        
    
    'Terr.AppendSinHist "I:" + CStr(Index) + _
        " CurrPos:" + CStr(CurrPos) + _
        " Sec:" + CStr(mSecActual(CLng(Index))) + _
        " TotalF:" + CStr(mTotalFrames(CLng(Index))) + _
        " Termino=" + CStr(Termino)
    
EndSong:
    If Termino Then
        
        RELOJ(Index).Interval = 0
        'si termino un karaoke!!!
        If Index = 2 Then
            'Terr.AppendLog "FinKar:" + CStr(CurrPos), CStr(mTotalFrames(CLng(Index)))
            'XXXX esto borraba la imagen de fndo al comenzar un karaoke que seguia despues de otro
            RelojKar.Interval = 0: RelojKar.Enabled = False
            KAR.CleanPic
        End If
        DoStop CLng(Index)
        'RaiseEvent EndPlay(CLng(Index))
        Exit Sub
    Else
        'saber que segundo es ...
        Dim sActual As Long
        RaiseEvent Played(CLng(mSecActual(CLng(Index))), CLng(Index), CLng(mMiliSecActual(CLng(Index))))
    End If
    
    Exit Sub
    
'    If IsPlaying(CLng(Index)) = False Then
'        Reloj(Index).Interval = 0
'        RaiseEvent EndPlay(CLng(Index))
'        Exit Sub 'ESTO NO ESTABA!!!!!!!, seguia mandando el evento!!!!!!!!!!
'    End If
'    'y SOLO si no termino largar el evento. Antes estaba alreves!!!!!!!!!
'    tERR.Anotar "002-0005"
'    RaiseEvent Played(PositionInSec(CLng(Index)), CLng(Index))
'    Exit Sub
ERmp3:
    RaiseEvent mmError("****ERRmmPlayer****" + vbCrLf + _
                        Terr.ErrToTXT(Err) + vbCrLf + _
                        "HIST:" + vbCrLf + _
                        Terr.LogAcumulado)
    
    Resume Next
End Sub
