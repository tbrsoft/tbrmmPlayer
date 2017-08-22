VERSION 5.00
Begin VB.UserControl MP3Play 
   BackColor       =   &H00FF00FF&
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   PropertyPages   =   "MP3Play.ctx":0000
   ScaleHeight     =   1620
   ScaleWidth      =   1500
   ToolboxBitmap   =   "MP3Play.ctx":0011
End
Attribute VB_Name = "MP3Play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
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

Event Played(SecondsPlayed As Long, iAlias As Long)
Event BeginPlay(iAlias As Long)
Event EndPlay(iAlias As Long)

Private TMPs As String
Private mTotalSec(3) As Single
Private mTotalFrames(3) As Single
Private mFramePerSecond(3) As Single
Private mSecActual(3) As Single

Private Sub Reloj_Timer(Index As Integer)
    On Error GoTo ERmp3
    tERR.Anotar "002-0001", Index, mFramePerSecond(CLng(Index))
    'primero ver si termina el tema
    
    Dim CurrPos As Single
    CurrPos = GetCurrentMultimediaPos(CLng(Index))
    mSecActual(CLng(Index)) = CurrPos / mFramePerSecond(CLng(Index))
    
    tERR.Anotar "002-0001b", CurrPos, mSecActual(CLng(Index)), Index
    'tERR.AppendSinHist "I:" + CStr(Index) + _
        " CurrPos:" + CStr(CurrPos) + _
        " Sec:" + CStr(mSecActual(CLng(Index))) + _
        " TotalF:" + CStr(mTotalFrames(CLng(Index)))
    Dim Termino As Boolean
    Termino = False
    
    If CurrPos = -1 Or mTotalFrames(CLng(Index)) = -1 Then Termino = True: GoTo EndSong
    If CurrPos >= (mTotalFrames(CLng(Index)) - 1) Then Termino = True: GoTo EndSong
    'If IsPlaying(CLng(Index)) = False Then Termino = True: GoTo EndSong
    If mSecActual(CLng(Index)) >= TotalTema(CLng(Index)) Then Termino = True: GoTo EndSong
    
EndSong:
    If Termino Then
        RELOJ(Index).Interval = 0
        DoStop CLng(Index)
        'RaiseEvent EndPlay(CLng(Index))
        Exit Sub
    Else
        'saber que segundo es ...
        Dim sActual As Long
        RaiseEvent Played(CLng(mSecActual(CLng(Index))), CLng(Index))
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
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayctl" + ".acpw"
    Resume Next
End Sub

Public Function GetCurrentMultimediaPos(iAlias As Long) As Long
    'se refiere a frame actual
    
    Dim dwReturn As Long
    Dim pos As String * 128
    tERR.Anotar "002-0001t", iAlias, Alias(iAlias)
    dwReturn = mciSendString("status " & Alias(iAlias) & " position", pos, 128, 0&)
    tERR.Anotar "002-0001s", dwReturn
    If Not dwReturn = 0 Then  'not success
        GetCurrentMultimediaPos = -1
        Exit Function
    End If
    
    'Success
    GetCurrentMultimediaPos = CLng(pos)
End Function

Public Function GetTotalframes(iAlias As Long) As Long
    
    Dim dwReturn As Long
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

Private Sub UserControl_Initialize()
    tERR.Anotar "MP3001"
    Alias(0) = "MP3Play0"
    Alias(1) = "Mp3Play1"
    Alias(2) = "Mp3Play2"
    Alias(3) = "Mp3Play3"
    'asegurarse que se reproduzca de manera predeterminada con lo que corresponde
    
    'this Function help you if you want to know the default device
    'the parameter must be the device type like:
    'MPEGVideo
    'sequencer
    'avivideo
    'waveaudio
    'videodisc
    If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
        'if Driver"mciqtz.drv" not the default device for type
        '"MpegVideo" then set mciqtz.drv as a default device
        
        SetDefaultDevice "MPEGVideo", "mciqtz.drv"
        'this mciqtz.drv most improtant driver and it will receives calls mci for MPEG types
        'Some programs change this device like xing mpeg
        'and if this occur you can not play all mutimedia files
        'and will occur unexpected errors
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
    End If
    
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 1620
    UserControl.Width = 1500
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property

Public Property Get Volumen(iAlias As Long) As Long
    On Error GoTo ERmp3
    tERR.Anotar "002-0011", iAlias, m_Volumen(iAlias)
    Volumen = m_Volumen(iAlias) / 10
    Exit Property
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acpv"
End Property

Public Property Let Volumen(iAlias As Long, ByVal New_Volumen As Long)
    On Error GoTo ERmp3
    'en mi máquina anda del 0 al 1000 (en todas)
    tERR.Anotar "002-0013"
    m_Volumen(iAlias) = New_Volumen * 10
    TMPs = "SetAudio " + Alias(iAlias) + " Volume To " + CStr(m_Volumen(iAlias))
    tERR.Anotar "002-0014", iAlias, TMPs, IAA, IAANext
    Ret = mciSendString(TMPs, 0&, 0&, 0&)
    tERR.Anotar "002-0015", Ret
    If Ret <> 0 Then
        LogErrorMCI Ret
        'no se pudo modificar el volumen
        tERR.AppendLog "NoVolumenEn:" + CStr(Ret), "MpPalyCtl" + ".acpw"
    End If
    tERR.Anotar "002-0018"
    Exit Property
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPalyCtl" + ".acpx"
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

Public Function IsPlaying(iAlias As Long) As Boolean
    Dim s As String * 128
    On Error GoTo ERmp3
    tERR.Anotar "002-0027"
    If m_FileName(iAlias) = "" Then
        tERR.Anotar "002-0028", iAlias
        IsPlaying = False
    Else
        tERR.Anotar "002-0030", HabilitarVUMetro, NoVumVID
        Ret = mciSendString("status " + Alias(iAlias) + " mode", s, 128, 0&)
        tERR.Anotar "002-0031", Ret, iAlias
        If Ret = 263 Then '263 es cuando no ha abierto nada
            IsPlaying = False
            Exit Function
        End If
        If Ret <> 0 Then
            LogErrorMCI Ret
            'no se pudo modificar el volumen
            tERR.AppendLog "ERR IsPlaying=Status:" + CStr(Ret)
            'WriteLog "No se pudo definir el estado de ejecucion." + ". Tema: " + m_FileName + " Function IsPlaying", False
        End If
        'EN ESTE CASO ES NULO O ALGO ASI
        'YA QUE MCI NO TIENE LA CAPACIDAD DE STATUS!!!
        If Ret = 274 Then
            tERR.Anotar "002-0034b", iAlias
            IsPlaying = True
        Else
            tERR.Anotar "002-0034", s
            IsPlaying = (Mid(s, 1, 7) = "playing")
        End If
        
        IsPlaying = (Mid(s, 1, 7) = "playing")
    End If
    
    Exit Function
ERmp3:
    tERR.Anotar "002-0034b"
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acpx"
    Resume Next
End Function

Public Function DoOpen(iAlias As Long)
    
    On Error GoTo ERmp3
    MiniCerrar iAlias
    'ver si esta el archivo
    tERR.Anotar "002-0041"
    If FSO.FileExists(m_FileName(iAlias)) = False Then
        tERR.AppendLog "MpPlayCtl.DoOpen.NoExistFile.acqb", m_FileName(iAlias)
        Exit Function
    End If
    
    tERR.Anotar "002-0040"
    Dim lenShort As Long, TMP As String * 255
    lenShort = GetShortPathName(m_FileName(iAlias), TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes el reproductor no anda. JOYA JOYA JOYA
    tERR.Anotar "002-0045", m_FileName(iAlias)
    
    Dim FileNameSHORT As String
    FileNameSHORT = left$(TMP, lenShort)
    
    tERR.Anotar "002-0046", FileNameSHORT
    
    Dim cmdToDo As String * 255
    
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias " + _
        Alias(iAlias) + " style " & WS_CHILD
    
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0049", dwReturn
    
    If dwReturn <> 0 Then
        If dwReturn = 263 Then
            'si da el error 263 es probable que la máquina no tenga MCI, lo que le paso a Mauro con W98 PE y a efren con ME
            tERR.AppendLog "WINDOWS NO REPRODUCE MP3!!!. INSTALE EL REPRODUCTOR CORESPONDIENTE " + _
                "A SU VERSION DE WINDOWS"
            MsgBox "No se ha podido abrir el fichero debido a un problema existente en Windows. " + vbCrLf + _
                "Revise que el reproductor multimedia de Windows este instalado y funcione correctamente." + _
                "Notifique a tbrSoft de esto para más detalles"
        End If
        
        LogErrorMCI dwReturn
        'no se puedo abrir!!!
        tERR.AppendLog "DoOpen.NoAbre." + CStr(dwReturn), "acqe"
    End If
    
    tERR.Anotar "002-0054"
    
    MiniInicVarMci iAlias
    
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acdy"
    Resume Next
End Function

Public Function DoOpenVideo(Style As String, HWind As Long, _
    X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, iAlias As Long)
    
    On Error GoTo ERmp3
    
    tERR.Anotar "002-0058"
    
    MiniCerrar iAlias
    
    Dim cmdToDo As String * 255
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    If Dir(m_FileName(iAlias)) = "" Then
        tERR.Anotar "002-0065b"
        tERR.AppendLog "NoEx.acqh", m_FileName(iAlias)
        Exit Function
    End If
    tERR.Anotar "002-0068"
    lenShort = GetShortPathName(m_FileName(iAlias), TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    tERR.Anotar "002-0069", m_FileName(iAlias)
    'volu = mciGetDeviceID(lenShort)
    FileNameSHORT = left$(TMP, lenShort)
    tERR.Anotar "002-0070", FileNameSHORT
    
    tERR.Anotar "002-0071", HabilitarVUMetro, NoVumVID, HWind
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias " + _
        Alias(iAlias) + " parent " + CStr(HWind) + " style " & WS_CHILD 'Style
        
    tERR.Anotar "002-0072", cmdToDo '                         xxxxxx
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0073", dwReturn, Salida2, Style, HWind
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlay.DoOpenVid.acqh." + CStr(dwReturn), m_FileName(iAlias)
    End If
    tERR.Anotar "002-0076", X1, X2, Y1, Y2
    
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
    tERR.Anotar "002-0077", Salida2, Style, HWind, dwReturn
    
    If dwReturn <> 0 Then
        '********************
        'si es 346 es que no tiene ventana de presentacion (base+90=MCIERR_NO_WINDOW)
        '¿¿¿¿¿????????
        'probe con style popup y overlapped y no sirven ni solucionan
        If dwReturn = 346 Then
            tERR.AppendLog "MCIERR_NO_WINDOW=346. No hay ventana de presentacion!!!" + CStr(dwReturn), m_FileName(iAlias)
            'pasa con videos con codecs nuevos que al cargarse si funciona en
            'WMP pero no en 3PM!!!!!!!!!!!!!!!!!!!
        '********************
        Else
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            tERR.AppendLog "MpPlay.DoOpenVid.WindowAt.acqi." + CStr(dwReturn), m_FileName(iAlias)
        End If
    End If
    
    MiniInicVarMci iAlias
    
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqj"
    Resume Next
End Function

Private Function GetFramesPerSecond(iAlias As Long) As Long
    mFramePerSecond(iAlias) = mTotalFrames(iAlias) / (mTotalSec(iAlias))
    GetFramesPerSecond = mFramePerSecond(iAlias)
End Function

Public Function DoPlay(iAlias As Long, Optional FullScreen As Boolean = False)
    On Error GoTo ERmp3
    tERR.Anotar "002-0082", CStr(FullScreen), iAlias
    
    tERR.Anotar "002-0082b", dwReturn
    If FullScreen Then
        'dwReturn = mciSendString("play " + Alias(iAlias) + " fullscreen from 0 to " + CStr(mTotalFrames(iAlias)), 0&, 0&, 0&)
        dwReturn = mciSendString("play " + Alias(iAlias) + " fullscreen from 0", 0&, 0&, 0&)
    Else
        'dwReturn = mciSendString("play " + Alias(iAlias) + " from 0 to " + CStr(mTotalFrames(iAlias)), 0&, 0&, 0&)
        dwReturn = mciSendString("play " + Alias(iAlias) + " from 0", 0&, 0&, 0&)
    End If
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.DoPlay.Play." + m_FileName(iAlias), ".acqk"
    End If
    tERR.Anotar "002-0086", dwReturn
    RELOJ(iAlias).Interval = 1000
    tERR.Anotar "002-0087"
    RaiseEvent BeginPlay(iAlias)
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acql"
    Resume Next
End Function

Public Function DoPause(iAlias As Long)
    On Error GoTo ERmp3
    tERR.Anotar "002-0088"
    dwReturn = mciSendString("pause " + Alias(iAlias), 0&, 0&, 0&)
    tERR.Anotar "002-0089"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.DoPause." + m_FileName(iAlias), ".acqm"
    End If
    tERR.Anotar "002-0092"
    RELOJ(iAlias).Interval = 0
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqn"
    Resume Next
End Function

Public Function DoStop(iAlias As Long) As String
    On Error GoTo ERmp3
    tERR.Anotar "002-0093", iAlias
    dwReturn = mciSendString("stop " + Alias(iAlias), 0&, 0&, 0&)
    If dwReturn <> 0 And dwReturn <> 263 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.DoStop." + m_FileName(iAlias), ".acqo"
    End If
    tERR.Anotar "002-0097"
    RELOJ(iAlias).Interval = 0
    tERR.Anotar "002-0098"
    RaiseEvent EndPlay(iAlias)
    
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqp"
    Resume Next
End Function

Public Function DoClose(iAlias As Long) As String
    'iAlias es el que hay que cerrar o:
        '99 cierra todos
        
    On Error GoTo ERmp3
    tERR.Anotar "002-0099", iAlias
    If iAlias = 99 Then 'cierra todos
        Dim F11 As Long
        For F11 = 0 To 3
            dwReturn = mciSendString("stop " + Alias(F11), 0&, 0&, 0&)
            RELOJ(F11).Interval = 0
        Next F11
        
        dwReturn = mciSendString("close all", 0&, 0&, 0&)
        
        If dwReturn <> 0 And dwReturn <> 263 Then '263 ES CUANDO NO HAY NADA ABIERTO
            LogErrorMCI dwReturn
            tERR.AppendLog "MpPlayCtl.DoClose." + m_FileName(F11), ".acqr"
        End If
        
    Else 'o solo el elegido
        dwReturn = mciSendString("close " + Alias(iAlias), 0&, 0&, 0&)
        If dwReturn <> 0 And dwReturn <> 263 Then '263 ES CUANDO NO HAY NADA ABIERTO
            LogErrorMCI dwReturn
            tERR.AppendLog "MpPlayCtl.DoClose." + m_FileName(iAlias), ".acqr"
        End If
        tERR.Anotar "002-0103"
        RELOJ(iAlias).Interval = 0
    End If
    'SI SIGUE EL RELOJ SE MARCAN 1000 errores!!!!!!!!!!
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlatCtl" + ".acqq"
    Resume Next
End Function

Public Function PercentPlay(iAlias As Long)
    On Error GoTo ERmp3
    tERR.Anotar "002-0104"
    PercentPlay = PositionInSec(iAlias) / mTotalSec(iAlias) * 100
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqs"
    Resume Next
End Function

Public Function PositionInSec(iAlias As Long) As Long
    PositionInSec = mSecActual(iAlias)
End Function

Public Function FaltaInSec(iAlias As Long)
    On Error GoTo ERmp3
    tERR.Anotar "002-0130"
    FaltaInSec = mTotalSec(iAlias) - mSecActual(iAlias)
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqu"
    Resume Next
End Function

Public Function Falta(iAlias As Long) As String
    On Error GoTo ERmp3
    Dim MINS As Long, SEC As Long
    tERR.Anotar "002-0131"
    SEC = FaltaInSec(iAlias)
    tERR.Anotar "002-0132", SEC, iAlias
    If SEC < 60 Then Falta = "0:" & Format(SEC, "00")
    If SEC > 59 Then
        tERR.Anotar "002-0134"
        MINS = Int(SEC / 60)
        tERR.Anotar "002-0135", MINS
        SEC = SEC - (MINS * 60)
        tERR.Anotar "002-0136", SEC
        Falta = Format(MINS, "00") & ":" & Format(SEC, "00")
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqv"
    Resume Next
End Function

Public Function LengthInSec(iAlias As Long) As Long
    tERR.Anotar "002-0137"
    LengthInSec = CLng(mTotalSec(iAlias))
End Function

Public Function Length(iAlias As Long) As String
    On Error GoTo ERmp3
    tERR.Anotar "002-0148"
    SEC = mTotalSec(iAlias)
    tERR.Anotar "002-0149"
    If SEC < 60 Then Length = "0:" & Format(SEC, "00")
    tERR.Anotar "002-0150"
    If SEC > 59 Then
        tERR.Anotar "002-0151"
        MINS = Int(SEC / 60)
        tERR.Anotar "002-0152"
        SEC = SEC - (MINS * 60)
        tERR.Anotar "002-0153"
        Length = Format(MINS, "00") & ":" & Format(SEC, "00")
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqx"
    Resume Next
End Function

Public Function SeekTo(Second, iAlias As Long)
    On Error GoTo ERmp3
    tERR.Anotar "002-0154", iAlias
    
    'transformar segundos a frames
    Dim STF As Long
    STF = (Second / 1000) * mFramePerSecond(iAlias)
    If IsPlaying(iAlias) Then
        tERR.Anotar "002-0155"
        dwReturn = mciSendString("play " + Alias(iAlias) + " from " & STF, 0&, 0&, 0&)
        tERR.Anotar "002-0156"
        If dwReturn <> 0 And dwReturn <> 282 Then '282 es que pide un lugar de tiempo que no existe!
            tERR.Anotar "002-0157"
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            tERR.Anotar "002-0158"
            tERR.AppendLog "MpPlayCtl.SeekTo.Open." + m_FileName(iAlias) + ".acqy"
        End If
    Else
        tERR.Anotar "002-0159"
        dwReturn = mciSendString("seek " + Alias(iAlias) + " to " & STF, 0, 0, 0)
        If dwReturn <> 0 And dwReturn <> 263 And dwReturn <> 282 Then
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            tERR.AppendLog "MpPlayCtl.SeekTo.Close." + m_FileName(iAlias) + ".acqz"
        End If
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acra"
    Resume Next
End Function

Function Record() 'no se como hara con los alias, estimo que graba todo ¿¿??
    On Error GoTo ERmp3
    tERR.Anotar "002-0162"
    dwReturn = mciSendString("Close MP3rec", 0&, 0&, 0&)
    tERR.Anotar "002-0163"
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        tERR.Anotar "002-0164"
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.Anotar "002-0165"
        tERR.AppendLog "MpPlayCtl.acrb"
    End If
    tERR.Anotar "002-0166"
    Dim cmdToDo As String * 255
    tERR.Anotar "002-0167"
    'abrir nuevo
    tERR.Anotar "002-0168"
    cmdToDo = "open new type WaveAudio Alias MP3rec"
    tERR.Anotar "002-0169"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.acrc"
        Exit Function
    End If
    'iniciar grabacion
    tERR.Anotar "002-0174"
    cmdToDo = "record MP3rec"
    tERR.Anotar "002-0175"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.acrd"
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl.acre"
    Resume Next
End Function

Function StopRecord()
    On Error GoTo ERmp3
    Dim cmdToDo As String * 255
    tERR.Anotar "002-0178"
    'parar nuevo
    cmdToDo = "stop MP3rec"
    tERR.Anotar "002-0179"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0180"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.acrf"
    End If
    'grabar grabacion
    tERR.Anotar "002-0182"
    cmdToDo = "save MP3rec c:\3pm.wav"
    tERR.Anotar "002-0183"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0184"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.acrg"
    End If
    
    'cerrra grabacion
    tERR.Anotar "002-0185"
    cmdToDo = "Close MP3rec "
    tERR.Anotar "002-0186"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0187"
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.acrh"
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acrh"
    Resume Next
End Function

Public Sub LogErrorMCI(CodeErrMCI)
    On Error GoTo ERmp3
    
    Dim Buffer As String, Largo As Integer
    Buffer = Space$(512)
    
    Largo = mciGetErrorString(CodeErrMCI, Buffer, Len(Buffer))
    
    Dim ErrTEXT As String
    
    ErrTEXT = left(Buffer, Len(Buffer))
    'en este writelog pongo la fecha y hora
    tERR.AppendLog "MciErr:" + Trim(CStr(CodeErrMCI)), ErrTEXT
    Exit Sub
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlay.acri"
    Resume Next
End Sub

Public Function QuickLargoDeTema(TemaQuick As String) As String
    On Local Error GoTo ErrFunc
    tERR.Anotar "002-0192"
    
    QuickLargoDeTema = "N/S"
    
    FileName(2) = TemaQuick
    DoOpen 2
    MiniInicVarMci 2
    
    If dwReturn = 264 Then 'no hay memoria sufuciente!!!
        LogErrorMCI dwReturn
        tERR.AppendLog "acrk." + CStr(dwReturn)
        'En este caso queda tildado y no puede volver a mostrar hasta que se
        'cierre el MCI original (el que reproduce)
        Exit Function
    End If
    
    '------------ver el largo--------------
    SEC = mTotalFrames(2) / mFramePerSecond(2)
    tERR.Anotar "002-0210", SEC
    If SEC < 60 Then
        QuickLargoDeTema = "00:" & Format(SEC, "00")
    Else
        tERR.Anotar "002-0212"
        MINS = Int(SEC / 60)
        tERR.Anotar "002-0213"
        SEC = SEC - (MINS * 60)
        tERR.Anotar "002-0214"
        QuickLargoDeTema = Format(MINS, "00") & ":" & Format(SEC, "00")
        tERR.Anotar "002-0215"
    End If
    
    MiniCerrar 2
    
    Exit Function
    
ErrFunc:
    tERR.AppendLog "acrn." + TemaQuick
 
End Function

Private Function SoloNumeros(TXT As String) As String
    Dim Largo As Long
    Largo = Len(TXT)
    Dim TmpNumber As String
    TmpNumber = ""
    Dim Letra As String
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

Private Sub MiniCerrar(iAlias As Long)
    tERR.Anotar "002-0035"
    Dim Ret As String * 128
    tERR.Anotar "002-0036", iAlias, Alias(iAlias)
    Ret = mciSendString("Close " + Alias(iAlias), 0&, 0&, 0&)
    tERR.Anotar "002-0037"
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        tERR.AppendLog "NoCierraMCI.RET:" + CStr(Ret), m_FileName(iAlias)
    End If
End Sub

Public Sub MiniInicVarMci(iAlias As Long)
    'usar formato de frames que es mmucho mas rapido
    tERR.Anotar "iniMCI001"
    
    Dim Total As String * 128
    dwReturn = mciSendString("set " & Alias(iAlias) & " time format frames", Total, 128, 0&)
    
    'llenar las variables basicas para no usar de nuevo mciSend
    GetTotalframes iAlias 'aqui se carga mTotalFrames
    GetTotalTimeByMS iAlias 'aqui se carga mtotalSec
    GetFramesPerSecond iAlias 'aquis e dividen los dos anteriores y se obtiene mFrames per sec
End Sub

Public Function GetTotalTimeByMS(iAlias As Long) As Long
    
    Dim dwReturn As Long
    Dim TotalTime As String * 128

    dwReturn = mciSendString("set " & Alias(iAlias) & " time format ms", TotalTime, 128, 0&)
    dwReturn = mciSendString("status " & Alias(iAlias) & " length", TotalTime, 128, 0&)
    
    mciSendString "set " & Alias(iAlias) & " time format frames", 0&, 0&, 0& ' return focus to frames not to time
    
    If Not dwReturn = 0 Then  'not success
        mTotalSec(iAlias) = -1
        GetTotalTimeByMS = -1
        Exit Function
    End If
    
    'Success
    mTotalSec(iAlias) = Val(TotalTime) / 1000
    GetTotalTimeByMS = Val(TotalTime)
End Function
