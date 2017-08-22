VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrPlayMN0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim FS As New Scripting.FileSystemObject

Private DireccionPelota As Long 'si la pelota de los karaokes sube o baja
Private TopPelota As Long 'altura de la pelota
Private LetrasAcumFrase As Long 'cantidad de letras que pasaron sin empezar una _

Private FRASE As String 'frase actual del NK0
Private SigFrase As String ' frase siguiente del NK0
Private LE As Long 'numero de evento que se ejecuta
Private LEF As Long 'numero de frases del NK0

Private ColorNormal As OLE_COLOR
Private ColorSeleccionado As OLE_COLOR

Dim cTMP As String 'carpeta de temporales`para colocar los archivos y borrarlos

Dim nNK0 As String 'archivo NK0 dentro del MN0
Dim nMP3 As String 'mp3 adentro
Dim nTMP As String 'detalle de fuentes y colores
Dim nIMG As String 'imagen dentro del MN0

Private P As Object 'algun objeto frm o pBox para dibujar todo
Private Pelota As Object 'REEMPLAZAR POR CIRCLE

Private T() As String 'lista de renglonres por cantarse
Private ListaEventos() As String
Private ListaEvFraces() As String

Private mFile As String 'archivo NK0 abierto

Private mSegundosParaAvisar As Long 'tramos sin eventos que vale la pena avisar al
'cantante que hay alguna parte instrumental
Private mEsteEventoEsLargo As Boolean

Private TFU As Long 'tamaño de fuente usado
Private Sombra As Long 'distancia del segundo texto

Private Terr As New tbrErrores.clsTbrERR

Public Function GetPathMP3() As String
    GetPathMP3 = nMP3
End Function

Public Function DoOpenKar(Archivo As String) As Long 'separar el mn0 y deja todo los eventos listos!
    
    Terr.Anotar "1000", Archivo
    'el mmPlayer lo llama cuando tiene el mp3 abierto
    Dim RET As Long
    RET = SepararMN0(Archivo) 'graba los archivos dentro del mn0 en la carpeta temp
    Terr.Anotar "1001", RET
    If RET <> 0 Then
        'si fue negativo es un error del sistema y el numero es abs(ret)
        'si es de 1 a 4 es que falto algun archivo o se descomprimio mal
        DoOpenKar = RET
        Exit Function
    End If
    
    Terr.Anotar "1002"
    SetFormato 'define la fuente y los tamaños del mn0
    Terr.Anotar "1003"
    DireccionPelota = 0
    
    LeeKar nNK0 'hacer una lista de eventos
    Terr.Anotar "1004", nIMG
    'poner la imagen que el usuario eligio
    P.PaintPicture LoadPicture(nIMG), 0, 0, P.Width, P.Height
    Dim BT As String
    BT = App.Path + "\tmp.bmp"
    If FS.FileExists(BT) Then FS.DeleteFile BT, True
    'If Dir(BT) <> "" Then Kill BT
    '¿ por que lo graba a BMP ? por el tamaño!
    Terr.Anotar "1005", BT
    SavePicture P.Image, BT
    'P.Picture = LoadPicture(ArchIMG)
    P.Picture = LoadPicture(BT) '<<<<<
    
    LE = 0: LEF = 0 'numero de evento y de frase en cero
    
    Pelota.Visible = True
    Terr.Anotar "1006"
    GetTFU 'definir el tamaño de fuente que se va a usar!!!
    
End Function

Private Function SepararMN0(Archivo As String) As Long
    
    On Local Error GoTo ErrSEPMN
    Terr.Anotar "1007"
    'graba los archivos en APP + TEMP
    Dim Archivos() As String
    Dim qDatos() As String
    
    Dim Mapo As String
    Dim tDato As String
    tDato = Space$(FileLen(Archivo))
    
    Open Archivo For Binary As #1
        Get #1, 1, tDato
        Mapo = Mid(tDato, 1, InStr(tDato, "+++") - 1)
        qDatos = Split(Mapo, "*")
    
        Dim Datoo As String
        Dim Dqs() As String
        Dim Pdv As Long
        Pdv = Len(Mapo) + 5
        Dim R As Long
        For R = 0 To UBound(qDatos) - 1
            Dqs = Split(qDatos(R), "?")
            Datoo = Space$(Dqs(1))
            Terr.Anotar "1008", R
            Dim ArchivoDestino As String
            Select Case R
                Case 0
                    nNK0 = cTMP + Dqs(0)
                    ArchivoDestino = nNK0
                Case 1
                    nMP3 = cTMP + Dqs(0)
                    ArchivoDestino = nMP3
                Case 2
                    nIMG = cTMP + Dqs(0)
                    ArchivoDestino = nIMG
                Case 3
                    nTMP = cTMP + Dqs(0)
                    ArchivoDestino = nTMP
            End Select
            Terr.Anotar "1009", ArchivoDestino
            'si existe algo lo borro
            If FS.FileExists(ArchivoDestino) Then FS.DeleteFile ArchivoDestino, True
            'BOSTAS DEL MANU
            'If Not Dir(ArchivoDestino) = "" Then Kill ArchivoDestino
    
            Open ArchivoDestino For Binary As #2
                Get #1, Pdv, Datoo
                Put #2, 1, Datoo
            Close #2
            Pdv = Pdv + Dqs(1) + 1
            'DoEvents sacado 3 mar 2010 ya que jode al USB ILUS ... que cagadaaaaaaaaaaaaa
        Next R
    Close 'cierra todos, no solo el 1
    Terr.Anotar "1010"
    'VER QUE ESTEN TODOS LOS ARCHIVOS
    If FS.FileExists(nNK0) = False Then
        SepararMN0 = 1
        Exit Function
    End If
    
    If FS.FileExists(nMP3) = False Then
        SepararMN0 = 2
        Exit Function
    End If
    
    If FS.FileExists(nIMG) = False Then
        SepararMN0 = 3
        Exit Function
    End If
    
    If FS.FileExists(nTMP) = False Then
        SepararMN0 = 4
        Exit Function
    End If
    'todo ok
    SepararMN0 = 0
    Exit Function
ErrSEPMN:
    SepararMN0 = -(Err.Number)
End Function

Public Function EjecutarEvento(Miliseg As Double) As Long
    'devuelve el proximo evento que hara
    'para saber si tiene que esperar!
    Dim TmpEE As Long 'temporal del valor de retorno

    'PRIMERO LA PELOTA!
    If DireccionPelota = 0 Then
        TopPelota = TopPelota + 200
        If TopPelota > P.Height / 4 - Pelota.Height - 60 Then
            DireccionPelota = 1
        End If
    Else
        TopPelota = TopPelota - 200
        If TopPelota < Pelota.Height Then
            DireccionPelota = 0
        End If
    End If
    Pelota.top = TopPelota
    '-------------------------------------------------
    '-------------------------------------------------
    
    'AHORA LA LETRA
    'si no llego el tiempo salir antes
    Dim HayQuePintar As Boolean 'si salteo la letra por que no la necesito
    'y voy a ver si no necesito la frase ya se que no hay que pintar si la frase tampoco necesita nada
 
    Dim TextoActual As String 'texto que se esta marcando
    
    'contar el tiempo que falta para el proximo evento
    If (LE + 1) > MaxEventos Then
        TmpEE = -1 'ya termina la cancion .... ruido final
    Else
        TmpEE = (GetTimeShow(LE + 1) - Miliseg) / 1000
    End If
    
    If Miliseg < GetTimeShow(LE) Then
        HayQuePintar = False
        GoTo VerFrase
    Else
        HayQuePintar = True
        'ver que no se pase del total!!!
        If LE > MaxEventos Then
            Exit Function
        Else
            'si es una letra valida es la que se va a mostrar
            'por lo tanto le sumo las letras que tiene
            TextoActual = GetLetra(LE)
            LetrasAcumFrase = LetrasAcumFrase + Len(TextoActual)
        End If
        
        'RECIEN MUESTRO ALGO, ACA ME FIJO SI ES LARGO!
        If TmpEE < 5 Then 'no me sirve este evento
            mEsteEventoEsLargo = False
        Else 'tomo 5 segundos como referencia
            mEsteEventoEsLargo = True
        End If
        
        LE = LE + 1
        'si no se termino dejo que se muestre
        'si hay algun atraso esto se ejecutara cad 20 miliseg y seguramente
        'alcanzara a la reproduccion real
        
    End If
    
    
VerFrase:
    
    If mEsteEventoEsLargo Then
        EjecutarEvento = TmpEE
    Else
        EjecutarEvento = 0
    End If

    If Miliseg < Val(GetFraseTimeShow(LEF)) Then
        'si todavia no llego el tiempo de inicio de la frase es por que es la siguiente
            'no la actual
        SigFrase = GetFraseTexto(LEF)
        If HayQuePintar = False Then Exit Function
    Else
        'EMPEZO NUEVA FRASE!!!
        LetrasAcumFrase = Len(TextoActual)   'lo pongo en el total de la primera parte. Si o si cuamdo empieza una frase ya se cargo la primeras letras de esta frases

        FRASE = GetFraseTexto(LEF)
        'el tiempo de reproduccion paso el de la frase actual!!!
        LEF = LEF + 1
        'ver que no se pase del total!!!

        Dim J As Long
        For J = 0 To 2
            ReDim Preserve T(J)
            If (LEF + J) > MaxFrases Then
                T(J) = "FIN CANCION (" + CStr(7 - J) + ")"
            End If
            T(J) = GetFraseTexto(LEF + J)
        Next J
    End If

    ImprimirTxt FRASE, LetrasAcumFrase - Len(TextoActual) + 1, Len(TextoActual)

End Function

Private Sub ImprimirTxt(Texto As String, Empieza As Long, Largo As Long)
    Dim A As Long ' no se si es mas publica XXXX
    
    P.Cls

    P.ForeColor = ColorNormal: P.FontSize = TFU

    'sombreado de la letra! (mas de 30 no sirve!)
    P.CurrentX = ((P.Width / 2) - (P.TextWidth(Texto) / 2)) + Sombra
    P.CurrentY = (P.Height / 4) + Sombra
    
    Dim FCOL As Long
    FCOL = (ColorNormal + ColorSeleccionado) / 2
    If FCOL < 0 Then FCOL = vbRed
    
    P.ForeColor = FCOL
    P.Print Texto;
    
    P.CurrentX = (P.Width / 2) - (P.TextWidth(Texto) / 2)
    P.CurrentY = P.Height / 4
    If ColorNormal < 0 Then ColorNormal = vbWhite
    P.ForeColor = ColorNormal
    P.Print Texto;

    P.CurrentX = (P.Width / 2) - (P.TextWidth(Texto) / 2)
    P.CurrentY = P.Height / 4
    P.Print Mid(Texto, 1, Empieza - 1);

    P.ForeColor = ColorSeleccionado

    Pelota.left = P.CurrentX
    P.Print Mid(Texto, Empieza, Largo)

    'de pecho y como negrada las letras que siguen
    P.ForeColor = ColorNormal: P.FontSize = TFU
    Dim CurrentTop As Long
    'el ultimo renglon primero abajo
    CurrentTop = P.Height - (UBound(T) + 1) * (P.TextHeight(T(A)) + 150)
    For A = 0 To UBound(T)
    
        'sombreado de la letra! (mas de 30 no sirve!)
        P.CurrentX = (P.Width / 2) - (P.TextWidth(T(A)) / 2) + Sombra
        P.CurrentY = CurrentTop + Sombra
        P.ForeColor = (ColorNormal + ColorSeleccionado) / 2
        P.Print T(A)
    
        P.CurrentX = (P.Width / 2) - (P.TextWidth(T(A)) / 2)
        P.CurrentY = CurrentTop
        P.ForeColor = ColorNormal
        P.Print T(A)
        
        CurrentTop = CurrentTop + (P.TextHeight(T(A)) + 150)
    Next A
End Sub

Public Function GetTFU() 'obtener un tamaño de fuente que sirva para todas las frases
    Dim SE As Long
    'de 5 en 5 desde un maximo de 70
    TFU = 10 'menos de eso no uso
    Dim FRASE As String
    For SE = 70 To 10 Step -5
        If GetMaxWidthFrase(SE, FRASE) < P.Width Then
            TFU = SE
            Exit For
        End If
    Next SE
    
    If TFU > 30 Then
        Sombra = 30
    Else
        Sombra = TFU
    End If
    
    'MsgBox "Se definio TFU en:" + CStr(TFU) + vbCrLf + FRASE
End Function

Private Function GetMaxWidthFrase(TamanoFuente As Long, ByRef ReturnFrase As String)
    Dim SE As Long, MaxW As Long, FraseLarga As String
    MaxW = 0
    Dim L As String, SP() As String
    P.FontSize = TamanoFuente
    For SE = 0 To UBound(ListaEvFraces)
        SP = Split(ListaEvFraces(SE), Chr(6))
        'tambien hay ""
        If UBound(SP) > 0 Then
            L = SP(1)
            If P.TextWidth(L) > MaxW Then
                MaxW = P.TextWidth(L)
                FraseLarga = L
            End If
        End If
    Next SE
    ReturnFrase = FraseLarga
    GetMaxWidthFrase = MaxW
End Function


'Configura los colores de las lettttras
Private Sub SetFormato()
    Dim Datos As String
    
    Datos = Space(FileLen(nTMP))
    
    Open nTMP For Binary As #1
        Get #1, 1, Datos
    Close
    
    Dim C() As String
    C = Split(Datos, "?")
    
    ColorNormal = Val(C(0))
    ColorSeleccionado = Val(C(1))
    P.Font = C(2)
    P.FontSize = C(3)
End Sub

Private Sub Class_Initialize()
    CarpetaTemp
    Set P = Nothing
    mEsteEventoEsLargo = True 'el primero lo tomo como largo por las dudas
End Sub

Public Sub DefinePathLog(Pth As String)
    Terr.FileLog = Pth
    'en realidad nunca uso el append ya que le paso el detalle al programa usuario de la DLL para que haga lo que quiera!
    Terr.LargoAcumula = 300
    Terr.Anotar "Kar001", Pth
End Sub

'Private Sub IrATiempo(xMiliSeg As Long)
'    'cuando se mueva el mp3 de posicion muevo el texto tambien
'    RelojControl = RelojControl - (xMiliSeg / 1000)
'    IrA CSng(xMiliSeg / 1000)
'End Sub

Private Sub CarpetaTemp()
    If Right(App.Path, 1) = "\" Then
        cTMP = App.Path + ""
    Else
        cTMP = App.Path + "\"
    End If
    
    cTMP = cTMP + "Pemt\"
    
    If Dir(cTMP, vbDirectory) = "" Then
        MkDir cTMP
    End If
End Sub

Private Function LeeKar(Archivo As String, Optional CambiarTiempos As Long = 0) As Long
    '0 si es OK
    '1 si no existe
    '2 si no es neokaraoke cero
    '3 desconocido
    Terr.Anotar "1011"
    mFile = Archivo

    On Local Error GoTo ErrKar
    Dim sKAR As String
    If Dir(mFile) = "" Then
        LeeKar = 1
        Exit Function
    End If

    Dim Activ As Boolean
    Activ = False
    
    sKAR = Space(FileLen(mFile))
    Open mFile For Binary As #1
        Get #1, 1, sKAR
    Close
    
    Terr.Anotar "1012"
    'ok
    LeeKar = 0

    If Mid(sKAR, 1, 6) <> "NeoKar" Then
        sKAR = "NeoKar0" + Chr(5) + sKAR
    End If

    sKAR = Mid(sKAR, 7)
    ListaEventos = Split(sKAR, Chr(6))

    ReDim ListaEvFraces(0)
    Dim A As Long
    Dim eB As Long
    eB = 0
    
    Dim DosDatos() As String
    LEF = 0
    For A = 0 To UBound(ListaEventos)
        If ListaEventos(A) = "" Then GoTo SIG
        DosDatos = Split(ListaEventos(A), Chr(5))
        '------------------------------------------
        '------------------------------------------
        'ver si hay que cambiar los tiempos
        If CambiarTiempos <> 0 Then
            DosDatos(0) = CStr(CLng(DosDatos(0)) + CambiarTiempos)
            ListaEventos(A) = DosDatos(0) + Chr(5) + DosDatos(1)
        End If
        '------------------------------------------
        '------------------------------------------
        Dim ChINI As String
        ChINI = Mid(DosDatos(1), 1, 1)
        If ChINI = "\" Or ChINI = "/" Then
            If Not InStr(DosDatos(1), "\") = 0 Then
                ListaEventos(A) = Replace(ListaEventos(A), "\", "")
            Else
                ListaEventos(A) = Replace(ListaEventos(A), "/", "")
            End If
            LEF = LEF + 1 'cantidad de frases
            ReDim Preserve ListaEvFraces(LEF)
            ListaEvFraces(LEF) = DosDatos(0) + Chr(6) + Mid(DosDatos(1), 2)
            eB = 0
        End If
        If eB = 1 Then
            ListaEvFraces(LEF) = ListaEvFraces(LEF) + DosDatos(1)
        End If
        eB = 1
SIG:
    Next A
    Terr.Anotar "1013"
    Exit Function
ErrKar:
    LeeKar = 3
End Function

Public Function GetLOG() As String
    GetLOG = Terr.LogAcumulado
End Function

Public Property Get MaxEventos() As Long
    MaxEventos = UBound(ListaEventos)
End Property

Public Property Get MaxFrases() As Long
    MaxFrases = UBound(ListaEvFraces)
End Property

Private Function GetLetra(Index As Long) As String
    If Index > UBound(ListaEventos) Then
        GetLetra = ""
        Exit Function
    End If
    Dim SP() As String
    If ListaEventos(Index) = "" Then
        GetLetra = ""
        Exit Function
    End If
    SP = Split(ListaEventos(Index), Chr(5))
    GetLetra = SP(1)
End Function

Private Function GetTimeShow(Index As Long) As Long

    If Index > UBound(ListaEventos) Then
        GetTimeShow = -1
        Exit Function
    End If

    If ListaEventos(Index) = "" Then
        GetTimeShow = -1
        Exit Function
    End If

    Dim SP() As String
    SP = Split(ListaEventos(Index), Chr(5))
    GetTimeShow = Val(SP(0))

End Function

Private Function GetFraseTexto(Index As Long) As String
    If Index > UBound(ListaEvFraces) Then
        GetFraseTexto = ""
        Exit Function
    End If
    If ListaEvFraces(Index) = "" Then
        GetFraseTexto = ""
        Exit Function
    End If
    Dim SP() As String
    SP = Split(ListaEvFraces(Index), Chr(6))
    If UBound(SP) = 0 Then
        GetFraseTexto = ""
    Else
        GetFraseTexto = SP(1)
    End If
End Function

Private Function GetFraseTimeShow(Index As Long) As String
    If Index > UBound(ListaEvFraces) Then
        GetFraseTimeShow = -1
        Exit Function
    End If
    If ListaEvFraces(Index) = "" Then
        GetFraseTimeShow = -1
        Exit Function
    End If
    Dim SP() As String
    SP = Split(ListaEvFraces(Index), Chr(6))
    GetFraseTimeShow = SP(0)
End Function

Public Sub DefinePictureBox(PIC As Object)
    Set P = PIC
End Sub

Public Sub DefinePelota(PIC2 As Object)
    Set Pelota = PIC2
End Sub

Public Sub BorraTemp()

    If Dir(cTMP, vbDirectory) <> "" Then
        Dim Archa As String
        Archa = Dir(cTMP)
        While Archa <> ""
            FS.DeleteFile cTMP + Archa, True
            Archa = Dir
        Wend
    End If
End Sub

Public Sub CleanPic()
    'limpia todo si ya estaba definido
    On Local Error Resume Next
    P.Picture = LoadPicture
    P.Cls
    Pelota.Visible = False
End Sub
