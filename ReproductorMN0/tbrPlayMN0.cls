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



Dim TimerEmpezo As Long 'timer exacto de cuando empieza
Private RelojControl As Single 'momento actual del timer tambien

Private DireccionPelota As Long 'si la pelota de los karaokes sube o baja
Private TopPelota As Long 'altura de la pelota
Private LetrasAcumFrase As Long 'cantidad de letras que pasaron sin empezar una _

Private Frase As String 'frase actual del NK0
Private SigFrase As String ' frase siguiente del NK0
Private LE As Long 'numero de evento que se ejecuta
Private LEF As Long 'numero de frases del NK0

Private T() As String 'lista de renglonres por cantarse

Private ColorNormal As OLE_COLOR
Private ColorSeleccionado As OLE_COLOR

Dim cTMP As String 'carpeta de temporales`para colocar los archivos y borrarlos

Dim nNK0 As String 'archivo NK0 dentro del MN0
Dim nMP3 As String 'mp3 adentro
Dim nTMP As String 'detalle de fuentes y colores
Dim nIMG As String 'imagen

'reloj virtual mejor que el de VB
Private WithEvents TIM As tbrTimer.clsTimer
Attribute TIM.VB_VarHelpID = -1
Private P As Object 'algun objeto frm o pBox para dibujar todo

Public Sub Cerrar()
    Reset 'para el timer y cerrar el mp3
End Sub

Public Sub Cargar(Archivo As String) 'separar el mn0 y comenzar!
    Reset 'para el timer y cerrar el mp3
    SepararMN0 Archivo 'graba los archivos dentro del mn0 en la carpeta temp
    SetFormato 'define la fuente y los tama�os del mn0
    Ejecutar nNK0, nMP3, nIMG 'empezar la reproduccion
End Sub

Sub BorraTemp()
    If Dir(cTMP, vbDirectory) <> "" Then
        Dim Archa As String
        Archa = Dir(cTMP)
        While Archa <> ""
            If Dir(cTMP + Archa) <> "" Then
                Kill cTMP + Archa
                Archa = Dir(cTMP)
            End If
        Wend
    End If
End Sub

Private Sub SepararMN0(Archivo As String)
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
    For r = 0 To UBound(qDatos) - 1
        Dqs = Split(qDatos(r), "?")
        Datoo = Space$(Dqs(1))

        Dim ArchivoDestino As String
        Select Case r
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

        If Not Dir(ArchivoDestino) = "" Then Kill ArchivoDestino

        Open ArchivoDestino For Binary As #2
            Get #1, Pdv, Datoo
            Put #2, 1, Datoo
        Close #2
        Pdv = Pdv + Dqs(1) + 1
        DoEvents
    Next r
Close
End Sub

Private Sub Ejecutar(ArchNK0 As String, ArchMP3 As String, ArchIMG As String)
    DireccionPelota = 0

    LeeKar ArchNK0 'hacer una lista de eventos

    'poner la imagen que el usuario eligio
    P.PaintPicture LoadPicture(ArchIMG), 0, 0, P.Width, P.Height
    Dim BT As String
    BT = App.path + "\tmp.bmp"
    If Dir(BT) <> "" Then Kill BT
    SavePicture P.Image, BT
    'P.Picture = LoadPicture(ArchIMG)
    P.Picture = LoadPicture(BT) '<<<<<
    
    LE = 0: LEF = 0 'numero de evento y de frase en cero
    'abrir el mp3
    Abrir nMP3
    'y ejecutarlo
    vPlay
    'arranca ...
    RelojControl = Timer
    TimerEmpezo = RelojControl
    TIM.Interval = 20
    TIM.Enabled = True

End Sub

Private Sub EjecutarEvento()

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
    Pelota.Top = TopPelota
    '-------------------------------------------------
    '-------------------------------------------------
    'TRUCAHDA MIA

'    Picture1.Cls
'    Picture1.Print CStr(LE) + "-" + CStr(Miliseg) + "-" + CStr(CLng((Timer - RelojControl) * 1000)) + "-" + _
'        CStr(NK0.GetFraseTimeShow(LEF)) + "-" + CStr(NK0.GetTimeShow(LE))
'
    Dim Miliseg As Long

    Miliseg = CLng((Timer - RelojControl) * 1000)

    'AHORA LA LETRA
    'si no llego el tiempo salir antes
    Dim HayQuePintar As Boolean 'si salteo la letra por que no la necesito
    'y voy a ver si no necesito la frase ya se que no hay que pintar si la frase tampoco necesita nada
 
    Dim TextoActual As String 'texto que se esta marcando

    If Miliseg < GetTimeShow(LE) Then
        HayQuePintar = False
        GoTo VerFrase
    Else
        HayQuePintar = True
        'ver que no se pase del total!!!
        If LE > MaxEventos Then
            Exit Sub
        Else
            'si es una letra valida es la que se va a mostrar
            'por lo tanto le sumo las letras que tiene
            TextoActual = GetLetra(LE)
            LetrasAcumFrase = LetrasAcumFrase + Len(TextoActual)
        End If
        LE = LE + 1
        'si no se termino dejo que se muestre
        'si hay algun atraso esto se ejecutara cad 20 miliseg y seguramente
        'alcanzara a la reproduccion real
    End If
VerFrase:
    If Miliseg < Val(GetFraseTimeShow(LEF)) Then
        'si todavia no llego el tiempo de inicio de la frase es por que es la siguiente
            'no la actual
        SigFrase = GetFraseTexto(LEF)
        If HayQuePintar = False Then Exit Sub
    Else
        'EMPEZO NUEVA FRASE!!!
        LetrasAcumFrase = Len(TextoActual)   'lo pongo en el total de la primera parte. Si o si cuamdo empieza una frase ya se cargo la primeras letras de esta frases

        Frase = GetFraseTexto(LEF)
        'el tiempo de reproduccion paso el de la frase actual!!!
        LEF = LEF + 1
        'ver que no se pase del total!!!

        Dim J As Long
        For J = 0 To 7
            ReDim Preserve T(J)
            If (LEF + J) > MaxFrases Then
                T(J) = "FIN CANCION (" + CStr(7 - J) + ")"
            End If
            T(J) = GetFraseTexto(LEF + J)
        Next J
    End If

    ImprimirTxt Frase, LetrasAcumFrase - Len(TextoActual) + 1, Len(TextoActual)

End Sub

Private Sub ImprimirTxt(Texto As String, Empieza As Long, Largo As Long)
    P.Cls

    P.ForeColor = ColorNormal: P.FontSize = 30

    P.CurrentX = (P.Width / 2) - (P.TextWidth(Texto) / 2)
    P.CurrentY = P.Height / 4
    P.Print Texto;

    P.CurrentX = (P.Width / 2) - (P.TextWidth(Texto) / 2)
    P.CurrentY = P.Height / 4
    P.Print Mid(Texto, 1, Empieza - 1);

    P.ForeColor = ColorSeleccionado

    Pelota.Left = P.CurrentX
    P.Print Mid(Texto, Empieza, Largo)

    'de pecho y cmo negrada las letras que siguen
    P.ForeColor = ColorNormal: P.FontSize = 20
    Dim CurrentTop As Long
    'el ultimo renglon primero abajo
    CurrentTop = P.Height - (UBound(T) + 1) * (P.TextHeight(T(A)) + 50)
    For A = 0 To UBound(T)
        P.CurrentX = (P.Width / 2) - (P.TextWidth(T(A)) / 2)
        P.CurrentY = CurrentTop
        P.Print T(A)
        CurrentTop = CurrentTop + (P.TextHeight(T(A)) + 50)
    Next A
End Sub

'Configura los colores de las lettttras
Sub SetFormato()
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

Sub Reset()
    vCerrar ' close la cancion reproduciendo
    ColorNormal = 0
    ColorSeleccionado = 0
    TIM.Enabled = False
End Sub


Function hms(Mili As Single) As String
    If Mili > 0 Then
    
        Dim H As Long
        Dim M As Long
        Dim S As Long
        Dim Z As Long
    
        Dim TMPSec As Long
        Z = (Mili Mod 1000)
        TMPSec = (Mili - Z) / 1000
        S = (TMPSec Mod 60)
        TMPSec = TMPSec - S
        M = TMPSec / 60
        If M > 59 Then
            TMPSec = TMPSec - (M * 60)
            H = Fix(M / 60)
            M = M Mod 60
        Else
            H = 0
        End If
    
    End If
    hms = Format(CStr(H), "00") + ":" + Format(CStr(M), "00") + ":" + Format(CStr(S), "00") '+ ":" + Format(CStr(Z), "00")

End Function

Private Sub Form_Load()
    CarpetaTemp
End Sub

Private Sub Tim_Timer()
    EjecutarEvento
End Sub

Public Sub IrATiempo(xMiliSeg As Long)
    'cuando se mueva el mp3 de posicion muevo el texto tambien
    RelojControl = RelojControl - (xMiliSeg / 1000)
    IrA CSng(xMiliSeg / 1000)
End Sub

Sub CarpetaTemp()
    If Right(App.path, 1) = "\" Then
        cTMP = App.path + ""
    Else
        cTMP = App.path + "\"
    End If
    
    cTMP = cTMP + "TEMP\"
    
    If Dir(cTMP, vbDirectory) = "" Then
        MkDir cTMP
    End If
End Sub



Private ListaEventos() As String
Private ListaEvFraces() As String
Private mFile As String 'archivo NK0 abierto

Public Function LeeKar(Archivo As String, Optional CambiarTiempos As Long = 0) As Long
    '0 si es OK
    '1 si no existe
    '2 si no es neokaraoke cero
    '3 desconocido

    mFile = Archivo

    On Local Error GoTo ErrKar
    Dim Kar As String
    If Dir(mFile) = "" Then
        LeeKar = 1
        Exit Function
    End If

    Dim Activ As Boolean
    Activ = False

    Kar = Space(FileLen(mFile))
    Open mFile For Binary As #1
        Get #1, 1, Kar
    Close

    'ok
    LeeKar = 0

    If Mid(Kar, 1, 6) <> "NeoKar" Then
        Kar = "NeoKar0" + Chr(5) + Kar
    End If


    Kar = Mid(Kar, 7)
    ListaEventos = Split(Kar, Chr(6))

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

    Exit Function
ErrKar:
    LeeKar = 3
End Function

Public Property Get MaxEventos() As Long
    MaxEventos = UBound(ListaEventos)
End Property

Public Property Get MaxFrases() As Long
    MaxFrases = UBound(ListaEvFraces)
End Property

Public Function GetLetra(Index As Long) As String
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

Public Function GetTimeShow(Index As Long) As Long

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

Public Function GetFraseTexto(Index As Long) As String
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

Public Function GetFraseTimeShow(Index As Long) As String
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

Public Sub GrabarArchNK0(ArchivoOUT As String)
    'graba el nk0
    
    Dim nEvFrase As Long 'contador para las frases que ya se leyeron
    nEvFrase = 1
    Dim NK As String, G As Long
    NK = ""
    For G = 0 To UBound(ListaEventos)
        'esto no deberia ser necesario
        If ListaEventos(G) <> "" Then
            'ver si hay que poner una barra por el separador de frases!!!
            'si no lo hago se graba sin frases y por lo ta�nto la matriz de frases
            'se pierde y da error ya que la necesita
            Dim T1 As Long 'tiempo del evento de texto actual
            'para comparar con el tiempo de las frases
            'cuando son iguales es que debo agregar una barra para identificar la frase
            Dim SP() As String
            SP = Split(ListaEventos(G), Chr(5))
            Dim PonerBarra As Boolean
            PonerBarra = False
            If nEvFrase < MaxFrases Then
                If SP(0) = GetFraseTimeShow(nEvFrase) Then
                    PonerBarra = True
                    nEvFrase = nEvFrase + 1
                End If
            End If
            'poner la barra en su lugar
            If PonerBarra Then
                NK = NK + SP(0) + Chr(5) + "/" + SP(1)
            Else
                NK = NK + ListaEventos(G)
            End If
            If G < UBound(ListaEventos) Then NK = NK + Chr(6) 'separador irrepetible
        End If
    Next G
    
    Dim H As String
    H = "NeoKar"

    If Dir(ArchivoOUT) <> "" Then Kill ArchivoOUT

    Open ArchivoOUT For Binary As #1
        Put #1, 1, H + NK
    Close
End Sub

Public Sub DefinePictureBox(PIC As Object)
    Set P = PIC
End Sub

