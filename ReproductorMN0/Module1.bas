Attribute VB_Name = "Module1"
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
            'si no lo hago se graba sin frases y por lo taçnto la matriz de frases
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
