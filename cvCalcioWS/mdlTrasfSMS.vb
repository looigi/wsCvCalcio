Imports System.Timers

Module mdlTrasfSMS
    Public ultimaNotifica As String = ""
    Public ultimoMessaggio As String = ""
    Public ultimaPosizione As String = ""

    Public Function SistemaStringa(ByVal Cosa As String) As String
        Dim Ritorno As String = Cosa

        Ritorno = MetteMaiuscoleDopoOgniSpazio(Ritorno)
        Ritorno = RaddoppiaApici(Ritorno)

        Return Ritorno
    End Function

    Public Function MetteMaiuscoleDopoOgniSpazio(ByVal Cosa As String) As String
        Dim Appoggio As String
        Dim Ritorno As String
        Dim I As Integer

        Appoggio = LCase(Cosa)
        Ritorno = UCase(Mid(Appoggio, 1, 1))
        For I = 2 To Len(Appoggio)
            Ritorno = Ritorno & Mid(Appoggio, I, 1)
            If Mid(Appoggio, I, 1) = " " Or Mid(Appoggio, I, 1) = "'" Then
                Ritorno = Ritorno & UCase(Mid(Appoggio, I + 1, 1))
                I = I + 1
            End If
        Next I

        MetteMaiuscoleDopoOgniSpazio = Ritorno
    End Function

    Public Function RaddoppiaApici(ByVal Cosa As String) As String
        Dim Ritorno As String = Cosa.Trim

        Ritorno = Ritorno.Replace("'", "''")

        Return Ritorno
    End Function

    Public Function ApreDB(ByVal Connessione As String) As Object
        ' Routine che apre il DB e vede se ci sono errori
        Dim Conn As Object = CreateObject("ADODB.Connection")

        Try
            Conn.Open(Connessione)
            Conn.CommandTimeout = 0
        Catch ex As Exception
            Conn = StringaErrore & " " & ex.Message
        End Try

        Return Conn
    End Function

    Public Function EsegueSql(ByVal Conn As Object, ByVal Sql As String, ByVal Connessione As String) As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn, Connessione)
        Dim Ritorno As String = "*"
        Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
        Dim gf As New GestioneFilesDirectory

        If effettuaLog Then
            'If nomeFileLogGenerale = "" Then
            Dim paths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
            Dim pp() As String = paths.Split(";")
            pp(1) = pp(1).Replace(vbCrLf, "")
            If Strings.Right(pp(1), 1) <> "\" Then
                pp(1) = pp(1) & "\"
            End If
            nomeFileLogGenerale = pp(1) & "logWS_Exec_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"
            ThreadScriveLog(Datella & "--------------------------------------------------------------------------")
            ThreadScriveLog(Datella & ": " & Sql)
            ' End If
        End If

        ' Routine che esegue una query sul db
        Try
            Conn.Execute(Sql)
        Catch ex As Exception
            Ritorno = StringaErrore & " " & ex.Message
            If effettuaLog Then
                ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message)
            End If
        End Try
        If effettuaLog Then
            ThreadScriveLog(Datella & "--------------------------------------------------------------------------")
            ThreadScriveLog("")
        End If

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Private Sub ThreadScriveLog(s As String)
        listaLog.Add(s)

        avviaTimerLog()
    End Sub

    Private Sub avviaTimerLog()
        If timerLog Is Nothing Then
            timerLog = New Timer(100)
            AddHandler timerLog.Elapsed, New ElapsedEventHandler(AddressOf scodaLog)
            timerLog.Start()
        End If
    End Sub

    Private Sub scodaLog()
        timerLog.Enabled = False
        Dim sLog As String = listaLog.Item(0)

        Dim gf As New GestioneFilesDirectory
        gf.ApreFileDiTestoPerScrittura(nomeFileLogGenerale)
        gf.ScriveTestoSuFileAperto(sLog)
        gf.ChiudeFileDiTestoDopoScrittura()

        listaLog.RemoveAt(0)
        If listaLog.Count > 0 Then
            timerLog.Enabled = True
        Else
            timerLog = Nothing
            listaLog = New List(Of String)
        End If
    End Sub

    Public Function LeggeQuery(ByVal Conn As Object, ByVal Sql As String, ByVal Connessione As String) As Object
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn, Connessione)
        Dim Rec As Object = CreateObject("ADODB.Recordset")
        Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
        Dim gf As New GestioneFilesDirectory

        If effettuaLog Then
            'If nomeFileLogGenerale = "" Then
            Dim paths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
            Dim pp() As String = paths.Split(";")
            pp(1) = pp(1).Replace(vbCrLf, "")
            If Strings.Right(pp(1), 1) <> "\" Then
                pp(1) = pp(1) & "\"
            End If
            nomeFileLogGenerale = pp(1) & "logWS_Query_" & Now.Day & "_" & Now.Month & "_" & Now.Year & ".txt"
            ThreadScriveLog(Datella & "--------------------------------------------------------------------------")
            ThreadScriveLog(Datella & ": " & Sql)
            'End If
        End If

        Try
            Rec.Open(Sql, Conn)
        Catch ex As Exception
            Rec = StringaErrore & " " & ex.Message
            If effettuaLog Then
                ThreadScriveLog(Datella & ": ERRORE SQL -> " & ex.Message)
            End If
        End Try
        If effettuaLog Then
            ThreadScriveLog(Datella & "--------------------------------------------------------------------------")
            ThreadScriveLog("")
        End If

        ChiudeDB(AperturaManuale, Conn)

        Return Rec
    End Function

    Private Function ControllaAperturaConnessione(ByRef Conn As Object, ByVal Connessione As String) As Boolean
        Dim Ritorno As Boolean = False

        If Conn Is Nothing Then
            Ritorno = True
            Conn = ApreDB(Connessione)
        End If

        Return Ritorno
    End Function

    Public Sub ChiudeDB(ByVal TipoApertura As Boolean, ByRef Conn As Object)
        If TipoApertura = True Then
            Conn.Close()
        End If
    End Sub
End Module

