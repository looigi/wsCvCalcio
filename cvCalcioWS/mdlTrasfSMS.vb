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

        ' Routine che esegue una query sul db
        Try
            Conn.Execute(Sql)
        Catch ex As Exception
            Ritorno = StringaErrore & " " & ex.Message
        End Try

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Public Function LeggeQuery(ByVal Conn As Object, ByVal Sql As String, ByVal Connessione As String) As Object
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn, Connessione)
        Dim Rec As Object = CreateObject("ADODB.Recordset")

        Try
            Rec.Open(Sql, Conn)
        Catch ex As Exception
            Rec = StringaErrore & " " & ex.Message
        End Try

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

