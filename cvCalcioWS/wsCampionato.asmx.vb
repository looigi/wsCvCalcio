Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_cam.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsCampionato
    Inherits System.Web.Services.WebService


    <WebMethod()>
    Public Function RitornaCampionatoCategoria(ByVal idAnno As String, idCategoria As String, idUtente As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
                Dim Rec2 As Object = Server.CreateObject("ADODB.Recordset")
                Dim Sql As String = ""

                Dim idSquadre As New ArrayList
                Dim Squadre As New ArrayList

                Dim idGiornata As String = RitornaGiornataUtenteCategoria(idUtente, idAnno, idCategoria)

                Ritorno = ""

                ' Giornata
                Ritorno &= "^"
                Ritorno &= idGiornata.ToString
                Ritorno &= "^"

                Dim NomeSquadra As String

                Sql = "Select * From Anni Where idAnno=" & idAnno
                Rec = LeggeQuery(Conn, Sql, Connessione)
                If TypeOf (Rec) Is String Then
                    Ritorno = StringaErrore & "" & Rec
                    Return Ritorno
                Else
                    NomeSquadra = Rec("NomeSquadra").Value
                End If
                Rec.Close

                Try
                    ' Squadre avversarie
                    Sql = "SELECT AvversariCalendario.idAvversario As idAvv, SquadreAvversarie.Descrizione As Squadra, CampiAvversari.idCampo As idCampo, CampiAvversari.Descrizione As Campo, " &
                        "CampiAvversari.Indirizzo As Indirizzo, AvversariCoord.Lat, AvversariCoord.Lon " &
                        "FROM ((AvversariCalendario LEFT JOIN SquadreAvversarie ON AvversariCalendario.idAvversario = SquadreAvversarie.idAvversario) " &
                        "Left Join CampiAvversari On SquadreAvversarie.idCampo = CampiAvversari.idCampo) " &
                        "Left Join AvversariCoord On SquadreAvversarie.idAvversario = AvversariCoord.idAvversario " &
                        "WHERE AvversariCalendario.idAnno=" & idAnno & " And AvversariCalendario.idCategoria=" & idCategoria & " " &
                        "ORDER BY AvversariCalendario.idProgressivo;"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = StringaErrore & "" & Rec
                        Return Ritorno
                    Else
                        ' Aggiungo la riga per i dati della categoria
                        idSquadre.Add(-idCategoria)
                        Squadre.Add(NomeSquadra)

                        Ritorno &= "£"

                        Do Until Rec.Eof
                            idSquadre.Add(Rec("idAvv").Value)
                            Squadre.Add(Rec("Squadra").Value)

                            Ritorno &= Rec("idAvv").Value & ";" &
                                    Rec("Squadra").Value & ";" &
                                    Rec("idCampo").Value & ";" &
                                    Rec("Campo").Value.ToString.Replace(";", ",") & ";" &
                                    Rec("Indirizzo").Value.ToString.Replace(";", ",") & ";" &
                                    Rec("Lat").Value & ";" &
                                    Rec("Lon").Value & ";" &
                                    "§"
                            Rec.MoveNext()
                        Loop
                        Rec.Close()

                        Ritorno &= "£"
                    End If

                    ' Partite
                    Sql = "SELECT CalendarioPartite.idGiornata As idGiornata, CalendarioPartite.idPartita As idPartita, CalendarioPartite.idPartitaGen As idPartitaGen, " &
                        "CalendarioPartite.idSqCasa As idSqCasa, CalendarioPartite.idSqFuori As idSqFuori, " &
                        "CalendarioDate.Datella AS Datella, SquadreAvversarie.Descrizione As Casa, SquadreAvversarie_1.Descrizione As Fuori, RisultatiAggiuntivi.RisGiochetti As RisGiochetti, " &
                        "RisultatiAggiuntivi.GoalAvvPrimoTempo As GoalAvv1, RisultatiAggiuntivi.GoalAvvSecondoTempo As GoalAvv2, RisultatiAggiuntivi.GoalAvvTerzoTempo As GoalAvv3, " &
                        "CalendarioPartite.Risultato As Risultato1, Risultati.Risultato As Risultato2, Risultati.Note As Notelle, Partite.Casa As InCasa, Partite.OraConv As OraConv, " &
                        "CalendarioDate.Datella, CalendarioPartite.Giocata, Partite.idPartita As PartitaUfficiale " &
                        "FROM ((((((CalendarioPartite LEFT JOIN CalendarioDate ON CalendarioPartite.idAnno = CalendarioDate.idAnno And CalendarioPartite.idCategoria = CalendarioDate.idCategoria " &
                        "And CalendarioPartite.idGiornata = CalendarioDate.idGiornata And CalendarioPartite.idPartita = CalendarioDate.idPartita) " &
                        "LEFT JOIN SquadreAvversarie ON CalendarioPartite.idSqCasa = SquadreAvversarie.idAvversario) " &
                        "LEFT JOIN SquadreAvversarie AS SquadreAvversarie_1 ON CalendarioPartite.idSqFuori = SquadreAvversarie_1.idAvversario) " &
                        "LEFT JOIN Partite ON CalendarioPartite.idPartitaGen = Partite.idUnioneCalendario And CalendarioPartite.idCategoria = Partite.idCategoria) " &
                        "LEFT JOIN Risultati ON Partite.idPartita = Risultati.idPartita) " &
                        "LEFT JOIN RisultatiAggiuntivi ON Partite.idPartita = RisultatiAggiuntivi.idPartita) " &
                        "WHERE CalendarioPartite.idCategoria=" & idCategoria & " And CalendarioPartite.idAnno=" & idAnno & " " &
                        "ORDER BY CalendarioPartite.idGiornata, CalendarioPartite.idPartita;"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = StringaErrore & "" & Rec
                        Return Ritorno
                    Else
                        Do Until Rec.Eof
                            Dim GoalAvv As Integer = 0
                            If Val("" & Rec("GoalAvv1").Value) > 0 Then
                                GoalAvv += ("" & Rec("GoalAvv1").Value)
                            End If
                            If Val("" & Rec("GoalAvv2").Value) > 0 Then
                                GoalAvv += ("" & Rec("GoalAvv2").Value)
                            End If
                            If Val("" & Rec("GoalAvv3").Value) > 0 Then
                                GoalAvv += ("" & Rec("GoalAvv3").Value)
                            End If

                            Ritorno &= Rec("idGiornata").Value & ";" &
                                Rec("idPartita").Value & ";" &
                                Rec("idPartitaGen").Value & ";" &
                                Rec("idSqCasa").Value & ";" &
                                Rec("idSqFuori").Value & ";" &
                                Rec("Datella").Value & ";" &
                                Rec("Casa").Value & ";" &
                                Rec("Fuori").Value & ";" &
                                Rec("RisGiochetti").Value & ";" &
                                GoalAvv.ToString & ";" &
                                Rec("Risultato1").Value & ";" &
                                Rec("Risultato2").Value & ";" &
                                Rec("Notelle").Value.ToString.Replace(";", ",") & ";" &
                                Rec("InCasa").Value & ";" &
                                Rec("OraConv").Value & ";" &
                                Rec("Datella").Value & ";" &
                                Rec("Giocata").Value & ";"

                            '' Prende i convocati
                            'Ritorno &= "!"
                            'If "" & Rec("PartitaUfficiale").Value <> "" Then
                            '    Sql = "SELECT Giocatori.idGiocatore, Giocatori.Cognome As Cognome, Giocatori.Nome As Nome, Giocatori.idRuolo As idRuolo, Ruoli.Descrizione As Ruolo " &
                            '    "FROM ((Convocati LEFT JOIN Giocatori ON Convocati.idGiocatore = Giocatori.idGiocatore) " &
                            '    "LEFT JOIN Ruoli ON Giocatori.idRuolo = Ruoli.idRuolo) " &
                            '    "WHERE Convocati.idPartita=" & Rec("PartitaUfficiale").Value & " And idAnno=" & idAnno
                            '    Rec2 = LeggeQuery(Conn, Sql, Connessione)
                            '    If TypeOf (Rec2) Is String Then
                            '        Ritorno = Rec2
                            '        Return Ritorno
                            '    Else
                            '        Do Until Rec2.Eof
                            '            Ritorno &= Rec2("idGiocatore").Value & ";"
                            '            Ritorno &= Rec2("Cognome").Value & ";"
                            '            Ritorno &= Rec2("Nome").Value & ";"
                            '            Ritorno &= Rec2("idRuolo").Value & ";"
                            '            Ritorno &= Rec2("Ruolo").Value & ";"

                            '            Rec2.MoveNext()
                            '        Loop
                            '        Rec2.Close()
                            '    End If
                            'End If
                            'Ritorno &= "!"
                            'Ritorno &= ";"

                            '' Prende i marcatori
                            'Ritorno &= "*"
                            'If "" & Rec("PartitaUfficiale").Value <> "" Then
                            '    Sql = "Select * From ( " &
                            '        "SELECT Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, RisultatiAggiuntiviMarcatori.Minuto, Progressivo " &
                            '        "FROM RisultatiAggiuntiviMarcatori LEFT JOIN Giocatori ON RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore " &
                            '        "WHERE RisultatiAggiuntiviMarcatori.idPartita=" & Rec("idPartita").Value & " And idAnno=" & idAnno & " " &
                            '        "Union All " &
                            '        "SELECT Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Marcatori.Minuto, idProgressivo As Progressivo " &
                            '        "FROM Marcatori LEFT JOIN Giocatori ON Marcatori.idGiocatore = Giocatori.idGiocatore " &
                            '        "WHERE Marcatori.idPartita=" & Rec("PartitaUfficiale").Value & " And idAnno=" & idAnno & " " &
                            '        ") A Order By Minuto , Progressivo"
                            '    Rec2 = LeggeQuery(Conn, Sql, Connessione)
                            '    If TypeOf (Rec2) Is String Then
                            '        Ritorno = StringaErrore & "" & Rec2
                            '        Return Ritorno
                            '    Else
                            '        Do Until Rec2.Eof
                            '            Ritorno &= Rec2("idGiocatore").Value & ";"
                            '            Ritorno &= Rec2("Cognome").Value & ";"
                            '            Ritorno &= Rec2("Nome").Value & ";"
                            '            Ritorno &= Rec2("Minuto").Value & ";"

                            '            Rec2.MoveNext()
                            '        Loop
                            '        Rec2.Close()
                            '    End If
                            'End If
                            'Ritorno &= "*"


                            Ritorno &= "§"

                            Rec.MoveNext()
                        Loop
                    End If

                    Rec.Close()
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function CalcolaClassificaAllaGiornata(ByVal idAnno As String, idCategoria As String, idGiornata As String, idUtente As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
                Dim Sql As String = ""
                Dim ProgressivoSquadra As String = ""

                Dim idSquadre As New ArrayList
                Dim Squadre As New ArrayList
                Dim Giocate As New ArrayList
                Dim Vinte As New ArrayList
                Dim Pareggiate As New ArrayList
                Dim Perse As New ArrayList
                Dim Punti As New ArrayList
                Dim gFatti As New ArrayList
                Dim gSubiti As New ArrayList

                Dim CeRisultato As Boolean = False
                Dim g1 As Integer = 0
                Dim g2 As Integer = 0
                Dim gR1 As Integer = 0
                Dim gR2 As Integer = 0

                Dim NomeSquadra As String

                Sql = "Select * From Anni Where idAnno=" & idAnno
                Rec = LeggeQuery(Conn, Sql, Connessione)
                If TypeOf (Rec) Is String Then
                    Ritorno = StringaErrore & "" & Rec
                    Return Ritorno
                Else
                    NomeSquadra = Rec("NomeSquadra").Value
                End If
                Rec.Close

                Sql = "Select SquadreAvversarie.idAvversario, SquadreAvversarie.Descrizione From (AvversariCalendario " &
                    "LEFT JOIN SquadreAvversarie As SquadreAvversarie On AvversariCalendario.idAvversario = SquadreAvversarie.idAvversario) " &
                    "Where idAnno=" & idAnno & " And idCategoria=" & idCategoria
                Rec = LeggeQuery(Conn, Sql, Connessione)
                If TypeOf (Rec) Is String Then
                    Ritorno = StringaErrore & "" & Rec
                    Return Ritorno
                Else
                    idSquadre.Add(-Val(idCategoria))
                    Squadre.Add(NomeSquadra)
                    Giocate.Add(0)
                    Vinte.Add(0)
                    Pareggiate.Add(0)
                    Perse.Add(0)
                    Punti.Add(0)
                    gFatti.Add(0)
                    gSubiti.Add(0)

                    Do Until Rec.Eof
                        idSquadre.Add(Rec("idAvversario").Value)
                        Squadre.Add(Rec("Descrizione").Value)
                        Giocate.Add(0)
                        Vinte.Add(0)
                        Pareggiate.Add(0)
                        Perse.Add(0)
                        Punti.Add(0)
                        gFatti.Add(0)
                        gSubiti.Add(0)

                        Rec.MoveNext
                    Loop
                End If
                Rec.Close

                Sql = "SELECT 'Altri' As Tipo, CalendarioPartite.idSqCasa, CalendarioPartite.idSqFuori, CalendarioRisultati.Risultato As Risultato1, '' As RisGiochetti, '' As GoalAvv " &
                    "FROM CalendarioPartite LEFT JOIN CalendarioRisultati On CalendarioPartite.idPartitaGen = CalendarioRisultati.idPartita " &
                    "WHERE CalendarioPartite.idAnno=" & idAnno & " AND CalendarioPartite.idCategoria=" & idCategoria & " And idSqCasa>0 And idSqFuori>0 And idGiornata<=" & idGiornata & " " &
                    "Union All " &
                    "SELECT 'Interna' As Tipo, CalendarioPartite.idSqCasa, CalendarioPartite.idSqFuori, Risultati.Risultato As Risultato1, RisultatiAggiuntivi.RisGiochetti, GoalAvvPrimoTempo + GoalAvvSecondoTempo + GoalAvvTerzoTempo As GoalAvv " &
                    "FROM (((CalendarioPartite " &
                    "LEFT JOIN Partite ON CalendarioPartite.idPartitaGen = Partite.idUnioneCalendario) " &
                    "Left Join Risultati On Partite.idPartita = Risultati.idPartita) " &
                    "Left Join RisultatiAggiuntivi On Partite.idPartita = RisultatiAggiuntivi.idPartita) " &
                    "WHERE CalendarioPartite.idAnno=" & idAnno & " AND CalendarioPartite.idCategoria=" & idCategoria & " And (idSqCasa<0) And idGiornata<=" & idGiornata & " " &
                    "Union All " &
                    "SELECT 'Esterna' As Tipo, CalendarioPartite.idSqCasa, CalendarioPartite.idSqFuori, Risultati.Risultato As Risultato1, RisultatiAggiuntivi.RisGiochetti, GoalAvvPrimoTempo + GoalAvvSecondoTempo + GoalAvvTerzoTempo As GoalAvv " &
                    "FROM (((CalendarioPartite " &
                    "LEFT JOIN Partite ON CalendarioPartite.idPartitaGen = Partite.idUnioneCalendario) " &
                    "Left Join Risultati On Partite.idPartita = Risultati.idPartita) " &
                    "Left Join RisultatiAggiuntivi On Partite.idPartita = RisultatiAggiuntivi.idPartita) " &
                    "WHERE CalendarioPartite.idAnno=" & idAnno & " AND CalendarioPartite.idCategoria=" & idCategoria & " And (idSqFuori<0) And idGiornata<=" & idGiornata
                Rec = LeggeQuery(Conn, Sql, Connessione)
                If TypeOf (Rec) Is String Then
                    Ritorno = StringaErrore & "" & Rec
                    Return Ritorno
                Else
                    Do Until Rec.Eof
                        If "" & Rec("Risultato1").Value <> "" Then
                            Dim g() As String = Rec("Risultato1").Value.split("-")

                            If Rec("Tipo").Value = "Esterna" Then
                                gR1 = Val(g(1))
                                gR2 = Val(g(0))
                            Else
                                gR1 = Val(g(0))
                                gR2 = Val(g(1))
                            End If

                            CeRisultato = True
                        Else
                            CeRisultato = False
                        End If

                        'If "" & Rec("RisGiochetti").Value <> "" Then
                        '    Dim g() As String = Rec("RisGiochetti").Value.split("-")
                        '    gG1 += Val(g(0))
                        '    gG2 += Val(g(1))
                        '    CeGiochetti = True
                        'Else
                        '    CeGiochetti = False
                        'End If
                        'If "" & Rec("GoalAvv").Value <> "" Then
                        '    gA1 = Val(Rec("GoalAvv").Value)
                        '    CeAvversario = True
                        'Else
                        '    CeAvversario = False
                        'End If

                        If CeRisultato Then
                            Dim Indice1 As Integer = -1
                            Dim Indice2 As Integer = -1
                            Dim AppoIndice As Integer = 0

                            For Each i As Integer In idSquadre
                                If Rec("idSqCasa").Value = i Then
                                    Indice1 = AppoIndice
                                End If
                                If Rec("idSqFuori").Value = i Then
                                    Indice2 = AppoIndice
                                End If
                                AppoIndice += 1
                            Next

                            'If Indice1 = -1 Then Indice1 = 0
                            'If Indice2 = -1 Then Indice2 = 0

                            Giocate(Indice1) += 1
                            Giocate(Indice2) += 1

                            g1 = 0
                            g2 = 0

                            If CeRisultato Then
                                gFatti(Indice1) += gR1
                                gSubiti(Indice1) += gR2

                                gFatti(Indice2) += gR2
                                gSubiti(Indice2) += gR1

                                g1 += gR1
                                g2 += gR2
                            End If

                            'If CeGiochetti Then
                            '    gFatti(Indice1) += gG1
                            '    gSubiti(Indice1) += gG2

                            '    gFatti(Indice2) += gR2
                            '    gSubiti(Indice2) += gR1

                            '    g1 += gG1
                            '    g2 += gG2
                            'End If

                            'If CeAvversario Then
                            'If Rec("idSqCasa").Value < 0 Then
                            '    gFatti(Indice2) += gA1
                            '    gSubiti(Indice1) += gA1
                            '    g2 += gA1
                            'Else
                            '    gFatti(Indice1) += gA1
                            '    gSubiti(Indice2) += gA1
                            '    g1 += gA1
                            'End If
                            'End If

                            If g1 > g2 Then
                                Vinte(Indice1) += 1
                                Perse(Indice2) += 1
                                Punti(Indice1) += 3
                            Else
                                If g1 < g2 Then
                                    Vinte(Indice2) += 1
                                    Perse(Indice1) += 1
                                    Punti(Indice2) += 3
                                Else
                                    Pareggiate(Indice1) += 1
                                    Pareggiate(Indice2) += 1
                                    Punti(Indice1) += 1
                                    Punti(Indice2) += 1
                                End If
                            End If
                        End If

                        Rec.MoveNext
                    Loop

                    For i As Integer = 0 To idSquadre.Count - 1
                        For k As Integer = 0 To idSquadre.Count - 1
                            If Val(Punti(i)) > Val(Punti(k)) Then
                                Dim appo As Integer
                                Dim appo2 As String

                                appo = idSquadre(i)
                                idSquadre(i) = idSquadre(k)
                                idSquadre(k) = appo

                                appo2 = Squadre(i)
                                Squadre(i) = Squadre(k)
                                Squadre(k) = appo2

                                appo = Punti(i)
                                Punti(i) = Punti(k)
                                Punti(k) = appo

                                appo = Giocate(i)
                                Giocate(i) = Giocate(k)
                                Giocate(k) = appo

                                appo = Vinte(i)
                                Vinte(i) = Vinte(k)
                                Vinte(k) = appo

                                appo = Pareggiate(i)
                                Pareggiate(i) = Pareggiate(k)
                                Pareggiate(k) = appo

                                appo = Perse(i)
                                Perse(i) = Perse(k)
                                Perse(k) = appo

                                appo = gFatti(i)
                                gFatti(i) = gFatti(k)
                                gFatti(k) = appo

                                appo = gSubiti(i)
                                gSubiti(i) = gSubiti(k)
                                gSubiti(k) = appo
                            End If
                        Next
                    Next

                    Dim c As Integer = 0

                    For Each i As Integer In idSquadre
                        Ritorno &= idSquadre(c) & ";"
                        Ritorno &= Squadre(c) & ";"
                        Ritorno &= Punti(c) & ";"
                        Ritorno &= Giocate(c) & ";"
                        Ritorno &= Vinte(c) & ";"
                        Ritorno &= Pareggiate(c) & ";"
                        Ritorno &= Perse(c) & ";"
                        Ritorno &= gFatti(c) & ";"
                        Ritorno &= gSubiti(c) & ";"
                        Ritorno &= "§"

                        c += 1
                    Next
                End If
            End If

            Dim Ritorno2 As String = SalvaGiornataUtenteCategoria(idUtente, idAnno, idCategoria, idGiornata)

            Conn.Close()
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function AggiungeSquadraAvversaria(ByVal idAnno As String, idCategoria As String, idAvversario As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
                Dim Sql As String = ""
                Dim ProgressivoSquadra As String = ""

                Try
                    Sql = "SELECT Max(idProgressivo)+1 FROM AvversariCalendario Where idAnno=" & idAnno & " And idCategoria=" & idCategoria
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " Nessun avversario rilevato"
                        Else
                            If Rec(0).Value Is DBNull.Value Then
                                ProgressivoSquadra = "1"
                            Else
                                ProgressivoSquadra = Rec(0).Value.ToString
                            End If
                        End If
                        Rec.Close()
                    End If

                    If ProgressivoSquadra <> "" Then
                        Sql = "Insert Into AvversariCalendario Values (" &
                            " " & idAnno & ", " &
                            " " & idCategoria & ", " &
                            " " & ProgressivoSquadra & ", " &
                            " " & idAvversario & " " &
                            ")"
                        Ritorno = EsegueSql(Conn, Sql, Connessione)
                    Else
                        Ritorno = StringaErrore & " Problemi nel rilevamento del progressivo squadra"
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function EliminaSquadraAvversaria(ByVal idAnno As String, idCategoria As String, idAvversario As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
                Dim Sql As String = ""

                Try
                    Sql = "Delete From AvversariCalendario Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idAvversario=" & idAvversario
                    Ritorno = EsegueSql(Conn, Sql, Connessione)
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function InserisceNuovaPartita(ByVal idAnno As String, idGiornata As String, idCategoria As String, Data As String,
                                          Ora As String, Casa As String, Fuori As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))
        Dim idNuovaPartita As Integer = -1
        Dim ProgressivoPartita As String = ""
        Dim idUnioneCalendario As String = ""

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
                Dim Sql As String = ""

                Try
                    Sql = "SELECT Max(idPartita)+1 FROM Partite"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " Nessun progressivo partita rilevato"
                        Else
                            If Rec(0).Value Is DBNull.Value Then
                                idNuovaPartita = "1"
                            Else
                                idNuovaPartita = Rec(0).Value.ToString
                            End If
                        End If
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Try
                    Sql = "SELECT Max(idPartita)+1 FROM CalendarioPartite Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " Nessun progressivo rilevato"
                        Else
                            If Rec(0).Value Is DBNull.Value Then
                                ProgressivoPartita = "1"
                            Else
                                ProgressivoPartita = Rec(0).Value.ToString
                            End If
                        End If
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Dim c() As String = Casa.Split(";")
                Dim f() As String = Fuori.Split(";")

                Try
                    Sql = "SELECT Max(idPartitaGen)+1 FROM CalendarioPartite"
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " Nessun progressivo generale rilevato"
                        Else
                            If Rec(0).Value Is DBNull.Value Then
                                idUnioneCalendario = "1"
                            Else
                                idUnioneCalendario = Rec(0).Value.ToString
                            End If
                        End If
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                If Not Ritorno.Contains(StringaErrore & "") Then
                    Try
                        Sql = "Insert Into CalendarioPartite Values (" & idAnno & ", " & idCategoria & ", " & idGiornata & ", " & ProgressivoPartita & ", " & c(0) & ", " & f(0) & ", " & idUnioneCalendario & ", 'N', '')"
                        Ritorno = EsegueSql(Conn, Sql, Connessione)
                    Catch ex As Exception
                        Ritorno = StringaErrore & " " & ex.Message
                    End Try

                    If Not Ritorno.Contains(StringaErrore & "") Then
                        If Mid(Ora, 1, 3) = "24:" Then Ora = "00:" & Mid(Ora, 4, Ora.Length)
                        Try
                            Sql = "Insert Into CalendarioDate Values (" & idAnno & ", " & idCategoria & ", " & idGiornata & ", '" & Data & " " & Ora & "', " & ProgressivoPartita & ")"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Catch ex As Exception
                            Ritorno = StringaErrore & " " & ex.Message
                        End Try

                        If Not Ritorno.Contains(StringaErrore & "") Then
                            'Try
                            '    Sql = "Insert Into CalendarioRisultati Values (" & idNuovaPartita & ", '')"
                            '    Ritorno = EsegueSql(Conn, Sql, Connessione)
                            'Catch ex As Exception
                            '    Ritorno = StringaErrore & " " & ex.Message
                            'End Try

                            If Not Ritorno.Contains(StringaErrore & "") Then
                                If Val(c(0)) = -1 Or Val(f(0)) = -1 Then
                                    Dim idAvversario As Integer
                                    Dim Datella As Date = Data & " " & Ora
                                    Dim dOraConv As Date = Datella.AddMinutes(-45)
                                    Dim OraConv As String = dOraConv.Hour & ":" & dOraConv.Minute
                                    Dim inCasa As String = ""

                                    If Val(c(0)) = -1 Then
                                        idAvversario = f(0)
                                        inCasa = "S"
                                    Else
                                        idAvversario = c(0)
                                        inCasa = "N"
                                    End If

                                    Dim idCampo As Integer

                                    Try
                                        Sql = "SELECT idCampo FROM SquadreAvversarie Where idAvversario=" & idAvversario & " And Eliminato='N'"
                                        Rec = LeggeQuery(Conn, Sql, Connessione)
                                        If TypeOf (Rec) Is String Then
                                            Ritorno = Rec
                                        Else
                                            If Rec.Eof Then
                                                Ritorno = StringaErrore & " Nessun campo rilevato"
                                            Else
                                                idCampo = Rec(0).Value.ToString
                                            End If
                                            Rec.Close()
                                        End If
                                    Catch ex As Exception
                                        Ritorno = StringaErrore & " " & ex.Message
                                    End Try

                                    If Not Ritorno.Contains("ERROR") Then
                                        Dim idAllenatore As Integer

                                        Try
                                            Sql = "SELECT idAllenatore FROM Allenatori Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And Eliminato='N'"
                                            Rec = LeggeQuery(Conn, Sql, Connessione)
                                            If TypeOf (Rec) Is String Then
                                                Ritorno = Rec
                                            Else
                                                If Rec.Eof Then
                                                    Ritorno = StringaErrore & " Nessun allenatore rilevato"
                                                Else
                                                    idAllenatore = Rec(0).Value.ToString
                                                End If
                                                Rec.Close()
                                            End If
                                        Catch ex As Exception
                                            Ritorno = StringaErrore & " " & ex.Message
                                        End Try

                                        If Not Ritorno.Contains("ERROR") Then
                                            Try
                                                Sql = "Insert Into Partite Values (" & idAnno & ", " & idNuovaPartita & ", " & idCategoria & ", " &
                                                    "" & idAvversario & ", " & idAllenatore & ", '" & Data & " " & Ora & "', " &
                                                    "'N', '" & inCasa & "', 1, " & idCampo & ", '" & OraConv & "', " & idUnioneCalendario & ")"
                                                Ritorno = EsegueSql(Conn, Sql, Connessione)
                                            Catch ex As Exception
                                                Ritorno = StringaErrore & " " & ex.Message
                                            End Try
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If Not Ritorno.Contains("ERROR") Then
                    Ritorno = idGiornata & ";" & idUnioneCalendario & ";" & ProgressivoPartita & ";" & Data & ";" & Ora & ";" & Casa & Fuori & idNuovaPartita & ";"
                Else
                    Dim Appoggio As String = Ritorno

                    If ProgressivoPartita <> "" Then
                        Try
                            Sql = "Delete From CalendarioPartite Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata & " And idPartita=" & ProgressivoPartita
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Catch ex As Exception

                        End Try
                    End If

                    If idNuovaPartita <> -1 Then
                        Try
                            Sql = "Delete From CalendarioDate Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata & " And idPartita=" & idNuovaPartita
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Catch ex As Exception

                        End Try
                    End If

                    'If idNuovaPartita <> -1 Then
                    '    Try
                    '        Sql = "Delete From CalendarioRisultati Where idPartita=" & idNuovaPartita
                    '        Ritorno = EsegueSql(Conn, Sql, Connessione)
                    '    Catch ex As Exception

                    '    End Try
                    'End If

                    If idNuovaPartita <> -1 And idUnioneCalendario <> "" Then
                        Try
                            Sql = "Delete From Partite Where idAnno=" & idAnno & " And idPartita=" & idNuovaPartita & " And idCategoria=" & idCategoria & " And idUnioneCalendario=" & idUnioneCalendario
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Catch ex As Exception

                        End Try
                    End If

                    Ritorno = Appoggio
                End If

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function EliminaPartita(ByVal idAnno As String, idGiornata As String, idCategoria As String, idPartita As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
                Dim idUnioneCalendario As Integer = -1
                Dim Sql As String = "Select * From Partite Where idAnno=" & idAnno & " And idUnioneCalendario=" & idPartita

                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = "*"
                        Else
                            idUnioneCalendario = Rec("idUnioneCalendario").Value
                        End If
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                If idUnioneCalendario <> -1 Then
                    Try
                        Sql = "Delete From Partite Where idAnno=" & idAnno & " And idUnioneCalendario=" & idPartita
                        Ritorno = EsegueSql(Conn, Sql, Connessione)
                    Catch ex As Exception
                        Ritorno = StringaErrore & " " & ex.Message
                    End Try
                End If

                If Not Ritorno.Contains(StringaErrore & "") Then
                    Try
                        Sql = "Delete From CalendarioRisultati Where idPartita=" & idPartita
                        Ritorno = EsegueSql(Conn, Sql, Connessione)
                    Catch ex As Exception
                        Ritorno = StringaErrore & " " & ex.Message
                    End Try

                    If Not Ritorno.Contains(StringaErrore & "") Then
                        Dim idPartitaGiornata As Integer = -1
                        Sql = "Select * From CalendarioPartite Where idAnno=" & idAnno & " And idPartitaGen=" & idPartita
                        Try
                            Rec = LeggeQuery(Conn, Sql, Connessione)
                            If TypeOf (Rec) Is String Then
                                Ritorno = Rec
                            Else
                                If Rec.Eof Then
                                    Ritorno = StringaErrore & " nessun idPartita della giornata rilevato"
                                Else
                                    idPartitaGiornata = Rec("idPartita").Value
                                End If
                                Rec.Close()
                            End If
                        Catch ex As Exception
                            Ritorno = StringaErrore & " " & ex.Message
                        End Try

                        If Not Ritorno.Contains(StringaErrore & "") Then
                            Try
                                Sql = "Delete From CalendarioDate Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata & " And idPartita=" & idPartitaGiornata
                                Ritorno = EsegueSql(Conn, Sql, Connessione)
                            Catch ex As Exception
                                Ritorno = StringaErrore & " " & ex.Message
                            End Try

                            If Not Ritorno.Contains(StringaErrore & "") Then
                                Try
                                    Sql = "Delete From CalendarioPartite Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata & " And idPartita=" & idPartitaGiornata
                                    Ritorno = EsegueSql(Conn, Sql, Connessione)
                                Catch ex As Exception
                                    Ritorno = StringaErrore & " " & ex.Message
                                End Try

                                If Not Ritorno.Contains("ERROR") Then
                                    Ritorno = idGiornata & ";" & idPartitaGiornata & ";"
                                End If
                            End If
                        End If
                    End If
                End If

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function ModificaPartitaAltre(ByVal idAnno As String, idGiornata As String, idCategoria As String, Data As String,
                                    Ora As String, Casa As String, Fuori As String, idUnioneCalendario As String,
                                    ProgressivoPartita As String, Risultato As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))
        Dim Giocata As String = ""

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
                Dim Sql As String = ""

                Dim c() As String = Casa.Split(";")
                Dim f() As String = Fuori.Split(";")

                If Risultato <> "" Then
                    Giocata = "S"
                Else
                    Giocata = "N"
                End If

                Try
                    Sql = "Update CalendarioPartite Set " &
                        "idSqCasa=" & c(0) & ", " &
                        "idSqFuori=" & f(0) & ", " &
                        "Giocata='" & Giocata & "', " &
                        "Risultato='" & Risultato & "' " &
                        "Where idPartitaGen=" & idUnioneCalendario
                    Ritorno = EsegueSql(Conn, Sql, Connessione)
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                If Not Ritorno.Contains(StringaErrore & "") Then
                    Try
                        Sql = "Update CalendarioDate Set " &
                            "Datella='" & Data & " " & Ora & "' " &
                            "Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata & " And idPartita=" & ProgressivoPartita
                        Ritorno = EsegueSql(Conn, Sql, Connessione)
                    Catch ex As Exception
                        Ritorno = StringaErrore & " " & ex.Message
                    End Try
                End If

                If Not Ritorno.Contains(StringaErrore & "") Then
                    If Risultato <> "" Then
                        Try
                            Sql = "Delete From CalendarioRisultati Where idPartita=" & idUnioneCalendario
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Catch ex As Exception
                        End Try

                        Try
                            Sql = "Insert Into CalendarioRisultati Values (" & idUnioneCalendario & ", '" & Risultato & "')"
                            Ritorno = EsegueSql(Conn, Sql, Connessione)
                        Catch ex As Exception
                            Ritorno = StringaErrore & " " & ex.Message
                        End Try
                    End If
                End If
            End If

            If Not Ritorno.Contains("ERROR") Then
                Ritorno = idGiornata & ";" & idUnioneCalendario & ";" & ProgressivoPartita & ";" & Data & ";" & Ora & ";" & Casa & Fuori & Giocata & ";" & Risultato & ";"
            End If

            Conn.Close()
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function RitornaIdPartitaDaUnione(idUnioneCalendario As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
                Dim Sql As String = "Select * From Partite Where idUnioneCalendario=" & idUnioneCalendario

                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            Ritorno = StringaErrore & " Nessun id partita rilevato"
                        Else
                            Ritorno = Rec("idPartita").Value
                        End If
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function SalvaGiornataUtenteCategoria(idUtente As String, idAnno As String, idCategoria As String, idGiornata As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
                Dim Sql As String = ""

                Try
                    Sql = "Delete From Giornata Where idUtente=" & idUtente & " And idAnno=" & idAnno & " And idCategoria=" & idCategoria
                    Ritorno = EsegueSql(Conn, Sql, Connessione)
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Try
                    Sql = "Insert Into Giornata Values (" & idUtente & ", " & idAnno & ", " & idCategoria & ", " & idGiornata & ")"
                    Ritorno = EsegueSql(Conn, Sql, Connessione)
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

    <WebMethod()>
    Public Function RitornaGiornataUtenteCategoria(idUtente As String, idAnno As String, idCategoria As String) As String
        Dim Ritorno As String = ""
        Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."))
        Dim idGiornata As String = "-1"

        If Connessione = "" Then
            Ritorno = ErroreConnessioneNonValida
        Else
            Dim Conn As Object = ApreDB(Connessione)

            If TypeOf (Conn) Is String Then
                Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
            Else
                Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
                Dim Sql As String = ""

                Sql = "Select * From Giornata Where idUtente=" & idUtente & " And idAnno=" & idAnno & " And idCategoria=" & idCategoria
                Try
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ritorno = Rec
                    Else
                        If Rec.Eof Then
                            idGiornata = 1
                            Try
                                Sql = "Insert Into Giornata Values (" & idUtente & ", " & idAnno & ", " & idCategoria & ", " & idGiornata & ")"
                                Ritorno = EsegueSql(Conn, Sql, Connessione)
                                Ritorno = idGiornata
                            Catch ex As Exception
                                Ritorno = StringaErrore & " " & ex.Message
                            End Try
                        Else
                            idGiornata = Rec("idGiornata").Value
                            Ritorno = idGiornata
                        End If
                        Rec.Close()
                    End If
                Catch ex As Exception
                    Ritorno = StringaErrore & " " & ex.Message
                End Try

                Conn.Close()
            End If
        End If

        Return Ritorno
    End Function

End Class