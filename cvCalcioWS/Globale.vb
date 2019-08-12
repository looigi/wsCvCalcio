Module Globale
    Public Const ErroreConnessioneNonValida As String = "ERRORE: Stringa di connessione non valida"
    Public Const ErroreConnessioneDBNonValida As String = "ERRORE: Connessione al db non valida"
    Public Percorso As String
    Public PercorsoSitoCV As String = "C:\inetpub\wwwroot\CVCalcio\App_Themes\Standard\Images\"
    Public PercorsoSitoURLImmagini As String = "http://looigi.no-ip.biz:12345/CvCalcio/App_Themes/Standard/Images/"
    Public StringaErrore As String = "ERROR: "
    Public RigaPari As Boolean = False

	Public Sub EliminaDatiNuovoAnnoDopoErrore(idAnno As String, Conn As Object, Connessione As String)
		Dim Ritorno As String
		Dim Sql As String

		Sql = "Delete * From Anni Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)

		Sql = "Delete * From UtentiMobile Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)

		Sql = "Delete * From Categorie Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)

		Sql = "Delete * From Allenatori Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)

		Sql = "Delete * From Dirigenti Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)

		Sql = "Delete * From Giocatori Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)
	End Sub

	Public Function LeggeImpostazioniDiBase(Percorso As String) As String
        Dim Connessione As String = ""

        ' Impostazioni di base
        Dim ListaConnessioni As ConnectionStringSettingsCollection = ConfigurationManager.ConnectionStrings

        If ListaConnessioni.Count <> 0 Then
            ' Get the collection elements. 
            For Each Connessioni As ConnectionStringSettings In ListaConnessioni
                Dim Nome As String = Connessioni.Name
                Dim Provider As String = Connessioni.ProviderName
                Dim connectionString As String = Connessioni.ConnectionString

                If Nome = "MDBConnectionString" Then
                    Connessione = "Provider=" & Provider & ";" & connectionString
                    Connessione = Replace(Connessione, "*^*^*", Percorso & "\")
                    Exit For
                End If
            Next
        End If

        Return Connessione
    End Function

    Public Function RitornaMultimediaPerTipologia(idAnno As String, id As String, Tipologia As String) As String
        ' PercorsoSitoCV = "D:\Looigi\VB.Net\Miei\WEB\SSDCastelverdeCalcio\CVCalcio\App_Themes\Standard\Images\"
        Dim Ritorno As String = ""
        Dim Ok As Boolean = True
        Dim Percorso As String = PercorsoSitoCV & Tipologia & "\"
        Dim IndirizzoURL As String = PercorsoSitoURLImmagini & Tipologia & "/"
        Dim Codice As String
        Dim gf As New GestioneFilesDirectory

        Select Case Tipologia
            Case "Partite"
                Codice = id.ToString
            Case Else
                Codice = idAnno.ToString & "_" & id.ToString
        End Select
        Percorso &= Codice
        IndirizzoURL &= Codice & "/"
        gf.CreaDirectoryDaPercorso(Percorso & "\")
        gf.ScansionaDirectorySingola(Percorso)
        Dim Filetti() As String = gf.RitornaFilesRilevati
        Dim qFiletti As String = gf.RitornaQuantiFilesRilevati

        For i As Integer = 1 To qFiletti
            If Not Filetti(i).ToUpper.Contains("\THUMBS\") Then
                Dim Dimensioni As Long = FileLen(Filetti(i))
                Dim DataUltimaModifica As String = gf.TornaDataDiUltimaModifica(Filetti(i))
                Dim NomeUrl As String = IndirizzoURL & Filetti(i).Replace(Percorso & "\", "").Replace("\", "/")
                Dim NomeFile As String = gf.TornaNomeFileDaPath(NomeUrl.Replace("/", "\"))

                Ritorno &= NomeUrl & ";" & NomeFile & ";" & Dimensioni.ToString & ";" & DataUltimaModifica & ";" & Codice & "§"
            End If
        Next

        If Ritorno = "" Then
            Ritorno = StringaErrore & " Nessun file rilevato"
        End If

        Return Ritorno
    End Function

    Public Sub CreaHtmlPartita(Conn As Object, Connessione As String, idAnno As String, idPartita As String)
        Dim Sql As String
        Dim Rec As Object
        Dim Ok As Boolean = True
        Dim Pagina As StringBuilder = New StringBuilder
        Dim gf As New GestioneFilesDirectory
        Dim PathBaseImmagini As String = "http://looigi.no-ip.biz:12345/CVCalcio/App_Themes/Standard/Images"
        Dim PathBaseImmScon As String = "http://looigi.no-ip.biz:12345/CVCalcio/App_Themes/Standard/Images/Sconosciuto.png"

        Dim Filone As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\base_partita.txt")
        gf.CreaDirectoryDaPercorso(HttpContext.Current.Server.MapPath(".") & "\Partite\")
        Dim NomeFileFinale As String = HttpContext.Current.Server.MapPath(".") & "\Partite\" & idAnno & "_" & idPartita & ".html"

        Filone = Filone.Replace("***SFONDO***", PathBaseImmagini & "/bg.jpg")

        Sql = "SELECT TipologiePartite.Descrizione As Tipologia, Partite.DataOra, Partite.Casa, Partite.idAvversario, Categorie.idCategoria, Categorie.Descrizione As Squadra1, " &
            "SquadreAvversarie.Descrizione As Squadra2, CampiAvversari.Descrizione As Campo, CampiAvversari.Indirizzo, " &
            "Risultati.Risultato, Risultati.Note, Allenatori.idAllenatore, Allenatori.Cognome + ' ' + Allenatori.Nome As Allenatore, " &
            "MeteoPartite.Tempo, MeteoPartite.Gradi, MeteoPartite.Umidita, MeteoPartite.Pressione, Allenatori.idAllenatore, " &
            "TempiGoalAvversari.TempiPrimoTempo, TempiGoalAvversari.TempiSecondoTempo, TempiGoalAvversari.TempiTerzoTempo, Risultati.Note, " &
            "RisultatiAggiuntivi.Tempo1Tempo, RisultatiAggiuntivi.Tempo2Tempo, RisultatiAggiuntivi.Tempo3Tempo, RisultatiAggiuntivi.RisGiochetti, CampiEsterni.Descrizione As CampoEsterno, " &
            "Partite.RisultatoATempi " &
            "FROM (((((((((Partite LEFT JOIN Risultati ON Partite.idPartita = Risultati.idPartita) " &
            "LEFT JOIN Categorie ON (Partite.idCategoria = Categorie.idCategoria) And (Partite.idAnno = Categorie.idAnno)) " &
            "LEFT JOIN SquadreAvversarie ON Partite.idAvversario = SquadreAvversarie.idAvversario) " &
            "LEFT JOIN TipologiePartite ON Partite.idTipologia = TipologiePartite.idTipologia) " &
            "LEFT JOIN CampiAvversari ON Partite.idCampo = CampiAvversari.idCampo) " &
            "LEFT JOIN Allenatori On (Partite.idAnno = Allenatori.idAnno) And (Partite.idAllenatore = Allenatori.idAllenatore)) " &
            "LEFT JOIN MeteoPartite ON Partite.idPartita = MeteoPartite.idPartita) " &
            "LEFT JOIN TempiGoalAvversari ON Partite.idPartita = TempiGoalAvversari.idPartita) " &
            "LEFT JOIN CampiEsterni ON Partite.idPartita = CampiEsterni.idPartita) " &
            "LEFT JOIN RisultatiAggiuntivi ON Partite.idPartita = RisultatiAggiuntivi.idPartita " &
            "WHERE Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita
        Rec = LeggeQuery(Conn, Sql, Connessione)
        If TypeOf (Rec) Is String Then
            Ok = False
        Else
            If Not Rec.Eof Then
                Dim Meteo As String = "'" & MetteMaiuscoleDopoOgniSpazio(Rec("Tempo").Value) & "' Gradi: " & Rec("Gradi").Value & " Umidità: " & Rec("Umidita").Value & " Pressione: " & Rec("Pressione").Value
                Dim Casa As String = ""& Rec("Casa").Value

                Filone = Filone.Replace("***PARTITA***", "" & idPartita)
                Filone = Filone.Replace("***TIPOLOGIA***", "" & Rec("Tipologia").Value)
                Filone = Filone.Replace("***DATA ORA***", "" & Rec("DataOra").Value)
                If "" & Rec("Casa").Value = "E" Then
                    Filone = Filone.Replace("***CAMPO***", "Campo esterno: " & Rec("CampoEsterno").Value)
                Else
                    Filone = Filone.Replace("***CAMPO***", "" & Rec("Campo").Value)
                End If
                Filone = Filone.Replace("***INDIRIZZO***", "" & Rec("Indirizzo").Value)
                Filone = Filone.Replace("***METEO***", "" & Meteo)
                Filone = Filone.Replace("***NOTE***", "" & Rec("Note").Value)

                Dim CiSonoGiochetti As Boolean = False
                Dim Giochetti() As String = {}

                If Rec("RisGiochetti").Value <> "" Then
                    If Rec("RisGiochetti").Value.ToString.Contains("-") And Rec("RisGiochetti").Value.ToString.Trim <> "-" Then
                        Giochetti = Rec("RisGiochetti").Value.ToString.Split("-")
                        Filone = Filone.Replace("***TIT RIS GIOCHETTI***", "Risultato giochetti:")
                        Filone = Filone.Replace("***TRATTINO2***", "-")
                        Filone = Filone.Replace("***RIS 1G***", Giochetti(0))
                        Filone = Filone.Replace("***RIS 2G***", Giochetti(1))

                        CiSonoGiochetti = True
                    Else
                        Filone = Filone.Replace("***TIT RIS GIOCHETTI***", "")
                        Filone = Filone.Replace("***TRATTINO2***", "")
                        Filone = Filone.Replace("***RIS 1G***", "")
                        Filone = Filone.Replace("***RIS 2G***", "")
                    End If
                Else
                    Filone = Filone.Replace("***TIT RIS GIOCHETTI***", "")
                    Filone = Filone.Replace("***TRATTINO2***", "")
                    Filone = Filone.Replace("***RIS 1G***", "")
                    Filone = Filone.Replace("***RIS 2G***", "")
                End If

                Dim ImmAll As String = PathBaseImmagini & "/Allenatori/" & idAnno & "_" & Rec("idAllenatore").Value & ".Jpg"
                Filone = Filone.Replace("***IMMAGINE ALL***", ImmAll)
                Filone = Filone.Replace("***ALLENATORE***", Rec("Allenatore").Value)

                Dim Imm1 As String = PathBaseImmagini & "/Categorie/" & idAnno & "_" & Rec("idCategoria").Value & ".Jpg"
                Dim Imm2 As String = PathBaseImmagini & "/Avversari/" & Rec("idAvversario").Value & ".Jpg"

                'If Casa = "S" Then
                Filone = Filone.Replace("***IMMAGINE SQ1***", Imm1)
                    Filone = Filone.Replace("***SQUADRA 1***", Rec("Squadra1").Value)

                    Filone = Filone.Replace("***IMMAGINE SQ2***", Imm2)
                    Filone = Filone.Replace("***SQUADRA 2***", Rec("Squadra2").Value)
                'Else
                '    Filone = Filone.Replace("***IMMAGINE SQ2***", Imm2)
                '    Filone = Filone.Replace("***SQUADRA 2***", Rec("Squadra2").Value)

                '    Filone = Filone.Replace("***IMMAGINE SQ1***", Imm1)
                '    Filone = Filone.Replace("***SQUADRA 1***", Rec("Squadra1").Value)
                'End If

                Dim GoalAvv1Tempi As String = Rec("TempiPrimoTempo").Value
                Dim GoalAvv2Tempi As String = Rec("TempiSecondoTempo").Value
                Dim GoalAvv3Tempi As String = Rec("TempiTerzoTempo").Value

                Dim Tempi As String = "Primo tempo: " & Rec("Tempo1Tempo").Value & " Secondo tempo: " & Rec("Tempo2Tempo").Value & " Tezro Tempo: " & Rec("Tempo3Tempo").Value
                Filone = Filone.Replace("***TEMPI DI GIOCO***", Tempi)

                Dim RisultatoATempi As String = "" & Rec("RisultatoATempi").Value.ToString.Trim

                Rec.Close

                ' Arbitro
                Sql = "Select Arbitri.idArbitro, Arbitri.Cognome, Arbitri.Nome " &
                    "FROM(Partite INNER JOIN ArbitriPartite On Partite.idPartita = ArbitriPartite.idPartita) " &
                    "INNER Join Arbitri ON ArbitriPartite.idArbitro = Arbitri.idArbitro " &
                    "Where Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita
                Rec = LeggeQuery(Conn, Sql, Connessione)
                If TypeOf (Rec) Is String Then
                    Ok = False
                Else
                    If Not Rec.Eof Then
                        Dim PathArb As String = PathBaseImmagini & "/Arbitri/" & idAnno & "_" & Rec("idArbitro").Value & ".jpg"
                        Filone = Filone.Replace("***IMMAGINE ARB***", PathArb)
                        Filone = Filone.Replace("***ARBITRO***", Rec("Cognome").Value & " " & Rec("Nome").Value)
                    Else
                        Filone = Filone.Replace("***IMMAGINE ARB***", PathBaseImmScon)
                        Filone = Filone.Replace("***ARBITRO***", "Non impostato")
                    End If

                    ' Dirigenti
                    Sql = "SELECT Dirigenti.idDirigente, Dirigenti.Cognome, Dirigenti.Nome " &
                    "FROM (Partite INNER JOIN DirigentiPartite ON Partite.idPartita = DirigentiPartite.idPartita) INNER JOIN Dirigenti ON DirigentiPartite.idDirigente = Dirigenti.idDirigente " &
                    "Where Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita
                    Rec = LeggeQuery(Conn, Sql, Connessione)
                    If TypeOf (Rec) Is String Then
                        Ok = False
                    Else
                        Dim Dirigenti As New StringBuilder

                        Dirigenti.Append("<table style=""width: 99%; text-align: center;"">")

                        Do Until Rec.Eof
                            Dim Path As String = PathBaseImmagini & "/Dirigenti/" & idAnno & "_" & Rec("idDirigente").Value & ".jpg"

                            Dirigenti.Append("<tr>")
                            Dirigenti.Append("<td>")
                            Dirigenti.Append("<img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & PathBaseImmScon & "'"" />")
                            Dirigenti.Append("</td>")
                            Dirigenti.Append("<td>")
                            Dirigenti.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Rec("Cognome").Value & " " & Rec("Nome").Value & "</span>")
                            Dirigenti.Append("</td>")
                            Dirigenti.Append("</tr>")

                            Rec.MoveNext
                        Loop

                        Dirigenti.Append("</table>")

                        Filone = Filone.Replace("***DIRIGENTE***", Dirigenti.ToString)

                        Rec.Close

                        ' Convocati
                        Sql = "SELECT Convocati.idGiocatore, Giocatori.NumeroMaglia, Giocatori.Cognome, Giocatori.Nome, Ruoli.Descrizione As Ruolo " &
                            "FROM (Partite INNER JOIN Convocati ON Partite.idPartita = Convocati.idPartita) " &
                            "INNER JOIN (Giocatori INNER JOIN Ruoli ON Giocatori.idRuolo = Ruoli.idRuolo) ON (Convocati.idGiocatore = Giocatori.idGiocatore) AND (Partite.idAnno = Giocatori.idAnno) " &
                            "Where Partite.idAnno=" & idAnno & " And PArtite.idPartita=" & idPartita & " " &
                            "Order By Ruoli.idRuolo"
                        Rec = LeggeQuery(Conn, Sql, Connessione)
                        If TypeOf (Rec) Is String Then
                            Ok = False
                        Else
                            Dim Convocati As New StringBuilder

                            Convocati.Append("<table style=""width: 99%; text-align: center;"">")

                            Do Until Rec.Eof
                                Dim C As String = Rec("Cognome").Value & " " & Rec("Nome").Value & " (" & Rec("Ruolo").Value & ")"
                                Dim Path As String = PathBaseImmagini & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".jpg"

                                Convocati.Append("<tr>")
                                Convocati.Append("<td>")
                                Convocati.Append("<img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & PathBaseImmScon & "'"" />")
                                Convocati.Append("</td>")
                                Convocati.Append("<td style=""text-align: center;"">")
                                Convocati.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Rec("NumeroMaglia").Value & "</span>")
                                Convocati.Append("</td>")
                                Convocati.Append("<td style=""text-align: left;"">")
                                Convocati.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & C & "</span>")
                                Convocati.Append("</td>")
                                Convocati.Append("</tr>")

                                Rec.MoveNext
                            Loop
                            Rec.Close

                            Convocati.Append("</table>")

                            Filone = Filone.Replace("***CONVOCATI***", Convocati.ToString)

                            ' Marcatori
                            Sql = "SELECT RisultatiAggiuntiviMarcatori.Minuto, Giocatori.NumeroMaglia, Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Ruoli.Descrizione As Ruolo, RisultatiAggiuntiviMarcatori.idTempo " &
                                    "FROM ((Partite INNER JOIN RisultatiAggiuntiviMarcatori ON Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
                                    "INNER JOIN Giocatori ON (Partite.idAnno = Giocatori.idAnno) AND (RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore)) " &
                                    "INNER JOIN Ruoli ON Giocatori.idRuolo = Ruoli.idRuolo " &
                                    "Where Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita & " " &
                                    "Order By idTempo, Minuto " &
                                    "Union ALL " &
                                    "SELECT RisultatiAggiuntiviMarcatori.Minuto, '', -1, 'Autorete', '', '' As Ruolo, RisultatiAggiuntiviMarcatori.idTempo FROM Partite INNER JOIN RisultatiAggiuntiviMarcatori ON Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita " &
                                    "Where Partite.idAnno = " & idAnno & " And Partite.idPartita = " & idPartita & " And IdGiocatore = -1 Order By idTempo, Minuto"
                            Rec = LeggeQuery(Conn, Sql, Connessione)
                            If TypeOf (Rec) Is String Then
                                Ok = False
                            Else
                                Dim Marc() As String = {}
                                Dim QuantiGoal As Integer = 0
                                Dim QuantiGoal1 As Integer = 0
                                Dim QuantiGoal2 As Integer = 0

                                Do Until Rec.Eof
                                    ReDim Preserve Marc(QuantiGoal)
                                    Marc(QuantiGoal) = "0" & Rec("idTempo").Value & ";" & Format(Rec("Minuto").Value, "00") & ";" & Rec("idGiocatore").Value & ";" & Rec("Cognome").Value & ";" & Rec("Nome").Value & ";" & Rec("Ruolo").Value & ";"
                                    QuantiGoal1 += 1
                                    QuantiGoal += 1

                                    Rec.MoveNext
                                Loop
                                Rec.Close

                                Dim ga1() As String = GoalAvv1Tempi.Split("#")

                                For Each g As String In ga1
                                    If g <> "" Then
                                        ReDim Preserve Marc(QuantiGoal)
                                        Marc(QuantiGoal) = "01;" & Format(Val(g), "00") & ";;Goal avversario;;;"
                                        QuantiGoal2 += 1
                                        QuantiGoal += 1
                                    End If
                                Next

                                Dim ga2() As String = GoalAvv2Tempi.Split("#")

                                For Each g As String In ga2
                                    If g <> "" Then
                                        ReDim Preserve Marc(QuantiGoal)
                                        Marc(QuantiGoal) = "02;" & Format(Val(g), "00") & ";;Goal avversario;;;"
                                        QuantiGoal2 += 1
                                        QuantiGoal += 1
                                    End If
                                Next

                                Dim ga3() As String = GoalAvv3Tempi.Split("#")

                                For Each g As String In ga3
                                    If g <> "" Then
                                        ReDim Preserve Marc(QuantiGoal)
                                        Marc(QuantiGoal) = "03;" & Format(Val(g), "00") & ";;Goal avversario;;;"
                                        QuantiGoal2 += 1
                                        QuantiGoal += 1
                                    End If
                                Next

                                For i As Integer = 0 To Marc.Length - 1
                                    For k As Integer = 0 To Marc.Length - 1
                                        If i <> k Then
                                            If Marc(i) < Marc(k) Then
                                                Dim Appo As String = Marc(i)
                                                Marc(i) = Marc(k)
                                                Marc(k) = Appo
                                            End If
                                        End If
                                    Next
                                Next

                                ' Risultato a tempi
                                Dim GoalPropri As Integer = 0
                                Dim GoalAvversari As Integer = 0
                                Dim NomiCampi() As String = {"", "GoalAvvPrimoTempo", "GoalAvvSecondoTempo", "GoalAvvTerzoTempo"}
                                Dim RisProprio As Integer = 0
                                Dim RisAvversario As Integer = 0

                                For i As Integer = 1 To 3
                                    Sql = "Select Count(*) From RisultatiAggiuntiviMarcatori Where idPartita=" & idPartita & " And idTempo=" & i
                                    Rec = LeggeQuery(Conn, Sql, Connessione)
                                    If Rec(0).Value Is DBNull.Value Then
                                        GoalPropri = 0
                                    Else
                                        GoalPropri = Rec(0).Value
                                    End If
                                    Rec.Close
                                    Sql = "Select Sum(" & NomiCampi(i) & ") From RisultatiAggiuntivi Where idPartita=" & idPartita & " And " & NomiCampi(i) & "<>-1"
                                    Rec = LeggeQuery(Conn, Sql, Connessione)
                                    If Rec(0).Value Is DBNull.Value Then
                                        GoalAvversari = 0
                                    Else
                                        GoalAvversari = Rec(0).Value
                                    End If
                                    Rec.Close

                                    If GoalPropri > GoalAvversari Then
                                        RisProprio += 1
                                    Else
                                        If GoalPropri < GoalAvversari Then
                                            RisAvversario += 1
                                        Else
                                            RisProprio += 1
                                            RisAvversario += 1
                                        End If
                                    End If
                                Next

                                If CiSonoGiochetti Then
                                    If Val(Giochetti(0)) > Val(Giochetti(1)) Then
                                        RisProprio += 1
                                    Else
                                        If Val(Giochetti(0)) < Val(Giochetti(1)) Then
                                            RisAvversario += 1
                                        Else
                                            RisProprio += 1
                                            RisAvversario += 1
                                        End If
                                    End If
                                End If

                                If RisultatoATempi = "S" Then
                                    Filone = Filone.Replace("***TIT RIS TEMPI***", "Risultato a tempi:")
                                    Filone = Filone.Replace("***RIS 1T***", RisProprio.ToString.Trim)
                                    Filone = Filone.Replace("***RIS 2T***", RisAvversario.ToString.Trim)
                                    Filone = Filone.Replace("***TRATTINO1***", "-")
                                Else
                                    Filone = Filone.Replace("***RIS 1T***", "")
                                    Filone = Filone.Replace("***RIS 2T***", "")
                                    Filone = Filone.Replace("***TIT RIS TEMPI***", "")
                                    Filone = Filone.Replace("***TRATTINO1***", "")
                                End If

                                Dim Marcatori As New StringBuilder

                                    Marcatori.Append("<table style=""width: 99%; text-align: center;"">")
                                    Marcatori.Append("<tr>")
                                    Marcatori.Append("<td>")
                                    Marcatori.Append("<span class=""testo verde"" style=""font-size: 13px;"">Tempo</span>")
                                    Marcatori.Append("</td>")
                                    Marcatori.Append("<td>")
                                    Marcatori.Append("<span class=""testo verde"" style=""font-size: 13px;"">Minuto</span>")
                                    Marcatori.Append("</td>")
                                    Marcatori.Append("<td>")
                                    Marcatori.Append("</td>")
                                    Marcatori.Append("<td>")
                                    Marcatori.Append("<span class=""testo verde"" style=""font-size: 13px;"">Giocatore</span>")
                                    Marcatori.Append("</td>")
                                    Marcatori.Append("<td>")
                                    Marcatori.Append("<span class=""testo verde"" style=""font-size: 13px;"">Ruolo</span>")
                                    Marcatori.Append("</td>")
                                    Marcatori.Append("</tr>")

                                    Dim OldTempo As String = ""

                                    For Each m As String In Marc
                                        Dim Mm() As String = m.Split(";")

                                        If OldTempo <> Mm(0) Then
                                            Marcatori.Append("<tr>")
                                            Marcatori.Append("<td>")
                                            Marcatori.Append("<hr />")
                                            Marcatori.Append("</td>")
                                            Marcatori.Append("<td>")
                                            Marcatori.Append("<hr />")
                                            Marcatori.Append("</td>")
                                            Marcatori.Append("<td>")
                                            Marcatori.Append("<hr />")
                                            Marcatori.Append("</td>")
                                            Marcatori.Append("<td>")
                                            Marcatori.Append("<hr />")
                                            Marcatori.Append("</td>")
                                            Marcatori.Append("<td>")
                                            Marcatori.Append("<hr />")
                                            Marcatori.Append("</td>")
                                            Marcatori.Append("</tr>")
                                            OldTempo = Mm(0)
                                        End If

                                        Dim Path As String
                                    If m.Contains("Goal avversario") Then
                                        Path = PathBaseImmagini & "/goal.png"
                                    Else
                                        If m.Contains("Autorete") Then
                                            Path = PathBaseImmagini & "/autorete.png"
                                        Else
                                            Path = PathBaseImmagini & "/Giocatori/" & idAnno & "_" & Mm(2) & ".jpg"
                                        End If
                                    End If

                                    Marcatori.Append("<tr>")
                                        Marcatori.Append("<td>")
                                        Marcatori.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Mm(0) & "</span>")
                                        Marcatori.Append("</td>")
                                        Marcatori.Append("<td>")
                                        Marcatori.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Mm(1) & "°</span>")
                                        Marcatori.Append("</td>")
                                        Marcatori.Append("<td>")
                                        Marcatori.Append("<img src=""" & Path & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & PathBaseImmScon & "'"" />")
                                        Marcatori.Append("</td>")
                                        Marcatori.Append("<td style=""text-align: left;"">")
                                        Marcatori.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Mm(3) & " " & Mm(4) & "</span>")
                                        Marcatori.Append("</td>")
                                        Marcatori.Append("<td>")
                                        Marcatori.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Mm(5) & "</span>")
                                        Marcatori.Append("</td>")
                                        Marcatori.Append("</tr>")
                                    Next
                                    Marcatori.Append("</table>")

                                    Filone = Filone.Replace("***MARCATORI***", Marcatori.ToString)

                                    ' Risultato
                                    'If Casa = "S" Then
                                    Filone = Filone.Replace("***RIS 1***", QuantiGoal1)
                                    Filone = Filone.Replace("***RIS 2***", QuantiGoal2)
                                    'Else
                                    '    Filone = Filone.Replace("***RIS 1***", QuantiGoal2)
                                    '    Filone = Filone.Replace("***RIS 2***", QuantiGoal1)
                                    'End If

                                    If QuantiGoal1 > QuantiGoal2 Then
                                        Filone = Filone.Replace("***COLORE RIS***", "verde")
                                    Else
                                        If QuantiGoal1 < QuantiGoal2 Then
                                            Filone = Filone.Replace("***COLORE RIS***", "rosso")
                                        Else
                                            Filone = Filone.Replace("***COLORE RIS***", "nero")
                                        End If
                                    End If

                                    gf.CreaAggiornaFile(NomeFileFinale, Filone)
                                End If
                            End If
                    End If
                End If
            Else
                Ok = False
            End If
        End If
    End Sub
End Module
