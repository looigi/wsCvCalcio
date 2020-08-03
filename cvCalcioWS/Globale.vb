﻿Module Globale
	Public Const ErroreConnessioneNonValida As String = "ERRORE: Stringa di connessione non valida"
	Public Const ErroreConnessioneDBNonValida As String = "ERRORE: Connessione al db non valida"
	Public Percorso As String
	' Public PercorsoSitoCV As String = "C:\GestioneCampionato\CalcioImages\" ' "C:\inetpub\wwwroot\CVCalcio\App_Themes\Standard\Images\"
	' Public PercorsoSitoURLImmagini As String = "http://loppa.duckdns.org:90/MultiMedia/" ' "http://looigi.no-ip.biz:90/CvCalcio/App_Themes/Standard/Images/"
	Public StringaErrore As String = "ERROR: "
	Public RigaPari As Boolean = False

	Public Function SistemaNumero(Numero As String) As String
		If Numero = "" Then
			Return "Null"
		Else
			Return Numero.Replace(",", ".")
		End If
	End Function

	Public Function AggiungeRigoriEGoal(NomeLista As List(Of String), Rec As Object) As List(Of String)
		Dim ListaNomi As New List(Of String)
		Dim posi As Integer = 0
		Dim Ok As Boolean = False

		ListaNomi.AddRange(NomeLista)
		For Each s As String In ListaNomi
			Dim n As Integer = Val(Mid(s, 1, s.IndexOf("-")))
			If n = Rec(4).Value Then
				Dim cc() As String = s.Split("-")
				ListaNomi.Item(posi) = Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & (Rec(3).Value + cc(2))
				Ok = True
				Exit For
			End If
			posi += 1
		Next

		If Not Ok Then
			ListaNomi.Add(Rec(4).Value & "-" & Rec(1).Value & " " & Rec(2).Value & "-" & Rec(3).Value)
		End If

		Return ListaNomi
	End Function

	Public Sub EliminaDatiNuovoAnnoDopoErrore(idAnno As String, Conn As Object, Connessione As String)
		Dim Ritorno As String
		Dim Sql As String

		Sql = "delete from Anni Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)

		Sql = "delete from UtentiMobile Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)

		Sql = "delete from Categorie Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)

		Sql = "delete from Allenatori Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)

		Sql = "delete from Dirigenti Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)

		Sql = "delete from Giocatori Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)

		Sql = "delete from Arbitri Where idAnno=" & idAnno
		Ritorno = EsegueSql(Conn, Sql, Connessione)
	End Sub

	Public Function LeggeImpostazioniDiBase(Percorso As String, Squadra As String) As String
		Dim Connessione As String = ""

		' Impostazioni di base
		Dim ListaConnessioni As ConnectionStringSettingsCollection = ConfigurationManager.ConnectionStrings

		If ListaConnessioni.Count <> 0 Then
			' Get the collection elements. 
			For Each Connessioni As ConnectionStringSettings In ListaConnessioni
				Dim Nome As String = Connessioni.Name
				Dim Provider As String = Connessioni.ProviderName
				Dim connectionString As String = Connessioni.ConnectionString

				If Nome = "SQLConnectionStringLOCALE" Then
					Connessione = "Provider=" & Provider & ";" & connectionString
					Connessione = Replace(Connessione, "*^*^*", Percorso & "\")
					If Squadra <> "" Then
						If Squadra = "DBVUOTO" Then
							Connessione = Connessione.Replace("***NOME_DB***", "DBVuoto")
						Else
							Connessione = Connessione.Replace("***NOME_DB***", Squadra)
						End If
					Else
						Connessione = Connessione.Replace("***NOME_DB***", "Generale")
					End If
					Exit For
				End If
			Next
		End If

		Return Connessione
	End Function

	Public Function RitornaMultimediaPerTipologia(Squadra As String, idAnno As String, id As String, Tipologia As String) As String
		' PercorsoSitoCV = "D:\Looigi\VB.Net\Miei\WEB\SSDCastelverdeCalcio\CVCalcio\App_Themes\Standard\Images\"
		Dim gf As New GestioneFilesDirectory
		Dim Righe As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\PathAllegati.txt")
		Dim Campi() As String = Righe.Split(";")

		Dim Ritorno As String = ""
		Dim Ok As Boolean = True
		Dim Percorso As String = Campi(0) & "\" & Squadra & "\" & Tipologia & "\"
		Percorso = Percorso.Replace("\\", "\")
		Dim IndirizzoURL As String = Campi(2) & "/" & Squadra & "/" & Tipologia & "/"
		IndirizzoURL = IndirizzoURL.Replace("//", "/")
		Dim Codice As String

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

	Public Function CreaHtmlPartita(Squadra As String, Conn As Object, Connessione As String, idAnno As String, idPartita As String) As String
		Dim Sql As String
		Dim Rec As Object
		Dim Rec2 As Object
		Dim Ok As Boolean = True
		Dim Pagina As StringBuilder = New StringBuilder
		Dim gf As New GestioneFilesDirectory
		Dim PathBaseImmagini As String = "http://loppa.duckdns.org:90/MultiMedia" ' "http://looigi.no-ip.biz:90/CVCalcio/App_Themes/Standard/Images"
		Dim PathBaseImmScon As String = "http://loppa.duckdns.org:90/MultiMedia/Sconosciuto.png" ' "http://looigi.no-ip.biz:90/CVCalcio/App_Themes/Standard/Images/Sconosciuto.png"
		Dim Ritorno As String = "*"

		Dim Filone As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\base_partita.txt")
		gf.CreaDirectoryDaPercorso(HttpContext.Current.Server.MapPath(".") & "\Partite\" & Squadra & "\")
		Dim NomeFileFinale As String = HttpContext.Current.Server.MapPath(".") & "\Partite\" & Squadra & "\" & idAnno & "_" & idPartita & ".html"

		' Return NomeFileFinale

		Filone = Filone.Replace("***SFONDO***", PathBaseImmagini & "/bg.jpg")

		Sql = "SELECT TipologiePartite.Descrizione As Tipologia, Partite.DataOra, Partite.Casa, Partite.idAvversario, Categorie.idCategoria, Categorie.Descrizione As Squadra1, " &
			"SquadreAvversarie.Descrizione As Squadra2, CampiAvversari.Descrizione As CampoAvversari, CampiAvversari.Indirizzo As IndirizzoAvversari, " &
			"Risultati.Risultato, Risultati.Note, Allenatori.idAllenatore, Allenatori.Cognome + ' ' + Allenatori.Nome As Allenatore, " &
			"MeteoPartite.Tempo, MeteoPartite.Gradi, MeteoPartite.Umidita, MeteoPartite.Pressione, Allenatori.idAllenatore, " &
			"TempiGoalAvversari.TempiPrimoTempo, TempiGoalAvversari.TempiSecondoTempo, TempiGoalAvversari.TempiTerzoTempo, Risultati.Note, " &
			"RisultatiAggiuntivi.Tempo1Tempo, RisultatiAggiuntivi.Tempo2Tempo, RisultatiAggiuntivi.Tempo3Tempo, RisultatiAggiuntivi.RisGiochetti, CampiEsterni.Descrizione As CampoEsterno, " &
			"Partite.RisultatoATempi, Anni.CampoSquadra, Anni.Indirizzo As IndirizzoBase " &
			"FROM ((((((((((Partite LEFT JOIN Risultati ON Partite.idPartita = Risultati.idPartita) " &
			"LEFT JOIN Categorie ON (Partite.idCategoria = Categorie.idCategoria) And (Partite.idAnno = Categorie.idAnno)) " &
			"LEFT JOIN SquadreAvversarie ON Partite.idAvversario = SquadreAvversarie.idAvversario) " &
			"LEFT JOIN TipologiePartite ON Partite.idTipologia = TipologiePartite.idTipologia) " &
			"LEFT JOIN CampiAvversari ON Partite.idCampo = CampiAvversari.idCampo) " &
			"LEFT JOIN Allenatori On (Partite.idAnno = Allenatori.idAnno) And (Partite.idAllenatore = Allenatori.idAllenatore)) " &
			"LEFT JOIN MeteoPartite ON Partite.idPartita = MeteoPartite.idPartita) " &
			"LEFT JOIN TempiGoalAvversari ON Partite.idPartita = TempiGoalAvversari.idPartita) " &
			"LEFT JOIN CampiEsterni ON Partite.idPartita = CampiEsterni.idPartita) " &
			"LEFT JOIN RisultatiAggiuntivi ON Partite.idPartita = RisultatiAggiuntivi.idPartita) " &
			"LEFT JOIN Anni ON Partite.idanno = Anni.idAnno " &
			"WHERE Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita
		Rec = LeggeQuery(Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ok = False
			Ritorno = "Problemi lettura generale"
		Else
			If Not Rec.Eof Then
				Dim Meteo As String = "'" & MetteMaiuscoleDopoOgniSpazio(Rec("Tempo").Value) & "' Gradi: " & Rec("Gradi").Value & " Umidità: " & Rec("Umidita").Value & " Pressione: " & Rec("Pressione").Value
				Dim Casa As String = "" & Rec("Casa").Value

				Filone = Filone.Replace("***PARTITA***", "" & idPartita)
				Filone = Filone.Replace("***TIPOLOGIA***", "" & Rec("Tipologia").Value)
				Filone = Filone.Replace("***DATA ORA***", "" & Rec("DataOra").Value)
				If "" & Rec("Casa").Value = "E" Then
					Filone = Filone.Replace("***CAMPO***", "Campo esterno: " & Rec("CampoEsterno").Value)
					Filone = Filone.Replace("***INDIRIZZO***", "")
				Else
					If (Rec("Casa").Value = "N") Then
						Filone = Filone.Replace("***CAMPO***", "" & Rec("CampoAvversari").Value)
						Filone = Filone.Replace("***INDIRIZZO***", "" & Rec("IndirizzoAvversari").Value)
					Else
						Filone = Filone.Replace("***CAMPO***", Rec("CampoSquadra").Value)
						Filone = Filone.Replace("***INDIRIZZO***", "" & Rec("IndirizzoBase").Value)
					End If
				End If
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

				Dim ImmAll As String = PathBaseImmagini & "/" & Squadra & "/Allenatori/" & idAnno & "_" & Rec("idAllenatore").Value & ".Jpg"
				Filone = Filone.Replace("***IMMAGINE ALL***", ImmAll)
				Filone = Filone.Replace("***ALLENATORE***", Rec("Allenatore").Value)

				Dim Imm1 As String = PathBaseImmagini & "/" & Squadra & "/Categorie/" & idAnno & "_" & Rec("idCategoria").Value & ".Jpg"
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
					Ritorno = "Problemi lettura arbitro"
				Else
					If Not Rec.Eof Then
						Dim PathArb As String = PathBaseImmagini & "/Arbitri/" & Rec("idArbitro").Value & ".jpg"
						Filone = Filone.Replace("***IMMAGINE ARB***", PathArb)
						Filone = Filone.Replace("***ARBITRO***", Rec("Cognome").Value & " " & Rec("Nome").Value)
					Else
						Filone = Filone.Replace("***IMMAGINE ARB***", PathBaseImmScon)
						Filone = Filone.Replace("***ARBITRO***", "Non impostato")
					End If

					' Dirigenti
					Sql = "SELECT Dirigenti.idDirigente, Dirigenti.Cognome, Dirigenti.Nome " &
						"FROM (Partite INNER JOIN DirigentiPartite ON Partite.idPartita = DirigentiPartite.idPartita) INNER JOIN Dirigenti ON DirigentiPartite.idDirigente = Dirigenti.idDirigente " &
						"Where Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita & " And Dirigenti.idAnno=" & idAnno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ok = False
						Ritorno = "Problemi lettura dirigenti"
					Else
						Dim Dirigenti As New StringBuilder

						Dirigenti.Append("<table style=""width: 99%; text-align: center;"">")

						Do Until Rec.Eof
							Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Dirigenti/" & idAnno & "_" & Rec("idDirigente").Value & ".jpg"

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
							Ritorno = "Problemi lettura convocati"
						Else
							Dim Convocati As New StringBuilder

							Convocati.Append("<table style=""width: 99%; text-align: center;"">")

							Do Until Rec.Eof
								Dim C As String = Rec("Cognome").Value & " " & Rec("Nome").Value & " (" & Rec("Ruolo").Value & ")"
								Dim Path As String = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".jpg"

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
							Sql = "Select * From (SELECT RisultatiAggiuntiviMarcatori.Minuto, Giocatori.NumeroMaglia, Giocatori.idGiocatore, Giocatori.Cognome, Giocatori.Nome, Ruoli.Descrizione As Ruolo, RisultatiAggiuntiviMarcatori.idTempo " &
									"FROM ((Partite INNER JOIN RisultatiAggiuntiviMarcatori ON Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita) " &
									"INNER JOIN Giocatori ON (Partite.idAnno = Giocatori.idAnno) AND (RisultatiAggiuntiviMarcatori.idGiocatore = Giocatori.idGiocatore)) " &
									"INNER JOIN Ruoli ON Giocatori.idRuolo = Ruoli.idRuolo " &
									"Where Partite.idAnno=" & idAnno & " And Partite.idPartita=" & idPartita & " " &
									"Union ALL " &
									"SELECT RisultatiAggiuntiviMarcatori.Minuto, '', -1, 'Autorete', '', '' As Ruolo, RisultatiAggiuntiviMarcatori.idTempo FROM Partite INNER JOIN RisultatiAggiuntiviMarcatori ON Partite.idPartita = RisultatiAggiuntiviMarcatori.idPartita " &
									"Where Partite.idAnno = " & idAnno & " And Partite.idPartita = " & idPartita & " And IdGiocatore = -1) A " &
									"Order By idTempo, Minuto"
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ok = False
								Ritorno = "Problemi lettura marcatori: " & Sql
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

								Dim CiSonoRigori As Boolean = False
								Dim Rigoristi As New List(Of String)
								Dim RigoriAvv As String = ""
								Dim RigoriSegnatiPropri As Integer = 0
								Dim RigoriSegnatiAvversari As Integer = 0
								Dim RigoriSbagliatiAvversari As Integer = 0

								Sql = "SELECT RigoriPropri.idGiocatore, RigoriPropri.idRigore, Ruoli.Descrizione, Giocatori.Cognome + ' ' + Giocatori.Nome As Giocatore, " &
									"Giocatori.NumeroMaglia, RigoriPropri.Termine From ((RigoriPropri " &
									"Left Join Giocatori On RigoriPropri.idGiocatore=Giocatori.idGiocatore And RigoriPropri.idAnno = Giocatori.idAnno) " &
									"Left Join Ruoli On Giocatori.idRuolo = Ruoli.idRuolo) " &
									"Where RigoriPropri.idAnno=" & idAnno & " And idPartita=" & idPartita & " " &
									"Order By RigoriPropri.idRigore"
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
									Ok = False
									Ritorno = "Problemi lettura rigori"
								Else
									If Not Rec2.Eof Then
										CiSonoRigori = True

										Do Until Rec2.Eof
											Dim Termine As String = ""
											Dim Colore As String = ""

											If Rec2("Termine").Value = "1" Then
												Termine = "Segnato"
												Colore = "verde"
												RigoriSegnatiPropri += 1
											Else
												If Rec2("Termine").Value = "0" Then
													Termine = "Sbagliato"
													Colore = "rosso"
												End If
											End If
											If Termine <> "" Then
												' Rigoristi.Add("<span class=""testo " & Colore & """ style=""font-size: 15px;"">Rigore " & Rec2("idRigore").Value & ": " & Rec2("Giocatore").Value & " (" & Rec2("Descrizione").Value & ") - " & Termine & "</span>")
												Rigoristi.Add(Colore & ";" & Rec2("idRigore").Value & ";;" & Rec2("Giocatore").Value & ";" & Rec2("Descrizione").Value & "; " & Termine & ";" & Rec2("idGiocatore").Value & ";")
											End If

											Rec2.MoveNext
										Loop
										Rec2.Close
									End If
								End If
								If CiSonoRigori Then
									Sql = "Select * From RigoriAvversari Where idAnno=" & idAnno & " And idPartita=" & idPartita
									Rec2 = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec2) Is String Then
									Else
										If Not Rec2.Eof Then
											RigoriSegnatiAvversari += Val(Rec2("Segnati").Value)
											RigoriSbagliatiAvversari += Val(Rec2("Sbagliati").Value)

											RigoriAvv = Rec2("Segnati").Value & "!" & Rec2("Sbagliati").Value & "!"
										End If
									End If

									Dim Rigori As String = "<span class=""testo blu"" style=""font-size: 15px;"">RISULTATO DOPO I TEMPI REGOLAMENTARI: " & QuantiGoal1 & " - " & QuantiGoal2 & "</span><br /><br />"

									Rigori &= "<span class=""testo blu"" style=""font-size: 15px;"">RIGORI PROPRI:</span><hr />"

									Rigori &= "<table style=""width: 99%; text-align: center;"">"
									For Each s As String In Rigoristi
										Dim c() As String = s.Split(";")
										Dim Path2 As String = PathBaseImmagini & "/Giocatori/" & idAnno & "_" & c(6) & ".jpg"

										Rigori &= "<tr>"
										Rigori &= "<td align=""left"">"
										Rigori &= "<span class=""testo blu"" style=""font-size: 15px;"">Rigore " & c(1) & "</span>"
										Rigori &= "</td>"
										Rigori &= "<td>"
										Rigori &= "<img src=""" & Path2 & """ style=""width: 50px; height: 50px;"" onerror=""this.src='" & PathBaseImmScon & "'"" />"
										Rigori &= "</td>"
										Rigori &= "<td align=""center"">"
										Rigori &= "<span class=""testo blu"" style=""font-size: 15px;"">" & c(3) & "</span>"
										Rigori &= "</td>"
										Rigori &= "<td align=""center"">"
										Rigori &= "<span class=""testo blu"" style=""font-size: 15px;"">" & c(4) & "</span>"
										Rigori &= "</td>"
										Rigori &= "<td align=""center"">"
										Rigori &= "<span class=""testo " & c(0) & """ style=""font-size: 15px;"">" & c(5) & "</span>"
										Rigori &= "</td>"
										Rigori &= "</tr>"
									Next
									Rigori &= "</table>"

									Rigori &= "<br /><span class=""testo blu"" style=""font-size: 15px;"">RIGORI AVVERSARI:</span><hr />"
									Rigori &= "<span class=""testo rosso"" style=""font-size: 15px;"">Segnati: " & RigoriSegnatiAvversari & "</span><br />"
									Rigori &= "<span class=""testo verde"" style=""font-size: 15px;"">Sbagliati: " & RigoriSbagliatiAvversari & "</span><hr />"

									Filone = Filone.Replace("***RIGORI***", Rigori)

									If RisultatoATempi = "S" Then
										RisProprio += RigoriSegnatiPropri
										RisAvversario += RigoriSegnatiAvversari
									Else
										If (RigoriSegnatiPropri > RigoriSegnatiAvversari) Then
											RisProprio += 1
										Else
											If (RigoriSegnatiPropri < RigoriSegnatiAvversari) Then
												RisAvversario += 1
											End If
										End If
									End If
								Else
									Filone = Filone.Replace("***RIGORI***", "")
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
											Path = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & Mm(2) & ".jpg"
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

								' Eventi
								Dim Eventi As New StringBuilder

								Eventi.Append("<table style=""width: 99%; text-align: center;"">")

								Sql = "SELECT EventiPartita.idTempo, EventiPartita.Minuto, Eventi.Descrizione, iif(Giocatori.Cognome + ' ' + Giocatori.Nome is null, 'Avversario', Giocatori.Cognome + ' ' + Giocatori.Nome) As Giocatore, Giocatori.idGiocatore " &
									"FROM (EventiPartita LEFT JOIN Giocatori ON (EventiPartita.idGiocatore = Giocatori.idGiocatore) AND (EventiPartita.idAnno = Giocatori.idAnno)) LEFT JOIN Eventi ON EventiPartita.idEvento = Eventi.idEvento " &
									"WHERE EventiPartita.idPartita=" & idPartita & " AND EventiPartita.idAnno=" & idAnno
								Rec2 = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec2) Is String Then
								Else
									Do Until Rec2.Eof
										Dim Path As String

										If Rec2("Giocatore").Value.Contains("Avversario") Then
											Path = PathBaseImmScon
										Else
											Path = PathBaseImmagini & "/" & Squadra & "/Giocatori/" & idAnno & "_" & Rec2("idGiocatore").Value & ".jpg"
										End If

										Eventi.Append("<tr>")
										Eventi.Append("<td align=""right"">")
										Eventi.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Rec2("idTempo").Value & "</span>")
										Eventi.Append("</td>")
										Eventi.Append("<td align=""right"">")
										Eventi.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Rec2("Minuto").Value & "°</span>")
										Eventi.Append("</td>")
										Eventi.Append("<td align=""left"">")
										Eventi.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Rec2("Descrizione").Value & "</span>")
										Eventi.Append("</td>")
										Eventi.Append("<td>")
										Eventi.Append("<img src=""" & Path & """ style=""width: 30px; height: 30px;"" onerror=""this.src='" & PathBaseImmScon & "'"" />")
										Eventi.Append("</td>")
										Eventi.Append("<td align=""left"">")
										Eventi.Append("<span class=""testo nero"" style=""font-size: 13px;"">" & Rec2("Giocatore").Value & "</span>")
										Eventi.Append("</td>")
										Eventi.Append("</tr>")

										Rec2.MoveNext
									Loop
									Rec2.Close
								End If

								Eventi.Append("</table>")
								Filone = Filone.Replace("***RACCONTO***", Eventi.ToString)

								' Risultato
								If CiSonoRigori Then
									QuantiGoal1 += RigoriSegnatiPropri
									QuantiGoal2 += RigoriSegnatiAvversari
								End If

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
				Ritorno = "Nessun dato rilevato"
			End If
		End If

		Return Ritorno
	End Function

	Public Function CriptaStringa(Stringa As String) As String
		'Dim Ancora As Boolean = True
		Dim Ritorno As String = ""

		'Do While Ancora = True
		'	Dim rnd1 As New Random(CInt(Date.Now.Ticks And &HFFFF))
		'	Dim Chiave As Integer = rnd1.Next(32) + 32
		'	Ritorno = Chr(Chiave)
		'	For i As Integer = 1 To 13
		'		Dim x As Integer = 64 + rnd1.Next(48)
		'		Ritorno &= Chr(x)
		'	Next
		'	Ritorno &= Chr(32 + Stringa.Length)
		'	Dim Contatore As Integer = 0
		'	For i As Integer = 1 To Stringa.Length
		'		Dim carattere As String = Mid(Stringa, i, 1)
		'		Dim ascii As Integer = Asc(carattere) + Chiave + Contatore
		'		Dim ascii2 As String = Chr(ascii)
		'		Ritorno &= ascii2
		'		Contatore += 2
		'	Next
		'	For i As Integer = 1 To 7
		'		Dim x As Integer = 64 + rnd1.Next(48)
		'		Ritorno &= Chr(x)
		'	Next

		'	If Not Ritorno.Contains(";") And Not Ritorno.Contains("'") Then
		'		Ancora = False
		'	End If
		'Loop

		Dim wrapper As New CryptEncrypt("WPippoBaudo227!")
		Ritorno = wrapper.EncryptData(Stringa)

		Return Ritorno
	End Function

	Public Function DecriptaStringa(Stringa As String) As String
		Dim Ritorno As String = ""
		'Dim Chiave As Integer = Asc(Mid(Stringa, 1, 1))
		'' Dim Chiave As Integer = Asc(car)
		'Dim Lunghezza As Integer = Asc(Mid(Stringa, 15, 1)) - 32
		'Dim Contatore As Integer = 0
		'For i As Integer = 16 To 16 + Lunghezza - 1
		'	Dim Car As String = Mid(Stringa, i, 1)
		'	Dim Car1 As Integer = Asc(Car) - Chiave - Contatore
		'	Dim c As String = Chr(Car1)
		'	Ritorno &= c
		'	Contatore += 2
		'Next

		Dim wrapper As New CryptEncrypt("WPippoBaudo227!")

		Try
			Ritorno = wrapper.DecryptData(Stringa)
		Catch ex As System.Security.Cryptography.CryptographicException
			Ritorno = ""
		End Try

		Return Ritorno
	End Function

	Public Function EliminaPartita(Squadra As String, idAnno As String, idPartita As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(HttpContext.Current.Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim sql As String
				Dim Ok As Boolean = True

				sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					sql = "delete from Partite Where idAnno = " & idAnno & " And idPartita = " & idPartita
					Ritorno = EsegueSql(Conn, sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

					If Ok Then
						sql = "delete from ArbitriPartite Where idAnno = " & idAnno & " And idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from Convocati Where idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from CoordinatePartite Where idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from EventiPartita Where idAnno = " & idAnno & " And idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from Marcatori Where idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from MeteoPartite Where idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from RigoriAvversari Where idAnno = " & idAnno & " And idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from RigoriPropri Where idAnno = " & idAnno & " And idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from Risultati Where idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from RisultatiAggiuntivi Where idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from RisultatiAggiuntiviMarcatori Where idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from TempiGoalAvversari Where idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If

					If Ok Then
						sql = "delete from EventiCalendario Where idPartita = " & idPartita
						Ritorno = EsegueSql(Conn, sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					End If
				Else
					Ok = False
				End If

				If Ok Then
					sql = "commit"
					Dim Ritorno2 As String = EsegueSql(Conn, sql, Connessione)
				Else
					sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, sql, Connessione)
				End If
			End If
		End If

		Return Ritorno
	End Function
End Module
