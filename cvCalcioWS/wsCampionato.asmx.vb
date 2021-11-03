Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://cvcalcio_cam.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsCampionato
    Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaCampionatoCategoria(Squadra As String, ByVal idAnno As String, idCategoria As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

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

				Dim idGiornata As String = RitornaGiornataUtenteCategoria(Squadra, idUtente, idAnno, idCategoria)

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
						"FROM AvversariCalendario LEFT JOIN SquadreAvversarie ON AvversariCalendario.idAvversario = SquadreAvversarie.idAvversario " &
						"Left Join CampiAvversari On SquadreAvversarie.idCampo = CampiAvversari.idCampo " &
						"Left Join AvversariCoord On SquadreAvversarie.idAvversario = AvversariCoord.idAvversario " &
						"WHERE AvversariCalendario.idAnno=" & idAnno & " And AvversariCalendario.idCategoria=" & idCategoria & " " &
						"ORDER BY AvversariCalendario.idProgressivo"
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
						"FROM CalendarioPartite LEFT JOIN CalendarioDate ON CalendarioPartite.idAnno = CalendarioDate.idAnno And CalendarioPartite.idCategoria = CalendarioDate.idCategoria " &
						"And CalendarioPartite.idGiornata = CalendarioDate.idGiornata And CalendarioPartite.idPartita = CalendarioDate.idPartita " &
						"LEFT JOIN SquadreAvversarie ON CalendarioPartite.idSqCasa = SquadreAvversarie.idAvversario " &
						"LEFT JOIN SquadreAvversarie AS SquadreAvversarie_1 ON CalendarioPartite.idSqFuori = SquadreAvversarie_1.idAvversario " &
						"LEFT JOIN Partite ON CalendarioPartite.idPartitaGen = Partite.idUnioneCalendario And CalendarioPartite.idCategoria = Partite.idCategoria " &
						"LEFT JOIN Risultati ON Partite.idPartita = Risultati.idPartita " &
						"LEFT JOIN RisultatiAggiuntivi ON Partite.idPartita = RisultatiAggiuntivi.idPartita " &
						"WHERE CalendarioPartite.idCategoria=" & idCategoria & " And CalendarioPartite.idAnno=" & idAnno & " " &
						"ORDER BY CalendarioPartite.idGiornata, CalendarioPartite.idPartita"
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
	Public Function CalcolaClassificaAllaGiornata(Squadra As String, ByVal idAnno As String, idCategoria As String, idGiornata As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

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
						If Rec("idAvversario").Value <> 999 Then
							idSquadre.Add(Rec("idAvversario").Value)
							Squadre.Add(Rec("Descrizione").Value)
							Giocate.Add(0)
							Vinte.Add(0)
							Pareggiate.Add(0)
							Perse.Add(0)
							Punti.Add(0)
							gFatti.Add(0)
							gSubiti.Add(0)
						End If

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
								Dim idC As Integer = Rec("idSqCasa").Value
								Dim idF As Integer = Rec("idSqFuori").Value

								If idC = 9999 Then idC = -1
								If idF = 9999 Then idF = -1

								If idC = i Then
									Indice1 = AppoIndice
								End If
								If idF = i Then
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
							Dim p1 As String = Format(Punti(i), "00") + Format(gFatti(i) - gSubiti(i), "00") + Format(gFatti(i), "00") + Format(Giocate(i), "00") + Squadre(i)
							Dim p2 As String = Format(Punti(k), "00") + Format(gFatti(k) - gSubiti(k), "00") + Format(gFatti(k), "00") + Format(Giocate(k), "00") + Squadre(k)

							If p1 > p2 Then
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

			Dim Ritorno2 As String = SalvaGiornataUtenteCategoria(Squadra, idUtente, idAnno, idCategoria, idGiornata)

			Conn.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function AggiungeSquadraAvversaria(Squadra As String, ByVal idAnno As String, idCategoria As String, idAvversario As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

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
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Dim idAvv() As String = {}
					If idAvversario.Contains(";") Then
						idAvv = idAvversario.Split(";")
					Else
						ReDim idAvv(0)
						idAvv(0) = idAvversario
					End If

					For Each id As String In idAvv
						If id <> "" Then
							Try
								Sql = "SELECT Max(idProgressivo)+1 FROM AvversariCalendario Where idAnno=" & idAnno & " And idCategoria=" & idCategoria
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Rec.Eof Then
										ProgressivoSquadra = "1"
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
										" " & id & " " &
										")"
									Ritorno = EsegueSql(Conn, Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If
								Else
									Ritorno = StringaErrore & " Problemi nel rilevamento del progressivo squadra"
									Ok = False
									Exit For
								End If
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
								Exit For
							End Try
						End If
					Next
				Else
					Ok = False
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaSquadraAvversaria(Squadra As String, ByVal idAnno As String, idCategoria As String, idAvversario As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Delete From AvversariCalendario Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idAvversario=" & idAvversario
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function InserisceNuovaPartita(Squadra As String, ByVal idAnno As String, idGiornata As String, idCategoria As String, Data As String,
										  Ora As String, Casa As String, Fuori As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idNuovaPartita As Integer = -1
		Dim idNuovaPartita1 As Integer = -1
		Dim idNuovaPartita2 As Integer = -1
		Dim ProgressivoPartita As String = ""
		' Dim idUnioneCalendario As String = ""

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
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
									idNuovaPartita1 = 1
								Else
									idNuovaPartita1 = Rec(0).Value
								End If
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					Try
						Sql = "SELECT Max(idPartita)+1 FROM CalendarioPartite"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = StringaErrore & " Nessun progressivo partita rilevato"
							Else
								If Rec(0).Value Is DBNull.Value Then
									idNuovaPartita2 = 1
								Else
									idNuovaPartita2 = Rec(0).Value
								End If
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If idNuovaPartita1 >= idNuovaPartita2 Then
						idNuovaPartita = idNuovaPartita1
					Else
						idNuovaPartita = idNuovaPartita2
					End If

					If Ok Then
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
							Ok = False
						End Try
					End If

					If Ok Then
						Dim c() As String = Casa.Split(";")
						Dim f() As String = Fuori.Split(";")

						'Try
						'	Sql = "SELECT Max(idPartitaGen)+1 FROM CalendarioPartite"
						'	Rec = LeggeQuery(Conn, Sql, Connessione)
						'	If TypeOf (Rec) Is String Then
						'		Ritorno = Rec
						'	Else
						'		If Rec.Eof Then
						'			Ritorno = StringaErrore & " Nessun progressivo generale rilevato"
						'		Else
						'			If Rec(0).Value Is DBNull.Value Then
						'				idUnioneCalendario = "1"
						'			Else
						'				idUnioneCalendario = Rec(0).Value.ToString
						'			End If
						'		End If
						'		Rec.Close()
						'	End If
						'Catch ex As Exception
						'	Ritorno = StringaErrore & " " & ex.Message
						'	Ok = False
						'End Try

						If Ok Then
							Try
								Sql = "Insert Into CalendarioPartite Values (" & idAnno & ", " & idCategoria & ", " & idGiornata & ", " & idNuovaPartita & ", " & c(0) & ", " & f(0) & ", " & idNuovaPartita & ", 'N', '')"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If

							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try

							If Ok Then
								If Mid(Ora, 1, 3) = "24:" Then Ora = "00:" & Mid(Ora, 4, Ora.Length)
								Dim dd As String = Data
								Dim oo As String = Ora
								If dd.Contains("/") Then
									Dim dc() As String = dd.Split("/")
									dd = dc(0) & "-" & dc(1) & "-" & dc(2)
								End If
								Dim ddoo As String = dd & " " & Ora

								Try
									Sql = "Insert Into CalendarioDate Values (" & idAnno & ", " & idCategoria & ", " & idNuovaPartita & ", " &
											idGiornata & ", '" & ddoo & "')"
									Ritorno = EsegueSql(Conn, Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If

								Catch ex As Exception
									Ritorno = StringaErrore & " " & ex.Message
									Ok = False
								End Try

								If Ok Then
									'Try
									'    Sql = "Insert Into CalendarioRisultati Values (" & idNuovaPartita & ", '')"
									'    Ritorno = EsegueSql(Conn, Sql, Connessione)
									'Catch ex As Exception
									'    Ritorno = StringaErrore & " " & ex.Message
									'End Try

									Dim Anticipo As Integer = 45

									If Not Ritorno.Contains(StringaErrore & "") Then
										If Val(c(0)) = -1 Or Val(f(0)) = -1 Or Val(c(0)) = 9999 Or Val(f(0)) = 9999 Then
											Try
												Sql = "SELECT AnticipoConvocazione FROM Categorie Where idAnno=" & idAnno & " And idCategoria=" & idCategoria
												Rec = LeggeQuery(Conn, Sql, Connessione)
												If TypeOf (Rec) Is String Then
													Ritorno = Rec
												Else
													If Not Rec.Eof Then
														Anticipo = Rec(0).Value.ToString * 90
													End If
													Rec.Close()
												End If
											Catch ex As Exception
												Ritorno = StringaErrore & " " & ex.Message
												Ok = False
											End Try

											Dim idAvversario As Integer
											Dim Datella As Date = Data & " " & Ora
											Dim dOraConv As Date = Datella.AddMinutes(-Anticipo)
											Dim OraConv As String = "00:" & Format(dOraConv.Hour, "00") & ":" & Format(dOraConv.Minute, "00")
											Dim inCasa As String = ""

											If Val(c(0)) = -1 Or Val(c(0)) = 9999 Then
												idAvversario = f(0)
												inCasa = "S"
											Else
												idAvversario = c(0)
												inCasa = "N"
											End If

											Dim idCampo As Integer
											Dim SquadraAvversaria As String = ""

											Try
												Sql = "SELECT idCampo, Descrizione FROM SquadreAvversarie Where idAvversario=" & idAvversario & " And Eliminato='N'"
												Rec = LeggeQuery(Conn, Sql, Connessione)
												If TypeOf (Rec) Is String Then
													Ritorno = Rec
												Else
													If Rec.Eof Then
														Ritorno = StringaErrore & " Nessun campo rilevato"
													Else
														idCampo = Rec(0).Value.ToString
														SquadraAvversaria = Rec(1).Value.ToString
													End If
													Rec.Close()
												End If
											Catch ex As Exception
												Ritorno = StringaErrore & " " & ex.Message
												Ok = False
											End Try

											If Ok Then
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
													Ok = False
												End Try

												If Ok Then
													Try
														Sql = "Insert Into ArbitriPartite Values (" & idAnno & ", " & idNuovaPartita & ", 1, 3)"
														Ritorno = EsegueSql(Conn, Sql, Connessione)
														If Ritorno.Contains(StringaErrore) Then
															Ok = False
														End If

													Catch ex As Exception
														Ritorno = StringaErrore & " " & ex.Message
														Ok = False
													End Try
												End If

												If Ok Then
													Dim LuogoAppuntamento As String = ""

													Try
														Sql = "Select * From CampiAvversari Where idCampo = " & idCampo
														Rec = LeggeQuery(Conn, Sql, Connessione)
														If TypeOf (Rec) Is String Then
															Ritorno = Rec
														Else
															If Rec.Eof Then
																Ritorno = StringaErrore & " Nessun campo rilevato"
															Else
																LuogoAppuntamento = "Campo '" & Rec(1).Value.ToString & "' " & Rec(2).Value.ToString
															End If
															Rec.Close()
														End If
													Catch ex As Exception
														Ritorno = StringaErrore & " " & ex.Message
														Ok = False
													End Try

													If Ok Then
														Try
															Sql = "Insert Into Partite Values (" & idAnno & ", " & idNuovaPartita & ", " & idCategoria & ", " &
																"" & idAvversario & ", " & idAllenatore & ", '" & Data & " " & Ora & "', " &
																"'N', '" & inCasa & "', 1, " & idCampo & ", '" & OraConv & "', -1, 'N', " &
																"'" & Data & " " & OraConv & "', '" & SistemaStringa(LuogoAppuntamento) & "', 'A', 'N', 2, 'N'" &
																")"
															Ritorno = EsegueSql(Conn, Sql, Connessione)
															If Ritorno.Contains(StringaErrore) Then
																Ok = False
															End If
														Catch ex As Exception
															Ritorno = StringaErrore & " " & ex.Message
															Ok = False
														End Try
													End If

													Dim idConvocazione As Integer

													If Ok Then
														Try
															Sql = "Select Max(idEvento) + 1 From EventiConvocazioni"
															Rec = LeggeQuery(Conn, Sql, Connessione)
															If TypeOf (Rec) Is String Then
																Ritorno = Rec
															Else
																If Rec(0).Value Is DBNull.Value Then
																	idConvocazione = 1
																Else
																	idConvocazione = Rec(0).Value
																End If
																Rec.Close()
															End If
														Catch ex As Exception
															Ritorno = StringaErrore & " " & ex.Message
															Ok = False
														End Try

													End If

													If Ok Then
														Try
															Dim Titolo As String = "Partita di campionato "
															If inCasa = "S" Then
																Titolo &= "in casa "
															Else
																Titolo &= "fuori casa "
															End If
															Titolo &= "contro " & SquadraAvversaria

															Dim ddd As String() = Data.Split("/")
															Dim dddd As String = ddd(2) & "-" & ddd(1) & "-" & ddd(0)
															Dim dor As String = dddd & " " & Ora

															Sql = "Insert Into EventiConvocazioni Values (" &
																" " & idConvocazione & ", 1, '" & SistemaStringa(Titolo) & "', " &
																"'" & dor & "', '" & dor & "', 'N', '#a00', " &
																"'#a0a', '', '', " & idNuovaPartita &
																")"
															Ritorno = EsegueSql(Conn, Sql, Connessione)
															If Ritorno.Contains(StringaErrore) Then
																Ok = False
															End If
														Catch ex As Exception
															Ritorno = StringaErrore & " " & ex.Message
															Ok = False
														End Try
													End If

													If Ok Then
														Try
															Sql = "Select * From Giocatori Where Categorie Like '%" & idCategoria & "-%' And Eliminato = 'N'"
															Rec = LeggeQuery(Conn, Sql, Connessione)
															If TypeOf (Rec) Is String Then
																Ritorno = Rec
															Else
																If Not Rec.Eof Then
																	Dim Progressivo As Integer = 0

																	Do Until Rec.Eof
																		Progressivo += 1
																		Sql = "Insert Into Convocati Values (" &
																			" " & idNuovaPartita & ", " &
																			" " & Progressivo & ", " &
																			" " & Rec("idGiocatore").Value & " " &
																			")"
																		Ritorno = EsegueSql(Conn, Sql, Connessione)
																		If Ritorno.Contains(StringaErrore) Then
																			Ok = False
																			Exit Do
																		End If

																		Rec.MoveNext
																	Loop
																	Rec.Close()
																End If
															End If
														Catch ex As Exception
															Ritorno = StringaErrore & " " & ex.Message
															Ok = False
														End Try
													End If
												End If
											End If
										End If
									End If
								End If
							End If
						End If

						If Ok Then
							Ritorno = idGiornata & ";" & idNuovaPartita & ";" & ProgressivoPartita & ";" & Data & ";" & Ora & ";" & Casa & Fuori & idNuovaPartita & ";"
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

								Try
									Sql = "Delete From Partite Where idAnno=" & idAnno & " And idPartita=" & idNuovaPartita & " And idCategoria=" & idCategoria & " " 'And idUnioneCalendario=" & idUnioneCalendario
									Ritorno = EsegueSql(Conn, Sql, Connessione)
								Catch ex As Exception
								End Try

								Try
									Sql = "Delete From Convocati Where idPartita=" & idNuovaPartita
									Ritorno = EsegueSql(Conn, Sql, Connessione)
								Catch ex As Exception
								End Try
							End If

							Ritorno = Appoggio
						End If
					End If
				Else
					Ok = False
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaPartita(Squadra As String, ByVal idAnno As String, idGiornata As String, idCategoria As String, idPartita As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				' Dim idUnioneCalendario As Integer = -1
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					'Sql = "Select * From Partite Where idAnno=" & idAnno & " And idPartita=" & idPartita
					'Try
					'	Rec = LeggeQuery(Conn, Sql, Connessione)
					'	If TypeOf (Rec) Is String Then
					'		Ritorno = Rec
					'	Else
					'		If Rec.Eof Then
					'			Ritorno = "*"
					'		Else
					'			' idUnioneCalendario = Rec("idUnioneCalendario").Value
					'		End If
					'		Rec.Close()
					'	End If
					'Catch ex As Exception
					'	Ritorno = StringaErrore & " " & ex.Message
					'	Ok = False
					'End Try

					If Ok Then
						'If idUnioneCalendario <> -1 Then
						Try
							Sql = "Delete From Partite Where idAnno=" & idAnno & " And idPartita=" & idPartita
							Ritorno = EsegueSql(Conn, Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							End If
						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
						'End If
					End If

					If Ok Then
						Try
							Sql = "Delete From ArbitriPartite Where idPartita=" & idPartita
							Ritorno = EsegueSql(Conn, Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							End If

						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try

						If Ok Then
							Try
								Sql = "Delete From CalendarioRisultati Where idPartita=" & idPartita
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If

							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try

							If Ok Then
								'Dim idPartitaGiornata As Integer = -1
								'Sql = "Select * From CalendarioPartite Where idAnno=" & idAnno & " And idPartitaGen=" & idPartita
								'Try
								'	Rec = LeggeQuery(Conn, Sql, Connessione)
								'	If TypeOf (Rec) Is String Then
								'		Ritorno = Rec
								'	Else
								'		If Rec.Eof Then
								'			Ritorno = StringaErrore & " nessun idPartita della giornata rilevato"
								'		Else
								'			idPartitaGiornata = Rec("idPartita").Value
								'		End If
								'		Rec.Close()
								'	End If
								'Catch ex As Exception
								'	Ritorno = StringaErrore & " " & ex.Message
								'	Ok = False
								'End Try

								If Ok Then
									Try
										Sql = "Delete From CalendarioDate Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata & " And idPartita=" & idPartita
										Ritorno = EsegueSql(Conn, Sql, Connessione)
										If Ritorno.Contains(StringaErrore) Then
											Ok = False
										End If
									Catch ex As Exception
										Ritorno = StringaErrore & " " & ex.Message
										Ok = False
									End Try

									If Ok Then
										Try
											Sql = "Delete From CalendarioPartite Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata & " And idPartita=" & idPartita
											Ritorno = EsegueSql(Conn, Sql, Connessione)
											If Ritorno.Contains(StringaErrore) Then
												Ok = False
											End If
										Catch ex As Exception
											Ritorno = StringaErrore & " " & ex.Message
											Ok = False
										End Try

										If Ok Then
											Try
												Sql = "Delete From EventiConvocazioni Where idPartita=" & idPartita
												Ritorno = EsegueSql(Conn, Sql, Connessione)
												If Ritorno.Contains(StringaErrore) Then
													Ok = False
												End If
											Catch ex As Exception
												Ritorno = StringaErrore & " " & ex.Message
												Ok = False
											End Try

											If Ok Then
												Ritorno = idGiornata & ";" & idPartita & ";"
											End If
										End If
									End If
								End If
							End If
						End If
					End If
				Else
					Ok = False
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ModificaPartitaAltre(Squadra As String, ByVal idAnno As String, idGiornata As String, idCategoria As String, Data As String,
									Ora As String, Casa As String, Fuori As String, idUnioneCalendario As String,
									ProgressivoPartita As String, Risultato As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
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
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
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
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Ok Then
						Try
							Sql = "Update CalendarioDate Set " &
								"Datella='" & Data & " " & Ora & "' " &
								"Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata & " And idPartita=" & ProgressivoPartita
							Ritorno = EsegueSql(Conn, Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							End If

						Catch ex As Exception
							Ritorno = StringaErrore & " " & ex.Message
							Ok = False
						End Try
					End If

					If Ok Then
						If Risultato <> "" Then
							Try
								Sql = "Delete From CalendarioRisultati Where idPartita=" & idUnioneCalendario
								Ritorno = EsegueSql(Conn, Sql, Connessione)
							Catch ex As Exception
							End Try

							Try
								Sql = "Insert Into CalendarioRisultati Values (" & idUnioneCalendario & ", '" & Risultato & "')"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If

							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
								Ok = False
							End Try
						End If
					End If
				Else
					Ok = False
				End If

				If Not Ritorno.Contains("ERROR") Then
					Ritorno = idGiornata & ";" & idUnioneCalendario & ";" & ProgressivoPartita & ";" & Data & ";" & Ora & ";" & Casa & Fuori & Giocata & ";" & Risultato & ";"
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				End If
			End If

			Conn.Close()
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaIdPartitaDaUnione(Squadra As String, idUnioneCalendario As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

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
	Public Function SalvaGiornataUtenteCategoria(Squadra As String, idUtente As String, idAnno As String, idCategoria As String, idGiornata As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Delete From Giornata Where idUtente=" & idUtente & " And idAnno=" & idAnno & " And idCategoria=" & idCategoria
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					Try
						Sql = "Insert Into Giornata Values (" & idUtente & ", " & idAnno & ", " & idCategoria & ", " & idGiornata & ")"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try
				Else
					Ok = False
				End If

				If Ok Then
					Sql = "commit"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaGiornataUtenteCategoria(Squadra As String, idUtente As String, idAnno As String, idCategoria As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
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
				Dim Ok As Boolean = True

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
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								Else
									Ritorno = idGiornata
								End If
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

	<WebMethod()>
	Public Function Statistiche(Squadra As String, idCategoria As String, idAnno As String, Stampa As String, idGiornata As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim StampaMarcatori As String = ""
		Dim StampaPresenze As String = ""
		Dim StampaFgF As String = ""
		Dim StampaFgs As String = ""
		Dim StampaEventi As String = ""
		Dim StampaStatistiche As String = ""
		Dim StampaTipoPartite As String = ""
		Dim StampaAvversari As String = ""
		Dim StampaMeteo As String = ""

		Dim paths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim P() As String = paths.Split(";")
		If Strings.Right(P(0), 1) <> "\" Then
			P(0) &= "\"
		End If
		Dim pathAllegati As String = P(0).Replace(vbCrLf, "")
		If Strings.Right(P(2), 1) <> "/" Then
			P(2) &= "/"
		End If
		Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
		Dim PathBaseMultimedia As String = pathMultimedia.Replace("Allegati", "Multimedia")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Dim NomeSquadra As String = ""
				Dim ss() As String = Squadra.Split("_")
				Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & Val(ss(1)).ToString
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = "Problemi lettura squadra"
				Else
					If Rec.Eof Then
					Else
						NomeSquadra = "" & Rec("Descrizione").Value
					End If
					Rec.Close
				End If

				Dim filetto As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_Goals.txt")

				If Ok Then
					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%' Order By Totale Desc, GoalCampionato Desc, GoalAmichevole Desc"
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Stampa = "S" Then
								StampaMarcatori &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaMarcatori &= "<tr>" & vbCrLf
								StampaMarcatori &= "<th></th>" & vbCrLf
								StampaMarcatori &= "<th>Nominativo</th>" & vbCrLf
								StampaMarcatori &= "<th>Ruolo</th>" & vbCrLf
								StampaMarcatori &= "<th>Goal Amichevole</th>" & vbCrLf
								StampaMarcatori &= "<th>Goal Campionato</th>" & vbCrLf
								StampaMarcatori &= "<th>Totale</th>" & vbCrLf
								StampaMarcatori &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof
								Ritorno &= Rec("Cognome").Value & ";"
								Ritorno &= Rec("Nome").Value & ";"
								Ritorno &= Rec("Soprannome").Value & ";"
								Ritorno &= Rec("Ruolo").Value & ";"
								Ritorno &= Rec("GoalAmichevole").Value & ";"
								Ritorno &= Rec("GoalCampionato").Value & ";"
								Ritorno &= Rec("Totale").Value & ";"
								Ritorno &= Rec("idGiocatore").Value & ";"
								Ritorno &= "§"

								If Stampa = "S" Then
									Dim Soprannome As String = Rec("Soprannome").Value
									If Soprannome <> "" Then Soprannome = "'" & Soprannome & "' "
									Dim Nominativo As String = Rec("Nome").Value & " " & Soprannome & Rec("Cognome").Value
									Dim Path As String = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
									Path = DecriptaImmagine(Path)

									StampaMarcatori &= "<tr>" & vbCrLf
									StampaMarcatori &= "<td><img src=""" & Path & """ width=""50"" height=""50"" /></td>" & vbCrLf
									StampaMarcatori &= "<td>" & Nominativo & "</td>" & vbCrLf
									StampaMarcatori &= "<td>" & Rec("Ruolo").Value & "</td>" & vbCrLf
									StampaMarcatori &= "<td style=""text-align: right;"">" & Rec("GoalAmichevole").Value & "</td>" & vbCrLf
									StampaMarcatori &= "<td style=""text-align: right;"">" & Rec("GoalCampionato").Value & "</td>" & vbCrLf
									StampaMarcatori &= "<td style=""text-align: right;"">" & Rec("Totale").Value & "</td>" & vbCrLf
									StampaMarcatori &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaMarcatori &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_Presenze.txt")
					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%' Order By Totale Desc, PresenzeCampionato Desc, PresenzeAmichevole Desc"
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Ritorno &= "|"

							If Stampa = "S" Then
								StampaPresenze &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaPresenze &= "<tr>" & vbCrLf
								StampaPresenze &= "<th></th>" & vbCrLf
								StampaPresenze &= "<th>Nominativo</th>" & vbCrLf
								StampaPresenze &= "<th>Ruolo</th>" & vbCrLf
								StampaPresenze &= "<th>Presenze Amichevole</th>" & vbCrLf
								StampaPresenze &= "<th>Presenze Campionato</th>" & vbCrLf
								StampaPresenze &= "<th>Totale</th>" & vbCrLf
								StampaPresenze &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof
								Ritorno &= Rec("Cognome").Value & ";"
								Ritorno &= Rec("Nome").Value & ";"
								Ritorno &= Rec("Soprannome").Value & ";"
								Ritorno &= Rec("Ruolo").Value & ";"
								Ritorno &= Rec("PresenzeAmichevole").Value & ";"
								Ritorno &= Rec("PresenzeCampionato").Value & ";"
								Ritorno &= Rec("Totale").Value & ";"
								Ritorno &= Rec("idGiocatore").Value & ";"
								Ritorno &= "§"

								If Stampa = "S" Then
									Dim Soprannome As String = Rec("Soprannome").Value
									If Soprannome <> "" Then Soprannome = "'" & Soprannome & "' "
									Dim Nominativo As String = Rec("Nome").Value & " " & Soprannome & Rec("Cognome").Value
									Dim Path As String = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
									Path = DecriptaImmagine(Path)

									StampaPresenze &= "<tr>" & vbCrLf
									StampaPresenze &= "<td><img src=""" & Path & """ width=""50"" height=""50"" /></td>" & vbCrLf
									StampaPresenze &= "<td>" & Nominativo & "</td>" & vbCrLf
									StampaPresenze &= "<td>" & Rec("Ruolo").Value & "</td>" & vbCrLf
									StampaPresenze &= "<td style=""text-align: right;"">" & Rec("PresenzeAmichevole").Value & "</td>" & vbCrLf
									StampaPresenze &= "<td style=""text-align: right;"">" & Rec("PresenzeCampionato").Value & "</td>" & vbCrLf
									StampaPresenze &= "<td style=""text-align: right;"">" & Rec("Totale").Value & "</td>" & vbCrLf
									StampaPresenze &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaPresenze &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_FasceGoalFatti.txt")
					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%'"
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Ritorno &= "|"

							If Stampa = "S" Then
								StampaFgF &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaFgF &= "<tr>" & vbCrLf
								StampaFgF &= "<th>Tipologia</th>" & vbCrLf
								StampaFgF &= "<th>Fascia</th>" & vbCrLf
								StampaFgF &= "<th>Tempo</th>" & vbCrLf
								StampaFgF &= "<th>Goals</th>" & vbCrLf
								StampaFgF &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof
								Ritorno &= Rec("Tipologia").Value & ";"
								Ritorno &= Rec("Fascia").Value & ";"
								Ritorno &= Rec("idTempo").Value & ";"
								Ritorno &= Rec("Goals").Value & ";"
								Ritorno &= "§"

								If Stampa = "S" Then
									StampaFgF &= "<tr>" & vbCrLf
									StampaFgF &= "<td>" & Rec("Tipologia").Value & "</td>" & vbCrLf
									StampaFgF &= "<td>" & Rec("Fascia").Value & "</td>" & vbCrLf
									StampaFgF &= "<td style=""text-align: right;"">" & Rec("idTempo").Value & "</td>" & vbCrLf
									StampaFgF &= "<td style=""text-align: right;"">" & Rec("Goals").Value & "</td>" & vbCrLf
									StampaFgF &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaFgF &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					Sql = "Select C.idTipologia, idTempo, Minuti " &
					"From RisultatiAvversariMinuti A " &
					"Left Join Partite B On A.idPartita = B.idPartita " &
					"Left Join [Generale].[dbo].Tipologie C On B.idTipologia = C.idTipologia " &
					"Left Join Convocati E On A.idPartita = E.idPartita And E.idPartita = B.idPartita " &
					"Left Join Giocatori D On E.idGiocatore = D.idGiocatore And E.idProgressivo = 1 " &
					"Where E.idProgressivo = 1 And Categorie Like '%" & idCategoria & "-%'"
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Dim Fascia1(1, 2) As Integer
							Dim Fascia2(1, 2) As Integer
							Dim Fascia3(1, 2) As Integer
							Dim Fascia4(1, 2) As Integer
							Dim Fascia5(1, 2) As Integer

							Do Until Rec.Eof
								Dim Minuti() As String = Rec("Minuti").Value.split(";")
								Dim idTempo As Integer = Val(Rec("idTempo").Value) - 1
								Dim idTipologia As Integer = Val(Rec("idTipologia").Value) - 1

								For Each mm As String In Minuti
									If mm <> "" Then
										Dim m As Integer = Val(mm)

										If m < 10 Then
											Fascia1(idTipologia, idTempo) += 1
										Else
											If m > 9 And m < 20 Then
												Fascia2(idTipologia, idTempo) += 1
											Else
												If m > 19 And m < 30 Then
													Fascia3(idTipologia, idTempo) += 1
												Else
													If m > 29 And m < 40 Then
														Fascia4(idTipologia, idTempo) += 1
													Else
														If m > 39 Then
															Fascia5(idTipologia, idTempo) += 1
														End If
													End If
												End If
											End If
										End If
									End If
								Next
								Rec.MoveNext
							Loop
							Rec.Close()

							Ritorno &= "|"

							Dim Tipi() As String = {"Campionato", "Amichevole"}
							Dim Fascia() As String = {"0-9", "10-19", "20-29", "30-39", "40-"}

							If Stampa = "S" Then
								StampaFgs &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaFgs &= "<tr>" & vbCrLf
								StampaFgs &= "<th>Tipologia</th>" & vbCrLf
								StampaFgs &= "<th>Fascia</th>" & vbCrLf
								StampaFgs &= "<th>Tempo</th>" & vbCrLf
								StampaFgs &= "<th>Goals</th>" & vbCrLf
								StampaFgs &= "</tr>" & vbCrLf
							End If

							For i As Integer = 0 To 1
								For k As Integer = 0 To 2
									If Fascia1(i, k) > 0 Then
										Ritorno &= Tipi(i) & ";" & Fascia(k) & ";" & k + 1 & ";" & Fascia1(i, k) & "§"

										If Stampa = "S" Then
											StampaFgs &= "<tr>" & vbCrLf
											StampaFgs &= "<td>" & Tipi(i) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia(k) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & k + 1 & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia1(i, k) & "</td>" & vbCrLf
											StampaFgs &= "</tr>" & vbCrLf
										End If
									End If
									If Fascia2(i, k) > 0 Then
										Ritorno &= Tipi(i) & ";" & Fascia(k) & ";" & k + 1 & ";" & Fascia2(i, k) & "§"

										If Stampa = "S" Then
											StampaFgs &= "<tr>" & vbCrLf
											StampaFgs &= "<td>" & Tipi(i) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia(k) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & k + 1 & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia2(i, k) & "</td>" & vbCrLf
											StampaFgs &= "</tr>" & vbCrLf
										End If
									End If
									If Fascia3(i, k) > 0 Then
										Ritorno &= Tipi(i) & ";" & Fascia(k) & ";" & k + 1 & ";" & Fascia3(i, k) & "§"

										If Stampa = "S" Then
											StampaFgs &= "<tr>" & vbCrLf
											StampaFgs &= "<td>" & Tipi(i) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia(k) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & k + 1 & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia3(i, k) & "</td>" & vbCrLf
											StampaFgs &= "</tr>" & vbCrLf
										End If
									End If
									If Fascia4(i, k) > 0 Then
										Ritorno &= Tipi(i) & ";" & Fascia(k) & ";" & k + 1 & ";" & Fascia4(i, k) & "§"

										If Stampa = "S" Then
											StampaFgs &= "<tr>" & vbCrLf
											StampaFgs &= "<td>" & Tipi(i) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia(k) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & k + 1 & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia4(i, k) & "</td>" & vbCrLf
											StampaFgs &= "</tr>" & vbCrLf
										End If
									End If
									If Fascia5(i, k) > 0 Then
										Ritorno &= Tipi(i) & ";" & Fascia(k) & ";" & k + 1 & ";" & Fascia5(i, k) & "§"

										If Stampa = "S" Then
											StampaFgs &= "<tr>" & vbCrLf
											StampaFgs &= "<td>" & Tipi(i) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia(k) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & k + 1 & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia5(i, k) & "</td>" & vbCrLf
											StampaFgs &= "</tr>" & vbCrLf
										End If
									End If
								Next
							Next

							If Stampa = "S" Then
								StampaFgs &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_Eventi.txt")
					Sql = "Select Top 25 * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%' Order By Quanti Desc"
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Ritorno &= "|"

							If Stampa = "S" Then
								StampaEventi &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaEventi &= "<tr>" & vbCrLf
								StampaEventi &= "<th></th>" & vbCrLf
								StampaEventi &= "<th>Nominativo</th>" & vbCrLf
								StampaEventi &= "<th>Descrizione</th>" & vbCrLf
								StampaEventi &= "<th>Quanti</th>" & vbCrLf
								StampaEventi &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof
								Ritorno &= Rec("Descrizione").Value & ";"
								Ritorno &= Rec("Cognome").Value & ";"
								Ritorno &= Rec("Nome").Value & ";"
								Ritorno &= Rec("Soprannome").Value & ";"
								Ritorno &= Rec("Quanti").Value & ";"
								Ritorno &= Rec("idGiocatore").Value & ";"
								Ritorno &= "§"

								If Stampa = "S" Then
									Dim Soprannome As String = Rec("Soprannome").Value
									If Soprannome <> "" Then Soprannome = "'" & Soprannome & "' "
									Dim Nominativo As String = Rec("Nome").Value & " " & Soprannome & Rec("Cognome").Value
									Dim Path As String = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
									Path = DecriptaImmagine(Path)

									StampaEventi &= "<tr>" & vbCrLf
									StampaEventi &= "<td><img src=""" & Path & """ width=""50"" height=""50"" /></td>" & vbCrLf
									StampaEventi &= "<td>" & Nominativo & "</td>" & vbCrLf
									StampaEventi &= "<td>" & Rec("Descrizione").Value & "</td>" & vbCrLf
									StampaEventi &= "<td style=""text-align: right;"">" & Rec("Quanti").Value & "</td>" & vbCrLf
									StampaEventi &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaEventi &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_TipologiePartite.txt")
					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%'"
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Ritorno &= "|"

							If Stampa = "S" Then
								StampaTipoPartite &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaTipoPartite &= "<tr>" & vbCrLf
								StampaTipoPartite &= "<th>Dove</th>" & vbCrLf
								StampaTipoPartite &= "<th>Descrizione</th>" & vbCrLf
								StampaTipoPartite &= "<th>Quante</th>" & vbCrLf
								StampaTipoPartite &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof
								Dim Dove As String = ""

								Select Case Rec("Casa").Value
									Case "S"
										Dove = "In casa"
									Case "N"
										Dove = "Fuori casa"
									Case Else
										Dove = "Campo esterno"
								End Select
								Ritorno &= Dove & ";"
								Ritorno &= Rec("Descrizione").Value & ";"
								Ritorno &= Rec("Quante").Value & ";"
								Ritorno &= "§"

								If Stampa = "S" Then
									StampaTipoPartite &= "<tr>" & vbCrLf
									StampaTipoPartite &= "<td>" & Dove & "</td>" & vbCrLf
									StampaTipoPartite &= "<td>" & Rec("Descrizione").Value & "</td>" & vbCrLf
									StampaTipoPartite &= "<td>" & Rec("Quante").Value & "</td>" & vbCrLf
									StampaTipoPartite &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaTipoPartite &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_Partite.txt")
					filetto = filetto.Replace("%idCategoria%", idCategoria)
					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%'"
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Dim MaxGoals As Integer = 0
							Dim MaxGoalsCasa As Integer = 0
							Dim MaxGoalsFuori As Integer = 0
							Dim GoalMax As String = ""
							Dim GoalMaxCasa As String = ""
							Dim GoalMaxFuori As String = ""
							Dim Giocate As Integer = 0
							Dim Vittorie As Integer = 0
							Dim Pareggi As Integer = 0
							Dim Sconfitte As Integer = 0
							Dim GoalTotali As Integer = 0
							Dim SubitiTotali As Integer = 0

							Dim GiocateCasa As Integer = 0
							Dim VittorieCasa As Integer = 0
							Dim PareggiCasa As Integer = 0
							Dim SconfitteCasa As Integer = 0
							Dim GoalTotaliCasa As Integer = 0
							Dim SubitiTotaliCasa As Integer = 0

							Dim GiocateFuori As Integer = 0
							Dim VittorieFuori As Integer = 0
							Dim PareggiFuori As Integer = 0
							Dim SconfitteFuori As Integer = 0
							Dim GoalTotaliFuori As Integer = 0
							Dim SubitiTotaliFuori As Integer = 0

							Dim Avversari As New List(Of String)
							Dim Incontrati As New List(Of Integer)
							Dim idAvversario As New List(Of Integer)

							Do Until Rec.Eof
								Dim Avversario As String = Rec("Avversario").Value
								Dim Categoria As String = Rec("Categoria").Value
								Dim Risultato As String = Rec("Risultato").Value
								Dim GCasa As Integer = Val(Rec("Casa").Value)
								Dim GFuori As Integer = Val(Rec("Fuori").Value)

								Dim OkAvv As Boolean = True
								Dim ii As Integer = 0
								For Each a As String In Avversari
									If a = Avversario Then
										OkAvv = False
										Incontrati.Item(ii) += 1
										Exit For
									End If
									ii += 1
								Next
								If OkAvv Then
									Avversari.Add(Avversario)
									idAvversario.Add(Rec("idAvversario").Value)
									Incontrati.Add(1)
								End If

								Dim sqCasa As String = ""
								Dim sqFuori As String = ""
								Dim goalCasa As Integer
								Dim goalFuori As Integer

								Dim Dove As String = ""
								Dim TotaleCasa As Integer = 0
								Dim TotaleFuori As Integer = 0

								Giocate += 1

								Select Case Rec("Dove").Value
									Case "S"
										Dove = "In casa"

										sqCasa = Categoria
										sqFuori = Avversario
										goalCasa = GCasa
										goalFuori = GFuori

										GoalTotali += goalCasa
										SubitiTotali += goalFuori

										GoalTotaliCasa += goalCasa
										SubitiTotaliCasa += goalFuori

										TotaleCasa = goalCasa + goalFuori
										If TotaleCasa > MaxGoalsCasa Then
											GoalMaxCasa = "Max goal per partita in casa: " & sqCasa & "-" & sqFuori & " " & goalCasa & "-" & goalFuori & " -> Totale: " & TotaleCasa
											MaxGoalsCasa = TotaleCasa
										End If

										If goalCasa > goalFuori Then
											Vittorie += 1
											VittorieCasa += 1
										Else
											If goalCasa < goalFuori Then
												Sconfitte += 1
												SconfitteCasa += 1
											Else
												Pareggi += 1
												PareggiCasa += 1
											End If
										End If

										GiocateCasa += 1
									Case "N", "E"
										If Rec("Dove").Value = "N" Then
											Dove = "Fuori casa"
										Else
											Dove = "Campo esterno"
										End If

										sqCasa = Avversario
										sqFuori = Categoria
										goalCasa = GFuori
										goalFuori = GCasa

										GoalTotali += goalCasa
										SubitiTotali += goalFuori

										GoalTotaliFuori += goalCasa
										SubitiTotaliFuori += goalFuori

										TotaleFuori = goalCasa + goalFuori
										If TotaleFuori > MaxGoalsFuori Then
											GoalMaxFuori = "Max goal per partita fuori casa: " & sqCasa & "-" & sqFuori & " " & goalCasa & "-" & goalFuori & " -> Totale: " & TotaleFuori
											MaxGoalsFuori = TotaleFuori
										End If

										If goalCasa < goalFuori Then
											Vittorie += 1
											VittorieFuori += 1
										Else
											If goalCasa > goalFuori Then
												Sconfitte += 1
												SconfitteFuori += 1
											Else
												Pareggi += 1
												PareggiFuori += 1
											End If
										End If

										GiocateFuori += 1
								End Select

								Dim Totale As Integer = GCasa + GFuori

								If Totale > MaxGoals Then
									GoalMax = "Max goal totali: " & sqCasa & "-" & sqFuori & " " & goalCasa & "-" & goalFuori & " -> Totale: " & Totale
									MaxGoals = Totale
								End If

								Rec.MoveNext
							Loop
							Rec.Close()

							Ritorno &= "|"

							If Stampa = "S" Then
								StampaStatistiche &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaStatistiche &= "<tr>" & vbCrLf
								StampaStatistiche &= "<th>Descrizione</th>" & vbCrLf
								StampaStatistiche &= "</tr>" & vbCrLf
							End If

							Dim Media As Single = -1
							Dim Lista As New List(Of String)

							Ritorno &= GoalMax & "§"
							Lista.Add(GoalMax)
							Ritorno &= GoalMaxCasa & "§"
							Lista.Add(GoalMaxCasa)
							Ritorno &= GoalMaxFuori & "§"
							Lista.Add(GoalMaxFuori)

							Ritorno &= "Giocate: " & Giocate & "§"
							Lista.Add("Giocate: " & Giocate)
							Ritorno &= "Vittorie: " & Vittorie & "§"
							Lista.Add("Vittorie: " & Vittorie)
							Ritorno &= "Pareggi: " & Pareggi & "§"
							Lista.Add("Pareggi: " & Pareggi)
							Ritorno &= "Sconfitte: " & Sconfitte & "§"
							Lista.Add("Sconfitte: " & Sconfitte)
							Ritorno &= "Goal Fatti Totali: " & GoalTotali & "§"
							Lista.Add("Goal Fatti Totali: " & GoalTotali)
							Ritorno &= "Goal Subiti Totali: " & SubitiTotali & "§"
							Lista.Add("Goal Subiti Totali: " & SubitiTotali)
							Media = CInt((GoalTotali / Giocate) * 100) / 100
							Ritorno &= "Media goal fatti Totali: " & Media & "§"
							Lista.Add("Media goal fatti Totali: " & Media)
							Media = CInt((SubitiTotali / Giocate) * 100) / 100
							Ritorno &= "Media goal subiti Totali: " & Media & "§"
							Lista.Add("Media goal subiti Totali: " & Media)

							Ritorno &= "Giocate Casa: " & GiocateCasa & "§"
							Lista.Add("Giocate Casa: " & GiocateCasa)
							Ritorno &= "Vittorie Casa: " & VittorieCasa & "§"
							Lista.Add("Vittorie Casa: " & VittorieCasa)
							Ritorno &= "Pareggi Casa: " & PareggiCasa & "§"
							Lista.Add("Pareggi Casa: " & PareggiCasa)
							Ritorno &= "Sconfitte Casa: " & SconfitteCasa & "§"
							Lista.Add("Sconfitte Casa: " & SconfitteCasa)
							Ritorno &= "Goal Fatti Totali Casa: " & GoalTotaliCasa & "§"
							Lista.Add("Goal Fatti Totali Casa: " & GoalTotaliCasa)
							Ritorno &= "Goal Subiti Totali Casa: " & SubitiTotaliCasa & "§"
							Lista.Add("Goal Subiti Totali Casa: " & SubitiTotaliCasa)
							Media = CInt((GoalTotaliCasa / GiocateCasa) * 100) / 100
							Ritorno &= "Media goal fatti Casa: " & Media & "§"
							Lista.Add("Media goal fatti Casa: " & Media)
							Media = CInt((SubitiTotaliCasa / GiocateCasa) * 100) / 100
							Ritorno &= "Media goal subiti Casa: " & Media & "§"
							Lista.Add("Media goal subiti Casa: " & Media)

							Ritorno &= "Giocate Fuori: " & GiocateFuori & "§"
							Lista.Add("Giocate Fuori: " & GiocateFuori)
							Ritorno &= "Vittorie Fuori: " & VittorieFuori & "§"
							Lista.Add("Vittorie Fuori: " & VittorieFuori)
							Ritorno &= "Pareggi Fuori: " & PareggiFuori & "§"
							Lista.Add("Pareggi Fuori: " & PareggiFuori)
							Ritorno &= "Sconfitte Fuori: " & SconfitteFuori & "§"
							Lista.Add("Sconfitte Fuori: " & SconfitteFuori)
							Ritorno &= "Goal Fatti Totali Fuori: " & GoalTotaliFuori & "§"
							Lista.Add("Goal Fatti Totali Fuori: " & GoalTotaliFuori)
							Ritorno &= "Goal Subiti Totali Fuori: " & SubitiTotaliFuori & "§"
							Lista.Add("Goal Subiti Totali Fuori: " & SubitiTotaliFuori)
							Media = CInt((GoalTotaliFuori / GiocateFuori) * 100) / 100
							Ritorno &= "Media goal fatti Fuori: " & Media & "§"
							Lista.Add("Media goal fatti Fuori: " & Media)
							Media = CInt((SubitiTotaliFuori / GiocateFuori) * 100) / 100
							Ritorno &= "Media goal subiti Fuori: " & Media & "§"
							Lista.Add("Media goal subiti Fuori: " & Media)

							If Stampa = "S" Then
								For Each l As String In Lista
									StampaStatistiche &= "<tr>" & vbCrLf
									StampaStatistiche &= "<td>" & l & "</dh>" & vbCrLf
									StampaStatistiche &= "</tr>" & vbCrLf
								Next
								StampaStatistiche &= "</table>"
							End If

							Ritorno &= "|"

							If Stampa = "S" Then
								StampaAvversari &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaAvversari &= "<tr>" & vbCrLf
								StampaAvversari &= "<th></th>" & vbCrLf
								StampaAvversari &= "<th>Descrizione</th>" & vbCrLf
								StampaAvversari &= "</tr>" & vbCrLf
							End If

							For i As Integer = 0 To Incontrati.Count - 1
								For k As Integer = i + 1 To Incontrati.Count - 1
									If Incontrati.Item(i) < Incontrati(k) Then
										Dim a As String = Avversari.Item(i)
										Avversari.Item(i) = Avversari.Item(k)
										Avversari.Item(k) = a

										Dim ii As Integer = Incontrati.Item(i)
										Incontrati.Item(i) = Incontrati.Item(k)
										Incontrati.Item(k) = ii

										ii = idAvversario.Item(i)
										idAvversario.Item(i) = idAvversario.Item(k)
										idAvversario.Item(k) = ii
									End If
								Next
							Next

							Dim iii As Integer = 0
							For Each a As String In Avversari
								Ritorno &= idAvversario.Item(iii) & ";" & a & ";" & Incontrati.Item(iii) & "§"

								If Stampa = "S" Then
									Dim Imm2 As String = PathBaseMultimedia & "/" & NomeSquadra & "/Avversari/" & idAvversario.Item(iii) & ".kgb"
									Imm2 = DecriptaImmagine(Imm2)

									StampaAvversari &= "<tr>" & vbCrLf
									StampaAvversari &= "<td><img src=""" & Imm2 & """ width=""50"" height=""50"" /></td>" & vbCrLf
									StampaAvversari &= "<td>" & a & ": " & Incontrati.Item(iii) & "</td>" & vbCrLf
									StampaAvversari &= "</tr>" & vbCrLf
								End If

								iii += 1
							Next
							If Stampa = "S" Then
								StampaAvversari &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_Meteo.txt")
					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%' Order By Quante Desc, Tempo"
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Ritorno &= "|"

							If Stampa = "S" Then
								StampaMeteo &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaMeteo &= "<tr>" & vbCrLf
								StampaMeteo &= "<th></th>" & vbCrLf
								StampaMeteo &= "<th>Tempo</th>" & vbCrLf
								StampaMeteo &= "<th>Quante</th>" & vbCrLf
								StampaMeteo &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof
								Ritorno &= Rec("Tempo").Value & ";"
								Ritorno &= Rec("Icona").Value & ";"
								Ritorno &= Rec("Quante").Value & ";"
								Ritorno &= "§"

								If Stampa = "S" Then
									StampaMeteo &= "<tr>" & vbCrLf
									StampaMeteo &= "<td><img src=""" & Rec("Icona").Value & """ width=""50"" height=""50"" /></td>" & vbCrLf
									StampaMeteo &= "<td>" & Rec("Tempo").Value & "</td>" & vbCrLf
									StampaMeteo &= "<td>" & Rec("Quante").Value & "</td>" & vbCrLf
									StampaMeteo &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaMeteo &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						Ok = False
					End Try
				End If

				If Ok Then
					If Stampa = "S" Then
						Dim PathBaseImmagini As String = pathMultimedia

						filetto = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\statistiche_nuove.txt")
						filetto = filetto.Replace("***SFONDO***", PathBaseImmagini & "/bg.jpg")

						filetto = filetto.Replace("***MARCATORI***", StampaMarcatori)
						filetto = filetto.Replace("***PRESENZE***", StampaPresenze)
						filetto = filetto.Replace("***FGF***", StampaFgF)
						filetto = filetto.Replace("***FGS***", StampaFgs)
						filetto = filetto.Replace("***EVENTI***", StampaEventi)
						filetto = filetto.Replace("***STATISTICHE***", StampaStatistiche)
						filetto = filetto.Replace("***TIPOPARTITE***", StampaTipoPartite)
						filetto = filetto.Replace("***AVVERSARI***", StampaAvversari)
						filetto = filetto.Replace("***METEO***", StampaMeteo)

						'Dim Altro As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") ' & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
						Dim NomeFileFinale As String = pathAllegati & Squadra & "\Statistiche\Statistiche_Giornata_" & idGiornata & ".html"
						Dim NomeFileFinalePDF As String = pathAllegati & Squadra & "\Statistiche\Statistiche_Giornata_" & idGiornata & ".pdf"
						Dim PathLog As String = HttpContext.Current.Server.MapPath(".") & "\Log\Pdf.txt"
						gf.CreaDirectoryDaPercorso(NomeFileFinale)

						gf.EliminaFileFisico(NomeFileFinale)
						gf.EliminaFileFisico(NomeFileFinalePDF)
						gf.CreaAggiornaFile(NomeFileFinale, filetto)

						Dim pp As New pdfGest
						Ritorno = pp.ConverteHTMLInPDF(NomeFileFinale, NomeFileFinalePDF, PathLog, True)

						If Ritorno = "*" Then
							Ritorno = NomeFileFinalePDF.Replace(pathAllegati, pathMultimedia).Replace("Multimedia", "Allegati").Replace("\", "/")
						End If

					End If
				End If
			End If
		End If

		' Marcatori | Presenze | Fasce Goal Fatti | Fasce Goal Subiti | Eventi | Tipologie Partite | Partite | Avversari Incontrati | Meteo

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function StampaClassifica(Squadra As String, Classifica As String, Giornata As String, idAnno As String, idCategoria As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As String = ""
		Dim Righe() As String = Classifica.Split("§")
		Dim Stampa As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		Dim paths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim P() As String = paths.Split(";")
		If Strings.Right(P(0), 1) <> "\" Then
			P(0) &= "\"
		End If
		Dim pathAllegati As String = P(0).Replace(vbCrLf, "")
		If Strings.Right(P(2), 1) <> "/" Then
			P(2) &= "/"
		End If
		Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
		Dim PathBaseMultimedia As String = pathMultimedia.Replace("Allegati", "Multimedia")
		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Dim NomeSquadra As String = ""
				Dim ss() As String = Squadra.Split("_")
				Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & Val(ss(1)).ToString
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = "Problemi lettura squadra"
				Else
					If Rec.Eof Then
					Else
						NomeSquadra = "" & Rec("Descrizione").Value
					End If
					Rec.Close
				End If

				If Ok Then
					Stampa &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
					Stampa &= "<tr>" & vbCrLf
					Stampa &= "<th></th>" & vbCrLf
					Stampa &= "<th>Squadra</th>" & vbCrLf
					Stampa &= "<th style=""text-align: right;"">Punti</th>" & vbCrLf
					Stampa &= "<th style=""text-align: right;"">Giocate</th>" & vbCrLf
					Stampa &= "<th style=""text-align: right;"">Vinte</th>" & vbCrLf
					Stampa &= "<th style=""text-align: right;"">Pareggiate</th>" & vbCrLf
					Stampa &= "<th style=""text-align: right;"">Perse</th>" & vbCrLf
					Stampa &= "<th style=""text-align: right;"">Goal Fatti</th>" & vbCrLf
					Stampa &= "<th style=""text-align: right;"">Goal Subiti</th>" & vbCrLf
					Stampa &= "</tr>" & vbCrLf

					'idAvversario: +ccampi[0],
					'                        Squadra: ccampi[1],
					'                        Punti: +ccampi[2],
					'                        Giocate: +ccampi[3],
					'                        Vinte: +ccampi[4],
					'                        Pareggiate: +ccampi[5],
					'                        Perse: +ccampi[6],
					'                        gFatti: +ccampi[7],
					'                        gSubiti: +ccampi[8],


					For Each r As String In Righe
						If r <> "" Then
							Dim Campi() As String = r.Split(";")

							Stampa &= "<tr>"
							Dim imm As String = ""

							If Campi(0) = -1 Then
								imm = PathBaseMultimedia & "/" & NomeSquadra & "/Categorie/" & idAnno & "_" & idCategoria & ".kgb"
							Else
								imm = PathBaseMultimedia & "/" & NomeSquadra & "/Avversari/" & Campi(0) & ".kgb"
							End If
							Dim Imm1 As String = DecriptaImmagine(imm)

							Stampa &= "<td><img src=""" & Imm1 & """ width=""50"" height=""50"" /></td>" & vbCrLf
							Stampa &= "<td>" & Campi(1) & "</td>"
							Stampa &= "<td style=""text-align: right;"">" & Campi(2) & "</td>"
							Stampa &= "<td style=""text-align: right;"">" & Campi(3) & "</td>"
							Stampa &= "<td style=""text-align: right;"">" & Campi(4) & "</td>"
							Stampa &= "<td style=""text-align: right;"">" & Campi(5) & "</td>"
							Stampa &= "<td style=""text-align: right;"">" & Campi(6) & "</td>"
							Stampa &= "<td style=""text-align: right;"">" & Campi(7) & "</td>"
							Stampa &= "<td style=""text-align: right;"">" & Campi(8) & "</td>"
							Stampa &= "</tr>"
						End If
					Next

					Stampa &= "</table>"

					Dim PathBaseImmagini As String = pathMultimedia

					Dim filetto As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\statistiche_classifica.txt")
					filetto = filetto.Replace("***SFONDO***", PathBaseImmagini & "/bg.jpg")
					filetto = filetto.Replace("***GIORNATA***", Giornata)
					filetto = filetto.Replace("***CLASSIFICA***", Stampa)

					'Dim Altro As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") ' & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
					Dim NomeFileFinale As String = pathAllegati & Squadra & "\Statistiche\Classifica_Giornata_" & Giornata & ".html"
					Dim NomeFileFinalePDF As String = pathAllegati & Squadra & "\Statistiche\Classifica_Giornata_" & Giornata & ".pdf"
					Dim PathLog As String = HttpContext.Current.Server.MapPath(".") & "\Log\Pdf.txt"
					gf.CreaDirectoryDaPercorso(NomeFileFinale)

					gf.EliminaFileFisico(NomeFileFinale)
					gf.EliminaFileFisico(NomeFileFinalePDF)
					gf.CreaAggiornaFile(NomeFileFinale, filetto)

					Dim pp As New pdfGest
					Ritorno = pp.ConverteHTMLInPDF(NomeFileFinale, NomeFileFinalePDF, PathLog, True)

					If Ritorno = "*" Then
						Ritorno = NomeFileFinalePDF.Replace(pathAllegati, pathMultimedia).Replace("Multimedia", "Allegati").Replace("\", "/")
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function StampaGiornata(Squadra As String, idAnno As String, idCategoria As String, idGiornata As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim Stampa As String = ""
		Dim gf As New GestioneFilesDirectory

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Dim paths As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
				Dim P() As String = paths.Split(";")
				If Strings.Right(P(0), 1) <> "\" Then
					P(0) &= "\"
				End If
				Dim pathAllegati As String = P(0).Replace(vbCrLf, "")
				If Strings.Right(P(2), 1) <> "/" Then
					P(2) &= "/"
				End If
				Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
				Dim PathBaseMultimedia As String = pathMultimedia.Replace("Allegati", "Multimedia")

				Dim NomeSquadra As String = ""
				Dim ss() As String = Squadra.Split("_")
				Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & Val(ss(1)).ToString
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = "Problemi lettura squadra"
				Else
					If Rec.Eof Then
					Else
						NomeSquadra = "" & Rec("Descrizione").Value
					End If
					Rec.Close
				End If

				Dim NomeCategoria As String = ""
				If Ok Then
					Sql = "Select * From Categorie Where idCategoria=" & idCategoria
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ok = False
						Ritorno = "Problemi lettura categoria"
					Else
						If Rec.Eof Then
						Else
							NomeCategoria = "" & Rec("Descrizione").Value & " " & Rec("AnnoCategoria").Value
						End If
						Rec.Close
					End If
				End If

				If Ok Then
					Sql = "Select D.Descrizione As Casa, E.Descrizione As Fuori, A.idSqCasa, A.idSqFuori, F.Risultato " &
					"From CalendarioPartite A " &
					"Left Join SquadreAvversarie D On D.idAvversario = A.idSqCasa " &
					"Left Join SquadreAvversarie E On E.idAvversario = A.idSqFuori " &
					"left Join CalendarioRisultati F On A.idPartita = F.idPartita " &
					"Where idGiornata = " & idGiornata & " And idCategoria = " & idCategoria & " Order By A.idPartita"
					Try
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Stampa &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
							Stampa &= "<tr>" & vbCrLf
							Stampa &= "<th></th>" & vbCrLf
							Stampa &= "<th style=""text-align: right;"">Squadra in Casa</th>" & vbCrLf
							Stampa &= "<th></th>" & vbCrLf
							Stampa &= "<th style=""text-align: right;"">Squadra fuori Casa</th>" & vbCrLf
							Stampa &= "<th style=""text-align: right;"">Risultato</th>" & vbCrLf
							Stampa &= "</tr>" & vbCrLf

							Do Until Rec.eof
								Dim imm As String = ""
								Dim Casa As String = "" & Rec("Casa").Value
								Dim Fuori As String = "" & Rec("Fuori").Value

								If Rec("idSqCasa").Value = -1 Or Rec("idSqCasa").Value = 9999 Then
									imm = PathBaseMultimedia & "/" & NomeSquadra & "/Categorie/" & idAnno & "_" & idCategoria & ".kgb"
									Casa = NomeCategoria
								Else
									imm = PathBaseMultimedia & "/" & NomeSquadra & "/Avversari/" & Rec("idSqCasa").Value & ".kgb"
								End If
								Dim Imm1 As String = DecriptaImmagine(imm)

								If Rec("idSqFuori").Value = -1 Or Rec("idSqFuori").Value = 9999 Then
									imm = PathBaseMultimedia & "/" & NomeSquadra & "/Categorie/" & idAnno & "_" & idCategoria & ".kgb"
									Fuori = NomeCategoria
								Else
									imm = PathBaseMultimedia & "/" & NomeSquadra & "/Avversari/" & Rec("idSqFuori").Value & ".kgb"
								End If
								Dim Imm2 As String = DecriptaImmagine(imm)

								Stampa &= "<td><img src=""" & Imm1 & """ width=""50"" height=""50"" /></td>" & vbCrLf
								Stampa &= "<td>" & Casa & "</td>"
								Stampa &= "<td><img src=""" & Imm2 & """ width=""50"" height=""50"" /></td>" & vbCrLf
								Stampa &= "<td>" & Fuori & "</td>"
								Stampa &= "<td>" & Rec("Risultato").Value & "</td>"
								Stampa &= "</tr>"

								Rec.MoveNext
							Loop
							Rec.Close()

							Stampa &= "</table>"

							Dim PathBaseImmagini As String = pathMultimedia

							Dim filetto As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Scheletri\statistiche_giornata.txt")
							filetto = filetto.Replace("***SFONDO***", PathBaseImmagini & "/bg.jpg")
							filetto = filetto.Replace("***GIORNATA***", idGiornata)
							filetto = filetto.Replace("***CLASSIFICA***", Stampa)

							'Dim Altro As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") ' & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
							Dim NomeFileFinale As String = pathAllegati & Squadra & "\Statistiche\Giornata_" & idGiornata & ".html"
							Dim NomeFileFinalePDF As String = pathAllegati & Squadra & "\Statistiche\Giornata_" & idGiornata & ".pdf"
							Dim PathLog As String = HttpContext.Current.Server.MapPath(".") & "\Log\Pdf.txt"
							gf.CreaDirectoryDaPercorso(NomeFileFinale)

							gf.EliminaFileFisico(NomeFileFinale)
							gf.EliminaFileFisico(NomeFileFinalePDF)
							gf.CreaAggiornaFile(NomeFileFinale, filetto)

							Dim pp As New pdfGest
							Ritorno = pp.ConverteHTMLInPDF(NomeFileFinale, NomeFileFinalePDF, PathLog, True)

							If Ritorno = "*" Then
								Ritorno = NomeFileFinalePDF.Replace(pathAllegati, pathMultimedia).Replace("Multimedia", "Allegati").Replace("\", "/")
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
					End Try
				End If
			End If
		End If

		Return Ritorno
	End Function

End Class