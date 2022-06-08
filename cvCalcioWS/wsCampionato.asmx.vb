Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports ADODB

<System.Web.Services.WebService(Namespace:="http://cvcalcio_cam.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsCampionato
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaCampionatoCategoria(Squadra As String, ByVal idAnno As String, idCategoria As String, idUtente As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Rec2 As Object
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
				Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = StringaErrore & "" & Rec.ToString
					Return Ritorno
				Else
					NomeSquadra = Rec("NomeSquadra").Value
				End If
				Rec.Close()

				Try
					' Squadre avversarie
					Sql = "SELECT AvversariCalendario.idAvversario As idAvv, SquadreAvversarie.Descrizione As Squadra, CampiAvversari.idCampo As idCampo, CampiAvversari.Descrizione As Campo, " &
						"CampiAvversari.Indirizzo As Indirizzo, AvversariCoord.Lat, AvversariCoord.Lon, SquadreAvversarie.FuoriClassifica " &
						"FROM AvversariCalendario LEFT JOIN SquadreAvversarie ON AvversariCalendario.idAvversario = SquadreAvversarie.idAvversario " &
						"Left Join CampiAvversari On SquadreAvversarie.idCampo = CampiAvversari.idCampo " &
						"Left Join AvversariCoord On SquadreAvversarie.idAvversario = AvversariCoord.idAvversario " &
						"WHERE AvversariCalendario.idAnno=" & idAnno & " And AvversariCalendario.idCategoria=" & idCategoria & " " &
						"ORDER BY AvversariCalendario.idProgressivo"
					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = StringaErrore & "" & Rec.ToString
						Return Ritorno
					Else
						' Aggiungo la riga per i dati della categoria
						idSquadre.Add(-idCategoria)
						Squadre.Add(NomeSquadra)

						Ritorno &= "£"

						Do Until Rec.Eof()
							idSquadre.Add(Rec("idAvv").Value)
							Squadre.Add(Rec("Squadra").Value)

							Ritorno &= Rec("idAvv").Value & ";" &
									Rec("Squadra").Value & ";" &
									Rec("idCampo").Value & ";" &
									Rec("Campo").Value.ToString.Replace(";", ",") & ";" &
									Rec("Indirizzo").Value.ToString.Replace(";", ",") & ";" &
									Rec("Lat").Value & ";" &
									Rec("Lon").Value & ";" &
									Rec("FuoriClassifica").Value & ";" &
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
					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = StringaErrore & "" & Rec.ToString
						Return Ritorno
					Else
						Do Until Rec.Eof()
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
							'    Rec2 = Conn.LeggeQuery(Server.MapPath("."),Sql, Connessione)
							'    If TypeOf (Rec2) Is String Then
							'        Ritorno = Rec2
							'        Return Ritorno
							'    Else
							'        Do Until Rec2.Eof()
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
							'    Rec2 = Conn.LeggeQuery(Server.MapPath("."),Sql, Connessione)
							'    If TypeOf (Rec2) Is String Then
							'        Ritorno = StringaErrore & "" & Rec2
							'        Return Ritorno
							'    Else
							'        Do Until Rec2.Eof()
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim ProgressivoSquadra As String = ""

				Dim idSquadre As New ArrayList
				Dim Squadre As New ArrayList
				Dim Giocate As New ArrayList
				Dim Vinte As New ArrayList
				Dim Pareggiate As New ArrayList
				Dim Perse As New ArrayList
				Dim Punti As New ArrayList
				Dim PuntiFC As New ArrayList
				Dim gFatti As New ArrayList
				Dim gSubiti As New ArrayList
				Dim FuoriClassifica As New ArrayList

				Dim CeRisultato As Boolean = False
				Dim g1 As Integer = 0
				Dim g2 As Integer = 0
				Dim gR1 As Integer = 0
				Dim gR2 As Integer = 0

				Dim NomeSquadra As String

				Sql = "Select * From Anni Where idAnno=" & idAnno
				Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = StringaErrore & "" & Rec.ToString
					Return Ritorno
				Else
					NomeSquadra = Rec("NomeSquadra").Value
				End If
				Rec.Close()

				Sql = "Select SquadreAvversarie.idAvversario, SquadreAvversarie.Descrizione, SquadreAvversarie.FuoriClassifica From (AvversariCalendario " &
					"LEFT JOIN SquadreAvversarie As SquadreAvversarie On AvversariCalendario.idAvversario = SquadreAvversarie.idAvversario) " &
					"Where idAnno=" & idAnno & " And idCategoria=" & idCategoria
				Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = StringaErrore & "" & Rec.ToString
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
					PuntiFC.Add(0)
					FuoriClassifica.Add(False)

					Do Until Rec.Eof()
						If Rec("idAvversario").Value <> 999 Then
							idSquadre.Add(Rec("idAvversario").Value)
							Squadre.Add(Rec("Descrizione").Value)
							FuoriClassifica.Add(IIf(Rec("FuoriClassifica").Value = "S", True, False))
							Giocate.Add(0)
							Vinte.Add(0)
							Pareggiate.Add(0)
							Perse.Add(0)
							Punti.Add(0)
							gFatti.Add(0)
							gSubiti.Add(0)
							PuntiFC.Add(0)
						End If

						Rec.MoveNext()
					Loop
				End If
				Rec.Close()

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
				Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = StringaErrore & "" & Rec.ToString
					Return Ritorno
				Else
					Do Until Rec.Eof()
						Dim bFuoriClassifica As Boolean = False

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

						If ControllaValiditaSquadra(idSquadre, FuoriClassifica, Rec("idSqCasa").Value, Rec("idSqFuori").Value) = True Then
							bFuoriClassifica = False
						Else
							bFuoriClassifica = True
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
								If bFuoriClassifica = False Then
									Punti(Indice1) += 3
								Else
									PuntiFC(Indice1) += 3
								End If
							Else
								If g1 < g2 Then
									Vinte(Indice2) += 1
									Perse(Indice1) += 1
									If bFuoriClassifica = False Then
										Punti(Indice2) += 3
									Else
										PuntiFC(Indice2) += 3
									End If
								Else
									Pareggiate(Indice1) += 1
									Pareggiate(Indice2) += 1
									If bFuoriClassifica = False Then
										Punti(Indice1) += 1
										Punti(Indice2) += 1
									Else
										PuntiFC(Indice1) += 1
										PuntiFC(Indice2) += 1
									End If
								End If
							End If
						End If

						Rec.MoveNext()
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

								appo = PuntiFC(i)
								PuntiFC(i) = PuntiFC(k)
								PuntiFC(k) = appo

								appo = Giocate(i)
								Giocate(i) = Giocate(k)
								Giocate(k) = appo

								appo = FuoriClassifica(i)
								FuoriClassifica(i) = FuoriClassifica(k)
								FuoriClassifica(k) = appo

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
						If FuoriClassifica(c) = True Then
							Ritorno &= (500 + PuntiFC(c)) & ";"
						Else
							Ritorno &= Punti(c) & ";"
						End If
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim ProgressivoSquadra As String = ""
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

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
								If TipoDB = "SQLSERVER" Then
									Sql = "SELECT IsNull(Max(idProgressivo),0)+1 FROM AvversariCalendario Where idAnno=" & idAnno & " And idCategoria=" & idCategoria
								Else
									Sql = "SELECT Coalesce(Max(idProgressivo),0)+1 FROM AvversariCalendario Where idAnno=" & idAnno & " And idCategoria=" & idCategoria
								End If
								Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
								If TypeOf (Rec) Is String Then
									Ritorno = Rec
								Else
									If Rec.Eof() Then
										ProgressivoSquadra = "1"
									Else
										'If Rec(0).Value Is DBNull.Value Then
										'	ProgressivoSquadra = "1"
										'Else
										ProgressivoSquadra = Rec(0).Value.ToString
										'End If
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
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Delete From AvversariCalendario Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idAvversario=" & idAvversario
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						If TipoDB = "SQLSERVER" Then
							Sql = "SELECT IsNull(Max(idPartita),0)+1 FROM Partite"
						Else
							Sql = "SELECT Coalesce(Max(idPartita),0)+1 FROM Partite"
						End If
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof() Then
								Ritorno = StringaErrore & " Nessun progressivo partita rilevato"
							Else
								'If Rec(0).Value Is DBNull.Value Then
								'	idNuovaPartita1 = 1
								'Else
								idNuovaPartita1 = Rec(0).Value
								'End If
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					Try
						If TipoDB = "SQLSERVER" Then
							Sql = "SELECT IsNull(Max(idPartita),0)+1 FROM CalendarioPartite"
						Else
							Sql = "SELECT Coalesce(Max(idPartita),0)+1 FROM CalendarioPartite"
						End If
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof() Then
								Ritorno = StringaErrore & " Nessun progressivo partita rilevato"
							Else
								'If Rec(0).Value Is DBNull.Value Then
								'	idNuovaPartita2 = 1
								'Else
								idNuovaPartita2 = Rec(0).Value
								'End If
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
							If TipoDB = "SQLSERVER" Then
								Sql = "SELECT IsNull(Max(idPartita),0)+1 FROM CalendarioPartite Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata
							Else
								Sql = "SELECT Coalesce(Max(idPartita),0)+1 FROM CalendarioPartite Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata
							End If
							Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Nessun progressivo rilevato"
								Else
									'If Rec(0).Value Is DBNull.Value Then
									'	ProgressivoPartita = "1"
									'Else
									ProgressivoPartita = Rec(0).Value.ToString
									'End If
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
						'	Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
						'	If TypeOf (Rec) Is String Then
						'		Ritorno = Rec
						'	Else
						'		If Rec.Eof() Then
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
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
									'    Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
									'Catch ex As Exception
									'    Ritorno = StringaErrore & " " & ex.Message
									'End Try

									Dim Anticipo As Integer = 45

									If Not Ritorno.Contains(StringaErrore & "") Then
										If Val(c(0)) = -1 Or Val(f(0)) = -1 Or Val(c(0)) = 9999 Or Val(f(0)) = 9999 Then
											Try
												Sql = "SELECT AnticipoConvocazione FROM Categorie Where idAnno=" & idAnno & " And idCategoria=" & idCategoria
												Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
												If TypeOf (Rec) Is String Then
													Ritorno = Rec
												Else
													If Not Rec.Eof() Then
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
												Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
												If TypeOf (Rec) Is String Then
													Ritorno = Rec
												Else
													If Rec.Eof() Then
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
													Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
													If TypeOf (Rec) Is String Then
														Ritorno = Rec
													Else
														If Rec.Eof() Then
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
														Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
														Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
														If TypeOf (Rec) Is String Then
															Ritorno = Rec
														Else
															If Rec.Eof() Then
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
															Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
															If TipoDB = "SQLSERVER" Then
																Sql = "Select IsNull(Max(idEvento),0) + 1 From EventiConvocazioni"
															Else
																Sql = "Select Coalesce(Max(idEvento),0) + 1 From EventiConvocazioni"
															End If
															Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
															If TypeOf (Rec) Is String Then
																Ritorno = Rec
															Else
																'If Rec(0).Value Is DBNull.Value Then
																'	idConvocazione = 1
																'Else
																idConvocazione = Rec(0).Value
																'End If
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
															Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
															Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
															If TypeOf (Rec) Is String Then
																Ritorno = Rec
															Else
																If Not Rec.Eof() Then
																	Dim Progressivo As Integer = 0

																	Do Until Rec.Eof()
																		Progressivo += 1
																		Sql = "Insert Into Convocati Values (" &
																			" " & idNuovaPartita & ", " &
																			" " & Progressivo & ", " &
																			" " & Rec("idGiocatore").Value & " " &
																			")"
																		Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
																		If Ritorno.Contains(StringaErrore) Then
																			Ok = False
																			Exit Do
																		End If

																		Rec.MoveNext()
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
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								Catch ex As Exception

								End Try
							End If

							If idNuovaPartita <> -1 Then
								Try
									Sql = "Delete From CalendarioDate Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata & " And idPartita=" & idNuovaPartita
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								Catch ex As Exception
								End Try

								Try
									Sql = "Delete From Partite Where idAnno=" & idAnno & " And idPartita=" & idNuovaPartita & " And idCategoria=" & idCategoria & " " 'And idUnioneCalendario=" & idUnioneCalendario
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
								Catch ex As Exception
								End Try

								Try
									Sql = "Delete From Convocati Where idPartita=" & idNuovaPartita
									Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				' Dim idUnioneCalendario As Integer = -1
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					'Sql = "Select * From Partite Where idAnno=" & idAnno & " And idPartita=" & idPartita
					'Try
					'	Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					'	If TypeOf (Rec) Is String Then
					'		Ritorno = Rec
					'	Else
					'		If Rec.Eof() Then
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
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
								'	Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
								'	If TypeOf (Rec) Is String Then
								'		Ritorno = Rec
								'	Else
								'		If Rec.Eof() Then
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
										Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
											Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
												Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Ok As Boolean = True
				Dim D() As String = Data.Split("/")
				Dim DataSistemata As String = D(2) & "-" & D(1) & "-" & D(0)

				Sql = IIf(TipoDB = "SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

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
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
								"Datella='" & DataSistemata & " " & Ora & "' " &
								"Where idAnno=" & idAnno & " And idCategoria=" & idCategoria & " And idGiornata=" & idGiornata & " And idPartita=" & ProgressivoPartita
							Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
							Catch ex As Exception
							End Try

							Try
								Sql = "Insert Into CalendarioRisultati Values (" & idUnioneCalendario & ", '" & Risultato & "')"
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Ritorno2.Contains(StringaErrore) Then
						Ritorno = Ritorno2
					End If
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
					If Ritorno2.Contains(StringaErrore) Then
						Ritorno = Ritorno2
					End If
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = "Select * From Partite Where idUnioneCalendario=" & idUnioneCalendario

				Try
					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = iif(tipodb="SQLSERVER", "Begin transaction", "Start transaction")
				Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Delete From Giornata Where idUtente=" & idUtente & " And idAnno=" & idAnno & " And idCategoria=" & idCategoria
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					Try
						Sql = "Insert Into Giornata Values (" & idUtente & ", " & idAnno & ", " & idCategoria & ", " & idGiornata & ")"
						Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = "Select * From Giornata Where idUtente=" & idUtente & " And idAnno=" & idAnno & " And idCategoria=" & idCategoria
				Try
					Rec = Conn.LeggeQuery(Server.MapPath("."),  Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							idGiornata = 1
							Try
								Sql = "Insert Into Giornata Values (" & idUtente & ", " & idAnno & ", " & idCategoria & ", " & idGiornata & ")"
								Ritorno = Conn.EsegueSql(Server.MapPath("."), Sql, Connessione)
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
		Dim StampaAmmEsp As String = ""
		Dim StampaSostituzioni As String = ""
		Dim Barra As String = "\"

		If TipoDB <> "SQLSERVER" Then
			Barra = "/"
		End If

		'Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		'Dim P() As String = paths.Split(";")
		'If Strings.Right(P(0), 1) <> "\" Then
		'	P(0) &= "\"
		'End If
		'Dim pathAllegati As String = P(0).Replace(vbCrLf, "")
		'If Strings.Right(P(2), 1) <> "/" Then
		'	P(2) &= "/"
		'End If
		'Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
		'Dim PathBaseMultimedia As String = pathMultimedia.Replace("Allegati", "Multimedia")
		Dim Colore As String = "#ccc"
		Dim wsImm As New wsImmagini

		Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim P() As String = paths.Split(";")
		P(0) = P(0).Trim.Replace(vbCrLf, "")
		If Strings.Right(P(0), 1) <> Barra Then
			P(0) &= Barra
		End If
		Dim pathAllegati As String = P(0).Replace(vbCrLf, "")

		P(1) = P(1).Trim.Replace(vbCrLf, "")
		If Strings.Right(P(1), 1) <> Barra Then
			P(1) &= Barra
		End If
		Dim pathLogG As String = P(1).Replace(vbCrLf, "")

		P(2) = P(2).Trim.Replace(vbCrLf, "")
		If Strings.Right(P(2), 1) <> Barra Then
			P(2) &= Barra
		End If
		Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
		Dim PathBaseImmagini As String = pathMultimedia & "ImmaginiLocali"
		Dim ImmagineSconosciuta As String = PathBaseImmagini & "/Sconosciuto.png"

		ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Stampa statistiche")
		ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "------------------------------------------------------")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Dim NomeSquadra As String = ""
				Dim ss() As String = Squadra.Split("_")
				Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & Val(ss(1)).ToString
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = "Problemi lettura squadra"
					ScriveLog(Server.MapPath("."), Squadra, "Statistiche", Ritorno)
				Else
					If Rec.Eof() Then
					Else
						NomeSquadra = "" & Rec("Descrizione").Value
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Nome Squadra: " & NomeSquadra)
					End If
					Rec.Close()
				End If

				Dim filetto As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_Goals.txt")

				If Ok Then
					ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sui goals")

					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%' Order By Totale Desc, GoalCampionato Desc, GoalAmichevole Desc"
					Sql = ConverteStringaSQL(Sql)

					Try
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione, False)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sui goals. " & Ritorno)
						Else
							If Stampa = "S" Then
								StampaMarcatori &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaMarcatori &= "<tr style=""background-color: gray;"">" & vbCrLf
								StampaMarcatori &= "<th></th>" & vbCrLf
								StampaMarcatori &= "<th>Nominativo</th>" & vbCrLf
								StampaMarcatori &= "<th>Ruolo</th>" & vbCrLf
								StampaMarcatori &= "<th>Goal Amichevole</th>" & vbCrLf
								StampaMarcatori &= "<th>Goal Campionato</th>" & vbCrLf
								StampaMarcatori &= "<th>Rigori</th>" & vbCrLf
								StampaMarcatori &= "<th>Totale</th>" & vbCrLf
								StampaMarcatori &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof()
								Ritorno &= Rec("Cognome").Value & ";"
								Ritorno &= Rec("Nome").Value & ";"
								Ritorno &= Rec("Soprannome").Value & ";"
								Ritorno &= Rec("Ruolo").Value & ";"
								Ritorno &= Rec("GoalAmichevole").Value & ";"
								Ritorno &= Rec("GoalCampionato").Value & ";"
								Ritorno &= Rec("Rigori").Value & ";"
								Ritorno &= Rec("Totale").Value & ";"
								Ritorno &= Rec("idGiocatore").Value & ";"
								Ritorno &= "§"

								If Stampa = "S" Then
									Dim Soprannome As String = Rec("Soprannome").Value
									If Soprannome <> "" Then Soprannome = "'" & Soprannome & "' "
									Dim Nominativo As String = Rec("Nome").Value & " " & Soprannome & Rec("Cognome").Value
									' Dim Path As String = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
									Dim Path As String = wsImm.RitornaImmagineDB(Squadra, "Giocatori", Rec("idGiocatore").Value, "")
									If Path.Contains(StringaErrore) Then
										Path = ImmagineSconosciuta
									Else
										Path = "data:image/png;base64,'" + Path
									End If

									' Path = DecriptaImmagine(Server.MapPath("."), Path)

									StampaMarcatori &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
									If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
									StampaMarcatori &= "<td><img src=""" & Path & """ width=""50"" height=""50"" /></td>" & vbCrLf
									StampaMarcatori &= "<td>" & Nominativo & "</td>" & vbCrLf
									StampaMarcatori &= "<td>" & Rec("Ruolo").Value & "</td>" & vbCrLf
									StampaMarcatori &= "<td style=""text-align: right;"">" & Rec("GoalAmichevole").Value & "</td>" & vbCrLf
									StampaMarcatori &= "<td style=""text-align: right;"">" & Rec("GoalCampionato").Value & "</td>" & vbCrLf
									StampaMarcatori &= "<td style=""text-align: right;"">" & Rec("Rigori").Value & "</td>" & vbCrLf
									StampaMarcatori &= "<td style=""text-align: right;"">" & Rec("Totale").Value & "</td>" & vbCrLf
									StampaMarcatori &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext()
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaMarcatori &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						Ok = False
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sui goals. CATCH: " & Ritorno)
					End Try
				End If

				Dim RitornoSost As String = ""

				If Ok Then
					ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle sostituzioni")
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_Sostituzioni.txt")
					Sql = "Select idGiocatore, Cosa, Cognome, Nome, Soprannome, Ruolo, Volte From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%' Order By Cosa, Volte Desc"
					Sql = ConverteStringaSQL(Sql)

					Try
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle sostituzioni: " & Ritorno)
						Else
							If Stampa = "S" Then
								StampaSostituzioni &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaSostituzioni &= "<tr style=""background-color: gray;"">" & vbCrLf
								StampaSostituzioni &= "<th></th>" & vbCrLf
								StampaSostituzioni &= "<th>Nominativo</th>" & vbCrLf
								StampaSostituzioni &= "<th>Ruolo</th>" & vbCrLf
								StampaSostituzioni &= "<th>Volte</th>" & vbCrLf
								StampaSostituzioni &= "<th>Cosa</th>" & vbCrLf
								StampaSostituzioni &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof()
								RitornoSost &= Rec("Cognome").Value & ";"
								RitornoSost &= Rec("Nome").Value & ";"
								RitornoSost &= Rec("Soprannome").Value & ";"
								RitornoSost &= Rec("Ruolo").Value & ";"
								RitornoSost &= Rec("Volte").Value & ";"
								RitornoSost &= Rec("idGiocatore").Value & ";"
								RitornoSost &= Rec("Cosa").Value & ";"
								RitornoSost &= "§"

								If Stampa = "S" Then
									Dim Soprannome As String = Rec("Soprannome").Value
									If Soprannome <> "" Then Soprannome = "'" & Soprannome & "' "
									Dim Nominativo As String = Rec("Nome").Value & " " & Soprannome & Rec("Cognome").Value
									'Dim Path As String = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
									'Path = DecriptaImmagine(Server.MapPath("."), Path)
									Dim Path As String = wsImm.RitornaImmagineDB(Squadra, "Giocatori", Rec("idGiocatore").Value, "")
									If Path.Contains(StringaErrore) Then
										Path = ImmagineSconosciuta
									Else
										Path = "data:image/png;base64,'" + Path
									End If

									StampaSostituzioni &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
									If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
									StampaSostituzioni &= "<td><img src=""" & Path & """ width=""50"" height=""50"" /></td>" & vbCrLf
									StampaSostituzioni &= "<td>" & Nominativo & "</td>" & vbCrLf
									StampaSostituzioni &= "<td>" & Rec("Ruolo").Value & "</td>" & vbCrLf
									StampaSostituzioni &= "<td style=""text-align: right;"">" & Rec("Volte").Value & "</td>" & vbCrLf
									StampaSostituzioni &= "<td>" & Rec("Cosa").Value & "</td>" & vbCrLf
									StampaSostituzioni &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext()
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaSostituzioni &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle sostituzioni. CATCH: " & Ritorno)
						Ok = False
					End Try
				End If

				If Ok Then
					ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle presenze")
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_Presenze.txt")
					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%' Order By Totale Desc, PresenzeCampionato Desc, PresenzeAmichevole Desc"
					Sql = ConverteStringaSQL(Sql)

					Try
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle presenze: " & Ritorno)
						Else
							Ritorno &= "|"

							If Stampa = "S" Then
								StampaPresenze &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaPresenze &= "<tr style=""background-color: gray;"">" & vbCrLf
								StampaPresenze &= "<th></th>" & vbCrLf
								StampaPresenze &= "<th>Nominativo</th>" & vbCrLf
								StampaPresenze &= "<th>Ruolo</th>" & vbCrLf
								StampaPresenze &= "<th>Presenze Amichevole</th>" & vbCrLf
								StampaPresenze &= "<th>Presenze Campionato</th>" & vbCrLf
								StampaPresenze &= "<th>Totale</th>" & vbCrLf
								StampaPresenze &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof()
								Ritorno &= Rec("Cognome").Value & ";"
								Ritorno &= Rec("Nome").Value & ";"
								Ritorno &= Rec("Soprannome").Value & ";"
								Ritorno &= Rec("Ruolo").Value & ";"
								Ritorno &= Rec("PresenzeCampionato").Value & ";"
								Ritorno &= Rec("PresenzeAmichevole").Value & ";"
								Ritorno &= Rec("Totale").Value & ";"
								Ritorno &= Rec("idGiocatore").Value & ";"
								Ritorno &= "§"

								If Stampa = "S" Then
									Dim Soprannome As String = Rec("Soprannome").Value
									If Soprannome <> "" Then Soprannome = "'" & Soprannome & "' "
									Dim Nominativo As String = Rec("Nome").Value & " " & Soprannome & Rec("Cognome").Value
									'Dim Path As String = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
									'Path = DecriptaImmagine(Server.MapPath("."), Path)
									Dim Path As String = wsImm.RitornaImmagineDB(Squadra, "Giocatori", Rec("idGiocatore").Value, "")
									If Path.Contains(StringaErrore) Then
										Path = ImmagineSconosciuta
									Else
										Path = "data:image/png;base64,'" + Path
									End If

									StampaPresenze &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
									If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
									StampaPresenze &= "<td><img src=""" & Path & """ width=""50"" height=""50"" /></td>" & vbCrLf
									StampaPresenze &= "<td>" & Nominativo & "</td>" & vbCrLf
									StampaPresenze &= "<td>" & Rec("Ruolo").Value & "</td>" & vbCrLf
									StampaPresenze &= "<td style=""text-align: right;"">" & Rec("PresenzeCampionato").Value & "</td>" & vbCrLf
									StampaPresenze &= "<td style=""text-align: right;"">" & Rec("PresenzeAmichevole").Value & "</td>" & vbCrLf
									StampaPresenze &= "<td style=""text-align: right;"">" & Rec("Totale").Value & "</td>" & vbCrLf
									StampaPresenze &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext()
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaPresenze &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle presenze: CATCH: " & Ritorno)
						Ok = False
					End Try
				End If

				If Ok Then
					ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle fasce goal fatti")
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_FasceGoalFatti.txt")
					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%'"
					Sql = ConverteStringaSQL(Sql)

					Try
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle fasce goal fatti: " & Ritorno)
						Else
							Ritorno &= "|"

							If Stampa = "S" Then
								StampaFgF &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaFgF &= "<tr style=""background-color: gray;"">" & vbCrLf
								StampaFgF &= "<th>Tipologia</th>" & vbCrLf
								StampaFgF &= "<th>Fascia</th>" & vbCrLf
								StampaFgF &= "<th style=""text-align: right;""> Tempo</th>" & vbCrLf
								StampaFgF &= "<th style=""text-align: right;"">Goals</th>" & vbCrLf
								StampaFgF &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof()
								Ritorno &= Rec("Tipologia").Value & ";"
								Ritorno &= Rec("Fascia").Value & ";"
								Ritorno &= Rec("idTempo").Value & ";"
								Ritorno &= Rec("Goals").Value & ";"
								Ritorno &= "§"

								If Stampa = "S" Then
									StampaFgF &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
									If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
									StampaFgF &= "<td>" & Rec("Tipologia").Value & "</td>" & vbCrLf
									StampaFgF &= "<td>" & Rec("Fascia").Value & "</td>" & vbCrLf
									StampaFgF &= "<td style=""text-align: right;"">" & Rec("idTempo").Value & "</td>" & vbCrLf
									StampaFgF &= "<td style=""text-align: right;"">" & Rec("Goals").Value & "</td>" & vbCrLf
									StampaFgF &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext()
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaFgF &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle fasce goal fatti. CATCH: " & Ritorno)
						Ok = False
					End Try
				End If

				If Ok Then
					ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle fasce goal subiti")

					Sql = "Select C.idTipologia, idTempo, Minuti " &
						"From RisultatiAvversariMinuti A " &
						"Left Join Partite B On A.idPartita = B.idPartita " &
						"Left Join [Generale].[dbo].Tipologie C On B.idTipologia = C.idTipologia " &
						"Left Join Convocati E On A.idPartita = E.idPartita And E.idPartita = B.idPartita " &
						"Left Join Giocatori D On E.idGiocatore = D.idGiocatore And E.idProgressivo = 1 " &
						"Where E.idProgressivo = 1 And Categorie Like '%" & idCategoria & "-%'"
					Try
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle fasce goal subiti: " & Ritorno)
						Else
							Dim Fascia1(1, 2) As Integer
							Dim Fascia2(1, 2) As Integer
							Dim Fascia3(1, 2) As Integer
							Dim Fascia4(1, 2) As Integer
							Dim Fascia5(1, 2) As Integer

							Do Until Rec.Eof()
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
								Rec.MoveNext()
							Loop
							Rec.Close()

							Ritorno &= "|"

							Dim Tipi() As String = {"Campionato", "Amichevole"}
							Dim Fascia() As String = {"0-9", "10-19", "20-29", "30-39", "40-"}

							If Stampa = "S" Then
								StampaFgs &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaFgs &= "<tr style=""background-color: gray;"">" & vbCrLf
								StampaFgs &= "<th>Tipologia</th>" & vbCrLf
								StampaFgs &= "<th>Fascia</th>" & vbCrLf
								StampaFgs &= "<th style=""text-align: right;"">Tempo</th>" & vbCrLf
								StampaFgs &= "<th style=""text-align: right;"">Goals</th>" & vbCrLf
								StampaFgs &= "</tr>" & vbCrLf
							End If

							For i As Integer = 0 To 1
								For k As Integer = 0 To 2
									If Fascia1(i, k) > 0 Then
										Ritorno &= Tipi(i) & ";" & Fascia(k) & ";" & k + 1 & ";" & Fascia1(i, k) & "§"

										If Stampa = "S" Then
											StampaFgs &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
											If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
											StampaFgs &= "<td>" & Tipi(i) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia(k) & "</td>" & vbCrLf
											StampaFgs &= "<td style=""text-align: right;"">" & k + 1 & "</td>" & vbCrLf
											StampaFgs &= "<td style=""text-align: right;"">" & Fascia1(i, k) & "</td>" & vbCrLf
											StampaFgs &= "</tr>" & vbCrLf
										End If
									End If
									If Fascia2(i, k) > 0 Then
										Ritorno &= Tipi(i) & ";" & Fascia(k) & ";" & k + 1 & ";" & Fascia2(i, k) & "§"

										If Stampa = "S" Then
											StampaFgs &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
											If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
											StampaFgs &= "<td>" & Tipi(i) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia(k) & "</td>" & vbCrLf
											StampaFgs &= "<td style=""text-align: right;"">" & k + 1 & "</td>" & vbCrLf
											StampaFgs &= "<td style=""text-align: right;"">" & Fascia2(i, k) & "</td>" & vbCrLf
											StampaFgs &= "</tr>" & vbCrLf
										End If
									End If
									If Fascia3(i, k) > 0 Then
										Ritorno &= Tipi(i) & ";" & Fascia(k) & ";" & k + 1 & ";" & Fascia3(i, k) & "§"

										If Stampa = "S" Then
											StampaFgs &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
											If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
											StampaFgs &= "<td>" & Tipi(i) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia(k) & "</td>" & vbCrLf
											StampaFgs &= "<td style=""text-align: right;"">" & k + 1 & "</td>" & vbCrLf
											StampaFgs &= "<td style=""text-align: right;"">" & Fascia3(i, k) & "</td>" & vbCrLf
											StampaFgs &= "</tr>" & vbCrLf
										End If
									End If
									If Fascia4(i, k) > 0 Then
										Ritorno &= Tipi(i) & ";" & Fascia(k) & ";" & k + 1 & ";" & Fascia4(i, k) & "§"

										If Stampa = "S" Then
											StampaFgs &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
											If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
											StampaFgs &= "<td>" & Tipi(i) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia(k) & "</td>" & vbCrLf
											StampaFgs &= "<td style=""text-align: right;"">" & k + 1 & "</td>" & vbCrLf
											StampaFgs &= "<td style=""text-align: right;"">" & Fascia4(i, k) & "</td>" & vbCrLf
											StampaFgs &= "</tr>" & vbCrLf
										End If
									End If
									If Fascia5(i, k) > 0 Then
										Ritorno &= Tipi(i) & ";" & Fascia(k) & ";" & k + 1 & ";" & Fascia5(i, k) & "§"

										If Stampa = "S" Then
											StampaFgs &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
											If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
											StampaFgs &= "<td>" & Tipi(i) & "</td>" & vbCrLf
											StampaFgs &= "<td>" & Fascia(k) & "</td>" & vbCrLf
											StampaFgs &= "<td style=""text-align: right;"">" & k + 1 & "</td>" & vbCrLf
											StampaFgs &= "<td style=""text-align: right;"">" & Fascia5(i, k) & "</td>" & vbCrLf
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
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle fasce goal subiti. CATCH: " & Ritorno)
						Ok = False
					End Try
				End If

				If Ok Then
					ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sugli eventi")
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_Eventi.txt")
					Sql = "Select " & IIf(TipoDB = "SQLSERVER", "Top 25", "") & " * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%' Order By Quanti Desc" & IIf(TipoDB = "SQLSERVER", "", " Limit 25")
					Sql = ConverteStringaSQL(Sql)

					Try
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sugli eventi: " & Ritorno)
						Else
							Ritorno &= "|"

							If Stampa = "S" Then
								StampaEventi &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaEventi &= "<tr style=""background-color: gray;"">" & vbCrLf
								StampaEventi &= "<th></th>" & vbCrLf
								StampaEventi &= "<th>Nominativo</th>" & vbCrLf
								StampaEventi &= "<th>Descrizione</th>" & vbCrLf
								StampaEventi &= "<th style=""text-align: right;"">Quanti</th>" & vbCrLf
								StampaEventi &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof()
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
									'Dim Path As String = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
									'Path = DecriptaImmagine(Server.MapPath("."), Path)
									Dim Path As String = wsImm.RitornaImmagineDB(Squadra, "Giocatori", Rec("idGiocatore").Value, "")
									If Path.Contains(StringaErrore) Then
										Path = ImmagineSconosciuta
									Else
										Path = "data:image/png;base64,'" + Path
									End If

									StampaEventi &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
									If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
									StampaEventi &= "<td><img src=""" & Path & """ width=""50"" height=""50"" /></td>" & vbCrLf
									StampaEventi &= "<td>" & Nominativo & "</td>" & vbCrLf
									StampaEventi &= "<td>" & Rec("Descrizione").Value & "</td>" & vbCrLf
									StampaEventi &= "<td style=""text-align: right;"">" & Rec("Quanti").Value & "</td>" & vbCrLf
									StampaEventi &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext()
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaEventi &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sugli eventi. CATCH: " & Ritorno)
						Ok = False
					End Try
				End If

				If Ok Then
					ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle tipologie delle partite")
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_TipologiePartite.txt")
					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%'"
					Sql = ConverteStringaSQL(Sql)

					Try
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle tipologie delle partite: " & Ritorno)
						Else
							Ritorno &= "|"

							If Stampa = "S" Then
								StampaTipoPartite &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaTipoPartite &= "<tr style=""background-color: gray;"">" & vbCrLf
								StampaTipoPartite &= "<th>Dove</th>" & vbCrLf
								StampaTipoPartite &= "<th>Descrizione</th>" & vbCrLf
								StampaTipoPartite &= "<th style=""text-align: right;"">Quante</th>" & vbCrLf
								StampaTipoPartite &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof()
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
									StampaTipoPartite &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
									If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
									StampaTipoPartite &= "<td>" & Dove & "</td>" & vbCrLf
									StampaTipoPartite &= "<td>" & Rec("Descrizione").Value & "</td>" & vbCrLf
									StampaTipoPartite &= "<td style=""text-align: right;"">" & Rec("Quante").Value & "</td>" & vbCrLf
									StampaTipoPartite &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext()
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaTipoPartite &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle tipologie delle partite. CATCH: " & Ritorno)
						Ok = False
					End Try
				End If

				If Ok Then
					ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite")
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_Partite.txt")
					filetto = filetto.Replace("%idCategoria%", idCategoria)
					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%'"
					Sql = ConverteStringaSQL(Sql)

					Try
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite: " & Ritorno)
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

							Do Until Rec.Eof()
								Dim Avversario As String = Rec("Avversario").Value
								If Avversario <> "" Then
									Dim Categoria As String = Rec("Categoria").Value
									Dim Risultato As String = Rec("Risultato").Value
									Dim GCasa As Integer = Val(Rec("Casa").Value)
									Dim GFuori As Integer = Val(Rec("Fuori").Value)
									ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite: Avversario: " & Avversario & " - Categoria " & Categoria & " - Risultato " & Risultato & " - GCasa " & GCasa & " - GFuori" & GFuori)

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
									ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite. Avversari: " & Avversari.Count)

									Dim sqCasa As String = ""
									Dim sqFuori As String = ""
									Dim goalCasa As Integer
									Dim goalFuori As Integer

									Dim Dove As String = ""
									Dim TotaleCasa As Integer = 0
									Dim TotaleFuori As Integer = 0

									Giocate += 1

									ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite: Dove: " & Rec("Dove").Value)
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
									ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite: Totale " & Totale & " - GCasa " & GCasa & " - GFuori " & GFuori)

									If Totale > MaxGoals Then
										GoalMax = "Max goal totali: " & sqCasa & "-" & sqFuori & " " & goalCasa & "-" & goalFuori & " -> Totale: " & Totale
										MaxGoals = Totale
									End If
								End If

								Rec.MoveNext()
							Loop
							Rec.Close()

							Ritorno &= "|"

							If Stampa = "S" Then
								StampaStatistiche &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaStatistiche &= "<tr style=""background-color: gray;"">" & vbCrLf
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

							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite: Ritorno 1")

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

							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite: Ritorno 2")

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

							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite: Ritorno 3")

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

							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite: Ritorno 4")

							If Stampa = "S" Then
								For Each l As String In Lista
									StampaStatistiche &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
									If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
									StampaStatistiche &= "<td>" & l & "</dh>" & vbCrLf
									StampaStatistiche &= "</tr>" & vbCrLf
								Next
								StampaStatistiche &= "</table>"
							End If

							Ritorno &= "|"

							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite: Ritorno 5")

							If Stampa = "S" Then
								StampaAvversari &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaAvversari &= "<tr style=""background-color: gray;"">" & vbCrLf
								StampaAvversari &= "<th></th>" & vbCrLf
								StampaAvversari &= "<th>Avversario</th>" & vbCrLf
								StampaAvversari &= "</tr>" & vbCrLf
							End If

							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite: Incontrati: " & Incontrati.Count - 1)

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

							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite: Avversari: " & Avversari.Count - 1)

							For Each a As String In Avversari
								Ritorno &= idAvversario.Item(iii) & ";" & a & ";" & Incontrati.Item(iii) & "§"

								If Stampa = "S" Then
									'Dim Imm2 As String = PathBaseMultimedia & "/" & NomeSquadra & "/Avversari/" & idAvversario.Item(iii) & ".kgb"
									'Imm2 = DecriptaImmagine(Server.MapPath("."), Imm2)
									Dim Path As String = wsImm.RitornaImmagineDB(Squadra, "Avversari", idAvversario.Item(iii), "")
									If Path.Contains(StringaErrore) Then
										Path = ImmagineSconosciuta
									Else
										Path = "data:image/png;base64,'" + Path
									End If

									StampaAvversari &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
									If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
									StampaAvversari &= "<td><img src=""" & Path & """ width=""50"" height=""50"" /></td>" & vbCrLf
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
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sulle partite. CATCH: " & Ritorno)
						Ok = False
					End Try
				End If

				If Ok Then
					ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sul meteo")
					filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Queries\Statistiche_Meteo.txt")
					Sql = "Select * From (" & filetto & ") As B Where Categorie Like '%" & idCategoria & "-%' Order By Quante Desc, Tempo"
					Sql = ConverteStringaSQL(Sql)

					Try
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
							ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sul meteo: " & Ritorno)
						Else
							Ritorno &= "|"

							If Stampa = "S" Then
								StampaMeteo &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
								StampaMeteo &= "<tr style=""background-color: gray;"">" & vbCrLf
								StampaMeteo &= "<th></th>" & vbCrLf
								StampaMeteo &= "<th>Tempo</th>" & vbCrLf
								StampaMeteo &= "<th style=""text-align: right;"">Quante</th>" & vbCrLf
								StampaMeteo &= "</tr>" & vbCrLf
							End If

							Do Until Rec.Eof()
								Ritorno &= Rec("Tempo").Value & ";"
								Ritorno &= Rec("Icona").Value & ";"
								Ritorno &= Rec("Quante").Value & ";"
								Ritorno &= "§"

								If Stampa = "S" Then
									StampaMeteo &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
									If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
									StampaMeteo &= "<td><img src=""" & Rec("Icona").Value & """ width=""50"" height=""50"" /></td>" & vbCrLf
									StampaMeteo &= "<td>" & Rec("Tempo").Value & "</td>" & vbCrLf
									StampaMeteo &= "<td style=""text-align: right;"">" & Rec("Quante").Value & "</td>" & vbCrLf
									StampaMeteo &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext()
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaMeteo &= "</table>"
							End If
						End If
					Catch ex As Exception
						Ritorno = "ERROR: " & ex.Message
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Calcolo statistiche sul meteo. CATCH: " & Ritorno)
						Ok = False
					End Try
				End If

				If Ok Then
					ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Ritorno statistiche")

					Dim AmmoEspu As String = ""

					If Stampa = "S" Then
						StampaAmmEsp &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
						StampaAmmEsp &= "<tr style=""background-color: gray;"">" & vbCrLf
						StampaAmmEsp &= "<th>Tipo Partita</th>" & vbCrLf
						StampaAmmEsp &= "<th>Tipologia</th>" & vbCrLf
						StampaAmmEsp &= "<th></th>" & vbCrLf
						StampaAmmEsp &= "<th>Nominativo</th>" & vbCrLf
						StampaAmmEsp &= "<th style=""text-align: right;"">Quante</th>" & vbCrLf
						StampaAmmEsp &= "</tr>" & vbCrLf
					End If

					Sql = "Select * From (" &
						"Select Cognome, Nome, Soprannome, C.Descrizione As AmmoEspu, A.idGiocatore, B.Categorie, E.Descrizione As Tipologia, " & IIf(TipoDB = "SQLSERVER", "Isnull(Count(*),0)", "COALESCE(Count(*),0)") & " As Quante From EventiPartita A " &
						"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
						"Left Join Eventi C On A.idEvento = C.idEvento " &
						"Left Join Partite D On D.idPartita = A.idPartita " &
						"Left Join [Generale].[dbo].TipologiePartite E On D.idTipologia = E.idTipologia " &
						"Where (C.Descrizione = 'Ammonito' Or C.Descrizione = 'Espulso') And  Categorie Like '%" & idCategoria & "-%' " &
						"Group By Cognome, Nome, Soprannome, C.Descrizione, A.idGiocatore, B.Categorie, E.Descrizione" &
						") As A Order By Tipologia, Quante Desc"
					Try
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Do Until Rec.Eof()
								AmmoEspu &= Rec("Cognome").Value & ";"
								AmmoEspu &= Rec("Nome").Value & ";"
								AmmoEspu &= Rec("Soprannome").Value & ";"
								AmmoEspu &= Rec("AmmoEspu").Value & ";"
								AmmoEspu &= Rec("idGiocatore").Value & ";"
								AmmoEspu &= Rec("Tipologia").Value & ";"
								AmmoEspu &= Rec("Quante").Value & "§"

								If Stampa = "S" Then
									'Dim Path As String = PathBaseMultimedia & "/" & NomeSquadra & "/Giocatori/" & idAnno & "_" & Rec("idGiocatore").Value & ".kgb"
									'Path = DecriptaImmagine(Server.MapPath("."), Path)
									Dim Path As String = wsImm.RitornaImmagineDB(Squadra, "Giocatori", Rec("idGiocatore").Value, "")
									If Path.Contains(StringaErrore) Then
										Path = ImmagineSconosciuta
									Else
										Path = "data:image/png;base64,'" + Path
									End If

									StampaAmmEsp &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
									If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
									StampaAmmEsp &= "<td>" & Rec("Tipologia").Value & "</td>" & vbCrLf
									StampaAmmEsp &= "<td>" & Rec("AmmoEspu").Value & "</td>" & vbCrLf
									StampaAmmEsp &= "<td><img src=""" & Path & """ width=""50"" height=""50"" /></td>" & vbCrLf
									StampaAmmEsp &= "<td>" & Rec("Nome").Value & " '" & Rec("Soprannome").Value & "' " & Rec("Cognome").Value & "</td>" & vbCrLf
									StampaAmmEsp &= "<td style=""text-align: right;"">" & Rec("Quante").Value & "</td>" & vbCrLf
									StampaAmmEsp &= "</tr>" & vbCrLf
								End If

								Rec.MoveNext()
							Loop
							Rec.Close()

							If Stampa = "S" Then
								StampaAmmEsp &= "</table>"
							End If

							Ritorno &= "|" & AmmoEspu
						End If
					Catch ex As Exception
						Ok = False
						Ritorno = ex.Message
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Ritorno statistiche. CATCH: " & Ritorno)
					End Try
				End If

				Ritorno &= "|" & RitornoSost

				If Ok Then
					If Stampa = "S" Then
						'Dim PathBaseImmagini As String = pathMultimedia

						filetto = gf.LeggeFileIntero(Server.MapPath(".") & "\Scheletri\statistiche_nuove.txt")
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
						filetto = filetto.Replace("***AMMESP***", StampaAmmEsp)
						filetto = filetto.Replace("***SOSTITUZIONI***", StampaSostituzioni)

						'Dim Altro As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") ' & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
						Dim NomeFileFinale As String = pathAllegati & Squadra & "\Statistiche\Statistiche_Giornata_" & idGiornata & ".html"
						Dim NomeFileFinalePDF As String = pathAllegati & Squadra & "\Statistiche\Statistiche_Giornata_" & idGiornata & ".pdf"
						Dim NomeFileLog As String = pathAllegati & Squadra & "\Statistiche\LogPDFStatistiche_" & idGiornata & ".txt"
						'Dim PathLog As String = Server.MapPath(".") & "\Log\Pdf.txt"
						gf.CreaDirectoryDaPercorso(NomeFileFinale)

						gf.EliminaFileFisico(NomeFileFinale)
						gf.EliminaFileFisico(NomeFileFinalePDF)
						gf.CreaAggiornaFile(NomeFileFinale, filetto)

						'Dim pp As New pdfGest
						'Ritorno = pp.ConverteHTMLInPDF(NomeFileFinale, NomeFileFinalePDF, PathLog, True)

						' Ritorno = ConvertePDF(NomeFileFinale, NomeFileFinalePDF)
						Dim pp As New pdfGest
						Ritorno = pp.ConverteHTMLInPDF(NomeFileFinale, NomeFileFinalePDF, NomeFileLog)

						If Ritorno = "*" Then
							Ritorno = NomeFileFinalePDF '.Replace(pathAllegati, pathMultimedia).Replace("Multimedia", "Allegati").Replace("\", "/")
							gf.EliminaFileFisico(NomeFileFinale)
							gf.EliminaFileFisico(NomeFileLog)
						End If
						ScriveLog(Server.MapPath("."), Squadra, "Statistiche", "Ritorno statistiche 1: " & Ritorno)

					End If
				End If
			End If
		End If

		' Marcatori | Presenze | Fasce Goal Fatti | Fasce Goal Subiti | Eventi | Tipologie Partite | Partite | Avversari Incontrati | Meteo | AmmoEspu | Sostituzioni

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function StampaClassifica(Squadra As String, Classifica As String, Giornata As String, idAnno As String, idCategoria As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As String = ""
		Dim Righe() As String = Classifica.Split("§")
		Dim Stampa As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim Barra As String = "\"

		If TipoDB <> "SQLSERVER" Then
			Barra = "/"
		End If

		'Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		'Dim P() As String = paths.Split(";")
		'If Strings.Right(P(0), 1) <> "\" Then
		'	P(0) &= "\"
		'End If
		'Dim pathAllegati As String = P(0).Replace(vbCrLf, "")
		'If Strings.Right(P(2), 1) <> "/" Then
		'	P(2) &= "/"
		'End If
		'Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
		'Dim PathBaseMultimedia As String = pathMultimedia.Replace("Allegati", "Multimedia")

		Dim wsImm As New wsImmagini

		Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		Dim P() As String = paths.Split(";")
		P(0) = P(0).Trim.Replace(vbCrLf, "")
		If Strings.Right(P(0), 1) <> Barra Then
			P(0) &= Barra
		End If
		Dim pathAllegati As String = P(0).Replace(vbCrLf, "")

		P(1) = P(1).Trim.Replace(vbCrLf, "")
		If Strings.Right(P(1), 1) <> Barra Then
			P(1) &= Barra
		End If
		Dim pathLogG As String = P(1).Replace(vbCrLf, "")

		P(2) = P(2).Trim.Replace(vbCrLf, "")
		If Strings.Right(P(2), 1) <> Barra Then
			P(2) &= Barra
		End If
		Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
		Dim PathBaseImmagini As String = pathMultimedia & "ImmaginiLocali"
		Dim ImmagineSconosciuta As String = PathBaseImmagini & "/Sconosciuto.png"

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Ok As Boolean = True
				Dim Colore As String = "#ccc"

				Dim NomeSquadra As String = ""
				Dim ss() As String = Squadra.Split("_")
				Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & Val(ss(1)).ToString
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = "Problemi lettura squadra"
				Else
					If Rec.Eof() Then
					Else
						NomeSquadra = "" & Rec("Descrizione").Value
					End If
					Rec.Close()
				End If

				If Ok Then
					Stampa &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
					Stampa &= "<tr style=""background-color: gray;"">" & vbCrLf
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

							Stampa &= "<tr style=""background-color: " & Colore & ";"">"
							If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"

							Dim imm As String = ""

							If Campi(0) = -1 Then
								'imm = PathBaseMultimedia & "/" & NomeSquadra & "/Categorie/" & idAnno & "_" & idCategoria & ".kgb"

								imm = wsImm.RitornaImmagineDB(Squadra, "Categorie", idCategoria, "")
								If imm.Contains(StringaErrore) Then
									imm = ImmagineSconosciuta
								Else
									imm = "data:image/png;base64,'" + imm
								End If
							Else
								'imm = PathBaseMultimedia & "/" & NomeSquadra & "/Avversari/" & Campi(0) & ".kgb"
								imm = wsImm.RitornaImmagineDB(Squadra, "Avversari", Campi(0), "")
								If imm.Contains(StringaErrore) Then
									imm = ImmagineSconosciuta
								Else
									imm = "data:image/png;base64,'" + imm
								End If
							End If
							'Dim Imm1 As String = DecriptaImmagine(Server.MapPath("."), imm)

							Stampa &= "<td><img src=""" & imm & """ width=""50"" height=""50"" /></td>" & vbCrLf
							Stampa &= "<td>" & Campi(1) & "</td>"
							If Campi(2) > 499 Then
								Stampa &= "<td style=""text-align: right;"">*" & Campi(2) - 500 & "*</td>"
							Else
								Stampa &= "<td style=""text-align: right;"">" & Campi(2) & "</td>"
							End If
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

					'Dim PathBaseImmagini As String = pathMultimedia

					Dim filetto As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Scheletri\statistiche_classifica.txt")
					filetto = filetto.Replace("***SFONDO***", PathBaseImmagini & "/bg.jpg")
					filetto = filetto.Replace("***GIORNATA***", Giornata)
					filetto = filetto.Replace("***CLASSIFICA***", Stampa)

					'Dim Altro As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") ' & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
					Dim NomeFileFinale As String = pathAllegati & Squadra & "\Statistiche\Classifica_Giornata_" & Giornata & ".html"
					Dim NomeFileFinalePDF As String = pathAllegati & Squadra & "\Statistiche\Classifica_Giornata_" & Giornata & ".pdf"
					Dim NomeFileLog As String = pathAllegati & Squadra & "\Statistiche\LogPDFClassifica_" & Giornata & ".pdf"
					' Dim PathLog As String = Server.MapPath(".") & "\Log\Pdf.txt"
					gf.CreaDirectoryDaPercorso(NomeFileFinale)

					gf.EliminaFileFisico(NomeFileFinale)
					gf.EliminaFileFisico(NomeFileFinalePDF)
					gf.CreaAggiornaFile(NomeFileFinale, filetto)

					'Dim pp As New pdfGest
					'Ritorno = pp.ConverteHTMLInPDF(NomeFileFinale, NomeFileFinalePDF, PathLog, True)

					'Ritorno = ConvertePDF(NomeFileFinale, NomeFileFinalePDF)

					'If Ritorno = "*" Then
					'	Ritorno = NomeFileFinalePDF.Replace(pathAllegati, pathMultimedia).Replace("Multimedia", "Allegati").Replace("\", "/")
					'End If

					Dim pp As New pdfGest
					Ritorno = pp.ConverteHTMLInPDF(NomeFileFinale, NomeFileFinalePDF, NomeFileLog)

					If Ritorno = "*" Then
						Ritorno = NomeFileFinalePDF '.Replace(pathAllegati, pathMultimedia).Replace("Multimedia", "Allegati").Replace("\", "/")
						gf.EliminaFileFisico(NomeFileFinale)
						gf.EliminaFileFisico(NomeFileLog)
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
		Dim Barra As String = "\"

		If TipoDB <> "SQLSERVER" Then
			Barra = "/"
		End If

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object
				Dim Sql As String = ""
				Dim Ok As Boolean = True
				Dim Colore As String = "#ccc"

				'Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
				'Dim P() As String = paths.Split(";")
				'If Strings.Right(P(0), 1) <> "\" Then
				'	P(0) &= "\"
				'End If
				'Dim pathAllegati As String = P(0).Replace(vbCrLf, "")
				'If Strings.Right(P(2), 1) <> "/" Then
				'	P(2) &= "/"
				'End If
				'Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
				'Dim PathBaseMultimedia As String = pathMultimedia.Replace("Allegati", "Multimedia")

				Dim wsImm As New wsImmagini

				Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
				Dim P() As String = paths.Split(";")
				P(0) = P(0).Trim.Replace(vbCrLf, "")
				If Strings.Right(P(0), 1) <> Barra Then
					P(0) &= Barra
				End If
				Dim pathAllegati As String = P(0).Replace(vbCrLf, "")

				P(1) = P(1).Trim.Replace(vbCrLf, "")
				If Strings.Right(P(1), 1) <> Barra Then
					P(1) &= Barra
				End If
				Dim pathLogG As String = P(1).Replace(vbCrLf, "")

				P(2) = P(2).Trim.Replace(vbCrLf, "")
				If Strings.Right(P(2), 1) <> Barra Then
					P(2) &= Barra
				End If
				Dim pathMultimedia As String = P(2).Replace(vbCrLf, "")
				Dim PathBaseImmagini As String = pathMultimedia & "ImmaginiLocali"
				Dim ImmagineSconosciuta As String = PathBaseImmagini & "/Sconosciuto.png"

				Dim NomeSquadra As String = ""
				Dim ss() As String = Squadra.Split("_")
				Sql = "Select * From [Generale].[dbo].[Squadre] Where idSquadra = " & Val(ss(1)).ToString
				Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ok = False
					Ritorno = "Problemi lettura squadra"
				Else
					If Rec.Eof() Then
					Else
						NomeSquadra = "" & Rec("Descrizione").Value
					End If
					Rec.Close()
				End If

				Dim NomeCategoria As String = ""
				If Ok Then
					Sql = "Select * From Categorie Where idCategoria=" & idCategoria
					Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ok = False
						Ritorno = "Problemi lettura categoria"
					Else
						If Rec.Eof() Then
						Else
							NomeCategoria = "" & Rec("Descrizione").Value & " " & Rec("AnnoCategoria").Value
						End If
						Rec.Close()
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
						Rec = Conn.LeggeQuery(Server.MapPath("."), Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Stampa &= "<table style=""width: 100%;"" cellpading=""0"" cellspacing=""0"">" & vbCrLf
							Stampa &= "<tr style=""background-color: gray;"">" & vbCrLf
							Stampa &= "<th></th>" & vbCrLf
							Stampa &= "<th style=""text-align: center;"">Squadra in Casa</th>" & vbCrLf
							Stampa &= "<th></th>" & vbCrLf
							Stampa &= "<th style=""text-align: center;"">Squadra fuori Casa</th>" & vbCrLf
							Stampa &= "<th style=""text-align: right;"">Risultato</th>" & vbCrLf
							Stampa &= "</tr>" & vbCrLf

							Do Until Rec.Eof()
								Dim imm1 As String = ""
								Dim imm2 As String = ""
								Dim Casa As String = "" & Rec("Casa").Value
								Dim Fuori As String = "" & Rec("Fuori").Value

								If Rec("idSqCasa").Value = -1 Or Rec("idSqCasa").Value = 9999 Then
									Imm1 = wsImm.RitornaImmagineDB(Squadra, "Categorie", idCategoria, "")
									If Imm1.Contains(StringaErrore) Then
										Imm1 = ImmagineSconosciuta
									Else
										Imm1 = "data:image/png;base64,'" + Imm1
									End If

									' imm = PathBaseMultimedia & "/" & NomeSquadra & "/Categorie/" & idAnno & "_" & idCategoria & ".kgb"
									Casa = NomeCategoria
								Else
									imm1 = wsImm.RitornaImmagineDB(Squadra, "Avversari", Rec("idSqCasa").Value, "")
									If imm1.Contains(StringaErrore) Then
										imm1 = ImmagineSconosciuta
									Else
										imm1 = "data:image/png;base64,'" + imm1
									End If

									' imm = PathBaseMultimedia & "/" & NomeSquadra & "/Avversari/" & Rec("idSqCasa").Value & ".kgb"
								End If
								'Dim Imm1 As String = DecriptaImmagine(Server.MapPath("."), imm)

								If Rec("idSqFuori").Value = -1 Or Rec("idSqFuori").Value = 9999 Then
									imm2 = wsImm.RitornaImmagineDB(Squadra, "Categorie", idCategoria, "")
									If imm2.Contains(StringaErrore) Then
										imm2 = ImmagineSconosciuta
									Else
										imm2 = "data:image/png;base64,'" + imm2
									End If

									' imm = PathBaseMultimedia & "/" & NomeSquadra & "/Categorie/" & idAnno & "_" & idCategoria & ".kgb"
									Fuori = NomeCategoria
								Else
									imm2 = wsImm.RitornaImmagineDB(Squadra, "Avversari", Rec("idSqFuori").Value, "")
									If imm2.Contains(StringaErrore) Then
										imm2 = ImmagineSconosciuta
									Else
										imm2 = "data:image/png;base64,'" + imm2
									End If

									' imm = PathBaseMultimedia & "/" & NomeSquadra & "/Avversari/" & Rec("idSqFuori").Value & ".kgb"
								End If
								'Dim Imm2 As String = DecriptaImmagine(Server.MapPath("."), imm)

								Stampa &= "<tr style=""background-color: " & Colore & ";"">" & vbCrLf
								If Colore = "#ccc" Then Colore = "#fff" Else Colore = "#ccc"
								Stampa &= "<td><img src=""" & Imm1 & """ width=""50"" height=""50"" /></td>" & vbCrLf
								Stampa &= "<td>" & Casa & "</td>"
								Stampa &= "<td><img src=""" & Imm2 & """ width=""50"" height=""50"" /></td>" & vbCrLf
								Stampa &= "<td>" & Fuori & "</td>"
								Stampa &= "<td style=""text-align: right;"">" & Rec("Risultato").Value & "</td>"
								Stampa &= "</tr>"

								Rec.MoveNext()
							Loop
							Rec.Close()

							Stampa &= "</table>"

							'Dim PathBaseImmagini As String = pathMultimedia

							Dim filetto As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Scheletri\statistiche_giornata.txt")
							filetto = filetto.Replace("***SFONDO***", PathBaseImmagini & "/bg.jpg")
							filetto = filetto.Replace("***GIORNATA***", idGiornata)
							filetto = filetto.Replace("***CLASSIFICA***", Stampa)

							'Dim Altro As String = Now.Year & Format(Now.Month, "00") & Format(Now.Day, "00") ' & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00")
							Dim NomeFileFinale As String = pathAllegati & Squadra & "\Statistiche\Giornata_" & idGiornata & ".html"
							Dim NomeFileFinalePDF As String = pathAllegati & Squadra & "\Statistiche\Giornata_" & idGiornata & ".pdf"
							Dim NomeFileLog As String = pathAllegati & Squadra & "\Statistiche\LogPDFGiornata_" & idGiornata & ".txt"
							'Dim PathLog As String = Server.MapPath(".") & "\Log\Pdf.txt"
							gf.CreaDirectoryDaPercorso(NomeFileFinale)

							gf.EliminaFileFisico(NomeFileFinale)
							gf.EliminaFileFisico(NomeFileFinalePDF)
							gf.CreaAggiornaFile(NomeFileFinale, filetto)

							'Dim pp As New pdfGest
							'Ritorno = pp.ConverteHTMLInPDF(NomeFileFinale, NomeFileFinalePDF, PathLog, True)

							'Ritorno = ConvertePDF(NomeFileFinale, NomeFileFinalePDF)

							'If Ritorno = "*" Then
							'	Ritorno = NomeFileFinalePDF.Replace(pathAllegati, pathMultimedia).Replace("Multimedia", "Allegati").Replace("\", "/")
							'End If

							Dim pp As New pdfGest
							Ritorno = pp.ConverteHTMLInPDF(NomeFileFinale, NomeFileFinalePDF, NomeFileLog)

							If Ritorno = "*" Then
								Ritorno = NomeFileFinalePDF '.Replace(pathAllegati, pathMultimedia).Replace("Multimedia", "Allegati").Replace("\", "/")
								gf.EliminaFileFisico(NomeFileFinale)
								gf.EliminaFileFisico(NomeFileLog)
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

	Private Function ControllaValiditaSquadra(Squadre As ArrayList, Lista As ArrayList, squadra1 As Integer, squadra2 As Integer) As Boolean
		Dim q As Integer = 0
		Dim FcC As String = ""
		Dim FcF As String = ""
		Dim Ritorno As Boolean = False

		If squadra1 = 9999 Or squadra2 = 9999 Or squadra1 = -1 Or squadra2 = -1 Then
			Return True
		End If

		For Each id As Integer In Squadre
			If id = squadra1 Then
				FcC = IIf(Lista.Item(q) = True, "S", "N")
			End If
			If id = squadra2 Then
				FcF = IIf(Lista.Item(q) = True, "S", "N")
			End If
			q += 1
		Next

		If FcC = "N" And FcF = "N" Then
			Ritorno = True
		Else
			Ritorno = False
		End If

		Return Ritorno
	End Function

	Private Function ConverteStringaSQL(SqlP As String) As String
		Dim Sql As String = SqlP

		If (TipoDB <> "SQLSERVER") Then
			Sql = Sql.ToLower
			Sql = Sql.Replace("[", "")
			Sql = Sql.Replace("]", "")
			Sql = Sql.Replace("dbo.", "")
			Sql = Sql.Replace("generale.", "Generale.")
			Sql = Sql.Replace("iif(", "if(")
		End If

		Return Sql
	End Function
End Class