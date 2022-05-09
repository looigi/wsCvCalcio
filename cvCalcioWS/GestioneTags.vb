Imports System.IO

Public Class GestioneTags
	' Private Connessione As String = ""
	' Private Conn As Object
	Dim PathImmagini As String = "" ' C:\GestioneCampionato\CalcioImages\
	Dim PathAllegati As String = "" ' C:\GestioneCampionato\Allegati
	Dim PathLog As String = "" ' C:\GestioneCampionato\Logs
	Dim UrlAllegati As String = "" ' http://192.168.0.227:92/Multimedia
	Dim PathFirma As String = "" ' C:\GestioneCampionato\Allegati
	Dim nomeFileLogMail As String = ""
	Dim gf As New GestioneFilesDirectory
	Dim Maggiorenne As String = ""
	Dim GenitoriSeparati As String = ""
	Dim AffidamentoCongiunto As String = ""
	Dim idTutore As String = ""
	Dim ceGenitore1 As String = ""
	Dim ceGenitore2 As String = ""
	Dim NomeSquadra As String = ""
	Dim Anno As String = ""
	Dim idAnno As String = ""
	Dim iscrFirmaEntrambi As String = ""
	Dim CodSquadra As String = ""
	Dim idGiocatore As String = ""
	Dim idQuota As String = ""
	Dim idRata As String = ""
	Dim idPagatore As String = ""

	Public Sub New(MP As String)
		' Dim Ritorno As String = ""

		'Connessione = LeggeImpostazioniDiBase(MP, "")

		'If Connessione = "" Then
		'	Ritorno = ErroreConnessioneNonValida
		'Else
		' Conn = New clsGestioneDB(CodSquadra)

		'If TypeOf (Conn) Is String Then
		'	Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
		'Else
		Dim pp As String = gf.LeggeFileIntero(MP & "\Impostazioni\Paths.txt")
		PathImmagini = SistemaPercorso(pp)
		pp = gf.LeggeFileIntero(MP & "\Impostazioni\PathAllegati.txt")

		Dim paths() As String = pp.Split(";")
		PathAllegati = SistemaPercorso(paths(0))
		PathLog = SistemaPercorso(paths(1))
		UrlAllegati = SistemaPercorso(paths(2))
		Dim ppp As String = gf.LeggeFileIntero(MP & "\Impostazioni\PercorsoSitoFirma.txt")
		PathFirma = ppp.Replace(vbCrLf, "")

		'nomeFileLogMail = PathLog & "\Tags_Log.txt"

		'ScriveLog(MP, CodSquadra, "GestioneTags", " - Classe istanziata")
		'ScriveLog(MP, CodSquadra, "GestioneTags", " - Paths Immagini: " & PathImmagini)
		'ScriveLog(MP, CodSquadra, "GestioneTags", " - Paths Allegati: " & PathAllegati)
		'ScriveLog(MP, CodSquadra, "GestioneTags", " - Paths Log: " & PathLog)
		'ScriveLog(MP, CodSquadra, "GestioneTags", " - Url Allegati: " & UrlAllegati)
		'ScriveLog(MP, CodSquadra, "GestioneTags", " - Paths Firma: " & PathFirma)
		'End If
		'End If
	End Sub

	Protected Overrides Sub Finalize()
		'ScriveLog(Mp, CodSquadra, "GestioneTags", " - Classe distrutta")
		'ScriveLog(Mp, CodSquadra, "GestioneTags", "")
		'ScriveLog(Mp, CodSquadra, "GestioneTags", "")
		'ScriveLog(Mp, CodSquadra, "GestioneTags", "")

		'ChiudeDB(True, Conn)
		'Connessione = ""
	End Sub

	Private Function PrendeDatiSollecito(MP As String, Squadra As String, Dati As String)
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(MP, Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim d() As String = Dati.Split(";")
				' 2020-09-30;iscrizione - scuola calcio 2020/21;50;PIRANDOLA;FABIO MASS;Polisportiva_GdC_Ponte_di_Nona

				Dim Rec As Object ' = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Select * From Giocatori Where Cognome='" & d(3) & "' And Nome='" & d(4) & "'"

				Rec = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					idGiocatore = Rec("idGiocatore").Value
					Rec.Close()

					Sql = "Select * From QuoteRate  Where DescRata = '" & d(1) & "' And DataScadenza = '" & d(0) & "' And Importo = " & d(2).Replace(",", ".").Trim
					Rec = Conn.LeggeQuery(MP, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessuna quota rilevata"
						Else
							idQuota = Rec("idQuota").Value
							idRata = Rec("Progressivo").Value
							Rec.Close()

							Sql = "Select NomeSquadra, Descrizione, iscrFirmaEntrambi, idAnno From Anni Where idAnno = 1"
							Rec = Conn.LeggeQuery(MP, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Nessuna squadra rilevata"
								Else
									NomeSquadra = Rec("NomeSquadra").Value
									Anno = Rec("Descrizione").Value
									idAnno = Rec("idAnno").Value
									iscrFirmaEntrambi = "" & Rec("iscrFirmaEntrambi").Value
									Rec.Close()

									Ritorno = "*"
								End If
							End If

							Ritorno = "*"
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	Private Function PrendeDatiDiBase(MP As String, Squadra As String, idGiocatoreP As String, idAnnoP As String)
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(MP, Squadra)

		Dim c() As String = Squadra.Split("_")
		Dim Anno2 As String = Str(Val(c(0))).Trim
		Dim idSquadra As String = Str(Val(c(1))).Trim
		CodSquadra = Squadra
		idAnno = Anno2
		idGiocatore = idGiocatoreP

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(Squadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object ' = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Select MailGenitore1, MailGenitore2, B.Cognome + ' ' + B.Nome As Genitore3 , Genitore1, Genitore2, MailGenitore3, " &
					"B.Maggiorenne, GenitoriSeparati, AffidamentoCongiunto, idTutore " &
					"From GiocatoriDettaglio A " &
					"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
					"Where A.idGiocatore = " & idGiocatore
				Rec = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						Ritorno = StringaErrore & " Nessun dettaglio giocatore rilevato"
					Else
						Maggiorenne = "" & Rec("Maggiorenne").Value
						GenitoriSeparati = "" & Rec("GenitoriSeparati").Value
						AffidamentoCongiunto = "" & Rec("AffidamentoCongiunto").Value
						idTutore = "" & Rec("idTutore").Value
						ceGenitore1 = "" & Rec("Genitore1").Value
						ceGenitore2 = "" & Rec("Genitore2").Value
						Rec.Close()

						Sql = "Select NomeSquadra, Descrizione, iscrFirmaEntrambi From Anni Where idAnno = " & idAnnoP
						Rec = Conn.LeggeQuery(MP, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof() Then
								Ritorno = StringaErrore & " Nessuna squadra rilevata"
							Else
								NomeSquadra = Rec("NomeSquadra").Value
								Anno = Rec("Descrizione").Value
								iscrFirmaEntrambi = "" & Rec("iscrFirmaEntrambi").Value
								Rec.Close()

								Ritorno = "*"
							End If
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	Public Function EsegueFileAssociato(MP As String, NomeSquadra As String, idGiocatore As String, idAnno As String)
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(MP, NomeSquadra, idGiocatore, idAnno)
		Dim Ritorno As String = ""

		If PrendeDati = "*" Then
			ScriveLog(MP, CodSquadra, "GestioneTags", "")
			ScriveLog(MP, CodSquadra, "GestioneTags", " - ASSOCIATO")
			ScriveLog(MP, CodSquadra, "GestioneTags", "")

			Dim gf As New GestioneFilesDirectory

			Dim fileDaCopiare As String = PathAllegati & "\" & CodSquadra & "\Firme\associato_" & idAnno & "_" & idGiocatore & ".html"
			Dim fileDaCopiarePDF As String = PathAllegati & "\" & CodSquadra & "\Firme\associato_" & idAnno & "_" & idGiocatore & ".pdf"
			Dim fileLog As String = PathAllegati & "\" & CodSquadra & "\Firme\associato_" & idAnno & "_" & idGiocatore & ".log"

			'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
			gf.CreaDirectoryDaPercorso(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiarePDF)
			gf.EliminaFileFisico(fileLog)

			Body = EsegueFunzione(MP, "Scheletri\associato.txt", CodSquadra, NomeSquadra, idGiocatore)

			gf.EliminaFileFisico(fileDaCopiare)
			gf.ApreFileDiTestoPerScrittura(fileDaCopiare)
			gf.ScriveTestoSuFileAperto(Body)
			gf.ChiudeFileDiTestoDopoScrittura()

			'File.Copy(fileDaCopiare, fileDaCopiare2)
			Dim pp As New pdfGest

			Ritorno = pp.ConverteHTMLInPDF(fileDaCopiare, fileDaCopiarePDF, fileLog)

			If Ritorno = "*" Then
				gf.EliminaFileFisico(fileDaCopiare)
			End If
		Else
			Body = PrendeDati
		End If

		Return Ritorno
	End Function

	Public Function EsegueMailAssociato(MP As String, NomeSquadra As String, idGiocatore As String, idAnno As String, GenitoreP As String, Privacy As String) As String
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(MP, NomeSquadra, idGiocatore, idAnno)

		If PrendeDati = "*" Then
			ScriveLog(MP, CodSquadra, "GestioneTags", "")
			ScriveLog(MP, CodSquadra, "GestioneTags", " - MAIL ASSOCIATO")
			ScriveLog(MP, CodSquadra, "GestioneTags", "")

			Body = EsegueFunzione(MP, "Scheletri\email_associato.txt", CodSquadra, NomeSquadra, idGiocatore)

			' <a href="%Percorso?firma=true&codSquadra=%Squadra&id=%idGiocatore&squadra=%NomeSquadra&anno=%Anno&genitore=%Genitore&privacy=%Privacy&tipoUtente=1">

			ScriveLog(MP, CodSquadra, "GestioneTags", "Percorso: " & PathFirma)
			ScriveLog(MP, CodSquadra, "GestioneTags", "CodSquadra: " & CodSquadra)
			ScriveLog(MP, CodSquadra, "GestioneTags", "ID Gioc: " & idGiocatore)
			ScriveLog(MP, CodSquadra, "GestioneTags", "Squadra: " & Me.NomeSquadra)
			ScriveLog(MP, CodSquadra, "GestioneTags", "Anno: " & Anno)
			ScriveLog(MP, CodSquadra, "GestioneTags", "Genitore: " & GenitoreP)
			ScriveLog(MP, CodSquadra, "GestioneTags", "Privacy: " & Privacy)

			Dim NumeroFirme As String = ""
			Dim ConnessioneSquadra As String = LeggeImpostazioniDiBase(MP, "")
			Dim Ritorno As String = ""

			If ConnessioneSquadra = "" Then
				Ritorno = ErroreConnessioneNonValida
			Else
				Dim ConnSq As Object = New clsGestioneDB(CodSquadra)

				If TypeOf (ConnSq) Is String Then
					Ritorno = ErroreConnessioneDBNonValida & ":" & ConnSq
				Else
					Dim c() As String = CodSquadra.Split("_")
					Dim Anno2 As String = Str(Val(c(0))).Trim
					Dim idSquadra As String = Str(Val(c(1))).Trim
					Dim Rec As Object

					Dim Sql As String = "Select * From NumeroFirme Where idSquadra=" & idSquadra
					Rec = ConnSq.LeggeQuery(MP, Sql, ConnessioneSquadra)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Not Rec.Eof() Then
							NumeroFirme = Rec("NumeroFirme").Value
						End If
						Rec.Close
					End If

					ScriveLog(MP, CodSquadra, "GestioneTags", "Numero Firme: " & NumeroFirme)

					Body = Body.Replace("%Percorso", PathFirma)
					Body = Body.Replace("%Squadra", CodSquadra)
					Body = Body.Replace("%idGiocatore", idGiocatore)
					Body = Body.Replace("%NomeSquadra", Me.NomeSquadra)
					Body = Body.Replace("%Anno", Anno)
					Body = Body.Replace("%Genitore", GenitoreP)
					Body = Body.Replace("%numeroFirme%", NumeroFirme)
				End If
			End If
		End If

		Return Body
	End Function

	Public Function EsegueMailSollecito(MP As String, NomeSquadra As String, Dati As String) As String
		ScriveLog(MP, CodSquadra, "GestioneTags", " - MAIL SOLLECITO: " & Dati)
		Dim Body As String = ""
		Dim sPrendeDatiSollecito As String = PrendeDatiSollecito(MP, NomeSquadra, Dati)
		Dim PrendeDati As String = PrendeDatiDiBase(MP, NomeSquadra, idGiocatore, idAnno)

		If sPrendeDatiSollecito = "*" And PrendeDati = "*" Then
			ScriveLog(MP, CodSquadra, "GestioneTags", "")
			ScriveLog(MP, CodSquadra, "GestioneTags", " - MAIL SOLLECITO")
			ScriveLog(MP, CodSquadra, "GestioneTags", "")

			Body = EsegueFunzione(MP, "Scheletri\mail_sollecito.txt", CodSquadra, NomeSquadra, idGiocatore)
		End If

		Return Body
	End Function

	Public Function EsegueStampaRicevuta(MP As String, NomeSquadra As String, idGiocatore As String, idAnno As String, Progressivo As String, Dati As String,
										 NumeroRicevuta As String, DataRicevuta As String, Motivazione As String, Intero As String, Virgola As String,
										 ImportoLettere As String, Nominativo As String, idPagatoreP As String) As String
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(MP, NomeSquadra, idGiocatore, idAnno)

		ScriveLog(MP, CodSquadra, "GestioneTags", "")
		ScriveLog(MP, CodSquadra, "GestioneTags", " - STAMPA RICEVUTA")
		ScriveLog(MP, CodSquadra, "GestioneTags", "")

		If PrendeDati = "*" Then
			Dim gf As New GestioneFilesDirectory

			'Dim fileFinale As String = PathAllegati & "\" & CodSquadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\Ricevuta_" & Progressivo & ".pdf"
			'Dim fileAppoggio As String = PathAllegati & "\" & CodSquadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\Ricevuta_" & Progressivo & ".html"
			'Dim fileLog As String = PathAllegati & "\" & CodSquadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\Ricevuta_" & Progressivo & ".log"

			Dim fileFinale As String = MP & "\Appoggio\Ricevuta_" & idGiocatore & "_" & Progressivo & ".pdf"
			Dim fileAppoggio As String = MP & "\Appoggio\Ricevuta_" & idGiocatore & "_" & Progressivo & ".html"
			Dim fileLog As String = MP & "\Appoggio\Ricevuta_" & idGiocatore & "_" & Progressivo & ".log"

			'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
			gf.CreaDirectoryDaPercorso(fileFinale)
			gf.EliminaFileFisico(fileFinale)
			gf.EliminaFileFisico(fileLog)

			idPagatore = idPagatoreP

			Body = EsegueFunzione(MP, "Scheletri\ricevuta_pagamento.txt", CodSquadra, NomeSquadra, idGiocatore)

			Body = Body.Replace("###Dati###", Dati)
			Body = Body.Replace("###Numero Ricevuta###", NumeroRicevuta)
			Body = Body.Replace("###Data Ricevuta###", DataRicevuta)
			Body = Body.Replace("###Corpo Ricevuta###", Motivazione)
			Body = Body.Replace("###Intero###", Intero)
			Body = Body.Replace("###Virgola###", Virgola)
			Body = Body.Replace("###Importo Lettere###", ImportoLettere)
			Body = Body.Replace("###Nominativo###", Nominativo)

			gf.EliminaFileFisico(fileAppoggio)
			gf.ApreFileDiTestoPerScrittura(fileAppoggio)
			gf.ScriveTestoSuFileAperto(Body)
			gf.ChiudeFileDiTestoDopoScrittura()

			Dim pp As New pdfGest

			Dim Ritorno As String = pp.ConverteHTMLInPDF(fileAppoggio, fileFinale, fileLog)

			If Ritorno = "*" Then
				gf.EliminaFileFisico(fileAppoggio)

				Dim imm As New wsImmagini
				Ritorno = imm.SalvaAllegatoDB(CodSquadra, "ricevute", fileFinale, gf.TornaNomeFileDaPath(fileFinale), idGiocatore, Progressivo)
				If Ritorno = "*" Then
					gf.EliminaFileFisico(fileFinale)
					gf.EliminaFileFisico(fileLog)
				End If
			End If
		Else
			ScriveLog(MP, CodSquadra, "GestioneTags", " - ERRORE SU PRENDE DATI: " & NomeSquadra & " - " & idGiocatore & " - " & idAnno)
		End If

		Return Body
	End Function

	Public Function EsegueStampaScontrino(MP As String, NomeSquadra As String, idGiocatore As String, idAnno As String, Progressivo As String, Dati As String,
										 NumeroRicevuta As String, DataRicevuta As String, Motivazione As String, Intero As String, Virgola As String,
										 ImportoLettere As String, Nominativo As String, idPagatoreP As String) As String
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(MP, NomeSquadra, idGiocatore, idAnno)

		ScriveLog(MP, CodSquadra, "GestioneTags", "")
		ScriveLog(MP, CodSquadra, "GestioneTags", " - STAMPA SCONTRINO")
		ScriveLog(MP, CodSquadra, "GestioneTags", "")

		If PrendeDati = "*" Then
			Dim gf As New GestioneFilesDirectory

			'Dim fileFinale As String = PathAllegati & "\" & CodSquadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\Scontrino_" & Progressivo & ".html"
			'Dim fileLog As String = PathAllegati & "\" & CodSquadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\Scontrino_" & Progressivo & ".log"

			Dim fileFinale As String = MP & "\Appoggio\Scontrino_" & idGiocatore & "_" & Progressivo & ".html"
			Dim fileLog As String = MP & "\Appoggio\Scontrino_" & idGiocatore & "_" & Progressivo & ".log"

			'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
			gf.CreaDirectoryDaPercorso(fileFinale)
			gf.EliminaFileFisico(fileFinale)
			gf.EliminaFileFisico(fileLog)

			idPagatore = idPagatoreP

			Body = EsegueFunzione(MP, "Scheletri\ricevuta_scontrino.txt", CodSquadra, NomeSquadra, idGiocatore)

			Body = Body.Replace("###Dati###", Dati)
			Body = Body.Replace("###Numero Ricevuta###", NumeroRicevuta)
			Body = Body.Replace("###Data Ricevuta###", DataRicevuta)
			Body = Body.Replace("###Corpo Ricevuta###", Motivazione)
			Body = Body.Replace("###Importo Scontrino###", Intero & "," & Virgola)
			Body = Body.Replace("###Nominativo###", Nominativo)

			gf.EliminaFileFisico(fileFinale)
			gf.ApreFileDiTestoPerScrittura(fileFinale)
			gf.ScriveTestoSuFileAperto(Body)
			gf.ChiudeFileDiTestoDopoScrittura()

			'Dim pp As New pdfGest

			'Dim Ritorno As String = pp.ConverteHTMLInPDF(fileAppoggio, fileFinale, fileLog)

			'If Ritorno = "*" Then
			'gf.EliminaFileFisico(fileAppoggio)
			'End If

			Dim imm As New wsImmagini
			Dim Ritorno As String = imm.SalvaAllegatoDB(CodSquadra, "scontrini", fileFinale, gf.TornaNomeFileDaPath(fileFinale), idGiocatore, Progressivo)
			If Ritorno = "*" Then
				gf.EliminaFileFisico(fileFinale)
				gf.EliminaFileFisico(fileLog)
			End If
		Else
			ScriveLog(MP, CodSquadra, "GestioneTags", " - ERRORE SU PRENDE DATI: " & NomeSquadra & " - " & idGiocatore & " - " & idAnno)
		End If

		Return Body
	End Function

	Public Function EsegueFileFirme(MP As String, NomeSquadra As String, idGiocatore As String, idAnno As String, DaRichiestaFirma As Boolean)
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(MP, NomeSquadra, idGiocatore, idAnno)

		If PrendeDati = "*" Then
			ScriveLog(MP, CodSquadra, "GestioneTags", "")
			ScriveLog(MP, CodSquadra, "GestioneTags", " - FIRME")
			ScriveLog(MP, CodSquadra, "GestioneTags", "")

			'Dim fileDaCopiare As String = PathAllegati & "\" & CodSquadra & "\Firme\iscrizione_" & idAnno & "_" & idGiocatore & ".html"
			'Dim fileDaCopiarePDF As String = PathAllegati & "\" & CodSquadra & "\Firme\iscrizione_" & idAnno & "_" & idGiocatore & ".pdf"
			'Dim fileLog As String = PathAllegati & "\" & CodSquadra & "\Firme\iscrizione_" & idAnno & "_" & idGiocatore & ".log"

			Dim fileDaCopiare As String = MP & "\Appoggio\iscrizione_" & idAnno & "_" & idGiocatore & ".html"
			Dim fileDaCopiarePDF As String = MP & "\Appoggio\iscrizione_" & idAnno & "_" & idGiocatore & ".pdf"
			Dim fileLog As String = MP & "\Appoggio\iscrizione_" & idAnno & "_" & idGiocatore & ".log"

			'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
			gf.CreaDirectoryDaPercorso(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiarePDF)
			gf.EliminaFileFisico(fileLog)

			Body = EsegueFunzione(MP, "Scheletri\base_iscrizione_.txt", CodSquadra, NomeSquadra, idGiocatore)

			If Maggiorenne = "S" Then
				Body = Body.Replace("---HEIGHT_PADRE---", "0px")
				Body = Body.Replace("---VIS_PADRE---", "hidden")

				Body = Body.Replace("---HEIGHT_MADRE---", "0px")
				Body = Body.Replace("---VIS_MADRE---", "hidden")

				Body = Body.Replace("***HEIGHT_GIOCATORE***", "auto")
				Body = Body.Replace("***VIS GIOCATORE***", "visible")
			Else
				Body = Body.Replace("***HEIGHT_GIOCATORE***", "0px")
				Body = Body.Replace("***VIS GIOCATORE***", "hidden")

				If GenitoriSeparati = "S" Then
					If AffidamentoCongiunto = "S" Then
						If iscrFirmaEntrambi = "S" Then
							Body = Body.Replace("---HEIGHT_PADRE---", "auto")
							Body = Body.Replace("---VIS_PADRE---", "visible")

							Body = Body.Replace("---HEIGHT_MADRE---", "auto")
							Body = Body.Replace("---VIS_MADRE---", "visible")
						Else
							If ceGenitore1 <> "" Then
								Body = Body.Replace("---HEIGHT_PADRE---", "auto")
								Body = Body.Replace("---VIS_PADRE---", "visible")

								Body = Body.Replace("---HEIGHT_MADRE---", "0px")
								Body = Body.Replace("---VIS_MADRE---", "hidden")
							Else
								Body = Body.Replace("---HEIGHT_PADRE---", "0px")
								Body = Body.Replace("---VIS_PADRE---", "hidden")

								Body = Body.Replace("---HEIGHT_MADRE---", "auto")
								Body = Body.Replace("---VIS_MADRE---", "visible")
							End If
						End If
					Else
						If idTutore = "1" Then
							Body = Body.Replace("---HEIGHT_MADRE---", "0px")
							Body = Body.Replace("---VIS_MADRE---", "hidden")
						Else
							Body = Body.Replace("---HEIGHT_PADRE---", "0px")
							Body = Body.Replace("---VIS_PADRE---", "hidden")
						End If
					End If
				Else
					If iscrFirmaEntrambi = "S" Then
						Body = Body.Replace("---HEIGHT_PADRE---", "auto")
						Body = Body.Replace("---VIS_PADRE---", "visible")

						Body = Body.Replace("---HEIGHT_MADRE---", "auto")
						Body = Body.Replace("---VIS_MADRE---", "visible")
					Else
						If ceGenitore1 <> "" Then
							Body = Body.Replace("---HEIGHT_PADRE---", "auto")
							Body = Body.Replace("---VIS_PADRE---", "visible")

							Body = Body.Replace("---HEIGHT_MADRE---", "0px")
							Body = Body.Replace("---VIS_MADRE---", "hidden")
						Else
							Body = Body.Replace("---HEIGHT_PADRE---", "0px")
							Body = Body.Replace("---VIS_PADRE---", "hidden")

							Body = Body.Replace("---HEIGHT_MADRE---", "auto")
							Body = Body.Replace("---VIS_MADRE---", "visible")
						End If
					End If
				End If
			End If

			gf.EliminaFileFisico(fileDaCopiare)
			gf.ApreFileDiTestoPerScrittura(fileDaCopiare)
			gf.ScriveTestoSuFileAperto(Body)
			gf.ChiudeFileDiTestoDopoScrittura()

			'File.Copy(fileDaCopiare, fileDaCopiare2)
			Dim pp As New pdfGest
			Dim Ritorno As String = ""

			Ritorno = pp.ConverteHTMLInPDF(fileDaCopiare, fileDaCopiarePDF, fileLog)

			Dim imm As New wsImmagini
			Ritorno = imm.SalvaAllegatoDB(CodSquadra, "iscrizioni", fileDaCopiarePDF, gf.TornaNomeFileDaPath(fileDaCopiarePDF), idGiocatore, 1)
			If Ritorno = "*" Then
				gf.EliminaFileFisico(fileDaCopiare)
				If DaRichiestaFirma = False Then
					gf.EliminaFileFisico(fileDaCopiarePDF)
				End If
				gf.EliminaFileFisico(fileLog)
			End If

			'If Ritorno = "*" Then
			'	gf.EliminaFileFisico(fileDaCopiare)
			'End If
		Else
			Body = PrendeDati
		End If

		Return Body
	End Function

	Public Function EsegueFilePrivacy(MP As String, NomeSquadra As String, idGiocatore As String, idAnno As String, DaRichiestaFirma As Boolean) As String
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(MP, NomeSquadra, idGiocatore, idAnno)

		If PrendeDati = "*" Then
			ScriveLog(MP, CodSquadra, "GestioneTags", "")
			ScriveLog(MP, CodSquadra, "GestioneTags", " - PRIVACY")
			ScriveLog(MP, CodSquadra, "GestioneTags", "")

			'Dim fileDaCopiare As String = PathAllegati & "\" & CodSquadra & "\Firme\privacy_" & idAnno & "_" & idGiocatore & ".html"
			'Dim fileDaCopiarePDF As String = PathAllegati & "\" & CodSquadra & "\Firme\privacy_" & idAnno & "_" & idGiocatore & ".pdf"
			'Dim fileLog As String = PathAllegati & "\" & CodSquadra & "\Firme\privacy_" & idAnno & "_" & idGiocatore & ".log"
			Dim fileDaCopiare As String = MP & "\Appoggio\Privacy_" & idAnno & "_" & idGiocatore & ".html"
			Dim fileDaCopiarePDF As String = MP & "\Appoggio\Privacy_" & idAnno & "_" & idGiocatore & ".pdf"
			Dim fileLog As String = MP & "\Appoggio\Privacy_" & idAnno & "_" & idGiocatore & ".log"

			'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
			gf.CreaDirectoryDaPercorso(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiarePDF)
			gf.EliminaFileFisico(fileLog)

			Body = EsegueFunzione(MP, "Scheletri\base_privacy.txt", CodSquadra, NomeSquadra, idGiocatore)

			gf.EliminaFileFisico(fileDaCopiare)
			gf.ApreFileDiTestoPerScrittura(fileDaCopiare)
			gf.ScriveTestoSuFileAperto(Body)
			gf.ChiudeFileDiTestoDopoScrittura()

			'File.Copy(fileDaCopiare, fileDaCopiare2)
			Dim pp As New pdfGest
			Dim Ritorno As String = ""

			Ritorno = pp.ConverteHTMLInPDF(fileDaCopiare, fileDaCopiarePDF, fileLog)

			Dim imm As New wsImmagini
			Ritorno = imm.SalvaAllegatoDB(CodSquadra, "privacy", fileDaCopiarePDF, gf.TornaNomeFileDaPath(fileDaCopiarePDF), idGiocatore, 1)
			If Ritorno = "*" Then
				gf.EliminaFileFisico(fileDaCopiare)
				If DaRichiestaFirma = False Then
					gf.EliminaFileFisico(fileDaCopiarePDF)
				End If
				gf.EliminaFileFisico(fileLog)
			End If
		Else
			Body = PrendeDati
		End If

		Return Body
	End Function

	Public Function EsegueFunzione(MP As String, NomefileScheletro As String, CodSquadra As String, NomeSquadra As String, idGiocatore As String)
		Dim fileScheletro As String = PathAllegati & "\" & CodSquadra & "\" & NomefileScheletro
		If Not ControllaEsistenzaFile(fileScheletro) Then
			fileScheletro = MP & "\" & NomefileScheletro
		End If
		ScriveLog(MP, CodSquadra, "GestioneTags", " - File Scheletro: " & fileScheletro)

		'abc***123***def
		'6
		'123*** def
		'4
		'***123***
		Dim Body As String = gf.LeggeFileIntero(fileScheletro)
		While Body.Contains("***")
			ScriveLog(MP, CodSquadra, "GestioneTags", "----------------------------------------------------------------------------")
			Dim Inizio As Integer = Body.IndexOf("***") + 4
			'ScriveLog(Mp, CodSquadra, "GestioneTags", "Tag da ricercare 1:" & Inizio)
			Dim Parte1 As String = Mid(Body, Inizio, 75)
			'ScriveLog(Mp, CodSquadra, "GestioneTags", "Tag da ricercare 2:" & Parte1)
			Dim Altro As Integer = Parte1.IndexOf("***")
			'ScriveLog(Mp, CodSquadra, "GestioneTags", "Tag da ricercare 3:" & Altro)

			'If Parte1 > 0 And Altro > 0 Then
			Dim Parte2 As String = "***" & Mid(Body, Inizio, Altro) & "***"
			'ScriveLog(Mp, CodSquadra, "GestioneTags", "Tag da ricercare:" & Inizio & "-" & Altro & "-" & (Inizio + Altro) & ": " & Parte2)

			Body = Body.Replace(Parte2, EsegueQuery(MP, Parte2))
			'Else
			'If Parte1 > 0 Then

			'Else
			'Exit While
			'End If
			'End If
		End While

		Return Body
	End Function

	Public Function EsegueQuery(MP As String, Tag As String) As String
		Dim Ritorno As String = ""
		Dim Rec As Object ' = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
		Dim Sql As String = "Select idTag, Descrizione, Valore, " & IIf(TipoDB = "SQLSERVER", "IsNull(Query,'')", "Coalesce(Query,'')") & " As Query From Tags Where Trim(Upper(Valore))='" & Tag.Trim.ToUpper & "'"

		ScriveLog(MP, CodSquadra, "GestioneTags", " - Tag: " & Tag & " / CodSquadra:" & CodSquadra & " / Squadra: " & NomeSquadra & " / Anno: " & Anno & " / Parametro: " & idGiocatore)

		Dim Connessione As String = LeggeImpostazioniDiBase(MP, "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = New clsGestioneDB(CodSquadra)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Rec = Conn.LeggeQuery(MP, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof() Then
						ScriveLog(MP, CodSquadra, "GestioneTags", " - Errore sql: " & Sql)

						Ritorno = "ERROR: Nessun tag rilevato"
					Else
						If Rec("Query").Value <> "" Then
							Dim Query As String = "" & Rec("Query").Value

							Rec.Close()

							If Query.Contains("IMMAGINE;") Then
								ScriveLog(MP, CodSquadra, "GestioneTags", " - Query: " & Query)

								Dim campi() As String = Query.Split(";")
								Dim altro1 As String = ""
								Dim altro2 As String = ""
								If campi.Length > 8 Then
									altro2 = campi(8)
								End If
								If campi(5).Contains("_") Then
									Dim a() As String = campi(5).Split("_")
									If a.Length > 2 Then
										altro1 = a(2)
									End If
								End If
								Ritorno = ConverteImmagine(MP, Query, idGiocatore, campi(7), altro2, altro1)

								ScriveLog(MP, CodSquadra, "GestioneTags", " - Ritorno Immagine: " & Ritorno)
							Else
								If Query.Contains("LINK;") Then
									ScriveLog(MP, CodSquadra, "GestioneTags", " - Query: " & Query)

									' <a href="%Percorso?firma=true&codSquadra=%Squadra&id=%idGiocatore&squadra=%NomeSquadra&anno=%Anno&genitore=%Genitore&privacy=%Privacy&tipoUtente=1">Click per firmare</a>
									'Ritorno = Query.Replace("%Percorso", Percorso)
									'Ritorno = Ritorno.Replace("%Squadra", CodSquadra)
									'Ritorno = Ritorno.Replace("%idGiocatore", idGiocatore)
									'Ritorno = Ritorno.Replace("%NomeSquadra", NomeSquadra.Replace(" ", "_"))
									'Ritorno = Ritorno.Replace("%Anno", idAnno)
									'Ritorno = Ritorno.Replace("%Genitore", Genitore)
									'Ritorno = Ritorno.Replace("%Privacy", Privacy)
									Ritorno = Query.Replace("LINK;", "")

									ScriveLog(MP, CodSquadra, "GestioneTags", " - Ritorno LINK: " & Ritorno)
								Else
									Query = Query.Replace("%CodSquadra", "[" & CodSquadra & "].[dbo]")
									Query = Query.Replace("%Anno", idAnno)
									Query = Query.Replace("%idGiocatore", idGiocatore)
									Query = Query.Replace("%idQuota", idQuota)
									Query = Query.Replace("%idRata", idRata)

									ScriveLog(MP, CodSquadra, "GestioneTags", " - Query: " & Query)

									Sql = Query
									Rec = Conn.LeggeQuery(MP, Sql, Connessione)
									If TypeOf (Rec) Is String Then
										Ritorno = Rec
									Else
										If Rec.Eof() Then
											ScriveLog(MP, CodSquadra, "GestioneTags", " - Nessun tag query rilevato. Ritorno -----")

											Ritorno = ""
											' Ritorno = "ERROR: Nessun tag query rilevato"
										Else
											Ritorno = "" & Rec(0).Value

											ScriveLog(MP, CodSquadra, "GestioneTags", " - Ritorno: " & Ritorno)

											Rec.Close()
										End If
									End If
								End If
							End If
						Else
							ScriveLog(MP, CodSquadra, "GestioneTags", " - Query: Vuota. SQL: " & Sql)

							Ritorno = "ERROR: Query Vuota"
						End If
					End If
				End If

			End If
		End If

		Return Ritorno
	End Function

	Private Function ConverteImmagine(Mp As String, Query As String, idGiocatore As String, MetteDefault As String, Progressivo As String, Progressivo2 As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As String = ""
		Dim Campi() As String = Query.Split(";")
		Dim DimensioneX As String = Campi(1)
		Dim DimensioneY As String = Campi(2)
		Dim Tipologia As String = Campi(3)
		Dim Criptata As String = Campi(4)
		Dim NomeImmagine As String = Campi(5)
		Dim Estensione As String = Campi(6)
		'Dim UrlImmagine As String = ""
		'Dim PathIniziale As String = gf.LeggeFileIntero(MP & "\Impostazioni\Paths.txt")
		'PathIniziale = PathIniziale.Replace(vbCrLf, "")
		'If Strings.Right(PathIniziale, 1) = "\" Then
		'	PathIniziale = Mid(PathIniziale, 1, PathIniziale.Length - 1)
		'End If
		'PathIniziale = PathIniziale.Trim
		'Dim UrlIniziale As String = ""
		Dim c As New CriptaFiles

		NomeImmagine = NomeImmagine.Replace("%Anno", idAnno)
		NomeImmagine = NomeImmagine.Replace("%IDGioc", idGiocatore)
		NomeImmagine = NomeImmagine.Replace("%IDPagatore", idPagatore)

		ScriveLog(Mp, CodSquadra, "GestioneTags", " - NomeImmagine: " & NomeImmagine)
		ScriveLog(Mp, CodSquadra, "GestioneTags", " - Dimensione: " & DimensioneX & "/" & DimensioneY)
		ScriveLog(Mp, CodSquadra, "GestioneTags", " - Tipologia: " & Tipologia)
		ScriveLog(Mp, CodSquadra, "GestioneTags", " - Criptata: " & Criptata)
		ScriveLog(Mp, CodSquadra, "GestioneTags", " - Estensione: " & Estensione)
		ScriveLog(Mp, CodSquadra, "GestioneTags", " - PathIniziale: " & PathImmagini)
		ScriveLog(Mp, CodSquadra, "GestioneTags", " - Squadra: " & CodSquadra)
		ScriveLog(Mp, CodSquadra, "GestioneTags", " - ID 1: " & Progressivo)
		ScriveLog(Mp, CodSquadra, "GestioneTags", " - ID 2: " & Progressivo2)

		'If Criptata = "S" Then
		'Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)

		' Dim nomeImm As String = UrlIniziale & "/" & Squadra.Replace(" ", "_") & "/" & Tipologia & "/" & NomeImmagine & ".kgb"
		'Dim pathImm As String = PathImmagini & "\" & NomeSquadra.Replace(" ", "_") & "\" & Tipologia & "\" & NomeImmagine & ".kgb"

		'' ScriveLog(Mp, CodSquadra, "GestioneTags", " - nomeImm: " & nomeImm)
		'ScriveLog(Mp, CodSquadra, "GestioneTags", " - pathImm: " & pathImm)

		'If Tipologia = "Societa" Then
		Dim id As String = ""

		Select Case Tipologia
			Case "Societa"
				id = 1 ' NomeImmagine.Replace("Societa_", "")
			Case "Firme"
				id = idGiocatore
		End Select

		Dim urlImmagine As String = RitornaImmagine(Mp, Tipologia, CodSquadra, id, Progressivo, Progressivo2)
		'Else
		'	If ControllaEsistenzaFile(pathImm) Then
		'		UrlImmagine = UrlAllegati & "/Appoggio/" & NomeSquadra.Replace(" ", "_") & "_" & Esten & "." & Estensione
		'		Dim pathImmConv As String = PathImmagini & "\Appoggio\" & NomeSquadra.Replace(" ", "_") & "_" & Esten & "." & Estensione

		'		ScriveLog(Mp, CodSquadra, "GestioneTags", " - UrlImmagine:" & UrlImmagine)
		'		ScriveLog(Mp, CodSquadra, "GestioneTags", " - pathImmconv: " & pathImmConv)

		'		c.DecryptFile(CryptPasswordString, pathImm, pathImmConv)
		'	Else
		'		If MetteDefault = "S" Then
		'			UrlImmagine = UrlAllegati & "/Sconosciuto.png"
		'		Else
		'			UrlImmagine = ""
		'		End If
		'	End If
		'End If
		'Else
		'	Dim pathImm As String = PathImmagini & "\" & NomeSquadra.Replace(" ", "_") & "\" & Tipologia & "\" & NomeImmagine & "." & Estensione

		'	ScriveLog(Mp, CodSquadra, "GestioneTags", " - pathImm: " & pathImm)

		'	If ControllaEsistenzaFile(pathImm) Then
		'		UrlImmagine = UrlAllegati & "/" & NomeSquadra.Replace(" ", "_") & "/" & Tipologia & "/" & NomeImmagine & "." & Estensione

		'		ScriveLog(Mp, CodSquadra, "GestioneTags", " - UrlImmagine: " & UrlImmagine)
		'	Else
		'		If MetteDefault = "S" Then
		'			UrlImmagine = UrlAllegati & "/Sconosciuto.png"
		'		Else
		'			UrlImmagine = ""
		'		End If
		'	End If
		'End If

		'If UrlImmagine <> "" Then
		ScriveLog(Mp, CodSquadra, "GestioneTags", " - Immagine: " & Mid(urlImmagine, 1, 30) & "...")
		Ritorno = "<img src=""data:image/png;base64," & urlImmagine & """ style=""width: " & DimensioneX & "px; height: " & DimensioneY & "px;"" />"
		'End If

		Return Ritorno
	End Function
End Class
