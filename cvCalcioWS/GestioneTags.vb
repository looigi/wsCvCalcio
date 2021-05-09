Imports System.IO

Public Class GestioneTags
	Private Connessione As String = ""
	Private Conn As Object
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

	Public Sub New()
		Dim Ritorno As String = ""

		Connessione = LeggeImpostazioniDiBase(HttpContext.Current.Server.MapPath("."), "")

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Conn = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim pp As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\Paths.txt")
				PathImmagini = SistemaPercorso(pp)
				pp = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
				Dim paths() As String = pp.Split(";")
				PathAllegati = SistemaPercorso(paths(0))
				PathLog = SistemaPercorso(paths(1))
				UrlAllegati = SistemaPercorso(paths(2))
				Dim ppp As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\PercorsoSito.txt")
				PathFirma = ppp.Replace(vbCrLf, "")

				nomeFileLogMail = PathLog & "\Tags_Log.txt"

				ScriveLog(" - Classe istanziata")
				ScriveLog(" - Paths Immagini: " & PathImmagini)
				ScriveLog(" - Paths Allegati: " & PathAllegati)
				ScriveLog(" - Paths Log: " & PathLog)
				ScriveLog(" - Url Allegati: " & UrlAllegati)
				ScriveLog(" - Paths Firma: " & PathFirma)
			End If
		End If
	End Sub

	Private Function SistemaPercorso(pathPassato As String) As String
		Dim pp As String = pathPassato
		pp = pp.Replace(vbCrLf, "").Trim()
		If Strings.Right(pp, 1) = "\" Or Strings.Right(pp, 1) = "/" Then
			pp = Mid(pp, 1, pp.Length - 1)
		End If

		Return pp
	End Function

	Protected Overrides Sub Finalize()
		ScriveLog(" - Classe distrutta")
		ScriveLog("")
		ScriveLog("")
		ScriveLog("")

		ChiudeDB(True, Conn)
		Connessione = ""
	End Sub

	Private Function PrendeDatiSollecito(Squadra As String, Dati As String)
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(HttpContext.Current.Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim d() As String = Dati.Split(";")
				' 2020-09-30;iscrizione - scuola calcio 2020/21;50;PIRANDOLA;FABIO MASS;Polisportiva_GdC_Ponte_di_Nona

				Dim Rec As Object = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Select * From Giocatori Where Cognome='" & d(3) & "' And Nome='" & d(4) & "'"

				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					idGiocatore = Rec("idGiocatore").Value
					Rec.Close

					Sql = "Select * From QuoteRate  Where DescRata = '" & d(1) & "' And DataScadenza = '" & d(0) & "' And Importo = " & d(2).Replace(",", ".").Trim
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessuna quota rilevata"
						Else
							idQuota = Rec("idQuota").Value
							idRata = Rec("Progressivo").Value
							Rec.Close

							Sql = "Select NomeSquadra, Descrizione, iscrFirmaEntrambi, idAnno From Anni Where idAnno = 1"
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									Ritorno = StringaErrore & " Nessuna squadra rilevata"
								Else
									NomeSquadra = Rec("NomeSquadra").Value
									Anno = Rec("Descrizione").Value
									idAnno = Rec("idAnno").Value
									iscrFirmaEntrambi = "" & Rec("iscrFirmaEntrambi").Value
									Rec.Close

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

	Private Function PrendeDatiDiBase(Squadra As String, idGiocatoreP As String, idAnnoP As String)
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(HttpContext.Current.Server.MapPath("."), Squadra)

		Dim c() As String = Squadra.Split("_")
		Dim Anno2 As String = Str(Val(c(0))).Trim
		CodSquadra = Squadra
		idAnno = Anno2
		idGiocatore = idGiocatoreP

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Rec As Object = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
				Dim Sql As String = "Select MailGenitore1, MailGenitore2, B.Cognome + ' ' + B.Nome As Genitore3 , Genitore1, Genitore2, MailGenitore3, " &
					"B.Maggiorenne, GenitoriSeparati, AffidamentoCongiunto, idTutore " &
					"From GiocatoriDettaglio A " &
					"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
					"Where A.idGiocatore = " & idGiocatore

				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
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
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = StringaErrore & " Nessuna squadra rilevata"
							Else
								NomeSquadra = Rec("NomeSquadra").Value
								Anno = Rec("Descrizione").Value
								iscrFirmaEntrambi = "" & Rec("iscrFirmaEntrambi").Value
								Rec.Close

								Ritorno = "*"
							End If
						End If
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	Public Function EsegueFileAssociato(NomeSquadra As String, idGiocatore As String, idAnno As String)
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(NomeSquadra, idGiocatore, idAnno)
		Dim Ritorno As String = ""

		If PrendeDati = "*" Then
			ScriveLog("")
			ScriveLog(" - ASSOCIATO")
			ScriveLog("")

			Dim gf As New GestioneFilesDirectory

			Dim fileDaCopiare As String = PathAllegati & "\" & CodSquadra & "\Firme\associato_" & idAnno & "_" & idGiocatore & ".html"
			Dim fileDaCopiarePDF As String = PathAllegati & "\" & CodSquadra & "\Firme\associato_" & idAnno & "_" & idGiocatore & ".pdf"
			Dim fileLog As String = PathAllegati & "\" & CodSquadra & "\Firme\associato_" & idAnno & "_" & idGiocatore & ".log"

			'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
			gf.CreaDirectoryDaPercorso(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiarePDF)
			gf.EliminaFileFisico(fileLog)

			Body = EsegueFunzione("Scheletri\associato.txt", CodSquadra, NomeSquadra, idGiocatore)

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

	Public Function EsegueMailAssociato(NomeSquadra As String, idGiocatore As String, idAnno As String, GenitoreP As String, Privacy As String) As String
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(NomeSquadra, idGiocatore, idAnno)

		If PrendeDati = "*" Then
			ScriveLog("")
			ScriveLog(" - MAIL ASSOCIATO")
			ScriveLog("")

			Body = EsegueFunzione("Scheletri\email_associato.txt", CodSquadra, NomeSquadra, idGiocatore)

			' <a href="%Percorso?firma=true&codSquadra=%Squadra&id=%idGiocatore&squadra=%NomeSquadra&anno=%Anno&genitore=%Genitore&privacy=%Privacy&tipoUtente=1">

			ScriveLog("Percorso: " & PathFirma)
			ScriveLog("CodSquadra: " & CodSquadra)
			ScriveLog("ID Gioc: " & idGiocatore)
			ScriveLog("Squadra: " & Me.NomeSquadra)
			ScriveLog("Anno: " & Anno)
			ScriveLog("Genitore: " & GenitoreP)
			ScriveLog("Privacy: " & Privacy)

			Body = Body.Replace("%Percorso", PathFirma)
			Body = Body.Replace("%Squadra", CodSquadra)
			Body = Body.Replace("%idGiocatore", idGiocatore)
			Body = Body.Replace("%NomeSquadra", Me.NomeSquadra)
			Body = Body.Replace("%Anno", Anno)
			Body = Body.Replace("%Genitore", GenitoreP)
			Body = Body.Replace("%Privacy", Privacy)
		End If

		Return Body
	End Function

	Public Function EsegueMailSollecito(NomeSquadra As String, Dati As String) As String
		ScriveLog(" - MAIL SOLLECITO: " & Dati)
		Dim Body As String = ""
		Dim sPrendeDatiSollecito As String = PrendeDatiSollecito(NomeSquadra, Dati)
		Dim PrendeDati As String = PrendeDatiDiBase(NomeSquadra, idGiocatore, idAnno)

		If sPrendeDatiSollecito = "*" And PrendeDati = "*" Then
			ScriveLog("")
			ScriveLog(" - MAIL SOLLECITO")
			ScriveLog("")

			Body = EsegueFunzione("Scheletri\mail_sollecito.txt", CodSquadra, NomeSquadra, idGiocatore)
		End If

		Return Body
	End Function

	Public Function EsegueStampaRicevuta(NomeSquadra As String, idGiocatore As String, idAnno As String, Progressivo As String, Dati As String,
										 NumeroRicevuta As String, DataRicevuta As String, Motivazione As String, Intero As String, Virgola As String,
										 ImportoLettere As String, Nominativo As String, idPagatoreP As String) As String
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(NomeSquadra, idGiocatore, idAnno)

		ScriveLog("")
		ScriveLog(" - STAMPA RICEVUTA")
		ScriveLog("")

		If PrendeDati = "*" Then
			Dim gf As New GestioneFilesDirectory

			Dim fileFinale As String = PathAllegati & "\" & CodSquadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\Ricevuta_" & Progressivo & ".pdf"
			Dim fileAppoggio As String = PathAllegati & "\" & CodSquadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\Ricevuta_" & Progressivo & ".html"
			Dim fileLog As String = PathAllegati & "\" & CodSquadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\Ricevuta_" & Progressivo & ".log"

			'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
			gf.CreaDirectoryDaPercorso(fileFinale)
			gf.EliminaFileFisico(fileFinale)
			gf.EliminaFileFisico(fileLog)

			idPagatore = idPagatoreP

			Body = EsegueFunzione("Scheletri\ricevuta_pagamento.txt", CodSquadra, NomeSquadra, idGiocatore)

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
			End If
		Else
			ScriveLog(" - ERRORE SU PRENDE DATI: " & NomeSquadra & " - " & idGiocatore & " - " & idAnno)
		End If

		Return Body
	End Function

	Public Function EsegueStampaScontrino(NomeSquadra As String, idGiocatore As String, idAnno As String, Progressivo As String, Dati As String,
										 NumeroRicevuta As String, DataRicevuta As String, Motivazione As String, Intero As String, Virgola As String,
										 ImportoLettere As String, Nominativo As String, idPagatoreP As String) As String
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(NomeSquadra, idGiocatore, idAnno)

		ScriveLog("")
		ScriveLog(" - STAMPA SCONTRINO")
		ScriveLog("")

		If PrendeDati = "*" Then
			Dim gf As New GestioneFilesDirectory

			Dim fileFinale As String = PathAllegati & "\" & CodSquadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\Scontrino_" & Progressivo & ".html"
			'Dim fileAppoggio As String = PathAllegati & "\" & CodSquadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\Scontrino_" & Progressivo & ".app"
			Dim fileLog As String = PathAllegati & "\" & CodSquadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\Scontrino_" & Progressivo & ".log"

			'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
			gf.CreaDirectoryDaPercorso(fileFinale)
			gf.EliminaFileFisico(fileFinale)
			gf.EliminaFileFisico(fileLog)

			idPagatore = idPagatoreP

			Body = EsegueFunzione("Scheletri\ricevuta_scontrino.txt", CodSquadra, NomeSquadra, idGiocatore)

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
		Else
			ScriveLog(" - ERRORE SU PRENDE DATI: " & NomeSquadra & " - " & idGiocatore & " - " & idAnno)
		End If

		Return Body
	End Function

	Public Function EsegueFileFirme(NomeSquadra As String, idGiocatore As String, idAnno As String)
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(NomeSquadra, idGiocatore, idAnno)

		If PrendeDati = "*" Then
			ScriveLog("")
			ScriveLog(" - FIRME")
			ScriveLog("")

			Dim fileDaCopiare As String = PathAllegati & "\" & CodSquadra & "\Firme\iscrizione_" & idAnno & "_" & idGiocatore & ".html"
			Dim fileDaCopiarePDF As String = PathAllegati & "\" & CodSquadra & "\Firme\iscrizione_" & idAnno & "_" & idGiocatore & ".pdf"
			Dim fileLog As String = PathAllegati & "\" & CodSquadra & "\Firme\iscrizione_" & idAnno & "_" & idGiocatore & ".log"

			'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
			gf.CreaDirectoryDaPercorso(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiarePDF)
			gf.EliminaFileFisico(fileLog)

			Body = EsegueFunzione("Scheletri\base_iscrizione_.txt", CodSquadra, NomeSquadra, idGiocatore)

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

			If Ritorno = "*" Then
				gf.EliminaFileFisico(fileDaCopiare)
			End If
		Else
			Body = PrendeDati
		End If

		Return Body
	End Function

	Public Function EsegueFilePrivacy(NomeSquadra As String, idGiocatore As String, idAnno As String) As String
		Dim Body As String = ""
		Dim PrendeDati As String = PrendeDatiDiBase(NomeSquadra, idGiocatore, idAnno)

		If PrendeDati = "*" Then
			ScriveLog("")
			ScriveLog(" - PRIVACY")
			ScriveLog("")

			Dim fileDaCopiare As String = PathAllegati & "\" & CodSquadra & "\Firme\privacy_" & idAnno & "_" & idGiocatore & ".html"
			Dim fileDaCopiarePDF As String = PathAllegati & "\" & CodSquadra & "\Firme\privacy_" & idAnno & "_" & idGiocatore & ".pdf"
			Dim fileLog As String = PathAllegati & "\" & CodSquadra & "\Firme\privacy_" & idAnno & "_" & idGiocatore & ".log"

			'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
			gf.CreaDirectoryDaPercorso(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiare)
			gf.EliminaFileFisico(fileDaCopiarePDF)
			gf.EliminaFileFisico(fileLog)

			Body = EsegueFunzione("Scheletri\base_privacy.txt", CodSquadra, NomeSquadra, idGiocatore)

			gf.EliminaFileFisico(fileDaCopiare)
			gf.ApreFileDiTestoPerScrittura(fileDaCopiare)
			gf.ScriveTestoSuFileAperto(Body)
			gf.ChiudeFileDiTestoDopoScrittura()

			'File.Copy(fileDaCopiare, fileDaCopiare2)
			Dim pp As New pdfGest
			Dim Ritorno As String = ""

			Ritorno = pp.ConverteHTMLInPDF(fileDaCopiare, fileDaCopiarePDF, fileLog)

			If Ritorno = "*" Then
				gf.EliminaFileFisico(fileDaCopiare)
			End If
		Else
			Body = PrendeDati
		End If

		Return Body
	End Function

	Public Function EsegueFunzione(NomefileScheletro As String, CodSquadra As String, NomeSquadra As String, idGiocatore As String)
		Dim fileScheletro As String = PathAllegati & "\" & CodSquadra & "\" & NomefileScheletro
		If Not File.Exists(fileScheletro) Then
			fileScheletro = HttpContext.Current.Server.MapPath(".") & "\" & NomefileScheletro
		End If
		ScriveLog(" - File Scheletro: " & fileScheletro)

		'abc***123***def
		'6
		'123*** def
		'4
		'***123***
		Dim Body As String = gf.LeggeFileIntero(fileScheletro)
		While Body.Contains("***")
			ScriveLog("----------------------------------------------------------------------------")
			Dim Inizio As Integer = Body.IndexOf("***") + 4
			'ScriveLog("Tag da ricercare 1:" & Inizio)
			Dim Parte1 As String = Mid(Body, Inizio, 75)
			'ScriveLog("Tag da ricercare 2:" & Parte1)
			Dim Altro As Integer = Parte1.IndexOf("***")
			'ScriveLog("Tag da ricercare 3:" & Altro)

			'If Parte1 > 0 And Altro > 0 Then
			Dim Parte2 As String = "***" & Mid(Body, Inizio, Altro) & "***"
			'ScriveLog("Tag da ricercare:" & Inizio & "-" & Altro & "-" & (Inizio + Altro) & ": " & Parte2)

			Body = Body.Replace(Parte2, EsegueQuery(Parte2))
			'Else
			'If Parte1 > 0 Then

			'Else
			'Exit While
			'End If
			'End If
		End While

		Return Body
	End Function

	Public Function EsegueQuery(Tag As String) As String
		Dim Ritorno As String = ""
		Dim Rec As Object = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
		Dim Sql As String = "Select * From Tags Where Trim(Upper(Valore))='" & Tag.Trim.ToUpper & "'"

		ScriveLog(" - Tag: " & Tag & " / CodSquadra:" & CodSquadra & " / Squadra: " & NomeSquadra & " / Anno: " & Anno & " / Parametro: " & idGiocatore)

		Rec = LeggeQuery(Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				ScriveLog(" - Errore sql: " & Sql)

				Ritorno = "ERROR: Nessun tag rilevato"
			Else
				If Not Rec("Query").Value Is DBNull.Value And "" & Rec("Query").Value <> "" Then
					Dim Query As String = "" & Rec("Query").Value

					Rec.Close()

					If Query.Contains("IMMAGINE;") Then
						ScriveLog(" - Query: " & Query)

						Dim campi() As String = Query.Split(";")
						Ritorno = ConverteImmagine(Query, idGiocatore, campi(7))

						ScriveLog(" - Ritorno Immagine: " & Ritorno)
					Else
						If Query.Contains("LINK;") Then
							ScriveLog(" - Query: " & Query)

							' <a href="%Percorso?firma=true&codSquadra=%Squadra&id=%idGiocatore&squadra=%NomeSquadra&anno=%Anno&genitore=%Genitore&privacy=%Privacy&tipoUtente=1">Click per firmare</a>
							'Ritorno = Query.Replace("%Percorso", Percorso)
							'Ritorno = Ritorno.Replace("%Squadra", CodSquadra)
							'Ritorno = Ritorno.Replace("%idGiocatore", idGiocatore)
							'Ritorno = Ritorno.Replace("%NomeSquadra", NomeSquadra.Replace(" ", "_"))
							'Ritorno = Ritorno.Replace("%Anno", idAnno)
							'Ritorno = Ritorno.Replace("%Genitore", Genitore)
							'Ritorno = Ritorno.Replace("%Privacy", Privacy)
							Ritorno = Query.Replace("LINK;", "")

							ScriveLog(" - Ritorno LINK: " & Ritorno)
						Else
							Query = Query.Replace("%CodSquadra", "[" & CodSquadra & "].[dbo]")
							Query = Query.Replace("%Anno", idAnno)
							Query = Query.Replace("%idGiocatore", idGiocatore)
							Query = Query.Replace("%idQuota", idQuota)
							Query = Query.Replace("%idRata", idRata)

							ScriveLog(" - Query: " & Query)

							Sql = Query
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									ScriveLog(" - Nessun tag query rilevato. Ritorno -----")

									Ritorno = "-----"
									' Ritorno = "ERROR: Nessun tag query rilevato"
								Else
									Ritorno = "" & Rec(0).Value

									ScriveLog(" - Ritorno: " & Ritorno)

									Rec.Close
								End If
							End If
						End If
					End If
				Else
					ScriveLog(" - Query: Vuota. SQL: " & Sql)

					Ritorno = "ERROR: Query Vuota"
				End If
			End If
		End If

		Return Ritorno
	End Function

	Private Sub ScriveLog(Cosa As String)
		Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

		gf.ApreFileDiTestoPerScrittura(nomeFileLogMail)
		gf.ScriveTestoSuFileAperto(Datella & Cosa)
		gf.ChiudeFileDiTestoDopoScrittura()
	End Sub

	Private Function ConverteImmagine(Query As String, idGiocatore As String, MetteDefault As String) As String
		Dim gf As New GestioneFilesDirectory
		Dim Ritorno As String = ""
		Dim Campi() As String = Query.Split(";")
		Dim DimensioneX As String = Campi(1)
		Dim DimensioneY As String = Campi(2)
		Dim Tipologia As String = Campi(3)
		Dim Criptata As String = Campi(4)
		Dim NomeImmagine As String = Campi(5)
		Dim Estensione As String = Campi(6)
		Dim UrlImmagine As String = ""
		'Dim PathIniziale As String = gf.LeggeFileIntero(HttpContext.Current.Server.MapPath(".") & "\Impostazioni\Paths.txt")
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

		ScriveLog(" - NomeImmagine: " & NomeImmagine)
		ScriveLog(" - Dimensione: " & DimensioneX & "/" & DimensioneY)
		ScriveLog(" - Tipologia: " & Tipologia)
		ScriveLog(" - Criptata: " & Criptata)
		ScriveLog(" - Estensione: " & Estensione)
		ScriveLog(" - PathIniziale: " & PathImmagini)

		If Criptata = "S" Then
			Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)

			' Dim nomeImm As String = UrlIniziale & "/" & Squadra.Replace(" ", "_") & "/" & Tipologia & "/" & NomeImmagine & ".kgb"
			Dim pathImm As String = PathImmagini & "\" & NomeSquadra.Replace(" ", "_") & "\" & Tipologia & "\" & NomeImmagine & ".kgb"

			' ScriveLog(" - nomeImm: " & nomeImm)
			ScriveLog(" - pathImm: " & pathImm)

			If File.Exists(pathImm) Then
				UrlImmagine = UrlAllegati & "/Appoggio/" & NomeSquadra.Replace(" ", "_") & "_" & Esten & "." & Estensione
				Dim pathImmConv As String = PathImmagini & "\Appoggio\" & NomeSquadra.Replace(" ", "_") & "_" & Esten & "." & Estensione

				ScriveLog(" - UrlImmagine:" & UrlImmagine)
				ScriveLog(" - pathImmconv: " & pathImmConv)

				c.DecryptFile(CryptPasswordString, pathImm, pathImmConv)
			Else
				If MetteDefault = "S" Then
					UrlImmagine = UrlAllegati & "/Sconosciuto.png"
				Else
					UrlImmagine = ""
				End If
			End If
		Else
			Dim pathImm As String = PathImmagini & "\" & NomeSquadra.Replace(" ", "_") & "\" & Tipologia & "\" & NomeImmagine & "." & Estensione

			ScriveLog(" - pathImm: " & pathImm)

			If File.Exists(pathImm) Then
				UrlImmagine = UrlAllegati & "/" & NomeSquadra.Replace(" ", "_") & "/" & Tipologia & "/" & NomeImmagine & "." & Estensione

				ScriveLog(" - UrlImmagine: " & UrlImmagine)
			Else
				If MetteDefault = "S" Then
					UrlImmagine = UrlAllegati & "/Sconosciuto.png"
				Else
					UrlImmagine = ""
				End If
			End If
		End If

		If UrlImmagine <> "" Then
			Ritorno = "<img src=""" & UrlImmagine & """ style=""width: " & DimensioneX & "px; height: " & DimensioneY & "px;"" />"
		End If

		Return Ritorno
	End Function
End Class
