Imports System.IO

Public Class GestioneTags
	Private Connessione As String = ""
	Private Conn As Object
	Dim PathImmagini As String = "" ' C:\GestioneCampionato\CalcioImages\
	Dim PathAllegati As String = "" ' C:\GestioneCampionato\Allegati
	Dim PathLog As String = "" ' C:\GestioneCampionato\Logs
	Dim UrlAllegati As String = "" ' http://192.168.0.227:92/Multimedia
	Dim nomeFileLogMail As String = ""
	Dim gf As New GestioneFilesDirectory

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

				nomeFileLogMail = PathLog & "\Tags_Log.txt"

				ScriveLog(" - Classe istanziata")
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

		ChiudeDB(True, Conn)
		Connessione = ""
	End Sub

	Public Function EsegueMailAssociato(CodSquadra As String, NomeSquadra As String, idGiocatore As String, Anno As String, Genitore As String, Privacy As String)
		Dim Body As String = EsegueFunzione("Scheletri\mail_associato.txt", CodSquadra, NomeSquadra, idGiocatore, Anno, Genitore, Privacy)

		Return Body
	End Function

	Public Function EsegueFileFirme(CodSquadra As String, NomeSquadra As String, idGiocatore As String, Anno As String, Genitore As String, Privacy As String)
		Dim Body As String = EsegueFunzione("Scheletri\base_iscrizione_.txt", CodSquadra, NomeSquadra, idGiocatore, Anno, Genitore, Privacy)

		Return Body
	End Function

	Public Function EsegueFilePrivacy(CodSquadra As String, NomeSquadra As String, idGiocatore As String, Anno As String, Genitore As String, Privacy As String)
		Dim Body As String = EsegueFunzione("Scheletri\base_privacy.txt", CodSquadra, NomeSquadra, idGiocatore, Anno, Genitore, Privacy)

		Return Body
	End Function

	Public Function EsegueFirma(CodSquadra As String, NomeSquadra As String, idGiocatore As String, Anno As String, Genitore As String, Privacy As String)
		Dim Body As String = EsegueFunzione("Scheletri\base_firma.txt", CodSquadra, NomeSquadra, idGiocatore, Anno, Genitore, Privacy)

		Return Body
	End Function

	Public Function EsegueFunzione(NomefileScheletro As String, CodSquadra As String, NomeSquadra As String, idGiocatore As String, Anno As String, Genitore As String, Privacy As String)
		Dim fileScheletro As String = PathAllegati & "\" & CodSquadra & "\" & NomefileScheletro
		If Not File.Exists(fileScheletro) Then
			fileScheletro = HttpContext.Current.Server.MapPath(".") & "\" & NomefileScheletro
		End If
		ScriveLog(" - File Scheletro: " & fileScheletro)

		Dim Body As String = gf.LeggeFileIntero(fileScheletro)
		While Body.Contains("***")
			Dim Parte1 As String = Body.IndexOf("***") + 3
			Dim Altro As Integer = Parte1.IndexOf("***")

			Parte1 = "***" & Mid(Parte1, 1, Altro) & "***"
			Body = Body.Replace(Parte1, EsegueQuery(Parte1, CodSquadra, NomeSquadra, Anno, idGiocatore, Genitore, Privacy))
		End While

		Return Body
	End Function

	Public Function EsegueQuery(Tag As String, CodSquadra As String, NomeSquadra As String, Anno As String, idGiocatore As String, Genitore As String, Privacy As String) As String
		Dim Ritorno As String = ""
		Dim Rec As Object = HttpContext.Current.Server.CreateObject("ADODB.Recordset")
		Dim Sql As String = "Select * From Tags Where Tag='" & Tag & "'"

		ScriveLog("----------------------------------------------------------------------------")
		ScriveLog(" - Tag: " & Tag & " / CodSquadra:" & CodSquadra & " / Squadra: " & NomeSquadra & " / Anno: " & Anno & " / Parametro: " & idGiocatore)

		Rec = LeggeQuery(Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Ritorno = Rec
		Else
			If Rec.Eof Then
				ScriveLog(" - Nessun tag rilevato")

				Ritorno = "ERROR: Nessun tag rilevato"
			Else
				Dim Query As String = Rec("Query").Value

				Rec.Close()

				If Query.Contains("IMMAGINE;") Then
					ScriveLog(" - Query: " & Query)

					Ritorno = ConverteImmagine(Query, CodSquadra, NomeSquadra, Anno, idGiocatore)

					ScriveLog(" - Ritorno Immagine: " & Ritorno)
				Else
					If Query.Contains("LINK;") Then
						ScriveLog(" - Query: " & Query)

						' <a href="%Percorso?firma=true&codSquadra=%Squadra&id=%idGiocatore&squadra=%NomeSquadra&anno=%Anno&genitore=%Genitore&privacy=%Privacy&tipoUtente=1">Click per firmare</a>
						Ritorno = Query.Replace("%Percorso", Percorso)
						Ritorno = Ritorno.Replace("%Squadra", CodSquadra)
						Ritorno = Ritorno.Replace("%idGiocatore", idGiocatore)
						Ritorno = Ritorno.Replace("%NomeSquadra", NomeSquadra.Replace(" ", "_"))
						Ritorno = Ritorno.Replace("%Anno", Anno)
						Ritorno = Ritorno.Replace("%Genitore", Genitore)
						Ritorno = Ritorno.Replace("%Privacy", Privacy)

						ScriveLog(" - Ritorno LINK: " & Ritorno)
					Else
						Query = Query.Replace("%CodSquadra", "[" & CodSquadra & "].[dbo]")
						Query = Query.Replace("%1", idGiocatore)

						ScriveLog(" - Query: " & Query)

						Sql = Query
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								ScriveLog(" - Nessun tag query rilevato")

								Ritorno = "ERROR: Nessun tag query rilevato"
							Else
								Ritorno = Rec(0).Value

								ScriveLog(" - Ritorno: " & Ritorno)

								Rec.Close
							End If
						End If
					End If
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

	Private Function ConverteImmagine(Query As String, CodSquadra As String, Squadra As String, Anno As String, idGiocatore As String) As String
		Dim Ritorno As String = ""
		Dim Campi() As String = Query.Split(";")
		Dim DimensioneX As String = Campi(1)
		Dim DimensioneY As String = Campi(2)
		Dim Tipologia As String = Campi(3)
		Dim Criptata As String = Campi(4)
		Dim NomeImmagine As String = Campi(5)
		Dim Estensione As String = Campi(6)
		Dim UrlImmagine As String = ""
		Dim PathIniziale As String = ""
		Dim UrlIniziale As String = ""
		Dim c As New CriptaFiles

		NomeImmagine = NomeImmagine.Replace("%Anno", Anno)
		NomeImmagine = NomeImmagine.Replace("%IDGioc", idGiocatore)

		ScriveLog(" - NomeImmagine: " & NomeImmagine)
		ScriveLog(" - Dimensione: " & DimensioneX & "/" & DimensioneY)
		ScriveLog(" - Tipologia: " & Tipologia)
		ScriveLog(" - Criptata: " & Criptata)
		ScriveLog(" - Estensione: " & Estensione)

		If Criptata = "S" Then
			Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)

			' Dim nomeImm As String = UrlIniziale & "/" & Squadra.Replace(" ", "_") & "/" & Tipologia & "/" & NomeImmagine & ".kgb"
			Dim pathImm As String = PathIniziale & "\" & Squadra.Replace(" ", "_") & "\" & Tipologia & "\" & NomeImmagine & ".kgb"

			' ScriveLog(" - nomeImm: " & nomeImm)
			ScriveLog(" - pathImm: " & pathImm)

			If File.Exists(pathImm) Then
				UrlImmagine = UrlIniziale & "/Appoggio/" & Squadra.Replace(" ", "_") & "_" & Esten & "." & Estensione
				Dim pathImmConv As String = PathImmagini & "\Appoggio\" & Squadra.Replace(" ", "_") & "_" & Esten & "." & Estensione

				ScriveLog(" - UrlImmagine:" & UrlImmagine)
				ScriveLog(" - pathImmconv: " & pathImmConv)

				c.DecryptFile(CryptPasswordString, pathImm, pathImmConv)
			Else
				UrlImmagine = UrlAllegati & "/Sconosciuto.png"
			End If
		Else
			Dim pathImm As String = PathIniziale & "\" & Squadra.Replace(" ", "_") & "\" & Tipologia & "\" & NomeImmagine & "." & Estensione

			ScriveLog(" - pathImm: " & pathImm)

			If File.Exists(pathImm) Then
				UrlImmagine = UrlIniziale & "/" & Squadra.Replace(" ", "_") & "/" & Tipologia & "/" & NomeImmagine & "." & Estensione

				ScriveLog(" - UrlImmagine: " & UrlImmagine)
			Else
				UrlImmagine = UrlAllegati & "/Sconosciuto.png"
			End If
		End If

		Ritorno = "<img src=""" & UrlImmagine & """ style=""width: " & DimensioneX & "px; height: " & DimensioneY & "px;"" />"

		Return Ritorno
	End Function
End Class
