Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.FileIO
Imports System.Management
Imports System.Web.Hosting
Imports System.Net.Security
Imports System.Net
Imports System.Diagnostics.Eventing.Reader
Imports SelectPdf
Imports System.Windows.Forms

<System.Web.Services.WebService(Namespace:="http://cvcalcio_gioc.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class wsGiocatori
	Inherits System.Web.Services.WebService

	<WebMethod()>
	Public Function RitornaFirmeDaValidare(Squadra As String, Tutte As String) As String
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
				Dim Altro As String = ""

				If Tutte = "" Or Altro = "N" Or Altro = "NO" Then
					Altro = "Top 3"
				End If

				Sql = "Select " & Altro & " A.*, B.Cognome + ' ' + B.Nome As Giocatore, " &
					"CASE A.idGenitore " &
					"     WHEN 1 THEN C.Genitore1 " &
					"     WHEN 2 THEN C.Genitore2 " &
					"     WHEN 3 THEN B.Cognome + ' ' + B.Nome " &
					"END As Genitore " &
					"From GiocatoriFirme A " &
					"Left Join Giocatori B On A.idGiocatore = B.idGiocatore " &
					"Left Join GiocatoriDettaglio C On A.idGiocatore = C.idGiocatore " &
					"Where (DataFirma Is Not Null And DataFirma <> '') And (Validazione Is Null Or Validazione = '')"
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof Then
						Do Until Rec.Eof
							Ritorno &= Rec("idGiocatore").Value.ToString & ";" &
									Rec("idGenitore").Value.ToString & ";" &
									Rec("Datella").Value.ToString.Trim & ";" &
									Rec("DataFirma").Value.ToString.Trim & ";" &
									Rec("Giocatore").Value.ToString.Trim & ";" &
									Rec("Genitore").Value.ToString.Trim & ";" &
									"§"

							Rec.MoveNext()
						Loop
					End If
					Rec.Close
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ConvalidaFirma(idAnno As String, Squadra As String, idGiocatore As String, idGenitore As String) As String
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
				If Ritorno <> "*" Then
					Ok = False
				End If

				If Ok Then
					Dim dataVal As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
					Sql = "Update GiocatoriFirme Set Validazione='" & dataVal & "' Where idGiocatore=" & idGiocatore & " And idGenitore=" & idGenitore
					Ritorno = EsegueSql(Conn, Sql, Connessione)
					If Ritorno.Contains(StringaErrore) Then
						Ok = False
					End If

					If Ok Then
						Sql = "Select * From GiocatoriDettaglio Where idGiocatore=" & idGiocatore
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = StringaErrore & " Nessun genitore rilevato"
								Ok = False
							Else
								If idGenitore < 3 Then
									Dim Genitore As String = Rec("Genitore" & idGenitore).Value
									Dim Mail As String = Rec("MailGenitore" & idGenitore).Value
									Dim Telefono As String = Rec("TelefonoGenitore" & idGenitore).Value

									Rec.Close

									Sql = "Update GiocatoriMails Set Mail='" & Mail & "' Where idGiocatore=" & idGiocatore & " And Progressivo=" & idGenitore
									Ritorno = EsegueSql(Conn, Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If

									If Ok Then
										Dim idGenitoreLetto As Integer = -1
										Dim GenitoreGiaEsisteComeUtente As Boolean = False
										Dim figliGiaPresenti As String = ""

										Sql = "Select * From [Generale].[dbo].[Utenti] Where EMail='" & Mail.Replace("'", "''") & "'"
										Rec = LeggeQuery(Conn, Sql, Connessione)
										If TypeOf (Rec) Is String Then
											Ritorno = Rec
											Ok = False
										Else
											If Rec.Eof Then
												GenitoreGiaEsisteComeUtente = False
											Else
												GenitoreGiaEsisteComeUtente = True
												idGenitoreLetto = Rec("idUtente").Value
												figliGiaPresenti = Rec("idGiocatore").Value
												If figliGiaPresenti = "-1" Then
													figliGiaPresenti = ""
												End If
												If Strings.Right(figliGiaPresenti, 1) <> ";" Then
													figliGiaPresenti = figliGiaPresenti & ";"
												End If
											End If
										End If
										Rec.Close

										If Ok Then
											If Not GenitoreGiaEsisteComeUtente Then
												Sql = "Select Max(idUtente) + 1 From [Generale].[dbo].[Utenti] Where idAnno=" & idAnno
												Rec = LeggeQuery(Conn, Sql, Connessione)
												If TypeOf (Rec) Is String Then
													Ritorno = Rec
												Else
													If Rec(0).Value Is DBNull.Value Then
														idGenitoreLetto = 1
													Else
														idGenitoreLetto = Rec(0).Value
													End If
												End If
											End If

											Dim g() As String = Genitore.Split(" ")
											Dim s() As String = Squadra.Split("_")
											Dim idSquadra As Integer = Val(s(1))
											Dim pass As String = generaPassRandom()
											Dim nuovaPass() = pass.Split(";")

											If Not GenitoreGiaEsisteComeUtente Then
												Sql = "Insert Into [Generale].[dbo].[Utenti] Values (" &
													" " & idAnno & ", " &
													" " & idGenitoreLetto & ", " &
													"'" & Mail.Replace("'", "''") & "', " &
													"'" & g(0).Replace("'", "''") & "', " &
													"'" & g(1).Replace("'", "''") & "', " &
													"'" & nuovaPass(1).Replace("'", "''") & "', " &
													"'" & Mail.Replace("'", "''") & "', " &
													"-1, " &
													"3, " &
													" " & idSquadra & ", " &
													"1, " &
													"'" & Telefono & "', " &
													"'N', " &
													"'" & idGiocatore & "', " &
													"'N' " &
													")"
											Else
												figliGiaPresenti &= idGiocatore & ";"
												Sql = "Update [Generale].[dbo].[Utenti] Set " &
													"idGiocatore='" & figliGiaPresenti & "' " &
													"Where idUtente=" & idGenitoreLetto
											End If
											Ritorno = EsegueSql(Conn, Sql, Connessione)
											If Ritorno.Contains(StringaErrore) Then
												Ok = False
											Else
												Dim m As New mail
												Dim Oggetto As String = "Nuovo utente inCalcio"
												Dim Body As String = ""
												Body &= "E' stato creato l'utente '" & Genitore.ToUpper & "'. <br />"
												Body &= "Per accedere al sito sarà possibile digitare la mail rilasciata alla segreteria in fase di iscrizione: " & Mail & "<br />"
												Body &= "La password valida per il solo primo accesso è: " & nuovaPass(0) & "<br /><br />"
												Dim ChiScrive As String = "notifiche@incalcio.cloud"

												Ritorno = m.SendEmail(Squadra, "", Oggetto, Body, Mail, "")
											End If
										End If
									End If
								End If
							End If
						End If
					End If
				End If

				If Ok Then
					Ritorno = "*"
					Sql = "Commit"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				Else
					Sql = "Rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function EliminaFirma(Squadra As String, idGiocatore As String, idGenitore As String) As String
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

				Dim c() As String = Squadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim codSquadra As String = c(1)
				Dim NomeSquadra As String = ""

				Sql = "Select NomeSquadra, Descrizione From Anni Where idAnno = " & Anno
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = StringaErrore & " Nessuna squadra rilevata"
					Else
						NomeSquadra = Rec("NomeSquadra").Value
					End If
				End If
				Rec.Close

				If Ritorno = "" Then
					Dim gf As New GestioneFilesDirectory
					Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
					gf = Nothing
					Percorso = Percorso.Trim()
					If Strings.Right(Percorso, 1) <> "\" Then
						Percorso &= "\"
					End If
					Dim path1 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_" & idGenitore & ".kgb"
					If File.Exists(path1) Then
						Try
							File.Delete(path1)
							Ritorno = "*"
						Catch ex As Exception
							Ritorno = StringaErrore & ": " & ex.Message
						End Try
					Else
						Ritorno = StringaErrore & "Firma non esistente"
					End If
				End If

				If Ritorno = "*" Then
					Sql = "Delete From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & idGenitore
					Ritorno = EsegueSql(Conn, Sql, Connessione)
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function AggiornaFirma(Squadra As String, ByVal idGiocatore As String, ByVal Genitore As String) As String
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

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				Dim Datella As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
				Sql = "Update GiocatoriFirme Set DataFirma='" & Datella & "' Where idGiocatore=" & idGiocatore & " And idGenitore=" & Genitore
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Ritorno = "*" Then
					Sql = "commit"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ControllaFirma(Squadra As String, ByVal idGiocatore As String, ByVal Genitore As String) As String
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

				Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & Genitore
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.Eof Then
						Dim Datella As String = Rec("DataFirma").Value
						Rec.Close

						If Datella Is DBNull.Value Or Datella <> "" Then
							If Genitore <> 3 Then
								Ritorno = StringaErrore & " Una firma è già stata inserita per il giocatore ed il genitore in data " & Datella
							Else
								Ritorno = StringaErrore & " Una firma è già stata richiesta per il giocatore in data " & Datella
							End If
						Else
							Ritorno = "*"
						End If
					Else
						Rec.Close
						Ritorno = "*"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RichiedeFirma(Squadra As String, ByVal idGiocatore As String, ByVal Genitore As String, Mittente As String) As String
		' RichiedeFirma?Squadra= 0002_00160&idGiocatore=432&Genitore=1 
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
				Dim c() As String = Squadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim codSquadra As String = c(1)

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				'Sql = "Delete From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & Genitore
				'Ritorno = EsegueSql(Conn, Sql, Connessione)
				'If Ritorno <> "*" Then
				'	Sql = "rollback"
				'	Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)

				'	Return Ritorno
				'End If

				Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=" & Genitore
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Sql = "Insert Into GiocatoriFirme Values (" & idGiocatore & ", " & Genitore & ", '" & Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00") & "', '', '')"
						Ritorno = EsegueSql(Conn, Sql, Connessione)
					End If
				End If

				If Ritorno = "*" Then
					Ritorno = ""

					Sql = "Select NomeSquadra, Descrizione From Anni Where idAnno = " & Anno
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessuna squadra rilevata"
						Else
							Dim NomeSquadra As String = Rec("NomeSquadra").Value
							Dim Descrizione As String = Rec("Descrizione").Value
							Rec.Close

							Sql = "Select * From Giocatori Where idGiocatore = " & idGiocatore
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec.Eof Then
									Ritorno = StringaErrore & " Nessun giocatore rilevato"
								Else
									Dim Nominativo As String = Rec("Cognome").Value & " " & Rec("Nome").Value
									Rec.Close

									Sql = "Select MailGenitore1, MailGenitore2, B.Cognome + ' ' + B.Nome As Genitore3 , Genitore1, Genitore2, MailGenitore3 " &
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
											Dim EMail As String = ""
											Dim nomeGenitore As String = ""

											If Genitore = "1" Then
												EMail = Rec("MailGenitore1").Value
												nomeGenitore = Rec("Genitore1").Value
											Else
												If Genitore = "2" Then
													EMail = Rec("MailGenitore2").Value
													nomeGenitore = Rec("Genitore2").Value
												Else
													EMail = Rec("MailGenitore3").Value
													nomeGenitore = Rec("Genitore3").Value
												End If
											End If
											Rec.Close

											Dim gf As New GestioneFilesDirectory
											Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PercorsoSito.txt")

											If Percorso = "" Then
												Ritorno = StringaErrore & " Nessun percorso sito rilevato"
											Else
												Percorso = Percorso.Trim()
												If Strings.Right(Percorso, 1) <> "/" Then
													Percorso &= "/"
												End If

												Dim m As New mail
												Dim Oggetto As String = "Richiesta Firma inCalcio"
												Dim Body As String = ""

												If Genitore = 3 Then
													Body &= "E' stata richiesta la firma di " & nomeGenitore & " dalla direzione della società " & NomeSquadra & " per l'iscrizione all'anno " & Descrizione & ". <br /><br />"
													Body &= "Per effettuare l'operazione eseguire il seguente link:<br /><br />"
												Else
													Body &= "E' stata richiesta la firma di " & nomeGenitore & " dalla direzione della società " & NomeSquadra & " per l'iscrizione all'anno " & Descrizione & " del giocatore " & Nominativo & ".<br /><br />"
													Body &= "Per effettuare l'operazione eseguire il seguente link:<br /><br />"
												End If

												Body &= "<a href= """ & Percorso & "?firma=true&codSquadra=" & Squadra & "&id=" & idGiocatore & "&squadra=" & NomeSquadra.Replace(" ", "_") & "&anno=" & Anno & "&genitore=" & Genitore & """>"
												Body &= "Click per firmare"
												Body &= "</a>"

												' Dim ChiScrive As String = "notifiche@incalcio.cloud"
												Dim PathAllegati As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
												Dim P() As String = PathAllegati.Split(";")
												If Strings.Right(P(0), 1) = "\" Then
													P(0) = Mid(P(0), 1, P(0).Length - 1)
												End If
												Dim fileDaCopiare As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".html"
												Dim fileDaCopiarePDF As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".pdf"
												Dim fileLog As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".log"
												'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
												gf.CreaDirectoryDaPercorso(fileDaCopiare)
												Dim fileScheletro As String = Server.MapPath(".") & "\Scheletri\base_iscrizione_.txt"

												If File.Exists(fileScheletro) Then
													Try
														Dim fileFirme As String = gf.LeggeFileIntero(fileScheletro)
														fileFirme = RiempieFileFirme(fileFirme, Anno, idGiocatore, Rec, Conn, Connessione, NomeSquadra, P, Descrizione)

														gf.EliminaFileFisico(fileDaCopiare)
														gf.ApreFileDiTestoPerScrittura(fileDaCopiare)
														gf.ScriveTestoSuFileAperto(fileFirme)
														gf.ChiudeFileDiTestoDopoScrittura()

														'File.Copy(fileDaCopiare, fileDaCopiare2)
														Dim pp As New pdfGest
														Ritorno = pp.ConverteHTMLInPDF(fileDaCopiare, fileDaCopiarePDF, fileLog)

														If Ritorno = "*" Then
															gf.EliminaFileFisico(fileDaCopiare)
															Ritorno = m.SendEmail(Squadra, Mittente, Oggetto, Body, EMail, fileDaCopiarePDF)
														End If

														gf = Nothing
													Catch ex As Exception
														Ritorno = StringaErrore & " " & ex.Message
													End Try
												Else
													Ritorno = StringaErrore & " Scheletro iscrizione non trovato"
												End If
												gf = Nothing
												'Ritorno = "*"
											End If
										End If
									End If
								End If
							End If
						End If
					End If
				End If
				' End If
				'End If

				If Ritorno = "*" Then
					Sql = "commit"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				Else
					Sql = "rollback"
					Dim Ritorno2 As String = EsegueSql(Conn, Sql, Connessione)
				End If
			End If
		End If

		Return Ritorno
	End Function

	Private Function RiempieFileFirme(Contenuto As String, Anno As String, idGiocatore As String, Rec As Object, Conn As Object, Connessione As String, Squadra As String, p() As String, DescAnno As String) As String
		Dim c As New CriptaFiles

		'Dim gf As New GestioneFilesDirectory
		'Dim paths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
		'Dim p() As String = paths.Split(";")
		p(2) = p(2).Replace(vbCrLf, "")
		If (Strings.Right(p(2), 1) <> "/") Then
			p(2) = p(2) & "/"
		End If

		Dim Sql As String = "Select * From Anni Where idAnno=" & Anno

		Rec = LeggeQuery(Conn, Sql, Connessione)
		If TypeOf (Rec) Is String Then
			Contenuto = Rec
		Else
			If Not Rec.Eof Then
				Dim NomePolisportiva As String = "" & Rec("NomePolisportiva").value
				Dim NomeCampo As String = "" & Rec("CampoSquadra").value
				Dim Mail As String = "" & Rec("Mail").value
				Dim Telefono As String = "" & Rec("Telefono").value
				Dim SitoWeb As String = "" & Rec("SitoWeb").value
				Dim Indirizzo As String = "" & Rec("Indirizzo").Value
				Dim CodiceFiscale As String = "" & Rec("CodiceFiscale").Value
				Dim PIva As String = "" & Rec("PIva").Value

				Contenuto = Contenuto.Replace("***Anno menu settaggi***", DescAnno)
				Contenuto = Contenuto.Replace("***nome societ&agrave; menu settaggi***", NomePolisportiva)
				Contenuto = Contenuto.Replace("***nome Campo menu settaggi***", NomeCampo)
				Contenuto = Contenuto.Replace("***Telefono - mail - sito web menu settaggi***", Mail & ", " & Telefono & ", " & SitoWeb)
				Contenuto = Contenuto.Replace("***indirizzo menu settaggi tab Dati Generali***", Indirizzo)
				Contenuto = Contenuto.Replace("***codice fiscale menu settaggi***", CodiceFiscale)
				Contenuto = Contenuto.Replace("***partita iva menu settaggi***", PIva)
			Else
				Contenuto = Contenuto.Replace("***Anno menu settaggi***", Anno)
				Contenuto = Contenuto.Replace("***nome societ&agrave; menu settaggi***", "")
				Contenuto = Contenuto.Replace("***nome Campo menu settaggi***", "")
				Contenuto = Contenuto.Replace("***Telefono - mail - sito web menu settaggi***", "")
				Contenuto = Contenuto.Replace("***indirizzo menu settaggi tab Dati Generali***", "")
				Contenuto = Contenuto.Replace("***codice fiscale menu settaggi***", "")
				Contenuto = Contenuto.Replace("***partita iva menu settaggi***", "")
			End If

			Sql = "Select * From Giocatori Where idGiocatore=" & idGiocatore
			Rec = LeggeQuery(Conn, Sql, Connessione)
			If TypeOf (Rec) Is String Then
				Contenuto = Rec
			Else
				If Not Rec.Eof Then
					Dim Cognome As String = "" & Rec("Cognome").value
					Dim Nome As String = "" & Rec("Nome").value
					Dim DataDiNascita As String = "" & Rec("DataDiNascita").value
					Dim CodFisc As String = "" & Rec("CodFiscale").value
					Dim Maschio As String = "" & Rec("Maschio").value
					Dim Indirizzo As String = "" & Rec("Indirizzo").value
					Dim Citta As String = "" & Rec("Citta").value
					Dim EMail As String = "" & Rec("EMail").value
					Dim TelefonoGioc As String = "" & Rec("Telefono").value
					Dim Cap As String = "" & Rec("Cap").value
					Dim CittaNascita As String = "" & Rec("CittaNascita").value

					If Maschio = "M" Then
						Maschio = "Maschile"
					Else
						Maschio = "Femminile"
					End If

					Contenuto = Contenuto.Replace("****cognome menu&nbsp; anagrafica3***", Cognome)
					Contenuto = Contenuto.Replace("***Nome menu anagrafica3***", Nome)
					Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica3***", DataDiNascita)
					Contenuto = Contenuto.Replace("***non c'&egrave;***", "")
					Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica3***", CodFisc)
					Contenuto = Contenuto.Replace("***sesso menu anagrafica***", Maschio)
					Contenuto = Contenuto.Replace("****indirizzo menu anagrafica3***", Indirizzo)
					Contenuto = Contenuto.Replace("***citt&agrave;3***", Citta)
					Contenuto = Contenuto.Replace("***?***", "")
					Contenuto = Contenuto.Replace("*** mail menu anagrafica3***", EMail)
					Contenuto = Contenuto.Replace("***telefono menu anagrafica3***", TelefonoGioc)
					Contenuto = Contenuto.Replace("***?Cap3***", Cap)
					Contenuto = Contenuto.Replace("***Citt&agrave; di nascita3***", CittaNascita)
				Else
					Contenuto = Contenuto.Replace("****cognome menu  anagrafica3***", "")
					Contenuto = Contenuto.Replace("***Nome menu anagrafica3***", "")
					Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica3***", "")
					Contenuto = Contenuto.Replace("***non c'&egrave;***", "")
					Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica3***", "")
					Contenuto = Contenuto.Replace("***sesso menu anagrafica***", "")
					Contenuto = Contenuto.Replace("****indirizzo menu anagrafica3***", "")
					Contenuto = Contenuto.Replace("***citt&agrave;3***", "")
					Contenuto = Contenuto.Replace("***?***", "")
					Contenuto = Contenuto.Replace("*** mail menu anagrafica3***", "")
					Contenuto = Contenuto.Replace("***telefono menu anagrafica3***", "")
					Contenuto = Contenuto.Replace("***?Cap3***", "")
					Contenuto = Contenuto.Replace("***Citt&agrave; di nascita3***", "")
				End If
			End If

			Sql = "Select * From GiocatoriDettaglio Where idGiocatore=" & idGiocatore
			Rec = LeggeQuery(Conn, Sql, Connessione)
			If TypeOf (Rec) Is String Then
				Contenuto = Rec
			Else
				If Not Rec.Eof Then
					Dim Genitore1 As String = "" & Rec("Genitore1").value
					Dim Mail1 As String = "" & Rec("MailGenitore1").value
					Dim Telefono1 As String = "" & Rec("TelefonoGenitore1").value
					Dim Gen1() As String = Genitore1.Split(" ")
					If Gen1.Length = 1 Then
						ReDim Preserve Gen1(2)
					End If
					Dim DataDiNascita1 As String = "" & Rec("DataDiNascita1").Value
					Dim CittaNascita1 As String = "" & Rec("CittaNascita1").Value
					Dim CodFiscale1 As String = "" & Rec("CodFiscale1").Value
					Dim Citta1 As String = "" & Rec("Citta1").Value
					Dim Cap1 As String = "" & Rec("Cap1").Value
					Dim Indirizzo1 As String = "" & Rec("Indirizzo1").Value

					Dim Genitore2 As String = "" & Rec("Genitore2").value
					Dim Mail2 As String = "" & Rec("MailGenitore2").value
					Dim Telefono2 As String = "" & Rec("TelefonoGenitore2").value
					Dim Gen2() As String = Genitore2.Split(" ")
					If Gen2.Length = 1 Then
						ReDim Preserve Gen2(2)
					End If
					Dim DataDiNascita2 As String = "" & Rec("DataDiNascita2").Value
					Dim CittaNascita2 As String = "" & Rec("CittaNascita2").Value
					Dim CodFiscale2 As String = "" & Rec("CodFiscale2").Value
					Dim Citta2 As String = "" & Rec("Citta2").Value
					Dim Cap2 As String = "" & Rec("Cap2").Value
					Dim Indirizzo2 As String = "" & Rec("Indirizzo2").Value

					Contenuto = Contenuto.Replace("****cognome menu&nbsp; anagrafica1***", Gen1(1))
					Contenuto = Contenuto.Replace("***Nome menu anagrafica1***", Gen1(0))
					Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica1***", DataDiNascita1)
					Contenuto = Contenuto.Replace("***Citta di nascita1***", CittaNascita1)
					Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica1***", CodFiscale1)
					Contenuto = Contenuto.Replace("****indirizzo menu anagrafica1***", Indirizzo1)
					Contenuto = Contenuto.Replace("***citt&agrave;1***", Citta1)
					Contenuto = Contenuto.Replace("***?Cap1***", Cap1)
					Contenuto = Contenuto.Replace("*** mail menu anagrafica1***", Mail1)
					Contenuto = Contenuto.Replace("***telefono menu anagrafica1***", Indirizzo1)

					Contenuto = Contenuto.Replace("****cognome menu&nbsp; anagrafica2***", Gen2(1))
					Contenuto = Contenuto.Replace("***Nome menu anagrafica2***", Gen2(0))
					Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica2***", DataDiNascita2)
					Contenuto = Contenuto.Replace("***Citt&agrave; di nascita2;***", CittaNascita2)
					Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica2***", CodFiscale2)
					Contenuto = Contenuto.Replace("****indirizzo menu anagrafica2***", Indirizzo2)
					Contenuto = Contenuto.Replace("***citt&agrave;2***", Citta2)
					Contenuto = Contenuto.Replace("***Cap2***", Cap2)
					Contenuto = Contenuto.Replace("*** mail menu anagrafica2***", Mail2)
					Contenuto = Contenuto.Replace("***telefono menu anagrafica2***", Indirizzo2)
				Else
					Contenuto = Contenuto.Replace("****cognome menu&nbsp; anagrafica1***", "")
					Contenuto = Contenuto.Replace("***Nome menu anagrafica1***", "")
					Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica1***", "")
					Contenuto = Contenuto.Replace("***Citta Nascita 1***", "")
					Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica1***", "")
					Contenuto = Contenuto.Replace("****indirizzo menu anagrafica1***", "")
					Contenuto = Contenuto.Replace("***citt&agrave;1***", "")
					Contenuto = Contenuto.Replace("***cap1***", "")
					Contenuto = Contenuto.Replace("*** mail menu anagrafica1***", "")
					Contenuto = Contenuto.Replace("***telefono menu anagrafica1***", "")

					Contenuto = Contenuto.Replace("****cognome menu&nbsp; anagrafica2***", "")
					Contenuto = Contenuto.Replace("***Nome menu anagrafica2***", "")
					Contenuto = Contenuto.Replace("***Data di nascita menu anagrafica2***", "")
					Contenuto = Contenuto.Replace("***Citta Nascita 2***", "")
					Contenuto = Contenuto.Replace("***codice fiscale menu anagrafica2***", "")
					Contenuto = Contenuto.Replace("****indirizzo menu anagrafica2***", "")
					Contenuto = Contenuto.Replace("***citt&agrave;2***", "")
					Contenuto = Contenuto.Replace("***cap2***", "")
					Contenuto = Contenuto.Replace("*** mail menu anagrafica2***", "")
					Contenuto = Contenuto.Replace("***telefono menu anagrafica2***", "")
				End If
			End If

			Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore
			Rec = LeggeQuery(Conn, Sql, Connessione)
			If TypeOf (Rec) Is String Then
				Contenuto = Rec
			Else
				Do Until Rec.Eof
					Select Case Rec("idGenitore").value
						Case 1
							Contenuto = Contenuto.Replace("***data firma1***", Rec("DataFirma").value)
						Case 2
							Contenuto = Contenuto.Replace("***data firma2***", Rec("DataFirma").value)
						Case 3
							Contenuto = Contenuto.Replace("***data firma3***", Rec("DataFirma").value)
					End Select

					Rec.movenext
				Loop
				Contenuto = Contenuto.Replace("***data firma1***", "")
				Contenuto = Contenuto.Replace("***data firma2***", "")
				Contenuto = Contenuto.Replace("***data firma3***", "")
			End If

			Dim gf As New GestioneFilesDirectory
			Dim pp As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
			pp = pp.Trim()
			If Strings.Right(pp, 1) = "\" Then
				pp = Mid(pp, 1, pp.Length - 1)
			End If
			Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)

			Dim pathFirma1 As String = p(2) & Squadra.Replace(" ", "_") & "/Firme/" & Anno & "_" & idGiocatore & "_1.kgb"
			Dim urlFirma1 As String = pp & "\" & Squadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_1.kgb"
			Dim pathFirmaConv1 As String = p(2) & "/Appoggio/Firma_" & idGiocatore & "_1_" & Esten & ".png"
			Dim urlFirmaConv1 As String = pp & "\Appoggio\Firma_" & idGiocatore & "_1_" & Esten & ".png"

			Dim pathFirma2 As String = p(2) & Squadra.Replace(" ", "_") & "/Firme/" & Anno & "_" & idGiocatore & "_2.kgb"
			Dim urlFirma2 As String = pp & "\" & Squadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_2.kgb"
			Dim pathFirmaConv2 As String = p(2) & "/Appoggio/Firma_" & idGiocatore & "_2_" & Esten & ".png"
			Dim urlFirmaConv2 As String = pp & "\Appoggio\Firma_" & idGiocatore & "_2_" & Esten & ".png"

			Dim pathFirma3 As String = p(2) & Squadra.Replace(" ", "_") & "/Firme/" & Anno & "_" & idGiocatore & "_3.kgb"
			Dim urlFirma3 As String = pp & "\" & Squadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_3.kgb"
			Dim pathFirmaConv3 As String = p(2) & "/Appoggio/Firma_" & idGiocatore & "_3_" & Esten & ".png"
			Dim urlFirmaConv3 As String = pp & "\Appoggio\Firma_" & idGiocatore & "_3_" & Esten & ".png"

			Dim nomeImm As String = p(2) & Squadra.Replace(" ", "_") & "/Societa/" & Anno & "_1.kgb"
			Dim pathImm As String = pp & "\" & Squadra.Replace(" ", "_") & "\Societa\" & Anno & "_1.kgb"
			Dim nomeImmConv As String = p(2) & "/Appoggio/Societa_" & idGiocatore & "_" & Esten & ".png"
			Dim pathImmConv As String = pp & "\Appoggio\Societa_" & idGiocatore & "_" & Esten & ".png"
			c.DecryptFile("WPippoBaudo227!", pathImm, pathImmConv)

			Contenuto = Contenuto.Replace("***immagine logo menu settaggi***", "<img src=""" & nomeImmConv & """ style=""width: 100px; height: 100px;"" />")

			If File.Exists(urlFirma1) Then
				c.DecryptFile("WPippoBaudo227!", urlFirma1, urlFirmaConv1)
				Contenuto = Contenuto.Replace("***firma padre***", "FIRMA: <img src=""" & pathFirmaConv1 & """ style=""width: 300px; height: 100px; border-bottom: 1px solid #black;"" />")
			Else
				Contenuto = Contenuto.Replace("***firma padre***", "FIRMA: " & "")
			End If
			If File.Exists(urlFirma2) Then
				c.DecryptFile("WPippoBaudo227!", urlFirma2, urlFirmaConv2)
				Contenuto = Contenuto.Replace("***firma madre***", "FIRMA: <img src=""" & pathFirmaConv2 & """ style=""width: 300px; height: 100px; border-bottom: 1px solid #black;"" />")
			Else
				Contenuto = Contenuto.Replace("***firma madre***", "FIRMA: " & "")
			End If
			If File.Exists(urlFirma3) Then
				c.DecryptFile("WPippoBaudo227!", urlFirma3, urlFirmaConv3)
				Contenuto = Contenuto.Replace("***firma giocatore***", "FIRMA: <img src=""" & pathFirmaConv3 & """ style=""width: 300px; height: 100px; border-bottom: 1px solid #black;"" />")
			Else
				Contenuto = Contenuto.Replace("***firma giocatore***", "FIRMA: " & "")
			End If
		End If

		Return Contenuto
	End Function

	<WebMethod()>
	Public Function RitornaGiocatoriCategoria(Squadra As String, ByVal idAnno As String, ByVal idCategoria As String) As String
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
				Dim Altro As String = ""

				If idCategoria <> "-1" Then
					Altro = "And CharIndex('" & idCategoria & "-', Categorie) > 0"
				End If

				Try
					Sql = "SELECT idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
						"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, idCategoria2, Categorie.Descrizione As Categoria2, idCategoria3, Cat3.Descrizione As Categoria3, Cat1.Descrizione As Categoria1, " &
						"Giocatori.Categorie, Giocatori.RapportoCompleto, Giocatori.Cap, Giocatori.CittaNascita, Giocatori.Maggiorenne " &
						"FROM (((Giocatori " &
						"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo) " &
						"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria2 And Categorie.idAnno=Giocatori.idAnno) " &
						"Left Join Categorie As Cat3 On Cat3.idCategoria=Giocatori.idCategoria3 And Cat3.idAnno=Giocatori.idAnno) " &
						"Left Join Categorie As Cat1 On Cat1.idCategoria=Giocatori.idCategoria And Cat1.idAnno=Giocatori.idAnno " &
						"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " " & Altro & " " &
						"And RapportoCompleto = 'S' " &
						"Order By Cognome, Nome"
					' "Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " And (Giocatori.idCategoria=" & idCategoria & " Or Giocatori.idCategoria2=" & idCategoria & " Or Giocatori.idCategoria3=" & idCategoria & ") " &
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun giocatore rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idGiocatore").Value.ToString & ";" &
									Rec("idR").Value.ToString & ";" &
									Rec("Cognome").Value.ToString.Trim & ";" &
									Rec("Nome").Value.ToString.Trim & ";" &
									Rec("Descrizione").Value.ToString.Trim & ";" &
									Rec("EMail").Value.ToString.Trim & ";" &
									Rec("Telefono").Value.ToString.Trim & ";" &
									Rec("Soprannome").Value.ToString.Trim & ";" &
									Rec("DataDiNascita").Value.ToString & ";" &
									Rec("Indirizzo").Value.ToString.Trim & ";" &
									Rec("CodFiscale").Value.ToString.Trim & ";" &
									Rec("Maschio").Value.ToString.Trim & ";" &
									Rec("Citta").Value.ToString.Trim & ";" &
									Rec("Matricola").Value.ToString.Trim & ";" &
									Rec("NumeroMaglia").Value.ToString.Trim & ";" &
									Rec("idCategoria").Value.ToString & ";" &
									Rec("idCategoria2").Value.ToString & ";" &
									Rec("Categoria2").Value.ToString & ";" &
									Rec("idCategoria3").Value.ToString & ";" &
									Rec("Categoria3").Value.ToString & ";" &
									Rec("Categoria1").Value.ToString & ";" &
									Rec("Categorie").Value.ToString & ";" &
									Rec("RapportoCompleto").Value.ToString & ";" &
									Rec("Cap").Value.ToString & ";" &
									Rec("CittaNascita").Value.ToString & ";" &
									Rec("Maggiorenne").Value & ";" &
									"§"

								Rec.MoveNext()
							Loop
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
	Public Function RitornaGiocatoriCategoriaSenzaAltri(Squadra As String, ByVal idAnno As String, ByVal idCategoria As String) As String
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

				Try
					Sql = "SELECT idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
						"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, -1 As idCategoria2, '' As Categoria2, -1 As idCategoria3, '' As Categoria3, Categorie.Descrizione As Categoria1, " &
						"Giocatori.Categorie, Giocatori.RapportoCompleto, Giocatori.Cap, Giocatori.CittaNascita, Giocatori.Maggiorenne " &
						"FROM Giocatori " &
						"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo " &
						"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria And Categorie.idAnno=Giocatori.idAnno " &
						"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " And Giocatori.idCategoria=" & idCategoria & " " &
						"And RapportoCompleto = 'S' " &
						"Order By Ruoli.idRuolo, Cognome, Nome"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun giocatore rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idGiocatore").Value.ToString & ";" &
									Rec("idR").Value.ToString & ";" &
									Rec("Cognome").Value.ToString.Trim & ";" &
									Rec("Nome").Value.ToString.Trim & ";" &
									Rec("Descrizione").Value.ToString.Trim & ";" &
									Rec("EMail").Value.ToString.Trim & ";" &
									Rec("Telefono").Value.ToString.Trim & ";" &
									Rec("Soprannome").Value.ToString.Trim & ";" &
									Rec("DataDiNascita").Value.ToString & ";" &
									Rec("Indirizzo").Value.ToString.Trim & ";" &
									Rec("CodFiscale").Value.ToString.Trim & ";" &
									Rec("Maschio").Value.ToString.Trim & ";" &
									Rec("Citta").Value.ToString.Trim & ";" &
									Rec("Matricola").Value.ToString.Trim & ";" &
									Rec("NumeroMaglia").Value.ToString.Trim & ";" &
									Rec("idCategoria").Value.ToString & ";" &
									Rec("idCategoria2").Value.ToString & ";" &
									Rec("Categoria2").Value.ToString & ";" &
									Rec("idCategoria3").Value.ToString & ";" &
									Rec("Categoria3").Value.ToString & ";" &
									Rec("Categoria1").Value.ToString & ";" &
									Rec("Categorie").Value.ToString & ";" &
									Rec("RapportoCompleto").Value.ToString & ";" &
									Rec("Cap").Value.ToString & ";" &
									Rec("CittaNascita").Value.ToString & ";" &
									Rec("Maggiorenne").Value.ToString & ";" &
									"§"

								Rec.MoveNext()
							Loop
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
	Public Function RitornaGiocatoriNonInCategoria(Squadra As String, ByVal idAnno As String, ByVal idCategoria As String) As String
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

				Try
					Sql = "SELECT idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
						"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, -1 As idCategoria2, '' As Categoria2, -1 As idCategoria3, '' As Categoria3, Categorie.Descrizione As Categoria1, " &
						"Giocatori.Categorie, Giocatori.RapportoCompleto, Giocatori.Cap, Giocatori.CittaNascita, Giocatori.Maggiorenne " &
						"FROM Giocatori " &
						"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo " &
						"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria And Categorie.idAnno=Giocatori.idAnno " &
						"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " And CharIndex('" & idCategoria & "-', Categorie) = 0 " &
						"And Giocatori.RapportoCompleto = 'S' " &
						"Order By Ruoli.idRuolo, Cognome, Nome"
					' "Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " And Giocatori.idCategoria<>" & idCategoria & " " &
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun giocatore rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idGiocatore").Value.ToString & ";" &
									Rec("idR").Value.ToString & ";" &
									Rec("Cognome").Value.ToString.Trim & ";" &
									Rec("Nome").Value.ToString.Trim & ";" &
									Rec("Descrizione").Value.ToString.Trim & ";" &
									Rec("EMail").Value.ToString.Trim & ";" &
									Rec("Telefono").Value.ToString.Trim & ";" &
									Rec("Soprannome").Value.ToString.Trim & ";" &
									Rec("DataDiNascita").Value.ToString & ";" &
									Rec("Indirizzo").Value.ToString.Trim & ";" &
									Rec("CodFiscale").Value.ToString.Trim & ";" &
									Rec("Maschio").Value.ToString.Trim & ";" &
									Rec("Citta").Value.ToString.Trim & ";" &
									Rec("Matricola").Value.ToString.Trim & ";" &
									Rec("NumeroMaglia").Value.ToString.Trim & ";" &
									Rec("idCategoria").Value.ToString & ";" &
									Rec("idCategoria2").Value.ToString & ";" &
									Rec("Categoria2").Value.ToString & ";" &
									Rec("idCategoria3").Value.ToString & ";" &
									Rec("Categoria3").Value.ToString & ";" &
									Rec("Categoria1").Value.ToString & ";" &
									Rec("Categorie").Value.ToString & ";" &
									Rec("RapportoCompleto").Value.ToString & ";" &
									Rec("Cap").Value.ToString & ";" &
									Rec("CittaNascita").Value.ToString & ";" &
									Rec("Maggiorenne").Value.ToString & ";" &
									"§"

								Rec.MoveNext()
							Loop
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
	Public Function RitornaGiocatoriTutti(Squadra As String, ByVal idAnno As String) As String
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

				Dim c() As String = Squadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim codSquadra As String = c(1)
				Dim NomeSquadra As String = ""

				Sql = "Select NomeSquadra, Descrizione From Anni Where idAnno = " & Anno
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = StringaErrore & " Nessuna squadra rilevata"
					Else
						NomeSquadra = Rec("NomeSquadra").Value
					End If
				End If
				Rec.Close

				If Ritorno = "" Then
					Dim gf As New GestioneFilesDirectory
					Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
					gf = Nothing
					Percorso = Percorso.Trim()
					If Strings.Right(Percorso, 1) <> "\" Then
						Percorso &= "\"
					End If

					Try
						Sql = "SELECT Giocatori.idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
							"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, Giocatori.idCategoria2 As idCategoria2, Categorie2.Descrizione As Categoria2, " &
							"Giocatori.idCategoria3 As idCategoria3, Categorie3.Descrizione As Categoria3, Categorie.Descrizione As Categoria1, Giocatori.Categorie, " &
							"Giocatori.RapportoCompleto, Giocatori.idTaglia, Min(KitGiocatori.idTipoKit) As idTipologiaKit, Giocatori.Cap, Giocatori.CittaNascita, Giocatori.Maggiorenne " &
							"FROM Giocatori " &
							"Left Join KitGiocatori On Giocatori.idGiocatore=KitGiocatori.idGiocatore " &
							"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo " &
							"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria And Categorie.idAnno=Giocatori.idAnno " &
							"Left Join Categorie As Categorie2 On Categorie2.idCategoria=Giocatori.idCategoria2 And Categorie2.idAnno=Giocatori.idAnno " &
							"Left Join Categorie As Categorie3 On Categorie3.idCategoria=Giocatori.idCategoria3 And Categorie3.idAnno=Giocatori.idAnno " &
							"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " " &
							"Group By Giocatori.idGiocatore, Ruoli.idRuolo, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, CodFiscale, Maschio, " &
							"Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, Giocatori.idCategoria2, Categorie2.Descrizione, Giocatori.idCategoria3, Categorie3.Descrizione, Categorie.Descrizione, " &
							"Giocatori.Categorie, Giocatori.RapportoCompleto, Giocatori.idTaglia, Giocatori.Cap, Giocatori.CittaNascita, Giocatori.Maggiorenne " &
							"Order By Cognome, Nome"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Ritorno = StringaErrore & " Nessun giocatore rilevato"
							Else
								Ritorno = ""
								Do Until Rec.Eof
									Dim Semaforo1 As String = "" : Dim Titolo1 As String = ""
									Dim Semaforo2 As String = "" : Dim Titolo2 As String = ""
									Dim Semaforo3 As String = "" : Dim Titolo3 As String = ""
									Dim Semaforo4 As String = "" : Dim Titolo4 As String = ""
									Dim Semaforo5 As String = "" : Dim Titolo5 As String = ""

									' Semaforo 1: Iscrizione
									Semaforo1 = IIf(Rec("RapportoCompleto").Value = "S", "verde", "rosso")
									Titolo1 = IIf(Rec("RapportoCompleto").Value = "S", "Giocatore iscritto", "Giocatore non iscritto")

									' Semaforo 2: Pagamenti
									Sql = "Select Sum(Pagamento) As Pagato, TotalePagamento As Somma " &
										"From GiocatoriPagamenti A Left Join GiocatoriDettaglio B On A.idAnno = B.idAnno And A.idGiocatore = B.idGiocatore " &
										"Where A.idAnno = " & idAnno & " And A.idGiocatore = " & Rec("idGiocatore").Value & " " &
										"Group By TotalePagamento"
									Rec2 = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec2) Is String Then
										Ritorno = Rec2
									Else
										If Not Rec2.Eof Then
											Semaforo2 = IIf(Rec2("Pagato").Value >= Rec2("Somma").Value, "verde", "giallo")
											Titolo2 = IIf(Rec2("Pagato").Value >= Rec2("Somma").Value, "Pagamento completo", "Pagamento parziale")
										Else
											Semaforo2 = "rosso"
											Titolo2 = "Pagamento non completo"
										End If
										Rec2.Close
									End If

									' Semaforo 3: Firme
									Dim GenitoriSeparati As Boolean = False
									Dim AffidamentoCongiunto As Boolean = False
									Dim Maggiorenne As Boolean = False
									Dim idTutore As String = "M"

									Sql = "Select * From GiocatoriDettaglio Where idGiocatore=" & Rec("idGiocatore").Value
									Rec2 = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec2) Is String Then
										Ritorno = Rec2
									Else
										If Not Rec2.Eof Then
											If "" & Rec2("GenitoriSeparati").Value = "S" Then
												GenitoriSeparati = True
											Else
												GenitoriSeparati = False
											End If
											Maggiorenne = IIf("" & Rec2("Maggiorenne").Value = "S", True, False)
											AffidamentoCongiunto = IIf("" & Rec2("AffidamentoCongiunto").Value = "S", True, False)
											idTutore = "" & Rec2("idTutore").Value
										End If
										Rec2.Close
									End If

									Dim path1 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & Rec("idGiocatore").Value & "_1.kgb"
									Dim path2 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & Rec("idGiocatore").Value & "_2.kgb"
									Dim path3 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & Rec("idGiocatore").Value & "_3.kgb"
									Dim q As Integer = 0
									Dim FirmaPresente1 As Boolean = False
									Dim FirmaPresente2 As Boolean = False
									Dim FirmaPresente3 As Boolean = False
									Dim FirmaValidata1 As Boolean = False
									Dim FirmaValidata2 As Boolean = False
									Dim FirmaValidata3 As Boolean = False
									Dim Validate As Integer = 0

									If File.Exists(path1) Then
										FirmaPresente1 = True
										q += 1

										Sql = "Select * From GiocatoriFirme Where idGiocatore=" & Rec("idGiocatore").Value & " And idGenitore=1"
										Rec2 = LeggeQuery(Conn, Sql, Connessione)
										If TypeOf (Rec2) Is String Then
											Ritorno = Rec2
										Else
											If Not Rec2.Eof Then
												If "" & Rec2("Validazione").Value <> "" Then
													FirmaValidata1 = True
													Validate += 1
												End If
											End If
											Rec2.Close
										End If
									End If
									If File.Exists(path2) Then
										FirmaPresente2 = True
										q += 1

										Sql = "Select * From GiocatoriFirme Where idGiocatore=" & Rec("idGiocatore").Value & " And idGenitore=2"
										Rec2 = LeggeQuery(Conn, Sql, Connessione)
										If TypeOf (Rec2) Is String Then
											Ritorno = Rec2
										Else
											If Not Rec2.Eof Then
												If "" & Rec2("Validazione").Value <> "" Then
													FirmaValidata2 = True
													Validate += 1
												End If
											End If
											Rec2.Close
										End If
									End If
									If File.Exists(path3) Then
										FirmaPresente3 = True
										q += 1

										Sql = "Select * From GiocatoriFirme Where idGiocatore=" & Rec("idGiocatore").Value & " And idGenitore=3"
										Rec2 = LeggeQuery(Conn, Sql, Connessione)
										If TypeOf (Rec2) Is String Then
											Ritorno = Rec2
										Else
											If Not Rec2.Eof Then
												If "" & Rec2("Validazione").Value <> "" Then
													FirmaValidata3 = True
													Validate += 1
												End If
											End If
											Rec2.Close
										End If
									End If

									If Maggiorenne Then
										If FirmaValidata3 = True And FirmaPresente3 Then
											Semaforo3 = "verde"
											Titolo3 = "Firma validata dalla segreteria"
										Else
											Semaforo3 = "rosso"
											Titolo3 = "Nessuna firma validata"
										End If
									Else
										If GenitoriSeparati Then
											If AffidamentoCongiunto Then
												If FirmaPresente1 And FirmaValidata1 And FirmaPresente2 And FirmaValidata2 And FirmaPresente3 And FirmaValidata3 Then
													Semaforo3 = "verde"
													Titolo3 = "Tutte le firme validate"
												Else
													If Validate > 0 Then
														Semaforo3 = "giallo"
														Titolo3 = "Firme non validate (" & Validate & "/3)"
													Else
														Semaforo3 = "rosso"
														Titolo3 = "Nessuna firma validata"
													End If
												End If
											Else
												If idTutore = "M" Then
													If FirmaPresente2 And FirmaValidata2 And FirmaPresente3 And FirmaValidata3 Then
														Semaforo3 = "verde"
														Titolo3 = "Tutte le firme validate"
													Else
														If Validate > 0 Then
															Semaforo3 = "giallo"
															Titolo3 = "Firme non validate (" & Validate & "/2)"
														Else
															Semaforo3 = "rosso"
															Titolo3 = "Nessuna firma validata"
														End If
													End If
												Else
													If FirmaPresente1 And FirmaValidata1 And FirmaPresente3 And FirmaValidata3 Then
														Semaforo3 = "verde"
														Titolo3 = "Tutte le firme validate"
													Else
														If Validate > 0 Then
															Semaforo3 = "giallo"
															Titolo3 = "Firme non validate (" & Validate & "/2)"
														Else
															Semaforo3 = "rosso"
															Titolo3 = "Nessuna firma validata"
														End If
													End If
												End If
											End If
										Else
											If FirmaPresente1 And FirmaValidata1 And FirmaPresente2 And FirmaValidata2 And FirmaPresente3 And FirmaValidata3 Then
												Semaforo3 = "verde"
												Titolo3 = "Tutte le firme validate"
											Else
												If Validate > 0 Then
													Semaforo3 = "giallo"
													Titolo3 = "Firme non validate (" & Validate & "/3)"
												Else
													Semaforo3 = "rosso"
													Titolo3 = "Nessuna firma validata"
												End If
											End If
										End If
									End If

									'Semaforo 4: Certificato
									Sql = "Select CertificatoMedico, ScadenzaCertificatoMedico From GiocatoriDettaglio " &
										"Where idAnno = " & idAnno & " And idGiocatore = " & Rec("idGiocatore").Value
									Rec2 = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec2) Is String Then
										Ritorno = Rec2
									Else
										If Not Rec2.Eof Then
											If Rec2("CertificatoMedico").value = "" Or Rec2("CertificatoMedico").value = "N" Then
												Semaforo4 = "rosso"
												Titolo4 = "Flag certificato non impostato"
											Else
												If Rec2("ScadenzaCertificatoMedico").Value Is DBNull.Value Then
													If Rec2("CertificatoMedico").Value = "S" Then
														Semaforo4 = "giallo"
														Titolo4 = "Certificato presente, Scadenza no"
													Else
														Semaforo4 = "rosso"
														Titolo4 = "Nessun certificato e data presenti"
													End If
												Else
													If Rec2("ScadenzaCertificatoMedico").Value = "" Then
														If Rec2("CertificatoMedico").Value = "S" Then
															Semaforo4 = "giallo"
															Titolo4 = "Certificato presente, Scadenza no"
														Else
															Semaforo4 = "rosso"
															Titolo4 = "Nessun certificato e data presenti"
														End If
													Else
														Dim D() As String = Rec2("ScadenzaCertificatoMedico").Value.split("-")
														Dim dat As Date = Convert.ToDateTime(D(2) & "/" & D(1) & "/" & D(0))

														Dim Scadenza As DateTime = Convert.ToDateTime(Rec2("ScadenzaCertificatoMedico").Value)
														Dim GiorniAllaScadenza As Integer = DateAndTime.DateDiff(DateInterval.Day, Now, Scadenza, )

														If Rec2("CertificatoMedico").Value = "S" And dat > Now Then
															If GiorniAllaScadenza <= 30 Then
																Semaforo4 = "giallo"
																Titolo4 = "Certificato presente ma data scadenza inferiore a 30 giorni"
															Else
																Semaforo4 = "verde"
																Titolo4 = "Certificato e data scadenza presenti"
															End If
														Else
															Semaforo4 = "rosso"
															Titolo4 = "Certificato presente ma con data scaduta"
														End If
													End If
												End If
											End If
										Else
											Semaforo4 = "rosso"
											Titolo4 = "Nessun certificato e data presenti"
										End If
										Rec2.Close
									End If

									' Semaforo 5: KIT
									Sql = "Select C.Descrizione, QuantitaConsegnata, Quantita From KitGiocatori A " &
										"Left Join KitTipologie B On A.idTipoKit = B.idTipoKit " &
										"Left Join KitElementi C On A.idElemento = C.idElemento " &
										"Left Join KitComposizione D On D.idAnno = " & idAnno & " And A.idTipoKit = B.idTipoKit And A.idElemento = C.idElemento And A.idTipoKit = D.idTipoKit  And A.idElemento = D.idElemento " &
										"Where idGiocatore = " & Rec("idGiocatore").Value & " And B.Eliminato = 'N' And C.Eliminato = 'N' And D.Eliminato = 'N'"
									Rec2 = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec2) Is String Then
										Ritorno = Rec2
									Else
										If Rec2.Eof Then
											Semaforo5 = "rosso"
											Titolo5 = "Nessun elemento kit consegnato"
										Else
											Dim Tutto As Boolean = True
											Dim Qualcosa As Boolean = False

											Do Until Rec2.Eof
												If Val(Rec2("QuantitaConsegnata").Value) < Val(Rec2("Quantita").Value) Then
													Qualcosa = True
													Tutto = False
													Exit Do
												Else
													If Val(Rec2("QuantitaConsegnata").Value) > 0 Then
														Qualcosa = True
													End If
												End If

												Rec2.MoveNext()
											Loop

											If Tutto Then
												Semaforo5 = "verde"
												Titolo5 = "Tutto il kit è stato consegnato"
											Else
												If Qualcosa Then
													Semaforo5 = "giallo"
													Titolo5 = "Alcuni elementi del kit sono stati consegnati"
												Else
													Semaforo5 = "rosso"
													Titolo5 = "Nessun elemento kit consegnato"
												End If
											End If
										End If
										Rec2.Close()
									End If

									Ritorno &= Rec("idGiocatore").Value.ToString & ";" &
										Rec("idR").Value.ToString & ";" &
										Rec("Cognome").Value.ToString.Trim & ";" &
										Rec("Nome").Value.ToString.Trim & ";" &
										Rec("Descrizione").Value.ToString.Trim & ";" &
										Rec("EMail").Value.ToString.Trim & ";" &
										Rec("Telefono").Value.ToString.Trim & ";" &
										Rec("Soprannome").Value.ToString.Trim & ";" &
										Rec("DataDiNascita").Value.ToString & ";" &
										Rec("Indirizzo").Value.ToString.Trim & ";" &
										Rec("CodFiscale").Value.ToString.Trim & ";" &
										Rec("Maschio").Value.ToString.Trim & ";" &
										Rec("Citta").Value.ToString.Trim & ";" &
										Rec("Matricola").Value.ToString.Trim & ";" &
										Rec("NumeroMaglia").Value.ToString.Trim & ";" &
										Rec("idCategoria").Value.ToString & ";" &
										Rec("idCategoria2").Value.ToString & ";" &
										Rec("Categoria2").Value.ToString & ";" &
										Rec("idCategoria3").Value.ToString & ";" &
										Rec("Categoria3").Value.ToString & ";" &
										Rec("Categoria1").Value.ToString & ";" &
										Rec("Categorie").Value.ToString & ";" &
										Rec("RapportoCompleto").Value.ToString & ";" &
										Rec("idTaglia").Value.ToString & ";" &
										Semaforo1 & "*" & Titolo1 & ";" &
										Semaforo2 & "*" & Titolo2 & ";" &
										Semaforo3 & "*" & Titolo3 & ";" &
										Semaforo4 & "*" & Titolo4 & ";" &
										Semaforo5 & "*" & Titolo5 & ";" &
										Rec("idTipologiaKit").Value.ToString & ";" &
										Rec("Cap").Value.ToString & ";" &
										Rec("CittaNascita").Value.ToString & ";" &
										Rec("Maggiorenne").Value.ToString & ";" &
										"§"

									Rec.MoveNext()
								Loop
							End If
							Rec.Close()
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
					End Try
				End If

				Conn.Close()
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaGiocatoriDaIscrivere(Squadra As String, ByVal idAnno As String) As String
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

				Try
					Sql = "SELECT idGiocatore, Ruoli.idRuolo As idR, Cognome, Nome, Ruoli.Descrizione, EMail, Telefono, Soprannome, DataDiNascita, Indirizzo, " &
						"CodFiscale, Maschio, Citta, Matricola, NumeroMaglia, Giocatori.idCategoria, Giocatori.idCategoria2 As idCategoria2, Categorie2.Descrizione As Categoria2, " &
						"Giocatori.idCategoria3 As idCategoria3, Categorie3.Descrizione As Categoria3, Categorie.Descrizione As Categoria1, Giocatori.Categorie, " &
						"Giocatori.RapportoCompleto, Giocatori.Cap, Giocatori.CittaNascita, Giocatori.Maggiorenne " &
						"FROM Giocatori " &
						"Left Join [Generale].[dbo].[Ruoli] On Giocatori.idRuolo=Ruoli.idRuolo " &
						"Left Join Categorie On Categorie.idCategoria=Giocatori.idCategoria And Categorie.idAnno=Giocatori.idAnno " &
						"Left Join Categorie As Categorie2 On Categorie2.idCategoria=Giocatori.idCategoria2 And Categorie2.idAnno=Giocatori.idAnno " &
						"Left Join Categorie As Categorie3 On Categorie3.idCategoria=Giocatori.idCategoria3 And Categorie3.idAnno=Giocatori.idAnno " &
						"Where Giocatori.Eliminato='N' And Giocatori.idAnno=" & idAnno & " And Giocatori.RapportoCompleto = 'N' " &
						"Order By Ruoli.idRuolo, Cognome, Nome"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec.Eof Then
							Ritorno = StringaErrore & " Nessun giocatore rilevato"
						Else
							Ritorno = ""
							Do Until Rec.Eof
								Ritorno &= Rec("idGiocatore").Value.ToString & ";" &
									Rec("idR").Value.ToString & ";" &
									Rec("Cognome").Value.ToString.Trim & ";" &
									Rec("Nome").Value.ToString.Trim & ";" &
									Rec("Descrizione").Value.ToString.Trim & ";" &
									Rec("EMail").Value.ToString.Trim & ";" &
									Rec("Telefono").Value.ToString.Trim & ";" &
									Rec("Soprannome").Value.ToString.Trim & ";" &
									Rec("DataDiNascita").Value.ToString & ";" &
									Rec("Indirizzo").Value.ToString.Trim & ";" &
									Rec("CodFiscale").Value.ToString.Trim & ";" &
									Rec("Maschio").Value.ToString.Trim & ";" &
									Rec("Citta").Value.ToString.Trim & ";" &
									Rec("Matricola").Value.ToString.Trim & ";" &
									Rec("NumeroMaglia").Value.ToString.Trim & ";" &
									Rec("idCategoria").Value.ToString & ";" &
									Rec("idCategoria2").Value.ToString & ";" &
									Rec("Categoria2").Value.ToString & ";" &
									Rec("idCategoria3").Value.ToString & ";" &
									Rec("Categoria3").Value.ToString & ";" &
									Rec("Categoria1").Value.ToString & ";" &
									Rec("Categorie").Value.ToString & ";" &
									Rec("RapportoCompleto").Value.ToString & ";" &
									Rec("Cap").Value.ToString & ";" &
									Rec("CittaNascita").Value.ToString & ";" &
									Rec("Maggiorenne").Value.ToString & ";" &
									"§"

								Rec.MoveNext()
							Loop
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
	Public Function SalvaIscrizione(Squadra As String, idAnno As String, idGiocatore As String, RapportoCompleto As String) As String
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
						Sql = "Update Giocatori Set " &
							"RapportoCompleto='" & Replace(RapportoCompleto, "'", "''") & "' " &
							"Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
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
	Public Function SalvaDettaglioGiocatore(Squadra As String, idAnno As String, idGiocatore As String, Genitore1 As String,
											Genitore2 As String, FirmaGenitore1 As String, FirmaGenitore2 As String, CertificatoMedico As String,
											TotalePagamento As String, TelefonoGenitore1 As String, TelefonoGenitore2 As String,
											ScadenzaCertificatoMedico As String, MailGenitore1 As String, MailGenitore2 As String, FirmaGenitore3 As String, MailGenitore3 As String,
											DataDiNascita1 As String, CittaNascita1 As String, CodFiscale1 As String, Citta1 As String, Cap1 As String, Indirizzo1 As String,
											DataDiNascita2 As String, CittaNascita2 As String, CodFiscale2 As String, Citta2 As String, Cap2 As String, Indirizzo2 As String,
											GenitoriSeparati As String, AffidamentoCongiunto As String, AbilitaFirmaGenitore1 As String, AbilitaFirmaGenitore2 As String,
											AbilitaFirmaGenitore3 As String, FirmaAnalogicaGenitore1 As String, FirmaAnalogicaGenitore2 As String, FirmaAnalogicaGenitore3 As String,
											idTutore As String) As String
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
						Sql = "Update GiocatoriDettaglio Set " &
							"Genitore1='" & Replace(Genitore1, "'", "''") & "', " &
							"Genitore2='" & Replace(Genitore2, "'", "''") & "', " &
							"FirmaGenitore1='" & Replace(FirmaGenitore1, "'", "''") & "', " &
							"FirmaGenitore2='" & Replace(FirmaGenitore2, "'", "''") & "', " &
							"CertificatoMedico='" & Replace(CertificatoMedico, "'", "''") & "', " &
							"TotalePagamento=" & TotalePagamento & ", " &
							"TelefonoGenitore1='" & Replace(TelefonoGenitore1, "'", "''") & "', " &
							"TelefonoGenitore2='" & Replace(TelefonoGenitore2, "'", "''") & "', " &
							"ScadenzaCertificatoMedico='" & ScadenzaCertificatoMedico & "', " &
							"MailGenitore1='" & MailGenitore1 & "', " &
							"MailGenitore2='" & MailGenitore2 & "', " &
							"FirmaGenitore3='" & Replace(FirmaGenitore3, "'", "''") & "', " &
							"MailGenitore3='" & MailGenitore3 & "', " &
							"DataDiNascita1='" & DataDiNascita1 & "', " &
							"CittaNascita1='" & CittaNascita1 & "', " &
							"CodFiscale1='" & CodFiscale1 & "', " &
							"Citta1='" & Citta1 & "', " &
							"Cap1='" & Cap1 & "', " &
							"Indirizzo1='" & Indirizzo1 & "', " &
							"DataDiNascita2='" & DataDiNascita2 & "', " &
							"CittaNascita2='" & CittaNascita2 & "', " &
							"CodFiscale2='" & CodFiscale2 & "', " &
							"Citta2='" & Citta2 & "', " &
							"Cap2='" & Cap2 & "', " &
							"Indirizzo2='" & Indirizzo2 & "', " &
							"GenitoriSeparati='" & GenitoriSeparati & "', " &
							"AffidamentoCongiunto='" & AffidamentoCongiunto & "', " &
							"AbilitaFirmaGenitore1='" & AbilitaFirmaGenitore1 & "', " &
							"AbilitaFirmaGenitore2='" & AbilitaFirmaGenitore2 & "', " &
							"AbilitaFirmaGenitore3='" & AbilitaFirmaGenitore3 & "', " &
							"FirmaAnalogicaGenitore1='" & FirmaAnalogicaGenitore1 & "', " &
							"FirmaAnalogicaGenitore2='" & FirmaAnalogicaGenitore2 & "', " &
							"FirmaAnalogicaGenitore3='" & FirmaAnalogicaGenitore3 & "', " &
							"idTutore='" & idTutore & "' " &
							"Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Ok Then
						Sql = "Select * From GiocatoriMails Where idGiocatore=" & idGiocatore
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Dim Attiva As String = ""

								If MailGenitore1 <> "" Then
									Attiva = "S"
								Else
									Attiva = "N"
								End If
								Sql = "Insert Into GiocatoriMails Values (" &
										" " & idGiocatore & ", " &
										"1, " &
										"'" & MailGenitore1.Replace("'", "''") & "', " &
										"'" & Attiva & "' " &
										")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If
								'End If
								If Ok Then
									If MailGenitore2 <> "" Then
										Attiva = "S"
									Else
										Attiva = "N"
									End If
									Sql = "Insert Into GiocatoriMails Values (" &
										" " & idGiocatore & ", " &
										"2, " &
										"'" & MailGenitore2.Replace("'", "''") & "', " &
										"'" & Attiva & "' " &
										")"
									Ritorno = EsegueSql(Conn, Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If
									'End If
								End If
								If Ok Then
									If MailGenitore3 <> "" Then
										Attiva = "S"
									Else
										Attiva = "N"
									End If
									Sql = "Insert Into GiocatoriMails Values (" &
										" " & idGiocatore & ", " &
										"3, " &
										"'" & MailGenitore3.Replace("'", "''") & "', " &
										"'" & Attiva & "' " &
										")"
									Ritorno = EsegueSql(Conn, Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If
									'End If
								End If
							End If
							Rec.Close()
						End If
					End If
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
	Public Function TornaDettaglioGiocatore(Squadra As String, idAnno As String, idGiocatore As String) As String
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
				Dim c() As String = Squadra.Split("_")
				Dim Anno As String = Str(Val(c(0))).Trim
				Dim codSquadra As String = c(1)

				Sql = "Select NomeSquadra, Descrizione From Anni Where idAnno = " & Anno
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = StringaErrore & " Nessuna squadra rilevata"
					Else
						Dim NomeSquadra As String = Rec("NomeSquadra").Value
						Dim Descrizione As String = Rec("Descrizione").Value
						Rec.Close

						Sql = "Select * From GiocatoriDettaglio Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							If Rec.Eof Then
								Sql = "Insert Into GiocatoriDettaglio Values (" &
									" " & idAnno & ", " &
									" " & idGiocatore & ", " &
									"'', " &
									"'', " &
									"'N', " &
									"'N', " &
									"'N', " &
									"0, " &
									"'', " &
									"'', " &
									"null, " &
									"'', " &
									"'', " &
									"'N', " &
									"'', " &
									"'', " &
									"'', " &
									"'', " &
									"'', " &
									"'', " &
									"'', " &
									"'', " &
									"'', " &
									"'', " &
									"'', " &
									"'', " &
									"'', " &
									"'N', " &
									"'N', " &
									"'N', " &
									"'S', " &
									"'S', " &
									"'S', " &
									"'N', " &
									"'N', " &
									"'N', " &
									"'M' " &
									")"
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Not Ritorno.Contains(StringaErrore) Then
									Ritorno = ";;N;N;N;0;;;"
								End If
							Else
								Dim gf As New GestioneFilesDirectory
								Dim Percorso As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
								gf = Nothing
								Percorso = Percorso.Trim()
								If Strings.Right(Percorso, 1) <> "\" Then
									Percorso &= "\"
								End If
								Dim path1 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_1.kgb"
								Dim path2 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_2.kgb"
								Dim path3 As String = Percorso & NomeSquadra.Replace(" ", "_") & "\Firme\" & Anno & "_" & idGiocatore & "_3.kgb"
								Dim dataFirma1 As String = ""
								Dim dataFirma2 As String = ""
								Dim dataFirma3 As String = ""

								Dim firma1 As String = "N"
								If File.Exists(path1) Then
									firma1 = "S"
									Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=1"
									Rec2 = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec2) Is String Then
										Ritorno = Rec2
									Else
										If Not Rec2.Eof Then
											dataFirma1 = Rec2("DataFirma").Value
										End If
										Rec2.Close
									End If
								End If

								Dim firma2 As String = "N"
								If File.Exists(path2) Then
									firma2 = "S"
									Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=2"
									Rec2 = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec2) Is String Then
										Ritorno = Rec2
									Else
										If Not Rec2.Eof Then
											dataFirma2 = Rec2("DataFirma").Value
										End If
										Rec2.Close
									End If
								End If
								Dim firma3 As String = "N"
								If File.Exists(path3) Then
									firma3 = "S"
									Sql = "Select * From GiocatoriFirme Where idGiocatore=" & idGiocatore & " And idGenitore=3"
									Rec2 = LeggeQuery(Conn, Sql, Connessione)
									If TypeOf (Rec2) Is String Then
										Ritorno = Rec2
									Else
										If Not Rec2.Eof Then
											dataFirma3 = Rec2("DataFirma").Value
										End If
										Rec2.Close
									End If
								End If

								Ritorno = Rec("Genitore1").Value & ";"
								Ritorno &= Rec("Genitore2").Value & ";"
								Ritorno &= firma1 & ";"
								Ritorno &= firma2 & ";"
								Ritorno &= Rec("CertificatoMedico").Value & ";"
								Ritorno &= Rec("TotalePagamento").Value & ";"
								Ritorno &= Rec("TelefonoGenitore1").Value & ";"
								Ritorno &= Rec("TelefonoGenitore2").Value & ";"
								Ritorno &= Rec("ScadenzaCertificatoMedico").Value & ";"
								Ritorno &= Rec("MailGenitore1").Value & ";"
								Ritorno &= Rec("MailGenitore2").Value & ";"
								Ritorno &= firma3 & ";"
								Ritorno &= Rec("MailGenitore3").Value & ";"
								Ritorno &= dataFirma1 & ";"
								Ritorno &= dataFirma2 & ";"
								Ritorno &= dataFirma3 & ";"

								Ritorno &= Rec("DataDiNascita1").Value & ";"
								Ritorno &= Rec("CittaNascita1").Value & ";"
								Ritorno &= Rec("CodFiscale1").Value & ";"
								Ritorno &= Rec("Citta1").Value & ";"
								Ritorno &= Rec("Cap1").Value & ";"
								Ritorno &= Rec("Indirizzo1").Value & ";"

								Ritorno &= Rec("DataDiNascita2").Value & ";"
								Ritorno &= Rec("CittaNascita2").Value & ";"
								Ritorno &= Rec("CodFiscale2").Value & ";"
								Ritorno &= Rec("Citta2").Value & ";"
								Ritorno &= Rec("Cap2").Value & ";"
								Ritorno &= Rec("Indirizzo2").Value & ";"
								Ritorno &= Rec("GenitoriSeparati").Value & ";"
								Ritorno &= Rec("AffidamentoCongiunto").Value & ";"
								Ritorno &= Rec("AbilitaFirmaGenitore1").Value & ";"
								Ritorno &= Rec("AbilitaFirmaGenitore2").Value & ";"
								Ritorno &= Rec("AbilitaFirmaGenitore3").Value & ";"
								Ritorno &= Rec("FirmaAnalogicaGenitore1").Value & ";"
								Ritorno &= Rec("FirmaAnalogicaGenitore2").Value & ";"
								Ritorno &= Rec("FirmaAnalogicaGenitore3").Value & ";"
								Ritorno &= Rec("idTutore").Value & ";"
							End If
						End If
					End If
				End If

			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function RitornaNuovoIDGiocatore(Squadra As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)
		Dim idGioc As String = -1

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
					Sql = "SELECT Max(idGiocatore)+1 FROM Giocatori"
					Rec = LeggeQuery(Conn, Sql, Connessione)
					If TypeOf (Rec) Is String Then
						Ritorno = Rec
					Else
						If Rec(0).Value Is DBNull.Value Then
							idGioc = 1
						Else
							idGioc = Rec(0).Value
						End If
						Rec.Close()

						Ritorno = idGioc
					End If
				Catch ex As Exception
					Ritorno = StringaErrore & " " & ex.Message
				End Try
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function SalvaGiocatore(Squadra As String, idAnno As String, idCategoria As String, idGiocatore As String, idRuolo As String, Cognome As String, Nome As String, EMail As String, Telefono As String,
								   Soprannome As String, DataDiNascita As String, Indirizzo As String, CodFiscale As String, Maschio As String, Citta As String, Matricola As String,
								   NumeroMaglia As String, idCategoria2 As String, idCategoria3 As String, Categorie As String, RapportoCompleto As String,
								   idTaglia As String, Modalita As String, Cap As String, CittaNascita As String) As String
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
				Dim idGioc As Integer = -1
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Dim Maggiorenne As String = "N"

					'If idGiocatore = "-1" Then
					If Modalita = "INSERIMENTO" Then
						Sql = "SELECT * FROM Giocatori Where idAnno=" & idAnno & " And Upper(lTrim(rTrim(CodFiscale)))='" & CodFiscale.ToUpper.Trim & "'"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If Not Rec.Eof Then
							Ritorno = StringaErrore & "Codice fiscale già presente in archivio" ' : " & CodFiscale & "--->" & Sql
							Ok = False
						Else
							Ritorno = ""
						End If
						Rec.Close

						Dim Scadenza As DateTime = Convert.ToDateTime(DataDiNascita)
						Dim Anni As Integer = DateAndTime.DateDiff(DateInterval.Year, Scadenza, Now, )
						If Anni >= 18 Then
							Maggiorenne = "S"
						Else
							Maggiorenne = "N"
						End If

						If Maggiorenne = "S" Then
							' Creo utente separato in quanto il giocatore è maggiorenne
							Dim idUtente As Integer = -1

							Sql = "Select Max(idUtente) + 1 From [Generale].[dbo].[Utenti] Where idAnno=" & idAnno
							Rec = LeggeQuery(Conn, Sql, Connessione)
							If TypeOf (Rec) Is String Then
								Ritorno = Rec
							Else
								If Rec(0).Value Is DBNull.Value Then
									idUtente = 1
								Else
									idUtente = Rec(0).Value
								End If
							End If

							Dim pass As String = generaPassRandom()
							Dim nuovaPass() = pass.Split(";")
							Dim s() As String = Squadra.Split("_")
							Dim idSquadra As Integer = Val(s(1))

							Sql = "Insert Into [Generale].[dbo].[Utenti] Values (" &
										" " & idAnno & ", " &
										" " & idUtente & ", " &
										"'" & EMail.Replace("'", "''") & "', " &
										"'" & Cognome.Replace("'", "''") & "', " &
										"'" & Nome.Replace("'", "''") & "', " &
										"'" & nuovaPass(1).Replace("'", "''") & "', " &
										"'" & EMail.Replace("'", "''") & "', " &
										"-1, " &
										"6, " &
										" " & idSquadra & ", " &
										"1, " &
										"'" & Telefono & "', " &
										"'N', " &
										"-1, " &
										"'N' " &
										")"
							Ritorno = EsegueSql(Conn, Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							Else
								Dim m As New mail
								Dim Oggetto As String = "Nuovo utente inCalcio"
								Dim Body As String = ""
								Body &= "E' stato creato l'utente '" & Cognome.ToUpper & " " & Nome.ToUpper & "'. <br />"
								Body &= "Per accedere al sito sarà possibile digitare la mail rilasciata alla segreteria in fase di iscrizione: " & EMail & "<br />"
								Body &= "La password valida per il solo primo accesso è: " & nuovaPass(0) & "<br /><br />"
								Dim ChiScrive As String = "notifiche@incalcio.cloud"

								Ritorno = m.SendEmail(Squadra, "", Oggetto, Body, EMail, "")
							End If
						End If
					Else
						Sql = "SELECT * FROM Giocatori Where idAnno=" & idAnno & " And Upper(lTrim(rTrim(CodFiscale)))='" & CodFiscale.ToUpper.Trim & "'"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If Not Rec.Eof Then
							Maggiorenne = "" & Rec("Maggiorenne").Value
							'Ok = False
						End If
						Rec.Close
					End If

					'If Ok Then
					'	Try
					'		Sql = "SELECT Max(idGiocatore)+1 FROM Giocatori"
					'		Rec = LeggeQuery(Conn, Sql, Connessione)
					'		If TypeOf (Rec) Is String Then
					'			Ritorno = Rec
					'		Else
					'			If Rec(0).Value Is DBNull.Value Then
					'				idGioc = 1
					'			Else
					'				idGioc = Rec(0).Value
					'			End If
					'			Rec.Close()
					'		End If
					'	Catch ex As Exception
					'		Ritorno = StringaErrore & " " & ex.Message
					'		Ok = False
					'	End Try
					'End If
					'Else
					If Ok Then
						Dim GiaCe As Boolean = True

						Sql = "SELECT * FROM Giocatori Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If Rec.Eof Then
							GiaCe = False
							' Dim conta As Integer = 0

							'Do While Ritorno.Contains(StringaErrore) Or Ritorno = ""
							'Try
							'	Sql = "Delete  From Giocatori Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
							'	Ritorno = EsegueSql(Conn, Sql, Connessione)
							'	If Ritorno.Contains(StringaErrore) Then
							'		Ok = False
							'	End If

							'Catch ex As Exception
							'	Ritorno = StringaErrore & " " & ex.Message
							'	Ok = False
							'	'Exit Do
							'End Try
							'	conta += 1
							'	If (conta = 10) Then
							'		Ritorno = StringaErrore & " Impossibile modificare il giocatore"
							'		Ok = False
							'	'Exit Do
							'End If
							'Loop
						End If
						Rec.Close
						idGioc = idGiocatore
						'End If

						If Ok = True Then
							If GiaCe Then
								Sql = "Update Giocatori Set " &
									"idCategoria=" & idCategoria & ", " &
									"idRuolo=" & idRuolo & ", " &
									"Cognome='" & Cognome.Replace("'", "''") & "', " &
									"Nome='" & Nome.Replace("'", "''") & "', " &
									"EMail='" & EMail.Replace("'", "''") & "', " &
									"Telefono='" & Telefono.Replace("'", "''") & "', " &
									"Soprannome='" & Soprannome.Replace("'", "''") & "', " &
									"DataDiNascita='" & DataDiNascita.Replace("'", "''") & "', " &
									"Indirizzo='" & Indirizzo.Replace("'", "''") & "', " &
									"CodFiscale='" & CodFiscale.Replace("'", "''") & "', " &
									"Maschio='" & Maschio & "', " &
									"Citta='" & Citta.Replace("'", "''") & "', " &
									"idTaglia=" & idTaglia & ", " &
									"idCategoria2=" & idCategoria2 & ", " &
									"Matricola='" & Matricola.Replace("'", "''") & "', " &
									"NumeroMaglia='" & NumeroMaglia.Replace("'", "''") & "', " &
									"idCategoria3=" & idCategoria3 & ", " &
									"Categorie='" & Categorie & "', " &
									"RapportoCompleto='" & RapportoCompleto & "', " &
									"Cap='" & Cap & "', " &
									"CittaNascita='" & CittaNascita.Replace("'", "''") & "' " &
									"Where idGiocatore=" & idGiocatore
							Else
								Sql = "Insert Into Giocatori Values (" &
									" " & idAnno & ", " &
									" " & idGioc & ", " &
									" " & idCategoria & ", " &
									" " & idRuolo & ", " &
									"'" & Cognome.Replace("'", "''") & "', " &
									"'" & Nome.Replace("'", "''") & "', " &
									"'" & EMail.Replace("'", "''") & "', " &
									"'" & Telefono.Replace("'", "''") & "', " &
									"'" & Soprannome.Replace("'", "''") & "', " &
									"'" & DataDiNascita.Replace("'", "''") & "', " &
									"'" & Indirizzo.Replace("'", "''") & "', " &
									"'" & CodFiscale.Replace("'", "''") & "', " &
									"'N', " &
									"null, " &
									"'" & Maschio & "', " &
									"'', " &
									"'" & Citta.Replace("'", "''") & "', " &
									" " & idTaglia & ", " &
									" " & idCategoria2 & ", " &
									"'" & Matricola.Replace("'", "''") & "', " &
									"'" & NumeroMaglia.Replace("'", "''") & "', " &
									" " & idCategoria3 & ", " &
									"'" & Categorie & "', " &
									"'" & RapportoCompleto & "', " &
									"'" & Cap & "', " &
									"'" & CittaNascita.Replace("'", "''") & "', " &
									"'" & Maggiorenne & "' " &
									")"
							End If
							Ritorno = EsegueSql(Conn, Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
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
	Public Function EliminaGiocatore(Squadra As String, ByVal idAnno As String, idGiocatore As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Update Giocatori Set Eliminato='S' Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
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
	Public Function AggiungeCategoriaAGiocatore(Squadra As String, ByVal idAnno As String, idGiocatore As String, idCategoria As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Dim Giocatori() As String = idGiocatore.Split(";")

					For Each g As String In Giocatori
						If g <> "" Then
							Try
								Sql = "Update Giocatori Set Categorie = Categorie + '" & idCategoria & "-' Where idAnno=" & idAnno & " And idGiocatore=" & g
								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
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
	Public Function EliminaGiocatoreDallaCategoria(Squadra As String, ByVal idAnno As String, idGiocatore As String, idCategoria As String) As String
		Dim Ritorno As String = ""
		Dim Connessione As String = LeggeImpostazioniDiBase(Server.MapPath("."), Squadra)

		If Connessione = "" Then
			Ritorno = ErroreConnessioneNonValida & ":" & Connessione
		Else
			Dim Conn As Object = ApreDB(Connessione)

			If TypeOf (Conn) Is String Then
				Ritorno = ErroreConnessioneDBNonValida & ":" & Conn
			Else
				Dim Sql As String = ""
				Dim Ok As Boolean = True

				Sql = "Begin transaction"
				Ritorno = EsegueSql(Conn, Sql, Connessione)

				If Not Ritorno.Contains(StringaErrore) Then
					Try
						Sql = "Update Giocatori Set Categorie = Replace(Categorie, '" & idCategoria & "-', '') Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

						Sql = "Update Giocatori Set idCategoria=-1 Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " And idCategoria=" & idCategoria
						Ritorno = EsegueSql(Conn, Sql, Connessione)
						If Ritorno.Contains(StringaErrore) Then
							Ok = False
						End If

						If Ok Then
							Sql = "Update Giocatori Set idCategoria2=-1 Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " And idCategoria2=" & idCategoria
							Ritorno = EsegueSql(Conn, Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							End If
						End If

						If Ok Then
							Sql = "Update Giocatori Set idCategoria3=-1 Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " And idCategoria3=" & idCategoria
							Ritorno = EsegueSql(Conn, Sql, Connessione)
							If Ritorno.Contains(StringaErrore) Then
								Ok = False
							End If
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
	Public Function SalvaPagamento(Squadra As String, idAnno As String, idGiocatore As String, Pagamento As String, Commento As String, idPagatore As String, idRegistratore As String) As String
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
					Dim Progressivo As Integer
					Dim ProgressivoGenerale As Integer

					Dim DataPagamento As String = Now.Year & "-" & Format(Now.Month, "00") & "-" & Format(Now.Day, "00") & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
					Dim Cognome As String = ""
					Dim Nome As String = ""
					Dim CognomeIscritto As String = ""
					Dim NomeIscritto As String = ""
					Dim CodFiscalePagatore As String = ""
					Dim CodFiscaleIscritto As String = ""
					Dim NomeSquadra As String = ""
					Dim NomePolisportiva As String = ""
					Dim Indirizzo As String = ""
					Dim CodiceFiscale As String = ""
					Dim PIva As String = ""
					Dim Telefono As String = ""
					Dim eMail As String = ""

					Try
						Sql = "SELECT * FROM Anni"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If Rec.Eof() Then
							Ritorno = StringaErrore & " Nessuna squadra rilevata"
							Ok = False
						Else
							NomeSquadra = Rec("NomeSquadra").Value
							NomePolisportiva = Rec("NomePolisportiva").Value
							Indirizzo = Rec("Indirizzo").Value
							CodiceFiscale = Rec("CodiceFiscale").Value
							PIva = Rec("PIva").Value
							Telefono = Rec("Telefono").Value
							eMail = Rec("Mail").Value
						End If
						Rec.Close()

						If Ok Then
							If idPagatore = 3 Then
								Sql = "SELECT * FROM Giocatori Where idGiocatore=" & idGiocatore
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Nessun giocatore rilevato"
									Ok = False
								Else
									Cognome = Rec("Cognome").Value
									Nome = Rec("Nome").Value
									CodFiscalePagatore = Rec("CodFiscale").Value

									CognomeIscritto = Rec("Cognome").Value
									NomeIscritto = Rec("Nome").Value
									CodFiscaleIscritto = Rec("CodFiscale").Value
								End If
								Rec.Close()
							Else
								Sql = "SELECT * FROM Giocatori Where idGiocatore=" & idGiocatore
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Nessun giocatore rilevato"
									Ok = False
								Else
									CognomeIscritto = Rec("Cognome").Value
									NomeIscritto = Rec("Nome").Value
									CodFiscaleIscritto = Rec("CodFiscale").Value
								End If
								Rec.Close()

								Sql = "SELECT * FROM GiocatoriDettaglio Where idGiocatore=" & idGiocatore
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If Rec.Eof() Then
									Ritorno = StringaErrore & " Nessun giocatore rilevato"
									Ok = False
								Else
									If idPagatore = 1 Then
										Cognome = Rec("Genitore1").Value
										Nome = ""
										CodFiscalePagatore = Rec("CodFiscale1").Value
									Else
										Cognome = Rec("Genitore2").Value
										Nome = ""
										CodFiscalePagatore = Rec("CodFiscale2").Value
									End If
								End If
								Rec.Close()
							End If

							If Ok Then
								Sql = "SELECT Max(Progressivo)+1 FROM DatiFattura Where Anno=" & Now.Year
								Rec = LeggeQuery(Conn, Sql, Connessione)
								If Rec(0).Value Is DBNull.Value Then
									ProgressivoGenerale = 1
									Sql = "Insert Into DatiFattura Values(" & Now.Year & ", 1)"
								Else
									ProgressivoGenerale = Rec(0).Value
									Sql = "Update DatiFattura Set Progressivo = " & ProgressivoGenerale & " Where Anno=" & Now.Year
								End If
								Rec.Close()

								Ritorno = EsegueSql(Conn, Sql, Connessione)
								If Ritorno.Contains(StringaErrore) Then
									Ok = False
								End If

								If Ok Then
									Sql = "SELECT Max(Progressivo)+1 FROM GiocatoriPagamenti Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
									Rec = LeggeQuery(Conn, Sql, Connessione)
									If Rec(0).Value Is DBNull.Value Then
										Progressivo = 1
									Else
										Progressivo = Rec(0).Value
									End If
									Rec.Close()

									Sql = "Insert Into GiocatoriPagamenti Values (" &
										" " & idAnno & ", " &
										" " & idGiocatore & ", " &
										" " & Progressivo & ", " &
										" " & Pagamento & ", " &
										"'" & DataPagamento & "', " &
										"'N', " &
										"'" & Commento.Replace("'", "''") & "', " &
										" " & idPagatore & ", " &
										" " & idRegistratore & " " &
										")"
									Ritorno = EsegueSql(Conn, Sql, Connessione)
									If Ritorno.Contains(StringaErrore) Then
										Ok = False
									End If
								End If
							End If
						End If
					Catch ex As Exception
						Ritorno = StringaErrore & " " & ex.Message
						Ok = False
					End Try

					If Ok Then
						Dim gf As New GestioneFilesDirectory
						Dim filePaths As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
						Dim p() As String = filePaths.Split(";")
						If Strings.Right(p(0), 1) <> "\" Then
							p(0) &= "\"
						End If
						If Strings.Right(p(2), 1) <> "/" Then
							p(2) = p(2) & "/"
						End If
						' Dim url As String = p(2) & NomeSquadra.Replace(" ", "_") & "/Societa/" & idAnno & "_1.jpg"

						Dim pp As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
						pp = pp.Trim()
						If Strings.Right(pp, 1) = "\" Then
							pp = Mid(pp, 1, pp.Length - 1)
						End If
						Dim Esten As String = Format(Now.Second, "00") & "_" & Now.Millisecond & RitornaValoreRandom(55)

						p(2) = p(2).Replace(vbCrLf, "")
						Dim nomeImm As String = p(2) & NomeSquadra.Replace(" ", "_") & "/Societa/" & idAnno & "_1.kgb"
						Dim pathImm As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Societa\" & idAnno & "_1.kgb"
						Dim nomeImmConv As String = p(2) & "Appoggio/Societa_" & idAnno & "_1_" & Esten & ".png"
						Dim pathImmConv As String = pp & "\Appoggio\Societa_" & idAnno & "_1_" & Esten & ".png"
						Dim c As New CriptaFiles
						c.DecryptFile("WPippoBaudo227!", pathImm, pathImmConv)

						Dim Body As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Scheletri\ricevuta_pagamento.txt")
						Dim path As String = p(0) & "\" & Squadra & "\Ricevute\Anno" & idAnno & "\" & idGiocatore & "\"
						gf.CreaDirectoryDaPercorso(path)
						Dim fileFinale As String = path & "Ricevuta_" & Progressivo & ".pdf"
						Dim fileAppoggio As String = path & "Ricevuta_" & Progressivo & ".html"

						Dim Intero As String
						Dim Virgola As String

						If Pagamento.Contains(",") Or Pagamento.Contains(".") Then
							If Pagamento.Contains(".") Then
								Dim pp1() As String = Pagamento.Split(".")
								Intero = pp1(0)
								Virgola = pp1(1)
							Else
								Dim pp22() As String = Pagamento.Split(",")
								Intero = pp22(0)
								Virgola = pp22(1)
							End If
						Else
							Intero = Pagamento
							Virgola = ""
						End If

						If Virgola = "" Then
							Virgola = "00"
						Else
							If Virgola.Length = 1 Then
								Virgola = "0" & Virgola
							Else
								If Virgola > 2 Then
									Virgola = Mid(Virgola, 1, 2)
								End If
							End If
						End If

						Dim Dati As String = "C.F.: " & CodiceFiscale & " P.I.:" & PIva & "<br />Telefono: " & Telefono & "<br />E-Mail: " & eMail
						Dim Altro As String = ""
						If Commento <> "" Then
							Altro = "- " & Commento
						End If

						Body = Body.Replace("***URL LOGO***", nomeImmConv)
						Body = Body.Replace("***NOME POLISPORTIVA***", NomePolisportiva)
						Body = Body.Replace("***INDIRIZZO***", Indirizzo)
						Body = Body.Replace("***DATI***", Dati)
						Body = Body.Replace("***NUMERO_RICEVUTA***", ProgressivoGenerale & "/" & Now.Year)
						Body = Body.Replace("***DATA_RICEVUTA***", Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year)
						Body = Body.Replace("***NOME***", Cognome & " " & Nome)
						Body = Body.Replace("***MOTIVAZIONE***", CognomeIscritto & " " & NomeIscritto & " " & CodFiscaleIscritto & " " & Altro)
						Body = Body.Replace("***IMPORTO***", Intero)
						Body = Body.Replace("***VIRGOLE***", Virgola)

						Dim Cifre1 As String = convertNumberToReadableString(Val(Intero))
						Dim Cifre2 As String = convertNumberToReadableString(Val(Virgola))
						Dim Altro2 As String = ""
						If Cifre2 <> "" Then
							Altro2 = "/" & Virgola
						End If
						Body = Body.Replace("***IMPORTO LETTERE***", Cifre1 & Altro2)

						filePaths = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\Paths.txt")
						If Strings.Right(filePaths, 1) <> "\" Then
							filePaths &= "\"
						End If
						' Dim pathFirma As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Firme\" & idAnno & "_" & idGiocatore & "_" & idPagatore & ".png"
						' Dim pathFirma As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Segreteria\" & idAnno & ".kgb"

						Dim pathFirma As String = filePaths & NomeSquadra.Replace(" ", "_") & "\Segreteria\" & idAnno & ".kgb"
						Dim urlFirma As String = pp & "\" & NomeSquadra.Replace(" ", "_") & "\Segreteria\" & idAnno & ".kgb"
						Dim pathFirmaConv As String = p(2) & "/Appoggio/Segreteria_" & Esten & ".png"
						Dim urlFirmaConv As String = pp & "\Appoggio\Segreteria_" & Esten & ".png"
						c.DecryptFile("WPippoBaudo227!", urlFirma, urlFirmaConv)

						Body = Body.Replace("***URL FIRMA***", urlFirmaConv)

						gf.EliminaFileFisico(fileAppoggio)
						gf.ApreFileDiTestoPerScrittura(fileAppoggio)
						gf.ScriveTestoSuFileAperto(Body)

						gf.ChiudeFileDiTestoDopoScrittura()

						Dim pp2 As New pdfGest
						Ritorno = pp2.ConverteHTMLInPDF(fileAppoggio, fileFinale, "")
						If Ritorno <> "*" Then
							Ok = False
						End If
					End If
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
	Public Function EliminaPagamentoGiocatore(Squadra As String, idAnno As String, idGiocatore As String, Progressivo As String, idRegistratore As String) As String
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
						Sql = "Update GiocatoriPagamenti Set " &
							"Eliminato='S' " &
							"Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " And Progressivo=" & Progressivo
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
							Dim ora As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")
							Sql = "Insert Into GiocatoriPagamentiEliminazioni Values (" & idAnno & ", " & idGiocatore & ", " & Progressivo & ", " & idRegistratore & ", '" & ora & "')"
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
	Public Function RitornaPagamentiGiocatore(Squadra As String, idAnno As String, idGiocatore As String) As String
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
				Dim TotPag As Single = 0

				Sql = "Select * From GiocatoriDettaglio Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Not Rec.eof Then
						TotPag = Rec("TotalePagamento").Value
						Rec.Close

						Sql = "Select * From GiocatoriPagamenti Where idAnno=" & idAnno & " And idGiocatore=" & idGiocatore & " And Eliminato='N' Order By Progressivo"
						Rec = LeggeQuery(Conn, Sql, Connessione)
						If TypeOf (Rec) Is String Then
							Ritorno = Rec
						Else
							Dim Totale As Single = 0
							'Ritorno = "Totale a pagare;" & Format(TotPag, "#0.#0") & ";;§"
							Dim Ritorno2 As String = ""
							Do Until Rec.Eof
								Ritorno2 &= Rec("Progressivo").Value & ";" & Rec("Pagamento").Value & ";" & Rec("DataPagamento").Value & ";" & Rec("Commento").Value & ";§"
								Totale += (Rec("Pagamento").Value)

								Rec.MoveNext
							Loop
							Rec.Close

							Ritorno = Totale & ";" & TotPag & ";|" & Ritorno2

							'Ritorno &= "Totale;" & Format(Totale, "#0.#0") & ";;§"
							'Dim Differenza As Single = TotPag - Totale
							'Differenza = CInt(Differenza * 100) / 100
							'Ritorno &= "Differenza;" & Format(Differenza, "#0.#0") & ";;§"
						End If
					Else
						Ritorno = StringaErrore & ": Nessun pagamento impostato"
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

	<WebMethod()>
	Public Function ControllaEsistenzaModuloIscrizione(Squadra As String, Anno As String, idGiocatore As String) As String
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
				Dim Sql As String

				Sql = "Select NomeSquadra, Descrizione From Anni Where idAnno = " & Anno
				Rec = LeggeQuery(Conn, Sql, Connessione)
				If TypeOf (Rec) Is String Then
					Ritorno = Rec
				Else
					If Rec.Eof Then
						Ritorno = StringaErrore & " Nessuna squadra rilevata"
					Else
						Dim NomeSquadra As String = Rec("NomeSquadra").Value
						Dim Descrizione As String = Rec("Descrizione").Value
						Rec.Close

						Dim gf As New GestioneFilesDirectory
						Dim PathAllegati As String = gf.LeggeFileIntero(Server.MapPath(".") & "\Impostazioni\PathAllegati.txt")
						Dim P() As String = PathAllegati.Split(";")
						If Strings.Right(P(0), 1) = "\" Then
							P(0) = Mid(P(0), 1, P(0).Length - 1)
						End If
						Dim fileDaCopiare As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".html"
						Dim fileDaCopiarePDF As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".pdf"
						Dim fileLog As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & ".log"
						'Dim fileDaCopiare2 As String = P(0) & "\" & Squadra & "\Firme\iscrizione_" & Anno & "_" & idGiocatore & "_send.html"
						gf.CreaDirectoryDaPercorso(fileDaCopiare)
						Dim fileScheletro As String = Server.MapPath(".") & "\Scheletri\base_iscrizione_.txt"

						If File.Exists(fileScheletro) Then
							Try
								Dim fileFirme As String = gf.LeggeFileIntero(fileScheletro)
								fileFirme = RiempieFileFirme(fileFirme, Anno, idGiocatore, Rec, Conn, Connessione, NomeSquadra, P, Descrizione)

								gf.EliminaFileFisico(fileDaCopiare)
								gf.ApreFileDiTestoPerScrittura(fileDaCopiare)
								gf.ScriveTestoSuFileAperto(fileFirme)
								gf.ChiudeFileDiTestoDopoScrittura()

								'File.Copy(fileDaCopiare, fileDaCopiare2)
								Dim pp As New pdfGest
								Ritorno = pp.ConverteHTMLInPDF(fileDaCopiare, fileDaCopiarePDF, fileLog)

								If Ritorno = "*" Then
									gf.EliminaFileFisico(fileDaCopiare)
								End If

								gf = Nothing
							Catch ex As Exception
								Ritorno = StringaErrore & " " & ex.Message
							End Try
						Else
							Ritorno = StringaErrore & " Scheletro iscrizione non trovato"
						End If
						gf = Nothing
					End If
				End If
			End If
		End If

		Return Ritorno
	End Function

End Class